"""
/gpreport Slack slash command
--------------------------------
Usage in Slack:  /gpreport 20-01-2026 to 30-01-2026

What it does:
1. Slack POSTs the slash command to /slack/gp-report on this Flask app.
2. We immediately ack (Slack requires a response within 3s) with an
   ephemeral "generating..." message, then do the real work in a
   background thread.
3. Background thread:
   - Pulls ALL orders created in the given date range from Shopify
     Admin REST API (paginated via the Link header).
   - Keeps only line items belonging to orders whose financial_status
     is paid, partially_paid, or pending.
   - Looks up Item Cost per SKU from Airtable's French Inventories table
     (batched, not one request per line item).
   - Builds a styled .xlsx with columns:
     Name, Payment Status, Created at, Fulfilled at, Lineitem name,
     Lineitem sku, Lineitem price, Lineitem quantity, Shipping, Taxes,
     Discount, Total, Item Cost, Total Cost, Shipping (Aramex),
     Gateway, Net Cost, GP, GP%
     — where:
       Total       = (price * qty) + Shipping - Discount
       Total Cost  = Item Cost * qty
       Shipping (Carrier) = Aramex: city-based rate (Dubai 18.67,
                     Sharjah 20, Ajman 21, Abu Dhabi/Fujairah/
                     Ras Al Khaimah/Umm Al Quwain 22.67). Professional
                     Courier: flat 31.5. Matched against
                     fulfillments[].tracking_company.
       Gateway     = Total * gateway fee (COD 0%, Tabby 9.5%, Card 3.2%),
                     matched against payment_gateway_names
       Net Cost    = Item Cost + Shipping (Aramex) + Gateway
       GP          = Total - Net Cost
       GP%         = (GP / Total) * 100
     All computed as plain values (not Excel formulas), so they display
     correctly in any preview, not just when opened in real Excel.
     Unrecognized cities/gateways are logged as warnings rather than
     silently guessed.
   - Uploads the file to the requesting Slack channel using Slack's
     current 3-step external upload flow (files.upload is deprecated).

Required environment variables (set these on Render):
  SHOPIFY_STORE          e.g. "fragrantsouq.myshopify.com"
  SHOPIFY_ADMIN_TOKEN    Shopify Admin API access token (read_orders scope)
  SLACK_BOT_TOKEN        Bot token with files:write scope, bot invited to
                          the channel this command will be run from
  SLACK_SIGNING_SECRET   (optional but recommended) used to verify the
                          request really came from Slack
  AIRTABLE_API_KEY       Airtable personal access token, read access to
                          the French Inventories table
  AIRTABLE_BASE_ID       (optional) defaults to app5gOqDt9aZrW5bV
  AIRTABLE_TABLE_NAME    (optional) defaults to "French Inventories"
  AIRTABLE_SKU_FIELD     (optional) defaults to "SKU" — the field name in
                          that table holding the SKU. Change this if your
                          actual field is named differently.
  AIRTABLE_COST_FIELD    (optional) defaults to "Cost"

Slack app setup:
  - Create slash command "/gpreport" (no spaces allowed in the command
    name itself) with Request URL:
      https://<your-render-service>.onrender.com/slack/gp-report
    Usage Hint: [start date] to [end date] e.g. 20-01-2026 to 30-01-2026
  - Bot token scopes needed: files:write, chat:write
  - Invite the bot to whichever channel(s) will run the command
"""

import hashlib
import hmac
import os
import re
import threading
import time
import unicodedata
from datetime import datetime
from io import BytesIO

import requests
from flask import Flask, jsonify, request
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

app = Flask(__name__)

SHOPIFY_STORE = os.environ["SHOPIFY_STORE"]
SHOPIFY_ADMIN_TOKEN = os.environ["SHOPIFY_ADMIN_TOKEN"]
SHOPIFY_API_VERSION = "2026-04"
SLACK_BOT_TOKEN = os.environ["SLACK_BOT_TOKEN"]
SLACK_SIGNING_SECRET = os.environ.get("SLACK_SIGNING_SECRET")  # optional

AIRTABLE_API_KEY = os.environ["AIRTABLE_API_KEY"]
AIRTABLE_BASE_ID = os.environ.get("AIRTABLE_BASE_ID", "app5gOqDt9aZrW5bV")
AIRTABLE_TABLE_NAME = os.environ.get("AIRTABLE_TABLE_NAME", "French Inventories")
AIRTABLE_SKU_FIELD = os.environ.get("AIRTABLE_SKU_FIELD", "SKU")
AIRTABLE_COST_FIELD = os.environ.get("AIRTABLE_COST_FIELD", "Cost")

TARGET_STATUSES = {"paid", "partially_paid", "pending"}

COLUMNS = [
    "Name",
    "Payment Status",
    "Created at",
    "Fulfilled at",
    "Lineitem name",
    "Lineitem sku",
    "Lineitem price",
    "Lineitem quantity",
    "Shipping",
    "Taxes",
    "Discount",
    "Total",
    "Item Cost",
    "Total Cost",
    "Shipping",       # Carrier shipping charge (Aramex: city-based, Professional
                       # Courier: flat rate) — distinct from the order-level
                       # "Shipping" column above; kept as a second "Shipping"
                       # column to match your template
    "Gateway",
    "Net Cost",
    "GP",
    "GP%",
]

# Aramex charge by destination emirate (AED). Matched against the shipping
# address's PROVINCE (not city — UAE addresses use area/neighborhood names
# for city, e.g. "Barsha Heights", "Khalifa City", which don't reliably
# indicate the emirate). UAE's official province codes are used first,
# with a full-name fallback.
ARAMEX_PROVINCE_CODE_RATES = {
    "du": 18.67,   # Dubai
    "sh": 20.00,   # Sharjah
    "aj": 21.00,   # Ajman
    "az": 22.67,   # Abu Dhabi
    "fu": 22.67,   # Fujairah
    "rk": 22.67,   # Ras Al Khaimah
    "uq": 22.67,   # Umm Al Quwain
}
ARAMEX_PROVINCE_NAME_RATES = [
    (("dubai",), 18.67),
    (("sharj",), 20.00),  # matches "sharjah" and "sharja"
    (("ajman",), 21.00),
    (("abu dhabi", "fujair", "ras al khaim", "ras al khaym", "umm al quwain", "ummal quin"), 22.67),
]

# Professional Courier is a flat rate, unlike Aramex's emirate-based rates.
PROFESSIONAL_COURIER_RATE = 31.50

# Porter is a flat rate, used for Dubai deliveries.
PORTER_RATE = 45.00



# Payment gateway fee as a fraction of Total. Matched against Shopify's
# payment_gateway_names (case-insensitive substring match). "manual" is
# included under COD because Shopify records COD as a "manual" payment
# gateway internally.
GATEWAY_RATES = [
    (("cod", "cash on delivery", "cash_on_delivery", "manual"), 0.0, "COD"),
    (("tabby",), 0.095, "Tabby"),
    (("card", "shopify_payments", "stripe", "credit"), 0.032, "Card"),
]

DATE_RE = re.compile(r"^\d{2}-\d{2}-\d{4}$")  # DD-MM-YYYY


def normalize_text(text: str) -> str:
    """Lowercase, strip accents (e.g. 'Dubaï' -> 'dubai'), and replace
    hyphens/underscores with spaces (e.g. 'Abu-dhabi' -> 'abu dhabi')."""
    if not text:
        return ""
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = text.replace("-", " ").replace("_", " ")
    return text.strip().lower()


def match_aramex_rate(province: str, province_code: str):
    """Return the Aramex shipping charge for a UAE emirate, using the
    shipping address's province_code (preferred, e.g. 'DU') with a
    full-name fallback (e.g. 'Dubai'). Returns None if unrecognized —
    which is expected/correct for non-UAE addresses."""
    code = normalize_text(province_code)
    if code in ARAMEX_PROVINCE_CODE_RATES:
        return ARAMEX_PROVINCE_CODE_RATES[code]

    normalized_name = normalize_text(province)
    if normalized_name:
        for keywords, rate in ARAMEX_PROVINCE_NAME_RATES:
            if any(kw in normalized_name for kw in keywords):
                return rate
    return None


def match_gateway_rate(gateway_names: list):
    """Return (rate, label) for a list of Shopify payment_gateway_names,
    or (None, None) if none of the known patterns match."""
    combined = normalize_text(" ".join(gateway_names or []))
    for keywords, rate, label in GATEWAY_RATES:
        if any(kw in combined for kw in keywords):
            return rate, label
    return None, None



# ---------------------------------------------------------------------------
# Slack request verification (optional, but recommended)
# ---------------------------------------------------------------------------
def verify_slack_signature(req) -> bool:
    if not SLACK_SIGNING_SECRET:
        return True  # verification skipped if secret not configured

    timestamp = req.headers.get("X-Slack-Request-Timestamp", "")
    if not timestamp:
        print("[gpreport] Signature check failed: no timestamp header present.", flush=True)
        return False

    age = abs(time.time() - int(timestamp))
    if age > 60 * 5:
        print(f"[gpreport] Signature check failed: timestamp too old/skewed ({age:.0f}s).", flush=True)
        return False

    raw_body = req.get_data(as_text=True)
    sig_basestring = f"v0:{timestamp}:{raw_body}"
    my_sig = (
        "v0="
        + hmac.new(
            SLACK_SIGNING_SECRET.encode(),
            sig_basestring.encode(),
            hashlib.sha256,
        ).hexdigest()
    )
    slack_sig = req.headers.get("X-Slack-Signature", "")

    if not hmac.compare_digest(my_sig, slack_sig):
        # Don't log full signatures/secret — just enough to spot common issues
        # (e.g. wrong secret entirely, empty body, mismatched timestamp).
        print(
            "[gpreport] Signature check failed: computed sig doesn't match Slack's. "
            f"body_len={len(raw_body)} timestamp_age={age:.0f}s "
            f"mine_prefix={my_sig[:10]} slack_prefix={slack_sig[:10]}",
            flush=True,
        )
        return False

    return True


# ---------------------------------------------------------------------------
# Shopify
# ---------------------------------------------------------------------------
def fetch_orders(start_date: datetime, end_date: datetime) -> list:
    """Fetch all orders (any status) created within the date range, paginated.
    start_date/end_date are datetime.date objects."""
    orders = []
    base_url = f"https://{SHOPIFY_STORE}/admin/api/{SHOPIFY_API_VERSION}/orders.json"
    headers = {"X-Shopify-Access-Token": SHOPIFY_ADMIN_TOKEN}
    params = {
        "status": "any",
        "created_at_min": f"{start_date.isoformat()}T00:00:00+04:00",
        "created_at_max": f"{end_date.isoformat()}T23:59:59+04:00",
        "limit": 250,
    }

    url = base_url
    page_num = 1
    while url:
        print(f"[gpreport] Fetching orders page {page_num}...", flush=True)
        resp = requests.get(
            url,
            headers=headers,
            params=params if url == base_url else None,
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        page_orders = data.get("orders", [])
        orders.extend(page_orders)
        print(f"[gpreport] Page {page_num}: got {len(page_orders)} orders (total so far: {len(orders)})", flush=True)

        next_url = None
        link_header = resp.headers.get("Link", "")
        if link_header:
            for part in link_header.split(","):
                if 'rel="next"' in part:
                    next_url = part[part.find("<") + 1 : part.find(">")]
        url = next_url
        page_num += 1

    print(f"[gpreport] Done fetching. {len(orders)} total orders in range.", flush=True)
    return orders


# ---------------------------------------------------------------------------
# Airtable — look up Item Cost per SKU from "French Inventories"
# ---------------------------------------------------------------------------
def fetch_costs_from_airtable(skus: set) -> dict:
    """Given a set of SKUs, return {sku: cost} looked up from Airtable's
    French Inventories table. Batches lookups (POST listRecords, which
    avoids GET URL-length limits) so we don't do one request per SKU."""
    skus = [s for s in skus if s]  # drop blanks
    if not skus:
        return {}

    costs = {}
    url = f"https://api.airtable.com/v0/{AIRTABLE_BASE_ID}/{AIRTABLE_TABLE_NAME}/listRecords"
    headers = {
        "Authorization": f"Bearer {AIRTABLE_API_KEY}",
        "Content-Type": "application/json",
    }

    batch_size = 100
    batches = [skus[i : i + batch_size] for i in range(0, len(skus), batch_size)]
    print(f"[gpreport] Looking up cost for {len(skus)} unique SKUs in {len(batches)} Airtable batch(es)...", flush=True)

    for batch_num, batch in enumerate(batches, start=1):
        conditions = ",".join(f"{{{AIRTABLE_SKU_FIELD}}}='{sku}'" for sku in batch)
        formula = f"OR({conditions})"

        offset = None
        while True:
            body = {
                "filterByFormula": formula,
                "fields": [AIRTABLE_SKU_FIELD, AIRTABLE_COST_FIELD],
                "pageSize": 100,
            }
            if offset:
                body["offset"] = offset

            resp = requests.post(url, headers=headers, json=body, timeout=30)
            if resp.status_code != 200:
                print(f"[gpreport] Airtable cost lookup batch {batch_num} FAILED: {resp.status_code} {resp.text}", flush=True)
                break
            data = resp.json()

            for record in data.get("records", []):
                fields = record.get("fields", {})
                sku = fields.get(AIRTABLE_SKU_FIELD)
                cost = fields.get(AIRTABLE_COST_FIELD)
                if sku is not None and cost is not None:
                    costs[sku] = cost

            offset = data.get("offset")
            if not offset:
                break

        print(f"[gpreport] Airtable batch {batch_num}/{len(batches)} done. {len(costs)} costs found so far.", flush=True)

    missing = set(skus) - set(costs.keys())
    if missing:
        print(f"[gpreport] {len(missing)} SKU(s) had no Airtable cost match (left blank in report).", flush=True)

    return costs


# ---------------------------------------------------------------------------
# Excel building
# ---------------------------------------------------------------------------
STATUS_LABELS = {
    "paid": "Paid",
    "partially_paid": "Partially Paid",
    "pending": "Pending",
}


def build_excel(orders: list) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "GP Report"

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    for col_idx, col_name in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"

    # ---- Pass 1: collect qualifying line-item rows + the set of SKUs we'll
    # need cost data for, without writing to the sheet yet.
    status_counts = {"paid": 0, "partially_paid": 0, "pending": 0}
    pending_rows = []  # list of dicts, one per line item row
    all_skus = set()

    for order in orders:
        financial_status = order.get("financial_status")
        if financial_status not in TARGET_STATUSES:
            continue
        status_counts[financial_status] += 1

        line_items = order.get("line_items", [])
        if not line_items:
            continue

        name = order.get("name", "")
        status_label = STATUS_LABELS.get(financial_status, financial_status)
        created_at = order.get("created_at", "")

        fulfilled_at = ""
        for fulfillment in order.get("fulfillments", []):
            if fulfillment.get("created_at"):
                fulfilled_at = fulfillment["created_at"]
                break

        shipping_total = sum(
            float(s.get("price", 0) or 0) for s in order.get("shipping_lines", [])
        )
        taxes_total = order.get("total_tax", "0.00")
        discount_total = float(order.get("total_discounts", 0) or 0)

        # Shipping and Discount are order-level totals, not per line item.
        # Allocate each proportionally to a row's share of the order's
        # TOTAL QUANTITY (not per-row-count), so a line item with qty=2
        # gets twice the share of one with qty=1, and the rows still sum
        # back to the order's actual shipping/discount total.
        order_total_qty = sum(float(li.get("quantity") or 0) for li in line_items)

        tracking_company = ""
        tracking_urls_combined = ""
        for fulfillment in order.get("fulfillments", []):
            if fulfillment.get("tracking_company"):
                tracking_company = fulfillment["tracking_company"]
            urls = fulfillment.get("tracking_urls") or (
                [fulfillment["tracking_url"]] if fulfillment.get("tracking_url") else []
            )
            if urls:
                tracking_urls_combined = " ".join(urls)
            if tracking_company or tracking_urls_combined:
                break

        # Diagnostic: log the raw carrier data Shopify actually returned for
        # this order, so we can tell whether "Aramex" shows up as literal
        # text in tracking_company vs. only being inferred by Shopify's own
        # UI from the tracking number/URL pattern.
        if order.get("fulfillments"):
            print(
                f"[gpreport] Carrier check {order.get('name')}: "
                f"tracking_company={tracking_company!r} tracking_url={tracking_urls_combined!r}",
                flush=True,
            )

        shipping_addr = order.get("shipping_address") or {}
        dest_city = shipping_addr.get("city", "")
        dest_province = shipping_addr.get("province", "")
        dest_province_code = shipping_addr.get("province_code", "")
        gateway_names = order.get("payment_gateway_names") or []
        if not gateway_names and order.get("gateway"):
            # Fallback: some orders (esp. older or certain checkout flows)
            # only populate the older singular "gateway" field instead of
            # payment_gateway_names.
            gateway_names = [order["gateway"]]

        for li in line_items:
            sku = li.get("sku", "")
            if sku:
                all_skus.add(sku)

            li_qty = float(li.get("quantity") or 0)
            qty_share = (li_qty / order_total_qty) if order_total_qty > 0 else 0
            shipping_share = round(shipping_total * qty_share, 2)
            discount_share = round(discount_total * qty_share, 2)

            pending_rows.append(
                {
                    "name": name,
                    "status_label": status_label,
                    "created_at": created_at,
                    "fulfilled_at": fulfilled_at,
                    "li_name": li.get("name", ""),
                    "sku": sku,
                    "price": li.get("price", ""),
                    "qty": li.get("quantity", ""),
                    "shipping": shipping_share,
                    "taxes": taxes_total,
                    "discount": discount_share,
                    "tracking_company": tracking_company,
                    "tracking_url": tracking_urls_combined,
                    "dest_city": dest_city,
                    "dest_province": dest_province,
                    "dest_province_code": dest_province_code,
                    "gateway_names": gateway_names,
                }
            )


    # ---- Batch-fetch Item Cost per SKU from Airtable, once, for every SKU
    # we'll need across the whole report.
    cost_by_sku = fetch_costs_from_airtable(all_skus)

    # The "Shipping" header appears twice (order-level Shopify shipping,
    # and the Aramex carrier charge) — resolve the second one's column
    # index explicitly rather than relying on COLUMNS.index(), which would
    # only find the first match.
    aramex_shipping_col = [i for i, c in enumerate(COLUMNS) if c == "Shipping"][-1] + 1

    # ---- Pass 2: write rows. All computed columns are plain numbers in
    # Python (not Excel formulas) — openpyxl doesn't calculate formula
    # results itself, so viewers without a calc engine (Slack's inline
    # preview, some other tools) show blank cells for formulas until the
    # file is opened in real Excel. Plain values display correctly
    # everywhere.
    row_idx = 2
    rows_missing_cost = 0
    rows_unmapped_aramex_city = 0
    rows_unmapped_gateway = 0
    for row in pending_rows:
        values = [
            row["name"],
            row["status_label"],
            row["created_at"],
            row["fulfilled_at"],
            row["li_name"],
            row["sku"],
            row["price"],
            row["qty"],
            row["shipping"],
            row["taxes"],
            row["discount"],
        ]
        for col_idx, value in enumerate(values, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)

        item_cost = cost_by_sku.get(row["sku"])  # None if no match found
        if item_cost is None:
            rows_missing_cost += 1

        try:
            price = float(row["price"] or 0)
            qty = float(row["qty"] or 0)
            shipping = float(row["shipping"] or 0)
            discount = float(row["discount"] or 0)
            total = (price * qty) + shipping - discount
        except (TypeError, ValueError) as e:
            print(f"[gpreport] WARN row {row_idx} ({row['sku']}): couldn't compute Total — {e}", flush=True)
            total = None
            qty = 0

        total_cost = None
        if item_cost is not None:
            try:
                total_cost = float(item_cost) * qty
            except (TypeError, ValueError) as e:
                print(f"[gpreport] WARN row {row_idx} ({row['sku']}): couldn't compute Total Cost — {e}", flush=True)

        # Carrier shipping charge: Aramex is emirate-based (matched via
        # province/province_code — NOT city, since UAE addresses use area/
        # neighborhood names for city that don't reliably indicate the
        # emirate). Professional Courier is a flat rate. Any other/
        # unrecognized carrier is left blank. Checking both
        # tracking_company AND tracking_url because Shopify sometimes shows
        # "Aramex tracking" in the admin UI purely from auto-detecting the
        # tracking number/URL pattern, without tracking_company itself
        # literally saying "Aramex".
        carrier_text = f"{row['tracking_company'] or ''} {row['tracking_url'] or ''}".lower()
        carrier_shipping = None
        if "aramex" in carrier_text:
            carrier_shipping = match_aramex_rate(row["dest_province"], row["dest_province_code"])
            if carrier_shipping is None:
                rows_unmapped_aramex_city += 1
                print(
                    f"[gpreport] WARN row {row_idx} ({row['name']}): Aramex order with unrecognized "
                    f"province (province='{row['dest_province']}' code='{row['dest_province_code']}' "
                    f"city='{row['dest_city']}') — shipping charge left blank.",
                    flush=True,
                )
            else:
                print(
                    f"[gpreport] Matched {row['name']}: Aramex, province='{row['dest_province']}' "
                    f"code='{row['dest_province_code']}' -> {carrier_shipping} AED",
                    flush=True,
                )
        elif "porter" in carrier_text:
            carrier_shipping = PORTER_RATE
            print(
                f"[gpreport] Matched {row['name']}: Porter -> {PORTER_RATE} AED",
                flush=True,
            )
        elif "professional" in carrier_text or (row["tracking_company"] or "").strip().lower() == "other":
            carrier_shipping = PROFESSIONAL_COURIER_RATE
            print(
                f"[gpreport] Matched {row['name']}: Professional Courier "
                f"(tracking_company={row['tracking_company']!r}) -> {PROFESSIONAL_COURIER_RATE} AED",
                flush=True,
            )

        # Payment gateway fee, as a percentage of Total
        gateway_rate, gateway_label = match_gateway_rate(row["gateway_names"])
        gateway_charge = None
        if gateway_rate is not None and total is not None:
            gateway_charge = round(total * gateway_rate, 2)
        elif gateway_rate is None:
            rows_unmapped_gateway += 1
            print(
                f"[gpreport] WARN row {row_idx} ({row['name']}): unrecognized payment gateway "
                f"{row['gateway_names']} — gateway charge left blank.",
                flush=True,
            )

        # Net Cost / GP / GP% only make sense once we have an Item Cost;
        # missing shipping/gateway values are treated as 0 in the sum
        # (they're already flagged above via the warnings) so one unknown
        # component doesn't blank out the whole row.
        net_cost = None
        gp = None
        gp_percent = None
        if item_cost is not None:
            net_cost = float(item_cost) + (carrier_shipping or 0) + (gateway_charge or 0)
            if total is not None:
                gp = total - net_cost
                if total != 0:
                    gp_percent = round((gp / total) * 100, 2)

        ws.cell(row=row_idx, column=COLUMNS.index("Total") + 1, value=total)
        ws.cell(row=row_idx, column=COLUMNS.index("Item Cost") + 1, value=item_cost)
        ws.cell(row=row_idx, column=COLUMNS.index("Total Cost") + 1, value=total_cost)
        ws.cell(row=row_idx, column=aramex_shipping_col, value=carrier_shipping)
        ws.cell(row=row_idx, column=COLUMNS.index("Gateway") + 1, value=gateway_charge)
        ws.cell(row=row_idx, column=COLUMNS.index("Net Cost") + 1, value=net_cost)
        ws.cell(row=row_idx, column=COLUMNS.index("GP") + 1, value=gp)
        ws.cell(row=row_idx, column=COLUMNS.index("GP%") + 1, value=gp_percent)

        row_idx += 1

    print(
        f"[gpreport] {rows_missing_cost} of {row_idx - 2} rows had no Airtable cost match. "
        f"{rows_unmapped_aramex_city} Aramex rows had an unrecognized city. "
        f"{rows_unmapped_gateway} rows had an unrecognized payment gateway.",
        flush=True,
    )

    for col_idx, col_name in enumerate(COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = max(14, len(col_name) + 6)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    print(
        f"[gpreport] Excel built: {row_idx - 2} line-item rows "
        f"(Paid={status_counts['paid']}, Partially Paid={status_counts['partially_paid']}, "
        f"Pending={status_counts['pending']})",
        flush=True,
    )
    return buf, status_counts


# ---------------------------------------------------------------------------
# Slack upload (current 3-step external upload flow; files.upload is
# deprecated by Slack as of March 2025)
# ---------------------------------------------------------------------------
def ensure_bot_in_channel(channel_id: str):
    """files.completeUploadExternal requires bot membership in the target
    channel — chat:write.public does NOT extend to file sharing, per
    Slack's docs. Auto-join public channels so /gpreport works without a
    manual /invite everywhere it can. Private channels can't be
    auto-joined by any bot (Slack platform limitation) — those still need
    a manual invite; we just log that case clearly rather than failing
    silently later at the upload step."""
    resp = requests.post(
        "https://slack.com/api/conversations.join",
        headers={"Authorization": f"Bearer {SLACK_BOT_TOKEN}"},
        data={"channel": channel_id},
        timeout=15,
    )
    data = resp.json()
    if data.get("ok"):
        if data.get("warning") == "already_in_channel":
            print(f"[gpreport] Already a member of channel {channel_id}.", flush=True)
        else:
            print(f"[gpreport] Auto-joined channel {channel_id}.", flush=True)
    else:
        error = data.get("error", "unknown_error")
        if error == "method_not_supported_for_channel_type":
            print(
                f"[gpreport] Channel {channel_id} is private/DM — can't auto-join. "
                f"Bot must be manually /invite'd to this channel.",
                flush=True,
            )
        else:
            print(f"[gpreport] conversations.join for {channel_id} returned: {error}", flush=True)


def upload_to_slack(channel_id: str, file_buf: BytesIO, filename: str, comment: str):
    ensure_bot_in_channel(channel_id)

    file_bytes = file_buf.getvalue()
    print(f"[gpreport] Uploading '{filename}' ({len(file_bytes)} bytes) to Slack channel {channel_id}...", flush=True)

    # Step 1: get an upload URL
    resp = requests.post(
        "https://slack.com/api/files.getUploadURLExternal",
        headers={"Authorization": f"Bearer {SLACK_BOT_TOKEN}"},
        data={"filename": filename, "length": len(file_bytes)},
        timeout=30,
    )
    data = resp.json()
    if not data.get("ok"):
        print(f"[gpreport] ERROR getUploadURLExternal: {data}", flush=True)
        raise RuntimeError(f"getUploadURLExternal failed: {data}")
    upload_url = data["upload_url"]
    file_id = data["file_id"]
    print(f"[gpreport] Got upload URL, file_id={file_id}", flush=True)

    # Step 2: upload the raw bytes
    up_resp = requests.post(upload_url, files={"file": (filename, file_bytes)}, timeout=60)
    up_resp.raise_for_status()
    print("[gpreport] File bytes uploaded to Slack.", flush=True)

    # Step 3: complete the upload and share it to the channel
    complete_resp = requests.post(
        "https://slack.com/api/files.completeUploadExternal",
        headers={
            "Authorization": f"Bearer {SLACK_BOT_TOKEN}",
            "Content-Type": "application/json; charset=utf-8",
        },
        json={
            "files": [{"id": file_id, "title": filename}],
            "channel_id": channel_id,
            "initial_comment": comment,
        },
        timeout=30,
    )
    complete_data = complete_resp.json()
    if not complete_data.get("ok"):
        print(f"[gpreport] ERROR completeUploadExternal: {complete_data}", flush=True)
        if complete_data.get("error") == "channel_not_found":
            raise RuntimeError(
                "I'm not in this channel and can't auto-join it (likely private). "
                "Please /invite the bot to this channel and try again."
            )
        raise RuntimeError(f"completeUploadExternal failed: {complete_data}")
    print(f"[gpreport] Report posted successfully to {channel_id}.", flush=True)


# ---------------------------------------------------------------------------
# Background worker
# ---------------------------------------------------------------------------
def process_report(channel_id: str, start_date, end_date, response_url: str):
    """start_date/end_date are datetime.date objects."""
    start_str = start_date.strftime("%d-%m-%Y")
    end_str = end_date.strftime("%d-%m-%Y")
    print(f"[gpreport] Job started: {start_str} to {end_str} for channel {channel_id}", flush=True)
    try:
        orders = fetch_orders(start_date, end_date)
        buf, status_counts = build_excel(orders)
        filename = f"GP_Report_{start_str}_to_{end_str}.xlsx"
        comment = (
            f"📊 GP Report — {start_str} to {end_str}\n"
            f"Paid: {status_counts['paid']} · "
            f"Partially Paid: {status_counts['partially_paid']} · "
            f"Pending: {status_counts['pending']}"
        )
        upload_to_slack(channel_id, buf, filename, comment)
        print(f"[gpreport] Job finished successfully: {start_str} to {end_str}", flush=True)
    except Exception as e:  # noqa: BLE001
        print(f"[gpreport] Job FAILED: {start_str} to {end_str} — {e}", flush=True)
        requests.post(
            response_url,
            json={
                "response_type": "ephemeral",
                "text": f"❌ Failed to generate GP report: {e}",
            },
            timeout=10,
        )


# ---------------------------------------------------------------------------
# Route
# ---------------------------------------------------------------------------
@app.route("/slack/gp-report", methods=["POST"])
def gp_report():
    if not verify_slack_signature(request):
        print("[gpreport] Rejected: invalid Slack signature.", flush=True)
        return jsonify({"response_type": "ephemeral", "text": "Invalid request signature."}), 401

    text = request.form.get("text", "").strip()
    channel_id = request.form.get("channel_id")
    response_url = request.form.get("response_url")
    print(f"[gpreport] Received command: text='{text}' channel={channel_id}", flush=True)

    usage_error = jsonify(
        {
            "response_type": "ephemeral",
            "text": (
                "⚠️ Usage: `/gpreport DD-MM-YYYY to DD-MM-YYYY`\n"
                "Example: `/gpreport 20-01-2026 to 30-01-2026`"
            ),
        }
    )

    # Strip the word "to" (case-insensitive) wherever it appears, then we
    # should be left with exactly two DD-MM-YYYY tokens.
    parts = [p for p in text.split() if p.lower() != "to"]

    if len(parts) != 2 or not DATE_RE.match(parts[0]) or not DATE_RE.match(parts[1]):
        print(f"[gpreport] Rejected: bad usage, parts={parts}", flush=True)
        return usage_error

    try:
        start_date = datetime.strptime(parts[0], "%d-%m-%Y").date()
        end_date = datetime.strptime(parts[1], "%d-%m-%Y").date()
    except ValueError:
        print(f"[gpreport] Rejected: invalid date format, parts={parts}", flush=True)
        return jsonify({"response_type": "ephemeral", "text": "⚠️ Invalid date. Use DD-MM-YYYY, e.g. 20-01-2026."})

    if end_date < start_date:
        print(f"[gpreport] Rejected: end date before start date ({parts})", flush=True)
        return jsonify(
            {"response_type": "ephemeral", "text": "⚠️ End date is before start date — check the order."}
        )

    threading.Thread(
        target=process_report,
        args=(channel_id, start_date, end_date, response_url),
        daemon=True,
    ).start()

    start_str = start_date.strftime("%d-%m-%Y")
    end_str = end_date.strftime("%d-%m-%Y")
    return jsonify(
        {
            "response_type": "ephemeral",
            "text": f"⏳ Generating GP report for {start_str} → {end_str}... it'll post here shortly.",
        }
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))