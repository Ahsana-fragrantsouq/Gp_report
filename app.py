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
   - Builds a styled .xlsx with columns:
     Name, Created at, Fulfilled at, Lineitem name, Lineitem sku,
     Lineitem price, Lineitem quantity, Shipping, Taxes, Discount
   - Uploads the file to the requesting Slack channel using Slack's
     current 3-step external upload flow (files.upload is deprecated).
"""
import hashlib
import hmac
import os
import re
import threading
import time
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
]

DATE_RE = re.compile(r"^\d{2}-\d{2}-\d{4}$")  # DD-MM-YYYY


# ---------------------------------------------------------------------------
# Slack request verification (optional, but recommended)
# ---------------------------------------------------------------------------
def verify_slack_signature(req) -> bool:
    if not SLACK_SIGNING_SECRET:
        return True  # verification skipped if secret not configured

    timestamp = req.headers.get("X-Slack-Request-Timestamp", "")
    if not timestamp or abs(time.time() - int(timestamp)) > 60 * 5:
        return False

    sig_basestring = f"v0:{timestamp}:{req.get_data(as_text=True)}"
    my_sig = (
        "v0="
        + hmac.new(
            SLACK_SIGNING_SECRET.encode(),
            sig_basestring.encode(),
            hashlib.sha256,
        ).hexdigest()
    )
    slack_sig = req.headers.get("X-Slack-Signature", "")
    return hmac.compare_digest(my_sig, slack_sig)


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

    status_counts = {"paid": 0, "partially_paid": 0, "pending": 0}
    row_idx = 2
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
        discount_total = order.get("total_discounts", "0.00")

        for li in line_items:
            values = [
                name,
                status_label,
                created_at,
                fulfilled_at,
                li.get("name", ""),
                li.get("sku", ""),
                li.get("price", ""),
                li.get("quantity", ""),
                shipping_total,
                taxes_total,
                discount_total,
            ]
            for col_idx, value in enumerate(values, start=1):
                ws.cell(row=row_idx, column=col_idx, value=value)
            row_idx += 1

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
def upload_to_slack(channel_id: str, file_buf: BytesIO, filename: str, comment: str):
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
    text = request.form.get("text", "").strip()
    channel_id = request.form.get("channel_id")
    response_url = request.form.get("response_url")
    print(f"[gpreport] Received command: text='{text}' channel={channel_id}", flush=True)

    if not verify_slack_signature(request):
        print("[gpreport] Rejected: invalid Slack signature.", flush=True)
        return jsonify({"response_type": "ephemeral", "text": "Invalid request signature."}), 401

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