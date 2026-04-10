import logging, json, os, csv, io
from datetime import datetime, date, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import gspread
from google.oauth2.service_account import Credentials

logger = logging.getLogger(__name__)

# ── Constants ──────────────────────────────────────────────────────────────────
SHEET_TAB     = "AJIO_RETURN"
TRACKER_TAB   = "AJIO_TICKETS"
SCOPES        = ["https://www.googleapis.com/auth/spreadsheets",
                 "https://www.googleapis.com/auth/drive"]

# ── Sheet column indices (0-based) AFTER 3 manual cols ────────────────────────
COL_SELLER_SKU          = 2   # C
COL_CUST_ORDER          = 4   # E
COL_RETURN_AWB          = 5   # F
COL_RETURN_CREATED      = 6   # G
COL_RETURN_DELIVERED    = 8   # I
COL_ACTUAL_DELIVERED    = 9   # J  ← manual
COL_RETURN_CARRIER      = 12  # M
COL_RETURN_ORDER_NUM    = 14  # O

# ── Highlight colours ──────────────────────────────────────────────────────────
PURPLE_BG = {"red": 0.6,  "green": 0.0,  "blue": 0.9}   # Alert 1
ORANGE_BG = {"red": 1.0,  "green": 0.6,  "blue": 0.0}   # Alert 2
NO_BG     = {"red": 1.0,  "green": 1.0,  "blue": 1.0}   # Clear highlight

ALERT_DELIVERED = "DELIVERED_NOT_RECEIVED"
ALERT_60_DAYS   = "60_DAYS_NO_DELIVERY"


def get_client(creds_json_str):
    creds = Credentials.from_service_account_info(
        json.loads(creds_json_str), scopes=SCOPES)
    return gspread.authorize(creds)


def _parse_date(val: str):
    """Try to parse date from various formats. Returns date or None."""
    if not val or not val.strip():
        return None
    for fmt in ("%d-%m-%Y %H:%M", "%d-%m-%Y", "%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(val.strip(), fmt).date()
        except ValueError:
            continue
    return None


def _col_letter(idx: int) -> str:
    """0-based column index → A, B, ... Z, AA, ..."""
    result = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        result = chr(65 + rem) + result
    return result


def _set_row_bg(ws, sheet_id_int: int, row_num: int, color: dict, total_cols: int = 15):
    """Set background colour for an entire row (1-based row_num)."""
    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id_int,
                "startRowIndex": row_num - 1,
                "endRowIndex": row_num,
                "startColumnIndex": 0,
                "endColumnIndex": total_cols,
            },
            "cell": {"userEnteredFormat": {"backgroundColor": color}},
            "fields": "userEnteredFormat.backgroundColor",
        }
    }


def _read_tracker(tracker_ws) -> dict:
    """
    Returns dict: {RETURN_ORDER_NUMBER: {"alert_type": ..., "date_identified": ...,
                                          "ticket_date": ..., "row_num": int}}
    """
    rows = tracker_ws.get_all_values()
    if len(rows) <= 1:
        return {}
    header = rows[0]
    try:
        col_ron    = header.index("RETURN_ORDER_NUMBER")
        col_alert  = header.index("ALERT_TYPE")
        col_date   = header.index("DATE_IDENTIFIED")
        col_ticket = header.index("TICKET_CREATED_DATE")
    except ValueError:
        logger.error("AJIO_TICKETS headers not found!")
        return {}

    result = {}
    for i, row in enumerate(rows[1:], start=2):
        ron = row[col_ron].strip() if col_ron < len(row) else ""
        if not ron:
            continue
        result[ron] = {
            "alert_type":      row[col_alert].strip()  if col_alert  < len(row) else "",
            "date_identified": row[col_date].strip()   if col_date   < len(row) else "",
            "ticket_date":     row[col_ticket].strip() if col_ticket < len(row) else "",
            "row_num":         i,
        }
    return result


def _append_tracker(tracker_ws, ron: str, alert_type: str):
    tracker_ws.append_rows(
        [[ron, alert_type, date.today().strftime("%Y-%m-%d"), ""]],
        value_input_option="USER_ENTERED"
    )


def _build_email_html(rows_data: list, columns: list, alert_type: str) -> str:
    if alert_type == ALERT_DELIVERED:
        subject_line = "Ajio Order Marked Delivered but Not Received – Immediate raise ticket"
        body_text = ("The following order(s) are marked as delivered in the Ajio system; "
                     "however, they have not been received by us.<br><br>"
                     "Kindly investigate this issue and raise a ticket at the earliest.")
    else:
        subject_line = "Ajio Alert – 60 Days Passed, Return Not Received"
        body_text = ("It has been over 60 days since the Return Creation Date, yet the "
                     "Return Delivered Date is still not updated.<br><br>"
                     "Kindly raise a ticket immediately to investigate and resolve this issue.")

    header_html = "".join(f"<th style='border:1px solid #ccc;padding:8px;background:#f0f0f0'>{c}</th>" for c in columns)
    rows_html   = ""
    for row in rows_data:
        cells = "".join(f"<td style='border:1px solid #ccc;padding:8px'>{v}</td>" for v in row)
        rows_html += f"<tr>{cells}</tr>"

    return f"""
    <html><body>
    <p>{body_text}</p>
    <table style='border-collapse:collapse;font-family:Arial,sans-serif;font-size:13px'>
      <thead><tr>{header_html}</tr></thead>
      <tbody>{rows_html}</tbody>
    </table>
    </body></html>
    """, subject_line


def _build_csv(rows_data: list, columns: list) -> bytes:
    buf = io.StringIO()
    w   = csv.writer(buf)
    w.writerow(columns)
    w.writerows(rows_data)
    return buf.getvalue().encode("utf-8")


def _send_email(sender: str, app_password: str, recipients: list,
                subject: str, html_body: str, csv_bytes: bytes, csv_filename: str):
    msg = MIMEMultipart("mixed")
    msg["From"]    = sender
    msg["To"]      = ", ".join(recipients)
    msg["Subject"] = subject

    msg.attach(MIMEText(html_body, "html"))

    part = MIMEBase("application", "octet-stream")
    part.set_payload(csv_bytes)
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{csv_filename}"')
    msg.attach(part)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender, app_password)
        server.sendmail(sender, recipients, msg.as_string())
    logger.info(f"Email sent to {recipients}: {subject}")


def run_alerts(sheet_id: str, creds_json_str: str,
               email_sender: str, email_password: str, email_to_str: str):
    """
    Main entry point called from main.py after sync().
    email_to_str: comma-separated email addresses from GitHub Secret EMAIL_TO
    """
    recipients = [e.strip() for e in email_to_str.split(",") if e.strip()]
    today      = date.today()

    gc           = get_client(creds_json_str)
    spreadsheet  = gc.open_by_key(sheet_id)
    ws           = spreadsheet.worksheet(SHEET_TAB)
    tracker_ws   = spreadsheet.worksheet(TRACKER_TAB)
    sheet_id_int = ws.id

    all_rows    = ws.get_all_values()
    header_row  = all_rows[0] if all_rows else []
    data_rows   = all_rows[1:] if len(all_rows) > 1 else []

    tracker = _read_tracker(tracker_ws)

    highlight_requests = []

    # ══════════════════════════════════════════════════════════════════════════
    # STEP 1 — CLEANUP: remove highlights where ticket has been created
    # ══════════════════════════════════════════════════════════════════════════
    logger.info("Step 1: Cleanup — removing highlights for resolved tickets...")
    ron_to_sheet_row = {}
    for idx, row in enumerate(data_rows, start=2):
        ron = row[COL_RETURN_ORDER_NUM].strip() if COL_RETURN_ORDER_NUM < len(row) else ""
        if ron:
            ron_to_sheet_row[ron] = idx

    for ron, info in tracker.items():
        if info["ticket_date"]:   # ticket created → remove highlight
            if ron in ron_to_sheet_row:
                highlight_requests.append(
                    _set_row_bg(ws, sheet_id_int, ron_to_sheet_row[ron], NO_BG)
                )
                logger.info(f"  Cleared highlight for {ron} (ticket: {info['ticket_date']})")

    # ══════════════════════════════════════════════════════════════════════════
    # STEP 2 — ALERT 1: Delivered but Not Received (Purple)
    # Window: today-7 to today-3 (inclusive), col J (Actual Delivered) empty
    # ══════════════════════════════════════════════════════════════════════════
    logger.info("Step 2: Alert 1 — Delivered Not Received (Purple)...")
    window_start = today - timedelta(days=7)
    window_end   = today - timedelta(days=3)

    alert1_rows  = []
    alert1_cols  = ["SELLER SKU", "Cust Order No", "Return AWB No",
                    "Return Delivered Date", "Return Carrier Name", "RETURN ORDER NUMBER"]
    col_indices_1 = [COL_SELLER_SKU, COL_CUST_ORDER, COL_RETURN_AWB,
                     COL_RETURN_DELIVERED, COL_RETURN_CARRIER, COL_RETURN_ORDER_NUM]

    for idx, row in enumerate(data_rows, start=2):
        delivered_str = row[COL_RETURN_DELIVERED].strip() if COL_RETURN_DELIVERED < len(row) else ""
        actual_str    = row[COL_ACTUAL_DELIVERED].strip()  if COL_ACTUAL_DELIVERED  < len(row) else ""
        ron           = row[COL_RETURN_ORDER_NUM].strip()  if COL_RETURN_ORDER_NUM  < len(row) else ""

        delivered_dt  = _parse_date(delivered_str)
        if not delivered_dt:
            continue
        if not (window_start <= delivered_dt <= window_end):
            continue
        if actual_str:   # Actual Delivered Date filled → already received
            continue

        # Check tracker
        tracker_info = tracker.get(ron)
        if tracker_info and tracker_info["ticket_date"]:
            continue   # ticket created → skip (highlight already cleared in step 1)

        # Qualify → highlight purple
        highlight_requests.append(_set_row_bg(ws, sheet_id_int, idx, PURPLE_BG))

        if not tracker_info:
            _append_tracker(tracker_ws, ron, ALERT_DELIVERED)
            tracker[ron] = {"alert_type": ALERT_DELIVERED, "ticket_date": ""}

        # Collect row data for email
        row_vals = [row[c] if c < len(row) else "" for c in col_indices_1]
        alert1_rows.append(row_vals)
        logger.info(f"  Alert1: {ron} delivered {delivered_str}")

    # Send Alert 1 email
    if alert1_rows:
        html, subject = _build_email_html(alert1_rows, alert1_cols, ALERT_DELIVERED)
        csv_bytes = _build_csv(alert1_rows, alert1_cols)
        _send_email(email_sender, email_password, recipients,
                    subject, html, csv_bytes,
                    f"alert_delivered_not_received_{today}.csv")
    else:
        logger.info("  No Alert 1 rows today.")

    # ══════════════════════════════════════════════════════════════════════════
    # STEP 3 — ALERT 2: 60 Days No Delivery (Orange)
    # Col G > 60 days ago AND col I (Return Delivered Date) empty
    # ══════════════════════════════════════════════════════════════════════════
    logger.info("Step 3: Alert 2 — 60 Days No Delivery (Orange)...")
    cutoff_60 = today - timedelta(days=60)

    alert2_rows = []
    alert2_cols = ["SELLER SKU", "Cust Order No", "Return AWB No",
                   "Return Created Date", "Return Delivered Date",
                   "Return Carrier Name", "RETURN ORDER NUMBER"]
    col_indices_2 = [COL_SELLER_SKU, COL_CUST_ORDER, COL_RETURN_AWB,
                     COL_RETURN_CREATED, COL_RETURN_DELIVERED,
                     COL_RETURN_CARRIER, COL_RETURN_ORDER_NUM]

    for idx, row in enumerate(data_rows, start=2):
        created_str   = row[COL_RETURN_CREATED].strip()   if COL_RETURN_CREATED   < len(row) else ""
        delivered_str = row[COL_RETURN_DELIVERED].strip() if COL_RETURN_DELIVERED < len(row) else ""
        ron           = row[COL_RETURN_ORDER_NUM].strip() if COL_RETURN_ORDER_NUM < len(row) else ""

        created_dt = _parse_date(created_str)
        if not created_dt:
            continue
        if created_dt > cutoff_60:   # not yet 60 days
            continue
        if delivered_str:            # already has delivery date
            continue

        tracker_info = tracker.get(ron)
        if tracker_info and tracker_info["ticket_date"]:
            continue   # ticket created → skip

        # Qualify → highlight orange
        highlight_requests.append(_set_row_bg(ws, sheet_id_int, idx, ORANGE_BG))

        if not tracker_info:
            _append_tracker(tracker_ws, ron, ALERT_60_DAYS)
            tracker[ron] = {"alert_type": ALERT_60_DAYS, "ticket_date": ""}

        row_vals = [row[c] if c < len(row) else "" for c in col_indices_2]
        alert2_rows.append(row_vals)
        logger.info(f"  Alert2: {ron} created {created_str} (>{60} days, no delivery)")

    # Send Alert 2 email
    if alert2_rows:
        html, subject = _build_email_html(alert2_rows, alert2_cols, ALERT_60_DAYS)
        csv_bytes = _build_csv(alert2_rows, alert2_cols)
        _send_email(email_sender, email_password, recipients,
                    subject, html, csv_bytes,
                    f"alert_60days_no_delivery_{today}.csv")
    else:
        logger.info("  No Alert 2 rows today.")

    # ══════════════════════════════════════════════════════════════════════════
    # STEP 4 — Apply all highlight changes in one batch request
    # ══════════════════════════════════════════════════════════════════════════
    if highlight_requests:
        ws.spreadsheet.batch_update({"requests": highlight_requests})
        logger.info(f"Applied {len(highlight_requests)} highlight changes ✅")

    logger.info("Alerts run complete ✅")
