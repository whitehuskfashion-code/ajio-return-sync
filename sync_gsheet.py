import logging, time, json, os
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

logger = logging.getLogger(__name__)

SHEET_TAB = "AJIO_RETURN"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive"]

# ── Column indices (0-based) in the DataFrame (12 cols) ──────────────────────
RON_COL_IDX  = 11  # RETURN ORDER NUMBER → sheet col O (14)
DATE_COL_IDX =  5  # Return Created Date → sheet col F (5)  ← was 6/G before restructure

# ── Sheet column indices (0-based) after the 3 manual cols J, K, L ───────────
# df cols 0-8  → sheet cols 0-8   (A-I)   direct 1-to-1
# df col  9    → sheet col  12  M   Return Carrier Name
# df col  10   → sheet col  13  N   Cust Return Reason
# df col  11   → sheet col  14  O   RETURN ORDER NUMBER
# J(9) K(10) L(11) = manual — NEVER overwritten by sync
PROTECTED_SHEET_COLS = {9, 10, 11}


def _df_col_to_sheet_col(df_col_idx: int) -> int:
    if df_col_idx <= 8:
        return df_col_idx
    return df_col_idx + 3   # skip J, K, L


def get_client(creds_json_str):
    creds = Credentials.from_service_account_info(
        json.loads(creds_json_str), scopes=SCOPES)
    return gspread.authorize(creds)


def df_value(val):
    if pd.isna(val) or val is None:
        return ""
    if isinstance(val, datetime):
        return val.strftime("%d-%m-%Y %H:%M")
    return str(val).strip()


def df_row_to_sheet_list(row, existing_sheet_row: list) -> list:
    """
    Build a full 15-col sheet row from a 12-col df row.
    Cols J, K, L (indices 9,10,11) are taken from the EXISTING sheet row
    so we never overwrite manual data.
    """
    result = [""] * 15
    for df_idx in range(12):
        sheet_idx = _df_col_to_sheet_col(df_idx)
        result[sheet_idx] = df_value(row.iloc[df_idx])
    for protected in PROTECTED_SHEET_COLS:
        if protected < len(existing_sheet_row):
            result[protected] = existing_sheet_row[protected]
    return result


def sort_sheet(ws):
    """Sort AJIO_RETURN by Return Created Date (col F, index 5) ascending."""
    ws.spreadsheet.batch_update({"requests": [{"sortRange": {
        "range": {"sheetId": ws.id, "startRowIndex": 1},
        "sortSpecs": [{"dimensionIndex": DATE_COL_IDX, "sortOrder": "ASCENDING"}]
    }}]})
    logger.info("Sheet sorted by Return Created Date ✅")


def sync(df, sheet_id, creds_json_str):
    logger.info("Connecting to Google Sheets...")
    gc = get_client(creds_json_str)
    ws = gc.open_by_key(sheet_id).worksheet(SHEET_TAB)
    existing = ws.get_all_values()

    # ── Empty sheet: first-time write ────────────────────────────────────────
    def _is_real_row(row):
        return any(v.strip() and v.strip().lower() != "select" for v in row)

    real_data_rows = [r for r in existing[1:] if _is_real_row(r)]
    if len(existing) <= 1 or len(real_data_rows) == 0:
        header = list(df.columns)
        full_header = header[:9] + ["Actual Delivered Date", "Quality", "Notes"] + header[9:]
        rows = []
        for i in range(len(df)):
            r = df.iloc[i]
            row_data  = [df_value(r.iloc[j]) for j in range(9)]
            row_data += ["", "", ""]
            row_data += [df_value(r.iloc[j]) for j in range(9, 12)]
            rows.append(row_data)
        ws.update("A1", [full_header] + rows, value_input_option="USER_ENTERED")
        sort_sheet(ws)
        logger.info(f"Sheet was empty. Wrote {len(df)} rows.")
        return

    # ── Existing sheet ────────────────────────────────────────────────────────
    header_row = existing[0]
    data_rows  = existing[1:]

    try:
        ron_sheet_col = header_row.index("RETURN ORDER NUMBER")
    except ValueError:
        ron_sheet_col = 14

    ron_to_row = {}
    for idx, row in enumerate(data_rows):
        ron = row[ron_sheet_col].strip() if ron_sheet_col < len(row) else ""
        if ron:
            ron_to_row[ron] = idx + 2   # 1-based sheet row

    updates  = []
    new_rows = []

    for i in range(len(df)):
        row_s   = df.iloc[i]
        ron_val = df_value(row_s.iloc[RON_COL_IDX])
        if not ron_val:
            continue
        if ron_val in ron_to_row:
            sheet_row_num = ron_to_row[ron_val]
            existing_data = data_rows[sheet_row_num - 2]
            full_vals = df_row_to_sheet_list(row_s, existing_data)
            updates.append((sheet_row_num, full_vals))
        else:
            new_rows.append((row_s.iloc[DATE_COL_IDX], row_s))

    logger.info(f"Updates: {len(updates)} | Inserts: {len(new_rows)}")

    if updates:
        reqs = [{"range": f"A{r}", "values": [v]} for r, v in updates]
        for start in range(0, len(reqs), 500):
            ws.batch_update(reqs[start:start+500], value_input_option="USER_ENTERED")
            time.sleep(1)

    if new_rows:
        sorted_new = sorted(
            new_rows,
            key=lambda x: x[0] if isinstance(x[0], datetime) and not pd.isna(x[0]) else datetime.min
        )
        append_data = []
        for _, row_s in sorted_new:
            row_data  = [df_value(row_s.iloc[j]) for j in range(9)]
            row_data += ["", "", ""]
            row_data += [df_value(row_s.iloc[j]) for j in range(9, 12)]
            append_data.append(row_data)
        last_real_row = 1
        for idx, row in enumerate(existing):
            if idx == 0:
                continue
            if _is_real_row(row):
                last_real_row = idx + 1
        next_row = last_real_row + 1
        ws.update(f"A{next_row}", append_data, value_input_option="USER_ENTERED")

    sort_sheet(ws)
    logger.info("Sync complete ✅")
