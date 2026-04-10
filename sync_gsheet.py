import logging, time, json, os
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

logger = logging.getLogger(__name__)

SHEET_TAB   = "AJIO_RETURN"
SCOPES      = ["https://www.googleapis.com/auth/spreadsheets",
               "https://www.googleapis.com/auth/drive"]

# ── Column indices (0-based) in the Excel/DataFrame (12 cols from macro) ──────
RON_COL_IDX  = 11   # RETURN ORDER NUMBER  (col L in df, col O in sheet after +3)
DATE_COL_IDX = 6    # Return Created Date  (col G, same in both)

# ── Sheet column indices (0-based) after the 3 manual columns J,K,L ──────────
# J=9, K=10, L=11 are the 3 manual cols → NEVER written by sync
# The df has 12 cols (A-L). In the sheet they map to A-I then skip J,K,L then M-O
# df col 0-8  → sheet col 0-8   (A-I)
# df col 9    → sheet col 12    (M)  ← Return Carrier Name
# df col 10   → sheet col 13    (N)  ← Cust Return Reason
# df col 11   → sheet col 14    (O)  ← RETURN ORDER NUMBER
PROTECTED_SHEET_COLS = {9, 10, 11}   # J, K, L (0-based) — never overwrite

def _df_col_to_sheet_col(df_col_idx: int) -> int:
    """Map 0-based df column index → 0-based sheet column index."""
    if df_col_idx <= 8:
        return df_col_idx
    return df_col_idx + 3   # shift past J, K, L

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
    Build a full sheet row (15 cols A-O) from a df row (12 cols).
    Cols J, K, L (indices 9,10,11) are taken from the EXISTING sheet row
    so we never overwrite manual data.
    existing_sheet_row: current values from the sheet (may be shorter than 15)
    """
    result = [""] * 15
    for df_idx in range(12):
        sheet_idx = _df_col_to_sheet_col(df_idx)
        result[sheet_idx] = df_value(row.iloc[df_idx])
    # Restore protected cols J(9), K(10), L(11) from existing sheet data
    for protected in PROTECTED_SHEET_COLS:
        if protected < len(existing_sheet_row):
            result[protected] = existing_sheet_row[protected]
    return result

def sort_sheet(ws):
    """Sort AJIO_RETURN by col G (index 6) ascending, keeping header."""
    ws.spreadsheet.batch_update({"requests": [{"sortRange": {
        "range": {"sheetId": ws.id, "startRowIndex": 1},
        "sortSpecs": [{"dimensionIndex": DATE_COL_IDX, "sortOrder": "ASCENDING"}]
    }}]})
    logger.info("Sheet sorted by Return Created Date ✅")

def sync(df, sheet_id, creds_json_str):
    logger.info("Connecting to Google Sheets...")
    gc  = get_client(creds_json_str)
    ws  = gc.open_by_key(sheet_id).worksheet(SHEET_TAB)

    existing = ws.get_all_values()

    # ── Empty sheet: first-time write ─────────────────────────────────────────
    if not existing:
        header = list(df.columns)
        # Extend header with 3 manual cols + shifted cols
        full_header = header[:9] + ["Actual Delivered Date", "Quality", "Notes"] + header[9:]
        empty_protected = ["", "", ""]
        rows = []
        for i in range(len(df)):
            r = df.iloc[i]
            row_data = [df_value(r.iloc[j]) for j in range(9)]
            row_data += empty_protected
            row_data += [df_value(r.iloc[j]) for j in range(9, 12)]
            rows.append(row_data)
        ws.update("A1", [full_header] + rows, value_input_option="USER_ENTERED")
        sort_sheet(ws)
        logger.info(f"Sheet was empty. Wrote {len(df)} rows.")
        return

    # ── Existing sheet ─────────────────────────────────────────────────────────
    header_row = existing[0]
    data_rows  = existing[1:]

    # Find RETURN ORDER NUMBER column in sheet (col O = index 14)
    try:
        ron_sheet_col = header_row.index("RETURN ORDER NUMBER")
    except ValueError:
        ron_sheet_col = 14

    ron_to_row = {}
    for idx, row in enumerate(data_rows):
        ron = row[ron_sheet_col].strip() if ron_sheet_col < len(row) else ""
        if ron:
            ron_to_row[ron] = idx + 2   # 1-based sheet row number

    updates  = []
    new_rows = []

    for i in range(len(df)):
        row_s   = df.iloc[i]
        ron_val = df_value(row_s.iloc[RON_COL_IDX])
        if not ron_val:
            continue

        if ron_val in ron_to_row:
            sheet_row_num = ron_to_row[ron_val]
            existing_data = data_rows[sheet_row_num - 2]   # 0-based
            full_vals     = df_row_to_sheet_list(row_s, existing_data)
            updates.append((sheet_row_num, full_vals))
        else:
            new_rows.append((row_s.iloc[DATE_COL_IDX], row_s))

    logger.info(f"Updates: {len(updates)} | Inserts: {len(new_rows)}")

    # ── Batch update existing rows (skip protected cols via full row rebuild) ──
    if updates:
        reqs = [{"range": f"A{r}", "values": [v]} for r, v in updates]
        for start in range(0, len(reqs), 500):
            ws.batch_update(reqs[start:start+500], value_input_option="USER_ENTERED")
            time.sleep(1)

    # ── Append new rows ────────────────────────────────────────────────────────
    if new_rows:
        sorted_new = sorted(
            new_rows,
            key=lambda x: x[0] if isinstance(x[0], datetime) and not pd.isna(x[0]) else datetime.min
        )
        append_data = []
        for _, row_s in sorted_new:
            row_data = [df_value(row_s.iloc[j]) for j in range(9)]
            row_data += ["", "", ""]          # J, K, L empty for new rows
            row_data += [df_value(row_s.iloc[j]) for j in range(9, 12)]
            append_data.append(row_data)
        ws.append_rows(append_data, value_input_option="USER_ENTERED")

    # ── Always sort after any change ───────────────────────────────────────────
    sort_sheet(ws)
    logger.info("Sync complete ✅")
