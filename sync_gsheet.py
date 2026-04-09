import logging, time, json, os
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

logger = logging.getLogger(__name__)
SHEET_TAB   = "AJIO_RETURN"
SCOPES      = ["https://www.googleapis.com/auth/spreadsheets",
               "https://www.googleapis.com/auth/drive"]
RON_COL_IDX  = 11
DATE_COL_IDX = 7

def get_client(creds_json_str):
    creds = Credentials.from_service_account_info(json.loads(creds_json_str), scopes=SCOPES)
    return gspread.authorize(creds)

def df_value(val):
    if pd.isna(val) or val is None: return ""
    if isinstance(val, datetime): return val.strftime("%d-%m-%Y %H:%M")
    return str(val).strip()

def df_row_to_list(row): return [df_value(v) for v in row]

def sync(df, sheet_id, creds_json_str):
    logger.info("Connecting to Google Sheets...")
    gc = get_client(creds_json_str)
    ws = gc.open_by_key(sheet_id).worksheet(SHEET_TAB)

    existing = ws.get_all_values()
    if not existing:
        header = list(df.columns)
        ws.update("A1", [header] + [df_row_to_list(df.iloc[i]) for i in range(len(df))],
                  value_input_option="USER_ENTERED")
        logger.info(f"Sheet was empty. Wrote {len(df)} rows.")
        return

    header_row = existing[0]
    data_rows  = existing[1:]
    try:    ron_sheet_col = header_row.index("RETURN ORDER NUMBER")
    except: ron_sheet_col = 11

    ron_to_row = {}
    for idx, row in enumerate(data_rows):
        ron = row[ron_sheet_col].strip() if ron_sheet_col < len(row) else ""
        if ron: ron_to_row[ron] = idx + 2

    updates, new_rows = [], []
    for i in range(len(df)):
        row_s   = df.iloc[i]
        ron_val = df_value(row_s.iloc[RON_COL_IDX])
        vals    = df_row_to_list(row_s)
        if ron_val in ron_to_row:
            updates.append((ron_to_row[ron_val], vals))
        else:
            new_rows.append((row_s.iloc[DATE_COL_IDX], vals))

    logger.info(f"Updates: {len(updates)} | Inserts: {len(new_rows)}")

    if updates:
        reqs = [{"range": f"A{r}", "values": [v]} for r, v in updates]
        for start in range(0, len(reqs), 500):
            ws.batch_update(reqs[start:start+500], value_input_option="USER_ENTERED")
            time.sleep(1)

    if new_rows:
        sorted_new = sorted(new_rows,
            key=lambda x: x[0] if isinstance(x[0], datetime) and not pd.isna(x[0]) else datetime.min)
        ws.append_rows([r[1] for r in sorted_new], value_input_option="USER_ENTERED")
        ws.spreadsheet.batch_update({"requests": [{"sortRange": {
            "range": {"sheetId": ws.id, "startRowIndex": 1},
            "sortSpecs": [{"dimensionIndex": DATE_COL_IDX, "sortOrder": "ASCENDING"}]
        }}]})

    logger.info("Sync complete ✅")