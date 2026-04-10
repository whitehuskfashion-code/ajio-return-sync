import logging
import pandas as pd
from datetime import datetime

logger = logging.getLogger(__name__)

SRC_COLS_0BASED = [20, 2, 7, 37, 10, 5, 11, 12, 9, 22, 0]
DATE_SRC_F_IDX  = 5
DATE_SRC_M_IDX  = 12
FINAL_DATE_COLS = [0, 6, 8]

def parse_java_date(s: str):
    try:
        parts = str(s).strip().split()
        if len(parts) < 6:
            return None
        dt_str = f"{parts[2]} {parts[1]} {parts[5]} {parts[3]}"
        return datetime.strptime(dt_str, "%d %b %Y %H:%M:%S")
    except Exception:
        return None

def process_single(input_path: str) -> pd.DataFrame:
    logger.info(f"Processing: {input_path}")
    df = pd.read_excel(input_path, header=0, dtype=str)
    logger.info(f"  Shape: {df.shape}")
    cols = df.columns.tolist()
    max_needed = max(SRC_COLS_0BASED)
    if len(cols) <= max_needed:
        logger.warning(f"  Only {len(cols)} cols found, padding to {max_needed+1}")
        for i in range(len(cols), max_needed + 1):
            df[f"_pad_{i}"] = ""
        cols = df.columns.tolist()
    col_f = cols[DATE_SRC_F_IDX]
    col_m = cols[DATE_SRC_M_IDX]
    df[col_f] = df[col_f].apply(
        lambda x: parse_java_date(x) if pd.notna(x) and str(x).strip() not in ("", "nan") else None
    )
    df[col_m] = df[col_m].apply(
        lambda x: parse_java_date(x) if pd.notna(x) and str(x).strip() not in ("", "nan") else None
    )
    df = df.sort_values(by=col_f, ascending=True, na_position="last").reset_index(drop=True)
    selected = [cols[i] for i in SRC_COLS_0BASED]
    df_out = df[selected].copy()
    df_out.insert(0, "LAST_UPDATE_DATE", datetime.now().replace(microsecond=0))
    for i, col in enumerate(df_out.columns):
        if i not in FINAL_DATE_COLS:
            df_out[col] = df_out[col].astype(str).str.strip().replace("nan", "")
    logger.info(f"  Processed rows: {len(df_out)}")
    return df_out

def combine_runs(df_a: pd.DataFrame, df_b: pd.DataFrame) -> pd.DataFrame:
    combined = pd.concat([df_b, df_a], ignore_index=True)
    ron_col  = combined.columns[11]
    date_col = combined.columns[7]
    combined_sorted = combined.sort_values(by=date_col, ascending=True, na_position="last")
    deduped = combined_sorted.drop_duplicates(subset=[ron_col], keep="last")
    deduped = deduped.sort_values(by=date_col, ascending=True, na_position="last").reset_index(drop=True)
    logger.info(f"Combined: {len(combined)} rows → {len(deduped)} unique orders")
    return deduped

def process_both(path_a: str, path_b: str) -> pd.DataFrame:
    df_a = process_single(path_a) if path_a is not None else None
    df_b = process_single(path_b) if path_b is not None else None

    if df_a is not None and df_b is not None:
        return combine_runs(df_a, df_b)
    elif df_a is not None:
        logger.info("Only RunA has data — using RunA only")
        return df_a
    elif df_b is not None:
        logger.info("Only RunB has data — using RunB only")
        return df_b
    else:
        raise RuntimeError("Both RunA and RunB processed as None — nothing to sync!")
