import os, sys, logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(name)s — %(message)s",
    handlers=[logging.StreamHandler(sys.stdout), logging.FileHandler("run.log")]
)
logger = logging.getLogger("main")

from ajio_scraper  import run_scraper
from process_excel import process_both
from sync_gsheet   import sync
from alerts        import run_alerts
from datetime import datetime

def env(key):
    v = os.environ.get(key, "").strip()
    if not v: logger.error(f"Missing env: {key}"); sys.exit(1)
    return v

def main():
    logger.info("=" * 55)
    logger.info("AJIO Return Sync — Starting")
    username   = env("AJIO_USERNAME")
    password   = env("AJIO_PASSWORD")
    sheet_id   = env("GOOGLE_SHEET_ID")
    creds_json = env("GOOGLE_CREDENTIALS_JSON")

    logger.info("STEP 1 — Downloading from Ajio...")
    try:
        path_a, path_b = run_scraper(username, password, "downloads")
        excel_files = [p for p in [path_a, path_b] if p is not None]
        logger.info(f"Will process {len(excel_files)} file(s)")
    except Exception as e:
        logger.exception(f"Scraping failed: {e}"); sys.exit(1)

    logger.info("STEP 2 — Processing Excel (macro equivalent)...")
    try:
        df = process_both(path_a, path_b)
        os.makedirs("output", exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        df.to_excel(f"output/processed_{stamp}.xlsx", index=False)
        logger.info(f"  {len(df)} rows processed")
    except Exception as e:
        logger.exception(f"Processing failed: {e}"); sys.exit(1)

    logger.info("STEP 3 — Syncing to Google Sheet...")
    try:
        sync(df, sheet_id, creds_json)
    except Exception as e:
        logger.exception(f"Sync failed: {e}"); sys.exit(1)

    logger.info("STEP 4 — Running alerts...")
    try:
        run_alerts(
            sheet_id       = sheet_id,
            creds_json_str = creds_json,
            email_sender   = env("EMAIL_SENDER"),
            email_password = env("EMAIL_APP_PASSWORD"),
            email_to_str   = env("EMAIL_TO"),
        )
    except Exception as e:
        logger.exception(f"Alerts failed: {e}")
        # ← no sys.exit here — alerts failing should NOT kill the whole run

    logger.info("DONE ✅")

if __name__ == "__main__":
    main()
