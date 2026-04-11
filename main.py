import os, sys, logging, urllib.request
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(name)s — %(message)s",
    handlers=[logging.StreamHandler(sys.stdout), logging.FileHandler("run.log")]
)
logger = logging.getLogger("main")

from ajio_scraper  import run_scraper
from process_excel import process_both
from sync_gsheet   import sync
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

    logger.info("STEP 4 — Triggering Apps Script alerts...")
    apps_script_url = os.environ.get("APPS_SCRIPT_URL", "").strip()
    if not apps_script_url:
        logger.warning("APPS_SCRIPT_URL not set — skipping alerts trigger")
    else:
        try:
            req = urllib.request.Request(apps_script_url, method="POST", data=b"trigger")
            with urllib.request.urlopen(req, timeout=30) as resp:
                body = resp.read().decode()
            logger.info(f"  Apps Script response: {body}")
        except Exception as e:
            logger.warning(f"  Apps Script trigger failed (non-fatal): {e}")

    logger.info("DONE ✅")

if __name__ == "__main__":
    main()
