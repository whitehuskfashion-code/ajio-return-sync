import os, time, logging
from datetime import datetime, timedelta
from selenium import webdriver
#from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

logger = logging.getLogger(__name__)
LOGIN_URL   = "https://seller.ajio.com/ajiocommerce/"
REPORTS_URL = "https://seller.ajio.com/vmsui/reports/ViewReports"

def build_driver(download_dir: str) -> webdriver.Chrome:
    abs_dl = os.path.abspath(download_dir)
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1920,1080")

    # ── Anti-bot detection ────────────────────────────────────
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    opts.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/147.0.0.0 Safari/537.36"
    )

    opts.add_experimental_option("prefs", {
        "download.default_directory": abs_dl,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    })
    driver = webdriver.Chrome(options=opts)

    # Mask webdriver property via JS
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    })
    return driver

def screenshot(driver, name):
    try: driver.save_screenshot(f"downloads/debug_{name}.png")
    except: pass

def login(driver, username: str, password: str):
    logger.info("Opening landing page...")
    driver.get(LOGIN_URL)

    # Screenshot immediately — tells us what headless actually sees
    time.sleep(5)
    screenshot(driver, "00_landing_page")
    logger.info(f"Landing page URL: {driver.current_url}")
    logger.info(f"Page title: {driver.title}")

    wait = WebDriverWait(driver, 45)

    # ── Step 1: Click "Log in" button (top-right of landing page) ──
    # "Log in →" — use contains() to handle the arrow character
    clicked = False
    # ── Step 1: Click "Log in" button ────────────────────────
    # Button structure: <button id="ui-btn"><span>Log in</span><img/></button>
    # Text is in child <span>, NOT directly on button — must use .// to search descendants
    clicked = False
    selectors_to_try = [
        (By.XPATH,        "//button[.//span[normalize-space(text())='Log in']]"),
        (By.XPATH,        "//button[@id='ui-btn'][contains(@class,'ui-btn-yellow')]"),
        (By.CSS_SELECTOR, "button.ui-btn-yellow"),
        (By.CSS_SELECTOR, "button#ui-btn"),
        (By.XPATH,        "//button[contains(@class,'ui-btn')][1]"),
    ]

    for sel_type, sel_val in selectors_to_try:
        try:
            el = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((sel_type, sel_val)))
            driver.execute_script("arguments[0].click();", el)
            logger.info(f"Clicked login trigger: {sel_val}")
            clicked = True
            break
        except Exception:
            continue

    if not clicked:
        screenshot(driver, "00_login_button_not_found")
        raise RuntimeError("Login button not found on landing page")

    time.sleep(2)
    screenshot(driver, "01_modal_open")

    for sel_type, sel_val in selectors_to_try:
        try:
            el = wait.until(EC.element_to_be_clickable((sel_type, sel_val)))
            driver.execute_script("arguments[0].scrollIntoView(true);", el)
            time.sleep(0.3)
            driver.execute_script("arguments[0].click();", el)  # JS click avoids interception
            logger.info(f"Clicked login trigger using: {sel_val}")
            clicked = True
            break
        except Exception:
            continue

    if not clicked:
        # Last resort: dump all links so we can debug
        all_links = driver.execute_script(
            "return Array.from(document.querySelectorAll('a,button')).map(e => e.outerHTML).join('\\n')"
        )
        logger.error(f"Could not find Log in button. All links/buttons on page:\n{all_links[:3000]}")
        screenshot(driver, "00_login_button_not_found")
        raise RuntimeError("Login button not found on landing page")

    time.sleep(2)
    screenshot(driver, "01_modal_open")

    # ── Step 2: Fill username ─────────────────────────────────
    for sel in [
        (By.XPATH, "//input[@type='text' and not(@readonly)]"),
        (By.XPATH, "//input[@type='email']"),
        (By.XPATH, "//input[contains(@placeholder,'Username') or contains(@placeholder,'Email') or contains(@placeholder,'email')]"),
        (By.NAME,  "username"),
        (By.NAME,  "email"),
    ]:
        try:
            u = wait.until(EC.element_to_be_clickable(sel))
            u.clear()
            u.send_keys(username)
            logger.info(f"Username entered using {sel}")
            break
        except Exception:
            continue

    # ── Step 3: Fill password ─────────────────────────────────
    for sel in [
        (By.XPATH, "//input[@type='password']"),
        (By.NAME,  "password"),
    ]:
        try:
            p = driver.find_element(*sel)
            p.clear()
            p.send_keys(password)
            logger.info("Password entered")
            break
        except Exception:
            continue

    screenshot(driver, "02_creds_filled")

    # ── Step 4: Click submit inside modal ─────────────────────
    for sel in [
        (By.XPATH, "//button[@type='submit']"),
        (By.XPATH, "//button[contains(text(),'Log in')]"),
        (By.XPATH, "//button[contains(text(),'Login')]"),
        (By.XPATH, "//button[contains(text(),'Sign in')]"),
    ]:
        try:
            btn = wait.until(EC.element_to_be_clickable(sel))
            driver.execute_script("arguments[0].click();", btn)
            logger.info(f"Submit clicked using {sel}")
            break
        except Exception:
            continue

    # ── Step 5: Wait for dashboard ────────────────────────────
    try:
        wait.until(lambda d: (
            "login" not in d.current_url.lower() and
            d.current_url.rstrip("/") != LOGIN_URL.rstrip("/")
        ))
    except Exception:
        pass

    time.sleep(3)
    screenshot(driver, "03_after_login")
    logger.info(f"Post-login URL: {driver.current_url}")

def set_dropdown(driver, wait, label, option):
    # Try native <select>
    try:
        el = wait.until(EC.presence_of_element_located(
            (By.XPATH, f"//*[contains(text(),'{label}')]/following::select[1]")))
        Select(el).select_by_visible_text(option); return
    except: pass
    # Custom dropdown
    try:
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, f"//*[contains(text(),'{label}')]/following::*[@role='listbox' or @role='combobox' or contains(@class,'dropdown')][1]")
        )).click()
        time.sleep(0.5)
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, f"//*[contains(@role,'option') or contains(@class,'option')][normalize-space(text())='{option}']")
        )).click()
    except Exception as e:
        logger.error(f"Dropdown '{label}'→'{option}' failed: {e}"); raise

def set_date(driver, wait, label: str, dt: datetime):
    iso_val = dt.strftime("%Y-%m-%d")   # YYYY-MM-DD — HTML input internal format

    date_xpaths = [
        f"//*[contains(text(),'{label}')]/following::input[@type='date'][1]",
        f"//*[contains(text(),'{label}')]/following::input[contains(@class,'date')][1]",
        f"//*[contains(text(),'{label}')]/following::input[1]",
    ]

    for xpath in date_xpaths:
        try:
            inp = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))

            # JS setter — works with React/Angular controlled inputs
            driver.execute_script("""
                var inp = arguments[0];
                var val = arguments[1];
                var nativeSetter = Object.getOwnPropertyDescriptor(
                    window.HTMLInputElement.prototype, 'value').set;
                nativeSetter.call(inp, val);
                inp.dispatchEvent(new Event('input',  { bubbles: true }));
                inp.dispatchEvent(new Event('change', { bubbles: true }));
            """, inp, iso_val)

            time.sleep(0.3)
            actual = inp.get_attribute("value")
            if actual == iso_val:
                logger.info(f"Date '{label}' set to {iso_val} ✅")
                return

            # Fallback: send_keys with YYYY-MM-DD
            inp.click()
            inp.send_keys(Keys.CONTROL + "a")
            inp.send_keys(Keys.DELETE)
            inp.send_keys(iso_val)
            logger.info(f"Date '{label}' sent via keys: {iso_val}")
            return

        except Exception:
            continue

    raise RuntimeError(f"Date field not found for: '{label}'")

def wait_for_download(download_dir, before, timeout=180):
    end = time.time() + timeout
    while time.time() < end:
        time.sleep(2)
        new = set(os.listdir(download_dir)) - before
        done = [f for f in new if f.endswith(".xlsx") and not f.endswith(".crdownload")]
        if done: return os.path.join(download_dir, done[0])
    raise TimeoutError("Download timed out")
    
def dismiss_dialog_if_present(driver, timeout=5):
    """Click OK on AJIO Announcement dialog if it appears."""
    try:
        ok_btn = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.XPATH,
                "//div[contains(@class,'MuiDialog-container')]//button[normalize-space(text())='OK']"
            ))
        )
        ok_btn.click()
        logger.info("Announcement dialog dismissed (OK clicked) ✅")
        time.sleep(1)
    except TimeoutException:
        logger.info("No announcement dialog found, continuing...")
    except Exception:
        try:
            driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
            logger.info("Announcement dialog dismissed (ESC fallback) ✅")
            time.sleep(1)
        except Exception as ex:
            logger.warning(f"Dialog dismiss failed entirely: {ex}")

def download_report(driver, from_dt, to_dt, download_dir, label):
    abs_dl = os.path.abspath(download_dir)
    wait   = WebDriverWait(driver, 60)
    logger.info(f"[{label}] {from_dt.date()} → {to_dt.date()}")
    driver.get(REPORTS_URL)
    time.sleep(3)

    # ── Guard: if redirected to login page, raise clearly ─────
    if "ajiocommerce" in driver.current_url and "reports" not in driver.current_url:
        screenshot(driver, f"{label}_LOGIN_REDIRECT_ERROR")
        raise RuntimeError(
            f"[{label}] Redirected to login instead of reports! "
            f"Current URL: {driver.current_url} — Login likely failed."
        )

    screenshot(driver, f"{label}_01_reports")
    screenshot(driver, f"{label}_01")
    dismiss_dialog_if_present(driver)
    set_dropdown(driver, wait, "Report Type", "Dropship Rtv Report"); time.sleep(1)
    set_date(driver, wait, "From Date", from_dt)
    set_date(driver, wait, "To Date",   to_dt)
    screenshot(driver, f"{label}_02_dates")
    view_btn = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//button[contains(., 'View') and not(contains(., 'Preview'))]")
    ))
    driver.execute_script("arguments[0].click();", view_btn)
    logger.info(f"[{label}] Clicked View")
    try: wait.until(EC.presence_of_element_located((By.XPATH, "//table//tbody//tr")))
    except: pass
    time.sleep(4); screenshot(driver, f"{label}_03_table")
    # ── Check for empty data before trying to Export ───────────
    try:
        no_data = driver.find_element(
            By.XPATH,
            "//*[contains(text(),'No Data to show') or contains(text(),'No data to show') or contains(text(),'No Data To Show')]"
        )
        if no_data.is_displayed():
            logger.warning(f"[{label}] No data found for this date range — skipping export")
            return None   # ← caller must handle None
    except:
        pass  # element not found = data exists, proceed normally

    before = set(os.listdir(abs_dl))
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//button[contains(., 'Export') or contains(., 'EXPORT')]")
    )).click()
    path = wait_for_download(abs_dl, before)
    logger.info(f"[{label}] Downloaded: {path}")
    return path

def run_scraper(username, password, download_dir="downloads"):
    os.makedirs(download_dir, exist_ok=True)
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    run_a_from, run_a_to = today - timedelta(days=89), today
    run_b_to,  run_b_from = today - timedelta(days=90), today - timedelta(days=178)
    driver = build_driver(download_dir)
    try:
        login(driver, username, password)
        p_a = download_report(driver, run_a_from, run_a_to, download_dir, "RunA")
        if p_a is not None:
            path_a = os.path.join(download_dir, "run_a.xlsx")
            os.replace(p_a, path_a)
        else:
            path_a = None
            logger.warning("[RunA] No data — skipped")

        p_b = download_report(driver, run_b_from, run_b_to, download_dir, "RunB")
        if p_b is not None:
            path_b = os.path.join(download_dir, "run_b.xlsx")
            os.replace(p_b, path_b)
        else:
            path_b = None
            logger.warning("[RunB] No data — skipped")

        if path_a is None and path_b is None:
            raise RuntimeError("Both RunA and RunB have no data — nothing to process!")

        return path_a, path_b
    finally:
        driver.quit()
