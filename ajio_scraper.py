import os, time, logging
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

logger = logging.getLogger(__name__)
LOGIN_URL   = "https://seller.ajio.com/ajiocommerce/"
REPORTS_URL = "https://seller.ajio.com/vmsui/reports/ViewReports"

def build_driver(download_dir):
    abs_dl = os.path.abspath(download_dir)
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1920,1080")
    opts.add_experimental_option("prefs", {
        "download.default_directory": abs_dl,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
    })
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)

def screenshot(driver, name):
    try: driver.save_screenshot(f"downloads/debug_{name}.png")
    except: pass

def login(driver, username, password):
    driver.get(LOGIN_URL)
    wait = WebDriverWait(driver, 30)
    for sel in [(By.NAME,"username"),(By.NAME,"email"),(By.ID,"username"),(By.XPATH,"//input[@type='email' or @type='text'][1]")]:
        try: u=wait.until(EC.presence_of_element_located(sel)); u.clear(); u.send_keys(username); break
        except: continue
    for sel in [(By.NAME,"password"),(By.ID,"password"),(By.XPATH,"//input[@type='password']")]:
        try: p=driver.find_element(*sel); p.clear(); p.send_keys(password); break
        except: continue
    for sel in [(By.XPATH,"//button[@type='submit']"),(By.XPATH,"//button[contains(translate(text(),'LOGIN','login'),'login')]")]:
        try: driver.find_element(*sel).click(); break
        except: continue
    try: wait.until(lambda d: "seller.ajio.com" in d.current_url and "login" not in d.current_url.lower())
    except: pass
    time.sleep(3)
    screenshot(driver, "after_login")
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

def set_date(driver, wait, label, dt):
    date_str = dt.strftime("%d-%m-%Y")
    for xpath in [
        f"//*[contains(text(),'{label}')]/following::input[@type='date'][1]",
        f"//*[contains(text(),'{label}')]/following::input[contains(@class,'date')][1]",
        f"//*[contains(text(),'{label}')]/following::input[1]",
    ]:
        try:
            inp = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
            inp.click(); inp.send_keys(Keys.CONTROL+"a"); inp.send_keys(Keys.DELETE)
            time.sleep(0.2); inp.send_keys(date_str)
            logger.info(f"Date '{label}' = {date_str}"); return
        except: continue
    raise RuntimeError(f"Date field not found: {label}")

def wait_for_download(download_dir, before, timeout=180):
    end = time.time() + timeout
    while time.time() < end:
        time.sleep(2)
        new = set(os.listdir(download_dir)) - before
        done = [f for f in new if f.endswith(".xlsx") and not f.endswith(".crdownload")]
        if done: return os.path.join(download_dir, done[0])
    raise TimeoutError("Download timed out")

def download_report(driver, from_dt, to_dt, download_dir, label):
    abs_dl = os.path.abspath(download_dir)
    wait   = WebDriverWait(driver, 60)
    logger.info(f"[{label}] {from_dt.date()} → {to_dt.date()}")
    driver.get(REPORTS_URL); time.sleep(3)
    screenshot(driver, f"{label}_01")
    set_dropdown(driver, wait, "Report Type", "Dropship Rtv Report"); time.sleep(1)
    set_date(driver, wait, "From Date", from_dt)
    set_date(driver, wait, "To Date",   to_dt)
    screenshot(driver, f"{label}_02_dates")
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//button[normalize-space(text())='View' or normalize-space(text())='VIEW']")
    )).click()
    try: wait.until(EC.presence_of_element_located((By.XPATH, "//table//tbody//tr")))
    except: pass
    time.sleep(4); screenshot(driver, f"{label}_03_table")
    before = set(os.listdir(abs_dl))
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//button[normalize-space(text())='Export' or normalize-space(text())='EXPORT']")
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
        path_a = os.path.join(download_dir, "run_a.xlsx"); os.replace(p_a, path_a)
        p_b = download_report(driver, run_b_from, run_b_to, download_dir, "RunB")
        path_b = os.path.join(download_dir, "run_b.xlsx"); os.replace(p_b, path_b)
        return path_a, path_b
    finally:
        driver.quit()