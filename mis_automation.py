"""
Kunal MIS Report Automation
════════════════════════════
Flow:
  1. Selenium opens panel.parcelx.in with HTTP Basic Auth embedded in URL
  2. Fills the form-based login (Sudhanshu / @shudh@@$%)
  3. Opens /mis_report
  4. Expands "Non Mandatory Fields" -> ticks Sales POC checkbox
  5. Selects "Kunal Chauhan" from the dropdown
  6. Sets from_date = today-30d, to_date = today
  7. Clicks Search -> waits for results -> clicks Export Excel
  8. Uploads rows to Google Sheet (gid=2029444177)
  9. Emails the .xlsx to kunal.chauhan@parcelx.in
     CC: sudhanshu@parcelx.in, sparsh@parcelx.in

Schedule (IST): 09:00 · 13:00 · 16:00 · 20:00
"""

import os, time, logging, smtplib, schedule, datetime, traceback
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.base     import MIMEBase
from email.mime.text     import MIMEText
from email               import encoders

from selenium                          import webdriver
from selenium.webdriver.common.by      import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui     import WebDriverWait, Select
from selenium.webdriver.support        import expected_conditions as EC
from selenium.common.exceptions        import TimeoutException, NoSuchElementException

import gspread
from google.oauth2.service_account import Credentials
import openpyxl

# ══════════════════════════════════════════════
# CONFIG
# ══════════════════════════════════════════════
BASIC_USER  = os.getenv("PARCELX_BASIC_USER", "pax")
BASIC_PASS  = os.getenv("PARCELX_BASIC_PASS",  "cloud@4w5")
FORM_USER   = os.getenv("PARCELX_FORM_USER",  "Sudhanshu")
FORM_PASS   = os.getenv("PARCELX_FORM_PASS",  "@shudh@@$%")

AUTH_BASE   = f"https://{BASIC_USER}:{BASIC_PASS}@panel.parcelx.in"

SALES_POC_VALUE = "26"
SALES_POC_NAME  = "Kunal Chauhan"

GOOGLE_SHEET_ID  = "1gUwZoO0v3Rjkx4XshHHuBUwPyMfLY7KaE7UfzBEPLyE"
GOOGLE_SHEET_GID = "2029444177"
SA_FILE          = os.getenv("GOOGLE_SA_FILE", "service_account.json")

SMTP_HOST = os.getenv("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT") or "587")
SMTP_USER = os.getenv("SMTP_USER", "sudhanshu@parcelx.in")
SMTP_PASS = os.getenv("SMTP_PASS", "")

MAIL_TO = ["kunal.chauhan@parcelx.in"]
MAIL_CC = ["sudhanshu@parcelx.in", "sparsh@parcelx.in"]

DOWNLOAD_DIR = Path(os.getenv("DOWNLOAD_DIR", "/tmp/mis_downloads"))
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(), logging.FileHandler("mis_automation.log")],
)
log = logging.getLogger(__name__)


# ══════════════════════════════════════════════
# SELENIUM
# ══════════════════════════════════════════════

def build_driver():
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1920,1080")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    prefs = {
        "download.default_directory":   str(DOWNLOAD_DIR),
        "download.prompt_for_download": False,
        "download.directory_upgrade":   True,
        "safebrowsing.enabled":         True,
    }
    opts.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=opts)
    driver.implicitly_wait(10)
    return driver


def wait_for_new_xlsx(before, timeout=120):
    deadline = time.time() + timeout
    while time.time() < deadline:
        time.sleep(2)
        after    = set(DOWNLOAD_DIR.glob("*.xls*"))
        new      = after - before
        complete = [f for f in new if not f.name.endswith(".crdownload")]
        if complete:
            return max(complete, key=lambda p: p.stat().st_mtime)
    return None


# ══════════════════════════════════════════════
# BROWSER STEPS
# ══════════════════════════════════════════════

def login(driver, wait):
    log.info("Opening login page with Basic Auth...")
    driver.get(f"{AUTH_BASE}/login")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "form")))

    for sel in [(By.NAME,"username"),(By.NAME,"email"),(By.ID,"username"),
                (By.ID,"email"),(By.CSS_SELECTOR,"input[type='text']")]:
        try:
            e = driver.find_element(*sel)
            e.clear(); e.send_keys(FORM_USER)
            log.info("Username filled"); break
        except NoSuchElementException:
            continue

    for sel in [(By.NAME,"password"),(By.ID,"password"),
                (By.CSS_SELECTOR,"input[type='password']")]:
        try:
            e = driver.find_element(*sel)
            e.clear(); e.send_keys(FORM_PASS)
            log.info("Password filled"); break
        except NoSuchElementException:
            continue

    try:
        driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
    except NoSuchElementException:
        driver.find_element(By.CSS_SELECTOR, "input[type='submit']").click()

    wait.until(lambda d: "/login" not in d.current_url)
    log.info("Login successful -> %s", driver.current_url)


def open_mis_report(driver, wait):
    log.info("Navigating to /mis_report...")
    driver.get(f"{AUTH_BASE}/mis_report")
    wait.until(EC.presence_of_element_located((By.ID, "searchbtn")))
    log.info("MIS Report page loaded")


def expand_non_mandatory(driver, wait):
    log.info("Expanding Non-Mandatory Fields...")
    wait.until(EC.element_to_be_clickable(
        (By.CSS_SELECTOR, "h4.dropdown_mis_el"))).click()
    time.sleep(0.6)


def enable_sales_poc_filter(driver, wait):
    log.info("Enabling Sales POC checkbox (box8)...")
    chk = wait.until(EC.presence_of_element_located((By.ID, "box8")))
    if not chk.is_selected():
        driver.execute_script("arguments[0].click();", chk)
        driver.execute_script(
            "if(typeof checkboxCheck==='function') checkboxCheck(8);")
    time.sleep(0.4)


def select_kunal_chauhan(driver, wait):
    log.info("Selecting Sales POC: %s...", SALES_POC_NAME)
    Select(wait.until(EC.visibility_of_element_located(
        (By.ID, "sales_poc")))).select_by_value(SALES_POC_VALUE)


def set_date_range(driver, wait):
    today    = datetime.date.today()
    from_str = (today - datetime.timedelta(days=30)).strftime("%d-%m-%Y")
    to_str   = today.strftime("%d-%m-%Y")
    log.info("Date range: %s -> %s", from_str, to_str)

    fi = wait.until(EC.presence_of_element_located((By.ID, "from_date")))
    driver.execute_script("arguments[0].removeAttribute('readonly');", fi)
    driver.execute_script("arguments[0].value = '';", fi)
    fi.send_keys(from_str)

    try:
        ti = driver.find_element(By.ID, "to_date")
        driver.execute_script("arguments[0].removeAttribute('readonly');", ti)
        driver.execute_script("arguments[0].value = '';", ti)
        ti.send_keys(to_str)
    except NoSuchElementException:
        pass

    return from_str, to_str


def search_and_export(driver, wait):
    before = set(DOWNLOAD_DIR.glob("*.xls*"))
    log.info("Clicking Search...")
    wait.until(EC.element_to_be_clickable((By.ID, "searchbtn"))).click()

    log.info("Waiting for results...")
    try:
        wait.until(EC.visibility_of_element_located((By.ID, "btn_exportexcel")))
    except TimeoutException:
        log.error("Export button did not appear — no results?")
        return None

    time.sleep(2)
    log.info("Clicking Export Excel...")
    driver.find_element(By.ID, "btn_exportexcel").click()

    log.info("Waiting for download...")
    filepath = wait_for_new_xlsx(before, timeout=120)
    if filepath:
        log.info("Downloaded -> %s (%.1f KB)",
                 filepath.name, filepath.stat().st_size / 1024)
    else:
        log.error("Download timed out")
    return filepath


# ══════════════════════════════════════════════
# GOOGLE SHEETS
# ══════════════════════════════════════════════

def upload_to_sheets(filepath):
    if not Path(SA_FILE).exists():
        log.warning("'%s' not found — skipping Sheets upload", SA_FILE)
        return False
    try:
        creds = Credentials.from_service_account_file(SA_FILE, scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ])
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(GOOGLE_SHEET_ID)

        ws = None
        for w in sh.worksheets():
            if str(w.id) == GOOGLE_SHEET_GID:
                ws = w
                break
        if ws is None:
            ws = sh.get_worksheet(0)

        wb   = openpyxl.load_workbook(filepath, data_only=True)
        rows = [[str(c.value) if c.value is not None else ""
                 for c in row] for row in wb.active.iter_rows()]
        ws.clear()
        if rows:
            ws.update(rows, value_input_option="USER_ENTERED")
        log.info("Google Sheet updated — %d rows", len(rows))
        return True
    except Exception:
        log.error("Sheets upload failed:\n%s", traceback.format_exc())
        return False


# ══════════════════════════════════════════════
# EMAIL
# ══════════════════════════════════════════════

def send_email(filepath, from_str, to_str):
    if not SMTP_PASS:
        log.warning("SMTP_PASS not set — skipping email")
        return False

    now_str   = datetime.datetime.now().strftime("%I:%M %p")
    today_str = datetime.date.today().strftime("%d %b %Y")
    sheet_url = (f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}"
                 f"/edit#gid={GOOGLE_SHEET_GID}")

    msg            = MIMEMultipart()
    msg["From"]    = SMTP_USER
    msg["To"]      = ", ".join(MAIL_TO)
    msg["Cc"]      = ", ".join(MAIL_CC)
    msg["Subject"] = f"MIS Report — {SALES_POC_NAME} | {today_str}  {now_str} IST"
    msg.attach(MIMEText(
        f"Hi Kunal,\n\n"
        f"Please find the MIS report for {from_str} to {to_str} attached.\n\n"
        f"Live Google Sheet:\n{sheet_url}\n\n"
        f"Generated at {now_str} IST, {today_str}.\n\n"
        f"Regards,\nAutomated MIS System",
        "plain"
    ))

    if filepath and filepath.exists():
        with open(filepath, "rb") as fh:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(fh.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition",
            f'attachment; filename="MIS_{SALES_POC_NAME.replace(" ","_")}_{today_str}.xlsx"')
        msg.attach(part)

    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as srv:
            srv.ehlo(); srv.starttls(); srv.login(SMTP_USER, SMTP_PASS)
            srv.sendmail(SMTP_USER, MAIL_TO + MAIL_CC, msg.as_string())
        log.info("Email sent TO=%s CC=%s", MAIL_TO, MAIL_CC)
        return True
    except Exception:
        log.error("Email failed:\n%s", traceback.format_exc())
        return False


# ══════════════════════════════════════════════
# ORCHESTRATOR
# ══════════════════════════════════════════════

def run_job():
    log.info("━" * 60)
    log.info("JOB START  %s", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    driver = None; filepath = None; from_str = to_str = ""
    try:
        driver = build_driver()
        wait   = WebDriverWait(driver, 30)
        login(driver, wait)
        open_mis_report(driver, wait)
        expand_non_mandatory(driver, wait)
        enable_sales_poc_filter(driver, wait)
        select_kunal_chauhan(driver, wait)
        from_str, to_str = set_date_range(driver, wait)
        filepath = search_and_export(driver, wait)
    except Exception:
        log.error("Browser error:\n%s", traceback.format_exc())
        if driver:
            driver.save_screenshot(
                str(DOWNLOAD_DIR / f"error_{int(time.time())}.png"))
    finally:
        if driver:
            driver.quit()

    if filepath:
        upload_to_sheets(filepath)
    send_email(filepath, from_str, to_str)

    log.info("JOB END    %s", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    log.info("━" * 60)


def main():
    log.info("Scheduler armed — 09:00, 13:00, 16:00, 20:00 IST")
    for t in ["09:00", "13:00", "16:00", "20:00"]:
        schedule.every().day.at(t).do(run_job)
    log.info("Running initial job...")
    run_job()
    while True:
        schedule.run_pending()
        time.sleep(30)


if __name__ == "__main__":
    main()
