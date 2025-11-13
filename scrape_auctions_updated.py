import csv, os, json, time, re, base64
from datetime import datetime, date, timedelta
from typing import List, Dict, Tuple

import requests
from requests.adapters import HTTPAdapter, Retry

# Selenium
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, ElementClickInterceptedException, WebDriverException
)

# ==============================
# Local storage
# ==============================
LOCAL_CSV = "cricauction_upcoming.csv"
CSV_HEADERS = ["Tournament Name", "Location", "Total players", "Auction Date", "Date Added"]

# ==============================
# Share link + Worksheet
# ==============================
SHARE_LINK = "https://cricbattle.sharepoint.com/:x:/s/CorpDevaaf94619130d4a79af9f3aeae502bdb5/EdVug56P9cdNqWk8YrJV8-oBBqJ_vMqA4GlPACCCnAD4Og?e=ZmkkdP"
WORKSHEET_NAME = "CricAuction Auction"   # will be created if missing

# ==============================
# Behavior tuning
# ==============================
HEADLESS = True
MAX_PAGES = 500
PAGE_CHANGE_TIMEOUT = 6
STAGNANT_PAGE_LIMIT = 1

# ==============================
# Helpers
# ==============================
def today_str() -> str:
    return date.today().strftime("%Y-%m-%d")

def norm_text(s) -> str:
    if s is None:
        return ""
    if isinstance(s, (datetime, date)):
        return s.strftime("%Y-%m-%d")
    try:
        return re.sub(r"\s+", " ", str(s).strip())
    except Exception:
        return str(s)

def parse_players(text: str) -> str:
    m = re.search(r"(\d+)\s*Players", text or "", re.IGNORECASE)
    return m.group(1) if m else norm_text(text)

def parse_auction_date(val) -> str:
    if isinstance(val, (int, float)):
        try:
            base = datetime(1899, 12, 30)  # Excel 1900 system
            d = base + timedelta(days=float(val))
            return d.strftime("%Y-%m-%d")
        except Exception:
            pass
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")

    s = norm_text(val)
    if not s:
        return ""

    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    m = re.search(r"\b\d{4}-\d{2}-\d{2}\b", s)
    return m.group(0) if m else s

def row_key_from_list(row: List) -> Tuple[str, str, str]:
    title = norm_text(row[0]) if len(row) > 0 else ""
    loc   = norm_text(row[1]) if len(row) > 1 else ""
    auc   = parse_auction_date(row[3]) if len(row) > 3 else ""
    return (title, loc, auc)

def row_key_from_dict(d: Dict[str, str]) -> Tuple[str, str, str]:
    return (
        norm_text(d.get("Tournament Name", "")),
        norm_text(d.get("Location", "")),
        parse_auction_date(d.get("Auction Date", "")),
    )

def print_table_header():
    print("\n" + "-" * 120)
    print(f"{'Tournament Name':50} | {'Location':20} | {'Total':5} | {'Auction Date':12} | {'Date Added':10}")
    print("-" * 120)

def print_table_row(rec: Dict[str, str]):
    tn = rec.get("Tournament Name", "")[:50]
    lc = rec.get("Location", "")[:20]
    tp = rec.get("Total players", "")[:5]
    ad = rec.get("Auction Date", "")[:12]
    da = rec.get("Date Added", "")[:10]
    print(f"{tn:50} | {lc:20} | {tp:5} | {ad:12} | {da:10}")

# ==============================
# Selenium driver
# ==============================
def build_driver(headless: bool = HEADLESS) -> webdriver.Chrome:
    opts = Options()
    opts.page_load_strategy = "eager"
    if headless:
        opts.add_argument("--headless=new")
    prefs = {
        "profile.managed_default_content_settings.images": 2,
        "profile.managed_default_content_settings.stylesheets": 2
    }
    opts.add_experimental_option("prefs", prefs)
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1400,1200")
    opts.add_argument("--disable-notifications")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--disable-extensions")
    opts.add_argument("--disable-background-networking")
    opts.add_argument("--remote-allow-origins=*")
    driver = webdriver.Chrome(options=opts)
    driver.set_page_load_timeout(20)
    return driver

JS_SCRAPE_PAGE = r"""
return Array.from(document.querySelectorAll('.team-content')).map(card=>{
  const title = (card.querySelector('h2')?.textContent || '').trim();
  const location = (card.querySelector('.team-location p')?.textContent || '').trim();
  const subtexts = Array.from(card.querySelectorAll('.team-subcontent')).map(e => e.innerText || '');
  return {title, location, subtexts};
});
"""

def page_signature_js(driver) -> str:
    titles = driver.execute_script(
        "return Array.from(document.querySelectorAll('.team-content h2')).slice(0,6).map(e=>e.textContent.trim());"
    )
    return "|".join([t.strip() for t in titles if t])

# ==============================
# Scraper
# ==============================
def scrape_generator():
    url = "https://cricauction.live/upcoming-auction"
    driver = build_driver(headless=HEADLESS)
    wait = WebDriverWait(driver, 10)
    seen_on_session: set[Tuple[str, str, str]] = set()
    try:
        driver.get(url)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".team-content")))
        last_sig = page_signature_js(driver)
        stagnant = 0
        page_num = 1

        while True:
            cards = driver.execute_script(JS_SCRAPE_PAGE) or []
            for c in cards:
                title = norm_text(c.get("title", ""))
                location = norm_text(c.get("location", ""))

                players_text = ""
                auction_date = ""
                for raw in c.get("subtexts", []):
                    if not players_text and ("Players" in raw):
                        players_text = parse_players(raw)
                    if not auction_date and re.search(r"\b\d{1,2}[-/\.]\d{1,2}[-/\.]\d{4}\b", raw):
                        auction_date = parse_auction_date(raw)

                rec = {
                    "Tournament Name": title,
                    "Location": location or "Unknown Location",
                    "Total players": players_text,
                    "Auction Date": auction_date,
                    "Date Added": today_str(),
                }
                key = row_key_from_dict(rec)
                if title and key not in seen_on_session:
                    seen_on_session.add(key)
                    yield rec

            if page_num >= MAX_PAGES:
                break

            try:
                next_btn = driver.find_element(By.ID, "nextBtn")
                if not next_btn.is_displayed() or not next_btn.is_enabled():
                    break
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(0.15)
                try:
                    next_btn.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", next_btn)

                changed = False
                start = time.time()
                while time.time() - start < PAGE_CHANGE_TIMEOUT:
                    time.sleep(0.25)
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".team-content")))
                    sig = page_signature_js(driver)
                    if sig and sig != last_sig:
                        changed = True
                        last_sig = sig
                        break

                if not changed:
                    stagnant += 1
                    if stagnant > STAGNANT_PAGE_LIMIT:
                        break
                else:
                    stagnant = 0
                    page_num += 1

            except (NoSuchElementException, ElementClickInterceptedException, TimeoutException, WebDriverException):
                break
    finally:
        driver.quit()

# ==============================
# Local CSV I/O
# ==============================
def ensure_local_csv():
    if not os.path.exists(LOCAL_CSV):
        with open(LOCAL_CSV, "w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow(CSV_HEADERS)

def read_local_keys() -> set[Tuple[str, str, str]]:
    if not os.path.exists(LOCAL_CSV):
        return set()
    keys = set()
    with open(LOCAL_CSV, "r", newline="", encoding="utf-8") as f:
        reader = csv.reader(f)
        _ = next(reader, [])
        for row in reader:
            if not row:
                continue
            try:
                keys.add(row_key_from_list([
                    row[0] if len(row) > 0 else "",
                    row[1] if len(row) > 1 else "",
                    row[2] if len(row) > 2 else "",
                    row[3] if len(row) > 3 else "",
                    row[4] if len(row) > 4 else "",
                ]))
            except Exception:
                pass
    return keys

def append_local_rows(rows: List[List[str]]):
    if not rows:
        return
    ensure_local_csv()
    with open(LOCAL_CSV, "a", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(rows)

# ==============================
# Microsoft Graph Excel (via Share link)
# ==============================
class GraphExcelClient:
    def __init__(self, cred: dict, share_link: str, worksheet_name: str):
        self.tenant_id = cred["tenant_id"]
        self.client_id = cred["client_id"]
        self.client_secret = cred["client_secret"]
        self.share_link = share_link
        self.worksheet_name = worksheet_name

        self.base = "https://graph.microsoft.com/v1.0"
        self.session = requests.Session()
        retry = Retry(total=5, backoff_factor=0.4, status_forcelist=[429, 500, 502, 503, 504])
        self.session.mount("https://", HTTPAdapter(max_retries=retry))
        self._tok = None
        self._drive_id = None
        self._item_id = None

    def token(self):
        if self._tok:
            return self._tok
        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        data = {
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "grant_type": "client_credentials",
            "scope": "https://graph.microsoft.com/.default",
        }
        r = self.session.post(url, data=data, timeout=20)
        if not r.ok:
            try:
                err = r.json()
            except Exception:
                err = {"text": r.text}
            raise requests.HTTPError(f"Token request failed [{r.status_code}]: {err}", response=r)
        self._tok = r.json()["access_token"]
        return self._tok

    def headers(self):
        return {"Authorization": f"Bearer {self.token()}"}

    def _ensure_item(self):
        if self._drive_id and self._item_id:
            return
        share_id = "u!" + base64.urlsafe_b64encode(self.share_link.encode("utf-8")).decode("utf-8").rstrip("=")
        url = f"{self.base}/shares/{share_id}/driveItem"
        r = self.session.get(url, headers=self.headers(), timeout=20)
        if r.status_code == 403:
            raise requests.HTTPError(
                "403 from /shares: app-only access blocked for this site (Sites.Selected?). "
                "Grant this app write access to the site, or switch to delegated auth.",
                response=r,
            )
        r.raise_for_status()
        item = r.json()
        self._drive_id = item["parentReference"]["driveId"]
        self._item_id = item["id"]

    @staticmethod
    def _odata_quote(name: str) -> str:
        return name.replace("'", "''")

    def _worksheets(self) -> list:
        self._ensure_item()
        url = f"{self.base}/drives/{self._drive_id}/items/{self._item_id}/workbook/worksheets"
        r = self.session.get(url, headers=self.headers(), timeout=20)
        r.raise_for_status()
        return r.json().get("value", []) or []

    def ensure_worksheet(self):
        sheets = self._worksheets()
        if any(ws.get("name") == self.worksheet_name for ws in sheets):
            return
        url = f"{self.base}/drives/{self._drive_id}/items/{self._item_id}/workbook/worksheets/add"
        r = self.session.post(
            url,
            headers={**self.headers(), "Content-Type": "application/json"},
            json={"name": self.worksheet_name},
            timeout=20,
        )
        if r.status_code not in (200, 201, 409):
            r.raise_for_status()

    def get_used_values(self) -> List[List]:
        self._ensure_item()
        from urllib.parse import quote as urlquote
        ws = urlquote(self._odata_quote(self.worksheet_name))
        url = (
            f"{self.base}/drives/{self._drive_id}/items/{self._item_id}"
            f"/workbook/worksheets('{ws}')/usedRange(valuesOnly=true)"
        )
        r = self.session.get(url, headers=self.headers(), timeout=30)
        r.raise_for_status()
        return r.json().get("values", []) or []

    def append_rows(self, rows: List[List]):
        if not rows:
            return
        self.ensure_worksheet()
        used = self.get_used_values()
        start_row = (len(used) + 1) if used else 2  # keep row 1 for headers
        end_row = start_row + len(rows) - 1

        address = f"A{start_row}:E{end_row}"  # range only
        from urllib.parse import quote as urlquote
        ws = urlquote(self._odata_quote(self.worksheet_name))
        url = (
            f"{self.base}/drives/{self._drive_id}/items/{self._item_id}"
            f"/workbook/worksheets('{ws}')/range(address='{address}')"
        )
        r = self.session.patch(
            url,
            headers={**self.headers(), "Content-Type": "application/json"},
            json={"values": rows},
            timeout=30,
        )
        r.raise_for_status()

# ==============================
# Credentials loader - GitHub Actions only
# ==============================
def try_load_creds() -> dict | None:
    """Load credentials ONLY from environment variables (GitHub Secrets)"""
    # Check for Graph/SharePoint credentials
    cid = os.getenv("CLIENT_ID")
    cs  = os.getenv("CLIENT_SECRET")
    tid = os.getenv("TENANT_ID")
    if cid and cs and tid:
        print("✓ Using SharePoint credentials from environment")
        return {"client_id": cid, "client_secret": cs, "tenant_id": tid}
    
    # Check for CB credentials  
    email = os.getenv("CB_EMAIL")
    pwd   = os.getenv("CB_PASSWORD")
    if email and pwd:
        print("✓ Using CB credentials from environment")
        return {"email": email, "password": pwd}
    
    print("⚠ No credentials found in environment variables")
    return None

# ==============================
# Main
# ==============================
def main():
    print("Scraping...")
    print_table_header()

    ensure_local_csv()
    local_keys = read_local_keys()

    staged_local_rows: List[List[str]] = []
    for rec in scrape_generator():
        k = row_key_from_dict(rec)
        if k in local_keys:
            continue
        print_table_row(rec)
        row = [
            rec["Tournament Name"],
            rec["Location"],
            rec["Total players"],
            rec["Auction Date"],
            rec["Date Added"],
        ]
        staged_local_rows.append(row)
        local_keys.add(k)

    append_local_rows(staged_local_rows)
    print(f"\nSaved locally to {LOCAL_CSV}. New local rows: {len(staged_local_rows)}")

    cred = try_load_creds()
    if not cred or not all(k in cred for k in ("client_id", "client_secret", "tenant_id")):
        print("No Graph credentials → skipping SharePoint sync.")
        print("\n" + "-" * 120 + "\nDone.\n" + "-" * 120)
        return

    print("\nAttempting SharePoint sync (skipping on auth error)...")
    try:
        graph = GraphExcelClient(cred, SHARE_LINK, WORKSHEET_NAME)
        graph.ensure_worksheet()
        sp_existing = graph.get_used_values()
        if sp_existing and sp_existing[0] and "Tournament" in str(sp_existing[0][0]):
            sp_rows = sp_existing[1:]
        else:
            sp_rows = sp_existing
        sp_keys = {row_key_from_list(r) for r in sp_rows}

        rows_for_sp = [r for r in staged_local_rows if row_key_from_list(r) not in sp_keys]
        if rows_for_sp:
            graph.append_rows(rows_for_sp)
            print(f"SharePoint sync complete. Uploaded rows: {len(rows_for_sp)}")
        else:
            print("SharePoint had all these rows already. Nothing uploaded.")
    except requests.HTTPError as e:
        print(f"SharePoint sync skipped due to error: {e}")
    except Exception as e:
        print(f"SharePoint sync skipped due to unexpected error: {e}")

    print("\n" + "-" * 120 + "\nDone.\n" + "-" * 120)

if __name__ == "__main__":
    main()
