"""
telephone_infotorg.py - phone number lookup via Infotorg Nummeropplysning
(nummeropplysning.infotorg.no). Drop-in replacement for telephone.py: it is
started by app5.py in exactly the same way:

    python telephone_infotorg.py <path-to-excel-file>

What it does
------------
* Reads the sheet(s) named "Private" / "Private 2024" ... (private persons only;
  companies are never looked up).
* For every person it searches "NAME POSTALCODE" and accepts a result card
  only if BOTH the name and the postal code match.
* If the card has several numbers, only a MOBILE number is used (starts with 4 or 9).
* If several different persons match (same name + postal code, different
  address), the result is "-" (a wrong number is worse than none).
* No second search without the postal code.
* Found mobile numbers are cached for 90 days (phone_cache.sqlite3, shared with the
  old script), so repeated people cost no request. PHONE_CACHE=0 switches it off.
* Polite pace: one request every ~6-8 s across ALL running copies (shared
  lock file), and it stops by itself if the site shows a block / captcha, or
  does not answer 3 times in a row. Rows not processed then get
  "BLOCKED - retry later".
* Everything printed is also appended to telephone_live.log next to this script
  (watch it with: tail -f telephone_live.log).
* The finished file is saved as "<name> (T).xlsx", like the old script.
"""
# ============================================================
# FORCE UTF-8 OUTPUT (the script is started via subprocess by app5.py)
# ============================================================
import sys
import io

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
else:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

import atexit
import os
import random
import re
import sqlite3
import tempfile
import time
from contextlib import closing
from urllib.parse import quote

try:
    import fcntl  # Linux: lock shared between several running copies
except ImportError:  # Windows
    fcntl = None

import pandas as pd
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By

# ============================================================
# SETTINGS
# ============================================================
BASE_URL = "https://nummeropplysning.infotorg.no/search?fulltext="
MIN_GAP_SECONDS = 6       # minimum pause between two requests (all users together)
JITTER_SECONDS = 2        # plus a random 0..2 s
CACHE_DAYS = 90
USE_CACHE = os.environ.get("PHONE_CACHE", "1") != "0"
BLOCK_TEXT = "BLOCKED - retry later"
SHOW_BROWSER = os.environ.get("SHOW_BROWSER", "0") == "1"   # local debugging only
BLOCK_WORDS = [
    "captcha", "blokkert", "blocked", "access denied", "forbidden",
    "too many requests", "unusual traffic", "are you a robot", "er du en robot",
    "verify you are human", "just a moment", "midlertidig sperret",
]

_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
_CACHE_DB = os.path.join(_BASE_DIR, "phone_cache.sqlite3")
_RATE_FILE = os.path.join(tempfile.gettempdir(), "infotorg_rate_limit.txt")
_LIVE_LOG = os.path.join(_BASE_DIR, "telephone_live.log")

_timeouts_in_row = 0


class BlockedBySite(Exception):
    """The site shows a block / captcha, or stopped answering."""


# ============================================================
# PURE HELPERS (name / number logic)
# ============================================================
def extract_zip(zip_city):
    m = re.match(r"\s*(\d{4})\b", zip_city or "")
    return m.group(1) if m else None


def normalize_number(raw):
    """Anything that looks like a Norwegian number -> 8 digits, else None."""
    digits = re.sub(r"\D", "", raw or "")
    if digits.startswith("0047"):
        digits = digits[4:]
    elif digits.startswith("47") and len(digits) == 10:
        digits = digits[2:]
    return digits if len(digits) == 8 else None


def is_mobile(digits):
    return bool(digits) and digits[0] in "49"


def pretty(digits):
    return f"{digits[:2]} {digits[2:4]} {digits[4:6]} {digits[6:8]}"


def _words(text):
    return set(re.findall(r"\w+", (text or "").lower()))


def match_tier(wanted_list, card_name):
    """3 = same words, 2 = card has all searched words (+ extra), 1 = first and
    last word found (names with 3+ words only), 0 = no match."""
    wanted, words = set(wanted_list), _words(card_name)
    if not wanted:
        return 0
    if wanted == words:
        return 3
    if wanted <= words:
        return 2
    if len(wanted_list) > 2 and {wanted_list[0], wanted_list[-1]} <= words:
        return 1
    return 0


def choose_number(cards, name, zip_code):
    """cards: [{'name','addr','numbers':[8-digit strings]}].
    Returns (digits_or_None, status, note)."""
    wanted_list = re.findall(r"\w+", name.lower())
    candidates, same_name_other_zip = [], 0
    for c in cards:
        tier = match_tier(wanted_list, c["name"])
        if not tier:
            continue
        if zip_code not in c["addr"]:
            same_name_other_zip += 1
            continue
        candidates.append((tier, c))
    if not candidates:
        note = (f"{same_name_other_zip} card(s) with this name but another postal code"
                if same_name_other_zip else "no card with this name")
        return None, "NO_MATCH", note

    best_tier = max(t for t, _ in candidates)
    top = [c for t, c in candidates if t == best_tier]

    # Cards with the same name AND address are the same person (e.g. two
    # listings); a different address means a different person.
    persons = {}
    for c in top:
        key = (frozenset(_words(c["name"])), re.sub(r"\s+", " ", c["addr"].lower()).strip())
        persons.setdefault(key, []).extend(n for n in c["numbers"] if n not in persons.get(key, []))
    if len(persons) > 1:
        return None, "AMBIGUOUS", f"{len(persons)} different persons match name + postal code"

    numbers = next(iter(persons.values()))
    mobiles = [n for n in numbers if is_mobile(n)]
    if mobiles:
        return mobiles[0], "FOUND_MOBILE", f"{len(numbers)} number(s) on card" if len(numbers) > 1 else ""
    if numbers:
        return None, "LANDLINE_ONLY", f"{len(numbers)} number(s), none mobile"
    return None, "NO_NUMBER_ON_CARD", ""


# ============================================================
# CACHE (90 days, shared with the old script; stores "XX XX XX XX")
# ============================================================
def _cache_conn():
    conn = sqlite3.connect(_CACHE_DB, timeout=30)
    conn.execute("CREATE TABLE IF NOT EXISTS phones "
                 "(name TEXT, zip TEXT, phone TEXT, ts REAL, PRIMARY KEY (name, zip))")
    return conn


def cache_get(name, zip_code):
    """Returns a cached MOBILE number or None (old landline entries are ignored)."""
    if not USE_CACHE:
        return None
    try:
        with closing(_cache_conn()) as c:
            row = c.execute("SELECT phone, ts FROM phones WHERE name=? AND zip=?",
                            (name.lower(), zip_code)).fetchone()
        if row and time.time() - row[1] < CACHE_DAYS * 86400:
            digits = normalize_number(row[0])
            if digits and is_mobile(digits):
                return pretty(digits)
    except Exception as e:
        print(f"⚠️ Cache read failed: {e}")
    return None


def cache_put(name, zip_code, formatted):
    if not USE_CACHE or not formatted:
        return
    try:
        with closing(_cache_conn()) as c:
            c.execute("INSERT OR REPLACE INTO phones VALUES (?, ?, ?, ?)",
                      (name.lower(), zip_code, formatted, time.time()))
            c.commit()
    except Exception as e:
        print(f"⚠️ Cache write failed: {e}")


# ============================================================
# RATE LIMIT (shared by every running copy of this script)
# ============================================================
def wait_for_turn():
    gap = MIN_GAP_SECONDS + random.uniform(0, JITTER_SECONDS)
    if fcntl is None:
        time.sleep(gap)
        return
    with open(_RATE_FILE, "a+") as f:
        fcntl.flock(f, fcntl.LOCK_EX)
        try:
            f.seek(0)
            try:
                last = float(f.read().strip() or 0)
            except ValueError:
                last = 0.0
            wait = last + gap - time.time()
            if wait > 0:
                time.sleep(wait)
            f.seek(0)
            f.truncate()
            f.write(str(time.time()))
            f.flush()
        finally:
            fcntl.flock(f, fcntl.LOCK_UN)


# ============================================================
# BROWSER
# ============================================================
def make_driver():
    opts = webdriver.ChromeOptions()
    for arg in ("--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu",
                "--window-size=1920,1080"):
        opts.add_argument(arg)
    if not SHOW_BROWSER:
        opts.add_argument("--headless=new")
    # Service() with no arguments: Selenium picks and caches the matching driver.
    drv = webdriver.Chrome(options=opts)
    drv.set_page_load_timeout(40)
    return drv


def _text(el, selector):
    try:
        return (el.find_element(By.CSS_SELECTOR, selector).get_attribute("textContent") or "").strip()
    except Exception:
        return ""


def read_cards(driver):
    cards = []
    for el in driver.find_elements(By.CSS_SELECTOR, "div.info-card"):
        raw = [n.get_attribute("textContent") or "" for n in el.find_elements(
            By.CSS_SELECTOR, "[data-test='phone-number-container'] .copyable")]
        raw += [a.get_attribute("href") or "" for a in el.find_elements(
            By.CSS_SELECTOR, "a[href^='tel:']")]
        numbers = []
        for r in raw:
            d = normalize_number(r)
            if d and d not in numbers:
                numbers.append(d)
        cards.append({"name": _text(el, "h4"), "addr": _text(el, "small"),
                      "numbers": numbers})
    return cards


def wait_for_answer(driver, timeout=20):
    """Wait until the page shows cards, 'N resultater funnet' or 'ga ingen treff'."""
    for _ in range(timeout * 2):
        try:
            body = driver.find_element(By.TAG_NAME, "body").text.lower()
        except Exception:
            body = ""
        if (driver.find_elements(By.CSS_SELECTOR, "div.info-card")
                or "ga ingen treff" in body or "resultat" in body):
            time.sleep(1.5)  # let the cards finish rendering
            try:
                body = driver.find_element(By.TAG_NAME, "body").text.lower()
            except Exception:
                pass
            return True, body
        time.sleep(0.5)
    return False, ""


def click_load_more(driver):
    for b in driver.find_elements(By.XPATH, "//button[contains(normalize-space(.), 'Last mer')]"):
        try:
            if b.is_displayed() and b.is_enabled():
                driver.execute_script("arguments[0].click();", b)
                time.sleep(2)
                return True
        except Exception:
            continue
    return False


def lookup_person(driver, name, zip_code, attempt=0):
    """Returns (formatted_number_or_'-', status). Raises BlockedBySite."""
    global _timeouts_in_row
    print(f"🔍 Searching for: {name} in postal code {zip_code}")
    wait_for_turn()
    try:
        driver.get(BASE_URL + quote(f"{name} {zip_code}"))
        answered, body = wait_for_answer(driver)
        cards = read_cards(driver)
        haystack = driver.title.lower() + " " + body
    except Exception as e:
        msg = (str(e).splitlines() or [repr(e)])[0][:120]
        if attempt < 1:
            print(f"⚠️ Error ({msg}) - retrying once...")
            time.sleep(3)
            return lookup_person(driver, name, zip_code, attempt + 1)
        print(f"❌ Error during lookup: {msg}")
        return "-", "ERROR"

    hits = [w for w in BLOCK_WORDS if w in haystack]
    if hits and not cards:
        raise BlockedBySite(f"block words on page: {hits}")
    if not answered:
        _timeouts_in_row += 1
        print("⚠️ No answer from the site within 20 s")
        if _timeouts_in_row >= 3:
            raise BlockedBySite("3 timeouts in a row")
        return "-", "TIMEOUT"
    _timeouts_in_row = 0

    if not cards:
        if "ga ingen treff" in body:
            print("   Site: ingen treff (not in the directory)")
            return "-", "NO_HITS"
        print("   No result cards on the page")
        return "-", "NO_CARDS"

    number, status, note = choose_number(cards, name, zip_code)
    clicks = 0
    while status == "NO_MATCH" and clicks < 2 and click_load_more(driver):
        clicks += 1
        cards = read_cards(driver)
        number, status, note = choose_number(cards, name, zip_code)

    print(f"   {len(cards)} card(s) -> {status}" + (f" ({note})" if note else ""))
    if number:
        return pretty(number), status
    return "-", status


# ============================================================
# LIVE LOG: everything printed is also appended to telephone_live.log
# ============================================================
class _Tee:
    def __init__(self, stream, path, tag):
        self._stream = stream
        self._f = open(path, "a", encoding="utf-8", buffering=1)
        self._tag = tag
        self._buf = ""

    def write(self, text):
        self._stream.write(text)
        self._buf += text
        while "\n" in self._buf:
            line, self._buf = self._buf.split("\n", 1)
            if line.strip():
                self._f.write(f"[{self._tag}] {line}\n")
        self._f.flush()
        return len(text)

    def flush(self):
        try:
            self._stream.flush()
        except Exception:
            pass

    def __getattr__(self, name):
        return getattr(self._stream, name)


def setup_live_log(file_path):
    try:
        if os.path.exists(_LIVE_LOG) and os.path.getsize(_LIVE_LOG) > 5_000_000:
            os.replace(_LIVE_LOG, _LIVE_LOG + ".old")
        tag = os.path.basename(os.path.dirname(os.path.abspath(file_path))) or "run"
        sys.stdout = _Tee(sys.stdout, _LIVE_LOG, tag)
    except Exception as e:
        print(f"(live log disabled: {e})")


# ============================================================
# MAIN
# ============================================================
def private_sheets(sheetnames):
    """'Private', 'Private 2024', ... but not 'Selskap ...' or 'Private & Selskap ...'."""
    return [s for s in sheetnames if re.fullmatch(r"Private( \d{4})?", s.strip())]


def main(argv=None):
    argv = sys.argv if argv is None else argv
    if len(argv) < 2:
        print("Error: No file path provided!")
        print("Usage: python telephone_infotorg.py <excel_file_path>")
        sys.exit(1)

    file_path = argv[1]
    setup_live_log(file_path)
    print(f"Received file path: {file_path}")
    print(f"File exists: {os.path.exists(file_path)}")
    if not os.path.exists(file_path):
        print("ERROR: File does not exist!")
        sys.exit(1)
    print(f"File size: {os.path.getsize(file_path)} bytes")

    try:
        wb = load_workbook(file_path)
        print(f"✅ Workbook loaded. Sheets: {wb.sheetnames}")
    except Exception as e:
        print(f"❌ ERROR: Failed to load workbook: {e}")
        sys.exit(1)

    sheets = private_sheets(wb.sheetnames)
    print(f"📋 Sheets to process: {sheets}")

    driver = None

    def _cleanup():
        try:
            if driver is not None:
                driver.quit()
        except Exception:
            pass

    atexit.register(_cleanup)

    stats = {}
    blocked = False

    for sheet_name in sheets:
        if blocked:
            break
        print(f"\n📊 Processing sheet: {sheet_name}")
        print("=" * 50)
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            print(f"📋 Found {len(df)} rows")
            if "SHAREHOLDER" not in df.columns or "ZIP CODE / CITY" not in df.columns:
                print("⚠️ Required columns not found, skipping...")
                continue
            if "TELEPHONE" not in df.columns:
                df["TELEPHONE"] = ""
                print("📋 Added TELEPHONE column")

            phone_numbers = []
            for index, row in df.iterrows():
                name = str(row["SHAREHOLDER"]).strip()
                zip_city = str(row["ZIP CODE / CITY"]).strip()
                zip_code = extract_zip(zip_city)
                print(f"\n📋 Processing {index + 1}/{len(df)}: {name} ({zip_city})")

                if not name or name.lower() in ("nan", "none"):
                    phone_numbers.append("-")
                    continue
                if not zip_code:
                    print("   Skipped: no Norwegian postal code")
                    stats["SKIPPED_NO_ZIP"] = stats.get("SKIPPED_NO_ZIP", 0) + 1
                    phone_numbers.append("-")
                    continue

                existing = str(row.get("TELEPHONE")).strip() if pd.notna(row.get("TELEPHONE")) else ""
                if existing and existing != "-":
                    print(f"✅ Phone number already in the file: {existing}")
                    phone_numbers.append(existing)
                    continue

                cached = cache_get(name, zip_code)
                if cached:
                    print(f"💾 Phone taken from local cache: {cached}")
                    stats["CACHE"] = stats.get("CACHE", 0) + 1
                    phone_numbers.append(cached)
                    continue

                try:
                    if driver is None:
                        driver = make_driver()
                        print("🔧 Chrome started")
                    value, status = lookup_person(driver, name, zip_code)
                except BlockedBySite as e:
                    print(f"⛔ Infotorg is not usable right now ({e}).")
                    print("   Stopping now so the problem does not get worse. Try again later.")
                    blocked = True
                    phone_numbers.append(BLOCK_TEXT)
                    break
                stats[status] = stats.get(status, 0) + 1
                phone_numbers.append(value)
                if status == "FOUND_MOBILE":
                    cache_put(name, zip_code, value)
                print(f"📞 Result: {value}")

            if blocked:
                phone_numbers += [BLOCK_TEXT] * (len(df) - len(phone_numbers))

            ws = wb[sheet_name]
            telephone_col = next((i for i, c in enumerate(df.columns, 1) if c == "TELEPHONE"), None)
            if telephone_col and len(phone_numbers) == len(df):
                for i, value in enumerate(phone_numbers, start=2):
                    ws.cell(row=i, column=telephone_col, value=value)
                print(f"✅ Updated {len(phone_numbers)} phone numbers in sheet '{sheet_name}'")
            else:
                print(f"⚠️ Could not update sheet '{sheet_name}' - column mismatch")
        except Exception as e:
            print(f"❌ Error processing sheet '{sheet_name}': {e}")
            continue

    print("\n================ SUMMARY ================")
    for status in sorted(stats):
        print(f"  {status:20s} {stats[status]}")
    print("STOPPED EARLY (blocked)" if blocked else "Finished without stopping.")

    try:
        wb.save(file_path)
        wb.close()
        print("\n💾 Workbook saved successfully")
        file_dir, file_name = os.path.split(file_path)
        file_base, file_ext = os.path.splitext(file_name)
        new_file_name = f"{file_base} (T){file_ext}"
        os.replace(file_path, os.path.join(file_dir, new_file_name))
        print(f"✅ SUCCESS: Phone numbers added and file saved as: {new_file_name}")
    except Exception as e:
        print(f"❌ Error saving workbook: {e}")

    if driver is not None:
        try:
            driver.quit()
        except Exception:
            pass
        driver = None
    print("🔄 Browser closed. Process completed!")


if __name__ == "__main__":
    main()
