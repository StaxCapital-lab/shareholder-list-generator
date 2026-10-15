# ============================================================
# FORCE UTF-8 OUTPUT (fixes UnicodeEncodeError when this script is
# launched via subprocess, e.g. from app3.py/app5.py). When output is
# piped/captured instead of going to a real console, Python falls back to
# the system codepage (e.g. cp1251 on Windows), which can't encode the
# emoji used in the print() messages below. Must run before any print().
# ============================================================
import sys
import io

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
else:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import re
from openpyxl import load_workbook
import sys
import os
import random
import logging
import platform
import atexit

# Get the file path from the command-line argument
if len(sys.argv) < 2:
    print("Error: No file path provided!")
    print("Usage: python3 telephone_1881_batch_updated.py <excel_file_path>")
    sys.exit(1)

file_path = sys.argv[1]  # dynamically gets the correct file path from app.py
sheets = ["Private 2024", "Selskap 2024"]  # Sheets to process
# "Selskap 2024" is kept in the workbook for feature parity, but company
# shareholders are intentionally NOT searched on 1881. Their pages use a
# different layout/URL and this lookup is person-only.

# Setup logging for debugging. The log file is written next to the input
# Excel file (inside this run's own session folder) instead of a single
# shared "telephone_debug.log" in the program folder - otherwise several
# people running this on the server at the same time would overwrite each
# other's logs.
log_path = os.path.join(os.path.dirname(file_path) or '.', 'telephone_debug.log')
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_path),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Debug: Print file path and check if file exists
print(f"Received file path: {file_path}")
print(f"File exists: {os.path.exists(file_path)}")
if os.path.exists(file_path):
    print(f"File size: {os.path.getsize(file_path)} bytes")
    print(f"File extension: {os.path.splitext(file_path)[1]}")
else:
    print("ERROR: File does not exist!")
    sys.exit(1)

# ============================================================
# VIRTUAL DISPLAY (Xvfb) FOR SERVER DEPLOYMENT
# ============================================================
# This script never uses Chrome's --headless flag, because headless Chrome
# is much more likely to get flagged as a bot by sites like this one. On a
# server (no physical screen), we instead start a virtual display (Xvfb)
# and run Chrome normally against it - Chrome behaves exactly like a
# regular visible browser, it just isn't shown on a real monitor.
# Requires on the server: `sudo apt install xvfb` and
# `pip install pyvirtualdisplay`. Locally on Windows this block is skipped
# entirely, so a real Chrome window still opens, same as before.
virtual_display = None
used_headless_fallback = False

if platform.system() == "Linux":
    try:
        from pyvirtualdisplay import Display
        virtual_display = Display(visible=False, size=(1920, 1080))
        virtual_display.start()
        atexit.register(virtual_display.stop)
        print("🖥️ Virtual display (Xvfb) started for server environment")
    except Exception as e:
        print(f"⚠️ Could not start virtual display ({e}) - falling back to --headless=new")
        used_headless_fallback = True

# Initialize WebDriver with enhanced options
options = webdriver.ChromeOptions()

# Enhanced options for better stability and stealth
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("--disable-extensions")
options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
options.add_experimental_option('useAutomationExtension', False)

if used_headless_fallback:
    options.add_argument("--headless=new")

# User agent to appear more like a regular browser
user_agents = [
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
]
selected_user_agent = random.choice(user_agents)
options.add_argument(f"--user-agent={selected_user_agent}")

driver = None


def _cleanup():
    """Makes sure Chrome (and the virtual display, if used) never stay
    hanging in memory, however the script exits - runs automatically at
    program exit, including on crashes and sys.exit()."""
    try:
        if driver is not None:
            driver.quit()
    except Exception:
        pass


atexit.register(_cleanup)

try:
    # Selenium 4.6+ includes Selenium Manager, which automatically
    # finds/downloads the ChromeDriver matching the installed Chrome.
    # Do NOT use webdriver-manager here: its global /root/.wdm lock
    # causes concurrent users to block each other.
    driver = webdriver.Chrome(options=options)

    # Hide webdriver property
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

    print("🔧 Chrome driver setup completed")
except Exception as e:
    print(f"❌ Failed to start Chrome driver: {e}")
    sys.exit(1)

# Navigate to 1881.no
print("\n🌐 ACCESSING 1881.NO...")
print("="*60)

try:
    driver.get("https://www.1881.no/")
    print("✅ Page request sent to 1881.no")
    time.sleep(random.uniform(3, 6))
except Exception as e:
    print(f"❌ Error accessing 1881.no: {e}")
    driver.quit()
    sys.exit(1)

def handle_consent_popup():
    """Handle the consent popup that appears on 1881.no"""
    print("\n🔍 HANDLING CONSENT POPUP...")
    print("="*50)
    
    try:
        time.sleep(3)
        
        # Check if consent popup is present
        page_source = driver.page_source.lower()
        if "give us your consent" in page_source:
            print("🔍 Consent popup detected, attempting to close...")
            
            # Try to find and click the "Agree and close" button
            button_selectors = [
                "button[class*='agree']",
                "button:contains('Agree and close')",
                ".consent-button",
                "[data-consent='accept']",
            ]
            
            for selector in button_selectors:
                try:
                    button = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                    )
                    driver.execute_script("arguments[0].click();", button)
                    print("✅ Consent popup closed successfully!")
                    time.sleep(2)
                    return True
                except:
                    continue
            
            print("⚠️ Could not automatically close consent popup")
            return False
        else:
            print("✅ No consent popup detected")
            return True
            
    except Exception as e:
        print(f"❌ Error handling consent popup: {e}")
        return False

def normalize_name(name):
    """Normalize a name for robust comparison.

    We compare individual words rather than requiring the exact string,
    because 1881 can display names in a different order or with small
    formatting differences.
    """
    if not name:
        return []
    name = re.sub(r"[^\wæøåÆØÅ-]+", " ", str(name), flags=re.UNICODE).lower()
    return [part for part in name.split() if part]


def verify_person_match(page_name, search_name, zip_code=None):
    """Return True when the displayed person name matches the searched name.

    All words from the searched name must be present in the displayed name,
    allowing the order to differ. This prevents accidentally accepting a
    relative/person with the same surname.
    """
    try:
        page_words = normalize_name(page_name)
        search_words = normalize_name(search_name)

        if not page_words or not search_words:
            return False

        page_set = set(page_words)
        search_set = set(search_words)

        name_match = search_set.issubset(page_set)

        if not name_match:
            print(
                f"⚠️ Name mismatch: expected '{search_name}', "
                f"found '{page_name}'"
            )
            return False

        print(f"✅ Name match confirmed: '{page_name}'")
        return True

    except Exception as e:
        # Do NOT accept the result if verification itself fails.
        print(f"⚠️ Could not verify person match: {e}")
        return False


def extract_postal_code_from_address(address_text):
    """Extract postal code from address text"""
    try:
        # Look for 4-digit postal code pattern
        postcode_match = re.search(r'(\d{4})', address_text)
        if postcode_match:
            return postcode_match.group(1)
    except:
        pass
    return None

def get_all_postal_codes_on_page():
    """Get all postal codes visible on the current page"""
    try:
        page_text = driver.page_source
        # Find all 4-digit postal codes on the page
        postal_codes = re.findall(r'\b(\d{4})\b', page_text)
        return list(set(postal_codes))  # Return unique postal codes
    except:
        pass
    return []

def verify_phone_postal_code_match(phone_element, expected_zip_code):
    """
    Verify if the phone number's address matches the expected postal code.
    This helps identify when a person has multiple addresses.

    Strategy: Check if the expected postal code is visible on the page
    and appears near the phone number in the page structure.
    """
    try:
        # Get all postal codes on the page
        page_postal_codes = get_all_postal_codes_on_page()

        # If the expected postal code is not on the page at all, it's a mismatch
        if expected_zip_code not in page_postal_codes:
            print(f"📍 Expected postal code {expected_zip_code} NOT found on page")
            print(f"   Available postal codes: {page_postal_codes}")
            return False

        # If the expected postal code is on the page, it's likely a match
        # (since we already verified the person name matches)
        print(f"📍 Expected postal code {expected_zip_code} found on page")
        print(f"   Available postal codes: {page_postal_codes}")
        return True

    except Exception as e:
        print(f"⚠️ Error verifying postal code: {e}")
        pass

    return None  # Could not determine

def clean_phone_number(phone_number):
    """Clean and format Norwegian phone number"""
    if not phone_number:
        return phone_number
    
    # Remove all non-digit characters except +
    cleaned = re.sub(r'[^\d\+]', '', phone_number)
    
    # Remove Norwegian country codes
    if cleaned.startswith('0047'):
        cleaned = cleaned[4:]  # Remove 0047
    elif cleaned.startswith('+47'):
        cleaned = cleaned[3:]  # Remove +47
    elif cleaned.startswith('47') and len(cleaned) == 10:
        cleaned = cleaned[2:]  # Remove 47 if it's a 10-digit number starting with 47
    
    # Ensure it's 8 digits (standard Norwegian mobile/landline)
    if len(cleaned) == 8 and cleaned.isdigit():
        # Format as XX XX XX XX
        return f"{cleaned[:2]} {cleaned[2:4]} {cleaned[4:6]} {cleaned[6:8]}"
    
    # If not standard 8 digits, return as-is but cleaned
    return cleaned

def get_search_result_links():
    """Collect candidate person links from the current 1881 results page.

    We deliberately only collect /person/ links. Company pages are ignored.
    Returns a de-duplicated list of (element, href, visible_text).
    """
    selectors = [
        "a[href*='/person/']",
        ".search-result a[href*='/person/']",
        ".result-item a[href*='/person/']",
        "div[class*='result'] a[href*='/person/']",
    ]

    candidates = []
    seen = set()

    for selector in selectors:
        try:
            elements = WebDriverWait(driver, 5).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, selector))
            )
        except Exception:
            continue

        for elem in elements:
            try:
                href = elem.get_attribute("href") or ""
                text_value = elem.text.strip()
                if not href or "/person/" not in href:
                    continue

                key = href.split("#")[0]
                if key in seen:
                    continue

                seen.add(key)
                candidates.append((elem, href, text_value))
            except Exception:
                continue

        if candidates:
            # Keep the first selector that gives useful person links, but
            # continue deduplicating from the other selectors if available.
            continue

    return candidates


def print_search_page_diagnostics():
    """Print useful information when 1881 gives an unexpected results page."""
    try:
        print("🔎 SEARCH PAGE DIAGNOSTICS")
        print(f"📍 URL: {driver.current_url}")

        links = driver.find_elements(By.CSS_SELECTOR, "a[href]")
        printed = 0
        seen = set()

        for link in links:
            try:
                href = link.get_attribute("href") or ""
                text_value = " ".join(link.text.split())

                if not href:
                    continue

                key = (href, text_value)
                if key in seen:
                    continue
                seen.add(key)

                # Print likely search/result links first, then a limited
                # number of other links so logs stay manageable.
                if (
                    "/person/" in href
                    or "/company/" in href
                    or "#gsc" in href
                    or "search" in href.lower()
                ):
                    print(f"   LINK: text='{text_value[:120]}' href='{href}'")
                    printed += 1
                    if printed >= 30:
                        break
            except Exception:
                continue

        if printed == 0:
            print("   No likely result/search links found on the page.")

        print(f"📄 Page title: {driver.title}")
    except Exception as e:
        print(f"⚠️ Could not print search page diagnostics: {e}")


def get_phone_number(name, zip_code, retry_count=0, max_retries=2):
    """Extract phone number from a matching 1881 person page."""
    try:
        print(f"🔍 Searching for: {name} in postal code {zip_code}")

        # Find search input
        search_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='search']"))
        )

        # Clear and enter search term (include postal code in search)
        search_term = f"{name} {zip_code}" if zip_code else name
        search_input.clear()
        search_input.send_keys(search_term)

        # Submit search
        search_input.submit()

        # Wait for navigation/results to settle
        time.sleep(3)

        current_url = driver.current_url
        print(f"📍 Current URL: {current_url}")

        # ------------------------------------------------------------
        # CASE 1: 1881 directly opened a person detail page.
        # ------------------------------------------------------------
        if "/person/" in current_url:
            print("✅ 1881 directly opened a person detail page!")

        # ------------------------------------------------------------
        # CASE 2: We are on a results/search page.
        # Find ALL person results and choose the one matching the
        # requested name, instead of blindly clicking the first result.
        # ------------------------------------------------------------
        else:
            print("📋 On search/results page - looking for matching person...")

            candidates = get_search_result_links()
            print(f"👥 Found {len(candidates)} person result link(s)")

            matching_candidate = None

            for _, href, visible_text in candidates:
                print(
                    f"   Candidate: '{visible_text}' -> {href}"
                )

                # First use the visible result text as a quick filter.
                if visible_text and verify_person_match(
                    visible_text, name, zip_code
                ):
                    matching_candidate = href
                    print(f"✅ Matching result selected: {href}")
                    break

                # If visible text is incomplete, open the candidate in the
                # same tab and verify the actual person page name.
                try:
                    original_url = driver.current_url
                    driver.get(href)
                    time.sleep(2)

                    candidate_url = driver.current_url
                    if "/person/" not in candidate_url:
                        driver.get(original_url)
                        time.sleep(1)
                        continue

                    name_elements = driver.find_elements(
                        By.CSS_SELECTOR, ".details-name, h1"
                    )
                    candidate_name = (
                        name_elements[0].text.strip()
                        if name_elements else ""
                    )

                    if verify_person_match(candidate_name, name, zip_code):
                        matching_candidate = candidate_url
                        print(
                            f"✅ Matching person page selected: "
                            f"{candidate_name}"
                        )
                        break

                    driver.get(original_url)
                    time.sleep(1)

                except Exception as candidate_error:
                    print(
                        f"⚠️ Could not inspect candidate '{href}': "
                        f"{candidate_error}"
                    )
                    try:
                        driver.back()
                        time.sleep(1)
                    except Exception:
                        pass

            if not matching_candidate:
                print(
                    f"❌ No person result matched '{name}' "
                    f"in the search results."
                )
                print_search_page_diagnostics()
                return "-"

            # If the matching candidate was not already opened as a page,
            # navigate to it now.
            if "/person/" not in driver.current_url or (
                matching_candidate != driver.current_url
            ):
                driver.get(matching_candidate)
                time.sleep(3)

            current_url = driver.current_url
            print(f"📍 Selected person URL: {current_url}")

            if "/person/" not in current_url:
                print("❌ Selected result is not a person detail page")
                print_search_page_diagnostics()
                return "-"

        # ------------------------------------------------------------
        # Verify the person page one more time before extracting a phone.
        # ------------------------------------------------------------
        print("✅ Successfully navigated to person detail page!")

        try:
            name_elements = driver.find_elements(
                By.CSS_SELECTOR, ".details-name, h1"
            )
            page_name = (
                name_elements[0].text.strip()
                if name_elements else ""
            )

            print(f"👤 Found person: {page_name}")

            if not page_name or not verify_person_match(
                page_name, name, zip_code
            ):
                print("❌ Person verification failed")
                return "-"
        except Exception as e:
            print(f"⚠️ Could not verify person name: {e}")
            # Safety first: do not return a phone from an unverified page.
            return "-"

        # Look for postal code in the address
        try:
            address_elements = driver.find_elements(
                By.CSS_SELECTOR,
                ".listing-address, .address, .location"
            )
            address_found = False
            found_postal_codes = []

            for addr_elem in address_elements:
                address_text = addr_elem.text
                print(f"📍 Found address: {address_text}")

                extracted_zip = extract_postal_code_from_address(address_text)
                if extracted_zip:
                    found_postal_codes.append(extracted_zip)
                    print(f"   Extracted postal code: {extracted_zip}")

                if zip_code and zip_code in address_text:
                    print(f"✅ Postal code {zip_code} confirmed in address!")
                    address_found = True
                    break

            # If the site does not expose the address in those selectors,
            # fall back to the full page source.
            if not address_found and zip_code and zip_code in driver.page_source:
                print(f"✅ Postal code {zip_code} found in page source!")
                address_found = True

            # Only enforce postal-code matching when a postcode was supplied.
            if zip_code and not address_found:
                print(f"⚠️ Could not confirm postal code {zip_code}")
                print(f"   Found postal codes on page: {found_postal_codes}")
                return "-"
        except Exception as e:
            print(f"⚠️ Could not check address: {e}")
            if zip_code:
                return "-"

        # Now extract phone number
        print("📞 Looking for phone number...")

        page_postal_codes = get_all_postal_codes_on_page()
        print(f"📍 Postal codes found on page: {page_postal_codes}")

        # Strategy 1: Look for phone links
        try:
            phone_links = driver.find_elements(By.CSS_SELECTOR, "a[href^='tel:']")
            print(f"📞 Found {len(phone_links)} phone link(s)")

            for i, link in enumerate(phone_links):
                phone_href = link.get_attribute("href")
                phone_number = phone_href.replace("tel:", "").strip()
                if phone_number:
                    cleaned_phone = clean_phone_number(phone_number)
                    print(f"📞 Phone link {i+1}: {cleaned_phone}")

                    if zip_code and zip_code not in page_postal_codes:
                        print(
                            f"⚠️ Skipping phone {cleaned_phone} - "
                            f"postal code {zip_code} not found on page"
                        )
                        continue

                    print(f"✅ Returning phone number: {cleaned_phone}")
                    return cleaned_phone
        except Exception as e:
            print(f"⚠️ Error processing phone links: {e}")

        # Strategy 2: Look for phone in specific elements
        phone_selectors = [
            ".button-call__number",
            ".contact-featured__call",
            ".phone",
            ".telephone",
            ".tlf",
            ".mobil",
            ".contact-phone"
        ]

        for selector in phone_selectors:
            try:
                phone_elements = driver.find_elements(By.CSS_SELECTOR, selector)
                print(
                    f"📞 Found {len(phone_elements)} element(s) "
                    f"with selector '{selector}'"
                )

                for i, elem in enumerate(phone_elements):
                    phone_text = elem.text.strip()
                    if phone_text and re.search(r'\d', phone_text):
                        phone_clean = re.sub(r'[^\d\+\s]', '', phone_text)
                        if len(phone_clean.replace(' ', '')) >= 8:
                            cleaned_phone = clean_phone_number(phone_clean)
                            print(f"📞 Phone element {i+1}: {cleaned_phone}")

                            if zip_code and zip_code not in page_postal_codes:
                                print(
                                    f"⚠️ Skipping phone {cleaned_phone} - "
                                    f"postal code {zip_code} not found on page"
                                )
                                continue

                            print(
                                f"📍 Postal code {zip_code} confirmed on page"
                            )
                            print(f"✅ Returning phone number: {cleaned_phone}")
                            return cleaned_phone
            except Exception as e:
                print(f"⚠️ Error with selector '{selector}': {e}")
                continue

        # Strategy 3: Look for phone in page text using regex
        page_text = driver.page_source

        phone_patterns = [
            r'(\+47\s?)?(\d{2}\s?\d{2}\s?\d{2}\s?\d{2})',
            r'(\+47\s?)?\d{8}',
            r'tel:(\+47)?(\d{8})',
        ]

        for pattern in phone_patterns:
            matches = re.findall(pattern, page_text)
            for match in matches:
                if isinstance(match, tuple):
                    phone_number = ''.join(match)
                else:
                    phone_number = match

                phone_clean = re.sub(r'[^\d\+]', '', phone_number)
                if len(phone_clean) >= 8:
                    cleaned_phone = clean_phone_number(phone_clean)
                    print(f"📞 Found phone number with regex: {cleaned_phone}")
                    return cleaned_phone

        print("❌ No phone number found on person page")
        return "-"

    except Exception as e:
        print(f"❌ Error during phone extraction: {e}")

        if retry_count < max_retries:
            print(
                f"🔄 Retrying after error "
                f"(attempt {retry_count + 1}/{max_retries})..."
            )
            time.sleep(3)
            return get_phone_number(
                name, zip_code, retry_count + 1, max_retries
            )

        return "-"

# Handle consent popup
consent_handled = handle_consent_popup()

if not consent_handled:
    print("⚠️ Consent popup handling failed, but continuing...")

print("✅ Successfully loaded 1881.no and ready to extract phone numbers")

# Load existing workbook
try:
    print(f"\n📊 Loading workbook: {file_path}")
    wb = load_workbook(file_path)
    print("✅ Workbook loaded successfully")
    print(f"📋 Available sheets: {wb.sheetnames}")
except Exception as e:
    print(f"❌ ERROR: Failed to load workbook: {e}")
    driver.quit()
    sys.exit(1)

# Process each sheet separately
for sheet_name in sheets:
    if sheet_name not in wb.sheetnames:
        print(f"⚠️ Sheet '{sheet_name}' not found in workbook, skipping...")
        continue
        
    print(f"\n📊 Processing sheet: {sheet_name}")
    print("="*50)

    try:
        # Load data from the current sheet
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        print(f"📋 Found {len(df)} rows in sheet '{sheet_name}'")
        
        # Check if required columns exist
        if "SHAREHOLDER" not in df.columns or "ZIP CODE / CITY" not in df.columns:
            print(f"⚠️ Required columns not found in sheet '{sheet_name}', skipping...")
            continue
        
        # Add TELEPHONE column if it doesn't exist
        if "TELEPHONE" not in df.columns:
            df["TELEPHONE"] = ""
            print("📋 Added TELEPHONE column")

        # List to store phone numbers
        phone_numbers = []

        # Iterate through dataframe and fetch phone numbers
        for index, row in df.iterrows():
            name = str(row["SHAREHOLDER"]).strip()

            # Extract postal code properly from "ZIP CODE / CITY" format
            zip_city = str(row["ZIP CODE / CITY"]).strip()
            # Extract only the numeric part (postal code) - usually first 4 digits
            postcode_match = re.match(r'(\d{4})', zip_city)
            postcode = postcode_match.group(1) if postcode_match else zip_city[:4]

            print(f"📮 Extracted postal code: '{postcode}' from '{zip_city}'")

            # Company shareholders are intentionally not searched on 1881.
            # The current phone lookup is person-only; company result pages
            # use a different URL/layout and must not be mistaken for people.
            if sheet_name == "Selskap 2024":
                phone_numbers.append("-")
                print(f"🏢 Skipping company lookup: {name}")
                continue

            # Skip if name is empty or NaN
            if not name or name.lower() in ['nan', 'none', '']:
                phone_numbers.append("-")
                continue

            print(f"\n📋 Processing {index + 1}/{len(df)}: {name} (Zip: {postcode})")

            # Check if phone number already exists
            if pd.notna(row.get("TELEPHONE")) and str(row.get("TELEPHONE")).strip() and str(row.get("TELEPHONE")).strip() != "-":
                existing_phone = str(row["TELEPHONE"]).strip()
                print(f"✅ Phone number already exists: {existing_phone}")
                phone_numbers.append(existing_phone)
                continue

            phone_number = get_phone_number(name, postcode)
            phone_numbers.append(phone_number)

            print(f"📞 Result: {phone_number if phone_number else 'No phone found'}")

            # Random delay to avoid being detected as bot
            time.sleep(random.uniform(2, 4))

        # Update the worksheet with phone numbers
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]

            # Find the TELEPHONE column
            telephone_col = None
            for col_idx, col in enumerate(df.columns, 1):
                if col == "TELEPHONE":
                    telephone_col = col_idx
                    break

            if telephone_col and len(phone_numbers) == len(df):
                for i, value in enumerate(phone_numbers, start=2):  # Assuming row 1 contains headers
                    ws.cell(row=i, column=telephone_col, value=value)
                print(f"✅ Updated {len(phone_numbers)} phone numbers in sheet '{sheet_name}'")
            else:
                print(f"⚠️ Could not update sheet '{sheet_name}' - column mismatch")

    except Exception as e:
        print(f"❌ Error processing sheet '{sheet_name}': {e}")
        continue

# Save and close workbook
try:
    wb.save(file_path)
    wb.close()
    print(f"\n💾 Workbook saved successfully")

    # Create a copy with (T) suffix
    file_dir, file_name = os.path.split(file_path)
    file_base, file_ext = os.path.splitext(file_name)

    new_file_name = f"{file_base} (T){file_ext}"
    new_file_path = os.path.join(file_dir, new_file_name)

    os.rename(file_path, new_file_path)

    print(f"✅ SUCCESS: Phone numbers added and file saved as: {new_file_name}")

except Exception as e:
    print(f"❌ Error saving workbook: {e}")

# Quit Selenium driver
driver.quit()
print("🔄 Browser closed. Process completed!")
