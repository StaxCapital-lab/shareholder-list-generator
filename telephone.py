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
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
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
sheets = ["Private 2024"]  # Only private individuals.
# NOTE: company lookups ("Selskap 2024") were tried and removed again -
# 1881.no opens a company page at a different URL pattern (no "/person/"
# in it), which this script doesn't know how to read, so it always came
# back empty anyway. Easier to just not search for companies here at all.

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
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()), options=options
    )

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

def is_driver_alive():
    """Checks whether the current Chrome session is still responding -
    used to tell a page-level problem (name not found, etc.) apart from
    the whole browser process having crashed."""
    try:
        _ = driver.current_url
        return True
    except Exception:
        return False


def recreate_driver():
    """
    Starts a fresh Chrome instance after a crash (the browser process can
    occasionally die mid-run, for reasons that don't show up as an OOM
    kill or any other clear system-level cause) and re-does the initial
    1881.no load + consent popup, so processing can continue from a clean
    browser instead of every subsequent lookup failing the same way.
    """
    global driver
    print("🔁 Chrome appears to have crashed - starting a new browser instance...")
    try:
        if driver is not None:
            driver.quit()
    except Exception:
        pass

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()), options=options
    )
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    driver.get("https://www.1881.no/")
    time.sleep(random.uniform(3, 6))
    handle_consent_popup()
    print("✅ New Chrome instance ready - resuming")


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

def verify_person_match(page_name, search_name, zip_code):
    """Verify if the person on the page matches the search criteria"""
    try:
        # Normalize names for comparison
        page_name_normalized = page_name.lower().strip()
        search_name_normalized = search_name.lower().strip()

        # Check if names match (allow for slight variations)
        name_match = (search_name_normalized in page_name_normalized or
                     page_name_normalized in search_name_normalized or
                     search_name_normalized.split()[0] in page_name_normalized)

        if not name_match:
            print(f"⚠️ Name mismatch: expected '{search_name}', found '{page_name}'")
            return False

        print(f"✅ Name match confirmed: '{page_name}'")
        return True

    except Exception as e:
        print(f"⚠️ Could not verify person match: {e}")
        return True  # Continue anyway if verification fails

def names_match(candidate_text, search_name):
    """
    Checks whether a search-result's visible text (name, possibly plus age
    and city, e.g. "Christopher Heidenreich, 34, Oslo") contains every word
    of the name we're searching for - regardless of word order. Used to
    pick the correct result out of several, instead of always clicking the
    first one on the list.
    """
    def words(text):
        return set(re.findall(r"[a-zA-ZæøåÆØÅ]+", text.lower()))

    search_words = words(search_name)
    if not search_words:
        return False
    candidate_words = words(candidate_text)
    return search_words.issubset(candidate_words)


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

def get_phone_number(name, zip_code, retry_count=0, max_retries=2):
    """Extract phone number using the improved method with retry logic"""
    try:
        print(f"🔍 Searching for: {name} in postal code {zip_code}")

        # Find search input
        search_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='search']"))
        )

        # Clear and enter search term (include postal code in search)
        # Search by name only. Appending the postal code here used to
        # confuse 1881.no's own search - instead of showing a normal list
        # of name candidates, it would sometimes jump straight to an
        # unrelated "closest match" person (e.g. a relative with the same
        # surname at a similar address). The postal code is still checked
        # further down, once we're on a candidate's actual page, so this
        # doesn't weaken the "wrong person" protection.
        search_term = name
        search_input.clear()
        search_input.send_keys(search_term)

        # Submit search
        search_input.submit()

        # Wait for search results to load
        time.sleep(3)

        current_url = driver.current_url
        print(f"📍 Current URL: {current_url}")

        # Check if we're on a search results page (not a person detail page yet)
        if "/person/" not in current_url:
            print("📋 On search results page, attempting to click first result...")

            # Try to find and click the first search result
            try:
                # Wait for search results to appear with multiple selector options
                result_selectors = [
                    "a[href*='/person/']",
                    ".search-result a",
                    ".result-item a",
                    "div[class*='result'] a"
                ]

                all_results = []
                for selector in result_selectors:
                    try:
                        WebDriverWait(driver, 5).until(
                            EC.presence_of_all_elements_located((By.CSS_SELECTOR, selector))
                        )
                        found = driver.find_elements(By.CSS_SELECTOR, selector)
                        if found:
                            all_results = found
                            print(f"✅ Found {len(found)} result(s) with selector: {selector}")
                            break
                    except:
                        continue

                if not all_results:
                    print("❌ Could not find any search results")
                    # DIAGNOSTIC: log every link on the page, so cases like
                    # a #gsc... Google-custom-search results page (different
                    # markup our selectors don't match) can be inspected
                    # afterwards instead of just guessing.
                    try:
                        diag_links = driver.find_elements(By.TAG_NAME, "a")
                        print(f"🔍 DIAGNOSTIC: page has {len(diag_links)} <a> links total. First 40:")
                        for link in diag_links[:40]:
                            href = link.get_attribute('href')
                            text = link.text.strip()
                            if href:
                                print(f"   - [{text}] -> {href}")
                    except Exception as diag_error:
                        print(f"⚠️ Could not collect diagnostic links: {diag_error}")
                    return "-"

                # Among all results, skip company pages (their links don't
                # contain "/person/") and pick the one whose visible name
                # matches every word of the name we searched for - rather
                # than always clicking the first result on the list.
                chosen_result = None
                for result in all_results:
                    href = result.get_attribute('href') or ""
                    if "/person/" not in href:
                        continue  # company result - not handled here
                    result_text = result.text.strip()
                    if names_match(result_text, name):
                        chosen_result = result
                        print(f"✅ Name match in results list: '{result_text}' -> selecting this one")
                        break
                    else:
                        print(f"⏭️ Skipping non-matching result: '{result_text}' ({href})")

                if chosen_result is None:
                    print(f"❌ None of the {len(all_results)} result(s) matched the name '{name}'")
                    # DIAGNOSTIC: log what was actually on the results list,
                    # so a mismatch (wrong site layout, company-only
                    # results, etc.) can be diagnosed from the log.
                    for result in all_results:
                        href = result.get_attribute('href') or ""
                        print(f"   - candidate: '{result.text.strip()}' -> {href}")
                    return "-"

                print(f"🔗 Clicking matched result link: {chosen_result.get_attribute('href')}")

                # Click the matched result
                driver.execute_script("arguments[0].click();", chosen_result)

                # Wait for person detail page to load
                time.sleep(3)

                # Verify we're now on a person page
                current_url = driver.current_url
                print(f"📍 After click - Current URL: {current_url}")

                if "/person/" not in current_url:
                    print("❌ Failed to navigate to person detail page after clicking")
                    return "-"

            except Exception as e:
                print(f"❌ Error clicking first result: {e}")
                return "-"

        # Now we should be on a person detail page
        if "/person/" in current_url:
            print("✅ Successfully navigated to person detail page!")
            
            # Verify this is the correct person by checking the name
            try:
                name_element = driver.find_element(By.CSS_SELECTOR, ".details-name, h1")
                page_name = name_element.text.strip()
                print(f"👤 Found person: {page_name}")

                # Verify the person matches our search criteria
                if not verify_person_match(page_name, name, zip_code):
                    print(f"❌ Person verification failed")
                    return "-"
            except:
                print("⚠️ Could not verify person name")
            
            # Look for postal code in the address
            try:
                address_elements = driver.find_elements(By.CSS_SELECTOR, ".listing-address, .address, .location")
                address_found = False
                found_postal_codes = []

                for addr_elem in address_elements:
                    address_text = addr_elem.text
                    print(f"📍 Found address: {address_text}")

                    # Extract postal code from address
                    extracted_zip = extract_postal_code_from_address(address_text)
                    if extracted_zip:
                        found_postal_codes.append(extracted_zip)
                        print(f"   Extracted postal code: {extracted_zip}")

                    # Check if postal code is in the address
                    if zip_code in address_text:
                        print(f"✅ Postal code {zip_code} confirmed in address!")
                        address_found = True
                        break

                if not address_found:
                    # Check page source for postal code
                    if zip_code in driver.page_source:
                        print(f"✅ Postal code {zip_code} found in page source!")
                        address_found = True

                if not address_found:
                    print(f"⚠️ Could not confirm postal code {zip_code}")
                    print(f"   Found postal codes on page: {found_postal_codes}")
                    return "-"
            except:
                print("⚠️ Could not check address")
            
            # Now extract phone number
            print("📞 Looking for phone number...")

            # Get all postal codes on the page first
            page_postal_codes = get_all_postal_codes_on_page()
            print(f"📍 Postal codes found on page: {page_postal_codes}")

            # Strategy 1: Look for phone links
            try:
                phone_links = driver.find_elements(By.CSS_SELECTOR, "a[href^='tel:']")
                print(f"📞 Found {len(phone_links)} phone link(s)")

                for i, link in enumerate(phone_links):
                    phone_href = link.get_attribute('href')
                    phone_number = phone_href.replace('tel:', '').strip()
                    if phone_number:
                        cleaned_phone = clean_phone_number(phone_number)
                        print(f"📞 Phone link {i+1}: {cleaned_phone}")

                        # Verify postal code match if we have a zip code to verify
                        if zip_code:
                            # Check if the expected postal code is on the page
                            if zip_code not in page_postal_codes:
                                print(f"⚠️ Skipping phone {cleaned_phone} - postal code {zip_code} not found on page")
                                continue
                            else:
                                print(f"✅ Postal code {zip_code} confirmed on page")

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
                    print(f"📞 Found {len(phone_elements)} element(s) with selector '{selector}'")

                    for i, elem in enumerate(phone_elements):
                        phone_text = elem.text.strip()
                        if phone_text and re.search(r'\d', phone_text):
                            # Clean up phone number
                            phone_clean = re.sub(r'[^\d\+\s]', '', phone_text)
                            if len(phone_clean.replace(' ', '')) >= 8:  # Norwegian phone numbers are 8 digits
                                cleaned_phone = clean_phone_number(phone_clean)
                                print(f"📞 Phone element {i+1}: {cleaned_phone}")

                                # Verify postal code match if we have a zip code to verify
                                if zip_code:
                                    # Check if the expected postal code is on the page
                                    if zip_code not in page_postal_codes:
                                        print(f"⚠️ Skipping phone {cleaned_phone} - postal code {zip_code} not found on page")
                                        continue
                                    else:
                                        print(f"✅ Postal code {zip_code} confirmed on page")

                                print(f"✅ Returning phone number: {cleaned_phone}")
                                return cleaned_phone
                except Exception as e:
                    print(f"⚠️ Error with selector '{selector}': {e}")
                    continue
            
            # Strategy 3: Look for phone in page text using regex
            page_text = driver.page_source
            
            # Norwegian phone number patterns
            phone_patterns = [
                r'(\+47\s?)?(\d{2}\s?\d{2}\s?\d{2}\s?\d{2})',  # +47 XX XX XX XX or XX XX XX XX
                r'(\+47\s?)?\d{8}',  # +47XXXXXXXX or XXXXXXXX
                r'tel:(\+47)?(\d{8})',  # tel: links
            ]
            
            for pattern in phone_patterns:
                matches = re.findall(pattern, page_text)
                for match in matches:
                    if isinstance(match, tuple):
                        phone_number = ''.join(match)
                    else:
                        phone_number = match
                    
                    # Clean and validate
                    phone_clean = re.sub(r'[^\d\+]', '', phone_number)
                    if len(phone_clean) >= 8:
                        cleaned_phone = clean_phone_number(phone_clean)
                        print(f"📞 Found phone number with regex: {cleaned_phone}")
                        return cleaned_phone
            
            print("❌ No phone number found on person page")
            return "-"
            
        else:
            print("❌ Did not reach person detail page")
            print("📋 Current page might be search results or error page")

            # Retry with just the name (without postal code)
            if retry_count < max_retries:
                print(f"🔄 Retrying with just name (attempt {retry_count + 1}/{max_retries})...")
                time.sleep(2)
                return get_phone_number(name, "", retry_count + 1, max_retries)

            return "-"

    except Exception as e:
        print(f"❌ Error during phone extraction: {e}")

        # If the browser process itself crashed (not just a page-level
        # problem), the same error would repeat for every remaining
        # person unless we start a new Chrome instance first.
        if not is_driver_alive():
            try:
                recreate_driver()
            except Exception as recreate_error:
                print(f"❌ Could not recreate Chrome driver: {recreate_error}")
                return "-"

        # Retry on error
        if retry_count < max_retries:
            print(f"🔄 Retrying after error (attempt {retry_count + 1}/{max_retries})...")
            time.sleep(3)
            return get_phone_number(name, zip_code, retry_count + 1, max_retries)

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

# Quit Selenium driver. Guarded: by this point the workbook is already
# saved and renamed, so a hiccup while closing the browser must not turn
# an otherwise-successful run into a reported failure (which previously
# left app5.py pointing at the old, now-renamed, file path).
try:
    driver.quit()
except Exception as e:
    print(f"⚠️ Error while closing browser (ignored, run already completed): {e}")

print("🔄 Browser closed. Process completed!")

