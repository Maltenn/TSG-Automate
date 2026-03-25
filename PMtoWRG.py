import os
import re
import glob
import csv
import datetime
import pandas as pd
import time
import math

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select

# ─── CONFIG ────────────────────────────────────────────────────────────────────
SCRIPT_DIR           = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH           = os.path.join(SCRIPT_DIR, 'Processed_orders.xlsx')
DOWNLOAD_FOLDER      = os.getenv("TSG_DOWNLOAD_DIR", os.path.join(os.path.expanduser("~"), "Downloads"))
LOGIN_URL            = "https://wranglerb2b.com/login.php/client/NQ=="
BATCH_ORDER_URL      = "https://wranglerb2b.com/batch_order.php/ecat_view"
CHECKOUT_URL         = "https://wranglerb2b.com/tp_checkout.php/ecat_checkout"
ORDER_HISTORY_URL    = "https://wranglerb2b.com/tp_order_history.php/ecat_view"

# Folder that contains per-PO CSVs produced by your PDF extraction pipeline.
# Example: C:\TSG_Automate\pdfs\297361.csv
PDFS_DIR             = os.path.join(SCRIPT_DIR, "pdfs")

# If the extracted ship-to value matches this (after normalization), we proceed
# with the normal checkout flow (select radio and done).
DEFAULT_SHIPTO_VALUE = "THE SOURCING GROUP, INC. | 4560 36TH STREET | ORLANDO, FL 32811 | FedEx Ground: 955617339,"

EMAIL    = os.getenv("WRANGLER_EMAIL")  or os.getenv("WRG_EMAIL")  or "internal3@broberry.com"
PASSWORD = os.getenv("WRANGLER_PASSWORD") or os.getenv("WRG_PASSWORD") or "Internal3Broberry!"

# Default Selenium waits (seconds)
WAIT_SHORT = 10
WAIT_LONG = 25
WAIT_XLONG = 60

# State abbreviation to full name mapping
STATE_ABBREV_MAP = {
    "AL": "Alabama", "AK": "Alaska", "AZ": "Arizona", "AR": "Arkansas",
    "CA": "California", "CO": "Colorado", "CT": "Connecticut", "DE": "Delaware",
    "DC": "District of Columbia", "FL": "Florida", "GA": "Georgia", "HI": "Hawaii",
    "ID": "Idaho", "IL": "Illinois", "IN": "Indiana", "IA": "Iowa",
    "KS": "Kansas", "KY": "Kentucky", "LA": "Louisiana", "ME": "Maine",
    "MD": "Maryland", "MA": "Massachusetts", "MI": "Michigan", "MN": "Minnesota",
    "MS": "Mississippi", "MO": "Missouri", "MT": "Montana", "NE": "Nebraska",
    "NV": "Nevada", "NH": "New Hampshire", "NJ": "New Jersey", "NM": "New Mexico",
    "NY": "New York", "NC": "North Carolina", "ND": "North Dakota", "OH": "Ohio",
    "OK": "Oklahoma", "OR": "Oregon", "PA": "Pennsylvania", "RI": "Rhode Island",
    "SC": "South Carolina", "SD": "South Dakota", "TN": "Tennessee", "TX": "Texas",
    "UT": "Utah", "VT": "Vermont", "VA": "Virginia", "WA": "Washington",
    "WV": "West Virginia", "WI": "Wisconsin", "WY": "Wyoming"
}

# ────────────────────────────────────────────────────────────────────────────────

def log(msg: str) -> None:
    """Simple logger used throughout the script."""
    print(msg, flush=True)

def debug_dump(driver, error_name="error"):
    """Save screenshot and HTML for debugging."""
    try:
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot_path = os.path.join(SCRIPT_DIR, f"debug_{error_name}_{timestamp}.png")
        html_path = os.path.join(SCRIPT_DIR, f"debug_{error_name}_{timestamp}.html")
        
        driver.save_screenshot(screenshot_path)
        log(f"[DEBUG] Screenshot saved: {screenshot_path}")
        
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(driver.page_source)
        log(f"[DEBUG] HTML saved: {html_path}")
    except Exception as e:
        log(f"[DEBUG] Failed to save debug info: {e}")

def cleanup_old_debug_files():
    """Remove old debug screenshots and HTML files from script directory."""
    try:
        import glob
        debug_files = glob.glob(os.path.join(SCRIPT_DIR, "debug_*.png")) + \
                     glob.glob(os.path.join(SCRIPT_DIR, "debug_*.html"))
        
        if debug_files:
            log(f"[INFO] Cleaning up {len(debug_files)} old debug files...")
            for file in debug_files:
                try:
                    os.remove(file)
                except Exception as e:
                    log(f"[WARN] Could not delete {file}: {e}")
            log(f"[INFO] Debug file cleanup complete")
        else:
            log("[INFO] No old debug files to clean up")
    except Exception as e:
        log(f"[WARN] Error during debug file cleanup: {e}")


def coerce_str(val) -> str:
    """Convert Excel/CSV values (floats, NaN, ints, None, etc.) to a clean string."""
    if val is None:
        return ""
    if isinstance(val, float):
        if math.isnan(val):
            return ""
        # If it's an integer-like float (e.g., 13092.0), return without .0
        if val.is_integer():
            return str(int(val))
        return str(val)
    if isinstance(val, (int,)):
        return str(val)
    return str(val).strip()


def _normalize_shipto(s: str) -> str:
    """Normalize ship-to strings for robust comparisons."""
    if s is None:
        return ""
    # Standardize whitespace + pipe formatting.
    s = str(s).replace("\r", " ").replace("\n", " ")
    s = re.sub(r"\s*\|\s*", " | ", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def _canonical_shipto(s: str) -> str:
    """A looser canonicalization used for fuzzy ship-to matching.

    Wrangler's saved ship-to labels often abbreviate (ST vs STREET), omit punctuation,
    and may include ZIP+4. The PDF-extracted 'shipTo' string can include commas/INC/etc.
    This helper reduces those differences so we can reliably detect the default
    "THE SOURCING GROUP" destination.
    """
    if s is None:
        return ""
    s = str(s).upper()
    # Replace common punctuation with spaces
    s = re.sub(r"[\.,]", " ", s)
    # Normalize street words
    s = s.replace("STREET", "ST")
    s = s.replace("AVENUE", "AVE")
    s = s.replace("ROAD", "RD")
    s = s.replace("DRIVE", "DR")
    # Remove common legal suffixes
    for tok in (" INC ", " INCORPORATED ", " LLC "):
        s = s.replace(tok, " ")
    # Collapse whitespace
    s = re.sub(r"\s+", " ", s).strip()
    return s


def is_default_sourcing_group_shipto(shipto_text: str) -> bool:
    """Return True if shipto_text refers to THE SOURCING GROUP at 4560 36TH ST, Orlando, FL 32811."""
    norm = _normalize_shipto(shipto_text)
    # Fast-path exact-ish match
    if norm == _normalize_shipto(DEFAULT_SHIPTO_VALUE):
        return True

    canon = _canonical_shipto(shipto_text)
    # Fuzzy token match: tolerate abbreviations and ZIP+4.
    required = [
        "THE SOURCING GROUP",
        "4560",
        "36TH",
        "ORLANDO",
        "FL",
        "32811",
    ]
    return all(tok in canon for tok in required)


def find_po_csv_path(po_number: str) -> str:
    """Find a PO CSV in PDFS_DIR.

    Prefers an exact '<po>.csv'. Otherwise, falls back to a wildcard
    search for '*<po>*...*.csv' and chooses the most recently modified.
    Returns an empty string if nothing is found.
    """
    exact = os.path.join(PDFS_DIR, f"{po_number}.csv")
    if os.path.exists(exact):
        return exact

    try:
        pattern = os.path.join(PDFS_DIR, f"*{po_number}*.csv")
        candidates = glob.glob(pattern)
        if not candidates:
            # Sometimes extensions can be uppercase, depending on how it was saved
            pattern2 = os.path.join(PDFS_DIR, f"*{po_number}*.CSV")
            candidates = glob.glob(pattern2)
        if not candidates:
            return ""
        candidates.sort(key=lambda fp: os.path.getmtime(fp), reverse=True)
        return candidates[0]
    except Exception:
        return ""


def load_shipto_data_from_csv(po_number: str) -> dict:
    """Load ship-to data from the PO's CSV.
    
    Returns a dict with keys:
    - 'shipTo': Full ship-to text from column C (index 2)
    - 'company': shipToCompany from column K (index 10)
    - 'attention': shipToAttention from column L (index 11)
    - 'street': shipToStreet from column M (index 12)
    - 'city': shipToCity from column N (index 13)
    - 'state': shipToState from column O (index 14)
    - 'zip': shipToZip from column P (index 15)
    """
    csv_path = find_po_csv_path(po_number)
    if not csv_path:
        log(f"[WARN] PO CSV not found for {po_number}: {os.path.join(PDFS_DIR, f'{po_number}.csv')}")
        return {}

    # Try multiple encodings in order of likelihood
    encodings = ['utf-8-sig', 'cp1252', 'latin-1', 'utf-8', 'iso-8859-1']
    
    for encoding in encodings:
        try:
            with open(csv_path, newline='', encoding=encoding) as f:
                reader = csv.reader(f)
                rows = list(reader)
                
                # Skip header row (row 0), data starts at row 1
                if len(rows) < 2:
                    log(f"[WARN] CSV file for {po_number} has insufficient data rows")
                    return {}
                
                data_row = rows[1]  # First data row
                
                # Build result dict with coerced values
                result = {
                    'shipTo': coerce_str(data_row[2]) if len(data_row) > 2 else "",
                    'company': coerce_str(data_row[10]) if len(data_row) > 10 else "",
                    'attention': coerce_str(data_row[11]) if len(data_row) > 11 else "",
                    'street': coerce_str(data_row[12]) if len(data_row) > 12 else "",
                    'city': coerce_str(data_row[13]) if len(data_row) > 13 else "",
                    'state': coerce_str(data_row[14]) if len(data_row) > 14 else "",
                    'zip': coerce_str(data_row[15]) if len(data_row) > 15 else "",
                }
                
                log(f"[INFO] Successfully read CSV with {encoding} encoding")
                return result
                
        except UnicodeDecodeError:
            # Try next encoding
            continue
        except Exception as e:
            log(f"[WARN] Failed reading PO CSV '{csv_path}' with {encoding}: {e}")
            return {}
    
    # If all encodings failed
    log(f"[ERROR] Could not read CSV '{csv_path}' with any supported encoding")
    return {}


def coerce_date(val) -> str:
    """Return a mm/dd/yyyy string for dates coming from Excel/datetime/strings."""
    if val is None:
        return ""
    # Already a string → trust it
    if isinstance(val, str):
        return val.strip()
    # datetime/date → format it
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.strftime("%m/%d/%Y")
    # Excel may pass floats (serials) or something odd; fall back to str
    return coerce_str(val)

def get_next_business_day(from_date=None):
    d = from_date or datetime.date.today()
    one_day = datetime.timedelta(days=1)
    d += one_day
    while d.weekday() > 4:  # Sat=5, Sun=6
        d += one_day
    return d

def login(driver):
    driver.get(LOGIN_URL)
    wait = WebDriverWait(driver, 20)

    # Fill credentials
    email_el = wait.until(EC.visibility_of_element_located((By.ID, "login_email")))
    email_el.clear()
    email_el.send_keys(EMAIL)

    pwd_el = driver.find_element(By.ID, "login_password")
    pwd_el.clear()
    pwd_el.send_keys(PASSWORD)

    # Pause for manual verification before submitting
    print("\n[ACTION REQUIRED]")
    print("Please complete any login verification in the browser window now.")
    print("Examples: CAPTCHA checkbox, 'I'm not a robot', or any pre-login security step.")
    input("When you're done and ready for the script to click Sign In, press Enter here... ")

    # Submit login after user confirms
    driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()

    # Continue with your original post-login wait
    wait.until(EC.presence_of_element_located((By.ID, "p7SOPt_2")))


def open_order_menu(driver):
    wait = WebDriverWait(driver, 10)
    
    # Wait for any overlays/preloaders to disappear
    wait_for_overlay_gone(driver, timeout=15)
    time.sleep(0.5)
    
    # Wait for the order menu button to be clickable
    order_menu_btn = wait.until(EC.element_to_be_clickable((By.ID, "p7SOPt_2")))
    
    # Use safe_click to handle any remaining interception issues
    safe_click(driver, order_menu_btn)
    
    # Wait for the New Draft option to appear
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.pop-newdraft")))


def to_text(x) -> str:
    """Coerce Excel/CSV values (floats, NaN, ints, None, etc.) to a clean string."""
    if x is None:
        return ""
    if isinstance(x, float):
        if math.isnan(x):
            return ""
        if x.is_integer():
            return str(int(x))
        return str(x)
    return str(x).strip()


def create_new_draft(driver, draft_name, ship_date):
    wait = WebDriverWait(driver, 10)
    
    # Wait for any overlays to disappear first
    wait_for_overlay_gone(driver, timeout=15)
    time.sleep(0.5)

    # Open the "New Draft" popup - wait for it to be clickable and use safe_click
    new_draft_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.pop-newdraft")))
    safe_click(driver, new_draft_btn)

    # Safely send the draft name and tab to the date field
    safe_name = to_text(draft_name)
    wait.until(EC.visibility_of_element_located(
        (By.ID, "pfm-newdraft"))
    ).send_keys(safe_name + "\t")

    # Pick the ship date from the calendar
    day = ship_date.day
    xpath = f"//td[@data-handler='selectDay']/a[text()='{day}']"
    wait.until(EC.element_to_be_clickable((By.XPATH, xpath))).click()

    # Click "Save New Draft"
    driver.find_element(
        By.XPATH, "//button[@onclick='save_new_draft()']"
    ).click()

    # Make closing the popup optional / resilient
    try:
        # Old behaviour: explicit close button
        close_btn = WebDriverWait(driver, 8).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//button[@onclick='preloadCloseWindow()']")
            )
        )
        close_btn.click()
    except TimeoutException:
        # If the site changed and there is no close button anymore,
        # just wait briefly for the popup/overlay to disappear and continue.
        try:
            WebDriverWait(driver, 8).until(
                EC.invisibility_of_element_located(
                    (By.ID, "fancybox-wrap")
                )
            )
        except TimeoutException:
            # As a last resort, just continue
            pass
    
    # Final wait for any overlays to clear before moving on
    wait_for_overlay_gone(driver, timeout=10)



def upload_batch_order(driver, order_no):
    wait = WebDriverWait(driver, 60)
    driver.get(BATCH_ORDER_URL)
    
    # Wait for page to load and any overlays to clear
    wait_ready(driver, timeout=25)
    wait_for_overlay_gone(driver, timeout=15)

    file_input = wait.until(EC.presence_of_element_located((By.ID, "load_items_file")))
    driver.execute_script("arguments[0].style.display = 'block';", file_input)

    pattern1 = os.path.join(DOWNLOAD_FOLDER, f"*{order_no}*wrangler*.*")
    candidates = [f for f in glob.glob(pattern1) if f.lower().endswith((".xml", ".xlsx"))]
    if not candidates:
        pattern2 = os.path.join(DOWNLOAD_FOLDER, f"*{order_no}*.*")
        candidates = [f for f in glob.glob(pattern2) if f.lower().endswith((".xml", ".xlsx"))]
    if not candidates:
        available = os.listdir(DOWNLOAD_FOLDER)
        raise FileNotFoundError(f"No file for order {order_no} in {DOWNLOAD_FOLDER}. Contains: {available}")

    xml_path = candidates[0]
    print(f"[INFO] Uploading file: {xml_path}")
    file_input.send_keys(xml_path)
    time.sleep(9)
    add_btn = wait.until(EC.element_to_be_clickable((
        By.XPATH, "//button[contains(@onclick,'add_ecat_items_to_cart_alert')]"
    )))
    add_btn.click()
    time.sleep(16)


def wait_ready(driver, timeout=25):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

def wait_for_overlay_gone(driver, timeout=20):
    """Wait for common overlays/spinners to go away."""
    wait = WebDriverWait(driver, timeout)
    try:
        wait.until(EC.invisibility_of_element_located(
            (By.CSS_SELECTOR, ".fancybox-overlay, .modal-backdrop, .blockUI, .loading-overlay, .loading, [id^='fs-preloader']")
        ))
    except TimeoutException:
        pass


def safe_click(driver, el):
    """Click element, falling back to JavaScript if intercepted or not interactable."""
    from selenium.common.exceptions import ElementClickInterceptedException, ElementNotInteractableException
    try:
        el.click()
    except (ElementClickInterceptedException, ElementNotInteractableException) as e:
        log(f"[INFO] Regular click failed ({type(e).__name__}), using JavaScript click")
        driver.execute_script("arguments[0].click();", el)

def wait_modal_open(driver, timeout=10):
    """
    Fancybox usually injects .fancybox-overlay + .fancybox-wrap/.fancybox-inner.
    Wait until the radio list is present, VISIBLE, and clickable.
    """
    wait = WebDriverWait(driver, timeout)
    
    log("[DEBUG] Waiting for modal to open...")
    
    # Check if fancybox overlay appeared
    try:
        overlay = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, ".fancybox-overlay, .fancybox-wrap")
        ))
        log("[DEBUG] Fancybox overlay detected")
    except TimeoutException:
        log("[ERROR] Fancybox overlay never appeared!")
        debug_dump(driver, "modal_overlay_timeout")
        log("[DEBUG] Current URL: " + driver.current_url)
        raise
    
    # Wait for fancybox-inner (the actual modal content container) to be visible
    try:
        inner = wait.until(EC.visibility_of_element_located(
            (By.CSS_SELECTOR, ".fancybox-inner")
        ))
        log("[DEBUG] Fancybox inner content is visible")
    except TimeoutException:
        log("[WARN] Fancybox inner content not detected, continuing anyway...")
    
    # Wait for radio buttons to be present
    try:
        radio = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "input[name='add_addresses1']")
        ))
        log("[DEBUG] Radio buttons found in DOM")
    except TimeoutException:
        log("[ERROR] Radio buttons never appeared in DOM!")
        debug_dump(driver, "radio_buttons_timeout")
        raise
    
    # CRITICAL: Wait for radio buttons to become VISIBLE (not just present)
    # The modal animates in, so buttons exist but are hidden initially
    log("[DEBUG] Waiting for radio buttons to become visible...")
    
    # Try scrolling within the modal in case there's an inner scroll container
    try:
        scroll_container = driver.find_element(By.CSS_SELECTOR, ".stylescrollA, .fancybox-inner")
        driver.execute_script("arguments[0].scrollTop = 0;", scroll_container)
        log("[DEBUG] Scrolled modal content to top")
    except:
        pass
    
    max_attempts = 15  # Try for up to 3 seconds (15 * 0.2s)
    for attempt in range(max_attempts):
        try:
            radio = driver.find_element(By.CSS_SELECTOR, "input[name='add_addresses1']")
            
            # Check if it's actually visible with real dimensions
            is_displayed = radio.is_displayed()
            size = radio.size
            location = radio.location
            
            log(f"[DEBUG] Attempt {attempt+1}: displayed={is_displayed}, size={size}, location={location}")
            
            if is_displayed and size['height'] > 0 and size['width'] > 0:
                log("[DEBUG] Radio button is now visible with real dimensions!")
                break
                
            time.sleep(0.2)
            
        except Exception as e:
            log(f"[DEBUG] Attempt {attempt+1} check failed: {e}")
            time.sleep(0.2)
    else:
        # If we exhausted all attempts
        log("[ERROR] Radio buttons never became visible after 3 seconds!")
        debug_dump(driver, "radio_not_visible")
        raise TimeoutException("Radio buttons present but never became visible")
    
    # Now wait for it to be clickable
    log("[DEBUG] Waiting for radio button to be clickable...")
    try:
        clickable_radio = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "input[name='add_addresses1']")
        ))
        log("[DEBUG] Radio button is now clickable")
        return clickable_radio
    except TimeoutException:
        log("[ERROR] Radio button never became clickable!")
        debug_dump(driver, "radio_not_clickable")
        raise

def wait_modal_close(driver, timeout=10):
    """Wait for modal and overlay to disappear."""
    wait = WebDriverWait(driver, timeout)
    wait.until(EC.invisibility_of_element_located(
        (By.CSS_SELECTOR, ".fancybox-overlay")
    ))


def open_and_choose_ship_to(
    driver,
    preferred_radio_id: str = None,
    preferred_value_contains: str = None,
    preferred_label_contains: str = None,
    preferred_account_number: str = None,
    max_retries: int = 5,
):
    """
    Open the Ship-To modal and select the specified ship-to address.
    Includes retry logic with page refresh for stale page state.
    """
    wait = WebDriverWait(driver, 25)
    
    for attempt in range(max_retries):
        try:
            log(f"[DEBUG] Ship-To modal attempt {attempt + 1}/{max_retries}")
            
            # Check for any alerts or error messages first
            if attempt == 0:  # Only check on first attempt
                log("[DEBUG] Checking for alerts or error messages...")
                try:
                    alerts = driver.find_elements(By.CSS_SELECTOR, ".alert, .error, .warning, [role='alert']")
                    if alerts:
                        visible_alerts = [a for a in alerts if a.is_displayed()]
                        if visible_alerts:
                            log(f"[WARN] Found {len(visible_alerts)} visible alerts on page!")
                            for i, alert in enumerate(visible_alerts[:3]):
                                log(f"[WARN] Alert {i+1}: {alert.text[:100]}")
                except Exception as e:
                    log(f"[DEBUG] Error checking alerts: {e}")
            
            # Re-find the button on each attempt (avoid stale element)
            log("[DEBUG] Looking for Ship-To button...")
            try:
                shiptos_btn = wait.until(EC.element_to_be_clickable((
                    By.CSS_SELECTOR, "button.pop-myShipTos-1, button[class*='pop-myShipTos']"
                )))
                log(f"[DEBUG] Found Ship-To button with selector: button.pop-myShipTos-1")
            except TimeoutException:
                log("[DEBUG] First selector failed, trying XPath...")
                shiptos_btn = wait.until(EC.element_to_be_clickable((
                    By.XPATH, "//button[contains(., \"Available Ship-To\") or contains(., \"Ship-To\")]"
                )))
                log(f"[DEBUG] Found Ship-To button with XPath")
            
            log(f"[DEBUG] Ship-To button text: '{shiptos_btn.text}'")
            log(f"[DEBUG] Ship-To button is_displayed: {shiptos_btn.is_displayed()}, is_enabled: {shiptos_btn.is_enabled()}")
            
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", shiptos_btn)
            log("[DEBUG] Scrolled to Ship-To button")
            
            safe_click(driver, shiptos_btn)
            log("[DEBUG] Clicked Ship-To button")
            
            time.sleep(1.5)  # Longer wait for modal to appear
            log("[DEBUG] Waited 1.5s for modal animation to start")
            
            # Verify modal actually opened
            log("[DEBUG] Checking if modal opened after button click...")
            try:
                WebDriverWait(driver, 3).until(EC.presence_of_element_located(
                    (By.CSS_SELECTOR, ".fancybox-overlay, .fancybox-wrap")
                ))
                log("[DEBUG] Modal overlay detected after click")
                
                # If we got here, modal opened successfully - break out of retry loop
                break
                
            except TimeoutException:
                log(f"[WARN] Modal did not open on attempt {attempt + 1}")
                debug_dump(driver, f"modal_not_opened_attempt_{attempt + 1}")
                
                if attempt < max_retries - 1:
                    # Try refreshing the page for next attempt
                    log("[INFO] Refreshing page and retrying...")
                    driver.refresh()
                    time.sleep(4)  # Longer wait after refresh to let page fully load
                    
                    # Wait for page to be ready
                    WebDriverWait(driver, 15).until(
                        lambda d: d.execute_script("return document.readyState") == "complete"
                    )
                    time.sleep(1)
                else:
                    # Last attempt failed
                    log("[ERROR] All attempts to open modal failed!")
                    raise TimeoutException("Ship-To modal never opened after multiple attempts")
        
        except TimeoutException:
            if attempt == max_retries - 1:
                raise
            log(f"[WARN] Attempt {attempt + 1} failed with timeout, will retry...")
            continue
    
    # Wait for modal to fully open
    wait_modal_open(driver, timeout=12)
    
    # Find and select the radio button
    all_radios = driver.find_elements(By.CSS_SELECTOR, "input[name='add_addresses1']")
    chosen_radio = None

    for radio in all_radios:
        radio_id = radio.get_attribute("id")
        radio_val = radio.get_attribute("value") or ""
        
        # Get the label text
        try:
            label = driver.find_element(By.XPATH, f"//label[input[@id='{radio_id}']]")
            label_text = label.text
        except NoSuchElementException:
            label_text = ""

        # Match criteria
        if preferred_radio_id and radio_id == preferred_radio_id:
            chosen_radio = radio
            log(f"[INFO] Matched ship-to by radio ID: {radio_id}")
            break
        if preferred_value_contains and preferred_value_contains in radio_val:
            chosen_radio = radio
            log(f"[INFO] Matched ship-to by value substring: '{preferred_value_contains}'")
            break
        if preferred_label_contains and preferred_label_contains in label_text:
            chosen_radio = radio
            log(f"[INFO] Matched ship-to by label substring: '{preferred_label_contains}'")
            break
        if preferred_account_number and f"account_number={preferred_account_number}" in radio_val:
            chosen_radio = radio
            log(f"[INFO] Matched ship-to by account number: {preferred_account_number}")
            break

    if chosen_radio is None:
        # Default to first radio if no match
        chosen_radio = all_radios[0]
        log("[WARN] No ship-to matched the criteria; selecting first available.")

    # Select the radio button
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", chosen_radio)
    
    if not chosen_radio.is_selected():
        # Wait for the specific radio to be clickable
        try:
            wait.until(EC.element_to_be_clickable(chosen_radio))
            log("[DEBUG] Radio button is clickable, attempting click...")
        except TimeoutException:
            log("[WARN] Radio button wait timed out, will try clicking anyway...")
            time.sleep(0.5)
        
        # Try to click it
        try:
            safe_click(driver, chosen_radio)
            log("[DEBUG] Radio button clicked successfully")
        except Exception as e:
            log(f"[WARN] Regular click failed: {e}, using JavaScript to select radio...")
            # Fallback: Use JavaScript to directly set the radio as checked
            driver.execute_script("""
                arguments[0].checked = true;
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('click', { bubbles: true }));
            """, chosen_radio)
            log("[DEBUG] Radio button selected via JavaScript")


def fill_drop_ship_form(driver, shipto_data: dict):
    """
    Fill in the drop ship form with data from the CSV.
    
    Args:
        driver: Selenium WebDriver instance
        shipto_data: Dict containing address data with keys:
                     company, attention, street, city, state, zip
    """
    wait = WebDriverWait(driver, 15)
    
    log("[INFO] Filling drop ship form...")
    
    # 1. Set Country to United States
    try:
        country_select = wait.until(EC.element_to_be_clickable(
            (By.ID, "fm-shipTo-country")
        ))
        Select(country_select).select_by_value("USA")
        log("[INFO] Set country to United States")
        time.sleep(1)  # Wait for state dropdown to populate
    except Exception as e:
        log(f"[WARN] Failed to set country: {e}")
    
    # 2. Fill Contact Name (Column K - shipToCompany)
    try:
        contact_name = shipto_data.get('company', '')
        if contact_name:
            contact_input = wait.until(EC.presence_of_element_located(
                (By.ID, "fm-addrbook-contactName")
            ))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", contact_input)
            contact_input.clear()
            contact_input.send_keys(contact_name)
            log(f"[INFO] Set contact name: {contact_name}")
    except Exception as e:
        log(f"[WARN] Failed to set contact name: {e}")
    
    # 3. Fill Address 1 (Column M - shipToStreet)
    try:
        street = shipto_data.get('street', '')
        if street:
            addr1_input = wait.until(EC.presence_of_element_located(
                (By.ID, "fm-shipTo-addr-1")
            ))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", addr1_input)
            addr1_input.clear()
            addr1_input.send_keys(street)
            log(f"[INFO] Set address 1: {street}")
    except Exception as e:
        log(f"[WARN] Failed to set address 1: {e}")
    
    # 4. Fill Address 2 (Column L - shipToAttention) - max 50 chars, can be blank
    try:
        attention = shipto_data.get('attention', '')
        if attention:
            # Truncate to 50 characters if needed
            attention = attention[:50]
            addr2_input = wait.until(EC.presence_of_element_located(
                (By.ID, "fm-shipTo-addr-2")
            ))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", addr2_input)
            addr2_input.clear()
            addr2_input.send_keys(attention)
            log(f"[INFO] Set address 2: {attention}")
    except Exception as e:
        log(f"[WARN] Failed to set address 2: {e}")
    
    # 5. Fill City (Column N - shipToCity)
    try:
        city = shipto_data.get('city', '')
        if city:
            city_input = wait.until(EC.presence_of_element_located(
                (By.ID, "fm-shipTo-city")
            ))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", city_input)
            city_input.clear()
            city_input.send_keys(city)
            log(f"[INFO] Set city: {city}")
    except Exception as e:
        log(f"[WARN] Failed to set city: {e}")
    
    # 6. Select State (Column O - shipToState)
    try:
        state_abbrev = shipto_data.get('state', '').upper()
        if state_abbrev:
            state_select = wait.until(EC.element_to_be_clickable(
                (By.ID, "fm-shipTo-state")
            ))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", state_select)
            # Select by abbreviation value
            Select(state_select).select_by_value(state_abbrev)
            log(f"[INFO] Set state: {state_abbrev}")
    except Exception as e:
        log(f"[WARN] Failed to set state: {e}")
    
    # 7. Fill Zip Code (Column P - shipToZip)
    try:
        zipcode = shipto_data.get('zip', '')
        if zipcode:
            # Ensure zip is max 7 characters (some forms have maxlength=7)
            zipcode = str(zipcode)[:7]
            zip_input = wait.until(EC.presence_of_element_located(
                (By.ID, "fm-shipTo-zipcode")
            ))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", zip_input)
            zip_input.clear()
            zip_input.send_keys(zipcode)
            log(f"[INFO] Set zip code: {zipcode}")
    except Exception as e:
        log(f"[WARN] Failed to set zip code: {e}")
    
    # 8. Fill Email field with required addresses
    try:
        email_text = "sales@broberry.com"
        email_input = wait.until(EC.presence_of_element_located(
            (By.ID, "email")
        ))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", email_input)
        email_input.clear()
        email_input.send_keys(email_text)
        log(f"[INFO] Set email: {email_text}")
    except Exception as e:
        log(f"[WARN] Failed to set email: {e}")
    
    # 9. Fill Special Instructions with FedEx Ground number
    try:
        instructions_text = "FedEx Ground 955617339"
        instructions_input = wait.until(EC.presence_of_element_located(
            (By.ID, "fm-shipTo-instructions")
        ))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", instructions_input)
        instructions_input.clear()
        instructions_input.send_keys(instructions_text)
        log(f"[INFO] Set special instructions: {instructions_text}")
    except Exception as e:
        log(f"[WARN] Failed to set special instructions: {e}")
    
    log("[INFO] Drop ship form filled successfully")


def handle_address_verification_popup(driver, timeout=10):
    """
    Handle the 'Verify Your Address' popup that may appear after order submission.
    
    This popup shows ORIGINAL and SUGGESTED addresses from USPS verification.
    We select the SUGGESTED radio button and click Continue.
    
    Returns True if popup was handled, False if popup didn't appear.
    """
    try:
        # Wait for the address verification popup to appear
        popup = WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located(
                (By.ID, "pop-chk-address-verify-1")
            )
        )
        log("[INFO] Address verification popup detected!")
        
        # Check which section is visible (verify or invalid)
        try:
            verify_section = driver.find_element(By.ID, "address-chk-verify-2")
            verify_visible = driver.execute_script(
                "return window.getComputedStyle(arguments[0]).display !== 'none';", 
                verify_section
            )
            
            if verify_visible:
                log("[INFO] USPS suggested address corrections detected.")
                
                # Select the SUGGESTED radio button
                try:
                    suggested_radio = driver.find_element(By.ID, "fm-choutNumbShipTo-suggest-s1")
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", suggested_radio)
                    
                    if not suggested_radio.is_selected():
                        safe_click(driver, suggested_radio)
                        log("[INFO] Selected SUGGESTED address")
                        time.sleep(0.3)
                except NoSuchElementException:
                    log("[WARN] Could not find SUGGESTED radio button, trying ORIGINAL")
                    # Fallback to ORIGINAL if SUGGESTED not found
                    try:
                        original_radio = driver.find_element(By.ID, "fm-choutNumbShipTo-orig-s1")
                        if not original_radio.is_selected():
                            safe_click(driver, original_radio)
                            log("[INFO] Selected ORIGINAL address")
                            time.sleep(0.3)
                    except NoSuchElementException:
                        log("[WARN] Could not find ORIGINAL radio button either")
                
                # Click Continue button
                try:
                    continue_btn = driver.find_element(By.ID, "continue_chk_address")
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", continue_btn)
                    safe_click(driver, continue_btn)
                    log("[INFO] Clicked Continue on address verification popup")
                    time.sleep(1)
                    return True
                except NoSuchElementException:
                    log("[WARN] Could not find Continue button")
                    
        except NoSuchElementException:
            pass
        
        # Check if it's the "Invalid Address" popup instead
        try:
            invalid_section = driver.find_element(By.ID, "address-chk-invalid-2")
            invalid_visible = driver.execute_script(
                "return window.getComputedStyle(arguments[0]).display !== 'none';", 
                invalid_section
            )
            
            if invalid_visible:
                log("[WARN] Invalid address popup detected!")
                log("[INFO] Clicking 'Use as Entered' button...")
                
                try:
                    use_as_entered_btn = driver.find_element(By.ID, "use_as_entered")
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", use_as_entered_btn)
                    safe_click(driver, use_as_entered_btn)
                    log("[INFO] Clicked 'Use as Entered' button")
                    time.sleep(1)
                    return True
                except NoSuchElementException:
                    log("[WARN] Could not find 'Use as Entered' button")
        except NoSuchElementException:
            pass
            
    except TimeoutException:
        # Popup didn't appear - this is normal for many orders
        return False
    except Exception as e:
        log(f"[WARN] Error handling address verification popup: {e}")
        return False
    
    return False


def submit_checkout(driver, timeout=25):
    """Submit the order checkout form."""
    wait = WebDriverWait(driver, timeout)
    
    # 1) Find and click the submit button
    # Try multiple selectors since the button might have different attributes
    submit_btn = None
    
    try:
        # First try: look for the button by ID (most reliable)
        submit_btn = wait.until(EC.element_to_be_clickable((By.ID, "submit_order")))
        log("[INFO] Found submit button by ID")
    except TimeoutException:
        try:
            # Second try: look for validate_checkout_form onclick
            submit_btn = wait.until(EC.element_to_be_clickable((
                By.XPATH, "//button[contains(@onclick,'validate_checkout_form')]"
            )))
            log("[INFO] Found submit button by validate_checkout_form onclick")
        except TimeoutException:
            try:
                # Third try: look for ecat_submit_order onclick (older version)
                submit_btn = wait.until(EC.element_to_be_clickable((
                    By.XPATH, "//button[contains(@onclick,'ecat_submit_order')]"
                )))
                log("[INFO] Found submit button by ecat_submit_order onclick")
            except TimeoutException:
                # Fourth try: look for Submit Order button by text
                submit_btn = wait.until(EC.element_to_be_clickable((
                    By.XPATH, "//button[contains(text(),'Submit Order')]"
                )))
                log("[INFO] Found submit button by text content")
    
    if submit_btn is None:
        raise RuntimeError("Could not find Submit Order button")
    
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", submit_btn)
    safe_click(driver, submit_btn)
    log("[INFO] Clicked Submit Order button")

    # 2) Handle the "Please review your order" confirmation alert
    try:
        WebDriverWait(driver, 5).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        alert.accept()
        log("[INFO] Accepted confirmation alert")
    except TimeoutException:
        pass

    # 3) Handle "DON'T MISS OUT!" free-shipping modal if it shows up
    try:
        popup = WebDriverWait(driver, 4).until(
            EC.visibility_of_element_located(
                (By.ID, "not-all-qualified-pop-alert-1")
            )
        )
        log("[INFO] Free shipping popup detected")
        try:
            proceed_btn = popup.find_element(
                By.XPATH, ".//button[contains(@onclick, 'ecat_submit_order')]"
            )
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", proceed_btn
            )
            safe_click(driver, proceed_btn)
            log("[INFO] Clicked Proceed on free shipping popup")
        except NoSuchElementException:
            pass
    except TimeoutException:
        pass

    # 4) Give the site a beat to process
    time.sleep(0.6)
    wait_for_overlay_gone(driver, timeout=timeout)

    # 5) Handle address verification popup if it appears
    handle_address_verification_popup(driver, timeout=8)

    # 6) Check for error banner
    try:
        err = driver.find_element(By.ID, "submit_order_error_text")
        visible = driver.execute_script(
            "return window.getComputedStyle(arguments[0]).display !== 'none';", err
        )
        if visible:
            raise RuntimeError(
                "Order submission appears to be disabled (error banner shown)."
            )
    except NoSuchElementException:
        pass


def checkout_and_ship(driver, po_number: str, client_po: str):
    """
    Navigate to checkout and handle ship-to selection based on address:
    - If default Sourcing Group address: select radio and proceed normally
    - If non-default address: select Sourcing Group radio, click Select, 
      click Drop Ship, and fill the drop ship form
    """
    # Load the extracted ship-to data for this client PO
    shipto_data = load_shipto_data_from_csv(client_po)
    
    if not shipto_data or not shipto_data.get('shipTo'):
        log(f"[WARN] No ship-to data found for {client_po}; falling back to default behavior.")
        is_default_shipto = True
    else:
        is_default_shipto = is_default_sourcing_group_shipto(shipto_data['shipTo'])
    
    # 1) Navigate to checkout
    driver.get(CHECKOUT_URL)
    wait_ready(driver, timeout=25)
    
    # Handle the fs-preloader overlay that checks inventory (takes 3-15 seconds)
    try:
        wait = WebDriverWait(driver, 30)
        
        # First, wait for the preloader div to appear
        log("[INFO] Waiting for inventory check preloader...")
        preloader = wait.until(EC.presence_of_element_located((By.ID, "fs-preloader-1")))
        
        # Wait for the "Continue" section to become visible (display: block)
        # This happens after the inventory check completes
        log("[INFO] Waiting for Continue button to appear...")
        continue_section = wait.until(EC.visibility_of_element_located((By.ID, "fs-preload-continue")))
        
        # Now wait for the Continue button itself to be present and visible
        continue_btn = wait.until(EC.visibility_of_element_located((
            By.XPATH, "//div[@id='fs-preload-continue']//button[@onclick='preloadCloseWindow()']"
        )))
        
        # Try multiple methods to click the button
        log("[INFO] Clicking Continue button on checkout overlay...")
        try:
            # Method 1: Try regular click first
            continue_btn.click()
            log("[INFO] Continue button clicked (regular click)")
        except Exception as e1:
            log(f"[INFO] Regular click failed: {e1}, trying JavaScript click...")
            try:
                # Method 2: Try JavaScript click
                driver.execute_script("arguments[0].click();", continue_btn)
                log("[INFO] Continue button clicked (JavaScript click)")
            except Exception as e2:
                log(f"[INFO] JavaScript click failed: {e2}, trying direct function call...")
                try:
                    # Method 3: Directly execute the onclick function
                    driver.execute_script("preloadCloseWindow();")
                    log("[INFO] Continue button clicked (direct function call)")
                except Exception as e3:
                    log(f"[WARN] All click methods failed: {e3}")
        
        # Wait for the overlay to disappear
        time.sleep(1)
        wait_for_overlay_gone(driver, timeout=10)
        log("[INFO] Checkout overlay dismissed successfully")
        
    except TimeoutException:
        log("[INFO] No Continue overlay found (may have already been dismissed)")
    
    wait = WebDriverWait(driver, 25)
    
    if is_default_shipto:
        # OLD WAY: Just select The Sourcing Group radio button and click Select
        log("[INFO] Default ship-to address detected; using old method (select radio only).")
        open_and_choose_ship_to(
            driver,
            preferred_radio_id="add_addresses-4",
            preferred_value_contains="store=THE SOURCING GROUP",
            preferred_label_contains="THE SOURCING GROUP",
            preferred_account_number="1000263820",
        )
        
        # Click Select button
        try:
            select_btn = wait.until(EC.element_to_be_clickable((
                By.XPATH, "//button[@onclick='return selected_my_shiptos()']"
            )))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", select_btn)
            safe_click(driver, select_btn)
            log("[INFO] Clicked Select button")
        except TimeoutException:
            log("[WARN] Could not find Select button")
        
        # Wait for modal to close
        try:
            wait_modal_close(driver, timeout=12)
        except TimeoutException:
            pass
        
    else:
        # NEW WAY: Select The Sourcing Group, click Select, then Drop Ship, then fill form
        log("[INFO] Non-default ship-to address detected; using new method (Drop Ship form).")
        
        # Step 1: Select The Sourcing Group radio button
        open_and_choose_ship_to(
            driver,
            preferred_radio_id="add_addresses-4",
            preferred_value_contains="store=THE SOURCING GROUP",
            preferred_label_contains="THE SOURCING GROUP",
            preferred_account_number="1000263820",
        )
        
        # Step 2: Click Select button
        try:
            select_btn = wait.until(EC.element_to_be_clickable((
                By.XPATH, "//button[@onclick='return selected_my_shiptos()']"
            )))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", select_btn)
            safe_click(driver, select_btn)
            log("[INFO] Clicked Select button")
        except TimeoutException:
            log("[WARN] Could not find Select button")
        
        # Wait for modal to close
        try:
            wait_modal_close(driver, timeout=12)
        except TimeoutException:
            pass
        
        wait_for_overlay_gone(driver, timeout=20)
        time.sleep(1)
        
        # Step 3: Click Drop Ship button
        try:
            drop_ship_btn = wait.until(EC.element_to_be_clickable((
                By.XPATH, "//button[@onclick='BTNaddNewShipToAddress()']"
            )))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", drop_ship_btn)
            safe_click(driver, drop_ship_btn)
            log("[INFO] Clicked Drop Ship button")
            time.sleep(1)
        except TimeoutException:
            log("[WARN] Could not find Drop Ship button")
        
        # Step 4: Fill in the drop ship form
        fill_drop_ship_form(driver, shipto_data)
    
    wait_for_overlay_gone(driver, timeout=20)
    time.sleep(0.3)
    
    # Fill PO number (same for both paths)
    try:
        po_input = WebDriverWait(driver, 8).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "input#fm-shipTo-po-Order1"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", po_input)
        po_input.clear()
        po_input.send_keys(po_number)
    except TimeoutException:
        inputs = wait.until(EC.presence_of_all_elements_located(
            (By.CSS_SELECTOR, "input[name='po_order_number[]']")
        ))
        for inp in inputs:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", inp)
            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable(inp))
                inp.clear()
                inp.send_keys(po_number)
            except TimeoutException:
                driver.execute_script("arguments[0].value = arguments[1];", inp, po_number)
    
    # Auto-submit the order
    log(f"[INFO] Submitting order for PO '{po_number}' (Client PO {client_po})...")
    submit_checkout(driver, timeout=25)
    log(f"[OK] Order submitted successfully!")
    time.sleep(0.5)



def main():
    """Main automation script that places orders with Wrangler."""
    log("")
    log("="*60)
    log("*** WRANGLER B2B ORDER AUTOMATION SCRIPT ***")
    log("="*60)
    log(f"Script started at: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log(f"Excel file: {EXCEL_PATH}")
    log(f"PDF directory: {PDFS_DIR}")
    log("="*60)
    
    # Clean up old debug files from previous runs
    cleanup_old_debug_files()
    log("")
    
    driver = webdriver.Chrome()
    try:
        log("[INFO] Logging in to Wrangler B2B...")
        login(driver)
        log("[OK] Login successful!")
        
        log("[INFO] Loading Excel file...")
        df = pd.read_excel(EXCEL_PATH, engine="openpyxl", dtype=str)
        col_g = df.columns[6]   # draft-name / PO column
        col_d = df.columns[3]   # Client PO # (used to locate PDFs_DIR\\<Client PO #>.csv)
        col_j = df.columns[9]   # the raw order field

        # track which row/index and PO we processed
        processed = []
        
        total_orders = len(df)
        log("")
        log("="*60)
        log(f"[INFO] Starting Order Placement - {total_orders} orders to process")
        log("="*60)

        # Place all orders
        for order_num, (idx, row) in enumerate(df.iterrows(), 1):
            draft_name  = coerce_str(row[col_g])
            client_po   = coerce_str(row[col_d])
            order_field = row[col_j]
            m = re.search(r"\d+", order_field)
            if not m:
                raise ValueError(f"Cannot parse order number from '{order_field}'")
            order_no  = m.group()
            ship_date = get_next_business_day()
            
            log("")
            log(f"[{order_num}/{total_orders}] Processing Order:")
            log(f"  - PO Number: {draft_name}")
            log(f"  - Client PO: {client_po}")
            log(f"  - Order File: {order_no}")

            open_order_menu(driver)
            create_new_draft(driver, draft_name, ship_date)
            upload_batch_order(driver, order_no)
            checkout_and_ship(driver, draft_name, client_po)

            processed.append((idx, draft_name))
            log(f"[OK] Order {draft_name} placed successfully!")
            
            # Small pause between orders to let system settle
            time.sleep(2)

        # Final summary
        log("")
        log("="*60)
        log("*** ORDER PLACEMENT COMPLETE ***")
        log("="*60)
        log(f"Total Orders Placed: {len(processed)}")
        log(f"Excel File: {EXCEL_PATH}")
        log("")
        log("[INFO] Use 'Get Order IDs' button to fetch Order IDs after placement")
        log("="*60)

    except Exception as e:
        log("")
        log("="*60)
        log("[ERROR] Script encountered an error!")
        log(f"[ERROR] {str(e)}")
        log("="*60)
        raise
    finally:
        log("[INFO] Closing browser...")
        driver.quit()
        log("[INFO] Browser closed.")

if __name__ == "__main__":
    main()
