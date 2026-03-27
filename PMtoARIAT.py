import os
import re
import glob
import time
import math
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException


# ─── CONFIG ────────────────────────────────────────────────────────────────────
SCRIPT_DIR      = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH      = os.path.join(SCRIPT_DIR, "Processed_orders.xlsx")

DOWNLOAD_FOLDER = os.getenv("TSG_DOWNLOAD_DIR", os.path.join(os.path.expanduser("~"), "Downloads"))
PDF_DIR = os.getenv("TSG_PDF_DIR", os.path.join(SCRIPT_DIR, "pdfs"))

ARIAT_URL       = "https://b2b.ariat.com/"

ARIAT_USERNAME  = os.getenv("ARIAT_USERNAME", "internal3")
ARIAT_PASSWORD  = os.getenv("ARIAT_PASSWORD", "5Wft87ptvX68h3h")

WAIT_LONG  = 90
WAIT_MED   = 30
WAIT_SHORT = 10
# ────────────────────────────────────────────────────────────────────────────────


US_STATE_ABBR_TO_NAME = {
    "AL":"Alabama","AK":"Alaska","AZ":"Arizona","AR":"Arkansas","CA":"California","CO":"Colorado","CT":"Connecticut",
    "DE":"Delaware","FL":"Florida","GA":"Georgia","HI":"Hawaii","ID":"Idaho","IL":"Illinois","IN":"Indiana",
    "IA":"Iowa","KS":"Kansas","KY":"Kentucky","LA":"Louisiana","ME":"Maine","MD":"Maryland","MA":"Massachusetts",
    "MI":"Michigan","MN":"Minnesota","MS":"Mississippi","MO":"Missouri","MT":"Montana","NE":"Nebraska","NV":"Nevada",
    "NH":"New Hampshire","NJ":"New Jersey","NM":"New Mexico","NY":"New York","NC":"North Carolina","ND":"North Dakota",
    "OH":"Ohio","OK":"Oklahoma","OR":"Oregon","PA":"Pennsylvania","RI":"Rhode Island","SC":"South Carolina",
    "SD":"South Dakota","TN":"Tennessee","TX":"Texas","UT":"Utah","VT":"Vermont","VA":"Virginia","WA":"Washington",
    "WV":"West Virginia","WI":"Wisconsin","WY":"Wyoming","DC":"District of Columbia",
}


def coerce_str(val) -> str:
    if val is None:
        return ""
    if isinstance(val, float):
        if math.isnan(val):
            return ""
        if val.is_integer():
            return str(int(val))
        return str(val)
    return str(val).strip()


def extract_po_key(po_number: str) -> str:
    """
    For values like '162945-297239' return '297239' (last digit chunk).
    If the string is already '297239', returns '297239'.
    """
    s = coerce_str(po_number)
    chunks = re.findall(r"\d+", s)
    return chunks[-1] if chunks else s


def wait_ready(driver, timeout=WAIT_MED):
    WebDriverWait(driver, timeout).until(lambda d: d.execute_script("return document.readyState") == "complete")


def safe_click(driver, el):
    try:
        el.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].click();", el)


def wait_and_click(driver, by, sel, timeout=WAIT_MED):
    el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, sel)))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    safe_click(driver, el)
    return el


def wait_visible(driver, by, sel, timeout=WAIT_MED):
    return WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((by, sel)))


def wait_present(driver, by, sel, timeout=WAIT_MED):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, sel)))


def click_button_by_text(driver, text, timeout=WAIT_MED):
    xpath = f"//button[normalize-space()='{text}' or .//div[normalize-space()='{text}'] or contains(normalize-space(.), '{text}')]"
    return wait_and_click(driver, By.XPATH, xpath, timeout=timeout)


# ─── DIJIT BUTTON HELPERS (for Save) ───────────────────────────────────────────
def click_dijit_button_by_label(driver, label_text: str, timeout=WAIT_LONG, prefer_id: str | None = None):
    """
    Clicks a Dojo/Dijit button reliably.
    - If prefer_id is provided and exists (e.g. 'dijit_form_Button_40'), click that first.
    - Otherwise find the dijitButtonText span by label and click its clickable ancestor.
    """
    label_text = str(label_text).strip()

    # 1) Prefer stable ID when we have one
    if prefer_id:
        try:
            el = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, prefer_id)))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            safe_click(driver, el)
            return el
        except Exception:
            pass

        # Also try label id pattern (e.g. dijit_form_Button_40_label)
        try:
            el = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, f"{prefer_id}_label")))
            # climb to clickable button area
            btn = el.find_element(By.XPATH, "./ancestor::*[@role='button'][1] | ./ancestor::*[contains(@class,'dijitButtonNode')][1]")
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            safe_click(driver, btn)
            return btn
        except Exception:
            pass

    # 2) Find by visible label and click the correct ancestor
    xpath_label = f"//span[contains(@class,'dijitButtonText') and normalize-space()='{label_text}']"
    label_span = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, xpath_label)))

    # Prefer role=button node or dijitButtonNode container
    try:
        btn = label_span.find_element(By.XPATH, "./ancestor::*[@role='button'][1]")
    except Exception:
        btn = label_span.find_element(By.XPATH, "./ancestor::*[contains(@class,'dijitButtonNode')][1]")

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
    safe_click(driver, btn)
    return btn


# ─── DOJO MAIN MENU (CRITICAL) ──────────────────────────────────────────────
MAIN_MENU_TRIGGER_ID = "dijit__WidgetsInTemplateMixin_1"
MAIN_MENU_POPUP_ID   = "dijit__WidgetsInTemplateMixin_1_dropdown"
# NOTE: Do NOT use a hardcoded dijit_MenuItem_NN_text ID here — the numeric
# suffix shifts whenever Dojo re-renders, which caused the wrong item (Export
# XLSX) to be clicked.  We now locate the row by its stable import_csv CSS
# class instead (see click_import_a_file below).
# The ID below is intentionally unused — selector-based lookup is used instead.
IMPORT_MENU_LABEL_TD_ID = None  # deprecated; kept for reference only


def is_main_menu_open(driver) -> bool:
    try:
        popup = driver.find_element(By.ID, MAIN_MENU_POPUP_ID)
        style = (popup.get_attribute("style") or "").lower()
        return ("visibility: visible" in style) and popup.is_displayed()
    except Exception:
        return False


def open_main_menu(driver, timeout=WAIT_LONG):
    end = time.time() + timeout
    last_err = None

    while time.time() < end:
        if is_main_menu_open(driver):
            return

        try:
            trigger = WebDriverWait(driver, 6).until(
                EC.element_to_be_clickable((By.ID, MAIN_MENU_TRIGGER_ID))
            )
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", trigger)
            safe_click(driver, trigger)
        except Exception as e:
            last_err = e
            try:
                trigger2 = WebDriverWait(driver, 4).until(
                    EC.element_to_be_clickable((
                        By.CSS_SELECTOR,
                        f"*[aria-owns='{MAIN_MENU_POPUP_ID}'], *[aria-controls='{MAIN_MENU_POPUP_ID}']"
                    ))
                )
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", trigger2)
                safe_click(driver, trigger2)
            except Exception as e2:
                last_err = e2

        try:
            WebDriverWait(driver, 6).until(lambda d: is_main_menu_open(d))
            return
        except Exception:
            time.sleep(0.5)

    raise TimeoutException(f"Timed out opening main menu popup ({MAIN_MENU_POPUP_ID}). Last error: {last_err}")


def wait_for_import_menu_item(driver, timeout=WAIT_LONG):
    """Wait until the 'Import a File' menu item is present in the DOM (menu is open)."""
    open_main_menu(driver, timeout=timeout)
    # Use the stable import_csv class; fall back to text-content match
    for by, sel in [
        (By.CSS_SELECTOR, "tr.import_csv"),
        (By.XPATH, "//td[normalize-space()='Import a File']/ancestor::tr[1]"),
    ]:
        try:
            return WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((by, sel))
            )
        except TimeoutException:
            continue
    raise TimeoutException("Could not confirm 'Import a File' menu item is present.")


def click_import_a_file(driver, timeout=WAIT_LONG):
    open_main_menu(driver, timeout=timeout)

    # Give the Dojo menu animation a moment to finish rendering items.
    time.sleep(0.6)

    # Selector priority:
    #  1. tr.import_csv  — most stable; purpose-built class, independent of dijit ID.
    #  2. XPath on the <tr> class attribute directly.
    #  3. Locate the label <td> by visible text and climb to the <tr>.
    #  4. Click the label <td> directly (works when <tr> intercepts differently).
    #  5. Broad aria-label contains match as last resort.
    _IMPORT_SELECTORS = [
        (By.CSS_SELECTOR, "tr.import_csv"),
        (By.XPATH,        "//tr[contains(@class,'import_csv')]"),
        (By.XPATH,        "//td[normalize-space()='Import a File']/ancestor::tr[1]"),
        (By.XPATH,        "//td[normalize-space()='Import a File']"),
        (By.XPATH,        "//*[contains(normalize-space(@aria-label),'Import a File')]"),
    ]

    el = None
    for by, sel in _IMPORT_SELECTORS:
        try:
            # Use presence_of_element_located — Dojo <tr> rows are often not
            # considered "clickable" by Selenium even when fully interactive.
            el = WebDriverWait(driver, 12).until(
                EC.presence_of_element_located((by, sel))
            )
            print(f"[INFO] Located 'Import a File' with selector: {sel}")
            break
        except TimeoutException:
            print(f"[WARN] Selector did not match, trying next: {sel}")
            continue

    if el is None:
        raise TimeoutException(
            "Could not locate the 'Import a File' menu item using any selector. "
            "Check that the main menu is open and the item is visible."
        )

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(0.2)

    # Prefer JS click for Dojo menu items — avoids intercept issues.
    try:
        driver.execute_script("arguments[0].click();", el)
    except Exception:
        safe_click(driver, el)

    return el
# ────────────────────────────────────────────────────────────────────────────────


def wait_for_clipboard_dropdown(driver, timeout=WAIT_LONG):
    xpath = (
        "//div[contains(@class,'css-13483rh-control') and "
        ".//div[contains(@class,'css-1uccc91-singleValue') and normalize-space()='Paste From Clipboard']]"
    )
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, xpath)))


def _react_control_by_display(driver, display_text: str, timeout=WAIT_MED):
    """
    Find the react-select control whose visible label (singleValue or placeholder)
    matches display_text.  Does NOT depend on auto-incremented react-select-N IDs.
    """
    xpath = (
        f"//div[contains(@class,'css-13483rh-control') and "
        f"(.//*[contains(@class,'singleValue') and normalize-space()='{display_text}'] or "
        f"  .//*[contains(@class,'placeholder') and normalize-space()='{display_text}'])]"
    )
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, xpath)))


def _set_react_select_by_display(driver, display_text: str, value_text: str, timeout=WAIT_LONG):
    """
    Set a React Select dropdown by finding it via its current visible label.
    Uses the same send_keys fallback strategy as the original set_react_select_by_input_id
    which is what actually works in this app — click to open, try clicking the option,
    fall back to typing the value + Enter into the hidden input.
    """
    value_text = str(value_text).strip()

    def open_menu():
        ctrl = _react_control_by_display(driver, display_text, timeout=timeout)
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", ctrl)
        safe_click(driver, ctrl)
        time.sleep(0.15)
        # Find the actual text input inside the control for send_keys fallback
        try:
            inp = ctrl.find_element(By.XPATH, ".//input")
            try:
                safe_click(driver, inp)
            except Exception:
                pass
        except Exception:
            inp = None
        return inp, ctrl

    def pick_option(inp, ctrl):
        # Strategy 1: click a visible [role=option] element
        for opt_xpath in [
            f"//*[@role='option' and normalize-space()='{value_text}']",
            f"//div[contains(@class,'option') and normalize-space()='{value_text}']",
            f"//*[normalize-space()='{value_text}' and (self::div or self::span) and contains(@class,'option')]",
        ]:
            try:
                opt = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((By.XPATH, opt_xpath))
                )
                safe_click(driver, opt)
                return True
            except TimeoutException:
                continue

        # Strategy 2: type the value into the input and press Enter
        if inp is not None:
            try:
                inp.send_keys(Keys.CONTROL, "a")
                inp.send_keys(value_text)
                inp.send_keys(Keys.ENTER)
                return True
            except Exception:
                pass

        return False

    for attempt in range(1, 4):
        inp, ctrl = open_menu()
        pick_option(inp, ctrl)

        # Verify the selection took — check singleValue text inside the control
        try:
            WebDriverWait(driver, 8).until(
                lambda d: (
                    bool(ctrl.find_elements(By.XPATH,
                        f".//*[contains(@class,'singleValue') and normalize-space()='{value_text}']"
                    )) or
                    bool(ctrl.find_elements(By.XPATH,
                        f".//*[normalize-space()='{value_text}']"
                    ))
                )
            )
            print(f"[INFO] React-select set: '{display_text}' → '{value_text}'")
            return
        except TimeoutException:
            if attempt == 3:
                current = ctrl.text.strip()
                raise TimeoutException(
                    f"Failed to set '{display_text}' → '{value_text}'. Control text now: '{current}'"
                )
            time.sleep(0.6)


def ensure_custom_file_selected(driver, timeout=WAIT_LONG):
    """Switch the import-mode dropdown from 'Paste From Clipboard' to 'Custom File'."""
    # Short-circuit if already set
    try:
        _react_control_by_display(driver, "Custom File", timeout=3)
        print("[INFO] Import mode already set to 'Custom File'")
        return
    except TimeoutException:
        pass

    # Use the combined display-text finder + send_keys setter
    _set_react_select_by_display(driver, "Paste From Clipboard", "Custom File", timeout=timeout)


def _react_control_for_input(driver, input_id: str):
    inp = WebDriverWait(driver, WAIT_LONG).until(EC.presence_of_element_located((By.ID, input_id)))
    ctrl = inp.find_element(By.XPATH, "./ancestor::div[contains(@class,'css-13483rh-control')][1]")
    return inp, ctrl


def _react_control_text(ctrl) -> str:
    txt = (ctrl.text or "").strip()
    txt = re.sub(r"\s+", " ", txt)
    return txt


def set_react_select_by_input_id(driver, input_id: str, value_text: str, timeout=WAIT_LONG):
    value_text = str(value_text).strip()

    def open_menu():
        inp, ctrl = _react_control_for_input(driver, input_id)
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", ctrl)
        safe_click(driver, ctrl)
        time.sleep(0.15)
        try:
            safe_click(driver, inp)
        except Exception:
            pass
        return inp, ctrl

    def pick_option(inp, ctrl):
        try:
            opt = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, f"//*[@role='option' and normalize-space()='{value_text}']"))
            )
            safe_click(driver, opt)
            return True
        except TimeoutException:
            pass

        try:
            inp.send_keys(Keys.CONTROL, "a")
            inp.send_keys(value_text)
            inp.send_keys(Keys.ENTER)
            return True
        except Exception:
            return False

    for attempt in range(1, 4):
        inp, ctrl = open_menu()
        _ = pick_option(inp, ctrl)

        try:
            WebDriverWait(driver, 8).until(
                lambda d: (
                    value_text == _react_control_text(ctrl) or
                    f" {value_text} " in f" {_react_control_text(ctrl)} " or
                    bool(ctrl.find_elements(By.XPATH, f".//div[contains(@class,'singleValue') and normalize-space()='{value_text}']")) or
                    bool(ctrl.find_elements(By.XPATH, f".//*[normalize-space()='{value_text}']"))
                )
            )
            return
        except TimeoutException:
            if attempt == 3:
                final_txt = _react_control_text(ctrl)
                raise TimeoutException(
                    f"Failed to set {input_id} to '{value_text}'. Control text now: '{final_txt}'"
                )
            time.sleep(0.6)
# ────────────────────────────────────────────────────────────────────────────────


def find_latest_matching_file(order_no: str) -> str:
    # For Ariat uploads, ONLY look for files containing "ariat" or "carhartt"
    # This prevents accidentally selecting wrangler or other vendor files
    patterns = [
        os.path.join(DOWNLOAD_FOLDER, f"*{order_no}*ariat*carhartt*.*"),
        os.path.join(DOWNLOAD_FOLDER, f"*{order_no}*ariat*.*"),
        os.path.join(DOWNLOAD_FOLDER, f"*{order_no}*carhartt*.*"),
    ]
    candidates = []
    for pat in patterns:
        for f in glob.glob(pat):
            if f.lower().endswith((".xlsx", ".xls", ".xlsm", ".xlsb", ".ods", ".csv", ".txt")):
                candidates.append(f)

    if not candidates:
        available = os.listdir(DOWNLOAD_FOLDER)
        raise FileNotFoundError(
            f"No Ariat upload file found for order {order_no} in {DOWNLOAD_FOLDER}.\n"
            f"Looked for files matching: *{order_no}*ariat* or *{order_no}*carhartt*\n"
            f"Folder contains (first 80): {available[:80]}"
        )

    candidates.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return candidates[0]


def _read_shipto_from_csv(csv_path: str) -> dict:
    df = pd.read_csv(csv_path, dtype=str)
    if df.empty:
        raise ValueError(f"PO CSV {csv_path} has no rows.")
    row = df.iloc[0].to_dict()

    def get(*keys):
        for k in keys:
            if k in row and pd.notna(row[k]):
                return coerce_str(row[k])
        return ""

    company   = get("shipToCompany", "K")
    attention = get("shipToAttention", "L")
    street    = get("shipToStreet", "M")
    city      = get("shipToCity", "N")
    state     = get("shipToState", "O")
    zipc      = get("shipToZip", "P")

    name_line = (company + (attention or "")).strip()  # no separator requested

    return {
        "name_line": name_line,
        "street": street,
        "city": city,
        "state_abbr": state.upper().strip(),
        "zip": zipc,
        "csv_path": csv_path,
    }


def load_shipto_from_po_csv(po_number: str) -> dict:
    """
    Handles naming mismatch:
      - orders sheet might have '162945-297239'
      - CSV is named '297239.csv' and PO column is '297239'
    """
    po_raw = coerce_str(po_number)
    po_key = extract_po_key(po_raw)

    pats = [
        os.path.join(PDF_DIR, f"{po_raw}.csv"),
        os.path.join(PDF_DIR, f"{po_key}.csv"),
        os.path.join(PDF_DIR, f"*{po_key}*.csv"),
        os.path.join(PDF_DIR, f"*{po_raw}*.csv"),
    ]

    matches = []
    for pat in pats:
        matches.extend(glob.glob(pat))

    # If we found matches, prefer exact po_key.csv, then exact po_raw.csv, then most recent
    if matches:
        # De-dupe while preserving order
        seen = set()
        uniq = []
        for m in matches:
            if m not in seen:
                uniq.append(m)
                seen.add(m)

        exact_key = os.path.join(PDF_DIR, f"{po_key}.csv")
        exact_raw = os.path.join(PDF_DIR, f"{po_raw}.csv")

        if exact_key in uniq:
            return _read_shipto_from_csv(exact_key)
        if exact_raw in uniq:
            return _read_shipto_from_csv(exact_raw)

        uniq.sort(key=lambda p: os.path.getmtime(p), reverse=True)
        return _read_shipto_from_csv(uniq[0])

    # Fallback: scan CSVs and match their PO column
    all_csvs = glob.glob(os.path.join(PDF_DIR, "*.csv"))
    for csv_path in all_csvs:
        try:
            df = pd.read_csv(csv_path, dtype=str, nrows=1)
            if df.empty:
                continue
            if "PO" in df.columns:
                po_val = extract_po_key(coerce_str(df.iloc[0].get("PO", "")))
                if po_val == po_key:
                    return _read_shipto_from_csv(csv_path)
        except Exception:
            continue

    raise FileNotFoundError(
        f"Could not find PO CSV for '{po_raw}' (PO key '{po_key}') in {PDF_DIR}.\n"
        f"Tried patterns: {pats}"
    )


def login_and_land(driver):
    wait = WebDriverWait(driver, WAIT_LONG)
    driver.get(ARIAT_URL)
    wait_ready(driver)

    user = wait.until(EC.visibility_of_element_located((By.NAME, "username")))
    user.clear()
    user.send_keys(ARIAT_USERNAME)

    pwd = wait.until(EC.visibility_of_element_located((By.NAME, "password")))
    pwd.clear()
    pwd.send_keys(ARIAT_PASSWORD)

    click_button_by_text(driver, "Login", timeout=WAIT_LONG)
    click_button_by_text(driver, "Shop Now", timeout=WAIT_LONG)

    try:
        wait_and_click(driver, By.CSS_SELECTOR, "[data-testid='card-carousel-image-0']", timeout=WAIT_LONG)
    except TimeoutException:
        wait_and_click(driver, By.CSS_SELECTOR, ".slick-slide.slick-current", timeout=WAIT_LONG)

    wait_for_import_menu_item(driver, timeout=WAIT_LONG)


def import_file_flow(driver, upload_path: str):
    click_import_a_file(driver, timeout=WAIT_LONG)

    # Wait for the import dialog to be ready (Paste From Clipboard dropdown visible)
    _react_control_by_display(driver, "Paste From Clipboard", timeout=WAIT_LONG)

    # Switch to Custom File mode
    ensure_custom_file_selected(driver, timeout=WAIT_LONG)

    # Upload the file
    file_input = wait_present(driver, By.CSS_SELECTOR, "input[type='file']", timeout=WAIT_LONG)
    file_input.send_keys(upload_path)

    click_button_by_text(driver, "Next", timeout=WAIT_LONG)

    # Map columns by placeholder text — immune to react-select-N ID shifts
    # The UPC/EAN/SKU dropdown maps to column A; Quantity maps to column B
    _set_react_select_by_display(driver, "UPC/EAN/SKU", "A", timeout=WAIT_LONG)
    _set_react_select_by_display(driver, "Quantity",    "B", timeout=WAIT_LONG)

    click_button_by_text(driver, "Next", timeout=WAIT_LONG)


def proceed_to_checkout_flow(driver):
    try:
        wait_and_click(
            driver,
            By.XPATH,
            "//span[contains(@class,'proceedBtn') and .//span[contains(@class,'dijitButtonText') and normalize-space()='Proceed to Checkout']]",
            timeout=WAIT_LONG
        )
    except TimeoutException:
        wait_and_click(driver, By.XPATH, "//*[normalize-space()='Proceed to Checkout']", timeout=WAIT_LONG)

    for _ in range(2):
        try:
            click_button_by_text(driver, "Proceed to Checkout", timeout=WAIT_MED)
        except TimeoutException:
            pass
    time.sleep (1.1)
    WebDriverWait(driver, WAIT_LONG).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "span.btnDropShip"))
    )



def handle_address_confirmation_popup(driver, timeout=WAIT_LONG):
    """
    Handle the address confirmation popup that appears after saving address.
    - Always selects "Suggested Address" (usually pre-selected)
    - Clicks "Use Selected Address" button
    """
    try:
        # Wait for the modal to appear
        modal = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".ReactModal__Content[aria-label='Confirm Address']"))
        )
        print("[INFO] Address confirmation popup detected")
        
        # The "Suggested Address" radio button is usually already selected by default
        # But let's ensure it's selected by clicking it
        try:
            suggested_radio = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='selectedAddress'][value='addressFromSmarty']"))
            )
            if not suggested_radio.is_selected():
                suggested_radio.click()
                print("[INFO] Selected 'Suggested Address'")
            else:
                print("[INFO] 'Suggested Address' already selected")
        except Exception as e:
            print(f"[WARN] Could not verify suggested address selection: {e}")
        
        # Click "Use Selected Address" button
        use_button = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Use Selected Address')]"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", use_button)
        time.sleep(0.5)  # Brief pause to ensure button is ready
        safe_click(driver, use_button)
        print("[INFO] Clicked 'Use Selected Address'")
        
        # Wait for modal to close
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.CSS_SELECTOR, ".ReactModal__Content[aria-label='Confirm Address']"))
        )
        print("[INFO] Address confirmation popup closed")
        
    except TimeoutException:
        print("[INFO] No address confirmation popup appeared (this is okay)")
    except Exception as e:
        print(f"[WARN] Error handling address confirmation popup: {e}")


def _handle_address_validation_warning(driver, timeout: int = 6) -> bool:
    """
    After clicking Save on the shipping address form, the site sometimes shows a
    Dijit validation warning:
        'We could not find a match for the address entered below.
         Please double check the fields highlighted in red.
         If the above address is confirmed, please click Save to continue.'

    When this banner is present the Save button must be clicked a second time to
    confirm and proceed.  If the banner does not appear within `timeout` seconds
    we assume the address was accepted on the first click and return False.

    Returns True if the warning was detected and bypassed, False otherwise.
    """
    WARNING_CSS = "div.dijitTextBoxError"
    WARNING_TEXT = "We could not find a match for the address"

    try:
        # Wait briefly for the warning banner to appear
        banner = WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, WARNING_CSS))
        )
        if WARNING_TEXT.lower() not in (banner.text or "").lower():
            # Different error — don't swallow it; let the caller surface it
            return False

        print("[WARN] Address validation warning detected — clicking Save again to confirm.")
        click_dijit_button_by_label(driver, "Save", timeout=WAIT_LONG, prefer_id="dijit_form_Button_40")
        print("[INFO] Second Save click sent to confirm unmatched address.")

        # Wait for the warning banner to disappear — confirms the form accepted
        # the second Save and has closed or moved on.
        try:
            WebDriverWait(driver, WAIT_MED).until(
                EC.invisibility_of_element_located((By.CSS_SELECTOR, WARNING_CSS))
            )
            print("[INFO] Address validation warning dismissed — form closed successfully.")
        except TimeoutException:
            print("[WARN] Warning banner did not disappear after second Save — proceeding anyway.")

        return True

    except TimeoutException:
        # Banner never appeared — address was accepted first time, nothing to do
        return False
    except Exception as e:
        print(f"[WARN] Unexpected error while handling address validation warning: {e}")
        return False


def fill_drop_ship_address(driver, po_number: str):
    addr = load_shipto_from_po_csv(po_number)

    wait_and_click(driver, By.CSS_SELECTOR, "span.btnDropShip", timeout=WAIT_LONG)

    def fill_by_input_name(input_name: str, value: str):
        if not value:
            return
        inp = wait_visible(driver, By.CSS_SELECTOR, f"input[name='{input_name}']", timeout=WAIT_LONG)
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", inp)
        inp.clear()
        inp.send_keys(value)

    fill_by_input_name("name", addr["name_line"])
    fill_by_input_name("address1", addr["street"])
    fill_by_input_name("city", addr["city"])
    fill_by_input_name("zip", addr["zip"])

    state_ab = addr["state_abbr"]
    state_name = US_STATE_ABBR_TO_NAME.get(state_ab, "")
    if state_name:
        try:
            wait_and_click(driver, By.XPATH, "//span[contains(@class,'dijitSelect') and contains(@class,'state')]", timeout=WAIT_LONG)
        except TimeoutException:
            wait_and_click(driver, By.XPATH, "//*[normalize-space()='State']/following::span[contains(@class,'dijitSelect')][1]", timeout=WAIT_LONG)

        wait_and_click(
            driver,
            By.XPATH,
            f"//td[contains(@class,'dijitMenuItemLabel') and normalize-space()='{state_name}']",
            timeout=WAIT_LONG
        )
    else:
        print(f"[WARN] Unknown/blank state abbreviation '{state_ab}' for PO {po_number}. Please select state manually.")

    # ✅ Save (Dijit) - click the actual button node, not just the inner text span
    click_dijit_button_by_label(driver, "Save", timeout=WAIT_LONG, prefer_id="dijit_form_Button_40")

    # ✅ If the site cannot validate the address it shows a dijitTextBoxError warning
    # and requires a second Save click to confirm and proceed anyway.
    warning_bypassed = _handle_address_validation_warning(driver)

    # ✅ Handle the address confirmation popup that appears after saving.
    # When the validation warning was bypassed the React confirmation modal rarely
    # appears, so use a short timeout (10 s) to avoid a 90-second stall.
    popup_timeout = WAIT_SHORT if warning_bypassed else WAIT_LONG
    handle_address_confirmation_popup(driver, timeout=popup_timeout)

    return addr

def fill_po_number_field(driver, po_number: str):
    po = coerce_str(po_number)
    try:
        inp = wait_visible(driver, By.ID, "dijit__WidgetsInTemplateMixin_4_poNumber_input", timeout=WAIT_LONG)
    except TimeoutException:
        inp = wait_visible(driver, By.XPATH, "//input[contains(@id,'poNumber') and @type='text']", timeout=WAIT_LONG)

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", inp)
    inp.clear()
    inp.send_keys(po)



def click_place_order_button(driver, timeout=WAIT_LONG):
    """
    Click the 'Place Order' button to initiate order submission.
    """
    print("[INFO] Clicking 'Place Order' button...")
    try:
        # Try by widgetid first
        place_order_btn = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "*[widgetid='finalSubmitButton']"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", place_order_btn)
        time.sleep(0.5)
        safe_click(driver, place_order_btn)
        print("[INFO] 'Place Order' button clicked")
    except TimeoutException:
        # Fallback: try clicking by button text
        print("[INFO] Trying fallback method to find 'Place Order' button...")
        click_dijit_button_by_label(driver, "Place Order", timeout=timeout)
        print("[INFO] 'Place Order' button clicked (fallback method)")


def handle_order_confirmation_popup(driver, timeout=WAIT_LONG):
    """
    Handle the order confirmation popup that appears after clicking 'Place Order'.
    Clicks the 'Submit' button in the confirmation dialog.
    """
    print("[INFO] Waiting for order confirmation popup...")
    try:
        # Wait for the confirmation dialog to appear
        confirmation_dialog = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".dijitDialog.modal-confirm"))
        )
        print("[INFO] Order confirmation popup detected")
        
        # Wait a moment for the dialog to fully render
        time.sleep(1)
        
        # Click the Submit button (id="dijit_form_Button_44" or search by label)
        try:
            submit_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "dijit_form_Button_44"))
            )
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", submit_btn)
            safe_click(driver, submit_btn)
            print("[INFO] Clicked 'Submit' button in confirmation popup")
        except TimeoutException:
            # Fallback: click by button text
            print("[INFO] Trying fallback method to find 'Submit' button...")
            click_dijit_button_by_label(driver, "Submit", timeout=10)
            print("[INFO] 'Submit' button clicked (fallback method)")
        
        # Wait for confirmation dialog to close
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.CSS_SELECTOR, ".dijitDialog.modal-confirm"))
        )
        print("[INFO] Order confirmation popup closed")
        
    except TimeoutException:
        print("[WARN] Order confirmation popup did not appear within timeout")
    except Exception as e:
        print(f"[WARN] Error handling order confirmation popup: {e}")


def extract_order_id_from_success_popup(driver, timeout=WAIT_LONG):
    """
    Wait for the order submission success popup and extract the order ID.
    Returns the order ID string (e.g., "10744371") or None if not found.
    """
    print("[INFO] Waiting for order submission success popup...")
    try:
        # Wait for the success dialog to appear
        success_dialog = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".dijitDialog.submitOkModal"))
        )
        print("[INFO] Order submission success popup detected")
        
        # Find the description paragraph containing the order ID
        description = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".submitOkModalContents p[data-dojo-attach-point='description']"))
        )
        
        # Extract the text
        success_text = description.text
        print(f"[INFO] Success message: {success_text}")
        
        # Extract order ID using regex (e.g., "Order 10744371 submitted successfully...")
        match = re.search(r"Order\s+(\d+)\s+submitted", success_text, re.IGNORECASE)
        if match:
            order_id = match.group(1)
            print(f"[SUCCESS] Extracted Order ID: {order_id}")
            
            # Click "Okay" to close the success dialog
            try:
                okay_btn = driver.find_element(By.XPATH, "//div[contains(@class,'submitOkModal')]//span[contains(@class,'dijitButtonText') and normalize-space()='Okay']/..")
                safe_click(driver, okay_btn)
                print("[INFO] Clicked 'Okay' to close success popup")
            except Exception:
                print("[WARN] Could not find 'Okay' button, popup may close automatically")
            
            return order_id
        else:
            print(f"[ERROR] Could not extract order ID from success message: {success_text}")
            return None
            
    except TimeoutException:
        print("[ERROR] Order submission success popup did not appear within timeout")
        return None
    except Exception as e:
        print(f"[ERROR] Error extracting order ID from success popup: {e}")
        return None


def update_order_id_in_excel(excel_path: str, row_index: int, order_id: str):
    """
    Update the Order ID in column M (index 12) of the Excel file.
    If there's already a value in column M, append the new order ID with a space separator.
    
    Args:
        excel_path: Path to the Excel file
        row_index: The pandas DataFrame index (row number)
        order_id: The order ID to add
    """
    try:
        print(f"[INFO] Updating Excel file with Order ID: {order_id}")
        
        # Read the Excel file
        df = pd.read_excel(excel_path, engine="openpyxl", dtype=str)
        
        # Ensure column M (index 12) exists for Order ID
        if len(df.columns) < 13:
            # Add columns if needed
            while len(df.columns) < 13:
                df.insert(len(df.columns), f'Column_{len(df.columns)}', '')
        
        # Get or create column M
        col_m_name = df.columns[12]
        
        # Get existing value in column M for this row
        existing_value = coerce_str(df.at[row_index, col_m_name])
        
        # Append or set the order ID
        if existing_value:
            new_value = f"{existing_value} {order_id}"
            print(f"[INFO] Appending to existing value: '{existing_value}' → '{new_value}'")
        else:
            new_value = order_id
            print(f"[INFO] Setting new Order ID: '{new_value}'")
        
        df.at[row_index, col_m_name] = new_value
        
        # Save back to Excel
        df.to_excel(excel_path, index=False)
        print(f"[SUCCESS] Excel file updated: {excel_path}")
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Failed to update Excel file: {e}")
        return False


def main():
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(options=options)

    try:
        login_and_land(driver)

        df = pd.read_excel(EXCEL_PATH, engine="openpyxl", dtype=str)

        # Column G (index 6) = PO number
        # Column J (index 9) = upload identifier
        # Column K (index 10) = Vendor
        col_g = df.columns[6]
        col_j = df.columns[9]
        col_k = df.columns[10]

        for idx, row in df.iterrows():
            po_number   = coerce_str(row[col_g])
            order_field = coerce_str(row[col_j])
            vendor      = coerce_str(row[col_k])

            if not po_number:
                print(f"[SKIP] Row {idx}: blank PO in column G.")
                continue

            # CRITICAL: Only process Ariat orders (skip Wrangler, Propper, etc.)
            if "ariat" not in vendor.lower():
                print(f"[SKIP] Row {idx}: Not an Ariat order (Vendor: {vendor})")
                continue

            m = re.search(r"\d+", order_field)
            if not m:
                raise ValueError(f"Row {idx}: cannot parse upload identifier from '{order_field}' (column J).")
            order_no = m.group()

            print(f"\n=== ARIAT ORDER START: PO={po_number}  UploadID={order_no} ===")

            upload_path = find_latest_matching_file(order_no)
            print(f"[INFO] Using upload file: {upload_path}")

            import_file_flow(driver, upload_path)
            proceed_to_checkout_flow(driver)

            addr = fill_drop_ship_address(driver, po_number)
            print(f"[INFO] Address loaded from: {addr['csv_path']}")

            fill_po_number_field(driver, po_number)

            # Wait for user to review and press Enter
            print("\n[ACTION REQUIRED]")
            print("Review the cart / shipping / totals")
            print("When ready, press Enter to automatically submit the order...")
            input()

            # Automatically submit the order
            try:
                click_place_order_button(driver, timeout=WAIT_LONG)
                handle_order_confirmation_popup(driver, timeout=WAIT_LONG)
                order_id = extract_order_id_from_success_popup(driver, timeout=WAIT_LONG)
                
                if order_id:
                    # Update Excel with the order ID
                    update_order_id_in_excel(EXCEL_PATH, idx, order_id)
                    print(f"[SUCCESS] Order submitted successfully! Order ID: {order_id}")
                else:
                    print("[WARNING] Order may have been submitted, but Order ID could not be extracted.")
                    print("Please check manually and update the Excel file if needed.")
                    
            except Exception as e:
                print(f"[ERROR] Error during order submission: {e}")
                print("You may need to complete the order manually.")
                input("Press Enter to continue to next order...")

            print(f"=== ARIAT ORDER DONE: PO={po_number} ===")

    finally:
        driver.quit()


if __name__ == "__main__":
    main()
