import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
)
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys

# ─── STATE ABBREVIATION → DROPDOWN VISIBLE NAME ───────────────────────────────
# Used to select the correct state in the shipping-state <select>.
STATE_ABBR_TO_NAME = {
    "AL": "Alabama",
    "AK": "Alaska",
    "AZ": "Arizona",
    "AR": "Arkansas",
    "CA": "California",
    "CO": "Colorado",
    "CT": "Connecticut",
    "DE": "Delaware",
    "FL": "Florida",
    "GA": "Georgia",
    "HI": "Hawaii",
    "ID": "Idaho",
    "IL": "Illinois",
    "IN": "Indiana",
    "IA": "Iowa",
    "KS": "Kansas",
    "KY": "Kentucky",
    "LA": "Louisiana",
    "ME": "Maine",
    "MD": "Maryland",
    "MA": "Massachusetts",
    "MI": "Michigan",
    "MN": "Minnesota",
    "MS": "Mississippi",
    "MO": "Missouri",
    "MT": "Montana",
    "NE": "Nebraska",
    "NV": "Nevada",
    "NH": "New Hampshire",
    "NJ": "New Jersey",
    "NM": "New Mexico",
    "NY": "New York",
    "NC": "North Carolina",
    "ND": "North Dakota",
    "OH": "Ohio",
    "OK": "Oklahoma",
    "OR": "Oregon",
    "PA": "Pennsylvania",
    "RI": "Rhode Island",
    "SC": "South Carolina",
    "SD": "South Dakota",
    "TN": "Tennessee",
    "TX": "Texas",
    "UT": "Utah",
    "VT": "Vermont",
    "VA": "Virginia",
    "WA": "Washington",
    "WV": "West Virginia",
    "WI": "Wisconsin",
    "WY": "Wyoming",
    # Optional extras (harmless if ever present in data)
    "DC": "District Of Columbia",
}

# ─── CUSTOM EXCEPTIONS ───────────────────────────────────────────────────
class UnorderableSizeError(Exception):
    """Raised when the desired size exists on the page but is not orderable (no qty input)."""


# ─── CONFIG ────────────────────────────────────────────────────────────────────
CREDENTIALS = {
    "jmccarthy@thesourcinggroup.com": "TSG2025$",
    "ashotwell@thesourcinggroup.com": "Welcome2TSG!",
    "mdelgado@thesourcinggroup.com": "TSG2024$",
    # add others like "alvina@..." / "jessica@..." if needed
}

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PDF_DIR = os.path.join(SCRIPT_DIR, "pdfs")   # legacy location
CSV_DIRS = [PDF_DIR, SCRIPT_DIR]             # search both

LOGIN_URL = "https://shop.broberry.com/login"
ORDER_HISTORY_URL = "https://shop.broberry.com/account/order-history"
SUMMARY_URL = "https://shop.broberry.com/shop/order/summary"
ADDRESS_URL = "https://shop.broberry.com/shop/order/address"

PRODUCT_MAP = {
    "3W045CH": {
        "url": "https://shop.broberry.com/shop/product/1406495",
        "sizes": [30, 31, 32, 33, 34, 35, 36, 38, 40, 42, 44, 46],
        "mode": "grid",
    },
    "3W045DK": {
        "url": "https://shop.broberry.com/shop/product/1406722",
        "sizes": [30, 31, 32, 33, 34, 35, 36, 38, 40, 42, 44, 46],
        "mode": "grid",
    },
    "3W060BR": {
        "url": "https://shop.broberry.com/shop/product/1407233",
        "sizes": [30, 31, 32, 33, 34, 35, 36, 38, 40, 42, 44, 46, 48, 50, 52, 54, 56, 58, 60, 62],
        "mode": "grid",
    },
    "10FR13MWZ": {
        "url": "https://shop.broberry.com/shop/product/1392822",
        "sizes": [30, 31, 32, 33, 34, 35, 36, 38, 40, 42, 44, 46, 48, 50, 52, 54],
        "mode": "auto",
    },
    "10FR13MMS": {
        "url": "https://shop.broberry.com/shop/product/1392784",
        "sizes": [30, 31, 32, 33, 34, 35, 36, 38, 40, 42],
        "mode": "auto",
    },
    "10FR47MLW": {
        "url": "https://shop.broberry.com/shop/product/1393239",
        "sizes": [30, 31, 32, 33, 34, 35, 36, 38, 40, 42, 44, 46, 48, 50, 52, 54],
        "mode": "auto",
    },
    "F52944X250": {
        "url": "https://shop.broberry.com/shop/product/1095437",
        "sizes": [30, 32, 34, 36, 38, 40, 42, 44, 46, 48, 50, 52, 54, 56],
        "mode": "auto",
    },
    "F52594X250": {
        "url": "https://shop.broberry.com/shop/product/1094476",
        "sizes": [2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24],
        "mode": "auto",
    },
    "10030232": {
        "url": "https://shop.broberry.com/shop/product/1083190",
        "sizes": [28, 29, 30, 31, 32, 33, 34, 35, 36, 38, 40, 42, 44],
        "mode": "auto",
    },
}

# CH<->DK can sub for each other
PAIRABLE = {
    # CH<->DK can sub for each other
    "3W045CH": "3W045DK",
    "3W045DK": "3W045CH",

    # 10FR13MWZ <-> 10FR13MMS can sub for each other (back order / unavailable)
    "10FR13MWZ": "10FR13MMS",
    "10FR13MMS": "10FR13MWZ",
}

# ─── DRIVER SETUP ───────────────────────────────────────────────────────────────
def init_driver():
    opts = webdriver.ChromeOptions()
    # launch clean each time
    opts.add_argument("--incognito")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--disable-save-password-bubble")
    opts.add_argument("--disable-features=AutofillKeyBoardAccessoryView,PasswordManagerOnboarding,OptimizationHints")
    opts.add_experimental_option("prefs", {
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
    })
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=opts)
    driver.maximize_window()
    driver.implicitly_wait(5)
    return driver

# ─── LOGIN / LOGOUT ─────────────────────────────────────────────────────────────
def login(driver, email, password):
    print(f"⇢ Logging in as {email} …")
    driver.get(LOGIN_URL)
    wait = WebDriverWait(driver, 20)

    email_in = wait.until(EC.element_to_be_clickable((By.NAME, "email")))
    email_in.clear(); email_in.send_keys(email)

    pwd_in = wait.until(EC.element_to_be_clickable((By.NAME, "password")))
    pwd_in.clear(); pwd_in.send_keys(password)

    sign_in_btn = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//button[@type='submit' and normalize-space(text())='Sign in']")))
    sign_in_btn.click()

    # consider both redirect and dom confirmation
    try:
        WebDriverWait(driver, 20).until(
            EC.any_of(
                EC.url_contains("/account"),
                EC.presence_of_element_located((By.XPATH, "//*[contains(.,'Order History') or contains(.,'My Account')]"))
            )
        )
        print(f"→ Logged in as {email}")
        return True
    except TimeoutException:
        print(f"✖ Login did not complete for {email}. Check credentials or MFA prompts.")
        return False

# ─── CSV DISCOVERY / ACCOUNT EXTRACTION ─────────────────────────────────────────
def _discover_csvs():
    seen, found = set(), []
    for base in CSV_DIRS:
        if not os.path.isdir(base): continue
        for f in os.listdir(base):
            if f.lower().endswith(".csv"):
                p = os.path.join(base, f)
                if p not in seen:
                    seen.add(p); found.append(p)
    return sorted(found, key=lambda p: os.path.basename(p).lower())

def _read_account_from_df(df):
    lower = {c.lower(): c for c in df.columns}
    for key in ("email", "account", "acct"):
        if key in lower:
            return str(df.iloc[0][lower[key]]).strip().lower()
    # fallback: first column's first value
    first_col = df.columns[0]
    return str(df.iloc[0][first_col]).strip().lower()

def discover_csvs_with_accounts():
    items = []
    for p in _discover_csvs():
        try:
            df = pd.read_csv(p)
            acct = _read_account_from_df(df)
            items.append((p, acct))
        except Exception as e:
            print(f"⚠️  Could not read {os.path.basename(p)} ({e}). Skipping.")
    return items

# ─── SKIPPED ORDERS EXCEL HELPER ────────────────────────────────────────────────

SKIPPED_ORDERS_PATH = os.path.join(SCRIPT_DIR, "skipped_orders.xlsx")

def log_skipped_order(po_number, reason):
    """
    Append a row to skipped_orders.xlsx in the script directory.
    Creates the file if it does not exist yet.
    """
    new_row = pd.DataFrame([{"PO": po_number, "Reason": reason}])
    try:
        if os.path.exists(SKIPPED_ORDERS_PATH):
            existing = pd.read_excel(SKIPPED_ORDERS_PATH)
            updated = pd.concat([existing, new_row], ignore_index=True)
        else:
            updated = new_row
        updated.to_excel(SKIPPED_ORDERS_PATH, index=False)
        print(f"📝 Logged skipped order {po_number} to skipped_orders.xlsx: {reason}")
    except Exception as e:
        print(f"⚠️  Could not log skipped order {po_number} to Excel: {e}")


# ─── PRODUCT / SUMMARY HELPERS ──────────────────────────────────────────────────

def _locate_qty_input_and_context(driver, sku, waist, inseam):
    """Locate the qty input for a given SKU/size.

    Supports both:
      1) Waist×inseam grid (inseam as sticky row header, waist as columns), and
      2) Single-dimension tables (waist only as row labels).
    """
    sizes = PRODUCT_MAP[sku].get("sizes", [])
    mode = PRODUCT_MAP[sku].get("mode", "auto")

    try:
        waist_i = int(waist) if waist is not None and str(waist).strip() != "" else None
    except Exception:
        waist_i = None
    try:
        inseam_i = int(inseam) if inseam is not None and str(inseam).strip() != "" else None
    except Exception:
        inseam_i = None

    def try_grid():
        # requires inseam row header + waist list
        if inseam_i is None:
            return None
        if sizes and waist_i not in sizes:
            return None

        # Locate the inseam row in the size grid
        header_td = driver.find_element(
            By.XPATH,
            f"//td[contains(@class,'sticky') and normalize-space()='{inseam_i}']"
        )

        # New products sometimes shift the waist columns (extra columns, labels, etc.).
        # Instead of relying on a hard-coded index, derive the waist column from the
        # table header whenever possible.
        def _col_index_for_waist(table, w):
            w = str(w).strip()
            xpaths = [
                # Preferred: header area
                f".//thead//tr//*[self::th or self::td][normalize-space()='{w}' and not(.//input)]",
                # Common fallback when there's no <thead>
                f".//tr[1]//*[self::th or self::td][normalize-space()='{w}' and not(.//input)]",
                # Last resort: anywhere in the table (avoid inputs)
                f".//*[self::th or self::td][normalize-space()='{w}' and not(.//input)]",
            ]
            for xp in xpaths:
                els = table.find_elements(By.XPATH, xp)
                if not els:
                    continue
                el = els[0]
                return len(el.find_elements(By.XPATH, "preceding-sibling::*[self::th or self::td]")) + 1
            return None

        # Resolve the target cell by column index
        row = header_td.find_element(By.XPATH, "ancestor::tr[1]")
        table = header_td.find_element(By.XPATH, "ancestor::table[1]")
        col_idx = _col_index_for_waist(table, waist_i)

        # Fallback to the old behavior if we couldn't resolve the header column.
        if col_idx is None and sizes:
            # sizes.index(...) is 0-based and references the waist columns only;
            # add 2 to account for the sticky inseam cell at the start of the row.
            col_idx = sizes.index(waist_i) + 2

        if col_idx is None:
            return None

        row_cells = row.find_elements(By.XPATH, "./*[self::td or self::th]")
        if col_idx < 1 or col_idx > len(row_cells):
            return None
        cell = row_cells[col_idx - 1]
        inputs = cell.find_elements(By.CSS_SELECTOR, "input[type='number']")
        if inputs:
            return inputs[0], cell
        raise UnorderableSizeError(f"{sku} size {waist_i}{('x'+str(inseam_i)) if inseam_i is not None else ''} is not orderable (no qty input)")

    def try_row():
        # single-dimension table where the size is a row label
        row = driver.find_element(By.XPATH, f"//tr[.//*[self::td or self::th][normalize-space()='{waist_i}']]")
        qty_inputs = row.find_elements(By.CSS_SELECTOR, "input[type='number']")
        if not qty_inputs:
            raise UnorderableSizeError(f"{sku} size {waist_i} is not orderable (no qty input)")
        qty_input = qty_inputs[0]
        return qty_input, row

    if mode == "grid":
        try:
            return try_grid()
        except Exception:
            return None
    if mode == "row":
        try:
            return try_row()
        except Exception:
            return None

    # auto: try grid first, then row-based
    try:
        res = try_grid()
        if res:
            return res
    except Exception:
        pass

    try:
        return try_row()
    except Exception:
        return None


def try_add_line(driver, sku, waist, inseam, qty):
    driver.get(PRODUCT_MAP[sku]["url"])
    time.sleep(0.5)

    try:
        located = _locate_qty_input_and_context(driver, sku, waist, inseam)
    except UnorderableSizeError as e:
        return ('fatal_unorderable_size', str(e))
    if not located:
        return ('unavailable', 'size not found on page')

    qty_input, context = located

    if qty_input.get_attribute("disabled") or qty_input.get_attribute("readonly"):
        return ('unavailable', 'qty input disabled')

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", qty_input)
    qty_input.click()
    qty_input.send_keys(Keys.CONTROL, "a")
    qty_input.send_keys(Keys.DELETE)
    qty_input.send_keys(str(qty))
    qty_input.send_keys(Keys.TAB)

    # find Add button as close to the input as possible
    try:
        form = context.find_element(By.XPATH, "ancestor::form[1]")
        add_btn = form.find_element(
            By.XPATH,
            ".//button[@type='submit' and (contains(@class,'bg-green-600') or contains(., 'Add'))]"
        )
    except NoSuchElementException:
        add_btn = driver.find_element(
            By.XPATH,
            "//button[@type='submit' and (contains(@class,'bg-green-600') or contains(., 'Add'))]"
        )

    try:
        driver.execute_script("arguments[0].click();", add_btn)
    except Exception:
        add_btn.click()

    time.sleep(1)
    return ('added', None)

def extract_sku_from_text(text):
    for code in PRODUCT_MAP.keys():
        if code in text:
            return code
    return None

def find_summary_row(driver, sku, waist, inseam):
    rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
    for tr in rows:
        tds = tr.find_elements(By.CSS_SELECTOR, "td")
        if not tds or len(tds) < 9:
            continue
        item_text = tds[0].text.strip()
        if sku and sku not in item_text:
            continue

        # Standard layout: waist in col 5, inseam in col 6
        try:
            w = int(tds[4].text)
        except ValueError:
            continue

        # If inseam is missing/NA for this product, match on waist only.
        if inseam is None or str(inseam).strip() == "":
            if w == int(waist):
                return tr
            continue

        try:
            i = int(tds[5].text)
        except ValueError:
            continue

        if w == int(waist) and i == int(inseam):
            return tr
    return None

def is_backorder_row(tr):
    try:
        return "Back Order" in tr.text
    except StaleElementReferenceException:
        return False

def remove_summary_row(driver, tr):
    wait = WebDriverWait(driver, 10)
    btn = tr.find_element(By.CSS_SELECTOR, "button.text-rose-600")
    driver.execute_script("arguments[0].click();", btn)
    try:
        delete_xpath = (
            "//button[normalize-space()='Delete' and contains(@class,'bg-red-')]"
            " | //div[contains(@class,'modal') or contains(@role,'dialog')]//button[normalize-space()='Delete']"
            " | //button[normalize-space()='Remove']"
            " | //button[normalize-space()='Yes, delete']"
        )
        delete_btn = wait.until(EC.element_to_be_clickable((By.XPATH, delete_xpath)))
        driver.execute_script("arguments[0].click();", delete_btn)
    except TimeoutException:
        pass
    wait.until(EC.staleness_of(tr))
    time.sleep(0.3)

# ─── ORDER PROCESSING ─────────────────────────────────────

def has_propper_or_wrangler_items(driver):
    """
    Check if the current cart/summary contains Propper or Wrangler items.
    Returns True if any items match Propper (F*) or Wrangler (3W*, 10FR*) product codes.
    """
    try:
        driver.get(SUMMARY_URL)
        time.sleep(0.5)
        rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
        
        for tr in rows:
            tds = tr.find_elements(By.CSS_SELECTOR, "td")
            if not tds or len(tds) < 1:
                continue
            item_text = tds[0].text.strip()
            
            # Check for Propper items (start with F)
            if item_text.startswith("F52944") or item_text.startswith("F52594"):
                return True
            
            # Check for Wrangler items (start with 3W or 10FR)
            if (item_text.startswith("3W045") or 
                item_text.startswith("3W060") or 
                item_text.startswith("10FR13") or 
                item_text.startswith("10FR47")):
                return True
                
        return False
    except Exception as e:
        print(f"⚠️  Could not check for Propper/Wrangler items: {e}")
        return False

def clear_cart(driver):
    """Remove all lines from the cart summary (used when skipping an order)."""
    driver.get(SUMMARY_URL)
    time.sleep(0.8)
    while True:
        rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
        if not rows:
            break
        try:
            remove_summary_row(driver, rows[0])
        except Exception:
            driver.get(SUMMARY_URL)
            time.sleep(0.8)
    print("🧹 Cart cleared (order skipped).")

def fill_address_and_notes(driver, po, notes,
                           ship_company=None, ship_attention=None,
                           ship_street=None, ship_city=None, ship_state=None, ship_zip=None):
    """Fill PO + notes, and (if provided) update the shipping address fields."""
    wait = WebDriverWait(driver, 10)
    driver.get(ADDRESS_URL)

    # Always fill PO into last-name slots
    billing_ln  = wait.until(EC.element_to_be_clickable((By.ID, "billing-last-name")))
    shipping_ln = wait.until(EC.element_to_be_clickable((By.ID, "shipping-last-name")))
    for el in (billing_ln, shipping_ln):
        el.clear()
        el.send_keys(str(po))

    po_fld  = wait.until(EC.element_to_be_clickable((By.ID, "order-purchase-order")))
    po_fld.clear()
    po_fld.send_keys(str(po))

    notes_f = wait.until(EC.element_to_be_clickable((By.NAME, "order[notes]")))
    notes_f.clear()
    if notes:
        notes_f.send_keys("\n".join(notes))

    def _norm(v):
        if v is None:
            return ""
        s = str(v).strip()
        return "" if (not s or s.lower() == "nan") else s

    # Company line: no connector, just a single space
    company_line_parts = [p for p in (_norm(ship_company), _norm(ship_attention)) if p]
    company_line = " ".join(company_line_parts).strip()

    def _get_by_id(id_, timeout=2):
        try:
            return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.ID, id_)))
        except Exception:
            return None

    comp_el = _get_by_id("shipping-company", timeout=2)
    addr_el = _get_by_id("shipping-address-1", timeout=2)
    city_el = _get_by_id("shipping-city", timeout=2)
    zip_el  = _get_by_id("shipping-postal-code", timeout=2)
    st_el   = _get_by_id("shipping-state", timeout=2)

    if comp_el and company_line:
        comp_el.clear(); comp_el.send_keys(company_line)
    if addr_el and _norm(ship_street):
        addr_el.clear(); addr_el.send_keys(_norm(ship_street))
    if city_el and _norm(ship_city):
        city_el.clear(); city_el.send_keys(_norm(ship_city))
    if zip_el and _norm(ship_zip):
        zip_el.click()
        zip_el.send_keys(Keys.CONTROL, "a")
        zip_el.send_keys(_norm(ship_zip))

    if st_el and _norm(ship_state):
        state_in = _norm(ship_state).upper()
        state_name = STATE_ABBR_TO_NAME.get(state_in, _norm(ship_state))
        try:
            from selenium.webdriver.support.ui import Select
            Select(st_el).select_by_visible_text(state_name)
        except Exception:
            pass

    wait.until(EC.element_to_be_clickable((
        By.CSS_SELECTOR, "button.w-full.rounded-md.bg-rose-600"
    ))).click()
    time.sleep(1)

def fill_shipper_number(driver, shipper_number):
    """Fill the shipper number field on the shipping-and-payment page."""
    if not shipper_number:
        return
    
    wait = WebDriverWait(driver, 10)
    try:
        shipper_field = wait.until(EC.element_to_be_clickable((By.ID, "order-shipper-number")))
        shipper_field.clear()
        shipper_field.send_keys(str(shipper_number))
        print(f"✓ Shipper number {shipper_number} added to order")
    except Exception as e:
        print(f"⚠️  Could not fill shipper number field: {e}")

def submit_order(driver):
    wait = WebDriverWait(driver, 20)
    try:
        complete_btn = wait.until(EC.element_to_be_clickable((
            By.XPATH, "//button[.//span[normalize-space()='Complete Checkout'] and not(@disabled)]"
        )))
    except TimeoutException:
        complete_btn = wait.until(EC.element_to_be_clickable((
            By.XPATH, "//button[contains(@class,'bg-green-600') and not(@disabled)]"
        )))
    try:
        driver.execute_script("arguments[0].click();", complete_btn)
    except Exception:
        complete_btn.click()
    try:
        WebDriverWait(driver, 25).until(
            EC.any_of(
                EC.url_contains("/shop/order/complete"),
                EC.url_contains("/shop/order/confirmation"),
                EC.presence_of_element_located((By.XPATH, "//*[contains(., 'Thank you for your order') or contains(., 'Order Complete')]"))
            )
        )
        print("✅ Order submitted.")
    except TimeoutException:
        print("⚠️  Submitted click issued, but confirmation not detected. Proceeding.")
time.sleep(1)

def process_csv(driver, csv_path):
    print(f"\n=== Processing {os.path.basename(csv_path)} ===")
    df = pd.read_csv(csv_path)
    
    # Handle both old and new column name formats
    column_map = {
        'po': 'PO',
        'productId': 'Item-Number',
        'size1': 'Size-1',
        'size2': 'Size-2',
        'qty': 'Qty'
    }
    
    # Rename columns if they match the new format
    df = df.rename(columns=column_map)
    
    po_number = str(df.iloc[0]["PO"])
    notes = []

    # --- Ship-to fields from the new extraction format ---
    def _get_col(*cands):
        cols_lower = {c.lower(): c for c in df.columns}
        for c in cands:
            if c in df.columns:
                return c
            if c.lower() in cols_lower:
                return cols_lower[c.lower()]
        return None

    ship_company   = df.iloc[0].get(_get_col("ShipToCompany", "shipToCompany", "ShiptoCompany"), "")
    ship_attention = df.iloc[0].get(_get_col("ShipToAttention", "shipToAttention", "ShiptoAttention"), "")
    ship_street    = df.iloc[0].get(_get_col("ShipToStreet", "shipToStreet", "ShiptoStreet", "ShipToAddress1", "shipToAddress1"), "")
    ship_city      = df.iloc[0].get(_get_col("ShipToCity", "shipToCity", "ShiptoCity"), "")
    ship_state     = df.iloc[0].get(_get_col("ShipToState", "shipToState", "ShiptoState"), "")
    ship_zip       = df.iloc[0].get(_get_col("ShipToZip", "shipToZip", "ShiptoZip", "ShipToPostalCode", "shipToPostalCode"), "")

    # Add all requested lines, attempt substitutions for unavailable
    unavailable_lines = []
    for _, row in df.iterrows():
        sku    = str(row["Item-Number"]).strip()
        waist_v  = row.get("Size-1", "")
        inseam_v = row.get("Size-2", "")

        # Some products are single-dimension (waist only). Allow blank/NA inseam.
        waist  = int(waist_v) if str(waist_v).strip() != "" and str(waist_v).lower() != "nan" else None
        inseam = int(inseam_v) if str(inseam_v).strip() != "" and str(inseam_v).lower() != "nan" else None
        qty_v  = row.get("Qty", 0)
        qty    = int(qty_v) if str(qty_v).strip() != "" and str(qty_v).lower() != "nan" else 0

        status, reason = try_add_line(driver, sku, waist, inseam, qty)
        if status == 'fatal_unorderable_size':
            print(f"⛔ Skipping order {po_number}: {reason}")
            clear_cart(driver)
            return
        if status == 'unavailable':
            unavailable_lines.append({"sku": sku, "waist": waist, "inseam": inseam, "qty": qty})

    still_unresolved = []
    for rec in unavailable_lines:
        sku = rec["sku"]
        if sku in PAIRABLE:
            alt = PAIRABLE[sku]
            status, _ = try_add_line(driver, alt, rec["waist"], rec["inseam"], rec["qty"])
            if status == 'added':
                notes.append(f"{sku} {rec['waist']}{('x'+str(rec['inseam'])) if rec['inseam'] is not None else ''} subbed to {alt} due to unavailable size.")
            else:
                still_unresolved.append(rec)
        else:
            still_unresolved.append(rec)

    # Backorder handling on summary
    driver.get(SUMMARY_URL); time.sleep(1)
    rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
    backorders = []
    for tr in rows:
        if not is_backorder_row(tr): continue
        tds = tr.find_elements(By.CSS_SELECTOR, "td")
        if len(tds) < 9: continue
        item_text = tds[0].text.strip()
        sku = extract_sku_from_text(item_text)
        if not sku: continue
        try:
            waist = int(tds[4].text)
        except ValueError:
            continue
        try:
            inseam = int(tds[5].text)
        except ValueError:
            inseam = None
        qty = int(tds[7].find_element(By.CSS_SELECTOR, "input[type='number']").get_attribute("value"))
        backorders.append({"tr": tr, "sku": sku, "waist": waist, "inseam": inseam, "qty": qty})

    hard_blockers = []
    for bo in backorders:
        sku, waist, inseam, qty, tr = bo["sku"], bo["waist"], bo["inseam"], bo["qty"], bo["tr"]
        if sku in PAIRABLE:
            alt = PAIRABLE[sku]
            status, _ = try_add_line(driver, alt, waist, inseam, qty)
            driver.get(SUMMARY_URL); time.sleep(1)
            alt_row = find_summary_row(driver, alt, waist, inseam)
            alt_is_backorder = is_backorder_row(alt_row) if alt_row else True
            if alt_row and not alt_is_backorder:
                orig_row = find_summary_row(driver, sku, waist, inseam)
                if orig_row: remove_summary_row(driver, orig_row)
                notes.append(f"{sku} {waist}{('x'+str(inseam)) if inseam is not None else ''} subbed to {alt} due to back order.")
            else:
                if alt_row: remove_summary_row(driver, alt_row)
                hard_blockers.append({"sku": sku, "waist": waist, "inseam": inseam, "qty": qty, "reason": "both CH/DK backorder"})
        else:
            hard_blockers.append({"sku": sku, "waist": waist, "inseam": inseam, "qty": qty, "reason": "060BR backorder"})

    for rec in still_unresolved:
        if rec["sku"] in PAIRABLE:
            hard_blockers.append({**rec, "reason": "both CH/DK unavailable"})
        else:
            hard_blockers.append({**rec, "reason": "060BR unavailable"})

    if hard_blockers:
        print("Unresolvable lines:", hard_blockers)
        log_skipped_order(po_number, "Order skipped due to back order")
        clear_cart(driver)
        print(f"⛔ Order {po_number} skipped — back order could not be resolved.")
        return

    # Proceed to checkout flow
    # Quick data sanity notes (optional; can be removed to speed up)
    # ... (kept minimal to focus on login/session fix)

    # Check if order contains Propper or Wrangler items
    shipper_number = None
    if has_propper_or_wrangler_items(driver):
        shipper_number = "955617339"
        print(f"✓ Order contains Propper/Wrangler items - will use shipper number {shipper_number}")

    # Go to checkout/address page directly (more reliable than clicking a button
    # whose text/classes can change).
    fill_address_and_notes(driver, po_number, notes, ship_company, ship_attention, 
                          ship_street, ship_city, ship_state, ship_zip)

    wait = WebDriverWait(driver, 10)
    wait.until(EC.element_to_be_clickable((
        By.XPATH, "//button[contains(., 'Continue To Shipping and Payment Method')]"
    ))).click()

    # Fill shipper number on the shipping-and-payment page if needed
    if shipper_number:
        fill_shipper_number(driver, shipper_number)

    # Shipping method: try normal option first; if not available, fall back to UPS Ground (value=4 / id=4).
    try:
        ship_radio = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[name="order[shipping_id]"][value="1"]')))
        driver.execute_script("arguments[0].click();", ship_radio)
    except Exception:
        try:
            ship_radio = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[name="order[shipping_id]"][value="4"]')))
        except Exception:
            ship_radio = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, "4")))
        driver.execute_script("arguments[0].click();", ship_radio)
    pay = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="order[payment_id]"][value="1"]')))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", pay)
    try:
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[name="order[payment_id]"][value="1"]'))).click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].click();", pay)

    wait.until(EC.element_to_be_clickable((
        By.XPATH, "//button[contains(., 'Continue and Review Order')]"
    ))).click()

    submit_order(driver)

# ─── MAIN (RESTART CHROME ON ACCOUNT SWITCH) ────────────────────────────────────
def main():
    # Clear skipped orders log from any previous run
    if os.path.exists(SKIPPED_ORDERS_PATH):
        os.remove(SKIPPED_ORDERS_PATH)
        print("🗑️  Cleared skipped_orders.xlsx from previous run.")

    driver = None
    current_account = None
    try:
        csvs = discover_csvs_with_accounts()
        if not csvs:
            print("No CSVs found in ./pdfs or script directory.")
            return

        # Normalize credentials lookup
        creds_lower = {k.lower(): v for k, v in CREDENTIALS.items()}

        for csv_path, acct in csvs:
            acct_norm = (acct or "").strip().lower()
            if acct_norm not in creds_lower:
                print(f"⚠️  {os.path.basename(csv_path)}: unknown or unmapped account '{acct}'. Skipping.")
                continue

            # If switching accounts → hard reset browser
            if current_account != acct_norm:
                if driver:
                    try:
                        driver.quit()
                    except Exception:
                        pass
                driver = init_driver()

                ok = login(driver, acct_norm, creds_lower[acct_norm])
                if not ok:
                    print(f"✖ Skipping files for {acct_norm} due to login failure.")
                    current_account = None
                    continue
                current_account = acct_norm
                print(f"→ New session for account {current_account}")

            print(f"→ Using account {current_account} for {os.path.basename(csv_path)}")
            process_csv(driver, csv_path)

    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass

if __name__ == "__main__":
    main()