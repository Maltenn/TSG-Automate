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
STATE_ABBR_TO_NAME = {
    "AL": "Alabama", "AK": "Alaska", "AZ": "Arizona", "AR": "Arkansas",
    "CA": "California", "CO": "Colorado", "CT": "Connecticut", "DE": "Delaware",
    "FL": "Florida", "GA": "Georgia", "HI": "Hawaii", "ID": "Idaho",
    "IL": "Illinois", "IN": "Indiana", "IA": "Iowa", "KS": "Kansas",
    "KY": "Kentucky", "LA": "Louisiana", "ME": "Maine", "MD": "Maryland",
    "MA": "Massachusetts", "MI": "Michigan", "MN": "Minnesota", "MS": "Mississippi",
    "MO": "Missouri", "MT": "Montana", "NE": "Nebraska", "NV": "Nevada",
    "NH": "New Hampshire", "NJ": "New Jersey", "NM": "New Mexico", "NY": "New York",
    "NC": "North Carolina", "ND": "North Dakota", "OH": "Ohio", "OK": "Oklahoma",
    "OR": "Oregon", "PA": "Pennsylvania", "RI": "Rhode Island", "SC": "South Carolina",
    "SD": "South Dakota", "TN": "Tennessee", "TX": "Texas", "UT": "Utah",
    "VT": "Vermont", "VA": "Virginia", "WA": "Washington", "WV": "West Virginia",
    "WI": "Wisconsin", "WY": "Wyoming", "DC": "District Of Columbia",
}

# ─── CUSTOM EXCEPTIONS ────────────────────────────────────────────────────────
class UnorderableSizeError(Exception):
    """Raised when the desired size exists on the page but is not orderable (no qty input)."""


# ─── CONFIG ───────────────────────────────────────────────────────────────────
CREDENTIALS = {
    "jmccarthy@thesourcinggroup.com": "TSG2025$",
    "ashotwell@thesourcinggroup.com": "Welcome2TSG!",
    "mdelgado@thesourcinggroup.com": "TSG2024$",
}

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PDF_DIR = os.path.join(SCRIPT_DIR, "pdfs")
CSV_DIRS = [PDF_DIR, SCRIPT_DIR]

LOGIN_URL = "https://shop.broberry.com/login"
SUMMARY_URL = "https://shop.broberry.com/shop/order/summary"
ADDRESS_URL = "https://shop.broberry.com/shop/order/address"

SKIPPED_ORDERS_PATH = os.path.join(SCRIPT_DIR, "skipped_orders.xlsx")

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

# ─── DRIVER SETUP ─────────────────────────────────────────────────────────────
def init_driver():
    opts = webdriver.ChromeOptions()
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


# ─── LOGIN ────────────────────────────────────────────────────────────────────
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


# ─── SKU / QTY HELPERS ────────────────────────────────────────────────────────
def _locate_qty_input_and_context(driver, sku, waist, inseam):
    sizes = PRODUCT_MAP[sku].get("sizes", [])
    mode  = PRODUCT_MAP[sku].get("mode", "auto")

    try:
        waist_i = int(waist) if waist is not None and str(waist).strip() != "" else None
    except Exception:
        waist_i = None
    try:
        inseam_i = int(inseam) if inseam is not None and str(inseam).strip() != "" else None
    except Exception:
        inseam_i = None

    def try_grid():
        if inseam_i is None:
            return None
        if sizes and waist_i not in sizes:
            return None

        header_td = driver.find_element(
            By.XPATH,
            f"//td[contains(@class,'sticky') and normalize-space()='{inseam_i}']"
        )

        def _col_index_for_waist(table, w):
            w = str(w).strip()
            xpaths = [
                f".//thead//tr//*[self::th or self::td][normalize-space()='{w}' and not(.//input)]",
                f".//tr[1]//*[self::th or self::td][normalize-space()='{w}' and not(.//input)]",
                f".//*[self::th or self::td][normalize-space()='{w}' and not(.//input)]",
            ]
            for xp in xpaths:
                els = table.find_elements(By.XPATH, xp)
                if not els:
                    continue
                el = els[0]
                return len(el.find_elements(By.XPATH, "preceding-sibling::*[self::th or self::td]")) + 1
            return None

        row   = header_td.find_element(By.XPATH, "ancestor::tr[1]")
        table = header_td.find_element(By.XPATH, "ancestor::table[1]")
        col_idx = _col_index_for_waist(table, waist_i)

        if col_idx is None and sizes:
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
        raise UnorderableSizeError(
            f"{sku} size {waist_i}{('x'+str(inseam_i)) if inseam_i is not None else ''} is not orderable (no qty input)"
        )

    def try_row():
        row = driver.find_element(By.XPATH, f"//tr[.//*[self::td or self::th][normalize-space()='{waist_i}']]")
        qty_inputs = row.find_elements(By.CSS_SELECTOR, "input[type='number']")
        if not qty_inputs:
            raise UnorderableSizeError(f"{sku} size {waist_i} is not orderable (no qty input)")
        return qty_inputs[0], row

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

    try:
        form    = context.find_element(By.XPATH, "ancestor::form[1]")
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


# ─── SUMMARY / CART HELPERS ───────────────────────────────────────────────────
def extract_sku_from_text(text):
    for code in PRODUCT_MAP.keys():
        if code in text:
            return code
    return None


def clear_cart(driver):
    """Remove all lines from the cart summary."""
    driver.get(SUMMARY_URL)
    time.sleep(0.8)
    while True:
        rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
        if not rows:
            break
        try:
            _remove_summary_row(driver, rows[0])
        except Exception:
            driver.get(SUMMARY_URL)
            time.sleep(0.8)
    print("🧹 Cart cleared.")


def _remove_summary_row(driver, tr):
    wait = WebDriverWait(driver, 10)
    btn  = tr.find_element(By.CSS_SELECTOR, "button.text-rose-600")
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


def has_propper_or_wrangler_items(driver):
    """Check if the current cart contains Propper or Wrangler items."""
    try:
        driver.get(SUMMARY_URL)
        time.sleep(0.5)
        rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
        for tr in rows:
            tds = tr.find_elements(By.CSS_SELECTOR, "td")
            if not tds:
                continue
            item_text = tds[0].text.strip()
            if item_text.startswith(("F52944", "F52594", "3W045", "3W060", "10FR13", "10FR47")):
                return True
        return False
    except Exception as e:
        print(f"⚠️  Could not check for Propper/Wrangler items: {e}")
        return False


# ─── CHECKOUT HELPERS ─────────────────────────────────────────────────────────
def fill_address_and_notes(driver, po, notes,
                           ship_company=None, ship_attention=None,
                           ship_street=None, ship_city=None, ship_state=None, ship_zip=None):
    wait = WebDriverWait(driver, 10)
    driver.get(ADDRESS_URL)

    billing_ln  = wait.until(EC.element_to_be_clickable((By.ID, "billing-last-name")))
    shipping_ln = wait.until(EC.element_to_be_clickable((By.ID, "shipping-last-name")))
    for el in (billing_ln, shipping_ln):
        el.clear()
        el.send_keys(str(po))

    po_fld = wait.until(EC.element_to_be_clickable((By.ID, "order-purchase-order")))
    po_fld.clear()
    po_fld.send_keys(str(po))

    notes_f = wait.until(EC.element_to_be_clickable((By.NAME, "order[notes]")))
    notes_f.clear()
    if notes:
        notes_f.send_keys("\n".join(notes))

    def _norm(v):
        if v is None: return ""
        s = str(v).strip()
        return "" if (not s or s.lower() == "nan") else s

    company_line_parts = [p for p in (_norm(ship_company), _norm(ship_attention)) if p]
    company_line = " ".join(company_line_parts).strip()

    def _get_by_id(id_, timeout=2):
        try:
            return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.ID, id_)))
        except Exception:
            return None

    comp_el = _get_by_id("shipping-company")
    addr_el = _get_by_id("shipping-address-1")
    city_el = _get_by_id("shipping-city")
    zip_el  = _get_by_id("shipping-postal-code")
    st_el   = _get_by_id("shipping-state")

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
        state_in   = _norm(ship_state).upper()
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
    if not shipper_number:
        return
    wait = WebDriverWait(driver, 10)
    try:
        field = wait.until(EC.element_to_be_clickable((By.ID, "order-shipper-number")))
        field.clear()
        field.send_keys(str(shipper_number))
        print(f"✓ Shipper number {shipper_number} added.")
    except Exception as e:
        print(f"⚠️  Could not fill shipper number: {e}")


def _tick_shipping_as_billing_if_present(driver):
    """
    On the summary page, tick the 'ship as billing' checkbox if it exists
    and is not already checked. This must be done before clicking Complete Checkout.
    """
    try:
        cb = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.ID, "order-is-shipping-as-billing"))
        )
        if not cb.is_selected():
            driver.execute_script("arguments[0].click();", cb)
            print("☑  'Shipping as billing' checkbox ticked.")
            time.sleep(0.5)
    except TimeoutException:
        pass  # checkbox not present on this order — that's fine


def submit_order(driver):
    wait = WebDriverWait(driver, 20)

    # ── Tick the 'ship as billing' checkbox if it's on the page ──────────────
    _tick_shipping_as_billing_if_present(driver)

    # ── Click Complete Checkout ───────────────────────────────────────────────
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
                EC.presence_of_element_located((
                    By.XPATH, "//*[contains(., 'Thank you for your order') or contains(., 'Order Complete')]"
                ))
            )
        )
        print("✅ Order submitted.")
    except TimeoutException:
        print("⚠️  Submitted click issued, but confirmation not detected. Proceeding.")
    time.sleep(1)


# ─── CSV DISCOVERY ────────────────────────────────────────────────────────────
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
    return str(df.iloc[0][df.columns[0]]).strip().lower()


def _get_col(df, *cands):
    cols_lower = {c.lower(): c for c in df.columns}
    for c in cands:
        if c in df.columns:
            return c
        if c.lower() in cols_lower:
            return cols_lower[c.lower()]
    return None


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


# ─── LOAD SKIPPED POs ─────────────────────────────────────────────────────────
def load_skipped_pos():
    """Return a set of PO strings from skipped_orders.xlsx."""
    if not os.path.exists(SKIPPED_ORDERS_PATH):
        print(f"✖ {SKIPPED_ORDERS_PATH} not found. Nothing to process.")
        return set()
    try:
        df = pd.read_excel(SKIPPED_ORDERS_PATH)
        # column is named 'PO'
        col = None
        for c in df.columns:
            if c.strip().upper() == "PO":
                col = c
                break
        if col is None:
            print("⚠️  skipped_orders.xlsx has no 'PO' column.")
            return set()
        return {str(v).strip() for v in df[col].dropna()}
    except Exception as e:
        print(f"⚠️  Could not read skipped_orders.xlsx: {e}")
        return set()


# ─── ORDER PROCESSING (NO SUBSTITUTION) ──────────────────────────────────────
def process_backorder_csv(driver, csv_path):
    """
    Place a previously-skipped (back-order) order from its CSV.
    - NO substitution / sub logic.
    - Items are added as-is; backorder items remain in the cart.
    - The 'ship as billing' checkbox is handled automatically before checkout.
    """
    print(f"\n=== Processing backorder CSV: {os.path.basename(csv_path)} ===")
    df = pd.read_csv(csv_path)

    # Normalise column names
    column_map = {
        'po': 'PO', 'productId': 'Item-Number',
        'size1': 'Size-1', 'size2': 'Size-2', 'qty': 'Qty'
    }
    df = df.rename(columns=column_map)

    po_number = str(df.iloc[0]["PO"])
    notes = []

    # ── Ship-to fields ──────────────────────────────────────────────────────
    ship_company   = df.iloc[0].get(_get_col(df, "ShipToCompany",   "shipToCompany",   "ShiptoCompany"), "")
    ship_attention = df.iloc[0].get(_get_col(df, "ShipToAttention", "shipToAttention", "ShiptoAttention"), "")
    ship_street    = df.iloc[0].get(_get_col(df, "ShipToStreet",    "shipToStreet",    "ShiptoStreet",
                                              "ShipToAddress1",     "shipToAddress1"), "")
    ship_city      = df.iloc[0].get(_get_col(df, "ShipToCity",  "shipToCity",  "ShiptoCity"), "")
    ship_state     = df.iloc[0].get(_get_col(df, "ShipToState", "shipToState", "ShiptoState"), "")
    ship_zip       = df.iloc[0].get(_get_col(df, "ShipToZip",   "shipToZip",   "ShiptoZip",
                                              "ShipToPostalCode", "shipToPostalCode"), "")

    # ── Add all lines (no substitution) ────────────────────────────────────
    add_failures = []
    for _, row in df.iterrows():
        sku     = str(row["Item-Number"]).strip()
        waist_v  = row.get("Size-1", "")
        inseam_v = row.get("Size-2", "")
        waist  = int(waist_v)  if str(waist_v).strip()  not in ("", "nan") else None
        inseam = int(inseam_v) if str(inseam_v).strip() not in ("", "nan") else None
        qty_v  = row.get("Qty", 0)
        qty    = int(qty_v)   if str(qty_v).strip()    not in ("", "nan") else 0

        if sku not in PRODUCT_MAP:
            print(f"⚠️  Unknown SKU '{sku}' in {os.path.basename(csv_path)} — skipping line.")
            continue

        status, reason = try_add_line(driver, sku, waist, inseam, qty)

        if status == 'fatal_unorderable_size':
            print(f"⛔ Fatal: {reason} — aborting order {po_number}.")
            clear_cart(driver)
            return
        if status == 'unavailable':
            # On a back-order run we still skip lines that are completely
            # unavailable (no qty input / not found on page) since those
            # simply cannot be ordered.
            print(f"⚠️  Line {sku} {waist}x{inseam} qty {qty} unavailable — skipping line.")
            add_failures.append({"sku": sku, "waist": waist, "inseam": inseam, "qty": qty})
        else:
            print(f"✓ Added: {sku} {waist}{('x'+str(inseam)) if inseam is not None else ''} x{qty}")

    # ── Navigate to summary to confirm cart has items ───────────────────────
    driver.get(SUMMARY_URL)
    time.sleep(1)
    cart_rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
    if not cart_rows:
        print(f"⛔ Cart is empty after adding lines for PO {po_number}. Aborting.")
        return

    # ── Shipper number for Propper / Wrangler items ─────────────────────────
    shipper_number = None
    if has_propper_or_wrangler_items(driver):
        shipper_number = "955617339"
        print(f"✓ Propper/Wrangler items detected — shipper number {shipper_number} will be used.")

    # ── Address / notes page ────────────────────────────────────────────────
    fill_address_and_notes(
        driver, po_number, notes,
        ship_company, ship_attention,
        ship_street, ship_city, ship_state, ship_zip
    )

    wait = WebDriverWait(driver, 10)
    wait.until(EC.element_to_be_clickable((
        By.XPATH, "//button[contains(., 'Continue To Shipping and Payment Method')]"
    ))).click()

    # ── Shipper number (if needed) ──────────────────────────────────────────
    if shipper_number:
        fill_shipper_number(driver, shipper_number)

    # ── Shipping method ─────────────────────────────────────────────────────
    try:
        ship_radio = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[name="order[shipping_id]"][value="1"]'))
        )
        driver.execute_script("arguments[0].click();", ship_radio)
    except Exception:
        try:
            ship_radio = WebDriverWait(driver, 2).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[name="order[shipping_id]"][value="4"]'))
            )
        except Exception:
            ship_radio = WebDriverWait(driver, 2).until(
                EC.element_to_be_clickable((By.ID, "4"))
            )
        driver.execute_script("arguments[0].click();", ship_radio)

    # ── Payment method ──────────────────────────────────────────────────────
    pay = wait.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, 'input[name="order[payment_id]"][value="1"]')
    ))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", pay)
    try:
        wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, 'input[name="order[payment_id]"][value="1"]')
        )).click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].click();", pay)

    wait.until(EC.element_to_be_clickable((
        By.XPATH, "//button[contains(., 'Continue and Review Order')]"
    ))).click()

    # ── submit_order handles the checkbox + Complete Checkout ───────────────
    submit_order(driver)


# ═══════════════════════════════════════════════════════════════════════════════
# ─── PM LOGGING PIPELINE (ShoptoPM logic, scoped to skipped POs only) ─────────
# ═══════════════════════════════════════════════════════════════════════════════

# ── Admin credentials / paths ─────────────────────────────────────────────────
PM_INITIALS        = os.getenv("ORDER_USER_INITIALS", "MY")
ADMIN_EMAIL        = os.getenv("BROBERRY_ADMIN_EMAIL", "internal3@broberry.com")
ADMIN_PASSWORD     = os.getenv("BROBERRY_ADMIN_PASSWORD", "MYoung454$")

DOWNLOAD_FOLDER    = os.getenv("TSG_DOWNLOAD_DIR",
                        os.path.join(os.path.expanduser("~"), "Downloads"))
TEMPLATE_XLSX      = os.path.join(SCRIPT_DIR, "Example.xlsx")
OUTPUT_XLSX        = os.path.join(SCRIPT_DIR, "Processed_orders.xlsx")

ADMIN_LOGIN_URL    = "https://admin.broberry.com/login"
ORDERS_URL         = "https://admin.broberry.com/orders"

# ── Vendor detection ──────────────────────────────────────────────────────────
WRANGLER_SKU_PREFIXES = ("3W0", "10FR")
PROPPER_SKU_PREFIXES  = ("F52944X", "F52594X")
ARIAT_EXACT_SKUS      = {"10030232"}

def _detect_vendor_from_sku(sku: str):
    if not sku:
        return None
    s = str(sku).strip().upper()
    if any(s.startswith(p) for p in WRANGLER_SKU_PREFIXES):
        return "wrangler"
    if any(s.startswith(p) for p in PROPPER_SKU_PREFIXES):
        return "propper"
    if s in ARIAT_EXACT_SKUS:
        return "ariat"
    return None

def _detect_vendors_from_df(df: pd.DataFrame):
    col_map = {c.lower().strip(): c for c in df.columns}
    for key in ("productid", "product_id", "sku", "product", "item-number", "item_number"):
        if key in col_map:
            sku_col = col_map[key]
            break
    else:
        return []
    vendors = set()
    for raw in df[sku_col].dropna().astype(str).tolist():
        v = _detect_vendor_from_sku(raw)
        if v:
            vendors.add(v)
    return sorted(vendors)

# ── Build PM records from a list of CSV paths (already matched to skipped POs) ─
def _build_pm_records(matched_csvs):
    """
    matched_csvs: list of (csv_path, acct, po)
    Returns list of dicts ready for admin processing and Excel writing.
    """
    records = []
    for csv_path, acct, po in matched_csvs:
        try:
            df = pd.read_csv(csv_path, dtype=str)
        except Exception:
            try:
                df = pd.read_csv(csv_path, dtype=str, encoding="latin-1")
            except Exception as e:
                print(f"⚠️  PM: Could not read {os.path.basename(csv_path)}: {e}")
                continue

        # Normalise column names (same rename as in the order-placing side)
        df = df.rename(columns={
            'po': 'PO', 'productId': 'Item-Number',
            'size1': 'Size-1', 'size2': 'Size-2', 'qty': 'Qty'
        })

        col_map = {c.lower(): c for c in df.columns}
        row = df.iloc[0]

        # Order cost
        order_cost = ""
        for key in ("order-cost", "order cost", "ordercost", "cost", "total"):
            if key in col_map:
                order_cost = (row.get(col_map[key]) or "").strip()
                break

        records.append({
            "email":      acct,
            "PO":         po,
            "Order-Cost": order_cost,
            "vendors":    _detect_vendors_from_df(df),
        })
        print(f"  PM record: email={acct}, PO={po}, cost={order_cost}, vendors={records[-1]['vendors']}")
    return records

# ── Admin browser setup ───────────────────────────────────────────────────────
def _setup_admin_driver():
    from selenium.webdriver.chrome.options import Options as ChromeOpts
    opts = ChromeOpts()
    os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)
    prefs = {
        "download.default_directory":                        DOWNLOAD_FOLDER,
        "download.prompt_for_download":                      False,
        "download.directory_upgrade":                        True,
        "profile.default_content_settings.popups":          0,
        "profile.default_content_setting_values.automatic_downloads": 1,
        "safebrowsing.enabled":                              True,
        "safebrowsing.disable_download_protection":          True,
    }
    opts.add_experimental_option("prefs", prefs)
    opts.add_argument("--safebrowsing-disable-download-protection")
    driver = webdriver.Chrome(
        service=ChromeService(ChromeDriverManager().install()),
        options=opts
    )
    driver.maximize_window()
    return driver

def _admin_login(driver):
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    driver.get(ADMIN_LOGIN_URL)
    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.NAME, "email")))
    driver.find_element(By.NAME, "email").clear()
    driver.find_element(By.NAME, "email").send_keys(ADMIN_EMAIL)
    driver.find_element(By.NAME, "password").clear()
    driver.find_element(By.NAME, "password").send_keys(ADMIN_PASSWORD)
    driver.find_element(By.XPATH, "//button[@type='submit']").click()
    WebDriverWait(driver, 20).until(
        EC.invisibility_of_element_located((By.NAME, "email"))
    )
    print("→ Logged into admin panel.")

# ── Fetch order numbers and trigger downloads ─────────────────────────────────
def _admin_process_orders(driver, records):
    """
    For each PM record, search the admin orders page for the PO,
    grab the order number, and click the download links.
    Mirrors ShoptoPM.process_orders exactly.
    """
    driver.get(ORDERS_URL)

    for rec in records:
        po = rec["PO"]
        print(f"\n  Admin: searching for PO {po} …")

        search = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, "search"))
        )
        search.clear()
        search.send_keys(po, Keys.RETURN)
        time.sleep(2)

        try:
            table_row = WebDriverWait(driver, 12).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr"))
            )
        except TimeoutException:
            print(f"  ⚠️  No result row found for PO {po} — skipping.")
            rec["order_num"] = ""
            continue

        # Extract numeric order number from first cell
        order_num = ""
        try:
            order_p = table_row.find_element(By.XPATH, ".//td[1]//p")
            text = order_p.text.strip()
            if not text.isdigit():
                candidates = [
                    e.text.strip()
                    for e in table_row.find_elements(
                        By.XPATH, ".//td[1]//*[self::p|self::span|self::div]"
                    )
                    if e.text and e.text.strip()
                ]
                text = next((t for t in candidates if t.isdigit()), "")
            order_num = text
        except Exception:
            pass

        if not order_num:
            print(f"  ⚠️  Could not extract order number for PO {po} — skipping downloads.")
            rec["order_num"] = ""
            continue

        rec["order_num"] = order_num
        print(f"  ✓ PO {po} → admin order #{order_num}")

        # Determine which download links to click
        titles  = ["Download XML"]
        vendors = set(rec.get("vendors") or [])

        # Propper + Ariat together → skip vendor XLS to avoid wrong export
        if "propper" in vendors and "ariat" in vendors:
            print(f"  ⚠️  PO {po}: both Propper and Ariat present — skipping vendor XLS.")
            vendors.discard("propper")
            vendors.discard("ariat")

        if "wrangler" in vendors:
            titles.append("Download Wrangler")
        if "propper" in vendors:
            titles.append("Download Propper")
        if "ariat" in vendors:
            titles.append("Download Ariat/Carhartt")

        for title in titles:
            for attempt in range(3):
                try:
                    row = driver.find_element(
                        By.XPATH,
                        f"//tr[.//p[normalize-space()='{order_num}'] "
                        f"or .//a[normalize-space()='{order_num}']]"
                    )
                    link = row.find_element(By.XPATH, f".//a[@title='{title}']")
                    driver.execute_script(
                        "arguments[0].scrollIntoView({block:'center'});", link
                    )
                    driver.execute_script("arguments[0].click();", link)
                    time.sleep(1)
                    print(f"    ✓ Clicked '{title}'")
                    break
                except StaleElementReferenceException:
                    time.sleep(1)
                except Exception:
                    if attempt == 2:
                        print(f"    ⚠️  Could not click '{title}' for order {order_num} — skipping.")
                    time.sleep(1)

        print(f"  ✓ Downloads done for order {order_num}. Waiting 5 s …")
        time.sleep(5)

# ── Append rows to Processed_orders.xlsx ─────────────────────────────────────
def _write_pm_rows(records):
    """
    Appends one row per PM record to Processed_orders.xlsx,
    creating the file from Example.xlsx template if it doesn't exist yet.
    """
    from openpyxl import load_workbook
    from datetime import datetime

    today = datetime.now().strftime("%m/%d/%Y")

    if os.path.exists(OUTPUT_XLSX):
        wb = load_workbook(OUTPUT_XLSX)
    elif os.path.exists(TEMPLATE_XLSX):
        wb = load_workbook(TEMPLATE_XLSX)
        print(f"  Created {OUTPUT_XLSX} from template.")
    else:
        # No template — build a minimal workbook with the known headers
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.append([
            "Name", "Customer", "Ack", "Client PO #", "Cust Acct",
            "Who began order", "BMI Order #/ Full retailers PO",
            "Who finalized order", "Date PO finalized",
            "Notes/F/up date & who", "Vendor", "Transaction ID",
            "Order ID", "GP%", "Item Amount", "Freight (CC or N30)"
        ])
        print("  ⚠️  Example.xlsx not found — created bare workbook with default headers.")

    ws      = wb.active
    headers = [cell.value for cell in ws[1]]

    for rec in records:
        row_map = {
            "Name":                              today,
            "Customer":                          "The Sourcing Group",
            "Ack":                               rec["email"],
            "Client PO #":                       rec["PO"],
            "Cust Acct":                         "",
            "Who began order":                   PM_INITIALS,
            "BMI Order #/ Full retailers PO":    "",
            "Who finalized order":               PM_INITIALS,
            "Date PO finalized":                 today,
            "Notes/F/up date & who":             f"Order #: {rec.get('order_num', '')}",
            "Vendor":                            " / ".join(
                                                     rec.get("vendors") or ["Wrangler"]
                                                 ).title(),
            "Transaction ID":                    "Terms",
            "Order ID":                          "",
            "GP%":                               "19%",
            "Item Amount":                       rec["Order-Cost"],
            "Freight (CC or N30)":               "Cust Acct",
        }
        ws.append([row_map.get(h, "") for h in headers])
        print(f"  📝 Wrote row: PO={rec['PO']}, order#={rec.get('order_num','')}")

    wb.save(OUTPUT_XLSX)
    print(f"  ✓ Saved {len(records)} row(s) to {OUTPUT_XLSX}")

# ── Entry point called from main() ───────────────────────────────────────────
def run_pm_pipeline(matched_csvs):
    """
    matched_csvs: the same list of (csv_path, acct, po) used by the order placer.
    Logs into admin, fetches order numbers, triggers downloads,
    then appends rows to Processed_orders.xlsx.
    """
    if not matched_csvs:
        print("PM pipeline: no matched CSVs — nothing to log.")
        return

    print("\n" + "═" * 60)
    print("  PM LOGGING PIPELINE — starting …")
    print("═" * 60)

    records = _build_pm_records(matched_csvs)
    if not records:
        print("PM pipeline: no records built — skipping.")
        return

    admin_driver = None
    try:
        admin_driver = _setup_admin_driver()
        _admin_login(admin_driver)
        _admin_process_orders(admin_driver, records)
    except Exception as e:
        print(f"⚠️  PM pipeline admin step failed: {e}")
    finally:
        if admin_driver:
            try:
                admin_driver.quit()
            except Exception:
                pass

    _write_pm_rows(records)
    print("═" * 60)
    print("  PM LOGGING PIPELINE — complete.")
    print("═" * 60 + "\n")


# ─── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    # ── Load skipped POs from the Excel file ───────────────────────────────
    skipped_pos = load_skipped_pos()
    if not skipped_pos:
        print("No skipped orders to process. Exiting.")
        return
    print(f"📋 Found {len(skipped_pos)} skipped PO(s): {', '.join(sorted(skipped_pos))}")

    # ── Find CSVs and match them to skipped POs ────────────────────────────
    all_csvs = discover_csvs_with_accounts()
    if not all_csvs:
        print("No CSV files found in ./pdfs or script directory.")
        return

    # Build a list of (csv_path, account, po) only for CSVs whose PO was skipped
    creds_lower = {k.lower(): v for k, v in CREDENTIALS.items()}
    matched = []
    for csv_path, acct in all_csvs:
        try:
            df = pd.read_csv(csv_path)
            df = df.rename(columns={'po': 'PO', 'productId': 'Item-Number',
                                    'size1': 'Size-1', 'size2': 'Size-2', 'qty': 'Qty'})
            po = str(df.iloc[0]["PO"]).strip()
            if po in skipped_pos:
                matched.append((csv_path, acct, po))
        except Exception as e:
            print(f"⚠️  Could not read {os.path.basename(csv_path)}: {e}")

    if not matched:
        print("None of the CSV files match the skipped POs. Exiting.")
        return
    print(f"→ {len(matched)} CSV(s) matched to skipped POs.")

    # ── Phase 1: Place the back-order orders ──────────────────────────────
    driver = None
    current_account = None
    successfully_placed = []   # track which ones actually went through
    try:
        for csv_path, acct, po in matched:
            acct_norm = (acct or "").strip().lower()
            if acct_norm not in creds_lower:
                print(f"⚠️  {os.path.basename(csv_path)}: unknown account '{acct}'. Skipping.")
                continue

            # Switch account → restart browser
            if current_account != acct_norm:
                if driver:
                    try: driver.quit()
                    except Exception: pass
                driver = init_driver()
                ok = login(driver, acct_norm, creds_lower[acct_norm])
                if not ok:
                    print(f"✖ Login failed for {acct_norm}. Skipping.")
                    current_account = None
                    continue
                current_account = acct_norm
                print(f"→ New session for account {current_account}")

            print(f"→ Using account {current_account} for PO {po} ({os.path.basename(csv_path)})")
            process_backorder_csv(driver, csv_path)
            successfully_placed.append((csv_path, acct, po))

    finally:
        if driver:
            try: driver.quit()
            except Exception: pass

    # ── Phase 2: PM logging for every order we just placed ────────────────
    if successfully_placed:
        run_pm_pipeline(successfully_placed)
    else:
        print("No orders were successfully placed — skipping PM logging.")


if __name__ == "__main__":
    main()
