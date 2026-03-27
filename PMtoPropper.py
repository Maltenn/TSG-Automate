#!/usr/bin/env python3
# PMtoPropper.py
# ──────────────────────────────────────────────────────────────────────────────
# Automates order placement on retailer.propper.com
#
# Flow per order (rows in Processed_orders.xlsx where col K contains "Propper"):
#   1. Pre-flight  : Find & re-save the Propper upload CSV from the download folder
#   2. Login       : https://retailer.propper.com/customer/account/login/
#   3. Quick Order : Upload CSV → Add to Cart
#   4. Cart        : Verify cart → Proceed to Checkout
#   5. Shipping    : Click New Address → fill form → uncheck Save → Ship Here
#   6. Method      : Select FedEx Ground + account number → Next
#                    (retries automatically if an error popup kicks us back)
#   7. Payment     : Select Purchase Order → enter PO number from col G
#   8. PAUSE       : Wait for user review in app (Verification Complete / Enter)
#   9. Place Order : Submit
#
# Credentials come from the tsg_automate_app profile (env vars):
#   PROPPER_USERNAME  /  PROPPER_EMAIL
#   PROPPER_PASSWORD
# Paths come from:
#   TSG_WORKSPACE_DIR  (workspace folder – same dir as Processed_orders.xlsx)
#   TSG_DOWNLOAD_DIR   (folder where the browser saves downloads)
# ──────────────────────────────────────────────────────────────────────────────

import os
import sys
import re
import csv
import glob
import math
import time
import datetime
import traceback

import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    ElementClickInterceptedException,
    StaleElementReferenceException,
)

# ═══════════════════════════════════════════════════════════════════════════════
#  CONFIG
# ═══════════════════════════════════════════════════════════════════════════════

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

WORKSPACE_DIR   = os.getenv("TSG_WORKSPACE_DIR", SCRIPT_DIR)
DOWNLOAD_FOLDER = os.getenv("TSG_DOWNLOAD_DIR",
                             os.path.join(os.path.expanduser("~"), "Downloads"))
PDFS_DIR        = os.path.join(WORKSPACE_DIR, "pdfs")
EXCEL_PATH      = os.path.join(WORKSPACE_DIR, "Processed_orders.xlsx")

LOGIN_URL     = "https://retailer.propper.com/customer/account/login/"
QUICKORDER_URL = "https://retailer.propper.com/quickorder/"
CART_URL      = "https://retailer.propper.com/checkout/cart/"
CHECKOUT_URL  = "https://retailer.propper.com/checkout/"

USERNAME = (
    os.getenv("PROPPER_USERNAME")
    or os.getenv("PROPPER_EMAIL")
    or ""
)
PASSWORD = os.getenv("PROPPER_PASSWORD") or ""

PHONE_NUMBER    = "3309950736"
ACCOUNT_NUMBER  = "955617339"
SHIPPING_METHOD = "FXG"   # FedEx Ground option value in the carrier dropdown

WAIT_SHORT = 10
WAIT_LONG  = 25
WAIT_XLONG = 45

# ── Propper region_id values (matches the <select name="region_id"> in checkout) ─
STATE_TO_REGION_ID = {
    "AL": "1",   "AK": "2",   "AZ": "4",   "AR": "5",
    "CA": "12",  "CO": "13",  "CT": "14",  "DE": "15",  "DC": "16",
    "FL": "18",  "GA": "19",  "HI": "21",  "ID": "22",
    "IL": "23",  "IN": "24",  "IA": "25",  "KS": "26",
    "KY": "27",  "LA": "28",  "ME": "29",  "MD": "31",
    "MA": "32",  "MI": "33",  "MN": "34",  "MS": "35",
    "MO": "36",  "MT": "37",  "NE": "38",  "NV": "39",
    "NH": "40",  "NJ": "41",  "NM": "42",  "NY": "43",
    "NC": "44",  "ND": "45",  "OH": "47",  "OK": "48",
    "OR": "49",  "PA": "51",  "RI": "53",  "SC": "54",
    "SD": "55",  "TN": "56",  "TX": "57",  "UT": "58",
    "VT": "59",  "VA": "61",  "WA": "62",  "WV": "63",
    "WI": "64",  "WY": "65",
    # Territories
    "PR": "52",  "VI": "60",  "GU": "20",
}

# ═══════════════════════════════════════════════════════════════════════════════
#  UTILITY HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def log(msg: str) -> None:
    print(msg, flush=True)


def coerce_str(val) -> str:
    """Convert any pandas/Excel value to a clean string."""
    if val is None:
        return ""
    if isinstance(val, float):
        if math.isnan(val):
            return ""
        if val.is_integer():
            return str(int(val))
        return str(val)
    if isinstance(val, int):
        return str(val)
    return str(val).strip()


def safe_click(driver, el) -> None:
    """Click an element, falling back to JS click if intercepted."""
    try:
        el.click()
    except (ElementClickInterceptedException, StaleElementReferenceException):
        driver.execute_script("arguments[0].click();", el)


def clear_and_type(driver, el, text: str) -> None:
    """Clear a field fully (via JS + Selenium) then type text."""
    try:
        el.click()
    except Exception:
        pass
    try:
        el.clear()
    except Exception:
        pass
    driver.execute_script("arguments[0].value = '';", el)
    if text:
        el.send_keys(text)


def wait_for_url_change(driver, original_url: str, timeout: int = WAIT_LONG) -> None:
    end = time.time() + timeout
    while time.time() < end:
        if driver.current_url != original_url:
            return
        time.sleep(0.4)


# ═══════════════════════════════════════════════════════════════════════════════
#  CSV HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def resave_csv(path: str) -> None:
    """Re-open and re-save a CSV file with clean UTF-8 / QUOTE_MINIMAL formatting.

    Propper's quick-order uploader can choke on extra quoting that Excel or
    some download tools add.  This strips that and writes a clean copy.
    """
    log(f"[INFO] Re-saving CSV: {path}")
    rows = []
    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        try:
            with open(path, "r", encoding=enc, newline="") as fh:
                reader = csv.reader(fh)
                rows = list(reader)
            break
        except Exception:
            continue

    if not rows:
        log(f"[WARN] Could not read CSV for re-save — leaving as-is: {path}")
        return

    with open(path, "w", encoding="utf-8", newline="") as fh:
        writer = csv.writer(fh, quoting=csv.QUOTE_MINIMAL)
        for row in rows:
            writer.writerow(row)

    log(f"[OK] CSV re-saved: {path}")


def find_propper_csv(order_no: str) -> str:
    """Find the Propper upload CSV (sku/qty) for a given order number.

    Looks in DOWNLOAD_FOLDER for files matching patterns like:
      Order_No_<order_no>_propper.csv
      *<order_no>*propper*.csv
    Returns the most recently modified match, or "" if nothing found.
    """
    search_patterns = [
        os.path.join(DOWNLOAD_FOLDER, f"*{order_no}*propper*.csv"),
        os.path.join(DOWNLOAD_FOLDER, f"*propper*{order_no}*.csv"),
        os.path.join(DOWNLOAD_FOLDER, f"Order_No_{order_no}*.csv"),
        os.path.join(DOWNLOAD_FOLDER, f"*{order_no}*.csv"),
    ]

    for pattern in search_patterns:
        candidates = glob.glob(pattern)
        if candidates:
            # For the broad last pattern, require "propper" in the filename
            if pattern.endswith(f"*{order_no}*.csv"):
                candidates = [c for c in candidates
                              if "propper" in os.path.basename(c).lower()]
            if candidates:
                candidates.sort(key=lambda x: os.path.getmtime(x), reverse=True)
                return candidates[0]

    return ""


def load_address_from_csv(client_po: str) -> dict:
    """Load ship-to address from the PO CSV in PDFS_DIR.

    The CSV is named after the CLIENT PO number (column D in Processed_orders.xlsx),
    e.g. 301500.csv — NOT the order number from column J.

    Columns: email, PO, shipTo, productId, size1, size2, qty, unitCost, lineCost,
             orderCost, shipToCompany (K), shipToAttention (L), shipToStreet (M),
             shipToCity (N), shipToState (O), shipToZip (P)

    Returns a dict with keys: company, attention, street, city, state, zip
    """
    csv_path = os.path.join(PDFS_DIR, f"{client_po}.csv")
    if not os.path.isfile(csv_path):
        # Try wildcard fallback
        matches = glob.glob(os.path.join(PDFS_DIR, f"*{client_po}*.csv"))
        if matches:
            matches.sort(key=lambda x: os.path.getmtime(x), reverse=True)
            csv_path = matches[0]
        else:
            log(f"[WARN] No address CSV found for PO {client_po} in: {PDFS_DIR}")
            return {}

    log(f"[INFO] Loading address from: {csv_path}")
    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        try:
            with open(csv_path, "r", encoding=enc, newline="") as fh:
                reader = csv.DictReader(fh)
                for row in reader:
                    return {
                        "company":   row.get("shipToCompany",   "").strip(),
                        "attention": row.get("shipToAttention", "").strip(),
                        "street":    row.get("shipToStreet",    "").strip(),
                        "city":      row.get("shipToCity",      "").strip(),
                        "state":     row.get("shipToState",     "").strip().upper(),
                        "zip":       row.get("shipToZip",       "").strip(),
                    }
        except Exception as e:
            log(f"[WARN] Could not read {csv_path} ({enc}): {e}")

    log(f"[WARN] Address CSV appears empty: {csv_path}")
    return {}


# ═══════════════════════════════════════════════════════════════════════════════
#  BROWSER STEPS
# ═══════════════════════════════════════════════════════════════════════════════

def login(driver) -> None:
    log("[INFO] Navigating to Propper login page...")
    driver.get(LOGIN_URL)
    wait = WebDriverWait(driver, WAIT_LONG)

    email_el = wait.until(
        EC.presence_of_element_located((By.NAME, "login[username]"))
    )
    email_el.clear()
    email_el.send_keys(USERNAME)

    pass_el = driver.find_element(By.NAME, "login[password]")
    pass_el.clear()
    pass_el.send_keys(PASSWORD)

    # Sign In button: <button name="send">…</button>
    sign_in = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "button[name='send']"))
    )
    safe_click(driver, sign_in)

    # Wait until navigated away from login page
    WebDriverWait(driver, WAIT_XLONG).until(EC.url_changes(LOGIN_URL))
    log("[OK] Logged in to Propper.")


def upload_and_add_to_cart(driver, csv_path: str) -> None:
    """Navigate to the Quick Order page, upload the CSV, and click Add to Cart."""
    log("[INFO] Navigating to Quick Order page...")
    driver.get(QUICKORDER_URL)
    wait = WebDriverWait(driver, WAIT_LONG)

    # Locate the hidden file input and send the file path directly
    file_input = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "input#customer_sku_csv[type='file']")
        )
    )
    # Ensure Selenium can interact with the file input even if it's styled hidden
    driver.execute_script(
        "arguments[0].style.cssText = 'opacity:1; display:block; visibility:visible;';",
        file_input,
    )
    file_input.send_keys(os.path.abspath(csv_path))
    log(f"[INFO] CSV file sent to input: {csv_path}")

    # Wait for Add to Cart button to become enabled (isReady() == true)
    log("[INFO] Waiting for Add to Cart button to become enabled...")
    add_to_cart = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "button.tocart[title='Add to Cart']")
        )
    )
    # Poll until not disabled (Vue binding removes the disabled attr when ready)
    for _ in range(30):
        disabled = add_to_cart.get_attribute("disabled")
        if not disabled:
            break
        time.sleep(0.5)
    else:
        log("[WARN] Add to Cart button still appears disabled; attempting click anyway.")

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", add_to_cart)
    safe_click(driver, add_to_cart)
    log("[INFO] Clicked Add to Cart.")

    # Confirm we reached the cart page
    log("[INFO] Waiting for cart page...")
    WebDriverWait(driver, WAIT_XLONG).until(EC.url_contains("checkout/cart"))
    log("[OK] On cart page.")


def proceed_to_checkout(driver) -> None:
    """Click the Proceed to Checkout button from the cart page."""
    wait = WebDriverWait(driver, WAIT_LONG)
    checkout_link = wait.until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "a[title='Proceed to Checkout']")
        )
    )
    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});", checkout_link
    )
    safe_click(driver, checkout_link)
    log("[INFO] Clicked Proceed to Checkout.")

    # Wait for checkout URL (may end in #shipping, #payment, etc.)
    WebDriverWait(driver, WAIT_XLONG).until(
        EC.url_contains("retailer.propper.com/checkout")
    )
    time.sleep(2)
    log("[OK] On checkout page.")


def _ko_set_value(driver, el, value: str) -> None:
    """Type a value into a Knockout-bound input so the KO observable updates.

    Knockout uses valueUpdate:'keyup' — it reads el.value on every keyup event
    and stores it in the observable.  Setting el.value via JS + synthetic events
    does NOT update the observable; KO then overwrites the DOM back to empty the
    moment focus moves away.  The only reliable fix is real send_keys() which
    fires genuine keyboard events that update the KO observable directly.
    """
    from selenium.webdriver.common.action_chains import ActionChains
    from selenium.webdriver.common.keys import Keys

    # Focus the element
    try:
        driver.execute_script("arguments[0].focus();", el)
    except Exception:
        pass
    try:
        el.click()
    except Exception:
        pass

    # Select-all + Delete to clear existing content without el.clear()
    # (el.clear() can trigger KO to re-write empty string into the observable)
    try:
        ActionChains(driver).key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).perform()
        ActionChains(driver).send_keys(Keys.DELETE).perform()
    except Exception:
        try:
            el.clear()
        except Exception:
            pass

    # Type character by character — each keystroke fires a real keyup event
    # that Knockout's valueUpdate:'keyup' listener picks up and saves
    if value:
        el.send_keys(value)


def _wait_for_address_form(driver, timeout: int = WAIT_XLONG) -> None:
    """Wait until #opc-new-shipping-address is visible and firstname is interactive.

    Knockout controls visibility via isFormPopUpVisible() which toggles the
    element's inline style.  When visible the div has style="" (empty); when
    hidden it has style="display:none;".
    """
    log("[INFO] Waiting for address form to become visible...")
    end = time.time() + timeout

    while time.time() < end:
        try:
            container = driver.find_element(By.ID, "opc-new-shipping-address")
            style = container.get_attribute("style") or ""
            if "display:none" in style.replace(" ", "") or \
               "display: none" in style:
                time.sleep(0.4)
                continue

            # Container is visible — confirm firstname is usable
            try:
                fn = container.find_element(By.NAME, "firstname")
                if fn.is_displayed() and fn.is_enabled():
                    log("[INFO] Address form is open and ready.")
                    return
            except NoSuchElementException:
                pass
        except NoSuchElementException:
            pass

        time.sleep(0.4)

    raise TimeoutException(
        f"Address form (#opc-new-shipping-address) did not become interactive "
        f"within {timeout}s after clicking New Address."
    )


def fill_shipping_address(driver, addr: dict) -> None:
    """Click New Address, wait for the modal, fill all fields, click Ship Here."""
    wait = WebDriverWait(driver, WAIT_LONG)

    # ── Click "New Address" ───────────────────────────────────────────────────
    log("[INFO] Looking for New Address button...")
    try:
        new_addr_btn = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH,
                 "//*[normalize-space(text())='New Address' or "
                 "@data-bind=\"i18n: 'New Address'\"]")
            )
        )
        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});", new_addr_btn
        )
        safe_click(driver, new_addr_btn)
        log("[INFO] Clicked New Address.")
    except TimeoutException:
        log("[WARN] New Address button not found — form may already be visible.")

    # Small pause to let KO start animating the modal open
    time.sleep(1.0)

    # ── Wait for KO to make the form visible ──────────────────────────────────
    _wait_for_address_form(driver, timeout=WAIT_XLONG)

    # ── Scope all field lookups to #co-shipping-form ──────────────────────────
    # The form has stable id="co-shipping-form"; all inputs live inside it.
    def get_field(name_attr: str):
        """Return a visible, enabled input/select inside #co-shipping-form."""
        try:
            form = driver.find_element(By.ID, "co-shipping-form")
            els = form.find_elements(By.NAME, name_attr)
            for el in els:
                if el.is_displayed() and el.is_enabled():
                    return el
        except NoSuchElementException:
            pass

        # Broad fallback — any visible element with that name on the page
        for el in driver.find_elements(By.NAME, name_attr):
            if el.is_displayed() and el.is_enabled():
                return el

        raise NoSuchElementException(
            f"Could not find a visible+enabled field with name='{name_attr}' "
            "inside the shipping address form."
        )

    # ── First name → Company ──────────────────────────────────────────────────
    company_val = addr.get("company", "")
    fn_el = get_field("firstname")
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", fn_el)
    _ko_set_value(driver, fn_el, company_val)
    log(f"[INFO] firstname (Company): {company_val}")

    # ── Last name → Attention ─────────────────────────────────────────────────
    attention_val = addr.get("attention", "")
    ln_el = get_field("lastname")
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", ln_el)
    _ko_set_value(driver, ln_el, attention_val)
    log(f"[INFO] lastname (Attention): {attention_val}")

    # ── Street ────────────────────────────────────────────────────────────────
    street_val = addr.get("street", "")
    st_el = get_field("street[0]")
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", st_el)
    _ko_set_value(driver, st_el, street_val)
    log(f"[INFO] street: {street_val}")

    # ── Country — ensure United States is selected before state renders ───────
    try:
        country_el = get_field("country_id")
        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});", country_el
        )
        sel_country = Select(country_el)
        sel_country.select_by_value("US")
        driver.execute_script(
            "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
            country_el,
        )
        log("[INFO] country_id set to US")
        time.sleep(0.5)   # let KO re-render the state dropdown
    except Exception as e:
        log(f"[WARN] Could not set country: {e}")

    # ── State / region ────────────────────────────────────────────────────────
    state_abbrev = addr.get("state", "").upper()
    region_val = STATE_TO_REGION_ID.get(state_abbrev, "")
    if region_val:
        try:
            st_select = get_field("region_id")
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", st_select
            )
            sel_state = Select(st_select)
            sel_state.select_by_value(region_val)
            driver.execute_script(
                "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
                st_select,
            )
            log(f"[INFO] region_id: {state_abbrev} → {region_val}")
        except Exception as e:
            log(f"[WARN] Could not select state '{state_abbrev}': {e}")
    else:
        log(f"[WARN] Unknown state abbreviation '{state_abbrev}' — skipping.")

    # ── City ──────────────────────────────────────────────────────────────────
    city_val = addr.get("city", "")
    ct_el = get_field("city")
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", ct_el)
    _ko_set_value(driver, ct_el, city_val)
    log(f"[INFO] city: {city_val}")

    # ── ZIP ───────────────────────────────────────────────────────────────────
    zip_val = addr.get("zip", "")
    zp_el = get_field("postcode")
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", zp_el)
    _ko_set_value(driver, zp_el, zip_val)
    log(f"[INFO] postcode: {zip_val}")

    # ── Phone ─────────────────────────────────────────────────────────────────
    ph_el = get_field("telephone")
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", ph_el)
    _ko_set_value(driver, ph_el, PHONE_NUMBER)
    log(f"[INFO] telephone: {PHONE_NUMBER}")

    # ── Uncheck "Save in address book" ────────────────────────────────────────
    # The HTML shows: data-bind="visible: !isFormInline" — it may be hidden.
    try:
        save_cb = driver.find_element(By.ID, "shipping-save-in-address-book")
        if save_cb.is_displayed() and save_cb.is_selected():
            safe_click(driver, save_cb)
            log("[INFO] Unchecked 'Save in address book'.")
        else:
            log("[INFO] 'Save in address book' already unchecked or not visible.")
    except NoSuchElementException:
        log("[INFO] 'Save in address book' checkbox not in DOM — skipping.")

    # ── Ship Here ─────────────────────────────────────────────────────────────
    # The exact button from the page source:
    #   <button class="action primary action-save-address" type="button"
    #           data-role="action"><span>Ship Here</span></button>
    log("[INFO] Clicking Ship Here...")
    ship_here_btn = None

    for css in (
        "button.action.primary.action-save-address",
        "button.action-save-address",
    ):
        els = driver.find_elements(By.CSS_SELECTOR, css)
        for el in els:
            try:
                if el.is_displayed() and el.is_enabled():
                    ship_here_btn = el
                    break
            except StaleElementReferenceException:
                continue
        if ship_here_btn:
            break

    if not ship_here_btn:
        # XPath fallback by span text
        try:
            ship_here_btn = driver.find_element(
                By.XPATH,
                "//button[.//span[normalize-space(text())='Ship Here']]",
            )
        except NoSuchElementException:
            pass

    if not ship_here_btn:
        raise NoSuchElementException(
            "Could not locate the Ship Here button. "
            "Check that the address form is fully open."
        )

    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});", ship_here_btn
    )
    safe_click(driver, ship_here_btn)
    log("[INFO] Clicked Ship Here.")
    time.sleep(3)


def _wait_for_overlay_gone(driver, timeout: int = 20) -> None:
    """Wait for any Magento loading overlay to disappear."""
    end = time.time() + timeout
    while time.time() < end:
        overlays = driver.find_elements(
            By.CSS_SELECTOR,
            ".loading-mask, ._block-content-loading, "
            "[data-role='loader'], .loader"
        )
        visible = [o for o in overlays if o.is_displayed()]
        if not visible:
            return
        time.sleep(0.4)


def _do_fill_shipping_method(driver) -> None:
    """Internal: select FedEx Ground and enter account number (single attempt)."""
    wait = WebDriverWait(driver, WAIT_LONG)

    # Wait for any post-Ship-Here loading overlay to clear first
    _wait_for_overlay_gone(driver, timeout=20)

    # Select FedEx Ground from the carrier dropdown
    carrier_select_el = wait.until(
        EC.element_to_be_clickable((By.ID, "shippingnumber"))
    )
    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});", carrier_select_el
    )
    sel = Select(carrier_select_el)
    sel.select_by_value(SHIPPING_METHOD)
    log(f"[INFO] Selected carrier: FedEx Ground ({SHIPPING_METHOD})")
    time.sleep(1.0)

    # Wait for the account number input to be fully interactive
    # (it can be present in DOM but blocked by an overlay)
    acct_input = wait.until(
        EC.element_to_be_clickable(
            (By.ID, "propper_shippingnumber_shipping_number")
        )
    )
    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});", acct_input
    )
    # Use JS focus then clear, then send_keys — plain input (not KO-bound)
    driver.execute_script("arguments[0].focus();", acct_input)
    driver.execute_script("arguments[0].value = '';", acct_input)
    acct_input.send_keys(ACCOUNT_NUMBER)
    log(f"[INFO] Entered account number: {ACCOUNT_NUMBER}")


def fill_shipping_method_and_next(driver) -> None:
    """Fill the shipping method and click Next, retrying if an error popup
    kicks us back to the shipping step."""
    for attempt in range(1, 5):
        log(f"[INFO] Shipping method fill — attempt {attempt}...")
        try:
            _do_fill_shipping_method(driver)
        except Exception as e:
            log(f"[WARN] Could not fill shipping method (attempt {attempt}): {e}")
            if attempt < 4:
                time.sleep(2)
                continue
            else:
                log("[ERROR] Giving up on shipping method after 4 attempts.")
                return

        # Click Next
        try:
            wait = WebDriverWait(driver, WAIT_LONG)
            next_btn = wait.until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "button[data-role='opc-continue']")
                )
            )
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", next_btn
            )
            safe_click(driver, next_btn)
            log("[INFO] Clicked Next.")
        except TimeoutException:
            log("[WARN] Next button not found.")
            return

        # Wait a moment then check if we moved past the shipping step
        time.sleep(3)

        # If the carrier dropdown is still visible, an error occurred and we
        # were kicked back.  Loop and retry.
        try:
            WebDriverWait(driver, 4).until(
                EC.presence_of_element_located((By.ID, "shippingnumber"))
            )
            log("[WARN] Still on shipping step after clicking Next — retrying (error popup?).")
        except TimeoutException:
            # Carrier dropdown gone → we advanced to the next step
            log("[OK] Moved past shipping step.")
            return

    log("[ERROR] Could not advance past the shipping method step after multiple attempts.")


def fill_payment(driver, po_number: str) -> None:
    """Select the Purchase Order payment method and enter the PO number."""
    wait = WebDriverWait(driver, WAIT_LONG)

    # Select the Purchase Order radio
    log("[INFO] Waiting for Purchase Order radio button...")
    po_radio = wait.until(
        EC.element_to_be_clickable((By.ID, "purchaseorder"))
    )
    if not po_radio.is_selected():
        safe_click(driver, po_radio)
        log("[INFO] Selected Purchase Order payment method.")
    time.sleep(1.5)

    # PO number field (Magento standard: payment[po_number])
    po_text = po_number.replace("-", " ")
    log(f"[INFO] Entering PO number: '{po_text}'")

    po_input = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR,
             "input[name='payment[po_number]'], "
             "input[id*='po_number'], "
             "input[id*='purchaseorder']")
        )
    )
    clear_and_type(driver, po_input, po_text)
    log(f"[OK] PO number entered: '{po_text}'")


def place_order(driver) -> None:
    """Click the Place Order button and wait for the confirmation page."""
    wait = WebDriverWait(driver, WAIT_LONG)
    log("[INFO] Clicking Place Order...")
    place_btn = wait.until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR,
             "button[data-role='review-save'][title='Place Order']")
        )
    )
    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});", place_btn
    )
    safe_click(driver, place_btn)
    log("[OK] Place Order clicked — waiting for confirmation...")
    time.sleep(5)


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    log("")
    log("=" * 60)
    log("*** PROPPER B2B ORDER AUTOMATION SCRIPT ***")
    log("=" * 60)
    log(f"Started:        {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log(f"Workspace:      {WORKSPACE_DIR}")
    log(f"Excel:          {EXCEL_PATH}")
    log(f"Download dir:   {DOWNLOAD_FOLDER}")
    log(f"PDFs/CSV dir:   {PDFS_DIR}")
    log("=" * 60)
    log("")

    # ── Validate credentials ───────────────────────────────────────────────────
    if not USERNAME or not PASSWORD:
        log("[ERROR] Propper credentials not set.")
        log("[ERROR] Go to Manage Profiles and fill in the Propper Email / Password fields.")
        sys.exit(1)
    log(f"[INFO] Using Propper account: {USERNAME}")

    # ── Load and filter Processed_orders.xlsx ─────────────────────────────────
    if not os.path.isfile(EXCEL_PATH):
        log(f"[ERROR] Processed_orders.xlsx not found at: {EXCEL_PATH}")
        sys.exit(1)

    log("[INFO] Loading Processed_orders.xlsx...")
    df = pd.read_excel(EXCEL_PATH, engine="openpyxl", dtype=str)

    # Column indices (0-based):
    #   D = 3  → Client PO number (used to find address CSV in pdfs/, e.g. 301500.csv)
    #   G = 6  → PO / draft name  (used as the purchase-order reference on checkout)
    #   J = 9  → order number     (used to find upload CSV in downloads folder)
    #   K = 10 → vendor name      (filter: must contain "Propper")
    COL_D = df.columns[3]
    COL_G = df.columns[6]
    COL_J = df.columns[9]
    COL_K = df.columns[10]

    propper_rows = []
    for idx, row in df.iterrows():
        vendor = coerce_str(row[COL_K])
        if "propper" in vendor.lower():
            propper_rows.append((idx, row))

    if not propper_rows:
        log("[INFO] No rows with 'Propper' in Column K. Nothing to process.")
        sys.exit(0)

    log(f"[INFO] Found {len(propper_rows)} Propper order(s) to process.")

    # ── Pre-flight: locate + re-save every Propper upload CSV ─────────────────
    # This happens BEFORE the browser opens.
    log("")
    log("[INFO] Pre-flight: locating and re-saving Propper CSVs...")
    order_data = []   # list of (idx, row, order_no, csv_path)

    for (idx, row) in propper_rows:
        order_field = coerce_str(row[COL_J])
        m = re.search(r"\d+", order_field)
        if not m:
            log(f"[WARN] Cannot parse order number from Col J value '{order_field}' "
                f"at row {idx + 1} — skipping.")
            continue
        order_no = m.group()

        csv_path = find_propper_csv(order_no)
        if not csv_path:
            log(f"[WARN] No Propper upload CSV found for order {order_no} "
                f"in {DOWNLOAD_FOLDER} — skipping row {idx + 1}.")
            continue

        resave_csv(csv_path)
        order_data.append((idx, row, order_no, csv_path))
        log(f"  Order {order_no}: {os.path.basename(csv_path)}")

    if not order_data:
        log("[ERROR] No valid Propper orders with CSVs found. Exiting.")
        sys.exit(1)

    log(f"[INFO] Pre-flight complete — {len(order_data)} order(s) ready.")
    log("")

    # ── Browser automation ─────────────────────────────────────────────────────
    driver = webdriver.Chrome()
    try:
        # Login once for the session
        login(driver)

        total = len(order_data)
        for order_num, (idx, row, order_no, csv_path) in enumerate(order_data, 1):
            po_number  = coerce_str(row[COL_G])
            client_po  = coerce_str(row[COL_D])

            log("")
            log("─" * 60)
            log(f"[{order_num}/{total}] Processing order:")
            log(f"  PO number : {po_number}")
            log(f"  Client PO : {client_po}")
            log(f"  Order No  : {order_no}")
            log(f"  CSV       : {os.path.basename(csv_path)}")
            log("─" * 60)

            # Load shipping address from PDFS_DIR/{client_po}.csv  (e.g. 301500.csv)
            addr = load_address_from_csv(client_po)
            if addr:
                log(f"[INFO] Ship to: {addr.get('company','?')} | "
                    f"{addr.get('city','?')}, {addr.get('state','?')} {addr.get('zip','?')}")
            else:
                log("[WARN] No address data found — fields may be left blank.")

            # ── Step 1: Upload CSV + Add to Cart ──────────────────────────────
            upload_and_add_to_cart(driver, csv_path)

            # ── Step 2: Proceed to Checkout ───────────────────────────────────
            proceed_to_checkout(driver)

            # ── Step 3: Fill shipping address ─────────────────────────────────
            fill_shipping_address(driver, addr)

            # ── Step 4: Select FedEx Ground + account number + Next ───────────
            fill_shipping_method_and_next(driver)

            # ── Step 5: Fill payment (Purchase Order) ─────────────────────────
            fill_payment(driver, po_number)

            # ── Step 6: PAUSE for user review ─────────────────────────────────
            log("")
            log("╔" + "═" * 56 + "╗")
            log("║  ⏸  REVIEW REQUIRED — PROPPER ORDER                     ║")
            log("╠" + "═" * 56 + "╣")
            log(f"║  PO Number : {po_number:<42} ║")
            log(f"║  Ship To   : {addr.get('company','?'):<42} ║")
            log(f"║  City/St   : {(addr.get('city','?') + ', ' + addr.get('state','?')):<42} ║")
            log("╠" + "═" * 56 + "╣")
            log("║  Review the cart + payment in the browser, then:        ║")
            log("║  • In the app → click '✅ Verification Complete'        ║")
            log("║  • Running standalone → press Enter                     ║")
            log("╚" + "═" * 56 + "╝")
            input("")   # Blocks until Enter is sent (app sends '\n' via stdin)
            log("[INFO] User confirmed — placing order...")

            # ── Step 7: Place Order ───────────────────────────────────────────
            place_order(driver)
            log(f"[OK] Order placed for PO '{po_number}'!")

            # Short pause between orders so the site can settle
            if order_num < total:
                time.sleep(3)

        log("")
        log("=" * 60)
        log("*** PROPPER AUTOMATION COMPLETE ***")
        log(f"    {total} order(s) processed.")
        log("=" * 60)

    except Exception as e:
        log("")
        log("=" * 60)
        log("[ERROR] Script encountered an unexpected error!")
        log(f"        {e}")
        log("=" * 60)
        traceback.print_exc()
        raise

    finally:
        try:
            driver.quit()
        except Exception:
            pass
        log("[INFO] Browser closed.")


if __name__ == "__main__":
    main()
