#!/usr/bin/env python3
"""
PDFExtract.py (Improved)

Extracts line items + Ship To address from The Sourcing Group PO PDFs and writes a CSV per PO.

Output columns:
  email, PO, shipTo, productId, size1, size2, qty, unitCost, lineCost, orderCost

Notes
- Uses pdfplumber word coordinates to reliably capture the Ship To block even when the
  left/right columns are merged in plain text extraction.
- Handles both "PO Form Group#####" and "PO wInstructions#####" layouts.
- IMPROVED: Added fail-safes for malformed addresses (duplicate city/state/zip, standalone zips, etc.)
"""

import os
import re
import sys
import csv
from pathlib import Path

import pdfplumber

# Map "Our Contact:" -> email
CONTACT_EMAILS = {
    "Jessica McCarthy": "jmccarthy@thesourcinggroup.com",
    "Alvina Shotwell": "ashotwell@thesourcinggroup.com",
    "Mavi Delgado": "mdelgado@thesourcinggroup.com",
}

SKU_TOKEN_RE = re.compile(r"^[A-Z0-9][A-Z0-9\-]{4,}$", re.I)
FLOAT_RE = re.compile(r"^\d+\.\d{2}$")

ALPHA_SIZES = {
    "XXS", "XS", "S", "SM", "M", "MD", "L", "LG", "XL", "XXL", "2XL", "3XL", "4XL", "5XL", "6XL",
    "REG", "TALL", "LONG", "SHORT"
}

FOOTER_RE = re.compile(r"^(Subtotal|Total|Shipping|Sales\s+Tax|Authorized By:|Report Date:|Page\s+#)", re.I)

# Stationary carrier text to append to the end of Column C (shipTo)
FEDEX_SUFFIX = "FedEx Ground: 955617339"


def read_lines(pdf_path: Path):
    """Extract all text lines from the PDF."""
    out = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            for l in txt.splitlines():
                l = l.strip()
                if l:
                    out.append(l)
    return out


def extract_po(lines, filename: str) -> str:
    # Often the first line is the PO number
    for l in lines[:10]:
        m = re.search(r"\b(\d{6})\b", l)
        if m:
            return m.group(1)
    # Fallback: from filename
    m = re.search(r"(\d{6})", filename)
    return m.group(1) if m else ""


def extract_contact_email(lines) -> str:
    for l in lines:
        m = re.search(r"Our Contact:\s*(.+)$", l, re.I)
        if m:
            name = m.group(1).strip()
            return CONTACT_EMAILS.get(name, "")
    return ""


def _cluster_lines(words, y_tol: float = 2.5):
    """Cluster pdfplumber word dicts into lines using their 'top' coordinate."""
    words_sorted = sorted(words, key=lambda w: (w["top"], w["x0"]))
    clusters = []
    for w in words_sorted:
        placed = False
        for cl in clusters:
            if abs(cl[0]["top"] - w["top"]) <= y_tol:
                cl.append(w)
                placed = True
                break
        if not placed:
            clusters.append([w])
    return clusters


def extract_ship_to_lines(pdf_path: Path, lines):
    """
    Extract the Ship To block as a list of lines (best effort).

    Primary: coordinate-based extraction.
    Fallback: text-based extraction (less reliable).
    """
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            page = pdf.pages[0]
            words = page.extract_words()
            if not words:
                raise RuntimeError("No words extracted")

            # Find the line that contains both "Ship" and "To:".
            candidates = []
            for w in words:
                if w["text"] != "Ship":
                    continue
                same_line = [x for x in words if abs(x["top"] - w["top"]) <= 2.5]
                to_words = [x for x in same_line if x["text"].startswith("To")]
                if to_words:
                    candidates.append((w, to_words[0]))

            if not candidates:
                raise RuntimeError("Could not find Ship To label")

            # Usually there is also "Req. Ship Date" higher up; choose the lowest candidate.
            ship_w, to_w = sorted(candidates, key=lambda t: t[0]["top"])[-1]

            ship_top = ship_w["top"]
            # Start extracting just to the right of the "Ship To:" label
            x_threshold = max(ship_w["x1"], to_w["x1"]) + 4

            # Stop when we hit the table header / next section
            end_top = None
            for w in words:
                if w["top"] <= ship_top + 3:
                    continue
                if w["text"] in {"Unit", "Qty", "Product", "ID", "Subtotal", "NOTE", "Note"}:
                    end_top = w["top"]
                    break
            if end_top is None:
                end_top = ship_top + 180  # safe fallback - increased from 140 to capture more lines

            # FIX: Relaxed the boundary constraints to capture all address lines
            # - Changed (end_top - 1.5) to (end_top + 2) to include lines closer to the table
            # - Changed x_threshold to (x_threshold - 5) to handle slight x-coordinate variations
            ship_words = [
                w for w in words
                if (ship_top - 1.5) <= w["top"] < (end_top + 2) and w["x0"] >= (x_threshold - 5)
            ]

            out_lines = []
            for line_words in _cluster_lines(ship_words):
                txt = " ".join(w["text"] for w in sorted(line_words, key=lambda z: z["x0"])).strip()
                if txt and txt != ".":
                    out_lines.append(txt)

            if out_lines:
                return out_lines

    except Exception:
        pass

    # Fallback (best-effort) from raw text
    joined = "\n".join(lines)
    m = re.search(r"Ship To:\s*(.+)", joined, re.I)
    return [m.group(1).strip()] if m else []


def extract_ship_to_address(pdf_path: Path, lines) -> str:
    """Backwards-compatible string version of the Ship To block."""
    out_lines = extract_ship_to_lines(pdf_path, lines)
    return " | ".join(out_lines) if out_lines else ""


_CITY_STATE_ZIP_RE = re.compile(r"^(.*?),\s*([A-Z]{2})\s+(\d{5}(?:-\d{4})?)\s*$")
_CITY_STATE_ZIP_RE_NOCOMMA = re.compile(r"^(.*?)\s+([A-Z]{2})\s+(\d{5}(?:-\d{4})?)\s*$")
# Patterns without anchors for searching within a line (not just matching the full line)
_CITY_STATE_ZIP_PATTERN = re.compile(r"(.*?),\s*([A-Z]{2})\s+(\d{5}(?:-\d{4})?)\s*$")
_CITY_STATE_ZIP_PATTERN_NOCOMMA = re.compile(r"(.*?)\s+([A-Z]{2})\s+(\d{5}(?:-\d{4})?)\s*$")
_COUNTRY_RE = re.compile(r"^(UNITED\s+STATES(?:\s+OF\s+AMERICA)?|USA)\s*$", re.I)
_STANDALONE_ZIP_RE = re.compile(r"^(\d{5}(?:-\d{4})?)$")


def filter_address_chars(text: str) -> str:
    """Remove any characters that are not numbers, letters, or spaces."""
    if not text:
        return ""
    # Keep only alphanumeric characters and spaces
    filtered = re.sub(r'[^A-Za-z0-9 ]', '', text)
    # Normalize multiple spaces to single space
    filtered = re.sub(r'\s+', ' ', filtered)
    return filtered.strip()


def _is_ship_to_junk_line(s: str) -> bool:
    """Filter out carrier/service/tracking artifacts that sometimes bleed into the Ship To block."""
    t = s.strip()
    if not t:
        return True
    if t in {".", "_", "-"}:
        return True
    if _COUNTRY_RE.match(t):
        return True

    low = t.lower()
    # Common bleed-through lines observed on these POs:
    # - "_FedEx Ground 955617339"
    # - "Service not mappedFedEx - Ground"
    if "service not mapped" in low:
        return True
    if any(k in low for k in ["fedex", "ups", "dhl", "usps"]):
        # If it's clearly a carrier/service line (often includes a tracking-like number)
        if re.search(r"\d{6,}", t) or any(k in low for k in ["ground", "air", "overnight", "2day", "express"]):
            return True

    # Pure tracking-ish line
    if re.fullmatch(r"[#_\-\s]*\d{6,}[#_\-\s]*", t):
        return True

    return False


def _normalize_address_lines(lines):
    """
    NEW: Clean up malformed address lines before parsing.
    
    Handles:
    1. Standalone zip code lines (e.g., "71225" on its own line)
    2. Duplicate city/state/zip lines
    3. Incomplete addresses that need zip codes merged
    
    Returns a cleaned list of address lines.
    """
    if not lines:
        return []
    
    cleaned = []
    i = 0
    
    while i < len(lines):
        current = lines[i].strip()
        
        # Skip if this is a standalone zip code AND there's a next line with full city/state/zip
        if _STANDALONE_ZIP_RE.match(current) and i + 1 < len(lines):
            next_line = lines[i + 1].strip()
            # Check if next line has a complete city/state/zip
            if _CITY_STATE_ZIP_RE.match(next_line) or _CITY_STATE_ZIP_RE_NOCOMMA.match(next_line):
                # Extract the zip from the next line
                m = _CITY_STATE_ZIP_RE.match(next_line) or _CITY_STATE_ZIP_RE_NOCOMMA.match(next_line)
                if m and m.group(3) == current:
                    # This standalone zip matches the zip in the next line - skip it
                    print(f"  [CLEANUP] Skipping duplicate standalone zip: {current}")
                    i += 1
                    continue
        
        # Check if this line looks like an incomplete address (has city/state but no zip)
        # Pattern 1: ends with ", CITYNAME ST" or " CITYNAME ST" but no 5-digit zip
        incomplete_pattern = re.match(r"^(.+?)(?:,\s*|\s+)([A-Z][A-Za-z\s]+?)\s+([A-Z]{2})$", current)
        if incomplete_pattern and i + 1 < len(lines):
            next_line = lines[i + 1].strip()
            # Check if the next line is just a zip code
            zip_match = _STANDALONE_ZIP_RE.match(next_line)
            if zip_match:
                # Merge them into a proper address
                merged = f"{current} {next_line}"
                print(f"  [CLEANUP] Merging incomplete address with standalone zip:")
                print(f"            '{current}' + '{next_line}' -> '{merged}'")
                cleaned.append(merged)
                i += 2  # Skip both lines
                continue
        
        # Pattern 2: ends with just a city name (no state), followed by standalone zip, then full city/state/zip
        # Example: "6206 BRYANT POND DR HOUSTON" + "77041" + "HOUSTON, TX 77041"
        # Strategy: Check if line contains a number (street address) and if next+1 line is complete city/state/zip
        # Then see if current line ends with the same city name
        if re.search(r'\d', current) and i + 2 < len(lines):
            next_line = lines[i + 1].strip()
            line_after = lines[i + 2].strip()
            
            # Check if next line is standalone zip
            zip_match = _STANDALONE_ZIP_RE.match(next_line)
            if zip_match:
                # Check if line after next is complete city/state/zip
                complete_match = _CITY_STATE_ZIP_RE.match(line_after) or _CITY_STATE_ZIP_RE_NOCOMMA.match(line_after)
                if complete_match:
                    complete_city = complete_match.group(1).strip()
                    complete_zip = complete_match.group(3).strip()
                    
                    # Check if current line ends with the city name and zips match
                    if (current.upper().endswith(complete_city.upper()) and 
                        zip_match.group(1) == complete_zip):
                        # This is an incomplete address - strip the city name from the end
                        street_only = current[:-(len(complete_city))].strip()
                        print(f"  [CLEANUP] Detected incomplete address ending with city name only:")
                        print(f"            Original: '{current}'")
                        print(f"            Street only: '{street_only}'")
                        print(f"            Skipping standalone zip: '{next_line}'")
                        print(f"            Keeping complete city/state/zip: '{line_after}'")
                        cleaned.append(street_only)
                        # Skip the standalone zip, but let the complete city/state/zip be processed normally
                        cleaned.append(line_after)
                        i += 3  # Skip all three lines
                        continue
        
        # Check for duplicate city/state/zip lines
        current_match = _CITY_STATE_ZIP_RE.match(current) or _CITY_STATE_ZIP_RE_NOCOMMA.match(current)
        if current_match and i + 1 < len(lines):
            next_line = lines[i + 1].strip()
            next_match = _CITY_STATE_ZIP_RE.match(next_line) or _CITY_STATE_ZIP_RE_NOCOMMA.match(next_line)
            
            if next_match:
                # We have two consecutive city/state/zip lines - check if they're duplicates
                if (current_match.group(1).strip().upper() == next_match.group(1).strip().upper() and
                    current_match.group(2) == next_match.group(2) and
                    current_match.group(3) == next_match.group(3)):
                    # Duplicate found - keep the first one, skip the second
                    print(f"  [CLEANUP] Removing duplicate city/state/zip: {next_line}")
                    cleaned.append(current)
                    i += 2  # Skip both, we already added the first
                    continue
        
        # Normal case - just add the line
        cleaned.append(current)
        i += 1
    
    return cleaned


def parse_ship_to_fields(ship_to_lines):
    """
    Convert Ship To lines into structured columns.

    Returns a dict with:
      shipToCompany, shipToAttention, shipToStreet, shipToCity, shipToState, shipToZip

    Heuristics:
    - Drop obvious country lines (UNITED STATES/USA) and carrier/service/tracking artifacts.
    - Identify the last line matching "City, ST 12345" (or without comma).
    - Street is the line immediately above the city/state/zip line.
    - Lines above street are treated as company and attention (company = first line, attention = second+).
    
    IMPROVED: Now calls _normalize_address_lines first to handle malformed addresses.
    """
    # First pass: remove junk lines
    cleaned_initial = []
    for raw in (ship_to_lines or []):
        t = re.sub(r"\s+", " ", str(raw)).strip()
        t = t.lstrip("._- ").strip()
        if _is_ship_to_junk_line(t):
            continue
        cleaned_initial.append(t)
    
    # NEW: Second pass - normalize malformed addresses
    cleaned = _normalize_address_lines(cleaned_initial)

    out = {
        "shipToCompany": "",
        "shipToAttention": "",
        "shipToStreet": "",
        "shipToCity": "",
        "shipToState": "",
        "shipToZip": "",
    }

    if not cleaned:
        return out

    # Find city/state/zip line from bottom
    city_idx = None
    city_m = None
    for i in range(len(cleaned) - 1, -1, -1):
        m = _CITY_STATE_ZIP_RE.match(cleaned[i]) or _CITY_STATE_ZIP_RE_NOCOMMA.match(cleaned[i])
        if m:
            city_idx = i
            city_m = m
            break

    if city_idx is None:
        # No structured city/state/zip; best effort: company only
        out["shipToCompany"] = cleaned[0]
        if len(cleaned) > 1:
            out["shipToStreet"] = cleaned[1]
        return out

    out["shipToCity"] = (city_m.group(1) or "").strip()
    out["shipToState"] = (city_m.group(2) or "").strip()
    out["shipToZip"] = (city_m.group(3) or "").strip()

    # Street is the line above the city line, if present
    if city_idx - 1 >= 0:
        street_line = cleaned[city_idx - 1].strip()
        
        # FAIL-SAFE: Check if the street line ALSO contains a city/state/zip at the end
        # This happens when an incomplete address gets merged with a zip, creating:
        #   "163 CURRY CREEK DRIVE CALHOUN LA 71225" (merged, should be street only)
        #   "CALHOUN, LA 71225" (duplicate city/state/zip line)
        # We want to strip "CALHOUN LA 71225" from the street to leave just "163 CURRY CREEK DRIVE"
        
        # Try to match city/state/zip pattern at the END of the street line
        # Use the patterns without ^ anchor so they can match the end of a longer line
        street_city_match = (_CITY_STATE_ZIP_PATTERN.search(street_line) or 
                           _CITY_STATE_ZIP_PATTERN_NOCOMMA.search(street_line))
        
        if street_city_match:
            # Check if this city/state/zip matches what we found in the official city/state/zip line
            street_city = street_city_match.group(1).strip()
            street_state = street_city_match.group(2).strip()
            street_zip = street_city_match.group(3).strip()
            
            # Compare just the city name part (group 1 from street might have street address prefix)
            # We want to check if the ENDING matches the official city/state/zip
            official_city = out["shipToCity"].upper()
            official_state = out["shipToState"].upper()
            official_zip = out["shipToZip"]
            
            # Check if the street ends with the same state and zip
            if (street_state == official_state and street_zip == official_zip):
                # Extract just the street part (everything before the city/state/zip)
                # The captured group 1 is everything before ", STATE ZIP"
                # If it contains the street AND city, we need to extract just the street
                
                # For pattern like "163 CURRY CREEK DRIVE, CALHOUN, LA 71225"
                # group(1) would be "163 CURRY CREEK DRIVE, CALHOUN"
                # We need to check if this ends with the city name and remove it
                prefix = street_city_match.group(1).strip()
                
                # Check if prefix ends with the official city name
                if prefix.upper().endswith(official_city):
                    # Remove the city name from the end
                    clean_street = prefix[:-len(official_city)].strip().rstrip(',').strip()
                else:
                    # Just use the prefix as-is
                    clean_street = prefix.rstrip(',').strip()
                
                print(f"  [CLEANUP] Removing duplicate city/state/zip from street:")
                print(f"            Before: '{street_line}'")
                print(f"            After:  '{clean_street}'")
                out["shipToStreet"] = clean_street
            else:
                out["shipToStreet"] = street_line
        else:
            out["shipToStreet"] = street_line

    pre = cleaned[: max(city_idx - 1, 0)]
    if pre:
        out["shipToCompany"] = pre[0].strip()
        if len(pre) > 1:
            # If there are multiple "attention" lines, keep them together (pipe-separated)
            out["shipToAttention"] = " | ".join(x.strip() for x in pre[1:] if x.strip())

    return out


def _parse_item_line(line: str):
    """
    Parse the primary item line:
      "6 10FR47MLW Prewash Wrangler ... 6 60.45 362.70"
      "8 10030232 Field Ariat ... 8 39.65 317.20"
    """
    toks = line.split()
    if len(toks) < 5:
        return None
    if not toks[0].isdigit():
        return None

    qty = int(toks[0])
    pid = toks[1]
    if not SKU_TOKEN_RE.match(pid):
        return None

    # Find the last two float tokens (unitCost and lineCost) scanning from the end
    floats = []
    for k in range(len(toks) - 1, -1, -1):
        if FLOAT_RE.match(toks[k]):
            floats.append((k, float(toks[k])))
            if len(floats) == 2:
                break
    if len(floats) < 2:
        return None

    idx_line, line_cost = floats[0]   # last float in the line
    idx_unit, unit_cost = floats[1]   # second-last float

    # Normalize ordering (just in case)
    if idx_unit > idx_line:
        idx_unit, unit_cost, idx_line, line_cost = idx_line, line_cost, idx_unit, unit_cost

    color = toks[2] if len(toks) > 2 else ""
    return {
        "qty": qty,
        "productId": pid,
        "color": color,
        "unitCost": unit_cost,
        "lineCost": line_cost,
    }


def find_sizes(main_line: str, cont_lines, pid=""):
    """
    Pull sizes from the item line + continuation lines under an item.

    Fix: some POs put the waist on the *main* item line (e.g. "...-44") and the inseam on the next
    line (e.g. "30"). The old logic only looked at continuation lines, so it missed the waist.

    Strategy:
    - Look for explicit "44x30" patterns first.
    - Look for dash-separated numeric pairs ("- 34 32") next.
    - Otherwise, collect 2–3 digit integers that are *not* part of decimals (so 40.28 won't become 40 and 28),
      keep only typical waist/inseam ranges (20–60), and take the last two.
    - Fall back to alpha sizes (XS, SM, etc.).
    """
    text = " ".join([str(main_line)] + list(cont_lines or []))

    # 44x30 / 44 X 30
    m = re.search(r"\b(\d{2,3})\s*[xX]\s*(\d{2,3})\b", text)
    if m:
        return [m.group(1), m.group(2)]

    # "- 34 32"
    m = re.search(r"-\s*(\d{2,3})\s+(\d{2,3})\b", text)
    if m:
        return [m.group(1), m.group(2)]

    # "- 14 L" style: numeric size (any width) + alpha inseam/length (e.g. women's sizes)
    m = re.search(r"-\s*(\d{1,3})\s+([A-Z]{1,4})\b", text, re.I)
    if m and m.group(2).upper() in ALPHA_SIZES:
        return [m.group(1), m.group(2).upper()]

    # Integer tokens not adjacent to '.' so we don't split decimals like 40.28 into 40 and 28
    nums = re.findall(r"(?<![\d.])\d{2,3}(?![\d.])", text)
    nums = [n for n in nums if 20 <= int(n) <= 60]
    if nums:
        return nums[-2:] if len(nums) >= 2 else [nums[-1]]

    # Alpha size tokens — use (?<!') to avoid matching letters split from apostrophes
    # e.g. "Women's" → "WOMEN'S" would otherwise yield a spurious "S" token
    toks = re.findall(r"(?<!')\b[A-Z]{1,4}\b", text.upper())
    sizes = []
    seen = set()
    for t in toks:
        if t in ALPHA_SIZES and t not in seen:
            sizes.append(t)
            seen.add(t)
        if len(sizes) >= 2:
            break
    return sizes


def extract_products(lines):
    """
    Extract product rows using a header anchor, then parse item rows until footer.
    Works for both PO templates.
    """
    start = None
    for i, l in enumerate(lines):
        if re.search(r"\bQty\b.*\bProduct\b.*\bID\b", l):
            start = i + 1
            break
    if start is None:
        for i, l in enumerate(lines):
            if "Unit Total" in l:
                start = i + 1
                break
    if start is None:
        start = 0

    items = []
    i = start
    while i < len(lines):
        l = lines[i].strip()
        if FOOTER_RE.match(l):
            break

        parsed = _parse_item_line(l)
        if parsed:
            # Gather continuation lines until next item or footer
            cont = []
            j = i + 1
            while j < len(lines):
                nxt = lines[j].strip()
                if FOOTER_RE.match(nxt):
                    break
                if _parse_item_line(nxt):
                    break
                cont.append(nxt)
                j += 1

            sizes = find_sizes(l, cont, parsed["productId"])
            size1 = sizes[0] if len(sizes) > 0 else ""
            size2 = sizes[1] if len(sizes) > 1 else ""

            items.append({
                "productId": parsed["productId"],
                "qty": parsed["qty"],
                "size1": size1,
                "size2": size2,
                "unitCost": parsed["unitCost"],
                "lineCost": parsed["lineCost"],
            })
            i = j
        else:
            i += 1

    return items


def extract_order_total(lines):
    for key in ("Total", "Subtotal"):
        for l in lines[-60:]:
            m = re.search(rf"{key}\s+(\d+\.\d{{2}})", l, re.I)
            if m:
                return float(m.group(1))
    return None


def process_file(pdf_path: Path, out_dir: Path) -> bool:
    print(f"Processing {pdf_path.name}...")
    lines = read_lines(pdf_path)
    po = extract_po(lines, pdf_path.name)
    email = extract_contact_email(lines)

    ship_to_lines = extract_ship_to_lines(pdf_path, lines)
    ship_to = " | ".join(ship_to_lines) if ship_to_lines else ""

    # Ensure Column C ends with: "FedEx Ground: 955617339" (stationary)
    # - If the ship_to text already contains a carrier line with 955617339, remove it first
    ship_to_base = re.sub(r"\s*\|\s*[^|]*955617339[^|]*", "", ship_to, flags=re.I).strip(" |")
    ship_to_csv = f"{ship_to_base} | {FEDEX_SUFFIX}" if ship_to_base else FEDEX_SUFFIX

    ship_parts = parse_ship_to_fields(ship_to_lines)

    items = extract_products(lines)
    if not items:
        print(f"  No items in {pdf_path.name}")
        return False

    order_total = extract_order_total(lines)
    order_cost = order_total if order_total is not None else sum(i["lineCost"] for i in items)

    out_csv = out_dir / f"{po}.csv"
    headers = [
        "email", "PO", "shipTo", "productId", "size1", "size2", "qty",
        "unitCost", "lineCost", "orderCost",
        "shipToCompany", "shipToAttention", "shipToStreet", "shipToCity", "shipToState", "shipToZip"
    ]

    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(headers)
        for it in items:
            w.writerow([
                email,
                po,
                ship_to_csv,  # <-- Column C now always ends with the FedEx Ground line
                it["productId"],
                it["size1"],
                it["size2"],
                it["qty"],
                f"{it['unitCost']:.2f}",
                f"{it['lineCost']:.2f}",
                f"{order_cost:.2f}",
                filter_address_chars(ship_parts["shipToCompany"]),
                filter_address_chars(ship_parts["shipToAttention"]),
                filter_address_chars(ship_parts["shipToStreet"]),
                filter_address_chars(ship_parts["shipToCity"]),
                filter_address_chars(ship_parts["shipToState"]),
                filter_address_chars(ship_parts["shipToZip"]),
            ])

    print(f"  Wrote {out_csv}")
    return True


def process_path(target: str) -> bool:
    in_path = Path(target)
    out_dir = in_path if in_path.is_dir() else in_path.parent

    if in_path.is_dir():
        any_ok = False
        for pdf in sorted(in_path.glob("*.pdf")):
            any_ok |= process_file(pdf, out_dir)
        return any_ok
    else:
        return process_file(in_path, out_dir)


if __name__ == "__main__":
    # Usage:
    #   python PDFExtract.py <pdf_file_or_folder>
    # If no argument is given, it defaults to a sibling "pdfs" folder.
    target = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), "pdfs")
    print("▶︎ Running PDFExtract.py…")
    process_path(target)
