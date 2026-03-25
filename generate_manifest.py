"""
generate_manifest.py  –  TSG Automate manifest builder
=======================================================
Run this script from your SOURCE folder (the one you edit and maintain)
any time you update files before pushing them to GitHub / copying to the share.

Usage
-----
    python generate_manifest.py

It will write (or overwrite) update_manifest.json in the same folder.

Then push / copy update_manifest.json alongside all the .py files to your
chosen host (GitHub repo, network share, etc.).

Configuration
-------------
Edit the two variables below:

  BASE_URL   – The raw base URL where your files are hosted.
               For GitHub:  "https://raw.githubusercontent.com/Maltenn/TSG-Automate/main/update_manifest.json"
               For UNC:     r"\\server\share\TSG_Automate"

  FILE_NAMES – List every file you want the updater to manage.
               Add or remove entries as your project grows.
"""

import hashlib
import json
import os
import sys
from datetime import date

# ──────────────────────────────────────────────────────────────────────────────
#  CONFIGURE THESE TWO VALUES
# ──────────────────────────────────────────────────────────────────────────────

# Base URL where files are publicly reachable (no trailing slash).
# GitHub example:
BASE_URL = "https://raw.githubusercontent.com/YOUR_ORG/YOUR_REPO/main"

# All files the updater should manage (paths relative to this script's folder).
FILE_NAMES = [
    "tsg_automate_app.py",
    "app_updater.py",
    "PDFExtract.py",
    "BroberryShop.py",
    "BroberryShop_Backorders.py",
    "ShoptoPM.py",
    "Add_PM_Nums.py",
    "PMtoWRG.py",
    "PMtoARIAT.py",
    "PMtoPropper.py",
    "GetOrderId.py",
    # Add more files here as needed:
    # "SomeNewScript.py",
]

# ──────────────────────────────────────────────────────────────────────────────

APP_VERSION = "1.0.0"   # Bump this manually whenever you want a visible version label


def sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def main():
    src_dir = os.path.dirname(os.path.abspath(__file__))
    entries = []
    missing = []

    print(f"Scanning {src_dir} …\n")

    for name in FILE_NAMES:
        path = os.path.join(src_dir, name)
        if not os.path.isfile(path):
            print(f"  ⚠  MISSING – {name}")
            missing.append(name)
            continue
        digest = sha256(path)
        url = f"{BASE_URL}/{name}"
        entries.append({
            "name": name,
            "url": url,
            "sha256": digest,
        })
        print(f"  ✓  {name}  ({digest[:12]}…)")

    manifest = {
        "version": APP_VERSION,
        "updated": str(date.today()),
        "files": entries,
    }

    out_path = os.path.join(src_dir, "update_manifest.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2)

    print(f"\n✅ Wrote {out_path}")
    print(f"   Version : {APP_VERSION}")
    print(f"   Files   : {len(entries)} hashed  |  {len(missing)} missing")

    if missing:
        print(f"\n⚠  The following files were not found and were skipped:")
        for m in missing:
            print(f"     – {m}")
        print("   Create or add them to FILE_NAMES in this script.")

    print("\nNext steps:")
    print("  1. Upload all .py files + update_manifest.json to your host.")
    print(f"  2. Make sure the manifest URL in tsg_automate_app.py points to:")
    print(f"     {BASE_URL}/update_manifest.json")


if __name__ == "__main__":
    main()
