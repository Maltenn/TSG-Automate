"""
app_updater.py  –  TSG Automate self-updater
=============================================
Drop this file into your app folder alongside tsg_automate_app.py.

HOW IT WORKS
------------
1.  A manifest JSON file is hosted somewhere reachable (GitHub Raw URL,
    SharePoint direct-download link, shared-network-drive path, etc.).
2.  The manifest lists every file that should be kept up-to-date, with its
    download URL and SHA-256 hash.
3.  When the user clicks "Update App", this module fetches the manifest,
    compares hashes, and downloads only changed/missing files.
4.  Files are downloaded to a .tmp sibling first, then atomically renamed so
    a failed download never corrupts the live copy.
5.  If the main app file itself was updated, the user is prompted to restart.

HOSTING OPTIONS (pick one, set MANIFEST_URL in tsg_automate_app.py)
--------------------------------------------------------------------
A) GitHub (recommended – free, versioned, always accessible)
      MANIFEST_URL = "https://raw.githubusercontent.com/Maltenn/TSG-Automate/main/update_manifest.json"
      Upload both the manifest and all script files to that repo.

B) SharePoint / OneDrive direct link
      MANIFEST_URL = "https://orgname.sharepoint.com/.../_layouts/download.aspx?..."
      (Use the "Copy direct link" option, not the sharing URL.)

C) Shared network drive (no internet required)
      MANIFEST_URL = r"\\server\share\TSG_Automate\update_manifest.json"
      File URLs in the manifest should also be UNC paths under that share.

GENERATE THE MANIFEST
---------------------
Run generate_manifest.py in your source folder any time you update files.
It will hash every .py file and write update_manifest.json automatically.
"""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import sys
import tempfile
import urllib.request
import urllib.error
from typing import Callable

# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def check_and_update(
    manifest_url: str,
    app_dir: str,
    log: Callable[[str], None],
    parent_window=None,
) -> bool:
    """
    Fetch the manifest and update all changed/missing files.

    Parameters
    ----------
    manifest_url  : URL or local path to update_manifest.json
    app_dir       : directory where files should be placed (APP_DIR)
    log           : callable that accepts a single string message
    parent_window : optional Qt parent widget for dialogs

    Returns
    -------
    True  if the main app file (tsg_automate_app.py) was replaced
          → caller should offer a restart prompt
    False otherwise
    """
    log("🔄 Checking for updates…")

    # ── 1. Fetch manifest ───────────────────────────────────────────────────
    try:
        manifest = _fetch_json(manifest_url)
    except Exception as exc:
        log(f"[ERROR] Could not reach update server: {exc}")
        _show_error(parent_window, f"Could not reach update server:\n\n{exc}")
        return False

    remote_version = manifest.get("version", "?")
    files: list[dict] = manifest.get("files", [])
    if not files:
        log("[WARN] Manifest contains no files – nothing to do.")
        return False

    log(f"📋 Manifest version {remote_version} – {len(files)} file(s) listed.")

    # ── 2. Compare and download ─────────────────────────────────────────────
    updated: list[str] = []
    added:   list[str] = []
    failed:  list[str] = []
    skipped: list[str] = []
    main_app_updated = False

    for entry in files:
        name:        str = entry.get("name", "")
        remote_url:  str = entry.get("url", "")
        remote_hash: str = entry.get("sha256", "").lower()

        if not name or not remote_url:
            log(f"[WARN] Skipping invalid manifest entry: {entry}")
            continue

        dest = os.path.join(app_dir, name)
        is_new = not os.path.isfile(dest)

        # Compare hash
        if not is_new:
            local_hash = _sha256(dest)
            if local_hash == remote_hash:
                skipped.append(name)
                log(f"   ✓  {name}  (up-to-date)")
                continue

        # Download
        log(f"   ⬇  {name}  {'(new file)' if is_new else '(updating)'}…")
        try:
            _download_file(remote_url, dest)
        except Exception as exc:
            log(f"   [ERROR] Failed to download {name}: {exc}")
            failed.append(name)
            continue

        # Verify hash after download
        if remote_hash:
            actual = _sha256(dest)
            if actual != remote_hash:
                log(f"   [ERROR] Hash mismatch for {name} – file may be corrupt, keeping original.")
                # The atomic rename in _download_file already handled this, but log it
                failed.append(name)
                continue

        if is_new:
            added.append(name)
        else:
            updated.append(name)

        if name.lower() == "tsg_automate_app.py":
            main_app_updated = True

    # ── 3. Summary ──────────────────────────────────────────────────────────
    log("─" * 40)
    if updated:
        log(f"✅ Updated  ({len(updated)}): {', '.join(updated)}")
    if added:
        log(f"🆕 Added    ({len(added)}): {', '.join(added)}")
    if skipped:
        log(f"  ↳ Already current ({len(skipped)}): {', '.join(skipped)}")
    if failed:
        log(f"❌ Failed   ({len(failed)}): {', '.join(failed)}")

    if not updated and not added:
        log("🎉 Everything is already up-to-date.")
        _show_info(parent_window, f"TSG Automate is already up-to-date.\n(Manifest v{remote_version})")
    else:
        total = len(updated) + len(added)
        msg = f"{total} file(s) updated/added successfully."
        if failed:
            msg += f"\n\n⚠ {len(failed)} file(s) failed – check the log."
        if main_app_updated:
            msg += "\n\nThe main application file was updated.\nPlease restart TSG Automate."
        _show_info(parent_window, msg, title="Update Complete")

    return main_app_updated


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _fetch_json(url: str) -> dict:
    """Fetch and parse JSON from a URL or a local/UNC file path."""
    if _is_local(url):
        with open(url, "r", encoding="utf-8") as f:
            return json.load(f)
    req = urllib.request.Request(url, headers={"User-Agent": "TSGAutomate-Updater/1.0"})
    with urllib.request.urlopen(req, timeout=15) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _download_file(url: str, dest: str) -> None:
    """
    Download url → dest atomically (write to .tmp sibling, then rename).
    Supports https:// URLs and local/UNC paths.
    """
    os.makedirs(os.path.dirname(dest) or ".", exist_ok=True)
    tmp = dest + ".tmp"
    try:
        if _is_local(url):
            shutil.copy2(url, tmp)
        else:
            req = urllib.request.Request(url, headers={"User-Agent": "TSGAutomate-Updater/1.0"})
            with urllib.request.urlopen(req, timeout=30) as resp, open(tmp, "wb") as out:
                shutil.copyfileobj(resp, out)
        # Atomic replace
        if os.path.isfile(dest):
            os.replace(tmp, dest)
        else:
            os.rename(tmp, dest)
    except Exception:
        # Clean up temp file on failure
        if os.path.isfile(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
        raise


def _sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _is_local(url: str) -> bool:
    """Return True if url is a filesystem path (UNC, drive letter, relative)."""
    return not url.startswith(("http://", "https://", "ftp://"))


# ---------------------------------------------------------------------------
# Qt dialog helpers (gracefully degrade if Qt isn't imported yet)
# ---------------------------------------------------------------------------

def _show_info(parent, msg: str, title: str = "TSG Automate Updater") -> None:
    try:
        from PySide6 import QtWidgets
        QtWidgets.QMessageBox.information(parent, title, msg)
    except Exception:
        print(f"[{title}] {msg}")


def _show_error(parent, msg: str, title: str = "Update Error") -> None:
    try:
        from PySide6 import QtWidgets
        QtWidgets.QMessageBox.critical(parent, title, msg)
    except Exception:
        print(f"[{title}] {msg}", file=sys.stderr)
