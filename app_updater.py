r"""
app_updater.py  -  TSG Automate self-updater
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
5.  Results are returned to the caller -- all dialogs are shown by the app
    on the main thread (never from here).

HOSTING OPTIONS (pick one, set MANIFEST_URL in tsg_automate_app.py)
--------------------------------------------------------------------
A) GitHub (recommended -- free, versioned, always accessible)
      MANIFEST_URL = "https://raw.githubusercontent.com/Maltenn/TSG-Automate/main/update_manifest.json"

B) Shared network drive (no internet required)
      MANIFEST_URL = "\\\\server\\share\\TSG_Automate\\update_manifest.json"
      File URLs in the manifest should also be UNC paths under that share.

GENERATE THE MANIFEST
---------------------
Run generate_manifest.py in your source folder any time you update files.
"""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import urllib.request
import urllib.error
from typing import Callable


def check_and_update(
    manifest_url: str,
    app_dir: str,
    log: Callable[[str], None],
) -> dict:
    """
    Fetch the manifest and update all changed/missing files.
    Never touches Qt -- all dialog logic is handled by the caller.

    Returns a dict with keys:
        error            : str | None
        version          : str
        updated          : list[str]
        added            : list[str]
        failed           : list[str]
        skipped          : list[str]
        main_app_updated : bool
    """
    result = dict(
        error=None,
        version="?",
        updated=[],
        added=[],
        failed=[],
        skipped=[],
        main_app_updated=False,
    )

    log("Checking for updates...")

    try:
        manifest = _fetch_json(manifest_url)
    except Exception as exc:
        msg = f"Could not reach update server: {exc}"
        log(f"[ERROR] {msg}")
        result["error"] = msg
        return result

    result["version"] = manifest.get("version", "?")
    files = manifest.get("files", [])

    if not files:
        log("[WARN] Manifest contains no files - nothing to do.")
        return result

    log(f"Manifest version {result['version']} - {len(files)} file(s) listed.")

    for entry in files:
        name        = entry.get("name", "")
        remote_url  = entry.get("url", "")
        remote_hash = entry.get("sha256", "").lower()

        if not name or not remote_url:
            log(f"[WARN] Skipping invalid manifest entry: {entry}")
            continue

        dest = os.path.join(app_dir, name)
        is_new = not os.path.isfile(dest)

        if not is_new:
            if remote_hash and _sha256(dest) == remote_hash:
                result["skipped"].append(name)
                log(f"   checkmark  {name}  (up-to-date)")
                continue

        log(f"   downloading  {name}  {'(new file)' if is_new else '(updating)'}...")
        try:
            _download_file(remote_url, dest)
        except Exception as exc:
            log(f"   [ERROR] Failed to download {name}: {exc}")
            result["failed"].append(name)
            continue

        if remote_hash and _sha256(dest) != remote_hash:
            log(f"   [ERROR] Hash mismatch for {name} - file may be corrupt.")
            result["failed"].append(name)
            continue

        if is_new:
            result["added"].append(name)
        else:
            result["updated"].append(name)

        if name.lower() == "tsg_automate_app.py":
            result["main_app_updated"] = True

    log("-" * 40)
    if result["updated"]:
        log(f"Updated ({len(result['updated'])}): {', '.join(result['updated'])}")
    if result["added"]:
        log(f"Added ({len(result['added'])}): {', '.join(result['added'])}")
    if result["skipped"]:
        log(f"Already current ({len(result['skipped'])}): {', '.join(result['skipped'])}")
    if result["failed"]:
        log(f"Failed ({len(result['failed'])}): {', '.join(result['failed'])}")
    if not result["updated"] and not result["added"]:
        log("Everything is already up-to-date.")

    return result


def _fetch_json(url: str) -> dict:
    if _is_local(url):
        with open(url, "r", encoding="utf-8") as f:
            return json.load(f)
    req = urllib.request.Request(url, headers={"User-Agent": "TSGAutomate-Updater/1.0"})
    with urllib.request.urlopen(req, timeout=15) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _download_file(url: str, dest: str) -> None:
    os.makedirs(os.path.dirname(dest) or ".", exist_ok=True)
    tmp = dest + ".tmp"
    try:
        if _is_local(url):
            shutil.copy2(url, tmp)
        else:
            req = urllib.request.Request(url, headers={"User-Agent": "TSGAutomate-Updater/1.0"})
            with urllib.request.urlopen(req, timeout=30) as resp, open(tmp, "wb") as out:
                shutil.copyfileobj(resp, out)
        if os.path.isfile(dest):
            os.replace(tmp, dest)
        else:
            os.rename(tmp, dest)
    except Exception:
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
    return not url.startswith(("http://", "https://", "ftp://"))
