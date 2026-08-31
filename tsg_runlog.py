"""Shared placement ledger for the TSG vendor scripts (added 2026-08-31).

The ledger (placed_orders.json in the workspace) records every successfully
placed order as  "<PO>|<vendor>": {order_id, when}.  It is the checkpoint the
scripts consult on start-up so a re-run after a crash skips exactly the orders
that were already placed — including split orders (one sheet row, several
vendors), which a plain column-M check cannot distinguish.

Column M in Processed_orders.xlsx remains the human-facing record; this file
is the machine-facing one.  Safe to delete once a batch is fully processed —
scripts treat a missing ledger as empty.
"""
import datetime
import json
import os


def _ledger_path(workspace: str) -> str:
    return os.path.join(workspace, "placed_orders.json")


def load_ledger(workspace: str) -> dict:
    try:
        with open(_ledger_path(workspace), "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def record_placed(workspace: str, po: str, vendor: str, order_id: str = "") -> None:
    """Mark PO as placed with this vendor. Never raises (a ledger hiccup must
    not fail an already-placed order)."""
    try:
        led = load_ledger(workspace)
        led[f"{str(po).strip()}|{vendor.strip().lower()}"] = {
            "order_id": str(order_id or ""),
            "when": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        tmp = _ledger_path(workspace) + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(led, fh, indent=2)
        os.replace(tmp, _ledger_path(workspace))
    except Exception as e:
        print(f"[WARN] Could not update placement ledger: {e}")


def already_placed(workspace: str, po: str, vendor: str):
    """Return the ledger entry if PO was already placed with vendor, else None."""
    return load_ledger(workspace).get(f"{str(po).strip()}|{vendor.strip().lower()}")
