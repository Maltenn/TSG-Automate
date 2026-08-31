#!/usr/bin/env python3
"""Read-only live inspector for TSG_DEBUG vendor-script runs.

When a vendor script runs with TSG_DEBUG=1, its Chrome exposes DevTools on:
    9222 = PMtoARIAT   9223 = PMtoPropper   9224 = PMtoWRG

This tool attaches over the Chrome DevTools Protocol (pure websocket — it does
NOT create a second Selenium session, so it cannot disturb the automation) and
dumps the live page state: URL, title, readyState, a screenshot PNG and the
rendered DOM HTML.

Usage:
    python debug_inspect.py [port] [outdir]
    python debug_inspect.py 9222 --js "document.querySelectorAll('.mainMenu').length"

With --js, evaluates the (read-only) expression and prints the result instead
of dumping files.
"""
import base64
import datetime
import json
import os
import sys

import requests
import websocket


def main():
    args = sys.argv[1:]
    port = args[0] if args and not args[0].startswith("--") else "9222"
    js_expr = None
    outdir = os.getcwd()
    if "--js" in args:
        js_expr = args[args.index("--js") + 1]
    elif len(args) > 1:
        outdir = args[1]

    try:
        targets = requests.get(f"http://127.0.0.1:{port}/json", timeout=5).json()
    except Exception as e:
        print(f"[ERROR] Could not reach DevTools on port {port}: {e}")
        print("Is the vendor script running with TSG_DEBUG=1 ?")
        sys.exit(1)

    pages = [t for t in targets
             if t.get("type") == "page" and not t.get("url", "").startswith("devtools")]
    if not pages:
        print("[ERROR] No page targets found. Raw targets:")
        print(json.dumps(targets, indent=2)[:1000])
        sys.exit(1)

    for i, t in enumerate(pages):
        print(f"[{i}] {t.get('title', '')[:70]!r}  {t.get('url', '')[:140]}")

    page = pages[0]
    ws = websocket.create_connection(page["webSocketDebuggerUrl"],
                                     timeout=20, suppress_origin=True)
    state = {"id": 0}

    def cmd(method, **params):
        state["id"] += 1
        ws.send(json.dumps({"id": state["id"], "method": method, "params": params}))
        while True:
            msg = json.loads(ws.recv())
            if msg.get("id") == state["id"]:
                if "error" in msg:
                    raise RuntimeError(f"{method}: {msg['error']}")
                return msg.get("result", {})

    try:
        ready = cmd("Runtime.evaluate", expression="document.readyState",
                    returnByValue=True)["result"].get("value")
        print(f"URL:        {page.get('url')}")
        print(f"TITLE:      {page.get('title')}")
        print(f"readyState: {ready}")

        if js_expr:
            res = cmd("Runtime.evaluate", expression=js_expr, returnByValue=True)
            r = res.get("result", {})
            if "value" in r:
                val = r["value"]
                print("JS result:", json.dumps(val, indent=2, default=str)[:4000]
                      if isinstance(val, (dict, list)) else val)
            else:
                print("JS result (unserialized):", r.get("description", r))
            return

        ts = datetime.datetime.now().strftime("%H%M%S")
        shot = cmd("Page.captureScreenshot", format="png")
        png = os.path.join(outdir, f"live_{port}_{ts}.png")
        with open(png, "wb") as f:
            f.write(base64.b64decode(shot["data"]))
        dom = cmd("Runtime.evaluate",
                  expression="document.documentElement.outerHTML",
                  returnByValue=True)["result"].get("value", "")
        html = os.path.join(outdir, f"live_{port}_{ts}.html")
        with open(html, "w", encoding="utf-8") as f:
            f.write(dom)
        print(f"PNG:  {png}")
        print(f"HTML: {html}")
    finally:
        ws.close()


if __name__ == "__main__":
    main()
