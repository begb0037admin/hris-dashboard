"""
login.py
HRIS Dashboard — One-time manual login
---------------------------------------
Run this ONCE to authenticate and save your full browser session.
A visible browser window will open. Log in manually (Oxford SSO -> Microsoft -> MFA).
Session is saved to session.json — generate_dashboard.py loads it automatically.

Re-run this only if the session expires (you'll see 401 errors in the dashboard log).

Run:
  cd C:/Users/admin/Documents/Claude/HRIS-Dashboard
  python login.py
"""

import json
from pathlib import Path
from playwright.sync_api import sync_playwright

SAASIT_URL    = "https://oxford.saasiteu.com"
SESSION_FILE  = "session.json"

def main():
    print("=" * 60)
    print("HRIS Dashboard — Manual Login")
    print("=" * 60)
    print()
    print("A browser window will open.")
    print("Log in with your Oxford SSO credentials (including MFA).")
    print("Once you can see the SAASIT dashboard, come back here")
    print("and press ENTER to save your session.")
    print()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=100)
        context = browser.new_context()
        page = context.new_page()

        print(f"Opening {SAASIT_URL} ...")
        page.goto(SAASIT_URL, timeout=60_000)

        input(">>> Log in fully in the browser window, then press ENTER here to save session: ")

        storage = context.storage_state()
        Path(SESSION_FILE).write_text(json.dumps(storage, indent=2), encoding="utf-8")

        cookie_count = len(storage.get("cookies", []))
        print()
        print(f"Session saved to {SESSION_FILE}  ({cookie_count} cookies stored)")
        print()
        print("You can now close this window and run generate_dashboard.py")

        browser.close()

if __name__ == "__main__":
    main()
