"""
login.py
HRIS Dashboard — Manual login + GitHub secret sync
----------------------------------------------------
Run this on any machine (work laptop or home desktop) when your session expires.

What it does:
  1. Opens a browser window — log in with Oxford SSO (Microsoft MFA)
  2. Saves the session to session.json locally (so you can run the script locally too)
  3. Automatically pushes the session to GitHub as the SAASIT_SESSION secret
     so GitHub Actions can run the dashboard without any local files

Dependencies:
  pip install playwright PyNaCl requests
  playwright install chromium

Run:
  python login.py

You only need to do this when the dashboard stops updating (session expired).
Typically every few weeks.
"""

# ─────────────────────────────────────────────
#  CONFIGURATION — must match generate_dashboard.py
# ─────────────────────────────────────────────

# repo + workflow scope PAT. NEVER hardcode it here: GitHub auto-revokes any
# PAT pushed to a public repo. Set the HRIS_GITHUB_PAT env var, or put the
# token in a github_pat.txt file next to this script (git-ignored).
GITHUB_PAT   = __import__("os").environ.get("HRIS_GITHUB_PAT", "").strip()
if not GITHUB_PAT:
    _pat_file = __import__("pathlib").Path(__file__).with_name("github_pat.txt")
    if _pat_file.exists():
        GITHUB_PAT = _pat_file.read_text(encoding="utf-8").strip()

GITHUB_REPO  = "begb0037admin/hris-dashboard"
SECRET_NAME  = "SAASIT_SESSION"

SAASIT_URL   = "https://oxford.saasiteu.com"
SESSION_FILE = "session.json"

# ─────────────────────────────────────────────

import json
import sys
from base64 import b64encode
from pathlib import Path

import requests
from playwright.sync_api import sync_playwright


# ─────────────────────────────────────────────
#  GITHUB SECRETS API
# ─────────────────────────────────────────────

def push_session_to_github(session_data: dict) -> None:
    """
    Encrypts session_data and stores it as the SAASIT_SESSION GitHub Actions secret.
    Requires: pip install PyNaCl
    """
    try:
        from nacl import encoding, public
    except ImportError:
        print()
        print("⚠️  PyNaCl not installed — cannot push session to GitHub.")
        print("   Run:  pip install PyNaCl")
        print("   Then re-run login.py")
        print()
        print("   The session has been saved locally to session.json.")
        print("   You can still run the dashboard locally with: python generate_dashboard.py")
        return

    session_json = json.dumps(session_data)

    headers = {
        "Authorization": f"Bearer {GITHUB_PAT}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }

    print("Fetching GitHub repo public key ...")
    r = requests.get(
        f"https://api.github.com/repos/{GITHUB_REPO}/actions/secrets/public-key",
        headers=headers,
        timeout=15,
    )
    if r.status_code == 401:
        print()
        print("❌  GitHub auth failed (401). Check that GITHUB_PAT is correct")
        print("   and has 'repo' scope.")
        return
    r.raise_for_status()
    key_data = r.json()

    print("Encrypting session ...")
    pk  = public.PublicKey(key_data["key"].encode(), encoding.Base64Encoder())
    box = public.SealedBox(pk)
    encrypted_b64 = b64encode(box.encrypt(session_json.encode())).decode()

    print(f"Pushing to GitHub secret '{SECRET_NAME}' ...")
    r = requests.put(
        f"https://api.github.com/repos/{GITHUB_REPO}/actions/secrets/{SECRET_NAME}",
        headers=headers,
        json={"encrypted_value": encrypted_b64, "key_id": key_data["key_id"]},
        timeout=15,
    )
    r.raise_for_status()

    print(f"✅  Session pushed to GitHub secret '{SECRET_NAME}' successfully.")
    print("   GitHub Actions can now fetch the dashboard without any local files.")


# ─────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────

def main():
    print("=" * 60)
    print("HRIS Dashboard — Session Refresh")
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
        page    = context.new_page()

        print(f"Opening {SAASIT_URL} ...")
        page.goto(SAASIT_URL, timeout=60_000)

        input(">>> Log in fully in the browser, then press ENTER here to save: ")

        storage = context.storage_state()
        Path(SESSION_FILE).write_text(json.dumps(storage, indent=2), encoding="utf-8")

        cookie_count = len(storage.get("cookies", []))
        print()
        print(f"✅  Session saved locally to {SESSION_FILE}  ({cookie_count} cookies)")

        browser.close()

    print()

    # The GitHub Actions runner is this same PC and reads session.json
    # straight from this folder — nothing needs uploading anywhere.
    # (If a PAT is available we also push to the GitHub secret as a backup.)
    if GITHUB_PAT:
        push_session_to_github(storage)

    print()
    print("Done. The dashboard will update on the next run (hourly, or click")
    print("the Refresh button on the dashboard page).")


if __name__ == "__main__":
    main()
