"""
push_automation_status.py -- pushes data/last_automated_run.json to GitHub.

Why this exists: the dashboard (index.html) only ever showed
data.updated_display, which is stamped by import_osm_report.py whenever
*anyone* (manual double-click or the automated morning run) successfully
finishes a full pipeline run. That means a FAILED automated run leaves
zero trace on the dashboard -- yesterday's tickets.json just keeps showing
yesterday's (or this morning's manual) updated_display, and a stale
dashboard from a silently-broken automation looks identical to a genuinely
fresh one. Kevin's explicit ask (21 Aug 2026): he needs to look at the
dashboard and immediately know whether THIS MORNING'S automated run
happened and whether it worked -- not infer it from tickets.json's own
timestamp.

This script is called by Run_HRIS_Auto_Refresh.bat -- and ONLY by that
script -- at every exit point (success and every distinct failure mode),
so data/last_automated_run.json is the automation's own self-reported
record of what actually happened on ITS last attempt, independent of
whether tickets.json itself changed. The existing manual pipeline
(Update HRIS Dashboard.bat / import_osm_report.py) is completely
untouched by this file.

Usage:
    python push_automation_status.py --status success --step dashboard_update_ok --detail "..."
    python push_automation_status.py --status failure --step fetch_failed --detail "..."
    python push_automation_status.py --status failure --step fetch_failed --detail "..." --attempt 2 --max-attempts 3

--attempt / --max-attempts (added 22 Aug 2026, both optional, default 1/1):
    records which of the morning's scheduled attempts (08:45 primary,
    09:15/09:45 retries -- see check_last_run_status.py) this run was, so
    the dashboard can show e.g. "Failed (attempt 3 of 3, no more retries
    today)" instead of a bare failure with no context on whether more
    retries are still coming.

Requires:
    pip install requests
    Windows User env var: GITHUB_PAT (same PAT already used by
    import_osm_report.py -- Contents RW on begb0037admin/hris-dashboard)
"""

import os
import sys
import json
import base64
import argparse
from datetime import datetime

try:
    import requests
except ImportError:
    sys.exit("ERROR: requests not installed. Run: pip install requests")

GITHUB_REPO = "begb0037admin/hris-dashboard"
GITHUB_REF = "main"
API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/contents/data/last_automated_run.json"


def get_github_pat():
    pat = os.environ.get("GITHUB_PAT", "").strip()
    if not pat:
        sys.exit(
            "ERROR: GITHUB_PAT environment variable is not set.\n"
            "Set it as a Windows User environment variable (never in any file)."
        )
    return pat


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--status", required=True, choices=["success", "failure"])
    parser.add_argument("--step", required=True,
                         help="short machine-readable stage identifier, e.g. "
                              "fetch_failed / dashboard_update_ok")
    parser.add_argument("--detail", default="", help="short human-readable detail line")
    parser.add_argument("--attempt", type=int, default=1,
                         help="which scheduled attempt this run was (1-based)")
    parser.add_argument("--max-attempts", type=int, default=1,
                         help="total scheduled attempts today")
    args = parser.parse_args()

    pat = get_github_pat()
    now = datetime.now()

    payload = {
        "timestamp": now.strftime("%Y-%m-%dT%H:%M:%S"),
        "timestamp_display": now.strftime("%A %d %B %Y at %H:%M"),
        "status": args.status,
        "step": args.step,
        "detail": args.detail,
        "trigger": "automated",
        "attempt": args.attempt,
        "max_attempts": args.max_attempts,
    }
    body = json.dumps(payload, indent=2, ensure_ascii=False).encode("utf-8")

    headers = {
        "Authorization": f"token {pat}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }

    resp = requests.get(f"{API_URL}?ref={GITHUB_REF}", headers=headers, timeout=30)
    if resp.status_code == 404:
        current_sha = None
    elif resp.status_code == 200:
        current_sha = resp.json()["sha"]
    else:
        sys.exit(f"ERROR: GET {API_URL} failed (HTTP {resp.status_code})\n{resp.text}")

    put_payload = {
        "message": f"automation status -- {args.status} ({now.strftime('%Y-%m-%d %H:%M')})",
        "content": base64.b64encode(body).decode("ascii"),
        "branch": GITHUB_REF,
    }
    if current_sha:
        put_payload["sha"] = current_sha

    put_resp = requests.put(API_URL, headers=headers, json=put_payload, timeout=30)
    if put_resp.status_code not in (200, 201):
        sys.exit(f"ERROR: PUT {API_URL} failed (HTTP {put_resp.status_code})\n{put_resp.text}")

    print(f"Pushed data/last_automated_run.json: status={args.status} step={args.step}")


if __name__ == "__main__":
    main()
