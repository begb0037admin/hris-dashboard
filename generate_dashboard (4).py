"""
generate_dashboard.py
HRIS Team Dashboard — Ivanti SAASIT → GitHub Pages
----------------------------------------------------
Runs daily at 08:00 via Windows Task Scheduler.
1. Launches a headed Chromium browser via Playwright
2. Logs in through Oxford SSO (manual step on first run; cookies cached after that)
3. Fetches open tasks from the OData API using the live session cookies
4. Builds a self-contained index.html dashboard
5. Commits and pushes to begb0037admin/hris-dashboard (GitHub Pages)

Setup (one-time):
  pip install playwright requests gitpython
  playwright install chromium

Configuration — edit the block below before first run.
"""

# ─────────────────────────────────────────────
#  CONFIGURATION  — edit these values
# ─────────────────────────────────────────────
GITHUB_TOKEN   = "YOUR_GITHUB_PAT_HERE"          # repo scope PAT
GITHUB_REPO    = "begb0037admin/hris-dashboard"
REPO_BRANCH    = "main"

SAASIT_BASE    = "https://oxford.saasiteu.com"
SAASIT_LOGIN   = f"{SAASIT_BASE}/Default.aspx"    # SSO entry point
COOKIE_FILE    = "saasit_cookies.json"            # cached after first login

# Local clone path — script will clone here if folder doesn't exist
REPO_LOCAL     = "hris-dashboard-local"

TEAM_MEMBERS = {
    "musf0100": "James Salas Guillen",
    "ouit0422": "Asta Palmer",
    "ouit0036": "Michael O'Sullivan",
    "admn2716": "Simon Burford",
    "kele0001": "Kevin Lelitte",
}
TEAM_NAME = "HRIS Analysis"

OPEN_STATUSES = ["Accepted", "Active", "Assigned", "Waiting", "Logged", "In Progress"]

TICKET_BASE_URL = f"{SAASIT_BASE}/servicedesk/WebObjects/ServiceDesk.woa/wa/Tasks/taskDetails?recordId="
# ─────────────────────────────────────────────

import json
import os
import sys
import time
import logging
from datetime import datetime, timezone
from pathlib import Path

import requests
from playwright.sync_api import sync_playwright
import git  # gitpython

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("dashboard_run.log", encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)


# ─────────────────────────────────────────────
#  AUTH — Playwright SSO login + cookie cache
# ─────────────────────────────────────────────

def load_cookies() -> list | None:
    if Path(COOKIE_FILE).exists():
        with open(COOKIE_FILE, "r") as f:
            cookies = json.load(f)
        log.info("Loaded %d cached cookies from %s", len(cookies), COOKIE_FILE)
        return cookies
    return None


def save_cookies(cookies: list):
    with open(COOKIE_FILE, "w") as f:
        json.dump(cookies, f, indent=2)
    log.info("Saved %d cookies to %s", len(cookies), COOKIE_FILE)


def get_session_cookies() -> dict:
    """
    Returns a requests-compatible cookie dict.
    Uses cached cookies if available; otherwise opens a browser for SSO login.
    On first run the browser window will appear — log in normally, then press
    Enter in this terminal window to continue.
    """
    cached = load_cookies()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=cached is not None)
        context = browser.new_context()

        if cached:
            context.add_cookies(cached)

        page = context.new_page()
        page.goto(SAASIT_LOGIN, wait_until="networkidle")

        # Check whether we landed on a protected page (SSO redirect) or are already in
        if "login" in page.url.lower() or "sso" in page.url.lower() or "shib" in page.url.lower():
            if cached:
                log.warning("Cached cookies expired — need fresh login.")
            log.info("Browser window open. Please log in via Oxford SSO, then press Enter here...")
            input("  >>> Press Enter once you are logged in: ")
            page.wait_for_load_state("networkidle")

        # Verify we reached SAASIT
        if SAASIT_BASE not in page.url:
            log.error("Login may not have completed. Current URL: %s", page.url)

        pw_cookies = context.cookies()
        save_cookies(pw_cookies)
        browser.close()

    # Convert Playwright cookie list → requests dict
    return {c["name"]: c["value"] for c in pw_cookies}


# ─────────────────────────────────────────────
#  DATA FETCH — OData API calls
# ─────────────────────────────────────────────

ODATA_BASE = f"{SAASIT_BASE}/api/odata/businessobject"
COMMON_SELECT = "AssignmentID,Owner,OwnerLoginName,Subject,Status,Priority,ParentLink_Category,CreatedDateTime,LastModDateTime"


def _status_filter() -> str:
    parts = [f"Status eq '{s}'" for s in OPEN_STATUSES]
    return "(" + " or ".join(parts) + ")"


def _owner_filter() -> str:
    parts = [f"OwnerLoginName eq '{u}'" for u in TEAM_MEMBERS.keys()]
    return "(" + " or ".join(parts) + ")"


def fetch_odata(session: requests.Session, url: str) -> list:
    """Pages through OData results, returns all records."""
    results = []
    next_url = url
    while next_url:
        log.info("GET %s", next_url[:120])
        resp = session.get(next_url, timeout=30)
        if resp.status_code != 200:
            log.error("OData error %d: %s", resp.status_code, resp.text[:300])
            break
        data = resp.json()
        results.extend(data.get("value", []))
        next_url = data.get("@odata.nextLink")
    return results


def fetch_assigned_tasks(session: requests.Session) -> list:
    owner_f  = _owner_filter()
    status_f = _status_filter()
    url = (
        f"{ODATA_BASE}/tasks"
        f"?$filter={owner_f} and {status_f}"
        f"&$select={COMMON_SELECT}"
        f"&$orderby=CreatedDateTime desc"
    )
    return fetch_odata(session, url)


def fetch_unassigned_tasks(session: requests.Session) -> list:
    status_f = _status_filter()
    team_f = (
        "(OwnerTeam eq 'Health and Safety 2nd Line'"
        " or OwnerTeam eq 'HRIS Analysis'"
        " or OwnerTeam eq 'HRIS Service Desk'"
        " or OwnerTeam eq 'HRIS Sys Admin')"
    )
    url = (
        f"{ODATA_BASE}/tasks"
        f"?$filter={team_f} and (Owner eq null) and {status_f}"
        f"&$select={COMMON_SELECT}"
        f"&$orderby=CreatedDateTime desc"
    )
    return fetch_odata(session, url)


# ─────────────────────────────────────────────
#  HTML BUILDER
# ─────────────────────────────────────────────

def priority_badge(priority: str | None) -> str:
    p = (priority or "").strip()
    colours = {
        "Critical": ("#c0392b", "#fff"),
        "High":     ("#e67e22", "#fff"),
        "Medium":   ("#2980b9", "#fff"),
        "Low":      ("#27ae60", "#fff"),
    }
    bg, fg = colours.get(p, ("#95a5a6", "#fff"))
    return f'<span class="badge" style="background:{bg};color:{fg}">{p or "—"}</span>'


def status_badge(status: str | None) -> str:
    s = (status or "").strip()
    colours = {
        "Active":      "#2980b9",
        "In Progress": "#8e44ad",
        "Assigned":    "#16a085",
        "Accepted":    "#27ae60",
        "Waiting":     "#e67e22",
        "Logged":      "#7f8c8d",
    }
    bg = colours.get(s, "#95a5a6")
    return f'<span class="badge" style="background:{bg};color:#fff">{s or "—"}</span>'


def fmt_date(dt_str: str | None) -> str:
    if not dt_str:
        return "—"
    try:
        dt = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
        return dt.strftime("%d %b %Y")
    except Exception:
        return dt_str[:10]


def ticket_link(ticket_id: str | None) -> str:
    if not ticket_id:
        return "—"
    url = f"{TICKET_BASE_URL}{ticket_id}"
    return f'<a href="{url}" target="_blank" rel="noopener">{ticket_id}</a>'


def task_rows(tasks: list) -> str:
    if not tasks:
        return '<tr><td colspan="6" class="empty">No open tasks</td></tr>'
    rows = []
    for t in tasks:
        rows.append(
            f"<tr>"
            f"<td>{ticket_link(t.get('AssignmentID'))}</td>"
            f"<td class='subject'>{t.get('Subject','—')}</td>"
            f"<td>{status_badge(t.get('Status'))}</td>"
            f"<td>{priority_badge(t.get('Priority'))}</td>"
            f"<td>{t.get('ParentLink_Category','—')}</td>"
            f"<td>{fmt_date(t.get('CreatedDateTime'))}</td>"
            f"</tr>"
        )
    return "\n".join(rows)


def person_section(username: str, display_name: str, tasks: list) -> str:
    count = len(tasks)
    return f"""
    <section class="person-section">
      <h2>
        <span class="person-name">{display_name}</span>
        <span class="person-count">{count} open</span>
      </h2>
      <table>
        <thead>
          <tr>
            <th style="width:110px">Ticket ID</th>
            <th>Subject</th>
            <th style="width:115px">Status</th>
            <th style="width:90px">Priority</th>
            <th style="width:140px">Category</th>
            <th style="width:100px">Created</th>
          </tr>
        </thead>
        <tbody>
          {task_rows(tasks)}
        </tbody>
      </table>
    </section>"""


def build_html(assigned: list, unassigned: list) -> str:
    now_utc  = datetime.now(timezone.utc)
    now_bst  = now_utc  # could localise; keep UTC for simplicity
    updated  = now_bst.strftime("%A %d %B %Y at %H:%M UTC")

    # Split assigned tasks by owner
    by_owner: dict[str, list] = {u: [] for u in TEAM_MEMBERS}
    for t in assigned:
        owner = t.get("Owner", "")
        if owner in by_owner:
            by_owner[owner].append(t)

    total_assigned   = len(assigned)
    total_unassigned = len(unassigned)

    # Per-person sections
    person_sections_html = ""
    for username, display_name in TEAM_MEMBERS.items():
        person_sections_html += person_section(username, display_name, by_owner[username])

    # Unassigned section
    unassigned_rows = task_rows(unassigned)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>HRIS Team — Open Tickets</title>
  <style>
    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}

    body {{
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
      background: #f0f2f5;
      color: #1a1a2e;
      min-height: 100vh;
    }}

    /* ── Header ── */
    header {{
      background: linear-gradient(135deg, #002147 0%, #003d80 100%);
      color: #fff;
      padding: 28px 40px 24px;
      display: flex;
      align-items: flex-end;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 12px;
    }}
    header h1 {{
      font-size: 1.6rem;
      font-weight: 700;
      letter-spacing: -0.3px;
    }}
    header .subtitle {{
      font-size: 0.85rem;
      opacity: 0.75;
      margin-top: 4px;
    }}
    .last-updated {{
      font-size: 0.8rem;
      opacity: 0.7;
      text-align: right;
    }}

    /* ── Summary bar ── */
    .summary-bar {{
      display: flex;
      gap: 16px;
      padding: 20px 40px;
      background: #fff;
      border-bottom: 1px solid #e0e4ea;
      flex-wrap: wrap;
    }}
    .stat-card {{
      background: #f7f9fc;
      border: 1px solid #dce1ea;
      border-radius: 10px;
      padding: 14px 24px;
      min-width: 160px;
      text-align: center;
    }}
    .stat-card .stat-number {{
      font-size: 2rem;
      font-weight: 700;
      color: #002147;
      line-height: 1;
    }}
    .stat-card .stat-number.red {{ color: #c0392b; }}
    .stat-card .stat-label {{
      font-size: 0.78rem;
      color: #6b7280;
      margin-top: 5px;
      text-transform: uppercase;
      letter-spacing: 0.4px;
    }}

    /* ── Main content ── */
    main {{
      padding: 32px 40px;
      max-width: 1400px;
      margin: 0 auto;
    }}

    /* ── Person sections ── */
    .person-section {{
      background: #fff;
      border-radius: 12px;
      box-shadow: 0 1px 4px rgba(0,0,0,.07);
      margin-bottom: 28px;
      overflow: hidden;
    }}
    .person-section h2 {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 16px 24px;
      background: #f7f9fc;
      border-bottom: 1px solid #e0e4ea;
      font-size: 1rem;
    }}
    .person-name {{ font-weight: 600; color: #002147; }}
    .person-count {{
      font-size: 0.82rem;
      background: #002147;
      color: #fff;
      border-radius: 20px;
      padding: 3px 12px;
      font-weight: 500;
    }}

    /* ── Unassigned section ── */
    .unassigned-section {{
      background: #fff;
      border-radius: 12px;
      box-shadow: 0 1px 4px rgba(0,0,0,.07);
      margin-bottom: 28px;
      overflow: hidden;
      border: 2px solid #e74c3c;
    }}
    .unassigned-section h2 {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 16px 24px;
      background: #fdf2f2;
      border-bottom: 1px solid #f5c6c6;
      font-size: 1rem;
    }}
    .unassigned-label {{ font-weight: 600; color: #c0392b; }}
    .unassigned-count {{
      font-size: 0.82rem;
      background: #c0392b;
      color: #fff;
      border-radius: 20px;
      padding: 3px 12px;
      font-weight: 500;
    }}

    /* ── Tables ── */
    table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 0.875rem;
    }}
    th {{
      text-align: left;
      padding: 10px 16px;
      font-size: 0.75rem;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      color: #6b7280;
      background: #fafbfc;
      border-bottom: 1px solid #e0e4ea;
    }}
    td {{
      padding: 11px 16px;
      border-bottom: 1px solid #f0f2f5;
      vertical-align: middle;
    }}
    tr:last-child td {{ border-bottom: none; }}
    tr:hover td {{ background: #f7f9fc; }}
    .subject {{ max-width: 380px; }}
    .empty {{
      text-align: center;
      color: #9ca3af;
      padding: 28px;
      font-style: italic;
    }}

    td a {{
      color: #0057b8;
      text-decoration: none;
      font-weight: 500;
    }}
    td a:hover {{ text-decoration: underline; }}

    /* ── Badges ── */
    .badge {{
      display: inline-block;
      padding: 3px 10px;
      border-radius: 12px;
      font-size: 0.73rem;
      font-weight: 600;
      letter-spacing: 0.2px;
      white-space: nowrap;
    }}

    /* ── Footer ── */
    footer {{
      text-align: center;
      font-size: 0.78rem;
      color: #9ca3af;
      padding: 24px 40px 40px;
    }}
  </style>
</head>
<body>

<header>
  <div>
    <h1>HRIS Team — Open Tickets</h1>
    <div class="subtitle">Ivanti SAASIT · University of Oxford</div>
  </div>
  <div class="last-updated">
    Last updated<br><strong>{updated}</strong>
  </div>
</header>

<div class="summary-bar">
  <div class="stat-card">
    <div class="stat-number">{total_assigned + total_unassigned}</div>
    <div class="stat-label">Total Open</div>
  </div>
  <div class="stat-card">
    <div class="stat-number">{total_assigned}</div>
    <div class="stat-label">Assigned</div>
  </div>
  <div class="stat-card">
    <div class="stat-number red">{total_unassigned}</div>
    <div class="stat-label">⚠ Unassigned</div>
  </div>
  <div class="stat-card">
    <div class="stat-number">{len(TEAM_MEMBERS)}</div>
    <div class="stat-label">Team Members</div>
  </div>
</div>

<main>

  {person_sections_html}

  <section class="unassigned-section">
    <h2>
      <span class="unassigned-label">⚠ Unassigned — {TEAM_NAME}</span>
      <span class="unassigned-count">{total_unassigned} open</span>
    </h2>
    <table>
      <thead>
        <tr>
          <th style="width:110px">Ticket ID</th>
          <th>Subject</th>
          <th style="width:115px">Status</th>
          <th style="width:90px">Priority</th>
          <th style="width:140px">Category</th>
          <th style="width:100px">Created</th>
        </tr>
      </thead>
      <tbody>
        {unassigned_rows}
      </tbody>
    </table>
  </section>

</main>

<footer>
  Auto-generated by generate_dashboard.py · Data sourced from Ivanti SAASIT OData API
</footer>

</body>
</html>"""


# ─────────────────────────────────────────────
#  GIT PUSH
# ─────────────────────────────────────────────

def push_to_github(html_content: str):
    repo_path = Path(REPO_LOCAL)

    if not repo_path.exists():
        log.info("Cloning repo to %s ...", repo_path)
        clone_url = f"https://{GITHUB_TOKEN}@github.com/{GITHUB_REPO}.git"
        git.Repo.clone_from(clone_url, repo_path, branch=REPO_BRANCH)

    repo = git.Repo(repo_path)

    # Update remote URL with current token (in case token changed)
    origin = repo.remote("origin")
    origin.set_url(f"https://{GITHUB_TOKEN}@github.com/{GITHUB_REPO}.git")

    # Pull latest
    origin.pull(REPO_BRANCH)

    # Write dashboard
    index_path = repo_path / "index.html"
    index_path.write_text(html_content, encoding="utf-8")

    # Commit and push
    repo.index.add(["index.html"])
    if repo.is_dirty(index=True):
        stamp = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
        repo.index.commit(f"dashboard: auto-update {stamp}")
        origin.push(REPO_BRANCH)
        log.info("Pushed updated dashboard to GitHub.")
    else:
        log.info("No changes detected — skipping push.")


# ─────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────

def main():
    log.info("=== HRIS Dashboard run started ===")

    # 1. Get authenticated session cookies
    pw_cookies = get_session_cookies()

    session = requests.Session()
    session.cookies.update(pw_cookies)
    session.headers.update({
        "Accept": "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
    })

    # 2. Fetch data
    log.info("Fetching assigned tasks ...")
    assigned   = fetch_assigned_tasks(session)
    log.info("  → %d assigned tasks", len(assigned))

    log.info("Fetching unassigned tasks ...")
    unassigned = fetch_unassigned_tasks(session)
    log.info("  → %d unassigned tasks", len(unassigned))

    # 3. Build HTML
    html = build_html(assigned, unassigned)
    log.info("Dashboard HTML built (%d bytes)", len(html))

    # 4. Push to GitHub
    push_to_github(html)

    log.info("=== Run complete ===")


if __name__ == "__main__":
    main()
