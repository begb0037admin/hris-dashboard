"""
generate_dashboard.py
HRIS Team Dashboard — Ivanti SAASIT → GitHub Pages
----------------------------------------------------
Runs via GitHub Actions (workflow_dispatch) or manually on any machine.

1. Loads browser session — from SAASIT_SESSION env var (GitHub Actions)
   or session.json (local run created by login.py)
2. Fetches open tasks from the OData API via Playwright headless browser
3. Filters open statuses client-side
4. Builds a self-contained index.html dashboard
5. In GitHub Actions: writes index.html; the workflow handles git push.
   Locally: clones repo, commits, and pushes via gitpython.

Dependencies:
  pip install playwright requests gitpython tzdata
  playwright install chromium

First login (one-time, or when session expires):
  python login.py
  (then follow prompts — saves session to session.json AND pushes to GitHub secret)
"""

# ─────────────────────────────────────────────
#  CONFIGURATION  — set your values here
# ─────────────────────────────────────────────

# Personal Access Token — needs scopes: repo, workflow
# Used ONLY for local git push. NEVER hardcode it here: GitHub auto-revokes any
# PAT pushed to a public repo (which is why the Refresh button kept breaking).
# Set the HRIS_GITHUB_PAT env var, or put the token in a github_pat.txt file
# next to this script (git-ignored). The dashboard's Refresh button no longer
# embeds a token — it prompts once and stores it in the browser's localStorage.
GITHUB_PAT     = __import__("os").environ.get("HRIS_GITHUB_PAT", "").strip()
if not GITHUB_PAT:
    _pat_file = __import__("pathlib").Path(__file__).with_name("github_pat.txt")
    if _pat_file.exists():
        GITHUB_PAT = _pat_file.read_text(encoding="utf-8").strip()

GITHUB_REPO    = "begb0037admin/hris-dashboard"
REPO_BRANCH    = "main"

SAASIT_BASE    = "https://oxford.saasiteu.com"
REPO_LOCAL     = "hris-dashboard-local"   # local git clone folder (local mode only)

TEAM_MEMBERS = {
    "musf0100": "James Salas Guillen",
    "ouit0422": "Asta Palmer",
    "ouit0036": "Michael O'Sullivan",
    "admn2716": "Simon Burford",
    "begb0037": "Kevin Lelitte",
}
TEAM_NAME = "HRIS Analysis"

OPEN_STATUSES = ["Accepted", "Active", "Assigned", "Waiting", "Logged", "In Progress"]

TICKET_BASE_URL = f"{SAASIT_BASE}/Default.aspx#"

# In GitHub Actions GITHUB_TOKEN env var is the built-in token (used for git push by the workflow).
# Locally it falls back to GITHUB_PAT.
GITHUB_TOKEN = __import__("os").environ.get("GITHUB_TOKEN", GITHUB_PAT)

# ─────────────────────────────────────────────

import json
import os
import sys
import logging
from datetime import datetime, timezone
from pathlib import Path

import requests
import git  # gitpython
from playwright.sync_api import sync_playwright

# UK timezone — handles both GMT (winter) and BST (summer) automatically.
# Requires: pip install tzdata  (needed on Windows; Linux/macOS have it built in)
try:
    from zoneinfo import ZoneInfo
    _UK_TZ = ZoneInfo("Europe/London")
except Exception:
    _UK_TZ = timezone.utc   # fallback — install tzdata to get correct UK time

# Windows runner consoles default to cp1252, which can't print characters like
# '→' and crashes the logging StreamHandler. Force UTF-8 with safe fallback.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

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
#  AUTH — Playwright session (created by login.py)
# ─────────────────────────────────────────────

SESSION_FILE = Path("session.json")


class SessionExpiredError(RuntimeError):
    """SAASIT rejected our session cookies — run login.py to refresh."""


def get_session_cookies() -> dict:
    """
    Loads the saved browser session and returns cookies for oxford.saasiteu.com.

    Source priority:
      1. SAASIT_SESSION environment variable (base — used in GitHub Actions)
      2. session.json on disk (local runs)

    If neither is available, run login.py to create a fresh session.
    """
    session_env = os.environ.get("SAASIT_SESSION")

    if session_env:
        log.info("Loading session from SAASIT_SESSION environment variable ...")
        storage = json.loads(session_env)
    elif SESSION_FILE.exists():
        log.info("Loading session from %s ...", SESSION_FILE)
        storage = json.loads(SESSION_FILE.read_text(encoding="utf-8"))
    else:
        log.error(
            "No session available. Set SAASIT_SESSION env var or run login.py:\n"
            "  python login.py"
        )
        raise FileNotFoundError("No session available — run login.py")

    # Extract cookies directly from stored session — no browser navigation needed.
    # Navigating via headless browser from GitHub Actions IPs causes SSO redirects
    # that invalidate the session. Reading directly from the JSON is both faster
    # and more reliable.
    all_cookies = storage.get("cookies", [])
    cookies = {
        c["name"]: c["value"]
        for c in all_cookies
        if "saasiteu.com" in c.get("domain", "")
    }

    log.info("Got %d session cookies for oxford.saasiteu.com", len(cookies))

    if not cookies:
        raise SessionExpiredError(
            "No cookies found for saasiteu.com — session may have expired. "
            "Re-run login.py to refresh the session."
        )

    return cookies


# ─────────────────────────────────────────────
#  DATA FETCH — OData API calls
# ─────────────────────────────────────────────

ODATA_BASE    = f"{SAASIT_BASE}/api/odata/businessobject"
COMMON_SELECT = "AssignmentID,Owner,OwnerTeam,Subject,Status,Priority,ParentLink_Category,CreatedDateTime,LastModDateTime"


def _status_filter() -> str:
    parts = [f"Status eq '{s}'" for s in OPEN_STATUSES]
    return "(" + " or ".join(parts) + ")"


def fetch_odata(session: requests.Session, endpoint: str, params: dict) -> list:
    """Pages through OData results."""
    results = []
    url = f"{ODATA_BASE}/{endpoint}"
    while url:
        log.info("GET %s %s", endpoint, str(params)[:80])
        resp = session.get(url, params=params, timeout=30)
        if resp.status_code in (401, 403):
            # Expired session — abort the whole run rather than silently
            # publishing an all-zero dashboard.
            raise SessionExpiredError(
                f"OData auth error {resp.status_code}: {resp.text[:300]}"
            )
        if resp.status_code != 200:
            log.error("OData error %d: %s", resp.status_code, resp.text[:300])
            break
        data = resp.json()
        results.extend(data.get("value", []))
        next_link = data.get("@odata.nextLink")
        if next_link:
            url = next_link
            params = {}
        else:
            url = None
    return results


def fetch_assigned_tasks(session: requests.Session) -> list:
    """Fetch per-person to avoid long filter rejection by the API."""
    open_statuses = {s.lower() for s in OPEN_STATUSES}
    all_tasks = []
    for username in TEAM_MEMBERS.keys():
        params = {
            "$filter":  f"Owner eq '{username}'",
            "$select":  COMMON_SELECT,
            "$orderby": "CreatedDateTime desc",
            "$top":     "100",
        }
        tasks = fetch_odata(session, "tasks", params)
        open_tasks = [t for t in tasks if (t.get("Status") or "").lower() in open_statuses]
        log.info("  %s: %d open tasks", username, len(open_tasks))
        all_tasks.extend(open_tasks)
    return all_tasks


UNASSIGNED_TEAMS = [
    "Health and Safety 2nd Line",
    "HRIS Analysis",
    "HRIS Service Desk",
    "HRIS Sys Admin",
]


def fetch_unassigned_tasks(session: requests.Session) -> list:
    """Fetch per-team to avoid long filter rejection by the API."""
    open_statuses = set(s.lower() for s in OPEN_STATUSES)
    team_usernames = set(TEAM_MEMBERS.keys())
    seen_ids = set()
    all_tasks = []
    for team in UNASSIGNED_TEAMS:
        params = {
            "$filter":  f"OwnerTeam eq '{team}'",
            "$select":  COMMON_SELECT,
            "$orderby": "CreatedDateTime desc",
            "$top":     "100",
        }
        tasks = fetch_odata(session, "tasks", params)
        for t in tasks:
            tid = t.get("AssignmentID")
            if tid in seen_ids:
                continue
            seen_ids.add(tid)
            if (
                (t.get("Status", "") or "").lower() in open_statuses
                and (t.get("Owner") or "") not in team_usernames
            ):
                all_tasks.append(t)
    log.info("  unassigned total: %d", len(all_tasks))
    return all_tasks


# ─────────────────────────────────────────────
#  HTML BUILDER
# ─────────────────────────────────────────────

def priority_badge(priority: str | None) -> str:
    p = str(priority or "").strip()
    num_map = {"1": "Critical", "2": "High", "3": "Medium", "4": "Low", "5": "Low"}
    label = num_map.get(p, p) or "—"
    colours = {
        "Critical": ("#c0392b", "#fff"),
        "High":     ("#e67e22", "#fff"),
        "Medium":   ("#2980b9", "#fff"),
        "Low":      ("#27ae60", "#fff"),
    }
    bg, fg = colours.get(label, ("#95a5a6", "#fff"))
    return f'<span class="badge" style="background:{bg};color:{fg}">{label}</span>'


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
    # UK time — shows GMT in winter, BST in summer
    now = datetime.now(_UK_TZ)
    tz_str = now.strftime("%Z") or "UTC"
    updated = now.strftime("%A %d %B %Y at %H:%M ") + tz_str

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

    # Unassigned rows
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
    .header-right {{
      display: flex;
      flex-direction: column;
      align-items: flex-end;
      gap: 6px;
    }}
    .refresh-pill {{
      background: rgba(255,255,255,0.15);
      border: 1px solid rgba(255,255,255,0.45);
      color: #fff;
      padding: 7px 20px;
      border-radius: 20px;
      font-size: 0.85rem;
      font-weight: 600;
      cursor: pointer;
      letter-spacing: 0.2px;
      transition: background 0.15s;
    }}
    .refresh-pill:hover:not(:disabled) {{ background: rgba(255,255,255,0.28); }}
    .refresh-pill:disabled {{ opacity: 0.55; cursor: not-allowed; }}
    .refresh-status {{
      font-size: 0.75rem;
      opacity: 0.85;
      text-align: right;
      min-height: 14px;
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
  <div class="header-right">
    <button id="refresh-btn" class="refresh-pill" onclick="triggerRefresh()">↻ Refresh</button>
    <div id="refresh-status" class="refresh-status"></div>
    <div class="last-updated">Last updated<br><strong>{updated}</strong></div>
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

<script>
// No token is embedded in this page: GitHub auto-revokes any PAT it finds in
// a public repo. Instead the token is requested once and kept in this
// browser's localStorage only.
function getGithubToken() {{
  let token = localStorage.getItem('hris_gh_pat');
  if (!token) {{
    token = prompt(
      'Paste a GitHub personal access token that can run workflows on\\n' +
      '{GITHUB_REPO} (classic PAT with repo + workflow scope,\\n' +
      'or a fine-grained PAT with Actions read/write).\\n\\n' +
      'It is stored only in this browser — never in the repo.'
    );
    if (token) {{
      token = token.trim();
      localStorage.setItem('hris_gh_pat', token);
    }}
  }}
  return token;
}}

async function triggerRefresh() {{
  const btn    = document.getElementById('refresh-btn');
  const status = document.getElementById('refresh-status');

  const token = getGithubToken();
  if (!token) {{
    status.textContent = 'Cancelled — no token provided.';
    return;
  }}

  btn.disabled = true;
  btn.textContent = '↻ Working…';
  status.textContent = 'Triggering update…';

  try {{
    const resp = await fetch(
      'https://api.github.com/repos/{GITHUB_REPO}/actions/workflows/update-dashboard.yml/dispatches',
      {{
        method: 'POST',
        headers: {{
          'Authorization': 'Bearer ' + token,
          'Accept': 'application/vnd.github+json',
          'Content-Type': 'application/json',
        }},
        body: JSON.stringify({{ ref: '{REPO_BRANCH}' }}),
      }}
    );

    if (resp.status === 204) {{
      btn.textContent = '✓ Triggered';
      let secs = 60;
      status.textContent = 'Updating — reload in ' + secs + 's…';
      const timer = setInterval(() => {{
        secs--;
        if (secs > 0) {{
          status.textContent = 'Updating — reload in ' + secs + 's…';
        }} else {{
          clearInterval(timer);
          status.textContent = '✓ Ready — reload now.';
          status.style.fontWeight = 'bold';
          status.style.color = '#27ae60';
        }}
      }}, 1000);
    }} else if (resp.status === 401) {{
      localStorage.removeItem('hris_gh_pat');
      throw new Error('GitHub token rejected (401) — it may have expired. Click Refresh again to enter a new one.');
    }} else {{
      const body = await resp.text();
      throw new Error('GitHub API ' + resp.status + ': ' + body.slice(0, 120));
    }}
  }} catch (e) {{
    btn.textContent = '↻ Refresh';
    btn.disabled = false;
    status.textContent = 'Error: ' + e.message;
  }}
}}
</script>

</body>
</html>"""


# ─────────────────────────────────────────────
#  GIT PUSH (local mode only)
# ─────────────────────────────────────────────

def push_to_github(html_content: str):
    """Clone/pull the repo, write index.html, commit and push. Local runs only."""
    repo_path = Path(REPO_LOCAL)

    if not repo_path.exists():
        log.info("Cloning repo to %s ...", repo_path)
        clone_url = f"https://{GITHUB_TOKEN}@github.com/{GITHUB_REPO}.git"
        git.Repo.clone_from(clone_url, repo_path, branch=REPO_BRANCH)

    repo = git.Repo(repo_path)
    origin = repo.remote("origin")
    origin.set_url(f"https://{GITHUB_TOKEN}@github.com/{GITHUB_REPO}.git")
    origin.pull(REPO_BRANCH)

    index_path = repo_path / "index.html"
    index_path.write_text(html_content, encoding="utf-8")

    repo.index.add(["index.html"])
    if repo.is_dirty(index=True):
        now = datetime.now(_UK_TZ)
        tz_str = now.strftime("%Z") or "UTC"
        stamp = now.strftime("%Y-%m-%d %H:%M ") + tz_str
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

    # 1. Auth
    cookies = get_session_cookies()

    session = requests.Session()
    session.cookies.update(cookies)
    session.headers.update({
        "Accept": "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
    })

    # 2. Fetch
    log.info("Fetching assigned tasks ...")
    assigned = fetch_assigned_tasks(session)
    log.info("  → %d assigned tasks", len(assigned))

    log.info("Fetching unassigned tasks ...")
    unassigned = fetch_unassigned_tasks(session)
    log.info("  → %d unassigned tasks", len(unassigned))

    # 3. Build HTML
    html = build_html(assigned, unassigned)
    log.info("Dashboard HTML built (%d bytes)", len(html))

    # 4. Output
    if os.environ.get("GITHUB_ACTIONS") == "true":
        # Running in GitHub Actions — write index.html to the repo root.
        # The workflow's final step handles git commit and push.
        Path("index.html").write_text(html, encoding="utf-8")
        log.info("Wrote index.html — git push handled by workflow.")
    else:
        # Local run — clone, commit, push via gitpython.
        push_to_github(html)

    log.info("=== Run complete ===")


if __name__ == "__main__":
    try:
        main()
    except SessionExpiredError as e:
        log.error("SAASIT session expired: %s", e)
        log.error(
            "Dashboard NOT updated — the previous data stays live.\n"
            "Fix: run 'Refresh Session.bat' (python login.py) and log in via "
            "Oxford SSO to refresh the SAASIT_SESSION secret."
        )
        sys.exit(1)
