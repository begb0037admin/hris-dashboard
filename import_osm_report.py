"""
import_osm_report.py — OSM Excel → HRIS Dashboard

Usage:
    python import_osm_report.py

Reads the most recently modified "All Open Tasks by Team*.xlsx" from
C:\\Users\\admin\\Downloads, pushes data/tickets.json (for the new v2
dashboard) and regenerates index.html (legacy fallback), then pushes
both to GitHub.

Requires:
    pip install openpyxl requests
    Windows User env var: GITHUB_PAT (fine-grained PAT, Contents RW on hris-dashboard)
"""

import os
import sys
import glob
import base64
import json
import re
from datetime import datetime

try:
    import openpyxl
except ImportError:
    sys.exit("ERROR: openpyxl not installed. Run: pip install openpyxl")

try:
    import requests
except ImportError:
    sys.exit("ERROR: requests not installed. Run: pip install requests")

# ── Config ─────────────────────────────────────────────────────────────────────────────

DOWNLOADS    = r"C:\Users\admin\Downloads"
FILE_PATTERN = "All Open Tasks by Team*.xlsx"
GITHUB_REPO  = "begb0037admin/hris-dashboard"
GITHUB_REF   = "main"
BASE_API     = f"https://api.github.com/repos/{GITHUB_REPO}/contents"
GITHUB_API   = f"{BASE_API}/index.html"
DATA_API     = f"{BASE_API}/data/tickets.json"

# Column indices (0-based) — fallback if header detection fails
COL_OWNER_TEAM  = 0
COL_TASK_NUMBER = 5
COL_SUMMARY     = 6
COL_STATUS      = 7
COL_PRIORITY    = 8
COL_ANALYST     = 9
COL_CATEGORY    = 4
COL_CREATED     = 11
COL_DAYS_OPEN   = 15

# Display order on the dashboard
KNOWN_ANALYSTS = [
    "James Salas Guillen",
    "Asta Palmer",
    "Michael O'Sullivan",
    "Simon Burford",
    "Kevin Lelitte",
]

STALE_THRESHOLD_DAYS = 30

# ── Step 1: find the report file ──────────────────────────────────────────────────

def find_report():
    pattern = os.path.join(DOWNLOADS, FILE_PATTERN)
    files = glob.glob(pattern)
    if not files:
        sys.exit(f"ERROR: No file matching '{FILE_PATTERN}' found in {DOWNLOADS}")
    latest = max(files, key=os.path.getmtime)
    print(f"Found report: {os.path.basename(latest)}")
    print(f"  Modified:   {datetime.fromtimestamp(os.path.getmtime(latest)).strftime('%d %b %Y %H:%M')}")
    return latest

# ── Step 2: parse the Excel ─────────────────────────────────────────────────────

def parse_report(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    header_idx = None
    for i, row in enumerate(rows):
        if any(str(c).strip() == "Task Number" for c in row if c is not None):
            header_idx = i
            break

    if header_idx is None:
        sys.exit("ERROR: Could not find header row containing 'Task Number'.")

    print(f"  Header row: {header_idx + 1} (1-indexed)")

    header = rows[header_idx]
    col_map = {str(cell).strip(): idx for idx, cell in enumerate(header) if cell is not None}

    def col(name, fallback):
        return col_map.get(name, fallback)

    c_task     = col("Task Number", COL_TASK_NUMBER)
    c_summary  = col("Summary",     COL_SUMMARY)
    c_status   = col("Status",      COL_STATUS)
    c_priority = col("Priority",    COL_PRIORITY)
    c_analyst  = col("Analyst",     COL_ANALYST)
    c_category = col("Category",    COL_CATEGORY)
    c_created  = col("Created",     COL_CREATED)
    c_days     = col("Days Open",   COL_DAYS_OPEN)
    c_team     = col("Owner Team",  COL_OWNER_TEAM)

    tickets = []
    for row in rows[header_idx + 1:]:
        if all(c is None or str(c).strip() == "" for c in row):
            continue

        def cell(idx):
            try:
                v = row[idx]
                return str(v).strip() if v is not None else ""
            except IndexError:
                return ""

        task_num = cell(c_task)
        if not task_num:
            continue

        created = cell(c_created)
        raw_created = row[c_created] if c_created < len(row) else None
        if isinstance(raw_created, datetime):
            created = raw_created.strftime("%d %b %Y")
        elif created and re.match(r"\d{4}-\d{2}-\d{2}", created):
            try:
                created = datetime.strptime(created[:10], "%Y-%m-%d").strftime("%d %b %Y")
            except ValueError:
                pass

        days_raw = cell(c_days)
        try:
            days_int = int(float(days_raw)) if days_raw else 0
        except ValueError:
            days_int = 0

        tickets.append({
            "task_num": task_num,
            "summary":  cell(c_summary),
            "status":   cell(c_status),
            "priority": cell(c_priority),
            "analyst":  cell(c_analyst),
            "category": cell(c_category),
            "created":  created,
            "days":     days_int,
            "team":     cell(c_team),
        })

    print(f"  Tickets parsed: {len(tickets)}")
    return tickets

# ── Step 3: group tickets ──────────────────────────────────────────────────────────────

def group_tickets(tickets):
    by_analyst = {name: [] for name in KNOWN_ANALYSTS}
    unassigned = []
    other = {}

    for t in tickets:
        a = t["analyst"].strip()
        if not a:
            unassigned.append(t)
        elif a in by_analyst:
            by_analyst[a].append(t)
        else:
            other.setdefault(a, []).append(t)

    for name in sorted(other):
        by_analyst[name] = other[name]

    return by_analyst, unassigned

# ── Step 4: build tickets.json ────────────────────────────────────────────────────────────

def build_tickets_json(by_analyst, unassigned, report_path):
    all_tickets = [t for ts in by_analyst.values() for t in ts] + unassigned
    total     = len(all_tickets)
    assigned  = total - len(unassigned)
    stale     = [t for t in all_tickets if t["days"] >= STALE_THRESHOLD_DAYS]
    oldest    = max((t["days"] for t in all_tickets), default=0)

    now = datetime.now()

    analysts_list = []
    for name in list(by_analyst.keys()):
        tickets = by_analyst[name]
        if tickets or name in KNOWN_ANALYSTS:
            analysts_list.append({
                "name": name,
                "ticket_count": len(tickets),
                "tickets": tickets,
            })

    return {
        "updated":         now.strftime("%Y-%m-%dT%H:%M:%S"),
        "updated_display": now.strftime("%A %d %B %Y at %H:%M"),
        "report_file":     os.path.basename(report_path),
        "summary": {
            "total":      total,
            "assigned":   assigned,
            "unassigned": len(unassigned),
            "stale":      len(stale),
            "oldest_days": oldest,
        },
        "analysts":   analysts_list,
        "unassigned": unassigned,
    }

# ── Step 5: HTML helpers (legacy index.html) ────────────────────────────────────────────

def esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
                  .replace(">", "&gt;").replace('"', "&quot;"))

PRIORITY_COLOURS = {
    "1 - Critical": "#7f1d1d",
    "2 - High":     "#b91c1c",
    "3 - Medium":   "#b45309",
    "4 - Low":      "#1d4ed8",
    "5 - Very Low": "#6b7280",
}

def priority_badge(p):
    colour = PRIORITY_COLOURS.get(p, "#6b7280")
    label = p.split(" - ", 1)[-1] if " - " in p else p
    return f'<span class="badge" style="background:{colour};color:#fff">{esc(label)}</span>'

def ticket_rows(tickets):
    if not tickets:
        return '<tr><td colspan="6" class="empty">No open tasks</td></tr>'
    rows = []
    for t in tickets:
        days_str = f" ({t['days']}d)" if t["days"] else ""
        rows.append(
            f'<tr>'
            f'<td><strong>{esc(t["task_num"])}</strong></td>'
            f'<td class="subject">{esc(t["summary"])}</td>'
            f'<td>{esc(t["status"])}</td>'
            f'<td>{priority_badge(t["priority"])}</td>'
            f'<td>{esc(t["category"])}</td>'
            f'<td>{esc(t["created"])}{esc(days_str)}</td>'
            f'</tr>'
        )
    return "\n          ".join(rows)

TABLE_HEADER = """<thead>
          <tr>
            <th style="width:130px">Ticket</th>
            <th>Summary</th>
            <th style="width:115px">Status</th>
            <th style="width:100px">Priority</th>
            <th style="width:150px">Category</th>
            <th style="width:120px">Created</th>
          </tr>
        </thead>"""

def person_section(name, tickets):
    count = len(tickets)
    label = f"{count} open" if count != 1 else "1 open"
    return f"""  <section class="person-section">
      <h2>
        <span class="person-name">{esc(name)}</span>
        <span class="person-count">{label}</span>
      </h2>
      <table>
        {TABLE_HEADER}
        <tbody>
          {ticket_rows(tickets)}
        </tbody>
      </table>
    </section>"""

def unassigned_section(tickets):
    count = len(tickets)
    label = f"{count} open" if count != 1 else "1 open"
    return f"""  <section class="unassigned-section">
      <h2>
        <span class="unassigned-label">&#9888; Unassigned</span>
        <span class="unassigned-count">{label}</span>
      </h2>
      <table>
        {TABLE_HEADER}
        <tbody>
          {ticket_rows(tickets)}
        </tbody>
      </table>
    </section>"""

def build_html(by_analyst, unassigned, report_path):
    total      = sum(len(v) for v in by_analyst.values()) + len(unassigned)
    assigned   = total - len(unassigned)
    team_count = len([n for n, t in by_analyst.items() if n in KNOWN_ANALYSTS])
    now        = datetime.now()
    updated    = now.strftime("%A %d %B %Y at %H:%M BST")
    report_name = os.path.basename(report_path)

    person_sections_html = "\n\n".join(
        person_section(name, tickets)
        for name, tickets in by_analyst.items()
        if tickets or name in KNOWN_ANALYSTS
    )
    unassigned_html = unassigned_section(unassigned)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>HRIS Team — Open Tickets</title>
  <style>
    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; background: #f0f2f5; color: #1a1a2e; min-height: 100vh; }}
    header {{ background: linear-gradient(135deg, #002147 0%, #003d80 100%); color: #fff; padding: 28px 40px 24px; display: flex; align-items: flex-end; justify-content: space-between; flex-wrap: wrap; gap: 12px; }}
    header h1 {{ font-size: 1.6rem; font-weight: 700; letter-spacing: -0.3px; }}
    header .subtitle {{ font-size: 0.85rem; opacity: 0.75; margin-top: 4px; }}
    .last-updated {{ font-size: 0.8rem; opacity: 0.7; text-align: right; }}
    .summary-bar {{ display: flex; gap: 16px; padding: 20px 40px; background: #fff; border-bottom: 1px solid #e0e4ea; flex-wrap: wrap; }}
    .stat-card {{ background: #f7f9fc; border: 1px solid #dce1ea; border-radius: 10px; padding: 14px 24px; min-width: 160px; text-align: center; }}
    .stat-number {{ font-size: 2rem; font-weight: 700; color: #002147; line-height: 1; }}
    .stat-number.red {{ color: #c0392b; }}
    .stat-label {{ font-size: 0.78rem; color: #6b7280; margin-top: 5px; text-transform: uppercase; letter-spacing: 0.4px; }}
    main {{ padding: 32px 40px; max-width: 1400px; margin: 0 auto; }}
    .person-section {{ background: #fff; border-radius: 12px; box-shadow: 0 1px 4px rgba(0,0,0,.07); margin-bottom: 28px; overflow: hidden; }}
    .person-section h2 {{ display: flex; align-items: center; justify-content: space-between; padding: 16px 24px; background: #f7f9fc; border-bottom: 1px solid #e0e4ea; font-size: 1rem; }}
    .person-name {{ font-weight: 600; color: #002147; }}
    .person-count {{ font-size: 0.82rem; background: #002147; color: #fff; border-radius: 20px; padding: 3px 12px; font-weight: 500; }}
    .unassigned-section {{ background: #fff; border-radius: 12px; box-shadow: 0 1px 4px rgba(0,0,0,.07); margin-bottom: 28px; overflow: hidden; border: 2px solid #e74c3c; }}
    .unassigned-section h2 {{ display: flex; align-items: center; justify-content: space-between; padding: 16px 24px; background: #fdf2f2; border-bottom: 1px solid #f5c6c6; font-size: 1rem; }}
    .unassigned-label {{ font-weight: 600; color: #c0392b; }}
    .unassigned-count {{ font-size: 0.82rem; background: #c0392b; color: #fff; border-radius: 20px; padding: 3px 12px; font-weight: 500; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 0.875rem; }}
    th {{ text-align: left; padding: 10px 16px; font-size: 0.75rem; text-transform: uppercase; letter-spacing: 0.5px; color: #6b7280; background: #fafbfc; border-bottom: 1px solid #e0e4ea; }}
    td {{ padding: 11px 16px; border-bottom: 1px solid #f0f2f5; vertical-align: middle; }}
    tr:last-child td {{ border-bottom: none; }}
    tr:hover td {{ background: #f7f9fc; }}
    .subject {{ max-width: 420px; }}
    .empty {{ text-align: center; color: #9ca3af; padding: 28px; font-style: italic; }}
    .badge {{ display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 0.73rem; font-weight: 600; letter-spacing: 0.2px; white-space: nowrap; }}
    footer {{ text-align: center; font-size: 0.78rem; color: #9ca3af; padding: 24px 40px 40px; }}
  </style>
</head>
<body>
<header>
  <div>
    <h1>HRIS Team — Open Tickets</h1>
    <div class="subtitle">Oxford Service Manager (OSM) · University of Oxford</div>
  </div>
  <div><div class="last-updated">Last updated<br><strong>{updated}</strong></div></div>
</header>
<div class="summary-bar">
  <div class="stat-card"><div class="stat-number">{total}</div><div class="stat-label">Total Open</div></div>
  <div class="stat-card"><div class="stat-number">{assigned}</div><div class="stat-label">Assigned</div></div>
  <div class="stat-card"><div class="stat-number red">{len(unassigned)}</div><div class="stat-label">&#9888; Unassigned</div></div>
  <div class="stat-card"><div class="stat-number">{team_count}</div><div class="stat-label">Team Members</div></div>
</div>
<main>
{person_sections_html}
{unassigned_html}
</main>
<footer>Generated by import_osm_report.py &middot; Source: {esc(report_name)}</footer>
</body>
</html>"""

# ── Step 6: GitHub push helpers ──────────────────────────────────────────────────────────────

def get_github_pat():
    pat = os.environ.get("GITHUB_PAT", "").strip()
    if not pat:
        sys.exit(
            "ERROR: GITHUB_PAT environment variable is not set.\n"
            "Set it as a Windows User environment variable (never in any file)."
        )
    return pat

def github_put(api_url, content_bytes, commit_message, pat):
    """GET current SHA, then PUT new content. Returns new SHA."""
    headers = {
        "Authorization": f"token {pat}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }

    resp = requests.get(f"{api_url}?ref={GITHUB_REF}", headers=headers, timeout=30)
    if resp.status_code == 404:
        current_sha = None  # new file
    elif resp.status_code == 200:
        current_sha = resp.json()["sha"]
    else:
        sys.exit(f"ERROR: GET {api_url} failed (HTTP {resp.status_code})\n{resp.text}")

    payload = {
        "message": commit_message,
        "content": base64.b64encode(content_bytes).decode("ascii"),
        "branch":  GITHUB_REF,
    }
    if current_sha:
        payload["sha"] = current_sha

    put_resp = requests.put(api_url, headers=headers, json=payload, timeout=30)
    if put_resp.status_code not in (200, 201):
        sys.exit(f"ERROR: PUT {api_url} failed (HTTP {put_resp.status_code})\n{put_resp.text}")

    return put_resp.json()["content"]["sha"]

# ── Main ──────────────────────────────────────────────────────────────────────────────

def main():
    print("=== import_osm_report.py ===")
    print()

    report_path             = find_report()
    tickets                 = parse_report(report_path)
    by_analyst, unassigned  = group_tickets(tickets)

    total = sum(len(v) for v in by_analyst.values()) + len(unassigned)
    print()
    print("Ticket summary:")
    for name, t in by_analyst.items():
        if t:
            print(f"  {name}: {len(t)}")
    print(f"  Unassigned: {len(unassigned)}")
    print(f"  Total: {total}")
    print()

    pat      = get_github_pat()
    now_str  = datetime.now().strftime("%Y-%m-%d %H:%M")

    # ── Push data/tickets.json (new v2 dashboard data layer) ──
    print("Pushing data/tickets.json...")
    payload   = build_tickets_json(by_analyst, unassigned, report_path)
    json_bytes = json.dumps(payload, indent=2, ensure_ascii=False).encode("utf-8")
    new_sha   = github_put(DATA_API, json_bytes, f"OSM import — {now_str}", pat)
    print(f"  data/tickets.json SHA: {new_sha}")
    print("  Done.")
    print()

    # ── Push index.html (legacy fallback — kept until v2 dashboard is live) ──
    print("Pushing index.html (legacy)...")
    html      = build_html(by_analyst, unassigned, report_path)
    html_sha  = github_put(GITHUB_API, html.encode("utf-8"), f"OSM import — {now_str}", pat)
    print(f"  index.html SHA: {html_sha}")
    print("  Done.")
    print()

    print("All updates pushed successfully.")
    print("Allow ~60 seconds for GitHub Pages to rebuild, then reload:")
    print("  https://begb0037admin.github.io/hris-dashboard/")

if __name__ == "__main__":
    main()
