"""
import_osm_report.py — OSM Excel → HRIS Dashboard

Usage:
    python import_osm_report.py

Reads the most recently modified "All Open Tasks by Team*" report
(.xlsx or legacy .xls) from C:\\Users\\admin\\Downloads and pushes
data/tickets.json to GitHub. The v2 dashboard (index.html) is a static
file — it fetches tickets.json dynamically and must not be overwritten
by this script.

OSM Administration's export format is not consistent — some cycles
arrive as modern .xlsx, others as genuine old-format .xls (OLE binary,
not just a renamed extension). Both are handled natively here: no
manual conversion step, no LibreOffice dependency. .xlsx is read with
openpyxl; .xls is read with xlrd (xlrd 2.x dropped .xlsx support but
still reads legacy .xls fine — the two libraries are complementary,
not a version choice).

Requires:
    pip install openpyxl xlrd requests
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
    import xlrd
except ImportError:
    sys.exit("ERROR: xlrd not installed (needed for legacy .xls reports). Run: pip install xlrd")

try:
    import requests
except ImportError:
    sys.exit("ERROR: requests not installed. Run: pip install requests")

# ── Config ────────────────────────────────────────────────────────────────────────────────────

DOWNLOADS     = r"C:\Users\admin\Downloads"
FILE_PATTERNS = [
    "All Open Tasks by Team*.xlsx",
    "All Open Tasks by Team*.xls",
]
GITHUB_REPO  = "begb0037admin/hris-dashboard"
GITHUB_REF   = "main"
BASE_API     = f"https://api.github.com/repos/{GITHUB_REPO}/contents"
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

# ── Step 1: find the report file ────────────────────────────────────────────

def find_report():
    files = []
    for pattern in FILE_PATTERNS:
        files.extend(glob.glob(os.path.join(DOWNLOADS, pattern)))
    if not files:
        patterns_str = " / ".join(FILE_PATTERNS)
        sys.exit(f"ERROR: No file matching '{patterns_str}' found in {DOWNLOADS}")
    latest = max(files, key=os.path.getmtime)
    print(f"Found report: {os.path.basename(latest)}")
    print(f"  Modified:   {datetime.fromtimestamp(os.path.getmtime(latest)).strftime('%d %b %Y %H:%M')}")
    return latest

# ── Step 2: read raw rows, format-agnostic ──────────────────────────────────
#
# Both readers normalise to the same shape openpyxl's
# `ws.iter_rows(values_only=True)` produces: a list of row-tuples, each
# cell either None, a str, an int/float, or a datetime. Everything
# downstream (header detection, column mapping, per-cell parsing) is
# then format-agnostic and untouched by which reader was used.

def _read_rows_xlsx(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    return list(ws.iter_rows(values_only=True))

def _read_rows_xls(path):
    wb = xlrd.open_workbook(path)
    ws = wb.sheet_by_index(0)
    rows = []
    for r in range(ws.nrows):
        row = []
        for c in range(ws.ncols):
            cell = ws.cell(r, c)
            ctype = cell.ctype
            value = cell.value
            if ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                value = None
            elif ctype == xlrd.XL_CELL_DATE:
                try:
                    value = xlrd.xldate_as_datetime(value, wb.datemode)
                except (ValueError, OverflowError):
                    pass  # leave raw serial if it isn't a valid date
            elif ctype == xlrd.XL_CELL_NUMBER:
                if value == int(value):
                    value = int(value)
            elif ctype == xlrd.XL_CELL_BOOLEAN:
                value = bool(value)
            elif ctype == xlrd.XL_CELL_ERROR:
                value = None
            row.append(value)
        rows.append(tuple(row))
    return rows

def read_rows(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsx":
        return _read_rows_xlsx(path)
    if ext == ".xls":
        return _read_rows_xls(path)
    sys.exit(f"ERROR: Unsupported file extension '{ext}' for {os.path.basename(path)}")

# ── Step 3: parse the rows into tickets ───────────────────────────────

def parse_report(path):
    rows = read_rows(path)

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

# ── Step 4: group tickets ─────────────────────────────────────────────────────

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

# ── Step 5: build tickets.json ────────────────────────────────────────────────────────────

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

# ── Step 6: GitHub push helpers ──────────────────────────────────────────────────────────

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
        current_sha = None
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

# ── Main ────────────────────────────────────────────────────────────────────────────────────

def main():
    print("=== import_osm_report.py ===")
    print()

    report_path            = find_report()
    tickets                = parse_report(report_path)
    by_analyst, unassigned = group_tickets(tickets)

    total = sum(len(v) for v in by_analyst.values()) + len(unassigned)
    print()
    print("Ticket summary:")
    for name, t in by_analyst.items():
        if t:
            print(f"  {name}: {len(t)}")
    print(f"  Unassigned: {len(unassigned)}")
    print(f"  Total: {total}")
    print()

    pat     = get_github_pat()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

    print("Pushing data/tickets.json...")
    payload    = build_tickets_json(by_analyst, unassigned, report_path)
    json_bytes = json.dumps(payload, indent=2, ensure_ascii=False).encode("utf-8")
    new_sha    = github_put(DATA_API, json_bytes, f"OSM import — {now_str}", pat)
    print(f"  data/tickets.json SHA: {new_sha}")
    print("  Done.")
    print()

    print("All updates pushed successfully.")
    print("Allow ~60 seconds for GitHub Pages to rebuild, then reload:")
    print("  https://begb0037admin.github.io/hris-dashboard/")

if __name__ == "__main__":
    main()
