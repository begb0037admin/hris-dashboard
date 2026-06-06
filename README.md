## Claude Quick Load

Paste any URL below directly into Claude chat to load project context:

| File | Raw URL |
|---|---|
| `CLAUDE.md` | https://raw.githubusercontent.com/begb0037admin/hris-dashboard/main/CLAUDE.md |

---

# HRIS Dashboard — Project Documentation & Recovery Guide
**Last updated:** June 2026
**Owner:** Kevin Lelitte (begb0037)
**Purpose:** Daily auto-updating HTML dashboard showing HRIS team open tickets from Ivanti SAASIT, published to GitHub Pages.

**Live site:** https://begb0037admin.github.io/hris-dashboard/

---

## Live URLs
| What | URL |
|------|-----|
| Live dashboard | https://begb0037admin.github.io/hris-dashboard |
| GitHub repo | https://github.com/begb0037admin/hris-dashboard |

---

## What It Does
- Fetches open tickets from Ivanti SAASIT (oxford.saasiteu.com) via OData API
- Shows tickets per team member plus unassigned tickets
- Rebuilds and publishes the dashboard automatically every hour, 8am-5pm
- Runs silently in the background via Windows Task Scheduler
- No server, no hosting costs - GitHub Pages serves the HTML for free

---

## Files - Where Everything Lives
All files live at: C:\Users\admin\Documents\Claude\HRIS-Dashboard\

| File | Purpose |
|------|---------|
| generate_dashboard.py | Main script - fetches data, builds HTML, pushes to GitHub |
| login.py | One-time login script - run when session expires |
| session.json | Saved browser session (created by login.py) - DO NOT delete |
| dashboard_run.log | Log of every script run - check here if dashboard goes empty |
| hris-dashboard-local\ | Local git clone of the GitHub repo (auto-created) |

---

## Team Members & Usernames
| Username | Name |
|----------|------|
| musf0100 | James Salas Guillen |
| ouit0422 | Asta Palmer |
| ouit0036 | Michael O'Sullivan |
| admn2716 | Simon Burford |
| begb0037 | Kevin Lelitte |

---

## Day-to-Day Operation
Nothing to do. The dashboard updates itself every hour from 8am-5pm.
Just open the live URL. It is always current as of the last hourly run.

### Session Expiry - The One Thing That Needs Attention
The session.json file holds your Oxford SSO login session. It will expire eventually (days to weeks). When it does, the dashboard will show 0 tickets.

Fix (2 minutes):
  cd C:\Users\admin\Documents\Claude\HRIS-Dashboard
  python login.py
A browser window opens. Log in with Oxford SSO (Microsoft MFA). Once you see the SAASIT dashboard, press Enter. Done - automated again.

How to know the session has expired:
- Dashboard shows 0 tickets across all team members
- Check dashboard_run.log - look for "Got 0 session cookies" or 401 errors

---

## Scheduled Task
Task name: HRIS Dashboard
Schedule: Every hour, 8:00am-5:00pm, every day
Runs as: admin (highest privileges)
Working directory: C:\Users\admin\Documents\Claude\HRIS-Dashboard

To verify: schtasks /query /tn "HRIS Dashboard"
To run now: schtasks /run /tn "HRIS Dashboard"

To recreate:
  schtasks /delete /tn "HRIS Dashboard" /f
  schtasks /create /tn "HRIS Dashboard" /tr "python C:\Users\admin\Documents\Claude\HRIS-Dashboard\generate_dashboard.py" /sc hourly /mo 1 /st 08:00 /et 17:00 /rl highest /f

---

## Machine Requirements
| Requirement | Status |
|-------------|--------|
| Windows machine, stays on 24/7 | Required |
| Never sleeps on AC power | Confirmed (powercfg = never) |
| Internet access | Required |

---

## Auth - How It Works
Oxford SSO uses Shibboleth -> Microsoft (oxford.ac.uk tenant) with MFA.
browser_cookie3 does NOT work (Chrome 127+ App-Bound Encryption).
Solution: Playwright saves a full browser state snapshot (session.json) after manual login.

GitHub PAT: Embedded in generate_dashboard.py at top as GITHUB_TOKEN.
If PAT expires, pushes will fail. Generate a new one at https://github.com/settings/tokens (repo scope).

---

## Full Rebuild - New Machine Setup

1. Install Python 3.11+ (tick "Add to PATH")
2. pip install playwright requests gitpython
3. python -m playwright install chromium
4. Copy generate_dashboard.py and login.py from this repo
5. Update GITHUB_TOKEN in generate_dashboard.py with a fresh PAT
6. python login.py (log in with Oxford SSO, press Enter when SAASIT visible)
7. python generate_dashboard.py (check log for success)
8. Recreate scheduled task (command above)
9. powercfg /change standby-timeout-ac 0

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---------|-------------|-----|
| Dashboard shows 0 tickets | Session expired | Run login.py |
| 401 errors in log | Session expired | Run login.py |
| 400 errors in log | API issue | Check dashboard_run.log |
| Dashboard not updating | Task Scheduler stopped | schtasks /query /tn "HRIS Dashboard" |
| Git push failing | PAT expired | Update GITHUB_TOKEN in generate_dashboard.py |
| session.json not found | File deleted or wrong dir | Run login.py from HRIS-Dashboard folder |

---

## Development History
| Approach | Result |
|----------|--------|
| browser_cookie3 - read live Chrome cookies | Failed - Chrome 127+ App-Bound Encryption |
| Playwright cookie-only cache | Failed - SSO tokens short-lived |
| Playwright full storage_state (current) | Working |

---

*Generated by Claude (Hope profile) - Lelitte Co. AIMM chain*
