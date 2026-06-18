## Claude Quick Load

Paste any URL below directly into Claude chat to load project context:

| File | Raw URL |
|---|---|
| `CLAUDE.md` | https://raw.githubusercontent.com/begb0037admin/hris-dashboard/main/CLAUDE.md |

---

# HRIS Dashboard

Automated SAASIT HR dashboard for Kevin Lelitte — University of Oxford.

Scrapes live headcount, vacancy, and compliance data from the SAASIT portal → publishes a structured HTML dashboard to GitHub Pages via GitHub Actions.

**Owner:** Kevin Lelitte, HR Systems Manager/Director
**Repo:** https://github.com/begb0037admin/hris-dashboard

## How It Works

| Component | Detail |
|---|---|
| `generate_dashboard.py` | Scrapes SAASIT via saved session → writes `index.html` |
| `login.py` | Refreshes SAASIT session — run when workflow goes red |
| `.github/workflows/update-dashboard.yml` | Cron: hourly Mon–Fri 7–16 UTC. Self-hosted runner on DESKTOP-MJDJM64 |
| `cloudflare-worker/` | Proxy for the Refresh button (holds PAT server-side) |

## Day-to-Day

Nothing to do. Just refresh your browser.

**If workflow runs go red:** session has expired. Run `Refresh Session.bat` from `C:\Users\admin\Documents\Claude\Projects\HRIS-Dashboard\` and log in via Oxford SSO.
