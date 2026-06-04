# HRIS Dashboard — Roadmap

## Shipped

### GitHub Actions Migration — COMPLETE (4 June 2026)

Migrated from Windows Task Scheduler to GitHub Actions with a self-hosted runner on DESKTOP-MJDJM64.

- Self-hosted runner installed as Windows service at C:\\actions-runner\\
- Workflow updated: runs-on self-hosted, Python setup removed, commit step in PowerShell
- Dashboard confirmed working from home without VPN (Oxford MDM-enrolled)
- Runner auto-starts on boot

## Ongoing

- Dashboard auto-updates hourly 8am-5pm via GitHub Actions
- Session refresh via Refresh Session.bat when Oxford SSO expires (~every few weeks)

## Backlog

- Nothing currently logged

---
Last updated: 4 June 2026 — Lelitte Co.
