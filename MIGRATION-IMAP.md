# MIGRATION-IMAP.md — hris-dashboard OSM report fetch: classic Outlook COM → Codex M365 connector

**Status: LIVE. Built, verified, and cut over 9 Sep 2026 (Drew).** The
classic-Outlook-COM path (`fetch_osm_report.py`, run from the desktop) is
retired from the automated morning refresh. `fetch_osm_report_connector.py`
(Microsoft Outlook Email app connector, personal-account CODEX_HOME, same
verb-based re-contamination guard pattern as work-inbox's `lane_b_call1.py`)
now runs on the Oxford laptop (`101L-DE013193`, `AD-OAK\begb0037`), scheduled
task `HRIS Dashboard Morning Refresh (connector)`, triggers 08:45/09:15/09:45
Mon–Fri. The desktop's own `HRIS Dashboard Morning Refresh` scheduled task is
**disabled** (not deleted — the `.bat`/COM script and its Task Scheduler
entry both still exist, one `Enable-ScheduledTask` away from being a manual
rollback if ever needed). `Update HRIS Dashboard.bat` (Kevin's manual
"run now") and the Startup Downloads watcher are **untouched**, exactly as
planned — the desktop manual fallback path is fully intact.

## What actually happened, 9 Sep 2026 (read this before the older sections below — they are the design history, this is the outcome)

Built `fetch_osm_report_connector.py` (hris-dashboard root) and
`Run HRIS OSM Connector Fetch (laptop).ps1` (the scheduled-task wrapper),
both reusing work-inbox's `lane_b_call1.py` connector primitives directly
(imported via `sys.path`, not duplicated — see that script's own docstring
for exactly what's imported vs. locally re-implemented and why). Verified
live, for real, against the real mailbox, same day:

1. **Guard self-test** (`--selftest-guard`, pure-function, no codex) passed
   on both the desktop and the laptop clone.
2. **Live connector fetch, first real attempt, succeeded on the first try**
   (no retries needed): `search_messages` found today's real OSM report
   email, `list_attachments` listed the real `.xls`, `fetch_attachment`
   materialised a presigned download URL, plain `requests.get()` downloaded
   40896 real bytes. No guard trip (no unexpected/off-scope/write tool call
   observed).
3. **Content-integrity verification:** the downloaded file parsed cleanly via
   `import_osm_report.py`'s own `parse_report()`/`group_tickets()` — 39 real
   tickets across 5 named analysts + unassigned, sane distribution, no
   anomalies.
4. **Honest gap, disclosed not hidden:** this was NOT a same-day
   byte-identical diff against the COM path, because the COM path
   (`fetch_osm_report.py`, desktop) genuinely **failed all 3 scheduled
   attempts that same morning** (`data/last_automated_run.json` recorded
   `status: failure, step: fetch_failed`, exhausted 3/3 attempts by 09:45) —
   there was no COM-fetched file for today to diff against. Verification
   instead rests on: (a) the already-documented 8 Sep 2026 proof that this
   exact three-step connector chain produces a sha256-byte-identical file to
   the COM path (see section 2(a) below, unretracted), plus (b) today's live
   content-integrity check (item 3 above) proving the connector path's
   output is genuinely correct data, not merely "a file downloaded without
   erroring."
5. **Used to fix a real live incident, not just tested in isolation:** since
   today's COM path had genuinely failed and the dashboard was showing a
   stale "failure" banner, the connector-fetched file was pushed for real —
   `import_osm_report.py` → `data/tickets.json` (39 tickets, SHA
   `1f5bd8ac5c911a13026ff67bcb9c62dc120f132c`) → `push_automation_status.py
   --status success`. Live dashboard screenshot confirmed: green
   "Last automated refresh: Wednesday 09 September 2026 at 18:27 — up to
   date", source line "All Open Tasks by Team - auto.xls".
   `C:\Users\admin\Documents\Meetings\hris_dashboard_connector_cutover_9sept.png`
6. **Cutover performed same session**, per Kevin's advance go-ahead ("you have
   Kevin's go-ahead for the cutover step itself once verification passes"):
   scheduled task `HRIS Dashboard Morning Refresh (connector)` registered on
   the laptop (triggers 08:45/09:15/09:45 Mon–Fri, `InteractiveToken` logon
   matching work-inbox's own proven task shape — see the registration
   gotcha below), desktop's `HRIS Dashboard Morning Refresh` task disabled.
7. **First real unattended morning run (08:45 tomorrow) has not happened
   yet at the time of writing** — today's proof was a manually-triggered
   live run, not yet a genuine unattended scheduled firing. That is the
   next thing to check, not a re-build.

### Real registration gotcha found this session (useful for any future cross-account scheduled-task work)

Registering an `Interactive`-logon scheduled task for a DIFFERENT Windows
account than the one creating it, via `schtasks /create ... /ru begb0037
/it` (no stored password), silently creates a task that **never actually
fires** — not on-demand (`schtasks /run` / `Start-ScheduledTask`), and not
even on its own natural trigger time. `Last Run Time` stays at the
Windows epoch placeholder (`30/11/1999`) forever; no Task Scheduler
Operational-log event is even written for the attempt. This is NOT a
permissions error surfaced anywhere — it fails completely silently.

**Root cause (confirmed by comparing XML exports):** the already-working
production task (`Work Inbox Bridge Briefing`) has its `<Principal>`
`<UserId>` set to the account's **SID**
(`S-1-5-21-1658931844-4182391637-1812126793-312275`), not the bare account
name. `schtasks /create /ru begb0037 /it` resolves/stores the principal
differently (by name, not SID) and that form does not bind correctly for a
task created by a *different* logged-on account (`begb0037-a`, local admin)
than the one it names. **Fix: export the SID from a known-working task
(`schtasks /query /tn "<task>" /xml`), build the new task's XML with that
exact SID in `<UserId>`, and import via `schtasks /create /tn "<name>" /xml
"<path>" /f`.** Verified twice this session: a throwaway single-shot
`TimeTrigger` test task built this way fired correctly and logged `Last
Result: 0` at its scheduled time; the real production task was then built
the same way. Also note when transferring XML content over SSH from a
non-Windows shell: `schtasks /create /xml` requires the file to actually be
UTF-16LE **with a byte-order mark** — a `<?xml version="1.0"
encoding="UTF-16"?>` declaration alone, converted with `iconv -f utf-8 -t
utf-16le` but with no BOM prepended, fails with a misleading "ERROR: The
task XML is malformed. (1,2)::ERROR: one root element" even though the file
content itself is byte-for-byte correct.

## 1. Why this migration exists

`fetch_osm_report.py` (the automated morning path) connects to Outlook with
`win32com` `Dispatch("Outlook.Application")` — classic Outlook COM. On
8 Sep 2026 this produced a real incident: the client-side Outlook rule that
files `reports-prd-ldz@saasiteu.com`'s daily "Report: All Open Tasks by Team"
email from the top-level Inbox into `Inbox/Reports/OSM` **only runs while
classic Outlook is open**, and that morning only *new* Outlook (`olk.exe`) was
running. The 08:45 and 09:15 scheduled runs both failed "no email today" while
the email sat unfiled in the top-level Inbox.

Same-day fix shipped:
[`ed1ad73`](https://github.com/begb0037admin/hris-dashboard/commit/ed1ad73d3f778c998d7161b27710b2709b734d20)
— `fetch_osm_report.py` now checks `Inbox/Reports/OSM` first, then falls back
to the top-level Inbox; also removed an early-bail bug and made the date
comparison timezone-safe. **That fix is a band-aid — it does not remove the
classic-Outlook dependency**, only the reliance on the subfolder having been
pre-filled by the rule. The machine still has to have classic Outlook COM
reachable, and `Dispatch()` still silently spawns a hidden classic-Outlook
instance on every run.

work-inbox and command-centre already migrated their entire mail path off
classic Outlook COM (to IMAP + OAuth2, stdlib only, zero win32com) running on
the Oxford laptop. hris-dashboard's OSM fetch is the last classic-Outlook-COM
dependency in Kevin's estate.

---

## 2. The three options weighed (8 Sep 2026)

All three were checked against live systems this session, not assumed.

### (a) ChatGPT/Codex `microsoft_outlook_email.*` connector, from the desktop

The same `codex_apps` connector work-inbox uses for calendar/Teams.

- **Verified working end to end this session.** `codex exec` (read-only) on
  `DESKTOP-MJDJM64` against `kevin.lelitte@admin.ox.ac.uk`:
  `search_messages` found today's real SAASIT email → `list_attachments`
  returned the real `.xls` metadata → `fetch_attachment` materialised it to a
  presigned Azure-blob URL (`sdmntpreastus2.oaiusercontent.com`, ~5 min TTL) →
  plain `curl` of that URL returned 32704 bytes, **sha256 byte-identical** to
  the file classic-Outlook COM fetched the same morning.
- Connector namespace is real: 46 `microsoft_outlook_email.*` tools incl.
  `search_messages`, `list_messages`, `list_attachments`, `fetch_attachment`
  (read) plus many write tools (`send_email`, `reply_to_email`, `move_email`,
  `mark_email_read_state`, `set_message_categories`, …).
- **Rejected as primary because:**
  - Per-call tool loading is flaky — ~8 headless `codex exec` attempts over
    ~10 minutes before the connector tools loaded in a session (matches
    work-inbox's documented "availability FLIPS run-to-run"). Needs the same
    retry/backoff + warm-up scaffolding `lane_b_call1.py` already has.
  - Two-step attachment retrieval (materialise → short-TTL URL → HTTP GET),
    more moving parts than a direct fetch.
  - Adds `microsoft_outlook_email` to the connector's trusted namespace set —
    widens the accepted worst case from "a stray Teams message" to "a stray
    email action". Write verbs still HALT under the existing verb-based guard,
    but this needs Kevin's own fresh explicit risk acceptance, not an
    inherited one.
  - Runs on Kevin's personal ChatGPT plan quota (small incremental, non-zero).
  - Depends on `codex` CLI + ChatGPT session health.
- **Kept as the fallback option** if Kevin later wants everything on one
  machine and no laptop scheduled task.

### (b) IMAP from the desktop, with a fresh Oxford OAuth consent

- **Rejected.** `DESKTOP-MJDJM64` (`dsregcmd /status`, checked live):
  `AzureAdJoined NO`, `DomainJoined NO`, `AzureAdPrt NO`; only
  `WorkplaceJoined` to the `lelitte.com` tenant, *not* Oxford. This is the
  "no-PRT admin desktop" work-inbox's own migration record identified as an
  outage cause.
- MSAL broker/WAM path is dead for the Thunderbird public client (proven on
  the laptop). The plain interactive system-browser flow works, but with no
  PRT every periodic (~14–90 day) re-auth would require a full Oxford
  username + password + MFA, not a silent SSO click. Day-to-day scheduled
  runs would still be silent via the cached token.
- Most new setup, on the machine least suited to it. Nothing recommends it
  over (a) or (c).

### (c) IMAP from the Oxford laptop, reusing work-inbox's token cache — **CHOSEN**

- **Auth: nothing new.** Reuse
  `%LOCALAPPDATA%\WorkInboxAI\msal_imap_token_cache.bin` on the laptop
  (`101l-de013193` / `AD-OAK\begb0037`). Same mailbox
  (`kevin.lelitte@admin.ox.ac.uk`), same public client id (`9e5f94bc-...`,
  Mozilla Thunderbird), same scope
  (`https://outlook.office365.com/IMAP.AccessAsUser.All`). Laptop holds a
  healthy Oxford PRT → silent day-to-day, one browser click (~22s) on the
  periodic re-auth.
- **Delivery: nothing to bridge.** hris-dashboard's OSM path is already
  *local script → GitHub Contents API push → GitHub Pages auto-serves*.
  `index.html` fetches `data/tickets.json` client-side from the repo. The
  self-hosted GitHub Actions runner on the desktop is used *only* by the
  separate, currently-dead SAASIT Playwright scrape
  (`generate_dashboard.py`), never by the email path. So the laptop just runs
  fetch → `import_osm_report.py` → push, exactly as work-inbox's laptop
  pipeline already pushes `data/briefing.json`. No file copy, no network
  share, no runner.
- **Reliability: laptop confirmed hardened for the morning window** (checked
  live over SSH, 8 Sep): AC *and* DC "sleep after" and "hibernate after" all
  `0` (never); wake-from-sleep history count 0 (it doesn't sleep); uptime
  since 3 Sep (5 days, not rebooted); on AC at 100%; and three work-inbox
  scheduled tasks (`Work Inbox Bridge Briefing` 07:00, `Work Inbox Laptop
  Draft Diff` 07:30, `Work Inbox Laptop Parity Shadow` 07:00) all `Ready` and
  succeeding every weekday. Adding an 08:45/09:15/09:45 hris task drops into a
  slot already proven three times over — currently more reliable than the
  desktop (desktop uptime was only 2 days, and it is Kevin's interactive
  machine).
- **Residual risks (real, small):**
  - The laptop is portable — it *can* be undocked / carried / lid-closed /
    genuinely powered off in a way a stationary desktop cannot. Mitigated by:
    the desktop keeps the working manual fallback; the laptop's own power
    config already fights sleep. Lid-close power action could not be read
    cleanly over SSH this session — **confirm at build time**.
  - Splits hris-dashboard automation across two machines (laptop = automated
    morning fetch; desktop = manual `Update HRIS Dashboard.bat` + Downloads
    watcher). The desktop fallback staying put is a feature.

---

## 3. Target chain (what to build)

```
Oxford laptop 101l-de013193  (user AD-OAK\begb0037, PRT-holding standard user)
│
├─ NEW scheduled task  "HRIS Dashboard Morning Refresh (laptop)"
│     triggers 08:45 / 09:15 / 09:45 Mon–Fri  (mirror Run_HRIS_Auto_Refresh.bat)
│     retry-guard: skip if data/last_automated_run.json already shows today = success
│
├─ NEW  fetch_osm_report_imap.py
│     • MSAL silent token off %LOCALAPPDATA%\WorkInboxAI\msal_imap_token_cache.bin
│       (client 9e5f94bc-..., scope outlook.office365.com/IMAP.AccessAsUser.All)
│       — reuse imap_mail.py's acquire pattern; ImapReauthRequired → toast + exit clean
│     • IMAP EXAMINE (read-only) outlook.office365.com:993, SASL XOAUTH2
│     • SEARCH the top-level INBOX by  FROM reports-prd-ldz@saasiteu.com
│       SUBJECT "Report: All Open Tasks by Team"  SINCE <today>
│       ── NO Inbox/Reports/OSM subfolder dependency at all ──
│     • parse the matched message with email.*, extract the
│       "All Open Tasks by Team*.xls|.xlsx" attachment bytes (get_payload(decode=True))
│     • write to C:\...\Downloads\All Open Tasks by Team - auto.xls|.xlsx
│       (same filename convention import_osm_report.py already globs; keep the
│        "- auto" marker so the desktop Downloads watcher still ignores it)
│     • exit non-zero + clear message if no matching email today (unchanged contract)
│
├─ import_osm_report.py  (existing, unchanged — runs fine on the laptop; Python + openpyxl/xlrd)
│     → parses the .xls/.xlsx → PUT data/tickets.json to GitHub (Contents API)
│
└─ push_automation_status.py  (existing, unchanged)
      → PUT data/last_automated_run.json  status=success/failure

GitHub Pages rebuilds from data/tickets.json → https://begb0037admin.github.io/hris-dashboard/
```

Desktop (`DESKTOP-MJDJM64`) after cutover:
- `Run_HRIS_Auto_Refresh.bat` scheduled task → **disabled** (file kept).
- `Update HRIS Dashboard.bat` (Kevin's manual "run now") → **unchanged, kept**.
- `watch_downloads.py` Startup watcher → **unchanged, kept**.

---

## 4. Open items to resolve at build time

1. **Laptop `GITHUB_PAT` scope.** A User-scope `GITHUB_PAT` exists on the
   laptop but reads **length 16** — too short to be a real GitHub PAT.
   Confirm what it is; it must be (or be replaced with) a token with
   `contents:write` on `begb0037admin/hris-dashboard`. Never write the value
   to a file — Windows User env var only (estate rule).
2. **Laptop lid-close power action.** Could not be read cleanly over SSH this
   session. Must be "Do nothing" on AC (`powercfg /setacvalueindex
   SCHEME_CURRENT SUB_BUTTONS LIDACTION 0` equivalent) or an 08:45 run can be
   missed when the lid is shut.
3. **Drop the `Inbox/Reports/OSM` subfolder dependency.** Search the top-level
   INBOX directly by sender + subject + date via IMAP `SEARCH`. This removes
   the classic-Outlook rule-filing race at the root rather than patching it
   (the 8 Sep `ed1ad73` fallback becomes unnecessary once on IMAP).
4. **Parity / dry-run period before cutover.** Run the laptop IMAP fetch
   alongside the live desktop COM path for several scheduled cycles —
   log-only or pushing to a scratch path — and diff the produced
   `tickets.json` against the COM path's output. Cut over only on a clean
   diff run **and** Kevin's fresh explicit go-ahead (work-inbox
   cautious-change-pace discipline).
5. **Cutover step.** Disable the desktop `Run_HRIS_Auto_Refresh.bat` task,
   enable the laptop task, screenshot the dashboard after the first real
   laptop-driven run for Kevin's approval.
6. **`import_osm_report.py` on the laptop.** Confirm `openpyxl` + `xlrd` +
   `requests` are present in the laptop's Python (work-inbox's laptop Python
   is 3.12; deps likely need `pip install`). The pinned-SHA mechanism the
   desktop `.bat` uses can be reused, or the laptop task can run a local
   pinned copy — decide at build time.
7. **Guard/telemetry parity.** Keep the estate timestamp-on-first-line rule;
   keep `push_automation_status.py`'s failure signalling so the dashboard
   banner still reflects a failed laptop run.

---

## 5. What is NOT changing

- The dashboard itself (`index.html`, GitHub Pages) — untouched.
- `import_osm_report.py` parsing / `tickets.json` schema — untouched.
- The desktop manual fallback (`Update HRIS Dashboard.bat` + Downloads
  watcher) — untouched, kept as belt-and-braces.
- The separate, long-dead GitHub Actions SAASIT Playwright scrape
  (`generate_dashboard.py` / the "Refresh" button) — still needs Kevin's
  Oxford SSO re-login via `Refresh Session.bat`; out of scope for this
  migration.
