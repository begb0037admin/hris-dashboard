# RESUME.md — hris-dashboard

This file did not exist before 20 Aug 2026 — created that session per
`agent-commons/SESSION_PROTOCOL.md` (every project needs a durable
resume/state record; this repo previously relied on `CLAUDE.md` +
`HANDOVER.md` only). Keep this updated at every meaningful stop
alongside `HANDOVER.md`.

**Owning agent (as of 20 Aug 2026):** Drew (`begb0037admin/drew`), added to
scope by Kevin same day. Drew was not this repo's "usual" agent before
11 Aug 2026 — see `HANDOVER.md` session log for the first Drew touch
(GitHub Actions schedule-trigger fix).

---

## One-line resume (latest — 27 Aug 2026, ~11:40 BST)

**Follow-up to the ~11:00 build. Coordinator asked whether bracketed browser
filenames (`All Open Tasks by Team (1).xlsx`, `(2)`, no-space `Team(1).xlsx`,
etc.) break either path, so Kevin can stop deleting the old file before each
download. Verified precisely.**

- **Manual `.bat` / `import_osm_report.py`: no change needed.** `(` `)` are not
  `glob` metacharacters; verified `glob` matches every `(N)` variant (space and
  no-space), `openpyxl` reads paren paths (44 tickets, byte-identical data),
  `OSM_IMPORT_FILE` / `--file` / bare-argv overrides all accept paren paths, the
  advisory lock path is filename-independent, console output is ASCII.
  `find_report()` already does `max(..., key=os.path.getmtime)` so the freshest
  file wins — an old un-bracketed leftover is never picked over a fresh
  bracketed one.
- **Watcher: one real weakness found and fixed** ([`4387ba1`](https://github.com/begb0037admin/hris-dashboard/commit/4387ba1f79030da1e2631aab24fe16a89c98df6a) →
  merge [`caf92bf`](https://github.com/begb0037admin/hris-dashboard/commit/caf92bfc197576acfae40aade9b74773d84ed218)).
  `scan_once()` was importing *every* un-seeded matching file that settled in a
  pass, in non-deterministic `glob` order — so an un-deleted leftover could be
  the last import and overwrite a fresh bracketed download. Now **newest-file-by-
  mtime only**: only the single newest matching export is ever imported (same
  rule as `find_report()`); older matching files, once settled, are recorded
  `"superseded"` and never re-examined. Dedupe changed to **content-hash only**,
  so identical export bytes are never re-imported whatever the file was named.
  `import_osm_report.py` unchanged ⇒ **no SHA-pin bump** (`SCRIPT_SHA` /
  `script_sha` still `bd93285`).
- **Tested (stubbed importer, temp dirs):** old plain leftover + fresh `(1)` →
  only `(1)` imported, leftover superseded ✓; 5 mixed leftovers on a cold start
  → only the newest imported, 4 retired ✓; `(3)`→`(4)` re-download → `(4)`
  imported ✓; identical-bytes `(4)`→`(5)` → sha dedupe, skipped ✓; morning
  `- auto` present *and newer* → still excluded, manual `(1)` wins ✓; quiet-hours
  defer then import ✓; seed then fresh `(1)` → only `(1)` ✓. Original settle=2 /
  changed-bytes-reimport / exclusion regression all still pass.
- **Deployed:** `watch_downloads.py` re-pulled to `%LOCALAPPDATA%\hris-downloads-watcher\`;
  running watcher restarted (old pid 34812 killed, now **pid 38824**, single
  instance, idle, seeded file skipped). Live `tickets.json` unchanged
  (Kevin's 10:20 commit). No dashboard-visible change.

**Net effect for Kevin:** he can just download the fresh export and leave the
old file in Downloads — both the `.bat` and the watcher take the newest, brackets
or not. Deleting first still works fine; it's simply no longer required.

---

## One-line resume — 27 Aug 2026, ~11:00 BST

**Kevin approved Option 1 (harden the import script + re-sync his Desktop
`.bat`) AND Option 3 (Startup-folder Downloads watcher). Both built, tested,
merged to `main`, and deployed. The watcher is LIVE now (pid recorded in
`%LOCALAPPDATA%\hris-downloads-watcher\watcher.pid`). No dashboard-visible
change; `tickets.json` output is byte-identical to before. Nothing outstanding
for Kevin except optional: confirm it works on his next real mid-day export.**

### What shipped (all on `main`)
| Commit | What |
|---|---|
| [`207f9c9`](https://github.com/begb0037admin/hris-dashboard/commit/207f9c9a8a9cd3f0de1824380339ecc81e0feab5) | hardened `import_osm_report.py` + new `watch_downloads.py` / `watch_downloads.config.json` / `Start HRIS Downloads Watcher.vbs` (branch) |
| [`bd93285`](https://github.com/begb0037admin/hris-dashboard/commit/bd9328521e52123f419ee573b149027fdadc0215) | merge of that branch to `main` |
| [`aa77b25`](https://github.com/begb0037admin/hris-dashboard/commit/aa77b25f635e25779a63caf0ae838f4bbb7981c6) | `Update HRIS Dashboard.bat` SCRIPT_SHA bumped to `bd93285` |
| [`b769903`](https://github.com/begb0037admin/hris-dashboard/commit/b7699032782b8513b01e5c5f626a22eec69a43f5) | `watch_downloads.config.json` `script_sha` set to `bd93285` |

Branch `harden-import-and-downloads-watcher` merged and deleted (repo protocol:
never leave files on a branch).

### Option 1 — `import_osm_report.py` hardening (schema/parsing UNCHANGED)
- **Lazy `xlrd`.** Top-level `import xlrd` removed; imported only inside
  `_read_rows_xls()`. An `.xlsx`-only run no longer dies if `xlrd` is absent.
- **Locked / partial file handling.** `read_rows_with_retry()` catches
  `PermissionError` / `zipfile.BadZipFile` / lock-style `OSError`, prints
  "the export file appears to still be downloading, or is open in Excel",
  retries 3× at 4 s, then exits with a clear message (no scary traceback, no
  push). A genuinely malformed file still fails loudly.
- **Explicit-file override.** `OSM_IMPORT_FILE=<path>` / `--file <path>` / a
  bare `argv[1]` path bypasses the newest-match guess entirely (used by the
  watcher; also gives Kevin a drag-a-file-onto-it option for free).
- **Stale-file gate.** If the newest matching export is > 20 min old: prompt
  `[y/N]` when interactive; print a warning and proceed when not (so the
  Scheduled Task and watcher never hang). `is_noninteractive()` = env
  `OSM_IMPORT_NONINTERACTIVE`, `--yes`/`-y`, or stdin not a TTY.
- **Advisory single-writer lock** (`%TEMP%\hris_osm_import.lock`, auto-stale
  15 min, waits ≤ 90 s then proceeds anyway — can never deadlock). Every
  update path runs this script, so one lock covers manual `.bat` + morning
  task + watcher.
- **Push race-retry.** `github_put()` on HTTP 409/422 re-GETs the file SHA and
  retries the PUT once. Concurrent imports self-heal; `tickets.json` is always
  one complete valid import.
- **Timestamped first line** of output (estate-wide requirement).
- All printed strings scrubbed to ASCII (`--` not `—`) so a legacy Windows
  console codepage cannot `UnicodeEncodeError`-crash the run mid-import. The
  commit-message em-dash (`OSM import — <date>`) is kept (goes to the API as
  UTF-8, matches existing history).
- **Verified:** re-parsed the live `C:\Users\admin\Downloads\All Open Tasks by
  Team.xlsx` through the hardened functions — `tickets.json` payload
  byte-identical to the pre-change output (44 tickets, same summary, same
  per-analyst split). 6 unit-style checks pass (lazy xlrd, explicit file,
  missing file, stale gate non-interactive, transient-error retry, loud
  failure). Watcher-style invocation (`python <pinned> ` + env, stdin
  DEVNULL) parses cleanly and stops at the PAT guard.

### Option 3 — Downloads watcher
- `watch_downloads.py` (stdlib only) polls `C:\Users\admin\Downloads` every
  60 s. When a `All Open Tasks by Team*.xlsx|xls` file is size-stable across 2
  polls (~2 min) and its `name+size+sha256` is not in the state file, it runs
  the **pinned** `import_osm_report.py` with `OSM_IMPORT_FILE` +
  `OSM_IMPORT_NONINTERACTIVE=1`. Kevin's mid-day step becomes: save the export
  to Downloads.
- **Never collides with "HRIS Dashboard Morning Refresh":** (a) hard-excludes
  any name containing `- auto` (the Scheduled Task's file is
  `All Open Tasks by Team - auto.xls`); (b) fully idle during the weekday
  08:40–10:00 window — a file that appears then is held and imported once the
  window clears, not dropped; (c) the advisory lock above; (d) the push
  race-retry above.
- **How a watcher-triggered import reaches `main`: AUTO, directly** — the same
  single `data/tickets.json` Contents-API commit the manual `.bat` and the
  morning Scheduled Task already make unattended. Not staged, not a PR. This is
  intended: it is the identical mechanism the morning task already uses 3× per
  weekday with no human in the loop. Safe/revertible because: only fires on a
  genuinely new settled file; dedupe never repeats bytes; hardened script
  (lock + retry + race-retry); `index.html` is never touched by this path;
  revert = `git revert <sha>` on that one JSON commit (Pages rebuilds
  ~60–90 s). **Kill switch:** set `"enabled": false` in
  `watch_downloads.config.json` (watcher exits within one poll) or delete the
  Startup `.vbs` and `taskkill` the `pythonw`.
- `--seed` mode records already-present files so a fresh deploy does not
  re-import them.
- **Tested (sandbox, stubbed importer, temp dirs):** exclude `- auto` ✓;
  quiet-hours weekday/ weekend/ boundary ✓; settle → import only after 2
  stable polls ✓; dedupe (same bytes never re-imported) ✓; changed bytes →
  re-import ✓; quiet-hours defer-then-import ✓; `--seed` then no re-import ✓;
  state persisted ✓. `ensure_pinned_script()` fetched `import_osm_report.py`
  @ `bd93285`, validated non-empty + `def main()` present (guards the
  silent-empty-content trap).

### Deployed on this machine (DESKTOP-MJDJM64), 27 Aug ~10:56–11:00
- `%LOCALAPPDATA%\hris-downloads-watcher\` ← `watch_downloads.py`,
  `watch_downloads.config.json`, `Start HRIS Downloads Watcher.vbs`,
  `import_osm_report.pinned.py` (+ `.sha`), `state.json` (seeded with the
  current `All Open Tasks by Team.xlsx`), `watcher.log`, `watcher.pid`.
- Startup: `…\Start Menu\Programs\Startup\Start HRIS Downloads Watcher.vbs`
  (runs `pythonw watch_downloads.py` hidden at logon; watcher self-enforces
  single instance).
- Desktop `Update HRIS Dashboard.bat` replaced with the `main` copy
  (SCRIPT_SHA `bd93285`); old one backed up as
  `Update HRIS Dashboard.bat.bak-20260827`. Verified byte-identical to `main`.
- Watcher started (`pythonw`, detached); `watcher.log` shows clean start,
  correct config, `script_sha=bd93285`; one full poll cycle observed with no
  import (seeded file correctly skipped); live `tickets.json` unchanged
  (still Kevin's 10:20 commit); `main` HEAD unchanged. Exactly one `pythonw`
  process running (matches `watcher.pid`).

### MAINTENANCE — SHA pin is now in TWO places, bump them together
When `import_osm_report.py` changes on `main` and the change needs to ship:
1. `set SCRIPT_SHA=<new commit sha>` in `Update HRIS Dashboard.bat`, then
   re-copy that file to `D:\OneDrive - lelitte.com\Desktop\Update HRIS Dashboard.bat`.
2. `"script_sha": "<same sha>"` in `watch_downloads.config.json`, then re-copy
   it to `%LOCALAPPDATA%\hris-downloads-watcher\watch_downloads.config.json`
   (the running watcher re-reads the config every poll and re-fetches the
   pinned script automatically once the sha changes).
Miss either and that path keeps silently running the old script. `grep -rn
"bd9328521e52\|SCRIPT_SHA\|script_sha"` to find every pin.

### Morning Scheduled Task — untouched
`Run_HRIS_Auto_Refresh.bat` still pulls `Update HRIS Dashboard.bat` fresh from
`main` each run, so it automatically picks up the new pin. Not modified. Next
run 28 Aug 08:45.

### Separate pre-existing breakage — unchanged
GitHub Actions SAASIT scrape (`generate_dashboard.py`) still fails
`ISM_4001 session expired` — needs Kevin's Oxford SSO+MFA re-login via
`Refresh Session.bat`. Not in scope for this session.

### Exact next action / owner
- **Kevin (optional):** on the next real mid-day export, just save it to
  Downloads and confirm the dashboard updates within ~3–4 min (check
  `%LOCALAPPDATA%\hris-downloads-watcher\watcher.log`). No action needed
  otherwise.
- **Nothing outstanding.** No live-site push pending. Watcher live, `.bat`
  re-synced, docs updated.

### Rollback
- Tooling only. To undo: `git revert b769903 aa77b25 bd93285` (restores the
  20 Aug pinned `.bat` + removes the watcher files from the repo), set
  `"enabled": false` in the deployed config, delete the Startup `.vbs`,
  `taskkill /PID <watcher.pid>`. `data/tickets.json` was never touched by this
  work.

---

## One-line resume — 27 Aug 2026, ~10:30 BST


**Kevin reported "Update HRIS Dashboard.bat is not working" and asked for the
dashboard to be updated from `C:\Users\admin\Downloads\All Open Tasks by Team.xlsx`.
Investigated live (Drew, via coordinator). Finding: the fresh export was ALREADY
live before investigation started — no push was needed. No hard failure of the
`.bat` could be reproduced; his own run succeeded. Real underlying issue is a
stale Desktop copy of the `.bat` plus fragile file-selection logic. "Better way"
proposal recorded below for Kevin's decision — nothing built yet.**

### Data status — already current, verified, no push made
- Live `data/tickets.json` on `main` = commit `7b50878` "OSM import — 2026-08-27 10:20",
  author "kevin lelitte" (i.e. `import_osm_report.py`'s own GitHub PUT — Kevin's
  manual `.bat` run at 10:19:58 BST today).
- `report_file: All Open Tasks by Team.xlsx`; summary `total 44 / assigned 40 /
  unassigned 4 / stale 6 / oldest 189`.
- Independently re-parsed `C:\Users\admin\Downloads\All Open Tasks by Team.xlsx`
  (mtime 27 Aug 10:16) with the live script's own functions in a **no-push
  dry-run**: byte-identical to live `tickets.json` for every ticket (only the
  `updated`/`updated_display` timestamps differ, which is expected — they are
  `datetime.now()`).
- GitHub Pages build for `7b50878`: `status: built`, 09:20 UTC. `index.html`
  cache-busts its fetch (`fetch('data/tickets.json?t='+Date.now())`), so the
  live site genuinely serves the 10:20 data now.
- Rendered a real Playwright screenshot of the live dashboard (1440-wide, full
  page) — header reads "Source: All Open Tasks by Team.xlsx — Thursday 27
  August 2026 at 10:20", queue summary 44/40/4, analyst blocks James 3 / Asta
  26 / Michael 9 / Simon 1 / Kevin 1 / Unassigned 4. Sent to coordinator.
- Delta vs the 08:47 morning auto import (`- auto.xls`): +8 new task numbers,
  −4 resolved, unassigned 8→4 (Asta 17→26). Plausible 2-hour movement; no
  analyst appeared/vanished.

### Root cause of "not working" — no reproducible hard failure
- Ran the **exact** Desktop `.bat` mechanism (curl `main` → `pip install
  openpyxl requests` → `python %TEMP%\import_osm_report.py`) with the push
  disabled: curl OK (12797 bytes), pip OK, script found the new `.xlsx`,
  parsed 44 tickets cleanly. It works on this machine right now.
- The Desktop file `D:\OneDrive - lelitte.com\Desktop\Update HRIS Dashboard.bat`
  is the **2 July 2026 version** (656 bytes, unpinned `main` pull, `pip install
  openpyxl requests` with **no `xlrd`**). It is missing three later repo fixes
  to `Update HRIS Dashboard.bat`: `d3e4c24` (dep auto-install hardening),
  `a23f1dd` (add `xlrd`), `7b3014b` (pin download to SHA `e503509`, Kevin-approved
  20 Aug). It works today only incidentally: `xlrd 2.0.2` is already installed
  globally, and `main`'s `import_osm_report.py` still equals the pinned SHA
  (`git diff e503509 origin/main -- import_osm_report.py` = empty).
- Latent defects in that stale manual path: (a) `import_osm_report.py` does an
  unconditional top-level `import xlrd` and hard-exits if it is absent — even
  for `.xlsx`-only runs — and the Desktop `.bat` would not reinstall it;
  (b) `find_report()` silently picks the **newest** file matching
  `All Open Tasks by Team*.xlsx|xls` in Downloads, so a run started before the
  manual export finishes downloading (or while Excel still locks it) imports
  the 08:01 morning `- auto.xls` instead, or throws a lock error; (c) `pip
  install` on every run — slow, network-dependent, noisy.
- Most probable trigger for Kevin's experience (not confirmable after the fact —
  a pre-push `.bat` failure leaves no git trace): first run before the fresh
  `.xlsx` had finished downloading / while Excel held it open → wrong file or a
  lock error → retry at 10:19 succeeded. Precedent: the 25 Aug "dashboard hasn't
  updated" report was also not-a-bug (timing + Pages lag).

### Morning auto path (Scheduled Task "HRIS Dashboard Morning Refresh")
- **Not affected by the stale Desktop `.bat`.** `Run_HRIS_Auto_Refresh.bat`
  (Desktop copy dated 22 Aug, matches repo) downloads `Update HRIS Dashboard.bat`
  **fresh from GitHub `main`** every run, so it uses the good pinned repo copy,
  not Kevin's Desktop file. Today path B succeeded on attempt 1: commits
  `90a8039` + `3da82d4`, `data/last_automated_run.json` = `status: success`
  08:47. Next scheduled run 28 Aug 08:45.
- **Separate, pre-existing, still broken (NOT this issue):** the GitHub Actions
  self-hosted-runner SAASIT scrape (`generate_dashboard.py`, `update-dashboard.yml`,
  hourly cron) still fails `SAASIT session expired (ISM_4001)` — last diagnostic
  run #159, 26 Aug 16:16 UTC. Needs Kevin's interactive Oxford SSO+MFA re-login
  via `Refresh Session.bat`; no agent can do this. This also means the
  dashboard's on-page "Refresh" button (Cloudflare Worker → `workflow_dispatch`
  on that workflow) currently triggers a run that fails.

### "Better way" proposal for the manual mid-day top-up — DECISION NEEDED FROM KEVIN
All options keep GitHub-only writes, timestamped output, and leave the morning
Scheduled Task untouched. All are git-revertable.

- **Option 1 — re-sync + harden (do this regardless, ~1–2 h).** Copy the repo's
  current `Update HRIS Dashboard.bat` to the Desktop. In `import_osm_report.py`:
  make `import xlrd` lazy (only when a `.xls` is actually selected); detect
  `.crdownload`/`PermissionError`/`BadZipFile` and print a clear "export still
  downloading or open in Excel — close it and re-run" message with a short
  retry; print the chosen filename + mtime prominently and refuse (with a
  prompt) if the newest match is > ~20 min old; after push, print the Pages URL
  and the ~60–90 s rebuild lag.
- **Option 2 — drag-and-drop target (~2 h).** Desktop shortcut that passes the
  dropped file path as an explicit arg to the script (new optional `argv[1]`),
  bypassing `find_report()`'s newest-guessing entirely. Fast, removes the
  wrong-file failure mode; still local Python + PAT.
- **Option 3 — folder-watch, no click (RECOMMENDED, ~3–4 h).** A small always-on
  watcher (poll `C:\Users\admin\Downloads` every ~60 s, launched from the
  Startup folder — the same pattern already running live for
  `ai-usage-dashboard`). On a settle-detected new/changed `All Open Tasks by
  Team*.xlsx|xls` it runs the **same** `import_osm_report.py` and pushes.
  Kevin's whole manual step becomes "just save the export to Downloads".
  Dedupe by last-imported filename+mtime+SHA so it does not double-fire on the
  morning `- auto.xls`. Reuses a proven pattern, one shared import script,
  coexists cleanly with the Scheduled Task.
- **Option 4 — browser upload via Cloudflare Worker (~1–2 d).** File input on the
  dashboard → POST xlsx to a Worker that parses it (SheetJS) and commits
  `tickets.json` with the server-held PAT. No local Python/PAT/machine needed;
  works from any device. Higher complexity/risk; parsing logic forks from the
  Python version unless deliberately shared.

**Recommendation: Option 3, with Option 1's hardening folded in first** (cheap,
de-risks every path). Option 2 is the fast fallback if Kevin wants to keep an
explicit manual trigger. Option 4 only if he needs to update from a phone /
non-work machine.

### Exact next action / owner
1. **Kevin (via coordinator):** pick a "better way" option (or approve Option 1
   alone). Nothing to build until then.
2. **No live-site push outstanding** — data is current. If a redundant re-import
   is ever run it only bumps the "updated" timestamp text with zero data change;
   not worth doing.
3. If Kevin wants Option 1/2/3: Drew implements on a branch, screenshots any
   dashboard-visible change for approval, merges to `main` immediately per this
   repo's branch protocol.

### Rollback / deploy status
- No deploy made this session. Live `main` HEAD `7b50878` is Kevin's own good
  data. Nothing to roll back.

---

## One-line resume — 25 Aug 2026, afternoon


**Kevin asked for friendly, non-technical status messaging on the dashboard
(mirroring a separate ask Markey is handling for Linda's chat error on
`hr-fa-knowledge-base` — not touched here). Built, screenshot-verified via a
real local Playwright render (not just described), Kevin reviewed and
approved the 4 screenshots, then pushed to `main` and confirmed live.**

**What changed in `index.html` (commit `ad3638a6`):**
- The automated-refresh badge (`loadAutomationStatus()`) previously showed
  raw internals in the visible text on any failure — e.g.
  `Failed (fetch_failed) — attempt 1 of 3` — even on attempt 1, which is
  normal in-progress retry behaviour, not a real problem, yet displayed in
  red as if it were one. Replaced with three clearer states: green
  **"Last automated refresh: `<time>` — up to date"** on success; a new
  neutral blue **"Data is refreshing — check back shortly."** while a retry
  is still pending (attempt < max — added a new `.main-header-auto.pending`
  CSS class, same established colour system, Oxford-blue toned); red
  **"Automatic refresh didn't complete this morning — data may be a little
  behind. Run Update HRIS Dashboard.bat if you need it sooner."** only once
  retries are genuinely exhausted (attempt 3 of 3). Internal step/detail
  codes (e.g. `fetch_failed`) moved to the badge's hover `title` tooltip
  only — still there for troubleshooting, no longer in the visible text.
- `loadData()`'s "no ticket data" error card previously printed
  `Error: HTTP 404` (or whatever the raw fetch error was) directly under
  the heading. Now shows friendly copy only ("Data isn't available right
  now... this usually clears up on its own within a few minutes..."), with
  the raw error message moved to that card's `title` tooltip instead of
  its visible text.
- The "no automated run recorded yet" (grey/unknown) state and the
  "no ticket rows returned" empty state were already friendly — left as is.

**Verified for real, not just reasoned about:** built a local Playwright
test harness (`python -m http.server` + `playwright.sync_api`, real Chromium)
serving the actual edited `index.html` with synthetic `data/tickets.json`
and four variants of `data/last_automated_run.json` (success, mid-retry,
exhausted-retries, missing-file), captured a real screenshot of each
rendered state plus one of the "ticket data missing" error card — five PNGs
total. Kevin reviewed all five and gave explicit approval before anything
was pushed (standing rule: screenshots required for visual-change
approval, text descriptions aren't sufficient). After pushing, confirmed
the new strings genuinely live on `main` via
`raw.githubusercontent.com` (cache-busted) before considering this done —
not assumed from the push response alone.

**Deliberately not touched:** Linda AI's own
`"Connection error: Failed to fetch"` message in this same `index.html`
(right-hand panel) — same pattern as the `hr-fa-knowledge-base` Linda chat
error Kevin separately asked Markey to fix. Flagged for Kevin's awareness,
not fixed here, to avoid overlapping with Markey's parallel work. If Kevin
wants this pulled into Drew's scope too, it's a small, contained follow-up
(one `.catch()` handler, `onDone('Connection error: '+err.message)`).

**Untouched, as required:** `data/last_automated_run.json`'s actual schema
and the automation pipeline that writes it (`push_automation_status.py`,
`check_last_run_status.py`, `Run_HRIS_Auto_Refresh.bat`) — this was a
front-end display-only change, no automation logic touched.

---

## Superseded — 25 Aug 2026, morning

**Kevin reported at 9am that the dashboard hadn't updated. Investigated
live, root cause confirmed NOT a bug — the 22 Aug retry mechanism worked
exactly as designed on its first real "primary attempt genuinely fails"
morning. No code/config change made; this is a report-and-verify entry.**

**Timeline, all confirmed via live `Get-ScheduledTaskInfo`, the real
`osm_auto_refresh_last_run.log`, and the GitHub API (not inferred):**
- **08:45 (attempt 1, primary):** fired on time, `LastTaskResult: 1`.
  `fetch_osm_report.py` found `Inbox/Reports/OSM` at 55 items; most recent
  matching email was still 24 Aug (today's, 25 Aug, hadn't landed yet —
  same mail-rule-filing-delay pattern root-caused on 22 Aug). Failed
  loudly and correctly, exactly as designed; pushed
  `data/last_automated_run.json` with `status: failure, attempt: 1`.
  Kevin's 9am report was accurate — dashboard genuinely still showed
  24 Aug data at that point.
- **09:15 (attempt 2, retry):** fired on schedule. Today's OSM email had
  arrived by then. `fetch_osm_report.py` succeeded, `import_osm_report.py`
  parsed 48 tickets, pushed `data/tickets.json`
  (SHA `c184fedc5ede6c3b6ac8f05cc642f7983416b7bb`, `updated:
  2026-08-25T09:15:12`) and `data/last_automated_run.json`
  (`status: success, step: dashboard_update_ok, attempt: 2`).
  `LastTaskResult: 0`.
- Watched the live Scheduled Task transition from `Running` to `Ready`
  in real time (polled directly, not assumed), then confirmed both new
  JSON files live on `main` via the GitHub Contents API, then polled
  `https://begb0037admin.github.io/hris-dashboard/` (cache-busted) until
  GitHub Pages' normal build lag cleared and the live site genuinely
  served the 09:15 success data (~90s after the `main` push — expected
  lag, not a bug, consistent with the 22 Aug session's note about this).

**Current live status, confirmed:** dashboard is fully current — 48
tickets (26→48 vs 24 Aug's stale 31; all now assigned, 0 unassigned),
`data/last_automated_run.json` on both `main` and GitHub Pages reads
`status: success`, `attempt: 2 of 3`. The 09:45 (attempt 3) trigger will
still fire per schedule but `check_last_run_status.py`'s guard will see
today's success and print `SKIP` — no redundant run, no risk of
overwriting the correct 09:15 timestamp.

**Separate, pre-existing, NOT the cause of this morning's report — flagged
for awareness only:** the unrelated GitHub Actions self-hosted-runner
SAASIT scrape path (`generate_dashboard.py`, hourly cron) is still
failing with the same `SAASIT session expired (ISM_4001)` error first
found 11 Aug 2026 — confirmed still broken as of run #141, 2026-08-25
08:42 UTC (`last_run_status.txt`). This is a completely different data
path from the OSM-email pipeline above and does not touch
`data/tickets.json` when it fails (fails loudly, leaves prior data in
place, by design). Needs Kevin's interactive Oxford SSO+MFA login
(`Refresh Session.bat` / `python login.py`) — no agent can do this on
his behalf. Unchanged, no action taken, not new information — just
re-confirmed live rather than assumed stale from the 20 Aug note.

**No code or configuration changed this session.** Nothing needed
fixing — the system self-healed exactly as the 22 Aug retry mechanism
was built to do. This is the first real-world morning that mechanism
was tested against a genuine "primary attempt fails, retry recovers"
case, and it worked end-to-end, unattended, with no manual intervention.

---

## Superseded — 22 Aug 2026, afternoon

**Kevin decided both items left open by the late-morning investigation
below: moved the trigger later and built a retry mechanism. Both live and
verified for real this session (see full detail in `HANDOVER.md`'s
"Session 2026-08-22 (afternoon)" entry).**

**Trigger schedule:** `HRIS Dashboard Morning Refresh` now fires three
times — **08:45 (primary), 09:15, 09:45 (retries)** — on the same task,
same Action/Principal/Settings (confirmed unchanged). `NextRunTime`
confirmed live: `23/08/2026 08:45:00`.

**Retry guard, new file `check_last_run_status.py`:** every trigger
checks the live `data/last_automated_run.json` FIRST. If today already
succeeded, it prints `SKIP` and the `.bat` exits 0 immediately — no
Outlook connection, no GitHub push, and the real first-success timestamp
is never overwritten by a redundant later run (this is what keeps "Last
automated refresh" accurate rather than blurred). Otherwise it prints
`RUN:<attempt>:<max>` (a stateless wall-clock-time bucket: <09:00→1,
<09:30→2, else→3). `push_automation_status.py` gained `--attempt`/
`--max-attempts`; `index.html`'s badge now shows "attempt N of 3" on
failure, and "no more automated retries today" once N=3, so three real
failures read as one clear final state rather than a confusing blur.

**A real bug caught before any live test:** the first draft of the
retry-guard used a multi-line `if (...) else (...)` block that set and
then read back a variable within the SAME block — precisely the
cmd.exe trap this file already has two prior root-caused incidents
about (parens expand `%VAR%` once, at parse time). Rewritten with
sequential `goto`s per this file's own existing lesson, before it ever
ran live.

**Verified live, for real, this session:** `check_last_run_status.py`
against genuine live production data (today's real success record) →
`SKIP`, correct. 5 mocked-`requests` unit tests covering every attempt
boundary and edge case, all pass. The RUN branch proven against a
**genuine live network fetch** (not a mock) by pushing synthetic
failure data to a throwaway branch, pointing a scratch copy of the real
script at it, confirming `RUN:3:3`, then deleting the branch. The new
`.bat` `goto` logic tested in isolation against 5 realistic inputs. The
real, just-deployed `.bat` run directly (hit the SKIP path, confirmed
`data/last_automated_run.json` on `main` byte-identical before/after).
The REAL registered Scheduled Task triggered end-to-end via
`Start-ScheduledTask` — genuine Task Scheduler → `wscript.exe` (hidden)
→ VBS → Desktop `.bat` chain — `LastTaskResult: 0`, log confirms correct
skip behaviour under real automated execution. Deployed to Desktop with
CRLF (confirmed via `file`). All 4 files pushed to `main` via the GitHub
Contents API (this repo's own `AGENT_MODEL.md` convention, not
`git push`), each independently confirmed live via
`raw.githubusercontent.com`. Commits: `d6dfd506`, `0972d487`,
`6c8fd776`, `a94dc1d4`.

**`Update HRIS Dashboard.bat` / `import_osm_report.py` / the manual
path: zero changes**, as required.

**First real test of the new schedule under a genuine "hasn't succeeded
yet" morning: tomorrow, 23 Aug 2026, 08:45.** Everything provably works
live as of this session; what only tomorrow's real clock can prove is a
full cycle where the primary run genuinely fails (report not yet filed)
and a retry picks it up from a cold start.

---

## Superseded — 22 Aug 2026, late morning

**Investigated Kevin's report of seeing today's OSM email at 8:06am despite
the 08:30 run saying "no email today" — confirmed a genuine mail-rule
filing delay, not a script bug, not a folder-path mismatch, no fix needed.
Re-ran the scheduled task live afterward: succeeded end-to-end, dashboard
now genuinely current.**

**Folder path confirmed correct** — live Outlook COM probe opened the
exact path `fetch_osm_report.py` hard-codes (`Inbox > Reports > OSM`
under the default store), confirmed it's the only folder anywhere in any
of the 5 connected stores with this name/nesting, and matches what Kevin
sees in Outlook's own UI. No mismatch.

**Root cause, confirmed with direct before/after evidence, not inferred:**
the 08:30 log (below) recorded the OSM folder at **52 items**; a live
probe at 11:58 today found **53 items**, with today's email
(`ReceivedTime` 2026-08-22 08:06:47) now present. This folder receives
exactly one message a day via a mail rule — the 52→53 delta is direct
proof the email was filed into this specific watched folder *after* the
08:30 script ran, even though its `ReceivedTime` (stamped at actual mail
delivery, not at rule-filing time) reads as 8:06, before 08:30. This is
a mail-rule/filing delay between delivery and the message landing in the
folder the script polls — not a date/timezone bug in
`fetch_osm_report.py`'s comparison logic, which read the real live folder
state correctly at the moment it ran and failed loudly with an accurate
diagnosis, exactly as designed.

**One low-priority observation, flagged not fixed:** `msg.ReceivedTime`
via `win32com`/`pywintypes` carries a `(UTC+00:00) ...London` tzinfo
label even now, in BST (UTC+1) — the raw digits (08:06:47) match Kevin's
own stated local observation exactly, suggesting the value is local time
mislabelled as UTC rather than a true UTC value. Doesn't affect this
investigation (date component unaffected, arrivals are always mid-morning,
nowhere near a midnight boundary) — worth remembering only if a future
bug involves genuine UTC arithmetic on `ReceivedTime` or a near-midnight
arrival.

**Verified live, for real:** triggered
`Start-ScheduledTask -TaskName 'HRIS Dashboard Morning Refresh'` at
12:00:08; polled `Get-ScheduledTask` directly until it left `Running`
(12:00:29, `LastTaskResult: 0`). Fresh `osm_auto_refresh_last_run.log`
from this run shows: folder now 53 items, today's email found and its
attachment saved, `Update HRIS Dashboard.bat` → `import_osm_report.py`
ran unmodified, 31 tickets parsed, `data/tickets.json` pushed
(SHA `29b0b32e...`), `push_automation_status.py` pushed
`status: success`. Confirmed via `raw.githubusercontent.com` (bypassing
GitHub Pages' build-cache lag, which was still showing stale data at
check time — expected, not a bug) that `main` genuinely has
`data/tickets.json` (`updated: 2026-08-22T12:00:19`) and
`data/last_automated_run.json` (`status: success`,
`timestamp: 2026-08-22T12:00:21`). Commits `1b5c0ba2` and `3fb07c34`.

**Still open, unchanged, Kevin's call:** whether to move the 08:30
trigger later, add a retry-later mechanism, or accept the manual fallback
for occasional late-filing mornings. The automation itself is confirmed
healthy — this is a scheduling-tolerance question, not a defect.

Full detail: `HANDOVER.md`'s "Session 2026-08-22 (late morning)" entry.

---

## Superseded — 22 Aug 2026, morning entry (root cause now confirmed above — mail-rule filing delay, not a defect)

**First real test of the 21 Aug fix, under genuine Task Scheduler/
`wscript.exe` hidden execution — verified live, not assumed.**

`Get-ScheduledTaskInfo` confirmed `LastRunTime: 22/08/2026 08:30:01`,
`LastTaskResult: 1`, `NextRunTime: 23/08/2026 08:30:00` — the task did
fire today, on time, under the real automated context (not a stale
value from an earlier day).

**Result: failed again, but this is NOT a recurrence of the 21 Aug bug —
it's a genuinely different, well-understood failure mode, and the fix
is proven to be working correctly.** `osm_auto_refresh_last_run.log`
(`C:\Users\admin\Documents\Claude\Projects\HRIS-Dashboard\`) now shows
full, real diagnostic detail from the moment the run started, instead of
dying silently at ~10 seconds with nothing logged (the 21 Aug symptom):

```
[2026-08-22 08:30:02] Run HRIS Auto-Refresh started
[2026-08-22 08:30:02] Downloading fetch_osm_report.py and push_automation_status.py from GitHub...
[2026-08-22 08:30:02] Step 1/2 - fetching today's OSM report attachment from Outlook...
[2026-08-22 08:30:03] fetch_osm_report.py run started
[2026-08-22 08:30:04] Found folder: Inbox/Reports/OSM (52 items)
[2026-08-22 08:30:04]   (most recent matching email is from 2026-08-21, not today 2026-08-22)
[2026-08-22 08:30:04] Most recent email from reports-prd-ldz@saasiteu.com: received 2026-08-21 08:01:22.378000+00:00, subject 'Report: All Open Tasks by Team'
ERROR: No email from reports-prd-ldz@saasiteu.com with subject 'Report: All Open Tasks by Team' received TODAY (2026-08-22). The dashboard was NOT refreshed automatically this morning. If the report is just running late, re-run this script once it arrives, or run the manual Update HRIS Dashboard.bat path.
[2026-08-22 08:30:02] Step 1/2 FAILED, exit code 1 - Step 2 SKIPPED. Dashboard NOT refreshed this run.
Pushed data/last_automated_run.json: status=failure step=fetch_failed
```

Outlook COM connected fine, the folder was found fine, `fetch_osm_report.py`
ran and exited cleanly with a specific, correct diagnosis — today's OSM
report email simply had not arrived in `Inbox/Reports/OSM` by 08:30 (the
most recent matching email at that moment was still yesterday's, 21 Aug
08:01). This is exactly the "silent break made visible" behaviour the
whole automation was built for, working as designed — the opposite of
21 Aug's silent parse-crash.

Confirmed via GitHub API (not assumed from the log alone):
`data/last_automated_run.json` shows
`{"timestamp": "2026-08-22T08:30:05", "status": "failure", "step":
"fetch_failed", "detail": "fetch_osm_report.py exited 1 - see
osm_auto_refresh_last_run.log", "trigger": "automated"}`. Most recent
`data/tickets.json` commit is still `708a59f` ("OSM import — 2026-08-21
18:29"), i.e. last night's manual catch-up run — the dashboard has NOT
been refreshed today as of this check and is currently showing 21 Aug
data.

**Per this task's own instructions, no manual re-run was performed** —
report-only, live state left exactly as found.

**Decision needed from Kevin (not yet actioned):** the OSM email has, on
at least this one observed morning, not landed before the 08:30 trigger
(its usual arrival window was previously observed as ~08:01–08:16, so
this may be a one-off late delivery rather than a systemic timing
mismatch — one data point isn't enough to tell). Whether to address this
via a later trigger time, a retry-later-in-the-morning mechanism, or just
leave the manual-fallback path (`Update HRIS Dashboard.bat`) as the
answer for late-report mornings is Kevin's call, not yet decided or
built.

---

## Superseded — 21 Aug 2026, afternoon entry (fix confirmed working correctly by the 22 Aug real-world test above)

The 21 Aug 08:30 scheduled run fired but died silently after Step 1 (no
error in the log, nothing in Windows event logs) — Kevin reported it same
day. Root-caused and fixed for real, two independent bugs stacked, both
confirmed via isolated bisection testing on this exact machine, not
guesswork:

1. **Unescaped literal parens inside a multi-line `if (...)` block.** The
   old failure-branch echo — `Step 1/2 FAILED (exit %FETCH_EXIT%) ...` —
   sits inside `if not "%FETCH_EXIT%"=="0" ( ... )`. cmd.exe's block
   parser ends a parenthesized block at the FIRST unescaped `)` it meets,
   not by balancing against an earlier `(` — so the `)` right after
   `%FETCH_EXIT%` was read as the block's real close. Everything after
   that point misparsed into cmd's classic `- was unexpected at this
   time.` error, written only to the (invisible, since this always runs
   hidden) console, never the log — exactly why this looked like total
   silence. Fired on **every** run regardless of Step 1's real outcome,
   since cmd must parse a whole if-block to know where to resume even
   when skipping it — this is why today's run died even though Step 1
   itself (Outlook fetch) genuinely succeeded. Fixed by rewording to avoid
   literal parens.
2. **This machine's cmd.exe cannot launch a second `.bat`/`.cmd` FILE from
   a cmd.exe process already executing a batch file** — confirmed via
   `call`, bare invocation, and a fresh `cmd /c "x.bat"` all failing
   identically with `'"x.bat"' is not recognized as an internal or
   external command`, while nested `cmd /c dir`/`echo` (trivial commands,
   not script files) and curl.exe/python.exe (real .exe targets) all
   worked fine from the exact same nested context. This is what would have
   blocked Step 2 (`Update HRIS Dashboard.bat`) even with bug 1 fixed.
   Fixed by routing that one invocation through PowerShell as an
   intermediary (`powershell -NoProfile -Command "Get-Content
   'empty_stdin.txt' | & '.\Update HRIS Dashboard.bat'; exit
   $LASTEXITCODE"`), confirmed empirically (including with a real `pause`
   inside the target, so it doesn't hang) before applying it live.

**Also added, per Kevin's explicit ask:** the whole run now executes as
`call :main > "%LOG_FILE%" 2>&1` — one redirection covering the entire
subroutine, both stdout and stderr — instead of scattered per-line log
appends, so any *future* cmd.exe parse error lands in the log
automatically instead of being silently invisible again. All three curl
downloads now use `-f` so an HTTP error (e.g. a 404 for a file not yet on
`main`) is caught as a real failure instead of being silently saved as
file content (this actually bit during testing: `push_automation_status.py`
hadn't been pushed yet, curl without `-f` saved "404: Not Found" as if it
were the script, which python then failed to parse with a confusing
`SyntaxError`).

**New:** `push_automation_status.py` (repo root, new file) pushes
`data/last_automated_run.json` to GitHub after every run — success or any
distinct failure mode — and `index.html` now shows a real **"Last
automated refresh: `<time>` — Success/Failed"** badge next to the existing
manual-update timestamp, sourced from the automation's own run record
rather than tickets.json's mtime (a failed run leaving yesterday's data in
place would otherwise look identical to a fresh successful one).

**Verified live, for real, this session:**
- Both bugs reproduced via isolated bisection (minimal repro scripts,
  safe dummy `.bat` files) before touching the real files, and the fixes
  confirmed to resolve each in isolation before being applied live.
- Ran the actual fixed `Run HRIS Auto-Refresh.bat` end-to-end for real,
  twice — first attempt (bug 2 not yet found) failed cleanly with the new
  clear error instead of a silent parser crash, confirming bug 1's fix
  and surfacing bug 2; second attempt (both fixes applied) completed
  successfully: fetched today's real OSM email attachment, ran the real
  **unmodified** `Update HRIS Dashboard.bat` → `import_osm_report.py`
  pipeline, parsed 31 tickets, and pushed `data/tickets.json`
  (SHA `0e0c53b94c61a52fdce8aeef2034f59dcf8faf8d`) — this is also today's
  real data catch-up, since the 08:30 scheduled run never got this far.
  `push_automation_status.py` pushed `data/last_automated_run.json`
  (`status: success`, `step: dashboard_update_ok`) in the same run.
- Confirmed both new JSON endpoints are live via GitHub Pages directly
  (`curl https://begb0037admin.github.io/hris-dashboard/data/tickets.json`
  and `.../data/last_automated_run.json`), and that the rebuilt
  `index.html` on Pages contains the new `main-header-auto` CSS/markup.
  **Not verified:** an actual browser screenshot of the rendered badge —
  no browser/screenshot tool was available this session; the JS
  (`loadAutomationStatus()`) mirrors the already-proven `loadData()` fetch
  pattern exactly and both data endpoints it depends on are confirmed live
  and correctly shaped, but the literal pixel rendering hasn't been
  visually confirmed.
- Fixed file deployed to `D:\OneDrive - lelitte.com\Desktop\Run HRIS
  Auto-Refresh.bat` (CRLF line endings, matching the machine's existing
  convention) — this is the file Task Scheduler's `HRIS Dashboard Morning
  Refresh` task actually points at (via the `.vbs` wrapper, unchanged).
  Also pushed to the repo (`Run_HRIS_Auto_Refresh.bat`, `main`) so a
  future re-copy to Desktop carries the fix forward. Commit `b0ad42c`
  (merged as `f4fe3d6` after a GitHub Actions diagnostics-bot commit raced
  it).
- `Update HRIS Dashboard.bat` and `import_osm_report.py` (the manual
  pipeline) were not touched at all — confirmed by diff, only their
  *invocation method* from the automated orchestrator changed.

**Next scheduled run:** tomorrow, 22 Aug 2026 08:30 — first real
end-to-end test of the fix under the actual hidden/Task-Scheduler
execution context (today's verification ran visibly via PowerShell, not
via `wscript.exe`/Task Scheduler itself, though the PowerShell-intermediary
fix for bug 2 was specifically chosen and tested because the earlier
`< NUL`-based approach was already proven live under the real hidden
context in the 20 Aug session). **Watch item:** if tomorrow's run shows a
failure toast, `osm_auto_refresh_last_run.log` will now contain the real
reason (including any future parse error) rather than stopping short —
check that file first.

Full detail: `HANDOVER.md`'s `Session 2026-08-21 (afternoon)` entry.

---

## Superseded — 21 Aug 2026, morning entry (root cause not yet found at this point)

The morning auto-refresh automation is now **LIVE**. Kevin gave explicit
go-ahead; Drew ran `Register-HRISAutoRefreshTask.ps1`, found and fixed a
real bug in it (the staged script's `-DisallowStartIfOnBatteries` flag
doesn't exist on this machine's ScheduledTasks module version, which
made the first attempt silently fail to register while still printing
its own "registered" message — caught by checking `Get-ScheduledTask`
directly rather than trusting the script's own output), then re-ran it
and verified the resulting Task Scheduler job's trigger/action/principal/
settings directly against live state. Task `HRIS Dashboard Morning
Refresh` is `Enabled: True`, `State: Ready`, daily trigger at 08:30
(every day, not weekdays-only), pointed at the correct `.vbs` wrapper.
**This entry assumed "nothing else is blocking" — WRONG, see the
superseding entry above: two real bugs in the script itself meant the
08:30 run the next morning died silently anyway.**

---

## What's new this session (20 Aug 2026, evening — Drew)

**Task:** Kevin gets a second OSM report daily by email now (separate
from the manual Excel-download workflow that already existed) —
`reports-prd-ldz@saasiteu.com`, subject "Report: All Open Tasks by
Team", lands in Outlook folder `Inbox > Reports > OSM` (confirmed live,
filed there by an existing mail rule) around 08:01–08:16 every single
day. He wants it picked up automatically at ~8:30am, dropped where the
existing pipeline already looks, and the existing update pipeline run —
so the dashboard is fresh first thing each morning with no manual step.
**The existing manual midday path (download by hand + run
`Update HRIS Dashboard.bat`) is completely unchanged and continues to
work exactly as before** — this is additive automation only.

**Built (both pushed to `main`):**

1. **`fetch_osm_report.py`** (new file, repo root) — Outlook COM script.
   Connects to Outlook (reusing the exact retry-on-transient-COM-error
   pattern already proven live in `work-inbox/fetch_inbox.py`), opens
   `Inbox > Reports > OSM` on the default store, finds **today's** email
   matching the sender+subject above, and saves its attachment into
   Downloads as `All Open Tasks by Team - auto.xls` (or `.xlsx`,
   extension taken from whatever the real attachment is) — a name
   distinct from anything Kevin downloads by hand, but still matching
   `import_osm_report.py`'s existing `FILE_PATTERNS` glob, so the
   existing "most-recently-modified file wins" selection logic in that
   script needs **zero changes** to pick it up. Does not parse the
   spreadsheet, does not touch `data/tickets.json`, does not push
   anything to GitHub — that all stays entirely inside the existing,
   unmodified `import_osm_report.py`.

   **Fails loudly and specifically, on purpose,** rather than silently
   doing nothing or reusing a stale file, if: Outlook COM is
   unreachable; the `Inbox/Reports/OSM` folder doesn't exist (renamed/
   moved); no email from that sender was received **today** (as opposed
   to some earlier day); the sender matched but the subject didn't;
   or the email has no attachment / its attachment doesn't look like an
   Excel export. This is the detection mechanism for "OSM silently
   changed the report format" that Kevin explicitly asked for — a
   silently-stale dashboard would be worse than the old manual step.
   Supports `--dry-run DIR` for verification without touching the real
   Downloads folder (used for all testing this session — see below).

2. **`Run_HRIS_Auto_Refresh.bat`** (new file, repo root) — orchestrator.
   Pulls the latest `fetch_osm_report.py` from GitHub `main` (cache-
   busted) and runs it. **Only if that succeeds** (exit 0), it then
   pulls the latest `Update HRIS Dashboard.bat` from GitHub `main`
   (cache-busted) and runs **that exact, unmodified file** — same file
   Kevin double-clicks by hand, same pinned-`import_osm_report.py`-
   commit mechanism inside it, completely untouched — via
   `call "Update HRIS Dashboard.bat" < NUL`, which redirects that
   file's own `pause` prompt to immediate EOF so it completes
   unattended instead of hanging, without editing that file. If step 1
   fails, step 2 is skipped entirely — a missing/bad report never gets
   silently "updated" over. Logs every run, timestamped, to
   `C:\Users\admin\Documents\Claude\Projects\HRIS-Dashboard\osm_auto_refresh_last_run.log`.

3. **Desktop launcher files, staged but INERT** (not committed to the
   repo — these are local-execution artifacts, same category as the
   existing `Run Inbox Briefing Hidden.vbs` / `Show-TaskNotification.ps1`
   already on the Desktop for Work Inbox Briefing):
   - `D:\OneDrive - lelitte.com\Desktop\Run HRIS Auto-Refresh.bat` — a
     copy of the repo file above, placed where Task Scheduler needs a
     stable local path.
   - `D:\OneDrive - lelitte.com\Desktop\Run HRIS Auto-Refresh Hidden.vbs`
     — hidden-window wrapper mirroring `Run Inbox Briefing Hidden.vbs`
     exactly: runs the `.bat` with no visible console, captures and
     propagates its real exit code, then fires a desktop toast via the
     **existing** `Show-TaskNotification.ps1` (BurntToast) already on
     this Desktop — reused as-is, not a new notification mechanism.
   - `D:\OneDrive - lelitte.com\Desktop\Register-HRISAutoRefreshTask.ps1`
     — the Task Scheduler registration script itself. **NOT RUN.**
     Mirrors the exact live configuration of the existing "Work Inbox
     Briefing" task (`Get-ScheduledTask` checked directly, not assumed):
     LogonType Interactive, RunLevel Limited, UserId admin,
     DisallowStartIfOnBatteries/StartWhenAvailable both true. Trigger:
     **daily at 08:30, every day of the week** — confirmed against live
     mailbox evidence, not assumed, that the OSM report itself is a true
     7-day feed (genuine emails found dated Saturday 15 Aug and Sunday
     16 Aug 2026 in `Inbox/Reports/OSM`, not just weekdays).

**Verified live, for real, this session (not just read-through):**
- Live-probed the actual mailbox (Outlook COM, read-only) to find the
  OSM folder — it is not literally named "OSM" at the top level as
  Kevin's own description implied; it's a subfolder:
  `Inbox > Reports > OSM` under the default store
  (`kevin.lelitte@admin.ox.ac.uk`). Confirmed by listing all folders
  across all 5 connected Outlook stores.
- Confirmed live: sender `reports-prd-ldz@saasiteu.com`, subject
  `"Report: All Open Tasks by Team"`, one `.xls` attachment named
  `All Open Tasks by Team.xls`, arriving daily ~08:01–08:16, exactly as
  Kevin described — checked the 10 most recent real emails in that
  folder, not just one.
- Ran `fetch_osm_report.py --dry-run` directly against the real, live
  mailbox: found today's (20 Aug 2026) real email, saved its real
  attachment. Opened the saved file with `xlrd` and confirmed it's a
  genuine, parseable OLE `.xls` whose header row
  (`Owner Team, Parent Type, Parent Number, Service, Category,
  Task Number, Summary, Status...`) matches exactly what
  `import_osm_report.py`'s `parse_report()` already expects — so the
  existing, unmodified downstream parser will read a real auto-fetched
  attachment correctly.
- Ran the **actual, just-pushed** `Run_HRIS_Auto_Refresh.bat` logic (a
  test copy pointed at a scratch directory instead of the real
  Documents/Downloads paths) end-to-end: it really downloaded
  `fetch_osm_report.py` fresh from `raw.githubusercontent.com/.../main`
  and ran it successfully against the live mailbox. Exit code 0.
- Caught and fixed a real bug before it could bite in production: a
  literal space in the `Update HRIS Dashboard.bat` raw-GitHub URL fails
  to resolve via `curl` — confirmed by testing both the literal-space
  and `%20`-encoded forms directly. Fixed with `%%20` (batch-escaped) in
  the `.bat`; confirmed post-fix that the corrected URL genuinely
  downloads the real, current `Update HRIS Dashboard.bat` (47 lines,
  `SCRIPT_SHA` present) via `curl`.
- Confirmed live, with a minimal throwaway test `.bat`, that
  `call "some.bat" < NUL` reliably dismisses that script's own `pause`
  immediately (no hang) AND correctly propagates its real exit code to
  the caller — this is exactly what Step 2 of
  `Run_HRIS_Auto_Refresh.bat` depends on to run
  `Update HRIS Dashboard.bat` unattended without editing it.

**NOT verified — deliberately, because it requires either Kevin's
approval or waiting for a live daily run:**
- Step 2 (the real `Update HRIS Dashboard.bat` pipeline, which pushes
  `data/tickets.json` to GitHub) was never actually executed this
  session — only downloaded and inspected. Running it would be a real,
  consequential production write, which is exactly the kind of action
  this task's own instructions said not to do without approval.
- The Task Scheduler task itself was never registered — `Register-
  HRISAutoRefreshTask.ps1` is staged on the Desktop, ready, but not run.
- The hidden-window + toast-notification wrapper
  (`Run HRIS Auto-Refresh Hidden.vbs`) was written to exactly mirror the
  already-proven-live Work Inbox Briefing version, but has not itself
  been executed this session (would require actually running the whole
  pipeline including the live GitHub push, which is the approval-gated
  step above).

**Kevin's own next action:** review
`D:\OneDrive - lelitte.com\Desktop\Register-HRISAutoRefreshTask.ps1`
(and the two new repo files it depends on) and either run it himself or
tell Drew to run it. That single step is all that's left to make the
morning auto-refresh live. Everything upstream of that step has been
proven against the real mailbox and the real, current GitHub content.

---

*(Prior session content — first `.xls` support + supply-chain pinning
fix, same day, earlier — preserved below for continuity.)*

## Earlier this same day (20 Aug 2026, first session)

The `.bat` manual-import pipeline now accepts `.xls` as well as `.xlsx` —
fixed and verified against a real OLE-binary `.xls` file — but the fix has
**not yet been run against Kevin's actual live OSM attachment**. Next
action: Kevin drops today's `All Open Tasks by Team.xls` (from
`reports-prd-ldz@saasiteu.com`, 20 Aug 2026) into `C:\Users\admin\Downloads`
and double-clicks `Update HRIS Dashboard.bat`; confirm it completes and
`data/tickets.json` / the live dashboard reflect it.

---

## What's fixed this session (20 Aug 2026)

**Problem:** `import_osm_report.py`'s `FILE_PATTERN` only matched
`*.xlsx`, and its only reader was `openpyxl.load_workbook()`, which
cannot open `.xls` at all. OSM Administration's "All Open Tasks by
Team" report sometimes arrives as a genuine old-format `.xls` (OLE
binary, confirmed — not a renamed `.xlsx`). Dropping that file into
Downloads and running the `.bat` would fail outright (no matching
file found, since the pattern is `.xlsx`-only).

**Fix chosen (option a from the brief — widen the script, not the
intake):** kept the single `.bat`-driven pipeline, no new conversion
step, no LibreOffice/soffice dependency, no request to OSM Administration
to change their export format.

- `FILE_PATTERN` → `FILE_PATTERNS` (list): now globs both
  `All Open Tasks by Team*.xlsx` and `All Open Tasks by Team*.xls`,
  picks the single most-recently-modified match across both via
  `os.path.getmtime` — unchanged selection semantics, just widened.
- Added `_read_rows_xls()` using `xlrd` (xlrd 2.x dropped `.xlsx`
  support but still reads legacy `.xls` fine — it's the complementary
  library to openpyxl, not a version choice). Normalises xlrd's typed
  cells (`XL_CELL_DATE`, `XL_CELL_NUMBER`, etc.) into the same
  `None`/`str`/`int`/`float`/`datetime` shape that
  `openpyxl.iter_rows(values_only=True)` already produced, via
  `xlrd.xldate_as_datetime(value, wb.datemode)` for dates.
- `read_rows(path)` dispatches on file extension to `_read_rows_xlsx`
  (existing logic, unchanged) or the new `_read_rows_xls`. Everything
  downstream — header-row detection by `"Task Number"`, column-name
  mapping, per-cell parsing, grouping — is unchanged and format-agnostic,
  so the `.xlsx` path carries zero behavioural risk from this change.
- `Update HRIS Dashboard.bat`: `pip install openpyxl requests` →
  `pip install openpyxl xlrd requests`.

**Commits (main):**
- `import_osm_report.py`: commit `f418802dac35445cbeaa0585d1276d3d963af015`
- `Update HRIS Dashboard.bat` (.xls + xlrd install): commit `a23f1dd0cef25bb9739397636687bc2096fea0e9`
- `Update HRIS Dashboard.bat` (SHA-pin the download, supply-chain fix): commit `7b3014beb9d696c4d2d40da2d1069bc87996ed4b`

---

## What's actually verified vs. not (earlier session)

**Verified directly, this session, with real evidence:**
- Confirmed via `file` on a freshly-generated sample that a genuine
  `.xls` is `CDFV2 Microsoft Excel` (OLE binary) — matches the brief's
  claim, not just taken on faith.
- Installed `xlrd` 2.0.2 and `xlwt`, built a synthetic `.xls` reproducing
  the real report's shape (junk title row, header row with the exact
  column names the script looks for, a date cell, an unassigned-analyst
  row) and ran the new `import_osm_report_new.py`'s `parse_report()`
  against it directly (not through the `.bat`, no GitHub push side
  effects). Output: 3 tickets parsed correctly, dates converted right
  (`xldate_as_datetime` path exercised), analyst grouping and the
  unassigned bucket both correct.
- Built a matching synthetic `.xlsx` and re-ran `find_report()` with
  both files present in the same directory, flipping which one had the
  newer mtime each way — confirmed it always picks whichever is
  genuinely most recently modified, regardless of extension, and that
  the pre-existing `.xlsx` path is unchanged (1 ticket parsed correctly
  from the `.xlsx` test file).

**NOT verified (earlier session) — this session ran on a non-Windows-GUI
environment with no access to Kevin's actual Downloads folder, real
GITHUB_PAT, or the real email attachment:**
- The `.bat` itself was not double-clicked end-to-end.
- Kevin's real `All Open Tasks by Team.xls` attachment (from 20 Aug
  2026) was never opened or parsed by that session — only a synthetic
  file with the same expected shape.
  **UPDATE (same day, evening session):** this has now effectively been
  covered — `fetch_osm_report.py --dry-run` opened and parsed the real
  live 20 Aug attachment successfully, header row confirmed matching.

---

## Supply-chain gap — FIXED same day, 20 Aug 2026 (Kevin approved)

`Update HRIS Dashboard.bat` no longer pulls `import_osm_report.py` from
the floating `main` branch — it now pulls from a commit-pinned raw URL:

`https://raw.githubusercontent.com/begb0037admin/hris-dashboard/<SCRIPT_SHA>/import_osm_report.py`

`SCRIPT_SHA` is set near the top of the `.bat` (currently
`e503509d2785cc30d4102365aa72e353983f864d`, `main`'s tip at the time of
the `.xls` fix, verified live before use). **Any future change to
`import_osm_report.py` must also bump `SCRIPT_SHA` in the `.bat` in the
same change**, or explicitly note the pin is deliberately being left
as-is. Commit: `7b3014beb9d696c4d2d40da2d1069bc87996ed4b`.

**This does NOT affect the new evening-session automation above** —
`Run_HRIS_Auto_Refresh.bat`'s step 2 downloads and runs the current
`Update HRIS Dashboard.bat` fresh from `main` every morning, so it
always uses whatever `SCRIPT_SHA` that file currently has pinned,
automatically, with no separate pin to maintain in the new automation
itself.

---

## Flagged, not fixed (informational / for Kevin's awareness)

1. **The Desktop copy of `Update HRIS Dashboard.bat` was found stale**
   (evening session discovery): `D:\OneDrive - lelitte.com\Desktop\Update
   HRIS Dashboard.bat`, dated 2 Jul 2026, predates BOTH the `.xls`
   support fix and the SHA-pinning fix — it still downloads
   `import_osm_report.py` from floating `main` (which does have `.xls`
   support, so that part would still work) but its `pip install` line
   is missing `xlrd`, which would make a manual `.xls` run fail on this
   specific Desktop copy unless `xlrd` happens to already be installed.
   **Not touched** — out of scope for this task ("do not touch the
   existing manual update path"), and Kevin may already plan to refresh
   it himself. Flagging only. The new automation does NOT use this
   stale Desktop copy — it downloads a fresh one from GitHub `main`
   every run, sidestepping this entirely.
2. **Sender domain note:** `reports-prd-ldz@saasiteu.com` — checked
   against this repo's own live evidence: `oxford.saasiteu.com` is the
   same domain the automated Playwright scrape has authenticated against
   all along, so this reads as SAASIT's genuine vendor domain. Not a
   definitive email-header/SPF/DKIM check — Kevin's own call on the
   actual email stands.

---

## Known pre-existing issue, unrelated to this task (do not re-fix blindly)

The GitHub Actions automated scrape path (`generate_dashboard.py` via
the self-hosted runner) is still failing — `last_run_status.txt` as of
run #120 (20 Aug 2026 16:33 UTC) shows the same `SAASIT session expired`
/ `ISM_4001` error first found by Drew on 11 Aug 2026. Needs Kevin's
interactive Oxford SSO + MFA login (`Refresh Session.bat` / `python
login.py`) — no agent can do this on his behalf. Entirely separate from
both the `.xls` fix and the new OSM email auto-refresh above — neither
of those touches or depends on the SAASIT session at all.

---

## Blockers / decisions needed from Kevin

- **RESOLVED 21 Aug 2026:** `Register-HRISAutoRefreshTask.ps1` reviewed,
  fixed (see `HANDOVER.md` Session 2026-08-21), and run — the morning
  auto-refresh Task Scheduler job is live. No longer a blocker.
- **New (evening session), informational only, no action required
  unless Kevin wants it fixed:** the stale Desktop copy of
  `Update HRIS Dashboard.bat` noted above.
- None to unblock the `.xls` fix itself — it's pushed to `main` and
  ready to use.
- Supply-chain pinning gap: fixed same day — no longer open.
- Separate, pre-existing: the SAASIT session refresh (interactive,
  Kevin-only step) for the automated GH Actions path — unrelated to
  this task.
- **Standing obligation, unchanged:** any future change to
  `import_osm_report.py` must also bump `SCRIPT_SHA` in
  `Update HRIS Dashboard.bat` in the same change, or explicitly note
  the pin is deliberately being left as-is.
