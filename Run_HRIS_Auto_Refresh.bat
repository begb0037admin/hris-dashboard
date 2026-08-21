@echo off
setlocal EnableExtensions
title HRIS Dashboard - Auto Morning Refresh

rem ============================================================================
rem Run_HRIS_Auto_Refresh.bat
rem
rem Additive automation only. Does NOT modify or replace the existing manual
rem path (Update HRIS Dashboard.bat / import_osm_report.py) -- it downloads
rem and runs that exact, unmodified file as its second step, exactly as if
rem Kevin had double-clicked it himself after downloading the report by hand.
rem Kevin's manual midday updates continue to work exactly as before; this
rem script's only job is to do the same two steps automatically at ~8:30am:
rem   1. Pull today's OSM report attachment out of Outlook (fetch_osm_report.py)
rem   2. Run the existing Update HRIS Dashboard.bat pipeline, unattended
rem       (its own "pause" prompts are satisfied via `< NUL`, not skipped or
rem        edited -- the file itself is untouched)
rem
rem fetch_osm_report.py, Update HRIS Dashboard.bat, and push_automation_status.py
rem are all pulled fresh from GitHub `main` every run (cache-busted), per this
rem repo's own AGENT_MODEL.md Section 1: local execution scripts pull their
rem latest version from GitHub before every run rather than operating as a
rem hand-edited standing local copy.
rem
rem If step 1 fails (no matching email received yet today, sender/subject/
rem attachment format changed, Outlook unreachable), step 2 is SKIPPED
rem entirely and this script exits non-zero -- the dashboard is left as-is
rem rather than being silently "refreshed" against stale/missing data.
rem Task Scheduler's own Last Run Result column then shows a real failure,
rem and the hidden-task wrapper fires a failure toast (see
rem "Run HRIS Auto-Refresh Hidden.vbs" on the Desktop) -- so a silent break
rem is detectable rather than invisible, per Kevin's explicit requirement.
rem ============================================================================
rem
rem ROOT CAUSE FIXED 21 Aug 2026 (Drew) -- the real bug behind "Step 2 never
rem starts, nothing in the log, nothing in Windows event logs" (21 Aug 08:30
rem run and every run before it):
rem
rem The old version's failure-branch echo read:
rem   echo [%RUN_TS%] Step 1/2 FAILED (exit %FETCH_EXIT%) - Step 2 SKIPPED. ...
rem which sits inside a multi-line `if not "%FETCH_EXIT%"=="0" ( ... )` block.
rem cmd.exe's block parser ends a parenthesized if/for block at the FIRST
rem unescaped ")" it meets while scanning -- it does not "balance" an earlier
rem "(" against it -- so the ")" right after %FETCH_EXIT% was read as the
rem block's real closing paren. Every line the script intended to run after
rem that point was misparsed as stray top-level tokens starting with the
rem literal text " - Step 2 SKIPPED...", whose leading "-" is exactly cmd's
rem classic "- was unexpected at this time." syntax error -- written only to
rem the console (invisible, since this always runs fully hidden via
rem wscript.exe windowStyle=0), never to the log file. That is exactly why
rem this looked like total silence: no error in the log, nothing in Windows
rem event logs (Task Scheduler only ever saw wscript.exe's own exit code,
rem 0x800700FF, with no further detail).
rem
rem This bug fired on EVERY run regardless of Step 1's real outcome -- cmd.exe
rem must fully parse an if-block to know where to resume even when the
rem block's condition is false and the block is being skipped, which is
rem exactly what happened on the 21 Aug 08:30 run (Step 1 genuinely
rem succeeded; FETCH_EXIT was "0"; the block should never have been entered
rem at all, yet the script still died parsing it).
rem
rem Confirmed via bisection, not guesswork: an isolated single-line
rem reproduction of the broken echo inside a trivial if-block reliably threw
rem the same "- was unexpected at this time." / exit 255; rewording that one
rem line to avoid literal parens (this version) reliably fixed it, both in
rem isolation and by re-running this exact real script end-to-end against
rem the live mailbox. Full bisection trail in RESUME.md / HANDOVER.md,
rem session 21 Aug 2026.
rem
rem SECOND change, same session -- finer-grained diagnostics so a *future*
rem silent break is impossible, not just this specific one: the whole run
rem now happens inside :main, called as `call :main > "%LOG_FILE%" 2>&1` --
rem ONE redirection covering the entire subroutine, both stdout and stderr,
rem instead of the old scattered per-line ">> %LOG_FILE%" appends. Any
rem future cmd.exe parser error -- not just this one -- lands in the log
rem file automatically, because the redirection applies to everything that
rem runs during that call, including cmd's own error output. Before this
rem change, a script-level syntax error was structurally invisible no matter
rem how much per-line logging existed, because execution never reached any
rem of those log-append statements once cmd's parser choked. Do NOT add new
rem literal "(" or ")" characters inside any multi-line if/for block's body
rem here without a compelling reason -- reword to avoid them (as done above),
rem or escape as ^( ^) and test with bisection before trusting it.
rem
rem THIRD change, same session -- after every exit point (success or any
rem failure), pushes a small status record to GitHub
rem (data/last_automated_run.json, via push_automation_status.py, downloaded
rem fresh from GitHub like the other two scripts) so the live dashboard can
rem show a real "did this morning's automated run happen, and did it work"
rem indicator -- sourced from the automation's own run record, not inferred
rem from tickets.json's mtime (a failed run leaving yesterday's data in place
rem would otherwise look identical to a fresh successful run).
rem
rem FOURTH change, same session -- a SECOND, independent root cause found
rem after fixing the paren bug above: even with that fixed, Step 2 still
rem failed, now with a clean "'"Update HRIS Dashboard.bat"' is not
rem recognized as an internal or external command" instead of a parser
rem crash. Confirmed by direct isolated testing (safe dummy .bat, not the
rem real file) that this machine's cmd.exe cannot resolve or launch a
rem SECOND .bat/.cmd FILE (via `call`, bare invocation, or a fresh
rem `cmd /c "x.bat"`) from a cmd.exe process that is already itself
rem executing a batch file -- confirmed this is specific to script-file
rem targets, not a blanket "cmd can't spawn cmd" block: nested `cmd /c dir`,
rem `cmd /c echo`, and even double-nested `cmd /c "cmd /c echo"` all worked
rem fine from the same nested context, and curl.exe/python.exe (real .exe
rem targets) already worked fine too, live, in Step 1 above. Only a second
rem SCRIPT FILE is blocked. Confirmed the fix empirically before applying
rem it for real: routing the same invocation through
rem `powershell -NoProfile -Command "... & '.\x.bat' ..."` as an
rem intermediary succeeds every time, including with a real `pause` inside
rem the target .bat (verified it does not hang and the real exit code
rem propagates correctly) -- see below. Update HRIS Dashboard.bat itself
rem remains completely unmodified; only HOW this script launches it changed.
rem ============================================================================

set "PROJECT_DIR=C:\Users\admin\Documents\Claude\Projects\HRIS-Dashboard"
set "LOG_FILE=%PROJECT_DIR%\osm_auto_refresh_last_run.log"

if not exist "%PROJECT_DIR%\" mkdir "%PROJECT_DIR%"
cd /d "%PROJECT_DIR%"

call :main > "%LOG_FILE%" 2>&1
set "FINAL_EXIT=%ERRORLEVEL%"
type "%LOG_FILE%"
exit /b %FINAL_EXIT%

rem ============================================================================
rem :main -- everything below runs as ONE unit, redirected to %LOG_FILE% by the
rem `call :main` line above. Do not add per-line ">> %LOG_FILE%" redirections
rem inside here -- a plain "echo" already lands in the log because of that
rem outer redirection, and doubling it up would just duplicate lines.
rem ============================================================================
:main

for /f "delims=" %%I in ('python -c "from datetime import datetime; print(datetime.now().isoformat(sep=' ', timespec='seconds'))" 2^>nul') do set "RUN_TS=%%I"
if "%RUN_TS%"=="" set "RUN_TS=(timestamp unavailable)"

set "FETCH_SCRIPT=fetch_osm_report.py"
set "FETCH_RAW_URL=https://raw.githubusercontent.com/begb0037admin/hris-dashboard/main/fetch_osm_report.py"
set "UPDATE_BAT_NAME=Update HRIS Dashboard.bat"
rem %%20 = URL-encoded space -- a literal space in the URL was tested and
rem confirmed to fail (curl could not resolve it), so this must stay encoded.
set "UPDATE_RAW_URL=https://raw.githubusercontent.com/begb0037admin/hris-dashboard/main/Update%%20HRIS%%20Dashboard.bat"
set "STATUS_SCRIPT=push_automation_status.py"
set "STATUS_RAW_URL=https://raw.githubusercontent.com/begb0037admin/hris-dashboard/main/push_automation_status.py"

rem Empty file used to satisfy Update HRIS Dashboard.bat's own "pause"
rem non-interactively (see the PowerShell-intermediary invocation below) --
rem `copy nul` is the standard batch idiom for creating a zero-byte file.
copy /y nul "empty_stdin.txt" >nul

echo [%RUN_TS%] Run HRIS Auto-Refresh started
echo [%RUN_TS%] Downloading fetch_osm_report.py and push_automation_status.py from GitHub...

rem -f (--fail) makes curl return a real nonzero exit code on an HTTP error
rem (404/5xx) instead of silently writing the error page's body as if it
rem were the file -- confirmed live this session: without -f, a 404 for
rem push_automation_status.py before it existed on `main` was saved as a
rem 14-byte file containing the literal text "404: Not Found", which python
rem then tried to run and failed on with a confusing SyntaxError instead of
rem a clear download-failed message.
curl -sf -o "%STATUS_SCRIPT%" "%STATUS_RAW_URL%?t=%RANDOM%%RANDOM%"
if errorlevel 1 (
    echo [%RUN_TS%] WARNING - could not download push_automation_status.py. Continuing anyway; this run's status will not reach the dashboard.
)

curl -sf -o "%FETCH_SCRIPT%" "%FETCH_RAW_URL%?t=%RANDOM%%RANDOM%"
if errorlevel 1 (
    echo [%RUN_TS%] ERROR - could not download fetch_osm_report.py. Dashboard NOT refreshed.
    if exist "%STATUS_SCRIPT%" python "%STATUS_SCRIPT%" --status failure --step fetch_script_download_failed --detail "curl could not download fetch_osm_report.py from GitHub"
    exit /b 1
)

echo [%RUN_TS%] Step 1/2 - fetching today's OSM report attachment from Outlook...
python "%FETCH_SCRIPT%"
set "FETCH_EXIT=%ERRORLEVEL%"

if not "%FETCH_EXIT%"=="0" (
    echo [%RUN_TS%] Step 1/2 FAILED, exit code %FETCH_EXIT% - Step 2 SKIPPED. Dashboard NOT refreshed this run.
    if exist "%STATUS_SCRIPT%" python "%STATUS_SCRIPT%" --status failure --step fetch_failed --detail "fetch_osm_report.py exited %FETCH_EXIT% - see osm_auto_refresh_last_run.log"
    exit /b %FETCH_EXIT%
)

echo [%RUN_TS%] Step 1/2 OK. Step 2/2 - downloading current Update HRIS Dashboard.bat from GitHub...
curl -sf -o "%UPDATE_BAT_NAME%" "%UPDATE_RAW_URL%?t=%RANDOM%%RANDOM%"
if errorlevel 1 (
    echo [%RUN_TS%] ERROR - could not download "%UPDATE_BAT_NAME%". Dashboard NOT refreshed.
    if exist "%STATUS_SCRIPT%" python "%STATUS_SCRIPT%" --status failure --step update_bat_download_failed --detail "curl could not download Update HRIS Dashboard.bat from GitHub"
    exit /b 1
)

echo [%RUN_TS%] Step 2/2 - running the existing, unmodified dashboard update pipeline...
rem This machine's cmd.exe cannot launch a second .bat/.cmd FILE (call, bare
rem invocation, or a fresh `cmd /c "x.bat"`) from a cmd.exe process that is
rem itself already executing a batch file -- confirmed by isolated testing,
rem see the FOURTH change note near the top of this file. Routing through
rem PowerShell as the intermediary sidesteps it. `Get-Content
rem 'empty_stdin.txt' | & '.\...'` pipes an immediate EOF into the call
rem operator so Update HRIS Dashboard.bat's own "pause" proceeds instantly
rem instead of hanging forever against a hidden, non-interactive console --
rem the same job `< NUL` used to do for the old (broken-here) `call`
rem invocation. The file itself, and everything it does (pinned
rem import_osm_report.py download, pip install, GitHub push of
rem data/tickets.json), is completely unchanged.
powershell -NoProfile -Command "Get-Content 'empty_stdin.txt' | & '.\%UPDATE_BAT_NAME%'; exit $LASTEXITCODE"
set "UPDATE_EXIT=%ERRORLEVEL%"

if "%UPDATE_EXIT%"=="0" (
    echo [%RUN_TS%] HRIS dashboard refreshed successfully via automated morning run.
    if exist "%STATUS_SCRIPT%" python "%STATUS_SCRIPT%" --status success --step dashboard_update_ok --detail "Automated morning run completed successfully."
) else (
    echo [%RUN_TS%] "%UPDATE_BAT_NAME%" failed, exit code %UPDATE_EXIT%.
    if exist "%STATUS_SCRIPT%" python "%STATUS_SCRIPT%" --status failure --step dashboard_update_failed --detail "Update HRIS Dashboard.bat exited %UPDATE_EXIT% - see osm_auto_refresh_last_run.log"
)

exit /b %UPDATE_EXIT%
