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
rem Both downloaded scripts are pulled fresh from GitHub `main` every run
rem (cache-busted), per this repo's own AGENT_MODEL.md Section 1: local
rem execution scripts pull their latest version from GitHub before every
rem run rather than operating as a hand-edited standing local copy.
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

rem ISO-ish timestamp via Python (isoformat() takes no %%-style format
rem codes, so there's nothing here for cmd.exe to misinterpret).
for /f "delims=" %%I in ('python -c "from datetime import datetime; print(datetime.now().isoformat(sep=' ', timespec='seconds'))" 2^>nul') do set "RUN_TS=%%I"
if "%RUN_TS%"=="" set "RUN_TS=(timestamp unavailable)"

set "PROJECT_DIR=C:\Users\admin\Documents\Claude\Projects\HRIS-Dashboard"
set "FETCH_SCRIPT=fetch_osm_report.py"
set "FETCH_RAW_URL=https://raw.githubusercontent.com/begb0037admin/hris-dashboard/main/fetch_osm_report.py"
set "UPDATE_BAT_NAME=Update HRIS Dashboard.bat"
rem %%20 = URL-encoded space -- a literal space in the URL was tested and
rem confirmed to fail (curl could not resolve it), so this must stay encoded.
set "UPDATE_RAW_URL=https://raw.githubusercontent.com/begb0037admin/hris-dashboard/main/Update%%20HRIS%%20Dashboard.bat"
set "LOG_FILE=%PROJECT_DIR%\osm_auto_refresh_last_run.log"

if not exist "%PROJECT_DIR%\" mkdir "%PROJECT_DIR%"
cd /d "%PROJECT_DIR%"

echo [%RUN_TS%] Run HRIS Auto-Refresh started > "%LOG_FILE%"
echo [%RUN_TS%] Run HRIS Auto-Refresh started

echo [%RUN_TS%] Step 1/2 - downloading latest fetch_osm_report.py from GitHub... >> "%LOG_FILE%"
curl -s -o "%FETCH_SCRIPT%" "%FETCH_RAW_URL%?t=%RANDOM%%RANDOM%"
if errorlevel 1 (
    echo [%RUN_TS%] ERROR - could not download fetch_osm_report.py. Dashboard NOT refreshed. >> "%LOG_FILE%"
    echo ERROR - could not download fetch_osm_report.py. Check internet connection.
    exit /b 1
)

echo [%RUN_TS%] Step 1/2 - fetching today's OSM report attachment from Outlook... >> "%LOG_FILE%"
echo Step 1/2 - fetching today's OSM report attachment from Outlook...
python "%FETCH_SCRIPT%" >> "%LOG_FILE%" 2>&1
set "FETCH_EXIT=%ERRORLEVEL%"
type "%LOG_FILE%"

if not "%FETCH_EXIT%"=="0" (
    echo. >> "%LOG_FILE%"
    echo [%RUN_TS%] Step 1/2 FAILED (exit %FETCH_EXIT%) - Step 2 SKIPPED. Dashboard NOT refreshed this run. >> "%LOG_FILE%"
    echo.
    echo Step 1/2 FAILED - see %LOG_FILE%
    exit /b %FETCH_EXIT%
)

echo. >> "%LOG_FILE%"
echo [%RUN_TS%] Step 1/2 OK. Step 2/2 - downloading current Update HRIS Dashboard.bat from GitHub... >> "%LOG_FILE%"
echo Step 2/2 - running the existing (unmodified) dashboard update pipeline...
curl -s -o "%UPDATE_BAT_NAME%" "%UPDATE_RAW_URL%?t=%RANDOM%%RANDOM%"
if errorlevel 1 (
    echo [%RUN_TS%] ERROR - could not download "%UPDATE_BAT_NAME%". Dashboard NOT refreshed. >> "%LOG_FILE%"
    echo ERROR - could not download "%UPDATE_BAT_NAME%".
    exit /b 1
)

rem `< NUL` satisfies the downloaded .bat's own "pause" without editing that
rem file -- stdin redirected from NUL gives an immediate EOF, so `pause`
rem proceeds instantly instead of waiting for a keypress. The file itself,
rem and everything it does (pinned import_osm_report.py download, pip
rem install, GitHub push of data/tickets.json), is completely unchanged.
call "%UPDATE_BAT_NAME%" < NUL >> "%LOG_FILE%" 2>&1
set "UPDATE_EXIT=%ERRORLEVEL%"
type "%LOG_FILE%"

echo. >> "%LOG_FILE%"
if "%UPDATE_EXIT%"=="0" (
    echo [%RUN_TS%] HRIS dashboard refreshed successfully via automated morning run. >> "%LOG_FILE%"
    echo.
    echo Done - dashboard refreshed.
) else (
    echo [%RUN_TS%] "%UPDATE_BAT_NAME%" failed with exit code %UPDATE_EXIT%. >> "%LOG_FILE%"
    echo.
    echo Step 2/2 FAILED - see %LOG_FILE%
)

exit /b %UPDATE_EXIT%
