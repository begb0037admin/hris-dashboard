@echo off
title HRIS Dashboard — OSM Import
echo.
echo === HRIS Dashboard Update ===
echo.

:: Pull import_osm_report.py from GitHub, PINNED to a known-good commit SHA
:: (not the "main" branch) -- this is intentional supply-chain pinning, not
:: staleness. Fetching "main" directly would run whatever the latest commit
:: on the repo happens to be at run time, with zero integrity check; pinning
:: to a SHA means this .bat always fetches the exact byte-for-byte version
:: below, verified working, regardless of what else lands on main later.
::
:: SHA below = main's tip as of 20 Aug 2026, includes the .xls import fix
:: (commit e503509d2785cc30d4102365aa72e353983f864d).
::
:: TO BUMP: whenever import_osm_report.py changes on main and that change
:: needs to actually ship, update SCRIPT_SHA below to the new commit SHA
:: you want pinned (verify it live: `gh api repos/begb0037admin/hris-dashboard/commits/main --jq .sha`
:: or check the commit history on GitHub) and push this .bat. Until you do
:: that, this .bat will keep silently running today's pinned version
:: forever, even after main moves on -- that is the point of pinning, but
:: it means a future script fix will NOT take effect here until someone
:: deliberately bumps this SHA.
set SCRIPT_SHA=e503509d2785cc30d4102365aa72e353983f864d

echo Downloading import_osm_report.py (pinned to %SCRIPT_SHA%)...
curl -s -o "%TEMP%\import_osm_report.py" "https://raw.githubusercontent.com/begb0037admin/hris-dashboard/%SCRIPT_SHA%/import_osm_report.py"
if errorlevel 1 (
    echo ERROR: Could not download script. Check your internet connection.
    pause
    exit /b 1
)
echo Done.
echo.

:: Ensure dependencies are installed
echo Checking dependencies...
pip install openpyxl xlrd requests --quiet --disable-pip-version-check
echo Done.
echo.

:: Run it
python "%TEMP%\import_osm_report.py"

echo.
pause
