@echo off
title HRIS Dashboard — OSM Import
echo.
echo === HRIS Dashboard Update ===
echo.

:: Pull latest script from GitHub
echo Downloading latest import_osm_report.py...
curl -s -o "%TEMP%\import_osm_report.py" "https://raw.githubusercontent.com/begb0037admin/hris-dashboard/main/import_osm_report.py"
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
