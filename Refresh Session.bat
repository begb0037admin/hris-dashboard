@echo off
cd /d "C:\actions-runner\_work\hris-dashboard\hris-dashboard"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$t=[DateTimeOffset]::UtcNow.ToUnixTimeSeconds(); Invoke-WebRequest -UseBasicParsing ('https://raw.githubusercontent.com/begb0037admin/hris-dashboard/main/login.py?t='+$t) -OutFile 'login.py'"
python login.py
pause
