<#
Run HRIS OSM Connector Fetch (laptop).ps1
==========================================
Connector-based replacement for the desktop's Run_HRIS_Auto_Refresh.bat ->
fetch_osm_report.py (classic Outlook COM) path. Runs on the Oxford laptop
(101L-DE013193, user AD-OAK\begb0037), same account/session pattern as
work-inbox's own "Work Inbox Bridge Briefing" task.

See MIGRATION-IMAP.md for the full design. Same-day incident this was
proven against, 9 Sep 2026: the desktop's classic-Outlook-COM path
(fetch_osm_report.py) genuinely failed all 3 scheduled attempts this
morning (exit 1, "fetch_failed", data\last_automated_run.json confirms);
this connector-based fetch succeeded on the very first attempt against the
same live mailbox the same day, and was used to manually clear that
failure (see HANDOVER.md / RESUME.md for the exact commits).

Steps:
  1. Retry-guard: skip entirely if data\last_automated_run.json already
     shows a genuine success for TODAY (mirrors the desktop .bat's own
     "don't re-import an already-current file" discipline, and prevents 3
     redundant connector calls across the 08:45/09:15/09:45 cadence once
     one has already succeeded).
  2. Run fetch_osm_report_connector.py (real target, not --dry-run).
  3. Exit code 0  -> run import_osm_report.py (OSM_IMPORT_FILE = the saved
     path) then push_automation_status.py --status success.
     Exit code 1  -> GUARD HALT. A write/off-scope connector tool call was
     observed. Disable THIS task (fail closed, same shape as work-inbox's
     mail-guard HALT response) and push a failure status. Does NOT touch
     the desktop's COM path automatically -- that decision needs a human
     look at what actually happened, this is a rare/unexpected event by
     design.
     Exit code 2  -> usage/environment error. Log, push failure status,
     leave the task enabled (transient, e.g. codex not resolvable this
     run) -- the normal 08:45/09:15/09:45 cadence retries on its own.
     Exit code 3  -> connector unavailable this cycle (headless tool-
     loading flakiness, documented, not a safety event). Log, push
     failure status, leave the task enabled, cadence retries.
     Exit code 4  -> no matching OSM report email found for today yet.
     Log, push failure status (informational, not alarming), leave the
     task enabled, cadence retries.
#>

$ErrorActionPreference = 'Continue'
$root      = 'C:\Users\begb0037.AD-OAK\hris-dashboard'
$python    = 'C:\Users\begb0037.AD-OAK\AppData\Local\Programs\Python\Python312\python.exe'
$targetDir = Join-Path $root 'data\osm_fetch'
$logDir    = Join-Path $root 'data\osm_fetch\logs'
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$log   = Join-Path $logDir "run_$stamp.log"

function Log($msg) {
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $msg"
    Write-Output $line
    $line | Out-File -FilePath $log -Append -Encoding utf8
}

Log "=== Run HRIS OSM Connector Fetch started ==="
Set-Location $root

# --- Step 1: retry-guard -----------------------------------------------
try {
    $statusRaw = Invoke-RestMethod -Uri "https://raw.githubusercontent.com/begb0037admin/hris-dashboard/main/data/last_automated_run.json?nocache=$(Get-Random)" -TimeoutSec 20
    $todayStr = Get-Date -Format 'yyyy-MM-dd'
    if ($statusRaw.status -eq 'success' -and $statusRaw.timestamp -like "$todayStr*") {
        Log "SKIP: last_automated_run.json already shows a genuine success for today ($($statusRaw.timestamp)) -- nothing to do this cycle."
        exit 0
    }
} catch {
    Log "WARN: could not read last_automated_run.json for the retry-guard check ($($_.Exception.Message)) -- proceeding anyway (fail open on the guard itself, the fetch step below has its own safety)."
}

# --- Step 2: connector fetch --------------------------------------------
Log "running: python fetch_osm_report_connector.py"
& $python (Join-Path $root 'fetch_osm_report_connector.py') 2>&1 | Tee-Object -FilePath $log -Append
$fetchRc = $LASTEXITCODE
Log "fetch_osm_report_connector.py exit $fetchRc"

switch ($fetchRc) {
    0 {
        # find the just-saved file (fixed filename base, extension varies)
        $saved = Get-ChildItem -Path $targetDir -Filter 'All Open Tasks by Team - auto.*' -ErrorAction SilentlyContinue |
                 Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if (-not $saved) {
            Log "ERROR: fetch reported success (exit 0) but no saved file found in $targetDir -- treating as a failure, not importing anything."
            & $python (Join-Path $root 'push_automation_status.py') --status failure --step fetch_ok_but_file_missing --detail "fetch_osm_report_connector.py exited 0 but no output file found" --attempt 1 --max-attempts 1 2>&1 | Tee-Object -FilePath $log -Append
            exit 1
        }
        Log "running: python import_osm_report.py (OSM_IMPORT_FILE=$($saved.FullName))"
        $env:OSM_IMPORT_FILE = $saved.FullName
        $env:OSM_IMPORT_NONINTERACTIVE = '1'
        & $python (Join-Path $root 'import_osm_report.py') --yes 2>&1 | Tee-Object -FilePath $log -Append
        $importRc = $LASTEXITCODE
        Log "import_osm_report.py exit $importRc"
        if ($importRc -eq 0) {
            & $python (Join-Path $root 'push_automation_status.py') --status success --step dashboard_update_ok --detail "connector-based fetch (fetch_osm_report_connector.py)" --attempt 1 --max-attempts 1 2>&1 | Tee-Object -FilePath $log -Append
        } else {
            & $python (Join-Path $root 'push_automation_status.py') --status failure --step import_failed --detail "import_osm_report.py exited $importRc" --attempt 1 --max-attempts 1 2>&1 | Tee-Object -FilePath $log -Append
        }
    }
    1 {
        Log "GUARD HALT: a write/off-scope connector tool call was observed. Disabling this task -- investigate before re-enabling. See data\osm_fetch\transcripts\ for the raw evidence."
        & $python (Join-Path $root 'push_automation_status.py') --status failure --step guard_halt --detail "fetch_osm_report_connector.py GUARD HALT -- unexpected connector tool call observed, task disabled pending investigation" --attempt 1 --max-attempts 1 2>&1 | Tee-Object -FilePath $log -Append
        try {
            Disable-ScheduledTask -TaskName 'HRIS Dashboard Morning Refresh (connector)' -ErrorAction Stop | Out-Null
            Log "Task 'HRIS Dashboard Morning Refresh (connector)' disabled."
        } catch {
            Log "WARN: could not disable the scheduled task automatically ($($_.Exception.Message)) -- disable it manually."
        }
    }
    2 {
        Log "USAGE/ENVIRONMENT ERROR (not a safety event). Task stays enabled; normal cadence retries."
        & $python (Join-Path $root 'push_automation_status.py') --status failure --step env_error --detail "fetch_osm_report_connector.py usage/environment error, exit 2" --attempt 1 --max-attempts 1 2>&1 | Tee-Object -FilePath $log -Append
    }
    3 {
        Log "Connector unavailable this cycle (headless tool-loading flakiness, documented, not a safety event). Task stays enabled; normal cadence retries."
        & $python (Join-Path $root 'push_automation_status.py') --status failure --step connector_unavailable --detail "fetch_osm_report_connector.py connector unavailable this cycle, exit 3" --attempt 1 --max-attempts 1 2>&1 | Tee-Object -FilePath $log -Append
    }
    4 {
        Log "No matching OSM report email found for today yet. Task stays enabled; normal cadence retries."
        & $python (Join-Path $root 'push_automation_status.py') --status failure --step fetch_failed --detail "no matching OSM report email found for today yet" --attempt 1 --max-attempts 1 2>&1 | Tee-Object -FilePath $log -Append
    }
    default {
        Log "UNEXPECTED exit code $fetchRc -- treating conservatively as a failure, task stays enabled."
        & $python (Join-Path $root 'push_automation_status.py') --status failure --step unexpected_exit --detail "fetch_osm_report_connector.py exited $fetchRc (unexpected)" --attempt 1 --max-attempts 1 2>&1 | Tee-Object -FilePath $log -Append
    }
}

Log "=== Run HRIS OSM Connector Fetch finished ==="
