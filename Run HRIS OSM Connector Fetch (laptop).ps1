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
  0. Best-effort self-refresh: pull the current connector script and its
     touched shared work-inbox dependencies from main with a cache-buster and
     marker/size checks. This keeps the independent HRIS scheduled task from
     running a stale or mismatched copy after a repo update; if the network is
     unavailable, retain the last-known-good local copies and continue.
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
     Exit code 5  -> MODEL POLICY VIOLATION (added 10 Sep 2026, Priority 4)
     -- a codex_model_policy.ModelPolicyViolation propagated (e.g. an
     xhigh/max ceiling breach). Deliberately distinct from 2/3: this is a
     deterministic code/config bug, not transient flakiness -- will keep
     failing every cycle until fixed, not self-resolving on retry. Log,
     push failure status, leave the task enabled (not a mailbox-safety
     issue, so no Disable-ScheduledTask), but investigate promptly rather
     than assuming the normal cadence will clear it.
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

# --- Step 0: refresh the connector code and its touched shared dependency ---
# This task is independent of work-inbox's own laptop bridge task, so it cannot
# rely on that wrapper's refresh loop having run first. Keep the three files
# that changed in this Priority 4 implementation coherent on the production
# laptop: the HRIS entry point, the shared Lane-B executor, and the shared
# model-policy module it now imports. Each download is cache-busted and staged
# through a guarded .download file; a failed/truncated/mismatched fetch leaves
# the last-known-good local copy in place.
$workInboxRoot = Join-Path (Split-Path -Parent $root) 'work-inbox'
$refreshStamp = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
$refreshSpecs = @(
    @{ Name = 'fetch_osm_report_connector.py'; Url = 'https://raw.githubusercontent.com/begb0037admin/hris-dashboard/main/fetch_osm_report_connector.py'; Dest = (Join-Path $root 'fetch_osm_report_connector.py'); Marker = 'def fetch_osm_attachment_via_connector' }
    @{ Name = 'lane_b_call1.py'; Url = 'https://raw.githubusercontent.com/begb0037admin/work-inbox/main/lane_b_call1.py'; Dest = (Join-Path $workInboxRoot 'lane_b_call1.py'); Marker = 'def run_codex_json' }
    @{ Name = 'codex_model_policy.py'; Url = 'https://raw.githubusercontent.com/begb0037admin/work-inbox/main/codex_model_policy.py'; Dest = (Join-Path $workInboxRoot 'codex_model_policy.py'); Marker = 'class ModelPolicyViolation' }
)
foreach ($spec in $refreshSpecs) {
    $downloadPath = "$($spec.Dest).download"
    try {
        Invoke-WebRequest -UseBasicParsing "$($spec.Url)?t=$refreshStamp" -OutFile $downloadPath -TimeoutSec 30
        if ((Get-Item -LiteralPath $downloadPath).Length -lt 1000) { throw 'downloaded file too small' }
        if (-not (Select-String -Quiet -LiteralPath $downloadPath -Pattern $spec.Marker)) {
            throw "downloaded file missing marker /$($spec.Marker)/"
        }
        Move-Item -Force -LiteralPath $downloadPath -Destination $spec.Dest
        Log "refreshed $($spec.Name) from main (cache-busted, guarded)"
    } catch {
        Log "WARN: could not refresh $($spec.Name) ($($_.Exception.Message)) -- keeping local copy"
        if (Test-Path -LiteralPath $downloadPath) { Remove-Item -LiteralPath $downloadPath -Force }
    }
}

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
    5 {
        # Added 10 Sep 2026 (Priority 4, touchpoint-3 Codex review finding):
        # MUST be distinguished from `default` below, which the pre-existing
        # code treats identically to exit 2/3 -- exactly the masking this
        # exit code exists to prevent. Not a security HALT (task stays
        # enabled, unlike exit 1), but a deterministic code/config bug in
        # codex_model_policy usage, not connector unavailability -- will keep
        # failing every cycle until fixed, not self-resolving like exit 3.
        Log "MODEL POLICY VIOLATION -- a code/config bug in codex_model_policy usage (see work-inbox/codex_model_policy.py), NOT connector unavailability. Task stays enabled but this needs investigation, not just a retry."
        & $python (Join-Path $root 'push_automation_status.py') --status failure --step model_policy_violation --detail "fetch_osm_report_connector.py MODEL POLICY VIOLATION, exit 5 -- code/config bug, not connector flakiness" --attempt 1 --max-attempts 1 2>&1 | Tee-Object -FilePath $log -Append
    }
    default {
        Log "UNEXPECTED exit code $fetchRc -- treating conservatively as a failure, task stays enabled."
        & $python (Join-Path $root 'push_automation_status.py') --status failure --step unexpected_exit --detail "fetch_osm_report_connector.py exited $fetchRc (unexpected)" --attempt 1 --max-attempts 1 2>&1 | Tee-Object -FilePath $log -Append
    }
}

Log "=== Run HRIS OSM Connector Fetch finished ==="
