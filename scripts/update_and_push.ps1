# Regenerates data/tracker.json from the live source spreadsheets and pushes it
# to GitHub if anything changed. Meant to be run every 12 hours by a Windows
# Scheduled Task on a machine with the Dropbox/OneDrive drives mounted.
#
# Registered task name: LSE-CTUR-Tracker-Update

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

$logFile = Join-Path $repoRoot "scripts\update.log"
function Log($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $msg"
    Write-Output $line
    Add-Content -Path $logFile -Value $line
}

try {
    Log "Running generate_data.py..."
    py -3 scripts\generate_data.py 2>&1 | ForEach-Object { Log $_ }

    $status = git status --porcelain data\tracker.json
    if ([string]::IsNullOrWhiteSpace($status)) {
        Log "No change in data/tracker.json -- nothing to push."
    } else {
        Log "Data changed -- committing and pushing."
        git add data\tracker.json
        git commit -m "Auto-update tracker data ($(Get-Date -Format 'yyyy-MM-dd HH:mm'))"
        git push
        Log "Pushed."
    }
} catch {
    Log "ERROR: $($_.Exception.Message)"
    throw
}
