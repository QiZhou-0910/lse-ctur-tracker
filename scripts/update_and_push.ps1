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
    # Note: deliberately not using 2>&1 here -- PowerShell wraps a native
    # command's stderr lines as terminating ErrorRecords under
    # $ErrorActionPreference = "Stop", which would abort this script even
    # on success. Check $LASTEXITCODE instead.
    $out = & py -3 scripts\generate_data.py
    $out | ForEach-Object { Log $_ }
    if ($LASTEXITCODE -ne 0) {
        throw "generate_data.py exited with code $LASTEXITCODE"
    }

    $status = git status --porcelain data\tracker.json
    if ([string]::IsNullOrWhiteSpace($status)) {
        Log "No change in data/tracker.json -- nothing to push."
    } else {
        Log "Data changed -- committing and pushing."
        git add data\tracker.json
        git commit -m "Auto-update tracker data ($(Get-Date -Format 'yyyy-MM-dd HH:mm'))"
        git push
        if ($LASTEXITCODE -ne 0) {
            throw "git push exited with code $LASTEXITCODE"
        }
        Log "Pushed."
    }
} catch {
    Log "ERROR: $($_.Exception.Message)"
    throw
}
