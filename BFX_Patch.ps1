# Sample script to force re-evaluation and list relevant patches from BigFix client log
# Run as administrator. Adjust paths if your BigFix installation is different.

$bigfixPath = "C:\Program Files (x86)\BigFix Enterprise\BES Client"
$logPath = "$bigfixPath\__BESData\__Global\Logs"
$statePath = "$bigfixPath\__BESData"

# Step 1: Identify and delete state files for patch sites (force re-evaluation)
$patchSites = Get-ChildItem -Path $statePath -Directory | Where-Object { $_.Name -match "Patch|Update|Security" }
foreach ($site in $patchSites) {
    $siteState = "$statePath\$($site.Name)"
    Remove-Item -Path "$siteState\*" -Force -Recurse -ErrorAction SilentlyContinue
    Write-Host "Cleared state for site: $($site.Name)"
}

# Step 2: Restart BigFix client service to trigger evaluation and logging
Stop-Service -Name BESClient -Force
Start-Service -Name BESClient
Write-Host "BigFix client restarted. Waiting 60 seconds for evaluation..."
Start-Sleep -Seconds 60  # Adjust if needed; evaluation may take 30-120 seconds

# Step 3: Find the latest log file
$latestLog = Get-ChildItem -Path $logPath -Filter "*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($latestLog) {
    Write-Host "Parsing latest log: $($latestLog.FullName)"
    $relevantPatches = @()

    # Read the log and extract "Relevant - " lines for patches
    Get-Content $latestLog.FullName | ForEach-Object {
        if ($_ -match "Relevant - (.+?)\(id: \d+\) \((Patches|Updates|Security).*?\)") {
            $patchName = $matches[1].Trim()
            $site = $matches[2].Trim()
            $relevantPatches += "$patchName (Site: $site)"
        }
    }

    # Remove duplicates and output
    $relevantPatches | Sort-Object -Unique | ForEach-Object { Write-Host $_ }
    if ($relevantPatches.Count -eq 0) {
        Write-Host "No relevant patches found in the log after re-evaluation."
    }
} else {
    Write-Host "No log file found. Check BigFix installation."
}
