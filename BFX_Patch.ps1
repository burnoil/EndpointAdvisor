# Sample script to list relevant patches from BigFix client log without restarting service
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

# Step 2: Wait for the client to naturally re-evaluate and log (no restart needed!)
Write-Host "Waiting 90 seconds for client re-evaluation..."
Start-Sleep -Seconds 90  # Increase to 120-180 if your client takes longer to refresh

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
        Write-Host "No relevant patches found. Try increasing the wait time or check connectivity to the BigFix relay."
    }
} else {
    Write-Host "No log file found. Check BigFix installation."
}
