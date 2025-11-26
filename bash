# 1. Build the full path safely
$ScriptPath = Join-Path -Path $adtSession.DirFiles -ChildPath "BlackHoleDiskCleaner.ps1"

# 2. Check if it actually exists (Optional, but saves sanity)
If (-not (Test-Path -Path $ScriptPath)) {
    Write-ADTLogEntry -Message "ERROR: Could not find script at [$ScriptPath]" -Severity 3
    Throw "Script file not found: $ScriptPath"
}

# 3. Run it using the clean variable
Start-ADTProcess -FilePath "$PSHOME\powershell.exe" -ArgumentList "-NoLogo -NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`" -LocalRun -Silent -SkipRecycleBin -SkipBrowserCache"
