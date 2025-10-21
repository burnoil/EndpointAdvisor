# --- Microsoft 365 Apps (ODT) Install ---
try {
    Write-ADTLogEntry -Message "Starting Microsoft 365 Apps installation..." -Source $adtSession.InstallPhase

    $setupExe  = Join-Path -Path $dirFiles -ChildPath 'Setup.exe'
    $configXml = Join-Path -Path $dirFiles -ChildPath 'configuration.xml'

    if (-not (Test-Path -LiteralPath $setupExe))  { throw "ODT Setup.exe not found: $setupExe" }
    if (-not (Test-Path -LiteralPath $configXml)) { throw "ODT configuration.xml not found: $configXml" }

    Start-ADTProcess -FilePath $setupExe `
                     -ArgumentList "/configure `"$configXml`"" `
                     -WorkingDirectory $dirFiles `
                     -WindowStyle Hidden `
                     -CreateNoWindow:$true

    Write-ADTLogEntry -Message "Microsoft 365 Apps installation completed." -Source $adtSession.InstallPhase
}
catch {
    Write-ADTLogEntry -Message "ODT install failed: $($_.Exception.Message)" -Severity 3 -Source $adtSession.InstallPhase
    throw
}
