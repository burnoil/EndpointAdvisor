# --- ODT (Office Deployment Tool) pre-flight + install ---
try {
    $setupExe  = Join-Path -Path $adtSession.DirFiles -ChildPath 'Setup.exe'
    $configXml = Join-Path -Path $adtSession.DirFiles -ChildPath 'configuration.xml'

    if (-not (Test-Path -LiteralPath $setupExe)) {
        Write-ADTLogEntry -Message "ODT payload missing: $setupExe" -Severity 3 -Source $adtSession.InstallPhase
        throw "ODT Setup.exe not found at: $setupExe"
    }
    if (-not (Test-Path -LiteralPath $configXml)) {
        Write-ADTLogEntry -Message "ODT configuration missing: $configXml" -Severity 3 -Source $adtSession.InstallPhase
        throw "ODT configuration.xml not found at: $configXml"
    }

    # Important: set WorkingDirectory to the Files folder (handles spaces in paths)
    Write-ADTLogEntry -Message "Launching ODT: `"$setupExe`" /configure `"$configXml`"" -Severity 1 -Source $adtSession.InstallPhase
    Start-ADTProcess -FilePath $setupExe `
                     -ArgumentList "/configure `"$configXml`"" `
                     -WorkingDirectory $adtSession.DirFiles
}
catch {
    Write-ADTLogEntry -Message "ODT install failed: $($_.Exception.Message)" -Severity 3 -Source $adtSession.InstallPhase

    # Try to surface the most recent ODT log to speed up diagnosis
    try {
        $odtLog = Get-ChildItem -Path (Join-Path $env:WINDIR 'Temp') -Filter 'OfficeDeploymentTool*.log' -ErrorAction SilentlyContinue |
                  Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($odtLog) {
            Write-ADTLogEntry -Message "Latest ODT log: $($odtLog.FullName)" -Severity 2 -Source $adtSession.InstallPhase
        }
    } catch { }

    throw
}
