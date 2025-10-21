# --- Microsoft 365 Apps (ODT) Install ---
Write-Log -Message "Starting Microsoft 365 Apps installation..." -Source $installPhase

$setupExe  = Join-Path -Path $dirFiles -ChildPath 'Setup.exe'
$configXml = Join-Path -Path $dirFiles -ChildPath 'configuration.xml'

if (-not (Test-Path -LiteralPath $setupExe)) {
    Write-Log -Message "ODT Setup.exe not found at $setupExe" -Severity 3 -Source $installPhase
    Throw "ODT Setup.exe missing"
}

if (-not (Test-Path -LiteralPath $configXml)) {
    Write-Log -Message "ODT configuration.xml not found at $configXml" -Severity 3 -Source $installPhase
    Throw "ODT configuration.xml missing"
}

Execute-Process -Path $setupExe `
                -Parameters "/configure `"$configXml`"" `
                -WorkingDirectory $dirFiles `
                -WindowStyle Hidden `
                -CreateNoWindow:$true

Write-Log -Message "Microsoft 365 Apps installation completed." -Source $installPhase
