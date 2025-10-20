Setup.exe /configure Remove.xml


<Configuration>
  <Remove All="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="SharedComputerLicensing" Value="0" />
  <Logging Level="Standard" Path="%windir%\Temp" />
</Configuration>


try {
    $setupExe  = Join-Path $adtSession.DirFiles 'Setup.exe'
    $removeXml = Join-Path $adtSession.DirFiles 'Remove.xml'

    if (-not (Test-Path $setupExe)) {
        Write-ADTLogEntry -Message "Office Setup.exe not found in Files directory." -Severity 3 -Source $adtSession.InstallPhase
        throw "Missing ODT Setup.exe"
    }
    if (-not (Test-Path $removeXml)) {
        Write-ADTLogEntry -Message "Remove.xml not found; skipping Office uninstall." -Severity 2 -Source $adtSession.InstallPhase
        return
    }

    Write-ADTLogEntry -Message "Uninstalling Microsoft 365 Apps using $removeXml" -Severity 1 -Source $adtSession.InstallPhase
    Start-ADTProcess -FilePath $setupExe -ArgumentList "/configure `"$removeXml`"" -WorkingDirectory $adtSession.DirFiles
}
catch {
    Write-ADTLogEntry -Message "Office uninstall failed: $($_.Exception.Message)" -Severity 3 -Source $adtSession.InstallPhase
}
