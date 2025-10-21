# Pre-Install
$adtSession.InstallPhase = 'Pre-Install'
Write-ADTLogEntry -Message 'Starting Pre-Install' -Source $adtSession.InstallPhase

# Install
$adtSession.InstallPhase = 'Install'
Write-ADTLogEntry -Message 'Starting Install' -Source $adtSession.InstallPhase

# Post-Install
$adtSession.InstallPhase = 'Post-Install'
Write-ADTLogEntry -Message 'Starting Post-Install' -Source $adtSession.InstallPhase

# Uninstall
$adtSession.InstallPhase = 'Uninstall'
Write-ADTLogEntry -Message 'Starting Uninstall' -Source $adtSession.InstallPhase


$adtSession.InstallPhase = 'Uninstall'
Write-ADTLogEntry -Message 'Removing Microsoft 365 Apps via Uninstall.xml' -Source $adtSession.InstallPhase

Start-ADTProcess -FilePath "$dirFiles\Setup.exe" `
                 -ArgumentList '/configure Uninstall.xml' `
                 -WorkingDirectory $dirFiles `
                 -WindowStyle Hidden -CreateNoWindow:$true
