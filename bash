// ============================================================
// Lincoln Laboratory Endpoint Advisor - BigFix Backup Deployment
// ============================================================
// This action script deploys LLEA from a prefetched archive
// ============================================================

// Step 1: Stop any running instances of LLEA
// ============================================================
waithidden powershell.exe -ExecutionPolicy Bypass -NoProfile -Command "Get-Process -Name 'powershell','pwsh' -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -like '*Lincoln Laboratory Endpoint Advisor*' } | Stop-Process -Force"

// Wait for processes to terminate
pause {now + 2 * second}

// Step 2: Create the installation directory
// ============================================================
folder create "C:\Program Files\LLEA"

// Step 3: Extract LLEA files from the archive
// ============================================================
// NOTE: The LLEA.zip file should be prefetched in the BigFix action
// The archive should contain:
//   - LLEA.ps1
//   - LL_LOGO.ico
//   - LL_LOGO_MSG.ico
//   - ContentData.json (optional, will be cached from GitHub)

extract LLEA.zip "C:\Program Files\LLEA"

// Step 4: Set proper permissions on the LLEA directory
// ============================================================
waithidden icacls "C:\Program Files\LLEA" /grant "BUILTIN\Users:(OI)(CI)RX" /T /C
waithidden icacls "C:\Program Files\LLEA" /grant "BUILTIN\Administrators:(OI)(CI)F" /T /C

// Step 5: Remove existing LLEA scheduled tasks (if any)
// ============================================================
waithidden schtasks.exe /Delete /TN "LLEA-UserLogon" /F
waithidden schtasks.exe /Delete /TN "Lincoln Laboratory Endpoint Advisor" /F

// Step 6: Create new scheduled task for LLEA (runs at user logon)
// ============================================================
delete __createfile
createfile until __END_OF_TASK_XML__
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>2025-01-01T00:00:00</Date>
    <Author>ISD Endpoint Engineering</Author>
    <Description>Lincoln Laboratory Endpoint Advisor - Provides system notifications and alerts via system tray application</Description>
    <URI>\LLEA-UserLogon</URI>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
      <Delay>PT30S</Delay>
    </LogonTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>S-1-5-32-545</UserId>
      <GroupId>S-1-5-32-545</GroupId>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>false</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
    <UseUnifiedSchedulingEngine>true</UseUnifiedSchedulingEngine>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>powershell.exe</Command>
      <Arguments>-ExecutionPolicy Bypass -WindowStyle Hidden -NoProfile -File "C:\Program Files\LLEA\LLEA.ps1"</Arguments>
    </Exec>
  </Actions>
</Task>
__END_OF_TASK_XML__

// Move the task XML to a temporary location
delete "C:\Windows\Temp\LLEA-Task.xml"
move __createfile "C:\Windows\Temp\LLEA-Task.xml"

// Import the scheduled task
waithidden schtasks.exe /Create /XML "C:\Windows\Temp\LLEA-Task.xml" /TN "LLEA-UserLogon" /F

// Clean up temporary task XML file
delete "C:\Windows\Temp\LLEA-Task.xml"

// Step 7: Set PowerShell execution policy (if needed)
// ============================================================
// This ensures PowerShell scripts can run on the machine
waithidden powershell.exe -ExecutionPolicy Bypass -NoProfile -Command "Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine -Force -ErrorAction SilentlyContinue"

// Step 8: Create initial configuration file
// ============================================================
delete __createfile
createfile until __END_OF_INIT_SCRIPT__
# Initialize LLEA configuration
$configPath = "C:\Program Files\LLEA\LLEndpointAdvisor.config.json"
$defaultConfig = @{
    RefreshInterval = 900
    LogRotationSizeMB = 2
    DefaultLogLevel = "INFO"
    ContentDataUrl = "https://raw.githubusercontent.com/burnoil/EndpointAdvisor/refs/heads/main/ContentData.json"
    CertificateCheckInterval = 86400
    YubiKeyAlertDays = 14
    IconPaths = @{
        Main = "C:\Program Files\LLEA\LL_LOGO.ico"
        Warning = "C:\Program Files\LLEA\LL_LOGO_MSG.ico"
    }
    AnnouncementsLastState = "{}"
    SupportLastState = "{}"
    LastSeenUpdateState = ""
    BigFixSSA_Path = "C:\Program Files (x86)\BigFix Enterprise\BigFix Self Service Application\BigFixSSA.exe"
    YubiKeyManager_Path = "C:\Program Files\Yubico\Yubikey Manager\ykman.exe"
    BlinkingEnabled = $false
    CachePath = "C:\Program Files\LLEA\ContentData.cache.json"
    HasRunBefore = $false
}

# Only create config if it doesn't exist (preserve existing configurations)
if (-not (Test-Path $configPath)) {
    $defaultConfig | ConvertTo-Json -Depth 10 | Out-File $configPath -Force
    Write-Host "LLEA configuration initialized at $configPath"
} else {
    Write-Host "Existing LLEA configuration preserved at $configPath"
}
__END_OF_INIT_SCRIPT__

delete "C:\Windows\Temp\LLEA-Init.ps1"
move __createfile "C:\Windows\Temp\LLEA-Init.ps1"

waithidden powershell.exe -ExecutionPolicy Bypass -NoProfile -File "C:\Windows\Temp\LLEA-Init.ps1"

delete "C:\Windows\Temp\LLEA-Init.ps1"

// Step 9: Start LLEA for the current user session (if user is logged in)
// ============================================================
action requires restart "false"

// Attempt to start LLEA in the current user context
waithidden powershell.exe -ExecutionPolicy Bypass -NoProfile -Command "Start-Process -FilePath 'powershell.exe' -ArgumentList '-ExecutionPolicy Bypass -WindowStyle Hidden -NoProfile -File \"C:\Program Files\LLEA\LLEA.ps1\"' -ErrorAction SilentlyContinue"

// ============================================================
// Deployment Complete
// ============================================================
continue if {exists file "C:\Program Files\LLEA\LLEA.ps1"}
continue if {exists file "C:\Program Files\LLEA\LL_LOGO.ico"}
continue if {exists file "C:\Program Files\LLEA\LL_LOGO_MSG.ico"}
