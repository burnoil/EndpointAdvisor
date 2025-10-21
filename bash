// Prefetch the PSADT package (adjust URL/hash as needed)
prefetch M365_PSADT.zip sha1:1234567890abcdef1234567890abcdef12345678 size:12345678 http://your-relay-server/M365_PSADT.zip sha256:abcdef...

// Extract the package
extract M365_PSADT.zip

// Create PowerShell script to launch via Task Scheduler
delete __createtask.ps1

createfile until _END_OF_FILE_
# Get logged-on user
$loggedOnUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName

if ([string]::IsNullOrEmpty($loggedOnUser)) {{
    Write-Host "No user logged on"
    exit 1
}}

# Define the PSADT path from BigFix download folder
$psadtPath = "{escapeForPowerShell of pathname of client folder of current site}\__Download\Deploy-Application.ps1"

# Verify the file exists
if (-not (Test-Path $psadtPath)) {{
    Write-Error "PSADT Deploy-Application.ps1 not found at $psadtPath"
    exit 1
}}

# Create scheduled task
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File `"$psadtPath`""
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddSeconds(5)
$principal = New-ScheduledTaskPrincipal -UserId $loggedOnUser -LogonType Interactive -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

# Register and start task
Register-ScheduledTask -TaskName "M365_PSADT_Install" -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
Start-ScheduledTask -TaskName "M365_PSADT_Install"

# Wait for task to start then cleanup
Start-Sleep -Seconds 10
Unregister-ScheduledTask -TaskName "M365_PSADT_Install" -Confirm:$false
_END_OF_FILE_

move __createtask.ps1 __Download\createtask.ps1

// Execute the PowerShell script
waithidden powershell.exe -ExecutionPolicy Bypass -NoProfile -File "{pathname of client folder of current site}\__Download\createtask.ps1"
