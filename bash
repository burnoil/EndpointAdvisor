action uses wow64 redirection {not x64 of operating system}

// 1. Ensure the target folder exists
folder create "C:\Program Files\LLEA"
waithidden cmd.exe /c icacls "C:\Program Files\LLEA" /grant "Users":(OI)(CI)F /t

// 2. Build a compact PowerShell downloader (BigFix-safe, NO curly braces)
delete __createfile
createfile until END_OF_DOWNLOAD_SCRIPT
# LLEA Download Script - Compact BigFix-safe version
# Uses parentheses instead of curly braces to avoid BigFix parsing issues
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12, [Net.SecurityProtocolType]::Tls13
$ProgressPreference = 'SilentlyContinue'

# Define files to download
$urls = @(
    @("https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LLEA.ps1", "C:\Program Files\LLEA\LLEA.ps1"),
    @("https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/DriverUpdate.ps1", "C:\Program Files\LLEA\DriverUpdate.ps1"),
    @("https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LL_LOGO.ico", "C:\Program Files\LLEA\LL_LOGO.ico"),
    @("https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LL_LOGO_MSG.ico", "C:\Program Files\LLEA\LL_LOGO_MSG.ico")
)

$failed = @()

foreach ($item in $urls)
(
    $url = $item[0]
    $dest = $item[1]
    $success = $false
    
    for ($i = 1; $i -le 3; $i++)
    (
        try
        (
            Write-Host "Downloading $(Split-Path $dest -Leaf) (attempt $i)..."
            Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing -TimeoutSec 120 -ErrorAction Stop
            
            if (Test-Path $dest)
            (
                $fileSize = (Get-Item $dest).Length
                Write-Host "  Success: $fileSize bytes"
                $success = $true
                break
            )
        )
        catch
        (
            Write-Host "  Failed: $($_.Exception.Message)"
            if ($i -lt 3)
            (
                Start-Sleep -Seconds (5 * $i)
            )
        )
    )
    
    if (-not $success)
    (
        $failed += Split-Path $dest -Leaf
    )
)

if ($failed.Count -gt 0)
(
    Write-Host "ERROR: Failed to download: $($failed -join ', ')" -ForegroundColor Red
    exit 1
)

Write-Host "All files downloaded successfully!" -ForegroundColor Green
exit 0
END_OF_DOWNLOAD_SCRIPT

// 3. Drop the PowerShell script into place
copy __createfile "C:\Program Files\LLEA\download_LLEA.ps1"

// 4. Run the download script
override wait
hidden=true
wait powershell.exe -ExecutionPolicy Bypass -NoProfile -File "C:\Program Files\LLEA\download_LLEA.ps1"

// 4a. Check if download was successful before proceeding
continue if {exists file "LLEA.ps1" of folder "LLEA" of folder "Program Files" of drive of system folder}

// 5. Clean up the download script
delete "C:\Program Files\LLEA\download_LLEA.ps1"

// 5a. Ensure log directory exists
folder create "C:\Windows\MITLL\Logs"

// 5b. Create scheduled task using PowerShell and grant user permissions
delete __createfile
createfile until END_OF_TASK_CREATION
$taskName = "MITLL_DriverUpdate"
$taskDescription = "Monthly Windows driver updates via Windows Update"
$scriptPath = "C:\Program Files\LLEA\DriverUpdate.ps1"
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -File `"$scriptPath`""
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 2)
Register-ScheduledTask -TaskName $taskName -Description $taskDescription -Action $action -Principal $principal -Settings $settings -Force
# Grant authenticated users permission to start the task
$taskPath = "\$taskName"
$sid = New-Object System.Security.Principal.SecurityIdentifier("S-1-5-11")  # Authenticated Users
$taskScheduler = New-Object -ComObject Schedule.Service
$taskScheduler.Connect()
$rootFolder = $taskScheduler.GetFolder("\")
$task = $rootFolder.GetTask($taskName)
$securityDescriptor = $task.GetSecurityDescriptor(0xF)
$securityDescriptor += "(A;;GRGX;;;AU)"  # Grant Read and Execute to Authenticated Users
$task.SetSecurityDescriptor($securityDescriptor, 0)
Write-Output "Scheduled task created and permissions granted."
END_OF_TASK_CREATION

move __createfile "C:\Program Files\LLEA\CreateTask.ps1"

override wait
hidden=true
wait powershell.exe -ExecutionPolicy Bypass -File "C:\Program Files\LLEA\CreateTask.ps1"

delete "C:\Program Files\LLEA\CreateTask.ps1"

// 6. Register per machine Run key (hidden PowerShell)
override wait
hidden=true
wait reg add "HKLM\Software\Microsoft\Windows\CurrentVersion\Run" /v "LLEA" /t REG_SZ /d "\"C:\Windows\System32\conhost.exe\" --headless \"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe\" -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File \"C:\Program Files\LLEA\LLEA.ps1\" -RunMode LLEA\"" /f

// 6a. Brief pause to ensure cleanup complete
wait {pathname of system folder}\timeout.exe 1 /nobreak

// 7. Immediately invoke the (signed) script once
override wait
hidden=true
runas=currentuser  
wait cmd.exe /C start "" /b powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\Program Files\LLEA\LLEA.ps1" -RunMode LLEA
