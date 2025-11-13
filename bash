action uses wow64 redirection {not x64 of operating system}

// 1. Ensure the target folder exists
folder create "C:\Program Files\LLEA"
waithidden cmd.exe /c icacls "C:\Program Files\LLEA" /grant "Users":(OI)(CI)F /t

// 2. Create ultra-simple PowerShell downloader using .NET WebClient
// Avoids complex PowerShell structures and curly braces as much as possible
delete __createfile
createfile until ___END_DOWNLOAD_PS1___
[Net.ServicePointManager]::SecurityProtocol = 'Tls12,Tls13'
$wc = New-Object System.Net.WebClient
$files = @(
  @('https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LLEA.ps1','C:\Program Files\LLEA\LLEA.ps1'),
  @('https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/DriverUpdate.ps1','C:\Program Files\LLEA\DriverUpdate.ps1'),
  @('https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LL_LOGO.ico','C:\Program Files\LLEA\LL_LOGO.ico'),
  @('https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LL_LOGO_MSG.ico','C:\Program Files\LLEA\LL_LOGO_MSG.ico')
)
$failed = @()
foreach ($f in $files) {{
  $url,$dest = $f
  $ok = $false
  1..3 | ForEach-Object {{
    if (-not $ok) {{
      try {{
        Write-Host "Downloading $(Split-Path $dest -Leaf) attempt $_"
        $wc.DownloadFile($url,$dest)
        if (Test-Path $dest) {{ $ok = $true; Write-Host "  OK: $((Get-Item $dest).Length) bytes" }}
      }} catch {{
        Write-Host "  Failed: $($_.Exception.Message)"
        Start-Sleep -Seconds ($_ * 2)
      }}
    }}
  }}
  if (-not $ok) {{ $failed += Split-Path $dest -Leaf }}
}}
$wc.Dispose()
if ($failed.Count -gt 0) {{ Write-Host "ERROR: $($failed -join ',')"; exit 1 }}
Write-Host "Success!"; exit 0
___END_DOWNLOAD_PS1___

// 3. Drop the PowerShell script into place
move __createfile "C:\Program Files\LLEA\download_LLEA.ps1"

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
