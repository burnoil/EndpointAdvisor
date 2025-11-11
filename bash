# ================================================================
# YOUR CORRECTED UPGRADE SCRIPT
# ================================================================

action uses wow64 redirection {not x64 of operating system}

// 1) Kill any running LLEA processes using script file (avoids escaping issues)
delete __createfile
createfile until END_OF_KILL_SCRIPT
Get-Process -Name 'powershell','pwsh' -ErrorAction SilentlyContinue | Where-Object { 
    try { 
        `$proc = Get-WmiObject Win32_Process -Filter "ProcessId = `$(`$_.Id)" -ErrorAction SilentlyContinue
        `$proc -and `$proc.CommandLine -like '*LLEA.ps1*'
    } catch { 
        `$false 
    } 
} | Stop-Process -Force -ErrorAction SilentlyContinue
END_OF_KILL_SCRIPT

move __createfile "{pathname of system folder}\KillLLEA.ps1"

override wait
hidden=true
wait powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{pathname of system folder}\KillLLEA.ps1"

delete "{pathname of system folder}\KillLLEA.ps1"

// 1a) Wait for processes to fully terminate
wait {pathname of system folder}\timeout.exe 3 /nobreak

// 1b) Clean up lock file
delete "{(value "TEMP" of environment)}\LLEA_Instance.lock"

// 2) Remove old LLEA.ps1
delete "C:\Program Files\LLEA\LLEA.ps1"

// 2a) OPTIONAL: Comment out to preserve user settings
// delete "C:\Program Files\LLEA\LLEndpointAdvisor.config.json"

// 3) Download new LLEA script and icons
delete __createfile
createfile until END_OF_BATCH
@echo off
REM — download the signed script and icons
certutil -urlcache -f https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LLEA.ps1 "C:\Program Files\LLEA\LLEA.ps1"
if %ERRORLEVEL% NEQ 0 exit /b %ERRORLEVEL%
certutil -urlcache -f https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LL_LOGO.ico "C:\Program Files\LLEA\LL_LOGO.ico"
if %ERRORLEVEL% NEQ 0 exit /b %ERRORLEVEL%
certutil -urlcache -f https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LL_LOGO_MSG.ico "C:\Program Files\LLEA\LL_LOGO_MSG.ico"
exit /b 0
END_OF_BATCH

copy __createfile "C:\Program Files\LLEA\download_LLEA.bat"

override wait
hidden=true
wait cmd.exe /C "C:\Program Files\LLEA\download_LLEA.bat"

delete "C:\Program Files\LLEA\download_LLEA.bat"

// 6a) Brief pause to ensure filesystem operations complete
wait {pathname of system folder}\timeout.exe 1 /nobreak

// 7) Immediately invoke the (signed) script once
override wait
hidden=true
runas=currentuser  
wait cmd.exe /C start "" /b powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\Program Files\LLEA\LLEA.ps1" -RunMode LLEA

# ================================================================
# YOUR CORRECTED DEPLOYMENT SCRIPT
# ================================================================

action uses wow64 redirection {not x64 of operating system}

// 0) Kill any running LLEA processes using script file
delete __createfile
createfile until END_OF_KILL_SCRIPT
Get-Process -Name 'powershell','pwsh' -ErrorAction SilentlyContinue | Where-Object { 
    try { 
        `$proc = Get-WmiObject Win32_Process -Filter "ProcessId = `$(`$_.Id)" -ErrorAction SilentlyContinue
        `$proc -and `$proc.CommandLine -like '*LLEA.ps1*'
    } catch { 
        `$false 
    } 
} | Stop-Process -Force -ErrorAction SilentlyContinue
END_OF_KILL_SCRIPT

move __createfile "{pathname of system folder}\KillLLEA.ps1"

override wait
hidden=true
wait powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{pathname of system folder}\KillLLEA.ps1"

delete "{pathname of system folder}\KillLLEA.ps1"

// Wait for full termination
wait {pathname of system folder}\timeout.exe 2 /nobreak

// Clean up lock file
delete "{(value "TEMP" of environment)}\LLEA_Instance.lock"

// 1. Ensure the target folder exists
folder create "C:\Program Files\LLEA"
waithidden cmd.exe /c icacls "C:\Program Files\LLEA" /grant "Users":(OI)(CI)F /t

// 2. Build a pure-batch downloader using certutil
delete __createfile
createfile until END_OF_BATCH
@echo off
REM — download the signed scripts and icons
certutil -urlcache -f https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LLEA.ps1 "C:\Program Files\LLEA\LLEA.ps1"
if %ERRORLEVEL% NEQ 0 exit /b %ERRORLEVEL%
certutil -urlcache -f https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/DriverUpdate.ps1 "C:\Program Files\LLEA\DriverUpdate.ps1"
if %ERRORLEVEL% NEQ 0 exit /b %ERRORLEVEL%
certutil -urlcache -f https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LL_LOGO.ico "C:\Program Files\LLEA\LL_LOGO.ico"
if %ERRORLEVEL% NEQ 0 exit /b %ERRORLEVEL%
certutil -urlcache -f https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LL_LOGO_MSG.ico "C:\Program Files\LLEA\LL_LOGO_MSG.ico"
exit /b 0
END_OF_BATCH

// 3. Drop the batch into place
copy __createfile "C:\Program Files\LLEA\download_LLEA.bat"

// 4. Run it as the current user, hidden
override wait
hidden=true
wait cmd.exe /C "C:\Program Files\LLEA\download_LLEA.bat"

// 5. Clean up the batch once successful
delete "C:\Program Files\LLEA\download_LLEA.bat"

// 5a. Ensure log directory exists
folder create "C:\Windows\MITLL\Logs"

// 5b. Create scheduled task using PowerShell and grant user permissions
delete __createfile
createfile until END_OF_TASK_CREATION
`$taskName = "MITLL_DriverUpdate"
`$taskDescription = "Monthly Windows driver updates via Windows Update"
`$scriptPath = "C:\Program Files\LLEA\DriverUpdate.ps1"
`$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -File `"`$scriptPath`""
`$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
`$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 2)
Register-ScheduledTask -TaskName `$taskName -Description `$taskDescription -Action `$action -Principal `$principal -Settings `$settings -Force
# Grant authenticated users permission to start the task
`$taskPath = "\`$taskName"
`$sid = New-Object System.Security.Principal.SecurityIdentifier("S-1-5-11")  # Authenticated Users
`$taskScheduler = New-Object -ComObject Schedule.Service
`$taskScheduler.Connect()
`$rootFolder = `$taskScheduler.GetFolder("\")
`$task = `$rootFolder.GetTask(`$taskName)
`$securityDescriptor = `$task.GetSecurityDescriptor(0xF)
`$securityDescriptor += "(A;;GRGX;;;AU)"  # Grant Read and Execute to Authenticated Users
`$task.SetSecurityDescriptor(`$securityDescriptor, 0)
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

// 6a. Brief pause to ensure registry write completes
wait {pathname of system folder}\timeout.exe 1 /nobreak

// 7. Immediately invoke the (signed) script once
override wait
hidden=true
runas=currentuser  
wait cmd.exe /C start "" /b powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\Program Files\LLEA\LLEA.ps1" -RunMode LLEA
