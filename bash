action uses wow64 redirection {not x64 of operating system}

// 1. Ensure the target folder exists
folder create "C:\Program Files\LLEA"
waithidden cmd.exe /c icacls "C:\Program Files\LLEA" /grant "Users":(OI)(CI)F /t

// 2. Download files using PowerShell one-liners with retry
// This avoids ALL curly brace issues and is more reliable than bitsadmin

// Download LLEA.ps1 (with 3 retry attempts)
waithidden powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$p='SilentlyContinue';$ProgressPreference=$p;[Net.ServicePointManager]::SecurityProtocol='Tls12,Tls13';$r=0;do{try{Invoke-WebRequest -Uri 'https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LLEA.ps1' -OutFile 'C:\Program Files\LLEA\LLEA.ps1' -UseBasicParsing -TimeoutSec 120 -ErrorAction Stop;$r=99}catch{$r++;Start-Sleep -Seconds ($r*5)}}while($r -lt 3 -and $r -ne 99)"

// Brief pause between downloads
wait {pathname of system folder}\timeout.exe 2 /nobreak

// Download DriverUpdate.ps1 (with 3 retry attempts)
waithidden powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$p='SilentlyContinue';$ProgressPreference=$p;[Net.ServicePointManager]::SecurityProtocol='Tls12,Tls13';$r=0;do{try{Invoke-WebRequest -Uri 'https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/DriverUpdate.ps1' -OutFile 'C:\Program Files\LLEA\DriverUpdate.ps1' -UseBasicParsing -TimeoutSec 120 -ErrorAction Stop;$r=99}catch{$r++;Start-Sleep -Seconds ($r*5)}}while($r -lt 3 -and $r -ne 99)"

// Brief pause between downloads
wait {pathname of system folder}\timeout.exe 2 /nobreak

// Download LL_LOGO.ico (with 3 retry attempts)
waithidden powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$p='SilentlyContinue';$ProgressPreference=$p;[Net.ServicePointManager]::SecurityProtocol='Tls12,Tls13';$r=0;do{try{Invoke-WebRequest -Uri 'https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LL_LOGO.ico' -OutFile 'C:\Program Files\LLEA\LL_LOGO.ico' -UseBasicParsing -TimeoutSec 120 -ErrorAction Stop;$r=99}catch{$r++;Start-Sleep -Seconds ($r*5)}}while($r -lt 3 -and $r -ne 99)"

// Brief pause between downloads
wait {pathname of system folder}\timeout.exe 2 /nobreak

// Download LL_LOGO_MSG.ico (with 3 retry attempts)
waithidden powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$p='SilentlyContinue';$ProgressPreference=$p;[Net.ServicePointManager]::SecurityProtocol='Tls12,Tls13';$r=0;do{try{Invoke-WebRequest -Uri 'https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LL_LOGO_MSG.ico' -OutFile 'C:\Program Files\LLEA\LL_LOGO_MSG.ico' -UseBasicParsing -TimeoutSec 120 -ErrorAction Stop;$r=99}catch{$r++;Start-Sleep -Seconds ($r*5)}}while($r -lt 3 -and $r -ne 99)"

// 3. Verify downloads succeeded
continue if {exists file "LLEA.ps1" of folder "LLEA" of folder "Program Files" of drive of system folder}
continue if {exists file "DriverUpdate.ps1" of folder "LLEA" of folder "Program Files" of drive of system folder}
continue if {exists file "LL_LOGO.ico" of folder "LLEA" of folder "Program Files" of drive of system folder}
continue if {exists file "LL_LOGO_MSG.ico" of folder "LLEA" of folder "Program Files" of drive of system folder}

// 4. Ensure log directory exists
folder create "C:\Windows\MITLL\Logs"

// 5. Create scheduled task using PowerShell and grant user permissions
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

// 7. Brief pause to ensure cleanup complete
wait {pathname of system folder}\timeout.exe 1 /nobreak

// 8. Immediately invoke the (signed) script once
override wait
hidden=true
runas=currentuser  
wait cmd.exe /C start "" /b powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\Program Files\LLEA\LLEA.ps1" -RunMode LLEA
