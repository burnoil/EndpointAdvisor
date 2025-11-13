action uses wow64 redirection {not x64 of operating system}

// 1. Ensure the target folder exists
folder create "C:\Program Files\LLEA"
waithidden cmd.exe /c icacls "C:\Program Files\LLEA" /grant "Users":(OI)(CI)F /t

// 2. Download with 3 retry attempts per file
override wait
hidden=true
wait powershell.exe -ExecutionPolicy Bypass -NoProfile -Command "& { [Net.ServicePointManager]::SecurityProtocol = 'Tls12,Tls13'; $ok = $false; 1..3 | ForEach-Object { if (-not $ok) { try { Invoke-WebRequest -Uri 'https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LLEA.ps1' -OutFile 'C:\Program Files\LLEA\LLEA.ps1' -UseDefaultCredentials -UseBasicParsing; $ok = $true } catch { Start-Sleep -Seconds ($_ * 2) } } }; if (-not $ok) { exit 1 } }"

override wait
hidden=true
wait powershell.exe -ExecutionPolicy Bypass -NoProfile -Command "& { [Net.ServicePointManager]::SecurityProtocol = 'Tls12,Tls13'; $ok = $false; 1..3 | ForEach-Object { if (-not $ok) { try { Invoke-WebRequest -Uri 'https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/DriverUpdate.ps1' -OutFile 'C:\Program Files\LLEA\DriverUpdate.ps1' -UseDefaultCredentials -UseBasicParsing; $ok = $true } catch { Start-Sleep -Seconds ($_ * 2) } } }; if (-not $ok) { exit 1 } }"

override wait
hidden=true
wait powershell.exe -ExecutionPolicy Bypass -NoProfile -Command "& { [Net.ServicePointManager]::SecurityProtocol = 'Tls12,Tls13'; $ok = $false; 1..3 | ForEach-Object { if (-not $ok) { try { Invoke-WebRequest -Uri 'https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LL_LOGO.ico' -OutFile 'C:\Program Files\LLEA\LL_LOGO.ico' -UseDefaultCredentials -UseBasicParsing; $ok = $true } catch { Start-Sleep -Seconds ($_ * 2) } } }; if (-not $ok) { exit 1 } }"

override wait
hidden=true
wait powershell.exe -ExecutionPolicy Bypass -NoProfile -Command "& { [Net.ServicePointManager]::SecurityProtocol = 'Tls12,Tls13'; $ok = $false; 1..3 | ForEach-Object { if (-not $ok) { try { Invoke-WebRequest -Uri 'https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LL_LOGO_MSG.ico' -OutFile 'C:\Program Files\LLEA\LL_LOGO_MSG.ico' -UseDefaultCredentials -UseBasicParsing; $ok = $true } catch { Start-Sleep -Seconds ($_ * 2) } } }; if (-not $ok) { exit 1 } }"

// 3. Check if download was successful before proceeding
continue if {exists file "LLEA.ps1" of folder "LLEA" of folder "Program Files" of drive of system folder}

// 4. Ensure log directory exists
folder create "C:\Windows\MITLL\Logs"

// 5. Create scheduled task
delete __createfile
createfile until END_OF_TASK_CREATION
$taskName = "MITLL_DriverUpdate"
$taskDescription = "Monthly Windows driver updates via Windows Update"
$scriptPath = "C:\Program Files\LLEA\DriverUpdate.ps1"
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -File `"$scriptPath`""
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 2)
$task = Register-ScheduledTask -TaskName $taskName -Description $taskDescription -Action $action -Principal $principal -Settings $settings -Force
$sddl = $task.GetSecurityDescriptor(0xF)
$newSDDL = $sddl + "(A;;GRGX;;;AU)"
$task.SetSecurityDescriptor($newSDDL, 0)
Write-Output "Scheduled task created and permissions granted."
END_OF_TASK_CREATION

move __createfile "C:\Program Files\LLEA\CreateTask.ps1"

override wait
hidden=true
wait powershell.exe -ExecutionPolicy Bypass -File "C:\Program Files\LLEA\CreateTask.ps1"

delete "C:\Program Files\LLEA\CreateTask.ps1"

// 6. Register per machine Run key
override wait
hidden=true
wait reg add "HKLM\Software\Microsoft\Windows\CurrentVersion\Run" /v "LLEA" /t REG_SZ /d "\"C:\Windows\System32\conhost.exe\" --headless \"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe\" -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File \"C:\Program Files\LLEA\LLEA.ps1\" -RunMode LLEA\"" /f

// 7. Brief pause
wait {pathname of system folder}\timeout.exe 1 /nobreak

// 8. Immediately invoke script (with relevance guard)
continue if {exists logged on user}
override wait
hidden=true
runas=currentuser  
wait cmd.exe /C start "" /b powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\Program Files\LLEA\LLEA.ps1" -RunMode LLEA
