action uses wow64 redirection {not x64 of operating system}

// 1. Ensure the target folder exists
folder create "C:\Program Files\LLEA"
waithidden cmd.exe /c icacls "C:\Program Files\LLEA" /grant "Users":(OI)(CI)F /t

// 2. Create download script with retry logic
delete __createfile
createfile until ___END_DOWNLOAD___
[Net.ServicePointManager]::SecurityProtocol = 'Tls12,Tls13'
$files = @(
  'https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LLEA.ps1',
  'https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/DriverUpdate.ps1',
  'https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LL_LOGO.ico',
  'https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LL_LOGO_MSG.ico'
)
$failed = @()
foreach ($url in $files) {{
  $filename = Split-Path $url -Leaf
  $dest = "C:\Program Files\LLEA\$filename"
  $ok = $false
  for ($i = 1; $i -le 3; $i++) {{
    try {{
      Write-Host "Download $filename attempt $i"
      Invoke-WebRequest -Uri $url -OutFile $dest -UseDefaultCredentials -UseBasicParsing -ErrorAction Stop
      $ok = $true
      Write-Host "  Success: $((Get-Item $dest).Length) bytes"
      break
    } catch {{
      Write-Host "  Failed: $($_.Exception.Message)"
      if ($i -lt 3) {{ Start-Sleep -Seconds ($i * 2) }
    }
  }
  if (-not $ok) {{ $failed += $filename }
}
if ($failed.Count -gt 0) {{ 
  Write-Host "ERROR: Failed to download: $($failed -join ', ')"
  exit 1 
}
Write-Host "All downloads successful"
exit 0
___END_DOWNLOAD___

move __createfile "C:\Program Files\LLEA\download.ps1"

// 3. Run download script
override wait
hidden=true
wait powershell.exe -ExecutionPolicy Bypass -NoProfile -File "C:\Program Files\LLEA\download.ps1"

// 4. Verify downloads succeeded
continue if {exists file "LLEA.ps1" of folder "LLEA" of folder "Program Files" of drive of system folder}

// 5. Clean up download script
delete "C:\Program Files\LLEA\download.ps1"

// 6. Ensure log directory exists
folder create "C:\Windows\MITLL\Logs"

// 7. Create scheduled task
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

// 8. Register per machine Run key
override wait
hidden=true
wait reg add "HKLM\Software\Microsoft\Windows\CurrentVersion\Run" /v "LLEA" /t REG_SZ /d "\"C:\Windows\System32\conhost.exe\" --headless \"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe\" -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File \"C:\Program Files\LLEA\LLEA.ps1\" -RunMode LLEA\"" /f

// 9. Brief pause
wait {pathname of system folder}\timeout.exe 1 /nobreak

// 10. Immediately invoke script (with relevance guard)
continue if {exists logged on user}
override wait
hidden=true
runas=currentuser  
wait cmd.exe /C start "" /b powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\Program Files\LLEA\LLEA.ps1" -RunMode LLEA
