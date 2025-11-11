╔══════════════════════════════════════════════════════════════════════════════╗
║           WORKING SOLUTION FOR WINDOWS 11 24H2 (No WMIC)                    ║
╚══════════════════════════════════════════════════════════════════════════════╝

Since you're on Windows 11 24H2, WMIC is gone. Use this ForEach loop method.

┌──────────────────────────────────────────────────────────────────────────────┐
│ YOUR COMPLETE UPGRADE SCRIPT (Copy This Entire Thing)                        │
└──────────────────────────────────────────────────────────────────────────────┘

action uses wow64 redirection {not x64 of operating system}

// 1) Kill any running LLEA processes using ForEach loop
delete __createfile
createfile until END_OF_KILL_SCRIPT
`$processes = Get-Process -Name powershell,pwsh -ErrorAction SilentlyContinue
foreach (`$p in `$processes) {
    `$wmi = Get-WmiObject Win32_Process -Filter "ProcessId=`$(`$p.Id)" -ErrorAction SilentlyContinue
    if (`$wmi -and `$wmi.CommandLine -like '*LLEA.ps1*') {
        Stop-Process -Id `$p.Id -Force -ErrorAction SilentlyContinue
    }
}
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


┌──────────────────────────────────────────────────────────────────────────────┐
│ DEPLOYMENT SCRIPT - ADD THIS AS STEP 0                                       │
└──────────────────────────────────────────────────────────────────────────────┘

// 0) Kill any existing LLEA instances
delete __createfile
createfile until END_OF_KILL_SCRIPT
`$processes = Get-Process -Name powershell,pwsh -ErrorAction SilentlyContinue
foreach (`$p in `$processes) {
    `$wmi = Get-WmiObject Win32_Process -Filter "ProcessId=`$(`$p.Id)" -ErrorAction SilentlyContinue
    if (`$wmi -and `$wmi.CommandLine -like '*LLEA.ps1*') {
        Stop-Process -Id `$p.Id -Force -ErrorAction SilentlyContinue
    }
}
END_OF_KILL_SCRIPT

move __createfile "{pathname of system folder}\KillLLEA.ps1"

override wait
hidden=true
wait powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{pathname of system folder}\KillLLEA.ps1"

delete "{pathname of system folder}\KillLLEA.ps1"

wait {pathname of system folder}\timeout.exe 2 /nobreak
delete "{(value "TEMP" of environment)}\LLEA_Instance.lock"

// Then continue with your existing deployment steps...


┌──────────────────────────────────────────────────────────────────────────────┐
│ VERIFY THE SCRIPT FILE BIGFIX CREATES                                        │
└──────────────────────────────────────────────────────────────────────────────┘

To test if BigFix is creating the file correctly, use this test action:

delete __createfile
createfile until END_OF_TEST
`$processes = Get-Process -Name powershell,pwsh -ErrorAction SilentlyContinue
foreach (`$p in `$processes) {
    `$wmi = Get-WmiObject Win32_Process -Filter "ProcessId=`$(`$p.Id)" -ErrorAction SilentlyContinue
    if (`$wmi -and `$wmi.CommandLine -like '*LLEA.ps1*') {
        Write-Output "Found LLEA process: PID `$(`$p.Id)"
    }
}
END_OF_TEST

move __createfile "C:\Temp\TestKillLLEA.ps1"

// Don't run it yet - just create it
// Then manually check C:\Temp\TestKillLLEA.ps1 to see what BigFix created


The file should contain:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
$processes = Get-Process -Name powershell,pwsh -ErrorAction SilentlyContinue
foreach ($p in $processes) {
    $wmi = Get-WmiObject Win32_Process -Filter "ProcessId=$($p.Id)" -ErrorAction SilentlyContinue
    if ($wmi -and $wmi.CommandLine -like '*LLEA.ps1*') {
        Write-Output "Found LLEA process: PID $($p.Id)"
    }
}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

If it looks correct, then run the actual kill script.


┌──────────────────────────────────────────────────────────────────────────────┐
│ ALTERNATIVE: CIM Instead of WMI (Modern Method)                              │
└──────────────────────────────────────────────────────────────────────────────┘

If the WMI approach still has issues, try CIM cmdlets (newer, cleaner):

delete __createfile
createfile until END_OF_KILL_SCRIPT
`$processes = Get-Process -Name powershell,pwsh -ErrorAction SilentlyContinue
foreach (`$p in `$processes) {
    try {
        `$cim = Get-CimInstance Win32_Process -Filter "ProcessId=`$(`$p.Id)" -ErrorAction SilentlyContinue
        if (`$cim -and `$cim.CommandLine -like '*LLEA.ps1*') {
            Stop-Process -Id `$p.Id -Force -ErrorAction SilentlyContinue
        }
    } catch {
        # Ignore errors
    }
}
END_OF_KILL_SCRIPT

move __createfile "{pathname of system folder}\KillLLEA.ps1"

override wait
hidden=true
wait powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{pathname of system folder}\KillLLEA.ps1"

delete "{pathname of system folder}\KillLLEA.ps1"


┌──────────────────────────────────────────────────────────────────────────────┐
│ SIMPLEST POSSIBLE VERSION (Last Resort)                                      │
└──────────────────────────────────────────────────────────────────────────────┘

If all else fails, use the most basic approach:

delete __createfile
createfile until END_OF_KILL_SCRIPT
Get-Process | Where-Object {
    (`$_.Name -eq 'powershell' -or `$_.Name -eq 'pwsh') -and
    (`$_.MainWindowTitle -like '*Lincoln Laboratory Endpoint Advisor*')
} | Stop-Process -Force -ErrorAction SilentlyContinue

# Also try command line match
Get-Process | Where-Object {
    `$_.Name -eq 'powershell' -or `$_.Name -eq 'pwsh'
} | ForEach-Object {
    try {
        `$cmdline = (Get-CimInstance Win32_Process -Filter "ProcessId=`$(`$_.Id)").CommandLine
        if (`$cmdline -like '*LLEA.ps1*') {
            Stop-Process -Id `$_.Id -Force -ErrorAction SilentlyContinue
        }
    } catch {}
}
END_OF_KILL_SCRIPT

move __createfile "{pathname of system folder}\KillLLEA.ps1"

override wait
hidden=true
wait powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{pathname of system folder}\KillLLEA.ps1"

delete "{pathname of system folder}\KillLLEA.ps1"


╔══════════════════════════════════════════════════════════════════════════════╗
║  BACKTICK ESCAPING GUIDE FOR BIGFIX                                          ║
╚══════════════════════════════════════════════════════════════════════════════╝

In BigFix createfile blocks:

  `$variable      →  $variable       (single backtick)
  `$(`$p.Id)      →  $($p.Id)        (backtick before each $)
  ``             →  `                (double backtick for literal backtick)

The subexpression operator $() requires careful escaping:
  - `$ for the opening $
  - `$ for any $ inside the ()


╔══════════════════════════════════════════════════════════════════════════════╗
║  TESTING PROCEDURE                                                           ║
╚══════════════════════════════════════════════════════════════════════════════╝

1. Create test fixlet with file creation only (don't run yet)
2. Deploy to test machine
3. Check the created .ps1 file manually
4. If it looks correct, run it manually from PowerShell
5. If it works manually, let BigFix run it
6. If it works, deploy to production


╔══════════════════════════════════════════════════════════════════════════════╗
║  RECOMMENDED: Use the ForEach Loop Method (First Option Above)              ║
╚══════════════════════════════════════════════════════════════════════════════╝

This is the most reliable for Windows 11 24H2.
The foreach loop avoids the complex Where-Object pipeline that causes escaping issues.
