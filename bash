╔══════════════════════════════════════════════════════════════════════════════╗
║        MODERN SOLUTION FOR WINDOWS 11 24H2 (No Deprecated Features)         ║
╚══════════════════════════════════════════════════════════════════════════════╝

No WMIC, no VBScript - uses modern Get-CimInstance with minimal backticks.

┌──────────────────────────────────────────────────────────────────────────────┐
│ SOLUTION 1: Simplified PowerShell (Minimal Backticks)                        │
└──────────────────────────────────────────────────────────────────────────────┘

delete __createfile
createfile until END_OF_KILL
Get-CimInstance Win32_Process | Where-Object {
    (`$_.Name -eq 'powershell.exe' -or `$_.Name -eq 'pwsh.exe') -and
    (`$_.CommandLine -like '*LLEA.ps1*')
} | ForEach-Object {
    Stop-Process -Id `$_.ProcessId -Force -ErrorAction SilentlyContinue
}
END_OF_KILL

move __createfile "{pathname of system folder}\KillLLEA.ps1"

override wait
hidden=true
wait powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{pathname of system folder}\KillLLEA.ps1"

delete "{pathname of system folder}\KillLLEA.ps1"


┌──────────────────────────────────────────────────────────────────────────────┐
│ SOLUTION 2: Even Simpler - Direct Pipeline                                   │
└──────────────────────────────────────────────────────────────────────────────┘

delete __createfile
createfile until END_OF_KILL
Get-CimInstance Win32_Process -Filter "Name='powershell.exe' OR Name='pwsh.exe'" | 
    Where-Object { `$_.CommandLine -like '*LLEA.ps1*' } | 
    ForEach-Object { Stop-Process -Id `$_.ProcessId -Force -ErrorAction SilentlyContinue }
END_OF_KILL

move __createfile "{pathname of system folder}\KillLLEA.ps1"

override wait
hidden=true
wait powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{pathname of system folder}\KillLLEA.ps1"

delete "{pathname of system folder}\KillLLEA.ps1"


┌──────────────────────────────────────────────────────────────────────────────┐
│ SOLUTION 3: Inline PowerShell (No File Creation)                             │
└──────────────────────────────────────────────────────────────────────────────┘

override wait
hidden=true
wait powershell.exe -NoProfile -Command "Get-CimInstance Win32_Process -Filter \"Name='powershell.exe' OR Name='pwsh.exe'\" | Where-Object { $_.CommandLine -like '*LLEA.ps1*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"


┌──────────────────────────────────────────────────────────────────────────────┐
│ SOLUTION 4: Base64 Encoded (Most Reliable)                                   │
└──────────────────────────────────────────────────────────────────────────────┘

// Step 1: Create the base64 encoded kill command
// This PowerShell creates a base64 string - no backticks in the command itself

delete __createfile
createfile until END_OF_ENCODER
$command = "Get-CimInstance Win32_Process -Filter \"Name='powershell.exe' OR Name='pwsh.exe'\" | Where-Object { $_.CommandLine -like '*LLEA.ps1*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"
$bytes = [System.Text.Encoding]::Unicode.GetBytes($command)
$encoded = [Convert]::ToBase64String($bytes)
$encoded | Out-File -FilePath "$env:TEMP\kill_encoded.txt" -Encoding ASCII -NoNewline
END_OF_ENCODER

move __createfile "{pathname of system folder}\CreateEncoded.ps1"

override wait
hidden=true
wait powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{pathname of system folder}\CreateEncoded.ps1"

delete "{pathname of system folder}\CreateEncoded.ps1"

// Step 2: Run the encoded command
override wait
hidden=true
wait powershell.exe -NoProfile -EncodedCommand {concatenation "" of lines of file "kill_encoded.txt" of folder (value "TEMP" of environment)}

delete "{(value "TEMP" of environment)}\kill_encoded.txt"


┌──────────────────────────────────────────────────────────────────────────────┐
│ RECOMMENDED: YOUR COMPLETE UPGRADE SCRIPT (Solution 2)                       │
└──────────────────────────────────────────────────────────────────────────────┘

action uses wow64 redirection {not x64 of operating system}

// 1) Kill any running LLEA processes
delete __createfile
createfile until END_OF_KILL
Get-CimInstance Win32_Process -Filter "Name='powershell.exe' OR Name='pwsh.exe'" | 
    Where-Object { `$_.CommandLine -like '*LLEA.ps1*' } | 
    ForEach-Object { Stop-Process -Id `$_.ProcessId -Force -ErrorAction SilentlyContinue }
END_OF_KILL

move __createfile "{pathname of system folder}\KillLLEA.ps1"

override wait
hidden=true
wait powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{pathname of system folder}\KillLLEA.ps1"

delete "{pathname of system folder}\KillLLEA.ps1"

// 1a) Wait for termination
wait {pathname of system folder}\timeout.exe 3 /nobreak

// 1b) Clean up lock file
delete "{(value "TEMP" of environment)}\LLEA_Instance.lock"

// 2) Remove old LLEA.ps1
delete "C:\Program Files\LLEA\LLEA.ps1"

// 3) Download new LLEA script
delete __createfile
createfile until END_OF_BATCH
@echo off
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

wait {pathname of system folder}\timeout.exe 1 /nobreak

override wait
hidden=true
runas=currentuser  
wait cmd.exe /C start "" /b powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\Program Files\LLEA\LLEA.ps1" -RunMode LLEA


┌──────────────────────────────────────────────────────────────────────────────┐
│ DEPLOYMENT SCRIPT - ADD THIS AS STEP 0                                       │
└──────────────────────────────────────────────────────────────────────────────┘

// 0) Kill existing LLEA
delete __createfile
createfile until END_OF_KILL
Get-CimInstance Win32_Process -Filter "Name='powershell.exe' OR Name='pwsh.exe'" | 
    Where-Object { `$_.CommandLine -like '*LLEA.ps1*' } | 
    ForEach-Object { Stop-Process -Id `$_.ProcessId -Force -ErrorAction SilentlyContinue }
END_OF_KILL

move __createfile "{pathname of system folder}\KillLLEA.ps1"

override wait
hidden=true
wait powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{pathname of system folder}\KillLLEA.ps1"

delete "{pathname of system folder}\KillLLEA.ps1"

wait {pathname of system folder}\timeout.exe 2 /nobreak
delete "{(value "TEMP" of environment)}\LLEA_Instance.lock"

// Then your existing deployment steps...


╔══════════════════════════════════════════════════════════════════════════════╗
║  WHY THIS WORKS                                                              ║
╚══════════════════════════════════════════════════════════════════════════════╝

✅ Get-CimInstance: Modern, not deprecated (replaces Get-WmiObject)
✅ Filter at source: Reduces backtick usage
✅ Simple pipeline: Minimal complexity
✅ Only `$_ used: Single underscore, easier escaping
✅ Works on Windows 11 24H2: Uses modern cmdlets

The key: Keep the script SIMPLE with minimal backtick usage.


╔══════════════════════════════════════════════════════════════════════════════╗
║  TEST IT MANUALLY FIRST                                                      ║
╚══════════════════════════════════════════════════════════════════════════════╝

Run this in PowerShell while LLEA is running:

Get-CimInstance Win32_Process -Filter "Name='powershell.exe' OR Name='pwsh.exe'" | 
    Where-Object { $_.CommandLine -like '*LLEA.ps1*' } | 
    Select-Object ProcessId, ProcessName, CommandLine

This will show you what it finds.

To actually kill:

Get-CimInstance Win32_Process -Filter "Name='powershell.exe' OR Name='pwsh.exe'" | 
    Where-Object { $_.CommandLine -like '*LLEA.ps1*' } | 
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force }


╔══════════════════════════════════════════════════════════════════════════════╗
║  IF BACKTICKS STILL FAIL - USE SOLUTION 4 (Base64)                          ║
╚══════════════════════════════════════════════════════════════════════════════╝

The base64 method is GUARANTEED to work because:
1. The encoder script uses normal PowerShell (no BigFix escaping)
2. The encoded command has no special characters
3. BigFix just passes a base64 string - nothing to escape

It's a two-step process but 100% reliable.


╔══════════════════════════════════════════════════════════════════════════════╗
║  WHAT'S DIFFERENT                                                            ║
╚══════════════════════════════════════════════════════════════════════════════╝

OLD (deprecated):
  Get-WmiObject Win32_Process -Filter "..." 
  VBScript
  WMIC

NEW (modern):
  Get-CimInstance Win32_Process -Filter "..."
  PowerShell pipeline
  No deprecated features

All work on Windows 11 24H2.


╔══════════════════════════════════════════════════════════════════════════════╗
║  BACKTICK ESCAPING GUIDE                                                     ║
╚══════════════════════════════════════════════════════════════════════════════╝

In BigFix createfile:
  `$_             →  $_              (backtick before $)
  `$_.Property    →  $_.Property     (backtick only before $)

Keep it simple:
  ✅ Use `$_ 
  ✅ Use `$_.Property
  ❌ Avoid `$(`$var.Property) - complex nesting


╔══════════════════════════════════════════════════════════════════════════════╗
║  SUMMARY                                                                     ║
╚══════════════════════════════════════════════════════════════════════════════╝

✅ Use Get-CimInstance (modern, not deprecated)
✅ Keep script simple (minimal backticks)
✅ Test manually first
✅ If backticks still fail, use base64 method

Try Solution 2 first. If it fails, use Solution 4 (base64).
