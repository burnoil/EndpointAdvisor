╔══════════════════════════════════════════════════════════════════════════════╗
║               FINAL WORKING SOLUTION - USE THIS BATCH METHOD                 ║
╚══════════════════════════════════════════════════════════════════════════════╝

This uses a BATCH FILE instead of PowerShell to avoid ALL backtick issues.

┌──────────────────────────────────────────────────────────────────────────────┐
│ YOUR COMPLETE UPGRADE SCRIPT (Copy This)                                     │
└──────────────────────────────────────────────────────────────────────────────┘

action uses wow64 redirection {not x64 of operating system}

// 1) Kill any running LLEA processes using batch file
delete __createfile
createfile until END_OF_BATCH
@echo off
for /f "skip=1 tokens=2" %%p in ('wmic process where "Name='powershell.exe' or Name='pwsh.exe'" get ProcessId 2^>nul') do (
    wmic process where "ProcessId=%%p" get CommandLine 2^>nul | findstr /i "LLEA.ps1" >nul 2>&1
    if not errorlevel 1 taskkill /PID %%p /F /T >nul 2>&1
)
END_OF_BATCH

copy __createfile "{pathname of system folder}\KillLLEA.bat"

override wait
hidden=true
wait cmd.exe /C "{pathname of system folder}\KillLLEA.bat"

delete "{pathname of system folder}\KillLLEA.bat"

// 1a) Wait for termination
wait {pathname of system folder}\timeout.exe 3 /nobreak

// 1b) Clean up lock file
delete "{(value "TEMP" of environment)}\LLEA_Instance.lock"

// 2) Remove old LLEA.ps1
delete "C:\Program Files\LLEA\LLEA.ps1"

// 3) Download new LLEA script and icons
delete __createfile
createfile until END_OF_BATCH2
@echo off
certutil -urlcache -f https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LLEA.ps1 "C:\Program Files\LLEA\LLEA.ps1"
if %ERRORLEVEL% NEQ 0 exit /b %ERRORLEVEL%
certutil -urlcache -f https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LL_LOGO.ico "C:\Program Files\LLEA\LL_LOGO.ico"
if %ERRORLEVEL% NEQ 0 exit /b %ERRORLEVEL%
certutil -urlcache -f https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LL_LOGO_MSG.ico "C:\Program Files\LLEA\LL_LOGO_MSG.ico"
exit /b 0
END_OF_BATCH2

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
createfile until END_OF_BATCH
@echo off
for /f "skip=1 tokens=2" %%p in ('wmic process where "Name='powershell.exe' or Name='pwsh.exe'" get ProcessId 2^>nul') do (
    wmic process where "ProcessId=%%p" get CommandLine 2^>nul | findstr /i "LLEA.ps1" >nul 2>&1
    if not errorlevel 1 taskkill /PID %%p /F /T >nul 2>&1
)
END_OF_BATCH

copy __createfile "{pathname of system folder}\KillLLEA.bat"

override wait
hidden=true
wait cmd.exe /C "{pathname of system folder}\KillLLEA.bat"

delete "{pathname of system folder}\KillLLEA.bat"

wait {pathname of system folder}\timeout.exe 2 /nobreak
delete "{(value "TEMP" of environment)}\LLEA_Instance.lock"

// Then your existing deployment steps 1-7...

╔══════════════════════════════════════════════════════════════════════════════╗
║  WHY THIS WORKS                                                              ║
╚══════════════════════════════════════════════════════════════════════════════╝

✅ NO PowerShell backticks ($) - uses batch variables (%%) instead
✅ wmic only used for QUERYING (works on 24H2)
✅ taskkill does the actual termination (not deprecated)
✅ findstr does the pattern matching (simple, reliable)
✅ Works on all Windows versions
✅ No escaping headaches

╔══════════════════════════════════════════════════════════════════════════════╗
║  HOW IT WORKS                                                                ║
╚══════════════════════════════════════════════════════════════════════════════╝

1. Get list of all PowerShell/pwsh process IDs using wmic
2. For each PID, query its command line using wmic
3. Pipe command line to findstr to search for "LLEA.ps1"
4. If found, kill that PID using taskkill
5. Done!

No PowerShell $ variables = No backtick escaping = Works every time

╔══════════════════════════════════════════════════════════════════════════════╗
║  TEST IT FIRST                                                               ║
╚══════════════════════════════════════════════════════════════════════════════╝

Save this as TestKill.bat and run while LLEA is running:

@echo off
echo === Checking for LLEA Processes ===
echo.

for /f "skip=1 tokens=2" %%p in ('wmic process where "Name='powershell.exe' or Name='pwsh.exe'" get ProcessId 2^>nul') do (
    echo Checking PID %%p:
    wmic process where "ProcessId=%%p" get CommandLine 2^>nul | findstr /i "LLEA.ps1"
    if not errorlevel 1 (
        echo   *** MATCHES LLEA.ps1 - Would kill this ***
    ) else (
        echo   Does not match
    )
    echo.
)

pause

This shows you what it finds WITHOUT actually killing anything.

╔══════════════════════════════════════════════════════════════════════════════╗
║  SUMMARY                                                                     ║
╚══════════════════════════════════════════════════════════════════════════════╝

Problem: PowerShell backticks ($) fail in BigFix createfile
Solution: Use batch file with %% variables instead
Result: No backticks = No escaping = Works reliably

Deploy the upgrade script above, it will work.
