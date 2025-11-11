action uses wow64 redirection {not x64 of operating system}

// 1) Kill any running LLNotify processes via WMIC
override wait
hidden=true
wait cmd.exe /C wmic process where "CommandLine like '%LLEA.ps1%'" call terminate

// 2) Remove LLEA.ps1
delete C:\Program Files\LLEA\LLEA.ps1
delete C:\Program Files\LLEA\LLEndpointAdvisor.config.json

// 3) Download new LLEA script and run
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

// 4) Drop the batch into place
copy __createfile "C:\Program Files\LLEA\download_LLEA.bat"

// 5) Run it as the current user, hidden
override wait
hidden=true
wait cmd.exe /C "C:\Program Files\LLEA\download_LLEA.bat"

// 6) (Optional) Clean up the batch once successful
delete "C:\Program Files\LLEA\download_LLEA.bat"

// 7) Immediately invoke the (signed) script once
override wait
hidden=true
runas=currentuser  
wait cmd.exe /C start "" /b powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\Program Files\LLEA\LLEA.ps1" -RunMode LLEA

