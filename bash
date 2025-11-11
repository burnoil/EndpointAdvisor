delete __createfile
createfile until END_OF_KILL_SCRIPT
# Kill by window title
Get-Process | Where-Object {
    `$_.MainWindowTitle -like '*Lincoln Laboratory Endpoint Advisor*' -or
    `$_.MainWindowTitle -like '*LLEA*'
} | Stop-Process -Force -ErrorAction SilentlyContinue

# Also try to find by command line (all PowerShell processes)
Get-CimInstance Win32_Process -Filter "Name='powershell.exe' OR Name='pwsh.exe'" | 
    Where-Object { `$_.CommandLine -like '*LLEA.ps1*' } |
    ForEach-Object { Stop-Process -Id `$_.ProcessId -Force -ErrorAction SilentlyContinue }
END_OF_KILL_SCRIPT

move __createfile "{pathname of system folder}\KillLLEA.ps1"

override wait
hidden=true
wait powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{pathname of system folder}\KillLLEA.ps1"

delete "{pathname of system folder}\KillLLEA.ps1"
