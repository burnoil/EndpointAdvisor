delete __createfile
createfile until END_OF_KILL_SCRIPT
`$processes = Get-Process -Name powershell,pwsh -ErrorAction SilentlyContinue
foreach (`$p in `$processes) {
    # Method 1: Try Get-CimInstance instead of Get-WmiObject
    try {
        `$cim = Get-CimInstance Win32_Process -Filter "ProcessId=`$(`$p.Id)"
        if (`$cim.CommandLine -like '*LLEA.ps1*') {
            Stop-Process -Id `$p.Id -Force
        }
    } catch {
        # Method 2: Fallback to WMI without filter
        try {
            `$wmi = Get-WmiObject Win32_Process | Where-Object { `$_.ProcessId -eq `$p.Id }
            if (`$wmi.CommandLine -like '*LLEA.ps1*') {
                Stop-Process -Id `$p.Id -Force
            }
        } catch {
            # Method 3: Last resort - check MainWindowTitle
            if (`$p.MainWindowTitle -like '*Lincoln Laboratory Endpoint Advisor*') {
                Stop-Process -Id `$p.Id -Force
            }
        }
    }
}
END_OF_KILL_SCRIPT

move __createfile "{pathname of system folder}\KillLLEA.ps1"

override wait
hidden=true
wait powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{pathname of system folder}\KillLLEA.ps1"

delete "{pathname of system folder}\KillLLEA.ps1"
