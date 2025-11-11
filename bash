// 🛑 Kill by Command Line (Modern PowerShell Method)
// This is the most reliable method. It works as SYSTEM
// because the CommandLine property is visible.
override wait
hidden=true
wait powershell.exe -NoProfile -Command "Get-CimInstance Win32_Process -Filter \"CommandLine LIKE '%LLEA.ps1%'\" | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"

// 🛑 Kill by Command Line (Taskkill Method)
// This uses the built-in taskkill.exe with a WMI filter
// to find and kill the process.
override wait
hidden=true
wait taskkill.exe /F /FI "CommandLine like *LLEA.ps1*" /T
