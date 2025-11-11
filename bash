# Check what LLEA's command line actually looks like
$processes = Get-Process -Name powershell,pwsh -ErrorAction SilentlyContinue
foreach ($p in $processes) {
    $wmi = Get-WmiObject Win32_Process -Filter "ProcessId=$($p.Id)" -ErrorAction SilentlyContinue
    if ($wmi) {
        Write-Host "PID: $($p.Id)"
        Write-Host "  CommandLine: $($wmi.CommandLine)"
        Write-Host "  Matches LLEA.ps1: $($wmi.CommandLine -like '*LLEA.ps1*')"
        Write-Host ""
    }
}
