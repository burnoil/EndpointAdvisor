# This version is CORRECT for BigFix actions:
override wait
hidden=true
wait powershell.exe -NoProfile -Command "Get-Process -Name 'powershell','pwsh' -ErrorAction SilentlyContinue | Where-Object { try { (Get-WmiObject Win32_Process -Filter \"ProcessId = $($_.Id)\" -ErrorAction SilentlyContinue).CommandLine -like '*LLEA.ps1*' } catch { $false } } | Stop-Process -Force -ErrorAction SilentlyContinue"
