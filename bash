Step 0
// Kill any existing LLEA instances before deploying
override wait
hidden=true
wait powershell.exe -NoProfile -Command "Get-Process -Name 'powershell','pwsh' -ErrorAction SilentlyContinue | Where-Object { try { (Get-WmiObject Win32_Process -Filter \"ProcessId = $($_.Id)\" -ErrorAction SilentlyContinue).CommandLine -like '*LLEA.ps1*' } catch { $false } } | Stop-Process -Force -ErrorAction SilentlyContinue"

wait {pathname of system folder}\timeout.exe 2 /nobreak
delete "{(value "TEMP" of environment)}\LLEA_Instance.lock"

between 6 and 7

// Brief pause to ensure cleanup complete
wait {pathname of system folder}\timeout.exe 1 /nobreak
