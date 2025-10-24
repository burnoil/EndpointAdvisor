# 2) Find Slack MSI product codes from both 64/32-bit uninstall hives
$slackMsiGuids = Get-ChildItem `
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall' -ErrorAction SilentlyContinue |
  ForEach-Object { $_ | Get-ItemProperty -ErrorAction SilentlyContinue } |
  Where-Object {
      $_.DisplayName -like 'Slack*' -and
      $_.PSObject.Properties.Name -contains 'WindowsInstaller' -and
      $_.WindowsInstaller -eq 1 -and
      $_.PSObject.Properties.Name -contains 'UninstallString'
  } |
  ForEach-Object {
      # Extract {GUID} from UninstallString (handles /I or /X forms)
      if ($_.UninstallString -match '{[0-9A-Fa-f\-]+}') { $matches[0] }
  } |
  Sort-Object -Unique

if ($slackMsiGuids) {
    Write-Log "MSI Slack detected: $($slackMsiGuids -join ', ')"
    foreach ($guid in $slackMsiGuids) {
        Execute-Process -Path 'msiexec.exe' -Parameters "/x $guid /qn /norestart" `
            -IgnoreExitCodes '1605,1614,1641,3010'
    }
} else {
    Write-Log 'No MSI Slack found.'
}
