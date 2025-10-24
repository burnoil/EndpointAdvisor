## ============================================================================
## PRE-INSTALL
## Detect and remove MSI Slack before installing MSIX Slack
## ============================================================================

# Path to the MSIX or MSIXBundle in your Files directory
$MsixPath = Join-Path -Path $dirFiles -ChildPath 'Slack-x64.msixbundle'  # adjust as needed

# 1. Stop Slack if it’s running
Try {
    Stop-Process -Name 'slack' -Force -ErrorAction SilentlyContinue
} Catch {}

# 2. Detect any MSI-installed Slack (32-bit or 64-bit)
$SlackMsiApps = Get-InstalledApplication -Name 'Slack*' |
    Where-Object { $_.ProductCode -and $_.WindowsInstaller -eq $true }

if ($SlackMsiApps) {
    Write-Log "Found MSI Slack install(s): $($SlackMsiApps.DisplayName -join ', ')"
    # Remove all MSI Slack installations quietly
    Remove-MSIApplications -Name 'Slack*' -Exact:$false -SilentArgs '/qn /norestart' `
        -IgnoreExitCodes '1605,1614,1641,3010'
}
else {
    Write-Log 'No MSI Slack installation found.'
}

# 3. (Optional) Import code-signing certificate if needed
# $cerPath = Join-Path $dirFiles 'Slack-Code-Signing.cer'
# if (Test-Path $cerPath) {
#     Execute-Process -Path 'powershell.exe' -Parameters "-NoProfile -ExecutionPolicy Bypass -Command `"Import-Certificate -FilePath '$cerPath' -CertStoreLocation Cert:\LocalMachine\TrustedPeople | Out-Null`""
# }

# 4. Install the MSIX package
Execute-Process -Path 'powershell.exe' -Parameters "-NoProfile -ExecutionPolicy Bypass -Command `"Add-AppxPackage -Path '$MsixPath' -ForceUpdateFromAnyVersion -DisableDevelopmentMode`""

# --- Optional all-users provisioning ---
# Execute-Process -Path 'dism.exe' -Parameters "/Online /Add-ProvisionedAppxPackage /PackagePath:`"$MsixPath`" /SkipLicense"

Write-Log 'Slack MSIX installation complete.'
