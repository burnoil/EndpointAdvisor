## ============================================================================
## PRE-INSTALL
## Detect & remove MSI Slack if present, then install MSIX Slack
## ============================================================================

# Path to your MSIX/MSIXBundle in the Files folder
$MsixPath = Join-Path -Path $dirFiles -ChildPath 'Slack-x64.msixbundle'  # or .msix

# 1) Make sure Slack isn’t running (no PSADT UI)
Try {
    Stop-Process -Name 'slack' -Force -ErrorAction SilentlyContinue
} Catch {}

# 2) Detect MSI Slack (robust across 32/64-bit uninstall hives)
#    If found, remove all matching MSI-installed Slack instances
$msiSlack = Get-InstalledApplication -Name 'Slack*' -IncludeUpdatesAndHotfixes:$false |
            Where-Object { $_.ProductCode }   # MSI entries have a ProductCode

if ($msiSlack) {
    Write-Log "Found MSI Slack:`n$($msiSlack | Format-Table DisplayName, DisplayVersion, ProductCode -Auto | Out-String)"
    # Quietly remove every MSI Slack found (handles multiple versions/arches)
    Remove-MSIApplications -Name 'Slack*' -Wildcard -Exact:$false -IncludeUpdatesAndHotfixes:$false `
        -SilentArgs '/qn /norestart' -IgnoreExitCodes '1605,1614,1641,3010'
} else {
    Write-Log 'No MSI Slack detected.'
}

# 3) (Optional) Import code-signing cert if your MSIX is enterprise-signed
# $cerPath = Join-Path $dirFiles 'Slack-Code-Signing.cer'
# if (Test-Path $cerPath) {
#     Execute-Process -Path 'powershell.exe' -Parameters "-NoProfile -ExecutionPolicy Bypass -Command `"Import-Certificate -FilePath '$cerPath' -CertStoreLocation Cert:\LocalMachine\TrustedPeople | Out-Null`""
# }

# 4) Install MSIX (per-machine context). For most orgs, this is sufficient.
Execute-Process -Path 'powershell.exe' -Parameters "-NoProfile -ExecutionPolicy Bypass -Command `"Add-AppxPackage -Path '$MsixPath' -ForceUpdateFromAnyVersion -DisableDevelopmentMode`""

# ---- Alternative (commented): provision for all users instead of Add-AppxPackage ----
# Execute-Process -Path 'dism.exe' -Parameters "/Online /Add-ProvisionedAppxPackage /PackagePath:`"$MsixPath`" /SkipLicense"

# 5) Verify: MSI gone and MSIX present (log only)
Execute-Process -Path 'powershell.exe' -Parameters "-NoProfile -ExecutionPolicy Bypass -Command `"`$msi = @(Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue | Where-Object { `$_.DisplayName -like 'Slack*' -and `$_.PSObject.Properties.Name -contains 'WindowsInstaller' -and `$_.WindowsInstaller -eq 1 }); `$appx = Get-AppxPackage -AllUsers *Slack* -ErrorAction SilentlyContinue; if((`$msi.Count -eq 0) -and `$appx){ Write-Host 'Validation OK: MSI removed, MSIX present.' } else { Write-Host 'Validation: further attention needed.' }`""
