# --- SAP AO post-check: install ONLY if it existed pre-upgrade (flag set) ---
try {
    $needReinstall = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\LL\OfficeUpgrade' -ErrorAction SilentlyContinue).ReinstallSAPAO -eq 1
    $archPref      = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\LL\OfficeUpgrade' -ErrorAction SilentlyContinue).SAPAOArch
    $archToInstall = if ($archPref) { $archPref } else { 'x64' }

    if ($needReinstall) {
        $stateAfter = Get-SAPAOState

        if ($stateAfter.Present) {
            # It survived the Office switch — do NOT reinstall
            Write-ADTLogEntry -Message "SAP AO was flagged pre-upgrade but is still present post-upgrade ($($stateAfter.DisplayName) $($stateAfter.Version)). Skipping reinstall." -Severity 1 -Source $adtSession.InstallPhase
        } else {
            # It was present before, missing now → reinstall
            Write-ADTLogEntry -Message "SAP AO was present pre-upgrade and is missing post-upgrade. Installing ($archToInstall)..." -Severity 1 -Source $adtSession.InstallPhase
            Install-SAPAOIfNeeded -Architecture $archToInstall -Force
        }

        # Clean up the flag regardless (we handled the case)
        Remove-ItemProperty -Path 'HKLM:\SOFTWARE\LL\OfficeUpgrade' -Name 'ReinstallSAPAO' -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path 'HKLM:\SOFTWARE\LL\OfficeUpgrade' -Name 'SAPAOArch' -ErrorAction SilentlyContinue
    }
    else {
        # No pre-upgrade SAP AO → never install
        Write-ADTLogEntry -Message "No SAP AO pre-upgrade flag present; skipping any SAP AO actions." -Severity 1 -Source $adtSession.InstallPhase
    }
}
catch {
    Write-ADTLogEntry -Message "SAP AO post-check failed: $($_.Exception.Message)" -Severity 2 -Source $adtSession.InstallPhase
}
