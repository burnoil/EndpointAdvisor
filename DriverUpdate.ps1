# DriverUpdate.ps1
# Log file path
$logPath = "C:\Windows\MITLL\Logs\MS_Update.txt"

# Function to log messages
function Write-UpdateLog {
    param([string]$Message)
    $timestamp = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
    "$timestamp - $Message" | Out-File -Append -FilePath $logPath
}

Write-UpdateLog "===== Driver Update Started ====="

# Set registry for driver updates
Write-UpdateLog "Setting registry for driver updates..."
Set-itemproperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' -Name SetPolicyDrivenUpdateSourceForDriverUpdates -Value 0 -ErrorAction SilentlyContinue

# Install PSWindowsUpdate module if not present
Write-UpdateLog "Checking for PSWindowsUpdate module..."
if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
    Write-UpdateLog "Installing PSWindowsUpdate module..."
    try {
        Install-Module -Name PSWindowsUpdate -Force -ErrorAction Stop
        Write-UpdateLog "PSWindowsUpdate module installed successfully."
    } catch {
        Write-UpdateLog "ERROR: Failed to install PSWindowsUpdate - $($_.Exception.Message)"
        Set-itemproperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' -Name SetPolicyDrivenUpdateSourceForDriverUpdates -Value 1 -ErrorAction SilentlyContinue
        exit 1
    }
}

# Import the module
Write-UpdateLog "MODULE_IMPORT: Importing PSWindowsUpdate module..."
try {
    Import-Module PSWindowsUpdate -Force -ErrorAction Stop
    Write-UpdateLog "MODULE_READY: PSWindowsUpdate module imported successfully."
} catch {
    Write-UpdateLog "ERROR: Failed to import PSWindowsUpdate - $($_.Exception.Message)"
    Set-itemproperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' -Name SetPolicyDrivenUpdateSourceForDriverUpdates -Value 1 -ErrorAction SilentlyContinue
    exit 1
}

# Scan for driver updates first
Write-UpdateLog "SCAN_START: Scanning for driver updates..."
try {
    $updates = Get-WindowsUpdate -UpdateType Driver -MicrosoftUpdate -ErrorAction Stop
    
    if ($updates.Count -eq 0 -or $null -eq $updates) {
        Write-UpdateLog "NO_UPDATES: No driver updates available - system is up to date"
        Write-UpdateLog "Driver update process completed."
        Set-itemproperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' -Name SetPolicyDrivenUpdateSourceForDriverUpdates -Value 1 -ErrorAction SilentlyContinue
        Write-UpdateLog "Resetting registry..."
        Write-UpdateLog "===== Driver Update Completed ====="
        exit 0
    }
    
    Write-UpdateLog "UPDATES_FOUND: $($updates.Count) driver update(s) available"
    
    # Check if any updates require reboot
    $rebootNeeded = $false
    foreach ($update in $updates) {
        $sizeInMB = [math]::Round($update.Size / 1MB, 2)
        $kbInfo = if ($update.KB) { " (KB$($update.KB))" } else { "" }
        Write-UpdateLog "  - $($update.Title)$kbInfo - Size: $sizeInMB MB"
        if ($update.RebootRequired) {
            $rebootNeeded = $true
        }
    }
    
    # Suspend BitLocker ONLY if updates found and reboot needed
    if ($rebootNeeded) {
        Write-UpdateLog "Checking BitLocker status..."
        try {
            $bitlockerstatus = Get-BitLockerVolume -MountPoint "C:" -ErrorAction Stop
            if ($bitlockerstatus.ProtectionStatus -eq 'On') {
                Write-UpdateLog "BitLocker is enabled, suspending for next reboot..."
                Suspend-BitLocker -MountPoint "C:" -RebootCount 1 -ErrorAction Stop
                Write-UpdateLog "BitLocker suspended successfully."
            } else {
                Write-UpdateLog "BitLocker is not enabled."
            }
        } catch {
            Write-UpdateLog "WARNING: BitLocker check/suspend failed - $($_.Exception.Message)"
        }
    }
    
    # Download updates
    Write-UpdateLog "DOWNLOAD_START: Downloading driver updates..."
    Get-WindowsUpdate -UpdateType Driver -MicrosoftUpdate -Download -AcceptAll -ErrorAction Stop
    Write-UpdateLog "DOWNLOAD_COMPLETE: Driver updates downloaded"
    
    # Install updates
    Write-UpdateLog "INSTALL_START: Installing driver updates..."
    Install-WindowsUpdate -UpdateType Driver -MicrosoftUpdate -AcceptAll -IgnoreReboot -ErrorAction Stop
    Write-UpdateLog "INSTALL_COMPLETE: Driver updates installed successfully"
    
    if ($rebootNeeded) {
        Write-UpdateLog "REBOOT_REQUIRED: System restart is required to complete installation"
        Write-UpdateLog "REBOOT_SCHEDULED: Reboot scheduled for 5 minutes"
        # Schedule restart in 5 minutes
        shutdown /r /t 300 /c "Driver updates have been installed. Your computer will restart in 5 minutes to complete the installation. Please save your work."
    } else {
        Write-UpdateLog "REBOOT_NOT_REQUIRED: No restart needed"
    }
    
} catch {
    Write-UpdateLog "ERROR: Failed to check/install updates - $($_.Exception.Message)"
    Set-itemproperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' -Name SetPolicyDrivenUpdateSourceForDriverUpdates -Value 1 -ErrorAction SilentlyContinue
    exit 1
}

# Reset registry
Write-UpdateLog "Resetting registry..."
Set-itemproperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' -Name SetPolicyDrivenUpdateSourceForDriverUpdates -Value 1 -ErrorAction SilentlyContinue

Write-UpdateLog "Driver update process completed."
Write-UpdateLog "===== Driver Update Completed ====="
