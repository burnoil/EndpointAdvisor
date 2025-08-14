function Update-PatchingAndSystem {
    Write-Log "Updating Patching and System section..." -Level "INFO"
    $restartStatusText = Get-PendingRestartStatus
    $windowsBuild = Get-WindowsBuildNumber
    
    # --- BigFix Update Logic ---
    $fixletPath = "C:\temp\X-Fixlet-Source_Count.txt"
    $bigfixStatusText = "Application Updates: No Updates Pending."
    $showBigFixButton = $false
    try {
        if (Test-Path $fixletPath) {
            $fileContent = Get-Content -Path $fixletPath
            if ($fileContent) { 
                $multiLineContent = $fileContent -join "`n"
                $bigfixStatusText = "Application Updates:`n" + $multiLineContent
                $showBigFixButton = $true
                Write-Log "Successfully read fixlet data from $fixletPath" -Level "INFO"
            } else {
                Write-Log "$fixletPath is empty." -Level "INFO"
            }
        } else {
            Write-Log "$fixletPath not found." -Level "WARNING"
        }
    } catch {
        $bigfixStatusText = "Application Updates: Error reading update data."
        Write-Log "Error reading BigFix data: $($_.Exception.Message)" -Level "ERROR"
    }

    # --- ECM Update Logic ---
    $ecmResult = Get-ECMUpdateStatus
    $ecmStatusText = $ecmResult.StatusText
    $showEcmButton = $ecmResult.HasPendingUpdates

    # --- Update the UI ---
    $window.Dispatcher.Invoke({
        # Conditionally display the restart panel
        if ($global:PendingRestart) {
            $global:PendingRestartPanel.Visibility = "Visible"
            $global:PendingRestartStatusText.Text = $restartStatusText
            $global:PendingRestartStatusText.Foreground = [System.Windows.Media.Brushes]::Red
        } else {
            $global:PendingRestartPanel.Visibility = "Collapsed"
        }
        
        $global:WindowsBuildText.Text = $windowsBuild
        $global:FooterText.Text = "(C) 2025 Lincoln Laboratory v$ScriptVersion"
        
        # Update BigFix UI elements
        $global:BigFixStatusText.Text = $bigfixStatusText
        $global:BigFixLaunchButton.Visibility = if ($showBigFixButton) { "Visible" } else { "Collapsed" }
        
        # Update ECM UI elements
        $global:ECMStatusText.Text = $ecmStatusText
        $global:ECMLaunchButton.Visibility = if ($showEcmButton) { "Visible" } else { "Collapsed" }
        
        # The default tooltip from the XAML is now sufficient.
    })
}
