# LLNOTIFY.ps1 - Lincoln Laboratory Notification System
# Version 4.3.33 (Fixed try/catch parsing error in Get-BigFixRelevanceResult, prioritized piped input for QnA.exe)

# Ensure $PSScriptRoot is defined for older versions
if ($MyInvocation.MyCommand.Path) {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    $ScriptDir = Get-Location
}

# Define version
$ScriptVersion = "4.3.33"

# Global flag to prevent recursive logging during rotation
$global:IsRotatingLog = $false

# Global flag to track pending restart state
$global:PendingRestart = $false

# Global variables for certificate check caching
$global:LastCertificateCheck = $null
$global:CachedCertificateStatus = $null

# ... [Previous sections: Write-Log, Invoke-WithRetry, Rotate-LogFile, Handle-Error unchanged] ...

Write-Log "--- LLNOTIFY Script Started (Version $ScriptVersion) ---"

# ============================================================
# BigFix Compliance Reporting Functions
# ============================================================
function Get-BigFixRelevanceResult {
    param([string]$RelevanceQuery)
    try {
        $qnaPath = $config.BigFixQnA_Path
        if (-not $qnaPath -or -not (Test-Path $qnaPath)) {
            throw "QnA.exe not found at configured path: '$qnaPath'. Please verify BigFix client installation."
        }

        # Log QnA.exe version for debugging
        try {
            $qnaVersion = & $qnaPath -version 2>&1
            Write-Log "QnA.exe version: $qnaVersion" -Level "INFO"
        } catch {
            Write-Log "Failed to retrieve QnA.exe version: $($_.Exception.Message)" -Level "WARNING"
        }

        Write-Log "Executing QnA.exe for query: $RelevanceQuery" -Level "INFO"

        # Try piped input first (more reliable across versions)
        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = $qnaPath
        $processInfo.RedirectStandardInput = $true
        $processInfo.RedirectStandardOutput = $true
        $processInfo.RedirectStandardError = $true
        $processInfo.UseShellExecute = $false
        $processInfo.CreateNoWindow = $true

        Write-Log "Running command: echo $RelevanceQuery | $qnaPath" -Level "INFO"
        $process = [System.Diagnostics.Process]::Start($processInfo)
        $process.StandardInput.WriteLine($RelevanceQuery)
        $process.StandardInput.Close()
        $output = $process.StandardOutput.ReadToEnd()
        $errorOutput = $process.StandardError.ReadToEnd()
        $process.WaitForExit()
        $exitCode = $process.ExitCode

        if ($exitCode -eq 0 -and -not [string]::IsNullOrWhiteSpace($output)) {
            $result = $output.Trim()
            if ($result -is [array]) {
                $result = $result -join "`n"
            }
            Write-Log "QnA piped query succeeded: $result" -Level "INFO"
            return $result
        }

        # Fallback to -f option if piped input fails
        Write-Log "Piped input failed with exit code $exitCode`: $errorOutput. Trying -f option." -Level "WARNING"
        $tempFile = [System.IO.Path]::GetTempFileName()
        $RelevanceQuery | Out-File -FilePath $tempFile -Encoding ASCII -Force
        
        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = $qnaPath
        $processInfo.Arguments = "-f `"$tempFile`""
        $processInfo.RedirectStandardOutput = $true
        $processInfo.RedirectStandardError = $true
        $processInfo.UseShellExecute = $false
        $processInfo.CreateNoWindow = $true

        Write-Log "Running command: $qnaPath -f `"$tempFile`"" -Level "INFO"
        $process = [System.Diagnostics.Process]::Start($processInfo)
        $output = $process.StandardOutput.ReadToEnd()
        $errorOutput = $process.StandardError.ReadToEnd()
        $process.WaitForExit()
        $exitCode = $process.ExitCode

        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
        
        if ($exitCode -ne 0) {
            throw "QnA.exe -f failed with exit code $exitCode`: $errorOutput"
        }
        if ([string]::IsNullOrWhiteSpace($output)) {
            Write-Log "QnA.exe -f returned empty output for query: $RelevanceQuery" -Level "WARNING"
            return "No data returned"
        }
        $result = $output.Trim()
        if ($result -is [array]) {
            $result = $result -join "`n"
        }
        Write-Log "QnA -f query succeeded: $result" -Level "INFO"
        return $result
    }
    catch {
        Write-Log "QnA query failed: $($_.Exception.Message)" -Level "ERROR"
        return "Error: BigFix QnA failed: $($_.Exception.Message)"
    }
}

function Generate-BigFixComplianceReport {
    try {
        Write-Log "Gathering BigFix compliance info using QnA.exe..." -Level "INFO"
        $reportPath = Join-Path $ScriptDir "BigFixComplianceReport.txt"
        $jsonPath = Join-Path $ScriptDir "BigFixComplianceReport.json"

        $computerName = Get-BigFixRelevanceResult "name of computer"
        $clientVersion = Get-BigFixRelevanceResult "version of client as string"
        $relay = Get-BigFixRelevanceResult "if exists relay service then (address of relay service as string) else `"No Relay`""
        $lastReport = Get-BigFixRelevanceResult "last report time of client as string"
        $ipAddress = Get-BigFixRelevanceResult "ip address of client as string"
        $fixletList = Get-BigFixRelevanceResult "names of relevant fixlets whose (baseline flag of it = false and (name of it as lowercase contains `"microsoft`" or name of it as lowercase contains `"security update`")) of sites"

        $fixlets = @()
        if ($fixletList -is [string] -and -not [string]::IsNullOrWhiteSpace($fixletList) -and -not $fixletList.StartsWith("Error:")) {
            $fixlets = $fixletList -split "`n" | Where-Object { $_ -match "\S" }
        } elseif ($fixletList.StartsWith("Error:")) {
            $fixlets = @($fixletList)
        }

        $report = @(
            "BigFix Compliance Report - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
            "------------------------------------------------------------",
            "Computer Name  : $computerName",
            "Client Version : $clientVersion",
            "Relay Address  : $relay",
            "IP Address     : $ipAddress",
            "Last Reported  : $lastReport",
            "",
            "Applicable Fixlets (Non-Baselines):",
            "----------------------------------"
        )
        if ($fixlets.Count -gt 0 -and -not $fixlets[0].StartsWith("Error:")) {
            $report += $fixlets | ForEach-Object { " - $_" }
        } else {
            if ($fixlets[0]) {
                $report += $fixlets[0]
            } else {
                $report += "No applicable fixlets found."
            }
        }

        $report | Out-File -FilePath $reportPath -Encoding UTF8
        $reportData = @{
            Timestamp         = (Get-Date)
            ComputerName      = $computerName
            ClientVersion     = $clientVersion
            Relay             = $relay
            IPAddress         = $ipAddress
            LastReportTime    = $lastReport
            ApplicableFixlets = $fixlets
        }
        $reportData | ConvertTo-Json -Depth 3 | Out-File -FilePath $jsonPath -Encoding UTF8
        Write-Log "BigFix compliance report written to $reportPath and $jsonPath" -Level "INFO"
        
        return $reportData
    } catch {
        Write-Log "Error generating BigFix compliance report: $($_.Exception.Message)" -Level "ERROR"
        return $null
    }
}

# ... [Previous sections: Get-DefaultConfig, Load-Configuration, Save-Configuration unchanged] ...

# ============================================================
# F) Load and Verify XAML
# ============================================================
try {
    Write-Log "Loading XAML..." -Level "INFO"
    $xmlDoc = New-Object System.Xml.XmlDocument
    $xmlDoc.LoadXml($xamlString)
    $reader = New-Object System.Xml.XmlNodeReader $xmlDoc
    [System.Windows.Window]$global:window = [Windows.Markup.XamlReader]::Load($reader)
    Write-Log "XAML loaded successfully." -Level "INFO"

    $window.Width = 350
    $window.Height = 500

    $window.DataContext = [PSCustomObject]@{ 
        MainIconUri   = [Uri]$mainIconUri
        WindowIconUri = $mainIconPath
    }

    $uiElements = @(
        "AnnouncementsExpander", "AnnouncementsAlertIcon", "AnnouncementsText", "AnnouncementsDetailsText",
        "AnnouncementsLinksPanel", "AnnouncementsSourceText", "PatchingExpander", "PatchingDescriptionText",
        "PendingRestartStatusText", "PatchingUpdatesText", "PatchingSSAButton", "SupportExpander",
        "SupportAlertIcon", "SupportText", "SupportLinksPanel", "SupportSourceText", "ComplianceExpander",
        "YubiKeyComplianceText", "WindowsBuildText", "ClearAlertsButton", "ScriptUpdateText", "FooterText",
        "BigFixClientInfoText"
    )
    foreach ($elementName in $uiElements) {
        Set-Variable -Name "global:$elementName" -Value $window.FindName($elementName)
    }
    Write-Log "UI elements mapped to variables." -Level "INFO"

    $global:FooterText.Text = "© 2025 Lincoln Laboratory v$ScriptVersion"

    $global:AnnouncementsExpander.Add_Expanded({ $window.Dispatcher.Invoke({ $global:AnnouncementsAlertIcon.Visibility = "Hidden"; Update-TrayIcon }) })
    $global:SupportExpander.Add_Expanded({ $window.Dispatcher.Invoke({ $global:SupportAlertIcon.Visibility = "Hidden"; Update-TrayIcon }) })

    $global:PatchingSSAButton.Add_Click({
        try {
            $ssaPath = $config.BigFixSSA_Path
            if ([string]::IsNullOrWhiteSpace($ssaPath) -or -not (Test-Path $ssaPath)) {
                throw "BigFix Self-Service Application path is invalid or not found: `"$ssaPath`""
            }
            Write-Log "Launching BigFix SSA: $ssaPath" -Level "INFO"
            Start-Process -FilePath $ssaPath
        }
        catch {
            Handle-Error $_.Exception.Message -Source "PatchingSSAButton"
        }
    })
    
    $global:ClearAlertsButton.Add_Click({
        Write-Log "Clear Alerts button clicked by user." -Level "INFO"

        if ($global:contentData) {
            $config.AnnouncementsLastState = $global:contentData.Data.Announcements | ConvertTo-Json -Compress
            $config.SupportLastState = $global:contentData.Data.Support | ConvertTo-Json -Compress
        }

        $window.Dispatcher.Invoke({
            $global:AnnouncementsAlertIcon.Visibility = 'Hidden'
            $global:SupportAlertIcon.Visibility = 'Hidden'
        })

        $global:BlinkingTimer.Stop()
        Update-TrayIcon
        
        Save-Configuration -Config $config
    })

    $window.Add_Closing({
        if (-not $global:IsUpdating) {
            $_.Cancel = $true
            $window.Hide()
        }
    })
}
catch {
    Handle-Error "Failed to load the XAML layout: $($_.Exception.Message)" -Source "XAML"
    exit
}

# ... [Rest of the script unchanged from version 4.3.32, including Update-PatchingAndSystem, Initialize-TrayIcon, etc.] ...
