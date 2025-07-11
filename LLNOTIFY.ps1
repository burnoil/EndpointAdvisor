# LLNOTIFY.ps1 - Lincoln Laboratory Notification System
# Version 4.3.5 (Added color to Clear Alerts button)

# Ensure $PSScriptRoot is defined for older versions
if ($MyInvocation.MyCommand.Path) {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    $ScriptDir = Get-Location
}

# Define version
$ScriptVersion = "4.3.5"

# Global flag to prevent recursive logging during rotation
$global:IsRotatingLog = $false

# Global flag to track pending restart state
$global:PendingRestart = $false

# Global variables for certificate check caching
$global:LastCertificateCheck = $null
$global:CachedCertificateStatus = $null

# ============================================================
# A) Advanced Logging & Error Handling
# ============================================================
function Invoke-WithRetry {
    param(
        [ScriptBlock]$Action,
        [int]$MaxRetries = 3,
        [int]$RetryDelayMs = 100
    )
    $attempt = 0
    while ($attempt -lt $MaxRetries) {
        try {
            $attempt++
            $Action.Invoke()
            return # Success
        }
        catch {
            if ($attempt -ge $MaxRetries) {
                throw "Action failed after $MaxRetries attempts: $($_.Exception.Message)"
            }
            Start-Sleep -Milliseconds $RetryDelayMs
        }
    }
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )
    if ($global:IsRotatingLog) {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message (Skipped due to log rotation)"
        return
    }

    $logPath = if ($LogFilePath) { $LogFilePath } else { Join-Path $ScriptDir "LLNOTIFY.log" }
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    $maxRetries = 3
    $retryDelayMs = 100
    $attempt = 0
    $success = $false
    
    while ($attempt -lt $maxRetries -and -not $success) {
        try {
            $attempt++
            Add-Content -Path $logPath -Value $logEntry -Force -ErrorAction Stop
            $success = $true
        }
        catch {
            if ($attempt -eq $maxRetries) {
                Write-Host "[$timestamp] [$Level] $Message (Failed to write to log after $maxRetries attempts: $($_.Exception.Message))"
            } else {
                Start-Sleep -Milliseconds $retryDelayMs
            }
        }
    }
}

function Rotate-LogFile {
    try {
        if (Test-Path $LogFilePath) {
            $fileInfo = Get-Item $LogFilePath
            $maxSizeBytes = $config.LogRotationSizeMB * 1MB
            if ($fileInfo.Length -gt $maxSizeBytes) {
                $archivePath = "$LogFilePath.$(Get-Date -Format 'yyyyMMddHHmmss').archive"
                
                $global:IsRotatingLog = $true
                
                try {
                    Invoke-WithRetry -Action {
                        Rename-Item -Path $LogFilePath -NewName $archivePath -ErrorAction Stop
                    }
                    Write-Log "Log file rotated. Archived as $archivePath" -Level "INFO"

                    $archiveFiles = Get-ChildItem -Path $LogDirectory -Filter "LLNOTIFY.log.*.archive" | Sort-Object CreationTime
                    $maxArchives = 3
                    if ($archiveFiles.Count -gt $maxArchives) {
                        $filesToDelete = $archiveFiles | Select-Object -First ($archiveFiles.Count - $maxArchives)
                        foreach ($file in $filesToDelete) {
                            try {
                                Invoke-WithRetry -Action {
                                    Remove-Item -Path $file.FullName -Force -ErrorAction Stop
                                }
                                Write-Log "Deleted old archive: $($file.FullName)" -Level "INFO"
                            }
                            catch {
                                Write-Log "Failed to delete old archive $($file.FullName): $($_.Exception.Message)" -Level "ERROR"
                            }
                        }
                    }
                }
                catch {
                    Write-Log "Failed to rotate log file: $($_.Exception.Message)" -Level "ERROR"
                }
            }
        }
    }
    catch {
        Write-Log "Error checking log file size for rotation: $($_.Exception.Message)" -Level "ERROR"
    }
    finally {
        $global:IsRotatingLog = $false
    }
}

function Handle-Error {
    param(
        [string]$ErrorMessage,
        [string]$Source = ""
    )
    if ($Source) { $ErrorMessage = "[$Source] $ErrorMessage" }
    Write-Log $ErrorMessage -Level "ERROR"
}

Write-Log "--- LLNOTIFY Script Started (Version $ScriptVersion) ---"

# ============================================================
# MODULE: Configuration Management
# ============================================================
function Get-DefaultConfig {
    return @{
        RefreshInterval       = 90
        LogRotationSizeMB     = 2
        DefaultLogLevel       = "INFO"
        ContentDataUrl        = "https://raw.githubusercontent.com/burnoil/LLNOTIFY/refs/heads/main/ContentData.json"
        CertificateCheckInterval = 86400
        YubiKeyAlertDays      = 14
        IconPaths             = @{
            Main    = Join-Path $ScriptDir "LL_LOGO.ico"
            Warning = Join-Path $ScriptDir "LL_LOGO_MSG.ico"
        }
        AnnouncementsLastState = "{}"
        SupportLastState       = "{}"
        Version               = $ScriptVersion
        PatchInfoFilePath     = "C:\temp\X-Fixlet-Source_Count.txt"
        BigFixSSA_Path        = "C:\Program Files (x86)\BigFix Enterprise\BigFix Self Service Application\BigFixSSA.exe"
        YubiKeyManager_Path   = "C:\Program Files\Yubico\Yubikey Manager\ykman.exe"
        BlinkingEnabled       = $true
        ScriptUrl             = "https://raw.githubusercontent.com/burnoil/LLNOTIFY/refs/heads/main/LLNOTIFY.ps1"
    }
}

function Load-Configuration {
    param([string]$Path = (Join-Path $ScriptDir "LLNOTIFY.config.json"))
    $finalConfig = Get-DefaultConfig
    if (Test-Path $Path) {
        try {
            $loadedConfig = Get-Content $Path -Raw | ConvertFrom-Json
            if ($loadedConfig) {
                foreach ($key in $loadedConfig.PSObject.Properties.Name) {
                    if ($finalConfig.ContainsKey($key) -and $loadedConfig.$key -ne $null) {
                        $finalConfig[$key] = $loadedConfig.$key
                    }
                }
            }
        }
        catch {
            Write-Log "Failed to load or merge existing config file. Reverting to full defaults. Error: $($_.Exception.Message)" -Level "WARNING"
        }
    }
    try {
        $finalConfig | ConvertTo-Json -Depth 8 | Out-File $Path -Force
        Write-Log "Configuration file validated and saved." -Level "INFO"
    }
    catch {
        Handle-Error "Could not save the updated configuration to '$Path'. Error: $($_.Exception.Message)"
    }
    return $finalConfig
}

function Save-Configuration {
    param(
        [psobject]$Config,
        [string]$Path = (Join-Path $ScriptDir "LLNOTIFY.config.json")
    )
    try {
        $Config | ConvertTo-Json -Depth 8 | Out-File $Path -Force
    } catch {
        Handle-Error "Could not save state to configuration file '$Path'. Error: $($_.Exception.Message)"
    }
}

# ============================================================
# B) External Configuration Setup
# ============================================================
$LogFilePath = Join-Path $ScriptDir "LLNOTIFY.log"
$config = Load-Configuration

$mainIconPath = $config.IconPaths.Main
$warningIconPath = $config.IconPaths.Warning
$mainIconUri = "file:///$($mainIconPath -replace '\\','/')"

Write-Log "Main icon path: $mainIconPath" -Level "INFO"
Write-Log "Warning icon path: $warningIconPath" -Level "INFO"

$defaultContentData = @{
    Announcements = @{ Text = "No announcements at this time."; Details = ""; Links = @() }
    Support = @{ Text  = "Contact IT Support."; Links = @() }
}

# ============================================================
# C) Log File Setup & Rotation
# ============================================================
$LogDirectory = Split-Path $LogFilePath
if (-not (Test-Path $LogDirectory)) {
    New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
}
Rotate-LogFile

function Log-DotNetVersion {
    try {
        $dotNetVersion = [System.Environment]::Version.ToString()
        Write-Log ".NET Version: $dotNetVersion" -Level "INFO"
        $frameworkDescription = [System.Runtime.InteropServices.RuntimeInformation]::FrameworkDescription
        Write-Log ".NET Framework Description: $frameworkDescription" -Level "INFO"
    } catch {}
}
# ============================================================
# D) Import Required Assemblies
# ============================================================
function Import-RequiredAssemblies {
    try {
        Write-Log "Loading required .NET assemblies..." -Level "INFO"
        Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
        Write-Log "Loaded PresentationFramework." -Level "INFO"
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Write-Log "Loaded System.Windows.Forms." -Level "INFO"
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        Write-Log "Loaded System.Drawing." -Level "INFO"
        return $true
    }
    catch {
        Write-Log "Failed to load required GUI assemblies: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

$global:FormsAvailable = Import-RequiredAssemblies
# ============================================================
# E) XAML Layout Definition
# ============================================================
$xamlString = @"
<?xml version="1.0" encoding="utf-8"?>
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="LLNOTIFY - Lincoln Laboratory Notification System"
    WindowStartupLocation="Manual" 
    SizeToContent="Manual"
    MinWidth="350" MinHeight="500"
    MaxWidth="400" MaxHeight="550"
    ResizeMode="CanResizeWithGrip" ShowInTaskbar="False" Visibility="Hidden" Topmost="True"
    Background="#f0f0f0"
    Icon="{Binding WindowIconUri}">
  <Grid Margin="5">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <Border Grid.Row="0" Background="#0078D7" Padding="5" CornerRadius="3" Margin="0,0,0,5">
      <StackPanel Orientation="Horizontal" VerticalAlignment="Center" HorizontalAlignment="Center">
        <Image Source="{Binding MainIconUri}" Width="20" Height="20" Margin="0,0,5,0"/>
        <TextBlock Text="Lincoln Laboratory Notification System" FontSize="14" FontWeight="Bold" Foreground="White" VerticalAlignment="Center"/>
      </StackPanel>
    </Border>
    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
      <StackPanel VerticalAlignment="Top">
        <Expander x:Name="AnnouncementsExpander" FontSize="12" IsExpanded="True" Margin="0,2,0,2">
          <Expander.Header>
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="Announcements" VerticalAlignment="Center"/>
              <Ellipse x:Name="AnnouncementsAlertIcon" Width="10" Height="10" Margin="5,0,0,0" Fill="Red" Visibility="Hidden"/>
            </StackPanel>
          </Expander.Header>
          <Border BorderBrush="#00008B" BorderThickness="1" Padding="5" CornerRadius="3" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="AnnouncementsText" FontSize="11" TextWrapping="Wrap"/>
              <TextBlock x:Name="AnnouncementsDetailsText" FontSize="11" TextWrapping="Wrap" Margin="0,5,0,0"/>
              <StackPanel x:Name="AnnouncementsLinksPanel" Orientation="Vertical" Margin="0,5,0,0"/>
              <TextBlock x:Name="AnnouncementsSourceText" FontSize="9" Foreground="Gray" Margin="0,5,0,0"/>
            </StackPanel>
          </Border>
        </Expander>
        <Expander x:Name="PatchingExpander" FontSize="12" IsExpanded="True" Margin="0,2,0,2">
          <Expander.Header>
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="Patching and Updates" VerticalAlignment="Center"/>
              <Button x:Name="PatchingSSAButton" Content="Launch Updates" Margin="10,0,0,0" ToolTip="Launch BigFix Self-Service Application"/>
            </StackPanel>
          </Expander.Header>
          <Border BorderBrush="#00008B" BorderThickness="1" Padding="5" CornerRadius="3" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="PatchingDescriptionText" FontSize="11" TextWrapping="Wrap"/>
              <TextBlock Text="Pending Restart Status:" FontSize="11" FontWeight="Bold" Margin="0,5,0,0"/>
              <TextBlock x:Name="PendingRestartStatusText" FontSize="11" FontWeight="Bold" TextWrapping="Wrap"/>
              <TextBlock x:Name="PatchingUpdatesText" FontSize="11" TextWrapping="Wrap" Margin="0,5,0,0"/>
            </StackPanel>
          </Border>
        </Expander>
        <Expander x:Name="SupportExpander" FontSize="12" IsExpanded="False" Margin="0,2,0,2">
          <Expander.Header>
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="Support" VerticalAlignment="Center"/>
              <Ellipse x:Name="SupportAlertIcon" Width="10" Height="10" Margin="5,0,0,0" Fill="Red" Visibility="Hidden"/>
            </StackPanel>
          </Expander.Header>
          <Border BorderBrush="#00008B" BorderThickness="1" Padding="5" CornerRadius="3" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="SupportText" FontSize="11" TextWrapping="Wrap"/>
              <StackPanel x:Name="SupportLinksPanel" Orientation="Vertical" Margin="0,5,0,0"/>
              <TextBlock x:Name="SupportSourceText" FontSize="9" Foreground="Gray" Margin="0,5,0,0"/>
            </StackPanel>
          </Border>
        </Expander>
        <Expander x:Name="ComplianceExpander" Header="Certificate Status" FontSize="12" IsExpanded="False" Margin="0,2,0,2">
          <Border BorderBrush="#00008B" BorderThickness="1" Padding="5" CornerRadius="3" Background="White" Margin="2">
            <TextBlock x:Name="YubiKeyComplianceText" FontSize="11" TextWrapping="Wrap"/>
          </Border>
        </Expander>
        <TextBlock x:Name="WindowsBuildText" FontSize="11" TextWrapping="Wrap" HorizontalAlignment="Center" Margin="0,10,0,0"/>
        <TextBlock x:Name="ScriptUpdateText" FontSize="11" TextWrapping="Wrap" HorizontalAlignment="Center" Margin="0,10,0,0" Foreground="Red" Visibility="Hidden"/>
      </StackPanel>
    </ScrollViewer>
    <Grid Grid.Row="2" Margin="0,5,0,0">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*" />
            <ColumnDefinition Width="Auto" />
        </Grid.ColumnDefinitions>
        <TextBlock Grid.Column="0" Text="© 2025 Lincoln Laboratory" FontSize="10" Foreground="Gray" HorizontalAlignment="Center" VerticalAlignment="Center"/>
        <Button x:Name="ClearAlertsButton" Grid.Column="1" Content="Clear Alerts" FontSize="10" Padding="5,1" Background="#B0C4DE" ToolTip="Acknowledge all new announcements and support messages."/>
    </Grid>
  </Grid>
</Window>
"@

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
        "YubiKeyComplianceText", "WindowsBuildText", "ClearAlertsButton", "ScriptUpdateText"
    )
    foreach ($elementName in $uiElements) {
        Set-Variable -Name "global:$elementName" -Value $window.FindName($elementName)
    }
    Write-Log "UI elements mapped to variables." -Level "INFO"

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

    $window.Add_Closing({ $_.Cancel = $true; $window.Hide() })
}
catch {
    Handle-Error "Failed to load the XAML layout: $($_.Exception.Message)" -Source "XAML"
    exit
}
# ============================================================
# H) Modularized System Information Functions
# ============================================================
function New-HyperlinkBlock {
    param([string]$Name, [string]$Url)
    $tb = New-Object System.Windows.Controls.TextBlock
    $hp = New-Object System.Windows.Documents.Hyperlink
    $hp.NavigateUri = [Uri]$Url
    $hp.Inlines.Add($Name)
    $hp.Add_RequestNavigate({ try { Start-Process $_.Uri.AbsoluteUri } catch {} })
    $tb.Inlines.Add($hp)
    return $tb
}

function Validate-ContentData {
    param($Data)
    if (-not ($Data.PSObject.Properties.Match('Announcements') -and $Data.PSObject.Properties.Match('Support'))) {
        throw "JSON data is missing 'Announcements' or 'Support' top-level property."
    }
    if (-not $Data.Announcements.PSObject.Properties.Match('Text')) {
        throw "Announcements data is missing 'Text' property."
    }
    if (-not $Data.Support.PSObject.Properties.Match('Text')) {
        throw "Support data is missing 'Text' property."
    }
    return $true
}

function Fetch-ContentData {
    try {
        $url = $config.ContentDataUrl
        Write-Log "Attempting to fetch content from: $url" -Level "INFO"
        
        # Async fetch using job to avoid blocking
        $job = Start-Job -ScriptBlock {
            param($url)
            Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop
        } -ArgumentList $url
        $response = Wait-Job $job | Receive-Job
        Remove-Job $job
        Write-Log "Successfully fetched content from Git repository (Status: $($response.StatusCode))." -Level "INFO"
        
        $contentData = $response.Content | ConvertFrom-Json
        Validate-ContentData -Data $contentData
        Write-Log "Content data validated successfully." -Level "INFO"

        return [PSCustomObject]@{ Data = $contentData; Source = "Remote" }
    }
    catch {
        Write-Log "Failed to fetch or validate content from $($config.ContentDataUrl): $($_.Exception.Message)" -Level "ERROR"
        return [PSCustomObject]@{ Data = $defaultContentData; Source = "Default" }
    }
}

function Get-YubiKeyCertExpiryDays {
    try {
        $ykmanPath = $config.YubiKeyManager_Path
        if ([string]::IsNullOrWhiteSpace($ykmanPath)) {
            throw "The 'YubiKeyManager_Path' is not set in the configuration file."
        }
        if (-not (Test-Path $ykmanPath)) {
            throw "YubiKey Manager executable not found at the configured path: `"$ykmanPath`""
        }
        
        if (-not (& $ykmanPath info 2>$null)) {
            return "YubiKey not present"
        }
        $slots = @("9a", "9c", "9d", "9e")
        $statuses = @()
        foreach ($slot in $slots) {
            $certPem = & $ykmanPath "piv" "certificates" "export" $slot "-" 2>$null
            if ($certPem -and $certPem -match "-----BEGIN CERTIFICATE-----") {
                $tempFile = [System.IO.Path]::GetTempFileName()
                $certPem | Out-File $tempFile -Encoding ASCII
                $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($tempFile)
                Remove-Item $tempFile -Force
                $statuses += "YubiKey Certificate (Slot $slot): Expires: $($cert.NotAfter.ToString("yyyy-MM-dd"))"
            }
        }
        if ($statuses) {
            return $statuses -join "`n"
        }
        return "YubiKey Certificate: No PIV certificate found."
    }
    catch {
        Write-Log "YubiKey check error: $($_.Exception.Message)" -Level "ERROR"
        return "YubiKey Certificate: Unable to determine status."
    }
}

function Get-VirtualSmartCardCertExpiry {
    try {
        $cert = Get-ChildItem "Cert:\CurrentUser\My" | Where-Object { $_.Subject -match "Virtual" } | Sort-Object NotAfter -Descending | Select-Object -First 1
        if (-not $cert) { return "No certificate found." }
        return "Microsoft Virtual Smart Card: Expires: $($cert.NotAfter.ToString("yyyy-MM-dd"))"
    } catch { return "Microsoft Virtual Smart Card: Unable to check status." }
}

function Update-CertificateInfo {
    try {
        Write-Log "Updating certificate info..." -Level "INFO"
        $ykStatus = Get-YubiKeyCertExpiryDays
        $vscStatus = Get-VirtualSmartCardCertExpiry
        $combinedStatus = "$ykStatus`n$vscStatus"
        $global:CachedCertificateStatus = $combinedStatus
        $window.Dispatcher.Invoke({ $global:YubiKeyComplianceText.Text = $combinedStatus })
    } catch { Handle-Error $_.Exception.Message -Source "Update-CertificateInfo" }
}

function Get-PendingRestartStatus {
    $rebootKeys = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\InProgress',
        'HKLM:\SOFTWARE\Wow6432Node\BigFix\EnterpriseClient\BESPendingRestart',
        'HKLM:\SOFTWARE\BigFix\EnterpriseClient\BESPendingRestart'
    )
    $global:PendingRestart = $rebootKeys | ForEach-Object { Test-Path $_ } | Where-Object { $_ } | Select-Object -First 1
    
    if ($global:PendingRestart) { 
        "System restart required." 
    } else { 
        "No system restart required." 
    }
}

$global:LastPatchTimestamp = $null

function Get-WindowsBuildNumber {
    try {
        $buildInfo = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
        $productName = if ($buildInfo.CurrentBuildNumber -ge 22000) { "Windows 11" } else { "Windows 10" }
        return "$productName Build: $($buildInfo.DisplayVersion)"
    } catch { return "Windows Build: Unknown" }
}

function Update-Announcements {
    Write-Log "Updating Announcements section..." -Level "INFO"
    $newAnnouncementsObject = $global:contentData.Data.Announcements
    if (-not $newAnnouncementsObject) { return }

    $newJsonState = $newAnnouncementsObject | ConvertTo-Json -Compress

    $isNew = $false
    if ($config.AnnouncementsLastState -ne $newJsonState) {
        Write-Log "New announcement content detected." -Level "INFO"
        $isNew = $true
    }

    $window.Dispatcher.Invoke({
        if ($isNew) {
            $global:AnnouncementsAlertIcon.Visibility = "Visible"
            if (-not $global:window.IsVisible) {
                $global:TrayIcon.ShowBalloonTip(3000, "New Announcement", $newAnnouncementsObject.Text.Substring(0, [math]::Min(100, $newAnnouncementsObject.Text.Length)), "Info")
            }
        }
        $global:AnnouncementsText.Text = $newAnnouncementsObject.Text
        $global:AnnouncementsDetailsText.Text = $newAnnouncementsObject.Details
        $global:AnnouncementsLinksPanel.Children.Clear()
        if ($newAnnouncementsObject.Links) {
            foreach ($link in $newAnnouncementsObject.Links) {
                $global:AnnouncementsLinksPanel.Children.Add((New-HyperlinkBlock -Name $link.Name -Url $link.Url))
            }
        }
        $global:AnnouncementsSourceText.Text = "Source: $($global:contentData.Source)"
    })
    
    $config.AnnouncementsLastState = $newJsonState
}

function Update-Support {
    Write-Log "Updating Support section..." -Level "INFO"
    $newSupportObject = $global:contentData.Data.Support
    if (-not $newSupportObject) { return }

    $newJsonState = $newSupportObject | ConvertTo-Json -Compress

    $isNew = $false
    if ($config.SupportLastState -ne $newJsonState) {
        Write-Log "New support content detected." -Level "INFO"
        $isNew = $true
    }

    $window.Dispatcher.Invoke({
        if ($isNew) {
            $global:SupportAlertIcon.Visibility = "Visible"
            if (-not $global:window.IsVisible) {
                $global:TrayIcon.ShowBalloonTip(3000, "New Support Update", $newSupportObject.Text.Substring(0, [math]::Min(100, $newSupportObject.Text.Length)), "Info")
            }
        }
        $global:SupportText.Text = $newSupportObject.Text
        $global:SupportLinksPanel.Children.Clear()
        if ($newSupportObject.Links) {
            foreach ($link in $newSupportObject.Links) {
                $global:SupportLinksPanel.Children.Add((New-HyperlinkBlock -Name $link.Name -Url $link.Url))
            }
        }
        $global:SupportSourceText.Text = "Source: $($global:contentData.Source)"
    })
    
    $config.SupportLastState = $newJsonState
}

function Update-PatchingAndSystem {
    Write-Log "Updating Patching and System section..." -Level "INFO"
    $restartStatusText = Get-PendingRestartStatus
    $statusColor = if ($global:PendingRestart) { [System.Windows.Media.Brushes]::Red } else { [System.Windows.Media.Brushes]::Green }
    
    $patchFile = $config.PatchInfoFilePath
    if (Test-Path $patchFile) {
        $fileTimestamp = (Get-Item $patchFile).LastWriteTime
        if ($global:LastPatchTimestamp -ne $fileTimestamp) {
            $rawPatchText = Get-Content -Path $patchFile -Raw -ErrorAction SilentlyContinue
            $global:LastPatchText = if ([string]::IsNullOrWhiteSpace($rawPatchText)) { "No pending updates listed." } else { $rawPatchText }
            $global:LastPatchTimestamp = $fileTimestamp
        }
        $finalPatchText = $global:LastPatchText
    } else {
        $finalPatchText = "Patch info file not found."
    }

    $windowsBuild = Get-WindowsBuildNumber

    $window.Dispatcher.Invoke({
        $global:PatchingDescriptionText.Text = "Lists available software updates. Updates marked (R) require a restart."
        $global:PendingRestartStatusText.Text = $restartStatusText
        $global:PendingRestartStatusText.Foreground = $statusColor
        $global:PatchingUpdatesText.Text = $finalPatchText
        $global:WindowsBuildText.Text = $windowsBuild
    })
}

function Check-ScriptUpdate {
    try {
        $url = $config.ScriptUrl
        Write-Log "Checking for script update from: $url" -Level "INFO"
        
        $job = Start-Job -ScriptBlock {
            param($url)
            Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop
        } -ArgumentList $url
        $response = Wait-Job $job | Receive-Job
        Remove-Job $job
        
        if ($response.Content -match 'ScriptVersion = "(\d+\.\d+\.\d+)"') {
            $remoteVersion = $matches[1]
            if ([version]$remoteVersion -gt [version]$ScriptVersion) {
                Write-Log "New script version available: $remoteVersion" -Level "INFO"
                $window.Dispatcher.Invoke({
                    $global:ScriptUpdateText.Text = "New version $remoteVersion available. Please update the script."
                    $global:ScriptUpdateText.Visibility = "Visible"
                    if (-not $global:window.IsVisible) {
                        $global:TrayIcon.ShowBalloonTip(5000, "Script Update Available", "Version $remoteVersion is available.", "Info")
                    }
                })
                return $true
            }
        }
        Write-Log "No new script version available." -Level "INFO"
        $window.Dispatcher.Invoke({
            $global:ScriptUpdateText.Visibility = "Hidden"
        })
        return $false
    } catch {
        Write-Log "Failed to check for script update: $($_.Exception.Message)" -Level "WARNING"
        return $false
    }
}

# ============================================================
# I) Tray Icon Management
# ============================================================

$global:BlinkingTimer = $null
$global:MainIcon = $null
$global:WarningIcon = $null

function Get-Icon {
    param([string]$Path)
    if (Test-Path $Path) {
        try {
            return New-Object System.Drawing.Icon($Path)
        }
        catch {
            Write-Log "Error loading icon from `"$Path`": $($_.Exception.Message)" -Level "ERROR"
        }
    }
    return [System.Drawing.SystemIcons]::Application
}

function Update-TrayIcon {
    if (-not $global:TrayIcon.Visible) { return }
    
    $announcementAlert = $global:AnnouncementsAlertIcon.Visibility -eq "Visible"
    $supportAlert = $global:SupportAlertIcon.Visibility -eq "Visible"
    
    $hasBlinkingAlert = $announcementAlert -or $supportAlert
    $hasAnyAlert = $global:PendingRestart -or $hasBlinkingAlert

    if ($hasBlinkingAlert -and -not $window.IsVisible) {
        if ($config.BlinkingEnabled) {
            if (-not $global:BlinkingTimer.IsEnabled) {
                $global:TrayIcon.Icon = $global:WarningIcon
                $global:BlinkingTimer.Start()
            }
        } else {
            $global:TrayIcon.Icon = $global:WarningIcon
            $global:TrayIcon.Text = "LLNOTIFY v$ScriptVersion - Alerts Pending"
        }
    }
    else {
        if ($global:BlinkingTimer.IsEnabled) {
            $global:BlinkingTimer.Stop()
        }
        $global:TrayIcon.Icon = if ($hasAnyAlert) { $global:WarningIcon } else { $global:MainIcon }
    }
}

function Initialize-TrayIcon {
    if (-not $global:FormsAvailable) { return }
    try {
        $global:MainIcon = Get-Icon -Path $config.IconPaths.Main
        $global:WarningIcon = Get-Icon -Path $config.IconPaths.Warning

        $global:TrayIcon = New-Object System.Windows.Forms.NotifyIcon
        $global:TrayIcon.Icon = $global:MainIcon
        $global:TrayIcon.Text = "LLNOTIFY v$ScriptVersion"
        $global:TrayIcon.Visible = $true

        $ContextMenuStrip = New-Object System.Windows.Forms.ContextMenuStrip
        $ContextMenuStrip.Items.AddRange(@(
            (New-Object System.Windows.Forms.ToolStripMenuItem("Show Dashboard", $null, { Toggle-WindowVisibility })),
            (New-Object System.Windows.Forms.ToolStripMenuItem("Refresh Now", $null, { Main-UpdateCycle -ForceCertificateCheck $true })),
            (New-Object System.Windows.Forms.ToolStripMenuItem("Exit", $null, { $window.Dispatcher.InvokeShutdown() }))
        ))
        $global:TrayIcon.ContextMenuStrip = $ContextMenuStrip
        $global:TrayIcon.add_MouseClick({ if ($_.Button -eq 'Left') { Toggle-WindowVisibility } })
    } catch { Handle-Error $_.Exception.Message -Source "Initialize-TrayIcon" }
}

# ============================================================
# K) Window Visibility Management
# ============================================================
function Set-WindowPosition {
    Add-Type -AssemblyName System.Windows.Forms
    $mousePos = [System.Windows.Forms.Cursor]::Position
    $screen = [System.Windows.Forms.Screen]::AllScreens | Where-Object { $_.Bounds.Contains($mousePos) } | Select-Object -First 1
    if (-not $screen) { $screen = [System.Windows.Forms.Screen]::PrimaryScreen }
    $window.Left = $screen.WorkingArea.X + ($screen.WorkingArea.Width - $window.ActualWidth) / 2
    $window.Top = $screen.WorkingArea.Y + ($screen.WorkingArea.Height - $window.ActualHeight) / 2
}

function Toggle-WindowVisibility {
    $window.Dispatcher.Invoke({
        if ($window.IsVisible) {
            $window.Hide()
            Update-TrayIcon
        } else {
            $window.Show()
            $global:BlinkingTimer.Stop()
            Update-TrayIcon
            Set-WindowPosition
            $window.Activate()
            $window.Topmost = $true; $window.Topmost = $false
        }
    })
}
# ============================================================
# O) Main Update Cycle and DispatcherTimer
# ============================================================
function Main-UpdateCycle {
    param([bool]$ForceCertificateCheck = $false)
    try {
        Write-Log "Main update cycle running..." -Level INFO
        $global:contentData = Fetch-ContentData
        
        Update-Announcements
        Update-Support
        Update-PatchingAndSystem
        
        if ($ForceCertificateCheck -or (-not $global:LastCertificateCheck -or ((Get-Date) - $global:LastCertificateCheck).TotalSeconds -ge $config.CertificateCheckInterval)) {
            Update-CertificateInfo
            $global:LastCertificateCheck = Get-Date
        }
        
        Check-ScriptUpdate
        
        Update-TrayIcon
        Save-Configuration -Config $config
        Rotate-LogFile
    }
    catch { Handle-Error $_.Exception.Message -Source "Main-UpdateCycle" }
}

# ============================================================
# P) Initial Setup & Application Start
# ============================================================
try {
    $global:blinkingTickAction = {
        if ($global:TrayIcon.Icon.Handle -eq $global:WarningIcon.Handle) {
            $global:TrayIcon.Icon = $global:MainIcon
        } else {
            $global:TrayIcon.Icon = $global:WarningIcon
        }
    }
    $global:BlinkingTimer = New-Object System.Windows.Threading.DispatcherTimer
    $global:BlinkingTimer.Interval = [TimeSpan]::FromSeconds(1)
    $global:BlinkingTimer.add_Tick($global:blinkingTickAction)

    $global:mainTickAction = {
        param($sender, $e)
        Main-UpdateCycle
    }
    $global:DispatcherTimer = New-Object System.Windows.Threading.DispatcherTimer
    $global:DispatcherTimer.Interval = [TimeSpan]::FromSeconds($config.RefreshInterval)
    $global:DispatcherTimer.add_Tick($global:mainTickAction)

    Initialize-TrayIcon
    Log-DotNetVersion
    Main-UpdateCycle -ForceCertificateCheck $true
    
    $global:DispatcherTimer.Start()
    Write-Log "Main timer started." -Level INFO
    
    $window.Dispatcher.Add_UnhandledException({ Handle-Error $_.Exception.Message -Source "Dispatcher"; $_.Handled = $true })

    Write-Log "Application startup complete. Running dispatcher." -Level "INFO"
    [System.Windows.Threading.Dispatcher]::Run()
}
catch {
    Handle-Error "A critical error occurred during startup: $($_.Exception.Message)" -Source "Startup"
}
finally {
    Write-Log "--- LLNOTIFY Script Exiting ---"
    if ($global:DispatcherTimer) { $global:DispatcherTimer.Stop() }
    if ($global:TrayIcon) { $global:TrayIcon.Dispose() }
    if ($global:MainIcon) { $global:MainIcon.Dispose() }
    if ($global:WarningIcon) { $global:WarningIcon.Dispose() }
}
