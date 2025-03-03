###############################################################################
# SystemMonitor.ps1 - Renamed to SHOT (System Health Observation Tool)
###############################################################################

# Ensure $PSScriptRoot is defined for older versions
if (-not $PSScriptRoot) {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    $ScriptDir = $PSScriptRoot
}

# Define version
$ScriptVersion = "1.1.0"

# ========================
# A) Advanced Logging & Error Handling
# ========================
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )
    $logPath = if ($LogFilePath) { $LogFilePath } else { Join-Path $ScriptDir "SHOT.log" }  # Updated log name
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Add-Content -Path $logPath -Value $logEntry -ErrorAction SilentlyContinue
}

function Handle-Error {
    param(
        [string]$ErrorMessage,
        [string]$Source = ""
    )
    if ($Source) { $ErrorMessage = "[$Source] $ErrorMessage" }
    Write-Log $ErrorMessage -Level "ERROR"
}

function Log-DotNetVersion {
    try {
        $dotNetVersion = [System.Environment]::Version.ToString()
        Write-Log ".NET Version: $dotNetVersion" -Level "INFO"
        try {
            $frameworkDescription = [System.Runtime.InteropServices.RuntimeInformation]::FrameworkDescription
            Write-Log ".NET Framework Description: $frameworkDescription" -Level "INFO"
        }
        catch {
            Write-Log "RuntimeInformation not available." -Level "WARNING"
        }
    }
    catch {
        Write-Log "Error capturing .NET version: $_" -Level "ERROR"
    }
}

# ========================
# B) External Configuration
# ========================
$configPath = Join-Path $ScriptDir "SHOT.config.json"  # Updated config name
$LogFilePath = Join-Path $ScriptDir "SHOT.log"  # Updated log name

$defaultConfig = @{
    RefreshInterval    = 30
    LogRotationSizeMB  = 5
    DefaultLogLevel    = "INFO"
    ContentDataUrl     = "ContentData.json"
    YubiKeyAlertDays   = 7
    IconPaths          = @{
        Main       = "icon.ico"
        Warning    = "warning.ico"
    }
    YubiKeyLastCheck   = @{
        Date   = "1970-01-01 00:00:00"
        Result = "YubiKey Certificate: Not yet checked"
    }
    AnnouncementsLastState = @{}
    Version            = $ScriptVersion
}

if (Test-Path $configPath) {
    try {
        $config = Get-Content $configPath -Raw | ConvertFrom-Json
        Write-Log "Loaded config from $configPath" -Level "INFO"
        foreach ($key in $defaultConfig.Keys) {
            if (-not $config.PSObject.Properties.Match($key)) {
                $config | Add-Member -NotePropertyName $key -NotePropertyValue $defaultConfig[$key]
            }
        }
        if ($config.YubiKeyLastCheck.Result.PSObject.Properties.Name -contains "value") {
            $config.YubiKeyLastCheck.Result = $config.YubiKeyLastCheck.Result.value
        }
        if (-not $config.AnnouncementsLastState) {
            $config.AnnouncementsLastState = @{}
        }
    }
    catch {
        Write-Host "Error reading configuration file. Using default settings with preserved URL if possible."
        Write-Log "Error reading config: $_" -Level "ERROR"
        $config = $defaultConfig
        try {
            $rawContent = Get-Content $configPath -Raw
            if ($rawContent -match '"ContentDataUrl"\s*:\s*"([^"]+)"') {
                $config.ContentDataUrl = $matches[1]
                Write-Log "Preserved ContentDataUrl: $($config.ContentDataUrl)" -Level "INFO"
            }
        }
        catch {
            Write-Log "Could not preserve ContentDataUrl from corrupted config: $_" -Level "WARNING"
        }
    }
}
else {
    $config = $defaultConfig
    $config | ConvertTo-Json -Depth 3 | Out-File $configPath -Force
    Write-Log "Created default config at $configPath" -Level "INFO"
}

# Construct full icon paths dynamically
$mainIconPath = Join-Path $ScriptDir $config.IconPaths.Main
$warningIconPath = Join-Path $ScriptDir $config.IconPaths.Warning
$mainIconUri = "file:///" + ($mainIconPath -replace '\\','/')

# Default content data in case fetch fails
$defaultContentData = @{
    Announcements = @{
        Text    = "No announcements at this time."
        Details = "Check back later for updates."
        Links   = @{
            Link1 = "https://company.com/news1"
            Link2 = "https://company.com/news2"
        }
    }
    EarlyAdopter = @{
        Text  = "Join our beta program!"
        Links = @{
            Link1 = "https://beta.company.com/signup"
            Link2 = "https://beta.company.com/info"
        }
    }
    Support = @{
        Text  = "Contact IT Support: support@company.com | Phone: 1-800-555-1234"
        Links = @{
            Link1 = "https://support.company.com/help"
            Link2 = "https://support.company.com/tickets"
        }
    }
}

# ========================
# C) Log File Setup & Rotation
# ========================
$LogDirectory = Split-Path $LogFilePath
if (-not (Test-Path $LogDirectory)) {
    New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
    Write-Log "Created log directory: $LogDirectory" -Level "INFO"
}

function Rotate-LogFile {
    try {
        if (Test-Path $LogFilePath) {
            $fileInfo = Get-Item $LogFilePath
            $maxSizeBytes = $config.LogRotationSizeMB * 1MB
            if ($fileInfo.Length -gt $maxSizeBytes) {
                $archivePath = "$LogFilePath.$(Get-Date -Format 'yyyyMMddHHmmss').archive"
                Rename-Item -Path $LogFilePath -NewName $archivePath
                Write-Log "Log file rotated. Archived as $archivePath" -Level "INFO"
            }
        }
    }
    catch {
        Write-Log "Failed to rotate log file: $_" -Level "ERROR"
    }
}
Rotate-LogFile

# Update Write-Log to use $config.DefaultLogLevel
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = $config.DefaultLogLevel
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Add-Content -Path $LogFilePath -Value $logEntry -ErrorAction SilentlyContinue
}

# ========================
# D) Import Required Assemblies
# ========================
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ========================
# E) XAML Layout Definition
# ========================
$xamlString = @"
<?xml version="1.0" encoding="utf-8"?>
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="SHOT v$ScriptVersion"
    WindowStartupLocation="Manual"
    SizeToContent="WidthAndHeight"
    MinWidth="350" MinHeight="500"
    ResizeMode="CanResize"
    ShowInTaskbar="False"
    Visibility="Hidden"
    Topmost="True"
    Background="#f0f0f0">
  <Grid Margin="3">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <!-- Title Section -->
    <Border Grid.Row="0" Background="#0078D7" Padding="4" CornerRadius="2" Margin="0,0,0,4">
      <StackPanel Orientation="Horizontal" VerticalAlignment="Center" HorizontalAlignment="Center">
        <Image Source="$mainIconUri" Width="20" Height="20" Margin="0,0,4,0"/>
        <TextBlock Text="System Health Observation Tool"
                   FontSize="14" FontWeight="Bold" Foreground="White"
                   VerticalAlignment="Center"/>
      </StackPanel>
    </Border>
    <!-- Content Area -->
    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
      <StackPanel VerticalAlignment="Top">
        <!-- Information Section -->
        <Expander Header="Information" ToolTip="View system details" FontSize="12" Foreground="#0078D7" IsExpanded="True" Margin="0,2,0,2">
          <Border BorderBrush="#0078D7" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="LoggedOnUserText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
              <TextBlock x:Name="MachineTypeText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
              <TextBlock x:Name="OSVersionText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
              <TextBlock x:Name="SystemUptimeText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
              <TextBlock x:Name="UsedDiskSpaceText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
              <TextBlock x:Name="IpAddressText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
              <TextBlock x:Name="YubiKeyCertExpiryText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
            </StackPanel>
          </Border>
        </Expander>
        <!-- Announcements with Red Dot Alert -->
        <Expander x:Name="AnnouncementsExpander" ToolTip="View latest announcements" FontSize="12" Foreground="#000080" IsExpanded="False" Margin="0,2,0,2">
          <Expander.Header>
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="Announcements" VerticalAlignment="Center"/>
              <Ellipse x:Name="AnnouncementsAlertIcon" Width="10" Height="10" Margin="4,0,0,0" Fill="Red" Visibility="Hidden"/>
            </StackPanel>
          </Expander.Header>
          <Border BorderBrush="#000080" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="AnnouncementsText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
              <TextBlock x:Name="AnnouncementsDetailsText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
              <TextBlock FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300">
                <Hyperlink x:Name="AnnouncementsLink1" NavigateUri="https://company.com/news1" ToolTip="Open announcement link 1">Announcement Link 1</Hyperlink>
              </TextBlock>
              <TextBlock FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300">
                <Hyperlink x:Name="AnnouncementsLink2" NavigateUri="https://company.com/news2" ToolTip="Open announcement link 2">Announcement Link 2</Hyperlink>
              </TextBlock>
            </StackPanel>
          </Border>
        </Expander>
        <!-- Patching and Updates -->
        <Expander Header="Patching and Updates" ToolTip="View patching status" FontSize="12" Foreground="#008000" IsExpanded="False" Margin="0,2,0,2">
          <Border BorderBrush="#008000" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="PatchingUpdatesText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
            </StackPanel>
          </Border>
        </Expander>
        <!-- Support -->
        <Expander Header="Support" ToolTip="Contact IT support" FontSize="12" Foreground="#800000" IsExpanded="False" Margin="0,2,0,2">
          <Border BorderBrush="#800000" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="SupportText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
              <TextBlock FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300">
                <Hyperlink x:Name="SupportLink1" NavigateUri="https://support.company.com/help" ToolTip="Visit support page">Support Link 1</Hyperlink>
              </TextBlock>
              <TextBlock FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300">
                <Hyperlink x:Name="SupportLink2" NavigateUri="https://support.company.com/tickets" ToolTip="Submit a ticket">Support Link 2</Hyperlink>
              </TextBlock>
            </StackPanel>
          </Border>
        </Expander>
        <!-- Open Early Adopter Testing -->
        <Expander Header="Open Early Adopter Testing" ToolTip="Join beta program" FontSize="12" Foreground="#FF00FF" IsExpanded="False" Margin="0,2,0,2">
          <Border BorderBrush="#FF00FF" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="EarlyAdopterText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
              <TextBlock FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300">
                <Hyperlink x:Name="EarlyAdopterLink1" NavigateUri="https://beta.company.com/signup" ToolTip="Sign up for beta">Early Adopter Link 1</Hyperlink>
              </TextBlock>
              <TextBlock FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300">
                <Hyperlink x:Name="EarlyAdopterLink2" NavigateUri="https://beta.company.com/info" ToolTip="Learn more about beta">Early Adopter Link 2</Hyperlink>
              </TextBlock>
            </StackPanel>
          </Border>
        </Expander>
        <!-- Compliance Section with Status Indicator -->
        <Expander x:Name="ComplianceExpander" ToolTip="Check compliance status" FontSize="12" Foreground="#B22222" IsExpanded="False" Margin="0,2,0,2">
          <Expander.Header>
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="Compliance" VerticalAlignment="Center"/>
              <Ellipse x:Name="ComplianceStatusIndicator" Width="10" Height="10" Margin="4,0,0,0" Fill="Gray" Visibility="Visible"/>
            </StackPanel>
          </Expander.Header>
          <Border BorderBrush="#B22222" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
            <StackPanel>
              <Border BorderBrush="#28a745" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
                <StackPanel>
                  <TextBlock Text="Antivirus Status" FontSize="11" FontWeight="Bold" Margin="2" Foreground="#28a745"/>
                  <TextBlock x:Name="AntivirusStatusText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
                </StackPanel>
              </Border>
              <Border x:Name="BitLockerBorder" BorderBrush="#6c757d" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
                <StackPanel>
                  <TextBlock Text="BitLocker Status" FontSize="11" FontWeight="Bold" Margin="2" Foreground="#6c757d"/>
                  <TextBlock x:Name="BitLockerStatusText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
                </StackPanel>
              </Border>
              <Border x:Name="BigFixBorder" BorderBrush="#4b0082" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
                <StackPanel>
                  <TextBlock Text="BigFix (BESClient) Status" FontSize="11" FontWeight="Bold" Margin="2" Foreground="#4b0082"/>
                  <TextBlock x:Name="BigFixStatusText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
                </StackPanel>
              </Border>
              <Border x:Name="Code42Border" BorderBrush="#800080" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
                <StackPanel>
                  <TextBlock Text="Code42 Service Status" FontSize="11" FontWeight="Bold" Margin="2" Foreground="#800080"/>
                  <TextBlock x:Name="Code42StatusText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
                </StackPanel>
              </Border>
              <Border x:Name="FIPSBorder" BorderBrush="#FF4500" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
                <StackPanel>
                  <TextBlock Text="FIPS Compliance Status" FontSize="11" FontWeight="Bold" Margin="2" Foreground="#FF4500"/>
                  <TextBlock x:Name="FIPSStatusText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
                </StackPanel>
              </Border>
            </StackPanel>
          </Border>
        </Expander>
        <!-- Logs Section -->
        <Expander Header="Logs" ToolTip="View recent logs" FontSize="12" Foreground="#ff8c00" IsExpanded="False" Margin="0,2,0,2">
          <Border BorderBrush="#ff8c00" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
            <StackPanel>
              <ListView x:Name="LogListView" FontSize="10" Margin="2" Height="120">
                <ListView.View>
                  <GridView>
                    <GridViewColumn Header="Timestamp" Width="100" DisplayMemberBinding="{Binding Timestamp}" />
                    <GridViewColumn Header="Message" Width="150" DisplayMemberBinding="{Binding Message}" />
                  </GridView>
                </ListView.View>
              </ListView>
              <Button x:Name="ExportLogsButton" Content="Export Logs" Width="80" Margin="2" HorizontalAlignment="Right" ToolTip="Save logs to a file"/>
            </StackPanel>
          </Border>
        </Expander>
        <!-- About Section -->
        <Expander Header="About" ToolTip="View app info and changelog" FontSize="12" Foreground="#000000" IsExpanded="False" Margin="0,2,0,2">
          <Border BorderBrush="#000000" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="AboutText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300">
                <TextBlock.Text><![CDATA[SHOT v1.1.0
© 2025 SHOT. All rights reserved.
Built with PowerShell and WPF.

Changelog:
- v1.1.0: Added tooltips, collapsible tray menu, status indicators, YubiKey alerts, async updates, versioning
- v1.0.0: Initial release]]></TextBlock.Text>
              </TextBlock>
            </StackPanel>
          </Border>
        </Expander>
      </StackPanel>
    </ScrollViewer>
    <!-- Footer Section -->
    <TextBlock Grid.Row="2" Text="© 2025 SHOT" FontSize="10" Foreground="Gray" HorizontalAlignment="Center" Margin="0,4,0,0"/>
  </Grid>
</Window>
"@

# Convert string to XmlDocument
$xmlDoc = New-Object System.Xml.XmlDocument
$xmlDoc.LoadXml($xamlString)

# Create XmlNodeReader from XmlDocument
$reader = New-Object System.Xml.XmlNodeReader $xmlDoc

# ========================
# F) Load and Verify XAML
# ========================
try {
    $window = [Windows.Markup.XamlReader]::Load($reader)
    Write-Log "XAML loaded successfully" -Level "INFO"
}
catch {
    Handle-Error "Failed to load the XAML layout: $_" -Source "XAML"
    exit
}
if ($window -eq $null) {
    Handle-Error "Failed to load the XAML layout. Check the XAML syntax for errors." -Source "XAML"
    exit
}

# ========================
# G) Access UI Elements
# ========================
$LoggedOnUserText    = $window.FindName("LoggedOnUserText")
$MachineTypeText     = $window.FindName("MachineTypeText")
$OSVersionText       = $window.FindName("OSVersionText")
$SystemUptimeText    = $window.FindName("SystemUptimeText")
$UsedDiskSpaceText   = $window.FindName("UsedDiskSpaceText")
$IpAddressText       = $window.FindName("IpAddressText")
$YubiKeyCertExpiryText = $window.FindName("YubiKeyCertExpiryText")

$AntivirusStatusText = $window.FindName("AntivirusStatusText")
$BitLockerStatusText = $window.FindName("BitLockerStatusText")
$BigFixStatusText    = $window.FindName("BigFixStatusText")
$Code42StatusText    = $window.FindName("Code42StatusText")
$FIPSStatusText      = $window.FindName("FIPSStatusText")
$AboutText           = $window.FindName("AboutText")

$AnnouncementsExpander = $window.FindName("AnnouncementsExpander")
$AnnouncementsAlertIcon = $window.FindName("AnnouncementsAlertIcon")
$AnnouncementsText   = $window.FindName("AnnouncementsText")
$AnnouncementsDetailsText = $window.FindName("AnnouncementsDetailsText")
$AnnouncementsLink1  = $window.FindName("AnnouncementsLink1")
$AnnouncementsLink2  = $window.FindName("AnnouncementsLink2")
$PatchingUpdatesText = $window.FindName("PatchingUpdatesText")
$SupportText         = $window.FindName("SupportText")
$SupportLink1        = $window.FindName("SupportLink1")
$SupportLink2        = $window.FindName("SupportLink2")
$EarlyAdopterText    = $window.FindName("EarlyAdopterText")
$EarlyAdopterLink1   = $window.FindName("EarlyAdopterLink1")
$EarlyAdopterLink2   = $window.FindName("EarlyAdopterLink2")

$ComplianceExpander  = $window.FindName("ComplianceExpander")
$ComplianceStatusIndicator = $window.FindName("ComplianceStatusIndicator")
$LogListView         = $window.FindName("LogListView")
$ExportLogsButton    = $window.FindName("ExportLogsButton")

$BitLockerBorder     = $window.FindName("BitLockerBorder")
$BigFixBorder        = $window.FindName("BigFixBorder")
$Code42Border        = $window.FindName("Code42Border")
$FIPSBorder          = $window.FindName("FIPSBorder")

# Global variables
$yubiKeyJob = $null
$contentData = $null
$announcementAlertActive = $false
$yubiKeyAlertShown = $false

# ========================
# H) Modularized System Information Functions
# ========================
function Fetch-ContentData {
    try {
        $url = $config.ContentDataUrl
        Write-Log "Attempting to fetch content from: $url" -Level "INFO"
        if ($url -match "^(http|https)://") {
            $response = Invoke-WebRequest -Uri $url -UseBasicParsing
            $contentData = $response.Content | ConvertFrom-Json
            Write-Log "Fetched content data from URL: $url" -Level "INFO"
            Write-Log "Raw content: $($response.Content)" -Level "INFO"
        }
        elseif ($url -match "^\\\\") {
            if (-not (Test-Path $url)) {
                throw "Network path not accessible: $url"
            }
            $rawContent = Get-Content -Path $url -Raw
            $contentData = $rawContent | ConvertFrom-Json
            Write-Log "Fetched content data from network path: $url" -Level "INFO"
            Write-Log "Raw content: $rawContent" -Level "INFO"
        }
        else {
            $fullPath = if ([System.IO.Path]::IsPathRooted($url)) { $url } else { Join-Path $ScriptDir $url }
            if (-not (Test-Path $fullPath)) {
                throw "Local path not found: $fullPath"
            }
            $rawContent = Get-Content -Path $fullPath -Raw
            $contentData = $rawContent | ConvertFrom-Json
            Write-Log "Fetched content data from local path: $fullPath" -Level "INFO"
            Write-Log "Raw content: $rawContent" -Level "INFO"
        }
        Write-Log "Parsed content data: $($contentData | ConvertTo-Json -Depth 3)" -Level "INFO"
        return $contentData
    }
    catch {
        Write-Log "Failed to fetch content data from ${url}: $_" -Level "ERROR"
        Write-Log "Reverting to default content data" -Level "WARNING"
        return $defaultContentData
    }
}

function Get-YubiKeyCertExpiryDays {
    param([string]$ykmanPathPassed = "C:\Program Files\Yubico\Yubikey Manager\ykman.exe")
    try {
        if (-not (Test-Path $ykmanPathPassed)) {
            throw "ykman.exe not found at $ykmanPathPassed"
        }
        Write-Log "ykman.exe found at $ykmanPathPassed" -Level "INFO"

        $yubiKeyInfo = & $ykmanPathPassed info 2>$null
        if (-not $yubiKeyInfo) {
            Write-Log "No YubiKey detected" -Level "INFO"
            return "YubiKey not present"
        }
        Write-Log "YubiKey detected: $yubiKeyInfo" -Level "INFO"

        $pivInfo = & $ykmanPathPassed "piv" "info" 2>$null
        if ($pivInfo) {
            Write-Log "PIV info: $pivInfo" -Level "INFO"
        } else {
            Write-Log "No PIV info available" -Level "WARNING"
        }

        $slots = @("9a", "9c", "9d", "9e")
        $certPem = $null
        $slotUsed = $null

        foreach ($slot in $slots) {
            Write-Log "Checking slot $slot for certificate" -Level "INFO"
            $certPem = & $ykmanPathPassed "piv" "certificates" "export" $slot "-" 2>$null
            if ($certPem -and $certPem -match "-----BEGIN CERTIFICATE-----") {
                $slotUsed = $slot
                Write-Log "Certificate found in slot $slot" -Level "INFO"
                break
            } else {
                Write-Log "No valid certificate in slot $slot" -Level "INFO"
            }
        }

        if (-not $certPem) {
            throw "No certificate found in slots 9a, 9c, 9d, or 9e"
        }

        $tempFile = [System.IO.Path]::GetTempFileName()
        $certPem | Out-File $tempFile -Encoding ASCII

        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $cert.Import($tempFile)

        $today = Get-Date
        $expiryDate = $cert.NotAfter
        $daysUntilExpiry = ($expiryDate - $today).Days

        Remove-Item $tempFile -Force

        if ($daysUntilExpiry -lt 0) {
            return "YubiKey Certificate (Slot $slotUsed): Expired ($(-$daysUntilExpiry) days ago)"
        } else {
            return "YubiKey Certificate (Slot $slotUsed): $daysUntilExpiry days until expiry ($expiryDate)"
        }
    }
    catch {
        if ($_.Exception.Message -ne "No YubiKey detected by ykman") {
            Write-Log "Error retrieving YubiKey certificate expiry: $_" -Level "ERROR"
            return "YubiKey Certificate: Unable to determine expiry date - $_"
        }
    }
}

function Start-YubiKeyCertCheckAsync {
    if ($global:yubiKeyJob -and $global:yubiKeyJob.State -eq "Running") {
        Write-Log "YubiKey certificate check already in progress." -Level "INFO"
        return
    }

    $global:yubiKeyJob = Start-Job -ScriptBlock {
        param($ykmanPath, $LogFilePathPass)

        function Write-Log {
            param(
                [string]$Message,
                [ValidateSet("INFO", "WARNING", "ERROR")]
                [string]$Level = "INFO"
            )
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logEntry = "[$timestamp] [$Level] $Message"
            Add-Content -Path $LogFilePathPass -Value $logEntry -ErrorAction SilentlyContinue
        }

        try {
            if (-not (Test-Path $ykmanPath)) {
                throw "ykman.exe not found at $ykmanPath"
            }
            Write-Log "ykman.exe found at $ykmanPath" -Level "INFO"

            $yubiKeyInfo = & $ykmanPath info 2>$null
            if (-not $yubiKeyInfo) {
                Write-Log "No YubiKey detected" -Level "INFO"
                return "YubiKey not present"
            }
            Write-Log "YubiKey detected: $yubiKeyInfo" -Level "INFO"

            $pivInfo = & $ykmanPath "piv" "info" 2>$null
            if ($pivInfo) {
                Write-Log "PIV info: $pivInfo" -Level "INFO"
            } else {
                Write-Log "No PIV info available" -Level "WARNING"
            }

            $slots = @("9a", "9c", "9d", "9e")
            $certPem = $null
            $slotUsed = $null

            foreach ($slot in $slots) {
                Write-Log "Checking slot $slot for certificate" -Level "INFO"
                $certPem = & $ykmanPath "piv" "certificates" "export" $slot "-" 2>$null
                if ($certPem -and $certPem -match "-----BEGIN CERTIFICATE-----") {
                    $slotUsed = $slot
                    Write-Log "Certificate found in slot $slot" -Level "INFO"
                    break
                } else {
                    Write-Log "No valid certificate in slot $slot" -Level "INFO"
                }
            }

            if (-not $certPem) {
                throw "No certificate found in slots 9a, 9c, 9d, or 9e"
            }

            $tempFile = [System.IO.Path]::GetTempFileName()
            $certPem | Out-File $tempFile -Encoding ASCII

            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
            $cert.Import($tempFile)

            $today = Get-Date
            $expiryDate = $cert.NotAfter
            $daysUntilExpiry = ($expiryDate - $today).Days

            Remove-Item $tempFile -Force

            if ($daysUntilExpiry -lt 0) {
                return "YubiKey Certificate (Slot $slotUsed): Expired ($(-$daysUntilExpiry) days ago)"
            } else {
                return "YubiKey Certificate (Slot $slotUsed): $daysUntilExpiry days until expiry ($expiryDate)"
            }
        }
        catch {
            if ($_.Exception.Message -ne "No YubiKey detected by ykman") {
                Write-Log "Error retrieving YubiKey certificate expiry: $_" -Level "ERROR"
                return "YubiKey Certificate: Unable to determine expiry date - $_"
            }
        }
    } -ArgumentList "C:\Program Files\Yubico\Yubikey Manager\ykman.exe", $LogFilePath

    Write-Log "Started async YubiKey certificate check job." -Level "INFO"
}

function Update-SystemInfo {
    try {
        $window.Dispatcher.Invoke([Action]{ $LoggedOnUserText.Text = "Logged-in User: Checking..." })
        $user = [System.Environment]::UserName
        $window.Dispatcher.Invoke([Action]{ $LoggedOnUserText.Text = "Logged-in User: $user" })
        Write-Log "Logged-in User: $user" -Level "INFO"

        $machine = Get-CimInstance -ClassName Win32_ComputerSystem
        $machineType = "$($machine.Manufacturer) $($machine.Model)"
        $window.Dispatcher.Invoke([Action]{ $MachineTypeText.Text = "Machine Type: $machineType" })
        Write-Log "Machine Type: $machineType" -Level "INFO"

        $os = Get-CimInstance -ClassName Win32_OperatingSystem
        $osVersion = "$($os.Caption) (Build $($os.BuildNumber))"
        try {
            $displayVersion = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name 'DisplayVersion' -ErrorAction SilentlyContinue).DisplayVersion
            if ($displayVersion) { $osVersion += " $displayVersion" }
        }
        catch {
            Handle-Error "Could not retrieve DisplayVersion from registry: $_" -Source "OSVersion"
        }
        $window.Dispatcher.Invoke([Action]{ $OSVersionText.Text = "OS Version: $osVersion" })
        Write-Log "OS Version: $osVersion" -Level "INFO"

        $uptime = (Get-Date) - $os.LastBootUpTime
        $systemUptime = "$([math]::Floor($uptime.TotalDays)) days $($uptime.Hours) hours"
        $window.Dispatcher.Invoke([Action]{ $SystemUptimeText.Text = "System Uptime: $systemUptime" })
        Write-Log "System Uptime: $systemUptime" -Level "INFO"

        $drive = Get-PSDrive -Name C
        $usedDiskSpace = "$([math]::Round(($drive.Used / 1GB), 2)) GB of $([math]::Round((($drive.Free + $drive.Used) / 1GB), 2)) GB"
        $window.Dispatcher.Invoke([Action]{ $UsedDiskSpaceText.Text = "Used Disk Space: $usedDiskSpace" })
        Write-Log "Used Disk Space: $usedDiskSpace" -Level "INFO"

        $ipv4s = Get-NetIPAddress -AddressFamily IPv4 | Where-Object {
            $_.IPAddress -notlike "127.*" -and $_.IPAddress -notlike "169.254.*" -and
            $_.IPAddress -notin @("0.0.0.0","255.255.255.255") -and $_.PrefixOrigin -ne "WellKnown"
        } | Select-Object -ExpandProperty IPAddress -ErrorAction SilentlyContinue

        if ($ipv4s) {
            $ipList = $ipv4s -join ", "
            $window.Dispatcher.Invoke([Action]{ $IpAddressText.Text = "IPv4 Address(es): $ipList" })
            Write-Log "IP Address(es): $ipList" -Level "INFO"
        }
        else {
            $window.Dispatcher.Invoke([Action]{ $IpAddressText.Text = "IPv4 Address(es): None detected" })
            Write-Log "No valid IPv4 addresses found." -Level "WARNING"
        }

        if ($global:yubiKeyJob -and $global:yubiKeyJob.State -eq "Completed") {
            $yubiKeyResult = Receive-Job -Job $global:yubiKeyJob
            $yubiKeyResultString = if ($yubiKeyResult -is [string]) { $yubiKeyResult } else { $yubiKeyResult.ToString() }
            $window.Dispatcher.Invoke([Action]{ $YubiKeyCertExpiryText.Text = $yubiKeyResultString })
            $config.YubiKeyLastCheck.Date = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            $config.YubiKeyLastCheck.Result = $yubiKeyResultString
            $config | ConvertTo-Json -Depth 3 | Out-File $configPath -Force
            Remove-Job -Job $global:yubiKeyJob -Force
            $global:yubiKeyJob = $null
            Write-Log "YubiKey certificate check completed and saved: $yubiKeyResultString" -Level "INFO"

            if ($yubiKeyResultString -match "(\d+) days until expiry" -and [int]$matches[1] -le $config.YubiKeyAlertDays -and -not $global:yubiKeyAlertShown) {
                $days = [int]$matches[1]
                $TrayIcon.ShowBalloonTip(5000, "YubiKey Expiry Alert", "YubiKey certificate expires in $days days!", [System.Windows.Forms.ToolTipIcon]::Warning)
                Write-Log "YubiKey expiry alert triggered: $days days remaining" -Level "WARNING"
                $global:yubiKeyAlertShown = $true
            }
        }
        elseif (-not $global:yubiKeyJob) {
            $checkYubiKey = ((Get-Date) - [DateTime]::Parse($config.YubiKeyLastCheck.Date)).TotalDays -ge 1
            if ($checkYubiKey) {
                $window.Dispatcher.Invoke([Action]{ $YubiKeyCertExpiryText.Text = "Checking YubiKey certificate..." })
                Start-YubiKeyCertCheckAsync
            } else {
                $window.Dispatcher.Invoke([Action]{ $YubiKeyCertExpiryText.Text = $config.YubiKeyLastCheck.Result })
            }
        }
    }
    catch {
        Handle-Error "Error updating system information: $_" -Source "Update-SystemInfo"
    }
}

function Get-BigFixStatus {
    try {
        $besService = Get-Service -Name BESClient -ErrorAction SilentlyContinue
        if ($besService) {
            if ($besService.Status -eq 'Running') {
                return $true, "BigFix (BESClient) Service: Running"
            }
            else {
                return $false, "BigFix (BESClient) is Installed but NOT Running (Status: $($besService.Status))"
            }
        }
        else {
            return $false, "BigFix (BESClient) not installed or not detected."
        }
    }
    catch {
        return $false, "Error retrieving BigFix status: $_"
    }
}

function Get-BitLockerStatus {
    try {
        $shell = New-Object -ComObject Shell.Application
        $bitlockerValue = $shell.NameSpace("C:").Self.ExtendedProperty("System.Volume.BitLockerProtection")
        switch ($bitlockerValue) {
            0 { return $false, "BitLocker is NOT Enabled on Drive C:" }
            1 { return $true,  "BitLocker is Enabled (Locked) on Drive C:" }
            2 { return $true,  "BitLocker is Enabled (Unlocked) on Drive C:" }
            3 { return $true,  "BitLocker is Enabled (Unknown State) on Drive C:" }
            6 { return $true,  "BitLocker is Fully Encrypted (Unlocked) on Drive C:" }
            default { return $false, "BitLocker code: $bitlockerValue (Unmapped status)" }
        }
    }
    catch {
        return $false, "Error retrieving BitLocker info: $_"
    }
}

function Get-AntivirusStatus {
    try {
        $antivirus = Get-CimInstance -Namespace "root\SecurityCenter2" -ClassName "AntiVirusProduct"
        if ($antivirus) {
            $antivirusNames = $antivirus | ForEach-Object { $_.displayName } | Sort-Object -Unique
            return $true, "Antivirus Detected: $($antivirusNames -join ', ')"
        }
        else {
            return $false, "No Antivirus Detected."
        }
    }
    catch {
        return $false, "Error retrieving antivirus information: $_"
    }
}

function Get-Code42Status {
    try {
        $code42Process = Get-Process -Name "Code42Service" -ErrorAction SilentlyContinue
        if ($code42Process) {
            return $true, "Code42 Service: Running (PID: $($code42Process.Id))"
        }
        else {
            $servicePath = "C:\Program Files\Code42\Code42Service.exe"
            if (Test-Path $servicePath) {
                return $false, "Code42 Service: Installed but NOT running."
            }
            else {
                return $false, "Code42 Service: Not installed."
            }
        }
    }
    catch {
        return $false, "Error checking Code42 Service: $_"
    }
}

function Get-FIPSStatus {
    try {
        $fipsSetting = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\FipsAlgorithmPolicy" -Name "Enabled" -ErrorAction SilentlyContinue
        if ($fipsSetting -and $fipsSetting.Enabled -eq 1) {
            return $true, "FIPS Compliance: Enabled"
        }
        else {
            return $false, "FIPS Compliance: Not Enabled"
        }
    }
    catch {
        return $false, "FIPS Compliance: Unknown (error: $_)"
    }
}

function Compare-Announcements {
    param($current, $last)
    $changes = @()
    if (-not $last.Text) { $last.Text = "" }
    if (-not $last.Details) { $last.Details = "" }
    if (-not $last.Links) { $last.Links = @{ Link1 = @{ Name = ""; Url = "" }; Link2 = @{ Name = ""; Url = "" } } }
    if ($current.Text -ne $last.Text) { $changes += "Text changed from '$($last.Text)' to '$($current.Text)'" }
    if ($current.Details -ne $last.Details) { $changes += "Details changed from '$($last.Details)' to '$($current.Details)'" }
    if ($current.Links.Link1.Name -ne $last.Links.Link1.Name -or $current.Links.Link1.Url -ne $last.Links.Link1.Url) {
        $changes += "Link1 changed from '$($last.Links.Link1.Name) ($($last.Links.Link1.Url))' to '$($current.Links.Link1.Name) ($($current.Links.Link1.Url))'"
    }
    if ($current.Links.Link2.Name -ne $last.Links.Link2.Name -or $current.Links.Link2.Url -ne $last.Links.Link2.Url) {
        $changes += "Link2 changed from '$($last.Links.Link2.Name) ($($last.Links.Link2.Url))' to '$($current.Links.Link2.Name) ($($current.Links.Link2.Url))'"
    }
    return $changes
}

function Update-Announcements {
    try {
        if (-not $global:contentData.Announcements) { throw "Announcements data missing" }
        if (-not $global:contentData.Announcements.Text) { throw "Announcements.Text missing" }
        if (-not $global:contentData.Announcements.Links -or 
            -not $global:contentData.Announcements.Links.Link1 -or 
            -not $global:contentData.Announcements.Links.Link1.Name -or 
            -not $global:contentData.Announcements.Links.Link1.Url -or 
            -not $global:contentData.Announcements.Links.Link2 -or 
            -not $global:contentData.Announcements.Links.Link2.Name -or 
            -not $global:contentData.Announcements.Links.Link2.Url) { 
            throw "Announcements.Links.Link1 or Link2 missing Name or Url" 
        }
        
        $currentAnnouncements = $global:contentData.Announcements
        $lastAnnouncements = $config.AnnouncementsLastState

        Write-Log "Current Announcements: $($currentAnnouncements | ConvertTo-Json -Depth 3)" -Level "INFO"
        Write-Log "Last Announcements: $($lastAnnouncements | ConvertTo-Json -Depth 3)" -Level "INFO"

        $changes = Compare-Announcements -current $currentAnnouncements -last $lastAnnouncements
        if ($changes.Count -gt 0 -and -not $AnnouncementsExpander.IsExpanded) {
            Write-Log "Announcements changed detected: $($changes -join '; ')" -Level "INFO"
            $window.Dispatcher.Invoke([Action]{ $AnnouncementsAlertIcon.Visibility = "Visible" })
            Write-Log "Announcements red dot set to visible" -Level "INFO"
            $global:announcementAlertActive = $true
        }
        else {
            Write-Log "No changes detected in Announcements or section already expanded" -Level "INFO"
        }

        $window.Dispatcher.Invoke([Action]{ 
            $AnnouncementsText.Text = $currentAnnouncements.Text
            $AnnouncementsDetailsText.Text = if ($currentAnnouncements.Details) { $currentAnnouncements.Details } else { "" }
            $AnnouncementsLink1.NavigateUri = [Uri]$currentAnnouncements.Links.Link1.Url
            $AnnouncementsLink1.Inlines.Clear()
            $AnnouncementsLink1.Inlines.Add($currentAnnouncements.Links.Link1.Name)
            $AnnouncementsLink2.NavigateUri = [Uri]$currentAnnouncements.Links.Link2.Url
            $AnnouncementsLink2.Inlines.Clear()
            $AnnouncementsLink2.Inlines.Add($currentAnnouncements.Links.Link2.Name)
        })
        $config.AnnouncementsLastState = $currentAnnouncements
        $config | ConvertTo-Json -Depth 3 | Out-File $configPath -Force
        Write-Log "Announcements updated: $($AnnouncementsText.Text)" -Level "INFO"
    }
    catch {
        Write-Log "Failed to update announcements: $_" -Level "ERROR"
        $window.Dispatcher.Invoke([Action]{ 
            $AnnouncementsText.Text = "Error fetching announcements."
            $AnnouncementsDetailsText.Text = ""
            $AnnouncementsLink1.NavigateUri = [Uri]$defaultContentData.Announcements.Links.Link1
            $AnnouncementsLink1.Inlines.Clear()
            $AnnouncementsLink1.Inlines.Add("Announcement Link 1")
            $AnnouncementsLink2.NavigateUri = [Uri]$defaultContentData.Announcements.Links.Link2
            $AnnouncementsLink2.Inlines.Clear()
            $AnnouncementsLink2.Inlines.Add("Announcement Link 2")
        })
    }
}

function Update-PatchingUpdates {
    try {
        $lastUpdate = Get-CimInstance -ClassName Win32_QuickFixEngineering | Sort-Object InstalledOn -Descending | Select-Object -First 1
        if ($lastUpdate) {
            $patchText = "Last Patch: $($lastUpdate.HotFixID) installed on $($lastUpdate.InstalledOn)"
        }
        else {
            $patchText = "No recent patches detected."
        }
        $window.Dispatcher.Invoke([Action]{ $PatchingUpdatesText.Text = $patchText })
        Write-Log "Patching status updated: $patchText" -Level "INFO"
    }
    catch {
        Write-Log "Failed to update patching: $_" -Level "ERROR"
        $window.Dispatcher.Invoke([Action]{ $PatchingUpdatesText.Text = "Error checking patches." })
    }
}

function Update-Support {
    try {
        if (-not $global:contentData.Support) { throw "Support data missing" }
        if (-not $global:contentData.Support.Text) { throw "Support.Text missing" }
        if (-not $global:contentData.Support.Links -or 
            -not $global:contentData.Support.Links.Link1 -or 
            -not $global:contentData.Support.Links.Link1.Name -or 
            -not $global:contentData.Support.Links.Link1.Url -or 
            -not $global:contentData.Support.Links.Link2 -or 
            -not $global:contentData.Support.Links.Link2.Name -or 
            -not $global:contentData.Support.Links.Link2.Url) { 
            throw "Support.Links.Link1 or Link2 missing Name or Url" 
        }
        
        $window.Dispatcher.Invoke([Action]{ 
            $SupportText.Text = $global:contentData.Support.Text
            $SupportLink1.NavigateUri = [Uri]$global:contentData.Support.Links.Link1.Url
            $SupportLink1.Inlines.Clear()
            $SupportLink1.Inlines.Add($global:contentData.Support.Links.Link1.Name)
            $SupportLink2.NavigateUri = [Uri]$global:contentData.Support.Links.Link2.Url
            $SupportLink2.Inlines.Clear()
            $SupportLink2.Inlines.Add($global:contentData.Support.Links.Link2.Name)
        })
        Write-Log "Support info updated: $($SupportText.Text)" -Level "INFO"
    }
    catch {
        Write-Log "Failed to update support: $_" -Level "ERROR"
        $window.Dispatcher.Invoke([Action]{ 
            $SupportText.Text = "Error loading support info."
            $SupportLink1.NavigateUri = [Uri]$defaultContentData.Support.Links.Link1
            $SupportLink1.Inlines.Clear()
            $SupportLink1.Inlines.Add("Support Link 1")
            $SupportLink2.NavigateUri = [Uri]$defaultContentData.Support.Links.Link2
            $SupportLink2.Inlines.Clear()
            $SupportLink2.Inlines.Add("Support Link 2")
        })
    }
}

function Update-EarlyAdopterTesting {
    try {
        if (-not $global:contentData.EarlyAdopter) { throw "EarlyAdopter data missing" }
        if (-not $global:contentData.EarlyAdopter.Text) { throw "EarlyAdopter.Text missing" }
        if (-not $global:contentData.EarlyAdopter.Links -or 
            -not $global:contentData.EarlyAdopter.Links.Link1 -or 
            -not $global:contentData.EarlyAdopter.Links.Link1.Name -or 
            -not $global:contentData.EarlyAdopter.Links.Link1.Url -or 
            -not $global:contentData.EarlyAdopter.Links.Link2 -or 
            -not $global:contentData.EarlyAdopter.Links.Link2.Name -or 
            -not $global:contentData.EarlyAdopter.Links.Link2.Url) { 
            throw "EarlyAdopter.Links.Link1 or Link2 missing Name or Url" 
        }
        
        $window.Dispatcher.Invoke([Action]{ 
            $EarlyAdopterText.Text = $global:contentData.EarlyAdopter.Text
            $EarlyAdopterLink1.NavigateUri = [Uri]$global:contentData.EarlyAdopter.Links.Link1.Url
            $EarlyAdopterLink1.Inlines.Clear()
            $EarlyAdopterLink1.Inlines.Add($global:contentData.EarlyAdopter.Links.Link1.Name)
            $EarlyAdopterLink2.NavigateUri = [Uri]$global:contentData.EarlyAdopter.Links.Link2.Url
            $EarlyAdopterLink2.Inlines.Clear()
            $EarlyAdopterLink2.Inlines.Add($global:contentData.EarlyAdopter.Links.Link2.Name)
        })
        Write-Log "Early adopter info updated: $($EarlyAdopterText.Text)" -Level "INFO"
    }
    catch {
        Write-Log "Failed to update early adopter: $_" -Level "ERROR"
        $window.Dispatcher.Invoke([Action]{ 
            $EarlyAdopterText.Text = "Error loading early adopter info."
            $EarlyAdopterLink1.NavigateUri = [Uri]$defaultContentData.EarlyAdopter.Links.Link1
            $EarlyAdopterLink1.Inlines.Clear()
            $EarlyAdopterLink1.Inlines.Add("Early Adopter Link 1")
            $EarlyAdopterLink2.NavigateUri = [Uri]$defaultContentData.EarlyAdopter.Links.Link2
            $EarlyAdopterLink2.Inlines.Clear()
            $EarlyAdopterLink2.Inlines.Add("Early Adopter Link 2")
        })
    }
}

function Update-Compliance {
    try {
        $antivirusStatus, $antivirusMessage = Get-AntivirusStatus
        $bitlockerStatus, $bitlockerMessage = Get-BitLockerStatus
        $bigfixStatus,    $bigfixMessage    = Get-BigFixStatus
        $code42Status,    $code42Message    = Get-Code42Status
        $fipsStatus,      $fipsMessage      = Get-FIPSStatus

        $window.Dispatcher.Invoke([Action]{ 
            $AntivirusStatusText.Text = $antivirusMessage
            $BitLockerStatusText.Text = $bitlockerMessage
            $BigFixStatusText.Text    = $bigfixMessage
            $Code42StatusText.Text    = $code42Message
            $FIPSStatusText.Text      = $fipsMessage

            if ($bitlockerStatus) { $BitLockerBorder.BorderBrush = 'Green' } else { $BitLockerBorder.BorderBrush = 'Red' }
            if ($bigfixStatus) { $BigFixBorder.BorderBrush = 'Green' } else { $BigFixBorder.BorderBrush = 'Red' }
            if ($code42Status) { $Code42Border.BorderBrush = 'Green' } else { $Code42Border.BorderBrush = 'Red' }
            if ($fipsStatus) { $FIPSBorder.BorderBrush = 'Green' } else { $FIPSBorder.BorderBrush = 'Red' }

            if ($antivirusStatus -and $bitlockerStatus -and $bigfixStatus -and $code42Status -and $fipsStatus) {
                $ComplianceStatusIndicator.Fill = "Green"
            } else {
                $ComplianceStatusIndicator.Fill = "Red"
            }
        })

        Write-Log "Compliance updated: Antivirus=$antivirusStatus, BitLocker=$bitlockerStatus, BigFix=$bigfixStatus, Code42=$code42Status, FIPS=$fipsStatus" -Level "INFO"
    }
    catch {
        Write-Log "Failed to update compliance: $_" -Level "ERROR"
        $window.Dispatcher.Invoke([Action]{ 
            $AntivirusStatusText.Text = "Error checking antivirus."
            $BitLockerStatusText.Text = "Error checking BitLocker."
            $BigFixStatusText.Text    = "Error checking BigFix."
            $Code42StatusText.Text    = "Error checking Code42."
            $FIPSStatusText.Text      = "Error checking FIPS."
            $ComplianceStatusIndicator.Fill = "Gray"
        })
    }
}

# ========================
# I) Tray Icon Management
# ========================
function Get-Icon {
    param(
        [string]$Path,
        [System.Drawing.Icon]$DefaultIcon
    )
    $fullPath = Join-Path $ScriptDir $Path
    if (-not (Test-Path $fullPath)) {
        Write-Log "$fullPath not found. Using default icon." -Level "WARNING"
        return $DefaultIcon
    }
    else {
        try {
            $icon = New-Object System.Drawing.Icon($fullPath)
            Write-Log "Custom icon loaded from ${fullPath}." -Level "INFO"
            return $icon
        }
        catch {
            Handle-Error "Error loading icon from ${fullPath}: $_" -Source "Get-Icon"
            return $DefaultIcon
        }
    }
}

function Update-TrayIcon {
    try {
        $antivirusStatus, $antivirusMessage = Get-AntivirusStatus
        $bitlockerStatus, $bitlockerMessage = Get-BitLockerStatus
        $bigfixStatus,    $bigfixMessage    = Get-BigFixStatus
        $code42Status,    $code42Message    = Get-Code42Status
        $fipsStatus,      $fipsMessage      = Get-FIPSStatus

        $yubiKeyCert = $YubiKeyCertExpiryText.Text
        $yubikeyStatus = $yubiKeyCert -notmatch "Unable to determine expiry date" -and $yubiKeyCert -ne "YubiKey not present"

        if ($antivirusStatus -and $bitlockerStatus -and $yubikeyStatus -and $code42Status -and $fipsStatus -and $bigfixStatus) {
            $TrayIcon.Icon = Get-Icon -Path $config.IconPaths.Main -DefaultIcon ([System.Drawing.SystemIcons]::Application)
            Write-Log "Tray icon set to icon.ico" -Level "INFO"
            $TrayIcon.Text = "SHOT - Healthy"
        }
        else {
            $TrayIcon.Icon = Get-Icon -Path $config.IconPaths.Warning -DefaultIcon ([System.Drawing.SystemIcons]::Application)
            Write-Log "Tray icon set to warning.ico" -Level "INFO"
            $TrayIcon.Text = "SHOT - Warning"
        }

        $TrayIcon.Visible = $true
        Write-Log "Tray icon and status updated." -Level "INFO"
    }
    catch {
        Handle-Error "Error updating tray icon: $_" -Source "Update-TrayIcon"
    }
}

# ========================
# J) Enhanced Logs Management (ListView)
# ========================
function Update-Logs {
    try {
        if (Test-Path $LogFilePath) {
            $logContent = Get-Content -Path $LogFilePath -Tail 100 -ErrorAction SilentlyContinue
            $logEntries = @()
            foreach ($line in $logContent) {
                if ($line -match "^\[(?<timestamp>[^\]]+)\]\s\[(?<level>[^\]]+)\]\s(?<message>.*)$") {
                    $logEntries += [PSCustomObject]@{
                        Timestamp = $matches['timestamp']
                        Message   = $matches['message']
                    }
                }
            }
            $window.Dispatcher.Invoke([Action]{ $LogListView.ItemsSource = $logEntries })
        }
        else {
            $window.Dispatcher.Invoke([Action]{ $LogListView.ItemsSource = @([PSCustomObject]@{Timestamp="N/A"; Message="Log file not found."}) })
        }
        Write-Log "Logs updated in GUI." -Level "INFO"
    }
    catch {
        Handle-Error "Error loading logs: $_" -Source "Update-Logs"
    }
}

function Export-Logs {
    try {
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "Log Files (*.log)|*.log|All Files (*.*)|*.*"
        $saveFileDialog.FileName = "SHOT.log"
        if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Copy-Item -Path $LogFilePath -Destination $saveFileDialog.FileName -Force
            Write-Log "Logs exported to $($saveFileDialog.FileName)" -Level "INFO"
        }
    }
    catch {
        Handle-Error "Error exporting logs: $_" -Source "Export-Logs"
    }
}

# ========================
# K) Window Visibility Management
# ========================
function Set-WindowPosition {
    try {
        $window.Dispatcher.Invoke([Action]{ 
            $window.UpdateLayout()
            $primary = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
            $window.Left = $primary.X + $primary.Width - $window.ActualWidth - 10
            $window.Top  = $primary.Y + $primary.Height - $window.ActualHeight - 50
        })
        Write-Log "Window position set: Left=$($window.Left), Top=$($window.Top)" -Level "INFO"
    }
    catch {
        Handle-Error "Error setting window position: $_" -Source "Set-WindowPosition"
    }
}

function Toggle-WindowVisibility {
    try {
        $window.Dispatcher.Invoke([Action]{ 
            if ($window.Visibility -eq 'Visible') {
                $window.Hide()
                Write-Log "Dashboard hidden via Toggle-WindowVisibility." -Level "INFO"
            }
            else {
                Set-WindowPosition
                $window.Show()
                Write-Log "Dashboard shown via Toggle-WindowVisibility." -Level "INFO"
            }
        })
    }
    catch {
        Handle-Error "Error toggling window visibility: $_" -Source "Toggle-WindowVisibility"
    }
}

# ========================
# L) Button and Event Handlers
# ========================
$ExportLogsButton.Add_Click({ Export-Logs })

# Hyperlink Click Handlers
$AnnouncementsLink1.Add_RequestNavigate({ param($sender, $e) Start-Process $e.Uri.AbsoluteUri; $e.Handled = $true; Write-Log "Clicked Announcements Link 1: $($e.Uri.AbsoluteUri)" -Level "INFO" })
$AnnouncementsLink2.Add_RequestNavigate({ param($sender, $e) Start-Process $e.Uri.AbsoluteUri; $e.Handled = $true; Write-Log "Clicked Announcements Link 2: $($e.Uri.AbsoluteUri)" -Level "INFO" })
$SupportLink1.Add_RequestNavigate({ param($sender, $e) Start-Process $e.Uri.AbsoluteUri; $e.Handled = $true; Write-Log "Clicked Support Link 1: $($e.Uri.AbsoluteUri)" -Level "INFO" })
$SupportLink2.Add_RequestNavigate({ param($sender, $e) Start-Process $e.Uri.AbsoluteUri; $e.Handled = $true; Write-Log "Clicked Support Link 2: $($e.Uri.AbsoluteUri)" -Level "INFO" })
$EarlyAdopterLink1.Add_RequestNavigate({ param($sender, $e) Start-Process $e.Uri.AbsoluteUri; $e.Handled = $true; Write-Log "Clicked Early Adopter Link 1: $($e.Uri.AbsoluteUri)" -Level "INFO" })
$EarlyAdopterLink2.Add_RequestNavigate({ param($sender, $e) Start-Process $e.Uri.AbsoluteUri; $e.Handled = $true; Write-Log "Clicked Early Adopter Link 2: $($e.Uri.AbsoluteUri)" -Level "INFO" })

# Announcements Expander Event Handler
$AnnouncementsExpander.Add_Expanded({
    if ($global:announcementAlertActive) {
        $window.Dispatcher.Invoke([Action]{ $AnnouncementsAlertIcon.Visibility = "Hidden" })
        $global:announcementAlertActive = $false
        Write-Log "Announcements red dot hidden on expand" -Level "INFO"
    }
})

# ========================
# M) Create & Configure Tray Icon with Collapsible Menu
# ========================
$TrayIcon = New-Object System.Windows.Forms.NotifyIcon
$TrayIcon.Icon = Get-Icon -Path $config.IconPaths.Main -DefaultIcon ([System.Drawing.SystemIcons]::Application)
$TrayIcon.Text = "SHOT v$ScriptVersion"
$TrayIcon.Visible = $true
Write-Log "Tray icon initialized with icon.ico" -Level "INFO"
Write-Log "Note: To ensure the SHOT tray icon is always visible, right-click the taskbar, select 'Taskbar settings', scroll to 'Notification area', click 'Select which icons appear on the taskbar', and set 'SHOT' to 'On'." -Level "INFO"

$ContextMenu = New-Object System.Windows.Forms.ContextMenu
$MenuItemShow = New-Object System.Windows.Forms.MenuItem("Show Dashboard")
$MenuItemQuickActions = New-Object System.Windows.Forms.MenuItem("Quick Actions")
$MenuItemRefresh = New-Object System.Windows.Forms.MenuItem("Refresh Now")
$MenuItemExportLogs = New-Object System.Windows.Forms.MenuItem("Export Logs")
$MenuItemExit = New-Object System.Windows.Forms.MenuItem("Exit")

$MenuItemQuickActions.MenuItems.Add($MenuItemRefresh)
$MenuItemQuickActions.MenuItems.Add($MenuItemExportLogs)
$ContextMenu.MenuItems.Add($MenuItemShow)
$ContextMenu.MenuItems.Add($MenuItemQuickActions)
$ContextMenu.MenuItems.Add($MenuItemExit)
$TrayIcon.ContextMenu = $ContextMenu

$TrayIcon.add_MouseClick({
    param($sender,$e)
    try {
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            Toggle-WindowVisibility
        }
    }
    catch {
        Handle-Error "Error handling tray icon mouse click: $_" -Source "TrayIcon"
    }
})

$MenuItemShow.add_Click({ Toggle-WindowVisibility })
$MenuItemRefresh.add_Click({ 
    $global:contentData = Fetch-ContentData
    Update-TrayIcon
    Update-SystemInfo
    Update-Logs
    Update-Announcements
    Update-PatchingUpdates
    Update-Support
    Update-EarlyAdopterTesting
    Update-Compliance
    Write-Log "Manual refresh triggered from tray menu" -Level "INFO"
})
$MenuItemExportLogs.add_Click({ Export-Logs })
$MenuItemExit.add_Click({
    try {
        Write-Log "Exit clicked by user." -Level "INFO"
        $dispatcherTimer.Stop()
        Write-Log "DispatcherTimer stopped." -Level "INFO"
        if ($global:yubiKeyJob) {
            Stop-Job -Job $global:yubiKeyJob -ErrorAction SilentlyContinue
            Remove-Job -Job $global:yubiKeyJob -Force -ErrorAction SilentlyContinue
            Write-Log "YubiKey job stopped and removed." -Level "INFO"
        }
        $TrayIcon.Visible = $false
        $TrayIcon.Dispose()
        Write-Log "Tray icon disposed." -Level "INFO"
        $window.Dispatcher.InvokeShutdown()
        Write-Log "Application exited via tray menu." -Level "INFO"
    }
    catch {
        Handle-Error "Error during application exit: $_" -Source "Exit"
    }
})

# ========================
# O) DispatcherTimer for Periodic Updates
# ========================
$dispatcherTimer = New-Object System.Windows.Threading.DispatcherTimer
$dispatcherTimer.Interval = [TimeSpan]::FromSeconds($config.RefreshInterval)
$dispatcherTimer.add_Tick({
    try {
        $global:contentData = Fetch-ContentData
        Update-TrayIcon
        Update-SystemInfo
        Update-Logs
        Update-Announcements
        Update-PatchingUpdates
        Update-Support
        Update-EarlyAdopterTesting
        Update-Compliance
        Write-Log "Dispatcher tick completed" -Level "INFO"
    }
    catch {
        Handle-Error "Error during timer tick: $_" -Source "DispatcherTimer"
    }
})
$dispatcherTimer.Start()
Write-Log "DispatcherTimer started with interval $($config.RefreshInterval) seconds" -Level "INFO"

# ========================
# P) Dispatcher Exception Handling
# ========================
function Handle-DispatcherUnhandledException {
    param(
        [object]$sender,
        [System.Windows.Threading.DispatcherUnhandledExceptionEventArgs]$args
    )
    Handle-Error "Unhandled Dispatcher exception: $($args.Exception.Message)" -Source "Dispatcher"
}

Register-ObjectEvent -InputObject $window.Dispatcher -EventName UnhandledException -Action {
    param($sender, $args)
    Handle-DispatcherUnhandledException -sender $sender -args $args
}

# ========================
# Q) Initial Update & Start Dispatcher
# ========================
try {
    $window.Add_Loaded({ Set-WindowPosition })
    $global:contentData = Fetch-ContentData
    Write-Log "Initial contentData set: $($global:contentData | ConvertTo-Json -Depth 3)" -Level "INFO"
    
    Update-SystemInfo
    Update-TrayIcon
    Update-Logs
    Update-Announcements
    Update-PatchingUpdates
    Update-Support
    Update-EarlyAdopterTesting
    Update-Compliance
    Log-DotNetVersion
    Write-Log "Initial update completed" -Level "INFO"
}
catch {
    Handle-Error "Error during initial update: $_" -Source "InitialUpdate"
}

$window.Add_Closing({
    param($sender,$eventArgs)
    try {
        $eventArgs.Cancel = $true
        $window.Hide()
        Write-Log "Dashboard hidden via window closing event." -Level "INFO"
    }
    catch {
        Handle-Error "Error handling window closing: $_" -Source "WindowClosing"
    }
})

Write-Log "About to call Dispatcher.Run()..." -Level "INFO"
[System.Windows.Threading.Dispatcher]::Run()
Write-Log "Dispatcher ended; script exiting." -Level "INFO"
