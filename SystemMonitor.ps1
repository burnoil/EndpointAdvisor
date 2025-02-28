###############################################################################
# SystemMonitor.ps1 - Revised Version with Advanced Logging, External Config,
# DispatcherTimer, Enhanced Log Viewer, New Sections,
# Auto-Sizing, Anchored to Bottom Right on Primary Display, .NET Version Logging, and Fixed Icon URIs
###############################################################################

# Ensure $PSScriptRoot is defined for older versions
if (-not $PSScriptRoot) {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    $ScriptDir = $PSScriptRoot
}

# ========================
# A) External Configuration
# ========================
$configPath = Join-Path $ScriptDir "SystemMonitor.config.json"
if (Test-Path $configPath) {
    try {
        $config = Get-Content $configPath -Raw | ConvertFrom-Json
    }
    catch {
        Write-Host "Error reading configuration file. Using default settings."
        $config = $null
    }
}
if (-not $config) {
    $config = @{
        RefreshInterval    = 30
        LogRotationSizeMB  = 5
        DefaultLogLevel    = "INFO"
        IconPaths          = @{
            Healthy = (Join-Path $ScriptDir "healthy.ico")
            Warning = (Join-Path $ScriptDir "warning.ico")
        }
        SupportLinks       = @{
            Link1 = "https://support.company.com/help"
            Link2 = "https://support.company.com/tickets"
        }
        EarlyAdopterLinks  = @{
            Link1 = "https://beta.company.com/signup"
            Link2 = "https://beta.company.com/info"
        }
        AnnouncementLinks  = @{
            Link1 = "https://company.com/news1"
            Link2 = "https://company.com/news2"
        }
    }
    $config | ConvertTo-Json | Out-File $configPath -Force
}

$healthyIconUri = "file:///" + ($config.IconPaths.Healthy -replace '\\','/')

# ========================
# B) Log File Setup & Rotation
# ========================
$LogFilePath = Join-Path $ScriptDir "SystemMonitor.log"
$LogDirectory = Split-Path $LogFilePath
if (-not (Test-Path $LogDirectory)) {
    New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
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

# ========================
# C) Advanced Logging & Error Handling
# ========================
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = $config.DefaultLogLevel
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Add-Content -Path $LogFilePath -Value $logEntry
}

function Handle-Error {
    param(
        [string]$ErrorMessage,
        [string]$Source = ""
    )
    if ($Source) {
        $ErrorMessage = "[$Source] $ErrorMessage"
    }
    Write-Log $ErrorMessage -Level "ERROR"
}

function Show-Notification {
    param(
        [string]$Title,
        [string]$Message,
        [System.Drawing.Icon]$Icon = [System.Drawing.SystemIcons]::Information
    )
    Write-Log "Notification suppressed: $Title - $Message" -Level "INFO"
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
# D) Import Required Assemblies
# ========================
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ========================
# E) XAML Layout Definition (Compact UI with New Sections)
# ========================
$xamlString = @"
<?xml version="1.0" encoding="utf-8"?>
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="System Monitor"
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
        <Image Source="$healthyIconUri" Width="20" Height="20" Margin="0,0,4,0"/>
        <TextBlock Text="System Monitoring Dashboard"
                   FontSize="14" FontWeight="Bold" Foreground="White"
                   VerticalAlignment="Center"/>
      </StackPanel>
    </Border>
    <!-- Content Area -->
    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
      <StackPanel VerticalAlignment="Top">
        <!-- Information Section -->
        <Expander Header="Information" FontSize="12" Foreground="#0078D7" IsExpanded="True" Margin="0,2,0,2">
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
        <!-- BigFix Section -->
        <Expander x:Name="BigFixExpander" Header="BigFix (BESClient)" FontSize="12" Foreground="#4b0082" IsExpanded="True" Margin="0,2,0,2">
          <Border x:Name="BigFixBorder" BorderBrush="#4b0082" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="BigFixStatusText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
            </StackPanel>
          </Border>
        </Expander>
        <!-- Announcements -->
        <Expander Header="Announcements" FontSize="12" Foreground="#000080" IsExpanded="False" Margin="0,2,0,2">
          <Border BorderBrush="#000080" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="AnnouncementsText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
              <TextBlock FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300">
                <Hyperlink x:Name="AnnouncementsLink1" NavigateUri="https://company.com/news1">Announcement Link 1</Hyperlink>
              </TextBlock>
              <TextBlock FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300">
                <Hyperlink x:Name="AnnouncementsLink2" NavigateUri="https://company.com/news2">Announcement Link 2</Hyperlink>
              </TextBlock>
            </StackPanel>
          </Border>
        </Expander>
        <!-- Patching and Updates -->
        <Expander Header="Patching and Updates" FontSize="12" Foreground="#008000" IsExpanded="False" Margin="0,2,0,2">
          <Border BorderBrush="#008000" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="PatchingUpdatesText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
            </StackPanel>
          </Border>
        </Expander>
        <!-- Support -->
        <Expander Header="Support" FontSize="12" Foreground="#800000" IsExpanded="False" Margin="0,2,0,2">
          <Border BorderBrush="#800000" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="SupportText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
              <TextBlock FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300">
                <Hyperlink x:Name="SupportLink1" NavigateUri="https://support.company.com/help">Support Link 1</Hyperlink>
              </TextBlock>
              <TextBlock FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300">
                <Hyperlink x:Name="SupportLink2" NavigateUri="https://support.company.com/tickets">Support Link 2</Hyperlink>
              </TextBlock>
            </StackPanel>
          </Border>
        </Expander>
        <!-- Open Early Adopter Testing -->
        <Expander Header="Open Early Adopter Testing" FontSize="12" Foreground="#FF00FF" IsExpanded="False" Margin="0,2,0,2">
          <Border BorderBrush="#FF00FF" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="EarlyAdopterText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
              <TextBlock FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300">
                <Hyperlink x:Name="EarlyAdopterLink1" NavigateUri="https://beta.company.com/signup">Early Adopter Link 1</Hyperlink>
              </TextBlock>
              <TextBlock FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300">
                <Hyperlink x:Name="EarlyAdopterLink2" NavigateUri="https://beta.company.com/info">Early Adopter Link 2</Hyperlink>
              </TextBlock>
            </StackPanel>
          </Border>
        </Expander>
        <!-- Compliance Section with Antivirus, BitLocker, Code42, and FIPS -->
        <Expander Header="Compliance" FontSize="12" Foreground="#B22222" IsExpanded="False" Margin="0,2,0,2">
          <Border BorderBrush="#B22222" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
            <StackPanel>
              <!-- Nested Antivirus Section -->
              <Border BorderBrush="#28a745" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
                <StackPanel>
                  <TextBlock Text="Antivirus Status" FontSize="11" FontWeight="Bold" Margin="2" Foreground="#28a745"/>
                  <TextBlock x:Name="AntivirusStatusText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
                </StackPanel>
              </Border>
              <!-- Nested BitLocker Section -->
              <Border x:Name="BitLockerBorder" BorderBrush="#6c757d" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
                <StackPanel>
                  <TextBlock Text="BitLocker Status" FontSize="11" FontWeight="Bold" Margin="2" Foreground="#6c757d"/>
                  <TextBlock x:Name="BitLockerStatusText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
                </StackPanel>
              </Border>
              <!-- Nested Code42 Section -->
              <Border x:Name="Code42Border" BorderBrush="#800080" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
                <StackPanel>
                  <TextBlock Text="Code42 Service Status" FontSize="11" FontWeight="Bold" Margin="2" Foreground="#800080"/>
                  <TextBlock x:Name="Code42StatusText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
                </StackPanel>
              </Border>
              <!-- Nested FIPS Section -->
              <Border x:Name="FIPSBorder" BorderBrush="#FF4500" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
                <StackPanel>
                  <TextBlock Text="FIPS Compliance Status" FontSize="11" FontWeight="Bold" Margin="2" Foreground="#FF4500"/>
                  <TextBlock x:Name="FIPSStatusText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
                </StackPanel>
              </Border>
            </StackPanel>
          </Border>
        </Expander>
        <!-- Logs Section (Enhanced Log Viewer) -->
        <Expander Header="Logs" FontSize="12" Foreground="#ff8c00" IsExpanded="False" Margin="0,2,0,2">
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
              <Button x:Name="ExportLogsButton" Content="Export Logs" Width="80" Margin="2" HorizontalAlignment="Right"/>
            </StackPanel>
          </Border>
        </Expander>
        <!-- About Section -->
        <Expander Header="About" FontSize="12" Foreground="#000000" IsExpanded="False" Margin="0,2,0,2">
          <Border BorderBrush="#000000" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="AboutText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300">
                <TextBlock.Text><![CDATA[System Monitor v1.0
© 2025 System Monitor. All rights reserved.
Built with PowerShell and WPF.]]></TextBlock.Text>
              </TextBlock>
            </StackPanel>
          </Border>
        </Expander>
      </StackPanel>
    </ScrollViewer>
    <!-- Footer Section -->
    <TextBlock Grid.Row="2" Text="© 2025 System Monitor" FontSize="10" Foreground="Gray" HorizontalAlignment="Center" Margin="0,4,0,0"/>
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
}
catch {
    Handle-Error "Failed to load the XAML layout. Error: $_" -Source "XAML"
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

$AnnouncementsText   = $window.FindName("AnnouncementsText")
$AnnouncementsLink1  = $window.FindName("AnnouncementsLink1")
$AnnouncementsLink2  = $window.FindName("AnnouncementsLink2")
$PatchingUpdatesText = $window.FindName("PatchingUpdatesText")
$SupportText         = $window.FindName("SupportText")
$SupportLink1        = $window.FindName("SupportLink1")
$SupportLink2        = $window.FindName("SupportLink2")
$EarlyAdopterText    = $window.FindName("EarlyAdopterText")
$EarlyAdopterLink1   = $window.FindName("EarlyAdopterLink1")
$EarlyAdopterLink2   = $window.FindName("EarlyAdopterLink2")

$LogListView         = $window.FindName("LogListView")
$ExportLogsButton    = $window.FindName("ExportLogsButton")

$BitLockerBorder     = $window.FindName("BitLockerBorder")
$BigFixExpander      = $window.FindName("BigFixExpander")
$BigFixBorder        = $window.FindName("BigFixBorder")
$Code42Border        = $window.FindName("Code42Border")
$FIPSBorder          = $window.FindName("FIPSBorder")

# Global variable to store the last YubiKey job
$yubiKeyJob = $null

# ========================
# H) Modularized System Information Functions
# ========================
function Get-YubiKeyCertExpiryDays {
    param([string]$ykmanPathPassed = "C:\Program Files\Yubico\Yubikey Manager\ykman.exe")
    try {
        # Verify ykman.exe exists
        if (-not (Test-Path $ykmanPathPassed)) {
            throw "ykman.exe not found at $ykmanPathPassed"
        }
        Write-Log "ykman.exe found at $ykmanPathPassed" -Level "INFO"

        # Check YubiKey general info
        $yubiKeyInfo = & $ykmanPathPassed info 2>$null
        if (-not $yubiKeyInfo) {
            throw "No YubiKey detected by ykman"
        }
        Write-Log "YubiKey detected: $yubiKeyInfo" -Level "INFO"

        # Check PIV-specific info
        $pivInfo = & $ykmanPathPassed "piv" "info" 2>$null
        if ($pivInfo) {
            Write-Log "PIV info: $pivInfo" -Level "INFO"
        } else {
            Write-Log "No PIV info available" -Level "WARNING"
        }

        # List of PIV slots to check
        $slots = @("9a", "9c", "9d", "9e")
        $certPem = $null
        $slotUsed = $null

        # Try each slot until a certificate is found
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

        # Save the PEM to a temporary file
        $tempFile = [System.IO.Path]::GetTempFileName()
        $certPem | Out-File $tempFile -Encoding ASCII

        # Convert PEM to certificate object
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $cert.Import($tempFile)

        # Calculate days until expiry
        $today = Get-Date
        $expiryDate = $cert.NotAfter
        $daysUntilExpiry = ($expiryDate - $today).Days

        # Clean up temporary file
        Remove-Item $tempFile -Force

        if ($daysUntilExpiry -lt 0) {
            return "YubiKey Certificate (Slot $slotUsed): Expired ($(-$daysUntilExpiry) days ago)"
        } else {
            return "YubiKey Certificate (Slot $slotUsed): $daysUntilExpiry days until expiry ($expiryDate)"
        }
    }
    catch {
        Write-Log "Error retrieving YubiKey certificate expiry: $_" -Level "ERROR"
        return "YubiKey Certificate: Unable to determine expiry date - $_"
    }
}

function Start-YubiKeyCertCheck {
    if ($global:yubiKeyJob -and $global:yubiKeyJob.State -eq "Running") {
        Write-Log "YubiKey certificate check already in progress." -Level "INFO"
        return
    }

    $global:yubiKeyJob = Start-Job -ScriptBlock {
        param($ykmanPath, $LogFilePathPass)

        # Re-import logging function in job context
        function Write-Log {
            param(
                [string]$Message,
                [ValidateSet("INFO", "WARNING", "ERROR")]
                [string]$Level = "INFO"
            )
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logEntry = "[$timestamp] [$Level] $Message"
            Add-Content -Path $LogFilePathPass -Value $logEntry
        }

        try {
            # Verify ykman.exe exists
            if (-not (Test-Path $ykmanPath)) {
                throw "ykman.exe not found at $ykmanPath"
            }
            Write-Log "ykman.exe found at $ykmanPath" -Level "INFO"

            # Check YubiKey general info
            $yubiKeyInfo = & $ykmanPath info 2>$null
            if (-not $yubiKeyInfo) {
                throw "No YubiKey detected by ykman"
            }
            Write-Log "YubiKey detected: $yubiKeyInfo" -Level "INFO"

            # Check PIV-specific info
            $pivInfo = & $ykmanPath "piv" "info" 2>$null
            if ($pivInfo) {
                Write-Log "PIV info: $pivInfo" -Level "INFO"
            } else {
                Write-Log "No PIV info available" -Level "WARNING"
            }

            # List of PIV slots to check
            $slots = @("9a", "9c", "9d", "9e")
            $certPem = $null
            $slotUsed = $null

            # Try each slot until a certificate is found
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

            # Save the PEM to a temporary file
            $tempFile = [System.IO.Path]::GetTempFileName()
            $certPem | Out-File $tempFile -Encoding ASCII

            # Convert PEM to certificate object
            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
            $cert.Import($tempFile)

            # Calculate days until expiry
            $today = Get-Date
            $expiryDate = $cert.NotAfter
            $daysUntilExpiry = ($expiryDate - $today).Days

            # Clean up temporary file
            Remove-Item $tempFile -Force

            if ($daysUntilExpiry -lt 0) {
                return "YubiKey Certificate (Slot $slotUsed): Expired ($(-$daysUntilExpiry) days ago)"
            } else {
                return "YubiKey Certificate (Slot $slotUsed): $daysUntilExpiry days until expiry ($expiryDate)"
            }
        }
        catch {
            Write-Log "Error retrieving YubiKey certificate expiry: $_" -Level "ERROR"
            return "YubiKey Certificate: Unable to determine expiry date - $_"
        }
    } -ArgumentList "C:\Program Files\Yubico\Yubikey Manager\ykman.exe", $LogFilePath

    Write-Log "Started YubiKey certificate check job." -Level "INFO"
}

function Update-SystemInfo {
    try {
        $user = [System.Environment]::UserName
        $LoggedOnUserText.Text = "Logged-in User: $user"
        Write-Log "Logged-in User: $user" -Level "INFO"

        $machine = Get-CimInstance -ClassName Win32_ComputerSystem
        $machineType = "$($machine.Manufacturer) $($machine.Model)"
        $MachineTypeText.Text = $machineType
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
        $OSVersionText.Text = $osVersion
        Write-Log "OS Version: $osVersion" -Level "INFO"

        $uptime = (Get-Date) - $os.LastBootUpTime
        $systemUptime = "$([math]::Floor($uptime.TotalDays)) days $($uptime.Hours) hours"
        $SystemUptimeText.Text = $systemUptime
        Write-Log "System Uptime: $systemUptime" -Level "INFO"

        $drive = Get-PSDrive -Name C
        $usedDiskSpace = "$([math]::Round(($drive.Used / 1GB), 2)) GB of $([math]::Round((($drive.Free + $drive.Used) / 1GB), 2)) GB"
        $UsedDiskSpaceText.Text = $usedDiskSpace
        Write-Log "Used Disk Space: $usedDiskSpace" -Level "INFO"

        $ipv4s = Get-NetIPAddress -AddressFamily IPv4 | Where-Object {
            $_.IPAddress -notlike "127.*" -and $_.IPAddress -notlike "169.254.*" -and
            $_.IPAddress -notin @("0.0.0.0","255.255.255.255") -and $_.PrefixOrigin -ne "WellKnown"
        } | Select-Object -ExpandProperty IPAddress -ErrorAction SilentlyContinue

        if ($ipv4s) {
            $ipList = $ipv4s -join ", "
            $IpAddressText.Text = "IPv4 Address(es): $ipList"
            Write-Log "IP Address(es): $ipList" -Level "INFO"
        }
        else {
            $IpAddressText.Text = "IPv4 Address(es): None detected"
            Write-Log "No valid IPv4 addresses found." -Level "WARNING"
        }

        # Check if YubiKey job has completed
        if ($global:yubiKeyJob -and $global:yubiKeyJob.State -eq "Completed") {
            $yubiKeyResult = Receive-Job -Job $global:yubiKeyJob
            $YubiKeyCertExpiryText.Text = $yubiKeyResult
            Remove-Job -Job $global:yubiKeyJob -Force
            $global:yubiKeyJob = $null
            Write-Log "YubiKey certificate check completed: $yubiKeyResult" -Level "INFO"
        }
        elseif (-not $global:yubiKeyJob) {
            $YubiKeyCertExpiryText.Text = "Checking YubiKey certificate..."
            Start-YubiKeyCertCheck
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

function Update-Announcements {
    try {
        $announcements = "Latest update: System Monitor v1.1 released on $(Get-Date -Format 'yyyy-MM-dd')."
        $AnnouncementsText.Text = $announcements
        $AnnouncementsLink1.NavigateUri = [Uri]$config.AnnouncementLinks.Link1
        $AnnouncementsLink1.Inlines.Clear()
        $AnnouncementsLink1.Inlines.Add("Announcement Link 1")
        $AnnouncementsLink2.NavigateUri = [Uri]$config.AnnouncementLinks.Link2
        $AnnouncementsLink2.Inlines.Clear()
        $AnnouncementsLink2.Inlines.Add("Announcement Link 2")
        Write-Log "Announcements updated: $announcements" -Level "INFO"
    }
    catch {
        $AnnouncementsText.Text = "Error fetching announcements."
        Handle-Error "Error updating announcements: $_" -Source "Update-Announcements"
    }
}

function Update-PatchingUpdates {
    try {
        $lastUpdate = Get-CimInstance -ClassName Win32_QuickFixEngineering | Sort-Object InstalledOn -Descending | Select-Object -First 1
        if ($lastUpdate) {
            $PatchingUpdatesText.Text = "Last Patch: $($lastUpdate.HotFixID) installed on $($lastUpdate.InstalledOn)"
        }
        else {
            $PatchingUpdatesText.Text = "No recent patches detected."
        }
        Write-Log "Patching status updated." -Level "INFO"
    }
    catch {
        $PatchingUpdatesText.Text = "Error checking patches."
        Handle-Error "Error updating patching: $_" -Source "Update-PatchingUpdates"
    }
}

function Update-Support {
    try {
        $SupportText.Text = "Contact IT Support: support@company.com | Phone: 1-800-555-1234"
        $SupportLink1.NavigateUri = [Uri]$config.SupportLinks.Link1
        $SupportLink1.Inlines.Clear()
        $SupportLink1.Inlines.Add("Support Link 1")
        $SupportLink2.NavigateUri = [Uri]$config.SupportLinks.Link2
        $SupportLink2.Inlines.Clear()
        $SupportLink2.Inlines.Add("Support Link 2")
        Write-Log "Support info updated." -Level "INFO"
    }
    catch {
        $SupportText.Text = "Error loading support info."
        Handle-Error "Error updating support: $_" -Source "Update-Support"
    }
}

function Update-EarlyAdopterTesting {
    try {
        $EarlyAdopterText.Text = "Join our beta program!"
        $EarlyAdopterLink1.NavigateUri = [Uri]$config.EarlyAdopterLinks.Link1
        $EarlyAdopterLink1.Inlines.Clear()
        $EarlyAdopterLink1.Inlines.Add("Early Adopter Link 1")
        $EarlyAdopterLink2.NavigateUri = [Uri]$config.EarlyAdopterLinks.Link2
        $EarlyAdopterLink2.Inlines.Clear()
        $EarlyAdopterLink2.Inlines.Add("Early Adopter Link 2")
        Write-Log "Early adopter info updated." -Level "INFO"
    }
    catch {
        $EarlyAdopterText.Text = "Error loading early adopter info."
        Handle-Error "Error updating early adopter: $_" -Source "Update-EarlyAdopterTesting"
    }
}

function Update-Compliance {
    try {
        $antivirusStatus, $antivirusMessage = Get-AntivirusStatus
        $bitlockerStatus, $bitlockerMessage = Get-BitLockerStatus
        $code42Status,    $code42Message    = Get-Code42Status
        $fipsStatus,      $fipsMessage      = Get-FIPSStatus

        $AntivirusStatusText.Text = $antivirusMessage
        $BitLockerStatusText.Text = $bitlockerMessage
        $Code42StatusText.Text    = $code42Message
        $FIPSStatusText.Text      = $fipsMessage

        if ($bitlockerStatus) {
            $BitLockerBorder.BorderBrush = 'Green'
        }
        else {
            $BitLockerBorder.BorderBrush = 'Red'
        }

        if ($code42Status) {
            $Code42Border.BorderBrush = 'Green'
        }
        else {
            $Code42Border.BorderBrush = 'Red'
        }

        if ($fipsStatus) {
            $FIPSBorder.BorderBrush = 'Green'
        }
        else {
            $FIPSBorder.BorderBrush = 'Red'
        }

        Write-Log "Compliance updated: Antivirus=$antivirusStatus, BitLocker=$bitlockerStatus, Code42=$code42Status, FIPS=$fipsStatus" -Level "INFO"
    }
    catch {
        $AntivirusStatusText.Text = "Error checking antivirus."
        $BitLockerStatusText.Text = "Error checking BitLocker."
        $Code42StatusText.Text    = "Error checking Code42."
        $FIPSStatusText.Text      = "Error checking FIPS."
        Handle-Error "Error updating compliance: $_" -Source "Update-Compliance"
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
    if (-not (Test-Path $Path)) {
        Write-Log "$Path not found. Using default icon." -Level "WARNING"
        return $DefaultIcon
    }
    else {
        try {
            $icon = New-Object System.Drawing.Icon($Path)
            Write-Log "Custom icon loaded from ${Path}." -Level "INFO"
            return $icon
        }
        catch {
            Handle-Error "Error loading icon from ${Path}: $_. Using default icon." -Source "Get-Icon"
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

        # Check YubiKey certificate presence as a proxy for YubiKey status
        $yubiKeyCert = $YubiKeyCertExpiryText.Text  # Use current UI text to avoid direct call
        $yubikeyStatus = $yubiKeyCert -notmatch "Unable to determine expiry date"

        if ($antivirusStatus -and $bitlockerStatus -and $yubikeyStatus -and $code42Status -and $fipsStatus -and $bigfixStatus) {
            $TrayIcon.Icon = Get-Icon -Path $config.IconPaths.Healthy -DefaultIcon ([System.Drawing.SystemIcons]::Application)
            $TrayIcon.Text = "System Monitor - Healthy"
        }
        else {
            $TrayIcon.Icon = Get-Icon -Path $config.IconPaths.Warning -DefaultIcon ([System.Drawing.SystemIcons]::Application)
            $TrayIcon.Text = "System Monitor - Warning"
        }

        $AntivirusStatusText.Text = $antivirusMessage
        $BitLockerStatusText.Text = $bitlockerMessage
        $BigFixStatusText.Text    = $bigfixMessage
        $Code42StatusText.Text    = $code42Message
        $FIPSStatusText.Text      = $fipsMessage

        if ($bitlockerStatus) {
            $BitLockerBorder.BorderBrush = 'Green'
        }
        else {
            $BitLockerBorder.BorderBrush = 'Red'
        }

        if ($bigfixStatus) {
            $BigFixExpander.Foreground = 'Green'
            $BigFixBorder.BorderBrush  = 'Green'
        }
        else {
            $BigFixExpander.Foreground = 'Red'
            $BigFixBorder.BorderBrush  = 'Red'
        }

        if ($code42Status) {
            $Code42Border.BorderBrush = 'Green'
        }
        else {
            $Code42Border.BorderBrush = 'Red'
        }
        
        if ($fipsStatus) {
            $FIPSBorder.BorderBrush = 'Green'
        }
        else {
            $FIPSBorder.BorderBrush = 'Red'
        }

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
            $LogListView.ItemsSource = $logEntries
        }
        else {
            $LogListView.ItemsSource = @([PSCustomObject]@{Timestamp="N/A"; Message="Log file not found."})
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
        $saveFileDialog.FileName = "SystemMonitor.log"
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
        $window.UpdateLayout()
        $primary = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
        $window.Left = $primary.X + $primary.Width - $window.ActualWidth - 10
        $window.Top  = $primary.Y + $primary.Height - $window.ActualHeight - 50
    }
    catch {
        Handle-Error "Error setting window position: $_" -Source "Set-WindowPosition"
    }
}

function Toggle-WindowVisibility {
    try {
        if ($window.Visibility -eq 'Visible') {
            $window.Hide()
            Write-Log "Dashboard hidden via Toggle-WindowVisibility." -Level "INFO"
        }
        else {
            Set-WindowPosition
            $window.Show()
            Write-Log "Dashboard shown via Toggle-WindowVisibility." -Level "INFO"
        }
    }
    catch {
        Handle-Error "Error toggling window visibility: $_" -Source "Toggle-WindowVisibility"
    }
}

# ========================
# L) Button Event Handlers
# ========================
$ExportLogsButton.Add_Click({ Export-Logs })

# Hyperlink Click Handlers
$AnnouncementsLink1.Add_RequestNavigate({ param($sender, $e) Start-Process $e.Uri.AbsoluteUri; $e.Handled = $true })
$AnnouncementsLink2.Add_RequestNavigate({ param($sender, $e) Start-Process $e.Uri.AbsoluteUri; $e.Handled = $true })
$SupportLink1.Add_RequestNavigate({ param($sender, $e) Start-Process $e.Uri.AbsoluteUri; $e.Handled = $true })
$SupportLink2.Add_RequestNavigate({ param($sender, $e) Start-Process $e.Uri.AbsoluteUri; $e.Handled = $true })
$EarlyAdopterLink1.Add_RequestNavigate({ param($sender, $e) Start-Process $e.Uri.AbsoluteUri; $e.Handled = $true })
$EarlyAdopterLink2.Add_RequestNavigate({ param($sender, $e) Start-Process $e.Uri.AbsoluteUri; $e.Handled = $true })

# ========================
# M) Create & Configure Tray Icon
# ========================
$TrayIcon = New-Object System.Windows.Forms.NotifyIcon
$TrayIcon.Icon = Get-Icon -Path $config.IconPaths.Healthy -DefaultIcon ([System.Drawing.SystemIcons]::Application)
$TrayIcon.Visible = $true
$TrayIcon.Text = "System Monitor"

# ========================
# N) Tray Icon Context Menu
# ========================
$ContextMenu = New-Object System.Windows.Forms.ContextMenu
$MenuItemShow = New-Object System.Windows.Forms.MenuItem("Show Dashboard")
$MenuItemExit = New-Object System.Windows.Forms.MenuItem("Exit")
$ContextMenu.MenuItems.Add($MenuItemShow)
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
        Update-TrayIcon
        Update-SystemInfo
        Update-Logs
        Update-Announcements
        Update-PatchingUpdates
        Update-Support
        Update-EarlyAdopterTesting
        Update-Compliance
    }
    catch {
        Handle-Error "Error during timer tick: $_" -Source "DispatcherTimer"
    }
})
$dispatcherTimer.Start()

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
    $window.Dispatcher.Invoke([Action]{
        Update-SystemInfo
        Update-TrayIcon
        Update-Logs
        Update-Announcements
        Update-PatchingUpdates
        Update-Support
        Update-EarlyAdopterTesting
        Update-Compliance
    })
    Log-DotNetVersion
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
