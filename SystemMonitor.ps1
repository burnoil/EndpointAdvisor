###############################################################################
# SystemMonitor.ps1 - Revised Version with Advanced Logging, External Config,
# DispatcherTimer, Enhanced Log Viewer, Code42 Service Check, About Section,
# Auto-Sizing, Anchored to Bottom Right, FIPS Compliance Detection, and a 
# More Compact UI Layout
###############################################################################

# Ensure $PSScriptRoot is defined for older versions.
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
    # Default settings
    $config = @{
        RefreshInterval    = 30         # in seconds
        LogRotationSizeMB  = 5          # Maximum log file size in MB
        DefaultLogLevel    = "INFO"
        IconPaths          = @{
            Healthy = (Join-Path $ScriptDir "healthy.ico")
            Warning = (Join-Path $ScriptDir "warning.ico")
        }
    }
    $config | ConvertTo-Json | Out-File $configPath -Force
}

# ========================
# B) Log File Setup & Rotation
# ========================
$LogFilePath = Join-Path $ScriptDir "SystemMonitor.log"
$LogDirectory  = Split-Path $LogFilePath
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
    # Additional error handling (e.g., user notifications) can be added here.
}

# Stubbed out: no toast notifications.
function Show-Notification {
    param(
        [string]$Title,
        [string]$Message,
        [System.Drawing.Icon]$Icon = [System.Drawing.SystemIcons]::Information
    )
    Write-Log "Notification suppressed: $Title - $Message" -Level "INFO"
}

# ========================
# D) Import Required Assemblies
# ========================
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ========================
# E) XAML Layout Definition (Compact UI, with Enhanced Log Viewer, Code42, FIPS & About)
# ========================
[xml]$xaml = @"
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
        <Image Source="$($config.IconPaths.Healthy)" Width="20" Height="20" Margin="0,0,4,0"/>
        <TextBlock Text="System Monitoring Dashboard"
                   FontSize="14" FontWeight="Bold" Foreground="White"
                   VerticalAlignment="Center"/>
      </StackPanel>
    </Border>
    <!-- Content Area -->
    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
      <StackPanel VerticalAlignment="Top">
        <!-- System Information Section -->
        <Expander Header="System Information" FontSize="12" Foreground="#0078D7" IsExpanded="True" Margin="0,2,0,2">
          <Border BorderBrush="#0078D7" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="LoggedOnUserText" FontSize="11" Margin="2" TextWrapping="Wrap"/>
              <TextBlock x:Name="MachineTypeText" FontSize="11" Margin="2" TextWrapping="Wrap"/>
              <TextBlock x:Name="OSVersionText" FontSize="11" Margin="2" TextWrapping="Wrap"/>
              <TextBlock x:Name="SystemUptimeText" FontSize="11" Margin="2" TextWrapping="Wrap"/>
              <TextBlock x:Name="UsedDiskSpaceText" FontSize="11" Margin="2" TextWrapping="Wrap"/>
              <TextBlock x:Name="IpAddressText" FontSize="11" Margin="2" TextWrapping="Wrap"/>
            </StackPanel>
          </Border>
        </Expander>
        <!-- Antivirus Section -->
        <Expander Header="Antivirus Information" FontSize="12" Foreground="#28a745" IsExpanded="True" Margin="0,2,0,2">
          <Border BorderBrush="#28a745" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="AntivirusStatusText" FontSize="11" Margin="2" TextWrapping="Wrap"/>
            </StackPanel>
          </Border>
        </Expander>
        <!-- BitLocker Section -->
        <Expander x:Name="BitLockerExpander" Header="BitLocker Information" FontSize="12" Foreground="#6c757d" IsExpanded="True" Margin="0,2,0,2">
          <Border x:Name="BitLockerBorder" BorderBrush="#6c757d" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="BitLockerStatusText" FontSize="11" Margin="2" TextWrapping="Wrap"/>
            </StackPanel>
          </Border>
        </Expander>
        <!-- YubiKey Section -->
        <Expander x:Name="YubiKeyExpander" Header="YubiKey Information" FontSize="12" Foreground="#FF69B4" IsExpanded="True" Margin="0,2,0,2">
          <Border x:Name="YubiKeyBorder" BorderBrush="#FF69B4" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="YubiKeyStatusText" FontSize="11" Margin="2" TextWrapping="Wrap"/>
            </StackPanel>
          </Border>
        </Expander>
        <!-- BigFix Section -->
        <Expander x:Name="BigFixExpander" Header="BigFix (BESClient)" FontSize="12" Foreground="#4b0082" IsExpanded="True" Margin="0,2,0,2">
          <Border x:Name="BigFixBorder" BorderBrush="#4b0082" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="BigFixStatusText" FontSize="11" Margin="2" TextWrapping="Wrap"/>
            </StackPanel>
          </Border>
        </Expander>
        <!-- Code42 Service Section -->
        <Expander x:Name="Code42Expander" Header="Code42 Service" FontSize="12" Foreground="#800080" IsExpanded="True" Margin="0,2,0,2">
          <Border x:Name="Code42Border" BorderBrush="#800080" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="Code42StatusText" FontSize="11" Margin="2" TextWrapping="Wrap"/>
            </StackPanel>
          </Border>
        </Expander>
        <!-- FIPS Compliance Section -->
        <Expander x:Name="FIPSExpander" Header="FIPS Compliance" FontSize="12" Foreground="#FF4500" IsExpanded="True" Margin="0,2,0,2">
          <Border x:Name="FIPSBorder" BorderBrush="#FF4500" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="FIPSStatusText" FontSize="11" Margin="2" TextWrapping="Wrap"/>
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
              <TextBlock x:Name="AboutText" FontSize="11" Margin="2" TextWrapping="Wrap"
                         Text="System Monitor v1.0`n© 2025 System Monitor. All rights reserved.`nBuilt with PowerShell and WPF."/>
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

# ========================
# F) Load and Verify XAML
# ========================
$reader = New-Object System.Xml.XmlNodeReader($xaml)
try {
    $window = [Windows.Markup.XamlReader]::Load($reader)
}
catch {
    Handle-Error "Failed to load the XAML layout. Error: $_" -Source "XAML"
    return
}
if ($window -eq $null) {
    Handle-Error "Failed to load the XAML layout. Check the XAML syntax for errors." -Source "XAML"
    return
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

$AntivirusStatusText = $window.FindName("AntivirusStatusText")
$BitLockerStatusText = $window.FindName("BitLockerStatusText")
$YubiKeyStatusText   = $window.FindName("YubiKeyStatusText")
$BigFixStatusText    = $window.FindName("BigFixStatusText")
$Code42StatusText    = $window.FindName("Code42StatusText")
$FIPSStatusText      = $window.FindName("FIPSStatusText")
$AboutText           = $window.FindName("AboutText")

$LogListView         = $window.FindName("LogListView")
$ExportLogsButton    = $window.FindName("ExportLogsButton")

$BitLockerExpander   = $window.FindName("BitLockerExpander")
$BitLockerBorder     = $window.FindName("BitLockerBorder")
$YubiKeyExpander     = $window.FindName("YubiKeyExpander")
$YubiKeyBorder       = $window.FindName("YubiKeyBorder")
$BigFixExpander      = $window.FindName("BigFixExpander")
$BigFixBorder        = $window.FindName("BigFixBorder")
$Code42Expander      = $window.FindName("Code42Expander")
$Code42Border        = $window.FindName("Code42Border")
$FIPSExpander        = $window.FindName("FIPSExpander")
$FIPSBorder          = $window.FindName("FIPSBorder")

# ========================
# H) Modularized System Information Functions
# ========================
function Update-SystemInfo {
    try {
        $user = [System.Environment]::UserName
        $LoggedOnUserText.Text = "Logged-in User: $user"
        Write-Log "Logged-in User: $user" -Level "INFO"

        $machine = Get-CimInstance -ClassName Win32_ComputerSystem
        $machineType = "$($machine.Manufacturer) $($machine.Model)"
        $MachineTypeText.Text = "Machine Type: $machineType"
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
        $OSVersionText.Text = "OS Version: $osVersion"
        Write-Log "OS Version: $osVersion" -Level "INFO"

        $uptime = (Get-Date) - $os.LastBootUpTime
        $systemUptime = "$([math]::Floor($uptime.TotalDays)) days $($uptime.Hours) hours"
        $SystemUptimeText.Text = "System Uptime: $systemUptime"
        Write-Log "System Uptime: $systemUptime" -Level "INFO"

        $drive = Get-PSDrive -Name C
        $usedDiskSpace = "$([math]::Round(($drive.Used / 1GB), 2)) GB of $([math]::Round((($drive.Free + $drive.Used) / 1GB), 2)) GB"
        $UsedDiskSpaceText.Text = "Used Disk Space: $usedDiskSpace"
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

function Get-YubiKeyStatus {
    Write-Log "Starting YubiKey detection..." -Level "INFO"
    $yubicoVendorID = "1050"
    $yubikeyProductIDs = @("0407","0408","0409","040A","040B","040C","040D","040E")
    try {
        $allYubicoDevices = Get-PnpDevice -Class USB | Where-Object {
            ($_.InstanceId -match "VID_$yubicoVendorID") -and ($_.Status -eq "OK")
        }
        Write-Log "Found $($allYubicoDevices.Count) Yubico USB device(s) with Status='OK'." -Level "INFO"
        foreach ($device in $allYubicoDevices) {
            Write-Log "Detected Device: $($device.FriendlyName) - InstanceId: $($device.InstanceId)" -Level "INFO"
        }
        $detectedYubiKeys = $allYubicoDevices | Where-Object {
            foreach ($productId in $yubikeyProductIDs) {
                if ($_.InstanceId -match "PID_$productId") { return $true }
            }
            return $false
        }
        if ($detectedYubiKeys) {
            $friendlyNames = $detectedYubiKeys | ForEach-Object { $_.FriendlyName } | Sort-Object -Unique
            $statusMessage = "YubiKey Detected: $($friendlyNames -join ', ')"
            Write-Log $statusMessage -Level "INFO"
            return $true, $statusMessage
        }
        else {
            $statusMessage = "No YubiKey Detected."
            Write-Log $statusMessage -Level "INFO"
            return $false, $statusMessage
        }
    }
    catch {
        Write-Log "Error during YubiKey detection: $_" -Level "ERROR"
        return $false, "Error detecting YubiKey."
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
        $yubikeyStatus,  $yubikeyMessage   = Get-YubiKeyStatus
        $bigfixStatus,   $bigfixMessage    = Get-BigFixStatus
        $code42Status,   $code42Message    = Get-Code42Status
        $fipsStatus,     $fipsMessage      = Get-FIPSStatus

        if ($antivirusStatus -and $bitlockerStatus -and $yubikeyStatus -and $code42Status) {
            $TrayIcon.Icon = Get-Icon -Path $config.IconPaths.Healthy -DefaultIcon ([System.Drawing.SystemIcons]::Application)
            $TrayIcon.Text = "System Monitor - Healthy"
        }
        else {
            $TrayIcon.Icon = Get-Icon -Path $config.IconPaths.Warning -DefaultIcon ([System.Drawing.SystemIcons]::Application)
            $TrayIcon.Text = "System Monitor - Warning"
        }

        $AntivirusStatusText.Text = $antivirusMessage
        $BitLockerStatusText.Text = $bitlockerMessage
        $YubiKeyStatusText.Text   = $yubikeyMessage
        $BigFixStatusText.Text    = $bigfixMessage
        $Code42StatusText.Text    = $code42Message
        $FIPSStatusText.Text      = $fipsMessage

        if ($bitlockerStatus) {
            $BitLockerExpander.Foreground = 'Green'
            $BitLockerBorder.BorderBrush = 'Green'
        }
        else {
            $BitLockerExpander.Foreground = 'Red'
            $BitLockerBorder.BorderBrush = 'Red'
        }

        if ($yubikeyStatus) {
            $YubiKeyExpander.Foreground = 'Green'
            $YubiKeyBorder.BorderBrush = 'Green'
        }
        else {
            $YubiKeyExpander.Foreground = 'Red'
            $YubiKeyBorder.BorderBrush = 'Red'
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
            $Code42Expander.Foreground = 'Green'
            $Code42Border.BorderBrush  = 'Green'
        }
        else {
            $Code42Expander.Foreground = 'Red'
            $Code42Border.BorderBrush  = 'Red'
        }
        
        if ($fipsStatus) {
            $FIPSExpander.Foreground = 'Green'
            $FIPSBorder.BorderBrush = 'Green'
        }
        else {
            $FIPSExpander.Foreground = 'Red'
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
# K) Window Visibility Management (Anchor to Bottom Right)
# ========================
function Set-WindowPosition {
    try {
        $window.UpdateLayout()
        $screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
        $window.Left = $screen.Width - $window.ActualWidth - 10
        $window.Top  = $screen.Height - $window.ActualHeight - 50
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
        Update-YubiKeyStatus
    }
    catch {
        Handle-Error "Error during timer tick: $_" -Source "DispatcherTimer"
    }
})
$dispatcherTimer.Start()

function Update-YubiKeyStatus {
    try {
        $yubikeyPresent, $yubikeyMessage = Get-YubiKeyStatus
        $YubiKeyStatusText.Text = $yubikeyMessage
        if ($yubikeyPresent) {
            Write-Log "YubiKey is present." -Level "INFO"
        }
        else {
            Write-Log "YubiKey is not present." -Level "INFO"
        }
    }
    catch {
        $YubiKeyStatusText.Text = "Error detecting YubiKey."
        Handle-Error "Error updating YubiKey status: $_" -Source "Update-YubiKeyStatus"
    }
}

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
        Update-YubiKeyStatus
    })
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
