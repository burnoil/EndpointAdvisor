###############################################################################
# SystemMonitor.ps1
# - COM-based BitLocker detection for non-admin
# - No toast notifications
# - Clean dispatcher shutdown
# - IP address display
# - BigFix (BESClient) in its own Expander
# - Condensed UI (smaller margins, no settings expander)
# - Fixed 30s refresh interval
# - OS DisplayVersion detection (e.g. 22H2) appended to OS
# - Icon files loaded from the same folder as this script
###############################################################################

# Ensure we can reference $PSScriptRoot in older PS versions if needed
# (In PowerShell 3+, $PSScriptRoot works automatically when script is saved as .ps1)
if (-not $PSScriptRoot) {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    $ScriptDir = $PSScriptRoot
}

# ========================
# 1) Import Required Assemblies
# ========================
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ========================
# 2) Configuration Variables
# ========================
$GreenIconPath = Join-Path $ScriptDir "healthy.ico"    # Green icon for healthy status
$RedIconPath   = Join-Path $ScriptDir "warning.ico"    # Red icon for warning status
$LogoImagePath = Join-Path $ScriptDir "icon.png"       # Optional: Logo/Icon for the dashboard

# Log file (same folder or specify another path if needed)
$LogFilePath   = Join-Path $ScriptDir "SystemMonitor.log"

# Ensure the log directory exists
$LogDirectory = Split-Path $LogFilePath
if (-not (Test-Path $LogDirectory)) {
    New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
}

# ========================
# 3) Logging & No-Op Notification
# ========================
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogFilePath -Value "[$timestamp] $Message"
}

# Stubbed out: no toast or balloon tips.
function Show-Notification {
    param(
        [string]$Title,
        [string]$Message,
        [System.Drawing.Icon]$Icon = [System.Drawing.SystemIcons]::Information
    )
    Write-Log "Notification suppressed: $Title - $Message"
}

# ========================
# 4) XAML Layout Definition (Condensed, no Settings)
# ========================
[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="System Monitor"
    Height="500"
    Width="350"
    ResizeMode="CanResize"
    ShowInTaskbar="False"
    Visibility="Hidden"
    Topmost="True"
    Background="#f0f0f0">

    <Grid Margin="5">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Title Section -->
        <Border Grid.Row="0" Background="#0078D7" Padding="8" CornerRadius="3" Margin="0,0,0,5">
            <StackPanel Orientation="Horizontal" VerticalAlignment="Center" HorizontalAlignment="Center">
                <!-- Replace with a dynamic binding to $LogoImagePath if needed,
                     but typically you can reference the absolute path or just omit the image. -->
                <Image Source="$LogoImagePath" Width="24" Height="24" Margin="0,0,8,0"/>
                <TextBlock Text="System Monitoring Dashboard"
                           FontSize="16" FontWeight="Bold" Foreground="White"
                           VerticalAlignment="Center"/>
            </StackPanel>
        </Border>

        <!-- Content Area -->
        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
            <StackPanel VerticalAlignment="Top">

                <!-- System Information Section -->
                <Expander Header="System Information"
                          FontSize="13" Foreground="#0078D7"
                          IsExpanded="True" Margin="0,0,0,5">
                    <Border BorderBrush="#0078D7" BorderThickness="1"
                            Padding="5" CornerRadius="3" Background="White" Margin="3">
                        <StackPanel Orientation="Vertical">
                            <TextBlock x:Name="LoggedOnUserText"  FontSize="12" Margin="3" TextWrapping="Wrap"/>
                            <TextBlock x:Name="MachineTypeText"   FontSize="12" Margin="3" TextWrapping="Wrap"/>
                            <TextBlock x:Name="OSVersionText"     FontSize="12" Margin="3" TextWrapping="Wrap"/>
                            <TextBlock x:Name="SystemUptimeText"  FontSize="12" Margin="3" TextWrapping="Wrap"/>
                            <TextBlock x:Name="UsedDiskSpaceText" FontSize="12" Margin="3" TextWrapping="Wrap"/>
                            <TextBlock x:Name="IpAddressText"     FontSize="12" Margin="3" TextWrapping="Wrap"/>
                        </StackPanel>
                    </Border>
                </Expander>

                <!-- Antivirus Section -->
                <Expander Header="Antivirus Information"
                          FontSize="13" Foreground="#28a745"
                          IsExpanded="True" Margin="0,0,0,5">
                    <Border BorderBrush="#28a745" BorderThickness="1"
                            Padding="5" CornerRadius="3" Background="White" Margin="3">
                        <StackPanel Orientation="Vertical">
                            <TextBlock x:Name="AntivirusStatusText" FontSize="12" Margin="3" TextWrapping="Wrap"/>
                        </StackPanel>
                    </Border>
                </Expander>

                <!-- BitLocker Section -->
                <Expander x:Name="BitLockerExpander" Header="BitLocker Information"
                          FontSize="13" Foreground="#6c757d"
                          IsExpanded="True" Margin="0,0,0,5">
                    <Border x:Name="BitLockerBorder" BorderBrush="#6c757d" BorderThickness="1"
                            Padding="5" CornerRadius="3" Background="White" Margin="3">
                        <StackPanel Orientation="Vertical">
                            <TextBlock x:Name="BitLockerStatusText" FontSize="12" Margin="3" TextWrapping="Wrap"/>
                        </StackPanel>
                    </Border>
                </Expander>

                <!-- YubiKey Section -->
                <Expander x:Name="YubiKeyExpander" Header="YubiKey Information"
                          FontSize="13" Foreground="#FF69B4"
                          IsExpanded="True" Margin="0,0,0,5">
                    <Border x:Name="YubiKeyBorder" BorderBrush="#FF69B4" BorderThickness="1"
                            Padding="5" CornerRadius="3" Background="White" Margin="3">
                        <StackPanel Orientation="Vertical">
                            <TextBlock x:Name="YubiKeyStatusText" FontSize="12" Margin="3" TextWrapping="Wrap"/>
                        </StackPanel>
                    </Border>
                </Expander>

                <!-- BigFix Section -->
                <Expander x:Name="BigFixExpander" Header="BigFix (BESClient)"
                          FontSize="13" Foreground="#4b0082"
                          IsExpanded="True" Margin="0,0,0,5">
                    <Border x:Name="BigFixBorder" BorderBrush="#4b0082" BorderThickness="1"
                            Padding="5" CornerRadius="3" Background="White" Margin="3">
                        <StackPanel Orientation="Vertical">
                            <TextBlock x:Name="BigFixStatusText" FontSize="12" Margin="3" TextWrapping="Wrap"/>
                        </StackPanel>
                    </Border>
                </Expander>

                <!-- Logs Section -->
                <Expander Header="Logs"
                          FontSize="13" Foreground="#ff8c00"
                          IsExpanded="False" Margin="0,0,0,5">
                    <Border BorderBrush="#ff8c00" BorderThickness="1"
                            Padding="5" CornerRadius="3" Background="White" Margin="3">
                        <StackPanel Orientation="Vertical">
                            <TextBox x:Name="LogTextBox" FontSize="10" Margin="3"
                                     Height="150" IsReadOnly="True" TextWrapping="Wrap"
                                     VerticalScrollBarVisibility="Auto"/>
                            <Button x:Name="ExportLogsButton" Content="Export Logs"
                                    Width="90" Margin="3" HorizontalAlignment="Right"/>
                        </StackPanel>
                    </Border>
                </Expander>

            </StackPanel>
        </ScrollViewer>

        <!-- Footer Section -->
        <TextBlock Grid.Row="2"
                   Text="© 2025 System Monitor"
                   FontSize="10" Foreground="Gray"
                   HorizontalAlignment="Center"
                   Margin="0,5,0,0"/>
    </Grid>
</Window>
"@

# ========================
# 5) Load and Verify XAML
# ========================
$reader = New-Object System.Xml.XmlNodeReader($xaml)
try {
    $window = [Windows.Markup.XamlReader]::Load($reader)
}
catch {
    Write-Log "Failed to load the XAML layout. Error: $_"
    return
}
if ($window -eq $null) {
    Write-Log "Failed to load the XAML layout. Check the XAML syntax for errors."
    return
}

# ========================
# 6) Access UI Elements
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

$LogTextBox          = $window.FindName("LogTextBox")
$ExportLogsButton    = $window.FindName("ExportLogsButton")

$BitLockerExpander = $window.FindName("BitLockerExpander")
$BitLockerBorder   = $window.FindName("BitLockerBorder")
$YubiKeyExpander   = $window.FindName("YubiKeyExpander")
$YubiKeyBorder     = $window.FindName("YubiKeyBorder")
$BigFixExpander    = $window.FindName("BigFixExpander")
$BigFixBorder      = $window.FindName("BigFixBorder")

# ========================
# 7) System Information Functions
# ========================
function Update-SystemInfo {
    try {
        # Logged-on user
        $user = [System.Environment]::UserName
        $LoggedOnUserText.Text = "Logged-in User: $user"
        Write-Log "Logged-in User: $user"

        # Machine info
        $machine = Get-CimInstance -ClassName Win32_ComputerSystem
        $machineType = "$($machine.Manufacturer) $($machine.Model)"
        $MachineTypeText.Text = "Machine Type: $machineType"
        Write-Log "Machine Type: $machineType"

        # OS version (with DisplayVersion e.g. "22H2")
        $os = Get-CimInstance -ClassName Win32_OperatingSystem
        $osVersion = "$($os.Caption) (Build $($os.BuildNumber))"

        try {
            # Attempt to read 'DisplayVersion' e.g. "22H2"
            $displayVersion = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name 'DisplayVersion' -ErrorAction SilentlyContinue).DisplayVersion
            if ($displayVersion) {
                $osVersion += " $displayVersion"
            }
        }
        catch {
            Write-Log "Could not retrieve DisplayVersion from registry: $_"
        }

        $OSVersionText.Text = "OS Version: $osVersion"
        Write-Log "OS Version: $osVersion"

        # Uptime
        $uptime = (Get-Date) - $os.LastBootUpTime
        $systemUptime = "$([math]::Floor($uptime.TotalDays)) days $($uptime.Hours) hours"
        $SystemUptimeText.Text = "System Uptime: $systemUptime"
        Write-Log "System Uptime: $systemUptime"

        # Disk usage
        $drive = Get-PSDrive -Name C
        $usedDiskSpace = "$([math]::Round(($drive.Used / 1GB), 2)) GB of $([math]::Round(($drive.Free + $drive.Used) / 1GB, 2)) GB"
        $UsedDiskSpaceText.Text = "Used Disk Space: $usedDiskSpace"
        Write-Log "Used Disk Space: $usedDiskSpace"

        # IP addresses
        $ipv4s = Get-NetIPAddress -AddressFamily IPv4 | Where-Object {
            $_.IPAddress -notlike "127.*" -and
            $_.IPAddress -notlike "169.254.*" -and
            $_.IPAddress -notin @("0.0.0.0","255.255.255.255") -and
            $_.PrefixOrigin -ne "WellKnown"
        } | Select-Object -ExpandProperty IPAddress -ErrorAction SilentlyContinue

        if ($ipv4s) {
            $ipList = $ipv4s -join ", "
            $IpAddressText.Text = "IPv4 Address(es): $ipList"
            Write-Log "IP Address(es): $ipList"
        }
        else {
            $IpAddressText.Text = "IPv4 Address(es): None detected"
            Write-Log "No valid IPv4 addresses found."
        }
    }
    catch {
        Write-Log "Error updating system information: $_"
    }
}

function Get-BigFixStatus {
    try {
        $besService = Get-Service -Name BESClient -ErrorAction SilentlyContinue
        if ($besService) {
            if ($besService.Status -eq 'Running') {
                return $true,  "BigFix (BESClient) Service: Running"
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
    Write-Log "Starting YubiKey detection..."
    $yubicoVendorID = "1050"
    $yubikeyProductIDs = @("0407","0408","0409","040A","040B","040C","040D","040E")
    try {
        $allYubicoDevices = Get-PnpDevice -Class USB | Where-Object {
            ($_.InstanceId -match "VID_$yubicoVendorID") -and ($_.Status -eq "OK")
        }

        Write-Log "Found $($allYubicoDevices.Count) Yubico USB device(s) with Status='OK'."
        foreach ($device in $allYubicoDevices) {
            Write-Log "Detected Device: $($device.FriendlyName) - InstanceId: $($device.InstanceId)"
        }

        $detectedYubiKeys = $allYubicoDevices | Where-Object {
            foreach ($productId in $yubikeyProductIDs) {
                if ($_.InstanceId -match "PID_$productId") {
                    return $true
                }
            }
            return $false
        }

        if ($detectedYubiKeys) {
            $friendlyNames = $detectedYubiKeys | ForEach-Object { $_.FriendlyName } | Sort-Object -Unique
            $statusMessage = "YubiKey Detected: $($friendlyNames -join ', ')"
            Write-Log $statusMessage
            return $true, $statusMessage
        }
        else {
            $statusMessage = "No YubiKey Detected."
            Write-Log $statusMessage
            return $false, $statusMessage
        }
    }
    catch {
        Write-Log "Error during YubiKey detection: $_"
        return $false, "Error detecting YubiKey."
    }
}

# ========================
# 8) Tray Icon Management
# ========================
function Get-Icon {
    param(
        [string]$Path,
        [System.Drawing.Icon]$DefaultIcon
    )
    if (-not (Test-Path $Path)) {
        Write-Log "$Path not found. Using default icon."
        return $DefaultIcon
    }
    else {
        try {
            $icon = New-Object System.Drawing.Icon($Path)
            Write-Log "Custom icon loaded from $($Path)."
            return $icon
        }
        catch {
            Write-Log "Error loading icon from $($Path): $($_). Using default icon."
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

        if ($antivirusStatus -and $bitlockerStatus -and $yubikeyStatus) {
            $TrayIcon.Icon = Get-Icon -Path $GreenIconPath -DefaultIcon ([System.Drawing.SystemIcons]::Application)
            $TrayIcon.Text = "System Monitor - Healthy"
        }
        else {
            $TrayIcon.Icon = Get-Icon -Path $RedIconPath -DefaultIcon ([System.Drawing.SystemIcons]::Application)
            $TrayIcon.Text = "System Monitor - Warning"
        }

        # Update text blocks
        $AntivirusStatusText.Text = $antivirusMessage
        $BitLockerStatusText.Text = $bitlockerMessage
        $YubiKeyStatusText.Text   = $yubikeyMessage
        $BigFixStatusText.Text    = $bigfixMessage

        # BitLocker color
        if ($bitlockerStatus) {
            $BitLockerExpander.Foreground = 'Green'
            $BitLockerBorder.BorderBrush = 'Green'
        }
        else {
            $BitLockerExpander.Foreground = 'Red'
            $BitLockerBorder.BorderBrush = 'Red'
        }

        # YubiKey color
        if ($yubikeyStatus) {
            $YubiKeyExpander.Foreground = 'Green'
            $YubiKeyBorder.BorderBrush = 'Green'
        }
        else {
            $YubiKeyExpander.Foreground = 'Red'
            $YubiKeyBorder.BorderBrush = 'Red'
        }

        # BigFix color
        if ($bigfixStatus) {
            $BigFixExpander.Foreground = 'Green'
            $BigFixBorder.BorderBrush  = 'Green'
        }
        else {
            $BigFixExpander.Foreground = 'Red'
            $BigFixBorder.BorderBrush  = 'Red'
        }

        Write-Log "Tray icon and status updated."
    }
    catch {
        Write-Log "Error updating tray icon: $_"
    }
}

# ========================
# 9) Logs Management
# ========================
function Update-Logs {
    try {
        if (Test-Path $LogFilePath) {
            $LogContent = Get-Content -Path $LogFilePath -Tail 100 -ErrorAction SilentlyContinue
            $LogTextBox.Text = $LogContent -join "`n"
        }
        else {
            $LogTextBox.Text = "Log file not found."
        }
        Write-Log "Logs updated in GUI."
    }
    catch {
        $LogTextBox.Text = "Error loading logs: $_"
        Write-Log "Error loading logs: $_"
    }
}

function Export-Logs {
    try {
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "Log Files (*.log)|*.log|All Files (*.*)|*.*"
        $saveFileDialog.FileName = "SystemMonitor.log"
        if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Copy-Item -Path $LogFilePath -Destination $saveFileDialog.FileName -Force
            Write-Log "Logs exported to $($saveFileDialog.FileName)"
        }
    }
    catch {
        Write-Log "Error exporting logs: $_"
    }
}

# ========================
# 10) Window Visibility Management
# ========================
function Toggle-WindowVisibility {
    try {
        if ($window.Visibility -eq 'Visible') {
            $window.Hide()
            Write-Log "Dashboard hidden via Toggle-WindowVisibility."
        }
        else {
            $screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
            $window.Left = $screen.Width - $window.Width - 10
            $window.Top  = $screen.Height - $window.Height - 50
            $window.Show()
            Write-Log "Dashboard shown via Toggle-WindowVisibility."
        }
    }
    catch {
        Write-Log "Error toggling window visibility: $_"
    }
}

# ========================
# 11) Button Event Handlers
# ========================
$ExportLogsButton.Add_Click({ Export-Logs })

# ========================
# 12) Create & Configure Tray Icon
# ========================
$TrayIcon = New-Object System.Windows.Forms.NotifyIcon
$TrayIcon.Icon = Get-Icon -Path $GreenIconPath -DefaultIcon ([System.Drawing.SystemIcons]::Application)
$TrayIcon.Visible = $true
$TrayIcon.Text = "System Monitor"

# ========================
# 13) Tray Icon Context Menu
# ========================
$ContextMenu    = New-Object System.Windows.Forms.ContextMenu
$MenuItemShow   = New-Object System.Windows.Forms.MenuItem("Show Dashboard")
$MenuItemExit   = New-Object System.Windows.Forms.MenuItem("Exit")
$ContextMenu.MenuItems.Add($MenuItemShow)
$ContextMenu.MenuItems.Add($MenuItemExit)
$TrayIcon.ContextMenu = $ContextMenu

# Toggle GUI on tray left-click
$TrayIcon.add_MouseClick({
    param($sender,$e)
    try {
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            Toggle-WindowVisibility
        }
    }
    catch {
        Write-Log "Error handling tray icon mouse click: $_"
    }
})

$MenuItemShow.add_Click({
    Toggle-WindowVisibility
})

$MenuItemExit.add_Click({
    try {
        Write-Log "Exit clicked by user."
        $timer.Stop()
        Write-Log "Timer stopped."
        $TrayIcon.Dispose()
        Write-Log "Tray icon disposed."
        $window.Dispatcher.InvokeShutdown()
        Write-Log "Application exited via tray menu."
    }
    catch {
        Write-Log "Error during application exit: $_"
    }
})

# ========================
# 14) Timer for Periodic Updates (Fixed Interval 30s)
# ========================
$timer = New-Object System.Windows.Forms.Timer
$defaultInterval = 30
$timer.Interval = $defaultInterval * 1000  # 30s
$timer.Add_Tick({
    try {
        $window.Dispatcher.Invoke([Action]{
            Update-TrayIcon
            Update-SystemInfo
            Update-Logs
            Update-YubiKeyStatus
        })
    }
    catch {
        Write-Log "Error during timer tick: $_"
    }
})
$timer.Start()

# ========================
# 15) YubiKey Update Function
# ========================
function Update-YubiKeyStatus {
    try {
        $yubikeyPresent, $yubikeyMessage = Get-YubiKeyStatus
        $YubiKeyStatusText.Text = $yubikeyMessage

        if ($yubikeyPresent) {
            Write-Log "YubiKey is present."
        }
        else {
            Write-Log "YubiKey is not present."
        }
    }
    catch {
        $YubiKeyStatusText.Text = "Error detecting YubiKey."
        Write-Log "Error updating YubiKey status: $_"
    }
}

# ========================
# 16) Dispatcher Exception Handling
# ========================
function Handle-DispatcherUnhandledException {
    param(
        [object]$sender,
        [System.Windows.Threading.DispatcherUnhandledExceptionEventArgs]$args
    )
    Write-Log "Unhandled Dispatcher exception: $($args.Exception.Message)"
    # $args.Handled = $true
}

Register-ObjectEvent -InputObject $window.Dispatcher -EventName UnhandledException -Action {
    param($sender, $args)
    Handle-DispatcherUnhandledException -sender $sender -args $args
}

# ========================
# 17) Initialize the First Update
# ========================
try {
    $window.Dispatcher.Invoke([Action]{
        Update-SystemInfo
        Update-TrayIcon
        Update-Logs
        Update-YubiKeyStatus
    })
}
catch {
    Write-Log "Error during initial update: $_"
}

# ========================
# 18) Handle Window Closing (Hide Instead of Close)
# ========================
$window.Add_Closing({
    param($sender,$eventArgs)
    try {
        $eventArgs.Cancel = $true
        $window.Hide()
        Write-Log "Dashboard hidden via window closing event."
    }
    catch {
        Write-Log "Error handling window closing: $_"
    }
})

# ========================
# 19) Start the Application Dispatcher
# ========================
Write-Log "About to call Dispatcher.Run()..."
[System.Windows.Threading.Dispatcher]::Run()
Write-Log "Dispatcher ended; script exiting."
