###############################################################################
# SystemMonitor.ps1
# Final Corrected Version - COM-based BitLocker detection for non-admin
# - No toast notifications
# - Clean dispatcher shutdown
# - Added IP address display
###############################################################################

# ========================
# 1) Import Required Assemblies
# ========================
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ========================
# 2) Configuration Variables
# ========================
$GreenIconPath = "C:\Scripts\healthy.ico"  # Green icon for healthy status
$RedIconPath   = "C:\Scripts\warning.ico"  # Red icon for warning status
$LogoImagePath = "C:\Scripts\icon.png"     # Optional: Logo/Icon for the dashboard

# Log file
$LogFilePath = "C:\Scripts\SystemMonitor.log"

# Ensure the log directory exists
$LogDirectory = Split-Path $LogFilePath
if (-not (Test-Path $LogDirectory)) {
    New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
}

# ========================
# 3) Logging & (No-Op) Notification
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
    # We just log that a notification would have appeared
    Write-Log "Notification suppressed: $Title - $Message"
}

# ========================
# 4) XAML Layout Definition
# ========================
[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="System Monitor"
    Height="700"
    Width="400"
    ResizeMode="CanResize"
    ShowInTaskbar="False"
    Visibility="Hidden"
    Topmost="True"
    Background="#f0f0f0">

    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <!-- Title Section -->
        <Border Grid.Row="0" Background="#0078D7" Padding="10" CornerRadius="5" Margin="0,0,0,10">
            <StackPanel Orientation="Horizontal" VerticalAlignment="Center" HorizontalAlignment="Center">
                <Image Source="C:\Scripts\icon.png" Width="30" Height="30" Margin="0,0,10,0"/>
                <TextBlock Text="System Monitoring Dashboard"
                           FontSize="20" FontWeight="Bold" Foreground="White"
                           VerticalAlignment="Center"/>
            </StackPanel>
        </Border>

        <!-- Content Area -->
        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
            <StackPanel VerticalAlignment="Top">

                <!-- System Information Section -->
                <Expander Header="System Information"
                          FontSize="14" Foreground="#0078D7"
                          IsExpanded="True" Margin="0,0,0,10">
                    <Border BorderBrush="#0078D7" BorderThickness="1"
                            Padding="10" CornerRadius="5" Background="White" Margin="5">
                        <StackPanel Orientation="Vertical">
                            <TextBlock x:Name="LoggedOnUserText"  FontSize="12" Margin="5" TextWrapping="Wrap"/>
                            <TextBlock x:Name="MachineTypeText"   FontSize="12" Margin="5" TextWrapping="Wrap"/>
                            <TextBlock x:Name="OSVersionText"     FontSize="12" Margin="5" TextWrapping="Wrap"/>
                            <TextBlock x:Name="SystemUptimeText"  FontSize="12" Margin="5" TextWrapping="Wrap"/>
                            <TextBlock x:Name="UsedDiskSpaceText" FontSize="12" Margin="5" TextWrapping="Wrap"/>

                            <!-- Added IP Address TextBlock -->
                            <TextBlock x:Name="IpAddressText" FontSize="12" Margin="5" TextWrapping="Wrap"/>
                        </StackPanel>
                    </Border>
                </Expander>

                <!-- Antivirus Section -->
                <Expander Header="Antivirus Information"
                          FontSize="14" Foreground="#28a745"
                          IsExpanded="True" Margin="0,0,0,10">
                    <Border BorderBrush="#28a745" BorderThickness="1"
                            Padding="10" CornerRadius="5" Background="White" Margin="5">
                        <StackPanel Orientation="Vertical">
                            <TextBlock x:Name="AntivirusStatusText" FontSize="12" Margin="5" TextWrapping="Wrap"/>
                        </StackPanel>
                    </Border>
                </Expander>

                <!-- BitLocker Section (Named for dynamic color) -->
                <Expander x:Name="BitLockerExpander" Header="BitLocker Information"
                          FontSize="14" Foreground="#6c757d"
                          IsExpanded="True" Margin="0,0,0,10">
                    <Border x:Name="BitLockerBorder" BorderBrush="#6c757d" BorderThickness="1"
                            Padding="10" CornerRadius="5" Background="White" Margin="5">
                        <StackPanel Orientation="Vertical">
                            <TextBlock x:Name="BitLockerStatusText" FontSize="12" Margin="5" TextWrapping="Wrap"/>
                        </StackPanel>
                    </Border>
                </Expander>

                <!-- YubiKey Section (Named for dynamic color) -->
                <Expander x:Name="YubiKeyExpander" Header="YubiKey Information"
                          FontSize="14" Foreground="#FF69B4"
                          IsExpanded="True" Margin="0,0,0,10">
                    <Border x:Name="YubiKeyBorder" BorderBrush="#FF69B4" BorderThickness="1"
                            Padding="10" CornerRadius="5" Background="White" Margin="5">
                        <StackPanel Orientation="Vertical">
                            <TextBlock x:Name="YubiKeyStatusText" FontSize="12" Margin="5" TextWrapping="Wrap"/>
                        </StackPanel>
                    </Border>
                </Expander>

                <!-- Logs Section -->
                <Expander Header="Logs"
                          FontSize="14" Foreground="#ff8c00"
                          IsExpanded="False" Margin="0,0,0,10">
                    <Border BorderBrush="#ff8c00" BorderThickness="1"
                            Padding="10" CornerRadius="5" Background="White" Margin="5">
                        <StackPanel Orientation="Vertical">
                            <TextBox x:Name="LogTextBox" FontSize="10" Margin="5"
                                     Height="200" IsReadOnly="True" TextWrapping="Wrap"
                                     VerticalScrollBarVisibility="Auto"/>
                            <Button x:Name="ExportLogsButton" Content="Export Logs"
                                    Width="100" Margin="5" HorizontalAlignment="Right"/>
                        </StackPanel>
                    </Border>
                </Expander>

                <!-- Settings Section -->
                <Expander Header="Settings"
                          FontSize="14" Foreground="#1e90ff"
                          IsExpanded="False" Margin="0,0,0,10">
                    <Border BorderBrush="#1e90ff" BorderThickness="1"
                            Padding="10" CornerRadius="5" Background="White" Margin="5">
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <TextBlock Text="Refresh Interval (seconds): "
                                       FontSize="12" VerticalAlignment="Center" Margin="5"/>
                            <TextBox x:Name="RefreshIntervalTextBox" Width="50" Text="30"
                                     FontSize="12" Margin="5"/>
                            <Button x:Name="ApplySettingsButton" Content="Apply"
                                    Width="60" FontSize="12" Margin="5"/>
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
                   Margin="0,10,0,0"/>
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
$LoggedOnUserText       = $window.FindName("LoggedOnUserText")
$MachineTypeText        = $window.FindName("MachineTypeText")
$OSVersionText          = $window.FindName("OSVersionText")
$SystemUptimeText       = $window.FindName("SystemUptimeText")
$UsedDiskSpaceText      = $window.FindName("UsedDiskSpaceText")
$IpAddressText          = $window.FindName("IpAddressText")  # Newly added IP text block
$AntivirusStatusText    = $window.FindName("AntivirusStatusText")
$BitLockerStatusText    = $window.FindName("BitLockerStatusText")
$YubiKeyStatusText      = $window.FindName("YubiKeyStatusText")
$LogTextBox             = $window.FindName("LogTextBox")
$RefreshIntervalTextBox = $window.FindName("RefreshIntervalTextBox")
$ExportLogsButton       = $window.FindName("ExportLogsButton")
$ApplySettingsButton    = $window.FindName("ApplySettingsButton")

# Additional Named Controls for Color Updates
$BitLockerExpander = $window.FindName("BitLockerExpander")
$BitLockerBorder   = $window.FindName("BitLockerBorder")
$YubiKeyExpander   = $window.FindName("YubiKeyExpander")
$YubiKeyBorder     = $window.FindName("YubiKeyBorder")

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

        # OS version
        $os = Get-CimInstance -ClassName Win32_OperatingSystem
        $osVersion = "$($os.Caption) (Build $($os.BuildNumber))"
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

        # Retrieve IPv4 addresses (excluding loopback 127.* and APIPA 169.254.*)
        $ipv4s = Get-NetIPAddress -AddressFamily IPv4 | Where-Object {
            $_.IPAddress -notlike "127.*" -and
            $_.IPAddress -notlike "169.254.*" -and
            $_.IPAddress -notin @("0.0.0.0","255.255.255.255") -and
            $_.PrefixOrigin -ne "WellKnown"
        } | Select-Object -ExpandProperty IPAddress -ErrorAction SilentlyContinue

        if ($ipv4s) {
            # Join multiple addresses with comma + space if you have multiple NICs
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

# COM-based BitLocker check
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
        return $false, "Error retrieving BitLocker info via Shell.Application: $_"
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
    $yubicoVendorID = "1050"
    $yubikeyProductIDs = @("0407","0408","0409","040A","040B","040C","040D","040E")

    Write-Log "Starting YubiKey detection..."
    try {
        # Only check devices with Status='OK' to skip ghost devices
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
    Write-Log "Get-Icon called with Path='$Path' and DefaultIcon type='$($DefaultIcon.GetType())'"
    if (-not (Test-Path $Path)) {
        Write-Log "$Path not found. Using default icon."
        return $DefaultIcon
    }
    else {
        try {
            $icon = New-Object System.Drawing.Icon($Path)
            Write-Log "Custom icon loaded from $Path."
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
        $yubikeyStatus, $yubikeyMessage     = Get-YubiKeyStatus

        # Overall system health
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

        # Dynamically change BitLocker panel color
        if ($bitlockerStatus) {
            $BitLockerExpander.Foreground = 'Green'
            $BitLockerBorder.BorderBrush = 'Green'
        }
        else {
            $BitLockerExpander.Foreground = 'Red'
            $BitLockerBorder.BorderBrush = 'Red'
        }

        # Dynamically change YubiKey panel color
        if ($yubikeyStatus) {
            $YubiKeyExpander.Foreground = 'Green'
            $YubiKeyBorder.BorderBrush = 'Green'
        }
        else {
            $YubiKeyExpander.Foreground = 'Red'
            $YubiKeyBorder.BorderBrush = 'Red'
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
# 10) Settings Management
# ========================
function Apply-Settings {
    try {
        $newInterval = [int]$RefreshIntervalTextBox.Text
        if ($newInterval -lt 10) {
            [System.Windows.Forms.MessageBox]::Show("Refresh interval must be at least 10 seconds.",
                "Invalid Input",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
            Write-Log "Invalid refresh interval entered: $newInterval"
            return
        }
        $timer.Interval = $newInterval * 1000
        Write-Log "Refresh interval set to $newInterval seconds."
    }
    catch {
        Write-Log "Error applying settings: $_"
    }
}

# ========================
# 11) Window Visibility Management
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
# 12) Button Event Handlers
# ========================
$ExportLogsButton.Add_Click({ Export-Logs })
$ApplySettingsButton.Add_Click({ Apply-Settings })

# ========================
# 13) Create & Configure Tray Icon
# ========================
$TrayIcon = New-Object System.Windows.Forms.NotifyIcon
$TrayIcon.Icon = Get-Icon -Path $GreenIconPath -DefaultIcon ([System.Drawing.SystemIcons]::Application)
$TrayIcon.Visible = $true
$TrayIcon.Text = "System Monitor"

# ========================
# 14) Tray Icon Context Menu
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

# "Show Dashboard" from context menu
$MenuItemShow.add_Click({
    Toggle-WindowVisibility
})

# "Exit" from context menu
$MenuItemExit.add_Click({
    try {
        # Stop periodic updates so timer won't fire again
        $timer.Stop()

        # Dispose tray icon
        $TrayIcon.Dispose()

        # Gracefully shut down Dispatcher loop
        $window.Dispatcher.InvokeShutdown()
        Write-Log "Application exited via tray menu."
    }
    catch {
        Write-Log "Error during application exit: $_"
    }
})

# ========================
# 15) Timer for Periodic Updates
# ========================
$timer = New-Object System.Windows.Forms.Timer
$defaultInterval = 30
$timer.Interval = $defaultInterval * 1000
$timer.Add_Tick({
    try {
        $window.Dispatcher.Invoke([Action]{
            Update-TrayIcon
            Update-SystemInfo
            Update-YubiKeyStatus
            Update-Logs
        })
    }
    catch {
        Write-Log "Error during timer tick: $_"
    }
})
$timer.Start()

# ========================
# 16) YubiKey Update Function
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
# 17) Dispatcher Exception Handling
# ========================
function Handle-DispatcherUnhandledException {
    param(
        [object]$sender,
        [System.Windows.Threading.DispatcherUnhandledExceptionEventArgs]$args
    )
    Write-Log "Unhandled Dispatcher exception: $($args.Exception.Message)"
    # We'll just mark it as handled to avoid crashing
    $args.Handled = $true
}

Register-ObjectEvent -InputObject $window.Dispatcher -EventName UnhandledException -Action {
    param($sender, $args)
    Handle-DispatcherUnhandledException -sender $sender -args $args
}

# ========================
# 18) Initialize the First Update
# ========================
try {
    $window.Dispatcher.Invoke([Action]{
        Update-SystemInfo
        Update-TrayIcon
        Update-YubiKeyStatus
        Update-Logs
    })
}
catch {
    Write-Log "Error during initial update: $_"
}

# ========================
# 19) Handle Window Closing (Hide Instead of Close)
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
# 20) Start the Application Dispatcher
# ========================
[System.Windows.Threading.Dispatcher]::Run()
Write-Log "Dispatcher ended; script exiting."
