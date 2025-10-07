# Lincoln Laboratory Endpoint Advisor
# Version 6.1.0 (Tabbed Interface Update)

# Ensure $PSScriptRoot is defined for older versions
if ($MyInvocation.MyCommand.Path) {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    $ScriptDir = Get-Location
}

# Define version
$ScriptVersion = "6.1.0"

# --- START OF SINGLE-INSTANCE CHECK ---
# Single-Instance Check: Prevents multiple copies of the application from running.
$AppName = "Lincoln Laboratory Endpoint Advisor"
# Find any other PowerShell process with the same window title, excluding the current process.
$existingProcesses = Get-Process -Name "powershell", "pwsh" -ErrorAction SilentlyContinue | Where-Object { $_.Id -ne $pid -and $_.MainWindowTitle -like "*$AppName*" }
if ($existingProcesses) {
    Write-Host "An instance of $AppName is already running. Exiting."
    exit
}
# --- END OF SINGLE-INSTANCE CHECK ---

# CODE ADDED TO TERMINATE NOTIFICATION IF RUN ON IDENTIFIED CONFERENCE ROOM PC
try {
    $registryPath = "HKLM:\SOFTWARE\MITLL" 
    $Name = "IsConferenceRoom" 
    $exists = Get-ItemProperty -Path $registryPath -Name $Name -ErrorAction SilentlyContinue 
    if ($exists -and $exists.$Name) {
        Write-Host "[INFO] Notification aborted due to detection of Conference Room PC via registry."
        exit
    }
    If (Test-Path "C:\Windows\IsConferenceRoom.stub") {
        Write-Host "[INFO] Notification aborted due to detection of Conference Room PC via stub file."
        exit
    }
} catch {
    # Fail silently if the check itself causes an error
}
# END OF CONFERENCE ROOM CHECK

# Global flag to prevent recursive logging during rotation
$global:IsRotatingLog = $false

# Global flag to track pending update state
$global:UpdatesPending = $false
$global:CurrentUpdateState = ""
$global:LastAnnouncementState = ""
$global:DynamicTabsCreated = $false
# Global variables for certificate check caching
$global:LastCertificateCheck = $null
$global:CachedCertificateStatus = $null

# Global counter for failed fetch attempts
$global:FailedFetchAttempts = 0

# ============================================================
# A) Advanced Logging & Error Handling
# ============================================================
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

    $logPath = if ($LogFilePath) { $LogFilePath } else { Join-Path $ScriptDir "LLEndpointAdvisor.log" }
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

function Invoke-WithRetry {
    param(
        [ScriptBlock]$Action,
        [int]$MaxRetries = 3,
        [int]$RetryDelayMs = 500
    )
    $attempt = 0
    while ($attempt -lt $MaxRetries) {
        try {
            $attempt++
            $result = $Action.Invoke()
            return $result
        }
        catch {
            if ($attempt -ge $MaxRetries) {
                throw "Action failed after $MaxRetries attempts: $($_.Exception.Message)"
            }
            Write-Log "Retry $attempt of $MaxRetries failed: $($_.Exception.Message)" -Level "WARNING"
            Start-Sleep -Milliseconds ($RetryDelayMs * (2 * $attempt)) # Exponential backoff
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

                    $archiveFiles = Get-ChildItem -Path $LogDirectory -Filter "LLEndpointAdvisor.log.*.archive" | Sort-Object CreationTime
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
                                Write-Log "Failed to delete old archive $($file.FullName) - $($_.Exception.Message)" -Level "ERROR"
                            }
                        }
                    }
                }
                catch {
                    Write-Log "Failed to rotate log file - $($_.Exception.Message)" -Level "ERROR"
                }
            }
        }
    }
    catch {
        Write-Log "Error checking log file size for rotation - $($_.Exception.Message)" -Level "ERROR"
    }
    finally {
        $global:IsRotatingLog = $false
    }
}

function Handle-Error {
    param(
        [string]$Message,
        [string]$Source = ""
    )
    if ($Source) { $Message = "[$Source] $Message" }
    Write-Log $Message -Level "ERROR"
}

Write-Log "--- Lincoln Laboratory Endpoint Advisor Script Started (Version $ScriptVersion) ---"

# ============================================================
# MODULE: Configuration Management
# ============================================================
function Get-DefaultConfig {
    return @{
        RefreshInterval       = 900
        LogRotationSizeMB     = 2
        DefaultLogLevel       = "INFO"
        ContentDataUrl        = "https://raw.githubusercontent.com/burnoil/EndpointAdvisor/refs/heads/main/ContentData.json"
        CertificateCheckInterval = 86400
        YubiKeyAlertDays      = 14
        IconPaths             = @{
            Main    = Join-Path $ScriptDir "LL_LOGO.ico"
            Warning = Join-Path $ScriptDir "LL_LOGO_MSG.ico"
        }
        AnnouncementsLastState = "{}"
        SupportLastState       = "{}"
        LastSeenUpdateState   = ""
		LastPatchingState     = ""
        Version               = $ScriptVersion
        BigFixSSA_Path        = "C:\Program Files (x86)\BigFix Enterprise\BigFix Self Service Application\BigFixSSA.exe"
        YubiKeyManager_Path   = "C:\Program Files\Yubico\Yubikey Manager\ykman.exe"
        BlinkingEnabled       = $false
        CachePath             = Join-Path $ScriptDir "ContentData.cache.json"
        HasRunBefore          = $false
    }
}

function Load-Configuration {
    param([string]$Path = (Join-Path $ScriptDir "LLEndpointAdvisor.config.json"))
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
            Write-Log "Failed to load or merge existing config file. Reverting to full defaults. - $($_.Exception.Message)" -Level "WARNING"
        }
    }
    try {
        $finalConfig | ConvertTo-Json -Depth 100 | Out-File $Path -Force
        Write-Log "Configuration file validated and saved." -Level "INFO"
    }
    catch {
        Write-Log "Could not save the updated configuration to $Path - $($_.Exception.Message)" -Level "ERROR"
    }
    return $finalConfig
}

function Save-Configuration {
    param(
        [psobject]$Config,
        [string]$Path = (Join-Path $ScriptDir "LLEndpointAdvisor.config.json")
    )
    try {
        $Config | ConvertTo-Json -Depth 100 | Out-File $Path -Force
        Write-Log "Configuration file saved to $Path" -Level "INFO"
    } catch {
        Write-Log "Could not save state to configuration file $Path - $($_.Exception.Message)" -Level "ERROR"
    }
}

# ============================================================
# B) External Configuration Setup
# ============================================================
$LogFilePath = Join-Path $ScriptDir "LLEndpointAdvisor.log"
$config = Load-Configuration

$mainIconPath = $config.IconPaths.Main
$warningIconPath = $config.IconPaths.Warning

Write-Log "Main icon path: $mainIconPath" -Level "INFO"
Write-Log "Warning icon path: $warningIconPath" -Level "INFO"

$defaultContentData = @{
    Announcements = @{ 
        Default = @{ Text = "No announcements at this time."; Details = ""; Links = @() }
        Targeted = @()
    }
    Support = @{ Text = "Contact IT Support."; Links = @() }
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
        Write-Log "Failed to load required GUI assemblies - $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

$global:FormsAvailable = Import-RequiredAssemblies

# ============================================================
# E) XAML Layout Definition
# ============================================================
$xamlString = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Lincoln Laboratory Endpoint Advisor"
    WindowStartupLocation="Manual" 
    SizeToContent="Manual"
    MinWidth="380" MinHeight="500"
	MaxWidth="450" MaxHeight="550"
    ResizeMode="CanResizeWithGrip" ShowInTaskbar="False" Visibility="Hidden" Topmost="True"
    Background="#f0f0f0">
  <Window.Resources>
    <!-- Custom Expander Style (unchanged) -->
    <Style TargetType="Expander">
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Expander">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition x:Name="ContentRow" Height="0"/>
              </Grid.RowDefinitions>
              <Border x:Name="HeaderBorder" Grid.Row="0" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
                <ToggleButton x:Name="ToggleButton" IsChecked="{Binding IsExpanded, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}" Margin="5,0,0,0">
                  <ToggleButton.Template>
                    <ControlTemplate TargetType="ToggleButton">
                      <Border Background="Transparent">
                        <ContentPresenter Content="{TemplateBinding Content}"/>
                      </Border>
                    </ControlTemplate>
                  </ToggleButton.Template>
                  <ToggleButton.Content>
                    <Grid>
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                      </Grid.ColumnDefinitions>
                      <Path x:Name="Arrow" Grid.Column="0" Data="M 0 0 L 8 8 L 0 16 Z" Fill="#0055A4" Stroke="#D3D3D3" StrokeThickness="1" HorizontalAlignment="Center" VerticalAlignment="Center" Margin="0,0,10,0">
                        <Path.RenderTransform>
                          <ScaleTransform ScaleX="1.2" ScaleY="1.2"/>
                        </Path.RenderTransform>
                      </Path>
                      <ContentPresenter Grid.Column="1" Content="{TemplateBinding Header}"/>
                    </Grid>
                  </ToggleButton.Content>
                </ToggleButton>
              </Border>
              <Border x:Name="ContentBorder" Grid.Row="1" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
                <ContentPresenter x:Name="ExpandSite" Visibility="Collapsed"/>
              </Border>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsExpanded" Value="True">
                <Setter TargetName="ContentRow" Property="Height" Value="*"/>
                <Setter TargetName="ExpandSite" Property="Visibility" Value="Visible"/>
                <Setter TargetName="Arrow" Property="Data" Value="M 0 0 L 8 8 L 16 0 Z"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>
  <Grid Margin="5">
    <Grid.RowDefinitions>
  <RowDefinition Height="Auto"/>
  <RowDefinition Height="Auto"/>
  <RowDefinition Height="*"/>
  <RowDefinition Height="Auto"/>
</Grid.RowDefinitions>
    
    <!-- Header (unchanged) -->
    <Border Grid.Row="0" Background="#0078D7" Padding="5" CornerRadius="3" Margin="0,0,0,5">
      <StackPanel Orientation="Horizontal" VerticalAlignment="Center" HorizontalAlignment="Center">
        <Image x:Name="HeaderIcon" Width="20" Height="20" Margin="0,0,5,0"/>
        <TextBlock Text="Lincoln Laboratory Endpoint Advisor" FontSize="14" FontWeight="Bold" Foreground="White" VerticalAlignment="Center"/>
      </StackPanel>
    </Border>
<!-- Patching and Updates Section -->
<StackPanel Grid.Row="2" Margin="0,5,0,5">
  <StackPanel Orientation="Horizontal" Margin="0,0,0,5">
    <TextBlock Text="Patching and Updates" FontSize="12" FontWeight="Bold" VerticalAlignment="Center"/>
    <Ellipse x:Name="PatchingAlertDot" Width="10" Height="10" Margin="5,0,0,0" Fill="Red" Visibility="Hidden"/>
  </StackPanel>
  <Border BorderBrush="#00008B" BorderThickness="2" Padding="8" CornerRadius="3" Background="White">
    <StackPanel>
      <Grid Margin="0,2,0,2">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock x:Name="BigFixStatusText" Grid.Column="0" VerticalAlignment="Center" FontSize="11" TextWrapping="Wrap"/>
        <Button x:Name="BigFixLaunchButton" Grid.Column="1" Content="App Updates" Margin="10,0,0,0" Padding="5,1" VerticalAlignment="Center" Visibility="Collapsed" ToolTip="Install available application updates"/>
      </Grid>
      <Separator Margin="0,5,0,5"/>
      <Grid Margin="0,2,0,2">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock x:Name="ECMStatusText" Grid.Column="0" VerticalAlignment="Center" FontSize="11" TextWrapping="Wrap"/>
        <Button x:Name="ECMLaunchButton" Grid.Column="1" Content="Install Patches" Margin="10,0,0,0" Padding="5,1" VerticalAlignment="Center" Visibility="Collapsed" ToolTip="Install pending Windows OS patches"/>
      </Grid>
      <Separator Margin="0,5,0,5"/>
      <Grid Margin="0,2,0,2">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Grid.Column="0" VerticalAlignment="Center">
          <TextBlock x:Name="DriverUpdateStatusText" FontSize="11" FontWeight="Bold" TextWrapping="Wrap" Text="Windows Driver Updates (Required every month. Your computer will automatically restart when this is complete.)"/>
          <TextBlock x:Name="DriverUpdateLastRunText" FontSize="9" Foreground="Gray" TextWrapping="Wrap" Text="Checking status..."/>
        </StackPanel>
        <Button x:Name="DriverUpdateButton" Grid.Column="1" Content="Install Drivers" Margin="10,0,0,0" Padding="5,1" VerticalAlignment="Center" Visibility="Collapsed" ToolTip="Install driver updates via Windows Update"/>
      </Grid>
	  <!-- Driver Update Progress Panel (hidden by default) -->
<Border x:Name="DriverProgressPanel" BorderBrush="#0078D7" BorderThickness="2" Background="#F0F8FF" Padding="10" CornerRadius="3" Margin="0,10,0,0" Visibility="Collapsed">
  <StackPanel>
    <TextBlock Text="Driver Update Progress" FontSize="11" FontWeight="Bold" Margin="0,0,0,5"/>
    <TextBlock x:Name="DriverProgressStatus" FontSize="11" TextWrapping="Wrap" Text="Initializing..."/>
    <ProgressBar x:Name="DriverProgressBar" Height="20" Margin="0,10,0,0" IsIndeterminate="True"/>
    <TextBlock FontSize="9" Foreground="Gray" Margin="0,5,0,0" Text="This process may take several minutes. You can continue working."/>
  </StackPanel>
</Border>
    </StackPanel>
  </Border>
</StackPanel>
    <!-- NEW: TabControl to replace the ScrollViewer -->
    <TabControl x:Name="MainTabControl" Grid.Row="1" TabStripPlacement="Top">
        <TabItem>
  <TabItem.Header>
    <StackPanel Orientation="Horizontal">
      <TextBlock Text="ISD Dashboard" VerticalAlignment="Center"/>
      <Ellipse x:Name="DashboardTabAlert" Width="8" Height="8" Margin="5,0,0,0" Fill="Red" Visibility="Hidden"/>
    </StackPanel>
  </TabItem.Header>
            <ScrollViewer VerticalScrollBarVisibility="Auto">
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
                      <StackPanel x:Name="AppendedAnnouncementsPanel" Orientation="Vertical" Margin="0,5,0,0" Visibility="Collapsed"/>
                      <StackPanel x:Name="AnnouncementsLinksPanel" Orientation="Vertical" Margin="0,5,0,0"/>
                      <TextBlock x:Name="AnnouncementsSourceText" FontSize="9" Foreground="Gray" Margin="0,5,0,0"/>
                    </StackPanel>
                  </Border>
                </Expander>
                
              </StackPanel>
            </ScrollViewer>
        </TabItem>
		<TabItem x:Name="SupportTab">
  <TabItem.Header>
    <StackPanel Orientation="Horizontal">
      <TextBlock Text="Support" VerticalAlignment="Center"/>
      <Ellipse x:Name="SupportTabAlert" Width="8" Height="8" Margin="5,0,0,0" Fill="Red" Visibility="Hidden"/>
    </StackPanel>
  </TabItem.Header>
  <ScrollViewer VerticalScrollBarVisibility="Auto">
    <StackPanel Margin="10">
      <Border BorderBrush="#00008B" BorderThickness="1" Padding="10" CornerRadius="3" Background="White">
        <StackPanel>
          <TextBlock x:Name="SupportText" FontSize="11" TextWrapping="Wrap"/>
          <StackPanel x:Name="SupportLinksPanel" Orientation="Vertical" Margin="0,10,0,0"/>
          <TextBlock x:Name="SupportSourceText" FontSize="9" Foreground="Gray" Margin="0,10,0,0"/>
        </StackPanel>
      </Border>
    </StackPanel>
  </ScrollViewer>
</TabItem>
       <TabItem Header="About">
            <StackPanel Margin="10">
                <TextBlock Text="Lincoln Laboratory Endpoint Advisor" FontWeight="Bold" FontSize="14"/>
                <TextBlock x:Name="AboutVersionText" Text="Version: (loading...)" Margin="0,5,0,10"/>
                <TextBlock Text="This application provides timely announcements, system status, and support information for your endpoint." TextWrapping="Wrap"/>
            </StackPanel>
        </TabItem>
		
    </TabControl>
    
    <!-- Footer (unchanged) -->
    <Grid Grid.Row="3" Margin="0,5,0,0">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*" />
            <ColumnDefinition Width="Auto" />
        </Grid.ColumnDefinitions>
        <TextBlock x:Name="FooterText" Grid.Column="0" Text="(C) 2025 Lincoln Laboratory" FontSize="10" Foreground="Gray" HorizontalAlignment="Center" VerticalAlignment="Center"/>
        <StackPanel x:Name="ClearAlertsPanel" Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
            <Ellipse x:Name="ClearAlertsDot" Width="10" Height="10" Fill="Red" Margin="0,0,5,0" Visibility="Collapsed"/>
            <Button x:Name="ClearAlertsButton" Content="Clear Alerts" FontSize="10" Padding="5,1" Background="#B0C4DE" ToolTip="Click to clear all new announcement and support alerts (red dots) from the UI."/>
        </StackPanel>
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

    $window.Width = 380
    $window.Height = 600

    # Modified UI Elements list to include new elements
    $uiElements = @(
    "HeaderIcon", "MainTabControl", "AnnouncementsExpander", "AnnouncementsAlertIcon", "AnnouncementsText", "AnnouncementsDetailsText",
    "AnnouncementsLinksPanel", "AnnouncementsSourceText", "SupportText", "SupportLinksPanel",
    "SupportSourceText", "ClearAlertsButton",
    "FooterText", "ClearAlertsPanel", "ClearAlertsDot", "BigFixStatusText", "BigFixLaunchButton", "ECMStatusText", "ECMLaunchButton",
    "AppendedAnnouncementsPanel", "AboutVersionText", "DashboardTabAlert", "SupportTabAlert", "SupportTab",
    "DriverUpdateStatusText", "DriverUpdateButton", "DriverUpdateLastRunText", "PatchingAlertDot",
    "DriverProgressPanel", "DriverProgressStatus", "DriverProgressBar"
)
    foreach ($elementName in $uiElements) {
        $value = $window.FindName($elementName)
        Set-Variable -Name "global:$elementName" -Value $value
        if (-not $value) {
            Write-Log "UI element $elementName is null." -Level "WARNING"
        } else {
            Write-Log "UI element $elementName initialized." -Level "INFO"
        }
    }
    Write-Log "UI elements mapped to variables." -Level "INFO"
    
    # Set dynamic text in the 'About' tab
    $global:AboutVersionText.Text = "Version: $ScriptVersion"
    $global:FooterText.Text = "(C) 2025 Lincoln Laboratory v$ScriptVersion"

    # Set window icon in code
    if (Test-Path $mainIconPath) {
        $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
        $bitmap.BeginInit()
        $bitmap.UriSource = New-Object System.Uri $mainIconPath
        $bitmap.EndInit()
        $global:window.Icon = $bitmap
        Write-Log "Window icon set successfully." -Level "INFO"
    } else {
        Write-Log "Main icon not found at $mainIconPath." -Level "WARNING"
    }

    # Set header icon in code
    if ($global:HeaderIcon -and (Test-Path $mainIconPath)) {
        $global:HeaderIcon.Source = $bitmap
        Write-Log "Header icon set successfully." -Level "INFO"
    } else {
        Write-Log "Header icon element null or file not found." -Level "WARNING"
    }

    $window.Add_Closing({
        $_.Cancel = $true
        $window.Hide()
    })

function Test-LaptopOnBattery {
    try {
        # Check if system has a battery (indicates laptop/portable device)
        $battery = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue
        
        if (-not $battery) {
            # No battery found, likely a desktop
            Write-Log "No battery detected - appears to be a desktop system." -Level "INFO"
            return $false
        }
        
        # Has battery - check power status
        Add-Type -AssemblyName System.Windows.Forms
        $powerStatus = [System.Windows.Forms.SystemInformation]::PowerStatus
        
        if ($powerStatus.PowerLineStatus -eq 'Online') {
            Write-Log "Laptop detected on AC power." -Level "INFO"
            return $false
        } else {
            Write-Log "Laptop detected on BATTERY power - warning needed!" -Level "WARNING"
            return $true
        }
    }
    catch {
        Write-Log "Error detecting power status: $($_.Exception.Message)" -Level "ERROR"
        # If we can't determine, be safe and show warning
        return $true
    }
}

function Start-DriverUpdateMonitoring {
    Write-Log "Starting driver update monitoring..." -Level "INFO"
    
    # Show the progress panel
    $window.Dispatcher.Invoke({
        $global:DriverProgressPanel.Visibility = "Visible"
        $global:DriverProgressStatus.Text = "Preparing system for driver updates..."
    })
    
    # Create and start monitoring timer
    $global:DriverMonitorTimer = New-Object System.Windows.Threading.DispatcherTimer
    $global:DriverMonitorTimer.Interval = [TimeSpan]::FromSeconds(5)
    
    $global:DriverMonitorTimer.Add_Tick({
        try {
            $logPath = "C:\Windows\mitll\Logs\MS_Update.txt"
            
            if (-not (Test-Path $logPath)) {
                $global:DriverProgressStatus.Text = "Waiting for update process to start..."
                return
            }
            
            # Read last few lines of log
            $logContent = Get-Content $logPath -Tail 20 -ErrorAction SilentlyContinue
            $lastLine = $logContent | Select-Object -Last 1
            
            # Translate technical log entries to user-friendly messages
            if ($logContent -match "Install-Module") {
                $global:DriverProgressStatus.Text = "[OK] Preparing update tools...`n[...] Checking for available driver updates..."
            }
            elseif ($logContent -match "Downloading" -or $logContent -match "Download") {
                $global:DriverProgressStatus.Text = "[OK] Found driver updates`n[...] Downloading updates..."
            }
            elseif ($logContent -match "Installing" -or $logContent -match "Install") {
                $global:DriverProgressStatus.Text = "[OK] Download complete`n[...] Installing drivers...`n`nPlease wait, this may take several minutes."
            }
            elseif ($logContent -match "Success" -or $logContent -match "Installed") {
                $global:DriverProgressStatus.Text = "[OK] Driver installation complete!`n`nYour computer may restart shortly."
                $global:DriverProgressBar.IsIndeterminate = $false
                $global:DriverProgressBar.Value = 100
                $global:DriverMonitorTimer.Stop()
                Write-Log "Driver update monitoring completed." -Level "INFO"
            }
            elseif ($logContent -match "Failed" -or $logContent -match "Error") {
                $global:DriverProgressStatus.Text = "[!] An issue occurred during installation.`n`nCheck C:\Windows\mitll\Logs\MS_Update.txt for details."
                $global:DriverProgressBar.IsIndeterminate = $false
                $global:DriverMonitorTimer.Stop()
                Write-Log "Driver update encountered errors." -Level "WARNING"
            }
            
        } catch {
            Write-Log "Error monitoring driver update progress: $($_.Exception.Message)" -Level "ERROR"
        }
    })
    
    $global:DriverMonitorTimer.Start()
}

    # Initialize events after a delay
    function InitializeUI {
        if ($global:AnnouncementsExpander) {
            $global:AnnouncementsExpander.IsExpanded = $true
            $global:AnnouncementsExpander.Add_Expanded({ 
                if ($global:AnnouncementsAlertIcon) { $global:AnnouncementsAlertIcon.Visibility = "Hidden" }
                Update-TrayIcon
            })
        }
        
        if ($global:BigFixLaunchButton) {
            $global:BigFixLaunchButton.Add_Click({
                try {
                    $ssaPath = $config.BigFixSSA_Path
                    if ([string]::IsNullOrWhiteSpace($ssaPath) -or -not (Test-Path $ssaPath)) {
                        throw "BigFix Self-Service Application path is invalid or not found: `"$ssaPath`""
                    }
                    Write-Log "Launching BigFix SSA: $ssaPath" -Level "INFO"
                    Start-Process -FilePath $ssaPath
                }
                catch {
                    Handle-Error $_.Exception.Message -Source "BigFixLaunchButton"
                }
            })
        }
        if ($global:ECMLaunchButton) {
            $global:ECMLaunchButton.Add_Click({
                try {
                    $softwareCenterPath = "$($Env:WinDir)\CCM\SCClient.exe"
                    if (-not (Test-Path $softwareCenterPath)) {
                        throw "Microsoft Software Center not found at: `"$softwareCenterPath`""
                    }
                    Write-Log "Launching Microsoft Software Center: $softwareCenterPath" -Level "INFO"
                    Start-Process -FilePath $softwareCenterPath
                }
                catch {
                    Handle-Error $_.Exception.Message -Source "ECMLaunchButton"
                }
            })
        }
if ($global:DriverUpdateButton) {
    $global:DriverUpdateButton.Add_Click({
        try {
            Write-Log "Driver Update button clicked by user." -Level "INFO"
            
            # Check when drivers were last updated
            $lastRunStatus = Get-DaysSinceLastDriverUpdate
            
            # Parse the days from the status string
            if ($lastRunStatus -match "Last run (\d+) day") {
                $daysSinceUpdate = [int]$matches[1]
                
                if ($daysSinceUpdate -lt 30) {
                    Write-Log "Driver updates were run $daysSinceUpdate days ago (within 30-day window)." -Level "INFO"
                    
                    $recentUpdateWarning = [System.Windows.MessageBox]::Show(
                        "Driver updates were last run $daysSinceUpdate days ago.`n`nDriver updates are typically only required once per month. Running them more frequently is usually unnecessary unless requested by ISD.`n`nDo you still want to proceed?",
                        "Recent Update Detected",
                        [System.Windows.MessageBoxButton]::YesNo,
                        [System.Windows.MessageBoxImage]::Information
                    )
                    
                    if ($recentUpdateWarning -ne [System.Windows.MessageBoxResult]::Yes) {
                        Write-Log "User cancelled driver update - recent update detected." -Level "INFO"
                        return
                    }
                }
            }
            
            # Check if laptop on battery power
            $needsPowerWarning = Test-LaptopOnBattery
            
            if ($needsPowerWarning) {
                # Show AC power warning for laptops on battery
                $powerWarning = [System.Windows.MessageBox]::Show(
                    "WARNING: Your laptop is currently running on battery power!`n`nPlease plug into AC power before continuing. Driver updates require a stable power source and will require a system restart.`n`nAre you plugged in and ready to proceed?",
                    "AC Power Required",
                    [System.Windows.MessageBoxButton]::YesNo,
                    [System.Windows.MessageBoxImage]::Warning
                )
                
                if ($powerWarning -ne [System.Windows.MessageBoxResult]::Yes) {
                    Write-Log "User cancelled driver update - needs AC power." -Level "INFO"
                    return
                }
            }
            
            # Final confirmation for the actual update
            $result = [System.Windows.MessageBox]::Show(
                "This will install driver updates via Windows Update and will require a system restart. Continue?",
                "Install Driver Updates",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Question
            )
            
            if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                Write-Log "User confirmed driver update installation." -Level "INFO"
                
                # Run the driver update script
                $scriptBlock = @"
Set-itemproperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' -Name SetPolicyDrivenUpdateSourceForDriverUpdates -Value 0 -ErrorAction SilentlyContinue
try { Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\TrustedInstaller" -Name "BlockTimeIncrement" -Value 3600 -type dword -ErrorAction Stop } catch { }
Install-Module -Name PSWindowsUpdate -Force
`$dateTime = Get-Date -Format "MM/dd/yyyy"
`$dateTime | Out-File -Append -FilePath "C:\Windows\mitll\Logs\MS_Update.txt"
Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot | Out-File -Append -FilePath "C:\Windows\mitll\Logs\MS_Update.txt"
Set-itemproperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' -Name SetPolicyDrivenUpdateSourceForDriverUpdates -Value 1 -ErrorAction SilentlyContinue
try { Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\TrustedInstaller" -Name "BlockTimeIncrement" -Value 900 -type dword -ErrorAction Stop } catch { }
`$bitlockerstatus = get-bitlockervolume -mountpoint "c:"
if (`$bitlockerstatus.ProtectionStatus -eq 'On') {
    Suspend-BitLocker -MountPoint "C:" -RebootCount 1
}
"@
                
Start-Process powershell.exe -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden", "-Command", $scriptBlock -Verb RunAs

# Start monitoring the update progress
Start-DriverUpdateMonitoring

[System.Windows.MessageBox]::Show(
    "Driver update installation has been started. Progress will be shown below.",
    "Driver Updates Started",
    [System.Windows.MessageBoxButton]::OK,
    [System.Windows.MessageBoxImage]::Information
)
            }
        }
        catch {
            Handle-Error $_.Exception.Message -Source "DriverUpdateButton"
            [System.Windows.MessageBox]::Show(
                "Failed to start driver update installation: $($_.Exception.Message)",
                "Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        }
    })
}

        if ($global:ClearAlertsButton) {
    $global:ClearAlertsButton.Add_Click({
        Write-Log "Clear Alerts button clicked by user to clear new alerts (red dots)." -Level "INFO"
        
        $config.AnnouncementsLastState = $global:LastAnnouncementState
        
        # Use the same logic as Update-Support to get the correct object
        $supportObject = if ($global:contentData.Data.Dashboard) {
            $global:contentData.Data.Dashboard.Support
        } else {
            $global:contentData.Data.Support
        }
        $config.SupportLastState = ($supportObject | ConvertTo-Json -Compress -Depth 10)
        
        # Calculate and save current patching state
        $showBigFix = $global:BigFixLaunchButton.Visibility -eq "Visible"
        $showECM = $global:ECMLaunchButton.Visibility -eq "Visible"
        $showDriver = $global:DriverUpdateButton.Visibility -eq "Visible"
        $config.LastPatchingState = "BigFix:$showBigFix|ECM:$showECM|Driver:$showDriver"
        
        $window.Dispatcher.Invoke({
            if ($global:AnnouncementsAlertIcon) { $global:AnnouncementsAlertIcon.Visibility = "Hidden" }
            if ($global:DashboardTabAlert) { $global:DashboardTabAlert.Visibility = "Hidden" }
            if ($global:SupportTabAlert) { $global:SupportTabAlert.Visibility = "Hidden" }
            if ($global:PatchingAlertDot) { $global:PatchingAlertDot.Visibility = "Hidden" }
            if ($global:ClearAlertsDot) { $global:ClearAlertsDot.Visibility = "Hidden" }
        })
        $global:BlinkingTimer.Stop()
        Update-TrayIcon
        Save-Configuration -Config $config
    })
}
    }
if ($global:MainTabControl) {
    $global:MainTabControl.Add_SelectionChanged({
        $selectedTab = $global:MainTabControl.SelectedItem
        if ($selectedTab) {
            # Clear alert for Dashboard tab (index 0)
            if ($global:MainTabControl.SelectedIndex -eq 0 -and $global:DashboardTabAlert) {
                $global:DashboardTabAlert.Visibility = "Hidden"
            }
            # Clear alert for Support tab (index 1)
            if ($global:MainTabControl.SelectedIndex -eq 1 -and $global:SupportTabAlert) {
                $global:SupportTabAlert.Visibility = "Hidden"
            }
            Update-TrayIcon
        }
    })
}
    $window.Dispatcher.InvokeAsync({
        Start-Sleep -Milliseconds 200
        InitializeUI
    }).Wait()
}
catch {
    Handle-Error "Failed to load the XAML layout: $($_.Exception.Message)" -Source "XAML"
    exit
}

# ============================================================
# H) Modularized System Information Functions
# ============================================================
function Get-ActiveAnnouncement {
    param($AnnouncementsObject)
    $base = $AnnouncementsObject.Default
    $appended = @()
    foreach ($targeted in $AnnouncementsObject.Targeted) {
        if ($targeted.Enabled -ne $true) { continue } # Skip if not enabled
        if ($targeted.Condition.Type -eq "Registry") {
            $path = $targeted.Condition.Path
            $name = $targeted.Condition.Name
            $value = $targeted.Condition.Value
            try {
                $reg = Get-ItemProperty -Path $path -Name $name -ErrorAction Stop
                if ($reg.$name -eq $value) {
                    if ($targeted.AppendToDefault) {
                        $appended += $targeted
                    } else {
                        return @{
                            Base = $targeted
                            Appended = @()
                        }
                    }
                }
            } catch {
                # Fail silently if registry check fails
            }
        }
    }
    return @{
        Base = $base
        Appended = $appended
    }
}

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
    
    # Check for new structure first
    if ($Data.PSObject.Properties.Match('Dashboard')) {
        if (-not $Data.Dashboard.PSObject.Properties.Match('Announcements') -or 
            -not $Data.Dashboard.PSObject.Properties.Match('Support')) {
            throw "JSON Dashboard is missing 'Announcements' or 'Support' property."
        }
    }
    # Fallback to old structure
    elseif (-not $Data.PSObject.Properties.Match('Announcements') -or 
            -not $Data.PSObject.Properties.Match('Support')) {
        throw "JSON data is missing 'Announcements' or 'Support' top-level property."
    }
    
    return $true
}

function Save-CachedContentData {
    param(
        [psobject]$ContentData,
        [string]$Path = $config.CachePath
    )
    try {
        $ContentData.Data | ConvertTo-Json -Depth 100 | Out-File $Path -Force
        Write-Log "Saved cached content data to $Path" -Level "INFO"
    } catch {
        Write-Log "Failed to save cached content data to $Path - $($_.Exception.Message)" -Level "ERROR"
    }
}

function Load-CachedContentData {
    param([string]$Path = $config.CachePath)
    try {
        if (Test-Path $Path) {
            $contentData = Get-Content $Path -Raw | ConvertFrom-Json
            Validate-ContentData -Data $contentData
            $lastWriteTime = (Get-Item $Path).LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
            Write-Log "Loaded cached content data from $Path (Last updated: $lastWriteTime)" -Level "INFO"
            return [PSCustomObject]@{
                Data = $contentData
                Source = "Cached ($lastWriteTime)"
            }
        } else {
            Write-Log "No cached content data found at $Path" -Level "WARNING"
            return $null
        }
    } catch {
        Write-Log "Failed to load cached content data from $Path - $($_.Exception.Message)" -Level "ERROR"
        return $null
    }
}

function Fetch-ContentData {
    if (-not $config -or [string]::IsNullOrWhiteSpace($config.ContentDataUrl)) {
        Write-Log "ContentDataUrl is not set! Check your Get-DefaultConfig return value." -Level "ERROR"
        $global:FailedFetchAttempts++
        $cachedData = Load-CachedContentData
        if ($cachedData) {
            return $cachedData
        }
        return [PSCustomObject]@{ Data = $defaultContentData; Source = "Default" }
    }
    $url = $config.ContentDataUrl

    try {
        Write-Log "Attempting to fetch content from: $url" -Level "INFO"
        
        $response = Invoke-WithRetry -Action {
            $job = Start-Job -ScriptBlock {
                param($url)
                Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop
            } -ArgumentList $url
            $jobResult = Wait-Job $job -Timeout 30
            if (-not $jobResult) {
                throw "Background job timed out after 30 seconds."
            }
            $result = Receive-Job $job
            Remove-Job $job
            if (-not $result) {
                throw "No response received from Invoke-WebRequest."
            }
            return $result
        } -MaxRetries 3 -RetryDelayMs 500

        Write-Log "Successfully fetched content from Git repository (Status: $($response.StatusCode))." -Level "INFO"
        
        $contentData = $response.Content | ConvertFrom-Json
        Validate-ContentData -Data $contentData
        Write-Log "Content data validated successfully." -Level "INFO"
		Write-Log "JSON Structure Check - Has Dashboard: $($contentData.PSObject.Properties.Match('Dashboard').Count -gt 0)" -Level "INFO"
    Write-Log "JSON Structure Check - Has Announcements: $($contentData.PSObject.Properties.Match('Announcements').Count -gt 0)" -Level "INFO"
    if ($contentData.Dashboard) {
        Write-Log "Dashboard.Announcements exists: $($contentData.Dashboard.Announcements -ne $null)" -Level "INFO"
        Write-Log "Dashboard.Support exists: $($contentData.Dashboard.Support -ne $null)" -Level "INFO"
    }
        Save-CachedContentData -ContentData ([PSCustomObject]@{ Data = $contentData })
        $global:FailedFetchAttempts = 0

        return [PSCustomObject]@{ Data = $contentData; Source = "Remote" }
    }
    catch {
        $global:FailedFetchAttempts++
        Write-Log "Failed to fetch or validate content from $url (Attempt $global:FailedFetchAttempts) - $($_.Exception.Message)" -Level "ERROR"
        if ($global:FailedFetchAttempts -ge 3) {
            Write-Log "Multiple consecutive fetch failures ($global:FailedFetchAttempts). Check network or URL configuration." -Level "WARNING"
        }
        $cachedData = Load-CachedContentData
        if ($cachedData) {
            return $cachedData
        }
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
            return "No YubiKey certificate found."
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
        return "No YubiKey certificate found."
    }
    catch {
        Write-Log "YubiKey check error - $($_.Exception.Message)" -Level "ERROR"
        return "Unable to determine YubiKey certificate status."
    }
}

function Get-VirtualSmartCardCertExpiry {
    try {
        $cert = Get-ChildItem "Cert:\CurrentUser\My" | Where-Object { $_.Subject -match "Virtual" } | Sort-Object NotAfter -Descending | Select-Object -First 1
        if (-not $cert) { return "No Windows Virtual Smart Card certificate found." }
        return "Windows Virtual Smart Card: Expires: $($cert.NotAfter.ToString("yyyy-MM-dd"))"
    } catch { return "Unable to check Windows Virtual Smart Card status." }
}

function Update-CertificateInfo {
    try {
        Write-Log "Checking certificate expiration..." -Level "INFO"
        $ykStatus = Get-YubiKeyCertExpiryDays
        $vscStatus = Get-VirtualSmartCardCertExpiry
        
        # Check for expiring certificates (within alert days threshold)
        $alertDays = $config.YubiKeyAlertDays
        $expiringCerts = @()
        
        # Parse YubiKey status
        if ($ykStatus -match "Expires: (\d{4}-\d{2}-\d{2})") {
            $expiryDate = [DateTime]::Parse($matches[1])
            $daysUntilExpiry = ($expiryDate - (Get-Date)).Days
            if ($daysUntilExpiry -le $alertDays -and $daysUntilExpiry -gt 0) {
                $expiringCerts += "YubiKey certificate expires in $daysUntilExpiry days"
                Write-Log "YubiKey certificate expiring in $daysUntilExpiry days" -Level "WARNING"
            } elseif ($daysUntilExpiry -le 0) {
                $expiringCerts += "YubiKey certificate has EXPIRED"
                Write-Log "YubiKey certificate has EXPIRED" -Level "ERROR"
            }
        }
        
        # Parse Virtual Smart Card status
        if ($vscStatus -match "Expires: (\d{4}-\d{2}-\d{2})") {
            $expiryDate = [DateTime]::Parse($matches[1])
            $daysUntilExpiry = ($expiryDate - (Get-Date)).Days
            if ($daysUntilExpiry -le $alertDays -and $daysUntilExpiry -gt 0) {
                $expiringCerts += "Virtual Smart Card certificate expires in $daysUntilExpiry days"
                Write-Log "Virtual Smart Card certificate expiring in $daysUntilExpiry days" -Level "WARNING"
            } elseif ($daysUntilExpiry -le 0) {
                $expiringCerts += "Virtual Smart Card certificate has EXPIRED"
                Write-Log "Virtual Smart Card certificate has EXPIRED" -Level "ERROR"
            }
        }
        
        # Show balloon notification if certificates are expiring
        if ($expiringCerts.Count -gt 0 -and $global:TrayIcon) {
            $message = $expiringCerts -join "`n"
            $global:TrayIcon.ShowBalloonTip(10000, "Certificate Expiration Alert", $message, [System.Windows.Forms.ToolTipIcon]::Warning)
        }
        
    } catch { 
        Handle-Error $_.Exception.Message -Source "Update-CertificateInfo" 
    }
}

function Get-WindowsBuildNumber {
    try {
        $buildInfo = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
        $productName = if ($buildInfo.CurrentBuildNumber -ge 22000) { "Windows 11" } else { "Windows 10" }
        return "$productName Build: $($buildInfo.DisplayVersion)"
    } catch { return "Windows Build: Unknown" }
}

function Get-ECMUpdateStatus {
    try {
        $CMTable = @{
            'class'     = 'CCM_SoftwareUpdate'
            'namespace' = 'ROOT\ccm\ClientSDK'
            'ErrorAction' = 'Stop'
        }
        $pendingUpdates = Get-WmiObject @CMTable | Where-Object { $_.ComplianceState -eq 0 }

        if ($null -eq $pendingUpdates) {
            return [PSCustomObject]@{
                StatusText        = "Windows OS Patches and Updates: No Updates Pending."
                HasPendingUpdates = $false
            }
        }

        $pendingCount = ($pendingUpdates | Measure-Object).Count
        $statusMessage = "Windows OS Patches and Updates: $pendingCount update(s) pending (restart required)."

        return [PSCustomObject]@{
            StatusText        = $statusMessage
            HasPendingUpdates = $true
        }
    }
    catch {
        Write-Log "Could not retrieve ECM update status. Client may not be installed. Error: $($_.Exception.Message)" -Level "INFO"
        return [PSCustomObject]@{
            StatusText        = "Windows OS Patches and Updates: Client not found or inaccessible."
            HasPendingUpdates = $false
        }
    }
}

function Test-LastWeekOfMonth {
    $today = Get-Date
    $lastDayOfMonth = [DateTime]::DaysInMonth($today.Year, $today.Month)
    $daysRemaining = $lastDayOfMonth - $today.Day
    return $daysRemaining -le 6
}

function Get-DaysSinceLastDriverUpdate {
    try {
        $logPath = "C:\Windows\mitll\Logs\MS_Update.txt"
        if (-not (Test-Path $logPath)) {
            return "Never run"
        }
        
        $content = Get-Content $logPath -ErrorAction Stop
        $dates = @()
        
        foreach ($line in $content) {
            # Match date format MM/DD/YYYY
            if ($line -match '(\d{2}/\d{2}/\d{4})') {
                try {
                    $date = [DateTime]::Parse($matches[1])
                    $dates += $date
                } catch {
                    # Skip invalid dates
                }
            }
        }
        
        if ($dates.Count -eq 0) {
            return "Never run"
        }
        
        $mostRecent = ($dates | Sort-Object -Descending | Select-Object -First 1)
        $daysSince = ((Get-Date) - $mostRecent).Days
        
        if ($daysSince -eq 0) {
            return "Last run today"
        } elseif ($daysSince -eq 1) {
            return "Last run 1 day ago"
        } else {
            return "Last run $daysSince days ago"
        }
    }
    catch {
        Write-Log "Error reading driver update log: $($_.Exception.Message)" -Level "ERROR"
        return "Status unknown"
    }
}

function Update-PatchingAndSystem {
    Write-Log "Updating Patching and System section..." -Level "INFO"
    
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
            }
        }
    } catch {
        $bigfixStatusText = "Application Updates: Error reading update data."
        Write-Log "Error reading BigFix data: $($_.Exception.Message)" -Level "ERROR"
    }

    $ecmResult = Get-ECMUpdateStatus
    $ecmStatusText = $ecmResult.StatusText
    $showEcmButton = $ecmResult.HasPendingUpdates
    
    # Check if we're in the last week of the month for driver updates
    $showDriverButton = $true
	#$showDriverButton = Test-LastWeekOfMonth

	# Get days since last driver update run
	$driverLastRun = Get-DaysSinceLastDriverUpdate

    # Create current state string
    $currentPatchingState = "BigFix:$showBigFixButton|ECM:$showEcmButton|Driver:$showDriverButton"
    
    # Check if state has changed
    $patchingStateChanged = $config.LastPatchingState -ne $currentPatchingState
    
    if ($patchingStateChanged) {
        Write-Log "Patching state changed. Previous: $($config.LastPatchingState), Current: $currentPatchingState" -Level "INFO"
    }
    
    try {
        $window.Dispatcher.Invoke({
            Write-Log "Inside Dispatcher.Invoke for patching update" -Level "INFO"
            
            $global:BigFixStatusText.FontWeight = "Bold"
            $global:BigFixStatusText.Text = $bigfixStatusText
            $global:BigFixLaunchButton.Visibility = if ($showBigFixButton) { "Visible" } else { "Collapsed" }
            Write-Log "Set BigFix UI - Button visible: $showBigFixButton" -Level "INFO"
            
            $global:ECMStatusText.FontWeight = "Bold"
            $global:ECMStatusText.Text = $ecmStatusText
            $global:ECMLaunchButton.Visibility = if ($showEcmButton) { "Visible" } else { "Collapsed" }
            Write-Log "Set ECM UI - Button visible: $showEcmButton" -Level "INFO"
            
            $global:DriverUpdateButton.Visibility = if ($showDriverButton) { "Visible" } else { "Collapsed" }
            Write-Log "Set Driver UI - Button visible: $showDriverButton" -Level "INFO"
            
			$global:DriverUpdateButton.Visibility = if ($showDriverButton) { "Visible" } else { "Collapsed" }
			$global:DriverUpdateLastRunText.Text = $driverLastRun
			Write-Log "Set Driver UI - Button visible: $showDriverButton, Last run: $driverLastRun" -Level "INFO"
			
            # Show red dot if state changed and at least one button is visible
            Write-Log "Checking alert conditions: StateChanged=$patchingStateChanged, AnyButton=$($showBigFixButton -or $showEcmButton -or $showDriverButton)" -Level "INFO"
            if ($patchingStateChanged -and ($showBigFixButton -or $showEcmButton -or $showDriverButton)) {
                Write-Log "Attempting to show patching alert dot. PatchingAlertDot exists: $($global:PatchingAlertDot -ne $null)" -Level "INFO"
                if ($global:PatchingAlertDot) { 
                    $global:PatchingAlertDot.Visibility = "Visible" 
                    Write-Log "Patching alert dot set to Visible" -Level "INFO"
                } else {
                    Write-Log "PatchingAlertDot is null at runtime!" -Level "WARNING"
                }
                if ($global:ClearAlertsDot) { $global:ClearAlertsDot.Visibility = "Visible" }
            }
        })
    } catch {
        Write-Log "ERROR in Dispatcher.Invoke for patching update: $($_.Exception.Message)" -Level "ERROR"
    }
    
   
}

function Convert-MarkdownToTextBlock {
    param(
        [string]$Text,
        [System.Windows.Controls.TextBlock]$TargetTextBlock
    )
    
    try {
        if (-not $Text -or -not $TargetTextBlock) {
            $TargetTextBlock.Inlines.Clear()
            if ($Text) { $TargetTextBlock.Inlines.Add((New-Object System.Windows.Documents.Run($Text))) }
            return
        }

        $TargetTextBlock.Inlines.Clear()
        
        $regexColor = "\[(green|red|yellow|blue)\](.*?)\[/\1\]"
        $regexBold = "\*\*(.*?)\*\*"
        $regexItalic = "\*(.*?)\*"
        $regexUnderline = "__(.*?)__"
        
        $currentText = $Text
        $colorPlaceholders = @{}
        $placeholderCounter = 0
        $colorMatches = [regex]::Matches($Text, $regexColor) | Sort-Object Index -Descending
        
        foreach ($match in $colorMatches) {
            $placeholder = "{COLORPH$placeholderCounter}"
            $leftOk  = ($match.Index -ge 2) -and ($Text.Substring($match.Index - 2, 2) -eq "**")
			$rightOk = (($match.Index + $match.Length + 2) -le $Text.Length) -and ($Text.Substring($match.Index + $match.Length, 2) -eq "**")
			$isBold  = $leftOk -and $rightOk
            $colorPlaceholders[$placeholder] = @{
                Text = $match.Groups[2].Value
                Color = $match.Groups[1].Value
                IsBold = $isBold
            }
            $currentText = $currentText.Remove($match.Index, $match.Length).Insert($match.Index, $placeholder)
            $placeholderCounter++
        }
        
        $matches = @()
        $boldMatches = [regex]::Matches($currentText, $regexBold)
        $italicMatches = [regex]::Matches($currentText, $regexItalic)
        $underlineMatches = [regex]::Matches($currentText, $regexUnderline)
        
        foreach ($match in $boldMatches) { $matches += [PSCustomObject]@{ Index = $match.Index; Length = $match.Length; Text = $match.Groups[1].Value; Type = "Bold" } }
        foreach ($match in $italicMatches) { if ([string]::IsNullOrWhiteSpace($match.Groups[1].Value)) { continue }; $matches += [PSCustomObject]@{ Index = $match.Index; Length = $match.Length; Text = $match.Groups[1].Value; Type = "Italic" } }
        foreach ($match in $underlineMatches) { $matches += [PSCustomObject]@{ Index = $match.Index; Length = $match.Length; Text = $match.Groups[1].Value; Type = "Underline" } }
        
        $matches = $matches | Sort-Object Index
        $lastIndex = 0
        $runs = @()

        foreach ($match in $matches) {
            if ($match.Index -gt $lastIndex) {
                $plainText = $currentText.Substring($lastIndex, $match.Index - $lastIndex)
                $runs += Process-TextSegment -Text $plainText -ColorPlaceholders $colorPlaceholders
            }
            
            $text = $match.Text
            if ($colorPlaceholders.ContainsKey($text)) {
                $colorInfo = $colorPlaceholders[$text]
                $innerRuns = Process-InnerMarkdown -Text $colorInfo.Text -Color $colorInfo.Color -IsBold $colorInfo.IsBold
                $runs += $innerRuns
            } else {
                $run = New-Object System.Windows.Documents.Run($text)
                if ($match.Type -eq "Bold") { $run.FontWeight = [System.Windows.FontWeights]::Bold }
                elseif ($match.Type -eq "Italic") { $run.FontStyle = [System.Windows.FontStyles]::Italic }
                elseif ($match.Type -eq "Underline") { $run.TextDecorations = [System.Windows.TextDecorations]::Underline }
                $runs += $run
            }
            $lastIndex = $match.Index + $match.Length
        }
        
        if ($lastIndex -lt $currentText.Length) {
            $plainText = $currentText.Substring($lastIndex)
            $runs += Process-TextSegment -Text $plainText -ColorPlaceholders $colorPlaceholders
        }
        
        foreach ($run in $runs) {
            $TargetTextBlock.Inlines.Add($run)
        }
    } catch {
        Write-Log "Failed to parse Markdown for text: $Text - $($_.Exception.Message)" -Level "ERROR"
        $TargetTextBlock.Inlines.Clear()
        $TargetTextBlock.Inlines.Add((New-Object System.Windows.Documents.Run($Text)))
    }
}

function Process-TextSegment {
    param(
        [string]$Text,
        [hashtable]$ColorPlaceholders
    )
    
    $runs = @()
    $currentPos = 0
    $placeholderRegex = [regex] "{COLORPH\d+}"
    $placeholderMatches = $placeholderRegex.Matches($Text) | Sort-Object Index
    
    foreach ($match in $placeholderMatches) {
        if ($match.Index -gt $currentPos) {
            $plainText = $Text.Substring($currentPos, $match.Index - $currentPos)
            $runs += New-Object System.Windows.Documents.Run($plainText)
        }
        
        $placeholder = $match.Value
        if ($colorPlaceholders.ContainsKey($placeholder)) {
            $colorInfo = $colorPlaceholders[$placeholder]
            $innerRuns = Process-InnerMarkdown -Text $colorInfo.Text -Color $colorInfo.Color -IsBold $colorInfo.IsBold
            $runs += $innerRuns
        }
        
        $currentPos = $match.Index + $match.Length
    }
    
    if ($currentPos -lt $Text.Length) {
        $plainText = $Text.Substring($currentPos)
        $runs += New-Object System.Windows.Documents.Run($plainText)
    }
    
    return $runs
}

function Process-InnerMarkdown {
    param(
        [string]$Text,
        [string]$Color,
        [bool]$IsBold
    )
    
    try {
        if (-not $Text) { return @() }

        $runs = @()
        $regexBold = "\*\*(.*?)\*\*"
        $regexItalic = "\*(.*?)\*"
        $regexUnderline = "__(.*?)__"
        
        $matches = @()
        $boldMatches = [regex]::Matches($Text, $regexBold)
        $italicMatches = [regex]::Matches($Text, $regexItalic)
        $underlineMatches = [regex]::Matches($Text, $regexUnderline)
        
        foreach ($match in $boldMatches) { $matches += [PSCustomObject]@{ Index = $match.Index; Length = $match.Length; Text = $match.Groups[1].Value; Type = "Bold" } }
        foreach ($match in $italicMatches) { if ([string]::IsNullOrWhiteSpace($match.Groups[1].Value)) { continue }; $matches += [PSCustomObject]@{ Index = $match.Index; Length = $match.Length; Text = $match.Groups[1].Value; Type = "Italic" } }
        foreach ($match in $underlineMatches) { $matches += [PSCustomObject]@{ Index = $match.Index; Length = $match.Length; Text = $match.Groups[1].Value; Type = "Underline" } }
        
        $matches = $matches | Sort-Object Index
        $lastIndex = 0

        foreach ($match in $matches) {
            if ($match.Index -gt $lastIndex) {
                $plainText = $Text.Substring($lastIndex, $match.Index - $lastIndex)
                $run = New-Object System.Windows.Documents.Run($plainText)
                if ($Color) { $run.Foreground = [System.Windows.Media.Brushes]::$Color }
                if ($IsBold) { $run.FontWeight = [System.Windows.FontWeights]::Bold }
                $runs += $run
            }
            
            $run = New-Object System.Windows.Documents.Run($match.Text)
            if ($match.Type -eq "Bold") { $run.FontWeight = [System.Windows.FontWeights]::Bold }
            elseif ($match.Type -eq "Italic") { $run.FontStyle = [System.Windows.FontStyles]::Italic }
            elseif ($match.Type -eq "Underline") { $run.TextDecorations = [System.Windows.TextDecorations]::Underline }
            if ($Color) { $run.Foreground = [System.Windows.Media.Brushes]::$Color }
            if ($IsBold) { $run.FontWeight = [System.Windows.FontWeights]::Bold }
            $runs += $run
            $lastIndex = $match.Index + $match.Length
        }
        
        if ($lastIndex -lt $Text.Length) {
            $plainText = $Text.Substring($lastIndex)
            $run = New-Object System.Windows.Documents.Run($plainText)
            if ($Color) { $run.Foreground = [System.Windows.Media.Brushes]::$Color }
            if ($IsBold) { $run.FontWeight = [System.Windows.FontWeights]::Bold }
            $runs += $run
        }
        return $runs
    } catch {
        $run = New-Object System.Windows.Documents.Run($Text)
        if ($Color) { $run.Foreground = [System.Windows.Media.Brushes]::$Color }
        if ($IsBold) { $run.FontWeight = [System.Windows.FontWeights]::Bold }
        return @($run)
    }
}

function Update-Announcements {
    Write-Log "Updating Announcements section..." -Level "INFO"
	Write-Log "contentData.Data type: $($global:contentData.Data.GetType().Name)" -Level "INFO"
    Write-Log "contentData.Data has Dashboard: $($global:contentData.Data.Dashboard -ne $null)" -Level "INFO"
    $announcementsData = if ($global:contentData.Data.Dashboard) {
        Write-Log "Using new Dashboard structure" -Level "INFO"
		$global:contentData.Data.Dashboard.Announcements
    } else {
		Write-Log "Using old flat structure" -Level "INFO"
        # Fallback for old structure
        $global:contentData.Data.Announcements
    }
	Write-Log "announcementsData is null: $($announcementsData -eq $null)" -Level "INFO"
    if ($announcementsData) {
        Write-Log "announcementsData.Default exists: $($announcementsData.Default -ne $null)" -Level "INFO"
    }
    
    $announcementData = Get-ActiveAnnouncement -AnnouncementsObject $announcementsData
    if (-not $announcementData) { return }
    
    $baseMessage = $announcementData.Base
    $appendedMessages = $announcementData.Appended
    $compositeObjectForStateCheck = @{ Base = $baseMessage; Appended = $appendedMessages }
    $global:LastAnnouncementState = $compositeObjectForStateCheck | ConvertTo-Json -Compress
    
    $isNew = $config.AnnouncementsLastState -ne $global:LastAnnouncementState
    if ($isNew) { Write-Log "New announcement content detected." -Level "INFO" }

    $window.Dispatcher.Invoke({
        if ($isNew) {
    if ($global:AnnouncementsAlertIcon) { $global:AnnouncementsAlertIcon.Visibility = "Visible" }
    if ($global:DashboardTabAlert) { $global:DashboardTabAlert.Visibility = "Visible" }
    if ($global:ClearAlertsDot) { $global:ClearAlertsDot.Visibility = "Visible" }
}

        Convert-MarkdownToTextBlock -Text $baseMessage.Text -TargetTextBlock $global:AnnouncementsText
        Convert-MarkdownToTextBlock -Text $baseMessage.Details -TargetTextBlock $global:AnnouncementsDetailsText

        $global:AppendedAnnouncementsPanel.Children.Clear()
        if ($appendedMessages.Count -gt 0) {
            $global:AppendedAnnouncementsPanel.Visibility = "Visible"
            foreach ($message in $appendedMessages) {
                $separator = New-Object System.Windows.Controls.Separator; $separator.Margin = [System.Windows.Thickness]::new(0,10,0,10)
                $global:AppendedAnnouncementsPanel.Children.Add($separator)
                if ($message.Text) {
                    $appendedText = New-Object System.Windows.Controls.TextBlock; $appendedText.FontSize = 11; $appendedText.TextWrapping = "Wrap"
                    Convert-MarkdownToTextBlock -Text $message.Text -TargetTextBlock $appendedText
                    $global:AppendedAnnouncementsPanel.Children.Add($appendedText)
                }
                if ($message.Details) {
                    $appendedDetails = New-Object System.Windows.Controls.TextBlock; $appendedDetails.FontSize = 11; $appendedDetails.TextWrapping = "Wrap"; $appendedDetails.Margin = [System.Windows.Thickness]::new(0,5,0,0)
                    Convert-MarkdownToTextBlock -Text $message.Details -TargetTextBlock $appendedDetails
                    $global:AppendedAnnouncementsPanel.Children.Add($appendedDetails)
                }
            }
        } else {
            $global:AppendedAnnouncementsPanel.Visibility = "Collapsed"
        }

        $allLinks = [System.Collections.Generic.List[object]]::new()
        if ($baseMessage.Links) { $allLinks.AddRange($baseMessage.Links) }
        foreach ($message in $appendedMessages) { if ($message.Links) { $allLinks.AddRange($message.Links) } }

        $global:AnnouncementsLinksPanel.Children.Clear()
        if ($allLinks.Count -gt 0) {
            foreach ($link in $allLinks) {
                $global:AnnouncementsLinksPanel.Children.Add((New-HyperlinkBlock -Name $link.Name -Url $link.Url))
            }
        }
        $global:AnnouncementsSourceText.Text = "Source: $($global:contentData.Source)"
    })
}

function Update-Support {
    Write-Log "Updating Support section..." -Level "INFO"
    $newSupportObject = if ($global:contentData.Data.Dashboard) {
        $global:contentData.Data.Dashboard.Support
    } else {
        $global:contentData.Data.Support
    }
    if (-not $newSupportObject) { return }

    $newJsonState = $newSupportObject | ConvertTo-Json -Compress -Depth 10
    $isNew = $config.SupportLastState -ne $newJsonState
    
    # Debug logging
    if ($isNew) { 
        Write-Log "New support content detected." -Level "INFO"
        Write-Log "Previous state: $($config.SupportLastState)" -Level "INFO"
        Write-Log "New state: $newJsonState" -Level "INFO"
    }

    $window.Dispatcher.Invoke({
        if ($isNew) {
            if ($global:SupportTabAlert) { $global:SupportTabAlert.Visibility = "Visible" }
            if ($global:ClearAlertsDot) { $global:ClearAlertsDot.Visibility = "Visible" }
        }
        
        Convert-MarkdownToTextBlock -Text $newSupportObject.Text -TargetTextBlock $global:SupportText
        $global:SupportLinksPanel.Children.Clear()
        if ($newSupportObject.Links) {
            foreach ($link in $newSupportObject.Links) {
                $global:SupportLinksPanel.Children.Add((New-HyperlinkBlock -Name $link.Name -Url $link.Url))
            }
        }
        $global:SupportSourceText.Text = "Source: $($global:contentData.Source)"
    })
    
    # Update the saved state after processing
    if ($isNew) {
        $config.SupportLastState = $newJsonState
        Save-Configuration -Config $config
    }
}

function Create-DynamicTabs {
    if ($global:contentData.Data.AdditionalTabs) {
        foreach ($tabConfig in $global:contentData.Data.AdditionalTabs) {
            if (-not $tabConfig.Enabled) { continue }
            
            $tabItem = New-Object System.Windows.Controls.TabItem
            $tabItem.Header = $tabConfig.TabHeader
            
            $scrollViewer = New-Object System.Windows.Controls.ScrollViewer
            $scrollViewer.VerticalScrollBarVisibility = "Auto"
            
            $stackPanel = New-Object System.Windows.Controls.StackPanel
            $stackPanel.Margin = [System.Windows.Thickness]::new(10)
            
            if ($tabConfig.Content.Text) {
                $textBlock = New-Object System.Windows.Controls.TextBlock
                $textBlock.TextWrapping = "Wrap"
                $textBlock.FontSize = 11
                $textBlock.Margin = [System.Windows.Thickness]::new(0,0,0,10)
                Convert-MarkdownToTextBlock -Text $tabConfig.Content.Text -TargetTextBlock $textBlock
                $stackPanel.Children.Add($textBlock)
            }
            
            if ($tabConfig.Content.Sections) {
                foreach ($section in $tabConfig.Content.Sections) {
                    $expander = New-Object System.Windows.Controls.Expander
                    $expander.Header = $section.Title
                    $expander.IsExpanded = $true
                    $expander.Margin = [System.Windows.Thickness]::new(0,5,0,5)
                    
                    $sectionStack = New-Object System.Windows.Controls.StackPanel
                    
                    if ($section.Text) {
                        $tb = New-Object System.Windows.Controls.TextBlock
                        $tb.TextWrapping = "Wrap"
                        $tb.FontSize = 11
                        Convert-MarkdownToTextBlock -Text $section.Text -TargetTextBlock $tb
                        $sectionStack.Children.Add($tb)
                    }
                    
                    if ($section.Details) {
                        $tbDetails = New-Object System.Windows.Controls.TextBlock
                        $tbDetails.TextWrapping = "Wrap"
                        $tbDetails.FontSize = 11
                        $tbDetails.Margin = [System.Windows.Thickness]::new(0,5,0,0)
                        Convert-MarkdownToTextBlock -Text $section.Details -TargetTextBlock $tbDetails
                        $sectionStack.Children.Add($tbDetails)
                    }
                    
                    if ($section.Links) {
                        $linkPanel = New-Object System.Windows.Controls.StackPanel
                        $linkPanel.Margin = [System.Windows.Thickness]::new(0,5,0,0)
                        foreach ($link in $section.Links) {
                            $linkPanel.Children.Add((New-HyperlinkBlock -Name $link.Name -Url $link.Url))
                        }
                        $sectionStack.Children.Add($linkPanel)
                    }
                    
                    $expander.Content = $sectionStack
                    $stackPanel.Children.Add($expander)
                }
            }
            
            if ($tabConfig.Content.Links) {
                $linkPanel = New-Object System.Windows.Controls.StackPanel
                $linkPanel.Margin = [System.Windows.Thickness]::new(0,10,0,0)
                foreach ($link in $tabConfig.Content.Links) {
                    $linkPanel.Children.Add((New-HyperlinkBlock -Name $link.Name -Url $link.Url))
                }
                $stackPanel.Children.Add($linkPanel)
            }
            
            $scrollViewer.Content = $stackPanel
            $tabItem.Content = $scrollViewer
            
            $window.Dispatcher.Invoke({
                $global:MainTabControl.Items.Add($tabItem)
            })
        }
    }
}

function Update-LogsTab {
    try {
        if (Test-Path $LogFilePath) {
            # Read last 100 lines for performance
            $logContent = Get-Content $LogFilePath -Tail 100
            $window.Dispatcher.Invoke({
                $global:LogTextBox.Text = $logContent -join "`r`n"
                $global:LogTextBox.ScrollToEnd()
            })
        }
    } catch {
        # Fail silently
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
        try { return New-Object System.Drawing.Icon($Path) }
        catch { Write-Log "Error loading icon from `"$Path`" - $($_.Exception.Message)" -Level "ERROR" }
    }
    return [System.Drawing.SystemIcons]::Application
}

function Update-TrayIcon {
    if (-not $global:TrayIcon.Visible) { return }
    
    $announcementAlert = $global:AnnouncementsAlertIcon -and $global:AnnouncementsAlertIcon.Visibility -eq "Visible"
    $dashboardTabAlert = $global:DashboardTabAlert -and $global:DashboardTabAlert.Visibility -eq "Visible"
    $supportTabAlert = $global:SupportTabAlert -and $global:SupportTabAlert.Visibility -eq "Visible"
    $patchingAlert = $global:PatchingAlertDot -and $global:PatchingAlertDot.Visibility -eq "Visible"
    $hasAnyAlert = $announcementAlert -or $dashboardTabAlert -or $supportTabAlert -or $patchingAlert

    $global:TrayIcon.Icon = if ($hasAnyAlert) { $global:WarningIcon } else { $global:MainIcon }
    $global:TrayIcon.Text = if ($hasAnyAlert) { "Endpoint Advisor v$ScriptVersion - Alerts Pending" } else { "Lincoln Laboratory Endpoint Advisor v$ScriptVersion" }
}

function Initialize-TrayIcon {
    if (-not $global:FormsAvailable) { return }
    try {
        $global:MainIcon = Get-Icon -Path $config.IconPaths.Main
        $global:WarningIcon = Get-Icon -Path $config.IconPaths.Warning

        $global:TrayIcon = New-Object System.Windows.Forms.NotifyIcon
        $global:TrayIcon.Icon = $global:MainIcon
        $global:TrayIcon.Text = "Lincoln Laboratory Endpoint Advisor v$ScriptVersion"
        $global:TrayIcon.Visible = $true

        $ContextMenuStrip = New-Object System.Windows.Forms.ContextMenuStrip
        
        $intervalSubMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Set Update Interval")
        $tenMin = New-Object System.Windows.Forms.ToolStripMenuItem("10 minutes", $null, { $config.RefreshInterval = 600; Save-Configuration -Config $config; $global:DispatcherTimer.Interval = [TimeSpan]::FromSeconds(600); $global:DispatcherTimer.Stop(); $global:DispatcherTimer.Start() })
        $fifteenMin = New-Object System.Windows.Forms.ToolStripMenuItem("15 minutes", $null, { $config.RefreshInterval = 900; Save-Configuration -Config $config; $global:DispatcherTimer.Interval = [TimeSpan]::FromSeconds(900); $global:DispatcherTimer.Stop(); $global:DispatcherTimer.Start() })
        $twentyMin = New-Object System.Windows.Forms.ToolStripMenuItem("20 minutes", $null, { $config.RefreshInterval = 1200; Save-Configuration -Config $config; $global:DispatcherTimer.Interval = [TimeSpan]::fromSeconds(1200); $global:DispatcherTimer.Stop(); $global:DispatcherTimer.Start() })
        $intervalSubMenu.DropDownItems.AddRange(@($tenMin, $fifteenMin, $twentyMin))

        $ContextMenuStrip.Items.AddRange(@(
            (New-Object System.Windows.Forms.ToolStripMenuItem("Show Dashboard", $null, { Toggle-WindowVisibility })),
            (New-Object System.Windows.Forms.ToolStripMenuItem("Refresh Now", $null, { Main-UpdateCycle -ForceCertificateCheck $true })),
            $intervalSubMenu
            
        ))
        $global:TrayIcon.ContextMenuStrip = $ContextMenuStrip
        $global:TrayIcon.add_MouseClick({
            if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) { Toggle-WindowVisibility }
        })
    } catch { Handle-Error $_.Exception.Message -Source "Initialize-TrayIcon" }
}

# ============================================================
# K) Window Visibility Management
# ============================================================
function Toggle-WindowVisibility {
    $window.Dispatcher.Invoke({
        if ($window.IsVisible) {
            $window.Hide()
        } else {
            Update-LogsTab # Refresh logs when window is shown
            $window.Show()
            $global:BlinkingTimer.Stop()
            $window.Activate()
        }
        Update-TrayIcon
    })
}

# ============================================================
# O) Main Update Cycle and DispatcherTimer
# ============================================================
function Main-UpdateCycle {
    param([bool]$ForceCertificateCheck = $false)
    try {
        Write-Log "Main content update cycle running..." -Level "INFO"
        $global:contentData = Fetch-ContentData
        
        Update-Announcements
        Update-Support
        Update-PatchingAndSystem
        
		if (-not $global:DynamicTabsCreated) {
            Create-DynamicTabs
            $global:DynamicTabsCreated = $true
            Write-Log "Dynamic tabs created from JSON configuration." -Level "INFO"
        }
		
        if ($ForceCertificateCheck -or (-not $global:LastCertificateCheck -or ((Get-Date) - $global:LastCertificateCheck).TotalSeconds -ge $config.CertificateCheckInterval)) {
            Update-CertificateInfo
            $global:LastCertificateCheck = Get-Date
        }
        
        if ($window.IsVisible) {
            Update-LogsTab
        }
        
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

    $global:mainTickAction = { param($sender, $e) Main-UpdateCycle }
    $global:DispatcherTimer = New-Object System.Windows.Threading.DispatcherTimer
    $global:DispatcherTimer.Interval = [TimeSpan]::FromSeconds($config.RefreshInterval)
    $global:DispatcherTimer.add_Tick($global:mainTickAction)

    Initialize-TrayIcon
    Log-DotNetVersion
    Main-UpdateCycle -ForceCertificateCheck $true
    
    if (-not $config.HasRunBefore) {
        Write-Log "First run detected. Showing balloon tip notification." -Level "INFO"
        $global:TrayIcon.ShowBalloonTip(10000, "Endpoint Advisor is Running", "The notification icon is active in your system tray. You may need to drag it from the overflow area (^) to the main taskbar.", [System.Windows.Forms.ToolTipIcon]::Info)
        $config.HasRunBefore = $true
        Save-Configuration -Config $config
    }
    
    $global:DispatcherTimer.Start()
    Write-Log "Main timer started." -Level "INFO"
    
    $window.Dispatcher.Add_UnhandledException({ Handle-Error $_.Exception.Message -Source "Dispatcher"; $_.Handled = $true })

    Write-Log "Application startup complete. Running dispatcher." -Level "INFO"
    [System.Windows.Threading.Dispatcher]::Run()
}
catch {
    Handle-Error "A critical error occurred during startup: $($_.Exception.Message)" -Source "Startup"
}
finally {
    Write-Log "--- Lincoln Laboratory Endpoint Advisor Script Exiting ---"
    if ($global:DispatcherTimer) { $global:DispatcherTimer.Stop() }
    if ($global:TrayIcon) { $global:TrayIcon.Dispose() }
    if ($global:MainIcon) { $global:MainIcon.Dispose() }
    if ($global:WarningIcon) { $global:WarningIcon.Dispose() }
}
