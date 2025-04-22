# LLNOTIFY.ps1 - Lincoln Laboratory Notification System

# Ensure $PSScriptRoot is defined for older versions
if ($MyInvocation.MyCommand.Path) {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    $ScriptDir = Get-Location
}

# Define version
$ScriptVersion = "1.1.0"

# ============================================================
# A) Advanced Logging & Error Handling
# ============================================================
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )
    $logPath = if ($LogFilePath) { $LogFilePath } else { Join-Path $ScriptDir "LLNOTIFY.log" }
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    try {
        Add-Content -Path $logPath -Value $logEntry -ErrorAction Stop
    }
    catch {
        Write-Host "[$timestamp] [$Level] $Message (Failed to write to log: $_)"
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

Write-Log "Script directory resolved as: $ScriptDir" -Level "INFO"

# ============================================================
# MODULE: Configuration Management
# ============================================================
function Get-DefaultConfig {
    return @{
        RefreshInterval       = 90
        LogRotationSizeMB     = 5
        DefaultLogLevel       = "INFO"
        ContentDataUrl        = "https://raw.githubusercontent.com/burnoil/MITSI/main/ContentData.json"
        ContentFetchInterval  = 120
        YubiKeyAlertDays      = 14
        IconPaths             = @{
            Main    = Join-Path $ScriptDir "healthy.ico"
            Warning = Join-Path $ScriptDir "warning.ico"
        }
        YubiKeyLastCheck      = @{
            Date   = "1970-01-01 00:00:00"
            Result = "YubiKey Certificate: Not yet checked"
        }
        AnnouncementsLastState = @{}
        SupportLastState       = @{}
        Version               = $ScriptVersion
        PatchInfoFilePath     = "C:\temp\patch_fixlets.txt"
    }
}

function Load-Configuration {
    param(
        [string]$Path = (Join-Path $ScriptDir "LLNOTIFY.config.json")
    )
    $defaultConfig = Get-DefaultConfig
    if (Test-Path $Path) {
        try {
            $config = Get-Content $Path -Raw | ConvertFrom-Json
            Write-Log "Loaded config from $Path" -Level "INFO"
            # Warn if ContentDataUrl contains query parameters
            if ($config.ContentDataUrl -match "\?") {
                Write-Log "Warning: ContentDataUrl contains query parameters ('$($config.ContentDataUrl)'). Remove parameters like '?token=' for reliable fetching. Consider updating LLNOTIFY.config.json." -Level "WARNING"
            }
            foreach ($key in $defaultConfig.Keys) {
                if (-not $config.PSObject.Properties.Match($key)) {
                    $config | Add-Member -NotePropertyName $key -NotePropertyValue $defaultConfig[$key]
                }
            }
            return $config
        }
        catch {
            Write-Log "Error loading config, reverting to default: $_" -Level "ERROR"
            return $defaultConfig
        }
    }
    else {
        $defaultConfig | ConvertTo-Json -Depth 3 | Out-File $Path -Force
        Write-Log "Created default config at $Path" -Level "INFO"
        return $defaultConfig
    }
}

function Save-Configuration {
    param(
        [psobject]$Config,
        [string]$Path = (Join-Path $ScriptDir "LLNOTIFY.config.json")
    )
    $Config | ConvertTo-Json -Depth 3 | Out-File $Path -Force
}

# ============================================================
# MODULE: Performance Optimizations – Caching
# ============================================================
$global:LastContentFetch = $null
$global:CachedContentData = $null

# ============================================================
# B) External Configuration Setup
# ============================================================
$LogFilePath = Join-Path $ScriptDir "LLNOTIFY.log"
$config = Load-Configuration

# Use healthy.ico for the normal state, warning.ico when anything alerts.
$config.IconPaths.Main    = Join-Path $ScriptDir "healthy.ico"
$config.IconPaths.Warning = Join-Path $ScriptDir "warning.ico"

$mainIconPath = $config.IconPaths.Main
$warningIconPath = $config.IconPaths.Warning
$mainIconUri = "file:///" + ($mainIconPath -replace '\\','/')

Write-Log "Main icon path: $mainIconPath" -Level "INFO"
Write-Log "Warning icon path: $warningIconPath" -Level "INFO"
Write-Log "Main icon URI: $mainIconUri" -Level "INFO"
Write-Log "Main icon exists: $(Test-Path $mainIconPath)" -Level "INFO"
Write-Log "Warning icon exists: $(Test-Path $warningIconPath)" -Level "INFO"

# Default content data (in case the external content is missing)
$defaultContentData = @{
    Announcements = @{
        Text    = "No announcements at this time."
        Details = "Check back later for updates."
        Links   = @(
            @{ Name = "Announcement Link 1"; Url = "https://company.com/news1" },
            @{ Name = "Announcement Link 2"; Url = "https://company.com/news2" }
        )
    }
    Support = @{
        Text  = "Contact IT Support: support@company.com | Phone: 1-800-555-1234"
        Links = @(
            @{ Name = "Support Link 1"; Url = "https://support.company.com/help" },
            @{ Name = "Support Link 2"; Url = "https://support.company.com/tickets" }
        )
    }
}

# ============================================================
# C) Log File Setup & Rotation
# ============================================================
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

function Export-Logs {
    try {
        if (-not $global:FormsAvailable) {
            Write-Log "Export-Logs requires System.Windows.Forms for SaveFileDialog. Feature unavailable." -Level "WARNING"
            return
        }
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "Log Files (*.log)|*.log|All Files (*.*)|*.*"
        $saveFileDialog.FileName = "LLNOTIFY.log"
        if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Copy-Item -Path $LogFilePath -Destination $saveFileDialog.FileName -Force
            Write-Log "Logs exported to $($saveFileDialog.FileName)" -Level "INFO"
        }
    }
    catch {
        Handle-Error "Error exporting logs: $_" -Source "Export-Logs"
    }
}

# ============================================================
# D) Import Required Assemblies
# ============================================================
function Import-RequiredAssemblies {
    try {
        # Check PowerShell version
        $psVersion = $PSVersionTable.PSVersion
        $isWindowsPowerShell = $PSVersionTable.PSEdition -eq "Desktop" -or $psVersion.Major -le 5
        Write-Log "PowerShell Version: $psVersion, Edition: $($PSVersionTable.PSEdition)" -Level "INFO"

        if (-not $isWindowsPowerShell) {
            Write-Log "Running in PowerShell Core/7. System.Windows.Forms may not be fully supported." -Level "WARNING"
        }

        # Load assemblies with error handling
        Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
        Write-Log "Loaded PresentationFramework assembly" -Level "INFO"

        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Write-Log "Loaded System.Windows.Forms assembly" -Level "INFO"

        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        Write-Log "Loaded System.Drawing assembly" -Level "INFO"

        return $true
    }
    catch {
        Write-Log "Failed to load required assemblies: $_" -Level "ERROR"
        if ($_.Exception.Message -match "System.Windows.Forms") {
            Write-Log "System.Windows.Forms is unavailable. Tray icon functionality will be disabled." -Level "WARNING"
            return $false
        }
        throw $_ # Re-throw if other assemblies fail
    }
}

# Import assemblies and store result
$global:FormsAvailable = Import-RequiredAssemblies

# ============================================================
# E) XAML Layout Definition with Visual Enhancements
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
        <Image Source="{Binding MainIconUri}" Width="20" Height="20" Margin="0,0,4,0"/>
        <TextBlock Text="Lincoln Laboratory Notification System"
                   FontSize="14" FontWeight="Bold" Foreground="White"
                   VerticalAlignment="Center"/>
      </StackPanel>
    </Border>
    <!-- Content Area -->
    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
      <StackPanel VerticalAlignment="Top">
        <!-- Announcements Section -->
        <Expander x:Name="AnnouncementsExpander" ToolTip="View latest announcements" FontSize="12" Foreground="#00008B" IsExpanded="True" Margin="0,2,0,2">
          <Expander.Header>
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="Announcements" VerticalAlignment="Center"/>
              <Ellipse x:Name="AnnouncementsAlertIcon" Width="10" Height="10" Margin="4,0,0,0" Fill="Red" Visibility="Hidden"/>
            </StackPanel>
          </Expander.Header>
          <Border BorderBrush="#00008B" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="AnnouncementsText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
              <TextBlock x:Name="AnnouncementsDetailsText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
              <StackPanel x:Name="AnnouncementsLinksPanel" Orientation="Vertical" Margin="2"/>
              <TextBlock x:Name="AnnouncementsSourceText" FontSize="9" Foreground="Gray" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
            </StackPanel>
          </Border>
        </Expander>
        <!-- Patching and Updates Section -->
        <Expander x:Name="PatchingExpander" ToolTip="View patching status" FontSize="12" Foreground="#00008B" IsExpanded="True" Margin="0,2,0,2">
          <Expander.Header>
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="Patching and Updates" VerticalAlignment="Center"/>
              <Button x:Name="PatchingSSAButton" Content="Launch Updates" Width="80" Height="20" Margin="4,0,0,0" ToolTip="Launch BigFix Self-Service Application for Updates"/>
            </StackPanel>
          </Expander.Header>
          <Border BorderBrush="#00008B" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="PatchingUpdatesText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
            </StackPanel>
          </Border>
        </Expander>
        <!-- Support Section -->
        <Expander x:Name="SupportExpander" ToolTip="Contact IT support" FontSize="12" Foreground="#00008B" IsExpanded="False" Margin="0,2,0,2">
          <Expander.Header>
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="Support" VerticalAlignment="Center"/>
              <Ellipse x:Name="SupportAlertIcon" Width="10" Height="10" Margin="4,0,0,0" Fill="Red" Visibility="Hidden"/>
            </StackPanel>
          </Expander.Header>
          <Border BorderBrush="#00008B" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="SupportText" FontSize="11" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
              <StackPanel x:Name="SupportLinksPanel" Orientation="Vertical" Margin="2"/>
              <TextBlock x:Name="SupportSourceText" FontSize="9" Foreground="Gray" Margin="2" TextWrapping="Wrap" MaxWidth="300"/>
            </StackPanel>
          </Border>
        </Expander>
        <!-- Compliance Section -->
        <Expander x:Name="ComplianceExpander" ToolTip="Certificate Status" FontSize="12" Foreground="#00008B" IsExpanded="False" Margin="0,2,0,2">
          <Expander.Header>
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="Certificate Status" VerticalAlignment="Center"/>
            </StackPanel>
          </Expander.Header>
          <Border BorderBrush="#00008B" BorderThickness="1" Padding="3" CornerRadius="2" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="YubiKeyComplianceText" FontSize="11" Margin="2" TextWrapping="Wrap"/>
            </StackPanel>
          </Border>
        </Expander>
      </StackPanel>
    </ScrollViewer>
    <!-- Footer Section -->
    <TextBlock Grid.Row="2" Text="© 2025 Lincoln Laboratory" FontSize="10" Foreground="Gray" HorizontalAlignment="Center" Margin="0,4,0,0"/>
  </Grid>
</Window>
"@

# ============================================================
# F) Load and Verify XAML
# ============================================================
$xmlDoc = New-Object System.Xml.XmlDocument
$xmlDoc.LoadXml($xamlString)
$reader = New-Object System.Xml.XmlNodeReader $xmlDoc
try {
    [System.Windows.Window]$global:window = [Windows.Markup.XamlReader]::Load($reader)
    $window.Width = 350
    $window.Height = 500
    $window.Left = 100
    $window.Top = 100
    $window.WindowState = 'Normal'
    Write-Log "Initial window setup: Left=100, Top=100, Width=350, Height=500, State=$($window.WindowState)" -Level "INFO"
    try {
        $mainIconPath = $config.IconPaths.Main
        $mainIconUri = "file:///" + ($mainIconPath -replace '\\','/')
        $window.DataContext = [PSCustomObject]@{ MainIconUri = [Uri]$mainIconUri }
        Write-Log "Setting window icon URI to: $mainIconUri" -Level "INFO"
        Write-Log "Window icon URI valid: $([Uri]::IsWellFormedUriString($mainIconUri, [UriKind]::Absolute))" -Level "INFO"
    }
    catch {
        Write-Log "Error setting window icon URI: $_" -Level "ERROR"
    }
    # Access UI Elements
    $global:AnnouncementsExpander = $window.FindName("AnnouncementsExpander")
    $global:AnnouncementsAlertIcon = $window.FindName("AnnouncementsAlertIcon")
    $global:AnnouncementsText = $window.FindName("AnnouncementsText")
    $global:AnnouncementsDetailsText = $window.FindName("AnnouncementsDetailsText")
    $global:AnnouncementsLinksPanel = $window.FindName("AnnouncementsLinksPanel")
    $global:AnnouncementsSourceText = $window.FindName("AnnouncementsSourceText")
    $global:PatchingExpander = $window.FindName("PatchingExpander")
    $global:PatchingUpdatesText = $window.FindName("PatchingUpdatesText")
    $global:PatchingSSAButton = $window.FindName("PatchingSSAButton")
    $global:SupportExpander = $window.FindName("SupportExpander")
    $global:SupportAlertIcon = $window.FindName("SupportAlertIcon")
    $global:SupportText = $window.FindName("SupportText")
    $global:SupportLinksPanel = $window.FindName("SupportLinksPanel")
    $global:SupportSourceText = $window.FindName("SupportSourceText")
    $global:ComplianceExpander = $window.FindName("ComplianceExpander")
    $global:YubiKeyComplianceText = $window.FindName("YubiKeyComplianceText")
    # Clear the red alert-dot when the Expander is opened
    if ($global:AnnouncementsExpander) {
        $global:AnnouncementsExpander.Add_Expanded({
            $window.Dispatcher.Invoke({
                if ($global:AnnouncementsAlertIcon) {
                    $global:AnnouncementsAlertIcon.Visibility = "Hidden"
                }
            })
            Write-Log "Announcements expander expanded, alert dot cleared." -Level "INFO"
        })
    }

    if ($global:SupportExpander) {
        $global:SupportExpander.Add_Expanded({
            $window.Dispatcher.Invoke({
                if ($global:SupportAlertIcon) {
                    $global:SupportAlertIcon.Visibility = "Hidden"
                }
            })
            Write-Log "Support expander expanded, alert dot cleared." -Level "INFO"
        })
    }

    # SSA Button Click Handler
    if ($global:PatchingSSAButton) {
        $global:PatchingSSAButton.Add_Click({
            try {
                $ssaPath = "C:\Program Files (x86)\BigFix Enterprise\BigFix Self Service Application\BigFixSSA.exe"
                if (Test-Path $ssaPath) {
                    Start-Process -FilePath $ssaPath
                    Write-Log "Launched BigFix Self-Service Application: $ssaPath" -Level "INFO"
                } else {
                    Write-Log "BigFix BigFixSSA.exe not found at $ssaPath" -Level "ERROR"
                }
            }
            catch {
                Handle-Error "Error launching BigFix BigFixSSA.exe: $_" -Source "PatchingSSAButton"
            }
        })
    }

    Write-Log "AnnouncementsText null? $($global:AnnouncementsText -eq $null)" -Level "INFO"
    Write-Log "AnnouncementsDetailsText null? $($global:AnnouncementsDetailsText -eq $null)" -Level "INFO"
    Write-Log "AnnouncementsLinksPanel null? $($global:AnnouncementsLinksPanel -eq $null)" -Level "INFO"
    Write-Log "AnnouncementsSourceText null? $($global:AnnouncementsSourceText -eq $null)" -Level "INFO"
    Write-Log "AnnouncementsExpander null? $($global:AnnouncementsExpander -eq $null)" -Level "INFO"
    Write-Log "AnnouncementsAlertIcon null? $($global:AnnouncementsAlertIcon -eq $null)" -Level "INFO"
    Write-Log "PatchingExpander null? $($global:PatchingExpander -eq $null)" -Level "INFO"
    Write-Log "PatchingUpdatesText null? $($global:PatchingUpdatesText -eq $null)" -Level "INFO"
    Write-Log "PatchingSSAButton null? $($global:PatchingSSAButton -eq $null)" -Level "INFO"
    Write-Log "SupportText null? $($global:SupportText -eq $null)" -Level "INFO"
    Write-Log "SupportLinksPanel null? $($global:SupportLinksPanel -eq $null)" -Level "INFO"
    Write-Log "SupportSourceText null? $($global:SupportSourceText -eq $null)" -Level "INFO"
    Write-Log "SupportExpander null? $($global:SupportExpander -eq $null)" -Level "INFO"
    Write-Log "SupportAlertIcon null? $($global:SupportAlertIcon -eq $null)" -Level "INFO"
    Write-Log "ComplianceExpander null? $($global:ComplianceExpander -eq $null)" -Level "INFO"
    Write-Log "YubiKeyComplianceText null? $($global:YubiKeyComplianceText -eq $null)" -Level "INFO"
}
catch {
    Handle-Error "Failed to load the XAML layout: $_" -Source "XAML"
    exit
}
if ($window -eq $null) {
    Handle-Error "Failed to load the XAML layout. Check the XAML syntax for errors." -Source "XAML"
    exit
}

$window.Add_Closing({
    param($sender, $eventArgs)
    try {
        if ($global:FormsAvailable -and $global:TrayIcon) {
            $eventArgs.Cancel = $true
            $window.Hide()
            Write-Log "Dashboard hidden via window closing event. Visibility=$($window.Visibility), IsVisible=$($window.IsVisible)" -Level "INFO"
        }
        else {
            # Exit application if tray icon is unavailable
            if ($global:DispatcherTimer) {
                $global:DispatcherTimer.Stop()
                Write-Log "DispatcherTimer stopped." -Level "INFO"
            }
            $window.Dispatcher.InvokeShutdown()
            Write-Log "Application exited via window closing (no tray icon)." -Level "INFO"
        }
    }
    catch {
        Handle-Error "Error handling window closing: $_" -Source "WindowClosing"
    }
})

# ============================================================
# Global Variables for Jobs, Caching, and Data
# ============================================================
$global:yubiKeyJob = $null
$global:contentData = $null
$global:announcementAlertActive = $false
$global:yubiKeyAlertShown = $false

# ============================================================
# H) Modularized System Information Functions
# ============================================================
function Fetch-ContentData {
    try {
        Write-Log "Config object: $($config | ConvertTo-Json -Depth 3)" -Level "INFO"
        $url = $config.ContentDataUrl
        Write-Log "Raw ContentDataUrl from config: '$url'" -Level "INFO"

        # Strip query parameters from URL
        $cleanUrl = ($url -split '\?')[0]
        if ($url -ne $cleanUrl) {
            Write-Log "Stripped query parameters from URL: '$cleanUrl'" -Level "INFO"
        }

        if ($global:LastContentFetch -and ((Get-Date) - $global:LastContentFetch).TotalSeconds -lt $config.ContentFetchInterval) {
            Write-Log "Using cached content data" -Level "INFO"
            return [PSCustomObject]@{
                Data   = $global:CachedContentData
                Source = "Cache"
            }
        }

        Write-Log "Attempting to fetch content from: $cleanUrl" -Level "INFO"
        if ($cleanUrl -match "^(?i)(http|https)://") {
            Write-Log "Detected HTTP/HTTPS URL, using Invoke-WebRequest" -Level "INFO"
            try {
                $response = Invoke-WebRequest -Uri $cleanUrl -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop
                Write-Log "HTTP Status: $($response.StatusCode), Content-Length: $($response.Content.Length)" -Level "INFO"
                $contentString = $response.Content.Trim()
                $contentData = $contentString | ConvertFrom-Json
                Write-Log "Fetched content data from URL: $cleanUrl" -Level "INFO"
            }
            catch {
                if ($_.Exception.Response) {
                    $statusCode = $_.Exception.Response.StatusCode.Value__
                    Write-Log "HTTP Error: StatusCode=$statusCode, Message=$($_.Exception.Message)" -Level "ERROR"
                    if ($statusCode -eq 404 -and $url -ne $cleanUrl) {
                        Write-Log "Retrying with original URL due to 404: $url" -Level "INFO"
                        $response = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop
                        Write-Log "HTTP Status: $($response.StatusCode), Content-Length: $($response.Content.Length)" -Level "INFO"
                        $contentString = $response.Content.Trim()
                        $contentData = $contentString | ConvertFrom-Json
                        Write-Log "Fetched content data from original URL: $url" -Level "INFO"
                    }
                    else {
                        throw $_
                    }
                }
                else {
                    throw $_
                }
            }
        }
        elseif ($cleanUrl -match "^\\\\") {
            Write-Log "Detected network path" -Level "INFO"
            if (-not (Test-Path $cleanUrl)) { throw "Network path not accessible: $cleanUrl" }
            $rawContent = Get-Content -Path $cleanUrl -Raw
            $contentData = $rawContent | ConvertFrom-Json
            Write-Log "Fetched content data from network path: $cleanUrl" -Level "INFO"
        }
        else {
            Write-Log "Assuming local file path" -Level "INFO"
            $fullPath = if ([System.IO.Path]::IsPathRooted($cleanUrl)) { $cleanUrl } else { Join-Path $ScriptDir $cleanUrl }
            Write-Log "Resolved full path: $fullPath" -Level "INFO"
            if (-not (Test-Path $fullPath)) { throw "Local path not found: $fullPath" }
            $rawContent = Get-Content -Path $fullPath -Raw
            $contentData = $rawContent | ConvertFrom-Json
            Write-Log "Fetched content data from local path: $fullPath" -Level "INFO"
        }
        $global:CachedContentData = $contentData
        $global:LastContentFetch = Get-Date
        Write-Log "Content fetched from remote source: $cleanUrl" -Level "INFO"
        return [PSCustomObject]@{
            Data   = $contentData
            Source = "Remote"
        }
    }
    catch {
        Write-Log "Failed to fetch content data from ${cleanUrl}: $_" -Level "ERROR"
        Write-Log "ContentDataUrl invalid or unreachable, using default content" -Level "ERROR"
        return [PSCustomObject]@{
            Data   = $defaultContentData
            Source = "Default"
        }
    }
}

# ------------------------------------------------------------
# Certificate Check Functions
# ------------------------------------------------------------
function Get-YubiKeyCertExpiryDays {
    try {
        if (-not (Test-Path "C:\Program Files\Yubico\Yubikey Manager\ykman.exe")) {
            throw "ykman.exe not found at C:\Program Files\Yubico\Yubikey Manager\ykman.exe"
        }
        Write-Log "ykman.exe found at C:\Program Files\Yubico\Yubikey Manager\ykman.exe" -Level "INFO"
        $yubiKeyInfo = & "C:\Program Files\Yubico\Yubikey Manager\ykman.exe" info 2>$null
        if (-not $yubiKeyInfo) {
            Write-Log "No YubiKey detected" -Level "INFO"
            return "YubiKey not present"
        }
        Write-Log "YubiKey detected: $yubiKeyInfo" -Level "INFO"
        $pivInfo = & "C:\Program Files\Yubico\Yubikey Manager\ykman.exe" "piv" "info" 2>$null
        if ($pivInfo) {
            Write-Log "PIV info: $pivInfo" -Level "INFO"
        }
        else {
            Write-Log "No PIV info available" -Level "WARNING"
        }
        $slots = @("9a", "9c", "9d", "9e")
        $certPem = $null
        $slotUsed = $null
        foreach ($slot in $slots) {
            Write-Log "Checking slot $slot for certificate" -Level "INFO"
            $certPem = & "C:\Program Files\Yubico\Yubikey Manager\ykman.exe" "piv" "certificates" "export" $slot "-" 2>$null
            if ($certPem -and $certPem -match "-----BEGIN CERTIFICATE-----") {
                $slotUsed = $slot
                Write-Log "Certificate found in slot $slot" -Level "INFO"
                break
            }
            else {
                Write-Log "No valid certificate in slot $slot" -Level "INFO"
            }
        }
        if (-not $certPem) { throw "No certificate found in slots 9a, 9c, 9d, or 9e" }
        $tempFile = [System.IO.Path]::GetTempFileName()
        $certPem | Out-File $tempFile -Encoding ASCII
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $cert.Import($tempFile)
        $expiryDateFormatted = $cert.NotAfter.ToString("MM/dd/yyyy")
        Remove-Item $tempFile -Force
        return "YubiKey Certificate (Slot $slotUsed): Expires: $expiryDateFormatted"
    }
    catch {
        if ($_.Exception.Message -ne "No YubiKey detected by ykman") {
            Write-Log "Error retrieving YubiKey certificate expiry: $_" -Level "ERROR"
            return "YubiKey Certificate: Unable to determine expiry date - $_"
        }
    }
}

function Get-VirtualSmartCardCertExpiry {
    try {
        $criteria = "Virtual|VSC|TPM|Identity Device"
        $virtualCerts = @()
        foreach ($store in @("Cert:\CurrentUser\My", "Cert:\LocalMachine\My")) {
            $virtualCerts += Get-ChildItem $store | Where-Object {
                ($_.Subject -match $criteria) -or (($_.FriendlyName -ne $null) -and ($_.FriendlyName -match $criteria))
            }
        }
        if (-not $virtualCerts -or $virtualCerts.Count -eq 0) {
            Write-Log "No Microsoft Virtual Smart Card certificate found using criteria '$criteria'" -Level "INFO"
            return "Microsoft Virtual Smart Card not present"
        }
        $cert = $virtualCerts | Sort-Object NotAfter | Select-Object -Last 1
        $expiryDateFormatted = $cert.NotAfter.ToString("MM/dd/yyyy")
        return "Microsoft Virtual Smart Card Certificate: Expires: $expiryDateFormatted"
    }
    catch {
        Write-Log "Error retrieving Microsoft Virtual Smart Card certificate expiry: $_" -Level "ERROR"
        return "Microsoft Virtual Smart Card Certificate: Unable to determine expiry date"
    }
}

function Update-CertificateInfo {
    try {
        $ykStatus = Get-YubiKeyCertExpiryDays
        $vscStatus = Get-VirtualSmartCardCertExpiry
        $combinedStatus = "$ykStatus`n$vscStatus"
        $window.Dispatcher.Invoke({
            if ($global:YubiKeyComplianceText) { 
                $global:YubiKeyComplianceText.Text = $combinedStatus 
            }
        })
        Write-Log "Certificate info updated: $combinedStatus" -Level "INFO"
    }
    catch {
        Write-Log "Failed to update certificate info: $_" -Level "ERROR"
    }
}

# ------------------------------------------------------------
# Update Functions for Additional UI Sections
# ------------------------------------------------------------
function Update-Announcements {
    try {
        $contentResult = $global:contentData
        $current = $contentResult.Data.Announcements
        $source = $contentResult.Source
        $last = $config.AnnouncementsLastState

        if ($last -and $current.Text -ne $last.Text -and -not $global:AnnouncementsExpander.IsExpanded) {
            $window.Dispatcher.Invoke({
                if ($global:AnnouncementsAlertIcon) {
                    $global:AnnouncementsAlertIcon.Visibility = "Visible"
                }
            })
        }

        $window.Dispatcher.Invoke({
            $global:AnnouncementsText.Text = $current.Text
            $global:AnnouncementsDetailsText.Text = $current.Details
            $global:AnnouncementsLinksPanel.Children.Clear()
            foreach ($link in $current.Links) {
                $tb = New-Object System.Windows.Controls.TextBlock
                $hp = New-Object System.Windows.Documents.Hyperlink
                $hp.NavigateUri = [Uri]$link.Url
                $hp.Inlines.Add($link.Name)
                $hp.Add_RequestNavigate({ param($s,$e) Start-Process $e.Uri.AbsoluteUri; $e.Handled = $true })
                $tb.Inlines.Add($hp)
                $global:AnnouncementsLinksPanel.Children.Add($tb)
            }
            if ($global:AnnouncementsSourceText) {
                $global:AnnouncementsSourceText.Text = "Source: $source"
            }
        })

        $config.AnnouncementsLastState = $current
        Save-Configuration -Config $config

        Write-Log "Announcements updated from $source." -Level "INFO"
    }
    catch {
        Write-Log "Error updating Announcements: $_" -Level "ERROR"
    }
}

function Update-Support {
    try {
        $contentResult = $global:contentData
        $current = $contentResult.Data.Support
        $source = $contentResult.Source
        $last = $config.SupportLastState

        if ($last -and $current.Text -ne $last.Text -and -not $global:SupportExpander.IsExpanded) {
            $window.Dispatcher.Invoke({
                if ($global:SupportAlertIcon) {
                    $global:SupportAlertIcon.Visibility = "Visible"
                }
            })
        }

        $window.Dispatcher.Invoke({
            $global:SupportText.Text = $current.Text
            $global:SupportLinksPanel.Children.Clear()
            foreach ($link in $current.Links) {
                $tb = New-Object System.Windows.Controls.TextBlock
                $hp = New-Object System.Windows.Documents.Hyperlink
                $hp.NavigateUri = [Uri]$link.Url
                $hp.Inlines.Add($link.Name)
                $hp.Add_RequestNavigate({ param($s,$e) Start-Process $e.Uri.AbsoluteUri; $e.Handled = $true })
                $tb.Inlines.Add($hp)
                $global:SupportLinksPanel.Children.Add($tb)
            }
            if ($global:SupportSourceText) {
                $global:SupportSourceText.Text = "Source: $source"
            }
        })

        $config.SupportLastState = $current
        Save-Configuration -Config $config

        Write-Log "Support updated from $source." -Level "INFO"
    }
    catch {
        Write-Log "Error updating Support: $_" -Level "ERROR"
    }
}

function Update-PatchingUpdates {
    try {
        $patchFilePath = if ([System.IO.Path]::IsPathRooted($config.PatchInfoFilePath)) {
            $config.PatchInfoFilePath
        }
        else {
            Join-Path $ScriptDir $config.PatchInfoFilePath
        }
        Write-Log "Resolved patch file path: $patchFilePath" -Level "INFO"
        if (Test-Path $patchFilePath -PathType Leaf) {
            $patchContent = Get-Content -Path $patchFilePath -Raw -ErrorAction Stop
            $patchText = if ([string]::IsNullOrWhiteSpace($patchContent)) { 
                "Patch info file is empty." 
            } else { 
                $patchContent.Trim() 
            }
            Write-Log "Successfully read patch info: $patchText" -Level "INFO"
        }
        else {
            $patchText = "Patch info file not found at $patchFilePath."
            Write-Log "Patch info file not found: $patchFilePath" -Level "WARNING"
        }
        $window.Dispatcher.Invoke({ $global:PatchingUpdatesText.Text = $patchText })
        Write-Log "Patching status updated: $patchText" -Level "INFO"
    }
    catch {
        $errorMessage = "Error reading patch info file: $_"
        Write-Log $errorMessage -Level "ERROR"
        $window.Dispatcher.Invoke({ $global:PatchingUpdatesText.Text = $errorMessage })
    }
}

# ============================================================
# I) Tray Icon Management
# ============================================================
function Get-Icon {
    param(
        [string]$Path,
        [System.Drawing.Icon]$DefaultIcon
    )
    Write-Log "Attempting to load icon from: $Path" -Level "INFO"
    if (-not (Test-Path $Path)) {
        Write-Log "$Path not found. Using default icon." -Level "WARNING"
        return $DefaultIcon
    }
    try {
        $icon = New-Object System.Drawing.Icon($Path)
        Write-Log "Custom icon successfully loaded from $Path" -Level "INFO"
        return $icon
    }
    catch {
        Write-Log "Error loading icon from ${Path}: $_" -Level "ERROR"
        return $DefaultIcon
    }
}

function Update-TrayIcon {
    try {
        if (-not $global:FormsAvailable -or -not $global:TrayIcon -or $global:TrayIcon.IsDisposed) {
            Write-Log "Skipping tray icon update: Tray icon is unavailable or disposed." -Level "INFO"
            return
        }

        # Check if any section still has its red-dot visible
        $hasAlert = $false
        foreach ($icon in @(
            $global:AnnouncementsAlertIcon,
            $global:SupportAlertIcon
        )) {
            if ($icon -and $icon.Visibility -eq 'Visible') {
                $hasAlert = $true
                break
            }
        }

        # Choose healthy vs warning
        $iconPath = if ($hasAlert) {
            $config.IconPaths.Warning
        } else {
            $config.IconPaths.Main
        }

        $global:TrayIcon.Icon = Get-Icon -Path $iconPath -DefaultIcon ([System.Drawing.SystemIcons]::Application)
        Write-Log "Tray icon updated to $iconPath" -Level "INFO"
    }
    catch {
        Write-Log "Error updating tray icon: $_" -Level "ERROR"
    }
}

function Set-TrayIconAlwaysShow {
    param(
        [string]$IconName = "LLNOTIFY v$ScriptVersion"
    )
    try {
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer"
        $regName = "EnableAutoTray"
        $regPathNotify = "HKCU:\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\TrayNotify"
        $iconRegName = "LLNOTIFY_IconVisibility"

        # Check if auto-tray is disabled (0 means all icons are shown)
        if (Test-Path $regPath) {
            $enableAutoTray = Get-ItemProperty -Path $regPath -Name $regName -ErrorAction SilentlyContinue
            if ($enableAutoTray.EnableAutoTray -eq 0) {
                Write-Log "Auto-tray is disabled; all icons are shown." -Level "INFO"
                return
            }
        }

        # Set icon-specific visibility (1 = Always Show)
        if (-not (Test-Path $regPathNotify)) {
            New-Item -Path $regPathNotify -Force | Out-Null
        }
        $existingValue = Get-ItemProperty -Path $regPathNotify -Name $iconRegName -ErrorAction SilentlyContinue
        if ($existingValue.$iconRegName -ne 1) {
            Set-ItemProperty -Path $regPathNotify -Name $iconRegName -Value 1 -Type DWord -Force
            Write-Log "Set tray icon '$IconName' to Always Show in registry." -Level "INFO"
        }
        else {
            Write-Log "Tray icon '$IconName' is already set to Always Show." -Level "INFO"
        }
    }
    catch {
        Write-Log "Error setting tray icon to Always Show: $_" -Level "ERROR"
    }
}

# ============================================================
# M) Create & Configure Tray Icon with Collapsible Menu
# ============================================================
function Initialize-TrayIcon {
    try {
        if (-not $global:FormsAvailable) {
            Write-Log "Skipping tray icon initialization: System.Windows.Forms is unavailable." -Level "WARNING"
            return
        }

        $global:TrayIcon = New-Object System.Windows.Forms.NotifyIcon
        $iconPath = $config.IconPaths.Main
        Write-Log "Initializing tray icon with: $iconPath" -Level "INFO"
        $global:TrayIcon.Icon = Get-Icon -Path $iconPath -DefaultIcon ([System.Drawing.SystemIcons]::Application)
        $global:TrayIcon.Text = "LLNOTIFY v$ScriptVersion"
        $global:TrayIcon.Visible = $true
        Write-Log "Tray icon initialized with $iconPath" -Level "INFO"
        Write-Log "Note: To ensure the LLNOTIFY tray icon is always visible, right-click the taskbar, select 'Taskbar settings', scroll to 'Notification area', click 'Select which icons appear on the taskbar', and set 'LLNOTIFY' to 'On'." -Level "INFO"
        
        # Set tray icon to Always Show
        Set-TrayIconAlwaysShow -IconName "LLNOTIFY v$ScriptVersion"
    }
    catch {
        Write-Log "Error initializing tray icon: $_" -Level "ERROR"
        $global:TrayIcon = $null
        return
    }

    try {
        $ContextMenuStrip = New-Object System.Windows.Forms.ContextMenuStrip
        $MenuItemShow = New-Object System.Windows.Forms.ToolStripMenuItem("Show Dashboard")
        $MenuItemQuickActions = New-Object System.Windows.Forms.ToolStripMenuItem("Quick Actions")
        $MenuItemRefresh = New-Object System.Windows.Forms.ToolStripMenuItem("Refresh Now")
        $MenuItemExit = New-Object System.Windows.Forms.ToolStripMenuItem("Exit")
        $MenuItemQuickActions.DropDownItems.Add($MenuItemRefresh)
        $ContextMenuStrip.Items.Add($MenuItemShow) | Out-Null
        $ContextMenuStrip.Items.Add($MenuItemQuickActions) | Out-Null
        $ContextMenuStrip.Items.Add($MenuItemExit) | Out-Null
        $global:TrayIcon.ContextMenuStrip = $ContextMenuStrip

        $global:TrayIcon.add_MouseClick({
            param($sender, $e)
            try {
                Write-Log "Tray icon clicked: Button=$($e.Button)" -Level "INFO"
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
            & "Update-TrayIcon"
            & "Update-Announcements"
            & "Update-Support"
            & "Update-PatchingUpdates"
            Update-CertificateInfo
            Write-Log "Manual refresh triggered from tray menu" -Level "INFO"
        })
        $MenuItemExit.add_Click({
            try {
                Write-Log "Exit clicked by user." -Level "INFO"
                if ($global:DispatcherTimer) {
                    $global:DispatcherTimer.Stop()
                    Write-Log "DispatcherTimer stopped." -Level "INFO"
                }
                if ($global:yubiKeyJob) {
                    Stop-Job -Job $global:yubiKeyJob -ErrorAction SilentlyContinue
                    Remove-Job -Job $global:yubiKeyJob -Force -ErrorAction SilentlyContinue
                    Write-Log "YubiKey job stopped and removed." -Level "INFO"
                }
                if ($global:TrayIcon -and -not $global:TrayIcon.IsDisposed) {
                    $global:TrayIcon.Visible = $false
                    $global:TrayIcon.Dispose()
                    Write-Log "Tray icon disposed." -Level "INFO"
                }
                $global:TrayIcon = $null
                $window.Dispatcher.InvokeShutdown()
                Write-Log "Application exited via tray menu." -Level "INFO"
            }
            catch {
                Handle-Error "Error during application exit: $_" -Source "Exit"
            }
        })
        Write-Log "ContextMenuStrip initialized successfully" -Level "INFO"
    }
    catch {
        Write-Log "Error setting up ContextMenuStrip: $_" -Level "ERROR"
        if ($global:TrayIcon -and -not $global:TrayIcon.IsDisposed) {
            $global:TrayIcon.Visible = $false
            $global:TrayIcon.Dispose()
        }
        $global:TrayIcon = $null
    }
}

# Initialize tray icon if forms are available
Initialize-TrayIcon

# ============================================================
# K) Window Visibility Management
# ============================================================
function Set-WindowPosition {
    try {
        $window.Dispatcher.Invoke({
            $window.Width = 350
            $window.Height = 500
            $window.UpdateLayout()
            $primary = if ($global:FormsAvailable) {
                [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
            } else {
                # Fallback for when System.Windows.Forms is unavailable
                [PSCustomObject]@{
                    X = 0
                    Y = 0
                    Width = 1920  # Default to common resolution
                    Height = 1080
                }
            }
            Write-Log "Primary screen: X=$($primary.X), Y=$($primary.Y), Width=$($primary.Width), Height=$($primary.Height)" -Level "INFO"
            $left = $primary.X + ($primary.Width - $window.ActualWidth) / 2
            $top = $primary.Y + ($primary.Height - $window.ActualHeight) / 2
            $taskbarBuffer = 50
            $maxTop = $primary.Height - $window.ActualHeight - $taskbarBuffer
            $top = [Math]::Max($primary.Y, [Math]::Min($top, $maxTop))
            $left = [Math]::Max($primary.X, [Math]::Min($left, $primary.X + $primary.Width - $window.ActualWidth))
            $top = [Math]::Max($primary.Y, [Math]::Min($top, $primary.Y + $primary.Height - $window.ActualHeight))
            $window.Left = $left
            $window.Top = $top
            Write-Log "Window position set: Left=$left, Top=$top, Width=$($window.ActualWidth), Height=$($window.ActualHeight)" -Level "INFO"
        })
    }
    catch {
        Handle-Error "Error setting window position: $_" -Source "Set-WindowPosition"
    }
}

function Toggle-WindowVisibility {
    try {
        $window.Dispatcher.Invoke({
            Write-Log "Current visibility: IsVisible=$($window.IsVisible), Visibility=$($window.Visibility)" -Level "INFO"
            if ($window.IsVisible) {
                $window.Hide()
                Write-Log "Dashboard hidden via Toggle-WindowVisibility." -Level "INFO"
            }
            else {
                Set-WindowPosition
                $window.Show()
                $window.WindowState = 'Normal'
                $window.Activate()
                $window.Topmost = $true
                Start-Sleep -Milliseconds 500
                $window.Topmost = $false
                Write-Log "Dashboard shown via Toggle-WindowVisibility at Left=$($window.Left), Top=$($window.Top), Visibility=$($window.Visibility), State=$($window.WindowState)" -Level "INFO"
            }
        }, "Normal")
    }
    catch {
        Handle-Error "Error toggling window visibility: $_" -Source "Toggle-WindowVisibility"
    }
}

function Update-UIElements {
    & "Update-Announcements"
    & "Update-Support"
    & "Update-PatchingUpdates"
    Update-CertificateInfo
}

# ============================================================
# L) Button and Event Handlers
# ============================================================
if ($global:SupportExpander) {
    $global:SupportExpander.Add_Expanded({
        try {
            $window.Dispatcher.Invoke({
                if ($global:SupportAlertIcon) { $global:SupportAlertIcon.Visibility = "Hidden" }
            })
            Write-Log "Support expander expanded, alert cleared." -Level "INFO"
        }
        catch {
            Write-Log "Error in SupportExpander expanded event: $_" -Level "ERROR"
        }
    })
}

# ============================================================
# O) DispatcherTimer for Periodic Updates
# ============================================================
$global:DispatcherTimer = New-Object System.Windows.Threading.DispatcherTimer
$global:DispatcherTimer.Interval = [TimeSpan]::FromSeconds($config.RefreshInterval)
$global:DispatcherTimer.add_Tick({
    try {
        $global:contentData = Fetch-ContentData
        & "Update-TrayIcon"
        & "Update-Announcements"
        & "Update-Support"
        & "Update-PatchingUpdates"
        Update-CertificateInfo
        Write-Log "Dispatcher tick completed" -Level "INFO"
    }
    catch {
        Handle-Error "Error during timer tick: $_" -Source "DispatcherTimer"
    }
})
$global:DispatcherTimer.Start()
Write-Log "DispatcherTimer started with interval $($config.RefreshInterval) seconds" -Level "INFO"

# ============================================================
# P) Dispatcher Exception Handling
# ============================================================
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

# ============================================================
# Q) Initial Update & Start Dispatcher
# ============================================================
try {
    $global:contentData = Fetch-ContentData
    Write-Log "Initial contentData set from $($global:contentData.Source): $($global:contentData.Data | ConvertTo-Json -Depth 3)" -Level "INFO"
    try { & "Update-TrayIcon" } catch { Handle-Error "Update-TrayIcon failed: $_" -Source "InitialUpdate" }
    try { Update-Announcements } catch { Handle-Error "Update-Announcements failed: $_" -Source "InitialUpdate" }
    try { Update-Support } catch { Handle-Error "Update-Support failed: $_" -Source "InitialUpdate" }
    try { Update-PatchingUpdates } catch { Handle-Error "Update-PatchingUpdates failed: $_" -Source "InitialUpdate" }
    try { Update-CertificateInfo } catch { Handle-Error "Update-CertificateInfo failed: $_" -Source "InitialUpdate" }
    Log-DotNetVersion
    Write-Log "Initial update completed" -Level "INFO"
}
catch {
    Handle-Error "Error during initial update setup: $_" -Source "InitialUpdate"
}

Write-Log "About to call Dispatcher.Run()..." -Level "INFO"
[System.Windows.Threading.Dispatcher]::Run()
Write-Log "Dispatcher ended; script exiting." -Level "INFO"
