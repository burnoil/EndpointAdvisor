# LLNOTIFY.ps1 - Lincoln Laboratory Notification System
# Version 4.3.89 (Stable Release with first-run notification)

# Ensure $PSScriptRoot is defined for older versions
if ($MyInvocation.MyCommand.Path) {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    $ScriptDir = Get-Location
}

# Define version
$ScriptVersion = "4.3.89"

# --- START OF SINGLE-INSTANCE CHECK ---
# Single-Instance Check: Prevents multiple copies of the application from running.
$AppName = "LLNOTIFY"
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

# Global flag to track pending restart state
$global:PendingRestart = $false

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

Write-Log "--- LLNOTIFY Script Started (Version $ScriptVersion) ---"

# ============================================================
# MODULE: Configuration Management
# ============================================================
function Get-DefaultConfig {
    return @{
        RefreshInterval       = 900
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
        BigFixSSA_Path        = "C:\Program Files (x86)\BigFix Enterprise\BigFix Self Service Application\BigFixSSA.exe"
        YubiKeyManager_Path   = "C:\Program Files\Yubico\Yubikey Manager\ykman.exe"
        BlinkingEnabled       = $false
        CachePath             = Join-Path $ScriptDir "ContentData.cache.json"
        HasRunBefore          = $false
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
        [string]$Path = (Join-Path $ScriptDir "LLNOTIFY.config.json")
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
$LogFilePath = Join-Path $ScriptDir "LLNOTIFY.log"
$config = Load-Configuration

$mainIconPath = $config.IconPaths.Main
$warningIconPath = $config.IconPaths.Warning

Write-Log "Main icon path: $mainIconPath" -Level "INFO"
Write-Log "Warning icon path: $warningIconPath" -Level "INFO"

$defaultContentData = @{
    Announcements = @{ Text = "No announcements at this time."; Details = ""; Links = @() }
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
    Title="LLNOTIFY - Lincoln Laboratory Notification System"
    WindowStartupLocation="Manual" 
    SizeToContent="Manual"
    MinWidth="350" MinHeight="500"
    MaxWidth="400" MaxHeight="550"
    ResizeMode="CanResizeWithGrip" ShowInTaskbar="False" Visibility="Hidden" Topmost="True"
    Background="#f0f0f0">
  <Window.Resources>
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
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <Border Grid.Row="0" Background="#0078D7" Padding="5" CornerRadius="3" Margin="0,0,0,5">
      <StackPanel Orientation="Horizontal" VerticalAlignment="Center" HorizontalAlignment="Center">
        <Image x:Name="HeaderIcon" Width="20" Height="20" Margin="0,0,5,0"/>
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
        <Expander x:Name="PatchingExpander" FontSize="12" IsExpanded="False" Margin="0,2,0,2">
          <Expander.Header>
            <TextBlock Text="Patching and Updates" VerticalAlignment="Center"/>
          </Expander.Header>
          <Border BorderBrush="#00008B" BorderThickness="1" Padding="5" CornerRadius="3" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="PatchingDescriptionText" FontSize="11" TextWrapping="Wrap" Visibility="Collapsed"/>
              
              <StackPanel x:Name="PendingRestartPanel" Orientation="Vertical" Visibility="Collapsed">
                <TextBlock Text="Pending Restart Status:" FontSize="11" FontWeight="Bold" Margin="0,0,0,0"/>
                <TextBlock x:Name="PendingRestartStatusText" FontSize="11" FontWeight="Bold" TextWrapping="Wrap"/>
              </StackPanel>
              
              <TextBlock Text="Available Updates:" FontSize="11" FontWeight="Bold" Margin="0,10,0,2"/>
              
              <Grid Margin="0,2,0,2">
                  <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock x:Name="BigFixStatusText" Grid.Column="0" VerticalAlignment="Center" FontSize="11" TextWrapping="Wrap"/>
                  <Button x:Name="BigFixLaunchButton" Grid.Column="1" Content="App Updates" Margin="10,0,0,0" Padding="5,1" VerticalAlignment="Center" Visibility="Collapsed" ToolTip="Install available application updates"/>
              </Grid>

              <Grid Margin="0,2,0,2">
                  <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock x:Name="ECMStatusText" Grid.Column="0" VerticalAlignment="Center" FontSize="11" TextWrapping="Wrap"/>
                  <Button x:Name="ECMLaunchButton" Grid.Column="1" Content="Install Patches" Margin="10,0,0,0" Padding="5,1" VerticalAlignment="Center" Visibility="Collapsed" ToolTip="Install pending Windows OS patches"/>
              </Grid>
            </StackPanel>
          </Border>
        </Expander>
        <Expander x:Name="SupportExpander" FontSize="12" IsExpanded="True" Margin="0,2,0,2">
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
        <Expander x:Name="ComplianceExpander" FontSize="12" IsExpanded="False" Margin="0,2,0,2">
          <Expander.Header>
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="Certificate Status" VerticalAlignment="Center"/>
            </StackPanel>
          </Expander.Header>
          <Border BorderBrush="#00008B" BorderThickness="1" Padding="5" CornerRadius="3" Background="White" Margin="2">
            <TextBlock x:Name="YubiKeyComplianceText" FontSize="11" TextWrapping="Wrap"/>
          </Border>
        </Expander>
        <TextBlock x:Name="WindowsBuildText" FontSize="11" TextWrapping="Wrap" HorizontalAlignment="Center" Margin="0,10,0,0"/>
      </StackPanel>
    </ScrollViewer>
    <Grid Grid.Row="2" Margin="0,5,0,0">
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

    $window.Width = 350
    $window.Height = 500

    # Modified UI Elements list
    $uiElements = @(
        "HeaderIcon", "AnnouncementsExpander", "AnnouncementsAlertIcon", "AnnouncementsText", "AnnouncementsDetailsText",
        "AnnouncementsLinksPanel", "AnnouncementsSourceText", "PatchingExpander", "PatchingDescriptionText",
        "PendingRestartPanel", "PendingRestartStatusText", "SupportExpander", "SupportAlertIcon", "SupportText", "SupportLinksPanel",
        "SupportSourceText", "ComplianceExpander", "YubiKeyComplianceText", "WindowsBuildText", "ClearAlertsButton",
        "FooterText", "ClearAlertsPanel", "ClearAlertsDot", "BigFixStatusText", "BigFixLaunchButton", "ECMStatusText", "ECMLaunchButton"
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

    # Initialize events after a delay
    function InitializeUI {
        if ($global:AnnouncementsExpander) {
            $global:AnnouncementsExpander.IsExpanded = $true
            $global:AnnouncementsExpander.Add_Expanded({ 
                if ($global:AnnouncementsAlertIcon) { $global:AnnouncementsAlertIcon.Visibility = "Hidden" }
                Update-TrayIcon
            })
        }
        if ($global:SupportExpander) {
            $global:SupportExpander.IsExpanded = $true
            $global:SupportExpander.Add_Expanded({ 
                if ($global:SupportAlertIcon) { $global:SupportAlertIcon.Visibility = "Hidden" }
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

        if ($global:ClearAlertsButton) {
            $global:ClearAlertsButton.Add_Click({
                Write-Log "Clear Alerts button clicked by user to clear new alerts (red dots)." -Level "INFO"
                if ($global:contentData) {
                    $config.AnnouncementsLastState = $global:contentData.Data.Announcements | ConvertTo-Json -Compress
                    $config.SupportLastState = $global:contentData.Data.Support | ConvertTo-Json -Compress
                }
                $window.Dispatcher.Invoke({
                    if ($global:AnnouncementsAlertIcon) { $global:AnnouncementsAlertIcon.Visibility = "Hidden" }
                    if ($global:SupportAlertIcon) { $global:SupportAlertIcon.Visibility = "Hidden" }
                    if ($global:ClearAlertsDot) { $global:ClearAlertsDot.Visibility = "Hidden" }
                })
                $global:BlinkingTimer.Stop()
                Update-TrayIcon
                Save-Configuration -Config $config
            })
        }
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

function Get-WindowsBuildNumber {
    try {
        $buildInfo = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
        $productName = if ($buildInfo.CurrentBuildNumber -ge 22000) { "Windows 11" } else { "Windows 10" }
        return "$productName Build: $($buildInfo.DisplayVersion)"
    } catch { return "Windows Build: Unknown" }
}

function Get-ECMUpdateStatus {
    try {
        # Get all pending updates
        $pendingUpdates = Get-WmiObject -Namespace 'ROOT\ccm\ClientSDK' -Class CCM_SoftwareUpdate | Where-Object { $_.ComplianceState -eq 0 }

        if ($null -eq $pendingUpdates) {
            return [PSCustomObject]@{
                StatusText        = "Windows OS Patches and Updates: No Updates Pending."
                HasPendingUpdates = $false
            }
        }

        # --- START OF NEW, MORE ROBUST LOGIC ---
        
        # Get all update *assignments*. These contain the deployment settings like reboot behavior.
        $assignments = Get-WmiObject -Namespace 'ROOT\ccm\ClientSDK' -Class CCM_UpdateCIAssignment
        
        # For faster lookups, convert the assignments into a hashtable with the UpdateID as the key.
        $assignmentLookup = @{}
        foreach ($assignment in $assignments) {
            # Some updates can have multiple assignments; we only need one.
            if (-not $assignmentLookup.ContainsKey($assignment.UpdateID)) {
                $assignmentLookup.Add($assignment.UpdateID, $assignment)
            }
        }

        $pendingCount = ($pendingUpdates | Measure-Object).Count
        $rebootMayBeNeeded = $false

        # Now, check the assignment for each pending update.
        foreach ($update in $pendingUpdates) {
            $assignment = $assignmentLookup[$update.UpdateID]
            
            # An update is considered to potentially require a reboot if the deployment settings indicate it.
            # NotifyUserOfReboot and RebootOutsideOfMaintWin are the key flags from the deployment.
            if ($null -ne $assignment -and ($assignment.NotifyUserOfReboot -eq $true -or $assignment.RebootOutsideOfMaintWin -eq $true)) {
                $rebootMayBeNeeded = $true
                break # We found one that might require a reboot, so we can stop looking.
            }
        }
        
        # Build the final status text based on our findings.
        $statusMessage = "Windows OS Patches and Updates: $pendingCount update(s) pending"
        if ($rebootMayBeNeeded) {
            $statusMessage += " (restart may be required)."
        } else {
            $statusMessage += "."
        }
        # --- END OF NEW LOGIC ---

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

        # --- START OF MODIFICATION: Dynamically set the tooltip for the ECM button ---
        if ($showEcmButton) {
            $global:ECMLaunchButton.ToolTip = "Install pending Windows OS patches (a restart may be required)."
        }
        # --- END OF MODIFICATION ---
    })
}

function Convert-MarkdownToTextBlock {
    param(
        [string]$Text,
        [System.Windows.Controls.TextBlock]$TargetTextBlock
    )
    
    try {
        if (-not $Text -or -not $TargetTextBlock) {
            Write-Log "Markdown text or TargetTextBlock is null or empty" -Level "WARNING"
            $TargetTextBlock.Inlines.Clear()
            if ($Text) {
                $TargetTextBlock.Inlines.Add((New-Object System.Windows.Documents.Run($Text)))
            }
            return
        }

        $TargetTextBlock.Inlines.Clear()
        
        $regexColor = "\[(green|red|yellow)\](.*?)\[/\1\]"
        $regexBold = "\*\*(.*?)\*\*"
        $regexItalic = "\*(.*?)\*"
        $regexUnderline = "__(.*?)__"
        
        # Step 1: Process color tags first, replacing with unique placeholders
        $currentText = $Text
        $colorPlaceholders = @{}
        $placeholderCounter = 0
        $colorMatches = [regex]::Matches($Text, $regexColor) | Sort-Object Index -Descending
        
        foreach ($match in $colorMatches) {
            $placeholder = "{COLORPH$placeholderCounter}"
            $isBold = $Text.Substring($match.Index - 2, 2) -eq "**" -and $Text.Substring($match.Index + $match.Length, 2) -eq "**"
            $colorPlaceholders[$placeholder] = @{
                Text = $match.Groups[2].Value
                Color = $match.Groups[1].Value
                IsBold = $isBold
            }
            $currentText = $currentText.Remove($match.Index, $match.Length).Insert($match.Index, $placeholder)
            Write-Log "Replaced color tag with placeholder: $placeholder (Text: $($match.Groups[2].Value), Color: $($match.Groups[1].Value), IsBold: $isBold)" -Level "INFO"
            $placeholderCounter++
        }
        
        # Step 2: Process bold, italic, and underline on the modified text
        $matches = @()
        $boldMatches = [regex]::Matches($currentText, $regexBold)
        $italicMatches = [regex]::Matches($currentText, $regexItalic)
        $underlineMatches = [regex]::Matches($currentText, $regexUnderline)
        
        foreach ($match in $boldMatches) {
            $matches += [PSCustomObject]@{
                Index = $match.Index
                Length = $match.Length
                Text = $match.Groups[1].Value
                Type = "Bold"
                FullMatch = $match.Value
            }
        }
        foreach ($match in $italicMatches) {
            if ([string]::IsNullOrWhiteSpace($match.Groups[1].Value)) { continue } # Skip empty italic matches
            $matches += [PSCustomObject]@{
                Index = $match.Index
                Length = $match.Length
                Text = $match.Groups[1].Value
                Type = "Italic"
                FullMatch = $match.Value
            }
        }
        foreach ($match in $underlineMatches) {
            $matches += [PSCustomObject]@{
                Index = $match.Index
                Length = $match.Length
                Text = $match.Groups[1].Value
                Type = "Underline"
                FullMatch = $match.Value
            }
        }
        
        $matches = $matches | Sort-Object Index
        $lastIndex = 0
        $runs = @()

        foreach ($match in $matches) {
            if ($match.Index -gt $lastIndex) {
                $plainText = $currentText.Substring($lastIndex, $match.Index - $lastIndex)
                $runs += Process-TextSegment -Text $plainText -ColorPlaceholders $colorPlaceholders
                Write-Log "Added plain text segment: $plainText" -Level "INFO"
            }
            
            $text = $match.Text
            if ($colorPlaceholders.ContainsKey($text)) {
                $colorInfo = $colorPlaceholders[$text]
                $innerRuns = Process-InnerMarkdown -Text $colorInfo.Text -Color $colorInfo.Color -IsBold $colorInfo.IsBold
                $runs += $innerRuns
                Write-Log "Processed color placeholder: $text (Text: $($colorInfo.Text), Color: $($colorInfo.Color), IsBold: $($colorInfo.IsBold))" -Level "INFO"
            } else {
                $run = New-Object System.Windows.Documents.Run($text)
                if ($match.Type -eq "Bold") {
                    $run.FontWeight = [System.Windows.FontWeights]::Bold
                } elseif ($match.Type -eq "Italic") {
                    $run.FontStyle = [System.Windows.FontStyles]::Italic
                } elseif ($match.Type -eq "Underline") {
                    $run.TextDecorations = [System.Windows.TextDecorations]::Underline
                }
                $runs += $run
                Write-Log "Added $($match.Type) run: $text" -Level "INFO"
            }
            
            $lastIndex = $match.Index + $match.Length
        }
        
        if ($lastIndex -lt $currentText.Length) {
            $plainText = $currentText.Substring($lastIndex)
            $runs += Process-TextSegment -Text $plainText -ColorPlaceholders $colorPlaceholders
            Write-Log "Added final plain text segment: $plainText" -Level "INFO"
        }
        
        foreach ($run in $runs) {
            $TargetTextBlock.Inlines.Add($run)
            Write-Log "Added run to TextBlock: $($run.Text) (FontWeight: $($run.FontWeight), FontStyle: $($run.FontStyle), Foreground: $($run.Foreground))" -Level "INFO"
        }
        
        Write-Log "Successfully parsed Markdown for text: $Text" -Level "INFO"
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
            Write-Log "Added inner plain text run: $plainText" -Level "INFO"
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
        Write-Log "Added inner final plain text run: $plainText" -Level "INFO"
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
        if (-not $Text) {
            Write-Log "Inner Markdown text is empty" -Level "WARNING"
            return @()
        }

        $runs = @()
        $regexBold = "\*\*(.*?)\*\*"
        $regexItalic = "\*(.*?)\*"
        $regexUnderline = "__(.*?)__"
        
        $matches = @()
        $boldMatches = [regex]::Matches($Text, $regexBold)
        $italicMatches = [regex]::Matches($Text, $regexItalic)
        $underlineMatches = [regex]::Matches($Text, $regexUnderline)
        
        foreach ($match in $boldMatches) {
            $matches += [PSCustomObject]@{ Index = $match.Index; Length = $match.Length; Text = $match.Groups[1].Value; Type = "Bold"; FullMatch = $match.Value }
        }
        foreach ($match in $italicMatches) {
            if ([string]::IsNullOrWhiteSpace($match.Groups[1].Value)) { continue } # Skip empty italic matches
            $matches += [PSCustomObject]@{ Index = $match.Index; Length = $match.Length; Text = $match.Groups[1].Value; Type = "Italic"; FullMatch = $match.Value }
        }
        foreach ($match in $underlineMatches) {
            $matches += [PSCustomObject]@{ Index = $match.Index; Length = $match.Length; Text = $match.Groups[1].Value; Type = "Underline"; FullMatch = $match.Value }
        }
        
        $matches = $matches | Sort-Object Index
        $lastIndex = 0

        foreach ($match in $matches) {
            if ($match.Index -gt $lastIndex) {
                $plainText = $Text.Substring($lastIndex, $match.Index - $lastIndex)
                $run = New-Object System.Windows.Documents.Run($plainText)
                if ($Color) {
                    $colorBrush = [System.Windows.Media.Brushes]::($Color.Substring(0,1).ToUpper() + $Color.Substring(1))
                    $run.Foreground = $colorBrush
                }
                if ($IsBold) {
                    $run.FontWeight = [System.Windows.FontWeights]::Bold
                }
                $runs += $run
                Write-Log "Added inner plain text run: $plainText (Color: $Color, IsBold: $IsBold)" -Level "INFO"
            }
            
            $run = New-Object System.Windows.Documents.Run($match.Text)
            if ($match.Type -eq "Bold") {
                $run.FontWeight = [System.Windows.FontWeights]::Bold
            } elseif ($match.Type -eq "Italic") {
                $run.FontStyle = [System.Windows.FontStyles]::Italic
            } elseif ($match.Type -eq "Underline") {
                $run.TextDecorations = [System.Windows.TextDecorations]::Underline
            }
            if ($Color) {
                $colorBrush = [System.Windows.Media.Brushes]::($Color.Substring(0,1).ToUpper() + $Color.Substring(1))
                $run.Foreground = $colorBrush
            }
            if ($IsBold) {
                $run.FontWeight = [System.Windows.FontWeights]::Bold
            }
            $runs += $run
            Write-Log "Added inner $($match.Type) run: $($match.Text) (Color: $Color, IsBold: $IsBold)" -Level "INFO"
            
            $lastIndex = $match.Index + $match.Length
        }
        
        if ($lastIndex -lt $Text.Length) {
            $plainText = $Text.Substring($lastIndex)
            $run = New-Object System.Windows.Documents.Run($plainText)
            if ($Color) {
                $colorBrush = [System.Windows.Media.Brushes]::($Color.Substring(0,1).ToUpper() + $Color.Substring(1))
                $run.Foreground = $colorBrush
            }
            if ($IsBold) {
                $run.FontWeight = [System.Windows.FontWeights]::Bold
            }
            $runs += $run
            Write-Log "Added inner final plain text run: $plainText (Color: $Color, IsBold: $IsBold)" -Level "INFO"
        }
        
        return $runs
    } catch {
        Write-Log "Failed to process inner Markdown for text: $Text - $($_.Exception.Message)" -Level "ERROR"
        $run = New-Object System.Windows.Documents.Run($Text)
        if ($Color) {
            $colorBrush = [System.Windows.Media.Brushes]::($Color.Substring(0,1).ToUpper() + $Color.Substring(1))
            $run.Foreground = $colorBrush
        }
        if ($IsBold) {
            $run.FontWeight = [System.Windows.FontWeights]::Bold
        }
        return @($run)
    }
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
            if ($global:AnnouncementsAlertIcon) { $global:AnnouncementsAlertIcon.Visibility = "Visible" }
            if ($global:ClearAlertsDot) { $global:ClearAlertsDot.Visibility = "Visible" }
        }
        if ($global:AnnouncementsText) {
            Convert-MarkdownToTextBlock -Text $newAnnouncementsObject.Text -TargetTextBlock $global:AnnouncementsText
        }
        if ($global:AnnouncementsDetailsText) {
            Convert-MarkdownToTextBlock -Text $newAnnouncementsObject.Details -TargetTextBlock $global:AnnouncementsDetailsText
        }
        if ($global:AnnouncementsLinksPanel) {
            $global:AnnouncementsLinksPanel.Children.Clear()
            if ($newAnnouncementsObject.Links) {
                foreach ($link in $newAnnouncementsObject.Links) {
                    $global:AnnouncementsLinksPanel.Children.Add((New-HyperlinkBlock -Name $link.Name -Url $link.Url))
                }
            }
        }
        if ($global:AnnouncementsSourceText) {
            $global:AnnouncementsSourceText.Text = "Source: $($global:contentData.Source)"
        }
    })
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
            if ($global:SupportAlertIcon) { $global:SupportAlertIcon.Visibility = "Visible" }
            if ($global:ClearAlertsDot) { $global:ClearAlertsDot.Visibility = "Visible" }
        }
        if ($global:SupportText) {
            Convert-MarkdownToTextBlock -Text $newSupportObject.Text -TargetTextBlock $global:SupportText
        }
        if ($global:SupportLinksPanel) {
            $global:SupportLinksPanel.Children.Clear()
            if ($newSupportObject.Links) {
                foreach ($link in $newSupportObject.Links) {
                    $global:SupportLinksPanel.Children.Add((New-HyperlinkBlock -Name $link.Name -Url $link.Url))
                }
            }
        }
        if ($global:SupportSourceText) {
            $global:SupportSourceText.Text = "Source: $($global:contentData.Source)"
        }
    })
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
            Write-Log "Error loading icon from `"$Path`" - $($_.Exception.Message)" -Level "ERROR"
        }
    }
    return [System.Drawing.SystemIcons]::Application
}

function Update-TrayIcon {
    if (-not $global:TrayIcon.Visible) { return }
    
    $announcementAlert = $global:AnnouncementsAlertIcon -and $global:AnnouncementsAlertIcon.Visibility -eq "Visible"
    $supportAlert = $global:SupportAlertIcon -and $global:SupportAlertIcon.Visibility -eq "Visible"
    
    $hasAnyAlert = $announcementAlert -or $supportAlert -or $global:PendingRestart

    $global:TrayIcon.Icon = if ($hasAnyAlert) { $global:WarningIcon } else { $global:MainIcon }
    $global:TrayIcon.Text = if ($hasAnyAlert) { "LLNOTIFY v$ScriptVersion - Alerts Pending" } else { "Lincoln Laboratory LLNOTIFY v$ScriptVersion" }
}

function Initialize-TrayIcon {
    if (-not $global:FormsAvailable) { return }
    try {
        $global:MainIcon = Get-Icon -Path $config.IconPaths.Main
        $global:WarningIcon = Get-Icon -Path $config.IconPaths.Warning

        $global:TrayIcon = New-Object System.Windows.Forms.NotifyIcon
        $global:TrayIcon.Icon = $global:MainIcon
        $global:TrayIcon.Text = "Lincoln Laboratory LLNOTIFY v$ScriptVersion"
        $global:TrayIcon.Visible = $true

        $ContextMenuStrip = New-Object System.Windows.Forms.ContextMenuStrip
        
        $intervalSubMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Set Update Interval")
        $tenMin = New-Object System.Windows.Forms.ToolStripMenuItem("10 minutes", $null, { 
            $config.RefreshInterval = 600
            Save-Configuration -Config $config
            $global:DispatcherTimer.Interval = [TimeSpan]::FromSeconds(600)
            $global:DispatcherTimer.Stop()
            $global:DispatcherTimer.Start()
            Write-Log "Update interval set to 10 minutes" -Level "INFO"
        })
        $fifteenMin = New-Object System.Windows.Forms.ToolStripMenuItem("15 minutes", $null, { 
            $config.RefreshInterval = 900
            Save-Configuration -Config $config
            $global:DispatcherTimer.Interval = [TimeSpan]::FromSeconds(900)
            $global:DispatcherTimer.Stop()
            $global:DispatcherTimer.Start()
            Write-Log "Update interval set to 15 minutes" -Level "INFO"
        })
        $twentyMin = New-Object System.Windows.Forms.ToolStripMenuItem("20 minutes", $null, { 
            $config.RefreshInterval = 1200
            Save-Configuration -Config $config
            $global:DispatcherTimer.Interval = [TimeSpan]::fromSeconds(1200)
            $global:DispatcherTimer.Stop()
            $global:DispatcherTimer.Start()
            Write-Log "Update interval set to 20 minutes" -Level "INFO"
        })
        $intervalSubMenu.DropDownItems.AddRange(@($tenMin, $fifteenMin, $twentyMin))

        $ContextMenuStrip.Items.AddRange(@(
            (New-Object System.Windows.Forms.ToolStripMenuItem("Show Dashboard", $null, { Toggle-WindowVisibility })),
            (New-Object System.Windows.Forms.ToolStripMenuItem("Refresh Now", $null, { Main-UpdateCycle -ForceCertificateCheck $true })),
            $intervalSubMenu,
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
        Write-Log "Main content update cycle running..." -Level "INFO"
        $global:contentData = Fetch-ContentData
        
        Update-Announcements
        Update-Support
        Update-PatchingAndSystem
        
        if ($ForceCertificateCheck -or (-not $global:LastCertificateCheck -or ((Get-Date) - $global:LastCertificateCheck).TotalSeconds -ge $config.CertificateCheckInterval)) {
            Update-CertificateInfo
            $global:LastCertificateCheck = Get-Date
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
    
    # --- FIRST RUN NOTIFICATION LOGIC ---
    if (-not $config.HasRunBefore) {
        Write-Log "First run detected. Showing balloon tip notification." -Level "INFO"
        
        $timeout = 10000 # 10 seconds in milliseconds
        $title = "LLNOTIFY is Running"
        $text = "The notification icon is active in your system tray. You may need to drag it from the overflow area (^) to the main taskbar."
        $icon = [System.Windows.Forms.ToolTipIcon]::Info

        $global:TrayIcon.ShowBalloonTip($timeout, $title, $text, $icon)

        # Set the flag to true and save it so this doesn't run again
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
    Write-Log "--- LLNOTIFY Script Exiting ---"
    if ($global:DispatcherTimer) { $global:DispatcherTimer.Stop() }
    if ($global:TrayIcon) { $global:TrayIcon.Dispose() }
    if ($global:MainIcon) { $global:MainIcon.Dispose() }
    if ($global:WarningIcon) { $global:WarningIcon.Dispose() }
}
