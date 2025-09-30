# Lincoln Laboratory Endpoint Advisor
# Version 8.0.0 (Modern Inbox)

# Ensure $PSScriptRoot is defined for older versions
if ($MyInvocation.MyCommand.Path) {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    $ScriptDir = Get-Location
}

# Define version
$ScriptVersion = "8.0.0"

# --- START OF SINGLE-INSTANCE CHECK ---
$AppName = "Lincoln Laboratory Endpoint Advisor"
$mutexName = "Global\{c14e4b1a-8b6b-4c3e-b0d3-3b2a2e5a7d6e}"
$isFirstInstance = $false 
$global:mutex = New-Object System.Threading.Mutex($true, $mutexName, [ref]$isFirstInstance)
if (-not $isFirstInstance) {
    Write-Host "An instance of $AppName is already running. Exiting."
    exit
}
# --- END OF SINGLE-INSTANCE CHECK ---

# --- START OF CONFERENCE ROOM CHECK ---
try {
    if ((Get-ItemProperty -Path "HKLM:\SOFTWARE\MITLL" -Name "IsConferenceRoom" -ErrorAction SilentlyContinue) -or (Test-Path "C:\Windows\IsConferenceRoom.stub")) {
        Write-Host "[INFO] Notification aborted due to detection of Conference Room PC."
        exit
    }
} catch {}
# --- END OF CONFERENCE ROOM CHECK ---

# ============================================================
# A) Advanced Logging & Error Handling
# ============================================================
function Write-Log {
    param([string]$Message, [ValidateSet("INFO", "WARNING", "ERROR")][string]$Level = "INFO")
    $logPath = if ($LogFilePath) { $LogFilePath } else { Join-Path $ScriptDir "LLEndpointAdvisor.log" }
    $logEntry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    try { Add-Content -Path $logPath -Value $logEntry -Force -ErrorAction Stop } catch { Write-Host $logEntry "(Failed to write to log)" }
}

function Handle-Error { param([string]$Message, [string]$Source = "") Write-Log ("[$Source] " + $Message) -Level "ERROR" }

Write-Log "--- Lincoln Laboratory Endpoint Advisor Script Started (Version $ScriptVersion) ---"

# ============================================================
# B) Configuration Management
# ============================================================
function Get-DefaultConfig {
    return @{
        RefreshInterval       = 900
        ContentDataUrl        = "https://raw.githubusercontent.com/burnoil/EndpointAdvisor/main/ContentData.json"
        BigFixSSA_Path        = "C:\Program Files (x86)\BigFix Enterprise\BigFix Self Service Application\BigFixSSA.exe"
        CachePath             = Join-Path $ScriptDir "ContentData.cache.json"
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
                    if ($finalConfig.ContainsKey($key) -and $loadedConfig.$key -ne $null) { $finalConfig[$key] = $loadedConfig.$key }
                }
            }
        } catch { Write-Log "Failed to load config file. Using defaults." -Level "WARNING" }
    }
    try { $finalConfig | ConvertTo-Json -Depth 100 | Out-File $Path -Force } catch { Write-Log "Could not save configuration to $Path" -Level "ERROR" }
    return $finalConfig
}
function Save-Configuration { 
    param($Config, [string]$Path = (Join-Path $ScriptDir "LLEndpointAdvisor.config.json")) 
    try { $Config | ConvertTo-Json -Depth 100 | Out-File $Path -Force } 
    catch { Handle-Error $_.Exception.Message -Source "Save-Configuration" } 
}

# ============================================================
# C) External Configuration Setup
# ============================================================
$LogFilePath = Join-Path $ScriptDir "LLEndpointAdvisor.log"
$config = Load-Configuration

# ============================================================
# D) Import Required Assemblies
# ============================================================
try { 
    Add-Type -AssemblyName PresentationFramework, System.Windows.Forms, System.Drawing -ErrorAction Stop 
} catch { 
    Write-Log "Failed to load GUI assemblies." -Level "ERROR"; exit 
}

# ============================================================
# E) XAML Layout Definition
# ============================================================
$xamlString = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Lincoln Laboratory Endpoint Advisor"
    WindowStartupLocation="Manual"
    MinWidth="450" MinHeight="600" Width="450" Height="600"
    ResizeMode="CanResizeWithGrip" ShowInTaskbar="False" Visibility="Hidden" Topmost="True"
    TextOptions.TextRenderingMode="ClearType" TextOptions.TextFormattingMode="Display">

    <Window.Resources>
        <SolidColorBrush x:Key="WindowBackgroundColor" Color="#F2F2F2"/>
        <SolidColorBrush x:Key="ContentBackgroundColor" Color="#FFFFFF"/>
        <SolidColorBrush x:Key="PrimaryAccentColor" Color="#0078D4"/>
        <SolidColorBrush x:Key="PrimaryAccentHoverColor" Color="#005A9E"/>
        <SolidColorBrush x:Key="MutedTextColor" Color="#666666"/>
        <SolidColorBrush x:Key="BorderColor" Color="#E0E0E0"/>
        <SolidColorBrush x:Key="SelectionColor" Color="#E6E6E6"/>
        <SolidColorBrush x:Key="ScrollBarColor" Color="#CDCDCD"/>
        <SolidColorBrush x:Key="SeverityInfoColor" Color="#0078D4"/>
        <SolidColorBrush x:Key="SeverityWarningColor" Color="#F7A000"/>
        <SolidColorBrush x:Key="SeverityCriticalColor" Color="#D13438"/>

        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="{StaticResource PrimaryAccentColor}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="3">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="{StaticResource PrimaryAccentHoverColor}"/></Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="ListViewItem">
            <Setter Property="HorizontalContentAlignment" Value="Stretch"/><Setter Property="Padding" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListViewItem">
                        <Border x:Name="Bd" Background="Transparent" Padding="{TemplateBinding Padding}" BorderBrush="Transparent" BorderThickness="2,0,0,0">
                            <ContentPresenter/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="{StaticResource SelectionColor}"/></Trigger>
                            <Trigger Property="IsSelected" Value="True"><Setter TargetName="Bd" Property="Background" Value="{StaticResource SelectionColor}"/><Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource PrimaryAccentColor}"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <Style x:Key="ScrollBarThumbStyle" TargetType="{x:Type Thumb}">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type Thumb}"><Border CornerRadius="2" Background="{StaticResource ScrollBarColor}" BorderBrush="Transparent" BorderThickness="1" /></ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="{x:Type ScrollBar}">
            <Setter Property="Background" Value="Transparent"/><Setter Property="Width" Value="8"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type ScrollBar}">
                        <Track Grid.Row="1" IsDirectionReversed="true" x:Name="PART_Track"><Track.Thumb><Thumb Style="{StaticResource ScrollBarThumbStyle}" /></Track.Thumb></Track>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="Window">
            <Setter Property="Background" Value="{StaticResource WindowBackgroundColor}"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
        </Style>
    </Window.Resources>

    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="{StaticResource PrimaryAccentColor}" Padding="10" CornerRadius="3" Margin="0,0,0,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <Image x:Name="HeaderIcon" Width="24" Height="24" Margin="0,0,10,0"/>
                    <TextBlock Text="Lincoln Laboratory Endpoint Advisor" FontSize="16" FontWeight="SemiBold" Foreground="White" VerticalAlignment="Center"/>
                    <Rectangle Width="34" Fill="Transparent"/>
                </StackPanel>
            </Grid>
        </Border>

        <Grid Grid.Row="1">
            <Grid.RowDefinitions>
                <RowDefinition Height="*" MinHeight="150" />
                <RowDefinition Height="Auto" />
                <RowDefinition Height="2*" MinHeight="200" />
            </Grid.RowDefinitions>
            <ListView x:Name="MessagesListView" Grid.Row="0" BorderThickness="1" BorderBrush="{StaticResource BorderColor}">
                <ListView.ItemTemplate>
                    <DataTemplate>
                        <Border Padding="10" BorderThickness="0,0,0,1" BorderBrush="{StaticResource BorderColor}">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Ellipse x:Name="SeverityIndicator" Width="10" Height="10" Margin="0,5,10,0" VerticalAlignment="Top"/>
                                <StackPanel Grid.Column="1">
                                    <TextBlock Text="{Binding title}" FontSize="14" TextWrapping="Wrap">
                                        <TextBlock.Style>
                                            <Style TargetType="TextBlock"><Setter Property="FontWeight" Value="Normal"/><Style.Triggers><DataTrigger Binding="{Binding IsRead}" Value="False"><Setter Property="FontWeight" Value="Bold"/></DataTrigger></Style.Triggers></Style>
                                        </TextBlock.Style>
                                    </TextBlock>
                                    <TextBlock Text="{Binding publishedDate}" FontSize="11" Foreground="{StaticResource MutedTextColor}" Margin="0,2,0,0"/>
                                </StackPanel>
                            </Grid>
                        </Border>
                        <DataTemplate.Triggers>
                            <DataTrigger Binding="{Binding severity}" Value="Info"><Setter TargetName="SeverityIndicator" Property="Fill" Value="{StaticResource SeverityInfoColor}"/></DataTrigger>
                            <DataTrigger Binding="{Binding severity}" Value="Warning"><Setter TargetName="SeverityIndicator" Property="Fill" Value="{StaticResource SeverityWarningColor}"/></DataTrigger>
                            <DataTrigger Binding="{Binding severity}" Value="Critical"><Setter TargetName="SeverityIndicator" Property="Fill" Value="{StaticResource SeverityCriticalColor}"/></DataTrigger>
                        </DataTemplate.Triggers>
                    </DataTemplate>
                </ListView.ItemTemplate>
            </ListView>
            <GridSplitter Grid.Row="1" Height="5" HorizontalAlignment="Stretch" Background="{StaticResource BorderColor}" ResizeBehavior="PreviousAndNext"/>
            <ScrollViewer Grid.Row="2" VerticalScrollBarVisibility="Auto" BorderThickness="1" BorderBrush="{StaticResource BorderColor}" Background="{StaticResource ContentBackgroundColor}" Padding="15">
    <StackPanel>
        <TextBlock x:Name="DetailTitleText" FontWeight="SemiBold" FontSize="18" TextWrapping="Wrap" Margin="0,0,0,5"/>
        
        <Border Height="1" Background="{StaticResource BorderColor}" Margin="0,5,0,10" CornerRadius="0.5"/>

        <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
            <TextBlock x:Name="DetailAuthorText" FontSize="12" Foreground="{StaticResource MutedTextColor}" Margin="0,0,10,0"/>
            <TextBlock x:Name="DetailDateText" FontSize="12" Foreground="{StaticResource MutedTextColor}"/>
        </StackPanel>
        
        <Border BorderBrush="{StaticResource BorderColor}" BorderThickness="1" CornerRadius="3" Padding="10">
            <TextBlock x:Name="DetailContentText" FontSize="14" TextWrapping="Wrap" LineHeight="22"/>
        </Border>

        <StackPanel x:Name="DetailLinksPanel" Margin="0,20,0,0"/>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,20,0,0">
            <Button x:Name="LaunchActionButton" Content="Launch Action" Margin="0,0,10,0" Style="{StaticResource ModernButton}"/>
            <Button x:Name="DismissButton" Content="Dismiss Message" Style="{StaticResource ModernButton}"/>
        </StackPanel>
    </StackPanel>
</ScrollViewer>
        </Grid>
    </Grid>
</Window>
"@

# ============================================================
# F) Load and Verify XAML
# ============================================================
try {
    $xmlDoc = New-Object System.Xml.XmlDocument
    $xmlDoc.LoadXml($xamlString)
    $reader = New-Object System.Xml.XmlNodeReader $xmlDoc
    [System.Windows.Window]$global:window = [Windows.Markup.XamlReader]::Load($reader)

    "MessagesListView", "DetailTitleText", "DetailAuthorText", "DetailDateText", "DetailContentText", "DetailLinksPanel", "DismissButton", "LaunchActionButton", "HeaderIcon" | ForEach-Object {
        Set-Variable -Name $_ -Value ($window.FindName($_)) -Scope Global
    }
    $global:DismissButton.Visibility = 'Collapsed'
    $global:LaunchActionButton.Visibility = 'Collapsed'

    [void]$global:LaunchActionButton.Add_Click({
        $commandPath = $global:LaunchActionButton.Tag
        if ($commandPath -and (Test-Path $commandPath)) {
            try { Start-Process -FilePath $commandPath -ErrorAction Stop }
            catch { Handle-Error "Failed to launch process: $commandPath. $($_.Exception.Message)" }
        } else {
            Handle-Error "Launch path not found or invalid: $commandPath"
        }
    })

    [void]$global:DismissButton.Add_Click({
        if ($global:MessagesListView.SelectedItem) {
            $selectedMessage = $global:MessagesListView.SelectedItem
            $state = Load-MessageState
            if ($selectedMessage.id -notin $state.dismissed) {
                $state.dismissed += $selectedMessage.id
                Save-MessageState -StateObject $state
            }
            Main-UpdateCycle
        }
    })
    
    [void]$global:MessagesListView.Add_SelectionChanged({
        if ($global:MessagesListView.SelectedItem) {
            $selectedMessage = $global:MessagesListView.SelectedItem
            
            if ($selectedMessage.isDismissible) { $global:DismissButton.Visibility = 'Visible' }
            else { $global:DismissButton.Visibility = 'Collapsed' }

            if ($selectedMessage.launchAction) {
                $global:LaunchActionButton.Visibility = 'Visible'
                $global:LaunchActionButton.Content = $selectedMessage.launchAction.buttonText
                $global:LaunchActionButton.Tag = $selectedMessage.launchAction.commandPath
            } else {
                $global:LaunchActionButton.Visibility = 'Collapsed'
            }
            
            if (-not $selectedMessage.IsRead) {
                $selectedMessage.IsRead = $true
                $state = Load-MessageState
                if ($selectedMessage.id -notin $state.read) {
                    $state.read += $selectedMessage.id
                    Save-MessageState -StateObject $state
                }
                $global:MessagesListView.Items.Refresh()
                Update-AlertIndicators
            }

            $global:DetailTitleText.Text = $selectedMessage.title
            $global:DetailAuthorText.Text = "From: $($selectedMessage.author)"
            $global:DetailDateText.Text = "Published: $([datetime]$selectedMessage.publishedDate | Get-Date -Format 'g')"
            Convert-MarkdownToTextBlock -Text $selectedMessage.content -TargetTextBlock $global:DetailContentText
            
            $global:DetailLinksPanel.Children.Clear()
            if ($selectedMessage.links) {
                foreach ($link in $selectedMessage.links) {
                    [void]$global:DetailLinksPanel.Children.Add((New-HyperlinkBlock -Name $link.Name -Url $link.Url))
                }
            }
        } else {
            $global:DetailTitleText.Text = ""
            $global:DetailAuthorText.Text = ""
            $global:DetailDateText.Text = ""
            $global:DetailContentText.Inlines.Clear()
            $global:DetailLinksPanel.Children.Clear()
            $global:DismissButton.Visibility = 'Collapsed'
            $global:LaunchActionButton.Visibility = 'Collapsed'
        }
    })

    $window.Add_Closing({$_.Cancel = $true; $window.Hide()})
    
    $mainIconPath = (Join-Path $ScriptDir "LL_LOGO.ico")
    Write-Log "Attempting to load icon from: $mainIconPath" -Level "INFO"

    if (Test-Path $mainIconPath) {
        try {
            $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
            $fileStream = New-Object System.IO.FileStream($mainIconPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
            $bitmap.BeginInit()
            $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
            $bitmap.StreamSource = $fileStream
            $bitmap.EndInit()
            $bitmap.Freeze()
            $fileStream.Close()
            $fileStream.Dispose()

            $global:window.Icon = $bitmap
            $global:HeaderIcon.Source = $bitmap
        } catch {
            Handle-Error "Failed to load the icon file at '$mainIconPath'. It may be corrupt or an unsupported format. Error: $($_.Exception.Message)" -Source "IconLoader"
        }
    }
} catch { 
    Handle-Error "Failed to load the XAML shell: $($_.Exception.Message)" -Source "XAML"; exit 
}

# ============================================================
# H) Modularized System Information Functions
# ============================================================
function Get-LocalSystemIdentity {
    $identity = @{}
    try {
        $regPath = "HKLM:\SOFTWARE\MITLL\SystemInfo"
        $identity.division = (Get-ItemProperty -Path $regPath -Name "Division" -ErrorAction SilentlyContinue).Division
    }
    catch {
        Write-Log "Could not read local system identity from registry." -Level "WARNING"
    }
    return $identity
}

function Fetch-ContentData {
    try {
        $content = (Invoke-WebRequest -Uri $config.ContentDataUrl -UseBasicParsing -TimeoutSec 15).Content
        $jsonData = $content | ConvertFrom-Json
        if (-not $jsonData.PSObject.Properties.Match('messages')) { 
            throw "The fetched JSON content is invalid because it is missing the top-level 'messages' array." 
        }
        $global:FailedFetchAttempts = 0
        try { $jsonData | ConvertTo-Json -Depth 100 | Out-File $config.CachePath -Force } catch { Write-Log "Could not save content to cache: $($_.Exception.Message)" -Level "WARNING"}
        return $jsonData
    } catch {
        $global:FailedFetchAttempts++
        Write-Log "Failed to fetch or validate content (Attempt $global:FailedFetchAttempts) - $($_.Exception.Message)" -Level "ERROR"
        if (Test-Path $config.CachePath) {
            try { 
                Write-Log "Falling back to cached content." -Level "WARNING"
                return Get-Content $config.CachePath -Raw | ConvertFrom-Json 
            }
            catch { 
                Write-Log "Failed to load cached content. Using default error page." -Level "ERROR"
                return @{ messages = @() }
            }
        }
        return @{ messages = @() }
    }
}

function Load-MessageState {
    $stateFilePath = Join-Path $ScriptDir "message_state.json"
    $defaultState = @{ read = @(); dismissed = @() }
    if (Test-Path $stateFilePath) {
        try {
            $state = Get-Content $stateFilePath -Raw | ConvertFrom-Json
            if ($state -is [array]) {
                $defaultState.read = $state
                return $defaultState
            }
            if (-not $state.PSObject.Properties.Match('read')) { $state.read = @() }
            if (-not $state.PSObject.Properties.Match('dismissed')) { $state.dismissed = @() }
            return $state
        } catch {
            Write-Log "Failed to load message state file. A new one will be created." -Level "WARNING"
            return $defaultState
        }
    }
    return $defaultState
}

function Save-MessageState {
    param($StateObject)
    $stateFilePath = Join-Path $ScriptDir "message_state.json"
    try {
        $StateObject | ConvertTo-Json | Out-File $stateFilePath -Force
    } catch {
        Handle-Error $_.Exception.Message -Source "Save-MessageState"
    }
}

function Get-FilteredMessages {
    $allMessages = (Fetch-ContentData).messages
    $localIdentity = Get-LocalSystemIdentity
    $state = Load-MessageState
    $filteredMessages = [System.Collections.Generic.List[psobject]]::new()
    foreach ($message in $allMessages) {
        if ($message.id -in $state.dismissed) { continue }
        $message | Add-Member -MemberType NoteProperty -Name "IsRead" -Value ($message.id -in $state.read)
        $message | Add-Member -MemberType NoteProperty -Name "isDismissible" -Value $true
        if (-not $message.targeting) {
            $filteredMessages.Add($message)
            continue
        }
        $isMatch = $true
        foreach ($key in $message.targeting.PSObject.Properties.Name) {
            $localValue = $localIdentity[$key]
            $targetValues = $message.targeting.$key
            if (-not $targetValues) { $isMatch = $false; break }
            if (-not ($localValue -in $targetValues)) { $isMatch = $false; break }
        }
        if ($isMatch) { $filteredMessages.Add($message) }
    }
    return $filteredMessages
}

function Get-SystemStatusMessages {
    $systemMessages = [System.Collections.Generic.List[psobject]]::new()
    $now = Get-Date
    $fixletPath = "C:\temp\X-Fixlet-Source_Count.txt"
    if (Test-Path $fixletPath) {
        $fixletContent = Get-Content $fixletPath -Raw -ErrorAction SilentlyContinue
        if ($fixletContent) {
            $systemMessages.Add([pscustomobject]@{
                id            = "system-bigfix-updates"
                publishedDate = $now.ToString("o")
                author        = "System Status"
                severity      = "Warning"
                title         = "Application Updates Available"
                content       = "Updates are available for the following applications:`n`n" + $fixletContent
                links         = @()
                IsRead        = $false
                isDismissible = $false
                launchAction  = @{ buttonText  = "Open App Updates"; commandPath = $config.BigFixSSA_Path }
            })
        }
    }
    $ecmUpdates = Get-CimInstance -ClassName CCM_SoftwareUpdate -Namespace "ROOT\ccm\ClientSDK" -ErrorAction SilentlyContinue | Where-Object { $_.ComplianceState -eq 0 }
    if ($ecmUpdates) {
        $systemMessages.Add([pscustomobject]@{
            id            = "system-ecm-updates"
            publishedDate = $now.ToString("o")
            author        = "System Status"
            severity      = "Warning"
            title         = "$($ecmUpdates.Count) Windows OS Patches Available"
            content       = "There are $($ecmUpdates.Count) pending Windows OS patches available to install via Software Center."
            links         = @()
            IsRead        = $false
            isDismissible = $false
            launchAction  = @{ buttonText  = "Open Software Center"; commandPath = "$($Env:WinDir)\CCM\SCClient.exe" }
        })
    }
    return $systemMessages
}

function New-HyperlinkBlock {
    param([string]$Name, [string]$Url)
    $tb = New-Object System.Windows.Controls.TextBlock
    $hp = New-Object System.Windows.Documents.Hyperlink
    $hp.NavigateUri = [Uri]$Url
    [void]$hp.Inlines.Add($Name)
    [void]$hp.Add_RequestNavigate({ try { Start-Process $_.Uri.AbsoluteUri } catch {} })
    [void]$tb.Inlines.Add($hp)
    return $tb
}

function Convert-MarkdownToTextBlock {
    param([string]$Text, $TargetTextBlock)
    $TargetTextBlock.Inlines.Clear()
    if ([string]::IsNullOrEmpty($Text)) { return }
    $lines = $Text -split "`r?`n"
    $isFirstLine = $true
    foreach ($line in $lines) {
        if (-not $isFirstLine) {
            [void]$TargetTextBlock.Inlines.Add([System.Windows.Documents.LineBreak]::new())
        }
        $regex = '(\[red\]|\[green\]|\[blue\]|\[yellow\]|\*\*|\*|__)(.*?)(\[/red\]|\[/green\]|\[/blue\]|\[/yellow\]|\*\*|\*|__)|(.+?)(?=\[red\]|\[green\]|\[blue\]|\[yellow\]|\*\*|\*|__|$)'
        $matches = [regex]::Matches($line, $regex)
        foreach ($match in $matches) {
            $run = New-Object System.Windows.Documents.Run
            if ($match.Groups[4].Success) {
                $run.Text = $match.Groups[4].Value
            } else {
                $run.Text = $match.Groups[2].Value
                switch ($match.Groups[1].Value) {
                    "[red]" { $run.Foreground = [System.Windows.Media.Brushes]::Red }
                    "[green]" { $run.Foreground = [System.Windows.Media.Brushes]::Green }
                    "[blue]" { $run.Foreground = [System.Windows.Media.Brushes]::Blue }
                    "[yellow]" { $run.Foreground = [System.Windows.Media.Brushes]::Goldenrod }
                    "**" { $run.FontWeight = [System.Windows.FontWeights]::Bold }
                    "*" { $run.FontStyle = [System.Windows.FontStyles]::Italic }
                    "__" { $run.TextDecorations = [System.Windows.TextDecorations]::Underline }
                }
            }
            [void]$TargetTextBlock.Inlines.Add($run)
        }
        $isFirstLine = $false
    }
}

# ============================================================
# I) Tray Icon and Alert Management
# ============================================================
function Update-AlertIndicators {
    if (-not $global:TrayIcon.Visible) { return }
    $unreadMessages = $global:MessagesListView.ItemsSource | Where-Object { -not $_.IsRead }
    if ($unreadMessages) {
        $global:TrayIcon.Icon = $global:WarningIcon
        $global:TrayIcon.Text = "Endpoint Advisor - New Messages"
    } else {
        $global:TrayIcon.Icon = $global:MainIcon
        $global:TrayIcon.Text = "Endpoint Advisor"
    }
}

function Toggle-WindowVisibility {
    $window.Dispatcher.Invoke({
        if ($window.IsVisible) {
            $window.Hide()
        } else {
            $window.Show()
            $window.Activate()
        }
    })
}

function Initialize-TrayIcon {
    $global:MainIcon = New-Object System.Drawing.Icon (Join-Path $ScriptDir "LL_LOGO.ico")
    $global:WarningIcon = New-Object System.Drawing.Icon (Join-Path $ScriptDir "LL_LOGO_MSG.ico")
    $global:TrayIcon = New-Object System.Windows.Forms.NotifyIcon -Property @{ Icon = $global:MainIcon; Text = "Endpoint Advisor"; Visible = $true }
    $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    [void]$contextMenu.Items.AddRange(@(
        (New-Object System.Windows.Forms.ToolStripMenuItem("Show", $null, { Toggle-WindowVisibility })),
        (New-Object System.Windows.Forms.ToolStripMenuItem("Refresh", $null, { Main-UpdateCycle })),
        (New-Object System.Windows.Forms.ToolStripMenuItem("Exit", $null, { $window.Dispatcher.InvokeShutdown() }))
    ))
    $global:TrayIcon.ContextMenuStrip = $contextMenu
    [void]$global:TrayIcon.add_MouseClick({ if ($_.Button -eq 'Left') { Toggle-WindowVisibility } })
}

# ============================================================
# O) Main Update Cycle and DispatcherTimer
# ============================================================
function Main-UpdateCycle {
    try {
        Write-Log "Main update cycle running..." -Level "INFO"
        $systemMessages = @(Get-SystemStatusMessages)
        $serverMessages = @(Get-FilteredMessages)
        $allDisplayMessages = $systemMessages + $serverMessages
        
        $selectedIndex = $global:MessagesListView.SelectedIndex
        $global:MessagesListView.ItemsSource = $allDisplayMessages
        if ($selectedIndex -ge 0 -and $selectedIndex -lt $allDisplayMessages.Count) {
            $global:MessagesListView.SelectedIndex = $selectedIndex
        }
        Update-AlertIndicators
    } catch { Handle-Error $_.Exception.Message -Source "Main-UpdateCycle" }
}

# ============================================================
# P) Initial Setup & Application Start
# ============================================================
try {
    $global:DispatcherTimer = New-Object System.Windows.Threading.DispatcherTimer -Property @{ Interval = [TimeSpan]::FromSeconds($config.RefreshInterval) }
    [void]$global:DispatcherTimer.add_Tick({ Main-UpdateCycle })

    Initialize-TrayIcon
    Main-UpdateCycle
    
    $global:DispatcherTimer.Start()
    Write-Log "Application startup complete. Running dispatcher." -Level "INFO"
    [System.Windows.Threading.Dispatcher]::Run()
} catch { 
    Handle-Error "Critical startup error: $($_.Exception.Message)" -Source "Startup" 
} finally {
    Write-Log "--- Lincoln Laboratory Endpoint Advisor Script Exiting ---"
    if ($global:DispatcherTimer) { $global:DispatcherTimer.Stop() }
    if ($global:TrayIcon) { $global:TrayIcon.Dispose() }
    if ($global:mutex) { $global:mutex.Dispose() }
}
