# ===== LLEA CORE HELPERS (added) =====
# Version: 6.2.2 (Fixed multiple instance issues)


function Test-IsJson {
    param([string]$s)
    if ([string]::IsNullOrWhiteSpace($s)) { return $false }
    $t = $s.Trim()
    if (-not ($t.StartsWith('{') -or $t.StartsWith('['))) { return $false }
    try { $null = $t | ConvertFrom-Json -Depth 64; return $true } catch { return $false }
}

function Get-RemoteContentData {
    param(
        [Parameter(Mandatory)] [string]$Url,
        [int]$MaxAttempts = 3
    )
    $headers = @{
        'User-Agent'    = 'LLEA/1.0'
        'Accept'        = 'application/json'
        'Cache-Control' = 'no-cache'
        'Pragma'        = 'no-cache'
    }
    for ($i=1; $i -le $MaxAttempts; $i++) {
        try {
            $resp = Invoke-WebRequest -Uri $Url -Headers $headers -UseBasicParsing -MaximumRedirection 5 -TimeoutSec 20
            $body = [string]$resp.Content
            if (Test-IsJson $body) {
                $obj = $body | ConvertFrom-Json -Depth 64
                $dataNode = $null
                if ($obj -and ($obj.PSObject.Properties.Name -contains 'Data') -and $obj.Data) {
                    $dataNode = $obj.Data
                } else {
                    $dataNode = $obj
                }
                $keys = ($dataNode | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name) -join ','
                Write-Log ("Fetched content keys: {0}" -f $keys) -Level "INFO"
                return @{
                    Data      = $dataNode
                    Source    = $Url
                    Retrieved = (Get-Date).ToString('s')
                }
            } else {
                $first200 = if ($body.Length -gt 200) { $body.Substring(0,200) + '...' } else { $body }
                Write-Log ('Response does not appear to be JSON. First 200 chars: {0}' -f $first200) -Level 'WARNING'
            }
        } catch {
            Write-Log ("Fetch attempt {0} failed - {1}" -f $i, $_.Exception.Message) -Level 'ERROR'
        }
        Start-Sleep -Seconds ([Math]::Min(2*$i, 6))
    }
    return $null
}

function Load-ContentDataFromCache {
    param([string]$Path = 'C:\Program Files\LLEA\ContentData.cache.json')
    try {
        if (Test-Path -LiteralPath $Path) {
            $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
            if (Test-IsJson $raw) {
                $o = $raw | ConvertFrom-Json -Depth 64
                $dataNode = $null
                if ($o -and ($o.PSObject.Properties.Name -contains 'Data') -and $o.Data) {
                    $dataNode = $o.Data
                } else {
                    $dataNode = $o
                }
                $keys = ($dataNode | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name) -join ','
                Write-Log ("Cached content keys: {0}" -f $keys) -Level "INFO"
                return @{ Data = $dataNode; Source = "cache:$Path"; Retrieved = (Get-Date).ToString('s') }
            }
        }
    } catch {
        Write-Log ("Cache load failed: {0}" -f $_) -Level "WARNING"
    }
    return $null
}

function Normalize-TextStable {
    param([string]$s)
    if ($null -eq $s) { return '' }
    $t = [string]$s
    $t = $t -replace "`r`n","`n"
    $t = $t -replace "`r","`n"
    $t = $t -replace "[ \t]+`n","`n"
    $t = $t.Trim()
    return $t
}

function Normalize-UrlStable {
    param([string]$u)
    if ([string]::IsNullOrWhiteSpace($u)) { return '' }
    try {
        $uri = [System.Uri]$u
        $schemeHost = ($uri.Scheme + '://' + $uri.Host).ToLowerInvariant()
        if ($uri.IsDefaultPort) { $portPart = '' } else { $portPart = ':' + $uri.Port }
        $pathQuery = $uri.PathAndQuery + $uri.Fragment
        return $schemeHost + $portPart + $pathQuery
    } catch { return $u.Trim() }
}

function Get-TextSha256 {
    param([string]$Text)
    if ($null -eq $Text) { $Text = '' }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $b = [System.Text.Encoding]::UTF8.GetBytes($Text)
        return ($sha.ComputeHash($b) | ForEach-Object { $_.ToString('x2') }) -join ''
    } finally { $sha.Dispose() }
}

function Ensure-SectionStateStore {
    if (-not $script:config) { $script:config = @{} }
    if (-not $script:config.SectionStates) { $script:config.SectionStates = @{} }
}

function Test-SectionChanged {
    param(
        [Parameter(Mandatory)][string]$SectionKey,
        [Parameter(Mandatory)][string]$NewStateJson
    )
    Ensure-SectionStateStore
    $prev = $script:config.SectionStates[$SectionKey]
    $prevHash = if ($prev) { Get-TextSha256 $prev } else { '(none)' }
    $newHash  = Get-TextSha256 $NewStateJson
    Write-Log ("{0}: prevSHA={1} newSHA={2}" -f $SectionKey, $prevHash, $newHash) -Level "INFO"

    if ([string]::IsNullOrEmpty($prev)) {
        Write-Log ("{0}: first run; baselining without alert." -f $SectionKey) -Level "INFO"
        return $false
    }
    if ($prev -ne $NewStateJson) {
        Write-Log ("{0}: content changed; alerting." -f $SectionKey) -Level "INFO"
        return $true
    }
    Write-Log ("{0}: content unchanged; no alert." -f $SectionKey) -Level "INFO"
    return $false
}

function Save-SectionBaseline {
    param(
        [Parameter(Mandatory)][string]$SectionKey,
        [Parameter(Mandatory)][string]$NewStateJson
    )
    Ensure-SectionStateStore
    $script:config.SectionStates[$SectionKey] = $NewStateJson
    Write-Log ("{0}: baseline saved (SHA={1})." -f $SectionKey, (Get-TextSha256 $NewStateJson)) -Level "INFO"
    if ($SectionKey -eq 'Support') { $script:config.SupportLastState = $NewStateJson }
    if ($SectionKey -eq 'Announcements') { $script:config.AnnouncementsLastState = $NewStateJson }
}
# ===== END CORE HELPERS =====

# Ensure $PSScriptRoot is defined for older versions
if ($MyInvocation.MyCommand.Path) {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    $ScriptDir = Get-Location
}

# Define version
$ScriptVersion = "6.2.2"

# --- START OF ENHANCED SINGLE-INSTANCE CHECK ---
# Uses multiple methods to prevent duplicate instances:
# 1. Global mutex (most reliable)
# 2. Process/command-line checking (backup)
# 3. Lock file with PID tracking (tertiary)
$AppName = "Lincoln Laboratory Endpoint Advisor"
$MutexName = "Global\LLEA_SingleInstance_Mutex"
$LockFilePath = "$env:TEMP\LLEA_Instance.lock"

function Test-OtherInstanceRunning {
    param([string]$CurrentPID = $pid)
    $otherInstances = @()
    
    # Method 1: Check by command line
    try {
        $processes = Get-WmiObject Win32_Process -Filter "Name = 'powershell.exe' OR Name = 'pwsh.exe'" -ErrorAction SilentlyContinue
        foreach ($proc in $processes) {
            if ($proc.ProcessId -ne $CurrentPID -and $proc.CommandLine -like "*LLEA.ps1*") {
                $otherInstances += $proc.ProcessId
            }
        }
    } catch { Write-Host "[WARNING] Could not check processes via WMI: $_" }
    
    # Method 2: Check by MainWindowTitle (original method)
    try {
        $existingByTitle = Get-Process -Name "powershell", "pwsh" -ErrorAction SilentlyContinue | 
            Where-Object { $_.Id -ne $CurrentPID -and $_.MainWindowTitle -like "*$AppName*" }
        foreach ($proc in $existingByTitle) {
            if ($proc.Id -notin $otherInstances) { $otherInstances += $proc.Id }
        }
    } catch { Write-Host "[WARNING] Could not check processes by window title: $_" }
    
    # Method 3: Check lock file
    if (Test-Path $LockFilePath) {
        try {
            $lockContent = Get-Content $LockFilePath -Raw -ErrorAction SilentlyContinue
            if ($lockContent -match "PID:(\d+)") {
                $lockedPID = [int]$matches[1]
                $lockedProcess = Get-Process -Id $lockedPID -ErrorAction SilentlyContinue
                if ($lockedProcess -and $lockedProcess.Id -ne $CurrentPID) {
                    try {
                        $procInfo = Get-WmiObject Win32_Process -Filter "ProcessId = $lockedPID" -ErrorAction SilentlyContinue
                        if ($procInfo -and $procInfo.CommandLine -like "*LLEA.ps1*") {
                            if ($lockedPID -notin $otherInstances) { $otherInstances += $lockedPID }
                        } else {
                            Remove-Item $LockFilePath -Force -ErrorAction SilentlyContinue
                        }
                    } catch { Write-Host "[WARNING] Could not verify locked process: $_" }
                } elseif (-not $lockedProcess) {
                    Remove-Item $LockFilePath -Force -ErrorAction SilentlyContinue
                }
            }
        } catch { Write-Host "[WARNING] Error reading lock file: $_" }
    }
    return $otherInstances
}

function Create-LockFile {
    param([string]$Path = $LockFilePath, [string]$CurrentPID = $pid)
    try {
        $lockContent = @"
LLEA Instance Lock File
Created: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
PID:$CurrentPID
Script:$($MyInvocation.MyCommand.Path)
"@
        Set-Content -Path $Path -Value $lockContent -Force -ErrorAction Stop
        return $true
    } catch { Write-Host "[WARNING] Could not create lock file: $_"; return $false }
}

function Remove-LockFile {
    param([string]$Path = $LockFilePath)
    try {
        if (Test-Path $Path) {
            Remove-Item $Path -Force -ErrorAction Stop
        }
    } catch { Write-Host "[WARNING] Could not remove lock file: $_" }
}

# Try to acquire mutex
$script:InstanceMutex = $null
$mutexAcquired = $false
try {
    $script:InstanceMutex = New-Object System.Threading.Mutex($false, $MutexName)
    $mutexAcquired = $script:InstanceMutex.WaitOne(0, $false)
} catch { Write-Host "[WARNING] Mutex creation/check failed: $_" }

# Check for other running instances
$otherInstances = Test-OtherInstanceRunning -CurrentPID $pid

# Decide whether to proceed
if (-not $mutexAcquired -or $otherInstances.Count -gt 0) {
    Write-Host ""
    Write-Host ("=" * 60)
    Write-Host "ANOTHER INSTANCE OF $AppName IS ALREADY RUNNING"
    Write-Host ("=" * 60)
    Write-Host "  - Mutex acquired: $mutexAcquired"
    Write-Host "  - Other instances found: $($otherInstances.Count)"
    if ($otherInstances.Count -gt 0) { Write-Host "  - Other PIDs: $($otherInstances -join ', ')" }
    Write-Host ("=" * 60)
    
    if ($mutexAcquired -and $script:InstanceMutex) {
        try { $script:InstanceMutex.ReleaseMutex(); $script:InstanceMutex.Dispose() } catch {}
    }
    Start-Sleep -Seconds 2
    exit
}

# Create lock file
Create-LockFile -Path $LockFilePath -CurrentPID $pid

# Register cleanup on exit
$script:CleanupRegistered = $false
if (-not $script:CleanupRegistered) {
    Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
        if ($script:InstanceMutex) {
            try { $script:InstanceMutex.ReleaseMutex(); $script:InstanceMutex.Dispose() } catch {}
        }
        Remove-LockFile -Path $using:LockFilePath
    } | Out-Null
    $script:CleanupRegistered = $true
}
# --- END OF ENHANCED SINGLE-INSTANCE CHECK ---

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

# Global flag to track pending restart state - REMOVED
# Pending restart checking has been disabled
# $global:PendingRestart = $false
# $global:RestartAlertAcknowledged = $false

# Global flag to track pending update state
$global:UpdatesPending = $false
$global:CurrentUpdateState = ""
$global:LastAnnouncementState = ""

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

# Force TLS 1.2 for web requests (required for GitHub and most modern HTTPS sites)
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Write-Log "TLS 1.2 protocol enabled for web requests" -Level "INFO"
} catch {
    Write-Log "Warning: Could not set TLS protocol - $($_.Exception.Message)" -Level "WARNING"
}

# Skip certificate validation for internal GitHub (GlobalProtect SSL inspection compatibility)
try {
    add-type @"
        using System.Net;
        using System.Security.Cryptography.X509Certificates;
        public class TrustAllCertsPolicy : ICertificatePolicy {
            public bool CheckValidationResult(
                ServicePoint srvPoint, X509Certificate certificate,
                WebRequest request, int certificateProblem) {
                return true;
            }
        }
"@
    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
    Write-Log "Certificate validation configured for GlobalProtect compatibility" -Level "INFO"
} catch {
    # Type might already be loaded, that's okay
    Write-Log "Certificate policy already configured" -Level "INFO"
}

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
        # Example of platform-targeted announcement (not active by default):
        # @{
        #     Platform = "Windows"  # Can be "Windows", "Mac", or "All"
        #     Enabled = $true
        #     AppendToDefault = $false
        #     Text = "Windows-specific message"
        #     Details = "Additional details"
        #     Links = @()
        #     Condition = @{ Type = "Registry"; Path = "HKLM:\..."; Name = "..."; Value = "..." }
        # }
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
        <TextBlock Text="Lincoln Laboratory Endpoint Advisor" FontSize="14" FontWeight="Bold" Foreground="White" VerticalAlignment="Center"/>
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
          <StackPanel x:Name="AppendedAnnouncementsPanel" Orientation="Vertical" Margin="0,5,0,0" Visibility="Collapsed"/>
          <TextBlock x:Name="AnnouncementsSourceText" FontSize="9" Foreground="Gray" Margin="0,5,0,0"/>
        </StackPanel>
      </Border>
    </Expander>
    <Expander x:Name="PatchingExpander" FontSize="12" IsExpanded="False" Margin="0,2,0,2">
      <Expander.Header>
        <StackPanel Orientation="Horizontal">
          <TextBlock Text="Patching and Updates" VerticalAlignment="Center"/>
          <Ellipse x:Name="PatchingAlertIcon" Width="10" Height="10" Margin="5,0,0,0" Fill="Red" Visibility="Hidden"/>
        </StackPanel>
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

          <Separator Margin="0,8,0,8"/>

          <Grid Margin="0,2,0,2">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock x:Name="ECMStatusText" Grid.Column="0" VerticalAlignment="Center" FontSize="11" TextWrapping="Wrap"/>
            <Button x:Name="ECMLaunchButton" Grid.Column="1" Content="Install Patches" Margin="10,0,0,0" Padding="5,1" VerticalAlignment="Center" Visibility="Collapsed" ToolTip="Install pending Windows OS patches"/>
          </Grid>

          <Separator Margin="0,8,0,8"/>

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

          <Border x:Name="DriverProgressPanel" BorderBrush="#0078D7" BorderThickness="2" Background="#F0F8FF" Padding="10" CornerRadius="3" Margin="0,10,0,0" Visibility="Collapsed">
            <StackPanel>
              <Grid>
                <TextBlock Text="Driver Update Progress" FontSize="11" FontWeight="Bold" HorizontalAlignment="Left" VerticalAlignment="Center"/>
                <Button x:Name="DriverProgressCloseButton" Content="X" Width="22" Height="22" HorizontalAlignment="Right" VerticalAlignment="Top" Padding="0" FontSize="11" FontWeight="Bold" ToolTip="Close progress panel"/>
              </Grid>
              <TextBlock x:Name="DriverProgressStatus" FontSize="11" TextWrapping="Wrap" Text="Initializing..." Margin="0,5,0,0"/>
              <ProgressBar x:Name="DriverProgressBar" Height="20" Margin="0,10,0,0" IsIndeterminate="True"/>
              <TextBlock FontSize="9" Foreground="Gray" Margin="0,5,0,0" Text="This process may take several minutes. You can continue working."/>
            </StackPanel>
          </Border>
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
          <TextBlock x:Name="SupportDetailsText" FontSize="11" TextWrapping="Wrap" Margin="0,5,0,0"/>
          <StackPanel x:Name="SupportLinksPanel" Orientation="Vertical" Margin="0,5,0,0"/>
          <TextBlock x:Name="SupportSourceText" FontSize="9" Foreground="Gray" Margin="0,5,0,0"/>
        </StackPanel>
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
    "PendingRestartPanel", "PendingRestartStatusText", "SupportExpander", "SupportAlertIcon", "SupportText", "SupportDetailsText", "SupportLinksPanel",
    "SupportSourceText", "WindowsBuildText", "ClearAlertsButton",
    "FooterText", "ClearAlertsPanel", "ClearAlertsDot", "BigFixStatusText", "BigFixLaunchButton", "ECMStatusText", "ECMLaunchButton",
    "PatchingAlertIcon", "AppendedAnnouncementsPanel",
    "DriverUpdateStatusText", "DriverUpdateLastRunText", "DriverUpdateButton", "DriverProgressPanel", 
    "DriverProgressStatus", "DriverProgressBar", "DriverProgressCloseButton"  # ADD THIS LINE
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
                if ($global:AnnouncementsAlertIcon) { $global:AnnouncementsAlertIcon.Visibility = "Hidden" 
$script:UIReady = $true
Write-Log "UI initialized; running first update cycle." -Level "INFO"
Main-UpdateCycle
}

                Update-TrayIcon
            })
        }
        if ($global:SupportExpander) {
            $global:SupportExpander.IsExpanded = $true
            $global:SupportExpander.Add_Expanded({ 
                if ($global:SupportAlertIcon) { $global:SupportAlertIcon.Visibility = "Hidden" }
                # Save the state when user views it
                $config.SupportLastState = Get-StableSupportStateJson (Get-CurrentSupportObject $global:contentData.Data.Support)
                Save-Configuration -Config $config
                Update-TrayIcon
            })
        }
        if ($global:PatchingExpander) {
            $global:PatchingExpander.Add_Expanded({
                if ($global:PatchingAlertIcon) { $global:PatchingAlertIcon.Visibility = "Hidden" }
                $global:UpdatesPending = $false 
                $config.LastSeenUpdateState = $global:CurrentUpdateState
                Save-Configuration -Config $config
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
                Start-DriverUpdate
            })
        }

        if ($global:DriverProgressCloseButton) {
            $global:DriverProgressCloseButton.Add_Click({
                Write-Log "User closed driver progress panel." -Level "INFO"
                $window.Dispatcher.Invoke({
                    $global:DriverProgressPanel.Visibility = "Collapsed"
                })
            })
        }
        if ($global:ClearAlertsButton) {
			$global:ClearAlertsButton.Add_Click({
				Write-Log "Clear Alerts button clicked by user to clear new alerts (red dots)." -Level "INFO"
        
			$config.AnnouncementsLastState = $global:LastAnnouncementState
			$config.SupportLastState = Get-StableSupportStateJson (Get-CurrentSupportObject $global:contentData.Data.Support)
                
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
function Get-ActiveAnnouncement {
    param($AnnouncementsObject)
    
    Write-Log "Get-ActiveAnnouncement starting..." -Level "INFO"
    
    # Determine current platform
    $currentPlatform = "Windows"
    
    # Initialize result
    $base = $null
    $appended = @()
    
    # Handle Default announcements
    if ($AnnouncementsObject.Default) {
        Write-Log "Processing Default announcements (Type: $($AnnouncementsObject.Default.GetType().Name))" -Level "INFO"
        
        if ($AnnouncementsObject.Default -is [System.Array]) {
            Write-Log "Default is an array with $($AnnouncementsObject.Default.Count) items" -Level "INFO"
            
            foreach ($defaultItem in $AnnouncementsObject.Default) {
                $itemPlatform = $defaultItem.Platform
                Write-Log "Checking default item with Platform: $itemPlatform" -Level "INFO"
                
                if (-not $itemPlatform -or $itemPlatform -eq "All" -or $itemPlatform -eq $currentPlatform) {
                    $base = $defaultItem
                    Write-Log "Found matching default announcement: $($defaultItem.Text)" -Level "INFO"
                    break
                }
            }
            
            # Fallback to first if no platform match
            if (-not $base -and $AnnouncementsObject.Default.Count -gt 0) {
                $base = $AnnouncementsObject.Default[0]
                Write-Log "No platform match, using first default announcement" -Level "INFO"
            }
        } else {
            # Single default object
            $base = $AnnouncementsObject.Default
            Write-Log "Default is a single object" -Level "INFO"
            
            if ($base.Platform -and $base.Platform -ne "All" -and $base.Platform -ne $currentPlatform) {
                Write-Log "Platform mismatch but using anyway: $($base.Platform)" -Level "WARNING"
            }
        }
    } else {
        Write-Log "No Default announcements found!" -Level "WARNING"
    }
    
    # Handle Targeted announcements
    if ($AnnouncementsObject.Targeted) {
        Write-Log "Processing Targeted announcements (Type: $($AnnouncementsObject.Targeted.GetType().Name))" -Level "INFO"
        
        # Ensure Targeted is treated as array
        $targetedArray = @()
        if ($AnnouncementsObject.Targeted -is [System.Array]) {
            $targetedArray = $AnnouncementsObject.Targeted
        } else {
            # Wrap single object in array
            $targetedArray = @($AnnouncementsObject.Targeted)
        }
        
        Write-Log "Found $($targetedArray.Count) targeted announcements" -Level "INFO"
        
        foreach ($targeted in $targetedArray) {
            Write-Log "Checking targeted: Platform=$($targeted.Platform), Enabled=$($targeted.Enabled), Text=$($targeted.Text)" -Level "INFO"
            
            # Check if enabled
            if ($targeted.Enabled -ne $true) {
                Write-Log "  Skipping - not enabled" -Level "INFO"
                continue
            }
            
            # Platform check
            if ($targeted.Platform -and $targeted.Platform -ne "All" -and $targeted.Platform -ne $currentPlatform) {
                Write-Log "  Skipping - platform mismatch (needs: $($targeted.Platform))" -Level "INFO"
                continue
            }
            
            # Condition check
            $conditionMet = $false
            
            if ($targeted.Condition) {
                if ($targeted.Condition.Type -eq "Registry") {
                    $path = $targeted.Condition.Path
                    $name = $targeted.Condition.Name
                    $value = $targeted.Condition.Value
                    
                    Write-Log "  Checking registry: $path\$name = $value" -Level "INFO"
                    
                    try {
                        $reg = Get-ItemProperty -Path $path -Name $name -ErrorAction Stop
                        if ($reg.$name -eq $value) {
                            Write-Log "  Registry condition MET!" -Level "INFO"
                            $conditionMet = $true
                        } else {
                            Write-Log "  Registry value mismatch: Expected=$value, Actual=$($reg.$name)" -Level "INFO"
                        }
                    } catch {
                        Write-Log "  Registry not found: $_" -Level "INFO"
                    }
                }
            } else {
                # No condition means always show
                Write-Log "  No condition - always show" -Level "INFO"
                $conditionMet = $true
            }
            
            if ($conditionMet) {
                if ($targeted.AppendToDefault) {
                    Write-Log "  Adding to appended list" -Level "INFO"
                    $appended += $targeted
                } else {
                    Write-Log "  Replacing default announcement" -Level "INFO"
                    return @{
                        Base = $targeted
                        Appended = @()
                    }
                }
            }
        }
    } else {
        Write-Log "No Targeted announcements found" -Level "INFO"
    }
    
    Write-Log "Returning: Base=$($base -ne $null), Appended=$($appended.Count)" -Level "INFO"
    
    return @{
        Base = $base
        Appended = $appended
    }
}


# ============================================================
# JSON State Helper Functions
# ============================================================
function Get-StableAnnouncementStateJson {
    param($AnnouncementResult)
    
    # Create a stable JSON representation of the announcement state
    # This is used for comparison to detect changes
    $stateObj = @{
        Base = if ($AnnouncementResult.Base) { 
            @{
                Text = $AnnouncementResult.Base.Text
                Details = $AnnouncementResult.Base.Details
                Links = $AnnouncementResult.Base.Links
            }
        } else { $null }
        Appended = @()
    }
    
    foreach ($item in $AnnouncementResult.Appended) {
        $stateObj.Appended += @{
            Text = $item.Text
            Details = $item.Details
            Links = $item.Links
        }
    }
    
    return ($stateObj | ConvertTo-Json -Compress -Depth 10)
}

function Get-StableSupportStateJson {
    param($SupportObject)
    
    # Create a stable JSON representation of the support state
    # This is used for comparison to detect changes
    if (-not $SupportObject) { return "{}" }
    
    $stateObj = @{
        Text = $SupportObject.Text
        Details = $SupportObject.Details
        Links = $SupportObject.Links
    }
    
    return ($stateObj | ConvertTo-Json -Compress -Depth 10)
}

function Get-CurrentSupportObject {
    param($SupportData)
    
    # Extract the current platform-specific support object
    $currentPlatform = "Windows"
    
    if ($SupportData -is [System.Array]) {
        # New format: Array of platform-specific support sections
        foreach ($supportItem in $SupportData) {
            $itemPlatform = $supportItem.Platform
            if (-not $itemPlatform -or $itemPlatform -eq "All" -or $itemPlatform -eq $currentPlatform) {
                return $supportItem
            }
        }
        # If no platform match found, use first entry as fallback
        if ($SupportData.Count -gt 0) {
            return $SupportData[0]
        }
    } else {
        # Old format: Single support object
        return $SupportData
    }
    
    return $null
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
    # Handle new structure with schemaVersion and nested Data
    if ($Data.PSObject.Properties.Match('schemaVersion') -and $Data.PSObject.Properties.Match('Data')) {
        # New structure - check nested Data object
        if (-not $Data.Data.PSObject.Properties.Match('Announcements') -or -not $Data.Data.PSObject.Properties.Match('Support')) {
            throw "JSON data is missing 'Announcements' or 'Support' in Data property."
        }
    } else {
        # Old structure - check top level
        if (-not $Data.PSObject.Properties.Match('Announcements') -or -not $Data.PSObject.Properties.Match('Support')) {
            throw "JSON data is missing 'Announcements' or 'Support' top-level property."
        }
    }
    return $true
}


# Function to normalize content data to handle both old and new JSON structures
function Normalize-ContentData {
    param($Data)
    
    # If this is the new structure with schemaVersion, extract the nested Data
    if ($Data.PSObject.Properties.Match('schemaVersion') -and $Data.PSObject.Properties.Match('Data')) {
        Write-Log "Detected new JSON structure with schemaVersion: $($Data.schemaVersion)" -Level "INFO"
        return $Data.Data
    }
    
    # Otherwise return as-is (old structure)
    return $Data
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
                Data = (Normalize-ContentData -Data $contentData)
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
        return [PSCustomObject]@{ Data = (Normalize-ContentData -Data $defaultContentData); Source = "Default" }
    }
    $url = $config.ContentDataUrl

    try {
        Write-Log "Attempting to fetch content from: $url" -Level "INFO"
        
        $response = Invoke-WithRetry -Action {
            # Configure WebClient with proxy and credentials
            $webClient = New-Object System.Net.WebClient
            $webClient.Headers.Add("User-Agent", "PowerShell-EndpointAdvisor/6.0")
            $webClient.Encoding = [System.Text.Encoding]::UTF8
            $webClient.UseDefaultCredentials = $true
            
            # Get system proxy with credentials
            $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
            $proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
            $webClient.Proxy = $proxy
            
            try {
                $content = $webClient.DownloadString($url)
                return $content
            } finally {
                $webClient.Dispose()
            }
        } -MaxRetries 3 -RetryDelayMs 1000

        if (-not $response) {
            throw "Empty response received from web request"
        }

        Write-Log "Successfully fetched content from remote source." -Level "INFO"
        
        # Validate it looks like JSON before parsing
        $trimmed = $response.Trim()
        if (-not ($trimmed.StartsWith("{") -or $trimmed.StartsWith("["))) {
            Write-Log "Response does not appear to be JSON. First 200 chars: $($trimmed.Substring(0, [Math]::Min(200, $trimmed.Length)))" -Level "ERROR"
            throw "Response is not valid JSON format"
        }
        
        # Parse JSON
        try {
            $contentData = $response | ConvertFrom-Json -ErrorAction Stop
        } catch {
            Write-Log "JSON parsing failed. Error: $($_.Exception.Message)" -Level "ERROR"
            throw
        }
        
        Validate-ContentData -Data $contentData
        Write-Log "Content data validated successfully." -Level "INFO"
        Save-CachedContentData -ContentData ([PSCustomObject]@{ Data = $contentData })
        $global:FailedFetchAttempts = 0

        return [PSCustomObject]@{ Data = (Normalize-ContentData -Data $contentData); Source = "Remote" }
    }
    catch {
        $global:FailedFetchAttempts++
        Write-Log "Failed to fetch or validate content from $url (Attempt $global:FailedFetchAttempts) - $($_.Exception.Message)" -Level "ERROR"
        
        if ($_.Exception.InnerException) {
            Write-Log "Inner exception: $($_.Exception.InnerException.Message)" -Level "ERROR"
        }
        
        if ($global:FailedFetchAttempts -ge 3) {
            Write-Log "Multiple consecutive fetch failures. Possible proxy authentication or connectivity issue." -Level "WARNING"
        }
        
        $cachedData = Load-CachedContentData
        if ($cachedData) {
            Write-Log "Using cached content data as fallback." -Level "INFO"
            return $cachedData
        }
        
        Write-Log "No cached data available, using default content." -Level "WARNING"
        return [PSCustomObject]@{ Data = (Normalize-ContentData -Data $defaultContentData); Source = "Default" }
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
        
        $alertDays = $config.YubiKeyAlertDays  # 14 days by default
        $now = Get-Date
        $alerts = @()
        
        # Check YubiKey certificates
        try {
            $ykmanPath = $config.YubiKeyManager_Path
            if (-not [string]::IsNullOrWhiteSpace($ykmanPath) -and (Test-Path $ykmanPath)) {
                if (& $ykmanPath info 2>$null) {
                    $slots = @("9a", "9c", "9d", "9e")
                    foreach ($slot in $slots) {
                        $certPem = & $ykmanPath "piv" "certificates" "export" $slot "-" 2>$null
                        if ($certPem -and $certPem -match "-----BEGIN CERTIFICATE-----") {
                            $tempFile = [System.IO.Path]::GetTempFileName()
                            $certPem | Out-File $tempFile -Encoding ASCII
                            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($tempFile)
                            Remove-Item $tempFile -Force
                            
                            $daysUntilExpiry = ($cert.NotAfter - $now).Days
                            if ($daysUntilExpiry -le $alertDays) {
                                if ($daysUntilExpiry -lt 0) {
                                    $alerts += "Your YubiKey certificate (Slot $slot) has EXPIRED on $($cert.NotAfter.ToString('yyyy-MM-dd'))."
                                } else {
                                    $alerts += "Your YubiKey certificate (Slot $slot) will expire in $daysUntilExpiry day(s) on $($cert.NotAfter.ToString('yyyy-MM-dd'))."
                                }
                            }
                        }
                    }
                }
            }
        } catch {
            Write-Log "YubiKey certificate check error - $($_.Exception.Message)" -Level "WARNING"
        }
        
        # Show alert if any certificates are expiring soon
        if ($alerts.Count -gt 0) {
            $alertMessage = "Certificate Expiration Alert:`n`n" + ($alerts -join "`n`n") + "`n`nPlease renew your certificate(s) as soon as possible."
            Write-Log "Certificate expiration alert: $($alerts -join '; ')" -Level "WARNING"
            
            $window.Dispatcher.Invoke({
                [System.Windows.MessageBox]::Show(
                    $alertMessage,
                    "Certificate Expiration Warning",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning
                )
            })
        } else {
            Write-Log "All certificates are valid for more than $alertDays days" -Level "INFO"
        }
        
        $global:CachedCertificateStatus = if ($alerts.Count -gt 0) { "Alerts: $($alerts.Count)" } else { "Valid" }
        
    } catch {
        Write-Log "Error checking certificate expiration - $($_.Exception.Message)" -Level "ERROR"
    }
}

function Get-PendingRestartStatus {
    # REMOVED: Pending restart checking has been disabled
    # This function now always returns "No system restart required."
    return "No system restart required."
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

function Get-DaysSinceLastDriverUpdate {
    $logFile = "C:\Windows\MITLL\Logs\MS_Update.txt"
    
    if (Test-Path $logFile) {
        try {
            $content = Get-Content $logFile -ErrorAction Stop
            
            # Find all dates in MM/DD/YYYY format
            $datePattern = '(\d{2}/\d{2}/\d{4})'
            $dates = @()
            
            foreach ($line in $content) {
                if ($line -match $datePattern) {
                    try {
                        $dateStr = $matches[1]
                        $parsedDate = [DateTime]::ParseExact($dateStr, "MM/dd/yyyy", $null)
                        $dates += $parsedDate
                    } catch {
                        # Skip invalid dates
                    }
                }
            }
            
            if ($dates.Count -eq 0) {
                return "Never run"
            }
            
            # Get the most recent date
            $lastRun = ($dates | Sort-Object -Descending)[0]
            $daysSince = ([DateTime]::Now - $lastRun).Days
            
            if ($daysSince -eq 0) {
                return "Last run today"
            } elseif ($daysSince -eq 1) {
                return "Last run 1 day ago"
            } else {
                return "Last run $daysSince days ago"
            }
        } catch {
            Write-Log "Error reading MS_Update.txt: $($_.Exception.Message)" -Level "ERROR"
            return "Unable to determine last run"
        }
    } else {
        return "Never run"
    }
}

function Update-DriverUpdateStatus {
    try {
        $driverLastRun = Get-DaysSinceLastDriverUpdate
        
        # Determine if button should show (25+ days or never run)
        $showDriverButton = $false
        if ($driverLastRun -eq "Never run") {
            $showDriverButton = $true
        } elseif ($driverLastRun -match "Last run (\d+) day") {
            $days = [int]$matches[1]
            if ($days -ge 25) {
                $showDriverButton = $true
            }
        }
        
        $window.Dispatcher.Invoke({
            $global:DriverUpdateButton.Visibility = if ($showDriverButton) { "Visible" } else { "Collapsed" }
            $global:DriverUpdateLastRunText.Text = $driverLastRun
            
            # Check if last run was 25+ days ago for alert styling
            $daysOverdue = $false
            if ($driverLastRun -match "Last run (\d+) day") {
                $daysSince = [int]$matches[1]
                if ($daysSince -ge 25) {
                    $daysOverdue = $true
                }
            } elseif ($driverLastRun -eq "Never run") {
                $daysOverdue = $true
            }
            
            if ($daysOverdue) {
                $global:DriverUpdateLastRunText.FontWeight = "Bold"
                $global:DriverUpdateLastRunText.Foreground = [System.Windows.Media.Brushes]::Red
                if ($global:PatchingAlertIcon) { $global:PatchingAlertIcon.Visibility = "Visible" }
            } else {
                $global:DriverUpdateLastRunText.FontWeight = "Normal"
                $global:DriverUpdateLastRunText.Foreground = [System.Windows.Media.Brushes]::Gray
            }
        })
        
        Write-Log "Driver update status checked: $driverLastRun, Button visible: $showDriverButton" -Level "INFO"
        
    } catch {
        Write-Log "Error checking driver update status: $($_.Exception.Message)" -Level "ERROR"
    }
}

function Start-DriverUpdate {
    try {
        # Check battery power
        $powerStatus = [System.Windows.Forms.SystemInformation]::PowerStatus
        if ($powerStatus.PowerLineStatus -ne [System.Windows.Forms.PowerLineStatus]::Online) {
            [System.Windows.MessageBox]::Show(
                "Your computer is running on battery power.`n`nPlease connect to AC power before installing driver updates to prevent installation failures.",
                "AC Power Required",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            )
            return
        }
        
        # Final confirmation
        $result = [System.Windows.MessageBox]::Show(
            "This will install Windows driver updates. Your computer will automatically restart in 5 minutes after installation completes.`n`nSave your work before proceeding.`n`nContinue?",
            "Install Driver Updates",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        
        if ($result -eq [System.Windows.MessageBoxResult]::No) {
            return
        }
        
        Write-Log "User initiated driver update" -Level "INFO"
        
        # Show progress panel
        $window.Dispatcher.Invoke({
            $global:DriverProgressPanel.Visibility = "Visible"
            $global:DriverProgressStatus.Text = "Starting driver update process..."
            $global:DriverProgressStatus.Foreground = [System.Windows.Media.Brushes]::Black
            $global:DriverProgressBar.IsIndeterminate = $true
            $global:DriverUpdateButton.IsEnabled = $false
            $global:DriverProgressCloseButton.Visibility = "Collapsed"
        })
        
        # Start the scheduled task
        try {
            Start-ScheduledTask -TaskName "MITLL_DriverUpdate" -ErrorAction Stop
            Write-Log "Scheduled task 'MITLL_DriverUpdate' started successfully" -Level "INFO"
        } catch {
            throw "Failed to start scheduled task: $($_.Exception.Message)"
        }
        
        # Create timer for monitoring
        $script:monitorTimer = New-Object System.Windows.Threading.DispatcherTimer
        $script:monitorTimer.Interval = [TimeSpan]::FromSeconds(3)
        $script:monitorStartTime = Get-Date
        $script:lastLogUpdate = ""
        $script:monitorComplete = $false
        
        $script:monitorTimer.Add_Tick({
            try {
                $logFile = "C:\Windows\MITLL\Logs\MS_Update.txt"
                $elapsed = ((Get-Date) - $script:monitorStartTime).TotalSeconds
                
                # Timeout after 30 minutes
                if ($elapsed -gt 1800) {
                    $window.Dispatcher.Invoke({
                        $global:DriverProgressBar.IsIndeterminate = $false
                        $global:DriverProgressBar.Value = 100
                        $global:DriverProgressStatus.Text = "Update process timed out after 30 minutes. Check log file for details."
                        $global:DriverProgressStatus.Foreground = [System.Windows.Media.Brushes]::Orange
                        $global:DriverProgressCloseButton.Visibility = "Visible"
                        $global:DriverUpdateButton.IsEnabled = $true
                    })
                    Write-Log "Driver update monitoring timed out after 30 minutes" -Level "WARNING"
                    $script:monitorTimer.Stop()
                    $script:monitorComplete = $true
                    Update-DriverUpdateStatus
                    return
                }
                
                if (-not (Test-Path $logFile)) {
                    Write-Log "Log file not found: $logFile" -Level "WARNING"
                    return
                }
                
                # Read last 20 lines for better pattern matching
                $lastLines = Get-Content $logFile -Tail 20 -ErrorAction SilentlyContinue
                if (-not $lastLines) { return }
                
                $allContent = $lastLines -join "`n"
                
                # Check for completion states - ORDER MATTERS!
                
                # Check NO_UPDATES first before general completion
                if ($allContent -match "NO_UPDATES") {
                    $window.Dispatcher.Invoke({
                        $global:DriverProgressBar.IsIndeterminate = $false
                        $global:DriverProgressBar.Value = 100
                        $global:DriverProgressStatus.Text = "No driver updates are currently available. Your system is up to date."
                        $global:DriverProgressStatus.Foreground = [System.Windows.Media.Brushes]::Blue
                        $global:DriverProgressCloseButton.Visibility = "Visible"
                        $global:DriverUpdateButton.IsEnabled = $true
                    })
                    Write-Log "No driver updates available" -Level "INFO"
                    $script:monitorTimer.Stop()
                    $script:monitorComplete = $true
                    Update-DriverUpdateStatus
                    return
                }
                
                # Then check for successful installation completion
                if ($allContent -match "Driver Update Completed|Driver update process completed") {
                    $window.Dispatcher.Invoke({
                        $global:DriverProgressBar.IsIndeterminate = $false
                        $global:DriverProgressBar.Value = 100
                        $global:DriverProgressStatus.Text = "Driver updates installed successfully!"
                        $global:DriverProgressStatus.Foreground = [System.Windows.Media.Brushes]::Green
                        $global:DriverProgressCloseButton.Visibility = "Visible"
                        $global:DriverUpdateButton.IsEnabled = $true
                    })
                    Write-Log "Driver update completed successfully" -Level "INFO"
                    $script:monitorTimer.Stop()
                    $script:monitorComplete = $true
                    Update-DriverUpdateStatus
                    return
                }
                
                if ($allContent -match "REBOOT_SCHEDULED") {
                    $window.Dispatcher.Invoke({
                        $global:DriverProgressBar.IsIndeterminate = $false
                        $global:DriverProgressBar.Value = 100
                        $global:DriverProgressStatus.Text = "Installation complete. Your computer will restart in 5 minutes."
                        $global:DriverProgressStatus.Foreground = [System.Windows.Media.Brushes]::Orange
                        $global:DriverProgressCloseButton.Visibility = "Visible"
                        $global:DriverUpdateButton.IsEnabled = $true
                    })
                    Write-Log "Driver update complete - reboot scheduled" -Level "INFO"
                    $script:monitorTimer.Stop()
                    $script:monitorComplete = $true
                    Update-DriverUpdateStatus
                    return
                }
                
                # Update progress status based on current stage
                $statusText = "Installing driver updates. Please wait..."
                
                if ($allContent -match "SCAN_START") {
                    $statusText = "Scanning for available driver updates..."
                } elseif ($allContent -match "Checking for PSWindowsUpdate module") {
                    $statusText = "Preparing Windows Update module..."
                } elseif ($allContent -match "MODULE_IMPORT") {
                    $statusText = "Loading Windows Update module..."
                } elseif ($allContent -match "MODULE_READY") {
                    $statusText = "Module ready. Starting scan..."
                } elseif ($allContent -match "Downloading driver updates") {
                    $statusText = "Downloading driver updates..."
                } elseif ($allContent -match "Installing driver updates") {
                    $statusText = "Installing driver updates. This may take several minutes..."
                } elseif ($allContent -match "Checking BitLocker") {
                    $statusText = "Finalizing installation..."
                }
                
                # Only update if status changed
                if ($statusText -ne $script:lastLogUpdate) {
                    $script:lastLogUpdate = $statusText
                    $window.Dispatcher.Invoke({
                        $global:DriverProgressStatus.Text = $statusText
                        $global:DriverProgressStatus.Foreground = [System.Windows.Media.Brushes]::Black
                    })
                    Write-Log "Driver update progress: $statusText" -Level "INFO"
                }
                
            } catch {
                Write-Log "Error in driver update monitoring: $($_.Exception.Message)" -Level "ERROR"
            }
        })
        
        $script:monitorTimer.Start()
        Write-Log "Driver update monitoring started" -Level "INFO"
        
    } catch {
        Write-Log "Error starting driver update: $($_.Exception.Message)" -Level "ERROR"
        $window.Dispatcher.Invoke({
            $global:DriverProgressPanel.Visibility = "Visible"
            $global:DriverProgressStatus.Text = "Error: $($_.Exception.Message)"
            $global:DriverProgressStatus.Foreground = [System.Windows.Media.Brushes]::Red
            $global:DriverProgressBar.IsIndeterminate = $false
            $global:DriverUpdateButton.IsEnabled = $true
            $global:DriverProgressCloseButton.Visibility = "Visible"
        })
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
    
    # --- Update the global alert flags ---
    $global:CurrentUpdateState = "$bigfixStatusText`n$ecmStatusText"
    if ($global:CurrentUpdateState -ne $config.LastSeenUpdateState) {
        $global:UpdatesPending = $showBigFixButton -or $showEcmButton
    } else {
        $global:UpdatesPending = $false
    }
    
    # --- Update the UI ---
    $window.Dispatcher.Invoke({
        # REMOVED: Pending restart panel display has been disabled
        # The restart status panel will always be hidden
        $global:PendingRestartPanel.Visibility = "Collapsed"
        
        $global:WindowsBuildText.Text = $windowsBuild
        $global:FooterText.Text = "(C) 2025 Lincoln Laboratory v$ScriptVersion"
        
        # Update BigFix UI elements
        $global:BigFixStatusText.FontWeight = "Bold"
        $global:BigFixStatusText.Text = $bigfixStatusText
        $global:BigFixLaunchButton.Visibility = if ($showBigFixButton) { "Visible" } else { "Collapsed" }
        
        # Update ECM UI elements
        $global:ECMStatusText.FontWeight = "Bold"
        $global:ECMStatusText.Text = $ecmStatusText
        $global:ECMLaunchButton.Visibility = if ($showEcmButton) { "Visible" } else { "Collapsed" }

        # Conditionally show the alert dot for the whole section
        if ($global:PatchingAlertIcon -and (-not $global:PatchingExpander.IsExpanded)) {
            $global:PatchingAlertIcon.Visibility = if ($global:UpdatesPending) { "Visible" } else { "Collapsed" }
        }
    })
	Update-DriverUpdateStatus
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
        
        $regexColor = "\[(green|red|yellow|blue)\](.*?)\[/\1\]"
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
            $leftOk  = ($match.Index -ge 2) -and ($Text.Substring($match.Index - 2, 2) -eq "**")
			$rightOk = (($match.Index + $match.Length + 2) -le $Text.Length) -and ($Text.Substring($match.Index + $match.Length, 2) -eq "**")
			$isBold  = $leftOk -and $rightOk
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
    if (-not $global:contentData -or -not $global:contentData.Data) { Write-Log "No contentData available; skipping Announcements." -Level "WARNING"; return }
    $annObj = $global:contentData.Data.Announcements
    if (-not $annObj) { Write-Log "No Announcements object found; skipping." -Level "WARNING"; return }

    $result = Get-ActiveAnnouncement -AnnouncementsObject $annObj
    if ($null -eq $result) { Write-Log "Could not determine an active announcement. Section will not be updated." -Level "WARNING"; return }

    $newJsonState = Get-StableAnnouncementStateJson $result
    $shouldAlert  = Test-SectionChanged -SectionKey "Announcements" -NewStateJson $newJsonState

    $title   = if ($result.Base.Title)   { [string]$result.Base.Title }   else { '' }
    $textMd  = if ($result.Base.Text)    { [string]$result.Base.Text }    else { '' }
    $detailMd= if ($result.Base.Details) { [string]$result.Base.Details } else { '' }
    $links   = @(); if ($result.Base.Links) { $links = @($result.Base.Links) }

    Write-Log ("Announcements chosen: Title='{0}', Platform='{1}', Appended={2}" -f $title, $result.Base.Platform, $result.Appended.Count) -Level "INFO"

    $window.Dispatcher.Invoke({
        try {
            if ($shouldAlert) {
                if ($global:AnnouncementsAlertIcon) { $global:AnnouncementsAlertIcon.Visibility = "Visible" }
                if ($global:ClearAlertsDot)        { $global:ClearAlertsDot.Visibility        = "Visible" }
            }
            if ($global:AnnouncementsTitle)   { $global:AnnouncementsTitle.Text = $title }
            if ($global:AnnouncementsText)    {
                if ([string]::IsNullOrWhiteSpace($textMd)) { $global:AnnouncementsText.Text = '' }
                else { Convert-MarkdownToTextBlock -Text $textMd -TargetTextBlock $global:AnnouncementsText }
            }
            if ($global:AnnouncementsDetailsText) {
                if ([string]::IsNullOrWhiteSpace($detailMd)) { $global:AnnouncementsDetailsText.Text = '' }
                else { Convert-MarkdownToTextBlock -Text $detailMd -TargetTextBlock $global:AnnouncementsDetailsText }
            }
            if ($global:AnnouncementsLinksPanel) {
                $global:AnnouncementsLinksPanel.Children.Clear()
                if ($links.Count -gt 0) {
                    foreach ($link in $links) {
                        $global:AnnouncementsLinksPanel.Children.Add((New-HyperlinkBlock -Name $link.Name -Url $link.Url))
                    }
                }
            }
            if ($global:AnnouncementsSourceText) { $global:AnnouncementsSourceText.Text = "Source: $($global:contentData.Source)" }
            
            # Display appended targeted announcements
            if ($global:AppendedAnnouncementsPanel) {
                $global:AppendedAnnouncementsPanel.Children.Clear()
                
                if ($result.Appended -and $result.Appended.Count -gt 0) {
                    foreach ($appendedItem in $result.Appended) {
                        # Create separator
                        $separator = New-Object System.Windows.Controls.Border
                        $separator.BorderThickness = "0,1,0,0"
                        $separator.BorderBrush = "LightGray"
                        $separator.Margin = "0,10,0,10"
                        $global:AppendedAnnouncementsPanel.Children.Add($separator)
                        
                        # Title
                        if ($appendedItem.Title) {
                            $titleBlock = New-Object System.Windows.Controls.TextBlock
                            $titleBlock.Text = $appendedItem.Title
                            $titleBlock.FontWeight = "Bold"
                            $titleBlock.FontSize = 11
                            $titleBlock.Foreground = "#00008B"
                            $global:AppendedAnnouncementsPanel.Children.Add($titleBlock)
                        }
                        
                        # Text
                        if ($appendedItem.Text) {
                            $textBlock = New-Object System.Windows.Controls.TextBlock
                            $textBlock.FontSize = 11
                            $textBlock.TextWrapping = "Wrap"
                            $textBlock.Margin = "0,3,0,0"
                            Convert-MarkdownToTextBlock -Text $appendedItem.Text -TargetTextBlock $textBlock
                            $global:AppendedAnnouncementsPanel.Children.Add($textBlock)
                        }
                        
                        # Details
                        if ($appendedItem.Details) {
                            $detailsBlock = New-Object System.Windows.Controls.TextBlock
                            $detailsBlock.FontSize = 11
                            $detailsBlock.TextWrapping = "Wrap"
                            $detailsBlock.Margin = "0,3,0,0"
                            Convert-MarkdownToTextBlock -Text $appendedItem.Details -TargetTextBlock $detailsBlock
                            $global:AppendedAnnouncementsPanel.Children.Add($detailsBlock)
                        }
                        
                        # Links
                        if ($appendedItem.Links -and $appendedItem.Links.Count -gt 0) {
                            foreach ($link in $appendedItem.Links) {
                                $hyperlinkBlock = New-HyperlinkBlock -Name $link.Name -Url $link.Url
                                $hyperlinkBlock.Margin = "0,3,0,0"
                                $global:AppendedAnnouncementsPanel.Children.Add($hyperlinkBlock)
                            }
                        }
                    }
                    
                    $global:AppendedAnnouncementsPanel.Visibility = "Visible"
                    Write-Log "Displayed $($result.Appended.Count) appended announcement(s) - Panel visibility set to Visible" -Level "INFO"
                } else {
                    $global:AppendedAnnouncementsPanel.Visibility = "Collapsed"
                }
            } else {
                Write-Log "AppendedAnnouncementsPanel global variable is null!" -Level "ERROR"
            }
        } catch {
            Write-Log ("Announcements render failed: {0}" -f $_) -Level "ERROR"
        }
    })
    Save-SectionBaseline -SectionKey "Announcements" -NewStateJson $newJsonState
}

function Update-Support {
    Write-Log "Updating Support section..." -Level "INFO"
    if (-not $global:contentData -or -not $global:contentData.Data) { return }
    $supportData = $global:contentData.Data.Support
    if (-not $supportData) { return }

    $currentPlatform = "Windows"
    $newSupportObject = $null
    if ($supportData -is [System.Array]) {
        foreach ($supportItem in $supportData) {
            $p = $supportItem.Platform
            if (-not $p -or $p -eq "All" -or $p -eq $currentPlatform) { $newSupportObject = $supportItem; Write-Log "Using platform-specific support for: $currentPlatform" -Level "INFO"; break }
        }
        if (-not $newSupportObject -and $supportData.Count -gt 0) { $newSupportObject = $supportData[0]; Write-Log "No platform-specific support found, using first support entry" -Level "INFO" }
    } else {
        $newSupportObject = $supportData
        if ($newSupportObject.Platform -and $newSupportObject.Platform -ne "All" -and $newSupportObject.Platform -ne $currentPlatform) {
            Write-Log "Support section Platform mismatch (Expected: $($newSupportObject.Platform), Current: $currentPlatform)" -Level "WARNING"
        }
    }
    if (-not $newSupportObject) { return }

    Write-Log ("Support chosen: Platform='{0}', Links={1}" -f $newSupportObject.Platform, (@($newSupportObject.Links).Count)) -Level "INFO"

    $newJsonState = Get-StableSupportStateJson $newSupportObject
    Write-Log ("Support state SHA (new): {0}" -f (Get-TextSha256 $newJsonState)) -Level "INFO"
    $shouldAlert  = Test-SectionChanged -SectionKey "Support" -NewStateJson $newJsonState

    $window.Dispatcher.Invoke({
        try {
            if ($shouldAlert) {
                if ($global:SupportAlertIcon) { $global:SupportAlertIcon.Visibility = "Visible" }
                if ($global:ClearAlertsDot)   { $global:ClearAlertsDot.Visibility   = "Visible" }
            }
            if ($global:SupportText) {
                $txt = if ($newSupportObject.Text) { [string]$newSupportObject.Text } else { '' }
                if ([string]::IsNullOrWhiteSpace($txt)) { $global:SupportText.Text = '' }
                else { Convert-MarkdownToTextBlock -Text $txt -TargetTextBlock $global:SupportText }
            }
            if ($global:SupportDetailsText) {
                $details = if ($newSupportObject.Details) { [string]$newSupportObject.Details } else { '' }
                if ([string]::IsNullOrWhiteSpace($details)) { $global:SupportDetailsText.Text = '' }
                else { Convert-MarkdownToTextBlock -Text $details -TargetTextBlock $global:SupportDetailsText }
            }
            if ($global:SupportLinksPanel) {
                $global:SupportLinksPanel.Children.Clear()
                if ($newSupportObject.Links) { foreach ($link in $newSupportObject.Links) { $global:SupportLinksPanel.Children.Add((New-HyperlinkBlock -Name $link.Name -Url $link.Url)) } }
            }
            if ($global:SupportSourceText) { $global:SupportSourceText.Text = "Source: $($global:contentData.Source)" }
        } catch {
            Write-Log ("Support render failed: {0}" -f $_) -Level "ERROR"
        }
    })
    Save-SectionBaseline -SectionKey "Support" -NewStateJson $newJsonState
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
    $patchingAlert = $global:PatchingAlertIcon -and $global:PatchingAlertIcon.Visibility -eq "Visible"

    # REMOVED: Pending restart check has been disabled
    $hasAnyAlert = $announcementAlert -or $supportAlert -or $patchingAlert

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
        $global:TrayIcon.add_MouseClick({
            if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
                Toggle-WindowVisibility
            }
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
            Update-TrayIcon
        } else {
            # REMOVED: Pending restart acknowledgment has been disabled
            $window.Show()
            $global:BlinkingTimer.Stop()
            Update-TrayIcon
            
            # Use the .Activate() method for robust foregrounding
            $window.Activate()
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
        $title = "Endpoint Advisor is Running"
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
    Write-Log '--- Lincoln Laboratory Endpoint Advisor Script Exiting ---'
    if ($global:DispatcherTimer) { $global:DispatcherTimer.Stop() }
    if ($global:TrayIcon) { $global:TrayIcon.Dispose() }
    if ($global:MainIcon) { $global:MainIcon.Dispose() }
    if ($global:WarningIcon) { $global:WarningIcon.Dispose() }

    try {
        foreach ($run in $runs) {
            $TargetTextBlock.Inlines.Add($run)
            Write-Log ("Added run to TextBlock: {0} (FontWeight: {1}, FontStyle: {2}, Foreground: {3})" -f $run.Text,$run.FontWeight,$run.FontStyle,$run.Foreground) -Level "INFO"
        }
        Write-Log ("Successfully parsed Markdown for text: {0}" -f $Text) -Level "INFO"
    }
    catch {
        Write-Log ("Failed to parse Markdown for text: {0} - {1}" -f $Text, $_.Exception.Message) -Level "ERROR"
        if ($TargetTextBlock -and $TargetTextBlock.Inlines) {
            $TargetTextBlock.Inlines.Clear()
            $TargetTextBlock.Inlines.Add((New-Object System.Windows.Documents.Run(($Text -as [string]))))
        }
    }
}
