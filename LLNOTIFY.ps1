# LLNOTIFY.ps1 - Lincoln Laboratory Notification System
# Version 4.6.13 (Updated to use BESClientComplianceMod.Session per BigFix Client Compliance API, retained qna.exe fallback, fixed SSA log file issue, restricted site info to reports)

# Detect if running in 64-bit PowerShell and relaunch in 32-bit if necessary
if ([Environment]::Is64BitProcess) {
    Write-Host "Relaunching in 32-bit PowerShell for BigFix COM compatibility..."
    $scriptPath = $PSCommandPath
    $32bitPS = "$env:windir\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    Start-Process -FilePath $32bitPS -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Wait
    exit
}

# Ensure $PSScriptRoot is defined for older versions
if ($MyInvocation.MyCommand.Path) {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    $ScriptDir = Get-Location
}

# Define version
$ScriptVersion = "4.6.13"

# Global flag to prevent recursive logging during rotation
$global:IsRotatingLog = $false

# Global flag to track pending restart state
$global:PendingRestart = $false

# Global variables for certificate check caching
$global:LastCertificateCheck = $null
$global:CachedCertificateStatus = $null

# ============================================================
# A) Advanced Logging & Error Handling
# ============================================================
function Invoke-WithRetry {
    param(
        [ScriptBlock]$Action,
        [int]$MaxRetries = 3,
        [int]$RetryDelayMs = 100
    )
    $attempt = 0
    while ($attempt -lt $MaxRetries) {
        try {
            $attempt++
            $Action.Invoke()
            return # Success
        }
        catch {
            if ($attempt -ge $MaxRetries) {
                throw "Action failed after $MaxRetries attempts: $($_.Exception.Message)"
            }
            Start-Sleep -Milliseconds $RetryDelayMs
        }
    }
}

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
    
    while ($attempt -lt $MaxRetries -and -not $success) {
        try {
            $attempt++
            Add-Content -Path $logPath -Value $logEntry -Force -ErrorAction Stop
            $success = $true
        }
        catch {
            if ($attempt -eq $MaxRetries) {
                Write-Host "[$timestamp] [$Level] $Message (Failed to write to log after $maxRetries attempts: $($_.Exception.Message))"
            } else {
                Start-Sleep -Milliseconds $retryDelayMs
            }
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
                                Write-Log "Failed to delete old archive $($file.FullName): $($_.Exception.Message)" -Level "ERROR"
                            }
                        }
                    }
                }
                catch {
                    Write-Log "Failed to rotate log file: $($_.Exception.Message)" -Level "ERROR"
                }
            }
        }
    }
    catch {
        Write-Log "Error checking log file size for rotation: $($_.Exception.Message)" -Level "ERROR"
    }
    finally {
        $global:IsRotatingLog = $false
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

Write-Log "--- LLNOTIFY Script Started (Version $ScriptVersion) ---"

# ============================================================
# BigFix Compliance Reporting Functions
# ============================================================
function Register-BigFixCOMObject {
    try {
        Write-Log "Checking and registering BigFix COM object ($($config.BigFixCOM_ProgID))..." -Level "INFO"
        $dllPaths = $config.BigFixDLL_Paths
        if (-not $dllPaths) {
            $dllPaths = @($config.BigFixDLL_Path)
        }

        # Check if BigFix client service is running
        $service = Get-Service -Name "BESClient" -ErrorAction SilentlyContinue
        if (-not $service -or $service.Status -ne "Running") {
            Write-Log "BigFix client service (BESClient) is not running or not installed. COM registration may fail." -Level "WARNING"
        } else {
            Write-Log "BigFix client service (BESClient) is running." -Level "INFO"
        }

        # Check registry for BigFix-related ProgIDs
        $bigFixProgIDs = Get-ChildItem HKCR:\ -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*BES*" -or $_.Name -like "*BigFix*" } | Select-Object -ExpandProperty Name
        if ($bigFixProgIDs) {
            Write-Log "Found BigFix-related ProgIDs in registry: $($bigFixProgIDs -join ', ')" -Level "INFO"
        } else {
            Write-Log "No BigFix-related ProgIDs found in registry." -Level "WARNING"
        }

        # Attempt to create the COM object with configured ProgID
        try {
            $comObject = New-Object -ComObject $config.BigFixCOM_ProgID -ErrorAction Stop
            Write-Log "$($config.BigFixCOM_ProgID) COM object already registered." -Level "INFO"
            $clsid = (Get-ItemProperty -Path "HKCR\$($config.BigFixCOM_ProgID)\CLSID" -ErrorAction SilentlyContinue)."(default)"
            if ($clsid) { Write-Log "CLSID for $($config.BigFixCOM_ProgID): $clsid" -Level "INFO" }
            return $true
        }
        catch [System.Runtime.InteropServices.COMException] {
            if ($_.Exception.HResult -eq 0x80040154) {
                Write-Log "$($config.BigFixCOM_ProgID) COM object not registered. Attempting fallback ProgIDs..." -Level "INFO"
                
                # Try fallback ProgIDs
                $fallbackProgIDs = @("BESClientComplianceMod.Session", "BESClient", "BigFix.Client")
                foreach ($progID in $fallbackProgIDs) {
                    try {
                        $comObject = New-Object -ComObject $progID -ErrorAction Stop
                        Write-Log "Fallback ProgID $progID succeeded. Updating config..." -Level "INFO"
                        $config.BigFixCOM_ProgID = $progID
                        Save-Configuration -Config $config
                        $clsid = (Get-ItemProperty -Path "HKCR\$progID\CLSID" -ErrorAction SilentlyContinue)."(default)"
                        if ($clsid) { Write-Log "CLSID for $progID: $clsid" -Level "INFO" }
                        return $true
                    }
                    catch {
                        Write-Log "Fallback ProgID $progID failed: $($_.Exception.Message)" -Level "WARNING"
                    }
                }

                # Attempt to register each DLL
                $regsvr32Path = "C:\Windows\SysWOW64\regsvr32.exe"
                foreach ($dllPath in $dllPaths) {
                    if (-not (Test-Path $dllPath)) {
                        Write-Log "BigFix COM DLL not found at: $dllPath" -Level "WARNING"
                        continue
                    }

                    # Check DLL dependencies
                    try {
                        [System.Reflection.Assembly]::LoadFile($dllPath) | Out-Null
                        Write-Log "DLL dependencies for $dllPath appear to be satisfied." -Level "INFO"
                    }
                    catch {
                        Write-Log "Dependency check failed for $dllPath: $($_.Exception.Message)" -Level "WARNING"
                    }

                    Write-Log "Attempting to unregister and re-register with 32-bit Regsvr32: $dllPath" -Level "INFO"
                    try {
                        # Unregister first to clean up
                        $tempUnregOutputFile = [System.IO.Path]::GetTempFileName()
                        $tempUnregErrorFile = [System.IO.Path]::GetTempFileName()
                        $unregResult = Start-Process -FilePath $regsvr32Path -ArgumentList "/u /s `"$dllPath`"" -Wait -NoNewWindow -RedirectStandardOutput $tempUnregOutputFile -RedirectStandardError $tempUnregErrorFile -PassThru
                        $unregOutput = Get-Content -Path $tempUnregOutputFile -Raw
                        $unregError = Get-Content -Path $tempUnregErrorFile -Raw
                        Remove-Item -Path $tempUnregOutputFile -Force -ErrorAction SilentlyContinue
                        Remove-Item -Path $tempUnregErrorFile -Force -ErrorAction SilentlyContinue
                        
                        $unregCombined = ""
                        if ($unregOutput) { $unregCombined += "StdOut: $unregOutput`n" }
                        if ($unregError) { $unregCombined += "StdErr: $unregError" }
                        Write-Log "Unregistration result for $dllPath (exit code: $($unregResult.ExitCode)): $unregCombined" -Level "INFO"

                        # Register
                        $tempRegOutputFile = [System.IO.Path]::GetTempFileName()
                        $tempRegErrorFile = [System.IO.Path]::GetTempFileName()
                        $regResult = Start-Process -FilePath $regsvr32Path -ArgumentList "/s `"$dllPath`"" -Wait -NoNewWindow -RedirectStandardOutput $tempRegOutputFile -RedirectStandardError $tempRegErrorFile -PassThru
                        $regOutput = Get-Content -Path $tempRegOutputFile -Raw
                        $regError = Get-Content -Path $tempRegErrorFile -Raw
                        Remove-Item -Path $tempRegOutputFile -Force -ErrorAction SilentlyContinue
                        Remove-Item -Path $tempRegErrorFile -Force -ErrorAction SilentlyContinue
                        
                        $regCombined = ""
                        if ($regOutput) { $regCombined += "StdOut: $regOutput`n" }
                        if ($regError) { $regCombined += "StdErr: $regError" }
                        
                        if ($regResult.ExitCode -eq 0) {
                            Write-Log "32-bit Regsvr32 reported success for: $dllPath. Verifying COM object..." -Level "INFO"
                            # Verify COM object with all ProgIDs
                            $allProgIDs = @($config.BigFixCOM_ProgID) + $fallbackProgIDs
                            foreach ($progID in $allProgIDs) {
                                try {
                                    $comObject = New-Object -ComObject $progID -ErrorAction Stop
                                    Write-Log "COM object $progID successfully verified after registration." -Level "INFO"
                                    $config.BigFixCOM_ProgID = $progID
                                    Save-Configuration -Config $config
                                    $clsid = (Get-ItemProperty -Path "HKCR\$progID\CLSID" -ErrorAction SilentlyContinue)."(default)"
                                    if ($clsid) { Write-Log "Post-registration CLSID for $progID: $clsid" -Level "INFO" }
                                    return $true
                                }
                                catch {
                                    Write-Log "Verification failed for $progID: $($_.Exception.Message)" -Level "WARNING"
                                }
                            }
                            Write-Log "COM object registration verification failed for all ProgIDs: $($allProgIDs -join ', ')" -Level "WARNING"
                        } else {
                            Write-Log "32-bit Regsvr32 failed to register $dllPath with exit code: $($regResult.ExitCode). Output: $regCombined" -Level "ERROR"
                        }
                    }
                    catch {
                        Write-Log "Failed to register BigFix COM DLL $dllPath: $($_.Exception.Message). Ensure the script is run as administrator and dependencies are installed." -Level "ERROR"
                    }
                }
                return $false
            } else {
                Write-Log "Unexpected COM error when testing $($config.BigFixCOM_ProgID): $($_.Exception.Message)" -Level "ERROR"
                return $false
            }
        }
    }
    catch {
        Write-Log "Error checking or registering BigFix COM object: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

function Initialize-BigFixAPIConfig {
    try {
        Write-Log "Checking and initializing BigFix Client Compliance API configuration..." -Level "INFO"
        
        # Only attempt COM registration if enabled
        if ($config.UseBigFixCOM) {
            if (-not (Register-BigFixCOMObject)) {
                Write-Log "BigFix COM object registration failed. Falling back to qna.exe if available." -Level "WARNING"
            }
        } else {
            Write-Log "BigFix COM usage disabled in config. Using qna.exe for relevance queries." -Level "INFO"
        }

        # Registry key for 64-bit Windows
        $regKeyPath = "HKLM:\SOFTWARE\WOW6432Node\BigFix\ClientComplianceAPI"
        $bigFixBasePath = "C:\Program Files (x86)\BigFix Enterprise\BES Client"
        $requestDir = Join-Path $bigFixBasePath "RequestDir"
        $responseDir = Join-Path $bigFixBasePath "ResponseDir"
        $connectDir = Join-Path $bigFixBasePath "ConnectDir"

        # Check and create registry key
        if (-not (Test-Path $regKeyPath)) {
            Write-Log "Creating BigFix API registry key: $regKeyPath" -Level "INFO"
            New-Item -Path $regKeyPath -Force | Out-Null
        }

        # Check and set registry values
        $regValues = @{
            "RequestDir" = $requestDir
            "ResponseDir" = $responseDir
            "ConnectDir" = $connectDir
        }
        foreach ($key in $regValues.Keys) {
            $currentValue = (Get-ItemProperty -Path $regKeyPath -Name $key -ErrorAction SilentlyContinue).$key
            if (-not $currentValue) {
                Write-Log "Setting registry value $key to $($regValues[$key])" -Level "INFO"
                Set-ItemProperty -Path $regKeyPath -Name $key -Value $regValues[$key] -Force
            } elseif ($currentValue -ne $regValues[$key]) {
                Write-Log "Updating registry value $key from $currentValue to $($regValues[$key])" -Level "INFO"
                Set-ItemProperty -Path $regKeyPath -Name $key -Value $regValues[$key] -Force
            }
        }

        # Check and create directories
        foreach ($dir in @($requestDir, $responseDir, $connectDir)) {
            if (-not (Test-Path $dir)) {
                Write-Log "Creating directory: $dir" -Level "INFO"
                New-Item -Path $dir -ItemType Directory -Force | Out-Null
            }
        }

        # Set directory permissions (requires admin privileges)
        try {
            # RequestDir: Everyone can write
            icacls $requestDir /grant "Everyone:(OI)(CI)(W)" /T | Out-Null
            Write-Log "Set write permissions for Everyone on $requestDir" -Level "INFO"
            
            # ResponseDir: Everyone can read
            icacls $responseDir /grant "Everyone:(OI)(CI)(R)" /T | Out-Null
            Write-Log "Set read permissions for Everyone on $responseDir" -Level "INFO"
            
            # ConnectDir: Everyone can execute
            icacls $connectDir /grant "Everyone:(OI)(CI)(RX)" /T | Out-Null
            Write-Log "Set execute permissions for Everyone on $connectDir" -Level "INFO"
        }
        catch {
            Write-Log "Failed to set permissions on BigFix directories: $($_.Exception.Message). Script may require admin privileges." -Level "WARNING"
        }

        Write-Log "BigFix API configuration initialized successfully." -Level "INFO"
        return $true
    }
    catch {
        Write-Log "Error initializing BigFix API config: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

function Get-BigFixRelevanceResult {
    param(
        [string]$RelevanceQuery
    )
    try {
        if ($config.UseBigFixCOM) {
            # Try COM object with retries
            $maxRetries = 3
            $retryDelayMs = 1000
            $attempt = 0
            while ($attempt -lt $maxRetries) {
                try {
                    $session = New-Object -ComObject $config.BigFixCOM_ProgID -ErrorAction Stop
                    $session.Open()
                    $session.Expression = $RelevanceQuery
                    if ($session.Evaluate()) {
                        $result = $session.Result
                        Write-Log "BigFix relevance query succeeded via COM ($($config.BigFixCOM_ProgID)): $RelevanceQuery. Result: $result" -Level "INFO"
                        $session.Close()
                        return $result
                    } else {
                        Write-Log "BigFix relevance query evaluation failed via COM ($($config.BigFixCOM_ProgID)): $RelevanceQuery" -Level "WARNING"
                        $session.Close()
                        return "Evaluation failed"
                    }
                }
                catch [System.Runtime.InteropServices.COMException] {
                    if ($_.Exception.HResult -eq 0x80040154) {
                        $attempt++
                        if ($attempt -eq $maxRetries) {
                            Write-Log "BigFix relevance query failed after $maxRetries attempts: $($config.BigFixCOM_ProgID) COM object not registered (REGDB_E_CLASSNOTREG). Falling back to qna.exe." -Level "ERROR"
                            break
                        }
                        Write-Log "COM object not ready (attempt $attempt/$maxRetries). Retrying in $retryDelayMs ms..." -Level "WARNING"
                        Start-Sleep -Milliseconds $retryDelayMs
                    } else {
                        Write-Log "BigFix relevance query error via COM ($($config.BigFixCOM_ProgID)): $($_.Exception.Message)" -Level "ERROR"
                        return "Error: $($_.Exception.Message)"
                    }
                }
                finally {
                    if ($session) {
                        try { $session.Close() } catch {}
                    }
                }
            }
        } else {
            Write-Log "COM usage disabled. Using qna.exe for relevance query: $RelevanceQuery" -Level "INFO"
        }

        # Fallback to qna.exe
        $qnaPath = $config.BigFixQna_Path
        if (-not (Test-Path $qnaPath)) {
            Write-Log "qna.exe not found at: $qnaPath. Cannot execute relevance query." -Level "ERROR"
            return "Error: BigFix API not available. Contact IT."
        }

        try {
            Write-Log "Executing relevance query with qna.exe: $RelevanceQuery" -Level "INFO"
            $tempQueryFile = [System.IO.Path]::GetTempFileName()
            "q: $RelevanceQuery" | Out-File -FilePath $tempQueryFile -Encoding ASCII
            $qnaResult = & $qnaPath $tempQueryFile
            Remove-Item -Path $tempQueryFile -Force -ErrorAction SilentlyContinue
            if ($qnaResult) {
                $result = $qnaResult -join "`n"
                Write-Log "qna.exe relevance query succeeded: $RelevanceQuery. Result: $result" -Level "INFO"
                return $result
            } else {
                Write-Log "qna.exe relevance query returned no results: $RelevanceQuery" -Level "WARNING"
                return "Evaluation failed"
            }
        }
        catch {
            Write-Log "qna.exe relevance query error: $($_.Exception.Message)" -Level "ERROR"
            return "Error: $($_.Exception.Message)"
        }
    }
    catch {
        Write-Log "BigFix relevance query error: $($_.Exception.Message)" -Level "ERROR"
        return "Error: $($_.Exception.Message)"
    }
}

function Generate-BigFixComplianceReport {
    try {
        Write-Log "Gathering BigFix compliance info..." -Level "INFO"
        $reportPath = Join-Path $ScriptDir "BigFixComplianceReport.txt"
        $jsonPath = Join-Path $ScriptDir "BigFixComplianceReport.json"

        $computerName = Get-BigFixRelevanceResult "name of computer"
        $clientVersion = Get-BigFixRelevanceResult "version of client as string"
        $relay = Get-BigFixRelevanceResult "if exists relay service then (address of relay service as string) else `"No Relay`""
        $lastReport = Get-BigFixRelevanceResult "last report time of client as string"
        $ipAddress = Get-BigFixRelevanceResult "ip address of client as string"
        $siteList = Get-BigFixRelevanceResult "names of sites whose (subscribed of it = true)"
        $fixletList = Get-BigFixRelevanceResult "names of relevant fixlets whose (baseline flag of it = false and (name of it as lowercase contains `"microsoft`" or name of it as lowercase contains `"security update`")) of sites"

        # Check if all queries failed
        if ($computerName -like "Error:*" -and $clientVersion -like "Error:*" -and $relay -like "Error:*" -and 
            $lastReport -like "Error:*" -and $ipAddress -like "Error:*" -and $siteList -like "Error:*" -and $fixletList -like "Error:*") {
            $report = @(
                "BigFix Compliance Report - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
                "------------------------------------------------------------",
                "Error: BigFix API not available. Contact IT for assistance."
            )
            $report | Out-File -FilePath $reportPath -Encoding UTF8
            $reportJson = @{
                Timestamp = (Get-Date)
                Error = "BigFix API not available. Contact IT for assistance."
            }
            $reportJson | ConvertTo-Json -Depth 3 | Out-File -FilePath $jsonPath -Encoding UTF8
            Write-Log "BigFix compliance report failed due to API unavailability." -Level "ERROR"
            return
        }

        $sites = @()
        if ($siteList -is [string] -and -not [string]::IsNullOrWhiteSpace($siteList) -and $siteList -notlike "Error:*") {
            $sites = $siteList -split "`n" | Where-Object { $_ -match "\S" } # Filter empty lines
        }

        $fixlets = @()
        if ($fixletList -is [string] -and -not [string]::IsNullOrWhiteSpace($fixletList) -and $fixletList -notlike "Error:*") {
            $fixlets = $fixletList -split "`n" | Where-Object { $_ -match "\S" } # Filter empty lines
        }

        $report = @(
            "BigFix Compliance Report - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
            "------------------------------------------------------------",
            "Computer Name  : $computerName",
            "Client Version : $clientVersion",
            "Relay Address  : $relay",
            "IP Address     : $ipAddress",
            "Last Reported  : $lastReport",
            "",
            "Subscribed Sites:",
            "-----------------"
        )
        if ($sites.Count -gt 0) {
            $report += $sites | ForEach-Object { " - $_" }
        } else {
            $report += "No subscribed sites found."
        }
        $report += @("", "Applicable Fixlets (Non-Baselines):", "----------------------------------")
        if ($fixlets.Count -gt 0) {
            $report += $fixlets | ForEach-Object { " - $_" }
        } else {
            $report += "No applicable fixlets found."
        }

        $report | Out-File -FilePath $reportPath -Encoding UTF8
        $reportJson = @{
            Timestamp         = (Get-Date)
            ComputerName      = $computerName
            ClientVersion     = $clientVersion
            Relay             = $relay
            IPAddress         = $ipAddress
            LastReportTime    = $lastReport
            SubscribedSites   = $sites
            ApplicableFixlets = $fixlets
        }
        $reportJson | ConvertTo-Json -Depth 3 | Out-File -FilePath $jsonPath -Encoding UTF8
        Write-Log "BigFix compliance report written to $reportPath and $jsonPath" -Level "INFO"
    } catch {
        Write-Log "Error generating BigFix compliance report: $($_.Exception.Message)" -Level "ERROR"
    }
}

# ============================================================
# MODULE: Configuration Management
# ============================================================
function Get-DefaultConfig {
    return @{
        RefreshInterval       = 900  # Default to 15 minutes (900 seconds)
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
        BigFixDLL_Path        = "C:\Program Files (x86)\BigFix Enterprise\BES Client\BESClientComplianceMOD.dll"
        BigFixDLL_Paths       = @(
            "C:\Program Files (x86)\BigFix Enterprise\BES Client\BESClientComplianceMOD.dll",
            "C:\Program Files (x86)\BigFix Enterprise\BES Client\BESClient.dll"
        )
        BigFixCOM_ProgID      = "BESClientComplianceMod.Session"
        UseBigFixCOM          = $true
        BigFixQna_Path        = "C:\Program Files (x86)\BigFix Enterprise\BES Client\qna.exe"
        YubiKeyManager_Path   = "C:\Program Files\Yubico\Yubikey Manager\ykman.exe"
        BlinkingEnabled       = $true
        ScriptUrl             = "https://raw.githubusercontent.com/burnoil/LLNOTIFY/refs/heads/main/LLNOTIFY.ps1"
        VersionUrl            = "https://raw.githubusercontent.com/burnoil/LLNOTIFY/refs/heads/main/currentversion.txt"
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
            Write-Log "Failed to load or merge existing config file. Reverting to full defaults. Error: $($_.Exception.Message)" -Level "WARNING"
        }
    }
    try {
        $finalConfig | ConvertTo-Json -Depth 100 | Out-File $Path -Force
        Write-Log "Configuration file validated and saved." -Level "INFO"
    }
    catch {
        Handle-Error "Could not save the updated configuration to '$Path'. Error: $($_.Exception.Message)"
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
    } catch {
        Handle-Error "Could not save state to configuration file '$Path'. Error: $($_.Exception.Message)"
    }
}

# ============================================================
# B) External Configuration Setup
# ============================================================
$LogFilePath = Join-Path $ScriptDir "LLNOTIFY.log"
$config = Load-Configuration

$mainIconPath = $config.IconPaths.Main
$warningIconPath = $config.IconPaths.Warning
$mainIconUri = "file:///$($mainIconPath -replace '\\','/')"

Write-Log "Main icon path: $mainIconPath" -Level "INFO"
Write-Log "Warning icon path: $warningIconPath" -Level "INFO"

$defaultContentData = @{
    Announcements = @{ Text = "No announcements at this time."; Details = ""; Links = @() }
    Support = @{ Text  = "Contact IT Support."; Links = @() }
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
    } catch {
        Write-Log "Error logging .NET version: $($_.Exception.Message)" -Level "ERROR"
    }
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
        Write-Log "Failed to load required GUI assemblies: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

$global:FormsAvailable = Import-RequiredAssemblies

# ============================================================
# E) XAML Layout Definition
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
    ResizeMode="CanResizeWithGrip" ShowInTaskbar="False" Visibility="Hidden" Topmost="True"
    Background="#f0f0f0"
    Icon="{Binding WindowIconUri}">
  <Grid Margin="5">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <Border Grid.Row="0" Background="#0078D7" Padding="5" CornerRadius="3" Margin="0,0,0,5">
      <StackPanel Orientation="Horizontal" VerticalAlignment="Center" HorizontalAlignment="Center">
        <Image Source="{Binding MainIconUri}" Width="20" Height="20" Margin="0,0,5,0"/>
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
        <Expander x:Name="PatchingExpander" FontSize="12" IsExpanded="True" Margin="0,2,0,2">
          <Expander.Header>
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="Patching and Updates" VerticalAlignment="Center"/>
              <Button x:Name="PatchingSSAButton" Content="Launch Updates" Margin="10,0,0,0" ToolTip="Launch BigFix Self-Service Application"/>
            </StackPanel>
          </Expander.Header>
          <Border BorderBrush="#00008B" BorderThickness="1" Padding="5" CornerRadius="3" Background="White" Margin="2">
            <StackPanel>
              <TextBlock x:Name="PatchingDescriptionText" FontSize="11" TextWrapping="Wrap"/>
              <TextBlock Text="Pending Restart Status:" FontSize="11" FontWeight="Bold" Margin="0,5,0,0"/>
              <TextBlock x:Name="PendingRestartStatusText" FontSize="11" FontWeight="Bold" TextWrapping="Wrap"/>
              <TextBlock x:Name="PatchingUpdatesText" FontSize="11" TextWrapping="Wrap" Margin="0,5,0,0"/>
            </StackPanel>
          </Border>
        </Expander>
        <Expander x:Name="SupportExpander" FontSize="12" IsExpanded="False" Margin="0,2,0,2">
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
        <Expander x:Name="ComplianceExpander" Header="Certificate Status" FontSize="12" IsExpanded="False" Margin="0,2,0,2">
          <Border BorderBrush="#00008B" BorderThickness="1" Padding="5" CornerRadius="3" Background="White" Margin="2">
            <TextBlock x:Name="YubiKeyComplianceText" FontSize="11" TextWrapping="Wrap"/>
          </Border>
        </Expander>
        <TextBlock x:Name="WindowsBuildText" FontSize="11" TextWrapping="Wrap" HorizontalAlignment="Center" Margin="0,10,0,0"/>
        <TextBlock x:Name="ScriptUpdateText" FontSize="11" TextWrapping="Wrap" HorizontalAlignment="Center" Margin="0,10,0,0" Foreground="Red" Visibility="Hidden"/>
      </StackPanel>
    </ScrollViewer>
    <Grid Grid.Row="2" Margin="0,5,0,0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*" />
        <ColumnDefinition Width="Auto" />
      </Grid.ColumnDefinitions>
      <TextBlock x:Name="FooterText" Grid.Column="0" Text="© 2025 Lincoln Laboratory" FontSize="10" Foreground="Gray" HorizontalAlignment="Center" VerticalAlignment="Center"/>
      <Button x:Name="ClearAlertsButton" Grid.Column="1" Content="Clear Alerts" FontSize="10" Padding="5,1" Background="#B0C4DE" ToolTip="Acknowledge all new announcements and support messages."/>
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

    $window.DataContext = [PSCustomObject]@{ 
        MainIconUri   = [Uri]$mainIconUri
        WindowIconUri = $mainIconPath
    }

    $uiElements = @(
        "AnnouncementsExpander", "AnnouncementsAlertIcon", "AnnouncementsText", "AnnouncementsDetailsText",
        "AnnouncementsLinksPanel", "AnnouncementsSourceText", "PatchingExpander", "PatchingDescriptionText",
        "PendingRestartStatusText", "PatchingUpdatesText", "PatchingSSAButton", "SupportExpander",
        "SupportAlertIcon", "SupportText", "SupportLinksPanel", "SupportSourceText", "ComplianceExpander",
        "YubiKeyComplianceText", "WindowsBuildText", "ClearAlertsButton", "ScriptUpdateText", "FooterText"
    )
    foreach ($elementName in $uiElements) {
        Set-Variable -Name "global:$elementName" -Value $window.FindName($elementName)
    }
    Write-Log "UI elements mapped to variables." -Level "INFO"

    $global:FooterText.Text = "$([char]169) 2025 Lincoln Laboratory v$ScriptVersion"

    $global:AnnouncementsExpander.Add_Expanded({ $window.Dispatcher.Invoke({ $global:AnnouncementsAlertIcon.Visibility = "Hidden"; Update-TrayIcon }) })
    $global:SupportExpander.Add_Expanded({ $window.Dispatcher.Invoke({ $global:SupportAlertIcon.Visibility = "Hidden"; Update-TrayIcon }) })

    $global:PatchingSSAButton.Add_Click({
        try {
            $ssaPath = $config.BigFixSSA_Path
            if ([string]::IsNullOrWhiteSpace($ssaPath) -or -not (Test-Path $ssaPath)) {
                throw "BigFix Self-Service Application path is invalid or not found: `"$ssaPath`""
            }

            # Ensure BigFix SSA log directory and file exist
            $ssaLogDir = Join-Path $env:LOCALAPPDATA "BigFix\BigFixSSA\logs"
            $today = Get-Date -Format "yyyyMMdd"
            $ssaLogFile = Join-Path $ssaLogDir "main_$today.log"
            if (-not (Test-Path $ssaLogDir)) {
                Write-Log "Creating BigFix SSA log directory: $ssaLogDir" -Level "INFO"
                New-Item -Path $ssaLogDir -ItemType Directory -Force | Out-Null
                if (-not (Test-Path $ssaLogDir)) {
                    throw "Failed to create BigFix SSA log directory: $ssaLogDir"
                }
            }
            if (-not (Test-Path $ssaLogFile)) {
                Write-Log "Creating BigFix SSA log file: $ssaLogFile" -Level "INFO"
                "Initialized at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Out-File -FilePath $ssaLogFile -Encoding UTF8
                if (-not (Test-Path $ssaLogFile)) {
                    throw "Failed to create BigFix SSA log file: $ssaLogFile"
                }
            }

            Write-Log "Launching BigFix SSA: $ssaPath" -Level "INFO"
            Start-Process -FilePath $ssaPath
        }
        catch {
            Handle-Error $_.Exception.Message -Source "PatchingSSAButton"
        }
    })
    
    $global:ClearAlertsButton.Add_Click({
        Write-Log "Clear Alerts button clicked by user." -Level "INFO"

        if ($global:contentData) {
            $config.AnnouncementsLastState = $global:contentData.Data.Announcements | ConvertTo-Json -Compress
            $config.SupportLastState = $global:contentData.Data.Support | ConvertTo-Json -Compress
        }

        $window.Dispatcher.Invoke({
            $global:AnnouncementsAlertIcon.Visibility = 'Hidden'
            $global:SupportAlertIcon.Visibility = 'Hidden'
        })

        $global:BlinkingTimer.Stop()
        Update-TrayIcon
        
        Save-Configuration -Config $config
    })

    $window.Add_Closing({
        if (-not $global:IsUpdating) {
            $_.Cancel = $true
            $window.Hide()
        }
    })

    # Initialize BigFix API configuration
    Initialize-BigFixAPIConfig
}
catch {
    Handle-Error "Failed to load the XAML layout or initialize BigFix API: $($_.Exception.Message)" -Source "XAML"
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

function Fetch-ContentData {
    # Guard: ensure ContentDataUrl is set
    if (-not $config -or [string]::IsNullOrWhiteSpace($config.ContentDataUrl)) {
        Write-Log "ContentDataUrl is not set! Check your Get-DefaultConfig return value." -Level "ERROR"
        return [PSCustomObject]@{ Data = $defaultContentData; Source = "Default" }
    }
    $url = $config.ContentDataUrl

    try {
        Write-Log "Attempting to fetch content from: $url" -Level "INFO"
        
        # Async fetch using job to avoid blocking
        $job = Start-Job -ScriptBlock {
            param($url)
            Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop
        } -ArgumentList $url
        $response = Wait-Job $job | Receive-Job
        Remove-Job $job
        Write-Log "Successfully fetched content from Git repository (Status: $($response.StatusCode))." -Level "INFO"
        
        $contentData = $response.Content | ConvertFrom-Json
        Validate-ContentData -Data $contentData
        Write-Log "Content data validated successfully." -Level "INFO"

        return [PSCustomObject]@{ Data = $contentData; Source = "Remote" }
    }
    catch {
        Write-Log "Failed to fetch or validate content from $($config.ContentDataUrl): $($_.Exception.Message)" -Level "ERROR"
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
            return "YubiKey not present"
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
        return "YubiKey Certificate: No PIV certificate found."
    }
    catch {
        Write-Log "YubiKey check error: $($_.Exception.Message)" -Level "ERROR"
        return "YubiKey Certificate: Unable to determine status."
    }
}

function Get-VirtualSmartCardCertExpiry {
    try {
        $cert = Get-ChildItem "Cert:\CurrentUser\My" | Where-Object { $_.Subject -match "Virtual" } | Sort-Object NotAfter -Descending | Select-Object -First 1
        if (-not $cert) { return "No certificate found." }
        return "Microsoft Virtual Smart Card: Expires: $($cert.NotAfter.ToString("yyyy-MM-dd"))"
    } catch {
        Write-Log "Virtual Smart Card check error: $($_.Exception.Message)" -Level "ERROR"
        return "Microsoft Virtual Smart Card: Unable to check status."
    }
}

function Update-CertificateInfo {
    try {
        Write-Log "Updating certificate info..." -Level "INFO"
        $ykStatus = Get-YubiKeyCertExpiryDays
        $vscStatus = Get-VirtualSmartCardCertExpiry
        $combinedStatus = "$ykStatus`n$vscStatus"
        $global:CachedCertificateStatus = $combinedStatus
        $window.Dispatcher.Invoke({ $global:YubiKeyComplianceText.Text = $combinedStatus })
    } catch {
        Handle-Error $_.Exception.Message -Source "Update-CertificateInfo"
    }
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
    } catch {
        Write-Log "Windows build number check error: $($_.Exception.Message)" -Level "ERROR"
        return "Windows Build: Unknown"
    }
}

function Convert-MarkdownToTextBlock {
    param(
        [string]$Text,
        [System.Windows.Controls.TextBlock]$TargetTextBlock
    )
    
    $TargetTextBlock.Inlines.Clear()
    
    $regexBold = "\*\*(.*?)\*\*"
    $regexItalic = "\*(.*?)\*"
    $regexUnderline = "__(.*?)__"
    $regexColor = "\[(green|red|yellow)\](.*?)\[/\1\]"
    
    $currentText = $Text
    $lastIndex = 0
    $matches = @()

    # Find all bold, italic, underline, and color matches
    $boldMatches = [regex]::Matches($Text, $regexBold)
    $italicMatches = [regex]::Matches($Text, $regexItalic)
    $underlineMatches = [regex]::Matches($Text, $regexUnderline)
    $colorMatches = [regex]::Matches($Text, $regexColor)
    
    # Combine and sort matches by index
    foreach ($match in $boldMatches) {
        $matches += [PSCustomObject]@{ Index = $match.Index; Length = $match.Length; Text = $match.Groups[1].Value; Type = "Bold" }
    }
    foreach ($match in $italicMatches) {
        $matches += [PSCustomObject]@{ Index = $match.Index; Length = $match.Length; Text = $match.Groups[1].Value; Type = "Italic" }
    }
    foreach ($match in $underlineMatches) {
        $matches += [PSCustomObject]@{ Index = $match.Index; Length = $match.Length; Text = $match.Groups[1].Value; Type = "Underline" }
    }
    foreach ($match in $colorMatches) {
        $matches += [PSCustomObject]@{ Index = $match.Index; Length = $match.Length; Text = $match.Groups[2].Value; Type = "Color"; Color = $match.Groups[1].Value }
    }
    $matches = $matches | Sort-Object Index

    foreach ($match in $matches) {
        # Add text before the match
        if ($match.Index -gt $lastIndex) {
            $plainText = $currentText.Substring($lastIndex, $match.Index - $lastIndex)
            $TargetTextBlock.Inlines.Add((New-Object System.Windows.Documents.Run($plainText)))
        }
        
        # Add formatted text
        $run = New-Object System.Windows.Documents.Run($match.Text)
        if ($match.Type -eq "Bold") {
            $run.FontWeight = [System.Windows.FontWeights]::Bold
        } elseif ($match.Type -eq "Italic") {
            $run.FontStyle = [System.Windows.FontStyles]::Italic
        } elseif ($match.Type -eq "Underline") {
            $run.TextDecorations = [System.Windows.TextDecorations]::Underline
        } elseif ($match.Type -eq "Color") {
            $colorBrush = [System.Windows.Media.Brushes]::($match.Color.Substring(0,1).ToUpper() + $match.Color.Substring(1))
            $run.Foreground = $colorBrush
        }
        $TargetTextBlock.Inlines.Add($run)
        
        $lastIndex = $match.Index + $match.Length
    }
    
    # Add remaining text
    if ($lastIndex -lt $currentText.Length) {
        $plainText = $currentText.Substring($lastIndex)
        $TargetTextBlock.Inlines.Add((New-Object System.Windows.Documents.Run($plainText)))
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
            $global:AnnouncementsAlertIcon.Visibility = "Visible"
        }
        Convert-MarkdownToTextBlock -Text $newAnnouncementsObject.Text -TargetTextBlock $global:AnnouncementsText
        Convert-MarkdownToTextBlock -Text $newAnnouncementsObject.Details -TargetTextBlock $global:AnnouncementsDetailsText
        $global:AnnouncementsLinksPanel.Children.Clear()
        if ($newAnnouncementsObject.Links) {
            foreach ($link in $newAnnouncementsObject.Links) {
                $global:AnnouncementsLinksPanel.Children.Add((New-HyperlinkBlock -Name $link.Name -Url $link.Url))
            }
        }
        $global:AnnouncementsSourceText.Text = "Source: $($global:contentData.Source)"
    })
    
    $config.AnnouncementsLastState = $newJsonState
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
            $global:SupportAlertIcon.Visibility = "Visible"
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
    
    $config.SupportLastState = $newJsonState
}

function Update-PatchingAndSystem {
    Write-Log "Updating Patching and System section..." -Level "INFO"
    $restartStatusText = Get-PendingRestartStatus
    $statusColor = if ($global:PendingRestart) { [System.Windows.Media.Brushes]::Red } else { [System.Windows.Media.Brushes]::Green }
    
    $relevanceQuery = "names of relevant fixlets whose (baseline flag of it = false and (name of it as lowercase contains `"microsoft`" or name of it as lowercase contains `"security update`")) of sites"
    $patchResult = Get-BigFixRelevanceResult -RelevanceQuery $relevanceQuery

    $finalPatchText = ""
    if ($patchResult -like "Error:*") {
        $finalPatchText = "BigFix API not available. Contact IT."
    } else {
        $fixlets = if ($patchResult -is [string] -and -not [string]::IsNullOrWhiteSpace($patchResult) -and $patchResult -notlike "Evaluation failed") {
            $patchResult -split "`n" | Where-Object { $_ -match "\S" } | ForEach-Object { " - $_" }
        } else {
            @("No applicable fixlets found.")
        }
        $finalPatchText = $fixlets -join "`n"
    }

    $windowsBuild = Get-WindowsBuildNumber

    $window.Dispatcher.Invoke({
        $global:PatchingDescriptionText.Text = "Lists available software updates. Updates marked (R) require a restart."
        $global:PendingRestartStatusText.Text = $restartStatusText
        $global:PendingRestartStatusText.Foreground = $statusColor
        $global:PatchingUpdatesText.Text = $finalPatchText
        $global:WindowsBuildText.Text = $windowsBuild
    })
}

function Check-ScriptUpdate {
    try {
        $versionUrl = $config.VersionUrl
        Write-Log "Checking for script update from: $versionUrl" -Level "INFO"
        
        $job = Start-Job -ScriptBlock {
            param($url)
            Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop
        } -ArgumentList $versionUrl
        $response = Wait-Job $job | Receive-Job
        Remove-Job $job
        
        $remoteVersion = $response.Content.Trim()
        if ([version]$remoteVersion -gt [version]$ScriptVersion) {
            Write-Log "New script version available: $remoteVersion" -Level "INFO"
            $window.Dispatcher.Invoke({
                $global:ScriptUpdateText.Text = "New version $remoteVersion available. Updating now..."
                $global:ScriptUpdateText.Visibility = "Visible"
            })
            Perform-AutoUpdate -RemoteVersion $remoteVersion
            return $true
        }
        Write-Log "No new script version available." -Level "INFO"
        $window.Dispatcher.Invoke({
            $global:ScriptUpdateText.Visibility = "Hidden"
        })
        return $false
    } catch {
        Write-Log "Failed to check for script update: $($_.Exception.Message)" -Level "WARNING"
        return $false
    }
}

function Perform-AutoUpdate {
    param([string]$RemoteVersion)
    try {
        $scriptUrl = $config.ScriptUrl
        $newScriptPath = Join-Path $ScriptDir "LLNOTIFY.new.ps1"
        $batchPath = Join-Path $ScriptDir "update.bat"
        
        # Remove any existing batch file to avoid conflicts
        if (Test-Path $batchPath) {
            Remove-Item $batchPath -Force
            if (Test-Path $batchPath) {
                Write-Log "Warning: Could not remove existing update.bat" -Level "WARNING"
            } else {
                Write-Log "Removed existing update.bat" -Level "INFO"
            }
        }

        # Download new script
        $job = Start-Job -ScriptBlock {
            param($url, $path)
            Invoke-WebRequest -Uri $url -UseBasicParsing -OutFile $path
        } -ArgumentList $scriptUrl, $newScriptPath
        Wait-Job $job | Receive-Job
        Remove-Job $job
        
        if (-not (Test-Path $newScriptPath)) {
            throw "Failed to download new script."
        }

        # Create batch file for replacement and restart with reliable self-delete
        try {
            $batchContent = @"
@echo off
echo Batch started %date% %time% >> "%PSScriptRoot%\batch_log.txt"
timeout /t 5 /nobreak >nul
set /a attempts=0
:retry
set /a attempts+=1
echo Attempt %attempts% to move >> "%PSScriptRoot%\batch_log.txt"
move /Y "$newScriptPath" "$PSScriptRoot\LLNOTIFY.ps1" >> "%PSScriptRoot%\batch_log.txt" 2>&1
if ERRORLEVEL 1 (
  if %attempts% GEQ 5 goto fail
  timeout /t 2 /nobreak >nul
  goto retry
)
echo Move succeeded >> "%PSScriptRoot%\batch_log.txt"
echo Starting powershell >> "%PSScriptRoot%\batch_log.txt"
powershell -ExecutionPolicy Bypass -File "$PSScriptRoot\LLNOTIFY.ps1" >> "%PSScriptRoot%\batch_log.txt" 2>&1
if ERRORLEVEL 1 echo Relaunch failed with code %ERRORLEVEL% >> "%PSScriptRoot%\batch_log.txt"
echo Relaunch complete >> "%PSScriptRoot%\batch_log.txt"
start /b "" cmd /c del "%~f0" & exit
:fail
echo Failed to update after 5 attempts >> "%PSScriptRoot%\update_error.log"
"@
            $batchContent | Out-File $batchPath -Encoding ascii -Force
            if (-not (Test-Path $batchPath)) {
                throw "Batch file creation failed without error"
            }
            Write-Log "Created new update.bat at $batchPath" -Level "INFO"
        } catch {
            throw "Failed to create batch file: $($_.Exception.Message)"
        }

        # Start the batch and exit
        Start-Process -FilePath $batchPath -WindowStyle Hidden
        Write-Log "Auto-update initiated. Exiting current instance." -Level "INFO"
        Start-Sleep -Milliseconds 500  # Brief delay for cleanup
        $global:IsUpdating = $true
        $window.Dispatcher.InvokeShutdown()
    } catch {
        Write-Log "Auto-update failed: $($_.Exception.Message)" -Level "ERROR"
        $window.Dispatcher.Invoke({
            $global:ScriptUpdateText.Text = "Update failed. Please update manually."
        })
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
        try {
            return New-Object System.Drawing.Icon($Path)
        }
        catch {
            Write-Log "Error loading icon from `"$Path`": $($_.Exception.Message)" -Level "ERROR"
        }
    }
    return [System.Drawing.SystemIcons]::Application
}

function Update-TrayIcon {
    if (-not $global:TrayIcon.Visible) { return }
    
    $announcementAlert = $global:AnnouncementsAlertIcon.Visibility -eq "Visible"
    $supportAlert = $global:SupportAlertIcon.Visibility -eq "Visible"
    
    $hasBlinkingAlert = $announcementAlert -or $supportAlert
    $hasAnyAlert = $global:PendingRestart -or $hasBlinkingAlert

    if ($hasBlinkingAlert -and -not $window.IsVisible) {
        if ($config.BlinkingEnabled) {
            if (-not $global:BlinkingTimer.IsEnabled) {
                $global:TrayIcon.Icon = $global:WarningIcon
                $global:BlinkingTimer.Start()
            }
        } else {
            $global:TrayIcon.Icon = $global:WarningIcon
            $global:TrayIcon.Text = "LLNOTIFY v$ScriptVersion - Alerts Pending"
        }
    }
    else {
        if ($global:BlinkingTimer.IsEnabled) {
            $global:BlinkingTimer.Stop()
        }
        $global:TrayIcon.Icon = if ($hasAnyAlert) { $global:WarningIcon } else { $global:MainIcon }
    }
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
        
        # Submenu for Set Update Interval
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
            $global:DispatcherTimer.Interval = [TimeSpan]::FromSeconds(1200)
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
    } catch {
        Handle-Error $_.Exception.Message -Source "Initialize-TrayIcon"
    }
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
        Write-Log "Main update cycle running..." -Level "INFO"
        $global:contentData = Fetch-ContentData
        
        Update-Announcements
        Update-Support
        Update-PatchingAndSystem
        
        if ($ForceCertificateCheck -or (-not $global:LastCertificateCheck -or ((Get-Date) - $global:LastCertificateCheck).TotalSeconds -ge $config.CertificateCheckInterval)) {
            Update-CertificateInfo
            $global:LastCertificateCheck = Get-Date
        }
        
        if (Check-ScriptUpdate) { return }  # Stop cycle if update started
        
        Generate-BigFixComplianceReport
        
        Update-TrayIcon
        Save-Configuration -Config $config
        Rotate-LogFile
    }
    catch {
        Handle-Error $_.Exception.Message -Source "Main-UpdateCycle"
    }
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
    Initialize-BigFixAPIConfig  # Ensure BigFix API is configured before queries
    Main-UpdateCycle -ForceCertificateCheck $true
    
    $global:DispatcherTimer.Start()
    Write-Log "Main timer started." -Level "INFO"
    
    $window.Dispatcher.Add_UnhandledException({ 
        Handle-Error $_.Exception.Message -Source "Dispatcher"
        $_.Handled = $true 
    })

    Write-Log "Application startup complete. Running dispatcher." -Level "INFO"
    if (-not $global:IsUpdating) {
        [System.Windows.Threading.Dispatcher]::Run()
    }
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
