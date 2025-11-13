action uses wow64 redirection {not x64 of operating system}

// 1. Ensure the target folder exists
folder create "C:\Program Files\LLEA"
waithidden cmd.exe /c icacls "C:\Program Files\LLEA" /grant "Users":(OI)(CI)F /t

// 2. Build a robust PowerShell downloader (replaces certutil)
delete __createfile
createfile until END_OF_DOWNLOAD_SCRIPT
# LLEA Download Script - Robust replacement for certutil
# Uses Invoke-WebRequest with BITS fallback and retry logic

$ErrorActionPreference = "Stop"

# Enable TLS 1.2 and 1.3 for HTTPS connections
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12, [Net.SecurityProtocolType]::Tls13

# Define files to download
$filesToDownload = @(
    @{
        Url = "https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LLEA.ps1"
        Destination = "C:\Program Files\LLEA\LLEA.ps1"
    },
    @{
        Url = "https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/DriverUpdate.ps1"
        Destination = "C:\Program Files\LLEA\DriverUpdate.ps1"
    },
    @{
        Url = "https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LL_LOGO.ico"
        Destination = "C:\Program Files\LLEA\LL_LOGO.ico"
    },
    @{
        Url = "https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LL_LOGO_MSG.ico"
        Destination = "C:\Program Files\LLEA\LL_LOGO_MSG.ico"
    }
)

# Function to download with Invoke-WebRequest (primary method)
function Download-WithInvokeWebRequest {
    param($Url, $Destination, $MaxRetries = 3)
    
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            Write-Host "[$attempt/$MaxRetries] Downloading $(Split-Path $Destination -Leaf) via Invoke-WebRequest..."
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri $Url -OutFile $Destination -UseBasicParsing -TimeoutSec 120 -ErrorAction Stop
            $ProgressPreference = 'Continue'
            
            if (Test-Path $Destination) {
                $size = (Get-Item $Destination).Length
                Write-Host "  SUCCESS: Downloaded $size bytes"
                return $true
            }
        } catch {
            $ProgressPreference = 'Continue'
            Write-Host "  FAILED: $($_.Exception.Message)"
            if ($attempt -lt $MaxRetries) {
                $waitSeconds = 5 * $attempt
                Write-Host "  Waiting $waitSeconds seconds before retry..."
                Start-Sleep -Seconds $waitSeconds
            }
        }
    }
    return $false
}

# Function to download with BITS (fallback method)
function Download-WithBITS {
    param($Url, $Destination)
    
    try {
        Write-Host "Trying BITS download for $(Split-Path $Destination -Leaf)..."
        
        # Check if BITS service is running
        $bitsService = Get-Service -Name BITS -ErrorAction SilentlyContinue
        if ($bitsService.Status -ne 'Running') {
            Write-Host "  Starting BITS service..."
            Start-Service -Name BITS -ErrorAction Stop
            Start-Sleep -Seconds 2
        }
        
        # Download via BITS
        Start-BitsTransfer -Source $Url -Destination $Destination -ErrorAction Stop
        
        if (Test-Path $Destination) {
            $size = (Get-Item $Destination).Length
            Write-Host "  SUCCESS: Downloaded $size bytes via BITS"
            return $true
        }
    } catch {
        Write-Host "  BITS FAILED: $($_.Exception.Message)"
        return $false
    }
    return $false
}

# Function to download with WebClient (secondary fallback)
function Download-WithWebClient {
    param($Url, $Destination, $MaxRetries = 2)
    
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        $webClient = $null
        try {
            Write-Host "[$attempt/$MaxRetries] Downloading $(Split-Path $Destination -Leaf) via WebClient..."
            $webClient = New-Object System.Net.WebClient
            $webClient.DownloadFile($Url, $Destination)
            
            if (Test-Path $Destination) {
                $size = (Get-Item $Destination).Length
                Write-Host "  SUCCESS: Downloaded $size bytes"
                return $true
            }
        } catch {
            Write-Host "  FAILED: $($_.Exception.Message)"
            if ($attempt -lt $MaxRetries) {
                Start-Sleep -Seconds 3
            }
        } finally {
            if ($webClient) { $webClient.Dispose() }
        }
    }
    return $false
}

# Main download loop with multiple methods
$allSuccess = $true
$downloadLog = @()

foreach ($file in $filesToDownload) {
    Write-Host ""
    Write-Host "================================================"
    Write-Host "Downloading: $(Split-Path $file.Destination -Leaf)"
    Write-Host "From: $($file.Url)"
    Write-Host "================================================"
    
    $success = $false
    
    # Try Method 1: Invoke-WebRequest (most reliable)
    Write-Host "Method 1: Invoke-WebRequest"
    $success = Download-WithInvokeWebRequest -Url $file.Url -Destination $file.Destination
    
    # Try Method 2: BITS (if Method 1 failed)
    if (-not $success) {
        Write-Host "Method 2: BITS"
        $success = Download-WithBITS -Url $file.Url -Destination $file.Destination
    }
    
    # Try Method 3: WebClient (if Methods 1 and 2 failed)
    if (-not $success) {
        Write-Host "Method 3: WebClient"
        $success = Download-WithWebClient -Url $file.Url -Destination $file.Destination
    }
    
    # Log result
    if ($success) {
        $downloadLog += "SUCCESS: $(Split-Path $file.Destination -Leaf)"
        Write-Host "FINAL RESULT: SUCCESS" -ForegroundColor Green
    } else {
        $downloadLog += "FAILED: $(Split-Path $file.Destination -Leaf)"
        Write-Host "FINAL RESULT: FAILED - All methods exhausted" -ForegroundColor Red
        $allSuccess = $false
    }
}

# Summary
Write-Host ""
Write-Host "================================================"
Write-Host "DOWNLOAD SUMMARY"
Write-Host "================================================"
foreach ($log in $downloadLog) {
    Write-Host $log
}

# Exit with appropriate code
if ($allSuccess) {
    Write-Host ""
    Write-Host "All files downloaded successfully!" -ForegroundColor Green
    exit 0
} else {
    Write-Host ""
    Write-Host "ERROR: Some files failed to download!" -ForegroundColor Red
    exit 1
}
END_OF_DOWNLOAD_SCRIPT

// 3. Drop the PowerShell script into place
copy __createfile "C:\Program Files\LLEA\download_LLEA.ps1"

// 4. Run it as the current user
override wait
hidden=true
wait powershell.exe -ExecutionPolicy Bypass -NoProfile -File "C:\Program Files\LLEA\download_LLEA.ps1"

// 4a. Check if download was successful before proceeding
continue if {exists file "LLEA.ps1" of folder "LLEA" of folder "Program Files" of drive of system folder}

// 5. Clean up the download script
delete "C:\Program Files\LLEA\download_LLEA.ps1"

// 5a. Ensure log directory exists
folder create "C:\Windows\MITLL\Logs"

// 5b. Create scheduled task using PowerShell and grant user permissions
delete __createfile
createfile until END_OF_TASK_CREATION
$taskName = "MITLL_DriverUpdate"
$taskDescription = "Monthly Windows driver updates via Windows Update"
$scriptPath = "C:\Program Files\LLEA\DriverUpdate.ps1"
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -File `"$scriptPath`""
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 2)
Register-ScheduledTask -TaskName $taskName -Description $taskDescription -Action $action -Principal $principal -Settings $settings -Force
# Grant authenticated users permission to start the task
$taskPath = "\$taskName"
$sid = New-Object System.Security.Principal.SecurityIdentifier("S-1-5-11")  # Authenticated Users
$taskScheduler = New-Object -ComObject Schedule.Service
$taskScheduler.Connect()
$rootFolder = $taskScheduler.GetFolder("\")
$task = $rootFolder.GetTask($taskName)
$securityDescriptor = $task.GetSecurityDescriptor(0xF)
$securityDescriptor += "(A;;GRGX;;;AU)"  # Grant Read and Execute to Authenticated Users
$task.SetSecurityDescriptor($securityDescriptor, 0)
Write-Output "Scheduled task created and permissions granted."
END_OF_TASK_CREATION

move __createfile "C:\Program Files\LLEA\CreateTask.ps1"

override wait
hidden=true
wait powershell.exe -ExecutionPolicy Bypass -File "C:\Program Files\LLEA\CreateTask.ps1"

delete "C:\Program Files\LLEA\CreateTask.ps1"

// 6. Register per machine Run key (hidden PowerShell)
override wait
hidden=true
wait reg add "HKLM\Software\Microsoft\Windows\CurrentVersion\Run" /v "LLEA" /t REG_SZ /d "\"C:\Windows\System32\conhost.exe\" --headless \"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe\" -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File \"C:\Program Files\LLEA\LLEA.ps1\" -RunMode LLEA\"" /f

// 6a. Brief pause to ensure cleanup complete
wait {pathname of system folder}\timeout.exe 1 /nobreak

// 7. Immediately invoke the (signed) script once
override wait
hidden=true
runas=currentuser  
wait cmd.exe /C start "" /b powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\Program Files\LLEA\LLEA.ps1" -RunMode LLEA
