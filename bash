#region ===== PSADT 4.1–aware helpers =====
function Write-PSADTLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateRange(1,3)][int]$Severity = 1
    )
    if (Get-Command -Name Write-ADTLogEntry -ErrorAction SilentlyContinue) {
        Write-ADTLogEntry -Message $Message -Severity $Severity
    } elseif (Get-Command -Name Write-Log -ErrorAction SilentlyContinue) {
        Write-Log -Message $Message -Severity $Severity
    } else {
        Write-Host "[$Severity] $Message"
    }
}

function Invoke-PSADTProcess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Parameters,
        [switch]$IgnoreExitCodes
    )
    if (Get-Command -Name Execute-ADTProcess -ErrorAction SilentlyContinue) {
        Execute-ADTProcess -Path $Path -Parameters $Parameters -IgnoreExitCodes:$IgnoreExitCodes
    } else {
        Execute-Process -Path $Path -Parameters $Parameters -WindowStyle Hidden -CreateNoWindow $true -IgnoreExitCodes:$IgnoreExitCodes
    }
}
#endregion

#region ===== SAP AO detection/installation =====
function Get-SAPAOState {
    [CmdletBinding()] param()

    $roots = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    )

    $items = foreach ($r in $roots) {
        if (Test-Path $r) {
            Get-ChildItem $r -ErrorAction SilentlyContinue | ForEach-Object {
                try { Get-ItemProperty $_.PsPath -ErrorAction Stop } catch { $null }
            }
        }
    }

    $regHit = $items | Where-Object {
        ($_.DisplayName -match 'Analysis for (Microsoft )?Office' -or
         $_.DisplayName -match 'SAP BusinessObjects Analysis' -or
         $_.DisplayName -match 'SAP Analysis') -and
        ($_.Publisher -match 'SAP')
    } | Select-Object -First 1

    $pf   = ${env:ProgramFiles}
    $pf86 = ${env:ProgramFiles(x86)}
    $paths = @(
        Join-Path $pf   'SAP BusinessObjects\Office AddIn',
        Join-Path $pf86 'SAP BusinessObjects\Office AddIn'
    )
    $folderHit = $paths | Where-Object { Test-Path $_ } | Select-Object -First 1

    $arch =
        if     ($regHit -and $regHit.DisplayName -match '(x64|64)') { 'x64' }
        elseif ($regHit -and $regHit.DisplayName -match '(x86|32)') { 'x86' }
        elseif ($folderHit -and $folderHit -like "$pf86*")          { 'x86' }
        elseif ($folderHit)                                         { 'x64' }
        else                                                        { $null }

    [PSCustomObject]@{
        Present         = [bool]($regHit -or $folderHit)
        DisplayName     = $regHit.DisplayName
        Version         = $regHit.DisplayVersion
        Architecture    = $arch
        InstallPath     = $folderHit
        UninstallString = $regHit.UninstallString
        RegistryPath    = if ($regHit.PSPath) { ($regHit.PSPath -split '::')[-1] }
    }
}

function Test-SAPAOInstalled {
    try {
        return (Get-SAPAOState).Present
    } catch {
        Write-PSADTLog -Message "Test-SAPAOInstalled error: $($_.Exception.Message)" -Severity 3
        return $false
    }
}

function Test-SAPAOExcelAddin {
    [CmdletBinding()] param()
    $excel = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $addin = $excel.AddIns | Where-Object { $_.Name -match 'Analysis' -or $_.Title -match 'Analysis' }
        return [bool]$addin
    } catch {
        Write-PSADTLog -Message "Excel COM add-in check skipped/failed: $($_.Exception.Message)" -Severity 2
        return $false
    } finally {
        if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    }
}

function Get-SAPAOProductSwitch {
    param([ValidateSet('x64','x86')] [string] $Architecture = 'x64')
    switch ($Architecture) {
        'x64' { 'SapCofx64' }
        'x86' { 'SapCofx86' }
    }
}

function Install-SAPAOIfNeeded {
    [CmdletBinding()]
    param(
        [ValidateSet('x64','x86')] [string] $Architecture = 'x64',
        [switch] $Force
    )

    $state = Get-SAPAOState
    if ($state.Present -and -not $Force) {
        Write-PSADTLog -Message "SAP AO present ($($state.DisplayName) $($state.Version)); skipping install." -Severity 1
        return
    }

    $exe = Join-Path $dirFiles 'SAPBAO\Setup\NwSapSetup.exe'
    if (-not (Test-Path $exe)) { throw "Installer not found: $exe" }

    $prod   = Get-SAPAOProductSwitch -Architecture $Architecture
    $params = "/product=`"$prod`" /Silent"

    Write-PSADTLog -Message "Installing SAP AO $Architecture → $exe $params" -Severity 1
    Invoke-PSADTProcess -Path $exe -Parameters $params -IgnoreExitCodes:$false
}
#endregion


----------

# --- Pre-Installation ---
$ao = Get-SAPAOState
if ($ao.Present) {
    Write-PSADTLog -Message "Pre-check: SAP AO detected ($($ao.DisplayName) $($ao.Version)); flagging for post-upgrade reinstall." -Severity 1
    New-Item -Path 'HKLM:\SOFTWARE\LL\OfficeUpgrade' -Force | Out-Null
    New-ItemProperty -Path 'HKLM:\SOFTWARE\LL\OfficeUpgrade' -Name 'ReinstallSAPAO' -Value 1 -PropertyType DWord -Force | Out-Null
    if ($ao.Architecture) {
        New-ItemProperty -Path 'HKLM:\SOFTWARE\LL\OfficeUpgrade' -Name 'SAPAOArch' -Value $ao.Architecture -PropertyType String -Force | Out-Null
    }
} else {
    Write-PSADTLog -Message "Pre-check: SAP AO not detected." -Severity 1
}


----------

# --- Post-Installation ---
$needReinstall = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\LL\OfficeUpgrade' -ErrorAction SilentlyContinue).ReinstallSAPAO -eq 1
$archPref      = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\LL\OfficeUpgrade' -ErrorAction SilentlyContinue).SAPAOArch

$stateAfter = Get-SAPAOState
$addinOk    = Test-SAPAOExcelAddin   # optional sanity check

if ($needReinstall -or (-not $stateAfter.Present) -or (-not $addinOk)) {
    $archToInstall = if ($archPref) { $archPref } else { 'x64' }
    Write-PSADTLog -Message "Triggering SAP AO (re)install. Present=$($stateAfter.Present) ExcelAddinOK=$addinOk Arch=$archToInstall" -Severity 1
    Install-SAPAOIfNeeded -Architecture $archToInstall -Force

    Remove-ItemProperty -Path 'HKLM:\SOFTWARE\LL\OfficeUpgrade' -Name 'ReinstallSAPAO' -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path 'HKLM:\SOFTWARE\LL\OfficeUpgrade' -Name 'SAPAOArch' -ErrorAction SilentlyContinue
} else {
    Write-PSADTLog -Message "SAP AO OK post-upgrade; no action required." -Severity 1
}
