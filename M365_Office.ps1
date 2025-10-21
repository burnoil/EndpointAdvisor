<#

.SYNOPSIS
PSAppDeployToolkit - This script performs the installation or uninstallation of an application(s).

.DESCRIPTION
- The script is provided as a template to perform an install, uninstall, or repair of an application(s).
- The script either performs an "Install", "Uninstall", or "Repair" deployment type.
- The install deployment type is broken down into 3 main sections/phases: Pre-Install, Install, and Post-Install.

The script imports the PSAppDeployToolkit module which contains the logic and functions required to install or uninstall an application.

PSAppDeployToolkit is licensed under the GNU LGPLv3 License - (C) 2025 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).

This program is free software: you can redistribute it and/or modify it under the terms of the GNU Lesser General Public License as published by the
Free Software Foundation, either version 3 of the License, or any later version. This program is distributed in the hope that it will be useful, but
WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License
for more details. You should have received a copy of the GNU Lesser General Public License along with this program. If not, see <http://www.gnu.org/licenses/>.

.PARAMETER DeploymentType
The type of deployment to perform.

.PARAMETER DeployMode
Specifies whether the installation should be run in Interactive (shows dialogs), Silent (no dialogs), or NonInteractive (dialogs without prompts) mode.

NonInteractive mode is automatically set if it is detected that the process is not user interactive.

.PARAMETER AllowRebootPassThru
Allows the 3010 return code (requires restart) to be passed back to the parent process (e.g. SCCM) if detected from an installation. If 3010 is passed back to SCCM, a reboot prompt will be triggered.

.PARAMETER TerminalServerMode
Changes to "user install mode" and back to "user execute mode" for installing/uninstalling applications for Remote Desktop Session Hosts/Citrix servers.

.PARAMETER DisableLogging
Disables logging to file for the script.

.EXAMPLE
powershell.exe -File Invoke-AppDeployToolkit.ps1 -DeployMode Silent

.EXAMPLE
powershell.exe -File Invoke-AppDeployToolkit.ps1 -AllowRebootPassThru

.EXAMPLE
powershell.exe -File Invoke-AppDeployToolkit.ps1 -DeploymentType Uninstall

.EXAMPLE
Invoke-AppDeployToolkit.exe -DeploymentType "Install" -DeployMode "Silent"

.INPUTS
None. You cannot pipe objects to this script.

.OUTPUTS
None. This script does not generate any output.

.NOTES
Toolkit Exit Code Ranges:
- 60000 - 68999: Reserved for built-in exit codes in Invoke-AppDeployToolkit.ps1, and Invoke-AppDeployToolkit.exe
- 69000 - 69999: Recommended for user customized exit codes in Invoke-AppDeployToolkit.ps1
- 70000 - 79999: Recommended for user customized exit codes in PSAppDeployToolkit.Extensions module.

.LINK
https://psappdeploytoolkit.com

#>

[CmdletBinding()]
param
(
    [Parameter(Mandatory = $false)]
    [ValidateSet('Install', 'Uninstall', 'Repair')]
    [PSDefaultValue(Help = 'Install', Value = 'Install')]
    [System.String]$DeploymentType,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Interactive', 'Silent', 'NonInteractive')]
    [PSDefaultValue(Help = 'Interactive', Value = 'Interactive')]
    [System.String]$DeployMode,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.SwitchParameter]$AllowRebootPassThru,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.SwitchParameter]$TerminalServerMode,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.SwitchParameter]$DisableLogging
)


##================================================
## MARK: Variables
##================================================

$adtSession = @{
    # App variables.
    AppVendor = 'Microsoft'
    AppName = '365 Apps for Enterprise'
    AppVersion = '16.0.18925.20184'
    AppArch = 'x64'
    AppLang = 'EN'
    AppRevision = '01'
    AppSuccessExitCodes = @(0)
    AppRebootExitCodes = @(1641, 3010)
    AppScriptVersion = '1.0.0'
    AppScriptDate = '2025-10-20'
    AppScriptAuthor = 'Todd Loenhorst'

    # Install Titles (Only set here to override defaults set by the toolkit).
    #InstallName = ''
    #InstallTitle = ''

    # Script variables.
    DeployAppScriptFriendlyName = $MyInvocation.MyCommand.Name
    DeployAppScriptVersion = '4.1.0'
    DeployAppScriptParameters = $PSBoundParameters
}

##================================================
## MARK: SAP AO Helpers (PSADT 4.1 style)
##================================================
#region ===== SAP Analysis for Office (SAP AO) detection / install helpers =====

function Get-SAPAOState {
    [CmdletBinding()] param()

    # Uninstall registry roots (native + WOW6432)
    $roots = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    )

    $items = @()
    foreach ($r in $roots) {
        if (Test-Path -LiteralPath $r) {
            $items += Get-ChildItem -LiteralPath $r -ErrorAction SilentlyContinue | ForEach-Object {
                try { Get-ItemProperty -LiteralPath $_.PsPath -ErrorAction Stop } catch { $null }
            }
        }
    }
    # Drop nulls
    $items = $items | Where-Object { $_ }

    # Identify SAP AO by DisplayName + Publisher
    $regHit = $items | Where-Object {
        $_.PSObject.Properties.Match('DisplayName').Count -gt 0 -and
        $_.PSObject.Properties.Match('Publisher').Count   -gt 0 -and
        -not [string]::IsNullOrWhiteSpace($_.DisplayName) -and
        -not [string]::IsNullOrWhiteSpace($_.Publisher)   -and
        (
            $_.DisplayName -match 'Analysis for (Microsoft )?Office' -or
            $_.DisplayName -match 'SAP BusinessObjects Analysis'     -or
            $_.DisplayName -match 'SAP Analysis'
        ) -and
        ($_.Publisher -match 'SAP')
    } | Select-Object -First 1

    # Build candidate install paths safely
    $pf   = [Environment]::GetFolderPath('ProgramFiles')
    $pf86 = [Environment]::GetFolderPath('ProgramFilesX86')

    $paths = @()
    if ($pf   -and -not [string]::IsNullOrWhiteSpace($pf))   { $paths += (Join-Path -Path $pf   -ChildPath 'SAP BusinessObjects\Office AddIn') }
    if ($pf86 -and -not [string]::IsNullOrWhiteSpace($pf86)) { $paths += (Join-Path -Path $pf86 -ChildPath 'SAP BusinessObjects\Office AddIn') }

    $folderHit = $paths | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1

    # Infer architecture
    $arch =
        if     ($regHit -and ($regHit.DisplayName -match '(x64|64)')) { 'x64' }
        elseif ($regHit -and ($regHit.DisplayName -match '(x86|32)')) { 'x86' }
        elseif ($folderHit -and $pf86 -and $folderHit.StartsWith($pf86, [System.StringComparison]::OrdinalIgnoreCase)) { 'x86' }
        elseif ($folderHit) { 'x64' }
        else { $null }

    [PSCustomObject]@{
        Present         = [bool]($regHit -or $folderHit)
        DisplayName     = if ($regHit) { $regHit.DisplayName } else { $null }
        Version         = if ($regHit) { $regHit.DisplayVersion } else { $null }
        Architecture    = $arch
        InstallPath     = $folderHit
        UninstallString = if ($regHit) { $regHit.UninstallString } else { $null }
        RegistryPath    = if ($regHit -and $regHit.PSPath) { ($regHit.PSPath -split '::')[-1] } else { $null }
    }
}

function Test-SAPAOInstalled {
    try { (Get-SAPAOState).Present } catch {
        Write-ADTLogEntry -Message "Test-SAPAOInstalled error: $($_.Exception.Message)" -Severity 3 -Source 'Detect-SAPAO'
        $false
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
        Write-ADTLogEntry -Message "SAP AO present ($($state.DisplayName) $($state.Version)); skipping install." -Severity 1 -Source 'Install-SAPAO'
        return
    }

    $exe = Join-Path $adtSession.DirFiles 'SAPBAO\Setup\NwSapSetup.exe'
    if (-not (Test-Path -LiteralPath $exe)) { throw "SAP AO installer not found: $exe" }

    $prod   = Get-SAPAOProductSwitch -Architecture $Architecture
    $params = "/product=`"$prod`" /Silent"

    Write-ADTLogEntry -Message "Installing SAP AO $Architecture → $exe $params" -Severity 1 -Source 'Install-SAPAO'
    Start-ADTProcess -FilePath $exe -ArgumentList $params
}

# Optional: Excel COM add-in sanity check (not used for gating install)
function Test-SAPAOExcelAddin {
    [CmdletBinding()] param()
    $excel = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $addin = $excel.AddIns | Where-Object { $_.Name -match 'Analysis' -or $_.Title -match 'Analysis' }
        [bool]$addin
    } catch {
        Write-ADTLogEntry -Message "Excel COM add-in check skipped/failed: $($_.Exception.Message)" -Severity 2 -Source 'Detect-SAPAO'
        $false
    } finally {
        if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    }
}
#endregion


function Install-ADTDeployment
{
    ##================================================
    ## MARK: Pre-Install
    ##================================================
    $adtSession.InstallPhase = "Pre-$($adtSession.DeploymentType)"

    ## Welcome / close apps
	Show-ADTInstallationWelcome `
    -CloseProcesses @{ Name = 'outlook'; Description = 'Microsoft Outlook' }, @{ Name = 'winword'; Description = 'Microsoft Office Word' }, @{ Name = 'excel'; Description = 'Microsoft Office Excel' }, @{ Name = 'powerpnt'; Description = 'Microsoft PowerPoint' }, @{ Name = 'onenote'; Description = 'Microsoft OneNote' } `
    -BlockExecution `
    -CloseProcessesCountdown 600 `
    -PersistPrompt

	## Progress
	Show-ADTInstallationProgress -StatusMessage "Microsoft 365 Apps installation in Progress...`nThis installation may take approximately 20-30 minutes to complete. Please wait..."

    ## <Perform Pre-Installation tasks here>
    # Stop any custom processes (sample placeholders)
    $processes = @("process1", "process2", "process3")
    foreach ($process in $processes) {
        if (Get-Process -Name $process -ErrorAction SilentlyContinue) {
            Write-ADTLogEntry -Message "Stopping the process '$process'..." -Source $adtSession.InstallPhase
            Stop-Process -Name $process -Force -ErrorAction SilentlyContinue
            Write-ADTLogEntry -Message "'$process' process has been stopped." -Source $adtSession.InstallPhase
        } else {
            Write-ADTLogEntry -Message "'$process' is not running." -Source $adtSession.InstallPhase
        }
    }

    # --- SAP AO pre-check: if present, flag for post-upgrade reinstall ---
    try {
        $ao = Get-SAPAOState
        if ($ao.Present) {
            Write-ADTLogEntry -Message "Pre-check: SAP AO detected ($($ao.DisplayName) $($ao.Version)); flagging for post-upgrade reinstall." -Severity 1 -Source $adtSession.InstallPhase
            New-Item -Path 'HKLM:\SOFTWARE\LL\OfficeUpgrade' -Force | Out-Null
            New-ItemProperty -Path 'HKLM:\SOFTWARE\LL\OfficeUpgrade' -Name 'ReinstallSAPAO' -Value 1 -PropertyType DWord -Force | Out-Null
            if ($ao.Architecture) {
                New-ItemProperty -Path 'HKLM:\SOFTWARE\LL\OfficeUpgrade' -Name 'SAPAOArch' -Value $ao.Architecture -PropertyType String -Force | Out-Null
            }
        } else {
            Write-ADTLogEntry -Message "Pre-check: SAP AO not detected." -Severity 1 -Source $adtSession.InstallPhase
        }
    } catch {
        Write-ADTLogEntry -Message "SAP AO pre-check failed: $($_.Exception.Message)" -Severity 2 -Source $adtSession.InstallPhase
    }

    ##================================================
    ## MARK: Install
    ##================================================
    $adtSession.InstallPhase = $adtSession.DeploymentType

    <## Handle Zero-Config MSI installations.
    if ($adtSession.UseDefaultMsi)
    {
        $ExecuteDefaultMSISplat = @{ Action = $adtSession.DeploymentType; FilePath = $adtSession.DefaultMsiFile }
        if ($adtSession.DefaultMstFile)
        {
            $ExecuteDefaultMSISplat.Add('Transform', $adtSession.DefaultMstFile)
        }
        Start-ADTMsiProcess @ExecuteDefaultMSISplat
        if ($adtSession.DefaultMspFiles)
        {
            $adtSession.DefaultMspFiles | Start-ADTMsiProcess -Action Patch
        }
    }
    #>

    ## <Perform Installation tasks here>
    # --- ODT (Office Deployment Tool) pre-flight + install ---
    try {
        $setupExe  = Join-Path -Path $adtSession.DirFiles -ChildPath 'Setup.exe'
        $configXml = Join-Path -Path $adtSession.DirFiles -ChildPath 'configuration.xml'

        if (-not (Test-Path -LiteralPath $setupExe)) {
            Write-ADTLogEntry -Message "ODT payload missing: $setupExe" -Severity 3 -Source $adtSession.InstallPhase
            throw "ODT Setup.exe not found at: $setupExe"
        }
        if (-not (Test-Path -LiteralPath $configXml)) {
            Write-ADTLogEntry -Message "ODT configuration missing: $configXml" -Severity 3 -Source $adtSession.InstallPhase
            throw "ODT configuration.xml not found at: $configXml"
        }

        Write-ADTLogEntry -Message "Launching ODT: `"$setupExe`" /configure `"$configXml`"" -Severity 1 -Source $adtSession.InstallPhase
        Start-ADTProcess -FilePath $setupExe `
                         -ArgumentList "/configure `"$configXml`"" `
                         -WorkingDirectory $adtSession.DirFiles
						 -WindowStyle Hidden `
    }
    catch {
        Write-ADTLogEntry -Message "ODT install failed: $($_.Exception.Message)" -Severity 3 -Source $adtSession.InstallPhase
        try {
            $odtLog = Get-ChildItem -Path (Join-Path $env:WINDIR 'Temp') -Filter 'OfficeDeploymentTool*.log' -ErrorAction SilentlyContinue |
                      Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($odtLog) { Write-ADTLogEntry -Message "Latest ODT log: $($odtLog.FullName)" -Severity 2 -Source $adtSession.InstallPhase }
        } catch { }
        throw
    }

    ##================================================
    ## MARK: Post-Install
    ##================================================
    $adtSession.InstallPhase = "Post-$($adtSession.DeploymentType)"

    ## <Perform Post-Installation tasks here>
    # Example: remove a public desktop shortcut if it exists
    $DESKTOPICONPATH = "$envCommonDesktop\none.lnk"
    If (Test-Path $DESKTOPICONPATH) {
        Remove-Item $DESKTOPICONPATH -Force
        Write-ADTLogEntry -Message "$DESKTOPICONPATH found.  Removing ..." -Source $adtSession.InstallPhase
    }

    # --- SAP AO post-check: install ONLY if it existed pre-upgrade (flag set) ---
    try {
        $needReinstall = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\LL\OfficeUpgrade' -ErrorAction SilentlyContinue).ReinstallSAPAO -eq 1
        $archPref      = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\LL\OfficeUpgrade' -ErrorAction SilentlyContinue).SAPAOArch
        $archToInstall = if ($archPref) { $archPref } else { 'x64' }

        if ($needReinstall) {
            $stateAfter = Get-SAPAOState

            if ($stateAfter.Present) {
                Write-ADTLogEntry -Message "SAP AO was flagged pre-upgrade but is still present post-upgrade ($($stateAfter.DisplayName) $($stateAfter.Version)). Skipping reinstall." -Severity 1 -Source $adtSession.InstallPhase
            } else {
                Write-ADTLogEntry -Message "SAP AO was present pre-upgrade and is missing post-upgrade. Installing ($archToInstall)..." -Severity 1 -Source $adtSession.InstallPhase
                Install-SAPAOIfNeeded -Architecture $archToInstall -Force
            }

            # cleanup flag
            Remove-ItemProperty -Path 'HKLM:\SOFTWARE\LL\OfficeUpgrade' -Name 'ReinstallSAPAO' -ErrorAction SilentlyContinue
            Remove-ItemProperty -Path 'HKLM:\SOFTWARE\LL\OfficeUpgrade' -Name 'SAPAOArch' -ErrorAction SilentlyContinue
        }
        else {
            Write-ADTLogEntry -Message "No SAP AO pre-upgrade flag present; skipping any SAP AO actions." -Severity 1 -Source $adtSession.InstallPhase
        }
    } catch {
        Write-ADTLogEntry -Message "SAP AO post-check failed: $($_.Exception.Message)" -Severity 2 -Source $adtSession.InstallPhase
    }

    ## End message
    Show-ADTInstallationPrompt -Message 'M365 + Office installation has completed successfully.' -ButtonRightText 'OK' -Icon Information -NoWait -Timeout 5
}

function Uninstall-ADTDeployment
{
    ##================================================
    ## MARK: Pre-Uninstall
    ##================================================
    $adtSession.InstallPhase = "Pre-$($adtSession.DeploymentType)"

    ## Show Progress Message
    Show-ADTInstallationProgress

    ## <Perform Pre-Uninstallation tasks here>
    $processes = @("process1", "process2", "process3")
    foreach ($process in $processes) {
        if (Get-Process -Name $process -ErrorAction SilentlyContinue) {
            Write-ADTLogEntry -Message "Stopping the process '$process'..." -Source $adtSession.InstallPhase
            Stop-Process -Name $process -Force -ErrorAction SilentlyContinue
            Write-ADTLogEntry -Message "'$process' process has been stopped." -Source $adtSession.InstallPhase
        } else {
            Write-ADTLogEntry -Message "'$process' is not running." -Source $adtSession.InstallPhase
        }
    }

    ##================================================
    ## MARK: Uninstall
    ##================================================
    $adtSession.InstallPhase = $adtSession.DeploymentType

    <## Handle Zero-Config MSI uninstallations.
    if ($adtSession.UseDefaultMsi)
    {
        $ExecuteDefaultMSISplat = @{ Action = $adtSession.DeploymentType; FilePath = $adtSession.DefaultMsiFile }
        if ($adtSession.DefaultMstFile)
        {
            $ExecuteDefaultMSISplat.Add('Transform', $adtSession.DefaultMstFile)
        }
        Start-ADTMsiProcess @ExecuteDefaultMSISplat
    }
    #>

    ## <Perform Uninstallation tasks here>
    try {
    $setupExe  = Join-Path $adtSession.DirFiles 'Setup.exe'
    $removeXml = Join-Path $adtSession.DirFiles 'Remove.xml'

    if (-not (Test-Path $setupExe)) {
        Write-ADTLogEntry -Message "Office Setup.exe not found in Files directory." -Severity 3 -Source $adtSession.InstallPhase
        throw "Missing ODT Setup.exe"
    }
    if (-not (Test-Path $removeXml)) {
        Write-ADTLogEntry -Message "Remove.xml not found; skipping Office uninstall." -Severity 2 -Source $adtSession.InstallPhase
        return
    }

    Write-ADTLogEntry -Message "Uninstalling Microsoft 365 Apps using $removeXml" -Severity 1 -Source $adtSession.InstallPhase
    Start-ADTProcess -FilePath $setupExe -ArgumentList "/configure `"$removeXml`"" -WorkingDirectory $adtSession.DirFiles
}
catch {
    Write-ADTLogEntry -Message "Office uninstall failed: $($_.Exception.Message)" -Severity 3 -Source $adtSession.InstallPhase
}

    ##================================================
    ## MARK: Post-Uninstallation
    ##================================================
    $adtSession.InstallPhase = "Post-$($adtSession.DeploymentType)"

    ## <Perform Post-Uninstallation tasks here>
}

function Repair-ADTDeployment
{
    ##================================================
    ## MARK: Pre-Repair
    ##================================================
    $adtSession.InstallPhase = "Pre-$($adtSession.DeploymentType)"
    #Show-ADTInstallationProgress

    ## <Perform Pre-Repair tasks here>

    ##================================================
    ## MARK: Repair
    ##================================================
    $adtSession.InstallPhase = $adtSession.DeploymentType

    <## Handle Zero-Config MSI repairs.
    if ($adtSession.UseDefaultMsi)
    {
        $ExecuteDefaultMSISplat = @{ Action = $adtSession.DeploymentType; FilePath = $adtSession.DefaultMsiFile }
        if ($adtSession.DefaultMstFile)
        {
            $ExecuteDefaultMSISplat.Add('Transform', $adtSession.DefaultMstFile)
        }
        Start-ADTMsiProcess @ExecuteDefaultMSISplat
    }
    #>

    ## <Perform Repair tasks here>

    ##================================================
    ## MARK: Post-Repair
    ##================================================
    $adtSession.InstallPhase = "Post-$($adtSession.DeploymentType)"

    ## <Perform Post-Repair tasks here>
}


##================================================
## MARK: Initialization
##================================================

# Set strict error handling across entire operation.
$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
$ProgressPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
Set-StrictMode -Version 1

# Import the module and instantiate a new session.
try
{
    $moduleName = if ([System.IO.File]::Exists("$PSScriptRoot\PSAppDeployToolkit\PSAppDeployToolkit.psd1"))
    {
        Get-ChildItem -LiteralPath $PSScriptRoot\PSAppDeployToolkit -Recurse -File | Unblock-File -ErrorAction Ignore
        "$PSScriptRoot\PSAppDeployToolkit\PSAppDeployToolkit.psd1"
    }
    else
    {
        'PSAppDeployToolkit'
    }
    Import-Module -FullyQualifiedName @{ ModuleName = $moduleName; Guid = '8c3c366b-8606-4576-9f2d-4051144f7ca2'; ModuleVersion = '4.1.0' } -Force
    try
    {
        $iadtParams = Get-ADTBoundParametersAndDefaultValues -Invocation $MyInvocation
        $adtSession = Open-ADTSession -SessionState $ExecutionContext.SessionState @adtSession @iadtParams -PassThru
    }
    catch
    {
        Remove-Module -Name PSAppDeployToolkit* -Force
        throw
    }
}
catch
{
    $Host.UI.WriteErrorLine((Out-String -InputObject $_ -Width ([System.Int32]::MaxValue)))
    exit 60008
}


##================================================
## MARK: Invocation
##================================================

try
{
    Get-Item -Path $PSScriptRoot\PSAppDeployToolkit.* | & {
        process
        {
            Get-ChildItem -LiteralPath $_.FullName -Recurse -File | Unblock-File -ErrorAction Ignore
            Import-Module -Name $_.FullName -Force
        }
    }
    & "$($adtSession.DeploymentType)-ADTDeployment"
    Close-ADTSession
}
catch
{
    Write-ADTLogEntry -Message ($mainErrorMessage = Resolve-ADTErrorRecord -ErrorRecord $_) -Severity 3
    Show-ADTDialogBox -Text $mainErrorMessage -Icon Stop | Out-Null
    Close-ADTSession -ExitCode 60001
}
finally
{
    Remove-Module -Name PSAppDeployToolkit* -Force
}
# SIG # Begin signature block
# MIIMjgYJKoZIhvcNAQcCoIIMfzCCDHsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU1jBAsvSnOvzGoWf/RWYZkrcL
# Km2gggnvMIIEwDCCA6igAwIBAgIBEzANBgkqhkiG9w0BAQsFADBWMQswCQYDVQQG
# EwJVUzEfMB0GA1UEChMWTUlUIExpbmNvbG4gTGFib3JhdG9yeTEMMAoGA1UECxMD
# UEtJMRgwFgYDVQQDEw9NSVRMTCBSb290IENBLTIwHhcNMTkwNzA4MTExMDAwWhcN
# MjkwNzA4MTExMDAwWjBRMQswCQYDVQQGEwJVUzEfMB0GA1UECgwWTUlUIExpbmNv
# bG4gTGFib3JhdG9yeTEMMAoGA1UECwwDUEtJMRMwEQYDVQQDDApNSVRMTCBDQS02
# MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAj2T0hoZXOA+UPr8SD/Re
# gKGDHDfz+8i1bm+cGV9V2Zxs1XxYrCBbnTB79AtuYR29HIf6HfsUrsqJH6gQtptF
# tux8QrWqx25iOE4tg2yeSVmrc/ZB4fRfufKi0idq2IA13kJgYQ8xCLpIiBEm8be7
# Lzlz9mGT0UVgRe3I5Jku935a7pOB2qHHH6OGWSs9AOPiJdo4oSWUbL5H3H5MmZCI
# 8T3Rj7dobmrRYOsUADI5kkqvOf7o1j09X7X2q4Q+ez4JHgGTLTxjvox7QEDYglZM
# Mh9qB2SGpvhCkKoZ3/05bT1oCt2Pb4iR7MlETNryi/mzZuOjf2gaYpuWweYVh2Ny
# 3wIDAQABo4IBnDCCAZgwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQUk5BH
# A0LBTbQzHtRCl5+h4Ctwv4gwHwYDVR0jBBgwFoAU/8nJZUxTgPGpDDwhroIqx+74
# MvswDgYDVR0PAQH/BAQDAgGGMGcGCCsGAQUFBwEBBFswWTAuBggrBgEFBQcwAoYi
# aHR0cDovL2NybC5sbC5taXQuZWR1L2dldHRvL0xMUkNBMjAnBggrBgEFBQcwAYYb
# aHR0cDovL29jc3AubGwubWl0LmVkdS9vY3NwMDQGA1UdHwQtMCswKaAnoCWGI2h0
# dHA6Ly9jcmwubGwubWl0LmVkdS9nZXRjcmwvTExSQ0EyMIGSBgNVHSAEgYowgYcw
# DQYLKoZIhvcSAgEDAQYwDQYLKoZIhvcSAgEDAQgwDQYLKoZIhvcSAgEDAQcwDQYL
# KoZIhvcSAgEDAQkwDQYLKoZIhvcSAgEDAQowDQYLKoZIhvcSAgEDAQswDQYLKoZI
# hvcSAgEDAQ4wDQYLKoZIhvcSAgEDAQ8wDQYLKoZIhvcSAgEDARAwDQYJKoZIhvcN
# AQELBQADggEBALnwy+yzh/2SvpwC8q8EKdDQW8LxWnDM56DcHm5zgfi0WfEsQi8w
# xcV2Vb2eCNs6j0NofdgsSP7k9DJ6LmDs+dfZEmD23+r9zlMhI6QQcwlvq+cgTrOI
# oUcZd83oyTHr0ig5IFy1r9FpnG00/P5MV+zxmTbTDXJjC8VgxqWl2IhnPk8zr0Fc
# JK0BoYHtv7NHeC4WbNHQZCQf9UMSDALcVR23YZemWizmEK2Mclhjv0E+s7mLZn0A
# K03zCQSvwQrjt+2YzS7J8MxWlRA5cNj1bNbnTtIuEUPpLSYgsN8Q+Ks9ffk9D7yU
# t8No/ntuf6R38t/33c0LTCSJ9AIgjz7hUHMwggUnMIIED6ADAgECAhMwAAW/Xff+
# 6WMO1wIRAAAABb9dMA0GCSqGSIb3DQEBCwUAMFExCzAJBgNVBAYTAlVTMR8wHQYD
# VQQKDBZNSVQgTGluY29sbiBMYWJvcmF0b3J5MQwwCgYDVQQLDANQS0kxEzARBgNV
# BAMMCk1JVExMIENBLTYwHhcNMjQxMDI4MTgxMjU1WhcNMjcxMDI4MTgxMjU1WjBg
# MQswCQYDVQQGEwJVUzEfMB0GA1UEChMWTUlUIExpbmNvbG4gTGFib3JhdG9yeTEO
# MAwGA1UECxMFT3RoZXIxIDAeBgNVBAMTF0lTRCBEZXNrdG9wIEVuZ2luZWVyaW5n
# MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA0BQ5+bMtDvgRT7pCIgHp
# b0iuWsrGHTAKWvKo3T6uk/5r/Kp7VtqJFvcuwLqu0jm+As1kypxloyme0GAKCZcm
# nvyEtRIS5Vxn0FpPO1/y1Bm1JOZ30O7xoy3kimp/16jSmROMeCSdm9qPEmG60M5Y
# L12k7DOaU6/v+5MSZLQiDl20lf34u+Qt8SYNe/L4oA4kdsN3YMXuM6MVbbh6CJzb
# wBT3ceZNwRmkkqQOEQtA0Zr0n2UmoijuraIxU5DC+pISBJIcF3RbfFQNQMivR0lq
# rzQZDrKej/3D9FouGiBl8xZyVtJE0cNum6OE8b7nABtYwKP4jvz3ttxtIWVhoC/v
# WQIDAQABo4IB5zCCAeMwPQYJKwYBBAGCNxUHBDAwLgYmKwYBBAGCNxUIg4PlHYfs
# p2aGrYcVg+rwRYW2oR8dhuHfGoHsg1wCAWQCAQQwFgYDVR0lAQH/BAwwCgYIKwYB
# BQUHAwMwDgYDVR0PAQH/BAQDAgeAMBgGA1UdIAQRMA8wDQYLKoZIhvcSAgEDAQYw
# HQYDVR0OBBYEFLlL4q2UwnJN7ZTZ9W2D+7Y9a+tdMIGCBgNVHREEezB5pFswWTEY
# MBYGA1UEAwwPQW50aG9ueS5NYXNzYXJvMQ8wDQYDVQQLDAZQZW9wbGUxHzAdBgNV
# BAoMFk1JVCBMaW5jb2xuIExhYm9yYXRvcnkxCzAJBgNVBAYTAlVTgRpBbnRob255
# Lk1hc3Nhcm9AbGwubWl0LmVkdTAfBgNVHSMEGDAWgBSTkEcDQsFNtDMe1EKXn6Hg
# K3C/iDAzBgNVHR8ELDAqMCigJqAkhiJodHRwOi8vY3JsLmxsLm1pdC5lZHUvZ2V0
# Y3JsL2xsY2E2MGYGCCsGAQUFBwEBBFowWDAtBggrBgEFBQcwAoYhaHR0cDovL2Ny
# bC5sbC5taXQuZWR1L2dldHRvL2xsY2E2MCcGCCsGAQUFBzABhhtodHRwOi8vb2Nz
# cC5sbC5taXQuZWR1L29jc3AwDQYJKoZIhvcNAQELBQADggEBAFqyP/3MhIsDF2Qu
# ThdPiYz24768PIl64Tiaz8PjjxPnKTiayoOfnCG40wsZh+wlWvZZP5R/6FZab6ZC
# nkrI9IObUZdJeiN4UEypO1v5L6J1iXGq4Zc3QpkJUmjCIIYU0IPG9BPo0SX7mBiz
# DFafAGHReYkovs6vq035+4I6tsOQBpl+JfFPIT37Kpy+PlKz/OXzhVmQOa87mC1b
# YADxWAwwDJd1Mm1GFbXUHHBPkdusW+POqR7qh5WQf0dJpRTsMG/MzIqWiUZxDzkD
# lsqyRl4Y9nN9ii92PGpJF59AZAuEHDX0fqP6yeyMWYZGKpy7XqhQidW7nPxeqHl+
# EQW6EH0xggIJMIICBQIBATBoMFExCzAJBgNVBAYTAlVTMR8wHQYDVQQKDBZNSVQg
# TGluY29sbiBMYWJvcmF0b3J5MQwwCgYDVQQLDANQS0kxEzARBgNVBAMMCk1JVExM
# IENBLTYCEzAABb9d9/7pYw7XAhEAAAAFv10wCQYFKw4DAhoFAKB4MBgGCisGAQQB
# gjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYK
# KwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFN/5Y8uh
# tbFrsi4P9hYPdzOKqt2HMA0GCSqGSIb3DQEBAQUABIIBAAG1qmPV/kKRJ3qJF3ch
# T1ga2p3I83PMEGfJHQoM8odpjWYZ5SA3CGe0vErwAOw3d5SZLrcGrH5Zn7Qs4WMw
# RBTL1HgwQyV7Cmv/lxtpFh7hj1dASNRiIEX8mppH4SdUpCb3/ceq5fkh38s4KbdF
# aaw9yASYRcayLaeXA8OMlOLB03rq6GbCbl9dga5KMql8im40liHlCvOb727vTA7M
# MGsuXChN6zcjVGKaQjOrmOXt2zMr9jQeJdn2fzkwKPuDSNpR34O4fzOg5VxU5iwr
# +RD67IYR+5c9OsnCFzgW1/vOmbt+AXpa+wdPNqnhOKOt3Zc5tOIoFpwMMgopRhpv
# 8do=
# SIG # End signature block
