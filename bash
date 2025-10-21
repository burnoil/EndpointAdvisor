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
    AppVendor = ''
    AppName = ''
    AppVersion = ''
    AppArch = ''
    AppLang = 'EN'
    AppRevision = '01'
    AppSuccessExitCodes = @(0)
    AppRebootExitCodes = @(1641, 3010)
    AppScriptVersion = '1.0.0'
    AppScriptDate = '2000-12-31'
    AppScriptAuthor = '<author name>'

    # Install Titles (Only set here to override defaults set by the toolkit).
    InstallName = ''
    InstallTitle = ''

    # Script variables.
    DeployAppScriptFriendlyName = $MyInvocation.MyCommand.Name
    DeployAppScriptVersion = '4.0.6'
    DeployAppScriptParameters = $PSBoundParameters
}

##================================================
## MARK: SAP AO Helpers
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
        Write-Log -Message "Test-SAPAOInstalled error: $($_.Exception.Message)" -Severity 3 -Source 'Detect-SAPAO'
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
        Write-Log -Message "SAP AO present ($($state.DisplayName) $($state.Version)); skipping install." -Severity 1 -Source 'Install-SAPAO'
        return
    }

    $exe = Join-Path $dirFiles 'SAPBAO\Setup\NwSapSetup.exe'
    if (-not (Test-Path -LiteralPath $exe)) { throw "SAP AO installer not found: $exe" }

    $prod   = Get-SAPAOProductSwitch -Architecture $Architecture
    $params = "/product=`"$prod`" /Silent"

    Write-Log -Message "Installing SAP AO $Architecture → $exe $params" -Severity 1 -Source 'Install-SAPAO'
    Execute-Process -Path $exe -Parameters $params
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
        Write-Log -Message "Excel COM add-in check skipped/failed: $($_.Exception.Message)" -Severity 2 -Source 'Detect-SAPAO'
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

    ## Show Welcome Message, close Internet Explorer if required, allow up to 3 deferrals, verify there is enough disk space to complete the install, and persist the prompt.
    #Show-ADTInstallationWelcome -CloseProcesses iexplore -AllowDefer -DeferTimes 3 -CheckDiskSpace -PersistPrompt

    ## Show Progress Message (with the default message).
    Show-ADTInstallationProgress

    # --- SAP AO pre-check: if present, flag for post-upgrade reinstall ---
    try {
        $ao = Get-SAPAOState
        if ($ao.Present) {
            Write-Log -Message "Pre-check: SAP AO detected ($($ao.DisplayName) $($ao.Version)); flagging for post-upgrade reinstall." -Severity 1 -Source $installPhase
            New-Item -Path 'HKLM:\SOFTWARE\LL\OfficeUpgrade' -Force | Out-Null
            New-ItemProperty -Path 'HKLM:\SOFTWARE\LL\OfficeUpgrade' -Name 'ReinstallSAPAO' -Value 1 -PropertyType DWord -Force | Out-Null
            if ($ao.Architecture) {
                New-ItemProperty -Path 'HKLM:\SOFTWARE\LL\OfficeUpgrade' -Name 'SAPAOArch' -Value $ao.Architecture -PropertyType String -Force | Out-Null
            }
        } else {
            Write-Log -Message "Pre-check: SAP AO not detected." -Severity 1 -Source $installPhase
        }
    } catch {
        Write-Log -Message "SAP AO pre-check failed: $($_.Exception.Message)" -Severity 2 -Source $installPhase
    }


    ## <Perform Pre-Installation tasks here>
		#Kill Processes.  Add each process name into the processes variable.  Add as many needed.
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

        #Start-ADTMsiProcess -Action 'Install' -FilePath 'googlechromestandaloneenterprise64.msi'
        ## --- ODT (Office Deployment Tool) pre-flight + install (silent UI controlled by XML) ---
    try {
        $setupExe  = Join-Path -Path $dirFiles -ChildPath 'Setup.exe'
        $configXml = Join-Path -Path $dirFiles -ChildPath 'configuration.xml'

        if (-not (Test-Path -LiteralPath $setupExe)) {
            Write-Log -Message "ODT payload missing: $setupExe" -Severity 3 -Source $installPhase
            throw "ODT Setup.exe not found at: $setupExe"
        }
        if (-not (Test-Path -LiteralPath $configXml)) {
            Write-Log -Message "ODT configuration missing: $configXml" -Severity 3 -Source $installPhase
            throw "ODT configuration.xml not found at: $configXml"
        }

        Write-Log -Message "Launching ODT: `"$setupExe`" /configure `"$configXml`"" -Severity 1 -Source $installPhase
        Execute-Process -Path $setupExe -Parameters "/configure `"$configXml`"" -WindowStyle Hidden -CreateNoWindow:$true -WorkingDirectory $dirFiles
    }
    catch {
        Write-Log -Message "ODT install failed: $($_.Exception.Message)" -Severity 3 -Source $installPhase
        try {
            $odtLog = Get-ChildItem -Path (Join-Path $env:WINDIR 'Temp') -Filter 'OfficeDeploymentTool*.log' -ErrorAction SilentlyContinue |
                      Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($odtLog) { Write-Log -Message "Latest ODT log: $($odtLog.FullName)" -Severity 2 -Source $installPhase }
        } catch { }
        throw
    }

    ##================================================
    ## MARK: Post-Install
    ##================================================
    $adtSession.InstallPhase = "Post-$($adtSession.DeploymentType)"

    ## <Perform Post-Installation tasks here>

		#Remove shortcut from All Users Desktop (if found)
		#For Example:  $DESKTOPICONPATH = "C:\Users\Public\Desktop\Google Chrome.lnk"
		$DESKTOPICONPATH = "$envCommonDesktop\Google Chrome.lnk"
		If (Test-Path $DESKTOPICONPATH)
		{
			Remove-Item $DESKTOPICONPATH -Force
			Write-ADTLogEntry -Message "$DESKTOPICONPATH found.  Removing ..." -Source $adtSession.InstallPhase
		}

    ## Display a message at the end of the install.
    #if (!$adtSession.UseDefaultMsi)
    #{
    #    
    # --- SAP AO post-check: install ONLY if it existed pre-upgrade (flag set) ---
    try {
        $needReinstall = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\LL\OfficeUpgrade' -ErrorAction SilentlyContinue).ReinstallSAPAO -eq 1
        $archPref      = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\LL\OfficeUpgrade' -ErrorAction SilentlyContinue).SAPAOArch
        $archToInstall = if ($archPref) { $archPref } else { 'x64' }

        if ($needReinstall) {
            $stateAfter = Get-SAPAOState

            if ($stateAfter.Present) {
                Write-Log -Message "SAP AO was flagged pre-upgrade but is still present post-upgrade ($($stateAfter.DisplayName) $($stateAfter.Version)). Skipping reinstall." -Severity 1 -Source $installPhase
            } else {
                Write-Log -Message "SAP AO was present pre-upgrade and is missing post-upgrade. Installing ($archToInstall)..." -Severity 1 -Source $installPhase
                Install-SAPAOIfNeeded -Architecture $archToInstall -Force
            }

            # cleanup flag
            Remove-ItemProperty -Path 'HKLM:\SOFTWARE\LL\OfficeUpgrade' -Name 'ReinstallSAPAO' -ErrorAction SilentlyContinue
            Remove-ItemProperty -Path 'HKLM:\SOFTWARE\LL\OfficeUpgrade' -Name 'SAPAOArch' -ErrorAction SilentlyContinue
        }
        else {
            Write-Log -Message "No SAP AO pre-upgrade flag present; skipping any SAP AO actions." -Severity 1 -Source $installPhase
        }
    } catch {
        Write-Log -Message "SAP AO post-check failed: $($_.Exception.Message)" -Severity 2 -Source $installPhase
    }

Show-ADTInstallationPrompt -Message 'Installation has completed successfully.' -ButtonRightText 'OK' -Icon Information -NoWait -Timeout 5
    #}
}

function Uninstall-ADTDeployment
{
    ##================================================
    ## MARK: Pre-Uninstall
    ##================================================
    $adtSession.InstallPhase = "Pre-$($adtSession.DeploymentType)"

    ## Show Welcome Message, close Internet Explorer with a 60 second countdown before automatically closing.
    #Show-ADTInstallationWelcome -CloseProcesses iexplore -CloseProcessesCountdown 60

    ## Show Progress Message (with the default message).
    Show-ADTInstallationProgress

    ## <Perform Pre-Uninstallation tasks here>

		#Kill Processes.  Add each process name into the processes variable.  Add as many needed.
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

        Start-ADTMsiProcess -Action 'UnInstall' -FilePath 'googlechromestandaloneenterprise64.msi'
        #Start-ADTProcess -Filepath 'Setup.exe' -Argumentlist '/S /Uninstall'

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

    ## Show Welcome Message, close Internet Explorer with a 60 second countdown before automatically closing.
    #Show-ADTInstallationWelcome -CloseProcesses iexplore -CloseProcessesCountdown 60

    ## Show Progress Message (with the default message).
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
    Import-Module -FullyQualifiedName @{ ModuleName = $moduleName; Guid = '8c3c366b-8606-4576-9f2d-4051144f7ca2'; ModuleVersion = '4.0.6' } -Force
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

