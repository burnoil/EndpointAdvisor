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
    DeployAppScriptVersion = '4.0.6'
    DeployAppScriptParameters = $PSBoundParameters
}

function Install-ADTDeployment
{
    ##================================================
    ## MARK: Pre-Install
    ##================================================
    $adtSession.InstallPhase = 'Pre-Install'
	Write-ADTLogEntry -Message 'Starting Pre-Install' -Source $adtSession.InstallPhase

    ## Welcome / close apps
	Show-ADTInstallationWelcome `
    -CloseProcesses @{ Name = 'outlook'; Description = 'Microsoft Outlook' }, @{ Name = 'winword'; Description = 'Microsoft Office Word' }, @{ Name = 'excel'; Description = 'Microsoft Office Excel' }, @{ Name = 'powerpnt'; Description = 'Microsoft PowerPoint' }, @{ Name = 'onenote'; Description = 'Microsoft OneNote' } `
    -BlockExecution `
    -CloseProcessesCountdown 600 `
    -PersistPrompt

	## Progress
	Show-ADTInstallationProgress -StatusMessage "Microsoft 365 Apps installation in Progress...`nThis installation may take approximately 20-30 minutes to complete. Please wait..."

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
    $adtSession.InstallPhase = 'Install'
	Write-ADTLogEntry -Message 'Starting Install' -Source $adtSession.InstallPhase

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

        # --- Microsoft 365 Apps (ODT) Install ---
Write-Log -Message "Starting Microsoft 365 Apps installation..." -Source $installPhase

$setupExe  = Join-Path -Path $dirFiles -ChildPath 'Setup.exe'
$configXml = Join-Path -Path $dirFiles -ChildPath 'configuration.xml'

if (-not (Test-Path -LiteralPath $setupExe)) {
    Write-Log -Message "ODT Setup.exe not found at $setupExe" -Severity 3 -Source $installPhase
    Throw "ODT Setup.exe missing"
}

if (-not (Test-Path -LiteralPath $configXml)) {
    Write-Log -Message "ODT configuration.xml not found at $configXml" -Severity 3 -Source $installPhase
    Throw "ODT configuration.xml missing"
}

Execute-Process -Path $setupExe `
                -Parameters "/configure `"$configXml`"" `
                -WorkingDirectory $dirFiles `
                -WindowStyle Hidden `
                -CreateNoWindow:$true

Write-Log -Message "Microsoft 365 Apps installation completed." -Source $installPhase

    ##================================================
    ## MARK: Post-Install
    ##================================================
    $adtSession.InstallPhase = 'Post-Install'
	Write-ADTLogEntry -Message 'Starting Post-Install' -Source $adtSession.InstallPhase

    ## <Perform Post-Installation tasks here>

		#Remove shortcut from All Users Desktop (if found)
		#For Example:  $DESKTOPICONPATH = "C:\Users\Public\Desktop\Google Chrome.lnk"
		$DESKTOPICONPATH = "$envCommonDesktop\NA.lnk"
		If (Test-Path $DESKTOPICONPATH)
		{
			Remove-Item $DESKTOPICONPATH -Force
			Write-ADTLogEntry -Message "$DESKTOPICONPATH found.  Removing ..." -Source $adtSession.InstallPhase
		}

    ## Display a message at the end of the install.
    #if (!$adtSession.UseDefaultMsi)
    #{
    Show-ADTInstallationPrompt -Message 'M365 + Office installation has completed successfully.' -ButtonRightText 'OK' -Icon Information -NoWait -Timeout 5
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
    $adtSession.InstallPhase = 'Uninstall'
Write-ADTLogEntry -Message 'Removing Microsoft 365 Apps via Uninstall.xml' -Source $adtSession.InstallPhase

Start-ADTProcess -FilePath "$dirFiles\Setup.exe" `
                 -ArgumentList '/configure Uninstall.xml' `
                 -WorkingDirectory $dirFiles `
                 -WindowStyle Hidden -CreateNoWindow:$true

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
