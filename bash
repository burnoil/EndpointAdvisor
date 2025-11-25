##================================================
## MARK: Install
##================================================
$adtSession.InstallPhase = $adtSession.DeploymentType

## Handle Zero-Config MSI installations.
if ($adtSession.UseDefaultMsi)
{
    $ExecuteDefaultMSISplat = @{ Action = $adtSession.DeploymentType; FilePath = $adtSession.DefaultMsiFile }
    if ($adtSession.DefaultMstFile)
    {
        $ExecuteDefaultMSISplat.Add('Transforms', $adtSession.DefaultMstFile)
    }
    Start-ADTMsiProcess @ExecuteDefaultMSISplat
    if ($adtSession.DefaultMspFiles)
    {
        $adtSession.DefaultMspFiles | Start-ADTMsiProcess -Action Patch
    }
}

## <Perform Installation tasks here>

# Build full paths for safety
$officeExe      = Join-Path -Path $adtSession.DirFiles -ChildPath 'setup.exe'
$configXmlPath  = Join-Path -Path $adtSession.DirFiles -ChildPath 'configuration.xml'

Write-ADTLogEntry -Message "Starting Microsoft 365 Apps setup: '$officeExe /configure `"$configXmlPath`"'." 

# Run Office Deployment Tool with a safety timeout
$officeResult = Start-ADTProcess `
    -FilePath $officeExe `
    -ArgumentList @("/configure `"$configXmlPath`"") `
    -WorkingDirectory $adtSession.DirFiles `
    -WaitForMsiExec `
    -WaitForChildProcesses `
    -Timeout (New-TimeSpan -Minutes 90) `
    -TimeoutAction Continue `
    -ExitOnProcessFailure:$false `
    -PassThru

Write-ADTLogEntry -Message "M365 setup.exe completed. ExitCode = [$($officeResult.ExitCode)], TimedOut = [$($officeResult.TimedOut)]."

# --- Post-install detection to see if M365 is actually present ---
$ctrConfigKey = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'

try {
    $ctrProps = Get-ItemProperty -Path $ctrConfigKey -ErrorAction Stop
}
catch {
    Write-ADTLogEntry -Message "Unable to read ClickToRun configuration key at '$ctrConfigKey'. Treating as failure." -Severity 3
    throw "Microsoft 365 Apps installation failed: ClickToRun configuration key not found."
}

if (-not $ctrProps.VersionToReport) {
    Write-ADTLogEntry -Message "ClickToRun key found, but VersionToReport is missing/empty. Treating as failure." -Severity 3
    throw "Microsoft 365 Apps installation failed: VersionToReport not present."
}

Write-ADTLogEntry -Message "Detected Microsoft 365 Apps version [$($ctrProps.VersionToReport)] after setup.exe."

if ($officeResult.TimedOut) {
    # Optional: you can treat timeout as warning if detection passed
    Write-ADTLogEntry -Message "Note: M365 setup.exe hit the timeout, but detection passed. Proceeding." -Severity 2
}

##================================================
## MARK: Post-Install
##================================================
$adtSession.InstallPhase = "Post-$($adtSession.DeploymentType)"

## <Perform Post-Installation tasks here>

## Display a message at the end of the install.
if (-not $adtSession.UseDefaultMsi)
{
    Show-ADTInstallationPrompt -Message 'Installation for Microsoft 365 Apps for Enterprise is complete.' -ButtonRightText 'OK' -Icon Information -NoWait
}
