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

    # setup.exe and configuration.xml are in the Files folder
    $officeExe     = Join-Path -Path $adtSession.DirFiles -ChildPath 'setup.exe'
    $configXmlPath = Join-Path -Path $adtSession.DirFiles -ChildPath 'configuration.xml'

    # Sanity checks so we fail clearly if paths are wrong
    if (-not (Test-Path -Path $officeExe)) {
        Write-Log -Message "M365 setup.exe not found at '$officeExe'." -Severity 3
        throw "M365 setup.exe not found at '$officeExe'."
    }
    if (-not (Test-Path -Path $configXmlPath)) {
        Write-Log -Message "M365 configuration.xml not found at '$configXmlPath'." -Severity 3
        throw "M365 configuration.xml not found at '$configXmlPath'."
    }

    Write-Log -Message "Starting Microsoft 365 Apps setup: '$officeExe /configure `"$configXmlPath`"'." 

    # Run Office Deployment Tool with a safety timeout (60 minutes)
    $officeResult = Start-ADTProcess `
        -FilePath $officeExe `
        -ArgumentList @("/configure `"$configXmlPath`"") `
        -WorkingDirectory $adtSession.DirFiles `
        -WaitForMsiExec `
        -WaitForChildProcesses `
        -Timeout (New-TimeSpan -Minutes 60) `
        -TimeoutAction Continue `
        -ExitOnProcessFailure:$false `
        -PassThru

    Write-Log -Message "M365 setup.exe completed (or timed out). ExitCode = [$($officeResult.ExitCode)], TimedOut = [$($officeResult.TimedOut)]."

    # --- Post-install detection: wait for ClickToRun to show M365 is actually there ---
    $ctrConfigKey = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    $maxMinutes   = 60           # maximum time to wait for detection
    $pollInterval = 30           # seconds
    $maxPolls     = [int](($maxMinutes * 60) / $pollInterval)

    $pollCount       = 0
    $installDetected = $false

    while ($pollCount -lt $maxPolls) {
        try {
            $ctrProps = Get-ItemProperty -Path $ctrConfigKey -ErrorAction Stop

            if ($ctrProps.VersionToReport) {
                Write-Log -Message "M365 installation detected. VersionToReport = [$($ctrProps.VersionToReport)]."
                $installDetected = $true
                break
            }
            else {
                Write-Log -Message "ClickToRun key present but VersionToReport empty; waiting..." -Severity 2
            }
        }
        catch {
            Write-Log -Message "ClickToRun config key not present yet; waiting..." -Severity 2
        }

        Start-Sleep -Seconds $pollInterval
        $pollCount++
        Write-Log -Message "Waiting for M365 detection... ($pollCount/$maxPolls)"
    }

    if (-not $installDetected) {
        Write-Log -Message "Timed out waiting for Microsoft 365 Apps to complete installation." -Severity 3
        throw "Microsoft 365 Apps installation did not complete within the expected time."
    }

    if ($officeResult.TimedOut) {
        Write-Log -Message "Note: setup.exe hit the timeout, but M365 detection passed. Proceeding." -Severity 2
    }

    ##================================================
    ## MARK: Post-Install
    ##================================================
    $adtSession.InstallPhase = "Post-$($adtSession.DeploymentType)"

    ## <Perform Post-Installation tasks here>

    # Close the progress window before showing the completion prompt
    Close-ADTInstallationProgress

    ## Display a message at the end of the install.
    if (!$adtSession.UseDefaultMsi)
    {
        Show-ADTInstallationPrompt -Message 'Installation for Microsoft 365 Apps for Enterprise is complete.' -ButtonRightText 'OK' -Icon Information -NoWait
    }
}
