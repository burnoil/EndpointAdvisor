##================================================
    ## MARK: Pre-Install
    ##================================================
    $adtSession.InstallPhase = "Pre-$($adtSession.DeploymentType)"

    ## Show Welcome Message, close processes if specified, allow up to 3 deferrals, verify there is enough disk space to complete the install, and persist the prompt.
    $saiwParams = @{
        AllowDefer = $true
        DeferTimes = 3
        CheckDiskSpace = $true
        PersistPrompt = $true
    }
    if ($adtSession.AppProcessesToClose.Count -gt 0)
    {
        $saiwParams.Add('CloseProcesses', $adtSession.AppProcessesToClose)
    }
    ## Welcome / close apps
	Show-ADTInstallationPrompt -Message 'This will upgrade your Office installation to Microsoft 365 Apps for Enterprise. Any MSOffice apps will be forcibly closed if still open. Please save your work before continuing. This prompt will auto-close and proceed in 3 minutes' -ButtonMiddleText 'Proceed' -NoExitOnTimeout -Timeout 180
	
	Show-ADTInstallationWelcome -CloseProcesses @{ Name = 'outlook'; Description = 'Microsoft Outlook' }, @{ Name = 'winword'; Description = 'Microsoft Office Word' }, @{ Name = 'excel'; Description = 'Microsoft Office Excel' }, @{ Name = 'msaccess'; Description = 'Microsoft Access' }, @{ Name = 'powerpnt'; Description = 'Microsoft PowerPoint' }, @{ Name = 'onenote'; Description = 'Microsoft OneNote' } -CloseProcessesCountdown 300 -BlockExecution

	## Progress
	Show-ADTInstallationProgress -StatusMessage "Microsoft 365 Apps for Enterprise installation in Progress...`nThis installation may take approximately 20-45 minutes to complete. Please wait..."

    ## <Perform Pre-Installation tasks here>


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
	Start-ADTProcess -Filepath 'Setup.exe' -Argumentlist "/configure `"$($adtSession.DirFiles)\configuration.xml`""

    ##================================================
    ## MARK: Post-Install
    ##================================================
    $adtSession.InstallPhase = "Post-$($adtSession.DeploymentType)"

    ## <Perform Post-Installation tasks here>


    ## Display a message at the end of the install.
    if (!$adtSession.UseDefaultMsi)
    {
        Show-ADTInstallationPrompt -Message 'Installation for Microsoft 365 Apps for Enterprise is complete.' -ButtonRightText 'OK' -Icon Information -NoWait
    }
}
