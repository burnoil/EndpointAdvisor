## Welcome / close apps
Show-ADTInstallationWelcome `
    -CloseProcesses @{ Name = 'outlook'; Description = 'Microsoft Outlook' }, @{ Name = 'winword'; Description = 'Microsoft Office Word' }, @{ Name = 'excel'; Description = 'Microsoft Office Excel' }, @{ Name = 'powerpnt'; Description = 'Microsoft PowerPoint' }, @{ Name = 'onenote'; Description = 'Microsoft OneNote' } `
    -BlockExecution `
    -CloseProcessesCountdown 600 `
    -PersistPrompt

## Progress
Show-ADTInstallationProgress -StatusMessage "Microsoft 365 Apps installation in Progress...`nThis installation may take approximately 20-30 minutes to complete. Please wait..."
