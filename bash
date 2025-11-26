# PSADT 4.1 — get the Files directory (replacement for old $dirFiles)
$dirFiles = $adtSession.ToolkitFilesPath

# Build the PS1 script path
$scriptPath = Join-Path $dirFiles 'MyScript.ps1'

# Run the script using Start-ADTProcess with switches
Start-ADTProcess -FilePath 'pwsh.exe' -Arguments @(
    '-NoLogo'
    '-NoProfile'
    '-File', $scriptPath
    '-Install'
    '-ForceUpdate'
)
