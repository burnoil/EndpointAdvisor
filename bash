# Ultimate Silent Temp Cleaner - Expanded Edition
$Paths = @(
    "$env:TEMP",
    "$env:SYSTEMROOT\Temp",
    "$env:LOCALAPPDATA\Temp",
    "$env:LOCALAPPDATA\Microsoft\Windows\INetCache",
    "$env:LOCALAPPDATA\Microsoft\Windows\WebCache",
    "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache",
    "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Code Cache",
    "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache",
    "$env:SYSTEMROOT\Prefetch",
    "$env:SYSTEMROOT\SoftwareDistribution\Download",
    "$env:LOCALAPPDATA\Microsoft\Windows\Explorer\thumbcache_*.db",
    "$env:APPDATA\Microsoft\Windows\Recent"
    # Add more from the list above as desired
)

foreach ($Path in $Paths) {
    if (Test-Path $Path) {
        Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue | 
            Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
    }
}

# Optional: Empty Recycle Bin + clear Windows Update cache safely
Clear-RecycleBin -Force -ErrorAction SilentlyContinue
# Stop services temporarily if you really want SoftwareDistribution clean (uncomment if needed)
# Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
# Remove-Item "$env:SYSTEMROOT\SoftwareDistribution\Download\*" -Recurse -Force -ErrorAction SilentlyContinue
# Start-Service wuauserv

Write-Host "All common temp locations cleaned!" -ForegroundColor Cyan
