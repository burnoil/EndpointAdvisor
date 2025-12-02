# Silent Temp Cleaner for Windows
$Paths = @(
    "$env:TEMP",
    "$env:SYSTEMROOT\Temp",
    "$env:LOCALAPPDATA\Temp",
    "$env:LOCALAPPDATA\Microsoft\Windows\INetCache",
    "$env:LOCALAPPDATA\Microsoft\Windows\WebCache"
)

foreach ($Path in $Paths) {
    if (Test-Path $Path) {
        Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue | 
            Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
    }
}

# Also empty Recycle Bin silently (optional)
Clear-RecycleBin -Force -ErrorAction SilentlyContinue

Write-Host "Temp files cleaned silently!" -ForegroundColor Green
