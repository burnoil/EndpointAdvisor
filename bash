function Get-SAPAOState {
    [CmdletBinding()] param()

    # --- Gather uninstall entries from both native and WOW6432Node ---
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
    # Drop any $nulls that snuck in
    $items = $items | Where-Object { $_ }

    # --- Find a matching SAP AO entry by DisplayName/Publisher ---
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

    # --- Build candidate install paths safely (skip nulls) ---
    $pf   = [Environment]::GetFolderPath('ProgramFiles')
    $pf86 = [Environment]::GetFolderPath('ProgramFilesX86')

    $paths = @()
    if ($pf   -and -not [string]::IsNullOrWhiteSpace($pf))   { $paths += (Join-Path -Path $pf   -ChildPath 'SAP BusinessObjects\Office AddIn') }
    if ($pf86 -and -not [string]::IsNullOrWhiteSpace($pf86)) { $paths += (Join-Path -Path $pf86 -ChildPath 'SAP BusinessObjects\Office AddIn') }

    $folderHit = $paths | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1

    # --- Infer architecture from name or folder that hit ---
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
