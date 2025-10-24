## Installation
[PSCustomObject]$adtSession.InstallPhase = 'Installation'

# Install the MSIX package for current user
Add-AppxPackage -Path "$($adtSession.DirFiles)\YourApp.msix"

# OR install for all users (requires admin/system context)
Add-AppxProvisionedPackage -Online -PackagePath "$($adtSession.DirFiles)\YourApp.msix" -SkipLicense

## Uninstallation
[PSCustomObject]$adtSession.InstallPhase = 'Uninstallation'

# Get the package full name
$package = Get-AppxPackage -Name "YourAppPackageName"

# Remove for current user
if ($package) {
    Remove-AppxPackage -Package $package.PackageFullName
}

# OR remove for all users
Remove-AppxPackage -Package $package.PackageFullName -AllUsers
