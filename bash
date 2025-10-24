##*===============================================
##* INSTALLATION
##*===============================================

$adtSession.InstallPhase = $adtSession.DeploymentType

# Your MSIX installation code here
Add-AppxPackage -Path "$($adtSession.DirFiles)\YourApp.msix"

# OR for all users deployment
Add-AppxProvisionedPackage -Online -PackagePath "$($adtSession.DirFiles)\YourApp.msix" -SkipLicense

##*===============================================
##* UNINSTALLATION
##*===============================================

$adtSession.InstallPhase = $adtSession.DeploymentType

# Get and remove the package
$package = Get-AppxPackage -Name "YourAppPackageName*"
if ($package) {
    Remove-AppxPackage -Package $package.PackageFullName -AllUsers
}
