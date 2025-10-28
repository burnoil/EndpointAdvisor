# Check if it's a provisioned package
$provisionedPackage = Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -like "*Slack*"}

if ($provisionedPackage) {
    # Remove provisioned package first
    Remove-AppxProvisionedPackage -Online -PackageName $provisionedPackage.PackageName
}

# Then remove from all users
Get-AppxPackage -AllUsers -Name "*Slack*" | Remove-AppxPackage -AllUsers
