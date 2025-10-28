# Check for provisioned packages
Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -like "*slack*"}

# Check with the exact package name
Get-AppxPackage -AllUsers -Name "com.tinyspeck.slackdesktop"

# List ALL packages and search for slack
Get-AppxPackage -AllUsers | Where-Object {$_.Name -like "*slack*" -or $_.PackageFullName -like "*slack*"}

# Check what's actually in WindowsApps
Get-ChildItem "C:\Program Files\WindowsApps" | Where-Object {$_.Name -like "*slack*"}
