$package = Get-AppxPackage -Name "*Slack*"
		if ($package) {
			Remove-AppxPackage -Package $package.PackageFullName -AllUsers
		}
