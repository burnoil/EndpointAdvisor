// Step 1: Check if MSI version of Slack is installed
if {exists (it whose (it as string contains "Slack")) of values "DisplayName" of keys of keys "Software\Microsoft\Windows\CurrentVersion\Uninstall" of registry}
    // Step 2: Get the product code of the MSI version
    parameter "ProductCode" = "{(value "ProductCode" of it as string) of keys whose (value "DisplayName" of it as string contains "Slack") of keys "Software\Microsoft\Windows\CurrentVersion\Uninstall" of registry}"
    
    // Step 3: Uninstall the MSI version
    waithidden msiexec.exe /x {parameter "ProductCode"} /qn /norestart /l*v "C:\Windows\Temp\Slack_MSI_Uninstall.log"
endif

// Step 4: Download the MSIx package
prefetch Slack.msix sha1:<sha1_hash> size:<file_size> url:<url_to_msix_package>

// Step 5: Install the MSIx version using PowerShell
waithidden powershell.exe -ExecutionPolicy Bypass -Command "Add-AppxPackage -Path '.\Slack.msix'"

// Step 6: Log the installation
appendfile Installation of Slack MSIx completed on {now}
copy __appendfile "C:\Windows\Temp\Slack_MSIX_Install.log"
