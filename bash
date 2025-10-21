// Define the path to your PSADT executable
parameter "PSADTPath"="{pathname of client folder of current site}\__Download\M365_Office.exe"

// Get the logged-on user
parameter "LoggedOnUser"={name of user of client}

// Create the scheduled task
waithidden cmd /c schtasks /create /tn "M365_PSADT_Install" /tr "\"{parameter "PSADTPath"}\"" /sc once /st 00:00 /ru "{parameter "LoggedOnUser"}" /rl highest /f

// Run the task immediately
waithidden cmd /c schtasks /run /tn "M365_PSADT_Install"

// Wait 5 seconds for the task to start
pause 5000

// Clean up the scheduled task
waithidden cmd /c schtasks /delete /tn "M365_PSADT_Install" /f
