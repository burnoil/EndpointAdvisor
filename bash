// Prefetch the PSADT package
prefetch M365_PSADT.zip sha1:YOUR_SHA1_HASH_HERE size:YOUR_FILE_SIZE_HERE http://your-bigfix-relay:52311/Uploads/M365_PSADT.zip sha256:YOUR_SHA256_HASH_HERE

// Extract the package to __Download folder
extract M365_PSADT.zip

// Define the path to your PSADT executable
parameter "PSADTPath"="{pathname of client folder of current site}\__Download\M365_Office.exe"

// Create the scheduled task (INTERACTIVE = logged-on user, elevated)
waithidden cmd /c schtasks /create /tn "M365_PSADT_Install" /tr "\"{parameter "PSADTPath"}\"" /sc once /st 00:00 /ru INTERACTIVE /rl highest /z /f

// Run the task immediately
waithidden cmd /c schtasks /run /tn "M365_PSADT_Install"
