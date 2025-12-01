action uses wow64 redirection false

// Ensure marker folder exists
waithidden cmd.exe /c mkdir "C:\ProgramData\BigFix" 2>NUL

// ====== BEGIN: your existing prefetch/extract block ======
// Example structure – replace this with YOUR block:

begin prefetch block
  // your existing prefetch lines
  // prefetch M365_Apps_x64.zip sha1:... size:... url:...
end prefetch block

// If you already do an extract, keep that here:
extract M365_Apps_x64.zip "__Download\M365_Apps_x64"

// ====== END: your existing prefetch/extract block ======

// Optional: quick sanity check – ensure a key file exists after extract
continue if {exists file "__Download\M365_Apps_x64\setup.exe"}

// Create a marker file so we know this endpoint is pre-cached
waithidden cmd.exe /c echo PreCacheComplete > "C:\ProgramData\BigFix\M365_PreCache.done"

// Done – no install is performed here
exit 0



action uses wow64 redirection false

// ====== BEGIN: your existing prefetch/extract block ======
// Use the exact same prefetch lines as Task A
begin prefetch block
  // your existing prefetch lines
  // prefetch M365_Apps_x64.zip sha1:... size:... url:...
end prefetch block

// Same extract line you already have
extract M365_Apps_x64.zip "__Download\M365_Apps_x64"
// ====== END: your existing prefetch/extract block ======

// Now your existing install logic – e.g. one of these:

// (A) If you call setup.exe directly:
waithidden "__Download\M365_Apps_x64\setup.exe" /configure "__Download\M365_Apps_x64\configuration-25H2.xml"

// OR (B) If you use PSADT:
waithidden powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass `
  -File "__Download\M365_Apps_x64\Deploy-Application.ps1" `
  -DeploymentType Install `
  -DeployMode Silent `
  -AllowRebootPassThru

// Capture exit code
parameter "exitCode" = "{exit code of action}"

// Optional: clear the pre-cache marker so it won’t be offered again
// (Only do this if you don’t want repeated installs on failure.)
if {parameter "exitCode" as integer = 0}
  // Success
  // delete "C:\ProgramData\BigFix\M365_PreCache.done"
  exit 0
else
  // Failure – leave marker so you can retry Offer if you want
  exit {parameter "exitCode" as integer}
endif

