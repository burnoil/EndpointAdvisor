PS C:\Users\TO33767> powershell.exe -NoProfile -Command "Get-Process -Name 'powershell','pwsh' -ErrorAction SilentlyContinue | Where-Object { try { (Get-WmiObject Win32_Process -Filter \"ProcessId = $($_.Id)\" -ErrorAction SilentlyContinue).CommandLine -like '*LLEA.ps1*' } catch { $false } } | Stop-Process -Force -ErrorAction SilentlyContinue" 
powershell.exe : The string is missing the terminator: ".
At line:1 char:1
+ powershell.exe -NoProfile -Command "Get-Process -Name 'powershell','p ...
+ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : NotSpecified: (The string is missing the terminator: ".:String) [], RemoteException
    + FullyQualifiedErrorId : NativeCommandError
 
    + CategoryInfo          : ParserError: (:) [], ParentContainsErrorRecordException
    + FullyQualifiedErrorId : TerminatorExpectedAtEndOfString
