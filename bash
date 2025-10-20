[2025-10-20 10:56:36.540] [Install] [Start-ADTProcess] [Error] :: Execution failed with exit code [-1].
[2025-10-20 10:56:36.657] [Install] [M365_Office.ps1] [Error] :: Error Record:
-------------

Message               : Execution failed with exit code [-1].

FullyQualifiedErrorId : ProcessExitCodeError,Start-ADTProcess
ScriptStackTrace      : at Start-ADTProcess<Process>, Z:\Microsoft\Microsoft 365\M365 and Office\PSAppDeployToolkit\PSAppDeployToolkit.psm1: line 20256
                        at Install-ADTDeployment, Z:\Microsoft\Microsoft 365\M365 and Office\M365_Office.ps1: line 305
                        at <ScriptBlock>, Z:\Microsoft\Microsoft 365\M365 and Office\M365_Office.ps1: line 509
                        at <ScriptBlock>, <No file>: line 1

PositionMessage       : At Z:\Microsoft\Microsoft 365\M365 and Office\M365_Office.ps1:305 char:9
                        +         Start-ADTProcess -Filepath 'Setup.exe' -Argumentlist "/config ...
                        +         ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
[2025-10-20 10:56:36.720] [Install] [Show-ADTDialogBox] [Info] :: Bypassing Show-ADTDialogBox [Mode: Silent]. Text: Error Record:
-------------

Message               : Execution failed with exit code [-1].

FullyQualifiedErrorId : ProcessExitCodeError,Start-ADTProcess
ScriptStackTrace      : at Start-ADTProcess<Process>, Z:\Microsoft\Microsoft 365\M365 and Office\PSAppDeployToolkit\PSAppDeployToolkit.psm1: line 20256
                        at Install-ADTDeployment, Z:\Microsoft\Microsoft 365\M365 and Office\M365_Office.ps1: line 305
                        at <ScriptBlock>, Z:\Microsoft\Microsoft 365\M365 and Office\M365_Office.ps1: line 509
                        at <ScriptBlock>, <No file>: line 1

PositionMessage       : At Z:\Microsoft\Microsoft 365\M365 and Office\M365_Office.ps1:305 char:9
                        +         Start-ADTProcess -Filepath 'Setup.exe' -Argumentlist "/config ...
                        +         ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
[2025-10-20 10:56:36.810] [Finalization] [Close-ADTSession] [Error] :: [Microsoft_365AppsforEnterprise_16.0_x64_EN_01] install completed with exit code [-1].
