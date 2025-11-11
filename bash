number of processes whose (
    (name of it = "powershell.exe" OR name of it = "pwsh.exe") AND
    (command line of it contains "LLEA.ps1")
) > 1

number of processes whose (
    (name of it = "powershell.exe" OR name of it = "pwsh.exe") AND
    (
        (command line of it contains "LLEA.ps1") OR
        (command line of it contains "Lincoln Laboratory Endpoint Advisor")
    )
) > 1

exists process whose (
    (name of it = "powershell.exe" OR name of it = "pwsh.exe") AND
    (command line of it contains "LLEA.ps1")
)
