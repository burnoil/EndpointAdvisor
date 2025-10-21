Start-ADTProcess -FilePath "$dirFiles\Setup.exe" -ArgumentList '/configure Uninstall.xml' -WindowStyle Hidden -CreateNoWindow:$true -WorkingDirectory $dirFiles
