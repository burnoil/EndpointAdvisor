/* Windows 10/11 (MSIX-capable) */
(version of operating system >= "10.0.15063")

AND
(
  /* A) MSI Slack present (either hive) */
  (
    exists keys whose(
      exists value "DisplayName" of it
      and (value "DisplayName" of it as string as lowercase starts with "slack")
      and exists value "WindowsInstaller" of it
      and (value "WindowsInstaller" of it as string = "1")
    ) of keys "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall" of native registry
  )
  OR
  (
    exists keys whose(
      exists value "DisplayName" of it
      and (value "DisplayName" of it as string as lowercase starts with "slack")
      and exists value "WindowsInstaller" of it
      and (value "WindowsInstaller" of it as string = "1")
    ) of keys "HKLM\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" of native registry
  )

  /* B) OR MSIX Slack installed but version < target */
  OR
  (
    exists folder whose (name as lowercase starts with "slacktechnologies.slack_")
      of folder "C:\Program Files\WindowsApps"
    AND
    (
      maximum of
        (
          (preceding text of first "_" of following text of first "_" of (name of it)) as version
        )
      of folders whose (name as lowercase starts with "slacktechnologies.slack_")
      of folder "C:\Program Files\WindowsApps"
    ) < "4.42.90.0" as version
  )
)
