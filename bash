exists key "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" of native registry
AND
exists value "ProductReleaseIds" whose
(
  (it as string contains "2019Volume")
  OR (it as string contains "2021Volume")
  OR (it as string contains "2024Volume")
)
of key "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" of native registry
