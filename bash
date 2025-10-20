exists key "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" of native registry
AND
exists value "ProductReleaseIds" whose
(
  (it as string) matches (regex "(?i).*(2019|2021|2024)volume.*")
)
of key "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" of native registry
