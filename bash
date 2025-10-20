exists key "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" of native registry
AND exists value "ProductReleaseIds" of key "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" of native registry whose
  (
    it as lowercase contains "2019volume"
    OR it as lowercase contains "2021volume"
    OR it as lowercase contains "2024volume"
  )


