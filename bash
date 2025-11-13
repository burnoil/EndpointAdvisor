[Net.ServicePointManager]::SecurityProtocol = 'Tls12,Tls13'
$files = @(
  'https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LLEA.ps1',
  'https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/DriverUpdate.ps1',
  'https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LL_LOGO.ico',
  'https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/main/LL_LOGO_MSG.ico'
)
$failed = @()
foreach ($url in $files) {
  $filename = Split-Path $url -Leaf
  $dest = "C:\Program Files\LLEA\$filename"
  $ok = $false
  for ($i = 1; $i -le 3; $i++) {
    try {
      Write-Host "Download $filename attempt $i"
      Invoke-WebRequest -Uri $url -OutFile $dest -UseDefaultCredentials -UseBasicParsing -ErrorAction Stop
      $ok = $true
      Write-Host "  Success: $((Get-Item $dest).Length) bytes"
      break
    }} catch {
      Write-Host "  Failed: $($_.Exception.Message)"
      if ($i -lt 3) { Start-Sleep -Seconds ($i * 2) }}
    }}
  }}
  if (-not $ok) { $failed += $filename }}
}}
if ($failed.Count -gt 0) { 
  Write-Host "ERROR: Failed to download: $($failed -join ', ')"
  exit 1 
}}
Write-Host "All downloads successful"
exit 0
