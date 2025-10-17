# Test web access with different methods
$url = "https://raw.llcad-github.llan.ll.mit.edu/EndpointEngineering/EndpointAdvisor/refs/heads/main/ContentData.json"

# Test 1: WebClient with proxy
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$webClient = New-Object System.Net.WebClient
$webClient.UseDefaultCredentials = $true
$proxy = [System.Net.WebRequest]::GetSystemWebProxy()
$proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
$webClient.Proxy = $proxy
try {
    $content = $webClient.DownloadString($url)
    Write-Host "SUCCESS with WebClient" -ForegroundColor Green
} catch {
    Write-Host "FAILED with WebClient: $($_.Exception.Message)" -ForegroundColor Red
}
$webClient.Dispose()
