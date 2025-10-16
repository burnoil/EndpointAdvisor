function Fetch-ContentData {
    if (-not $config -or [string]::IsNullOrWhiteSpace($config.ContentDataUrl)) {
        Write-Log "ContentDataUrl is not set! Check your Get-DefaultConfig return value." -Level "ERROR"
        $global:FailedFetchAttempts++
        $cachedData = Load-CachedContentData
        if ($cachedData) {
            return $cachedData
        }
        return [PSCustomObject]@{ Data = $defaultContentData; Source = "Default" }
    }
    $url = $config.ContentDataUrl

    try {
        Write-Log "Attempting to fetch content from: $url" -Level "INFO"
        
        # Force TLS 1.2
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
        $response = Invoke-WithRetry -Action {
            $job = Start-Job -ScriptBlock {
                param($url)
                try {
                    # Force TLS 1.2 in job context
                    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                    
                    # Simple Invoke-WebRequest with proxy support
                    $result = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 30 -UseDefaultCredentials -ErrorAction Stop
                    
                    if (-not $result -or -not $result.Content) {
                        throw "Empty response from web request"
                    }
                    
                    return $result.Content
                } catch {
                    # Return error info that can be logged
                    return @{
                        Error = $true
                        Message = $_.Exception.Message
                        Details = $_.Exception.ToString()
                    }
                }
            } -ArgumentList $url
            
            $jobResult = Wait-Job $job -Timeout 35
            if (-not $jobResult) {
                Remove-Job $job -Force
                throw "Web request timed out after 35 seconds."
            }
            
            $result = Receive-Job $job
            Remove-Job $job
            
            # Check if job returned an error
            if ($result -is [hashtable] -and $result.Error) {
                throw "Web request failed: $($result.Message)"
            }
            
            if (-not $result) {
                throw "No response received from web request."
            }
            
            return $result
        } -MaxRetries 3 -RetryDelayMs 500

        Write-Log "Successfully fetched content from remote source." -Level "INFO"
        
        $contentData = $response | ConvertFrom-Json
        Validate-ContentData -Data $contentData
        Write-Log "Content data validated successfully." -Level "INFO"
        Save-CachedContentData -ContentData ([PSCustomObject]@{ Data = $contentData })
        $global:FailedFetchAttempts = 0

        return [PSCustomObject]@{ Data = $contentData; Source = "Remote" }
    }
    catch {
        $global:FailedFetchAttempts++
        Write-Log "Failed to fetch or validate content from $url (Attempt $global:FailedFetchAttempts) - $($_.Exception.Message)" -Level "ERROR"
        
        # Log more details for troubleshooting
        if ($_.Exception.InnerException) {
            Write-Log "Inner exception: $($_.Exception.InnerException.Message)" -Level "ERROR"
        }
        
        if ($global:FailedFetchAttempts -ge 3) {
            Write-Log "Multiple consecutive fetch failures ($global:FailedFetchAttempts). Possible causes: network connectivity, proxy authentication, firewall blocking, or TLS issues." -Level "WARNING"
        }
        
        $cachedData = Load-CachedContentData
        if ($cachedData) {
            Write-Log "Using cached content data as fallback." -Level "INFO"
            return $cachedData
        }
        
        Write-Log "No cached data available, using default content." -Level "WARNING"
        return [PSCustomObject]@{ Data = $defaultContentData; Source = "Default" }
    }
}
