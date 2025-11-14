# Standalone AD Account Monitoring Test Script
# =============================================
# Run this script to test AD monitoring functionality before integrating into LLEA
# This script runs in the logged-in user context and retrieves current user's AD info

param(
    [switch]$Detailed,
    [switch]$ShowRawData
)

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  AD Account Monitoring Test Script" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# Function to get AD account status
function Get-ADAccountStatus {
    try {
        # Get current user's domain and username
        $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $userPrincipalName = $currentUser.Name
        
        Write-Host "Current User Identity: $userPrincipalName" -ForegroundColor Green
        
        # Split domain\username
        $domainUser = $userPrincipalName -split '\\'
        if ($domainUser.Count -ne 2) {
            throw "Unable to parse domain and username from: $userPrincipalName"
        }
        
        $domain = $domainUser[0]
        $username = $domainUser[1]
        
        Write-Host "Domain: $domain" -ForegroundColor Green
        Write-Host "Username: $username" -ForegroundColor Green
        Write-Host "`nQuerying Active Directory..." -ForegroundColor Yellow
        
        # Use ADSI to query AD
        $searcher = New-Object System.DirectoryServices.DirectorySearcher
        $searcher.Filter = "(&(objectClass=user)(samAccountName=$username))"
        $searcher.PropertiesToLoad.AddRange(@(
            "accountExpires",
            "pwdLastSet",
            "userAccountControl",
            "msDS-UserPasswordExpiryTimeComputed",
            "distinguishedName",
            "mail",
            "displayName"
        ))
        
        $result = $searcher.FindOne()
        
        if (-not $result) {
            throw "User not found in Active Directory: $username"
        }
        
        Write-Host "User found in AD!" -ForegroundColor Green
        
        # Extract properties
        $properties = @{
            DistinguishedName = if ($result.Properties["distinguishedName"]) { $result.Properties["distinguishedName"][0] } else { "N/A" }
            DisplayName = if ($result.Properties["displayName"]) { $result.Properties["displayName"][0] } else { "N/A" }
            Email = if ($result.Properties["mail"]) { $result.Properties["mail"][0] } else { "N/A" }
        }
        
        if ($Detailed) {
            Write-Host "`nUser Details:" -ForegroundColor Cyan
            Write-Host "  Distinguished Name: $($properties.DistinguishedName)"
            Write-Host "  Display Name: $($properties.DisplayName)"
            Write-Host "  Email: $($properties.Email)"
        }
        
        # Get account expiration
        $accountExpires = $null
        $accountExpiresValue = $result.Properties["accountExpires"][0]
        
        if ($ShowRawData) {
            Write-Host "`nRaw accountExpires value: $accountExpiresValue" -ForegroundColor Gray
        }
        
        if ($accountExpiresValue -and $accountExpiresValue -ne 0 -and $accountExpiresValue -ne 9223372036854775807) {
            $accountExpires = [DateTime]::FromFileTime($accountExpiresValue)
        }
        
        # Get password expiration using computed attribute
        $passwordExpires = $null
        if ($result.Properties.Contains("msDS-UserPasswordExpiryTimeComputed")) {
            $pwdExpiryValue = $result.Properties["msDS-UserPasswordExpiryTimeComputed"][0]
            
            if ($ShowRawData) {
                Write-Host "Raw msDS-UserPasswordExpiryTimeComputed value: $pwdExpiryValue" -ForegroundColor Gray
            }
            
            if ($pwdExpiryValue -and $pwdExpiryValue -ne 0 -and $pwdExpiryValue -ne 9223372036854775807) {
                $passwordExpires = [DateTime]::FromFileTime($pwdExpiryValue)
            }
        }
        
        # Alternative method if computed attribute is not available
        if (-not $passwordExpires) {
            Write-Host "`nComputed password expiry not available, checking pwdLastSet..." -ForegroundColor Yellow
            $pwdLastSetValue = $result.Properties["pwdLastSet"][0]
            
            if ($ShowRawData) {
                Write-Host "Raw pwdLastSet value: $pwdLastSetValue" -ForegroundColor Gray
            }
            
            if ($pwdLastSetValue -and $pwdLastSetValue -ne 0) {
                $pwdLastSet = [DateTime]::FromFileTime($pwdLastSetValue)
                
                if ($Detailed) {
                    Write-Host "Password last set: $($pwdLastSet.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan
                }
                
                # Try to get domain password policy
                Write-Host "Retrieving domain password policy..." -ForegroundColor Yellow
                $maxPasswordAge = 90  # Default
                
                try {
                    $domainRoot = $searcher.SearchRoot.Path
                    $domainEntry = New-Object System.DirectoryServices.DirectoryEntry($domainRoot)
                    $domainSearcher = New-Object System.DirectoryServices.DirectorySearcher($domainEntry)
                    $domainSearcher.Filter = "(objectClass=domainDNS)"
                    $domainSearcher.PropertiesToLoad.Add("maxPwdAge")
                    $domainResult = $domainSearcher.FindOne()
                    
                    if ($domainResult -and $domainResult.Properties["maxPwdAge"]) {
                        $maxPwdAge = $domainResult.Properties["maxPwdAge"][0]
                        # Convert from 100-nanosecond intervals to days
                        $maxPasswordAge = [Math]::Abs($maxPwdAge) / 864000000000
                        Write-Host "Domain password max age: $maxPasswordAge days" -ForegroundColor Green
                    }
                } catch {
                    Write-Host "Could not retrieve domain password policy, using default 90 days" -ForegroundColor Yellow
                }
                
                $passwordExpires = $pwdLastSet.AddDays($maxPasswordAge)
            }
        }
        
        # Check User Account Control flags
        if ($result.Properties["userAccountControl"]) {
            $uac = $result.Properties["userAccountControl"][0]
            if ($Detailed) {
                Write-Host "`nUser Account Control Flags:" -ForegroundColor Cyan
                if ($uac -band 0x10000) { Write-Host "  - Password never expires" -ForegroundColor Yellow }
                if ($uac -band 0x2) { Write-Host "  - Account is disabled" -ForegroundColor Red }
                if ($uac -band 0x10) { Write-Host "  - Account is locked out" -ForegroundColor Red }
            }
            
            # If password never expires flag is set
            if ($uac -band 0x10000) {
                $passwordExpires = $null
            }
        }
        
        return @{
            Success = $true
            Username = "$domain\$username"
            Properties = $properties
            AccountExpires = $accountExpires
            PasswordExpires = $passwordExpires
            UAC = $uac
        }
        
    } catch {
        return @{
            Success = $false
            Error = $_.Exception.Message
        }
    }
}

# Main execution
try {
    $status = Get-ADAccountStatus
    
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "           RESULTS SUMMARY" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    
    if ($status.Success) {
        Write-Host "`nUser: " -NoNewline -ForegroundColor White
        Write-Host $status.Username -ForegroundColor Green
        
        # Account expiration
        Write-Host "`nAccount Expiration:" -ForegroundColor White
        if ($status.AccountExpires) {
            $daysLeft = ($status.AccountExpires - (Get-Date)).Days
            $dateStr = $status.AccountExpires.ToString('yyyy-MM-dd HH:mm:ss')
            
            if ($daysLeft -le 0) {
                Write-Host "  Status: " -NoNewline
                Write-Host "EXPIRED" -ForegroundColor Red -BackgroundColor DarkRed
                Write-Host "  Expired on: $dateStr" -ForegroundColor Red
            } elseif ($daysLeft -eq 1) {
                Write-Host "  Status: " -NoNewline
                Write-Host "EXPIRES TOMORROW!" -ForegroundColor Red
                Write-Host "  Date: $dateStr" -ForegroundColor Yellow
            } elseif ($daysLeft -le 7) {
                Write-Host "  Status: " -NoNewline
                Write-Host "CRITICAL - Expires in $daysLeft days" -ForegroundColor Red
                Write-Host "  Date: $dateStr" -ForegroundColor Yellow
            } elseif ($daysLeft -le 30) {
                Write-Host "  Status: " -NoNewline
                Write-Host "WARNING - Expires in $daysLeft days" -ForegroundColor Yellow
                Write-Host "  Date: $dateStr" -ForegroundColor Yellow
            } else {
                Write-Host "  Status: " -NoNewline
                Write-Host "OK - Expires in $daysLeft days" -ForegroundColor Green
                Write-Host "  Date: $dateStr"
            }
        } else {
            Write-Host "  Status: " -NoNewline
            Write-Host "No expiration set" -ForegroundColor Green
        }
        
        # Password expiration
        Write-Host "`nPassword Expiration:" -ForegroundColor White
        if ($status.PasswordExpires) {
            $daysLeft = ($status.PasswordExpires - (Get-Date)).Days
            $dateStr = $status.PasswordExpires.ToString('yyyy-MM-dd HH:mm:ss')
            
            if ($daysLeft -le 0) {
                Write-Host "  Status: " -NoNewline
                Write-Host "EXPIRED - CHANGE REQUIRED!" -ForegroundColor Red -BackgroundColor DarkRed
                Write-Host "  Expired on: $dateStr" -ForegroundColor Red
                Write-Host "`n  ACTION REQUIRED: " -NoNewline -ForegroundColor Red
                Write-Host "Press Ctrl+Alt+Delete and select 'Change a password'" -ForegroundColor Yellow
            } elseif ($daysLeft -eq 1) {
                Write-Host "  Status: " -NoNewline
                Write-Host "EXPIRES TOMORROW!" -ForegroundColor Red
                Write-Host "  Date: $dateStr" -ForegroundColor Yellow
                Write-Host "`n  RECOMMENDATION: " -NoNewline -ForegroundColor Yellow
                Write-Host "Change your password today to avoid lockout" -ForegroundColor White
            } elseif ($daysLeft -le 7) {
                Write-Host "  Status: " -NoNewline
                Write-Host "CRITICAL - Expires in $daysLeft days" -ForegroundColor Red
                Write-Host "  Date: $dateStr" -ForegroundColor Yellow
                Write-Host "`n  RECOMMENDATION: " -NoNewline -ForegroundColor Yellow
                Write-Host "Change your password soon" -ForegroundColor White
            } elseif ($daysLeft -le 14) {
                Write-Host "  Status: " -NoNewline
                Write-Host "WARNING - Expires in $daysLeft days" -ForegroundColor Yellow
                Write-Host "  Date: $dateStr" -ForegroundColor Yellow
            } else {
                Write-Host "  Status: " -NoNewline
                Write-Host "OK - Expires in $daysLeft days" -ForegroundColor Green
                Write-Host "  Date: $dateStr"
            }
        } else {
            Write-Host "  Status: " -NoNewline
            Write-Host "No expiration (password never expires policy)" -ForegroundColor Green
        }
        
    } else {
        Write-Host "`nERROR: " -NoNewline -ForegroundColor Red
        Write-Host $status.Error -ForegroundColor Yellow
        
        Write-Host "`nPossible causes:" -ForegroundColor Yellow
        Write-Host "  - Computer is not domain-joined"
        Write-Host "  - No connection to domain controller"
        Write-Host "  - Insufficient permissions to query AD"
        Write-Host "  - User is a local account, not domain account"
    }
    
    Write-Host "`n========================================" -ForegroundColor Cyan
    
    # Offer to open password change dialog
    if ($status.Success -and $status.PasswordExpires) {
        $daysLeft = ($status.PasswordExpires - (Get-Date)).Days
        if ($daysLeft -le 14) {
            Write-Host "`nWould you like to change your password now? (Y/N): " -NoNewline -ForegroundColor Yellow
            $response = Read-Host
            if ($response -eq 'Y' -or $response -eq 'y') {
                Write-Host "Opening password change dialog..." -ForegroundColor Green
                try {
                    Start-Process "ms-settings:signinoptions-password"
                } catch {
                    Write-Host "Alternatively, press Ctrl+Alt+Delete and select 'Change a password'" -ForegroundColor Yellow
                }
            }
        }
    }
    
} catch {
    Write-Host "`nUnexpected error: $_" -ForegroundColor Red
}

Write-Host "`nPress any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
