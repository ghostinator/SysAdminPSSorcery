<#
.SYNOPSIS
    Updates a Cloudflare DNS record with your dynamic external IP address.

.DESCRIPTION
    This script retrieves your current external IP address and updates a specified
    DNS A or AAAA record in your Cloudflare account. It performs an initial update
    when the script first runs and then monitors for IP address changes, updating
    Cloudflare whenever a change is detected. To prevent excessive updates, it
    communicates with Cloudflare at most once every 24 hours if the IP address
    remains the same.

.NOTES
    Author: ghostinator
    Version: 1.0
    Requires: PowerShell 5.1 or later
    Cloudflare API Token with "DNS:Edit" permission for the target zone.
#>

# Configuration - IMPORTANT: Replace with your actual values
$cloudflareApiToken = "YOUR_CLOUDFLARE_API_TOKEN" # Get this from your Cloudflare dashboard (My Profile -> API Tokens)
$cloudflareZoneId = "YOUR_CLOUDFLARE_ZONE_ID"     # Find this on the "Overview" page of your domain in Cloudflare
$dnsRecordId = "YOUR_DNS_RECORD_ID"             # The ID of the DNS record you want to update (find using Cloudflare API or dashboard)
$dnsRecordName = "your.hostname.com"           # The name of the DNS record (e.g., mydomain.com or subdomain.mydomain.com)
$dnsRecordType = "A"                             # Specify the record type ("A" for IPv4, "AAAA" for IPv6)
$dnsRecordTTL = 120                              # Time to live for the record in seconds
$dnsRecordProxied = $false                         # Set to $true if you want Cloudflare to proxy traffic, otherwise $false

# File paths for storing the last IP and update time
$lastIPFile = "last_ip.txt"
$lastUpdateFile = "last_update.txt"
$dailyInterval = New-TimeSpan -Days 1

# Function to get the current external IP
function Get-ExternalIP {
    try {
        $ip = Invoke-RestMethod -Uri "http://api.ipify.org"
        return $ip
    } catch {
        Write-Host "Error retrieving external IP: $_"
        return $null
    }
}

# Function to update the specified Cloudflare DNS record
function Update-CloudflareDNS {
    param (
        [string]$ip
    )

    $apiUrl = "https://api.cloudflare.com/client/v4/zones/$cloudflareZoneId/dns_records/$dnsRecordId"
    $headers = @{
        "Authorization" = "Bearer $cloudflareApiToken"
        "Content-Type"  = "application/json"
    }
    $body = @{
        type    = $dnsRecordType
        name    = $dnsRecordName
        content = $ip
        ttl     = $dnsRecordTTL
        proxied = $dnsRecordProxied
    } | ConvertTo-Json

    try {
        $response = Invoke-RestMethod -Uri $apiUrl -Method Put -Headers $headers -Body $body
        Write-Host "Cloudflare update response: $($response | ConvertTo-Json -Depth 5)"
        return $true
    } catch {
        Write-Host "Error updating Cloudflare: $_"
        return $false
    }
}

# Main script logic

# Get the current external IP address
$currentIP = Get-ExternalIP

# Perform initial Cloudflare DNS update on the first run
$initialUpdateSuccessful = $false
if ($currentIP) {
    Write-Host "Performing initial Cloudflare DNS update..."
    if (Update-CloudflareDNS -ip $currentIP) {
        $initialUpdateSuccessful = $true
        Set-Content -Path $lastUpdateFile -Value (Get-Date).ToString()
    }
} else {
    Write-Host "Could not retrieve current IP. Exiting."
    exit
}

# Check if the last IP file exists and read the last known IP
if (Test-Path $lastIPFile) {
    $lastIP = Get-Content $lastIPFile
} else {
    $lastIP = ""
}

# Monitor for IP changes in an infinite loop
while ($true) {
    Start-Sleep -Seconds 60 # Check every 60 seconds
    $currentIP = Get-ExternalIP

    if ($currentIP) {
        # Check if the IP has changed
        if ($currentIP -ne $lastIP) {
            Write-Host "IP has changed. New IP: $currentIP. Updating Cloudflare DNS..."
            if (Update-CloudflareDNS -ip $currentIP) {
                Set-Content -Path $lastIPFile -Value $currentIP
                Set-Content -Path $lastUpdateFile -Value (Get-Date).ToString()
            }
        } else {
            Write-Host "IP has not changed. Current IP: $currentIP. Checking for daily update..."
            # Check if it's been more than 24 hours since the last successful update
            if (Test-Path $lastUpdateFile) {
                $lastUpdateTime = Get-Date (Get-Content $lastUpdateFile)
                $timeSinceLastUpdate = New-TimeSpan -Start $lastUpdateTime -End (Get-Date)
                if ($timeSinceLastUpdate -gt $dailyInterval) {
                    Write-Host "It's been more than a day since the last update. Updating Cloudflare DNS..."
                    if (Update-CloudflareDNS -ip $currentIP) {
                        Set-Content -Path $lastUpdateFile -Value (Get-Date).ToString()
                    }
                } else {
                    Write-Host "Last update was on $($lastUpdateTime). Skipping daily update."
                }
            } else {
                # This should ideally not happen after the first successful update, but handle just in case.
                Write-Host "Last update time not found. Performing daily update..."
                if (Update-CloudflareDNS -ip $currentIP) {
                    Set-Content -Path $lastUpdateFile -Value (Get-Date).ToString()
                }
            }
        }
    } else {
        Write-Host "Could not retrieve current IP. Retrying..."
    }
}
