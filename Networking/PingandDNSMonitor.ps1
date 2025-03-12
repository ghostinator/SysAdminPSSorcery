# IP addresses to ping
$ipAddresses = @(
    "192.168.2.1",
    "172.16.215.132",
    "172.16.215.134",
    "1.1.1.1",
    "8.8.8.8"  # Google DNS
)

# DNS servers to test resolution with
$dnsServers = @(
    "1.1.1.1",       # Cloudflare
    "8.8.8.8",       # Google
    "172.16.215.132", # Internal DNS server
    "172.16.215.134"  # Internal DNS server
)

# Domain to test DNS resolution
$testDomain = "www.google.com"

# Create a log file in the same directory as the script
$logFile = Join-Path $PSScriptRoot "network_monitor_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
"Network monitoring started at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Out-File -FilePath $logFile

Write-Host "Continuous network monitoring started. Press Ctrl+C to stop."
Write-Host "Results are being logged to: $logFile"

try {
    while ($true) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        # PING TESTS - Start parallel jobs for each IP
        $pingJobs = @()
        foreach ($ip in $ipAddresses) {
            $pingJobs += Start-Job -ScriptBlock {
                param($ip)
                $pingResult = Test-Connection -ComputerName $ip -Count 1 -ErrorAction SilentlyContinue
                if ($pingResult) {
                    return @{IP = $ip; Success = $true; ResponseTime = $pingResult.ResponseTime}
                } else {
                    return @{IP = $ip; Success = $false}
                }
            } -ArgumentList $ip
        }

        # DNS RESOLUTION TESTS - Start parallel jobs for each DNS server
        $dnsJobs = @()
        foreach ($dnsServer in $dnsServers) {
            $dnsJobs += Start-Job -ScriptBlock {
                param($dnsServer, $testDomain)
                $startTime = Get-Date
                try {
                    $result = Resolve-DnsName -Name $testDomain -Server $dnsServer -Type A -ErrorAction Stop
                    $endTime = Get-Date
                    $responseTime = [math]::Round(($endTime - $startTime).TotalMilliseconds)
                    return @{
                        DNSServer = $dnsServer
                        Success = $true
                        ResponseTime = $responseTime
                        ResolvedIP = $result.IPAddress -join ', '
                    }
                } catch {
                    return @{
                        DNSServer = $dnsServer
                        Success = $false
                        Error = $_.Exception.Message
                    }
                }
            } -ArgumentList $dnsServer, $testDomain
        }

        # Wait for all jobs to complete
        Wait-Job -Job ($pingJobs + $dnsJobs) | Out-Null

        # Process ping results
        "[$timestamp] PING RESULTS:" | Out-File -FilePath $logFile -Append
        Write-Host "[$timestamp] PING RESULTS:"

        foreach ($job in $pingJobs) {
            $result = Receive-Job -Job $job
            if ($result.Success) {
                $message = "  Ping to $($result.IP) successful - Response time: $($result.ResponseTime) ms"
            } else {
                $message = "  Ping to $($result.IP) failed"
            }

            # Output to console and log file
            Write-Host $message
            $message | Out-File -FilePath $logFile -Append
        }

        # Process DNS results
        "[$timestamp] DNS RESOLUTION RESULTS:" | Out-File -FilePath $logFile -Append
        Write-Host "[$timestamp] DNS RESOLUTION RESULTS:"

        foreach ($job in $dnsJobs) {
            $result = Receive-Job -Job $job
            if ($result.Success) {
                $message = "  DNS $($result.DNSServer) resolved $testDomain to $($result.ResolvedIP) - Response time: $($result.ResponseTime) ms"
            } else {
                $message = "  DNS $($result.DNSServer) failed to resolve $testDomain - Error: $($result.Error)"
            }

            # Output to console and log file
            Write-Host $message
            $message | Out-File -FilePath $logFile -Append
        }

        # Clean up jobs
        Remove-Job -Job ($pingJobs + $dnsJobs)

        # Add a separator between cycles
        "-" * 80 | Out-File -FilePath $logFile -Append
        Write-Host ("-" * 80)
        Start-Sleep -Seconds 5  # Wait 5 seconds between cycles
    }
} finally {
    "Network monitoring ended at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Out-File -FilePath $logFile -Append
}