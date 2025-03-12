<#
.SYNOPSIS
    Continuously monitors network connectivity by pinging devices and testing DNS resolution.

.DESCRIPTION
    This script continuously performs network monitoring by concurrently pinging a list of IP addresses and testing DNS resolution against specified DNS servers.
    It logs the results of each test, including timestamps, ping response times, DNS resolution times, and any errors, to a text file.
    The script is designed to help network administrators quickly assess network health and identify connectivity issues.

.FEATURES
    - Concurrent Ping Tests: Pings multiple IP addresses in parallel for faster results.
    - Concurrent DNS Resolution Tests: Tests multiple DNS servers simultaneously.
    - Detailed Logging: Logs timestamped results to a text file for historical analysis.
    - Real-time Output: Displays current test results in the PowerShell console.
    - Configurable Targets: Easily modify IP addresses, DNS servers, and the test domain within the script.
    - Continuous Monitoring: Runs indefinitely until stopped manually (Ctrl+C).

.CONFIGURATION
    - $ipAddresses: Array of IP addresses to ping. Modify this array to include the devices you want to monitor.
    - $dnsServers: Array of DNS server IP addresses to test against. Update this to reflect your DNS infrastructure.
    - $testDomain: The domain name used for DNS resolution testing (default is "www.google.com").
    - $logFile:  The script automatically creates a log file in the same directory named "network_monitor_YYYYMMDD_HHMMSS.txt".

.HOW TO USE
    1. Save this script as a .ps1 file (e.g., NetworkMonitor.ps1).
    2. Open PowerShell as an Administrator.
    3. Navigate to the script's directory using 'cd'.
    4. Run the script using: '.\NetworkMonitor.ps1'.
    5. Monitor the real-time output in the console and check the log file for detailed results.
    6. Press Ctrl+C to stop the script.

.NOTES
    - Ensure PowerShell execution policy is set to allow script execution if needed (e.g., RemoteSigned).
    - Log files are created in the same directory as the script.
    - Requires PowerShell version 3.0 or later for Job cmdlets.

.EXAMPLE
    .\NetworkMonitor.ps1

#>

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