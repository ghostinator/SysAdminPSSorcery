<#
.SYNOPSIS
    Network Connectivity Watchdog - Automated network adapter monitoring and reset tool.

.DESCRIPTION
    A PowerShell script that continuously monitors network connectivity through any specified network adapter,
    automatically resetting the adapter when persistent connectivity issues are detected. The script provides
    real-time monitoring with a dashboard interface showing connectivity status, test results, and adapter statistics.

.PARAMETER AdapterPattern
    The name pattern to identify your network adapter. Supports wildcards.
    Examples:
    - "Ethernet*"     : Matches any adapter starting with "Ethernet"
    - "Wi-Fi*"        : Matches any Wi-Fi adapter
    - "usb_xhci*"     : Matches USB network adapters
    - "*"             : Matches all adapters (will use first found)
    Default: "*"

.PARAMETER FailureThreshold
    Number of seconds to wait after continuous failures before attempting an adapter reset.
    Default: 30 seconds

.PARAMETER TestInterval
    Time in seconds between connectivity tests.
    Default: 5 seconds

.EXAMPLE
    .\NetworkConnectivityWatchdog.ps1
    Runs with default settings, monitoring the first available network adapter.

.EXAMPLE
    .\NetworkConnectivityWatchdog.ps1 -AdapterPattern "Ethernet*" -FailureThreshold 60 -TestInterval 10
    Monitors Ethernet adapters, waits 60 seconds of failures before reset, tests every 10 seconds.

.NOTES
    File Name      : NetworkConnectivityWatchdog.ps1
    Author         : ghostinatr
    Prerequisite   : PowerShell 5.1 or later
    Copyright      : MIT License
    Version        : 1.0.0

.LINK
    https://github.com/ghostinator/SysAdminPSSorcery/NetworkConnectivityWatchdog
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [string]$AdapterPattern = "*",
    
    [Parameter(Mandatory=$false)]
    [int]$FailureThreshold = 30,
    
    [Parameter(Mandatory=$false)]
    [int]$TestInterval = 5
)

# Script version and initialization
$script:Version = "1.0.0"

# Configure console window
Clear-Host
$host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.Size(120, 30)
$host.UI.RawUI.WindowSize = New-Object System.Management.Automation.Host.Size(120, 30)

# Initialize statistics
$script:stats = @{
    StartTime = Get-Date
    TotalResets = 0
    LastReset = $null
    LastSuccess = $null
    CurrentStatus = "Initializing..."
    LastError = ""
    SuccessfulTests = 0
    FailedTests = 0
    FailureStartTime = $null
    ConsecutiveFailures = 0
    TestResults = @{}
}

# Test targets configuration - Can be customized as needed
$script:testTargets = @{
    PingTargets = @(
        @{ Name = "Google DNS"; Address = "8.8.8.8" },
        @{ Name = "Cloudflare DNS"; Address = "1.1.1.1" },
        @{ Name = "Default Gateway"; Address = (Get-NetRoute | 
            Where-Object DestinationPrefix -eq '0.0.0.0/0' | 
            Select-Object -First 1 -ExpandProperty NextHop) }
    )
    DnsTargets = @(
        @{ Name = "Google"; Address = "www.google.com" },
        @{ Name = "Microsoft"; Address = "www.microsoft.com" }
    )
}

function Get-ActiveNetworkAdapter {
    $adapter = Get-NetAdapter | Where-Object {
        $_.Name -like $AdapterPattern -and 
        $_.Status -eq "Up"
    } | Select-Object -First 1
    
    if (-not $adapter) {
        Write-Log "No active adapter found matching pattern: $AdapterPattern" -TestName "Adapter Status" -Success $false
        return $null
    }
    
    Write-Log "Using adapter: $($adapter.Name) ($($adapter.InterfaceDescription))" -TestName "Adapter Status" -Success $true
    return $adapter
}

[Rest of the functions remain the same...]

# Main execution loop
Write-Log "Starting Network Connectivity Watchdog v$script:Version" "Cyan"
Write-Log "Monitoring adapter pattern: $AdapterPattern" "Cyan"
Write-Log "Failure threshold: $FailureThreshold seconds" "Cyan"
Write-Log "Test interval: $TestInterval seconds" "Cyan"

while ($true) {
    $adapter = Get-ActiveNetworkAdapter
    if (-not $adapter) {
        Start-Sleep -Seconds $TestInterval
        continue
    }

    $Connectivity = Test-NetworkConnectivity -AdapterName $adapter.Name

    if ($Connectivity) {
        $script:stats.FailureStartTime = $null
        $script:stats.ConsecutiveFailures = 0
    } 
    else {
        $script:stats.ConsecutiveFailures++
        
        if (-not $script:stats.FailureStartTime) {
            $script:stats.FailureStartTime = Get-Date
        }
        
        $failureDuration = (Get-Date) - $script:stats.FailureStartTime
        
        if ($failureDuration.TotalSeconds -ge $FailureThreshold) {
            Write-Log "Resetting adapter (failing for $([Math]::Round($failureDuration.TotalSeconds))s)" -TestName "Network Status" -Success $false
            
            if (-not (Reset-NetworkAdapter -AdapterName $adapter.Name)) {
                Start-Sleep -Seconds $TestInterval
                continue
            }
            
            $script:stats.FailureStartTime = $null
            $script:stats.ConsecutiveFailures = 0
        }
        else {
            Write-Log "Failing for $([Math]::Round($failureDuration.TotalSeconds))s" -TestName "Network Status" -Success $false
        }
    }

    Start-Sleep -Seconds $TestInterval
}