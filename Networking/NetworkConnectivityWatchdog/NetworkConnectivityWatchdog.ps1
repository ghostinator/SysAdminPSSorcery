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
    Author         : ghostinator
    Prerequisite   : PowerShell 5.1 or later
    Copyright      : MIT License
    Version        : 1.1.0

.LINK
    https://github.com/ghostinator/SysAdminPSSorcery/tree/main/Networking/NetworkConnectivityWatchdog
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
$script:Version = "1.1.0"

# Configure console window
try {
    $host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.Size(120, 30)
    $host.UI.RawUI.WindowSize = New-Object System.Management.Automation.Host.Size(120, 30)
} catch {
    Write-Warning "Could not set console size. This is not a critical error."
}

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
    LogBuffer = New-Object System.Collections.Generic.List[object]
    MaxLogLines = 20
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

function Write-Log {
    param(
        [string]$Message,
        [string]$TestName = "Info",
        [System.Nullable[bool]]$Success = $true,
        [string]$Color = $null
    )
    
    $timestamp = (Get-Date).ToString('HH:mm:ss')
    $status = if ($Success -eq $true) { "[OK]" } elseif ($Success -eq $false) { "[FAIL]" } else { "[INFO]" }
    
    $colorToUse = $Color
    if (-not $colorToUse) {
        $colorToUse = if ($Success -eq $true) { "Green" } elseif ($Success -eq $false) { "Red" } else { "Yellow" }
    }

    $logEntry = @{
        Timestamp = $timestamp
        Status = $status
        TestName = $TestName
        Message = $Message
        Color = $colorToUse
    }
    
    $script:stats.LogBuffer.Add($logEntry)
    if ($script:stats.LogBuffer.Count -gt $script:stats.MaxLogLines) {
        $script:stats.LogBuffer.RemoveAt(0)
    }

    $script:stats.CurrentStatus = $Message
    if ($Success -eq $true) {
        $script:stats.SuccessfulTests++
        $script:stats.LastSuccess = Get-Date
        $script:stats.ConsecutiveFailures = 0
        $script:stats.FailureStartTime = $null
    } elseif ($Success -eq $false) {
        $script:stats.FailedTests++
        $script:stats.LastError = $Message
        if ($script:stats.ConsecutiveFailures -eq 0) {
            $script:stats.FailureStartTime = Get-Date
        }
        $script:stats.ConsecutiveFailures++
    }
}

function Test-NetworkConnectivity {
    param(
        [string]$AdapterName
    )
    $allPassed = $true
    $results = @{}

    # Ping tests - simplified without source specification
    foreach ($target in $script:testTargets.PingTargets) {
        try {
            $ping = Test-Connection -ComputerName $target.Address -Count 1 -Quiet -ErrorAction Stop
        } catch {
            $ping = $false
        }
        $results[$target.Name] = $ping
        Write-Log "Ping $($target.Name) ($($target.Address))" -TestName "Ping" -Success $ping
        if (-not $ping) { $allPassed = $false }
    }
    
    # DNS tests
    foreach ($target in $script:testTargets.DnsTargets) {
        try {
            $dns = [bool](Resolve-DnsName -Name $target.Address -ErrorAction Stop)
        } catch {
            $dns = $false
        }
        $results[$target.Name] = $dns
        Write-Log "DNS $($target.Name) ($($target.Address))" -TestName "DNS" -Success $dns
        if (-not $dns) { $allPassed = $false }
    }
    $script:stats.TestResults = $results
    return $allPassed
}

function Reset-NetworkAdapter {
    param(
        [string]$AdapterName
    )
    try {
        Disable-NetAdapter -Name $AdapterName -Confirm:$false -ErrorAction Stop
        Start-Sleep -Seconds 3
        Enable-NetAdapter -Name $AdapterName -Confirm:$false -ErrorAction Stop
        $script:stats.TotalResets++
        $script:stats.LastReset = Get-Date
        Write-Log "Adapter $AdapterName reset successfully." -TestName "Adapter Reset" -Success $true
        return $true
    } catch {
        Write-Log "Failed to reset adapter ${AdapterName}: $PSItem" -TestName "Adapter Reset" -Success $false
        return $false
    }
}


function Update-Dashboard {
    Clear-Host
    
    $width = $host.UI.RawUI.WindowSize.Width
    
    # Header
    $uptime = (Get-Date) - $script:stats.StartTime
    $uptimeStr = "{0:d2}:{1:d2}:{2:d2}" -f $uptime.Hours, $uptime.Minutes, $uptime.Seconds
    $lastResetStr = if ($script:stats.LastReset) { $script:stats.LastReset.ToString('HH:mm:ss') } else { "N/A" }
    
    $consecutiveFailureTime = 0
    if ($script:stats.FailureStartTime) {
        $consecutiveFailureTime = [int]((Get-Date) - $script:stats.FailureStartTime).TotalSeconds
    }

    $statusText = "All tests passed"
    $statusColor = "Green"
    if ($script:stats.ConsecutiveFailures -gt 0) {
        $statusText = "FAIL ($($script:stats.ConsecutiveFailures) checks for $($consecutiveFailureTime)s): $($script:stats.LastError)"
        $statusColor = "Red"
    }

    $line1 = " Uptime: {0,-9} | Tests: {1} OK / {2} Fail | Resets: {3} | Last Reset: {4}" -f $uptimeStr, $script:stats.SuccessfulTests, $script:stats.FailedTests, $script:stats.TotalResets, $lastResetStr
    $line2 = " Status: "
    
    Write-Host ("-" * ($width-1)) -ForegroundColor DarkGray
    Write-Host $line1
    Write-Host $line2 -NoNewline
    Write-Host $statusText -ForegroundColor $statusColor
    Write-Host ("-" * ($width-1)) -ForegroundColor DarkGray

    # Log messages
    $script:stats.LogBuffer | ForEach-Object {
        $log = $_
        Write-Host "$($log.Timestamp) $($log.Status) [$($log.TestName)] $($log.Message)" -ForegroundColor $log.Color
    }
}

while ($true) {
    Update-Dashboard
    
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