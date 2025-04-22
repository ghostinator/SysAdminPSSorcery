# /Users/BrandonCook/VPNTroubleshooterGUI.ps1
# L2TP VPN Troubleshooter
# Version 0.15 "OMFG MY HEAD HURTS EDITION"
# Author: Brandon Cook

# Import required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Advanced Logging System
class VPNLogger {
    [string]$LogPath
    [string]$ErrorLogPath
    [string]$PerformanceLogPath
    [bool]$EnableDebug

    VPNLogger() {
        $this.LogPath = "C:\VPNDiagnostics\vpn.log"
        $this.ErrorLogPath = "C:\VPNDiagnostics\error.log"
        $this.PerformanceLogPath = "C:\VPNDiagnostics\performance.log"
        $this.EnableDebug = $false
        New-Item -ItemType Directory -Force -Path "C:\VPNDiagnostics"
    }

    [void] Log([string]$Message, [string]$Level = "INFO") {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logMessage = "[$timestamp] [$Level] $Message"
        
        switch ($Level.ToUpper()) {
            "ERROR" { 
                Add-Content -Path $this.ErrorLogPath -Value $logMessage
                Write-Host $Message -ForegroundColor Red
            }
            "WARN" { 
                Add-Content -Path $this.LogPath -Value $logMessage
                Write-Host $Message -ForegroundColor Yellow
            }
            default { 
                Add-Content -Path $this.LogPath -Value $logMessage
                Write-Host $Message -ForegroundColor Green
            }
        }
    }

    [hashtable] AnalyzeLogs() {
        $analysis = @{
            ErrorCount        = (Get-Content $this.ErrorLogPath).Count
            WarningCount      = (Get-Content $this.LogPath | Where-Object { $_ -match "\[WARN\]" }).Count
            CommonErrors      = @{}
            PerformanceIssues = @{}
        }

        Get-Content $this.ErrorLogPath | ForEach-Object {
            if ($_ -match "\[ERROR\] (.*)") {
                $errorMessage = $matches[1]
                if (!$analysis.CommonErrors[$errorMessage]) {
                    $analysis.CommonErrors[$errorMessage] = 0
                }
                $analysis.CommonErrors[$errorMessage]++
            }
        }

        return $analysis
    }
}

# VPN Prerequisites Conclusion Class
class VPNPrerequisitesConclusion {
    [string]$Status
    [System.Collections.ArrayList]$Issues
    [System.Collections.ArrayList]$Recommendations
    [string]$Summary

    VPNPrerequisitesConclusion() {
        $this.Status = "Ready"
        $this.Issues = [System.Collections.ArrayList]::new()
        $this.Recommendations = [System.Collections.ArrayList]::new()
        $this.Summary = ""
    }

    [void] AddIssue([string]$issue, [string]$recommendation) {
        $this.Issues.Add($issue)
        $this.Recommendations.Add($recommendation)
        $this.Status = "Not Ready"
    }

    [string] GenerateSummary() {
        if ($this.Status -eq "Ready") {
            $this.Summary = "✓ All VPN prerequisites are met. Your system is ready for VPN connections."
        } else {
            $this.Summary = "⚠ VPN Prerequisites Check Failed`n`n"
            $this.Summary += "Issues Found:`n"
            for ($i = 0; $i -lt $this.Issues.Count; $i++) {
                $this.Summary += "• $($this.Issues[$i])`n"
            }
            $this.Summary += "`nRecommended Actions:`n"
            for ($i = 0; $i -lt $this.Recommendations.Count; $i++) {
                $this.Summary += "$($i + 1). $($this.Recommendations[$i])`n"
            }
        }
        return $this.Summary
    }
}

# Initialize Logger
$Logger = [VPNLogger]::new()
function Start-VPNDiagnostics {
    param([string]$VPNName)    
    # Define public DNS servers to test connectivity
    $publicDnsServers = @(
        @{Name = "Google DNS Primary"; IP = "8.8.8.8"},
        @{Name = "Google DNS Secondary"; IP = "8.8.4.4"},
        @{Name = "Cloudflare Primary"; IP = "1.1.1.1"},
        @{Name = "Cloudflare Secondary"; IP = "1.0.0.1"},
        @{Name = "OpenDNS Primary"; IP = "208.67.222.222"},
        @{Name = "OpenDNS Secondary"; IP = "208.67.220.220"},
        @{Name = "Quad9"; IP = "9.9.9.9"},
        @{Name = "AdGuard DNS"; IP = "94.140.14.14"},
        @{Name = "CleanBrowsing"; IP = "185.228.168.168"},
        @{Name = "Verisign"; IP = "64.6.64.6"}
    )    

    
    $diagnosticResults = @()
    
    # Check VPN Connection
    try {
        $vpnConnection = Get-VpnConnection -Name $VPNName -ErrorAction Stop
        $diagnosticResults += @{
            Component = "VPN Configuration"
            Status    = "Found"
            Details   = "Type: $($vpnConnection.TunnelType), Server: $($vpnConnection.ServerAddress)"
        }
    }
    catch {
        $diagnosticResults += @{
            Component = "VPN Configuration"
            Status    = "Not Found"
            Details   = "No VPN connection named '$VPNName' exists"
        }
    }

    # Check VPN Client Service
    $services = @{
        'RasMan' = 'Remote Access Connection Manager (Required for VPN Client)'
    }

    foreach ($service in $services.GetEnumerator()) {
        $status = Get-Service -Name $service.Key -ErrorAction SilentlyContinue
        $diagnosticResults += @{
            Component = $service.Value
            Status    = i        # Test VPN Server Connectivity            
            Details   = "Service State: $($status.Status), Start Type: $($status.StartType)"
        }
    }

    # Check WAN Miniports
    $mini
        # Test General Internet Connectivity using Public DNS Servers
        $reachableDns = 0
        $totalTestedDns = 0
        $bestLatency = [int]::MaxValue
        $bestDns = ""

        foreach ($dns in $publicDnsServers) {
            $dnsTest = Test-Connection -ComputerName $dns.IP -Count 1 -ErrorAction SilentlyContinue
            $totalTestedDns++
            
            if ($dnsTest) {
                $reachableDns++
                if ($dnsTest.ResponseTime -lt $bestLatency) {
                    $bestLatency = $dnsTest.ResponseTime
                    $bestDns = $dns.Name
                }
            }
        }

        $diagnosticResults += @{
            Component = "Internet Connectivity"
            Status = if ($reachableDns -gt 0) { "Connected" } else { "No Connection" }
            Details = if ($reachableDns -gt 0) {
                "$reachableDns out of $totalTestedDns DNS servers reachable. Best response: $bestDns ($bestLatency ms)"
            } else {
                "Unable to reach any public DNS servers. Check internet connection."
            }
        }    
        $diagnosticResults += @{
        Component = "WAN Miniports"
        Status    = if ($miniports) { "Found" } else { "Missing" }
        Details   = "Found $($miniports.Count) WAN Miniport devices"
    }

    # Check Network Connectivity
    if ($vpnConnection) {
        $pingTest = Test-Connection -ComputerName $vpnConnection.ServerAddress -Count 1 -ErrorAction SilentlyContinue
        $diagnosticResults += @{
            Component = "Server Connectivity"
            Status    = if ($pingTest) { "Reachable" } else { "Unreachable" }
            Details   = if ($pingTest) { "Latency: $($pingTest.ResponseTime)ms" } else { "Unable to reach server" }
        }

        # Check if required ports are accessible (not blocked by firewall)
        $ports = @(
            @{Port = 500; Purpose = "IKE (Internet Key Exchange)"},
            @{Port = 1701; Purpose = "L2TP"},
            @{Port = 4500; Purpose = "IPSec NAT Traversal"}
        )
        
        foreach ($portInfo in $ports) {
            $portTest = Test-NetConnection -ComputerName $vpnConnection.ServerAddress -Port $portInfo.Port -WarningAction SilentlyContinue
            $diagnosticResults += @{
                Component = "Port $($portInfo.Port)"
                Status    = if ($portTest.TcpTestSucceeded) { "Accessible" } else { "Blocked" }
                Details   = "$($portInfo.Purpose) - Check firewall rules if blocked"
            }
        }
    }

    # Check Client Registry Settings
    $registryPaths = @{
        "HKLM:\SYSTEM\CurrentControlSet\Services\RasMan\Parameters" = "AllowL2TPWeakCrypto"
    }

    foreach ($path in $registryPaths.GetEnumerator()) {
        $value = Get-ItemProperty -Path $path.Key -Name $path.Value -ErrorAction SilentlyContinue
        $diagnosticResults += @{
            Component = "Registry: $($path.Value)"
            Status    = if ($value) { "Configured" } else { "Missing" }
            Details   = "Path: $($path.Key)"
        }
    }

    # Analyze results and provide conclusion
    $conclusion = [VPNPrerequisitesConclusion]::new()
    
    # Check VPN Configuration
    $vpnConfig = $diagnosticResults | Where-Object { $_.Component -eq "VPN Configuration" }
    if ($vpnConfig.Status -eq "Not Found") {
        $conclusion.AddIssue(
            "VPN connection not configured",
            "Create a new L2TP VPN connection with the correct server address"
        )
    }

    # Check Required Client Service
    $rasmanService = $diagnosticResults | Where-Object { 
        $_.Component -eq "Remote Access Connection Manager" -and $_.Status -ne "Running"
    }
    if ($rasmanService) {
        $conclusion.AddIssue(
            "VPN Client service not running",
            "Enable and start the Remote Access Connection Manager (RasMan) service: Open Services app, find 'Remote Access Connection Manager', set Startup type to Automatic and click Start"
        )
    }

    # Check WAN Miniports
    $miniports = $diagnosticResults | Where-Object { $_.Component -eq "WAN Miniports" }
    if ($miniports.Status -eq "Missing") {
        $conclusion.AddIssue(
            "WAN Miniport devices missing",
            "Use the 'Reset Network Devices' option to reinstall WAN Miniports"
        )
    }

    # Check Internet Connectivity
    $internetCheck = $diagnosticResults | Where-Object { $_.Component -eq "Internet Connectivity" }
    if ($internetCheck.Status -eq "No Connection") {
        $conclusion.AddIssue(
            "No internet connectivity detected",
            "Check your internet connection and try connecting to different networks (e.g., mobile hotspot)"
        )
    }

    # Check Server Connectivity
    $connectivity = $diagnosticResults | Where-Object { $_.Component -eq "Server Connectivity" }
    if ($connectivity -and $connectivity.Status -eq "Unreachable") {
        $conclusion.AddIssue(
            "Cannot reach VPN server",
            "Verify the VPN server address is correct and the server is operational"
        )
    }

    # Check Port Accessibility
    $blockedPorts = $diagnosticResults | Where-Object { 
        $_.Component -like "Port *" -and $_.Status -eq "Blocked"
    }
    if ($blockedPorts) {
        $portList = $blockedPorts.Component -join ', '
        $conclusion.AddIssue(
            "Required ports are blocked: $portList",
            "Check Windows Firewall and antivirus settings to ensure ports 500, 1701, and 4500 are allowed"
        )
    }

    # Generate and add conclusion to results
    $diagnosticResults += @{
        Component = "Prerequisites Conclusion"
        Status = $conclusion.Status
        Details = $conclusion.GenerateSummary()
    }

    return $diagnosticResults
}
function Reset-NetworkDevices {
    try {
        # Show instructions to user
        $message = @"
Manual Steps Required:
1. Device Manager will open
2. Expand 'Network adapters'
3. Right-click and uninstall each 'WAN Miniport' device (check 'Delete driver' if available)
4. After removing all WAN Miniports, click 'Scan for hardware changes' in Device Manager
5. Click OK to open Device Manager

Note: Windows will automatically reinstall the WAN Miniport devices.
"@
        [System.Windows.Forms.MessageBox]::Show($message, "Manual Device Removal Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        # Open Device Manager
        Start-Process "devmgmt.msc"
        
        # Wait for user confirmation
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Click OK once you've completed the WAN Mini-port removal steps.",
            "Confirm Completion",
            [System.Windows.Forms.MessageBoxButtons]::OKCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($result -eq 'OK') {
            $Logger.Log("User confirmed manual device removal. Proceeding with network stack reset...", "INFO")
            
            # Reset network stack
            $commands = @(
                "netsh winsock reset catalog",
                "netsh int ipv4 reset reset.log",
                "netsh int ipv6 reset reset.log",
                "ipconfig /release",
                "ipconfig /renew",
                "ipconfig /flushdns"
            )
            
            foreach ($cmd in $commands) {
                Update-Status "Executing command: $cmd"
                $result = cmd.exe /c $cmd 2>&1
                if ($result) {
                    Update-Status "Result: $result"
                }
                Start-Sleep -Seconds 1  # Add slight delay to make commands more visible
            }

            # Stop and restart networking services
            $services = @('RasMan', 'RemoteAccess', 'PolicyAgent', 'IKEEXT')
            foreach ($service in $services) {
                Stop-Service -Name $service -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 2
                Start-Service -Name $service -ErrorAction SilentlyContinue
                $Logger.Log("Restarted service: $service", "INFO")
            }
            
            # Notify user about required restart
            [System.Windows.Forms.MessageBox]::Show(
                "Network stack reset completed. A system restart is required for changes to take effect.",
                "Restart Required",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            
            $Logger.Log("Network devices reset completed successfully. Restart pending.", "INFO")
            return $true
        }
        else {
            $Logger.Log("User cancelled the operation", "INFO")
            return $false
        }
    }
    catch {
        $Logger.Log("Failed to reset network devices: $_", "ERROR")
        return $false
    }
}
function Reset-NetworkStack {
    param([switch]$IncludeMiniports)
    
    try {
        # Create restore point
        Checkpoint-Computer -Description "Before VPN Network Reset" -RestorePointType "MODIFY_SETTINGS"

        # Stop required services
        $services = @('RasMan', 'RemoteAccess', 'PolicyAgent')
        foreach ($service in $services) {
            Stop-Service -Name $service -Force
        }

        if ($IncludeMiniports) {
            # Remove WAN Miniport devices
            Get-PnpDevice | Where-Object { $_.FriendlyName -like "*WAN Miniport*" } | ForEach-Object {
                $_ | Disable-PnpDevice -Confirm:$false
                $_ | Remove-PnpDevice -Confirm:$false
            }
        }

        # Reset network stack
        $commands = @(
            "netsh winsock reset",
            "netsh int ip reset",
            "ipconfig /release",
            "ipconfig /renew",
            "ipconfig /flushdns"
        )

        foreach ($command in $commands) {
            Invoke-Expression $command
        }

        if ($IncludeMiniports) {
            # Scan for hardware changes
            Start-Process "pnputil.exe" -ArgumentList "/scan-devices" -Wait -NoNewWindow
            Start-Sleep -Seconds 15  # Wait for devices to initialize
        }

        # Restart services
        foreach ($service in $services) {
            Start-Service -Name $service
        }

        return $true
    }
    catch {
        $Logger.Log("Network stack reset failed: $_", "ERROR")
        return $false
    }
}
# Backup Function
function Backup-VPNConfiguration {
    param(
        [string]$VPNName
    )
    
    $backupDate = Get-Date -Format "yyyy-MM-dd_HH-mm"
    $backupDir = "C:\VPNBackup_$backupDate"
    
    try {
        # Create backup directory
        New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
        
        # Try to export both user-specific and all-user VPN connection details
        $vpnConnections = @()
        
        # Try user-specific connection
        try {
            $userVPN = Get-VpnConnection -Name $VPNName -ErrorAction SilentlyContinue
            if ($userVPN) {
                $vpnConnections += @{
                    Connection = $userVPN
                    Type = "User"
                }
            }
        } catch { }
        
        # Try all-user connection
        try {
            $allUserVPN = Get-VpnConnection -Name $VPNName -AllUserConnection -ErrorAction SilentlyContinue
            if ($allUserVPN) {
                $vpnConnections += @{
                    Connection = $allUserVPN
                    Type = "AllUsers"
                }
            }
        } catch { }
        
        if ($vpnConnections.Count -gt 0) {
            $vpnConnections | ConvertTo-Json | Out-File "$backupDir\vpn_connections.json"
            $Logger.Log("VPN connection details backed up", "INFO")
        }
        
        # Export relevant registry settings
        $registryPaths = @(
            "HKLM:\SYSTEM\CurrentControlSet\Services\RasMan",
            "HKLM:\SYSTEM\CurrentControlSet\Services\Rasl2tp"
        )
        
        foreach ($path in $registryPaths) {
            $regName = ($path -split '\\')[-1]
            reg export ($path -replace 'HKLM:\\', 'HKLM\') "$backupDir\${regName}.reg" /y
        }
        
        $Logger.Log("Backup created successfully in $backupDir", "INFO")
        return $true
    }
    catch {
        $Logger.Log("Backup failed: $($_.Exception.Message)", "ERROR")
        return $false
    }
}
# Cleanup Function
function Remove-ExistingVPNConfiguration {
    param([string]$VPNName)
    
    try {
        # Remove VPN connections
        if ($VPNName) {
            Remove-VpnConnection -Name $VPNName -Force -ErrorAction Stop
            $Logger.Log("Removed VPN connection: $VPNName")
        }
        else {
            Get-VpnConnection | ForEach-Object {
                Remove-VpnConnection -Name $_.Name -Force -ErrorAction Stop
                $Logger.Log("Removed VPN connection: $($_.Name)")
            }
        }
        
        # Stop and restart services
        $services = @('RasMan', 'RemoteAccess', 'PolicyAgent')
        foreach ($service in $services) {
            Stop-Service -Name $service -Force
            Start-Service -Name $service
            $Logger.Log("Restarted service: $service")
        }
        
        # Reset network stack
        $commands = @(
            "netsh winsock reset",
            "netsh int ip reset",
            "ipconfig /release",
            "ipconfig /renew",
            "ipconfig /flushdns"
        )
        
        foreach ($command in $commands) {
            Invoke-Expression $command
            $Logger.Log("Executed: $command")
        }
        
        $Logger.Log("VPN configuration cleanup completed successfully")
        return $true
    }
    catch {
        $Logger.Log("Cleanup failed: $_", "ERROR")
        return $false
    }
}

# Create New VPN Connection
function New-VPNConnection {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        
        [Parameter(Mandatory = $true)]
        [string]$ServerAddress,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet("PSK", "Certificate")]
        [string]$AuthType,
        
        [Parameter(Mandatory = $false)]
        [string]$PreSharedKey,
        
        [Parameter(Mandatory = $false)]
        [string]$CertificateThumbprint,
        
        [Parameter(Mandatory = $false)]
        [string]$Username,
        
        [Parameter(Mandatory = $false)]
        [string]$Domain,

        [Parameter(Mandatory = $false)]
        [bool]$SplitTunneling = $false,

        [Parameter(Mandatory = $false)]
        [bool]$RememberCredential = $true,

        [Parameter(Mandatory = $false)]
        [ValidateSet("Required", "Optional", "NoEncryption", "Maximum")]
        [string]$EncryptionLevel = "Required",

        [Parameter(Mandatory = $false)]
        [ValidateSet("MSChapv2", "EAP", "PAP", "CHAP")]
        [string]$AuthenticationMethod = "PAP",

        [Parameter(Mandatory = $false)]
        [bool]$IPv4Only = $false,

        [Parameter(Mandatory = $false)]
        [bool]$UseWinlogonCredential = $false,

        [Parameter(Mandatory = $false)]
        [int]$IdleDisconnectSeconds = 0
    )
    
    try {
        $vpnParams = @{
            Name                 = $Name
            ServerAddress        = $ServerAddress
            TunnelType           = "L2tp"
            EncryptionLevel      = $EncryptionLevel
            AuthenticationMethod = $AuthenticationMethod
            RememberCredential   = $RememberCredential
            Force                = $true
            PassThru             = $true
        }

        # Add authentication based on type
        if ($AuthType -eq "PSK") {
            if ([string]::IsNullOrEmpty($PreSharedKey)) {
                throw "Pre-shared key is required when using PSK authentication"
            }
            $vpnParams.L2tpPsk = $PreSharedKey
        }
        else {
            if ([string]::IsNullOrEmpty($CertificateThumbprint)) {
                throw "Certificate thumbprint is required when using Certificate authentication"
            }
            $vpnParams.EapConfigXmlStream = [String]::Empty
            $vpnParams.UseWinlogonCredential = $UseWinlogonCredential
        }
        
        $connection = Add-VpnConnection @vpnParams -ErrorAction Stop
        
        # Configure additional settings
        Set-VpnConnection -Name $Name `
            -SplitTunneling $SplitTunneling `
            -UseWinlogonCredential $UseWinlogonCredential

        # Set idle disconnect timeout
        if ($IdleDisconnectSeconds -gt 0) {
            Set-VpnConnection -Name $Name -IdleDisconnectSeconds $IdleDisconnectSeconds
        }

        # Configure IPv4/IPv6 settings
        if ($IPv4Only) {
            # Disable IPv6 for this connection using registry
            $connectionPath = "HKLM:\SYSTEM\CurrentControlSet\Services\RasMan\Parameters\Config\$Name"
            if (!(Test-Path $connectionPath)) {
                New-Item -Path $connectionPath -Force | Out-Null
            }
            Set-ItemProperty -Path $connectionPath -Name "IPv6" -Value 0 -Type DWord
            
            # Update connection properties
            Set-VpnConnection -Name $Name `
                -SplitTunneling $SplitTunneling `
                -UseWinlogonCredential $UseWinlogonCredential `
                -RememberCredential $RememberCredential
        }

        # Configure IPsec settings if using PSK
        if ($AuthType -eq "PSK") {
            $connection | Set-VpnConnectionIPsecConfiguration -AuthenticationTransformConstants GCMAES256 `
                -CipherTransformConstants GCMAES256 -EncryptionMethod AES256 -IntegrityCheckMethod SHA256 `
                -DHGroup Group14 -PfsGroup PFS2048 -Force
        }
        
        $Logger.Log("VPN connection '$Name' created successfully", "INFO")
        return $true
    }
    catch {
        $Logger.Log("Failed to create VPN connection: $_", "ERROR")
        return $false
    }
}


# GUI Implementation
function Show-VPNTroubleshooterGUI {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'L2TP VPN Troubleshooter'
    $form.Size = New-Object System.Drawing.Size(800, 700)
    $form.StartPosition = 'CenterScreen'
    $form.Activate()       # Forces the window to activate

    # Create Status TextBox first
    $statusTextBox = New-Object System.Windows.Forms.TextBox
    $statusTextBox.Location = New-Object System.Drawing.Point(10, 620)
    $statusTextBox.Size = New-Object System.Drawing.Size(760, 40)
    $statusTextBox.Multiline = $true
    $statusTextBox.ScrollBars = 'Vertical'
    $statusTextBox.ReadOnly = $true
    $form.Controls.Add($statusTextBox)

    # Define Update-Status function early
    function Update-Status {
        param($Message)
        $timestamp = Get-Date -Format "HH:mm:ss"
        $statusTextBox.AppendText("[$timestamp] $Message`r`n")
        $statusTextBox.ScrollToCaret()
        if ($diagResults) {
            $diagResults.AppendText("[$timestamp] $Message`r`n")
            $diagResults.ScrollToCaret()
        }
        $Logger.Log($Message, "INFO")
        [System.Windows.Forms.Application]::DoEvents()
    }

    # Create TabControl
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Size = New-Object System.Drawing.Size(770, 600)
    $tabControl.Location = New-Object System.Drawing.Point(10, 10)

    # Create tabs
    $tabPages = @{
        "Configuration" = New-Object System.Windows.Forms.TabPage
        "Diagnostics"   = New-Object System.Windows.Forms.TabPage
        "Logs"          = New-Object System.Windows.Forms.TabPage
    }

    foreach ($tab in $tabPages.GetEnumerator()) {
        $tab.Value.Text = $tab.Key
        $tabControl.Controls.Add($tab.Value)
    }

    # Configuration Tab
    $configTab = $tabPages["Configuration"]
    
    # VPN Connection Manager Group
    $vpnManagerGroup = New-Object System.Windows.Forms.GroupBox
    $vpnManagerGroup.Text = "VPN Connection Manager"
    $vpnManagerGroup.Location = New-Object System.Drawing.Point(10, 20)
    $vpnManagerGroup.Size = New-Object System.Drawing.Size(350, 200)
    $configTab.Controls.Add($vpnManagerGroup)

    # VPN List
    $vpnListBox = New-Object System.Windows.Forms.ListBox
    $vpnListBox.Location = New-Object System.Drawing.Point(10, 20)
    $vpnListBox.Size = New-Object System.Drawing.Size(330, 130)
    $vpnManagerGroup.Controls.Add($vpnListBox)

    # Function to refresh VPN list
    function Update-VPNList {
        try {
            $vpnListBox.Items.Clear()
            
            # Try multiple methods to find VPN connections
            $vpns = @()
            
            # Method 1: Standard VPN connections
            $vpns += Get-VpnConnection -ErrorAction SilentlyContinue | Select-Object Name
            
            # Method 2: All-user VPN connections
            $vpns += Get-VpnConnection -AllUserConnection -ErrorAction SilentlyContinue | Select-Object Name
            
            # Method 3: Check RAS phone book
            $rasPhoneBook = "$env:APPDATA\Microsoft\Network\Connections\Pbk\rasphone.pbk"
            if (Test-Path $rasPhoneBook) {
                $content = Get-Content $rasPhoneBook
                $content | Select-String '^\[(.+)\]$' | ForEach-Object {
                    $vpnName = $_.Matches[0].Groups[1].Value
                    if ($vpnName -ne "Version") {
                        $vpns += [PSCustomObject]@{ Name = $vpnName }
                    }
                }
            }

            # Method 4: Check Registry locations
            $regPaths = @(
                "HKLM:\SYSTEM\CurrentControlSet\Services\RasMan\Config",
                "HKLM:\SOFTWARE\Microsoft\RAS AutoDial\Addresses",
                "HKCU:\Software\Microsoft\RasPhonebook",
                "HKLM:\SYSTEM\CurrentControlSet\Services\RemoteAccess\Performance\Connections"
            )
            
            foreach ($path in $regPaths) {
                if (Test-Path $path) {
                    Get-ChildItem $path -ErrorAction SilentlyContinue | ForEach-Object {
                        $vpnName = $_.PSChildName
                        if ($vpnName -and ($vpns | Where-Object Name -eq $vpnName).Count -eq 0) {
                            $vpns += [PSCustomObject]@{ Name = $vpnName }
                        }
                    }
                }
            }
            
            # Add unique VPN connections to the list box
            $vpns | Sort-Object Name -Unique | ForEach-Object {
                $vpnListBox.Items.Add($_.Name)
            }
            
            if ($vpnListBox.Items.Count -eq 0) {
                Update-Status "No VPN connections found"
            }
            else {
                Update-Status "Found $($vpnListBox.Items.Count) VPN connection(s)"
            }
        }
        catch {
            Update-Status "Error updating VPN list: $_"
        }
    }

    # VPN Management Buttons
    $vpnManageButtons = @(
        # Replace the View Settings action with:
        @{
            'Text' = 'View Settings'
            'Location' = New-Object System.Drawing.Point(10,160)
            'Action' = {
                if ($vpnListBox.SelectedItem) {
                    try {
                        $vpnName = $vpnListBox.SelectedItem
                        $vpnConfig = $null

                        # Method 1: Try standard VPN connection
                        try {
                            $vpnConfig = Get-VpnConnection -Name $vpnName -ErrorAction Stop
                            Update-Status "Found VPN in standard connections"
                        }
                        catch {
                            try {
                                $vpnConfig = Get-VpnConnection -Name $vpnName -AllUserConnection -ErrorAction Stop
                                Update-Status "Found VPN in all-user connections"
                            }
                            catch {
                                Update-Status "Standard VPN info not available, checking other sources..."
                            }
                        }

                        # Method 2: Check RAS phone book
                        if (-not $vpnConfig) {
                            $rasPhoneBook = "$env:APPDATA\Microsoft\Network\Connections\Pbk\rasphone.pbk"
                            if (Test-Path $rasPhoneBook) {
                                $content = Get-Content $rasPhoneBook
                                $section = $content | Select-String -Pattern "(?ms)\[$vpnName\](.*?)(\[|$)" | 
                                    ForEach-Object { $_.Matches.Groups[1].Value }
                                
                                if ($section) {
                                    Update-Status "Found VPN in RAS phone book"
                                    $serverAddress = [regex]::Match($section, 'PhoneNumber=(.+)').Groups[1].Value
                                    $vpnType = [regex]::Match($section, 'VpnStrategy=(.+)').Groups[1].Value
                                    $authProtocol = [regex]::Match($section, 'AuthProtocol=(.+)').Groups[1].Value
                                    $encryption = [regex]::Match($section, 'EncryptionType=(.+)').Groups[1].Value

                                    $vpnConfig = [PSCustomObject]@{
                                        Name = $vpnName
                                        ServerAddress = $serverAddress
                                        TunnelType = switch($vpnType) {
                                            "5" { "L2TP" }
                                            "6" { "PPTP" }
                                            "7" { "SSTP" }
                                            "8" { "IKEv2" }
                                            default { "Unknown" }
                                        }
                                        AuthenticationMethod = switch($authProtocol) {
                                            "0" { "None" }
                                            "1" { "PAP" }
                                            "2" { "CHAP" }
                                            "3" { "MSCHAPv2" }
                                            "4" { "EAP" }
                                            default { "Unknown" }
                                        }
                                        Source = "RAS Phone Book"
                                    }
                                }
                            }
                        }

                        # Method 3: Check Registry
                        if (-not $vpnConfig) {
                            $regPaths = @(
                                "HKLM:\SYSTEM\CurrentControlSet\Services\RasMan\Config\$vpnName",
                                "HKLM:\SOFTWARE\Microsoft\RAS AutoDial\Addresses\$vpnName",
                                "HKCU:\Software\Microsoft\RasPhonebook\$vpnName"
                            )

                            foreach ($path in $regPaths) {
                                if (Test-Path $path) {
                                    $regInfo = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
                                    if ($regInfo) {
                                        Update-Status "Found VPN in registry: $path"
                                        $vpnConfig = [PSCustomObject]@{
                                            Name = $vpnName
                                            ServerAddress = $regInfo.PhoneNumber
                                            TunnelType = "L2TP"
                                            AuthenticationMethod = "Unknown"
                                            Source = "Registry: $($path.Split('\')[-2])"
                                        }
                                        break
                                    }
                                }
                            }
                        }

                        if ($vpnConfig) {
                            # Populate Basic Settings
                            $inputControls['VPN Name'].Text = $vpnConfig.Name
                            $inputControls['Server Address'].Text = $vpnConfig.ServerAddress
                            
                            # Set Authentication Type
                            $inputControls['Authentication Type'].SelectedItem = 'PSK'  # Default to PSK
                            
                            # Populate Advanced Settings where available
                            if ($vpnConfig.AuthenticationMethod -ne "Unknown") {
                                $inputControls['Authentication Method'].SelectedItem = $vpnConfig.AuthenticationMethod
                            }
                            if ($vpnConfig.EncryptionLevel) {
                                $inputControls['Encryption Level'].SelectedItem = $vpnConfig.EncryptionLevel
                            }
                            if ($vpnConfig.SplitTunneling -is [bool]) {
                                $inputControls['Split Tunneling'].Checked = $vpnConfig.SplitTunneling
                            }
                            if ($vpnConfig.RememberCredential -is [bool]) {
                                $inputControls['Remember Credential'].Checked = $vpnConfig.RememberCredential
                            }
                            if ($vpnConfig.UseWinlogonCredential -is [bool]) {
                                $inputControls['Use Winlogon Credential'].Checked = $vpnConfig.UseWinlogonCredential
                            }

                            # Switch to Configuration tab
                            $tabControl.SelectedTab = $tabPages["Configuration"]
                            
                            Update-Status "Loaded settings for VPN connection: $vpnName (Source: $($vpnConfig.Source))"
                        } else {
                            [System.Windows.Forms.MessageBox]::Show(
                                "Could not find VPN connection settings in any location.",
                                "Settings Unavailable",
                                [System.Windows.Forms.MessageBoxButtons]::OK,
                                [System.Windows.Forms.MessageBoxIcon]::Warning
                            )
                        }
                    }
                    catch {
                        Update-Status "Error retrieving VPN settings: $_"
                    }
                }
            }
        }

        # Delete action:
        @{
            'Text' = 'Delete'
            'Location' = New-Object System.Drawing.Point(90,160)
            'Action' = {
                if ($vpnListBox.SelectedItem) {
                    $vpnName = $vpnListBox.SelectedItem
                    $result = [System.Windows.Forms.MessageBox]::Show(
                        "Are you sure you want to delete the VPN connection '$vpnName'?",
                        "Confirm Delete",
                        [System.Windows.Forms.MessageBoxButtons]::YesNo,
                        [System.Windows.Forms.MessageBoxIcon]::Warning
                    )
                    if ($result -eq 'Yes') {
                        try {
                            $deletionAttempted = $false
                            $deletionSuccessful = $false

                            # Try standard removal
                            try {
                                Remove-VpnConnection -Name $vpnName -Force -AllUserConnection -ErrorAction SilentlyContinue
                                Remove-VpnConnection -Name $vpnName -Force -ErrorAction SilentlyContinue
                                $deletionAttempted = $true
                                Update-Status "Attempted standard VPN removal"
                            } catch { }

                            # Remove from RAS phone book
                            $rasPhoneBook = "$env:APPDATA\Microsoft\Network\Connections\Pbk\rasphone.pbk"
                            if (Test-Path $rasPhoneBook) {
                                $content = Get-Content $rasPhoneBook
                                $newContent = $content | Where-Object { $_ -notmatch "^\[$vpnName\]" }
                                if ($content.Count -ne $newContent.Count) {
                                    $newContent | Set-Content $rasPhoneBook
                                    $deletionAttempted = $true
                                    Update-Status "Removed from RAS phone book"
                                }
                            }

                            # Remove from Registry
                            $regPaths = @(
                                "HKLM:\SYSTEM\CurrentControlSet\Services\RasMan\Config\$vpnName",
                                "HKLM:\SOFTWARE\Microsoft\RAS AutoDial\Addresses\$vpnName",
                                "HKCU:\Software\Microsoft\RasPhonebook\$vpnName"
                            )

                            foreach ($path in $regPaths) {
                                if (Test-Path $path) {
                                    Remove-Item -Path $path -Force -Recurse -ErrorAction SilentlyContinue
                                    $deletionAttempted = $true
                                    Update-Status "Removed from registry path: $path"
                                }
                            }

                            if ($deletionAttempted) {
                                Update-Status "VPN connection '$vpnName' deletion attempted from all locations"
                                Start-Sleep -Seconds 2  # Give system time to process changes
                                Update-VPNList
                            } else {
                                Update-Status "No deletion performed - VPN not found in any location"
                            }
                        }
                        catch {
                            Update-Status "Error during VPN deletion: $_"
                        }
                    }
                }
            }
        },    
        @{
            'Text'     = 'Refresh'
            'Location' = New-Object System.Drawing.Point(170, 160)
            'Action'   = { Update-VPNList }
        }
    )
    foreach ($button in $vpnManageButtons) {
        $btn = New-Object System.Windows.Forms.Button
        $btn.Location = $button.Location
        $btn.Size = New-Object System.Drawing.Size(75,23)
        $btn.Text = $button.Text
        $btn.Add_Click($button.Action)
        $vpnManagerGroup.Controls.Add($btn)
    }

        # Call initial VPN list update
        Update-VPNList

        # New VPN Configuration Group
        $newVPNGroup = New-Object System.Windows.Forms.GroupBox
        $newVPNGroup.Text = "New VPN Configuration"
        $newVPNGroup.Location = New-Object System.Drawing.Point(10, 230)
        $newVPNGroup.Size = New-Object System.Drawing.Size(740, 400)
        $configTab.Controls.Add($newVPNGroup)

        # Basic Settings Group
        $basicSettingsGroup = New-Object System.Windows.Forms.GroupBox
        $basicSettingsGroup.Text = "Basic Settings"
        $basicSettingsGroup.Location = New-Object System.Drawing.Point(10, 20)
        $basicSettingsGroup.Size = New-Object System.Drawing.Size(350, 210)
        $newVPNGroup.Controls.Add($basicSettingsGroup)

        # Basic Input Fields 
        $basicFields = @{
            'VPN Name'               = @{Required = $true }
            'Server Address'         = @{Required = $true }
            'Username'               = @{Required = $true }
            'Authentication Type'    = @{Required = $true; Type = 'ComboBox'; Options = @('PSK', 'Certificate') }
            'Pre-shared Key'         = @{Required = $false; Password = $true }
            'Certificate Thumbprint' = @{Required = $false }
            'Domain (Optional)'      = @{Required = $false }
        }

        $yPosition = 30
        $inputControls = @{}

        foreach ($field in $basicFields.GetEnumerator()) {
            $label = New-Object System.Windows.Forms.Label
            $label.Location = New-Object System.Drawing.Point(10, $yPosition)
            $label.Size = New-Object System.Drawing.Size(120, 20)
            $label.Text = $field.Key + $(if ($field.Value.Required) { "*" } else { "" })
            $basicSettingsGroup.Controls.Add($label)

            if ($field.Value.Type -eq 'ComboBox') {
                $control = New-Object System.Windows.Forms.ComboBox
                $control.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
                foreach ($option in $field.Value.Options) {
                    $control.Items.Add($option)
                }
                $control.SelectedIndex = 0
            }
            else {
                $control = New-Object System.Windows.Forms.TextBox
                if ($field.Value.Password) {
                    $control.PasswordChar = '*'
                }
            }
        
            $control.Location = New-Object System.Drawing.Point(160, $yPosition)
            $control.Size = New-Object System.Drawing.Size(180, 20)
            $basicSettingsGroup.Controls.Add($control)
            $inputControls[$field.Key] = $control

            $yPosition += 30
        }

        # Advanced Settings Group
        $advancedSettingsGroup = New-Object System.Windows.Forms.GroupBox
        $advancedSettingsGroup.Text = "Advanced Settings"
        $advancedSettingsGroup.Location = New-Object System.Drawing.Point(370, 20)
        $advancedSettingsGroup.Size = New-Object System.Drawing.Size(350, 280)
        $newVPNGroup.Controls.Add($advancedSettingsGroup)

        # Advanced Settings Controls
        $advancedSettings = @{
            'Split Tunneling'           = @{ Type = 'CheckBox'; Default = $false }
            'Remember Credential'       = @{ Type = 'CheckBox'; Default = $true }
            'Authentication Method'     = @{ 
                Type    = 'ComboBox'
                Options = @('MSChapv2', 'EAP', 'PAP', 'CHAP')
                Default = 'PAP'
            }
            'Encryption Level'          = @{
                Type    = 'ComboBox'
                Options = @('Required', 'Optional', 'NoEncryption', 'Maximum')
                Default = 'Optional'
            }
            'IPv4 Only'                 = @{ Type = 'CheckBox'; Default = $false }
            'Use Winlogon Credential'   = @{ Type = 'CheckBox'; Default = $false }
            'Idle Disconnect (seconds)' = @{ Type = 'TextBox'; Default = '0' }
        }

        $yPos = 30
        foreach ($setting in $advancedSettings.GetEnumerator()) {
            $label = New-Object System.Windows.Forms.Label
            $label.Location = New-Object System.Drawing.Point(10, $yPos)
            $label.Size = New-Object System.Drawing.Size(150, 20)
            $label.Text = $setting.Key
            $advancedSettingsGroup.Controls.Add($label)

            switch ($setting.Value.Type) {
                'CheckBox' {
                    $control = New-Object System.Windows.Forms.CheckBox
                    $control.Location = New-Object System.Drawing.Point(160, $yPos)
                    $control.Size = New-Object System.Drawing.Size(180, 20)
                    $control.Checked = $setting.Value.Default
                }
                'ComboBox' {
                    $control = New-Object System.Windows.Forms.ComboBox
                    $control.Location = New-Object System.Drawing.Point(160, $yPos)
                    $control.Size = New-Object System.Drawing.Size(180, 20)
                    $control.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
                    foreach ($option in $setting.Value.Options) {
                        $control.Items.Add($option)
                    }
                    $control.SelectedItem = $setting.Value.Default
                }
                'TextBox' {
                    $control = New-Object System.Windows.Forms.TextBox
                    $control.Location = New-Object System.Drawing.Point(160, $yPos)
                    $control.Size = New-Object System.Drawing.Size(180, 20)
                    $control.Text = $setting.Value.Default
                }
            }
            $advancedSettingsGroup.Controls.Add($control)
            $inputControls[$setting.Key] = $control
            $yPos += 30
        }
        # Button Container for Create/Test buttons
        $buttonContainer = New-Object System.Windows.Forms.Panel
        $buttonContainer.Location = New-Object System.Drawing.Point(10, 230)
        $buttonContainer.Size = New-Object System.Drawing.Size(720, 40)
        $newVPNGroup.Controls.Add($buttonContainer)

        # Create New VPN Button
        $createButton = New-Object System.Windows.Forms.Button
        $createButton.Location = New-Object System.Drawing.Point(0, 0)
        $createButton.Size = New-Object System.Drawing.Size(150, 30)
        $createButton.Text = "Save VPN"
        $createButton.Add_Click({
                # Validate required fields
                $missingFields = $basicFields.GetEnumerator() | 
                Where-Object { $_.Value.Required -and [string]::IsNullOrWhiteSpace($inputControls[$_.Key].Text) } |
                Select-Object -ExpandProperty Key

                if ($missingFields) {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Please fill in all required fields:`n$($missingFields -join "`n")",
                        "Missing Information",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Warning
                    )
                    return
                }

                # Validate auth-specific fields
                $authType = $inputControls['Authentication Type'].SelectedItem
                if ($authType -eq 'PSK' -and [string]::IsNullOrWhiteSpace($inputControls['Pre-shared Key'].Text)) {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Pre-shared Key is required for PSK authentication",
                        "Missing Information",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Warning
                    )
                    return
                }

                Update-Status "Creating new VPN connection..."
    
                # Gather all VPN parameters including advanced settings
                # In the Create New VPN Button click handler, modify the vpnParams section:
                $vpnParams = @{
                    Name                  = $inputControls['VPN Name'].Text
                    ServerAddress         = $inputControls['Server Address'].Text
                    AuthType              = $authType
                    Username              = $inputControls['Username'].Text
                    Domain                = $inputControls['Domain (Optional)'].Text
                    SplitTunneling        = $inputControls['Split Tunneling'].Checked
                    RememberCredential    = $inputControls['Remember Credential'].Checked
                    AuthenticationMethod  = $inputControls['Authentication Method'].SelectedItem
                    EncryptionLevel       = $inputControls['Encryption Level'].SelectedItem
                    IPv4Only              = $inputControls['IPv4 Only'].Checked
                    UseWinlogonCredential = $inputControls['Use Winlogon Credential'].Checked
                }

                # Add authentication specific parameters
                if ($authType -eq 'PSK') {
                    $vpnParams.PreSharedKey = $inputControls['Pre-shared Key'].Text
                }
                else {
                    $vpnParams.CertificateThumbprint = $inputControls['Certificate Thumbprint'].Text
                }

                # Add idle disconnect if specified
                if ([int]::TryParse($inputControls['Idle Disconnect (seconds)'].Text, [ref]$null)) {
                    $vpnParams.IdleDisconnectSeconds = [int]$inputControls['Idle Disconnect (seconds)'].Text
                }

                $result = New-VPNConnection @vpnParams
                if ($result) {
                    Update-Status "VPN created successfully"
                    Update-VPNList
                }
            })
        $buttonContainer.Controls.Add($createButton)

          
        # Diagnostics Tab
        $diagTab = $tabPages["Diagnostics"]
    
        # Network Analysis
        $diagButtons = @(
            @{
                Text   = 'Network Stack Info'
                Action = {
                    Update-Status "Gathering Network Stack Information..."

                    Update-Status "`nNetwork Adapters:"
                    Get-NetAdapter | ForEach-Object {
                        Update-Status "  $($_.Name):"
                        Update-Status "    Status: $($_.Status)"
                        Update-Status "    MAC Address: $($_.MacAddress)"
                        Update-Status "    Speed: $($_.LinkSpeed)"
                    }

                    Update-Status "`nIP Configuration:"
                    $ipConfigs = Get-NetIPConfiguration
                    foreach ($cfg in $ipConfigs) {
                        Update-Status "  Interface: $($cfg.InterfaceAlias)"
                        $ip = $cfg.IPv4Address.IPAddress
                        $gw = $cfg.IPv4DefaultGateway.NextHop
                        $dns = $cfg.DNSServer.ServerAddresses
                        Update-Status "    IPv4: $ip"
                        Update-Status "    Gateway: $gw"
                        Update-Status "    DNS: $($dns -join ', ')"
                        # IP Analysis
                        if ($ip -match '^(10\.|192\.168\.|172\.(1[6-9]|2[0-9]|3[01])\.)') {
                            $ipType = 'Private'
                        } else {
                            $ipType = 'Public'
                        }
                        if ($gw) {
                            $ipParts = $ip -split '\.'
                            $gwParts = $gw -split '\.'
                            if ($ipParts[0..2] -join '.' -eq $gwParts[0..2] -join '.') {
                                $gwRel = 'Gateway is in the same subnet as the IP address.'
                            } else {
                                $gwRel = 'Gateway is NOT in the same subnet as the IP address.'
                            }
                        } else {
                            $gwRel = 'No gateway set.'
                        }
                        Update-Status "      [Conclusion] IP is $ipType. $gwRel"
                        # DNS Analysis
                        $publicDns = @('8.8.8.8','8.8.4.4','1.1.1.1','9.9.9.9')
                        $dnsType = if ($dns | Where-Object { $publicDns -contains $_ }) { 'Public DNS detected.' } else { 'No well-known public DNS detected.' }
                        Update-Status "      [Conclusion] $dnsType"
                    }

                    Update-Status "`nRouting Table:"
                    $routes = Get-NetRoute -Protocol NetMgmt
                    foreach ($route in $routes) {
                        Update-Status "  $($route.DestinationPrefix) via $($route.NextHop)"
                    }
                    # Route Analysis
                    $defaultRoute = $routes | Where-Object { $_.DestinationPrefix -eq '0.0.0.0/0' } | Select-Object -First 1
                    if ($defaultRoute) {
                        Update-Status "    [Conclusion] Default route points to $($defaultRoute.NextHop). This is the gateway for all outbound traffic."
                    } else {
                        Update-Status "    [Conclusion] No default route (0.0.0.0/0) found. Internet access may not be possible."
                    }
                }
            }
            
            @{
                'Text'   = 'Network Tests'
                'Action' = {
                    # Create test selection form
                    $testForm = New-Object System.Windows.Forms.Form
                    $testForm.Text = "Network Test Configuration"
                    $testForm.Size = New-Object System.Drawing.Size(500,400)
                    $testForm.StartPosition = 'CenterParent'
                    $testForm.FormBorderStyle = 'FixedDialog'
                    $testForm.MaximizeBox = $false
                    $testForm.MinimizeBox = $false

                    # Target Selection Group
                    $targetGroup = New-Object System.Windows.Forms.GroupBox
                    $targetGroup.Text = "Test Target"
                    $targetGroup.Location = New-Object System.Drawing.Point(10,10)
                    $targetGroup.Size = New-Object System.Drawing.Size(460,120)
                    $testForm.Controls.Add($targetGroup)

                    # VPN Radio Button
                    $vpnRadio = New-Object System.Windows.Forms.RadioButton
                    $vpnRadio.Location = New-Object System.Drawing.Point(10,20)
                    $vpnRadio.Size = New-Object System.Drawing.Size(150,20)
                    $vpnRadio.Text = "VPN Connection:"
                    $vpnRadio.Checked = $true
                    $targetGroup.Controls.Add($vpnRadio)

                    # VPN Selection ComboBox
                    $vpnComboBox = New-Object System.Windows.Forms.ComboBox
                    $vpnComboBox.Location = New-Object System.Drawing.Point(170,20)
                    $vpnComboBox.Size = New-Object System.Drawing.Size(270,20)
                    $vpnComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

                    # Add "None" option
                    $vpnComboBox.Items.Add("None (Use Custom Target)")

                    # Populate VPN list with all sources
                    try {
                        # Standard VPN connections
                        Get-VpnConnection -ErrorAction SilentlyContinue | ForEach-Object {
                            $vpnComboBox.Items.Add($_.Name)
                        }
                        # All-user VPN connections
                        Get-VpnConnection -AllUserConnection -ErrorAction SilentlyContinue | ForEach-Object {
                            if ($vpnComboBox.Items -notcontains $_.Name) {
                                $vpnComboBox.Items.Add($_.Name)
                            }
                        }
                    } catch { }
                    
                    if ($vpnComboBox.Items.Count -gt 0) {
                        $vpnComboBox.SelectedIndex = 0
                    }
                    $targetGroup.Controls.Add($vpnComboBox)

                    # IP/Hostname Radio Button
                    $ipRadio = New-Object System.Windows.Forms.RadioButton
                    $ipRadio.Location = New-Object System.Drawing.Point(10,50)
                    $ipRadio.Size = New-Object System.Drawing.Size(100,20)
                    $ipRadio.Text = "IP/Hostname:"
                    $targetGroup.Controls.Add($ipRadio)

                    # IP/Hostname TextBox
                    $ipTextBox = New-Object System.Windows.Forms.TextBox
                    $ipTextBox.Location = New-Object System.Drawing.Point(120,50)
                    $ipTextBox.Size = New-Object System.Drawing.Size(320,20)
                    $targetGroup.Controls.Add($ipTextBox)

                    # Common Targets Label
                    $commonLabel = New-Object System.Windows.Forms.Label
                    $commonLabel.Location = New-Object System.Drawing.Point(10,80)
                    $commonLabel.Size = New-Object System.Drawing.Size(100,20)
                    $commonLabel.Text = "Quick Select:"
                    $targetGroup.Controls.Add($commonLabel)

                    # Common Targets ComboBox
                    $commonTargets = New-Object System.Windows.Forms.ComboBox
                    $commonTargets.Location = New-Object System.Drawing.Point(120,80)
                    $commonTargets.Size = New-Object System.Drawing.Size(320,20)
                    $commonTargets.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
                    $commonTargets.Items.AddRange(@(
                        "Select target...",
                        "Google DNS (8.8.8.8)",
                        "Cloudflare DNS (1.1.1.1)",
                        "Google.com",
                        "Microsoft.com"
                    ))
                    $commonTargets.SelectedIndex = 0
                    $targetGroup.Controls.Add($commonTargets)

                    # Test Options Group
                    $optionsGroup = New-Object System.Windows.Forms.GroupBox
                    $optionsGroup.Text = "Test Options"
                    $optionsGroup.Location = New-Object System.Drawing.Point(10,140)
                    $optionsGroup.Size = New-Object System.Drawing.Size(460,160)
                    $testForm.Controls.Add($optionsGroup)

                    # Test Options Checkboxes
                    $testOptions = @{
                        pingTest = @{
                            Text    = "Ping Test (ICMP)"
                            Y       = 20
                            Checked = $true
                        }
                        dnsTest = @{
                            Text    = "DNS Resolution"
                            Y       = 50
                            Checked = $true
                        }
                        traceTest = @{
                            Text    = "Traceroute"
                            Y       = 80
                            Checked = $false
                        }
                        portTest = @{
                            Text    = "VPN Ports (500, 1701, 4500)"
                            Y       = 110
                            Checked = $true
                        }
                    }

                    $checkboxes = @{}
                    foreach ($option in $testOptions.GetEnumerator()) {
                        $checkbox = New-Object System.Windows.Forms.CheckBox
                        $checkbox.Location = New-Object System.Drawing.Point(10,$option.Value.Y)
                        $checkbox.Size = New-Object System.Drawing.Size(440,20)
                        $checkbox.Text = $option.Value.Text
                        $checkbox.Checked = $option.Value.Checked
                        $optionsGroup.Controls.Add($checkbox)
                        $checkboxes[$option.Key] = $checkbox
                    }

                    # Event Handlers
                    $vpnRadio.Add_CheckedChanged({
                        if ($vpnRadio.Checked) {
                            $vpnComboBox.Enabled = $true
                            $ipTextBox.Enabled = $false
                            $commonTargets.Enabled = $false
                            $checkboxes['portTest'].Enabled = $true
                        }
                    })

                    $ipRadio.Add_CheckedChanged({
                        if ($ipRadio.Checked) {
                            $vpnComboBox.Enabled = $false
                            $ipTextBox.Enabled = $true
                            $commonTargets.Enabled = $true
                            $checkboxes['portTest'].Enabled = $true
                        }
                    })

                    $commonTargets.Add_SelectedIndexChanged({
                        if ($commonTargets.SelectedIndex -gt 0) {
                            $target = $commonTargets.SelectedItem
                            $ipTextBox.Text = switch ($target) {
                                "Google DNS (8.8.8.8)" { "8.8.8.8" }
                                "Cloudflare DNS (1.1.1.1)" { "1.1.1.1" }
                                "Google.com" { "google.com" }
                                "Microsoft.com" { "microsoft.com" }
                            }
                            $ipRadio.Checked = $true
                        }
                    })

                    # Start Test Button
                    $startButton = New-Object System.Windows.Forms.Button
                    $startButton.Location = New-Object System.Drawing.Point(10,310)
                    $startButton.Size = New-Object System.Drawing.Size(460,30)
                    $startButton.Text = "Start Tests"
                    $startButton.Add_Click({
                        # Determine test target
                        if ($vpnRadio.Checked) {
                            if ($vpnComboBox.SelectedItem) {
                                $vpnName = $vpnComboBox.SelectedItem
                                
                                # Check if "None" is selected
                                if ($vpnName -eq "None (Use Custom Target)") {
                                    # Switch to custom IP mode
                                    $ipRadio.Checked = $true
                                    $testTarget = $ipTextBox.Text.Trim()
                                    $testName = "Custom: ${testTarget}"
                                    
                                    if ([string]::IsNullOrWhiteSpace($testTarget)) {
                                        [System.Windows.Forms.MessageBox]::Show(
                                            "Please enter an IP address or hostname.",
                                            "Input Required",
                                            [System.Windows.Forms.MessageBoxButtons]::OK,
                                            [System.Windows.Forms.MessageBoxIcon]::Warning
                                        )
                                        return
                                    }
                                }
                                else {
                                    try {
                                        # Try both standard and all-user connections
                                        $vpn = Get-VpnConnection -Name $vpnName -ErrorAction SilentlyContinue
                                        if (-not $vpn) {
                                            $vpn = Get-VpnConnection -Name $vpnName -AllUserConnection -ErrorAction Stop
                                        }
                                        $testTarget = $vpn.ServerAddress
                                        $testName = "VPN: $vpnName (${testTarget})"
                                    }
                                    catch {
                                        [System.Windows.Forms.MessageBox]::Show(
                                            "Could not retrieve VPN connection details.",
                                            "VPN Error",
                                            [System.Windows.Forms.MessageBoxButtons]::OK,
                                            [System.Windows.Forms.MessageBoxIcon]::Error
                                        )
                                        return
                                    }
                                }
                            }
                            else {
                                [System.Windows.Forms.MessageBox]::Show(
                                    "Please select a VPN connection.",
                                    "Selection Required",
                                    [System.Windows.Forms.MessageBoxButtons]::OK,
                                    [System.Windows.Forms.MessageBoxIcon]::Warning
                                )
                                return
                            }
                        }
                        else {
                            $testTarget = $ipTextBox.Text.Trim()
                            $testName = "Custom: ${testTarget}"
                            if ([string]::IsNullOrWhiteSpace($testTarget)) {
                                [System.Windows.Forms.MessageBox]::Show(
                                    "Please enter an IP address or hostname.",
                                    "Input Required",
                                    [System.Windows.Forms.MessageBoxButtons]::OK,
                                    [System.Windows.Forms.MessageBoxIcon]::Warning
                                )
                                return
                            }
                        }

                        Update-Status "Starting network tests for ${testName}..."

                        # Ping Test
                        if ($checkboxes['pingTest'].Checked) {
                            Update-Status "`nPing Test:"
                            try {
                                $pingResults = Test-Connection -ComputerName ${testTarget} -Count 4 -ErrorAction Stop
                                
                                Update-Status "  Target: ${testTarget}"
                                $pingResults | ForEach-Object {
                                    Update-Status "  Reply from $($_.Address): time=$($_.ResponseTime)ms"
                                }
                                
                                $avgTime = ($pingResults | Measure-Object -Property ResponseTime -Average).Average
                                Update-Status "  Average response time: $([math]::Round($avgTime, 2))ms"
                            }
                            catch {
                                Update-Status "  Failed to ping ${testTarget}: $($_.Exception.Message)"
                            }
                        }

                        # DNS Resolution Test
                        if ($checkboxes['dnsTest'].Checked) {
                            Update-Status "`nDNS Resolution Test:"
                            try {
                                # If input is IP, do reverse lookup
                                if (${testTarget} -match "^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$") {
                                    $reverseDns = Resolve-DnsName -Name ${testTarget} -ErrorAction Stop -Type PTR
                                    
                                    Update-Status "  IP Address: ${testTarget}"
                                    foreach ($record in $reverseDns) {
                                        Update-Status "  Hostname: $($record.NameHost)"
                                    }
                                }
                                # If input is hostname, do forward lookup
                                else {
                                    $dnsResult = Resolve-DnsName -Name ${testTarget} -ErrorAction Stop
                                    
                                    Update-Status "  Hostname: ${testTarget}"
                                    foreach ($record in $dnsResult) {
                                        if ($record.Type -in @("A", "AAAA")) {
                                            Update-Status "  IP Address: $($record.IPAddress)"
                                        }
                                    }
                                }
                            }
                            catch {
                                Update-Status "  DNS resolution failed: $($_.Exception.Message)"
                            }
                        }

                        # Traceroute Test
                        if ($checkboxes['traceTest'].Checked) {
                            Update-Status "`nTraceroute Test:"
                            try {
                                $traceResults = Test-NetConnection -ComputerName ${testTarget} -TraceRoute -ErrorAction Stop
                                
                                Update-Status "  Tracing route to ${testTarget}"
                                $hop = 1
                                foreach ($node in $traceResults.TraceRoute) {
                                    Update-Status "  $hop. $node"
                                    $hop++
                                }
                            }
                            catch {
                                Update-Status "  Traceroute failed: $($_.Exception.Message)"
                            }
                        }

                        # VPN Port Test
                        if ($checkboxes['portTest'].Checked) {
                            Update-Status "`nVPN Port Test:"
                            $ports = @{
                                500 = "IKE"
                                1701 = "L2TP"
                                4500 = "NAT-T"
                            }
                            foreach ($port in $ports.GetEnumerator()) {
                                try {
                                    $portTest = Test-NetConnection -ComputerName ${testTarget} -Port $port.Key -WarningAction SilentlyContinue
                                    $status = if ($portTest.TcpTestSucceeded) { "Open" } else { "Closed" }
                                    Update-Status "  Port $($port.Key) ($($port.Value)): $status"
                                }
                                catch {
                                    Update-Status "  Port $($port.Key) test failed: $($_.Exception.Message)"
                                }
                            }
                        }

                        # Network Interface Information
                        Update-Status "`nNetwork Interface Information:"
                        try {
                            $activeAdapter = Get-NetAdapter | Where-Object Status -eq 'Up' | Sort-Object -Property Speed -Descending | Select-Object -First 1
                            if ($activeAdapter) {
                                Update-Status "  Primary Network Adapter: $($activeAdapter.Name)"
                                Update-Status "  Status: $($activeAdapter.Status)"
                                Update-Status "  Speed: $($activeAdapter.LinkSpeed)"
                                Update-Status "  MAC Address: $($activeAdapter.MacAddress)"
                            }
                        }
                        catch {
                            Update-Status "  Failed to retrieve network interface information: $($_.Exception.Message)"
                        }

                        $testForm.Close()
                    })
                    $testForm.Controls.Add($startButton)

                    # Show the form
                    $testForm.ShowDialog()
                }
            },
            @{
                'Text'   = 'Review VPN Event Logs'
                'Action' = {
                    Update-Status "Retrieving VPN client connectivity logs..."
                    $diagResults.Clear()
                    $diagResults.AppendText("=== VPN Client Connectivity Logs (Last 3 Days) ===`r`n`r`n")
                    
                    # Define client-specific VPN event sources
                    $vpnClientSources = @(
                        @{LogName = "Application"; ProviderName = "RasClient"},
                        @{LogName = "System"; ProviderName = "RasMan"}
                    )
                    
                    $since = (Get-Date).AddDays(-3) # Last 3 days of logs
                    $foundEvents = $false
                    
                    # Process each client-specific event source
                    foreach ($src in $vpnClientSources) {
                        $diagResults.AppendText("Searching for $($src.ProviderName) events in $($src.LogName)...`r`n")
                        [System.Windows.Forms.Application]::DoEvents()
                        
                        try {
                            # Use proper parameter names for Get-WinEvent
                            $filterHashtable = @{
                                LogName = $src.LogName
                                ProviderName = $src.ProviderName
                                StartTime = $since
                            }
                            
                            $events = Get-WinEvent -FilterHashtable $filterHashtable -MaxEvents 15 -ErrorAction Stop
                            
                            if ($events -and $events.Count -gt 0) {
                                $foundEvents = $true
                                $diagResults.AppendText("Found $($events.Count) client VPN events from $($src.ProviderName)`r`n")
                                
                                foreach ($evt in $events) {
                                    $diagResults.AppendText("[$($evt.TimeCreated)] [ID: $($evt.Id)] [Level: $($evt.LevelDisplayName)]`r`n")
                                    # Truncate message if too long to prevent UI slowdown
                                    $message = if ($evt.Message.Length -gt 500) { $evt.Message.Substring(0, 500) + "..." } else { $evt.Message }
                                    $diagResults.AppendText("$message`r`n")
                                    $diagResults.AppendText("-----------------------------------------`r`n")
                                }
                            } else {
                                $diagResults.AppendText("No client VPN events found`r`n")
                            }
                        } catch {
                            $diagResults.AppendText("Error: $($_.Exception.Message)`r`n")
                        }
                        
                        $diagResults.AppendText("`r`n")
                        $diagResults.ScrollToCaret()
                        [System.Windows.Forms.Application]::DoEvents()
                        Update-Status "Searched $($src.LogName) for $($src.ProviderName) events"
                    }
                    
                    # Check for common VPN client error Event IDs
                    $diagResults.AppendText("Searching for common VPN client error Event IDs...`r`n")
                    [System.Windows.Forms.Application]::DoEvents()
                    
                    # Client-specific VPN error IDs
                    $vpnClientErrorIDs = @(809, 789, 800, 801, 812, 13801, 13802, 13803, 13804, 13805, 13806, 13807, 13808, 13809, 13810, 13811, 13812, 13813, 13814, 13815, 13816, 13817, 13818, 13819, 13820, 13821, 13822, 13823, 13824, 13825, 13826, 13827, 13828, 13829, 13830, 13831, 13832, 13833, 13834, 13835, 13836, 13837, 13838, 13839, 13840, 13841, 13842, 13843, 13844, 13845, 13846, 13847, 13848, 13849, 13850, 13851, 13852, 13853, 13854, 13855, 13856, 13857, 13858, 13859, 13860, 13861, 13862, 13863, 13864, 13865, 13866, 13867, 13868, 13869, 13870, 13871, 13872, 13873, 13874, 13875, 13876, 13877, 13878, 13879, 13880, 13881, 13882, 13883, 13884, 13885, 13886, 13887, 13888, 13889, 13890, 13891, 13892, 13893, 13894, 13895, 13896, 13897, 13898, 13899, 13900, 13901, 13902, 13903, 13904, 13905, 13906, 13907, 13908, 13909, 13910, 13911, 13912, 13913, 13914, 13915, 13916, 13917, 13918, 13919, 13920, 13921, 13922, 13923, 13924, 13925, 13926, 13927, 13928, 13929, 13930, 13931, 13932, 13933, 13934, 13935, 13936, 13937, 13938, 13939, 13940, 13941, 13942, 13943, 13944, 13945, 13946, 13947, 13948, 13949, 13950, 13951, 13952, 13953, 13954, 13955, 13956, 13957, 13958, 13959, 13960, 13961, 13962, 13963, 13964, 13965, 13966, 13967, 13968, 13969, 13970, 13971, 13972, 13973, 13974, 13975, 13976, 13977, 13978, 13979, 13980, 13981, 13982, 13983, 13984, 13985, 13986, 13987, 13988, 13989, 13990, 13991, 13992, 13993, 13994, 13995, 13996, 13997, 13998, 13999, 14000)
                    
                    try {
                        # Create a filter for specific client error IDs in the last 3 days
                        $idFilter = @{
                            LogName = "System"
                            StartTime = $since
                        }
                        
                        # Due to the large number of IDs, we'll search for any System events and filter by ID
                        $systemEvents = Get-WinEvent -FilterHashtable $idFilter -MaxEvents 100 -ErrorAction Stop
                        $clientErrorEvents = $systemEvents | Where-Object { $vpnClientErrorIDs -contains $_.Id }
                        
                        if ($clientErrorEvents -and $clientErrorEvents.Count -gt 0) {
                            $foundEvents = $true
                            $diagResults.AppendText("Found $($clientErrorEvents.Count) VPN client error events`r`n")
                            
                            foreach ($evt in $clientErrorEvents) {
                                $diagResults.AppendText("[$($evt.TimeCreated)] [ID: $($evt.Id)] [Provider: $($evt.ProviderName)]`r`n")
                                # Truncate message if too long
                                $message = if ($evt.Message.Length -gt 500) { $evt.Message.Substring(0, 500) + "..." } else { $evt.Message }
                                $diagResults.AppendText("$message`r`n")
                                $diagResults.AppendText("-----------------------------------------`r`n")
                            }
                        } else {
                            $diagResults.AppendText("No VPN client error events found`r`n")
                        }
                    } catch {
                        $diagResults.AppendText("Error searching for VPN client error events: $($_.Exception.Message)`r`n")
                    }
                    
                    # Check for VPN client connection attempts in the Application log
                    $diagResults.AppendText("`r`nSearching for VPN client connection attempts...`r`n")
                    [System.Windows.Forms.Application]::DoEvents()
                    
                    try {
                        $appFilter = @{
                            LogName = "Application"
                            StartTime = $since
                        }
                        
                        $appEvents = Get-WinEvent -FilterHashtable $appFilter -MaxEvents 100 -ErrorAction Stop | 
                                    Where-Object { ($_.Message -like "*VPN*" -or $_.Message -like "*connection*") -and 
                                                ($_.Message -like "*dial*" -or $_.Message -like "*connect*" -or 
                                                    $_.Message -like "*authentication*" -or $_.Message -like "*tunnel*") }
                        
                        if ($appEvents -and $appEvents.Count -gt 0) {
                            $foundEvents = $true
                            $diagResults.AppendText("Found $($appEvents.Count) VPN client connection events`r`n")
                            
                            foreach ($evt in $appEvents) {
                                $diagResults.AppendText("[$($evt.TimeCreated)] [ID: $($evt.Id)] [Provider: $($evt.ProviderName)]`r`n")
                                # Truncate message if too long
                                $message = if ($evt.Message.Length -gt 500) { $evt.Message.Substring(0, 500) + "..." } else { $evt.Message }
                                $diagResults.AppendText("$message`r`n")
                                $diagResults.AppendText("-----------------------------------------`r`n")
                            }
                        } else {
                            $diagResults.AppendText("No VPN client connection events found`r`n")
                        }
                    } catch {
                        $diagResults.AppendText("Error searching Application log: $($_.Exception.Message)`r`n")
                    }
                    
                    if (-not $foundEvents) {
                        $diagResults.AppendText("`r`nNo VPN client connectivity events found in the last 3 days.`r`n")
                        $diagResults.AppendText("This could mean either no VPN connection attempts were made or logging is disabled.`r`n")
                    }
                    
                    $diagResults.ScrollToCaret()
                    Update-Status "VPN client event log review completed."
                }
            },
            @{
                'Text'   = 'Reset Network Stack'
                'Action' = {
                    Update-Status "Resetting network stack..."
                    $result = Reset-NetworkDevices
                    if ($result) {
                        Update-Status "Network stack reset completed successfully"
                    }
                }
            },
            @{
                'Text'   = 'Check VPN Prerequisites'
                'Action' = {
                    Update-Status "Checking VPN Prerequisites..."
            
                    Update-Status "`nRequired Services:"
                    $services = @{
                        'RasMan'       = 'Remote Access Connection Manager - Manages VPN connections'
                        'PolicyAgent'  = 'IPsec Policy Agent - Handles IPsec security policies'
                        'IKEEXT'       = 'IKE and AuthIP IPsec Keying Modules - Manages security associations'
                    }
            
                    foreach ($svc in $services.GetEnumerator()) {
                        $status = Get-Service -Name $svc.Key -ErrorAction SilentlyContinue
                        Update-Status ("  " + $svc.Key + " (" + $svc.Value + ")")
                        Update-Status ("    Status: " + $status.Status)
                        Update-Status ("    Startup Type: " + $status.StartType)
                    }
            
                    Update-Status "`nChecking L2TP/IPsec Components:"
                    # Check if RAS device exists
                    $rasDevice = Get-WmiObject Win32_NetworkAdapter | Where-Object { $_.Name -like "*WAN Miniport (L2TP)*" }
                    Update-Status "  L2TP Adapter: $(if($rasDevice){'Present'}else{'Not Found'})"
            
                    # Check IPSec Policy
                    $ipsecPolicy = Get-WmiObject -Namespace root\SecurityCenter2 -Class FirewallProduct -ErrorAction SilentlyContinue
                    Update-Status "  IPSec Support: $(if($ipsecPolicy){'Enabled'}else{'Not Found'})"
            
                    Update-Status "`nChecking Firewall Rules:"
                    $ports = @(500, 1701, 4500)
                    foreach ($port in $ports) {
                        $rule = Get-NetFirewallRule -DisplayName "*$port*" -ErrorAction SilentlyContinue
                        $portType = switch ($port) {
                            500 { "IKEEXT" }
                            1701 { "L2TP" }
                            4500 { "NAT-T" }
                        }
                        $status = if ($rule) { "Allowed" } else { "Not Found" }
                        Update-Status ("  Port " + $port + " (" + $portType + "): " + $status)
                    }

                    # Check Registry Settings
                    Update-Status "`nChecking Registry Settings:"
                    $regSettings = @{
                        "HKLM:\System\CurrentControlSet\Services\PolicyAgent"       = "AssumeUDPEncapsulationContextOnSendRule"
                        "HKLM:\System\CurrentControlSet\Services\RasMan\Parameters" = "ProhibitIpSec"
                    }
                    foreach ($reg in $regSettings.GetEnumerator()) {
                        $value = Get-ItemProperty -Path $reg.Key -Name $reg.Value -ErrorAction SilentlyContinue
                        Update-Status "  $($reg.Value): $(if($value){'Configured'}else{'Not Configured'})"
                    }
                }
            },

            @{
                'Text'   = 'Repair VPN Connection'
                'Action' = {
                    # Create repair options form
                    $repairForm = New-Object System.Windows.Forms.Form
                    $repairForm.Text = "VPN Repair Options"
                    $repairForm.Size = New-Object System.Drawing.Size(400, 450)
                    $repairForm.StartPosition = 'CenterParent'
            
                    # VPN Selection Group
                    $vpnGroup = New-Object System.Windows.Forms.GroupBox
                    $vpnGroup.Text = "VPN Selection"
                    $vpnGroup.Location = New-Object System.Drawing.Point(10, 10)
                    $vpnGroup.Size = New-Object System.Drawing.Size(360, 60)
                    $repairForm.Controls.Add($vpnGroup)
            
                    $vpnComboBox = New-Object System.Windows.Forms.ComboBox
                    $vpnComboBox.Location = New-Object System.Drawing.Point(10, 20)
                    $vpnComboBox.Size = New-Object System.Drawing.Size(340, 20)
                    $vpnComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
            
                    # Populate VPN list with both user and all-user connections
                    try {
                        # Standard VPN connections
                        Get-VpnConnection -ErrorAction SilentlyContinue | ForEach-Object {
                            $vpnComboBox.Items.Add("$($_.Name) (User)")
                        }
                        
                        # All-user VPN connections
                        Get-VpnConnection -AllUserConnection -ErrorAction SilentlyContinue | ForEach-Object {
                            $vpnComboBox.Items.Add("$($_.Name) (All Users)")
                        }
                    } catch {
                        Update-Status "Error getting VPN connections: $($_.Exception.Message)"
                    }
            
                    if ($vpnComboBox.Items.Count -gt 0) {
                        $vpnComboBox.SelectedIndex = 0
                    }
                    $vpnGroup.Controls.Add($vpnComboBox)
            
                    # Credentials Group
                    $credGroup = New-Object System.Windows.Forms.GroupBox
                    $credGroup.Text = "VPN Credentials"
                    $credGroup.Location = New-Object System.Drawing.Point(10, 80)
                    $credGroup.Size = New-Object System.Drawing.Size(360, 150)
                    $repairForm.Controls.Add($credGroup)
            
                    # PSK Field
                    $pskLabel = New-Object System.Windows.Forms.Label
                    $pskLabel.Text = "Pre-shared Key (PSK)*:"
                    $pskLabel.Location = New-Object System.Drawing.Point(10, 20)
                    $pskLabel.Size = New-Object System.Drawing.Size(130, 20)
                    $credGroup.Controls.Add($pskLabel)
            
                    $pskBox = New-Object System.Windows.Forms.TextBox
                    $pskBox.Location = New-Object System.Drawing.Point(140, 20)
                    $pskBox.Size = New-Object System.Drawing.Size(210, 20)
                    $pskBox.PasswordChar = '*'
                    $credGroup.Controls.Add($pskBox)
            
                    # Username Field
                    $userLabel = New-Object System.Windows.Forms.Label
                    $userLabel.Text = "Username:"
                    $userLabel.Location = New-Object System.Drawing.Point(10, 50)
                    $userLabel.Size = New-Object System.Drawing.Size(130, 20)
                    $credGroup.Controls.Add($userLabel)
            
                    $userBox = New-Object System.Windows.Forms.TextBox
                    $userBox.Location = New-Object System.Drawing.Point(140, 50)
                    $userBox.Size = New-Object System.Drawing.Size(210, 20)
                    $credGroup.Controls.Add($userBox)
            
                    # Password Field
                    $passLabel = New-Object System.Windows.Forms.Label
                    $passLabel.Text = "Password:"
                    $passLabel.Location = New-Object System.Drawing.Point(10, 80)
                    $passLabel.Size = New-Object System.Drawing.Size(130, 20)
                    $credGroup.Controls.Add($passLabel)
            
                    $passBox = New-Object System.Windows.Forms.TextBox
                    $passBox.Location = New-Object System.Drawing.Point(140, 80)
                    $passBox.Size = New-Object System.Drawing.Size(210, 20)
                    $passBox.PasswordChar = '*'
                    $credGroup.Controls.Add($passBox)
            
                    # Remember credentials checkbox
                    $rememberCred = New-Object System.Windows.Forms.CheckBox
                    $rememberCred.Text = "Remember credentials"
                    $rememberCred.Location = New-Object System.Drawing.Point(140, 110)
                    $rememberCred.Size = New-Object System.Drawing.Size(210, 20)
                    $rememberCred.Checked = $true
                    $credGroup.Controls.Add($rememberCred)
            
                    # Repair Options Group
                    $optionsGroup = New-Object System.Windows.Forms.GroupBox
                    $optionsGroup.Text = "Repair Options"
                    $optionsGroup.Location = New-Object System.Drawing.Point(10, 240)
                    $optionsGroup.Size = New-Object System.Drawing.Size(360, 120)
                    $repairForm.Controls.Add($optionsGroup)
            
                    $options = @{
                        'RestartServices' = "Restart VPN Client Services"
                        'ClearDNS'       = "Clear DNS Cache"
                        'RecreateVPN'    = "Recreate VPN Connection"
                    }
            
                    $yPos = 20
                    $checkboxes = @{}
                    foreach ($option in $options.GetEnumerator()) {
                        $checkbox = New-Object System.Windows.Forms.CheckBox
                        $checkbox.Location = New-Object System.Drawing.Point(10, $yPos)
                        $checkbox.Size = New-Object System.Drawing.Size(330, 20)
                        $checkbox.Text = $option.Value
                        $checkbox.Checked = $true
                        $checkboxes[$option.Key] = $checkbox
                        $optionsGroup.Controls.Add($checkbox)
                        $yPos += 25
                    }
            
                    # Repair Button
                    $repairButton = New-Object System.Windows.Forms.Button
                    $repairButton.Location = New-Object System.Drawing.Point(10, 370)
                    $repairButton.Size = New-Object System.Drawing.Size(360, 30)
                    $repairButton.Text = "Start Repair"
                    $repairButton.Add_Click({
                        $selectedItem = $vpnComboBox.SelectedItem
                        if (-not $selectedItem) {
                            [System.Windows.Forms.MessageBox]::Show(
                                "Please select a VPN connection.",
                                "Selection Required",
                                [System.Windows.Forms.MessageBoxButtons]::OK,
                                [System.Windows.Forms.MessageBoxIcon]::Warning
                            )
                            return
                        }
            
                        # Extract VPN name and type from selection
                        $vpnName = $selectedItem -replace ' \((User|All Users)\)$', ''
                        $isAllUserVPN = $selectedItem -match 'All Users'
            
                        if ([string]::IsNullOrWhiteSpace($pskBox.Text)) {
                            [System.Windows.Forms.MessageBox]::Show(
                                "Pre-shared Key (PSK) is required.",
                                "Required Field",
                                [System.Windows.Forms.MessageBoxButtons]::OK,
                                [System.Windows.Forms.MessageBoxIcon]::Warning
                            )
                            return
                        }
            
                        Update-Status "Starting repair for VPN connection: $vpnName ($(if ($isAllUserVPN) { 'All Users' } else { 'User' }))"
            
                        # Backup current configuration
                        Update-Status "Creating backup..."
                        Backup-VPNConfiguration -VPNName $vpnName
            
                        # Get existing VPN settings
                        try {
                            $existingVPN = if ($isAllUserVPN) {
                                Get-VpnConnection -Name $vpnName -AllUserConnection -ErrorAction Stop
                            } else {
                                Get-VpnConnection -Name $vpnName -ErrorAction Stop
                            }
            
                            if ($existingVPN) {
                                $vpnParams = @{
                                    Name = $vpnName
                                    ServerAddress = $existingVPN.ServerAddress
                                    TunnelType = "L2tp"
                                    EncryptionLevel = $existingVPN.EncryptionLevel
                                    AuthenticationMethod = $existingVPN.AuthenticationMethod
                                    SplitTunneling = $existingVPN.SplitTunneling
                                    RememberCredential = $rememberCred.Checked
                                    L2tpPsk = $pskBox.Text
                                    Force = $true
                                }
            
                                # Add AllUserConnection parameter if it's an all-user VPN
                                if ($isAllUserVPN) {
                                    $vpnParams.AllUserConnection = $true
                                }
            
                                # Restart Services if selected
                                if ($checkboxes['RestartServices'].Checked) {
                                    $services = @('RasMan', 'RemoteAccess')
                                    foreach ($service in $services) {
                                        Update-Status "Restarting $service service..."
                                        Restart-Service -Name $service -Force -ErrorAction SilentlyContinue
                                        Start-Sleep -Seconds 2
                                    }
                                }
            
                                # Clear DNS if selected
                                if ($checkboxes['ClearDNS'].Checked) {
                                    Update-Status "Clearing DNS cache..."
                                    Clear-DnsClientCache
                                }
            
                                # Recreate VPN if selected
                                if ($checkboxes['RecreateVPN'].Checked) {
                                    Update-Status "Recreating VPN connection..."
                                    
                                    # Remove existing connection
                                    if ($isAllUserVPN) {
                                        Remove-VpnConnection -Name $vpnName -AllUserConnection -Force -ErrorAction SilentlyContinue
                                    } else {
                                        Remove-VpnConnection -Name $vpnName -Force -ErrorAction SilentlyContinue
                                    }
            
                                    # Create new connection
                                    Add-VpnConnection @vpnParams
            
                                    # Set credentials if provided
                                    if (-not [string]::IsNullOrWhiteSpace($userBox.Text)) {
                                        $password = $passBox.Text | ConvertTo-SecureString -AsPlainText -Force
                                        $cred = New-Object System.Management.Automation.PSCredential ($userBox.Text, $password)
                                        
                                        # Set VPN connection credentials
                                        if ($isAllUserVPN) {
                                            Set-VpnConnection -Name $vpnName -AllUserConnection -RememberCredential $rememberCred.Checked
                                        } else {
                                            Set-VpnConnection -Name $vpnName -RememberCredential $rememberCred.Checked
                                        }
            
                                        # Store credentials if remember is checked
                                        if ($rememberCred.Checked) {
                                            cmdkey /add:$existingVPN.ServerAddress /user:$userBox.Text /pass:$passBox.Text
                                        }
                                    }
                                }
            
                                Update-Status "VPN connection repair completed"
                                Update-VPNList
                                $repairForm.Close()
                            }
                        }
                        catch {
                            Update-Status "Error during VPN repair: $($_.Exception.Message)"
                            [System.Windows.Forms.MessageBox]::Show(
                                "Error during VPN repair: $($_.Exception.Message)",
                                "Repair Error",
                                [System.Windows.Forms.MessageBoxButtons]::OK,
                                [System.Windows.Forms.MessageBoxIcon]::Error
                            )
                        }
                    })
                    $repairForm.Controls.Add($repairButton)
            
                    # Show the form
                    $repairForm.ShowDialog()
                }
            }
        )
        # Button layout configuration
        $buttonConfig = @{
            StartX            = 10
            StartY            = 20
            Width             = 220
            Height            = 30
            HorizontalSpacing = 10
            VerticalSpacing   = 10
            ButtonsPerColumn  = 10
        }

        # Create and position buttons
        for ($i = 0; $i -lt $diagButtons.Count; $i++) {
            $column = [Math]::Floor($i / $buttonConfig.ButtonsPerColumn)
            $row = $i % $buttonConfig.ButtonsPerColumn
        
            $xPos = $buttonConfig.StartX + ($column * ($buttonConfig.Width + $buttonConfig.HorizontalSpacing))
            $yPos = $buttonConfig.StartY + ($row * ($buttonConfig.Height + $buttonConfig.VerticalSpacing))
        
            $button = New-Object System.Windows.Forms.Button
            $button.Location = New-Object System.Drawing.Point($xPos, $yPos)
            $button.Size = New-Object System.Drawing.Size($buttonConfig.Width, $buttonConfig.Height)
            $button.Text = $diagButtons[$i].Text
            $button.Add_Click($diagButtons[$i].Action)
            $diagTab.Controls.Add($button)
        }

        # Adjust results textbox position based on buttons
        $resultsStartX = $buttonConfig.StartX + ($buttonConfig.Width + $buttonConfig.HorizontalSpacing)
        $diagResults = New-Object System.Windows.Forms.TextBox
        $diagResults.Location = New-Object System.Drawing.Point($resultsStartX, $buttonConfig.StartY)
        $diagResults.Size = New-Object System.Drawing.Size(500, 500)
        $diagResults.Multiline = $true
        $diagResults.ScrollBars = 'Vertical'
        $diagResults.ReadOnly = $true
        $diagTab.Controls.Add($diagResults)

        # Logs Tab
        $logsTab = $tabPages["Logs"]
    
        $logsTextBox = New-Object System.Windows.Forms.TextBox
        $logsTextBox.Location = New-Object System.Drawing.Point(12, 40)
        $logsTextBox.Size = New-Object System.Drawing.Size(740, 480)
        $logsTextBox.Multiline = $true
        $logsTextBox.ScrollBars = 'Vertical'
        $logsTextBox.ReadOnly = $true
    
        $refreshButton = New-Object System.Windows.Forms.Button
        $refreshButton.Location = New-Object System.Drawing.Point(12, 10)
        $refreshButton.Size = New-Object System.Drawing.Size(100, 25)
        $refreshButton.Text = "Refresh Logs"
        $refreshButton.Add_Click({
                $logsTextBox.Clear()
                if (Test-Path $Logger.LogPath) {
                    $logsTextBox.AppendText("=== General Logs ===`r`n")
                    $logsTextBox.AppendText((Get-Content $Logger.LogPath | Out-String))
                }
                if (Test-Path $Logger.ErrorLogPath) {
                    $logsTextBox.AppendText("`r`n=== Error Logs ===`r`n")
                    $logsTextBox.AppendText((Get-Content $Logger.ErrorLogPath | Out-String))
                }
            })
    
        $clearButton = New-Object System.Windows.Forms.Button
        $clearButton.Location = New-Object System.Drawing.Point(120, 10)
        $clearButton.Size = New-Object System.Drawing.Size(100, 25)
        $clearButton.Text = "Clear Logs"
        $clearButton.Add_Click({
                if (Test-Path $Logger.LogPath) { Clear-Content $Logger.LogPath }
                if (Test-Path $Logger.ErrorLogPath) { Clear-Content $Logger.ErrorLogPath }
                $logsTextBox.Clear()
                Update-Status "Logs cleared successfully"
            })
    
        $logsTab.Controls.Add($refreshButton)
        $logsTab.Controls.Add($clearButton)
        $logsTab.Controls.Add($logsTextBox)

        # Add TabControl to form
        $form.Controls.Add($tabControl)

        # Show the form
        $form.ShowDialog()
    
    }

    # Start the GUI
    Show-VPNTroubleshooterGUI
