# /Users/BrandonCook/VPNTroubleshooterGUI.ps1
# L2TP VPN Troubleshooter
# Version 0.13
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

# Initialize Logger
$Logger = [VPNLogger]::new()
function Start-VPNDiagnostics {
    param([string]$VPNName)
    
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

    # Check Required Services
    $services = @{
        'RasMan'       = 'Remote Access Connection Manager'
        'RemoteAccess' = 'Routing and Remote Access'
        'PolicyAgent'  = 'IPsec Policy Agent'
        'IKEExt'       = 'IKE and AuthIP IPsec Keying Modules'
    }

    foreach ($service in $services.GetEnumerator()) {
        $status = Get-Service -Name $service.Key -ErrorAction SilentlyContinue
        $diagnosticResults += @{
            Component = $service.Value
            Status    = if ($status.Status -eq 'Running') { "Running" } else { "Not Running" }
            Details   = "Service State: $($status.Status), Start Type: $($status.StartType)"
        }
    }

    # Check WAN Miniports
    $miniports = Get-PnpDevice | Where-Object { $_.FriendlyName -like "*WAN Miniport*" }
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

        # Test VPN Ports
        $ports = @(500, 1701, 4500)
        foreach ($port in $ports) {
            $portTest = Test-NetConnection -ComputerName $vpnConnection.ServerAddress -Port $port -WarningAction SilentlyContinue
            $diagnosticResults += @{
                Component = "Port $port"
                Status    = if ($portTest.TcpTestSucceeded) { "Open" } else { "Closed" }
                Details   = "Required for L2TP/IPsec"
            }
        }
    }

    # Check Registry Settings
    $registryPaths = @{
        "HKLM:\SYSTEM\CurrentControlSet\Services\PolicyAgent"       = "AssumeUDPEncapsulationContextOnSendRule"
        "HKLM:\SYSTEM\CurrentControlSet\Services\RasMan\Parameters" = "ProhibitIpSec"
    }

    foreach ($path in $registryPaths.GetEnumerator()) {
        $value = Get-ItemProperty -Path $path.Key -Name $path.Value -ErrorAction SilentlyContinue
        $diagnosticResults += @{
            Component = "Registry: $($path.Value)"
            Status    = if ($value) { "Configured" } else { "Missing" }
            Details   = "Path: $($path.Key)"
        }
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
    $backupDate = Get-Date -Format "yyyy-MM-dd_HH-mm"
    $backupDir = "C:\VPNBackup_$backupDate"
    
    try {
        New-Item -ItemType Directory -Path $backupDir -Force
        
        # Export VPN connections
        $vpnConnections = Get-VpnConnection -ErrorAction SilentlyContinue
        $vpnConnections | ConvertTo-Json | Out-File "$backupDir\vpn_connections.json"
        
        # Export registry settings
        $registryPaths = @(
            "HKLM:\SYSTEM\CurrentControlSet\Services\PolicyAgent",
            "HKLM:\SYSTEM\CurrentControlSet\Services\RasMan",
            "HKLM:\SYSTEM\CurrentControlSet\Services\Rasl2tp"
        )
        
        foreach ($path in $registryPaths) {
            $regName = ($path -split '\\')[-1]
            reg export ($path -replace 'HKLM:\\', 'HKLM\') "$backupDir\${regName}.reg" /y
        }
        
        # Export network adapter configuration
        Get-NetAdapter | Export-Clixml "$backupDir\network_adapters.xml"
        
        $Logger.Log("Backup created successfully in $backupDir")
        return $true
    }
    catch {
        $Logger.Log("Backup failed: $_", "ERROR")
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
        [ValidateSet("Required", "Optional", "None", "Maximum")]
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
# Test VPN Connection
function Test-VPNConnection {
    param(
        [string]$VPNName,
        [string]$ServerAddress
    )
    
    try {
        $connection = Get-VpnConnection -Name $VPNName -ErrorAction Stop
        
        $results = @{
            "Connection Name"       = $connection.Name
            "Server Address"        = $connection.ServerAddress
            "Connection Status"     = $connection.ConnectionStatus
            "Tunnel Type"           = $connection.TunnelType
            "Authentication Method" = $connection.AuthenticationMethod
            "Split Tunneling"       = $connection.SplitTunneling
            "Encryption Level"      = $connection.EncryptionLevel
        }
        
        # Test connectivity
        $ports = @(500, 1701, 4500)
        foreach ($port in $ports) {
            $test = Test-NetConnection -ComputerName $ServerAddress -Port $port -WarningAction SilentlyContinue
            $results["Port $port"] = $test.TcpTestSucceeded
        }
        
        return $results
    }
    catch {
        $Logger.Log("Connection test failed: $_", "ERROR")
        return $null
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
        $newVPNGroup.Size = New-Object System.Drawing.Size(740, 300)
        $configTab.Controls.Add($newVPNGroup)

        # Basic Settings Group
        $basicSettingsGroup = New-Object System.Windows.Forms.GroupBox
        $basicSettingsGroup.Text = "Basic Settings"
        $basicSettingsGroup.Location = New-Object System.Drawing.Point(10, 20)
        $basicSettingsGroup.Size = New-Object System.Drawing.Size(350, 200)
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
        
            $control.Location = New-Object System.Drawing.Point(140, $yPosition)
            $control.Size = New-Object System.Drawing.Size(200, 20)
            $basicSettingsGroup.Controls.Add($control)
            $inputControls[$field.Key] = $control

            $yPosition += 30
        }

        # Advanced Settings Group
        $advancedSettingsGroup = New-Object System.Windows.Forms.GroupBox
        $advancedSettingsGroup.Text = "Advanced Settings"
        $advancedSettingsGroup.Location = New-Object System.Drawing.Point(370, 20)
        $advancedSettingsGroup.Size = New-Object System.Drawing.Size(350, 200)
        $newVPNGroup.Controls.Add($advancedSettingsGroup)

        # Advanced Settings Controls
        $advancedSettings = @{
            'Split Tunneling'           = @{ Type = 'CheckBox'; Default = $false }
            'Remember Credential'       = @{ Type = 'CheckBox'; Default = $true }
            'Authentication Method'     = @{ 
                Type    = 'ComboBox'
                Options = @('MSChapv2', 'EAP', 'PAP', 'CHAP')
                Default = 'MSChapv2'
            }
            'Encryption Level'          = @{
                Type    = 'ComboBox'
                Options = @('Required', 'Optional', 'None', 'Maximum')
                Default = 'None'
            }
            'IPv4 Only'                 = @{ Type = 'CheckBox'; Default = $false }
            'Use Winlogon Credential'   = @{ Type = 'CheckBox'; Default = $false }
            'Idle Disconnect (seconds)' = @{ Type = 'TextBox'; Default = '0' }
        }

        $yPos = 30
        foreach ($setting in $advancedSettings.GetEnumerator()) {
            $label = New-Object System.Windows.Forms.Label
            $label.Location = New-Object System.Drawing.Point(10, $yPos)
            $label.Size = New-Object System.Drawing.Size(120, 20)
            $label.Text = $setting.Key
            $advancedSettingsGroup.Controls.Add($label)

            switch ($setting.Value.Type) {
                'CheckBox' {
                    $control = New-Object System.Windows.Forms.CheckBox
                    $control.Location = New-Object System.Drawing.Point(140, $yPos)
                    $control.Size = New-Object System.Drawing.Size(200, 20)
                    $control.Checked = $setting.Value.Default
                }
                'ComboBox' {
                    $control = New-Object System.Windows.Forms.ComboBox
                    $control.Location = New-Object System.Drawing.Point(140, $yPos)
                    $control.Size = New-Object System.Drawing.Size(200, 20)
                    $control.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
                    foreach ($option in $setting.Value.Options) {
                        $control.Items.Add($option)
                    }
                    $control.SelectedItem = $setting.Value.Default
                }
                'TextBox' {
                    $control = New-Object System.Windows.Forms.TextBox
                    $control.Location = New-Object System.Drawing.Point(140, $yPos)
                    $control.Size = New-Object System.Drawing.Size(200, 20)
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
        $createButton.Text = "Create New VPN"
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

        # Test Connection Button
        $testButton = New-Object System.Windows.Forms.Button
        $testButton.Location = New-Object System.Drawing.Point(160, 0)
        $testButton.Size = New-Object System.Drawing.Size(150, 30)
        $testButton.Text = "Test Connection"
        # Replace the Test Connection Button click handler with:
        $testButton.Add_Click({
            if ([string]::IsNullOrWhiteSpace($inputControls['VPN Name'].Text)) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Please enter a VPN Name to test",
                    "Missing Information",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                return
            }

            # Create test progress form
            $testForm = New-Object System.Windows.Forms.Form
            $testForm.Text = "Testing VPN Connection"
            $testForm.Size = New-Object System.Drawing.Size(400,150)
            $testForm.StartPosition = 'CenterParent'
            $testForm.FormBorderStyle = 'FixedDialog'
            $testForm.MaximizeBox = $false
            $testForm.MinimizeBox = $false

            $progressBar = New-Object System.Windows.Forms.ProgressBar
            $progressBar.Location = New-Object System.Drawing.Point(10,20)
            $progressBar.Size = New-Object System.Drawing.Size(360,20)
            $progressBar.Style = 'Marquee'
            $testForm.Controls.Add($progressBar)

            $statusLabel = New-Object System.Windows.Forms.Label
            $statusLabel.Location = New-Object System.Drawing.Point(10,50)
            $statusLabel.Size = New-Object System.Drawing.Size(360,20)
            $statusLabel.Text = "Testing connection..."
            $testForm.Controls.Add($statusLabel)

            $cancelButton = New-Object System.Windows.Forms.Button
            $cancelButton.Location = New-Object System.Drawing.Point(150,80)
            $cancelButton.Size = New-Object System.Drawing.Size(75,23)
            $cancelButton.Text = "Cancel"
            $testForm.Controls.Add($cancelButton)

            $testCancelled = $false
            $cancelButton.Add_Click({
                $testCancelled = $true
                $testForm.Close()
            })

            # Create job to run tests
            $job = Start-Job -ScriptBlock {
                param($vpnName)
                try {
                    # Try all methods to find the VPN connection
                    $vpnConfig = Get-VpnConnection -Name $vpnName -ErrorAction Stop
                    if (-not $vpnConfig) {
                        $vpnConfig = Get-VpnConnection -Name $vpnName -AllUserConnection -ErrorAction Stop
                    }

                    if ($vpnConfig) {
                        $results = @{
                            "Connection Name" = $vpnConfig.Name
                            "Server Address" = $vpnConfig.ServerAddress
                            "Connection Status" = $vpnConfig.ConnectionStatus
                            "Tunnel Type" = $vpnConfig.TunnelType
                            "Authentication Method" = $vpnConfig.AuthenticationMethod
                            "Split Tunneling" = $vpnConfig.SplitTunneling
                            "Encryption Level" = $vpnConfig.EncryptionLevel
                        }

                        # Test connectivity with timeout
                        $ports = @(500, 1701, 4500)
                        foreach ($port in $ports) {
                            $portType = switch($port) {
                                500 { "IKEEXT" }
                                1701 { "L2TP" }
                                4500 { "NAT-T" }
                            }
                            
                            $testResult = $false
                            $timeout = New-TimeSpan -Seconds 10
                            $sw = [Diagnostics.Stopwatch]::StartNew()
                            
                            while ($sw.Elapsed -lt $timeout) {
                                $test = Test-NetConnection -ComputerName $vpnConfig.ServerAddress -Port $port -WarningAction SilentlyContinue
                                if ($test.TcpTestSucceeded) {
                                    $testResult = $true
                                    break
                                }
                                Start-Sleep -Milliseconds 500
                            }
                            
                            $results["Port $port ($portType)"] = if($testResult) { "Open (Response time: $($sw.Elapsed.TotalSeconds) seconds)" } else { "Timeout after 10 seconds" }
                        }
                        return $results
                    }
                    return $null
                }
                catch {
                    return @{ "Error" = $_.Exception.Message }
                }
            } -ArgumentList $inputControls['VPN Name'].Text

            # Show progress form and wait for completion or cancellation
            $testForm.Show()
            while (-not $job.HasMoreData -and -not $testCancelled -and $job.State -eq 'Running') {
                Start-Sleep -Milliseconds 100
                [System.Windows.Forms.Application]::DoEvents()
            }

            if ($testCancelled) {
                Stop-Job $job
                Update-Status "VPN connection test cancelled by user"
            }
            else {
                $results = Receive-Job $job
                if ($results) {
                    if ($results.ContainsKey("Error")) {
                        Update-Status "Error testing VPN connection: $($results.Error)"
                    }
                    else {
                        $resultsText = ($results.GetEnumerator() | ForEach-Object { "$($_.Key): $($_.Value)" }) -join "`r`n"
                        Update-Status "Test Results:`r`n$resultsText"
                    }
                }
                else {
                    Update-Status "Could not find VPN connection: $($inputControls['VPN Name'].Text)"
                }
            }

            if ($job.State -eq 'Running') {
                Stop-Job $job
            }
            Remove-Job $job -Force
            $testForm.Close()
        })
        $buttonContainer.Controls.Add($testButton)
  
        # Diagnostics Tab
        $diagTab = $tabPages["Diagnostics"]
    
        $diagButtons = @(
            # Network Analysis
            @{
                'Text'   = 'Network Stack Info'
                'Action' = {
                    Update-Status "Gathering Network Stack Information..."
            
                    Update-Status "`nNetwork Adapters:"
                    Get-NetAdapter | ForEach-Object {
                        Update-Status "  $($_.Name):"
                        Update-Status "    Status: $($_.Status)"
                        Update-Status "    MAC Address: $($_.MacAddress)"
                        Update-Status "    Speed: $($_.LinkSpeed)"
                    }
            
                    Update-Status "`nIP Configuration:"
                    Get-NetIPConfiguration | ForEach-Object {
                        Update-Status "  Interface: $($_.InterfaceAlias)"
                        Update-Status "    IPv4: $($_.IPv4Address.IPAddress)"
                        Update-Status "    Gateway: $($_.IPv4DefaultGateway.NextHop)"
                        Update-Status "    DNS: $($_.DNSServer.ServerAddresses -join ', ')"
                    }
            
                    Update-Status "`nRouting Table:"
                    Get-NetRoute -Protocol NetMgmt | ForEach-Object {
                        Update-Status "  $($_.DestinationPrefix) via $($_.NextHop)"
                    }
                }
            },
            @{
                'Text'   = 'Test Network Connectivity'
                'Action' = {
                    Update-Status "Testing network connectivity..."
                    $tests = @(
                        "8.8.8.8",
                        $inputControls['Server Address'].Text
                    )
                    foreach ($test in $tests) {
                        $ping = Test-Connection -ComputerName $test -Count 1 -Quiet
                        Update-Status "Ping test to $test : $(if($ping){'Successful'}else{'Failed'})"
                    }
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
                        'RemoteAccess' = 'Routing and Remote Access - Provides VPN server functionality'
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
                'Text' = 'Analyze VPN Connection'
                'Action' = {
                    # Create analysis selection form
                    $analysisForm = New-Object System.Windows.Forms.Form
                    $analysisForm.Text = "Select VPN for Analysis"
                    $analysisForm.Size = New-Object System.Drawing.Size(400,150)
                    $analysisForm.StartPosition = 'CenterParent'
                    $analysisForm.FormBorderStyle = 'FixedDialog'
                    $analysisForm.MaximizeBox = $false
                    $analysisForm.MinimizeBox = $false

                    # VPN Selection ComboBox
                    $vpnLabel = New-Object System.Windows.Forms.Label
                    $vpnLabel.Location = New-Object System.Drawing.Point(10,20)
                    $vpnLabel.Size = New-Object System.Drawing.Size(100,20)
                    $vpnLabel.Text = "Select VPN:"
                    $analysisForm.Controls.Add($vpnLabel)

                    $vpnComboBox = New-Object System.Windows.Forms.ComboBox
                    $vpnComboBox.Location = New-Object System.Drawing.Point(120,20)
                    $vpnComboBox.Size = New-Object System.Drawing.Size(250,20)
                    $vpnComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

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
                    $analysisForm.Controls.Add($vpnComboBox)

                    # Analyze Button
                    $analyzeButton = New-Object System.Windows.Forms.Button
                    $analyzeButton.Location = New-Object System.Drawing.Point(10,60)
                    $analyzeButton.Size = New-Object System.Drawing.Size(360,30)
                    $analyzeButton.Text = "Start Analysis"
                    $analyzeButton.Add_Click({
                        $vpnName = $vpnComboBox.SelectedItem
                        if ($vpnName) {
                            # Create progress form
                            $progressForm = New-Object System.Windows.Forms.Form
                            $progressForm.Text = "Analyzing VPN Connection"
                            $progressForm.Size = New-Object System.Drawing.Size(400,150)
                            $progressForm.StartPosition = 'CenterParent'
                            $progressForm.FormBorderStyle = 'FixedDialog'
                            $progressForm.MaximizeBox = $false
                            $progressForm.MinimizeBox = $false

                            $progressBar = New-Object System.Windows.Forms.ProgressBar
                            $progressBar.Location = New-Object System.Drawing.Point(10,20)
                            $progressBar.Size = New-Object System.Drawing.Size(360,20)
                            $progressBar.Style = 'Marquee'
                            $progressForm.Controls.Add($progressBar)

                            $statusLabel = New-Object System.Windows.Forms.Label
                            $statusLabel.Location = New-Object System.Drawing.Point(10,50)
                            $statusLabel.Size = New-Object System.Drawing.Size(360,20)
                            $statusLabel.Text = "Analyzing connection..."
                            $progressForm.Controls.Add($statusLabel)

                            $cancelButton = New-Object System.Windows.Forms.Button
                            $cancelButton.Location = New-Object System.Drawing.Point(150,80)
                            $cancelButton.Size = New-Object System.Drawing.Size(75,23)
                            $cancelButton.Text = "Cancel"
                            $progressForm.Controls.Add($cancelButton)

                            $analysisCancelled = $false
                            $cancelButton.Add_Click({
                                $analysisCancelled = $true
                                $progressForm.Close()
                            })

                            # Create analysis job
                            $job = Start-Job -ScriptBlock {
                                param($vpnName)
                                try {
                                    # Try both standard and all-user connections
                                    $vpn = Get-VpnConnection -Name $vpnName -ErrorAction SilentlyContinue
                                    if (-not $vpn) {
                                        $vpn = Get-VpnConnection -Name $vpnName -AllUserConnection -ErrorAction Stop
                                    }

                                    $results = @{
                                        "Connection Details" = @{
                                            "Status" = $vpn.ConnectionStatus
                                            "Server Address" = $vpn.ServerAddress
                                            "Authentication Method" = $vpn.AuthenticationMethod
                                            "Encryption Level" = $vpn.EncryptionLevel
                                            "Split Tunneling" = $vpn.SplitTunneling
                                        }
                                        "Port Tests" = @{}
                                        "Event Logs" = @()
                                        "Network Interface" = @{}
                                    }

                                    # Test ports with timeout
                                    $ports = @{
                                        500 = 'IKE'
                                        4500 = 'NAT-T'
                                        1701 = 'L2TP'
                                    }
                                    foreach ($port in $ports.Keys) {
                                        $testResult = $false
                                        $timeout = New-TimeSpan -Seconds 10
                                        $sw = [Diagnostics.Stopwatch]::StartNew()
                                        
                                        while ($sw.Elapsed -lt $timeout) {
                                            $test = Test-NetConnection -ComputerName $vpn.ServerAddress -Port $port -WarningAction SilentlyContinue
                                            if ($test.TcpTestSucceeded) {
                                                $testResult = $true
                                                break
                                            }
                                            Start-Sleep -Milliseconds 500
                                        }
                                        
                                        $results.PortTests["$port ($($ports[$port]))"] = if($testResult) {
                                            "Open (Response time: $($sw.Elapsed.TotalSeconds) seconds)"
                                        } else {
                                            "Timeout after 10 seconds"
                                        }
                                    }

                                    # Get recent event logs
                                    $events = Get-WinEvent -FilterHashtable @{
                                        LogName = 'Application','System'
                                        StartTime = (Get-Date).AddHours(-1)
                                        Keywords = 'RasClient','PPP','L2TP','IKE'
                                    } -MaxEvents 10 -ErrorAction SilentlyContinue

                                    $results.EventLogs = $events | ForEach-Object {
                                        "$($_.TimeCreated) - $($_.Message)"
                                    }

                                    # Check network interface
                                    $vpnAdapter = Get-NetAdapter | Where-Object { 
                                        $_.InterfaceDescription -like "*WAN Miniport (L2TP)*" 
                                    }
                                    if ($vpnAdapter) {
                                        $results.NetworkInterface = @{
                                            "Status" = $vpnAdapter.Status
                                            "Speed" = $vpnAdapter.LinkSpeed
                                            "MediaState" = $vpnAdapter.MediaConnectionState
                                        }
                                    }

                                    return $results
                                }
                                catch {
                                    return @{ "Error" = $_.Exception.Message }
                                }
                            } -ArgumentList $vpnName

                            # Show progress form and monitor job
                            $progressForm.Show()
                            while (-not $job.HasMoreData -and -not $analysisCancelled -and $job.State -eq 'Running') {
                                Start-Sleep -Milliseconds 100
                                [System.Windows.Forms.Application]::DoEvents()
                            }

                            if ($analysisCancelled) {
                                Stop-Job $job
                                Update-Status "VPN analysis cancelled by user"
                            }
                            else {
                                $results = Receive-Job $job
                                if ($results) {
                                    if ($results.ContainsKey("Error")) {
                                        Update-Status "Error analyzing VPN: $($results.Error)"
                                    }
                                    else {
                                        Update-Status "Analysis Results for " + $vpnName + ":"
                                        
                                        Update-Status "`nConnection Details:"
                                        foreach ($detail in $results."Connection Details".GetEnumerator()) {
                                            Update-Status "  $($detail.Key): $($detail.Value)"
                                        }

                                        Update-Status "`nPort Tests:"
                                        foreach ($port in $results.PortTests.GetEnumerator()) {
                                            Update-Status "  $($port.Key): $($port.Value)"
                                        }

                                        if ($results.EventLogs.Count -gt 0) {
                                            Update-Status "`nRecent Event Logs:"
                                            foreach ($event in $results.EventLogs) {
                                                Update-Status "  $event"
                                            }
                                        }

                                        if ($results.NetworkInterface.Count -gt 0) {
                                            Update-Status "`nNetwork Interface:"
                                            foreach ($detail in $results.NetworkInterface.GetEnumerator()) {
                                                Update-Status "  $($detail.Key): $($detail.Value)"
                                            }
                                        }
                                    }
                                }
                                else {
                                    Update-Status "No results returned from analysis"
                                }
                            }

                            if ($job.State -eq 'Running') {
                                Stop-Job $job
                            }
                            Remove-Job $job -Force
                            $progressForm.Close()
                            $analysisForm.Close()
                        }
                        else {
                            [System.Windows.Forms.MessageBox]::Show(
                                "Please select a VPN connection to analyze",
                                "No VPN Selected",
                                [System.Windows.Forms.MessageBoxButtons]::OK,
                                [System.Windows.Forms.MessageBoxIcon]::Warning
                            )
                        }
                    })
                    $analysisForm.Controls.Add($analyzeButton)

                    # Show the form
                    $analysisForm.ShowDialog()
                }
            },
            @{
                'Text'   = 'Repair VPN Connection'
                'Action' = {
                    # Create repair options form
                    $repairForm = New-Object System.Windows.Forms.Form
                    $repairForm.Text = "VPN Repair Options"
                    $repairForm.Size = New-Object System.Drawing.Size(400, 300)
                    $repairForm.StartPosition = 'CenterParent'

                    # VPN Selection ComboBox
                    $vpnLabel = New-Object System.Windows.Forms.Label
                    $vpnLabel.Location = New-Object System.Drawing.Point(10, 20)
                    $vpnLabel.Size = New-Object System.Drawing.Size(100, 20)
                    $vpnLabel.Text = "Select VPN:"
                    $repairForm.Controls.Add($vpnLabel)

                    $vpnComboBox = New-Object System.Windows.Forms.ComboBox
                    $vpnComboBox.Location = New-Object System.Drawing.Point(120, 20)
                    $vpnComboBox.Size = New-Object System.Drawing.Size(250, 20)
                    $vpnComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
            
                    # Populate VPN list
                    Get-VpnConnection -ErrorAction SilentlyContinue | ForEach-Object {
                        $vpnComboBox.Items.Add($_.Name)
                    }
                    if ($vpnComboBox.Items.Count -gt 0) {
                        $vpnComboBox.SelectedIndex = 0
                    }
                    $repairForm.Controls.Add($vpnComboBox)

                    # Repair Options Group
                    $optionsGroup = New-Object System.Windows.Forms.GroupBox
                    $optionsGroup.Text = "Repair Options"
                    $optionsGroup.Location = New-Object System.Drawing.Point(10, 50)
                    $optionsGroup.Size = New-Object System.Drawing.Size(360, 150)
            
                    $options = @{
                        'RestartServices' = "Restart VPN Services"
                        'ClearDNS'        = "Clear DNS Cache"
                        'ResetIPSec'      = "Reset IPSec Policies"
                        'RecreateVPN'     = "Recreate VPN Connection"
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
                    $repairForm.Controls.Add($optionsGroup)

                    # Repair Button
                    $repairButton = New-Object System.Windows.Forms.Button
                    $repairButton.Location = New-Object System.Drawing.Point(10, 210)
                    $repairButton.Size = New-Object System.Drawing.Size(360, 30)
                    $repairButton.Text = "Start Repair"
                    $repairButton.Add_Click({
                            $vpnName = $vpnComboBox.SelectedItem
                            if ($vpnName) {
                                Update-Status "Starting repair for VPN connection: $vpnName"
                    
                                # Backup current configuration
                                Update-Status "Creating backup before repair..."
                                Backup-VPNConfiguration
                    
                                # Restart Services if selected
                                if ($checkboxes['RestartServices'].Checked) {
                                    $services = @('RasMan', 'RemoteAccess', 'PolicyAgent', 'IKEEXT')
                                    foreach ($service in $services) {
                                        Update-Status "  Restarting $service service..."
                                        Stop-Service -Name $service -Force -ErrorAction SilentlyContinue
                                        Start-Sleep -Seconds 2
                                        Start-Service -Name $service -ErrorAction SilentlyContinue
                                    }
                                }
                    
                                # Clear DNS if selected
                                if ($checkboxes['ClearDNS'].Checked) {
                                    Update-Status "  Clearing DNS cache..."
                                    Clear-DnsClientCache
                                }
                    
                                # Reset IPSec if selected
                                if ($checkboxes['ResetIPSec'].Checked) {
                                    Update-Status "  Resetting IPsec policies..."
                                    $null = netsh ipsec static delete all
                                }
                    
                                # Recreate VPN if selected
                                if ($checkboxes['RecreateVPN'].Checked) {
                                    Update-Status "  Recreating VPN connection..."
                                    $vpn = Get-VpnConnection -Name $vpnName
                                    $vpnParams = @{
                                        Name                 = $vpn.Name
                                        ServerAddress        = $vpn.ServerAddress
                                        TunnelType           = $vpn.TunnelType
                                        EncryptionLevel      = $vpn.EncryptionLevel
                                        AuthenticationMethod = $vpn.AuthenticationMethod
                                        SplitTunneling       = $vpn.SplitTunneling
                                        RememberCredential   = $vpn.RememberCredential
                                        Force                = $true
                                    }
                        
                                    Remove-VpnConnection -Name $vpnName -Force
                                    Add-VpnConnection @vpnParams
                                }
                    
                                Update-Status "VPN connection repair completed"
                                Update-VPNList
                                $repairForm.Close()
                            }
                            else {
                                Update-Status "Please select a VPN connection to repair"
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
            Width             = 180
            Height            = 30
            HorizontalSpacing = 10
            VerticalSpacing   = 10
            ButtonsPerColumn  = 5
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
        $resultsStartX = $buttonConfig.StartX + (2 * ($buttonConfig.Width + $buttonConfig.HorizontalSpacing))
        $diagResults = New-Object System.Windows.Forms.TextBox
        $diagResults.Location = New-Object System.Drawing.Point($resultsStartX, $buttonConfig.StartY)
        $diagResults.Size = New-Object System.Drawing.Size(350, 500)
        $diagResults.Multiline = $true
        $diagResults.ScrollBars = 'Vertical'
        $diagResults.ReadOnly = $true
        $diagTab.Controls.Add($diagResults)

        # Logs Tab
        $logsTab = $tabPages["Logs"]
    
        $logsTextBox = New-Object System.Windows.Forms.TextBox
        $logsTextBox.Location = New-Object System.Drawing.Point(10, 40)
        $logsTextBox.Size = New-Object System.Drawing.Size(740, 480)
        $logsTextBox.Multiline = $true
        $logsTextBox.ScrollBars = 'Vertical'
        $logsTextBox.ReadOnly = $true
    
        $refreshButton = New-Object System.Windows.Forms.Button
        $refreshButton.Location = New-Object System.Drawing.Point(10, 10)
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