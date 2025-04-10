# L2TP VPN Troubleshooter GUI

A comprehensive PowerShell GUI tool for managing, troubleshooting, and configuring L2TP VPN connections on Windows systems.

##Issues
Some features such as Test connection are not able to get the information from ALL-USERS VPN connections.


## Features

### VPN Connection Management
- Create new L2TP VPN connections with advanced configuration options
- View and manage existing VPN connections from multiple sources:
  - Windows VPN connections
  - All-user VPN connections
  - RAS Phone Book entries
  - Registry-stored connections
- View and modify existing VPN settings
- Delete VPN connections from all sources
- Test VPN connectivity with timeout controls

### Advanced Configuration Options
- PSK or Certificate-based authentication (Cert based auth in progress)
- Split tunneling configuration
- Authentication method selection (MSChapv2, EAP, PAP, CHAP)
- Encryption level settings (Default: None)
- IPv4/IPv6 configuration
- Idle disconnect timing
- Winlogon credential integration
- Remember credential option

### Diagnostic Tools
- Network Stack Information
  - Adapter details and status
  - IP configuration
  - Routing table information
  - Interface speeds
- VPN Prerequisites Check
  - Required services status
  - L2TP/IPsec components
  - Firewall rules
  - Registry settings
  - IPsec policy verification
- Network Connectivity Testing
  - Server reachability
  - DNS resolution
  - Port accessibility with 10-second timeouts
  - Cancellable connection tests
- Full VPN Connection Analysis
  - Connection status
  - Port accessibility
  - Event logs
  - Network interface details
  - Security associations
  - Connection statistics

### Network Repair Tools
- Guided WAN Miniport device removal process
- Network stack reset with detailed command output:
  - Winsock reset
  - IP stack reset
  - DNS cache clearing
  - IP release/renew
- VPN service configuration
- IPsec policy reset
- DNS cache management
- VPN connection repair with options:
  - Service restart
  - DNS cache clearing
  - IPsec policy reset
  - Connection recreation
  - Configuration backup

### Logging and Monitoring
- Comprehensive logging system
- Real-time status updates
- Error tracking
- Performance monitoring
- Log export and management
- Clear log functionality
- Separate error and general logs

## Requirements
- Windows Operating System
- PowerShell 5.1 or later
- Administrator privileges
- .NET Framework 4.5 or later

## Installation
1. Download the script file
2. Right-click and select "Run with PowerShell" as Administrator
   - Or launch PowerShell as Administrator and navigate to the script directory
   - Run: `.\VPNTroubleshooterGUI.ps1`

## Usage

### Creating a New VPN Connection
1. Navigate to the "Configuration" tab
2. Fill in the required basic settings:
   - VPN Name
   - Server Address
   - Authentication Type (PSK/Certificate)
   - Required credentials
3. Configure advanced settings if needed:
   - Split tunneling
   - Authentication method
   - Encryption level (Default: None)
   - IPv4/IPv6 settings
   - Idle disconnect
4. Click "Create New VPN"

### Managing Existing VPNs
1. Use the VPN Connection Manager section to:
   - View existing VPN connections from all sources
   - View and modify settings
   - Delete connections (removes from all locations)
   - Refresh the connection list

### Troubleshooting Existing VPN
1. Go to the "Diagnostics" tab
2. Choose from available diagnostic tools:
   - Network Stack Info
   - Test Network Connectivity
   - Check VPN Prerequisites
   - Analyze VPN Connection (with timeout controls)
   - Reset Network Stack
   - Repair VPN Connection

### Network Stack Reset
1. Select "Reset Network Stack"
2. Follow the guided process:
   - Manual WAN Miniport device removal in Device Manager
   - Automatic network stack reset
   - Service restart
   - System restart required after completion

### VPN Repair Process
1. Select "Repair VPN Connection"
2. Choose the VPN to repair
3. Select repair options:
   - Restart VPN services
   - Clear DNS cache
   - Reset IPSec policies
   - Recreate VPN connection
4. Monitor repair progress in status window

## Logging
- Logs are stored in `C:\VPNDiagnostics\`
  - General logs: `vpn.log`
  - Error logs: `error.log`
  - Performance logs: `performance.log`
- View logs in the "Logs" tab
- Refresh or clear logs as needed
- Real-time status updates

## Backup and Recovery
- Automatic backup before repairs
- Backup location: `C:\VPNBackup_[DateTime]`
- Includes:
  - VPN configurations
  - Registry settings
  - Network adapter settings

## Version
0.13

## Author
Brandon Cook 

## License
MIT