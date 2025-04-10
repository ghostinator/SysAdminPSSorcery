# L2TP VPN Troubleshooter GUI

A comprehensive PowerShell GUI tool for managing, troubleshooting, and configuring L2TP VPN connections on Windows systems.

## Features

### VPN Connection Management
- Create new L2TP VPN connections with advanced configuration options
- View and manage existing VPN connections from multiple sources:
  - Windows VPN connections
  - RAS Phone Book entries
  - Registry-stored connections
- Delete VPN connections
- Test VPN connectivity
- View detailed VPN settings

### Advanced Configuration Options
- PSK or Certificate-based authentication
- Split tunneling configuration
- Authentication method selection (MSChapv2, EAP, PAP, CHAP)
- Encryption level settings (Required, Optional, None, Maximum)
- IPv4/IPv6 configuration
- Idle disconnect timing
- Winlogon credential integration
- Remember credential option

### Diagnostic Tools
- Network Stack Information
  - Adapter details
  - IP configuration
  - Routing table
- VPN Prerequisites Check
  - Required services status
  - L2TP/IPsec components
  - Firewall rules
  - Registry settings
- Network Connectivity Testing
- Full VPN Connection Analysis
  - Connection status
  - Port accessibility
  - Event logs
  - Network interface details

### Network Repair Tools
- Guided WAN Miniport device removal
- Network stack reset with detailed command output
- VPN service configuration
- IPsec policy reset
- DNS cache management
- VPN connection repair with options:
  - Service restart
  - DNS cache clearing
  - IPsec policy reset
  - Connection recreation

### Logging and Monitoring
- Comprehensive logging system
- Real-time status updates
- Error tracking
- Performance monitoring
- Log export and management
- Clear log functionality

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
3. Configure advanced settings if needed
4. Click "Create New VPN"

### Managing Existing VPNs
1. Use the VPN Connection Manager section to:
   - View existing VPN connections
   - View detailed settings
   - Delete connections
   - Refresh the connection list

### Troubleshooting Existing VPN
1. Go to the "Diagnostics" tab
2. Choose from available diagnostic tools:
   - Network Stack Info
   - Test Network Connectivity
   - Check VPN Prerequisites
   - Analyze VPN Connection
   - Reset Network Stack
   - Repair VPN Connection

### Network Stack Reset
1. Select