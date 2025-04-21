# L2TP VPN Troubleshooter GUI
Version 0.15 "OMFG MY HEAD HURTS EDITION"

## Overview
A comprehensive PowerShell-based GUI tool for troubleshooting, managing, and repairing L2TP VPN connections on Windows systems. This tool provides a user-friendly interface for common VPN management tasks and advanced diagnostics.

![CleanShot 2025-04-21 at 15 41 06](https://github.com/user-attachments/assets/ea0fcd5f-75f0-4912-ac1a-e04da0b4da59)
![CleanShot 2025-04-21 at 15 52 54](https://github.com/user-attachments/assets/2cc4cfcd-e99a-4756-8844-cb7121112525)


## Features

### VPN Connection Management
- View and manage existing VPN connections (both user-specific and all-user)
- Create new L2TP VPN connections with advanced configuration options
- Delete existing VPN connections
- View detailed VPN connection settings

### Diagnostic Tools
- Network Stack Information Analysis
- Comprehensive Network Testing
  - ICMP (Ping) Tests
  - DNS Resolution Tests
  - Traceroute Analysis
  - VPN Port Testing (500, 1701, 4500)
  - Network Interface Information
- VPN Prerequisites Check
  - Service Status Verification
  - L2TP/IPsec Component Check
  - Firewall Rule Analysis
  - Registry Settings Validation

### Repair Tools
- Network Stack Reset
  - WAN Miniport Device Management
  - Network Service Reset
  - IP Configuration Reset
- VPN Connection Repair
  - Backup Current Configuration
  - Service Restart
  - DNS Cache Clearing
  - Connection Recreation with Credential Management

### Advanced Features
- Automatic Logging System
  - Error Logging
  - Performance Logging
  - Diagnostic Information
- Backup and Restore Capabilities
- Credential Management
- Split Tunneling Configuration
- IPv4/IPv6 Settings Management

## Requirements
- Windows Operating System
- PowerShell 5.1 or later
- Administrative privileges
- .NET Framework 4.5 or later

## Installation
1. Download the script file (`VPNTroubleshooterGUI.ps1`)
2. Right-click the file and select "Run with PowerShell" or
3. Open PowerShell as Administrator and navigate to the script directory:
```powershell
.\VPNTroubleshooterGUI.ps1
```

## Usage

### Creating a New VPN Connection

1. Navigate to the **Configuration** tab.
2. Fill in the required fields:
   - **VPN Name**
   - **Server Address**
   - **Authentication Type** (PSK/Certificate)
   - **Pre-shared Key** (if using PSK)
3. Configure advanced settings if needed.
4. Click **Save VPN**.

---

### Troubleshooting an Existing Connection

1. Go to the **Diagnostics** tab.
2. Use **Network Stack Info** for basic connectivity information.
3. Run **Comprehensive Network Test** for detailed analysis.
4. Check VPN prerequisites using **Check VPN Prerequisites**.

---

### Repairing VPN Issues

1. Select **Repair VPN Connection**.
2. Choose the VPN connection to repair.
3. Enter the required credentials and PSK.
4. Select repair options.
5. Click **Start Repair**.

---

### Logging

- Logs are stored in `C:\VPNDiagnostics\`
- Three log types:
  - `vpn.log`: General operation logs
  - `error.log`: Error messages
  - `performance.log`: Performance metrics

---

### Backup Location

- VPN configuration backups are stored in:  
  `C:\VPNBackup_[DATE]_[TIME]\`

---

### Common Issues and Solutions

**WAN Miniport Missing**  
- Use the **Reset Network Stack** option  
- Follow the guided process to reinstall devices

**Service Issues**  
- Check VPN Prerequisites  
- Use the Repair tool to restart services

**Connection Failures**  
- Run Comprehensive Network Test  
- Verify port accessibility  
- Check PSK and credentials

---

### Notes

- Always run as Administrator
- Create a backup before making changes
- Some operations require a system restart
- Credential storage uses Windows Credential Manager

---

### Author

Brandon Cook

---

### License

MIT

---

### Support

maaybe
