# Office 365 Cloud Storage Provider Manager

A comprehensive PowerShell tool for managing cloud storage providers in Office 365 environments across all user profiles on a Windows system.

**Created by:** Brandon Cook  
**Email:** brandon@ghostinator.co  
**GitHub:** https://github.com/ghostinator/SysAdminPSSorcery

## Overview

This enterprise-ready PowerShell script provides system administrators with complete control over Office 365 cloud storage integration across all user profiles on a system. It features both a modern WPF GUI and command-line interface for flexible deployment scenarios.

## Key Features

### 🔧 Cloud Provider Management
- **Remove OneDrive** from startup and Office integration
- **Add OneDrive back** to startup for all users
- **Configure alternative cloud providers**: Dropbox, Google Drive, Box, ShareFile, Egnyte
- **Set default save locations** for Office applications
- **Registry-based configuration** that persists across user sessions

### 🖥️ Multiple Interface Options
- **Modern WPF GUI** with DataGrid for visual management
- **Text-based menu** for traditional console interaction
- **Command-line parameters** for scripting and automation
- **Silent operation mode** for Intune/MDM deployment

### 📊 Comprehensive Reporting
- **Real-time scanning** of all user profiles on the system
- **Detailed status display** showing current cloud provider configuration
- **Export capabilities** to CSV and professional HTML reports
- **Comprehensive logging** with timestamps

### 🔒 Enterprise Features
- **Auto-elevation** with UAC prompts
- **Password protection** (optional)
- **Registry hive loading** for offline user profile modification
- **Error handling** and recovery mechanisms

## System Requirements

- **Windows 10/11** or **Windows Server 2016+**
- **PowerShell 5.1** or later
- **Administrator privileges** (script auto-elevates)
- **.NET Framework 4.7.2** or later (for WPF GUI)

## Installation

1. Download the `Office365CloudStorageManager.ps1` script
2. Save to a secure location (e.g., `C:\Scripts\`)
3. Ensure PowerShell execution policy allows script execution:
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

## Usage Methods

### 1. GUI Mode (Default)

Launch the script without parameters to open the modern WPF interface:

```powershell
.\Office365CloudStorageManager.ps1
```

**GUI Features:**
- **Provider Selection**: Dropdown to choose cloud provider
- **Configuration Options**: Checkboxes for "Remove OneDrive" and "Set as Default"
- **User Status Grid**: Real-time view of all user profiles and their cloud configurations
- **Action Buttons**:
  - `Scan Users` - Refresh user profile data
  - `Add OneDrive to Startup` - Re-enable OneDrive for all users
  - `Run Configuration` - Apply selected settings
  - `Export` - Save reports to CSV or HTML
  - `Refresh` - Rescan user profiles

### 2. Command-Line Mode

Execute with parameters for automation and scripting:

```powershell
# Configure Dropbox as default, remove OneDrive
.\Office365CloudStorageManager.ps1 -Provider Dropbox -RemoveOneDrive -SetAsDefault -NoGUI

# Just remove OneDrive from startup
.\Office365CloudStorageManager.ps1 -RemoveOneDrive -NoGUI

# Configure Google Drive without removing OneDrive
.\Office365CloudStorageManager.ps1 -Provider GoogleDrive -SetAsDefault -NoGUI

# Set OneDrive as default (restore OneDrive integration)
.\Office365CloudStorageManager.ps1 -Provider OneDrive -SetAsDefault -NoGUI
```

### 3. Text Menu Mode

Access the traditional console menu:

```powershell
.\Office365CloudStorageManager.ps1
# Then select 'G' for GUI or use numbered options
```

**Menu Options:**
1. Remove OneDrive from startup
2. Add OneDrive to startup
3. Uninstall OneDrive completely
4. Configure Dropbox as cloud provider
5. Configure Google Drive as cloud provider
6. Configure Box as cloud provider
7. Configure Citrix ShareFile as cloud provider
8. Configure Egnyte as cloud provider
9. Configure OneDrive as cloud provider
G. Show GUI
0. Exit

### 4. Intune/MDM Deployment

For enterprise deployment, modify the configuration section:

```powershell
$UseHardcodedConfig = $true

$Config = @{
    Provider = "Dropbox"        # Set desired provider
    RemoveOneDrive = $true      # Remove OneDrive
    SetAsDefault = $true        # Set as default save location
    ShowGUI = $false           # Silent operation
    RequiredPassword = ""       # Optional password protection
}
```

## Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `-Provider` | String | Cloud provider to configure: `OneDrive`, `Dropbox`, `GoogleDrive`, `Box`, `ShareFile`, `Egnyte`, `None` |
| `-RemoveOneDrive` | Switch | Remove OneDrive from startup and disable integration |
| `-SetAsDefault` | Switch | Set the selected provider as default save location for Office |
| `-NoGUI` | Switch | Run in command-line mode without showing GUI |

## Supported Cloud Providers

| Provider | Office Integration | Startup Management | Default Save Location |
|----------|-------------------|-------------------|---------------------|
| **OneDrive** | ✅ | ✅ | ✅ |
| **Dropbox** | ✅ | ❌ | ✅ |
| **Google Drive** | ✅ | ❌ | ✅ |
| **Box** | ✅ | ❌ | ✅ |
| **ShareFile** | ✅ | ❌ | ✅ |
| **Egnyte** | ✅ | ❌ | ✅ |

## What the Script Does

### Registry Modifications
- Loads and modifies `NTUSER.DAT` files for all user profiles
- Configures Office 365 cloud provider settings in:
  - `HKCU\Software\Microsoft\Office\16.0\Common\Cloud`
  - `HKCU\Software\Microsoft\Office\16.0\Common\Cloud\Backstage`
- Manages startup entries in:
  - `HKCU\Software\Microsoft\Windows\CurrentVersion\Run`

### File System Operations
- Creates/removes startup shortcuts in user startup folders
- Manages All Users startup folder entries
- Optionally removes OneDrive installation files

### User Profile Processing
- Scans all local user profiles with `NTUSER.DAT` files
- Excludes system profiles (Default, Public, etc.)
- Safely loads and unloads registry hives
- Applies configurations even for users not currently logged in

## Export and Reporting

### CSV Export
- Clean tabular data for analysis
- Includes all user profile information
- Compatible with Excel and other spreadsheet applications

### HTML Report
- Professional formatted report with styling
- Includes generation timestamp and summary statistics
- Embedded branding and contact information
- Suitable for management reporting

### Log Files
- Comprehensive logging to `%TEMP%\Office365CloudStorageManager_[timestamp].log`
- Includes all operations, errors, and status updates
- Useful for troubleshooting and audit trails

## Security Features

### Password Protection
Optional password protection can be enabled in the configuration section:

```powershell
$Config = @{
    RequiredPassword = "YourSecurePassword"
}
```

### UAC Integration
- Automatically elevates to Administrator privileges
- Preserves command-line arguments during elevation
- Graceful handling of elevation failures

### Registry Safety
- Safe registry hive loading and unloading
- Error recovery mechanisms
- Temporary hive naming to prevent conflicts

## Common Use Cases

### 1. Enterprise OneDrive Replacement
```powershell
# Replace OneDrive with Dropbox across all users
.\Office365CloudStorageManager.ps1 -Provider Dropbox -RemoveOneDrive -SetAsDefault -NoGUI
```

### 2. Audit Current Configuration
```powershell
# Launch GUI to scan and export current cloud provider status
.\Office365CloudStorageManager.ps1
# Use Export button to generate reports
```

### 3. Restore OneDrive Integration
```powershell
# Re-enable OneDrive after it was disabled
.\Office365CloudStorageManager.ps1 -Provider OneDrive -SetAsDefault -NoGUI
```

### 4. Multi-Provider Environment
```powershell
# Configure Box without removing existing OneDrive
.\Office365CloudStorageManager.ps1 -Provider Box -SetAsDefault -NoGUI
```

## Troubleshooting

### Common Issues

**Script won't run:**
- Ensure PowerShell execution policy allows script execution
- Run PowerShell as Administrator
- Check that the script file isn't blocked (Right-click → Properties → Unblock)

**Registry access errors:**
- Verify Administrator privileges
- Ensure no other processes are accessing user registry hives
- Check for corrupted user profiles

**OneDrive executable not found:**
- Script checks both `System32` and `SysWOW64` locations
- Verify OneDrive is installed on the system
- Check Windows version compatibility

### Log Analysis
Review the log file at `%TEMP%\Office365CloudStorageManager_[timestamp].log` for detailed operation information and error messages.

## Best Practices

1. **Test in Lab Environment**: Always test configurations in a non-production environment first
2. **Backup Registry**: Consider creating system restore points before major changes
3. **User Communication**: Inform users about cloud provider changes before deployment
4. **Gradual Rollout**: Deploy to small groups initially, then expand
5. **Monitor Logs**: Review log files for any issues during deployment

## Version History

- **v1.0** - Initial release with basic OneDrive removal
- **v2.0** - Added multi-provider support and GUI interface
- **v3.0** - Added WPF GUI, export features, and OneDrive restore capability
- **v3.1** - Current version with enhanced branding and comprehensive documentation

## Support

For issues, feature requests, or contributions:
- **Email**: brandon@ghostinator.co
- **GitHub**: https://github.com/ghostinator/SysAdminPSSorcery
- **Issues**: Submit via GitHub Issues page

## License

This script is provided as-is for educational and administrative purposes. Use at your own risk and ensure compliance with your organization's policies and Microsoft's terms of service.

**⚠️ Important Notice**: This tool modifies system registry settings and user configurations. Always test thoroughly in a lab environment before production deployment. Ensure you have appropriate backups and recovery procedures in place.