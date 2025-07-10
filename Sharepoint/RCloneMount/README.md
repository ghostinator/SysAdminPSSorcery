# SharePoint/OneDrive Mount Tools

This folder contains PowerShell scripts for mounting SharePoint Online and OneDrive folders to local drive letters using RClone.

## Files

- **SharePointOneDriveGUI.ps1** - User-friendly GUI application for mounting SharePoint/OneDrive
- **MountSharePointinRClone.ps1** - Command-line version for scripting and automation
- **README.md** - This documentation file

## Prerequisites

- Windows PowerShell 5.1 or later
- RClone (will be automatically installed by the GUI version if not present)
- Microsoft 365 account with access to SharePoint/OneDrive

## GUI Version (Recommended for End Users)

### Features

- **Easy-to-use graphical interface**
- **Automatic RClone installation** if not present
- **Support for multiple service types:**
  - SharePoint Online
  - OneDrive for Business
  - OneDrive Personal
- **Real-time mount status tracking**
- **Automatic drive letter management**
- **Built-in error handling and validation**

### Usage

1. Run the script:
   ```powershell
   .\SharePointOneDriveGUI.ps1
   ```

2. **Select Service Type:**
   - SharePoint Online (requires site URL)
   - OneDrive for Business
   - OneDrive Personal

3. **Configure Settings:**
   - **Remote Name:** Friendly name for your connection (e.g., "MySharePoint")
   - **SharePoint Site URL:** Full URL to your SharePoint site (for SharePoint only)
   - **Tenant ID:** Optional - your Microsoft 365 tenant ID
   - **Remote Path:** Optional - specific folder path within SharePoint/OneDrive
   - **Drive Letter:** Available drive letter to mount to

4. **Configure Remote:**
   - Click "Configure Remote" to set up authentication
   - Follow the RClone configuration prompts
   - Complete Microsoft 365 authentication in your browser

5. **Mount Drive:**
   - Click "Mount Drive" to create the connection
   - Your SharePoint/OneDrive content will appear as a local drive

6. **Manage Mounts:**
   - View current mounts in the status area
   - Select and unmount drives as needed
   - Use "Refresh" to update the interface

## Command Line Version (For Automation)

### Usage Examples

**Mount SharePoint Site:**
```powershell
.\MountSharePointinRClone.ps1 -RemoteName "CompanySharePoint" -ServiceType "sharepoint" -SiteUrl "https://contoso.sharepoint.com/sites/documents" -DriveLetter "S"
```

**Mount OneDrive for Business:**
```powershell
.\MountSharePointinRClone.ps1 -RemoteName "MyOneDrive" -ServiceType "onedrive" -DriveLetter "O"
```

**Mount Specific Folder:**
```powershell
.\MountSharePointinRClone.ps1 -RemoteName "ProjectFiles" -ServiceType "sharepoint" -SiteUrl "https://contoso.sharepoint.com/sites/projects" -DriveLetter "P" -RemotePath "Documents/ActiveProjects"
```

**Configure Remote Only:**
```powershell
.\MountSharePointinRClone.ps1 -RemoteName "MySharePoint" -ServiceType "sharepoint" -SiteUrl "https://contoso.sharepoint.com/sites/mysite" -Configure
```

**Unmount Drive:**
```powershell
.\MountSharePointinRClone.ps1 -RemoteName "MySharePoint" -DriveLetter "S" -Unmount
```

### Parameters

- **RemoteName** (Required): Name for the RClone remote configuration
- **ServiceType** (Required): "sharepoint", "onedrive", or "onedrive-personal"
- **SiteUrl** (Required for SharePoint): Full SharePoint site URL
- **DriveLetter** (Required): Drive letter to mount (D-Z)
- **RemotePath** (Optional): Specific folder path within the remote
- **TenantId** (Optional): Microsoft 365 tenant ID
- **Configure** (Switch): Configure remote without mounting
- **Unmount** (Switch): Unmount the specified drive

## Authentication

Both scripts use RClone's built-in Microsoft 365 authentication:

1. **Initial Setup:** RClone will open your web browser
2. **Sign In:** Use your Microsoft 365 credentials
3. **Grant Permissions:** Allow RClone to access your SharePoint/OneDrive
4. **Token Storage:** Authentication tokens are securely stored by RClone

## Mount Options

The scripts use optimized RClone mount settings:

- **VFS Cache Mode:** `writes` - Enables local caching for better performance
- **Cache Max Age:** `1h` - Files cached for 1 hour
- **Cache Max Size:** `1G` - Maximum 1GB cache size
- **Buffer Size:** `32M` - 32MB buffer for transfers
- **Timeout:** `1h` - 1 hour timeout for operations
- **Retries:** `3` attempts with `10` low-level retries

## Troubleshooting

### Common Issues

**RClone Not Found:**
- The GUI version will automatically install RClone
- For command line, download from: https://rclone.org/downloads/

**Authentication Failures:**
- Ensure you have proper permissions to the SharePoint site
- Check that your Microsoft 365 account is active
- Try reconfiguring the remote with fresh authentication

**Mount Failures:**
- Verify the drive letter is not already in use
- Check that the SharePoint site URL is correct and accessible
- Ensure you have network connectivity

**Performance Issues:**
- Large files may take time to appear due to caching
- Consider adjusting cache settings for your use case
- Network speed affects transfer performance

### Getting Help

**View RClone Configuration:**
```powershell
rclone config show
```

**List Configured Remotes:**
```powershell
rclone listremotes
```

**Test Remote Connection:**
```powershell
rclone ls RemoteName:
```

**View RClone Logs:**
```powershell
rclone mount RemoteName: Z: --log-level DEBUG
```

## Security Considerations

- **Authentication tokens** are stored in RClone's configuration file
- **Local cache** may contain copies of your files
- **Network traffic** is encrypted using Microsoft's APIs
- **Unmount drives** when not in use to free resources

## Advanced Configuration

### Custom RClone Options

You can modify the mount arguments in the scripts to customize behavior:

```powershell
# Example: Increase cache size
--vfs-cache-max-size 5G

# Example: Disable caching
--vfs-cache-mode off

# Example: Enable debug logging
--log-level DEBUG
```

### Batch Operations

Create batch files to mount multiple drives:

```batch
@echo off
powershell -ExecutionPolicy Bypass -File "MountSharePointinRClone.ps1" -RemoteName "SharePoint1" -ServiceType "sharepoint" -SiteUrl "https://contoso.sharepoint.com/sites/site1" -DriveLetter "S"
powershell -ExecutionPolicy Bypass -File "MountSharePointinRClone.ps1" -RemoteName "OneDrive1" -ServiceType "onedrive" -DriveLetter "O"
```

## Support

For issues with:
- **RClone:** Visit https://rclone.org/
- **Microsoft 365:** Contact your IT administrator
- **These Scripts:** Check the troubleshooting section above

## Version History

- **v1.0** - Initial release with GUI and command-line versions
- Support for SharePoint Online, OneDrive for Business, and OneDrive Personal
- Automatic RClone installation and configuration
- Optimized mount settings for performance