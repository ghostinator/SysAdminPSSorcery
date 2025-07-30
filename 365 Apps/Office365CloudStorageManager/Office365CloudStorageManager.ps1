#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Office 365 Cloud Storage Provider Manager for Business
.DESCRIPTION
    This script manages cloud storage providers for Office 365 applications in business environments.
    Features:
    - Remove OneDrive from startup locations for all users
    - Configure enterprise cloud storage providers for Office 365
    - Set default save locations for Office applications
    - GUI interface for easy management
    - Scan and display current cloud provider configuration for all users
.PARAMETER Provider
    The cloud provider to configure: OneDrive, Dropbox, GoogleDrive, Box, ShareFile, Egnyte, or None
.PARAMETER RemoveOneDrive
    Switch to remove OneDrive completely
.PARAMETER SetAsDefault
    Switch to set the selected provider as default for Office applications
.EXAMPLE
    .\Office365CloudStorageManager.ps1 -Provider Dropbox -RemoveOneDrive -SetAsDefault
.NOTES
    Run as Administrator to access all user profiles and registry hives
#>

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateSet("OneDrive", "Dropbox", "GoogleDrive", "Box", "ShareFile", "Egnyte", "None")]
    [string]$Provider = "None",

    [Parameter()]
    [switch]$RemoveOneDrive,

    [Parameter()]
    [switch]$SetAsDefault
)

#region CONFIGURATION SECTION - MODIFY FOR INTUNE DEPLOYMENT
# Set to $true to enable this configuration section (for Intune deployment)
# Set to $false to use command-line parameters or interactive menu
$UseHardcodedConfig = $false

# Hard-coded configuration options (only used when $UseHardcodedConfig = $true)
$Config = @{
    # Set to one of: "OneDrive", "Dropbox", "GoogleDrive", "Box", "ShareFile", "Egnyte", "None"
    Provider = "Dropbox"

    # Set to $true to remove OneDrive completely, $false to keep it
    RemoveOneDrive = $true

    # Set to $true to set the selected provider as default for Office, $false otherwise
    SetAsDefault = $true

    # Set to $true to show interactive menu, $false for silent operation
    # Note: For Intune, this should typically be $false
    ShowMenu = $false
}
#endregion

# Configuration
$ErrorActionPreference = "Continue"
$LogPath = "$env:TEMP"
$LogFile = "$LogPath\Office365CloudStorageManager_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"

function Write-Log {
    param([string]$Message)
    $Time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$Time $Message" | Out-File -FilePath $LogFile -Append -Encoding UTF8
    Write-Host $Message
}

function Get-CloudProviderForAllUsers {
    $results = @()

    $userProfiles = Get-ChildItem 'C:\Users' -Directory | Where-Object {
        (Test-Path "$($_.FullName)\NTUSER.DAT") -and
        ($_.Name -notin @('Default', 'Public', 'All Users', 'Default User'))
    }

    Write-Log "Found $($userProfiles.Count) user profiles to scan"

    foreach ($userProfile in $userProfiles) {
        $userName = $userProfile.Name
        $userPath = $userProfile.FullName
        $ntUserDat = "$userPath\NTUSER.DAT"
        $tempHive = "HKU\TempHive_Scan_$userName"

        try {
            $loadResult = reg load $tempHive $ntUserDat 2>&1
            if ($LASTEXITCODE -eq 0) {
                $officeCloudPath = "Registry::$tempHive\Software\Microsoft\Office\16.0\Common\Cloud"
                $defaultCloud = "None"
                $defaultSaveLoc = "None"

                if (Test-Path $officeCloudPath) {
                    $props = Get-ItemProperty -Path $officeCloudPath -ErrorAction SilentlyContinue
                    if ($props) {
                        $defaultCloud = if ($props.DefaultCloudProvider) { $props.DefaultCloudProvider } else { "None" }
                        $defaultSaveLoc = if ($props.DefaultSaveLocation) { $props.DefaultSaveLocation } else { "None" }
                    }
                }

                $results += [PSCustomObject]@{
                    User         = $userName
                    CloudProvider = $defaultCloud
                    SaveLocation  = $defaultSaveLoc
                }

                reg unload $tempHive | Out-Null
            } else {
                $results += [PSCustomObject]@{
                    User          = $userName
                    CloudProvider = "[Failed to load registry]"
                    SaveLocation  = ""
                }
            }
        } catch {
            reg unload $tempHive 2>&1 | Out-Null
            $results += [PSCustomObject]@{
                User          = $userName
                CloudProvider = "[ERROR: $($_.Exception.Message)]"
                SaveLocation  = ""
            }
        }
    }

    return $results
}

function Remove-OneDriveStartup {
    Write-Log "Starting OneDrive startup removal for all existing users..."

    # Get all user profiles with NTUSER.DAT files (existing users)
    $userProfiles = Get-ChildItem 'C:\Users' -Directory | Where-Object {
        (Test-Path "$($_.FullName)\NTUSER.DAT") -and
        ($_.Name -notin @('Default', 'Public', 'All Users', 'Default User'))
    }

    Write-Log "Found $($userProfiles.Count) user profiles to process"

    $totalRemoved = 0

    foreach ($userProfile in $userProfiles) {
        $userName = $userProfile.Name
        $userPath = $userProfile.FullName
        Write-Log "Processing user: $userName"

        # --- A. Remove from Registry (HKCU Run Key) ---
        try {
            $ntUserDat = "$userPath\NTUSER.DAT"
            $tempHive = "HKU\TempHive_$userName"
            
            # Load the user's registry hive
            $loadResult = reg load $tempHive $ntUserDat 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Log "Loaded registry hive for $userName"
                
                # Check if OneDrive entry exists in Run key
                $runKeyPath = "Registry::$tempHive\Software\Microsoft\Windows\CurrentVersion\Run"
                if (Test-Path $runKeyPath) {
                    $runKey = Get-ItemProperty -Path $runKeyPath -ErrorAction SilentlyContinue
                    if ($runKey -and $runKey.PSObject.Properties.Name -contains "OneDrive") {
                        Remove-ItemProperty -Path $runKeyPath -Name "OneDrive" -ErrorAction Stop
                        Write-Log "Removed OneDrive from Run registry key for $userName"
                        $totalRemoved++
                    } else {
                        Write-Log "OneDrive not found in Run registry key for $userName"
                    }
                }
                
                # Unload the registry hive
                reg unload $tempHive | Out-Null
                Write-Log "Unloaded registry hive for $userName"
            } else {
                Write-Log "Failed to load registry hive for $userName`: $loadResult"
            }
        } catch {
            Write-Log "Error processing registry for $userName`: $($_.Exception.Message)"
            # Attempt to unload hive if it was loaded
            reg unload $tempHive 2>&1 | Out-Null
        }
        
        # --- B. Remove from Startup Folders ---
        
        # Remove from user's personal startup folder
        $userStartupFolder = "$userPath\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
        $userOneDriveLnk = Join-Path $userStartupFolder "OneDrive.lnk"
        
        if (Test-Path $userOneDriveLnk) {
            try {
                Remove-Item $userOneDriveLnk -Force
                Write-Log "Removed OneDrive shortcut from startup folder for $userName"
                $totalRemoved++
            } catch {
                Write-Log "Failed to remove OneDrive shortcut from startup folder for $userName`: $($_.Exception.Message)"
            }
        } else {
            Write-Log "OneDrive shortcut not found in startup folder for $userName"
        }
    }

    # Remove from All Users Startup Folder
    $allUsersStartup = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp"
    $allUsersOneDriveLnk = Join-Path $allUsersStartup "OneDrive.lnk"

    if (Test-Path $allUsersOneDriveLnk) {
        try {
            Remove-Item $allUsersOneDriveLnk -Force
            Write-Log "Removed OneDrive shortcut from All Users startup folder"
            $totalRemoved++
        } catch {
            Write-Log "Failed to remove OneDrive shortcut from All Users startup folder: $($_.Exception.Message)"
        }
    } else {
        Write-Log "OneDrive shortcut not found in All Users startup folder"
    }

    # Summary
    Write-Log "OneDrive startup removal completed for all existing users."
    Write-Log "Total startup entries removed: $totalRemoved"
    
    return $totalRemoved
}

function Uninstall-OneDrive {
    Write-Log "Starting OneDrive uninstallation process..."
    
    # Kill OneDrive process if running
    Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue | Stop-Process -Force
    Write-Log "Stopped OneDrive processes"
    
    # Uninstall OneDrive
    $oneDriveSetup = "$env:SystemRoot\SysWOW64\OneDriveSetup.exe"
    if (-not (Test-Path $oneDriveSetup)) {
        $oneDriveSetup = "$env:SystemRoot\System32\OneDriveSetup.exe"
    }
    
    if (Test-Path $oneDriveSetup) {
        Write-Log "Running OneDrive uninstaller..."
        Start-Process $oneDriveSetup "/uninstall" -NoNewWindow -Wait
        Write-Log "OneDrive uninstaller completed"
    } else {
        Write-Log "OneDrive setup executable not found"
    }
    
    # Remove OneDrive folder
    $oneDriveFolder = "$env:USERPROFILE\OneDrive"
    if (Test-Path $oneDriveFolder) {
        try {
            Remove-Item $oneDriveFolder -Force -Recurse -ErrorAction Stop
            Write-Log "Removed OneDrive folder: $oneDriveFolder"
        } catch {
            Write-Log "Failed to remove OneDrive folder: $($_.Exception.Message)"
        }
    }
    
    # Disable OneDrive via Group Policy Registry settings
    $regPaths = @(
        "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions\{A52BBA46-E9E1-435f-B3D9-28DAA648C0F6}",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions\{A52BBA46-E9E1-435f-B3D9-28DAA648C0F6}"
    )
    
    foreach ($path in $regPaths) {
        if (-not (Test-Path $path)) {
            New-Item -Path $path -Force | Out-Null
        }
    }
    
    # Disable OneDrive
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive" -Name "DisableFileSyncNGSC" -Value 1 -Type DWord
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive" -Name "DisableFileSync" -Value 1 -Type DWord
    
    Write-Log "OneDrive uninstallation and disabling completed"
}

function Configure-CloudProvider {
    param (
        [string]$Provider,
        [bool]$SetAsDefault
    )
    
    Write-Log "Configuring cloud provider: $Provider (Set as default: $SetAsDefault)"
    
    # Get all user profiles with NTUSER.DAT files (existing users)
    $userProfiles = Get-ChildItem 'C:\Users' -Directory | Where-Object {
        (Test-Path "$($_.FullName)\NTUSER.DAT") -and
        ($_.Name -notin @('Default', 'Public', 'All Users', 'Default User'))
    }

    foreach ($userProfile in $userProfiles) {
        $userName = $userProfile.Name
        $userPath = $userProfile.FullName
        Write-Log "Configuring cloud provider for user: $userName"

        try {
            $ntUserDat = "$userPath\NTUSER.DAT"
            $tempHive = "HKU\TempHive_$userName"

            # Load the user's registry hive
            $loadResult = reg load $tempHive $ntUserDat 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Log "Loaded registry hive for $userName"
                
                # Office cloud storage provider settings
                $officeCloudPath = "Registry::$tempHive\Software\Microsoft\Office\16.0\Common\Cloud"
                
                if (-not (Test-Path $officeCloudPath)) {
                    New-Item -Path $officeCloudPath -Force | Out-Null
                }
                
                # First, disable all providers to avoid conflicts
                $allProviders = @("OneDrive", "Dropbox", "GoogleDrive", "Box", "ShareFile", "Egnyte")
                foreach ($p in $allProviders) {
                    $enableKey = "Enable${p}InOffice"
                    if ($p -ne $Provider) {
                        Set-ItemProperty -Path $officeCloudPath -Name $enableKey -Value 0 -Type DWord -ErrorAction SilentlyContinue
                    }
                }
                
                # Configure based on provider
                switch ($Provider) {
                    "Dropbox" {
                        # Enable Dropbox integration
                        Set-ItemProperty -Path $officeCloudPath -Name "EnableDropboxInOffice" -Value 1 -Type DWord
                        
                        if ($SetAsDefault) {
                            Set-ItemProperty -Path $officeCloudPath -Name "DefaultCloudProvider" -Value "Dropbox" -Type String
                            Set-ItemProperty -Path $officeCloudPath -Name "DefaultSaveLocation" -Value "Dropbox" -Type String
                        }
                        
                        Write-Log "Enabled Dropbox integration for $userName"
                    }
                    "GoogleDrive" {
                        # Enable Google Drive integration
                        Set-ItemProperty -Path $officeCloudPath -Name "EnableGoogleDriveInOffice" -Value 1 -Type DWord
                        
                        if ($SetAsDefault) {
                            Set-ItemProperty -Path $officeCloudPath -Name "DefaultCloudProvider" -Value "GoogleDrive" -Type String
                            Set-ItemProperty -Path $officeCloudPath -Name "DefaultSaveLocation" -Value "GoogleDrive" -Type String
                        }
                        
                        Write-Log "Enabled Google Drive integration for $userName"
                    }
                    "Box" {
                        # Enable Box integration
                        Set-ItemProperty -Path $officeCloudPath -Name "EnableBoxInOffice" -Value 1 -Type DWord
                        
                        if ($SetAsDefault) {
                            Set-ItemProperty -Path $officeCloudPath -Name "DefaultCloudProvider" -Value "Box" -Type String
                            Set-ItemProperty -Path $officeCloudPath -Name "DefaultSaveLocation" -Value "Box" -Type String
                        }
                        
                        Write-Log "Enabled Box integration for $userName"
                    }
                    "ShareFile" {
                        # Enable Citrix ShareFile integration
                        Set-ItemProperty -Path $officeCloudPath -Name "EnableShareFileInOffice" -Value 1 -Type DWord
                        
                        if ($SetAsDefault) {
                            Set-ItemProperty -Path $officeCloudPath -Name "DefaultCloudProvider" -Value "ShareFile" -Type String
                            Set-ItemProperty -Path $officeCloudPath -Name "DefaultSaveLocation" -Value "ShareFile" -Type String
                        }
                        
                        Write-Log "Enabled ShareFile integration for $userName"
                    }
                    "Egnyte" {
                        # Enable Egnyte integration
                        Set-ItemProperty -Path $officeCloudPath -Name "EnableEgnyteInOffice" -Value 1 -Type DWord
                        
                        if ($SetAsDefault) {
                            Set-ItemProperty -Path $officeCloudPath -Name "DefaultCloudProvider" -Value "Egnyte" -Type String
                            Set-ItemProperty -Path $officeCloudPath -Name "DefaultSaveLocation" -Value "Egnyte" -Type String
                        }
                        
                        Write-Log "Enabled Egnyte integration for $userName"
                    }
                    "OneDrive" {
                        # Enable OneDrive integration
                        Set-ItemProperty -Path $officeCloudPath -Name "EnableOneDriveInOffice" -Value 1 -Type DWord
                        
                        if ($SetAsDefault) {
                            Set-ItemProperty -Path $officeCloudPath -Name "DefaultCloudProvider" -Value "OneDrive" -Type String
                            Set-ItemProperty -Path $officeCloudPath -Name "DefaultSaveLocation" -Value "OneDrive" -Type String
                        }
                        
                        Write-Log "Enabled OneDrive integration for $userName"
                    }
                }
                
                # Additional Office integration settings
                $officeBackstagePath = "Registry::$tempHive\Software\Microsoft\Office\16.0\Common\Cloud\Backstage"
                if (-not (Test-Path $officeBackstagePath)) {
                    New-Item -Path $officeBackstagePath -Force | Out-Null
                }
                
                # Enable the selected provider in Office backstage view
                if ($Provider -ne "None") {
                    Set-ItemProperty -Path $officeBackstagePath -Name "Show$Provider" -Value 1 -Type DWord -ErrorAction SilentlyContinue
                }
                
                # Unload the registry hive
                reg unload $tempHive | Out-Null
                Write-Log "Unloaded registry hive for $userName"
            } else {
                Write-Log "Failed to load registry hive for $userName`: $loadResult"
            }
        } catch {
            Write-Log "Error configuring cloud provider for $userName`: $($_.Exception.Message)"
            # Attempt to unload hive if it was loaded
            reg unload $tempHive 2>&1 | Out-Null
        }
    }
    
    Write-Log "Cloud provider configuration completed for all users"
}

function Show-Menu {
    Clear-Host
    Write-Host "===== Office 365 Cloud Storage Provider Manager =====" -ForegroundColor Cyan
    Write-Host "1: Remove OneDrive from startup"
    Write-Host "2: Uninstall OneDrive completely"
    Write-Host "3: Configure Dropbox as cloud provider"
    Write-Host "4: Configure Google Drive as cloud provider"
    Write-Host "5: Configure Box as cloud provider"
    Write-Host "6: Configure Citrix ShareFile as cloud provider"
    Write-Host "7: Configure Egnyte as cloud provider"
    Write-Host "8: Configure OneDrive as cloud provider"
    Write-Host "9: Exit"
    Write-Host "===================================================" -ForegroundColor Cyan
    
    $choice = Read-Host "Enter your choice (1-9)"
    
    switch ($choice) {
        "1" {
            $removed = Remove-OneDriveStartup
            Write-Host "Removed $removed OneDrive startup entries" -ForegroundColor Green
            Pause
            Show-Menu
        }
        "2" {
            Remove-OneDriveStartup
            Uninstall-OneDrive
            Write-Host "OneDrive has been uninstalled and disabled" -ForegroundColor Green
            Pause
            Show-Menu
        }
        "3" {
            $setDefault = Read-Host "Set Dropbox as default save location? (y/n)"
            $removeOD = Read-Host "Remove OneDrive? (y/n)"
            
            if ($removeOD -eq "y") {
                Remove-OneDriveStartup
                Uninstall-OneDrive
            }
            
            Configure-CloudProvider -Provider "Dropbox" -SetAsDefault ($setDefault -eq "y")
            Write-Host "Dropbox has been configured" -ForegroundColor Green
            Pause
            Show-Menu
        }
        "4" {
            $setDefault = Read-Host "Set Google Drive as default save location? (y/n)"
            $removeOD = Read-Host "Remove OneDrive? (y/n)"
            
            if ($removeOD -eq "y") {
                Remove-OneDriveStartup
                Uninstall-OneDrive
            }
            
            Configure-CloudProvider -Provider "GoogleDrive" -SetAsDefault ($setDefault -eq "y")
            Write-Host "Google Drive has been configured" -ForegroundColor Green
            Pause
            Show-Menu
        }
        "5" {
            $setDefault = Read-Host "Set Box as default save location? (y/n)"
            $removeOD = Read-Host "Remove OneDrive? (y/n)"
            
            if ($removeOD -eq "y") {
                Remove-OneDriveStartup
                Uninstall-OneDrive
            }
            
            Configure-CloudProvider -Provider "Box" -SetAsDefault ($setDefault -eq "y")
            Write-Host "Box has been configured" -ForegroundColor Green
            Pause
            Show-Menu
        }
        "6" {
            $setDefault = Read-Host "Set Citrix ShareFile as default save location? (y/n)"
            $removeOD = Read-Host "Remove OneDrive? (y/n)"
            
            if ($removeOD -eq "y") {
                Remove-OneDriveStartup
                Uninstall-OneDrive
            }
            
            Configure-CloudProvider -Provider "ShareFile" -SetAsDefault ($setDefault -eq "y")
            Write-Host "ShareFile has been configured" -ForegroundColor Green
            Pause
            Show-Menu
        }
        "7" {
            $setDefault = Read-Host "Set Egnyte as default save location? (y/n)"
            $removeOD = Read-Host "Remove OneDrive? (y/n)"
            
            if ($removeOD -eq "y") {
                Remove-OneDriveStartup
                Uninstall-OneDrive
            }
            
            Configure-CloudProvider -Provider "Egnyte" -SetAsDefault ($setDefault -eq "y")
            Write-Host "Egnyte has been configured" -ForegroundColor Green
            Pause
            Show-Menu
        }
        "8" {
            $setDefault = Read-Host "Set OneDrive as default save location? (y/n)"
            Configure-CloudProvider -Provider "OneDrive" -SetAsDefault ($setDefault -eq "y")
            Write-Host "OneDrive has been configured" -ForegroundColor Green
            Pause
            Show-Menu
        }
        "9" {
            return
        }
        default {
            Write-Host "Invalid choice. Please try again." -ForegroundColor Red
            Pause
            Show-Menu
        }
    }
}

function Show-GUI {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Office 365 Cloud Storage Manager"
    $form.Size = New-Object System.Drawing.Size(500,400)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedSingle"
    $form.MaximizeBox = $false

    # Provider selection
    $labelProvider = New-Object System.Windows.Forms.Label
    $labelProvider.Text = "Select Cloud Provider:"
    $labelProvider.Location = New-Object System.Drawing.Point(20,20)
    $labelProvider.Size = New-Object System.Drawing.Size(150,20)
    $form.Controls.Add($labelProvider)

    $comboProvider = New-Object System.Windows.Forms.ComboBox
    $comboProvider.Location = New-Object System.Drawing.Point(180,18)
    $comboProvider.Size = New-Object System.Drawing.Size(200,20)
    $comboProvider.DropDownStyle = "DropDownList"
    $comboProvider.Items.AddRange(@("None", "OneDrive", "Dropbox", "GoogleDrive", "Box", "ShareFile", "Egnyte"))
    $comboProvider.SelectedIndex = 0
    $form.Controls.Add($comboProvider)

    # Options checkboxes
    $checkRemoveOD = New-Object System.Windows.Forms.CheckBox
    $checkRemoveOD.Text = "Remove OneDrive"
    $checkRemoveOD.Location = New-Object System.Drawing.Point(20,60)
    $checkRemoveOD.Size = New-Object System.Drawing.Size(150,20)
    $form.Controls.Add($checkRemoveOD)

    $checkSetDefault = New-Object System.Windows.Forms.CheckBox
    $checkSetDefault.Text = "Set as Default for Office"
    $checkSetDefault.Location = New-Object System.Drawing.Point(20,90)
    $checkSetDefault.Size = New-Object System.Drawing.Size(200,20)
    $form.Controls.Add($checkSetDefault)

    # Scan button
    $btnCheckStatus = New-Object System.Windows.Forms.Button
    $btnCheckStatus.Text = "Scan Existing Users"
    $btnCheckStatus.Location = New-Object System.Drawing.Point(20,120)
    $btnCheckStatus.Size = New-Object System.Drawing.Size(120,30)
    $form.Controls.Add($btnCheckStatus)

    # User information list
    $listInfo = New-Object System.Windows.Forms.ListBox
    $listInfo.Location = New-Object System.Drawing.Point(20,160)
    $listInfo.Size = New-Object System.Drawing.Size(450,150)
    $listInfo.Font = New-Object System.Drawing.Font("Consolas", 8)
    $listInfo.ScrollAlwaysVisible = $true
    $form.Controls.Add($listInfo)

    # Action buttons
    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = "Run Configuration"
    $btnRun.Location = New-Object System.Drawing.Point(270,320)
    $btnRun.Size = New-Object System.Drawing.Size(100,30)
    $form.Controls.Add($btnRun)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Close"
    $btnClose.Location = New-Object System.Drawing.Point(380,320)
    $btnClose.Size = New-Object System.Drawing.Size(80,30)
    $btnClose.Add_Click({ $form.Close() })
    $form.Controls.Add($btnClose)

    # Scan existing users button click
    $btnCheckStatus.Add_Click({
        $listInfo.Items.Clear()
        $listInfo.Items.Add("Scanning user profiles...")
        $form.Refresh()

        try {
            $userInfo = Get-CloudProviderForAllUsers
            $listInfo.Items.Clear()
            $listInfo.Items.Add("User Profile Scan Results:")
            $listInfo.Items.Add("=" * 50)

            if ($userInfo.Count -eq 0) {
                $listInfo.Items.Add("No user profiles found.")
            } else {
                foreach ($entry in $userInfo) {
                    $line = "{0,-15} | Provider: {1,-12} | SaveTo: {2}" -f $entry.User, $entry.CloudProvider, $entry.SaveLocation
                    $listInfo.Items.Add($line)
                }
            }

            $listInfo.Items.Add("=" * 50)
            $listInfo.Items.Add("Scan complete. Found $($userInfo.Count) user profiles.")
        } catch {
            $listInfo.Items.Clear()
            $listInfo.Items.Add("Error during scan: $($_.Exception.Message)")
        }
    })

    # Run configuration button click
    $btnRun.Add_Click({
        $provider = $comboProvider.SelectedItem
        $removeOD = $checkRemoveOD.Checked
        $setDefault = $checkSetDefault.Checked

        $listInfo.Items.Clear()
        $listInfo.Items.Add("Running configuration...")
        $listInfo.Items.Add("Provider: $provider")
        $listInfo.Items.Add("Remove OneDrive: $removeOD")
        $listInfo.Items.Add("Set as Default: $setDefault")
        $listInfo.Items.Add("=" * 40)
        $form.Refresh()

        try {
            if ($removeOD) {
                $listInfo.Items.Add("Removing OneDrive startup entries...")
                $form.Refresh()
                $removed = Remove-OneDriveStartup
                $listInfo.Items.Add("✔ Removed $removed OneDrive startup entries")
                
                $listInfo.Items.Add("Uninstalling OneDrive...")
                $form.Refresh()
                Uninstall-OneDrive
                $listInfo.Items.Add("✔ OneDrive uninstallation completed")
            }

            if ($provider -ne "None") {
                $listInfo.Items.Add("Configuring $provider as cloud provider...")
                $form.Refresh()
                Configure-CloudProvider -Provider $provider -SetAsDefault:$setDefault
                $listInfo.Items.Add("✔ Configured $provider as cloud provider")
                
                if ($setDefault) {
                    $listInfo.Items.Add("✔ Set $provider as default save location")
                }
            }

            $listInfo.Items.Add("=" * 40)
            $listInfo.Items.Add("✔ Configuration completed successfully!")
            $listInfo.Items.Add("Log file: $LogFile")
        } catch {
            $listInfo.Items.Add("❌ Error during configuration: $($_.Exception.Message)")
        }
    })

    # Auto-scan on startup
    $form.Add_Shown({
        $btnCheckStatus.PerformClick()
        $form.Activate()
    })

    [void]$form.ShowDialog()
}

# Main execution logic
Write-Log "Script started"

# Determine which configuration to use
if ($UseHardcodedConfig) {
    # Use hard-coded configuration (for Intune deployment)
    Write-Log "Using hard-coded configuration: Provider=$($Config.Provider), RemoveOneDrive=$($Config.RemoveOneDrive), SetAsDefault=$($Config.SetAsDefault)"
    
    if ($Config.RemoveOneDrive) {
        $removed = Remove-OneDriveStartup
        Write-Log "Removed $removed OneDrive startup entries"
        Uninstall-OneDrive
    }
    
    if ($Config.Provider -ne "None") {
        Configure-CloudProvider -Provider $Config.Provider -SetAsDefault $Config.SetAsDefault
    }
    
    if ($Config.ShowMenu) {
        Show-GUI
    }
} else {
    # Use command-line parameters or interactive GUI
    if ($Provider -eq "None" -and -not $RemoveOneDrive) {
        Show-GUI
    } else {
        # Process command line parameters
        Write-Log "Using command-line parameters: Provider=$Provider, RemoveOneDrive=$RemoveOneDrive, SetAsDefault=$SetAsDefault"
        
        if ($RemoveOneDrive) {
            $removed = Remove-OneDriveStartup
            Write-Log "Removed $removed OneDrive startup entries"
            Uninstall-OneDrive
        }
        
        if ($Provider -ne "None") {
            Configure-CloudProvider -Provider $Provider -SetAsDefault $SetAsDefault
        }
    }
}

Write-Log "Script completed successfully"
Write-Host "Cloud storage configuration completed. See log: $LogFile" -ForegroundColor Green
Read-Host -Prompt "Press Enter to exit"
