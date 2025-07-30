#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Office 365 Cloud Storage Provider Manager for Business - WPF Edition
.DESCRIPTION
    This script manages cloud storage providers for Office 365 applications in business environments.
    Features:
    - Remove OneDrive from startup locations for all users
    - Add OneDrive back to startup locations for all users
    - Configure enterprise cloud storage providers for Office 365
    - Set default save locations for Office applications
    - Modern WPF GUI interface with DataGrid
    - Scan and display current cloud provider configuration for all users
    - Export results to CSV or HTML
    - Optional password protection
    - Self-elevation with UAC
.PARAMETER Provider
    The cloud provider to configure: OneDrive, Dropbox, GoogleDrive, Box, ShareFile, Egnyte, or None
.PARAMETER RemoveOneDrive
    Switch to remove OneDrive completely
.PARAMETER SetAsDefault
    Switch to set the selected provider as default for Office applications
.PARAMETER NoGUI
    Switch to run without GUI (command line mode)
.EXAMPLE
    .\Office365CloudStorageManager.ps1 -Provider Dropbox -RemoveOneDrive -SetAsDefault -NoGUI
.NOTES
    Run as Administrator to access all user profiles and registry hives
    Created by: Brandon Cook
    Email: brandon@ghostinator.co
    GitHub: https://github.com/ghostinator/SysAdminPSSorcery
#>

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateSet("OneDrive", "Dropbox", "GoogleDrive", "Box", "ShareFile", "Egnyte", "None")]
    [string]$Provider = "None",

    [Parameter()]
    [switch]$RemoveOneDrive,

    [Parameter()]
    [switch]$SetAsDefault,

    [Parameter()]
    [switch]$NoGUI
)

#region CONFIGURATION SECTION - MODIFY FOR INTUNE DEPLOYMENT
# Set to $true to enable this configuration section (for Intune deployment)
# Set to $false to use command-line parameters or interactive GUI
$UseHardcodedConfig = $false

# Hard-coded configuration options (only used when $UseHardcodedConfig = $true)
$Config = @{
    # Set to one of: "OneDrive", "Dropbox", "GoogleDrive", "Box", "ShareFile", "Egnyte", "None"
    Provider = "Dropbox"

    # Set to $true to remove OneDrive completely, $false to keep it
    RemoveOneDrive = $true

    # Set to $true to set the selected provider as default for Office, $false otherwise
    SetAsDefault = $true

    # Set to $true to show interactive GUI, $false for silent operation
    ShowGUI = $false

    # Password protection (leave empty to disable)
    RequiredPassword = ""
}
#endregion

# Auto-elevate if not running as administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)) {
    
    Write-Host "Elevating to Administrator..." -ForegroundColor Yellow
    Start-Process powershell.exe -Verb runAs -ArgumentList "-NoExit", "-ExecutionPolicy Bypass", "-File `"$PSCommandPath`"", ($args -join " ")
    exit
}

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

function Test-Password {
    param([string]$RequiredPassword)
    
    if ([string]::IsNullOrEmpty($RequiredPassword)) {
        return $true
    }
    
    $attempts = 0
    do {
        $attempts++
        $password = Read-Host -AsSecureString "Enter admin tool password (Attempt $attempts/3)"
        $plainPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [Runtime.InteropServices.Marshal]::SecureStringToBSTR($password)
        )
        
        if ($plainPassword -eq $RequiredPassword) {
            return $true
        }
        
        if ($attempts -lt 3) {
            Write-Host "Incorrect password. Please try again." -ForegroundColor Red
        }
    } while ($attempts -lt 3)
    
    Write-Host "Too many failed attempts. Exiting..." -ForegroundColor Red
    return $false
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
                $oneDriveEnabled = "Unknown"

                if (Test-Path $officeCloudPath) {
                    $props = Get-ItemProperty -Path $officeCloudPath -ErrorAction SilentlyContinue
                    if ($props) {
                        $defaultCloud = if ($props.DefaultCloudProvider) { $props.DefaultCloudProvider } else { "None" }
                        $defaultSaveLoc = if ($props.DefaultSaveLocation) { $props.DefaultSaveLocation } else { "None" }
                        
                        # Check OneDrive status
                        $oneDriveEnabled = if ($props.EnableOneDriveInOffice -eq 1) { "Enabled" } else { "Disabled" }
                    }
                }

                # Check for OneDrive startup entry
                $runKeyPath = "Registry::$tempHive\Software\Microsoft\Windows\CurrentVersion\Run"
                $hasOneDriveStartup = $false
                if (Test-Path $runKeyPath) {
                    $runKey = Get-ItemProperty -Path $runKeyPath -ErrorAction SilentlyContinue
                    $hasOneDriveStartup = ($runKey -and $runKey.PSObject.Properties.Name -contains "OneDrive")
                }

                $results += [PSCustomObject]@{
                    User = $userName
                    CloudProvider = $defaultCloud
                    SaveLocation = $defaultSaveLoc
                    OneDriveOffice = $oneDriveEnabled
                    OneDriveStartup = if ($hasOneDriveStartup) { "Yes" } else { "No" }
                    Status = "OK"
                }

                reg unload $tempHive | Out-Null
            } else {
                $results += [PSCustomObject]@{
                    User = $userName
                    CloudProvider = "Error"
                    SaveLocation = "Error"
                    OneDriveOffice = "Error"
                    OneDriveStartup = "Error"
                    Status = "Failed to load registry"
                }
            }
        } catch {
            reg unload $tempHive 2>&1 | Out-Null
            $results += [PSCustomObject]@{
                User = $userName
                CloudProvider = "Error"
                SaveLocation = "Error"
                OneDriveOffice = "Error"
                OneDriveStartup = "Error"
                Status = $_.Exception.Message
            }
        }
    }

    return $results
}

function Remove-OneDriveStartup {
    Write-Log "Starting OneDrive startup removal for all existing users..."

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

        try {
            $ntUserDat = "$userPath\NTUSER.DAT"
            $tempHive = "HKU\TempHive_$userName"
            
            $loadResult = reg load $tempHive $ntUserDat 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Log "Loaded registry hive for $userName"
                
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
                
                reg unload $tempHive | Out-Null
                Write-Log "Unloaded registry hive for $userName"
            } else {
                Write-Log "Failed to load registry hive for $userName`: $loadResult"
            }
        } catch {
            Write-Log "Error processing registry for $userName`: $($_.Exception.Message)"
            reg unload $tempHive 2>&1 | Out-Null
        }
        
        # Remove from startup folders
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
    }

    Write-Log "OneDrive startup removal completed. Total entries removed: $totalRemoved"
    return $totalRemoved
}

function Enable-OneDriveStartup {
    Write-Log "Starting OneDrive startup enablement for all existing users..."

    $userProfiles = Get-ChildItem 'C:\Users' -Directory | Where-Object {
        (Test-Path "$($_.FullName)\NTUSER.DAT") -and
        ($_.Name -notin @('Default', 'Public', 'All Users', 'Default User'))
    }

    Write-Log "Found $($userProfiles.Count) user profiles to process"
    $totalAdded = 0

    foreach ($userProfile in $userProfiles) {
        $userName = $userProfile.Name
        $userPath = $userProfile.FullName
        Write-Log "Processing user: $userName"

        try {
            $ntUserDat = "$userPath\NTUSER.DAT"
            $tempHive = "HKU\TempHiveEnable_$userName"
            
            $loadResult = reg load $tempHive $ntUserDat 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Log "Loaded registry hive for $userName"
                
                $runKeyPath = "Registry::$tempHive\Software\Microsoft\Windows\CurrentVersion\Run"
                if (-not (Test-Path $runKeyPath)) {
                    New-Item -Path $runKeyPath -Force | Out-Null
                }

                # Find OneDrive executable
                $oneDriveExe = "$env:SystemRoot\System32\OneDrive.exe"
                if (-not (Test-Path $oneDriveExe)) {
                    $oneDriveExe = "$env:SystemRoot\SysWOW64\OneDrive.exe"
                }
                
                if (Test-Path $oneDriveExe) {
                    Set-ItemProperty -Path $runKeyPath -Name "OneDrive" -Value "`"$oneDriveExe`"" -Type String
                    Write-Log "Added OneDrive to Run registry key for $userName"
                    $totalAdded++
                } else {
                    Write-Log "OneDrive executable not found for $userName"
                }
                
                reg unload $tempHive | Out-Null
                Write-Log "Unloaded registry hive for $userName"
            } else {
                Write-Log "Failed to load registry hive for $userName`: $loadResult"
            }
        } catch {
            Write-Log "Error processing registry for $userName`: $($_.Exception.Message)"
            reg unload $tempHive 2>&1 | Out-Null
        }
        
        # Add to startup folders
        $userStartupFolder = "$userPath\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
        $userOneDriveLnk = Join-Path $userStartupFolder "OneDrive.lnk"
        
        try {
            if (-not (Test-Path $userStartupFolder)) {
                New-Item -ItemType Directory -Path $userStartupFolder -Force | Out-Null
            }
            
            # Find OneDrive executable
            $oneDriveExe = "$env:SystemRoot\System32\OneDrive.exe"
            if (-not (Test-Path $oneDriveExe)) {
                $oneDriveExe = "$env:SystemRoot\SysWOW64\OneDrive.exe"
            }
            
            if (Test-Path $oneDriveExe) {
                $shell = New-Object -ComObject WScript.Shell
                $shortcut = $shell.CreateShortcut($userOneDriveLnk)
                $shortcut.TargetPath = $oneDriveExe
                $shortcut.Save()
                Write-Log "Created OneDrive startup shortcut for $userName"
                $totalAdded++
            }
        } catch {
            Write-Log "Failed to create OneDrive shortcut for $userName`: $($_.Exception.Message)"
        }
    }

    # Add to All Users Startup Folder
    $allUsersStartup = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp"
    $allUsersOneDriveLnk = Join-Path $allUsersStartup "OneDrive.lnk"

    try {
        if (-not (Test-Path $allUsersStartup)) {
            New-Item -ItemType Directory -Path $allUsersStartup -Force | Out-Null
        }
        
        # Find OneDrive executable
        $oneDriveExe = "$env:SystemRoot\System32\OneDrive.exe"
        if (-not (Test-Path $oneDriveExe)) {
            $oneDriveExe = "$env:SystemRoot\SysWOW64\OneDrive.exe"
        }
        
        if (Test-Path $oneDriveExe) {
            $shell = New-Object -ComObject WScript.Shell
            $shortcut = $shell.CreateShortcut($allUsersOneDriveLnk)
            $shortcut.TargetPath = $oneDriveExe
            $shortcut.Save()
            Write-Log "Added OneDrive shortcut to All Users startup folder"
            $totalAdded++
        }
    } catch {
        Write-Log "Failed to add OneDrive shortcut to All Users startup folder: $($_.Exception.Message)"
    }

    Write-Log "OneDrive startup enablement completed. Total entries added: $totalAdded"
    return $totalAdded
}

function Uninstall-OneDrive {
    Write-Log "Starting OneDrive uninstallation process..."
    
    Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue | Stop-Process -Force
    Write-Log "Stopped OneDrive processes"
    
    $oneDriveSetup = "$env:SystemRoot\SysWOW64\OneDriveSetup.exe"
    if (-not (Test-Path $oneDriveSetup)) {
        $oneDriveSetup = "$env:SystemRoot\System32\OneDriveSetup.exe"
    }
    
    if (Test-Path $oneDriveSetup) {
        Write-Log "Running OneDrive uninstaller..."
        Start-Process $oneDriveSetup "/uninstall" -NoNewWindow -Wait
        Write-Log "OneDrive uninstaller completed"
    }
    
    $oneDriveFolder = "$env:USERPROFILE\OneDrive"
    if (Test-Path $oneDriveFolder) {
        try {
            Remove-Item $oneDriveFolder -Force -Recurse -ErrorAction Stop
            Write-Log "Removed OneDrive folder: $oneDriveFolder"
        } catch {
            Write-Log "Failed to remove OneDrive folder: $($_.Exception.Message)"
        }
    }
    
    # Disable OneDrive via registry
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

            $loadResult = reg load $tempHive $ntUserDat 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Log "Loaded registry hive for $userName"
                
                $officeCloudPath = "Registry::$tempHive\Software\Microsoft\Office\16.0\Common\Cloud"
                
                if (-not (Test-Path $officeCloudPath)) {
                    New-Item -Path $officeCloudPath -Force | Out-Null
                }
                
                # Disable all providers first
                $allProviders = @("OneDrive", "Dropbox", "GoogleDrive", "Box", "ShareFile", "Egnyte")
                foreach ($p in $allProviders) {
                    $enableKey = "Enable${p}InOffice"
                    if ($p -ne $Provider) {
                        Set-ItemProperty -Path $officeCloudPath -Name $enableKey -Value 0 -Type DWord -ErrorAction SilentlyContinue
                    }
                }
                
                # Configure selected provider
                switch ($Provider) {
                    "Dropbox" {
                        Set-ItemProperty -Path $officeCloudPath -Name "EnableDropboxInOffice" -Value 1 -Type DWord
                        if ($SetAsDefault) {
                            Set-ItemProperty -Path $officeCloudPath -Name "DefaultCloudProvider" -Value "Dropbox" -Type String
                            Set-ItemProperty -Path $officeCloudPath -Name "DefaultSaveLocation" -Value "Dropbox" -Type String
                        }
                        Write-Log "Enabled Dropbox integration for $userName"
                    }
                    "GoogleDrive" {
                        Set-ItemProperty -Path $officeCloudPath -Name "EnableGoogleDriveInOffice" -Value 1 -Type DWord
                        if ($SetAsDefault) {
                            Set-ItemProperty -Path $officeCloudPath -Name "DefaultCloudProvider" -Value "GoogleDrive" -Type String
                            Set-ItemProperty -Path $officeCloudPath -Name "DefaultSaveLocation" -Value "GoogleDrive" -Type String
                        }
                        Write-Log "Enabled Google Drive integration for $userName"
                    }
                    "Box" {
                        Set-ItemProperty -Path $officeCloudPath -Name "EnableBoxInOffice" -Value 1 -Type DWord
                        if ($SetAsDefault) {
                            Set-ItemProperty -Path $officeCloudPath -Name "DefaultCloudProvider" -Value "Box" -Type String
                            Set-ItemProperty -Path $officeCloudPath -Name "DefaultSaveLocation" -Value "Box" -Type String
                        }
                        Write-Log "Enabled Box integration for $userName"
                    }
                    "ShareFile" {
                        Set-ItemProperty -Path $officeCloudPath -Name "EnableShareFileInOffice" -Value 1 -Type DWord
                        if ($SetAsDefault) {
                            Set-ItemProperty -Path $officeCloudPath -Name "DefaultCloudProvider" -Value "ShareFile" -Type String
                            Set-ItemProperty -Path $officeCloudPath -Name "DefaultSaveLocation" -Value "ShareFile" -Type String
                        }
                        Write-Log "Enabled ShareFile integration for $userName"
                    }
                    "Egnyte" {
                        Set-ItemProperty -Path $officeCloudPath -Name "EnableEgnyteInOffice" -Value 1 -Type DWord
                        if ($SetAsDefault) {
                            Set-ItemProperty -Path $officeCloudPath -Name "DefaultCloudProvider" -Value "Egnyte" -Type String
                            Set-ItemProperty -Path $officeCloudPath -Name "DefaultSaveLocation" -Value "Egnyte" -Type String
                        }
                        Write-Log "Enabled Egnyte integration for $userName"
                    }
                    "OneDrive" {
                        Set-ItemProperty -Path $officeCloudPath -Name "EnableOneDriveInOffice" -Value 1 -Type DWord
                        if ($SetAsDefault) {
                            Set-ItemProperty -Path $officeCloudPath -Name "DefaultCloudProvider" -Value "OneDrive" -Type String
                            Set-ItemProperty -Path $officeCloudPath -Name "DefaultSaveLocation" -Value "OneDrive" -Type String
                        }
                        Write-Log "Enabled OneDrive integration for $userName"
                    }
                }
                
                $officeBackstagePath = "Registry::$tempHive\Software\Microsoft\Office\16.0\Common\Cloud\Backstage"
                if (-not (Test-Path $officeBackstagePath)) {
                    New-Item -Path $officeBackstagePath -Force | Out-Null
                }
                
                if ($Provider -ne "None") {
                    Set-ItemProperty -Path $officeBackstagePath -Name "Show$Provider" -Value 1 -Type DWord -ErrorAction SilentlyContinue
                }
                
                reg unload $tempHive | Out-Null
                Write-Log "Unloaded registry hive for $userName"
            } else {
                Write-Log "Failed to load registry hive for $userName`: $loadResult"
            }
        } catch {
            Write-Log "Error configuring cloud provider for $userName`: $($_.Exception.Message)"
            reg unload $tempHive 2>&1 | Out-Null
        }
    }
    
    Write-Log "Cloud provider configuration completed for all users"
}

function Show-WpfGui {
    # Load required assemblies
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

    # Define XAML
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Office 365 Cloud Storage Manager" Height="650" Width="950"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResize">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Title -->
        <TextBlock Grid.Row="0" Text="Office 365 Cloud Storage Provider Manager" 
                   FontSize="18" FontWeight="Bold" Margin="0,0,0,5" HorizontalAlignment="Center"/>
        
        <!-- Author Info -->
        <TextBlock Grid.Row="1" Text="by Brandon Cook | brandon@ghostinator.co | GitHub: ghostinator/SysAdminPSSorcery" 
                   FontSize="11" Foreground="DarkSlateGray" HorizontalAlignment="Center" Margin="0,0,0,15"/>

        <!-- Input Controls -->
        <StackPanel Orientation="Horizontal" Grid.Row="2" Margin="0,0,0,15">
            <Label Content="Cloud Provider:" Width="120" VerticalAlignment="Center"/>
            <ComboBox Name="ProviderSelector" Width="150" SelectedIndex="0">
                <ComboBoxItem Content="None"/>
                <ComboBoxItem Content="OneDrive"/>
                <ComboBoxItem Content="Dropbox"/>
                <ComboBoxItem Content="GoogleDrive"/>
                <ComboBoxItem Content="Box"/>
                <ComboBoxItem Content="ShareFile"/>
                <ComboBoxItem Content="Egnyte"/>
            </ComboBox>
            <CheckBox Name="RemoveOneDriveCheckbox" Margin="30,0,0,0" Content="Remove OneDrive" VerticalAlignment="Center"/>
            <CheckBox Name="SetDefaultCheckbox" Margin="20,0,0,0" Content="Set as Default" VerticalAlignment="Center"/>
        </StackPanel>

        <!-- Data Grid -->
        <GroupBox Grid.Row="3" Header="User Cloud Provider Status" Margin="0,0,0,15">
            <DataGrid Name="UserGrid" AutoGenerateColumns="False" IsReadOnly="True" CanUserAddRows="False" 
                      GridLinesVisibility="Horizontal" AlternatingRowBackground="LightGray">
                <DataGrid.Columns>
                    <DataGridTextColumn Header="User" Binding="{Binding User}" Width="120"/>
                    <DataGridTextColumn Header="Cloud Provider" Binding="{Binding CloudProvider}" Width="120"/>
                    <DataGridTextColumn Header="Save Location" Binding="{Binding SaveLocation}" Width="120"/>
                    <DataGridTextColumn Header="OneDrive Office" Binding="{Binding OneDriveOffice}" Width="100"/>
                    <DataGridTextColumn Header="OneDrive Startup" Binding="{Binding OneDriveStartup}" Width="100"/>
                    <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="*"/>
                </DataGrid.Columns>
            </DataGrid>
        </GroupBox>

        <!-- Status Bar -->
        <TextBlock Name="StatusText" Grid.Row="4" Text="Ready" Margin="0,0,0,10" 
                   FontStyle="Italic" Foreground="Blue"/>

        <!-- Action Buttons -->
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Grid.Row="5">
            <Button Name="ScanButton" Content="Scan Users" Width="100" Height="30" Margin="0,0,10,0"/>
            <Button Name="EnableOneDriveButton" Content="Add OneDrive to Startup" Width="160" Height="30" Margin="0,0,10,0"/>
            <Button Name="RunButton" Content="Run Configuration" Width="120" Height="30" Margin="0,0,10,0"/>
            <Button Name="ExportButton" Content="Export..." Width="80" Height="30" Margin="0,0,10,0"/>
            <Button Name="RefreshButton" Content="Refresh" Width="80" Height="30" Margin="0,0,10,0"/>
            <Button Name="CloseButton" Content="Close" Width="80" Height="30"/>
        </StackPanel>
    </Grid>
</Window>
"@

    # Load XAML
    $reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)

    # Get controls
    $providerSelector = $window.FindName("ProviderSelector")
    $removeOneDriveCheckbox = $window.FindName("RemoveOneDriveCheckbox")
    $setDefaultCheckbox = $window.FindName("SetDefaultCheckbox")
    $userGrid = $window.FindName("UserGrid")
    $statusText = $window.FindName("StatusText")
    $scanButton = $window.FindName("ScanButton")
    $enableOneDriveButton = $window.FindName("EnableOneDriveButton")
    $runButton = $window.FindName("RunButton")
    $exportButton = $window.FindName("ExportButton")
    $refreshButton = $window.FindName("RefreshButton")
    $closeButton = $window.FindName("CloseButton")

    # Global data variable
    $script:currentData = @()

    # Scan Users Button
    $scanButton.Add_Click({
        $statusText.Text = "Scanning user profiles..."
        $window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [action]{})
        
        try {
            $script:currentData = Get-CloudProviderForAllUsers
            $userGrid.ItemsSource = $script:currentData
            $statusText.Text = "Scan completed. Found $($script:currentData.Count) user profiles."
        } catch {
            $statusText.Text = "Error during scan: $($_.Exception.Message)"
        }
    })

    # Enable OneDrive Startup Button
    $enableOneDriveButton.Add_Click({
        $statusText.Text = "Adding OneDrive to startup for all users..."
        $window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [action]{})

        try {
            $added = Enable-OneDriveStartup
            $statusText.Text = "OneDrive added to startup for $added user entries."
            
            # Refresh the grid
            $script:currentData = Get-CloudProviderForAllUsers
            $userGrid.ItemsSource = $script:currentData
            
            [System.Windows.MessageBox]::Show("OneDrive has been added to startup for $added user entries.", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        } catch {
            $errorMessage = "Error adding OneDrive to startup: $($_.Exception.Message)"
            $statusText.Text = $errorMessage
            [System.Windows.MessageBox]::Show($errorMessage, "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    })

    # Run Configuration Button
    $runButton.Add_Click({
        $provider = $providerSelector.SelectedItem.Content
        $removeOD = $removeOneDriveCheckbox.IsChecked
        $setDefault = $setDefaultCheckbox.IsChecked

        $statusText.Text = "Running configuration..."
        $window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [action]{})

        try {
            $actions = @()
            
            if ($removeOD) {
                $statusText.Text = "Removing OneDrive startup entries..."
                $window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [action]{})
                $removed = Remove-OneDriveStartup
                $actions += "Removed $removed OneDrive startup entries"
                
                $statusText.Text = "Uninstalling OneDrive..."
                $window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [action]{})
                Uninstall-OneDrive
                $actions += "OneDrive uninstallation completed"
            }

            if ($provider -ne "None") {
                $statusText.Text = "Configuring $provider as cloud provider..."
                $window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [action]{})
                Configure-CloudProvider -Provider $provider -SetAsDefault:$setDefault
                $actions += "Configured $provider as cloud provider"
                
                if ($setDefault) {
                    $actions += "Set $provider as default save location"
                }
            }

            # Refresh the grid
            $statusText.Text = "Refreshing user data..."
            $window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [action]{})
            $script:currentData = Get-CloudProviderForAllUsers
            $userGrid.ItemsSource = $script:currentData

            $message = "Configuration completed successfully!`n`n" + ($actions -join "`n") + "`n`nLog file: $LogFile"
            [System.Windows.MessageBox]::Show($message, "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            
            $statusText.Text = "Configuration completed successfully."
        } catch {
            $errorMessage = "Error during configuration: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show($errorMessage, "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            $statusText.Text = $errorMessage
        }
    })

    # Export Button
    $exportButton.Add_Click({
        if ($script:currentData.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No data to export. Please scan users first.", "No Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }

        $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveDialog.Filter = "CSV File (*.csv)|*.csv|HTML Report (*.html)|*.html"
        $saveDialog.FileName = "CloudProviderReport_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss')"
        
        if ($saveDialog.ShowDialog()) {
            try {
                if ($saveDialog.FileName.EndsWith(".csv")) {
                    $script:currentData | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8
                    $statusText.Text = "Data exported to CSV: $($saveDialog.FileName)"
                } elseif ($saveDialog.FileName.EndsWith(".html")) {
                    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Office 365 Cloud Provider Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #0078d4; }
        .header { background-color: #f8f9fa; padding: 15px; border-radius: 5px; margin-bottom: 20px; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; font-weight: bold; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        .timestamp { color: #666; font-size: 0.9em; }
        .footer { margin-top: 30px; padding-top: 15px; border-top: 1px solid #ddd; color: #666; font-size: 0.9em; }
    </style>
</head>
<body>
    <div class="header">
        <h1>Office 365 Cloud Provider Report</h1>
        <p class="timestamp">Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        <p>Total Users: $($script:currentData.Count)</p>
    </div>
    
    <table>
        <tr>
            <th>User</th>
            <th>Cloud Provider</th>
            <th>Save Location</th>
            <th>OneDrive Office</th>
            <th>OneDrive Startup</th>
            <th>Status</th>
        </tr>
"@
                    foreach ($row in $script:currentData) {
                        $html += "<tr><td>$($row.User)</td><td>$($row.CloudProvider)</td><td>$($row.SaveLocation)</td><td>$($row.OneDriveOffice)</td><td>$($row.OneDriveStartup)</td><td>$($row.Status)</td></tr>"
                    }
                    
                    $html += @"
    </table>
    
    <div class="footer">
        <p>Report generated by Office 365 Cloud Storage Manager</p>
        <p>Created by: Brandon Cook | brandon@ghostinator.co</p>
        <p>GitHub: <a href="https://github.com/ghostinator/SysAdminPSSorcery">ghostinator/SysAdminPSSorcery</a></p>
    </div>
</body>
</html>
"@
                    $html | Set-Content -Path $saveDialog.FileName -Encoding UTF8
                    $statusText.Text = "Report exported to HTML: $($saveDialog.FileName)"
                }
                
                [System.Windows.MessageBox]::Show("Export completed successfully!", "Export Complete", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            } catch {
                [System.Windows.MessageBox]::Show("Export failed: $($_.Exception.Message)", "Export Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        }
    })

    # Refresh Button
    $refreshButton.Add_Click({
        $scanButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
    })

    # Close Button
    $closeButton.Add_Click({
        $window.Close()
    })

    # Auto-scan on startup
    $window.Add_Loaded({
        $scanButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
    })

    # Show window
    $window.ShowDialog() | Out-Null
}

function Show-Menu {
    Clear-Host
    Write-Host "===== Office 365 Cloud Storage Provider Manager =====" -ForegroundColor Cyan
    Write-Host "Created by: Brandon Cook | brandon@ghostinator.co" -ForegroundColor Green
    Write-Host "GitHub: https://github.com/ghostinator/SysAdminPSSorcery" -ForegroundColor Green
    Write-Host ""
    Write-Host "1: Remove OneDrive from startup"
    Write-Host "2: Add OneDrive to startup"
    Write-Host "3: Uninstall OneDrive completely"
    Write-Host "4: Configure Dropbox as cloud provider"
    Write-Host "5: Configure Google Drive as cloud provider"
    Write-Host "6: Configure Box as cloud provider"
    Write-Host "7: Configure Citrix ShareFile as cloud provider"
    Write-Host "8: Configure Egnyte as cloud provider"
    Write-Host "9: Configure OneDrive as cloud provider"
    Write-Host "G: Show GUI"
    Write-Host "0: Exit"
    Write-Host "===================================================" -ForegroundColor Cyan
    
    $choice = Read-Host "Enter your choice (0-9, G)"
    
    switch ($choice.ToUpper()) {
        "1" {
            $removed = Remove-OneDriveStartup
            Write-Host "Removed $removed OneDrive startup entries" -ForegroundColor Green
            Pause
            Show-Menu
        }
        "2" {
            $added = Enable-OneDriveStartup
            Write-Host "Added OneDrive to startup for $added user entries" -ForegroundColor Green
            Pause
            Show-Menu
        }
        "3" {
            Remove-OneDriveStartup
            Uninstall-OneDrive
            Write-Host "OneDrive has been uninstalled and disabled" -ForegroundColor Green
            Pause
            Show-Menu
        }
        "4" {
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
        "5" {
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
        "6" {
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
        "7" {
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
        "8" {
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
        "9" {
            $setDefault = Read-Host "Set OneDrive as default save location? (y/n)"
            Configure-CloudProvider -Provider "OneDrive" -SetAsDefault ($setDefault -eq "y")
            Write-Host "OneDrive has been configured" -ForegroundColor Green
            Pause
            Show-Menu
        }
        "G" {
            Show-WpfGui
            Show-Menu
        }
        "0" {
            return
        }
        default {
            Write-Host "Invalid choice. Please try again." -ForegroundColor Red
            Pause
            Show-Menu
        }
    }
}

# Main execution logic
Write-Log "Script started (PID: $PID)"
Write-Log "Created by: Brandon Cook | brandon@ghostinator.co"
Write-Log "GitHub: https://github.com/ghostinator/SysAdminPSSorcery"

# Password protection check
if (-not (Test-Password -RequiredPassword $Config.RequiredPassword)) {
    exit 1
}

# Determine execution mode
if ($UseHardcodedConfig) {
    Write-Log "Using hard-coded configuration: Provider=$($Config.Provider), RemoveOneDrive=$($Config.RemoveOneDrive), SetAsDefault=$($Config.SetAsDefault)"
    
    if ($Config.RemoveOneDrive) {
        $removed = Remove-OneDriveStartup
        Write-Log "Removed $removed OneDrive startup entries"
        Uninstall-OneDrive
    }
    
    if ($Config.Provider -ne "None") {
        Configure-CloudProvider -Provider $Config.Provider -SetAsDefault $Config.SetAsDefault
    }
    
    if ($Config.ShowGUI) {
        Show-WpfGui
    }
} else {
    # Command-line mode or interactive
    if ($NoGUI -or ($Provider -ne "None" -or $RemoveOneDrive)) {
        # Command line parameters provided
        Write-Log "Using command-line parameters: Provider=$Provider, RemoveOneDrive=$RemoveOneDrive, SetAsDefault=$SetAsDefault"
        
        if ($RemoveOneDrive) {
            $removed = Remove-OneDriveStartup
            Write-Log "Removed $removed OneDrive startup entries"
            Uninstall-OneDrive
        }
        
        if ($Provider -ne "None") {
            Configure-CloudProvider -Provider $Provider -SetAsDefault $SetAsDefault
        }
    } else {
        # Show GUI by default
        Show-WpfGui
    }
}

Write-Log "Script completed successfully"
Write-Host "Cloud storage configuration completed. See log: $LogFile" -ForegroundColor Green

if ($NoGUI -or ($Provider -ne "None" -or $RemoveOneDrive)) {
    Read-Host -Prompt "Press Enter to exit"
}
