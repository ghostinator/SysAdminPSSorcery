#Requires -Version 5.1
<#
.SYNOPSIS
    Configures default save locations for Windows Documents and Microsoft Office applications to Dropbox.
.DESCRIPTION
    This script sets the default save locations for the user's Documents folder and Office applications (Word, Excel, PowerPoint) to a specified Dropbox path.
    It also configures Office to prefer saving to the local computer by default instead of cloud locations.
.NOTES
    Designed for deployment via Microsoft Intune.
    Runs in the user context.
    Version 2.0: Includes more robust Documents folder redirection and sets Office to prefer local saves.
#>

[CmdletBinding()]
param()

# Configuration
$ErrorActionPreference = "Continue"
$VerbosePreference = "Continue"
$ProgressPreference = "SilentlyContinue"

# Initialize logging
$LogDate = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$LogPath = "C:\temp"
$LogFile = "$LogPath\DropboxSaveLocation_$LogDate.log"

# Ensure log directory exists
if (-not (Test-Path $LogPath)) {
    New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
}

# Enhanced logging function
function Write-LogEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Info", "Warning", "Error", "Success")]
        [string]$Level = "Info"
    )
    
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$TimeStamp] [$Level] $Message"
    
    # Write to log file
    $LogEntry | Out-File -FilePath $LogFile -Append -Encoding UTF8
    
    # Write to console based on level
    switch ($Level) {
        "Error" { Write-Error $Message }
        "Warning" { Write-Warning $Message }
        "Success" { Write-Host $Message -ForegroundColor Green }
        default { Write-Verbose $Message }
    }
}

# Registry paths configuration
$RegistryPaths = @{
    ShellFolders      = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    UserShellFolders  = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" # Added for robustness
    OfficeGeneral     = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\General" # Added for cloud save preference
    WordOptions       = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Word\Options"
    ExcelOptions      = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Excel\Options"
    PowerPointOptions = "HKCU:\SOFTWARE\Microsoft\Office\16.0\PowerPoint\Options"
}

# Office save path properties
$SavePathProperties = @{
    Word       = "DOC-PATH"
    Excel      = "DefaultPath"
    PowerPoint = "DefaultPath" # Note: This key is less reliable for setting the default save path in PowerPoint.
}

# Function to safely get registry value
function Get-RegistryValue {
    param(
        [string]$Path,
        [string]$Name
    )
    try {
        if (Test-Path $Path) {
            return Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue | Select-Object -ExpandProperty $Name
        }
        return $null
    }
    catch {
        Write-LogEntry "Failed to read registry value $Name from $Path`: $($_.Exception.Message)" -Level "Warning"
        return $null
    }
}

# Function to set registry value (now supports different property types)
function Set-RegistryValue {
    param(
        [string]$Path,
        [string]$Name,
        [ValidateNotNullOrEmpty()]
        $Value,
        [string]$Description,
        [string]$PropertyType = "String"
    )
    
    try {
        if (-not (Test-Path $Path)) {
            New-Item -Path $Path -Force | Out-Null
            Write-LogEntry "Created registry path: $Path" -Level "Info"
        }
        
        Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $PropertyType -Force | Out-Null
        Write-LogEntry "Successfully set $Description to '$Value'" -Level "Success"
        return $true
    }
    catch {
        Write-LogEntry "Failed to set $Description`: $($_.Exception.Message)" -Level "Error"
        return $false
    }
}

# --- Main Execution ---
try {
    Write-LogEntry "Starting Dropbox save location configuration" -Level "Info"
    Write-LogEntry "Running as user: $env:USERNAME" -Level "Info"
    
    $DropBoxPath = "$env:USERPROFILE\Dropbox"
    Write-LogEntry "Target Dropbox path: $DropBoxPath" -Level "Info"
    
    if (-not (Test-Path $DropBoxPath)) {
        Write-LogEntry "Dropbox folder not found at $DropBoxPath. Creating directory..." -Level "Warning"
        # Note: This assumes the user wants the folder here. In a managed environment, this is usually acceptable.
        New-Item -Path $DropBoxPath -ItemType Directory -Force | Out-Null
    }
    
    $UpdateCount = 0
    
    # 1. Set Office applications to prefer local saves over the cloud
    if (Set-RegistryValue -Path $RegistryPaths.OfficeGeneral -Name "PreferCloudSaveLocations" -Value 0 -Description "Office 'Save to Computer by default' setting" -PropertyType "DWord") {
        $UpdateCount++
    }
    
    # 2. Update Documents folder location (both keys for reliability)
    if ((Get-RegistryValue -Path $RegistryPaths.UserShellFolders -Name "Personal") -ne $DropBoxPath) {
        if (Set-RegistryValue -Path $RegistryPaths.UserShellFolders -Name "Personal" -Value $DropBoxPath -Description "User Shell Documents folder location") {
            $UpdateCount++
        }
        # Also update the legacy key for compatibility
        Set-RegistryValue -Path $RegistryPaths.ShellFolders -Name "Personal" -Value $DropBoxPath -Description "Legacy Shell Documents folder location" | Out-Null
    }
    else {
        Write-LogEntry "Documents folder already set to Dropbox" -Level "Info"
    }
    
    # 3. Update Word default save location
    if ((Get-RegistryValue -Path $RegistryPaths.WordOptions -Name $SavePathProperties.Word) -ne $DropBoxPath) {
        if (Set-RegistryValue -Path $RegistryPaths.WordOptions -Name $SavePathProperties.Word -Value $DropBoxPath -Description "Word default save location") {
            $UpdateCount++
        }
    }
    else {
        Write-LogEntry "Word save location already set to Dropbox" -Level "Info"
    }
    
    # 4. Update Excel default save location
    if ((Get-RegistryValue -Path $RegistryPaths.ExcelOptions -Name $SavePathProperties.Excel) -ne $DropBoxPath) {
        if (Set-RegistryValue -Path $RegistryPaths.ExcelOptions -Name $SavePathProperties.Excel -Value $DropBoxPath -Description "Excel default save location") {
            $UpdateCount++
        }
    }
    else {
        Write-LogEntry "Excel save location already set to Dropbox" -Level "Info"
    }
    
    # 5. Update PowerPoint default save location
    if ((Get-RegistryValue -Path $RegistryPaths.PowerPointOptions -Name $SavePathProperties.PowerPoint) -ne $DropBoxPath) {
        if (Set-RegistryValue -Path $RegistryPaths.PowerPointOptions -Name $SavePathProperties.PowerPoint -Value $DropBoxPath -Description "PowerPoint default save location") {
            $UpdateCount++
        }
    }
    else {
        Write-LogEntry "PowerPoint save location already set to Dropbox" -Level "Info"
    }
    
    # --- Summary ---
    Write-LogEntry "Configuration completed. Performed $UpdateCount update(s)." -Level "Success"
    Write-LogEntry "Users may need to restart applications or sign out/in for all changes to take full effect." -Level "Info"
    
    exit 0
    
} catch {
    Write-LogEntry "Script execution failed: $($_.Exception.Message)" -Level "Error"
    Write-LogEntry "Stack trace: $($_.ScriptStackTrace)" -Level "Error"
    
    exit 1
}