 #Requires -Version 5.1
<#
.SYNOPSIS
    Configures default save locations for Windows Documents and Microsoft Office applications to Dropbox
.DESCRIPTION
    This script sets the default save locations for Documents folder and Office applications (Word, Excel, PowerPoint) to the user's Dropbox folder
.NOTES
    Designed for deployment via Microsoft Intune
    Runs in user context
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
    ShellFolders = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    WordOptions  = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Word\Options"
    ExcelOptions = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Excel\Options"
    PowerPointOptions = "HKCU:\SOFTWARE\Microsoft\Office\16.0\PowerPoint\Options"
}

# Office save path properties
$SavePathProperties = @{
    Word = "DOC-PATH"
    Excel = "DefaultPath"
    PowerPoint = "DefaultPath"
}

# Function to safely get registry value
function Get-RegistryValue {
    param(
        [string]$Path,
        [string]$Name
    )
    
    try {
        if (Test-Path $Path) {
            $value = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
            return $value.$Name
        }
        return $null
    }
    catch {
        Write-LogEntry "Failed to read registry value $Name from $Path`: $($_.Exception.Message)" -Level "Warning"
        return $null
    }
}

# Function to set registry value
function Set-RegistryValue {
    param(
        [string]$Path,
        [string]$Name,
        [string]$Value,
        [string]$Description
    )
    
    try {
        # Ensure registry path exists
        if (-not (Test-Path $Path)) {
            New-Item -Path $Path -Force | Out-Null
            Write-LogEntry "Created registry path: $Path" -Level "Info"
        }
        
        # Check if property exists
        $existingValue = Get-RegistryValue -Path $Path -Name $Name
        
        if ($null -eq $existingValue) {
            New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType String -Force | Out-Null
            Write-LogEntry "Created $Description`: $Value" -Level "Success"
        }
        else {
            Set-ItemProperty -Path $Path -Name $Name -Value $Value | Out-Null
            Write-LogEntry "Updated $Description`: $Value" -Level "Success"
        }
        
        return $true
    }
    catch {
        Write-LogEntry "Failed to set $Description`: $($_.Exception.Message)" -Level "Error"
        return $false
    }
}

# Main execution
try {
    Write-LogEntry "Starting Dropbox save location configuration" -Level "Info"
    Write-LogEntry "Running as user: $env:USERNAME" -Level "Info"
    
    # Define target path
    $DropBoxPath = "$env:USERPROFILE\Dropbox"
    Write-LogEntry "Target Dropbox path: $DropBoxPath" -Level "Info"
    
    # Verify Dropbox folder exists
    if (-not (Test-Path $DropBoxPath)) {
        Write-LogEntry "Dropbox folder not found at $DropBoxPath. Creating directory..." -Level "Warning"
        New-Item -Path $DropBoxPath -ItemType Directory -Force | Out-Null
    }
    
    # Get current values
    Write-LogEntry "Reading current registry values..." -Level "Info"
    $CurrentValues = @{
        Documents = Get-RegistryValue -Path $RegistryPaths.ShellFolders -Name "Personal"
        Word = Get-RegistryValue -Path $RegistryPaths.WordOptions -Name $SavePathProperties.Word
        Excel = Get-RegistryValue -Path $RegistryPaths.ExcelOptions -Name $SavePathProperties.Excel
        PowerPoint = Get-RegistryValue -Path $RegistryPaths.PowerPointOptions -Name $SavePathProperties.PowerPoint
    }
    
    # Log current values
    foreach ($key in $CurrentValues.Keys) {
        $value = if ($CurrentValues[$key]) { $CurrentValues[$key] } else { "Not set" }
        Write-LogEntry "Current $key path: $value" -Level "Info"
    }
    
    $UpdateCount = 0
    
    # Update Documents folder location
    if ($CurrentValues.Documents -ne $DropBoxPath) {
        if (Set-RegistryValue -Path $RegistryPaths.ShellFolders -Name "Personal" -Value $DropBoxPath -Description "Documents folder location") {
            $UpdateCount++
        }
    } else {
        Write-LogEntry "Documents folder already set to Dropbox" -Level "Info"
    }
    
    # Update Word save location
    if ($CurrentValues.Word -ne $DropBoxPath) {
        if (Set-RegistryValue -Path $RegistryPaths.WordOptions -Name $SavePathProperties.Word -Value $DropBoxPath -Description "Word default save location") {
            $UpdateCount++
        }
    } else {
        Write-LogEntry "Word save location already set to Dropbox" -Level "Info"
    }
    
    # Update Excel save location
    if ($CurrentValues.Excel -ne $DropBoxPath) {
        if (Set-RegistryValue -Path $RegistryPaths.ExcelOptions -Name $SavePathProperties.Excel -Value $DropBoxPath -Description "Excel default save location") {
            $UpdateCount++
        }
    } else {
        Write-LogEntry "Excel save location already set to Dropbox" -Level "Info"
    }
    
    # Update PowerPoint save location
    if ($CurrentValues.PowerPoint -ne $DropBoxPath) {
        if (Set-RegistryValue -Path $RegistryPaths.PowerPointOptions -Name $SavePathProperties.PowerPoint -Value $DropBoxPath -Description "PowerPoint default save location") {
            $UpdateCount++
        }
    } else {
        Write-LogEntry "PowerPoint save location already set to Dropbox" -Level "Info"
    }
    
    # Summary
    Write-LogEntry "Configuration completed successfully. Updated $UpdateCount settings." -Level "Success"
    Write-LogEntry "Users may need to restart Office applications for changes to take effect." -Level "Info"
    
    # Exit with success code for Intune
    exit 0
    
} catch {
    Write-LogEntry "Script execution failed: $($_.Exception.Message)" -Level "Error"
    Write-LogEntry "Stack trace: $($_.ScriptStackTrace)" -Level "Error"
    
    # Exit with error code for Intune
    exit 1
}
 
