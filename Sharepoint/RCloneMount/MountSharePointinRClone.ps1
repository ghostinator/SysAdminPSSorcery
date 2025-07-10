#Requires -Version 5.1

<#
.SYNOPSIS
    Mount SharePoint Online and OneDrive folders using RClone (Command Line Version)
.DESCRIPTION
    This script provides command-line functionality for mounting SharePoint Online and OneDrive folders
    to local drive letters using RClone. For a GUI version, use SharePointOneDriveGUI.ps1
.PARAMETER RemoteName
    Name for the RClone remote configuration
.PARAMETER ServiceType
    Type of service: sharepoint, onedrive, or onedrive-personal
.PARAMETER SiteUrl
    SharePoint site URL (required for SharePoint)
.PARAMETER DriveLetter
    Drive letter to mount to (e.g., 'Z')
.PARAMETER RemotePath
    Optional remote path within the SharePoint/OneDrive
.PARAMETER TenantId
    Optional tenant ID for authentication
.EXAMPLE
    .\MountSharePointinRClone.ps1 -RemoteName "MySharePoint" -ServiceType "sharepoint" -SiteUrl "https://contoso.sharepoint.com/sites/mysite" -DriveLetter "Z"
.EXAMPLE
    .\MountSharePointinRClone.ps1 -RemoteName "MyOneDrive" -ServiceType "onedrive" -DriveLetter "Y"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$RemoteName,
    
    [Parameter(Mandatory=$true)]
    [ValidateSet("sharepoint", "onedrive", "onedrive-personal")]
    [string]$ServiceType,
    
    [Parameter(Mandatory=$false)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$true)]
    [ValidatePattern("^[D-Z]$")]
    [string]$DriveLetter,
    
    [Parameter(Mandatory=$false)]
    [string]$RemotePath = "",
    
    [Parameter(Mandatory=$false)]
    [string]$TenantId = "",
    
    [Parameter(Mandatory=$false)]
    [switch]$Unmount,
    
    [Parameter(Mandatory=$false)]
    [switch]$Configure
)

# Check if RClone is installed
function Test-RCloneInstalled {
    try {
        $null = Get-Command rclone -ErrorAction Stop
        return $true
    }
    catch {
        Write-Error "RClone is not installed. Please install RClone from https://rclone.org/downloads/"
        return $false
    }
}

# Configure RClone remote
function Configure-RCloneRemote {
    param(
        [string]$RemoteName,
        [string]$ServiceType,
        [string]$SiteUrl,
        [string]$TenantId
    )
    
    Write-Host "Configuring RClone remote '$RemoteName' for $ServiceType..." -ForegroundColor Yellow
    
    if ($ServiceType -eq "sharepoint" -and [string]::IsNullOrEmpty($SiteUrl)) {
        Write-Error "SharePoint site URL is required for SharePoint service type."
        return $false
    }
    
    try {
        # Start interactive RClone configuration
        Write-Host "Starting RClone configuration. Follow the prompts to authenticate..." -ForegroundColor Green
        Write-Host "When prompted:" -ForegroundColor Cyan
        Write-Host "1. Choose '$ServiceType' as the storage type" -ForegroundColor Cyan
        if ($ServiceType -eq "sharepoint") {
            Write-Host "2. Enter the SharePoint URL: $SiteUrl" -ForegroundColor Cyan
        }
        Write-Host "3. Follow the authentication prompts" -ForegroundColor Cyan
        
        $configArgs = @("config", "create", $RemoteName, $ServiceType)
        
        if ($ServiceType -eq "sharepoint" -and $SiteUrl) {
            $configArgs += "--sharepoint-url", $SiteUrl
        }
        
        if ($TenantId) {
            if ($ServiceType -eq "sharepoint") {
                $configArgs += "--sharepoint-tenant-id", $TenantId
            }
            elseif ($ServiceType -eq "onedrive") {
                $configArgs += "--onedrive-tenant-id", $TenantId
            }
        }
        
        $process = Start-Process -FilePath "rclone" -ArgumentList $configArgs -Wait -PassThru
        
        if ($process.ExitCode -eq 0) {
            Write-Host "Remote '$RemoteName' configured successfully!" -ForegroundColor Green
            return $true
        }
        else {
            Write-Error "Failed to configure remote '$RemoteName'"
            return $false
        }
    }
    catch {
        Write-Error "Configuration failed: $($_.Exception.Message)"
        return $false
    }
}

# Mount the remote
function Mount-Remote {
    param(
        [string]$RemoteName,
        [string]$DriveLetter,
        [string]$RemotePath
    )
    
    $mountPoint = "${DriveLetter}:"
    $remotePath = if ($RemotePath) { "${RemoteName}:${RemotePath}" } else { "${RemoteName}:" }
    
    # Check if drive letter is already in use
    if (Test-Path $mountPoint) {
        Write-Error "Drive letter $DriveLetter is already in use."
        return $false
    }
    
    Write-Host "Mounting $remotePath to $mountPoint..." -ForegroundColor Yellow
    
    try {
        $mountArgs = @(
            "mount",
            $remotePath,
            $mountPoint,
            "--vfs-cache-mode", "writes",
            "--vfs-cache-max-age", "1h",
            "--vfs-cache-max-size", "1G",
            "--buffer-size", "32M",
            "--timeout", "1h",
            "--retries", "3",
            "--low-level-retries", "10",
            "--stats", "0",
            "--log-level", "INFO"
        )
        
        Write-Host "Starting mount process..." -ForegroundColor Green
        $process = Start-Process -FilePath "rclone" -ArgumentList $mountArgs -PassThru
        
        # Wait a moment for the mount to establish
        Start-Sleep -Seconds 5
        
        # Check if mount was successful
        if (Test-Path $mountPoint) {
            Write-Host "Successfully mounted $remotePath to $mountPoint" -ForegroundColor Green
            Write-Host "Mount process ID: $($process.Id)" -ForegroundColor Cyan
            Write-Host "To unmount, run: .\MountSharePointinRClone.ps1 -RemoteName '$RemoteName' -DriveLetter '$DriveLetter' -Unmount" -ForegroundColor Cyan
            return $true
        }
        else {
            Write-Error "Mount failed - drive $mountPoint is not accessible"
            return $false
        }
    }
    catch {
        Write-Error "Mount failed: $($_.Exception.Message)"
        return $false
    }
}

# Unmount the drive
function Unmount-Drive {
    param([string]$DriveLetter)
    
    $mountPoint = "${DriveLetter}:"
    
    if (-not (Test-Path $mountPoint)) {
        Write-Warning "Drive $mountPoint is not currently mounted."
        return $true
    }
    
    Write-Host "Unmounting drive $mountPoint..." -ForegroundColor Yellow
    
    try {
        # Try graceful unmount first
        $unmountArgs = @("umount", $mountPoint)
        $process = Start-Process -FilePath "rclone" -ArgumentList $unmountArgs -Wait -PassThru -WindowStyle Hidden
        
        # Wait a moment and check
        Start-Sleep -Seconds 3
        
        if (-not (Test-Path $mountPoint)) {
            Write-Host "Successfully unmounted drive $mountPoint" -ForegroundColor Green
            return $true
        }
        else {
            Write-Warning "Graceful unmount failed, trying force unmount..."
            
            # Kill any RClone processes that might be holding the mount
            Get-Process -Name "rclone" -ErrorAction SilentlyContinue | Where-Object {
                $_.CommandLine -like "*$mountPoint*"
            } | Stop-Process -Force
            
            Start-Sleep -Seconds 2
            
            if (-not (Test-Path $mountPoint)) {
                Write-Host "Successfully force unmounted drive $mountPoint" -ForegroundColor Green
                return $true
            }
            else {
                Write-Error "Failed to unmount drive $mountPoint"
                return $false
            }
        }
    }
    catch {
        Write-Error "Unmount failed: $($_.Exception.Message)"
        return $false
    }
}

# Main execution
Write-Host "SharePoint/OneDrive RClone Mount Tool" -ForegroundColor Cyan
Write-Host "====================================" -ForegroundColor Cyan

# Check if RClone is installed
if (-not (Test-RCloneInstalled)) {
    exit 1
}

# Handle unmount operation
if ($Unmount) {
    if (Unmount-Drive -DriveLetter $DriveLetter) {
        Write-Host "Unmount operation completed successfully." -ForegroundColor Green
    }
    else {
        Write-Error "Unmount operation failed."
        exit 1
    }
    exit 0
}

# Handle configuration
if ($Configure) {
    if (Configure-RCloneRemote -RemoteName $RemoteName -ServiceType $ServiceType -SiteUrl $SiteUrl -TenantId $TenantId) {
        Write-Host "Configuration completed successfully." -ForegroundColor Green
    }
    else {
        Write-Error "Configuration failed."
        exit 1
    }
    exit 0
}

# Validate SharePoint requirements
if ($ServiceType -eq "sharepoint" -and [string]::IsNullOrEmpty($SiteUrl)) {
    Write-Error "SharePoint site URL is required when using SharePoint service type."
    Write-Host "Example: -SiteUrl 'https://yourtenant.sharepoint.com/sites/yoursite'" -ForegroundColor Yellow
    exit 1
}

# Check if remote exists
try {
    $remoteList = & rclone listremotes
    if ($remoteList -notcontains "${RemoteName}:") {
        Write-Warning "Remote '$RemoteName' is not configured."
        $configChoice = Read-Host "Would you like to configure it now? (y/n)"
        if ($configChoice -eq 'y' -or $configChoice -eq 'Y') {
            if (-not (Configure-RCloneRemote -RemoteName $RemoteName -ServiceType $ServiceType -SiteUrl $SiteUrl -TenantId $TenantId)) {
                exit 1
            }
        }
        else {
            Write-Host "Please configure the remote first using: .\MountSharePointinRClone.ps1 -RemoteName '$RemoteName' -ServiceType '$ServiceType' -SiteUrl '$SiteUrl' -Configure" -ForegroundColor Yellow
            exit 1
        }
    }
}
catch {
    Write-Error "Failed to check RClone remotes: $($_.Exception.Message)"
    exit 1
}

# Perform the mount
if (Mount-Remote -RemoteName $RemoteName -DriveLetter $DriveLetter -RemotePath $RemotePath) {
    Write-Host "Mount operation completed successfully." -ForegroundColor Green
    Write-Host "You can now access your SharePoint/OneDrive content at drive ${DriveLetter}:" -ForegroundColor Green
}
else {
    Write-Error "Mount operation failed."
    exit 1
}