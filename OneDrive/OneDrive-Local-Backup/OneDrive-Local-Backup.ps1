<#
.SYNOPSIS
    Backs up only the files from a OneDrive folder that are physically stored on the local disk.

.DESCRIPTION
    This script is designed to perform a safe backup of a OneDrive folder by copying only the files
    that are fully available offline. It inspects the file attributes to specifically ignore any
    cloud-only placeholders ("Files On-Demand"), thus preventing the OneDrive client from
    triggering unwanted downloads. This is especially useful when the OneDrive sync client is
    unstable or broken, and a reliable local backup is needed before a system reset or migration.

.NOTES
    Author: ghostinator
    License: MIT
    Version: 1.0
#>

# --- CONFIGURE YOUR PATHS HERE ---
# Example: "C:\Users\JohnDoe\OneDrive - Contoso"
$sourcePath = "C:\Users\YourUsername\OneDrive"

# Example: "E:\Backups\OneDrive"
$destinationPath = "D:\OneDrive_Backup"
# -----------------------------------


# --- SCRIPT LOGIC ---

# Ensure the root backup folder exists
if (-not (Test-Path $destinationPath)) {
    try {
        New-Item -ItemType Directory -Path $destinationPath -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Error "FATAL: Could not create destination directory '$destinationPath'. Please check permissions and path."
        exit 1
    }
}

Write-Host "🚀 Starting backup process..." -ForegroundColor Cyan
Write-Host "Source:      $sourcePath"
Write-Host "Destination: $destinationPath"
Write-Host "This will ONLY copy files that are physically on the disk and ignore cloud placeholders."
Write-Host "---------------------------------------------------------------------------------------"

# Get all items, including hidden/system ones. Suppress errors for paths that are too long or inaccessible.
$allItems = Get-ChildItem -Path $sourcePath -Recurse -Force -ErrorAction SilentlyContinue

# Filter for ONLY items that are physically on the disk (do not have the 'Offline' attribute)
$localItems = $allItems | Where-Object { ($_.Attributes -band [System.IO.FileAttributes]::Offline) -eq 0 }

# Process the filtered local items
foreach ($item in $localItems) {
    # Recreate the target folder and file structure in the destination
    $targetPath = $item.FullName.Replace($sourcePath, $destinationPath)

    if ($item.PSIsContainer) {
        # If the item is a directory, ensure it exists in the destination
        if (-not (Test-Path $targetPath)) {
             New-Item -ItemType Directory -Path $targetPath -Force | Out-Null
        }
    } else {
        # If the item is a file, copy it
        try {
            Write-Host "Copying: $($item.Name)" -ForegroundColor Green
            Copy-Item -Path $item.FullName -Destination $targetPath -Force -ErrorAction Stop
        }
        catch {
             Write-Warning "Could not copy file: $($item.FullName). It might be locked or inaccessible."
        }
    }
}

Write-Host "---------------------------------------------------------------------------------------"
Write-Host "✅ Backup process completed." -ForegroundColor Cyan
