# OneDrive Local File Backup Utility

A robust PowerShell script to safely back up only the local, on-disk files from your OneDrive folder. It intelligently ignores all cloud-only files, preventing unwanted downloads and errors. This script is a lifesaver when the OneDrive sync client is broken, unstable, or untrustworthy, and you need a clean backup before resetting your PC.

-----

## The Problem

Have you ever been in this situation?

  - You need to back up your OneDrive files before reinstalling Windows or moving to a new machine.
  - Your OneDrive sync client is broken—it's crashing, getting stuck, or constantly trying to download files.
  - Standard copy-paste or even `robocopy` commands fail because they either try to download online-only files (and get stuck) or throw errors like "The cloud file provider exited unexpectedly" or "The process cannot access the file because it is being used by another process."

This happens because modern OneDrive uses "Files On-Demand" to save disk space. Many files are just placeholders (or reparse points) that trigger a download on access. If the client that handles these downloads is malfunctioning, any attempt to read these files fails.

## The Solution

This PowerShell script bypasses the OneDrive client's problematic behavior. Instead of relying on simple copy commands, it inspects the low-level attributes of every single file.

> 💡 **The Magic:** The script checks each file and folder for the `Offline` attribute. This is the true, system-level indicator that a file is a cloud-only placeholder. By ignoring any item with this attribute, the script **only sees and copies files that are 100% physically present on your hard drive.**

This approach ensures that no download is ever triggered, allowing you to get a clean, reliable backup of your local data, even with a completely broken OneDrive client.

## Features

  - **Truly Local-Only:** Copies only the files that physically exist on your disk.
  - **Prevents Unwanted Downloads:** Will not trigger OneDrive to download any files from the cloud.
  - **Works When OneDrive is Broken:** Bypasses the sync client, avoiding crashes and errors.
  - **Preserves Directory Structure:** Your backup will have the same folder layout as the original.
  - **No Dependencies:** Runs using built-in Windows PowerShell. No external tools needed.

## How to Use

#### Step 1: Disconnect from the Internet (Recommended)

To be absolutely certain that no network activity can occur, unplug your Ethernet cable or disconnect from Wi-Fi. This is the ultimate failsafe.

#### Step 2: Configure the Script

1.  Save the code below as a PowerShell script file (e.g., `OneDrive-Local-Backup.ps1`).

2.  Open the file in any text editor or the PowerShell ISE.

3.  Modify the two variables at the top of the script to match your system:

    ```powershell
    # --- CONFIGURE YOUR PATHS HERE ---
    $sourcePath = "C:\Users\YourUsername\OneDrive"
    $destinationPath = "D:\OneDrive_Backup"
    # -----------------------------------
    ```

#### Step 3: Run the Script

1.  Right-click the script file and choose "Run with PowerShell".
2.  Alternatively, open a PowerShell terminal, navigate to the folder where you saved the script, and run it with `./OneDrive-Local-Backup.ps1`.

The script will print its progress as it safely copies your local files to the destination.

-----

## The Script: `OneDrive-Local-Backup.ps1`

```powershell
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
    Author: [Your Name/GitHub Username]
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
```

## License

This project is licensed under the MIT License. See the [LICENSE.md](LICENSE.md) file for details. You are free to use, modify, and distribute this script.
