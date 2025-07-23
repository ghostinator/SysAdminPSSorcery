<#
.SYNOPSIS
  Uninstall OneDrive, re-enable Office Cloud Storage, and configure
  Dropbox as the default save location for Office apps.

.DESCRIPTION
  1. Removes any Office policies that disable Cloud Storage providers.
  2. Ensures online content is allowed in Office ("UseOnlineContent").
  3. Uninstalls OneDrive (32/64-bit) and cleans up leftovers.
  4. Injects Dropbox as a Cloud Storage provider into each user hive
     and the Default User hive.
  5. Logs all actions to C:\temp\dropboxupdates.log
#>

# GUID Dropbox uses in Office's Cloud Storage
$guid = 'eee1e7ca-caac-4cf9-ab6c-7160f41e36f3'


# Log file path
$logFile = "C:\temp\dropboxupdates.log"

# Create log directory if it doesn't exist
if (-not (Test-Path "C:\temp")) {
    New-Item -Path "C:\temp" -ItemType Directory -Force | Out-Null
}

# Initialize log file
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
"[$timestamp] === SCRIPT STARTED: Set-DropboxDefaultSave.ps1 ===" | Out-File -FilePath $logFile -Force

# Function to write to both console and log file
function Write-Log {
    param (
        [string]$Message,
        [switch]$NoConsole
    )
   
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$timestamp] $Message" | Out-File -FilePath $logFile -Append
   
    if (-not $NoConsole) {
        Write-Output $Message
    }
}

function ReEnable-OfficeCloudFeatures {
    Write-Log "==> Re-enabling Office Cloud Storage integration…"

    # 1) Remove any policy keys that disable Cloud Storage providers
    $policyCloudPaths = @(
        'HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\CloudStorage',
        'HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\CloudStorage'
    )
    foreach ($p in $policyCloudPaths) {
        if (Test-Path $p) {
            Remove-Item $p -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log "    Removed policy key: $p"
        }
    }

    # 2) Ensure online content is allowed (UseOnlineContent = 2)
    $policyInternetPaths = @(
        'HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Internet',
        'HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Internet'
    )
    foreach ($p in $policyInternetPaths) {
        if (-not (Test-Path $p)) {
            New-Item -Path $p -Force | Out-Null
            Write-Log "    Created policy path: $p"
        }
        New-ItemProperty -Path $p `
                        -Name UseOnlineContent `
                        -PropertyType DWord `
                        -Value 2 `
                        -Force | Out-Null
        Write-Log "    Set UseOnlineContent=2 at $p"
    }
}

function Uninstall-OneDrive {
    Write-Log "==> Uninstalling OneDrive if present…"
    $paths = @(
      "$env:SystemRoot\SysWOW64\OneDriveSetup.exe",
      "$env:SystemRoot\System32\OneDriveSetup.exe"
    )
    foreach ($exe in $paths) {
        if (Test-Path $exe) {
            Write-Log "    Running: $exe /uninstall"
            & "$exe" /uninstall | Out-Null
            break
        }
    }
}

function Cleanup-OneDriveLeftovers {
    Write-Log "==> Removing OneDrive leftover folders…"
    $folders = @(
      "$env:LOCALAPPDATA\Microsoft\OneDrive"
      (Join-Path $env:USERPROFILE "OneDrive")
      "$env:PROGRAMDATA\Microsoft OneDrive"
      "$env:PROGRAMFILES\Microsoft OneDrive"
      "$env:PROGRAMFILES(x86)\Microsoft OneDrive"
    )
    foreach ($f in $folders) {
        if (Test-Path $f) {
            Remove-Item $f -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log "    Removed $f"
        }
    }

    Write-Log "==> Deleting OneDrive scheduled task…"
    schtasks.exe /Delete /TN "OneDrive Standalone Update Task-S-*" /F 2>$null

    Write-Log "==> Removing OneDrive Group Policy key…"
    Remove-Item "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive" `
                -Recurse -Force -ErrorAction SilentlyContinue
}

function Add-DropboxCloudStorageKey {
    param (
      [string] $hivePrefix,   # e.g. "HKU\TempHive_S-1-5-21-…"
      [string] $dropboxDir    # full path to that user's Dropbox folder
    )

    Write-Log "    Adding Dropbox registry keys to $hivePrefix"
   
    # Convert HKU\TempHive_xxx to proper registry path format
    $regPath = $hivePrefix -replace "^HKU\\", "HKEY_USERS\"
   
    # Create the base key path
    $keyPath = "Software\Microsoft\Office\Common\Cloud Storage\$guid"
   
    try {
        # Create the base key
        $fullPath = Join-Path $regPath $keyPath
        if (!(Test-Path "Registry::$fullPath")) {
            New-Item -Path "Registry::$fullPath" -Force | Out-Null
            Write-Log "    Created key: $fullPath"
        }
       
        # Set the values
        $values = @{
            "Description" = "Dropbox complicates the way you create, share and collaborate. Bring your photos, docs, and videos anywhere and keep your files safe."
            "DisplayName" = "Dropbox"
            "LearnMoreURL" = "https://www.dropbox.com/"
            "LocalFolderRoot" = $dropboxDir
            "ManageURL" = "https://www.dropbox.com/account"
            "Url48x48" = "http://dl.dropbox.com/u/46565/metro/Dropbox_48x48.png"
        }
       
        foreach ($key in $values.Keys) {
            New-ItemProperty -Path "Registry::$fullPath" -Name $key -Value $values[$key] -PropertyType String -Force | Out-Null
            Write-Log "    Set value: $key = $($values[$key])" -NoConsole
        }
       
        # Create Thumbnails subkey
        $thumbPath = Join-Path $fullPath "Thumbnails"
        if (!(Test-Path "Registry::$thumbPath")) {
            New-Item -Path "Registry::$thumbPath" -Force | Out-Null
            Write-Log "    Created thumbnails key: $thumbPath"
        }
       
        # Set thumbnail values
        $thumbs = @{
            "Url16x16" = "http://dl.dropbox.com/u/46565/metro/Dropbox_16x16.png"
            "Url20x20" = "http://dl.dropbox.com/u/46565/metro/Dropbox_24x24.png"
            "Url24x24" = "http://dl.dropbox.com/u/46565/metro/Dropbox_24x24.png"
            "Url32x32" = "http://dl.dropbox.com/u/46565/metro/Dropbox_32x32.png"
            "Url40x40" = "http://dl.dropbox.com/u/46565/metro/Dropbox_40x40.png"
            "Url48x48" = "http://dl.dropbox.com/u/46565/metro/Dropbox_48x48.png"
        }
       
        foreach ($key in $thumbs.Keys) {
            New-ItemProperty -Path "Registry::$thumbPath" -Name $key -Value $thumbs[$key] -PropertyType String -Force | Out-Null
            Write-Log "    Set thumbnail: $key = $($thumbs[$key])" -NoConsole
        }
       
        Write-Log "    Successfully added Dropbox registry entries"
        return $true
    }
    catch {
        Write-Log "    ERROR adding registry keys: $_"
        return $false
    }
}
function Configure-InteractiveUser {
    Write-Log "==> Configuring interactive user"
    try {
        $user = (Get-CimInstance Win32_ComputerSystem).UserName
        if (-not $user) {
            Write-Log "    No interactive user detected"
            return
        }
        Write-Log "    Found user: $user"

        $ntAcc = New-Object System.Security.Principal.NTAccount($user)
        $sid   = $ntAcc.Translate([System.Security.Principal.SecurityIdentifier]).Value
        Write-Log "    SID: $sid"

        $plistKey   = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid"
        $profileDir = (Get-ItemProperty -Path $plistKey -Name ProfileImagePath).ProfileImagePath
        if (-not (Test-Path $profileDir)) {
            Write-Log "    Profile folder not found: $profileDir"
            return
        }

        $dropboxDir = Join-Path $profileDir 'Dropbox'
        Add-DropboxCloudStorageKey -hivePrefix "HKU\$sid" -dropboxDir $dropboxDir
        Write-Log "    ✔ Interactive user configured"
    }
    catch {
        Write-Log "    ERROR configuring interactive user: $_"
    }
}

function Configure-For-AllUsers {
    Write-Log "==> Writing Dropbox keys to each existing user hive…"
   
    $plist = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    foreach ($pi in $plist) {
        $sid = $pi.PSChildName

        # ──────────────────────────────────────────────────────────
        # Skip SYSTEM & well-known service accounts
        #   S-1-5-18  = LocalSystem
        #   S-1-5-19  = LocalService
        #   S-1-5-20  = NetworkService
        #   S-1-5-80-* = Service SIDs
        # ──────────────────────────────────────────────────────────
        if ($sid -match '^(S-1-5-(18|19|20)|S-1-5-80-)') {
            Write-Log "    Skipping system/service account SID $sid"
            continue
        }

        $profilePath = (Get-ItemProperty -Path $pi.PSPath -Name ProfileImagePath `
                        -ErrorAction SilentlyContinue).ProfileImagePath
        if (-not $profilePath -or -not (Test-Path $profilePath)) {
            Write-Log "    Skipping invalid profile for SID $sid"
            continue
        }

        # also skip the built-in “Public”, “Default User” folders
        if ($profilePath -match '\\(Public|Default User|All Users)$') {
            Write-Log "    Skipping special profile path: $profilePath"
            continue
        }

        $ntuserDat = Join-Path $profilePath 'NTUSER.DAT'
        if (-not (Test-Path $ntuserDat)) {
            Write-Log "    No NTUSER.DAT for SID $sid"
            continue
        }

        # If the user is logged in, write straight to HKU\<SID>
        if (Test-Path "Registry::HKEY_USERS\$sid") {
            Write-Log "    Hive loaded for $sid (user logged in)"
            Add-DropboxCloudStorageKey -hivePrefix "HKU\$sid" `
                                      -dropboxDir (Join-Path $profilePath 'Dropbox')
            continue
        }

        # Otherwise, load/unload temporarily
        $tempHive = "TempHive_$sid"
        Write-Log "    Loading hive for SID $sid ($profilePath)…"
        try {
            # Unload if already mounted
            if (Test-Path "Registry::HKEY_USERS\$tempHive") {
                [gc]::Collect(); Start-Sleep 2
                reg.exe UNLOAD "HKU\$tempHive" | Out-Null
                Start-Sleep 1
            }

            reg.exe LOAD "HKU\$tempHive" "$ntuserDat" | Out-Null
            if ($LASTEXITCODE -ne 0) {
                Write-Log "    Failed to load hive for $sid"
                continue
            }

            Add-DropboxCloudStorageKey -hivePrefix "HKU\$tempHive" `
                                      -dropboxDir (Join-Path $profilePath 'Dropbox')
        }
        catch {
            Write-Log ("    ERROR processing hive for {0}: {1}" -f $sid, $_)
        }
        finally {
            Write-Log "    Unloading hive for SID $sid…"
            [gc]::Collect(); Start-Sleep 2
            reg.exe UNLOAD "HKU\$tempHive" | Out-Null
        }
    }
}

function Configure-DefaultUser {
    $defaultNt = 'C:\Users\Default\NTUSER.DAT'
    if (Test-Path $defaultNt) {
        Write-Log "==> Writing Dropbox keys into Default User hive…"
        $tempHive = 'DefaultHive'
       
        try {
            # Load the hive
            & reg.exe LOAD "HKU\$tempHive" "$defaultNt" | Out-Null
            if ($LASTEXITCODE -ne 0) {
                Write-Log "    Failed to load Default User hive (error code $LASTEXITCODE)"
                return
            }
           
            $dropboxDir = '%USERPROFILE%\Dropbox'
            Add-DropboxCloudStorageKey -hivePrefix "HKU\$tempHive" -dropboxDir $dropboxDir
        }
        catch {
            Write-Log ("    Error processing Default User hive: {0}" -f $_)
        }
        finally {
            # Always try to unload the hive
            [gc]::Collect() # Force garbage collection to release file handles
            Start-Sleep -Seconds 1
            & reg.exe UNLOAD "HKU\$tempHive" | Out-Null
        }
    }
}


# ==== MAIN ====
try {
    # Get computer info for the log
    $computerInfo = "Computer: $env:COMPUTERNAME, User: $env:USERNAME, Domain: $env:USERDOMAIN"
    Write-Log "System Info: $computerInfo" -NoConsole
   
    # Run main functions
    ReEnable-OfficeCloudFeatures
    Uninstall-OneDrive
    Cleanup-OneDriveLeftovers
    Configure-InteractiveUser
    Write-Log "------InteractiveUser Processing Completed------"
    Configure-For-AllUsers
    Configure-DefaultUser

    Write-Log "✔ All done. Office will now list Dropbox as a Cloud Storage provider."
   

}
catch {
    # Log any unhandled exceptions
    Write-Log "!!! ERROR: An unhandled exception occurred: $_"
    Write-Log $_.ScriptStackTrace
   
       
    # Re-throw the error so Intune knows the script failed
    throw $_
}
