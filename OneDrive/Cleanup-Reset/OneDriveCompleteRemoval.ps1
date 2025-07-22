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
  6. Emails log file contents upon completion
#>

# GUID Dropbox uses in Office's Cloud Storage
$guid = 'eee1e7ca-caac-4cf9-ab6c-7160f41e36f3'

# Email Configuration
$emailConfig = @{
    SMTPServer = "mx.app.ghosties.email"
    Port       = 587                    # Update if needed
    From       = "powershell@app.ghosties.email" # Update with your sender address
    To         = "brandon.cook@gadellnet.com" # Update with recipient address
    Subject    = "Mercy - Dropbox Office Integration Deployment Report"
    UseSsl     = $true
    Credential = $null                  # Will be set later
}

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

function Configure-For-AllUsers {
    Write-Log "==> Writing Dropbox keys to each existing user hive…"
    $plist = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    foreach ($pi in $plist) {
        $sid = $pi.PSChildName
        $profilePath = (Get-ItemProperty -Path $pi.PSPath -Name ProfileImagePath).ProfileImagePath
        if (-not (Test-Path $profilePath))                             { continue }
        if ($profilePath -match '\\(Public|Default User|All Users)$') { continue }

        $ntuserDat = Join-Path $profilePath 'NTUSER.DAT'
        if (-not (Test-Path $ntuserDat)) { continue }

        $tempHive = "TempHive_$sid"
        Write-Log "    Loading hive for SID $sid ($profilePath)…"
        
        try {
            # Load the hive
            & reg.exe LOAD "HKU\$tempHive" "$ntuserDat" | Out-Null
            if ($LASTEXITCODE -ne 0) {
                Write-Log "    Failed to load hive for $sid (error code $LASTEXITCODE)"
                continue
            }
            
            $dropboxDir = Join-Path $profilePath 'Dropbox'
            Add-DropboxCloudStorageKey -hivePrefix "HKU\$tempHive" -dropboxDir $dropboxDir
        }
        catch {
            Write-Log ("    Error processing hive for {0}: {1}" -f $sid, $_)
        }
        finally {
            # Always try to unload the hive
            Write-Log "    Unloading hive for $sid…"
            [gc]::Collect() # Force garbage collection to release file handles
            Start-Sleep -Seconds 1
            & reg.exe UNLOAD "HKU\$tempHive" | Out-Null
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

function Send-CompletionEmail {
    Write-Log "==> Preparing to send email notification..."
    
    # Get computer name and IP for the report
    $computerName = $env:COMPUTERNAME
    try {
        $ipAddresses = [System.Net.Dns]::GetHostAddresses($computerName) | 
                       Where-Object { $_.AddressFamily -eq 'InterNetwork' } | 
                       ForEach-Object { $_.IPAddressToString }
        $ipInfo = $ipAddresses -join ', '
    }
    catch {
        $ipInfo = "Unable to determine IP"
        Write-Log "    Warning: Could not get IP address: $_"
    }
    
    # Read log file content
    try {
        $logContent = Get-Content -Path $logFile -Raw -ErrorAction Stop
    }
    catch {
        $logContent = "Error reading log file: $_"
        Write-Log "    Error reading log file: $_"
    }
    
    # Create email body with HTML formatting
    $emailBody = @"
<html>
<body style="font-family: Calibri, Arial, sans-serif; font-size: 12pt;">
<h2>Dropbox Office Integration Deployment Report</h2>
<p><strong>Computer:</strong> $computerName</p>
<p><strong>IP Address:</strong> $ipInfo</p>
<p><strong>Execution Time:</strong> $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>

<h3>Log Output:</h3>
<pre style="background-color: #f5f5f5; padding: 10px; border: 1px solid #ddd; font-family: Consolas, monospace; white-space: pre-wrap;">
$logContent
</pre>
</body>
</html>
"@

    # Set up credentials for SMTP
    try {
        Write-Log "    Setting up email credentials..."
    $securePassword = ConvertTo-SecureString "F0ZqzIVb1cGgGGql61OgZE3J" -AsPlainText -Force
    $emailConfig.Credential = New-Object System.Management.Automation.PSCredential("powershell@app.ghosties.email", $securePassword)
        
        # Ignore certificate validation
        Write-Log "    Configuring to ignore invalid certificates..."
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
        
        # Send the email
        Write-Log "    Sending email to $($emailConfig.To)..."
        Send-MailMessage @emailConfig -Body $emailBody -BodyAsHtml -ErrorAction Stop
        Write-Log "==> Email notification sent successfully"
    }
    catch {
        Write-Log "==> ERROR: Failed to send email notification: $_"
        Write-Log "    Email server: $($emailConfig.SMTPServer):$($emailConfig.Port)"
        Write-Log "    From: $($emailConfig.From), To: $($emailConfig.To)"
    }
    finally {
        # Reset certificate validation to default
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
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
    Configure-For-AllUsers
    Configure-DefaultUser

    Write-Log "✔ All done. Office will now list Dropbox as a Cloud Storage provider."
    
    # Send email with log contents
    Send-CompletionEmail
}
catch {
    # Log any unhandled exceptions
    Write-Log "!!! ERROR: An unhandled exception occurred: $_"
    Write-Log $_.ScriptStackTrace
    
    # Try to send email with error info
    Send-CompletionEmail
    
    # Re-throw the error so Intune knows the script failed
    throw $_
}
