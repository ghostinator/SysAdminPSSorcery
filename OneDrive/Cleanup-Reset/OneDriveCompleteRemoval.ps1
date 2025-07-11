# =================================================================================================================
# Combined OneDrive Removal Script (All Actions Enabled by Default)
# Version: 1.4
# Disclaimer: Use at your own risk. Test before production deployment.
# =================================================================================================================

param(
    [switch]$StopProcesses = $true,
    [switch]$UninstallOneDrive = $true,
    [switch]$RemoveStartupEntries = $true,
    [switch]$RemoveModernApp = $true,
    [switch]$RemoveUserData = $true,
    [switch]$RegistryCleanup = $true,
    [switch]$PreventSetupNewUsers = $true,
    [switch]$PolicyBlock = $true,
    [switch]$RestartExplorer = $true
)

$ErrorActionPreference = "Continue"
$LogPath = "$env:TEMP"
$LogFile = "$LogPath\OneDriveCombinedRemoval_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("Info", "Warning", "Error", "Success", "Debug")]
        [string]$Level = "Info",
        [System.Management.Automation.ErrorRecord]$ErrorRecord = $null
    )
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$TimeStamp] [$Level] $Message"
    if ($ErrorRecord) {
        $LogEntry += "`n    Exception: $($ErrorRecord.Exception.Message)"
        $LogEntry += "`n    Category: $($ErrorRecord.CategoryInfo.Category)"
        $LogEntry += "`n    TargetObject: $($ErrorRecord.TargetObject)"
        $LogEntry += "`n    FullyQualifiedErrorId: $($ErrorRecord.FullyQualifiedErrorId)"
        if ($ErrorRecord.ScriptStackTrace) {
            $LogEntry += "`n    StackTrace: $($ErrorRecord.ScriptStackTrace)"
        }
    }
    $LogEntry | Out-File -FilePath $LogFile -Append -Encoding UTF8
    switch ($Level) {
        "Error" { Write-Host $Message -ForegroundColor Red }
        "Warning" { Write-Host $Message -ForegroundColor Yellow }
        "Success" { Write-Host $Message -ForegroundColor Green }
        "Debug" { Write-Host $Message -ForegroundColor Cyan }
        default { Write-Host $Message -ForegroundColor White }
    }
}

function Invoke-CommandWithLogging {
    param(
        [string]$Command,
        [string]$Arguments = "",
        [string]$Description
    )
    Write-Log "Executing: $Command $Arguments" -Level "Debug"
    try {
        if ($Arguments) {
            $result = & $Command $Arguments.Split(' ') 2>&1
        } else {
            $result = & $Command 2>&1
        }
        if ($LASTEXITCODE -eq 0) {
            Write-Log "$Description completed successfully" -Level "Success"
            if ($result) {
                Write-Log "Output: $result" -Level "Debug"
            }
        } else {
            Write-Log "$Description failed with exit code $LASTEXITCODE" -Level "Error"
            if ($result) {
                Write-Log "Error output: $result" -Level "Error"
            }
        }
        return $LASTEXITCODE -eq 0
    } catch {
        Write-Log "$Description failed with exception" -Level "Error" -ErrorRecord $_
        return $false
    }
}

If (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Log "Please run this script as Administrator!" -Level "Error"
    exit 1
}

Write-Log "=== OneDrive Combined Removal Script Started ===" -Level "Info"
Write-Log "Script version: 1.4 (All Actions Enabled by Default)" -Level "Info"
Write-Log "Current user: $env:USERNAME" -Level "Info"
Write-Log "Computer name: $env:COMPUTERNAME" -Level "Info"
Write-Log "Options selected:" -Level "Info"
Write-Log "  StopProcesses: $StopProcesses" -Level "Info"
Write-Log "  UninstallOneDrive: $UninstallOneDrive" -Level "Info"
Write-Log "  RemoveStartupEntries: $RemoveStartupEntries" -Level "Info"
Write-Log "  RemoveModernApp: $RemoveModernApp" -Level "Info"
Write-Log "  RemoveUserData: $RemoveUserData" -Level "Info"
Write-Log "  RegistryCleanup: $RegistryCleanup" -Level "Info"
Write-Log "  PreventSetupNewUsers: $PreventSetupNewUsers" -Level "Info"
Write-Log "  PolicyBlock: $PolicyBlock" -Level "Info"
Write-Log "  RestartExplorer: $RestartExplorer" -Level "Info"

$Stats = @{
    ProcessesStopped = 0
    UninstallAttempts = 0
    UninstallSuccesses = 0
    UsersProcessed = 0
    RegistryEntriesRemoved = 0
    FoldersRemoved = 0
    ErrorsEncountered = 0
    FilesSkipped = 0
}

if ($StopProcesses) {
    Write-Log "Terminating all OneDrive processes..." -Level "Info"
    try {
        $processes = Get-Process OneDrive -ErrorAction SilentlyContinue
        if ($processes) {
            foreach ($proc in $processes) {
                try {
                    Write-Log "Stopping OneDrive process (PID: $($proc.Id))" -Level "Debug"
                    $proc | Stop-Process -Force
                    $Stats.ProcessesStopped++
                } catch {
                    Write-Log "Failed to stop OneDrive process (PID: $($proc.Id))" -Level "Error" -ErrorRecord $_
                    $Stats.ErrorsEncountered++
                }
            }
            Write-Log "Terminated $($Stats.ProcessesStopped) OneDrive processes" -Level "Success"
        } else {
            Write-Log "No OneDrive processes found running" -Level "Info"
        }
    } catch {
        Write-Log "Error while checking for OneDrive processes" -Level "Error" -ErrorRecord $_
        $Stats.ErrorsEncountered++
    }
}

if ($UninstallOneDrive) {
    Write-Log "Uninstalling OneDrive system-wide..." -Level "Info"
    $setupPaths = @(
        "$env:SystemRoot\System32\OneDriveSetup.exe",
        "$env:SystemRoot\SysWOW64\OneDriveSetup.exe"
    )
    foreach ($exe in $setupPaths) {
        $Stats.UninstallAttempts++
        if (Test-Path $exe) {
            Write-Log "Found OneDrive setup at: $exe" -Level "Info"
            if (Invoke-CommandWithLogging -Command $exe -Arguments "/uninstall" -Description "OneDrive uninstall") {
                $Stats.UninstallSuccesses++
            } else {
                $Stats.ErrorsEncountered++
            }
        } else {
            Write-Log "OneDrive setup not found at: $exe" -Level "Warning"
        }
    }
}

if ($RemoveModernApp) {
    Write-Log "Attempting to uninstall the modern OneDrive app..." -Level "Info"
    try {
        $modernApp = Get-AppxPackage -AllUsers -Name "Microsoft.OneDriveSync" -ErrorAction SilentlyContinue
        if ($modernApp) {
            Write-Log "Found modern OneDrive app: $($modernApp.PackageFullName)" -Level "Info"
            $modernApp | Remove-AppxPackage -AllUsers -ErrorAction Stop
            Write-Log "Successfully uninstalled the modern OneDrive app" -Level "Success"
        } else {
            Write-Log "Modern OneDrive app not found" -Level "Info"
        }
    } catch {
        Write-Log "Error uninstalling modern OneDrive app" -Level "Error" -ErrorRecord $_
        $Stats.ErrorsEncountered++
    }
}

if ($RemoveStartupEntries) {
    Write-Log "Removing OneDrive from startup for all users..." -Level "Info"
    $userProfiles = Get-ChildItem 'C:\Users' -Directory | Where-Object {
        (Test-Path "$($_.FullName)\NTUSER.DAT") -and
        ($_.Name -notin @('Default', 'Public', 'All Users', 'Default User'))
    }
    Write-Log "Found $($userProfiles.Count) user profiles to process" -Level "Info"
    foreach ($profile in $userProfiles) {
        $userName = $profile.Name
        $userPath = $profile.FullName
        $Stats.UsersProcessed++
        Write-Log "Processing user: $userName" -Level "Info"
        try {
            $ntUserDat = "$userPath\NTUSER.DAT"
            $tempHive = "HKU\TempHive_$userName"
            Write-Log "Loading registry hive for $userName" -Level "Debug"
            $loadResult = Invoke-CommandWithLogging -Command "reg" -Arguments "load `"$tempHive`" `"$ntUserDat`"" -Description "Registry hive load for $userName"
            if ($loadResult) {
                $runKeyPath = "Registry::$tempHive\Software\Microsoft\Windows\CurrentVersion\Run"
                if (Test-Path $runKeyPath) {
                    $runKey = Get-ItemProperty -Path $runKeyPath -ErrorAction SilentlyContinue
                    if ($runKey -and $runKey.PSObject.Properties.Name -contains "OneDrive") {
                        Remove-ItemProperty -Path $runKeyPath -Name "OneDrive" -ErrorAction Stop
                        Write-Log "Removed OneDrive from Run registry key for $userName" -Level "Success"
                        $Stats.RegistryEntriesRemoved++
                    } else {
                        Write-Log "OneDrive not found in Run registry key for $userName" -Level "Info"
                    }
                } else {
                    Write-Log "Run registry key not found for $userName" -Level "Warning"
                }
                Write-Log "Unloading registry hive for $userName" -Level "Debug"
                Invoke-CommandWithLogging -Command "reg" -Arguments "unload `"$tempHive`"" -Description "Registry hive unload for $userName" | Out-Null
            } else {
                $Stats.ErrorsEncountered++
            }
        } catch {
            Write-Log "Error processing registry for $userName" -Level "Error" -ErrorRecord $_
            $Stats.ErrorsEncountered++
            try { reg unload $tempHive 2>&1 | Out-Null } catch { }
        }
        $startupFolder = "$userPath\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
        $oneDriveLnk = Join-Path $startupFolder "OneDrive.lnk"
        if (Test-Path $oneDriveLnk) {
            try {
                Remove-Item $oneDriveLnk -Force -ErrorAction Stop
                Write-Log "Removed OneDrive shortcut from startup folder for $userName" -Level "Success"
                $Stats.FoldersRemoved++
            } catch {
                Write-Log "Failed to remove OneDrive shortcut for $userName" -Level "Error" -ErrorRecord $_
                $Stats.ErrorsEncountered++
            }
        } else {
            Write-Log "OneDrive shortcut not found in startup folder for $userName" -Level "Info"
        }
    }
    $allUsersStartup = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\OneDrive.lnk"
    if (Test-Path $allUsersStartup) {
        try {
            Remove-Item $allUsersStartup -Force -ErrorAction Stop
            Write-Log "Removed OneDrive shortcut from All Users startup folder" -Level "Success"
            $Stats.FoldersRemoved++
        } catch {
            Write-Log "Failed to remove OneDrive shortcut from All Users startup folder" -Level "Error" -ErrorRecord $_
            $Stats.ErrorsEncountered++
        }
    } else {
        Write-Log "OneDrive shortcut not found in All Users startup folder" -Level "Info"
    }
}

if ($RemoveUserData) {
    Write-Log "Removing OneDrive data and credentials for all users..." -Level "Info"
    $userProfiles = Get-ChildItem 'C:\Users' -Directory | Where-Object { 
        $_.Name -notin @('Default', 'Public', 'All Users', 'Default User')
    }
    foreach ($profile in $userProfiles) {
        $userName = $profile.Name
        $userPath = $profile.FullName
        Write-Log "Removing OneDrive data for user: $userName" -Level "Info"
        $folders = @(
            "$userPath\OneDrive",
            "$userPath\AppData\Local\Microsoft\OneDrive",
            "$userPath\AppData\Roaming\Microsoft\OneDrive"
        )
        foreach ($folder in $folders) {
            if (Test-Path $folder) {
                try {
                    $folderSize = (Get-ChildItem $folder -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                    $folderSizeMB = [math]::Round($folderSize / 1MB, 2)
                    Remove-Item $folder -Recurse -Force -ErrorAction Stop
                    Write-Log "Removed OneDrive folder: $folder (Size: $folderSizeMB MB)" -Level "Success"
                    $Stats.FoldersRemoved++
                } catch {
                    Write-Log "Failed to remove OneDrive folder: $folder" -Level "Error" -ErrorRecord $_
                    $Stats.ErrorsEncountered++
                    try {
                        $lockedFiles = Get-ChildItem $folder -Recurse -ErrorAction SilentlyContinue | Where-Object {
                            try {
                                [System.IO.File]::OpenWrite($_.FullName).Close()
                                $false
                            } catch {
                                $true
                            }
                        }
                        if ($lockedFiles) {
                            Write-Log "Locked files preventing removal:" -Level "Warning"
                            foreach ($file in $lockedFiles) {
                                Write-Log "  - $($file.FullName)" -Level "Warning"
                                $Stats.FilesSkipped++
                            }
                        }
                    } catch {
                        Write-Log "Could not enumerate locked files in $folder" -Level "Warning"
                    }
                }
            } else {
                Write-Log "OneDrive folder not found: $folder" -Level "Info"
            }
        }
    }
    $systemFolders = @("C:\ProgramData\Microsoft OneDrive", "C:\OneDriveTemp")
    foreach ($sysFolder in $systemFolders) {
        if (Test-Path $sysFolder) {
            try {
                Remove-Item $sysFolder -Recurse -Force -ErrorAction Stop
                Write-Log "Removed system OneDrive folder: $sysFolder" -Level "Success"
                $Stats.FoldersRemoved++
            } catch {
                Write-Log "Failed to remove system OneDrive folder: $sysFolder" -Level "Error" -ErrorRecord $_
                $Stats.ErrorsEncountered++
            }
        } else {
            Write-Log "System OneDrive folder not found: $sysFolder" -Level "Info"
        }
    }
}

if ($RegistryCleanup) {
    Write-Log "Scrubbing the registry of OneDrive entries..." -Level "Info"
    $navPaneKeys = @(
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{018D5C66-4533-4307-9B53-224DE2ED1FE6}",
        "HKCU:\Software\Classes\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}",
        "HKCU:\Software\Classes\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}"
    )
    foreach ($key in $navPaneKeys) {
        if (Test-Path $key) {
            try {
                Remove-Item -Path $key -Recurse -Force -ErrorAction Stop
                Write-Log "Removed registry key: $key" -Level "Success"
                $Stats.RegistryEntriesRemoved++
            } catch {
                Write-Log "Failed to remove registry key: $key" -Level "Error" -ErrorRecord $_
                $Stats.ErrorsEncountered++
            }
        } else {
            Write-Log "Registry key not found: $key" -Level "Info"
        }
    }
    $pathsToClean = @(
        "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive",
        "HKCU:\Software\Microsoft\OneDrive",
        "HKLM:\SOFTWARE\Microsoft\OneDrive"
    )
    foreach ($path in $pathsToClean) {
        if (Test-Path $path) {
            try {
                Remove-Item -Path $path -Recurse -Force -ErrorAction Stop
                Write-Log "Removed registry path: $path" -Level "Success"
                $Stats.RegistryEntriesRemoved++
            } catch {
                Write-Log "Failed to remove registry path: $path" -Level "Error" -ErrorRecord $_
                $Stats.ErrorsEncountered++
            }
        } else {
            Write-Log "Registry path not found: $path" -Level "Info"
        }
    }
}

if ($PreventSetupNewUsers) {
    Write-Log "Preventing OneDrive setup for new users..." -Level "Info"
    $defaultUserHive = "hklm\Default_profile"
    $defaultUserDat = "C:\Users\Default\NTUSER.DAT"
    if (Test-Path $defaultUserDat) {
        if (Invoke-CommandWithLogging -Command "reg" -Arguments "load `"$defaultUserHive`" `"$defaultUserDat`"" -Description "Load Default user hive") {
            if (Invoke-CommandWithLogging -Command "reg" -Arguments "delete `"$defaultUserHive\SOFTWARE\Microsoft\Windows\CurrentVersion\Run`" /v `"OneDriveSetup`" /f" -Description "Remove OneDriveSetup from Default user") {
                $Stats.RegistryEntriesRemoved++
            } else {
                $Stats.ErrorsEncountered++
            }
            Invoke-CommandWithLogging -Command "reg" -Arguments "unload `"$defaultUserHive`"" -Description "Unload Default user hive" | Out-Null
        } else {
            $Stats.ErrorsEncountered++
        }
    } else {
        Write-Log "Default user NTUSER.DAT not found at: $defaultUserDat" -Level "Warning"
    }
}

if ($PolicyBlock) {
    Write-Log "Applying registry policy to disable OneDrive file sync..." -Level "Info"
    $policyRegPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive"
    try {
        if (-not (Test-Path $policyRegPath)) {
            New-Item -Path $policyRegPath -Force -ErrorAction Stop | Out-Null
            Write-Log "Created policy registry path: $policyRegPath" -Level "Success"
        }
        Set-ItemProperty -Path $policyRegPath -Name "DisableFileSyncNGSC" -Value 1 -Type DWord -Force -ErrorAction Stop
        Write-Log "Set registry policy to disable OneDrive file sync" -Level "Success"
        $Stats.RegistryEntriesRemoved++
    } catch {
        Write-Log "Failed to set policy registry value" -Level "Error" -ErrorRecord $_
        $Stats.ErrorsEncountered++
    }
}

if ($RestartExplorer) {
    Write-Log "Restarting Windows Explorer to apply changes..." -Level "Info"
    try {
        $explorerProcesses = Get-Process explorer -ErrorAction SilentlyContinue
        if ($explorerProcesses) {
            $explorerProcesses | Stop-Process -Force -ErrorAction Stop
            Write-Log "Stopped Explorer processes" -Level "Success"
        }
        Start-Sleep -Seconds 2
        Start-Process -FilePath "explorer.exe" -ErrorAction Stop
        Write-Log "Windows Explorer restarted" -Level "Success"
    } catch {
        Write-Log "Failed to restart Explorer" -Level "Error" -ErrorRecord $_
        $Stats.ErrorsEncountered++
    }
}

Write-Log "==========================================================" -Level "Info"
Write-Log "OneDrive combined removal script has completed" -Level "Success"
Write-Log "EXECUTION STATISTICS:" -Level "Info"
Write-Log "  Processes stopped: $($Stats.ProcessesStopped)" -Level "Info"
Write-Log "  Uninstall attempts: $($Stats.UninstallAttempts)" -Level "Info"
Write-Log "  Uninstall successes: $($Stats.UninstallSuccesses)" -Level "Info"
Write-Log "  Users processed: $($Stats.UsersProcessed)" -Level "Info"
Write-Log "  Registry entries removed: $($Stats.RegistryEntriesRemoved)" -Level "Info"
Write-Log "  Folders removed: $($Stats.FoldersRemoved)" -Level "Info"
Write-Log "  Files skipped (locked): $($Stats.FilesSkipped)" -Level "Info"
Write-Log "  Errors encountered: $($Stats.ErrorsEncountered)" -Level "Info"

if ($Stats.ErrorsEncountered -gt 0) {
    Write-Log "Script completed with $($Stats.ErrorsEncountered) errors - review log for details" -Level "Warning"
} else {
    Write-Log "Script completed successfully with no errors" -Level "Success"
}
Write-Log "A system restart is recommended to finalize all changes" -Level "Warning"
Write-Log "==========================================================" -Level "Info"
Write-Log "Log file saved to: $LogFile" -Level "Info"

exit $(if ($Stats.ErrorsEncountered -gt 0) { 1 } else { 0 })
