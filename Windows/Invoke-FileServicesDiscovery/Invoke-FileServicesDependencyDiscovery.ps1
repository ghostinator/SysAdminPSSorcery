<#
.SYNOPSIS
    Read-only second-pass dependency discovery for Windows file servers.

.DESCRIPTION
    Supplements a general file-services inventory by locating references that can
    prevent safe migration or retirement. Collects service configuration, service
    executable metadata, scheduled-task XML, printers and ports, SyncToy and
    SyncBack profiles, SharePoint Migration Tool artifacts, scripts and common
    configuration files, mapped drives, logon scripts, shares, DFS information,
    data deduplication settings, FSRM configuration, and optional Group Policy data.

    No client or organization names are embedded. Search terms are supplied with
    -SearchTerm and can include server names, UNC paths, share names, account names,
    or other literal strings.

    The script does not modify configuration. It writes a timestamped folder and ZIP.

.NOTES
    Requires Windows PowerShell 5.1. Run locally in an elevated session.
#>

[CmdletBinding()]
param(
    [string]$OutputRoot = 'C:\Temp',

    [string[]]$SearchTerm = @($env:COMPUTERNAME),

    [string[]]$AdditionalScanPath = @(),

    [ValidateRange(1, 20)]
    [int]$MaxFileSizeMB = 5,

    [ValidateRange(1, 100000)]
    [int]$MaxMatchesPerSource = 5000,

    [switch]$IncludeGroupPolicy,

    [switch]$IncludeRegistrySearch,

    [switch]$IncludeUserProfiles
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'

$ComputerName = $env:COMPUTERNAME
$Timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$OutputPath = Join-Path $OutputRoot "FileServicesDependencyDiscovery-$ComputerName-$Timestamp"
$ZipPath = "$OutputPath.zip"
$ErrorLog = Join-Path $OutputPath 'Errors.txt'
$RunLog = Join-Path $OutputPath 'RunLog.txt'

New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null

function Write-RunLog {
    param([string]$Message)
    $Line = '{0}  {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Write-Host $Line
    Add-Content -LiteralPath $RunLog -Value $Line
}

function Write-DiscoveryError {
    param([string]$Section, [object]$Exception)
    $Text = "[{0}]`r`nSECTION: {1}`r`nERROR:   {2}`r`n" -f `
        (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Section, $Exception.Exception.Message
    Add-Content -LiteralPath $ErrorLog -Value $Text
}

function Export-SafeCsv {
    param([string]$Name, [scriptblock]$Command)
    Write-RunLog "Collecting $Name"
    try {
        $Result = @(& $Command)
        if ($Result.Count -gt 0) {
            $Result | Export-Csv -LiteralPath (Join-Path $OutputPath "$Name.csv") -NoTypeInformation -Encoding UTF8
        }
        else {
            'No results returned.' | Set-Content -LiteralPath (Join-Path $OutputPath "$Name.txt") -Encoding UTF8
        }
    }
    catch { Write-DiscoveryError -Section $Name -Exception $_ }
}

function Export-SafeText {
    param([string]$Name, [scriptblock]$Command)
    Write-RunLog "Collecting $Name"
    try {
        (& $Command | Out-String -Width 8192) |
            Set-Content -LiteralPath (Join-Path $OutputPath "$Name.txt") -Encoding UTF8
    }
    catch { Write-DiscoveryError -Section $Name -Exception $_ }
}

function Get-LiteralMatches {
    param([AllowNull()][string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
    $Hits = foreach ($Term in $script:NormalizedSearchTerms) {
        if ($Text.IndexOf($Term, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $Term }
    }
    @($Hits | Select-Object -Unique)
}

function Protect-Value {
    param([AllowNull()][string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $Value }
    if ($Value -match '(?i)(password|passwd|pwd|secret|token|apikey|api_key|connectionstring)\s*[:=]') {
        return '[REDACTED: potentially sensitive value]'
    }
    return $Value
}

function Get-TextFileMatch {
    param([string[]]$Roots, [string]$SourceType)

    $Extensions = @(
        '.config','.xml','.json','.ini','.cfg','.conf','.properties','.yml','.yaml',
        '.txt','.log','.cmd','.bat','.ps1','.psm1','.psd1','.vbs','.js','.wsf','.sql',
        '.csv','.rdp','.url','.reg','.job','.profile','.ini.bak'
    )
    $MaxBytes = $MaxFileSizeMB * 1MB
    $Count = 0

    foreach ($Root in ($Roots | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -Unique)) {
        Write-RunLog "Scanning text/config files under $Root"
        Get-ChildItem -LiteralPath $Root -File -Recurse -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.Length -le $MaxBytes -and $Extensions -contains $_.Extension.ToLowerInvariant() } |
            ForEach-Object {
                if ($Count -ge $MaxMatchesPerSource) { return }
                $File = $_
                try {
                    $LineNumber = 0
                    foreach ($Line in [IO.File]::ReadLines($File.FullName)) {
                        $LineNumber++
                        $Hits = @(Get-LiteralMatches -Text $Line)
                        if ($Hits.Count -gt 0) {
                            $Count++
                            [pscustomobject]@{
                                SourceType = $SourceType
                                FilePath   = $File.FullName
                                LineNumber = $LineNumber
                                Match      = $Hits -join '; '
                                Text       = Protect-Value -Value ($Line.Trim())
                                LastWrite  = $File.LastWriteTime
                                SizeBytes  = $File.Length
                            }
                            if ($Count -ge $MaxMatchesPerSource) { return }
                        }
                    }
                }
                catch { Write-DiscoveryError -Section "File scan: $($File.FullName)" -Exception $_ }
            }
    }
}

$script:NormalizedSearchTerms = @(
    $SearchTerm |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object { $_.Trim() } |
        Select-Object -Unique
)

if ($script:NormalizedSearchTerms.Count -eq 0) {
    throw 'Provide at least one non-empty -SearchTerm.'
}

Write-RunLog "Starting second-pass dependency discovery on $ComputerName"
Write-RunLog "Search terms: $($script:NormalizedSearchTerms -join ', ')"
Write-RunLog "Output directory: $OutputPath"

[pscustomobject]@{
    Script = 'Invoke-FileServicesDependencyDiscovery.ps1'
    ComputerName = $ComputerName
    CollectionDate = Get-Date
    User = "$env:USERDOMAIN\$env:USERNAME"
    PowerShellVersion = $PSVersionTable.PSVersion.ToString()
    SearchTerms = $script:NormalizedSearchTerms -join '; '
    IncludeGroupPolicy = [bool]$IncludeGroupPolicy
    IncludeRegistrySearch = [bool]$IncludeRegistrySearch
    IncludeUserProfiles = [bool]$IncludeUserProfiles
    MaxFileSizeMB = $MaxFileSizeMB
    MaxMatchesPerSource = $MaxMatchesPerSource
} | Export-Csv -LiteralPath (Join-Path $OutputPath '00-Manifest.csv') -NoTypeInformation -Encoding UTF8

Export-SafeCsv '01-ServiceInventory' {
    Get-CimInstance Win32_Service | Select-Object Name,DisplayName,State,StartMode,StartName,PathName
}

Export-SafeCsv '02-ServiceSearchMatches' {
    foreach ($Service in Get-CimInstance Win32_Service) {
        $Text = "$($Service.Name) $($Service.DisplayName) $($Service.StartName) $($Service.PathName)"
        $Hits = @(Get-LiteralMatches $Text)
        if ($Hits.Count -gt 0) {
            [pscustomobject]@{
                Name = $Service.Name
                DisplayName = $Service.DisplayName
                State = $Service.State
                StartMode = $Service.StartMode
                StartName = $Service.StartName
                PathName = $Service.PathName
                Match = $Hits -join '; '
            }
        }
    }
}

Export-SafeCsv '03-ServiceExecutableDetails' {
    foreach ($Service in Get-CimInstance Win32_Service) {
        $Raw = [Environment]::ExpandEnvironmentVariables([string]$Service.PathName)
        $Exe = $null
        if ($Raw -match '^\s*"([^"]+\.exe)"') { $Exe = $matches[1] }
        elseif ($Raw -match '^\s*([^\s]+\.exe)') { $Exe = $matches[1] }
        if ($Exe -and (Test-Path -LiteralPath $Exe)) {
            $Item = Get-Item -LiteralPath $Exe -ErrorAction SilentlyContinue
            [pscustomobject]@{
                ServiceName = $Service.Name
                Executable = $Exe
                Company = $Item.VersionInfo.CompanyName
                Product = $Item.VersionInfo.ProductName
                FileVersion = $Item.VersionInfo.FileVersion
                ProductVersion = $Item.VersionInfo.ProductVersion
                LastWriteTime = $Item.LastWriteTime
                StartName = $Service.StartName
            }
        }
    }
}

Export-SafeCsv '10-ScheduledTaskInventory' {
    foreach ($Task in Get-ScheduledTask -ErrorAction SilentlyContinue) {
        foreach ($Action in $Task.Actions) {
            [pscustomobject]@{
                TaskPath = $Task.TaskPath
                TaskName = $Task.TaskName
                State = $Task.State
                UserId = $Task.Principal.UserId
                Execute = $Action.Execute
                Arguments = $Action.Arguments
                WorkingDirectory = $Action.WorkingDirectory
            }
        }
    }
}

Export-SafeCsv '11-ScheduledTaskSearchMatches' {
    foreach ($Task in Get-ScheduledTask -ErrorAction SilentlyContinue) {
        $Xml = $null
        try { $Xml = Export-ScheduledTask -TaskName $Task.TaskName -TaskPath $Task.TaskPath -ErrorAction Stop } catch {}
        $Hits = @(Get-LiteralMatches $Xml)
        if ($Hits.Count -gt 0) {
            [pscustomobject]@{
                TaskPath = $Task.TaskPath
                TaskName = $Task.TaskName
                State = $Task.State
                UserId = $Task.Principal.UserId
                Match = $Hits -join '; '
                Xml = Protect-Value (($Xml -replace '\s+',' ').Trim())
            }
        }
    }
}

Export-SafeCsv '20-PrinterInventory' {
    if (Get-Command Get-Printer -ErrorAction SilentlyContinue) {
        Get-Printer | Select-Object Name,ComputerName,Type,DriverName,PortName,Shared,ShareName,Published,PermissionSDDL,RenderingMode
    }
}

Export-SafeCsv '21-PrinterPortInventory' {
    if (Get-Command Get-PrinterPort -ErrorAction SilentlyContinue) {
        Get-PrinterPort | Select-Object Name,Description,PrinterHostAddress,PortNumber,Protocol,SNMPEnabled,SNMPCommunity,DeviceURL
    }
}

Export-SafeCsv '22-PrinterSearchMatches' {
    if (Get-Command Get-Printer -ErrorAction SilentlyContinue) {
        foreach ($Printer in Get-Printer) {
            $Text = "$($Printer.Name) $($Printer.ComputerName) $($Printer.DriverName) $($Printer.PortName) $($Printer.ShareName)"
            $Hits = @(Get-LiteralMatches $Text)
            if ($Hits.Count -gt 0) {
                [pscustomobject]@{ Name=$Printer.Name; DriverName=$Printer.DriverName; PortName=$Printer.PortName; ShareName=$Printer.ShareName; Match=$Hits -join '; ' }
            }
        }
    }
}

Export-SafeCsv '30-SMBShares' {
    Get-SmbShare -ErrorAction SilentlyContinue | Select-Object Name,Path,Description,FolderEnumerationMode,CachingMode,EncryptData,Special
}

Export-SafeCsv '31-SMBMappingsAndConnections' {
    $Results = @()
    if (Get-Command Get-SmbMapping -ErrorAction SilentlyContinue) {
        $Results += Get-SmbMapping -ErrorAction SilentlyContinue | Select-Object @{n='Source';e={'Get-SmbMapping'}},LocalPath,RemotePath,Status,UserName
    }
    $Results
}

Export-SafeText '32-NetUse' { cmd.exe /c 'net use' 2>&1 }

Export-SafeCsv '33-LocalUserProfileMappedDrives' {
    $HiveRoot = 'Registry::HKEY_USERS'
    Get-ChildItem $HiveRoot -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match '^S-1-5-21-' } |
        ForEach-Object {
            $Sid = $_.PSChildName
            $NetworkPath = "$HiveRoot\$Sid\Network"
            if (Test-Path $NetworkPath) {
                Get-ChildItem $NetworkPath -ErrorAction SilentlyContinue | ForEach-Object {
                    $P = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
                    [pscustomobject]@{ SID=$Sid; DriveLetter=$_.PSChildName; RemotePath=$P.RemotePath; UserName=$P.UserName }
                }
            }
        }
}

Export-SafeCsv '40-DFSNConfiguration' {
    if (Get-Command Get-DfsnRoot -ErrorAction SilentlyContinue) {
        foreach ($Root in Get-DfsnRoot -ErrorAction SilentlyContinue) {
            [pscustomobject]@{ RecordType='Root'; Root=$Root.Path; Folder=''; Target=''; State=$Root.State }
            foreach ($Folder in Get-DfsnFolder -Path "$($Root.Path)\*" -ErrorAction SilentlyContinue) {
                [pscustomobject]@{ RecordType='Folder'; Root=$Root.Path; Folder=$Folder.Path; Target=''; State=$Folder.State }
                foreach ($Target in Get-DfsnFolderTarget -Path $Folder.Path -ErrorAction SilentlyContinue) {
                    [pscustomobject]@{ RecordType='FolderTarget'; Root=$Root.Path; Folder=$Folder.Path; Target=$Target.TargetPath; State=$Target.State }
                }
            }
        }
    }
}

Export-SafeCsv '41-DFSRConfiguration' {
    if (Get-Command Get-DfsReplicationGroup -ErrorAction SilentlyContinue) {
        foreach ($Group in Get-DfsReplicationGroup -ErrorAction SilentlyContinue) {
            foreach ($Member in Get-DfsrMembership -GroupName $Group.GroupName -ErrorAction SilentlyContinue) {
                [pscustomobject]@{ GroupName=$Group.GroupName; ComputerName=$Member.ComputerName; FolderName=$Member.FolderName; ContentPath=$Member.ContentPath; Enabled=$Member.Enabled; State=$Member.State }
            }
        }
    }
}

Export-SafeCsv '50-DeduplicationConfiguration' {
    if (Get-Command Get-DedupVolume -ErrorAction SilentlyContinue) {
        Get-DedupVolume -ErrorAction SilentlyContinue | Select-Object Volume,Enabled,UsageType,SavedSpace,SavingsRate,UnoptimizedSize
    }
}

Export-SafeCsv '51-DeduplicationJobs' {
    if (Get-Command Get-DedupJob -ErrorAction SilentlyContinue) { Get-DedupJob -ErrorAction SilentlyContinue }
}

Export-SafeText '52-FSRMConfiguration' {
    if (Get-Command Get-FsrmQuota -ErrorAction SilentlyContinue) {
        '=== Quotas ==='; Get-FsrmQuota -ErrorAction SilentlyContinue | Format-List *
        '=== File Screens ==='; Get-FsrmFileScreen -ErrorAction SilentlyContinue | Format-List *
        '=== Classification Rules ==='; Get-FsrmClassificationRule -ErrorAction SilentlyContinue | Format-List *
        '=== Storage Reports ==='; Get-FsrmStorageReport -ErrorAction SilentlyContinue | Format-List *
    }
}

$ProgramRoots = @($env:ProgramData, $env:ProgramFiles, ${env:ProgramFiles(x86)}) + $AdditionalScanPath
if ($IncludeUserProfiles) { $ProgramRoots += 'C:\Users' }

Export-SafeCsv '60-TextAndConfigSearchMatches' {
    Get-TextFileMatch -Roots $ProgramRoots -SourceType 'ProgramAndConfigFile'
}

$KnownProductRoots = @(
    "$env:ProgramData\2BrightSparks",
    "$env:APPDATA\2BrightSparks",
    "$env:LOCALAPPDATA\2BrightSparks",
    "$env:LOCALAPPDATA\Microsoft\SyncToy",
    "$env:APPDATA\Microsoft\SPMigration",
    "$env:LOCALAPPDATA\Microsoft\SPMigration",
    "$env:ProgramData\Microsoft\SPMigration"
)
if ($IncludeUserProfiles) {
    $KnownProductRoots += Get-ChildItem 'C:\Users' -Directory -Force -ErrorAction SilentlyContinue | ForEach-Object {
        @(
            "$($_.FullName)\AppData\Roaming\2BrightSparks",
            "$($_.FullName)\AppData\Local\Microsoft\SyncToy",
            "$($_.FullName)\AppData\Roaming\Microsoft\SPMigration",
            "$($_.FullName)\AppData\Local\Microsoft\SPMigration"
        )
    }
}

Export-SafeCsv '61-SyncAndMigrationProductMatches' {
    Get-TextFileMatch -Roots $KnownProductRoots -SourceType 'SyncOrMigrationProduct'
}

Export-SafeCsv '62-KnownProductArtifacts' {
    foreach ($Root in ($KnownProductRoots | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -Unique)) {
        Get-ChildItem -LiteralPath $Root -File -Recurse -Force -ErrorAction SilentlyContinue |
            Select-Object @{n='ProductRoot';e={$Root}},FullName,Length,LastWriteTime,Extension
    }
}

if ($IncludeRegistrySearch) {
    Export-SafeCsv '70-RegistrySearchMatches' {
        $RegistryRoots = @(
            'HKLM:\SOFTWARE',
            'HKLM:\SYSTEM\CurrentControlSet\Services'
        )
        foreach ($Root in $RegistryRoots) {
            Get-ChildItem -LiteralPath $Root -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
                $Key = $_
                try {
                    $Props = Get-ItemProperty -LiteralPath $Key.PSPath -ErrorAction Stop
                    foreach ($Property in $Props.PSObject.Properties) {
                        if ($Property.Name -like 'PS*') { continue }
                        $Value = [string]$Property.Value
                        $Hits = @(Get-LiteralMatches $Value)
                        if ($Hits.Count -gt 0) {
                            [pscustomobject]@{
                                Key = $Key.Name
                                Property = $Property.Name
                                Value = Protect-Value $Value
                                Match = $Hits -join '; '
                            }
                        }
                    }
                } catch {}
            }
        }
    }
}

if ($IncludeGroupPolicy) {
    Export-SafeText '80-GPResultComputer' { gpresult.exe /scope computer /z 2>&1 }

    Export-SafeCsv '81-ActiveDirectoryLogonScripts' {
        if (Get-Command Get-ADUser -ErrorAction SilentlyContinue) {
            Get-ADUser -Filter 'ScriptPath -like "*"' -Properties ScriptPath,HomeDirectory,HomeDrive -ErrorAction SilentlyContinue |
                Where-Object { $_.ScriptPath -or $_.HomeDirectory -or $_.HomeDrive } |
                Select-Object SamAccountName,Enabled,ScriptPath,HomeDirectory,HomeDrive
        }
    }

    Export-SafeCsv '82-GPOReportSearchMatches' {
        if (Get-Command Get-GPOReport -ErrorAction SilentlyContinue) {
            [xml]$Report = Get-GPOReport -All -ReportType Xml -ErrorAction Stop
            foreach ($Gpo in $Report.GPOs.GPO) {
                $Xml = $Gpo.OuterXml
                $Hits = @(Get-LiteralMatches $Xml)
                if ($Hits.Count -gt 0) {
                    [pscustomobject]@{ Name=$Gpo.Name; ID=$Gpo.Identifier.Identifier.'#text'; Match=$Hits -join '; ' }
                }
            }
        }
    }
}

Export-SafeCsv '90-RecentRelevantEvents' {
    $Logs = @('Microsoft-Windows-SMBServer/Operational','DFS Replication','System')
    $Start = (Get-Date).AddDays(-30)
    foreach ($Log in $Logs) {
        try {
            Get-WinEvent -FilterHashtable @{LogName=$Log;StartTime=$Start} -ErrorAction Stop |
                Where-Object { $_.Level -in 1,2,3 } |
                Select-Object @{n='LogName';e={$Log}},TimeCreated,Id,LevelDisplayName,ProviderName,Message
        } catch {}
    }
}

Export-SafeText '99-QuickSummary' {
    "FILE SERVICES DEPENDENCY DISCOVERY SUMMARY"
    "=========================================="
    "Computer:     $ComputerName"
    "Date:         $(Get-Date)"
    "Search terms: $($script:NormalizedSearchTerms -join ', ')"
    ""
    "Review these reports first:"
    "  02-ServiceSearchMatches"
    "  11-ScheduledTaskSearchMatches"
    "  22-PrinterSearchMatches"
    "  33-LocalUserProfileMappedDrives"
    "  60-TextAndConfigSearchMatches"
    "  61-SyncAndMigrationProductMatches"
    "  70-RegistrySearchMatches, when enabled"
    "  82-GPOReportSearchMatches, when enabled"
    ""
    "Important: no-match results do not prove that a dependency does not exist."
}

Write-RunLog 'Collection complete. Creating ZIP archive.'
try {
    Compress-Archive -Path "$OutputPath\*" -DestinationPath $ZipPath -CompressionLevel Optimal -Force
    Write-RunLog "ZIP created: $ZipPath"
}
catch {
    Write-DiscoveryError -Section 'Compress output' -Exception $_
}

Write-Host ''
Write-Host 'DEPENDENCY DISCOVERY COMPLETE' -ForegroundColor Green
Write-Host "Folder: $OutputPath"
if (Test-Path -LiteralPath $ZipPath) { Write-Host "ZIP:    $ZipPath" -ForegroundColor Cyan }
Write-Host 'No configuration changes were intentionally performed.' -ForegroundColor Yellow
