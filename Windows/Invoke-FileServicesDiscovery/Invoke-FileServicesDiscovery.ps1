<#
.SYNOPSIS
    Read-only discovery for Windows file server, SMB, DFS Namespace, DFS-R,
    storage, permissions, workload dependencies, and migration readiness.


    Run locally as Administrator servers hosting shares

    Produces a timestamped folder and ZIP containing CSV/TXT discovery output.

    The script does NOT:
      - Change shares
      - Change NTFS permissions
      - Change DFS/DFSR configuration
      - Stop services
      - Modify files
      - Install Windows features/modules

.NOTES
    Compatible with Windows PowerShell 5.1.
#>

[CmdletBinding()]
param(
    [string]$OutputRoot = 'C:\Temp',
    [switch]$DeepFileScan
)

$ErrorActionPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'

$ComputerName = $env:COMPUTERNAME
$Timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$OutputPath = Join-Path $OutputRoot "FileServicesDiscovery-$ComputerName-$Timestamp"
$ZipPath = "$OutputPath.zip"

New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null

$ErrorLog = Join-Path $OutputPath 'Errors.txt'
$RunLog   = Join-Path $OutputPath 'RunLog.txt'

function Write-RunLog {
    param([string]$Message)

    $Line = "{0}  {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Write-Host $Line
    Add-Content -Path $RunLog -Value $Line
}

function Write-DiscoveryError {
    param(
        [string]$Section,
        [object]$Exception
    )

    $Message = @"
[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')]
SECTION: $Section
ERROR:   $($Exception.Exception.Message)

"@
    Add-Content -Path $ErrorLog -Value $Message
}

function Export-SafeCsv {
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [scriptblock]$Command
    )

    Write-RunLog "Collecting $Name"

    try {
        $Result = & $Command

        if ($null -ne $Result) {
            $Result |
                Export-Csv -Path (Join-Path $OutputPath "$Name.csv") `
                           -NoTypeInformation `
                           -Encoding UTF8
        }
        else {
            "No results returned." |
                Set-Content -Path (Join-Path $OutputPath "$Name.txt")
        }
    }
    catch {
        Write-DiscoveryError -Section $Name -Exception $_
    }
}

function Export-SafeText {
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [scriptblock]$Command
    )

    Write-RunLog "Collecting $Name"

    try {
        (& $Command | Out-String -Width 4096) |
            Set-Content -Path (Join-Path $OutputPath "$Name.txt")
    }
    catch {
        Write-DiscoveryError -Section $Name -Exception $_
    }
}

function Convert-ToReadableSize {
    param([double]$Bytes)

    if ($Bytes -ge 1TB) {
        return "{0:N2} TB" -f ($Bytes / 1TB)
    }
    elseif ($Bytes -ge 1GB) {
        return "{0:N2} GB" -f ($Bytes / 1GB)
    }
    elseif ($Bytes -ge 1MB) {
        return "{0:N2} MB" -f ($Bytes / 1MB)
    }
    elseif ($Bytes -ge 1KB) {
        return "{0:N2} KB" -f ($Bytes / 1KB)
    }

    return "$Bytes Bytes"
}

Write-RunLog "Starting file services discovery on $ComputerName"
Write-RunLog "Output directory: $OutputPath"

# ---------------------------------------------------------------------------
# Basic system information
# ---------------------------------------------------------------------------

Export-SafeCsv -Name '01-ComputerSystem' -Command {
    Get-CimInstance Win32_ComputerSystem |
        Select-Object Name,
                      Domain,
                      Manufacturer,
                      Model,
                      TotalPhysicalMemory,
                      NumberOfProcessors,
                      NumberOfLogicalProcessors
}

Export-SafeCsv -Name '02-OperatingSystem' -Command {
    Get-CimInstance Win32_OperatingSystem |
        Select-Object CSName,
                      Caption,
                      Version,
                      BuildNumber,
                      InstallDate,
                      LastBootUpTime
}

Export-SafeCsv -Name '03-Volumes' -Command {
    Get-Volume -ErrorAction Stop |
        Select-Object DriveLetter,
                      FileSystemLabel,
                      FileSystem,
                      HealthStatus,
                      OperationalStatus,
                      Size,
                      SizeRemaining,
                      @{
                          Name = 'SizeGB'
                          Expression = { [math]::Round($_.Size / 1GB, 2) }
                      },
                      @{
                          Name = 'FreeGB'
                          Expression = { [math]::Round($_.SizeRemaining / 1GB, 2) }
                      },
                      @{
                          Name = 'UsedGB'
                          Expression = {
                              [math]::Round(($_.Size - $_.SizeRemaining) / 1GB, 2)
                          }
                      }
}

Export-SafeCsv -Name '04-Disks' -Command {
    Get-Disk -ErrorAction Stop |
        Select-Object Number,
                      FriendlyName,
                      SerialNumber,
                      BusType,
                      PartitionStyle,
                      OperationalStatus,
                      HealthStatus,
                      Size
}

Export-SafeCsv -Name '05-WindowsFeatures' -Command {
    if (Get-Command Get-WindowsFeature -ErrorAction SilentlyContinue) {
        Get-WindowsFeature |
            Where-Object Installed |
            Select-Object Name, DisplayName, InstallState
    }
}

# ---------------------------------------------------------------------------
# SMB configuration
# ---------------------------------------------------------------------------

Export-SafeCsv -Name '10-SMBShares' -Command {
    Get-SmbShare |
        Select-Object Name,
                      Path,
                      Description,
                      ScopeName,
                      FolderEnumerationMode,
                      CachingMode,
                      EncryptData,
                      ContinuouslyAvailable,
                      ConcurrentUserLimit,
                      Special
}

Export-SafeCsv -Name '11-SMBSharePermissions' -Command {
    $Output = foreach ($Share in Get-SmbShare) {
        try {
            Get-SmbShareAccess -Name $Share.Name |
                Select-Object @{
                                  Name = 'ShareName'
                                  Expression = { $Share.Name }
                              },
                              AccountName,
                              AccessControlType,
                              AccessRight
        }
        catch {
            [pscustomobject]@{
                ShareName         = $Share.Name
                AccountName       = '<Unable to enumerate>'
                AccessControlType = ''
                AccessRight       = $_.Exception.Message
            }
        }
    }

    $Output
}

Export-SafeCsv -Name '12-SMBSessions' -Command {
    Get-SmbSession -ErrorAction Stop |
        Select-Object ClientComputerName,
                      ClientUserName,
                      NumOpens,
                      Dialect,
                      SecondsIdle,
                      SecondsExists
}

Export-SafeCsv -Name '13-SMBOpenFiles' -Command {
    Get-SmbOpenFile -ErrorAction Stop |
        Select-Object ClientComputerName,
                      ClientUserName,
                      Path,
                      ShareRelativePath,
                      SessionId,
                      FileId
}

Export-SafeText -Name '14-SMBServerConfiguration' -Command {
    Get-SmbServerConfiguration |
        Format-List *
}

Export-SafeText -Name '15-SMBClientConfiguration' -Command {
    Get-SmbClientConfiguration |
        Format-List *
}

# ---------------------------------------------------------------------------
# NTFS permissions at share root
# ---------------------------------------------------------------------------

Export-SafeCsv -Name '20-NTFSShareRootPermissions' -Command {

    $Results = foreach ($Share in Get-SmbShare |
            Where-Object {
                $_.Path -and
                $_.Path -match '^[A-Za-z]:\\'
            }) {

        if (Test-Path -LiteralPath $Share.Path) {

            try {
                $Acl = Get-Acl -LiteralPath $Share.Path

                foreach ($Entry in $Acl.Access) {

                    [pscustomobject]@{
                        ShareName        = $Share.Name
                        Path             = $Share.Path
                        Owner            = $Acl.Owner
                        Identity         = $Entry.IdentityReference
                        FileSystemRights = $Entry.FileSystemRights
                        AccessType       = $Entry.AccessControlType
                        IsInherited      = $Entry.IsInherited
                        InheritanceFlags = $Entry.InheritanceFlags
                        PropagationFlags = $Entry.PropagationFlags
                    }
                }
            }
            catch {
                [pscustomobject]@{
                    ShareName        = $Share.Name
                    Path             = $Share.Path
                    Owner            = ''
                    Identity         = '<ERROR>'
                    FileSystemRights = $_.Exception.Message
                    AccessType       = ''
                    IsInherited      = ''
                    InheritanceFlags = ''
                    PropagationFlags = ''
                }
            }
        }
    }

    $Results
}

# ---------------------------------------------------------------------------
# Share data sizing and basic migration profile
# ---------------------------------------------------------------------------

Export-SafeCsv -Name '21-ShareDataSummary' -Command {

    $Results = foreach ($Share in Get-SmbShare |
            Where-Object {
                $_.Path -and
                $_.Path -match '^[A-Za-z]:\\' -and
                -not $_.Special
            }) {

        Write-RunLog "Scanning share data: $($Share.Name) -> $($Share.Path)"

        if (-not (Test-Path -LiteralPath $Share.Path)) {
            continue
        }

        $FileCount = 0L
        $DirectoryCount = 0L
        $TotalBytes = 0L
        $OldestFile = $null
        $NewestFile = $null

        try {
            Get-ChildItem -LiteralPath $Share.Path `
                          -Recurse `
                          -Force `
                          -ErrorAction SilentlyContinue |
                ForEach-Object {

                    if ($_.PSIsContainer) {
                        $DirectoryCount++
                    }
                    else {
                        $FileCount++
                        $TotalBytes += $_.Length

                        if (($null -eq $OldestFile) -or
                            ($_.LastWriteTime -lt $OldestFile)) {
                            $OldestFile = $_.LastWriteTime
                        }

                        if (($null -eq $NewestFile) -or
                            ($_.LastWriteTime -gt $NewestFile)) {
                            $NewestFile = $_.LastWriteTime
                        }
                    }
                }
        }
        catch {
            Write-DiscoveryError `
                -Section "Share scan: $($Share.Name)" `
                -Exception $_
        }

        [pscustomobject]@{
            ShareName       = $Share.Name
            Path            = $Share.Path
            FileCount       = $FileCount
            DirectoryCount  = $DirectoryCount
            TotalBytes      = $TotalBytes
            TotalSize       = Convert-ToReadableSize $TotalBytes
            TotalSizeGB     = [math]::Round($TotalBytes / 1GB, 3)
            OldestLastWrite = $OldestFile
            NewestLastWrite = $NewestFile
        }
    }

    $Results
}

# ---------------------------------------------------------------------------
# Extension profiling and potentially problematic workload types
# ---------------------------------------------------------------------------

Export-SafeCsv -Name '22-FileExtensionSummary' -Command {

    $Results = foreach ($Share in Get-SmbShare |
            Where-Object {
                $_.Path -and
                $_.Path -match '^[A-Za-z]:\\' -and
                -not $_.Special
            }) {

        if (-not (Test-Path -LiteralPath $Share.Path)) {
            continue
        }

        $Files = Get-ChildItem -LiteralPath $Share.Path `
                               -File `
                               -Recurse `
                               -Force `
                               -ErrorAction SilentlyContinue

        $Files |
            Group-Object Extension |
            ForEach-Object {

                $Bytes = ($_.Group |
                    Measure-Object Length -Sum).Sum

                [pscustomobject]@{
                    ShareName = $Share.Name
                    Extension = if ($_.Name) { $_.Name.ToLower() } else { '<none>' }
                    FileCount = $_.Count
                    TotalBytes = $Bytes
                    TotalSizeGB = [math]::Round($Bytes / 1GB, 3)
                }
            }
    }

    $Results
}

$ProblemExtensions = @(
    '.pst',
    '.ost',
    '.mdb',
    '.accdb',
    '.mdf',
    '.ldf',
    '.ndf',
    '.qbw',
    '.qbb',
    '.qbm',
    '.vhd',
    '.vhdx',
    '.avhd',
    '.avhdx',
    '.bak',
    '.bkf',
    '.iso',
    '.exe',
    '.msi',
    '.dll',
    '.dwg',
    '.dxf',
    '.rvt',
    '.sldprt',
    '.sldasm'
)

Export-SafeCsv -Name '23-ProblemFileTypes' -Command {

    foreach ($Share in Get-SmbShare |
            Where-Object {
                $_.Path -and
                $_.Path -match '^[A-Za-z]:\\' -and
                -not $_.Special
            }) {

        if (-not (Test-Path -LiteralPath $Share.Path)) {
            continue
        }

        Get-ChildItem -LiteralPath $Share.Path `
                      -File `
                      -Recurse `
                      -Force `
                      -ErrorAction SilentlyContinue |
            Where-Object {
                $ProblemExtensions -contains $_.Extension.ToLower()
            } |
            Select-Object @{
                              Name = 'ShareName'
                              Expression = { $Share.Name }
                          },
                          FullName,
                          Extension,
                          Length,
                          LastWriteTime
    }
}

# ---------------------------------------------------------------------------
# EFS discovery
# ---------------------------------------------------------------------------

Export-SafeText -Name '24-EFSEncryptedFiles' -Command {

    foreach ($Share in Get-SmbShare |
            Where-Object {
                $_.Path -and
                $_.Path -match '^[A-Za-z]:\\' -and
                -not $_.Special
            }) {

        if (Test-Path -LiteralPath $Share.Path) {
            "============================================================"
            "SHARE: $($Share.Name)"
            "PATH:  $($Share.Path)"
            "============================================================"

            cipher.exe /u /n /h /s:"$($Share.Path)" 2>&1
            ""
        }
    }
}

# ---------------------------------------------------------------------------
# Optional deeper filename/path analysis
# ---------------------------------------------------------------------------

if ($DeepFileScan) {

    Export-SafeCsv -Name '25-LongPaths' -Command {

        foreach ($Share in Get-SmbShare |
                Where-Object {
                    $_.Path -and
                    $_.Path -match '^[A-Za-z]:\\' -and
                    -not $_.Special
                }) {

            if (-not (Test-Path -LiteralPath $Share.Path)) {
                continue
            }

            Get-ChildItem -LiteralPath $Share.Path `
                          -Recurse `
                          -Force `
                          -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.FullName.Length -ge 240
                } |
                Select-Object @{
                                  Name = 'ShareName'
                                  Expression = { $Share.Name }
                              },
                              FullName,
                              @{
                                  Name = 'PathLength'
                                  Expression = { $_.FullName.Length }
                              },
                              PSIsContainer,
                              LastWriteTime
        }
    }

    Export-SafeCsv -Name '26-PotentialCloudMigrationNameIssues' -Command {

        foreach ($Share in Get-SmbShare |
                Where-Object {
                    $_.Path -and
                    $_.Path -match '^[A-Za-z]:\\' -and
                    -not $_.Special
                }) {

            if (-not (Test-Path -LiteralPath $Share.Path)) {
                continue
            }

            Get-ChildItem -LiteralPath $Share.Path `
                          -Recurse `
                          -Force `
                          -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.Name -match '[#%]' -or
                    $_.Name.EndsWith('.') -or
                    $_.Name.EndsWith(' ')
                } |
                Select-Object @{
                                  Name = 'ShareName'
                                  Expression = { $Share.Name }
                              },
                              FullName,
                              Name,
                              PSIsContainer
        }
    }
}

# ---------------------------------------------------------------------------
# DFS Namespace
# ---------------------------------------------------------------------------

$DfsNAvailable = $false

try {
    Import-Module Dfsn -ErrorAction Stop
    $DfsNAvailable = $true
}
catch {
    Write-DiscoveryError -Section 'Import DFSN module' -Exception $_
}

if ($DfsNAvailable) {

    Export-SafeCsv -Name '30-DFSNamespaceRoots' -Command {
        Get-DfsnRoot -ErrorAction Stop |
            Select-Object Path,
                          Type,
                          State,
                          Description,
                          TimeToLiveSec
    }

    Export-SafeCsv -Name '31-DFSNamespaceFolders' -Command {

        foreach ($Root in Get-DfsnRoot -ErrorAction SilentlyContinue) {

            Get-DfsnFolder -Path "$($Root.Path)\*" `
                           -ErrorAction SilentlyContinue |
                Select-Object @{
                                  Name = 'NamespaceRoot'
                                  Expression = { $Root.Path }
                              },
                              Path,
                              State,
                              TimeToLiveSec,
                              Description
        }
    }

    Export-SafeCsv -Name '32-DFSFolderTargets' -Command {

        foreach ($Root in Get-DfsnRoot -ErrorAction SilentlyContinue) {

            foreach ($Folder in Get-DfsnFolder `
                    -Path "$($Root.Path)\*" `
                    -ErrorAction SilentlyContinue) {

                Get-DfsnFolderTarget `
                    -Path $Folder.Path `
                    -ErrorAction SilentlyContinue |
                    Select-Object @{
                                      Name = 'NamespaceRoot'
                                      Expression = { $Root.Path }
                                  },
                                  @{
                                      Name = 'FolderPath'
                                      Expression = { $Folder.Path }
                                  },
                                  TargetPath,
                                  State,
                                  ReferralPriorityClass,
                                  ReferralPriorityRank
            }
        }
    }

    Export-SafeCsv -Name '33-DFSRootTargets' -Command {

        foreach ($Root in Get-DfsnRoot -ErrorAction SilentlyContinue) {

            Get-DfsnRootTarget `
                -Path $Root.Path `
                -ErrorAction SilentlyContinue |
                Select-Object @{
                                  Name = 'RootPath'
                                  Expression = { $Root.Path }
                              },
                              TargetPath,
                              State,
                              ReferralPriorityClass,
                              ReferralPriorityRank
        }
    }
}
else {
    'DFSN PowerShell module not available on this server. Nothing was installed or changed.' |
        Set-Content (Join-Path $OutputPath '30-DFSN-Module-NotAvailable.txt')
}

# ---------------------------------------------------------------------------
# DFS Replication
# ---------------------------------------------------------------------------

$DfsRAvailable = $false

try {
    Import-Module Dfsr -ErrorAction Stop
    $DfsRAvailable = $true
}
catch {
    Write-DiscoveryError -Section 'Import DFSR module' -Exception $_
}

if ($DfsRAvailable) {

    Export-SafeCsv -Name '40-DFSRGroups' -Command {
        Get-DfsReplicationGroup -ErrorAction Stop |
            Select-Object GroupName,
                          DomainName,
                          Description,
                          State
    }

    Export-SafeCsv -Name '41-DFSRMemberships' -Command {

        foreach ($Group in Get-DfsReplicationGroup -ErrorAction SilentlyContinue) {

            Get-DfsrMembership `
                -GroupName $Group.GroupName `
                -ErrorAction SilentlyContinue |
                Select-Object GroupName,
                              ComputerName,
                              FolderName,
                              ContentPath,
                              PrimaryMember,
                              Enabled,
                              State,
                              StagingPath,
                              StagingPathQuotaInMB,
                              ConflictAndDeletedQuotaInMB
        }
    }

    Export-SafeCsv -Name '42-DFSRConnections' -Command {

        foreach ($Group in Get-DfsReplicationGroup -ErrorAction SilentlyContinue) {

            Get-DfsrConnection `
                -GroupName $Group.GroupName `
                -ErrorAction SilentlyContinue |
                Select-Object GroupName,
                              SourceComputerName,
                              DestinationComputerName,
                              Enabled,
                              State
        }
    }

    Export-SafeCsv -Name '43-DFSRReplicatedFolders' -Command {

        foreach ($Group in Get-DfsReplicationGroup -ErrorAction SilentlyContinue) {

            Get-DfsReplicatedFolder `
                -GroupName $Group.GroupName `
                -ErrorAction SilentlyContinue |
                Select-Object GroupName,
                              FolderName,
                              Description,
                              FileNameToExclude,
                              DirectoryNameToExclude,
                              DfsnPath,
                              State
        }
    }

    Export-SafeCsv -Name '44-DFSRState' -Command {

        Get-DfsrState -ComputerName $ComputerName -ErrorAction SilentlyContinue |
            Select-Object FileName,
                          UpdateState,
                          PathName,
                          FullPathName,
                          FileSize,
                          FileId,
                          ReplicatedFolderName,
                          ReplicationGroupName,
                          SourceComputerName,
                          DestinationComputerName
    }
}
else {
    'DFSR PowerShell module not available on this server. Nothing was installed or changed.' |
        Set-Content (Join-Path $OutputPath '40-DFSR-Module-NotAvailable.txt')
}

# ---------------------------------------------------------------------------
# DFSR event log
# ---------------------------------------------------------------------------

Export-SafeCsv -Name '45-DFSRRecentEvents' -Command {

    $StartTime = (Get-Date).AddDays(-30)

    Get-WinEvent -FilterHashtable @{
        LogName   = 'DFS Replication'
        StartTime = $StartTime
    } -ErrorAction Stop |
        Where-Object {
            $_.Level -in 1, 2, 3
        } |
        Select-Object TimeCreated,
                      Id,
                      LevelDisplayName,
                      ProviderName,
                      Message
}

# ---------------------------------------------------------------------------
# Services and scheduled tasks
# ---------------------------------------------------------------------------

Export-SafeCsv -Name '50-Services' -Command {

    Get-CimInstance Win32_Service |
        Select-Object Name,
                      DisplayName,
                      State,
                      StartMode,
                      StartName,
                      PathName
}

Export-SafeCsv -Name '51-ScheduledTasks' -Command {

    Get-ScheduledTask |
        ForEach-Object {

            $Task = $_

            foreach ($Action in $Task.Actions) {

                [pscustomobject]@{
                    TaskPath  = $Task.TaskPath
                    TaskName  = $Task.TaskName
                    State     = $Task.State
                    UserId    = $Task.Principal.UserId
                    Execute   = $Action.Execute
                    Arguments = $Action.Arguments
                    WorkingDirectory = $Action.WorkingDirectory
                }
            }
        }
}

# ---------------------------------------------------------------------------
# Search services/tasks for UNC paths and relevant server names
# ---------------------------------------------------------------------------

$SearchTerms = @(
    '\\',
    'STLDFS',
    'STLFILE2012',
    'STLHR01'
)

Export-SafeCsv -Name '52-PotentialUNCDependencies' -Command {

    $Results = @()

    Get-CimInstance Win32_Service | ForEach-Object {

        $Item = $_

        foreach ($Term in $SearchTerms) {

            if (($Item.PathName -like "*$Term*") -or
                ($Item.StartName -like "*$Term*")) {

                $Results += [pscustomobject]@{
                    SourceType = 'Service'
                    Name       = $Item.Name
                    Location   = ''
                    Detail     = $Item.PathName
                    RunAs      = $Item.StartName
                    Match      = $Term
                }

                break
            }
        }
    }

    Get-ScheduledTask | ForEach-Object {

        $Task = $_

        foreach ($Action in $Task.Actions) {

            $Text = "$($Action.Execute) $($Action.Arguments) $($Action.WorkingDirectory)"

            foreach ($Term in $SearchTerms) {

                if ($Text -like "*$Term*") {

                    $Results += [pscustomobject]@{
                        SourceType = 'ScheduledTask'
                        Name       = $Task.TaskName
                        Location   = $Task.TaskPath
                        Detail     = $Text
                        RunAs      = $Task.Principal.UserId
                        Match      = $Term
                    }

                    break
                }
            }
        }
    }

    $Results
}

# ---------------------------------------------------------------------------
# AD CS detection
# Important for STLFILE2012. Read-only only.
# ---------------------------------------------------------------------------

Export-SafeCsv -Name '60-ADCSDetection' -Command {

    $CAService = Get-CimInstance Win32_Service `
        -Filter "Name='CertSvc'" `
        -ErrorAction SilentlyContinue

    $Feature = $null

    if (Get-Command Get-WindowsFeature -ErrorAction SilentlyContinue) {
        $Feature = Get-WindowsFeature ADCS-Cert-n SilentlyContinue
    }

    [pscustomobject]@{
        ComputerName       = $ComputerName
        CertSvcInstalled   = [bool]$CAService
        CertSvcState       = if ($CAService) { $CAService.State } else { '' }
        CertSvcStartMode   = if ($CAService) { $CAService.StartMode } else { '' }
        ADCSFeaturePresent = if ($Feature) { $Feature.Installed } else { 'Unknown' }
    }
}

# ---------------------------------------------------------------------------
# Network configuration
# ---------------------------------------------------------------------------

Export-SafeCsv -Name '70-IPConfiguration' -Command {

    Get-NetIPConfiguration |
        Select-Object InterfaceAlias,
                      InterfaceDescription,
                      @{
                          Name = 'IPv4Address'
                          Expression = { $_.IPv4Address.IPAddress -join ',' }
                      },
                      @{
                          Name = 'IPv4Gateway'
                          Expression = { $_.IPv4DefaultGateway.NextHop -join ',' }
                      },
                      @{
                          Name = 'DNSServers'
                          Expression = { $_.DNSServer.ServerAddresses -join ',' }
                      }
}

Export-SafeCsv -Name '71-NetAdapters' -Command {

    Get-NetAdapter |
        Select-Object Name,
                      InterfaceDescription,
                      Status,
                      LinkSpeed,
                      MacAddress
}

# ---------------------------------------------------------------------------
# Installed applications
# ---------------------------------------------------------------------------

Export-SafeCsv -Name '80-InstalledApplications' -Command {

    $RegistryPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    foreach ($Path in $RegistryPaths) {

        Get-ItemProperty $Path -ErrorAction SilentlyContinue |
            Where-Object DisplayName |
            Select-Object DisplayName,
                          DisplayVersion,
                          Publisher,
                          InstallDate,
                          InstallLocation
    }
}

# ---------------------------------------------------------------------------
# Server role / retirement summary
# ---------------------------------------------------------------------------

Export-SafeText -Name '90-QuickSummary' -Command {

    $Shares = @(Get-SmbShare |
        Where-Object {
            $_.Path -and
            -not $_.Special
        })

    $Sessions = @(Get-SmbSession -ErrorAction SilentlyContinue)
    $OpenFiles = @(Get-SmbOpenFile -ErrorAction SilentlyContinue)

    $CertSvc = Get-Service CertSvc -ErrorAction SilentlyContinue

    "FILE SERVICES DISCOVERY QUICK SUMMARY"
    "====================================="
    ""
    "Computer:              $ComputerName"
    "Domain:                $((Get-CimInstance Win32_ComputerSystem).Domain)"
    "Date:                  $(Get-Date)"
    ""
    "Non-special SMB shares: $($Shares.Count)"
    "Current SMB sessions:   $($Sessions.Count)"
    "Current SMB open files: $($OpenFiles.Count)"
    ""
    "Certificate Services:"
    if ($CertSvc) {
        "  CertSvc present: YES"
        "  CertSvc status:  $($CertSvc.Status)"
    }
    else {
        "  CertSvc present: NO"
    }
    ""
    "Shares:"
    foreach ($Share in $Shares) {
        "  $($Share.Name) -> $($Share.Path)"
    }
}

# ---------------------------------------------------------------------------
# Create manifest
# ---------------------------------------------------------------------------

$Manifest = [pscustomobject]@{
    Script                 = 'Invoke-FileServicesDiscovery.ps1'
    ComputerName           = $ComputerName
    CollectionDate         = Get-Date
    User                   = "$env:USERDOMAIN\$env:USERNAME"
    PowerShellVersion      = $PSVersionTable.PSVersion.ToString()
    DeepFileScan           = [bool]$DeepFileScan
    DfsNamespaceModule     = $DfsNAvailable
    DfsReplicationModule   = $DfsRAvailable
    Opportunity            = '27042 - File Server Consolidation & DFS'
}

$Manifest |
    Export-Csv -Path (Join-Path $OutputPath '00-Manifest.csv') `
               -NoTypeInformation `
               -Encoding UTF8

# ---------------------------------------------------------------------------
# Compress output
# ---------------------------------------------------------------------------

Write-RunLog "Discovery collection complete."
Write-RunLog "Creating ZIP archive."

try {
    Compress-Archive -Path "$OutputPath\*" `
                     -DestinationPath $ZipPath `
                     -CompressionLevel Optimal `
                     -Force

    Write-RunLog "ZIP created successfully: $ZipPath"
}
catch {
    Write-DiscoveryError -Section 'Compress output' -Exception $_
    Write-Warning "Could not create ZIP. Raw output remains at $OutputPath"
}

Write-Host ''
Write-Host '============================================================' -ForegroundColor Green
Write-Host 'FILE SERVICES DISCOVERY COMPLETE' -ForegroundColor Green
Write-Host '============================================================' -ForegroundColor Green
Write-Host "Server: $ComputerName"
Write-Host "Folder: $OutputPath"

if (Test-Path $ZipPath) {
    Write-Host "ZIP:    $ZipPath" -ForegroundColor Cyan
}

Write-Host ''
Write-Host 'No configuration changes were intentionally performed.' -ForegroundColor Yellow