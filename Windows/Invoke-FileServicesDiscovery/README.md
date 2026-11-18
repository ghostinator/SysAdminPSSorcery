# Windows File Services Discovery

A read\-only PowerShell discovery and assessment tool for Windows file servers.

Invoke\-FileServicesDiscovery.ps1 inventories SMB shares, NTFS permissions, DFS Namespaces, DFS Replication, storage, active SMB usage, scheduled tasks, services, file characteristics, and other dependencies that matter when migrating, consolidating, modernizing, or retiring Windows file servers.

The script collects its findings into structured CSV and text files, then packages the results into a ZIP for offline analysis.

## Use Cases

This script is useful for:

- File server migration discovery
- File server consolidation projects
- DFS Namespace assessment
- DFS\-R assessment and troubleshooting
- Legacy file server retirement
- SMB dependency discovery
- Share and NTFS permission review
- SharePoint Online migration assessment
- OneDrive migration assessment
- Azure Files migration assessment
- Windows Server modernization projects
- Pre\-migration documentation
- Technical discovery and project scoping

## Safety

The script is designed to be **read\-only**.

It does not intentionally:

- Create, modify, or delete SMB shares
- Modify NTFS permissions
- Modify DFS Namespaces
- Modify DFS Replication
- Stop or restart services
- Change scheduled tasks
- Change Windows features
- Install PowerShell modules
- Move, rename, modify, or delete user data
- Change Windows Server configuration

The script primarily uses standard Windows PowerShell, CIM, SMB, storage, DFS, registry, and filesystem queries.

**Important:** Always review scripts before running them in a production environment. Test in an appropriate environment when required by your organization's change\-management or security procedures.

## Requirements

Recommended:

- Windows Server
- Windows PowerShell 5.1
- Local Administrator privileges
- Run from an elevated PowerShell session

Certain reports depend on Windows features and PowerShell modules already being present.

DFS Namespace discovery requires the Dfsn module.

DFS Replication discovery requires the Dfsr module.

If these modules are unavailable, the script records that condition and continues. It does not install them.

## Basic Usage

Run from an elevated Windows PowerShell prompt:

.\\Invoke\-FileServicesDiscovery.ps1

If PowerShell execution policy prevents execution, you can temporarily change the policy for the current PowerShell process:

Set\-ExecutionPolicy Bypass \-Scope Process \-Force

Then run:

.\\Invoke\-FileServicesDiscovery.ps1

The execution policy change applies only to the current PowerShell process.

## Deep File Scan

An optional deeper filesystem assessment can be enabled with:

.\\Invoke\-FileServicesDiscovery.ps1 \-DeepFileScan

The deep scan adds additional analysis for items such as:

- Long filesystem paths
- Potential cloud\-migration filename issues

For large file servers, run the standard discovery first unless you specifically need these reports.

## Custom Output Location

The default output location is:

C:\\Temp

A different location can be specified:

.\\Invoke\-FileServicesDiscovery.ps1 \-OutputRoot D:\\Discovery

## Output

The script creates a timestamped directory similar to:

FileServicesDiscovery\-FILESERVER01\-20261118\-105500

It also creates a ZIP archive:

FileServicesDiscovery\-FILESERVER01\-20261118\-105500.zip

This makes it easy to collect discovery results from multiple servers and analyze them centrally.

## Data Collected

### System Information

Collects information including:

- Computer name
- Active Directory domain
- Windows Server version
- Build number
- Server manufacturer and model
- Memory
- Processor information
- Installed Windows features
- Disk configuration
- Volume utilization

### SMB Configuration

Collects:

- SMB shares
- Share paths
- Share descriptions
- Share\-level permissions
- SMB server configuration
- SMB client configuration
- Current SMB sessions
- Currently open SMB files

Active sessions and open files can be especially useful when identifying whether supposedly legacy shares are still in use.

### NTFS Permissions

Collects NTFS permissions from each SMB share root, including:

- Owner
- Security principal
- Filesystem rights
- Allow or deny
- Inheritance status
- Inheritance flags
- Propagation flags

This helps identify permission structures that need to be preserved, cleaned up, or redesigned during a migration.

## Share Data Analysis

For each non\-administrative SMB share, the script can collect:

- File count
- Directory count
- Total data size
- Oldest modification date
- Most recent modification date
- File\-extension distribution

This can help distinguish actively used production data from stale or archival data.

## Application and Migration\-Sensitive Files

The script identifies files with extensions commonly worth reviewing before migration, including examples such as:

- PST and OST
- Access databases
- SQL Server MDF, LDF, and NDF
- QuickBooks files
- VHD and VHDX
- Backup files
- ISO images
- Executables and installers
- CAD\-related files

These files are not automatically considered unsafe or incompatible.

They are reported because they may require additional assessment before moving data to platforms such as SharePoint Online, OneDrive, or Azure Files.

## EFS Discovery

The script uses the built\-in Windows cipher.exe utility to look for Encrypting File System usage within shared data.

EFS discovery is particularly important before:

- File server retirement
- Data migration
- Certificate Authority retirement
- Recovery\-agent changes
- User\-profile retirement

Encrypted files should be understood before the certificates and private keys needed to decrypt them become unavailable.

## DFS Namespace Discovery

When the DFS Namespace PowerShell module is available, the script collects:

- DFS Namespace roots
- Root targets
- DFS folders
- Folder targets
- Target state
- Referral configuration

This can help identify stale namespace entries, old file servers, and dependencies that should be corrected before retiring infrastructure.

## DFS Replication Discovery

When the DFS Replication PowerShell module is available, the script collects:

- Replication groups
- Replicated folders
- Memberships
- Connections
- Content paths
- Staging configuration
- Member state
- Current DFS\-R activity

It also collects recent warning and error events from the **DFS Replication** event log.

## Workload Dependency Discovery

The script inventories:

- Windows services
- Service accounts
- Service executable paths
- Scheduled tasks
- Scheduled task commands
- Scheduled task arguments
- Scheduled task working directories

It searches this information for UNC paths and known file\-server references.

This can identify situations where a server appears safe to retire but is still referenced by an application, scheduled process, or Windows service.

## Certificate Services Detection

The script checks whether Active Directory Certificate Services appears to be installed on the server and whether the CertSvc service is present.

This does **not** perform a full AD CS audit.

The check exists because older Windows file servers sometimes host additional infrastructure roles that are easy to overlook during a file\-server migration.

Never assume a server can be retired solely because its SMB shares have been migrated.

## Network Information

The script collects:

- Network adapters
- Link speeds
- MAC addresses
- IPv4 addresses
- Default gateways
- DNS servers

This provides additional context when assessing legacy infrastructure.

## Installed Applications

Installed applications are inventoried from the standard 32\-bit and 64\-bit Windows uninstall registry locations.

This can help identify third\-party software that may be tied to the server being assessed.

## Output Files

Depending on server configuration and available modules, output can include files similar to:

- 00\-Manifest.csv
- 01\-ComputerSystem.csv
- 02\-OperatingSystem.csv
- 03\-Volumes.csv
- 04\-Disks.csv
- 05\-WindowsFeatures.csv
- 10\-SMBShares.csv
- 11\-SMBSharePermissions.csv
- 12\-SMBSessions.csv
- 13\-SMBOpenFiles.csv
- 14\-SMBServerConfiguration.txt
- 15\-SMBClientConfiguration.txt
- 20\-NTFSShareRootPermissions.csv
- 21\-ShareDataSummary.csv
- 22\-FileExtensionSummary.csv
- 23\-ProblemFileTypes.csv
- 24\-EFSEncryptedFiles.txt
- 30\-DFSNamespaceRoots.csv
- 31\-DFSNamespaceFolders.csv
- 32\-DFSFolderTargets.csv
- 33\-DFSRootTargets.csv
- 40\-DFSRGroups.csv
- 41\-DFSRMemberships.csv
- 42\-DFSRConnections.csv
- 43\-DFSRReplicatedFolders.csv
- 44\-DFSRState.csv
- 45\-DFSRRecentEvents.csv
- 50\-Services.csv
- 51\-ScheduledTasks.csv
- 52\-PotentialUNCDependencies.csv
- 60\-ADCSDetection.csv
- 70\-IPConfiguration.csv
- 71\-NetAdapters.csv
- 80\-InstalledApplications.csv
- 90\-QuickSummary.txt
- Errors.txt
- RunLog.txt

Additional files are generated when \-DeepFileScan is specified.

## Performance Considerations

Filesystem enumeration can be expensive.

Reports that recursively examine files may take significant time on servers containing:

- Millions of files
- Multiple terabytes of data
- Deep directory structures
- Slow storage
- Offline files
- Broken mount points
- Highly fragmented storage
- Antivirus products performing synchronous scanning

The script uses read\-only operations, but read operations still generate storage I/O.

For large production file servers, consider running discovery during a lower\-utilization period.

Start without \-DeepFileScan.

## Multi\-Server Assessments

For consolidation projects, run the script separately on every relevant file server.

For example:

- Existing source file servers
- DFS Namespace servers
- DFS\-R members
- Intended consolidation targets
- Legacy application servers that also expose SMB shares

The resulting ZIP archives can then be compared to build a migration matrix identifying:

- Current share
- Current server
- Data size
- Permissions
- DFS path
- DFS target
- Application dependencies
- Current activity
- Proposed destination
- Migration status
- Retirement dependency

## Limitations

This script provides technical discovery data. It does not determine business ownership or automatically decide where data should migrate.

A share containing ordinary Office documents does not automatically belong in SharePoint Online.

A share using SMB does not automatically require Azure Files.

Migration decisions should also consider:

- Application requirements
- Business ownership
- Workflow
- Identity requirements
- Network connectivity
- File locking behavior
- Performance
- Permissions
- Compliance
- Retention requirements
- Backup and recovery requirements

The script should be treated as an evidence\-gathering tool, not an automated migration decision engine.

## Troubleshooting

Review:

Errors.txt

for collection failures.

Review:

RunLog.txt

for the sequence of discovery operations.

Missing DFS reports usually indicate the corresponding DFS PowerShell module is not available on the server.

A missing or empty report does not necessarily mean the associated technology is unused. Validate unexpected results before making infrastructure changes.

## Security and Privacy

The output can contain sensitive infrastructure information, including:

- Server names
- Domain names
- IP addresses
- Share names and paths
- User and group names
- Permissions
- Service accounts
- Application names
- Scheduled task commands
- Active SMB client information
- File paths and filenames

**Treat discovery archives as sensitive administrative data.**

Do not publish real discovery output to a public repository.

Review output before sharing it with third parties.

## Contributing

Issues, improvements, compatibility fixes, and pull requests are welcome.

When reporting a problem, include:

- Windows Server version
- PowerShell version
- Relevant Windows roles
- Whether DFS Namespace or DFS Replication is installed
- The applicable error from Errors.txt

Please sanitize company names, domains, usernames, server names, IP addresses, file paths, and other sensitive information before posting diagnostic output.

## Disclaimer

This project is provided as\-is.

Review and test the script before using it in production. The author and contributors are not responsible for changes in Microsoft Windows behavior, environmental differences, unexpected workload impact, or decisions made based on collected discovery data.

