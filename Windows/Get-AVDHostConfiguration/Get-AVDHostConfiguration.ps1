# PowerShell script to export key configuration settings from the current session host for migration.
# Professional Edition: Includes comprehensive data capture, robust error handling, full transcript logging,
# and a final interactive HTML report with a floating navigation menu and collapsible sections.

#region SCRIPT PARAMETERS & SETUP
param(
    [string]$OutputDir = "C:\AVD_Migration_Export"
)

# Create output directory
if (-Not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir }

# --- New: Define dynamic report filename ---
$ComputerName = $env:COMPUTERNAME
$DateStamp = Get-Date -Format "yyyy-MM-dd"
$ReportFileName = "AVD Migration Report - $ComputerName - $DateStamp.html"

# Start logging all console output to a file
Start-Transcript -Path "$OutputDir\_ScriptRun.log" -Append

Write-Host "Exporting configuration data to $OutputDir..." -ForegroundColor Cyan
#endregion

#region DATA EXPORT (INDIVIDUAL FILES)

# --- SYSTEM & OS INFORMATION ---
Write-Host "- Exporting System, OS, and Time Zone info..."
Get-ComputerInfo | Out-File -FilePath "$OutputDir\SystemInfo.txt"
Get-TimeZone | Out-File -FilePath "$OutputDir\TimeZone.txt"
Get-Culture | Out-File -FilePath "$OutputDir\SystemLocale.txt"
Get-WindowsOptionalFeature -Online | Where-Object {$_.State -eq 'Enabled'} | Select-Object FeatureName, State | Export-Csv -Path "$OutputDir\WindowsOptionalFeatures.csv" -NoTypeInformation
powercfg /getactivescheme | Out-File -FilePath "$OutputDir\PowerPlan.txt"
Get-Content -Path "C:\Windows\System32\drivers\etc\hosts" | Out-File -FilePath "$OutputDir\HostsFile.txt"

# --- AZURE & DOMAIN JOIN STATUS ---
Write-Host "- Exporting Azure AD & Domain Join status..."
dsregcmd /status | Out-File -FilePath "$OutputDir\AzureAD_JoinStatus.txt"

# --- NETWORKING & FIREWALL ---
Write-Host "- Exporting Network and Firewall configuration..."
Get-NetIPConfiguration | Out-File -FilePath "$OutputDir\Network_IPConfiguration.txt"
Get-DnsClientServerAddress -AddressFamily IPv4 | Out-File -FilePath "$OutputDir\Network_DnsServers.txt"
Get-SmbClientConfiguration | Select-Object * | Out-File -FilePath "$OutputDir\Network_SMBClientConfig.txt"
Get-NetFirewallRule | Where-Object {$_.Enabled -eq 'True'} | Select-Object DisplayName, DisplayGroup, Direction, Action, Profile, Enabled | Sort-Object DisplayGroup, DisplayName | Export-Csv -Path "$OutputDir\Security_FirewallRules_Enabled.csv" -NoTypeInformation
Get-NetFirewallProfile | Out-File -FilePath "$OutputDir\Security_FirewallProfiles.txt"

# --- INSTALLED SOFTWARE & APPLICATIONS ---
Write-Host "- Exporting Installed Software list..."
Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*, HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |
    Select-Object DisplayName, DisplayVersion, Publisher, InstallDate, InstallLocation |
    Sort-Object DisplayName |
    Where-Object { $_.DisplayName -ne $null } |
    Export-Csv -Path "$OutputDir\InstalledSoftware.csv" -NoTypeInformation

# --- FSLogix, AVD, & APPLICATION CONFIGURATION (REGISTRY) ---
Write-Host "- Exporting key registry settings..."
$registryPathsToExport = @{
    "FSLogix"           = 'HKEY_LOCAL_MACHINE\SOFTWARE\FSLogix';
    "AVD_RDInfraAgent"  = 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\RDInfraAgent';
    "TerminalServer"    = 'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server';
    "ODBC_Connections"  = 'HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBC.INI';
    "ODBC_Connections_Wow6432Node" = 'HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\ODBC\ODBC.INI'
}
foreach ($name in $registryPathsToExport.Keys) {
    $regPath = $registryPathsToExport[$name]
    reg query $regPath /ve >$null 2>&1
    if ($LASTEXITCODE -eq 0) {
        $exportFile = "$OutputDir\RegistryExport_$name.reg"
        reg export $regPath $exportFile /y | Out-Null
        Write-Host "  Successfully exported $regPath"
    } else {
        Write-Warning "Registry path not found or accessible: $regPath"
    }
}

# --- SERVICES & TASKS ---
Write-Host "- Exporting Services and Scheduled Tasks..."
Get-Service | Select-Object Name, DisplayName, Status, StartType, PathName | Export-Csv -Path "$OutputDir\Services_All.csv" -NoTypeInformation
Get-ScheduledTask | Select-Object TaskName, TaskPath, State, @{Name="Actions";Expression={$_.Actions.Execute -join '; '}} | Export-Csv -Path "$OutputDir\ScheduledTasks_All.csv" -NoTypeInformation

# --- USER, GROUP, & CERTIFICATE CONFIGURATION ---
Write-Host "- Exporting Local Users, Groups, and Certificates..."
Get-LocalUser | Select-Object Name, Enabled, FullName, Description | Export-Csv -Path "$OutputDir\Security_LocalUsers.csv" -NoTypeInformation
(Get-LocalGroupMember -Group "Administrators").Name | Out-File -FilePath "$OutputDir\Security_LocalAdmins.txt"
Get-ChildItem -Path Cert:\LocalMachine\My | Select-Object Subject, Issuer, NotAfter, Thumbprint | Export-Csv -Path "$OutputDir\Security_ComputerCertificates.csv" -NoTypeInformation

# --- PRINTERS ---
Write-Host "- Exporting Printer configurations..."
Get-Printer | Select-Object Name, DriverName, PortName, Shared, Published | Out-File -FilePath "$OutputDir\Printers_Installed.txt"
Get-PrinterDriver | Select-Object Name, Manufacturer, MajorVersion | Export-Csv -Path "$OutputDir\Printers_Drivers.csv" -NoTypeInformation
Get-PrinterPort | Out-File -FilePath "$OutputDir\Printers_Ports.txt"

# --- GPO & EXECUTION POLICY ---
Write-Host "- Exporting applied GPOs and PowerShell Execution Policy..."
try {
    Get-GPResultantSetOfPolicy -ReportType Html -Path "$OutputDir\GPResult.html" -ErrorAction Stop
} catch {
    Write-Warning "Get-GPResultantSetOfPolicy failed. Trying alternative 'gpresult.exe' method..."
    gpresult /h "$OutputDir\GPResult.html" /f
}
Get-ExecutionPolicy -List | Out-File -FilePath "$OutputDir\ExecutionPolicy.txt"

# --- STORAGE, TIME, & MISC ---
Write-Host "- Exporting Storage, Time, and Misc info..."
Get-Disk | Out-File -FilePath "$OutputDir\Storage_DiskInfo.txt"
Get-Partition | Out-File -FilePath "$OutputDir\Storage_Partitions.txt"
Get-Volume | Out-File -FilePath "$OutputDir\Storage_Volumes.txt"
w32tm /query /configuration | Out-File -FilePath "$OutputDir\TimeConfiguration.txt"

Write-Host "All individual data files have been exported."
#endregion

#region HTML REPORT GENERATION
Write-Host "Generating comprehensive HTML summary report..."

# --- Helper Functions for HTML Report ---
function Convert-CsvToHtmlTable { param($FilePath) ; if (-not (Test-Path $FilePath)) { return "<p><i>Data file not found: $FilePath</i></p>" } ; $data = Import-Csv -Path $FilePath ; if (-not $data) { return "<p><i>No data found in $FilePath</i></p>" } ; $sbTable = [System.Text.StringBuilder]::new() ; [void]$sbTable.Append("<table>`n<tr>") ; foreach ($header in $data[0].PSObject.Properties.Name) { [void]$sbTable.Append("<th>$header</th>") } ; [void]$sbTable.Append("</tr>`n") ; foreach ($row in $data) { [void]$sbTable.Append("<tr>") ; foreach ($prop in $row.PSObject.Properties) { $value = "$($prop.Value)".Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;') ; [void]$sbTable.Append("<td>$value</td>") } ; [void]$sbTable.Append("</tr>`n") } ; [void]$sbTable.Append("</table>") ; return $sbTable.ToString() }
function Add-ContentToReport { param($StringBuilder, $Title, $FilePath, $FileType = 'pre', $SubheadingLevel = 3) ; [void]$StringBuilder.Append("<h$SubheadingLevel>$Title</h$SubheadingLevel>") ; [void]$StringBuilder.Append("<div class='content-block'>") ; if (Test-Path $FilePath) { switch ($FileType) { 'pre' { $content = Get-Content -Raw -Path $FilePath ; $encodedContent = $content.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;') ; [void]$StringBuilder.Append("<pre>$encodedContent</pre>") } 'csv' { $htmlTable = Convert-CsvToHtmlTable -FilePath $FilePath ; [void]$StringBuilder.Append($htmlTable) } 'link' { $fileName = Split-Path -Path $FilePath -Leaf ; [void]$StringBuilder.Append("<p>Data for this section is in a separate file. <a href='$fileName' target='_blank'>Click here to open $fileName</a></p>") } } } else { [void]$StringBuilder.Append("<p><i>Data file not found or not generated: $FilePath</i></p>") } ; [void]$StringBuilder.Append("</div>") }

# --- Main HTML Structure & Content ---
$sb = [System.Text.StringBuilder]::new()
$reportSections = @{
    "section-summary"  = "Execution Summary"
    "section-system"   = "System & OS"
    "section-identity" = "Identity & Security"
    "section-software" = "Installed Software & Services"
    "section-avd"      = "AVD, FSLogix & App Config"
    "section-network"  = "Networking & Firewall"
    "section-automation" = "Automation & Peripherals"
    "section-storage"  = "Storage"
}
$htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
<title>AVD Migration Report for $($env:COMPUTERNAME)</title>
<style>
    body { font-family: 'Segoe UI', Arial, sans-serif; font-size: 14px; margin: 0; padding: 0; background-color: #fcfcfc; color: #333; }
    h1 { color: #004080; text-align: center; }
    h2 { color: #0059b3; background-color: #e6f2ff; padding: 10px; border-left: 5px solid #004080; margin-top: 40px; cursor: pointer; }
    h2::after { content: ' ▼'; font-size: smaller; }
    h2.collapsed::after { content: ' ►'; }
    h3 { color: #34495E; border-bottom: 1px solid #D5DBDB; padding-bottom: 5px; margin-top: 25px; }
    table { border-collapse: collapse; width: 95%; margin-top: 15px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
    th, td { border: 1px solid #ccc; padding: 10px; text-align: left; vertical-align: top; }
    th { background-color: #e6f2ff; font-weight: bold; }
    tr:nth-child(even) { background-color: #f7faff; }
    pre { background-color: #f0f0f0; padding: 15px; border: 1px solid #ccc; white-space: pre-wrap; word-wrap: break-word; font-family: Consolas, monospace; }
    .toc { position: fixed; top: 20px; left: 20px; width: 220px; background: #fff; border: 1px solid #ccc; border-radius: 5px; padding: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.15); z-index: 1000; }
    .toc h3 { margin-top: 0; color: #004080; font-size: 16px; border-bottom: 1px solid #ddd; padding-bottom: 8px; }
    .toc ul { list-style-type: none; padding: 0; margin: 10px 0 0 0; }
    .toc ul li a { text-decoration: none; color: #0059b3; display: block; padding: 6px 10px; border-radius: 4px; }
    .toc ul li a:hover { text-decoration: underline; background-color: #e6f2ff; }
    .main-content { margin-left: 260px; padding: 20px; }
    .report-header { text-align: center; margin-bottom: 30px; border-bottom: 2px solid #0059b3; padding-bottom: 20px; }
    .content-block { margin-left: 20px; padding-bottom: 10px; }
    .section-content { display: block; } /* Change to none to collapse by default */
</style>
<script>
    function toggleSection(header) {
        var content = header.nextElementSibling;
        if (content.style.display === "none") {
            content.style.display = "block";
            header.classList.remove("collapsed");
        } else {
            content.style.display = "none";
            header.classList.add("collapsed");
        }
    }
</script>
</head>
<body>
"@
[void]$sb.Append($htmlHeader)
[void]$sb.Append("<div class='toc'><h3>Report Sections</h3><ul>")
foreach($id in $reportSections.Keys) { $title = $reportSections[$id]; [void]$sb.Append("<li><a href='#$id'>$title</a></li>") }
[void]$sb.Append("</ul></div>")
[void]$sb.Append("<div class='main-content'>")
[void]$sb.Append("<div class='report-header'><h1>AVD Host Configuration Report</h1><p><strong>Computer Name:</strong> $($env:COMPUTERNAME)</p><p><em>Report Generated On: $(Get-Date)</em></p></div>")

# --- Build HTML Body ---
[void]$sb.Append("<h2 id='section-summary' onclick='toggleSection(this)'>$($reportSections['section-summary'])</h2><div class='section-content'>")
Add-ContentToReport -StringBuilder $sb -Title "Execution Summary" -FilePath "$OutputDir\_ScriptRun.log"
[void]$sb.Append("</div>")

[void]$sb.Append("<h2 id='section-system' onclick='toggleSection(this)'>$($reportSections['section-system'])</h2><div class='section-content'>")
Add-ContentToReport -StringBuilder $sb -Title "System Details (Get-ComputerInfo)" -FilePath "$OutputDir\SystemInfo.txt"
Add-ContentToReport -StringBuilder $sb -Title "Time Zone" -FilePath "$OutputDir\TimeZone.txt"
Add-ContentToReport -StringBuilder $sb -Title "System Locale" -FilePath "$OutputDir\SystemLocale.txt"
Add-ContentToReport -StringBuilder $sb -Title "Enabled Windows Optional Features" -FilePath "$OutputDir\WindowsOptionalFeatures.csv" -FileType 'csv'
Add-ContentToReport -StringBuilder $sb -Title "Active Power Plan" -FilePath "$OutputDir\PowerPlan.txt"
Add-ContentToReport -StringBuilder $sb -Title "Hosts File Content" -FilePath "$OutputDir\HostsFile.txt"
[void]$sb.Append("</div>")

[void]$sb.Append("<h2 id='section-identity' onclick='toggleSection(this)'>$($reportSections['section-identity'])</h2><div class='section-content'>")
Add-ContentToReport -StringBuilder $sb -Title "Azure AD & Domain Join Status" -FilePath "$OutputDir\AzureAD_JoinStatus.txt"
Add-ContentToReport -StringBuilder $sb -Title "Applied Group Policies (GPO)" -FilePath "$OutputDir\GPResult.html" -FileType 'link'
Add-ContentToReport -StringBuilder $sb -Title "Local Users" -FilePath "$OutputDir\Security_LocalUsers.csv" -FileType 'csv'
Add-ContentToReport -StringBuilder $sb -Title "Local Administrators Group Members" -FilePath "$OutputDir\Security_LocalAdmins.txt"
Add-ContentToReport -StringBuilder $sb -Title "PowerShell Execution Policy" -FilePath "$OutputDir\ExecutionPolicy.txt"
Add-ContentToReport -StringBuilder $sb -Title "Installed Computer Certificates (Personal Store)" -FilePath "$OutputDir\Security_ComputerCertificates.csv" -FileType 'csv'
[void]$sb.Append("</div>")

[void]$sb.Append("<h2 id='section-software' onclick='toggleSection(this)'>$($reportSections['section-software'])</h2><div class='section-content'>")
Add-ContentToReport -StringBuilder $sb -Title "Installed Applications" -FilePath "$OutputDir\InstalledSoftware.csv" -FileType 'csv'
Add-ContentToReport -StringBuilder $sb -Title "All Services" -FilePath "$OutputDir\Services_All.csv" -FileType 'csv'
[void]$sb.Append("</div>")

[void]$sb.Append("<h2 id='section-avd' onclick='toggleSection(this)'>$($reportSections['section-avd'])</h2><div class='section-content'>")
Add-ContentToReport -StringBuilder $sb -Title "FSLogix Registry Settings" -FilePath "$OutputDir\RegistryExport_FSLogix.reg" -FileType 'link'
Add-ContentToReport -StringBuilder $sb -Title "AVD Agent (RDInfraAgent) Registry Settings" -FilePath "$OutputDir\RegistryExport_AVD_RDInfraAgent.reg" -FileType 'link'
Add-ContentToReport -StringBuilder $sb -Title "Terminal Server Registry Settings" -FilePath "$OutputDir\RegistryExport_TerminalServer.reg" -FileType 'link'
Add-ContentToReport -StringBuilder $sb -Title "ODBC DSNs (64-bit)" -FilePath "$OutputDir\RegistryExport_ODBC_Connections.reg" -FileType 'link'
Add-ContentToReport -StringBuilder $sb -Title "ODBC DSNs (32-bit)" -FilePath "$OutputDir\RegistryExport_ODBC_Connections_Wow6432Node.reg" -FileType 'link'
[void]$sb.Append("</div>")

[void]$sb.Append("<h2 id='section-network' onclick='toggleSection(this)'>$($reportSections['section-network'])</h2><div class='section-content'>")
Add-ContentToReport -StringBuilder $sb -Title "Network IP Configuration" -FilePath "$OutputDir\Network_IPConfiguration.txt"
Add-ContentToReport -StringBuilder $sb -Title "DNS Servers" -FilePath "$OutputDir\Network_DnsServers.txt"
Add-ContentToReport -StringBuilder $sb -Title "SMB Client Configuration" -FilePath "$OutputDir\Network_SMBClientConfig.txt"
Add-ContentToReport -StringBuilder $sb -Title "Enabled Firewall Rules" -FilePath "$OutputDir\Security_FirewallRules_Enabled.csv" -FileType 'csv'
Add-ContentToReport -StringBuilder $sb -Title "Firewall Profiles" -FilePath "$OutputDir\Security_FirewallProfiles.txt"
Add-ContentToReport -StringBuilder $sb -Title "Time (NTP) Configuration" -FilePath "$OutputDir\TimeConfiguration.txt"
[void]$sb.Append("</div>")

[void]$sb.Append("<h2 id='section-automation' onclick='toggleSection(this)'>$($reportSections['section-automation'])</h2><div class='section-content'>")
Add-ContentToReport -StringBuilder $sb -Title "All Scheduled Tasks" -FilePath "$OutputDir\ScheduledTasks_All.csv" -FileType 'csv'
Add-ContentToReport -StringBuilder $sb -Title "Installed Printers" -FilePath "$OutputDir\Printers_Installed.txt"
Add-ContentToReport -StringBuilder $sb -Title "Printer Drivers" -FilePath "$OutputDir\Printers_Drivers.csv" -FileType 'csv'
Add-ContentToReport -StringBuilder $sb -Title "Printer Ports" -FilePath "$OutputDir\Printers_Ports.txt"
[void]$sb.Append("</div>")

[void]$sb.Append("<h2 id='section-storage' onclick='toggleSection(this)'>$($reportSections['section-storage'])</h2><div class='section-content'>")
Add-ContentToReport -StringBuilder $sb -Title "Disk Information" -FilePath "$OutputDir\Storage_DiskInfo.txt"
Add-ContentToReport -StringBuilder $sb -Title "Partitions" -FilePath "$OutputDir\Storage_Partitions.txt"
Add-ContentToReport -StringBuilder $sb -Title "Volumes" -FilePath "$OutputDir\Storage_Volumes.txt"
[void]$sb.Append("</div>")

# --- Finalize HTML ---
[void]$sb.Append("</div></body></html>")
$sb.ToString() | Out-File -FilePath "$OutputDir\$ReportFileName" -Encoding UTF8

Write-Host "`nHTML summary report created successfully!" -ForegroundColor Green
Write-Host "All files are located in: $OutputDir"
Write-Host "Report file is: $ReportFileName"

Stop-Transcript
#endregion