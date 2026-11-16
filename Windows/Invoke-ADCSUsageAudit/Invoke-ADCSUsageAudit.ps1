#requires -version 5.1
<#+
.SYNOPSIS
Performs a read-only evidence-based audit of a Windows AD CS Certification Authority.
.DESCRIPTION
Collects CA role, service, configuration, issuance, template, AD publication, local certificate,
IIS, NDES, OCSP, event, and file-share evidence. Produces JSON, CSV, text, and HTML reports and
classifies the CA as Active, Likely Active, Historical Only, No Issuance Evidence, Inconclusive,
or Not a CA. It never changes AD CS configuration.
#>
[CmdletBinding()]
param(
    [ValidateRange(1,3650)][int]$RecentDays = 400,
    [ValidateRange(1,1000000)][int]$MaxIssuedRows = 50000,
    [string]$OutputPath = (Join-Path $env:ProgramData ("ADCSUsageAudit-{0:yyyyMMdd-HHmmss}" -f (Get-Date))),
    [switch]$IncludeEventLogs,
    [ValidateRange(1,3650)][int]$EventLookbackDays = 90,
    [switch]$SkipActiveDirectory
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Add-Finding {
    param([string]$Category,[string]$Severity,[string]$Title,[string]$Detail,[int]$Weight=0)
    $script:Findings.Add([pscustomobject]@{Category=$Category;Severity=$Severity;Title=$Title;Detail=$Detail;Weight=$Weight})
}
function Invoke-Native {
    param([string]$FilePath,[string[]]$Arguments)
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName=$FilePath; $psi.UseShellExecute=$false; $psi.RedirectStandardOutput=$true; $psi.RedirectStandardError=$true
    $psi.CreateNoWindow=$true; $psi.Arguments=($Arguments | ForEach-Object { if($_ -match '[\s,"]'){ '"'+($_ -replace '"','\"')+'"' } else { $_ } }) -join ' '
    $p=New-Object System.Diagnostics.Process; $p.StartInfo=$psi; [void]$p.Start(); $out=$p.StandardOutput.ReadToEnd(); $err=$p.StandardError.ReadToEnd(); $p.WaitForExit()
    [pscustomobject]@{ExitCode=$p.ExitCode;StdOut=$out;StdErr=$err;Command="$FilePath $($psi.Arguments)"}
}
function Convert-CertutilCsv {
    param([string]$Text)
    $lines = $Text -split "`r?`n"
    $start = -1
    for($i=0;$i -lt $lines.Count;$i++) { if($lines[$i] -match '^".*",".*"'){ $start=$i; break } }
    if($start -lt 0){ return @() }
    $csv = ($lines[$start..($lines.Count-1)] | Where-Object { $_ -notmatch '^CertUtil:' -and $_ -notmatch '^\s*$' }) -join "`n"
    try { return @($csv | ConvertFrom-Csv) } catch { Add-Finding 'Issuance' 'Warning' 'CSV parse failed' $_.Exception.Message 0; return @() }
}
function Get-FirstPropertyValue {
    param($Object,[string[]]$Names)
    foreach($n in $Names){ $p=$Object.PSObject.Properties[$n]; if($p){ return $p.Value } }
    return $null
}
function Convert-ToDateSafe { param($Value) if(!$Value){return $null}; $d=[datetime]::MinValue; if([datetime]::TryParse([string]$Value,[ref]$d)){return $d}; return $null }
function HtmlEncode { param([object]$Value) [System.Net.WebUtility]::HtmlEncode([string]$Value) }

$Findings = New-Object 'System.Collections.Generic.List[object]'
$Evidence = [ordered]@{}
$stage = 'Initialization'
trap { Write-Error ("ADCSUsageAudit failed during stage '{0}'. Line {1}: {2}" -f $stage, $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message); break }
New-Item -ItemType Directory -Force -Path $OutputPath | Out-Null
$Evidence.Run = [ordered]@{ComputerName=$env:COMPUTERNAME;Timestamp=(Get-Date);RecentDays=$RecentDays;User=$env:USERNAME;PowerShell=$PSVersionTable.PSVersion.ToString()}

# Role and service
$stage = 'Role and service discovery'
try {
    Import-Module ServerManager -ErrorAction Stop
    $features = Get-WindowsFeature AD-Certificate,ADCS-Cert-Authority,ADCS-Web-Enrollment,ADCS-Enroll-Web-Svc,ADCS-Enroll-Web-Pol,ADCS-Device-Enrollment,ADCS-Online-Cert
    $Evidence.Features = @($features | Select-Object Name,DisplayName,InstallState)
} catch { $Evidence.Features=@(); Add-Finding 'Platform' 'Warning' 'ServerManager unavailable' $_.Exception.Message 0 }
$svc = Get-Service CertSvc -ErrorAction SilentlyContinue
$Evidence.CertSvc = $null
if($null -ne $svc){ $Evidence.CertSvc = $svc | Select-Object Name,Status,StartType }
if(!$svc){ Add-Finding 'Platform' 'Info' 'Certification Authority service absent' 'CertSvc was not found.' 0 }
else { Add-Finding 'Platform' 'Info' 'Certification Authority service present' "CertSvc status is $($svc.Status), startup type $($svc.StartType)." 1 }

# CA configuration
$stage = 'CA configuration'
$caConfig = Invoke-Native certutil.exe @('-getreg','CA')
$Evidence.CARegistryRaw=$caConfig.StdOut
$caName = if($caConfig.StdOut -match 'CommonName\s+REG_SZ\s+=\s+(.+)'){ $Matches[1].Trim() } else { $null }
$caType = if($caConfig.StdOut -match 'ENUM_ENTERPRISE_ROOTCA'){ 'Enterprise Root CA' } elseif($caConfig.StdOut -match 'ENUM_ENTERPRISE_SUBCA'){ 'Enterprise Subordinate CA' } elseif($caConfig.StdOut -match 'ENUM_STANDALONE_ROOTCA'){ 'Standalone Root CA' } elseif($caConfig.StdOut -match 'ENUM_STANDALONE_SUBCA'){ 'Standalone Subordinate CA' } else { 'Unknown' }
$Evidence.CA=[ordered]@{Name=$caName;Type=$caType;GetRegExitCode=$caConfig.ExitCode}

# Published templates
$stage = 'Published template discovery'
$tplResult=Invoke-Native certutil.exe @('-catemplates')
$tplLines=@($tplResult.StdOut -split "`r?`n" | Where-Object { $_ -match '^\S[^:]+:' -and $_ -notmatch '^[0-9]+\.' -and $_ -notmatch '^CertUtil:' })
$templates=@($tplLines | ForEach-Object { $parts=$_ -split ':',2; [pscustomobject]@{Template=$parts[0].Trim();Description=$parts[1].Trim()} })
$Evidence.PublishedTemplates=$templates
if($templates.Count -gt 0){ Add-Finding 'Templates' 'Attention' 'Templates are published' "$($templates.Count) template entries were returned by this CA. Published templates show enrollment capability, not proof of current use." 2 }

# Issuance database. Keep output bounded and read-only.
$stage = 'CA issuance database query'
$fields='RequestID,RequesterName,CertificateTemplate,NotBefore,NotAfter,CommonName,Disposition'
$issuedResult=Invoke-Native certutil.exe @('-view','-restrict','Disposition=20','-out',$fields,'csv')
$issued=@(Convert-CertutilCsv $issuedResult.StdOut)
if($issued.Count -gt $MaxIssuedRows){ $issued=$issued | Select-Object -First $MaxIssuedRows; Add-Finding 'Issuance' 'Warning' 'Issuance result truncated' "More than $MaxIssuedRows rows were returned." 0 }
$normalized=foreach($r in $issued){
    $nb=Convert-ToDateSafe (Get-FirstPropertyValue $r @('Certificate Effective Date','NotBefore'))
    $na=Convert-ToDateSafe (Get-FirstPropertyValue $r @('Certificate Expiration Date','NotAfter'))
    [pscustomobject]@{
        RequestID=Get-FirstPropertyValue $r @('Issued Request ID','Request ID','RequestID')
        Requester=Get-FirstPropertyValue $r @('Requester Name','RequesterName')
        Template=Get-FirstPropertyValue $r @('Certificate Template','CertificateTemplate')
        CommonName=Get-FirstPropertyValue $r @('Issued Common Name','Common Name','CommonName')
        NotBefore=$nb;NotAfter=$na
    }
}
$Evidence.IssuedCertificates=@($normalized)
$cutoff=(Get-Date).AddDays(-$RecentDays)
$recent=@($normalized | Where-Object { $_.NotBefore -and $_.NotBefore -ge $cutoff })
$unexpired=@($normalized | Where-Object { $_.NotAfter -and $_.NotAfter -gt (Get-Date) })
$latest=$normalized | Where-Object NotBefore | Sort-Object NotBefore -Descending | Select-Object -First 1
if($recent.Count -gt 0){ Add-Finding 'Issuance' 'Critical' 'Recent certificate issuance detected' "$($recent.Count) issued certificate(s) have NotBefore within $RecentDays days. Latest: $($latest.NotBefore)." 10 }
elseif($normalized.Count -gt 0){ Add-Finding 'Issuance' 'Info' 'Historical issuance only in parsed results' "$($normalized.Count) issued certificate(s) parsed. Latest: $($latest.NotBefore)." -2 }
else { Add-Finding 'Issuance' 'Warning' 'No issued rows parsed' 'This can mean no issuance or a certutil localization/permissions/parsing issue. Review raw-certutil-issued.txt.' 0 }
if($unexpired.Count -gt 0){ Add-Finding 'Issuance' 'Critical' 'Unexpired CA-issued certificates detected' "$($unexpired.Count) parsed certificate(s) are not expired." 8 }

# Local machine certificates issued by this CA
$stage = 'Local certificate inventory'
$localCerts=@(Get-ChildItem Cert:\LocalMachine\My -ErrorAction SilentlyContinue | Select-Object Subject,Issuer,Thumbprint,NotBefore,NotAfter,HasPrivateKey,@{N='EKU';E={$_.EnhancedKeyUsageList.FriendlyName -join ', '}})
$Evidence.LocalMachineCertificates=$localCerts
$localIssued=@($localCerts | Where-Object { $caName -and $_.Issuer -match [regex]::Escape($caName) })
$localValid=@($localIssued | Where-Object NotAfter -gt (Get-Date))
if($localValid.Count -gt 0){ Add-Finding 'Local Certificates' 'Critical' 'Valid locally installed CA-issued certificate found' "$($localValid.Count) valid certificate(s) on this server were issued by this CA." 8 }
elseif($localIssued.Count -gt 0){ Add-Finding 'Local Certificates' 'Info' 'Only expired local CA-issued certificates found' "$($localIssued.Count) local certificate(s) issued by this CA are expired." -1 }

# AD publication and templates
$stage = 'Active Directory publication discovery'
if(!$SkipActiveDirectory){
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        $cfg=([ADSI]'LDAP://RootDSE').configurationNamingContext
        $cas=Get-ADObject -SearchBase "CN=Enrollment Services,CN=Public Key Services,CN=Services,$cfg" -LDAPFilter '(objectClass=pKIEnrollmentService)' -Properties dNSHostName,certificateTemplates
        $Evidence.ADEnrollmentServices=@($cas | Select-Object Name,dNSHostName,certificateTemplates)
        $thisCA=@($cas | Where-Object { $_.Name -eq $caName -or $_.dNSHostName -eq ([System.Net.Dns]::GetHostEntry($env:COMPUTERNAME).HostName) })
        if($thisCA.Count -gt 0){ Add-Finding 'Active Directory' 'Attention' 'CA is published in Active Directory' 'The CA has an Enrollment Services object. This permits domain discovery and must be removed cleanly during decommissioning.' 3 }
    } catch { $Evidence.ADEnrollmentServices=@(); Add-Finding 'Active Directory' 'Warning' 'AD publication check failed' $_.Exception.Message 0 }
}

# IIS bindings and AD CS supporting services
$stage = 'IIS and supporting service discovery'
try {
    Import-Module WebAdministration -ErrorAction Stop
    $bindings=@(Get-WebBinding | Select-Object protocol,bindingInformation,certificateHash,certificateStoreName,@{N='Site';E={$_.ItemXPath -replace ".*name='([^']+)'.*",'$1'}})
    $Evidence.IISBindings=$bindings
    $https=@($bindings | Where-Object protocol -eq 'https')
    if($https.Count -gt 0){ Add-Finding 'IIS' 'Attention' 'HTTPS bindings present' "$($https.Count) HTTPS binding(s) exist. Validate whether they support Web Enrollment, CEP/CES, NDES, OCSP, or unrelated applications." 2 }
} catch { $Evidence.IISBindings=@() }
$serviceNames='CertSvc','W3SVC','OcspSvc','SCardSvr'
$Evidence.RelatedServices=@(Get-Service -Name $serviceNames -ErrorAction SilentlyContinue | Select-Object Name,DisplayName,Status,StartType)
$ndesReg=Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Cryptography\MSCEP' -ErrorAction SilentlyContinue
$Evidence.NDESConfigured=[bool]$ndesReg
if($ndesReg){ Add-Finding 'NDES' 'Critical' 'NDES configuration detected' 'The MSCEP registry key exists. Treat the CA as potentially active until NDES clients and certificate templates are validated.' 8 }

# SMB share and recent file activity
$stage = 'CertEnroll file discovery'
$certEnrollShare=Get-SmbShare -Name CertEnroll -ErrorAction SilentlyContinue
$Evidence.CertEnrollShare = $null
if($null -ne $certEnrollShare){ $Evidence.CertEnrollShare = $certEnrollShare | Select-Object Name,Path,Description }
$certEnrollPath="$env:windir\System32\CertSrv\CertEnroll"
$Evidence.CertEnrollFiles = @()
if(Test-Path $certEnrollPath){ $Evidence.CertEnrollFiles = @(Get-ChildItem $certEnrollPath -File | Select-Object Name,Length,LastWriteTime) }

# Optional events
$stage = 'Event log discovery'
if($IncludeEventLogs){
    $start=(Get-Date).AddDays(-$EventLookbackDays)
    $events=@()
    foreach($log in 'Application','Microsoft-Windows-CertificateServicesClient-Lifecycle-System/Operational'){
        try { $events += Get-WinEvent -FilterHashtable @{LogName=$log;StartTime=$start} -ErrorAction Stop | Where-Object { $_.ProviderName -match 'Certificate|CertSvc' } | Select-Object TimeCreated,Id,LevelDisplayName,ProviderName,Message -First 500 }
        catch { Add-Finding 'Events' 'Warning' "Event log unavailable: $log" $_.Exception.Message 0 }
    }
    $Evidence.Events=$events
    if($events.Count -gt 0){ Add-Finding 'Events' 'Attention' 'Recent certificate-related events found' "$($events.Count) matching event(s) found in the last $EventLookbackDays days. Review context before treating them as CA use." 1 }
}

# Classification. Strong evidence controls; a score is explanatory, not proof.
$stage = 'Classification'
$score=($Findings | Measure-Object Weight -Sum).Sum
if(!$score){$score=0}
$critical=@($Findings | Where-Object Severity -eq 'Critical')
if(!$svc){$classification='Not a CA'}
elseif($recent.Count -gt 0 -or $unexpired.Count -gt 0 -or $localValid.Count -gt 0 -or $ndesReg){$classification='Active'}
elseif($score -ge 5){$classification='Likely Active'}
elseif($normalized.Count -gt 0 -and !$latest){$classification='Inconclusive'}
elseif($normalized.Count -gt 0 -and $latest.NotBefore -lt $cutoff){$classification='Historical Only'}
elseif($normalized.Count -eq 0){$classification='No Issuance Evidence'}
else{$classification='Inconclusive'}
$confidence=if($issuedResult.ExitCode -eq 0 -and $normalized.Count -gt 0){'Moderate'}else{'Low'}
$summary=[ordered]@{
    Classification=$classification;Confidence=$confidence;EvidenceScore=$score;CAName=$caName;CAType=$caType
    IssuedParsed=$normalized.Count;RecentIssued=$recent.Count;UnexpiredIssued=$unexpired.Count
    LatestIssuance=if($latest){$latest.NotBefore}else{$null};PublishedTemplates=$templates.Count
    LocalValidIssuedByCA=$localValid.Count;NDESConfigured=[bool]$ndesReg
    RetirementRecommendation=if($classification -in 'Active','Likely Active'){'Do not retire. Investigate dependencies.'}elseif($classification -eq 'Historical Only'){'Candidate for retirement after endpoint, template, GPO, CRL/AIA, NDES/OCSP, and backup validation.'}else{'Do not retire based on this result alone. Resolve incomplete evidence.'}
}
$Evidence.Summary=$summary
$Evidence.Findings = $Findings.ToArray()

# Outputs
$stage = 'Report generation'
$Evidence | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $OutputPath 'ADCSUsageAudit.json') -Encoding UTF8
$normalized | Export-Csv (Join-Path $OutputPath 'IssuedCertificates.csv') -NoTypeInformation -Encoding UTF8
$Findings | Export-Csv (Join-Path $OutputPath 'Findings.csv') -NoTypeInformation -Encoding UTF8
$issuedResult.StdOut | Set-Content (Join-Path $OutputPath 'raw-certutil-issued.txt') -Encoding UTF8
$tplResult.StdOut | Set-Content (Join-Path $OutputPath 'raw-certutil-templates.txt') -Encoding UTF8
$caConfig.StdOut | Set-Content (Join-Path $OutputPath 'raw-certutil-ca-registry.txt') -Encoding UTF8

$text=@("AD CS Usage Audit","===================","Computer: $env:COMPUTERNAME","CA: $caName","Classification: $classification","Confidence: $confidence","Evidence score: $score","Issued parsed: $($normalized.Count)","Recent issued: $($recent.Count)","Unexpired issued: $($unexpired.Count)","Latest issuance: $(if($latest){$latest.NotBefore}else{'Unknown'})","Published templates: $($templates.Count)","Recommendation: $($summary.RetirementRecommendation)","","Findings:")
$text += $Findings | ForEach-Object { "[$($_.Severity)] $($_.Title): $($_.Detail)" }
$text | Set-Content (Join-Path $OutputPath 'Summary.txt') -Encoding UTF8

$rows=($Findings | ForEach-Object { "<tr><td>$(HtmlEncode $_.Severity)</td><td>$(HtmlEncode $_.Category)</td><td>$(HtmlEncode $_.Title)</td><td>$(HtmlEncode $_.Detail)</td></tr>" }) -join "`n"
$html=@"
<!doctype html><html><head><meta charset="utf-8"><title>AD CS Usage Audit</title><style>body{font-family:Segoe UI,Arial;margin:2rem;color:#1f2937}h1{color:#075985}.card{background:#e5f2f9;border-left:6px solid #0080bb;padding:1rem;margin:1rem 0}.critical{background:#f4cccc}table{border-collapse:collapse;width:100%}th,td{border:1px solid #cbd5e1;padding:.5rem;text-align:left;vertical-align:top}th{background:#374b5d;color:white}code{background:#f1f5f9;padding:.1rem .25rem}</style></head><body><h1>AD CS Usage Audit</h1><div class="card"><strong>Classification:</strong> $(HtmlEncode $classification)<br><strong>Confidence:</strong> $(HtmlEncode $confidence)<br><strong>CA:</strong> $(HtmlEncode $caName)<br><strong>Latest issuance:</strong> $(HtmlEncode $(if($latest){$latest.NotBefore}else{'Unknown'}))<br><strong>Recommendation:</strong> $(HtmlEncode $summary.RetirementRecommendation)</div><p>This report is an evidence-based heuristic. It cannot prove that no application trusts or depends on the CA.</p><table><thead><tr><th>Severity</th><th>Category</th><th>Finding</th><th>Detail</th></tr></thead><tbody>$rows</tbody></table></body></html>
"@
$html | Set-Content (Join-Path $OutputPath 'Report.html') -Encoding UTF8

$color = 'Cyan'
if($classification -in 'Active','Likely Active'){ $color = 'Red' } elseif($classification -eq 'Historical Only'){ $color = 'Yellow' }
Write-Host "Classification: $classification" -ForegroundColor $color
Write-Host "Report: $OutputPath"
[pscustomobject]$summary
