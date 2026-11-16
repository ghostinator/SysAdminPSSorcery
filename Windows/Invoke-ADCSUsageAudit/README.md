# ADCSUsageAudit

A read-only PowerShell 5.1 audit for determining whether a Microsoft Active Directory Certificate Services Certification Authority shows evidence of current use.

## Why it exists

An installed and running `CertSvc` service does not prove that a CA is still used. An empty recent-issuance query also does not prove that retirement is safe. ADCSUsageAudit combines several signals and produces a cautious classification with the supporting evidence.

The script does not uninstall roles, revoke certificates, modify templates, change Active Directory, export private keys, or alter CA configuration.

## Output classifications

- **Active**: Strong evidence exists, such as recent issuance, unexpired CA-issued certificates, a valid local CA-issued certificate, or NDES configuration.
- **Likely Active**: Multiple weaker activity indicators exist.
- **Historical Only**: Issuance was found, but the parsed evidence is older than the selected recent-use window.
- **No Issuance Evidence**: No issued rows were parsed. This is not the same as proving the CA is unused.
- **Inconclusive**: Evidence is incomplete or contradictory.
- **Not a CA**: `CertSvc` is absent.

## Requirements

- Windows PowerShell 5.1
- Run locally on the CA server
- Elevated administrative PowerShell session recommended
- `certutil.exe`
- Server Manager PowerShell module for role inventory
- Active Directory module for Enrollment Services discovery, unless `-SkipActiveDirectory` is used

## Basic usage

    .\Invoke-ADCSUsageAudit.ps1

Use 365 days as the recent-issuance threshold and include 90 days of event evidence:

    .\Invoke-ADCSUsageAudit.ps1 -RecentDays 365 -IncludeEventLogs -EventLookbackDays 90

Write reports to a specific folder:

    .\Invoke-ADCSUsageAudit.ps1 -OutputPath C:\Temp\ADCS-Audit

Skip Active Directory discovery:

    .\Invoke-ADCSUsageAudit.ps1 -SkipActiveDirectory

## Produced files

- `Summary.txt`
- `Report.html`
- `ADCSUsageAudit.json`
- `Findings.csv`
- `IssuedCertificates.csv`
- `raw-certutil-issued.txt`
- `raw-certutil-templates.txt`
- `raw-certutil-ca-registry.txt`

Review the raw files when the JSON report says confidence is low. `certutil` output can vary by Windows version, locale, permission level, and CA database state.

## Retirement gate

Do not retire a CA based on the classification alone. Before removal, independently validate:

1. No recent issuance or renewal activity.
2. No unexpired certificates issued by the CA.
3. No NDES, SCEP, OCSP, CEP, CES, Web Enrollment, VPN, RADIUS, Wi-Fi, smart-card, EFS, code-signing, or application dependency.
4. Published templates and auto-enrollment GPOs have been reviewed.
5. AIA and CRL Distribution Point dependencies have been reviewed.
6. The CA database, registry configuration, CA certificate, and private key are backed up under an approved recovery procedure.
7. Active Directory Enrollment Services, AIA, CDP, template publication, and trust objects have an approved cleanup plan.
8. A controlled change and rollback plan exists.

## Interpreting false positives

- A running service may only be publishing CRLs for historical certificates.
- Published templates show enrollment capability, not necessarily current enrollment.
- HTTPS bindings may belong to workloads unrelated to AD CS.
- Old certificates in the local machine store may be harmless remnants.
- Certificate-related events may reflect local certificate maintenance rather than CA issuance.

## Interpreting false negatives

- `certutil` parsing can fail because of localization or permissions.
- A relying application may trust a root certificate without creating recent CA database activity.
- Offline systems may still depend on an old certificate chain.
- A CA may only be used during infrequent annual or multi-year renewal cycles.
- Another CA can issue from the same templates, so domain-wide discovery matters.

## Security

The reports may contain hostnames, certificate subjects, requesters, template names, and infrastructure metadata. Review and redact output before posting it publicly. The script intentionally does not collect password values or export private keys.

## License suggestion

MIT is a reasonable default for a public utility. Add your preferred `LICENSE` file before publishing.

## Microsoft note

Microsoft documents `certutil.exe` as an administrative and developer diagnostic tool and cautions that it is not recommended as a production-code interface with guaranteed compatibility. This project therefore retains raw output, reports parse confidence, and avoids making an automatic decommission decision.
