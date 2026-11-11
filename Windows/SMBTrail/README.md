# SMBTrail

SMBTrail is a PowerShell-based observation tool for Windows file servers. It samples active SMB sessions and open files over time, exports SMB1 audit events, inventories shares and server settings, creates a ZIP report, and can deliver the result through several email transports.

It is designed for migration discovery, SMB1 dependency identification, file-server retirement planning, segmentation projects, and validating whether a server can be replaced by a NAS or cloud service.

## Highlights

- Runs on Windows PowerShell 5.1 or later
- Samples `Get-SmbSession` and `Get-SmbOpenFile`
- Records client, user, SMB dialect, path, and observation frequency
- Exports SMB1 audit Event ID 3000 from `Microsoft-Windows-SMBServer/Audit`
- Captures SMB share and server configuration baselines
- Produces CSV, JSON, TXT, and ZIP output
- Installs and removes its own SYSTEM scheduled tasks
- Supports file-only output, SMTP, SendGrid, SMTP2GO, Postal, and a generic JSON API
- Supports machine-level environment variables so secrets do not need to be stored in the config file

## Important limitations

SMBTrail is an observation tool, not a packet analyzer. Short-lived sessions or files that open and close between collection intervals can be missed. Reduce the interval for busier or more transient workloads. If the remaining uncertainty requires packet-level evidence, use `pktmon`, Wireshark, or a firewall capture after SMBTrail narrows the clients and time windows.

`Get-SmbOpenFile` can expose filenames and paths. Treat the report as sensitive data. Restrict access to the output folder and use an approved transport.

## Requirements

- Windows Server with the `SmbShare` PowerShell module
- Local administrator rights for installation
- Task Scheduler service
- SMB1 auditing already enabled if SMB1 Event ID 3000 collection is desired
- Outbound HTTPS or SMTP access when delivery is enabled

Check SMB settings:

```powershell
Get-SmbServerConfiguration |
  Select-Object EnableSMB1Protocol, EnableSMB2Protocol, AuditSmb1Access
```

Enable SMB1 auditing without enabling SMB1 itself:

```powershell
Set-SmbServerConfiguration -AuditSmb1Access $true -Force
```

## Quick start

1. Copy `SMBTrail.ps1` and `SMBTrail.config.example.json` to the server.
2. Rename the example config:

```powershell
Rename-Item .\SMBTrail.config.example.json SMBTrail.config.json
```

3. Edit the collection duration, interval, output path, and delivery method.
4. Run an elevated Windows PowerShell session:

```powershell
Set-ExecutionPolicy -Scope Process Bypass -Force
.\SMBTrail.ps1 -Mode Install -ConfigPath .\SMBTrail.config.json
```

5. Verify status:

```powershell
C:\ProgramData\SMBTrail\SMBTrail.ps1 -Mode Status `
  -ConfigPath C:\ProgramData\SMBTrail\SMBTrail.config.json
```

## Modes

| Mode | Purpose |
|---|---|
| `Install` | Copies the tool and config, captures a baseline, and schedules collection/finalization |
| `Collect` | Performs one session/open-file sample |
| `Finalize` | Exports SMB1 events, builds reports, delivers the ZIP, and removes tasks |
| `Status` | Shows schedule and delivery status |
| `TestDelivery` | Sends a small test ZIP using the selected transport |
| `Uninstall` | Removes scheduled tasks and retains collected output |

## Collection guidance

- Default: 7 days at a 5-minute interval
- Legacy clinical, manufacturing, or scanning workflows: 14 days is safer
- Busy file servers: 1 to 2 minutes, with attention to output growth
- Infrequent month-end workflows: align the observation window with that process

## Delivery methods

Set `Delivery.Method` to one of the following.

### None

Leaves the ZIP report in `OutputPath`.

```json
"Method": "None"
```

### SMTP

Works with standard SMTP relays, including SMTP2GO's SMTP service, Microsoft Exchange relay, authenticated smarthosts, and self-hosted mail servers.

```json
{
  "Method": "SMTP",
  "To": "admin@example.com",
  "From": "smbtrail@example.com",
  "Host": "smtp.example.com",
  "Port": 587,
  "UseTls": true,
  "Username": "smtp-user",
  "PasswordEnvironmentVariable": "SMBTRAIL_SMTP_PASSWORD"
}
```

Set the machine secret before installation:

```powershell
[Environment]::SetEnvironmentVariable(
  'SMBTRAIL_SMTP_PASSWORD',
  'replace-me',
  'Machine'
)
```

### SendGrid API

Uses `POST https://api.sendgrid.com/v3/mail/send` with bearer authentication.

```json
{
  "Method": "SendGrid",
  "To": "admin@example.com",
  "From": "verified-sender@example.com",
  "FromName": "SMBTrail",
  "ApiKeyEnvironmentVariable": "SMBTRAIL_API_KEY"
}
```

Use an API key limited to Mail Send.

### SMTP2GO API

Uses `POST https://api.smtp2go.com/v3/email/send` with the `X-Smtp2go-Api-Key` header.

```json
{
  "Method": "SMTP2GO",
  "To": "admin@example.com",
  "From": "verified-sender@example.com",
  "ApiKeyEnvironmentVariable": "SMBTRAIL_API_KEY"
}
```

### Postal API

Uses the Postal server API endpoint `/api/v1/send/message` and the `X-Server-API-Key` header.

```json
{
  "Method": "Postal",
  "To": "admin@example.com",
  "From": "smbtrail@example.com",
  "BaseUrl": "https://postal.example.com",
  "ApiKeyEnvironmentVariable": "SMBTRAIL_API_KEY"
}
```

Postal deployments can vary by version and administrator configuration. Run `TestDelivery` before installing a long observation job.

### Generic JSON API

Posts a predictable JSON body containing `from`, `to`, `subject`, `text`, and a base64 ZIP attachment. This is intended for internal gateways or a lightweight adapter service.

```json
{
  "Method": "GenericApi",
  "To": "admin@example.com",
  "From": "smbtrail@example.com",
  "Uri": "https://api.example.com/send",
  "HeaderName": "Authorization",
  "HeaderPrefix": "Bearer ",
  "ApiKeyEnvironmentVariable": "SMBTRAIL_API_KEY"
}
```

## Test delivery

Run this before starting the observation period:

```powershell
.\SMBTrail.ps1 -Mode TestDelivery -ConfigPath .\SMBTrail.config.json
```

## Output

- `SMB-Sessions.csv`
- `SMB-OpenFiles.csv`
- `SMB1-Audit-Events.csv`
- `SMB-Shares.csv`
- `SMB-Server-Configuration.csv`
- `Summary.json`
- `Summary.txt`
- `Errors.log`, when errors occur
- `SMBTrail-SERVER-yyyyMMdd-HHmmss.zip`

## Security guidance

- Prefer machine environment variables for secrets.
- Use provider keys restricted to mail sending only.
- Revoke one-time API keys after delivery.
- Apply restrictive NTFS permissions to the output path.
- Do not place a populated config file or report ZIP in source control.
- Treat captured usernames, IP addresses, filenames, and paths as sensitive operational data.

## Repository hygiene

The included `.gitignore` excludes runtime configs, archives, CSVs, logs, and generated reports.

## License

MIT
