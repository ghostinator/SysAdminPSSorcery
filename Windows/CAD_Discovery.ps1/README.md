# CAD File Discovery & Alerting Tool

A robust PowerShell automation script designed to locate massive CAD project files across local and external drives, compile a CSV report, and email the results via the SendGrid API. 

Optimized for deployment as a Datto RMM Component, this script includes built-in workarounds for common PowerShell 5.1 terminating errors and filters out irrelevant removable storage.

## Features

* **Targeted Application Scanning:** Choose to search specifically for Autodesk, SolidWorks, Catia, MicroStation, SketchUp, Rhino, PTC Creo, or Universal 3D formats (or scan for all of them at once).
* **Smart Drive Discovery:** Automatically identifies Local Fixed and Removable drives, but completely ignores any drive with a total capacity under 50GB (preventing it from wasting time scanning small OS partitions or random 4GB thumb drives).
* **PS 5.1 Bug Bypass:** Avoids the classic PowerShell `Access is denied` terminating error caused by scanning root directories (`C:\*`) by breaking the recursion out into top-level directory parsing.
* **Self-Cleaning:** Generates local transcripts and CSVs in `C:\temp` for auditing, but automatically deletes logs older than 14 days to prevent drive bloat.
* **SendGrid API Integration:** Uses SendGrid's v3 API via HTTPS to deliver the CSV report, bypassing common local firewall blocks on port 587/25 and ensuring clean SPF/DKIM alignment.

## Datto RMM Component Setup

To deploy this script via Datto RMM, create a new PowerShell component and configure the following variables exactly as shown:

| Variable Name | Type | Description |
| :--- | :--- | :--- |
| `SENDGRID_API_KEY` | String (Secret) | Your SendGrid API key. **Highly recommended to set this as a hidden Site or Global variable.** |
| `MAIL_FROM` | String | The verified sender email address in your SendGrid account (e.g., `alerts@yourdomain.com`). |
| `MAIL_TO` | String | The destination email address for the report. |
| `CAD_APP_TARGET` | Selection | Options: `Autodesk, SolidWorks, Catia, MicroStation, SketchUp, Rhino, PTCCreo, Universal, All` |

### Execution Details
* **Run As:** `SYSTEM`
* **Network Impact:** Low. The script ignores network shares (`DriveType 4`) to prevent overwhelming file servers and only scans physically attached disks.
* **Performance:** High disk I/O during the scan. It is recommended to schedule this component to run off-hours to avoid impacting active workstation performance.

## Standalone Usage

While designed for Datto RMM, you can run this script standalone by manually setting the environment variables in your PowerShell session before execution:

```powershell
$env:SENDGRID_API_KEY = "your_api_key_here"
$env:MAIL_FROM = "alerts@yourdomain.com"
$env:MAIL_TO = "you@yourdomain.com"
$env:CAD_APP_TARGET = "Autodesk"

.\CAD_Discovery.ps1
