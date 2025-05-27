# Automatic Timezone Configuration for Windows (v1.3)

A robust PowerShell solution that automatically detects and configures the correct timezone on Windows computers based on their public IP address geolocation. This script is designed for organizations with mobile users, remote workers, or devices that travel, ensuring accurate timekeeping. **This version (v1.3) enhances conflict prevention by disabling Windows native automatic timezone features, defaults to Eastern Time on detection failures, deletes pre-existing scheduled tasks of the same name, and runs the task immediately after setup.**

## 🌍 Overview

This solution automates timezone management by:
- **Disabling Windows built-in automatic timezone features** to ensure consistent behavior and prevent conflicts.
- Detecting the device's public IP address using multiple reliable external services.
- Geolocating the IP address to determine the city, region, country, and IANA (Internet Assigned Numbers Authority) timezone.
- **Defaulting to "Eastern Standard Time"** if IP lookup or geolocation fails.
- Converting the IANA timezone identifier to the appropriate Windows timezone format using an extensive internal mapping.
    - If a received IANA timezone cannot be mapped, it also defaults to "Eastern Standard Time".
- Setting the system timezone accordingly.
- Triggering updates on system startup, user logon, and specific network connection events.
- Maintaining a detailed log of all operations and comprehensive tracking data in the registry.

## ✨ Features

- **🔧 Conflict Prevention**: Actively disables Windows built-in location-based timezone services (`tzautoupdate`) and related settings to ensure the script has sole control.
- **🌐 Automatic & Resilient Detection**: Uses IP geolocation (via `ip-api.com`) and multiple IP lookup services (`api.ipify.org`, `ipinfo.io/ip`, etc.) for reliable timezone determination.
- **⏰ Intelligent Fallback**: Defaults to "Eastern Standard Time" if accurate geolocation or IANA timezone mapping fails.
- **🗺️ Comprehensive Timezone Mapping**: Includes an extensive, updatable list mapping IANA timezones (e.g., "America/Indiana/Indianapolis") to Windows Time Zone IDs (e.g., "US Eastern Standard Time").
- **⚙️ Automated Scheduled Task**:
    - Named "AutoTimezoneUpdate".
    - Runs with `NT AUTHORITY\SYSTEM` privileges.
    - Triggers on system startup, user logon, and network connection events (NetworkProfile Event ID 10000).
    - The setup script **deletes any pre-existing task with the same name** before creating the new one.
    - The setup script **runs the task immediately** after successful configuration.
- **🏢 Enterprise Ready**: Designed for deployment via Microsoft Intune, Group Policy, or other Remote Management and Monitoring (RMM) tools.
- **📊 Enhanced Registry Tracking**: Stores detailed operational metadata in `HKLM:\SOFTWARE\AutoTimezone` (last IP, detected/set timezones, location, script version, update status, and timestamps).
- **📝 Detailed Logging**: All actions, decisions, and errors of the `UpdateTimezone.ps1` script are logged to `C:\Scripts\TimezoneUpdate.log`.
- **🛡️ Robust Error Handling**: Includes mechanisms to handle failures in IP/geolocation lookups and timezone setting, with fallback to `tzutil.exe` if `Set-TimeZone` fails.
- **👻 Silent Operation**: Runs silently in the background with no user interaction required for `UpdateTimezone.ps1`.

## 🚨 Problem Solved

Addresses the common issue where Windows devices incorrectly determine their timezone (e.g., defaulting to "UTC", showing "E. Africa Standard Time" for a US-based location) due to:
- Inaccurate geolocation data from nearby Wi-Fi access points.
- Unreliable Windows Location Services for timezone determination.
- Conflicts between user settings and system attempts to auto-set the timezone.

This solution provides a consistent and predictable timezone by relying on IP-based geolocation and taking control from less reliable native Windows mechanisms.

## 🚀 Quick Start / Deployment

This solution is deployed using a **setup script** (e.g., `Timezone_AutoConfig_Setup_v1.3.ps1` - the script generated in our previous interactions). This setup script performs a one-time installation on each target machine.

### Deployment Steps:

1.  **Obtain the Setup Script:** Use the complete PowerShell script generated previously (which includes the setup logic and the embedded `UpdateTimezone.ps1` content).
2.  **Deploy the Setup Script:**
    * **Microsoft Intune (Recommended):**
        1.  Go to **Devices > Scripts** in the Intune admin center.
        2.  Click **Add > Windows 10 and later**.
        3.  Name the script policy (e.g., "Deploy Timezone Auto-Configuration v1.3").
        4.  Upload the setup script (`Timezone_AutoConfig_Setup_v1.3.ps1`).
        5.  Configure settings:
            * Run this script using the logged on credentials: **No** (to run as SYSTEM)
            * Enforce script signature check: **No** (unless your script is signed and trusted)
            * Run script in 64-bit PowerShell Host: **Yes**
        6.  Assign the script policy to your target security group of devices (e.g., "Intune - AutoTimeZone").
    * **Group Policy:** Deploy the setup script as a computer startup script.
    * **Other RMM Tools:** Deploy and run the setup script with SYSTEM privileges.

### What the Setup Script Does:
* Creates the `C:\Scripts` directory.
* Writes the core detection script to `C:\Scripts\UpdateTimezone.ps1` (Version 1.3 logic).
* Deletes any existing "AutoTimezoneUpdate" scheduled task.
* Creates the "AutoTimezoneUpdate" scheduled task (triggers on startup, logon, network events; runs as SYSTEM).
* **Immediately runs the "AutoTimezoneUpdate" task once** to apply the timezone and create initial log/registry entries.

## 📁 Installed Components on Target Machine

* **Main Script:** `C:\Scripts\UpdateTimezone.ps1` (core logic, v1.3)
* **Log File:** `C:\Scripts\TimezoneUpdate.log` (detailed activity log)
* **Scheduled Task:** "AutoTimezoneUpdate" (triggers script execution)
* **Registry Key for Tracking:** `HKLM:\SOFTWARE\AutoTimezone` (stores operational data)

## 🔧 System Modifications by `UpdateTimezone.ps1`

* **Windows Time Zone Auto Update Service (`tzautoupdate`):** Startup type set to "Disabled" and service stopped.
* **Location Policy Keys:** Attempts to set `HKLM:\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors` keys to disable OS location use for timezones.
* **Location Capability Consent:** `HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location\Value` set to "Deny".
* **Dynamic DST:** `HKLM:\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\DynamicDaylightTimeDisabled` ensured to be `0` (allowing DST based on zone rules).

## 🌐 Supported Timezones (Examples in Mapping)

The `UpdateTimezone.ps1` script contains an extensible mapping table. Key entries include:

| Region          | IANA Examples                   | Windows Time Zone ID Examples     |
| :-------------- | :------------------------------ | :-------------------------------- |
| North America   | `America/New_York`              | `Eastern Standard Time`           |
|                 | `America/Chicago`               | `Central Standard Time`           |
|                 | `America/Denver`                | `Mountain Standard Time`          |
|                 | `America/Phoenix`               | `US Mountain Standard Time`       |
|                 | `America/Los_Angeles`           | `Pacific Standard Time`           |
|                 | `America/Indiana/Indianapolis` | `US Eastern Standard Time`        |
| Europe          | `Europe/London`                 | `GMT Standard Time`               |
|                 | `Europe/Berlin`                 | `W. Europe Standard Time`         |
|                 | `Europe/Paris`                  | `Romance Standard Time`           |
| Asia            | `Asia/Tokyo`                    | `Tokyo Standard Time`             |
|                 | `Asia/Dubai`                    | `Arabian Standard Time`           |
| Australia       | `Australia/Sydney`              | `AUS Eastern Standard Time`       |
| **Default/Fallback** | (If IP/Geo/IANA map fails) | `Eastern Standard Time`           |

*(Refer to the `$timezoneMapping` variable within `UpdateTimezone.ps1` for the full list; this table is illustrative.)*

## 📊 How `UpdateTimezone.ps1` Works (Flowchart)

```mermaid
graph TD
    A[Scheduled Task Triggered <br/>(Startup/Logon/Network Event)] --> B(Initialize & Start Logging <br/>to C:\Scripts\TimezoneUpdate.log);
    B --> C{Disable Windows <br/>Auto Timezone Features};
    C --> D[Get Public IP Address <br/>(Multiple Services)];
    D -- IP Lookup Fails --> E[Set Target: Eastern Standard Time <br/>Log Failure];
    D -- IP Lookup Succeeds --> F[Get Geolocation Data <br/>(ip-api.com)];
    F -- Geo Lookup Fails <br/>OR No IANA TZ --> E;
    F -- Geo Lookup Succeeds --> G[Extract IANA Timezone];
    G --> H{Convert IANA to Windows TZ <br/>(Using Internal Map)};
    H -- Mapping Fails --> E;
    H -- Mapping Succeeds --> I[Set Target: Mapped Windows TZ <br/>Log Success];
    E --> J(Set System Timezone <br/>to Target);
    I --> J;
    J --> K[Update Registry Tracking <br/>(HKLM:\SOFTWARE\AutoTimezone)];
    K --> L[Stop Logging & Exit];
```

## 🔧 Configuration & Customization (`C:\Scripts\UpdateTimezone.ps1`)

* **IP Detection Services:** Modify the `$uris` array in the `Get-PublicIPAddress` function within `C:\Scripts\UpdateTimezone.ps1` to change or add new public IP lookup endpoints.
* **Geolocation API:** The `Get-GeoLocationData` function currently uses `http://ip-api.com/json/`. This URL can be modified if you choose to use a different geolocation service (the parsing logic might also need adjustment).
* **Timezone Mapping:** The `$timezoneMapping` hashtable inside the `Convert-IANAToWindowsTimeZone` function within `C:\Scripts\UpdateTimezone.ps1` is designed to be extensible. You can add more IANA Time Zone strings and their corresponding Windows Time Zone ID pairs as needed.
* **Fallback Timezone:** The primary fallback timezone (used if IP/geolocation lookup fails, or if an IANA timezone string cannot be mapped by the `Convert-IANAToWindowsTimeZone` function) is set to "Eastern Standard Time". This can be changed in the `Main` function of `C:\Scripts\UpdateTimezone.ps1` (variable `$windowsTimeZoneToSet`) and in the final `else` block of the `Convert-IANAToWindowsTimeZone` function.
* **Logging:** The `UpdateTimezone.ps1` script uses `Start-Transcript` to log its activity. The log file path is defined as `$LogFile` near the top of `UpdateTimezone.ps1` (and set by the setup script to `C:\Scripts\TimezoneUpdate.log`).

## 📝 Logging Details

All significant operations, decisions, and errors of the `UpdateTimezone.ps1` script are logged with timestamps to `C:\Scripts\TimezoneUpdate.log`. This provides a comprehensive audit trail for troubleshooting.

Example log entries:
2025-05-27 10:05:39] (1.3-ET-Fallback) Script execution started.
Attempting to disable/control Windows automatic timezone features...
✓ Windows Time Zone Auto Update service (tzautoupdate) set to Disabled.
...
Attempting to get public IP from https://api.ipify.org/...
Public IP Address: 72.2.154.154 (from https://api.ipify.org/)
Querying geolocation API: http://ip-api.com/json/72.2.154.154
Location Details (from ip-api.com):
City: Elkhart, Region: Indiana, Country: United States
IANA Timezone: America/Indiana/Indianapolis
Successfully retrieved geolocation: IP: 72.2.154.154, City: Elkhart, Region: Indiana, Country: United States, IANA TZ: America/Indiana/Indianapolis
Found explicit mapping for IANA 'America/Indiana/Indianapolis': 'US Eastern Standard Time'

Attempting to set system timezone to 'US Eastern Standard Time'...
Successfully set timezone to 'US Eastern Standard Time'
Current local time: 05/27/2025 10:05:45 AM
Updated registry tracking information.
[2025-05-27 10:05:45] (1.3-ET-Fallback) Script execution finished. Status: Successfully set timezone to 'US Eastern Standard Time'. (IP: 72.2.154.154, City: Elkhart, Region: Indiana, Country: United States, IANA TZ: America/Indiana/Indianapolis)

## 🛠️ Troubleshooting and Verification

For detailed steps on verifying the deployment, checking logs, registry keys, scheduled task status, and troubleshooting common issues, please refer to **Section 4 (Verification Procedures)** and **Section 5 (Troubleshooting Guide)** of the comprehensive Knowledge Base Article you have.

Key verification commands to run in PowerShell on a target device:
* **Check current system timezone:**
    ```powershell
    Get-TimeZone
    tzutil /g
    ```
* **View recent log entries from `UpdateTimezone.ps1`:**
    ```powershell
    Get-Content "C:\Scripts\TimezoneUpdate.log" -Tail 20
    ```
* **Search logs for errors or specific successes:**
    ```powershell
    Select-String -Path "C:\Scripts\TimezoneUpdate.log" -Pattern "Error|Warning|Failed" -CaseSensitive
    Select-String -Path "C:\Scripts\TimezoneUpdate.log" -Pattern "Successfully set timezone to"
    ```
* **View registry tracking data written by `UpdateTimezone.ps1`:**
    ```powershell
    Get-ItemProperty -Path "HKLM:\SOFTWARE\AutoTimezone"
    ```
* **Check the status and last run result of the scheduled task:**
    ```powershell
    Get-ScheduledTask -TaskName "AutoTimezoneUpdate" | Get-ScheduledTaskInfo
    ```

## 🔒 Security Considerations

* **SYSTEM Privileges:** The initial setup script requires Administrator rights to create files in `C:\Scripts`, create registry keys under `HKLM`, and create the scheduled task. The `UpdateTimezone.ps1` script itself is executed as `NT AUTHORITY\SYSTEM` via the scheduled task, which is necessary for modifying the system timezone and writing to HKLM registry keys.
* **External API Calls:** The `UpdateTimezone.ps1` script makes calls to external, third-party services (e.g., `api.ipify.org`, `ipinfo.io/ip`, `http://ip-api.com/`) to determine the public IP address and perform geolocation. Ensure these endpoints are considered trusted within your organization and that outbound connectivity to them (HTTPS/HTTP) is permitted through firewalls or proxies. Note that the free tier of `ip-api.com` uses HTTP.
* **Script Integrity:** Scripts deployed to endpoints should always come from a trusted internal source or be signed. If deploying via Intune, you can enforce script signature checking if your organization signs its PowerShell scripts.
* **Logging Data:** The log file (`C:\Scripts\TimezoneUpdate.log`) and registry entries (`HKLM:\SOFTWARE\AutoTimezone`) will contain the device's public IP address and geolocation information. Ensure access to these locations is appropriately controlled if this information is considered sensitive in your environment.

## ⚠️ Important Notes & Reverting Changes

* **Authoritative Control:** This solution is designed to take authoritative control of the system's timezone settings. It actively disables several Windows native automatic timezone features to prevent conflicts.
* **Reverting to Windows Default Behavior:** To undo the changes made by this solution and revert to standard Windows timezone management:
    1.  **Delete the Scheduled Task:**
        ```powershell
        Unregister-ScheduledTask -TaskName "AutoTimezoneUpdate" -Confirm:$false -ErrorAction SilentlyContinue
        ```
    2.  **Re-enable the Windows Time Zone Auto Update Service:**
        ```powershell
        Set-Service -Name "tzautoupdate" -StartupType Automatic -ErrorAction SilentlyContinue
        Start-Service -Name "tzautoupdate" -ErrorAction SilentlyContinue
        ```
    3.  **Optional: Remove Script and Registry Data:**
        ```powershell
        Remove-Item -Path "C:\Scripts" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path "HKLM:\SOFTWARE\AutoTimezone" -Recurse -Force -ErrorAction SilentlyContinue
        ```
    4.  **Optional: Revert Location Policy/Consent (if necessary):**
        ```powershell
        Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors" -Name "DisableWindowsLocationProvider" -Force -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors" -Name "DisableLocation" -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location" -Name "Value" -Value "Allow" -ErrorAction SilentlyContinue
        ```
    5.  **Manually adjust settings in Windows:** Go to **Settings > Time & Language > Date & time** and enable "Set time zone automatically" and "Adjust for daylight saving time automatically" if desired.

## 🤝 Contributing

If this were a public project, contributions would be welcome! Potential areas for improvement:
-   Further expansion of the IANA to Windows timezone mapping table in `UpdateTimezone.ps1`.
-   Addition of more IP/geolocation service options with failover logic.
-   More granular error handling and reporting.
-   GUI for viewing logs or status.

## 📋 Requirements

-   **Operating System**: Windows 10 / Windows 11. PowerShell 5.1 or later should be available.
-   **Permissions**: Administrator rights are required on the target machine to run the initial setup script.
-   **Network Access**: The `UpdateTimezone.ps1` script requires outbound internet connectivity to contact IP detection and geolocation APIs.
-   **PowerShell Execution Policy**: The setup script and scheduled task use `-ExecutionPolicy Bypass` to ensure the `UpdateTimezone.ps1` script can run.

## 🆕 Version History

* **v1.3 (Current - This Version)**
    * Setup script now **deletes any pre-existing scheduled task** with the same name before creation.
    * Setup script now **runs the scheduled task immediately** after creation/update.
    * `UpdateTimezone.ps1`:
        * Enhanced `Disable-WindowsAutomaticTimezone` for more comprehensive control of native Windows features.
        * `Get-PublicIPAddress` function uses multiple IP lookup services for improved reliability.
        * `Convert-IANAToWindowsTimeZone` function:
            * Internal fallback (if an IANA string is received but cannot be mapped) changed to "Eastern Standard Time".
            * Explicit mapping for "America/Indiana/Indianapolis" to "US Eastern Standard Time" added.
            * Generally expanded IANA mapping table.
        * `Set-SystemTimeZone` function includes a fallback to use `tzutil.exe` if the `Set-TimeZone` cmdlet fails.
        * Improved logging structure using `Start-Transcript` for better detail in `C:\Scripts\TimezoneUpdate.log`.
        * Script's overall default target timezone (if IP or geo lookup fails entirely) set to "Eastern Standard Time".
* **v1.2 (Conceptual based on previous discussion)**
    * Primary default for `UpdateTimezone.ps1` (if IP/Geo lookup failed) set to "Eastern Standard Time".
    * `Convert-IANAToWindowsTimeZone` internal fallback (for unmappable IANA strings) also set to "Eastern Standard Time".
* **v1.1 (User's Initial Script Base)**
    * Included conflict prevention by disabling some Windows auto-timezone features.
    * Implemented registry tracking for operational data.
* **v1.0 (Original Concept)**
    * Basic IP-based timezone detection.
    * Network change monitoring.
    * Scheduled task for automation.

## 📄 License

`This project is licensed under the MIT License - see the LICENSE.md file for details.`

## 🙏 Acknowledgments

-   IP & Geolocation API Providers: `ipify.org`, `ipinfo.io`, `icanhazip.com`, `checkip.amazonaws.com`, `ip-api.com`.
-   The Microsoft PowerShell Team for `Set-TimeZone`, `Get-TimeZone`, and other useful cmdlets.
-   The broader Windows Administration and PowerShell communities for shared knowledge and inspiration.

---

⭐ *If this solution helps you effectively manage device timezones and resolve conflicts, consider sharing your feedback or starring the project* ⭐
