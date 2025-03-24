
# Network Connectivity Watchdog

## Description

The Network Connectivity Watchdog is a PowerShell script designed to continuously monitor network connectivity and automatically reset the network adapter if issues are detected. This tool is universal and can work with any type of network adapter, making it suitable for various environments. It provides real-time monitoring with a dashboard interface showing connectivity status, test results, and adapter statistics.

## Features

-   **Universal Adapter Support:** Works with any network adapter type (Ethernet, Wi-Fi, USB, etc.).
-   **Automated Reset:** Automatically resets the network adapter upon detecting persistent connectivity issues.
-   **Real-time Monitoring:** Provides a dashboard interface with real-time status updates.
-   **Customizable:** Offers command-line parameters for easy configuration.
-   **Detailed Logging:** Logs all activities and errors for troubleshooting.
-   **Dynamic Gateway Detection:** Automatically detects the default gateway for connectivity testing.
-   **DNS Resolution Testing:** Includes DNS resolution testing to ensure proper domain name resolution.
-   **PowerShell Best Practices:** Follows PowerShell documentation and coding standards.

## Requirements

-   PowerShell 5.1 or later
-   Administrator privileges (required for resetting the network adapter)

## Installation

1.  Download the `NetworkConnectivityWatchdog.ps1` script from the [GitHub repository](https://github.com/[username]/NetworkConnectivityWatchdog).
2.  Save the script to a directory on your system.

## Usage

To run the script, open PowerShell as an administrator and navigate to the directory where you saved the script. Then, execute the script using:

```powershell
.\NetworkConnectivityWatchdog.ps1
```

### Parameters

The script supports the following command-line parameters:

-   `-AdapterPattern`: Specifies the name pattern to identify your network adapter. Supports wildcards.
    -   Examples:
        -   `"Ethernet*"`: Matches any adapter starting with "Ethernet".
        -   `"Wi-Fi*"`: Matches any Wi-Fi adapter.
        -   `"usb_xhci*"`: Matches USB network adapters.
        -   `"*"`: Matches all adapters (will use the first one found).
    -   Default: `"*"`
-   `-FailureThreshold`: Specifies the number of seconds to wait after continuous failures before attempting an adapter reset.
    -   Default: `30` seconds
-   `-TestInterval`: Specifies the time in seconds between connectivity tests.
    -   Default: `5` seconds

### Examples

1.  Run the script with default settings, monitoring the first available network adapter:

    ```powershell
    .\NetworkConnectivityWatchdog.ps1
    ```

2.  Monitor Ethernet adapters, wait 60 seconds of failures before reset, and test every 10 seconds:

    ```powershell
    .\NetworkConnectivityWatchdog.ps1 -AdapterPattern "Ethernet*" -FailureThreshold 60 -TestInterval 10
    ```

## Configuration

### Adapter Pattern

The `AdapterPattern` parameter is crucial for specifying which network adapter the script should monitor. You can find the name of your network adapter using the `Get-NetAdapter` command in PowerShell. Use wildcards (`*`) to match a pattern of adapter names.

Example:

```powershell
Get-NetAdapter | Where-Object {$_.Status -eq "Up"}
```

This command lists all active network adapters. Identify the name of the adapter you want to monitor and use it in the `AdapterPattern` parameter.

### Test Targets

The script uses a set of predefined test targets to check network connectivity. These targets include pinging well-known DNS servers and resolving common domain names. You can customize these targets by modifying the `$script:testTargets` variable within the script.

```powershell
$script:testTargets = @{
    PingTargets = @(
        @{ Name = "Google DNS"; Address = "8.8.8.8" },
        @{ Name = "Cloudflare DNS"; Address = "1.1.1.1" },
        @{ Name = "Default Gateway"; Address = (Get-NetRoute |
            Where-Object DestinationPrefix -eq '0.0.0.0/0' |
            Select-Object -First 1 -ExpandProperty NextHop) }
    )
    DnsTargets = @(
        @{ Name = "Google"; Address = "www.google.com" },
        @{ Name = "Microsoft"; Address = "www.microsoft.com" }
    )
}
```

You can add, remove, or modify these targets to suit your specific network environment.

## Contributing

Contributions are welcome! If you find a bug or have a feature request, please open an issue or submit a pull request on the SysAdminPSSorcery repo. https://github.com/ghostinator/SysAdminPSSorcery/NetworkConnectivityWatchdog

Please follow these guidelines when contributing:

-   Fork the repository.
-   Create a new branch for your feature or bug fix.
-   Make your changes and test them thoroughly.
-   Submit a pull request with a clear description of your changes.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
