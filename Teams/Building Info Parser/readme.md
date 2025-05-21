# Microsoft Teams Network Information Parser Tool

**Version:** 1.0 (Based on script development ending May 21, 2025)
**Purpose:** To assist network and Teams administrators in parsing and transforming network-related information, particularly from ITGlue exports, into a format suitable for uploading to the Microsoft Teams Call Quality Dashboard (CQD) as Building/Location data.

![alt text](<../../../../../Application Support/CleanShot/media/media_3h4izazPYo/CleanShot 2025-05-21 at 13.44.36.png>)

## Overview

The Teams CQD requires specific CSV file formats for uploading custom data, such as building and subnet information. This PowerShell script provides a graphical user interface (GUI) to:

1.  **Load and Process ITGlue Location Data:** Parses a CSV export of location details from ITGlue.
2.  **Process ITGlue LAN Data & Merge:** Parses a CSV export of LAN details from ITGlue and merges it with the loaded location data using a common location identifier.
3.  **Format for Teams CQD:** Populates a grid with the combined information, ready for review and export into the Teams CQD Building/Location CSV format.
4.  **Other Data Types (Under Construction):** Includes tabs for general "Network Information" and "Endpoint Information," but these functionalities are currently under construction.

## Features

* GUI for ease of use.
* **Two-Step CSV Processing for Building Information:**
    * Dedicated input for ITGlue Location Export CSV (via paste or file load).
    * Main input for ITGlue LAN Export CSV.
    * Automatic merging of LAN data with corresponding Location data.
* Flexible parsing of CSV data, attempting to map common ITGlue headers.
* Normalization of location names for improved matching between the two CSVs (handles extra quotes and spaces).
* Duplicate subnet detection and resolution assistant for Building and Network information.
* Data preview functionality (for data pasted into the main input box for supported types).
* Export to CSV format compatible with Teams CQD building/location uploads (no header row in the output CSV).

## Prerequisites

* Windows Operating System.
* PowerShell 5.1 or higher (the script is designed to be compatible with PowerShell 5.1).
* .NET Framework (usually available by default on Windows) for Windows Forms GUI.

## Preparing Your ITGlue CSV Exports

For the **Building Information (Merge)** functionality, you will need two separate CSV exports from ITGlue:

1.  **ITGlue Location Export CSV:**
    * This file should contain the detailed address information for your locations/buildings.
    * **Crucial Columns Expected by the Script:**
        * `name`: The primary name of the location (e.g., "Main Office Building A"). This is used as the **key** to link with the LAN data.
        * `address_1`: The primary street address line. This will be appended to the `name` to form the `BuildingName` in the output grid (e.g., "Main Office Building A (123 Main St)").
        * `city`: The city.
        * `region_name`: The state or province (e.g., "Indiana", "California"). This maps to the "State" field.
        * `country_name`: The country (e.g., "United States", "Canada"). This maps to the "Country" field.
        * `postal_code`: The zip or postal code. This maps to the "ZipCode" field.
    * Export all available columns from ITGlue Locations for best results, ensuring these key headers are present.

2.  **ITGlue LAN Export CSV:**
    * This file should contain details about your network segments (VLANs, subnets).
    * **Crucial Columns Expected by the Script:**
        * `location` (or `site`, `building`, `name`): This field **must contain a name that exactly matches (after normalization) a `name` from your Location Export CSV.** This is how the script links a LAN to its physical location details. Ensure consistency in ITGlue.
        * `subnet`: The subnet in CIDR notation (e.g., `192.168.1.0/24`). This will be parsed into `NetworkIP` and `NetworkRange`.
        * `name_vlan_name` (or `vlan_name`, `networkname`, or a `description` column that clearly indicates the VLAN/network purpose): This will be used for the `NetworkName`.
    * Optional but helpful columns (the script will use defaults if these are missing):
        * `insidecorp`: (1 for internal, 0 for guest/external - script attempts to infer if not present).
        * `expressroute`: (1 if applicable, 0 otherwise).
        * `vpn`: (1 if this network is for VPN, 0 otherwise).
    * When exporting LANs from ITGlue, customize the "visible columns" to include the `location` (or equivalent linking field) and any other relevant network details.

**Tip for ITGlue Exports:** Ensure you are exporting "All" records and not just the currently visible page if you have many entries.

## How to Use the Script

1.  **Launch the Script:** Save the PowerShell script to a `.ps1` file (e.g., `TeamsCQD_Parser.ps1`) and run it from a PowerShell console (`powershell.exe -File .\TeamsCQD_Parser.ps1`) or by right-clicking and selecting "Run with PowerShell". Running from a console is recommended to see any debug `Write-Host` messages if troubleshooting.

2.  **Workflow for Building Information (Two-CSV Merge for Teams CQD):**
    * This is the primary and most developed feature of the script.
    * **Step 1: Load/Paste Location Info CSV Data**
        * In the GUI, locate **"Box 1: For Building Info Merge: Paste ITGlue 'Location Export' CSV here..."**.
        * **Option A (Paste):** Copy the entire content (including headers) of your ITGlue **Location Export CSV** and paste it into this textbox. Then click the **"Process Pasted Location Data"** button.
        * **Option B (Load File):** Click the **"OR Load Location File..."** button. An "Open File" dialog will appear. Select your ITGlue **Location Export CSV** file.
        * A status message and a pop-up will confirm how many location entries were loaded. This data is now ready for merging.

    * **Step 2: Paste LAN Info CSV Data**
        * Ensure the **"Building Information (Merge)"** radio button is selected in the "Select Data Type..." group box.
        * In the GUI, locate **"Box 2: Paste ITGlue 'LAN Export' CSV (for Building merge)..."**.
        * Copy the entire content (including headers) of your ITGlue **LAN Export CSV** and paste it into this textbox. (You can also use the "Paste to Main" button, which pastes clipboard content into this box).

    * **Step 3: Process the LAN Data (and Merge)**
        * Click the **"Process Main"** button.
        * The script will parse the LAN data and attempt to merge it with the loaded Location data.
        * The "Building Information" tab at the bottom will be populated. Check console output for debug messages on matching.

    * **Step 4: Review and Edit Data**
        * Go to the "Building Information" tab.
        * Review the populated grid. Manually edit any cell if needed.
        * **Duplicate Subnet Handling:** If duplicate `NetworkIP` entries are found, a dialog will allow you to select which entry to keep. You must select an item before "Keep Selected" is enabled. Closing this dialog with "X" will prompt for confirmation to skip resolving that specific duplicate set.

    * **Step 5: Export the CSV for Teams CQD**
        * Once satisfied, click **"Export CSV"**.
        * Save the file (e.g., `TeamsCQD_Buildings.csv`). The exported CSV will **not** include a header row.

3.  **Using for "Network Information" or "Endpoint Information" Data Types:**
    * **Please Note:** The parsing and processing logic for the "Network Information" and "Endpoint Information" tabs are **currently under construction and should not be expected to work reliably or produce complete results for Teams CQD uploads.** While the UI elements are present, the primary focus of development has been the "Building Information (Merge)" workflow.
    * If you choose to experiment: Select the appropriate radio button, paste data into Box 2, and click "Process Main". The respective tab will be populated based on the existing parsing logic.

## Current Status / Limitations

* **Primary Functionality:** The script is primarily developed and tested for the "Building Information (Merge)" workflow using ITGlue Location and LAN CSV exports to generate a CSV for Teams CQD.
* **Network & Endpoint Info:** The "Network Information" and "Endpoint Information" parsing functionalities are considered **under construction/experimental.** They may parse some data but have not been fully developed or tested for creating Teams CQD-ready files. Use with caution.
* **Matching:** The success of the Building Information merge depends heavily on the consistency of location names between your ITGlue LAN `location` field (or equivalent) and the Location `name` field.
* **Teams CQD Geo-Region:** The "Region" column in the exported Building CSV (referring to Teams Geo-Regions like North America, EMEA, APAC) is currently exported as blank. This may need to be manually added to the CSV if required.

## Output CSV Format (Building Information for Teams CQD)

The script exports the "Building Information" data with the following columns, in order, without a header row:
`NetworkIP,NetworkName,NetworkRange,BuildingName,OwnershipType,BuildingType,BuildingOfficeType,City,ZipCode,Country,State,Region,InsideCorp,ExpressRoute,VPN`

* `OwnershipType`, `BuildingType`, `BuildingOfficeType`, and `Region` (Teams Geo-Region) are currently exported as blank strings by this script.

## Disclaimer

This script is provided as-is. Always test with sample data and back up your original CSV files before relying on its output for critical production uploads. The accuracy of the output depends on the quality and consistency of your input data from ITGlue.