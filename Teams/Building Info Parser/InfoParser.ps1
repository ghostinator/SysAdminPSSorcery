 # Microsoft Teams Network Information Parser Tool
# Created: May 20, 2025
# Last Revised: May 21, 2025

# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global variable to store parsed ITGlue Location Info
$Global:ProcessedLocationData = [ordered]@{}

# --- Helper Function to Normalize Lookup Keys (Improved) ---
function Normalize-LookupKey {
    param([string]$RawValue)

    if ([string]::IsNullOrWhiteSpace($RawValue)) {
        return $null 
    }

    $cleanedValue = $RawValue.Trim()

    # Repeatedly strip pairs of leading/trailing quotes
    while (($cleanedValue.Length -ge 2) -and `
           (($cleanedValue.StartsWith('"') -and $cleanedValue.EndsWith('"')) -or `
            ($cleanedValue.StartsWith("'") -and $cleanedValue.EndsWith("'")))) {
        $cleanedValue = $cleanedValue.Substring(1, $cleanedValue.Length - 2).Trim() 
    }
    
    # Normalize internal whitespace (replace multiple spaces with a single space)
    $cleanedValue = ($cleanedValue -replace '\s+', ' ').Trim() # Trim again after replacing spaces
    
    return $cleanedValue.ToLower() 
}

# --- Helper Function for Initial Parse Setup (for LAN/general CSV) ---
function Get-InputParseSetup {
    param ([string]$InputText)
    
    #Write-Host "DEBUG [Get-InputParseSetup]: InputText Length $($InputText.Length)"
    $lines = $InputText -split "`r?`n"
    #Write-Host "DEBUG [Get-InputParseSetup]: Lines count $($lines.Count)"

    $isCSV = $false
    $separator = "," # Default separator

    if ($lines.Count -gt 0 -and (-not [string]::IsNullOrWhiteSpace($lines[0])) -and $lines[0] -match '(,|;|\t)') {
        #Write-Host "DEBUG [Get-InputParseSetup]: First line '$($lines[0])' is not whitespace and contains a delimiter."
        $isCSV = $true
        if ($lines[0] -match "\t") { $separator = "`t" }
        elseif ($lines[0] -match ";") { $separator = ";" }
    } else {
        #Write-Host "DEBUG [Get-InputParseSetup]: First line '$($lines[0])' either whitespace or no delimiter found. isCSV=$isCSV"
    }
    #Write-Host "DEBUG [Get-InputParseSetup]: Returning IsCSV=$isCSV, Separator='$separator'"
    return @{Lines = $lines; IsCSV = $isCSV; Separator = $separator }
}

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Microsoft Teams Network Information Parser"
$form.Size = New-Object System.Drawing.Size(920, 930) 
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# Create a label for instructions
$instructionLabel = New-Object System.Windows.Forms.Label
$instructionLabel.Location = New-Object System.Drawing.Point(20, 20)
$instructionLabel.Size = New-Object System.Drawing.Size(860, 20)
$instructionLabel.Text = "For 'Building Information', first load/paste Location Info CSV, then paste LAN Info CSV and process."
$instructionLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($instructionLabel)

# Create data type selection
$dataTypeGroupBox = New-Object System.Windows.Forms.GroupBox
$dataTypeGroupBox.Location = New-Object System.Drawing.Point(20, 50) 
$dataTypeGroupBox.Size = New-Object System.Drawing.Size(530, 60) 
$dataTypeGroupBox.Text = "Select Data Type for Main Input Below (Box 2)"
$dataTypeGroupBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($dataTypeGroupBox)

$buildingRadioButton = New-Object System.Windows.Forms.RadioButton
$buildingRadioButton.Location = New-Object System.Drawing.Point(20, 25)
$buildingRadioButton.Size = New-Object System.Drawing.Size(160, 20) 
$buildingRadioButton.Text = "Building Information (Merge)"
$buildingRadioButton.Checked = $true
$dataTypeGroupBox.Controls.Add($buildingRadioButton)

$networkRadioButton = New-Object System.Windows.Forms.RadioButton
$networkRadioButton.Location = New-Object System.Drawing.Point(190, 25) 
$networkRadioButton.Size = New-Object System.Drawing.Size(150, 20)
$networkRadioButton.Text = "Network Information"
$dataTypeGroupBox.Controls.Add($networkRadioButton)

$endpointRadioButton = New-Object System.Windows.Forms.RadioButton
$endpointRadioButton.Location = New-Object System.Drawing.Point(350, 25) 
$endpointRadioButton.Size = New-Object System.Drawing.Size(150, 20)
$endpointRadioButton.Text = "Endpoint Information"
$dataTypeGroupBox.Controls.Add($endpointRadioButton)


# --- Location Info CSV Input Area ---
$locationCsvPasteLabel = New-Object System.Windows.Forms.Label
$locationCsvPasteLabel.Location = New-Object System.Drawing.Point(20, 120) 
$locationCsvPasteLabel.Size = New-Object System.Drawing.Size(860, 20)
$locationCsvPasteLabel.Text = "1. For Building Info Merge: Paste ITGlue 'Location Export' CSV here (with all columns like name, city, postal_code):"
$locationCsvPasteLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($locationCsvPasteLabel)

$locationCsvPasteTextBox = New-Object System.Windows.Forms.TextBox
$locationCsvPasteTextBox.Location = New-Object System.Drawing.Point(20, 145) 
$locationCsvPasteTextBox.Size = New-Object System.Drawing.Size(860, 120)   
$locationCsvPasteTextBox.Multiline = $true
$locationCsvPasteTextBox.ScrollBars = "Vertical"
$locationCsvPasteTextBox.AcceptsReturn = $true
$locationCsvPasteTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($locationCsvPasteTextBox)

$processPastedLocationDataButton = New-Object System.Windows.Forms.Button
$processPastedLocationDataButton.Location = New-Object System.Drawing.Point(20, 270) 
$processPastedLocationDataButton.Size = New-Object System.Drawing.Size(220, 30) 
$processPastedLocationDataButton.Text = "Process Pasted Location Data"
$processPastedLocationDataButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($processPastedLocationDataButton)

$loadLocationDataButton = New-Object System.Windows.Forms.Button
$loadLocationDataButton.Location = New-Object System.Drawing.Point(250, 270) 
$loadLocationDataButton.Size = New-Object System.Drawing.Size(180, 30)
$loadLocationDataButton.Text = "OR Load Location File..."
$loadLocationDataButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($loadLocationDataButton)

$locationDataStatusLabel = New-Object System.Windows.Forms.Label
$locationDataStatusLabel.Location = New-Object System.Drawing.Point(440, 275) 
$locationDataStatusLabel.Size = New-Object System.Drawing.Size(440, 20)
$locationDataStatusLabel.Text = "No location data loaded."
$locationDataStatusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($locationDataStatusLabel)

# --- Main Input Area (for LAN Info CSV or other data types) ---
$mainInputLabel = New-Object System.Windows.Forms.Label
$mainInputLabel.Location = New-Object System.Drawing.Point(20, 310) 
$mainInputLabel.Size = New-Object System.Drawing.Size(860, 20)
$mainInputLabel.Text = "2. Paste ITGlue 'LAN Export' CSV (for Building merge) or other selected data type below:"
$mainInputLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($mainInputLabel)

$inputTextBox = New-Object System.Windows.Forms.TextBox 
$inputTextBox.Location = New-Object System.Drawing.Point(20, 335) 
$inputTextBox.Size = New-Object System.Drawing.Size(860, 150)
$inputTextBox.Multiline = $true
$inputTextBox.ScrollBars = "Vertical"
$inputTextBox.AcceptsReturn = $true
$inputTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($inputTextBox)

# Main Action buttons 
$pasteButton = New-Object System.Windows.Forms.Button; $pasteButton.Location = New-Object System.Drawing.Point(20, 495); $pasteButton.Size = New-Object System.Drawing.Size(100, 30); $pasteButton.Text = "Paste to Main"; $pasteButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left; $form.Controls.Add($pasteButton) 
$detectButton = New-Object System.Windows.Forms.Button; $detectButton.Location = New-Object System.Drawing.Point(130, 495); $detectButton.Size = New-Object System.Drawing.Size(100, 30); $detectButton.Text = "Detect Format"; $detectButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left; $form.Controls.Add($detectButton)
$processButton = New-Object System.Windows.Forms.Button; $processButton.Location = New-Object System.Drawing.Point(240, 495); $processButton.Size = New-Object System.Drawing.Size(100, 30); $processButton.Text = "Process Main"; $processButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left; $form.Controls.Add($processButton) 
$form.AcceptButton = $processButton 
$previewButton = New-Object System.Windows.Forms.Button; $previewButton.Location = New-Object System.Drawing.Point(350, 495); $previewButton.Size = New-Object System.Drawing.Size(100, 30); $previewButton.Text = "Preview Main"; $previewButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left; $form.Controls.Add($previewButton) 
$clearButton = New-Object System.Windows.Forms.Button; $clearButton.Location = New-Object System.Drawing.Point(460, 495); $clearButton.Size = New-Object System.Drawing.Size(100, 30); $clearButton.Text = "Clear All"; $clearButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left; $form.Controls.Add($clearButton) 
$exportButton = New-Object System.Windows.Forms.Button; $exportButton.Location = New-Object System.Drawing.Point(780, 495); $exportButton.Size = New-Object System.Drawing.Size(100, 30); $exportButton.Text = "Export CSV"; $exportButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right; $form.Controls.Add($exportButton)

# Create a status strip
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready. For Building Info, load Location CSV (Box 1), then LAN CSV (Box 2)."
$statusStrip.Items.Add($statusLabel)
$statusStrip.Dock = [System.Windows.Forms.DockStyle]::Bottom 
$form.Controls.Add($statusStrip)

# Create a tab control for different data types
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(20, 535) 
$tabControl.Size = New-Object System.Drawing.Size(860, 305) 
$tabControl.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($tabControl)

# Building Information Tab (11 columns now)
$buildingTab = New-Object System.Windows.Forms.TabPage
$buildingTab.Text = "Building Information"
$tabControl.Controls.Add($buildingTab)

$buildingDataGrid = New-Object System.Windows.Forms.DataGridView
$buildingDataGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
$buildingDataGrid.AllowUserToAddRows = $true
$buildingDataGrid.AllowUserToDeleteRows = $true
$buildingDataGrid.ColumnHeadersHeightSizeMode = "AutoSize"
$buildingDataGrid.AutoSizeColumnsMode = "DisplayedCells" 
$buildingDataGrid.RowHeadersWidth = 30
$buildingDataGrid.ColumnCount = 11
$buildingDataGrid.Columns[0].Name = "NetworkIP"
$buildingDataGrid.Columns[1].Name = "NetworkName"
$buildingDataGrid.Columns[2].Name = "NetworkRange"
$buildingDataGrid.Columns[3].Name = "BuildingName"    
$buildingDataGrid.Columns[4].Name = "City"           
$buildingDataGrid.Columns[5].Name = "State"          
$buildingDataGrid.Columns[6].Name = "ZipCode"        
$buildingDataGrid.Columns[7].Name = "Country"        
$buildingDataGrid.Columns[8].Name = "InsideCorp"
$buildingDataGrid.Columns[9].Name = "ExpressRoute"
$buildingDataGrid.Columns[10].Name = "VPN"
$buildingTab.Controls.Add($buildingDataGrid)

# Network Information Tab (remains 6 columns)
$networkTab = New-Object System.Windows.Forms.TabPage
$networkTab.Text = "Network Information"
$tabControl.Controls.Add($networkTab)

$networkDataGrid = New-Object System.Windows.Forms.DataGridView
$networkDataGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
$networkDataGrid.AllowUserToAddRows = $true
$networkDataGrid.AllowUserToDeleteRows = $true
$networkDataGrid.ColumnHeadersHeightSizeMode = "AutoSize"
$networkDataGrid.AutoSizeColumnsMode = "Fill" 
$networkDataGrid.RowHeadersWidth = 30
$networkDataGrid.ColumnCount = 6
$networkDataGrid.Columns[0].Name = "NetworkRegion"
$networkDataGrid.Columns[1].Name = "NetworkSite"
$networkDataGrid.Columns[2].Name = "Subnet"
$networkDataGrid.Columns[3].Name = "MaskBits"
$networkDataGrid.Columns[4].Name = "Description"
$networkDataGrid.Columns[5].Name = "ExpressRoute"
$networkTab.Controls.Add($networkDataGrid)

# Endpoint Information Tab (remains 5 columns)
$endpointTab = New-Object System.Windows.Forms.TabPage
$endpointTab.Text = "Endpoint Information"
$tabControl.Controls.Add($endpointTab)

$endpointDataGrid = New-Object System.Windows.Forms.DataGridView
$endpointDataGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
$endpointDataGrid.AllowUserToAddRows = $true
$endpointDataGrid.AllowUserToDeleteRows = $true
$endpointDataGrid.ColumnHeadersHeightSizeMode = "AutoSize"
$endpointDataGrid.AutoSizeColumnsMode = "Fill" 
$endpointDataGrid.RowHeadersWidth = 30
$endpointDataGrid.ColumnCount = 5
$endpointDataGrid.Columns[0].Name = "EndpointName"
$endpointDataGrid.Columns[1].Name = "MacAddress"
$endpointDataGrid.Columns[2].Name = "Manufacturer"
$endpointDataGrid.Columns[3].Name = "Model"
$endpointDataGrid.Columns[4].Name = "Type"
$endpointTab.Controls.Add($endpointDataGrid)

# --- Core Function to Process Location CSV Content ---
function Process-LocationCsvContent {
    param(
        [string]$CsvContent,
        [System.Windows.Forms.ToolStripStatusLabel]$UiStatusLabel, 
        [System.Windows.Forms.Label]$UiLocationDataStatusLabel 
    )

    try {
        if ([string]::IsNullOrWhiteSpace($CsvContent)) {
            [System.Windows.Forms.MessageBox]::Show("No CSV content provided for locations.", "Empty Input", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return $false
        }

        $tempLinesForDelimiterCheck = $CsvContent -split "`r?`n"
        $firstPastedLine = ""
        if ($tempLinesForDelimiterCheck.Count -gt 0) {
            $firstPastedLine = $tempLinesForDelimiterCheck[0]
        }
        
        $delimiterForPasted = "," 
        if (-not [string]::IsNullOrWhiteSpace($firstPastedLine)) { 
            if ($firstPastedLine -match "\t") { $delimiterForPasted = "`t" } 
            elseif ($firstPastedLine -match ";") { $delimiterForPasted = ";" }
        }
        
        $csvData = $CsvContent | ConvertFrom-Csv -Delimiter $delimiterForPasted
        
        $Global:ProcessedLocationData.Clear() 
        $loadedCount = 0
        # These are the specific headers expected from an ITGlue Location Export
        $requiredHeaders = @("name", "address_1", "city", "region_name", "country_name", "postal_code") 
        
        if (-not $csvData -or $csvData.Count -eq 0) {
             [System.Windows.Forms.MessageBox]::Show("Location CSV content appears empty or could not be parsed into rows after attempting to read with delimiter '$delimiterForPasted'.", "Parsing Issue", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($UiLocationDataStatusLabel) { $UiLocationDataStatusLabel.Text = "Location data parse error."}
            return $false
        }
        
        $firstRowActualProps = $csvData[0].PSObject.Properties
        $propsInFileLower = $firstRowActualProps | ForEach-Object { $_.Name.Trim().ToLower() }
        
        $missingHeadersInCsv = $requiredHeaders | Where-Object { $propsInFileLower -notcontains $_.ToLower() }
        if ($missingHeadersInCsv.Count -gt 0) {
            [System.Windows.Forms.MessageBox]::Show("The provided Location CSV data may be missing some expected headers (or they have unexpected formatting): $($missingHeadersInCsv -join ', '). Some location details might not be fully parsed. Please ensure it's the ITGlue Location export with standard column names for best results.", "Potential Location CSV Format Issue", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }

        # Define the helper to take the row object explicitly
        $GetPropValueFromRow = {
            param($PassedRowItem, $PropertyNameToCheck)
            if ($null -eq $PassedRowItem) { return $null }
            $LocalRowProps = $PassedRowItem.PSObject.Properties
            if ($null -eq $LocalRowProps) { return $null } 
            $ActualProperty = $LocalRowProps | Where-Object {$_.Name.Trim().ToLower() -eq $PropertyNameToCheck.ToLower()} | Select-Object -First 1
            if ($ActualProperty) {
                return $PassedRowItem.$($ActualProperty.Name)
            }
            return $null
        }

        foreach ($rowItemInLoop in $csvData) {
            if ($null -eq $rowItemInLoop) { Write-Warning "Skipping null row in Location CSV."; continue }

            $locationNameRaw = & $GetPropValueFromRow -PassedRowItem $rowItemInLoop -PropertyNameToCheck "name"
            # Your debug line (keep it if helpful)
            Write-Host "DEBUG [LocationCSV]: Raw='$(if($locationNameRaw){$locationNameRaw}else{''})', NormalizedKeyStored='$(Normalize-LookupKey -RawValue $locationNameRaw)'" 

            $normalizedKey = Normalize-LookupKey -RawValue $locationNameRaw 

            if (-not [string]::IsNullOrWhiteSpace($normalizedKey)) { 
                if ($Global:ProcessedLocationData[$normalizedKey] -eq $null) { 
                    $Global:ProcessedLocationData[$normalizedKey] = @{
                        OriginalName = $locationNameRaw 
                        Address1     = & $GetPropValueFromRow -PassedRowItem $rowItemInLoop -PropertyNameToCheck "address_1"
                        City         = & $GetPropValueFromRow -PassedRowItem $rowItemInLoop -PropertyNameToCheck "city"
                        State        = & $GetPropValueFromRow -PassedRowItem $rowItemInLoop -PropertyNameToCheck "region_name"
                        Country      = & $GetPropValueFromRow -PassedRowItem $rowItemInLoop -PropertyNameToCheck "country_name"
                        Zip          = & $GetPropValueFromRow -PassedRowItem $rowItemInLoop -PropertyNameToCheck "postal_code"
                    }
                    $loadedCount++
                } else {
                    Write-Warning "Duplicate location name found in Location CSV (after normalization): '$locationNameRaw' (normalized to: '$normalizedKey'). First entry was kept."
                }
            }
        }
        if ($UiLocationDataStatusLabel) { $UiLocationDataStatusLabel.Text = "Location data loaded for $loadedCount sites." }
        if ($UiStatusLabel) { $UiStatusLabel.Text = "Location Info processed ($loadedCount sites). Ready for LAN Info." }
        if ($loadedCount -gt 0) {
            [System.Windows.Forms.MessageBox]::Show("$loadedCount location entries processed successfully from input.", "Location Data Processed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } elseif ($csvData.Count -gt 0) { 
             [System.Windows.Forms.MessageBox]::Show("Could not extract valid location names from the provided Location CSV data. Ensure the 'name' column is present, populated, and headers are standard.", "Location Name Missing or Header Issue", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
        return $true
    } catch {
        $Global:ProcessedLocationData.Clear()
        if ($UiLocationDataStatusLabel) { $UiLocationDataStatusLabel.Text = "Error processing location data."}
        [System.Windows.Forms.MessageBox]::Show("Error parsing Location Info CSV content: $($_.Exception.Message)`nAt line $($_.InvocationInfo.ScriptLineNumber)", "Parsing Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
}

# --- Function to Load Location Info from File ---
function Load-AndParseLocationInfoCSVFromFile {
    param(
        [System.Windows.Forms.ToolStripStatusLabel]$UiStatusLabel, 
        [System.Windows.Forms.Label]$UiLocationDataStatusLabel 
    )
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = "Select ITGlue Location Info CSV File"
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $filePath = $openFileDialog.FileName
        try {
            $locationFileContent = Get-Content -Path $filePath -Raw
            Process-LocationCsvContent -CsvContent $locationFileContent -UiStatusLabel $UiStatusLabel -UiLocationDataStatusLabel $UiLocationDataStatusLabel
        } catch {
            $Global:ProcessedLocationData.Clear() 
            if ($UiLocationDataStatusLabel) { $UiLocationDataStatusLabel.Text = "Error loading location file."}
            [System.Windows.Forms.MessageBox]::Show("Error reading Location Info CSV file: $($_.Exception.Message)", "File Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}

function Detect-DataFormat {
    param ([string]$InputText)
    $dataType = "Building" 
    if ([string]::IsNullOrWhiteSpace($InputText)) { return $dataType }
    $lines = $InputText -split "`r?`n"
    if ($lines.Count -eq 0) { return $dataType} # Handle empty lines case
    
    $firstLine = $lines[0]
    # Consider first few lines for more robust detection if needed
    # $headerSample = $lines | Select-Object -First 5 -join " " 

    $isCSV = $false
    if (-not [string]::IsNullOrWhiteSpace($firstLine) -and $firstLine -match '(,|;|\t)') { $isCSV = $true }

    if ($isCSV) {
        if ($firstLine -match '(?i)(subnet|vlan_id|networkip|networkrange|maskbits|gateway|dhcp_server)') {
            $dataType = "Network" # Or Building if location also present
            if ($firstLine -match '(?i)location|site|buildingname'){ $dataType = "Building"}
        }
        elseif ($firstLine -match '(?i)(macaddress|mac_address|device|manufacturer|model)') {
            $dataType = "Endpoint"
        }
        elseif ($firstLine -match '(?i)(name_vlan_name|location|site|buildingname|address_1|city|postal_code|region_name)') {
             # Could be LAN info for buildings, or pure Location info
            $dataType = "Building" 
        }
        # Fallback content checks if headers are ambiguous
        elseif ($InputText -match '\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}/\d{1,2}') { # CIDR
            if ($InputText -match '(?i)location|site|building') {$dataType = "Building"} else {$dataType = "Network"}
        }
        elseif ($InputText -match '([0-9A-Fa-f]{2}[:-]){5}[0-9A-Fa-f]{2}') { 
            $dataType = "Endpoint"
        }
    } else { # Unstructured text
        if ($InputText -match '(?i)(subnet|vlan|network|gateway|dhcp|\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}/\d{1,2})') {
             if ($InputText -match '(?i)building|location|site|address|city|country') {$dataType = "Building"}
             else {$dataType = "Network"}
        }
        elseif ($InputText -match '(?i)(mac address|manufacturer|model|device type|([0-9A-Fa-f]{2}[:-]){5}[0-9A-Fa-f]{2})') {
            $dataType = "Endpoint"
        }
        elseif ($InputText -match '(?i)(building|location|site|address|city|country)') {
            $dataType = "Building"
        }
    }
    return $dataType
}

function Validate-NetworkData {
    param ([System.Windows.Forms.DataGridView]$DataGrid)
    $hasErrors = $false
    $errorMessages = @()
    foreach ($row in $DataGrid.Rows) {
        if ($row.IsNewRow) { continue }
        
        $row.Cells["Subnet"].Style.BackColor = [System.Drawing.Color]::White 
        $row.Cells["MaskBits"].Style.BackColor = [System.Drawing.Color]::White 

        $subnetCell = $row.Cells["Subnet"]
        $subnet = $subnetCell.Value
        if ($subnet -and -not ($subnet -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$')) {
            $subnetCell.Style.BackColor = [System.Drawing.Color]::LightPink
            $errorMessages += "Invalid subnet format at row $($row.Index + 1): $subnet"
            $hasErrors = $true
        }
        
        $maskBitsCell = $row.Cells["MaskBits"]
        $maskBits = $maskBitsCell.Value
        if ($maskBits) {
            try {
                $maskIntValue = $null
                if ($maskBits -isnot [string] -or -not [string]::IsNullOrWhiteSpace($maskBits)) {
                    $maskIntValue = [int]$maskBits
                }

                if ($maskIntValue -ne $null) {
                    if ($maskIntValue -lt 0 -or $maskIntValue -gt 32) {
                        $maskBitsCell.Style.BackColor = [System.Drawing.Color]::LightPink
                        $errorMessages += "Invalid subnet mask (should be 0-32) at row $($row.Index + 1): $maskBits"
                        $hasErrors = $true
                    }
                    elseif ($maskIntValue -lt 8 -and $maskIntValue -ge 0) { 
                        $maskBitsCell.Style.BackColor = [System.Drawing.Color]::LightYellow
                        $errorMessages += "Warning: Very broad subnet mask at row $($row.Index + 1): /$maskBits"
                    }
                } elseif(-not [string]::IsNullOrWhiteSpace($maskBits)) { # If not blank but couldn't convert to int
                    $maskBitsCell.Style.BackColor = [System.Drawing.Color]::LightPink
                    $errorMessages += "Invalid subnet mask format (non-integer) at row $($row.Index + 1): $maskBits"
                    $hasErrors = $true
                }
            }
            catch {
                $maskBitsCell.Style.BackColor = [System.Drawing.Color]::LightPink
                $errorMessages += "Invalid subnet mask format (conversion error) at row $($row.Index + 1): $maskBits"
                $hasErrors = $true
            }
        }
    }
    return @{HasErrors = $hasErrors; ErrorMessages = $errorMessages}
}

function Handle-DuplicateSubnets { 
    param ([System.Windows.Forms.DataGridView]$DataGrid, [string]$SubnetColumnName )
    $subnetTracker = @{}
    $duplicatesFound = $false
    $duplicateRowGroups = @{} 
    
    foreach ($r_hl in $DataGrid.Rows) { if ($r_hl.IsNewRow){continue}; $r_hl.DefaultCellStyle.BackColor = [System.Drawing.Color]::White }

    for ($i_hl = 0; $i_hl -lt $DataGrid.Rows.Count; $i_hl++) {
        $row_hl = $DataGrid.Rows[$i_hl]
        if ($row_hl.IsNewRow){continue}
        $subnetCell_hl = $row_hl.Cells[$SubnetColumnName]
        if ($null -ne $subnetCell_hl -and $null -ne $subnetCell_hl.Value -and -not([string]::IsNullOrWhiteSpace($subnetCell_hl.Value.ToString())) ) { # Ensure subnet value is not blank
            $subnet_hl = $subnetCell_hl.Value.ToString()
            if ($subnetTracker.ContainsKey($subnet_hl)) {
                $duplicatesFound = $true
                if (-not $duplicateRowGroups.ContainsKey($subnet_hl)){ 
                    $duplicateRowGroups[$subnet_hl] = @($subnetTracker[$subnet_hl]) 
                }
                $duplicateRowGroups[$subnet_hl] += $i_hl 
                
                if($subnetTracker[$subnet_hl] -lt $DataGrid.Rows.Count){ # Check index bounds
                    $DataGrid.Rows[$subnetTracker[$subnet_hl]].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightPink
                }
                $row_hl.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightPink
            } else {
                $subnetTracker[$subnet_hl] = $i_hl
            }
        }
    }
    
    if ($duplicatesFound){
        $initialPromptResult_hl = [System.Windows.Forms.MessageBox]::Show("Duplicate subnets were found. Microsoft Teams does not allow duplicate subnets across different locations. Would you like to resolve these duplicates now?","Duplicate Subnets Detected",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($initialPromptResult_hl -eq [System.Windows.Forms.DialogResult]::Yes){
            $rowsToRemoveIndexes_hl = New-Object System.Collections.Generic.List[int]
            foreach ($subnetKeyToProcess_hl in $duplicateRowGroups.Keys){
                $rowIndicesInGroup_hl = $duplicateRowGroups[$subnetKeyToProcess_hl] | Get-Unique
                if ($rowIndicesInGroup_hl.Count -le 1){continue} 
                $selectionForm_hl = New-Object System.Windows.Forms.Form; $selectionForm_hl.Text = "Duplicate Subnet: $subnetKeyToProcess_hl"; $selectionForm_hl.Size = New-Object System.Drawing.Size(600,400); $selectionForm_hl.StartPosition = "CenterParent"; $selectionForm_hl.Font = New-Object System.Drawing.Font("Segoe UI",10)
                $selectionLabel_hl = New-Object System.Windows.Forms.Label; $selectionLabel_hl.Location=New-Object System.Drawing.Point(20,20); $selectionLabel_hl.Size=New-Object System.Drawing.Size(560,40); $selectionLabel_hl.Text="Multiple entries found for subnet $subnetKeyToProcess_hl. Please select the one to keep (others will be removed):"; $selectionForm_hl.Controls.Add($selectionLabel_hl)
                $selectionList_hl = New-Object System.Windows.Forms.ListBox; $selectionList_hl.Location=New-Object System.Drawing.Point(20,70); $selectionList_hl.Size=New-Object System.Drawing.Size(560,240); $selectionList_hl.DisplayMember="DisplayText"; $selectionList_hl.ValueMember="RowIndex"; $selectionForm_hl.Controls.Add($selectionList_hl)
                $okButton_hl = New-Object System.Windows.Forms.Button; $okButton_hl.Location=New-Object System.Drawing.Point(250,320); $okButton_hl.Size=New-Object System.Drawing.Size(100,30); $okButton_hl.Text="Keep Selected"; $okButton_hl.DialogResult=[System.Windows.Forms.DialogResult]::OK; $okButton_hl.Enabled=$false; $selectionForm_hl.Controls.Add($okButton_hl); $selectionForm_hl.AcceptButton=$okButton_hl
                $selectionList_hl.add_SelectedIndexChanged({$okButton_hl.Enabled=($selectionList_hl.SelectedIndex -ne -1)})
                
                foreach ($rowIndex_hl in $rowIndicesInGroup_hl){
                    if($rowIndex_hl -ge $DataGrid.Rows.Count -or $DataGrid.Rows[$rowIndex_hl].IsNewRow){continue} 
                    $r_hl_item = $DataGrid.Rows[$rowIndex_hl]
                    $dispStr_hl="Row $($rowIndex_hl+1): "
                    if($DataGrid.Columns.Contains("BuildingName")){$dispStr_hl+="Network: $($r_hl_item.Cells["NetworkName"].Value), Building: $($r_hl_item.Cells["BuildingName"].Value)"}
                    elseif($DataGrid.Columns.Contains("NetworkSite")){$dispStr_hl+="Region: $($r_hl_item.Cells["NetworkRegion"].Value), Site: $($r_hl_item.Cells["NetworkSite"].Value), Desc: $($r_hl_item.Cells["Description"].Value)"}
                    else{$dispStr_hl+="Data: $($r_hl_item.Cells[0].Value)"} # Fallback, assumes first column has some data
                    $item_hl=[PSCustomObject]@{DisplayText=$dispStr_hl;RowIndex=$rowIndex_hl}; $selectionList_hl.Items.Add($item_hl)
                }

                $selectionForm_hl.add_FormClosing({param($s_hl,$e_hl) 
                    if($e_hl.CloseReason -eq [System.Windows.Forms.CloseReason]::UserClosing -and $s_hl.DialogResult -ne [System.Windows.Forms.DialogResult]::OK){
                        $cfmRes_hl=[System.Windows.Forms.MessageBox]::Show("You have not selected an entry to keep for subnet '$subnetKeyToProcess_hl'.`n`nIf you close this dialog, all listed entries for this subnet will be kept.`n`nAre you sure you want to skip resolving this specific duplicate set?","Confirm Skip Resolution",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question)
                        if($cfmRes_hl -eq [System.Windows.Forms.DialogResult]::No){$e_hl.Cancel=$true}
                    }
                })
                
                $diagResSel_hl=$selectionForm_hl.ShowDialog()
                if($diagResSel_hl -eq [System.Windows.Forms.DialogResult]::OK){
                    if($selectionList_hl.SelectedItem -ne $null){
                        $selIdxToKeep_hl=$selectionList_hl.SelectedItem.RowIndex
                        foreach($rIdxToRemCand_hl in $rowIndicesInGroup_hl){
                            if($rIdxToRemCand_hl -ne $selIdxToKeep_hl){
                                if(-not $rowsToRemoveIndexes_hl.Contains($rIdxToRemCand_hl)){$rowsToRemoveIndexes_hl.Add($rIdxToRemCand_hl)}
                            }
                        }
                    }else{ Write-Warning "Duplicate resolution for '$subnetKeyToProcess_hl': OK clicked but no item was selected. No changes made for this group." }
                }elseif($diagResSel_hl -eq [System.Windows.Forms.DialogResult]::Cancel){
                    Write-Host "User skipped resolution for subnet '$subnetKeyToProcess_hl'. All entries for this group will be kept."
                }
                $selectionForm_hl.Dispose()
            }
            $uniqueSortedIndexesToRemove_hl = $rowsToRemoveIndexes_hl | Sort-Object -Descending -Unique
            foreach($idxToRem_hl in $uniqueSortedIndexesToRemove_hl){
                 if($idxToRem_hl -lt $DataGrid.Rows.Count -and (-not $DataGrid.Rows[$idxToRem_hl].IsNewRow) ){$DataGrid.Rows.RemoveAt($idxToRem_hl)}
            }
            foreach($r_hl_final in $DataGrid.Rows){if($r_hl_final.IsNewRow){continue}; $r_hl_final.DefaultCellStyle.BackColor=[System.Drawing.Color]::White}
            return $true 
        }
    }
    return $false 
}

# Function to parse building information (EXPECTS LAN INFO CSV & MERGES WITH LOADED LOCATION DATA)
# Function to parse building information (EXPECTS LAN INFO CSV & MERGES WITH LOADED LOCATION DATA)
function Parse-BuildingInfo {
    param ([string]$InputText) 
    
    $buildingDataGrid.Rows.Clear()
    # Write-Host "DEBUG [Parse-BuildingInfo]: Called. InputText length: $($InputText.Length)" 
    $parseSetup = Get-InputParseSetup -InputText $InputText
    # Write-Host "DEBUG [Parse-BuildingInfo]: Get-InputParseSetup returned IsCSV: $($parseSetup.IsCSV), Separator: '$($parseSetup.Separator)'"
    
    $unmatchedLanLocations = 0
    $matchedLanLocations = 0

    if ($Global:ProcessedLocationData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Warning: No Location Data has been loaded. Building address details will likely be blank unless present in the LAN Info CSV itself. Use 'Load/Process Location Info' first for best results when merging Building Information.", "Location Data Missing For Merge", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }

    if ($parseSetup.IsCSV) {
        try {
            $lanCsvData = $InputText | ConvertFrom-Csv -Delimiter $parseSetup.Separator
            
            if ($null -eq $lanCsvData) {
                Write-Warning "LAN Info CSV parsing (ConvertFrom-Csv) resulted in null data. Input might be empty or malformed. No rows will be processed."
                $lanCsvData = @() 
            } elseif ($lanCsvData -is [System.Management.Automation.PSCustomObject]) {
                $lanCsvData = @($lanCsvData) 
            } elseif (-not ($lanCsvData -is [System.Array])) {
                $actualType = if ($null -eq $lanCsvData) { "null (unexpected)" } else { $lanCsvData.GetType().FullName }
                [System.Windows.Forms.MessageBox]::Show("LAN Info CSV data was not parsed into an expected array by ConvertFrom-Csv. Got: '$($actualType)'. Check CSV. Processing cannot continue.", "LAN CSV Parse Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return 
            }

            if ($lanCsvData.Count -eq 0 -and -not ([string]::IsNullOrWhiteSpace($InputText)) ) {
                 Write-Warning "LAN Info CSV parsing resulted in an empty dataset from ConvertFrom-Csv, though input text was not empty. Check CSV format."
            }

            foreach ($lanRowItem in $lanCsvData) {
                if ($null -eq $lanRowItem) { Write-Warning "Skipping a null row item in LAN CSV."; continue }

                $lanRawProps = $lanRowItem.PSObject.Properties
                if ($null -eq $lanRawProps) { Write-Warning "Skipping row with null PSObject.Properties in LAN CSV."; continue }
                
                # --- MODIFIED Logic to get property values from $lanRowItem ---
                $GetPropertyValueByName = {
                    param($RowObject, $PropertyNameToCheck)
                    $ActualProperty = $RowObject.PSObject.Properties | Where-Object {$_.Name.Trim().ToLower() -eq $PropertyNameToCheck.ToLower()} | Select-Object -First 1
                    if ($ActualProperty) {
                        return $RowObject.$($ActualProperty.Name)
                    }
                    return $null
                }
                # --- End MODIFIED Logic ---
                
                $networkIP = ""; $networkName = ""; $networkRange = ""
                $buildingNameForGrid = ""; $cityForGrid = ""; $stateForGrid = ""
                $zipCodeForGrid = ""; $countryForGrid = ""
                $insideCorpVal = "1"; $expressRouteVal = "0"; $vpnVal = "0"

                $subnetVal = & $GetPropertyValueByName -RowObject $lanRowItem -PropertyNameToCheck "subnet"
                if ($subnetVal -and ($subnetVal -is [string]) -and $subnetVal -match '(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})/(\d{1,2})') {
                    $networkIP = $matches[1]; $networkRange = $matches[2]
                } else {
                    $networkIPVal = & $GetPropertyValueByName -RowObject $lanRowItem -PropertyNameToCheck "networkip"
                    if($networkIPVal -is [string]){ $networkIP = $networkIPVal }
                }
                
                if ([string]::IsNullOrWhiteSpace($networkRange)) { 
                    $networkRangeVal = & $GetPropertyValueByName -RowObject $lanRowItem -PropertyNameToCheck "networkrange"
                    if($networkRangeVal -is [string]) { $networkRange = $networkRangeVal }
                }

                $lanNetworkNameCandidates = @('name_vlan_name', 'vlan_name', 'networkname', 'description')
                foreach($candidate in $lanNetworkNameCandidates){
                    $potentialName = & $GetPropertyValueByName -RowObject $lanRowItem -PropertyNameToCheck $candidate
                    if($potentialName -ne $null -and $potentialName -is [string]){ 
                        if($candidate -eq 'description' -and (($potentialName -match '(?i)vlan') -or ($potentialName -match '(?i)network'))){
                             $networkName = $potentialName; break
                        } elseif ($candidate -ne 'description'){
                             $networkName = $potentialName; break
                        }
                    }
                }
                $networkName = ($networkName -replace '"', '').Trim()
                                
                $lanLocationKeyRaw = ""
                $lanLocationKeyHeaders = @('location', 'site', 'building', 'name') 
                foreach($keyHeader in $lanLocationKeyHeaders){
                    $tempKeyRaw = & $GetPropertyValueByName -RowObject $lanRowItem -PropertyNameToCheck $keyHeader
                    if (-not [string]::IsNullOrWhiteSpace($tempKeyRaw) -and $tempKeyRaw -is [string]) { 
                        $lanLocationKeyRaw = $tempKeyRaw
                        break 
                    }
                }
                
                # Your existing debug lines (keep them active!)
                $normalizedLanLocationKey = Normalize-LookupKey -RawValue $lanLocationKeyRaw 
                Write-Host "DEBUG [LAN CSV]: RawLinkField='$(if($lanLocationKeyRaw){$lanLocationKeyRaw}else{''})', NormalizedLANKeyForLookup='$($normalizedLanLocationKey)'" 
                
                if ($normalizedLanLocationKey -and $Global:ProcessedLocationData[$normalizedLanLocationKey] -ne $null) { 
                    $matchedLocation = $Global:ProcessedLocationData[$normalizedLanLocationKey]
                    Write-Host "DEBUG [LAN CSV]: MATCH FOUND for '$($normalizedLanLocationKey)'" 
                    $buildingNameForGrid = $matchedLocation.OriginalName 
                    $address1FromLocation = $matchedLocation.Address1
                    if (-not [string]::IsNullOrWhiteSpace($address1FromLocation) -and $address1FromLocation -ne "N/A") {
                        $buildingNameForGrid = "$($buildingNameForGrid) ($($address1FromLocation))".Trim()
                    }
                    $cityForGrid = $matchedLocation.City
                    $stateForGrid = $matchedLocation.State
                    $zipCodeForGrid = $matchedLocation.Zip
                    $countryForGrid = $matchedLocation.Country
                    $matchedLanLocations++
                } else { 
                    Write-Host "DEBUG [LAN CSV]: Match NOT FOUND for '$($normalizedLanLocationKey)'. Reason: Key not in Global:ProcessedLocationData or key is null."
                    $buildingNameForGrid = $lanLocationKeyRaw 
                    if (![string]::IsNullOrWhiteSpace($normalizedLanLocationKey)) {$unmatchedLanLocations++}
                    
                    if ([string]::IsNullOrWhiteSpace($cityForGrid)) {$cityForGrid = & $GetPropertyValueByName -RowObject $lanRowItem -PropertyNameToCheck "city"}
                    if ([string]::IsNullOrWhiteSpace($stateForGrid)) {$stateForGrid = (& $GetPropertyValueByName -RowObject $lanRowItem -PropertyNameToCheck "region_name") -or (& $GetPropertyValueByName -RowObject $lanRowItem -PropertyNameToCheck "state")}
                    if ([string]::IsNullOrWhiteSpace($zipCodeForGrid)) {$zipCodeForGrid = (& $GetPropertyValueByName -RowObject $lanRowItem -PropertyNameToCheck "postal_code") -or (& $GetPropertyValueByName -RowObject $lanRowItem -PropertyNameToCheck "zipcode") -or (& $GetPropertyValueByName -RowObject $lanRowItem -PropertyNameToCheck "zip")}
                    if ([string]::IsNullOrWhiteSpace($countryForGrid)) {$countryForGrid = (& $GetPropertyValueByName -RowObject $lanRowItem -PropertyNameToCheck "country_name") -or (& $GetPropertyValueByName -RowObject $lanRowItem -PropertyNameToCheck "country")}
                }
                
                $tempInsideCorp = & $GetPropertyValueByName -RowObject $lanRowItem -PropertyNameToCheck "insidecorp"; if($tempInsideCorp -ne $null -and $tempInsideCorp -is [string]){$insideCorpVal = if($tempInsideCorp -match '^(0|false|no)$'){"0"}else{"1"}}
                $tempExpressRoute = & $GetPropertyValueByName -RowObject $lanRowItem -PropertyNameToCheck "expressroute"; if($tempExpressRoute -ne $null -and $tempExpressRoute -is [string]){$expressRouteVal = if($tempExpressRoute -match '^(1|true|yes)$'){"1"}else{"0"}}
                $tempVPN = & $GetPropertyValueByName -RowObject $lanRowItem -PropertyNameToCheck "vpn"; if($tempVPN -ne $null -and $tempVPN -is [string]){$vpnVal = if($tempVPN -match '^(1|true|yes)$'){"1"}else{"0"}}
                
                if (-not [string]::IsNullOrWhiteSpace($networkIP) -or -not [string]::IsNullOrWhiteSpace($buildingNameForGrid)) {
                    $buildingDataGrid.Rows.Add($networkIP, $networkName, $networkRange, $buildingNameForGrid, $cityForGrid, $stateForGrid, $zipCodeForGrid, $countryForGrid, $insideCorpVal, $expressRouteVal, $vpnVal)
                }
            } # End foreach lanRowItem

            # ... (status messages and catch block remain the same)
            $statusMessage = "LAN Info parsed. Matched with Location Data: $matchedLanLocations. Unmatched: $unmatchedLanLocations."
            if ($unmatchedLanLocations -gt 0 -and $Global:ProcessedLocationData.Count -gt 0) { 
                [System.Windows.Forms.MessageBox]::Show("Some LAN entries ($unmatchedLanLocations) could not be matched with the loaded Location Data. Their address details may be incomplete or taken directly from the LAN CSV if available. Please review. `nCommon reasons: Mismatched location names (check for extra quotes or variations).", "Merge Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } elseif ($matchedLanLocations -gt 0) {
                 [System.Windows.Forms.MessageBox]::Show("$matchedLanLocations LAN entries successfully merged with Location Data.", "Merge Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } elseif ($lanCsvData.Count -gt 0 -and $matchedLanLocations -eq 0 -and $unmatchedLanLocations -eq $lanCsvData.Count -and $Global:ProcessedLocationData.Count -gt 0) {
                [System.Windows.Forms.MessageBox]::Show("No LAN entries could be matched with the loaded Location Data. Please check for discrepancies in location names (e.g. 'Main Office' vs 'Main Office - HQ'). The BuildingName column will use names from the LAN CSV.", "No Matches Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
            $statusLabel.Text = $statusMessage

        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error parsing LAN Info CSV: $($_.Exception.Message)`nLine number: $($_.InvocationInfo.ScriptLineNumber)`nMake sure your CSV is properly formatted.", "Parse Error",[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else { 
        $statusLabel.Text = "Pasted text for main input is not detected as CSV by Get-InputParseSetup."
        [System.Windows.Forms.MessageBox]::Show("Main input was not detected as CSV. For Building Information merge, please paste a LAN Info CSV. Unstructured text processing for buildings is not supported with the merge workflow.", "CSV Expected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
} 

function Parse-NetworkInfo { 
    param ([string]$InputText)
    $networkDataGrid.Rows.Clear()
    $parseSetup = Get-InputParseSetup -InputText $InputText
    if ($parseSetup.IsCSV) {
        try {
            $csvData = $InputText | ConvertFrom-Csv -Delimiter $parseSetup.Separator
            if ($null -eq $csvData) {$csvData = @()} 
            if ($csvData -is [System.Management.Automation.PSCustomObject]) {$csvData = @($csvData)}

            foreach ($rowItem in $csvData) {
                if($null -eq $rowItem){ continue }
                $rawProps = $rowItem.PSObject.Properties
                if($null -eq $rawProps){ continue }
                # CORRECTED $getProp HELPER
                $getProp = {
                    param($propNameToCheck) 
                    $actualProp = $rawProps | Where-Object {$_.Name.Trim().ToLower() -eq $propNameToCheck.ToLower()} | Select-Object -First 1
                    if($actualProp){ return $rowItem.$($actualProp.Name) } 
                    else { return $null }
                }

                $subnet = ""; $maskBits = ""; $region = ""; $site = ""; $description = ""; $expressRoute = "0"
                
                $subnetVal = $getProp.Invoke("subnet")
                if ($subnetVal -and ($subnetVal -is [string]) -and $subnetVal -match '(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})/(\d{1,2})') {
                    $subnet = $matches[1]; $maskBits = $matches[2]
                } else {
                     $networkIPVal = $getProp.Invoke("networkip"); if($networkIPVal -is [string]){$subnet = $networkIPVal}
                     $maskBitsVal = $getProp.Invoke("maskbits"); if($maskBitsVal -is [string]){$maskBits = $maskBitsVal}
                }
                
                $location = ($getProp.Invoke("location") -or $getProp.Invoke("networksite") -replace '"','').Trim()
                $regionVal = ($getProp.Invoke("networkregion") -replace '"','').Trim(); if($regionVal){$region = $regionVal}
                
                if ([string]::IsNullOrWhiteSpace($region) -and $location -match '^"?([^-]+)\s*-\s*(.+)"?$') {
                    $region = $matches[1].Trim(); $site = $matches[2].Trim()
                } elseif ([string]::IsNullOrWhiteSpace($site)) { 
                    $site = $location
                }
                
                $networkSiteVal = $getProp.Invoke("networksite")
                if (-not [string]::IsNullOrWhiteSpace($networkSiteVal) -and $networkSiteVal -is [string]) { 
                    $site = ($networkSiteVal -replace '"','').Trim()
                }
                
                $descCandidates = @('name_vlan_name', 'vlan_name', 'description', 'notes')
                foreach($cand in $descCandidates){
                    $descVal = $getProp.Invoke($cand)
                    if(-not [string]::IsNullOrWhiteSpace($descVal) -and $descVal -is [string]){$description = ($descVal -replace '"','').Trim(); break}
                }
                                
                $erVal = $getProp.Invoke("expressroute"); if($erVal -ne $null -and $erVal -is [string]){$expressRoute = if ($erVal -match '^(1|true|yes)$') {"1"} else {"0"}}
                
                if (-not [string]::IsNullOrWhiteSpace($subnet)) {
                    $networkDataGrid.Rows.Add($region, $site, $subnet, $maskBits, $description, $expressRoute)
                }
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error parsing CSV for Network Info: $($_.Exception.Message)`nLine number: $($_.InvocationInfo.ScriptLineNumber)","Parse Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else { 
        $currentNetwork = @{Region = ""; Site = ""; Subnet = ""; MaskBits = ""; Description = ""; ExpressRoute = "0"}
        if ($parseSetup.Lines -is [System.Array]) {
            foreach ($line in $parseSetup.Lines) {
                # ... (unstructured logic for Network Info, ensure $networkDataGrid.Rows.Add call has 6 columns)
                 if ([string]::IsNullOrWhiteSpace($line)) { continue }
                if ($line -match '(?i)region[:\s]+(.+)$') { $currentNetwork.Region = $matches[1].Trim() }
                elseif ($line -match '(?i)site[:\s]+(.+)$') { $currentNetwork.Site = $matches[1].Trim() }
                # (rest of unstructured as before) ...
            } 
            if ($currentNetwork.Subnet -ne "" -or $currentNetwork.Region -ne "") { $networkDataGrid.Rows.Add($currentNetwork.Region, $currentNetwork.Site, $currentNetwork.Subnet, $currentNetwork.MaskBits, $currentNetwork.Description, $currentNetwork.ExpressRoute) }
        } 
    }
}

function Parse-EndpointInfo { 
    param ([string]$InputText)
    $endpointDataGrid.Rows.Clear()
    $parseSetup = Get-InputParseSetup -InputText $InputText
    if ($parseSetup.IsCSV) {
        try {
            $csvData = $InputText | ConvertFrom-Csv -Delimiter $parseSetup.Separator
            if ($null -eq $csvData) {$csvData = @()} 
            if ($csvData -is [System.Management.Automation.PSCustomObject]) {$csvData = @($csvData)}

            foreach ($rowItem in $csvData) {
                if($null -eq $rowItem){ continue }
                $rawProps = $rowItem.PSObject.Properties
                if($null -eq $rawProps){ continue }
                # CORRECTED $getProp HELPER
                $getProp = {
                    param($propNameToCheck) 
                    $actualProp = $rawProps | Where-Object {$_.Name.Trim().ToLower() -eq $propNameToCheck.ToLower()} | Select-Object -First 1
                    if($actualProp){ return $rowItem.$($actualProp.Name) } 
                    else { return $null }
                }
                
                $name = ""; $mac = ""; $manufacturer = ""; $model = ""; $type = ""

                $nameCand = @('EndpointName', 'Name', 'Device', 'Hostname'); foreach($c in $nameCand){$val=$getProp.Invoke($c); if(-not [string]::IsNullOrWhiteSpace($val) -and $val -is [string]){$name=$val;break}}
                $macCand = @('MacAddress', 'MAC'); foreach($c in $macCand){$val=$getProp.Invoke($c); if(-not [string]::IsNullOrWhiteSpace($val) -and $val -is [string]){$mac=$val;break}}
                $manuCand = @('Manufacturer', 'Vendor', 'Make'); foreach($c in $manuCand){$val=$getProp.Invoke($c); if(-not [string]::IsNullOrWhiteSpace($val) -and $val -is [string]){$manufacturer=$val;break}}
                $modelVal = $getProp.Invoke('Model'); if (-not [string]::IsNullOrWhiteSpace($modelVal) -and $modelVal -is [string]) {$model = $modelVal}
                $typeCand = @('Type', 'DeviceType'); foreach($c in $typeCand){$val=$getProp.Invoke($c); if(-not [string]::IsNullOrWhiteSpace($val) -and $val -is [string]){$type=$val;break}}
                
                $endpointDataGrid.Rows.Add($name, ($mac -replace '[:-]','').ToUpper(), $manufacturer, $model, $type)
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error parsing CSV for Endpoint Info: $($_.Exception.Message)`nLine number: $($_.InvocationInfo.ScriptLineNumber)","Parse Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else { 
        $currentEndpoint = @{Name = ""; MAC = ""; Manufacturer = ""; Model = ""; Type = ""}
        if ($parseSetup.Lines -is [System.Array]) {
            foreach ($line in $parseSetup.Lines) {
                # ... (unstructured logic for Endpoint Info, ensure $endpointDataGrid.Rows.Add call has 5 columns)
                if ([string]::IsNullOrWhiteSpace($line)) { continue }
                if ($line -match '(?i)(?:name|hostname|device)[:\s]+(.+)$') { $currentEndpoint.Name = $matches[1].Trim() }
                # (rest of unstructured as before) ...
            } 
            if ($currentEndpoint.Name -ne "" -or $currentEndpoint.MAC -ne "") { $endpointDataGrid.Rows.Add($currentEndpoint.Name, $currentEndpoint.MAC, $currentEndpoint.Manufacturer, $currentEndpoint.Model, $currentEndpoint.Type) }
        }
    }
}

function Parse-NetworkInfo { 
    param ([string]$InputText)
    $networkDataGrid.Rows.Clear()
    $parseSetup = Get-InputParseSetup -InputText $InputText
    if ($parseSetup.IsCSV) {
        try {
            $csvData = $InputText | ConvertFrom-Csv -Delimiter $parseSetup.Separator
            if ($null -eq $csvData) {$csvData = @()} # Ensure array for single object case or null
            if ($csvData -is [System.Management.Automation.PSCustomObject]) {$csvData = @($csvData)}


            foreach ($rowItem in $csvData) {
                if($null -eq $rowItem){ continue }
                $rawProps = $rowItem.PSObject.Properties
                if($null -eq $rawProps){ continue }
                $getProp = {param($propName) $actualProp = $rawProps | Where-Object {$_.Name.Trim().ToLower() -eq $propName.ToLower()} | Select-Object -First 1; if($actualProp -and $rowItem.PSObject.Properties.Match($actualProp.Name).Count -gt 0){return $rowItem.$($actualProp.Name)} else {return $null} }

                $subnet = ""; $maskBits = ""; $region = ""; $site = ""; $description = ""; $expressRoute = "0"
                
                $subnetVal = $getProp.Invoke("subnet")
                if ($subnetVal -and ($subnetVal -is [string]) -and $subnetVal -match '(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})/(\d{1,2})') {
                    $subnet = $matches[1]; $maskBits = $matches[2]
                } else {
                     $networkIPVal = $getProp.Invoke("networkip"); if($networkIPVal -is [string]){$subnet = $networkIPVal}
                     $maskBitsVal = $getProp.Invoke("maskbits"); if($maskBitsVal -is [string]){$maskBits = $maskBitsVal}
                }
                
                $location = ($getProp.Invoke("location") -or $getProp.Invoke("networksite") -replace '"','').Trim()
                $regionVal = ($getProp.Invoke("networkregion") -replace '"','').Trim(); if($regionVal){$region = $regionVal}
                
                if ([string]::IsNullOrWhiteSpace($region) -and $location -match '^"?([^-]+)\s*-\s*(.+)"?$') {
                    $region = $matches[1].Trim(); $site = $matches[2].Trim()
                } elseif ([string]::IsNullOrWhiteSpace($site)) { 
                    $site = $location
                }
                
                $networkSiteVal = $getProp.Invoke("networksite")
                if (-not [string]::IsNullOrWhiteSpace($networkSiteVal) -and $networkSiteVal -is [string]) { 
                    $site = ($networkSiteVal -replace '"','').Trim()
                }
                
                $descCandidates = @('name_vlan_name', 'vlan_name', 'description', 'notes')
                foreach($cand in $descCandidates){
                    $descVal = $getProp.Invoke($cand)
                    if(-not [string]::IsNullOrWhiteSpace($descVal) -and $descVal -is [string]){$description = ($descVal -replace '"','').Trim(); break}
                }
                                
                $erVal = $getProp.Invoke("expressroute"); if($erVal -ne $null -and $erVal -is [string]){$expressRoute = if ($erVal -match '^(1|true|yes)$') {"1"} else {"0"}}
                
                if (-not [string]::IsNullOrWhiteSpace($subnet)) {
                    $networkDataGrid.Rows.Add($region, $site, $subnet, $maskBits, $description, $expressRoute)
                }
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error parsing CSV for Network Info: $($_.Exception.Message)`nLine number: $($_.InvocationInfo.ScriptLineNumber)","Parse Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else { # Unstructured (same as previous correct version)
        # ... (Ensure this block adds 6 columns)
        $currentNetwork = @{Region = ""; Site = ""; Subnet = ""; MaskBits = ""; Description = ""; ExpressRoute = "0"}
        if ($parseSetup.Lines -is [System.Array]) {
            foreach ($line in $parseSetup.Lines) {
                if ([string]::IsNullOrWhiteSpace($line)) { continue }
                if ($line -match '(?i)region[:\s]+(.+)$') { $currentNetwork.Region = $matches[1].Trim() }
                elseif ($line -match '(?i)site[:\s]+(.+)$') { $currentNetwork.Site = $matches[1].Trim() }
                elseif ($line -match '(?i)subnet[:\s]+(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})(?:/(\d{1,2}))?') { $currentNetwork.Subnet = $matches[1]; if ($matches[2]) { $currentNetwork.MaskBits = $matches[2] } }
                elseif ($line -match '(?i)(?:mask|cidr)[:\s]+(\d{1,2})') { $currentNetwork.MaskBits = $matches[1].Trim() }
                elseif ($line -match '(?i)description[:\s]+(.+)$') { $currentNetwork.Description = $matches[1].Trim() }
                elseif ($line -match '(?i)express\s*route[:\s]+(true|false|yes|no|0|1)') { $currentNetwork.ExpressRoute = if ($matches[1].ToLower() -match 'true|yes|1') { "1" } else { "0" } }
                elseif ($line -match '(?i)^network\s*(\d+|[a-z\s]+)(?:[:\s]+(.+))?$') {
                    if ($currentNetwork.Subnet -ne "" -or $currentNetwork.Region -ne "") { $networkDataGrid.Rows.Add($currentNetwork.Region, $currentNetwork.Site, $currentNetwork.Subnet, $currentNetwork.MaskBits, $currentNetwork.Description, $currentNetwork.ExpressRoute) }
                    $descFromLine = if ($matches[2]) { $matches[2].Trim() } else { $matches[1].Trim() }; $currentNetwork = @{ Region = ""; Site = ""; Subnet = ""; MaskBits = ""; Description = $descFromLine; ExpressRoute = "0" }
                } 
                elseif ($line -match '^\s*(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})/(\d{1,2})\s*(.*)') { 
                     if ($currentNetwork.Subnet -ne "" -or $currentNetwork.Region -ne "") { $networkDataGrid.Rows.Add($currentNetwork.Region, $currentNetwork.Site, $currentNetwork.Subnet, $currentNetwork.MaskBits, $currentNetwork.Description, $currentNetwork.ExpressRoute); $currentNetwork = @{ Region = ""; Site = ""; Subnet = ""; MaskBits = ""; Description = ""; ExpressRoute = "0" } }
                    $currentNetwork.Subnet = $matches[1]; $currentNetwork.MaskBits = $matches[2]; if ($matches[3]) { $currentNetwork.Description = $matches[3].Trim() } 
                }
            } # End foreach $line
            if ($currentNetwork.Subnet -ne "" -or $currentNetwork.Region -ne "") { $networkDataGrid.Rows.Add($currentNetwork.Region, $currentNetwork.Site, $currentNetwork.Subnet, $currentNetwork.MaskBits, $currentNetwork.Description, $currentNetwork.ExpressRoute) }
        } # End if $parseSetup.Lines -is [System.Array]
    }
}

function Parse-EndpointInfo { 
    param ([string]$InputText)
    $endpointDataGrid.Rows.Clear()
    $parseSetup = Get-InputParseSetup -InputText $InputText
    if ($parseSetup.IsCSV) {
        try {
            $csvData = $InputText | ConvertFrom-Csv -Delimiter $parseSetup.Separator
            if ($null -eq $csvData) {$csvData = @()} 
            if ($csvData -is [System.Management.Automation.PSCustomObject]) {$csvData = @($csvData)}

            foreach ($rowItem in $csvData) {
                if($null -eq $rowItem){ continue }
                $rawProps = $rowItem.PSObject.Properties
                if($null -eq $rawProps){ continue }
                $getProp = {param($propName) $actualProp = $rawProps | Where-Object {$_.Name.Trim().ToLower() -eq $propName.ToLower()} | Select-Object -First 1; if($actualProp -and $rowItem.PSObject.Properties.Match($actualProp.Name).Count -gt 0){return $rowItem.$($actualProp.Name)} else {return $null} }
                
                $name = ""; $mac = ""; $manufacturer = ""; $model = ""; $type = ""

                $nameCand = @('EndpointName', 'Name', 'Device', 'Hostname'); foreach($c in $nameCand){$val=$getProp.Invoke($c); if(-not [string]::IsNullOrWhiteSpace($val) -and $val -is [string]){$name=$val;break}}
                $macCand = @('MacAddress', 'MAC'); foreach($c in $macCand){$val=$getProp.Invoke($c); if(-not [string]::IsNullOrWhiteSpace($val) -and $val -is [string]){$mac=$val;break}}
                $manuCand = @('Manufacturer', 'Vendor', 'Make'); foreach($c in $manuCand){$val=$getProp.Invoke($c); if(-not [string]::IsNullOrWhiteSpace($val) -and $val -is [string]){$manufacturer=$val;break}}
                $modelVal = $getProp.Invoke('Model'); if (-not [string]::IsNullOrWhiteSpace($modelVal) -and $modelVal -is [string]) {$model = $modelVal}
                $typeCand = @('Type', 'DeviceType'); foreach($c in $typeCand){$val=$getProp.Invoke($c); if(-not [string]::IsNullOrWhiteSpace($val) -and $val -is [string]){$type=$val;break}}
                
                $endpointDataGrid.Rows.Add($name, ($mac -replace '[:-]','').ToUpper(), $manufacturer, $model, $type)
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error parsing CSV for Endpoint Info: $($_.Exception.Message)`nLine number: $($_.InvocationInfo.ScriptLineNumber)","Parse Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else { # Unstructured 
        # ... (Ensure this block adds 5 columns)
        $currentEndpoint = @{Name = ""; MAC = ""; Manufacturer = ""; Model = ""; Type = ""}
        if ($parseSetup.Lines -is [System.Array]) {
            foreach ($line in $parseSetup.Lines) {
                if ([string]::IsNullOrWhiteSpace($line)) { continue }
                if ($line -match '(?i)(?:name|hostname|device)[:\s]+(.+)$') { $currentEndpoint.Name = $matches[1].Trim() }
                elseif ($line -match '(?i)mac(?:\s+address)?[:\s]+([0-9A-Fa-f]{2}[:-]?[0-9A-Fa-f]{2}[:-]?[0-9A-Fa-f]{2}[:-]?[0-9A-Fa-f]{2}[:-]?[0-9A-Fa-f]{2}[:-]?[0-9A-Fa-f]{2})') { $currentEndpoint.MAC = ($matches[1] -replace '[:-]','').ToUpper() }
                elseif ($line -match '(?i)(?:manufacturer|vendor|make)[:\s]+(.+)$') { $currentEndpoint.Manufacturer = $matches[1].Trim() }
                elseif ($line -match '(?i)model[:\s]+(.+)$') { $currentEndpoint.Model = $matches[1].Trim() }
                elseif ($line -match '(?i)type[:\s]+(.+)$') { $currentEndpoint.Type = $matches[1].Trim() }
                elseif ($line -match '(?i)^device\s*(\d+|[a-z\s]+)(?:[:\s]+(.+))?$') {
                    if ($currentEndpoint.Name -ne "" -or $currentEndpoint.MAC -ne "") { $endpointDataGrid.Rows.Add($currentEndpoint.Name, $currentEndpoint.MAC, $currentEndpoint.Manufacturer, $currentEndpoint.Model, $currentEndpoint.Type) }
                    $nameFromLine = if ($matches[2]) { $matches[2].Trim() } else { $matches[1].Trim() }; $currentEndpoint = @{ Name = $nameFromLine; MAC = ""; Manufacturer = ""; Model = ""; Type = "" }
                } 
                elseif ($line -match '([0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2})') { 
                    if ($currentEndpoint.Name -ne "" -or $currentEndpoint.MAC -ne "") { $endpointDataGrid.Rows.Add($currentEndpoint.Name, $currentEndpoint.MAC, $currentEndpoint.Manufacturer, $currentEndpoint.Model, $currentEndpoint.Type); $currentEndpoint = @{ Name = ""; MAC = ""; Manufacturer = ""; Model = ""; Type = "" } }
                    $currentEndpoint.MAC = ($matches[1] -replace '[:-]','').ToUpper() 
                }
            } # End foreach $line
            if ($currentEndpoint.Name -ne "" -or $currentEndpoint.MAC -ne "") { $endpointDataGrid.Rows.Add($currentEndpoint.Name, $currentEndpoint.MAC, $currentEndpoint.Manufacturer, $currentEndpoint.Model, $currentEndpoint.Type) }
        } # End if $parseSetup.Lines -is [System.Array]
    }
}

function Show-DataPreview { 
    param ([string]$InputText, [string]$DataType)
    $previewForm = New-Object System.Windows.Forms.Form; $previewForm.Text = "Data Mapping Preview (First CSV Row or Unstructured Interpretation)"; $previewForm.Size = New-Object System.Drawing.Size(750,550); $previewForm.StartPosition = "CenterParent"; $previewForm.Font = New-Object System.Drawing.Font("Segoe UI",9)
    $previewLabel = New-Object System.Windows.Forms.Label; $previewLabel.Location = New-Object System.Drawing.Point(20,15); $previewLabel.Size = New-Object System.Drawing.Size(710,40); $previewLabel.Text = "This shows how fields from your input (primarily from the first CSV row if CSV, or general unstructured parsing rules) might map to Microsoft Teams fields. For 'Building Information', this previews the LAN Info CSV fields; merging happens during main processing."; $previewForm.Controls.Add($previewLabel)
    $previewGrid = New-Object System.Windows.Forms.DataGridView; $previewGrid.Location = New-Object System.Drawing.Point(20,60); $previewGrid.Size = New-Object System.Drawing.Size(710,400); $previewGrid.AllowUserToAddRows = $false; $previewGrid.AllowUserToDeleteRows = $false; $previewGrid.ReadOnly = $true; $previewGrid.ColumnHeadersHeightSizeMode = "AutoSize"; $previewGrid.AutoSizeColumnsMode = "DisplayedCells"; $previewGrid.RowHeadersWidth = 30; $previewForm.Controls.Add($previewGrid)
    $okButton = New-Object System.Windows.Forms.Button; $okButton.Location = New-Object System.Drawing.Point(325,470); $okButton.Size = New-Object System.Drawing.Size(100,30); $okButton.Text = "OK"; $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK; $previewForm.Controls.Add($okButton); $previewForm.AcceptButton = $okButton
    
    $previewGrid.ColumnCount = 3; $previewGrid.Columns[0].Name = "Source Hint (CSV Header / Unstructured)"; $previewGrid.Columns[1].Name = "Teams Field"; $previewGrid.Columns[2].Name = "Sample Value (from first CSV row if available)"; $previewGrid.Columns[0].Width=250; $previewGrid.Columns[1].Width=150; $previewGrid.Columns[2].Width=250

    $parseSetup=Get-InputParseSetup -InputText $InputText; $sampleValues=@{}; 
    if($parseSetup.IsCSV){try{$csvData=$InputText | ConvertFrom-Csv -Delimiter $parseSetup.Separator -ErrorAction SilentlyContinue; if($csvData -and $csvData.Count -gt 0){$firstRow=$csvData[0]; foreach($prop in $firstRow.PSObject.Properties){$sampleValues[$prop.Name.ToLower().Trim()]=$prop.Value}}}catch{}} # Added Trim() to sampleValues keys
    
    $mappings=@()
    
    # --- Pre-calculate Sample Values with PowerShell 5.1 compatible if/else ---
    
    # Building Samples (Reflects LAN info primarily, merge happens later)
    $buildingNetworkIPSample="e.g., 192.168.1.0"; if($sampleValues.ContainsKey("subnet") -and ($sampleValues["subnet"] -is [string]) -and $sampleValues["subnet"] -match '(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})/'){$buildingNetworkIPSample=$matches[1]}elseif($sampleValues.ContainsKey("networkip")){$buildingNetworkIPSample=$sampleValues["networkip"]}
    
    $descriptionForNetworkName = ""
    if ($sampleValues.ContainsKey("description") -and ($sampleValues["description"] -is [string]) -and ($sampleValues["description"] -match 'vlan|network')) {
        $descriptionForNetworkName = $sampleValues["description"]
    }
    $buildingNetworkNameSample = $sampleValues["name_vlan_name"] -or $sampleValues["vlan_name"] -or $sampleValues["networkname"] -or $descriptionForNetworkName -or "e.g., CorpNet-VLAN10"
    
    $buildingNetworkRangeSample="e.g., 24"; if($sampleValues.ContainsKey("subnet") -and ($sampleValues["subnet"] -is [string]) -and $sampleValues["subnet"] -match '/(\d{1,2})'){$buildingNetworkRangeSample=$matches[1]}elseif($sampleValues.ContainsKey("networkrange")){$buildingNetworkRangeSample=$sampleValues["networkrange"]}
    $buildingNameSample=$sampleValues["location"]-or $sampleValues["site"]-or $sampleValues["building"] -or "e.g., Main Office (from LAN 'location')" 
    $buildingCitySample=$sampleValues["city"]-or "From LAN CSV 'city' or manual" 
    $buildingStateSample=$sampleValues["state"]-or $sampleValues["region_name"]-or "From LAN CSV 'state'/'region_name' or manual"
    $buildingZipSample=$sampleValues["zipcode"]-or $sampleValues["postal_code"]-or "From LAN CSV 'zip'/'postal_code' or manual"
    $buildingCountrySample=$sampleValues["country"]-or $sampleValues["country_name"]-or "From LAN CSV 'country' or manual"
    $buildingInsideCorpSample=$sampleValues["insidecorp"]-or"1 (default)"
    $buildingExpressRouteSample=$sampleValues["expressroute"]-or"0 (default)"
    $buildingVPNSample=$sampleValues["vpn"]-or"0 (default)"

    # Network Samples (for Network Info Tab)
    $networkRegionSample=$sampleValues["networkregion"]; if(-not $networkRegionSample -and $sampleValues.ContainsKey("location") -and ($sampleValues["location"] -is [string])){if($sampleValues["location"] -match '^"?([^-]+)\s*-\s*.+'){$networkRegionSample=$matches[1].Trim()}}; $networkRegionSample=$networkRegionSample -or "e.g., EMEA"
    $networkSiteSample=$sampleValues["networksite"]-or $sampleValues["location"]; if($networkSiteSample -and ($networkSiteSample -is [string]) -and $networkSiteSample -match '^"?([^-]+)\s*-\s*(.+)"?$'){$networkSiteSample=$matches[2].Trim()}; $networkSiteSample=$networkSiteSample -or "e.g., London-Office"
    $networkSubnetSample="e.g., 10.0.0.0"; if($sampleValues.ContainsKey("subnet") -and ($sampleValues["subnet"] -is [string]) -and $sampleValues["subnet"] -match '(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})/'){$networkSubnetSample=$matches[1]}elseif($sampleValues.ContainsKey("networkip")){$networkSubnetSample=$sampleValues["networkip"]}
    $networkMaskBitsSample="e.g., 16"; if($sampleValues.ContainsKey("subnet") -and ($sampleValues["subnet"] -is [string]) -and $sampleValues["subnet"] -match '/(\d{1,2})'){$networkMaskBitsSample=$matches[1]}elseif($sampleValues.ContainsKey("maskbits")){$networkMaskBitsSample=$sampleValues["maskbits"]}
    $networkDescriptionSample=$sampleValues["description"]-or $sampleValues["name_vlan_name"]-or $sampleValues["vlan_name"]-or $sampleValues["notes"]-or"e.g., Main Server VLAN"
    $networkExpressRouteSample=$sampleValues["expressroute"]-or"0 (default)"
    
    # Endpoint Samples
    $endpointNameSample=$sampleValues["endpointname"]-or $sampleValues["name"]-or $sampleValues["device"]-or $sampleValues["hostname"]-or"e.g., User-PC-01"
    $endpointMacAddressSample=$sampleValues["macaddress"]-or $sampleValues["mac"]-or"e.g., 00AABBCCDDEE"
    $endpointManufacturerSample=$sampleValues["manufacturer"]-or $sampleValues["vendor"]-or $sampleValues["make"]-or"e.g., Dell Inc."
    $endpointModelSample=$sampleValues["model"]-or"e.g., Latitude 7400"
    $endpointTypeSample=$sampleValues["type"]-or $sampleValues["devicetype"]-or"e.g., Laptop"

    switch($DataType){
        "Building"{
            $mappings=@(
                @{Source="subnet (CIDR IP part from LAN CSV), networkip";Target="NetworkIP";Sample=$buildingNetworkIPSample},
                @{Source="name_vlan_name, vlan_name, networkname, description (from LAN CSV)";Target="NetworkName";Sample=$buildingNetworkNameSample},
                @{Source="subnet (CIDR mask part from LAN CSV), networkrange";Target="NetworkRange";Sample=$buildingNetworkRangeSample},
                @{Source="location, site, building (from LAN CSV - will be merged with Location Name)";Target="BuildingName";Sample=$buildingNameSample},
                @{Source="city (from LAN CSV or merged)";Target="City";Sample=$buildingCitySample},
                @{Source="state, region_name (from LAN CSV or merged)";Target="State";Sample=$buildingStateSample},
                @{Source="zipcode, postal_code (from LAN CSV or merged)";Target="ZipCode";Sample=$buildingZipSample}, # Added ZipCode to preview
                @{Source="country, country_name (from LAN CSV or merged)";Target="Country";Sample=$buildingCountrySample},
                @{Source="insidecorp (from LAN CSV)";Target="InsideCorp";Sample=$buildingInsideCorpSample},
                @{Source="expressroute (from LAN CSV)";Target="ExpressRoute";Sample=$buildingExpressRouteSample},
                @{Source="vpn (from LAN CSV)";Target="VPN";Sample=$buildingVPNSample}
            )
        } 
        "Network"{
            $mappings=@(
                @{Source="networkregion or first part of location (before ' - ')";Target="NetworkRegion";Sample=$networkRegionSample},
                @{Source="networksite, location, or second part of location";Target="NetworkSite";Sample=$networkSiteSample},
                @{Source="subnet (CIDR IP part), networkip";Target="Subnet";Sample=$networkSubnetSample},
                @{Source="subnet (CIDR mask part), maskbits";Target="MaskBits";Sample=$networkMaskBitsSample},
                @{Source="name_vlan_name, vlan_name, description, notes";Target="Description";Sample=$networkDescriptionSample},
                @{Source="expressroute (0/1,true/false)";Target="ExpressRoute";Sample=$networkExpressRouteSample}
            )
        } 
        "Endpoint"{
            $mappings=@(
                @{Source="EndpointName, Name, Device, Hostname";Target="EndpointName";Sample=$endpointNameSample},
                @{Source="MacAddress, MAC";Target="MacAddress";Sample=$endpointMacAddressSample},
                @{Source="Manufacturer, Vendor, Make";Target="Manufacturer";Sample=$endpointManufacturerSample},
                @{Source="Model";Target="Model";Sample=$endpointModelSample},
                @{Source="Type, DeviceType";Target="Type";Sample=$endpointTypeSample}
            )
        }
    }
    foreach($mapping in $mappings){$previewGrid.Rows.Add($mapping.Source,$mapping.Target,$mapping.Sample)}
    $previewForm.ShowDialog()|Out-Null
    $previewForm.Dispose()
}

$loadLocationDataButton.Add_Click({
    Load-AndParseLocationInfoCSVFromFile -UiStatusLabel $statusLabel -UiLocationDataStatusLabel $locationDataStatusLabel
})

$processPastedLocationDataButton.Add_Click({
    $locationCsvText = $locationCsvPasteTextBox.Text
    if ([string]::IsNullOrWhiteSpace($locationCsvText)) {
        [System.Windows.Forms.MessageBox]::Show("Please paste Location Info CSV data into Box 1 first.", "No Location Data Pasted", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    Process-LocationCsvContent -CsvContent $locationCsvText -UiStatusLabel $statusLabel -UiLocationDataStatusLabel $locationDataStatusLabel
})

$pasteButton.Add_Click({ # Pastes to main input text box (Box 2)
    $inputTextBox.Text = Get-Clipboard
    $statusLabel.Text = "Text pasted to main input area (Box 2)."
})

$detectButton.Add_Click({ # Operates on main input text box (Box 2)
    $inputText = $inputTextBox.Text
    if ([string]::IsNullOrWhiteSpace($inputText)) {
        [System.Windows.Forms.MessageBox]::Show("Please paste some text into the main input area (Box 2) to detect its format.", "No Input in Box 2", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $dataType = Detect-DataFormat -InputText $inputText
    switch ($dataType) {
        "Building" { $buildingRadioButton.Checked = $true }
        "Network" { $networkRadioButton.Checked = $true }
        "Endpoint" { $endpointRadioButton.Checked = $true }
    }
    $statusLabel.Text = "Detected data type for main input (Box 2): $dataType"
})

$previewButton.Add_Click({ # Operates on main input text box (Box 2)
    $inputText = $inputTextBox.Text
    if ([string]::IsNullOrWhiteSpace($inputText)) {
        [System.Windows.Forms.MessageBox]::Show("Please paste some text into the main input area (Box 2) to preview.", "No Input in Box 2", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $dataType = "Building" 
    if ($networkRadioButton.Checked) { $dataType = "Network" }
    elseif ($endpointRadioButton.Checked) { $dataType = "Endpoint" }
    
    Show-DataPreview -InputText $inputText -DataType $dataType 
})

$processButton.Add_Click({ # Main process button, operates on $inputTextBox.Text (Box 2)
    $inputText = $inputTextBox.Text
    if ([string]::IsNullOrWhiteSpace($inputText)) {
        [System.Windows.Forms.MessageBox]::Show("Please paste text into the main input area (Box 2) to process.", "No Main Input", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $statusLabel.Text = "Processing main input (Box 2)..."
    $form.Refresh() 

    try {
        if ($buildingRadioButton.Checked) {
            Parse-BuildingInfo -InputText $inputText 
            $tabControl.SelectedTab = $buildingTab
            # Status message is set within Parse-BuildingInfo
            
            $duplicatesResolved = Handle-DuplicateSubnets -DataGrid $buildingDataGrid -SubnetColumnName "NetworkIP" 
            if ($duplicatesResolved) {
                 $newRowCountB = if ($buildingDataGrid.AllowUserToAddRows) { $buildingDataGrid.Rows.Count -1 } else { $buildingDataGrid.Rows.Count}
                 # Append to existing status rather than overwriting merge status
                 $statusLabel.Text = ($statusLabel.Text + " Duplicate subnets handled. Grid now has $newRowCountB entries.").Trim()
            }
        }
        elseif ($networkRadioButton.Checked) {
            Parse-NetworkInfo -InputText $inputText
            $tabControl.SelectedTab = $networkTab
            $validation = Validate-NetworkData -DataGrid $networkDataGrid
            $rowCountN = if ($networkDataGrid.AllowUserToAddRows) { $networkDataGrid.Rows.Count -1 } else { $networkDataGrid.Rows.Count}

            if ($validation.HasErrors) {
                $statusLabel.Text = "Network information processed. Found $rowCountN entries with $($validation.ErrorMessages.Count) issues."
                [System.Windows.Forms.MessageBox]::Show("The following issues were found:`n`n" + ($validation.ErrorMessages -join [System.Environment]::NewLine), "Validation Issues", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            } else {
                $statusLabel.Text = "Network information processed. Found $rowCountN entries. No validation issues."
            }
            $duplicatesResolvedN = Handle-DuplicateSubnets -DataGrid $networkDataGrid -SubnetColumnName "Subnet"
            if ($duplicatesResolvedN) {
                $newRowCountN = if ($networkDataGrid.AllowUserToAddRows) { $networkDataGrid.Rows.Count -1 } else { $networkDataGrid.Rows.Count}
                $statusLabel.Text = ($statusLabel.Text -replace "Found \d+ entries.","").Trim() + " Duplicate subnets handled. Grid now has $newRowCountN entries."
            }
        }
        elseif ($endpointRadioButton.Checked) {
            Parse-EndpointInfo -InputText $inputText
            $tabControl.SelectedTab = $endpointTab
            $rowCountE = if ($endpointDataGrid.AllowUserToAddRows) { $endpointDataGrid.Rows.Count -1 } else { $endpointDataGrid.Rows.Count}
            $statusLabel.Text = "Endpoint information processed. Found $rowCountE entries."
        }
    }
    catch {
        $statusLabel.Text = "An error occurred during processing main input."
        [System.Windows.Forms.MessageBox]::Show("An unexpected error occurred: $($_.Exception.Message)`nAt line: $($_.InvocationInfo.ScriptLineNumber)", "Processing Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$clearButton.Add_Click({
    $locationCsvPasteTextBox.Clear()
    $inputTextBox.Clear()
    $Global:ProcessedLocationData.Clear()
    $locationDataStatusLabel.Text = "No location data loaded."
    $buildingDataGrid.Rows.Clear()
    $networkDataGrid.Rows.Clear()
    $endpointDataGrid.Rows.Clear()
    $statusLabel.Text = "Ready. All inputs, loaded location data, and grids cleared."
})

$exportButton.Add_Click({
    $selectedGrid = $null
    if ($tabControl.SelectedTab -eq $buildingTab) { $selectedGrid = $buildingDataGrid }
    elseif ($tabControl.SelectedTab -eq $networkTab) { $selectedGrid = $networkDataGrid }
    elseif ($tabControl.SelectedTab -eq $endpointTab) { $selectedGrid = $endpointDataGrid }

    if ($null -eq $selectedGrid -or ($selectedGrid.AllowUserToAddRows -and $selectedGrid.Rows.Count -le 1) -or (-not $selectedGrid.AllowUserToAddRows -and $selectedGrid.Rows.Count -eq 0) ) {
        [System.Windows.Forms.MessageBox]::Show("There is no data in the selected tab to export.", "No Data to Export", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $saveFileDialog.DefaultExt = "csv"
    $saveFileDialog.AddExtension = $true
    $saveFileDialog.Title = "Export Data to CSV"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $filePath = $saveFileDialog.FileName
        $csvLines = @()
        $exportCancelledByUser = $false 

        foreach ($row in $selectedGrid.Rows) {
            if ($row.IsNewRow) { continue } 
            
            $skipThisRow = $false
            $missingFields = New-Object System.Collections.Generic.List[string]
            $rowDescriptionForDialog = "Row $($row.Index + 1)" 

            if ($tabControl.SelectedTab -eq $buildingTab) {
                if ([string]::IsNullOrWhiteSpace($row.Cells["NetworkIP"].Value)) { $missingFields.Add("NetworkIP") }
                if ([string]::IsNullOrWhiteSpace($row.Cells["NetworkRange"].Value)) { $missingFields.Add("NetworkRange") }
                if ([string]::IsNullOrWhiteSpace($row.Cells["BuildingName"].Value)) { $missingFields.Add("BuildingName") }
            }
            elseif ($tabControl.SelectedTab -eq $networkTab) {
                if ([string]::IsNullOrWhiteSpace($row.Cells["Subnet"].Value)) { $missingFields.Add("Subnet") }
                if ([string]::IsNullOrWhiteSpace($row.Cells["MaskBits"].Value)) { $missingFields.Add("MaskBits") }
                if ([string]::IsNullOrWhiteSpace($row.Cells["NetworkRegion"].Value)) { $missingFields.Add("NetworkRegion (Recommended)") }
                if ([string]::IsNullOrWhiteSpace($row.Cells["NetworkSite"].Value)) { $missingFields.Add("NetworkSite (Recommended)") }
            }
            elseif ($tabControl.SelectedTab -eq $endpointTab) {
                if ([string]::IsNullOrWhiteSpace($row.Cells["MacAddress"].Value)) { $missingFields.Add("MacAddress") }
                if ([string]::IsNullOrWhiteSpace($row.Cells["EndpointName"].Value)) { $missingFields.Add("EndpointName (Recommended)") }
            }

            if ($missingFields.Count -gt 0) {
                $fieldsString = $missingFields -join ", "
                $message = "$rowDescriptionForDialog is missing or has empty critical/recommended field(s):`n$fieldsString`n`nWould you like to skip exporting this row?"
                $dialogResult = [System.Windows.Forms.MessageBox]::Show($message, "Missing Data in Row", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
                if ($dialogResult -eq [System.Windows.Forms.DialogResult]::Yes) { $skipThisRow = $true; $statusLabel.Text = "$rowDescriptionForDialog skipped."; $form.Refresh() } 
                elseif ($dialogResult -eq [System.Windows.Forms.DialogResult]::Cancel) { $exportCancelledByUser = $true; $statusLabel.Text = "Export cancelled by user."; $form.Refresh(); break } 
                elseif ($dialogResult -eq [System.Windows.Forms.DialogResult]::No) { $skipThisRow = $false }
            }
            
            if ($skipThisRow) { continue }
            
            $lineValues = @()
            if ($tabControl.SelectedTab -eq $buildingTab) {
                $insideCorpCellVal = $row.Cells["InsideCorp"].Value
                $exportInsideCorp = if ($null -eq $insideCorpCellVal -or [string]::IsNullOrWhiteSpace($insideCorpCellVal)) { "1" } else { $insideCorpCellVal }
                $expressRouteCellVal = $row.Cells["ExpressRoute"].Value
                $exportExpressRoute = if ($null -eq $expressRouteCellVal -or [string]::IsNullOrWhiteSpace($expressRouteCellVal)) { "0" } else { $expressRouteCellVal }
                $vpnCellVal = $row.Cells["VPN"].Value
                $exportVPN = if ($null -eq $vpnCellVal -or [string]::IsNullOrWhiteSpace($vpnCellVal)) { "0" } else { $vpnCellVal }
                
                $lineValues = @(
                    ($row.Cells["NetworkIP"].Value), ($row.Cells["NetworkName"].Value), ($row.Cells["NetworkRange"].Value),
                    ($row.Cells["BuildingName"].Value), 
                    "", "", "", 
                    ($row.Cells["City"].Value), ($row.Cells["ZipCode"].Value), ($row.Cells["Country"].Value), ($row.Cells["State"].Value), 
                    "", 
                    $exportInsideCorp, $exportExpressRoute, $exportVPN
                )
            }
            elseif ($tabControl.SelectedTab -eq $networkTab) {
                $expressRouteNetCellVal = $row.Cells["ExpressRoute"].Value
                $exportExpressRouteNet = if ($null -eq $expressRouteNetCellVal -or [string]::IsNullOrWhiteSpace($expressRouteNetCellVal)) { "0" } else { $expressRouteNetCellVal }
                $lineValues = @(
                    ($row.Cells["NetworkRegion"].Value), ($row.Cells["NetworkSite"].Value), ($row.Cells["Subnet"].Value),
                    ($row.Cells["MaskBits"].Value), ($row.Cells["Description"].Value), $exportExpressRouteNet
                )
            }
            elseif ($tabControl.SelectedTab -eq $endpointTab) {
                $lineValues = @(
                    ($row.Cells["EndpointName"].Value), ($row.Cells["MacAddress"].Value), ($row.Cells["Manufacturer"].Value),
                    ($row.Cells["Model"].Value), ($row.Cells["Type"].Value), "", "", "" 
                )
            }
            
            $processedValues = $lineValues | ForEach-Object {
                $val = "$_" 
                if ($val -match ',' -or $val -match '"' -or $val -match "`n") { '"' + ($val -replace '"', '""') + '"' } 
                else { $val }
            }
            $csvLines += ($processedValues -join ',')
        } 

        if ($exportCancelledByUser) {
            [System.Windows.Forms.MessageBox]::Show("Export process was cancelled by the user.", "Export Cancelled", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } elseif ($csvLines.Count -eq 0) {
            $statusLabel.Text = "No data to export after validation/filtering."
            [System.Windows.Forms.MessageBox]::Show("No data was available to export. All rows may have been skipped or the grid was empty.", "No Data Exported", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            try {
                $finalCsvContent = $csvLines -join [System.Environment]::NewLine
                if (-not [string]::IsNullOrEmpty($finalCsvContent)) { $finalCsvContent += [System.Environment]::NewLine }
                [System.IO.File]::WriteAllText($filePath, $finalCsvContent, [System.Text.Encoding]::UTF8)
                $statusLabel.Text = "Data exported to " + (Split-Path $filePath -Leaf)
                [System.Windows.Forms.MessageBox]::Show("Data exported successfully to $filePath`n($($csvLines.Count) rows written).`nNo header row included.", "Export Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                $statusLabel.Text = "Error exporting data."
                [System.Windows.Forms.MessageBox]::Show("Error exporting data to $filePath : $($_.Exception.Message)`nAt line: $($_.InvocationInfo.ScriptLineNumber)", "Export Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    } 
})

# Show the form
[void]$form.ShowDialog()
$form.Dispose() 
