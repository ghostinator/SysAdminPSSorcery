<#
.SYNOPSIS
    Brandon Cook's SharePoint Folder Structure Duplicator
.DESCRIPTION
    Interactive ability to duplicate folder structures between SharePoint locations.
    Allows users to select sites, navigate folders, and copy structures without content.
.NOTES
    Author: Brandon Cook
    Version: 1.0
    Date: March 11, 2025
#>

# Enable verbose output
$VerbosePreference = "Continue"
$ErrorActionPreference = "Stop"

# Import required SharePoint Online module
if (-not (Get-Module -Name PnP.PowerShell -ListAvailable)) {
    Write-Host "PnP.PowerShell module not found. Please install it using: Install-Module PnP.PowerShell -Scope CurrentUser" -ForegroundColor Red
    exit
}
Import-Module PnP.PowerShell

# Logging function
function Write-Log {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS", "DEBUG")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        "INFO"    { Write-Host $logMessage -ForegroundColor Gray }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        "ERROR"   { Write-Host $logMessage -ForegroundColor Red }
        "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
        "DEBUG"   { Write-Host $logMessage -ForegroundColor Cyan }
    }
}

# Connect to SharePoint Online site
function Connect-ToSharePoint {
    param (
        [Parameter(Mandatory=$true)]
        [string]$SiteUrl
    )
    
    try {
        Write-Log "Connecting to SharePoint site: $SiteUrl" -Level "INFO"
        Write-Log "PnP.PowerShell version: $((Get-Module PnP.PowerShell).Version)" -Level "INFO"
        
        # Connect using web login authentication
        Write-Log "Using web login authentication. A browser window will open shortly..." -Level "INFO"
        Connect-PnPOnline -Url $SiteUrl -UseWebLogin
        
        # Test connection
        $web = Get-PnPWeb
        Write-Log "Connected successfully to SharePoint site: $($web.Title)" -Level "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to connect to SharePoint: $_" -Level "ERROR"
        return $false
    }
}

# Function to get document libraries
function Get-DocumentLibraries {
    try {
        Write-Log "Getting document libraries" -Level "DEBUG"
        $libraries = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and $_.Hidden -eq $false }
        
        if ($null -eq $libraries -or ($libraries | Measure-Object).Count -eq 0) {
            Write-Log "No document libraries found" -Level "WARNING"
            return @()
        }
        
        return $libraries | ForEach-Object { 
            [PSCustomObject]@{
                Name = $_.Title
                ServerRelativeUrl = $_.RootFolder.ServerRelativeUrl
                Type = "DocLib"
            }
        }
    }
    catch {
        Write-Log "Error getting document libraries: $_" -Level "ERROR"
        return @()
    }
}

# Function to get folders within a location
function Get-Folders {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ServerRelativeUrl
    )
    
    try {
        Write-Log "Getting folders for: $ServerRelativeUrl" -Level "DEBUG"
        
        # Get folder directly using server relative URL
        $folder = Get-PnPFolder -Url $ServerRelativeUrl -Includes Folders
        
        if ($null -eq $folder) {
            Write-Log "Folder not found: $ServerRelativeUrl" -Level "WARNING"
            return @()
        }
        
        # Get subfolders
        $subfolders = $folder.Folders
        
        if ($null -eq $subfolders -or ($subfolders | Measure-Object).Count -eq 0) {
            Write-Log "No subfolders found in: $ServerRelativeUrl" -Level "DEBUG"
            return @()
        }
        
        return $subfolders | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                ServerRelativeUrl = $_.ServerRelativeUrl
                Type = "Folder"
            }
        }
    }
    catch {
        Write-Log "Error getting folders: $_" -Level "ERROR"
        return @()
    }
}

# Function to let user navigate and select a folder
function Select-SharePointFolder {
    param (
        [Parameter(Mandatory=$false)]
        [string]$PromptText = "Select a folder",
        
        [Parameter(Mandatory=$false)]
        [string]$CurrentUrl = "",
        
        [Parameter(Mandatory=$false)]
        [string]$CurrentPath = ""
    )
    
    # If empty current URL, we're at the root, so get document libraries
    if ([string]::IsNullOrEmpty($CurrentUrl)) {
        $items = Get-DocumentLibraries
        $currentDisplayPath = "Root (Document Libraries)"
    }
    else {
        # Get folders within the current location
        $items = Get-Folders -ServerRelativeUrl $CurrentUrl
        $currentDisplayPath = $CurrentPath
    }
    
    # Display the current location and available items
    Write-Host "`n$PromptText" -ForegroundColor Cyan
    Write-Host "Current location: $currentDisplayPath" -ForegroundColor Yellow
    
    if ($items.Count -eq 0) {
        Write-Host "No folders available at this location." -ForegroundColor Red
    }
    else {
        Write-Host "Available folders:" -ForegroundColor White
        for ($i = 0; $i -lt $items.Count; $i++) {
            Write-Host "[$i] $($items[$i].Name)" -ForegroundColor White
        }
    }
    
    # Add navigation options
    if (-not [string]::IsNullOrEmpty($CurrentUrl)) {
        Write-Host "[B] Go back to parent folder" -ForegroundColor Magenta
    }
    Write-Host "[S] Select current folder" -ForegroundColor Green
    Write-Host "[C] Create new folder here" -ForegroundColor Cyan
    Write-Host "[Q] Quit selection" -ForegroundColor Red
    
    $selection = Read-Host "`nEnter your choice"
    
    if ($selection -eq "Q" -or $selection -eq "q") {
        # Quit selection
        return $null
    }
    elseif ($selection -eq "B" -or $selection -eq "b") {
        # Go back to parent folder
        if ([string]::IsNullOrEmpty($CurrentUrl)) {
            Write-Log "Already at root level" -Level "WARNING"
            return Select-SharePointFolder -PromptText $PromptText
        }
        else {
            # Get parent URL
            $urlParts = $CurrentUrl.Split('/')
            $parentUrl = $urlParts[0..($urlParts.Length - 2)] -join '/'
            
            # Get parent path
            $pathParts = $CurrentPath.Split('/')
            $parentPath = if ($pathParts.Length -gt 1) { $pathParts[0..($pathParts.Length - 2)] -join '/' } else { "" }
            
            # If we're going back to root
            if ($parentUrl -eq "" -or $parentUrl -eq $CurrentUrl.Split('/')[0]) {
                return Select-SharePointFolder -PromptText $PromptText
            }
            
            return Select-SharePointFolder -PromptText $PromptText -CurrentUrl $parentUrl -CurrentPath $parentPath
        }
    }
    elseif ($selection -eq "S" -or $selection -eq "s") {
        # Select current folder
        if ([string]::IsNullOrEmpty($CurrentUrl)) {
            Write-Log "Cannot select root. Please select a document library first." -Level "WARNING"
            return Select-SharePointFolder -PromptText $PromptText
        }
        else {
            return [PSCustomObject]@{
                Path = $CurrentPath
                ServerRelativeUrl = $CurrentUrl
            }
        }
    }
    elseif ($selection -eq "C" -or $selection -eq "c") {
        # Create new folder
        $newFolderName = Read-Host "Enter new folder name"
        if ([string]::IsNullOrEmpty($newFolderName)) {
            Write-Log "Folder name cannot be empty" -Level "ERROR"
            return Select-SharePointFolder -PromptText $PromptText -CurrentUrl $CurrentUrl -CurrentPath $CurrentPath
        }
        
        try {
            if ([string]::IsNullOrEmpty($CurrentUrl)) {
                Write-Log "Cannot create folder at root level. Please select a document library first." -Level "WARNING"
                return Select-SharePointFolder -PromptText $PromptText
            }
            
            # Create the new folder
            $newFolderUrl = "$CurrentUrl/$newFolderName"
            $newFolderPath = if ([string]::IsNullOrEmpty($CurrentPath)) { $newFolderName } else { "$CurrentPath/$newFolderName" }
            
            Resolve-PnPFolder -SiteRelativePath $newFolderUrl
            Write-Log "Created folder: $newFolderPath" -Level "SUCCESS"
            
            return Select-SharePointFolder -PromptText $PromptText -CurrentUrl $CurrentUrl -CurrentPath $CurrentPath
        }
        catch {
            Write-Log "Error creating folder: $_" -Level "ERROR"
            return Select-SharePointFolder -PromptText $PromptText -CurrentUrl $CurrentUrl -CurrentPath $CurrentPath
        }
    }
    else {
        # Try to parse as index
        try {
            $index = [int]$selection
            if ($index -ge 0 -and $index -lt $items.Count) {
                $selectedItem = $items[$index]
                
                # Update current path
                $newPath = if ([string]::IsNullOrEmpty($CurrentPath)) { $selectedItem.Name } else { "$CurrentPath/$($selectedItem.Name)" }
                
                return Select-SharePointFolder -PromptText $PromptText -CurrentUrl $selectedItem.ServerRelativeUrl -CurrentPath $newPath
            }
            else {
                Write-Log "Invalid selection" -Level "ERROR"
                return Select-SharePointFolder -PromptText $PromptText -CurrentUrl $CurrentUrl -CurrentPath $CurrentPath
            }
        }
        catch {
            Write-Log "Invalid selection" -Level "ERROR"
            return Select-SharePointFolder -PromptText $PromptText -CurrentUrl $CurrentUrl -CurrentPath $CurrentPath
        }
    }
}

# Create folder if it doesn't exist - using alternative methods if needed
function New-FolderIfNotExists {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ServerRelativeUrl
    )
    
    try {
        Write-Log "Checking if folder exists: $ServerRelativeUrl" -Level "DEBUG"
        $folder = Get-PnPFolder -Url $ServerRelativeUrl -ErrorAction SilentlyContinue
        
        if ($null -eq $folder) {
            Write-Log "Creating folder: $ServerRelativeUrl" -Level "DEBUG"
            
            # Try using Resolve-PnPFolder first
            try {
                Resolve-PnPFolder -SiteRelativePath $ServerRelativeUrl
                Write-Log "Created folder: $ServerRelativeUrl" -Level "SUCCESS"
                return $true
            }
            catch {
                Write-Log "First attempt to create folder failed: $_" -Level "WARNING"
                
                # Try alternative method - using Add-PnPFolder
                try {
                    # Extract folder name and parent path
                    $folderName = $ServerRelativeUrl.Split('/')[-1]
                    $parentUrl = $ServerRelativeUrl.Substring(0, $ServerRelativeUrl.LastIndexOf('/'))
                    
                    Write-Log "Trying alternative method to create folder '$folderName' in '$parentUrl'" -Level "DEBUG"
                    Add-PnPFolder -Name $folderName -Folder $parentUrl
                    Write-Log "Created folder using alternative method: $ServerRelativeUrl" -Level "SUCCESS"
                    return $true
                }
                catch {
                    Write-Log "Alternative method also failed: $_" -Level "ERROR"
                    
                    # Try using CSOM directly as a last resort
                    try {
                        $ctx = Get-PnPContext
                        $web = Get-PnPWeb
                        
                        # Get the relative URL to the web
                        $webUrl = $web.ServerRelativeUrl
                        $folderUrl = $ServerRelativeUrl
                        
                        if ($folderUrl.StartsWith($webUrl)) {
                            $folderUrl = $folderUrl.Substring($webUrl.Length).TrimStart('/')
                        }
                        
                        Write-Log "Trying CSOM method to create folder: $folderUrl" -Level "DEBUG"
                        
                        # Use EnsureFolder which creates all folders in the path if they don't exist
                        $folder = $web.Folders.EnsureFolder($folderUrl)
                        $ctx.ExecuteQuery()
                        
                        Write-Log "Created folder using CSOM method: $ServerRelativeUrl" -Level "SUCCESS"
                        return $true
                    }
                    catch {
                        Write-Log "All methods to create folder failed: $_" -Level "ERROR"
                        return $false
                    }
                }
            }
        }
        else {
            Write-Log "Folder already exists: $ServerRelativeUrl" -Level "DEBUG"
            return $true
        }
    }
    catch {
        Write-Log "Error checking/creating folder $ServerRelativeUrl : $_" -Level "ERROR"
        return $false
    }
}

# Copy folder structure recursively
function Copy-FolderStructure {
    param (
        [Parameter(Mandatory=$true)]
        [string]$SourceUrl,
        
        [Parameter(Mandatory=$true)]
        [string]$DestinationUrl,
        
        [Parameter(Mandatory=$false)]
        [int]$Depth = 0
    )
    
    Write-Log "Processing: Source=$SourceUrl, Destination=$DestinationUrl" -Level "INFO"
    
    # Create the destination folder
    $folderCreated = New-FolderIfNotExists -ServerRelativeUrl $DestinationUrl
    if (-not $folderCreated) {
        Write-Log "Failed to create destination folder. Stopping recursion." -Level "ERROR"
        return
    }
    
    # Get all subfolders in the current source folder
    try {
        $subfolders = Get-Folders -ServerRelativeUrl $SourceUrl
        $totalFolders = ($subfolders | Measure-Object).Count
        
        Write-Log "Found $totalFolders subfolders in $SourceUrl" -Level "INFO"
        
        $processedCount = 0
        foreach ($subfolder in $subfolders) {
            $folderName = $subfolder.Name
            $newSourceUrl = "$SourceUrl/$folderName"
            $newDestUrl = "$DestinationUrl/$folderName"
            
            $indentation = "  " * $Depth
            Write-Log "$($indentation)Processing subfolder: $folderName" -Level "INFO"
            
            # Recursively copy the folder structure
            Copy-FolderStructure -SourceUrl $newSourceUrl -DestinationUrl $newDestUrl -Depth ($Depth + 1)
            
            $processedCount++
            $percentComplete = [Math]::Round(($processedCount / $totalFolders) * 100, 2)
            Write-Progress -Activity "Copying folder structure" -Status "$percentComplete% Complete" -PercentComplete $percentComplete
        }
        
        if ($totalFolders -gt 0) {
            Write-Progress -Activity "Copying folder structure" -Completed
        }
        else {
            Write-Log "No subfolders found in $SourceUrl" -Level "INFO"
        }
    }
    catch {
        Write-Log "Error processing subfolders in $SourceUrl : $_" -Level "ERROR"
    }
}

# Main execution block
try {
    # Get SharePoint site URL
    $siteUrl = Read-Host "Enter the SharePoint site URL (e.g., https://contoso.sharepoint.com/sites/YourSite)"
    
    # Connect to the site
    $connected = Connect-ToSharePoint -SiteUrl $siteUrl
    if (-not $connected) {
        throw "Failed to connect to SharePoint site. Exiting script."
    }
    
    # Get the web's server relative URL
    $web = Get-PnPWeb
    $webUrl = $web.ServerRelativeUrl
    Write-Log "Web server relative URL: $webUrl" -Level "DEBUG"
    
    # Let user select source folder
    Write-Log "Please select the source folder..." -Level "INFO"
    $sourceFolder = Select-SharePointFolder -PromptText "Select SOURCE folder"
    
    if ($null -eq $sourceFolder) {
        throw "No source folder selected. Exiting script."
    }
    
    $sourcePath = $sourceFolder.Path
    $sourceUrl = $sourceFolder.ServerRelativeUrl
    Write-Log "Selected source folder: $sourcePath (URL: $sourceUrl)" -Level "SUCCESS"
    
    # Let user select destination folder
    Write-Log "Please select the destination folder..." -Level "INFO"
    $destinationFolder = Select-SharePointFolder -PromptText "Select DESTINATION folder"
    
    if ($null -eq $destinationFolder) {
        throw "No destination folder selected. Exiting script."
    }
    
    $destinationPath = $destinationFolder.Path
    $destinationUrl = $destinationFolder.ServerRelativeUrl
    Write-Log "Selected destination folder: $destinationPath (URL: $destinationUrl)" -Level "SUCCESS"
    
    # Confirm the operation
    Write-Host "`nYou are about to copy the folder structure from:" -ForegroundColor Yellow
    Write-Host "Source: $sourcePath" -ForegroundColor Cyan
    Write-Host "Destination: $destinationPath" -ForegroundColor Cyan
    $confirmation = Read-Host "Do you want to continue? (Y/N)"
    
    if ($confirmation -ne "Y" -and $confirmation -ne "y") {
        throw "Operation cancelled by user."
    }
    
    # Start the folder structure copy
    Write-Log "Starting folder structure duplication..." -Level "INFO"
    $startTime = Get-Date
    
    Copy-FolderStructure -SourceUrl $sourceUrl -DestinationUrl $destinationUrl
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    Write-Log "Folder structure duplication completed successfully!" -Level "SUCCESS"
    Write-Log "Total duration: $($duration.ToString('hh\:mm\:ss'))" -Level "INFO"
}
catch {
    Write-Log "An error occurred: $_" -Level "ERROR"
}
finally {
    try {
        Disconnect-PnPOnline
        Write-Log "Disconnected from SharePoint" -Level "INFO"
    }
    catch {
        # Ignore disconnect errors
    }
}