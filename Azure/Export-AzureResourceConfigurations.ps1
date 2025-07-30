# Export-AzureResourceConfigurations.ps1
# This script exports detailed configuration for all resources in Azure subscriptions

# Function to get resource configuration details
function Get-ResourceConfiguration {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ResourceId
    )
    
    try {
        # Get the resource type from the resource ID
        $resourceType = ($ResourceId -split "/providers/")[1].Split("/")[0,1] -join "/"
        
        # Get resource group name
        $resourceGroupName = ($ResourceId -split "/resourceGroups/")[1].Split("/")[0]
        
        # Get resource name
        $resourceName = $ResourceId.Split("/")[-1]
        
        # Different resource types need different commands to get their configuration
        switch -Wildcard ($resourceType) {
            "Microsoft.Compute/virtualMachines" {
                $config = Get-AzVM -ResourceGroupName $resourceGroupName -Name $resourceName -Status
                $details = Get-AzVM -ResourceGroupName $resourceGroupName -Name $resourceName
                
                # Get disk info
                $osDisk = $details.StorageProfile.OSDisk
                $dataDisks = $details.StorageProfile.DataDisks
                
                # Get network info
                $nics = $details.NetworkProfile.NetworkInterfaces
                $nicConfigs = @()
                foreach ($nic in $nics) {
                    $nicName = $nic.Id.Split("/")[-1]
                    $nicDetails = Get-AzNetworkInterface -ResourceGroupName $resourceGroupName -Name $nicName
                    $nicConfigs += $nicDetails
                }
                
                # Create custom object with VM details
                $configObject = [PSCustomObject]@{
                    ResourceType = $resourceType
                    Name = $resourceName
                    ResourceGroup = $resourceGroupName
                    Location = $details.Location
                    VMSize = $details.HardwareProfile.VmSize
                    OSType = $osDisk.OSType
                    OSVersion = $details.StorageProfile.ImageReference.Offer + " " + $details.StorageProfile.ImageReference.Sku
                    AdminUsername = $details.OSProfile.AdminUsername
                    OSDisk = $osDisk.Name + " (" + $osDisk.DiskSizeGB + " GB)"
                    DataDisks = ($dataDisks | ForEach-Object { $_.Name + " (" + $_.DiskSizeGB + " GB)" }) -join ", "
                    NetworkInterfaces = ($nicConfigs | ForEach-Object { 
                        $_.Name + " - Private IPs: " + (($_.IpConfigurations | ForEach-Object { $_.PrivateIpAddress }) -join ", ")
                    }) -join "; "
                    PowerState = ($config.Statuses | Where-Object { $_.Code -match "PowerState" }).DisplayStatus
                    Tags = ($details.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ", "
                    AvailabilitySet = if ($details.AvailabilitySetReference) { $details.AvailabilitySetReference.Id.Split("/")[-1] } else { "None" }
                    BootDiagnostics = if ($details.DiagnosticsProfile.BootDiagnostics.Enabled) { "Enabled" } else { "Disabled" }
                }
                return $configObject
            }
            
            "Microsoft.Storage/storageAccounts" {
                $details = Get-AzStorageAccount -ResourceGroupName $resourceGroupName -Name $resourceName
                $configObject = [PSCustomObject]@{
                    ResourceType = $resourceType
                    Name = $resourceName
                    ResourceGroup = $resourceGroupName
                    Location = $details.Location
                    SkuName = $details.Sku.Name
                    Kind = $details.Kind
                    AccessTier = $details.AccessTier
                    EnableHttpsTrafficOnly = $details.EnableHttpsTrafficOnly
                    AllowBlobPublicAccess = $details.AllowBlobPublicAccess
                    MinimumTlsVersion = $details.MinimumTlsVersion
                    PrimaryEndpoints = ($details.PrimaryEndpoints | Get-Member -MemberType NoteProperty | ForEach-Object { "$($_.Name): $($details.PrimaryEndpoints.$($_.Name))" }) -join ", "
                    NetworkRuleSet = "DefaultAction: $($details.NetworkRuleSet.DefaultAction), Bypass: $($details.NetworkRuleSet.Bypass)"
                    Tags = ($details.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ", "
                }
                return $configObject
            }
            
            "Microsoft.Network/virtualNetworks" {
                $details = Get-AzVirtualNetwork -ResourceGroupName $resourceGroupName -Name $resourceName
                $configObject = [PSCustomObject]@{
                    ResourceType = $resourceType
                    Name = $resourceName
                    ResourceGroup = $resourceGroupName
                    Location = $details.Location
                    AddressSpace = ($details.AddressSpace.AddressPrefixes) -join ", "
                    Subnets = ($details.Subnets | ForEach-Object { 
                        $_.Name + " (" + $_.AddressPrefix + ")" + 
                        $(if ($_.NetworkSecurityGroup) { " - NSG: " + $_.NetworkSecurityGroup.Id.Split("/")[-1] } else { "" })
                    }) -join "; "
                    DnsServers = if ($details.DhcpOptions.DnsServers.Count -gt 0) { ($details.DhcpOptions.DnsServers) -join ", " } else { "Default Azure DNS" }
                    EnableDdosProtection = $details.EnableDdosProtection
                    Tags = ($details.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ", "
                }
                return $configObject
            }
            
            "Microsoft.Network/networkSecurityGroups" {
                $details = Get-AzNetworkSecurityGroup -ResourceGroupName $resourceGroupName -Name $resourceName
                $configObject = [PSCustomObject]@{
                    ResourceType = $resourceType
                    Name = $resourceName
                    ResourceGroup = $resourceGroupName
                    Location = $details.Location
                    SecurityRules = ($details.SecurityRules | ForEach-Object { 
                        $_.Name + " (" + $_.Direction + ", Priority: " + $_.Priority + ", " + $_.Access + ") - " +
                        "Source: " + $(if ($_.SourceAddressPrefix -eq "*") { "Any" } else { $_.SourceAddressPrefix }) + ", " +
                        "Destination: " + $(if ($_.DestinationAddressPrefix -eq "*") { "Any" } else { $_.DestinationAddressPrefix }) + ", " +
                        "Port: " + $(if ($_.DestinationPortRange -eq "*") { "Any" } else { $_.DestinationPortRange }) + ", " +
                        "Protocol: " + $(if ($_.Protocol -eq "*") { "Any" } else { $_.Protocol })
                    }) -join "; "
                    Tags = ($details.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ", "
                }
                return $configObject
            }
            
            "Microsoft.Web/sites" {
                $details = Get-AzWebApp -ResourceGroupName $resourceGroupName -Name $resourceName
                $configObject = [PSCustomObject]@{
                    ResourceType = $resourceType
                    Name = $resourceName
                    ResourceGroup = $resourceGroupName
                    Location = $details.Location
                    State = $details.State
                    HostNames = ($details.HostNames) -join ", "
                    AppServicePlan = $details.ServerFarmId.Split("/")[-1]
                    RuntimeStack = $details.SiteConfig.LinuxFxVersion
                    AlwaysOn = $details.SiteConfig.AlwaysOn
                    HTTPSOnly = $details.HttpsOnly
                    ClientCertEnabled = $details.ClientCertEnabled
                    DefaultDocuments = ($details.SiteConfig.DefaultDocuments) -join ", "
                    NetFrameworkVersion = $details.SiteConfig.NetFrameworkVersion
                    PhpVersion = $details.SiteConfig.PhpVersion
                    PythonVersion = $details.SiteConfig.PythonVersion
                    JavaVersion = $details.SiteConfig.JavaVersion
                    Tags = ($details.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ", "
                }
                return $configObject
            }
            
            "Microsoft.Sql/servers/databases" {
                $serverName = $ResourceId.Split("/")[-3]
                $details = Get-AzSqlDatabase -ResourceGroupName $resourceGroupName -ServerName $serverName -DatabaseName $resourceName
                $configObject = [PSCustomObject]@{
                    ResourceType = $resourceType
                    Name = $resourceName
                    ResourceGroup = $resourceGroupName
                    Location = $details.Location
                    ServerName = $serverName
                    Edition = $details.Edition
                    Status = $details.Status
                    Collation = $details.Collation
                    MaxSizeBytes = [math]::Round($details.MaxSizeBytes / 1GB, 2).ToString() + " GB"
                    ElasticPoolName = $details.ElasticPoolName
                    CreationDate = $details.CreationDate
                    CurrentServiceObjectiveName = $details.CurrentServiceObjectiveName
                    Tags = ($details.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ", "
                }
                return $configObject
            }
            
            "Microsoft.KeyVault/vaults" {
                $details = Get-AzKeyVault -ResourceGroupName $resourceGroupName -VaultName $resourceName
                $configObject = [PSCustomObject]@{
                    ResourceType = $resourceType
                    Name = $resourceName
                    ResourceGroup = $resourceGroupName
                    Location = $details.Location
                    Sku = $details.Sku
                    EnabledForDeployment = $details.EnabledForDeployment
                    EnabledForDiskEncryption = $details.EnabledForDiskEncryption
                    EnabledForTemplateDeployment = $details.EnabledForTemplateDeployment
                    EnableSoftDelete = $details.EnableSoftDelete
                    SoftDeleteRetentionInDays = $details.SoftDeleteRetentionInDays
                    EnableRbacAuthorization = $details.EnableRbacAuthorization
                    EnablePurgeProtection = $details.EnablePurgeProtection
                    VaultUri = $details.VaultUri
                    Tags = ($details.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ", "
                }
                return $configObject
            }
            
            default {
                # For other resource types, get generic properties
                $resource = Get-AzResource -ResourceId $ResourceId
                $configObject = [PSCustomObject]@{
                    ResourceType = $resourceType
                    Name = $resourceName
                    ResourceGroup = $resourceGroupName
                    Location = $resource.Location
                    Properties = "Use Get-AzResource -ResourceId '$ResourceId' | Select-Object -ExpandProperty Properties for details"
                    Tags = ($resource.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ", "
                }
                return $configObject
            }
        }
    }
    catch {
        Write-Warning "Error getting configuration for resource $ResourceId : $_"
        $configObject = [PSCustomObject]@{
            ResourceType = $resourceType
            Name = $resourceName
            ResourceGroup = $resourceGroupName
            Error = $_.Exception.Message
        }
        return $configObject
    }
}

# Main script
try {
    # Check if Az module is installed
    if (-not (Get-Module -ListAvailable -Name Az.Accounts)) {
        Write-Host "Az module not found. Please install it using: Install-Module -Name Az -AllowClobber -Force" -ForegroundColor Red
        exit
    }
    
    # Check if logged in
    $context = Get-AzContext
    if (-not $context) {
        Write-Host "Not logged in to Azure. Please run Connect-AzAccount first." -ForegroundColor Red
        exit
    }
    
    # Get subscription ID from parameter or use current context
    $subscriptionId = $args[0]
    if (-not $subscriptionId) {
        $subscriptionId = $context.Subscription.Id
        Write-Host "Using current subscription: $($context.Subscription.Name) ($subscriptionId)" -ForegroundColor Yellow
    }
    else {
        # Set the specified subscription
        Set-AzContext -SubscriptionId $subscriptionId | Out-Null
        $context = Get-AzContext
        Write-Host "Using subscription: $($context.Subscription.Name) ($subscriptionId)" -ForegroundColor Yellow
    }
    
    # Create output folder
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $outputFolder = "AzureResourceExport-$timestamp"
    New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null
    
    Write-Host "Exporting Azure resource configurations to folder: $outputFolder" -ForegroundColor Cyan
    
    # Get all resources in the subscription
    Write-Host "Getting all resources in subscription..." -ForegroundColor Cyan
    $resources = Get-AzResource
    
    Write-Host "Found $($resources.Count) resources. Gathering detailed configurations..." -ForegroundColor Green
    
    # Group resources by type for separate CSV files
    $resourcesByType = $resources | Group-Object -Property ResourceType
    
    # Create a summary file
    $summaryFile = "$outputFolder/00_ResourceSummary.csv"
    $resources | Select-Object ResourceType, Name, ResourceGroupName, Location | 
        Export-Csv -Path $summaryFile -NoTypeInformation
    
    Write-Host "Created summary file: $summaryFile" -ForegroundColor Green
    
    # Process each resource type
    foreach ($resourceType in $resourcesByType) {
        $typeName = $resourceType.Name -replace "/", "_"
        $outputFile = "$outputFolder/$typeName.csv"
        
        Write-Host "Processing $($resourceType.Count) resources of type: $($resourceType.Name)" -ForegroundColor Cyan
        
        $configurations = @()
        
        # Process each resource of this type
        foreach ($resource in $resourceType.Group) {
            Write-Host "  Getting configuration for: $($resource.Name)" -ForegroundColor Gray
            $config = Get-ResourceConfiguration -ResourceId $resource.ResourceId
            $configurations += $config
        }
        
        # Export to CSV
        $configurations | Export-Csv -Path $outputFile -NoTypeInformation
        Write-Host "Exported configurations to: $outputFile" -ForegroundColor Green
    }
    
    # Create a master export with all resources
    $masterFile = "$outputFolder/AllResources.csv"
    
    Write-Host "Creating master export file with all resources..." -ForegroundColor Cyan
    $masterExport = @()
    
    foreach ($resource in $resources) {
        $config = Get-ResourceConfiguration -ResourceId $resource.ResourceId
        $masterExport += $config
    }
    
    $masterExport | Export-Csv -Path $masterFile -NoTypeInformation
    Write-Host "Created master export file: $masterFile" -ForegroundColor Green
    
    Write-Host "Export completed successfully!" -ForegroundColor Green
    Write-Host "Files are available in the $outputFolder folder" -ForegroundColor Green
}
catch {
    Write-Host "An error occurred: $_" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}
