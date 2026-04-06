<#
.SYNOPSIS
    Validates Azure subscriptions and resources, exports results to Excel.

.DESCRIPTION
    This script connects to Azure, validates subscription access, inventories all resources,
    and exports detailed information to an Excel file with multiple worksheets.

.PARAMETER OutputPath
    Path where the Excel file will be saved. Default: Current directory

.PARAMETER SubscriptionId
    Specific subscription ID to validate. If not provided, all accessible subscriptions are checked.

.PARAMETER TenantId
    Azure AD Tenant ID. Required when accessing subscriptions in different tenants.

.EXAMPLE
    .\AzureResourceValidator.ps1
    .\AzureResourceValidator.ps1 -OutputPath "C:\Reports" -SubscriptionId "12345678-1234-1234-1234-123456789012"
    .\AzureResourceValidator.ps1 -TenantId "87654321-4321-4321-4321-210987654321"
    .\AzureResourceValidator.ps1 -TenantId "87654321-4321-4321-4321-210987654321" -SubscriptionId "12345678-1234-1234-1234-123456789012"

.NOTES
    Requires: Az PowerShell module and ImportExcel module
    Authentication: Uses Azure CLI or Managed Identity
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = (Get-Location).Path,
    
    [Parameter(Mandatory=$false)]
    [string]$SubscriptionId = $null,
    
    [Parameter(Mandatory=$false)]
    [string]$TenantId = $null
)

#region Module Check and Installation
function Test-RequiredModules {
    $requiredModules = @('Az.Accounts', 'Az.Resources', 'Az.PolicyInsights', 'ImportExcel')
    $missingModules = @()
    
    foreach ($module in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            $missingModules += $module
        }
    }
    
    if ($missingModules.Count -gt 0) {
        Write-Warning "Missing required modules: $($missingModules -join ', ')"
        $install = Read-Host "Would you like to install missing modules? (Y/N)"
        
        if ($install -eq 'Y') {
            foreach ($module in $missingModules) {
                Write-Host "Installing $module..." -ForegroundColor Cyan
                try {
                    Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                    Write-Host "✓ $module installed successfully" -ForegroundColor Green
                }
                catch {
                    Write-Error "Failed to install $module. Please install manually: Install-Module -Name $module"
                    return $false
                }
            }
        }
        else {
            Write-Error "Required modules are missing. Please install: $($missingModules -join ', ')"
            return $false
        }
    }
    
    # Import modules
    Import-Module Az.Accounts -ErrorAction Stop
    Import-Module Az.Resources -ErrorAction Stop
    Import-Module Az.PolicyInsights -ErrorAction SilentlyContinue
    Import-Module ImportExcel -ErrorAction Stop
    
    return $true
}
#endregion

#region Authentication
function Connect-AzureWithRetry {
    [CmdletBinding()]
    param(
        [int]$MaxRetries = 3,
        [string]$TenantId = $null
    )
    
    $retryCount = 0
    $connected = $false
    
    while (-not $connected -and $retryCount -lt $MaxRetries) {
        try {
            $retryCount++
            Write-Host "Attempting to connect to Azure (Attempt $retryCount of $MaxRetries)..." -ForegroundColor Cyan
            
            if ($TenantId) {
                Write-Host "  Target Tenant: $TenantId" -ForegroundColor Gray
            }
            
            # Check if already connected to the correct tenant
            $context = Get-AzContext -ErrorAction SilentlyContinue
            
            $needsConnection = $false
            if ($null -eq $context) {
                $needsConnection = $true
            }
            elseif ($TenantId -and $context.Tenant.Id -ne $TenantId) {
                Write-Host "  Current tenant ($($context.Tenant.Id)) differs from target tenant" -ForegroundColor Yellow
                Write-Host "  Disconnecting and reconnecting to target tenant..." -ForegroundColor Yellow
                Disconnect-AzAccount -ErrorAction SilentlyContinue | Out-Null
                $needsConnection = $true
            }
            
            if ($needsConnection) {
                # Connect with or without tenant specification
                if ($TenantId) {
                    Connect-AzAccount -TenantId $TenantId -ErrorAction Stop | Out-Null
                }
                else {
                    Connect-AzAccount -ErrorAction Stop | Out-Null
                }
            }
            
            $context = Get-AzContext
            if ($null -ne $context) {
                Write-Host "✓ Connected to Azure successfully" -ForegroundColor Green
                Write-Host "  Account: $($context.Account.Id)" -ForegroundColor Gray
                Write-Host "  Tenant: $($context.Tenant.Id)" -ForegroundColor Gray
                Write-Host "  Environment: $($context.Environment.Name)" -ForegroundColor Gray
                $connected = $true
            }
        }
        catch {
            Write-Warning "Connection attempt $retryCount failed: $($_.Exception.Message)"
            
            if ($retryCount -lt $MaxRetries) {
                $waitTime = [Math]::Pow(2, $retryCount)
                Write-Host "Waiting $waitTime seconds before retry..." -ForegroundColor Yellow
                Start-Sleep -Seconds $waitTime
            }
        }
    }
    
    if (-not $connected) {
        throw "Failed to connect to Azure after $MaxRetries attempts"
    }
    
    return $context
}
#endregion

#region Subscription Validation
function Get-SubscriptionDetails {
    [CmdletBinding()]
    param(
        [string]$SubscriptionId,
        [string]$TenantId
    )
    
    Write-Host "`nGathering subscription information..." -ForegroundColor Cyan
    
    try {
        # Build parameters for Get-AzSubscription
        $subscriptionParams = @{
            ErrorAction = 'Stop'
        }
        
        if ($TenantId) {
            $subscriptionParams['TenantId'] = $TenantId
        }
        
        if ($SubscriptionId) {
            $subscriptionParams['SubscriptionId'] = $SubscriptionId
            $subscriptions = @(Get-AzSubscription @subscriptionParams)
        }
        else {
            $subscriptions = Get-AzSubscription @subscriptionParams
        }
        
        $subscriptionDetails = @()
        
        foreach ($sub in $subscriptions) {
            try {
                # Set context to subscription
                Set-AzContext -SubscriptionId $sub.Id -ErrorAction Stop | Out-Null
                
                $details = [PSCustomObject]@{
                    SubscriptionId = $sub.Id
                    SubscriptionName = $sub.Name
                    State = $sub.State
                    TenantId = $sub.TenantId
                    IsDefault = if ($sub.Id -eq (Get-AzContext).Subscription.Id) { "Yes" } else { "No" }
                    ValidationStatus = "Accessible"
                    ValidationTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                }
                
                $subscriptionDetails += $details
                Write-Host "  ✓ $($sub.Name) - $($sub.State)" -ForegroundColor Green
            }
            catch {
                $details = [PSCustomObject]@{
                    SubscriptionId = $sub.Id
                    SubscriptionName = $sub.Name
                    State = $sub.State
                    TenantId = $sub.TenantId
                    IsDefault = "No"
                    ValidationStatus = "Error: $($_.Exception.Message)"
                    ValidationTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                }
                
                $subscriptionDetails += $details
                Write-Host "  ✗ $($sub.Name) - Error accessing subscription" -ForegroundColor Red
            }
        }
        
        return $subscriptionDetails
    }
    catch {
        Write-Error "Failed to retrieve subscriptions: $($_.Exception.Message)"
        throw
    }
}
#endregion

#region Resource Inventory
function Get-ResourceInventory {
    [CmdletBinding()]
    param(
        [array]$Subscriptions
    )
    
    Write-Host "`nInventorying resources across subscriptions..." -ForegroundColor Cyan
    
    $allResources = @()
    $resourceSummary = @()
    
    foreach ($sub in $Subscriptions | Where-Object { $_.ValidationStatus -eq "Accessible" }) {
        try {
            Write-Host "  Processing: $($sub.SubscriptionName)..." -ForegroundColor Yellow
            
            # Set context
            Set-AzContext -SubscriptionId $sub.SubscriptionId -ErrorAction Stop | Out-Null
            
            # Get all resources
            $resources = Get-AzResource -ErrorAction Stop
            
            Write-Host "    Found $($resources.Count) resources" -ForegroundColor Gray
            
            # Detailed resource information with full properties
            foreach ($resource in $resources) {
                Write-Verbose "  Processing resource: $($resource.Name)"
                
                # Get full resource details including properties
                try {
                    $fullResource = Get-AzResource -ResourceId $resource.ResourceId -ExpandProperties -ErrorAction SilentlyContinue
                    
                    # Extract SKU information
                    $skuName = if ($fullResource.Sku.Name) { $fullResource.Sku.Name } else { "" }
                    $skuTier = if ($fullResource.Sku.Tier) { $fullResource.Sku.Tier } else { "" }
                    $skuCapacity = if ($fullResource.Sku.Capacity) { $fullResource.Sku.Capacity } else { "" }
                    $skuSize = if ($fullResource.Sku.Size) { $fullResource.Sku.Size } else { "" }
                    $skuFamily = if ($fullResource.Sku.Family) { $fullResource.Sku.Family } else { "" }
                    
                    # Extract Plan information
                    $planName = if ($fullResource.Plan.Name) { $fullResource.Plan.Name } else { "" }
                    $planPublisher = if ($fullResource.Plan.Publisher) { $fullResource.Plan.Publisher } else { "" }
                    $planProduct = if ($fullResource.Plan.Product) { $fullResource.Plan.Product } else { "" }
                    
                    # Extract Identity information
                    $identityType = if ($fullResource.Identity.Type) { $fullResource.Identity.Type } else { "" }
                    $identityPrincipalId = if ($fullResource.Identity.PrincipalId) { $fullResource.Identity.PrincipalId } else { "" }
                    $identityTenantId = if ($fullResource.Identity.TenantId) { $fullResource.Identity.TenantId } else { "" }
                    
                    # Extract common properties based on resource type
                    $props = $fullResource.Properties
                    
                    # Common property extractions
                    $provisioningState = if ($props.provisioningState) { $props.provisioningState } else { "" }
                    $status = if ($props.status) { $props.status } else { "" }
                    $state = if ($props.state) { $props.state } else { "" }
                    $enabled = if ($null -ne $props.enabled) { $props.enabled.ToString() } else { "" }
                    $publicNetworkAccess = if ($props.publicNetworkAccess) { $props.publicNetworkAccess } else { "" }
                    $version = if ($props.version) { $props.version } else { "" }
                    
                    # Network related properties
                    $privateEndpointConnections = if ($props.privateEndpointConnections) { 
                        "Count: $($props.privateEndpointConnections.Count)" 
                    } else { "" }
                    
                    # Encryption properties
                    $encryptionEnabled = if ($null -ne $props.encryption) { 
                        if ($props.encryption.services) { "Enabled" } else { "Configured" }
                    } else { "" }
                    
                    # Backup/Replication properties
                    $replication = if ($props.replication) { $props.replication } else { "" }
                    $redundancy = if ($props.redundancy) { $props.redundancy } else { "" }
                    
                    # Compute properties
                    $vmSize = if ($props.hardwareProfile.vmSize) { $props.hardwareProfile.vmSize } else { "" }
                    $osType = if ($props.storageProfile.osDisk.osType) { $props.storageProfile.osDisk.osType } else { "" }
                    
                    # Database properties
                    $databaseEdition = if ($props.edition) { $props.edition } else { "" }
                    $databaseCollation = if ($props.collation) { $props.collation } else { "" }
                    
                    # Web App properties
                    $hostingPlan = if ($props.serverFarmId) { 
                        $props.serverFarmId.Split('/')[-1] 
                    } else { "" }
                    $runtimeStack = if ($props.siteConfig.linuxFxVersion) { 
                        $props.siteConfig.linuxFxVersion 
                    } elseif ($props.siteConfig.windowsFxVersion) { 
                        $props.siteConfig.windowsFxVersion 
                    } else { "" }
                    $httpsOnly = if ($null -ne $props.httpsOnly) { $props.httpsOnly.ToString() } else { "" }
                    
                    # Storage Account properties
                    $accessTier = if ($props.accessTier) { $props.accessTier } else { "" }
                    $minimumTlsVersion = if ($props.minimumTlsVersion) { $props.minimumTlsVersion } else { "" }
                    $allowBlobPublicAccess = if ($null -ne $props.allowBlobPublicAccess) { 
                        $props.allowBlobPublicAccess.ToString() 
                    } else { "" }
                    
                    # Format additional properties in readable format
                    $additionalProperties = ""
                    if ($props) {
                        try {
                            # Create a readable multi-line format
                            $readableProps = @()
                            
                            $props.PSObject.Properties | ForEach-Object {
                                $propName = $_.Name
                                $propValue = $_.Value
                                
                                if ($null -eq $propValue) {
                                    $readableProps += "${propName}: <null>"
                                }
                                elseif ($propValue -is [string]) {
                                    $readableProps += "${propName}: $propValue"
                                }
                                elseif ($propValue -is [int] -or $propValue -is [long] -or $propValue -is [double] -or $propValue -is [bool]) {
                                    $readableProps += "${propName}: $propValue"
                                }
                                elseif ($propValue -is [datetime]) {
                                    $readableProps += "${propName}: $($propValue.ToString('yyyy-MM-dd HH:mm:ss'))"
                                }
                                elseif ($propValue -is [array]) {
                                    if ($propValue.Count -eq 0) {
                                        $readableProps += "${propName}: []"
                                    }
                                    elseif ($propValue.Count -le 3 -and ($propValue[0] -is [string] -or $propValue[0] -is [int])) {
                                        $readableProps += "${propName}: [$($propValue -join ', ')]"
                                    }
                                    else {
                                        $arrayJson = $propValue | ConvertTo-Json -Depth 2
                                        $readableProps += "${propName}: $arrayJson"
                                    }
                                }
                                elseif ($propValue -is [PSCustomObject] -or $propValue -is [hashtable]) {
                                    # Format nested objects
                                    $nestedProps = @()
                                    $propValue.PSObject.Properties | ForEach-Object {
                                        $nestedName = $_.Name
                                        $nestedValue = $_.Value
                                        if ($nestedValue -is [string] -or $nestedValue -is [int] -or $nestedValue -is [bool]) {
                                            $nestedProps += "  ${nestedName}: $nestedValue"
                                        }
                                        elseif ($null -ne $nestedValue) {
                                            $nestedJson = $nestedValue | ConvertTo-Json -Compress -Depth 1
                                            if ($nestedJson.Length -lt 100) {
                                                $nestedProps += "  ${nestedName}: $nestedJson"
                                            }
                                            else {
                                                $nestedProps += "  ${nestedName}: <complex object>"
                                            }
                                        }
                                    }
                                    if ($nestedProps.Count -gt 0) {
                                        $readableProps += "${propName}:"
                                        $readableProps += $nestedProps
                                    }
                                }
                                else {
                                    $readableProps += "${propName}: $($propValue.ToString())"
                                }
                            }
                            
                            # Join with line breaks for readability in Excel
                            $additionalProperties = $readableProps -join "`n"
                        }
                        catch {
                            # Fallback to formatted JSON if custom formatting fails
                            $additionalProperties = $props | ConvertTo-Json -Depth 3
                        }
                    }
                    
                    $resourceDetail = [PSCustomObject]@{
                        SubscriptionName = $sub.SubscriptionName
                        SubscriptionId = $sub.SubscriptionId
                        ResourceName = $resource.Name
                        ResourceType = $resource.ResourceType
                        ResourceGroup = $resource.ResourceGroupName
                        Location = $resource.Location
                        Kind = if ($fullResource.Kind) { $fullResource.Kind } else { "" }
                        
                        # SKU columns
                        SKU_Name = $skuName
                        SKU_Tier = $skuTier
                        SKU_Capacity = $skuCapacity
                        SKU_Size = $skuSize
                        SKU_Family = $skuFamily
                        
                        # Plan columns
                        Plan_Name = $planName
                        Plan_Publisher = $planPublisher
                        Plan_Product = $planProduct
                        
                        # Identity columns
                        Identity_Type = $identityType
                        Identity_PrincipalId = $identityPrincipalId
                        Identity_TenantId = $identityTenantId
                        
                        # Status and State columns
                        ProvisioningState = $provisioningState
                        Status = $status
                        State = $state
                        Enabled = $enabled
                        
                        # Security columns
                        PublicNetworkAccess = $publicNetworkAccess
                        PrivateEndpoints = $privateEndpointConnections
                        Encryption = $encryptionEnabled
                        HttpsOnly = $httpsOnly
                        MinimumTlsVersion = $minimumTlsVersion
                        AllowBlobPublicAccess = $allowBlobPublicAccess
                        
                        # Compute columns
                        VM_Size = $vmSize
                        OS_Type = $osType
                        
                        # Database columns
                        Database_Edition = $databaseEdition
                        Database_Collation = $databaseCollation
                        
                        # Storage columns
                        AccessTier = $accessTier
                        Replication = $replication
                        Redundancy = $redundancy
                        
                        # Web App columns
                        HostingPlan = $hostingPlan
                        RuntimeStack = $runtimeStack
                        
                        # General columns
                        Version = $version
                        Tags = if ($resource.Tags) { ($resource.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "; " } else { "" }
                        ManagedBy = if ($fullResource.ManagedBy) { $fullResource.ManagedBy } else { "" }
                        CreatedTime = if ($resource.CreatedTime) { $resource.CreatedTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }
                        ChangedTime = if ($resource.ChangedTime) { $resource.ChangedTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }
                        ResourceId = $resource.ResourceId
                        
                        # Keep full properties for reference (at the end)
                        AdditionalProperties = $additionalProperties
                    }
                }
                catch {
                    Write-Verbose "  Warning: Could not get full details for $($resource.Name): $($_.Exception.Message)"
                    
                    # Fallback to basic information
                    $resourceDetail = [PSCustomObject]@{
                        SubscriptionName = $sub.SubscriptionName
                        SubscriptionId = $sub.SubscriptionId
                        ResourceName = $resource.Name
                        ResourceType = $resource.ResourceType
                        ResourceGroup = $resource.ResourceGroupName
                        Location = $resource.Location
                        Kind = ""
                        SKU_Name = ""
                        SKU_Tier = ""
                        SKU_Capacity = ""
                        SKU_Size = ""
                        SKU_Family = ""
                        Plan_Name = ""
                        Plan_Publisher = ""
                        Plan_Product = ""
                        Identity_Type = ""
                        Identity_PrincipalId = ""
                        Identity_TenantId = ""
                        ProvisioningState = ""
                        Status = ""
                        State = ""
                        Enabled = ""
                        PublicNetworkAccess = ""
                        PrivateEndpoints = ""
                        Encryption = ""
                        HttpsOnly = ""
                        MinimumTlsVersion = ""
                        AllowBlobPublicAccess = ""
                        VM_Size = ""
                        OS_Type = ""
                        Database_Edition = ""
                        Database_Collation = ""
                        AccessTier = ""
                        Replication = ""
                        Redundancy = ""
                        HostingPlan = ""
                        RuntimeStack = ""
                        Version = ""
                        Tags = if ($resource.Tags) { ($resource.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "; " } else { "" }
                        ManagedBy = ""
                        CreatedTime = if ($resource.CreatedTime) { $resource.CreatedTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }
                        ChangedTime = if ($resource.ChangedTime) { $resource.ChangedTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }
                        ResourceId = $resource.ResourceId
                        AdditionalProperties = "Error retrieving properties"
                    }
                }
                
                $allResources += $resourceDetail
            }
            
            # Resource summary by type
            $typeGroups = $resources | Group-Object ResourceType
            foreach ($group in $typeGroups) {
                $summary = [PSCustomObject]@{
                    SubscriptionName = $sub.SubscriptionName
                    SubscriptionId = $sub.SubscriptionId
                    ResourceType = $group.Name
                    Count = $group.Count
                    Locations = ($group.Group | Select-Object -ExpandProperty Location -Unique) -join ", "
                }
                
                $resourceSummary += $summary
            }
        }
        catch {
            Write-Warning "  Error processing subscription $($sub.SubscriptionName): $($_.Exception.Message)"
        }
    }
    
    # Create a properties detail worksheet with key-value pairs
    $propertiesDetails = @()
    
    Write-Host "`nProcessing additional properties for detailed view..." -ForegroundColor Cyan
    
    foreach ($resource in $allResources) {
        if ($resource.AdditionalProperties -and $resource.AdditionalProperties -ne "" -and $resource.AdditionalProperties -ne "Error retrieving properties") {
            try {
                $props = $resource.AdditionalProperties | ConvertFrom-Json
                
                # Create a comprehensive flattened view
                $flattenedProps = [ordered]@{
                    SubscriptionName = $resource.SubscriptionName
                    SubscriptionId = $resource.SubscriptionId
                    ResourceName = $resource.ResourceName
                    ResourceType = $resource.ResourceType
                    ResourceGroup = $resource.ResourceGroup
                    Location = $resource.Location
                }
                
                # Function to flatten nested objects with dot notation
                function Add-FlattenedProperties {
                    param(
                        [Parameter(Mandatory=$true)]
                        $Object,
                        [Parameter(Mandatory=$false)]
                        [string]$Prefix = ""
                    )
                    
                    if ($null -eq $Object) { return }
                    
                    $Object.PSObject.Properties | ForEach-Object {
                        $key = if ($Prefix) { "$Prefix.$($_.Name)" } else { $_.Name }
                        $value = $_.Value
                        
                        # Handle different value types
                        if ($null -eq $value) {
                            $flattenedProps[$key] = ""
                        }
                        elseif ($value -is [string]) {
                            $flattenedProps[$key] = $value
                        }
                        elseif ($value -is [int] -or $value -is [long] -or $value -is [double]) {
                            $flattenedProps[$key] = $value
                        }
                        elseif ($value -is [bool]) {
                            $flattenedProps[$key] = $value.ToString()
                        }
                        elseif ($value -is [datetime]) {
                            $flattenedProps[$key] = $value.ToString("yyyy-MM-dd HH:mm:ss")
                        }
                        elseif ($value -is [array]) {
                            if ($value.Count -eq 0) {
                                $flattenedProps[$key] = ""
                            }
                            elseif ($value[0] -is [string] -or $value[0] -is [int]) {
                                $flattenedProps[$key] = $value -join "; "
                            }
                            else {
                                $flattenedProps[$key] = ($value | ConvertTo-Json -Compress -Depth 1)
                            }
                        }
                        elseif ($value -is [PSCustomObject] -or $value -is [hashtable]) {
                            # For nested objects, flatten one level deep only to avoid too many columns
                            $nestedJson = $value | ConvertTo-Json -Compress -Depth 1
                            if ($nestedJson.Length -lt 200) {
                                # If it's small enough, try to flatten it
                                try {
                                    $value.PSObject.Properties | ForEach-Object {
                                        $nestedKey = "$key.$($_.Name)"
                                        $nestedValue = $_.Value
                                        if ($nestedValue -is [string] -or $nestedValue -is [int] -or $nestedValue -is [bool]) {
                                            $flattenedProps[$nestedKey] = $nestedValue
                                        }
                                        else {
                                            $flattenedProps[$nestedKey] = ($nestedValue | ConvertTo-Json -Compress -Depth 1)
                                        }
                                    }
                                }
                                catch {
                                    $flattenedProps[$key] = $nestedJson
                                }
                            }
                            else {
                                $flattenedProps[$key] = $nestedJson
                            }
                        }
                        else {
                            $flattenedProps[$key] = $value.ToString()
                        }
                    }
                }
                
                # Add all flattened properties
                Add-FlattenedProperties -Object $props
                
                $propertiesDetails += [PSCustomObject]$flattenedProps
            }
            catch {
                Write-Verbose "Could not parse properties for $($resource.ResourceName): $($_.Exception.Message)"
            }
        }
    }
    
    Write-Host "  Processed $($propertiesDetails.Count) resources with detailed properties" -ForegroundColor Gray
    
    return @{
        DetailedResources = $allResources
        ResourceSummary = $resourceSummary
        PropertiesDetails = $propertiesDetails
    }
}
#endregion

#region Management Groups
function Get-ManagementGroupDetails {
    [CmdletBinding()]
    param()
    
    Write-Host "`nGathering management group information..." -ForegroundColor Cyan
    
    $allManagementGroups = @()
    
    try {
        $managementGroups = Get-AzManagementGroup -ErrorAction Stop
        
        foreach ($mg in $managementGroups) {
            try {
                # Get detailed information
                $mgDetail = Get-AzManagementGroup -GroupId $mg.Name -Expand -ErrorAction SilentlyContinue
                
                $mgInfo = [PSCustomObject]@{
                    ManagementGroupId = $mg.Name
                    DisplayName = $mg.DisplayName
                    Type = $mg.Type
                    ParentId = if ($mgDetail.ParentId) { $mgDetail.ParentId } else { "Root" }
                    ParentName = if ($mgDetail.ParentName) { $mgDetail.ParentName } else { "Root" }
                    Children = if ($mgDetail.Children) { ($mgDetail.Children | ForEach-Object { $_.DisplayName }) -join "; " } else { "" }
                    ChildCount = if ($mgDetail.Children) { $mgDetail.Children.Count } else { 0 }
                    UpdatedTime = if ($mg.UpdatedTime) { $mg.UpdatedTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }
                }
                
                $allManagementGroups += $mgInfo
                Write-Host "  ✓ $($mg.DisplayName)" -ForegroundColor Green
            }
            catch {
                Write-Warning "  Error retrieving details for management group $($mg.Name): $($_.Exception.Message)"
            }
        }
        
        Write-Host "  Found $($allManagementGroups.Count) management groups" -ForegroundColor Gray
    }
    catch {
        Write-Warning "Error retrieving management groups: $($_.Exception.Message)"
        Write-Host "  Note: You may not have permissions to view management groups" -ForegroundColor Yellow
    }
    
    return $allManagementGroups
}
#endregion

#region Policy Assignments
function Get-PolicyAssignmentDetails {
    [CmdletBinding()]
    param(
        [array]$Subscriptions
    )
    
    Write-Host "`nGathering policy assignments..." -ForegroundColor Cyan
    
    $allPolicyAssignments = @()
    
    foreach ($sub in $Subscriptions | Where-Object { $_.ValidationStatus -eq "Accessible" }) {
        try {
            Set-AzContext -SubscriptionId $sub.SubscriptionId -ErrorAction Stop | Out-Null
            
            $policyAssignments = Get-AzPolicyAssignment -ErrorAction Stop
            
            foreach ($assignment in $policyAssignments) {
                $assignmentDetail = [PSCustomObject]@{
                    SubscriptionName = $sub.SubscriptionName
                    SubscriptionId = $sub.SubscriptionId
                    AssignmentName = $assignment.Name
                    DisplayName = $assignment.Properties.DisplayName
                    Description = $assignment.Properties.Description
                    PolicyDefinitionId = $assignment.Properties.PolicyDefinitionId
                    Scope = $assignment.Properties.Scope
                    NotScopes = if ($assignment.Properties.NotScopes) { $assignment.Properties.NotScopes -join "; " } else { "" }
                    EnforcementMode = $assignment.Properties.EnforcementMode
                    Parameters = if ($assignment.Properties.Parameters) { ($assignment.Properties.Parameters | ConvertTo-Json -Compress -Depth 5) } else { "" }
                    Metadata = if ($assignment.Properties.Metadata) { ($assignment.Properties.Metadata | ConvertTo-Json -Compress -Depth 5) } else { "" }
                    ResourceId = $assignment.ResourceId
                }
                
                $allPolicyAssignments += $assignmentDetail
            }
            
            Write-Host "  ✓ $($sub.SubscriptionName): $($policyAssignments.Count) policy assignments" -ForegroundColor Green
        }
        catch {
            Write-Warning "  Error retrieving policy assignments for $($sub.SubscriptionName): $($_.Exception.Message)"
        }
    }
    
    return $allPolicyAssignments
}
#endregion

#region Policy Definitions
function Get-PolicyDefinitionDetails {
    [CmdletBinding()]
    param(
        [array]$Subscriptions
    )
    
    Write-Host "`nGathering policy definitions..." -ForegroundColor Cyan
    
    $allPolicyDefinitions = @()
    
    foreach ($sub in $Subscriptions | Where-Object { $_.ValidationStatus -eq "Accessible" }) {
        try {
            Set-AzContext -SubscriptionId $sub.SubscriptionId -ErrorAction Stop | Out-Null
            
            # Get custom policy definitions (subscription level)
            $customPolicies = Get-AzPolicyDefinition -Custom -ErrorAction Stop
            
            foreach ($policy in $customPolicies) {
                $policyDetail = [PSCustomObject]@{
                    SubscriptionName = $sub.SubscriptionName
                    SubscriptionId = $sub.SubscriptionId
                    PolicyName = $policy.Name
                    DisplayName = $policy.Properties.DisplayName
                    Description = $policy.Properties.Description
                    PolicyType = $policy.Properties.PolicyType
                    Mode = $policy.Properties.Mode
                    Category = if ($policy.Properties.Metadata.category) { $policy.Properties.Metadata.category } else { "" }
                    Version = if ($policy.Properties.Metadata.version) { $policy.Properties.Metadata.version } else { "" }
                    Parameters = if ($policy.Properties.Parameters) { ($policy.Properties.Parameters | ConvertTo-Json -Compress -Depth 5) } else { "" }
                    PolicyRule = if ($policy.Properties.PolicyRule) { ($policy.Properties.PolicyRule | ConvertTo-Json -Compress -Depth 5) } else { "" }
                    ResourceId = $policy.ResourceId
                }
                
                $allPolicyDefinitions += $policyDetail
            }
            
            Write-Host "  ✓ $($sub.SubscriptionName): $($customPolicies.Count) custom policy definitions" -ForegroundColor Green
        }
        catch {
            Write-Warning "  Error retrieving policy definitions for $($sub.SubscriptionName): $($_.Exception.Message)"
        }
    }
    
    return $allPolicyDefinitions
}
#endregion

#region Policy Set Definitions (Initiatives)
function Get-PolicySetDefinitionDetails {
    [CmdletBinding()]
    param(
        [array]$Subscriptions
    )
    
    Write-Host "`nGathering policy set definitions (initiatives)..." -ForegroundColor Cyan
    
    $allPolicySetDefinitions = @()
    
    foreach ($sub in $Subscriptions | Where-Object { $_.ValidationStatus -eq "Accessible" }) {
        try {
            Set-AzContext -SubscriptionId $sub.SubscriptionId -ErrorAction Stop | Out-Null
            
            # Get custom policy set definitions (subscription level)
            $customInitiatives = Get-AzPolicySetDefinition -Custom -ErrorAction Stop
            
            foreach ($initiative in $customInitiatives) {
                $initiativeDetail = [PSCustomObject]@{
                    SubscriptionName = $sub.SubscriptionName
                    SubscriptionId = $sub.SubscriptionId
                    InitiativeName = $initiative.Name
                    DisplayName = $initiative.Properties.DisplayName
                    Description = $initiative.Properties.Description
                    PolicyType = $initiative.Properties.PolicyType
                    Category = if ($initiative.Properties.Metadata.category) { $initiative.Properties.Metadata.category } else { "" }
                    Version = if ($initiative.Properties.Metadata.version) { $initiative.Properties.Metadata.version } else { "" }
                    PolicyCount = if ($initiative.Properties.PolicyDefinitions) { $initiative.Properties.PolicyDefinitions.Count } else { 0 }
                    PolicyDefinitions = if ($initiative.Properties.PolicyDefinitions) { ($initiative.Properties.PolicyDefinitions | ConvertTo-Json -Compress -Depth 5) } else { "" }
                    Parameters = if ($initiative.Properties.Parameters) { ($initiative.Properties.Parameters | ConvertTo-Json -Compress -Depth 5) } else { "" }
                    ResourceId = $initiative.ResourceId
                }
                
                $allPolicySetDefinitions += $initiativeDetail
            }
            
            Write-Host "  ✓ $($sub.SubscriptionName): $($customInitiatives.Count) custom policy set definitions" -ForegroundColor Green
        }
        catch {
            Write-Warning "  Error retrieving policy set definitions for $($sub.SubscriptionName): $($_.Exception.Message)"
        }
    }
    
    return $allPolicySetDefinitions
}
#endregion

#region Policy Compliance
function Get-PolicyComplianceDetails {
    [CmdletBinding()]
    param(
        [array]$Subscriptions
    )
    
    Write-Host "`nGathering policy compliance states..." -ForegroundColor Cyan
    
    $allComplianceStates = @()
    
    foreach ($sub in $Subscriptions | Where-Object { $_.ValidationStatus -eq "Accessible" }) {
        try {
            Set-AzContext -SubscriptionId $sub.SubscriptionId -ErrorAction Stop | Out-Null
            
            $policyStates = Get-AzPolicyState -SubscriptionId $sub.SubscriptionId -Top 1000 -ErrorAction SilentlyContinue
            
            if ($policyStates) {
                foreach ($state in $policyStates) {
                    $complianceDetail = [PSCustomObject]@{
                        SubscriptionName = $sub.SubscriptionName
                        SubscriptionId = $sub.SubscriptionId
                        ResourceId = $state.ResourceId
                        ResourceType = $state.ResourceType
                        ResourceLocation = $state.ResourceLocation
                        PolicyAssignmentName = $state.PolicyAssignmentName
                        PolicyDefinitionName = $state.PolicyDefinitionName
                        PolicyDefinitionAction = $state.PolicyDefinitionAction
                        ComplianceState = $state.ComplianceState
                        IsCompliant = $state.IsCompliant
                        Timestamp = if ($state.Timestamp) { $state.Timestamp.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }
                    }
                    
                    $allComplianceStates += $complianceDetail
                }
            }
            
            Write-Host "  ✓ $($sub.SubscriptionName): $($policyStates.Count) policy compliance records" -ForegroundColor Green
        }
        catch {
            Write-Warning "  Error retrieving policy compliance for $($sub.SubscriptionName): $($_.Exception.Message)"
        }
    }
    
    return $allComplianceStates
}
#endregion

#region Resource Groups
function Get-ResourceGroupDetails {
    [CmdletBinding()]
    param(
        [array]$Subscriptions
    )
    
    Write-Host "`nGathering resource group information..." -ForegroundColor Cyan
    
    $allResourceGroups = @()
    
    foreach ($sub in $Subscriptions | Where-Object { $_.ValidationStatus -eq "Accessible" }) {
        try {
            Set-AzContext -SubscriptionId $sub.SubscriptionId -ErrorAction Stop | Out-Null
            
            $resourceGroups = Get-AzResourceGroup -ErrorAction Stop
            
            foreach ($rg in $resourceGroups) {
                $rgDetail = [PSCustomObject]@{
                    SubscriptionName = $sub.SubscriptionName
                    SubscriptionId = $sub.SubscriptionId
                    ResourceGroupName = $rg.ResourceGroupName
                    Location = $rg.Location
                    ProvisioningState = $rg.ProvisioningState
                    Tags = if ($rg.Tags) { ($rg.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "; " } else { "" }
                    ResourceId = $rg.ResourceId
                }
                
                $allResourceGroups += $rgDetail
            }
        }
        catch {
            Write-Warning "  Error retrieving resource groups for $($sub.SubscriptionName): $($_.Exception.Message)"
        }
    }
    
    return $allResourceGroups
}
#endregion

#region Excel Export
function Export-ToExcel {
    [CmdletBinding()]
    param(
        [hashtable]$Data,
        [string]$OutputPath
    )
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $fileName = "AzureResourceValidation_$timestamp.xlsx"
    $fullPath = Join-Path -Path $OutputPath -ChildPath $fileName
    
    Write-Host "`nExporting data to Excel..." -ForegroundColor Cyan
    Write-Host "  Output file: $fullPath" -ForegroundColor Gray
    
    try {
        # Remove file if exists
        if (Test-Path $fullPath) {
            Remove-Item $fullPath -Force
        }
        
        $excelParams = @{
            Path = $fullPath
            AutoSize = $true
            FreezeTopRow = $true
            BoldTopRow = $true
        }
        
        # Export Subscription Details
        if ($Data.Subscriptions -and $Data.Subscriptions.Count -gt 0) {
            $Data.Subscriptions | Export-Excel @excelParams -WorksheetName "Subscriptions"
            Write-Host "  ✓ Exported Subscriptions ($($Data.Subscriptions.Count) records)" -ForegroundColor Green
        }
        
        # Export Resource Groups
        if ($Data.ResourceGroups -and $Data.ResourceGroups.Count -gt 0) {
            $Data.ResourceGroups | Export-Excel @excelParams -WorksheetName "ResourceGroups"
            Write-Host "  ✓ Exported Resource Groups ($($Data.ResourceGroups.Count) records)" -ForegroundColor Green
        }
        
        # Export Resource Summary
        if ($Data.ResourceSummary -and $Data.ResourceSummary.Count -gt 0) {
            $Data.ResourceSummary | Export-Excel @excelParams -WorksheetName "ResourceSummary"
            Write-Host "  ✓ Exported Resource Summary ($($Data.ResourceSummary.Count) records)" -ForegroundColor Green
        }
        
        # Export Detailed Resources
        if ($Data.DetailedResources -and $Data.DetailedResources.Count -gt 0) {
            $Data.DetailedResources | Export-Excel @excelParams -WorksheetName "AllResources"
            Write-Host "  ✓ Exported All Resources ($($Data.DetailedResources.Count) records)" -ForegroundColor Green
        }
        
        # Export Properties Details
        if ($Data.PropertiesDetails -and $Data.PropertiesDetails.Count -gt 0) {
            $Data.PropertiesDetails | Export-Excel @excelParams -WorksheetName "ResourceProperties"
            Write-Host "  ✓ Exported Resource Properties ($($Data.PropertiesDetails.Count) records)" -ForegroundColor Green
        }
        
        # Export Management Groups
        if ($Data.ManagementGroups -and $Data.ManagementGroups.Count -gt 0) {
            $Data.ManagementGroups | Export-Excel @excelParams -WorksheetName "ManagementGroups"
            Write-Host "  ✓ Exported Management Groups ($($Data.ManagementGroups.Count) records)" -ForegroundColor Green
        }
        
        # Export Policy Assignments
        if ($Data.PolicyAssignments -and $Data.PolicyAssignments.Count -gt 0) {
            $Data.PolicyAssignments | Export-Excel @excelParams -WorksheetName "PolicyAssignments"
            Write-Host "  ✓ Exported Policy Assignments ($($Data.PolicyAssignments.Count) records)" -ForegroundColor Green
        }
        
        # Export Policy Definitions
        if ($Data.PolicyDefinitions -and $Data.PolicyDefinitions.Count -gt 0) {
            $Data.PolicyDefinitions | Export-Excel @excelParams -WorksheetName "PolicyDefinitions"
            Write-Host "  ✓ Exported Policy Definitions ($($Data.PolicyDefinitions.Count) records)" -ForegroundColor Green
        }
        
        # Export Policy Set Definitions
        if ($Data.PolicySetDefinitions -and $Data.PolicySetDefinitions.Count -gt 0) {
            $Data.PolicySetDefinitions | Export-Excel @excelParams -WorksheetName "PolicyInitiatives"
            Write-Host "  ✓ Exported Policy Initiatives ($($Data.PolicySetDefinitions.Count) records)" -ForegroundColor Green
        }
        
        # Export Policy Compliance
        if ($Data.PolicyCompliance -and $Data.PolicyCompliance.Count -gt 0) {
            $Data.PolicyCompliance | Export-Excel @excelParams -WorksheetName "PolicyCompliance"
            Write-Host "  ✓ Exported Policy Compliance ($($Data.PolicyCompliance.Count) records)" -ForegroundColor Green
        }
        
        # Export Execution Summary
        $summary = [PSCustomObject]@{
            ExecutionDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            TotalSubscriptions = if ($Data.Subscriptions) { $Data.Subscriptions.Count } else { 0 }
            AccessibleSubscriptions = if ($Data.Subscriptions) { ($Data.Subscriptions | Where-Object { $_.ValidationStatus -eq "Accessible" }).Count } else { 0 }
            TotalResourceGroups = if ($Data.ResourceGroups) { $Data.ResourceGroups.Count } else { 0 }
            TotalResources = if ($Data.DetailedResources) { $Data.DetailedResources.Count } else { 0 }
            UniqueResourceTypes = if ($Data.ResourceSummary) { $Data.ResourceSummary.Count } else { 0 }
            TotalManagementGroups = if ($Data.ManagementGroups) { $Data.ManagementGroups.Count } else { 0 }
            TotalPolicyAssignments = if ($Data.PolicyAssignments) { $Data.PolicyAssignments.Count } else { 0 }
            TotalPolicyDefinitions = if ($Data.PolicyDefinitions) { $Data.PolicyDefinitions.Count } else { 0 }
            TotalPolicyInitiatives = if ($Data.PolicySetDefinitions) { $Data.PolicySetDefinitions.Count } else { 0 }
            TotalComplianceRecords = if ($Data.PolicyCompliance) { $Data.PolicyCompliance.Count } else { 0 }
        }
        
        $summary | Export-Excel @excelParams -WorksheetName "ExecutionSummary"
        Write-Host "  ✓ Exported Execution Summary" -ForegroundColor Green
        
        Write-Host "`n✓ Export completed successfully!" -ForegroundColor Green
        Write-Host "  File location: $fullPath" -ForegroundColor Cyan
        
        return $fullPath
    }
    catch {
        Write-Error "Failed to export to Excel: $($_.Exception.Message)"
        throw
    }
}
#endregion

#region Main Execution
function Main {
    $ErrorActionPreference = 'Stop'
    
    Write-Host "===========================================================" -ForegroundColor Cyan
    Write-Host "  Azure Subscription and Resource Validation Tool" -ForegroundColor Cyan
    Write-Host "===========================================================" -ForegroundColor Cyan
    
    try {
        # Check and install required modules
        if (-not (Test-RequiredModules)) {
            return
        }
        
        # Connect to Azure
        $context = Connect-AzureWithRetry -TenantId $TenantId
        
        # Gather subscription details
        $subscriptions = Get-SubscriptionDetails -SubscriptionId $SubscriptionId -TenantId $TenantId
        
        # Gather management groups
        $managementGroups = Get-ManagementGroupDetails
        
        # Gather resource groups
        $resourceGroups = Get-ResourceGroupDetails -Subscriptions $subscriptions
        
        # Gather resource inventory
        $resourceData = Get-ResourceInventory -Subscriptions $subscriptions
        
        # Gather policy information
        $policyAssignments = Get-PolicyAssignmentDetails -Subscriptions $subscriptions
        $policyDefinitions = Get-PolicyDefinitionDetails -Subscriptions $subscriptions
        $policySetDefinitions = Get-PolicySetDefinitionDetails -Subscriptions $subscriptions
        $policyCompliance = Get-PolicyComplianceDetails -Subscriptions $subscriptions
        
        # Prepare data for export
        $exportData = @{
            Subscriptions = $subscriptions
            ManagementGroups = $managementGroups
            ResourceGroups = $resourceGroups
            ResourceSummary = $resourceData.ResourceSummary
            DetailedResources = $resourceData.DetailedResources
            PropertiesDetails = $resourceData.PropertiesDetails
            PolicyAssignments = $policyAssignments
            PolicyDefinitions = $policyDefinitions
            PolicySetDefinitions = $policySetDefinitions
            PolicyCompliance = $policyCompliance
        }
        
        # Export to Excel
        $outputFile = Export-ToExcel -Data $exportData -OutputPath $OutputPath
        
        # Display summary
        Write-Host "`n===========================================================" -ForegroundColor Cyan
        Write-Host "  Validation Summary" -ForegroundColor Cyan
        Write-Host "===========================================================" -ForegroundColor Cyan
        Write-Host "  Total Subscriptions: $($subscriptions.Count)" -ForegroundColor White
        Write-Host "  Accessible Subscriptions: $(($subscriptions | Where-Object { $_.ValidationStatus -eq 'Accessible' }).Count)" -ForegroundColor White
        Write-Host "  Management Groups: $($managementGroups.Count)" -ForegroundColor White
        Write-Host "  Total Resource Groups: $($resourceGroups.Count)" -ForegroundColor White
        Write-Host "  Total Resources: $($resourceData.DetailedResources.Count)" -ForegroundColor White
        Write-Host "  Unique Resource Types: $($resourceData.ResourceSummary.Count)" -ForegroundColor White
        Write-Host "  Policy Assignments: $($policyAssignments.Count)" -ForegroundColor White
        Write-Host "  Custom Policy Definitions: $($policyDefinitions.Count)" -ForegroundColor White
        Write-Host "  Custom Policy Initiatives: $($policySetDefinitions.Count)" -ForegroundColor White
        Write-Host "  Policy Compliance Records: $($policyCompliance.Count)" -ForegroundColor White
        Write-Host "`n  Report saved to: $outputFile" -ForegroundColor Green
        Write-Host "===========================================================" -ForegroundColor Cyan
        
        # Open the file
        $open = Read-Host "`nWould you like to open the Excel file now? (Y/N)"
        if ($open -eq 'Y') {
            Invoke-Item $outputFile
        }
    }
    catch {
        Write-Host "`n✗ Script execution failed" -ForegroundColor Red
        Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Gray
        exit 1
    }
}

# Execute main function
Main
#endregion
