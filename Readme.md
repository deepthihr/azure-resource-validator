Summary:

AzureResourceValidator is a PowerShell script designed to automate the validation of Azure Platform. It checks for the presence, configuration, and compliance of essential Azure resources and settings, ensuring that your environment aligns with best practices and organizational requirements. The script provides clear output for remediation and can be integrated into CI/CD pipelines or used as a standalone validation tool for platform governance.


This PowerShell script connects to Azure, inventories all resources across one or more subscriptions, validates access, gathers metadata (including management groups, resource groups, policies, compliance), and exports the results to a multi-sheet Excel file.

Key Features:

✅ Automatic Module Management - Checks and installs required modules (Az.Accounts, Az.Resources, ImportExcel)

✅ Robust Authentication - Connects to Azure with retry logic and exponential backoff

✅ Comprehensive Validation:
* Validates subscription access (Minimum Reader Access)
* Inventories all resource groups
* Lists all resources with detailed information
* Creates summary by resource type

✅ Excel Export with multiple worksheets:
* Subscriptions - All subscriptions with validation status
* ResourceGroups - All resource groups with details
* ResourceSummary - Count of resources by type
* AllResources - Complete resource inventory with tags, locations, etc.
* ExecutionSummary - Overview of the validation run


✅ Security Best Practices:
* Uses Azure authentication (no hardcoded credentials)
* Error handling throughout
* Logging and status updates

The script will prompt you to authenticate when connecting to the new tenant, and then proceed to validate all accessible subscriptions within that tenant (minimum reader access)

The script will create an Excel file named AzureResourceValidation\_YYYYMMDD\_HHMMSS.xlsx with all the validation results and offer to open it automatically when complete.

Usage Examples: -

\# Connect to a different tenant and validate all subscriptions in that tenant

.\\AzureResourceValidator.ps1 -TenantId "87654321-4321-4321-4321-110987654321"



\# Connect to a specific tenant and validate a specific subscription

.\\AzureResourceValidator.ps1 -TenantId "xxxxxxxxxxxxxxxxxxxxxxx" -SubscriptionId "xxxxxxxxxxxxxxxxxxxxxxx"



\# Specify tenant and output location

.\\AzureResourceValidator.ps1 -TenantId "your-tenant-id" -OutputPath "C:\\Reports"



\# Specify tenant and Subscription Id

.\\AzureResourceValidator.ps1 -TenantId "xxxxxxxxxxxxx" -SubscriptionId "xxxxxxxxxxxxxxxxxxxxxxx"

