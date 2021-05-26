@{

# Script module or binary module file associated with this manifest.
RootModule = '.\Hawk.psm1'

# Version number of this module.
ModuleVersion = '1.0.0'


# ID used to uniquely identify this module
GUID = '1f6b6b91-79c4-4edf-83a1-66d2dc8c3d85'

# Author of this module
Author = 'hello@sankgreall.com'

# Company or vendor of this module
CompanyName = 'Sankgreall'

# Minimum version of the Windows PowerShell engine required by this module
PowerShellVersion = '5.0'

# Modules that must be imported into the global environment prior to importing this module
RequiredModules = @(
    @{ModuleName = 'CloudConnect'; ModuleVersion = '1.1.2'; },
    @{ModuleName = 'RobustCloudCommand'; ModuleVersion = '1.1.3'; },
    @{ModuleName = 'ExchangeOnlineManagement'; ModuleVersion = '2.0.4'; }
    )


# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess

NestedModules = @(
    'General\Initialize-HawkGlobalObject.ps1',
    'Tenant\Get-HawkTenantAzureAuditLog.ps1'
)

# Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
FunctionsToExport = 'Get-HawkTenantAzureAuditLog','Initialize-HawkGlobalObject'


}
