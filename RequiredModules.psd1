# NOTE: follow nuget syntax for versions: https://docs.microsoft.com/en-us/nuget/reference/package-versioning#version-ranges-and-wildcards
@{
    # Development Dependencies
    # "Stucco"         = "0.1.0" - Only needed to initialize new module template
    'Pester'           = '5.3.3'
    'psake'            = '4.9.0'
    'BuildHelpers'     = '2.0.16'
    'PowerShellBuild'  = '0.6.1'
    'PSScriptAnalyzer' = '1.19.1'
    'ModuleBuilder'    = '1.*'
    'PowerShellGet'    = '2.0.4'
    'PSDepend'         = '0.4.0'
    'PSReadLine'       = '2.2.*'

    # Production Dependencies
    'DataMashup'       = '*'
    'ImportExcel'      = '7.8.4'
}

# @{
#     PSDependOptions = @{
#         Target = 'CurrentUser'
#     }
#     'Pester' = @{
#         Version = '5.3.3'
#         Parameters = @{
#             SkipPublisherCheck = $true
#         }
#     }
#     'psake' = @{
#         Version = '4.9.0'
#     }
#     'BuildHelpers' = @{
#         Version = '2.0.16'
#     }
#     'PowerShellBuild' = @{
#         Version = '0.6.1'
#     }
#     'PSScriptAnalyzer' = @{
#         Version = '1.19.1'
#     }
# }
