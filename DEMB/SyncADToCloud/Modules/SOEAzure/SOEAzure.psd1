# Module manifest for module 'SOEAzure'
#
# -- Version history
# Version 1.03.00 (2016-07-06, Kees Hiemstra)
# - Add internal helper functions.
# - Add Sort-AzureLicense.
# - Add Test-AzureLicense.
# - Update ConvertFrom-AzureLicense.
# - Update Set-AzureLicense.
# Version 1.02.06 (2016-03-03, Kees Hiemstra)
# - Update ConvertFrom-AzureLicense.
# Version 1.02.05 (2016-02-15, Kees Hiemstra)
# - Update Set-AzureLicense.
# Version 1.02.04 (2016-02-02, Kees Hiemstra)
# - Update Set-AzureLicense.
# Version 1.02.03 (2016-01-13, Kees Hiemstra)
# - Update ConvertFrom-AzureLicense.
# - Update help in general.
# Version 1.02.02 (2015-11-08, Kees Hiemstra)
# - Update Get-AzureLicense.
# Version 1.02.01 (2015-10-22, Kees Hiemstra)
# - Update Set-AzureLicense.
# Version 1.02.00 (2015-10-05, Kees Hiemstra)
# - Update Get-AzureLicense.
# - Add ConvertFrom-AzureLicense.
# Version 1.01.01 (2015-10-02, Kees Hiemstra)
# - Update Get-AzureLicense.
# Version 1.01.00 (2015-09-30, Kees Hiemstra)
# - Add Get-AzureLicense.
# Version 1.00.02 (2015-09-20, Kees Hiemstra)
# - Update Set-AzureLicense.
# Version 1.00.01 (2015-09-09, Kees Hiemstra)
# - Update Set-AzureLicense.
# Version 1.00.00 (2015-09-07, Kees Hiemstra)
# - Add Set-AzureLicense.

@{

RootModule = 'SOEAzure'
ModuleVersion = '1.03.00'
GUID = 'f7d4553f-b9ef-4b14-9110-5ef4ec10500c'
Author = 'Kees Hiemstra'
CompanyName = 'Hewlett-Packard Enterprise'
Copyright = '(c) 2016 Kees.Hiemstra. All rights reserved.'
Description = 'SOE Azure Active Directory related cmdlets'
PowerShellVersion = '3.0'
# PowerShellHostName = ''
# PowerShellHostVersion = ''
DotNetFrameworkVersion = '4.0'
# CLRVersion = ''
# ProcessorArchitecture = ''
# RequiredModules = @()
# RequiredAssemblies = @()
# ScriptsToProcess = @()
# TypesToProcess = @()
# FormatsToProcess = @()
# NestedModules = @()
FunctionsToExport = @('ConvertFrom-AzureLicense',
                      'Get-AzureLicense',
                      'Set-AzureLicense',
                      'Sort-AzureLicenseTags',
                      'Test-AzureLicenseTags'
                     )
# CmdletsToExport = '*'
# VariablesToExport = '*'
AliasesToExport = '*'
# ModuleList = @()
FileList = @('SOEAzure.psm1')
# PrivateData = ''
# HelpInfoURI = ''
# DefaultCommandPrefix = ''
}

