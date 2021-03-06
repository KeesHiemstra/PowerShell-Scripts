# Module manifest for module 'SOEAD'
#
# -- Version history
# Version 3.00.02 (2016-12-27, Kees Hiemstra)
# - Update Submit-IDMBinding.
# Version 3.00.01 (2016-12-26, Kees Hiemstra)
# - Update Set-ADExtensionAttribute2.
# Version 3.00.00 (2016-11-04, Kees Hiemstra)
# - Rename Set-Manager to Set-ADManager.
# Version 2.03.00 (2016-11-04, Kees Hiemstra)
# - Add Set-Manager.
# Version 2.02.02 (2016-11-01, Kees Hiemstra)
# - Update Get-SOEADComputer.
# Version 2.02.01 (2016-09-15, Kees Hiemstra)
# - Update Set-ADExtensionAttribute2.
# Version 2.02.00 (2016-09-06, Kees Hiemstra)
# - Add Sumbit-IDMBinding.
# Version 2.01.05 (2016-08-28, Kees Hiemstra)
# - Update Get-SOEADUser.
# Version 2.01.04 (2016-08-03, Kees Hiemstra)
# - Update Get-SOEADUser.
# - Update Set-ADExtensionAttribute2.
# Version 2.01.03 (2016-08-01, Kees Hiemstra)
# - Update Get-SOEADUser.
# Version 2.01.02 (2016-07-25, Kees Hiemstra)
# - Update Get-SOEADUser.
# Version 2.01.01 (2016-07-22, Kees Hiemstra)
# - Update Get-SOEADUser.
# Version 2.01.00 (2016-07-06, Kees Hiemstra)
# - Add Set-ADExtensionAttribute2.
# Version 2.00.04 (2016-06-29, Kees Hiemstra)
# - Update Get-SOEADUser.
# Version 2.00.03 (2016-06-16, Kees Hiemstra)
# - Update Get-SOEADComputer.
# Version 2.00.02 (2016-06-14, Kees Hiemstra)
# - Updated Get-SOEADComputer.
# Version 2.00.01 (2016-06-10, Kees Hiemstra)
# - Updated Get-SOEADUser.
# Version 2.00.00 (2016-06-08, Kees Hiemstra)
# - Renamed the module to SOEAD.
# - Moved Get-SOEADComputer from SOEComputer module.
# Version 1.00.60 (2016-06-02, Kees Hiemstra)
# - Update Get-SOEADUser
# Version 1.00.50 (2016-02-28, Kees Hiemstra)
# - Update Get-SOEADUser
# Version 1.00.40 (2015-12-12, Kees Hiemstra)
# - Update Get-SOEADUser
# Version 1.00.31 (2015-12-09, Kees Hiemstra)
# - Update Get-SOEADUser
# Version 1.00.30 (2015-11-08, Kees Hiemstra)
# - Update Get-SOEADUser
# Version 1.00.20 (2015-10-05, Kees Hiemstra)
# - Update Get-SOEADUser
# Version 1.00.10 (2015-xx-xx, Kees Hiemstra)
# - Update Get-SOEADUser
# Version 1.00.03 (2015-xx-xx, Kees Hiemstra)
# - Update Get-SOEADUser
# Version 1.00.02 (2015-xx-xx, Kees Hiemstra)
# - Update Get-SOEADUser
# Version 1.00.01 (2015-xx-xx, Kees Hiemstra)
# - Update Get-SOEADUser
# Version 1.00.00 (2015-xx-xx, Kees Hiemstra)
# - Add Get-SOEADUser


@{

RootModule = 'SOEAD.psm1'
ModuleVersion = '3.00.01'
GUID = 'f92b90cf-29f4-4972-91bb-1ab96b615797'
Author = 'Kees Hiemstra'
CompanyName = 'Hewlett-Packard Enterprise'
Copyright = '(c) 2016 Kees.Hiemstra. All rights reserved.'
Description = 'SOE Active Directory related cmdlets'
PowerShellVersion = '3.0'
# PowerShellHostName = ''
# PowerShellHostVersion = ''
DotNetFrameworkVersion = '4.0'
# CLRVersion = ''
# ProcessorArchitecture = ''
RequiredModules = @('SOEAzure')
# RequiredAssemblies = @()
# ScriptsToProcess = @()
# TypesToProcess = @()
# FormatsToProcess = @()
# NestedModules = @()
FunctionsToExport = @('Get-SOEADUser',
                      'Get-SOEADComputer',
                      'Set-ADExtensionAttribute2',
                      'Submit-IDMBinding',
                      'Set-ADManager'
                      )
# CmdletsToExport = '*'
# VariablesToExport = '*'
AliasesToExport = '*'
# ModuleList = @()
FileList = @('SOEAD.psm1')
# PrivateData = ''
# HelpInfoURI = ''
# DefaultCommandPrefix = ''
}
