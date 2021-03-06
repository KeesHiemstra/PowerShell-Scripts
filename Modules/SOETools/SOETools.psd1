# Module manifest for module 'SOETools'
#
# -- Version history
# Version 1.05.00 (2016-12-01, Kees Hiemstra)
# - Added New-Share.
# Version 1.04.00 (2016-09-01, Kees Hiemstra)
# - Added Get-ShortString
# Version 1.03.00 (2016-04-17, Kees Hiemstra)
# - Added New-ZipArchive.
# Version 1.02.00 (2016-01-31, Kees Hiemstra)
# - Added Test-IsElevated.
# Version 1.01.00 (2016-01-27, Kees Hiemstra)
# - Added New-Share
# Version 1.00.00 (2015-11-10, Kees Hiemstra)
# - Added Get-Password.
# - Added Get-SOECredential.

@{

RootModule = 'SOETools.psm1'
ModuleVersion = '1.05.00'
GUID = '6cdb6e99-affb-41c8-9774-a2a7c4cf06d7'
Author = 'Kees Hiemstra'
CompanyName = 'Hewlett-Packard Enterprise'
Copyright = '(c) 2016 Kees.Hiemstra. All rights reserved.'
Description = 'SOE tools cmdlets'
PowerShellVersion = '3.0'
# PowerShellHostName = ''
# PowerShellHostVersion = ''
DotNetFrameworkVersion = '4.5'
# CLRVersion = ''
# ProcessorArchitecture = ''
# RequiredModules = @('')
# RequiredAssemblies = @()
# ScriptsToProcess = @()
# TypesToProcess = @()
# FormatsToProcess = @()
# NestedModules = @()
FunctionsToExport = @('Get-SOECredential',
                      'Get-Password',
                      'Test-IsElevated',
                      'New-ZipArchive',
                      'Get-ShortString',
                      'New-Share'
                      )
# CmdletsToExport = '*'
# VariablesToExport = '*'
AliasesToExport = '*'
# ModuleList = @()
FileList = @('SOETools.psm1')
# PrivateData = ''
# HelpInfoURI = ''
# DefaultCommandPrefix = ''
}
