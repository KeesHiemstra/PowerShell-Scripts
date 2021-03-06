# Module manifest for module 'SOEComputer'
#
# Version History
# ---------------
# Version 4.00.01 (2016-10-31, Kees Hiemstra)
# - Update Start-SOEServerService.
# - Update Get-SOEComputerModel.
# - Update Get-SOEComputerOS.
# - Update Get-SOEComputerBSoDEvents.
# - Update Get-SOEComputerNICDetails.
# - Update Get-SOEComputerSignedDrivers.
# - Update Get-SOEComputerRunKeys.
# - Update Get-SOEComputerPatches.
# - Update Get-SOEComputerBIOS.
# - Update Get-SOEComputerUserFirewallRules.
# - Update Get-SOEComputerPolicyFirewallRules.
# - Update Get-SOEComputerWifiProfile.
# - Update Get-SOEComputerWifiSignature.
# Version 4.00.00 (2016-06-08, Kees Hiemstra)
# - Moved Get-SOEADComputer to SOEAD module.
# - Removed dependency to ActiveDirectory module
# Version 3.01.00 (2016-06-02, Kees Hiemstra)
# - Add Get-SOEComputerWifiSignature
# - Add Get-SOEADComputer
# Version 3.00.00 ((2016-06-02, Kees Hiemstra)
# - Major update Get-SOEComputerWifiProfile
# Version 2.05.00 (2016-05-30, Kees Hiemstra)
# - Add Get-SOEComputerWifiProfile
# Version 2.04.00 (2016-05-18, Kees Hiemstra)
# - Add Get-SOEComputerPolicyFirewallRules
# Version 2.03.00 (2016-05-18, Kees Hiemstra)
# - Add Get-SOEComputerUserFirewallRules
# Version 2.02.02 (2015-12-29, Kees Hiemstra)
# - Update Get-SOEComputerSignedDrivers
# Version 2.02.01 (2015-12-22, Kees Hiemstra)
# - Update Get-SOEComputerNICDetails
# - Update Get-SOEComputerSignedDrivers
# - Update Get-SOEComputerBIOS
# Version 2.02.00 (2015-12-21, Kees Hiemstra)
# - Update Get-SOEComputerSignedDrivers
# - Add Get-SOEComputerBIOS
# Version 2.01.00 (2105-12-20, Kees Hiemstra)
# - Update Start-SOEServerService
# - Update Get-SOEComputerModel
# - Update Get-SOEComputerOS
# - Update Get-SOEComputerBSoDEvents
# - Update Get-SOEComputerNICDetails
# - Update Get-SOEComputerSignedDrivers
# - Update Get-SOEComputerRunKeys
# - Add Get-SOEComputerPatches
# Version 2.00.00 (2015-12-16, Kees Hiemstra)
# - Major update Get-SOEComputerBSoDEvents
# - Update Get-SOEComputerModel
# Version 1.04.01 (2015-12-14, Kees Hiemstra)
# - Update Get-SOEComputerRunKeys
# Version 1.04.00 (2015-12-13, Kees Hiemstra)
# - Add Get-SOEComputerRunKeys
# Version 1.03.00 (2015-12-09, Kees Hiemstra)
# - Update Get-SOEComputerModel
# - Add Get-SOEComputerNICDetails
# - Add Get-SOEComputerSignedDrivers
# Version 1.02.00 (2015-12-07, Kees Hiemstra)
# - Add Get-SOEComputerBSoDEvents
# Version 1.01.00 (2015-12-06, Kees Hiemstra)
# - Add Get-SOEComputerModel
# - Add Get-SOEComputerOS
# Version 1.00.00 (2014-09-09, Kees Hiemstra)
# - Add Start-SOEServerService

@{

RootModule = 'SOEComputer.psm1'
ModuleVersion = '4.00.01'
GUID = '18788c1f-b88f-49bb-be7f-c9576a1af0a6'
Author = 'Kees.Hiemstra@hpe.com'
CompanyName = 'Hewlett-Packard Enterprise'
Copyright = '(c) 2016 Kees Hiemstra. All rights reserved.'
Description = 'SOE computer related cmdlets'
PowerShellVersion = '3.0'
DotNetFrameworkVersion = '4.0'
RequiredModules = @('ActiveDirectory')
# ScriptsToProcess = @()
# TypesToProcess = @()
# FormatsToProcess = @()
# NestedModules = @('')
FunctionsToExport = @('Start-SOEServerService',
					  'Get-SOEComputerOS',
					  'Get-SOEComputerModel',
					  'Get-SOEComputerOS',
					  'Get-SOEComputerBSoDEvents',
					  'Get-SOEComputerNICDetails',
					  'Get-SOEComputerSignedDrivers',
					  'Get-SOEComputerRunKeys',
					  'Get-SOEComputerPatches',
					  'Get-SOEComputerBIOS',
					  'Get-SOEComputerUserFirewallRules',
					  'Get-SOEComputerPolicyFirewallRules',
					  'Get-SOEComputerWifiProfile',
					  'Get-SOEComputerWifiSignature'
					  )
# CmdletsToExport = '*'
# VariablesToExport = '*'
# AliasesToExport = '*'
FileList = @('SOEComputer.psm1')
# PrivateData = ''
# HelpInfoURI = ''
# DefaultCommandPrefix = ''
}

