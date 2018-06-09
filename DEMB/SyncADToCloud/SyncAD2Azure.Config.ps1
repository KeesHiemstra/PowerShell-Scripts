<#
    SyncAD2Azure.Config.ps1
#>

#All change actions in AD or Azure will be skipped if RunInProduction = $false
$RunInProduction = $false    # Update Active Directory if true
$RunNoDatabase = $true       # Update database if false
$RunOffline = $true          # Get data from Active Directory if false
$RunInDevelopment = $false   # Get module from development environment if true

#Define log file variables
$ScriptName = $MyInvocation.MyCommand.Definition.Replace("$PSScriptRoot\", '').Replace(".Config.ps1", '')
$LogFile = $MyInvocation.MyCommand.Definition -replace(".Config.ps1$",".log")
$ErrorFile = $MyInvocation.MyCommand.Definition -replace(".Config.ps1$","-$((Get-Date).ToString('yyyyMMdd-HHmm')).log")

#Define path variables
$SKUFile = "$PSScriptRoot\PrevSKU.xml"

#$ConnectionString = 'Trusted_Connection=True;Data Source=DEMBRSAPS350SQ1;Initial Catalog=SOElog'
$ConnectionString = 'Trusted_Connection=True;Data Source=(Local);Initial Catalog=SOEAdmin'

#Mail settings
$MailSMTP = "smtp.corp.demb.com"
$MailSender = "$($ScriptName.Replace(' ', '.'))@JDECoffee.com"
$MailErrorTo = "Kees.Hiemstra@hpe.com"

#Default settings
$DefaultInternalOfficeTags = "[EMS][EXC][INT][O365][RMS][SHA][SWY][WAC][YAM]"

#List of groups used in combination of the licenses
if ( -not $RunOffline )
{
    $Groups = [ordered]@{A_EMS=(Get-ADGroup -Filter { Name -eq 'Allow-EMS-JDE-employee' }).DistinguishedName
        E_EMS=(Get-ADGroup -Filter { Name -eq 'Allow-EMS-Non-JDE-employee' }).DistinguishedName
        A_OFF=(Get-ADGroup -Filter { Name -eq 'Allow-ADFS-Access' }).DistinguishedName
        A_VPN=(Get-ADGroup -Filter { Name -eq 'Allow-VPN-Access' }).DistinguishedName
        D_EMS=(Get-ADGroup -Filter { Name -eq 'Deny-EMS-Access' }).DistinguishedName
        D_OFF=(Get-ADGroup -Filter { Name -eq 'Deny-O365-Access' }).DistinguishedName
        D_VPN=(Get-ADGroup -Filter { Name -eq 'Deny-VPN-Access' }).DistinguishedName
        }
}
else
{
    $Groups = [ordered]@{A_EMS='CN=Allow-EMS-JDE-employee,OU=Groups,DC=corp,DC=demb,DC=com'
        E_EMS='CN=Allow-EMS-Non-JDE-employee,OU=Groups,DC=corp,DC=demb,DC=com'
        A_OFF='CN=Allow-ADFS-Access,OU=Groups,DC=corp,DC=demb,DC=com'
        A_VPN='CN=Allow-VPN-Access,OU=Groups,DC=corp,DC=demb,DC=com'
        D_EMS='CN=Deny-EMS-Access,OU=Groups,DC=corp,DC=demb,DC=com'
        D_OFF='CN=Deny-O365-Access,OU=Groups,DC=corp,DC=demb,DC=com'
        D_VPN='CN=Deny-VPN-Access,OU=Groups,DC=corp,DC=demb,DC=com'
        }
}

#List of properties to collect from AD (mainly for reporting purposes)
$ADPropertyList = @('UserPrincipalName',
    'C',
    'Comment',
    'DisplayName',
    'DistinguishedName',
    'EmployeeID',
    'EmployeeType',
    'Enabled',
    'EmployeeID',
    'EmployeeType',
    'ExtensionAttribute2',
    'ExtensionAttribute5',
    'ExtensionAttribute11',
    'GivenName',
    'Mail',
    'MemberOf',
    'SAMAccountName',
    'SurName',
    'WhenCreated')

#EmployeeType for internal users
$EmployeeTypeInternalUser = @('Active Employee', 'Expats/ Inpats', 'Retiree/ Pensioner')

#ExtensionAttribute5 values to indicate it is a special account
$EA5Special = @('Generic', 'Service', 'Mailbox')

#ADLocations for internal users
$ADLocationForInternals = ('Managed Users,DC=corp,DC=demb,DC=com')

