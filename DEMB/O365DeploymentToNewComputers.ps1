<#
Set initial groups to new computers

Newly created computer objects in AD will be added to the group: CN=O365 deployment to new computers,OU=Groups,OU=SoE,DC=corp,DC=demb,DC=com
in order to have Office 365 automatically deployed in the same way as this has been done with the sequoia computers and 
CN=GPO-C-BitLocker pilot,OU=SCCM,OU=Software Assignment,DC=corp,DC=demb,DC=com.

Author : Kees.Hiemstra
Date   : 2015-11-19
Version: 2.1

Version history:
2.1 (2015-11-19, Kees Hiemstra)
- Kazakstan needs to be included for encryption.
2.0 (2015-11-18, Kees Hiemstra)
- Added new computers to "GPO-C-BitLocker pilot" group, but exclude computers from:
    Belarus (BY)
    China (CN)
    France (FR)
    Georgia (GE)
    Kazakhstan (KZ)
    Morocco (MA)
    Russia (RU)
    Ukraine (UA)
    Thailand (TH)
    Hong Kong (HK)
    Indonesia (ID)
- Restucture process to be able to process BitLocker.
1.0 (2015-07-08, Kees Hiemstra)
- Inital version
#>

$Groups = [ordered]@{D_O365=(Get-ADGroup -Filter { Name -eq 'O365 deployment to new computers' }).DistinguishedName
    D_BITL=(Get-ADGroup -Filter { Name -eq 'GPO-C-BitLocker pilot' }).DistinguishedName
    }

$BitLockerExclusions = ('BY', 'CN', 'FR', 'GE', 'MA', 'RU', 'UA', 'TH', 'HK', 'ID', '_Test')

#Collect new computers added to AD
$CheckTime = (Get-Date).AddHours(-6)
$NewComputers = Get-ADComputer -Filter {createTimeStamp -gt $CheckTime} -SearchBase "OU=Win7,OU=Managed Computers,DC=corp,DC=demb,DC=com" -Properties MemberOf |
    Where-Object { $_.MemberOf.Count -eq 0 } |
    Select-Object *, @{n='CountryCode'; e={ $_.DistinguishedName -replace ",OU=Win7,OU=Managed Computers,DC=corp,DC=demb,DC=com$" -replace "^.*tops,OU=" }}

#Add new computers to the Office 365 deployment group
$NewComputers | Where-Object { $_.CountryCode -notin $BitLockerExclusions } |
    ForEach-Object { Add-ADGroupMember -Identity $Groups['D_O365'] -Members $_.DistinguishedName }


#Add new computers to the BitLocker GPO and deployment group and exclude forbidden countries
$NewComputers | Where-Object { $_.CountryCode -notin $BitLockerExclusions } |
    ForEach-Object { Add-ADGroupMember -Identity $Groups['D_BITL'] -Members $_.DistinguishedName }
