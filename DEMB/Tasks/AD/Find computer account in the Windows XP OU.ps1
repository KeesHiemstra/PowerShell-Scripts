(Get-ADComputer -Filter * -SearchBase "OU=XP,OU=Managed Computers,DC=corp,DC=demb,DC=com").Count
(Get-ADComputer -Filter {OperatingSystem -like '*XP*'}).Count

break
<#
    XP OU
#>
Get-ADComputer -Filter * -SearchBase "OU=XP,OU=Managed Computers,DC=corp,DC=demb,DC=com" -Properties Name, LastLogonDate, IPv4Address, Enabled, OperatingSystem, Description, Street, WhenChanged, WhenCreated, DistinguishedName |
    Select-Object Name, LastLogonDate, IPv4Address, Enabled, OperatingSystem, Description, WhenChanged, WhenCreated, DistinguishedName |
    ConvertTo-Csv -Delimiter "`t" -NoTypeInformation | Clip

break

Get-ADComputer -Filter {OperatingSystem -like '*XP*'} -Properties Name, LastLogonDate, IPv4Address, Enabled, Description, Street, WhenChanged, WhenCreated, DistinguishedName |
    Select-Object Name, LastLogonDate, IPv4Address, Enabled, Description, WhenChanged, WhenCreated, DistinguishedName |
    ConvertTo-Csv -Delimiter "`t" -NoTypeInformation | Clip

