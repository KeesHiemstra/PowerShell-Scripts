Get-ADComputer -Filter * -SearchBase "OU=US,OU=Win7,OU=Managed Computers,DC=corp,DC=demb,DC=com" -Properties Name, LastLogonDate, SerialNumber, IPv4Address, Enabled, Description, Street, WhenChanged, WhenCreated, DistinguishedName |
    Select-Object Name, LastLogonDate, @{n='SerialNumber'; e={ $_.SerialNumber -join ";" }}, IPv4Address, Enabled, Description, Street, WhenChanged, WhenCreated, DistinguishedName |
    ConvertTo-Csv -Delimiter "`t" -NoTypeInformation | Clip

break

