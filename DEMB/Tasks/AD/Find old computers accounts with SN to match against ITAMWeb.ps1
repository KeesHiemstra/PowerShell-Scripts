$ADComputers = Get-ADComputer -Filter { SerialNumber -like '*' -and SerialNumber -ne '0' -and SerialNumber -notlike '*-*' } -Properties Name, Description, DistinguishedName, Enabled, IPv4Address, LastLogonDate, SerialNumber, Street, WhenChanged, WhenCreated |
    Where-Object { $_.LastLogonDate -lt (Get-Date).AddDays(-273) }
    Select-Object Name, Description, DistinguishedName, Enabled, IPv4Address, LastLogonDate, @{n='SerialNumber'; e={$_.SerialNumber -join ';'}}, Street, WhenChanged, WhenCreated

$ADComputers.Count

$ADComputers | Select-Object Name, LastLogonDate, IPv4Address, Enabled, Description, Street, WhenChanged, WhenCreated, DistinguishedName, @{n='SerialNumber'; e={$_.SerialNumber -join ';'}} |
    ConvertTo-Csv -Delimiter "`t" -NoTypeInformation | Clip

#Match the result with ITAMWeb on SN and CN.