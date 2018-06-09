<#
    Collect computer names that are excluded from BitLocker

    --- Version History
    Version 1.00 (2016-11-03, Kees Hiemstra)
    - Initial version.
#>

$BitLocker = (Get-ADGroup 'GPO-C-BitLocker Off' -Properties Member).Member | Get-ADComputer

$BitLocker += (Get-ADGroup 'Microsoft MBAM Client 2.5 SP1 EN DEMB675 - Exclusions' -Properties Member).Member | Get-ADComputer
$BitLocker += (Get-ADGroup 'GPO-C-BitLocker off' -Properties Member).Member | Get-ADComputer

$BitLocker += Get-ADComputer -Filter * -SearchBase 'OU=CN,OU=Win7,OU=Managed Computers,DC=corp,DC=demb,DC=com'
$BitLocker += Get-ADComputer -Filter * -SearchBase 'OU=HK,OU=Win7,OU=Managed Computers,DC=corp,DC=demb,DC=com'
$BitLocker += Get-ADComputer -Filter * -SearchBase 'OU=RU,OU=Win7,OU=Managed Computers,DC=corp,DC=demb,DC=com'


$BitLocker | Group-Object Name | Select-Object @{n='ComputerName'; e={ $_.Name }} | Sort-Object ComputerName | ConvertTo-Csv -NoTypeInformation | Clip


