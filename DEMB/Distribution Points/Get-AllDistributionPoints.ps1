<#
    List all distribution points.
#>

Get-ADGroupMember "CN=HP-SCCM-Servers,OU=Server Groups,OU=Groups,OU=HP,OU=Support,DC=corp,DC=demb,DC=com" |
    Get-ADComputer -Properties Name, c, l, Type |
    Where-Object { $_.Type -like 'DP*' } |
    Sort-Object c, l, Name |
    Select-Object Name, c, l |
    ConvertTo-Csv -Delimiter "`t" -NoTypeInformation |
    Clip

