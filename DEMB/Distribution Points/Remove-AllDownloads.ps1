<#
    Remove Bits download from all distribution points.
#>

Get-ADGroupMember "CN=HP-SCCM-Servers,OU=Server Groups,OU=Groups,OU=HP,OU=Support,DC=corp,DC=demb,DC=com" |
    Get-ADComputer -Properties Name, Type |
    Where-Object { $_.Type -like 'DP*' } |
    ForEach-Object { Write-Host $_.Name; SchTasks /Run /S $_.Name /TN "SOE\Remove downloads" }

