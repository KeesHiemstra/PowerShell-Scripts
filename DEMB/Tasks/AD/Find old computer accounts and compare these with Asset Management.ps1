<#

#>

$CheckDate = (Get-Date).AddDays(-91)

$List = Get-ADComputer -Filter { LastLogonDate -lt $CheckDate } -Properties Name, SerialNumber, Street, LastLogonDate, Enabled, Description, OperatingSystem -SearchBase 'OU=Win7,OU=Managed Computers,DC=corp,DC=demb,DC=com' |
    Select-Object Name, @{n='SerialNumber'; e={ $_.SerialNumber -join '; ' }}, @{n='Users'; e={ $_.Street -split ';' }}, LastLogonDate, Enabled, Description, OperatingSystem

foreach ( $Item in $List )
{
    $Asset = Get-AMAsset -ComputerName $Item.Name

    Add-Member -InputObject $Item -MemberType NoteProperty -Name 'SerialNo' -Value $Asset.SerialNo
    Add-Member -InputObject $Item -MemberType NoteProperty -Name 'BillingStatus' -Value $Asset.BillingStatus
    Add-Member -InputObject $Item -MemberType NoteProperty -Name 'AssetStatus' -Value $Asset.AssetStatus
    Add-Member -InputObject $Item -MemberType NoteProperty -Name 'UserAccount' -Value $Asset.UserAccount
}