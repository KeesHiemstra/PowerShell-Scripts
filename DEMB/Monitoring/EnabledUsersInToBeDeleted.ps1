<#
    Monitor enabled users in To-Be-Deleted groups

    Version 2.00 (2017-01-09, Kees Hiemstra)
    - Restructure
    Version 1.00 (2016-10-28, Kees Hiemstra)
    - Initial version.
#>
$Subject = 'Enabled users in To-Be-Delete groups'

$ToBeDeleted = (Get-ADGroup -Filter * -SearchBase 'OU=Disabled Accounts,OU=To be deleted,OU=Support,DC=corp,DC=demb,DC=com' -Properties Member |
    Where-Object { $_.Name.Length -eq 17 }).Member

$List = @()
foreach ( $Item in $ToBeDeleted )
{
    $List += Get-ADUser $Item -Properties Mail, Comment, EmployeeID, LastLogonDate, WhenCreated, ExtensionAttribute15, MemberOf |
        Where-Object { $_.Enabled -eq $true } |
        Select-Object EmployeeID, SAMAccountName, Enabled, @{n='To-be-del'; e={ ($_.MemberOf -match 'CN=To-Be-Deleted-') -replace 'CN=To-Be-Deleted-' -replace ',OU=Disabled Accounts,OU=To be deleted,OU=Support,DC=corp,DC=demb,DC=com' -join '; ' }}, LastLogonDate, WhenCreated, ExtensionAttribute15
}

Write-Host "List [$Subject] contains $($List.Count) entries"

$List | Send-ObjectAsHTMLTableMessage -Subject $Subject -MessageType Exception -SmtpServer 'smtp.corp.demb.com' -From 'Kees.Hiemstra@JDEcoffee.com' -To 'Kees.Hiemstra@hpe.com'
