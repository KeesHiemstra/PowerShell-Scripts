<# All enabled DeskSide team members #>

$ADProperties = @('EmployeeID', 'SAMAccountName', 'DisplayName', 'c', 'Company', 'Enabled', 'Comment', 'ExtensionAttribute5','LastLogonDate', 'WhenCreated', 'WhenChanged', 'Title', 'ExtensionAttribute15')
$MailSplatting = @{'SmtpServer' = 'smtp.corp.demb.com'; 'From' = 'Kees.Hiemstra@JDEcoffee.com'; 'To' = 'Kees.Hiemstra@hpe.com' }

$Subject = 'DeskSide team members'
$List = Get-ADUser -Filter { Enabled -eq $true } -Properties $ADProperties -SearchBase 'OU=DeskSide,OU=Users,OU=HP,OU=Support,DC=corp,DC=demb,DC=com' |
    Select-Object *, @{n='Type'; e={ 'OU' }}

$List += Get-ADUser -Filter { Enabled -eq $true -and Title -eq 'HPE DeskSide Team Member' } -Properties $ADProperties |
    Where-Object { $_.SAMAccountName -notin $List.SAMAccountName } |
    Select-Object *, @{n='Type'; e={ 'Title' }}

$List += (Get-ADGroup 'HP-Deskside-EMEA' -Properties Member).Member |
    Get-ADUser -Properties $ADProperties |
    Where-Object { $_.Enabled } |
    Where-Object { $_.SAMAccountName -notin $List.SAMAccountName } |
    Select-Object *, @{n='Type'; e={ 'Group' }}

$List |
    Sort-Object C, L, DisplayName |
    Select-Object ($ADProperties+@('Type')) -ExcludeProperty @('EmployeeID', 'Company', 'ExtensionAttribute5', 'Title', 'Enabled', 'Comment') |
    Send-ObjectAsHTMLTableMessage -Subject $Subject @MailSplatting -MessageType Report

break
(Get-ADGroup 'HP-Deskside-EMEA' -Properties Member).Member.Count

((Get-ADGroup 'HP-Deskside-EMEA' -Properties Member).Member |
    Get-ADUser -Properties $ADProperties |
    Where-Object { $_.Enabled }).Count
