$ADProperties = @('EmployeeID', 'SAMAccountName', 'c', 'Company', 'Enabled', 'Comment', 'ExtensionAttribute5','LastLogonDate', 'WhenCreated', 'WhenChanged', 'Title', 'ExtensionAttribute15')
$MailSplatting = @{'SmtpServer' = 'smtp.corp.demb.com'; 'From' = 'Kees.Hiemstra@JDEcoffee.com'; 'To' = 'Kees.Hiemstra@hpe.com' }

$Old = (Get-Date).AddDays(-90).Date

$List = Get-ADUser -Filter { Enabled -eq $true -and LastLogonDate -lt $Old -and Company -eq 'HPE' -and ExtensionAttribute5 -ne 'Service' -and ExtensionAttribute5 -ne 'Mailbox' -and DistinguishedName -notlike '*,OU=DeskSide,OU=Users,OU=HP,OU=Support,DC=corp,DC=demb,DC=com' } -Properties $ADProperties |
    Select-Object $ADProperties |
    ConvertTo-Csv -Delimiter "`t" -NoTypeInformation |
    Clip

break

<# All enabled DeskSide team members #>

$ADProperties = @('EmployeeID', 'SAMAccountName', 'DisplayName', 'c', 'Company', 'Enabled', 'Comment', 'ExtensionAttribute5','LastLogonDate', 'WhenCreated', 'WhenChanged', 'Title', 'ExtensionAttribute15')
$MailSplatting = @{'SmtpServer' = 'smtp.corp.demb.com'; 'From' = 'Kees.Hiemstra@JDEcoffee.com'; 'To' = 'Kees.Hiemstra@hpe.com' }

$Subject = 'DeskSide team members'
#Get-ADUser -Filter { Enabled -eq $true -and Title -eq 'HPE DeskSide Team Member' -or DistinguishedName -like '*,OU=DeskSide,OU=Users,OU=HP,OU=Support,DC=corp,DC=demb,DC=com' } -Properties $ADProperties |
Get-ADUser -Filter { Enabled -eq $true -and DistinguishedName -like '*OU=DeskSide*' } -Properties $ADProperties |
    Sort-Object C, L, DisplayName |
    Select-Object $ADProperties -ExcludeProperty @('EmployeeID', 'Company', 'ExtensionAttribute5', 'Title', 'Enabled', 'Comment') |
    Send-ObjectAsHTMLTableMessage -Subject $Subject @MailSplatting -MessageType Report

break

$ADProperties = @('EmployeeID', 'SAMAccountName', 'c', 'Company', 'Enabled', 'Comment', 'ExtensionAttribute5','LastLogonDate', 'WhenCreated', 'WhenChanged', 'Title', 'ExtensionAttribute15')

$Subject = 'Disabled Active employees'
$Exception = @()
if ( Test-Path -Path "$PSScriptRoot\$Subject.csv" )
{
    $Exception = Import-Csv -Path "$PSScriptRoot\$Subject.csv" | Where-Object { $_.WhenDismissed -ge (Get-Date) }
}
Get-ADUser -Filter { EmployeeID -like '*' -and Comment -eq 'Active' -and Enabled -eq $false } -Properties $ADProperties |
    Where-Object { $_.SAMAccountName -notin ($Exception).SAMAccountName } |
    Select-Object $ADProperties |
    Sort-Object C, LastLogonDate |
    Send-ObjectAsHTMLTableMessage -Subject $Subject @MailSplatting

$Subject = 'Non-Employees with comment'
$Exception = @()
if ( Test-Path -Path "$PSScriptRoot\$Subject.csv" )
{
    $Exception = Import-Csv -Path "$PSScriptRoot\$Subject.csv" | Where-Object { $_.WhenDismissed -ge (Get-Date) }
}
Get-ADUser -Filter { EmployeeID -notlike '*' -and Comment -like '*' } -Properties $ADProperties |
    Where-Object { $_.SAMAccountName -notin ($Exception).SAMAccountName } |
    Select-Object $ADProperties |
    Sort-Object C, LastLogonDate |
    Send-ObjectAsHTMLTableMessage -Subject $Subject @MailSplatting

$Subject = 'Inactive employees with LastLogonDate'
$Exception = @()
if ( Test-Path -Path "$PSScriptRoot\$Subject.csv" )
{
    $Exception = Import-Csv -Path "$PSScriptRoot\$Subject.csv" | Where-Object { $_.WhenDismissed -ge (Get-Date) }
}
Get-ADUser -Filter { EmployeeID -like '*' -and Comment -notlike '*' -and LastLogonDate -like '*' } -Properties $ADProperties |
    Where-Object { $_.SAMAccountName -notin ($Exception).SAMAccountName -and $_.Company -notin ('HPE') } |
    Select-Object $ADProperties |
    Sort-Object C, LastLogonDate |
    Send-ObjectAsHTMLTableMessage -Subject $Subject @MailSplatting

$Subject = 'Active employees without company'
Get-ADUser -Filter { EmployeeID -like '*' -and Company -notlike '*' -and LastLogonDate -like '*' } -Properties $ADProperties |
    Select-Object $ADProperties |
    Sort-Object C, LastLogonDate |
    Send-ObjectAsHTMLTableMessage -Subject $Subject @MailSplatting

$Subject = 'Active employees with company but without department number'
Get-ADUser -Filter { EmployeeID -like '*' -and Comany -like '*' -and DepartmentNumber -notlike '*' -and LastLogonDate -like '*' } -Properties $ADProperties |
    Select-Object $ADProperties |
    Sort-Object C, LastLogonDate |
    Send-ObjectAsHTMLTableMessage -Subject $Subject @MailSplatting

$Subject = 'Non-employees with company but without department number'
Get-ADUser -Filter { EmployeeID -notlike '*' -and Comany -like '*' -and DepartmentNumber -notlike '*' } -Properties $ADProperties |
    Select-Object $ADProperties |
    Sort-Object C, LastLogonDate |
    Send-ObjectAsHTMLTableMessage -Subject $Subject @MailSplatting

break

$ToBeDeleted = (Get-ADGroup -Filter * -SearchBase 'OU=Disabled Accounts,OU=To be deleted,OU=Support,DC=corp,DC=demb,DC=com' -Properties Member |
    Where-Object { $_.Name.Length -eq 17 }).Member

$Subject = 'Wrong registered account where EA15 suggest a shared mailbox'
$Shared = Get-ADUser -Filter { Enabled -eq $false -and ExtensionAttribute5 -ne 'mailbox' -and ExtensionAttribute15 -like '*share*mail*' } -Properties ($ADProperties + @('Mail', 'UserPrincipalName')) |
    Where-Object { $_.DistinguishedName -notlike '*,OU=Resources,OU=Exchange,OU=Support,DC=corp,DC=demb,DC=com' -and
        $_.DistinguishedName -notin $ToBeDeleted } |
    Select-Object SAMAccountName, c, Company, Enabled, WhenCreated, WhenChanged, Mail, UserPrincipalName, ExtensionAttribute15, @{n='Location'; e={ Convert-DistinguishedNameToPartition -DistinguishedName $_.DistinguishedName }}

#$Shared | Send-ObjectAsHTMLTableMessage -Subject $Subject @MailSplatting

$Subject = 'Wrong registered account where EA15 suggest a room or equipment mailbox'
$Room = Get-ADUser -Filter { Enabled -eq $false -and ExtensionAttribute5 -ne 'mailbox' -and (ExtensionAttribute15 -like '*room*mail*' -or ExtensionAttribute15 -like '*equ*mail*') } -Properties ($ADProperties + @('Mail', 'UserPrincipalName')) |
    Where-Object { $_.DistinguishedName -notlike '*,OU=Resources,OU=Exchange,OU=Support,DC=corp,DC=demb,DC=com' -and
        $_.DistinguishedName -notin $ToBeDeleted } |
    Select-Object SAMAccountName, c, Company, Enabled, WhenCreated, WhenChanged, Mail, UserPrincipalName, ExtensionAttribute15, @{n='Location'; e={ Convert-DistinguishedNameToPartition -DistinguishedName $_.DistinguishedName }}

$Room | Send-ObjectAsHTMLTableMessage -Subject $Subject @MailSplatting

$Subject = 'Inproper disabled user accounts'
Get-ADUser -Filter { Enabled -eq $false -and ExtensionAttribute5 -ne 'mailbox' } -Properties ($ADProperties + @('MemberOf')) |
    Where-Object { $_.DistinguishedName -notlike '*,OU=Resources,OU=Exchange,OU=Support,DC=corp,DC=demb,DC=com' -and
        $_.DistinguishedName -notin $ToBeDeleted -and
        $_.SAMAccountName -notin ($Shared + $Room).SAMAccountName } |
    Select-Object ($ADProperties + @{n='Location'; e={ Convert-DistinguishedNameToPartition -DistinguishedName $_.DistinguishedName }}) |
    Send-ObjectAsHTMLTableMessage -Subject $Subject @MailSplatting

