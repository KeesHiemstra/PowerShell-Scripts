$MailSplatting = @{'SmtpServer' = 'smtp.corp.demb.com'; 'From' = 'Kees.Hiemstra@JDEcoffee.com'; 'To' = 'Kees.Hiemstra@hpe.com' }

$ReportProperty = @('Comment', 'Company', 'ExtensionAttribute5', 'c', 'co', 'l')

#$ADUser = Get-ADUser -Filter * -Properties $ReportProperty
#$ADUser = Get-ADUser -Filter * -Properties *

foreach ( $Item in $ReportProperty )
{
    $Subject = "Report on property: $Item"
    Write-Host "$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) $ubject"

    $ADUser |
        Select-Object $Item |
        Group-Object $Item |
        Sort-Object Name |
        Select @{n=$Item; e={ $_.Name }}, Count |
        Send-ObjectAsHTMLTableMessage -Subject $Subject @MailSplatting
}

$Subject = "Report on site"

$ADUser |
    Select-Object c, co, l |
    Group-Object c, co, l |
    Sort-Object Name |
    Select @{n='Site'; e={ $_.Name }}, Count |
    Send-ObjectAsHTMLTableMessage -Subject $Subject @MailSplatting
