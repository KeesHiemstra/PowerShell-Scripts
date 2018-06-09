<#
    Monitor AD Accounts with Expiration date.

    === Version history
    Version 1.10 (2016-09-24, Kees Hiemstra)
    - Introduced an exception list with SAMAccountNames that are allouwed to have an expire date.
    Version 1.01 (2016-08-26, Kees Hiemstra)
    - Improved the message readability.
    version 1.00 (2016-08-07, Kees Hiemstra)
    - initial version.
#>

$Header = @"
<style>
BODY{background-color:peachpuff;}
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:thistle}
TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:PaleGoldenrod}
</style>
"@
$MsgTitle = 'Accounts with expiration date'

if ( (Test-Path -Path ($MyInvocation.MyCommand.Definition -replace(".ps1$",".csv") ) ) )
{
    $Exceptions = Import-Csv -Path ($MyInvocation.MyCommand.Definition -replace(".ps1$",".csv") )
}
else
{
    $Exceptions = @{}
}

$ADUser = Get-ADUser -Filter { AccountExpirationDate -like '*' } -Properties AccountExpirationDate, WhenCreated, WhenChanged |
    Where-Object { $_.SAMAccountName -notin $Exceptions.SAMAccountName }

if ( $ADUser.Count -eq 0 )
{
    exit
}

[string]$Message = $ADUser |
    Select SAMAccountName, WhenCreated, WhenChanged, AccountExpirationDate |
    ConvertTo-Html -Title $MsgTitle -Head $Header -Body "<H2>$MsgTitle</H2>"

if ($false)
{
    $Message | Out-File -FilePath "$($env:TEMP)\$MsgTitle.html"
    Invoke-Item "$($env:TEMP)\$MsgTitle.html"
}
else
{
    Send-MailMessage -Body $Message -BodyAsHtml -SmtpServer 'smtp.corp.demb.com' -From 'Kees.Hiemstra@JDEcoffee.com' -To 'Kees.Hiemstra@hpe.com' -Subject $MsgTitle
}
