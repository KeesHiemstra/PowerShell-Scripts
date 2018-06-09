<#
    Monitor withdrawn AD Accounts that are not disabled.

    === Version history
    Version 1.02 (2016-12-27, Kees Hiemstra)
    - Add country and company information.
    Version 1.01 (2016-08-28, Kees Hiemstra)
    - Bug fix: Get-AzureLicense fix.
    Version 1.00 (2016-08-26, Kees Hiemstra)
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
$MsgTitle = 'Withdrawn accounts that are not disabled'

Connect-MsolService -Credential (Get-SOECredential -UserName UserCreation@coffeeandtea.onmicrosoft.com)

$ADUser = Get-ADUSer -Filter { EmployeeID -like '*' -and Comment -eq 'Withdrawn' -and Enabled -eq $true } -Properties Mail, Comment, EmployeeID, c, Company, ExtensionAttribute15, LastLogonDate, WhenCreated, UserPrincipalName |
    Select-Object EmployeeID, SAMAccountName, c, Company, LastLogonDate, WhenCreated, @{n='Licenses'; e={ Get-AzureLicense -userPrincipalName $_.userPrincipalName }}, @{n='ExtensionAttribute15'; e={ $_.ExtensionAttribute15.PadRight(33, ' ').Substring(0, 33).Trim() }} |
    Sort-Object LastLogonDate, WhenCreated

if ( $ADUser.Count -eq 0 )
{
    exit
}

[string]$Message = $ADUser | ConvertTo-Html -Title $MsgTitle -Head $Header -Body "<H2>$MsgTitle (#$($ADUser.Count))</H2>"

if ( $PSScriptRoot -ne 'C:\Etc\Jobs' )
{
    $Message | Out-File -FilePath "$($env:TEMP)\$MsgTitle.html"
    Invoke-Item "$($env:TEMP)\$MsgTitle.html"
}
else
{
    Send-MailMessage -Body $Message -BodyAsHtml -SmtpServer 'smtp.corp.demb.com' -From 'Kees.Hiemstra@JDEcoffee.com' -To 'Kees.Hiemstra@hpe.com' -Subject $MsgTitle
}


