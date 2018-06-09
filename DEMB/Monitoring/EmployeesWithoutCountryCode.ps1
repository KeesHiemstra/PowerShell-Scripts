$CheckDate = (Get-Date).AddHours(-3)

$List = Get-ADUser -Filter { EmployeeID -like '*' -and (c -notlike '*' -or c -eq 'CK') -and WhenCreated -lt $CheckDate -and Enabled -eq $true } -Properties c, EmployeeID, Company, LastLogonDate, WhenCreated, WhenChanged -SearchBase 'OU=Managed Users,DC=corp,DC=demb,DC=com' |
    Select-Object EmployeeID, SAMAccountName, c, Company, LastLogonDate, WhenCreated, WhenChanged |
    Sort-Object Company, EmployeeID


if ( $List.Count -eq 0 ) { break }

$Header = @"
<style>
BODY{background-color:peachpuff;}
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:thistle}
TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:PaleGoldenrod}
</style>
"@
$MsgTitle = 'Users with invalid country code'

[string]$Message = $List | ConvertTo-Html -Title $MsgTitle -Head $Header -Body "<H2>$MsgTitle (#$($List.Count))</H2>"

if ( $PSScriptRoot -ne 'C:\Etc\Jobs' )
{
    #Test page in browser
    $Message | Out-File -FilePath "$($env:TEMP)\$MsgTitle.html"
    Invoke-Item "$($env:TEMP)\$MsgTitle.html"
}
else
{
    Send-MailMessage -Body $Message -BodyAsHtml -SmtpServer 'smtp.corp.demb.com' -From 'Kees.Hiemstra@JDEcoffee.com' -To 'Kees.Hiemstra@hpe.com' -Subject $MsgTitle
}


