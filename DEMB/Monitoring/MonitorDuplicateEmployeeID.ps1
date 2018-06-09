<#
    Monitor duplicate EmployeeIDs

    === Version history
    Version 1.10 (2016-10-17, Kees Hiemstra)
    - Include to-be-deleted information.
    Version 1.00 (2016-10-13, Kees Hiemstra)
    - Initial version.
#>


$Duplicates = Get-ADUser -Filter { EmployeeID -like '*' } -Properties EmployeeID |
    Group-Object EmployeeID |
    Where-Object { $_.Count -gt 1 } |
    Select-Object @{n='EmployeeID'; e={ $_.Name }} |
    Sort-Object EmployeeID


if ( $Duplicates.Count -eq 0 ) { break }

$Header = @"
<style>
BODY{background-color:peachpuff;}
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:thistle}
TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:PaleGoldenrod}
</style>
"@
$MsgTitle = 'Duplicate employeeIDs'

$ADUser = @()

foreach ( $Item in $Duplicates )
{
    try
    {
        $EmployeeID = $Item.EmployeeID
        $ADUser += Get-ADUSer -Filter { EmployeeID -eq $EmployeeID } -Properties Mail, Comment, EmployeeID, LastLogonDate, WhenCreated, ExtensionAttribute15, MemberOf -ErrorAction Stop |
            Select-Object EmployeeID, SAMAccountName, Enabled, @{n='To-be-del'; e={ $_.MemberOf -match 'CN=To-Be-Deleted-' }}, LastLogonDate, WhenCreated, ExtensionAttribute15
    }
    catch
    {
    }
}


[string]$Message = $ADUser | ConvertTo-Html -Title $MsgTitle -Head $Header -Body "<H2>$MsgTitle (#$($ADUser.Count))</H2>"

if ( $false )
{
    $Message | Out-File -FilePath "$($env:TEMP)\$MsgTitle.html"
    Invoke-Item "$($env:TEMP)\$MsgTitle.html"
}
else
{
    Send-MailMessage -Body $Message -BodyAsHtml -SmtpServer 'smtp.corp.demb.com' -From 'Kees.Hiemstra@JDEcoffee.com' -To 'Kees.Hiemstra@hpe.com' -Subject $MsgTitle
}

