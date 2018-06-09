<#
    EmployeesWithoutValidManager.ps1

    Reporting on users that don't have a manager or where the manager is the default position.

    Exceptions are collect in the EmployeesWithoutValidManager.csv in the same folder.

    === Version history
    Version 1.00 (2016-12-27, Kees Hiemstra)
    - Initial version.
#>

$CheckDate = (Get-Date).AddDays(-28)

if ( (Test-Path -Path ($MyInvocation.MyCommand.Definition -replace(".ps1$",".csv") ) ) )
{
    $Exceptions = Import-Csv -Path ($MyInvocation.MyCommand.Definition -replace(".ps1$",".csv") )
}
else
{
    $Exceptions = @{}
}

$List = Get-ADUser -Filter { Enabled -eq $true -and EmployeeID -like '*' -and WhenCreated -lt $CheckDate } -Properties Manager, EmployeeID, c, Company, LastLogonDate, WhenCreated, WhenChanged -SearchBase 'OU=Managed Users,DC=corp,DC=demb,DC=com' |
    Where-Object { $_.Manager -eq $null -or $_.Manager -eq 'CN=default.position,OU=Generic Accounts,OU=Support,DC=corp,DC=demb,DC=com' -and $_.EmployeeID -notin $Exceptions.EmployeeID } |
    Select-Object EmployeeID, SAMAccountName, c, Company, LastLogonDate, WhenCreated, WhenChanged |
    Sort-Object c, Company, EmployeeID

if ( $List.Count -eq 0 ) { break }

$Header = @"
<style>
BODY{background-color:peachpuff;}
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:thistle}
TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:PaleGoldenrod}
</style>
"@
$MsgTitle = 'Users without valid manager'

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

