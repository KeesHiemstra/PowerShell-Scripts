﻿<#
    MonitorAuto-UAA.ps1

    Check if accounts are stuck because the manager is still unknown in AD.

    === Version history
    Version 1.10 (2016-12-27, Kees Hiemstra)
    - Delete .Manager.txt file if the user has been disabled or has a logon date.
    Version 1.00 (2016-12-14, Kees Hiemstra)
    - Initial version.
#>

#region Check stuck accounts on missing
$ProcessPath = '\\DEMBMCIS168.corp.demb.com\D$\Scripts\Auto-UAA\Process'
$CheckTime = (Get-Date).AddHours(-3)

$Files = Get-ChildItem -Path "$ProcessPath\*.manager.txt" | Where-Object { $_.LastWriteTime -lt $CheckTime }

if ( $Files.count -eq 0 ) { break }

$List = @()

foreach ( $Item in $Files )
{
    $ADUser = Get-ADUser -Identity ($Item.Name -replace '.manager.txt', '') -Properties Mail, Comment, EmployeeID, LastLogonDate, WhenCreated, WhenChanged, ExtensionAttribute15 -ErrorAction Stop |
        Select-Object EmployeeID, SAMAccountName, Enabled, LastLogonDate, WhenCreated, WhenChanged, ExtensionAttribute15

    if ( $ADUser.Enabled -and $ADUser.LastLogonDate -eq $null )
    {
        $List += $ADUser
    }
    else
    {
        Remove-Item -Path $Item.FullName
    }
}

$Header = @"
<style>
BODY{background-color:peachpuff;}
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:thistle}
TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:PaleGoldenrod}
</style>
"@
$MsgTitle = 'Auto-UAA account(s) without manager'

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
#endregion
