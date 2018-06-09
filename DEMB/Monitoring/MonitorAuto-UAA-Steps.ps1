<#
    MonitorAuto-UAA-Steps.ps1

    Check if accounts are stuck in some of the steps.

    === Version history
    Version 1.00 (2016-12-30, Kees Hiemstra)
    - Initial version.
#>

#region Check stuck accounts on missing
$ProcessPath = '\\DEMBMCIS168.corp.demb.com\D$\Scripts\Auto-UAA\Process'

$Files = Get-ChildItem -Path "$ProcessPath\*.txt" | Where-Object { $_.Length -gt 1KB }

if ( $Files.count -eq 0 ) { break }

$List = @()

foreach ( $Item in $Files )
{
    $ADUser = Get-ADUser -Identity ($Item.Name -replace '\.\w*\.txt$', '') -Properties Mail, Comment, EmployeeID, LastLogonDate, WhenCreated, WhenChanged -ErrorAction Stop |
        Select-Object EmployeeID, SAMAccountName, Enabled, LastLogonDate, WhenCreated, WhenChanged, @{n='Step'; e={ '' }}

    if ( $ADUser -ne $null )
    {
        $ADUser.Step = $Item.Name -replace "$($ADUser.SAMAccountName)\." -replace '\.txt$'
        $List += $ADUser
    }
    else
    {
        $Properties = @{'EmployeeID' = ''; 'SAMAccountName' = $Item.Name; 'Enabled' = 'n/a'; 'LastLogonDate' = $null; 'WhenCreated' = $null; 'WhenChanged' = $null; 'Step' = 'n/a' }
        $Obj = New-Object -TypeName PSObject -Property $Properties
        $List += $Obj
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
$MsgTitle = 'Auto-UAA failing steps'

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
