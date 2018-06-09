<#
    Monitor changes to disabled computers.

    version 1.00 (2016-05-09, Kees Hiemstra)
    - initial version.
#>

$Data = Get-ADComputer -Filter { Enabled -eq $false } -Properties Name, Description, Info, Gecos, LastLogonDate | Select-Object Name, Description, Info, Gecos, LastLogonDate

$DataPath = 'D:\Data\Jobs\Disabled'
$PrevFile = "$DataPath\Computer.csv"

if ( -not (Test-Path -Path $PrevFile) )
{
    $Data | Export-Csv -Path $PrevFile -NoTypeInformation -Delimiter "`t"
    Copy-Item -Path $PrevFile -Destination "$DataPath\Computer-$((Get-Date).ToString('yyyy-MM-dd HHmm')).csv"
    break
}

$PrevData = Import-Csv -Path $PrevFile -Delimiter "`t"

$Diff = Compare-Object -ReferenceObject $Data -DifferenceObject $PrevData -Property Name -PassThru

if ( $Diff.Count -eq 0 )
{
    break
}

$Header = @"
<style>
BODY{background-color:peachpuff;}
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:thistle}
TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:PaleGoldenrod}
</style>
"@
$MsgTitle = 'Changes in disabled computers'

$HTML = ""

#Get the details to report on
[array]$Added = $Diff | Where-Object { $_.SideIndicator -eq "<=" }
[array]$Deleted = $Diff | Where-Object { $_.SideIndicator -eq "=>" }

#Report
$HTML = "<p>The following changes have been reported</p>`n`n"
if ($Added.Count -gt 0)
{
    $HTML = $Added | ConvertTo-Html -Fragment -Property Name, Description, Info, Gecos, LastLogonDate -PreContent "$HTML<p>Added<p>`n"
}
if ($Deleted.Count -eq 0)
{
    $HTMLSub = ""
}
else
{
    $HTMLSub = "<p>Deleted</p>"

    foreach ( $Item in $Deleted )
    {
        try
        {
            $Name = $Item.Name
            $ADObject = Get-ADObject -Filter { Name -eq $Name }

            if ( $ADObject -ne $null -and $ADObject.Enabled )
            {
                Add-Member -InputObject $Item -MemberType NoteProperty -Name 'ActionInfo' -Value 'Has been enabled'
            }
            else
            {
                Add-Member -InputObject $Item -MemberType NoteProperty -Name 'ActionInfo' -Value 'Has been deleted'
            }
        }
        catch
        {
            Add-Member -InputObject $Item -MemberType NoteProperty -Name 'ActionInfo' -Value 'Cannot be found'
        }
    }
}
$HTML = $Deleted | ConvertTo-Html -Property Name, Description, Info, Gecos, LastLogonDate, ActionInfo -PreContent "$HTML`n$HTMLSub`n" -Title "Changes in disabled computers" -Head $Header

Send-MailMessage -Body "$HTML" -Subject $MsgTitle -To "Kees.Hiemstra@hpcds.com" -From "HPDesktop.Administrator@jdecoffee.com" -SmtpServer "smtp.corp.demb.com" -BodyAsHtml

Copy-Item -Path $PrevFile -Destination "$DataPath\Computer-$((Get-Date).ToString('yyyy-MM-dd HHmm')).csv"
$Data | Export-Csv -Path $PrevFile -NoTypeInformation -Delimiter "`t" -Force
