<#
    Send-StatusUpdate.ps1

    Send a message to home with some information.

    --- Version history
    Version 1.00 (2016-05-03, Kees Hiemstra)
    - Initial version.
#>

. "$($PSScriptRoot)\CommonEnvironment.ps1"

$ScriptName = "Send-StatusUpdate"

Write-Log -Message "Status update required"

$MailBody = "Status update has been requested`n"

$MailBody += "Running Bits transfers:`n"
$BitsTransfer = Get-BitsTransfer -AllUsers

if ( $BitsTransfer -eq $null -or $BitsTransfer.Count -eq 0 )
{
    $MailBody += "- None`n"
}
else
{
    foreach ($Download in $BitsTransfer)
    {
        $MailBody += ("- $($Download.Description) at {0:P} ($($Download.JobState))`n" -f ($Download.BytesTransferred / $Download.BytesTotal))
    }
}

Send-SOEMessage -Subject "Server $($Env:ComputerName) status update" -Body $MailBody -MailType "status"

Stop-Script -Silent