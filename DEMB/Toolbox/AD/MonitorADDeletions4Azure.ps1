<#
    Script: MonitorADDeletions4Azure.ps1

    Monitor AD object deletions.
#>

$ClassesToCheck = @('user', 'contact', 'group')
$DeletionThreshold = 300
$TimeToCheck = (Get-Date).AddHours(-3)

#region LogFile 1.20
$Error.Clear()

#Use: $Error.RemoveAt(0) to remove an error or warning captured in a try - catch statement

#Define initial log file variables
$ScriptName = $MyInvocation.MyCommand.Name.Replace(".ps1", '')
$MailSMTP = "smtp.corp.demb.com"
$MailSender = "$($ScriptName.Replace(' ', '.'))@JDECoffee.com"
$MailErrorTo = "Kees.Hiemstra@hpe.com"

$LogFile = $MyInvocation.MyCommand.Definition -replace(".ps1$",".log")
$LogStart = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")

#The error mail to home will be send if $LogWarning is true but no other errors were reported.
#The subject of the mail will indicate Warning instead of Error.
$LogWarning = $false

#$VerbosePreference = 'SilentlyContinue' #Verbose off
#$VerbosePreference = 'Continue' #Verbose on

#Create the log file if the file does not exits else write today's first empty line as batch separator
if (-not (Test-Path $LogFile))
{
    New-Item $LogFile -ItemType file | Out-Null
}
else 
{
    Add-Content -Path $LogFile -Value "---------- --------"
}

function Write-Log([string]$Message)
{
    $LogMessage = "{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Add-Content -Path $LogFile -Value $LogMessage
    if ($VerbosePreference -eq 'Continue') { Write-Host $LogMessage }
    Write-Host $LogMessage
}

#Finish script and report on $Error
function Stop-Script([bool]$WithError)
{
    if ( $WithError )
    {
        Write-Log -Message "Script stopped with an initiated error"
    }
    else
    {
        Write-Log -Message "Script ended normally"
    }

    if ( $Error.Count -gt 0 -or $WithError -or $LogWarning)
    {
        if ( $Error.Count -gt 0 -or $WithError )
        {
            $MailErrorSubject = "Error in $ScriptName on $($Env:ComputerName)"

            if ( $Error.Count -gt 0 )
            {
                $MailErrorBody = "The script $ScriptName has reported the following error(s):`n`n"
                $MailErrorBody += $Error | Out-String
            }
            else
            {
                $MailErrorBody = "The script $ScriptName has reported error(s) in the log file."
            }
        }
        else
        {
            $MailErrorSubject = "Warning in $ScriptName on $($Env:ComputerName)"
            $MailErrorBody = "The script $ScriptName has reported warning(s) in the log file."
        }

        $MailErrorBody += "`n`n--- LOG FILE (extract) ----------------`n"
        $MailErrorBody += (Get-Content $LogFile | Where-Object { $_.SubString(0, 19) -ge $LogStart }) -join "`n"
        try
        {
            Send-MailMessage -SmtpServer $MailSMTP -From $MailSender -To $MailErrorTo -Subject $MailErrorSubject -Body $MailErrorBody -ErrorAction Stop
            Write-Log -Message "Sent error mail to home"
        }
        catch
        {
            Write-Log -Message "Retry sending the message after 15 seconds"
            Start-Sleep -Seconds 15

            try
            {
                Send-MailMessage -SmtpServer $MailSMTP -From $MailSender -To $MailErrorTo -Subject $MailErrorSubject -Body $MailErrorBody -ErrorAction Stop
                Write-Log -Message "Sent error mail to home after a retry"
            }
            catch
            {
                Write-Log -Message "Unable to send error mail to home"
            }
        }
    }

    Exit
}

#This function write the error to the logfile and exit the script
function Write-Break([string]$Message)
{
    Write-Log -Message $Message
    Write-Error -Message $Message
    Stop-Script -WithError $true
}

Write-Log -Message "Script started ($($env:USERNAME))"
#endregion

Write-Log -Message "Check from $($TimeToCheck.ToString('yyyy-MM-dd HH:mm:ss'))"

$CountDeletions = (Get-ADObject -Filter { IsDeleted -eq $true -and WhenChanged -ge $TimeToCheck } -Properties ObjectClass -IncludeDeletedObjects -Credential (Get-SOECredential svc.uaa) | Where-Object { $_.ObjectClass -in $ClassesToCheck }).Count

Write-Log -Message "Deletion count: $CountDeletions"

if ( $CountDeletions -ge $DeletionThreshold )
{
    Write-Log -Message "The number of deletion in the last 3 hours: $CountDeletions"

    try
    {
        Send-MailMessage -SmtpServer $MailSMTP -From $MailSender -To $MailErrorTo -Subject "Deletion count larger then $DeletionThreshold" -Body "The number of deletion in the last 3 hours: $CountDeletions" -ErrorAction Stop
        Write-Log -Message "Sent warning mail to home"
    }
    catch
    {
        Write-Log -Message "Retry sending the message after 15 seconds"
        Start-Sleep -Seconds 15

        try
        {
            Send-MailMessage -SmtpServer $MailSMTP -From $MailSender -To $MailErrorTo -Subject "Deletion count larger then $DeletionThreshold" -Body "The number of deletion in the last 3 hours: $CountDeletions" -ErrorAction Stop
            Write-Log -Message "Sent warning mail to home after a retry"
        }
        catch
        {
            Write-Log -Message "Unable to send warning mail to home"
        }
    }
}
