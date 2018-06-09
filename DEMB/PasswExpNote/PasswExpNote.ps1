<#
    Script PasswExpNote

    PowerShell version: 3.0

    RfC: 1167379 - R91077 Create Password Expiry Notification
    
    Send an email to the users that will have their password expired in 21 or 5 days.
    
    Contact: Nils Bauer    

History
Version 1.21 (2016-04-01, Kees Hiemstra)
- Discard the error message when the first attempt of sendin the mail fails.
Version 1.20 (2016-03-23, Kees Hiemstra)
- Improve the error message to home (include the log file in the message).
Version 1.11 (2016-03-18, Kees Hiemstra)
- Bug fix: Added -ErrorAction Stop to get Send-Message into the catch part.
Version 1.10 (2016-03-15, Kees Hiemstra)
- Bug fix: SMTP every now and then goes wrong because of time outs. Try to send the message again after 15 seconds.
Version 1.00 (2015-11-11, Kees Hiemstra)
- Initial version.
#>

$MessagePath = "D:\Scripts\PasswExpNote\Message.html"
$MessageSubject = "Action Required: Your JDE Windows password will expire soon"
$MailSMTP = "smtp.corp.demb.com"
$MailSender = "Password Expiriation<PasswordExpiration@JDECoffee.com>"
$MailErrorTo = "Kees.Hiemstra@hpe.com"

#Define log file with the name as the script
$LogFile = "D:\Scripts\PasswExpNote\PasswExpNote.log"

$HTMLMessage = Get-Content -Path $MessagePath -Raw

#region LogFile 1.20
$Error.Clear()

#Use: $Error.RemoveAt(0) to remove an error or warning captured in a try - catch statement

$ScriptName = $MyInvocation.MyCommand.Name.Replace(".ps1", '')
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

Write-Log("Script started ($($env:USERNAME))")
#endregion

#The mail will send once again after 15 seconds if the initial attempt failed.
#If that attempt fails again, the script will continue after 30 seconds with the next record.
function SendMessage(
    [string]$Mail,
    [string]$Dear,
    [string]$ExpireDays
    )
{
    try
    {
        Send-MailMessage -SmtpServer $MailSMTP -From $MailSender -To $Mail -Subject $MessageSubject -Body ($HTMLMessage -replace "\$\(Dear\)", $Dear -replace "\$\(ExpireDays\)", $ExpireDays) -BodyAsHtml -Priority High -ErrorAction Stop
        Write-Log "Mail sent to $Mail ($Dear)"
    }
    catch
    {
        $Error.RemoveAt(0)
        Write-Log "Mail sending to $Mail failed, retry in 15 seconds"
        Start-Sleep -Seconds 15

        try
        {
            Send-MailMessage -SmtpServer $MailSMTP -From $MailSender -To $Mail -Subject $MessageSubject -Body ($HTMLMessage -replace "\$\(Dear\)", $Dear -replace "\$\(ExpireDays\)", $ExpireDays) -BodyAsHtml -Priority High -ErrorAction Stop
            Write-Log "Mail sent to $Mail ($Dear)"
        }
        catch
        {
            $LogWarning = $true
            Write-Log "Mail sending to $Mail failed again, waiting 30 second to continue with the next one"
            Start-Sleep 30
        }
    }
}

$MaxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge
$SelectDate = (Get-Date).AddDays(21 - 91 + 1).Date

Write-Log "Selected date: $SelectDate"

$ADUsers = Get-ADUser -Filter {Enabled -eq $true -and Mail -like "*" -and PasswordLastSet -le $SelectDate} -Properties PasswordNeverExpires, PasswordExpired, PasswordLastSet, DisplayName, Mail |
    Where-Object { $_.PasswordNeverExpires -eq $false -and $_.PasswordExpired -eq $false }

foreach($ADUser in $ADUsers)
{
    Add-Member -InputObject $ADUser -NotePropertyName DaysBeforeExpiring -NotePropertyValue ($ADUser.PasswordLastSet.Date + $MaxPasswordAge - (Get-Date).Date).Days -Force
    Add-Member -InputObject $ADUser -NotePropertyName Dear -NotePropertyValue "" -Force

    if ($ADUser.GivenName -eq $null -or $ADUser.SurName -eq $null)
    {
        $ADUser.Dear = $ADUser.DisplayName
    }
    else
    {
        $ADUser.Dear = "$($ADUser.GivenName) $($ADUser.SurName)"
    }
}

Write-Log ("{0} users that need to change their password in 21 days" -f ($ADUsers | Where-Object { $_.DaysBeforeExpiring -eq 21 }).Count)

$ADUsers | Where-Object { $_.DaysBeforeExpiring -eq 21 } | 
    ForEach-Object { SendMessage -mail $_.mail -dear $_.Dear -expireDays 21 }

Write-Log ("{0} users that need to change their password in 5 days" -f ($ADUsers | Where-Object { $_.DaysBeforeExpiring -eq 5 }).Count)

$ADUsers | Where-Object { $_.DaysBeforeExpiring -eq 5 } | 
    ForEach-Object { SendMessage -mail $_.mail -dear $_.Dear -expireDays 5 }

Stop-Script -WithError $false
