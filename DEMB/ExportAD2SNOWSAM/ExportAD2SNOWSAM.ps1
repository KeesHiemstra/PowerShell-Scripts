<#
    ExportAD2SNOWSAM.ps1

    RfC: 1274657 Export to SNOW SAM

    === Version history
    Version 1.00 (2017-04-12, Kees Hiemstra)
    - Inital version.
#>

$ExportPath = '\\CTCNLWIN037.corp.demb.com\UserUpdate\ExportAD2SNOWSAM.csv'

#region LogFile 1.20
$Error.Clear()

#Use: $Error.RemoveAt(0) to remove an error or warning captured in a try - catch statement

#Define initial log file variables
$ScriptName = $MyInvocation.MyCommand.Name.Replace(".ps1", '')
$MailSMTP = "smtp.corp.demb.com"
$MailSender = "$($ScriptName.Replace(' ', '.'))@JDECoffee.com"
$MailErrorTo = "Kees.Hiemstra@esfds.com"

$LogFile = 'D:\Scripts\ExportAD2SNOWSAM\ExportAD2SNOWSAM.log'
$LogStart = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")

#The error mail to home will be send if $LogWarning is true but no other errors were reported.
#The subject of the mail will indicate Warning instead of Error.
$LogWarning = $false

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

#region Export

Get-ADUser -Filter * -Properties Department, c, co, l, SAMAccountName, DisplayName, Mail, EmployeeID, DepartmentNumber, Office, Title |
    Select-Object @{n='User name'; e={ "DEMB\$($_.SAMAccountName)" }},
        @{n='Full name'; e={ $_.DisplayName }},
        @{n='Organisation'; e={ "ROOT,$($_.c),$($_.DepartmentNumber -join '|')" }},
        @{n='Location'; e={ $_.co }},
        @{n='SAMAccountName'; e={ $_.SAMAccountName }},
        @{n='Mail'; e={ $_.Mail }},
        EmployeeID,
        @{n='DepartmentNumber'; e={ $_.DepartmentNumber -join '|' }},
        @{n='Office'; e={ $_.Office }},
        @{n='Function/Role'; e= { $_.Title }} |
    Export-Csv -Path $ExportPath -Delimiter ';' -NoTypeInformation -Encoding Unicode -Force

#endregion

Stop-Script