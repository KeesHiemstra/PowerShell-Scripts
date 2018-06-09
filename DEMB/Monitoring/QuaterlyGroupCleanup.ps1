<#
    QuaterlyGroupCleanup.ps1

    Remove every hour 200 members of the to-be-deleted groups from this quarter.

    === Version history
    Version 1.10 (2017-03-29, Kees Hiemstra)
    - Decrease threashold.
    Version 1.00 (2016-12-01, Kees Hiemstra)
    - Initial version.
#>
$MaxDeletion = 50
$DC = 'DEMBDCRS001'

#Define log file variables
$ScriptName = $MyInvocation.MyCommand.Name.Replace(".ps1", '')
$LogFile = $MyInvocation.MyCommand.Definition -replace(".ps1$",".log")
$LogStart = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
$MailSMTP = "smtp.corp.demb.com"
$MailSender = "$($ScriptName.Replace(' ', '.'))@JDECoffee.com"
$MailErrorTo = "Kees.Hiemstra@hpe.com"

#region LogFile
$Error.Clear()

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
#    Write-Host $Message
}

#Finish script and report on $Error
function Stop-Script([bool]$WithError)
{
    if ($Error.Count -gt 0)
    {
        $MailErrorSubject = "Error in $ScriptName on $($Env:ComputerName)"
        $MailErrorBody = "The script $ScriptName has reported the following error(s):`n`n"
        $MailErrorBody += $Error | Out-String

        $MailErrorBody += "`n-------------------`n"
        $MailErrorBody += (Get-Content $LogFile | Where-Object { $_.SubString(0, 19) -ge $LogStart }) -join "`n"

        try
        {
            Send-MailMessage -SmtpServer $MailSMTP -From $MailSender -To $MailErrorTo -Subject $MailErrorSubject -Body $MailErrorBody
            Write-Log -Message "Sent error mail to home"
        }
        catch
        {
            Write-Log -Message "Unable to send error mail to home"
        }
    }

    if ($WithError)
    {
        Write-Log -Message "Script stopped with an error"
    }
    else
    {
        Write-Log -Message "Script ended normally"
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

$Month = (Get-Date).Month

#Don't do anything if it is not the right month to delete the quarterly stuff
if ( $Month -notin @( 3, 6, 9, 12 ) ) { break }

#Looking for the Quarter to be deleted
                      $Quarter = 3   # in Dec delete Jul, Aug, Sep
if ( $Month -le 9 ) { $Quarter = 2 } # in Sep delete Apr, May, Jun
if ( $Month -le 6 ) { $Quarter = 1 } # in Jun delete Jan, Feb, Mar
if ( $Month -le 3 ) { $Quarter = 4 } # in Mar delete Oct, Nov, Dec

[array] $Groups = (Get-ADGroup -Filter * -SearchBase "OU=Q$Quarter,OU=Groups,OU=To be deleted,OU=Support,DC=corp,DC=demb,DC=com" -Properties Member -Server $DC) |
    Where-Object { $_.Member.Count -ge $MaxDeletion }

if ( $Groups.Count -eq 0 ) { Write-Log "Nothing to do"; Stop-Script }

Write-Log "Domain controller [$DC]"

foreach ( $Item in $Groups )
{
    Write-Log "[$($Item.Name)] contains $($Item.Member.Count) members"
}

Remove-ADGroupMember -Identity $Groups[0] -Members ($Groups[0].Member | Select-Object -First $MaxDeletion) -Confirm:$false -Server $DC

Stop-Script
