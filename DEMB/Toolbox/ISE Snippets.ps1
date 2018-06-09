<#
    ISE Snippets

    --- Version history
    Version 1.01 (2016-01-31, Kees Hiemstra)
    - Small corrections in 'LogFile functions (complete)'
    Version 1.00 (2015-12-11, Kees Hiemstra)
    - Inital version.
#>

Remove-Item -Path ((Get-IseSnippet | Where-Object { $_.Name -eq 'Computer -SearchBase.snippets.ps1xml' }).VersionInfo.FileName) -ErrorAction SilentlyContinue

New-IseSnippet -Title 'Computer -SearchBase' -Description 'Add Win7 OU for Get-ADComputer -SearchBase' -Author 'Kees Hiemstra' -Force -Text @"
-SearchBase 'OU=Win7,OU=Managed Computers,DC=corp,DC=demb,DC=com'
"@

Remove-Item -Path ((Get-IseSnippet | Where-Object { $_.Name -eq 'DistinguishedName.snippets.ps1xml' }).VersionInfo.FileName) -ErrorAction SilentlyContinue

New-IseSnippet -Title 'DistinguishedName' -Description 'Help to spell the word DistinguishedName' -Author 'Kees Hiemstra' -Force -Text @"
DistinguishedName
"@

Remove-Item -Path ((Get-IseSnippet | Where-Object { $_.Name -eq 'LogFile functions (complete).snippets.ps1xml' }).VersionInfo.FileName) -ErrorAction SilentlyContinue

New-IseSnippet -Title 'LogFile functions (complete)' -Description 'Includes Write-Log, Write-Break, Stop-Script and Error reporting by mail' -Author 'Kees Hiemstra' -Force -Text @'
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

'@








