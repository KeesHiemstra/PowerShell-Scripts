<#
    CommonEnvironment.ps1

    Generic functions and variables. This need to be dot sourced in all distribution scripts.

    --- Version history
    Version 1.00 (2016-05-03, Kees Hiemstra)
    - Initial version.
#>

#region Set common dynamic variables
$SOEToolsPath = (Get-WmiObject -Class Win32_Share | Where-Object { $_.Name -eq 'SOETools$' }).Path
if ([string]::IsNullOrEmpty($SOEToolsPath))
{
    Write-Break -Message "There is no share called SOETools`$"
}
if (-not (Test-Path $SOEToolsPath))
{
    Write-Break -Message "The local path for the share SOETools`$ doesn't exist"
}

$SOEToolsPath += "\"
$StatusPath = "$($SOEToolsPath)Status.xml"

if ( -not (Test-Path -Path $StatusPath) )
{
    $Properties = [ordered]@{ComputerName=$env:COMPUTERNAME; 
                             CountryCode="Unknown";
                             Location="Unknown";
                             Description="Unknown";
                             IsInOfficeHours=$true;
                             StartStatus=(Get-Date);
                             LastStatus=(Get-Date);
                             Download=@();
                             }

    $Status = New-Object -TypeName PSObject -Property $Properties
    $Status | Export-Clixml $StatusPath
}
else
{
    $Status = Import-Clixml $StatusPath
}
#endregion

#region Update-Status
function Update-Status()
{
    $Status | Export-Clixml $StatusPath
}
#endregion

#region Get-CorrectedDate
try
{
    $TZInfo = [System.TimeZoneInfo]::FindSystemTimeZoneById($Status.TimeZone)
}
catch
{
    Write-Break -Message "Can't find the time zone '$($Status.TimeZone)'"
}

function Get-CorrectedDate($Date = (Get-Date))
{
    if ( $TZInfo -ne $null )
    {
        Write-Output ([System.TimeZoneInfo]::ConvertTimeFromUtc($Date.ToUniversalTime(), $TZInfo))
    }
    else
    {
        Write-Output (Get-Date)
    }
}
#endregion

#region Test-InOfficeHours
function Test-InOfficeHours
{
    if ( (Test-Path -Path "$($SOEToolsPath)InOfficeHours.txt") )
    {
        #Overwrite actual office hours settings
        #Used for testing and to bypass on special occasions
        Write-Output ( (Get-Content -Path "$($SOEToolsPath)InOfficeHours.txt") -eq "True" )
    }
    else
    {
        $Day = [int](Get-CorrectedDate).DayOfWeek #0=Sunday, 6=Saturday
        $Hour = (Get-CorrectedDate).Hour

        Write-Output ( $Day -ge 1 -and $Day -le 5 -and $Hour -ge 7 -and $Hour -le 20 )
    }
}
#endregion

#region LogFile
$ScriptName = $MyInvocation.MyCommand.Name.Replace(".ps1", '')
$LogFile = "$($PSScriptRoot)\SOEDistribution.log"
$LogStart = $Status.LastStatus.ToString('yyyy-MM-dd HH:mm:ss')
$SendCompleted = $false
$MailSMTP = "smtp.corp.demb.com"
$MailSender = "SOEDownloadScript@JDECoffee.com"
$MailErrorTo = "Kees.Hiemstra@hpe.com"

$Error.Clear()

#Create the log file if the file does not exits else write today's first empty line as batch separator
if (-not (Test-Path $LogFile))
{
    New-Item $LogFile -ItemType file | Out-Null
}

function Write-Log([string]$Message)
{
    $LogMessage = "{0} {1}" -f ((Get-CorrectedDate).ToString("yyyy-MM-dd HH:mm:ss")), $Message
    Add-Content -Path $LogFile -Value $LogMessage
    if ( [Environment]::UserInteractive ) { Write-Host $LogMessage }
}

function Send-SOEMessage([string]$Subject, [string]$Body, [string]$MailType)
{
    $MailBody = "CountryCode = $($Status.CountryCode)`n"
    $MailBody += "Location = $($Status.Location)`n"
    $MailBody += "Description = $($Status.Description)`n`n"
    $MailBody += "TimeZone = $($Status.TimeZone)`n"
    $MailBody += "Local time = $((Get-CorrectedDate).ToString('yyyy-MM-dd HH:mm:ss')) (Actual: $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')))`n"
    $MailBody += "`n"
    $MailBody += if ( (Test-InOfficeHours) ) { "In office hours`n" } else { "Outside office hours`n" }
    if ( ([array]$Status.Download).Count -gt 0 )
    {
        $MailBody += "Download(s) in progress according to Status.xml:`n"
        foreach ($Download in (Get-BitsTransfer))
        {
            $Download = 
            $MailBody += ("- $($Download.Description) at {0:P} ($($Download.JobState))`n" -f ($Download.BytesTransferred / $Download.BytesTotal))
        }
    }
    $MailBody += "`n`n"
    $MailBody += $Body

    try
    {
        Send-MailMessage -SmtpServer $MailSMTP -From $MailSender -To $MailErrorTo -Subject $Subject -Body $MailBody -ErrorAction Stop
        Write-Log -Message "Sent $MailType mail to home"
    }
    catch
    {
        Start-Sleep -Seconds 90
        try
        {
            Send-MailMessage -SmtpServer $MailSMTP -From $MailSender -To $MailErrorTo -Subject $Subject -Body $MailBody -ErrorAction Stop
            Write-Log -Message "Sent $MailType mail to home after retry"
        }
        catch
        {
            Write-Log -Message "Unable to send $MailType mail to home"
        }# second try
    }# first try
}

#Finish script and report on $Error
function Stop-Script([bool]$WithError, [switch]$Silent)
{
    if ($Error.Count -gt 0)
    {
        $Subject = "Error in $ScriptName on $($Env:ComputerName)"
        $MailBody = "The script $ScriptName has reported the following error(s):`n`n"
        $MailBody += $Error | Out-String

        $MailBody += "`n-------------------`n"
        $MailBody += (Get-Content $LogFile | Where-Object { $_.SubString(0, 19) -ge $LogStart }) -join "`n"

        Send-SOEMessage -Subject $Subject -Body $MailBody -MailType "error"
    }
    elseif ( $SendCompleted )
    {
        $Subject = "Finished $ScriptName on $($Env:ComputerName)"
        $MailBody += (Get-Content $LogFile | Where-Object { $_.SubString(0, 19) -ge $LogStart }) -join "`n"

        Send-SOEMessage -Subject $Subject -Body $MailBody -MailType "finish"

        $Status.LastStatus = (Get-CorrectedDate)
        Update-Status
    }

    if ($WithError)
    {
        Write-Log -Message "Script stopped with an error"
    }
    else
    {
        if (-not $Silent )
        {
            Write-Log -Message "Script ended normally"
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
#endregion

#region Check time zone
if ( $TZInfo -eq $null )
{
    Write-Break "Time zone is not correct"
}
#endregion

#region Check in or out office hours
$InOfficeHours = Test-InOfficeHours

if ( $Status.InOfficeHours -ne $InOfficeHours )
{
    $Status.InOfficeHours = $InOfficeHours
    Update-Status

    Write-Log -Message "Currently $( if ( $InOfficeHours ) {'in office hours'} else {'outside office hours'} )"
}
#endregion

#region Check listed downloads
if ( ([array]$Status.Download).Count -gt 0 )
{
    $DeletedJobs = ""
    foreach ($Download in $Status.Download)
    {
        $Bits = Get-BitsTransfer -JobId $Download.JobId -ErrorAction SilentlyContinue
        if ( $Bits -eq $null )
        {
            $DeletedJobs += "$($Download.JobId);"
            Write-Log -Message "$($Download.Description) transfer job is not existing any longer"
        }
    }
    $Status.Download = $Status.Download | Where-Object { $_.JobId -notin ($DeletedJobs -split ";") }
    Update-Status
}
#endregion

