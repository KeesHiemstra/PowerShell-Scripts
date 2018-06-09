<#
    ExportCalvi.ps1

    RfC Aldea 1190235: RITM0010543 AD export/ Company code: Global

    Export phone numbers.

    --- Version history
    Version 2.10 (2016-02-24, Kees Hiemstra)
    - Robbert van Straaten requested by mail to export all phone number and not only the ones that have a specific format.
    Version 2.00 (2016-01-29, Kees Hiemstra)
    - Taken over the ExportCalvi.vbs functionality.
#>

#Load settings
. "$PSScriptRoot\ExportCalvi.Config.ps1"

#Define log file variables
$ScriptName = $MyInvocation.MyCommand.Name.Replace(".ps1", '')
$MailSender = "$($ScriptName.Replace(' ', '.'))@JDECoffee.com"
$MailErrorTo = "Kees.Hiemstra@hpe.com"

#region LogFile
$Error.Clear()

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
    $LogMessage = "{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"),$Message
    Add-Content -Path $LogFile -Value $LogMessage
    Write-Debug $Message
}

#Finish script and report on $Error
function Stop-Script([bool]$WithError)
{
    if ($Error.Count -gt 0)
    {
        try
        {
            $Subject = "Error in $ScriptName on $($Env:ComputerName)"
            $MailBody = "The script $ScriptName has reported the following error(s):`n`n"
            $MailBody += $Error | Out-String

            Send-MailMessage -SmtpServer $MailSMTP -From $MailSender -To $MailErrorTo -Subject $Subject -Body $MailBody
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

<#
.SYNOPSIS
    Check if the layout of the phone is correct, return the number if it is okay or else return an empty string.
#>
Function Get-ValidPhoneNumber([string]$PhoneNumber)
{
    if ([string]::IsNullOrEmpty($PhoneNumber))
    {
        return ''
    }

    if ($PhoneNumber -match '\+\d{1,3}\ \d{4,15}')
    {
        return $PhoneNumber
    }
    else
    {
        return $PhoneNumber
    }
}

Write-Log -Message "Collect data from Active Directory"
$ADUsers = Get-ADUser -Filter { Mail -like '*' -and  (Mobile -like '*' -or OtherMobile -like '*' -or OfficePhone -like '*' -or OtherTelephone -like '*') } -Properties DepartmentNumber,
    ExtensionAttribute1,
    ExtensionAttribute3,
    Department,
    EmployeeID,
    GivenName,
    SN,
    Mail,
    EmployeeType,
    Title,
    Office,
    StreetAddress,
    PostalCode,
    L,
    Co,
    Mobile,
    OtherMobile,
    OfficePhone,
    OtherTelephone

Write-Log -Message "$($ADUsers.Count) records do have a mail address and one or more phone number are provided"

#Correct null-value fields
$ADUsers | Where { [string]::IsNullOrEmpty($_.ExtensionAttribute1) } |
    ForEach-Object { $_.ExtensionAttribute1 = '' }
$ADUsers | Where { [string]::IsNullOrEmpty($_.ExtensionAttribute3) } |
    ForEach-Object { $_.ExtensionAttribute3 = '' }

$Export = $ADUsers |
    Select-Object @{n='OpCo'; e={ $_.DepartmentNumber -join ', ' }},
        @{n='cost center'; e={ "$($_.ExtensionAttribute1.PadLeft(8, '0'))\$($_.ExtensionAttribute3.PadLeft(8, '0'))" }},
        department,
        @{n='employee number'; e={ $_.EmployeeID }},
        @{n='first name'; e={ $_.GivenName }},
        @{n='last name'; e={ $_.SN }},
        @{n='maiden name'; e={ '' }},
        @{n='initials'; e={ '' }},
        @{n='title (front)'; e={ '' }},
        @{n='middle name (belongs to last name)'; e={ '' }},
        @{n='title (back)'; e={ '' }},
        @{n='male/female'; e={ '' }},
        @{n='email address'; e={ $_.Mail }},
        @{n='contract type'; e={ $_.EmployeeType }},
        @{n='role'; e={ $_.Title }},
        @{n='site'; e={ $_.Office }},
        @{n='building'; e={ '' }},
        @{n='room'; e={ '' }},
        @{n='postal address'; e={ $_.StreetAddress }},
        @{n='ZIP'; e={ $_.PostalCode }},
        @{n='City'; e={ $_.L }},
        @{n='Country'; e={ $_.Co }},
        @{n='Visitors address'; e={ '' }},
        @{n='Visitors ZIP'; e={ '' }},
        @{n='Visitors City'; e={ '' }},
        @{n='Visitors Country'; e={ '' }},
        @{n='commencing date'; e={ '' }},
        @{n='termination date'; e={ '' }},
        @{n='username'; e={ $_.Mail -replace '\@.*$' }},
        @{n='reports to employee number'; e={ $_.ExtensionAttribute3 }},
        @{n='mobile phone'; e={ Get-ValidPhoneNumber($_.Mobile) }},
        @{n='2nd mobile phone'; e={ Get-ValidPhoneNumber($_.OtherMobile -join ',') }},
        @{n='business phone'; e={ Get-ValidPhoneNumber($_.OfficePhone) }},
        @{n='2nd business phone'; e={ Get-ValidPhoneNumber($_.OtherTelephone -join ',') }} |
    Where-Object { -not ([string]::IsNullOrEmpty($_.'mobile phone') -and [string]::IsNullOrEmpty($_.'2nd mobile phone') -and [string]::IsNullOrEmpty($_.'business phone') -and [string]::IsNullOrEmpty($_.'2nd business phone')) }

Write-Log -Message "$($Export.Count) records do have one or more valid formatted phone numbers"

$Export | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"

try
{
    Send-MailMessage -SmtpServer $MailSMTP -From $MailSender -To $MailExportTo -Subject $MailExportSubject -Body $MailExportBody -Attachments $ExportPath
    Write-Log -Message "Export sent to $MailExportTo"
}
catch
{
    Write-Break -Message "Unable to send export to $MailExportTo"
}

Stop-Script
