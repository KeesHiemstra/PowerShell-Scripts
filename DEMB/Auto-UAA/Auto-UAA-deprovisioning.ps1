<#
Aldea 1016252: R86861 Update User account de-provisioning procedure

Version 2.30.008 (2016-05-12, Kees Hiemstra)
- Change license part at de-provisioning (adding an exclamation mark in from of EA2).
Version 2.20.007 (2016-02-01, Kees Hiemstra)
- Added completion of process if the account was already disabled.
Version 2.11.006 (2015-10-06, Kees Hiemstra)
- Bug fix, extensionAttribute14 can't be filled if manager is null.
Version 2.10.005 (2015-09-24, Kees Hiemstra)
- Added hide mailbox from OAB.
- Added error monitoring by mail.
Version 2.00.004 (2015-08-24, Kees Hiemstra)
- Copy manager attribute to extensionAttribute14 for RfC# 1153685 - R90466 change/remove generic accounts in deprovisioning script.
Version 1.02.003 (2014-11-05, Kees Hiemstra)
- Moved file needs to have a timestamp in the name: _yyyy-MM-dd-HHmm
Version 1.01.002 (2014-10-30, Kees Hiemstra)
- Set culture to en-US because for some reason the svc account is set to Dutch.
Version 1.00.001 (2014-10-29, Kees Hiemstra)
- Initial version.
#>

#Parameters
Param ([Switch]$Force)

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010  # E2k10

#Set the culture to get the name of the month right
[System.Threading.Thread]::CurrentThread.CurrentCulture = New-Object "System.Globalization.CultureInfo" "en-US"

#region Basic functions
$Error.Clear()

#Write the message to the log file
function Write-Log([string]$LogMessage)
{
    Add-Content -Path $LogFile -Value ("{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $LogMessage)
    Write-Debug $LogFile
}

#Write the error to the log file and exit the script
function Error-Break([string]$ErrorMessage)
{
    Write-Log($ErrorMessage)
    Write-Log("Script stopped")
    Stop-Script($true)
}

#Finish script and report on $Error
function Stop-Script([bool]$WithError)
{
    if ($Error.Count -gt 0)
    {
        $Subject = "Error in Auto-UAA-deprovisioning"
        $MessageBody = "The script Auto-UAA-deprovisioning has reported the following error(s):`n`n"
        $MessageBody += $Error | Out-String

        Send-MailMessage -SmtpServer $MailSMTP -From $MailSender -To $MailErrorTo -Subject $Subject -Body $MessageBody
    }

    if ($WithError)
    {
        Write-Log("Script stopped with an error")
    }
    else
    {
        Write-Log("Script ended normally")
    }
    Exit
}

#endregion

#region Script initializing and start

$LogFile = $MyInvocation.MyCommand.Definition -Replace(".ps1$",".log")

#Read settings from Auto-UAA.cfg
Import-CSV ($PSScriptRoot + "\Auto-UAA.cfg") -Delimiter ";" | ForEach-Object { New-Variable -Name $_.Variablename -Value $_.Value -Force }

#Read and overwrite settings from Auto-UAA-deprovisioning.cfg
Import-CSV ($PSScriptRoot + "\Auto-UAA-deprovisioning.cfg") -Delimiter ";" | ForEach-Object { New-Variable -Name $_.Variablename -Value $_.Value -Force }

If ($errorFilePath -notmatch "\\$")  { $errorFilePath += "\" }
If ($inputFilePath -notmatch "\\$")  { $inputFilePath += "\" }
If ($reportFilePath -notmatch "\\$") { $reportFilePath += "\" }
If ($succesFilePath -notmatch "\\$") { $succesFilePath += "\" }

$Today = get-date -Format "dd MMM yyyy"
$ToBeDeletedGroupName = "CN=To-Be-Deleted-$(Get-Date -Format "MMM"),OU=Disabled Accounts,OU=To be deleted,OU=Support,DC=corp,DC=demb,DC=com"

# Script start #

if((Test-Path $LogFile)) { Add-Content -Path $LogFile -Value "" }
Write-Log("Script started ({0})" -f ($env:USERNAME))

#Run only from scheduler, exit when started without -Force
if(!$force.IsPresent)
{
    Write-Error "Due to interaction with other processes this script must only run at the scheduled times. DO NEVER RUN THIS SCRIPT MANUALLY!"
    Error-Break ("Script manually started by: {0}" -f $env:USERNAME )
}

#endregion

#region Get the most recent file to process
[array]$inputFiles = (Get-ChildItem ("{0}\{1}*.csv" -f $inputFilePath, $inputFileFirstName) | Sort-Object LastWriteTime -Descending) # Get the most recent file
if ($inputFiles.Count -eq 0)
{ 
    Write-Log("No input file found")
    Stop-Script($false)
}
$inputFileName = $inputFiles[0].Name
$inputFile = $inputFilePath + $inputFileName
$succesFile = ("{0}{1}" -f $succesFilePath, $inputFileName)

Remove-Variable InputFiles -Force

#endregion

#region Read the data and perform basic checks

try
{
    [Array]$readUsers = Import-Csv -Delimiter ";" -Path $inputFile
    Write-Log ("{0} read as input file" -f $inputFile)
}
catch
{
    Error-Break("Error reading input file: {0}" -f $Error[0])
}

#Check if input file contains data
if ($readUsers.Count -eq 0) { Error-Break("Input file is empty") }

$SamAccountField = ($readUsers | Get-Member | Where { $_.Name -match "samAc\w*Name" }).Name
if ($SamAccountField -eq "") { Error-Break("Input file doesn't contain samAccountName field") }

#endregion

###############
# Main script #
###############

foreach ($User in $readUsers | Where-Object { $_.$SamAccountField -ne "" })
{
    $ADuser = Get-ADUser $User.$SamAccountField -Properties extensionAttribute2, extensionAttribute15, legacyExchangeDN, manager, memberOf -ErrorAction SilentlyContinue

    if ($ADuser -eq $null)
    {
        Write-Log("{0} does not exist" -f $User.$SamAccountField)
    }
    elseif ($ADuser.Enabled)
    {
        #Determain the nature of the mailbox (local or remote)
        $MailboxDN = $ADUser.legacyExchangeDN

        if ($MailboxDN -eq $null)
        {
            #NO MAILBOX
        }
        elseif ($MailboxDN -like "/o=DEMB/ou=External*" -or $MailboxDN -like "/o=DEMB/ou=Exchange Administrative Group*")
        {
            #REMOTE MAILBOX will be automatically disabled due to the removal of the license
            Set-RemoteMailbox -Identity $User.$SamAccountField -HiddenFromAddressListsEnabled $true
        }
        else
        {
            #try
            #{
                #LOCAL MAILBOX

                #Disable access via OWA (Outlook Web Access)
                Set-CASMailbox -Identity $ADuser.Name -OWAEnabled $false

                #Disable access via EAS (Exchange Active Sync)
                Set-CASMailbox -Identity $ADuser.Name -ActiveSyncEnabled $false

                #Disable the mailbox by setting all mailbox limits to 0 (zero)
                Set-Mailbox –Identity $ADuser.Name –MaxSendSize 0mb –MaxReceiveSize 0mb

                #Hide mailbox the Exchange Address Lists
                Set-Mailbox –Identity $ADuser.Name -HiddenFromAddressListsEnabled $true

                Write-Log("Mailbox for {0} has been disabled" -f $ADuser.sAMAccountName)
            #}
            #catch
            #{
            #    Write-Log("Unable to disable mailbox for {0}" -f $ADuser.sAMAccountName)
            #}
        }

        try
        {
            #Disable account
            Disable-ADAccount -Identity $ADUser
        
            #Add account to the To-Be-Deleted group
            if ($ToBeDeletedGroupName -notin $ADUser.memberOf)
            {
                Add-ADGroupMember -Identity $ToBeDeletedGroupName -Members $ADuser.sAMAccountName
            }
            
            if ($ADuser.extensionAttribute2 -eq $null)
            {
                $ADuser.extensionAttribute2 = ""
            }

            #Set ExtensionAttribute15 to "Left Company - DO NOT ENABLE – <DATE> - Auto deprovisioning EXA2=<value of ExtensionAttribute2>"
            Set-ADUser -Identity $ADUser.sAMAccountName -Replace @{ ExtensionAttribute15=("Left Company - DO NOT ENABLE – {0} - Auto deprovisioning EXA2={1}" -f $Today, $ADuser.extensionAttribute2) }

            #Handle licenses
            #Set-ADUser -Identity $ADUser.sAMAccountName -Clear "extensionAttribute2"
            if ( $ADuser.ExtensionAttribute2 -notlike "!*" -and -not [string]::IsNullOrEmpty($ADuser.ExtensionAttribute2) )
            {
                Set-ADUser -Identity $ADUser.sAMAccountName -Replace @{ExtensionAttribute2 = "!$($ADuser.ExtensionAttribute2)"}
            }

            if ($ADuser.manager -ne $null)
            {
                #Set ExtensionAttribute14 to the current manager
                Set-ADUser -Identity $ADUser.sAMAccountName -Replace @{ extensionAttribute14=$ADuser.manager }
            }

            #Set Manager to "CN=default.position,OU=Generic Accounts,OU=Support,DC=corp,DC=demb,DC=com"
            Set-ADUser -Identity $ADUser.sAMAccountName -Replace @{ manager="CN=default.position,OU=Generic Accounts,OU=Support,DC=corp,DC=demb,DC=com" }


            Write-Log("Account {0} has been disabled" -f $ADuser.sAMAccountName)
        }
        catch
        {
            Write-Log("Error during disabling of the account {0}" -f $ADuser.sAMAccountName)
        }

    }
	else
	{
		#User is already disabled, but maybe need to be added to the To-Be-Deleted group
		try
		{
            if ($ToBeDeletedGroupName -notin $ADUser.memberOf)
            {
                Add-ADGroupMember -Identity $ToBeDeletedGroupName -Members $ADuser.sAMAccountName
            }
            
            if ($ADuser.extensionAttribute2 -eq $null)
            {
                $ADuser.extensionAttribute2 = ""
            }

            #Set ExtensionAttribute15 to "Left Company - DO NOT ENABLE – <DATE> - Auto deprovisioning EXA2=<value of ExtensionAttribute2>"
            Set-ADUser -Identity $ADUser.sAMAccountName -Replace @{ extensionAttribute15=("Left Company - DO NOT ENABLE – {0} - Auto Deprovisioning EXA2={1}" -f $Today, $ADuser.extensionAttribute2) }

            #Handle licenses
            if ( $ADuser.ExtensionAttribute2 -notlike "!*" -and -not [string]::IsNullOrEmpty($ADuser.ExtensionAttribute2) )
            {
                Set-ADUser -Identity $ADUser.sAMAccountName -Replace @{ExtensionAttribute2 = "!$($ADuser.ExtensionAttribute2)"}
            }

            if ($ADuser.manager -ne $null)
            {
                #Set ExtensionAttribute14 to the current manager
                Set-ADUser -Identity $ADUser.sAMAccountName -Replace @{ extensionAttribute14=$ADuser.manager }
            }

            #Set Manager to "CN=default.position,OU=Generic Accounts,OU=Support,DC=corp,DC=demb,DC=com"
            Set-ADUser -Identity $ADUser.sAMAccountName -Replace @{ manager="CN=default.position,OU=Generic Accounts,OU=Support,DC=corp,DC=demb,DC=com" }


            Write-Log("Account {0} disabling has been completed" -f $ADuser.sAMAccountName)		
        }
        catch
        {
            Write-Log("Error during the completion of the disabling of the account {0}" -f $ADuser.sAMAccountName)
        }
	}

    $ADUser = $null
}

Move-Item -Path $inputFile -Destination ($succesFile) -Force
Write-Log("Input file has been moved")

# Script end #
Stop-Script($false)
