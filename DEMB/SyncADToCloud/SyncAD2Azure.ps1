<#
    Script SyncAD2Azure.ps1

    PowerShell version: 3.0

    RfC: 1108791 - R88904 O365 License Script Re-Write
    RfC: 1184259 - RITM0010140 User Provisioning - 'Ivaylo' scripts consolidation to central & O365 (CHG0010087)
    RfC: 1202852 - Dinamo project

    A user account needs to be member of the security groups EMS, Office and VPN in order to use these functionallity.
    Next to the access to Office, the user needs to have Office licenses to be able to use the Microsoft products.

    Active internal users will get all of the security groups above automatically by added this script. External users or special accounts
    need to have the access requested by the proper process (e.g. ServicePortal).

    Based on rules for active internal users or the Tags in the ExtensionAttribute2 from the AD user object, the security groups and Azure
    licenses are managed by this script.

    === Version history
    v3.30 (2017-03-02, Kees Hiemstra)
    - Disable automatic license setting for internal/active employees.
    v3.21 (2016-12-13, Kees Hiemstra)
    - Only exclude license setting without primary mail address if the license to set contains Exchange related tags ([EXC] or [Light])
    v3.20 (2016-10-24, Kees Hiemstra)
    - Implemented a workaround for the issue that a mailbox is not created because the Exchange could not contact the Global Catalog.
      In these cases the EA11 licenses would disabled by adding an exclamation mark in EA2; this now will only happen if the primary mail address exists.
    v3.13 (2016-09-02, Kees Hiemstra)
    - Log file now reports failed licenses not as an error if the licenses can’t be set due to a license shortage.
    - The script will send the log file as warning if an account doesn't get licenses if it doesn't have a primary mail address.
    - Bug fix: Licenses are not withdrawn if these are provided through ExtentionAttribute11 because the de-provisioning process is not allowed to update ExtensionAttribute11.
    - Send warning if licenses are pending (SharePoint Online Provisioning Service Issue).
    v3.12 (2016-08-22, Kees Hiemstra)
    - Avoid the log file from being send if the set license fails because there are not enough licenses left.
    v3.11 (2016-08-09, Kees Hiemstra)
    - Send log file about accounts not exist in Azure only after the account is more than 3 hours old.
    v3.10 (2016-08-03, Kees Hiemstra)
    - Don't set licenses when the account has no primary mail address (new requirement from Plamen).
    v3.00 (2016-07-06, Kees Hiemstra)
    - Using ExtensionAttribute11 as HR part of the license. ExtensionAttribut2 will only be used for extra licenses (RfC: 1202852 - Dinamo project).
    v2.02 (2016-05-12, Kees Hiemstra)
    - Bug fix: Report error when error occurs at 'License count reporting'.
    - Bug fix: Throw an error when no data can be read in 'Read Azure data'.
    - Bug fix: Throw an error when connection to Azure fails.
    v2.01 (2016-04-14, Kees Hiemstra)
    - Added a counter on failed license change due to the lack of Visio licenses.
    - Added a license count reporting.
    v2.00 (2016-02-28, Kees Hiemstra)
    - In production on 2016-04-12.
    - Implemented the requirements for RfC Aldea 1184259: RITM0010140 User Provisioning - 'Ivaylo' scripts consolidation to central & O365 (CHG0010087).
    - Renamed the script from SyncADToCloud to SyncAD2Azure.
    v1.30 2015-10-22 (Kees Hiemstra)
    - Groupmembership will be checked on Allow-EMS-JDE-employee or Allow-EMS-Non-JDE-employee. Membership will be added if the account isn't member yet.
      The request comes from Richard Wesseling. JDE has decided to add [EMS] to the OE3 profile. These groups are important for using the EMS service.
    v1.21 2015-10-19 (Kees Hiemstra)
    - Added logging when user is added to one of the groups.
    v1.20 2015-10-09 (Kees Hiemstra)
    - Groupmembership will be checked on Allow-O365-Access or Allow-External-O365-Access. Membership will be added if the account isn't member yet.
    v1.12 2015-09-25 (Kees Hiemstra)
    - Bug fix, deleting license on a user that doesn't exist in Azure.
    v1.11 2015-09-22 (Kees Hiemstra)
    - Bug fix, all external variable became arrays.
    v1.10 2015-09-21 (Kees Hiemstra)
    - Don't save the current user list when an error occured during setting the Azure license. At the next run the changes will be detirminded again
      but this time the change in Azure might go through.
    - Captured the setting in an external file.
    v1.00 2015-09-10 (Kees Hiemstra)
    - Initial version.

    --- Extra information
    -- Create database log table
    CREATE TABLE dbo.AutoUAALog(
	    [ID] int IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [Source] varchar(25) NOT NULL,
        [SamAccountName] varchar(25) NOT NULL,
	    [Action] varchar(50) NOT NULL,
        [Value] varchar(128) NULL,
        [ValueOld] varchar(128) NULL,
	    [Message] varchar(2048) NOT NULL,
	    [DTCreation] datetime NOT NULL CONSTRAINT [DF_AutoUAALog_DTCreation] DEFAULT (GETDATE())
	    )

#>

#region Load settings
. "$PSScriptRoot\SyncAD2Azure.Config.ps1"

#User to connect to the cloud
$UserName = "UserCreation@coffeeandtea.onmicrosoft.com"
#endregion

#region LogFile
$LogStart = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
$Error.Clear()
$LogWarning = $false
$Conn = New-Object -TypeName System.Data.SqlClient.SqlConnection

#$VerbosePreference = 'SilentlyContinue' #Verbose off
#$VerbosePreference = 'Continue' #Verbose on

#Create the log file if the file does not exits else write today's first empty line as batch separator
if ( -not (Test-Path $LogFile) )
{
    New-Item $LogFile -ItemType file | Out-Null
}
else 
{
    Add-Content -Path $LogFile -Value "---------- --------"
}

function Write-Log([string]$Message, [string]$SamAccountName, [ValidateSet("AddToGroup", "RemoveFromGroup", "SetEA2", "SetAzureLicense", "RemoveAzureLicense", "UpdateAzureLicense", "UpdateAzureLocation")][string]$Action, [string]$Value, [string]$ValueOld)
{
    $LogMessage = "{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Add-Content -Path $LogFile -Value $LogMessage
    if ($VerbosePreference -eq 'Continue') { Write-Host $LogMessage }
    Write-Host $LogMessage

    if ( -not [string]::IsNullOrEmpty($ConnectionString) -and -not [string]::IsNullOrEmpty($Action) )
    {
        if ( -not $RunNoDatabase )
        {
            $Cmd.Parameters.Clear()
            $Cmd.Parameters.Add('SamAccountName', [string]) | Out-Null
            $Cmd.Parameters.Add('Action', [string]) | Out-Null
            $Cmd.Parameters.Add('Value', [string]) | Out-Null
            $Cmd.Parameters.Add('ValueOld', [string]) | Out-Null
            $Cmd.Parameters.Add('Message', [string]) | Out-Null

            $Cmd.Parameters['SamAccountName'].Value = $SamAccountName
            $Cmd.Parameters['Action'].Value = $Action
            $Cmd.Parameters['Value'].Value = $Value
            $Cmd.Parameters['ValueOld'].Value = $ValueOld
            $Cmd.Parameters['Message'].Value = $Message

            $Cmd.ExecuteNonQuery() | Out-Null
        }
    }#Use database
}

#Finish script and report on $Error
function Stop-Script([bool]$WithError)
{
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
            if ( -not $RunOffline )
            {
                Send-MailMessage -SmtpServer $MailSMTP -From $MailSender -To $MailErrorTo -Subject $MailErrorSubject -Body $MailErrorBody -ErrorAction Stop
            }
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
                $MailErrorBody | Out-File -FilePath $ErrorFile
                Write-Log -Message "Unable to send error mail to home"
            }
        }
    }

    #Close the connection to the database
    if ( $Conn.State -notin ('broken', 'closed') )
    {
        try { $Conn.Close() } catch {}
    }

    if ( $WithError )
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

if ( -not [string]::IsNullOrEmpty($ConnectionString) )
{
    $Conn.ConnectionString = $ConnectionString
    try
    {
        if ( -not $RunNoDatabase )
        {
            $Conn.Open()
            $Cmd = New-Object -TypeName System.Data.SqlClient.SqlCommand
            $Cmd.Connection = $Conn
            $Cmd.CommandText = "INSERT INTO dbo.AutoUAALog([Source], [SamAccountName], [Action], [Value], [ValueOld], [Message]) VALUES('$ScriptName', @SamAccountName, @Action, @Value, @ValueOld, @Message)"
        }
    }
    catch
    {
        Write-Break -Message "Failed connection with connection string: $ConnectionString"
    }
}
#endregion

#region Load module
Get-Module "SOEAzure" | Remove-Module -Force
if ( -not $RunInDevelopment )
{
    Import-Module "$PSScriptRoot\SOEAzure" -DisableNameChecking
}
else
{
    # Get the module from the profile
    Get-Module SOEAzure | Remove-Module
    Import-Module "SOEAzure" -DisableNameChecking -Force
}

if ((Get-Module "SOEAzure") -eq $null)
{
    Write-Break("Module Azure is not able to load")
}
#endregion

#region Read Active Directory
try
{
    Write-Log -Message "Reading Active Directory"

    #$AllADUsers = Get-ADUser -Filter * -Properties $ADPropertyList | Where-Object { $_.MemberOf -contains 'CN=SyncAD2Azure-Pilot,OU=Groups,OU=SoE,DC=corp,DC=demb,DC=com' } |
    #$AllADUsers = Get-ADGroupMember "CN=SyncAD2Azure-Pilot,OU=Groups,OU=SoE,DC=corp,DC=demb,DC=com" | Get-ADUser -Properties $ADPropertyList |
    #$AllADUsers = Import-Clixml -Path C:\Transfer\AllADUsers.xml |# Where-Object { $_.SAMAccountName -like 'HPWin7User*' } |
    #$AllADUsers = Get-ADUser -Filter "Name -like 'a*'" -Properties $ADPropertyList |
    $AllADUsers = Get-ADUser -Filter * -Properties $ADPropertyList |
        Select-Object UserPrincipalName, 
            C,
            Comment,
            DisplayName,
            DistinguishedName,
            EmployeeID,
            EmployeeType,
            Enabled,
            ExtensionAttribute5,
            GivenName,
            Mail, 
            MemberOf,
            SAMAccountName,
            SurName,
            @{n='ADLocation'; e={ $_.DistinguishedName -replace '^CN=.*,OU=' }},
            @{n='IsInternalUser'; e={ $false }},
            @{n='IsActiveInternalUser'; e={ $false }},
            @{n='IsExternalUser'; e={ $_.employeeType -notin $EmployeeTypeInternalUser -and $_.extensionAttribute5 -notin $EA5Special }},
            @{n='IsSpecialAccount'; e={ $_.extensionAttribute5 -in $EA5Special }},
            @{n='IsToBeDeleted'; e={ -not $_.Enabled -and -not [string]::IsNullOrEmpty(($_.MemberOf -join ";") -like '*CN=To-Be-Deleted-*,OU=*') }},
            @{n='AllowEMS_I'; e={ $false } },
            @{n='AllowEMS_E'; e={ $false } },
            @{n='AllowOFF'; e={ $false } },
            @{n='AllowVPN'; e={ $false }},
            @{n='IsEA2Changed'; e={ $false }},
            @{n='UsageLocation'; e={ if ([string]::IsNullOrEmpty($_.c)) { 'CK' } else { $_.C } }},
            @{n='AzureTags'; e={ '' }},
            @{n='GroupTags'; e={ '' }},
            @{n='LicenseObject'; e={ Test-AzureLicenseTags -EA2 ($_.ExtensionAttribute2 -join '' | Sort-AzureLicenseTags) -EA11 ($_.ExtensionAttribute11 -join '') }}
            #Added the individual already existing AD attributes to be sure these do exist in memory as well

    Write-Log -Message "Finished reading Active Directory"
}
catch
{
    Write-Break -Message "Error reading Active Directory"
}

Write-Log -Message "Number of AD user accounts: $(([array]$AllADUsers).Count)"
Write-Log -Message "Number of enabled AD user accounts: $(([array]($AllADUsers | Where-Object { $_.Enabled })).Count)"
#endregion

#region Determine internal users
foreach ($ADUser in $AllADUsers |
            Where-Object { $_.ADLocation -in $ADLocationForInternals -and $_.EmployeeType -in $EmployeeTypeInternalUser -and $_.ExtensionAttribute5 -eq 'Employee' })
{
    $ADUser.IsInternalUser = $true
    $ADUser.IsActiveInternalUser = $ADUser.IsInternalUser -and $ADUser.Enabled -and $ADUser.Comment -eq 'Active'
}

Write-Log -Message "Number of internal users: $(([array]($AllADUsers | Where-Object { $_.IsInternalUser })).Count)"
Write-Log -Message "Number of active internal users: $(([array]($AllADUsers | Where-Object { $_.IsActiveInternalUser })).Count)"
#endregion

#region Set AD licenses in memory to (new) internal users and remove AD licenses to deprovisioned users
#Set LicenseTags for (new) active internal users without licenses and not member of the deny group
#foreach ($ADUser in $AllADUsers | Where-Object { $_.IsActiveInternalUser -and $_.MemberOf -notcontains $Groups.D_OFF -and ([string]::IsNullOrEmpty($_.LicenseObject.ResultTags) -or $_.LicenseObject.HasDenyLicenses) })
#{
#    $ADUser.LicenseObject = Test-AzureLicenseTags -EA2 $DefaultInternalOfficeTags -EA11 $ADUser.LicenseObject.EA11OriginalTags -OriginalEA2 $ADUser.LicenseObject.EA2OriginalTags
#    Write-Log -Message "Set default Azure licenses for [$($ADUser.SAMAccountName)] in memory"
#}

#Remove LicenseTags for disabled users or users that are member of the deny group
foreach ($ADUser in $AllADUsers | Where-Object { (-not $_.Enabled -or $_.MemberOf -contains $Groups.D_OFF) -and -not $_.LicenseObject.HasDenyLicenses })
{
    #Revoke all licenses
    if ( -not [string]::IsNullOrEmpty($ADUser.LicenseObject.EA2LicenseTags) -or -not [string]::IsNullOrEmpty($ADUser.LicenseObject.EA11LicenseTags) )
    {
        $ADUser.LicenseObject = Test-AzureLicenseTags -EA2 "!$($ADUser.LicenseObject.EA2OriginalTags)" -EA11 $ADUser.LicenseObject.EA11OriginalTags -OriginalEA2 $ADUser.LicenseObject.EA2OriginalTags
        Write-Log -Message "Removed Azure licenses for [$($ADUser.SamAccountName)] in memory"
    }
}
#endregion

#region Prepare for comparing AD licenses and actual Azure licenses by splitting group tags and Azure tags
foreach ($ADUser in $AllADUsers | Where-Object { -not [string]::IsNullOrEmpty($_.LicenseObject.ResultTags) -and -not $_.LicenseObject.HasDenyLicenses })
{
    $ADUser.AzureTags = $ADUser.LicenseObject.AzureLicenseTags
}#foreach AD user
#endregion

#region Corrections on country code that need to be in Azure
Write-Log "Number of accounts that needs correction on country code: $(([array]($AllADUsers | Where-Object { -not [string]::IsNullOrEmpty($_.LicenseObject.ResultTags) -and [string]::IsNullOrEmpty($_.c) -or $_.c -match 'UK|SW' })).Count)"
foreach ($ADUser in ($AllADUsers | Where-Object { -not [string]::IsNullOrEmpty($_.LicenseObject.AzureLicenseTags) -and [string]::IsNullOrEmpty($_.c) -or $_.c -match 'UK|SW' }))
{
    switch ($ADUser.UsageLocation)
    {
        'UK' { $ADUser.UsageLocation = 'GB' }
        'SW' { $ADUser.UsageLocation = 'SE' }
    }

    Write-Log -Message "Corrected country code for the account [$($ADUser.SamAccountName)] from [$($ADUser.C)] to [$($ADUser.UsageLocation)] in memory because [$($ADUser.LicenseObject.AzureLicenseTags)]"
}
#endregion

#region EMS
#EMS always for internal users
foreach($ADUser in $AllADUsers | Where-Object { $_.IsActiveInternalUser -and $_.MemberOf -notcontains $Groups.D_EMS -and $_.LicenseObject.GroupLicenseTags -match '\[EMS\]' -and -not $_.LicenseObject.HasDenyLicenses } )
{
    $ADUser.AllowEMS_I = $true
}

#EMS enternal users
foreach($ADUser in $AllADUsers | Where-Object { -not $_.IsActiveInternalUser -and $_.MemberOf -notcontains $Groups.D_EMS -and $_.LicenseObject.GroupLicenseTags -match '\[EMS\]' -and -not $_.LicenseObject.HasDenyLicenses } )
{
    $ADUser.AllowEMS_E = $true
}

Write-Log -Message "Number of internal EMS users: $(([array]($AllADUsers | Where-Object { $_.AllowEMS_I })).Count)"
Write-Log -Message "Number of external EMS users: $(([array]($AllADUsers | Where-Object { $_.AllowEMS_E })).Count)"

Write-Log -Message "To remove from internal EMS group: $(([array]($AllADUsers | Where-Object { -not $_.AllowEMS_I -and $_.MemberOf -contains $Groups.A_EMS })).Count)"
#Remove users from internal EMS group
foreach ($ADUser in $AllADUsers | Where-Object { -not $_.AllowEMS_I -and $_.MemberOf -contains $Groups.A_EMS })
{
    try
    {
        if ( $RunInProduction )
        {
            Remove-ADGroupMember -Identity $Groups.A_EMS -Members $ADUser.DistinguishedName -Confirm:$false
        }
        Write-Log -Message "Removed user [$($ADUser.SamAccountName)] from internal EMS group" -SamAccountName $ADUser.SamAccountName -Action RemoveFromGroup -Value $Groups.A_EMS
    }
    catch
    {
        Write-Log -Message "Error removing user [$($ADUser.SamAccountName)] from internal EMS group"
    }
}

Write-Log -Message "To add to internal EMS group: $(([array]($AllADUsers | Where-Object { $_.AllowEMS_I -and $_.MemberOf -notcontains $Groups.A_EMS })).Count)"
#Add users to internal EMS group
foreach ($ADUser in $AllADUsers | Where-Object { $_.AllowEMS_I -and $_.MemberOf -notcontains $Groups.A_EMS })
{
    try
    {
        if ( $RunInProduction )
        {
            Add-ADGroupMember -Identity $Groups.A_EMS -Members $ADUser.DistinguishedName
        }
        Write-Log -Message "Added user [$($ADUser.SamAccountName)] to internal EMS group" -SamAccountName $ADUser.SamAccountName -Action AddToGroup -Value $Groups.A_EMS
    }
    catch
    {
        Write-Log -Message "Error adding user [$($ADUser.SamAccountName)] to internal EMS group"
    }
}

Write-Log -Message "To remove from external EMS group: $(([array]($AllADUsers | Where-Object { -not $_.AllowEMS_E -and $_.MemberOf -contains $Groups.E_EMS })).Count)"
#Remove users from external EMS group
foreach ($ADUser in $AllADUsers | Where-Object { -not $_.AllowEMS_E -and $_.MemberOf -contains $Groups.E_EMS })
{
    try
    {
        if ( $RunInProduction )
        {
            Remove-ADGroupMember -Identity $Groups.E_EMS -Members $ADUser.DistinguishedName -Confirm:$false
        }
        Write-Log -Message "Removed user [$($ADUser.SamAccountName)] from external EMS group" -SamAccountName $ADUser.SamAccountName -Action RemoveFromGroup -Value $Groups.E_EMS
    }
    catch
    {
        Write-Log -Message "Error removing user [$($ADUser.SamAccountName)] from external EMS group"
    }
}

Write-Log -Message "To add to external EMS group: $(([array]($AllADUsers | Where-Object { $_.AllowEMS_E -and $_.MemberOf -notcontains $Groups.E_EMS })).Count)"
#Add users to external EMS group
foreach ($ADUser in $AllADUsers | Where-Object { $_.AllowEMS_E -and $_.MemberOf -notcontains $Groups.E_EMS })
{
    try
    {
        if ( $RunInProduction )
        {
            Add-ADGroupMember -Identity $Groups.E_EMS -Members $ADUser.DistinguishedName
        }
        Write-Log -Message "Added user [$($ADUser.SamAccountName)] to external EMS group" -SamAccountName $ADUser.SamAccountName -Action AddToGroup -Value $Groups.E_EMS
    }
    catch
    {
        Write-Log -Message "Error adding user [$($ADUser.SamAccountName)] to external EMS group"
    }
}
#endregion

#region Office
#Office for internal and requested external users
foreach($ADUser in $AllADUsers | Where-Object { $_.Enabled -and -not $_.LicenseObject.HasDenyLicenses -and $_.LicenseObject.GroupLicenseTags -match '\[ADFS\]' -and $_.MemberOf -notcontains $Groups.D_OFF } )
{
    $ADUser.AllowOFF = $true
}

Write-Log -Message "Number of Office users: $(([array]($AllADUsers | Where-Object { $_.AllowOFF })).Count)"

Write-Log -Message "To remove from Office group:  $(([array]($AllADUsers | Where-Object { -not $_.AllowOFF -and $_.MemberOf -contains $Groups.A_OFF })).Count)"

#Remove users from Office group
foreach ($ADUser in $AllADUsers | Where-Object { -not $_.AllowOFF -and $_.MemberOf -contains $Groups.A_OFF })
{
    try
    {
        if ( $RunInProduction )
        {
            Remove-ADGroupMember -Identity $Groups.A_OFF -Members $ADUser.DistinguishedName -Confirm:$false
        }
        Write-Log -Message "Removed user [$($ADUser.SamAccountName)] from allow Office group" -SamAccountName $ADUser.SamAccountName -Action RemoveFromGroup -Value $Groups.A_OFF
    }
    catch
    {
        Write-Log -Message "Error removing user [$($ADUser.SamAccountName)] from allow Office group"
    }
}

Write-Log -Message "To add to Office group:   $(([array]($AllADUsers | Where-Object { $_.AllowOFF -and $_.MemberOf -notcontains $Groups.A_OFF })).Count)"
#Add users to Office group
foreach ($ADUser in $AllADUsers | Where-Object { $_.AllowOFF -and $_.MemberOf -notcontains $Groups.A_OFF })
{
    try
    {
        if ( $RunInProduction )
        {
            Add-ADGroupMember -Identity $Groups.A_OFF -Members $ADUser.DistinguishedName
        }
        Write-Log -Message "Added user [$($ADUser.SamAccountName)] to allow Office group" -SamAccountName $ADUser.SamAccountName -Action AddToGroup -Value $Groups.A_OFF
    }
    catch
    {
        Write-Log -Message "Error adding user [$($ADUser.SamAccountName)] to allow Office group"
    }
}
#endregion

#region VPN
#VPN for internal and requested external users
foreach($ADUser in $AllADUsers | Where-Object { $_.Enabled -and ($_.IsActiveInternalUser -or $_.LicenseObject.GroupLicenseTags -match '\[VPN\]') -and $_.MemberOf -notcontains $Groups.D_VPN } )
{
    $ADUser.AllowVPN = $true
}

Write-Log -Message "Number of VPN users: $(([array]($AllADUsers | Where-Object { $_.AllowVPN })).Count)"

Write-Log -Message "To remove from VPN group:  $(([array]($AllADUsers | Where-Object { -not $_.AllowVPN -and $_.MemberOf -contains $Groups.A_VPN })).Count)"
#Remove users from VPN group
foreach ($ADUser in $AllADUsers | Where-Object { -not $_.AllowVPN -and $_.MemberOf -contains $Groups.A_VPN })
{
    try
    {
        if ( $RunInProduction )
        {
            Remove-ADGroupMember -Identity $Groups.A_VPN -Members $ADUser.DistinguishedName -Confirm:$false
        }
        Write-Log -Message "Removed user [$($ADUser.SamAccountName)] from allow VPN group" -SamAccountName $ADUser.SamAccountName -Action RemoveFromGroup -Value $Groups.A_VPN
    }
    catch
    {
        Write-Log -Message "Error removing user [$($ADUser.SamAccountName)] from allow VPN group"
    }
}

Write-Log -Message "To add to VPN group: $(([array]($AllADUsers | Where-Object { $_.AllowVPN -and $_.MemberOf -notcontains $Groups.A_VPN })).Count)"
#Add users to VPN group
foreach ($ADUser in $AllADUsers | Where-Object { $_.AllowVPN -and $_.MemberOf -notcontains $Groups.A_VPN })
{
    try
    {
        if ( $RunInProduction )
        {
            Add-ADGroupMember -Identity $Groups.A_VPN -Members $ADUser.DistinguishedName
        }
        Write-Log -Message "Added user [$($ADUser.SamAccountName)] to allow VPN group" -SamAccountName $ADUser.SamAccountName -Action AddToGroup -Value $Groups.A_VPN
    }
    catch
    {
        Write-Log -Message "Error adding user [$($ADUser.SamAccountName)] to allow VPN group"
    }
}
#endregion

#region Updating ExtensionAttribute2 in Active Directory
foreach ($ADUser in ($AllADUsers | Where-Object { $_.LicenseObject.HasEA2Changed -and $_.Mail -ne $null }) )
{
    try
    {
        if ( $RunInProduction )
        {
            
            if ( [string]::IsNullOrEmpty($ADUser.LicenseObject.EA2LicenseTags) )
            {
                Set-ADUser -Identity $ADUser.SamAccountName -Clear 'ExtensionAttribute2'
            }
            else
            {
                Set-ADUser -Identity $ADUser.SamAccountName -Replace @{ ExtensionAttribute2 = $ADUser.LicenseObject.EA2LicenseTags }
            }
        }
        Write-Log -Message "Set ExtensionAttribute2 for [$($ADUser.SamAccountName)] to [$($ADUser.LicenseObject.EA2LicenseTags)] from [$($ADUser.LicenseObject.EA2OriginalTags)]" -SamAccountName $ADUser.SamAccountName -Action SetEA2 -Value $ADUser.LicenseObject.EA2LicenseTags
    }
    catch
    {
        Write-Log -Message "Error updating ExtensionAttribute2 for [$($ADUser.SamAccountName)] to [$($ADUser.ExtensionAttribute2)]"
    }
}
#endregion

########################

#region Connect to Azure
Write-Log("Connect to the Azure cloud")
$UserNameFile = "$PSScriptRoot\$UserName.txt"

if ( -not $RunOffline )
{
    #Get stored credentials
    if($AzCred -eq $null)
    {
        if( -Not (Test-Path -Path $UserNameFile) )
        {
            $AzCred = Get-Credential -UserName $UserName -Message "Provide password"
            $AzCred.Password | ConvertFrom-SecureString | Out-File $UserNameFile
        }
        else
        {
            $AzCred = New-Object -Type System.Management.Automation.PSCredential -ArgumentList $UserName, (Get-Content -Path $UserNameFile | ConvertTo-SecureString)
        }
    }

    try
    {
        Connect-MsolService -Credential $AzCred -ErrorAction Stop
    }
    catch
    {
        Write-Break -Message "The connection to the cloud failed"
    }
}#-not RunOffline
#endregion

#region Read Azure data
try
{
    Write-Log -Message "Reading Azure"
    #$AllAzUsers = Import-Clixml -Path C:\Transfer\AllAzUsers.xml
    $AllAzUsers = Get-MsolUser -All -ErrorAction Stop |
        Select-Object UserPrincipalName, UsageLocation, 
            @{n='AzureTags'; e={ if ($_.Licenses.Count -gt 0) { (ConvertFrom-AzureLicense -Licenses $_.Licenses).GranularNotation } else { '' }}}

    Write-Log -Message "Finished reading Azure"
    Write-Log -Message "Number of Azure accounts: $($AllAzUsers.Count)"
}
catch
{
    Write-Break -Message "Error reading from Azure"
}
#endregion

#region Find differences between AD and Azure
if ( $AllAzUsers -eq $null )
{
    Write-Break -Message "No data read from Azure"
}

#Workout the differences
$Diffs = Compare-Object -ReferenceObject $AllADUsers -DifferenceObject $AllAzUsers -Property UserPrincipalName, UsageLocation, AzureTags -PassThru

Write-Log "Number of Azure actions to be investigated (licenses and usage locations): $(($Diffs | Where-Object { -not [string]::IsNullOrEmpty($_.UserPrincipalName) -and $_.SideIndicator -eq '<=' }).Count)"
#endregion

#region Update Azure
$Counts = [ordered]@{'NotExist'=0; 'NoMail'= 0; 'UpdateUsageLocation'=0;'SetLicense'=0;'UpdateLicense'=0;'RemoveLicense'=0; 'FailedLicense'=0}
$WhenCreatedCheck = (Get-Date).AddHours(-3) #An account need to be at least 3 hours old before the warning flag is set

#User exists in AD and a difference has been found between AD licenses and Azure licenses or the country code and usageLocation
foreach ($Diff in $Diffs | Where-Object { -not [string]::IsNullOrEmpty($_.UserPrincipalName) -and $_.SideIndicator -eq '<=' })
{
    $AzUser = $Diffs | Where-Object { $_.UserPrincipalName -eq $Diff.UserPrincipalName -and $_.SideIndicator -eq '=>' }

    if ( $AzUser -ne $null )
    {
        #Update usageLocation
        if ( $Diff.UsageLocation -ne $AzUser.UsageLocation )
        {
            try
            {
                if ( $RunInProduction )
                {
                    Set-MsolUser -UserPrincipalName $Diff.UserPrincipalName -UsageLocation $Diff.UsageLocation -ErrorAction Stop
                }

                $Counts.UpdateUsageLocation++
                Write-Log -Message "Updated UsageLocation for [$($Diff.SamAccountName)] from [$($AzUser.UsageLocation)] to [$($Diff.UsageLocation)]" -SamAccountName $Diff.SamAccountName -Action UpdateAzureLocation -Value $Diff.UsageLocation -ValueOld $AzUser.UsageLocation
            }
            catch
            {
                Write-Log -Message "Error updating UsageLocation for [$($Diff.SamAccountName)] from [$($AzUser.UsageLocation)] to [$($Diff.UsageLocation)]"
            }
        }#UsageLocation

        #Update license
        if ( $Diff.AzureTags -ne $AzUser.AzureTags )
        {
            if ( [string]::IsNullOrEmpty($Diff.Mail) -and $Diff.AzureTags -match '\[EXC\]|\[Light\]' )
            {
                $Counts.NoMail++
                $LogWarning = $true
                Write-Log -Message "Warning changing Azure license for [$($Diff.SamAccountName)]/[$($Diff.UserPrincipalName)] because the there is no primary mail address for setting [$($Diff.AzureTags)]"
            }
            else
            {
                $License = $Diff.AzureTags
                if ( [string]::IsNullOrEmpty($License) ) { $License = '!' }
      
                try
                {
                    if ( $RunInProduction )
                    {
                        Set-AzureLicense -UserPrincipalName $Diff.userPrincipalName -License $License | Out-Null
                    }

                    if ( [string]::IsNullOrEmpty($Diff.AzureTags) )
                    {
                        $Counts.RemoveLicense++
                        Write-Log -Message "Removed Azure license for [$($Diff.SamAccountName)] from [$($AzUser.AzureTags)]" -SamAccountName $Diff.SamAccountName -Action RemoveAzureLicense -Value '' -ValueOld $AzUser.AzureTags
                    }
                    elseif ( [string]::IsNullOrEmpty($AzUser.AzureTags) )
                    {
                        $Counts.SetLicense++
                        Write-Log -Message "Set Azure license for [$($Diff.SamAccountName)] to [$($Diff.AzureTags)]" -SamAccountName $Diff.SamAccountName -Action SetAzureLicense -Value $Diff.AzureTags -ValueOld $AzUser.AzureTags
                    }
                    else
                    {
                        $Counts.UpdateLicense++
                        Write-Log -Message "Updated Azure license for [$($Diff.SamAccountName)] from [$($AzUser.AzureTags)] to [$($Diff.AzureTags)]" $Diff.SamAccountName -Action UpdateAzureLicense -Value $Diff.AzureTags -ValueOld $AzUser.AzureTags
                        
                        if ( $Diff.AzureTags -like '*~*' )
                        {
                            $LogWarning = $true
                            Write-Log -Message "Warning on pending licenses for [$($Diff.SamAccountName)]"
                        }
                    }
                }
                catch
                {
                    $Counts.FailedLicense++

                    #Remove the error about not enough licenses from the error stack
                    if ( $Error[0].Exception -match 'Unable to assign this license because the  number of allowed licenses have been assigned.' )
                    {
                        Write-Log -Message "Can't change Azure license for [$($Diff.SamAccountName)] from [$($AzUser.AzureTags)] to [$($Diff.AzureTags)] because of license shortage"
                        $Error.RemoveAt(0)
                    }
                    else
                    {
                        Write-Log -Message "Error changing Azure license for [$($Diff.SamAccountName)] from [$($AzUser.AzureTags)] to [$($Diff.AzureTags)]"
                    }
                }
            }
        }#License set/update/remove
    }#Azure user found
    else
    {
        if ( -not [string]::IsNullOrEmpty($Diff.AzureTags) )
        {
            $Counts.NotExist++
            if ( $Diff.WhenCreated -gt $WhenCreatedCheck )
            {
                #An account need to be at least 3 hours old before the warning flag is set
                $LogWarning = $true
            }
            Write-Log -Message "User [$($Diff.SamAccountName)] does not exist in Azure and can't set the license to [$($Diff.AzureTags)]"
        }
    }#Azure user not found
}#foreach difference

#Report only if the count not equal to 0
if ( $Counts.UpdateUsageLocation ) { Write-Log -Message "Number of successful updating of UsageLocation: $($Counts.UpdateUsageLocation)" }
if ( $Counts.SetLicense )          { Write-Log -Message "Number of successful setting of Azure license: $($Counts.SetLicense)" }
if ( $Counts.UpdateLicense )       { Write-Log -Message "Number of successful updating of Azure license: $($Counts.UpdateLicense)" }
if ( $Counts.RemoveLicense )       { Write-Log -Message "Number of successful removing of Azure license: $($Counts.RemoveLicense)" }
if ( $Counts.FailedLicense )       { Write-Log -Message "Number of failed Azure license changes: $($Counts.FailedLicense)" }
if ( $Counts.NoMail )              { Write-Log -Message "Number of accounts without primary mail address: $($Counts.NoMail)" }
if ( $Counts.NotExist )            { Write-Log -Message "Number of accounts not found in Azure: $($Counts.NotExist)" }

#endregion

########################

#region License count reporting
if ( -not $RunOffline )
{
    try
    {
        $CurrSKU = Get-MsolAccountSku
        $PrevSKU = Import-Clixml -Path $SKUFile

        $LicTotalMsg = [string]::Empty
        $LicWarningMsg = [string]::Empty

        foreach ( $Item in $CurrSKU )
        {
            #Report on changes in the total number of licenses
            $Prev = $PrevSKU | Where-Object { $_.SkuPartNumber -eq $Item.SkuPartNumber }
            if ( $Item.ActiveUnits -lt $Prev.ActiveUnits )
            {
                $LicTotalMsg += "The total number of licenses for $($Item.SkuPartNumber) has gone down from $($Prev.ActiveUnits) to $($Item.ActiveUnits) active units.`n"
                Write-Log -Message "The total number of licenses for $($Item.SkuPartNumber) has gone down from $($Prev.ActiveUnits) to $($Item.ActiveUnits) active units"
            }
            elseif ( $Item.ActiveUnits -gt $Prev.ActiveUnits )
            {
                $LicTotalMsg += "The total number of licenses for $($Item.SkuPartNumber) has increased from $($Prev.ActiveUnits) to $($Item.ActiveUnits) active units.`n"
                Write-Log -Message "The total number of licenses for $($Item.SkuPartNumber) has increased from $($Prev.ActiveUnits) to $($Item.ActiveUnits) active units"
            }

            #Log the changes in the number of free licenses
            if ( $Item.ConsumedUnits -lt $Prev.ConsumedUnits )
            {
                Write-Log -Message "Used licenses of $($Item.SkuPartNumber) has gone down from $($Prev.ConsumedUnits) to $($Item.ConsumedUnits)"
            }
            elseif ( $Item.ConsumedUnits -gt $Prev.ConsumedUnits )
            {
                Write-Log -Message "Used licenses of $($Item.SkuPartNumber) increased from $($Prev.ConsumedUnits) to $($Item.ConsumedUnits)"
            }

            if ( $Item.ActiveUnits -ge 10000 ) { $LowLimit = 100 }
            elseif ( $Item.ActiveUnits -ge 2500 ) { $LowLimit = 50 }
            elseif ( $Item.ActiveUnits -ge 500 ) { $LowLimit = 25 }
            else { $LowLimit = 15 }
    
            $FreeLicenses = $Item.ActiveUnits - $Item.ConsumedUnits
            if ( $FreeLicenses -eq 0 )
            {
                #$LogWarning = $true
                $LicWarningMsg += "There are no free licenses for $($Item.SkuPartNumber) left.`n"
                Write-Log -Message "There are no free licenses for $($Item.SkuPartNumber) left"
            }
            elseif ( $FreeLicenses -le $LowLimit -and $Item.ConsumedUnits -ne $Prev.ConsumedUnits )
            {
                $LicWarningMsg += "Free licenses of $($Item.SkuPartNumber) are getting low, $FreeLicenses of $($Item.ActiveUnits) remaining.`n"
                Write-Log "Free licenses of $($Item.SkuPartNumber) are getting low, $FreeLicenses of $($Item.ActiveUnits) remaining"
            } 
        }

        #Report per mail
        if ( -not [string]::IsNullOrEmpty($LicTotalMsg) )
        {
            #Send message to report on changes in the total number of licenses
            try
            {
                Send-MailMessage -SmtpServer $MailSMTP -From $MailSender -To $MailErrorTo -Subject "Changes in the number of Azure licenses" -Body $LicTotalMsg -ErrorAction Stop
                Write-Log -Message "Sent licenses changed mail to home"
            }
            catch
            {
                Write-Log -Message "Retry sending the message after 15 seconds"
                Start-Sleep -Seconds 15

                try
                {
                    Send-MailMessage -SmtpServer $MailSMTP -From $MailSender -To $MailErrorTo -Subject "Changes in the number of Azure licenses" -Body $LicTotalMsg -ErrorAction Stop
                    Write-Log -Message "Sent licenses changed mail to home after a retry"
                }
                catch
                {
                    Write-Log -Message "Unable to send licenses changed mail to home"
                }
            }
        }

        if ( -not [string]::IsNullOrEmpty($LicWarningMsg) )
        {
            #Send message to warn on the number of free licenses
            try
            {
                Send-MailMessage -SmtpServer $MailSMTP -From $MailSender -To $MailErrorTo -Subject "Warning on the number of free Azure licenses" -Body $LicWarningMsg -ErrorAction Stop
                Write-Log -Message "Sent licenses warning mail to home"
            }
            catch
            {
                Write-Log -Message "Retry sending the message after 15 seconds"
                Start-Sleep -Seconds 15

                try
                {
                    Send-MailMessage -SmtpServer $MailSMTP -From $MailSender -To $MailErrorTo -Subject "Warning on the number of free Azure licenses" -Body $LicWarningMsg -ErrorAction Stop
                    Write-Log -Message "Sent licenses warning mail to home after a retry"
                }
                catch
                {
                    Write-Log -Message "Unable to send licenses warning mail to home"
                }
            }
        }

        $CurrSKU | Export-Clixml -Path $SKUFile
    }
    catch
    {
        Write-Break -Message "Not able to get data for 'License count reporting'"
    }
}#-not RunOffline
#endregion

Stop-Script
