<#
    ManageToBeOrphanedAccounts.ps1

    Check in the previous To-Be-Deleted group if there are accounts that are set as owner of Mailbox, Generic or Service account
    that will be orphaned once the user is deleted.

    This scripts is scheduled twice a month.
    - In the 1st run it will inform the manager of the to-be-deleted user about the accounts that are going to be orphaned.
    - In the 2nd run, when nothing has changed, the account will get the manager of the to-be-deleted user as new owner
        and the new owner will be notified.

History:
2015-12-09 v1.0 KHi Initial version (Aldea 1152638: R90447 Maintenance of Generic and Service Account Owners).
#>
param([ValidateSet("First", "Last")]$Mail)

#Load settings
. "$PSScriptRoot\ManageToBeOrphanedAccounts.Config.ps1"

$AccountTypesToCheck = ('generic', 'service', 'mailbox')
$Months = ('Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec')

#region LogFile

#Create the log file if the file does not exits else write today's first empty line as batch separator
if (!(Test-Path $LogFile))
    { New-Item $logFile -ItemType file | Out-Null }
else 
    { Add-Content -Path $LogFile -Value "---------- --------" }
#-------------------------------------------------------------
#log functions
#-------------------------------------------------------------
function Write-Log([string]$Message)
{
    $LogMessage = "{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"),$Message
    Add-Content -Path $LogFile -Value $LogMessage
    Write-Debug $Message
}

#This function write the error to the logfile and exit the script
function Error-Break([string]$Message)
{
    Write-Log($Message)
    Write-Log("Script stopped")
    Exit
}
Write-Log("Script started ($($env:USERNAME))")
#endregion

if ($Mail -eq $null)
{
    Error-Break -Message "The parameter Mail is not filled"
}

Write-Log("Accounts to check for: $($AccountTypesToCheck -join ", ")")

#Previous month, array starts at 0
$InvestigateMonth = (Get-Date).Month - 2
if ($InvestigateMonth -eq -1 ) { $InvestigateMonth = 11 } #=December

Write-Log("Examine To-be-deleted-$($Months[$InvestigateMonth])")

#Collect disabled users that are about the be deleted this month
$Users = Get-ADGroupMember -Identity "To-Be-Deleted-$($Months[$InvestigateMonth])" | 
    Get-ADUser -Properties sAMAccountName, displayName, manager, extensionAttribute3, extensionAttribute4, extensionAttribute14 |
    Where-Object { -not $_.Enabled }

Write-Log("Number of disabled users: $($Users.Count)")

$CountUserOwnerOf = 0
foreach ($User in $Users)
{
    $Manager = $null

    #Collect the accounts where the disabled user is set as manager (owner)
    $OwnerOf = Get-ADUser -Filter ("manager -eq '{0}'" -f $User.DistinguishedName) -Properties SAMAccountName, DisplayName, ExtensionAttribute5, ExtensionAttribute15 |
        Where-Object { $_.ExtensionAttribute5 -in $AccountTypesToCheck }

    if ($OwnerOf -ne $null)
    {
        Write-Log("User $($User.sAMAccountName) is owner of $($OwnerOf.sAMAccountName -join ", ")")
        $CountUserOwnerOf++

        if ($User.manager -ne $null -and $User.manager -ne 'CN=default.position,OU=Generic Accounts,OU=Support,DC=corp,DC=demb,DC=com')
        {
            #Manager of disabled user has not been cleared
            $Manager = Get-ADUser -Filter ("distinguishedName -eq '{0}'" -f $User.manager) -Properties mail, GivenName, SurName, DisplayName, DistinguishedName
        }

        if ($Manager -eq $null -and $User.extensionAttribute14 -ne $null -and $User.extensionAttribute14 -ne 'CN=default.position,OU=Generic Accounts,OU=Support,DC=corp,DC=demb,DC=com')
        {
            #Manager of disabled user has been saved in extensionAttribute14
            $Manager = Get-ADUser -Filter ("distinguishedName -eq '{0}'" -f $User.extensionAttribute14) -Properties mail, GivenName, SurName, DisplayName, DistinguishedName
        }

        if ($Manager -eq $null -and $User.extensionAttribute3 -ne $null -and $User.extensionAttribute3 -ne '00000000')
        {
            #Manager of disabled user can be found by employeeID in extensionAttribute3
            $Manager = Get-ADUser -Filter ("employeeID -eq '{0}'" -f $User.extensionAttribute3) -Properties mail, GivenName, SurName, DisplayName, DistinguishedName
        }

        if ($Manager -eq $null -and $User.extensionAttribute4 -ne $null)
        {
            #Manager of disabled user can be found by displayName in extensionAttribute4
            $Manager = Get-ADUser -Filter ("displayName -eq '{0}'" -f $User.extensionAttribute4) -Properties mail, GivenName, SurName, DisplayName, DistinguishedName
        }

        if ($Manager -eq $null)
        {
            #No manager found
            Write-Log("Disabled user $($User.sAMAccountName) with ownership has no reference to a manager")
        }
        else
        {
            #Manager found, notify him/her of the to be orphaned account(s)
            $HTML = '<html><head><meta http-equiv="Content-Type" content="text/html; charset=us-ascii"><title>Automatic e-mail, do not reply</title>'
            $HTML += '<style type="text/css"><!--body {font-size:11.0pt;font-family:"Calibri","Arial","Helvetica","sans-serif";}--></style>'
            $HTML += '</head><body><table border="0" style="padding: 2px 5px 2px 5px; margin: 3px; width:100%"><tr><td width="75px"></td><td>'
            $HTML += "<p>Dear Sir, Madam:</p>"
            switch ($Mail)
            {
                'First'
                {
                    $HTML += "<p>Recently the account $($User.SamAccountName) of $($User.GivenName) $($User.SurName) was disabled and scheduled for deletion.<br />"
                    $HTML += 'We would like to bring to your attention that this employee was the owner of the following accounts or mailboxes:</p>'
                }
                'Last'
                {
                    $HTML += "<p>In accordance with the previous notification regarding the disabling of the account $($User.SamAccountName) of $($User.GivenName) $($User.SurName)<br />"
                    $HTML += 'You have become the new owner of the following accounts or mailboxes:</p>'
                }
            }

            $HTML += '<table><tr><td style="font-weight:bold">Account name</td><td style="font-weight:bold">Display name</td><td style="font-weight:bold">Type</td><td style="font-weight:bold">Remark</td></tr>'
            foreach ($Owns in $OwnerOf)
            {
                $HTML += "<tr><td>$($Owns.SAMAccountName)</td><td>$($Owns.DisplayName)</td><td>$($Owns.extensionAttribute5)</td><td>$($Owns.extensionAttribute15)</td></tr>"
            }
            $HTML += '</table>'

            switch ($Mail)
            {
                'First'
                {
                    $HTML += '<p>To prevent these accounts from becoming an orphan we strongly advise you to update the &quot;Manager&quot; of these accounts.<br />'
                }
                'Last'
                {
                    $HTML += '<p>You can still submit a request to update the &quot;Manager&quot; of these accounts.<br />'
                }
            }
            $HTML += 'This can be done through the following catalog items as found in the IT Service Portal (accessible through the icon on your desktop)</p>'
            $HTML += '<ul><li>&quot;Change the owner of a Generic or Service Account&quot; (for accounts without mailbox)</li>'
            $HTML += '<li>&quot;Modify Generic or Resource Mailbox owner&quot; (for accounts with mailbox)</li></ul>'
            switch ($Mail)
            {
                'First'
                {
                    $HTML += '<p>If no action is taken, you will become the new owner effective 3 weeks after this email.</p>'
                }
            }
            $HTML += '<p>Regards.</p>'
            $HTML += '<p style="color:red">Please do reply to this email as it is automatically generated.</p></td><td width="75px"></td></tr></table></body></html>'

            try
            {
                if ($Mail -eq 'Last')
                {
                    #Change the manager of the effected accounts
                    foreach ($Owns in $OwnerOf)
                    {
                        Set-ADUser -Identity $Owns.sAMAccountName -Replace @{ manager = $Manager.DistinguishedName }
                        Write-Log -Message "Changed manager of $($Owns.samAccountName) to $($Manager.DistinguishedName)"
                    }
                }

                Send-MailMessage -SmtpServer $MailSMTP -From $MailSender -To $Manager.mail -Bcc "Kees.Hiemstra@JDEcoffee.com" -Subject "Account or mailbox owner change required" -Body $HTML -BodyAsHtml
                Write-Log -Message "Mail sent to $($Manager.mail)"
            }
            catch
            {
                Write-Log -Message "Can't send mail to $($Manager.samAccountName) ($($Manager.mail))"
            }
        }
    }#OwnerOf
}

Write-Log("Number of disabled users with an ownership: $CountUserOwnerOf")
