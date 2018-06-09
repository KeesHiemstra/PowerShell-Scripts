<#
Initial version: arnoud.van.voorst@hp.com
Aldea 945952: R85379 AD account provisioning script
#>

<#
This script is part of a bundle of three scripts that run as scheduled jobs (Auto-UAA.ps1, Auto-UAA-messaging.ps1 and Auto-UAA-reprorting.ps1)
    Auto-UAA.ps1 is the script that read an input file with users and creates AD-accounts.
    Auto-UAA-messaging.ps1 is the script that creates mailboxes, enables Lync and set the password.
    Auto-UAA-reporting.ps1 generates a report of the accounts created yesterday

Txt files in the process and complete folders are use as data transfer between the scripts.
    <samaccountname>.mail.txt -> account waiting for mailbox creation
    <samaccountname>.lync.txt -> account waiting for enabling Lync
    <samaccountname>.enable.txt -> account waiting for password and enabling

Auto-UAA.cfg is a comma separated file that contains variables that are used in all scripts
#>

<# History
2017-04-14 v3.1.2 KHi  Bug fix: It was not recognized that the mailbox was already created since run.
2016-12-09 v3.1.1 KHi  Add more error messages at creating mailbox.
2016-10-24 v3.1.0 KHi  Implemented a workaround for the issue that a mailbox is not created because the Exchange could not contact the Global Catalog.
                       The mail step will be repeated in the next round in these cases.
2016-10-24 v3.0.0 KHi  RfC 1244719 - Wait of the existence of the manager before sending out the message to the manager.
2016-09-06 v2.9.6 KHi  Removed the extra text in the mail to the manager if the message is about an external user.
2016-09-02 v2.9.5 KHi  New exchange server.
2016-05-12 v2.9.4 KHi  Disabled adding user to VPN externals group.
2015-12-24 v2.9.3 KHi  Added 'submit request for Office license' text for external users in mail to manager.
2015-12-16 v2.9.2 PW   Updated list of CAS servers as provided by Richard.
2015-10-02 v2.9.1 KHi  Add UPN in log file at setting password for troubleshooting.
2015-09-24 v2.9.0 KHi  Unhide remote mailbox at re-enabling an account.
2015-07-01 v2.8.0 KHi  Enable the feedback for userPrincipalName (e-mail address will be @demb.com until June 30).
2015-06-11 v2.7.0 KHi  Stop temporarily the feedback for userPrincipalName (e-mail address will be @demb.com until June 30).
                       New link pools.
2015-04-21 v2.6.4 KHi  Use mail address as feedback for userPrincipalName.
2015-03-23 v2.6.3 KHi  Remove NL from the remote mailbox creation exception list.
2015-03-19 v2.6.2 KHi  Added Country -eq $null (to Country -eq '') for skipping the creation of the mailbox.
2015-03-16 v2.6.1 KHi  Remove BE from the remote mailbox creation exception list.
2015-01-21 v2.6   KHi  Enable remote mailboxes directly except for BE and NL and skip mailbox creation when Country -eq ''.
2014-12-24 v2.5.2 KHi  Mark manager unknown when it is set to Default.Position.
2014-12-17 v2.5.1 KHi  Improvent on logging.
2014-12-11 v2.5   KHi  Add re-enable user.
2014-12-12 v2.4.3 KHi  Add Accenture users to the Allow-VPN-Access-Externals security group.
2014-12-12 v2.4.2 KHi  Disable "User must change password at next log on" for users where company = "Accenture".
2014-11-24 v2.4.1 KHi  Bug fix. Unknown managemers where not recognized as such.
2014-11-06 v2.4   KHi  Added section to send the message to the manager (.enable => .report).
2014-11-04 v2.3   AVV  VDO users check changed to "^NA", original match "^88" - no special email to UAA team anymore.
2014-08-25 v2.2   KHi  Chaning the scope of the imported cfg file.
                       Changed the way the backslash is tested at the of the path variable.
2014-08-06 v2.1   AVV  Changed the logfile writing order (password/enabled) and fixed 'not enable' bug.
2014-08-05 v2.0   AVV  Added run only from scheduler test (-force).
                       Added username to startline logfile.
                  KHi  Don't enable user if EmployeeID starts with 88 (IBM / VDI).
2014-07-04 v1.0   AVV  initial script.
#>

#-------------------------------------------------------------
#parameters
#-------------------------------------------------------------

#run only from scheduler
Param ([Switch]$Force)

#set this only in test/bedugmode
#$DebugPreference = "continue"

#variables in cfg file (name = example)
#$inputFileName = "new_users.csv"
#$inputFilePath = "D:\Scripts\Auto-UAA"
#$succesFilePath = "D:\Scripts\Auto-UAA\Idm_new_users_Archive"
#$errorFilePath = "D:\Scripts\Auto-UAA\Idm_new_users_Error_Archive"
#$emailErrorsTo = "arnoud.van.voorst@hp.com"
#$emailErrorsFrom = "arnoud.van.voorst@hp.com"

##read (above) variables from cfg file
Import-CSV ($PSScriptRoot + "\Auto-UAA.cfg") -Delimiter ";" | ForEach-Object { New-Variable -Name $_.Variablename -Value $_.Value -Force -Scope Script }
If ($errorFilePath -notmatch "\\$")  { $errorFilePath += "\" }
If ($inputFilePath -notmatch "\\$")  { $inputFilePath += "\" }
If ($reportFilePath -notmatch "\\$") { $reportFilePath += "\" }
If ($succesFilePath -notmatch "\\$") { $succesFilePath += "\" }

$ProcessPath = $PSScriptRoot + "\process\"
$CompletedPath = $PSScriptRoot + "\completed\"

#create the process and completed folder if not exist already
If(!(test-path $ProcessPath)){New-Item $ProcessPath -ItemType directory >$null}
If(!(test-path $CompletedPath)){New-Item $CompletedPath -ItemType directory >$null}


#define logfile with the name as the script
$logFile = $MyInvocation.MyCommand.Definition -replace(".ps1$",".log")
#create the logfile if the file does not exitst else write todays first line
If(!(test-path $logFile)) { New-Item $logFile -ItemType file >$null }
Else { Add-Content -path $logFile -Value "" }
Add-Content -Path $logFile -Value $((get-date -Format "yyyy-MM-dd HH:mm:ss ") + "Script started")

### Functions ###

function New-Password([int]$Length){    $PW = $null    # exclude: o0O, 1lI etc.    $Forbidden = @('o','0','O',',',' ','1','l','I',"'",'"','`','?','^',';','/','\','~','|')    #define the number of charaters needed for each charactergroup (round up)    $c = [math]::ceiling(($Length / 4))    #numbers    $PW += (0..9 | where-object {$_ -notin $Forbidden} | Get-Random -Count $c)    #uppercase    $PW += (65..90 | %{[char]$_} | where-object {$_ -notin $Forbidden} | Get-Random -Count $c)    #lowercase    $PW += (97..122 | %{[char]$_} | where-object {$_ -notin $Forbidden} | Get-Random -Count $c)    #Special    $PW += ((32..47 + 58..64)  | %{[char]$_} | where-object {$_ -notin $Forbidden} |Get-Random -Count $c)    #Radomize the password order and limit to requested chars    $PW = ([system.string]::Join("",($PW | Get-Random -count $Length)))    RETURN $PW}

#-------------------------------------------------------------
#log functions
#-------------------------------------------------------------
function Write-Log([string]$LogMessage)
{
    $Message = "{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"),$LogMessage
    Add-Content -Path $logFile -Value $Message
    Write-Debug $Message
}

#This function write the error to the logfile and exit the script
function Error-Break([string]$ErrorMessage)
{
    Write-Log($ErrorMessage)
    Write-Log("Script stopped")
    Exit
}

#### MAIN script ###

#run only from scheduler, exit when started manually
if(!$force.IsPresent)
{
    Write-Error "Due to interaction with other processes this script must only run at the scheduled times. DO NEVER RUN THIS SCRIPT MANUALLY!"
    Error-Break ("Script manually started by: {0}" -f $env:USERNAME )
}


# Get DEMBDCSWA* domaincontroller availability (This will add the 'Available' property to the object)
# Availability is based on the existence of the host in the AD (then the DC is up and the AD services are running)
$Domaincontrollers = Get-ADComputer -SearchBase "OU=domain controllers,DC=corp,DC=demb,DC=com" -Filter {name -like "DEMBDCRS*"} |
                     Select-Object -Property *,@{n='Available';e={(Get-ADComputer -ldapfilter "(name=$($env:Computername))" -Server $_.name) -ne $null} }
$WorkDC = ($DomainControllers | Get-Random).name

### Prepare the exchange and lync environment
#$ExchangeServers = @("dembswacas090","dembswacas091","dembswacas092")
$ExchangeServers = @("DEMBRSMS419")
#remove existing sessions to the CAS servers
Get-PSSession | Where-Object {$exchangeServers -contains $_.ComputerName.Split(".")[0]} | Remove-PSSession
## Open Exchange Managementshell and pick one of the available CAS servers
$ExchangeServer = $ExchangeServers | Where-Object {Test-Connection -ComputerName $_ -Count 1 -Quiet} | Get-Random
$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ("http://{0}.corp.demb.com/PowerShell" -f $exchangeServer)
#import the module from CAS server
Import-PSSession $ExchangeSession | Out-Null
Import-Module Lync

<#
### Prepare Microsoft Online enviroment
Import-Module MSOnline
$Key = (2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43,6,6,6,6,6,6,31,33,60,23)
$MSOLUser = "UserCreation@coffeeandtea.onmicrosoft.com"
$MSOLPassword = ConvertTo-SecureString (Get-Content $PSScriptRoot\MSOLPassword.txt) -Key $Key
$MSOLCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $MSOLUser, $MSOLPassword 
Connect-MsolService -Credential $MSOLCred
#>

#region ###### Re-enable Mail ######
# Next step: Enable Lync if the mailbox still exists, else Create mailbox.
#
# Read the files/user schedule for processing and trim the mail.txt / lync.txt extention (usage of regEx)
$ScheduledForReMail = Get-ChildItem ($processPath + "\*.reenablemail.txt") | ForEach-Object {$_.name -Replace(".reenablemail.txt$","")}

Write-Debug ("Scheduled for mail: {0}" -f [string]$ScheduledForReMail)


foreach($User in $ScheduledForReMail)
{
    $ADUser = Get-ADUser $User -Properties legacyExchangeDN

    if ($ADUser.legacyExchangeDN -eq $null)
    {
        #No mailbox, needs to be created

        Rename-Item ("{0}\{1}.reenablemail.txt" -f $processPath, $User) -NewName ("{0}.mail.txt" -f $User)
    }#No mailbox
    elseif ($ADUser.legacyExchangeDN -like "/o=DEMB/ou=External*" -or "/o=DEMB/ou=Exchange Administrative Group*")
    {
        #Remote mailbox is still available, unhide
        Set-RemoteMailbox -Identity $User -HiddenFromAddressListsEnabled $false

        Rename-Item ("{0}\{1}.reenablemail.txt" -f $processPath, $User) -NewName ("{0}.lync.txt" -f $User)
        Write-Log ("Remote mailbox re-enabled: {0}" -f $User)
    }#Remote mailbox
    else
    {
        #Local mailbox
        try
        {
            Set-Mailbox –Identity $ADuser.Name –MaxSendSize Unlimited –MaxReceiveSize Unlimited -OWAEnabled $true -ActiveSyncEnabled $true -HiddenFromAddressListsEnabled $false

            Add-Content -path ("{0}\{1}.reenablemail.txt" -f $processPath, $User) -value ("{0} mailbox is re-enabled" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
            Rename-Item ("{0}\{1}.reenablemail.txt" -f $processPath,$user) -NewName ("{0}.lync.txt" -f $User)
            Write-Log ("Local mailbox re-enabled: {0}" -f $User)
        }
        catch
        {
            Add-Content -path ("{0}\{1}.reenablemail.txt" -f $processPath,$user) -value ("{0} not able to restore local mailbox" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
            Write-Log ("Error restoring local mailbox: {0}" -f $User)
        }
    }#Local mailbox

}#foreach user
#endregion

#region ###### Mail ######
# read the files/user schedule for processing and trim the mail.txt / lync.txt extention (usage of regEx)
$ScheduledForMail = Get-ChildItem ($processPath + "\*.mail.txt") | ForEach-Object {$_.name -Replace(".mail.txt$","")}

Write-Debug ("Scheduled for mail: {0}" -f [string]$ScheduledForMail)

foreach($User in $ScheduledForMail)
{
    #check if the account is already available on all DCs (if not available on one of the DCs the outcome is false)
    #if not avalible it will be picked up in the next time the script runs
    if( ($Domaincontrollers | ForEach-Object { $ADUser = (Get-ADUser -Filter {SamAccountName -eq $User} -Properties Country, Mail -Server $_.name); $ADUser -ne $null }) -notcontains $False )
    {
        if ( $ADUser.Mail -ne $null )
        {
            # Mailbox already created, rename the users file in the process folder
            Rename-Item ("{0}\{1}.mail.txt" -f $processPath, $User) -NewName ("{0}.lync.txt" -f $User)
            Write-Log -LogMessage "Mailbox for $User is already created"
        }# Mailbox was already created
        else
        {
            try
            {
                if ($ADUser.Country -ne '' -and $ADUser.Country -ne $null)
                {
                    if ($ADUser.Country -notin ('XZX'))
                    {
                        try
                        {
                            #create remote mailbox for this user
                            Enable-RemoteMailbox $User -RemoteRoutingAddress "$User@coffeeandtea.mail.onmicrosoft.com" -ErrorAction Stop

                            Start-Sleep -Seconds 15
                            $ADUser = Get-ADUser $User -Properties Country, legacyExchangeDN -Server $preferredExchangeDC -ErrorAction Stop

                            if ( $ADUser.legacyExchangeDN -ne $null )
                            {
                                #Mailbox is created
                                Enable-RemoteMailbox $User -Archive -ErrorAction Stop

                                Write-Log ("Remote mailbox created for: {0} ({1})" -f $User, $ADUser.Country)

                                #rename the users file in the proces folder
                                Rename-Item ("{0}\{1}.mail.txt" -f $processPath, $User) -NewName ("{0}.lync.txt" -f $User)
                            }
                        }
                        catch
                        {
                            Write-Log ("Error remote mailbox creation for: {0}" -f $Error[0])
                            Add-Content -Path ("{0}\{1}.mail.txt" -f $ProcessPath, $User) -Value ("{0} Error creating mailbox: {1}"  -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Error[0])
                        }
                    }#Remote mailbox
                    else
                    {
                        try
                        {
                            #create emailbox for this user
                            Enable-Mailbox -Identity $User -Alias $User -Database "DAG01DB23 (1GB)”

                            Write-Log ("Mailbox created for: {0} ({1})" -f $User, $ADUser.Country)

                            #For test phase only - set mailbox hidden from address lists
                            #If($DebugPreference = "continue"){Set-Mailbox –Identity $user -HiddenFromAddressListsEnabled $True}

                            #rename the users file in the proces folder
                            Rename-Item ("{0}\{1}.mail.txt" -f $processPath, $User) -NewName ("{0}.lync.txt" -f $User)
                        }
                        catch
                        {
                            Write-Log ("Error mailbox creation for: {0}" -f $Error[0])
                        }
                    }#Local mailbox
                }#Country provided
                else
                {
            	    Write-Log ("Mailbox NOT created for: {0} because country is empty or null" -f $User)
                    Send-MailMessage -SmtpServer $smtpServer -From "uaa@demb.com" -To $emailMonitoring -Subject "No country code provided" -Body "Mailbox for $User can't be requested because a country code was not provided."
            
                    #rename the users file in the proces folder, No mailbox = No Lync --> Enable account
                    Rename-Item ("{0}\{1}.mail.txt" -f $processPath, $User) -NewName ("{0}.enable.txt" -f $User)
                }
            }
            catch
            {
                Write-Log ("Genric error at mailbox creation for: {0}" -f $Error[0])
            }
        }#Mailbox was not yet created
    }
    Else
    {
        Add-Content -path ("{0}\{1}.mail.txt" -f $processPath, $user) -value ("{0} account {1} not replicated to all neccesary DCs yet"  -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $User)
    }
}
#endregion

#region ###### Lync ######

# read the files/user schedule for processing and trim the mail.txt / lync.txt extention (usage of regEx)
$ScheduledForLync = Get-ChildItem ($processPath + "\*.lync.txt") | ForEach-Object {$_.name -Replace(".lync.txt$","")}

Write-Debug ("Scheduled for lync: {0}" -f [string]$ScheduledForLync)

$Pools = @("cdffepool01.res.corp.demb.com", "lonfepool01.res.corp.demb.com")
foreach($User in $ScheduledForLync)
{
    #check if the mail property is already populated to all DCs in SWA and local where the script runs
    If((($Domaincontrollers | ForEach-Object { ((Get-ADUser -Filter {SamAccountName -eq $User} -Properties mail -Server $_.name)).mail -ne $null }) -notcontains $False) `
        -and ((get-aduser -Filter {SamAccountName -eq $User} -Properties mail -Server (Get-ADDomainController).name ).mail -ne $Null))
    {
        #$TempError = $Error
        Try
        {
            #enable Lync for this user
            #Enable-CsUser -Identity ("demb\{0}" -f $User) -RegistrarPool "eepool.corp.DEMB.com" -SipAddressType EmailAddress -SipDomain demb.com
            Enable-CsUser -Identity ("demb\{0}" -f $User) -RegistrarPool $Pools[(Get-Random -Minimum 0 -Maximum $Pools.Count)] -SipAddressType EmailAddress -SipDomain JDEcoffee.com
            #if($Error -ne $TempError)
            #{ 
            #    Send-MailMessage -Body "Lync for $User can't be enable because of the following error: $($Error[0])" -Subject "Enable Lync error" -SmtpServer $smtpServer -From "uaa@demb.com" -To $emailMonitoring
            #                
            #    Write-Log ("Error enabling Lync for: {0}" -f $Error[0])
            #    continue
            #}

            Write-Log ("Lync enabled for: {0}" -f $User)

            #rename the users file in the proces folder
            Rename-Item ("{0}\{1}.lync.txt" -f $processPath, $user) -NewName ("{0}.enable.txt" -f $User)

            Write-Debug ("Lync enabled for {0}" -f $User)
        }
        catch
        {
            Add-Content -path ("{0}\{1}.lync.txt" -f $processPath, $User) -value ("{0} Error enabeling Lync for: {1}"  -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Error[0])
            Write-Log ("Error enabeling Lync for: {0}" -f $Error[0])
            Send-MailMessage -Body "Lync for $User can't be enable because of the following error: $($Error[0])" -Subject "Enable Lync error" -SmtpServer $smtpServer -From "uaa@demb.com" -To $emailMonitoring
        }
     }
     Else
     {
            Write-Debug "Mail for $user not synced yet"
            Add-Content -path ("{0}\{1}.lync.txt" -f $processPath,$user) -value ("{0} mail attribute for {1} not replicated to all neccesary DCs yet"  -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"),$user)
     }
}
#endregion

#region ###### Enable Account and set password ######

# read the files/user schedule for processing and trim the mail.txt / lync.txt extention (usage of regEx)
$ScheduledToEnable = Get-ChildItem ($processPath + "\*.enable.txt") | ForEach-Object {$_.name -Replace(".enable.txt$","")}

Write-Debug ("Scheduled to enable: {0}" -f [string]$ScheduledToEnable)

foreach($User in $ScheduledToEnable)
{
    try
    {
        $ADUser = Get-ADUser $User -Properties userPrincipalName, company, employeeID, mail

        #Check if the primary mail address can be set as userPrincipalName
        If(-not [string]::IsNullOrWhiteSpace($ADUser.mail) -and $ADUser.mail -ne $ADUser.userPrincipalName)
        {
            if((Get-ADUser -Filter {userPrincipalName -eq $ADUser.mail}) -eq $null)
            {
                #The primary mail address is not yet used as userPrincipalName
                Set-ADUser -Identity $User -UserPrincipalName ($ADUser.mail).ToLower() -Server $WorkDC

                Write-Log ("$User has the userPrincipalName changed to the mail address")
                Send-MailMessage -SmtpServer $smtpServer -From "uaa@demb.com" -To $emailMonitoring -Subject "UPN has changed" -Body "$User has the userPrincipalName ($($ADUser.userPrincipalName)) changed to the mail address ($($ADUser.mail.ToLower()))."
                $ADUser = Get-ADUser $User -Properties userPrincipalName, company, employeeID, mail
            }
            else
            {
                #Primary mail address is already used as userPrincipal name and can't be used
                Write-Log ("$User can't have the userPrincipalName changed to the mail address $($ADUser.mail)")
                Send-MailMessage -SmtpServer $smtpServer -From "uaa@demb.com" -To $emailMonitoring -Subject "UPN can't be changed" -Body "$User can't have the userPrincipalName ($($ADUser.userPrincipalName)) changed to the mail address ($($ADUser.mail.ToLower()))."
            }
        }

        #set password and must change at next logon
        $Password = New-Password(10)
        Set-ADAccountPassword -Identity $User -NewPassword (ConvertTo-SecureString -AsPlainText $Password -Force) -Server $WorkDC

        if ($ADUser.company -match '^Accenture$')
        {
            #2014-12-12 - KHi: Accenture users only connect through VPN and they can't logon if ChangePasswordAtLogon -eq $true
            Set-ADUser -Identity $User -ChangePasswordAtLogon $false -Server $WorkDC
            #Add-ADGroupMember -Identity "CN=Allow-VPN-Access-Externals,OU=Groups,DC=corp,DC=demb,DC=com" -Members $User -Server $WorkDC

            Write-Log ("Account {0} is set as Accenture user", $User)
        }
        else
        {
            Set-ADUser -Identity $User -ChangePasswordAtLogon $True -Server $WorkDC
        }

        Add-Content -path ("{0}\{1}.enable.txt" -f $ProcessPath, $User) -Value ("{0} Password: {1}"-f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Password)
        Write-Log ("Password set for {0} (UPN:{1})" -f $User, $ADUser.userPrincipalName)

        #Enable the account (not for IBM / VDI users)
        if ($ADUser.employeeID -notmatch "^NA") # changed to "^NA", original match "^88" - IBM/VDI users not "ON HOLD anymore"
        {
            Enable-ADAccount -Identity $User -Server $WorkDC
            Add-Content -path ("{0}\{1}.enable.txt" -f $ProcessPath, $User) -Value ("{0} Account enabled"-f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
            Write-Log ("Account {0} enabled" -f $User)
        }
        Else
        {
            Add-Content -path ("{0}\{1}.enable.txt" -f $ProcessPath, $User) -Value ("{0} Account not enabled"-f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
            Write-Log ("Account {0} not enabled" -f $User)
        }

        #Split the workflow as requested in RfC 1244719 - manager.txt will contain the information for the mail, report.txt will trigger the export to IDM
                            
        #Rename the users file in the process folder
        Move-Item -Path ("{0}\{1}.enable.txt" -f $ProcessPath, $User) -Destination ("{0}\{1}.manager.txt" -f $processPath, $User) -Force

        #Create a report file (RfC 1244719)
        New-Item -Path ("{0}\{1}.report.txt" -f $ProcessPath, $User) -ItemType File -Force | Out-Null
    }
    catch
    {
        Write-Log ("Error set password and/or enabeling account for: {0}" -f $Error[0])
    }
     
}
#endregion

#region ###### Send a message to the manager of the new employee ######

# read the files/user schedule for processing and trim the manager.txt extention (usage of regEx)
$ScheduledToManager = Get-ChildItem ($processPath + "\*.manager.txt") | ForEach-Object {$_.name -Replace(".manager.txt$","")}

Write-Debug ("Scheduled to send massage to the manager: {0}" -f [string]$ScheduledToManager)

foreach($User in $ScheduledToManager)
{
    try
    {
        #Get the AD data from the user
        $ADUser = Get-ADUser $User -Properties *

        #Only send a mail when the manager is known (RfC 1244719)
        if ( -not ($ADUser.Manager -eq $null -or $ADUser.Manager -like 'CN=default.position*') )
        {
            #Get the password from the <Processpath>\<samaccountname>.manager.txt
            #Read password form last line in the user completed file (regex used to check and strip the datepart)
            $Password = $null
            $Match = "^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} Password: "            $Password = (Get-Content ("{0}{1}.manager.txt" -f $processPath, $ADUser.sAMAccountname) | Where-Object {$_ -match $Match} | Select-Object -Last 1) -replace($Match)            Add-Member -Force -InputObject $ADUser -NotePropertyname Password -NotePropertyValue $Password

            Write-Debug (Out-String -InputObject $ADUser)

            #Set manager properties to unknow and try to find these properties afterwards
            Add-Member -Force -InputObject $ADUser -NotePropertyname ManagerEmail -NotePropertyValue "<unknown>"
            Add-Member -Force -InputObject $ADUser -NotePropertyname ManagerDisplayname -NotePropertyValue "<unknown>"

            #Get email address of manager
            If(($ADUser.manager -ne "") -and ($ADUser.manager -ne $null) -and ($ADUser.manager -ne "CN=default.position,OU=Generic Accounts,OU=Support,DC=corp,DC=demb,DC=com"))
            {
                $ADUser.ManagerEmail = (Get-ADUser -LDAPFilter "(DistinguishedName=$($ADUser.manager))" -Properties mail).mail
                $ADUser.ManagerDisplayname = (Get-ADUser -LDAPFilter "(DistinguishedName=$($ADUser.manager))" -Properties displayName).displayName
            }

            $body = ""
            #select the properties for the email and add them to the body ('samaccountname' is translated to 'loginname')
            #and make a capital from the first character of the property name (not all AD properties start with a capital)
            ($ADUser | Select-Object @{n="LoginName"; e="sAMAccountName"}, Password, Mail, EmployeeID, @{n="manager"; e="ManagerDisplayName"},
                ManagerEmail).PSObject.Properties |
                    ForEach-Object { $body += ("{0} = {1}`r`n" -f $_.Name, $_.Value) }

            #Add extra remark on external users
            #if ($ADUser.employeeType -eq 'External employee')
            #{
            #    $Body += "`n`n`nFor an external employee you need to submit a separate request in the IT Service Portal (accessible through the icon on your desktop) for an Office license if the external employee needs to have access to the newly created mailbox and Intranet.`nRelated catalog item is 'Request Intranet Access (Request Intranet (Office 365) Access for External/Temporary Employees.)'"
            #}

            #generate the standard subject
            $emailSubject = ("Account created for {0} on {1}" -f ($ADUser.sAMAccountName),(Get-Date($ADUser.createTimeStamp) -Format "yyyy-MM-dd"))

            #standard powershell send-mailmessage doesn't support replyto:
            $msg = new-object Net.Mail.MailMessage
            $smtp = new-object Net.Mail.SmtpClient($smtpServer)

            $msg.From = $emailNewAccountsFrom
            $msg.ReplyTo = $emailNewAccountsReplyTo


            #if manager email is known send to manager, otherwise the email is send to the default to address from the cfg file (emailNewAccountsTo/UAA team)
            if ($ADUser.employeeID -match "^NA") # changed to "^NA", original match "^88" - IBM/VDI users not "ON HOLD anymore"
            {
                # VDI user (EmployeeID starts with 88)
                $msg.To.Add($emailNewAccountsTo)
                $msg.Subject = "ON HOLD VDI: " + $emailSubject
            }
            elseif($ADUser.ManagerEmail -eq "<unknown>")
            {
                # unknown manager (no VDI)
                $msg.To.Add($emailNewAccountsTo)
                $msg.Subject = "UNKNOWN MANAGER: " + $emailSubject
            }
            else
            {
                # known manager (no VDI)
                $msg.To.Add($ADUser.ManagerEmail)
                $msg.Subject = $emailSubject
                # UAA team on BCC
                $msg.Bcc.Add($emailNewAccountsBCC)
            }

            #for debug add developer to BCC
            #$msg.Bcc.Add($emailDeveloper)
                
            #Add body and send email
            $msg.body = $body
            $smtp.Send($msg)

                
            #Write email send to log
            Add-Content -Path ("{0}\{1}.manager.txt" -f $ProcessPath, ($ADUser.sAMAccountname)) -value ("{0} Email sent to: {1}"  -f ((Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg.To.ToString() ))
            Write-Log ("Email sent to {0} for account {1}" -f $msg.To.ToString(), $User)
                            
            #Move the users file in the completed folder (RfC 1244719)
            Move-Item -Path ("{0}\{1}.manager.txt" -f $ProcessPath, ($ADUser.sAMAccountname)) -Destination ("{0}\{1}.report.txt" -f $CompletedPath, ($ADUser.sAMAccountname)) -Force
        }
    }
    catch
    {
        Write-Log ("Error sending mail to the manager for: {0}" -f $Error[0])
    }
     
}#foreach $User
#endregion

## Exit script normally
Write-Log("Script ended normally")
