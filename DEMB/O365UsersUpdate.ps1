#######################################################
#
# Script to update O365 Users for O365 and VPN Access
# 
# Author: Ivaylo Petkov - ivaylop@hpe.com
# Version: 1.0
#
#######################################################

import-module activedirectory

# Get Current Date
$date=get-date -format s
$date=$date -replace(":","-")

#Create Report File For Requirement1 
$addfileint=new-item -path "D:\Scripts\O365Users\Reports" -name ("InternalUsersAdded" + $date +".csv") -itemtype file
add-content $addfileint "Samaccountname, Co, Comment, EmployeeID, Employeetype, Extensionattribute5, Extensionattribute2, Enabled"

#Put The Groups That Will Be Updated In Variables
$group = "CN=Allow-VPN-Access,OU=Groups,DC=corp,DC=demb,DC=com"
$group1= "CN=Allow-O365-Access,OU=Groups,DC=corp,DC=demb,DC=com"
$group2= "CN=Deny-VPN-Access,OU=Groups,DC=corp,DC=demb,DC=com"
$group3= "CN=Deny-O365-Access,OU=Groups,DC=corp,DC=demb,DC=com"
$group4= "CN=Allow-External-O365-Access,OU=Groups,DC=corp,DC=demb,DC=com"
$group5= "CN=Allow-VPN-Access-Externals,OU=Groups,DC=corp,DC=demb,DC=com"      

#Filter For User Accounts Per Requirement1
$users=get-aduser -ldapfilter "(&(objectCategory=user)(objectClass=user)(comment=active)(extensionattribute5=employee)(|(employeetype=active\20employee)
(employeetype=Expats/ Inpats)(employeetype=Retiree/ Pensioner)))" -properties co, comment, employeeid, employeetype, extensionattribute5, extensionattribute2, enabled, memberof

#Apply Actions Per Requirement1
foreach ($user in $users)
{
#2015-11-13 KHi Replacing check on EA2 is not filled
#if (($user.enabled -eq $true) -and ($user.extensionattribute2 -ne "OE3")) 
if (($user.enabled -eq $true) -and ([string]::IsNullOrEmpty($user.extensionattribute2))) 
    {
    set-aduser $user -clear "extensionattribute2"
    set-aduser $user -add @{"extensionattribute2"="OE3"} 
    add-adgroupmember $group $user
    add-adgroupmember $group1 $user
    add-content $addfileint "$($user.samaccountname), $($user.co), $($user.comment), $($user.employeeid), $($user.employeetype), $($user.extensionattribute5), $($user.extensionattribute2), $($user.enabled)"
    }
else 
    {}
}


#Create Report File For Requirement2 
$remfileint=new-item -path "D:\Scripts\O365Users\Reports" -name ("InternalUsersRemoved" + $date +".csv") -itemtype file
add-content $remfileint "Samaccountname, Co, Comment, EmployeeID, Employeetype, Extensionattribute5, Extensionattribute2, Enabled"

#Filter For User Accounts Per Requirement2
$remusers=get-aduser -filter {(memberof -ne $group4)} -properties co, comment, employeeid, employeetype, extensionattribute5, extensionattribute2, enabled, memberof

#Apply Actions Per Requirement2
foreach ($user in $remusers)
{
#2015-11-13 KHi Replacing check on EA2 is filled
#    if (((($user.comment -ne "active") -or ($user.enabled -eq $False) -or ($user.memberof -eq $group2) -or ($user.memberof -eq $group3)  -or (($user.employeetype -ne "active employee") -and ($user.employeetype -ne "Expats/ Inpats") -and ($user.employeetype -ne "Retiree/ Pensioner"))) -and ($user.extensionattribute2 -eq "OE3")))
    if ((-not [string]::IsNullOrEmpty($user.extensionattribute2)) -and
            ( 
                ($user.comment -ne "active") -or 
                ($user.enabled -eq $False) -or 
                ($user.memberof -eq $group2) -or 
                ($user.memberof -eq $group3) -or 
                (($user.employeetype -ne "active employee") -and ($user.employeetype -ne "Expats/ Inpats") -and ($user.employeetype -ne "Retiree/ Pensioner"))
            )
        )
     {
    set-aduser $user -clear "extensionattribute2"
    remove-adgroupmember $group $user  -Confirm:$False
    remove-adgroupmember $group1 $user -Confirm:$False
	add-content $remfileint "$($user.samaccountname), $($user.co), $($user.comment), $($user.employeeid), $($user.employeetype), 	$($user.extensionattribute5), $($user.extensionattribute2), $($user.enabled)"
     }
     else
     {}
}


#Create Report File For Requirement3
$addfileext1=new-item -path "D:\Scripts\O365Users\Reports" -name ("Ext1UsersAdded" + $date +".csv") -itemtype file
add-content $addfileext1 "Samaccountname, Co, Comment, EmployeeID, Employeetype, Extensionattribute5, Extensionattribute2, Enabled"

#Filter For User Accounts Per Requirement3
$ous="OU=managed users,dc=corp,dc=demb,dc=com", "OU=asp,ou=support,dc=corp,dc=demb,dc=com"
$usersext1=$ous | ForEach {Get-ADUser -Filter *  -properties co, comment, employeeid, employeetype, extensionattribute5, extensionattribute2, enabled, memberof -SearchBase $_}

#Apply Actions Per Requirement3
foreach ($user in $usersext1)
{
#2015-11-13 KHi Replacing check on EA2 is not filled
#if (($user.enabled -eq $true) -and ($user.memberof -eq $group4) -and ($user.employeeid -ge 0) -and ($user.extensionattribute2 -ne "OE3"))
if (($user.enabled -eq $true) -and ($user.memberof -eq $group4) -and ($user.employeeid -ge 0) -and ([string]::IsNullOrEmpty($user.extensionattribute2)))
	{
    set-aduser $user -clear "extensionattribute2"
    set-aduser $user -add @{"extensionattribute2"="OE3"} 
	add-content $addfileext1 "$($user.samaccountname), $($user.co), $($user.comment), $($user.employeeid), $($user.employeetype), $($user.extensionattribute5), $($user.extensionattribute2), $($user.enabled)"
	}
else
	{}
}


#Create Report File For Requirement4 
$addfileext2=new-item -path "D:\Scripts\O365Users\Reports" -name ("Ext2UsersAdded" + $date +".csv") -itemtype file
add-content $addfileext2 "Samaccountname, Co, Comment, EmployeeID, Employeetype, Extensionattribute5, Extensionattribute2, Enabled"

#Filter For User Accounts Per Requirement4
$usersext2=get-aduser -filter * -properties co, comment, employeeid, employeetype, extensionattribute5, extensionattribute2, enabled, memberof | where-object {($_.distinguishedname -notlike "*anaged user*") -and ($_.distinguishedname -notlike "*u=asp,ou=supp*")} 

#Apply Actions Per Requirement4
foreach ($user in $usersext2)
{
#2015-11-13 KHi Replacing check on EA2 is not filled
#if ((($user.memberof -eq $group4) -and ($user.extensionattribute2 -ne "OE3")))
if (($user.memberof -eq $group4) -and ([string]::IsNullOrEmpty($user.extensionattribute2)))
	{
    set-aduser $user -clear "extensionattribute2"
    set-aduser $user -add @{"extensionattribute2"="OE3"} 
	add-content $addfileext2 "$($user.samaccountname), $($user.co), $($user.comment), $($user.employeeid), $($user.employeetype), $($user.extensionattribute5), $($user.extensionattribute2), $($user.enabled)"
	}
else
	{}
}

#Create Report File For Requirement5
$addfileext3=new-item -path "D:\Scripts\O365Users\Reports" -name ("O365ExtUsersRemoved" + $date +".csv") -itemtype file
add-content $addfileext3 "Samaccountname, Co, Comment, EmployeeID, Employeetype, Extensionattribute5, Extensionattribute2, Enabled"

#Filter For User Accounts Per Requirement5
$usersext3=get-aduser -filter {(memberof -ne $group4) -and (memberof -ne $group1) -and (enabled -eq $false) -or (memberof -eq $group3)} -properties co, comment, employeeid, employeetype, extensionattribute5, extensionattribute2, enabled, memberof  

#Apply Actions Per Requirement5
foreach ($user in $usersext3)
{
#2015-11-13 KHi Replacing check on EA2 is filled
#if ($user.extensionattribute2 -eq "OE3")
if (-not [string]::IsNullOrEmpty($user.extensionattribute2))
	{
	set-aduser $user -clear "extensionattribute2"
	add-content $addfileext3 "$($user.samaccountname), $($user.co), $($user.comment), $($user.employeeid), $($user.employeetype), $($user.extensionattribute5), $($user.extensionattribute2), $($user.enabled)"
	}
else
	{}
}


#Create Report File For Requirement6
$addfileext4=new-item -path "D:\Scripts\O365Users\Reports" -name ("VPNAccessExtUsersRemoved" + $date +".csv") -itemtype file
add-content $addfileext4 "Samaccountname, Co, Comment, EmployeeID, Employeetype, Extensionattribute5, Extensionattribute2, Enabled"

#Filter For User Accounts Per Requirement6
$usersext4=get-aduser -filter 'enabled -eq $false' -properties co, comment, employeeid, employeetype, extensionattribute5, extensionattribute2, enabled, memberof  

#Apply Actions Per Requirement6
foreach ($user in $usersext4)
{
if (($user.memberof -eq $group3))
	{
	remove-adgroupmember $group5 $user -Confirm:$False
	add-content $addfileext4 "$($user.samaccountname), $($user.co), $($user.comment), $($user.employeeid), $($user.employeetype), $($user.extensionattribute5), $($user.extensionattribute2), $($user.enabled)"
	}
else
	{}
}


#Define Server And Recipients
$smtpServer = "smtp.corp.demb.com"
$RecipientList = ("IT.Security@demb.com", "ivaylop@hp.com")
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$msg.From = "ADSecurity@demb.com"

#Send Report File1 In E-mail
foreach ($Recipient in $RecipientList)
    {
       $msg.To.Add($Recipient)
    }
$emailattachment = $addfileint 
$msg.Subject = "Attached - users added to AllowVPNAccess and AllowO365access groups and OE3 Set"
$attachment = New-Object System.Net.Mail.Attachment($emailattachment)
$msg.Attachments.Add($attachment)
$smtp.Send($msg)
$attachment.Dispose()

#Define Server And Recipients
$smtpServer = "smtp.corp.demb.com"
$RecipientList = ("IT.Security@JDECoffee.com", "ivaylop@hpe.com")
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$msg.From = "ADSecurity@JDECoffee.com"

#Send Report File2 In E-mail
foreach ($Recipient in $RecipientList)
    {
       $msg.To.Add($Recipient)
    }
$emailattachment = $remfileint
$msg.Subject = "Attached - users removed from AllowVPNAccess group and AllowO365access groups and OE3 cleared"
$attachment = New-Object System.Net.Mail.Attachment($emailattachment)
$msg.Attachments.Add($attachment)
$smtp.Send($msg)
$attachment.Dispose()

#Define Server And Recipients
$smtpServer = "smtp.corp.demb.com"
$RecipientList = ("IT.Security@JDECoffee.com", "ivaylop@hpe.com")
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$msg.From = "ADSecurity@JDECoffee.com"

#Send Report File3 In E-mail
foreach ($Recipient in $RecipientList)
    {
       $msg.To.Add($Recipient)
    }
$emailattachment = $addfileext1
$msg.Subject = "Attached - External users in Managed Users and ASP OUs with OE3 set"
$attachment = New-Object System.Net.Mail.Attachment($emailattachment)
$msg.Attachments.Add($attachment)
$smtp.Send($msg)
$attachment.Dispose()

#Define Server And Recipients
$smtpServer = "smtp.corp.demb.com"
$RecipientList = ("IT.Security@JDECoffee.com", "ivaylop@hpe.com")
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$msg.From = "ADSecurity@JDECoffee.com"

#Send Report File4 In E-mail
foreach ($Recipient in $RecipientList)
    {
       $msg.To.Add($Recipient)
    }
$emailattachment = $addfileext2
$msg.Subject = "Attached - External users outside Managed Users and ASP OUs with OE3 set"
$attachment = New-Object System.Net.Mail.Attachment($emailattachment)
$msg.Attachments.Add($attachment)
$smtp.Send($msg)
$attachment.Dispose()

#Define Server And Recipients
$smtpServer = "smtp.corp.demb.com"
$RecipientList = ("IT.Security@JDECoffee.com", "ivaylop@hpe.com")
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$msg.From = "ADSecurity@JDECoffee.com"

#Send Report File5 In E-mail
foreach ($Recipient in $RecipientList)
    {
       $msg.To.Add($Recipient)
    }
$emailattachment = $addfileext3
$msg.Subject = "Attached - All External users with OE3 cleared"
$attachment = New-Object System.Net.Mail.Attachment($emailattachment)
$msg.Attachments.Add($attachment)
$smtp.Send($msg)
$attachment.Dispose()

#Define Server And Recipients
$smtpServer = "smtp.corp.demb.com"
$RecipientList = ("IT.Security@JDECoffee.com", "ivaylop@hpe.com")
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$msg.From = "ADSecurity@JDECoffee.com"

#Send Report File6 In E-mail
foreach ($Recipient in $RecipientList)
    {
       $msg.To.Add($Recipient)
    }
$emailattachment = $addfileext4
$msg.Subject = "Attached - Users with revoked VPN Access"
$attachment = New-Object System.Net.Mail.Attachment($emailattachment)
$msg.Attachments.Add($attachment)
$smtp.Send($msg)
$attachment.Dispose()
