<#
    ManageToBeOrphanedAccounts.Config.ps1
#>

#Mail settings
$MailSMTP = "smtp.corp.demb.com"
$MailSender = "No reply<AccountChange@JDECoffee.com>"
$MailErrorTo = "Kees.Hiemstra@hpe.com"

#Define log file with the name as the script
$LogFile = $MyInvocation.MyCommand.Definition -replace(".ps1$",".log")
