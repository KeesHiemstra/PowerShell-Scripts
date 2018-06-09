#region Messaging
function Send-Error([string]$Subject, [string]$Message)
{
    Send-MailMessage -From "HPDesktop.Administrator@demb.com" `
        -SMTPserver "smtp.corp.demb.com" `
        -To "Kees.Hiemstra@hp.com" `
        -Subject $Subject `
        -Body $Message `
        -Priority High `
        -DeliveryNotificationOption onFailure
}

function Send-Notification([string]$Subject, [string]$Message)
{
    Send-MailMessage -From "HPDesktop.Administrator@demb.com" `
        -SMTPserver "smtp.corp.demb.com" `
        -To "Kees.Hiemstra@hp.com" `
        -Subject $Subject `
        -Body $Message `
        -DeliveryNotificationOption onFailure
}
#endregion

$DayDiffMax = -1

if ((Get-Date).DayOfWeek -eq "Saturday") {$DayDiffMax = -2}
if ((Get-Date).DayOfWeek -eq "Sunday") {$DayDiffMax = -3}

$DateDiff = (Get-Date).AddDays($DayDiffMax)
if ((Get-ChildItem -Path "\\DEMBMCAPS032SQ2.corp.demb.com\ITAM`$\Usr\Report\ITAM19\Customer\Computer_in_contract.xlsx").LastWriteTime -lt $DateDiff)
{
    Write-Host "It looks like Computer_in_contract.xlsx is too old"
    Send-Error -Subject "Old reports on ITAMWeb" -Message "It look like Computer_in_contract.xlsx is too old"
}
