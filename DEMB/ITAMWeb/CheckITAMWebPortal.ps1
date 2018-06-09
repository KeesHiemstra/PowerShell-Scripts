<#
.Synopsis
    Check if the ITAMWeb portal is working and start the service if this is not the case.
.DESCRIPTION
    The IIS service is not always automatically started after the cluster node is (re)started.
    
    The root cause of this issue is not known.
    
    This controler will check if the IIS service is running and if not, it will start it.
    
    Before it can do this, it has to determine on which node the IIS service should be running. 
.OUTPUTS
    An email will be send if the service is started.
.NOTES
    Author         : Kees Hiemstra (Kees.Hiemstra@hp.com // chi@xs4all.nl)
    PSVerion       : 4.0
    Assemblies     : None

    Version history:

    Version 1.01 (2014-11-24, Kees Hiemstra)
    - Added this comment and tested it to run under a the svc.HPSoEAdmin service account.
    Version 1.00 (2014-10-09, Kees Hiemstra)
    - Inital version
#>
Import-Module FailoverClusters

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

$ClusterName = "DEMBMCAPS032CS.corp.demb.com"
#$Nodes = Get-ClusterNode -Cluster $ClusterName

$SQLNode = (Get-ClusterResource -Cluster $ClusterName -Name "SQL Server (PROD_2)").OwnerNode.Name

if ((Get-Service -ComputerName $SQLNode -Name "W3Svc").Status -ne "Running")
{
    #Start-Service can't be started remotely
    Invoke-Command -ComputerName $SQLNode -ScriptBlock {Start-Service -Name "W3SVC"}

    if ((Get-Service -ComputerName $SQLNode -Name "W3Svc").Status -ne "Running")
    {
        Send-Error -Subject "IIS service of ITAMWeb is not starting." -Message ("The IIS service on {0} is not started after initial restart." -f ($SQLNode))
    }
    else
    {
        Send-Notification -Subject "IIS service of ITAMWeb is started." -Message ("The IIS service on {0} started sucessfully after it was not running in the first place." -f ($SQLNode))
    }
}
else
{
    Write-Host "IIS is running, ITAMWeb should be accessible."
}

$CountExcel1 = (Get-Process -ComputerName $SQLNode EXCEL -ErrorAction SilentlyContinue).Count
Write-Host "$CountExcel1 number of Excel processes."
if ($CountExcel1 -gt 20)
{
    Write-Warning "More that 20 processes of Excel are running."
    Invoke-Command -ComputerName $SQLNode -ScriptBlock { Stop-Process -Name EXCEL -Confirm:$false }
    $CountExcel2 = (Get-Process -ComputerName $SQLNode EXCEL -ErrorAction SilentlyContinue).Count
    if ($CountExcel2 -gt 5)
    {
        Send-Error -Subject "Too many Excel processes are still running on $SQLNode" -Message "$CountExcel2 number of Excel processes are running whilst $CountExcel1 triggered this process."
    }
    else
    {
        Send-Notification -Subject "Too many Excel processes were running on $SQLNode" -Message "$CountExcel1 number of Excel processes were running, but these have been stopped."
    }
}