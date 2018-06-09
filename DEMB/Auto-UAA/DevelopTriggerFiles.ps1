#Test

Set-Location -Path C:\Src\PowerShell\!DEMB\Auto-UAA

$processPath = $PSScriptRoot + "\process\"
$completedPath = $PSScriptRoot + "\completed\"

$ScheduledForMail = Get-ChildItem ($processPath + "\*.mail.txt") | ForEach-Object {$_.name -Replace(".mail.txt$","")}

foreach($User in $ScheduledForMail)
{
    [string[]]$Content = Get-Content -Path "$processPath\$User.mail.txt"
    $UaaApp = $Content[0] -match "\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} Created by the UAA-App"



    Write-Host "$User`: $UaaApp"
}