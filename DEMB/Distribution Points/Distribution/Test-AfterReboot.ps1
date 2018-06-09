<#
    Test-AfterReboot.ps1

    Take action after a reboot took place.

    --- Version history
    Version 1.00 (2016-05-03, Kees Hiemstra)
    - Initial version.
#>

. "$($PSScriptRoot)\CommonEnvironment.ps1"

$ScriptName = "Test-AfterReboot"

Write-Log -Message "Server has rebooted"

Send-SOEMessage -Subject "Server $($Env:ComputerName) rebooted" -Body "Server has rebooted" -MailType "reboot"

Stop-Script
