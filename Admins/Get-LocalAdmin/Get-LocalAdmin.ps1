#
# Get-LocalAdmin.ps1
#

function Get-LocalAdmin
{
    param ($Computer)

    $Admins = Get-WmiObject Win32_GroupUser –Computer $Computer
    $Admins = $Admins |Where-Object {$_.GroupComponent –like '*"Administrators"'}

    $Admins | ForEach-Object {
        $_.PartComponent –match ".+Domain\=(.+)\,Name\=(.+)$" > $Nul
        $Matches[1].Trim('"') + "\" + $Matches[2].Trim('"')
    }
}

Get-LocalAdmin .