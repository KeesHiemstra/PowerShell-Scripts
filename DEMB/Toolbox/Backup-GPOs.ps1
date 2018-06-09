$Date = get-date -UFormat "%Y-%m-%d"
 
#backup all
$path = "\\DEMBMCAPS032FG1.corp.demb.com\SOE\GPO\Backup\Full\$($Date)"
New-Item -Path $Path -ItemType directory
Backup-GPO -All -Path $path
 
#backup DEMB-* gpo's 
 
$gpos = (Get-GPO -all | ?{$_.DisplayName -like "SOE-*"}).displayname
 
foreach ($Gpo in $gpos)
{
    $Path = "\\DEMBMCAPS032FG1.corp.demb.com\SOE\GPO\Backup\SingleGPO\$($GPO)\$($Date)"
 
    New-Item -Path $Path -ItemType directory
 
    Get-GPOReport $gpo -ReportType Html -path "$($Path)\$($Gpo).htm"
    Get-GPOReport $gpo -ReportType Xml -Path "$($Path)\$($Gpo).xml"
    Backup-GPO $gpo -Path $Path -Comment $gpo
} 

