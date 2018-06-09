
$FullBUPath = '\\DEMBMCAPS032FG1.corp.demb.com\SOE\GPO\Backup\Full'
$SingleBUPath = '\\DEMBMCAPS032FG1.corp.demb.com\SOE\GPO\Backup\SingleGPO'

$Date = (Get-Date).ToString('yyyy-MM-dd HHmm')

#region Full backup
New-Item -Path "$FullBUPath\$Date" -Type Directory | Out-Null
Backup-GPO -All -Path "$FullBUPath\$Date" -Verbose
#endregion

#region Individual backup
$GPOs = (Get-GPO -All | Where-Object { $_.DisplayName -like 'HPE-SOE*' }).DisplayName

foreach ( $Item in $GPOs )
{
    $BackupPath = "$SingleBUPath\$Item\$Date"

    New-Item -Path $BackupPath -Type Directory | Out-Null

    Get-GPOReport $Item -ReportType Html -Path "$BackupPath\$Item.html"
    Get-GPOReport $Item -ReportType Xml -Path "$BackupPath\$Item.xml"
    Backup-GPO -Name $Item -Path $BackupPath -Comment $Item
}
#endregion