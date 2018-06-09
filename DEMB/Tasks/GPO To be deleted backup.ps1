$SingleBUPath = '\\DEMBMCAPS032FG1.corp.demb.com\SOE\GPO\Backup\To be deleted'

$Date = (Get-Date).ToString('yyyy-MM-dd HHmm')

#region Individual to be deleted GPO backup
$GPOs = (Get-GPO -All | Where-Object { $_.DisplayName -like 'HPE-DELETE*' }).DisplayName

foreach ( $Item in $GPOs )
{
    $BackupPath = "$SingleBUPath\$Item\$Date"

    New-Item -Path $BackupPath -Type Directory | Out-Null

    Get-GPOReport $Item -ReportType Html -Path "$BackupPath\$Item.html"
    Get-GPOReport $Item -ReportType Xml -Path "$BackupPath\$Item.xml"
    Backup-GPO -Name $Item -Path $BackupPath -Comment $Item
}
#endregion