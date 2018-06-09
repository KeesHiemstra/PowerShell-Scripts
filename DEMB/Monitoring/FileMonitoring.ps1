<#
    Check the existence of files.
#>

$CheckPaths = Import-Csv -Path 'C:\Src\PowerShell\5.0\Monitoring\FileMonitoring.csv'

foreach ( $Item in $CheckPaths )
{
    $Files = Get-ChildItem -Path $Item.Path

    if ( $Files.Count -eq 1 -and (Get-Date).Minute -ge $Item.Time )
    {
        Write-Host "There is a file in $($Item.Name) and it seems like it isn''t processed yet."
    }
    elseif ( $Files.Count -gt 1 )
    {
        Write-Host "There are more than 1 files in $($Item.Name)."
    }

}
