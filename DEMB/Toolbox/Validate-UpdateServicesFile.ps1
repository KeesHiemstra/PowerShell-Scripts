Clear-Host
$Update = Import-Csv -Path "\\Dembmcaps032fg1\soe\Scripts\SAP files distribution\2015-03-24.01\UpdateServicesfile.csv" | Sort-Object Port
$Update | Format-Table -AutoSize

$WarningCount = 0

if (($Update | Where-Object { $_.Action -notin ('ADD', 'REMOVE') }).Count -gt 0)
{
    $WarningCount++
    Write-Warning "The action column contains other actions than Add and Remove."
}

if (($Update | Where-Object { $_.Port -lt 3600 -or $_.Port -gt 3699 }).Count -gt 0 )
{
    $WarningCount++
    Write-Warning "The port column contains other ports that don't fall in the range 3600-3699."
}

if (($Update | Where-Object { $_.Protocol -notin ('tcp', 'udp') }).Count -gt 0)
{
    $WarningCount++
    Write-Warning "The protocol column contains other protocols than tcp and udp."
}

if ($WarningCount -ne 0)
{
    Write-Warning "Warnings found!"
}