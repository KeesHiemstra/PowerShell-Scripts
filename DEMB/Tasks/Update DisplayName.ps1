$Update = Import-Csv -Path 'B:\UpdateDisplayName.csv' -Delimiter "`t"

foreach ( $Item in $Update )
{
    $SAMAccountName = $Item.SAMAccountName
    $DisplayName = $Item.DisplayName
    Set-ADUser $SAMAccountName -DisplayName $DisplayName -Server DEMBDCRS002
}