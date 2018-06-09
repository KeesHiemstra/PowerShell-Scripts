$List = Import-Csv -Path B:\Mail2ComputerName.csv -Delimiter "`t"

foreach ( $Item in $List )
{
    $Mail = $Item.Mail

    #try
    #{
        $Item.SAMAccountName = (Get-ADUser -Filter { Mail -eq $Mail } -ErrorAction Stop ).SAMAccountName
        $Item.ComputerName = (Get-AMAsset -SAMAccountName $Item.SAMAccountName).ComputerName -join '; '
        $Item
    ##}
    #catch
    #{
    #}
}

$List | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation | clip
