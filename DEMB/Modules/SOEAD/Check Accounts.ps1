$Users = Import-Csv -Path B:\CheckAccounts.txt -Delimiter "`t"

foreach ( $Item in $Users )
{
    try
    {
        $UPN = $Item.mail
        $ADUser = Get-ADUser -Filter { mail -eq $UPN } -Properties EmployeeType, ExtensionAttribute5, ExtensionAttribute2, ExtensionAttribute11 -ErrorAction Stop
        $Item.EmployeeType = $ADUser.EmployeeType
        $Item.ExtensionAttribute5 = $ADUser.ExtensionAttribute5
        $Item.ExtensionAttribute2 = $ADUser.ExtensionAttribute2
        $Item.ExtensionAttribute11 = $ADUser.ExtensionAttribute11
    }
    catch
    {
    }
}

$Users | Export-Csv -Path B:\CheckAccounts.txt -Delimiter "`t" -NoTypeInformation
