$Check = Import-Csv B:\EmployeeNumbersToCheck.txt -Delimiter "`t"

foreach ( $Item in $Check )
{
    $User = Get-SOEADUser -EmployeeID $Item.EmployeeID -ErrorAction SilentlyContinue

    if ( $User -ne $null )
    {
        $Item.SAMAccountName = $User.SAMAccountName
        $Item.Comment = $User.Comment
        $Item.Enabled = $User.Enabled
    }
}

