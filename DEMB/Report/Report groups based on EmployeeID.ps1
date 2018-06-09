$User = Import-Csv -Path 'B:\GroupList.csv' -Delimiter "`t"
$Result = @()

foreach ( $Item in $User )
{
    $EmployeeID = ($Item.EmployeeID).PadLeft(8, '0')

    $ADUser = Get-ADUser -Filter { EmployeeID -eq $EmployeeID } -Properties EmployeeID, SAMAccountName, MemberOf
    $Item.SAMAccountName = $ADUser.SAMAccountName

    foreach ( $Group in $ADUser.MemberOf )
    {
        $Result += New-Object -TypeName PSObject -Property ([ordered] @{EmployeeID = $EmployeeID; SAMAccountName = $ADUser.SamAccountName; Group = ($Group -Replace 'CN=' -Replace ',.*com$')})
    }
}

$Result | Export-Csv -Path B:\AlexGroupCheck.csv -NoTypeInformation
