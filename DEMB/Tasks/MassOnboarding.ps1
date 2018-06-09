<#
    Mass on boarding of user, resetting their password and capture the password in the input file
#>

#region Functions
function New-Password([int]$Length){    $PW = $null    # exclude: o0O, 1lI etc.    $Forbidden = @('o','0','O',',',' ','1','l','I',"'",'"','`','?','^',';','/','\','~','|')    #define the number of charaters needed for each charactergroup (round up)    $c = [math]::ceiling(($Length / 4))    #numbers    $PW += (0..9 | where-object {$_ -notin $Forbidden} | Get-Random -Count $c)    #uppercase    $PW += (65..90 | %{[char]$_} | where-object {$_ -notin $Forbidden} | Get-Random -Count $c)    #lowercase    $PW += (97..122 | %{[char]$_} | where-object {$_ -notin $Forbidden} | Get-Random -Count $c)    #Special    $PW += ((32..47 + 58..64)  | %{[char]$_} | where-object {$_ -notin $Forbidden} |Get-Random -Count $c)    #Radomize the password order and limit to requested chars    $PW = ([system.string]::Join("",($PW | Get-Random -count $Length)))    RETURN $PW}

#endregion

$User = Import-Csv -Path B:\MassOnboarding.csv -Delimiter "`t"

foreach ( $Item in ($User) )
{

    $EmployeeID = $Item.EmployeeID
    try
    {
        $Item.Comment = "Account not found"
        $ADUser = Get-ADUser -Filter { EmployeeID -eq $EmployeeID } -Properties Mail -ErrorAction Stop
        $Item.SAMAccountName = $ADUser.SamAccountName
        $Item.Mail = $ADUser.Mail

        $Item.Comment = "Password not set for $($ADUser.SAMAccountName)"
        $Password = New-Password(10)
        Set-ADAccountPassword -Identity $ADUser.SAMAccountName -NewPassword (ConvertTo-SecureString -AsPlainText $Password -Force)

        $Item.Password = $Password
        $Item.Comment = "Password changed for $($ADUser.SAMAccountName)"
    }
    catch
    {
        Write-Host 'Error'
    }
    $Item.Comment
}


$User | Export-Csv -Path B:\MassOnboarding.csv -Delimiter "`t" -NoTypeInformation