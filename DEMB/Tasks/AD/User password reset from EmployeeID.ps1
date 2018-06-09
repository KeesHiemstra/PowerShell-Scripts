<#
    Uset password reset from EmployeeID.ps1

    B:\UserPasswordReset.txt: EmployeeID is the key attribute, the file contains minimal the following fields separated with tabs
    EmployeeID	SAMAccountName	Mail	Password	Result

    SAMAccountName, Mail, Password, Result will be exported to the clipboard.

    === Version history
    Version 1.00 (2017-04-03, Kees Hiemstra)
    - Inital version.
#>

$List = Import-Csv -Path 'B:\UserPasswordReset.txt' -Delimiter "`t"

#region Functions
function New-Password([int]$Length){    $PW = $null    # exclude: o0O, 1lI etc.    $Forbidden = @('o','0','O',',',' ','1','l','I',"'",'"','`','?','^',';','/','\','~','|')    #define the number of charaters needed for each charactergroup (round up)    $c = [math]::ceiling(($Length / 4))    #numbers    $PW += (0..9 | where-object {$_ -notin $Forbidden} | Get-Random -Count $c)    #uppercase    $PW += (65..90 | %{[char]$_} | where-object {$_ -notin $Forbidden} | Get-Random -Count $c)    #lowercase    $PW += (97..122 | %{[char]$_} | where-object {$_ -notin $Forbidden} | Get-Random -Count $c)    #Special    $PW += ((32..47 + 58..64)  | %{[char]$_} | where-object {$_ -notin $Forbidden} |Get-Random -Count $c)    #Radomize the password order and limit to requested chars    $PW = ([system.string]::Join("",($PW | Get-Random -count $Length)))    RETURN $PW}

#endregion

foreach ( $Item in $List )
{
    $EmployeeID = $Item.EmployeeID.PadLeft(8, '0')
    try
    {
        $ADUser = Get-ADUser -Filter { EmployeeID -eq $EmployeeID } -Properties Mail, Company, LastLogonDate, PasswordLastSet -ErrorAction Stop
        $Item.SAMAccountName = $ADUser.SAMAccountName
        $Item.Mail = $ADUser.Mail

        if ( $Item.LastLogonDate -ne $null )
        {
            $Item.Result = "$($Item.SAMAccountName) has already logged in"
        }
        elseif ( $Item.PasswordLastSet -ne $null )
        {
            $Item.Result = "$($Item.SAMAccountName) has already changed password"
        }
        else
        {
            $Password = New-Password(10)
            Set-ADAccountPassword -Identity $ADUser.SAMAccountName -NewPassword (ConvertTo-SecureString -AsPlainText $Password -Force)
            Set-ADUser -Identity $ADUser.SAMAccountName -ChangePasswordAtLogon $false
            $Item.Password = $Password
            $Item.Result = "$($Item.SAMAccountName) password changed"
        }
    }
    catch
    {
       $Item.Result = "$EmployeeID does not exist"
    }
}

$List | Select-Object SAMAccountName, Mail, Password, Result | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation | Clip