﻿<#
    Mass on boarding of user, resetting their password and capture the password in the input file
#>

#region Functions


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