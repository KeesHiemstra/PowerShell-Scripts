<#
    Collect the members of the local support groups and convert it into an SQL command
    to update the special users table.
#>

$Users = (Get-ADGroup -Filter "name -like 'ManageWkstn-LS-*'" |
    Get-ADGroupMember |
    Where-Object { $_.ObjectClass -eq 'user' }).sAMAccountName

$SQL = ''
foreach ($User in $Users)
{
    $SQL += (@"
        IF NOT EXISTS (SELECT 1 FROM ITAMData.dbo.UserAccountSpecial WHERE [UserAccount] = 'demb\{0}') 
        INSERT INTO ITAMData.dbo.UserAccountSpecial([UserAccount], [Description], [ComputerCount]) 
        VALUES('demb\{0}', 'JDE local support engineer', -1)
        ELSE UPDATE ITAMData.dbo.UserAccountSpecial
        SET [Description] = 'JDE local support engineer'
        WHERE [UserAccount] = 'demb\{0}' AND [Description] NOT LIKE 'JDE local support engineer%'


"@ -f $User.ToLower())
}

$SQL | Clip



