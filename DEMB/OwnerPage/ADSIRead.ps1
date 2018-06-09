#Start of the search
$Domain = New-Object -TypeName System.DirectoryServices.DirectoryEntry

$Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher
$Searcher.SearchRoot = $Domain
$Searcher.PageSize = 1000
$Searcher.SearchScope = "Subtree"

$Properties = @('Name', 'DisplayName', 'Mail', 'C', 'L', 'UserAccountControl')
foreach ($Property in $Properties) { $Searcher.PropertiesToLoad.Add($Property) | Out-Null }
#                   (&(objectCategory=group)(objectClass=group)(CN=0002O_RD_Temp_Temp))
$Searcher.Filter = "(&(objectCategory=group)(objectClass=group)(CN=0002C_RD_Temp_Temp))"

$Group = $Searcher.FindOne()
if ($Group.Count -eq 1)
{
    $GroupMembers = @()
    foreach ($Member in ([ADSI]$Group.Path).Member)
    {
        $Searcher.Filter = "(&(objectCategory=person)(objectClass=user)(DistinguishedName=$Member))"
        $ADUser = $Searcher.FindOne()
        $GroupMembers += New-Object -TypeName PSObject -Property ([ordered] @{displayName = $ADUser.Properties.displayname -join ",";
                             mail = $ADUser.Properties.mail -join ",";
                             c = $ADUser.Properties.c -join ",";
                             l = $ADUser.Properties.l -join ",";
                             Enabled = (([uint64]($ADUser.Properties.useraccountcontrol -join ",")) -band 2) -eq 0;
                             })    
    }
    $GroupMembers
}