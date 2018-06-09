$ToBeDeletedGroup = 'To-Be-Deleted-Jan'
$Threshold = 400

$GroupSize = (Get-ADGroup $ToBeDeletedGroup -Properties Member).Member.Count

Write-Host "$ToBeDeletedGroup has $GroupSize members"

break

#Move the remaining users to the 'To-Be-Deleted-Stack' group
$List = (Get-ADGroup $ToBeDeletedGroup -Properties Member).Member | Select-Object -First ($GroupSize - $Threshold)

Add-ADGroupMember 'To-Be-Deleted-Stack' -Members $List
Remove-ADGroupMember $ToBeDeletedGroup -Members $List -Confirm:$false

Write-Host "$($List.Count) have been put aside"
Write-Host "$ToBeDeletedGroup has $((Get-ADGroup $ToBeDeletedGroup -Properties Member).Member.Count) members left"

break

#Move $Treshold user back in to $ToBeDeledGroup
$List = (Get-ADGroup 'To-Be-Deleted-Stack' -Properties Member).Member | Select-Object -First ($Threshold)

Add-ADGroupMember $ToBeDeletedGroup -Members $List
Remove-ADGroupMember 'To-Be-Deleted-Stack' -Members $List -Confirm:$false

Write-Host "Put back $($List.Count) users into $ToBeDeletedGroup"
Write-Host "'To-Be-Deleted-Stack' has $((Get-ADGroup 'To-Be-Deleted-Stack' -Properties Member).Member.Count) members left to be put back"


