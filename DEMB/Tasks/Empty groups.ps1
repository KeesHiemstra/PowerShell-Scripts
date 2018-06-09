#Empty groups
(Get-ADGroup -Filter { GroupCategory -eq 'Security' -and GroupScope -eq 'Global' } -Properties Member | Where-Object { $_.Name -notmatch '\d\d\d\d._' -and $_.Name -notlike '*-Admins' -and ($_.Member).Count -eq 0 }).Count
(Get-ADGroup -Filter { GroupCategory -eq 'Security' -and GroupScope -eq 'Global' } -Properties Name, Description, Info, Member | Where-Object { $_.Name -notmatch '\d\d\d\d._' -and $_.Name -notlike '*-Admins' -and ($_.Member).Count -eq 0 }) | Select-Object Name, Description, Info | ft *
(Get-ADGroup -Filter { GroupCategory -eq 'Security' -and GroupScope -eq 'Global' } -Properties Name, Description, Info, Member | Where-Object { $_.Name -notmatch '\d\d\d\d._' -and $_.Name -notlike '*-Admins' -and ($_.Member).Count -eq 0 }).Info

#Groups with 1 member
(Get-ADGroup -Filter { GroupCategory -eq 'Security' -and GroupScope -eq 'Global' } -Properties Member | Where-Object { $_.Name -notmatch '\d\d\d\d._' -and ($_.Member).Count -eq 1 }).Count

#Empty groups D18
(Get-ADGroup -Filter { Name -like 'D18*' -and GroupCategory -eq 'Security' -and GroupScope -eq 'Global' } -Properties Member | Where-Object { $_.Name -notmatch '\d\d\d\d._' -and ($_.Member).Count -eq 0 }).Count


Get-ADGroup -Filter { GroupCategory -eq 'Security' -and GroupScope -eq 'Global' } -Properties Name, Info | Where-Object { $_.Info -match 'westbroek|@hp\.com|thoma|luit|patrick|demb' } | ft Name, Info -AutoSize