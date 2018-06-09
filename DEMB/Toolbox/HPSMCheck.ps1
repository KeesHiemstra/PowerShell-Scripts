#Create exports
$OldTime = "20150915_1553"
$NewTime = (Get-Date).ToString("yyyyMMdd_HHmm")
Write-Host $NewTime
Get-ADUser -Filter * -SearchBase "OU=Managed Users,DC=corp,DC=demb,DC=com" -Properties company | Where {$_.Enabled -eq $true } | Group-Object company | Select-Object Count, Name | Export-Csv -Path "D:\Data\Jobs\ADAS Data\ADAS-Company-v$($NewTime).csv" -Delimiter "`t" -NoTypeInformation -Encoding UTF8
Get-ADUser -Filter * -SearchBase "OU=Managed Users,DC=corp,DC=demb,DC=com" -Properties c, l | Where {$_.Enabled -eq $true } | Group-Object c, l | Select-Object Count, Name | Export-Csv -Path "D:\Data\Jobs\ADAS Data\ADAS-Location-v$($NewTime).csv" -Delimiter "`t" -NoTypeInformation -Encoding UTF8

#Check differences

#Company
$Ref = Import-Csv -Path "D:\Data\Jobs\ADAS Data\ADAS-Company-v$($OldTime).csv" -Delimiter "`t" -Encoding UTF8 | Sort-Object Name
$Dif = Import-Csv -Path "D:\Data\Jobs\ADAS Data\ADAS-Company-v$($NewTime).csv" -Delimiter "`t" -Encoding UTF8 | Sort-Object Name
$ChangesC = Compare-Object -ReferenceObject $Ref -DifferenceObject $Dif -Property Name


$Ref = Import-Csv -Path "D:\Data\Jobs\ADAS Data\ADAS-Location-v$($OldTime).csv" -Delimiter "`t" -Encoding UTF8 | Sort-Object Name
$Dif = Import-Csv -Path "D:\Data\Jobs\ADAS Data\ADAS-Location-v$($NewTime).csv" -Delimiter "`t" -Encoding UTF8 | Sort-Object Name
$ChangesL = Compare-Object -ReferenceObject $Ref -DifferenceObject $Dif -Property Name

$ChangesC | Select-Object Name, @{n='Status'; e={if($_.SideIndicator -eq '=>'){'New'}else{'Deleted'}}} | Format-Table -AutoSize
Write-Host "====="
$ChangesL | Select-Object Name, @{n='Status'; e={if($_.SideIndicator -eq '=>'){'New'}else{'Deleted'}}} | Format-Table -AutoSize

break
#Company
Get-ADUser -Filter * -SearchBase "OU=Managed Users,DC=corp,DC=demb,DC=com" -Properties company | Where {$_.Enabled -eq $true } | Group-Object company | Select-Object Count, Name | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation | Clip.exe

break
#Locations
Get-ADUser -Filter * -SearchBase "OU=Managed Users,DC=corp,DC=demb,DC=com" -Properties c, l | Where {$_.Enabled -eq $true } | Group-Object c, l | Select-Object Count, Name | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation | Clip.exe

break
#OpCo vs location
Get-ADUser -Filter * -SearchBase "OU=Managed Users,DC=corp,DC=demb,DC=com" -Properties c, l, company | Where {$_.Enabled -eq $true } | Group-Object c, l, company | Select-Object Count, Name | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation | Clip.exe
