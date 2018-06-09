<#
    Find semi-duplicate computers in Active Directory.

    "C:\Users\Kees.Hiemstra\OneDrive - JACOBS DOUWE EGBERTS (JDE)\Housekeeping\HouseKeeping.xlsm" second tab.
    Match with ITAM on serial number.
#>

$ADComputers = Get-ADComputer -Filter { SerialNumber -like '*' -and SerialNumber -ne '0' -and SerialNumber -notlike '*-*' } -Properties Name, Description, DistinguishedName, Enabled, IPv4Address, LastLogonDate, SerialNumber, Street, WhenChanged, WhenCreated |
    Where-Object { $_.SerialNumber.Trim() -ne '' } |
    Select-Object Name, Description, DistinguishedName, Enabled, IPv4Address, LastLogonDate, @{n='SerialNumber'; e={$_.SerialNumber -join ';'}}, Street, WhenChanged, WhenCreated

$Duplicates = $ADComputers | Group-Object SerialNumber | Where-Object { $_.Count -gt 1 }

$Duplicates.Count

$Report = $Duplicates.Group | Select-Object Name, LastLogonDate, IPv4Address, Enabled, Description, Street, WhenChanged, WhenCreated, DistinguishedName, SerialNumber |
    Sort-Object SerialNumber, LastLogonDate 
    
foreach ( $Item in $Report )
{
    $Asset = Get-AMAsset -SerialNo $Item.SerialNumber

    Add-Member -InputObject $Item -MemberType NoteProperty -Name 'Serial No' -Value $Asset.SerialNo
    Add-Member -InputObject $Item -MemberType NoteProperty -Name 'Asset name' -Value $Asset.ComputerName
    Add-Member -InputObject $Item -MemberType NoteProperty -Name 'Brand' -Value $Asset.Brand
    Add-Member -InputObject $Item -MemberType NoteProperty -Name 'Model' -Value $Asset.Model
    Add-Member -InputObject $Item -MemberType NoteProperty -Name 'Acquisition' -Value $Asset.DTAcquisition
    Add-Member -InputObject $Item -MemberType NoteProperty -Name 'Asset status' -Value $Asset.AssetStatus
    Add-Member -InputObject $Item -MemberType NoteProperty -Name 'Billing status' -Value $Asset.BillingStatus
    Add-Member -InputObject $Item -MemberType NoteProperty -Name 'OpCo' -Value $Asset.AssetOpCoFull
    Add-Member -InputObject $Item -MemberType NoteProperty -Name 'User' -Value "$($Asset.UserLastName), $($Asset.UserFirstName)"
    Add-Member -InputObject $Item -MemberType NoteProperty -Name 'E-mail address' -Value $Asset.UserEMail
    Add-Member -InputObject $Item -MemberType NoteProperty -Name 'Radia connect' -Value $Asset.DaysFromLastMutation
    Add-Member -InputObject $Item -MemberType NoteProperty -Name 'Country' -Value $Asset.LocationCountryName
    Add-Member -InputObject $Item -MemberType NoteProperty -Name 'Location' -Value $Asset.LocationName
}   

$Report | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation | Clip

Write-Speech -Message "The AD task has completed."