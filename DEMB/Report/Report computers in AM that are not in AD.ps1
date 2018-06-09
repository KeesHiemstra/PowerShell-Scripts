$AM = Get-AMAsset -IsDesktop | Where-Object { $_.Category -ne 'Tablet' -and $_.Category -ne 'Thin client' }

$AD = Get-ADComputer -Filter * -Properties Name, Enabled, LastLogonDate, OperatingSystem, DistinguishedName, WhenCreated, SerialNumber |
    Select-Object @{n='ComputerName'; e={ $_.Name }}, @{n='SerialNo'; e={ $_.SerialNumber }}, Enabled, LastLogonDate, WhenCreated, OperatingSystem, DistinguishedName

$Match = Compare-Object -ReferenceObject $AM -DifferenceObject $AD -Property ComputerName -PassThru

##################
#Tab: AM not in AD

$Match |
    Where-Object { $_.SideIndicator -eq '<=' -and -not $_.IsObsolete -and $_.BillingStatus -ne 'Not in contract, to be disposed' -and -not $_.IsInStock -and $_.LocationDetail -notlike '*[Not in AD]*' -and -not $_.IsChangePending } |
    ConvertTo-Csv -Delimiter "`t" -NoTypeInformation |
    Clip
