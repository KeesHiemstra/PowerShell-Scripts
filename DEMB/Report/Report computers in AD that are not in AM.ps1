$AM = Get-AMAsset -IsDesktop | Where-Object { $_.Category -ne 'Tablet' -and $_.Category -ne 'Thin client' }

$AD = Get-ADComputer -Filter * -Properties Name, Enabled, LastLogonDate, OperatingSystem, DistinguishedName, WhenCreated, SerialNumber |
    Select-Object @{n='ComputerName'; e={ $_.Name }}, @{n='SerialNo'; e={ $_.SerialNumber }}, Enabled, LastLogonDate, WhenCreated, OperatingSystem, Description, Street, DistinguishedName

$Match = Compare-Object -ReferenceObject $AM -DifferenceObject $AD -Property ComputerName -PassThru

##################
#Tab: AD not in AM

$Match |
    Where-Object { $_.SideIndicator -eq '=>' -and $_.DistinguishedName -notlike '*,CN=Computers,DC=corp,DC=demb,DC=com' -and $_.DistinguishedName -notlike '*,OU=Servers,DC=corp,DC=demb,DC=com' -and $_.DistinguishedName -notlike '*,OU=Domain Controllers,DC=corp,DC=demb,DC=com' -and $_.WhenCreated -lt (Get-Date).AddDays(-5)} |
    ConvertTo-Csv -Delimiter "`t" -NoTypeInformation |
    Clip
    

