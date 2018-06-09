$Error.Clear()

Set-Location "C:\Src\PowerShell\4.0\SOE Tools"

$FileName = "D:\Data\Jobs\MSOL Data\MSOL-$((Get-Date).ToString("yyyy-MM-dd_HHmm")).csv"

$UserName = "Kees.Hiemstra@JDEcoffee.com"
$UserNameFile = "$($Profile.CurrentUserAllHosts -replace "profile.ps1$")$($UserName -replace "^.*\\").txt"

$Encoding = if ( (Get-Date).Hour -lt 11 ) { 'ASCII' } else { 'Unicode' }

#Initiate connection to Azure only once
try {$Sku = Get-MsolAccountSku -ErrorAction Stop }
catch
{
    $Error.RemoveAt(0)
    #Get stored credentials
    if($Cred -eq $null)
    {
        if(-Not (Test-Path -Path $UserNameFile))
        {
            $Cred = Get-Credential -UserName $UserName -Message "Provide password"
            $Cred.Password | ConvertFrom-SecureString | Out-File $UserNameFile
        }
        else
        {
            $Cred = New-Object -Type System.Management.Automation.PSCredential -ArgumentList $UserName, (Get-Content -Path $UserNameFile | ConvertTo-SecureString)
        }
    }
    try { Connect-MsolService -Credential $Cred } catch { throw }
    $Sku = Get-MsolAccountSku

    foreach ($Unit in $Sku)
    {
        $Unit | Select-Object @{n='Time'; e={ (Get-Date).ToString('yyyy-MM-dd HH:mm') }}, AccountSkuId, TargetClass, ActiveUnits, ConsumedUnits, WarningUnits, LockedOutUnits, SuspendedUnits, @{n='RemainingUnits'; e={ ($_.ActiveUnits + $_.WarningUnits - $_.ConsumedUnits) }} |
            Export-Csv -Path "D:\Data\Jobs\MSOL Data\_$($Unit.SkuPartNumber).csv" -Append -NoTypeInformation

        if (($Unit.ActiveUnits + $Unit.WarningUnits - $Unit.ConsumedUnits) -le (($Unit.ActiveUnits + $Unit.WarningUnits) / 100))
        {
            Send-MailMessage -SmtpServer "smtp.corp.demb.com" -From "Kees.Hiemstra@JDECoffee.com" -To "Kees.Hiemstra@hpe.com" -Subject "Azure licenses are getting low" -Body "Hi Kees,`n`nThe Azure licenses are getting low ($(($Unit.ActiveUnits + $Unit.WarningUnits - $Unit.ConsumedUnits))) for $($Unit.SkuPartNumber).`n`nKind regards,`nKees..."
        }
    }
}

$ADUsers = Get-ADUser -Filter * -Properties sAMAccountName, userPrincipalName, givenName, surName, displayName, c, l, mail, mobile, Enabled, PasswordExpired, passwordLastSet, lastLogonDate, whenCreated, whenChanged, company, employeeID, employeeType, comment, extensionAttribute5, extensionAttribute2, extensionAttribute11, manager
$Duplicates = $ADUsers | Group-Object userPrincipalName | Where-Object { $_.Name -ne '' -and $_.Count -gt 1 }
if ($Duplicates.Count -gt 0)
{
    $HTML = $Duplicates.Group | Select-Object sAMAccountName, userPrincipalName | ConvertTo-Html -Title "Duplicate userPrincipalNames"| Out-String

    Send-MailMessage -SmtpServer 'smtp.corp.demb.com' -From 'Kees.Hiemstra@JDECoffee.com' -To 'Kees.Hiemstra@hpe.com' -Subject "Duplicate userPrincipalNames" -Body $HTML -BodyAsHtml
}
$Duplicates = $null

[Collections.Generic.List[Object]]$AzUsers = Get-MsolUser -All

foreach($ADUser in $ADUsers)
{
    Add-Member -InputObject $ADUser -TypeName Bool -NotePropertyName BlockCredential -NotePropertyValue $true -Force
    Add-Member -InputObject $ADUser -TypeName Bool -NotePropertyName IsInAD -NotePropertyValue $true -Force
    Add-Member -InputObject $ADUser -TypeName Bool -NotePropertyName IsInAzure -NotePropertyValue $false -Force
    Add-Member -InputObject $ADUser -TypeName Date -NotePropertyname whenCreatedInAzure -NotePropertyValue "--" -Force
    Add-Member -InputObject $ADUser -TypeName String -NotePropertyname usageLocation -NotePropertyValue "--" -Force
    Add-Member -InputObject $ADUser -TypeName String -NotePropertyName LicensesText -NotePropertyValue "" -Force
    Add-Member -InputObject $ADUser -TypeName String -NotePropertyName GranularNotation -NotePropertyValue "" -Force

    if (-not [string]::IsNullOrEmpty($ADUser.userPrincipalName))
    {
        $Index = $AzUsers.FindIndex( { $args[0].userPrincipalName -eq $ADUser.userPrincipalName } )

        if ($Index -ne -1)
        {
            #User found in Azure
            $ADUser.BlockCredential = $AzUsers[$Index].BlockCredential
            $ADUser.IsInAzure = $true
            $ADUser.whenCreatedInAzure = $AzUsers[$Index].whenCreated
            $ADUser.usageLocation = $AzUsers[$Index].usageLocation

            if ($AzUsers[$Index].Licenses.Count -gt 0)
            {
                $Licenses = ConvertFrom-AzureLicense -Licenses $AzUsers[$Index].Licenses
                $ADUser.LicensesText = $Licenses.LicensesText
                $ADUser.GranularNotation = $Licenses.GranularNotation
            }#Has license

            $AzUsers.RemoveAt($Index)
        }#Found in Azure
    }#Has userPrincipalName
}#Foreach ADUsers

$ADUsers += $AzUsers | 
    Select-Object @{n='samAccountName'; e={'n/a'}}, 
        userPrincipalName, 
        @{n='givenName'; e={$_.FirstName}}, 
        @{n='surName'; e={$_.LastName}},
        displayName,
        #@{n='c'; e={''}}, 
        @{n='l'; e={$_.City}}, 
        #@{n='mail'; e={''}},
        #@{n='mobile'; e={''}}, 
        #@{n='Enabled'; e={''}}, 
        #@{n='PasswordExpired'; e={''}},
        @{n='passwordLastSet'; e={$_.LastPasswordChangeTimestamp}}, 
        #@{n='lastLogonDate'; e={''}}, 
        #whenCreated,
        #@{n='whenChanged'; e={''}}, 
        #@{n='company'; e={''}}, 
        #@{n='employeeID'; e={''}}, 
        #@{n='employeeType'; e={''}}, 
        #@{n='comment'; e={''}}, 
        #@{n='extensionAttribute5'; e={''}}, 
        #@{n='extensionAttribute2'; e={''}}, 
        @{n='IsInAD'; e={$false}}, 
        @{n='IsInAzure'; e={$true}}, 
        @{n='whenCreatedInAzure'; e={$_.whenCreated}}, 
        usageLocation,
        @{n='LicensesText'; e={if($_.Licenses -ne $null){try{(ConvertFrom-AzureLicense -Licenses $_.Licenses).LicensesText}catch{''}}}}, #LicensesText
        @{n='GranularNotation'; e={if($_.Licenses -ne $null){try{(ConvertFrom-AzureLicense -Licenses $_.Licenses).GranularNotation}catch{''}}}}, #GranularNotation
        BlockCredential

$ADUsers |
    Select-Object sAMAccountName, 
        userPrincipalName, 
        givenName, 
        surName,
        displayName,
        c, 
        l, 
        mail, 
        mobile,
        manager,
        Enabled, 
        PasswordExpired, 
        passwordLastSet, 
        lastLogonDate, 
        whenCreated, 
        whenChanged, 
        company, 
        employeeID, 
        employeeType, 
        comment, 
        extensionAttribute5, 
        extensionAttribute2,
        extensionAttribute11,
        IsInAD,
        IsInAzure,
        whenCreatedInAzure,
        usageLocation, 
        LicensesText,
        GranularNotation,
        BlockCredential |
    Export-Csv -Path $FileName -NoTypeInformation -Encoding $Encoding

if ((Get-Date).Hour -lt 11)
{
    Send-MailMessage -SmtpServer "smtp.corp.demb.com" -From "Kees.Hiemstra@JDECoffee.com" -To "Richard.Wesseling@hpe.com" -Subject "MSOnline licenses dump" -Body "Hoi Richard`n`nHier is de dumb van vandaag.`n`nGroetjes,`nKees..." -Attachments $FileName
}

$Message = ""
if ($Error.Count -gt 0)
{
    $Message = "The following errors occurred:`n$($Error | Out-String)`n`n"
}

Send-MailMessage -SmtpServer 'smtp.corp.demb.com' -From 'Kees.Hiemstra@JDECoffee.com' -To 'Kees.Hiemstra@hpe.com' -Subject "MSOnline licenses dump" -Body "Hi Kees,`n`nThe MSOL dumb is done.`n`n$($Message)Kind regards,`nKees..." #-Attachments $FileName
