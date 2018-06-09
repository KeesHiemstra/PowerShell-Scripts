<#


Todo:
- Create .OptOut files
- Determine path for .OptOut files
- Add logging
- Create and send report on new licenses and options
- Set variable external for configuration
- Check all testing line and remove them
- Document solution
#>

#region Settings

$LicensePath = "C:\Src\PowerShell\3.0\!DEMB\SyncADToCloud\AzureLicenses.csv"

$UserName = "UserCreation@coffeeandtea.onmicrosoft.com"
$UserNameFile = "$PSScriptRoot\$UserName.txt"

#region Set date/time format
$CurrentThread = [System.Threading.Thread]::CurrentThread
$Culture = [CultureInfo]::InvariantCulture.Clone()
$Culture.DateTimeFormat.ShortDatePattern = 'yyyy-MM-dd'
$Culture.DateTimeFormat.ShortTimePattern = 'HH:mm:ss'
$CurrentThread.CurrentCulture = $Culture
$CurrentThread.CurrentUICulture = $Culture
#endregion

if ( $true )
{
    #Get stored credentials
    if($AzCred -eq $null)
    {
        if( -Not (Test-Path -Path $UserNameFile) )
        {
            $AzCred = Get-Credential -UserName $UserName -Message "Provide password"
            $AzCred.Password | ConvertFrom-SecureString | Out-File $UserNameFile
        }
        else
        {
            $AzCred = New-Object -Type System.Management.Automation.PSCredential -ArgumentList $UserName, (Get-Content -Path $UserNameFile | ConvertTo-SecureString)
        }
    }

    try
    {
        Connect-MsolService -Credential $AzCred -ErrorAction Stop
    }
    catch
    {
        Write-Break -Message "The connection to the cloud failed"
    }
}#-not RunOffline
#endregion

#region Import existing licenses

if ( Test-Path -Path $LicensePath )
{
    #Import csv and convert boolean directly
    $CurrentLicenses = Import-Csv -Path $LicensePath -Delimiter ',' |
        Select-Object TenantName,
            TenantGUID,
            AccountSkuId,
            @{n='Parent';e={ [bool]::Parse($_.Parent) }},
            ServiceName,
            ServiceType,
            TargetClass,
            FirstSeen,
            LastSeen,
            LicenseTag,
            Description,
            @{n='Configurable'; e={ [bool]::Parse($_.Configurable) }},
            @{n='AllwaysOff'; e={ [bool]::Parse($_.AllwaysOff) }},
            @{n='IsNew'; e={ [bool]::Parse($_.IsNew) }},
            @{n='DoesExist'; e={ [bool]::Parse($_.DoesExist) }}
}
else
{
    $CurrentLicenses = @()
}

#endregion

#region Get SKU
$Sku = $null
try
{
    $Sku = Get-MsolAccountSku -ErrorAction Stop
}
catch
{
    Write-Error "Can't get SKU"
}

if ( $Sku -eq $null )
{
    Write-Error "Can't get SKU"
}
#endregion

#region Serialize SKU
$MSSku = @()

$Property = [Ordered] @{'TenantName'   = [string]   ''
                        'TenantGUID'   = [string]   ''
                        'AccountSkuId' = [string]   ''
                        'Parent'       = [bool]     $true
                        'ServiceName'  = [string]   ''
                        'ServiceType'  = [string]   ''
                        'TargetClass'  = [string]   ''
                        'FirstSeen'    = [datetime] (Get-Date)
                        'LastSeen'     = [datetime] (Get-Date)
                        'LicenseTag'   = [string]   ''
                        'Description'  = [string]   ''
                        'Configurable' = [bool]     $false
                        'AllwaysOff'   = [bool]     $true
                        'IsNew'        = [Bool]     $true
                        'DoesExist'    = [bool]     $true
                       }

foreach ( $Item in $Sku )
{
    $Obj = New-Object -TypeName PSObject -Property $Property
    $Obj.TenantName = $Item.AccountName
    $Obj.TenantGUID = $Item.AccountObjectId
    $Obj.AccountSkuId = $Item.AccountSkuId
    $Obj.Parent = $true

    $MSSku += $Obj

    foreach ( $Option in $Item.ServiceStatus.ServicePlan )
    {
        $Obj = New-Object -TypeName PSObject -Property $Property
        $Obj.TenantName = $Item.AccountName
        $Obj.TenantGUID = $Item.AccountObjectId
        $Obj.AccountSkuId = $Item.AccountSkuId
        $Obj.Parent = $false
        $Obj.ServiceName = $Option.ServiceName
        $Obj.ServiceType = $Option.ServiceType
        $Obj.TargetClass = $Option.TargetClass

        $MSSku += $Obj
    }
}

#endregion

#region Compare current with previous

foreach ( $Item in $MSSku )
{
    $CheckLicense = $CurrentLicenses | Where-Object { $_.AccountSkuId -eq $Item.AccountSkuId -and $_.ServiceName -eq $Item.ServiceName }
    if ( $CheckLicense -ne $null )
    {
        #Write-Host "Update $($Item.AccountSkuId)::$($Item.ServiceName)"
        $Item.FirstSeen    = $CheckLicense.FirstSeen
        $Item.LicenseTag   = $CheckLicense.LicenseTag
        $Item.Description  = $CheckLicense.Description
        $Item.Configurable = $CheckLicense.Configurable
        $Item.AllwaysOff   = $CheckLicense.AllwaysOff
        $Item.IsNew        = $false
    }
    else
    {
        #Write-Host "New $($Item.AccountSkuId)::$($Item.ServiceName)"
    }
}

foreach ( $CheckLicense in $CurrentLicenses )
{
    if ( $Item.Parent )
    {
        $Item = $MSSku | Where-Object { $_.AccountSkuId -eq $CheckLicense.AccountSkuId -and $_.ServiceName -eq $null }
    }
    else
    {
        $Item = $MSSku | Where-Object { $_.AccountSkuId -eq $CheckLicense.AccountSkuId -and $_.ServiceName -eq $CheckLicense.ServiceName }
    }

    if ( $Item -eq $null )
    {
        $CheckLicense.DoesExist = $false
        $MSSku += $Item
    }
}

#endregion

#region Report

$NewParent = $MSSku | Where-Object { $_.IsNew -and $_.Parent }


$NewOption = $MSSku | Where-Object { $_.IsNew -and $_.AccountSkuId -notin $NewParent.AccountSkuId }

#endregion

#$MSSku | Format-Table * -AutoSize
$MSSku | Export-Csv -Path B:\AzureLicenses.csv -Delimiter "," -NoTypeInformation

#region Create SkuId exclusion files
cls
foreach ( $Parent in ($MSSku | Where-Object { $_.Parent }) )
{
    $Configurable = $false
    foreach ( $Item in ($MSSku | Where-Object { $_.AccountSkuId -eq $Parent.AccountSkuId }) )
    {
        if ( $Item.Configurable )
        {
            $Configurable = $true
        }
    }

    if ( $Configurable )
    {
        $SkuId = $Parent.AccountSkuId -replace $Parent.TenantName
        $ServiceName = $MSSku | Where-Object { $_.AccountSkuId -eq $Parent.AccountSkuId -and -not $_.Parent -and $_.AllwaysOff }
        $ServiceName | Ft AccountSkuId, Parent, ServiceName, TargetClass, Configurable, IsAllwaysOff, IsNew -AutoSize
    }
}
#endregion

#region Create investigation
if ( $false )
{
    #region Read Azure data
    if ( $AllAzUsers -eq $null )
    {
        if ( Test-Path -Path "C:\Src\PowerShell\3.0\!DEMB\SyncADToCloud\AllAzUsers.xml" )
        {
            $AllAzUsers = Import-Clixml -Path "C:\Src\PowerShell\3.0\!DEMB\SyncADToCloud\AllAzUsers.xml"
        }
        else
        {
            try
            {
            #    Write-Log -Message "Reading Azure"
                $AllAzUsers = Get-MsolUser -All -ErrorAction Stop |
                    Where-Object { $_.Licenses.Count -gt 0 } |
                    Select-Object Licenses

            #    Write-Log -Message "Finished reading Azure"
            #    Write-Log -Message "Number of Azure accounts: $($AllAzUsers.Count)"
                $AllAzUsers | Export-Clixml -Path "C:\Src\PowerShell\3.0\!DEMB\SyncADToCloud\AllAzUsers.xml"
            }
            catch
            {
                Write-Host "Connection failed"
            #    Write-Break -Message "Error reading from Azure"
            }
        }
    }
    #endregion

    #region Find differences between AD and Azure
    if ( $AllAzUsers -eq $null )
    {
    #    Write-Break -Message "No data read from Azure"
    }

    #ii "B:\AzureLicenses.csv"

    foreach ( $Item in $MSSku )
    {
        Add-Member -InputObject $Item -MemberType NoteProperty -Name 'Total' -Value ($AllAzUsers | Where-Object { $_.Licenses.AccountSkuId -eq $Item.Licenses.AccountSkuId }).Count
        Add-Member -InputObject $Item -MemberType NoteProperty -Name 'Success' -Value 0
        Add-Member -InputObject $Item -MemberType NoteProperty -Name 'PendingActivation' -Value 0
        Add-Member -InputObject $Item -MemberType NoteProperty -Name 'PendingInput' -Value 0
        Add-Member -InputObject $Item -MemberType NoteProperty -Name 'Disabled' -Value 0
        Add-Member -InputObject $Item -MemberType NoteProperty -Name 'Error' -Value 0
    }

    #($AllAzUsers | Where-Object { $_.Licenses.AccountSkuId -eq 'coffeeandtea:ENTERPRISEPACK' -and $_.Licenses.ServiceStatus.ServicePlan.ServiceName -eq 'YAMMER_ENTERPRISE' } | Where-Object { $_.Licenses.ServiceStatus.ProvisioningStatus -eq 'Disabled' }).Count

    foreach ( $User in $AllAzUsers )
    {
        foreach ( $CheckLicense in $User.Licenses )
        {
            foreach ( $Option in $CheckLicense.ServiceStatus )
            {
                switch ( $Option.ProvisioningStatus )
                {
                    'Success'           { ($MSSku | Where-Object { $_.AccountSkuId -eq $CheckLicense.AccountSkuId -and $_.ServiceName -eq $Option.ServicePlan.ServiceName }).Success++ }
                    'PendingActivation' { ($MSSku | Where-Object { $_.AccountSkuId -eq $CheckLicense.AccountSkuId -and $_.ServiceName -eq $Option.ServicePlan.ServiceName }).PendingActivation++ }
                    'PendingInput'      { ($MSSku | Where-Object { $_.AccountSkuId -eq $CheckLicense.AccountSkuId -and $_.ServiceName -eq $Option.ServicePlan.ServiceName }).PendingInput++ }
                    'Disabled'          { ($MSSku | Where-Object { $_.AccountSkuId -eq $CheckLicense.AccountSkuId -and $_.ServiceName -eq $Option.ServicePlan.ServiceName }).Disabled++ }
                    'Error'             { ($MSSku | Where-Object { $_.AccountSkuId -eq $CheckLicense.AccountSkuId -and $_.ServiceName -eq $Option.ServicePlan.ServiceName }).Error++ }
                }
            }
        }
    }
    #endregion

    $MSSku | Export-Csv -Path B:\AzureLicensesInvestigation.csv -Delimiter "," -NoTypeInformation
    ii B:\AzureLicensesInvestigation.csv
}
#endregion
