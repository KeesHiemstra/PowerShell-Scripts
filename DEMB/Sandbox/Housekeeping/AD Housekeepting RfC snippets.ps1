$CheckDate = (Get-Date).AddDays(-30)

$List = Get-ADComputer -Filter { WhenCreated -lt $CheckDate -and WhenChanged -lt $CheckDate -and (LastLogonDate -notlike '*' -or LastLogonDate -le $CheckDate) } -Properties Name, SerialNumber, Enabled, LastLogonDate, WhenCreated, WhenChanged, Street, Description -SearchBase 'OU=Win7,OU=Managed Computers,DC=corp,DC=demb,DC=com'

$list | Select-Object Name, @{n='SerialNo'; e={ $_.SerialNumber -join ';' }}, Enabled, LastLogonDate, @{n='LL_Days'; e={ [int]((Get-Date) - $_.LastLogonDate).Days }}, WhenCreated, WhenChanged, @{n='WC_Days'; e={ [int]((Get-Date) - $_.WhenChanged).Days }}, Street, Description, Info | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation | Clip
$List.Count

$Check = $list |
    Select-Object Name, @{n='SerialNo'; e={ $_.SerialNumber -join ';' }}, 
        Enabled,
        LastLogonDate,
        @{n='LL_Days'; e={ [int]((Get-Date) - $_.LastLogonDate).Days }},
        WhenCreated,
        WhenChanged,
        @{n='WC_Days'; e={ [int]((Get-Date) - $_.WhenChanged).Days }},
        Street,
        Description,
        @{n='Category'; e={ 'No asset info' }},
        @{n='Model'; e={ $null }},
        @{n='Acquisition'; e={ $null }},
        @{n='AssetStatus'; e={ $null }},
        @{n='BillingStatus'; e={ $null }},
        @{n='AssetOpCoFull'; e={ $null }},
        @{n='DomainAccount'; e={ $null }},
        @{n='UserEMail'; e={ $null }},
        @{n='LocationCountryCode'; e={ $null }},
        @{n='LocationName'; e={ $null }},
        @{n='LocationDetail'; e={ $null }},
        @{n='IsDesktop'; e={ $null }},
        @{n='IsInStock'; e={ $null }},
        @{n='IsObsolete'; e={ $null }},
        @{n='IsDummy'; e={ $null }},
        @{n='MonthToRefresh'; e={ $null }},
        @{n='IsChangePending'; e={ $null }},
        @{n='DaysFromLastMutation'; e={ $null }},
        @{n='Action'; e={ $null }}

foreach ( $Item in $Check )
{
    $Asset = Get-AMAsset -ComputerName $Item.Name
    
    if ( $Asset -ne $null )
    {
        $Item.Category = $Asset.Category
        $Item.Model = $Asset.Model
        $Item.Acquisition = $Asset.Acquisition
        $Item.AssetStatus = $Asset.AssetStatus
        $Item.BillingStatus = $Asset.BillingStatus
        $Item.AssetOpCoFull = $Asset.AssetOpCoFull
        $Item.DomainAccount = $Asset.DomainAccount
        $Item.UserEMail = $Asset.UserEMail
        $Item.LocationCountryCode = $Asset.LocationCountryCode
        $Item.LocationName = $Asset.LocationName
        $Item.LocationDetail = $Asset.LocationDetail
        $Item.IsDesktop = $Asset.IsDesktop
        $Item.IsInStock = $Asset.IsInStock
        $Item.IsObsolete = $Asset.IsObsolete
        $Item.IsDummy = $Asset.IsDummy
        $Item.MonthToRefresh = $Asset.MonthToRefresh
        $Item.IsChangePending = $Asset.IsChangePending
        $Item.DaysFromLastMutation = $Asset.DaysFromLastMutation
    }
}


$Check | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation | Clip
