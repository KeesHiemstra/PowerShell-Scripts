$Props = @{'SKU'=''; }

$AllSKUs = Get-MsolAccountSku

$Licenses = @()

foreach ($SKU in $AllSKUs)
{
    $AllServicePlans = $SKU.ServiceStatus.ServicePlan
    foreach ($ServicePlan in $AllServicePlans)
    {
        $Obj = New-Object -TypeName 'PSObject' -Property ([ordered]@{'Name'        = [string]$SKU.AccountSkuId -replace 'coffeeandtea:';
                                                                     'ServiceName' = [string]$ServicePlan.ServiceName;
                                                                     'ServiceType' = [string]$ServicePlan.ServiceType;
                                                                     'TargetClass' = [string]$ServicePlan.TargetClass})
        $Licenses += $Obj
    }
}

$Licenses | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation | Clip

