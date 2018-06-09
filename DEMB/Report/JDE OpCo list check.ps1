<#
    OpCo list from AD
#>

$ADOpCos = Get-ADUser -Filter * -Properties DepartmentNumber |
    Select-Object @{n='OpCo'; e={ $_.DepartmentNumber -join '' }} |
    Group-Object OpCo |
    Where-Object { $_.Name -ne '' -and $_.Name -match '\d{4}' } |
    Select-Object @{n='OpCo'; e={ $_.Name }}, Count |
    Sort-Object OpCo

$AMOpCos = Get-Content -Path B:\ITAMOpCo.txt

foreach ( $Item in $ADOpCos )
{
    if ( $Item.OpCo -notin $AMOpCos  )
    {
        $User = Get-ADUser -Filter " Company -like '$($Item.OpCo)*' " -Properties Company, c | Select -First 1
        Write-Host "$($Item.OpCo): $($User.C), $($User.Company)"
        
    }
}