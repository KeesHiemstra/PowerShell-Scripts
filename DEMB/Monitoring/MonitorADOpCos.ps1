<#
    Monitor AD OpCos
#>

$ADUser = Get-ADUser -Filter { EmployeeID -like '*' -and DepartmentNumber -ne 'EXTERNAL'  -and DepartmentNumber -ne 'IBM'} -Properties DepartmentNumber, Company, C |
    Where-Object { $_.Enabled } |
    Select-Object SAMAccountName, C, @{n='DepartmentNumber'; e={ $_.DepartmentNumber -join ',' }}, Company |
    Where-Object { $_.DepartmentNumber -ne '' }

$ADOpCo = $ADUser |
    Group-Object DepartmentNumber |
    Select-Object @{n='Number'; e={ $_.Name }}, Count, @{n='CountryCode'; e={ $_.Group[0].C} }, @{n='FullName'; e={ $_.Group[0].Company} }, @{n='IsInAM'; e={ $false }}

foreach ( $Item in $ADOpCo )
{
    $AMOpCo = Get-AMOpCo -Number $Item.Number
    if ( $AMOpCo -ne $null )
    {
        $Item.IsInAM = $true
    }
}

$Diff = $ADOpCo | Where-Object { $_.IsInAM -eq $false }
