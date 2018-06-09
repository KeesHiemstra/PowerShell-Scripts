Import-Clixml -Path "\\DEMBMCIS168.corp.demb.com\D$\Scripts\SyncAD2Azure\PrevSKU.xml" |
    Select-Object @{n='AccountSku'; e={ $_.AccountSkuId -replace 'coffeeandtea:' }},
        ActiveUnits, WarningUnits, ConsumedUnits,
        @{n='FreeUnits'; e={ $_.ActiveUnits - $_.ConsumedUnits }} |
    Send-ObjectAsHTMLTableMessage -Subject 'Current Azure number of licenses' -MessageType Information @MailSplatting