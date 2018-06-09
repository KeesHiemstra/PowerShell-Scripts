$ComputerName = Get-Content -Path 'B:\Jobs\SCCM\Rescan.txt'

for ($i = 1; $i -lt 5; $i++)
{ 
    foreach ( $Item in $ComputerName )
    {
        if ( -not (Test-Path -Path "B:\Jobs\SCCM\Rescan\$Item.txt" ) -and (Test-Connection -ComputerName $Item -Quiet) )
        {
            try
            {
                Write-Host "($i) Trying $Item..."
                ([wmi]"\\$Item\root\ccm\invagt:InventoryActionStatus.InventoryActionID='{00000000-0000-0000-0000-000000000001}'").Delete() 
                ([wmi]"\\$Item\root\ccm\invagt:InventoryActionStatus.InventoryActionID='{00000000-0000-0000-0000-000000000002}'").Delete() 
    
                Invoke-WmiMethod -ComputerName $Item -Namespace root\CCM -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000001}"
                Start-Sleep -Seconds 1
                Invoke-WmiMethod -ComputerName $Item -Namespace root\CCM -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000002}"

                New-Item -Path "B:\Jobs\SCCM\Rescan\$Item.txt" -ItemType File | Out-Null
            }
            catch
            {
            }
        }
    }
}
