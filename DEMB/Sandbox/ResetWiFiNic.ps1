if ( Test-IsElevated )
{
    $WMI = Get-WmiObject -Class Win32_NetworkAdapter | 
        Where-Object { $_.NetEnabled -ne $null -and $_.ProductName -eq 'Intel(R) Dual Band Wireless-N 7265' }

    $WMI.Disable() | Out-Null
    $WMI.Enable() | Out-Null
}
