<#
    SetWiFiEliteBook8x0

    RfC 1286507 - Set WiFi Roaming agressiveness on "Highest" (4) and Preferred band on "5.2GHz" for any HP EliteBook 820 G2 and HP EliteBook 840 G2.

    The location of settings differ in different setup, the exact location needs to be read from WMI.
#>

$RegKeyNICs = 'SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002BE10318}'
$AllowNICs = ('Intel(R) Dual Band Wireless-N 7265')
$LogFile = 'C:\Windows\Temp\SetWiFi.log'

#region Log
function Write-Log([string]$Message)
{
    $LogMessage = "{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Add-Content -Path $LogFile -Value $LogMessage
}

#endregion

Write-Log -Message "Starting script"
Write-Log -Message "ComputerName: $($env:COMPUTERNAME) // UserName: $($env:USERNAME)"
Write-Log -Message "Last booted at $([management.managementDateTimeConverter]::ToDateTime((Get-WmiObject -Class 'Win32_OperatingSystem').LastBootUpTime).ToString('yyyy-MM-dd HH:mm:ss'))"

$Model = (Get-WmiObject -Class Win32_ComputerSystem).Model
if ($Model -notmatch 'HP EliteBook 8\d0 G2')
{
    Write-Log -Message "Model is not in scope ($Model)"
    break
}

#Process all NICs
$NICs = Get-WmiObject -Class Win32_NetworkAdapter |
    Where-Object { $_.NetEnabled -ne $null }

foreach ($NIC in $NICs)
{
    if ($AllowNICs -contains $NIC.ProductName)
    {
        $IndexKey = $NIC.Index.ToString().PadLeft(4, '0')

        $Changed = $false

        Push-Location HKLM:

        #"RoamAggressiveness"="4"
        if ((Get-ItemProperty -Path "$RegKeyNICs\$IndexKey" -Name "RoamAggressiveness" -ErrorAction SilentlyContinue))
        {
            if ((Get-ItemProperty -Path "$RegKeyNICs\$IndexKey" -Name "RoamAggressiveness" -ErrorAction SilentlyContinue).RoamAggressiveness -ne "4")
            {
                #Value exists but needs to be changed
                try
                {
                    Set-ItemProperty -Path "$RegKeyNICs\$IndexKey" -Name "RoamAggressiveness" -Value "4" -Force | Out-Null
                    $Changed = $true
                    Write-Log -Message "RoamAggressiveness on $($NIC.ProductName) has been set to 4"
                }
                catch
                {
                    Write-Log -Message "No access to change RoamAggressiveness on $($NIC.ProductName) to 4"
                }
            }
        }

        #"RoamingPreferredBandType"="2"
        if ((Get-ItemProperty -Path "$RegKeyNICs\$IndexKey" -Name "RoamingPreferredBandType" -ErrorAction SilentlyContinue))
        {
            if ((Get-ItemProperty -Path "$RegKeyNICs\$IndexKey" -Name "RoamingPreferredBandType" -ErrorAction SilentlyContinue).RoamingPreferredBandType -ne "2")
            {
                #Value exists but needs to be changed
                try
                {
                    Set-ItemProperty -Path "$RegKeyNICs\$IndexKey" -Name "RoamingPreferredBandType" -Value "2" -Force | Out-Null
                    $Changed = $true
                    Write-Log -Message "RoamingPreferredBandType on $($NIC.ProductName) has been set to 2"
                }
                catch
                {
                    Write-Log -Message "No access to change RoamingPreferredBandType on $($NIC.ProductName) to 2"
                }
            }
        }

        Pop-Location

        if ( $changed )
        {
            try
            {
                $Nic.Disable() | Out-Null
                $Nic.Enable() | Out-Null
                Write-Log -Message "$($NIC.ProductName) adaptor has been reset"
            }
            catch
            {
                Write-Log -Message "No access to reset the $($NIC.ProductName) adaptor or reset failed"
            }
        }
        else
        {
            Write-Log -Message "Nothing to change"
        }
    }
}

Write-Log "Script completed"