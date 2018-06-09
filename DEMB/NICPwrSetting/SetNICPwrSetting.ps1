<#
    SetNICPowerSetting

    Action in this script:
    - Install Update KB3050265 for performance and against BSoD.
    - Disable the Microsoft Virtual WiFi Miniport Adapter against BSoD
    - Switch off the powers setting of the network interface cards on the HP EliteBook 820 G2 and HP EliteBook 840 G2 against BSoD.

    The scripts needs to start in elivated rights.

    A stop file is created if the model is not in scope, or no changes are needed (any more). The script will not continue at the start
    if the stop file is present.

    A reboot is needed to have the settings active, this will not be forces by the script. Once rebooted, the script will
    check again if changes are needed and it will create a stop file if the settings are done.

    The scripts is compatible with PowerShell version 2.0
#>
$LogFile = 'C:\Windows\Temp\SetNICPwrSetting.log'
$StopFilePwrSet = 'C:\Windows\Temp\SetNICPwrSetting.txt'
$StopFileVWifi = 'C:\Windows\Temp\SetNICMsVWifiOff.txt'
$StopFile3050265 = 'C:\Windows\Temp\InstallKB3050265.txt'
$StopFile3102810 = 'C:\Windows\Temp\InstallKB3102810.txt'

$RegKeyNICs = 'SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002BE10318}'
$AllowNICs = ('Intel(R) Dual Band Wireless-N 7265', 'Intel(R) Ethernet Connection (3) I218-LM')

New-Variable -Name CountChanges -Scope Script -Value 0 -Force

function Write-Log([string]$Message)
{
    $LogMessage = "{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Add-Content -Path $LogFile -Value $LogMessage
}

function Write-Stop([string]$Message, [string]$FileName)
{
    Write-Log -Message $Message
    Add-Content -Path $FileName -Value $Message
}

Write-Log -Message "Starting script"
Write-Log -Message "ComputerName: $($env:COMPUTERNAME) // UserName: $($env:USERNAME)"
Write-Log -Message "Last booted at $([management.managementDateTimeConverter]::ToDateTime((Get-WmiObject -Class 'Win32_OperatingSystem').LastBootUpTime).ToString('yyyy-MM-dd HH:mm:ss'))"

#if (-not (Test-Path -Path $StopFile3102810))
#{
#    if ((Test-Path -Path "C:\Windows\Temp\Windows6.1-KB3102810-x64.msu"))
#    {
#        C:\Windows\Temp\Windows6.1-KB3102810-x64.msu /Quiet /NoRestart
#        Write-Stop "KB3102810 installer has been launched" -FileName $StopFile3102810
#    }
#    else
#    {
#        Write-Log -Message "Installation file not found (C:\Windows\Temp\Windows6.1-KB3102810-x64.msu)"
#    }
#}

if (-not (Test-Path -Path $StopFileVWifi))
{
    try
    {
        if (Get-WmiObject -Class Win32_NetworkAdapter -Namespace root\CIMV2 -Filter "Name='Microsoft Virtual WiFi Miniport Adapter'")
        {
            (Get-WmiObject -Class Win32_NetworkAdapter -Namespace root\CIMV2 -Filter "Name='Microsoft Virtual WiFi Miniport Adapter'").Disable()
            Write-Stop "The Microsoft Virtual WiFi Miniport Adapter has been disabled" -FileName $StopFileVWifi
        }
        else
        {
            Write-Stop "The Microsoft Virtual WiFi Miniport Adapter is not installed" -FileName $StopFileVWifi
        }
    }
    catch
    {
        Write-Log -Message "Error disabling "The Microsoft Virtual WiFi Miniport Adapter has been disabled""
    }
}

if ((Test-Path -Path $StopFilePwrSet))
{
    Write-Log -Message "Nothing to do here"
    break
}

$Model = (Get-WmiObject -Class Win32_ComputerSystem).Model
Write-Log -Message "Model: $Model"

if ($Model -notmatch 'HP EliteBook 8\d0 G2')
{
    Write-Stop -Message "Model is not in scope" -FileName $StopFilePwrSet
    break
}

Write-Log -Message "Processing all NICs"

$NICs = Get-WmiObject -Class Win32_NetworkAdapter |
    Where-Object { $_.NetEnabled -ne $null } |
    Select-Object Manufacturer, ProductName, Index, NetEnabled, NetConnectionStatus

foreach ($NIC in $NICs)
{
    if ($AllowNICs -contains $NIC.ProductName)
    {
        $IndexKey = $NIC.Index.ToString().PadLeft(4, '0')

        Push-Location HKLM:
        if ((Get-ItemProperty -Path "$RegKeyNICs\$IndexKey" -Name "PnpCapabilities" -ErrorAction SilentlyContinue))
        {
            if ((Get-ItemProperty -Path "$RegKeyNICs\$IndexKey" -Name "PnpCapabilities" -ErrorAction SilentlyContinue).PnpCapabilities -ne 24)
            {
                #Value exists but needs to be changed
                Set-ItemProperty -Path "$RegKeyNICs\$IndexKey" -Name "PnpCapabilities" -Value 24 -Force | Out-Null
                $Script:CountChanges++
                Write-Log -Message "Power settings on $($NIC.ProductName) has been set to off"
            }
            else
            {
                Write-Log -Message "Power settings on $($NIC.ProductName) has already been set to off"
            }
        }
        else
        {
            #Value needs to be created
            New-ItemProperty -Path "$RegKeyNICs\$IndexKey" -Name "PnpCapabilities" -Value 24 -PropertyType "DWord" | Out-Null
            $Script:CountChanges++
            Write-Log -Message "Power settings created in registry on $($NIC.ProductName) and has been set to off"
        }
        Pop-Location

    }
    else
    {
        Write-Log -Message "Skipped NIC $($NIC.ProductName)"
    }
}

if ($CountChanges -eq 0)
{
    Write-Stop -Message "No changes left to set for power settings" -FileName $StopFilePwrSet
}
else
{
    Write-Host "You need now to reboot the computer and run this script again!!!`n`n"
    Read-Host -Prompt "Press enter..."
}

