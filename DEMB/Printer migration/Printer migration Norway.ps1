<#
    Printer migration Norway.ps1

    === Version history
    Version 1.10 (2017-01-20, Kees Hiemstra)
    - Added logging.
    Version 1.01 (2017-01-04, Kees Hiemstra)
    - Updated printer list.
    Version 1.00 (2016-12-13, Kees Hiemstra)
    - Initial version.
#>

$LogPath = 'C:\HP\Logs\PrinterMigration.log'

function Write-Log ([string]$Message)
{
    Add-Content -Path $LogPath -Value "$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) $Message"
}

#region Get-PrinterPortName

<#
.SYNOPSIS
    Get printer port name based on the host address.
.DESCRIPTION
    Get printer port name based on the host address (IP address or DNS name).
.EXAMPLE
    New-PrinterPort -PrinterPort 'PrtRoundTable.Camelot.ay'
.OUTPUTS
    [string]
#>
function Get-PrinterPortName
{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        # Printer port host address (IP Address or DNS name)
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]
        $HostAddress
    )

    Process
    {
        (Get-WmiObject -Class Win32_TcpIPPrinterPort | Where-Object { $_.HostAddress -eq $HostAddress }).Name
    }
}

#endregion


#region New-PrinterPort

<#
.SYNOPSIS
    Create new printer port.
.DESCRIPTION
    Create new printer port on the local computer using WMI. The function will return a [System.Management.ManagementPath] object when the port is created.
    There is no output if the port already exist.
.EXAMPLE
    New-PrinterPort -PrinterPort 'PrtRoundTable.Camelot.ay'
.OUTPUTS
    [System.Management.ManagementPath]
#>
function New-PrinterPort
{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        # New printer port (IP Address or DNS name)
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]
        $PrinterPort
    )

    Process
    {
        if ( (Get-WmiObject -Class Win32_TcpIPPrinterPort | Where-Object { $_.HostAddress -eq $PrinterPort } ) -eq $null ) 
        {
            $Port = [wmiclass]"\\.\ROOT\cimv2:Win32_TcpIPPrinterPort"
            $Port.PSBase.Scope.Options.EnablePrivileges = $true

            $NewPort = $Port.CreateInstance()
            $NewPort.HostAddress = $PrinterPort
            $NewPort.Name = $PrinterPort
            $NewPort.PortNumber = '9100'
            $NewPort.Protocol = 1
            $NewPort.SnmpEnabled = $false
            $NewPort.Put()
            Write-Log "Created printer port $PrinterPort"
            Write-Verbose "Created printer port $PrinterPort"
        }
    }
}

#endregion


#region Set-PrinterPrinterPort

<#
.SYNOPSIS
    Set new printer port on existing printer based on the old printer port name (IP address or DNSName).
.DESCRIPTION
    Set new printer port on the printer with currently uses the existing printer port on the local computer using WMI.
    This function requires elevated rights.
.EXAMPLE
    Set-PrinterPrinterPort -PrinterPort '10.128.76.9' -NewPrinterPort 'PrtRoundTable.Camelot.ay' -Verbose
.OUTPUTS
    None
#>
function Set-PrinterPrinterPort
{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        # Existing printer port (IP Address or DNS name)
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]
        $PrinterPort,

        # New printer port (IP Address or DNS name)
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]
        $NewPrinterPort
    )

    Process
    {
        foreach ( $PortName in (Get-PrinterPortName -HostAddress $PrinterPort) )
        {
            if ( $Printer = Get-WmiObject -Class Win32_Printer -Property * | Where-Object { $_.PortName -eq $PortName } )
            {
                New-PrinterPort $NewPrinterPort
                foreach ( $Item in $Printer )
                {
                    try
                    {
                        $Item.PortName = $NewPrinterPort
                        $Item.Put()
                        Write-Log "Changed printer port on $($Item.Name) from $PrinterPort to $NewPrinterPort"
                        Write-Verbose "Changed printer port on $($Item.Name) from $PrinterPort to $NewPrinterPort"
                    }
                    catch
                    {
                        Write-Log "Error update printer port: $($Error[0])"
                        Write-Error "$($Error[0])"
                    }
                }
            }
        }
    }
}

#endregion

if ( -not (Test-Path -Path $LogPath) )
{
    Get-WmiObject -Class Win32_Printer -Property * | ForEach-Object { Write-Log "Printer: $($_.Name) => $($_.PortName)" }
    Get-WmiObject -Class Win32_TcpIPPrinterPort | ForEach-Object { Write-Log "PrinterPort: $($_.Name) => $($_.HostAddress)" }
}

Set-PrinterPrinterPort -PrinterPort '172.28.1.82' -NewPrinterPort 'M3BGOKET01' -Verbose
Set-PrinterPrinterPort -PrinterPort '172.28.1.84' -NewPrinterPort 'M3BGOKET02' -Verbose
Set-PrinterPrinterPort -PrinterPort '172.28.1.83' -NewPrinterPort 'M3BGORFA01' -Verbose
Set-PrinterPrinterPort -PrinterPort '172.28.1.14' -NewPrinterPort 'M3BGOKBI01' -Verbose
Set-PrinterPrinterPort -PrinterPort '172.28.1.13' -NewPrinterPort 'M3BGOKBI02' -Verbose
Set-PrinterPrinterPort -PrinterPort '172.28.1.12' -NewPrinterPort 'M3BGLAOD01' -Verbose
Set-PrinterPrinterPort -PrinterPort '172.28.1.17' -NewPrinterPort 'PRNNOBGO001' -Verbose
Set-PrinterPrinterPort -PrinterPort '172.28.1.31' -NewPrinterPort 'PRNNOBGO003' -Verbose
Set-PrinterPrinterPort -PrinterPort '172.28.1.215' -NewPrinterPort 'PRNNOBGO004' -Verbose
Set-PrinterPrinterPort -PrinterPort '172.28.1.23' -NewPrinterPort 'PRNNOBGO006' -Verbose
Set-PrinterPrinterPort -PrinterPort '172.28.1.38' -NewPrinterPort 'PRNNOBGO012' -Verbose
Set-PrinterPrinterPort -PrinterPort '172.28.1.226' -NewPrinterPort 'PRNNOBGO014' -Verbose
Set-PrinterPrinterPort -PrinterPort '172.28.1.53' -NewPrinterPort 'PRNNOBGO017' -Verbose
Set-PrinterPrinterPort -PrinterPort '172.28.1.33' -NewPrinterPort 'PRNNOBGO018' -Verbose
Set-PrinterPrinterPort -PrinterPort '172.28.1.232' -NewPrinterPort 'PRNNOBGO026' -Verbose
Set-PrinterPrinterPort -PrinterPort '172.28.2.21' -NewPrinterPort 'M3OSLPRT02' -Verbose
Set-PrinterPrinterPort -PrinterPort '172.28.2.20' -NewPrinterPort 'M3OSORFA02' -Verbose
Set-PrinterPrinterPort -PrinterPort '172.28.2.22' -NewPrinterPort 'M3OSORFA03' -Verbose
Set-PrinterPrinterPort -PrinterPort '172.28.2.40' -NewPrinterPort 'PRNNOOSL019' -Verbose
Set-PrinterPrinterPort -PrinterPort '172.28.2.23' -NewPrinterPort 'PRNNOOSL020' -Verbose
Set-PrinterPrinterPort -PrinterPort '172.28.2.16' -NewPrinterPort 'PRNNOOSL021' -Verbose
Set-PrinterPrinterPort -PrinterPort '172.28.3.11' -NewPrinterPort 'M3TRORFA01' -Verbose
Set-PrinterPrinterPort -PrinterPort '172.28.3.21' -NewPrinterPort 'PRNNOTRD025' -Verbose
Set-PrinterPrinterPort -PrinterPort '172.28.3.16' -NewPrinterPort 'PRNNOTRD027' -Verbose

Write-Log -Message 'Completed'
