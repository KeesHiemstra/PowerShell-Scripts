#region Get-SOEBitLockerStatus

<#
.SYNOPSIS
    Perform SOE specific checks and read encrtyption using Get-BitLockerStatus.
.DESCRIPTION

.EXAMPLE

.EXAMPLE

.INPUTS
    [string[]]
.OUTPUTS
    [PSObject]
.NOTES
    ===Version History
    Version 1.20 (2017-01-19, Kees Hiemstra)
    - Added the attribute EncryptionMethod: 1 = Windows BitLocker, 4 = MBAM Client.
    - Added the attribute Identification: Should be value set by the GPO.
    - Added the attribute ProtectorName.
    Version 1.10 (2017-01-17, Kees Hiemstra)
    - Added the attribute MBAMAgent to show the status of the MBAMAgent service.
    Version 1.00 (2017-01-13, Kees Hiemstra)
    - Inital version.
.COMPONENT
    BitLocker
.ROLE

.FUNCTIONALITY

#>
function Get-SOEBitLockerStatus
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([String])]
    Param
    (
        #
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [string[]]
        $ComputerName
    )

    Begin
    {
    }
    Process
    {
        foreach ( $Computer in $ComputerName )
        {
            Write-Verbose $Computer

            $Properties = [ordered] @{'ComputerName'       = [string]   $Computer
                                      'ADOperatingSystem'  = [string]   ''
                                      'ADEnabled'          = [bool]     $false
                                      'ADLastLogonDate'    = [datetime] (Get-Date -Year 1980 -Month 1 -Day 1).Date
                                      'Status'             = [string]   ''
                                      'TpmActivated'       = [string]   ''
                                      'TpmEnabled'         = [string]   ''
                                      'TpmOwned'           = [string]   ''
                                      'TpmPresenceVersion' = [string]   ''
                                      'KeyType'            = [string]   ''
                                      'ProtectionStatus'   = [string]   ''
                                      'EncryptionStatus'   = [string]   ''
                                      'EncryptionProgress' = [string]   ''
                                      'ErrorMessage'       = [string]   ''
                                      'ExtraInfo'          = [string]   ''
                                      'EncryptionMethod'   = [string]   ''
                                      'Identification'     = [string]   ''
                                      'ProtectorName'      = [string]   ''
                                      'MBAMAgent'          = [string]   ''
                                      'WhenChecked'        = [datetime] (Get-Date)
                                     }
            $Result = New-Object -TypeName psobject -Property $Properties

            $ADComputer = $null
            try
            {
                $ADComputer = (Get-ADComputer $Computer -Properties MemberOf, OperatingSystem, LastLogonDate -ErrorAction Stop)
            }
            catch
            {
                $Result.Status = "AD Error"
                $Result.ErrorMessage = $Global:Error[0].Exception.Message
                Write-Output $Result
                continue
            }

            if ( $ADComputer -eq $null )
            {
                $Result.Status = "AD Error: Computer not found"
                Write-Output $Result
                continue
            }

            $Result.ADOperatingSystem = $ADComputer.OperatingSystem
            $Result.ADEnabled = $ADComputer.Enabled
            $Result.ADLastLogonDate = $ADComputer.LastLogonDate

            if ( $ADComputer.DistinguishedName -notlike '*,OU=Win7,OU=Managed Computers,DC=corp,DC=demb,DC=com')
            {
                $Result.Status = "AD: Not in Win7 OU"
                Write-Output $Result
                continue
            }

            if ( $ADComputer.MemberOf -contains 'CN=Microsoft MBAM Client 2.5 SP1 EN DEMB675 - Exclusions,OU=Core,OU=Software Distribution,OU=SCCM,OU=Software Assignment,DC=corp,DC=demb,DC=com' )
            {
                $Result.Status = 'AD: Member of AD exception list'
                Write-Output $Result
                continue
            }

            if ( $ADComputer.MemberOf -contains 'CN=GPO-C-BitLocker Off,OU=Groups,OU=SoE,DC=corp,DC=demb,DC=com' )
            {
                $Result.Status = 'AD: Member of temporarily AD exception list'
                Write-Output $Result
                continue
            }

            if ( -not (Test-Connection -ComputerName $Computer -Count 1 -Quiet) )
            {
                $Result.Status = 'Not online'
                Write-Output $Result
                continue
            }

            if ( -not (Test-Connection -ComputerName $Computer -Count 1 -Authentication Connect -Quiet) )
            {
                $Result.Status = 'Not accessible'
                Write-Output $Result
                continue
            }

            try
            {
                $Null = Get-WmiObject -ComputerName $Computer -Class Win32_ComputerSystem -ErrorAction Stop
            }
            catch
            {
                $Result.Status = 'Not accessible'
                Write-Output $Result
                continue
            }

            $BitLocker = Get-BitLockerStatus -ComputerName $Computer -NoOnlineCheck

            $Result.Status             = $BitLocker.Status 
            $Result.TpmActivated       = $BitLocker.TpmActivated
            $Result.TpmEnabled         = $BitLocker.TpmEnabled
            $Result.TpmOwned           = $BitLocker.TpmOwned
            $Result.TpmPresenceVersion = $BitLocker.TpmPresenceVersion
            $Result.KeyType            = $BitLocker.KeyType
            $Result.ProtectionStatus   = $BitLocker.ProtectionStatus
            $Result.EncryptionStatus   = $BitLocker.EncryptionStatus
            $Result.EncryptionProgress = $BitLocker.EncryptionProgress
            $Result.ErrorMessage       = $BitLocker.ErrorMessage
            $Result.EncryptionMethod   = $BitLocker.EncryptionMethod
            $Result.Identification     = $BitLocker.Identification
            $Result.ProtectorName      = $BitLocker.ProtectorName
            $Result.MBAMAgent          = $BitLocker.MBAMAgent

            [array]$ExtraInfo = @()

            if ( $Result.TpmPresenceVersion -ne '1.2' )
            {
                $ExtraInfo += 'Wrong Tpm version'
            }

            if ( $Result.EncryptionStatus -in @('', 'Fully Decrypted') )
            {
                try
                {
                    [array]$Disk = Get-WmiObject -ComputerName $Computer -Class Win32_LogicalDisk -Filter "DriveType <> '5'"
                    
                    if ( $Disk -ne $null )
                    {
                        if ( $Disk.Count -gt 1 )
                        {
                            $ExtraInfo += 'More then one disk'
                        }

                        $Free = [math]::Truncate((($Disk[0].FreeSpace / 1GB) - ((3GB / 2) / 1GB) * 100) / ($Disk[0].Size / 1GB) * 100)
                        $ExtraInfo += "$(Disk[0].DeviceID) has $Free% free space"
                    }
                }
                catch
                {
                    #$Result.ErrorMessage = $Global:Error[0].Exception.Message
                }
            }

            $Result.ExtraInfo = $ExtraInfo -join '; '

            Write-Output $Result

        }
    }
    End
    {
    }
}

#endregion
