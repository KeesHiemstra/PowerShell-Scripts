#region Get-BitLockerStatus

<#
.Synopsis
    Read encrtyption data for the selected computer to determine the health of the excryption.
.DESCRIPTION

.EXAMPLE
    Get-BitLockerStatus MBAMBitLock

    The computer in this example is encrypted using the MBAM Client
    >>>
    ComputerName       : MBAMBitLock
    Status             : 
    TpmActivated       : True
    TpmEnabled         : True
    TpmOwned           : True
    TpmPresenceVersion : 1.2
    KeyType            : Numerical password; Trusted Platform Module (TPM)
    ProtectionStatus   : Protected
    EncryptionStatus   : Fully Encrypted
    EncryptionProgress : 100
    ErrorMessage       : 
    EncryptionMethod   : 4
    Identification     : JDE
    ProtectorName      : Malta
    MBAMAgent          : Running

.EXAMPLE
    Get-BitLockerStatus -ComputerName WinBitLock

    The computer in this example is encrypted using Windows BitLocker
    >>>
    ComputerName       : WinBitLock
    Status             : 
    TpmActivated       : True
    TpmEnabled         : True
    TpmOwned           : True
    TpmPresenceVersion : 1.2
    KeyType            : Trusted Platform Module (TPM); Numerical password
    ProtectionStatus   : Protected
    EncryptionStatus   : Fully Encrypted
    EncryptionProgress : 100
    ErrorMessage       : 
    EncryptionMethod   : 1
    Identification     : 
    ProtectorName      : 
    MBAMAgent          : Cannot find any service with service name 'MBAMAgent'.

.INPUTS
    [string[]]
.OUTPUTS
    [PSObject]
.NOTES
    ===Version History
    Version 2.20 (2017-01-19, Kees Hiemstra)
    - Added the attribute EncryptionMethod: 1 = Windows BitLocker, 4 = MBAM Client.
    - Added the attribute Identification: Should be value set by the GPO.
    - Added the attribute ProtectorName.
    Version 2.10 (2017-01-17, Kees Hiemstra)
    - Added the attribute MBAMAgent to show the status of the MBAMAgent service.
    Version 2.00 (2017-01-13, Kees Hiemstra)
    - Moved the specific company checks to separate function.
    Version 1.00 (2017-01-11, Kees Hiemstra)
    - Inital version.
.COMPONENT
    BitLocker
.ROLE
    Colling remote data
#>
function Get-BitLockerStatus
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
        $ComputerName,

        #
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=1,
                   ParameterSetName='Parameter Set 1')]
        [switch]
        $NoOnlineCheck
    )

    Begin
    {
    }
    Process
    {
        foreach ( $Computer in $ComputerName )
        {
            $Properties = [ordered] @{'ComputerName'       = [string]   $Computer
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
                                      'EncryptionMethod'   = [string]   ''
                                      'Identification'     = [string]   ''
                                      'ProtectorName'      = [string]   ''
                                      'MBAMAgent'          = [string]   ''
                                     }
            $Result = New-Object -TypeName psobject -Property $Properties

            if ( -not $NoOnlineCheck.ToBool() -and -not (Test-Connection -ComputerName $Computer -Count 1 -Quiet) )
            {
                $Result.Status = 'Not online'
                Write-Output $Result
                continue
            }

            if ( -not $NoOnlineCheck.ToBool() -and -not (Test-Connection -ComputerName $Computer -Count 1 -Authentication Connect -Quiet) )
            {
                $Result.Status = 'Not accessible'
                Write-Output $Result
                continue
            }

            try
            {
                $null = Get-WmiObject -ComputerName $Computer -Class Win32_ComputerSystem -ErrorAction Stop
            }
            catch
            {
                $Result.Status = 'Not accessible'
                Write-Output $Result
                continue
            }

            try
            {
                $Tpm = Get-WmiObject -ComputerName $Computer -Namespace ROOT\CIMV2\Security\MicrosoftTpm -Class Win32_Tpm -ErrorAction Stop

                if ( $Tpm -ne $null )
                {
                    $Result.TpmActivated = $Tpm.IsActivated().IsActivated
                    $Result.TpmEnabled = $Tpm.IsEnabled().IsEnabled
                    $Result.TpmOwned = $Tpm.IsOwned().IsOwned
                    $Result.TpmPresenceVersion = $Tpm.PhysicalPresenceVersionInfo
                }
                else
                {
                    $Result.Status = "No Tpm data available"
                }
            }
            catch [System.Management.ManagementException]
            {
                $Result.Status = "Tpm WMI Error: $($Global:Error[0].Exception.Message)"
#                $Result.ErrorMessage = $Global:Error[0].Exception.Message
#                Write-Output $Result
#                continue
            }
            catch
            {
                $Result.Status = "WMI General error (Tpm): $($Global:Error[0].Exception.Message)"
#                $Result.ErrorMessage = $Global:Error[0].Exception.Message
#                Write-Output $Result
#                continue
            }

            try
            {
                $MBAMAgent = Invoke-Command -ComputerName $Computer -ScriptBlock { Get-Service MBAMAgent } -ErrorAction Stop
                $Result.MBAMAgent = $MBAMAgent.Status
            }
            catch
            {
                $Result.MBAMAgent = $Global:Error[0].Exception.Message
            }

            try
            {
                $BitLocker = Get-WmiObject -ComputerName $Computer -Namespace "Root\cimv2\Security\MicrosoftVolumeEncryption" -Class "Win32_EncryptableVolume" -Filter "DriveLetter = 'C:'" -ErrorAction Stop
            }
            catch [System.Management.ManagementException]
            {
                $Result.Status = "BitLocker WMI Error"
                $Result.ErrorMessage = $Global:Error[0].Exception.Message
                Write-Output $Result
                continue
            }
            catch
            {
                $Result.Status = "WMI General error (BitLocker)"
                $Result.ErrorMessage = $Global:Error[0].Exception.Message
                Write-Output $Result
                continue
            }
    
            if ( $BitLocker -eq $null )
            {
                $Result.Status = 'No BitLocker data available'
                Write-Output $Result
                continue
            }

            $ProtectorIds = $BitLocker.GetKeyProtectors("0").VolumeKeyProtectorID
            $KeyType = @()

            foreach ( $ProtectorID in $ProtectorIds )
            {
                $KeyProtectorType = $BitLocker.GetKeyProtectorType($ProtectorID).KeyProtectorType

                switch ( $KeyProtectorType )
                {
                    "0"  { $KeyType += @("Unknown or other protector type") }
                    "1"  { $KeyType += @("Trusted Platform Module (TPM)") }
                    "2"  { $KeyType += @("External key") }
                    "3"  { $KeyType += @("Numerical password") }
                    "4"  { $KeyType += @("TPM And PIN") }
                    "5"  { $KeyType += @("TPM And Startup Key") }
                    "6"  { $KeyType += @("TPM And PIN And Startup Key") }
                    "7"  { $KeyType += @("Public Key") }
                    "8"  { $KeyType += @("Passphrase") }
                    "9"  { $KeyType += @("TPM Certificate") }
                    "10" { $KeyType += @("CryptoAPI Next Generation (CNG) Protector") }
                }#-switch
            }#-foreach

            $Result.KeyType = $KeyType -join '; '
            try
            {
                $Result.ProtectorName = $BitLocker.GetKeyProtectorFriendlyName($ProtectorIds[0]).FriendlyName
            }
            catch
            {
                $Result.ProtectorName = "FriendlyName error: $($Global:Error[0].Exception.Message)"
            }

            $ProtectionStatus = $BitLocker.GetProtectionStatus()

            switch ( $ProtectionStatus.ProtectionStatus )
            {
                "0"     { $Result.ProtectionStatus = "Unprotected" }
                "1"     { $Result.ProtectionStatus = "Protected" }
                "2"     { $Result.ProtectionStatus = "Uknowned" }
                default { $Result.ProtectionStatus = "NoReturn" }
	        }

            $ConversionStatus = $BitLocker.GetConversionStatus()
            $Result.EncryptionProgress = $ConversionStatus.EncryptionPercentage

            switch ($ConversionStatus.ConversionStatus)
            {
                "0"     { $Result.EncryptionStatus = 'Fully Decrypted' }
                "1"     { $Result.EncryptionStatus = 'Fully Encrypted' }
                "2"     { $Result.EncryptionStatus = 'Encryption in progress' }
                "3"     { $Result.EncryptionStatus = 'Decryption in progress' }
                "4"     { $Result.EncryptionStatus = 'Encryption paused' }
                "5"     { $Result.EncryptionStatus = 'Decryption paused' }
                default { }
            }

            $Result.EncryptionMethod = $BitLocker.GetEncryptionMethod().EncryptionMethod

            $Result.Identification = $BitLocker.GetIdentificationField().IdentificationField

            Write-Output $Result

        }
    }
    End
    {
    }
}

#endregion
