
break

$ListResult = Get-Content B:\BitLocker.txt | Get-SOEBitLockerStatus -Verbose

break

$ListResult | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation | clip

break

Get-Content B:\BitLocker1.txt | ForEach-Object { Invoke-Command -ComputerName $_ -ScriptBlock { mofcomp.exe c:\windows\system32\wbem\win32_encryptablevolume.mof } }; $ListResult = Get-Content B:\BitLocker1.txt | Get-SOEBitLockerStatus -Verbose

break

$B1 = Get-WmiObject -ComputerName NL5CG518098Q -Namespace "Root\cimv2\Security\MicrosoftVolumeEncryption" -Class "Win32_EncryptableVolume" -Filter "DriveLetter = 'C:'" -ErrorAction Stop
$B2 = Get-WmiObject -ComputerName NL5CG6115VF8 -Namespace "Root\cimv2\Security\MicrosoftVolumeEncryption" -Class "Win32_EncryptableVolume" -Filter "DriveLetter = 'C:'" -ErrorAction Stop

break

#region Get-BitLockerAnalysisReport
$ConnectionString = 'Trusted_Connection=True;Data Source=DEMBMCAPS032SQ2.corp.demb.com\Prod_2'

<#
.Synopsis
    Get the largest used OpCo for the selected country.

.Description
    The OpCo for an asset is a mandatory attribute. But if the user for the OpCo is not yet knowm, the most likely OpCo is the largest OpCo of the specified country.

    The result of this function is the OpCoID that has the most assets registered in the country, according to the asset registration.

.Example
    Get-AMCountryOpCo -CountryCode 'NL'

    >>> 79

.Example
    Get-AMCountryOpCo -CountryCode 'CK'

    No output because this country is not part of the contract.

.Outputs
   System.Int32

.Notes
    --- Version history:
    Version 1.00 (2016-09-02, Kees Hiemstra)
    - Initial version.

.Component
    Asset Management.

.Role
    Querying Asset Management data.

.Functionality
    Get the largest OpCo from the specified country.

.Link
    Get-AMAsset

.Link
    Get-AMOpCo

#>
function Get-BitLockerAnalysisReport
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$false, 
                  PositionalBinding=$false,
                  ConfirmImpact='Low')]
    [Alias()]
    [OutputType([String])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [string]
        $CountryCode
    )

    Begin
    {
        if ([string]::IsNullOrEmpty($ConnectionString))
        {
            $ConnectionString = 'Trusted_Connection=True;Data Source=(Local)'
        }
        Write-Verbose "Connection string: $ConnectionString"

        $Conn = New-Object -TypeName System.Data.SqlClient.SqlConnection
        $Conn.ConnectionString = $ConnectionString
        try
        {
            $Conn.Open()
        }
        catch
        {
            throw "Failed connection with connection string ($ConnectionString)"
        }
        Write-Verbose "Successfully connected to the database"
    }
    Process
    {
        $Cmd = New-Object -TypeName System.Data.SqlClient.SqlCommand
        $Cmd.Connection = $Conn
        $SQL = "SELECT * FROM ITAMProcess.BitLocker.Analysis "
        $SQL += "WHERE [AssetStatus] != 'In stock' "
    	$SQL += "AND [IsOUExclusion] = 0 "
    	$SQL += "AND [IsOSExclusion] = 0 "
    	$SQL += "AND [IsExcemption] = 0 "
    	$SQL += "AND [IsTemporaryExcemption] = 0 "
    	$SQL += "AND ( "
		$SQL += "[DaysFromInstall] < -1 "
		$SQL += "OR [DaysFromLastConnect] < -7 "
		$SQL += "OR [DaysFromADLogon] < -10 "
		$SQL += "OR [DaysFromToday] < -60 "
#		$SQL += "OR [HasSWTpm] = 0 "
#		$SQL += "OR [HasSWMBAM] = 0 "
        $SQL += "OR [BL_IsCompliant] != 'Yes' "
        $SQL += "OR [BL_ErrorInfoName] != 'No Error' "
		$SQL += "OR [BL_DTReporting] IS NULL "
		$SQL += ")"


        $Cmd.CommandText = $SQL
        $Data = $Cmd.ExecuteReader()
        while ($Data.Read())
        {
            $Properties = [ordered]@{
                'ComputerName'          =  [string] $Data['ComputerName']
                'Category'              =  [string] $Data['Category']
                'Model'                 =  [string] $Data['Model']
                'BillingStatus'         =  [string] $Data['BillingStatus']
                'AssetStatus'           =  [string] $Data['AssetStatus']
                'LocationCountryCode'   =  [string] $Data['LocationCountryCode']
                'UserEMail'             =  [string] $Data['UserEMail']
                'adDTLastConnect'       =  [datetime] $Data['adDTLastConnect']
                'adDescription'         =  if ( -not ([DBNull]::Value).Equals($Data['adDescription']) ) { [string] $Data['adDescription'] } else { $null }
                'DTInstall'             =  if ( -not ([DBNull]::Value).Equals($Data['DTInstall']) ) { [datetime] $Data['DTInstall'] } else { $null }
                'DTLastConnect'         =  if ( -not ([DBNull]::Value).Equals($Data['DTLastConnect']) ) { [datetime] $Data['DTLastConnect'] } else { $null }
                'BL_DTReporting'        =  if ( -not ([DBNull]::Value).Equals($Data['BL_DTReporting']) ) { [datetime] $Data['BL_DTReporting'] } else { $null }
                'BL_IsCompliant'        =  if ( -not ([DBNull]::Value).Equals($Data['BL_IsCompliant']) ) { [string] $Data['BL_IsCompliant'] } else { $null }
                'BL_ErrorInfoName'      =  if ( -not ([DBNull]::Value).Equals($Data['BL_ErrorInfoName']) ) { [string] $Data['BL_ErrorInfoName'] } else { $null }
                'HasSWTpm'              =  [bool] $Data['HasSWTpm']
                'HasSWMBAM'             =  [bool] $Data['HasSWMBAM']
                'IsOUExclusion'         =  [bool] $Data['IsOUExclusion']
                'IsOSExclusion'         =  if ( -not ([DBNull]::Value).Equals($Data['IsOSExclusion']) ) { [bool] $Data['IsOSExclusion'] } else { $null }
                'IsExcemption'          =  [bool] $Data['IsExcemption']
                'IsTemporaryExcemption' =  [bool] $Data['IsTemporaryExcemption']
                'DaysFromInstall'       =  if ( -not ([DBNull]::Value).Equals($Data['DaysFromInstall']) ) { [int] $Data['DaysFromInstall'] } else { $null }
                'DaysFromLastConnect'   =  if ( -not ([DBNull]::Value).Equals($Data['DaysFromLastConnect']) ) { [int] $Data['DaysFromLastConnect'] } else { $null }
                'DaysFromADLogon'       =  if ( -not ([DBNull]::Value).Equals($Data['DaysFromADLogon']) ) { [int] $Data['DaysFromADLogon'] } else { $null }
                'DaysFromToday'         =  if ( -not ([DBNull]::Value).Equals($Data['DaysFromToday']) ) { [int] $Data['DaysFromToday'] } else { $null }
                }

            $Return = New-Object -TypeName PSObject -Property $Properties
            Write-Output $Return
        }
        $Data.Close()
    }
    End
    {
        try
        {
            $Conn.Close()
        }
        catch
        {
        }
    }
}

#endregion

$MBAMExclusion = 'CN=Microsoft MBAM Client 2.5 SP1 EN DEMB675 - Exclusions,OU=Core,OU=Software Distribution,OU=SCCM,OU=Software Assignment,DC=corp,DC=demb,DC=com'
$Excemptions = (Get-ADGroup $MBAMExclusion -Properties Member).Member | Convert-DistinguishedNameToName
$MBAMExclusion = 'CN=GPO-C-BitLocker Off,OU=Groups,OU=SoE,DC=corp,DC=demb,DC=com'
$TemporaryExcemptions = (Get-ADGroup $MBAMExclusion -Properties Member).Member | Convert-DistinguishedNameToName
$DisabledComputers = (Get-ADComputer -Filter { Enabled -eq $False }).Name

$Report = Get-BitLockerAnalysisReport | Where-Object { $_.ComputerName -notin $Excemptions -and $_.ComputerName -notin $TemporaryExcemptions -and $_.ComputerName -notin $DisabledComputers }

foreach ( $Item in $Report )
{
    $Analysis = @()
    if ( $Item.BL_DTReporting -eq $null ) { $Analysis += 'Never reported to MBMAM' }
    if ( -not $Item.HasSWTpm ) { $Analysis += 'Tpm software not installed' }
    if ( -not $Item.HasSWMBAM ) { $Analysis += 'MBAM software not installed' }
    if ( $Item.DaysFromInstall -le 0 ) { $Analysis += 'MBAMAgent reported before installation date' }
    if ( ($Item.DaysFromLastConnect -lt 4 -or $Item.DaysFromADLogon -lt 7) -and $Item.DaysFromToday -lt -2 ) { $Analysis += 'MBAMAgent did not recently report' }
    if ( $Item.DaysFromToday -lt -60 ) { $Analysis += 'MBAMAgent did not report for 60 days or more' }

    Add-Member -InputObject $Item -MemberType NoteProperty -Name 'Analisys' -Value ($Analysis -join '; ')
}
