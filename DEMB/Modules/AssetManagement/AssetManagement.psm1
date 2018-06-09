$ConnectionString = 'Trusted_Connection=True;Data Source=DEMBMCAPS032SQ2.corp.demb.com\Prod_2'

#region Get-AMAsset

<#
.Synopsis
    Query the Asset table from the Asset Managament Database (ITAMWeb). The query can be filtered using the parameters.

.Description
    The returned database is presented as an AM.Asset.100 object and contains the folling fields:
    - ComputerName         : 
    - SerialNo             : 
    - Category             : 
    - Brand                : 
    - Model                : 
    - BillingStatus        : 
    - AssetStatus          : 
    - AssetCountryCode     :   
    - AssetOpCoNo          : 
    - AssetOpCoFull        : 
    - LocationCountryCode  : 
    - LocationCountryName  : 
    - LocationName         : 
    - LocationDetail       : 
    - UserAccount          : 
    - DomainAccount        : 
    - UserFirstName        : 
    - UserLastName         : 
    - UserPhone            : 
    - UserEMail            : 
    - UserDepartment       : 
    - InternalTag          : 
    - AssetTag             : 
    - AssetCostCenter      : 
    - DTAcquisition        : 
    - DTInstall            : 
    - DTRefresh            : 
    - DTMutation           : 
    - ProductRef           : 
    - RadiaStatus          : 
    - IsComputer           : 
    - IsDesktop            : 
    - IsInContractHP       : 
    - IsInContractSL       : 
    - IsInStock            : 
    - IsObsolete           : 
    - IsToBeDisposed       : 
    - IsDummy              : 
    - IsChangePending      : 
    - MonthToRefresh       : 
    - DaysFromLastMutation : 

.Example
    Get-AMAsset -ComputerName HPNLDev06

    Retruns the following data:
        ComputerName         : HPNLDEV06
        SerialNo             : 5CG52218MW
        Category             : Laptop
        Brand                : HEWLETT-PACKARD
        Model                : HP EliteBook 840 G2
        BillingStatus        : Not in contract, used by HP
        AssetStatus          : In use
        AssetCountryCode     :   
        AssetOpCoNo          : 9000
        AssetOpCoFull        : 9000 HP
        LocationCountryCode  : NL
        LocationCountryName  : The Netherlands
        LocationName         : Utrecht
        LocationDetail       : 
        UserAccount          : DEMB\KEES.HIEMSTRA
        DomainAccount        : KEES.HIEMSTRA
        UserFirstName        : Kees
        UserLastName         : Hiemstra
        UserPhone            : 
        UserEMail            : kees.hiemstra@jdecoffee.com
        UserDepartment       : 
        InternalTag          : MAC2000031169
        AssetTag             : NL5CG52218MW
        AssetCostCenter      : 
        DTAcquisition        : 2015-06-04 00:00:00
        DTInstall            : 2015-06-08 00:00:00
        DTRefresh            : 2019-06-04 00:00:00
        DTMutation           : 2016-03-21 14:42:00
        ProductRef           : 
        RadiaStatus          : 
        IsComputer           : True
        IsDesktop            : True
        IsInContractHP       : False
        IsInContractSL       : False
        IsInStock            : False
        IsObsolete           : False
        IsToBeDisposed       : False
        IsDummy              : False
        IsChangePending      : False
        MonthToRefresh       : 34
        DaysFromLastMutation : 162

.Example
    Get-AMAsset -ComputerName HPNLDev06 -IsNotObsolete:$false

    Will not return any data because the computer is not obsolete.

.Outputs
    [AM.OpCo.100] Object

.Notes
    --- Version history:
    Version 1.30 (2017-02-15, Kees Hiemstra)
    - Add MonthToRefreshOlderThen parameter
    Version 1.20 (2017-02-10, Kees Hiemstra)
    - Add SerialNo parameter.
    Version 1.10 (2016-12-29, Kees Hiemstra)
    - Add UseIMACD parameter.
    - Add IsToBeDisposed as field.
    - Add IsNotToBeDisped AS paramet.
    Version 1.00 (2016-08-29, Kees Hiemstra)
    - Initial version.

.Component
    Asset Management.

.Role
    Querying Asset Management master data.

.Functionality
    Get information from Asset Manament database.

.Link
    Get-AMOpCo

.Link
    Get-AMCountryOpCo

#>
function Get-AMAsset
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Low')]
    [Alias()]
    [OutputType([String])]
    Param
    (
        #Computer name or asset name, wildcards are allowed
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [Alias("AssetName")]
        [string]
        $ComputerName,

        #Serial number, wildcards are allowed
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [Alias("SerialNumber")]
        [string]
        $SerialNo,

        #Filter on Asset OpCo
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false)]
        [int]
        $OpCo,

        #SAMAccountName of the user
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [Alias("UserAccount")]
        [string]
        $SAMAccountName,

        #Filter on EUW computers
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false)]
        [switch]
        $IsDesktop,

        #Filter on EUW computers
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false)]
        [switch]
        $IsInStock,

        #Filter on assets that have no change pending
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false)]
        [switch]
        $HasNoPendingChange,

        #Filter on assets where the last mutation date is older than the number of given days
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false)]
        [int]
        $LastMutationOlderThen,

        #Filter on assets where the To be refresh date is older than the number of given months
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false)]
        [int]
        $MonthToRefreshOlderThen,

        #Filter on not obsolete computers
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false)]
        [switch]
        $IsNotObsolete,

        #Filter on not to be disposed computers
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false)]
        [switch]
        $IsNotToBeDisposed,

        #Filter on dummy assets
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false)]
        [switch]
        $IsDummy,

        #Filter on to be disposed computers
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false)]
        [switch]
        $IsToBeDisposed,

        #Use IMACD table to show pendinging changes
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false)]
        [switch]
        $UseIMACD
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

        if ( $PSBoundParameters.ContainsKey('UseIMACD') )
        {
            $SQL = "SELECT * FROM SOEAdmin.ps.IMACD"
        }
        else
        {
            $SQL = "SELECT * FROM SOEAdmin.ps.AssetList"
        }

        if ( $PSBoundParameters.ContainsKey('ComputerName') )
        {
            if ( $SQL -match 'WHERE' )
            {
                $SQL = "$SQL AND"
            }
            else
            {
                $SQL = "$SQL WHERE"
            }
            $SQL = "$SQL [AssetName] LIKE @AssetName"

            $Cmd.Parameters.Add('AssetName', [string]) | Out-Null
            $Cmd.Parameters['AssetName'].Value = [string]$ComputerName.Replace('*', '%').Replace('?', '_')

            Write-Verbose "ComputerName: $ComputerName"
        }

        if ( $PSBoundParameters.ContainsKey('SerialNo') )
        {
            if ( $SQL -match 'WHERE' )
            {
                $SQL = "$SQL AND"
            }
            else
            {
                $SQL = "$SQL WHERE"
            }
            $SQL = "$SQL [SerialNo] LIKE @SerialNo"

            $Cmd.Parameters.Add('SerialNo', [string]) | Out-Null
            $Cmd.Parameters['SerialNo'].Value = [string]$SerialNo.Replace('*', '%').Replace('?', '_')

            Write-Verbose "SerialNo: $SerialNo"
        }

        if ( $PSBoundParameters.ContainsKey('OpCo') )
        {
            if ( $SQL -match 'WHERE' )
            {
                $SQL = "$SQL AND"
            }
            else
            {
                $SQL = "$SQL WHERE"
            }
            $SQL = "$SQL [AssetOpCoID] = @AssetOpCoID"

            $Cmd.Parameters.Add('AssetOpCoID', [int]) | Out-Null
            $Cmd.Parameters['AssetOpCoID'].Value = [int]$OpCo

            Write-Verbose "ComputerName: $ComputerName"
        }

        if ( $PSBoundParameters.ContainsKey('SAMAccountName') )
        {
            if ( $SQL -match 'WHERE' )
            {
                $SQL = "$SQL AND"
            }
            else
            {
                $SQL = "$SQL WHERE"
            }
            $SQL = "$SQL [UserAccount] LIKE @UserAccount"

            $Cmd.Parameters.Add('UserAccount', [string]) | Out-Null
            $Cmd.Parameters['UserAccount'].Value = [string]"%\$($SAMAccountName.Replace('*', '%').Replace('?', '_'))"

            Write-Verbose "SAMAccountName: $SAMAccountName"
        }

        if ( $PSBoundParameters.ContainsKey('LastMutationOlderThen') )
        {
            if ( $SQL -match 'WHERE' )
            {
                $SQL = "$SQL AND"
            }
            else
            {
                $SQL = "$SQL WHERE"
            }
            $SQL = "$SQL [DaysFromLastMutation] > @DaysFromLastMutation"

            $Cmd.Parameters.Add('DaysFromLastMutation', [string]) | Out-Null
            $Cmd.Parameters['DaysFromLastMutation'].Value = [string]$LastMutationOlderThen

            Write-Verbose "last mutation date is older than: $LastMutationOlderThen days"
        }

        if ( $PSBoundParameters.ContainsKey('MonthToRefreshOlderThen') )
        {
            if ( $SQL -match 'WHERE' )
            {
                $SQL = "$SQL AND"
            }
            else
            {
                $SQL = "$SQL WHERE"
            }
            $SQL = "$SQL [MonthToRefresh] < @MonthToRefresh"

            $Cmd.Parameters.Add('MonthToRefresh', [string]) | Out-Null
            $Cmd.Parameters['MonthToRefresh'].Value = [string]$MonthToRefreshOlderThen

            Write-Verbose "Refresh should take place in $MonthToRefreshOlderThen months"
        }

        if ( $PSBoundParameters.ContainsKey('IsDesktop') )
        {
            if ( $SQL -match 'WHERE' )
            {
                $SQL = "$SQL AND"
            }
            else
            {
                $SQL = "$SQL WHERE"
            }
            $SQL = "$SQL [IsDesktop] = @IsDesktop"

            $Cmd.Parameters.Add('IsDesktop', [bool]) | Out-Null
            $Cmd.Parameters['IsDesktop'].Value = [bool]($IsDesktop.ToBool())

            Write-Verbose "IsDesktop: $($IsDesktop.ToBool())"
        }

        if ( $PSBoundParameters.ContainsKey('IsInStock') )
        {
            if ( $SQL -match 'WHERE' )
            {
                $SQL = "$SQL AND"
            }
            else
            {
                $SQL = "$SQL WHERE"
            }
            $SQL = "$SQL [IsInStock] = @IsInStock"

            $Cmd.Parameters.Add('IsInStock', [bool]) | Out-Null
            $Cmd.Parameters['IsInStock'].Value = [bool]($IsInStock.ToBool())

            Write-Verbose "IsDesktop: $($IsDesktop.ToBool())"
        }

        if ( $PSBoundParameters.ContainsKey('IsToBeDisposed') )
        {
            if ( $SQL -match 'WHERE' )
            {
                $SQL = "$SQL AND"
            }
            else
            {
                $SQL = "$SQL WHERE"
            }
            if ( $IsToBeRetired.ToBool() )
            {
                $SQL = "$SQL [BillingStatus] = 'Not in contract, to be disposed'"
            }
            else
            {
                $SQL = "$SQL [BillingStatus] != 'Not in contract, to be disposed'"
            }

            Write-Verbose "IsToBeRetired: $($IsToBeRetired.ToBool())"
        }

        if ( $PSBoundParameters.ContainsKey('IsDummy') )
        {
            if ( $SQL -match 'WHERE' )
            {
                $SQL = "$SQL AND"
            }
            else
            {
                $SQL = "$SQL WHERE"
            }
            $SQL = "$SQL [IsDummy] = @IsDummy AND [IsObsolete] = 0"

            $Cmd.Parameters.Add('IsDummy', [bool]) | Out-Null
            $Cmd.Parameters['IsDummy'].Value = [bool]($IsDummy.ToBool())

            Write-Verbose "IsDummy: $($IsDummy.ToBool())"
        }

        if ( $PSBoundParameters.ContainsKey('IsNotObsolete') )
        {
            if ( $SQL -match 'WHERE' )
            {
                $SQL = "$SQL AND"
            }
            else
            {
                $SQL = "$SQL WHERE"
            }
            $SQL = "$SQL [IsObsolete] = @IsObsolete"

            $Cmd.Parameters.Add('IsObsolete', [bool]) | Out-Null
            $Cmd.Parameters['IsObsolete'].Value = [bool](-not $IsNotObsolete.ToBool())

            Write-Verbose "IsNotObsolete: $($IsNotObsolete.ToBool())"
        }

        if ( $PSBoundParameters.ContainsKey('IsNotToBeDisposed') )
        {
            if ( $SQL -match 'WHERE' )
            {
                $SQL = "$SQL AND"
            }
            else
            {
                $SQL = "$SQL WHERE"
            }
            $SQL = "$SQL [IsToBeDisposed] = @IsToBeDisposed"

            $Cmd.Parameters.Add('IsToBeDisposed', [bool]) | Out-Null
            $Cmd.Parameters['IsToBeDisposed'].Value = [bool](-not $IsNotToBeDisposed.ToBool())

            Write-Verbose "IsNotObsolete: $($IsNotObsolete.ToBool())"
        }

        if ( $PSBoundParameters.ContainsKey('HasNoPendingChange') )
        {
            if ( $SQL -match 'WHERE' )
            {
                $SQL = "$SQL AND"
            }
            else
            {
                $SQL = "$SQL WHERE"
            }
            $SQL = "$SQL [IsChangePending] = @IsChangePending"

            $Cmd.Parameters.Add('IsChangePending', [bool]) | Out-Null
            $Cmd.Parameters['IsChangePending'].Value = [bool](-not $HasNoPendingChange.ToBool())

            Write-Verbose "HasNoPendingChange: $($HasNoPendingChange.ToBool())"
        }

        Write-Verbose $SQL

        if ( $PSBoundParameters.ContainsKey('UseIMACD') )
        {
            $SQL = "$SQL ORDER BY [DTMutation]"
        }

        $Cmd.CommandText = $SQL

        $Data = $Cmd.ExecuteReader()
        while ($Data.Read())
        {
            $Properties = [ordered]@{'ComputerName'         = [string]$Data['AssetName'].ToString().Trim()
                                     'SerialNo'             = [string]$Data['SerialNo']
                                     'Category'             = [string]$Data['Category']
                                     'Brand'                = [string]$Data['Brand']
                                     'Model'                = [string]$Data['Model']
                                     'BillingStatus'        = [string]$Data['BillingStatus']
                                     'AssetStatus'          = [string]$Data['AssetStatus']
                                     'AssetCountryCode'     = [string]$Data['AssetCountryCode']
                                     'AssetOpCoNo'          = [string]$Data['AssetOpCoNo']
                                     'AssetOpCoFull'        = [string]$Data['AssetOpCoFull']
                                     'LocationCountryCode'  = [string]$Data['LocationCountryCode']
                                     'LocationCountryName'  = [string]$Data['LocationCountryName']
                                     'LocationName'         = [string]$Data['LocationName']
                                     'LocationDetail'       = [string]$Data['LocationDetail']
                                     'UserAccount'          = [string]$Data['UserAccount']
                                     'DomainAccount'        = [string]$Data['DomainAccount']
                                     'UserFirstName'        = [string]$Data['UserFirstName']
                                     'UserLastName'         = [string]$Data['UserLastName']
                                     'UserPhone'            = [string]$Data['UserPhone']
                                     'UserEMail'            = [string]$Data['UserEMail']
                                     'UserDepartment'       = [string]$Data['UserDepartment']
                                     'InternalTag'          = [string]$Data['InternalTag']
                                     'AssetTag'             = [string]$Data['AssetTag']
                                     'AssetCostCenter'      = [string]$Data['AssetCostCenter']
                                     'DTAcquisition'        = [datetime]$Data['DTAcquisition']
                                     'DTInstall'            = [datetime]$Data['DTInstall']
                                     'DTRefresh'            = [datetime]$Data['DTRefresh']
                                     'DTMutation'           = [datetime]$Data['DTMutation']
                                     'ProductRef'           = [string]$Data['ProductRef']
                                     'RadiaStatus'          = [string]$Data['RadiaStatus']
                                     'IsComputer'           = [bool]$Data['IsComputer']
                                     'IsDesktop'            = [bool]$Data['IsDesktop']
                                     'IsInContractHP'       = [bool]$Data['IsInContractHP']
                                     'IsInContractSL'       = [bool]$Data['IsInContractSL']
                                     'IsInStock'            = [bool]$Data['IsInStock']
                                     'IsObsolete'           = [bool]$Data['IsObsolete']
                                     'IsToBeDisposed'       = [bool]$Data['IsToBeDisposed']
                                     'IsDummy'              = [bool]$Data['IsDummy']
                                     'IsChangePending'      = [bool]$Data['IsChangePending']
                                     'MonthToRefresh'       = [int]$Data['MonthToRefresh']
                                     'DaysFromLastMutation' = [int]$Data['DaysFromLastMutation']
                                    }
            $Return = New-Object -TypeName PSObject -Property $Properties
            if ( $PSBoundParameters.ContainsKey('UseIMACD') )
            {
                Add-Member -Force -InputObject $Return -NotePropertyname 'OldAssetName' -NotePropertyValue $Data['OldAssetName'].ToString()
                Add-Member -Force -InputObject $Return -NotePropertyname 'OldSerialNo' -NotePropertyValue $Data['OldSerialNo'].ToString()
                Add-Member -Force -InputObject $Return -NotePropertyname 'OldBillingStatus' -NotePropertyValue $Data['OldBillingStatus'].ToString()
                Add-Member -Force -InputObject $Return -NotePropertyname 'OldAssetStatus' -NotePropertyValue $Data['OldAssetStatus'].ToString()
                Add-Member -Force -InputObject $Return -NotePropertyname 'OldUserAccount' -NotePropertyValue $Data['OldUserAccount'].ToString()
                Add-Member -Force -InputObject $Return -NotePropertyname 'EngineerUID' -NotePropertyValue $Data['EngineerUID'].ToString()
                Add-Member -Force -InputObject $Return -NotePropertyname 'ChangeRef' -NotePropertyValue $Data['ChangeRef'].ToString()
                $Return.PSObject.TypeNames.Insert(0, 'AM.IMACD.100')
            }
            else
            {
                $Return.PSObject.TypeNames.Insert(0, 'AM.Asset.100')
            }
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

#region Get-AMOpCo

<#
.Synopsis
    Query the OpCo table from the Asset Managament Database (ITAMWeb).

.Description
    The returned database is presented as an AM.OpCo.100 object and contains the folling fields:
    - ID
    - Number
    - Name     
    - FullName
    - CountryCode
    - Managed
    - ReportAs
    - WhenCreated

.Example
    Get-AMOpCo

    Returns all of the data from the table.
.Example
    Get-AMOpCo -ManagedOnly
        or
    Get-AMOpCo -ManagedOnly:$true


    Will return of the data from the table where the OpCo is managed by the contract.
.Example
    Get-AMOpCo -ManagedOnly:$false

    Returns of all of the data from the table where the OpCo is NOT managed by the contract, for Asset Management this means an asset assigned to this OpCo can't be invoiced.

.Example
    Get-AMOpCo -ID 2
        or
    Get-AMOpCo -Number '0002'

        ID          : 2
        Number      : 0002
        Name        : KDE BV
        Fullname    : 0002 KDE BV
        CountryCode : NL
        Managed     : True

.Example
    Get-AMOpCo -ID 9020 -ManagedOnly
        or
    Get-AMOpCo -Number '9020' -ManagedOnly

    Will not return data because this OpCo is not managed by the contract.

.Outputs
    [AM.OpCo.100] Object

.Notes
    --- Version history:
    Version 1.00 (2016-08-29, Kees Hiemstra)
    - Initial version.

.Component
    Asset Management.

.Role
    Querying Asset Management master data.

.Functionality
    Get information from Asset Manament database.

.Link
    Get-AMAsset

.Link
    Get-AMCountryOpCo

#>
function Get-AMOpCo
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Low')]
    [Alias()]
    [OutputType([String])]
    Param
    (
        #OpCo ID
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$false,
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [int]
        $ID,

        #OpCo Number
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$false,
                   Position=0,
                   ParameterSetName='Parameter Set 2')]
        [Alias('Department')]
        [string]
        $Number,

        #Filter on managed OpCos only
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false)]
        [switch]
        $ManagedOnly
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
        $SQL = "SELECT * FROM ITAMData.dbo.riOpCo"

        if ( $PSBoundParameters.ContainsKey('ID') )
        {
            if ( $SQL -match 'WHERE' )
            {
                $SQL = "$SQL AND"
            }
            else
            {
                $SQL = "$SQL WHERE"
            }
            $SQL = "$SQL [OpCoID] = @OpCoID"

            $Cmd.Parameters.Add('OpCoID', [int]) | Out-Null
            $Cmd.Parameters['OpCoID'].Value = [int]$ID

            Write-Verbose "ID: $ID"
        }

        if ( $PSBoundParameters.ContainsKey('Number') )
        {
            if ( $SQL -match 'WHERE' )
            {
                $SQL = "$SQL AND"
            }
            else
            {
                $SQL = "$SQL WHERE"
            }
            $SQL = "$SQL [OpCoNo] = @OpCoNo"

            $Cmd.Parameters.Add('OpCoNo', [string]) | Out-Null
            $Cmd.Parameters['OpCoNo'].Value = [string]$Number

            Write-Verbose "Number: $Number"
        }

        if ( $PSBoundParameters.ContainsKey('ManagedOnly') )
        {
            if ( $SQL -match 'WHERE' )
            {
                $SQL = "$SQL AND"
            }
            else
            {
                $SQL = "$SQL WHERE"
            }
            $SQL = "$SQL [HPManaged] = @HPManaged"

            $Cmd.Parameters.Add('HPManaged', [bool]) | Out-Null
            $Cmd.Parameters['HPManaged'].Value = [bool]$ManagedOnly.ToBool()

            Write-Verbose "ManagedOnly: $($ManagedOnly.ToBool())"
        }

        Write-Verbose $SQL

        $Cmd.CommandText = $SQL

        $Data = $Cmd.ExecuteReader()
        while ($Data.Read())
        {
            $Properties = [ordered]@{'ID'          = [int]$Data['OpCoID']
                                     'Number'      = [string]$Data['OpCoNo']
                                     'Name'        = [string]$Data['OpCoName']
                                     'FullName'    = [string]"$($Data['OpCoNo'].ToString()) $($Data['OpCoName'].ToString())"
                                     'CountryCode' = [string]$Data['OpCoCountryCode']
                                     'Managed'     = [bool]$Data['HPManaged']
                                     'ReportAs'    = [string]$Data['ReportEntity']
                                     'WhenCreated' = [datetime]$Data['DTCreation']
                                     'Comment'     = [string]$Data['Comment']
                                    }
            $Return = New-Object -TypeName PSObject -Property $Properties
            $Return.PSObject.TypeNames.Insert(0, 'AM.OpCo.100')
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

#region Get-AMCountryOpCo

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
function Get-AMCountryOpCo
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
        [Parameter(Mandatory=$true, 
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
        $SQL = "SELECT TOP 1 * FROM SOEAdmin.ps.CountryOpCoAssetCount WHERE [CountryCode] = @CountryCode ORDER BY [AssetCount] DESC"

        $Cmd.Parameters.Add('CountryCode', [string]) | Out-Null
        $Cmd.Parameters['CountryCode'].Value = [string]$CountryCode

        $Cmd.CommandText = $SQL
        $Data = $Cmd.ExecuteReader()
        while ($Data.Read())
        {
            $Return = [int]$Data['OpCoID']
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

#region Get-AMReportOnRfC

<#
.Synopsis
    Query the ReportOnRfC table from the Asset Managament Database (ITAMWeb).

.Description
    The returned database is presented as an AM.ReportOnRfC.100 object and contains the folling fields:
    - SerialNo
    - ChangeRef
    - WhenCreated
    - WhenDeleted

.Example
    Get-AMReportOnRfC

    Returns all of the data from the table.
.Example
    Get-AMOpCo -ManagedOnly
        or
    Get-AMOpCo -ManagedOnly:$true


    Will return of the data from the table where the OpCo is managed by the contract.
.Example
    Get-AMOpCo -ManagedOnly:$false

    Returns of all of the data from the table where the OpCo is NOT managed by the contract, for Asset Management this means an asset assigned to this OpCo can't be invoiced.

.Example
    Get-AMOpCo -ID 2
        or
    Get-AMOpCo -Number '0002'

        ID          : 2
        Number      : 0002
        Name        : KDE BV
        Fullname    : 0002 KDE BV
        CountryCode : NL
        Managed     : True

.Example
    Get-AMOpCo -ID 9020 -ManagedOnly
        or
    Get-AMOpCo -Number '9020' -ManagedOnly

    Will not return data because this OpCo is not managed by the contract.

.Outputs
    [AM.OpCo.100] Object

.Notes
    --- Version history:
    Version 1.00 (2017-02-15, Kees Hiemstra)
    - Initial version.

.Component
    Asset Management.

.Role
    Querying Asset Management master data.

.Functionality
    Get information from Asset Manament database.

.Link
    Get-AMAsset

.Link
    Get-AMCountryOpCo

#>
function Get-AMReportOnRfC
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Low')]
    [Alias()]
    [OutputType([String])]
    Param
    (
        #Serial Number
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$false,
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [Alias('Department')]
        [string]
        $SerialNo,

        #Filter on managed OpCos only
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false)]
        [switch]
        $EarlyTermination
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
        $SQL = "SELECT * FROM ITAMData.dbo.acReportOnRfC"

        if ( $PSBoundParameters.ContainsKey('SerialNo') )
        {
            if ( $SQL -match 'WHERE' )
            {
                $SQL = "$SQL AND"
            }
            else
            {
                $SQL = "$SQL WHERE"
            }
            $SQL = "$SQL [SerialNo] = @SerialNo"

            $Cmd.Parameters.Add('SerialNo', [string]) | Out-Null
            $Cmd.Parameters['SerialNo'].Value = [string]$SerialNo

            Write-Verbose "SerialNo: $SerialNo"
        }

        if ( $PSBoundParameters.ContainsKey('EarlyTermination') )
        {
            if ( $SQL -match 'WHERE' )
            {
                $SQL = "$SQL AND"
            }
            else
            {
                $SQL = "$SQL WHERE"
            }

            if ( $EarlyTermination.ToBool() )
            {
                $SQL = "$SQL [ChangeRef] LIKE 'ET%'"
            }
            else
            {
                $SQL = "$SQL [ChangeRef] NOT LIKE 'ET%'"
            }

            Write-Verbose "EarlyTermination: $($EarlyTermination.ToBool())"
        }

        Write-Verbose $SQL

        $Cmd.CommandText = $SQL

        $Data = $Cmd.ExecuteReader()
        while ($Data.Read())
        {
            $Properties = [ordered]@{'SerialNo'    = [string]$Data['SerialNo']
                                     'Name'        = [string]$Data['OpCoName']
                                     'ChangeRef'   = [string]$Data['ChangeRef']
                                     'WhenCreated' = [datetime]$Data['DTCreation']
                                     'WhenDeleted' = if ( -not [system.DBNull]::Value.Equals( $Data['DTDeletion'] ) ) { [datetime]$Data['DTDeletion'] } else { $null }
                                     'Type'        = if ( [string]$Data['ChangeRef'] -like 'ET*' ) { 'Early termination' } else { 'Stolen or missing' }
                                    }
            $Return = New-Object -TypeName PSObject -Property $Properties
            $Return.PSObject.TypeNames.Insert(0, 'AM.ReportOnRfC.100')
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

