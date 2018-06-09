$ConnectionString = 'Trusted_Connection=True;Data Source=DEMBMCAPS032SQ2.corp.demb.com\Prod_2'

#region Get-IDMProvisioning

<#
.SYNOPSIS
    Get IDM binding data from the database.

.DESCRIPTION
    The Get-IDMProvisioning function searches the binding table with the provided EmployeeID and returns the data in chronological order. 

    The data is collected by the ImportKissData script and is stored in the SOEAmin database.

    You do already need to have access to this database in order to read the data.

.EXAMPLE
    > Get-IDMProvisioning -EmployeeID 20200209


    EmployeeID         : 20200209
    Sn                 : Rob
    GivenName          : Erne
    Company            : 6864 JDE NO AS
    l                  : Bergen
    c                  : NO
    Co                 : Norway
    Comment            : Active
    EmployeeType       : Active Employee
    Manager            : 
    DisplayName        : Erne, Rob
    ExtensionAttrite11 : 
    KissFile           : IDM_new_users_2016-09-20-0905.csv
    KissWritten        : 2016-09-20 06:55:00

.INPUTS
    [string]

.OUTPUTS
    PSObject (IDM.Provisioning.100)

.NOTES
    === Version history
    Version 1.00 (2017-02-08, Kees Hiemstra)

.COMPONENT
    The component this cmdlet belongs to the IDM module.

.ROLE
    The role this cmdlet belongs to User Account Administration.

.FUNCTIONALITY
    Database access is handled by standard .Net libraries.

.LINK
    Get-IDMProvisioning

.LINK
    Get-IDMBinding

.LINK
    Get-IDMDeprovisioning

#>
function Get-IDMProvisioning
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$false,
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([String])]
    Param
    (
        #EmployeeID to search for
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [string]
        $EmployeeID,

        #First name to search for
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 2')]
        [string]
        $FirstName,

        #Last name to search for
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=1,
                   ParameterSetName='Parameter Set 2')]
        [string]
        $LastName
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

        if ( $PSBoundParameters.ContainsKey('EmployeeID') )
        {
            $SQL = 'SELECT * FROM SOEAdmin.uaa.Provisioning AS P JOIN SOEAdmin.uaa.KissFile AS F ON P.[FileID] = F.[ID] WHERE [EmployeeID] = @EmployeeID '
            $SQL += 'ORDER BY [DTWrite], [Name]'

            $Cmd.Parameters.Add('EmployeeID', [string]) | Out-Null
            $Cmd.Parameters['EmployeeID'].Value = $EmployeeID.Trim().PadLeft(8, "0")
        }
        elseif ( $PSBoundParameters.ContainsKey('FirstName') -and $PSBoundParameters.ContainsKey('LastName') )
        {
            $SQL = 'SELECT * FROM SOEAdmin.uaa.Provisioning AS P JOIN SOEAdmin.uaa.KissFile AS F ON P.[FileID] = F.[ID] WHERE [GivenName] = @FirstName AND [sn] = @LastName '
            $SQL += 'ORDER BY [DTWrite], [Name]'

            $Cmd.Parameters.Add('FirstName', [string]) | Out-Null
            $Cmd.Parameters['FirstName'].Value = $FirstName.Replace('*', '%').Replace('?', '_')
            Write-Verbose "FirstName: $FirstName"

            $Cmd.Parameters.Add('LastName', [string]) | Out-Null
            $Cmd.Parameters['LastName'].Value = $LastName.Replace('*', '%').Replace('?', '_')
            Write-Verbose "LastName: $LastName"
        }

        $Cmd.CommandText = $SQL

        Write-Verbose $SQL

        $Data = $Cmd.ExecuteReader()
        while ($Data.Read())
        {
            $Properties = [ordered]@{'EmployeeID'         = [string]$Data['EmployeeID']
                                     'Sn'                 = [string]$Data['Sn']                                     
                                     'GivenName'          = [string]$Data['GivenName']
                                     'StreetAddress'      = [string]$Data['StreetAddress']
                                     'Company'            = [string]$Data['Company']
                                     'l'                  = [string]$Data['l']                                     
                                     'c'                  = [string]$Data['c']                                     
                                     'Co'                 = [string]$Data['Co']
                                     'Comment'            = [string]$Data['Comment']
                                     'EmployeeType'       = [string]$Data['EmployeeType']
                                     'Manager'            = [string]$Data['Manager']
                                     'DisplayName'        = [string]$Data['DisplayName']
                                     'ExtensionAttrite11' = [string]$Data['ExtensionAttrite11']
                                     'KissFile'           = [string]$Data['Name']
                                     'KissWritten'        = [datetime]$Data['DTWrite']
                                    }
            $Return = New-Object -TypeName PSObject -Property $Properties
            $Return.PSObject.TypeNames.Insert(0, 'IDM.Provisioning.100')
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

#region Get-IDMBinding

<#
.SYNOPSIS
    Get IDM binding data from the database.

.DESCRIPTION
    The Get-IDMBinding function searches the binding table with the provided EmployeeID and returns the data in chronological order. 

    The data is collected by the ImportKissData script and is stored in the SOEAmin database.

    You do already need to have access to this database in order to read the data.

.EXAMPLE
    > Get-IDMBinding -EmployeeID 20200209

    EmployeeID        : 20200209
    Mail              : Rob.Erne@JDEcoffee.com
    SAMAccountName    : Rob.Erna
    UserPrincipalName : rob.erne@jdecoffee.com
    DistinguishedName : CN=erne.rob,OU=Managed Users,DC=corp,DC=demb,DC=com
    KissFile          : 2016-09-20-AD_upload_2016-09-20-1600.csv
    KissWritten       : 2016-09-20 14:01:00

    EmployeeID        : 20200209
    Mail              : Rob.Erne@JDEcoffee.com
    SAMAccountName    : Rob.Erna
    UserPrincipalName : rob.erne@jdecoffee.com
    DistinguishedName : CN=erne.rob,OU=Managed Users,DC=corp,DC=demb,DC=com
    KissFile          : 2016-10-19-AD_upload_2016-10-19-1500.csv
    KissWritten       : 2016-10-19 13:01:00

.INPUTS
    [string]

.OUTPUTS
    PSObject (IDM.Binding.100)

.NOTES
    === Version history
    Version 1.00 (2017-02-08, Kees Hiemstra)

.COMPONENT
    The component this cmdlet belongs to the IDM module.

.ROLE
    The role this cmdlet belongs to User Account Administration.

.FUNCTIONALITY
    Database access is handled by standard .Net libraries.

.LINK
    Get-IDMProvisioning

.LINK
    Get-IDMBinding

.LINK
    Get-IDMDeprovisioning

#>
function Get-IDMBinding
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$false,
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([String])]
    Param
    (
        #EmployeeID to search for
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [string]
        $EmployeeID
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
        $SQL = 'SELECT * FROM SOEAdmin.uaa.Binding AS B JOIN SOEAdmin.uaa.KissFile AS F ON B.[FileID] = F.[ID] WHERE [EmployeeID] = @EmployeeID '
        $SQL += 'ORDER BY [DTWrite], [Name]'

        $Cmd.Parameters.Add('EmployeeID', [string]) | Out-Null
        $Cmd.Parameters['EmployeeID'].Value = $EmployeeID.Trim().PadLeft(8, "0")

        $Cmd.CommandText = $SQL

        $Data = $Cmd.ExecuteReader()
        while ($Data.Read())
        {
            $Properties = [ordered]@{'EmployeeID'        = [string]$Data['EmployeeID']
                                     'Mail'              = [string]$Data['Mail']                                     
                                     'SAMAccountName'    = [string]$Data['SAMAccountName']
                                     'UserPrincipalName' = [string]$Data['UserPrincipalName']
                                     'DistinguishedName' = [string]$Data['DistinguishedName']
                                     'KissFile'          = [string]$Data['Name']
                                     'KissWritten'       = [datetime]$Data['DTWrite']
                                    }
            $Return = New-Object -TypeName PSObject -Property $Properties
            $Return.PSObject.TypeNames.Insert(0, 'IDM.Binding.100')
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

#region Get-IDMDeprovisioning

<#
.SYNOPSIS
    Get IDM deprovisioning data from the database.

.DESCRIPTION
    The Get-IDMDeprovisioning function searches the deprovisioning table with the provided SAMAccountName and returns the data in chronological order. 

    The data is collected by the ImportKissData script and is stored in the SOEAmin database.

    You do already need to have access to this database in order to read the data.

.EXAMPLE
    > Get-IDMDeprovisioning siem.los

    SAMAccountName EmployeeID KissFile                           KissWritten        
    -------------- ---------- --------                           -----------        
    siem.los       05013067   disabled_users_2015-08-13-1555.csv 2015-08-13 13:55:00
    siem.los                  disabled_users_2015-09-23-0355.csv 2015-09-23 15:19:00
    siem.los                  disabled_users_2015-09-23-0555.csv 2015-09-23 15:19:00

.INPUTS
    [string]

.OUTPUTS
    PSObject (IDM.Deprovisioning.100)

.NOTES
    === Version history
    Version 1.00 (2017-02-08, Kees Hiemstra)

.COMPONENT
    The component this cmdlet belongs to the IDM module.

.ROLE
    The role this cmdlet belongs to User Account Administration.

.FUNCTIONALITY
    Database access is handled by standard .Net libraries.

.LINK
    Get-IDMProvisioning

.LINK
    Get-IDMBinding

.LINK
    Get-IDMDeprovisioning

#>
function Get-IDMDeprovisioning
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$false,
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([String])]
    Param
    (
        #SAMAccountName to search for
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [string]
        $SAMAccountName
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
        $SQL = 'SELECT * FROM SOEAdmin.uaa.Deprovisioning AS D JOIN SOEAdmin.uaa.KissFile AS F ON D.[FileID] = F.[ID] WHERE [SAMAccountName] = @SAMAccountName '
        $SQL += 'ORDER BY [DTWrite], [Name]'

        $Cmd.Parameters.Add('SAMAccountName', [string]) | Out-Null
        $Cmd.Parameters['SAMAccountName'].Value = $SAMAccountName

        $Cmd.CommandText = $SQL

        $Data = $Cmd.ExecuteReader()
        while ($Data.Read())
        {
            $Properties = [ordered]@{'SAMAccountName' = [string]$Data['SAMAccountName']
                                     'EmployeeID'     = [string]$Data['EmployeeID']
                                     'KissFile'       = [string]$Data['Name']
                                     'KissWritten'    = [datetime]$Data['DTWrite']
                                    }
            $Return = New-Object -TypeName PSObject -Property $Properties
            $Return.PSObject.TypeNames.Insert(0, 'IDM.Deprovisioning.100')
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
