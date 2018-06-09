#region Save-SOEObject

<#
.SYNOPSIS
    Saves SOE object data into the provided database.
.DESCRIPTION
    The SOE object have a type name like SOE.<TableName>.<Version>

    The cmdlet takes the data from the object and stores it in the given <TableName>.

    If the ImputObject contains the fields [ComputerName] and [DTCollection] the cmdlets deletes the exising data in the table based on these fields.
    This can be overwritten by the parameter -Cumulative.

    Fields that don't exist in the table are not provided and null values are stored as NULL.
.EXAMPLE
    Get-SOEComputerModel -ComputerName . | Save-SOEObject

    Saves the collected Model data into table ComputerModel.
.EXAMPLE
    Get-SOEComputerModel -ComputerName . | Get-SOEComputerRunKeys | Save-SOEObject

    Saves the collected runkeys data from the current console user into the table ComputerRunKeys.
.EXAMPLE
    Get-SOEComputerBSoDEvents -ComputerName . | Save-SOEObject -PassThrough

    Saves the collected blue screen of death data into the table ComputerBSoDEvents and passes the received object through the pipeline.
.INPUTS
    Any SOE object SOE.<name>.Version
.OUTPUTS

.NOTES
    --- Version history:
    Version 1.10 (2015-12-22, Kees Hiemstra)
    - Added logging to the database table.
    Version 1.10 (2015-12-21, Kees Hiemstra)
    - Added parameter Cumulatively to overwrite the deletion of existing data.
    - Added the deletion of existing data based on [ComputerName] and [DTCollection].
    - Added parameter ConnectionString to overwrite the local connection sting.
    - Updated help.
    Version 1.00 (2015-12-19, Kees Hiemstra)
    - Inital version.
.COMPONENT

.ROLE

.FUNCTIONALITY

.LINK
    Get-SOEComputerModel

.LINK
    Get-SOEComputerOS

.LINK
	Get-SOEComputerBSoDEvents

.LINK
	Get-SOEComputerNICDetails

.LINK
	Get-SOEComputerSignedDrivers

.LINK
	Get-SOEComputerRunKeys

.LINK
	Get-SOEComputerPatches

.LINK
	Get-SOEComputerBIOS

#>
function Save-SOEObject
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    [OutputType([String])]
    Param
    (
        #Object which data needs to be saved
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [Alias("SOEObject")] 
        [object[]]
        $InputObject,

        #Connection string to the database
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false)]
        [Alias("cs")] 
        [string]
        $ConnectionString,

        #Don't delete existing data.
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false)]
        [switch]
        $Cumulative,

        #Pass the input object through to the next pipeline.
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false)]
        [switch]
        $PassThrough
    )

    Begin
    {
        if ([string]::IsNullOrEmpty($ConnectionString))
        {
            $ConnectionString = 'Trusted_Connection=True;Data Source=(Local);Initial Catalog=SOEAdmin'
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

        $TableFields = @{}
        $SQL = ''

        Write-Verbose "Successfully connected to the database"

        $TableCheck = $false
    }
    Process
    {
        foreach ($Object in $InputObject)
        {
            if (-not $TableCheck)
            {
                $TypeName = $Object | Get-Member | Select-Object -ExpandProperty TypeName -First 1
                $TypeNameParts = $TypeName.Split('.')
            
                if ($TypeNameParts.Count -lt 3 -or $TypeNameParts[-3] -ne 'SOE')
                {
                    throw "Not an SOE object [$($TypeNameParts[-3])]"
                }
                $TableName = $TypeNameParts[-2]
                Write-Verbose "Extracted table name: $TableName"

                $Cmd = New-Object -TypeName System.Data.SqlClient.SqlCommand
                $Cmd.Connection = $Conn
                $Cmd.CommandText = "SELECT COUNT(*) AS 'Count' FROM INFORMATION_SCHEMA.TABLES WHERE [Table_Schema] = 'dbo' AND [Table_Name] = '$TableName'"
                $Data = $Cmd.ExecuteReader()
                if ($Data.Read())
                {
                    $RowCount = $Data.GetValue(0)
                    $Data.Close()
                    if ($RowCount -ne 1)
                    {
                        throw "Table [$TableName] doesn't exist"
                    }
                }
                else
                {
                    throw "Table [$TableName] doesn't exist"
                }

                $Cmd.CommandText = "SELECT [Column_Name], [Data_Type] FROM INFORMATION_SCHEMA.COLUMNS WHERE [Table_Schema] = 'dbo' AND [Table_Name] = '$TableName'"
                $Data = $Cmd.ExecuteReader()
                while ($Data.Read())
                {
                    $TableFields.Add($Data.GetValue(0), $Data.GetValue(1))

                    Write-Debug "Field [$($Data.GetValue(0))] with type $($Data.GetValue(1))"
                }
                $Data.Close()

                $Properties = $Object | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name |
                    Where-Object { $_ -in $TableFields.Keys }

                if ($Properties -contains 'ComputerName' -and $Properties -contains 'DTCollection' -and -not $Cumulative -and $pscmdlet.ShouldProcess("$TableName", "Delete existing data from"))
                {
                    $Cmd.CommandText = "DELETE FROM dbo.$TableName WHERE [ComputerName] = @ComputerName AND [DTCollection] < @DTCollection"

                    $Cmd.Parameters.Clear()
                    $Cmd.Parameters.Add('ComputerName', [string]) | Out-Null
                    $Cmd.Parameters['ComputerName'].Value = ($Object.ComputerName)
                    $Cmd.Parameters.Add('DTCollection', [datetime]) | Out-Null
                    $Cmd.Parameters['DTCollection'].Value = [datetime]($Object.DTCollection)
                    $Cmd.ExecuteNonQuery() | Out-Null
                }

                $Cmd.CommandText = "INSERT INTO dbo.$TableName ([$($Properties -join '],[')])`nVALUES (@$($Properties -join ',@'))"

                $Cmd.Parameters.Clear()
                foreach ($Property in $Properties)
                {
                    switch -Regex ($TableFields[$Property])
                    {
                        "int|smallint|bigint" { $Cmd.Parameters.Add($Property, [string]) | Out-Null }
                        "char|varchar"        { $Cmd.Parameters.Add($Property, [string]) | Out-Null }
                        "date|smalldate"      { $Cmd.Parameters.Add($Property, [datetime]) | Out-Null }
                        "bit"                 { $Cmd.Parameters.Add($Property, [bool]) | Out-Null }
                        default               { Write-Verbose "Unknown field" }
                    }
                }

                $TableCheck = $true
            }

            if ($TableCheck -and $pscmdlet.ShouldProcess("$TableName", "Data to be stored in"))
            {
                foreach ($Property in $Properties)
                {
                    if ($Object."$Property" -eq $null)
                    {
                        $Cmd.Parameters[$Property].Value = [System.DBNull]::Value
                    }
                    else
                    {
                        switch -Regex ($TableFields[$Property])
                        {
                            "int|bigint|smallint" { $Cmd.Parameters[$Property].Value = ($Object."$Property").ToString() }
                            "char|varchar"        { $Cmd.Parameters[$Property].Value = ($Object."$Property") }
                            "date|smalldate"      { $Cmd.Parameters[$Property].Value = [datetime]($Object."$Property") }
                            "bit"                 { $Cmd.Parameters[$Property].Value = [bool]($Object."$Property") }
                            default               { Write-Verbose "Unknown field" }
                        }
                    }                    
                }#foreach field
                try
                {
                    $Cmd.ExecuteNonQuery() | Out-Null
                }
                catch
                {
                    $Message = "Error saving data in [$TableName]`n"
                    foreach ($Property in $Properties)
                    {
                        switch -Regex ($TableFields[$Property])
                        {
                            "int|bigint|smallint" { $Message += "$Property = $(($Object."$Property").ToString())`n" }
                            "char|varchar"        { $Message += "$Property ($(($Object."$Property").Length)) = $($Object."$Property")`n" }
                            "date|smalldate"      { $Message += "$Property = $(($Object."$Property").ToString('yyyy-MM-dd HH:mm:ss'))`n" }
                            "bit"                 { $Message += "$Property = $(($Object."$Property").ToString())`n" }
                            default               { $Message += "$Property is an unknown field type`n" }
                        }
                    }#foreach field
                    $Message += "`n---Error message---`n"
                    $Message += $Error | Out-String
                    Save-SOELog -Message $Message -Source 'Save-SOEObject' -LogType Error -ConnectionString $ConnectionString                    
                }
            }#store object?

            if ($PassThrough)
            {
                Write-Output $InputObject
            }
        }#foreach object
    }
    End
    {
        $Conn.Close()
    }
}

#endregion

#region Save-SOELog

<#
.SYNOPSIS
    Saves logging data into the provided database.
.DESCRIPTION

.EXAMPLE

.INPUTS

.OUTPUTS

.NOTES
    --- ToDo:
    - Add ConnectionString as parameter.

    --- Version history:
    Version 1.00 (2015-12-22, Kees Hiemstra)
    - Inital version.

    --- Extra information
    Use the following SQL statement to create a table in the database if the data needs to store the logging information.
    
    CREATE TABLE [dbo].[SOELog](
	    [ID] [int] IDENTITY(1,1) NOT NULL PRIMARY KEY,
	    [Source] [varchar](25) NOT NULL,
	    [LogType] [char](7) NOT NULL,
	    [Message] [varchar](1024) NOT NULL,
	    [DTCreation] [datetime] NOT NULL CONSTRAINT [DF_SOELog_DTCreation] DEFAULT (GETDATE())
	    )


.LINK
	Save-SOEObject

#>
function Save-SOELog
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    [OutputType([String])]
    Param
    (
        #Log message
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [string]
        $Message,

        #Name of the source
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position=1,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [string]
        $Source,

        #Type of log
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position=2,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [ValidateSet('Error', 'Warning', 'Info', 'Debug')]
        [string]
        $LogType,

        #Connection string to the database
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false,
                   Position=3)]
        [Alias("cs")] 
        [string]
        $ConnectionString
    )

    Process
    {
        if ($Source.Length -gt 25)
        {
            $Source = "$($Source.Substring(0, 22))..."
        }
        if ($Source.Length -gt 1024)
        {
            $Source = "$($Source.Substring(0, 1021))..."
        }

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
        $Cmd = New-Object -TypeName System.Data.SqlClient.SqlCommand
        $Cmd.Connection = $Conn
        $Cmd.CommandText = "INSERT INTO dbo.SOELog([Source], [LogType], [Message]) VALUES(@Source, @LogType, @Message)"

        $Cmd.Parameters.Clear()
        $Cmd.Parameters.Add('Source', [string]) | Out-Null
        $Cmd.Parameters.Add('LogType', [string]) | Out-Null
        $Cmd.Parameters.Add('Message', [string]) | Out-Null

        $Cmd.Parameters['Source'].Value = $Source
        $Cmd.Parameters['LogType'].Value = $LogType
        $Cmd.Parameters['Message'].Value = $Message

        $Cmd.ExecuteNonQuery() | Out-Null
    }
    End
    {
        $Conn.Close()
    }
}

#endregion

#region Export functions to model

Export-ModuleMember -Function Save-SOEObject

#endregion