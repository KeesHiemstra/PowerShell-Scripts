<#
.Synopsis
   Create an SQL connection string.
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Get-ConnectionString
{
    [CmdletBinding()]
    [OutputType([string])]
    Param
    (
        # Server name (e.g. sbServ1) or Database server instance name (e.g. dbClust1\Inst2)
        [Parameter(Position=0)]
        [string] $ServerInstance = '(Local)',

        # Name of the database
        [string] $DatabaseName,

        # SQL user name
        [string] $UserID,

        # Password of the SQL user
        [string] $Password,

        # SLQ connection provider
        [string] $Provider,

        # SQL connection provider for Excel, will overwrite the -Provider parameter
        [switch] $ProviderForExcel
    )

    Begin
    {
        [string] $Result = ""
    }
    Process
    {
        if ($ProviderForExcel) { $Provider = "SQLOLEDB.1" }

        if (-not [string]::IsNullOrWhiteSpace($Provider)) { $Result += "Provider=$Provider;" }

        $Result += "Data Source=$ServerInstance;"

        if (-not [string]::IsNullOrWhiteSpace($UserID))
        {
            $Result += "Persist Security Info=True;User ID=$UserID;"
            if (-not [string]::IsNullOrWhiteSpace($Password)) { $Result += "Password=$Password;" }
        }
        else { $Result += "Integrated Security=True;" }

        if (-not [string]::IsNullOrWhiteSpace($DatabaseName)) { $Result += "Initial Catalog=$DatabaseName;" }
    }
    End
    {
        return $Result
    }
}

<#
.Synopsis
   Create SQL connection object.
.DESCRIPTION
   Creation an SQL System.Data.SqlClient.SqlConnection object.
.EXAMPLE
   $Conn = Get-SQLConnection (Get-ConnectionString -DatabaseName "master")

   Creates a new SQL connection object to the database master on the local database server and assigns it to $Conn.
#>
function Get-SQLConnection
{
    [CmdletBinding()]
    [OutputType([System.Data.SqlClient.SqlConnection])]
    Param
    (
        # Connection string to SQL server, e.g. Get-ConnectionString -DatabaseName Master
        [Parameter(Mandatory=$true,
                   Position=0)]
        [string] $ConnectionString
    )

    Begin
    {
        Write-Verbose "Connection string: $ConnectionString"
        [System.Data.SqlClient.SqlConnection] $Result = New-Object System.Data.SqlClient.SqlConnection
    }
    Process
    {
        $Result.ConnectionString = $ConnectionString
    }
    End
    {
        return $Result
    }
}

<#
.Synopsis
   Create SQL cmd object.
.DESCRIPTION
   Creation an SQL System.Data.SqlClient.SqlCmd object.
.EXAMPLE
   $Cmd = Get-SQLCommand (Get-SQLConnection (Get-ConnectionString -DatabaseName "msdb")) -Query 'SELECT [Name] FROM msdb.dbo.sysjobs'

   Creates a new SQL command object to the database master on the local database server with a query and assigns it to $Cmd.
#>
function Get-SQLCommand
{
    [CmdletBinding()]
    [OutputType([System.Data.SqlClient.SqlCommand])]
    Param
    (
        # Connection string to SQL server, e.g. Get-ConnectionString
        [Parameter(Mandatory=$true,
                   Position=0)]
        [System.Data.SqlClient.SqlConnection] $ConnectionObject,

        # SQL that will be assigned to the command.
        [string] $Query
    )

    Begin
    {
        if ($ConnectionObject.State -eq 'Closed') { $ConnectionObject.Open() }
        
        [System.Data.SqlClient.SqlCommand] $Result = New-Object System.Data.SqlClient.SqlCommand
        $Result.CommandTimeOut = 120
        $Result.Connection = $ConnectionObject
    }
    Process
    {
        if (-not [string]::IsNullOrWhiteSpace($Query)) { $Result.CommandText = $Query }
    }
    End
    {
        return $Result
    }
}

<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>

function Start-SQLJob
{
    [CmdletBinding()]
    #[OutputType([int])]
    Param
    (
        # Name of the SQL job
        [Parameter(Mandatory=$true,
                   Position=0)]
        [string] $JobName,

        # Connection string to SQL server, e.g. Get-ConnectionString -DatabaseName Master
        [string] $ConnectionString
    )

    Begin
    {
    }
    Process
    {
        $Conn = (Get-SQLConnection -ConnectionString $ConnectionString)
        $Cmd = Get-SQLCommand $Conn -Query "EXEC msdb.dbo.sp_start_job '$JobName'"
        $Result = $Cmd.ExecuteNonQuery()
        $Conn.Close()
    }
    End
    {
        return $Result
    }
}

Start-SQLJob -JobName 'ITAM - ! Import AssetCenter data' -ConnectionString (Get-ConnectionString -ServerInstance "DEMBMCAPS032SQ2\Prod_2")
