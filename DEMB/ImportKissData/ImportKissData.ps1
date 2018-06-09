<#
    ImportKissData.ps1

    Read the data from the IDM import and export files and store those into the SOEAdmin database.

    ===Version history
    Version 1.11 (2017-01-26, Kees Hiemstra)
    - Bug fix: In PowerShell 3.0 Import-Csv with a single column and empty field, the line is not skipped but can't process either.
    Version 1.10 (2017-01-24, Kees Hiemstra)
    - Bug fix: Prepare the culture to save the date/time in an exchangable way (yyyy-MM-dd HH:mm:ss).
    - Added logging.
    Version 1.00 (2017-01-15, Kees Hiemstra)
    - Initial version.
#>

$LogFile = 'D:\Scripts\ImportKissData\ImportKissData.log'

#region LogFile 1.20
$Error.Clear()

#Use: $Error.RemoveAt(0) to remove an error or warning captured in a try - catch statement

#Define initial log file variables
$ScriptName = $MyInvocation.MyCommand.Name.Replace(".ps1", '')
$MailSMTP = "smtp.corp.demb.com"
$MailSender = "$($ScriptName.Replace(' ', '.'))@JDECoffee.com"
$MailErrorTo = "Kees.Hiemstra@hpe.com"

$LogStart = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")

#The error mail to home will be send if $LogWarning is true but no other errors were reported.
#The subject of the mail will indicate Warning instead of Error.
$LogWarning = $false

#$VerbosePreference = 'SilentlyContinue' #Verbose off
#$VerbosePreference = 'Continue' #Verbose on

#Create the log file if the file does not exits else write today's first empty line as batch separator
if (-not (Test-Path $LogFile))
{
    New-Item $LogFile -ItemType file | Out-Null
}
else 
{
    Add-Content -Path $LogFile -Value "---------- --------"
}

function Write-Log([string]$Message)
{
    $LogMessage = "{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Add-Content -Path $LogFile -Value $LogMessage
    if ($VerbosePreference -eq 'Continue') { Write-Host $LogMessage }
    Write-Host $LogMessage
}

#Finish script and report on $Error
function Stop-Script([bool]$WithError)
{
    if ( $WithError )
    {
        Write-Log -Message "Script stopped with an initiated error"
    }
    else
    {
        Write-Log -Message "Script ended normally"
    }

    if ( $Error.Count -gt 0 -or $WithError -or $LogWarning)
    {
        if ( $Error.Count -gt 0 -or $WithError )
        {
            $MailErrorSubject = "Error in $ScriptName on $($Env:ComputerName)"

            if ( $Error.Count -gt 0 )
            {
                $MailErrorBody = "The script $ScriptName has reported the following error(s):`n`n"
                $MailErrorBody += $Error | Out-String
            }
            else
            {
                $MailErrorBody = "The script $ScriptName has reported error(s) in the log file."
            }
        }
        else
        {
            $MailErrorSubject = "Warning in $ScriptName on $($Env:ComputerName)"
            $MailErrorBody = "The script $ScriptName has reported warning(s) in the log file."
        }

        $MailErrorBody += "`n`n--- LOG FILE (extract) ----------------`n"
        $MailErrorBody += (Get-Content $LogFile | Where-Object { $_.SubString(0, 19) -ge $LogStart }) -join "`n"
        try
        {
            Send-MailMessage -SmtpServer $MailSMTP -From $MailSender -To $MailErrorTo -Subject $MailErrorSubject -Body $MailErrorBody -ErrorAction Stop
            Write-Log -Message "Sent error mail to home"
        }
        catch
        {
            Write-Log -Message "Retry sending the message after 15 seconds"
            Start-Sleep -Seconds 15

            try
            {
                Send-MailMessage -SmtpServer $MailSMTP -From $MailSender -To $MailErrorTo -Subject $MailErrorSubject -Body $MailErrorBody -ErrorAction Stop
                Write-Log -Message "Sent error mail to home after a retry"
            }
            catch
            {
                Write-Log -Message "Unable to send error mail to home"
            }
        }
    }

    Exit
}

#This function write the error to the logfile and exit the script
function Write-Break([string]$Message)
{
    Write-Log -Message $Message
    Write-Error -Message $Message
    Stop-Script -WithError $true
}

Write-Log -Message "Script started ($($env:USERNAME))"
#endregion

$ConnectionString = 'Trusted_Connection=True;Initial Catalog=SOEAdmin;Data Source=DEMBMCAPS032SQ2.corp.demb.com\Prod_2'

#Prepare the culture to save the date/time in the scv in such way that it can be read back properly
$CurrentThread = [System.Threading.Thread]::CurrentThread
$Culture = [CultureInfo]::InvariantCulture.Clone()
$Culture.DateTimeFormat.ShortDatePattern = 'yyyy-MM-dd'
$Culture.DateTimeFormat.ShortTimePattern = 'HH:mm:ss'
$CurrentThread.CurrentCulture = $Culture
$CurrentThread.CurrentUICulture = $Culture

$SQLFileInsert = @"
IF NOT EXISTS (SELECT 1 FROM SOEAdmin.uaa.KissFile WHERE [Name] = @Name)
    INSERT INTO SOEAdmin.uaa.KissFile ([Name], [DTWrite])
    VALUES (@Name, @DTWrite);

SELECT [ID]
FROM SOEAdmin.uaa.KissFile
WHERE [Name] = @Name
"@

#region Get-KissFileID

<#
.SYNOPSIS
    Get the Kiss FileID for the given file name.
.DESCRIPTION
    Query the uaa.KissFile table fore FileID of the given file name. The file name will be added if the name does not exist yet.
.EXAMPLE
.INPUTS
    [string[]] or [[System.IO.FileInfo]
.OUTPUTS
    PSObject for imporing the data.
.NOTES
    ===Version history
    Version 1.00 (2017-01-15, Kees Hiemstra)
    - Initial version.
#>
function Get-KissFileID
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    Param
    (
        #Object from Get-ChildItem
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$false,
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [System.IO.FileInfo[]]
        $File,

        #File name
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$false,
                   Position=0,
                   ParameterSetName='Parameter Set 2')]
        [String[]]
        $Path
    )

    Begin
    {
        $Conn = New-Object -TypeName System.Data.SqlClient.SqlConnection
        $Conn.ConnectionString = $ConnectionString
        try
        {
            $Conn.Open()
            Write-Verbose "Database connection is open"
        }
        catch
        {
            throw "Failed connection with connection string ($ConnectionString)"
        }
    }
    Process
    {
        if ( $PSBoundParameters.ContainsKey('Path') )
        {
            foreach ( $item in $File )
            {
                $File += Get-ChildItem -Path $Item
            }
        }

        foreach ( $Item in $File )
        {
            $Properties = [ordered]@{'Name'          = [string]   $Item.Name
                                     'LastWriteTime' = [datetime] $Item.LastWriteTimeUtc
                                     'FileID'        = [int]      -1
                                     'FullName'      = [string]   $Item.FullName
                                    }
            $Result = New-Object -TypeName PSObject -Property $Properties

            $Cmd = New-Object -TypeName System.Data.SqlClient.SqlCommand
            $Cmd.Connection = $Conn
            $Cmd.CommandText = $SQLFileInsert

            $Cmd.Parameters.Clear()
            $Cmd.Parameters.Add('Name', [string]) | Out-Null
            $Cmd.Parameters.Add('DTWrite', [datetime]) | Out-Null

            $Cmd.Parameters['Name'].Value = $Result.Name
            $Cmd.Parameters['DTWrite'].Value = $Result.LastWriteTime

            $Data = $Cmd.ExecuteReader()
            while ($Data.Read())
            {
                $Result.FileID = [int]$Data['ID']
            }
            $Data.Close()

            Write-Output $Result
        }
    }
    End
    {
        $Conn.Close()
    }
}

#endregion

#region Import-KissData

<#
.SYNOPSIS
    Import the Kiss provision data from the given file name.
.DESCRIPTION
.EXAMPLE
.INPUTS
    PSObject from Get-KissFileID
.OUTPUTS
    None
.NOTES
    ===Version history
    Version 1.00 (2017-01-15, Kees Hiemstra)
    - Initial version.
#>
function Import-KissData
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    Param
    (
        #Object from Get-KissFileID
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$false,
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [PSObject]
        $KissFileObject,

        #Table name
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false,
                   ValueFromRemainingArguments=$false,
                   Position=1,
                   ParameterSetName='Parameter Set 1')]
        [string]
        $Table
    )

    Begin
    {
        $Conn = New-Object -TypeName System.Data.SqlClient.SqlConnection
        $Conn.ConnectionString = $ConnectionString
        try
        {
            $Conn.Open()
            Write-Verbose "Database connection is open"
        }
        catch
        {
            throw "Failed connection with connection string ($ConnectionString)"
        }
    }
    Process
    {
        $Data = Import-Csv -Path $KissFileObject.FullName -Delimiter ';'

        $Cmd = New-Object -TypeName System.Data.SqlClient.SqlCommand
        $Cmd.Connection = $Conn

        $LineCount = 1
        foreach ( $Item in $Data )
        {
            $LineCount++
            $Insert = "INSERT INTO uaa.$Table ([FileID]"
            $Values = 'VALUES (@FileID'

            $Cmd.Parameters.Clear()
            $Cmd.Parameters.Add('FileID', [string]) | Out-Null
            $Cmd.Parameters['FileID'].Value = $KissFileObject.FileID.ToString()
            $HasFields = $false
            foreach ( $Field in ($Data[0] | Get-Member -MemberType NoteProperty).Name)
            {
                if ( -not [string]::IsNullOrEmpty($Item.$Field) )
                {
                    $HasFields = $true
                    $FieldName = $Field
                    if ( $Table -eq 'Deprovisioning' -and $FieldName -eq 'SAMAccoutName')
                    {
                        #In the first version the field name was misspelled
                        $FieldName = 'SAMAccountName'
                        Write-Verbose 'SAMAccountName field is corrected'
                    }
                    if ( $Table -eq 'Deprovisioning' -and $FieldName -eq 'SAMAccountName' -and [string]::IsNullOrEmpty($Item.$Field) )
                    {
                        #The field can be empty if the account is not bound
                        $Item.$Field = '<Empty>'
                        Write-Log 'SAMAccountName field is empty'
                    }
                    $Insert += ", [$FieldName]"
                    $Values += ", @$FieldName"
                    $Cmd.Parameters.Add($FieldName, [string]) | Out-Null
                    $Cmd.Parameters[$FieldName].Value = $Item.$Field
                    Write-Verbose "$($FieldName): $($Cmd.Parameters[$FieldName].Value) [length]: $(($Item.$Field).Length)"
                }
            }#foreach property in $Data
            if ( $Table -eq 'Deprovisioning' )
            {
                #An import should only have one record per SAMAccountName
                $Cmd.CommandText = "IF NOT EXISTS (SELECT 1 FROM uaa.$Table WHERE [FileID] = @FileID AND [SAMAccountName] = @SAMAccountName) $Insert) $Values);"
            }
            else
            {
                #An import should only have one record per EmployeeID
                $Cmd.CommandText = "IF NOT EXISTS (SELECT 1 FROM uaa.$Table WHERE [FileID] = @FileID AND [EmployeeID] = @EmployeeID) $Insert) $Values);"
            }
            try
            {
                if ( $HasFields )
                {
                    $Cmd.ExecuteNonQuery() | Out-Null
                }
            }
            catch
            {
                Write-Break -Message "Can't import $($KissFileObject.FullName) to $Table at line $LineCount"
            }
        }#foreach $Data
    }
    End
    {
        $Conn.Close()
    }
}

#endregion

$ImportPath = "$($MyInvocation.MyCommand.Definition -replace(".ps1$",".csv"))"
$Process = Import-Csv -Path $ImportPath

foreach ( $Import in $Process )
{
    if ( [string]::IsNullOrEmpty($Import.LastWriteTime) )
    {
        $Import.LastWriteTime = (Get-Date -Year 1980 -Month 1 -Day 1).Date
    }
    Get-ChildItem -Path $Import.Path |
        Where-Object { $_.LastWriteTime -gt ($Import.LastWriteTime) } |
        Sort-Object LastWriteTime |
        ForEach-Object {
            Write-Log "Processing: $($_.Name)"
            Get-KissFileID -File $_ |
            Import-KissData -Table $Import.Table -Verbose

            $Import.LastWriteTime = $_.LastWriteTime
            $Process | Export-Csv -Path $ImportPath -NoTypeInformation -Force
        }
}

Stop-Script
