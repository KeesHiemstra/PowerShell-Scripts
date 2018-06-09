<#
    MonitorADChanges.ps1
#>

. "$($MyInvocation.MyCommand.Definition.Replace('.ps1', '.Config.ps1'))"

#region LogFile
$ScriptName = $MyInvocation.MyCommand.Definition.Replace("$PSScriptRoot\", '').Replace(".ps1", '')
$LogFile = $MyInvocation.MyCommand.Definition -replace(".ps1$",".log")
$LogStart = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
$MailSMTP = "smtp.corp.demb.com"
$MailSender = "$($ScriptName.Replace(' ', '.'))@JDECoffee.com"
$MailErrorTo = "Kees.Hiemstra@hpe.com"

$Error.Clear()
$LogWarning = $false
$Conn = New-Object -TypeName System.Data.SqlClient.SqlConnection

#$VerbosePreference = 'SilentlyContinue' #Verbose off
#$VerbosePreference = 'Continue' #Verbose on

#Create the log file if the file does not exits else write today's first empty line as batch separator
if ( -not (Test-Path $LogFile) )
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
                Write-Log -Message $MailErrorBody
            }
        }
    }

    #Close the connection to the database
    if ( $Conn.State -notin ('broken', 'closed') )
    {
        try { $Conn.Close() } catch {}
    }

    if ( $WithError )
    {
        Write-Log -Message "Script stopped with an error"
    }
    else
    {
        Write-Log -Message "Script ended normally"
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

#region Open Database connection
$Conn.ConnectionString = $ConnectionString
try
{
    $Conn.Open()
    $Cmd = New-Object -TypeName System.Data.SqlClient.SqlCommand
    $Cmd.Connection = $Conn
}
catch
{
    Write-Break -Message "Failed connection with connection string: $ConnectionString"
}
#endregion

#region Get latest update times
$Cmd.CommandText = @"
SELECT MAX([DTCreation]) AS 'DTCreation',
    MAX([DTMutation]) AS 'DTMutation',
    MAX([DTDeletion]) AS 'DTDeletion'
FROM SOEAdmin.dbo.ADObject
"@

$Reader = $Cmd.ExecuteReader()
$Data  = New-Object "System.Data.DataTable"
$Data.Load($Reader)

$Date = $Data.DTCreation
if ( $Data.DTMutation -gt $Date ) { $Date = $Data.DTMutation }
If ( $Data.DTDeletion -gt $Date ) { $Date = $Data.DTDeletion }

$Date = $Date.ToLocalTime().AddHours(-1)

Write-Log -Message "Get AD data since: $($Date.ToString('yyyy-MM-dd HH:mm'))"
#endregion

#region Get deleted data from AD
#DistinguishedName is always deleted items
try
{
    $ADDeletion = Get-ADObject -Filter { IsDeleted -eq $true -and WhenChanged -ge $Date } -Properties Name, ObjectClass, ObjectGUID, ObjectSID, WhenChanged, WhenCreated -IncludeDeletedObjects -Credential (Get-SOECredential svc.uaa) -Server DEMBDCRS001 -ErrorAction Stop |
        Select-Object @{n='DeletedName'; e={ $_.Name.Split("`n")[0] }}, ObjectClass, ObjectGUID, ObjectSID, WhenChanged, WhenCreated
}
catch
{
    Write-Break "Get deleted data from AD failed: $($Error[0].Message)"
}

Write-Log -Message "Number of deletions: $($ADDeletion.Count)"
#endregion

#region Get updated data from AD
try
{
    $ADMutation = Get-ADObject -Filter { WhenChanged -ge $Date } -Properties Name, ObjectClass, ObjectGUID, ObjectSID, WhenCreated, WhenChanged, DistinguishedName -Credential (Get-SOECredential svc.uaa) -Server DEMBDCRS001 -ErrorAction Stop |
        Select-Object Name, ObjectClass, ObjectGUID, ObjectSID, DistinguishedName, WhenCreated, WhenChanged
}
catch
{
    Write-Break "Get updated data from AD failed: $($Error[0].Message)"
}

Write-Log -Message "Number of mutations: $($ADMutation.Count)"
#endregion

#region Save OU and containers to database
$Cmd.CommandText = "EXEC SOEAdmin.dbo.ADPartition_Store @Name, @PartitionName, @ObjectClass, @ObjectGUID, @DTCreation, @DTMutation"

foreach ($Partition in ($ADMutation | Where-Object {$_.ObjectClass -in ('Container', 'organizationalUnit')}))
{
    Write-Log -Message "Partition: $($Partition.DistinguishedName)"

    $Cmd.Parameters.Clear()
    $Cmd.Parameters.Add('Name', [string]) | Out-Null
    $Cmd.Parameters.Add('PartitionName', [string]) | Out-Null
    $Cmd.Parameters.Add('ObjectClass', [string]) | Out-Null
    $Cmd.Parameters.Add('ObjectGUID', [string]) | Out-Null
    $Cmd.Parameters.Add('DTCreation', [datetime]) | Out-Null
    $Cmd.Parameters.Add('DTMutation', [datetime]) | Out-Null

    $Cmd.Parameters['Name'].Value = $Partition.Name
    $Cmd.Parameters['PartitionName'].Value = $Partition.DistinguishedName
    $Cmd.Parameters['ObjectClass'].Value = $Partition.ObjectClass
    $Cmd.Parameters['ObjectGUID'].Value = $Partition.ObjectGUID
    $Cmd.Parameters['DTCreation'].Value = $Partition.WhenCreated.ToUniversalTime()
    $Cmd.Parameters['DTMutation'].Value = $Partition.WhenChanged.ToUniversalTime()

    $Cmd.ExecuteNonQuery() | Out-Null
}
#endregion

#region Save AD changes to database
$Cmd.CommandText = "EXEC SOEAdmin.dbo.ADObject_Store @Name, @SAMAccountName, @UserPrincipalName, @ObjectClass, @ObjectGUID, @DistinguishedName, @ObjectSID, @DTCreation, @DTMutation"

foreach ($Item in ($ADMutation))
{
#    Write-Log -Message "Object: $($Item.DistinguishedName)"

    $Cmd.Parameters.Clear()
    $Cmd.Parameters.Add('Name', [string]) | Out-Null
    $Cmd.Parameters.Add('SAMAccountName', [string]) | Out-Null
    $Cmd.Parameters.Add('UserPrincipalName', [string]) | Out-Null
    $Cmd.Parameters.Add('ObjectClass', [string]) | Out-Null
    $Cmd.Parameters.Add('ObjectGUID', [string]) | Out-Null
    $Cmd.Parameters.Add('DistinguishedName', [string]) | Out-Null
    $Cmd.Parameters.Add('ObjectSID', [string]).IsNullable | Out-Null
    $Cmd.Parameters.Add('DTCreation', [datetime]).IsNullable | Out-Null
    $Cmd.Parameters.Add('DTMutation', [datetime]).IsNullable | Out-Null

    $Cmd.Parameters['Name'].Value = $Item.Name
    $Cmd.Parameters['SAMAccountName'].Value = if ( $Item.SAMAccountName -ne $null ) { $Item.SAMAccountName } else { [System.Convert]::DBNull }
    $Cmd.Parameters['UserPrincipalName'].Value = if ( $Item.UserPrincipalName -ne $null ) { $Item.UserPrincipalName } else { [System.Convert]::DBNull }
    $Cmd.Parameters['ObjectClass'].Value = $Item.ObjectClass
    $Cmd.Parameters['ObjectGUID'].Value = $Item.ObjectGUID
    $Cmd.Parameters['DistinguishedName'].Value = $Item.DistinguishedName
    $Cmd.Parameters['ObjectSID'].Value = if ( $Item.ObjectSID -ne $null ) { $Item.ObjectSID.ToString() } else { [System.Convert]::DBNull }
    $Cmd.Parameters['DTCreation'].Value = if ( $Item.WhenCreated -ne $null ) { $Item.WhenCreated.ToUniversalTime() } else { [System.Convert]::DBNull }
    $Cmd.Parameters['DTMutation'].Value = if ( $Item.WhenChanged -ne $null ) { $Item.WhenChanged.ToUniversalTime() } else { [System.Convert]::DBNull }

    $Cmd.ExecuteNonQuery() | Out-Null
}
#endregion

#region Save AD deletion to database
$Cmd.CommandText = "EXEC SOEAdmin.dbo.ADObject_Delete @Name, @SAMAccountName, @UserPrincipalName, @ObjectClass, @ObjectGUID, @ObjectSID, @DTCreation, @DTDeletion"

foreach ($Item in ($ADDeletion))
{
#    Write-Log -Message "Deletion: $($Item.DeletedName)"

    $Cmd.Parameters.Clear()
    $Cmd.Parameters.Add('Name', [string]) | Out-Null
    $Cmd.Parameters.Add('SAMAccountName', [string]) | Out-Null
    $Cmd.Parameters.Add('UserPrincipalName', [string]) | Out-Null
    $Cmd.Parameters.Add('ObjectClass', [string]) | Out-Null
    $Cmd.Parameters.Add('ObjectGUID', [string]) | Out-Null
    $Cmd.Parameters.Add('ObjectSID', [string]).IsNullable | Out-Null
    $Cmd.Parameters.Add('DTCreation', [datetime]).IsNullable | Out-Null
    $Cmd.Parameters.Add('DTDeletion', [datetime]).IsNullable | Out-Null

    $Cmd.Parameters['Name'].Value = $Item.DeletedName
    $Cmd.Parameters['SAMAccountName'].Value = if ( $Item.SAMAccountName -ne $null ) { $Item.SAMAccountName } else { [System.Convert]::DBNull }
    $Cmd.Parameters['UserPrincipalName'].Value = if ( $Item.UserPrincipalName -ne $null ) { $Item.UserPrincipalName } else { [System.Convert]::DBNull }
    $Cmd.Parameters['ObjectClass'].Value = $Item.ObjectClass
    $Cmd.Parameters['ObjectGUID'].Value = $Item.ObjectGUID
    $Cmd.Parameters['ObjectSID'].Value = if ( $Item.ObjectSID -ne $null ) { $Item.ObjectSID.ToString() } else { [System.Convert]::DBNull }
    $Cmd.Parameters['DTCreation'].Value = if ( $Item.WhenCreated -ne $null ) { $Item.WhenCreated.ToUniversalTime() } else { [System.Convert]::DBNull }
    $Cmd.Parameters['DTDeletion'].Value = if ( $Item.WhenChanged -ne $null ) { $Item.WhenChanged.ToUniversalTime() } else { [System.Convert]::DBNull }

    $Cmd.ExecuteNonQuery() | Out-Null
}
#endregion

#region Close the connection to the database
if ( $Conn.State -notin ('broken', 'closed') )
{
    try { $Conn.Close() } catch {}
}
#endregion

Stop-Script
