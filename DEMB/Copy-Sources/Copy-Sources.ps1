<#
    Copy Daily.ps1

    Version 1.10 (2016-07-12, Kees Hiemstra)
    - Exclude copying log files.
    Version 1.02 (2015-12-23, Kees Hiemstra)
    - Updated the target folder.
    - Reduced the number of days to 14.
    Version 1.01 (2015-12-17, Kees Hiemstra)
    - Looking up the drive letter of the disk called "Sources".
    - Added log file on target folder.
    Version 1.00 (2015-12-16, Kees Hiemstra)
    - Inital version.
#>

$Date = (Get-Date).Date.AddDays(-14)

$SourcePath = "C:\Src\"

# Search for the disk with the volume name: Sources
$TargetDisk = (Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.VolumeName -eq 'Sources' }).DeviceID
if ( $TargetDisk -eq $null )
{
    Write-Error "Disk not found"
    break
}

#region Copy split on date
$TargetPath = "$TargetDisk\Sources\Diff"

#Define log file variables
$ScriptName = $MyInvocation.MyCommand.Definition.Replace("$PSScriptRoot\", '').Replace(".ps1", '')
$LogFile = "$TargetPath\Copy daily.log"

#region LogFile
$Error.Clear()

#Create the log file if the file does not exits else write today's first empty line as batch separator
if ( -not (Test-Path $LogFile) )
{
    New-Item $logFile -ItemType file | Out-Null
}
else 
{
    Add-Content -Path $LogFile -Value "---------- --------"
}

function Write-Log([string]$Message)
{
    $LogMessage = "{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"),$Message
    Add-Content -Path $LogFile -Value $LogMessage
    Write-Debug $Message
}

#Finish script and report on $Error
function Stop-Script([bool]$WithError)
{
    if ( $Error.Count -gt 0 )
    {
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

$Files = Get-ChildItem -Path $SourcePath -Recurse -File | Where-Object { $_.LastWriteTime -ge $Date }

foreach ( $File in $Files )
{
    $Destination = $File.FullName.Replace($SourcePath, "$TargetPath\$($File.LastWriteTime.ToString('yyyy-MM-dd'))\")
    $Target = Split-Path -Path $Destination
    $TargetFile = $null

    if ( -Not (Test-Path -Path $Target) )
    {
        New-Item -Path $Target -ItemType Directory | Out-Null
    }
    else
    {
        $TargetFile = Get-ChildItem -Path $Destination -ErrorAction SilentlyContinue
    }

    if ( $TargetFile -eq $null -or $File.LastWriteTime -gt $TargetFile.LastWriteTime -and $File.Extension -ne '.log' )
    {
        Copy-Item -Path $File.FullName -Destination $Destination -Verbose -Force
        Write-Log -Message "Copied $($File.Name) to $($Destination)"
    }
}

#endregion

#region Copy all
$TargetPath = "$TargetDisk\Sources\Src"

#Define log file variables
$LogFile = "$TargetPath\Copy daily.log"

foreach ($File in $Files)
{
    $Destination = $File.FullName.Replace($SourcePath, "$TargetPath\")
    $Target = Split-Path -Path $Destination
    $TargetFile = $null

    if ( -Not (Test-Path -Path $Target) )
    {
        New-Item -Path $Target -ItemType Directory | Out-Null
    }
    else
    {
        $TargetFile = Get-ChildItem -Path $Destination -ErrorAction SilentlyContinue
    }

    if ( $TargetFile -eq $null -or $File.LastWriteTime -gt $TargetFile.LastWriteTime -and $File.Extension -ne '.log' )
    {
        Copy-Item -Path $File.FullName -Destination $Destination -Force
        Write-Log -Message "Copied $($File.Name) to $($Destination)"
    }
}

#endregion
