<#
    SAPShortcutFix.ps1

    PowerShell version 2.0

    Aldea 1226803: RITM0013111 Adjust SAPGui shortcuts

    === Script steps
    - Delete SAP*.ini and SAPRules.xml from C:\Windows
      Except SAPLogon.ini if the HKLM:\Software\JDE\SAPLogon_Config::WindowsSAPLogon = 1
    - Delete SAP*.ini and SAPRules.xml from C:\Users\<user name>\AppData\Local\VirtualStore
    - Delete SAP*.ini and SAPRules.xml from C:\Users\<user name>\AppData\Local\VirtualStore\Program Files (x86)\JDE\SAPLogon_Config
    - Copy SAP*.ini and SAPRules.xml from NetLogon to C:\Program Files (x86)\JDE\SAPLogon_Config
    - Delete the existing shortcuts from the (public) desktop and start menu
    - Delete C:\Program Files (x86)\SAPSelect\SAPSelect.exe

    === Version history
    1.31 (2016-12-30, Kees Hiemstra)
    - Bug fix: Error message in log should not be there.
    1.30 (2016-11-28, Kees Hiemstra)
    - Implement changes from Aldea 1254256: RITM0014963 Adjust script for SAPLOGON.ini:
      - Keep and update C:\Windows\SAPLogon.ini if the registry key HKLM:\Software\JDE\SAPLogon_Config::WindowsSAPLogon = 1
    - Delete ini files from C:\Users\<user name>\AppData\Local\VirtualStore\Program Files (x86)\JDE\SAPLogon_Config
    - Set the distribution source folder based on the registry key HKLM:\Software\JDE\SAPLogon_Config::DistributionPath
    1.20 (2016-07-22, Kees Hiemstra)
    - Adjusting the script for PowerShell version 2.0.
    1.10 (2016-07-20, Kees Hiemstra)
    - New location of the local SAP ini files (instead of in the personal folder of the user): C:\Program Files (x86)\JDE\SAPLogon_Config
    1.00 (2016-07-15, Kees Hiemstra)
    - Initial version.
#>

#region Default settings
$ScriptVersion = '1.31'

$DefaultDistributionPath = '\\corp.demb.com\NETLOGON\SOE\SAP\'

$WindowsSAPLogon = $false

$ConfigFolder = 'C:\Program Files (x86)\JDE\SAPLogon_Config'

#endregion

#region LogFile
#Define log file variables
$RunInDevelopment = $false
$ScriptName = $MyInvocation.MyCommand.Name.Replace(".ps1", '')
$LogFile = "C:\HP\Logs\$ScriptName.log"

$Error.Clear()

#Create the log file if the file does not exits else write today's first empty line as batch separator
#if (-not (Test-Path $LogFile)) { New-Item $LogFile -ItemType file | Out-Null }

Add-Content -Path $LogFile -Value "-v$($ScriptVersion.PadRight(8,'-')) --------"

function Write-Log([string]$Message)
{
    $LogMessage = "{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Add-Content -Path $LogFile -Value $LogMessage
    if ( $RunInDevelopment ) { Write-Host $LogMessage }
}

#Finish script and report on $Error
function Stop-Script([bool]$WithError)
{
    if ($Error.Count -gt 0)
    {
        $ErrorLines = $Error | Out-String
        foreach ( $Line in $ErrorLines )
        {
            Write-Log -Message "ERROR: $Line"
        }
    }

    if ($WithError)
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

#region Helper functions
function DeleteSAPIni([string]$Folder)
{
    if ( -not $Folder.EndsWith('\') ) { $Folder += '\' }

    if ( -not (Test-Path -Path $Folder) ) { return }

    if ( $WindowsSAPLogon -and $Folder -eq 'C:\Windows\' )
    {
        $Files = Get-ChildItem -Path @("$Folder\SAP*.ini", "$Folder\SAPRules*.xml") -Exclude 'SAPLogon.ini' | Select-Object FullName
    }
    else
    {
        $Files = Get-ChildItem -Path @("$Folder\SAP*.ini", "$Folder\SAPRules*.xml") | Select-Object FullName
    }   

    foreach ( $File in $Files )
    {
        $FileName = $File | Select-Object -ExpandProperty FullName
        if ( $null -ne $FileName -and (Test-Path -Path $FileName) )
        {
            Remove-Item -Path $FileName -Force
            Write-Log -Message "Deleted file: $($FileName)"
        }
    }
}
#endregion

#region Settings

#region Setting - Folder to get the files from
$DistributionPath = $DefaultDistributionPath
try
{
    $DistributionPath = Get-ItemProperty -Path 'HKLM:\Software\JDE\SAPLogon_Config' -Name 'DistributionPath' -ErrorAction Stop | Select-Object -ExpandProperty DistributionPath
}
catch
{
    $Error.RemoveAt(0)
    $DistributionPath = $DefaultDistributionPath
}
if ( [string]::IsNullOrWhiteSpace($DistributionPath) )
{
    $DistributionPath = $DefaultDistributionPath
}
if ( -not $DistributionPath.EndsWith('\') ) { $DistributionPath += '\' }
#endregion

#region Setting - Should the SAPLogon.ini file should be in C:\Windows?
try
{
    $WindowsSAPLogon = (Get-ItemProperty -Path 'HKLM:\Software\JDE\SAPLogon_Config' -Name 'WindowsSAPLogon' -ErrorAction Stop | Select-Object -ExpandProperty WindowsSAPLogon) -eq 1
}
catch
{
    $Error.RemoveAt(0)
    $WindowsSAPLogon = $false
}
#endregion

#endregion

#region Collect the files to distribute
$SAPiniFiles = Get-ChildItem -Path @("$($DistributionPath)SAP*.ini", "$($DistributionPath)SAPRules.xml")

if ( $SAPiniFiles -eq $null )
{
    Write-Break "No access to the netlogon folder, can't copy the files"
}
#endregion

#region Delete all unused SAP*.ini files from the Windows folder and user folders

#Delete from C:\Windows
DeleteSAPIni -Folder C:\Windows

#Process for all users
$Users = Get-ChildItem -Path 'C:\Users'
foreach ( $User in $Users )
{
    if ( ($User | Select-Object -ExpandProperty Name) -notmatch '^Administrator$|^msdts.*|^mssql.*|^public$|^sql*.|^svc\..*' )
    {
        $UserPath = ($User | Select-Object -ExpandProperty FullName)
        if ( -not $UserPath.EndsWith('\') ) { $UserPath += '\' }

        DeleteSAPIni -Folder "$($UserPath)AppData\Local\VirtualStore"
        DeleteSAPIni -Folder "$($UserPath)AppData\Local\VirtualStore\Program Files (x86)\JDE\SAPLogon_Config"
    }
}
#endregion

#region Copy SAP files
if ( -not (Test-Path -Path $ConfigFolder) )
{
    New-Item -Path $ConfigFolder -ItemType Directory | Out-Null

    Write-Log "Created folder: $ConfigFolder"
}

foreach ( $SAPIniFile in $SAPiniFiles )
{
    if ( -not (Test-Path -Path "$ConfigFolder\$(($SAPIniFile | Select-Object -ExpandProperty Name))") -or ($SAPIniFile | Select-Object -ExpandProperty LastWriteTimeUtc) -ne ((Get-ChildItem -Path "$ConfigFolder\$(($SAPIniFile | Select-Object -ExpandProperty Name))") | Select-Object -ExpandProperty LastWriteTimeUtc) )
    {
        Copy-Item -Path ($SAPIniFile | Select-Object -ExpandProperty FullName) -Destination $ConfigFolder -Force

        Write-Log -Message "Copied file: $(($SAPIniFile | Select-Object -ExpandProperty FullName)) to $ConfigFolder"
    }
}

if ( $WindowsSAPLogon )
{
    if ( -not (Test-Path -Path 'C:\Windows\SAPLogon.ini') -or ($SAPIniFiles | Where-Object { $_.Name -eq 'SAPLogon.ini' } | Select-Object -ExpandProperty LastWriteTimeUtc) -ne ((Get-ChildItem -Path 'C:\Windows\SAPLogon.ini') | Select-Object -ExpandProperty LastWriteTimeUtc) )
    {
        Copy-Item -Path ($SAPIniFiles | Where-Object { $_.Name -eq 'SAPLogon.ini' } | Select-Object -ExpandProperty FullName) -Destination 'C:\Windows' -Force

        Write-Log -Message "Copied file: $(($SAPIniFiles | Where-Object { $_.Name -eq 'SAPLogon.ini' } | Select-Object -ExpandProperty FullName)) to C:\Windows"
    }
}
#endregion

#region Delete shortcuts from Menu and SAPLogonIni selector application
$DelFiles = @('C:\ProgramData\Microsoft\Windows\Start Menu\Programs\SAP Front End\SAP Logon.lnk',
    'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\SAP Front End\SAPLogon.ini selection.lnk',
    'C:\Program Files (x86)\SAPSelect\SAPSelect.exe',
    'C:\Users\Public\Desktop\SAP Logon.lnk')
foreach ( $DelFile in $DelFiles )
{
    if ( (Test-Path -Path $DelFile) )
    {
        Remove-Item -Path $DelFile -Force

        Write-Log -Message "Deleted: $DelFile"
    }
}
#endregion

Stop-Script
