<#
    Start-OfficeDownload.ps1

    Initiate Office sources download.

    --- Version history
    Version 1.11 (2016-12-09, Kees Hiemstra)
    - Correct the calculation of the number of jobs.
    Version 1.10 (2016-10-17, Kees Hiemstra)
    - Include the distribution of Office 2016.
    Version 1.01 (2016-06-14, Kees Hiemstra)
    - Bug fix: Office files were stored with an extra folder level (SOE).
    Version 1.00 (2016-05-16, Kees Hiemstra)
    - Initial version.
#>

#region Define script variables
$SourceHttp = 'http://ITAMWeb.corp.demb.com/O365Source'
#endregion

. "$($PSScriptRoot)\CommonEnvironment.ps1"

$ScriptName = "Start-OfficeDownload"
$MaxFileCount = 175

Write-Log -Message "Initiate Office download ($($env:USERNAME))"

#region Get files to process from SOETools
try
{
    Start-BitsTransfer -Source "http://ITAMWeb.corp.demb.com/SOETools/Office/OfficeFiles.csv" -Destination "$SOEToolsPath" -Description 'Office' -DisplayName 'Office-csv' -ErrorAction Stop
}
catch
{
    Write-Break -Message "Can't get the files list from SOETools"
}

$OfficeFiles = Import-Csv -Path "$($SOEToolsPath)\OfficeFiles.csv"
#endregion

#region Get local environment
$LocalPath = (Get-WmiObject -Class Win32_Share | Where-Object { $_.Name -eq 'SOE$' }).Path
if ([string]::IsNullOrEmpty($LocalPath))
{
    Write-Break -Message "There is no share called SOE`$"
}
if (-not (Test-Path $LocalPath))
{
    Write-Break -Message "The local path for the share SOE`$ doesn't exist"
}

$LocalPath += '\O365\'
Write-Log -Message "LocalPath: $LocalPath"

if ( -not (Test-Path -Path $LocalPath) )
{
    $null = New-Item -Path $LocalPath -ItemType Directory
}
#endregion

#region Check local subfolders
foreach ( $Folder in ($OfficeFiles | Where-Object { $_.RelativeName.EndsWith('\') }) )
{
    if ( -not (Test-Path -Path "$($LocalPath)$($Folder.RelativeName)") )
    {
        $null = New-Item -Path "$($LocalPath)$($Folder.RelativeName)" -ItemType Directory
    }
}
#endregion

#region Determine which files to download
$LocalFiles = Get-ChildItem -Path $LocalPath -Recurse -File |
    Select-Object @{ n='RelativeName'; e={ $_.FullName.Replace($LocalPath, '') } }, @{ n='LastWriteTimeUtc'; e={ $_.LastWriteTimeUtc.ToString('yyyy-MM-dd HH:mm:ss') } }, Length

if ( $LocalFiles -ne $null )
{
    $Diff = Compare-Object -ReferenceObject $LocalFiles -DifferenceObject ($OfficeFiles | Where-Object { -not $_.RelativeName.EndsWith('\') }) -Property RelativeName, Length, LastWriteTimeUtc

    $Files = ($Diff | Where-Object { $_.SideIndicator -eq '=>' }).RelativeName
}
else
{
    $Files = ($OfficeFiles | Where-Object { -not $_.RelativeName.EndsWith('\') }).RelativeName
}
#endregion

#region Start transfer
if ( $Files.Count -gt 0 )
{
    Write-Log -Message "Number of files to transfer: $($Files.Count)"
    $StartFileCount = 0
    $JobCount = 1

    While ($StartFileCount -lt $Files.Count)
    {
        if ( $Files.Count -le $MaxFileCount )
        {
            Write-Log -Message "Start BitsTransfer"
        }
        else
        {
            Write-Log -Message "Start BitsTransfer job $JobCount of $([math]::Truncate($Files.Count / $MaxFileCount) + 1)"
        }

        try
        {
            $FilesLimited = $Files[$StartFileCount..($StartFileCount + $MaxFileCount - 1)]

            $Params = @{Description = "Office download job $JobCount of $([math]::Truncate($Files.Count / $MaxFileCount) + 1) ($((Get-CorrectedDate).ToString('yyyy-MM-dd HH:mm:ss')))";
                Source = ($FilesLimited | ForEach-Object { "$SourceHttp/$($_.Replace('\', '/'))" });
                Destination = ($FilesLimited | ForEach-Object { "$LocalPath$_" });
                }
            
            $Bits = Start-BitsTransfer @Params -DisplayName 'Office' -ProxyUsage NoProxy -RetryInterval 300 -RetryTimeout 1209600 -Asynchronous -Suspended -ErrorAction Stop

            [array]$Status.Download += $Bits
            Update-Status
            Write-Log -Message "BitsTransfer initiated"
        }
        catch
        {
            Write-Break -Message "BitsTransfer initiation failed"
        }
        $StartFileCount += $MaxFileCount
        $JobCount++
    }#While
}
else
{
    Write-Log -Message "No files to transfer"
}
#endregion

Stop-Script
