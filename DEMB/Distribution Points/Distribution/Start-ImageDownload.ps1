<#
    Start-ImageDownload.ps1

    Initiate Image download.

    --- Version history
    Version 1.00 (2016-05-03, Kees Hiemstra)
    - Initial version.
#>

#region Define script variables
$SourceHttp = "http://ITAMWeb.corp.demb.com/SOEImage"
$SourcePath = "\\DEMBMCAPS032FG1.corp.demb.com\SOE\Image\Win7\"
#endregion

. "$($PSScriptRoot)\CommonEnvironment.ps1"

$ScriptName = "Start-ImageDownload"

Write-Log -Message "Initiate image download ($($env:USERNAME))"

#region Get files to process from SOETools
try
{
    Start-BitsTransfer -Source "http://ITAMWeb.corp.demb.com/SOETools/Image/ImageFiles.csv" -Destination "$SOEToolsPath" -Description 'Image' -DisplayName 'Image-csv' -ErrorAction Stop
}
catch
{
    Write-Break -Message "Can't get the files list from SOETools"
}

$ImageFiles = Import-Csv -Path "$($SOEToolsPath)\ImageFiles.csv"
#endregion

#region Get local environment
$LocalPath = (Get-WmiObject -Class Win32_Share | Where-Object { $_.Name -eq 'SOE_Image$' }).Path
if ([string]::IsNullOrEmpty($LocalPath))
{
    Write-Break -Message "There is no share called SOE_Image`$"
}
if (-not (Test-Path $LocalPath))
{
    Write-Break -Message "The local path for the share SOE_Image`$ doesn't exist"
}

$LocalPath += "\"
Write-Log -Message "LocalPath: $LocalPath"
#endregion

#region Select the files to copy
$SourceFiles = ($ImageFiles | Where-Object { $_.Action -eq 'Copy' }).Name

#Get all the relative target file names without the folders
$Target = (Get-ChildItem -Path "$LocalPath*" -Attributes !Directory | Where-Object {$_.Length -gt 0}).FullName
Write-Log -Message "Number of files target in target folder: $($Target.Count)"
if ($Target -eq $null)
{
    $Files = $SourceFiles
}
else
{
    $Target = $Target -replace ($LocalPath).Replace("\", "\\").Replace("$", "\$")
    $Files = (Compare-Object -ReferenceObject $SourceFiles -DifferenceObject $Target | Where-Object {$_.SideIndicator -eq "<="}).InputObject
}

$SourceFiles = ($ImageFiles | Where-Object { $_.Action -eq 'Newer' }).Name
foreach ($File in $SourceFiles)
{
    if ( -not (Test-Path -Path "$($LocalPath)$File") )
    {
        #Destination file doesn't exist
        $Files += $File
    }
    else
    {
        #Check if source file is newer
        $LocalTime = (Get-ChildItem -Path "$($LocalPath)$File").LastWriteTimeUtc.AddSeconds(2)
        $RemoteTime = (Get-ChildItem -Path "$($SourcePath)$File").LastWriteTimeUtc

        Write-Verbose -Message "Time local file: $LocalTime"
        Write-Verbose -Message "Time remote file: $RemoteTime"

        if ( $LocalTime -lt $RemoteTime )
        {
            $Files += $File
        }
    }
}
#endregion

#region Delete files from previous images
$SourceFiles = ($ImageFiles | Where-Object { $_.Action -eq 'Remove' }).Name

foreach ($File in $SourceFiles)
{
    if ( (Test-Path -Path "$($LocalPath)$File") )
    {
        Remove-Item -Path "$($LocalPath)$File" -Force -Confirm:$false
        Write-Log -Message "Deleted [$File]"
    }
}
#endregion

#region Start transfer
if ( $Files.Count -gt 0 )
{
#    $SendCompleted = $true
    Write-Log -Message "Number of files to transfer: $($Files.Count)"

    Write-Log -Message "Start BitsTransfer"
    try
    {
        $Bits = Start-BitsTransfer -Description "Image download ($((Get-CorrectedDate).ToString('yyyy-MM-dd HH:mm:ss')))" -Source ($Files | ForEach-Object { "$SourceHttp/$($_.Replace('\', '/'))" }) -Destination ($Files | ForEach-Object { "$LocalPath$_" }) -DisplayName 'Image' -ProxyUsage NoProxy -RetryInterval 300 -RetryTimeout 1209600 -Asynchronous -Suspended -ErrorAction Stop
        [array]$Status.Download += $Bits
        Update-Status
        Write-Log -Message "BitsTransfer initiated"
    }
    catch
    {
        Write-Break -Message "BitsTransfer initiation failed"
    }
}
else
{
    Write-Log -Message "No files to transfer"
}
#endregion

Stop-Script

