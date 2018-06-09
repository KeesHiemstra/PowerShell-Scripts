<#
    New-OfficeFiles.ps1

    Create new OfficeFiles.csv for a new office distribution.

    --- Version history
    Version 1.10 (2016-10-17, Kees Hiemstra)
    - Include Office 2016
    Version 1.00 (2016-05-16, Kees Hiemstra)
    - Initial version.
#>

$SourcePath = '\\DEMBMCAPS032FG1.corp.demb.com\Source\O365\Distribution\'
$OfficeFilesPath = '\\DEMBMCAPS032SQ2.corp.demb.com\SOETools$\Office\OfficeFiles.csv'

if ( -not (Test-Path $SourcePath) )
{
    Write-Error -Message "Path ($SourcePath) not found"
}

if ( -not $SourcePath.EndsWith('\') ) { $SourcePath += '\' }

#Get folders
$OfficeFiles = Get-ChildItem -Path $SourcePath -Exclude @('InstallO365_xx-xx.xml', 'Web.config') -Recurse -Directory |
    Select-Object @{ n='Type'; e={ 'D' } }, LastWriteTimeUtc, Length, @{ n='RelativeName'; e={ "$($_.FullName.Replace($SourcePath, ''))\" } }

#Add files
$OfficeFiles += Get-ChildItem -Path $SourcePath -Exclude @('InstallO365_xx-xx.xml', 'Web.config') -Recurse -File |
    Select-Object @{ n='Type'; e={ 'F' } }, LastWriteTimeUtc, Length, @{ n='RelativeName'; e={ $_.FullName.Replace($SourcePath, '') } }

#Export to SOETools\Office share
$OfficeFiles | 
    Sort-Object Type, RelativeName |
    Select-Object RelativeName, LastWriteTimeUtc, Length |
    Export-Csv -Path $OfficeFilesPath -NoTypeInformation
