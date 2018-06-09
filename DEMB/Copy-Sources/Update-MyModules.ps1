<#
    Update-MyModules.ps1

    Copies the modules from the Update-MyModules-Modules.csv to my local profile if there are updates.

    === Version history
    Version 1.00 (2016-07-07, Kees Hiemstra)
    - Initial version.
#>

$LocalProfile = "$(Split-Path -Path $Profile)\Modules"
$MyModules = Import-Csv -Path ($MyInvocation.MyCommand.Definition -replace(".ps1$","-Modules.csv")) |
    Select-Object *, @{n='ProfilePath'; e={ "$LocalProfile\$($_.Module)" }}

Function Compare-Folders([string]$Path, [string]$Target)
{
    $Result = $false

    $SourceFiles = Get-ChildItem -Path "$Path\*.*"
    foreach($SourceFile in $SourceFiles)
    {
        if (-not (Test-Path $Target))
        {
            New-Item $Target -ItemType Directory | Out-Null
        }

        $TargetFile = Get-ChildItem -Path "$Target\$($SourceFile.Name)" -ErrorAction SilentlyContinue
        if ($TargetFile -eq $null -or $SourceFile.LastWriteTime -gt $TargetFile.LastWriteTime)
        {
            Write-Host "Update $TargetFile with $SourceFile"
            Copy-Item -Path $SourceFile -Destination $Target -Force
            $Result = $true
        }
    }
    Write-Output $Result
}

foreach($Module in $MyModules)
{
    if ((Compare-Folders -Path $Module.SourcePath -Target $Module.ProfilePath))
    {
        Get-Module $Module.Module | Remove-Module -Force
    }
}

break
$MyModules = @{
    'SOEUser'     = "C:\Src\PowerShell\SOE Tools\Modules\SOEUser";
    'SOETools'    = "C:\Src\PowerShell\SOE Tools\Modules\SOETools";
    'SOEComputer' = "C:\Src\PowerShell\SOE Tools\Modules\SOEComputer";
    'Azure'       = "C:\Src\PowerShell\SOE Tools\Modules\Azure";
    }

Function Compare-Folders([string]$Path, [string]$Target)
{
    $Result = $false

    $SourceFiles = Get-ChildItem -Path "$Path\*.*"
    foreach($SourceFile in $SourceFiles)
    {
        $TargetFile = Get-ChildItem -Path "$Target\$($SourceFile.Name)"
        if ($TargetFile -ne $null -and $SourceFile.LastWriteTime -gt $TargetFile.LastWriteTime)
        {
            Write-Host "Update $TargetFile with $SourceFile"
            Copy-Item -Path $SourceFile -Destination $Target -Force
            $Result = $true
        }
    }
    Write-Output $Result
}

foreach($Unit in $MyModules.GetEnumerator())
{
    if ((Compare-Folders -Path $Unit.Value -Target "C:\Users\Kees.Hiemstra\Documents\WindowsPowerShell\Modules\$($Unit.Key)"))
    {
        Get-Module $Unit.Key | Remove-Module
        Import-Module $Unit.Key
    }
}

