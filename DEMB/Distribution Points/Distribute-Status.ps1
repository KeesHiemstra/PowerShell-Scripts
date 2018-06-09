<#
    Distribute-Status.ps1

    Create and copy the Status.xml to the Distribution Points.
    $PSScriptRoot\Scheduling.csv

    --- Version history
    Version 1.00 (2016-06-05, Kees Hiemstra)
    - Initial version.
#>

$Files = @('CommonEnvironment.ps1',
           'Send-StatusUpdate.ps1',
           'Start-ImageDownload.ps1',
           'Start-OfficeDownload.ps1',
           'Test-AfterReboot.ps1',
           'Test-Download.ps1',
           'Update-DistributionScripts.ps1'
           )

foreach ( $Schedule in (Import-Csv -Path "$($PSScriptRoot)\Scheduling.csv" -Delimiter "`t") )
{
    $DestinationPath = "\\$($Schedule.ComputerName)\SOETools$"
    $Destination = "$DestinationPath\Status.xml"

    if ( -not (Test-Path -Path $Destination) )
    {
        Get-ChildItem -Path "$DestinationPath\G*.*" | Remove-Item

        $Properties = [ordered]@{ComputerName=$Schedule.ComputerName; 
                        CountryCode=$Schedule.c;
                        Location=$Schedule.l;
                        Description=$Schedule.Type;
                        TimeZone=$Schedule.TimeZone;
                        InOfficeHours=$false;
                        StartStatus=$null;
                        LastStatus=(Get-Date -Year 2016 -Month 6 -Day 6 -Hour 0 -Minute 0 -Second 0);
                        Download=@();
                        }

        $Status = New-Object -TypeName PSObject -Property $Properties
        $Status | Export-Clixml -Path $Destination

        foreach ($File in $Files)
        {
            Copy-Item -Path "$($PSScriptRoot)\Distribution\$File" -Destination "$DestinationPath\$File"
        }
        Write-Host "$($Schedule.ComputerName) got the status and initial files at $($DestinationPath)"
    }#Not exists
}#foreach

break

PSEdit -filenames "C:\Src\PowerShell\4.0\!DEMB\SOETools\Scheduling.csv"