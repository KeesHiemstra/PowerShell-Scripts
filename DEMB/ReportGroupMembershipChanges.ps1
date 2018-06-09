<#
    ReportGroupMembershipChanges.ps1

    Report on changes in the AD groups listed in the csv file GroupsToCheck.txt

    Powershell version 3.0

    Version 1.10 (2016-09-06, Kees Hiemstra)
    - Bug fix: Adding object went wrong sometimes.
    - Bug fix: No data was processed when checking users in groups.
    - Made a proper HTML table.
    Version 1.00 (2014-02-24, Kees Hiemstra)
    - Initial version.
#>

$Header = @"
<style>
BODY{background-color:peachpuff;}
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:thistle}
TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:PaleGoldenrod}
</style>
"@

$DataPath = "\\DEMBMCAPS032SQ2.corp.demb.com\H$\ITAM\Etc\ReportGroupMembershipChanges"

if (Test-Path "$DataPath\GroupsToCheck.txt")
{
    $GroupsToCheck = Import-Csv "$DataPath\GroupsToCheck.txt"
    foreach ($GroupName in $GroupsToCheck.Name)
    {
        #try
        #{
            #Read the current situation
            $CurrMembers = Get-ADGroupMember $GroupName -Recursive | Select-Object Name

            if (Test-Path "$DataPath\$GroupName.csv")
            {
                Write-Host "Processing $GroupName"

                #Read the previous situation
                $PrevMembers = Import-Csv "$DataPath\$GroupName.csv"

                #Compare
                $DiffMembers = Compare-Object -ReferenceObject $PrevMembers -DifferenceObject $CurrMembers -Property Name

                $HTML = ""
                if ($DiffMembers.Count -gt 0)
                {
                    #Get the details to report on
                    [array]$Added = $DiffMembers | Where-Object { $_.SideIndicator -eq "=>" } | ForEach-Object { $Name = $_.Name; Get-ADObject -Filter { Name -eq $Name } -Properties Name, Description } | Select-Object Name, Description
                    [array]$Deleted = $DiffMembers | Where-Object SideIndicator -eq "<="

                    #Report
                    $HTML = "<p>The following changes have been reported since $((Get-Item "$DataPath\$GroupName.csv").LastWriteTime.ToString("yyyy-MM-dd hh:mm")).</p>`n`n"
                    if ($Added.Count -gt 0)
                    {
                        Write-Host "Report added objects"
                        $HTML = $Added | ConvertTo-Html -Fragment -Property Name, Description, Class -PreContent "$HTML<p>Added<p>`n"
                    }
                    if ($Deleted.Count -eq 0)
                    {
                        $HTMLSub = ""
                    }
                    else
                    {
                        Write-Host "Report deleted objects"
                        $HTMLSub = "<p>Deleted</p>"
                    }
                    $HTML = $Deleted | ConvertTo-Html -Property Name -PreContent "$HTML`n$HTMLSub`n" -Title "Change(s) in $GroupName group" -Head $Header

                    Send-MailMessage -Body "$HTML" -Subject "Change(s) in $GroupName group" -To "Kees.Hiemstra@hpcds.com" -From "HPDesktop.Administrator@jdecoffee.com" -SmtpServer "smtp.corp.demb.com" -BodyAsHtml
                }
            }
            else
            {
                Write-Host "$GroupName has no data yet, no e-mail will be sent."
            }
            #Export the current group.
            $CurrMembers | Select-Object Name | Export-Csv "$DataPath\$GroupName.csv" -Force -NoTypeInformation
        #}
        #catch
        #{
        #    Write-Warning "Error occured @ $GroupName"
        #}
    }
}
else
{
    Write-Error "Can't find or access $DataPath\GroupsToCheck.txt"
}
