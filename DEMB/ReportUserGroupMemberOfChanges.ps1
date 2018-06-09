"H:\ITAM\Scripts\PowerShell\ReportUserGroupMemberOfChanges.ps1"###############################################################################################
# ReportUserGroupMemberOfChanges
#
# Report on changes in the MemberOf AD groups listed in the csv file UsersToCheck.txt
#
# Powershell version 3.0
#
# Version 1.00 (2014-09-24, Kees Hiemstra)
# - Initial version.
###############################################################################################

$DataPath = "\\DEMBMCAPS032SQ2.corp.demb.com\H$\ITAM\Etc\ReportGroupMembershipChanges"

if (Test-Path "$DataPath\UsersToCheck.txt")
{
    $UsersToCheck = Import-Csv "$DataPath\UsersToCheck.txt"

    foreach($san in $UsersToCheck.Name)
    {
        $FileName = "$DataPath\$san.csv"

        #Create the list of group memberships for every user
        $ADGroups = [string[]]@()
        $ADUser = Get-ADUser $san -Properties MemberOf
        foreach ($mo in $ADUser.MemberOf.GetEnumerator())
        {
            $ADGroups += Get-ADGroup $mo | Select Name
        }

        if (Test-Path -Path $FileName)
        {
            $Groups = Import-Csv $FileName

            $Diff = Compare-Object -ReferenceObject $Groups -DifferenceObject $ADGroups -Property Name

            if ($Diff -ne $null)
            {
                $Message = ($Diff | Select Name, @{N="Action";E={if($_.SideIndicator -eq "=>"){"Added"} else {"Deleted"}}} | ConvertTo-Html) -join " "
                Send-MailMessage -From "HPDesktop.Administrator@demb.com" `
                    -SMTPserver "smtp.corp.demb.com" `
                    -To "Kees.Hiemstra@hp.com" `
                    -Subject "Groupmembership change of $san" `
                    -BodyAsHtml `
                    -Body $Message `
                    -DeliveryNotificationOption onFailure

                $ADGroups | Export-Csv -LiteralPath $FileName -NoTypeInformation
            }
        }
        else
        {
            $ADGroups | Sort-Object Name | Export-Csv -LiteralPath $FileName -NoTypeInformation

            $Message = ($ADGroups | Select Name | ConvertTo-Html) -join " "
            Send-MailMessage -From "HPDesktop.Administrator@demb.com" `
                -SMTPserver "smtp.corp.demb.com" `
                -To "Kees.Hiemstra@hp.com" `
                -Subject "Initial groupmemberships of $san" `
                -BodyAsHtml `
                -Body $Message `
                -DeliveryNotificationOption onFailure
        }
    }
}
