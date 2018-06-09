<#
    Monitor changes to the url attribute of computers.

    version 1.00 (2016-05-09, Kees Hiemstra)
    - initial version.
#>

$Report = @()

$Admin = Get-ADComputer -Filter { url -like '*' } -Properties url, LastLogonDate

foreach ( $Item in $Admin )
{
    foreach ( $Account in $Item.url )
    {
        $Account = $Account.Replace('demb\', '').Trim()
        try
        {
            $ADUser = Get-ADUser $Account -ErrorAction Stop
            if ( -not $ADUser.Enabled )
            {
                $Report += "$($Item.Name): User [$Account] is disabled [$($Item.url -join ';')]"
            }
        }
        catch
        {
            try
            {
                $ADUser = Get-ADGroup $Account -ErrorAction Stop
            }
            catch
            {
                $Report += "$($Item.Name): [$Account] doesn't exist [$($Item.url -join ';')]"
            }
        }
    }
}

$PrevAdminPath = 'D:\Data\Jobs\MonitorUrl'
$PrevAdminFile = "$PrevAdminPath\Admin.csv"

$CurrentAdmin = $Admin |
    Select-Object Name, LastLogonDate, @{n='url'; e={ $_.url -join ";" }} |
    Sort-Object Name

if ( -not (Test-Path -Path $PrevAdminFile) )
{
    $CurrentAdmin | Export-Csv -Path $PrevAdminFile -NoTypeInformation -Delimiter "`t"
    Copy-Item -Path $PrevAdminFile -Destination "$PrevAdminPath\Admin-$((Get-Date).ToString('yyyy-MM-dd HHmm')).csv"
    break
}

$PrevAdmin = Import-Csv -Path $PrevAdminFile -Delimiter "`t"

$Diff = Compare-Object -ReferenceObject $CurrentAdmin -DifferenceObject $PrevAdmin -Property Name, url

foreach ( $Item in $Diff )
{
    if ( $Item.SideIndicator -eq '=>' )
    {
        if ( $Item.Name -notin ($Diff | Where-Object { $_.SideIndicator -eq '<=' }).Name )
        {
            $Report += "Url field of $($Item.Name) has been cleared"
        }
    }
    else
    {
        if ( $Item.Name -in ($Diff | Where-Object { $_.SideIndicator -eq '=>' }).Name )
        {
            $Report += "Url field of $($Item.Name) has been changed from [$(($Diff | Where-Object { $_.Name -eq $Item.Name -and $_.SideIndicator -eq '=>' }).url)] to [$($Item.url)]"
        }
        else
        {
            $Report += "Url field of $($Item.Name) has been set to [$($Item.url)]"
        }
    }
}

if ( $Diff.Count -eq 0 )
{
    break
}

Copy-Item -Path $PrevAdminFile -Destination "$PrevAdminPath\Admin-$((Get-Date).ToString('yyyy-MM-dd HHmm')).csv"
$CurrentAdmin | Export-Csv -Path $PrevAdminFile -NoTypeInformation -Delimiter "`t" -Force

Send-MailMessage -Body ($Report -join "`n") -From 'Kees.Hiemstra@JDEcoffee.com' -To 'Kees.Hiemstra@hpe.com' -SmtpServer 'smtp.corp.demb.com' -Subject 'Monitor changes to the url attribute of computers'