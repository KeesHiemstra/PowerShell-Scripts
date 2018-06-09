<#
DEMBUserMailboxReport

Version 5.0 (2016-07-20, Kees Hiemstra)
- Added Manager attribute to the report.
- Changed the sener to Reporter@JDEcoffee.com
Version 4.9 (2015-09-09, Kees Hiemstra)
- Added sAMAccountName.
Version 4.8 (2015-08-26, Kees Hiemstra)
- Added mobile nummber.
Version 4.7.1 (2015-04-17, Kees Hiemstra)
- Bug fix: Some remote mailboxes are reported as local in AD and therefore not reported.
Version 4.7.0 (2015-01-28, Kees Hiemstra)
- Added UserPrincipalName to the report.
Version 4.6.0 (2015-01-05, Kees Hiemstra)
- Added Co to the report.
Version 4.5.0 (2014-12-11, Kees Hiemstra)
- Added ExtensionAttribute5 to the report.
Version 4.4.0 (2014-10-06, Kees Hiemstra)
- Adding two new columns: Lync Enabled and Lync Enterprise Voice.
Version 4.3.0 (2014-09-29, Kees Hiemstra)
- Added -Monthly switch to change the number of reciepence.
Version 4.2.0 (2014-09-19, Kees Hiemstra)
- Added -Encoding UTF8 to have Excel work with extended characters.
Version 4.1.0 (2014-09-08, Kees Hiemstra)
- Create a report with the name contains the date of creation.
- EMail the report.
Version 4.0.2 (2014-09-08, Kees Hiemstra)
- Don't export records where RecipientType is empty.
Version 4.0.1 (2014-09-05, Kees Hiemstra)
- Collect details from remote mailboxes.
Version 4.0.0 (2014-09-04, Kees Hiemstra)
- Initial version taken over from version 3.0
#>
Param ([Switch]$Monthly)

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010  # E2k10

#$ExchangeServers = @("dembswacas090","dembswacas091","dembswacas092")

#remove existing sessions to the CAS servers
#Get-PSSession | Where-Object {$ExchangeServers -contains $_.ComputerName.Split(".")[0]} | Remove-PSSession

#Open Exchange Managementshell and pick one of the available CAS servers
#$ExchangeServer = $ExchangeServers | Where-Object {Test-Connection -ComputerName $_ -Count 1 -Quiet} | Get-Random
#$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ("http://{0}.corp.demb.com/PowerShell" -f $ExchangeServer)

#Import the module from CAS server
#Import-PSSession $ExchangeSession | Out-Null


if ($Monthly.IsPresent)
{
    $ToAddress = "jde-billing@hp.com", "Agata.Szczytnicka@hpe.com", "Valentin.Diaconu@hpe.com", "Girishbs@hpe.com", "Richard.Wesseling@hpe.com"
}
Else
{
    $ToAddress = "Richard.Wesseling@hpe.com"
}

#region Collection
$ADUsers = Get-ADUser -Filter { Name -notlike "FederatedEmail*" -and Name -notlike "SystemMailbox*" -and Name -ne "MSOL_AD_Sync"} `
        -Properties Name, DisplayName, Company, City, Co, Country, c, CountryCode, CanonicalName, UserPrincipalName, EmployeeID, msExchArchiveName, msRTCSIP-UserEnabled, msRTCSIP-Line, extensionAttribute2, extensionAttribute5, extensionAttribute15, legacyExchangeDN, mobile, sAMAccountName, Manager |
    Select-Object @{n="CommonName"; e={$_.Name}}, `
        DisplayName,
        @{n="PrimarySmtpAddress"; e={""}},
        @{n="CompanyNumber"; e={$_.Company}},
        City,
        @{n="CountryOrRegion"; e={$_.Co}},
        @{n="CountyCode"; e={$_.Country}},
        @{n="OU"; e={$_.CanonicalName.Substring(0, $_.CanonicalName.Length - $_.Name.Length - 1)}},
        @{n="RecipientType"; e={if($_.legacyExchangeDN -like "/o=DEMB/ou=External*") {"RemoteMailbox"} else {""}}},
        EmployeeID,
        @{n="Alias"; e={""}},
        @{n="msExchArchiveName"; e={$_.msExchArchiveName -join ";"}},
        @{n="Lync Enabled"; e={$_."msRTCSIP-UserEnabled"}},
        @{n="Lync Enterprise Voice"; e={$_."msRTCSIP-Line" -join ";"}},
        @{n="CustomAttribute2"; e={$_.extensionAttribute2}},
        @{n="CustomAttribute5"; e={$_.extensionAttribute5}},
        @{n="CustomAttribute15"; e={$_.extensionAttribute15}},
        @{n="UPN"; e={$_.UserPrincipalName}},        @{n="Mobile"; e={$_.mobile}},
		@{n="Pre-Windows2000 account"; e={$_.sAMAccountName}},
        Manager

foreach ($U in $ADUsers)
{
    if ($U.RecipientType -eq "RemoteMailbox")
    {
        $Mailbox = Get-RemoteMailbox $U.CommonName -ErrorAction SilentlyContinue
    }
    else
    {
        $Mailbox = Get-Mailbox $U.CommonName -ErrorAction SilentlyContinue
    }

    if ($Mailbox -eq $null)
    {
        #Bug fix: Some mailboxes report in AD as being local, but these are remote.
        $Mailbox = Get-RemoteMailbox $U.CommonName -ErrorAction SilentlyContinue
        if($Mailbox -eq $null)
        {
            continue
        }

        $U.RecipientType = "RemoteMailboxx"
    }

    $U.PrimarySmtpAddress = $Mailbox.PrimarySmtpAddress.ToString().ToLower()
    $U.Alias = $Mailbox.Alias -join ";"

    $U.RecipientType = $Mailbox.RecipientTypeDetails[0]
    if ($Mailbox.ResourceType -eq "Room") {$U.RecipientType = "RoomMailbox"}
    elseif ($Mailbox.RecipientType -eq "Equipment") {$U.RecipientType = "EquipmentMailbox"}
    elseif ($Mailbox.IsShared) {$U.RecipientType = "SharedMailbox"}
}
#endregion

#region Reporting
$ReportFileName = "D:\Scripts\UserMailBoxReport\Reports\MailboxReport-{0}.csv" -f (Get-Date -Format "yyyy-MM-dd_hhmm")

$ADUsers |
    Where-Object {$_.RecipientType -ne ""} |
    Export-Csv -Path $ReportFileName -NoTypeInformation -Encoding UTF8

Send-MailMessage -From "Reporter@JDEcoffee.com" `
    -To $ToAddress `
    -Subject ("MailboxReport {0}" -f (Get-Date -Format "yyyy-MM-dd")) `    -Body ("Attached you'll find the Mailbox report of today") `    -Attachments $ReportFileName `    -SmtpServer smtp.corp.demb.com

#endregion

#remove existing sessions to the CAS servers
#Get-PSSession | Where-Object {$ExchangeServers -contains $_.ComputerName.Split(".")[0]} | Remove-PSSession

