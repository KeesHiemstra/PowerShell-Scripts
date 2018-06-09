<#
    EnableDeniedLicenses.ps1

    Monitor enabled account that are denied to have licenses.

    === Version history
    Version 1.00 (2017-04-14, Kees Hiemstra)
    - Initial version
#>

$MailSplatting = @{'SmtpServer' = 'smtp.corp.demb.com'; 'From' = 'Kees.Hiemstra@JDEcoffee.com'; 'To' = 'Kees.Hiemstra@hpe.com' }
$ExportPath = $MyInvocation.MyCommand.Definition -replace '.ps1', '.csv'

[Array]$List = Get-ADUser -Filter { Enabled -eq $true -and ExtensionAttribute2 -like '!*' } -Properties EmployeeID, SAMAccountName, Mail, WhenCreated, WhenChanged |
    Select-Object EmployeeID, SAMAccountName, Mail, WhenCreated, WhenChanged

if ( $List.Count -gt 0 )
{
    if ( Test-Path -Path $ExportPath )
    {
        $Prev = Import-Csv -Path $ExportPath -Delimiter ','
    }
    else
    {
        $Prev = @()
    }
}

$List | Where-Object { $_.SAMAccountName -notin $Prev.SAMAccountName } |
    Send-ObjectAsHTMLTableMessage -Subject 'Enabled accounts with denied licenses' @MailSplatting -MessageType Exception


$List | Export-Csv -Path $ExportPath -Delimiter ',' -NoTypeInformation -Force
