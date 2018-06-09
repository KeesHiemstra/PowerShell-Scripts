<#
    ExportCalviConfig.ps1

    These are the ExportCalvi.ps1 settings.
#>

$LogFile = "B:\ExportCalvi.log"
$MailSMTP = "smtp.corp.demb.com"

$ExportPath = "B:\$(Get-Date -Format 'yyyyMMdd_hhmm').csv"
$MailExportTo = "Kees.Hiemstra@hpe.com" #"adexport.telecom@demb.com"
$MailExportSubject = "Interface from Active Directory to Calvi"
$MailExportBody = "Attached you'll find the Active Directory data for Calvi."

