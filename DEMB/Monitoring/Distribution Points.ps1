$MailSplatting = @{'SmtpServer' = 'smtp.corp.demb.com'; 'From' = 'Kees.Hiemstra@JDEcoffee.com'; 'To' = 'Kees.Hiemstra@hpe.com' }

Get-ADDistributionPoint | Select-Object ComputerName, c, l, Type, LastLogonDate | Sort-Object LastLogonDate | Send-ObjectAsHTMLTableMessage -Subject 'Distribution Points' @MailSplatting