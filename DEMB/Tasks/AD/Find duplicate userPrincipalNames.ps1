$ADusers = Get-ADUSer -Filter *
$Duplicates = $ADUsers | Group-Object userPrincipalName | Where-Object { $_.Name -ne '' -and $_.Count -gt 1 }
if ($Duplicates.Count -gt 0)
{
    $HTML = $Duplicates.Group | Select-Object sAMAccountName, userPrincipalName | ConvertTo-Html -Title "Duplicate userPrincipalNames"| Out-String

    Send-MailMessage -To Kees.Hiemstra@hpe.com -Body $HTML.ToString() -BodyAsHtml -SmtpServer smtp.corp.demb.com -Subject "Duplicate userPrincipalNames" -From "AD.Check@JDECoffee.com"
}
else
{
    "No duplicates found."
}



