$ExchangeServers = @("DEMBRSMS419")

#remove existing sessions to the CAS servers
Get-PSSession | Where-Object { $ExchangeServers -contains $_.ComputerName.Split(".")[0] } | Remove-PSSession

## Open Exchange Managementshell and pick one of the available CAS servers
$ExchangeServer = $ExchangeServers | Where-Object {Test-Connection -ComputerName $_ -Count 1 -Quiet} | Get-Random
$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ("http://{0}.corp.demb.com/PowerShell" -f $ExchangeServer)

#import the module from CAS server
Import-PSSession $ExchangeSession | Out-Null

