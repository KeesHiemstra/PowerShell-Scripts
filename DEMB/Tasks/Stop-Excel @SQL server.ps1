$Server = New-PSSession -ComputerName 'DEMBMCAPS032N2.corp.demb.com'

Invoke-Command -Session $Server -ScriptBlock { Get-Process -Name Excel | Stop-Process -Confirm:$false }