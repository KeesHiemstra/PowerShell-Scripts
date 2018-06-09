<#
    Troubleshooting Appoint process on computer
#>

$ComputerName = 'NL5CG4490N50'

Enter-PSSession -ComputerName $ComputerName -Credential (Get-SOECredential svc.uaa) -Authentication Credssp

break

Get-Content -Path 'C:\HP\Logs\Appoint Administrators.log'

CScript.exe \\corp.demb.com\NETLOGON\SOE\Appoint.vbs Administrators /m:.\HPSupport /c:url /l:"c:\hp\logs\Appoint Administrators.log"