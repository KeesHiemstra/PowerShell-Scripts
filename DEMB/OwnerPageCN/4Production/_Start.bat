@Echo Off
Rem See _What.txt

XCopy \\corp.demb.com\Netlogon\SOE\OwnerPage\OwnerPage.ps1 C:\RPMTools\ownrPage\ /D /Y
PowerShell.exe -f C:\RPMTools\OwnrPage\OwnerPage.ps1
