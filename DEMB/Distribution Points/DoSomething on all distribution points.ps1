<#
    So something on all Distribution Points.
#>

Get-ADGroupMember "CN=HP-SCCM-Servers,OU=Server Groups,OU=Groups,OU=HP,OU=Support,DC=corp,DC=demb,DC=com" |
    Get-ADComputer -Properties Name, Type |
    Where-Object { $_.Type -like 'DP*' } |
    ForEach-Object {
 
        Write-Host -Path "\\$_.Name\d$\soe_image$\DEMB Win7 Installation instruction v0.9.docx"
#        Remove-Item -Path "\\$_.ComputerName\d$\soe_image$\DEMB Win7 Installation instruction v0.9.docx"
#        Copy-Item -Path "\\DEMBMCAPS032FG1.corp.demb.com\SOE\Source\O365\Distribution\2016\UninstallO365.xml" -Destination "\\$($_.Name)\D$\SOE\O365\2016" -Force

#        if ( (Test-Path -Path "\\$_\SOEAdmin$\SOE") )
#        { 
#            Remove-Item -Path "\\$_\SOEAdmin$\SOE" -Recurse -Confirm:$false -Verbose 
#        }
    }