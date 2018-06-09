<#
    Update-DistributionScripts.ps1

    Handle download.

    --- Version history
    Version 1.00 (2016-05-03, Kees Hiemstra)
    - Initial version.
#>

$SOEToolsPath = (Get-WmiObject -Class Win32_Share | Where-Object { $_.Name -eq 'SOETools$' }).Path
if ([string]::IsNullOrEmpty($SOEToolsPath))
{
    $SOEToolsPath = "D:\SOETools"
}

Start-BitsTransfer -Source  'http://ITAMWeb.corp.demb.com/SOETools/Distribution/CommonEnvironment.ps1' -Destination $SOEToolsPath -Description 'SOETools' -DisplayName 'SOETools update'
Start-BitsTransfer -Source  'http://ITAMWeb.corp.demb.com/SOETools/Distribution/Send-StatusUpdate.ps1' -Destination $SOEToolsPath -Description 'SOETools' -DisplayName 'SOETools update'
Start-BitsTransfer -Source  'http://ITAMWeb.corp.demb.com/SOETools/Distribution/Start-ImageDownload.ps1' -Destination $SOEToolsPath -Description 'SOETools' -DisplayName 'SOETools update'
Start-BitsTransfer -Source  'http://ITAMWeb.corp.demb.com/SOETools/Distribution/Start-OfficeDownload.ps1' -Destination $SOEToolsPath -Description 'SOETools' -DisplayName 'SOETools update'
Start-BitsTransfer -Source  'http://ITAMWeb.corp.demb.com/SOETools/Distribution/Test-AfterReboot.ps1' -Destination $SOEToolsPath -Description 'SOETools' -DisplayName 'SOETools update'
Start-BitsTransfer -Source  'http://ITAMWeb.corp.demb.com/SOETools/Distribution/Test-Download.ps1' -Destination $SOEToolsPath -Description 'SOETools' -DisplayName 'SOETools update'
Start-BitsTransfer -Source  'http://ITAMWeb.corp.demb.com/SOETools/Distribution/Update-DistributionScripts.ps1' -Destination $SOEToolsPath -Description 'SOETools' -DisplayName 'SOETools update'
