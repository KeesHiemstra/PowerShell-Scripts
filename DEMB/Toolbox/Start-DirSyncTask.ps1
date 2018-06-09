<#
    Start AAD Sync and monitor if the process is running on the server.

    === Version history
    Version 2.00 (2016-12-27, Kees Hiemstra)
    - Added monitoring to console.
    Version 1.00 (2015-09-28, Kees Hiemstra)
    - Initial version.
#>

$Seconds = 13

if ( Invoke-Command -ComputerName DEMBMCAPS608 -ScriptBlock { Get-Process -Name DirectorySyncClientCmd -ErrorAction SilentlyContinue } )
{ 
    Write-Host "$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) AAD Sync is already running..." -BackgroundColor Black -ForegroundColor Magenta
}
else 
{
    SchTasks /Run /S DEMBMCAPS608 /TN "Azure AD Sync Scheduler"
    Write-Host "$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) Started AAD Sync..." -BackgroundColor Black -ForegroundColor Cyan
    Start-Sleep -Seconds $Seconds
}

if ( Invoke-Command -ComputerName DEMBMCAPS608 -ScriptBlock { Get-Process -Name DirectorySyncClientCmd -ErrorAction SilentlyContinue } )
{ 
    Write-Host "$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) Running..." -BackgroundColor Black -ForegroundColor Green
}
else 
{
    Write-Host "$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) Not started or already finished" -BackgroundColor Black -ForegroundColor Magenta
    break
}

while ( Invoke-Command -ComputerName DEMBMCAPS608 -ScriptBlock { Get-Process -Name DirectorySyncClientCmd -ErrorAction SilentlyContinue } )
{ 
    Write-Host "$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) Still running..." -BackgroundColor Black -ForegroundColor Green
    Start-Sleep -Seconds $Seconds
}

if ( Invoke-Command -ComputerName DEMBMCAPS608 -ScriptBlock { Get-Process -Name DirectorySyncClientCmd -ErrorAction SilentlyContinue } )
{ 
    Write-Host "$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) Still running" -BackgroundColor Black -ForegroundColor Magenta
}
else 
{
    Write-Host "$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) Finished" -BackgroundColor Black -ForegroundColor Cyan
}
