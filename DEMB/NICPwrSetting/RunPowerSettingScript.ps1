<#
    Force to run the SetNICPwrSetting.ps1 to run under elivated rights.
#>
$LogFile = 'C:\Windows\Temp\SetNICPwrSetting.log'
$StopFile3050265 = 'C:\Windows\Temp\InstallKB3050265.txt'
$StopFile3102810 = 'C:\Windows\Temp\InstallKB3102810.txt'

function Write-Log([string]$Message)
{
    $LogMessage = "{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Add-Content -Path $LogFile -Value $LogMessage
}

function Write-Stop([string]$Message, [string]$FileName)
{
    Write-Log -Message $Message
    Add-Content -Path $FileName -Value $Message
}

Write-Log "Starting parent script"

#if (-not (Test-Path -Path $StopFile3102810))
#{
#    if (Get-WmiObject -Class “win32_quickfixengineering” -Filter "HotFixID = 'KB3102810'")
#    {
#        Write-Stop -Message "Update KB3102810 is already installed" -FileName $StopFile3102810
#    }
#    else
#    {
#        if (-not (Test-Path -Path "C:\Windows\Temp\Windows6.1-KB3102810-x64.msu"))
#        {
#            try
#            {
#                Copy-Item -Path "\\corp.demb.com\netlogon\SOE\Files\Windows6.1-KB3102810-x64.msu" -Destination "C:\Windows\Temp\Windows6.1-KB3102810-x64.msu"
#                Write-Log -Message "Copied installer locally (C:\Windows\Temp\Windows6.1-KB3102810-x64.msu)"
#            }
#            catch
#            {
#                Write-Log -Message "Could not copy installer locally (\\corp.demb.com\netlogon\SOE\Files\Windows6.1-KB3102810-x64.msu)"
#            }
#        }
#        else
#        {
#            Write-Log -Message "Installation file already copied locally (C:\Windows\Temp\Windows6.1-KB3102810-x64.msu)"
#        }
#    }
#}

Start-Process "$PSHome\Powershell.exe" -Verb RunAs -ArgumentList '-file "\\corp.demb.com\NetLogon\SOE\Scripts\SetNICPwrSetting.ps1"'