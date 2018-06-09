#
# Script.ps1
#

Get-WmiObject -Class win32_bios | Select PSComputerName,SerialNumber
