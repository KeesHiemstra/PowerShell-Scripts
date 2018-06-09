<#
    New-DistributionPointShares

    This scripts setup the folders and shares for the distribution of Office 365 source files
    and the image.
    It is typically used to use on new Distribution Point.

    --- Version history
    Version 1.10 (2016-05-10, Kees Hiemstra)
    - Added the image share.
    Version 1.00 (2016-01-27, Kees Hiemstra)
    - Initial version.
#>

break

#Create share on 1 new distribution point.
$ComputerName = "DEMBETMHVH1025"

if ( -not (Get-WmiObject -Class Win32_Share -ComputerName $ComputerName -Filter "Name = 'SOETools$'") )
{
    New-Share -ComputerName $ComputerName -Name SOETools$ -Path D:\SOETools -Access '/GRANT:DEMB\SoE Central Admins,FULL'
}

if ( -not (Get-WmiObject -Class Win32_Share -ComputerName $ComputerName -Filter "Name = 'SOE_Image$'") )
{
    New-Share -ComputerName $ComputerName -Name SOE_Image$ -Path 'D:\SOE_Image$' -Access '/GRANT:Everyone,READ'
}

if ( -not (Get-WmiObject -Class Win32_Share -ComputerName $ComputerName -Filter "Name = 'SOE$'") )
{
    New-Share -ComputerName $ComputerName -Name SOE$ -Path 'D:\SOE' -Access '/GRANT:Everyone,READ'
}

if ( -not (Get-WmiObject -Class Win32_Share -ComputerName $ComputerName -Filter "Name = 'SOEAdmin$'") )
{
    New-Share -ComputerName $ComputerName -Name SOEAdmin$ -Path 'D:\SOE' -Access '/GRANT:DEMB\SOE Central Admins,FULL'
}


break

#Create shares on all distributions points
(Get-ADDistributionPoint).ComputerName | ForEach-Object -Process {
    if ( -not (Get-WmiObject -Class Win32_Share -ComputerName $_ -Filter "Name = 'SOETools$'") )
    {
        New-Share -ComputerName $_ -Name SOETools$ -Path D:\SOETools -Access '/GRANT:DEMB\SoE Central Admins,FULL'
    }

    if ( -not (Get-WmiObject -Class Win32_Share -ComputerName $_ -Filter "Name = 'SOE_Image$'") )
    {
        New-Share -ComputerName $_ -Name SOE_Image$ -Path D:\SOE_Image$ -Access '/GRANT:Everyone,READ'
    }

    if ( -not (Get-WmiObject -Class Win32_Share -ComputerName $_ -Filter "Name = 'SOE$'") )
    {
        New-Share -ComputerName $_ -Name SOE$ -Path 'D:\SOE' -Access '/GRANT:Everyone,READ'
    }

    if ( -not (Get-WmiObject -Class Win32_Share -ComputerName $_ -Filter "Name = 'SOEAdmin$'") )
    {
        New-Share -ComputerName $_ -Name SOEAdmin$ -Path 'D:\SOE' -Access '/GRANT:DEMB\SOE Central Admins,FULL'
    }
}