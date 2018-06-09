
PSEdit 'B:\AddComputersToGPO.txt'

break

$GPOName = 'HPE-SOE-C-Corporate VPN EMEA-t1'

Get-Content 'B:\AddComputersToGPO.txt' |
    ForEach-Object { Set-GPPermissions -Name $GPOName -TargetName $_ -TargetType Computer -PermissionLevel GpoApply }
