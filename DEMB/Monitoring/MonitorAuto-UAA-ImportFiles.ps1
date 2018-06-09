$CheckTime = (Get-Date).AddHours(-1).AddMinutes(-30)

$Files = @()

$Folders = @('\\dembswaaps456.corp.demb.com\d$\KISS Project Folder\AD_Provisioning\idm_new_users\*.csv', '\\dembswaaps456.corp.demb.com\d$\KISS Project Folder\AD_Provisioning\idm_disabled_users')

foreach ( $Path in $Folders )
{
    $Files += Get-ChildItem -Path $Path |
        Where-Object { $_.CreationTimeUtc -lt $CheckTime }
}

