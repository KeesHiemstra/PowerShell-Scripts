<#
    Monitor if computers are online and accesible.
#>
$MailSplatting = @{'SmtpServer' = 'smtp.corp.demb.com'; 'From' = 'Kees.Hiemstra@JDEcoffee.com'; 'To' = 'Kees.Hiemstra@hpe.com' }

$List = Import-Csv -Path 'D:\Data\ConnectComputer.csv'

foreach ( $Item in $List)
{
    if ( (Test-Connection -ComputerName $Item.ComputerName -Quiet -Count 1 ) )
    {
        try
        {
            if ( Get-WmiObject -ComputerName $Item.ComputerName -Class Win32_OperatingSystem -ErrorAction Stop )
            {
                $Item.Online = $true
            }
        }
        catch
        {
            $Item.Online = $false
        }
    }   
}

$List | Where-Object { $_.Online -eq $true } | Select-Object ComputerName, Action | Send-ObjectAsHTMLTableMessage -Subject 'Computers online' -MessageType Report @MailSplatting
