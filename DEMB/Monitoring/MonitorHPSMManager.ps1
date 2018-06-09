$ConnectionString = 'Trusted_Connection=True;Data Source=DEMBMCAPS032SQ2.corp.demb.com\Prod_2'

$Conn = New-Object -TypeName System.Data.SqlClient.SqlConnection
$Conn.ConnectionString = $ConnectionString
try
{
    $Conn.Open()
}
catch
{
    throw "Failed connection with connection string ($ConnectionString)"
}

$List = @()

$Cmd = New-Object -TypeName System.Data.SqlClient.SqlCommand
$Cmd.Connection = $Conn

$Cmd.CommandText = 'SELECT c, l, Manager FROM HPSMExport.dbo.Countries'

$Data = $Cmd.ExecuteReader()
while ($Data.Read())
{
    $Mail = $Data['Manager'].ToString()
    $ADUser = Get-ADUser -Filter { Mail -eq $Mail } -ErrorAction SilentlyContinue

    if ( $ADUser -eq $null -or $ADUser.Enabled -eq $false )
    {
        $Properties = [ordered]@{'Mail'     = $Mail
                                 'Country'  = $Data['c'].ToString()
                                 'Location' = $Data['l'].ToString()
                                 'Error'    = (if ( $ADUser -eq $null ) { 'Account does not exist' } else { 'Account is disabled' } )
                                 }
        $Obj = New-Object -TypeName PSObject -Property $Properties
        $List += $Obj
    }
}

try
{
    $Conn.Close()
}
catch
{
}

if ( $List.Count -eq 0 ) { break }

$Header = @"
<style>
BODY{background-color:peachpuff;}
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:thistle}
TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:PaleGoldenrod}
</style>
"@
$MsgTitle = 'Disabled managers in HPSM'

[string]$Message = $List | ConvertTo-Html -Title $MsgTitle -Head $Header -Body "<H2>$MsgTitle (#$($List.Count))</H2>"

if ( $PSScriptRoot -ne 'C:\Etc\Jobs' )
{
    #Test page in browser
    $Message | Out-File -FilePath "$($env:TEMP)\$MsgTitle.html"
    Invoke-Item "$($env:TEMP)\$MsgTitle.html"
}
else
{
    Send-MailMessage -Body $Message -BodyAsHtml -SmtpServer 'smtp.corp.demb.com' -From 'Kees.Hiemstra@JDEcoffee.com' -To 'Kees.Hiemstra@hpe.com' -Subject $MsgTitle
}


