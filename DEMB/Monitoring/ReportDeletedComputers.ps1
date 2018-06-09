$ConnectionString = 'Trusted_Connection=True;Data Source=DEMBMCAPS032SQ2.corp.demb.com\Prod_2'

$DTDeletion = (Get-Date).AddDays(-7).Date

$SQL = @"
SELECT DISTINCT UPPER(O.[Name]) AS 'ComputerName'
FROM SOEAdmin.dbo.ADObject AS O
       JOIN SOEAdmin.dbo.ADPartition AS P
              ON O.[PartitionID] = P.[ID]
       LEFT OUTER JOIN (
              SELECT O.[Name]
              FROM SOEAdmin.dbo.ADObject AS O
                     JOIN SOEAdmin.dbo.ADPartition AS P
                           ON O.[PartitionID] = P.[ID]
              WHERE O.[ObjectClass] = 'Computer'
                     AND O.[DTDeletion] IS NULL
                     AND P.[PartitionName] NOT LIKE '%,OU=VDI,%'
              ) AS C
              ON O.[Name] = C.[Name]
WHERE O.[ObjectClass] = 'Computer'
       AND O.[DTDeletion] IS NOT NULL
       AND P.[PartitionName] NOT LIKE '%,OU=VDI,%'
       AND C.[Name] IS NULL
       AND O.[DTDeletion] >= @DTDeletion
"@

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

$Cmd = New-Object -TypeName System.Data.SqlClient.SqlCommand
$Cmd.Connection = $Conn
$Cmd.CommandText = $SQL

$Cmd.Parameters.Add('DTDeletion', [datetime]) | Out-Null
$Cmd.Parameters['DTDeletion'].Value = [datetime]$DTDeletion

$DeletedComputers = @()
$Index = 0
$Data = $Cmd.ExecuteReader()
while ($Data.Read())
{
    $Index++
    $Properties = [ordered]@{'#' = [int]$Index;'ComputerName' = [string] $Data['ComputerName'].ToString();}
    $Obj = New-Object -TypeName PSObject -Property $Properties
    $DeletedComputers += $Obj
}
$Data.Close()

try
{
    $Conn.Close()
}
catch
{
}

if ( $DeletedComputers.Count -eq 0 ) { break }

$Header = @"
<style>
BODY{background-color:lightgreen;}
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:thistle}
TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:PaleGoldenrod}
</style>
"@
$MsgTitle = "Deleted computers since $($DTDeletion.ToString('yyyy-MM-dd'))"

[string]$Message = $DeletedComputers | ConvertTo-Html -Title $MsgTitle -Head $Header -Body "<H2>$MsgTitle (#$($DeletedComputers.Count))</H2>"

if ( $PSScriptRoot -ne 'C:\Etc\Jobs' )
{
    #Test page in browser
    $Message | Out-File -FilePath "$($env:TEMP)\$MsgTitle.html"
    Invoke-Item "$($env:TEMP)\$MsgTitle.html"
}
else
{
    Send-MailMessage -Body $Message -BodyAsHtml -SmtpServer 'smtp.corp.demb.com' -From 'HPDesktop.Administrator@JDEcoffee.com' -To 'wps-ps-swm-sccm-bg@hpe.com' -Cc 'Kees.Hiemstra@hpe.com' -Subject $MsgTitle
}
