<#
    Report weekly hour overview of Exit project to antoniya.tominska@dxc.com
#>

#region Variables
$ConnectionString = 'Trusted_Connection=True;Data Source=HPNLDev05.corp.demb.com'

$SQL = @"SELECT REPLACE(CAST(CAST(L.[DTStart] AS date) AS varchar(7)), '-', '') AS 'Month ID',
       'CATW T&M' AS 'Billing Mode',
       'WCSO/NLEU502897' AS 'Cost Center',
       'NL' AS 'Country',
       'NL1-DEMB1.46.11' AS 'WbsID',
       'EUC Transition and Exit Activities' AS 'WBS Description',
       W.[RegistryPath] AS 'Attribute',
       'NL1-DEMB1.46.11/Not assigned' AS 'Attribute 1',
       CAST(L.[DTStart] AS date) AS 'Activity Date',
       '1030961' AS 'Employee ID',
       'Hiemstra Kees' AS 'Employee Name',
       '85008IUL26' AS 'Employee Cost Center',
       '' AS 'Employee Activity Type',
       '' AS 'AA type',
       L.[Comment] AS 'Text',
       'EUR' AS 'Local Currency',
       '' AS 'Billed Amount Local Currency',
       CAST(CAST(CAST((SUM(L.[Duration]) + 7.5) / 15 AS int) * 15 AS numeric(8, 2)) / 60 AS numeric(8, 2)) AS 'Activity Hours',
       '' AS 'Billed Amount USD'
FROM WorkLog2.dbo.WorkLog AS L
       JOIN WorkLog2.dbo.Activity AS A
              ON L.[ActivityID] = A.[ID]
                     AND A.[ProcessID] = 94
       JOIN WorkLog2.dbo.WBS AS W
              ON A.[WBSID] = W.[ID]
GROUP BY CAST(L.[DTStart] AS date), W.[RegistryPath], L.[Comment]
ORDER BY CAST(L.[DTStart] AS date) DESC, W.[RegistryPath]
"@

#endregion

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

Write-Verbose $SQL

$Data = $Cmd.ExecuteReader()

$Report = @()
while ($Data.Read())
{
    $Properties = [ordered]@{'Month ID'                     = [string]$Data['Month ID']
                             'Billing Mode'                 = [string]$Data['Billing Mode']
                             'Cost Center'                  = [string]$Data['Cost Center']
                             'Country'                      = [string]$Data['Country']
                             'WbsID'                        = [string]$Data['WbsID']
                             'WBS Description'              = [string]$Data['WBS Description']
                             'Attribute'                    = [string]$Data['Attribute']
                             'Attribute 1'                  = [string]$Data['Attribute 1']
                             'Activity Date'                = [string]($Data['Activity Date']).ToString('yyyy-MM-dd')
                             'Employee ID'                  = [string]$Data['Employee ID']
                             'Employee Name'                = [string]$Data['Employee Name']
                             'Employee Activity Type'       = [string]$Data['Employee Activity Type']
                             'AA type'                      = [string]$Data['AA type']
                             'Text'                         = [string]$Data['Text']
                             'Local Currency'               = [string]$Data['Local Currency']
                             'Billed Amount Local Currency' = [string]$Data['Billed Amount Local Currency']
                             'Activity Hours'               = [string]$Data['Activity Hours']
                             'Billed Amount USD'            = [string]$Data['Billed Amount USD']
                            }
    $Return = New-Object -TypeName PSObject -Property $Properties
    $Return.PSObject.TypeNames.Insert(0, 'WorkLog.ExitProject.100')
    $Report += $Return
}
$Data.Close()

$Conn.Close()

$Report | Send-ObjectAsHTMLTableMessage -Subject 'JDE EUC Exit - Time Tracking' -SmtpServer 'smtp.corp.demb.com' -From 'Kees.Hiemstra@esfds.com' -MessageType Report -Bcc 'Kees.Hiemstra@JDECoffee.com' -Cc 'Kees.Hiemstra@hpcds.com' -To 'Antoniya.Tominska@dxc.com'
