<#
    
#>

#region Script variables

$List = @()

#region Mail object
$MailObjectPath = $MyInvocation.MyCommand.Definition -replace '.ps1', '.xml'
if ( Test-Path -Path $MailObjectPath )
{
    $MailTest = $false
    $Mail = Import-Clixml -Path $MailObjectPath
}
else
{
    #Mail settings are not yet set, create default settings
    $MailTest = $true

    #Get current user mail address
    $MailSearch = [adsisearcher]"(&(objectClass=user)(objectCategory=person)(samaccountname=$($env:Username)))"
    $MailAddress = $EmailSearch.FindOne().Properties.mail

    $Mail = @{ From='ComputerOnline@JDEcoffee.com'; To=$MailAddress; Smtp='smtp.corp.demb.com' }

    $Mail | Export-Clixml -Path $MailObjectPath
}
#endregion

#region List of computers to check
$ComputerListPath = $MyInvocation.MyCommand.Definition -replace '.ps1', '.csv'
if ( Test-Path -Path $ComputerListPath )
{
    #Import list

    $List = Import-Csv -Path $ComputerListPath -Delimiter ','
}
else
{
    #Nothing to check, create empty file

    $Properties = @{ ComputerName=$Env:ComputerName; Action='Test (don''t delete this entry)' }
    $List += New-Object -TypeName PSObject -Property $Properties
}

if ( ($List | Where-Object { $_.ComputerName -eq $Env:ComputerName }).Count -eq 0 )
{
    #Add this computer to list

    $Properties = @{ ComputerName=$Env:ComputerName; Action='Test (don''t delete this entry)' }
    $List += New-Object -TypeName PSObject -Property $Properties
}
#endregion

#endregion

#region Script cmdlets

function Send-OnlineMessage
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$false,
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    Param
    (
        #Object(s) turned into HTML table and send as HTML message.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [System.Management.Automation.PSObject[]]
        $InputObject
    )

    Begin
    {
        $Tabel = @()
    }
    Process
    {
        if ( $MailTest -or $InputObject.ComputerName -ne $Env:ComputerName )
        {
            $Tabel += $InputObject
        }
    }
    End
    {
        $Header = "<style>BODY{background-color:LightGreen;}"
        $Header += "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
        $Header += "TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:thistle}"
        $Header += "TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:PaleGoldenrod}"
        $Header += "</style>"

        if ( $MailTest )
        {
            $Subject = 'Test mail from ComputerOnline'
        }
        else
        {
            $Subject = 'Computer(s) currently online'
        }

        if ( $Tabel.Count -ne 0 )
        {
            [string]$Message = $Tabel | ConvertTo-Html -Title $Subject -Head $Header -Body "<H2>$Subject (#$([array]$List.Count))</H2>"

            try
            {
                Send-MailMessage -Body $Message -Subject $Subject -BodyAsHtml @Mail
            }
            catch
            {
                Write-Host 'Error sending mail'
            }
        }
    }
}

#endregion

#region Main script

#region Check setting, create if they do not exist
#endregion

#region Check computers online

foreach ( $Item in $List)
{
    Add-Member -InputObject $Item -MemberType NoteProperty -Name Online -Value $false

    if ( (Test-Connection -ComputerName $Item.ComputerName -Quiet -Count 1 ) )
    {
        try
        {
            if ( Get-WmiObject -ComputerName $Item.ComputerName -Class Win32_OperatingSystem -ErrorAction Stop )
            {
                $Item.Online = $true

                Write-Host "$($Item.ComputerName) is online"
            }
        }
        catch
        {
        }
    }   
}

[array]$List | Where-Object { $_.Online -eq $true } | Select-Object ComputerName, Action | Send-OnlineMessage

#Export list and always include own computer to avoid empty list
$List | Where-Object { $_.Online -eq $false -or $_.ComputerName -eq $Env:ComputerName } |
    Select-Object ComputerName, Action |
    Export-Csv -Path $ComputerListPath -Delimiter ',' -NoTypeInformation

#endregion

#endregion
