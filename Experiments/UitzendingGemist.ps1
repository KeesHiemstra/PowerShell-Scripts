# Uitzending gemist

<#
.SYNOPSIS
   Extract the program data from the web page.
.OUTPUTS
   [PowerShell.CHi.ProgramData]
#>
function Split-ProgramData
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                   SupportsShouldProcess=$true, 
                   PositionalBinding=$true,
                   ConfirmImpact='Medium')]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        $ProgramData
    )

    Begin
    {
    }
    Process
    {
        $PD = [string[]]($ProgramData).InnerText.Split("`n")

        $Props = [ordered]@{
            'Title'    = [string]$PD[0].Trim()
            'SubTitle' = [string]$PD[4].Trim()
            'Text'     = [string]$PD[6].Trim()
            'Date'     = [string]$PD[8].Trim()
            'Station'  = [string]$PD[10].Trim()
            'Link'     = [string]"<a$(($ProgramData.InnerHtml -split '<a|</a>')[1])</a>"
        }
        $Return = New-Object -TypeName PSObject -Property $Props
        $Return.PSObject.TypeNames.Insert(0, 'PowerShell.CHi.ProgramData')

        Write-Output $Return
    }
    End
    {
    }
}

#Collect the data from yesterday
$Page = Invoke-WebRequest -Uri "https://www.uitzendinggemist.net/op/$((Get-Date).AddDays(-1).ToString('ddMMyyyy')).html"

$Progs = $Page.AllElements | 
    Where-Object {$_.TagName -eq 'DIV' -and $_.Class -eq 'kr_blok_main'} | 
    Split-ProgramData | 
    Where-Object { $_.Station -notin 'NET5','RTL4','RTL5','RTL7','RTL8','RTLXL','SBS6','VERONICA'}
    
