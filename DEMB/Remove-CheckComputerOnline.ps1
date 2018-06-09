#region Remove-CheckComputerOnline

<#
.SYNOPSIS
    Remove the given computer name(s) from the ConnectComputer.csv checklist.
.DESCRIPTION

.EXAMPLE
    > Remove-CheckComputerOnline -ComputerName AU5CG62278WQ

    The computer AU5CG62278WQ will be added to the checklist.
.EXAMPLE

.INPUTS
    [string[]]
.OUTPUTS
    None
.NOTES
    ===Version History
    Version 1.00 (2017-02-15, Kees Hiemstra)
    - Inital version.
.COMPONENT
    Operations
.ROLE

.FUNCTIONALITY

#>
function Remove-CheckComputerOnline
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([String])]
    Param
    (
        #Computer name(s) to be checked online
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$false,
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [string[]]
        $ComputerName
    )

    Begin
    {
    }
    Process
    {
        $List = Import-Csv -Path '\\HPNLDev05.corp.demb.com\D$\Data\ConnectComputer.csv' -Delimiter ',' |
            Where-Object { $_.ComputerName -notin $ComputerName }

        if ($PSCmdlet.ShouldProcess($ComputerName -join '; ', 'Remove from checklist'))
        {
            $List |
                Export-Csv -Path '\\HPNLDev05.corp.demb.com\D$\Data\ConnectComputer.csv' -Delimiter ',' -NoTypeInformation
        }
    }
    End
    {
    }
}

#endregion
