#region Add-CheckComputerOnline

<#
.SYNOPSIS
    Add the given computer name(s) to the ConnectComputer.csv file to check if those come online.
.DESCRIPTION

.EXAMPLE
    > Add-CheckComputerOnline -ComputerName AU5CG62278WQ -Action BitLocker?

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
function Add-CheckComputerOnline
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
        $ComputerName,

        #Action to be taken
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false,
                   ValueFromRemainingArguments=$false,
                   Position=1,
                   ParameterSetName='Parameter Set 1')]
        [string]
        $Action
    )

    Begin
    {
    }
    Process
    {
        foreach ( $Item in $ComputerName )
        {
            $Properties = @{'ComputerName' = $Item.ToUpper()
                            'Date'         = (Get-Date).ToString('yyyy-MM-dd')
                            'Action'       = $Action
                            'Online'       = $null
                           }
            if ($PSCmdlet.ShouldProcess($Item, 'Add to checklist'))
            {
                New-Object -TypeName PSObject -Property $Properties |
                    Export-Csv -Path '\\HPNLDev05.corp.demb.com\D$\Data\ConnectComputer.csv' -Append -Delimiter ',' -NoTypeInformation
            }
        }
    }
    End
    {
    }
}

#endregion
