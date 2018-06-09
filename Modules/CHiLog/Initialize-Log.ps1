#region Initialize-Log

<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
function Initialize-Log
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([String])]
    Param
    (
        #Log file file name
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false,
                   ValueFromRemainingArguments=$false,
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [Alias("LogFileName")]
        [string]
        $Path,

        #Show log to console
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false,
                   ValueFromRemainingArguments=$false,
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [switch]
        $ConsoleOutput,

        #Log index
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false,
                   ValueFromRemainingArguments=$false, 
                   Position=1,
                   ParameterSetName='Parameter Set 1')]
        [string]
        $LogIndex
    )

    Begin
    {
        if ( -not $Script:Log )
        {
            $Script:Log = @{}
            Write-Verbose 'Created `$Script:Log'
        }
    }
    Process
    {
        if ( -not $MyInvocation.BoundParameters.ContainsKey('LogIndex') )
        {
            $LogIndex = $MyInvocation.ScriptName
            Write-Verbose "LogIndex: $LogIndex"
        }

        if ( -not $Script:Log[$LogIndex] )
        {
            $Obj = New-Object -TypeName PSObject -Property $Properties
            $Script:Log += @{$LogIndex = $Obj}
            Write-Verbose "Added [$LogIndex] to `$Script:Log"
            $Script:Log
        }
        $Log = $Script:Log[$LogIndex]

        if ( $MyInvocation.BoundParameters.ContainsKey('Path') )
        {
            if ($PSCmdlet.ShouldProcess("Set `$Script:Log[].Path to [$Path]"))
            {
                $Log.Path = $Path
                
                Write-Verbose "Path set to [$($Log.Path)]"
            }#if ShouldProcess
        }#if Path

        if ( $MyInvocation.BoundParameters.ContainsKey('ConsoleOutput') )
        {
            $Log.ConsoleOutput = $ConsoleOutput.ToBool()
        }

        if ( [string]::IsNullOrWhiteSpace($Log.Path) )
        {
            if ( [string]::IsNullOrWhiteSpace($LogIndex) )
            {
                $Log.ConsoleOutput = $true
                $Log.Path = $null
            }
            else
            {
                $Log.Path = $LogIndex -replace "$([System.IO.Path]::GetExtension($LogIndex))$", '.log'
            }
        }

        Write-Output $Log
    }
    End
    {
    }
}

#endregion
