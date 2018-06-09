#region Write-Log

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
function Write-Log
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  HelpUri = 'http://www.microsoft.com/',
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([String])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [string]
        $Message
    )

    Begin
    {
        $Log = Initialize-Log -LogIndex $MyInvocation.ScriptName
    }
    Process
    {
        $LogMessage = "{0} {1}" -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Message
        if ( -not [string]::IsNullOrEmpty($Log.Path) -and $pscmdlet.ShouldProcess("Write [$Message] to $($Log.Path)") )
        {
            Add-Content -Path $Log.Path -Value $LogMessage
        }
        if ( $VerbosePreference -eq 'Continue' ) { Write-Host $LogMessage }
        elseif ( $log.ConsoleOutput ) { Write-Host $Message }
    }
    End
    {
    }
}
#endregion
