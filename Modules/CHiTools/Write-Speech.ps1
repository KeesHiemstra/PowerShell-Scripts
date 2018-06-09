<#
.SYNOPSIS
    Use text to speech to read the message out loud.
.DESCRIPTION

.EXAMPLE
    Write-Speech -Message "The task is completed at $((Get-Date).ToString('hh:mm'))"

    ---
    "The taks is completed at <current time>" will be spoken out loud.
.NOTES
    === Version history
    Version 1.00 (2017-04-24, Kees Hiemstra)
    - Initial version.

#>
function Write-Speech
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$false,
                  PositionalBinding=$true,
                  ConfirmImpact='Low')]
    [Alias()]
    [OutputType([String])]
    Param
    (
        # Message to speak out loud
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$false,
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        $Message
    )

    Begin
    {
        Add-Type -AssemblyName System.speech
    }
    Process
    {
        $Synth = New-Object System.Speech.Synthesis.SpeechSynthesizer
        $Synth.SetOutputToDefaultAudioDevice();

        # Speak a string synchronously.
        $null = $Synth.SpeakAsync($Message)
    }
    End
    {
    }
}

#Write-Speech -Message "The task is completed at $((Get-Date).ToString('hh:mm'))"