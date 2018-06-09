#region Get-Password

<#
.SYNOPSIS
    Decrypt the password from the provided credential object.

.DESCRIPTION
    Decryption works only in this way when the encryption is done on the same computer.

.EXAMPLE
    Get-Password -Credential $Cred

    >> This will return the password in plain text.
.INPUTS
    [System.Management.Automation.PSCredential]

.NOTES
   === Version history
   Version 1.00 (2015-11-10, Kees Hiemstra)
   - Initial version.
#>
function Get-Password
{
    [CmdletBinding()]
    [OutputType([string])]
    Param
    (
        # Credential for which the password needs to be decrypted.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false)]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        $Credential
    )

    Process
    {
        Write-Output $Credential.GetNetworkCredential().Password
    }
}

#endregion
