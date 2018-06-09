#region Get-ADManagerDisplayName

<#
.SYNOPSIS
    Get the display name of User's manager or return 'Invalid' if the manager is not known or the manager's account is disabled.
.DESCRIPTION
    This function will get the Object behind the given DistinguishedName and will return for the following objects
    User:     '<DisplayName>'
    Group:    'Group: <SAMAccountName>'
    Computer: 'Computer: <Name>'

    The function will return nothing in case the DisplayName = 'position, default'
.EXAMPLE
    Get-ADManagerDisplayName -Manager 'CN=Merlin,OU=Users,DC=Camelot,DC=ay'

   >> Merlin the Sorcerer
.EXAMPLE
    Get-ADUser Merlin -Properties Manager | Get-ADManagerDisplayName

    >> Pandragon, Arthur
.INPUTS
    [sting]
.OUTPUTS
    [sting]
.NOTES
    === Version history
    Version 1.00 (2017-01-03, Kees Hiemstra)
    - Initial version.
.COMPONENT
    Active Directory
.ROLE

.FUNCTIONALITY

#>
function Get-ADManagerDisplayName
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    [OutputType([String])]
    Param
    (
        #DistinguishedName from Manager attribute
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [string]
        $Manager,

        # Param3 help description
        [Parameter(ParameterSetName='Parameter Set 1')]
        [switch]
        $PersonOnly
    )

    Begin
    {
    }
    Process
    {
        if ( [string]::IsNullOrEmpty($Manager) )
        {
            Write-Verbose 'Manager is empty or null'
            Return
        }
        Write-Verbose "Manager input: $Manager"
        try
        {
            $ADObject = Get-ADObject $Manager -Properties SAMAccountName, DisplayName -ErrorAction Stop
            Write-Verbose "Object found: $($ADObject.Name)"

            switch ( $ADObject.ObjectClass )
            {
                'user'     { if ( $ADObject.DisplayName -ne 'position, default' ) { Write-Output "$($ADObject.DisplayName)" } else { Write-Verbose 'Default position' } }
                'group'    { if ( -not $PersonOnly.ToBool()) { Write-Output "Group: $($ADObject.SAMAccountName)" } }
                'computer' { if ( -not $PersonOnly.ToBool()) { Write-Output "Computer: $($ADObject.Name)" } }
            }
        }
        catch
        {
            Write-Verbose 'Object not found'
        }
    }
    End
    {
    }
}

#endregion
