#region Convert-DistinguishedNameToName

<#
.SYNOPSIS
    Get the name element of the DistinguishedName.
.DESCRIPTION
    This function will get the name part of the DistinguishedName. Commas in the name (\,) will be kept as comma (,).
.EXAMPLE
    Get-ADUser Arthur.Pandragon | Convert-DistinguishedNameToName

    >> arthur.pandragon
.EXAMPLE
    Convert-DistinguishedNameToName -DistinguishedName 'Arthur\, King,OU=Users,DC=Camelot,DC=ay'

    AD will lead a comma in the name with a backslash (\,). This function will replace this combination with a single comma.

    >> Arthur, King
.EXAMPLE
    (Get-ADUser Arthur.Pandragon -Properties MemberOf).MemberOf | Convert-DistinguishedNameToName

    It can also get the names of the groups.

    >> Royal family
    >> Knights of the round table
    >> Threasury
.INPUTS
    [string[]]
.OUTPUTS
    [string]
.NOTES
    === Version history
    Version 1.00 (2017-01-03, Kees Hiemstra)
    - Initial version.
.COMPONENT
    Active Directory.
.ROLE

.FUNCTIONALITY

#>
function Convert-DistinguishedNameToName
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$true,
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([String])]
    Param
    (
        #DistinguishedName of the object in AD
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [Alias("DN")]
        [string[]]
        $DistinguishedName
    )

    Begin
    {
    }
    Process
    {
        foreach ( $Item in $DistinguishedName )
        {
            if ( [string]::IsNullOrEmpty($Item) )
            {
                $Result = ''
            }
            else
            {
                $Item = $Item.Replace( '\,', '|')
                $Result = ($Item -replace "CN=" -replace ",.*com$").Replace('|', ',')
            }
            Write-Output $Result
        }
    }
    End
    {
    }
}

#endregion
