#region Convert-DistinguishedNameToLocation

<#
.SYNOPSIS
    Get the location of the AD object without the domain name.
.DESCRIPTION
    This function will get the location part of the DistinguishedName, but witout the domain name. Commas in the name (\,) will be kept as comma (,).
.EXAMPLE
    Get-ADUser Arthur.Pandragon | Convert-DistinguishedNameToLocation

    >> OU=Users
.INPUTS
    [string[]]
.OUTPUTS
    [string]
.NOTES
    === Version history
    Version 1.00 (2017-01-08, Kees Hiemstra)
    - Initial version.
.COMPONENT
    Active Directory.
.ROLE

.FUNCTIONALITY

#>
function Convert-DistinguishedNameToPartition
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
                $Item = $Item.Replace( '\,', '|') -replace ',DC.*$'
                $Result = $Item -split ','
                $Result = ($Item -replace "$($Result[0]),").Replace('|', ',')
            }
            Write-Output $Result
        }
    }
    End
    {
    }
}

#endregion
