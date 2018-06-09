<#
    Local Administrators through OU.ps1


#>

[array] $List = Get-ADOrganizationalUnit -Filter * -SearchBase 'OU=Win7,OU=Managed Computers,DC=corp,DC=demb,DC=com' -SearchScope Base -Properties Url |
    Select-Object Name, @{n='Level'; e={ 0 }}, Url, @{n='Source'; e={ 'Win7' }}

$List += Get-ADOrganizationalUnit -Filter * -SearchBase 'OU=Win7,OU=Managed Computers,DC=corp,DC=demb,DC=com' -SearchScope OneLevel -Properties Url |
    Where-Object { $_.Name -notin ('_Test') } |
    Select-Object Name, @{n='Level'; e={ 1 }}, Url, @{n='Source'; e={ "Win7\$($_.Name)" }}

$List |
    Select-Object @{n='Root'; e={ if ( $_.Level -eq 0 ) { $_.Name } }},
        @{n='Country'; e={ if ( $_.Level -eq 1 ) { $_.Name } }},
        @{n='Member'; e={ $_.url -join "`n" }} |
    ConvertTo-Csv -Delimiter "`t" -NoTypeInformation |
    Clip

break
$Script:GroupNo = 0

<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Report-Member
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        $Source,

        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        $Member,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=2)]
        $MemberOf,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=3)]
        [int]
        $Level = 0,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=4)]
        [int]
        $Group = 0
    )

    Begin
    {
    }
    Process
    {
        if ( $Level -eq 0 ) { $GroupNo = 0 }

        $ADObject = Get-ADObject -Filter { Name -eq $Member } -Properties MemberOf, DisplayName |
            Where-Object { $_.ObjectClass -in ( 'user', 'group' ) } |
            Select-Object Name, DisplayName, ObjectClass, MemberOf
        if ( $ADObject.ObjectClass -eq 'group' ) { $Level++ }

        Write-Output $ADObject |
            Select-Object @{n='Source'; e={ $Source }},
                @{n='Member'; e={ if ( $_.ObjectClass -eq 'user' ) { $_.DisplayName } else { $_.Name } }},
                @{n='Type'; e={ $_.ObjectClass }},
                @{n='MemberOf'; e={ $MemberOf }},
                @{n='Level'; e={ $Level }},
                @{n='Group'; e={ $Group }}

        if ( $ADObject.ObjectClass -eq 'group' )
        {
            $Script:GroupNo++
            foreach ( $Item in (Get-ADGroupMember $Member).Name | Convert-DistinguishedNameToName )
            {
                Report-Member -Source $Source -Member $Item -MemberOf $Member -Level $Level -Group $Script:GroupNo
            }
        }
    }
    End
    {
    }
}

$MemberList = @()

foreach ( $Item in $List )
{
    foreach ( $Member in $Item.Url )
    {
        Write-Output "$Member"
        $MemberList += Report-Member -Source $Item.Source -Member $Member | Sort-Object Group, Level, Type, MemberOf, Member
    }
}

$MemberList |
    Where-Object { $_.Type -in ( 'user', 'group' ) } |
    Select-Object @{n='Source OU'; e={ $_.Source }}, Member, Type, MemberOf |
    ConvertTo-Csv -Delimiter "`t" -NoTypeInformation |
    Clip
