#region New-SOEComputer

<#
.SYNOPSIS
    Create an AD computer account for computers.
.DESCRIPTION
    The computer name is based the country code and serial number.
    The location is based on the country code and category.

    The (elivated) credentials are needed to set the specific ACLs for the account svc.SOE-Image.
    The domain controler is determined by the login server of the computer.
.EXAMPLE
    New-SOEComputer -CountryCode NL -Category Laptops -SerialNo XYZ1234567 -Credential $Cred

    Will create a new computer object for NLXYZ1234567 in OU=Laptops,OU=NL,OU=Win7,OU=Managed Computers,DC=corp,DC=demb,DC=com which the image can access.
.EXAMPLE
    Import-Csv -Path D:\Data\Computers.csv | New-SOEComputer -Credential $Cred

    Will create a new computer objects for the computers listed in in the input file.
.INPUTS

.OUTPUTS
DistignuishedName of the created computer account.

.NOTES

    === Version history
    Version 2.00 (2017-03-22, Kees Hiemstra)
    - Made it more generic.
    Version 1.00 (2015-04-27, Kees Hiemstra)
    - Initial version.

.COMPONENT

.ROLE

.FUNCTIONALITY

#>
function New-SOEComputer
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$true,
                  ConfirmImpact='Medium')]
    [OutputType([String])]
    Param
    (
        # ISO country code where the computer needs to be created
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [ValidateLength(2,2)]
        [string]
        $CountryCode,

        # Category of the computer, either Desktops or Laptops
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=1,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("Desktops", "Laptops")]
        [string]
        $Category,

        # Serial number of the computer that will be used to create the computer name.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=2,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [ValidateLength(10,10)]
        [string]
        $SerialNo,

        # Credentials with which the ACLs can be set.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false)]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        $Credential
    )

    Begin
    {
        if((Get-Module ActiveDirectory) -eq $null) { Import-Module ActiveDirectory }

        if((Get-PSDrive -Name ADElivated -ErrorAction SilentlyContinue) -ne $null)
        {
           Get-PSDrive -Name ADElivated | Remove-PSDrive
        }

        if ($pscmdlet.ShouldProcess("Active Directory", "Set special PSDrive"))
        {
            New-PSDrive -Name ADElivated -PSProvider ActiveDirectory -Root //RootDSE/ -Description "Elivated access to set ACLs" -Credential $Credential | Out-Null
        }

        $UserSID = (Get-ADUser -Identity "svc.SOE-Image").SID
    }
    Process
    {
        $ComputerName = "DE$CountryCode$SerialNo"
        $Path = "OU=$Category,OU=$CountryCode,OU=Win7,OU=Managed Computers,DC=corp,DC=demb,DC=com"

        if ($pscmdlet.ShouldProcess("$Path", "Create new computer object for $ComputerName"))
        {
            #Create new computer account
            $DN = (Get-ADComputer -Filter ("Name -eq '{0}'" -f $ComputerName)).DistinguishedName
            if($DN -eq $null)
            {
                New-ADComputer -Name $ComputerName -Path $Path

                Write-Verbose "Created new computer account $ComputerName"
                
                sleep 2
                $DN = (Get-ADComputer -Filter ("Name -eq '{0}'" -f $ComputerName)).DistinguishedName

                while ($DN -eq $null)
                {
                    sleep 3
                    $DN = (Get-ADComputer -Filter ("Name -eq '{0}'" -f $ComputerName)).DistinguishedName
                }
            }
            else
            {
                Write-Verbose "Computer account $ComputerName already exists"
            }

            Write-Verbose $DN

            #Set ACLs
            $ACL = Get-Acl -Path "ADElivated:$DN" -ErrorAction SilentlyContinue
            if($ACL -eq $null)
            {
                Write-Verbose "! ACL for $ComputerName not existing, retrying..."

                sleep 2
                $ACL = Get-Acl -Path "ADElivated:$DN" -ErrorAction SilentlyContinue

                while ($ACL -eq $null)
                {
                    sleep 3
                    $ACL = Get-Acl -Path "ADElivated:$DN" -ErrorAction SilentlyContinue
                }
            }

            Write-Verbose "Completing ACLs for $ComputerName"
            
            $Rule = New-Object System.DirectoryServices.ActiveDirectoryAccessRule ($UserSID, 'DeleteTree, ExtendedRight, Delete, GenericRead', 'Allow', [GUID]'00000000-0000-0000-0000-000000000000')
            $ACL.AddAccessRule($Rule)

            $Rule = New-Object System.DirectoryServices.ActiveDirectoryAccessRule ($UserSID, 'WriteProperty', 'Allow', [GUID]'4c164200-20c0-11d0-a768-00aa006e0529')
            $ACL.AddAccessRule($Rule)

            $Rule = New-Object System.DirectoryServices.ActiveDirectoryAccessRule ($UserSID, 'Self', 'Allow', [GUID]'f3a64788-5306-11d1-a9c5-0000f80367c1')
            $ACL.AddAccessRule($Rule)

            $Rule = New-Object System.DirectoryServices.ActiveDirectoryAccessRule ($UserSID, 'Self', 'Allow', [GUID]'72e39547-7b18-11d1-adef-00c04fd8d5cd')
            $ACL.AddAccessRule($Rule)

            $Rule = New-Object System.DirectoryServices.ActiveDirectoryAccessRule ($UserSID, 'WriteProperty', 'Allow', [GUID]'3e0abfd0-126a-11d0-a060-00aa006c33ed')
            $ACL.AddAccessRule($Rule)

            $Rule = New-Object System.DirectoryServices.ActiveDirectoryAccessRule ($UserSID, 'WriteProperty', 'Allow', [GUID]'bf967953-0de6-11d0-a285-00aa003049e2')
            $ACL.AddAccessRule($Rule)

            $Rule = New-Object System.DirectoryServices.ActiveDirectoryAccessRule ($UserSID, 'Extendedright', 'Allow', [GUID]'5f202010-79a5-11d0-9020-00c04fc2d4cf')
            $ACL.AddAccessRule($Rule)

            $Rule = New-Object System.DirectoryServices.ActiveDirectoryAccessRule ($UserSID, 'WriteProperty', 'Allow', [GUID]'bf967953-0de6-11d0-a285-00aa003049e2')
            $ACL.AddAccessRule($Rule)

            $ACL | Set-Acl "ADElivated:$DN"

            Write-Verbose "ACLs set for $ComputerName"

            Write-Output $DN
        }
    }
    End
    {
        if((Get-PSDrive -Name ADElivated -ErrorAction SilentlyContinue) -ne $null)
        {
           Get-PSDrive -Name ADElivated | Remove-PSDrive
        }
    }
}

#endregion
