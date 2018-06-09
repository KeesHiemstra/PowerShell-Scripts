
#region DistributionPoint

#region Get-DPImageVersion

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
function Get-DPImageVersion
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$false, 
                  PositionalBinding=$true,
                  ConfirmImpact='low')]
    [OutputType([String])]
    Param
    (
        # Get the Image version on the specified computer.
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$false, 
                   Position = 0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        $InputObject,

        # Get the Image version on the specified computer.
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$false, 
                   Position = 0,
                   ParameterSetName='Parameter Set 2')]
        [ValidateNotNullOrEmpty()]
        [string]
        $ComputerName,

        # Passthru
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [switch]
        $PassThru
    )

    Begin
    {
    }
    Process
    {
        $Result = "n/a"

        if($InputObject -ne $null)
        {
            $ComputerName = $InputObject.ComputerName
        }

        if((Test-Path -Path "\\$ComputerName\SOE_Image$\_Version.txt"))
        {
            $Result = (Get-Content -Path "\\$ComputerName\SOE_Image$\_Version.txt")[0] -replace ("Image version\s+: ", "")
        }

        if($PassThru)
        {
            Write-Output ($InputObject | Select-Object *, @{n='ImageVersion'; e={$Result}})
        }
        else
        {
            Write-Output $Result
        }
    }
    End
    {
    }
}
#endregion

#region Get-ADDistributionPoint

<#
.Synopsis
Get the AD details for all Distribution Points.
.DESCRIPTION
The SOE Central Administrators maintain a set of AD attributes to describe the SCCM Distribution Points (DP) to simplify the administration.

The Security group 'HP-SCCM-Servers' contains all SCCM server, but the cmdlet will only select the servers marked as DP based on the type attribute unless the parameter 'Only' is set to all.
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
- ComputerName
- c (Country ISO code)
- co (Country full name)
- l (location)
- description
- type [DP|DP Onsite|DP Server|DP OiaB]  
- delivContLength (Used for ordering the list)
- IPv4Address

.NOTES
   General notes
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet

.LINK
Set-ADDistributionPoint

#>
function Get-ADDistributionPoint
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$false, 
                  PositionalBinding=$false,
                  ConfirmImpact='low')]
    [OutputType([PSObject])]
    Param
    (
        # The computer name of the Distribution Point.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [string]
        $ComputerName,

        # The ISO country code.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=1,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [Alias("c")]
        [string]
        $CountryCode,

        # The country name.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=2,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [Alias("co")]
        [string]
        $CountryName,

        # The description.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=3,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [Alias("desc")]
        [string]
        $Description,

        # The name of the location (city).
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=4,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [Alias("l")]
        [string]
        $LocationName,

        # The type of the distribution point.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=7,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [string]
        $Type,

        # Only return the Distribution Points from Managed (Countries), Non Managed (Countries) or OiaB servers (Office in a Box).
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("Managed", "NonManaged", "OiaB", "All")]
        [string]
        $Only
    )

    Begin
    {
    }
    Process
    {
        $ManagedCountries = @('NL', 'BE', 'DE', 'FR', 'DK', 'AU', 'ES', 'HU', 'GB', 'SE', 'NZ', 'LV')

        $Result = Get-ADGroupMember "CN=HP-SCCM-Servers,OU=Server Groups,OU=Groups,OU=HP,OU=Support,DC=corp,DC=demb,DC=com" |
            Get-ADComputer -Properties Name, c, co, l, description, type, delivContLength, IPv4Address, LastLogonDate |
            Select-Object @{n='ComputerName'; e={$_.Name}}, c, co, l, description, type, delivContLength, IPv4Address, LastLogonDate

        if(-not [string]::IsNullOrWhiteSpace($ComputerName))  { $Result = $Result | Where-Object {$_.ComputerName -like $ComputerName} }
        if(-not [string]::IsNullOrWhiteSpace($CountryCode))   { $Result = $Result | Where-Object {$_.c -like $CountryCode} }
        if(-not [string]::IsNullOrWhiteSpace($CountryName))   { $Result = $Result | Where-Object {$_.co -like $CountryName} }
        if(-not [string]::IsNullOrWhiteSpace($Description))   { $Result = $Result | Where-Object {$_.description -like $Description} }
        if(-not [string]::IsNullOrWhiteSpace($LocationName))  { $Result = $Result | Where-Object {$_.l -like $LocationName} }
        if(-not [string]::IsNullOrWhiteSpace($Type))          { $Result = $Result | Where-Object {$_.c -eq $Type} }

        switch ($Only)
        {
            "Managed"    { $Result = $Result | Where-Object {$_.c -in $ManagedCountries} }
            "NonManaged" { $Result = $Result | Where-Object {$_.c -notin $ManagedCountries} }
            "OiaB"       { $Result = $Result | Where-Object {$_.type -eq 'DP OiaB'} }
        }

        if($Only -ne "All") { $Result = $Result | Where-Object {$_.type -like 'DP*'} }

        Write-Output $Result
    }
    End
    {
    }
}
#endregion

#region Set-ADDistributionPoint

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
function Set-ADDistributionPoint
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  HelpUri = 'http://www.microsoft.com/',
                  ConfirmImpact='Medium')]
    [OutputType([String])]
    Param
    (
        # The computer name on which the attributes will be set on.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [string]
        $ComputerName,

        # The ISO country code.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=1,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [Alias("c")]
        [string]
        $CountryCode,

        # The country name.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=2,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [Alias("co")]
        [string]
        $CountryName,

        # The description.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=3,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [Alias("desc")]
        [string]
        $Description,

        # The name of the location (city).
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=4,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [Alias("l")]
        [string]
        $LocationName,

        # The priority in the distribution queue.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=5,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [Alias("prio")]
        [int]
        $Priority,

        # The serial number.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=6,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [Alias("SerialNo")]
        [String]
        $SerialNumber,

        # The type of the distribution point.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=7,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [string]
        $Type,

        # Specifies the user account credentials to use to perform this task. The default credentials are the credentials of the currently logged on user unless the cmdlet is run from an Active Directory PowerShell provider 
        # drive. If the cmdlet is run from such a provider drive, the account associated with the drive is the default.
        #
        # To specify this parameter, you can type a user name, such as "User1" or "Domain01\User01" or you can specify a PSCredential object. If you specify a user name for this parameter, the cmdlet prompts for a password. 
        #
        # You can also create a PSCredential object by using a script or by using the Get-Credential cmdlet. You can then set the Credential parameter to the PSCredential object The following example shows how to create 
        # credentials.
        #  $AdminCredentials = Get-Credential "Domain01\User01"
        #
        # The following shows how to set the Credential parameter to these credentials.
        #  -Credential $AdminCredentials
        #
        #If the acting credentials do not have directory-level permission to perform the task, Active Directory PowerShell returns a terminating error.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false)]
        [System.Management.Automation.PSCredential]
        $Credential
     )

    Begin
    {
        $Param = @{}
        if (-not [string]::IsNullOrEmpty($Credential))
        {
            $Param += @{'Credential' = $Credential}
        }
    }
    Process
    {
        $ADComputer = Get-ADComputer -Filter ("Name -eq '{0}'" -f $ComputerName) -Properties c, co, delivContLength, description, l, serialNumber, type, memberOf

        if($ADComputer.distinguishedName -like "*,OU=SCCM,OU=Servers,DC=corp,DC=demb,DC=com" -or 
            ($ADComputer.memberOf -eq "CN=HP-SCCM-Servers,OU=Server Groups,OU=Groups,OU=HP,OU=Support,DC=corp,DC=demb,DC=com"))
        {
            if(-not [string]::IsNullOrWhiteSpace($CountryCode)) {$ADComputer.c = $CountryCode}
            if(-not [string]::IsNullOrWhiteSpace($CountryName)) {$ADComputer.co = $CountryName}
            if(-not [string]::IsNullOrWhiteSpace($Description)) {$ADComputer.description = $Description}
            if(-not [string]::IsNullOrWhiteSpace($LocationName)) {$ADComputer.l = $LocationName}
            if(-not $Priority -eq 0) {$ADComputer.delivContLength = $Priority}
            if(-not [string]::IsNullOrWhiteSpace($SerialNumber)) {$ADComputer.serialNumber = $SerialNumber}
            if(-not [string]::IsNullOrWhiteSpace($Type)) {$ADComputer.type = $Type}

            if ($pscmdlet.ShouldProcess($ComputerName, "Set attributes on the distribution point AD object"))
            {
                Set-ADComputer -Instance $ADComputer @Param 
            }#ShoudProcess

            Get-ADComputer -Filter ("Name -eq '{0}'" -f $ComputerName) -Properties c, co, delivContLength, description, l, serialNumber, type, memberOf
        }#Match group or group
        else
        {
            Write-Error "$ComputerName is not a distribution point"
        }
    }
    End
    {
    }
}

#endregion

#endregion DistributionPoint

Export-ModuleMember -Function Get-DPImageVersion
Export-ModuleMember -Function Get-ADDistributionPoint
Export-ModuleMember -Function Set-ADDistributionPoint

