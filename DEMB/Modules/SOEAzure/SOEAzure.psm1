<#
    Module SOEAzure

    PowerShell version: 3.0

    RfC: 1108791 Sync AD to Cloud.
    RfC: 1202852 Dinamo project.
    RfC: 1257953 Remove unwanted O365 license options

    Collection of cmdlets to manage licenses in the Microsoft Azure cloud.
#>

#region Helper functions

#All posible Tags in alphabetical order
$InAllTags = @('[ADFS]','[EMS]','[EXC]','[INT]','[Light]','[None]','[O365]','[PBI]','[Project]','[RMA]','[RMS]','[SHA]','[SWY]','[Visio]','[VPN]','[WAC]','[YAM]')
# Tags that are not allowed if the Tag [Light] is present
$InEnterpriseTags = @('[EXC]','[INT]','[O365]','[RMS]','[SHA]','[SWY]','[WAC]','[YAM]')
$InExtraTags = @('[Project]','[Visio]')
$InNeedADFSTags = @('[EMS]','[Light]')
# Azure related Tags
$InAzureTags = @('[EMS]','[EXC]','[INT]','[Light]','[O365]','[PBI]','[Project]','[RMA]','[RMS]','[SHA]','[SWY]','[Visio]','[WAC]','[YAM]')
# Group related Tags
$InGroupTags = @('[ADFS]','[EMS]','[VPN]')

#$InAllTags = @('[ADFS]','[VPN]','[EMS]','[EXC]','[INT]','[O365]','[RMS]','[SHA]','[SWY]','[WAC]','[YAM]','[Light]','[PBI]','[Project]','[Visio]','[RMA]')
#$InEnterpriseTags = @('[EXC]','[INT]','[O365]','[RMS]','[SHA]','[SWY]','[WAC]','[YAM]')
#$InExtraTags = @('[Project]','[Visio]')
#$InAzureTags = @('[EMS]','[EXC]','[INT]','[O365]','[RMS]','[SHA]','[SWY]','[WAC]','[YAM]','[Light]','[PBI]','[Project]','[Visio]','[RMA]')
#$InGroupTags = @('[ADFS]','[VPN]','[EMS]')

<#
    Remove all EnterprisePack related tags from the input.
#>
function RemoveEnterpriseTags ([string]$Tags)
{
    Remove-AzureLicenseTags -Tags $Tags -RemoveTags ($InEnterpriseTags+$InExtraTags)
}#RemoveEnterpriseTags

<#
    Check if one or more EnterprisePack Tags are present
#>
function HasEnterpriseTags ([string]$Tags)
{
    $Tags = $Tags.ToUpper()

    foreach ( $Tag in ($InEnterpriseTags) )
    {
        if ( $Tags.Contains($Tag.ToUpper()) )
        {
            Write-Output $true
            return
        }
    }#foreach

    Write-Output $false
}#HasEnterpriseTags

<#
    Check if one or more Tags are present that need access to the ADFS group
#>
function HasADFSTags ([string]$Tags)
{
    $Tags = $Tags.ToUpper()

    foreach ( $Tag in ($InEnterpriseTags+$InExtraTags+$InNeedADFSTags) )
    {
        if ( $Tags.Contains($Tag.ToUpper()) )
        {
            Write-Output $true
            return
        }
    }#foreach

    Write-Output $false
}#HasADFSTags

<#
    Sort and check all tags, the result will not have unknown tags.
#>
function CheckAndSortTags ([string]$Tags, [array]$CheckTags)
{
    $Result = ''
    $Tags = $Tags.ToUpper()

    foreach ( $Tag in $CheckTags )
    {
        if ( $Tags.Contains($Tag.ToUpper()) )
        {
            $Result += $Tag
        }
    }#foreach
    Write-Output $Result
}#Sort-Tags
#endregion

#region ConvertFrom-AzureLicense

<#
.Synopsis
   Convert Azure licenses to granular notation.
.DESCRIPTION
   This cmdlet converts the Azure license object into an PSObject containing the following properties:

   LicenseArrayFull : Array of string with the full name of the license, including the Azure domain name.
   LicenseArray     : Array of string with the license name only.
   LicensesText     : String with the sorted licenses names.
   GranularNotation : String with the sorted granular notation of all the license options.
   GranularArray    : Array of granular notation of all the license options.
.EXAMPLE
   $AzUser = Get-MsolUser -UserPrincipalName Arthur.PendragonPendragon@camelot.ay
   ConvertFrom-AzureLicense -Licenses $AzUser.Licenses

   Result:
    LicenseArrayFull : {camelot:VISIOCLIENT, camelot:ENTERPRISEPACK}
    LicenseArray     : {VISIOCLIENT, ENTERPRISEPACK}
    LicensesText     : ENTERPRISEPACK;VISIOCLIENT
    GranularNotation : [EXC][INT][O365][RMS][SHA][Visio][WAC][YAM]
    GranularArray    : {[Visio], [EXC], [INT], [O365]...}

.NOTES
    --- Version history:
    Version 2.01 (2016-08-12, Kees Hiemstra)
    - Set [SHA] and/or [WAC] in lowercase when these options are still PendingInput.
    Version 2.00 (2016-07-16, Kees Hiemstra)
    - Adding DesklessPack ([Light])
    Version 1.02 (2016-03-01, Kees Hiemstra)
    - Order all attributes from the result.
    Version 1.01 (2016-01-13, Kees Hiemstra)
    - Added Sway Service plan [SWY] under EnterprisePack.
    Version 1.00 (2015-10-05, Kees Hiemstra)
    - Initial version.

.COMPONENT
   SOEAzure module
#>
function ConvertFrom-AzureLicense
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$true,
                  ConfirmImpact='Low')]
    [OutputType([PSObject])]
    Param
    (
        # Azure license object
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [System.Collections.Generic.List[Microsoft.Online.Administration.UserLicense]]
        $Licenses
    )

    Begin
    {
        $Props = [ordered]@{LicenseArrayFull=(@()); LicenseArray=(@()); LicensesText=""; GranularNotation=""; GranularArray=(@())}
    }
    Process
    {
        $Result = New-Object -TypeName PSObject -Property $Props

        $Result.LicenseArrayFull = $Licenses.AccountSkuId | Sort-Object
        $Result.LicenseArray = $Licenses.AccountSkuId -replace "^\w*\:" | Sort-Object #remove the name up and until the colon
        $Result.LicensesText = ($Result.LicenseArray | Sort-Object) -join ";"

        $HasEnterprisePack = $false
        switch -Regex ($Result.LicensesText)
        {
            "DESKLESSPACK"           { $Result.GranularArray += '[Light]' }
            "EMS"                    { $Result.GranularArray += '[EMS]' }
            "ENTERPRISEPACK"         { $HasEnterprisePack = $true }
            "POWER_BI_STANDARD"      { $Result.GranularArray += '[PBI]' }
            "PROJECTCLIENT"          { $Result.GranularArray += '[Project]' }
            "RIGHTSMANAGEMENT_ADHOC" { $Result.GranularArray += '[RMA]' }
            "VISIOCLIENT"            { $Result.GranularArray += '[Visio]' }
        }

        if ($HasEnterprisePack)
        {
            $Options = ((($Licenses | 
                Where-Object { $_.AccountSkuID -like "*:ENTERPRISEPACK" }).ServiceStatus | 
                Where-Object { $_.ProvisioningStatus -ne "Disabled" }).ServicePlan.ServiceName | Sort-Object) -join ";"

            $OptionsPending = ((($Licenses | 
                Where-Object { $_.AccountSkuID -like "*:ENTERPRISEPACK" }).ServiceStatus | 
                Where-Object { $_.ProvisioningStatus -ne "Disabled" -and $_.ProvisioningStatus -ne "Success" }).ServicePlan.ServiceName | Sort-Object) -join ";"

            switch -Regex ($Options)
            {
                "EXCHANGE_S_ENTERPRISE" { $Result.GranularArray += "[EXC]" }
                "INTUNE_O365"           { $Result.GranularArray += "[INT]" }
                "MCOSTANDARD"           { $Result.GranularArray += "[MCO]" } #Lync online
                "OFFICESUBSCRIPTION"    { $Result.GranularArray += "[O365]" }
                "RMS_S_ENTERPRISE"      { $Result.GranularArray += "[RMS]" }
                "SHAREPOINTENTERPRISE"  { $Result.GranularArray +=  if ( $OptionsPending -like '*SHAREPOINTENTERPRISE*' ) { "~[sha]" } else { "[SHA]" } }
                "SWAY"                  { $Result.GranularArray += "[SWY]" }
                "SHAREPOINTWAC"         { $Result.GranularArray += if ( $OptionsPending -like '*SHAREPOINTWAC*' ) { "~[wac]" } else { "[WAC]" } }
                "YAMMER_ENTERPRISE"     { $Result.GranularArray += "[YAM]" }
            }

        }#Has EnterprisePack

        $Result.GranularNotation = ($Result.GranularArray | Sort-Object) -join ""

        Write-Output $Result
    }
    End
    {
    }
}

#endregion

#region Get-AzureLicense

<#
.Synopsis
   Get the user's Azure licenses and Azure account details.
.DESCRIPTION
   The Get-AzureLicense cmdlet gets the license details from Azure including the indication if he/she is member of the
   Allow-ADFS-Access and/or Allow-VPN-Access-Externals groups (important for external employees).
.EXAMPLE
   Get-AzureLicense Arthur.Pendragon

   Get the licenses for the user Arthur.Pendragon as a string [ADFS][EXC][INT][O365][RMS][SHA][Visio][VPN][WAC][YAM]
.EXAMPLE
   Get-AzureLicense Arthur.Pendragon@camelot.ay -ReturnResult Object

   Get the licenses for the user Arthur.Pendragon as an object.

    userPrincipalName : Arthur.PendragonPendragon@camelot.ay
    usageLocation     : AY
    blockCredential   : False
    azureStatus       : Normal user
    LicenseArray      : {VISIOCLIENT, ENTERPRISEPACK}
    Licenses          : VISIOCLIENT;ENTERPRISEPACK
    GranularNotation  : [EXC][INT][O365][RMS][SHA][Visio][VPN][WAC][YAM]
    GranularArray     : {[VPN], [Visio], [EXC]...}

.NOTES
   You need to be connected the to Azure with enough privaliges to read account details.


    --- Version history:
    Version 1.20 (2015-11-08, Kees Hiemstra)
    - Added azureStatus to the result of Get-AzureLicense indicating that the user is found or not or is in the grace period.
    Version 1.10 (2015-10-05, Kees Hiemstra)
    - Use ConvertFrom-AzureLicense.
    - Introduced the parameter -ReturnResult.
    Version 1.00 (2015-09-30, Kees Hiemstra)
    - Initial version.

.COMPONENT
   The component this cmdlet belongs to the SOEAzure module.

.ROLE
   The role this cmdlet belongs to Azure license management

.LINK
    Set-AzureLicense

.LINK
    ConvertFrom-AzureLicense

#>
function Get-AzureLicense
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$true,
                  ConfirmImpact='Low')]
    [OutputType([String])]
    Param
    (
        # Specify the AD identity of the user for which you want to get the (Azure) license.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [Alias("Id")]
        [String]
        $Identity,

        # Specify the userPrincipal of the user for which you want to get the (Azure) license.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 2')]
        [ValidateNotNullOrEmpty()]
        [Alias("UPN")]
        [String]
        $userPrincipalName,

        # Specify how the returned result should be represented, either as String (Granular notation) or the full Object
        [Parameter()]
        [ValidateSet("String", "Object")]
        [Alias("Result")]
        [String]
        $ReturnResult = "String",

        # Specify that the Granular notation should not contain the AD user groups
        [Parameter()]
        [Alias("NoGroups")]
        [switch]
        $NoADGroupCheck
    )

    Begin
    {
        $Props = [ordered]@{userPrincipalName=""; 
            usageLocation=""; 
            blockCredential=$true; 
            azureStatus="";
            LicenseArrayFull=(@()); 
            LicenseArray=(@()); 
            LicensesText=""; 
            GranularNotation=""; 
            GranularArray=(@());
            }
    }
    Process
    {
        $Result = New-Object -TypeName PSObject -Property $Props
        $ADUser = $null

        if ($Identity -like '*@*.*')
        {
            $userPrincipalName = $Identity
        }

        if (-not [String]::IsNullOrWhiteSpace($userPrincipalName))
        {
            $Result.userPrincipalName = $userPrincipalName
            if (-not $NoADGroupCheck)
            {
                Write-Verbose "Get AD object based on userPrincipalName $userPrincipalName"
                $ADUser = Get-ADUser -Filter "userPrincipalName -eq '$userPrincipalName'" -Properties userPrincipalName, memberOf -ErrorAction Stop
            }
        }
        else
        {
            Write-Verbose "Get AD object based on identity $Identity"
            $ADUser = Get-ADUser -Identity $Identity -Properties userPrincipalName, memberOf -ErrorAction Stop
        }

        if ($ADUser -ne $null)
        {
            $Result.userPrincipalName = $ADUser.userPrincipalName
        }

        if (-not $NoADGroupCheck)
        {
            $Groups = ";$($ADUser.MemberOf -replace "CN=" -replace ",.*com$" -join ";");"

            switch -Regex ($Groups)
            {
                ";Allow-VPN-Access-Externals;" { $Result.GranularArray += "[VPN]" }
            }
        }#NoUserGroupCheck

        if (-not [string]::IsNullOrWhiteSpace($Result.userPrincipalName))
        {
            Write-Verbose "Get Azure object based on userPrincipalName $userPrincipalName"
            $AzUser = Get-MsolUser -UserPrincipalName $Result.userPrincipalName -ErrorAction SilentlyContinue
            if ($AzUser -ne $null)
            {
                $Result.azureStatus = 'Normal user'
            }
            else
            {
                Write-Verbose "Get graced Azure object based on userPrincipalName $userPrincipalName"
                $AzUser = Get-MsolUser -UserPrincipalName $Result.userPrincipalName -ReturnDeletedUsers -ErrorAction SilentlyContinue
                if($AzUser -ne $null)
                {
                    $Result.azureStatus = 'Graced user'
                }
                else
                {
                    $Result.azureStatus = '<no Azure user>'
                }
            }

            if ($AzUser.Licenses.Count -ne 0)
            {
                $AzLicense = ConvertFrom-AzureLicense -Licenses $AzUser.Licenses

                $Result.LicenseArrayFull += $AzLicense.LicenseArrayFull
                $Result.LicenseArray += $AzLicense.LicenseArray
                $Result.LicensesText = ($Result.LicenseArray | Sort-Object) -join ";"

                $Result.GranularArray += $AzLicense.GranularArray
            }

            $Result.blockCredential = $AzUser.blockCredential
            $Result.usageLocation = $AzUser.usageLocation

        }#Has userPrincipalName
        $Result.GranularNotation = ($Result.GranularArray | Sort-Object) -join ""

        switch ($ReturnResult)
        {
            "String" { Write-Output $Result.GranularNotation }
            "Object" { Write-Output $Result }
        }
    }
    End
    {
    }
}

#endregion

#region Get-AzureLicenseDetail

<#
.Synopsis
    Get all Azure license details from the selected user.
.DESCRIPTION
    Licenses have options that not necessarely are managed by the Azure sync script. This function will show all the details.
.EXAMPLE
    Get-AzureLicenseDetail -UserPrincipalName Arthur.Pendragon@camelot.ay | Format-Table * -AutoSize
    
    UserPrincipalName           AccountSkuID              ServiceName                 ServiceType                   TargetClass Status           
    -----------------           ------------              -----------                 -----------                   ----------- ------           
    Arthur.Pendragon@camelot.ay camelot:EMS                                                                                                      
    Arthur.Pendragon@camelot.ay camelot:EMS               RMS_S_PREMIUM               RMSOnline                     User        Success          
    Arthur.Pendragon@camelot.ay camelot:EMS               INTUNE_A                    SCO                           User        PendingInput     
    Arthur.Pendragon@camelot.ay camelot:EMS               RMS_S_ENTERPRISE            RMSOnline                     User        Success          
    Arthur.Pendragon@camelot.ay camelot:EMS               AAD_PREMIUM                 AADPremiumService             User        Success          
    Arthur.Pendragon@camelot.ay camelot:EMS               MFA_PREMIUM                 MultiFactorService            User        Success          
    Arthur.Pendragon@camelot.ay camelot:ENTERPRISEPACK                                                                                           
    Arthur.Pendragon@camelot.ay camelot:ENTERPRISEPACK    FLOW_O365_P2                ProcessSimple                 User        Success          
    Arthur.Pendragon@camelot.ay camelot:ENTERPRISEPACK    POWERAPPS_O365_P2           PowerAppsService              User        Success          
    Arthur.Pendragon@camelot.ay camelot:ENTERPRISEPACK    TEAMS1                      TeamspaceAPI                  User        Success          
    Arthur.Pendragon@camelot.ay camelot:ENTERPRISEPACK    PROJECTWORKMANAGEMENT       ProjectWorkManagement         User        Success          
    Arthur.Pendragon@camelot.ay camelot:ENTERPRISEPACK    SWAY                        Sway                          User        Success          
    Arthur.Pendragon@camelot.ay camelot:ENTERPRISEPACK    INTUNE_O365                 SCO                           Tenant      PendingActivation
    Arthur.Pendragon@camelot.ay camelot:ENTERPRISEPACK    YAMMER_ENTERPRISE           YammerEnterprise              User        Disabled         
    Arthur.Pendragon@camelot.ay camelot:ENTERPRISEPACK    RMS_S_ENTERPRISE            RMSOnline                     User        Success          
    Arthur.Pendragon@camelot.ay camelot:ENTERPRISEPACK    OFFICESUBSCRIPTION          MicrosoftOffice               User        Success          
    Arthur.Pendragon@camelot.ay camelot:ENTERPRISEPACK    MCOSTANDARD                 MicrosoftCommunicationsOnline User        Disabled         
    Arthur.Pendragon@camelot.ay camelot:ENTERPRISEPACK    SHAREPOINTWAC               SharePoint                    User        Success          
    Arthur.Pendragon@camelot.ay camelot:ENTERPRISEPACK    SHAREPOINTENTERPRISE        SharePoint                    User        Success          
    Arthur.Pendragon@camelot.ay camelot:ENTERPRISEPACK    EXCHANGE_S_ENTERPRISE       Exchange                      User        Success          
    Arthur.Pendragon@camelot.ay camelot:POWER_BI_STANDARD                                                                                        
    Arthur.Pendragon@camelot.ay camelot:POWER_BI_STANDARD EXCHANGE_S_FOUNDATION       Exchange                      Tenant      PendingActivation
    Arthur.Pendragon@camelot.ay camelot:POWER_BI_STANDARD BI_AZURE_P0                 AzureAnalysis                 User        Success          
    Arthur.Pendragon@camelot.ay camelot:PROJECTCLIENT                                                                                            
    Arthur.Pendragon@camelot.ay camelot:PROJECTCLIENT     EXCHANGE_S_FOUNDATION       Exchange                      Tenant      PendingActivation
    Arthur.Pendragon@camelot.ay camelot:PROJECTCLIENT     PROJECT_CLIENT_SUBSCRIPTION MicrosoftOffice               User        Success          
    Arthur.Pendragon@camelot.ay camelot:VISIOCLIENT                                                                                              
    Arthur.Pendragon@camelot.ay camelot:VISIOCLIENT       EXCHANGE_S_FOUNDATION       Exchange                      Tenant      PendingActivation
    Arthur.Pendragon@camelot.ay camelot:VISIOCLIENT       VISIO_CLIENT_SUBSCRIPTION   MicrosoftOffice               User        Success          

.NOTES
    === Version history
    Version 1.00 (2016-12-13, Kees Hiemstra)
    - Inital version.
.COMPONENT
   The component this cmdlet belongs to the SOEAzure module.

.ROLE
   The role this cmdlet belongs to Azure license management

.LINK
    Get-AzureLicense

.LINK
    Set-AzureLicense

.LINK
    ConvertFrom-AzureLicense
#>
function Get-AzureLicenseDetail
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
        [Alias("UPN")]
        [string]
        $UserPrincipalName
    )

    Begin
    {
        $Properties = [ordered] @{'UserPrincipalName' = [string] ''
                                  'AccountSkuID'      = [string] ''
                                  'ServiceName'       = [string] ''
                                  'ServiceType'       = [string] ''
                                  'TargetClass'       = [string] ''
                                  'Status'            = [string] ''
                                 }
    }
    Process
    {
        $License = (Get-MsolUser -UserPrincipalName $UserPrincipalName).Licenses | Sort-Object AccountSkuID

        foreach ( $Item in $License )
        {
            $Obj = New-Object -TypeName PSObject -Property $Properties
            $Obj.UserPrincipalName = $UserPrincipalName
            $Obj.AccountSkuID      = $Item.AccountSkuID
            Write-Output $Obj

            foreach ( $Option in $Item.ServiceStatus | Sort-Object ServicePlan.SeriveName )
            {
                $Obj = New-Object -TypeName PSObject -Property $Properties
                $Obj.UserPrincipalName = $UserPrincipalName
                $Obj.AccountSkuID      = $Item.AccountSkuID
                $Obj.ServiceName       = $Option.ServicePlan.ServiceName
                $Obj.ServiceType       = $Option.ServicePlan.ServiceType
                $Obj.TargetClass       = $Option.ServicePlan.TargetClass
                $Obj.Status            = $Option.ProvisioningStatus
        
                Write-Output $Obj
            }
        }
    }
}

#endregion

#region Set-AzureLicense

<#
.Synopsis
   Set Azure license and usageLocation for the specified UserPrincipalName if these aren't already set this way.
.DESCRIPTION
   The cmdlet will check the license and usageLocation that needs to be set for the user specified by their userPrincipalName and update the Azure licenses accordingly.

   Possible License Tags:
   [Visio]
   [Project]
   [Light]
   [EMS]  
   [PBI]
   [RMA]
   [EXC]  
   [INT]
   [MCO] (Should not be used)
   [O365]
   [RMS]   
   [SHA]
   [SWY]  
   [WAC]   
   [YAM]   
.EXAMPLE
   set-AzureLicense -userPrincipalName "Arthur.Pendragon@camelot.ay" -License "[EMS][EXC][INT][O365][RMS][SHA][SWY][WAC][YAM]" -CountryCode "AY"

    This will set the licenses as 'full user' for the user Arthur Pendragon.
.EXAMPLE
   set-AzureLicense -userPrincipalName "Arthur.Pendragon@camelot.ay" -License "[O365][EXC][VISIO]" -CountryCode "AY" -Force

    This will force set the license EnterprisePack with the active options for office and a mailbox, and the license for Visio for the user Arthur Pendragon.

.OUTPUTS
    The output string contains the license(s) that were present before the change separated with a semicolon.

.NOTES
    The user needs to be logged on to Azure. It will fail the user in not logged on with the correct access rights.

    When the account doesn't have a mailbox, one is created but with a Microsoft address instead of a JDE address. This should be avoided.

    Some option can't be combined together (like DesklessPack and EnterPricePack). This cmdlet doesn't contain these checks, this can for instance be done with Test-AzureLicense.

    --- Version history:
    Version 2.03 (2017-05-05, Kees Hiemstra)
    - Bug fix: Replaced previous bug fix with $Global:Error.RemoveAt(0).
    Version 2.02 (2017-04-14, Kees Hiemstra)
    - Bug fix: Don't report the error created by Set-License -RemoveLicense on EMS becasue removing EMS has a huge delay of several hours.
    Version 2.01 (2016-09-01, Kees Hiemstra)
    - Bug fix: Upgrading License from Light to Full failed.
    Version 2.00 (2016-07-06, Kees Hiemstra)
    - Removed profiles.
    - Renamed the parameter License to LicenseTags.
    Version 1.50 (2016-02-15, Kees Hiemstra)
    - Changed CountryCode to optional.
    Version 1.40 (2016-02-02, Kees Hiemstra)
    - Added [SWY] as option within the enterprise pack.
    - YAMMER_ENTERPRISE can now be changed in Azure (was not possible in the past).
    Version 1.30 (2015-10-22, Kees Hiemstra)
    - Adding [EMS] as default to the OE3 profile.
      The request comes from Richard Wesseling. JDE has decided to add [EMS] to the OE3 profile.
    Version 1.20 (2015-09-20, Kees Hiemstra)
    - Throw terminating error on accessing Azure (Set-AzureLicense).
    Version 1.10 (2015-09-09, Kees Hiemstra)
    - Added the list of licenses (semicolon separated) before the change took place as output (Set-AzureLicense).
    Version 1.00 (2015-09-07, Kees Hiemstra)
    - Initial version.

.COMPONENT
   The component this cmdlet belongs to the SOEAzure module.

.ROLE
   The role this cmdlet belongs to Azure license management

.LINK
    Test-AzureLicenseTags

.LINK
    Connect-MsolService

#>
function Set-AzureLicense
{
    [CmdletBinding(SupportsShouldProcess=$true, 
                  PositionalBinding=$true,
                  ConfirmImpact='medium')]
    [OutputType([string])]
    Param
    (
        #userPrincipalName from Active Directory that is already known in the Azure cloud.
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=0)]
        [string]
        $userPrincipalName,

        #License tags to be checked and set if applicable.
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=1)]
        [string]
        $LicenseTags,

        #CountryCode of the user (usageLocation).
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=2)]
        [string]
        $CountryCode,

        #Force to set the UsageLocation and Licenses.
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false,
                   Position=3)]
        [switch]
        $Force
    )

    Begin
    {
    }
    Process
    {
        $Result = "<nothing>"

        $HasVisio      = $false
        $HasProject    = $false
        $HasDeskless   = $false
        $HasEntPack    = $false
        $HasEMS        = $false
        $HasPBI        = $false
        $HasRMAdHoc    = $false
        $HasEntPackOps = ""

        #Set license and options
        $SetVisio      = $false
        $SetProject    = $false
        $SetDeskless   = $false
        $SetEntPack    = $false
        $SetEMS        = $false
        $SetPBI        = $false
        $SetRMAdHoc    = $false
        $SetEntPackOps = "RMS_S_ENTERPRISE,OFFICESUBSCRIPTION,MCOSTANDARD,SHAREPOINTWAC,SHAREPOINTENTERPRISE,EXCHANGE_S_ENTERPRISE,SWAY,YAMMER_ENTERPRISE"

        $Force = $Force.IsPresent

        if ($LicenseTags -ne '' -and $LicenseTags -notlike "!*")
        {
            #Translate Licence Tags to actions
            switch -regex ($LicenseTags)
            {
                "\[VISIO\]"    { $SetVisio    = $true }
                "\[PROJECT\]"  { $SetProject  = $true }
                "\[LIGHT\]"    { $SetDeskless = $true }
                "\[EMS\]"      { $SetEMS      = $true }
                "\[PBI\]"      { $SetPowerBI  = $true }
                "\[RMA\]"      { $SetRMAdHoc  = $true }
                #Enterprise Options are not to be set, but to be removed, hence the replace action
                "\[EXC\]"      { $SetEntPack  = $true; $SetEntPackOps = $SetEntPackOps -replace "EXCHANGE_S_ENTERPRISE" }
                "\[INT\]"      { $SetEntPack  = $true; $SetEntPackOps = $SetEntPackOps -replace "INTUNE_O365"           }
                "\[MCO\]"      { $SetEntPack  = $true; $SetEntPackOps = $SetEntPackOps -replace "MCOSTANDARD"           }
                "\[O365\]"     { $SetEntPack  = $true; $SetEntPackOps = $SetEntPackOps -replace "OFFICESUBSCRIPTION"    }
                "\[RMS\]"      { $SetEntPack  = $true; $SetEntPackOps = $SetEntPackOps -replace "RMS_S_ENTERPRISE"      }
                "\[SHA\]"      { $SetEntPack  = $true; $SetEntPackOps = $SetEntPackOps -replace "SHAREPOINTENTERPRISE"  }
                "\[SWY\]"      { $SetEntPack  = $true; $SetEntPackOps = $SetEntPackOps -replace "SWAY"                  }
                "\[WAC\]"      { $SetEntPack  = $true; $SetEntPackOps = $SetEntPackOps -replace "SHAREPOINTWAC"         }
                "\[YAM\]"      { $SetEntPack  = $true; $SetEntPackOps = $SetEntPackOps -replace "YAMMER_ENTERPRISE"     }
            }

        }#not empty line

        #Delete empty EnterprisePack deletions
        if ($SetEntPack)
        {
            $SetEntPackOps = ( $SetEntPackOps -split "," | Where-Object { $_ -ne ""} ) -join ","
        }

        #Get current license
        $CUser = Get-MSOLUser -UserPrincipalName $userPrincipalName -ErrorAction Stop

        if ($CUser -ne $null)
        {
            $Result = ($CUser.Licenses).AccountSkuId -join ";"
            foreach ($Sku in ($CUser.Licenses).AccountSkuId)
            {
                switch ($Sku)
                {
                    'coffeeandtea:VISIOCLIENT'            { $HasVisio    = $true }
                    'coffeeandtea:PROJECTCLIENT'          { $HasProject  = $true }
                    'coffeeandtea:ENTERPRISEPACK'         { $HasEntPack  = $true }
                    'coffeeandtea:DESKLESSPACK'           { $HasDeskless = $true }
                    'coffeeandtea:POWER_BI_STANDARD'      { $HasPowerBI  = $true }
                    'coffeeandtea:EMS'                    { $HasEMS      = $true }
                    'coffeeandtea:RIGHTSMANAGEMENT_ADHOC' { $HasRMAdHoc  = $true }
                }
            }

            if ($HasEntPack)
            {
                $HasEntPackOps = ((($CUser.Licenses | 
                    Where-Object { $_.AccountSkuID -eq "coffeeandtea:ENTERPRISEPACK" }).ServiceStatus |
                    Where-Object { $_.ProvisioningStatus -eq "Disabled" }).ServicePlan.ServiceName |
                    Sort-Object) -join ","
            }

            if ($PSCmdlet.ShouldProcess($userPrincipalName, "Update Azure usageLocation and cloud license"))
            {
                if ( -not [string]::IsNullOrEmpty($CountryCode) )
                {
                    #Set usageLocation
                    if ($CountryCode -ne $CUser.UsageLocation -or $Force)
                    {
                        Write-Verbose "Set usageLocation to $CountryCode"
                        Set-MsolUser -UserPrincipalName $userPrincipalName -UsageLocation $CountryCode -ErrorAction Stop
                    }
                    else
                    {
                        Write-Verbose "UsageLocation $($CUser.UsageLocation) is already set to $CountryCode"
                    }
                }

                #Set or change EnterprisePack license if there is a difference between Set and Has
                if ($SetEntPack)
                {
                    if ($HasDeskless)
                    {
                        #Deskless needs to be removed before the license can be upgraded from Light to Full
                        Write-Verbose "Remove DesklessPack"
                        Set-MsolUserLicense -UserPrincipalName $userPrincipalName -RemoveLicenses "coffeeandtea:DESKLESSPACK" -ErrorAction Stop
                    }

                    Write-Verbose -Message "Has disable options: $HasEntPackOps"
                    Write-Verbose -Message "Set disable options: $SetEntPackOps"
                    if (-not $HasEntPack -or ($SetEntPackOps -ne $HasEntPackOps) -or $Force)
                    {
                        Write-Verbose -Message "Change detected"
                        if ($SetEntPackOps -ne $HasEntPackOps -or $Force)
                        {
                            if ($HasEntPack)
                            {
                                Write-Verbose "Remove EnterprisePack license"
                                Set-MsolUserLicense -UserPrincipalName $userPrincipalName -RemoveLicenses "coffeeandtea:ENTERPRISEPACK" -ErrorAction Stop
                            }
                            Write-Verbose "Set EnterprisePack with license options"
                            Set-MsolUserLicense -UserPrincipalName $userPrincipalName -AddLicenses "coffeeandtea:ENTERPRISEPACK" -LicenseOptions (New-MsolLicenseOptions -AccountSkuId "coffeeandtea:ENTERPRISEPACK" -DisabledPlans ($SetEntPackOps -split ",")) -ErrorAction Stop
                        }
                        else
                        {
                            if ($HasEntPack)
                            {
                                Set-MsolUserLicense -UserPrincipalName $userPrincipalName -RemoveLicenses "coffeeandtea:ENTERPRISEPACK" -ErrorAction Stop
                            }
                            Write-Verbose "Set EnterprisePack without license options to be removed"
                            Set-MsolUserLicense -UserPrincipalName $userPrincipalName -AddLicenses "coffeeandtea:ENTERPRISEPACK" -ErrorAction Stop
                        }
                    }
                }
                else
                {
                    if ($HasEntPack)
                    {
                        Write-Verbose "Remove EnterprisePack"
                        Set-MsolUserLicense -UserPrincipalName $userPrincipalName -RemoveLicenses "coffeeandtea:ENTERPRISEPACK" -ErrorAction Stop
                    }
                }

                if ($SetEMS)
                {
                    if (-not $HasEMS -or $Force)
                    {
                        Write-Verbose "Set EMS (Mobility Services)"
                        Set-MsolUserLicense -UserPrincipalName $userPrincipalName -AddLicenses "coffeeandtea:EMS" -ErrorAction Stop
                    }
                }
                else
                {
                    if ($HasEMS)
                    {
                        Write-Verbose "Remove EMS (Mobility Services)"
                        try
                        {
                            Set-MsolUserLicense -UserPrincipalName $userPrincipalName -RemoveLicenses "coffeeandtea:EMS" -ErrorAction Stop
                        }
                        catch
                        {
                            # Don't report the error becasue removing EMS has a huge delay of several hours
                            $Global:Error.RemoveAt(0)
                        }
                    }
                }

                if ($SetDeskless)
                {
                    if (-not $HasDeskless -or $Force)
                    {
                        Write-Verbose "Set DesklessPack"
                        Set-MsolUserLicense -UserPrincipalName $userPrincipalName -AddLicenses "coffeeandtea:DESKLESSPACK" -ErrorAction Stop
                    }
                }
                else
                {
                    if ($HasDeskless)
                    {
                        Write-Verbose "Remove DesklessPack"
                        Set-MsolUserLicense -UserPrincipalName $userPrincipalName -RemoveLicenses "coffeeandtea:DESKLESSPACK" -ErrorAction Stop
                    }
                }

                if ($SetVisio)
                {
                    if (-not $HasVisio -or $Force)
                    {
                        Write-Verbose "Set Visio"
                        Set-MsolUserLicense -UserPrincipalName $userPrincipalName -AddLicenses "coffeeandtea:VISIOCLIENT" -ErrorAction Stop
                    }
                }
                else
                {
                    if ($HasVisio)
                    {
                        Write-Verbose "Remove Visio"
                        Set-MsolUserLicense -UserPrincipalName $userPrincipalName -RemoveLicenses "coffeeandtea:VISIOCLIENT" -ErrorAction Stop
                    }
                }

                if ($SetProject)
                {
                    if (-not $HasProject -or $Force)
                    {
                        Write-Verbose "Set Project"
                        Set-MsolUserLicense -UserPrincipalName $userPrincipalName -AddLicenses "coffeeandtea:PROJECTCLIENT" -ErrorAction Stop
                    }
                }
                else
                {
                    if ($HasProject)
                    {
                        Write-Verbose "Remove Project"
                        Set-MsolUserLicense -UserPrincipalName $userPrincipalName -RemoveLicenses "coffeeandtea:PROJECTCLIENT" -ErrorAction Stop
                    }
                }

                if ($SetPowerBI)
                {
                    if (-not $HasPowerBI -or $Force)
                    {
                        Write-Verbose "Set Power BI"
                        Set-MsolUserLicense -UserPrincipalName $userPrincipalName -AddLicenses "coffeeandtea:POWER_BI_STANDARD" -ErrorAction Stop
                    }
                }
                else
                {
                    if ($HasPowerBI)
                    {
                        Write-Verbose "Remove Power BI"
                        Set-MsolUserLicense -UserPrincipalName $userPrincipalName -RemoveLicenses "coffeeandtea:POWER_BI_STANDARD" -ErrorAction Stop
                    }
                }

                if ($SetRMAdHoc)
                {
                    if (-not $HasRMAdHoc -or $Force)
                    {
                        Write-Verbose "Set Rights Management AdHoc"
                        Set-MsolUserLicense -UserPrincipalName $userPrincipalName -AddLicenses "coffeeandtea:RIGHTSMANAGEMENT_ADHOC" -ErrorAction Stop
                    }
                }
                else
                {
                    if ($HasRMAdHoc)
                    {
                        Write-Verbose "Remove Rights Management AdHoc"
                        Set-MsolUserLicense -UserPrincipalName $userPrincipalName -RemoveLicenses "coffeeandtea:RIGHTSMANAGEMENT_ADHOC" -ErrorAction Stop
                    }
                }
            }#WhatIf
        }#Get CUser
        else
        {
            Write-Error "Can't retrieve $userPrincipalName from the Azure cloud"
            $Result = "Error: Can't retrieve $userPrincipalName from the Azure cloud"
        }

        Write-Output $Result
    }#Process
    End
    {
    }
}

#endregion

#region Split-AzureLicenseTags
<#
.Synopsis
   Split Tags string to array.
.DESCRIPTION
   The cmdlet split the tags string into an array.
.EXAMPLE
   Split-AzureLicenseTags -Tags '[EMS][EXC][INT][O365][Project][RMS][SHA][SWY][Visio][WAC][YAM]'

   This will return an array with the tags:
      [EMS]
      [EXC]
      [INT]
      [O365]
      [Project]
      [RMS]
      [SHA]
      [SWY]
      [Visio]
      [WAC]
      [YAM]

.EXAMPLE
   '[EMS][EXC]' | Split-AzureLicenseTags

   This will return an array with the tags:
      [EMS]
      [EXC]
.INPUTS
   [string]
.OUTPUTS
   [string[]]
.NOTES
    --- Version history:
    Version 1.01 2016-09-01, Kees Hiemstra)
    - Change the Tags paramater to not manatory.
    Version 1.00 2016-08-03, Kees Hiemstra)
    - Initial version.

.COMPONENT
   The component this cmdlet belongs to the SOEAzure module.

.ROLE
   The role this cmdlet belongs to Azure license management.

#>
function Split-AzureLicenseTags
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$true,
                  ConfirmImpact='Medium')]
    [OutputType([String[]])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$true, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [string]
        $Tags
    )

    Begin
    {
    }
    Process
    {
        if ( -not [string]::IsNullOrWhiteSpace($Tags) )
        {
            Write-Output ([array]$Tags.Replace('][', '],[') -split ",")
        }
        else
        {
            Write-Output (@())
        }
    }
    End
    {
    }
}
#endregion

#region Remove-AzureLicenseTags
<#
.Synopsis
   Remove Azure license tags from input string.
.DESCRIPTION
   The cmdlet will remove the given array of license tags from the input string and returns the result.
.EXAMPLE
   Remove-AzureLicenseTags -Tags '[EMS][EXC][INT][O365][RMS][SHA][SWY][Visio][WAC][VPN]' -RemoveTags '[Visio]'

   This will return '[EMS][EXC][INT][O365][RMS][SHA][SWY][WAC][VPN]'.
.EXAMPLE
   Remove-AzureLicenseTags -Tags '[EMS][EXC][INT][O365][RMS][SHA][SWY][Visio][WAC][VPN]' -RemoveTags @('[Visio]', '[EMS]')

   This will return '[EXC][INT][O365][RMS][SHA][SWY][WAC][VPN]'.
.EXAMPLE
   '[EMS][EXC][INT][O365][RMS][SHA][SWY][Visio][WAC][VPN]' | Remove-AzureLicenseTags -RemoveTags '[Visio]'

   This will return '[EMS][EXC][INT][O365][RMS][SHA][SWY][WAC][VPN]'.
.INPUTS
   
.OUTPUTS
   Resulting license string
.NOTES
    --- Version history:
    Version 1.02 (2016-09-15, Kees Hiemstra)
    - Bug fix: The function was case sensitive.
    Version 1.01 (2016-09-01, Kees Hiemstra)
    - Change the Tags paramater to not manatory.
    Version 1.00 (2016-08-03, Kees Hiemstra)
    - Initial version.

.COMPONENT
   The component this cmdlet belongs to the SOEAzure module.

.ROLE
   The role this cmdlet belongs to Azure license management.

#>
function Remove-AzureLicenseTags
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$true,
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([String])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [string]
        $Tags,

        # Param2 help description
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=1,
                   ParameterSetName='Parameter Set 1')]
        [string[]]
        $RemoveTags
    )
    Begin
    {
    }
    Process
    {
        $Tags = $Tags.ToUpper()
        foreach ( $Tag in $RemoveTags )
        {
            $Tags = $Tags.Replace($Tag.ToUpper(), '')
        }#foreach
        Write-Output $Tags
    }
    End
    {
    }
}
#endregion

#region Test-AzureLicenseTags

<#
.SYNOPSIS
   Validates the Azure license tags.
.DESCRIPTION
   Azure license tags are stored in ExtensionAttribute11 (HR profile) and ExtensionAttribute2 (Extra Azure licenses).
   This cmdlet will validate the combination of these two and will remove any EnterprisePack options if the DesklessPack is selected [Light].
   
   The output object will contain the (corrected) total set of tags that can be used for the Set-AzureLicense cmdlet.

   The validation will add the mandatory [INT] tag if it is missing and one or more EnterprisePack options are selected.

.EXAMPLE
    Test-AzureLicenseTags -EA2 '![EMS][EXC][O365][RMS][SHA][SWY][WAC][YAM][VPN][Visio][Project]' -EA11 '[EMS][EXC][O365][RMS][SHA][SWY][WAC][YAM][VPN][VPN]'

    The output object will contain:
        ResultTags       : ![EMS][EXC][INT][O365][Project][RMS][SHA][SWY][Visio][VPN][WAC][YAM]
        EA2OriginalTags  : ![EMS][EXC][O365][RMS][SHA][SWY][WAC][YAM][VPN][Visio][Project]
        EA2LicenseTags   : ![INT][Project][Visio]
        HasDenyLicenses  : True
        EA11LicenseTags  : [EMS][EXC][O365][RMS][SHA][SWY][VPN][WAC][YAM]
        EA11OriginalTags : [EMS][EXC][O365][RMS][SHA][SWY][WAC][YAM][VPN][VPN]
        AzureLicenseTags : 
        GroupLicenseTags : 
        HasEA2Changed    : True
        HasEA11Changed   : True

        In this example the combination of ExtensionAttribute2 and ExtensionAttribute11 do have a lot of the same Tags, these are removed. The deny character (!) takes presidence of all Tags, so in the end all licenses will be withdrawn.

.EXAMPLE
    Test-AzureLicenseTags -EA2 '[VPN][EMS][EXC][INT][O365][RMS][SHA][SWY][WAC][YAM][Project][Visio]' -EA11 '[VPN][EMS][Light]'

    The output object will contain:
        ResultTags       : [EMS][Light][VPN]
        EA2OriginalTags  : [VPN][EMS][EXC][INT][O365][RMS][SHA][SWY][WAC][YAM][Project][Visio]
        EA2LicenseTags   : 
        HasDenyLicenses  : False
        EA11LicenseTags  : [EMS][Light][VPN]
        EA11OriginalTags : [VPN][EMS][Light]
        AzureLicenseTags : [EMS][Light]
        GroupLicenseTags : [ADFS][EMS][VPN]
        HasEA2Changed    : True
        HasEA11Changed   : True

        In this example the [Light] Tag in the ExtensionAttribute11 Tags overrules the Enterprise tags in the ExtensionAttribute2 and these are removed.

.EXAMPLE
    Test-AzureLicenseTags -SAMAccountName King.Arthur

        ResultTags       : [EMS][EXC][INT][O365][Project][RMA][RMS][SHA][SWY][WAC][YAM]
        EA2OriginalTags  : [PROJECT][EMS][EXC][INT][O365][RMA][RMS][SHA][SWY][WAC][YAM][project]
        EA2LicenseTags   : [EMS][EXC][INT][O365][Project][RMA][RMS][SHA][SWY][WAC][YAM]
        HasDenyLicenses  : False
        EA11LicenseTags  : 
        EA11OriginalTags : 
        AzureLicenseTags : [EMS][EXC][INT][O365][Project][RMA][RMS][SHA][SWY][WAC][YAM]
        GroupLicenseTags : [ADFS][EMS]
        HasEA2Changed    : True
        HasEA11Changed   : False

    In this example the user King.Arthur has the project tag twice in the orgiginal ExtensionAttribute2 tag. One of these tags is removed.
.OUTPUTS
    [SOEAzure.TestLicenseTags.100]

    - ResultTags       = The result of the combination (if applicable) and corrections.
    - EA2OriginalTags  = The original ExtensionAttribute2.
    - EA2LicenseTags   = The corrected ExtensionAttribute2.
    - HasDenyLicenses  = If the original ExtensionAttribute2 contains deny license mark.
    - EA11LicenseTags  = The original ExtensionAttribute11.
    - EA11OriginalTags = The corrected ExtensionAttribute11.
    - AzureLicenseTags = All Azure related tags.
    - GroupLicenseTags = All security group related tags.
    - HasEA2Changed    = Has ExtensionAttribute2 changed?
    - HasEA11Changed   = Has ExtensionAttribute11 changed?

.NOTES
    --- Version history:
    Version 1.02 (2016-08-23, Kees Hiemstra)
    - Bug fix: Avoid error when EA11 is empty.
    Version 1.01 (2016-08-03, Kees Hiemstra)
    - Replaced SplitTags with Split-AzureLicenseTags.
    - Replaced RemoveTags with Remove-AzureLicenseTags.
    Version 1.00 (2016-07-05, Kees Hiemstra)
    - Initial version.

.COMPONENT
   The component this cmdlet belongs to the SOEAzure module.

.ROLE
   The role this cmdlet belongs to Azure license management.

.LINK
    Set-AzureLicense

.LINK
    ConvertFrom-AzureLicense

#>
function Test-AzureLicenseTags
{
    [CmdletBinding(DefaultParameterSetName='Combine EA2 and EA11',
                  SupportsShouldProcess=$false, 
                  PositionalBinding=$false,
                  ConfirmImpact='Low')]
    [OutputType([String])]
    Param
    (
        # The value of ExtensionAttribute2.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Combine EA2 and EA11')]
        [Alias("ExtensionAttribute2")] 
        [string]
        $EA2,

        # The value of ExtensionAttribute11. Use -join '' if the value comes directly from AD.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position=1,
                   ParameterSetName='Combine EA2 and EA11')]
        [Alias("ExtensionAttribute11")] 
        [string]
        $EA11,

        # The total set of licences. The result values regarding ExtensionAttribute2 and ExtensionAttribut11 will not be updated.
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Total set of licenses')]
        [Alias("AllLicenseTags")] 
        [string]
        $LicenseTags,

        # SAMAccountName of the user, the ExtensionAttribute2 and ExtensionAttribute11 will be read from AD.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Get from AD')]
        [Alias("Name")]
        [string]
        $SAMAccountName,

        # The original value of ExtensionAttribute2.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false)]
        [string]
        $OriginalEA2
    )

    Begin
    {
        $ReturnProps = [ordered]@{ResultTags       = ''
                                  EA2OriginalTags  = ''
                                  EA2LicenseTags   = ''
                                  HasDenyLicenses  = $false
                                  EA11LicenseTags  = ''
                                  EA11OriginalTags = ''
                                  AzureLicenseTags = ''
                                  GroupLicenseTags = ''
                                  HasEA2Changed    = $false
                                  HasEA11Changed   = $false
                        }
    }#Begin
    Process
    {
        Write-Verbose "Paramater set: $($PsCmdlet.ParameterSetName)"

        $OnlyLicenseTags = $false
        $Return = New-Object -TypeName PSObject -Property $ReturnProps
        $Return.PSObject.TypeNames.Insert(0, 'SOEAzure.TestLicenseTags.100')

        if ( $PSBoundParameters.ContainsKey('OriginalEA2') )
        {
            $Return.EA2OriginalTags = $OriginalEA2
        }

        if ( $PSBoundParameters.ContainsKey('SAMAccountName') )
        {
            Write-Verbose -Message "Get EA2 and EA11 from AD for user $($SAMAccountName)"

            $ADUser = Get-ADUser -Identity $SAMAccountName -Properties ExtensionAttribute2, ExtensionAttribute11
            $EA2 = $ADUser.ExtensionAttribute2 -join ''
            $EA11 = $ADUser.ExtensionAttribute11 -join ''
        }#Get-ADUser

        if ( $PSBoundParameters.ContainsKey('LicenseTags') )
        {
            Write-Verbose -Message "Process parameter -LicenseTags '$($LicenseTags)'"

            $OnlyLicenseTags = $true

            $Return.HasDenyLicenses = ($LicenseTags -match '^\!')
            $LicenseTags = CheckAndSortTags -Tags $LicenseTags -CheckTags $InAllTags

            # Remove EnterpriseTags, Visio and Project if the DesklessPack is used
            if ( $LicenseTags -match '\[Light\]' )
            {
                Write-Verbose "Remove EnterprisePack, Visio and Project"

                $LicenseTags = RemoveEnterpriseTags -Tags $LicenseTags
            }

            # Add [INT] if it is not present whilst other EnterprisePack are present
            if ( (HasEnterpriseTags -Tags  $LicenseTags) -and ($LicenseTags -notmatch '\[INT\]') )
            {
                $LicenseTags += '[INT]'
                $LicenseTags = CheckAndSortTags -Tags $LicenseTags -CheckTags $InAllTags

                $EA2 += '[INT]'
                $EA2 = CheckAndSortTags -Tags $EA2 -CheckTags $InAllTags
            }

            # Re-add the ! if it was present in the original input
            if ( $Return.HasDenyLicenses )
            {
                if ( $LicenseTags -notmatch '^\!' )
                {
                    $LicenseTags = "!$($LicenseTags)"
                }
            }

            $Return.AzureLicenseTags = $LicenseTags
        }#LiceseTags
        else
        {
            Write-Verbose -Message "Process parameter -EA2 '$($EA2)' -EA11 '$($EA11)'"

            $Return.HasDenyLicenses = ($EA2 -match '^\!')
            if ( -not $PSBoundParameters.ContainsKey('OriginalEA2') )
            {
                $Return.EA2OriginalTags = $EA2
            }
            $Return.EA11OriginalTags = $EA11

            $EA2 = CheckAndSortTags -Tags $EA2 -CheckTags $InAllTags
            $EA11 = CheckAndSortTags -Tags $EA11 -CheckTags $InAllTags

            # Remove EnterpriseTags, Visio and Project if the DesklessPack is used
            if ( $EA2 -match '\[Light\]' -or $EA11 -match '\[Light\]' )
            {
                Write-Verbose "Remove EnterprisePack, Visio and Project"

                $EA2 = RemoveEnterpriseTags -Tags $EA2
                $EA11 = RemoveEnterpriseTags -Tags $EA11
            }

            if ( -not [string]::IsNullOrEmpty($EA11) )
            {
                $EA2 = Remove-AzureLicenseTags -Tags $EA2 -RemoveTags (Split-AzureLicenseTags -Tags $EA11)
            }

            $LicenseTags = "$($EA2)$($EA11)"

            # Add [INT] if it is not present whilst other EnterprisePack are present
            if ( (HasEnterpriseTags -Tags $LicenseTags) -and ($LicenseTags -notmatch '\[INT\]') )
            {
                Write-Verbose -Message "Add [INT]"

                $LicenseTags += '[INT]'
                $LicenseTags = CheckAndSortTags -Tags $LicenseTags -CheckTags $InAllTags

                $EA2 += '[INT]'
                $EA2 = CheckAndSortTags -Tags $EA2 -CheckTags $InAllTags
            }

            # Re-add the ! if it was present in the original ExtensionAttribute2
            if ( $Return.HasDenyLicenses )
            {
                if ( $LicenseTags -notmatch '^\!' )
                {
                    $LicenseTags = "!$($LicenseTags)"
                }

                if ( $EA2 -notmatch '^\!' )
                {
                    $EA2 = "!$($EA2)"
                }
            }

            $Return.ResultTags = $LicenseTags
            $Return.EA2LicenseTags = $EA2
            $Return.EA11LicenseTags = $EA11
            $Return.HasEA2Changed = ($EA2 -ne $Return.EA2OriginalTags)
            $Return.HasEA11Changed = ($EA11 -ne $Return.EA11OriginalTags)

        }#EA2 & EA11

        if ( -not $Return.HasDenyLicenses )
        {
            $Return.AzureLicenseTags = (CheckAndSortTags -Tags $LicenseTags -CheckTags $InAzureTags )

            if ( (HasADFSTags -Tags $LicenseTags) )
            {
                $LicenseTags += '[ADFS]'
            }
            $Return.GroupLicenseTags = (CheckAndSortTags -Tags $LicenseTags -CheckTags $InGroupTags )
        }

        Write-Output $Return
    }#Process
    End
    {
    }
}#Test-AzureLicenses
#endregion

#region Sort-AzureLicenseTags

<#
.Synopsis
   Sort Tags on alphabetical order.
.DESCRIPTION
   This cmdlet sorts the Tags from the input in alphabetical order. This is used in SyncAD2Azure script to avoid false positive changes in ExtensionAttribute2 where only the order of the Tags has changed.
.EXAMPLE
    Sort-AzureLicenseTags -Tags '[PROJECT][VPN][EXC][INT][O365][RMA][RMS][SHA][SWY][WAC][YAM][EMS][project]'

    The result: [EMS][EXC][INT][O365][PROJECT][project][RMA][RMS][SHA][SWY][VPN][WAC][YAM]

.EXAMPLE
    '[PROJECT][VPN][EXC][INT][O365][RMA][RMS][SHA][SWY][WAC][YAM][EMS][project]' | Sort-AzureLicenseTags

    The result: [EMS][EXC][INT][O365][PROJECT][project][RMA][RMS][SHA][SWY][VPN][WAC][YAM]

.INPUTS
   [string]

.OUTPUTS
   [string]

.NOTES
    This cmdlet produces the following warning when the module is loaded: "WARNING: The names of some imported commands from the module 'SOEAzure' include unapproved verbs that might make them less discoverable. To find the commands with unapproved verbs, run the Import-Module command again with the Verbose parameter. For a list of approved verbs, type Get-Verb."

    There is no equivalent verb that matched Sorting and Sort-Object is allowed, so I decided not to change the verb of this cmdlet.

    --- Version history:
    Version 1.01 2016-08-03, Kees Hiemstra)
    - Replaced SplitTags with Split-AzureLicenseTags.
    Version 1.00 2016-07-07, Kees Hiemstra)
    - Initial version.

.COMPONENT
   The component this cmdlet belongs to the SOEAzure module.

.ROLE
   The role this cmdlet belongs to Azure license management.

#>
function Sort-AzureLicenseTags
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$false, 
                  PositionalBinding=$false,
                  ConfirmImpact='low')]
    [OutputType([String])]
    Param
    (
        # Tags to be sorted on alphabet.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        $Tags
    )

    Process
    {
        if ( -not [string]::IsNullOrWhiteSpace($Tags) )
        {
            $Tags = $Tags.Replace(' ', '')
            $DenyLicenses = $Tags -like '!*'
            $Tags = $Tags.Replace('!', '')

            $Tags = (Split-AzureLicenseTags -Tags $Tags | Sort-Object ) -join ''

            if ( $DenyLicenses )
            {
                $Tags = "!$($Tags)"
                Write-Debug "Deny licenses added: $($Tags)"
            }
        }
        Write-Output $Tags
    }
}

#endregion

#region Get-MsolUserLicenseParameterHashTable

<#
.SYNOPSIS
    Get a hash table for the Set-MsolUserLicense parameters -AddLicense and -LicenseOptions to be used in splatting.
.DESCRIPTION
    RfC 1257953 - Remove unwanted O365 license options, JDE wants to avoid options that can not be configured.
    The RfC resulted in the process 'MonitorSku' and this will provide a file license with the options that never should
    be set (if applicable).
    This cmdlet will use this file to complete the 'forbidden' options and will prevent duplicate disabled plans.
.EXAMPLE
    Get-MsolUserLicenseParameterHashTable -Licence 'camelot:ENTERPRISEPACK' -DisabledPlans 'MCOSTANDARD,YAMMER'

    ---
    @('AddLicense'='camelot:ENTERPRISEPACK'=[LicenseOption]Object)
    The object will contain disabled plans for MCOSTANDARD, YAMMER and FLOW_O365_P2 (because this one has been added)
.EXAMPLE

.INPUTS

.OUTPUTS
    Hash table for splatting
.NOTES
    === Version history
    Version 1.00 (2017-04-10, Kees Hiemstra)
    - Initial version.
.COMPONENT
    Azure license manangement.
.ROLE

.FUNCTIONALITY

#>
function Get-MsolUserLicenseParameterHashTable
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1',
                  SupportsShouldProcess=$true,
                  PositionalBinding=$true,
                  ConfirmImpact='Medium')]

    Param
    (
        # AccountSkuId to be set (e.g. camelot:ENTERPRISEPACK)
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$false,
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [string]
        $AccountSkuID,

        # Disabled plans that have been determined up front
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$false,
                   Position=1,
                   ParameterSetName='Parameter Set 1')]
        [string]
        $DisabledPlans,

        # Alternative path to look for the .OptOut files
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$false,
                   ParameterSetName='Parameter Set 1')]
        [string]
        $OptOutPath
    )

    Begin
    {
    }
    Process
    {
        
    }
    End
    {
    }
}

#endregion
