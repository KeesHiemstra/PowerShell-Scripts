#region Get-SOEComputerBIOS

<#
.SYNOPSIS
    Gets the BIOS data from the local computer or a remote computer.
.DESCRIPTION
   The Get-SOEComputerBIOS cmdlet gets a select part of BIOS data from the local or remote computer through WMI as SOE.ComputerBIOS.100 object.
   
   The BIOS data contains the following attributes:
   •	ComputerName
   •	SerialNo
   •	Vendor
   •	Name
   •	ReleaseDate
   •	SMBIOSPresent
   •	SMBIOSBIOSVersion
   •	SMBIOSMajorVersion
   •	SMBIOSMinorVersion
   •	Version
   •	BiosCharacteristics
   •	DTCollection

.EXAMPLE
    Get BIOS data from the local computer

    PS C:\>Get-SOEComputerModel -ComputerName .

        ComputerName        : KING001
        SerialNo            : CZC552KING
        Vendor              : Hewlett-Packard
        Name                : Default System BIOS
        ReleaseDate         : 2015-10-29 00:00:00
        SMBIOSPresent       : True
        SMBIOSBIOSVersion   : 786G3 v03.15
        SMBIOSMajorVersion  : 2
        SMBIOSMinorVersion  : 6
        Version             : HPQOEM - 20151029
        BiosCharacteristics : 7;9;11;12;15;16;19;21;24;26;27;28;29;32;33;36;37;40;41;42;48;49;56;57;58;59;64;65;66;67;68;69;70;71;72;73;74;75;76;77;78;79
        DTCollection        : 2015-12-21 12:01:17

    This command gets the BIOS data from the local computer.
.INPUTS
    <string>
.OUTPUTS
    This cmdlet returns a named System.Management.Automation.PSCustomObject object called SOE.ComputerBIOS.100.

    Table information
    Saving the output of this cmdlet with the Save-SOEObject cmdlet in an SQL table requires the following table structure.
    CREATE TABLE dbo.ComputerBIOS(
        [ID] [int] IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [ComputerName] [varchar](25) NOT NULL,
        [SerialNo] [varchar](75) NOT NULL,
        [Vendor] [varchar](50) NOT NULL,
        [Name] [varchar](50) NOT NULL,
        [ReleaseDate] [date] NOT NULL,
        [SMBIOSPresent] [bit] NULL,
        [SMBIOSBIOSVersion] [varchar](50) NULL,
        [SMBIOSMajorVersion] [smallint] NULL,
        [SMBIOSMinorVersion] [smallint] NULL,
        [Version] [varchar](50) NOT NULL,
        [BIOSCharacteristics] [varchar](255) NOT NULL,
        [DTCollection] [datetime] NOT NULL,
        [DTCreation] [datetime] NOT NULL CONSTRAINT [DF_ComputerBIOS_DTCreation] DEFAULT (GETDATE()),
        [DTMutation] [datetime] NULL,
        [DTDeletion] [datetime] NULL
        )
    
.NOTES
    --- Version history:
    Version 1.10 (2016-10-31, Kees Hiemstra)
    - Added parameter Credential.
    Version 1.01 (2015-12-22, Kees Hiemstra)
    - Bug fix: Increased the size of the [SerialNo] field from 50 to 75 in the database table.
    Version 1.00 (2015-12-21, Kees Hiemstra)
    - Initial version.

.LINK
    Get-SOEComputerModel

.LINK
    Get-SOEComputerOS

.LINK
	Get-SOEComputerBSoDEvents

.LINK
	Get-SOEComputerNICDetails

.LINK
	Get-SOEComputerSignedDrivers

.LINK
	Get-SOEComputerRunKeys

.LINK
	Get-SOEComputerPatches

.LINK
    Get-SOEComputerUserFirewallRules

.LINK
    Get-SOEComputerPolicyFirewallRules

.LINK
    Get-SOEComputerWifiProfile

.LINK
    Get-SOEComputerWifiSignature

.LINK
    Save-SOEObject

#>
function Get-SOEComputerBIOS
{
    [CmdletBinding(SupportsShouldProcess=$true,
                   PositionalBinding=$true,
                   HelpUri='http://www.xs4all.nl/~chi/ps/oh1552101.html',
                   ConfirmImpact='Low')]
    Param
    (
        #Specifies the computer for which this cmdlet gets the BIOS data.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0)]
        [Alias('Name')]
        [string]
        $ComputerName,

        #Specifies the user account credentials to use to perform this task. The default credentials are the credentials of the currently logged on user unless the cmdlet is run from an Active Directory PowerShell provider drive. If the cmdlet is run from such a provider drive, the account associated with the drive is the default.
        #
        #To specify this parameter, you can type a user name, such as "User1" or "Domain01\User01" or you can specify a PSCredential object. If you specify a user name for this parameter, the cmdlet prompts for a password.
        #
        #You can also create a PSCredential object by using a script or by using the Get-Credential cmdlet. You can then set the Credential parameter to the PSCredential object 
        #The following example shows how to create credentials.
        #$AdminCredentials = Get-Credential "Domain01\User01"
        #The following shows how to set the Credential parameter to these credentials.
        #-Credential $AdminCredentials
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
        if ((Test-Connection -ComputerName $ComputerName -Quiet -Count 1) -or $ComputerName -eq '.')
        {
            Write-Verbose "Querying $ComputerName for basic SOE Computer BIOS information"
            $DTCollection = (Get-Date).ToUniversalTime()

            $Data = Get-WmiObject -ComputerName $ComputerName -Class Win32_BIOS -ErrorAction SilentlyContinue @Param
            if ($Data -ne $null)
            {
                foreach ($Item in $Data)
                {
                    $Properties = [ordered]@{'ComputerName'        = $Item.__SERVER;
                                             'SerialNo'            = $Item.SerialNumber;
                                             'Vendor'              = $Item.Manufacturer;
                                             'Name'                = $Item.Name;
                                             'ReleaseDate'         = [datetime][management.managementDateTimeConverter]::ToDateTime($Item.ReleaseDate).ToUniversalTime();
                                             'SMBIOSPresent'       = [bool]$Item.SMBIOSPresent;
                                             'SMBIOSBIOSVersion'   = $Item.SMBIOSBIOSVersion;
                                             'SMBIOSMajorVersion'  = $Item.SMBIOSMajorVersion;
                                             'SMBIOSMinorVersion'  = $Item.SMBIOSMinorVersion;
                                             'Version'             = $Item.Version;
                                             'BIOSCharacteristics' = $Item.BiosCharacteristics -join ';';
                                             'DTCollection'        = $DTCollection }
                    $Return = New-Object -TypeName PSObject -Property $Properties
                    $Return.PSObject.TypeNames.Insert(0, 'SOE.ComputerBIOS.100')
                    Write-Output $Return
                }#foreach
            }
            else
            {
                Write-Verbose "Can't query $ComputerName for basic SOE Computer BIOS information"
            }#Access?
        }
        else
        {
            Write-Verbose "$ComputerName is not online to query basic SOE Computer BIOS information"
        }#Online?
    }
}

#endregion

#region Get-SOEComputerBSoDEvents

<#
.SYNOPSIS
    Gets the Blue Screen of Death data from the local computer or a remote computer.
.DESCRIPTION
    The SOE.ComputerBSoDEvents cmdlet gets a select part of the Blue Screen of Death data from the local or remote computer through WMI as SOE.ComputerBSoDEvents.200 object.
    
    The BSoDEvents data contains the following attributes:
    •	ComputerName
    •	Strings
    •	StopCode
    •	StopCodeInt
    •	Param1
    •	Param1Int
    •	Param2
    •	Param3
    •	Param4
    •	DTEvent
    •	DTCollection
    
    StopCode is the 32bit integer that appears as the first number of the BSoD message.
    Together with the 64bit integers Param1 to Param4 they tend to give already more information in which Param1 is the most important one.
    Therefore, the StopCode and Param1 are also translated to their int32 and int64 equivalent.
    For more information on the codes themselves can be found on the page https://msdn.microsoft.com/en-us/library/windows/hardware/hh994433(v=vs.85).aspx

.EXAMPLE
    PS C:\> Get-SOEComputerBSoDEvents -ComputerName .

    ComputerName : KING001
    Strings      : 0x000000c2 (0x0000000000000007, 0x000000000000109b, 0x00000000040e0005,
                    0xfffffa8013d98600);C:\Windows\MEMORY.DMP;121815-10561-01
    StopCode     : 0x000000c2
    StopCodeInt  : 194
    Param1       : 0x0000000000000007
    Param1Int    : 7
    Param2       : 0x000000000000109b
    Param3       : 0x00000000040e0005
    Param4       : 0xfffffa8013d98600
    DTEvent      : 2015-12-17 23:14:51
    DTCollection : 2015-12-20 16:30:43

    This command gets the Blue Screen of Death information from the local computer.

.INPUTS
    <string>
.OUTPUTS
    SOE.ComputerBSoDEvents.200
    
    This cmdlet returns a named System.Management.Automation.PSCustomObject object called SOE.ComputerBSoDEvents.200.
    
    Table information
    Saving the output of this cmdlet with the Save-SOEObject cmdlet in an SQL table requires the following table structure.
    
    CREATE TABLE dbo.ComputerBSoDEvents(
        [ID] [int] IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [ComputerName] [varchar](25) NOT NULL,
        [Strings] [varchar](255) NOT NULL,
        [StopCode] [char](10) NOT NULL,
        [StopCodeInt] [int] NOT NULL,
        [Param1] [char](18) NOT NULL,
        [Param2] [char](18) NOT NULL,
        [Param3] [char](18) NOT NULL,
        [Param4] [char](18) NOT NULL,
        [DTEvent] [datetime] NOT NULL,
        [DTCollection] [datetime] NOT NULL,
        [DTCreation] [datetime] NOT NULL CONSTRAINT [DF_ComputerBSoDEvents_DTCreation] DEFAULT (GETDATE()),
        [DTMutation] [datetime] NULL,
        [DTDeletion] [datetime] NULL
        )
    
.NOTES
    --- Version history
    Version 2.10 (2016-10-31, Kees Hiemstra)
    - Added parameter Credential.
    Version 2.01 (2015-12-20, Kees Hiemstra)
    - Added help.
    Version 2.00 (2015-12-16, Kees Hiemstra)
    - Updated the output object version to 200.
    - Removed the [Message] because it is not giving more information than [Strings].
    - Added [StopCode] and [StopCodeInt].
    - Added [Param1], [Param1Int], [Param2], [Param3], [Param4].
    Version 1.00 (2015-12-07, Kees Hiemstra)
    - Initial version.

.LINK
    Get-SOEComputerModel

.LINK
    Get-SOEComputerOS

.LINK
	Get-SOEComputerNICDetails

.LINK
	Get-SOEComputerSignedDrivers

.LINK
	Get-SOEComputerRunKeys

.LINK
	Get-SOEComputerPatches

.LINK
    Get-SOEComputerBIOS

.LINK
    Get-SOEComputerUserFirewallRules

.LINK
    Get-SOEComputerPolicyFirewallRules

.LINK
    Get-SOEComputerWifiProfile

.LINK
    Get-SOEComputerWifiSignature

.LINK
    Save-SOEObject

#>
function Get-SOEComputerBSoDEvents
{
    [CmdletBinding(SupportsShouldProcess=$true, 
                   PositionalBinding=$true,
                   HelpUri='http://www.xs4all.nl/~chi/ps/oh1550101.html',
                   ConfirmImpact='Low')]
    Param
    (
        #Specifies the computer for which this cmdlet gets the Blue Screen of Death data.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false,
                   Position=0)]
        [Alias('Name')]
        [string]
        $ComputerName,

        #Specifies the user account credentials to use to perform this task. The default credentials are the credentials of the currently logged on user unless the cmdlet is run from an Active Directory PowerShell provider drive. If the cmdlet is run from such a provider drive, the account associated with the drive is the default.
        #
        #To specify this parameter, you can type a user name, such as "User1" or "Domain01\User01" or you can specify a PSCredential object. If you specify a user name for this parameter, the cmdlet prompts for a password.
        #
        #You can also create a PSCredential object by using a script or by using the Get-Credential cmdlet. You can then set the Credential parameter to the PSCredential object 
        #The following example shows how to create credentials.
        #$AdminCredentials = Get-Credential "Domain01\User01"
        #The following shows how to set the Credential parameter to these credentials.
        #-Credential $AdminCredentials
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
        if ((Test-Connection -ComputerName $ComputerName -Quiet -Count 1) -or $ComputerName -eq '.')
        {
            Write-Verbose "Querying $ComputerName for BSoD Events information"
            $DTCollection = (Get-Date).ToUniversalTime()

            $Data = Get-WmiObject -Class Win32_ReliabilityRecords -Filter "EventIdentifier='1001'" -ComputerName $ComputerName -ErrorAction SilentlyContinue @Param
            if ($Data -ne $null)
            {
                foreach ($Item in $Data)
                {
                    $InsertionStrings = $Item.InsertionStrings -join ";";
                    $Strings = $InsertionStrings -split ' '

                    $Properties = [ordered]@{'ComputerName'       = $Item.__SERVER;
                                             'Strings'            = $InsertionStrings;
                                             'StopCode'           = $Strings[0];
                                             'StopCodeInt'        = [int]$Strings[0];
                                             'Param1'             = $Strings[1] -replace "^\(" -replace "\).*" -replace ",";
                                             'Param1Int'          = [uint64]($Strings[1] -replace "^\(" -replace "\).*" -replace ",");
                                             'Param2'             = $Strings[2] -replace "^\(" -replace "\).*" -replace ",";
                                             'Param3'             = $Strings[3] -replace "^\(" -replace "\).*" -replace ",";
                                             'Param4'             = $Strings[4] -replace "^\(" -replace "\).*" -replace ",";
                                             'DTEvent'            = [datetime][management.managementDateTimeConverter]::ToDateTime($Item.TimeGenerated).ToUniversalTime();
                                             'DTCollection'       = $DTCollection }
                    $Return = New-Object -TypeName PSObject -Property $Properties
                    $Return.PSObject.TypeNames.Insert(0, 'SOE.ComputerBSoDEvents.200')

                    Write-Output $Return
                }#foreach
            }
            else
            {
                Write-Verbose "Can't query $ComputerName for BSoD Events information"
            }#Access?
        }
        else
        {
            Write-Verbose "$ComputerName is not online to query BSoD Events information"
        }#Online?
    }
}

#endregion

#region Get-SOEComputerModel

<#
.SYNOPSIS
    Gets the model data from the local computer or a remote computer.

.DESCRIPTION
    The Get-SOEComputerModel cmdlet gets a selected part of the model data from the local or remote computer through WMI as a SOE.ComputerModel.100 object.

    The ComputerModel data contains the following attributes:
    - ComputerName
    - Vendor
    - Model
    - PhysicalMemorySize
    - ConsoleUser
    - DTCollection

.EXAMPLE
    PS C:\>Get-SOEComputerModel -ComputerName .

        ComputerName       : KING001
        Vendor             : Hewlett-Packard
        Model              : HP EliteBook 840 G2
        PhysicalMemorySize : 8458928128
        ConsoleUser        : Camelot\KingArthur
        DTCollection       : 2015-12-09 07:51:05

    This command gets the model information from the local computer.

.INPUTS
    [string]
.OUTPUTS
    [SOE.ComputerModel.100]

    This cmdlet returns a named System.Management.Automation.PSCustomObject object called SOE.ComputerModel.100.

    Table information
    Saving the output of this cmdlet with the Save-SOEObject cmdlet in an SQL table requires the following table structure.

    CREATE TABLE dbo.ComputerModel(
	    [ID] [int] IDENTITY(1,1) NOT NULL PRIMARY KEY,
	    [ComputerName] [varchar](25) NOT NULL,
	    [Vendor] [varchar](50) NOT NULL,
	    [Model] [varchar](75) NOT NULL,
	    [PhysicalMemorySize] [bigint] NOT NULL,
	    [ConsoleUser] [varchar](75) NULL,
	    [DTCollection] [datetime] NOT NULL,
	    [DTCreation] [datetime] NOT NULL CONSTRAINT [DF_ComputerModel_DTCreation] DEFAULT (GETDATE()),
	    [DTMutation] [datetime] NULL,
	    [DTDeletion] [datetime] NULL
        )

.NOTES
    --- Version history:
    Version 1.10 (2016-10-31, Kees Hiemstra)
    - Added parameter Credential.
    Version 1.02 (2015-12-20, Kees Hiemstra)
    - Adding extended help.
    Version 1.01 (2015-12-09, Kees Hiemstra)
    - Adding help.
    Version 1.00 (2015-12-06, Kees Hiemstra)
    - Initial version.

.LINK
    Get-SOEComputerOS

.LINK
	Get-SOEComputerBSoDEvents

.LINK
	Get-SOEComputerNICDetails

.LINK
	Get-SOEComputerSignedDrivers

.LINK
	Get-SOEComputerRunKeys

.LINK
	Get-SOEComputerPatches

.LINK
    Get-SOEComputerBIOS

.LINK
    Get-SOEComputerPolicyFirewallRules

.LINK
    Get-SOEComputerUserFirewallRules

.LINK
    Get-SOEComputerWifiProfile

.LINK
    Get-SOEComputerWifiSignature

.LINK
    Save-SOEObject

#>
function Get-SOEComputerModel
{
    [CmdletBinding(SupportsShouldProcess=$true, 
                   PositionalBinding=$true,
                   HelpUri='http://www.xs4all.nl/~chi/ps/oh1549701.html',
                   ConfirmImpact='Low')]
    Param
    (
        #<p>Specifies the computer for which this cmdlet gets the model data.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false,
                   Position=0)]
        [Alias('Name')]
        [string]
        $ComputerName,

        #Specifies the user account credentials to use to perform this task. The default credentials are the credentials of the currently logged on user unless the cmdlet is run from an Active Directory PowerShell provider drive. If the cmdlet is run from such a provider drive, the account associated with the drive is the default.
        #
        #To specify this parameter, you can type a user name, such as "User1" or "Domain01\User01" or you can specify a PSCredential object. If you specify a user name for this parameter, the cmdlet prompts for a password.
        #
        #You can also create a PSCredential object by using a script or by using the Get-Credential cmdlet. You can then set the Credential parameter to the PSCredential object 
        #The following example shows how to create credentials.
        #$AdminCredentials = Get-Credential "Domain01\User01"
        #The following shows how to set the Credential parameter to these credentials.
        #-Credential $AdminCredentials
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
        if ((Test-Connection -ComputerName $ComputerName -Quiet -Count 1) -or $ComputerName -eq '.')
        {
            Write-Verbose "Querying $ComputerName for basic SOE Computer model information"
            $DTCollection = (Get-Date).ToUniversalTime()

            $Data = Get-WmiObject -ComputerName $ComputerName -Class Win32_ComputerSystem -ErrorAction SilentlyContinue @Param
            if ($Data -ne $null)
            {
                foreach ($Item in $Data)
                {
                    $Properties = [ordered]@{'ComputerName'       = $Item.__SERVER;
                                             'Vendor'             = $Item.Manufacturer;
                                             'Model'              = $Item.Model;
                                             'PhysicalMemorySize' = $Item.TotalPhysicalMemory;
                                             'ConsoleUser'        = [string]$Item.UserName;
                                             'DTCollection'       = $DTCollection }
                    $Return = New-Object -TypeName PSObject -Property $Properties
                    $Return.PSObject.TypeNames.Insert(0, 'SOE.ComputerModel.100')
                    Write-Output $Return
                }#foreach
            }
            else
            {
                Write-Verbose "Can't query $ComputerName for basic SOE Computer model information"
            }#Access?
        }
        else
        {
            Write-Verbose "$ComputerName is not online to query basic SOE Computer model information"
        }#Online?
    }
}

#endregion

#region Get-SOEComputerNICDetails

<#
.SYNOPSIS
    Gets the network interface card data from the local computer or a remote computer.

.DESCRIPTION
    The Get-SOEComputerNICDetails cmdlet gets a select part of the network interface card (NIC) data for all NICs present from the local or remote computer through WMI as SOE.ComputerNICDetails.100 object.

    The ComputerNICDetails data contains the following attributes:
    - ComputerName    
    - Vendor          
    - Name            
    - Service         
    - Index           
    - Enabled         
    - ConnectionStatus
    - PnpCapabilities  = Registery setting and not always availble
    - DTCollection    

.EXAMPLE
    PS C:\>Get-SOEComputerNICDetails -ComputerName .

        ComputerName     : King001
        Vendor           : Microsoft
        Name             : Microsoft Virtual WiFi Miniport Adapter
        Service          : vwifimp
        Index            : 17
        Enabled          : False
        ConnectionStatus : 4
        PnpCapabilities  : 24
        DTCollection     : 2015-12-20 16:23:54

    This command gets the NIC information from the local computer. (Only showing one example of the output)

.INPUTS
<string>

.OUTPUTS
    [SOE.ComputerNICDetails.100]

    This cmdlet returns a named System.Management.Automation.PSCustomObject object called SOE.ComputerNICDetails.100.

    Table information

    Saving the output of this cmdlet with the Save-SOEObject cmdlet in an SQL table requires the following table structure.

    CREATE TABLE dbo.ComputerNICDetails(
	    [ID] [int] IDENTITY(1,1) NOT NULL PRIMARY KEY,
	    [ComputerName] [varchar](25) NOT NULL,
	    [Vendor] [varchar](50) NULL,
	    [Name] [varchar](75) NOT NULL,
	    [Service] [varchar](25) NULL,
	    [Index] [smallint] NOT NULL,
	    [Enabled] [bit] NULL,
	    [ConnectionStatus] [smallint] NULL,
	    [PnpCapabilities] [smallint] NULL,
	    [DTCollection] [datetime] NOT NULL,
	    [DTCreation] [datetime] NOT NULL CONSTRAINT [DF_ComputerNICDetails_DTCreation] DEFAULT (GETDATE()),
	    [DTMutation] [datetime] NULL,
	    [DTDeletion] [datetime] NULL
        )

.NOTES
    --- Version history:
    Version 1.10 (2016-10-31, Kees Hiemstra)
    - Added parameter Credential.
    Version 1.02 (2015-12-22, Kees Hiemstra)
    - Bug fix: Increased the size of the [Name] field from 50 to 75 in the database table.
    Version 1.01 (2015-12-20, Kees Hiemstra)
    - Adding help.
    Version 1.00 (2015-12-09, Kees Hiemstra)
    - Initial version.

.LINK
    Get-SOEComputerModel

.LINK
    Get-SOEComputerOS

.LINK
	Get-SOEComputerBSoDEvents

.LINK
	Get-SOEComputerSignedDrivers

.LINK
	Get-SOEComputerRunKeys

.LINK
	Get-SOEComputerPatches

.LINK
    Get-SOEComputerBIOS

.LINK
    Get-SOEComputerUserFirewallRules

.LINK
    Get-SOEComputerPolicyFirewallRules

.LINK
    Get-SOEComputerWifiProfile

.LINK
    Get-SOEComputerWifiSignature

.LINK
    Save-SOEObject

#>
function Get-SOEComputerNICDetails
{
    [CmdletBinding(SupportsShouldProcess=$true, 
                   PositionalBinding=$true,
                   HelpUri='http://www.xs4all.nl/~chi/ps/oh1550301.html',
                   ConfirmImpact='Low')]
    Param
    (
        #Specifies the computer for which this cmdlet gets the NIC data.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false,
                   Position=0)]
        [Alias('Name')]
        [string]
        $ComputerName,

        #Specifies the user account credentials to use to perform this task. The default credentials are the credentials of the currently logged on user unless the cmdlet is run from an Active Directory PowerShell provider drive. If the cmdlet is run from such a provider drive, the account associated with the drive is the default.
        #
        #To specify this parameter, you can type a user name, such as "User1" or "Domain01\User01" or you can specify a PSCredential object. If you specify a user name for this parameter, the cmdlet prompts for a password.
        #
        #You can also create a PSCredential object by using a script or by using the Get-Credential cmdlet. You can then set the Credential parameter to the PSCredential object 
        #The following example shows how to create credentials.
        #$AdminCredentials = Get-Credential "Domain01\User01"
        #The following shows how to set the Credential parameter to these credentials.
        #-Credential $AdminCredentials
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
        if ((Test-Connection -ComputerName $ComputerName -Quiet -Count 1) -or $ComputerName -eq '.')
        {
            Write-Verbose "Querying $ComputerName for NIC Details information"
            $DTCollection = (Get-Date).ToUniversalTime()

            $Data = Get-WmiObject -Class Win32_NetworkAdapter -ComputerName $ComputerName -ErrorAction SilentlyContinue @Param
            if ($Data -ne $null)
            {
                foreach ($Item in $Data)
                {
                    $Properties = [ordered]@{'ComputerName'       = $Item.__SERVER;
                                             'Vendor'             = $Item.Manufacturer;
                                             'Name'               = $Item.ProductName;
                                             'Service'            = $Item.ServiceName;
                                             'Index'              = $Item.Index;
                                             'Enabled'            = $Item.NetEnabled;
                                             'ConnectionStatus'   = $Item.NetConnectionStatus;
                                             'PnpCapabilities'    = [int]-1;
                                             'DTCollection'       = $DTCollection }
                    if($ComputerName -in ('.', 'localhost') -or ($ComputerName.Split('.'))[0] -eq $env:COMPUTERNAME)
                    {
                        $Properties.PnpCapabilities = ((Get-ItemProperty -Path "HKLM:SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002BE10318}\$($Properties.Index.ToString().PadLeft(4, '0'))" -Name "PnpCapabilities" -ErrorAction SilentlyContinue)).PnpCapabilities
                    }
                    else
                    {
                        $Properties.PnpCapabilities = Invoke-Command -ComputerName $ComputerName -ArgumentList $Properties.Index -ScriptBlock {param($Index) ((Get-ItemProperty -Path "HKLM:SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002BE10318}\$($Index.ToString().PadLeft(4, '0'))" -Name "PnpCapabilities" -ErrorAction SilentlyContinue)).PnpCapabilities }
                    }
                    $Return = New-Object -TypeName PSObject -Property $Properties
                    $Return.PSObject.TypeNames.Insert(0, 'SOE.ComputerNICDetails.100')
                    Write-Output $Return
                }#foreach
            }
            else
            {
                Write-Verbose "Can't query $ComputerName for NIC Details information"
            }#Access?
        }
        else
        {
            Write-Verbose "$ComputerName is not online to query NIC Details information"
        }#Online?
    }
}

#endregion

#region Get-SOEComputerOS

<#
.SYNOPSIS
    Query the basic WMI OS information of the selected computer.
.DESCRIPTION
    Collect the following data from the selected computer through WMI and report those as SOE.ComputerOS.100 object.
    - ComputerName
    - OSName
    - OSServicePack
    - OSVersion
    - OSArchitecture
    - DTInstallation
    - DTBooting
    - DTLocal
    - Uptime
    - PhysicalMemorySize
    - PhysicalMemoryFree
    - VirtualMemorySize
    - VirtualMemoryFree
    - DTCollection
.EXAMPLE
    Get-SOEComputerOS -ComputerName .

    Will get the basic OS information from the local computer:
        ComputerName       : KING001
        OSName             : Microsoft Windows 7 Enterprise 
        OSServicePack      : Service Pack 1
        OSVersion          : 6.1.7601
        OSArchitecture     : 64-bit
        DTInstallation     : 2015-06-08 10:01:08
        DTBooting          : 2015-12-19 08:59:04
        DTLocal            : 2015-12-20 16:28:23
        Uptime             : 1.07:29:18.9194010
        PhysicalMemorySize : 8260672
        PhysicalMemoryFree : 4700728
        VirtualMemorySize  : 16519508
        VirtualMemoryFree  : 11110952
        DTCollection       : 2015-12-20 16:28:23

.OUTPUTS
    [SOE.ComputerOS.100]

.NOTES
    --- Version history:
    Version 1.10 (2016-10-31, Kees Hiemstra)
    - Added parameter Credential.
    Version 1.01 (2015-12-20, Kees Hiemstra)
    - Adding help.
    Version 1.00 (2015-12-06, Kees Hiemstra)
    - Initial version.

    --- Extra information
    Use the following SQL statement to create a table in the database if the data needs to be stored with Save-SOEObject.

    CREATE TABLE dbo.ComputerOS(
	    [ID] [int] IDENTITY(1,1) NOT NULL PRIMARY KEY,
	    [ComputerName] [varchar](25) NOT NULL,
	    [OSName] [varchar](50) NOT NULL,
	    [OSServicePack] [varchar](25) NOT NULL,
	    [OSVersion] [varchar](25) NOT NULL,
	    [OSArchitecture] [char](6) NOT NULL,
	    [DTInstallation] [smalldatetime] NOT NULL,
	    [DTBooting] [datetime] NOT NULL,
	    [DTLocal] [datetime] NOT NULL,
	    [PhysicalMemorySize] [bigint] NOT NULL,
	    [PhysicalMemoryFree] [bigint] NOT NULL,
	    [VirtualMemorySize] [bigint] NOT NULL,
	    [VirtualMemoryFree] [bigint] NOT NULL,
	    [DTCollection] [datetime] NOT NULL,
	    [DTCreation] [datetime] NOT NULL CONSTRAINT [DF_ComputerOS_DTCreation] DEFAULT (GETDATE()),
	    [DTMutation] [datetime] NULL,
	    [DTDeletion] [datetime] NULL
        )

.LINK
    Get-SOEComputerModel

.LINK
	Get-SOEComputerBSoDEvents

.LINK
	Get-SOEComputerNICDetails

.LINK
	Get-SOEComputerSignedDrivers

.LINK
	Get-SOEComputerRunKeys

.LINK
	Get-SOEComputerPatches

.LINK
    Get-SOEComputerBIOS

.LINK
    Get-SOEComputerUserFirewallRules

.LINK
    Get-SOEComputerPolicyFirewallRules

.LINK
    Get-SOEComputerWifiProfile

.LINK
    Get-SOEComputerWifiSignature

.LINK
    Save-SOEObject

#>
function Get-SOEComputerOS
{
    [CmdletBinding(SupportsShouldProcess=$true, 
                   PositionalBinding=$true,
                   HelpUri='http://www.xs4all.nl/~chi/ps/oh1549702.html',
                   ConfirmImpact='Low')]
    Param
    (
        #Computer name to query
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false,
                   Position=0)]
        [Alias('Name')]
        [string]
        $ComputerName,

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
        if ((Test-Connection -ComputerName $ComputerName -Quiet -Count 1) -or $ComputerName -eq '.')
        {
            Write-Verbose "Querying $ComputerName for basic SOE Computer OS information"
            $DTCollection = (Get-Date).ToUniversalTime()

            $Data = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction SilentlyContinue @Param
            if ($Data -ne $null)
            {
                foreach ($Item in $Data)
                {
                    $Properties = [ordered]@{'ComputerName'       = $Item.__SERVER;
                                             'OSName'             = [string]$Item.Caption;
                                             'OSServicePack'      = [string]$Item.CSDVersion;
                                             'OSVersion'          = [string]$Item.Version;
                                             'OSArchitecture'     = [string]$Item.OSArchitecture;
                                             'DTInstallation'     = [datetime][management.managementDateTimeConverter]::ToDateTime($Item.InstallDate).ToUniversalTime();
                                             'DTBooting'          = [datetime][management.managementDateTimeConverter]::ToDateTime($Item.LastBootUpTime).ToUniversalTime();
                                             'DTLocal'            = [datetime][management.managementDateTimeConverter]::ToDateTime($Item.LocalDateTime).ToUniversalTime();
                                             'Uptime'             = [timespan]0;
                                             'PhysicalMemorySize' = $Item.TotalVisibleMemorySize;
                                             'PhysicalMemoryFree' = $Item.FreePhysicalMemory;
                                             'VirtualMemorySize'  = $Item.TotalVirtualMemorySize;
                                             'VirtualMemoryFree'  = $Item.FreeVirtualMemory;
                                             'DTCollection'       = $DTCollection }
                    $Properties.Uptime = (New-TimeSpan -Start $Properties.DTBooting -End $Properties.DTLocal)
                    $Return = New-Object -TypeName PSObject -Property $Properties
                    $Return.PSObject.TypeNames.Insert(0, 'SOE.ComputerOS.100')
                    Write-Output $Return
                }#foreach
            }
            else
            {
                Write-Verbose "Can't query $ComputerName for basic SOE Computer OS information"
            }#Access?
        }
        else
        {
            Write-Verbose "$ComputerName is not online to query basic SOE Computer OS information"
        }#Online?
    }
}

#endregion

#region Get-SOEComputerSignedDrivers

<#
.SYNOPSIS
    Query the basic WMI on installed driver information of the selected computer.
.DESCRIPTION
    Collect the following data from the selected computer through WMI and report those as SOE.ComputerNICDetails.100 object.
    - ComputerName    
    - Vendor          
    - Name            
    - Service         
    - Index           
    - Enabled         
    - ConnectionStatus
    - PnpCapabilities  = Registery setting and not always availble
    - DTCollection    
.EXAMPLE
    Get-SOEComputerSignedDrivers -ComputerName .

    Will get the NIC detail information from the local computer:
        ComputerName : KING001
        DeviceID     : ROOT\ACPI_HAL\0000
        DeviceClass  : COMPUTER
        Vendor       : (Standard computers)
        Name         : ACPI x64-based PC
        Version      : 6.1.7600.16385
        InfName      : hal.inf
        Signer       : Microsoft Windows
        DTCollection : 2015-12-20 16:35:55

.OUTPUTS
    [SOE.ComputerSignedDrivers.100]

.NOTES
    --- Version history:
    Version 1.10 (2016-10-31, Kees Hiemstra)
    - Added parameter Credential.
    Version 1.04 (2015-12-29, Kees Hiemstra)
    - Bug fix: Increased the size of the [DeviceClass] field from 25 to 50 in the database table.
    Version 1.03 (2015-12-22, Kees Hiemstra)
    - Bug fix: Increased the size of the [DeviceID] field from 128 to 255 in the database table.
    - Bug fix: Increased the size of the [Name] field from 75 to 128 in the database table.
    Version 1.02 (2015-12-21, Kees Hiemstra)
    - Updated help.
    Version 1.01 (2015-12-20, Kees Hiemstra)
    - Adding help.
    Version 1.00 (2015-12-09, Kees Hiemstra)
    - Initial version.

    --- Extra information
    Use the following SQL statement to create a table in the database if the data needs to be stored with Save-SOEObject.

    CREATE TABLE [dbo].[ComputerSignedDrivers](
	    [ID] [int] IDENTITY(1,1) NOT NULL PRIMARY KEY,
	    [ComputerName] [varchar](25) NOT NULL,
	    [DeviceID] [varchar](255) NOT NULL,
	    [DeviceClass] [varchar](50) NULL,
	    [Vendor] [varchar](50) NULL,
	    [Name] [varchar](128) NULL,
	    [Version] [varchar](50) NULL,
	    [InfName] [varchar](25) NULL,
	    [Signer] [varchar](75) NULL,
	    [DTCollection] [datetime] NOT NULL,
	    [DTCreation] [datetime] NOT NULL CONSTRAINT [DF_ComputerSignedDrivers_DTCreation] DEFAULT (GETDATE()),
	    [DTMutation] [datetime] NULL,
	    [DTDeletion] [datetime] NULL
	    )

.LINK
    Get-SOEComputerModel

.LINK
    Get-SOEComputerOS

.LINK
	Get-SOEComputerBSoDEvents

.LINK
	Get-SOEComputerNICDetails

.LINK
	Get-SOEComputerRunKeys

.LINK
	Get-SOEComputerPatches

.LINK
    Get-SOEComputerBIOS

.LINK
    Get-SOEComputerUserFirewallRules

.LINK
    Get-SOEComputerPolicyFirewallRules

.LINK
    Get-SOEComputerWifiProfile

.LINK
    Get-SOEComputerWifiSignature

.LINK
    Save-SOEObject

#>
function Get-SOEComputerSignedDrivers
{
    [CmdletBinding(SupportsShouldProcess=$true, 
                   PositionalBinding=$true,
                   HelpUri='http://www.xs4all.nl/~chi/ps/oh1550302.html',
                   ConfirmImpact='Low')]
    Param
    (
        #Computer name to query
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false,
                   Position=0)]
        [Alias('Name')]
        [string]
        $ComputerName,

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
        if ((Test-Connection -ComputerName $ComputerName -Quiet -Count 1) -or $ComputerName -eq '.')
        {
            Write-Verbose "Querying $ComputerName for Signed Drivers information"
            $DTCollection = (Get-Date).ToUniversalTime()

            $Data = Get-WmiObject -Class Win32_PnPSignedDriver -ComputerName $ComputerName -ErrorAction SilentlyContinue @Param
            if ($Data -ne $null)
            {
                foreach ($Item in $Data)
                {
                    $Properties = [ordered]@{'ComputerName'       = $Item.__SERVER;
                                             'DeviceID'           = $Item.DeviceID;
                                             'DeviceClass'        = $Item.DeviceClass;
                                             'Vendor'             = $Item.Manufacturer;
                                             'Name'               = $Item.DeviceName;
                                             'Version'            = [string] $Item.DriverVersion;
                                             'InfName'            = $Item.InfName;
                                             'Signer'             = $Item.Signer;
                                             'DTCollection'       = $DTCollection }
                    $Return = New-Object -TypeName PSObject -Property $Properties
                    $Return.PSObject.TypeNames.Insert(0, 'SOE.ComputerSignedDrivers.100')
                    Write-Output $Return
                }#foreach
            }
            else
            {
                Write-Verbose "Can't query $ComputerName for Signed Drivers information"
            }#Access?
        }
        else
        {
            Write-Verbose "$ComputerName is not online to query Signed Drivers information"
        }#Online?
    }
}

#endregion

#region Get-SOEComputerRunKeys

<#
.SYNOPSIS
    Collect the RunKey entries from registry of the selected computer.
.DESCRIPTION
    Collect the following data from the selected computer through registry and report those as SOE.ComputerRunKeys.100 object.
    - ComputerName
    - Key (abbriviation)
    - Name
    - Value
    - User
    - DTCollection

    The registry will be queried on the following keys for both the local machine and user:
    - SOFTWARE\Microsoft\Windows\CurrentVersion\Run
    - SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce
    - SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Run
    - SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\RunOnce

    The used technique doesn't require the Server service to be active on the target computer.
.EXAMPLE
    Get-SOEComputerRunKeys -ComputerName .

    Will query the local computer with your account.
        ComputerName : HPNLDEV06
        Key          : LM_W64_Run
        Name         : HPConnectionManager
        Value        : C:\Program Files (x86)\Hewlett-Packard\HP Connection Manager\HPCMDelayStart.exe
        User         :
        DTCollection : 2015-12-20 16:37:57

.EXAMPLE
    Get-SOEComputerRunKeys -ComputerName LocalHost -UserAccount "Camelot\KingArthur"

    Will query the local computer with the account Camelot\KingArthur. An error message will be shown if the account isn't found and instead your account will be used.
        ComputerName : KING001
        Key          : US_W32_Run
        Name         : Skype
        Value        : "C:\Program Files (x86)\Skype\Phone\Skype.exe" /minimized /regrun
        User         : Camelot\KingArthur
        DTCollection : 2015-12-20 16:37:57
.EXAMPLE
    Get-SOEComputerModel -ComputerName KING001 | Get-SOEComputerRunKeys 

    The Get-SOEComputerModel will retrieve who is logged on to the console of the computer KING001 and will query that account. If no one is logged on, your account will be used.
        ComputerName : KING001
        Key          : US_W32_Run
        Name         : OneDrive
        Value        : "C:\Users\Camelot\KingArthur\AppData\Local\Microsoft\OneDrive\OneDrive.exe" /background
        User         : Camelot\KingArthur
        DTCollection : 2015-12-20 16:37:57

.OUTPUTS
    [SOE.ComputerRunKeys.100]

.NOTES
    --- Version history:
    Version 1.20 (2016-10-31, Kees Hiemstra)
    - Added parameter Credential.
    Version 1.11 (2015-12-20, Kees Hiemstra)
    - Adding help.
    Version 1.10 (2015-12-14, Kees Hiemstra)
    - Changed the way to read the remote registry setting without having the Service service to be running on the target computer, now it also works with PowerShell 2 clients.
    Version 1.00 (2015-12-13, Kees Hiemstra)
    - Initial version.

    --- Extra information
    Use the following SQL statement to create a table in the database if the data needs to be stored with Save-SOEObject.

    CREATE TABLE dbo.ComputerRunKeys(
	    [ID] [int] IDENTITY(1,1) NOT NULL PRIMARY KEY,
	    [ComputerName] [varchar](25) NOT NULL,
	    [Key] [char](14) NOT NULL,
	    [Name] [varchar](75) NULL,
	    [Value] [varchar](255) NULL,
	    [User] [varchar](75) NULL,
	    [DTCollection] [datetime] NOT NULL,
	    [DTCreation] [datetime] NULL CONSTRAINT [DF_ComputerRunKeys_DTCreation] DEFAULT (GETDATE()),
	    [DTMutation] [datetime] NULL,
	    [DTDeletion] [datetime] NULL
        )

.LINK
    Get-SOEComputerModel

.LINK
    Get-SOEComputerOS

.LINK
	Get-SOEComputerBSoDEvents

.LINK
	Get-SOEComputerNICDetails

.LINK
	Get-SOEComputerSignedDrivers

.LINK
	Get-SOEComputerPatches

.LINK
    Get-SOEComputerBIOS

.LINK
    Get-SOEComputerUserFirewallRules

.LINK
    Get-SOEComputerPolicyFirewallRules

.LINK
    Get-SOEComputerWifiProfile

.LINK
    Get-SOEComputerWifiSignature

.LINK
    Save-SOEObject

#>
function Get-SOEComputerRunKeys
{
    [CmdletBinding(SupportsShouldProcess=$true, 
                   PositionalBinding=$true,
                   HelpUri='http://www.xs4all.nl/~chi/ps/oh1550701.html',
                   ConfirmImpact='Low')]
    Param
    (
        #Computer name to query
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0)]
        [Alias('Name')]
        [string]
        $ComputerName,

        #User account to query e.g. DomainName\SAMAccountName.
        #By default the account will be used that is starting this query.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false,
                   Position=1)]
        [Alias('ConsoleUser')]
        [string]
        $UserAccount,

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

        $ScriptBlock = {param([string]$SID)
            $RegKeys = @{
                'LM_W32_Run'     = 'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Run';
                'LM_W32_RunOnce' = 'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce';
                'LM_W64_Run'     = 'HKLM:SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Run';
                'LM_W64_RunOnce' = 'HKLM:SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\RunOnce';
                'US_W32_Run'     = 'HKUS:SOFTWARE\Microsoft\Windows\CurrentVersion\Run';
                'US_W32_RunOnce' = 'HKUS:SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce';
                'US_W64_Run'     = 'HKUS:SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Run';
                'US_W64_RunOnce' = 'HKUS:SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\RunOnce'
                }
            
            Get-PSDrive -Name HKUS  -ErrorAction SilentlyContinue | Remove-PSDrive -ErrorAction SilentlyContinue
            if ([string]::IsNullOrEmpty($SID))
            {
                New-PSDrive -Name HKUS -PSProvider Registry -Root HKCU: | Out-Null
            }
            else
            {
                New-PSDrive -Name HKUS -PSProvider Registry -Root Registry::"HKEY_USERS\$SID" | Out-Null
            }

            foreach ($Key in $RegKeys.GetEnumerator())
            {
                $Items = Get-Item -Path $Key.Value -ErrorAction SilentlyContinue
                foreach ($Name in ($Items | Select-Object -ExpandProperty Property))
                {
                    $Properties = @{'Key'   = $Key.Key;
                                    'Name'  = $Name;
                                    'Value' = (Get-ItemProperty -Path $Key.Value -Name $Name)."$Name"}
                    $Return = New-Object -TypeName PSObject -Property $Properties
                    Write-Output $Return
                }#foreach Items
            }#foreach RegKeys
            Get-PSDrive -Name HKUS  -ErrorAction SilentlyContinue | Remove-PSDrive -ErrorAction SilentlyContinue
        }#ScriptBlock
    }
    Process
    {
        if ($ComputerName -in ('.', 'LocalHost'))
        {
            Write-Verbose "Change name to $($env:ComputerName)"
            $ComputerName = $env:ComputerName
        }

        if ((Test-Connection -ComputerName $ComputerName -Quiet -Count 1))
        {
            Write-Verbose "Querying $ComputerName for RunKeys information"
            $DTCollection = (Get-Date).ToUniversalTime()

            if ([string]::IsNullOrEmpty($UserAccount))
            {
                $QHash = @{}
                Write-Verbose "Querying on current user"
            }
            else
            {
                try
                {
                    $User = New-Object System.Security.Principal.NTAccount($UserAccount)
                    $SID = $User.Translate([System.Security.Principal.SecurityIdentifier])
                    $QHash = @{'ArgumentList' = ($SID.Value)}
                    Write-Verbose "Querying on user $UserAccount"
                }
                catch
                {
                    Write-Error "UserAccount $UserAccount can't be translated to a SID"
                    Write-Verbose "Querying on current user"
                    $QHash = @{}
                }
            }

            if ($ComputerName -eq $env:ComputerName)
            {
                $Data = Invoke-Command -ScriptBlock $ScriptBlock @QHash @Param
            }
            else
            {
                $Data = Invoke-Command -ComputerName $ComputerName -ScriptBlock $ScriptBlock @QHash @Param
            }

            if ($Data -ne $null)
            {
                foreach ($Item in $Data)
                {
                    $Properties = [ordered]@{'ComputerName'       = $ComputerName;
                                             'Key'                = $Item.Key;
                                             'Name'               = $Item.Name;
                                             'Value'              = $Item.Value;
                                             'User'               = '';
                                             'DTCollection'       = $DTCollection }
                    $Return = New-Object -TypeName PSObject -Property $Properties

                    if ($Return.Key -like 'US_*')
                    {
                        if ([string]::IsNullOrEmpty($SID))
                        {
                            $Return.User = 'Current user'
                        }
                        else
                        {
                            $Return.User = $UserAccount
                        }
                    }

                    $Return.PSObject.TypeNames.Insert(0, 'SOE.ComputerRunKeys.100')
                    Write-Output $Return
                }#foreach
            }
            else
            {
                Write-Verbose "Can't query $ComputerName for RunKeys information"
            }#Access?
        }
        else
        {
            Write-Verbose "$ComputerName is not online to query RunKeys information"
        }#Online?
    }
}

#endregion

#region Get-SOEComputerPatches

<#
.SYNOPSIS
    Query the basic WMI hot fix/patch information of the selected computer.
.DESCRIPTION
   Collect the following data from the selected computer through WMI and report those as SOE.ComputerPatches.100 object.
   - ComputerName
   - HotFixID
   - PatchType
   - Lynk
   - User
   - DTInstallation
   - DTCollection

.EXAMPLE
    Get-SOEComputerPatches -ComputerName .

    Will get the basic model information from the local computer:
        ComputerName   : KING001
        HotFixID       : KB982018
        PatchType      : Update
        Link           : http://support.microsoft.com/?kbid=982018
        User           : NT AUTHORITY\SYSTEM
        DTInstallation : 2015-04-20 00:00:00
        DTCollection   : 2015-12-20 17:52:47

.OUTPUTS
    [SOE.ComputerPatches.100]

.NOTES
    --- Version history:
    Version 1.10 (2016-10-31, Kees Hiemstra)
    - Added parameter Credential.
    Version 1.00 (2015-12-20, Kees Hiemstra)
    - Initial version.

    --- Extra information
    Use the following SQL statement to create a table in the database if the data needs to be stored with Save-SOEObject.

    CREATE TABLE [dbo].[ComputerPatches](
	    [ID] [int] IDENTITY(1,1) NOT NULL PRIMARY KEY,
	    [ComputerName] [varchar](25) NOT NULL,
	    [HotFixID] [char](10) NOT NULL,
	    [PatchType] [char](15) NOT NULL,
	    [Link] [varchar](50) NOT NULL,
	    [User] [varchar](75) NOT NULL,
	    [DTInstallation] [date] NOT NULL,
	    [DTCollection] [datetime] NOT NULL,
	    [DTCreation] [datetime] NOT NULL CONSTRAINT [DF_ComputerPatches_DTCreation] DEFAULT (GETDATE()),
	    [DTMutation] [datetime] NULL,
	    [DTDeletion] [datetime] NULL
        )

.LINK
    Get-SOEComputerModel

.LINK
    Get-SOEComputerOS

.LINK
	Get-SOEComputerBSoDEvents

.LINK
	Get-SOEComputerNICDetails

.LINK
	Get-SOEComputerSignedDrivers

.LINK
	Get-SOEComputerRunKeys

.LINK
    Get-SOEComputerBIOS

.LINK
    Get-SOEComputerUserFirewallRules

.LINK
    Get-SOEComputerPolicyFirewallRules

.LINK
    Get-SOEComputerWifiProfile

.LINK
    Get-SOEComputerWifiSignature

.LINK
    Save-SOEObject

#>
function Get-SOEComputerPatches
{
    [CmdletBinding(SupportsShouldProcess=$true, 
                   PositionalBinding=$true,
                   HelpUri='http://www.xs4all.nl/~chi/ps/oh1551701.html',
                   ConfirmImpact='Low')]
    Param
    (
        #Computer name to query
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false,
                   Position=0)]
        [Alias('Name')]
        [string]
        $ComputerName,

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
        if ((Test-Connection -ComputerName $ComputerName -Quiet -Count 1) -or $ComputerName -eq '.')
        {
            Write-Verbose "Querying $ComputerName for basic SOE Computer patch information"
            $DTCollection = (Get-Date).ToUniversalTime()

            $Data = Get-WmiObject -ComputerName $ComputerName -Class Win32_QuickFixEngineering -ErrorAction SilentlyContinue @Param
            if ($Data -ne $null)
            {
                foreach ($Item in $Data)
                {
                    $Properties = [ordered]@{'ComputerName'       = $Item.__SERVER;
                                             'HotFixID'           = $Item.HotFixID;
                                             'PatchType'          = $Item.Description;
                                             'Link'               = $Item.Caption;
                                             'User'               = [string]$Item.InstalledBy;
                                             'DTInstallation'     = [datetime]$Item.InstalledOn;
                                             'DTCollection'       = $DTCollection }
                    $Return = New-Object -TypeName PSObject -Property $Properties
                    $Return.PSObject.TypeNames.Insert(0, 'SOE.ComputerPatches.100')
                    Write-Output $Return
                }#foreach
            }
            else
            {
                Write-Verbose "Can't query $ComputerName for basic SOE Computer patch information"
            }#Access?
        }
        else
        {
            Write-Verbose "$ComputerName is not online to query basic SOE Computer patch information"
        }#Online?
    }
}

#endregion

#region Get-SOEComputerUserFirewallRules

<#
.SYNOPSIS
    Collect the user created firewall rule entries from registry of the selected computer.
.DESCRIPTION
    Collect the following data from the selected computer through registry and report those as SOE.ComputerUserFirewallRules.100 object.
    - ComputerName
    - Key (abbriviation)
    - Name
    - Value
    - DTCollection

    The registry will be queried on the following key on the local machine:
    - SYSTEM\CurrentControlSet\services\SharedAccess\Parameters\FirewallPolicy\FirewallRules

    The used technique doesn't require the Server service to be active on the target computer.
.EXAMPLE
    Get-SOEComputerUserFirewallRules -ComputerName LocalHost

    Will query the local computer.
        ComputerName : KING001
        Key          : LM_SYSTEM
        Name         : TCP Query User{C1E71886-6D5D-432D-8B1D-E0BA9659D98A}C:\program files (x86)\mirc\mirc.exe
        Value        : v2.10|Action=Allow|Active=TRUE|Dir=In|Protocol=6|Profile=Private|App=C:\program files (x86)\mirc\mirc.exe|Name=mIRC|Desc=mIRC|Defer=User|
        DTCollection : 2016-05-18 11:20:23
.OUTPUTS
    [SOE.ComputerUserFirewallRules.100]

.NOTES
    --- Version history:
    Version 1.10 (2016-10-31, Kees Hiemstra)
    - Added parameter Credential.
    Version 1.00 (2016-05-18, Kees Hiemstra)
    - Initial version.

    --- Extra information
    Use the following SQL statement to create a table in the database if the data needs to be stored with Save-SOEObject.

    CREATE TABLE dbo.ComputerUserFirewallRules(
	    [ID] [int] IDENTITY(1,1) NOT NULL PRIMARY KEY,
	    [ComputerName] [varchar](25) NOT NULL,
	    [Key] [char](14) NOT NULL,
	    [Name] [varchar](255) NULL,
	    [Value] [varchar](512) NULL,
	    [DTCollection] [datetime] NOT NULL,
	    [DTCreation] [datetime] NULL CONSTRAINT [DF_ComputerUserFirewallRules_DTCreation] DEFAULT (GETDATE()),
	    [DTMutation] [datetime] NULL,
	    [DTDeletion] [datetime] NULL
      )

.LINK
    Get-SOEComputerModel

.LINK
    Get-SOEComputerOS

.LINK
	Get-SOEComputerBSoDEvents

.LINK
	Get-SOEComputerNICDetails

.LINK
	Get-SOEComputerSignedDrivers

.LINK
    Get-Get-SOEComputerRunKeys

.LINK
	Get-SOEComputerPatches

.LINK
    Get-SOEComputerBIOS

.LINK
    Get-SOEComputerPolicyFirewallRules

.LINK
    Get-SOEComputerWifiProfile

.LINK
    Get-SOEComputerWifiSignature

.LINK
    Save-SOEObject

#>
function Get-SOEComputerUserFirewallRules
{
    [CmdletBinding(SupportsShouldProcess=$true, 
                   PositionalBinding=$true,
                   HelpUri='http://www.xs4all.nl/~chi/ps/oh1621401.html',
                   ConfirmImpact='Low')]
    Param
    (
        #Computer name to query
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0)]
        [Alias('Name')]
        [string]
        $ComputerName,

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

        $ScriptBlock = {param()
            $RegKeys = @{LM_SYSTEM = 'HKLM:SYSTEM\CurrentControlSet\services\SharedAccess\Parameters\FirewallPolicy\FirewallRules'}
            
            foreach ($Key in $RegKeys.GetEnumerator())
            {
                $Items = Get-Item -Path $Key.Value -ErrorAction SilentlyContinue
                foreach ($Name in ($Items | Select-Object -ExpandProperty Property | Where-Object { $_ -match "^(TCP|UDP)\sQuery\sUser*." }))
                {
                    $Properties = @{'Key'   = $Key.Key;
                                    'Name'  = $Name;
                                    'Value' = (Get-ItemProperty -Path $Key.Value -Name $Name)."$Name"}
                    $Return = New-Object -TypeName PSObject -Property $Properties
                    Write-Output $Return
                }#foreach Items
            }#foreach RegKeys
        }#ScriptBlock
    }
    Process
    {
        if ($ComputerName -in ('.', 'LocalHost'))
        {
            Write-Verbose "Change name to $($env:ComputerName)"
            $ComputerName = $env:ComputerName
        }

        if ((Test-Connection -ComputerName $ComputerName -Quiet -Count 1))
        {
            Write-Verbose "Querying $ComputerName for User firewall rules information"
            $DTCollection = (Get-Date).ToUniversalTime()

            if ($ComputerName -eq $env:ComputerName)
            {
                $Data = Invoke-Command -ScriptBlock $ScriptBlock @Param
            }
            else
            {
                $Data = Invoke-Command -ComputerName $ComputerName -ScriptBlock $ScriptBlock @Param
            }

            if ($Data -ne $null)
            {
                foreach ($Item in $Data)
                {
                    $Properties = [ordered]@{'ComputerName'       = $ComputerName;
                                             'Key'                = $Item.Key;
                                             'Name'               = $Item.Name;
                                             'Value'              = $Item.Value;
                                             'DTCollection'       = $DTCollection }
                    $Return = New-Object -TypeName PSObject -Property $Properties

                    $Return.PSObject.TypeNames.Insert(0, 'SOE.ComputerUserFirewallRules.100')
                    Write-Output $Return
                }#foreach
            }
            else
            {
                Write-Verbose "Can't query $ComputerName for User firewall rules information"
            }#Access?
        }
        else
        {
            Write-Verbose "$ComputerName is not online to query User firewall rules information"
        }#Online?
    }
}

#endregion

#region Get-SOEComputerPolicyFirewallRules

<#
.SYNOPSIS
    Collect the policy firewall rule entries from registry of the selected computer.
.DESCRIPTION
    Collect the following data from the selected computer through registry and report those as SOE.ComputerPolicyFirewallRules.100 object.
    - ComputerName
    - Key (abbriviation)
    - Name
    - Value
    - DTCollection

    The registry will be queried on the following key on the local machine:
    - SOFTWARE\Policies\Microsoft\WindowsFirewall\FirewallRules

    The used technique doesn't require the Server service to be active on the target computer.
.EXAMPLE
    Get-SOEComputerPolicyFirewallRules -ComputerName LocalHost

    Will query the local computer.
        ComputerName : KING001
        Key          : LM_SOFTWARE
        Name         : {BAABBB17-9FA6-413D-B6BC-1761D4672F1E}
        Value        : v2.20|Action=Allow|Active=TRUE|Dir=In|Profile=Domain|App=C:\Windows\SysWOW64\ftp.exe|Name=File Transfer Program (ftp.exe // 64bit)|Desc=2016-05-19|
        DTCollection : 2016-05-19 15:30:34
.OUTPUTS
    [SOE.ComputerPolicyFirewallRules.100]

.NOTES
    --- Version history:
    Version 1.10 (2016-10-31, Kees Hiemstra)
    - Added parameter Credential.
    Version 1.00 (2016-05-19, Kees Hiemstra)
    - Initial version.

.LINK
    Get-SOEComputerModel

.LINK
    Get-SOEComputerOS

.LINK
	Get-SOEComputerBSoDEvents

.LINK
	Get-SOEComputerNICDetails

.LINK
	Get-SOEComputerSignedDrivers

.LINK
    Get-Get-SOEComputerRunKeys

.LINK
	Get-SOEComputerPatches

.LINK
    Get-SOEComputerBIOS

.LINK
    Get-SOEComputerUserFirewallRules

.LINK
    Get-SOEComputerWifiProfile

.LINK
    Get-SOEComputerWifiSignature

.LINK
    Save-SOEObject

#>
function Get-SOEComputerPolicyFirewallRules
{
    [CmdletBinding(SupportsShouldProcess=$true, 
                   PositionalBinding=$true,
                   HelpUri='http://www.xs4all.nl/~chi/ps/oh1621301.html',
                   ConfirmImpact='Low')]
    Param
    (
        #Computer name to query
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0)]
        [Alias('Name')]
        [string]
        $ComputerName,

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
    
        $ScriptBlock = {param()
            $RegKeys = @{LM_SOFTWARE = 'HKLM:SOFTWARE\Policies\Microsoft\WindowsFirewall\FirewallRules'}
            
            foreach ($Key in $RegKeys.GetEnumerator())
            {
                $Items = Get-Item -Path $Key.Value -ErrorAction SilentlyContinue
                foreach ($Name in ($Items | Select-Object -ExpandProperty Property))
                {
                    $Properties = @{'Key'   = $Key.Key;
                                    'Name'  = $Name;
                                    'Value' = (Get-ItemProperty -Path $Key.Value -Name $Name)."$Name"}
                    $Return = New-Object -TypeName PSObject -Property $Properties
                    Write-Output $Return
                }#foreach Items
            }#foreach RegKeys
        }#ScriptBlock
    }
    Process
    {
        if ($ComputerName -in ('.', 'LocalHost'))
        {
            Write-Verbose "Change name to $($env:ComputerName)"
            $ComputerName = $env:ComputerName
        }

        if ((Test-Connection -ComputerName $ComputerName -Quiet -Count 1))
        {
            Write-Verbose "Querying $ComputerName for Policy firewall rules information"
            $DTCollection = (Get-Date).ToUniversalTime()

            if ($ComputerName -eq $env:ComputerName)
            {
                $Data = Invoke-Command -ScriptBlock $ScriptBlock @Param
            }
            else
            {
                $Data = Invoke-Command -ComputerName $ComputerName -ScriptBlock $ScriptBlock @Param
            }

            if ($Data -ne $null)
            {
                foreach ($Item in $Data)
                {
                    $Properties = [ordered]@{'ComputerName'       = $ComputerName;
                                             'Key'                = $Item.Key;
                                             'Name'               = $Item.Name;
                                             'Value'              = $Item.Value;
                                             'DTCollection'       = $DTCollection }
                    $Return = New-Object -TypeName PSObject -Property $Properties

                    $Return.PSObject.TypeNames.Insert(0, 'SOE.ComputerPolicyFirewallRules.100')
                    Write-Output $Return
                }#foreach
            }
            else
            {
                Write-Verbose "Can't query $ComputerName for Policy firewall rules information"
            }#Access?
        }
        else
        {
            Write-Verbose "$ComputerName is not online to query Policy firewall rules information"
        }#Online?
    }
}

#endregion

#region Get-SOEComputerWifiProfile

<#
.SYNOPSIS
    Collect the Wifi profile entries from registry of the selected computer.
.DESCRIPTION
    Collect the following data from the selected computer through registry and report those as SOE.ComputerWifiProfiles.100 object.
    - ComputerName
    - ProfileGUID
    - ProfileName
    - Description
    - Managed
    - Category
    - NameType
    - CategoryType
    - IconType
    - Count
    - DTCollection

    The registry will be queried on the following key on the local machine:
    - SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Profiles

    The used technique doesn't require the Server service to be active on the target computer.
.EXAMPLE
    Get-SOEComputerWifiProfile -ComputerName LocalHost

    (for this example you'll need to run PowerShell as Administrator, not required for remove access)

    Will query the local computer.
        ComputerName      : KING001
        ProfileGUID       : {B7EDB5C0-8B2C-4361-99EB-83761934C246}
        ProfileName       : camelot.org
        Description       : camelot.org
        Managed           : 1
        Category          : 2
        NameType          : 6
        CategoryType      : 
        IconType          : 
        DateCreated       : 2016-04-07 15:37:17
        DateLastConnected : 2016-04-07 15:37:17
        Count             : 2
        DTCollection      : 2016-06-02 06:13:25
.OUTPUTS
    [SOE.ComputerWifiProfile.100]

.NOTES
    --- Version history:
    Version 2.10 (2016-10-31, Kees Hiemstra)
    - Added parameter Credential.
    Version 2.00 (2016-06-02, Kees Hiemstra)
    - Replaced [Key] attribute with [ProfileGUID]
    - Added [DateCreated] and [DateLastConnected]
    Version 1.00 (2016-05-30, Kees Hiemstra)
    - Initial version.

    --- Extra information
    Use the following SQL statement to create a table in the database if the data needs to be stored with Save-SOEObject.

    CREATE TABLE dbo.ComputerWifiProfile(
	    [ID] [int] IDENTITY(1,1) NOT NULL PRIMARY KEY,
	    [ComputerName] [varchar](25) NOT NULL,
	    [ProfileGUID] [char](38) NOT NULL,
	    [ProfileName] [varchar](255) NULL,
	    [Description] [varchar](255) NULL,
	    [Managed] [int] NULL,
	    [Category] [int] NULL,
	    [NameType] [int] NULL,
	    [CategoryType] [int] NULL,
	    [IconType] [int] NULL,
        [DateCreated] [datetime] NULL,
        [DateLastConnected] [datetime] NULL,
	    [DTCollection] [datetime] NOT NULL,
	    [DTCreation] [datetime] NULL CONSTRAINT [DF_ComputerWifiProfile_DTCreation] DEFAULT (GETDATE()),
	    [DTMutation] [datetime] NULL,
	    [DTDeletion] [datetime] NULL
      )

.LINK
    Get-SOEComputerModel

.LINK
    Get-SOEComputerOS

.LINK
	Get-SOEComputerBSoDEvents

.LINK
	Get-SOEComputerNICDetails

.LINK
	Get-SOEComputerSignedDrivers

.LINK
    Get-Get-SOEComputerRunKeys

.LINK
	Get-SOEComputerPatches

.LINK
    Get-SOEComputerBIOS

.LINK
    Get-SOEComputerUserFirewallRules

.LINK
    Get-SOEComputerPolicyFirewallRules

.LINK
    Save-SOEObject

#>
function Get-SOEComputerWifiProfile
{
    [CmdletBinding(SupportsShouldProcess=$true, 
                   PositionalBinding=$true,
                   HelpUri='http://www.xs4all.nl/~chi/ps/oh1623101.html',
                   ConfirmImpact='Low')]
    Param
    (
        #Computer name to query
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0)]
        [Alias('Name')]
        [string]
        $ComputerName,

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

        $ScriptBlock = {param()
            $AllWifiProfiles = Get-ChildItem -Path 'HKLM:SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Profiles' -ErrorAction SilentlyContinue
            foreach ($WifiProfile in $AllWifiProfiles)
            {
                $Items = Get-Item -Path $WifiProfile.Name.Replace('HKEY_LOCAL_MACHINE', 'HKLM:') -ErrorAction SilentlyContinue
                foreach ($Name in ($Items))
                {
                    $Hex = [System.BitConverter]::ToString($Name.GetValue("DateCreated")) -split '-'  
                    $Year = [Convert]::ToInt32($Hex[1]+$Hex[0],16)
                    $Month = [Convert]::ToInt32($Hex[3]+$Hex[2],16)
                    $Day = [Convert]::ToInt32($Hex[7]+$Hex[6],16)
                    $Hour = [Convert]::ToInt32($Hex[9]+$Hex[8],16)
                    $Minute = [Convert]::ToInt32($Hex[11]+$Hex[10],16)
                    $Second = [Convert]::ToInt32($Hex[13]+$Hex[12],16)
                    $DateCreated = Get-Date -Year $Year -Month $Month -Day $Day -Hour $Hour -Minute $Minute -Second $Second
          
                    $Hex = [System.BitConverter]::ToString($Name.GetValue("DateLastConnected")) -split '-'  
                    $Year = [Convert]::ToInt32($Hex[1]+$Hex[0],16)
                    $Month = [Convert]::ToInt32($Hex[3]+$Hex[2],16)
                    $Day = [Convert]::ToInt32($Hex[7]+$Hex[6],16)
                    $Hour = [Convert]::ToInt32($Hex[9]+$Hex[8],16)
                    $Minute = [Convert]::ToInt32($Hex[11]+$Hex[10],16)
                    $Second = [Convert]::ToInt32($Hex[13]+$Hex[12],16)
                    $DateLastConnected = Get-Date -Year $Year -Month $Month -Day $Day -Hour $Hour -Minute $Minute -Second $Second

                    $Properties = @{'ProfileGUID'       = $WifiProfile.PSChildName;
                                    'ProfileName'       = $Name.GetValue("ProfileName");
                                    'Description'       = $Name.GetValue("Description");
                                    'Managed'           = $Name.GetValue("Managed");
                                    'Category'          = $Name.GetValue("Category");
                                    'NameType'          = $Name.GetValue("NameType");
                                    'CategoryType'      = $Name.GetValue("CategoryType");
                                    'IconType'          = $Name.GetValue("IconType");
                                    'DateCreated'       = $DateCreated;
                                    'DateLastConnected' = $DateLastConnected}
                    $Return = New-Object -TypeName PSObject -Property $Properties
                    Write-Output $Return
                }
            }#foreach Profile
        }#ScriptBlock
    }
    Process
    {
        if ($ComputerName -in ('.', 'LocalHost'))
        {
            Write-Verbose "Change name to $($env:ComputerName)"
            $ComputerName = $env:ComputerName
        }

        if ((Test-Connection -ComputerName $ComputerName -Quiet -Count 1))
        {
            Write-Verbose "Querying $ComputerName for Wifi profile information"
            $DTCollection = (Get-Date).ToUniversalTime()

            if ($ComputerName -eq $env:ComputerName)
            {
                $Data = Invoke-Command -ScriptBlock $ScriptBlock @Param
            }
            else
            {
                $Data = Invoke-Command -ComputerName $ComputerName -ScriptBlock $ScriptBlock @Param
            }

            if ($Data -ne $null)
            {
                $CountItem = 1
                foreach ($Item in $Data)
                {
                    $Properties = [ordered]@{'ComputerName'      = $ComputerName;
                                             'ProfileGUID'       = $Item.ProfileGUID;
                                             'ProfileName'       = $Item.ProfileName;
                                             'Description'       = $Item.Description;
                                             'Managed'           = $Item.Managed;
                                             'Category'          = $Item.Category;
                                             'NameType'          = $Item.NameType;
                                             'CategoryType'      = $Item.CategoryType;
                                             'IconType'          = $Item.IconType;
                                             'DateCreated'       = $Item.DateCreated;
                                             'DateLastConnected' = $Item.DateLastConnected;
                                             'Count'             = $CountItem;
                                             'DTCollection'      = $DTCollection }
                    $Return = New-Object -TypeName PSObject -Property $Properties

                    $Return.PSObject.TypeNames.Insert(0, 'SOE.ComputerWifiProfile.100')
                    Write-Output $Return
                    $CountItem++
                }#foreach
            }
            else
            {
                Write-Verbose "Can't query $ComputerName for Wifi profile information"
            }#Access?
        }
        else
        {
            Write-Verbose "$ComputerName is not online to query Wifi profile information"
        }#Online?
    }
}

#endregion

#region Get-SOEComputerWifiSignature

<#
.SYNOPSIS
    Collect the Wifi signature entries from registry of the selected computer.
.DESCRIPTION
    Collect the following data from the selected computer through registry and report those as SOE.ComputerWifiSignature.100 object.
    - ComputerName
    - Signature
    - DefaultGatewayMac
    - Description
    - DnsSuffix
    - FirstNetwork
    - ProfileGUID
    - Source
    - ManagedType (M = Managed, U = Unmanaged)
    - Count
    - DTCollection

    The registry will be queried on the following key on the local machine:
    - SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Signatures

    The used technique doesn't require the Server service to be active on the target computer.
.EXAMPLE
    Get-SOEComputerWifiSignature -ComputerName LocalHost

    (for this example you'll need to run PowerShell as Administrator, not required for remove access)

    Will query the local computer.
        ComputerName      : KING001
        Signature         : 010103000F0000F0080000000F0000F0BFF36F6ACB1EB7DD2659E8C8B766227C7C9F75AB5E713ADC74DA43FF05EB664B
        DefaultGatewayMac : 88:9F:FB:C1:2E:1B
        Description       : Camelot
        DnsSuffix         : Camelot.org
        FirstNetwork      : Camelot
        ProfileGUID       : {E1557FB3-5A3F-4FA2-A4D6-FE36AF21BFFE}
        Source            : 8
        ManagedType       : U
        Count             : 1
        DTCollection      : 2016-06-02 09:01:16
.OUTPUTS
    [SOE.ComputerWifiProfile.100]

.NOTES
    --- Version history:
    Version 1.10 (2016-10-31, Kees Hiemstra)
    - Added parameter Credential.
    Version 1.00 (2016-06-02, Kees Hiemstra)
    - Initial version.

    --- Extra information
    Use the following SQL statement to create a table in the database if the data needs to be stored with Save-SOEObject.

    CREATE TABLE dbo.ComputerWifiSignature(
	    [ID] [int] IDENTITY(1,1) NOT NULL PRIMARY KEY,
	    [ComputerName] [varchar](25) NOT NULL,
	    [Signature] [char](96) NOT NULL,
	    [ProfileGUID] [char](38) NOT NULL,
	    [Description] [varchar](255) NULL,
	    [DnsSuffix] [varchar](64) NULL,
	    [FirstNetwork] [varchar](255) NULL,
	    [Source] [int] NULL,
	    [ManagedType] [char](1) NULL,
	    [DTCollection] [datetime] NOT NULL,
	    [DTCreation] [datetime] NULL CONSTRAINT [DF_ComputerWifiSignature_DTCreation] DEFAULT (GETDATE()),
	    [DTMutation] [datetime] NULL,
	    [DTDeletion] [datetime] NULL
      )

.LINK
    Get-SOEComputerModel

.LINK
    Get-SOEComputerOS

.LINK
	Get-SOEComputerBSoDEvents

.LINK
	Get-SOEComputerNICDetails

.LINK
	Get-SOEComputerSignedDrivers

.LINK
    Get-Get-SOEComputerRunKeys

.LINK
	Get-SOEComputerPatches

.LINK
    Get-SOEComputerBIOS

.LINK
    Get-SOEComputerUserFirewallRules

.LINK
    Get-SOEComputerPolicyFirewallRules

.LINK
    Get-SOEComputerWifiProfile

.LINK
    Save-SOEObject

#>
function Get-SOEComputerWifiSignature
{
    [CmdletBinding(SupportsShouldProcess=$true, 
                   PositionalBinding=$true,
                   HelpUri='http://www.xs4all.nl/~chi/ps/oh1623401.html',
                   ConfirmImpact='Low')]
    Param
    (
        #Computer name to query
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0)]
        [Alias('Name')]
        [string]
        $ComputerName,

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

        $ScriptBlock = {param()
            $AllWifiSignatures = Get-ChildItem -Path 'HKLM:SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Signatures\Managed' -ErrorAction SilentlyContinue
            foreach ($WifiSignature in $AllWifiSignatures)
            {
                $Items = Get-Item -Path $WifiSignature.Name.Replace('HKEY_LOCAL_MACHINE', 'HKLM:') -ErrorAction SilentlyContinue
                foreach ($Name in $Items)
                {
                    $Properties = @{'Signature'          = $WifiSignature.PSChildName;
                                    'DefaultGatewayMac' = ($Name.GetValue("DefaultGatewayMac") | foreach { $_.ToString("X2") }) -join ":"
                                    'Description'       = $Name.GetValue("Description")
                                    'DnsSuffix'         = $Name.GetValue("DnsSuffix")
                                    'FirstNetwork'      = $Name.GetValue("FirstNetwork")
                                    'ProfileGuid'       = $Name.GetValue("ProfileGuid")
                                    'Source'            = $Name.GetValue("Source")
                                    'Type'              = 'M'}
                    $Return = New-Object -TypeName PSObject -Property $Properties
                    Write-Output $Return
                }
            }#foreach signature

            $AllWifiSignatures = Get-ChildItem -Path 'HKLM:SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Signatures\UnManaged' -ErrorAction SilentlyContinue
            foreach ($WifiSignature in $AllWifiSignatures)
            {
                $Items = Get-Item -Path $WifiSignature.Name.Replace('HKEY_LOCAL_MACHINE', 'HKLM:') -ErrorAction SilentlyContinue
                foreach ($Name in $Items)
                {
                    $Properties = @{'Signature'         = $WifiSignature.PSChildName;
                                    'DefaultGatewayMac' = ($Name.GetValue("DefaultGatewayMac") | foreach { $_.ToString("X2") }) -join ":"
                                    'Description'       = $Name.GetValue("Description")
                                    'DnsSuffix'         = $Name.GetValue("DnsSuffix")
                                    'FirstNetwork'      = $Name.GetValue("FirstNetwork")
                                    'ProfileGuid'       = $Name.GetValue("ProfileGuid")
                                    'Source'            = $Name.GetValue("Source")
                                    'Type'              = 'U'}
                    $Return = New-Object -TypeName PSObject -Property $Properties
                    Write-Output $Return
                }
            }#foreach signature
        }#ScriptBlock
    }
    Process
    {
        if ($ComputerName -in ('.', 'LocalHost'))
        {
            Write-Verbose "Change name to $($env:ComputerName)"
            $ComputerName = $env:ComputerName
        }

        if ((Test-Connection -ComputerName $ComputerName -Quiet -Count 1))
        {
            Write-Verbose "Querying $ComputerName for Wifi signature information"
            $DTCollection = (Get-Date).ToUniversalTime()

            if ($ComputerName -eq $env:ComputerName)
            {
                $Data = Invoke-Command -ScriptBlock $ScriptBlock @Param
            }
            else
            {
                $Data = Invoke-Command -ComputerName $ComputerName -ScriptBlock $ScriptBlock @Param
            }

            if ($Data -ne $null)
            {
                $CountItem = 1
                foreach ($Item in $Data)
                {
                    $Properties = [ordered]@{'ComputerName'      = $ComputerName;
                                             'Signature'         = $Item.Signature;
                                             'DefaultGatewayMac' = $Item.DefaultGatewayMac;
                                             'Description'       = $Item.Description;
                                             'DnsSuffix'         = $Item.DnsSuffix;
                                             'FirstNetwork'      = $Item.FirstNetwork;
                                             'ProfileGUID'       = $Item.ProfileGUID;
                                             'Source'            = $Item.Source;
                                             'ManagedType'       = $Item.Type;
                                             'Count'             = $CountItem;
                                             'DTCollection'      = $DTCollection }
                    $Return = New-Object -TypeName PSObject -Property $Properties

                    $Return.PSObject.TypeNames.Insert(0, 'SOE.ComputerWifiSignature.100')
                    Write-Output $Return
                    $CountItem++
                }#foreach
            }
            else
            {
                Write-Verbose "Can't query $ComputerName for Wifi signature information"
            }#Access?
        }
        else
        {
            Write-Verbose "$ComputerName is not online to query Wifi signature information"
        }#Online?
    }
}

#endregion

#region Start-SOEServerService

<#
.SYNOPSIS
    Start the server service at the selected computer.
.DESCRIPTION
    The server services is switched off by default through a GPO on the SOE environment. With this cmdlet, the server service can be switched back on temporarely (until the GPO turns it off again, or the computer is rebooted).
.PARAMETER ComputerName
    One or more computer names where the server service need to be started.
.EXAMPLE
    Start-SOEServerService -ComputerName King001

    This will start the server service on the computer King001.

.NOTES
    --- Version history:
    Version 1.10 (2016-10-31, Kees Hiemstra)
    - Corrected output and capture errors during WMI.
    Version 1.01 (2015-12-20, Kees Hiemstra)
    - Updated help.
    Version 1.00 (2014-09-09, Kees Hiemstra)
    - Initial version.
#>
function Start-SOEServerService
{
    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   HelpUri='http://www.xs4all.nl/~chi/ps/oh1437201.html',
                   Position=0)]
        [string[]]
        $ComputerName
    )

    Process
    {
        $Output = @()

        foreach ($__ComputerName in $ComputerName)
        {
            if ($__ComputerName -ne '')
            {
                Write-Verbose "Computer name: $__ComputerName"
                $__ComputerName = $__ComputerName.Trim()
                $Result = New-Object PSObject -Property ([ordered]@{
                    'ComputerName'  = $__ComputerName;
                    'OnlineStatus'  = 'Unknown';
                    'ServerService' = 'n/a';
                    'RemRegService' = 'n/a';
                    'Completed'     = $false})

                if(Test-Connection -ComputerName $__ComputerName -Count 1 -Quiet)
                {
                    $Result.OnlineStatus = 'Online'
                    try
                    {
                        $Result.ServerService = (Get-WmiObject -computer $__ComputerName Win32_Service -Filter "Name='LanmanServer'" -ErrorAction Stop).State
                        $Result.RemRegService = (Get-WmiObject -computer $__ComputerName Win32_Service -Filter "Name='RemoteRegistry'" -ErrorAction Stop).State

                        if ($Result.ServerService -ne 'Running')
                        {
                            $R = (Get-WmiObject -computer $__ComputerName Win32_Service -Filter "Name='LanmanServer'" -ErrorAction Stop).invokeMethod('ChangeStartMode', "Manual")
                            $R = (Get-WmiObject -computer $__ComputerName Win32_Service -Filter "Name='LanmanServer'" -ErrorAction Stop).invokeMethod('StartService', $null)
                        }

                        if ($Result.RemRegService -ne "Running")
                        {
                            #Start remote registry
                            $R = (Get-WmiObject -computer $__ComputerName Win32_Service -Filter "Name='RemoteRegistry'" -ErrorAction Stop).invokeMethod('ChangeStartMode', "Manual")
                            $R = (Get-WmiObject -computer $__ComputerName Win32_Service -Filter "Name='RemoteRegistry'" -ErrorAction Stop).invokeMethod('StartService', $null)
                        }


                        $Result.ServerService = (Get-WmiObject -computer $__ComputerName Win32_Service -Filter "Name='LanmanServer'" -ErrorAction Stop).State
                        $Result.RemRegService = (Get-WmiObject -computer $__ComputerName Win32_Service -Filter "Name='RemoteRegistry'" -ErrorAction Stop).State

                        if (($Result.ServerService -eq 'Running') -and ($Result.RemRegService -eq 'Running'))
                        {
                            $Result.Completed = $true
                        }

                    }
                    catch
                    {
                        $Result.OnlineStatus = 'Error'
                    }
                }#if online
                else
                {
                    $Result.OnlineStatus = 'Offline'
                }

                Write-Output $Result
            }
            else
            {
                Write-Verbose "Empty computer name."
            }

        }#foreach
    }#Process
}

#endregion

