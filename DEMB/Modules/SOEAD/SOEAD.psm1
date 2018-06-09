#region Get-SOEADUser

<#
.Synopsis
   Selects individual AD user based on the given filter options and returns the most used fields and a simple representation of the O365 EnterprisePack license options that have been set.
.DESCRIPTION
   Long description
.EXAMPLE
   Get-SOEADUser -SAMAccountName Bernhard.Kaiser
.EXAMPLE
   Get-SOEADUser -EmployeeID "00000666"
.EXAMPLE
   Get-SOEADUser -DisplayName "Doe, John"
.EXAMPLE
   Get-SOEADUser -LastName John -FirstName Do*
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
    --- Version history:
    Version 2.20 (2017-04-16, Kees Hiemstra)
    - Added AdminCount to [ShortAnalysis].
    - Added ProtectedFromAccidentalDeletion to [ShortAnalysis].
    Version 2.12 (2016-08-28, Kees Hiemstra)
    - Padded EmployeeID with 0 to a lenght of 8 characters.
    Version 2.11 (2016-08-03, Kees Hiemstra)
    - Added [AccountExpirationDate] in the analysis.
    - Replaced [LockoutTime] with [AccountLockoutTime].
    Version 2.10 (2016-08-01, Kees Hiemstra)
    - Added [Mobile]. For MFA it needs to be in the format +31 621708596
    Version 2.00 (2016-07-25, Kees Hiemstra)
    - Added [HomePage] (wwwHomePage), which is the alternative (external) mailaddress.
    Version 1.90 (2016-07-22, Kees Hiemstra)
    - Added [EmployeeNumber] (SAP Access).
    Version 1.81 (2016-07-07, Kees Hiemstra)
    - Added analysis for DesklessPack.
    Version 1.80 (2016-06-29, Kees Hiemstra)
    - Added [ExtensionAttribut11] (HR Profile).
    Version 1.70 (2016-06-10, Kees Hiemstra)
    - Added [personalTitle]
    Version 1.60 (2016-06-02, Kees Hiemstra)
    - Added [Department]
    Version 1.50 (2016-02-28, Kees Hiemstra)
    - Removed mailbox interpretation.
    - Added OU=Managed Users to define internal user.
    Version 1.40 (2015-12-12, Kees Hiemstra)
    - Added a check on not completed user migration (AD -eq @JDECoffee.com and Azure -eq @DEMB.com)
    Version 1.31 2015-12-09, Kees Hiemstra)
    - Bug fix: Users where reported again and again when fed by the pipe.
    Version 1.30 2015-11-08, Kees Hiemstra)
    - Added AzureStatus and ShortAnalysis to the result.
    Version 1.20 2015-10-05, Kees Hiemstra)
    - Used ConvertFrom-AzureLicense cmdlet from the Azure module to translate the licenses.
    Version 1.10 2015-xx-xx, Kees Hiemstra)
    - Restructure searching and made searching on userPrincipalName easier.
    - Made two different results and show the short one when more than 1 user is found.
    Version 1.03 2015-xx-xx, Kees Hiemstra)
    - Added searching on userPrincipalName.
    Version 1.02 2015-xx-xx, Kees Hiemstra)
    - Added searching on email address and first name / last name.
    Version 1.01 2015-xx-xx, Kees Hiemstra)
    - Added searching on employeeID and displayName.
    Version 1.00 2015-xx-xx, Kees Hiemstra)
    - Initial version.
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
function Get-SOEADUser
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$true,
                  ConfirmImpact='Medium')]
    [OutputType([String])]
    Param
    (
        # Specifies the sAMAccountName of the user. If the sAMAccountName contains a @-sign, it will lookup the
        # account by the userPrincipalName
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [Alias("AccountName")] 
        [string[]]
        $SAMAccountName,

        # Specifies the userPrincipalName of the user.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 2')]
        [ValidateNotNullOrEmpty()]
        [Alias("UPN")] 
        [string[]]
        $UserPrincipalName,

        # Specifies the EmployeeID (employee number) of the user. The employee number needs to be enclosed in
        # quotes if the starts with a zero (0).
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   ParameterSetName='Parameter Set 3')]
        [ValidateNotNullOrEmpty()]
        [Alias("UserNo")] 
        [string[]]
        $EmployeeID,

        # Specifies the email address of the user.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   ParameterSetName='Parameter Set 4')]
        [ValidateNotNullOrEmpty()]
        [Alias("EMailAddress")] 
        [string[]]
        $Mail,

        # Specifies the display name of the user.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   ParameterSetName='Parameter Set 5')]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $DisplayName,

        # Specifies the first name of the user.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   ParameterSetName='Parameter Set 6')]
        [ValidateNotNullOrEmpty()]
        [Alias("givenName")] 
        [string]
        $FirstName,

        # Specifies the last name of the user.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   ParameterSetName='Parameter Set 6')]
        [ValidateNotNullOrEmpty()]
        [Alias("surName")] 
        [string]
        $LastName,

        # Specifies what data set is returned. By default it's all unless the number of users returned from
        # the search is larger then 1.
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false)]
        [ValidateSet("small", "All")]
        [string]
        $ShowDataSet,

        # Specifies the domain controller to query AD.
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false)]
        [string]
        $Server
    )

    Begin
    {
        #Initiate connection to Azure only once
        try {$Sku = Get-MsolAccountSku -ErrorAction Stop }
        catch
        {
            #Get stored credentials
            if($AzureCred -eq $null)
            {
                if (-not [string]::IsNullOrEmpty($env:AzureUser))
                {
                    $AzureCred = Get-SOECredential -UserName $env:AzureUser
                }
                else
                {
                    $AzureCred = Get-SOECredential -UseUPN
                }
            }
            try { Connect-MsolService -Credential $AzureCred } catch { throw }
            $Sku = Get-MsolAccountSku
        }

        #$Props = @{'displayName' = ""; 'title' = ""; 'LastLogonDate' = ""; 'LastBadPasswordAttempt' = ""; 'LockedOut' = ""; 'LockoutTime' = ""; 'PasswordExpired' = ""; 'PasswordLastSet' = ""; 'Manager' = ""; 'mail' = ""; 'c' = ""; 'co' = ""; 'l' = ""; 'legacyExchangeDN' = ""; 'comment' = ""; 'company' = ""; 'employeeID' = ""; 'employeeType' = ""; 'enabled' = ""; 'extensionAttribute2' = ""; 'extensionAttribute5' = ""; 'memberOf' = ""; 'extensionAttribute15' = ""; 'msRTCSIP-PrimaryUserAddress' = ""; 'msRTCSIP-PrimaryHomeServer' = ""; 'whenCreated'= ""}
        $Props = @{displayName = ""; title = ""; LastLogonDate = ""; LastBadPasswordAttempt = ""; LockedOut = ""; LockoutTime = ""; PasswordExpired = ""; PasswordLastSet = ""; Manager = ""; mail = ""; c = ""; co = ""; l = ""; legacyExchangeDN = ""; comment = ""; company = ""; employeeID = ""; employeeType = ""; enabled = ""; extensionAttribute2 = ""; extensionAttribute5 = ""; memberOf = ""; extensionAttribute15 = ""; 'msRTCSIP-PrimaryUserAddress' = ""; 'msRTCSIP-PrimaryHomeServer' = ""; whenCreated= ""}

        $ADUsers = @()

        $HashADUser = @{}
        if (-not [string]::IsNullOrEmpty($Server))
        {
            $HashADUser += @{'Server' = $Server}
        }
    }
    Process
    {
        $ADUsers = @()

        if(($SAMAccountName).Count -ne 0)
        {
            foreach($Item in $SAMAccountName)
            {
                if ($SAMAccountName -like '*@*')
                {
                    $ADUser = Get-ADUser -Filter ("userPrincipalName -like '{0}'" -f $Item.Trim()) -Properties * @HashADUser
                }
                else
                {
                    $ADUser = Get-ADUser $Item -Properties * -ErrorAction SilentlyContinue @HashADUser
                }
                if($ADUser -ne $null)
                {
                    $ADUsers += $ADUser
                }
            }
        }

        if(($UserPrincipalName).Count -ne 0)
        {
            foreach($Item in $UserPrincipalName)
            {
                $ADUser = Get-ADUser -Filter ("userPrincipalName -like '{0}'" -f $Item.Trim()) -Properties * @HashADUser
                if($ADUser -ne $null)
                {
                    $ADUsers += $ADUser
                }
            }
        }

        if(($EmployeeID).Count -ne 0)
        {
            foreach($Item in $EmployeeID)
            {
                $ADUser = Get-ADUser -Filter ("EmployeeID -like '{0}'" -f $Item.Trim().PadLeft(8, "0")) -Properties * @HashADUser
                if($ADUser -ne $null)
                {
                    $ADUsers += $ADUser
                }
            }

        }

        if(($Mail).Count -ne 0)
        {
            foreach($Item in $Mail)
            {
                $ADUser = Get-ADUser -Filter ("Mail -like '{0}'" -f $Item.Trim()) -Properties * @HashADUser
                if($ADUser -ne $null)
                {
                    $ADUsers += $ADUser
                }
            }

        }

        if(![string]::IsNullOrWhiteSpace($FirstName))
        {
            $ADUser = Get-ADUser -Filter ("givenName -like '{0}' -and surName -like '{1}'" -f $FirstName, $LastName) -Properties * @HashADUser
            if($ADUser -ne $null)
            {
                $ADUsers += $ADUser
            }
        }

        if(($DisplayName).Count -ne 0)
        {
            foreach($Item in $DisplayName)
            {
                $ADUser = Get-ADUser -Filter ("DisplayName -like '{0}'" -f $Item.Trim()) -Properties * @HashADUser
                if($ADUser -ne $null)
                {
                    $ADUsers += $ADUser
                }
            }

        }

        if(($ADUsers).Count -ne 0)
        {
            foreach($ADUser in $ADUsers)
            {
                $Groups = $ADUser.memberOf -replace "CN=" -replace ",.*com$" -join "; "
                $Analysis = @()

                if ( -not [string]::IsNullOrEmpty($ADUser.AdminCount) )
                {
                    $Analysis += "AD Admin privs"
                }

                if ( $ADUser.ProtectedFromAccidentalDeletion )
                {
                    $Analysis += "Protected"
                }

                $Line = ""
				if (-not $ADUser.Enabled)
				{
					$Line = "Disabled"
				}
				
				if ($ADUser.employeeID -ne $null)
				{
					if ($ADUser.comment -eq $null)
					{
						$Line = ("$Line Inactive").Trim()
					}
					else
					{
						$Line = ("$Line $($ADUser.comment)").Trim()
					}

					if ($ADUser.employeeType -in ('Active Employee', 'Expats/ Inpats', 'Retiree/ Pensioner') -and $ADUser.extensionAttribute5 -eq 'Employee' -and $ADUser.DistinguishedName -like '*,OU=Managed Users,DC=corp,DC=demb,DC=com')
					{
						$Line = "$Line internal SAP HR/IDM user"
					}
					else
					{
						$Line = "$Line external IDM user"
					}
				}
				else
				{
					$Line = ("$Line Non IDM user").Trim()
				}
				$Analysis += $Line
				
				if ($ADUser.userPrincipalName -notlike "*@JDECoffee.com" -and $ADUser.userPrincipalName -notlike "*@demb.com")
				{
					$Analysis += 'Invalid userPrincipalName'
				}

                if ([string]::IsNullOrEmpty($ADUser.mail))
                {
                    $Analysis += "No mailbox active"
                }
                elseif ($ADUser.mail -notlike '*@JDEcoffee.com')
                {
                    $Analysis += "Mail address not JDE"
                }

                if (-not [string]::IsNullOrEmpty($ADUser.mail) -and $ADUser.msExchHideFromAddressLists)
                {
                    $Analysis += 'Hidden mailbox'
                }

				
				if ($ADUser."msRTCSIP-PrimaryUserAddress" -eq $null)
				{
					$Analysis += 'Lync not enabled'
				}

                If ( $ADuser.AccountExpirationDate -ne $null )
                {
                    $Analysis += "Account expiration at $($ADuser.AccountExpirationDate.ToString('yyyy-MM-dd'))"
                }

                $EntPackOptions = ""
                $BlockCredential = ""
                $AzureStatus = "<Not found or no access>"

                $MSOLUser = Get-AzureLicense -Identity $ADUser.userPrincipalName -ReturnResult Object -NoADGroupCheck

                if ($MSOLUser.AzureStatus -eq '<no Azure user>')
                {
                    $MSOLUser = Get-AzureLicense -Identity ($ADUser.userPrincipalName -replace "@JDEcoffee.com", "@demb.com") -ReturnResult Object -NoADGroupCheck
                    if ($MSOLUser.AzureStatus -ne '<no Azure user>')
                    {
                        $MSOLUser.AzureStatus = "<User migration not completed>"
                        $Analysis += "UserPrincipalName migration from demb to JDEcoffee is not completed yet"
                    }
                }

                $AzureStatus = $MSOLUser.AzureStatus
                $EntPackOptions = $MSOLUser.GranularNotation
                $BlockCredential = $MSOLUser.blockCredential

                if ($MSOLUser.LicensesText -match 'ENTERPRISEPACK|DESKLESSPACK' -and $Groups -notlike '*Allow-O365-Access*' -and $Groups -notlike '*Allow-External-O365-Access*' -and $Groups -notlike '*Allow-ADFS-Access*')
                {
                    $Analysis += "Not a member of Azure access groups"
                }

                if ($Groups -like '*GPO-U-Avecto-Baseline*')
                {
                    $Analysis += "Avecto: Baseline user"
                }

                if ($Groups -like '*GPO-U-Avecto-Field Engineer*')
                {
                    $Analysis += "Avecto: Field Engineer user"
                }

                if ($Groups -like '*GPO-U-Avecto-GIS*')
                {
                    $Analysis += "Avecto: GIS user"
                }

                if ($Groups -like '*GPO-U-Avecto-Software Tester*')
                {
                    $Analysis += "Avecto: Software Tester user"
                }

                if ((($ADUsers).Count -eq 1 -or $ShowDataSet -eq "All") -and $ShowDataSet -ne "Small")
                {
                    Write-Output $ADUser | 
                        Select-Object sAMAccountName,
                            WhenCreated,
                            WhenChanged,
                            Manager,
                            Company,
                            DepartmentNumber,
                            DisplayName,
                            Description,
                            Department,
                            Title,
                            #Mobile,
                            homeDirectory,
                            scriptPath,
                            lastLogonDate,
                            lastBadPasswordAttempt,
                            enabled,
                            lockedOut,
                            AccountLockoutTime,
                            passwordExpired,
                            passwordLastSet,
                            personalTitle,
                            surName,
                            givenName,
                            mail,
                            userPrincipalName,
                            c,
                            co,
                            l,
                            msExchHideFromAddressLists,
                            msRTCSIP-PrimaryUserAddress,
                            HomePage,
                            @{n='SAP access'; e={ $_.employeeNumber }},
                            employeeID,
                            employeeType,
                            comment,
                            extensionAttribute5,
                            extensionAttribute2,
                            extensionAttribute11,
                            @{n='AzureStatus'; e={ $AzureStatus }},
                            @{n='AzureLicense'; e={ $EntPackOptions }},
                            @{n='BlockCredential'; e={ $BlockCredential }},
                            distinguishedName,
                            extensionAttribute10,
                            extensionAttribute14,
                            extensionAttribute15,
                            @{n='GroupsNames'; e={ $Groups }},
							@{n='ShortAnalysis'; e={ [array]$Analysis }}
                }
                else
                {
                    Write-Output $ADUser | 
                        Select-Object sAMAccountName,
                            whenCreated,
                            whenChanged,
                            displayName,
                            lastLogonDate,
                            enabled,
                            mail,
                            userPrincipalName,
                            c,
                            msRTCSIP-PrimaryUserAddress,
                            employeeID,
                            employeeType,
                            comment,
                            extensionAttribute5,
                            extensionAttribute2,
                            @{n='AzureStatus'; e={ $AzureStatus }},
                            @{n='AzureLicense'; e={ $EntPackOptions }},
                            @{n='BlockCredential'; e={ $BlockCredential }},
                            distinguishedName,
                            extensionAttribute15,
							@{n='ShortAnalysis'; e={ [array]$Analysis }}
                }
            }
        }
    }
    End
    {
    }
}

#endregion

#region Get-SOEADComputer

<#
.Synopsis
   Selects individual AD computer based on the given filter options and returns the most used fields.
.DESCRIPTION
   Selects the following fields from the computer account:
   - ComputerName
   - LastLogonDate
   - Street (Logged on users)
   - Url (administrators accounts)
   - WhenCreated
   - WhenChanged
   - (stripped) groups
   - Short analysis
.EXAMPLE
   Get-SOEADComputer -ComputerName King001

   ComputerName      : KING001
   WhenCreated       : 2015-07-15 12:05:18
   WhenChanged       : 2016-06-02 11:21:30
   LastLogonDate     : 2016-05-26 13:06:14
   Enabled           : True
   LockedOut         : False
   LockoutTime       : 
   PasswordExpired   : False
   PasswordLastSet   : 2016-06-02 11:11:32
   SystemKey         : F6E7B034A47D8F7D54CAE41C8763011C
   DNSHostName       : KING001.camelot.org
   IPv4Address       : 10.0.1.1
   OS                : Windows 7 Enterprise Service Pack 1
   OSVersion         : 6.1 (7601)
   SerialNo          : 9CG518194F
   Description       : 
   DistinguishedName : CN=KING001,OU=Laptops,OU=Managed Computers,DC=camelot,DC=org
   Users             : king.Arthur; merlin; .\help
   GroupsNames       : Round Table; Kingdom
   ShortAnalysis     : 

.EXAMPLE
   Get-SOEADComputer -AccountName King.Arthur

.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
    --- Version history:
    Version 1.40 (2017-04-16, Kees Hiemstra)
    - Added AdminCount to [ShortAnalysis].
    - Added ProtectedFromAccidentalDeletion to [ShortAnalysis].
    Version 1.31 (2016-11-01, Kees Hiemstra)
    - Bug fix: Url attribute was not visable.
    Version 1.30 (2016-06-16, Kees Hiemstra)
    - Added the attributes c, l, type
    - Bug fix: Did not accept Get-ADComputer from the pipeline.
    Version 1.20 (2016-06-16, Kees Hiemstra)
    - Added SCCM Distribution Point to the analysis.
    Version 1.10 (2016-06-14, Kees Hiemstra)
    - Added parameter -SerialNo.
    Version 1.00 (2016-06-02, Kees Hiemstra)
    - Initial version.
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
function Get-SOEADComputer
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$true,
                  ConfirmImpact='Medium')]
    [OutputType([String])]
    Param
    (
        # Specifies the computer name, wildcards allowed
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [Alias('Name')] 
        [string[]]
        $ComputerName,

        # Specifies the computer name, wildcards allowed
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set for pipeline')]
        [ValidateNotNullOrEmpty()]
        [string]
        $DistinguishedName,

        # Specifies the user account name, wildcards allowed
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 2')]
        [ValidateNotNullOrEmpty()]
        [Alias('SAMAccountName')] 
        [string[]]
        $AccountName,

        # Specifies the serial number of the computer, wildcards allowed
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 3')]
        [ValidateNotNullOrEmpty()]
        [Alias('SerialNumber')] 
        [string[]]
        $SerialNo,

        # Specifies the domain controller to query AD.
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false)]
        [string]
        $Server
    )

    Begin
    {
        $ADComputers = @()
        $ResultCount = 0

        $HashADComputer = @{}
        if (-not [string]::IsNullOrEmpty($Server))
        {
            $HashADComputer += @{'Server' = $Server}
        }
    }
    Process
    {
        $ADComputers = @()

        if( -not [string]::IsNullOrEmpty($DistinguishedName) )
        {
            Write-Verbose "Looking for $DistinguishedName"
            $ADComputerResult = Get-ADComputer -Filter "DistinguishedName -eq '$DistinguishedName'" -Properties * -ErrorAction SilentlyContinue @HashADComputer
            foreach ( $ADComputer in $ADComputerResult )
            {
                $ADComputers += $ADComputer
            }
        }#DistinguishedName

        if( ($ComputerName).Count -ne 0 )
        {
            foreach( $Item in $ComputerName )
            {
                Write-Verbose "Looking for $Item"
                $ADComputerResult = Get-ADComputer -Filter "Name -like '$Item' -or DistinguishedName -eq '$Item'" -Properties * -ErrorAction SilentlyContinue @HashADComputer
                foreach ( $ADComputer in $ADComputerResult )
                {
                    $ADComputers += $ADComputer
                }
            }
        }#ComputerName

        if( ($AccountName).Count -ne 0 )
        {
            foreach( $Item in $AccountName )
            {
                Write-Verbose "Looking for $Item"
                $ADComputerResult = Get-ADComputer -Filter "Street -like '*$Item*'" -Properties * -ErrorAction SilentlyContinue @HashADComputer
                foreach ( $ADComputer in $ADComputerResult )
                {
                    $ADComputers += $ADComputer
                }
            }
        }#Account name

        if( ($SerialNo).Count -ne 0 )
        {
            foreach( $Item in $SerialNo )
            {
                Write-Verbose "Looking for $Item"
                $ADComputerResult = Get-ADComputer -Filter "SerialNumber -like '$Item'" -Properties * -ErrorAction SilentlyContinue @HashADComputer
                foreach ($ADComputer in $ADComputerResult )
                {
                    if($ADComputer.Name -notin $ADComputers.Name)
                    {
                        $ADComputers += $ADComputer
                    }
                }

                $ADComputerResult = Get-ADComputer -Filter "Name -like '*$Item*'" -Properties * -ErrorAction SilentlyContinue @HashADComputer
                foreach ( $ADComputer in $ADComputerResult )
                {
                    if( $ADComputer.Name -notin $ADComputers.Name )
                    {
                        $ADComputers += $ADComputer
                    }
                }
            }#foreach item
        }#Account name

        if( ($ADComputers).Count -ne 0 )
        {
            foreach( $ADComputer in $ADComputers )
            {
                $ResultCount++

                $Groups = $ADComputer.memberOf -replace "CN=" -replace ",.*com$" -join "; "
                $Analysis = @()

                if ( -not [string]::IsNullOrEmpty($ADComputer.AdminCount) )
                {
                    $Analysis += "AD Admin privs"
                }

                if ( $ADComputer.ProtectedFromAccidentalDeletion )
                {
                    $Analysis += "Protected"
                }

                $Line = ""
				if ( -not $ADComputer.Enabled )
				{
					$Line = "Disabled "
				}

                if ( $ADComputer.Name -like 'DEMB*DP*' )
                {
                    if ( $Groups -like '*HP-SCCM-Servers*' )
                    {
                        $Line += "SCCM "
                    }
                    else
                    {
                        $Line += "Obsolete "
                    }
                    $Line += "Distribution Point "
                }
                else
                {
                    if ( $Groups -like '*HP-SCCM-Servers*' )
                    {
                        $Line += "SCCM Server "
                    }
                    if ( $ADComputer.Type -eq 'DP Server' )
                    {
                        $Line += " & Distribution Point "
                    }
                    elseif ( $ADComputer.Type -like 'DP*' )
                    {
                        $Line += "Distribution Point "
                    }

                }

                if ( $ADComputer.Type -in @('DP', 'DP Onsite', 'DP Full', 'DP OiaB', 'DP Server') )
                {
                    $Line += "(with SOE file shares)"
                }

                if ( -not [string]::IsNullOrEmpty($Line) )
                {
                    $Analysis += $Line.Trim()
                }
				
                if ( $Groups -like '*GPO-C-Avecto*' -or $Groups -like '*Avecto Defendpoint Client (x64) EN JDE10*' )
                {
                    $Analysis += "Avecto"
                }

                Write-Output $ADComputer | 
                    Select-Object @{n='ComputerName'; e={ $_.SAMAccountName.Replace('$', '') }},
                        WhenCreated,
                        WhenChanged,
                        LastLogonDate,
                        Enabled,
                        LockedOut,
                        LockoutTime,
                        PasswordExpired,
                        PasswordLastSet,
                        C,
                        L,
                        Type,
                        CarLicense,
                        DNSHostName,
                        IPv4Address,
                        @{n='OS'; e={ ("$($_.OperatingSystem) $($_.OperatingSystemServicePack)").Trim() }},
                        @{n='OSVersion'; e={ $_.OperatingSystemVersion }},
                        @{n='SerialNo'; e={ $_.SerialNumber -join "; " }},
                        Description,
                        DistinguishedName,
                        @{n='Admins'; e={ $_.url -join ";" }},
                        @{n='Users'; e={ ($_.Street -split ";") -join "; " }},
                        @{n='GroupsNames'; e={ $Groups }},
						@{n='ShortAnalysis'; e={ [array]$Analysis }}
            }
        }
    }
    End
    {
        Write-Verbose -Message "Found $ResultCount result(s)"
    }
}

#endregion

#region Set-ADExtensionAttribute2

<#
.Synopsis
   Set (or clean) the ExtenstionAttribute2 (aka Extra Azure licenses) or add/remove individual tags for the selected SAMAccountAccount.
.DESCRIPTION
   Set and clean the ExtensionAttribute2 can also be done through Set-ADUser. This cmdlet also has the possibillity to add or remove an individual tags as well as clean the entrire field.
.EXAMPLE
   Set-ADExtensionAttribute2 HPWin7User1 -Action Add -Value '[Visio]'

   Adds the [Visio] tag to ExtensionAttribute if the tag does not exists in ExtensionAttribute2.
.EXAMPLE
   Set-ADExtensionAttribute2 HPWin7User1 -Action Remove -Value '[SWY]'

   Removes the [SMY] tag from ExtensionAttribute2 if the tag exists in ExtensionAttribute2.
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
    --- Version history:
    Version 1.20 (2016-12-26, Kees Hiemstra)
    - Added new Azure tags ([PBI]).
    Version 1.10 (2016-09-15, Kees Hiemstra)
    - Added Server parameter.
    Version 1.01 (2016-08-03, Kees Hiemstra)
    - Bug fix. Replaced SplitTags with Split-AzureLicenseTags.
    - Bug fix. Replaced RemovTags with Split-AzureLicenseTags.
    Version 1.00 (2016-07-05, Kees Hiemstra)
    - Initial version.
.COMPONENT
   The component this cmdlet belongs to the SOEAD module.
.ROLE
   The role this cmdlet belongs to User Account Administration.
.FUNCTIONALITY
   Maintain the extra Azure License attribute in AD.
#>
function Set-ADExtensionAttribute2
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  HelpUri = 'http://www.microsoft.com/',
                  ConfirmImpact='Medium')]
    [OutputType([String])]
    Param
    (
        # The account on which ExtensionAttribute2 needs to be set.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [Alias("Name")]
        [string] 
        $SAMAccountName,

        # Action that needs to be taken [Set, Add or Remove].
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position=1,
                   ParameterSetName='Parameter Set 1')]
        [ValidateSet('Set', 'Add', 'Remove')]
        [string]
        $Action,

        # Value that needs to be set, added or removed.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position=2,
                   ParameterSetName='Parameter Set 1')]
        [String]
        $Value,

        # Specifies the domain controller to query AD.
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false)]
        [string]
        $Server
    )

    Begin
    {
        $ADUsers = @()

        $HashADUser = @{}
        if (-not [string]::IsNullOrEmpty($Server))
        {
            $HashADUser += @{'Server' = $Server}
        }
    }
    Process
    {
        $ADUser = Get-ADUser -Identity $SAMAccountName -Properties ExtensionAttribute2, ExtensionAttribute11 @HashADUser
        $EA2 = $ADUser.ExtensionAttribute2 -join ''
        $EA11 = $ADUser.ExtensionAttribute11 -join ''
        
        if ( $Value -eq 'Full' )
        {
            $Value = '[VPN][EMS][EXC][INT][O365][RMA][RMS][SHA][SWY][WAC][YAM][PBI]'
        }

        switch ( $Action )
        {
            'Set'
            {
                $NewEA2 = $Value
            }
            'Add'
            {
                $NewEA2 = "$($EA2)$Value"
            }
            'Remove'
            {
                $NewEA2 = Remove-AzureLicenseTags -Tags $EA2 -RemoveTags (Split-AzureLicenseTags -Tags $Value)
            }
            default
            {
            }
        }

        $Result = Test-AzureLicenseTags -EA2 $NewEA2 -EA11 $EA11 -OriginalEA2 $EA2
        
        if ( $Result.HasEA2Changed )
        {
            if ( -not [string]::IsNullOrEmpty($Result.EA2LicenseTags) )
            {
                if ( $PSCmdlet.ShouldProcess($SAMAccountName, "Set ExtensionAttribut2 to '$($Result.EA2LicenseTags)'") )
                {
                    Set-ADUser $SAMAccountName -Replace @{ ExtensionAttribute2 = $Result.EA2LicenseTags } @HashADUser
                }
            }
            else
            {
                if ( $PSCmdlet.ShouldProcess($SAMAccountName, "Clear ExtensionAttribut2") )
                {
                    Set-ADUser $SAMAccountName -Clear 'ExtensionAttribute2' @HashADUser
                }
            }
        }
        else
        {
            $null = $PSCmdlet.ShouldProcess($SAMAccountName, "No change requered on ExtensionAttribut2")
        }

        Write-Output $Result
    }
    End
    {
    }
}#Set-ADExtensionAttribute2
#endregion

#region Submit-IDMBinding

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
    === Version history
    Version 1.10 (2016-12-27, Kees Hiemstra)
    - Added parameter Action with default to Report.
    Version 1.00 (2016-09-06, Kees Hiemstra)
    - Initial version.
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
function Submit-IDMBinding
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([String])]
    Param
    (
        # Specifies the sAMAccountName of the user. If the sAMAccountName contains a @-sign, it will lookup the
        # account by the userPrincipalName
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [Alias("AccountName")] 
        [string[]]
        $SAMAccountName,

        # Specifies the EmployeeID (employee number) of the user. The employee number needs to be enclosed in
        # quotes if the starts with a zero (0).
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 2')]
        [ValidateNotNullOrEmpty()]
        [Alias("UserNo")] 
        [string[]]
        $EmployeeID,

        # Action that needs to be taken [Report (bind directly), Rebind (unbind and next bind), unbind (unbind derectly)].
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false,
                   Position=1)]
        [ValidateSet('Report', 'Rebind', 'Unbind')]
        [string]
        $Action = 'Report'
    )
    Begin
    {
        
    }
    Process
    {
        if ( $PSBoundParameters.ContainsKey('SAMAccountName') )
        {
            $SearchFor = $SAMAccountName
        }
        if ( $PSBoundParameters.ContainsKey('EmployeeID') )
        {
            $SearchFor = $EmployeeID
        }

        foreach ( $Item in $SearchFor )
        {
            if ( $PSBoundParameters.ContainsKey('SAMAccountName') )
            {
                $Account = (Get-SOEADUser -SAMAccountName $Item).SAMAccountName
            }
            elseif ( $PSBoundParameters.ContainsKey('EmployeeID') )
            {
                $Item = $Item.PadLeft(8, '0')
                $Account = (Get-SOEADUser -EmployeeID $Item).SAMAccountName
            }
            if ( $Account -ne $null )
            {
                if ($pscmdlet.ShouldProcess($Account, "Sumbit binding for"))
                {
                    New-Item -Path "\\DEMBMCIS168\D`$\Scripts\Auto-UAA\Process\$Account.$Action.txt" | Out-Null
                }
            }
        }#foreach
    }
    End
    {
    }
}

#endregion

#region Set-ADManager

<#
.SYNOPSIS
    Set manager at selected account.

.DESCRIPTION
    Set the manager attribute on the select AD account of the employee. Both manager and employee can be identified by their employeeID or SAMAccountName.

.EXAMPLE
    Set-Manager -EmployeeID 20200261 -ManagerID 20200018 -Verbose
    ---
    VERBOSE: Employee: freddy.birkeli (20200261)
    VERBOSE: Manager : vegar.toverud (20200018)
    VERBOSE: Performing the operation "Set manager vegar.toverud" on target "freddy.birkeli".

.INPUTS
    Inputs to this cmdlet (if any)

.OUTPUTS
    None

.NOTES
    --- Version history
    Version 1.00 (2016-11-04, Kees Hiemstra)
    - Initial version.

.COMPONENT
    The component this cmdlet belongs to the SOEAD module.

.ROLE
    The role this cmdlet belongs to User Account Administration.

.FUNCTIONALITY
    Update employee AD account.

#>
function Set-ADManager
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([String])]
    Param
    (
        # EmployeeID of the user
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false,
                   ValueFromRemainingArguments=$false,
                   Position=0)]
        [string]
        $EmployeeID,

        # EmployeeID of the manager
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false,
                   ValueFromRemainingArguments=$false,
                   Position=1)]
        [string]
        $ManagerID,

        # SAMAccountName of the user account
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false,
                   ValueFromRemainingArguments=$false,
                   Position=3)]
        [string]
        $EmployeeSAMAccountName,

        # SAMAccountName of the manager account
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false,
                   ValueFromRemainingArguments=$false,
                   Position=4)]
        [string]
        $ManagerSAMAccountName
    )

    Begin
    {
        #region Validate parameters
        if ( -not $PSBoundParameters.ContainsKey('EmployeeID') -and -not $PSBoundParameters.ContainsKey('EmployeeSAMAccountName') )
        {
            Write-Error 'Either EmployeeID or EmployeeSAMAccountName need to be provided'
            break
        }
        elseif ( $PSBoundParameters.ContainsKey('EmployeeID') -and $PSBoundParameters.ContainsKey('EmployeeSAMAccountName') )
        {
            Write-Warning 'EmployeeID will be used and EmployeeSAMAccountName will be ignored'
        }

        if ( -not $PSBoundParameters.ContainsKey('ManagerID') -and -not $PSBoundParameters.ContainsKey('ManagerSAMAccountName') )
        {
            Write-Error "Either ManagerID or ManagerSAMAccountName need to be provided"
            break
        }
        elseif ( $PSBoundParameters.ContainsKey('ManagerID') -and $PSBoundParameters.ContainsKey('ManagerSAMAccountName') )
        {
            Write-Warning 'ManagerID will be used and ManagerSAMAccountName will be ignored'
        }
        #endregion
    }
    Process
    {
        #region Find employee
        if ( $PSBoundParameters.ContainsKey('EmployeeID') )
        {
            $EmployeeID = $EmployeeID.Trim().PadLeft(8, "0")
            $Employee = Get-ADUser -Filter ("EmployeeID -eq '{0}'" -f $EmployeeID.Trim().PadLeft(8, "0")) -Properties EmployeeID -ErrorAction Stop
            if ( $Employee.Count -ge 2 )
            {
                Write-Error "EmployeeID $EmployeeID is not unique"
                break
            }
        }
        else
        {
            $Employee = Get-ADUser $EmployeeSAMAccountName -Properties EmployeeID -ErrorAction Stop
        }

        if ( $Employee -eq $null )
        {
            Write-Error 'Employee not found'
            break
        }

        #endregion
        #region Find manager
        if ( $PSBoundParameters.ContainsKey('ManagerID') )
        {
            $ManagerID = $ManagerID.Trim().PadLeft(8, "0")
            $Manager = Get-ADUser -Filter ("EmployeeID -eq '{0}'" -f $ManagerID.Trim().PadLeft(8, "0")) -Properties EmployeeID -ErrorAction Stop
            if ( $Manager.Count -ge 2 )
            {
                Write-Error "ManagerID $ManagerID is not unique"
                break
            }
        }
        else
        {
            $Manager = Get-ADUser $ManagerSAMAccountName -Properties EmployeeID -ErrorAction Stop
        }

        if ( $Manager -eq $null )
        {
            Write-Error 'Manager not found'
            break
        }
        #endregion

        Write-Verbose "Employee: $($Employee.SAMAccountName) ($($Employee.EmployeeID))"
        Write-Verbose "Manager : $($Manager.SAMAccountName) ($($Manager.EmployeeID))"

        if ( $PSCmdlet.ShouldProcess("$($Employee.SAMAccountName)", "Set manager $($Manager.SAMAccountName)") )
        {
            Set-ADUser $Employee -Replace @{ Manager = $Manager.DistinguishedName }
        }
    }
    End
    {
    }
}

#endregion

New-Alias -Name gsu -Value Get-SOEADUser -Description "Get JDE specialized user data from AD and Azure"
New-Alias -Name gsc -Value Get-SOEADComputer -Description "Get JDE specialized computer data from AD"

Export-ModuleMember -Alias * -Function *
