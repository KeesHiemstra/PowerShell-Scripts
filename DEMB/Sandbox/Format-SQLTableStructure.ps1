#region Format-SQLTableStructure

<#
.Synopsis
   Format the structure of the given as a CREATE TABLE command for SQL.
.DESCRIPTION
   The know fields from an object are shown as a CREATE TABLE command for SQL.

   The length of strings determine the size of the char and varchar.
   NULL and NOT NULL is determined as one or more occasions are $null.

   Unknown data type are commented out.

   The more data is put into the pipeline, the more presice the command.
.EXAMPLE
    Get-Process | Format-SQLTableStructure

    Creates a command based on the Get-Process object send through the pipeline.
.INPUTS
   Any object.
.OUTPUTS
   Text
.NOTES
    --- Version history:
    Version 1.10 (2016-06-22, Kees Hiemstra)
    - Added extra code to process data from the Select-Object pipeline.
      Get-Member Property --> Get-Member NoteProperty
      Definitions of NoteProperties where the value is null are empty --> Created a re-examine those properties.
    - Added extra data types.
      Some definitions are different if the object is stored in a variable.
    Version 1.00 (2016-06-04, Kees Hiemstra)
    - Initial version.
.COMPONENT

.ROLE

.FUNCTIONALITY

#>
function Format-SQLTableStructure
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
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
        [ValidateNotNullOrEmpty()]
        $Object
    )

    Begin
    {
        $AllFields = @()
        $Reexamine = $false
        $SampleCount = 0
    }
    Process
    {
        $SampleCount++
        Write-Verbose -Message "Sample: $SampleCount"

        if ( $AllFields.Count -eq 0 )
        {
            Write-Verbose -Message "Examening object properties"
            $AllProperties = $Object | Get-Member -MemberType Property -Force -ErrorAction SilentlyContinue
            if ( $AllProperties.Count -eq 0 )
            {
                $AllProperties = $Object | Get-Member -MemberType NoteProperty -Force -ErrorAction SilentlyContinue
            }

            $FieldCount = $AllProperties.Count

            foreach ($Property in $AllProperties)
            {
                $Props = [ordered] @{ Name = $Property.Name;
                                      DataType = ($Property.Definition -split " ")[0]
                                      MinSize = $null
                                      MaxSize = $null
                                      HasNull = $false
                                      Comment = ''
                          }
                if ( $Props.DataType -eq '' ) { $Props.HasNull = $true }

                $AllFields += (New-Object -TypeName PSObject -Property $Props)
            }#foreach

            Write-Verbose -Message "Collected $($AllFields.Count) fields"

            $Reexamine = ($AllFields | Where-Object { $_.DataType -eq ''}).Count -ne 0
        }
        elseif ( $Reexamine )
        {
            $AllProperties = $Object | Get-Member -MemberType NoteProperty -Force -ErrorAction SilentlyContinue

            foreach ( $NullField in ($AllFields | Where-Object { $_.DataType -eq '' }) )
            {
                $NullField.DataType = (($AllProperties | Where-Object { $_.Name -eq $NullField.Name }).Definition -split " ")[0]
            }

            $Reexamine = ($AllFields | Where-Object { $_.DataType -eq ''}).Count -ne 0
        }

        foreach ($Field in $AllFields)
        {
            if ( ($Object.($Field.Name)) -eq $null ) { $Field.HasNull = $true }

            if ( $Field.DataType -in @('string', 'System.string') )
            {
                $Length = $Object.($Field.Name).Length

                if ( $Field.MinSize -eq $null )
                {
                    if ( $Length -eq 0 )
                    {
                        $Field.HasNull = $true
                        $Field.Comment = 'Set as NULL because of zero size'
                    }
                    else
                    {
                        $Field.MinSize = $Length
                        $Field.MaxSize = $Lenght
                    }
                }
                elseif ( $Length -lt $Field.MinSize )
                {
                    if ( $Length -eq 0 )
                    {
                        if ( -not $Field.HasNull )
                        {
                            $Field.HasNull = $true
                            $Field.Comment = 'Set as NULL because of zero size'
                        }
                    }
                    else 
                    {
                        $Field.MinSize = $Length
                    }
                }
                if ( $Length -gt $Field.MaxSize )
                {
                    $Field.MaxSize = $Length
                }
            }#string
        }#foreach
    }
    End
    {
        $AllLines = @("CREATE TABLE tablename (")

        foreach ($Field in $AllFields)
        {
            $Line = "`t[$($Field.Name)] "

            switch -regex ($Field.DataType)
            {
                '^System.byte$'                { $Line += "byte" }
                '^System.Boolean$|^bool$'      { $Line += "bit" }
                '^System.datetime$|^datetime$' { $Line += "datetime" }
                '^System.int$|^int$'           { $Line += "int" }
                '^System.int32$'               { $Line += "int" }
                '^System.long$|^long$'         { $Line += "bigint" }
                '^^System.Int64$'              { $Line += "bigint" }
                '^System.string$|^string$'   
                {
                    if ( $Field.MinSize -lt 128 -and $Field.MinSize -eq $Field.MaxSize )
                    {
                        $Line += "char($($Field.MaxSize))"
                    }
                    else
                    {
                        $Line += "varchar($($Field.MaxSize))"
                    }
                }
                Default                        { $Line =  "--$($Line)type: $($Field.DataType) is not supported!" }
            }#swich

            if ( -not $Field.HasNull )
            {
                $Line += ' NOT'
            }

            $Line += ' NULL'

            $AllLines += $Line
        }#foreach

        for ($i = 0; $i -lt $AllFields.Count - 1; $i++) { $AllLines[$i + 1] += ',' }

        for ($i = 0; $i -lt $AllFields.Count; $i++)
        {
            if ( $AllFields[$i].DataType -eq 'string' )
            {
                if ( -not [string]::IsNullOrEmpty($AllFields[$i].Comment) )
                {
                    $AllFields[$i].Comment += ', '
                }
                $AllFields[$i].Comment += "Minimum size = $($AllFields[$i].MinSize)"
            }

            if ( -not [string]::IsNullOrEmpty($AllFields[$i].Comment) )
            {
                $AllLines[$i + 1] += " --$($AllFields[$i].Comment)"
            }#Add comment
        }#for

        $AllLines += "`t)"
        $AllLines += ""
        $AllLines += "-- Created on: $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))"
        $AllLines += "-- Number of fields: $($AllFields.Count)"
        $AllLines += "-- Sample size: $($SampleCount)"

        Write-Output $AllLines
    }#end
}

#endregion

$Props = @('AccountLockoutTime',
'c',
'co',
'codePage',
'Company',
'countryCode',
'Department',
'DepartmentNumber',
'Description',
'DisplayName',
'Division',
'EmployeeID',
'EmployeeNumber',
'Enabled',
'extensionAttribute1',
'extensionAttribute2',
'extensionAttribute3',
'extensionAttribute4',
'extensionAttribute5',
'extensionAttribute6',
'extensionAttribute8',
'extensionAttribute9',
'extensionAttribute10',
'extensionAttribute11',
'extensionAttribute12',
'extensionAttribute13',
'extensionAttribute14',
'extensionAttribute15',
'Fax',
'GivenName',
'HomeDirectory',
'HomedirRequired',
'HomeDrive',
'HomePage',
'HomePhone',
'Initials',
'l',
'LastBadPasswordAttempt',
'LastLogonDate',
'LockedOut',
'mail',
'mailNickname',
'Manager',
'MobilePhone',
'msRTCSIP-PrimaryUserAddress',
'msRTCSIP-UserEnabled',
'Name',
'ObjectCategory',
'ObjectClass',
'ObjectGUID',
'objectSid',
'Office',
'OfficePhone',
'Organization',
'OtherName',
'Pager',
'PasswordExpired',
'PasswordLastSet',
'PasswordNeverExpires',
'PasswordNotRequired',
'PersonalTitle',
'POBox',
'PostalCode',
'ProfilePath',
'ProtectedFromAccidentalDeletion',
'proxyAddresses',
'physicalDeliveryOfficeName',
'SamAccountName',
'sAMAccountType',
'ScriptPath',
'SmartcardLogonRequired',
'sn',
'State',
'StreetAddress',
'targetAddress',
'Title',
'TrustedForDelegation',
'TrustedToAuthForDelegation',
'userAccountControl',
'UserPrincipalName',
'whenChanged',
'whenCreated'
)

#$AllADUsers = Get-ADUser -Filter * -Properties $Props
<#
$AllADUsers  |
    Select-Object -wait AccountLockoutTime,
        c,
        co,
        codePage,
        Company,
        countryCode,
        Department,
        @{n='DepartmentNumber'; e={ $_.DepartmentNumber[0].ToString() }},
        Description,
        DisplayName,
        Division,
        EmployeeID,
        EmployeeNumber,
        Enabled,
        extensionAttribute1,
        extensionAttribute2,
        extensionAttribute3,
        extensionAttribute4,
        extensionAttribute5,
        extensionAttribute6,
        extensionAttribute8,
        extensionAttribute9,
        extensionAttribute10,
        extensionAttribute11,
        extensionAttribute12,
        extensionAttribute13,
        extensionAttribute14,
        extensionAttribute15,
        Fax,
        GivenName,
        HomeDirectory,
        HomedirRequired,
        HomeDrive,
        HomePage,
        HomePhone,
        Initials,
        l,
        LastBadPasswordAttempt,
        LastLogonDate,
        LockedOut,
        mail,
        mailNickname,
        Manager,
        MobilePhone,
        msRTCSIP-PrimaryUserAddress,
        msRTCSIP-UserEnabled,
        Name,
        ObjectCategory,
        ObjectClass,
        @{n='ObjectGUID'; e={ $_.ObjectGUID.ToString() }},
        @{n='objectSid'; e={ $_.objectSid.ToString() }},
        Office,
        OfficePhone,
        Organization,
        OtherName,
        Pager,
        PasswordExpired,
        PasswordLastSet,
        PasswordNeverExpires,
        PasswordNotRequired,
        PersonalTitle,
        POBox,
        PostalCode,
        ProfilePath,
        ProtectedFromAccidentalDeletion,
        proxyAddresses,
        physicalDeliveryOfficeName,
        SamAccountName,
        sAMAccountType,
        ScriptPath,
        SmartcardLogonRequired,
        sn,
        State,
        StreetAddress,
        targetAddress,
        Title,
        TrustedForDelegation,
        TrustedToAuthForDelegation,
        userAccountControl,
        UserPrincipalName,
        whenChanged,
        whenCreated |
    Format-SQLTableStructure -Verbose |
    Clip
#>

#$Processes = Get-Process
#$Processes | Select-Object ProcessName, StartTime | Format-SQLTableStructure -Verbose
#$Processes | Format-SQLTableStructure -Verbose
