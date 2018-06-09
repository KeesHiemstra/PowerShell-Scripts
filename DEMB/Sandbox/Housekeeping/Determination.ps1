#Determination object

$Properties = [ordered]@{Company          = [string]  'Unknown'
                         Source           = [string]  'Unknown'
                         UserType         = [string]  'Unknown'
                         UserStatus       = [string]  'Unknown'
                         AccountType      = [string]  'Active'
                         PasswordStatus   = [string]  'Active'
                         PasswordChange   = [timespan]0
                         CountryStatus    = [string]  'Okay'
                         MailboxStatus    = [string]  'Okay'
                         UPNStatus        = [string]  'Okay'
                         ExtraInformation = [string[]]@()
                         Housekeeping     = [string]  'No action'
                        }

$Determination = New-Object -TypeName PSObject -Property $Properties

$ADUser = Get-ADUser manpower.nl -Properties *

#region Company
switch ( $ADUser.ExtensionAttribute5 )
{
    'ASP'      { $Determination.Company = 'JDE' }
    'ATOS'     { $Determination.Company = 'ATOS' }
    'Employee' { $Determination.Company = 'JDE' }
    'HPE'      { $Determination.Company = 'HPE' }
}

if ( $ADUser.DistinguishedName -like '*,OU=Hosting,DC=corp,DC=demb,DC=com' )
{
    $Determination.Company = 'ATOS'
}

if ( $ADUser.DistinguishedName -like '*,OU=Users,OU=HP,OU=Support,DC=corp,DC=demb,DC=com' )
{
    $Determination.Company = 'HPE'
}
#endregion

#region Source
if ( $ADUser.EmployeeID -match '\d{8}' )
{
    if ( $ADUser.EmployeeID -match '88\d{6}' )
    {
        $Determination.Source = 'IDM'
    }
    else
    {
        $Determination.Source = 'SAP/HR'
    }
}
#endregion

#region UserStatus

switch ( $ADUser.Comment )
{
    'Active'    { $Determination.UserStatus = 'Active' }
    'Withdrawn' { $Determination.UserStatus = 'Withdrawn' }
}

#endregion

#region UserType
<#
    Unknown
    Internal
    External
#>

if ( $ADUser.ExtensionAttribute5 -eq 'Employee' )
{
    if ( $ADUser.EmployeeType -in ('Active Employee', 'Expats/ Inpats', 'Retiree/ Pensioner') )
    {
        $Determination.UserType = 'Internal'
        if ( $ADUser.comment -eq $null ) { $ADUser.UserStatus = 'Inactive' }
    }
    elseif ( $ADUser.EmployeeType -eq 'External Employee' )
    {
        $Determination.UserType = 'External'
        if ( $ADUser.comment -eq $null ) { $ADUser.UserStatus = 'Inactive' }
    }
    elseif ( $ADUser.EmployeeType -eq 'Inactive Employee' )
    {
        $Determination.UserType = 'Inactive'
        if ( $ADUser.comment -eq $null ) { $ADUser.UserStatus = 'Inactive' }
    }
    elseif ( $ADUser.EmployeeType -eq 'Internation. Transf.' )
    {
        $Determination.UserType = 'Tranfer'
        if ( $ADUser.comment -eq $null ) { $ADUser.UserStatus = 'Inactive' }
    }
    else
    {
        switch ( $ADUser.EmployeeType )
        {
            'Generic account' { $Determination.UserType = 'Generic' }
            'Service account' { $Determination.UserType = 'Service' }
        }
    }
}#Employee
elseif ( $ADUser.ExtensionAttribute5 -in ('ASP', 'ATOS', 'Contractor', 'External', 'HPE') )
{
    $Determination.UserType = 'External'
}
elseif ( [string]::IsNullOrEmpty($ADUser.ExtensionAttribute5) )
{
    switch ( $ADUser.EmployeeType )
    {
        'Generic account'   { $Determination.UserType = 'Generic' }
        'Service account'   { $Determination.UserType = 'Service' }
        'Mailbox'           { $Determination.UserType = 'Mailbox' }
        'Room'              { $Determination.UserType = 'Mailbox' }
        'Lync EV Test user' { $Determination.UserType = 'Test'    }
    }
}
else
{
    switch ( $ADUser.ExtensionAttribute5 )
    {
        'Generic account'   { $Determination.UserType = 'Generic' }
        'Service account'   { $Determination.UserType = 'Service' }
        'Mailbox'           { $Determination.UserType = 'Mailbox' }
        'Room'              { $Determination.UserType = 'Mailbox' }
        'Lync EV Test user' { $Determination.UserType = 'Test'    }
    }
}

#endregion

#region AccountType
<#
    Active
    Inactive (No LastLogonDate, Password expired)
    Dormant (60 days or more)
    Unused (30 - 59 days)
    Inclomplete (No mail, No Lync, No manager)
    Expired
    Disabled
#>

if ( $ADUser.LastLogonDate -eq $null )
{
    $Determination.AccountType = 'Unused'
    $Determination.ExtraInformation += "Account has never been used"
}
else
{
    $LastLogonDays = (New-TimeSpan -Start $ADUser.LastLogonDate -End (Get-Date)).Days
    if ( $LastLogonDays -ge 90 )
    {
        $Determination.AccountType = 'InActive'
        $Determination.ExtraInformation += "InActive for $LastLogonDays days"
    }
    elseif ( $LastLogonDays -ge 60  )
    {
        $Determination.AccountType = 'Dormant'
        $Determination.ExtraInformation += "InActive for $LastLogonDays days"
    }
    elseif ( $LastLogonDays -ge 30  )
    {
        $Determination.AccountType = 'Sleeping'
        $Determination.ExtraInformation += "InActive for $LastLogonDays days"
    }

    if ( $ADser.AccountExpirationDate -ne $null )
    {
        if ( $ADser.AccountExpirationDate -lt (Get-Date) )
        {
            $Expires = ((Get-Date) - $ADser.AccountExpirationDate).Day
            $Determination.AccountType = 'Expired'
            $Determination.ExtraInformation += "Account exipired $Expires days ago"
        }
        else
        {
            $Expires = ($ADser.AccountExpirationDate - (Get-Date)).Day
            $Determination.ExtraInformation += "Account exipired in $Expires days"
        }
    }
}
#endregion

#region PasswordStatus
<#
    Active
    NeverExpires
    LockedOut
    Expired
#>

if ( $ADser.PasswordNeverExpires )
{
    $Determination.PasswordStatus = 'NeverExpires'
    $Determination.ExtraInformation += 'Password never expires'
}
if ( $ADser.LockedOut )
{
    $Determination.PasswordStatus = 'LockedOut'
}
if ( $ADser.PasswordExpired )
{
    $Determination.PasswordStatus = 'Expired'
}

$Determination.PasswordChange = New-TimeSpan -Start (Get-Date) -End ($ADUser.PasswordLastSet).AddDays(90)
if ( $Determination.PasswordChange -lt 0 ) { $Determination.ExtraInformation += 'Password had to be changed already' }
#endregion

$Determination