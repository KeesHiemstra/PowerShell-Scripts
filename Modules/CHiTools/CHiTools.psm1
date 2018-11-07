#region Convert-ObjectToList

<#
.SYNOPSIS
    Rotate an object as a list.

.DESCRIPTION
    This function works like Format-Table, but the output is an object that can be used by for example ConvertTo-HTML.

.EXAMPLE
    Get-Service | Where-Object { $_.Name -like 'BITS' } | Convert-ObjectToList

    >>---
    Property            Value1                                 
    --------            ------                                 
    Name                BITS                                   
    RequiredServices    RpcSs; EventSystem                     
    CanPauseAndContinue False                                  
    CanShutdown         False                                  
    CanStop             True                                   
    DisplayName         Background Intelligent Transfer Service
    DependentServices                                          
    MachineName         .                                      
    ServiceName         BITS                                   
    ServicesDependedOn  RpcSs; EventSystem                     
    ServiceHandle                                              
    Status              Running                                
    ServiceType         Win32ShareProcess                      
    StartType           Automatic                              
    Site                                                       
    Container                                                  

.EXAMPLE
    Get-Service | Where-Object { $_.Name -like 'Lan*' } | Convert-ObjectToList 

    >>The first column will be the list of all properties, the following columns (Value1, Value2) will show the content.
    Property            Value1            Value2                         
    --------            ------            ------                         
    Name                LanmanServer      LanmanWorkstation              
    RequiredServices    SamSS; Srv        NSI; MRxSmb20; MRxSmb10; Bowser
    CanPauseAndContinue True              True                           
    CanShutdown         False             False                          
    CanStop             True              True                           
    DisplayName         Server            Workstation                    
    DependentServices   Browser           SessionEnv; Netlogon; Browser  
    MachineName         .                 .                              
    ServiceName         LanmanServer      LanmanWorkstation              
    ServicesDependedOn  SamSS; Srv        NSI; MRxSmb20; MRxSmb10; Bowser
    ServiceHandle                                                        
    Status              Running           Running                        
    ServiceType         Win32ShareProcess Win32ShareProcess              
    StartType           Automatic         Automatic                      
    Site                                                                 
    Container                                                            

.EXAMPLE
    Get-Service | Where-Object { $_.Name -like 'Lan*' } | Convert-ObjectToList | ConvertTo-Html

    >> Will convert the data in an HTML table where the properties are listes as rows.

.EXAMPLE
    Get-Service | Convert-ObjectToList -MaximumSize 3 -ExcludeProperty @('RequiredServices', 'DependentServices','ServicesDependedOn', 'ServiceHandle', 'Site', 'Container')

    >> Maximizes the number of columns to 4 (first column contains the property name) and exclude unwanted properties.
    Property            Value1                       Value2                 Value3                           
    --------            ------                       ------                 ------                           
    Name                AdobeARMservice              AeLookupSvc            ALG                              
    CanPauseAndContinue False                        False                  False                            
    CanShutdown         False                        False                  False                            
    CanStop             True                         False                  False                            
    DisplayName         Adobe Acrobat Update Service Application Experience Application Layer Gateway Service
    MachineName         .                            .                      .                                
    ServiceName         AdobeARMservice              AeLookupSvc            ALG                              
    Status              Running                      Stopped                Stopped                          
    ServiceType         Win32OwnProcess              Win32ShareProcess      Win32OwnProcess                  
    StartType           Automatic                    Manual                 Manual                           
.INPUTS
    [System.Management.Automation.PSObject]
.OUTPUTS
    [System.Management.Automation.PSObject]
.NOTES
    === Version history
    Version 1.00 (2017-01-06, Kees Hiemstra)
    - Initial version
.COMPONENT
    Tools
.ROLE

.FUNCTIONALITY
    Conversion
#>
function Convert-ObjectToList
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$false,
                  PositionalBinding=$false,
                  ConfirmImpact='Low')]
    [OutputType([System.Management.Automation.PSObject])]
    Param
    (
        #Object to be turned into a list
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [System.Management.Automation.PSObject[]]
        $InputObject,

        #Properties from InputObject to exclude from output
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false,
                   ValueFromRemainingArguments=$false,
                   Position=1,
                   ParameterSetName='Parameter Set 1')]
        [string[]]
        $ExcludeProperty,

        #Maximum number of InputObjects in output
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false,
                   ValueFromRemainingArguments=$false,
                   Position=2,
                   ParameterSetName='Parameter Set 1')]
        [int]
        $MaximumSize
    )

    Begin
    {
        $OutputObject = @(New-Object -TypeName PSObject)
        $Count = 0
    }
    Process
    {

        foreach ( $Object in $Input )
        {
            $Count++
            if ( $MyInvocation.BoundParameters.ContainsKey('MaximumSize') -and $Count -gt $MaximumSize ) { break }

            foreach ( $Property in $Object.PSObject.Properties )
            {
                if ( $Property.Name -in @('PropertyNames', 'PropertyCount' ) + $ExcludeProperty ) { continue }

                if ( $OutputObject.Property -notcontains $Property.Name)
                {
                    $Obj = New-Object -TypeName PSObject
                    Add-Member -InputObject $Obj -NotePropertyName 'Property' -NotePropertyValue $Property.Name -Force
                    $OutputObject += $Obj
                }
                $Obj = $OutputObject | Where-Object { $_.Property -eq $Property.Name }

                Add-Member -InputObject $Obj -NotePropertyName "Value$Count" -NotePropertyValue ($Property.Value -join '; ').Trim()
            }
        }
    }
    End
    {
        Write-Output $OutputObject
    }
}

#endregion


#region Send-ObjectAsHTMLTableMessage

<#
.SYNOPSIS
    Sends an object as HTML (table) email message.

.DESCRIPTION
    Uses the Send-MailMessage cmdlet sends an HTML email message in which the object is formatted as table from within Windows PowerShell.

.EXAMPLE
    Get-Service | Where-Object { $_.Name -like 'BITS' } | Convert-ObjectToList | Send-ObjectAsHTMLTableMessage -Subject 'BITS value(s)' -To CHi@xs4all.nl @SmtpSplatting

.EXAMPLE
    Get-Service | Where-Object { $_.Name -like 'Lan*' } | Convert-ObjectToList | Send-ObjectAsHTMLTableMessage -Subject 'Lan* value(s)' -To CHi@xs4all.nl @SmtpSplatting

.INPUTS
    [System.Management.Automation.PSObject[]]

.OUTPUTS
    None

.NOTES
    === Version history
    Version 1.02 (2018-11-06, Kees Hiemstra)
    - Bug fix: PowerShellCookbook causes an error, using Microsoft.PowerShell.Utility\Send-MailMessage forces to use the original CmdLet.
    Version 1.01 (2017-03-20, Kees Hiemstra)
    - Bug fix: With only 1 entry in the list, the subject didn't show the number.
    Version 1.00 (2017-01-06, Kees Hiemstra)
    - Initial version.

.COMPONENT

.ROLE
    Document, mail.

.FUNCTIONALITY

.LINK
    Convert-ObjectToList

.LINK
    Send-MailMessage

#>
function Send-ObjectAsHTMLTableMessage
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                   SupportsShouldProcess=$true, 
                   PositionalBinding=$false,
                   ConfirmImpact='Medium')]
    Param
    (
        #Object(s) turned into HTML table and send as HTML message.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [System.Management.Automation.PSObject[]]
        $InputObject,

        #Specifies the addresses to which the mail is sent. Enter names (optional) and the email address, such as Name <someone@example.com>. This parameter is required.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position=1,
                   ParameterSetName='Parameter Set 1')]
        [string[]]
        $To,

        #Specifies the subject of the email message. This parameter is required.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position=2,
                   ParameterSetName='Parameter Set 1')]
        [string]
        $Subject,

        #Specifies the name of the SMTP server that sends the email message.
        #
        #The default value is the value of the $PSEmailServer preference variable. If the preference variable is not set and this parameter is omitted, the command fails.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position=3,
                   ParameterSetName='Parameter Set 1')]
        [string]
        $SmtpServer,

        #Specifies the address from which the mail is sent. Enter a name (optional) and email address, such as Name <someone@example.com>. This parameter is required.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position=4,
                   ParameterSetName='Parameter Set 1')]
        [string]
        $From,

        #
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position=5,
                   ParameterSetName='Parameter Set 1')]
        [ValidateSet("Information", "Exception", "Report")]
        [string]
        $MessageType = 'Information',

        #Specifies the path and file names of files to be attached to the email message.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   ParameterSetName='Parameter Set 1')]
        [string[]]
        $Attachments,

        #Specifies the email addresses that receive a copy of the mail but are not listed as recipients of the message. Enter names (optional) and the email address, such as Name <someone@example.com>.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   ParameterSetName='Parameter Set 1')]
        [string[]]
        $Bcc,

        #Specifies the email addresses to which a carbon copy (CC) of the email message is sent. Enter names (optional) and the email address, such as Name <someone@example.com>.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   ParameterSetName='Parameter Set 1')]
        [string[]]
        $Cc,

        #Specifies a user account that has permission to perform this action. The default is the current user.
        #
        #Type a user name, such as User01 or Domain01\User01. Or, enter a PSCredential object, such as one from the Get-Credential cmdlet.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   ParameterSetName='Parameter Set 1')]
        [System.Management.Automation.PSCredential]
        $Credential,

        #Specifies the delivery notification options for the email message. You can specify multiple values. None is the default value. The alias for this parameter is dno.
        #
        #The delivery notifications are sent in an email message to the address specified in the value of the To parameter. The acceptable values for this parameter are:
        #
        #-- None. No notification.
        #-- OnSuccess. Notify if the delivery is successful.
        #-- OnFailure. Notify if the delivery is unsuccessful.
        #-- Delay. Notify if the delivery is delayed.
        #-- Never. Never notify.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   ParameterSetName='Parameter Set 1')]
        [ValidateSet('None', 'OnSuccess', 'OnFailure', 'Delay', 'Never')]
        [string]
        $DeliveryNotificationOption,

        #Specifies the encoding used for the body and subject. The acceptable values for this parameter are:
        #
        #-- ASCII
        #-- UTF8
        #-- UTF7
        #-- UTF32
        #-- Unicode
        #-- BigEndianUnicode
        #-- Default
        #-- OEM
        #
        #ASCII is the default.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   ParameterSetName='Parameter Set 1')]
        [ValidateSet('ASCII', 'UTF8', 'UTF7', 'UTF32', 'Unicode', 'BigEndianUnicode', 'Default', 'OEM')]
        [string]
        $Encoding,

        #Specifies an alternate port on the SMTP server. The default value is 25, which is the default SMTP port. This parameter is available in Windows PowerShell 3.0 and newer releases.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   ParameterSetName='Parameter Set 1')]
        [int32]
        $Port,

        #Specifies the priority of the email message. The acceptable values for this parameter are:
        #
        #-- Normal
        #-- High
        #-- Low
        #
        #Normal is the default.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   ParameterSetName='Parameter Set 1')]
        [ValidateSet('Normal', 'High', 'Low')]
        [string]
        $Priority,

        #Uses the Secure Sockets Layer (SSL) protocol to establish a connection to the remote computer to send mail. By default, SSL is not used.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   ParameterSetName='Parameter Set 1')]
        [switch]
        $UseSsl
    )

    Begin
    {
        $Splating = @{}
        $List = @()

        $Header = ''
    }
    Process
    {
        if ( $Splating.Count -eq 0 )
        {
            Write-Verbose "Collection all paramaters for Send-MailMessage"

            foreach ( $Item in $MyInvocation.BoundParameters.Keys )
            {
                $Value = Get-Variable -Name $Item -ValueOnly -ErrorAction SilentlyContinue

                switch ( $Item )
                {
                    'InputObject' {  }
                    'MessageType'
                    {
                        switch ( $Value )
                        {
                            #Color reference: http://www.w3schools.com/colors/colors_names.asp
                            'Information' { $Color = 'snow' }
                            'Exception'   { $Color = 'peachpuff' }
                            'Report'      { $Color = 'lightgreen' }
                            default       { $Color = 'snow' }
                        }
                    }
                    default       { $Splating += @{$Item = $Value} }
                }
            }
            $Header = "<style>BODY{background-color:$Color;}"
            $Header += "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
            $Header += "TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:thistle}"
            $Header += "TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:PaleGoldenrod}"
            $Header += "</style>"
        }

        $List += $InputObject
    }
    End
    {
        if ( $List.Count -ne 0 )
        {
            [string]$Message = $List | ConvertTo-Html -Title $Subject -Head $Header -Body "<H2>$Subject (#$([array]$List.Count))</H2>"

            if ($PSCmdlet.ShouldProcess("Send HTML message"))
            {
                # PowerShellCookbook is using Send-MailMessage to overwrite the original CmdLet
                Microsoft.PowerShell.Utility\Send-MailMessage -Body $Message -BodyAsHtml @Splating
            }
        }
    }
}

#endregion


#region Write-Speech

<#
.SYNOPSIS
    Use text to speech to read the message out loud.
.DESCRIPTION

.EXAMPLE
    Write-Speech -Message "The task is completed at $((Get-Date).ToString('hh:mm'))"

    ---
    "The taks is completed at <current time>" will be spoken out loud.
.NOTES
    === Version history
    Version 1.00 (2017-04-24, Kees Hiemstra)
    - Initial version.

#>
function Write-Speech
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                   SupportsShouldProcess=$false,
                   PositionalBinding=$true,
                   ConfirmImpact='Low')]
    [Alias()]
    [OutputType([String])]
    Param
    (
        # Message to speak out loud
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$false,
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        $Message
    )

    Begin
    {
        Add-Type -AssemblyName System.speech
    }
    Process
    {
        $Synth = New-Object System.Speech.Synthesis.SpeechSynthesizer
        $Synth.SetOutputToDefaultAudioDevice();

        # Speak a string synchronously.
        $null = $Synth.SpeakAsync($Message)
    }
    End
    {
    }
}

#endregion


#region Get-TypeName

<#
.SYNOPSIS
    Get the TypeName(s) of the object.
.DESCRIPTION
    Get-Member shows the TypeName(s) of an object(s). If you are only looking the TypeName, you scroll often and you not will
    notice then has more types.
    Get-TypeName will return only the unique TypeNames.
.EXAMPLE
    PS> Get-ChildItem -Path 'C:\' -Force | Get-TypeName
    ---
    System.IO.DirectoryInfo
    System.IO.FileInfo

.INPUTS
    Inputs to this cmdlet (if any)
.OUTPUTS
    [String[]]
.NOTES
    Version 1.00 (2018-11-07, Kees Hiemstra)
    - Initial version.
.COMPONENT

.ROLE

.FUNCTIONALITY

.LINK

.LINK

#>
function Get-TypeName
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                   PositionalBinding=$false,
                   ConfirmImpact='Low')]
    [Alias('GT')]
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
        $InputObject
    )

    Begin
    {
        $Result = [string[]]@()
    }
    Process
    {
        if ($InputObject -ne $null)
        {
            $TypeName = ($InputObject | Get-Member).TypeName[0]
        }
        else
        {
            $TypeName = '<null>'
        }

        if ($TypeName -notin $Result)
        {
            $Result += $TypeName
        }
    }
    End
    {
        Write-Output $Result
    }
}

#endregion


#region Test-IsElevated

<#
.SYNOPSIS
    Test if the current process runs under elevated privileges.

.DESCRIPTION
    Some tasks need to be performed whilst the process runs under elevated privileges. This cmdlet returns true if the current process is running under elevated privileges.

.EXAMPLE
    PS> Test-IsElevated
    ---
    >> When the process is started normally, the cmdlet will return false.
    False

    >> When the process is running with elevated privileges, the cmdlet will return true.
    True

.EXAMPLE
    Test-IsElevated -Verbose

    >> When the process is started normally, the cmdlet will return false.
    VERBOSE: The process is running under normal privileges
    False

    >> When the process is running with elevated privileges, the cmdlet will return true.
    VERBOSE: The process is running under elevated privileges
    True

.INPUTS
    There are no inputs.

.OUTPUTS
    [bool]

.NOTES
    --- Version history
    Version 1.00 (2016-01-31, Kees Hiemstra)
    - Inital version.

.LINK

.LINK

#>
function Test-IsElevated
{
    [CmdletBinding(SupportsShouldProcess=$false, 
                   PositionalBinding=$false,
                   ConfirmImpact='Low')]
    [OutputType([bool])]
    Param
    (
    )

    Process
    {
        $WindowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        if ($WindowsIdentity.Owner -ne $WindowsIdentity.User)
        {
            Write-Verbose "The process is running under elevated privileges"
            Write-Output $true
        }
        else
        {
            Write-Verbose "The process is running under normal privileges"
            Write-Output $false
        }
    }
}

#endregion


#region Get-ShortString

<#
.SYNOPSIS
    Shorten string to selected length and replace the last 3 characters with ... if the string is longer.

.DESCRIPTION
    This function shortens the length of the input to given length.  will be returned if the input is null or an empty string, unless the parameter -NoEmpty is given.

    The last 3 characters will be replaced by ... if the string is longer than the given length.

    The length is by default 36.

.EXAMPLE
    PS> Get-ShortString -String 'This is a long string' -Length 21
    ---
    This is a long string

.EXAMPLE
    PS> Get-ShortString -String 'This is a longer string' -Length 21
    ---
    This is a longer s...

.EXAMPLE
    PS> Get-ShortString -String ''
    ---
    <empty>

.EXAMPLE
    PS> Get-ShortString -String '' -EmptyText '...'
    ---
    ...

.INPUTS
    [string[]]

.OUTPUTS
    [string]

.NOTES
   --- Version history
   Version 1.00 (2016-09-01, Kees Hiemstra)
   - Initial version.

#>
function Get-ShortString
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                   SupportsShouldProcess=$false,
                   PositionalBinding=$true,
                   ConfirmImpact='Low')]
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
        [string[]]
        $String,

        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false,
                   ValueFromRemainingArguments=$false,
                   Position=1,
                   ParameterSetName='Parameter Set 1')]
        [int]
        $Length = 36,

        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false,
                   ValueFromRemainingArguments=$false,
                   Position=2,
                   ParameterSetName='Parameter Set 1')]
        [string]
        $EmptyText = '<empty>',

        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=3,
                   ParameterSetName='Parameter Set 1')]
        [switch]
        $NoEmpty
    )

    Process
    {
        foreach ( $Item in $String )
        {
            $Item = $Item.TrimEnd()

            if ( [string]::IsNullOrWhiteSpace($Item) )
            {
                if ( $NoEmpty.IsPresent )
                {
                    $Output = [string]::Empty
                }
                else
                {
                    $Output = $EmptyText
                }
            }
            elseif ( $Item.Length -le $Length )
            {
                $Output = $Item
            }
            else
            {
                $Output = "$($Item.Substring(0, $Length - 3))..."
            }

            Write-Output $Output
        }#foreach
    }
}

#endregion

