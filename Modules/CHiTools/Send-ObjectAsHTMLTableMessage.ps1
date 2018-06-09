#region Send-ObjectAsHTMLTableMessage

<#
.SYNOPSIS
    Sends an object as HTML (table) email message.

.DESCRIPTION
    Uses the Send-MailMessage cmdlet sends an HTML email message in which the object is formatted as table from within Windows PowerShell.

.EXAMPLE

.EXAMPLE

.INPUTS
    [System.Management.Automation.PSObject[]]

.OUTPUTS
    None

.NOTES
    === Version history
    Version 1.01 (2017-03-20, Kees Hiemstra
    - Bug fix: With only 1 entry in the list, the subject didn't show the number.
    Version 1.00 (2017-01-06, Kees Hiemstra
    - Initial version.

.COMPONENT

.ROLE

.FUNCTIONALITY

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
                Send-MailMessage -Body $Message -BodyAsHtml @Splating
            }
        }
    }
}

#endregion