#region Get-SOECredential

<#
.Synopsis
   Create a creditial object for the current user or specified account.

.Description
   This cmdlet creates a creditial object ([System.Management.Automation.PSCredential]) for the current user or the specified account.
   In case the userPrincipalName needs to be used, the parameter -UseUPN can be used. The cmdlet will look up the userPrincipalName in Active Directory if the user name
   is not already a userPrincipalName.

   It also will look for for a password file. If the password file exists, it will use the content to retrieve the password. Be aware the the encryption of the file can
   only be used on the same computer, you can't move the password file to another computer.

.Example
   Get-SOECredentials -SaveToFile

   The user name will be retreived from the environment varialble "UserName" and it will look for a file in the CurrentUserAllHosts folder with that name and the extension .txt.
   The password will be retreived from this file if it exists. If it doesn't exist, it will show the credentials message box to prompt for the password and save it to
   the file afterward.

.Example
   Get-SOECredentials -UseUPN -SaveToFile

   The user name will be retreived from the environment varialble "UserName" and the userPrincipalName will be retreived from Active Direcotry. Then the cmdlet will look
   for a file in the CurrentUserAllHosts folder with that name and the extension .txt.
   The password will be retreived from this file if it exists. If it doesn't exist, it will show the credentials message box to prompt for the password and save it to
   the file afterward.

.Example
   Get-SOECredentials -UserName GateKeeper@camelot.ay -Path B:\MyPassword.txt -SaveToFile

   The password will be retreived from the file B:\MyPassword.txt if it exists. If it doesn't exist, it will show the credentials message box to prompt for the password and save it to
   the file afterward.

.Outputs
   [System.Management.Automation.PSCredential]

.Notes
   --- Version history
   Version 1.00 (2015-11-10, Kees Hiemstra)
   - Initial version.

#>
function Get-SOECredential
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$true,
                  ConfirmImpact='Medium')]
    [OutputType([System.Management.Automation.PSCredential])]
    Param
    (
        #Specifies the user name for which you need the password
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [Alias("userPrincipalName")] 
        [string]
        $UserName,

        #Specifies the path to the password file
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position=1,
                   ParameterSetName='Parameter Set 1')]
        [Alias("FileName")] 
        [string]
        $Path,

        #Specifies that the password needs to be saved to a file
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   ParameterSetName='Parameter Set 1')]
        [switch]
        $SaveToFile,

        #Specifies that the user name needs to be a userPrincipalName
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   ParameterSetName='Parameter Set 1')]
        [switch]
        $UseUPN
    )

    Begin
    {
    }
    Process
    {
        if ([string]::IsNullOrEmpty($UserName))
        {
            Write-Verbose "Use current user"

            $UserName = $env:USERNAME
            if ($UseUPN.IsPresent -and $UseUPN -and $UserName -notlike '*@*')
            {
                Write-Verbose "Get userPrincipalName from AD"
                $UserName = (Get-ADUser -Filter { sAMAccountName -eq $UserName }).userPrincipalName
            }
        }
        else
        {
            if ($UseUPN.IsPresent -and $UseUPN -and $UserName -notlike '*@*')
            {
                Write-Verbose "Get userPrincipalName from AD"
                $UserName = (Get-ADUser -Filter { sAMAccountName -eq $UserName }).userPrincipalName
            }
        }

        if ([string]::IsNullOrEmpty($Path))
        {
            $Path = "$($Profile.CurrentUserAllHosts -replace "profile.ps1$")$($UserName -replace "^.*\\").txt"
        }

        Write-Verbose "User name: $UserName"
        Write-Verbose "File name: $Path"

        if(-Not (Test-Path -Path $Path))
        {
            $Cred = Get-Credential -UserName $UserName -Message "Provide password"
        }
        else
        {
            $Cred = New-Object -Type System.Management.Automation.PSCredential -ArgumentList $UserName, (Get-Content -Path $Path | ConvertTo-SecureString)
        }


        if ($SaveToFile.IsPresent -and $SaveToFile -and $pscmdlet.ShouldProcess($UserName, "Save password to"))
        {
            $Cred.Password | ConvertFrom-SecureString | Out-File $Path
        }

        Write-Output $Cred
    }
    End
    {
    }
}

#endregion


#region Get-Password

<#
.Synopsis
Decrypt the password from the provided credential object.

.Description
Decryption works only in this way when the encryption is done on the same computer.

.Example
Get-Password -Credential $Cred

This will return the password in plain text.

.Notes
   --- Version history
   Version 1.00 (2015-11-10, Kees Hiemstra)
   - Initial version.

#>
function Get-Password
{
    [CmdletBinding()]
    [OutputType([string])]
    Param
    (
        # Credential for which the password needs to be decrypted.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false)]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        $Credential
    )

    Process
    {
        Write-Output $Credential.GetNetworkCredential().Password
    }
}

#endregion


#region New-Share

<#
.Synopsis
   Create a new share.

.Description
   Create a new share with the mentioned access right.

.Example
   New-Share -ComputerName 'KingServ' -Name "King$" -Path 'C:\Distribution\King' -Access '/GRANT:Everyone,READ'

   This will create a new folder called C:\Distribution\King on the KingServ computer
   and share this folder as King$ with Read access to everyone.

.Notes
   --- Version history
   Version 1.00 (2016-01-27, Kees Hiemstra)
   - Initial version.

#>
function New-Share
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$true,
                  ConfirmImpact='Medium')]
    [OutputType([String])]
    Param
    (
        #Computer on which the share needs to be created
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$false,
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [string]
        $ComputerName,

        #Name of the share
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$false,
                   Position=1,
                   ParameterSetName='Parameter Set 1')]
        [string]
        $Name,

        #Local path to share
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$false,
                   Position=2,
                   ParameterSetName='Parameter Set 1')]
        [string]
        $Path,

        #Provide non-default access rules
        #e.g. /GRANT:Everyone,READ
        #     /GRANT:Camelot\King Admins,FULL
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$false,
                   Position=2,
                   ParameterSetName='Parameter Set 1')]
        [string]
        $Access
    )

    Begin
    {
    }
    Process
    {
        $HashParameters = @{}

        if ($ComputerName -eq '.')
        {
            $ComputerName = $env:COMPUTERNAME
        }

        if ($ComputerName -ne $env:COMPUTERNAME)
        {
            $HashInvoke += @{'ComputerName' = $ComputerName}
        }

        if ($pscmdlet.ShouldProcess($ComputerName, "Create new share to"))
        {
            #Create local path and set basic access
            $ScriptBlock = {
                param([string]$Path)
                if (-not (Test-Path -Path "$Path")) { New-Item -Path "$Path" -ItemType Directory | Out-Null };
                Cmd /C "ICacls $Path /Grant Everyone:(OI)(CI)(RX,R,X,RD,RA,REA,RC) /inheritance:e" | Out-Null;
                }
            Invoke-Command -ArgumentList @($Path) -ScriptBlock $ScriptBlock @HashInvoke

            #Create share
            $ScriptBlock = {
                param([string]$Path, [string]$Name, [string]$Access)
                Cmd /C "Net Share $Name=$Path `"$Access`"" | Out-Null;
            }
            Invoke-Command -ArgumentList @($Path, $Name, $Access) -ScriptBlock $ScriptBlock @HashInvoke
        }
    }
    End
    {
    }
}

#endregion


#region Test-IsElevated

<#
.SYNOPSIS
Test if the current process runs under elevated privileges.

.Description
Some tasks need to be performed whilst the process runs under elevated privileges. This cmdlet returns true if the current process is running under elevated privileges.

.Example
Test-IsElevated

>> When the process is started normally, the cmdlet will return false.
False

>> When the process is running with elevated privileges, the cmdlet will return true.
True

.Example
Test-IsElevated -Verbose

>> When the process is started normally, the cmdlet will return false.
VERBOSE: The process is running under normal privileges
False

>> When the process is running with elevated privileges, the cmdlet will return true.
VERBOSE: The process is running under elevated privileges
True

.Inputs
There are no inputs.

.Outputs
[bool]

.Notes
--- Version history
Version 1.00 (2016-01-31, Kees Hiemstra)
- Inital version.
#>
function Test-IsElevated
{
    [CmdletBinding(SupportsShouldProcess=$true, 
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


#region Remote BITS Transfer

#region Get-SOEBITSTransfer

<#
.Synopsis
Get the details from a remote BITS transfer.

.DESCRIPTION
Get-SOEBITSTransfer get the BITS transfer data from the specified job on the specified computer or will
list the registered jobs on the selected computer.

.EXAMPLE
Get-SOEBITSTransfer -ComputerName Server02 -JobId "{327C1E84-ABC3-475D-A5D3-E7ED09806523}" -Crededtial $Cred
    
Get the details from Server02 and look for the specified job with the credentials stored in $Cred.

.EXAMPLE
Get-SOEBITSTransfer -ComputerName Server02 -AllUsers -Crededtial $Cred
    
Get the details from Server02 and look for the all the jobs with the credentials stored in $Cred.

.EXAMPLE
Get-SOEProcessFile -ComputerName $ComputerName | Get-SOEBITSTransfer -Credential $Cred

The process file does contain the the JobId and is retrieved by Get-SOEProcessFile and then passed through to Get-SOEBITSTransfer to get the details.

.OUTPUTS
    PSObject with the following properies:
        ComputerName
        DisplayName
        BytesTransferred
        BytesTotal
        BytesProgress
        FilesTransferred
        FilesTotal
        JobId
        JobCompletionTime
        JobCreationTime
        State [int]
        Status

        State 6 = transfer completed, job ready for completion.

.COMPONENT
    The component this cmdlet belongs to Remote BITS.

.NOTES
BITS can't be done remotely with the normal BITSTransfer cmdlets. This cmdlet is using Get-WsmanInstance and translates the text output into a object. 

.LINK
Start-SOEBITSTransfer

.LINK
Complete-SOEBITSTransfer
    
.LINK
Remove-SOEBITSTransfer

#>
function Get-SOEBITSTransfer
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$false, 
                  PositionalBinding=$true,
                  ConfirmImpact='low')]
    [OutputType([PSObject])]
    Param
    (
        # Get the BITS transfer data on the specified computer.
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $ComputerName,

        # Get the BITS transfer data from the specified job. 
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false,
                   Position = 1,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [String]
        $JobId,

        # Get the BITS transfer data from the jobs. 
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false,
                   Position = 1,
                   ParameterSetName='Parameter Set 2')]
        [ValidateNotNull()]
        [switch]
        $AllUsers,

        # Credentials with which the transfer is started.
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
    }
    Process
    {
        $Props = [ordered]@{ComputerName=$ComputerName;
                            DisplayName='';
                            BytesTransferred=0;
                            BytesTotal=0;
                            BytesProgress=0;
                            FilesTransferred=0;
                            FilesTotal=0;
                            JobId=$JobId;
                            JobCompletionTime=0;
                            JobCreationTime=0;
                            State=-1;
                            Status='Unknown'}
        $Result = New-Object -TypeName PSObject -Property $Props

        try
        {
            if($AllUsers)
            {
                $BRList = Get-WsmanInstance -ResourceURI wmi/root/microsoft/bits/BitsClientJob -Enumerate –ComputerName $ComputerName -Credential $Credential
            }
            else
            {
                $BRList = Get-WsmanInstance -ResourceURI wmi/root/microsoft/bits/BitsClientJob -SelectorSet @{JobId=$JobId} –ComputerName $ComputerName -Credential $Credential
            }

            foreach ($BR in $BRList)
            {
                $Result.DisplayName = $BR.DisplayName
                $Result.BytesTransferred = [Convert]::ToInt64($BR.BytesTransferred)
                $Result.BytesTotal = [Convert]::ToInt64($BR.BytesTotal)
                if($Result.BytesTotal -gt 0) { $Result.BytesProgress = ($Result.BytesTransferred * 100)/$Result.BytesTotal }
                $Result.FilesTransferred = $BR.FilesTransferred
                $Result.FilesTotal = $BR.FilesTotal
                $Result.JobID = $BR.JobId
                $Result.JobCompletionTime = $BR.JobCompletionTime
                $Result.JobCreationTime = $BR.JobCreationTime
                $Result.State = [Convert]::ToInt64($BR.State)
                switch ($Result.State)
                {
                    2 { $Result.Status = 'Transferring' }
                    4 { $Result.Status = 'Error transferring' }
                    5 { $Result.Status = 'Error transferring' }
                    6 { $Result.Status = 'Transferred' }
                    default { $Result.Status = 'Unknown' }
                }
                Write-Output $Result
            }
        }
        catch
        {
            $Result.Status = "Error retreiving info"
        }
        
    }
    End
    {
    }
}

#endregion

#region Start-SOEBITSTransfer

<#
.Synopsis
Start a remote BITS transfer.

.DESCRIPTION
Initiate the remote BITS transfer, submit all files and start the download.

.EXAMPLE
$Result = Start-SOEBITSTransfer -ComputerName "Server02" -Files ("file1.txt", "file2.txt", "file3.txt") -SourceHttp "http://Server01/FileSource" -TargetPath "D:\FileTarget" -Credential $Cred -DisplayName "Remote transfer"

Starts the remote BITS transfer on Server02. It will download file1.txt, file2.txt, file3.txt from Server01 and store these on Server02 on D:\FileTarget.

.OUTPUTS
System.Management.Automation.PSObject, see the output from Get-SOEBITSTransfer.

.ROLE
    The role this cmdlet belongs to remote BITS.

.NOTES
BITS can't be done remotely with the normal BITSTransfer cmdlets. This cmdlet is using Invoke-WSManAction. 

During the transfer the download is stored in a temporarely folder in Windows. The State property will change to 6 once the download is finished and the Complete-SOEBITSTransfer can be used to move the downloaded files to the destination folder.

Although it's possible to download files from and to subfolders, these subfolders must exist on the target computer. These will NOT automatically be created.

.COMPONENT
The component this cmdlet belongs to SOETools.

.ROLE
   The role this cmdlet belongs to SOE remote BITS transfer.

.FUNCTIONALITY
   The functionality that best describes this cmdlet

.LINK
Get-SOEBITSTransfer

.LINK
Complete-SOEBITSTransfer

.LINK
Remove-SOEBITSTransfer

#>
function Start-SOEBITSTransfer
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$true,
                  ConfirmImpact='low')]
    [OutputType([PSObject])]
    Param
    (
        # Start the BITS transfer on the specified computer.
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position = 0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [string]
        $ComputerName,

        # Array of file names that needs to be downloaded.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false,
                   Position = 1,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [String[]]
        $Files,

        # Source Http of the server were the files will be downloaded from.
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position = 0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [string]
        $SourceHttp,

        # Local folder were to store the files once the download is complete.
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position = 0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [string]
        $TargetPath,

        # Credentials with which the transfer can be started. Needs to be local administrator on the target computer.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        $Credential,

        # Descriptive name.
        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [String]
        $DisplayName = "SOE BITS Transfer"
    )

    Begin
    {
    }
    Process
    {
        try
        {
            Enable-WSManCredSSP -Role Client -DelegateComputer $ComputerName -Force | Out-Null
            Invoke-Command -ComputerName $ComputerName -ScriptBlock { Enable-WSManCredSSP -Role Server -Force } | Out-Null
        }
        catch
        {
            Write-Error "Can't enable WSManCredSSP"
            return
        }

        if ($PSCmdlet.ShouldProcess($ComputerName, "Submit transfer"))
        {
            #Submit the first file (the transfer job will automatically start in suspended mode)
            $Result = Invoke-WsmanAction -Action CreateJob -ResourceURI wmi/root/microsoft/bits/BitsClientJob -ValueSet @{DisplayName=$DisplayName; RemoteUrl="$SourceHttp/$($Files[0])"; LocalFile="$TargetPath\$($Files[0])"; Type=0} –ComputerName $ComputerName -Credential $Credential            $JobId = $Result.JobId
            Write-Verbose "JobID: $JobId"

            if ($JobId -eq $null)
            {
                return
            }

            #Submit the remaining files
            $Files | Where-Object {$_ -ne $Files[0]} | ForEach-Object { Invoke-WsmanAction -Action AddFile -ResourceURI wmi/root/microsoft/bits/BitsClientJob -SelectorSet @{JobId=$JobId} –Valueset @{RemoteUrl=("$SourceHttp/$_"); LocalFile=("$TargetPath\$_")} –ComputerName $ComputerName -Credential $Credential } | Out-Null

            #Start the transfer
            Invoke-WSManAction -Action SetJobState -ResourceURI wmi/root/microsoft/bits/BitsClientJob -SelectorSet @{JobId=$JobId} -ValueSet @{JobState=2} –ComputerName $ComputerName -Credential $Credential | Out-Null            Get-SOEBITSTransfer -ComputerName $ComputerName -JobId $JobId -Credential $Credential        }

        Write-Output $Result
    }
    End
    {
    }
}#endregion

#region Complete-SOEBitsTransfer

<#
.Synopsis
Complete the remote BITS transfer.

.DESCRIPTION
Completes the remote BITS transfer. If the download is not yet completed, the job will be deleted otherwise the files will be moved to the target folder.

.EXAMPLE
Complete-SOEBITSTransfer -ComputerName Server02 -JobId "{327C1E84-ABC3-475D-A5D3-E7ED09806523}" -Credential $Cred | Out-File

Complete the remote BITS transfer job with the specified JobId on the Server02.

.OUTPUTS
System.Management.Automation.PSObject, see the output from Get-SOEBITSTransfer.

.ROLE
    The role this cmdlet belongs to remote BITS.

.NOTES
BITS can't be done remotely with the normal BITSTransfer cmdlets. This cmdlet is using Invoke-WSManAction. 

.COMPONENT
The component this cmdlet belongs to SOETools.

.LINK
Get-SOEBITSTransfer

.LINK
Start-SOEBITSTransfer

.LINK
Remove-SOEBITSTransfer

#>
function Complete-SOEBitsTransfer
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$false, 
                  PositionalBinding=$false,
                  ConfirmImpact='low')]
    [OutputType([PSObject])]
    Param
    (
        # Complete the BITS transfer job on the specified computer.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [Alias("DNSHostName")]
        [string]
        $ComputerName,

        # Complete the BITS transfer job with the specified JobId.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [String]
        $JobId,

        # Credentials with which the transfer is started.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        $Credential
    )

    Begin
    {
    }
    Process
    {
        $Result = Get-SOEBITSTransfer -ComputerName $ComputerName -JobId $JobId -Credential $Credential
        
        Invoke-WSManAction -Action SetJobState -ResourceURI wmi/root/microsoft/bits/BitsClientJob -SelectorSet @{JobId=$JobId} -ValueSet @{JobState=1} –ComputerName $ComputerName -Credential $Credential

        $Result.Status = "Completed"

        Write-Output $Result
    }
    End
    {
    }
}
#endregion

#region Remove-SOEBitsTransfer

<#
.Synopsis
Remove/cancel the remote BITS transfer.

.DESCRIPTION
Removes the remote BITS transfer. The dowloaded file will be removed from the cache and the job will be deleted.

.EXAMPLE
Remove-SOEBITSTransfer -ComputerName Server02 -JobId "{327C1E84-ABC3-475D-A5D3-E7ED09806523}" -Credential $Cred | Out-File

Remove the remote BITS transfer job with the specified JobId on the Server02.

.EXAMPLE
Get-SOEBITSTransfer -ComputerName $ComputerName -AllUsers -Credential $Cred | Remove-SOEBITSTransfer -Credential $Cred | Out-File

Remove all the remote BITS transfer jobs on the Server02.

.OUTPUTS
System.Management.Automation.PSObject, see the output from Get-SOEBITSTransfer.

.ROLE
    The role this cmdlet belongs to remote BITS.

.NOTES
BITS can't be done remotely with the normal BITSTransfer cmdlets. This cmdlet is using Invoke-WSManAction. 

.COMPONENT
The component this cmdlet belongs to SOETools.

.LINK
Get-SOEBITSTransfer

.LINK
Start-SOEBITSTransfer

.LINK
Complete-SOEBITSTransfer

#>
function Remove-SOEBitsTransfer
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$false, 
                  PositionalBinding=$false,
                  ConfirmImpact='low')]
    [OutputType([PSObject])]
    Param
    (
        # Remove the BITS transfer job on the specified computer.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [Alias("DNSHostName")]
        [string]
        $ComputerName,

        # Remove the BITS transfer job with the specified JobId.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [String]
        $JobId,

        # Credentials with which the transfer is started.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        $Credential
    )

    Begin
    {
    }
    Process
    {
        $Result = Get-SOEBITSTransfer -ComputerName $ComputerName -JobId $JobId -Credential $Credential
        
        Invoke-WSManAction -Action SetJobState -ResourceURI wmi/root/microsoft/bits/BitsClientJob -SelectorSet @{JobId=$JobId} -ValueSet @{JobState=0} –ComputerName $ComputerName -Credential $Credential

        $Result.Status = "Completed"

        Write-Output $Result
    }
    End
    {
    }
}
#endregion

#endregion


#region Zip archive
Add-Type -As System.IO.Compression.FileSystem

#region New-ZipArchive
<#
.Synopsis
Create new zip file or append to existing zip file.

.Description
Add files to an existing zip file or create a new one.

This function uses the System.IO.Compression.FileSystem library in the .Net Framework

.Example
   New-ZipArchive -ZipPath "O:\Backup\Souces.zip" "C:\Src"

   Will create a zipfile in O:\Backup called Sources.zip. All files in C:\Src will be added to the archive. The root folder in the archive
   is \Src

.Notes
    Requires .Net version 4.5 or higher.

   --- Version history
   Version 1.00 (2016-04-17, Kees Hiemstra)
   - Initial version.

#>
function New-ZipArchive
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true,
                  PositionalBinding=$true,
                  ConfirmImpact='Medium')]
    [OutputType([String])]
    Param
    (
        # The path of the zip to create or append
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false,
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1',
                   HelpMessage="Name of the zip archive")]
        [ValidateNotNullOrEmpty()]
        [Alias("Path" ,"Archive", "ZipFile", "Zip")]
        [string]
        $ZipPath,

        # Files to add to the zip file
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$false, 
                   Position=1,
                   ParameterSetName='Parameter Set 1',
                   HelpMessage="File(s) or Folder to include into the zip archive")]
        [ValidateNotNullOrEmpty()]
        [Alias("PSPath","Item")]
        [string[]]
        $InputObject = $Pwd,

		# Overwrite an existing zip file, instead of appending to it
		[Switch]$Overwrite,
 
		# The compression level (defaults to Optimal):
		#   Optimal - The compression operation should be optimally compressed, even if the operation takes a longer time to complete.
		#   Fastest - The compression operation should complete as quickly as possible, even if the resulting file is not optimally compressed.
		#   NoCompression - No compression should be performed on the file.
		[System.IO.Compression.CompressionLevel]$CompressionLevel = "Optimal"
    )

    Begin
    {
		[string]$File = Split-Path -Path $ZipPath -Leaf
		[string]$Folder = if( $Folder = Split-Path -Path $ZipPath) { Resolve-Path -Path $Folder } else { $Pwd }
		$ZipPath = Join-Path -Path $Folder -ChildPath $File
        if ( $Overwrite -and (Test-Path $ZipPath) )
        {
            if ( $pscmdlet.ShouldProcess($ZipPath, "Delete existing zip archive") )
            {
                Remove-Item -Path $ZipPath
            }
        }
        if ( $pscmdlet.ShouldProcess($ZipPath, "Open zip archive for update") )
        {
            $Archive = [System.IO.Compression.ZipFile]::Open( $ZipPath, "Update" )
        }
    }
    Process
    {
        foreach ($Path in $InputObject)
        {
            foreach ($Item in (Resolve-Path -Path $Path))
            {
                Push-Location (Split-Path -Path $Item)
                foreach($File in (Get-ChildItem -Path $Item -Recurse -File -Force | ForEach-Object -MemberName FullName))
                {
                    $RelativePath = (Resolve-Path -Path $File -Relative).TrimStart(".\")
                    $null = [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile( $Archive, $File, $RelativePath, $CompressionLevel )
                    Write-Output $File
                }
                Pop-Location
            }
        }


        if ($pscmdlet.ShouldProcess("Target", "Operation"))
        {
        }
    }
    End
    {
		$Archive.Dispose()
    }
}
#endregion
#endregion


#region Get-ShortString

<#
.Synopsis
Shorten string to selected length and replace the last 3 characters with ... if the string is longer.

.Description
This function shortens the length of the input to given length.  will be returned if the input is null or an empty string, unless the parameter -NoEmpty is given.

The last 3 characters will be replaced by ... if the string is longer than the given length.

The length is by default 36.

.Example
Get-ShortString -String 'This is a long string' -Length 21

>>> This is a long string

.Example
Get-ShortString -String 'This is a longer string' -Length 21

>>> This is a longer s...

.Example
Get-ShortString -String ''

>>> <empty>

.Example
Get-ShortString -String '' -EmptyText '...'

>>> ...

.Inputs
[string[]]

.OUtputs
[string]

.Notes
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

    Begin
    {
    }
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
    End
    {
    }
}

#endregion

#Export-ModuleMember -Function Get-SOECredential
#Export-ModuleMember -Function Get-Password
#Export-ModuleMember -Function New-Share
#Export-ModuleMember -Function Test-IsElevated
#Export-ModuleMember -Function New-ZipArchive
#Export-ModuleMember -Function Get-ShortString

#Work in progress?
#Export-ModuleMember -Function Get-SOEBITSTransfer
#Export-ModuleMember -Function Start-SOEBITSTransfer
#Export-ModuleMember -Function Complete-SOEBitsTransfer
#Export-ModuleMember -Function Remove-SOEBitsTransfer
