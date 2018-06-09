Set-Variable -Name SOELogFile -Value "Niets.log" -Scope Global

#region Logging

#region Set-SOELogFile

<#
.SYNOPSIS
    Set the log file name based on the input.
.DESCRIPTION
    The global variable SOELogFile will be set and is used by Write-SOELog.
.EXAMPLE
    Set-SOELogFile -Path $MyInvocation.MyCommand.Definition

    Will set the log file to the common SOE log folder with the name of script and the extension .log.
.EXAMPLE
    Set-SOELogFile -LiteralPath "C:\Log\Process.log"

    Will set the log file to "C:\Log\Process.log".
.OUTPUTS
   None
.NOTES
   General notes
.COMPONENT
   The component this cmdlet belongs to Write-SOELog.
#>
function Set-SOELogFile
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Low')]
    [OutputType([String])]
    Param
    (
        #Name of the script where the name of the log file will be dirived from.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [Alias("ScriptPath")] 
        [string]
        $Path,

        #Full file name to the log file.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 2')]
        [ValidateNotNullOrEmpty()]
        [Alias("FullPath")] 
        [string]
        $LiteralPath,

        #Optional message to be written to the logfile.
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$false, 
                   Position=1)]
        [ValidateNotNullOrEmpty()]
        [Alias("LogMessage")] 
        [string]
        $Message
    )

    Begin
    {
    }
    Process
    {
        if(-not [string]::IsNullOrWhiteSpace($LiteralPath))
        {
            $Global:SOELogFile = $LiteralPath
        }
        elseif (-not [string]::IsNullOrWhiteSpace($Path))
        {
            $FileInfo = Get-ItemProperty -Path $Path

            $Global:SOELogFile = $FileInfo.DirectoryName + '\Log'

            if(-not (Test-Path $SOELogFile))
            {
                New-Item -Path $SOELogFile -ItemType Directory | Out-Null
            }
            
            $Global:SOELogFile += '\' + ($FileInfo.Name).Replace($FileInfo.Extension, '.log')
        }
        else
        {
            Write-Warning "No log file name is given."
            break
        }

        if(-not [string]::IsNullOrWhiteSpace($Message))
        {
            Write-SOELog $Message
        }
        
        Write-Output $SOELogFile
    }
    End
    {
    }
}
#endregion

#region Write-SOELog

<#
.SYNOPSIS
    Write message to log file in common log folder.
.DESCRIPTION
    The SOE log files are stored in a common log folder and have the body name of the script (if this is not overwritten).
.EXAMPLE
    Write-SOELog -Message "Process started"

    Will add the text "Process started" to the log file as stored in $SOELogFile.
.OUTPUTS
    None
#>
function Write-SOELog
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Low')]
    [OutputType([String])]
    Param
    (
        # $LogMessage help description
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [Alias("LogMessage")]
        [string]
        $Message,

        # Break the current process
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false)]
        [switch]
        $BreakProcess
    )

    Begin
    {
        If([string]::IsNullOrWhiteSpace($SOELogFile))
        {
            Write-Warning "Log will not be written to file because the file name is empty."
        }
    }
    Process
    {
        $Msg = "{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
        if(-not [string]::IsNullOrWhiteSpace($SOELogFile))
        {
            Add-Content -Path $SOELogFile -Value $Msg
        }
        Write-Debug $Msg

        if($BreakProcess)
        {
            Write-SOELog -Message "Script terminated"
            Exit
        }
    }
    End
    {
    }
}
#endregion

#endregion

#region Process file

#region Get-SOEProcessFile

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
function Get-SOEProcessFile
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$true,
                  ConfirmImpact='Medium')]
    [OutputType([PSObject])]
    Param
    (
        # Computer name help description
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [Alias("Name")]
        [string]
        $ComputerName,

        # Literal path help description
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 2')]
        [ValidateNotNullOrEmpty()]
        [Alias("FileName")] 
        [string]
        $LiteralPath,

        # Path help description
        [Parameter(Mandatory=$false,
                   ValueFromPipeLine=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false,
                   Position=1,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [Alias("RelativePath")]
        [string]
        $Path = ".\Process"
    )

    Begin
    {
    }
    Process
    {
        $Props = [ordered]@{ComputerName='';
                            JobId='';
                            Step='';
                            OldStep='';
                            FullName='';
                            FileName='';
                            DirectoryName='';
                            Exists=$false}
        $Result = New-Object -TypeName PSObject -Property $Props

        if(-not [string]::IsNullOrWhiteSpace($LiteralPath))
        {
            $Result.FullName = $LiteralPath
        }
        else
        {
            $Result.FullName = (Get-ChildItem -Path "$Path\$ComputerName.*.txt").FullName
            if($Result.FullName -eq $null)
            {
                Write-Verbose "No process file for $($ComputerName)"
            }
        }

        if(-not [string]::IsNullOrWhiteSpace($Result.FullName))
        {
            $FileInfo = Get-ItemProperty -Path $Result.FullName
            $Result.Exists = $FileInfo.Exists
            $Result.FileName = $FileInfo.Name
            $Result.FullName = $FileInfo.FullName
            $Result.DirectoryName = $FileInfo.DirectoryName

            $Result.ComputerName = $Result.FileName -replace ("\.\w+\.txt$")
            $Result.OldStep = $Result.FileName -replace ("^$($Result.ComputerName)\.") -replace ("\.txt$")
            $Result.Step = $Result.OldStep

            if($Result.Exists)
            {
                $Match = "^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} JobId: "
                $Result.JobId = (Get-Content $Result.FullName | Where-Object {$_ -match $Match} |                    Select-Object -Last 1) -replace($Match)            }
        }
        else
        {
            $Result.ComputerName = $ComputerName
            $FileInfo = Get-ItemProperty -Path $Path
            $Result.DirectoryName = $FileInfo.FullName
        }

        Write-Output $Result
    }
    End
    {
    }
}
#endregion

#region Set-SOEProcessFile

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
function Set-SOEProcessFile
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    [OutputType([String])]
    Param
    (
        # Computer name help description
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [Alias("Name")]
        [string]
        $ComputerName,

        # Literal path help description
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 2')]
        [ValidateNotNullOrEmpty()]
        [Alias("FileName")]
        [string]
        $LiteralPath,

        # Path help description
        [Parameter(Mandatory=$false,
                   ValueFromPipeLine=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false,
                   Position=1,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()]
        [Alias("RelativePath")]
        [string]
        $Path = ".\Process",

        # Message help description
        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Message,

        # New status help description
        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [String]
        $NewStep
    )

    Begin
    {
    }
    Process
    {
        $Result = Get-SOEProcessFile -ComputerName $ComputerName

        if($LiteralPath -ne $null)
        {
            $Result = Get-SOEProcessFile -LiteralPath $LiteralPath
        }
        else
        {
            $Result = Get-SOEProcessFile -ComputerName $ComputerName -Path $Path
        }

        if($Message -ne "")
        {
            If([string]::IsNullOrWhiteSpace($Result.FullName))
            {
                if($NewStep -eq "")                    
                {
                    $Result.Step = "Unknown"
                }

                $Result.OldStep = $Result.Step
                $Result.FullName = "$($Result.DirectoryName)\$($Result.ComputerName).$($Result.Step).txt"
            }

            $LogMessage = "{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
            if ($PSCmdlet.ShouldProcess("$($Result.FileName)", "Add '$LogMessage'"))
            {
                Add-Content -Path $Result.FullName -Value $LogMessage
            }
        }

        if($NewStep -ne "" -and $Result.OldStep -ne $NewStep)
        {
            $Result.Step = $NewStep
            if ($PSCmdlet.ShouldProcess("$($Result.FileName)", "Rename"))
            {
                Rename-Item -LiteralPath $Result.FullName -NewName "$($Result.ComputerName).$NewStep.txt"
            }
            else
            {
                if($Result.Exists)
                {
                    Write-Warning "Renaming $($Result.FileName) to $($Result.ComputerName).$($Result.Step).txt"
                }
                else
                {
                    Write-Warning "Creating $($Result.FileName)"
                }
            }
        }
        
        Write-Output $Result
    }
    End
    {
    }
}
#endregion

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
            return
        }

        if ($PSCmdlet.ShouldProcess("$($Result.FileName)", "Submit transfer"))
        {
            #Submit the first file (the transfer job will automatically start in suspended mode)
            $Result = Invoke-WsmanAction -Action CreateJob -ResourceURI wmi/root/microsoft/bits/BitsClientJob -ValueSet @{DisplayName=$DisplayName; RemoteUrl="$SourceHttp/$($Files[0])"; LocalFile="$TargetPath\$($Files[0])"; Type=0} –ComputerName $ComputerName -Credential $Credential            $JobId = $Result.JobId
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

New-Alias -Name wsl Write-SOELog
Export-ModuleMember -Variable SOELogFile
Export-ModuleMember -Function Set-SOELogFile
Export-ModuleMember -Function Write-SOELog
Export-ModuleMember -Alias wsl
Export-ModuleMember -Function Get-SOEProcessFile
Export-ModuleMember -Function Set-SOEProcessFile

Export-ModuleMember -Function Get-SOEBITSTransfer
Export-ModuleMember -Function Start-SOEBITSTransfer
Export-ModuleMember -Function Complete-SOEBitsTransfer
Export-ModuleMember -Function Remove-SOEBitsTransfer
