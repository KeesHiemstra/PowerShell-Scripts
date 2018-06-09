<#
    OwnerPage.ps1

    Replaces the RPMTools\mk_own.vbs

    The script will use the input.txt file in its current folder to create the Owner.htm pages on the server.

    --- Change requests
    - Aldea 1188733: RITM0010413 incorrect htm.files for network level 3 folders
    - Aldea 1157766: R90655 NL 0002 Please expand all owner files in DFS with users who have access...

    --- Last modified by
    Kees.Hiemstra@hpe.com

    --- Version history
    Version 2.14 (2016-04-05, Kees Hiemstra)
    - Bug fix: Previous bug fixes didn't work. The issue was caused because the only member in the group
      was disabled. The empry group was handled wrongly.
    Version 2.13 (2016-03-29, Kees Hiemstra)
    - Bug fix: [array]$Member = ($GroupPath).Member with only one user was not processed correctly.
    Version 2.12 (2016-03-23, Kees Hiemstra)
    - Typo fix: Renamed contacts of the read/write and read group to members.
    - Bug fix: ([ADSI]$GroupPath).Member with only one user was not processed correctly.
    - Bug fix for troubleshooting: Add server name to the creation date/time at the bottom of the file.
    Version 2.11 (2016-02-22, Kees Hiemstra)
    - Send last part of the log file when an error is send.
    Version 2.10 (2016-02-09, Kees Hiemstra)
    - Start the original QM_diruse.cmd from this script to get ForderSize.txt.
    - Disabled FolderSize.txt because of errors on the deep folder structure on most servers.
    Version 2.01 (2016-02-08, Kees Hiemstra)
    - Bug fix: Get-FolderSize often reports "The specified path, file name, or both are too long."
    Version 2.00 (2016-02-08, Kees Hiemstra)
    - Added the creation of FolderSize.txt.
    Version 1.34 (2016-02-02, Kees Hiemstra)
    - Bug fix: Having no access to a folder was captured on level to deep.
    Version 1.13 (2016-01-31, Kees Hiemstra)
    - Bug fix: Catched error at getting acl still showed up as error at the end of the script.
    Version 1.12 (2016-01-25, Kees Hiemstra)
    - Bug fix: Catch error when getting acl.
    - Added log when error mail was send.
    Version 1.11 (2016-01-22, Kees Hiemstra)
    - Bug fix: Handle duplicate owner group assignments on folder level.
    - Bug fix: Skip empty lines in Input.txt.
    - Bug fix: The email from Stop-Script with the unrecoverable error doesn't contain the server name on which it failed.
    Version 1.10 (2016-01-20, Kees Hiemstra)
    - Sorted contact output.
    - Use SAMAccountName when both mail and display name are empty.
    Version 1.00 (2016-01-19, Kees Hiemstra)
    - Inital version, re-written from vbs.
#>

#Assebly needed to convert special charactes to HTML notation
Add-Type -AssemblyName System.Web

#region HTML templates
$TemplateOwnerFile = @"
<html>
<head>
<title>JDE Data Owners</title>
</head>
<body bgcolor="white" text="black" link="blue" vlink="purple" alink="red">
  <!-- Page main table -->
  <table align="center" width="64%" style="border-collapse:collapse;" cellspacing="0">
    <tr>
      <td width="902" align="center">
        <div align="center">
          <table style="BORDER-COLLAPSE: collapse" cellspacing="0" width="900" bgcolor="#FFFFCC">
            <tbody>
              <tr align="center">
                <td width="19"><p align="justify">&nbsp;</p></td>
                <td width="839">
                  <div align="justify"><font face="Arial">
                    This file explains whom you should address when you want to have access to any of the subfolders for OpCo</font>
                    <b><font face="Arial"><!-- Insert OpCo number --></font></b>
                  </div><p>&nbsp;</p></td>
                <td width="20"><p align="justify">&nbsp;</p></td>
              </tr>
              <tr align="center"> 
                <td width="19"><p align="justify">&nbsp;</p></td>
                <td width="839">
                  <div align="justify"><font face="Arial">
                    Please send a mail to one of the listed contacts (referred to as Data Owners) specifying to which group you 
                    would like to be added.</font>
                  </div>
                  <p align="justify"><font face="Arial">
                  It is best to copy the group name from this file and paste it into a mail to prevent typing errors.<br>
                  You will only need to request to be added to one of the groups belonging to a folder.<br></font></p>
                  <p align="justify"><font face="Arial">
                  The Data Owners will validate your request and notify you when it has been processed.</font></p>
                </td>
                <td width="20"><p align="justify">&nbsp;</p></td> 
              </tr>
              <tr align="center">
                <td width="19"><p align="justify">&nbsp;</p></td>
                <td width="839"><br><p align="justify"><font face="Arial">
                For more information please visit:</font><br>
   	            <b><a href="https://coffeeandtea.sharepoint.com/sites/rs1/it-global/Pages/File-Sharing.aspx">
                  <font face="Arial">https://coffeeandtea.sharepoint.com/sites/rs1/it-global/Pages/File-Sharing.aspx</font></a></b></p></td>
                <td width="20"> <p align="justify">&nbsp;</p>
                </td>
              </tr> 
              <tr align="center"> 
                <td width="19"><p align="justify">&nbsp;</p></td>
                <td width="839"><p align="justify"><font face="Arial">&nbsp;</font></p></td>
                <td width="20"><p align="justify">&nbsp; </p></td>
              </tr>
            </tbody>
          </table>
          <table style="BORDER-COLLAPSE: collapse" cellspacing="0" width="901" bgcolor="#99FFFF"> 
            <tbody>
              <tr align="center">
                <td width="19" bgcolor="#FFFFCC"><p align="left">&nbsp;</p></td>
                <td width="170" bgcolor="#FFFFCC"><p align="left"><font face="Arial"></font></p></td>
                <td width="11" bgcolor="#FFFFCC"><p align="left"><font face="Arial">&nbsp;</font></p></td>
                <td width="705" bgcolor="#FFFFCC"><p align="left"><b><font face="Arial"></font></b></p></td>
                <td width="22" bgcolor="#FFFFCC"><p>&nbsp;</p></td> 
              </tr>
              <!-- Insert all scanned folders -->
              <tr align="center">
                <td width="19" bgcolor="#FFFFCC"><p align="left">&nbsp;</p></td>
                <td width="170" bgcolor="#FFFFCC"><p align="left"><font face="Arial"><small>Created on</small></font></p></td>
                <td width="11" bgcolor="#FFFFCC"><p align="left"><font face="Arial">&nbsp;</font></p></td>
                <td width="705" bgcolor="#FFFFCC"><p align="left"><font face="Arial"><small><!-- Insert date/time of creation --></small></font></p></td>
                <td width="22" bgcolor="#FFFFCC"><p>&nbsp;</p></td>
              </tr>
            </tbody>
          </table>
        </div>
      </td>
    </tr>
  </table>
</body>
</html>
"@

$TemplateFolder = @"
              <tr align="center">
                <td width="19" bgcolor=""><p align="left">&nbsp;</p></td> 
                <td width="170" bgcolor=""><p align="left"><font face="Arial">Folder:</font></p></td>
                <td width="11" bgcolor=""><p align="left"><font face="Arial">&nbsp;</font></p></td>
                <td width="705" bgcolor=""><p align="left"><b><font face="Arial"><!-- Insert folder name --></font></b></p></td>
                <td width="22" bgcolor=""><p>&nbsp;</p></td>
              </tr>
              <tr align="center">
                <td width="19" bgcolor=""><p align="left">&nbsp; </p></td>
                <td width="170" bgcolor=""><p align="left"><font face="Arial">Owner Group:</font></p></td>
                <td width="11" bgcolor=""><p align="left"><font face="Arial">&nbsp;</font></p></td>
                <td width="705" bgcolor=""><p align="left"><b><font face="Arial"><!-- Insert owner group name --></font></b> </p></td>
                <td width="22" bgcolor=""><p>&nbsp;</p></td>
              </tr>
              <tr align="center">
                <td width="19" bgcolor=""><p align="left">&nbsp;</p></td> 
                <td width="170" bgcolor="" valign="top"><p align="left"><font face="Arial">Owner Contact:</font></p></td>
                <td width="11" bgcolor=""><p align="left"><font face="Arial">&nbsp;</font></p></td>
                <td width="705" bgcolor=""><p align="left"><font face="Arial"><!-- Insert owner contacts --></font></p></td>
                <td width="22" bgcolor=""><p>&nbsp;</p></td>
              </tr>
              <tr align="center">
                <td width="19" bgcolor=""><p align="left">&nbsp;</p></td>
                <td width="170" bgcolor=""><p align="left"><font face="Arial">Read/Write Group:</font></p></td>
                <td width="11" bgcolor=""><p align="left"><font face="Arial">&nbsp;</font></p></td>
                <td width="705" bgcolor=""><p align="left"><b><font face="Arial"><!-- Insert change group name --></font></b></p></td>
                <td width="22" bgcolor=""><p>&nbsp;</p></td>
              </tr>
              <tr align="center">
                <td width="19" bgcolor=""><p align="left">&nbsp;</p></td>
                <td width="170" bgcolor="" valign="top"><p align="left"><font face="Arial">Read/Write members:</font></p></td>
                <td width="11" bgcolor=""><p align="left"><font face="Arial">&nbsp;</font></p></td>
                <td width="705" bgcolor=""><p align="left"><font face="Arial"><!-- Insert change contacts --></font></p></td>
                <td width="22" bgcolor=""><p>&nbsp;</p></td>
              </tr>
              <tr align="center">
                <td width="19" bgcolor=""><p align="left">&nbsp;</p></td>
                <td width="170" bgcolor=""><p align="left"><font face="Arial">Read Group:</font></p></td>
                <td width="11" bgcolor=""><p align="left"><font face="Arial">&nbsp;</font></p></td>
                <td width="705" bgcolor=""><p align="left"><b><font face="Arial"><!-- Insert read group name --></font></b></p></td>
                <td width="22" bgcolor=""><p>&nbsp;</p></td>
              </tr>
              <tr align="center">
                <td width="19" bgcolor=""><p align="left">&nbsp;</p></td>
                <td width="170" bgcolor="" valign="top"><p align="left"><font face="Arial">Read members:</font></p></td>
                <td width="11" bgcolor=""><p align="left"><font face="Arial">&nbsp;</font></p></td>
                <td width="705" bgcolor=""><p align="left"><font face="Arial"><!-- Insert read contacts --></font></font></p></td>
                <td width="22" bgcolor=""><p>&nbsp;</p></td>
              </tr>
              <tr align="center">
                <td width="19" bgcolor="#FFFFCC"><p align="left">&nbsp;</p></td>
                <td width="170" bgcolor="#FFFFCC"><p align="left"><font face="Arial"></font></p></td>
                <td width="11" bgcolor="#FFFFCC"><p align="left"><font face="Arial">&nbsp;</font></p></td>
                <td width="705" bgcolor="#FFFFCC"><p align="left"><b><font face="Arial"></font></b></p></td>
                <td width="22" bgcolor="#FFFFCC"><p>&nbsp;</p></td>
              </tr>
"@

#endregion

#Define script variables because of compatability with PowerShell version 2.0
$ScriptPath = $MyInvocation.MyCommand.Path -replace $MyInvocation.MyCommand.Name
$QM_DirUse = "$($ScriptPath)QM_diruse.cmd"
$InputFile = "$($ScriptPath)Input.txt"

#Define log file variables
$ScriptName = $MyInvocation.MyCommand.Name.Replace(".ps1", '')
$LogFile = $MyInvocation.MyCommand.Definition -replace(".ps1$",".log")
$MailSMTP = "smtp.corp.demb.com"
$MailSender = "$($ScriptName.Replace(' ', '.'))@JDECoffee.com"
$MailErrorTo = "Kees.Hiemstra@hpe.com"

#region LogFile
$LogStart = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
$Error.Clear()

#Create the log file if the file does not exits else write today's first empty line as batch separator
if (-not (Test-Path $LogFile))
{
    New-Item $LogFile -ItemType file | Out-Null
}
else 
{
    Add-Content -Path $LogFile -Value "---------- --------"
}

function Write-Log([string]$Message)
{
    $LogMessage = "{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Add-Content -Path $LogFile -Value $LogMessage
#    Write-Host $LogMessage
}

#Finish script and report on $Error
function Stop-Script([bool]$WithError)
{
    if ($Error.Count -gt 0)
    {
        try
        {

            $MailErrorSubject = "Error in $ScriptName on $($Env:ComputerName)"
            $MailErrorBody = "The script $ScriptName has reported the following error(s):`n`n"
            $MailErrorBody += $Error | Out-String

            $MailErrorBody += "`n-------------------`n"
            $MailErrorBody += (Get-Content $LogFile | Where-Object { $_.SubString(0, 19) -ge $LogStart }) -join "`n"

            Send-MailMessage -SmtpServer $MailSMTP -From $MailSender -To $MailErrorTo -Subject $MailErrorSubject -Body $MailErrorBody
            Write-Log -Message "Sent error mail to home"
        }
        catch
        {
            Write-Log -Message "Unable to send error mail to home"
        }
    }

    if ($WithError)
    {
        Write-Log -Message "Script stopped with an error"
    }
    else
    {
        Write-Log -Message "Script ended normally"
    }
    Exit
}

#This function write the error to the logfile and exit the script
function Write-Break([string]$Message)
{
    Write-Log -Message $Message
    Write-Error -Message $Message
    Stop-Script -WithError $true
}

Write-Log -Message "Script started ($($env:USERNAME))"
#endregion

if (-not (Test-Path -Path $InputFile))
{
    Write-Break -Message "Input file is missing"
}

#Preparing domain search
$Domain = New-Object -TypeName System.DirectoryServices.DirectoryEntry
$Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher
$Searcher.SearchRoot = $Domain
$Searcher.PageSize = 1000
$Searcher.SearchScope = "Subtree"

$Properties = @('Name', 'DisplayName', 'Mail', 'C', 'Company', 'UserAccountControl', 'SAMAccountName')
foreach ($Property in $Properties) { $Searcher.PropertiesToLoad.Add($Property) | Out-Null }

#region Get-Contacts
function Get-Contacts
{
    Param
    (
        #Specified group path to list members
        [Parameter(Mandatory=$true)]
        [string]
        $GroupPath,

        #The specified group is the owner group
        [Parameter(Mandatory=$false)]
        [switch]
        $IsOwnerGroup
    )

    Process
    {
        $Output = @()
        [array]$Members = [array](([ADSI]$GroupPath).Member)
        foreach ($Member in $Members)
        {
            #Get user details
            $Searcher.Filter = "(&(objectCategory=person)(objectClass=user)(DistinguishedName=$Member))"
            $ADUser = $Searcher.FindOne()

            #Report only enabled users
            $Enabled = (([uint64]($ADUser.Properties.useraccountcontrol -join ",")) -band 2) -eq 0
            if ($ADUser -ne $null -and $Enabled)
            {
                #Note: Properties are case sensitive and are all lower case
                #Used -join "," to overcome the mix of arrays and null values; the result in this case will always be a string
                $Mail = $ADUser.Properties.mail -join ","
                $DisplayName = $ADUser.Properties.displayname -join ","
                $C = $ADUser.Properties.c -join ","
                $Company = $ADUser.Properties.company -join ","
                $SAMAccountName = $ADUser.Properties.samaccountname -join ","

                if([string]::IsNullOrEmpty($Mail))
                { 
                    if ([string]::IsNullOrEmpty($DisplayName))
                    {
                        $Line = "<i>$($SAMAccountName)</i>"
                    }
                    else
                    {
                        $Line = "<i>$($DisplayName)</i>"
                    }
                }
                else
                {
                    $Line = $Mail
                }

                if (-not ([string]::IsNullOrEmpty($C) -or [string]::IsNullOrEmpty($Company)) )
                {
                    $Info = ""
                    if (-not ([string]::IsNullOrEmpty($C)) ) { $Info = $C }
                
                    if (-not ([string]::IsNullOrEmpty($Company)) )
                    {
                        if (-not ([string]::IsNullOrEmpty($Info)) ) { $Info += " // " }

                        $Info += $Company
                    }
                    $Line += " <small>($Info)</small>"
                }
                $Output += $Line
            }#Enabled user
        }#foreach group member

        if ($Output -eq $null -or $Output.Count -eq 0)
        {
            return "<font color=red>Group doesn't contain members.</font>"
        }

        return ( ($Output | Sort-Object) -join "<br/>`n" )
    }
}
#endregion

#region Get-FolderSize
function Get-FolderSize([string]$Path)
{
    New-PSDrive -Name XX -PSProvider FileSystem -Root $Path
    $Folder = Get-ChildItem -LiteralPath XX: -Recurse -ErrorAction SilentlyContinue |
        Measure-Object -Sum Length -ErrorAction SilentlyContinue |
        Select Count, @{n='Size'; e={ $_.Sum }}
    Remove-PSDrive -Name XX
    Write-Output (New-Object -TypeName PSObject -Property (@{'Path'=$Path; 'Size'=[math]::Round($Folder.Size, 0); 'Count'=[math]::Round($Folder.Count, 0)}))
}
#endregion

#Process lines in input file
foreach ($Line in (Get-Content -Path $InputFile | Where-Object { -not [string]::IsNullOrEmpty($_) -and $_ -notmatch '^#' } ))
{
    #Level 1
    $Fields = $Line -split ','
    $OpCo = $Fields[0]
    $Path = $Fields[1]
    $Incl = $Fields[2]

    if ((Test-path -Path $Path))
    {
        foreach ($FolderL2 in (Get-ChildItem -Path $Path | Where-Object { ($_.Attributes -join ";") -like '*Directory*' } ))
        {
            $HTMLFolders = ""
            $FolderSize = @()
            try
            {
                $L3Folders = (Get-ChildItem -Path "$($FolderL2.FullName)" | Where-Object { ($_.Attributes -join ";") -like '*Directory*' } )
                foreach ($FolderL3 in $L3Folders)
                {
                    $Folder = $TemplateFolder.Replace('<!-- Insert folder name -->', [System.Web.HttpUtility]::HTMLEncode($FolderL3.Name))

                    try
                    {
                        #Filter out the ownergroup in case there are more groups added
                        $OwnerGroup = (((Get-Acl -Path $FolderL3.FullName).Access | Where-Object { $_.IdentityReference -like "DEMB\$($OpCo)O*" }).IdentityReference).Value | Select-Object -First 1
                    }
                    catch
                    {
                        Write-Log -Message "Error getting acl from $($FolderL3.FullName)"

                        #Delete the captured error from the error list
                        $Error.RemoveAt(0)
                    }

                    if ([string]::IsNullOrEmpty($OwnerGroup) -or $OwnerGroup -like 'BUILTIN\*')
                    {
                        $Folder = $Folder.Replace('<!-- Insert owner group name -->', '<b><font face="Arial"><font color=red>No owner group specified! Please contact your IT manager.</font></b>')
                    }
                    else
                    {
                        $Folder = $Folder.Replace('<!-- Insert owner group name -->', [System.Web.HttpUtility]::HTMLEncode($OwnerGroup))
                        $Searcher.Filter = "(&(objectCategory=group)(objectClass=group)(CN=$($OwnerGroup.Replace('DEMB\', ''))))"
                        $Group = $Searcher.FindOne()
                        if ($Group -ne $null)
                        {
                            $Folder = $Folder.Replace('<!-- Insert owner contacts -->', (Get-Contacts -GroupPath $Group.Path -IsOwnerGroup))
                        }
                        else
                        {
                            $Folder = $Folder.Replace('<!-- Insert owner contacts -->', "<b><font color=red>Group doesn't exist.</b></font>")
                        }

                        $GroupName = "$($OwnerGroup.Substring(0, 9))C$($OwnerGroup.Substring(10))"
                        $Folder = $Folder.Replace('<!-- Insert change group name -->', [System.Web.HttpUtility]::HTMLEncode($GroupName))
                        $Searcher.Filter = "(&(objectCategory=group)(objectClass=group)(CN=$($GroupName.Replace('DEMB\', ''))))"
                        $Group = $Searcher.FindOne()
                        if ($Group -ne $null)
                        {
                            $Folder = $Folder.Replace('<!-- Insert change contacts -->', (Get-Contacts -GroupPath $Group.Path))
                        }
                        else
                        {
                            $Folder = $Folder.Replace('<!-- Insert change contacts -->', "<b><font color=red>Group doesn't exist.</b></font>")
                        }

                        $GroupName = "$($OwnerGroup.Substring(0, 9))R$($OwnerGroup.Substring(10))"
                        $Folder = $Folder.Replace('<!-- Insert read group name -->', [System.Web.HttpUtility]::HTMLEncode($GroupName))
                        $Searcher.Filter = "(&(objectCategory=group)(objectClass=group)(CN=$($GroupName.Replace('DEMB\', ''))))"
                        $Group = $Searcher.FindOne()
                        if ($Group -ne $null)
                        {
                            $Folder = $Folder.Replace('<!-- Insert read contacts -->', (Get-Contacts -GroupPath $Group.Path))
                        }
                        else
                        {
                            $Folder = $Folder.Replace('<!-- Insert read contacts -->', "<b><font color=red>Group doesn't exist.</b></font>")
                        }
                    }#Owner group specified

                    $HTMLFolders += $Folder

                    #$FolderSize += Get-FolderSize -Path $FolderL3.FullName

                }#foreach level 3 folder
                try
                {
                    Set-Content -Path "$($FolderL2.FullName)\Owner.htm" -Value $TemplateOwnerFile.Replace('<!-- Insert OpCo number -->', $OpCo).Replace('<!-- Insert all scanned folders -->', $HTMLFolders).Replace('<!-- Insert date/time of creation -->', "$(Get-Date -DisplayHint DateTime) @ $($env:COMPUTERNAME)")
                }
                catch
                {
                    Write-Log "Can't write Owner.htm to folder [$($FolderL2.FullName)]"

                    #Delete the captured error from the error list
                    #$Error.RemoveAt(0)
                }
                #try
                #{
                #    $FolderSizeTxt = "Size report for $($FolderL2.Name)`r`n`r`n    Size (mb)  Files  Directory`r`n"
                #    foreach($Folder in $FolderSize)
                #    {
                #        $FolderSizeTxt += "$(('{0:0.00}' -f [math]::Round($Folder.Size / 1mb, 2)).PadLeft(13,' '))$(('{0}' -f $Folder.Count).PadLeft(7, ' '))  $($Folder.Path.Replace($FolderL2.Name, 'Sub-Total: .\'))`r`n"
                #    }
                #    $TotalFolderSize = $FolderSize | Measure-Object -Sum Size, Count
                #    
                #    $FolderSizeTxt += "$(('{0:0.00}' -f [math]::Round($TotalFolderSize[0].Sum / 1mb, 2)).PadLeft(13,' '))$(('{0}' -f $TotalFolderSize[1].Sum).PadLeft(7, ' '))  Total`r`n"
                #    $FolderSizeTxt += "=====================`r`n$((Get-Date -DisplayHint DateTime))"
                #    Set-Content -Path "$($FolderL2.FullName)\FolderSize.txt" -Value $FolderSizeTxt
                #}
                #catch
                #{
                #    Write-Log "Can't write FolderSize.txt to folder [$($FolderL2.FullName)]"
                #
                #    #Delete the captured error from the error list
                #    #$Error.RemoveAt(0)
                #}
            }#try access level 2 folder
            catch
            {
                Write-Log "Can't access folder [$($FolderL2.FullName)]"

                #Delete the captured error from the error list
                $Error.RemoveAt(0)
            }
        }#foreach level 2 folder
    }#$Path exists
    else
    {
        Write-Log -Message "Input path [$Path] doesn't exist"
    }
}#foreach row in input.txt

if ((Test-Path $QM_DirUse))
{
    Invoke-Expression -Command $QM_DirUse
}

Stop-Script -WithError $false
