<#
    OwnerPage.ps1

    Replaces the RPMTools\mk_own.vbs

    The script will use the input.txt file in its current folder to create the Owner.htm pages on the server.

    --- Change requests
    - Aldea 1188733: RITM0010413 incorrect htm.files for network level 3 folders
    - Aldea 1157766: R90655 NL 0002 Please expand all owner files in DFS with users who have access...

    --- Contacts
    JDESOEAdmins@hpe.com

    --- Version history
    Version 1.00 (2016-01-19, Kees Hiemstra)
    - Inital version.
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
                <td width="170" bgcolor="" valign="top"><p align="left"><font face="Arial">Read/Write Contact:</font></p></td>
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
                <td width="170" bgcolor="" valign="top"><p align="left"><font face="Arial">Read Contact:</font></p></td>
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
$InputFile = "$($ScriptPath)Input.txt"

#Define log file variables
$ScriptName = $MyInvocation.MyCommand.Name.Replace(".ps1", '')
$LogFile = $MyInvocation.MyCommand.Definition -replace(".ps1$",".log")
$MailSMTP = "smtp.corp.demb.com"
$MailSender = "$($ScriptName.Replace(' ', '.'))@JDECoffee.com"
$MailErrorTo = "Kees.Hiemstra@hpe.com"

#region LogFile
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
        $Subject = "Error in $ScriptName"
        $MailBody = "The script $ScriptName has reported the following error(s):`n`n"
        $MailBody += $Error | Out-String

        Send-MailMessage -SmtpServer $MailSMTP -From $MailSender -To $MailErrorTo -Subject $Subject -Body $MailBody
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

$Properties = @('Name', 'DisplayName', 'Mail', 'C', 'Company', 'UserAccountControl')
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
#Write-Log -Message "Get-Contacts GroupPath: $GroupPath"
        $Output = @()
        $Members = ([ADSI]$GroupPath).Member
#Write-Log -Message ($Members.GetType())
        foreach ($Member in $Members)
        {
#Write-Log -Message $Member
            #Get user details
            $Searcher.Filter = "(&(objectCategory=person)(objectClass=user)(DistinguishedName=$Member))"
            $ADUser = $Searcher.FindOne()

            #Report only enabled users
            $Enabled = (([uint64]($ADUser.Properties.useraccountcontrol -join ",")) -band 2) -eq 0
            if ($ADUser -ne $null -and $Enabled)
            {
                $Mail = $ADUser.Properties.mail -join ","
                $DisplayName = $ADUser.Properties.displayname -join ","
                $C = $ADUser.Properties.c -join ","
                $Company = $ADUser.Properties.company -join ","

                if([string]::IsNullOrEmpty($Mail))
                { 
                    $Line = "<i>$($DisplayName)</i>" 
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
            }#Enabled user

            $Output += $Line
        }#foreach group member

        if ($Output.Count -eq 0)
        {
            return "<b><font color=red>Group doesn't contain contacts.</b></font>"
        }

        return ($Output -join "<br/>`n")
    }
}
#endregion

#Process lines in input file
foreach ($Line in (Get-Content -Path $InputFile | Where-Object { $_ -notmatch '^#' } ))
#foreach ($Line in (Get-Content -Path $InputFile | Where-Object { $_ -notmatch '^#' } | Select-Object -First 1 ))
{
    #Level 1
    $Fields = $Line -split ','
    $OpCo = $Fields[0]
    $Path = $Fields[1]
    $Incl = $Fields[2]

    if ((Test-path -Path $Path))
    {
        foreach ($FolderL2 in (Get-ChildItem -Path $Path | Where-Object { ($_.Attributes -join ";") -like '*Directory*' } ))
#        foreach ($FolderL2 in (Get-ChildItem -Path $Path | Where-Object { ($_.Attributes -join ";") -like '*Directory*' } | Select-Object -First 1 ))
        {
#Write-Log -Message $FolderL2.FullName
            $HTMLFolders = ""
            foreach ($FolderL3 in (Get-ChildItem -Path "$($FolderL2.FullName)" | Where-Object { ($_.Attributes -join ";") -like '*Directory*' } ))
#            foreach ($FolderL3 in (Get-ChildItem -Path "$($FolderL2.FullName)" | Where-Object { ($_.Attributes -join ";") -like '*Directory*' | Select-Object -First 1 }))
            {
#Write-Log -Message $FolderL3.FullName
                $Folder = $TemplateFolder.Replace('<!-- Insert folder name -->', [System.Web.HttpUtility]::HTMLEncode($FolderL3.Name))

                #Filter out the ownergroup in case there are more groups added
                $OwnerGroup = (((Get-Acl -Path $FolderL3.FullName).Access | Where-Object { $_.IdentityReference -like "DEMB\$($OpCo)O*" }).IdentityReference).Value
#Write-Log -Message $OwnerGroup

                if ([string]::IsNullOrEmpty($OwnerGroup) -or $OwnerGroup -like 'BUILTIN\*')
                {
                    $Folder = $Folder.Replace('<!-- Insert owner group name -->', '<b><font face="Arial"><font color=red>No owner group specified! Please contact your IT manager.</font></b>')
                }
                else
                {
                    $Folder = $Folder.Replace('<!-- Insert owner group name -->', [System.Web.HttpUtility]::HTMLEncode($OwnerGroup))
                    $Searcher.Filter = "(&(objectCategory=group)(objectClass=group)(CN=$($OwnerGroup.Replace('DEMB\', ''))))"
#Write-Log -Message $Searcher.Filter
                    $Group = $Searcher.FindOne()
                    if ($Group -ne $null)
                    {
                        $Folder = $Folder.Replace('<!-- Insert owner contacts -->', (Get-Contacts -GroupPath $Group.Path -IsOwnerGroup))
                    }
                    else
                    {
                        $Folder = $Folder.Replace('<!-- Insert owner contacts -->', "<b><font color=red>Group doesn't exist.</b></font>")
#Write-Log -Message "Owner group [$OwnerGroup] doen't exist"
                    }

                    $GroupName = "$($OwnerGroup.Substring(0, 9))C$($OwnerGroup.Substring(10))"
                    $Folder = $Folder.Replace('<!-- Insert change group name -->', [System.Web.HttpUtility]::HTMLEncode($GroupName))
                    $Searcher.Filter = "(&(objectCategory=group)(objectClass=group)(CN=$($GroupName.Replace('DEMB\', ''))))"
#Write-Log -Message $Searcher.Filter
                    $Group = $Searcher.FindOne()
                    if ($Group -ne $null)
                    {
                        $Folder = $Folder.Replace('<!-- Insert change contacts -->', (Get-Contacts -GroupPath $Group.Path))
                    }
                    else
                    {
                        $Folder = $Folder.Replace('<!-- Insert change contacts -->', "<b><font color=red>Group doesn't exist.</b></font>")
#Write-Log -Message "Change group [$GroupName] doen't exist"
                    }

                    $GroupName = "$($OwnerGroup.Substring(0, 9))R$($OwnerGroup.Substring(10))"
                    $Folder = $Folder.Replace('<!-- Insert read group name -->', [System.Web.HttpUtility]::HTMLEncode($GroupName))
#Write-Log -Message $Searcher.Filter
                    $Searcher.Filter = "(&(objectCategory=group)(objectClass=group)(CN=$($GroupName.Replace('DEMB\', ''))))"
                    $Group = $Searcher.FindOne()
                    if ($Group -ne $null)
                    {
                        $Folder = $Folder.Replace('<!-- Insert read contacts -->', (Get-Contacts -GroupPath $Group.Path))
                    }
                    else
                    {
                        $Folder = $Folder.Replace('<!-- Insert read contacts -->', "<b><font color=red>Group doesn't exist.</b></font>")
#Write-Log -Message "Read group [$GroupName] doen't exist"
                    }
                }#Owner group specified

                $HTMLFolders += $Folder
            }
            try
            {
                Set-Content -Path "$($FolderL2.FullName)\Owner.htm" -Value $TemplateOwnerFile.Replace('<!-- Insert OpCo number -->', $OpCo).Replace('<!-- Insert all scanned folders -->', $HTMLFolders).Replace('<!-- Insert date/time of creation -->', (Get-Date -DisplayHint DateTime))
            }
            catch
            {
                Write-Log "Can't write Owner.htm to folder [$($FolderL2.FullName)]"
            }
        }
    }#$Path exists
    else
    {
        Write-Log -Message "Input path [$Path] doesn't exist"
    }
}#foreach row in input.txt

Stop-Script -WithError $false
