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
    Version 1.00 (2016-01-18, Kees Hiemstra)
    -- Inital version.
#>

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
                <td width="170" bgcolor="#FFFFCC"><p align="left"><font face="Arial">Created on</font></p></td>
                <td width="11" bgcolor="#FFFFCC"><p align="left"><font face="Arial">&nbsp;</font></p></td>
                <td width="705" bgcolor="#FFFFCC"><p align="left"><font face="Arial"><!-- Insert date/time of creation --></font></p></td>
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

#Define script variables
$InputFile = "$PSScriptRoot\Input.txt"
$HTMLPage = "Owner.htm"

#Define log file variables
$ScriptName = $MyInvocation.MyCommand.Definition.Replace("$PSScriptRoot\", '').Replace(".ps1", '')
$LogFile = $MyInvocation.MyCommand.Definition -replace(".ps1$",".log")
$MailSMTP = "smtp.corp.demb.com"
$MailSender = "$($ScriptName.Replace(' ', '.'))@JDECoffee.com"
$MailErrorTo = "Kees.Hiemstra@hpe.com"

#region LogFile
$Error.Clear()

#Create the log file if the file does not exits else write today's first empty line as batch separator
if (-not (Test-Path $LogFile))
{
    New-Item $logFile -ItemType file | Out-Null
}
else 
{
    Add-Content -Path $LogFile -Value "---------- --------"
}

function Write-Log([string]$Message)
{
    $LogMessage = "{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"),$Message
    Add-Content -Path $LogFile -Value $LogMessage
    Write-Debug $Message
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

#region Get-Contacts
function Get-Contacts
{
    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        #Specified group name to list members
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]
        $GroupName,

        [Parameter(Mandatory=$false)]
        [switch]
        $IsOwnerGroup
    )

    Process
    {
        $GroupName = $GroupName.Replace('DEMB\', '')
        try
        {
            Get-ADGroup -Identity $GroupName | Out-Null
        }
        catch
        {
            return "<b><font color=red>Group doesn't exist.</b></font>"
        }

        $Members = Get-ADGroupMember -Identity $GroupName -ErrorAction SilentlyContinue
        if ($Members.Count -eq 0)
        {
            return "<b><font color=red>Group doesn't contain contacts.</b></font>"
        }

        $GroupUsers = $Members | Get-ADUser -Properties Mail, DisplayName, C, Company | Where-Object { $_.Enabled }

        $Output = @()
        foreach ($User in $GroupUsers)
        {
            if([string]::IsNullOrEmpty($User.Mail))
            { 
                $Line = "<i>$($User.DisplayName)</i>" 
            }
            else
            {
                $Line = $User.Mail
            }

            if (-not ([string]::IsNullOrEmpty($User.C) -or [string]::IsNullOrEmpty($User.Company)) )
            {
                $Info = ""
                if (-not ([string]::IsNullOrEmpty($User.C)) ) { $Info = $User.C }
                
                if (-not ([string]::IsNullOrEmpty($User.Company)) )
                {
                    if (-not ([string]::IsNullOrEmpty($Info)) ) { $Info += " // " }

                    $Info += $User.Company
                }
                $Line += " <small>($Info)</small>"
            }
            $Output += $Line
        }
#       Select-Object @{n='Contact'; e={ if([string]::IsNullOrEmpty($_.Mail)) { "<i>$($_.DisplayName)</i>" } else { if ($IsOwnerGroup) { "<a href=`"mailto:$($_.Mail)?subject=Can you grant me access?`">$($_.mail)</a>" } else { $_.mail } } + " ($($_.C)/$($_.Company))" }}
#       Select-Object @{n='Contact'; e={ if([string]::IsNullOrEmpty($_.Mail)) { "<i>$($_.DisplayName)</i>" } else { if ($IsOwnerGroup) { "<a href=`"mailto:$($_.Mail)?subject=Grant me access to $($GroupName)`">$($_.mail)</a>" } else { $_.mail } } }}
        
        return ($Output -join "<br/>`n")
    }
}
#endregion

if (-not (Test-Path -Path $InputFile))
{
    Write-Break -Message "Input file is missing"
}

$ServerName = 'DESHUTFS247.corp.demb.com'

#Process lines in input file
foreach ($Line in (Get-Content -Path $InputFile | Where-Object { $_ -notmatch '^#' } | Select-Object -First 2 ))
{
    #Level 1
    $Fields = $Line -split ','
    $OpCo = $Fields[0]
    $Path = $Fields[1]
    $Incl = $Fields[2]

    if ((Test-path -Path $Path))
    {
        Write-Host $Path
        foreach ($FolderL2 in (Get-ChildItem -Path "\\$ServerName\$($Path.Replace(':', '$'))" -Directory | Select-Object -First 2))
        {
            Write-Host "`t$($FolderL2.Name)"
            if (-not (Test-Path -Path "B:\Owner\$($FolderL2.Name)"))
            {
                New-Item -Path "B:\Owner\$($FolderL2.Name)" -ItemType Directory | Out-Null
            }

            $HTMLFolders = ""
            #(Get-Acl -Path "$ShareName\$($Folder.Name)").Owner
            foreach ($FolderL3 in (Get-ChildItem -Path "$($FolderL2.FullName)" -Directory))
            {
                Write-Host "`t`t$($FolderL3.Name)"
                $Folder = $TemplateFolder.Replace('<!-- Insert folder name -->', $FolderL3.Name)

                $OwnerGroup = (((Get-Acl -Path $FolderL3.FullName).Access | Where-Object { $_.IdentityReference -like "DEMB\$($OpCo)O*" }).IdentityReference).Value

                if ([string]::IsNullOrEmpty($OwnerGroup) -or $OwnerGroup -like 'BUILTIN\*')
                {
                    $Folder = $Folder.Replace('<!-- Insert owner group name -->', '<b><font face="Arial"><font color=red>No group specified! Please contact your IT manager.</font></font></b>')
                }
                else
                {
                    $Folder = $Folder.Replace('<!-- Insert owner group name -->', $OwnerGroup)
                    $Folder = $Folder.Replace('<!-- Insert owner contacts -->', (Get-Contacts -GroupName $OwnerGroup -IsOwnerGroup))

                    $ChangeGroup = "$($OwnerGroup.Substring(0, 9))C$($OwnerGroup.Substring(10))"
                    $Folder = $Folder.Replace('<!-- Insert change group name -->', $ChangeGroup)
                    $Folder = $Folder.Replace('<!-- Insert change contacts -->', (Get-Contacts -GroupName $ChangeGroup))


                    $ReadGroup = "$($OwnerGroup.Substring(0, 9))R$($OwnerGroup.Substring(10))"
                    $Folder = $Folder.Replace('<!-- Insert read group name -->', $ReadGroup)
                    $Folder = $Folder.Replace('<!-- Insert read contacts -->', (Get-Contacts -GroupName $ReadGroup))
                }

                $HTMLFolders += $Folder
            }
            Set-Content -Path "B:\Owner\$($FolderL2.Name)\Owner.htm" -Value $TemplateOwnerFile.Replace('<!-- Insert OpCo number -->', $OpCo).Replace('<!-- Insert all scanned folders -->', $HTMLFolders).Replace('<!-- Insert date/time of creation -->', (Get-Date -DisplayHint DateTime))
        }
    }#$Path exists
    else
    {
        Write-Log -Message "Input path [$Path] doesn't exist"
    }
}#foreach input.txt