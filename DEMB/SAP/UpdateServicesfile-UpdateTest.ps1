<#
.Synopsis
   Windows Services file update script
.DESCRIPTION
   This script picks up a updatefile (netlogon?) and if needed updates the windows services file (add/remove lines).

   - Header is copied from origional
   - Other commented lines are moved to the bottom
   - Lines will be sorted on port number
   - Comments are ignored during comparing

   Inputfile (with the updated services) has the same name and is located at the same place as the script file,
   only difference is the extension csv. (\\<path>\<thisscriptfilename>.csv)

   Format input file (CSV): action,name,port,protocol,alias,comment
   allowed values for action field:<add/remove>,<port number>,<tcp/udp>,[aliases...],[comment]
   (not not at a hashtag # to the comment field, this is done by the script)

   Output as the standard services file: <service name>  <port number>/<protocol>  [aliases...]   [#<comment>]
   Output file location: C:\Windows\System32\Drivers\etc\services

   Log file (C:\hp\logs\<thisscriptfilename>.log) is written only when the update file on the netlogon is newer as the logfile
   When changes are made, origional Services file is copied to c:\hp\logs

   To avoid signing issues, start script with the following command:
   powershell.exe -noprofile -executionpolicy bypass -file script.ps1

.NOTES
   build to support SAP and to be compatible with powershell 2.0
   december 2014 / january 2015
   arnoud.van.voorst@hp.com
#>

#REQUIRES -Version 2.0

#variables:
$servicesfile = "$PSScriptRoot\Windows\services"
$updatefile = $MyInvocation.MyCommand.Definition -replace(".ps1$",".csv")
$logfilepath = "$PSScriptRoot\Log\"
$logfile = "{0}\{1}" -f $logfilepath,($MyInvocation.MyCommand.Definition -replace(".ps1$",".log")).Split("\")[-1]


##########################
# prologue
##########################

#exit the script if there is no update available
#continue with script if updatefile is newer as logfile, or logfile does not exist
if(!((Get-Item -Path $updatefile).LastWriteTimeUtc -gt (Get-Item -Path $logfile -ErrorAction SilentlyContinue).LastWriteTimeUtc))
    {break}

#create the logfile if the file does not exitst else write todays first empty line as batch seperator
If(!(test-path $logFile)){New-Item $logFile -ItemType file >$null}
else { Add-Content -path $logFile -Value "" }
Add-Content -path $logFile -value $((get-date -Format "yyyy-MM-dd HH:mm:ss ") + "Script started (" + $env:USERNAME + "), updatefile is newer as logfile.")

#region: read and process the services file to objects

    $services = Get-Content $servicesfile

    #get the header and get the other commented lines as footer
    $header = (($services -join "`n") -split "\n[^#]",2)[0] -split "\n" | Where-Object {$_ -like "#*"}
    $footer = (($services -join "`n") -split "\n[^#]",2)[1] -split "\n" | Where-Object {$_ -like "#*"}

    $services = $services |
        Where-Object { ($_ -gt "*") -and ($_ -notlike "#*") } |
        ForEach-Object {
        
            $LineArray = ((($_ -split '/') -split '\s\s+'))

            $name = $LineArray[0].trim()
            $port = $LineArray[1].trim()
            $protocol = $LineArray[2].trim()
        
            $LineArray = $_ -split '(/tcp|/udp)\s\s+';
            if($LineArray.Count -gt 2) { $alias, $Comment = ($LineArray[2] -split '#') }

            if($alias -eq $null){$alias = ""} else {$alias = $alias.trim()}
            if($comment -eq $null){$comment = ""} else {$comment = $Comment.trim()}

        
            New-Object -TypeName PSObject |
            Add-Member –membertype NoteProperty -Name name -Value $name  -Force -PassThru |
            Add-Member –membertype NoteProperty -Name port -Value $port -Force -PassThru | 
            Add-Member –membertype NoteProperty -Name protocol -Value $protocol  -Force -PassThru|
            Add-Member –membertype NoteProperty –name alias -value $alias -Force -PassThru |
            Add-Member –membertype NoteProperty –name comment -value $comment -Force -PassThru

            $alias = $null
            $comment = $null
 
        }
    #keep to current ones to compare later...
    $oldservices = $services | Select-Object -Property *

#endregion

#region: get services update file information

    $updateservices = Import-Csv $updatefile |
        ForEach-Object {if($_.alias -eq $null){$_.alias = ""} Write-Output $_ } |
        ForEach-Object {if($_.comment -eq $null){$_.comment = ""} Write-Output $_ }

    $add = $updateservices | Where-Object {$_.action -eq "add"} | Select-Object -Property * -ExcludeProperty action
    $remove = $updateservices | Where-Object {$_.action -eq "remove"} | Select-Object -Property * -ExcludeProperty action

#endregion

##########################
# process the changes
##########################

#region: add / remove services

#remove services
    
    if($remove -ne $null)
                {
                # to remove services from the list just add a # before the name

                $services = Compare-Object -ReferenceObject $services -DifferenceObject $remove -Property name,port,protocol,alias -IncludeEqual -PassThru |
                            Where-Object {$_.SideIndicator -eq "==" -or $_.SideIndicator -eq "<=" } |
                            ForEach-Object {if($_.SideIndicator -eq "==") {$_.name = "#" + $_.name } Write-Output $_ } |
                            Select-Object -Property * -ExcludeProperty SideIndicator
               }

    # add services
    if($add -ne $null)
                {

                <## this code is replaced due to PowerShell V2 bug

                $services += Compare-Object -ReferenceObject $services -DifferenceObject $add -Property name,port,protocol,alias -PassThru |
                             Where-Object {$_.SideIndicator -eq "=>" } |
                             Select-Object -Property * -ExcludeProperty SideIndicator
                ##>

                $tmp = Compare-Object -ReferenceObject $services -DifferenceObject $add -Property name,port,protocol,alias -PassThru |
                       Where-Object {$_.SideIndicator -eq "=>" } |
                       Select-Object -Property * -ExcludeProperty SideIndicator
                if($tmp -ne $null){$services += $tmp}

                }

#endregion


##########################
# output
##########################

#region: write and format updated services file

#check if there are any changes
$changes = Compare-Object -ReferenceObject $oldservices -DifferenceObject $services -Property name,port,protocol,alias -PassThru |
           Select-Object -Property * -ExcludeProperty SideIndicator

#if there are changes - update the services file, copy the file and write log
#if($changes.Count -gt 0)
if($changes -ne $null)
    {
        #make backup of the old services file
        Copy-Item -Path $servicesfile -Destination "$logfilepath\services" -Force

        #build new services file (3 steps, header, services, footer)
        $header | Out-file $servicesfile -Encoding ascii

        $namewidth = ($services| ForEach-Object {$_.name} | Measure-Object -Maximum -Property length).Maximum
        $aliaswidth = ($services| ForEach-Object {$_.alias} | Measure-Object -Maximum -Property length).Maximum

        $services | Sort-Object -property @{Expression={[int]$_.port}} |
            ForEach-Object { if( ($_.Comment -ne "") -and ($_.Comment -ne $null) -and ($_.Comment.substring(0,1) -ne "#" ) ){$_.Comment = "#" + $_.Comment } Write-Output $_ } |
            ForEach-Object { Write-Output ("{0,-$namewidth}  {1,5}/{2}  {3,-$aliaswidth}  {4}" -f $_.name,$_.port,$_.protocol,$_.alias,$_.comment).TrimEnd() } |
            Out-file $servicesfile -Append ascii


        $footer | Out-file $servicesfile -Append ascii

        #write changes to logfile
        $changes | ft -HideTableHeaders -AutoSize | Out-File $logfile -Append ascii

    }

#endregion

Add-Content -path $logFile -value $((get-date -Format "yyyy-MM-dd HH:mm:ss ") + "Script stopped (" + $env:USERNAME + ")")