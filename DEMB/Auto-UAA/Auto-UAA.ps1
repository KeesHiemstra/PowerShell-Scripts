<#
Initial version: arnoud.van.voorst@hp.com
Aldea 945952: R85379 AD account provisioning script
#>

<#
This script is part of a bundle of three scripts that run as scheduled jobs (Auto-UAA.ps1, Auto-UAA-messaging.ps1 and Auto-UAA-reprorting.ps1)
    Auto-UAA.ps1 is the script that read an input file with users and creates AD-accounts.
    Auto-UAA-messaging.ps1 is the script that creates mailboxes, enables Lync and set the password.
    Auto-UAA-reporting.ps1 generates a report of the accounts created yesterday

Txt files in the process and complete folders are use as data transfer between the scripts.
    <samaccountname>.mail.txt -> account waiting for mailbox creation
    <samaccountname>.lync.txt -> account waiting for enabling Lync
    <samaccountname>.enable.txt -> account waiting for password and enabling

Auto-UAA.cfg is a comma separated file that contains variables that are used in all scripts
#>

<# History
2017-02-14 v4.0.3 KHi  Bug fix: Removing ! from EA2 that has no other tags lead to undetected exception.
2016-09-02 v4.0.2 KHi  New Exchange server.
2016-08-03 v4.0.1 KHi  Bug fix. ExtensionAttribute11 was not update when a user was aleady enabled.
2016-07-12 v4.0.0 KHi  Added extra field coming from IDM (ExtensionAttribute11).
2016-06-02 V3.1.1 KHi  Bug fix: ExtensionAttribute2 was not restored at re-enabling if ExtensionAttribute2 started with !
2016-05-12 v3.1.0 KHi  Implemented granulair Office license recovery at account re-enabling.
2015-10-02 v3.0.2 KHi  Set country to Cook Islands is no country code is provided. Bug fix in re-enabling users. In some cases the re-enabled user got the
                       wrong userPricipalName that due to the replication delay was not noted by the messaging script.
2015-09-08 v3.0.1 KHi  Bug fix. Report file for already enabled account was spelled wrongly.
2015-08-27 V3.0   KHi  Prevent to process duplicate employeeID “in memory”.
2015-08-25 v2.9   KHi  Change extensionAttribut15 to "Auto re-enabled (already enabled) at YYYY-MM-DD hh:mm:ss" if the account is already enabled and schedule the account for reporting.
2015-08-05 v2.8   KHi  Set country code to 'CK' when no country code is provided.
2015-06-11 v2.7.1 KHi  Changing the UPN domain from @demb.com to @jdecoffee.com, always in small letters.
2015-06-11 v2.7   KHi  Changing the UPN domain from @demb.com to @JDECoffee.com.
2015-04-01 v2.6   AVV  Improved SAMAccount/UPN check, not only check is SAMAccount/UPN already exists agains AD only but also agains not yet created users in memory
2015-03-24 v2.5   PRW  Added processing of "DisplayName" from input file
2015-03-23 v2.4.1 KHi  Added start line in processing file to measure creating time.
2015-01-28 v2.4   KHi  UserPricipalName needs to be FirstName.LastName instead of SAMAccountName.
2015-01-07 v2.3.2 KHi  Improvent on logging.
2014-12-17 v2.3.1 KHi  Improvent on logging.
2014-12-11 v2.3   KHi  Added Re-enable user.
2014-12-12 v2.2.7 KHi  Added ACN,BAKERY ES,EXTERNAL,IBM,IBM-P,PMEXTSUP,SL IAUDIT,SUPERUSERS,YSTEM as exteptions in departmentnumberException.
2014-09-29 v2.2.6 KHi  Move even the empty input file on request of Plamen.
2014-08-25 v2.2.5 KHi  Chaning the scope of the imported cfg file.
                       Changed the way the backslash is tested at the of the path variable.
2014-08-13 v2.2.4 KHi  Moved the send errors to $emailErrorsTo to the bottom of the process.
2014-08-11 v2.2.3 KHi  Added -Descending to Get-ChildItem... Sort-Object otherwise the ($inputFiles[0]).Name contains the oldest written file instead of the newest written file.
                       Added some logging for flow check.
2014-08-08 v2.2.2 KHi  Fixed duplicated employeeID check. Used -LDAPFilter instead of -Filter and added the remark to the log file.
2014-08-07 v2.2.1 AVV  Fixed duplicated employeeID check -> added [string] (went wrong when the result was an array (Multiple entries)).
2014-08-06 v2.2   AVV  Added duplicate employeeID check.
2014-08-05 v2.1   AVV  Added run only from scheduler test (-force).
                       Added username to startline logfile.
                  KHi  locationBasedOnEmployeeID hash table simplyfied.
                  KHi  Add increasing digit if samaccountname already exists.
2014-07-10 v2.0   AVV  Added $departmentnumberException for non-numeric departmentnumbers / companyfields (enabled for IBM).
2014-07-10 v2.0   AVV  Added $locationBasedOnEmployeeID hashtable, to support creating accounts in OUs other as the managed users OU (enabled for employee ID starts with 88, create in supprt\asp).
2014-07-10 v2.0   AVV  Disabled setting the O365 attribute (extentionattribute2 = OE3).
                       On email request of Plamen: 04 July, 2014 11:26 [Plamen] OE3 license should be added by Ivaylo’s script with is running 30 min after this one. You cam disable this section for now.
2014-07-10 v2.0   AVV  Added EmployeeID in errorfile.
2014-07-04 v1.0   AVV  initial script
#>

#-------------------------------------------------------------
#parameters
#-------------------------------------------------------------

#run only from scheduler
Param ([Switch]$Force)

#region Initialize script

#set this only in test/debugmode
#$DebugPreference = "continue"

##read (above) variables from cfg file
Import-CSV ($PSScriptRoot + "\Auto-UAA.cfg") -Delimiter ";" | ForEach-Object { New-Variable -Name $_.Variablename -Value $_.Value -Force -Scope Script }
If ($errorFilePath -notmatch "\\$")  { $errorFilePath += "\" }
If ($inputFilePath -notmatch "\\$")  { $inputFilePath += "\" }
If ($reportFilePath -notmatch "\\$") { $reportFilePath += "\" }
If ($succesFilePath -notmatch "\\$") { $succesFilePath += "\" }

#Hashtable that defines user AD locations
#syntax is regEx = DN

#example: [ordered]@{"^88"="OU=ASP,OU=Support,DC=corp,DC=demb,DC=com";"\w"="OU=Managed Users,DC=corp,DC=demb,DC=com"}
#the example means that a when a employeeId starts with 88, it will be moved to the ASP OU

$locationBasedOnEmployeeID = [ordered]@{"^88"="OU=ASP,OU=Support,DC=corp,DC=demb,DC=com";"\w"="OU=Managed Users,DC=corp,DC=demb,DC=com"}

#expected headers in the inputfile:
$expectedHeaders = @(“employeeID”,”sn”,”givenname”,”streetAddress”,”company”,”department”,”l”,”c”,”co”,
                     ”comment”,”departmentNumber”,”division”,”employeeType”,”extensionAttribute1”,”extensionAttribute3”,”extensionAttribute4”,
                     ”extensionAttribute5”,”extensionAttribute6”,”extensionAttribute8”,”manager”,”pager”,”personalTitle”,”postalCode”,”title”,
                     ”extensionAttribute12”,”description”,”employeenumber”,”physicalDeliveryOfficeName”,"displayname", "ExtensionAttribute11")

#properties used for the new-users hashtable (notice that SamAccountName is excluded, as it is explicit specified in the new-aduser cmdlet)
$propertiesNeededForNewUser = @("name","givenname","sn","userprincipalname","displayname","EmployeeID","description","EmployeeNumber","comment",                                "company","city","l","c","co","countryCode","Department","Devision","employeeType","PostalCode","StreetAddress","Title","MobilePhone","OfficePhone",                                "Fax",”pager”,"personalTitle","departmentnumber","division","Manager","physicalDeliveryOfficeName","extensionAttribute1","extensionAttribute2","extensionAttribute3","extensionAttribute4",
                                "extensionAttribute5","extensionAttribute6","extensionAttribute8","extensionAttribute12", "ExtensionAttribute11")
                                        
#max allowed samaccountlength
$maxSAMAccountLenght = 20

#countrycode hastable is need to convert the ISO county code (c-attribute) to the Numeric Countrycode (countryCode-attribute)
$countrycodeHashTable = @{AF=4;AX=248;AL=8;DZ=12;AS=16;AD=20;AO=24;AI=660;AQ=10;AG=28;AR=32;AM=51;AW=533;AU=36;AT=40;AZ=31;BS=44;BH=48;BD=50;BB=52;BY=112;BE=56;BZ=84;BJ=204;BM=60;BT=64;BO=68;BA=70;BW=72;BV=74;BR=76;VG=92;IO=86;BN=96;BG=100;BF=854;BI=108;KH=116;CM=120;CA=124;CV=132;KY=136;CF=140;TD=148;CL=152;CN=156;HK=344;MO=446;CX=162;CC=166;CO=170;KM=174;CG=178;CD=180;CK=184;CR=188;CI=384;HR=191;CU=192;CY=196;CZ=203;DK=208;DJ=262;DM=212;DO=214;EC=218;EG=818;SV=222;GQ=226;ER=232;EE=233;ET=231;FK=238;FO=234;FJ=242;FI=246;FR=250;GF=254;PF=258;TF=260;GA=266;GM=270;GE=268;DE=276;GH=288;GI=292;GR=300;GL=304;GD=308;GP=312;GU=316;GT=320;GG=831;GN=324;GW=624;GY=328;HT=332;HM=334;VA=336;HN=340;HU=348;IS=352;IN=356;ID=360;IR=364;IQ=368;IE=372;IM=833;IL=376;IT=380;JM=388;JP=392;JE=832;JO=400;KZ=398;KE=404;KI=296;KP=408;KR=410;KW=414;KG=417;LA=418;LV=428;LB=422;LS=426;LR=430;LY=434;LI=438;LT=440;LU=442;MK=807;MG=450;MW=454;MY=458;MV=462;ML=466;MT=470;MH=584;MQ=474;MR=478;MU=480;YT=175;MX=484;FM=583;MD=498;MC=492;MN=496;ME=499;MS=500;MA=504;MZ=508;MM=104;NA=516;NR=520;NP=524;NL=528;AN=530;NC=540;NZ=554;NI=558;NE=562;NG=566;NU=570;NF=574;MP=580;NO=578;OM=512;PK=586;PW=585;PS=275;PA=591;PG=598;PY=600;PE=604;PH=608;PN=612;PL=616;PT=620;PR=630;QA=634;RE=638;RO=642;RU=643;RW=646;BL=652;SH=654;KN=659;LC=662;MF=663;PM=666;VC=670;WS=882;SM=674;ST=678;SA=682;SN=686;RS=688;SC=690;SL=694;SG=702;SK=703;SI=705;SB=90;SO=706;ZA=710;GS=239;SS=728;ES=724;LK=144;SD=736;SR=740;SJ=744;SZ=748;SE=752;CH=756;SY=760;TW=158;TJ=762;TZ=834;TH=764;TL=626;TG=768;TK=772;TO=776;TT=780;TN=788;TR=792;TM=795;TC=796;TV=798;UG=800;UA=804;AE=784;GB=826;US=840;UM=581;UY=858;UZ=860;VU=548;VE=862;VN=704;VI=850;WF=876;EH=732;YE=887;ZM=894;ZW=716}

#company/departmentnumber exception list as regEx (example: "^IBM|^ACME" -> ^ means must start with and | is OR)
$departmentnumberException = "^Accenture|^ACN|^BAKERY ES|^EXTERNAL|^IBM|^IBM-P|^PIMEXTSUP|^SL IAUDIT|^SUPERUSERS|^SYSTEM"

#-------------------------------------------------------------
#some parameters processing and checks
#-------------------------------------------------------------

# build the file paths
$inputFile = $inputFilePath + $inputFileName
$processPath = $PSScriptRoot + "\Process\"
$completedPath = $PSScriptRoot + "\Completed\"

#error settings
$errorFile = $errorFilePath + "Error_ $(get-date -Format "yyyy-MM-dd-HHmm").csv"
$errorArchiveFile = "{0}{1}_{2}.csv" -f $errorFilePath,($inputFileName -replace(".csv$","")),(Get-Date -Format "yyyy-MM-dd-HHmm")
#succespath and generate succesfilename
$successArchiveFile = "{0}{1}_{2}.csv" -f $succesFilePath,($inputFileName -replace(".csv$","")),(Get-Date -Format "yyyy-MM-dd-HHmm")


#create the process and completed folder if not exist already
If(!(test-path $processPath)){New-Item $processPath -ItemType directory >$null}
If(!(test-path $completedPath)){New-Item $completedPath -ItemType directory >$null}

#define log file with the name as the script
$logFile = $MyInvocation.MyCommand.Definition -replace(".ps1$",".log")
$errorLogfile = $MyInvocation.MyCommand.Definition -replace(".ps1$","-Errors.log")
#create the log file if the file does not exits else write today's first empty line as batch separator
If(!(Test-Path $logFile)){ New-Item $logFile -ItemType file >$null }
else { Add-Content -path $logFile -Value "" }
Add-Content -path $logFile -value $((Get-Date -Format "yyyy-MM-dd HH:mm:ss ") + "Script started (" + $env:USERNAME + ")")
#endregion

#region Log functions
#-------------------------------------------------------------
#log functions
#-------------------------------------------------------------
function Write-Log([string]$LogMessage)
{
    $Message = "{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"),$LogMessage
    Add-Content -Path $logFile -Value $Message
    Write-Debug $Message
}

#This function write the error to the logfile and exit the script
function Error-Break([string]$ErrorMessage)
{
    Write-Log($ErrorMessage)
    Write-Log("Script stopped")
    Exit
}
#endregion

#-------------------------------------------------------------
#-------------------------------------------------------------
# Main Script
#-------------------------------------------------------------
#-------------------------------------------------------------

#region Run only from scheduler, exit when started manually
if(!$force.IsPresent)
{
    Write-Error "Due to interaction with other processes this script must only run at the scheduled times. DO NEVER RUN THIS SCRIPT MANUALLY!"
    Error-Break ("Script manually started by: {0}" -f $env:USERNAME )
}
#endregion

#region Read and process input file

        #Check if the input file exists - if not try to find an alternative
        If(-not (Test-Path $inputFile))
        {
            #get all available input files (formatted as <InputfileName>*.csv, so if a date is added it is picked up too.
            $inputFiles = Get-ChildItem ("{0}\{1}*.csv" -f $inputFilePath,($inputFileName -replace(".csv$",""))) | Sort-Object LastWriteTime -Descending
                
            #break when there is no input file available
            If($inputFiles.count -eq 0){ Error-Break("No input file found") }
                
            #if there are multiple input files, always get the last written one (and change the name to the the actual name)
            $inputFileName = ($inputFiles[0]).name
            $inputFile = $inputFilePath + $inputFileName
            if($inputFiles.count -ne 1){ Write-Log ("Multiple ({0}) input files found; the script picked the most recent: {1}" -f ($inputFiles.Count), $inputFileName) }
        }

        # read inputFile or exit if any errors
        Try
        {
            [Array]$users = Import-Csv -path $inputFile -Delimiter ";"
            Write-Log ("{0} read as inputfile" -f $inputFile)
        }
        Catch
        {
            Error-Break("Error in input file: Can't open file ({0})" -f $Error[0])
        }

        #break if inputfile is empty
        If($users.count -eq 0)
        {
            #2014-09-29 KHi: Move even the empty input file on request of Plamen
            Move-Item -Path $inputFile -Destination ($successArchiveFile) -Force
            Error-Break("Error in input file: File is empty")
        }

        #check if all expected headers exist in the inputfile
        $receivedHeaders = ($users[0] | Get-Member -MemberType NoteProperty).name
        $missingHeaders = Compare-Object -ReferenceObject $expectedHeaders -DifferenceObject $receivedHeaders -CaseSensitive:$false -PassThru | Where-Object {$_.sideINdicator -eq "<="}
        if($missingHeaders.count -ne 0)
        {
            Send-MailMessage -From $emailErrorsFrom -To $emailErrorsTo `
                             -Bcc $emailMonitoring `
                             -Subject "Account creation errors $(get-date -Format "yyyy-MM-dd")" `                             -Body ("Missing headers in input file: {0} " -f [string]$missingHeaders ) `                             -SmtpServer $smtpServer

            Error-Break("Error in input file: Header(s) are missing ({0})" -f [string]$missingHeaders )
        }


        ###############################
        #check and create input values#
        ###############################

        
        foreach ($user in $users)
        {
            #$true if account needs to be revived
            add-member -force -InputObject $user -NotePropertyname Reenable -NotePropertyValue $False

            #create error status and error text property (including line number)
            add-member -force -InputObject $user -NotePropertyname Error -NotePropertyValue $False
            add-member -force -InputObject $user -NotePropertyname Errortext -NotePropertyValue "$(([Array]::IndexOf($Users, $User)) + 1);"
            $user.Errortext += $user.employeeID
            $user.Errortext += ";"
            $UserPrincipalName = ""

          #employeeID (Trim leading/trailing spaces and numeric check)
            $user.employeeID = $user.employeeID.trim()
            if($user.employeeID -notmatch "^\d+$")
            {
                $user.Errortext += "EmployeeID not numeric;"
            }

            #Error when Employee ID already exists in AD
            if (($Users[0..([Array]::IndexOf($Users, $User))] | Where-Object { $_.EmployeeID -eq $User.EmployeeID }).Count -lt 2)
            {
                [array]$ADUser = Get-ADUser -LDAPfilter "(employeeID=$($User.employeeID))"
                If($ADUser.Count -gt 0)
                {
                    #EmployeeID is found in AD, but there could be more then 1

                    if( $ADUser.Count -eq 1 )
                    {
                        #Re-enable user
                        $User.Reenable = $True
                        $SAMAccountName = $ADUser.sAMAccountName
                        $UserPrincipalName = $ADUser.userPrincipalName
                    }
                    else
                    {
                        $TempUserName = $ADUser.name -join ", "
                        $User.Errortext += ("EmpoyeeID already exists more than once in AD and can't be re-enabled (found: {0});" -f $TempUserName )
                        Write-Log ("Warning: EmpoyeeID ({1}) already exists more than once in AD and can't be re-enabled (found: {0})" -f ($TempUserName, $User.employeeID) )
                        Remove-Variable TempUserName
                    }
                }
                Remove-Variable ADUser
            }
            else
            {
                $User.Errortext += ("EmpoyeeID already exists in this import file;")
                Write-Log ("Warning: EmpoyeeID ({0}) already exists in this import file" -f ($User.employeeID) )
            }


          #sn (Trim leading/trailing spaces and not empty)
            $user.sn = $user.sn.trim()
            if($user.sn -eq "" -or $user.sn -eq $null)
                { $user.Errortext += "Missing surname (sn);" }

          #givenname (Trim leading/trailing spaces and not empty)
            $user.givenname = $user.givenname.trim()
            if($user.givenname -eq "" -or $user.givenname -eq $null)
                { $user.Errortext += "Missing givenname;" }

          #displayname (Create) if empty
            #for the displayname; place last part of the surname in front, then the givenname and then the remaining part of the surname (then trim spaces)
            #Example: givenname="arnoud", sn="voorst" will result in: "voorst, arnoud"
            #Example: givenname="arnoud", sn="van voorst" will result in: "voorst, arnoud van"
            #Example: givenname="arnoud", sn="van der voorst" will result in: "voorst, arnoud van der"
            if ($user.displayname-eq "" -or $user.sn -eq $null)
                { $User.displayname = ($user.sn.Split(" ")[-1] + ", " + $user.givenname + " " + $($user.sn -replace(("{0}$" -f $user.sn.Split(" ")[-1]),""))).trim() }
            #below line no longer required as property was added to the input file
            #Add-Member -force -InputObject $user -NotePropertyname Displayname -NotePropertyValue $Displayname
           
            
          #SamAccountName
            if (-not $User.Reenable)
            {#Create new account
                #givenname.sn
                #Example: givenname="arnoud", sn="voorst" will result in: "arnoud.voorst"
                #Example: givenname="arnoud", sn="van voorst" will result in: "arnoud.vanvoorst"
                #Example: givenname="arnoud", sn="van der voorst" will result in: "arnoud.vandervoorst"
                $SAMAccountName = $User.GivenName.replace(" ","").ToLower() + "." + $User.sn.replace(" ","").ToLower()

                #if samaccountname is too long use only the first character of the givenname
                if($SAMAccountName.Length -gt $maxSAMAccountLenght)
                {
                    $SAMAccountName = $User.givenname.replace(" ","").substring(0,1).ToLower()+"."+$User.sn.replace(" ","").ToLower()
                    # if samaccount is still too long, throw an error
                    if($SAMAccountName.Length -gt $maxSAMAccountLenght)
                        { $User.Errortext += "SAMAccount $SAMAccountName too long;" }
                }

                #Extend samaccount if it already exists with increasing number up and until 8
                #2015-04-01 v2.6   AVV  Improved SAMAccount/UPN check, not only check is SAMAccount/UPN already exists against AD only but also against not yet created users in memory

                #while ( ((Get-ADUser -Filter {SamAccountname -eq $SAMAccountName}) -ne $null) -and ($SAMAccountName[-1] -ne "9") )
                while ( ( ((Get-ADUser -Filter {SamAccountname -eq $SAMAccountName}) -ne $null) -or ( $SAMAccountName -in $users.SamAccountName )  ) -and ($SAMAccountName[-1] -ne "9") )
                {
                    if ($SAMAccountName -notmatch "\d$") #First attempt (samAccountName1)
                    {
                        $SAMAccountName += "2"
                    }
                    else #Increase last digit from samaccountname
                    {
                        $SAMAccountName = ($SAMAccountName -replace ("\d$",[char]([int]$SAMAccountName[-1] + 1)))
                        #Wite error when 8 attempts have failed
                        If($SAMAccountName[-1] -eq "9")
                            { $user.Errortext += "Too many SAMAccounts already exists ({0});" -f ($SAMAccountName -replace ("\d$", "*")) }
                    }
                }

                #If((Get-ADUser -Filter {SamAccountname -eq $SamAccountName}) -ne $null){$user.ErrorText += "SAMAccount $SamAccountName already exists;"}
            }#Create new account

            Add-Member -force -InputObject $user -NotePropertyname SamAccountName -NotePropertyValue $SAMAccountName

          #userprincipalname (Create)
            if(-not $User.Reenable)
            {
                #givenname.sn
                #Example: givenname="arnoud", sn="voorst" will result in: "arnoud.voorst@demb.com"
                #Example: givenname="arnoud", sn="van voorst" will result in: "arnoud.vanvoorst@demb.com"
                #Example: givenname="arnoud", sn="van der voorst" will result in: "arnoud.vandervoorst@demb.com"
                $UserPrincipalName = $User.GivenName.replace(" ","").ToLower() + "." + $User.sn.replace(" ","").ToLower() + "@jdecoffee.com"

                #Extend UserPrincipalName if it already exists with increasing number up and until 9
                #2015-04-01 v2.6   AVV  Improved SAMAccount/UPN check, not only check is SAMAccount/UP already exists agains AD only but also agains not yet created users in memory
                $CountUPN = 2
                while ( ( ((Get-ADUser -Filter { UserPrincipalName -eq $UserPrincipalName }) -ne $null) -or ( $UserPrincipalName -in $users.UserPrincipalName )) -and ($CountUPN -le 9) )
                {
                    $UserPrincipalName = $UserPrincipalName -replace '\d*@jdecoffee.com', "$CountUPN@jdecoffee.com"
                    $CountUPN++
                }

                if ( $CountUPN -gt 9 )
                {
                    #Wite error when 9 attempts have failed
                    $User.Errortext += "Too many UserPrincipalNames already exists ({0});" -f ($UserPrincipalName)
                }
            }
            Add-Member -force -InputObject $User -NotePropertyname UserPrincipalName -NotePropertyValue ($UserPrincipalName)

          #company (Should start with four digits or in the exception list)
            if(($User.company -notmatch "^\d{4}") -and ($User.company -notmatch $departmentnumberException))
                { $user.Errortext += "First four characters of companyfield not numeric and not in exception list;" }

          #departmentnumber (Should be four digits or in the exception list)
            if(($user.departmentNumber -notmatch "^\d{4}$") -and ($User.departmentNumber -notmatch $departmentnumberException))
                { $user.Errortext += "Departmentnumber does not contain four digits and not in exception list;" }
          
          #Set country code to 'CK' when no country code is provided
            if($User.c -eq '' -or $User.c -eq $null)
            {
                $User.c = 'CK'
                $User.co = 'Cook Islands'
            }

          #add countryCode based on c
            Add-Member -force -InputObject $user -NotePropertyname countryCode -NotePropertyValue $countrycodeHashTable[$user.c]

          #check manager DN (if not exist the result is $null)
          If(($User.manager -ne "") -and ($User.manager -ne $null))
            { $user.manager = (Get-ADUser -LDAPFilter "(DistinguishedName=$($user.manager))").DistinguishedName }

          <## THIS PART IS REMOVED ON EMAIL REQUEST OF PLAMEN - 04 July, 2014 11:26
          [Plamen] OE3 license should be added by Ivaylo’s script with is running 30 min after this one. You cam disable this section for now 
          
              #set O365 licence (extensionAttribute2 = OE3) if comment = active, extensionAttribute5 = employee and emplyeetype = active employee)
              If(($User.comment -eq "active") -and ($User.extensionAttribute5 -eq "Employee") -and ($user.employeeType -eq "Active Employee"))
              {
                add-member -force -InputObject $user -NotePropertyname extensionAttribute2 -NotePropertyValue "OE3"
              }

          ##>

          #create hastable propery that will be used for new-aduser
            Add-Member -force -InputObject $user -NotePropertyname NewADUserHashtable -NotePropertyValue (New-Object Hashtable)

            #build new-aduser hastable only with properties in $propertiesNeededForNewUser and when the value is not $null
            #note: new-aduser doesnot accept $null or ""
            $user.PSObject.Properties | Where-Object {($propertiesNeededForNewUser -contains $_.name) -and ($_.value -ne $null) -and ($_.value -ne "")} |
                ForEach-Object { $user.NewADUserHashtable.add($_.name, $_.value) }

           #Write errorcode to userobject
           #Error when errortext when there is text after <linenumber;employeeid;>
            if($user.Errortext -notmatch "^\d+;\d+;$")
                { $user.Error = $True }

        }

        ### at this point we have all users populated with the necessary information (including error, errortext, NewADUserHashTable etc.)

        #2014-08-13 KHi: Moved the error check part to the end of the script to overcome issues

# Get DEMBDCSWA* domaincontroller availability (This will add the 'Available' property to the object)
# Availability is based on the existence of the host in the AD (then the DC is up and the AD services are running)
$Domaincontrollers = Get-ADComputer -SearchBase "OU=domain controllers,DC=corp,DC=demb,DC=com" -Filter {name -like "DEMBDCRS*"} |
                     Select-Object -Property *,@{n='Available';e={(Get-ADComputer -ldapfilter "(name=$($env:Computername))" -Server $_.name) -ne $null} }
#endregion

#region Create new users
########################
# Create the new users #
########################

foreach ($User in ($Users | Where-Object { -not $_.error -and -not $_.Reenable } ))
    {
        $WorkDC = ($Domaincontrollers | Get-Random).name
        Write-Debug ("Work DC: {0}" -f $WorkDC)

        try
            {
                #Get the User OU based on the regEx in the keys of the $locationBasedOnEmployeeID Hashtable (defined in the beginning of this script)
                #Notice the Sort-Object to get the correct processing order
                $UserOU = $locationBasedOnEmployeeID[($locationBasedOnEmployeeID.Keys | Where-Object {$user.employeeID -match $_ } | Select-Object -First 1)]

                New-ADUser -name $user.SamAccountName -Path $userOU -OtherAttributes $user.NewADUserHashtable -Server $WorkDC
                #add timestamp to EA15 and logfile
                $now = get-date -Format "yyyy-MM-dd HH:mm:ss"
                Set-ADUser -Identity $user.SamAccountName -Replace @{extensionAttribute15="Autogenerated at $now"} -Server $WorkDC
                Add-Member -force -InputObject $user -NotePropertyname AccountStatus -NotePropertyValue "Account created on $WorkDC at $now"
                Write-Log -LogMessage "$($user.SamAccountName): $($User.AccountStatus)"
                Write-Debug -Message "$($user.SamAccountName): $($User.AccountStatus)"

                #create file in the process folder
                New-Item -Name "$($user.SamAccountName).mail.txt" -Path $processPath -ItemType file > $null
				Add-Content -path ("{0}\{1}.mail.txt" -f $processPath, $User.SamAccountName) -value ("{0} Created by Auto-UAA" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
            }
        catch
            {
                Add-Member -force -InputObject $user -NotePropertyname AccountStatus -NotePropertyValue "Error account creation on $WorkDC at $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") $($Error[0])"
                Write-Log -LogMessage "$($user.SamAccountName): $($User.AccountStatus)"
            }
    }
#endregion

#region Revieve users
###################
# Re-enable users #
###################

foreach ($User in ($Users | Where-Object { -not $_.error -and $_.Reenable } ))
{
    $WorkDC = ($Domaincontrollers | Get-Random).name
    Write-Debug ("Work DC: {0}" -f $WorkDC)

    try
    {
        #Get current data
        $ADUser = Get-ADUser $User.SAMAccountName -Properties memberOf, ExtensionAttribute2, ExtensionAttribute15

        #Remove from the to be deleted group(s)
        ($ADUser).MemberOf | 
            Where-Object {$_ -Match 'CN=To-Be-Deleted-\w{3},OU=Disabled Accounts,OU=To be deleted,OU=Support,DC=corp,DC=demb,DC=com'} |
            Remove-ADGroupMember -Members $User.SAMAccountName -Confirm:$false

        Set-ADUser -Identity $User.SamAccountName -Replace ($User.NewADUserHashtable) -Server $WorkDC

        if ( -Not $ADUser.Enabled )
        {
                
            #Restore Office365 license (extensionAttribute2) from extensionAttribute15 or ExtensionAttribute2

            if ( $ADUser.ExtensionAttribute2 -like "!*" )
            {
                if ( $ADUser.ExtensionAttribute2 -eq "!" )
                {
                    Set-ADUser -Identity $User.SamAccountName -Clear ExtensionAttribute2 -Server $WorkDC
                }
                else
                {
                    Set-ADUser -Identity $User.SamAccountName -Replace @{ExtensionAttribute2 = $ADUser.ExtensionAttribute2.Replace('!', '')} -Server $WorkDC
                }
            }
            else
            {
                #This old part can be deleted after July 25, 2016 due to backwards compatability
                $ExA2 = (($ADUser.extensionAttribute15 -split '(Left Company\s*-\s*DO NOT ENABLE.+)(\s+EXA2=\w+)(.*)')[2] -split '=')[1]
                if ($ExA2 -ne $null)
                {
                    Set-ADUser -Identity $User.SamAccountName -Replace @{extensionAttribute2=$ExA2} -Server $WorkDC
                }
            }
            Remove-Variable ExA2, ADUser

            #Add timestamp to extensionAttribute15 and logfile
            $Now = get-date -Format "yyyy-MM-dd HH:mm:ss"
            Set-ADUser -Identity $User.SamAccountName -Replace @{extensionAttribute15="Auto re-enabled at $Now"} -Server $WorkDC
            Add-Member -force -InputObject $User -NotePropertyname AccountStatus -NotePropertyValue "Account re-enabled on $WorkDC at $Now"
        
            Write-Log -LogMessage "$($User.SamAccountName): $($User.AccountStatus)"
            Write-Debug -Message "$($User.SamAccountName): $($User.AccountStatus)"

            #create file in the process folder
            New-Item -Name "$($User.SamAccountName).reenablemail.txt" -Path $processPath -ItemType file | Out-Null
        } #real re-enabling of account
        else
        {
            #Add timestamp to extensionAttribute15 and logfile
            $Now = get-date -Format "yyyy-MM-dd HH:mm:ss"
            Set-ADUser -Identity $User.SamAccountName -Replace @{extensionAttribute15="Auto re-enabled (already enabled) at $Now"} -Server $WorkDC
            Add-Member -force -InputObject $User -NotePropertyname AccountStatus -NotePropertyValue "Auto re-enabled (already enabled) on $WorkDC at $now"
        
            Write-Log -LogMessage "$($User.SamAccountName): $($User.AccountStatus)"
            Write-Debug -Message "$($User.SamAccountName): $($User.AccountStatus)"

            #create file in the process folder
            New-Item -Name "$($user.SamAccountName).report.txt" -Path $processPath -ItemType file | Out-Null
        }
    }
    catch
    {
        Add-Member -force -InputObject $user -NotePropertyname AccountStatus -NotePropertyValue "Error account re-enable on $WorkDC at $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") $($Error[0])"
        Write-Log -LogMessage "$($user.SamAccountName): $($User.AccountStatus)"
    }
}
#endregion

#region Finalize process

#if there are no errors in the input file move to inputfile to the archive folder
If(($users | Where-Object {$_.Error -eq $true}).Count -eq 0)
{
    Write-Log("Successfully processed {0}" -f ($inputFileName))
    Move-Item -Path $inputFile -Destination ($successArchiveFile) -Force
}
else
{
    #Write users with an error to file
    Write-Log("Warning: Input issues are found in {0}" -f ($inputFileName))

    ($Users | Where-Object {$_.error}).errortext | Out-File -FilePath $errorFile -Encoding ascii

    #2014-08-13 KHi: Added logging for flow check
    if(!(Test-Path -Path $errorFile)) { Write-Log("Error creating error file ({0})" -f $errorFile) }

    try
    {
        Send-MailMessage -From $emailErrorsFrom -To $emailErrorsTo `
                         -Bcc $emailMonitoring `
                         -Subject "Account creation errors $(Get-Date -Format "yyyy-MM-dd")" `                         -Attachments $errorFile `                         -SmtpServer $smtpServer
        Write-Log("Sent mail to {0} (subject 'Account creation errors {1}') with {2} as attachment," -f ($emailErrorsTo, (Get-Date -Format "yyyy-MM-dd"), $errorFile))
    }
    catch
    {
        Write-Log("Error sending message to {0}: {1}" -f ($emailErrorsTo, $Error[0].Exception))
    }

    #move inputfile to error folder, was later changed to the success archive folder on request
    Move-Item -Path $inputFile -Destination ($successArchiveFile) -Force
}

#2014-08-13 KHi: Added logging for flow check
if((Test-Path -Path $successArchiveFile))
{
    Write-Log("Moved input file ({0}) to success archive {1}" -f ($inputFileName, $successArchiveFile))
}
else
{
    Write-Log("Error moving input file ({0}) to success archive {1}" -f ($inputFileName, $successArchiveFile))
}
#endregion

## Exit script normally
Write-Log("Script ended normally")
