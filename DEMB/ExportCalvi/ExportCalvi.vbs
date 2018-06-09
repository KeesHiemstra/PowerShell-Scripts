'=================================================================
'
' Script to read information from the Active Directory and create
' a csv file to be transfered to the Calvi-system
'
' Note: In case there is an error detected within the attribute
'       the field is filled with "Error: ", the error and the
'       name of the attribute
'
' Author: Kees Hiemstra
' version 1.7
'
' Changes according to Aldea 1060160: R87818/D87509 Adjust RFC: 79281 - Change adress receipient Interface
'
' Author: Kees Hiemstra
' version 1.6
'
' Remove space from phone number(s) as per Aldea 915487: R84653/D84652 Adjust RFC 79281
'
' Author: Luuk Hansma
' version: 1.5
' 
' Modified the output of the costcentre by joining the extensionAttribute1 and extensionAttribute3
'
' Author: Luuk Hansma
' version: 1.4
' 
' Added logging to ExportCalvi.log
'
' version: 1.3
'
' userPrincipalName changed to sAMAccountname
' The otherMobile number was missing in the output, changed the handling 
' due to the fact that the othermobile number is saved in a multi-valued string
'
' Version: 1.2
'
' Modifications as result from the first test
'
' Version: 1.1
'
' Changed to outputfile first to a tmp file and last action is move the file to csv
'
' Version: 1.0
'
'=================================================================
Option Explicit

Dim strOU, strFile, oLog, oFSO

'=================================================================
' Definitions for the script
'
strOU = "OU=Managed Users,DC=corp,DC=demb,DC=com"
strFile = "D:\Scripts\ExportCalvi\" & DatePart("YYYY", Date) & Right("0" & DatePart("M", Date), 2) & Right("0" & DatePart("D", Date), 2) & " " & Right("0" & DatePart("h", Now), 2) & Right("0" & DatePart("n", Now), 2)

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oLog = oFSO.OpenTextFile ("D:\Scripts\ExportCalvi\ExportCalvi.log", 8, True)
oLog.WriteLine "================================================================="
oLog.WriteLine "Started " & Now

Call Main

'Dim sFTPUser, sFTPPass, sFTPHost, sFTPDir

'Path to file or folder to upload
'sFTPUser = "svc.dmzftpgdas"
'sFTPPass = "Steelw3dg3"
'sFTPHost = "10.142.152.12"
'sFTPDir = ""

'Call FTPUpload(strFile & ".csv")

Dim sSMTPServer, sFrom, sTo, sCC, sSubject, sMessage

sSMTPServer = "smtp.corp.demb.com"
sFrom = "ExportCalvi.dembmcis168@demb.com"
sTo = " adexport.telecom@demb.com"
sCC = ""
sSubject = "Interface from Active Directory to Calvi"
sMessage = "Attached you'll find the Active Directory data for Calvi."

Call SendMail(strFile & ".csv")

'Remove the file
If oFSO.FileExists(strFile & ".csv") Then
  oLog.WriteLine "File " & strFile & ".csv still exists so try to remove"
  oFSO.DeleteFile oFSO.getFile(strFile & ".csv")
End If

oLog.WriteLine "Stopped " & Now

WScript.Quit
'=================================================================

Sub Main
  Dim objConnection, objCommand, objRecordset, objFso, objFile
  
  '=================================================================
  ' Create connection to AD
  '
  Set objConnection = CreateObject("ADODB.Connection")
  objConnection.Open "Provider=ADsDSOObject;"

  '=================================================================
  ' Create command
  '
  Set objCommand = CreateObject("ADODB.Command")
  objCommand.ActiveConnection = objConnection
  objCommand.Properties("Page Size") = 1000

  '=================================================================
  ' Execute command to get all users in OU and co=Netherlands
  '
  objCommand.CommandText = _
    "<LDAP://" & strOU & ">;" & _
    "(&(objectclass=user)(objectcategory=person)(co=Netherlands));" & _
    "departmentNumber,extensionAttribute1,extensionAttribute3,department,employeeID,givenname,sn,personalTitle,mail,employeeType,title,physicalDeliveryOfficeName,streetAddress,postalCode,l,co,sAMAccountname,extensionAttribute3,mobile,otherMobile,telephoneNumber,otherTelephone;subtree"
  oLog.WriteLine objCommand.CommandText
  Set objRecordset = objCommand.Execute
  oLog.WriteLine "Recordcount: " & objRecordset.Recordcount

  '=================================================================
  ' Create result file
  '
  Set objFso = CreateObject("Scripting.FileSystemObject")
  Set objFile = objFso.CreateTextFile(strFile & ".tmp", True)

  '=================================================================
  ' Generate the csv-file
  '
  Write_CSV objFile, objRecordset
  
  objFile.Close
  Set objFile = Nothing
  
  objFso.MoveFile strFile & ".tmp", strFile & ".csv"
End Sub

'=================================================================
'
' Output the recordset retrieved from the AD to a csv-file
'
'=================================================================
Sub Write_CSV(oFil, oRst)
  If oRst.EOF Then Exit Sub

  Dim bOut, bSep
  Dim iFld, iCnt, iErr, iNot
  Dim oFld, oReg
  Dim sFld, sSep, sOut, sZer
  Dim aFld
  
  sSep = ";"
  '===========================================
  ' Use RegEXp to check for invalid characters
  Set oReg = New RegExp
  oReg.Pattern = "\r|\n|" & sSep

  ' The number in the array is used to format the output
  ' Based on the request by Piet OTter 26-Nov-2012
  aFld = Array(6, 71, 72, 0, 8, 0, 0, -1, -1, 1, -1, -1, -1, 0, 0, 0, 0, -1, -1, 0, 0, 0, 0, -1, -1, -1, -1, -1, -1, 0, 9, 3, 4, 1, 2)
  
  oFil.Write "OpCo;cost center;department;employee number;first name;last name;maiden name;inititals;title (front);middle name (belongs to last name);title (back);male/female;email address;contract type;role;site;building;room;postal address;ZIP;City;Country;Visitors address;ZIP;City;Country;commencing date;termination date;username;reports to employee number;mobile phone;2nd mobile phone;business phone;2nd business phone"
  oFil.Write vbNewLine
  
  iErr = 0
  Do Until oRst.EOF
    iFld = 0
    sOut = ""
    bOut = False
    For iCnt = 0 To UBound(aFld)
      If aFld(iCnt) >= 0 Then
        If iFld <= oRst.Fields.Count Then
          bSep = True
          oFld = oRst.Fields(iFld).Value
          '==============================================
          ' Check whether to data is filled and not empty
          If VarType(oFld) > 1 Or aFld(iCnt) >= 10 Then
            Select Case aFld(iCnt)
            Case 0
              '=========================================
              ' Check whether to data is an array or not
              If VarType(oFld) < 8192 Then
                If oReg.Test(oFld) Then
                  sOut = sOut & "Error: Illegal characters in attribute " & oRst.Fields(iFld).Name
                Else
On Error Resume Next
                  sOut = sOut & "" & oFld
If Err <> 0 Then sOut = sOut & "Error: " & Err.Description & " in attribute " & oRst.Fields(iFld).Name
On Error GoTo 0
                End If
              Else
                For Each sFld In oFld
On Error Resume Next
                  sOut = sOut & sFld
If Err <> 0 Then sOut = sOut & "Error in attribute " & oRst.Fields(iFld).Name
On Error GoTo 0
                Next
              End If
            Case 1
              If VarType(oFld) = 8 Then
								oFld = Replace(oFld, " ", "")
                If InStr(oFld, " ") = 0 Then
                  If Left(oFld, 8) = "+3120558" Then
                    sOut = sOut & oFld
                    bOut = True
                  End If
                  If Left(oFld, 5) = "+3130" Then
                    sOut = sOut & Right(oFld, 4)
                    bOut = True
                  End If
                  If Left(oFld, 6) = "+31513" Then
                    sOut = sOut & "3" & Right(oFld, 3)
                    bOut = True
                  End If
                End If
              End If
            Case 2
              If VarType(oFld) >= 8192 Then
                For Each sFld In oFld
									sFld = Replace(sFld, " ", "")
                  If InStr(sFld, " ") = 0 Then
										If Left(sFld, 8) = "+3120558" Then
											sOut = sOut & sFld
											bOut = True
										End If
                    If Left(sFld, 5) = "+3130" Then
                      sOut = sOut & Right(sFld, 4)
                      bOut = True
                    End If
                    If Left(sFld, 6) = "+31513" Then
                      sOut = sOut & "3" & Right(sFld, 3)
                      bOut = True
                    End If
                  End If
                Next
              End If
            Case 3
              If VarType(oFld) = 8 Then
								oFld = Replace(oFld, " ", "")
                If InStr(oFld, " ") = 0 Then
                  If Left(oFld, 4) = "+316" Then
                    sOut = sOut & "0" & Right(oFld, 9)
                    bOut = True
                  End If
                End If
              End If
            Case 4
              If VarType(oFld) >= 8192 Then
                For Each sFld In oFld
									sFld = Replace(sFld, " ", "")
                  If InStr(sFld, " ") = 0 Then
                    If Left(sFld, 4) = "+316" Then
                      sOut = sOut & "0" & Right(sFld, 9)
                      bOut = True
                    End If
                  End If
                Next
              End If
            Case 5
              If VarType(oFld) = 8 Then
                If InStr(oFld, "@") Then sOut = sOut & LCase(Left(oFld, InStr(oFld, "@") - 1))
              End If
            Case 6
              '=========================================
              ' Check whether to data is an array or not
              If VarType(oFld) < 8192 Then
                If oReg.Test(oFld) Then
                  sOut = sOut & "Error: Illegal characters in attribute " & oRst.Fields(iFld).Name
                Else
On Error Resume Next
                  sZer = "" & oFld
If Err <> 0 Then sOut = sOut & "Error: " & Err.Description & " in attribute " & oRst.Fields(iFld).Name
On Error GoTo 0
                End If
              Else
                For Each sFLd In oFld
On Error Resume Next
                  sZer = "" & sZer & sFLd
If Err <> 0 Then sOut = sOut & "Error in attribute " & oRst.Fields(iFld).Name
On Error GoTo 0
                Next
              End If
              sOut = sOut & Right("0000" & sZer, 4)
            Case 71
              '=========================================
              ' Check whether to data is an array or not
              If VarType(oFld) < 8192 Then
On Error Resume Next
                sZer = oFld
If Err <> 0 Then sZer = ""
On Error GoTo 0
              End If
              sOut = sOut & Right("00000000" & sZer, 8) & "\"
              bSep = False
            Case 72
              sZer = ""
              '=========================================
              ' Check whether to data is an array or not
              If VarType(oFld) < 8192 Then
On Error Resume Next
                sZer = oFld
If Err <> 0 Then sZer = ""
On Error GoTo 0
              End If
              sOut = sOut & Right("00000000" & sZer, 8)
            Case 7, 8
              '=========================================
              ' Check whether to data is an array or not
              If VarType(oFld) < 8192 Then
                If oReg.Test(oFld) Then
                  sOut = sOut & "Error: Illegal characters in attribute " & oRst.Fields(iFld).Name
                Else
On Error Resume Next
                  sZer = oFld
If Err <> 0 Then sOut = sOut & "Error: " & Err.Description & " in attribute " & oRst.Fields(iFld).Name
On Error GoTo 0
                End If
              Else
                For Each sFLd In oFld
On Error Resume Next
                  sZer = sZer & sFLd
If Err <> 0 Then sOut = sOut & "Error in attribute " & oRst.Fields(iFld).Name
On Error GoTo 0
                Next
              End If
              sOut = sOut & Right("00000000" & sZer, 8)
            Case 9
              '=========================================
              ' Check whether to data is an array or not
              If VarType(oFld) < 8192 Then
                If oReg.Test(oFld) Then
                  sOut = sOut & "Error: Illegal characters in attribute " & oRst.Fields(iFld).Name
                Else
On Error Resume Next
                  sZer = oFld
If Err <> 0 Then sOut = sOut & "Error: " & Err.Description & " in attribute " & oRst.Fields(iFld).Name
On Error GoTo 0
                End If
              Else
                For Each sFLd In oFld
On Error Resume Next
                  sZer = sZer & sFLd
If Err <> 0 Then sOut = sOut & "Error in attribute " & oRst.Fields(iFld).Name
On Error GoTo 0
                Next
              End If
              If sZer <> 0 Then sOut = sOut & Right("00000000" & sZer, 8)
            End Select
          End If
          iFld = iFld + 1
        End If
      End If
      If bSep Then sOut = sOut & sSep
    Next
    If bOut Then
On Error Resume Next
      oFil.Write sOut
      If Err = 0 Then
        oFil.Write vbNewLine
      Else
        iErr = iErr + 1
      End If
On Error GoTo 0
    Else
      iNot = iNot + 1
    End If
    oRst.MoveNext
  Loop
  If iErr > 0 Then oLog.WriteLine "Error count: " & iErr
  If iNot > 0 Then oLog.WriteLine "No phone count: " & iNot
End Sub

Sub FTPUpload(path)
  Dim oShell, oFSO, oFTP, oFile, oFolder, oItem
  Dim iWaitTime, sFTP, sParent
  
  Set oShell = CreateObject("Shell.Application")
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  
  On Error Resume Next
  'Copy Options: 16 = Yes to All
  Const copyType = 16
  'FTP Wait Time in ms
  iWaitTime = 80000
  
  sFTP = "ftp://" & sFTPUser & ":" & sFTPPass & "@" & sFTPHost & sFTPDir
  Set oFTP = oShell.NameSpace(sFTP)
  'Make new folder on FTP site
  'oFTP.NewFolder "FTP Backup"
  
  'Upload single file       
  If oFSO.FileExists(path) Then
    Set oFile = oFSO.getFile(path)
    sParent = oFile.ParentFolder
    Set oFolder = oShell.NameSpace(sParent)
    Set oItem = oFolder.ParseName(oFile.Name)
    oLog.WriteLine "Uploading file " & oItem.Name & " to " & sFTP
    oFTP.MoveHere oItem, copyType
  End If
  
  'Upload all files in folder
  If oFSO.FolderExists(path) Then
    'Code below can be used to upload entire folder
    Set oFolder = oShell.NameSpace(path)
    oLog.WriteLine "Uploading folder " & path & " to " & sFTP
    oFTP.MoveHere oFolder.Items, copyType
  End If
  
  If Err.Number <> 0 Then
    oLog.WriteLine "Error: " & Err.Description
  End If
  
  WScript.Sleep iWaitTime
End Sub

'------------------------------------------------------------------------------------------
' Mail the export
'------------------------------------------------------------------------------------------
Sub SendMail(sFile)
  Dim oCon, oFlds, oMsg
  
  Set oCon = CreateObject("CDO.Configuration")
  Call oCon.Load(-1)     ' CDO Source Defaults
  
  Set oFlds = oCon.Fields
  With oFlds
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = sSMTPServer
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    .Update
  End With
  
'On Error Resume Next
  Set oMsg = CreateObject("CDO.Message")
  With oMsg
    Set .Configuration = oCon
    .To = sTo
    .CC = sCC
'    .BCC = "Kees.Hiemstra@hp.com"
    .From = sFrom
    .Subject = sSubject
    .HtmlBody = sMessage
    .AddAttachment sFile
    .Send
  End With
  If Err <> 0 Then
    oLog.WriteLine Err.Description
  Else
    oLog.WriteLine "Mail sent"
  End If
End Sub

