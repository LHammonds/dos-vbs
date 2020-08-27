Option Explicit
'#############################################################
'## Name         : Archive_IIS_Logs.vbs
'## Version      : 1.0
'## Date         : 2013-01-11
'## Author       : LHammonds
'## Purpose      : Archive/Delete IIS Log files
'## Requirements : Windows Script Host v5.6 (CSCRIPT.EXE), 7-Zip, Email Server
'## Output       : Email Notification, Log file
'######################## CHANGE LOG #########################
'## DATE       VER WHO WHAT WAS CHANGED
'## ---------- --- --- ---------------------------------------
'## 2013-01-11 1.0 LTH  Created program.
'#############################################################
'## The ArchiveLogFiles function takes three parameters:
'## "Path to log dir"
'## "Compress log files older than n days and delete the original files"
'## "Delete compressed log files older than n days"
'## Note that the function runs through subfolders recursively, so if
'## the same log retention should be used on a whole log folder tree
'## structure, only one call with the root log folder is needed.
'## Additional calls with specific subfolders can then be made to have
'## shorter retentions on those.

WScript.Timeout = 82800
Const OverwriteExisting = True
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

DIM str_log
DIM obj_fso
DIM obj_log

Set obj_fso = CreateObject("Scripting.FileSystemObject")
str_log = "C:\Apps\Archive_IIS_Logs.log"
Set obj_log = obj_fso.OpenTextFile(str_log, ForAppending, True)
obj_log.Writeline(f_timestamp & " - ** Script Start **")

'ArchiveLogFiles "D:\Logfiles", 30, 180
'ArchiveLogFiles "D:\Logfiles\W3SVC1", 14, 30
'ArchiveLogFiles "D:\Logfiles\W3SVC243", 5, 30
'ArchiveLogFiles "D:\Logfiles\SMTPSVC1", 7, 60
'** Compress then delete log files older than 10 days and delete archives older than 30 days.
ArchiveLogFiles "D:\inetpub\LogFiles", 30, 360

obj_log.Writeline(f_timestamp & " - ** Script End **")
obj_log.Close
Set obj_log = Nothing

'######################
'## FUNCTION SECTION ##
'######################

Function ArchiveLogFiles(strLogPath, intZipAge, intDelAge)
  DIM obj_fsoCheck
  DIM objFolder
  DIM objSubFolder
  DIM objFile
  DIM objWShell
  DIM obj_logfso
  Set objWShell = CreateObject("WScript.Shell")
  Set obj_logfso = CreateObject("Scripting.FileSystemObject")
  Set obj_fsoCheck = CreateObject("Scripting.FileSystemObject")
  If Right(strLogPath, 1) <> "\" Then
    strLogPath = strLogPath & "\"
  End If
  If obj_logfso.FolderExists(strLogPath) Then
    Set objFolder = obj_logfso.GetFolder(strLogPath)
      For Each objSubFolder in objFolder.subFolders
        '** Process each sub-folder. **
        ArchiveLogFiles strLogPath & objSubFolder.Name, intZipAge, intDelAge
      Next
      obj_log.Writeline("  Processing folder: " & strLogPath)
      For Each objFile in objFolder.Files
        If (InStr(objFile.Name, "ex") > 0) _
          And (Right(objFile.Name, 4) = ".log") Then
          If DateDiff("d",objFile.DateLastModified,Date) > intZipAge Then
            obj_log.Writeline("  " & objFile.name & " -> " & Left(objFile.Name,Len(objFile.name)-3) & "zip")
            objWShell.Run "7za.exe a -tzip """ & strLogPath & Left(objFile.Name,Len(objFile.Name)-3) & "zip"" """ & strLogPath & objFile.Name & """", 7, true
            If obj_fsoCheck.FileExists(strLogPath & _
              Left(objFile.Name,Len(objFile.Name)-3) & "zip") And _
              (obj_fsoCheck.FileExists(strLogPath & objFile.Name)) Then
                obj_log.Writeline("  Deleting " & objFile.Name)
                obj_fsoCheck.DeleteFile(strLogPath & objFile.Name)
            End If
          End If
        ElseIf (InStr(objFile.Name, "ex") > 0) _
          And (Right(objFile.Name, 4) = ".zip") Then
          If DateDiff("d",objFile.DateLastModified,Date) > intDelAge Then
            obj_log.Writeline("  Deleting " & objFile.name)
            obj_fsoCheck.DeleteFile(strLogPath & objFile.Name)
          End If
        End If
      Next
    Set obj_logfso = Nothing
    Set obj_fsoCheck = Nothing
    Set objFolder = Nothing
    Set objWShell = nothing
  End If
End Function  '## ArchiveLogFiles() ##

Function f_email (pstr_subject, pstr_body)
  DIM obj_email
  Set obj_email = CreateObject("CDO.Message")
  obj_email.From = "webmaster@mydomain.com"
  obj_email.To = "john.doe@mydomain.com,jane.doe@mydomain.com"
  obj_email.Subject = pstr_subject
  obj_email.Textbody = pstr_body
  obj_email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
  obj_email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")="srv-mail"
  obj_email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
  obj_email.Configuration.Fields.Update
  obj_email.Send
  Set obj_email = Nothing
End Function  '## f_email() ##

Function f_timestamp()
  DIM lstr_timestamp
  DIM lstr_temp
  lstr_timestamp = DatePart("yyyy",Now) & "-"
  lstr_temp = DatePart("m",Now)
  If Len(lstr_temp) = 1 Then
    lstr_timestamp = lstr_timestamp & "0" & lstr_temp & "-"
  Else
    lstr_timestamp = lstr_timestamp & lstr_temp & "-"
  End If
  lstr_temp = DatePart("d",Now)
  If Len(lstr_temp) = 1 Then
    lstr_timestamp = lstr_timestamp & "0" & lstr_temp & " "
  Else
    lstr_timestamp = lstr_timestamp & lstr_temp & " "
  End If
  lstr_temp = DatePart("h",Now)
  If Len(lstr_temp) = 1 Then
    lstr_timestamp = lstr_timestamp & "0" & lstr_temp & ":"
  Else
    lstr_timestamp = lstr_timestamp & lstr_temp & ":"
  End If
  lstr_temp = DatePart("n",Now)
  If Len(lstr_temp) = 1 Then
    lstr_timestamp = lstr_timestamp & "0" & lstr_temp & ":"
  Else
    lstr_timestamp = lstr_timestamp & lstr_temp & ":"
  End If
  lstr_temp = DatePart("s",Now)
  If Len(lstr_temp) = 1 Then
    lstr_timestamp = lstr_timestamp & "0" & lstr_temp
  Else
    lstr_timestamp = lstr_timestamp & lstr_temp
  End If
  f_timestamp = lstr_timestamp
End Function '## f_timestamp() ##
