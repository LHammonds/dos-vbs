Option Explicit
'#############################################################
'## Name         : Check_App_Processes.vbs
'## Version      : 1.0
'## Date         : 2005-01-27
'## Author       : LHammonds
'## Purpose      : Ensure processes are running.
'## Requirements : Windows Script Host v5.6 (CSCRIPT.EXE), Email Server
'## Output       : Email Notification, Log file
'######################## CHANGE LOG #########################
'## DATE       VER WHO WHAT WAS CHANGED
'## ---------- --- --- ---------------------------------------
'## 2005-01-17 1.0 LTH  Created program.
'#############################################################

Const OverwriteExisting = True
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
DIM str_log
DIM obj_fso
DIM obj_log

str_log = "C:\Apps\Check_App_Processes.log"

Set obj_fso = CreateObject("Scripting.FileSystemObject")
Set obj_log = obj_fso.OpenTextFile(str_log, ForAppending, True)
obj_log.Writeline(f_timestamp & " - ** Script Start **")
f_check_process "srv-app", "Rcv_Sch.exe", "Schedule Interface"
f_check_process "srv-app", "Rcv_Dem.exe", "Demographics Interface"
f_check_process "srv-app", "Rcv_Lab.exe", "Lab Interface"
obj_log.Writeline(f_timestamp & " - ** Script End **")
obj_log.Close
Set obj_log = Nothing

'######################
'## FUNCTION SECTION ##
'######################

Function f_check_process (pstr_server, pstr_process, pstr_name)
  DIM obj_WMI_Service
  DIM col_ProcessList
  DIM obj_process
  DIM str_body
  DIM int_count
  int_count = 0
  Set obj_WMI_Service = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & pstr_server & "\root\cimv2")
  Set col_ProcessList = obj_WMI_Service.ExecQuery("Select * from Win32_Process Where Name = '" & pstr_process & "'")
  For Each obj_Process in col_ProcessList
    '## If there is one process, then we have no problem. ##
    int_count = int_count + 1
  Next
  Set obj_Process = Nothing
  Set col_ProcessList = Nothing
  Set obj_WMI_Service = Nothing
  str_body = ""
  If int_count = 0 Then
    '## Interface is not running. (BAD) ##
    obj_log.WriteLine(f_timestamp & " - " & pstr_name & " is not running! (" & pstr_process & ")")
    str_body = str_body & pstr_name & " is not running on " & pstr_server & vbNewLine & vbNewLine
    str_body = str_body & "The process called " & pstr_process & " could not be found running in memory." & vbNewLine
    f_email pstr_name & " is not running!", str_body
  ElseIf int_count = 1 Then
    '## Only one Interface is running. (GOOD) ##
  ElseIf int_count > 1 Then
    '## More than one Interface is running. (BAD) ##
    obj_log.WriteLine(f_timestamp & " - " & int_count & " instances of " & pstr_name & " is running on " & pstr_server & ". (" & pstr_process & ")")
    str_body = str_body & int_count & " instances of " & pstr_name & " is running on " & pstr_server & vbNewLine & vbNewLine
    str_body = str_body & "The process called " & pstr_process & " was found multiple times running in memory." & vbNewLine
    f_email pstr_name & " has multiple instances running on " & pstr_server, str_body
  End If
End Function   '## f_check_process() ##

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
End Function   '## f_email() ##

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
