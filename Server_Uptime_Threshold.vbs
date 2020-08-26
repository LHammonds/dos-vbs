Option Explicit
'#############################################################
'## Name          : Server_Uptime_Threshold.vbs
'## Version       : 1.0
'## Date          : 2005-01-19
'## Author        : LHammonds
'## Purpose       : Check & record server uptimes and report exceptions.
'## Compatibility : Windows 2000/2003/2008/2012
'## Required      : MS Access Database
'######################## CHANGE LOG #########################
'## DATE       VER WHO WHAT WAS CHANGED
'## ---------- --- --- ---------------------------------------
'## 2005-01-19 1.0 LTH Created program.
'#############################################################

'---- File Handling Values ----
Const OverwriteExisting = True
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

'---- CursorTypeEnum Values ----
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3

'---- LockTypeEnum Values ----
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4

'## Define program variables. ##
DIM cnn_db
DIM rst_server
DIM rst_uptime
DIM str_sql
DIM str_ConnString
DIM str_log
DIM obj_fso
DIM obj_log
DIM int_days
DIM int_total
DIM str_email_alarm
DIM str_email_warning
DIM str_email_reboot
DIM str_email_settings

'## Initialize program variables. ##
str_ConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Apps\Server_Uptime.mdb;"
'str_ConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Apps\Server_Uptime.mdb;User Id=admin; Password=;"
'str_ConnString = "Driver={Microsoft Access Driver (*.mdb)};DBQ=D:\Apps\Server_Uptime.mdb;Uid=;Pwd=;"
str_log = "D:\Apps\Server_Uptime_Threshold.log"
int_total = 0
str_email_alarm = ""
str_email_warning = ""
str_email_reboot = ""
str_email_settings = ""

'## Open Log file. ##
Set obj_fso = CreateObject("Scripting.FileSystemObject")
Set obj_log = obj_fso.OpenTextFile(str_log, ForAppending, True)
obj_log.Writeline(f_timestamp & " - ** Script Start **")

'## Create and Open database connection. ##
Set cnn_db = CreateObject("ADODB.Connection")
cnn_db.Errors.Clear
cnn_db.Open str_ConnString

Set rst_server = CreateObject("ADODB.Recordset")
rst_server.CursorType = adOpenStatic
rst_server.LockType = adLockOptimistic

Set rst_uptime = CreateObject("ADODB.Recordset")
rst_uptime.CursorType = adOpenStatic
rst_uptime.LockType = adLockOptimistic

'## Delete any rows with today's date (due to being run multiple times). ##
cnn_db.Execute "DELETE FROM tbl_uptime WHERE DateCreated = #" & FormatDateTime(Now(), vbShortDate) & "#"

str_sql = "SELECT * FROM tbl_server WHERE Active = True ORDER BY Server_Name ASC"
rst_server.Open str_sql, cnn_db
If rst_server.BOF and rst_server.EOF Then
  '## No records to process. ##
  obj_log.Writeline(f_timestamp & " - Error: No server records found.")
Else
  Do While Not rst_server.EOF
    '## Process each row one at a time. ##
    int_total = int_total + 1
    int_days = f_uptime(rst_server("Server_Name"))
    str_email_settings = str_email_settings & UCase(rst_server("Server_Name")) & ", Max=" & rst_server("Max_Uptime") & ", Warning=" & rst_server("Threshold_Warning") & ", Current=" & int_days & vbNewLine

    '## Record uptime information. ##
    str_sql = "SELECT * FROM tbl_uptime WHERE 1=2"
    '## Open an empty recordset. ##
    rst_uptime.Open str_sql, cnn_db
    rst_uptime.AddNew()
    rst_uptime("Server_ID") = rst_server("Server_ID")
    rst_uptime("Days_Online") = int_days
    rst_uptime("DateCreated") = FormatDateTime(Now(), vbShortDate)
    rst_uptime.Update()
    rst_uptime.Close

    If int_days >= rst_server("Max_Uptime") Then
      If int_days > 1000 Then
        '## Problem with uptime calculation. ##
        '## Do nothing. ##
      Else
        '## Email Alarm ##
        str_email_alarm = str_email_alarm & vbNewLine & UCase(rst_server("Server_Name")) & " up for " & int_days & " days, max is " & rst_server("Max_Uptime")
      End If
    ElseIf int_days >= (rst_server("Max_Uptime") - rst_server("Threshold_Warning")) Then
      '## Email Warning ##
      str_email_warning = str_email_warning & vbNewLine & UCase(rst_server("Server_Name")) & " up for " & int_days & " days, max is " & rst_server("Max_Uptime")
    End If
    If int_days <= 1 Then
      '## Email Reboot Notice ##
      str_email_reboot = str_email_reboot & vbNewLine & UCase(rst_server("Server_Name")) & " up for " & int_days & " days. This means a reboot occured within the last 24 hours."
    End If
    rst_server.MoveNext
  Loop
End If

If Len(str_email_alarm) > 0 Then
  f_email "UPTIME ALARM", str_email_alarm & vbNewLine & vbNewLine & "SERVER SETTINGS:" & vbNewLine & str_email_settings
  obj_log.Writeline(str_email_alarm)
End If
If Len(str_email_warning) > 0 Then
  f_email "UPTIME WARNING", str_email_warning & vbNewLine & vbNewLine & "SERVER SETTINGS:" & vbNewLine & str_email_settings
  obj_log.Writeline(str_email_warning)
End If
If Len(str_email_reboot) > 0 Then
  f_email "UPTIME REBOOT", str_email_reboot
  obj_log.Writeline(str_email_reboot)
End If
If (Len(str_email_alarm) + Len(str_email_warning) + Len(str_email_reboot)) > 0 Then
  obj_log.Writeline()
End If

rst_server.Close
Set rst_server = Nothing
cnn_db.Close
Set cnn_db = Nothing

'## Finish Log file and release object. ##
obj_log.Writeline(f_timestamp & " - Checked uptime threshold for " & int_total & " servers.")
obj_log.Writeline(f_timestamp & " - ** Script End **")
obj_log.Close
Set obj_log = Nothing

'**********************
'** FUNCTION SECTION **
'**********************

Function f_email(pstr_subject, pstr_body)
  DIM obj_email
  Set obj_email = CreateObject("CDO.Message")
  obj_email.From = "uptime@mydomain.com"
  obj_email.To = "dirk.diggler@mydomain.com,john.doe@mydomain.com,jane.doe@mydomain.com"
  obj_email.Subject = pstr_subject
  obj_email.Textbody = pstr_body & vbNewLine & vbNewLine & "Source: srv-file\D\Apps\Server_Update_Threshold.vbs"
  obj_email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
  obj_email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "192.168.0.25"
  obj_email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
  obj_email.Configuration.Fields.Update
  obj_email.Send
  Set obj_email = Nothing
End Function  '## f_email() ##


Function f_uptime(pstr_computer)
  DIM int_time
  DIM obj_WMI_Service
  DIM col_OSs
  DIM obj_OS
  DIM dtm_Bootup
  DIM dtm_Last_Bootup_Time
  DIM dtm_System_Uptime

  int_time = 0
  On Error Resume Next
  Set obj_WMI_Service = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & pstr_computer & "\root\cimv2")
  Set col_OSs = obj_WMI_Service.ExecQuery("Select * from Win32_OperatingSystem")
  For Each obj_OS in col_OSs
    dtm_Bootup = obj_OS.LastBootUpTime
    dtm_Last_Bootup_Time = f_WMI_String_To_Date(dtm_Bootup)
    dtm_System_Uptime = DateDiff("d", dtm_Last_Bootup_Time, Now)
    int_time = int_time & dtm_System_Uptime
  Next
  On Error GoTo 0
  f_uptime = CDbl(int_time)
End Function  '## f_uptime() ##


Function f_WMI_String_To_Date(pdtm_Bootup)
  f_WMI_String_To_Date = CDate(Mid(pdtm_Bootup, 5, 2) & "/" & Mid(pdtm_Bootup, 7, 2) & "/" & Left(pdtm_Bootup, 4) & " " & Mid (pdtm_Bootup, 9, 2) & ":" & Mid(pdtm_Bootup, 11, 2) & ":" & Mid(pdtm_Bootup, 13, 2))
End Function


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
