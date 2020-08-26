Option Explicit
'#############################################################
'## Name     : PurgeFiles.vbs
'## Version  : 1.0
'## Date     : 2005-09-12
'## Author   : LHammonds
'## Purpose  : Purge files old than a specified date on multiple servers.
'## NOTE     : This script should be scheduled to run before a full backup.
'## Required : Nothing
'## Output   : Deleted files.
'######################## CHANGE LOG #########################
'## DATE       VER WHO WHAT WAS CHANGED
'## ---------- --- --- ---------------------------------------
'## 2005-09-12 1.0 LTH Created script.
'#############################################################

'## Note: Set days to the maximum amount of days the files can live before death.

'## Syntax: f_purge DAYS_TO_RETAIN, UNC_PATH, FILENAME_EXTENSION

f_purge 7, "srv-dc", "\c$\Windows\", "log"
f_purge 7, "srv-file", "\c$\Windows\", "old"
f_purge 7, "srv-print", "\c$\Windows\", "log"
f_purge 7, "srv-citrix", "\c$\Windows\", "old"
f_purge 7, "srv-app1", "\c$\Windows\", "log"
f_purge 7, "srv-db1", "\c$\Windows\", "old"
f_purge 7, "srv-mssql", "\d$\Program Files\Microsoft SQL Server\MSSQL\BACKUP\MYDB1\", "bak"
f_purge 7, "srv-mssql", "\d$\Program Files\Microsoft SQL Server\MSSQL\BACKUP\MYDB2\", "bak"
f_purge 7, "srv-emr", "\e$\prod\archive\", "001"
f_purge 21, "srv-emr", "\p$\reports\Log\", "txt"
f_purge 10, "srv-emr", "\c$\Documents and Settings\WatchGuard\logs", "xml"

Function f_purge(pint_days, pstr_server, pstr_path, pstr_ext)
  DIM str_log
  DIM dbl_size
  DIM obj_fso
  DIM obj_files
  DIM obj_file
  DIM str_extension

  str_log = ""
  dbl_size = 0

  Set obj_fso = CreateObject("Scripting.FileSystemObject")
  Set obj_files = obj_fso.GetFolder("\\" & pstr_server & pstr_path).Files
  for each obj_file in obj_files
    str_extension = Mid(obj_file.Name,InStrRev(obj_file.Name,".")+1)
    if LCase(str_extension) = LCase(pstr_ext) then
      if DateDiff("D", obj_file.DateCreated, date()) >= pint_days OR DateDiff("D", obj_file.DateLastModified, date()) >= pint_days then
        If Len(str_log) > 0 Then
          str_log = str_log & obj_file.DateCreated & " - " & obj_file.Name & " (" & Round(obj_file.Size / 1000,0) & "k)" & vbNewLine
        Else
          str_log = str_log & "Network Path: \\" & pstr_server & pstr_path & vbNewLine & vbNewLine
          str_log = str_log & obj_file.DateCreated & " - " & obj_file.Name & " (" & Round(obj_file.Size / 1000,0) & "k)" & vbNewLine
        End If
        dbl_size = dbl_size + obj_file.size
on error resume next
        obj_fso.DeleteFile("\\" & pstr_server & pstr_path & obj_file.Name)
on error goto 0
      end if
    end if
  next
  Set obj_files = Nothing
  Set obj_fso = Nothing

  f_email str_log, pint_days, dbl_size, pstr_server
End Function  '## f_purge ##

Function f_email(pstr_log, pint_days, pdbl_size, pstr_server)
  DIM obj_shell
  DIM obj_mail
  DIM str_body
  str_body = ""
  if len(pstr_log) > 0 then
    On Error Resume Next

    str_body = str_body & "The files below were purged from the server because they were older than " &_
              pint_days & " days." & vbNewLine & vbNewLine & pstr_log
    str_body = str_body & vbNewLine & "Drive space reclaimed: " & Round(pdbl_size / 1000,0) & "k" & vbNewLine

    set obj_shell = CreateObject("wscript.shell")
    str_body = str_body & vbNewLine & "Script Source: PurgeFiles.vbs on " & obj_shell.ExpandEnvironmentStrings("%computername%") & vbNewLine
    str_body = str_body & vbNewLine & "NOTE: This script should run prior to a full backup." & vbNewLine
    Set obj_shell = Nothing

    Set obj_mail = CreateObject("aspSmartMail.SmartMail")
    obj_mail.Server = "srv-mail"
'    obj_mail.ServerPort = 25
    obj_mail.ServerTimeOut = 20
    obj_mail.SenderName = "no-reply"
    obj_mail.SenderAddress = "no-reply@mydomain.com"
    obj_mail.Subject = "BATCH Purge - " & pstr_server
    obj_mail.Recipients.Add "ddiggler@mydomain.com", "Dirk Diggler"
    obj_mail.Recipients.Add "john.doe@mydomain.com", "John Doe"
    obj_mail.Recipients.Add "jane.doe@mydomain.com", "Jane Doe"
    obj_mail.Organization = "ABC"
    obj_mail.XMailer = "ABC server"
    obj_mail.Priority = 5   '## Priority ranges 1-5 ##
'    obj_mail.ContentType = "text/html"
    obj_mail.ContentType = "text/plain"
    obj_mail.Body = cstr(str_body)
    obj_mail.SendMail
    Set obj_mail = Nothing 
    On Error GoTo 0
  end if
End Function   '## f_email() ##
