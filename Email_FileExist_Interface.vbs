Option Explicit
'#############################################################
'## Name         : Email_FileExist_Interface.vbs
'## Version      : 1.0
'## Date         : 2005-01-13
'## Author       : LHammonds
'## Purpose      : Send email notification when exceptions need to be handled.
'## Requirements : Access to srv-interface and Windows Script Host v5.6 (CSCRIPT.EXE)
'##                For email to work, use BLAT v2.21 or higher (www.blat.net)
'## Setup        : Schedule this script to be called such as "CSCRIPT.EXE Email_FileExist_Interface.vbs"
'## Output       : Email
'######################## CHANGE LOG #########################
'## DATE       VER WHO WHAT WAS CHANGED
'## ---------- --- --- ---------------------------------------
'## 2005-01-14 1.0 LTH Created script.
'#############################################################

Const OverwriteExisting = True
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
DIM str_source_dir
DIM str_log
DIM obj_fso
DIM obj_folder
DIM obj_log
DIM obj_file
DIM int_file_count
DIM str_filenames
DIM obj_email
DIM str_body

str_source_dir = "\\srv-interface\c$\Interface\Rcv\Lab\Exceptions"
str_log = "C:\Apps\Email_FileExist_Interface.log"
int_file_count = 0
str_filenames = ""
Set obj_fso = CreateObject("Scripting.FileSystemObject")
Set obj_log = obj_fso.OpenTextFile(str_log, ForAppending, True)
obj_log.Writeline(f_timestamp & " - ** Script Start **")
Set obj_folder = obj_fso.GetFolder(str_source_dir)
For Each obj_file in obj_folder.Files
  If LCase(Right(obj_file.Name,3)) = "xml" and obj_file.Size > 0 Then
    int_file_count = int_file_count + 1
    str_filenames = str_filenames & "                      " & obj_file.Name & vbNewLine
  End If
Next
If int_file_count = 0 Then
  '## Nothing to do. Exit script. ##
  obj_log.WriteLine(f_timestamp & " - No files to process.")
Else
  '## Exceptions exists, send email notification. ##
  obj_log.WriteLine(str_filenames)
  obj_log.WriteLine(f_timestamp & " - " & int_file_count & " file(s) found.")
  '## Send Email ##
  str_body = ""
  str_body = str_body & "The App1 to App2 Interface has the following exception files that need to be handled:" & vbNewLine & vbNewLine
  str_body = str_body & str_filenames & vbNewLine
  str_body = str_body & "Please log into srv-interface and perform the following generalized steps:" & vbNewLine
  str_body = str_body & "  1. Stop the Interface program." & vbNewLine
  str_body = str_body & "  2. Start App1 CrossCheck:" & vbNewLine
  str_body = str_body & "     A) Open one of the .XML files." & vbNewLine
  str_body = str_body & "     B) Correct the errors and update the Cross Reference if necessary." & vbNewLine
  str_body = str_body & "     C) When done with the .XML file, move it into the sub-folder called PROCESSED." & vbNewLine
  str_body = str_body & "     D) Repeat A thru C until all .XML files are done." & vbNewLine
  str_body = str_body & "     E) Exit CrossCheck." & vbNewLine
  str_body = str_body & "  3. Start the Interface program." & vbNewLine
  Set obj_email = CreateObject("CDO.Message")
  obj_email.From = "webmaster@mydomain.com"
  obj_email.To = "john.doe@mydomain.com,jane.doe@mydomain.com"
  obj_email.Subject = "Exceptions Found: App1 to App2 Interface" 
  obj_email.Textbody = str_body
  obj_email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
  obj_email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")="srv-mail"
  obj_email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
  obj_email.Configuration.Fields.Update
  obj_email.Send
  Set obj_email = Nothing
End If
obj_log.Writeline(f_timestamp & " - ** Script End **")
obj_log.Close
Set obj_log = Nothing

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
