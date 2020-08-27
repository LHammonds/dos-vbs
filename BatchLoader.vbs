Option Explicit
'#############################################################
'## Name         : BatchLoader.vbs
'## Version      : 1.4
'## Date         : 2005-02-18
'## Author       : LHammonds
'## Purpose      : Upload Transcription
'## Requirements : This program MUST be run from a server with PMSI 7.6 client and mapping to P:
'##                Windows Script Host v5.6 (CSCRIPT.EXE)
'##                For email to work, you must have an SMTP server
'## Setup        : Schedule this script to be called such as "CSCRIPT.EXE BatchLoader.vbs"
'## Output       : Transcriptions loaded into PMSI
'######################## CHANGE LOG #########################
'## DATE       VER WHO WHAT WAS CHANGED
'## ---------- --- --- ---------------------------------------
'## 2004-09-23 1.0 LTH  Created program.
'## 2005-01-12 1.1 LTH  Added the parse files function to separate loaded text from error text.
'## 2005-01-17 1.2 LTH  Added email notification using CDO.
'## 2005-01-27 1.3 LTH  Add filename checks.
'## 2005-02-18 1.4 LTH  Replaced "MoveFile" with "CopyFile" and "DeleteFile" to keep script from halting if duplicate file exists in target location.
'#############################################################

Const OverwriteExisting = True
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
DIM arr_text()
DIM bln_testmode
DIM int_i
DIM int_file_count
DIM int_max_filename_length
DIM obj_current_file
DIM obj_email
DIM obj_fso
DIM obj_folder
DIM obj_file
DIM obj_log
DIM obj_shell
DIM str_batchloader
DIM str_current_filename
DIM str_newfilename
DIM str_drive
DIM str_email_to
DIM str_email_from
DIM str_email_server
DIM str_errors_dir
DIM str_errors_no_dot_code_dir
DIM str_log
DIM str_original_dir
DIM str_processing_dir
DIM str_processed_dir
DIM str_temp_dir
DIM str_source_dir
DIM str_root
DIM str_temp
DIM str_server
DIM str_server_process

'##################################
'## Initialize Program Variables ##
'##################################
bln_testmode = False
If bln_testmode Then
  str_drive = "D:"
  str_log = "D:\Incoming\Trans\BatchLoader.log"
  str_email_to = "john.doe@mydomain.com"
Else
  str_drive = "P:"
  str_log = "C:\Apps\BatchLoader.log"
  str_email_to = "john.doe@mydomain.com,jane.doe@mydomain.com"
End If
str_server = "srv-pmsi-app"
str_server_process = "Textload.exe"
str_email_from = "john.doe@mydomain.com"
str_email_server = "srv-file"
str_batchloader = "C:\Apps\Batchload.bat"
str_root = "\Incoming\Trans"
str_source_dir = str_drive & str_root & "\New"
str_original_dir =  str_drive & str_root & "\Originals"
str_processing_dir = str_drive & str_root & "\Processing"
str_processed_dir = str_drive & str_root & "\Processed"
str_temp_dir = str_drive & str_root & "\Temp"
str_errors_dir = str_drive & str_root & "\New\Err"
str_errors_no_dot_code_dir = str_drive & str_root & "\New\ErrDotCode"
int_file_count = 0
int_max_filename_length = 35  '## Textloader can only handle 65 characters including drive, path, & filename. ##
Set obj_shell = CreateObject("WScript.Shell")
Set obj_fso = CreateObject("Scripting.FileSystemObject")
Set obj_log = obj_fso.OpenTextFile(str_log, ForAppending, True)
'## Check for existance of .txt file as well as count them. ##
str_temp = ""
Set obj_folder = obj_fso.GetFolder(str_source_dir)
For Each obj_file in obj_folder.Files
  If LCase(Right(obj_file.Name,3)) = "txt" and obj_file.Size > 0 Then
    int_file_count = int_file_count + 1
  End If
Next
'## End script if no text files available for processing. ##
If int_file_count = 0 Then
  obj_log.Close
  Set obj_log = Nothing
  wscript.Quit(0)
End If
obj_log.Writeline(f_timestamp & " - ** Script Start **")
'## Move files to a temporary directory. ##
obj_log.WriteLine(f_timestamp & " - Moving files from NEW to TEMP directory.")
obj_fso.MoveFile str_source_dir & "\*.txt", str_temp_dir & "\"
'## Archive the original unmodified files. ##
obj_log.WriteLine(f_timestamp & " - Archiving original unmodified files.")
obj_fso.CopyFile str_temp_dir & "\*.txt", str_original_dir & "\", OverwriteExisting
'## Remove any invalid characters from the file names & fix filename length. ##
obj_log.WriteLine(f_timestamp & " - Removing invalid characters from filenames.")
Set obj_folder = obj_fso.GetFolder(str_temp_dir)
For Each obj_file in obj_folder.Files
  str_newfilename = f_Generate_New_Filename(str_temp_dir,obj_file.Name)
  If str_newfilename <> obj_file.Name Then
    obj_fso.MoveFile str_temp_dir & "\" & obj_file.Name, str_temp_dir & "\" & str_newfilename
  End If
Next
'## Examine each file and filter out duplicate .D codes and page breaks. ##
'## This is specific to how this company creates the original MS Word Doc files. ##
obj_log.WriteLine(f_timestamp & " - Apply filter to text files.")
Set obj_folder = obj_fso.GetFolder(str_temp_dir)
For Each obj_file in obj_folder.Files
  If LCase(Right(obj_file.Name,3)) = "txt" and obj_file.Size > 0 Then
    '## Process the non-empty text file. ##
    str_current_filename = obj_file.Name
    Set obj_current_file = obj_fso.OpenTextFile(str_temp_dir & "\" & str_current_filename, ForReading)
    int_i = 0
    Do Until obj_current_file.AtEndOfStream
      '## Put each line into an array slot. ##
      REDIM Preserve arr_text(int_i)
      arr_text(int_i) = obj_current_file.ReadLine
      int_i = int_i + 1
    Loop
    obj_current_file.Close
    Set obj_current_file = obj_fso.CreateTextFile(str_processing_dir & "\" & str_current_filename)
    For int_i = 0 to UBound(arr_text) Step 1
      If LCase(Left(Trim(arr_text(int_i)),3)) = ".d:" Then
        '## Check for special case to filter on. ##
        If int_i + 1 <= UBound(arr_text) Then
          '## We are safe to check the next line. ##
          If LCase(Left(Trim(arr_text(int_i+1)),4)) = "page" Then
            '## We have a special case that we need to filter on. ##
            '## 1) Don't write out the current line ##
            '## 2) Skip writing out the next line ##
            int_i = int_i + 1
          Else
            '## This is not a special case.  Simply write out the line. ##
            obj_current_file.WriteLine(arr_text(int_i))
          End If
        Else
          '## This is the last line.  Simply write out the line. ##
          obj_current_file.WriteLine(arr_text(int_i))
        End If
      Else
        '## This is not a special case.  Simply write out the line. ##
        obj_current_file.WriteLine(arr_text(int_i))
      End If
    Next
    obj_current_file.Close
  End If
Next
'## Delete the temp files since they are already archived and also in the PROCESSING directory. ##
obj_log.WriteLine(f_timestamp & " - Cleanup the temp directory.")
obj_fso.DeleteFile str_temp_dir & "\*.txt", True
'## Pause the script for 2 seconds to allow the OS to process the requests. ##
wscript.sleep 2000
obj_log.WriteLine(f_timestamp & " - Run Text Loader to load " & int_file_count & " files.")
str_temp = "%comspec% /C CALL " & str_batchloader & " " & str_processing_dir & "\*.txt"
If Not bln_testmode Then
  '## Do not execute the textloader if in test mode. ##
  obj_shell.run str_temp
End If
'## Pause the script for 10 seconds to allow the OS to start the program. ##
wscript.sleep 10000
'## Pause the script until the textloader program finishes. ##
Do Until Not f_process_running (str_server, str_server_process)
  '## Kill time and wait for textloader to finish. ##
  wscript.sleep 5000
Loop
'## Check and parse files with errors and archive those that don't. ##
Set obj_folder = obj_fso.GetFolder(str_processing_dir)
For Each obj_file in obj_folder.Files
  If LCase(Right(obj_file.Name,3)) = "txt" Then
    '## Process the text file. ##
    str_current_filename = Left(obj_file.Name, Len(obj_file.Name) - 4)
    If obj_fso.FileExists(str_processing_dir & "\" & str_current_filename & ".err") Then
      If bln_testmode Then
        obj_log.WriteLine(f_timestamp & " - DEBUG MODE: Parsing files: " & str_current_filename & ".")
      End If
      f_parse_files str_current_filename
    Else
      If bln_testmode Then
        obj_log.WriteLine(f_timestamp & " - DEBUG MODE: No .err File Found: " & str_current_filename & ".")
      End If
      If obj_file.Size = 0 Then
        '## Delete the empty file. ##
        obj_fso.DeleteFile str_processing_dir & "\" & obj_file.Name
      Else
        '## Move the file to the "PROCESSED" directory. ##
        obj_fso.CopyFile str_processing_dir & "\" & str_current_filename & ".txt", str_processed_dir & "\" & str_current_filename & ".txt", True
        If obj_fso.FileExists (str_processed_dir & "\" & str_current_filename & ".txt") Then
          obj_fso.DeleteFile str_processing_dir & "\" & str_current_filename & ".txt"
        End If
      End If
    End If
  End If
Next

'########################
'## Email Notification ##
'########################

'## Check for error files and send email notification if any are found. ##
int_file_count = 0
Set obj_folder = obj_fso.GetFolder(str_errors_dir)
For Each obj_file in obj_folder.Files
  If LCase(Right(obj_file.Name,3)) = "txt" Then
    int_file_count = int_file_count + 1
  End If
Next
Set obj_folder = obj_fso.GetFolder(str_errors_no_dot_code_dir)
For Each obj_file in obj_folder.Files
  If LCase(Right(obj_file.Name,3)) = "txt" Then
    int_file_count = int_file_count + 1
  End If
Next
If int_file_count > 0 Then
  '## Send email notification about the existance of errors. ##
  '## Create the a file to hold the message of the email. ##
  str_temp = ""
  str_temp = str_temp & "The automated BatchLoader process detected " & int_file_count & " error(s) with the upload of the transcription files." & vbNewLine & vbNewLine
  str_temp = str_temp & "Please look at the following directories to fix them:" & vbNewLine
  str_temp = str_temp & "    " & str_errors_dir & vbNewLine
  str_temp = str_temp & "    " & str_errors_no_dot_code_dir & vbNewLine & vbNewLine
  str_temp = str_temp & "Thanks,"
  str_temp = str_temp & "BatchLoader.vbs"
  '## Send the email using CDO. ##
  Set obj_email = CreateObject("CDO.Message")
  obj_email.From = str_email_from
  obj_email.To = str_email_to
  obj_email.Subject = "Automated BatchLoader Detected Import Errors"
  obj_email.Textbody = str_temp
  obj_email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
  obj_email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = str_email_server
  obj_email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
  obj_email.Configuration.Fields.Update
  obj_email.Send
  Set obj_email = Nothing
End If

obj_log.Writeline(f_timestamp & " - ** Script End **")
obj_log.Close
Set obj_log = Nothing
Set obj_fso = Nothing

'######################
'## Function Section ##
'######################

Function f_Generate_New_Filename(pstr_path, pstr_filename)
  DIM lobj_fso
  DIM lobj_file
  DIM lstr_ext
  DIM lstr_new_filename
  Set lobj_fso = CreateObject("Scripting.FileSystemObject")
  lstr_ext = lobj_fso.GetExtensionName(pstr_path & pstr_filename)
  lstr_new_filename = Replace(pstr_filename,"." & lstr_ext,"") '## Remove extension. ##
  lstr_new_filename = Left(f_valid_chars_only(lstr_new_filename),int_max_filename_length) '## No special characters allowed. ##
  Set lobj_fso = Nothing
  f_Generate_New_Filename = lstr_new_filename & "." & lstr_ext
End Function   '## f_Generate_New_Filename() ##

Function f_valid_chars_only(pstr_value)
  DIM lstr_allowed
  DIM lstr_valid
  DIM lint_index
  lstr_allowed = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwyz0123456789_.-"
  lstr_valid = ""
  For lint_index = 1 To Len(pstr_value)
    If Instr(lstr_allowed, Mid(pstr_value, lint_index, 1)) Then
      lstr_valid = lstr_valid & Mid(pstr_value, lint_index, 1)
    End If
  Next
  f_valid_chars_only = lstr_valid
End Function  '## f_valid_chars_only() ##

Function f_process_running (pstr_server, pstr_process)
  DIM obj_WMI_Service
  DIM col_ProcessList
  DIM obj_process
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
  If int_count > 0 Then
    '## Found instance running in memory. ##
    f_process_running = True
  Else
    '## Not found in memory. ##
    f_process_running = False
  End If
End Function  '## f_process_running() ##

Function f_parse_files (pstr_filename)
  DIM lobj_fso
  DIM lobj_err
  DIM lobj_txt
  DIM lobj_new_good
  DIM lobj_new_bad
  DIM lstr_txt_line
  DIM lstr_err_line
  DIM lstr_current_write_file
  DIM lbln_test
  DIM lstr_temp
  Set lobj_fso = CreateObject("Scripting.FileSystemObject")
  Set lobj_err = lobj_fso.OpenTextFile(str_processing_dir & "\" & pstr_filename & ".err", ForReading, True)
  lstr_temp = lobj_err.ReadAll
  lobj_err.Close
  If InStr(1, lstr_temp, "No .D: line found.",1) Then
    '## Special Case: At least one identification line cannot be found, do not process file. ##
    If bln_testmode Then
      obj_log.Writeline(f_timestamp & " - DEBUG MODE: " & pstr_filename & ".err, Moving files since no .D: ID was found.")
    End If
    lobj_fso.MoveFile str_processing_dir & "\" & str_current_filename & ".txt", str_errors_no_dot_code_dir & "\"
    lobj_fso.MoveFile str_processing_dir & "\" & str_current_filename & ".err", str_errors_no_dot_code_dir & "\"
  Else
    '## All errors have IDs that can be found in the txt file. ##
    Set lobj_err = lobj_fso.OpenTextFile(str_processing_dir & "\" & pstr_filename & ".err", ForReading, True)
    Set lobj_txt = lobj_fso.OpenTextFile(str_processing_dir & "\" & pstr_filename & ".txt", ForReading, True)
    Set lobj_new_good = lobj_fso.OpenTextFile(str_processed_dir & "\" & pstr_filename & "-a.txt", ForAppending, True)
    Set lobj_new_bad = lobj_fso.OpenTextFile(str_errors_dir & "\" & pstr_filename & "-b.txt", ForAppending, True)
    lstr_err_line = ".D INITIALIZING VARIABLE"
    lbln_test = True
    If lobj_err.AtEndOfStream = True Then
      lbln_test = False
    End If
    Do While lbln_test = True
      '## Find first ID. ##
      lstr_err_line = lobj_err.ReadLine
      If LCase(Left(lstr_err_line,3)) = ".d:" Then
        lbln_test = False
      End If
      If lobj_err.AtEndOfStream = True Then
        lbln_test = False
      End If
    Loop
    lstr_current_write_file = "NOERROR"
    Do While lobj_txt.AtEndOfStream <> True
      lstr_txt_line = lobj_txt.ReadLine
      If LCase(Left(lstr_txt_line,3)) = ".d:" Then
        '## Special Case: Compare error file to txt file. ##
        If lstr_txt_line = lstr_err_line Then
          '## Move file pointer to the error file. ##
          lstr_current_write_file = "ERROR"
          lstr_err_line = ".D REINITIALIZING VARIABLE"
          lbln_test = True
          If lobj_err.AtEndOfStream = True Then
            lbln_test = False
          End If
          Do While lbln_test = True
            '## Find next error ID (if there is one) ##
            lstr_err_line = lobj_err.ReadLine
            If LCase(Left(lstr_err_line,3)) = ".d:" Then
              lbln_test = False
            End If
            If lobj_err.AtEndOfStream = True Then
              lbln_test = False
            End If
          Loop
        Else
          lstr_current_write_file = "NOERROR"
        End If
      End If
      If lstr_current_write_file = "NOERROR" Then
        '## Write Current .txt line to the no error file. ##
        lobj_new_good.WriteLine lstr_txt_line
      Else
        '## Write Current .txt line to the error file. ##
        lobj_new_bad.WriteLine lstr_txt_line
      End If
    Loop
    lobj_err.Close
    lobj_txt.Close
    Set lobj_txt = Nothing
    lobj_new_good.Close
    Set lobj_new_good = Nothing
    lobj_new_bad.Close
    Set lobj_new_bad = Nothing
    lobj_fso.MoveFile str_processing_dir & "\" & pstr_filename & ".err", str_errors_dir & "\" & pstr_filename & "-b.err"
    lobj_fso.DeleteFile str_processing_dir & "\" & pstr_filename & ".txt", True
  End If
  Set lobj_err = Nothing
  Set lobj_fso = Nothing
End Function '## f_parse_files() ##

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
