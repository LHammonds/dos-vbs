Option Explicit
'#############################################################
'## Name         : Email_Using_BLAT.vbs
'## Version      : 1.0
'## Date         : 2005-01-13
'## Author       : LHammonds
'## Purpose      : Send Emails
'## NOTE         : This script should be scheduled to run before a full backup.
'## Requirements : Windows Script Host v5.6 (CSCRIPT.EXE)
'##                For email to work, use BLAT v2.21 or higher (www.blat.net)
'## Output       : Email
'######################## CHANGE LOG #########################
'## DATE       VER WHO WHAT WAS CHANGED
'## ---------- --- --- ---------------------------------------
'## 2005-01-13 1.0 LTH Created script.
'#############################################################

Const OverwriteExisting = True
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

DIM str_email_to
DIM str_email_subject
DIM obj_shell
DIM obj_fso
DIM obj_email_file
DIM str_email_file

str_email_to = "jane@email.com,doe@email.com"
str_email_subject = "SUBJECT TEXT LOCATED HERE"
str_email_file = "C:\BLATEMAIL.txt"

Set obj_fso = CreateObject("Scripting.FileSystemObject")
Set obj_email_file = obj_fso.OpenTextFile(str_email_file, ForWriting, True)
obj_email_file.Writeline("Hello World 1")
obj_email_file.Writeline("Hello World 2")
obj_email_file.Writeline("Hello World 3")
obj_email_file.Close
Set obj_email_file = Nothing

Set obj_shell = CreateObject("WScript.Shell")
wscript.sleep 2000
obj_shell.run "BLAT.exe " & str_email_file & " -to """ & str_email_to & """ -subject """ & str_email_subject & """"
wscript.sleep 2000
obj_fso.DeleteFile str_email_file, True

Set obj_fso = Nothing
Set obj_shell = Nothing
