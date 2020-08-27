Option Explicit
'#############################################################
'## Name         : Email_Using_CDO_without_SMTP.vbs
'## Version      : 1.0
'## Date         : 2004-12-21
'## Author       : LHammonds
'## Purpose      : Example on how to send an email using CDO
'##                It uses CDO to send email from a computer
'##                where the SMTP Service has not been installed.
'##                Designed to work on Microsoft's corporate network.
'## Requirements : SMTP Mail Server.
'## Output       : Email
'######################## CHANGE LOG #########################
'## DATE       VER WHO WHAT WAS CHANGED
'## ---------- --- --- ---------------------------------------
'## 2004-12-21 1.0 LTH Created script.
'#############################################################

DIM obj_email
DIM str_email_server
DIM str_email_body
DIM str_email_to
DIM str_email_from

str_email_to = "joe@email.com"
str_email_from = "webmaster@email.com"
str_email_server = "mail.server.com"

str_email_body = ""
str_email_body = str_email_body & "Test Intro Header" & vbNewLine & vbNewLine
str_email_body = str_email_body & "Testing 1-2-3" & vbNewLine
str_email_body = str_email_body & "Testing A-B-C" & vbNewLine & vbNewLine
str_email_body = str_email_body & "Sincerely," & vbNewLine
str_email_body = str_email_body & str_email_from & vbNewLine

Set obj_email = CreateObject("CDO.Message")

obj_email.From = str_email_from
obj_email.To = str_email_to
obj_email.Subject = "VBScript CDO Email Test" 
obj_email.Textbody = str_email_body
obj_email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
obj_email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = str_email_server
obj_email.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
obj_email.Configuration.Fields.Update
obj_email.Send

Set obj_email = Nothing
