Option Explicit
'#############################################################
'## Name         : Email_Using_CDONTS.vbs
'## Version      : 1.0
'## Date         : 2001-09-11
'## Author       : LHammonds
'## Purpose      : Example on how to send an email using CDONTS
'## Output       : Email
'######################## CHANGE LOG #########################
'## DATE       VER WHO WHAT WAS CHANGED
'## ---------- --- --- ---------------------------------------
'## 2004-12-21 1.0 LTH Created script for use on WinNT servers.
'#############################################################

DIM obj_mail
DIM str_bodytext
DIM str_emailto
DIM str_emailbcc

str_emailto = "joe@email.com"
str_emailbcc = "jane@email.com"

str_bodytext = ""
str_bodytext = str_bodytext & "<p>Testing <strong>MYSERVER</strong> email via VB Script.</p>"

Set obj_mail		= CreateObject("CDONTS.NewMail")
obj_mail.From		= "webmaster@email.com" 
obj_mail.To		= str_emailto
obj_mail.BCC		= str_emailbcc
obj_mail.Subject	= "Testing Email"
obj_mail.BodyFormat	= 0
obj_mail.MailFormat	= 0
obj_mail.Body		= str_bodytext
obj_mail.Send
Set obj_mail		= Nothing
