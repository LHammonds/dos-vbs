Option Explicit
'#############################################################
'## Name         : Uptime2Screen.vbs
'## Version      : 1.0
'## Date         : 2004-08-13
'## Author       : LHammonds
'## Purpose      : Calculate uptime for servers and display onscreen.
'## Requirements : Windows Script Host v5.6 (CSCRIPT.EXE)
'## Output       : Uptime displayed to standard output (screen)
'######################## CHANGE LOG #########################
'## DATE       VER WHO WHAT WAS CHANGED
'## ---------- --- --- ---------------------------------------
'## 2004-08-13 1.0 LTH  Created program.
'#############################################################

DIM str_message
str_message = "This script calculates uptime for the following servers:" & vbNewline & vbNewLine
str_message = str_message & f_calc_days("srv-app")
str_message = str_message & f_calc_days("srv-dc")
str_message = str_message & f_calc_days("srv-db")
str_message = str_message & f_calc_days("srv-file")
str_message = str_message & f_calc_days("srv-print")
wscript.Echo str_message

Function f_calc_days(pstr_computer)
  DIM dtm_ConvertedDate
  DIM obj_WMIService
  DIM col_OperatingSystems
  DIM obj_OS
  DIM dtm_LastBootUpTime
  DIM int_Uptime
  DIM str_temp
  Set dtm_ConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
  Set obj_WMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & pstr_computer & "\root\cimv2")
  Set col_OperatingSystems = obj_WMIService.ExecQuery("Select * from Win32_OperatingSystem")
  For Each obj_OS in col_OperatingSystems
    dtm_ConvertedDate.Value = obj_OS.LastBootUpTime
    dtm_LastBootUpTime = dtm_ConvertedDate.GetVarDate
    int_Uptime = DateDiff("d", dtm_LastBootUpTime, Now)
    str_temp = str_temp & pstr_computer & " has been up for " & int_Uptime & " day(s)." & vbNewLine
  Next
  Set obj_OS = Nothing
  Set col_OperatingSystems = Nothing
  Set obj_WMIService = Nothing
  Set dtm_ConvertedDate = Nothing
  f_calc_days = str_temp
End Function  '## f_calc_days() ##
