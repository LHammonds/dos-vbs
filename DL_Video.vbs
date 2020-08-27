Option Explicit
'#############################################################
'## Name         : DL_Video.vbs
'## Version      : 1.0
'## Date         : 2004-04-06
'## Author       : LHammonds
'## Purpose      : Move .mpg videos from camera
'## Requirements : Windows Script Host v5.6 (CSCRIPT.EXE)
'######################## CHANGE LOG #########################
'## DATE       VER WHO WHAT WAS CHANGED
'## ---------- --- --- ---------------------------------------
'## 2004-04-06 1.0 LTH  Created program.
'#############################################################

DIM str_path_source_root
DIM str_path_target

str_path_source_root = "G:\"
str_path_target = "E:\Video\New\"

f_move_files str_path_source_root & "MSSONY\MOML0001", str_path_target

Function f_move_files(pstr_path_source, pstr_path_target)
  DIM str_ext
  DIM obj_fso
  DIM obj_folder
  DIM obj_file
  DIM dtm_date_created
  DIM str_filename
  DIM str_new_filename
  Set obj_fso = CreateObject("Scripting.FileSystemObject")
  Set obj_folder = obj_fso.GetFolder(pstr_path_source)
  For Each obj_file in obj_folder.Files
    str_filename = obj_file.Name
    str_ext = obj_fso.GetExtensionName(pstr_path_source & "\" & str_filename)
    If LCase(str_ext) = "mpg" Then
      dtm_date_created = obj_file.DateCreated
      dtm_date_created = f_FormatDate(dtm_date_created)
      str_new_filename = Replace(str_filename,"." & str_ext,"") '## Remove extension. ##
      str_new_filename = Left(f_AlphaNumericOnly(str_new_filename),28) '## No special characters allowed. ##
      str_new_filename = dtm_date_created & " " & str_new_filename & "." & LCase(str_ext)
'      msgbox pstr_path_target & str_new_filename
      obj_file.Move pstr_path_target & str_new_filename
    End If
  Next
  Set obj_file = Nothing
  Set obj_folder = Nothing
  Set obj_fso = Nothing
End Function  '## f_move_files() ##

Function f_Generate_New_Filename(pstr_path, pstr_filename)
  DIM lobj_fso
  DIM lobj_file
  DIM lstr_ext
  DIM lstr_new_filename
  DIM ldtm_date
  Set lobj_fso = CreateObject("Scripting.FileSystemObject")
  If lobj_fso.FileExists(pstr_path & pstr_filename) Then
    '## Generate new filename ##
    lstr_ext = lobj_fso.GetExtensionName(pstr_path & pstr_filename)
    lstr_new_filename = Replace(pstr_filename,"." & lstr_ext,"") '## Remove extension. ##
    lstr_new_filename = Left(f_AlphaNumericOnly(lstr_new_filename),28) '## No special characters allowed. ##
    lstr_new_filename = ldtm_date & " " & lstr_new_filename
  Else
    '## Use current filename ##
    lstr_new_filename = pstr_filename
  End If
  Set lobj_fso = Nothing
  f_Generate_New_Filename = lstr_new_filename
End Function   '## f_Generate_New_Filename() ##

Function f_AlphaNumericOnly(pstr_value)
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
  f_AlphaNumericOnly = lstr_valid
End Function  '## f_AlphaNumericOnly() ##

Function f_FormatDate(pdtm_date)
  DIM lstr_year
  DIM lstr_month
  DIM lstr_day
  lstr_year = DatePart("yyyy", pdtm_date)
  lstr_month = DatePart("m", pdtm_date)
  lstr_day = DatePart("d", pdtm_date)
  If Len(lstr_month) = 1 Then
    lstr_month = "0" & lstr_month
  End If
  If Len(lstr_day) = 1 Then
    lstr_day = "0" & lstr_day
  End If
  f_FormatDate = lstr_year & "-" & lstr_month & "-" & lstr_day
End Function  '## f_FormatDate() ##
