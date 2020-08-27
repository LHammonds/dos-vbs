Option Explicit
'#############################################################
'## Name         : export_zimbra_ad_users.vbs
'## Version      : 1.0
'## Date         : 2011-10-29
'## Author       : LHammonds
'## Purpose      : Export AD users to a comma-delimited file that
'##                are authorized to have a Zimbra mailbox.
'## Requirements : Windows Script Host v5.6 (CSCRIPT.EXE), Microsoft Access Drivers
'######################## CHANGE LOG #########################
'## DATE       VER WHO WHAT WAS CHANGED
'## ---------- --- --- ---------------------------------------
'## 2004-04-06 1.0 LTH  Created program.
'#############################################################

'## Field #1 = LoginID
'## Field #2 = First Name
'## Field #3 = Middle Initial
'## Field #4 = Last Name
'## Field #5 = Full Name
'## Field #6 = Title
'## Field #7 = Description
'## Field #8 = Comments
'## Field #9 = Telephone
'## Field #10 = Home Phone
'## Field #11 = Mobile Phone
'## Field #12 = Fax Number
'## Field #13 = Pager
'## Field #14 = Company
'## Field #15 = Office
'## Field #16 = Street Address
'## Field #17 = PO Box
'## Field #18 = City
'## Field #19 = State
'## Field #20 = Postal Code
'## Field #21 = Country
'## Field #22 = Password Replacement Value
'## Field #23 = Unused (mainly to avoid the end-of-line character being read into the last value)

'## NOTE: This could use a data cleanup routine that replaces all commas in a  ##
'## variable with something else such as a period instead to avoid CSV issues. ##

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

Const ConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=G:\soft\Database\inventory.accdb;Persist Security Info=False;"
Const ZimbraSecurityGroup = "GRP_ZimbraEmail"
Const DefaultPassword = "clinic1"
Const LogFile = "C:\Apps\export_ad_users.log"
Const RemoteFile = "\\192.168.107.25\share\adlist.csv"
Const TempFile = "G:\soft\Database\adlist.csv"

Dim objRootDSE, strDNC, objDomain
Dim obj_fso, obj_TempFile, obj_log, int_count
Dim cnn_db, rst_users, str_ConnString, str_sql

int_count = 0
Set obj_fso = CreateObject("Scripting.FileSystemObject")

'## Open log file ##
Set obj_log = obj_fso.OpenTextFile(LogFile, ForAppending, True)
obj_log.Writeline(f_timestamp & " - ** Script Start **")

'## Create and Open database connection. ##
Set cnn_db = CreateObject("ADODB.Connection")
cnn_db.Errors.Clear
cnn_db.Open ConnString
Set rst_users = CreateObject("ADODB.Recordset")
rst_users.CursorType = adOpenStatic
rst_users.LockType = adLockOptimistic

'## Get Windows Domain. ##
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNC = objRootDSE.Get("DefaultNamingContext")
Set objDomain = GetObject("LDAP://" & strDNC)

'## Create / Overwrite the export file. ##
Set obj_TempFile = obj_fso.CreateTextFile(TempFile)
obj_TempFile.Close
Set obj_TempFile = Nothing

Call TrollTheFolders(objDomain)

'## Close database connection. ##
Set rst_users = Nothing
cnn_db.Close
Set cnn_db = Nothing

'## Copy the file to remote location and if it worked, delete the source file ##
obj_fso.CopyFile TempFile, RemoteFile, OverwriteExisting
If obj_fso.FileExists(RemoteFile) Then
  obj_fso.DeleteFile(TempFile)
End If

'## Close the log file. ##
obj_log.Writeline(f_timestamp & " --- Processed " & int_count & " records.")
obj_log.Writeline(f_timestamp & " - ** Script Completed **")
obj_log.Close
Set obj_log = Nothing
Set obj_fso = Nothing

Sub TrollTheFolders(pobjDomain)
  '## Function: Traverse the AD structure to find users wherever they may reside. ##
  '##           The trick is that this function is called recursively in order to ##
  '##           inspect every sub-folder that may contain user accounts.          ##
  Dim lobjFile, lobjMember, lstrLine, lblnInZimbraGroup
  Dim lstrSamAccountName, lstrFirstName, lstrInitials, lstrLastName
  Dim lstrFullName, lstrTitle, lstrDescription, lstrComment
  Dim lstrTelephoneNumber, lstrHomePhone, lstrMobile, lstrFaxNumber, lstrPager
  Dim lstrCompany, lstrOffice, lstrStreetAddress, lstrPostOfficeBox
  Dim lstrCity, lstrState, lstrPostalCode, lstrCountry, lstrPassword
  Dim lcolGroups, lobjGroup, lstrTemp

  For Each lobjMember In pobjDomain
    '## Examine each object but process only "user" objects. ##
    If lobjMember.Class = "user" Then
      Set lobjFile = obj_fso.OpenTextFile (TempFile, ForAppending, True)
      If Not (isempty(lobjMember.samAccountName)) Then lstrSamAccountName = lobjMember.samAccountName Else lstrSamAccountName = "" End If
      If Not (isempty(lobjMember.GivenName)) Then lstrFirstName = lobjMember.GivenName Else lstrFirstName = "" End If
      If Not (isempty(lobjMember.initials)) Then lstrInitials = lobjMember.initials Else lstrInitials = "" End If
      If Not (isempty(lobjMember.sn)) Then lstrLastname = lobjMember.sn Else lstrLastName = "" End If
      If Not (isempty(lobjMember.CN)) Then lstrFullName = lobjMember.CN Else lstrFullName = "" End If
      If Not (isempty(lobjMember.title)) Then lstrTitle = lobjMember.title Else lstrTitle = "" End If
      If Not (isempty(lobjMember.Description)) Then lstrDescription = lobjMember.Description Else lstrDescription = "" End If
      If Not (isempty(lobjMember.comment)) Then lstrComment = lobjMember.comment Else lstrComment = "" End If
      If Not (isempty(lobjMember.telephoneNumber)) Then lstrTelephoneNumber = lobjMember.telephoneNumber Else lstrTelephoneNumber = "" End If
      If Not (isempty(lobjMember.homePhone)) Then lstrHomePhone = lobjMember.homePhone Else lstrHomePhone = "" End If
      If Not (isempty(lobjMember.mobile)) Then lstrMobile = lobjMember.mobile Else lstrMobile = "" End If
      If Not (isempty(lobjMember.otherFacsimileTelephoneNumber)) Then lstrFaxNumber = lobjMember.otherFacsimileTelephoneNumber Else lstrFaxNumber = "" End If
      If Not (isempty(lobjMember.pager)) Then lstrPager = lobjMember.pager Else lstrPager = "" End If
      If Not (isempty(lobjMember.company)) Then lstrCompany = lobjMember.company Else lstrCompany = "" End If
      If Not (isempty(lobjMember.physicalDeliveryOfficeName)) Then lstrOffice = lobjMember.physicalDeliveryOfficeName Else lstrOffice = "" End If
      If Not (isempty(lobjMember.streetAddress)) Then lstrStreetAddress = lobjMember.streetAddress Else lstrStreetAddress = "" End If
      If Not (isempty(lobjMember.postOfficeBox)) Then lstrPostOfficeBox = lobjMember.postOfficeBox Else lstrPostOfficeBox = "" End If
      If Not (isempty(lobjMember.l)) Then lstrCity = lobjMember.l Else lstrCity = "" End If
      If Not (isempty(lobjMember.st)) Then lstrState = lobjMember.st Else lstrState = "" End If
      If Not (isempty(lobjMember.postalCode)) Then lstrPostalCode = lobjMember.postalCode Else lstrPostalCode = "" End If
      If Not (isempty(lobjMember.countryCode)) Then lstrCountry = lobjMember.countryCode Else lstrCountry = "" End If

      '## Lookup the user in the database to obtain the password. ##
      str_sql = "SELECT pass_nt_current FROM tbl_employee WHERE (username_nt = """ & lstrSamAccountName & """)"
      rst_users.Open str_sql, cnn_db
      If rst_users.BOF and rst_users.EOF Then
        '## No matching record found in the password database. ##
        lstrPassword = DefaultPassword
      Else
        '## Password found ##
        lstrPassword = Trim(rst_users("pass_nt_current"))
      End If
      rst_users.Close
      lblnInZimbraGroup = 0
      For Each lobjGroup in lobjMember.Groups
        '## See if this member belongs to the group that allows Zimbra mailboxes ##
        If LCase(lobjGroup.cn) = LCase(ZimbraSecurityGroup) Then
          lblnInZimbraGroup = 1
        End If
      Next
      If lblnInZimbraGroup = 1 Then
        '## Member is associated to the Zimbra Email group and thus allowed to have a Zimbra mailbox. ##
        int_count = int_count + 1
        lstrLine = lstrSamAccountName & "," & lstrFirstName & "," & lstrInitials & "," & lstrLastName & "," & lstrFullName & "," &_
            lstrTitle & "," & lstrDescription & "," & lstrComment & "," & lstrTelephoneNumber & "," & lstrHomePhone & "," & lstrMobile & "," &_
            lstrFaxNumber & "," & lstrPager & "," & lstrCompany & "," & lstrOffice & "," & lstrStreetAddress & "," & lstrPostOfficeBox & "," &_
            lstrCity & "," & lstrState & "," & lstrPostalCode & "," & lstrCountry & "," & lstrPassword & ",unused"
        lobjFile.WriteLine(lstrLine)
      End If
      lobjFile.Close
      Set lobjFile = Nothing
    End If
    If lobjMember.Class = "organizationalUnit" or lobjMember.Class = "container" Then
      '## Recurse further down to find the users. ##
      TrollTheFolders(lobjMember)
    End If
  Next
End Sub  '## TrollTheFolders() ##

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
