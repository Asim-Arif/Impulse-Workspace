Attribute VB_Name = "UtilityFunctions"
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal _
    nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Public ConnectStr As String
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Private Declare Function OSRegEnumKey Lib "advapi32" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal iSubKey As Long, ByVal lpszName As String, ByVal cchName As Long) As Long
'Private Declare Function OSRegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
'Private Declare Function OSRegOpenKey Lib "advapi32" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public DBName As String, RptPath As String, strCompany As String, COMPANYADDRESS As String, strSubCompany As String
Public rptPayrollPath As String

Public Declare Function FindWindowEx Lib "User32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SetClassLong Lib "User32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function InvalidateRect Lib "User32" (ByVal hwnd As Long, lpRect As Long, ByVal bErase As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long

Public Enum enuTBType
    enuTB_FLAT = 1
    enuTB_STANDARD = 2
End Enum

Public Const GCL_HBRBACKGROUND = (-10)

Public Enum ListOptions
    
    ITEM_TEXT = 1
    VAL_OF_ITEM_TEXT = 2
    ITEM_KEY = 3
    VAL_OF_ITEM_KEY = 4
    
End Enum


'User Info Options
Public Enum UserInfoOptions
    Opt_NewUser = 1
    Opt_ChangeUserPassword = 3
    Opt_ResetPassword = 4
End Enum




'For Links To MSWord etc

 'Global ObjVar1 As Object
 'Global ObjVar2 As Object
 ''Global Objvar3 As Object
  
  
  
  
  
'Global Variable Holding Current Operator Name

'Global UserName As String


 ' Public Declare Function ReleaseCapture& Lib "user32" ()
 ' Public Declare Function GetCapture& Lib "user32" ()
 ' Public Declare Function SetCapture& Lib "user32" (ByVal hwnd&)


'Public con As New ADODB.Connection
'HKEYLOCALMACHINE CONSTANT
'Const HKLM = &H80000002

'Constant For String Value

'Private Const REG_SZ = 1

    
Sub AddToCombo(comb As Control, Field As String, Table As String, Optional Condition As String, Optional Append As Boolean, Optional OtherFields As String)

    'On Error Resume Next
    If OtherFields <> "" Then
        Field = Field & "," & OtherFields
    End If
    'Clear the combo box if Append Option is not Given
              
    If Not Append Then comb.Clear

    Dim recset As New ADODB.Recordset
 
    
    recset.Open "Select " & Field & " From " & Table & " " & Condition, con, adOpenForwardOnly, adLockReadOnly
 
    'Fill the combo box
 
    For i = 0 To recset.RecordCount - 1
        If Not IsNull(recset.Fields(0)) Then
            comb.AddItem recset.Fields(0)
            For j = 1 To recset.Fields.count - 1
                comb.List(comb.ListCount - 1, j) = recset.Fields(j) & ""
            Next
        End If
        recset.MoveNEXT
    Next
    
    recset.Close
    Set recset = Nothing
    
End Sub

Function FindInCombo(comb As Object, Value As String, Optional ExactMatch As Boolean) As Boolean
FindInCombo = False
'find the value in the given combo box

For i = 0 To comb.ListCount - 1
   If ExactMatch Then
      If UCase(comb.List(i)) = UCase(Value) Then comb.ListIndex = i: Exit For
   Else
      If InStr(UCase(comb.List(i)), UCase(Value)) > 0 Then comb.ListIndex = i: Exit For
   End If
  
Next
If i <> comb.ListCount Then FindInCombo = True
End Function
Function OnlyNumber(AsciiVal As Integer, Optional Dot As Boolean)

Dim RetVal As Integer

If (AsciiVal >= 48 And AsciiVal <= 57) Or (AsciiVal = 8) Then
 
    RetVal = AsciiVal
 
Else

    If Dot And AsciiVal = 46 Then
        RetVal = AsciiVal
    Else
        RetVal = 0
    End If
End If

    OnlyNumber = RetVal
    
End Function






''''''''''''''''''''''''''''''''''''''''
'Sub main()
'frmSplash.Show
'End Sub
''''''''''''''''''''''''''''''''''''''''




 Function GetID(StrText As String) As String
 Dim StartPos As Integer, EndPos As Integer
 On Error GoTo err
 StartPos = InStr(StrText, "[") + 1
 
 If StartPos = 1 Then
  GetID = ""
  Exit Function
  End If
 
  EndPos = InStr(StrText, "]")
  GetID = Mid(StrText, StartPos, EndPos - StartPos)
  Exit Function
err:
MsgBox "Following Error Occured in GetID" & vbNewLine & err.Description
End Function


Function GetITNumber() As Double
con.Execute "Delete * from ITNum"
con.Execute "insert into itnum(IT) values ('a')"
Dim rsIT As ADODB.Recordset
GetITNumber = Abs(con.Execute("Select ItNo from Itnum").Fields(0))
 
Set rsIT = Nothing
End Function


Public Function GetMax(Field As String, Table As String, Optional Condition As String) As String

Dim recmax As New ADODB.Recordset
recmax.Open "Select max(" & Field & ") from " & Trim(Table) & " " & Condition, con, adOpenForwardOnly, adLockReadOnly
If Not IsNull(recmax.Fields(0)) Then
 GetMax = recmax.Fields(0)
Else
 GetMax = 0
End If
recmax.Close
Set recmax = Nothing
End Function






Public Function GetNextJobID(OnDate As Date) As String
Dim rs As New ADODB.Recordset
      rs.Open "SELECT max(val(right(TRIM(JobID),3))) FROM Jobs WHERE LEFT(JobID,4) = '" & _
       Format(OnDate, "MMdd") & "'", con, adOpenForwardOnly, adLockReadOnly
      GetNextJobID = Format(OnDate, "MMdd") & "-" & Format((Val(rs.Fields(0) & "") + 1), "000")
rs.Close
Set rs = Nothing
End Function

Public Function GetSelList(LV As ListView, ListOption As ListOptions, Optional SingleQuote As Boolean)
Dim SQ As String, OptStr As String
Dim TheList As String
If SingleQuote Then SQ = "'" Else SQ = ""

    For i = 1 To LV.ListItems.count
    
        Select Case ListOption
         Case ListOptions.ITEM_TEXT
            OptStr = LV.ListItems(i).Text
         Case ListOptions.ITEM_KEY
            OptStr = LV.ListItems(i).key
         Case ListOptions.VAL_OF_ITEM_TEXT
            OptStr = Val(LV.ListItems(i).Text)
         Case ListOptions.VAL_OF_ITEM_KEY
            OptStr = Val(LV.ListItems(i).key)
        End Select
        
        If LV.ListItems(i).Checked Then TheList = TheList & SQ & OptStr & SQ & ","
    Next

If Len(TheList) > Len(SQ) + 1 Then
    GetSelList = Left(TheList, Len(TheList) - 1)
Else
    GetSelList = ""
End If
    
End Function


Public Function SubStrCount(SourceStr As String, StrToFind As String) As Double
    If Len(StrToFind) = 0 Then Exit Function
    SubStrCount = (Len(SourceStr) - Len(Replace(SourceStr, StrToFind, ""))) / Len(StrToFind)
End Function


'Public Sub adjustItemLedger(ItemID As String, Sno As Double, QtyToAdd As Double, Optional DonotUpdateCurrentStock As Boolean)
'
'
' If Not DonotUpdateCurrentStock Then
'   'Update Current Stock
'   con.Execute "Update Items Set inhand=inhand+" & QtyToAdd & " Where ItemID='" & ItemID & "'"
' End If
'
'
' Dim SQL As String
' 'Update Ledger Stock Balances
' SQL = "Update InoutDetail set inhand=inhand+(" & QtyToAdd & ") Where ItemID='" & ItemID & "' AND (Sno>" & Sno & ")"
'
' con.Execute SQL
'
'End Sub

Public Sub ShowList(Lst As ListBox, cmb As Object, txt As String, Optional ContLeft As Long, Optional ContTop As Double)
Lst.Clear
   'Lst.SetFocus
   'SendKeys "{ESC}"
   'Cmb.SetFocus
   Lst.Visible = False
   For i = 0 To cmb.ListCount - 1
      If InStr(UCase(cmb.List(i)), UCase(txt)) > 0 Then Lst.AddItem cmb.List(i)
   Next
   
   Lst.Move cmb.Left + ContLeft, cmb.Top + cmb.Height + 25 + ContTop ', Cmb.Width
   Lst.Visible = True
   
   'Lst.DropDown
   
End Sub

Public Function getPassword() As Boolean

    With frmPassword
        .Show 1
        If .Canceled Then
            getPassword = False
        ElseIf .Pass.Text = GetSetting(App.Title, "Settings", "Password", "admin") Then
            getPassword = True
        Else
            MsgBox "Invalid Password", vbInformation
            getPassword = False
        End If
        Unload frmPassword '

    End With
    
End Function

Public Function getPasswordNew(strPWD As String) As Boolean

    With frmPassword
        .Show 1
        If .Canceled Then
            getPasswordNew = False
        ElseIf .Pass.Text = strPWD Then
           getPasswordNew = True
        Else
          MsgBox "Invalid Password", vbInformation
           getPasswordNew = False
        End If
        Unload frmPassword '
    
    End With
    
End Function

Public Sub PlaySaveSound()
   Beep 1000, 150
   Beep 500, 150
End Sub


Public Function getDBPassword(strSetting As String) As Boolean
    
    Dim strPassword As String
    strPassword = GetSingleStringValue("Password", "Security", " WHERE SettingName='" & strSetting & "'")
    strPassword = CryptRC4(FromHexDump(strPassword), "AwmWrangler")
    
    With frmPassword
        .Show 1
        If .Canceled Then
            getDBPassword = False
        ElseIf .Pass.Text = strPassword Then
           getDBPassword = True
        Else
          MsgBox "Invalid Password", vbInformation
           getDBPassword = False
        End If
        Unload frmPassword '
    End With
    
End Function

Public Function getGeneralDataValue(strDataSetting As String) As String
          
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT DataValue FROM GeneralData WHERE DataName='" & strDataSetting & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            getGeneralDataValue = ""
        Else
            getGeneralDataValue = !DataValue & ""
        End If
    End With
    Set rs = Nothing
    
    Exit Function
err:
    MsgBox err.Description
    
End Function

Public Function ValidateAdminPassword() As Boolean
    Dim strAdminPWD As String
    strAdminPWD = GetSingleStringValue("Password", "Users", " WHERE UserName='Administrator'")
    ValidateAdminPassword = getPasswordNew(strAdminPWD)
End Function

Function Floor(ByVal n As Double) As Double
    '\\\\\
    'The floor function of a real number x, denoted floor(x), is a function that returns the largest integer less than or equal to x. Formally, for all real numbers x,
    Dim a As Long

    a = CLng(n)
    If a > n Then a = a - 1

    Floor = a

End Function

Function Ceil(ByVal n As Double) As Double
    '\/\/\/
    'The closely-related ceiling function, denoted ceiling(x), is the function that returns the smallest integer not less than x, or, formally,
    Dim a As Long

    a = CLng(n)
    If a < n Then a = a + 1

    Ceil = a
    
End Function

Function GetTemporaryDirectory() As String
    Dim buffer As String * 512, length As Long
    length = GetTempPath(Len(buffer), buffer)
    GetTemporaryDirectory = Left$(buffer, length)
End Function

Public Sub ShowTechDrawing(strTableName As String, strFieldName As String, strCondition As String, strFileName As String, CD1 As CommonDialog)

    On Error GoTo err
     
    Dim strSavePath As String, strProgramPath As String
    strSavePath = GetTemporaryDirectory '"\\Database\Software Attachments"
    
    Dim rs As New ADODB.Recordset
    Dim binObj As New ADODB.Stream
    rs.Open "SELECT " & strFieldName & " FROM " & strTableName & strCondition, con, adOpenForwardOnly, adLockReadOnly
    If IsNull(rs.Fields(0)) Then
        rs.Close
        Set rs = Nothing
        MsgBox "Technical Drawing not found.", vbInformation
        Exit Sub
    End If
    binObj.Type = adTypeBinary
    binObj.Open
    binObj.Write rs.Fields(strFieldName)
    
    strSavePath = strSavePath & strFileName & ".pdf"
    binObj.SaveToFile strSavePath, adSaveCreateOverWrite
    binObj.Close
    Set binObj = Nothing
         
    rs.Close
    Set rs = Nothing
     
'    strProgramPath = GetSetting(App.Title, "SettingsEx", "Acrobat Path", "C:\Program Files\Adobe\Acrobat 7.0\Reader\AcroRd32.exe")
'    Dim oFso As New FileSystemObject
'    'Dim CD1 As CommonDialog
'
'    If oFso.FileExists(strProgramPath) = False Then
'        On Error Resume Next
'        CD1.DialogTitle = "Please Select Adobe Acrobat Executeable"
'        CD1.Filter = "Acrobat Executeable|*.exe"
'        CD1.ShowOpen
'        If err.Number = 32755 Then Exit Sub
'        If CD1.FileName <> "" Then
'            strProgramPath = CD1.FileName
'            SaveSetting App.Title, "SettingsEx", "Acrobat Path", strProgramPath
'        End If
'    End If
'    'strProgramPath = "C:\Program Files\Adobe\Acrobat 7.0\Reader\AcroRd32.exe"
'    Shell """" & strProgramPath & """" & " " & """" & strSavePath & """", vbNormalFocus
    ShellExecute 0, "Open", strSavePath, vbNullString, "C:\", SW_SHOWNORMAL
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Public Sub OpenPDFFile(strFileName As String, CD1 As CommonDialog)

    On Error GoTo err
     
    strProgramPath = GetSetting(App.Title, "SettingsEx", "Acrobat Path", "C:\Program Files\Adobe\Acrobat 7.0\Reader\AcroRd32.exe")
    Dim oFso As New FileSystemObject
    'Dim CD1 As CommonDialog
    
    If oFso.FileExists(strProgramPath) = False Then
        On Error Resume Next
        CD1.DialogTitle = "Please Select Adobe Acrobat Executeable"
        CD1.Filter = "Acrobat Executeable|*.exe"
        CD1.ShowOpen
        If err.Number = 32755 Then Exit Sub
        If CD1.FileName <> "" Then
            strProgramPath = CD1.FileName
            SaveSetting App.Title, "SettingsEx", "Acrobat Path", strProgramPath
        End If
    End If
    'strProgramPath = "C:\Program Files\Adobe\Acrobat 7.0\Reader\AcroRd32.exe"
    Shell """" & strProgramPath & """" & " " & """" & strFileName & """", vbNormalFocus
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Public Sub PlayIssuanceSound()
    On Error Resume Next
    Call PlaySound(App.Path & "\Alert.wav", 0, &H0)
End Sub

Public Sub PlayReceivingSound()
    On Error Resume Next
    Call PlaySound(App.Path & "\ding.wav", 0, &H0)
End Sub

Public Sub ChangeTBBack(TB As Object, PNewBack As Long, pType As enuTBType)
    Dim lTBWnd      As Long
    Select Case pType
        Case enuTB_FLAT     'FLAT Button Style Toolbar
            'Apply directly to TB Hwnd
            DeleteObject SetClassLong(TB.hwnd, GCL_HBRBACKGROUND, PNewBack)
        Case enuTB_STANDARD 'STANDARD Button Style Toolbar
            lTBWnd = FindWindowEx(TB.hwnd, 0, "msvb_lib_toolbar", vbNullString) 'Find Hwnd first
            DeleteObject SetClassLong(lTBWnd, GCL_HBRBACKGROUND, PNewBack)      'Set new Back
    End Select
End Sub

Public Sub SetCalendarColors(ByVal myControl As DTPicker)
    With myControl
        .CalendarBackColor = vbWhite
        .CalendarForeColor = vbBlack
        .CalendarTitleBackColor = &H98&
        .CalendarTitleForeColor = vbWhite
        .CalendarTrailingForeColor = vbGray
    End With
End Sub

Public Sub FixAllDatePickers(frm As Form)

    Dim c As Control

    For Each c In frm.Controls
        If TypeOf c Is DTPicker Then
            SetCalendarColors c
        End If
    Next

End Sub
