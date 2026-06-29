Attribute VB_Name = "Funtions"
Public Function GetNextDeptID(TableName As String) As String

    'Dim NextVal As String
    'NextVal = GetMax("cast(DeptID as numeric)", TableName)
    'GetNextDeptID = Format(Val(NextVal) + 1, "00")
    
    GetNextDeptID = ""
    
End Function

Public Sub ShowWait(Optional frm As Form)

    On Error GoTo err
'    frmWait.SetFocus
'    frmWait.Refresh
    Screen.MousePointer = vbHourglass
Exit Sub
'    frmWait.Show , frm
err:
    If err.Number = 401 Then
        If Not frm Is Nothing Then
            frm.Hide
            Resume
        End If
    Else
        Resume Next
    End If
    
End Sub

Public Function GetNextEmpID(DeptID As String, Optional bTempEmp As Boolean = False) As String

    Dim bContractorDept As Boolean
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT TempDept FROM Departments WHERE DeptID='" & DeptID & "'", con, adOpenForwardOnly, adLockReadOnly
        If IsNull(!TempDept) Then
            bContractorDept = False
        Else
            bContractorDept = !TempDept
        End If
        .Close
    End With
    Set rs = Nothing
    
    Dim NextVal As String
'    If bContractorDept Then
'        NextVal = GetMax("Cast(Right(EmpID,6) as numeric)", "Employees", "Where LEFT(EmpID,5)='PCNT-'")
'        GetNextEmpID = "PCNT-" & Format(Val(NextVal) + 1, "000000")
'    Else
'        NextVal = GetMax("Cast(Right(EmpID,6) as numeric)", "Employees", "Where LEFT(EmpID,4)='PSS-'")
'        GetNextEmpID = "PSS-" & Format(Val(NextVal) + 1, "000000")
'    End If
    
    If strCompany = "PAKSMITH GROUP" Then
        NextVal = GetMax("Cast(Right(EmpID,6) as numeric)", "Employees", "Where LEFT(EmpID,4)='PSS-'")
        GetNextEmpID = "PSS-" & Format(Val(NextVal) + 1, "000000")
    ElseIf strCompany = "Weldon Instruments" Then
        NextVal = GetMax("Cast(Right(EmpID,6) as numeric)", "Employees", "Where LEFT(EmpID,4)='WIS-'")
        GetNextEmpID = "WIS-" & Format(Val(NextVal) + 1, "000000")
    ElseIf strCompany = "Weldon Industries" Then
        '''
    ElseIf strCompany = "Sunlike" Then
        NextVal = GetMax("Cast(Right(EmpID,6) as numeric)", "Employees", "Where LEFT(EmpID,4)='SLI-'")
        GetNextEmpID = "SLI-" & Format(Val(NextVal) + 1, "000000")
    ElseIf strCompany = "QEL" Then
        NextVal = GetMax("Cast(Right(EmpID,6) as numeric)", "Employees", "Where LEFT(EmpID,4)='QEL-'")
        GetNextEmpID = "QEL-" & Format(Val(NextVal) + 1, "000000")
    ElseIf strCompany = "Tecno" Then
        NextVal = GetMax("Cast(Right(EmpID,5) as numeric)", "Employees", "")
        If bTempEmp Then
            GetNextEmpID = "TMP-" & Format(Val(NextVal) + 1, "00000")
        Else
            GetNextEmpID = "TCN-" & Format(Val(NextVal) + 1, "00000")
        End If
    ElseIf strCompany = "Dr-Frgz" Then
        If strSubCompany = "Intl" Then
            NextVal = GetMax("Cast(Right(EmpID,6) as numeric)", "Employees", "")
            If bTempEmp Then
                GetNextEmpID = "T-" & Format(Val(NextVal) + 1, "000000")
            Else
                GetNextEmpID = "DF-" & Format(Val(NextVal) + 1, "000000")
            End If
        Else
            NextVal = GetMax("Cast(Right(EmpID,6) as numeric)", "Employees", "WHERE LEFT(EmpID,4)='DFG-'")
            GetNextEmpID = "DFG-" & Format(Val(NextVal) + 1, "000000")
        End If
    ElseIf strCompany = "Banzai" Then
        NextVal = GetMax("CAST(Right(EmpID,5) as numeric)", "Employees", "WHERE LEFT(EmpID,4)='BNZ-'")
        GetNextEmpID = "BNZ-" & Format(Val(NextVal) + 1, "00000")
    ElseIf strCompany = "IAA" Then
        NextVal = GetMax("CAST(Right(EmpID,5) as numeric)", "Employees", "WHERE LEFT(EmpID,4)='EMR-'")
        GetNextEmpID = "EMR-" & Format(Val(NextVal) + 1, "00000")
    ElseIf strCompany = "Kami" Then
        NextVal = GetMax("CAST(Right(EmpID,5) as numeric)", "Employees", "WHERE LEFT(EmpID,4)='KMI-'")
        GetNextEmpID = "KMI-" & Format(Val(NextVal) + 1, "00000")
    ElseIf strCompany = "Towne" Then
        NextVal = GetMax("CAST(Right(EmpID,5) as numeric)", "Employees", "WHERE LEFT(EmpID,4)='TWN-'")
        GetNextEmpID = "TWN-" & Format(Val(NextVal) + 1, "00000")
    End If
    
End Function
Public Function GetNextContID(DeptID As String) As String
    Dim NextVal As String
    NextVal = GetMax("cast(Right(EmpID,4) as numeric)", "Contractors", "Where DeptID='" & DeptID & "'")
    GetNextContID = DeptID & "-" & Format(Val(NextVal) + 1, "0000")
End Function
Public Sub HideWait()

    Unload frmWait
    Screen.MousePointer = vbDefault
End Sub
Public Function SetDateFormat(txt As TextBox) As Boolean

    
    
    If Len(Trim(txt.Text)) = 10 Then
        If Left$(txt.Text, 2) <= 31 Then
            If Mid$(txt.Text, 4, 2) <= 12 Then
                If Mid$(txt.Text, 3, 1) = "-" And Mid$(txt.Text, 6, 1) = "-" Then
                    SetDateFormat = True
                    Exit Function
                End If
            End If
        End If
    End If


    If Len(Trim(txt.Text)) = 8 Then
        Dim st As String
        st = "/-"
        If InStr(txt.Text, st) = 0 Then
            If Left(txt, 2) > 31 Then
                MsgBox "Days must be lesser than or equal to 31", vbInformation, "Information"
                txt.SetFocus
                txt.SelStart = 0
                txt.SelLength = 2
                Exit Function
            End If
            If Mid(txt, 3, 2) > 12 Then
                MsgBox "Month must be lesser than or equal 12", vbInformation, "Information"
                txt.SetFocus
                txt.SelStart = 2
                txt.SelLength = 2
                Exit Function
            Else
                txt.Text = Format(txt.Text, "00-00-0000")
                'txt.Text = Format$(txt.Text, "dd-MM-yyyy")
                SetDateFormat = True
            End If
        Else
            MsgBox "Please enter date without backslash (/) or dash(-)", vbInformation, "Information"
            txt.SelStart = InStr(txt.Text, st, vbTextCompare)
            'txt.SelStart = 0
            txt.SelLength = 1
            'txt.SelLength = Len(txt.TEXT)
            Exit Function
        End If
    Else
        MsgBox "Please enter valid date. Date format is (ddmmyyyy)", vbInformation, "Information"
        Exit Function
    End If
    
End Function
Public Function SetSaveFormat(str As String) As String
    If Len(str) < 4 Then Exit Function
    str = Mid$(str, 4, 3) & Left$(str, 3) & Right$(str, 4)
    SetSaveFormat = str
    
End Function
    
Sub FillCountries(strCust As String, cmb As MSForms.ComboBox)

    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Distinct Country from ForeignCustomers Where CustCode='" & strCust & "'", con, adOpenForwardOnly, adLockReadOnly
        cmb.Clear
        Do Until .EOF
            cmb.AddItem .Fields(0) & ""
            .MoveNEXT
        Loop
        .Close
    End With
    If cmb.ListCount > 0 Then cmb.ListIndex = 0
    Set rs = Nothing
    
End Sub


Public Function ConvertInWords(ByVal Amount As Currency, ByVal ldenimination As String) As String

 
    Dim WordArray(90) As String
    Dim ctr As Integer
    Dim tmpStr As String
    Dim X As Integer
    Dim lDigitsArray(15) As Integer
    Dim lTens As Integer
    Dim lExitCon As Boolean
    Dim lDenomination As String


    If Not (IsNumeric(Amount)) Then
        
        ConvertInWords = ""
        Exit Function
    End If

    'CHANGE THIS AS REQUIRED
    lDenomination = "Pesos"
    'Initialize Array
    WordArray(1) = "One"
    WordArray(2) = "Two"
    WordArray(3) = "Three"
    WordArray(4) = "Four"
    WordArray(5) = "Five"
    WordArray(6) = "Six"
    WordArray(7) = "Seven"
    WordArray(8) = "Eight"
    WordArray(9) = "Nine"
    WordArray(10) = "Ten"
    WordArray(11) = "Eleven"
    WordArray(12) = "Twelve"
    WordArray(13) = "Thirteen"
    WordArray(14) = "Fourteen"
    WordArray(15) = "Fifteen"
    WordArray(16) = "Sixteen"
    WordArray(17) = "Seventeen"
    WordArray(18) = "Eighteen"
    WordArray(19) = "Nineteen"
    WordArray(20) = "Twenty"
    WordArray(30) = "Thirty"
    WordArray(40) = "Forty"
    WordArray(50) = "Fifty"
    WordArray(60) = "Sixty"
    WordArray(70) = "Seventy"
    WordArray(80) = "Eighty"
    WordArray(90) = "Ninety"
    '*****************************************************
    ' CHECK IF LESS THAN ZERO
    '*****************************************************
    Dim lFraction As String
    Dim lIsFraction As Boolean
    Dim lInFracWords As String
    Dim lFracArray(100) As Integer
    Dim lFracCounter As Integer
    Dim lTenthsFrac As Integer


    If Int(Amount) = 0 Then
        lFraction = Trim$(str(Amount))


        If Len(lFraction) = 1 Then
            ConvertInWords = ""
        Else
            lFraction = str$(CCur(Left(lFraction, 3)))
            lFraction = Right$(lFraction, Len(lFraction) - 1)
            ConvertInWords = lFraction & "/100" 'WordArray(Int(lFraction)) & " Cents "
        End If

        Exit Function
    End If

   
    '*****************************************************
    ' Pass the digits (whole numbers) into a temporary array
    '*****************************************************
    tmpStr = Trim$(str(Int(Amount)))
    ctr = Len(tmpStr) - 1
    X = Len(tmpStr)


    While X >= 1
        lDigitsArray(X) = Left$(tmpStr, 1)
        tmpStr = Right$(tmpStr, ctr)
        ctr = ctr - 1
        If ctr = 0 Then ctr = 1
        X = X - 1
    Wend

    tmpStr = Trim$(str(Int(Amount)))
    '*******************************************************
    ConvertInWords = ""
    lExitCon = False
    X = Len(tmpStr)


    Do While X >= 1
        
        lTens = 0
        
        '******************************************************
        'CHECK IF VALUE IN TENS PLACE (MIDDLE) WOULD BE
        '11,or 12..13..14..15..ETC
        '*******************************************************


        Select Case X
            
            Case 2, 5, 8, 11, 14
            lDigitsArray(X) = lDigitsArray(X) * 10


            If lDigitsArray(X) = 10 Then
                lTens = lDigitsArray(X) + lDigitsArray(X - 1)
                lDigitsArray(X - 1) = 0


                If X = 2 Then
                    lExitCon = True
                End If

            End If

        End Select

    
    


    If lTens <> 0 Then
        ConvertInWords = ConvertInWords & " " & WordArray(lTens)
        If lExitCon Then Exit Do
    Else


        If lDigitsArray(X) <> 0 And X > 1 Then


            If X Mod 3 = 0 Then
                ConvertInWords = ConvertInWords & " " & WordArray(lDigitsArray(X)) & " Hundred"
            Else
                ConvertInWords = ConvertInWords & " " & WordArray(lDigitsArray(X))
            End If

        Else


            If lDigitsArray(X) <> 0 And lTens = 0 And X = 1 Then
                ConvertInWords = ConvertInWords & " " & WordArray(lDigitsArray(X))
            Else
                'Ldigitsarray(x)=0
                ConvertInWords = ConvertInWords & ""
            End If

        End If

    End If

    


    Select Case X
        Case 5, 8, 11, 14
        'cases for 100,000,90,000 ; 80,000;


        If (lDigitsArray(X) Mod 10 = 0 And lDigitsArray(X - 1) = 0) And (lDigitsArray(X) <> 0) Then
            


            Select Case X
                Case 5
                ConvertInWords = ConvertInWords & " Thousand"
                Case 8
                ConvertInWords = ConvertInWords & " Million"
                Case 11
                ConvertInWords = ConvertInWords & " Billion"
                Case 14
                ConvertInWords = ConvertInWords & " Trillion"
            End Select

    End If

    
    Case 4


    If lDigitsArray(X) <> 0 Then
        ConvertInWords = ConvertInWords & " Thousand "
    End If

    
    Case 7


    If lDigitsArray(X) <> 0 Then
        ConvertInWords = ConvertInWords & " Million "
    End If

    Case 10


    If lDigitsArray(X) <> 0 Then
        ConvertInWords = ConvertInWords & " Billion "
    End If

    Case 13


    If lDigitsArray(X) <> 0 Then
        ConvertInWords = ConvertInWords & " Trillion "
    End If

    
    Case 6, 9, 12


    If lDigitsArray(X) <> 0 And lDigitsArray(X - 1) = 0 And lDigitsArray(X - 2) = 0 Then


        Select Case X
            Case 6
            ConvertInWords = ConvertInWords & " Thousand "
            Case 9
            ConvertInWords = ConvertInWords & " Million "
            Case 12
            ConvertInWords = ConvertInWords & " Billion "
        End Select

    
End If

End Select


X = X - 1
Loop

If ldenimination = "" Then
    ldenimination = "$"
End If
 
  Dim Location, tmpvalue, PostFix, Cdesignation
      PostFix = ""
      Select Case ldenimination
           Case "$"
                PostFix = "Cents"
                Cdesignation = "US dollors"
           Case "Rs"
                 PostFix = "Paisas"
                Cdesignation = "Rs"
           Case "DM"
                ' PostFix = "Pasas"
                Cdesignation = "DM"
           Case "£"
                ' PostFix = "Pasas"
                Cdesignation = "Pounds"
           Case "¥"
                ' PostFix = "Pasas"
                Cdesignation = "Yen"
           Case Else
                Cdesignation = ldenimination
      End Select
      
      Location = InStr(1, Amount, ".", vbTextCompare)
         If Location <> 0 Then
             tmpvalue = Mid$(Amount, Location + 1, Len(Amount))
             
             'if length =1 then add 0 to cents
             If Len(tmpvalue) = 1 Then
             tmpvalue = tmpvalue & 0
             'elseif length > = 3 then trim to only two numbers
             ElseIf Len(tmpvalue) >= 3 Then
             tmpvalue = Mid(tmpvalue, 1, 2)
             End If
             
              
             ConvertInWords = Cdesignation & Space(3) & Trim$(ConvertInWords) & "  and  " & ConvertInWords(CCur(tmpvalue), "Cent")
               
           
         Else
             ConvertInWords = Cdesignation & Space(3) & Trim$(ConvertInWords) & " Only"
         End If
         
'ConvertInWords = Trim$(ConvertInWords) & " " & ldenimination & lInFracWords & "Only"
End Function

Public Sub AddDeclarations(strOrder As String, FGrid As MSFlexGrid, Optional FGrid1)
    On Error GoTo err
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "VOrderDeclarations"
        .ActiveConnection = con
        
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        Set rs = .Execute(, strOrder)
    End With
    Dim iRow As Integer
    Dim IDFound As Boolean
    With FGrid
        Do Until rs.EOF
            IDFound = False
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) = rs.Fields(0) & "" Then
                    IDFound = True
                    Exit For
                End If
            Next i
            If Not IDFound Then
                If Not (.Rows = 2 And .TextMatrix(1, 0) = "") Then
                    .AddItem ""
                    If Not IsMissing(FGrid1) Then FGrid1.AddItem ""
                End If
                iRow = .Rows - 1
                .TextMatrix(iRow, 0) = rs.Fields(0)
                .TextMatrix(iRow, 1) = rs.Fields(1)
                If Not IsMissing(FGrid1) Then
                    FGrid1.TextMatrix(iRow, 1) = rs.Fields(1)
                    FGrid1.TextMatrix(iRow, 0) = rs.Fields(2)
                End If
            End If
            rs.MoveNEXT
        Loop
        
        'If .Rows = 2 And .TextMatrix(1, 0) = "" Then .AddItem
        
    End With
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Public Sub AddPorts(CityID As Integer, cmb As MSForms.ComboBox)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Port,PortID from Ports Where CityID=" & CityID, con, adOpenForwardOnly, adLockReadOnly
        cmb.Clear
        Do Until .EOF
            cmb.AddItem .Fields(0) & ""
            cmb.List(.AbsolutePosition - 1, 1) = .Fields(1) & ""
            .MoveNEXT
        Loop
        If cmb.ListCount > 0 Then cmb.ListIndex = 0
        .Close
    End With
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Public Sub AddCities(CountryName As String, cmb As MSForms.ComboBox)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select City,CityID from Cities Where CountryName='" & CountryName & "'", con, adOpenForwardOnly, adLockReadOnly
        cmb.Clear
        Do Until .EOF
            cmb.AddItem .Fields(0) & ""
            cmb.List(.AbsolutePosition - 1, 1) = .Fields(1) & ""
            .MoveNEXT
        Loop
        If cmb.ListCount > 0 Then cmb.ListIndex = 0
        .Close
    End With
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Public Function LastSubStrPos(SourceStr As String, StrToFind As String) As Long

    If Len(StrToFind) = 0 Then Exit Function
    Dim start As Long
    Dim pos As Long
    Dim ReturnStr As String
    Dim LastPos As Long
    
    pos = InStr(1, SourceStr, StrToFind)
    If pos = 0 Then Exit Function
        
    Do While pos > 0
        LastPos = pos
        pos = InStr(LastPos + 1, SourceStr, StrToFind)
        
    Loop
    'ReturnStr = ReturnStr & SourceStr
    LastSubStrPos = LastPos
    
End Function
'Public Function SubStrCount(SourceStr As String, StrToFind As String) As Double
'    If Len(StrToFind) = 0 Then Exit Function
'    SubStrCount = (Len(SourceStr) - Len(Replace(SourceStr, StrToFind, ""))) / Len(StrToFind)
'End Function


Public Sub StartTrans(con As ADODB.Connection)
    On Error Resume Next
    con.BeginTrans
    If con.Errors.count > 0 Then
        If con.Errors(0) = Err_Trans Then
            con.RollbackTrans
            con.BeginTrans
        End If
    End If
    
End Sub

Public Function LotProcess(ProcessID As Long) As Boolean
    On Error GoTo err
    
    Dim lSNo1 As Long, lSNo2 As Long
'    lSNo1 = Val(con.Execute("Select SNo From VInitialLotProcess").Fields(0).value & "")
    lSNo2 = Val(con.Execute("Select SNo From Processes Where ProcessID=" & ProcessID).Fields(0).Value & "")
    
'    If lSNo2 >= lSNo1 Then
'        LotProcess = True
'    Else
'        LotProcess = False
'    End If
    LotProcess = False
    Exit Function
err:
    MsgBox err.Description
    
End Function
Public Function IsInitLotProcess(ProcessID As Long) As Boolean
    On Error GoTo err
    
    Dim lSNo1 As Long, lSNo2 As Long
'    lSNo1 = Val(con.Execute("Select SNo From VInitialLotProcess").Fields(0).Value & "")
'    lSNo2 = Val(con.Execute("Select SNo From Processes Where ProcessID=" & ProcessID).Fields(0).Value & "")
'
'    If lSNo2 = lSNo1 Then
'        IsInitLotProcess = True
'    Else
'        IsInitLotProcess = False
'    End If
    IsInitLotProcess = False
    Exit Function
err:
    MsgBox err.Description
    
End Function
Public Function RcvLot(ProcessID As Long) As Boolean

    On Error GoTo err
    
    Dim lSNo1 As Long, lSNo2 As Long
'    lSNo1 = Val(con.Execute("Select SNo From VInitialLotProcess").Fields(0).value & "")
    lSNo2 = Val(con.Execute("Select SNo From Processes Where ProcessID=" & ProcessID).Fields(0).Value & "")
    
'    If lSNo2 > lSNo1 Then
'        RcvLot = True
'    Else
'        RcvLot = False
'    End If
    RcvLot = False
    Exit Function
err:
    MsgBox err.Description
End Function


Public Function TimeToMin(RTime As String) As Integer

    TimeToMin = Hour(RTime) * 60 + Minute(RTime)

End Function

Public Function GetServerDate(Optional WithTime As Boolean = False) As Date

    If WithTime Then
        GetServerDate = con.Execute("Select GetDate()").Fields(0)
    Else
        GetServerDate = CDate(Format(con.Execute("Select GetDate()").Fields(0), "dd-MMM-yyyy"))
    End If
    
End Function

Public Function DateDiffInMonthsNDays(ByVal myStartDT As Date, ByVal myEndDT As Date, ByRef p_iMonths As Integer, ByRef p_iDays As Integer) As String

    Dim iYears As Integer, iMonths As Integer, iDays As Integer
    myEndDT = DateAdd("d", 1, myEndDT)
    iYears = DateDiff("m", myStartDT, myEndDT)
    
    If Day(myStartDT) > Day(myEndDT) Then
        iYears = iYears - 1
    End If
    
    myStartDT = DateAdd("M", iYears, myStartDT)
    iMonths = iYears Mod 12

    iDays = DateDiff("d", myStartDT, myEndDT)
    If iDays < 0 Then
        myStartDT = DateAdd("M", -1, myStartDT)
        iDays = DateDiff("d", myStartDT, myEndDT)

        iMonths = iMonths - 1
    End If

    iYears = Int(iYears / 12)

    p_iMonths = iMonths + (iYears * 12)
    p_iDays = iDays
    
    DateDiffInMonthsNDays = iYears & " Years, " & iMonths & " Months, " & iDays & " Days."
    
End Function

Public Function DateDiffInYears(myDT As Date) As Integer

    Dim iYears As Integer, iMonths As Integer, iDays As Integer
    iYears = DateDiff("m", myDT, Date)
    
    If Day(myDT) > Day(Date) Then
        iYears = iYears - 1
    End If
    
    iYears = Int(iYears / 12)
    
    
    DateDiffInYears = iYears
    
End Function

Public Function DateDiffInText2(myDT As Date) As String

    Dim iYears As Integer, iMonths As Integer, iDays As Integer
    iYears = DateDiff("m", myDT, Date)
    
    If Day(myDT) > Day(Date) Then
        iYears = iYears - 1
    End If
    
    myDT = DateAdd("M", iYears, myDT)
    iMonths = iYears Mod 12

    iDays = DateDiff("d", myDT, Date)
    If iDays < 0 Then
        myDT = DateAdd("M", -1, myDT)
        iDays = DateDiff("d", myDT, Date)

        iMonths = iMonths - 1
    End If

    iYears = Int(iYears / 12)
    
    
    DateDiffInText2 = iYears & " Y, " & iMonths & " M, " & iDays & " D"
    
End Function


Public Function DateDiffInText(myDT As Date) As String

    Dim iYears As Integer, iMonths As Integer, iDays As Integer
    iYears = DateDiff("m", myDT, Date)
    
    If Day(myDT) > Day(Date) Then
        iYears = iYears - 1
    End If
    
    myDT = DateAdd("M", iYears, myDT)
    iMonths = iYears Mod 12

    iDays = DateDiff("d", myDT, Date)
    If iDays < 0 Then
        myDT = DateAdd("M", -1, myDT)
        iDays = DateDiff("d", myDT, Date)

        iMonths = iMonths - 1
    End If

    iYears = Int(iYears / 12)
    
    
    DateDiffInText = iYears & " Years, " & iMonths & " Months, " & iDays & " Days."
    
End Function

Public Sub ShowMeasuringUnits()
    Dim f As New frmMiscS_New
    Load f
    f.ShowForm "Units", "UnitName", "Units", "UnitID", "Units"
End Sub

Public Sub ShowQMSDocumentTypes()
    Dim f As New frmMiscS_New
    Load f
    f.ShowForm "QMS_DocumentTypes", "DocumentType", "Document Types", "EntryID", "Document Types"
End Sub

Public Sub ShowQMSDocumentTypes_Ex(strMasterDocType As String)
    Dim f As New frmQMSDocuments_Ex_Types
    Load f
    f.ShowForm "QMS_Master_DocumentTypes_DocTypes", "DocumentType", strMasterDocType & " Types", "EntryID", strMasterDocType & " Types", strMasterDocType
End Sub


Function IsHoliday(DT As Date) As Boolean
IsHoliday = False
    If LCase(Format(DT, "ddd")) = "sun" Then
        IsHoliday = True
        Exit Function
    Else
        Dim rs As New ADODB.Recordset
        Set rs = con.Execute("Select Count(*) From Holidays Where DT='" & Format(DT, "dd-MMM-yyyy") & "'")
        IsHoliday = CBool(rs.Fields(0).Value > 0)
        Set rs = Nothing
    End If
    
End Function
