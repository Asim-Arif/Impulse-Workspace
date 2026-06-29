VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmTakeAtt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adjust Attendance"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   6675
   StartUpPosition =   1  'CenterOwner
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   315
      Left            =   1155
      TabIndex        =   0
      Top             =   30
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   556
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   315
      Left            =   1155
      TabIndex        =   1
      Top             =   375
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   556
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin MSComCtl2.DTPicker DTFrom 
      Height          =   315
      Left            =   1155
      TabIndex        =   2
      Top             =   765
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dddd dd-MMM-yyyy"
      Format          =   104726531
      CurrentDate     =   37384
   End
   Begin MSComCtl2.DTPicker DTTo 
      Height          =   315
      Left            =   1155
      TabIndex        =   3
      Top             =   1155
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dddd dd-MMM-yyyy"
      Format          =   104726531
      CurrentDate     =   37384
   End
   Begin VB.Label Label2 
      Caption         =   "* It will just Complete Missing Attendance, Keeping Original Incoming times."
      ForeColor       =   &H000000FF&
      Height          =   405
      Index           =   1
      Left            =   3330
      TabIndex        =   13
      Top             =   4800
      Visible         =   0   'False
      Width           =   3165
   End
   Begin MSForms.CommandButton cmdCompleteAttendance 
      Height          =   390
      Left            =   3330
      TabIndex        =   12
      Top             =   4395
      Visible         =   0   'False
      Width           =   3165
      Caption         =   "Complete Attendance"
      PicturePosition =   327683
      Size            =   "5583;688"
      Accelerator     =   80
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label2 
      Caption         =   "* It will delete all prior leaves and Attendances."
      ForeColor       =   &H000000FF&
      Height          =   405
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   4845
      Width           =   3165
   End
   Begin MSForms.CommandButton cmdAddAttendance 
      Height          =   390
      Left            =   120
      TabIndex        =   10
      Top             =   4395
      Width           =   3165
      Caption         =   "Make Present All Days"
      PicturePosition =   327683
      Size            =   "5583;688"
      Accelerator     =   80
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdAdjustAtt 
      Height          =   390
      Left            =   120
      TabIndex        =   4
      Top             =   3045
      Visible         =   0   'False
      Width           =   3945
      Caption         =   "Adjust Lunch Time Attendance Only"
      PicturePosition =   327683
      Size            =   "6959;688"
      Accelerator     =   65
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date To :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   420
      TabIndex        =   9
      Top             =   1185
      Width           =   675
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   120
      X2              =   6500
      Y1              =   5385
      Y2              =   5385
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   360
      Left            =   3330
      TabIndex        =   5
      Top             =   5520
      Width           =   3165
      Caption         =   "Close"
      PicturePosition =   327683
      Size            =   "5583;635"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date From :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   795
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Employee :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   300
      TabIndex        =   7
      Top             =   420
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Department :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   4
      Left            =   135
      TabIndex        =   6
      Top             =   90
      Width           =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   6500
      Y1              =   5400
      Y2              =   5400
   End
End
Attribute VB_Name = "frmTakeAtt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strInTime As String
Dim strOutTime As String

Private Sub cmbdept_matched()

    cmbEmp.ClearVals
    cmbEmp.AddItem "<All Active Employees>", "0"
    If cmbDept.ID = "0" Then
        cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name", "Employees", "EmpID", "Where Active=1"
    Else
        cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "' and Active=1"
    End If
   
End Sub

Private Sub cmdAddAttendance_Click()

    On Error GoTo err
    Dim strCond As String
    Dim strEmpID() As String
    Dim i As Integer, j As Integer
    Dim rs As New ADODB.Recordset
    
    strCond = " WHERE Active=1"
    If cmbEmp.ID <> "0" Then
        strCond = strCond & " AND Employees.EmpID='" & cmbEmp.ID & "'"
    Else
        If cmbDept.ID = "0" Then
            '''
        Else
            strCond = strCond & " AND DeptID='" & cmbDept.ID & "'"
        End If
    End If
    
    With rs
        .Open "SELECT Employees.EmpID,LunchInTime,LunchOutTime,FridayBreakStartTime,FridayBreakEndTime,FridayCheckOut,ExitTime,EnterTime FROM Employees INNER JOIN VEmpSettings ON Employees.EmpID=VEmpSettings.EmpID" & strCond, con, adOpenForwardOnly, adLockReadOnly
        ReDim strEmpID(.RecordCount, 7)
        Do Until .EOF
            strEmpID(.AbsolutePosition - 1, 0) = !EmpId
            strEmpID(.AbsolutePosition - 1, 1) = !LunchInTime & ""
            strEmpID(.AbsolutePosition - 1, 2) = !LunchOutTime & ""
            strEmpID(.AbsolutePosition - 1, 3) = !FridayBreakStartTime & ""
            strEmpID(.AbsolutePosition - 1, 4) = !FridayBreakEndTime & ""
            strEmpID(.AbsolutePosition - 1, 5) = !EnterTime & ""
            strEmpID(.AbsolutePosition - 1, 6) = !ExitTime & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
'        ReDim strEmpID(1)
'        strEmpID(0) = cmbEmp.ID
'    End If
    
    Dim myDT As Date, strDay As String
     
    Dim strRandomTime As String
    For i = 0 To UBound(strEmpID) - 1
        myDT = DtFrom.Value
        Do While myDT <= DtTo.Value
            strDay = Format(myDT, "dddd")
            If strDay = "Sunday" Then
                'Skip
            ElseIf IsHoliday(myDT) Then
                'Skip
            Else
                lCount = GetSingleLongValue("COUNT(*)", "AttendanceSheet", " WHERE EmpID='" & strEmpID(i, 0) & "' AND DT='" & myDT & "'")
                If lCount > 0 Then GoTo SKIPIT
                
                Call StartTrans(con)
                
                con.Execute "DELETE FROM AttendanceSheet WHERE EmpID='" & strEmpID(i, 0) & "' AND DT='" & myDT & "'"
                con.Execute "DELETE FROM EmpTimes WHERE EmpID='" & strEmpID(i, 0) & "' AND DT='" & myDT & "'"
                con.Execute "DELETE FROM Leaves WHERE EmpID='" & strEmpID(i, 0) & "' AND DT='" & myDT & "'"
                
                iRet = Int(Rnd * (Val(9) - (Val(0) - 1)) + (Val(0)))
                strRandomTime = strEmpID(i, 5)
                strRandomTime = DateAdd("n", -iRet, strRandomTime)
                strRandomTime = Format(strRandomTime, "HH:nn")
                
                Call TakeAttendance(strEmpID(i, 0), myDT, CDate(strRandomTime))
                
                If strEmpID(i, 1) = strEmpID(i, 2) Then     'No Break
                    'Do Nothing....
                Else
                    Call TakeAttendance(strEmpID(i, 0), myDT, IIf(strDay = "Friday", CDate(strEmpID(i, 3)), CDate(strEmpID(i, 1))))
                    Call TakeAttendance(strEmpID(i, 0), myDT, IIf(strDay = "Friday", CDate(strEmpID(i, 4)), CDate(strEmpID(i, 2))))
                End If
                
                
'                Call TakeAttendance(strEmpID(i, 0), myDT, CDate(strEmpID(i, 1)))
'                Call TakeAttendance(strEmpID(i, 0), myDT, CDate(strEmpID(i, 2)))
                
                iRet = Int(Rnd * (Val(5) - (Val(0) - 1)) + (Val(0)))
                'strRandomTime = "17:" & Format(iRet, "00")
                strRandomTime = strEmpID(i, 6)
                strRandomTime = DateAdd("n", iRet, strRandomTime)
                strRandomTime = Format(strRandomTime, "HH:nn")
                
                Call TakeAttendance(strEmpID(i, 0), myDT, CDate(strRandomTime))
                
                con.CommitTrans
            End If
SKIPIT:
            myDT = DateAdd("d", 1, myDT)
        Loop
        
    Next
    MsgBox "Done!!!", vbInformation
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdAdjustAtt_Click()

    On Error GoTo err
    
    If cmbDept.MatchFound = False Then Exit Sub
    If cmbEmp.MatchFound = False Then Exit Sub
    Dim rs As New ADODB.Recordset
    
    Dim strEmpCondition As String
    If cmbEmp.ID = "0" Then
        strEmpCondition = " Where DeptID='" & cmbDept.ID & "'"
    Else
        strEmpCondition = " Where EmpID='" & cmbEmp.ID & "'"
    End If
    strEmpCondition = strEmpCondition & " AND DT Between '" & DtFrom.Value & "' AND '" & DtTo.Value & "'"
    Dim iRet As Integer
    Dim strMyOutTime As String, strMyInTime As String
    Call GetDepartmentTimings
    With rs
        .Open "Select * From VEmpTimes2 " & strEmpCondition & " Order By EmpID,DT,FirstInTime", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
'            If !TotalEntries = 2 Then
'                If !SecondOutTime & "" = "" Then
'                    iRet = Int(Rnd * (Val(9) - (Val(1) - 1)) + (Val(1)))
'                    strMyOutTime = "17:" & Format(iRet, "00")
'                    'con.Execute "Update EmpTimes Set OutTime='" & strMyOutTime & "' Where EntryID=" & !SecondEntryID
'                    Call SaveNow(!EmpId, !DT, CDate(strMyOutTime))
'                End If
'            ElseIf !TotalEntries = 1 Then
'                'Change Ist Out Time, even if it already exists, to adjust those who has like '07:54 & 17:10
                'to do this, Just Pick the FirstInTime, Delete Entry and Add Four Attendances....
                'If !FirstOutTime & "" = "" Then
                strMyInTime = !FirstInTime & ""
                con.Execute "Delete From AttendanceSheet Where EmpID='" & !EmpId & "' AND DT='" & !DT & "'"
                con.Execute "Delete From EmpTimes Where EmpID='" & !EmpId & "' AND DT='" & !DT & "'"
                Call SaveNow(!EmpId, !DT, CDate(strMyInTime))
                iRet = Int(Rnd * (Val(9) - (Val(1) - 1)) + (Val(1)))
                If Format(!DT, "dddd") = "Friday" Then
                    strMyOutTime = DateAdd("n", iRet, "12:15")
                Else
                    strMyOutTime = DateAdd("n", iRet, strInTime)
                End If
                Call SaveNow(!EmpId, !DT, CDate(strMyOutTime))
                'End If
    
                iRet = Int(Rnd * (Val(9) - (Val(1) - 1)) + (Val(1)))
                If Format(!DT, "dddd") = "Friday" Then
                    strMyOutTime = DateAdd("n", -iRet, "13:59")
                Else
                    strMyOutTime = DateAdd("n", -iRet, strOutTime)
                End If
                Call SaveNow(!EmpId, !DT, CDate(strMyOutTime))
                iRet = Int(Rnd * (Val(9) - (Val(1) - 1)) + (Val(1)))
                strMyOutTime = "17:" & Format(iRet, "00")
                Call SaveNow(!EmpId, !DT, CDate(strMyOutTime))
'            End If
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    MsgBox "It's Done.", vbInformation
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CommandButton1_Click()

    On Error GoTo err
    Dim strCond As String
    Dim strEmpID() As String
    Dim i As Integer, j As Integer
    Dim rs As New ADODB.Recordset
    
    strCond = " WHERE Active=1"
    If cmbEmp.ID <> "0" Then
        strCond = strCond & " AND Employees.EmpID='" & cmbEmp.ID & "'"
    Else
        If cmbDept.ID = "0" Then
            '''
        Else
            strCond = strCond & " AND DeptID='" & cmbDept.ID & "'"
        End If
    End If
    strCond = strCond & " AND AttendanceSheet.DT='" & DtFrom.Value & "' AND AttendanceSheet.Attendance=1"
    With rs
        .Open "SELECT Employees.EmpID,LunchInTime,LunchOutTime FROM Employees INNER JOIN VEmpSettings ON Employees.EmpID=VEmpSettings.EmpID INNER JOIN AttendanceSheet ON Employees.EmpID=AttendanceSheet.EmpID" & strCond, con, adOpenForwardOnly, adLockReadOnly
        ReDim strEmpID(.RecordCount, 3)
        Do Until .EOF
            strEmpID(.AbsolutePosition - 1, 0) = !EmpId
            strEmpID(.AbsolutePosition - 1, 1) = !LunchInTime
            strEmpID(.AbsolutePosition - 1, 2) = !LunchOutTime
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Dim myDT As Date, strDay As String
     
    Dim strRandomTime As String
    For i = 0 To UBound(strEmpID) - 1
        myDT = DtFrom.Value
        Do While myDT <= DtTo.Value
            strDay = Format(myDT, "dddd")
            If strDay = "Sunday" Then
                'Skip
            ElseIf IsHoliday(myDT) Then
                'Skip
            Else
                Call StartTrans(con)
'                con.Execute "DELETE FROM AttendanceSheet WHERE EmpID='" & strEmpID(i, 0) & "' AND DT='" & myDT & "'"
'                con.Execute "DELETE FROM EmpTimes WHERE EmpID='" & strEmpID(i, 0) & "' AND DT='" & myDT & "'"
'                con.Execute "DELETE FROM Leaves WHERE EmpID='" & strEmpID(i, 0) & "' AND DT='" & myDT & "'"
                
'                iRet = Int(Rnd * (Val(10) - (Val(0) - 1)) + (Val(0)))
'                strRandomTime = "08:" & Format(iRet, "00")
'                Call TakeAttendance(strEmpID(i, 0), myDT, CDate(strRandomTime))
                
                Call TakeAttendance(strEmpID(i, 0), myDT, IIf(strDay = "Friday", #12:50:00 PM#, CDate(strEmpID(i, 1))))
                Call TakeAttendance(strEmpID(i, 0), myDT, IIf(strDay = "Friday", #2:40:00 PM#, CDate(strEmpID(i, 2))))
                iRet = Int(Rnd * (Val(5) - (Val(0) - 1)) + (Val(0)))
                strRandomTime = "17:" & Format(iRet, "00")
                Call TakeAttendance(strEmpID(i, 0), myDT, CDate(strRandomTime))
                
                con.CommitTrans
            End If
            myDT = DateAdd("d", 1, myDT)
        Loop
        
    Next
    MsgBox "Done!!!", vbInformation
    
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub cmdCompleteAttendance_Click()
    
    On Error GoTo err
    
    Dim strCond As String
    Dim strEmpID() As String
    Dim i As Integer, j As Integer
    Dim rs As New ADODB.Recordset
    
    strCond = " WHERE Active=1"
    If cmbEmp.ID <> "0" Then
        strCond = strCond & " AND Employees.EmpID='" & cmbEmp.ID & "'"
    Else
        If cmbDept.ID = "0" Then
            '''
        Else
            strCond = strCond & " AND DeptID='" & cmbDept.ID & "'"
        End If
    End If
    
    With rs
        ''"INNER JOIN (SELECT EmpID,DT,COUNT(*) AS AttCount FROM EmpTimes WHERE DT BETWEEN '" & DTFrom.Value & "' AND '" & DTTo & "' GROUP BY EmpID,DT HAVING COUNT(*)=1) T1 "
        .Open "SELECT Employees.EmpID,LunchInTime,LunchOutTime,FridayBreakStartTime,FridayBreakEndTime,FridayCheckOut,ExitTime,EnterTime,DT,FirstInTime,FirstOutTime,SecondInTime,SecondOutTime FROM Employees " & _
        "INNER JOIN VEmpSettings ON Employees.EmpID=VEmpSettings.EmpID " & _
        "INNER JOIN (SELECT EmpID,DT,FirstInTime,FirstOutTime,SecondInTime,SecondOutTime FROM VEmpTimes2 WHERE DT BETWEEN '" & DtFrom & "' AND '" & DtTo & "' AND (FirstOutTime IS NULL OR SecondInTime IS NULL OR SecondOutTime IS NULL)) T1 " & _
        "ON Employees.EmpID=T1.EmpID" & strCond & " ORDER BY Employees.EmpID,DT", con, adOpenForwardOnly, adLockReadOnly
        ReDim strEmpID(.RecordCount, 12)
        Do Until .EOF
            strEmpID(.AbsolutePosition - 1, 0) = !EmpId
            strEmpID(.AbsolutePosition - 1, 1) = !LunchInTime & ""
            strEmpID(.AbsolutePosition - 1, 2) = !LunchOutTime & ""
            strEmpID(.AbsolutePosition - 1, 3) = !FridayBreakStartTime & ""
            strEmpID(.AbsolutePosition - 1, 4) = !FridayBreakEndTime & ""
            strEmpID(.AbsolutePosition - 1, 5) = !EnterTime & ""
            strEmpID(.AbsolutePosition - 1, 6) = !ExitTime & ""
            strEmpID(.AbsolutePosition - 1, 7) = !FirstInTime & ""
            strEmpID(.AbsolutePosition - 1, 8) = !FirstOutTime & ""
            strEmpID(.AbsolutePosition - 1, 9) = !SecondInTime & ""
            strEmpID(.AbsolutePosition - 1, 10) = !SecondOutTime & ""
            strEmpID(.AbsolutePosition - 1, 11) = Format(!DT, "dd-MMM-yyyy")
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
'        ReDim strEmpID(1)
'        strEmpID(0) = cmbEmp.ID
'    End If
    
    Dim myDT As Date, strDay As String
     
    Dim strRandomTime As String
    
    For i = 0 To UBound(strEmpID) - 1
    
        myDT = Format(strEmpID(i, 11))
        
        Call StartTrans(con)
                
        con.Execute "DELETE FROM AttendanceSheet WHERE EmpID='" & strEmpID(i, 0) & "' AND DT='" & myDT & "'"
        con.Execute "DELETE FROM EmpTimes WHERE EmpID='" & strEmpID(i, 0) & "' AND DT='" & myDT & "'"
                
'        iRet = Int(Rnd * (Val(9) - (Val(0) - 1)) + (Val(0)))
'        strRandomTime = strEmpID(i, 5)
'        strRandomTime = DateAdd("n", -iRet, strRandomTime)
'        strRandomTime = Format(strRandomTime, "HH:nn")
'
        strRandomTime = strEmpID(i, 7)
        Call TakeAttendance(strEmpID(i, 0), myDT, CDate(strRandomTime))
            
        If strEmpID(i, 1) = strEmpID(i, 2) Then     'No Break
            'Do Nothing....
        Else
            Call TakeAttendance(strEmpID(i, 0), myDT, IIf(strDay = "Friday", CDate(strEmpID(i, 3)), CDate(strEmpID(i, 1))))
            Call TakeAttendance(strEmpID(i, 0), myDT, IIf(strDay = "Friday", CDate(strEmpID(i, 4)), CDate(strEmpID(i, 2))))
        End If
                
        Randomize
        iRet = Int(Rnd * (Val(5) - (Val(0) - 1)) + (Val(0)))
                
        strRandomTime = strEmpID(i, 6)
        strRandomTime = DateAdd("n", iRet, strRandomTime)
        strRandomTime = Format(strRandomTime, "HH:nn")
                
        Call TakeAttendance(strEmpID(i, 0), myDT, CDate(strRandomTime))
                
        con.CommitTrans
            
    Next
    MsgBox "Done!!!", vbInformation
    
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub Form_Load()

    With cmbDept
        .ListHeight = 1600
        .ClearVals
        .AddItem "<All Departments>", "0"
        .AddVals con, "Name", "Departments", "DeptID", " Order By Name"
    End With
    cmbEmp.ListHeight = 1250
      
    DtFrom = GetServerDate
    DtTo = DtFrom
    
End Sub


Private Sub GetDepartmentTimings()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim bExempt As Boolean
    With rs
        If cmbEmp.ID = "0" Then
            .Open "Select LunchInTime,LunchOutTime From GeneralSettings Where DeptID='" & cmbDept.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        Else
            .Open "Select ExemptSettings From Employees Where EmpID='" & cmbEmp.ID & "'", con, adOpenForwardOnly, adLockReadOnly
            bExempt = !ExemptSettings
            .Close
            If bExempt Then
                .Open "Select LunchInTime,LunchOutTime From EmpSettings Where EmpID='" & cmbEmp.ID & "'", con, adOpenForwardOnly, adLockReadOnly
            Else
                .Open "Select LunchInTime,LunchOutTime From GeneralSettings Where DeptID='" & cmbDept.ID & "'", con, adOpenForwardOnly, adLockReadOnly
            End If
        End If
        strInTime = !LunchInTime
        strOutTime = !LunchOutTime
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub SaveNow(EmpId As String, AttDate As Date, AttTime As Date)

    On Error GoTo err

    Dim ExtraDeds As Double, ExtraHrs As Double
    ExtraDeds = 0
    ExtraHrs = 0
    Dim MultipleAllowed As Boolean
    Dim bDeductLHrs As Boolean
    Dim bExempt As Boolean
    Dim DeptID As String
    Dim dLateHrs As Double
    Dim TotalEntries As Integer
    Dim AttRecs As Integer
    Dim bOverTime As Boolean
    Dim strEmpInTime As String, strEmpOutTime As String
    Dim strEmpLunchIn As String, strEmpLunchOut As String
    bOverTime = False
    '''''''''''   This Will Is The Employee Exists Or Not    '''''''''''
    
    With con.Execute("Select DeptID,EmpID,EmpType,ExemptSettings FROM Employees WHERE EmpID='" & EmpId & "'")
        If .EOF Then
            Exit Sub
        Else
            EmpId = ![EmpId] & ""
            Emptype = ![Emptype] & ""
            bExempt = ![ExemptSettings] & ""
            DeptID = ![DeptID] & ""
        End If
        .Close
        
        If bExempt Then
            .Open "Select EnterTime,ExitTime,LunchInTime,LunchOutTime,DeductLHrs From EmpSettings Where EmpID='" & EmpId & "'", con, adOpenForwardOnly, adLockReadOnly
        Else
            .Open "Select EnterTime,ExitTime,LunchInTime,LunchOutTime,DeductLHrs From GeneralSettings Where DeptID='" & DeptID & "'", con, adOpenForwardOnly, adLockReadOnly
        End If
        
        If .EOF Then
            bDeductLHrs = True
        Else
            bDeductLHrs = ![DeductLHrs]
            strEmpInTime = Format(!EnterTime, "HH:NN") '"08:00" 'Format(!EnterTime, "HH:NN")
            strEmpOutTime = Format(!ExitTime, "HH:NN") '"17:00" 'Format(!ExitTime, "HH:NN")
            strEmpLunchIn = Format(!LunchInTime, "HH:NN") '"12:30" 'Format(!LunchInTime, "HH:NN")
            strEmpLunchOut = Format(!LunchOutTime, "HH:NN") '"13:30" 'Format(!LunchOutTime, "HH:NN")
        End If
        .Close
    End With
    
    Dim Rec As New ADODB.Recordset, CurrTime As String, Hrs As Double, Sql As String, Cond As String
    Dim TimeDiff As Integer

    Cond = "WHERE EmpID='" & EmpId & "' AND DT='" & AttDate & "'"

    CurrTime = Format(AttTime, "HH:NN")

    TimeDiff = Val(con.Execute("Select TimeDiff FROM TimeSettings").Fields(0) & "")
        
    With Rec
        .Open "Select TotalEntries From VEmpDailyEntries Where EmpID='" & EmpId & "' and DT='" & AttDate & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            TotalEntries = 0
        Else
            TotalEntries = .Fields(0) & ""
        End If
        .Close
        
        .Open "SELECT * FROM EmpTimes WHERE EntryID=(SELECT Max(EntryID) FROM EmpTimes " & Cond & ")", con, adOpenForwardOnly, adLockReadOnly
            
        If TotalEntries > 0 Then
            
            If TotalEntries = 1 And !OutTime & "" = "" Then 'Means It's Lunch In Entry...
                '12:50 to 1:10
'                If CurrTime < #12:25:00 PM# Then
                If CurrTime < CDate(DateAdd("N", -5, strEmpLunchIn)) Then
                    ExtraDeds = 0 'DateDiff("n", CurrTime, "12:55") / 60
                ElseIf CurrTime > CDate(DateAdd("N", 5, strEmpLunchIn)) Then
                    ExtraDeds = 0 'DateDiff("n", "13:00", CurrTime) / 60
                End If
                
                If DateDiff("n", strEmpInTime, CDate(!InTime)) > 0 And DateDiff("n", strEmpInTime, CDate(!InTime)) < 11 Then
                    ExtraHrs = DateDiff("n", strEmpInTime, CDate(!InTime)) / 60
                ElseIf DateDiff("n", strEmpInTime, CDate(!InTime)) < 0 Then
                    ExtraDeds = ExtraDeds + DateDiff("n", CDate(!InTime), strEmpInTime) / 60
                End If
                
                If Format(AttDate, "dddd") = "Friday" Then
                    ExtraHrs = ExtraHrs + 0.25
                    If CurrTime > #12:15:00 PM# Then
                        Hrs = DateDiff("n", CDate(!InTime), #12:15:00 PM#)
                    Else
                        Hrs = DateDiff("n", CDate(!InTime), CurrTime)
                    End If
                Else
                    If CurrTime > CDate(DateAdd("N", -6, strEmpLunchIn)) Then '#12:24:00 PM#
                        Hrs = DateDiff("n", CDate(!InTime), strEmpLunchIn) '#12:30:00 PM#
                    Else
                        Hrs = DateDiff("n", CDate(!InTime), CurrTime)
                    End If
                End If
                
            ElseIf TotalEntries = 2 And !OutTime & "" = "" Then 'Means It's Lunch Out Entry
                If Format(AttDate, "dddd") = "Friday" Then
                    ExtraHrs = ExtraHrs + 0.5
                    If CDate(!InTime) > #2:05:00 PM# Then
                        ExtraDeds = 0 'DateDiff("n", "14:30", CDate(!InTime)) / 60
                    ElseIf CDate(!InTime) < #2:00:00 PM# Then
                        ExtraDeds = DateDiff("n", CDate(!InTime), "14:00") / 60
                    End If
                    If DateDiff("n", "14:00", CDate(!InTime)) > 0 And DateDiff("n", "14:00", CDate(!InTime)) < 6 Then
                        ExtraHrs = ExtraHrs + DateDiff("n", "14:00", CDate(!InTime)) / 60
                    End If
                Else
                    If DateDiff("n", strEmpLunchOut, CDate(!InTime)) > 0 And DateDiff("n", strEmpLunchOut, CDate(!InTime)) < 6 Then '"13:30"
                        ExtraHrs = DateDiff("n", strEmpLunchOut, CDate(!InTime)) / 60 '"13:30"
                    End If
                    
                    If CDate(!InTime) > CDate(DateAdd("n", 5, strEmpLunchOut)) Then  '#1:35:00 PM#
                        ExtraDeds = 0 'DateDiff("n", "14:00", CDate(!InTime)) / 60
                    ElseIf !InTime < strEmpLunchOut Then '#1:30:00 PM#
                        ExtraDeds = DateDiff("n", CDate(!InTime), strEmpLunchOut) / 60 '"13:30"
                    End If
                End If
                
                If CurrTime > strEmpOutTime Then '#5:00:00 PM#
                    ExtraDeds = ExtraDeds + 0 'DateDiff("n", "17:00", CurrTime) / 60
                End If
                
                If CurrTime > strEmpOutTime Then '#5:00:00 PM#
                    Hrs = DateDiff("n", CDate(!InTime), strEmpOutTime) '#5:00:00 PM#
                Else
                    Hrs = DateDiff("n", CDate(!InTime), CurrTime)
                End If
            End If
            'Check ... Is This Out Time Entry ?
            
            If TotalEntries < 3 Then
                If !OutTime & "" = "" Then
                    'Hrs = DateDiff("n", CDate(!InTime), CDate(CurrTime))
                
                    'Check for the Minimum Time Difference
                    If Hrs < TimeDiff Then GoTo END_OF_PROCEDURE
                
                    Hrs = Round(Hrs / 60, 2)
                            
                    'Update Out Time and Do Nothing Else
                
                    Sql = "UPDATE EmpTimes SET OutTime='" & CurrTime & "',Hrs=" & Hrs & ",PayableHrs=" & (Hrs + ExtraHrs - ExtraDeds) & " WHERE EntryID=" & Val(!EntryID & "") & ""
                    con.Execute Sql
                    'con.Execute "Update AttendanceSheet Set Attendance=1 Where EmpID='" & EmpId & "' AND DT='" & AttDate & "'"
                    GoTo END_OF_PROCEDURE
                Else
                    'Check for the Minimum Time Difference
                
                    Hrs = DateDiff("n", CDate(!OutTime), CDate(CurrTime))
                    If TotalEntries <> 2 Then
                        If Hrs < TimeDiff Then GoTo END_OF_PROCEDURE
                    End If
                End If
            End If
        Else
            If bDeductLHrs And DateDiff("n", strEmpInTime, CurrTime) > 10 Then '"08:00"
                dLateHrs = DateDiff("n", strEmpInTime, CurrTime) / 60 '"08:00"
            Else
                dLateHrs = 0
            End If
            AttRecs = Val(con.Execute("Select Count(*) From AttendanceSheet Where EmpID='" & EmpId & "' AND DT='" & AttDate & "'").Fields(0).Value & "")
            If AttRecs <> 0 Then con.Execute "Delete From AttendanceSheet Where EmpID='" & EmpId & "' AND DT='" & AttDate & "'"
            
            con.Execute "INSERT INTO AttendanceSheet(EmpID,DT,Attendance,LateHours) VALUES('" & EmpId & "','" & AttDate & "',1," & _
             dLateHrs & ")", a
            
        End If

        '.Close
        
        If Emptype = 2 Then
            
            ' If No Previous Entry Found Accept The Entry
            If TotalEntries = 0 Then GoTo ACCEPT_ENTRY
        
            'Remember The Total Entries
        
            ' If this is Second Time Entry Accept It
            If TotalEntries < 2 Then GoTo ACCEPT_ENTRY

        End If
            
        ' If Two Entries Have Been Made ... Check For Over Time Authorization
        If TotalEntries >= 2 Then
            Dim rsOT As New ADODB.Recordset
            Dim bAuthorized As Boolean
            With rsOT
                .Open "Select * From VOverTimeAuth Where EmpID='" & EmpId & "' and DT='" & AttDate & "'", con, adOpenForwardOnly, adLockReadOnly
                ' If Not Authorized for Over Time Neglect This Entry
                If .EOF Then
                    bAuthorized = False
                Else
                    bAuthorized = True
                End If
                rsOT.Close
                Set rsOT = Nothing
                If bAuthorized = False Then
                    MsgBox "Not Authorized.", vbInformation
                    GoTo END_OF_PROCEDURE
                End If
            End With
            'If Val(![Authority] & "") <> 1 Then GoTo END_OF_PROCEDURE
            
            ' Other Wise Accept The Entry
            If !OutTime & "" <> "" Then
                con.Execute "INSERT INTO EmpTimes(EmpID,DT,InTime,ManualEntry,OverTime,UserName,ComputerName) VALUES('" & EmpId _
                 & "','" & AttDate & "','" & CurrTime & "',1,1,'" & CurrentUserName & "','" & strComputerName & "')"
            Else
                Hrs = DateDiff("n", CDate(!InTime), CurrTime)
                If Hrs < TimeDiff Then GoTo END_OF_PROCEDURE
                Hrs = Round(Hrs / 60, 2)
                con.Execute "UPDATE EmpTimes SET OutTime='" & CurrTime & "',Hrs=" & Hrs & ",PayableHrs=" & (Hrs) & " WHERE EntryID=" & Val(!EntryID & "") & ""
            End If
            GoTo END_OF_PROCEDURE
            bOverTime = True
        End If
        
    End With

    ' Enter The New Entry With In Time

ACCEPT_ENTRY:
    
    Sql = "INSERT INTO EmpTimes(EmpID,DT,InTime,ManualEntry,UserName,ComputerName) VALUES('" & EmpId _
     & "','" & AttDate & "','" & CurrTime & "',1,'" & CurrentUserName & "','" & strComputerName & "')"
    
    con.Execute Sql
'    TotalEntries = Val(con.Execute("Select TotalEntries From VEmpDailyEntries Where EmpID='" & EmpID & "' and DT='" & AttDate & "'").Fields(0).Value & "")
'
'    If TotalEntries = 1 Then    'That Means that it's first entry of the day. So Calculate LateHours...
'        If DateDiff("n", "08:00", CurrTime) > 15 Then
'            con.Execute "Update AttendanceSheet Set LateHours=" & DateDiff("n", "08:00", CurrTime) / 60 & " Where EmpID='" & EmpID & "' and DT='" & AttDate & "'"
'        End If
'    End If
    
END_OF_PROCEDURE:
    
    If Rec.State = 1 Then Rec.Close
    Set Rec = Nothing
    
    Exit Sub
err:
    Debug.Print err.Description
    'So That Next Entries Should Not Be Skipped
    Resume Next
    
End Sub


