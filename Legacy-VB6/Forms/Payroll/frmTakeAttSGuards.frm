VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmTakeAttSGuards 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Take Attendance Of Security Guards"
   ClientHeight    =   3375
   ClientLeft      =   4020
   ClientTop       =   4170
   ClientWidth     =   7770
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTakeAttSGuards.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   7770
   Begin VB.CheckBox Check1 
      Caption         =   "Regular Employees"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5370
      TabIndex        =   11
      Top             =   450
      Width           =   1785
   End
   Begin ComboList.Usercontrol1 cmbEmp1 
      Height          =   285
      Left            =   405
      TabIndex        =   2
      Top             =   720
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   330
      Left            =   4530
      TabIndex        =   10
      Top             =   2835
      Width           =   2565
   End
   Begin VB.CommandButton cmdAddReliever 
      Caption         =   "Take Att Of Reliever"
      Height          =   330
      Left            =   405
      TabIndex        =   9
      Top             =   2835
      Width           =   2565
   End
   Begin VB.ComboBox cmbRestDay 
      Height          =   315
      Left            =   5415
      TabIndex        =   7
      Top             =   1530
      Width           =   1695
   End
   Begin VB.ComboBox cmbShift 
      Height          =   315
      Left            =   4140
      TabIndex        =   5
      Top             =   1545
      Width           =   1245
   End
   Begin VB.ComboBox cmbPeriod 
      Height          =   315
      Left            =   405
      TabIndex        =   4
      Top             =   1560
      Width           =   3630
   End
   Begin VB.CommandButton cmdAllatOnce 
      Caption         =   "Take Att"
      Height          =   330
      Left            =   4530
      TabIndex        =   0
      Top             =   1920
      Width           =   2565
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   405
      TabIndex        =   12
      Top             =   1560
      Visible         =   0   'False
      Width           =   3630
      _ExtentX        =   6403
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   60882947
      CurrentDate     =   38597
   End
   Begin VB.Label Label1 
      Caption         =   "Rest Day"
      Height          =   180
      Index           =   4
      Left            =   5430
      TabIndex        =   8
      Top             =   1320
      Width           =   1110
   End
   Begin VB.Label Label1 
      Caption         =   "Shift"
      Height          =   180
      Index           =   1
      Left            =   4155
      TabIndex        =   6
      Top             =   1335
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Security Guard"
      Height          =   195
      Index           =   3
      Left            =   435
      TabIndex        =   3
      Top             =   465
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "Period"
      Height          =   180
      Index           =   2
      Left            =   435
      TabIndex        =   1
      Top             =   1335
      Width           =   1845
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   300
      X2              =   7170
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   300
      X2              =   7170
      Y1              =   2415
      Y2              =   2415
   End
End
Attribute VB_Name = "frmTakeAttSGuards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim dtHolidays() As Date

Private Sub Check1_Click()

    If Check1.value = vbChecked Then
        DTPicker1.Visible = True
        cmbPeriod.Visible = False
        cmbEmp1.ClearVals
        cmbEmp1.AddVals con, "'{ ' + EmpID + ' } ' + Name + ' { ' + Designation + ' }'", "Employees", "EmpID", " Where Active=1"
    Else
        DTPicker1.Visible = False
        cmbPeriod.Visible = True
        cmbEmp1.ClearVals
        cmbEmp1.AddVals con, "'{ ' + EmpID + ' } ' + Name + ' { ' + Designation + ' }'", "Employees", "EmpID", " Where Designation='Security Guard' AND Active=1"
    End If
    
End Sub

Private Sub cmdAddReliever_Click()
    
    On Error GoTo err
    
    Dim bNextDay As Boolean
    Dim oInTime As String
    Dim oOutTime As String
    
    Dim iRet As Integer
    Dim dHrs As Double
    
    Dim StartDT As Date
    Dim EndDT As Date
    Dim DT As Date
    Dim iDay As Integer
    If Check1.value = vbChecked Then
        If Format(DT, "dddd") = "Sunday" Then
            MsgBox "Can't Take Attendance Of Sunday.", vbInformation
            Exit Sub
        Else
            Call StartTrans(con)
            DT = DTPicker1.value
            con.Execute "Delete From AttendanceSheet Where EmpID='" & cmbEmp1.ID & "' AND DT='" & DT & "'"
            con.Execute "Delete From EmpTimes Where EmpID='" & cmbEmp1.ID & "' AND DT='" & DT & "'"
            con.Execute "Delete From Leaves Where EmpID='" & cmbEmp1.ID & "' AND DT='" & DT & "'"
            
            con.Execute "Insert Into AttendanceSheet(EmpID,DT,Attendance) Values('" & cmbEmp1.ID & "','" & DT & "',1)"
            
            iDay = Weekday(DT, vbMonday)
            If iDay = 1 Or iDay = 6 Then    'Shift A
                oInTime = "08:00"
                oOutTime = "16:00"
                bNextDay = False
            ElseIf iDay = 2 Or iDay = 5 Then
                oInTime = "16:00"
                oOutTime = "00:00"
                bNextDay = True
            ElseIf iDay = 3 Or iDay = 4 Then
                oInTime = "00:00"
                oOutTime = "08:00"
                bNextDay = True
            End If
            iRet = Int(Rnd * (Val(15) - (Val(1) - 1)) + (Val(1)))
            InTime = Format(DateAdd("n", iRet, oInTime), "HH:nn")
            iRet = Int(Rnd * (Val(15) - (Val(1) - 1)) + (Val(1)))
            OutTime = Format(DateAdd("n", iRet, oOutTime), "HH:nn")
            
            dHrs = DateDiff("n", InTime, OutTime) / 60
            dHrs = Abs(dHrs)
'            If bNextDay Then
'                dHrs = DateDiff("n", DT & " " & InTime, DateAdd("d", 1, DT) & " " & OutTime) / 60
'            Else
'                dHrs = DateDiff("n", InTime, OutTime) / 60
'            End If
            
            con.Execute "Insert Into EmpTimes(EmpID,DT,InTime,OutTime,Hrs,PayableHrs) Values('" & _
             cmbEmp1.ID & "','" & DT & "','" & InTime & "','" & OutTime & "'," & dHrs & ",8)"

            con.CommitTrans
            MsgBox "Attendance Done.", vbInformation
            Exit Sub
        End If
        
    End If
    
    StartDT = Left(cmbPeriod, 11)
    EndDT = Right(cmbPeriod, 11)
    
    
    DT = StartDT
    
    Call StartTrans(con)
    
    'If Leaves Are Feeded Then Don't Lose Those Leaves and Don't Take Attendance Of Those Dates.
    'Just Hold The Dates In AbsDays() Array......
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select DT From Leaves Where EmpID='" & cmbEmp1.ID & "' And DT Between '" & StartDT & "' AND '" & EndDT & "'", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            TotAbs = .RecordCount
            ReDim AbsDays(TotAbs)
            For i = 1 To .RecordCount
                AbsDays(i) = ![DT]
                'Con.Execute "Insert Into AttendanceSheet(EmpID,DT,Attendance) Values('" & cmbEmp.ID & "','" & ![DT] & "',2)"
                .MoveNext
            Next
        End If
        .Close
        
        .Open "Select DT From AttendanceSheet Where EmpID='" & cmbEmp1.ID & "' And DT Between '" & StartDT & "' AND '" & EndDT & "' AND Attendance=0", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            ReDim Preserve AbsDays(TotAbs + .RecordCount)
            For i = 1 To .RecordCount
                AbsDays(i + TotAbs) = ![DT]
                .MoveNext
            Next
        End If
        TotAbs = TotAbs + .RecordCount
        .Close
        
        .Open "Select DT From Holidays Where DT Between '" & StartDT & "' AND '" & EndDT & "'", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            ReDim Preserve AbsDays(TotAbs + .RecordCount)
            For i = 1 To .RecordCount
                AbsDays(i + TotAbs) = ![DT]
                .MoveNext
            Next
        End If
        TotAbs = TotAbs + .RecordCount
        .Close
    End With
    
    
    While DT <= EndDT
        For i = 1 To TotAbs
            If DT = AbsDays(i) Then GoTo TheEnd
            
        Next
        If Format(DT, "dddd") = "Sunday" Then
        Else
            con.Execute "Insert Into AttendanceSheet(EmpID,DT,Attendance) Values('" & cmbEmp1.ID & "','" & DT & "',1)"
            
            iDay = Weekday(DT, vbMonday)
            
            If iDay = 1 Or iDay = 6 Then    'Shift A
                oInTime = "08:00"
                oOutTime = "16:00"
                bNextDay = False
            ElseIf iDay = 2 Or iDay = 5 Then
                oInTime = "16:00"
                oOutTime = "00:00"
                bNextDay = True
            ElseIf iDay = 3 Or iDay = 4 Then
                oInTime = "00:00"
                oOutTime = "08:00"
                bNextDay = False
            End If
            
            iRet = Int(Rnd * (Val(15) - (Val(1) - 1)) + (Val(1)))
            InTime = Format(DateAdd("n", iRet, oInTime), "HH:nn")
            iRet = Int(Rnd * (Val(15) - (Val(1) - 1)) + (Val(1)))
            OutTime = Format(DateAdd("n", iRet, oOutTime), "HH:nn")
            
'            dHrs = DateDiff("n", InTime, OutTime) / 60
'            dHrs = Abs(dHrs)
                        
            If bNextDay Then
                dHrs = DateDiff("n", DT & " " & InTime, DateAdd("d", 1, DT) & " " & OutTime) / 60
            Else
                dHrs = DateDiff("n", InTime, OutTime) / 60
            End If
            
            con.Execute "Insert Into EmpTimes(EmpID,DT,InTime,OutTime,Hrs,PayableHrs) Values('" & _
             cmbEmp1.ID & "','" & DT & "','" & InTime & "','" & OutTime & "'," & dHrs & ",8)"
             
        End If
        
TheEnd:
        DT = DateAdd("d", 1, DT)
    Wend
    
    
    con.CommitTrans
    
    MsgBox "Attendance Done.", vbInformation
    
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub cmdAllatOnce_Click()

    On Error GoTo err
    
    Dim oInTime As String
    Dim oOutTime As String
    Dim bNextDay As Boolean
    Dim DT As Date
    Dim TotAbs As Integer
    Dim AbsDays() As Date
    
    Dim iRet As Integer
    Dim dHrs As Double
    
    Dim StartDT As Date
    Dim EndDT As Date
    
    If cmbShift.ListIndex = 0 Then
        oInTime = "08:00"
        oOutTime = "16:00"
        bNextDay = False
    ElseIf cmbShift.ListIndex = 1 Then
        oInTime = "16:00"
        oOutTime = "00:00"
        bNextDay = True
    ElseIf cmbShift.ListIndex = 2 Then
        oInTime = "00:00"
        oOutTime = "08:00"
        bNextDay = False
    End If
        
    If Check1.value = vbChecked Then
        DT = DTPicker1.value
        Call StartTrans(con)
        con.Execute "Delete From AttendanceSheet Where EmpID='" & cmbEmp1.ID & "' And DT='" & DT & "'"
        con.Execute "Delete From EmpTimes Where EmpID='" & cmbEmp1.ID & "' And DT='" & DT & "'"
        con.Execute "Delete From Leaves Where EmpID='" & cmbEmp1.ID & "' And DT='" & DT & "'"
        
        con.Execute "Insert Into AttendanceSheet(EmpID,DT,Attendance) Values('" & cmbEmp1.ID & "','" & DT & "',1)"
        
        iRet = Int(Rnd * (Val(15) - (Val(1) - 1)) + (Val(1)))
        InTime = Format(DateAdd("n", iRet, oInTime), "HH:nn")
        iRet = Int(Rnd * (Val(15) - (Val(1) - 1)) + (Val(1)))
        OutTime = Format(DateAdd("n", iRet, oOutTime), "HH:nn")
        If bNextDay Then
            dHrs = DateDiff("n", DT & " " & InTime, DateAdd("d", 1, DT) & " " & OutTime) / 60
        Else
            dHrs = DateDiff("n", InTime, OutTime) / 60
        End If
        con.Execute "Insert Into EmpTimes(EmpID,DT,InTime,OutTime,Hrs,PayableHrs) Values('" & _
         cmbEmp1.ID & "','" & DT & "','" & InTime & "','" & OutTime & "'," & dHrs & ",8)"
        
        con.CommitTrans
        MsgBox "Attendance Done.", vbInformation
        Exit Sub
        
    End If
    
    
    StartDT = Left(cmbPeriod, 11)
    EndDT = Right(cmbPeriod, 11)
    
    DT = StartDT
    Call StartTrans(con)
    
    'If Leaves Are Feeded Then Don't Lose Those Leaves and Don't Take Attendance Of Those Dates.
    'Just Hold The Dates In AbsDays() Array......
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select DT From Leaves Where EmpID='" & cmbEmp1.ID & "' And DT Between '" & StartDT & "' AND '" & EndDT & "'", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            TotAbs = .RecordCount
            ReDim AbsDays(TotAbs)
            For i = 1 To .RecordCount
                AbsDays(i) = ![DT]
                'Con.Execute "Insert Into AttendanceSheet(EmpID,DT,Attendance) Values('" & cmbEmp.ID & "','" & ![DT] & "',2)"
                .MoveNext
            Next
        End If
        .Close
        
        .Open "Select DT From AttendanceSheet Where EmpID='" & cmbEmp1.ID & "' And DT Between '" & StartDT & "' AND '" & EndDT & "' AND Attendance=0", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            ReDim Preserve AbsDays(TotAbs + .RecordCount)
            For i = 1 To .RecordCount
                AbsDays(i + TotAbs) = ![DT]
                .MoveNext
            Next
        End If
        TotAbs = TotAbs + .RecordCount
        .Close
        
        .Open "Select DT From Holidays Where DT Between '" & StartDT & "' AND '" & EndDT & "'", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            ReDim Preserve AbsDays(TotAbs + .RecordCount)
            For i = 1 To .RecordCount
                AbsDays(i + TotAbs) = ![DT]
                .MoveNext
            Next
        End If
        TotAbs = TotAbs + .RecordCount
        .Close
    End With
    
    
    While DT <= EndDT
        For i = 1 To TotAbs
            If DT = AbsDays(i) Then GoTo TheEnd
            
        Next
        If Format(DT, "dddd") = cmbRestDay Then
        Else
            con.Execute "Insert Into AttendanceSheet(EmpID,DT,Attendance) Values('" & cmbEmp1.ID & "','" & DT & "',1)"
            iRet = Int(Rnd * (Val(15) - (Val(1) - 1)) + (Val(1)))
            InTime = Format(DateAdd("n", iRet, oInTime), "HH:nn")
            iRet = Int(Rnd * (Val(15) - (Val(1) - 1)) + (Val(1)))
            OutTime = Format(DateAdd("n", iRet, oOutTime), "HH:nn")
            If bNextDay Then
                dHrs = DateDiff("n", DT & " " & InTime, DateAdd("d", 1, DT) & " " & OutTime) / 60
            Else
                dHrs = DateDiff("n", InTime, OutTime) / 60
            End If
            con.Execute "Insert Into EmpTimes(EmpID,DT,InTime,OutTime,Hrs,PayableHrs) Values('" & _
             cmbEmp1.ID & "','" & DT & "','" & InTime & "','" & OutTime & "'," & dHrs & ",8)"
             
        End If
TheEnd:
        DT = DateAdd("d", 1, DT)
    Wend
    
    con.CommitTrans
    
    MsgBox "Attendance Done.", vbInformation
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    cmbEmp1.ListHeight = 2400
    cmbEmp1.AddVals con, "'{ ' + EmpID + ' } ' + Name ", "Employees", "EmpID", " Where Designation='Security Guard' AND Active=1"
    
    With cmbShift
        .AddItem "A"
        .AddItem "B"
        .AddItem "C"
    End With
    
    With cmbRestDay
        .AddItem "Monday"
        .AddItem "Tuesday"
        .AddItem "Wednesday"
        .AddItem "Thursday"
        .AddItem "Friday"
        .AddItem "Saturday"
        .AddItem "Sunday"
    End With
    
    Dim DT As Date, DT1 As Date
    DT = DateAdd("m", -2, Date) '#1/1/2005#
    
    With cmbPeriod
        For i = 1 To 9
            DT1 = DateSerial(year(DT), Month(DT), 15)
            .AddItem Format(DT, "dd-MMM-yyyy") & " to " & Format(DT1, "dd-MMM-yyyy")
            DT = DateAdd("d", 1, DT1)
            DT1 = DateSerial(year(DT), Month(DT) + 1, 0)
            .AddItem Format(DT, "dd-MMM-yyyy") & " to " & Format(DT1, "dd-MMM-yyyy")
            DT = DateAdd("d", 1, DT1)
        Next
    End With
    
End Sub
