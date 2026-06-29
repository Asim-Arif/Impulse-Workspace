VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmEmpOnLeaves 
   ClientHeight    =   7110
   ClientLeft      =   450
   ClientTop       =   5625
   ClientWidth     =   11880
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   7110
   ScaleMode       =   0  'User
   ScaleWidth      =   10641.08
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbLeaveWithin 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   900
      Width           =   2805
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   900
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   11670
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employees On Leave"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   390
         Index           =   1
         Left            =   3930
         TabIndex        =   5
         Top             =   120
         Width           =   3225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employees On Leave"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Index           =   2
         Left            =   3945
         TabIndex        =   6
         Top             =   135
         Width           =   3225
      End
   End
   Begin VB.Frame FLV 
      Height          =   5670
      Left            =   75
      TabIndex        =   0
      Top             =   1320
      Width           =   11670
      Begin VB.Frame fButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   630
         Left            =   8310
         TabIndex        =   7
         Top             =   4905
         Width           =   3240
         Begin MSForms.CommandButton cmdPrint 
            Height          =   375
            Left            =   60
            TabIndex        =   8
            Top             =   150
            Width           =   1500
            ForeColor       =   0
            Caption         =   "Print         "
            PicturePosition =   327683
            Size            =   "2646;661"
            Accelerator     =   80
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdClose 
            Height          =   375
            Left            =   1635
            TabIndex        =   9
            Top             =   150
            Width           =   1500
            ForeColor       =   0
            Caption         =   " Close    "
            PicturePosition =   327683
            Size            =   "2646;661"
            Accelerator     =   67
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   4695
         Left            =   90
         TabIndex        =   1
         Top             =   195
         Width           =   11460
         _ExtentX        =   20214
         _ExtentY        =   8281
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15199215
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "1500.09"
            Text            =   "S. No"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "2200"
            Text            =   "Emp Name"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "2000"
            Text            =   "Designation"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "700.15"
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Leave Type"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "App. By"
            Object.Width           =   2822
         EndProperty
      End
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   6945
      TabIndex        =   3
      Top             =   855
      Width           =   1500
      Caption         =   "Refresh       "
      PicturePosition =   327683
      Size            =   "2646;635"
      Accelerator     =   82
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   630
      Width           =   6720
      BackColor       =   11517387
      Caption         =   " Department                                                                Leave Appr."
      Size            =   "11853;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmEmpOnLeaves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DT As Date
Dim Loading As Boolean
Dim StrSQL As String

Private Sub cmbShift_Change()
    Call FillList
End Sub

Private Sub cmdCheckAll_Click()
    With LV
        For i = 1 To .ListItems.count
            .ListItems(i).Checked = True
        Next i
    End With
End Sub

Private Sub cmdClearAll_Click()
    With LV
        For i = 1 To .ListItems.count
            .ListItems(i).Checked = False
        Next i
    End With
End Sub

Private Sub cmbdept_matched()
    If Loading Then Exit Sub
    Call cmdRefresh_Click
End Sub

Private Sub cmbMonth_Change()
    If Loading Or cmbMonth.MatchFound = False Then Exit Sub
    Dim iDays As Integer
    Dim DT As Date
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    iDays = DateAdd("m", 1, DT) - DT
    cmbMonth.Tag = iDays
    txtMax = iDays
    txtMin = "1"
End Sub

Private Sub cmbMonth_Click()

    Call cmdRefresh_Click
End Sub

Private Sub cmbYear_Change()
    If Loading Then Exit Sub
    Dim iDays As Integer
    Dim DT As Date
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    iDays = DateAdd("m", 1, DT) - DT
    txtMax = iDays
    txtMin = "1"
End Sub

Private Sub cmdCancel_Click()
    cmdSave.Visible = False
    cmdCancel.Visible = False
    cmdRefresh.Visible = True
    cmdClose.Visible = True
    'cmdPrint.Visible = True
End Sub

Private Sub cmbYear_Click()
Call cmdRefresh_Click
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
    
    Dim strRet As String
    'strRet = ShowPopUpMenu(Me.Left + cmdPrint.Left, Me.Top + cmdPrint.Top - cmdPrint.Height, "&Print Material Detail", "-", "Print Material &Rcving Detail", "-", "Print Printing &Costing")
    
    strRet = ShowPopUpMenu(Me.Left + fButs1.Left + cmdPrint.Left, Me.Top + fButs1.Top + FLV.Top + cmdPrint.Top - 1000, "&Print Report", "-", "Print Leave Balance", "-", "Print Leave Balance of " & LV.SelectedItem.SubItems(1))
    
    Dim rpt As craxddrt.Report, frmRpt As New frmPrevRpt
    
    If strRet = "&Print Report" Then
        Me.MousePointer = vbHourglass
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpsOnLeave.rpt")
'        On Error Resume Next
'        SrcReport.Database.Tables(1).ConnectBufferString = "DSN=DSN_Vision;;User ID=sa;;Database=Vision;;Password=visionasim"
'        rpt.SQLQueryString = StrSQL
        rpt.FormulaFields(1).Text = "'" & cmbLeaveWithin.Text & "'"
        frmRpt.ShowReport "", rpt, StrSQL
        Me.MousePointer = vbNormal
    ElseIf strRet = "Print Leave Balance" Then
        Call PrintEmpLeaveBalance
    ElseIf strRet = "Print Leave Balance of " & LV.SelectedItem.SubItems(1) Then
        Call PrintEmpLeaveBalance(LV.SelectedItem.Tag)
    End If
    
    Exit Sub

err:
    MsgBox err.Description

End Sub

Public Sub cmdRefresh_Click()
    Call FillList
End Sub

Private Sub cmdSave_Click()
'    On Error GoTo ERR
'
'
'    con.BeginTrans
'
'
'    Cond = IIf(cmbDept.ID = "0", "", " and EmpID Like '" & cmbDept.ID & "%'")
'
'    con.Execute "Delete From SSMonthly Where Month(dt)=" & Month(Dated) & " and Year(dt)=" & year(Dated) & Cond
'
'
'    With LV
'        For i = 1 To .ListItems.Count
'            If .ListItems(i).Checked Then
'                con.Execute "insert into SSMonthly(EmpID,DT,EmpStatus,Salary," & _
'                "DaysWorked,NetSalary,SSAmt) Values('" & _
'                .ListItems(i).key & "','" & Dated & "'," & _
'                IIf(.ListItems(i).SubItems(5) = "Permanent", 1, 0) & "," & _
'                Val(.ListItems(i).SubItems(6)) & "," & Val(.ListItems(i).SubItems(7)) & _
'                "," & Val(.ListItems(i).SubItems(8)) & "," & Val(.ListItems(i).SubItems(9)) & ")"
'            End If
'        Next i
'    End With
'
'    con.CommitTrans
'    MsgBox "Saved Successfully", vbInformation
'    cmdSave.Visible = False
'
'    Exit Sub
'ERR:
'    MsgBox ERR.Description

End Sub



Private Sub DTPicker1_Change()
    Call FillList
End Sub

Private Sub Form_Activate()
    Loading = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()

    For i = 1 To LV.ColumnHeaders.count
        LV.ColumnHeaders(i).Tag = LV.ColumnHeaders(i).Width
    Next
    
    Loading = True
    
    With cmbLeaveWithin
        .AddItem "In Next 7 Days"
        .AddItem "In Next 15 Days"
        .AddItem "In Next 30 Days"
        .ListIndex = 0
    End With
    
    With cmbDept
        .ListHeight = 2200
        .ClearVals
        .AddItem "<All Departments>", "0"
        '.AddVals con, "Name", "Departments", "DeptID", "Where Deptid<>'VDRS'"
        .Text = "<All Departments>"
    End With
    
End Sub

Private Sub FillList()
    
    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim strCondition As String
    Dim DTStart As Date, DTEnd As Date
    DTStart = Date 'DateAdd("d", -7, Date)
    
    Select Case cmbLeaveWithin.ListIndex
        Case 0
            DTEnd = DateAdd("d", 7, DTStart)
        Case 1
            DTEnd = DateAdd("d", 15, DTStart)
        Case 2
            DTEnd = DateAdd("d", 30, DTStart)
    End Select
    
    strCondition = " Where DT Between '" & DTStart & "' AND '" & DTEnd & "'"
    If cmbDept.ID <> "0" Then
        strCondition = strCondition & " AND DeptID='" & cmbDept.ID & "'"
    End If
    
    Dim strLeaveType As String
    StrSQL = "Select * From VEmpLeaves " & strCondition & " AND EmpID In(Select EmpID From EmpOffice) Order By DT,EmpID"
    With rs
        .Open StrSQL, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            ITM.Tag = !EmpId & ""
            ITM.SubItems(1) = !Name & " " & !Rel & " " & !FName
            ITM.SubItems(2) = !Designation & ""
            ITM.SubItems(3) = Format(!DT, "dd-MMM-yyyy")
                        
            ITM.SubItems(4) = !LeaveType
            ITM.SubItems(5) = !AuthPerson
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub ComputeIT()

'    On Error GoTo ERR
'    If MsgBox("Social Security For The Selected Month Has Not Been Computed." & vbNewLine & "Do You Want To Compute It...?", vbQuestion + vbYesNo) = vbNo Then
'        Exit Sub
'    End If
'
'    Dim rsList As New ADODB.Recordset
'    Dim cmd As New ADODB.command
'    Dim ITM As ListItem
'    Dim Sql As String
'    Dim Shift As String
'    Dim Total As Long
'    With cmd
'
'        .CommandText = "Select * from Employees Where Social=1"
'
'        Set .ActiveConnection = con
'        .Parameters.Refresh
'
'        '.Parameters(0).Type = adVarChar
'        '.Parameters(0).value = cmbMonth
'        '.Parameters(1).Type = adVarChar
'        '.Parameters(1).value = cmbYear
'
'        rsList.CursorType = adOpenForwardOnly
'        rsList.LockType = adLockReadOnly
'        Set rsList = .Execute
'        LV.ListItems.Clear
'
'    End With
'
'    Dim PerSocial As Double
'    Dim MonthDays As Integer
'
'    Dim Sal As Double
'    Dim iMax As Integer, iMin As Integer
'
'    iMax = Val(txtMax)
'    iMin = Val(txtMin)
'    TDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
'    MonthDays = DateAdd("m", 1, TDate) - TDate
'
'    Dim rsPer As New ADODB.Recordset
'
'    With rsPer
'        .Open "Select SocialSecurity from GeneralSettings", con, adOpenForwardOnly, adLockReadOnly
'        If .RecordCount > 0 Then
'            PerSocial = Val(.Fields(0)) / 100
'        End If
'        .Close
'    End With
'
'    Set rsPer = Nothing
'
'    With rsList
'        For i = 0 To .RecordCount - 1
'
'            Set ITM = LV.ListItems.Add(, ![EmpId], ![EmpId])
'
'            ITM.Tag = PerSocial
'            ITM.SubItems(1) = ![Name] & ""
'            ITM.SubItems(2) = ![FName] & ""
'            ITM.SubItems(3) = ![SocialNo] & ""
'            ITM.SubItems(4) = ![Designation] & ""
'            ITM.SubItems(5) = ![Status] & ""
'
'            Randomize
'
'            ITM.SubItems(6) = Int(Rnd * (iMax - (iMin - 1)) + (iMin))
'            Sal = Val(![SocialSal])
'            ITM.SubItems(7) = Sal
'
'            ITM.SubItems(8) = Round((Sal / MonthDays) * Val(ITM.SubItems(6)))
'            ITM.SubItems(9) = Round(Val(ITM.SubItems(8)) * PerSocial)
'            Total = Total + Val(ITM.SubItems(8))
'            .MoveNext
'        Next
'
'        .Close
'        txtTotal = Total
'
'    End With
'
'    Set rsLedger = Nothing
'    Sql = ""
'
'    cmdSave.Visible = True
'    cmdCancel.Visible = True
'    cmdClose.Visible = False
'    cmdRefresh.Visible = False
'    cmdPrint.Visible = False
'
'    Exit Sub
'ERR:
'    MsgBox ERR.Description
End Sub
Private Sub ShowIt()
'    On Error GoTo ERR
'    Dim rsList As New ADODB.Recordset
'    Dim cmd As New ADODB.command
'
'    Dim ITM As ListItem
'
'    With cmd
'        .CommandText = "Select * from VSSMonthly"
'        .ActiveConnection = con
'        .Parameters.Refresh
'        .Parameters(0).Type = adVarChar
'        .Parameters(0).value = cmbMonth
'        .Parameters(1).Type = adVarChar
'        .Parameters(1).value = cmbYear
'        rsList.CursorType = adOpenForwardOnly
'        rsList.LockType = adLockReadOnly
'        Set rsList = .Execute
'    End With
'
'    With rsList
'        Dim MDays As Integer
'
'        LV.ListItems.Clear
'        For i = 0 To .RecordCount - 1
'            Set ITM = LV.ListItems.Add(, ![EmpId], ![EmpId])
'            ITM.ListSubItems(1) = ![Name] & ""
'            ITM.ListSubItems(2) = ![FName] & ""
'            ITM.ListSubItems(3) = ![SocialNo] & ""
'            ITM.ListSubItems(4) = ![Designation] & ""
'            ITM.ListSubItems(5) = ![DaysWorked] & ""
'            ITM.ListSubItems(6) = ![Salary] & ""
'            MDays = DateAdd("m", 1, ![DT]) - ![DT]
'            ITM.ListSubItems(7) = (Val(![Salary]) / MDays) * Val(![DaysWorked])
'            ITM.ListSubItems(8) = Val(ITM.ListSubItems(7)) * Val(![PerDeduct])
'            .MoveNext
'        Next
'        .Close
'    End With
'    Set rsLedger = Nothing
'    Sql = ""
'    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub AddEmps(ITM As ListItem, rs As ADODB.Recordset, Shift As String)
    With rs
        Set ITM = LV.ListItems.add(, ![EmpId] & Shift, ![EmpId] & "", , ICO)
        ITM.SubItems(1) = ![Name]
        ITM.SubItems(2) = ![FName]
        ITM.SubItems(3) = Shift
        ITM.SubItems(4) = 0
    End With
    'Itm.SubItems(4) = Format(![DT], "dddd dd MMMM ,yyyy")
        
End Sub

Private Sub Form_Resize()
  
On Error Resume Next

 fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2

 FLV.Move 50, 1230, Me.ScaleWidth - 100, Me.ScaleHeight - 1400
 
 LV.Move 100, 200, Me.Width - 400, FLV.Height - fButs1.Height - 300
 
 fButs1.Move LV.Left + LV.Width - fButs1.Width, LV.Top + LV.Height
 
 

    
End Sub

Private Sub LV_DblClick()

'    If LV.ListItems.Count = 0 Then Exit Sub
'
'    For i = 1 To LV.ColumnHeaders.Count
'        LV.ColumnHeaders(i).Width = Val(LV.ColumnHeaders(i).Tag)
'    Next
'
'    With PicEdit
'        LV.SelectedItem.EnsureVisible
'        .Move LV.Left + LV.ColumnHeaders(5).Left + 35, LV.Top + LV.SelectedItem.Top + 30, LV.ColumnHeaders(5).Width + LV.ColumnHeaders(6).Width + 30, LV.SelectedItem.Height
'
'        txtedit(0).Move 0, -15, LV.ColumnHeaders(5).Width
'        txtedit(0).Text = LV.SelectedItem.SubItems(4)
'
'        txtedit(1).Move txtedit(0).Width, -15, LV.ColumnHeaders(6).Width
'        txtedit(1).Text = LV.SelectedItem.SubItems(5)
'
'        .Visible = True
'        txtedit(0).SetFocus
'    End With
    
    
End Sub


Private Sub mnuPrintSSForm_Click()
    Screen.MousePointer = vbHourglass
    
    Dim rpt As craxddrt.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptSSForm.rpt")
    Dim f As New frmPrevRpt
    Dim SelFormula As String
    SelFormula = "Month({VSocialSecurity.DT})=" & cmbMonth.List(cmbMonth.ListIndex, 1) & _
    " and Year({VSocialSecurity.DT})=" & Val(cmbYear)
    
    f.ShowReport SelFormula & " ", rpt
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub muuPrintSSSheet_Click()
    
    Screen.MousePointer = vbHourglass
    
    Dim rpt As craxddrt.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptSocialSecurity.rpt")
    Dim f As New frmPrevRpt
    Dim SelFormula As String
    SelFormula = "Month({VSocialSecurity.DT})=" & cmbMonth.List(cmbMonth.ListIndex, 1) & _
    " and Year({VSocialSecurity.DT})=" & Val(cmbYear)
    
    f.ShowReport SelFormula & " ", rpt
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If LCase(Screen.ActiveControl.Name) <> "txtedit" Then PicEdit.Visible = False
End Sub

Private Sub txtMax_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumber(KeyAscii)
End Sub

Private Sub txtMax_Validate(cancel As Boolean)

    Dim iDays As Integer
    Dim DT As Date
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    iDays = DateAdd("m", 1, DT) - DT
    With txtMax
        If Val(.Text) > iDays Then
            MsgBox "Value Can't Be More Than The Days Of Selected Month.", vbInformation
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            cancel = True
        ElseIf Val(.Text) < 2 Then
            MsgBox "Value Can't Be Less Than The 2.", vbInformation
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            cancel = True
        ElseIf Val(.Text) < Val(txtMin) Then
            MsgBox "Value Can't Be Less Than Min Limit.", vbInformation
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            cancel = True
        End If
    End With
    
End Sub

Private Sub txtMin_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumber(KeyAscii)
End Sub

Private Sub txtMin_Validate(cancel As Boolean)

    Dim iDays As Integer
    Dim DT As Date
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    iDays = DateAdd("m", 1, DT) - DT
    
    With txtMin
    
        If Val(.Text) < 1 Then
            MsgBox "Value Can't Be Less Than 1.", vbInformation
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            cancel = True
        ElseIf Val(.Text) > (iDays - 1) Then
            MsgBox "Value Can't Be More Than The Days Of Selected Month.", vbInformation
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            cancel = True
        ElseIf Val(.Text) > Val(txtMax.Text) Then
            MsgBox "Value Can't Be More Than Max Limit.", vbInformation
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            cancel = True
        End If
        
    End With
        
End Sub

Private Sub GetTotal()
    
    Dim total As Double

    For i = 1 To LV.ListItems.count
        total = total + Val(LV.ListItems(i).SubItems(5))
    Next
    txtTotal = Format(total, "0.00")
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    On Error GoTo err
    If Button <> 2 Then Exit Sub
    If LV.ListItems.count = 0 Then Exit Sub
    Dim strRet As String
    'strRet = ShowPopUpMenu(Me.Left + cmdPrint.Left, Me.Top + cmdPrint.Top - cmdPrint.Height, "&Print Material Detail", "-", "Print Material &Rcving Detail", "-", "Print Printing &Costing")
    
    strRet = ShowPopUpMenu1("&Print Report", "-", "Print Leave Balance", "-", "Print Leave Balance of " & LV.SelectedItem.SubItems(1))
    
    Dim rpt As craxddrt.Report, frmRpt As New frmPrevRpt
    
    If strRet = "&Print Report" Then
        Me.MousePointer = vbHourglass
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpsOnLeave.rpt")
'        On Error Resume Next
'        SrcReport.Database.Tables(1).ConnectBufferString = "DSN=DSN_Vision;;User ID=sa;;Database=Vision;;Password=visionasim"
'        rpt.SQLQueryString = StrSQL
        rpt.FormulaFields(1).Text = "'" & cmbLeaveWithin.Text & "'"
        frmRpt.ShowReport "", rpt, StrSQL
        Me.MousePointer = vbNormal
    ElseIf strRet = "Print Leave Balance" Then
        Call PrintEmpLeaveBalance
    ElseIf strRet = "Print Leave Balance of " & LV.SelectedItem.SubItems(1) Then
        Call PrintEmpLeaveBalance(LV.SelectedItem.Tag)
    End If
    
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Private Sub PrintEmpLeaveBalance(Optional strEmpID As String = "")

    On Error GoTo err
    Dim cmd As New ADODB.Command
    Dim strYear As String, strMonth As String
    strYear = Year(Date)
    strMonth = Month(Date)
    With cmd
        .CommandTimeout = 120
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_EmpLeaveBalances"
        .Parameters("@DT").value = DateSerial(strYear, strMonth, 1)
        .Execute
    End With
    
    Dim rpt As craxddrt.Report
    
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpLeaveBalanceMonthly.rpt")
    
    rpt.FormulaFields(16).Text = "'" & strMonth & " - " & strYear & "'"
    Dim f As New frmPrevRpt
    Dim forSel As String
    
    If strEmpID <> "" Then
        forSel = "{VEmpLeaveBalance.EmpID}='" & strEmpID & "' "
    End If
    
    Dim StartDT As Date
    Dim EndDT As Date
    
    If CInt(strMonth) < 7 Then
        StartDT = DateSerial(CInt(strYear) - 1, 7, 1)
    Else
        StartDT = DateSerial(CInt(strYear), 7, 1)
    End If
    EndDT = Date
    
    Dim FormulaFields As craxddrt.FormulaFieldDefinitions
    Dim FormulaField As craxddrt.FormulaFieldDefinition
    
    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & strCompany & "'"
        End If
        If strEmpID <> "" Then
            If FormulaField.Name = "{@ForStartDT}" Then
                FormulaField.Text = "#" & StartDT & "#"
            ElseIf FormulaField.Name = "{@ForEndDT}" Then
                FormulaField.Text = "#" & EndDT & "#"
            End If
        End If
    Next

    f.ShowReport forSel, rpt
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

