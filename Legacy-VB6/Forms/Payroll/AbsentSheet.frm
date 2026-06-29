VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "combolist.ocx"
Begin VB.Form AbsentSheet 
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11985
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   6915
   ScaleMode       =   0  'User
   ScaleWidth      =   10735.13
   WindowState     =   2  'Maximized
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   315
      Left            =   6825
      TabIndex        =   13
      Top             =   420
      Width           =   2760
      _ExtentX        =   4868
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
   Begin VB.ComboBox cmbRange 
      Height          =   315
      ItemData        =   "AbsentSheet.frx":0000
      Left            =   105
      List            =   "AbsentSheet.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   420
      Width           =   2175
   End
   Begin VB.Frame FLV 
      Height          =   6045
      Left            =   15
      TabIndex        =   0
      Top             =   795
      Width           =   11910
      Begin VB.Frame fButs 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   855
         Left            =   4620
         TabIndex        =   7
         Top             =   5160
         Width           =   7215
         Begin MSForms.CommandButton cmdPrint 
            Height          =   360
            Left            =   4245
            TabIndex        =   11
            Top             =   465
            Width           =   1425
            ForeColor       =   0
            Caption         =   "Print         "
            PicturePosition =   327683
            Size            =   "2514;635"
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
            Height          =   360
            Left            =   5730
            TabIndex        =   10
            Top             =   465
            Width           =   1425
            ForeColor       =   0
            Caption         =   " Close      "
            PicturePosition =   327683
            Size            =   "2514;635"
            Picture         =   "AbsentSheet.frx":0034
            Accelerator     =   67
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Total Absents :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   3690
            TabIndex        =   9
            Top             =   105
            Width           =   2370
         End
         Begin VB.Label lblTotalFirstAbs 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0 "
            Height          =   285
            Left            =   6120
            TabIndex        =   8
            Top             =   75
            Width           =   975
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   4980
         Left            =   105
         TabIndex        =   1
         Top             =   180
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   8784
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   4194304
         BackColor       =   16777215
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Emp ID"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Emp Name"
            Object.Width           =   12347
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Day"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date"
            Object.Width           =   2822
         EndProperty
      End
   End
   Begin MSComCtl2.DTPicker DtTo 
      Height          =   315
      Left            =   4575
      TabIndex        =   2
      Top             =   420
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dddd dd-MMM-yyyy"
      Format          =   114753539
      CurrentDate     =   37081
   End
   Begin MSComCtl2.DTPicker DtFrom 
      Height          =   315
      Left            =   2295
      TabIndex        =   3
      Top             =   420
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dddd dd-MMM-yyyy"
      Format          =   114753539
      CurrentDate     =   37055
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9870
      Top             =   825
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AbsentSheet.frx":0146
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AbsentSheet.frx":0598
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkShowEmps 
      Caption         =   "Show Employees"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   10425
      TabIndex        =   6
      Top             =   135
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"AbsentSheet.frx":06AA
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   105
      TabIndex        =   12
      Top             =   165
      Width           =   9450
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   10380
      TabIndex        =   4
      Top             =   345
      Width           =   1500
      Caption         =   "Refresh        "
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
End
Attribute VB_Name = "AbsentSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                                     
Private Sub chkShowEmps_Click()
    Call ShowWait
    cmdRefresh_Click
    Call HideWait
End Sub

Private Sub cmbRange_Click()

    DtTo = GetServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    Select Case cmbRange.ListIndex
        Case 0
            DtFrom = GetServerDate
        Case 1
            DtFrom = DateSerial(year(GetServerDate), Month(GetServerDate), 1)
            DtTo = DateSerial(year(GetServerDate), Month(GetServerDate), DateDiff("d", GetServerDate, DateAdd("M", 1, GetServerDate)))
        Case 2
            DtFrom.Enabled = True
            DtTo.Enabled = True
            DtFrom.SetFocus
            Exit Sub
    End Select
    Call cmdRefresh_Click
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrintOld_Click()

    On Error GoTo err
    Call ShowWait
    Dim Cond As String
    
    If cmbShift.ListIndex = 0 Then
        'Cond = " {VAttendance.Shift} in('First','Second') "
        Cond = " {VAttendance.Shift}='First' Or {VAttendance.Shift}='Second'"
    Else
        Cond = " {VAttendance.Shift}='" & cmbShift & "' "
    End If
    
    If cmbEmp.ID = "0" Then
        If cmbDept.ID = "0" Then
            Cond = Cond & " AND ({VAttendance.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & "))"
        Else
            Cond = Cond & " AND ({VAttendance.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & ")) and {VAttendance.DeptID}='" & cmbDept.ID & "'"
        End If
    Else
        Cond = Cond & " and {VAttendance.empid}='" & cmbEmp.ID & "' AND ({VAttendance.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & "))"
    End If
    
With cr1
    .ReportFileName = RptPath & "\AttSheet.rpt"
    '.DataFiles(0) = rptpath & "\Payroll.mdb"
    .Connect = ConnectStr
    .SelectionFormula = Cond
    
    .Formulas(0) = "FromTo='From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
    .Formulas(1) = "Company='" & strCompany & "'"
    
    .Action = 1
    .PageZoomNext
End With

Call HideWait
Exit Sub
err:
    MsgBox err.Description
    Call HideWait
End Sub

Private Sub cmdPrint1_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    
    con.Execute "Delete from PrintAbsSheet"
    
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            con.Execute "Insert into PrintAbsSheet(EmpID,EmpName,FirstShiftAbsents,Bold)" & _
             " values ('" & .Text & "','" & .SubItems(1) & "'," & Val(.SubItems(2)) & "," & Abs(.Bold) & ")"
        End With
    Next
Dim FromTo As String


'    With cr1
'        .ReportFileName = RptPath & "\AbsentSheet.rpt"

         If cmbRange.ListIndex = 0 Then
            FromTo = "'For " & Format(DtFrom, "dd-MMM-yyyy") & "'"
         ElseIf cmbRange.ListIndex = 1 Then
            FromTo = "For The Month of '" & Format(DtFrom, "MMM-yyyy") & "'"
         Else
            FromTo = "'From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
         End If
'         .Formulas(1) = "Company='" & strCompany & "'"
         
'        .Action = 1
'        .PageZoomNext
'    End With


Screen.MousePointer = vbHourglass

Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptAbsentSheet.rpt")

Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

Set FormulaFields = rpt.FormulaFields
For Each FormulaField In FormulaFields
    If FormulaField.Name = "{@Company}" Then
        FormulaField.Text = "'" & strCompany & "'"
    ElseIf FormulaField.Name = "{@FromTo}" Then
        FormulaField.Text = FromTo
    End If
Next

f.ShowReport Cond & " ", rpt

Screen.MousePointer = vbDefault


End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
    Me.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report, frm As New frmPrevRpt
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpAbsentSheetSummary.rpt")
    rpt.FormulaFields(1).Text = "'" & Format(DtFrom.Value, "dd-MMM-yyyy") & " to " & Format(DtTo.Value, "dd-MMM-yyyy") & "'"
    frm.ShowReport "", rpt
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLedger
End Sub



Private Sub Form_Load()

    cmbRange.ListIndex = 0
    cmbDept.ListHeight = 4000
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "'{' + DeptID + '} ' + Name", "Departments", "DeptID"
    cmbDept.ID = "0"
    
End Sub
Private Sub cmbDept_matched()

    'cmbEmp.ClearVals
    'cmbEmp.AddItem "<All Employees>", "0"
    'If cmbDept.Text = "<All Departments>" Then
    '    cmbEmp.AddVals con, "Name", "Employees", "EmpID"
    'Else
    '    cmbEmp.AddVals con, "Name", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "'"
    'End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
 
    FLV.Move 50, 800, Me.ScaleWidth - 100, Me.ScaleHeight - 900
 
    LV.Move 100, 200, LV.Width, FLV.Height - fButs.Height - 300
 
    fButs.Move LV.Left + LV.Width - fButs.Width, LV.Top + LV.Height


 
End Sub

Private Sub RefreshLedger()
    
    On Error GoTo err
    Me.MousePointer = vbHourglass
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    With cmd
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_PrintEmpAbsentList"
        .Parameters("@DTFrom").Value = DtFrom.Value
        .Parameters("@DTTo").Value = DtTo.Value
        If cmbDept.ID <> "0" Then
            .Parameters("@DeptID").Value = cmbDept.ID
        End If
        .Execute
    End With
    Dim ITM As ListItem
    With rs
        .Open "Select T1.EmpID,T1.DT,T2.Name,T2.Rel,T2.FName From PrintEmpAbsentList T1 Inner Join Employees T2 ON T1.EmpID=T2.EmpID Order By T1.DT,T1.EmpID", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        lblTotalFirstAbs.Caption = .RecordCount
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , !EmpId & "")
            ITM.SubItems(1) = !Name & " " & !Rel & " " & !FName
            ITM.SubItems(2) = Format(!DT, "dddd")
            ITM.SubItems(3) = Format(!DT, "dd-MMM-yyyy")
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub RefreshLedgerOLD()
    
    Dim rsLedger As New ADODB.Recordset
    Dim rsShifts As New ADODB.Recordset
    Dim ITM As ListItem
    Dim Cond As String
    Dim Sql As String
    Dim ICO As String
    Dim Dates As String
    Dim FAbsTotal As Double, SAbsTotal As Double
    
    Cond = " and DT between '" & DtFrom & "' and '" & DtTo & "' "
    
    Sql = "Select DeptID,Name  DeptName," & _
    "(Select Count(Attendance) From AttendanceSheet Where Attendance=0 and Empid like dept.DeptID+ '%' " & Cond & ") SecondShiftAbsents " & _
    " From Departments Dept Order By DeptID"
    
    rsLedger.CursorLocation = adUseClient
    rsLedger.Open Sql, con, adOpenForwardOnly, adLockReadOnly
    Dim pos As Integer
    Dim LateHours As String
    Dim DT As String
    LV.ListItems.Clear
    With rsShifts
        Do Until rsLedger.EOF
            
            Set ITM = LV.ListItems.add(, , rsLedger![DeptID] & "")
            ITM.Bold = True
            
            ITM.ListSubItems.add(, , rsLedger![DeptName] & "").Bold = True
            ITM.ListSubItems.add(, , rsLedger![SecondShiftAbsents] & "").Bold = True
            SAbsTotal = SAbsTotal + Val(rsLedger![SecondShiftAbsents] & "")
            
            'Show Summary/Detail Report As Selected
            If chkShowEmps <> vbChecked Then GoTo NEXT_DEPT
            
            .Open "Select EmpID,Name,Rel,FName," & _
            "(Select Count(Attendance) From AttendanceSheet Where Attendance=0 and Empid like Emp.empid " & Cond & ") SecondShiftAbsents " & _
            " From Employees Emp Where Empid in " & _
            " (Select Distinct(EmpID) From AttendanceSheet Where Attendance=0 " & Cond & " and EmpID Like '" & rsLedger![DeptID] & "%') Order By EmpID", con, adOpenForwardOnly, adLockReadOnly
            
            Do Until .EOF
                Set ITM = LV.ListItems.add(, , ![EmpId] & "")
                ITM.SubItems(1) = ![Name] & IIf(IsNull(![FName]), "", " " & ![Rel] & " " & ![FName])
                ITM.SubItems(2) = ![SecondShiftAbsents] & ""
                .MoveNEXT
            Loop
            .Close
            
NEXT_DEPT:
            
            rsLedger.MoveNEXT
        Loop
    End With
    
    Set rsLedger = Nothing
    Sql = ""
    Cond = ""
    ICO = ""
    lblTotalFirstAbs = FAbsTotal & " "
    lblTotalSecondAbs = SAbsTotal & " "
End Sub




