VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "combolist.ocx"
Begin VB.Form frmEmpMonthlyPerformanceDeductionAmt 
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   120
   ClientWidth     =   12840
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   8685
   ScaleMode       =   0  'User
   ScaleWidth      =   11500.96
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkPrintInternal 
      Caption         =   "Print Internal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   345
      Left            =   7125
      TabIndex        =   18
      Top             =   225
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   885
   End
   Begin MSComCtl2.DTPicker DTAttend 
      Height          =   285
      Left            =   4695
      TabIndex        =   1
      Top             =   315
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
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
      CustomFormat    =   "MMM-yyyy"
      Format          =   116785155
      CurrentDate     =   37055
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   165
      TabIndex        =   0
      Top             =   300
      Width           =   4485
      _ExtentX        =   7911
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
   Begin VB.Frame FLV 
      Height          =   8055
      Left            =   165
      TabIndex        =   6
      Top             =   570
      Width           =   12435
      Begin VB.TextBox txtValue 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1605
         TabIndex        =   2
         Top             =   7110
         Width           =   1350
      End
      Begin VB.TextBox txtTotal 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10980
         TabIndex        =   21
         Top             =   7110
         Width           =   1350
      End
      Begin VB.TextBox txtHours 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4125
         TabIndex        =   15
         Top             =   2265
         Visible         =   0   'False
         Width           =   1290
      End
      Begin MSComCtl2.DTPicker TTo 
         Height          =   315
         Left            =   5730
         TabIndex        =   13
         Top             =   1110
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
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
         CustomFormat    =   "HH:mm"
         Format          =   116785155
         UpDown          =   -1  'True
         CurrentDate     =   37683.8333333333
      End
      Begin MSComCtl2.DTPicker TFrom 
         Height          =   315
         Left            =   4020
         TabIndex        =   14
         Top             =   1110
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
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
         CustomFormat    =   "HH:mm"
         Format          =   116785155
         UpDown          =   -1  'True
         CurrentDate     =   37958.7083333333
      End
      Begin VB.Frame FButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   570
         Left            =   2820
         TabIndex        =   10
         Top             =   7455
         Width           =   9525
         Begin MSForms.CommandButton cmdEmpWorkDone 
            CausesValidation=   0   'False
            Height          =   360
            Left            =   30
            TabIndex        =   19
            Top             =   135
            Visible         =   0   'False
            Width           =   1575
            ForeColor       =   0
            Caption         =   "Add Work Done"
            PicturePosition =   327683
            Size            =   "2778;635"
            Accelerator     =   68
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdPrintEmpWorkDone 
            CausesValidation=   0   'False
            Height          =   360
            Left            =   1650
            TabIndex        =   17
            Top             =   135
            Visible         =   0   'False
            Width           =   2460
            ForeColor       =   0
            Caption         =   "Print Employee Work Done"
            PicturePosition =   327683
            Size            =   "4339;635"
            Accelerator     =   87
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdPrintDaily 
            CausesValidation=   0   'False
            Height          =   360
            Left            =   5505
            TabIndex        =   16
            Top             =   135
            Width           =   1305
            ForeColor       =   0
            Caption         =   "Print"
            PicturePosition =   327683
            Size            =   "2302;635"
            Accelerator     =   68
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdPrintAll 
            CausesValidation=   0   'False
            Height          =   360
            Left            =   4155
            TabIndex        =   11
            Top             =   135
            Visible         =   0   'False
            Width           =   1305
            ForeColor       =   0
            Caption         =   "Print All"
            PicturePosition =   327683
            Size            =   "2302;635"
            Accelerator     =   65
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdClose 
            CausesValidation=   0   'False
            Height          =   360
            Left            =   8205
            TabIndex        =   5
            Top             =   135
            Width           =   1305
            ForeColor       =   0
            Caption         =   " Close"
            PicturePosition =   327683
            Size            =   "2302;635"
            Accelerator     =   67
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdSave 
            Height          =   360
            Left            =   6855
            TabIndex        =   4
            Top             =   135
            Width           =   1305
            ForeColor       =   0
            Caption         =   " Save "
            PicturePosition =   327683
            Size            =   "2302;635"
            Accelerator     =   83
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.Frame FButs2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   570
         Left            =   105
         TabIndex        =   7
         Top             =   7455
         Width           =   3000
         Begin MSForms.CommandButton cmdClearAll 
            Height          =   390
            Left            =   1470
            TabIndex        =   9
            Top             =   120
            Visible         =   0   'False
            Width           =   1485
            ForeColor       =   0
            Caption         =   "UnAuthorize All"
            PicturePosition =   327683
            Size            =   "2619;688"
            Accelerator     =   108
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdCheckAll 
            Height          =   360
            Left            =   0
            TabIndex        =   8
            Top             =   135
            Visible         =   0   'False
            Width           =   1425
            ForeColor       =   0
            Caption         =   "Authorize All"
            PicturePosition =   327683
            Size            =   "2514;635"
            Accelerator     =   104
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
         Height          =   6915
         Left            =   120
         TabIndex        =   3
         Top             =   150
         Width           =   12225
         _ExtentX        =   21564
         _ExtentY        =   12197
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "EmpID"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Employee Description"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Rate"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Search Emp. ID :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   22
         Top             =   7140
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total Amount :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   9675
         TabIndex        =   20
         Top             =   7140
         Width           =   1245
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7425
      Top             =   3120
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
            Picture         =   "frmEmpMonthlyPerformanceDeductionAmt.frx":0000
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpMonthlyPerformanceDeductionAmt.frx":0452
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Department                                                                            Date"
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
      Index           =   0
      Left            =   165
      TabIndex        =   12
      Top             =   60
      Width           =   6090
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete This Attendance"
      End
   End
End
Attribute VB_Name = "frmEmpMonthlyPerformanceDeductionAmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Editing As Boolean
Dim forLeave As Boolean

Public Sub enterLateHours()
    Load LateHours
    LateHours.Tag = LV.SelectedItem
    LateHours.Show
End Sub

Private Sub cmbEmp_matched()

On Error Resume Next
If cmbEmp.ID = "0" Then Exit Sub
LV.ListItems(cmbEmp.ID).Selected = True
LV.ListItems(cmbEmp.ID).EnsureVisible

End Sub

Private Sub cmbShift_Change()
    Call FillList
End Sub

Private Sub cmdCheckAll_Click()
    
    
        For i = 1 To LV.ListItems.count
            With LV.ListItems(i)
                .Checked = True
                .SubItems(2) = "OverTime"
            End With
        
        Next i
    
End Sub

Private Sub cmdClearAll_Click()

    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            .Checked = False
            .SubItems(2) = ""
        End With
    Next i

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub


Public Sub cmdRefresh_Click()
    'Call RefreshLedger
    
    'Call Form_Resize
End Sub


Private Sub cmdPrint_Click()

    If LV.ListItems.count = 0 Then Exit Sub

    con.Execute "DELETE FROM PrintOverTime"

    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                con.Execute "INSERT INTO PrintOverTime VALUES('" & .Text _
                 & "','" & .ListSubItems(1) & "','" & .SubItems(2) & "','" & .SubItems(3) & "')"
            End If
        End With
    Next


    Screen.MousePointer = vbHourglass

    Dim f As New frmPrevRpt, rpt As craxddrt.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptPrintOverTime.rpt")
    
    Dim FormulaFields As craxddrt.FormulaFieldDefinitions
    Dim FormulaField As craxddrt.FormulaFieldDefinition

    Set FormulaFields = rpt.FormulaFields

    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & company & "'"
        ElseIf FormulaField.Name = "{@FromTo}" Then
            FormulaField.Text = "'" & Format(DTAttend, "dddd dd-mmm-yyyy") & "'"
        End If
    Next

'    On Error Resume Next
'    SrcReport.Database.Tables(1).ConnectBufferString = strConnectionStringrpt
'    rpt.SQLQueryString = SQL
    f.ShowReport Cond & " ", rpt
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdEmpWorkDone_Click()
    frmEmpWorkDone.Show 1
End Sub

Private Sub cmdPrintAll_Click()

    On Error GoTo err
    
    Dim rpt As craxddrt.Report
    If chkPrintInternal.Value = vbChecked Then
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptOverTimeInternal.rpt")
    Else
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptOverTime1.rpt")
    End If
    Dim f As New frmPrevRpt
    f.ShowReport " Month({OverTimeAuth.DT})=" & Month(DTAttend) & " AND Year({OverTimeAuth.DT})=" & year(DTAttend) & " ", rpt
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdPrintDaily_Click()

    On Error GoTo err
    
    Dim rpt As craxddrt.Report
'    If chkPrintInternal.value = vbChecked Then
'        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptOverTimeInternal.rpt")
'    Else
'        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptOverTime1.rpt")
'    End If
'
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\MonthlyPerformanceDeductionAmounts.rpt")
    rpt.FormulaFields.GetItemByName("Heading").Text = "'" & Format(DTAttend, "MMM,yyyy") & "'"
    Dim f As New frmPrevRpt
    f.ShowReport " Month({EmpMonthlyPerformanceDeductionAmount.DT})=" & Month(DTAttend) & " AND Year({EmpMonthlyPerformanceDeductionAmount.DT})=" & year(DTAttend), rpt
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdPrintEmpWorkDone_Click()

    On Error GoTo err
    
    Dim rpt As craxddrt.Report
    
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptOverTimeWork.rpt")
    
    Dim f As New frmPrevRpt
    f.ShowReport " Month({OverTimeAuth.DT})=" & Month(DTAttend) & " AND Year({OverTimeAuth.DT})=" & year(DTAttend) & " ", rpt
    
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub cmdSave_Click()

    On Error GoTo err
    
'    If DateLocked(DTAttend.Value, True) Then
'        MsgBox "Date is Locked.", vbInformation
'        Exit Sub
'    End If
    
    Dim SQL As String
    Dim lHrs As Long
     
    Call StartTrans(con)
     
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
                If .Checked = True Then
                    Dim iRowAffect As Integer
                    con.Execute "UPDATE EmpMonthlyPerformanceDeductionAmount SET Amount=" & Val(.SubItems(3)) & " WHERE EntryID=" & Val(.Tag), iRowAffect
                    If iRowAffect = 0 Then
                        con.Execute "INSERT INTO EmpMonthlyPerformanceDeductionAmount (EmpID,DT,Amount) VALUES('" & _
                         .Text & "','" & DTAttend.Tag & "'," & Val(.SubItems(3)) & ")"
                    End If
                Else
                    con.Execute "DELETE FROM EmpMonthlyPerformanceDeductionAmount WHERE EntryID=" & Val(.Tag)
                End If
           
        End With
    Next i
    con.CommitTrans
    Call FillList
    
    Exit Sub

err:
    MsgBox err.Description, vbInformation

End Sub




Private Sub DTAttend_Change()
            
    DTAttend.Tag = DateSerial(DTAttend.year, DTAttend.Month + 1, 0)
    'Call cmdSave_Click
    Call FillList
    
End Sub

Private Sub Form_Load()

    DTAttend = Format(GetServerDate, "dd-MMM-yyyy")
    Call DTAttend_Change
    cmbDept.ListHeight = 6000
    
    
    cmbDept.ClearVals
    'cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID", " Where TempDept=0"
    cmbDept.ID = 0

'    DTAttend.Enabled = UserHasAccess("AttDateChange")
    
End Sub

Private Sub cmbDept_matched()

    Call FillList
    
'    cmbEmp.ClearVals
''    cmbEmp.AddItem "<All Employees>", "0"
'    If cmbDept.Text = "<All Departments>" Then
'        cmbEmp.AddVals con, "Name + ' {' + EmpID + '}'", "Employees", "EmpID", "Where active=1"
'    Else
'        cmbEmp.AddVals con, "Name + ' {' + EmpID + '}'", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "' and active=1 "
'    End If
'    cmbEmp.ID = "0"
    
End Sub

Private Sub FillList()

    If cmbDept.MatchFound = False Then Exit Sub
    Dim iWorkingHrs As Integer

    iWorkingHrs = TotalMonthWorkingHrs(DTAttend.Month, DTAttend.year)
    
    LV.ListItems.Clear
    
    Dim ITM As ListItem, Attend As String
    
    Dim iAdd As Integer
    Dim sSecOut As String
    Dim iTemp As Integer
    Dim rsList As New ADODB.Recordset
    Dim iHolidayCount As Integer
    iHolidayCount = GetSingleLongValue("COUNT(DT)", "Holidays", " WHERE DT='" & DTAttend.Value & "'")
    With rsList
'        If iHolidayCount > 0 Or Format(DTAttend.Value, "ddd") = "Sun" Then
'            .Open "SELECT EmpID,DeptID,Name,FName,Rel,StartingSalary From Employees WHERE DeptID='" & cmbDept.ID & "'", con, adOpenForwardOnly, adLockReadOnly
'        Else
            .Open "SELECT EmpID,DeptID,Name,FName,Rel,StartingSalary From Employees WHERE DeptID='" & cmbDept.ID & "'", con, adOpenForwardOnly, adLockReadOnly
'        End If
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, !EmpId, !EmpId & "")
            ITM.ListSubItems.add , , !Name & " " & !Rel & " " & !FName
            ITM.ListSubItems.add , , Round((Val(!StartingSalary & "") / 26 / 8), 4)
            ITM.ListSubItems.add().ForeColor = vbBlue
            ITM.ListSubItems.add().ForeColor = vbRed
            
             .MoveNEXT
        Loop
        .Close
    
        '    Exit Sub
    
        '''''''''' Will Check For Over Time'''''''''''''''
        txtWorkToDo = ""
        .Open "SELECT * FROM EmpMonthlyPerformanceDeductionAmount WHERE DT='" & DTAttend.Tag & "' ORDER BY EmpID", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            'txtWorkToDo = !WorkToBeDone & ""
            On Error Resume Next
            With LV.ListItems(!EmpId & "")
                If err.Number = 35601 Then GoTo MoveNEXT
                .Tag = rsList!EntryID
                .Checked = True
                .SubItems(3) = Val(rsList!Amount & "")
            End With
MoveNEXT:
            .MoveNEXT
        Loop
        .Close
        
    End With
    Set rsList = Nothing
    Call GetTotalOTAmt
    
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
    
    Exit Sub

    On Error Resume Next

    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2

    FLV.Move 50, 1230, Me.ScaleWidth - 100, Me.ScaleHeight - 1400

    LV.Move 100, 200, Me.Width - 400, FLV.Height - FButs1.Height - 300

    FButs1.Move LV.Left + LV.Width - FButs1.Width, LV.Top + LV.Height

    FButs2.Move LV.Left, LV.Top + LV.Height
    
End Sub


Private Sub LV_DblClick()
    
    If LV.SelectedItem.Checked = False Then Exit Sub
    If LV.SelectedItem.SubItems(2) = "" Then Exit Sub
    
'    If DateLocked(DTAttend.Value, False) Then
'        MsgBox "Date is Locked.", vbInformation
'        Exit Sub
'    End If
    
    TFrom.Move LV.Left + LV.ColumnHeaders(3).Left + 50, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(3).Width
    TTo.Move LV.Left + LV.ColumnHeaders(4).Left + 50, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(4).Width
    txtHours.Move LV.Left + LV.ColumnHeaders(4).Left + 50, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(4).Width
        
    With LV.SelectedItem
        If .SubItems(2) <> "" Then
            TFrom = CDate(.SubItems(2))
        End If
    
        If .SubItems(3) <> "" Then
'            TTo = CDate(.SubItems(3))
'            TTo.Visible = True
            txtHours.Visible = True
        Else
            txtHours.Visible = True
        End If
    End With
    
    txtHours = ""
    txtHours.SetFocus
    
End Sub

Private Sub LV_GotFocus()
    TFrom.Visible = False
    TTo.Visible = False
End Sub

Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)
        
'    If DateLocked(DTAttend.Value, True) Then
'        MsgBox "Date is Locked.", vbInformation
'        Exit Sub
'    End If
    
    Dim clr As Long
    Item.Selected = True
    If Item.Checked Then
        Call LV_DblClick
    Else
        Item.SubItems(2) = ""
        Item.SubItems(3) = ""
    End If
    
   
End Sub


Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
If Button = vbRightButton And LV.ListItems.count > 0 Then
    PopUpMenu mnuPopup
    Exit Sub
End If
    'Set Flag that mouse is in attendance column or not
    forLeave = CBool(x >= LV.ColumnHeaders(3).Left And x <= (LV.ColumnHeaders(3).Left + LV.ColumnHeaders(3).Width))
    

End Sub

Private Sub mnuDelete_Click()

con.Execute "DELETE FROM AttendanceSheet WHERE DT='" & DTAttend.Value & "' AND EmpID='" & LV.SelectedItem.Text & "' AND Shift='" & cmbTime.Tag & "'"
Call FillList

End Sub

Private Sub TFrom_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        If TTo.Visible Then TTo.SetFocus Else txtHours.SetFocus
    ElseIf KeyCode = 27 Then
        TFrom.Visible = False
        TTo.Visible = False
        txtHours.Visible = False
    End If
    
End Sub

Private Sub TTo_KeyDown(KeyCode As Integer, Shift As Integer)


Dim LHours As Integer, LMinute As Integer
Dim OHours As Integer, OMinute As Integer, Fine As Double, ExtraHrs As Double

Dim DefinedWorkingMins As Long, EmpWorkingMins As Long
Dim MinDiff As Long

If KeyCode = 13 Then
   
    
    With LV.SelectedItem

        .Checked = True
        
        .SubItems(2) = Format(TFrom.Value, "HH:mm")
        .SubItems(3) = Format(TTo.Value, "HH:mm")
        
        TTo.Visible = False
        TFrom.Visible = False
    
        LV.SetFocus
    End With
End If

End Sub


Private Sub txtvalue_Change()
    

    Dim StartIdx As Integer
    
    
    StartIdx = 1
    
    Dim fldValue As String

    For i = StartIdx To LV.ListItems.count
       
       
     
        fldValue = LV.ListItems(i)
    
        If LCase(fldValue) Like "*" & LCase(Trim(txtValue)) & "*" Then
            LV.ListItems(i).Selected = True
            LV.ListItems(i).EnsureVisible
            If i <> StartIdx Then Exit For
        End If
       
    Next

    If i = LV.ListItems.count + 1 Then
        MsgBox "Item Not Found.", vbInformation
        Set LV.SelectedItem = Nothing
    End If
    
    Exit Sub

End Sub

Private Sub txtEmpID_KeyPress(KeyAscii As Integer)
    
End Sub

Private Sub txtHours_KeyPress(KeyAscii As Integer)

    Dim LHours As Integer, LMinute As Integer, dHrs As Double
    Dim OHours As Integer, OMinute As Integer, Fine As Double, ExtraHrs As Double

    Dim DefinedWorkingMins As Long, EmpWorkingMins As Long
    Dim MinDiff As Long
    
    Dim sTto As String
    Dim iAdd As Integer
    Dim iTemp As Integer
    
    If KeyAscii = 13 Then
        If Val(txtHours) < 0 Then
            MsgBox "Invalid Entry.", vbInformation
            Exit Sub
        End If
        
        With LV.SelectedItem
            
            .Checked = True
        
            .SubItems(3) = Val(txtHours)
            
            txtHours.Visible = False
            
            LV.SetFocus
            GetTotalOTAmt
        End With
        
    ElseIf KeyAscii = 27 Then
        'TTo.Visible = False
        'TFrom.Visible = False
        txtHours.Visible = False
    End If

End Sub

Private Sub GetTotalOTAmt()

    Dim i As Integer
    Dim DTotal As Double
    For i = 1 To LV.ListItems.count
        DTotal = DTotal + Val(LV.ListItems(i).SubItems(3))
    Next
    txtTotal = DTotal
    
End Sub

Private Sub txtValue_GotFocus()
    With txtValue
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
