VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form OverTimeAuthOld1 
   BorderStyle     =   0  'None
   ClientHeight    =   6915
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12750
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   ScaleHeight     =   6915
   ScaleMode       =   0  'User
   ScaleWidth      =   11420.35
   ShowInTaskbar   =   0   'False
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
      Left            =   10470
      TabIndex        =   21
      Top             =   900
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   885
   End
   Begin MSComCtl2.DTPicker DTAttend 
      Height          =   285
      Left            =   7170
      TabIndex        =   13
      Top             =   975
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   503
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
      Format          =   65208323
      CurrentDate     =   37055
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   975
      Width           =   3435
      _ExtentX        =   6059
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
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   3675
      TabIndex        =   0
      Top             =   975
      Width           =   3480
      _ExtentX        =   6138
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
   Begin VB.Frame FLV 
      Height          =   5550
      Left            =   240
      TabIndex        =   5
      Top             =   1290
      Width           =   12435
      Begin VB.TextBox txtHours 
         Height          =   300
         Left            =   4125
         TabIndex        =   18
         Top             =   2265
         Visible         =   0   'False
         Width           =   1290
      End
      Begin MSComCtl2.DTPicker TTo 
         Height          =   315
         Left            =   5730
         TabIndex        =   16
         Top             =   1110
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   65208323
         UpDown          =   -1  'True
         CurrentDate     =   37683.8333333333
      End
      Begin MSComCtl2.DTPicker TFrom 
         Height          =   315
         Left            =   4020
         TabIndex        =   17
         Top             =   1110
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   65208323
         UpDown          =   -1  'True
         CurrentDate     =   37958.7083333333
      End
      Begin VB.Frame FButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   570
         Left            =   2820
         TabIndex        =   9
         Top             =   4950
         Width           =   9525
         Begin MSForms.CommandButton cmdEmpWorkDone 
            CausesValidation=   0   'False
            Height          =   360
            Left            =   30
            TabIndex        =   22
            Top             =   135
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
            TabIndex        =   20
            Top             =   135
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
            Left            =   4155
            TabIndex        =   19
            Top             =   135
            Width           =   1305
            ForeColor       =   0
            Caption         =   "Print Daily  "
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
            Left            =   5505
            TabIndex        =   14
            Top             =   135
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
            TabIndex        =   11
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
            TabIndex        =   10
            Top             =   135
            Width           =   1305
            ForeColor       =   0
            Caption         =   " Save "
            PicturePosition =   327683
            Size            =   "2302;635"
            Picture         =   "OverTimeAuthOld1.frx":0000
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
         TabIndex        =   6
         Top             =   4950
         Width           =   3000
         Begin MSForms.CommandButton cmdClearAll 
            Height          =   390
            Left            =   1470
            TabIndex        =   8
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
            TabIndex        =   7
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
         Height          =   4740
         Left            =   120
         TabIndex        =   12
         Top             =   150
         Width           =   12225
         _ExtentX        =   21564
         _ExtentY        =   8361
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "EmpID"
            Object.Width           =   3969
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Employee Description"
            Object.Width           =   11483
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "From"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "To"
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   255
      TabIndex        =   2
      Top             =   120
      Width           =   9915
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Authorize Over Time"
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
         Left            =   3615
         TabIndex        =   3
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Authorize Over Time"
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
         Left            =   3630
         TabIndex        =   4
         Top             =   135
         Width           =   3255
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7500
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
            Picture         =   "OverTimeAuthOld1.frx":0112
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OverTimeAuthOld1.frx":0564
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Department :                                                  Employee Name :                                             Date :"
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
      Left            =   240
      TabIndex        =   15
      Top             =   720
      Width           =   9270
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete This Attendance"
      End
   End
End
Attribute VB_Name = "OverTimeAuthOld1"
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
    
    
        For i = 1 To LV.ListItems.Count
            With LV.ListItems(i)
                .Checked = True
                .SubItems(2) = "OverTime"
            End With
        
        Next i
    
End Sub

Private Sub cmdClearAll_Click()

    For i = 1 To LV.ListItems.Count
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

    If LV.ListItems.Count = 0 Then Exit Sub

    con.Execute "DELETE FROM PrintOverTime"

    For i = 1 To LV.ListItems.Count
        With LV.ListItems(i)
            If .Checked Then
                con.Execute "INSERT INTO PrintOverTime VALUES('" & .Text _
                 & "','" & .ListSubItems(1) & "','" & .SubItems(2) & "','" & .SubItems(3) & "')"
            End If
        End With
    Next


    Screen.MousePointer = vbHourglass

    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptPrintOverTime.rpt")
    
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

    Set FormulaFields = rpt.FormulaFields

    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & company & "'"
        ElseIf FormulaField.Name = "{@FromTo}" Then
            FormulaField.Text = "'" & Format(DTAttend, "dddd dd-mmm-yyyy") & "'"
        End If
    Next

'    On Error Resume Next
'    SrcReport.Database.Tables(1).ConnectBufferString = "DSN=DSN_Vision;;User ID=sa;;Database=Vision;;Password=visionasim"
'    rpt.SQLQueryString = SQL
    f.ShowReport Cond & " ", rpt
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdEmpWorkDone_Click()
    frmEmpWorkDone.Show 1
End Sub

Private Sub cmdPrintAll_Click()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    If chkPrintInternal.value = vbChecked Then
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
    
    Dim rpt As CRAXDDRT.Report
    If chkPrintInternal.value = vbChecked Then
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptOverTimeInternal.rpt")
    Else
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptOverTime1.rpt")
    End If
    
    Dim f As New frmPrevRpt
    f.ShowReport " Month({OverTimeAuth.DT})=" & Month(DTAttend) & " AND Year({OverTimeAuth.DT})=" & year(DTAttend) & " AND Day({OverTimeAuth.DT})=" & Day(DTAttend) & " ", rpt
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdPrintEmpWorkDone_Click()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptOverTimeWork.rpt")
    
    Dim f As New frmPrevRpt
    f.ShowReport " Month({OverTimeAuth.DT})=" & Month(DTAttend) & " AND Year({OverTimeAuth.DT})=" & year(DTAttend) & " ", rpt
    
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub cmdSave_Click()

On Error GoTo err

    Dim SQL As String
    Dim lHrs As Long
    Dim DeptCond As String
    
    Call StartTrans(con)
    
    DeptCond = IIf(cmbDept.ID = "0", "", " and EmpID Like '" & cmbDept.ID & "%'")
     
    con.Execute "DELETE FROM OverTimeAuth WHERE  DT='" & DTAttend & "'" & DeptCond
    con.Execute "DELETE FROM EmpTimes Where EntryID In(Select OverTimeEntryID From VEmpDailyEntries Where DT='" & _
     DTAttend & "' " & DeptCond & ")"
    
    con.Execute "Update AttendanceSheet Set ExtraHrs=null Where DT='" & DTAttend & "'" & DeptCond
    
    'Dim lOT As Long
    Dim LHrs1 As Double
    Dim lTemp As Long
    For i = 1 To LV.ListItems.Count
        With LV.ListItems(i)
            If .Checked = True Then
               
                    lHrs = DateDiff("n", Format(.SubItems(2), "HH:NN"), Format(.SubItems(3), "HH:NN"))
                    LHrs1 = (lHrs / 60)
                    If LHrs1 - Int(LHrs1) > 0.5 Then
                        LHrs1 = Int(LHrs1) + 0.5
                    Else
                        LHrs1 = Int(LHrs1)
                    End If
                    lTemp = lHrs - (LHrs1 * 60)
                    
'                    If lHrs > 65 Then
'                        lHrs = 2
'                    ElseIf lHrs > 5 Then
'                        lHrs = 1
'                    End If
                    
                    'lOT = Val(con.Execute("Select OverTimeEntryID From VEmpDailyEntries Where EmpID='" & .Text & "' AND DT='" & DTAttend & "'").Fields(0).value & "")
                    
                    con.Execute "INSERT INTO OverTimeAuth VALUES('" & .Text & "','" & DTAttend & "',1,'" & Format(.SubItems(2), "HH:NN") & "','" & Format(.SubItems(3), "HH:NN") & "'," & LHrs1 & ")"
                    
                    'If lOT = 0 Then
                        con.Execute "INSERT INTO EmpTimes(EmpID,DT,InTime,OutTime,Hrs) VALUES('" & .Text _
                         & "','" & DTAttend & "','" & Format(.SubItems(2), "HH:mm") & "','" & Format(.SubItems(3), "HH:mm") & "'," & LHrs1 & ")"
                    'Else
                    '    con.Execute "Update EmpTimes Set InTime='" & Format(.SubItems(2), "HH:mm") & _
                         "',OutTime='" & Format(.SubItems(3), "HH:mm") & "',Hrs=" & lHrs & _
                         " Where EntryID=" & lOT
                    'End If
                    
                    
                    con.Execute "Update AttendanceSheet SEt ExtraHrs=" & LHrs1 & " Where empID='" & .Text & "' and DT='" & DTAttend & "'"
               

            End If
        End With
    Next i
    con.CommitTrans
Exit Sub

err:
    MsgBox err.Description, vbInformation

End Sub




Private Sub DTAttend_Change()
        
    'Call cmdSave_Click
    Call FillList
    
End Sub

Private Sub Form_Load()

    DTAttend = Format(GetServerDate, "dd-MMM-yyyy")
    
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1600
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID"
    cmbDept.ID = 0

    DTAttend.Enabled = UserHasAccess("AttDateChange")
    
End Sub
Private Sub cmbdept_matched()

    Call FillList
    
    cmbEmp.ClearVals
    cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.Text = "<All Departments>" Then
        cmbEmp.AddVals con, "Name + ' {' + EmpID + '}'", "Employees", "EmpID", "Where active=1"
    Else
        cmbEmp.AddVals con, "Name + ' {' + EmpID + '}'", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "' and active=1 "
    End If
    cmbEmp.ID = "0"
    
End Sub

Private Sub FillList()

    
    If cmbDept.MatchFound = False Then Exit Sub
    
    LV.ListItems.Clear

    Dim rsList As New ADODB.Recordset, cmd As New ADODB.command
    Dim ITM As ListItem, Attend As String
    
    Set cmd.ActiveConnection = con
    cmd.CommandText = IIf(cmbDept.ID = "0", "SP_OverTimeDept", "SP_OverTime")
    cmd.CommandType = adCmdStoredProc
    
    cmd.Parameters("@AttDate").Type = adDBTimeStamp
    cmd.Parameters("@AttDate").value = DTAttend
    
    If cmbDept.ID <> "0" Then
        cmd.Parameters("@DeptID").Type = adVarChar
        cmd.Parameters("@DeptID").value = cmbDept.ID
    End If
    
    Set rsList = cmd.Execute
    Dim iAdd As Integer
    Dim sSecOut As String
    Dim iTemp As Integer
    With rsList
        For i = 1 To .RecordCount
            Set ITM = LV.ListItems.add(, !EmpId, !EmpId & "")
            ITM.ListSubItems.add , , !Name & " " & !Rel & " " & !FName
            ITM.ListSubItems.add , , IIf(IsNull(!Authority), "", "OverTime")
            
            iAdd = Int(Rnd * (Val(5) - (Val(1) - 1)) + (Val(1)))
            
            If IsNull(![TFrom]) Then
                If Not IsNull(![SecondOutTime]) Then
                    'sSecOut = (Val(Left(![SecondOutTime] & "", 2))) & Right(![SecondOutTime] & "", 3)
                    sSecOut = ![SecondOutTime]
                    iTemp = Val(Right(sSecOut, 2)) + iAdd
                    If iTemp > 59 Then
                        sSecOut = (Val(Left(sSecOut, 2)) + 1) & ":" & Format((iAdd - 60), "00")
                    Else
                        sSecOut = Left(sSecOut, 3) & Format(iTemp, "00")
                    End If
                    ITM.SubItems(2) = sSecOut '![SecondOutTime] & ""
                ElseIf IsNull(![FirstOutTime]) = False And IsNull(![SecondInTime]) = True Then
                    
                    sSecOut = ![FirstOutTime]
                    iTemp = Val(Right(sSecOut, 2)) + iAdd
                    If iTemp > 59 Then
                        sSecOut = (Val(Left(sSecOut, 2)) + 1) & ":" & Format((iAdd - 60), "00")
                    Else
                        sSecOut = Left(sSecOut, 3) & Format(iTemp, "00")
                    End If
                    ITM.SubItems(2) = sSecOut '![SecondOutTime] & ""
                    
                End If
            Else
                ITM.SubItems(2) = ![TFrom]
            End If
            
            'ITM.SubItems(2) = IIf(![TFrom] & "" = "", DateAdd("nn", iAdd, ![SecondOutTime] & ""), ![TFrom] & "")
            ITM.SubItems(3) = ![TTo] & ""
            
            'If ITM.ListSubItems(2) <> "" Then ITM.Checked = True: ITM.Tag = "Edit"
            If ![TFrom] & "" <> "" Then ITM.Checked = True: ITM.Tag = "Edit"
            
            .MoveNext
        Next
    .Close
    
    Exit Sub
    
    '''''''''' Will Check For Over Time'''''''''''''''
    .Open "SELECT * FROM VOverTimeAuth WHERE IsNull(DT,true) or DT='" & DTAttend & "' ORDER BY EmpID", con, adOpenStatic, adLockReadOnly
    For i = 1 To .RecordCount
        With LV.ListItems(i)
            If rsList!DT & "" <> "" Then
                .Checked = True
                .SubItems(2) = rsList![TFrom] & ""
                .SubItems(3) = rsList![TTo] & ""
            End If
        End With
        .MoveNext
    Next
    Set rsLedger = Nothing
    .Close
End With
Set rsList = Nothing

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
 
 LV.Move 100, 200, Me.Width - 400, FLV.Height - FButs1.Height - 300
 
 FButs1.Move LV.Left + LV.Width - FButs1.Width, LV.Top + LV.Height
 
 FButs2.Move LV.Left, LV.Top + LV.Height
End Sub


Private Sub LV_DblClick()
    
    If LV.SelectedItem.SubItems(2) = "" Then Exit Sub
    TFrom.Move LV.Left + LV.ColumnHeaders(3).Left + 50, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(3).Width
    TTo.Move LV.Left + LV.ColumnHeaders(4).Left + 50, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(4).Width
    txtHours.Move LV.Left + LV.ColumnHeaders(4).Left + 50, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(4).Width
    
    
    With LV.SelectedItem
        If .SubItems(2) <> "" Then
            TFrom = CDate(.SubItems(2))
        End If
    
        If .SubItems(3) <> "" Then
            TTo = CDate(.SubItems(3))
            TTo.Visible = True
        Else
            txtHours.Visible = True
        End If
    End With
    
    TFrom.Visible = True
    'TTo.Visible = True

    'TFrom.SetFocus
    If txtHours.Visible Then
        txtHours = ""
        txtHours.SetFocus
    Else
        TTo.SetFocus
    End If
End Sub

Private Sub LV_GotFocus()
    TFrom.Visible = False
    TTo.Visible = False
End Sub

Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    
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

If Button = vbRightButton And LV.ListItems.Count > 0 Then
    PopUpMenu mnuPopup
    Exit Sub
End If
    'Set Flag that mouse is in attendance column or not
    forLeave = CBool(x >= LV.ColumnHeaders(3).Left And x <= (LV.ColumnHeaders(3).Left + LV.ColumnHeaders(3).Width))
    

End Sub

Private Sub mnuDelete_Click()

con.Execute "DELETE FROM AttendanceSheet WHERE DT='" & DTAttend.value & "' AND EmpID='" & LV.SelectedItem.Text & "' AND Shift='" & cmbTime.Tag & "'"
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
        
        .SubItems(2) = Format(TFrom.value, "HH:mm")
        .SubItems(3) = Format(TTo.value, "HH:mm")
        
        TTo.Visible = False
        TFrom.Visible = False
    
        LV.SetFocus
    End With
End If

End Sub


Private Sub txtHours_KeyPress(KeyAscii As Integer)

    Dim LHours As Integer, LMinute As Integer
    Dim OHours As Integer, OMinute As Integer, Fine As Double, ExtraHrs As Double

    Dim DefinedWorkingMins As Long, EmpWorkingMins As Long
    Dim MinDiff As Long
    
    Dim sTto As String
    Dim iAdd As Integer
    Dim iTemp As Integer
    
    If KeyAscii = 13 Then
        If Val(txtHours) <= 0 Then
            MsgBox "Invalid Entry.", vbInformation
            Exit Sub
        End If
        With LV.SelectedItem
            
            .Checked = True
        
            .SubItems(2) = Format(TFrom.value, "HH:mm")
            sTto = Format(DateAdd("n", Val(txtHours) * 60, Format(TFrom.value, "HH:mm")), "HH:mm")
            'sTto = Val(Left(Format(TFrom.value, "HH:mm"), 2)) + Val(txtHours) & Right(Format(TFrom.value, "HH:mm"), 3)
            iAdd = Int(Rnd * (Val(5) - (Val(1) - 1)) + (Val(1)))
            
            iTemp = Val(Right(sTto, 2)) + iAdd
            
            If iTemp > 59 Then
                sTto = (Val(Left(sSecOut, 2)) + 1) & ":" & Format((iAdd - 60), "00")
            Else
                sTto = Left(sTto, 3) & Format(iTemp, "00")
            End If
            
            .SubItems(3) = sTto 'Format(TTo.value, "HH:mm")
        
            TTo.Visible = False
            TFrom.Visible = False
            txtHours.Visible = False
            
            LV.SetFocus
        End With
        
    ElseIf KeyAscii = 27 Then
        TTo.Visible = False
        TFrom.Visible = False
        txtHours.Visible = False
    End If

End Sub
