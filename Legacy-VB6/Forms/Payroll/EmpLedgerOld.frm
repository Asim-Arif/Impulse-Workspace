VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form EmpLedgerOld 
   BorderStyle     =   0  'None
   ClientHeight    =   6900
   ClientLeft      =   0
   ClientTop       =   180
   ClientWidth     =   11550
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleMode       =   0  'User
   ScaleWidth      =   10345.49
   ShowInTaskbar   =   0   'False
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   2295
      TabIndex        =   15
      Top             =   390
      Width           =   2235
      _ExtentX        =   3942
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextForeColor   =   12582912
      ListIndex       =   -1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1050
      Top             =   2520
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
            Picture         =   "EmpLedgerOld.frx":0000
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EmpLedgerOld.frx":0452
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   90
      TabIndex        =   3
      Top             =   390
      Width           =   2220
      _ExtentX        =   3916
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextForeColor   =   12582912
      ListIndex       =   -1
   End
   Begin VB.Frame FLV 
      Height          =   6045
      Left            =   30
      TabIndex        =   0
      Top             =   780
      Width           =   11385
      Begin VB.Frame fButs2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   450
         Left            =   30
         TabIndex        =   9
         Top             =   5505
         Visible         =   0   'False
         Width           =   4410
         Begin VB.Label lblTotalPresents 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1395
            TabIndex        =   13
            Top             =   150
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total Presents :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   12
            Top             =   165
            Width           =   1320
         End
         Begin VB.Label lblTotalAbs 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3645
            TabIndex        =   11
            Top             =   150
            Width           =   660
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   2340
            TabIndex        =   10
            Top             =   165
            Width           =   1260
         End
      End
      Begin VB.Frame fButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   480
         Left            =   6225
         TabIndex        =   6
         Top             =   5490
         Width           =   5040
         Begin MSForms.CommandButton cmdPrintForSA 
            Height          =   360
            Left            =   90
            TabIndex        =   17
            Top             =   60
            Visible         =   0   'False
            Width           =   1935
            ForeColor       =   0
            Caption         =   "Print For SA      "
            PicturePosition =   327683
            Size            =   "3413;635"
            Picture         =   "EmpLedgerOld.frx":0564
            Accelerator     =   80
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdPrint 
            Height          =   360
            Left            =   2070
            TabIndex        =   8
            Top             =   60
            Width           =   1425
            ForeColor       =   0
            Caption         =   "Print         "
            PicturePosition =   327683
            Size            =   "2514;635"
            Picture         =   "EmpLedgerOld.frx":0676
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
            Left            =   3555
            TabIndex        =   7
            Top             =   60
            Width           =   1425
            ForeColor       =   0
            Caption         =   " Close      "
            PicturePosition =   327683
            Size            =   "2514;635"
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
      Begin Crystal.CrystalReport cr1 
         Left            =   1950
         Top             =   1020
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin MSComctlLib.ListView LV2 
         Height          =   5505
         Left            =   0
         TabIndex        =   16
         Top             =   180
         Visible         =   0   'False
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   9710
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
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
         NumItems        =   17
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Day"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Status"
            Object.Width           =   1614
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "In Time"
            Object.Width           =   1482
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Out Time"
            Object.Width           =   1535
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Lunch In"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Lunch Out"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "2nd In Time"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "2nd Out Time"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "Hrs"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "OT In"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "OT Out"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   12
            Text            =   "Late Hours"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   13
            Text            =   "Over Time"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   14
            Text            =   "Short Leaves"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   15
            Text            =   "Short Loan"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   16
            Text            =   "Long Loan"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5505
         Left            =   360
         TabIndex        =   14
         Top             =   465
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   9710
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Day"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Status"
            Object.Width           =   1614
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "In Time"
            Object.Width           =   1482
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Out Time"
            Object.Width           =   1535
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Hours"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Late Hours"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Over Time"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Short Leaves"
            Object.Width           =   4628
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Short Loan"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "Long Loan"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin MSForms.ComboBox cmbYear 
      Height          =   300
      Left            =   4530
      TabIndex        =   5
      Top             =   390
      Width           =   1140
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2011;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbMonth 
      Height          =   300
      Left            =   5670
      TabIndex        =   4
      Top             =   390
      Width           =   1320
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2328;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   105
      TabIndex        =   2
      Top             =   150
      Width           =   6855
      BackColor       =   11517387
      VariousPropertyBits=   8388627
      Caption         =   "  Department                         Employee                               Year                Month"
      Size            =   "12091;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   7035
      TabIndex        =   1
      Top             =   360
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
   Begin VB.Menu mnupop 
      Caption         =   "mnupop"
      Begin VB.Menu mnuAbsentLeave 
         Caption         =   "Change Attendance To Absent /Leave"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewSTLoan 
         Caption         =   "New Short Term Advance"
      End
      Begin VB.Menu mnuNewLTLoan 
         Caption         =   "New Long Term Advance"
      End
      Begin VB.Menu dash 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuManualOT 
         Caption         =   "Manual Over Time Feeding"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "EmpLedgerOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TotalPresents As Double, TotalAbsents As Double, TotalLeaves As Double, TotalDays As Double, TotalHolidays As Double
Dim PrintForSA As Boolean

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

'    On Error GoTo ERR
'    Screen.MousePointer = vbHourglass
    Dim DTAtt As Date
    
    con.Execute "Delete  From PrintEmpLedger"
    
    If LV.Visible Then
        For i = 1 To LV.ListItems.Count
            With LV.ListItems(i)
                
                If .Text <> "" Then
                    DT = .Text
                End If
                
                con.Execute "Insert into  PrintEmpLedger (DT,Status,InTime,OutTime,Hrs,LateHours,OverTime,ShortLeave,EmpID)" & _
                "values ('" & DT & "','" & .SubItems(2) & "','" & .SubItems(3) & "','" & .SubItems(4) & "','" _
                & .SubItems(5) & "','" & .SubItems(6) & "','" & .SubItems(7) & "','" & .SubItems(8) & "','" & cmbEmp.ID & "')"

            End With
        Next
    Else
        For i = 1 To LV2.ListItems.Count
            With LV2.ListItems(i)
                If .Text <> "" Then
                    DT = .Text
                End If
                con.Execute "Insert into  PrintEmpLedger (DT,Status,InTime,OutTime,SecondInTime," & _
                "SecondOutTime,Hrs,OtInTime,OTOutTime,LateHours,OverTime,ShortLeave,EmpID)" & _
                " Values ('" & DT & "','" & .SubItems(2) & "','" & .SubItems(3) & "','" & .SubItems(4) & "','" _
                & .SubItems(7) & "','" & .SubItems(8) & "','" & .SubItems(9) & "','" & .SubItems(10) & _
                "','" & .SubItems(11) & "','" & .SubItems(12) & "','" & .SubItems(13) & "','" & .SubItems(14) & "','" & cmbEmp.ID & "')"

            End With
        Next
    End If
    
    
    'With cr1
    '    .ReportFileName = RptPath & "\EmpLedger.rpt"
    '    '.DataFiles(0) = rptpath & "\Payroll.mdb"
    '    '.Connect = ConnectStr
    '    .Formulas(0) = "FromTo='For " & cmbMonth & ", " & cmbYear & ".'"
    '    .Formulas(1) = "EmpNAme='Ledger Of " & cmbEmp.Text & "'"
    '    .Formulas(2) = "TotalLateHrs='" & LV.ListItems(LV.ListItems.Count).ListSubItems(5) & "'"
    '    .Formulas(3) = "TotalOverTime='" & LV.ListItems(LV.ListItems.Count).ListSubItems(6) & "'"
    '    .Formulas(4) = "TotalAbsents=" & TotalAbsents & ""
    '    .Formulas(5) = "TotalPresents=" & TotalPresents & ""
    '    .Formulas(6) = "Company='" & company & "'"
    '    .Action = 1
    '    .PageZoomNext
    'End With
    
    
    

'''''''''''''''''''''''''' Print The Ledger For Production Employees '''''''''''''''''''''''''''''''''

If LV2.Visible Then
    PrintEmpLedgerWithLunchTime (PrintForSA)
    Exit Sub
End If




'''''''''''''''''''''''''' Print The Ledger For Staff Employees '''''''''''''''''''''''''''''''''


    Screen.MousePointer = vbHourglass

Dim f As New frmPrevRpt, rpt As New rptEmpLedger

Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

Set FormulaFields = rpt.FormulaFields
For Each FormulaField In FormulaFields
    If FormulaField.Name = "{@Company}" Then
        FormulaField.Text = "'" & company & "'"
    ElseIf FormulaField.Name = "{@FromTo}" Then
        FormulaField.Text = "'For " & cmbMonth & ", " & cmbYear & ".'"
    ElseIf FormulaField.Name = "{@EmpName}" Then
        FormulaField.Text = "'" & cmbEmp.Text & "'"
    ElseIf FormulaField.Name = "{@TotalLateHrs}" Then
        FormulaField.Text = "'" & LV.ListItems(LV.ListItems.Count).SubItems(5) & "'"
    ElseIf FormulaField.Name = "{@TotalOverTime}" Then
        FormulaField.Text = "'" & LV.ListItems(LV.ListItems.Count).SubItems(6) & "'"
    ElseIf FormulaField.Name = "{@TotalPresents}" Then
        FormulaField.Text = TotalPresents
    ElseIf FormulaField.Name = "{@TotalAbsents}" Then
        FormulaField.Text = TotalAbsents
    End If
Next


'rpt.SQLQueryString = Sql
rpt.RecordSelectionFormula = "{VEmp.EmpID}='" & cmbEmp.ID & "'"
rpt.PrintOut

'f.ShowReport "{VEmp.EmpID}='" & cmbEmp.ID & "' ", rpt

Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPrintForSA_Click()
    PrintForSA = True
    cmdPrint_Click
    PrintForSA = False
End Sub

Public Sub cmdRefresh_Click()

    If cmbEmp.MatchFound Then
        
        EmpType = con.Execute("Select EmpType From Employees Where EmpID='" & cmbEmp.ID & "'").Fields(0)
        
        If EmpType = 2 Then
            Call Refresh2ShiftLedger
        Else
            Call RefreshLedger
        End If
        
    End If
    
    'Call Form_Resize
    
End Sub

Private Sub Form_Load()

    mnupop.Visible = False

    cmbMonth.AddItem "Jan"
    cmbMonth.List(0, 1) = 1
    cmbMonth.AddItem "Feb"
    cmbMonth.List(1, 1) = 2
    cmbMonth.AddItem "Mar"
    cmbMonth.List(2, 1) = 3
    cmbMonth.AddItem "Apr"
    cmbMonth.List(3, 1) = 4
    cmbMonth.AddItem "May"
    cmbMonth.List(4, 1) = 5
    cmbMonth.AddItem "Jun"
    cmbMonth.List(5, 1) = 6
    cmbMonth.AddItem "Jul"
    cmbMonth.List(6, 1) = 7
    cmbMonth.AddItem "Aug"
    cmbMonth.List(7, 1) = 8
    cmbMonth.AddItem "Sep"
    cmbMonth.List(8, 1) = 9
    cmbMonth.AddItem "Oct"
    cmbMonth.List(9, 1) = "10"
    cmbMonth.AddItem "Nov"
    cmbMonth.List(10, 1) = 11
    cmbMonth.AddItem "Dec"
    cmbMonth.List(11, 1) = 12
    
    Dim CYear As Long
    CYear = Format(GetServerDate, "yyyy")
    For i = 2002 To CYear
        cmbYear.AddItem i
    Next i


    cmbMonth.ListIndex = Val(Format$(GetServerDate, "MM")) - 1
    cmbYear = Format$(GetServerDate, "yyyy")



    DtTo = GetServerDate
    DtFrom = DateAdd("M", -1, GetServerDate)
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID"
    cmbDept.Text = "<All Departments>"
    
    'cmbEmp.Text = "<All Employees>"
    'Set LV.SmallIcons = ImageList1
    
    
End Sub
Private Sub cmbdept_matched()

    cmbEmp.ClearVals
    'cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.Text = "<All Departments>" Then
        cmbEmp.AddVals con, "Name", "Employees", "EmpID", "Where Active=1"
    Else
        cmbEmp.AddVals con, "Name", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "' and Active=1"
    End If
    
End Sub

Private Sub Form_Resize()


On Error Resume Next

 'fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
Dim TopMargin As Integer
TopMargin = cmbEmp.Top + cmbEmp.Height + 30

 FLV.Move 50, TopMargin, Me.ScaleWidth - 100, Me.ScaleHeight - TopMargin
 
 LV.Move 100, 200, Me.Width - 400, FLV.Height - fButs1.Height - 300
 LV2.Move 100, 200, Me.Width - 400, FLV.Height - fButs1.Height - 300
 
 fButs1.Move LV.Left + LV.Width - fButs1.Width, LV.Top + LV.Height
 
 fButs2.Move LV.Left, LV.Top + LV.Height
 
End Sub



Private Sub RefreshLedger()


LV2.Visible = False
LV.Visible = True



TotalAbsents = 0
TotalPresents = 0
TotalLeaves = 0
TotalHolidays = 0

    'On Error GoTo ERR
    
    If cmbDept.MatchFound = False Or cmbEmp.MatchFound = False Then Exit Sub
    Dim rsLedger As New ADODB.Recordset
    Dim rsShifts As New ADODB.Recordset
    Dim ITM As ListItem
    Dim Pos As Integer
    Dim LateHours1 As Double
    Dim LateHours2 As Double
    
    Dim DT As Date
    Dim DTs As New Collection
    Dim LastDay As Integer
    Dim TheColor As Long
    Dim iDay As Integer
    Dim OverTime As String
    Dim ShortTerm As String
    Dim LongTerm As String, str As String
    
    Dim rs As New ADODB.Recordset, rsTime As New ADODB.Recordset
    
    'Get Holidays Into DTs Collection Object.
    
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    LastDay = DateAdd("m", 1, DT) - DT
    LV.ListItems.Clear
    


'''''''''''''''''' This Will Update The LateHrs & OverTime ''''''''''
    con.Execute "EXEC SP_CalcOvertimeEmpWise '" & DT & "','" & cmbEmp.ID & "'"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    With rs
        .Open "Select DAY(DT),Description from Holidays Where year(DT)=" & cmbYear & " and Month(DT)=" & cmbMonth.List(cmbMonth.ListIndex, 1) & " Order by DT", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            'DTs.Add .Fields(0).value, CStr(.Fields(0).value)
            DTs.Add .Fields(0).value & "  {" & .Fields(1).value & "}", .Fields(0).value & ""
            
            .MoveNext
        Loop
        TotalHolidays = .RecordCount
        .Close
            
        'Get Records
        .Open "Select * from VEmpLedger Where year(DT)=" & cmbYear & " and Month(DT)=" & cmbMonth.List(cmbMonth.ListIndex, 1) & " and EmpID='" & cmbEmp.ID & "' Order by DT", con, adOpenStatic, adLockReadOnly
    
    TotalDays = LastDay
    
    For i = 1 To LastDay
        
        If i - 1 = Day(Date) And Month(DT) = Month(Date) Then
            TotalDays = Day(Date)
            Exit For
        End If

        'On Error Resume Next 'To Trap the Error
        
        iDay = Day(DT)
        
        'If iDay = 9 Then Stop
        
        TheColor = &H800000
        Desc = ""
        
        If Weekday(DT) = 1 Then
            TheColor = vbRed
            TotalHolidays = TotalHolidays + 1
        Else
            For j = 1 To DTs.Count
                If Val(DTs(j)) = iDay Then
                    TheColor = vbRed
                    Desc = Right(DTs(j), Len(DTs(j)) - 2)
                    TheColor = RGB(20, 150, 20)
                    TotalHol = TotalHol + 1
                    Bold = True
                    DTs.Remove (j)
                    Exit For
                End If
            Next j
        End If
        
        
        Set ITM = LV.ListItems.Add(, Format(DT, "dd-MMM-yyyy"), Format(DT, "dd-MMM-yyyy"))
        ITM.ForeColor = TheColor
        
        ITM.ListSubItems.Add(, , Format(DT, "dddd") & Desc).ForeColor = TheColor
        DT = DateAdd("d", 1, DT)
        rs.Find "AttDay=" & iDay, , adSearchForward, 1
        
        
            If .EOF Then
                For j = 2 To 7
                    ITM.SubItems(j) = "" '"N/A"
                Next j
            Else

                    If Format(![DT], "ddd") = "Sun" Or TheColor = RGB(20, 150, 20) Then
                        'itm.ListSubItems.Add(, , ![Attend] & "").ForeColor = TheColor
                        ITM.ListSubItems.Add(, , "P" & "").ForeColor = TheColor
                    Else
                        ITM.ListSubItems.Add(, , IIf(Val(!Attendance & "") = 0, "A", IIf(Val(!Attendance & "") = 1, "P", "L"))).ForeColor = TheColor
                        If Val(!Attendance & "") = 1 Then
                            TotalPresents = TotalPresents + 1
                        ElseIf Val(!Attendance & "") = 2 Then
                            TotalLeaves = TotalLeaves + 1

                        End If
                    End If
                    
                    If (LV.ListItems(i).ListSubItems(2) = "A" Or LV.ListItems(i).ListSubItems(2) = "L") And Format(![DT], "ddd") <> "Sun" Then
                    
                        ITM.ListSubItems.Add , , ""
                        ITM.ListSubItems.Add , , ""
                    Else


''''''''''''''''''''''''''''''''''''' To Get All The Time Intervals ''''''''''

                        rsTime.Open "SELECT * FROM EmpTimes WHERE EmpID='" & cmbEmp.ID & "' AND DT='" & !DT & "'", con, adOpenStatic, adLockReadOnly
                        If rsTime.RecordCount > 0 Then
                            ITM.ListSubItems.Add(, , rsTime![InTime] & "").ForeColor = TheColor
                            ITM.ListSubItems.Add(, , rsTime![OutTime] & "").ForeColor = TheColor
                            ITM.ListSubItems.Add(, , getHrsMin(Val(rsTime!Hrs & ""))).ForeColor = TheColor
                            rsTime.MoveNext
                        Else
                            ITM.ListSubItems.Add , , ""
                            ITM.ListSubItems.Add , , ""
                            ITM.ListSubItems.Add , , ""
                        End If
                        ITM.ListSubItems.Add(, , getHrsMin(Val(!LateHours & ""), True)).ForeColor = TheColor
                        ITM.ListSubItems.Add(, , getHrsMin(Val(!OverTimeHours & ""), True)).ForeColor = TheColor
                        str = con.Execute("SELECT dbo.FShortLeaves ('" & DateSerial(cmbYear, cmbMonth.ListIndex + 1, i) & "','" & cmbEmp.ID & "')").Fields(0) & ""
                        ITM.SubItems(8) = str

                        For j = 1 To rsTime.RecordCount - 1
                            Set ITM = LV.ListItems.Add(, , "")
                            ITM.SubItems(3) = rsTime!InTime & ""
                            ITM.SubItems(4) = rsTime!OutTime & ""
                            ITM.SubItems(5) = getHrsMin(Val(rsTime!Hrs & ""))
                            rsTime.MoveNext
                        Next
                        rsTime.Close
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                    End If
            End If
        Next i
    End With
    Set rsLedger = Nothing
    Sql = ""
    Cond = ""
    ICO = ""
    
    'TotalAbsents = TotalDays - (TotalPresents + TotalLeaves + TotalHolidays)
    
    'Replace the TotalAbsents and TotalPresents with TotalDays and Hours
    
    TotalPresents = TotalDays - TotalHolidays
    TotalAbsents = TotalPresents * 8
    
    Exit Sub
ERR:
        MsgBox ERR.Number & vbNewLine & ERR.Description
        Exit Sub
        
End Sub



Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Or LV.ListItems.Count = 0 Then Exit Sub
    PopUpMenu mnupop
End Sub

Private Sub mnuAbsentLeave_Click()
    With frmLeaves
        .cmbEmp.Text = cmbEmp.Text
        .DTPicker1 = CDate(LV.SelectedItem.Text)
        .Show 1
    End With
End Sub

Private Sub mnuManualOT_Click()
    With frmManualOTFeeding
        .cmbEmp.Text = cmbEmp.Text
        .Show 1
    End With
End Sub

Private Sub mnuNewSTLoan_Click()

    With NewAdvanceShort
        .cmbDept.Text = cmbDept.Text
        .cmbEmp.ID = cmbEmp.ID
        .LongTerm = False
        .Show 1
    End With
    
    
End Sub

Private Sub mnuNewLTLoan_Click()
    
    With NewAdvanceShort
        .cmbDept.Text = cmbDept.Text
        .cmbEmp.ID = cmbEmp.ID
        .LongTerm = True
        .DTPicker1 = #6/30/2004#
        .txtDesc = "Previous Balance Of Fixed Loan"
        .Show 1
    End With
    
End Sub





Private Sub Refresh2ShiftLedger()


LV2.Visible = True
LV.Visible = False



TotalAbsents = 0
TotalPresents = 0
TotalLeaves = 0
    'On Error GoTo ERR
    
    If cmbDept.MatchFound = False Or cmbEmp.MatchFound = False Then Exit Sub
    Dim rsLedger As New ADODB.Recordset
    Dim rsShifts As New ADODB.Recordset
    Dim ITM As ListItem
    Dim Pos As Integer
    Dim LateHours1 As Double
    Dim LateHours2 As Double
    
    Dim DT As Date
    Dim DTs As New Collection
    Dim LastDay As Integer
    Dim TheColor As Long
    Dim iDay As Integer
    Dim OverTime As String
    Dim ShortTerm As String
    Dim LongTerm As String, str As String
    
    Dim rs As New ADODB.Recordset, rsTime As New ADODB.Recordset
    
    Dim LeaveType As String
    
    
    
    
    TotalPresents = 0
    TotalHolidays = 0
    TotalAbsents = 0
    
    'Get Holidays Into DTs Collection Object.
    
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    LastDay = DateAdd("m", 1, DT) - DT
    LV2.ListItems.Clear
    


    '''''''''''''''''' This Will Update The LateHrs & OverTime ''''''''''
        con.Execute "EXEC SP_CalcOvertimeEmpWise '" & DT & "','" & cmbEmp.ID & "'"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    With rs
        .Open "Select DAY(DT),Description from Holidays Where year(DT)=" & cmbYear & " and Month(DT)=" & cmbMonth.List(cmbMonth.ListIndex, 1) & " Order by DT", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            'DTs.Add .Fields(0).value, CStr(.Fields(0).value)
            DTs.Add .Fields(0).value & "  {" & .Fields(1).value & "}", .Fields(0).value & ""
            
            .MoveNext
        Loop
        .Close
            
        'Get Records
        .Open "Select * from VEmpLedger Where year(DT)=" & cmbYear & " and Month(DT)=" & cmbMonth.List(cmbMonth.ListIndex, 1) & " and EmpID='" & cmbEmp.ID & "' Order by DT", con, adOpenStatic, adLockReadOnly

    For i = 1 To LastDay
    
        TotalDays = i
        
        If i - 1 = Day(Date) And Month(DT) = Month(Date) Then
            TotalDays = i
            Exit For
        End If

        'On Error Resume Next 'To Trap the Error
        
        iDay = Day(DT)
        
        'If iDay = 9 Then Stop
        
        TheColor = &H800000
        Desc = ""
        
        If Weekday(DT) = 1 Then
            TheColor = vbRed
            TotalHolidays = TotalHolidays + 1
        Else
            For j = 1 To DTs.Count
                If Val(DTs(j)) = iDay Then
                    TheColor = vbRed
                    Desc = Right(DTs(j), Len(DTs(j)) - 2)
                    TheColor = RGB(20, 150, 20)
                    TotalHol = TotalHol + 1
                    Bold = True
                    DTs.Remove (j)
                    Exit For
                End If
            Next j
        End If
        
        
        
        
        
        
        Set ITM = LV2.ListItems.Add(, Format(DT, "dd-MMM-yyyy"), Format(DT, "dd-MMM-yyyy"))
        ITM.ForeColor = TheColor
        
        ITM.ListSubItems.Add(, , Format(DT, "dddd") & Desc).ForeColor = TheColor
        DT = DateAdd("d", 1, DT)
        rs.Find "AttDay=" & iDay, , adSearchForward, 1
        
        
            If .EOF Then
            
                For j = 2 To 7
                    ITM.SubItems(j) = "" '"N/A"
                Next j
                
            Else
            
                    ' Determine The Leave Type
                    If ![CL] & "" Then
                        LeaveType = "( Casual Leave )"
                    ElseIf ![SL] & "" Then
                        LeaveType = "( Sick Leave )"
                    ElseIf ![AL] & "" Then
                        LeaveType = "( Annual Leave )"
                    ElseIf ![CPL] & "" Then
                        LeaveType = "( Compensatory Leave )"
                    ElseIf ![LWP] & "" Then
                        LeaveType = "( Leave Without Pay )"
                    ElseIf ![ML] & "" Then
                        LeaveType = "( Maternity Leave )"
                    End If
                                        

                    If Format(![DT], "ddd") = "Sun" Or TheColor = RGB(20, 150, 20) Then
                        'itm.ListSubItems.Add(, , ![Attend] & "").ForeColor = TheColor
                        ITM.ListSubItems.Add(, , "P" & "").ForeColor = TheColor
                    Else
                        ITM.ListSubItems.Add(, , IIf(Val(!Attendance & "") = 0, "A", IIf(Val(!Attendance & "") = 1, "P", "L"))).ForeColor = TheColor
                        If Val(!Attendance & "") = 1 Then
                            TotalPresents = TotalPresents + 1
                        ElseIf Val(!Attendance & "") = 2 Then
                            TotalLeaves = TotalLeaves + 1
                        End If
                    End If
                    
                    If (LV2.ListItems(i).ListSubItems(2) = "A" Or LV2.ListItems(i).ListSubItems(2) = "L") And Format(![DT], "ddd") <> "Sun" Then
                    
                        ITM.ListSubItems.Add , , ""
                        ITM.ListSubItems.Add , , ""
                        ITM.SubItems(14) = LeaveType
                    Else


''''''''''''''''''''''''''''''''''''' To Get All The Time Intervals ''''''''''

                        rsTime.Open "SELECT * FROM VEmpTimes2 WHERE EmpID='" & cmbEmp.ID & "' AND DT='" & !DT & "' ORder By DT,EmpID,FirstEntryID ", con, adOpenForwardOnly, adLockReadOnly
                        If rsTime.RecordCount > 0 Then
                            ITM.ListSubItems.Add(, , rsTime![FirstInTime] & "").ForeColor = TheColor
                            ITM.ListSubItems.Add(, , rsTime![FirstOutTime] & "").ForeColor = TheColor
                            ITM.ListSubItems.Add(, , rsTime![FirstOutTime] & "").ForeColor = TheColor
                            ITM.ListSubItems.Add(, , rsTime![SecondInTime] & "").ForeColor = TheColor
                            ITM.ListSubItems.Add(, , rsTime![SecondInTime] & "").ForeColor = TheColor
                            ITM.ListSubItems.Add(, , rsTime![SecondOutTime] & "").ForeColor = TheColor
                            
                            ITM.ListSubItems.Add(, , getHrsMin(Val(rsTime!Hrs & ""))).ForeColor = TheColor
                            
                            ITM.ListSubItems.Add(, , rsTime![OTInTime] & "").ForeColor = TheColor
                            ITM.ListSubItems.Add(, , rsTime![OTOutTime] & "").ForeColor = TheColor
                            ITM.ListSubItems.Add().ForeColor = TheColor
                            ITM.ListSubItems.Add(, , getHrsMin(Val(rsTime!OTHrs & ""))).ForeColor = TheColor
                            
                            rsTime.MoveNext
                        Else
                            ITM.ListSubItems.Add , , ""
                            ITM.ListSubItems.Add , , ""
                            ITM.ListSubItems.Add , , ""
                            ITM.ListSubItems.Add , , ""
                            ITM.ListSubItems.Add , , ""
                            ITM.ListSubItems.Add , , ""
                            ITM.ListSubItems.Add , , ""
                            ITM.ListSubItems.Add , , ""
                            
                        End If
                        
                        ITM.ListSubItems.Add(, , getHrsMin(Val(!LateHours & ""), True)).ForeColor = TheColor
                        ITM.ListSubItems.Add(, , getHrsMin(Val(!OverTimeHours & ""), True)).ForeColor = TheColor
                        str = con.Execute("SELECT dbo.FShortLeaves ('" & DateSerial(cmbYear, cmbMonth.ListIndex + 1, i) & "','" & cmbEmp.ID & "')").Fields(0) & ""
                        
                        
                        
                        ITM.SubItems(14) = str

                        'For j = 1 To rsTime.RecordCount - 1
                        '    Set ITM = LV2.ListItems.Add(, , "")
                        '    ITM.SubItems(3) = rsTime!InTime & ""
                        '    ITM.SubItems(4) = rsTime!OutTime & ""
                        '    ITM.SubItems(5) = getHrsMin(Val(rsTime!Hrs & ""))
                        '    rsTime.MoveNext
                        'Next
                        
                        rsTime.Close
                        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                    End If
            End If
        Next i
    End With
    Set rsLedger = Nothing
    Sql = ""
    Cond = ""
    ICO = ""
    
    
'    For i = 1 To LV2.ListItems.Count
'        If LV2.ListItems(i).Text <> "" And (LV2.ListItems(i).SubItems(2) = "" Or LV2.ListItems(i).SubItems(2) = "A") Then
'            TotalAbsents = TotalAbsents + 1
'        End If
'    Next
    
    'TotalAbsents = TotalDays - (TotalPresents + TotalLeaves + TotalHolidays)
    
    'Replace the TotalAbsents and TotalPresents with TotalDays and Hours
    
    ' at The moment -1 is applied to remove extra day ...
    ' to be  Checked after
    
      TotalPresents = TotalDays - TotalHolidays - 1
      TotalAbsents = TotalPresents * 8
    
    
    Exit Sub
ERR:
        MsgBox ERR.Number & vbNewLine & ERR.Description
        Exit Sub

End Sub




Private Sub PrintEmpLedgerWithLunchTime(Optional ForSA As Boolean)

Screen.MousePointer = vbHourglass



Dim f As New frmPrevRpt, rpt As New rptEmpLedgerWithLunchTime

Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

Set FormulaFields = rpt.FormulaFields
For Each FormulaField In FormulaFields
    If FormulaField.Name = "{@Company}" Then
        FormulaField.Text = "'" & company & "'"
    ElseIf FormulaField.Name = "{@FromTo}" Then
        FormulaField.Text = "'For " & cmbMonth & ", " & cmbYear & ".'"
    ElseIf FormulaField.Name = "{@EmpName}" Then
        FormulaField.Text = "'" & cmbEmp.Text & "'"
    ElseIf FormulaField.Name = "{@TotalLateHrs}" Then
        FormulaField.Text = "'" & LV2.ListItems(LV2.ListItems.Count).SubItems(5) & "'"
    ElseIf FormulaField.Name = "{@TotalOverTime}" Then
        FormulaField.Text = "'" & LV2.ListItems(LV2.ListItems.Count).SubItems(6) & "'"
    ElseIf FormulaField.Name = "{@TotalPresents}" Then
        FormulaField.Text = TotalPresents
    ElseIf FormulaField.Name = "{@TotalAbsents}" Then
        FormulaField.Text = TotalAbsents
    ElseIf FormulaField.Name = "{@ForSA}" Then
        FormulaField.Text = ForSA
    End If
Next

'SelFormula = "Year({Holidays.DT})=" & cmbYear
rpt.SQLQueryString = Sql

rpt.PrintOut

'f.ShowReport Cond & " ", rpt

Screen.MousePointer = vbDefault

End Sub
