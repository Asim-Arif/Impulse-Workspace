VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmLeaves 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Leaves / Absent"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   8430
   StartUpPosition =   1  'CenterOwner
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   315
      Left            =   135
      TabIndex        =   0
      Top             =   330
      Width           =   2610
      _ExtentX        =   4604
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
      Left            =   2775
      TabIndex        =   1
      Top             =   330
      Width           =   5535
      _ExtentX        =   9763
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
   Begin VB.CheckBox chkPrintGatePass 
      Caption         =   "&Print Gatepass"
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
      Height          =   240
      Left            =   135
      TabIndex        =   31
      Top             =   4650
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Frame FraShortLeave 
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   3270
      TabIndex        =   21
      Top             =   1950
      Visible         =   0   'False
      Width           =   3405
      Begin VB.TextBox txtHrs 
         Height          =   315
         Left            =   2475
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   255
         Width           =   420
      End
      Begin MSComCtl2.DTPicker DTFrom 
         Height          =   315
         Left            =   0
         TabIndex        =   24
         Top             =   255
         Width           =   1230
         _ExtentX        =   2170
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
         CustomFormat    =   "hh:mm  tt"
         Format          =   344588291
         UpDown          =   -1  'True
         CurrentDate     =   37384
      End
      Begin MSComCtl2.DTPicker DTTo 
         Height          =   315
         Left            =   1245
         TabIndex        =   25
         Top             =   255
         Width           =   1230
         _ExtentX        =   2170
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
         CustomFormat    =   "hh:mm  tt"
         Format          =   344588291
         UpDown          =   -1  'True
         CurrentDate     =   37384
      End
      Begin VB.Label lblFrom 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   " From                To                        Hrs"
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
         Left            =   15
         TabIndex        =   26
         Top             =   0
         Width           =   2880
      End
   End
   Begin VB.OptionButton OptType 
      Caption         =   "Absent"
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
      Height          =   285
      Index           =   0
      Left            =   1785
      TabIndex        =   4
      Top             =   2295
      Width           =   900
   End
   Begin VB.OptionButton OptType 
      Caption         =   "Leave"
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
      Height          =   285
      Index           =   1
      Left            =   1785
      TabIndex        =   3
      Top             =   2055
      Value           =   -1  'True
      Width           =   840
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   90
      TabIndex        =   10
      Top             =   2040
      Width           =   2985
      Begin VB.OptionButton OptShortLeave 
         Caption         =   "Short Leave"
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
         Height          =   255
         Left            =   1650
         TabIndex        =   27
         Top             =   15
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton OptFullDay 
         Caption         =   "Full Day Leave"
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
         Height          =   255
         Left            =   30
         TabIndex        =   13
         Top             =   15
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton OptHalfDay 
         Caption         =   "Half Day Leave"
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
         Height          =   255
         Left            =   60
         TabIndex        =   12
         Top             =   240
         Width           =   1725
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   135
      TabIndex        =   2
      Top             =   900
      Width           =   2610
      _ExtentX        =   4604
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
      CustomFormat    =   "dddd dd-MMM yyyy"
      Format          =   163315715
      CurrentDate     =   37384
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   135
      TabIndex        =   14
      Top             =   1470
      Width           =   2610
      _ExtentX        =   4604
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
      CustomFormat    =   "dddd dd-MMM yyyy"
      Format          =   163315715
      CurrentDate     =   37384
   End
   Begin ComboList.Usercontrol1 cmbAuth 
      Height          =   315
      Left            =   3885
      TabIndex        =   20
      Top             =   3060
      Width           =   4365
      _ExtentX        =   7699
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
   Begin MSComctlLib.ListView LVBalances 
      Height          =   1200
      Left            =   2775
      TabIndex        =   32
      Top             =   660
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2117
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1640
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Casual"
         Object.Width           =   1270
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Sick"
         Object.Width           =   1270
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Annual"
         Object.Width           =   1270
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Special"
         Object.Width           =   1270
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "LWP"
         Object.Width           =   1270
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "HL"
         Object.Width           =   1270
      EndProperty
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "Print When Saved"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5175
      TabIndex        =   9
      Top             =   1785
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.OptionButton opts 
      Caption         =   "Without Deduction"
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
      Height          =   285
      Index           =   0
      Left            =   5175
      TabIndex        =   16
      Top             =   1245
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.OptionButton opts 
      Caption         =   "With Deduction"
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
      Height          =   285
      Index           =   1
      Left            =   5175
      TabIndex        =   17
      Top             =   855
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.CheckBox chkFull 
      Caption         =   "Full Day Leave"
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
      Height          =   240
      Left            =   5175
      TabIndex        =   23
      Top             =   1545
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " To"
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
      Height          =   240
      Index           =   2
      Left            =   135
      TabIndex        =   33
      Top             =   1215
      Width           =   2610
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Leave Type"
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
      Left            =   135
      TabIndex        =   29
      Top             =   2820
      Width           =   3720
   End
   Begin MSForms.ComboBox cmbLeaveType 
      Height          =   285
      Left            =   135
      TabIndex        =   28
      Top             =   3060
      Width           =   3720
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "6562;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Authority Person :"
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
      Index           =   1
      Left            =   3885
      MouseIcon       =   "frmleaves.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   2820
      Width           =   4350
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Leave Reason"
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
      Index           =   6
      Left            =   135
      TabIndex        =   19
      Top             =   3375
      Width           =   8115
   End
   Begin MSForms.TextBox txtDesc 
      Height          =   570
      Left            =   135
      TabIndex        =   18
      Top             =   3600
      Width           =   8145
      VariousPropertyBits=   -1467987941
      ScrollBars      =   3
      Size            =   "14367;1005"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton cmdPrintLeaves 
      Height          =   360
      Left            =   1215
      TabIndex        =   15
      Top             =   4275
      Width           =   1710
      ForeColor       =   0
      Caption         =   "Print Balances"
      PicturePosition =   327683
      Size            =   "3016;635"
      Accelerator     =   76
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " From"
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
      Height          =   240
      Index           =   0
      Left            =   135
      TabIndex        =   11
      Top             =   645
      Width           =   2610
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Department                                  Employee"
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
      Index           =   4
      Left            =   135
      TabIndex        =   8
      Top             =   90
      Width           =   8160
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   360
      Left            =   4785
      TabIndex        =   6
      Top             =   4275
      Width           =   1695
      Caption         =   "Save & Close"
      PicturePosition =   327683
      Size            =   "2990;635"
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdNext 
      Height          =   360
      Left            =   2985
      TabIndex        =   5
      Top             =   4275
      Width           =   1725
      Caption         =   "Save & New    "
      PicturePosition =   327683
      Size            =   "3043;635"
      Accelerator     =   78
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   360
      Left            =   6525
      TabIndex        =   7
      Top             =   4275
      Width           =   1710
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "3016;635"
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
Attribute VB_Name = "frmLeaves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public edit As Boolean

Private Sub PrintReport()

    Screen.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report, ID As Integer
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptLeaveSlip.rpt")
    
    ID = Val(con.Execute("Select MAX(EntryID) FROM Leaves").Fields(0) & "")
    
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
    
    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & strCompany & "'"
        End If
    Next
    
    'rpt.SQLQueryString = Sql
    Cond = "{VLeaves.EntryID}=" & ID & ""
    
    Me.Hide
    f.ShowReport Cond & " ", rpt
    
    Screen.MousePointer = vbDefault


End Sub


Private Sub cmbdept_matched()
    
    cmbEmp.ClearVals
    
    If cmbDept.Text = "<All Departments>" Then
        cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' - ' + FName + ' {' + Designation + '}'", "Employees", "EmpID", "Where Active=1"
    Else
        cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' - ' + FName + ' {' + Designation + '}'", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "' and ACtive=1"
    End If
    
End Sub

Private Function Saved() As Boolean

    Dim Affect As Long, First As Integer, Second As Integer, IsFirst As Boolean, IsSecond As Boolean
    Dim Full As Integer, Hrs As Double, TFrom As String, TTo As String
    Dim PrevCasualLeaves  As Integer
    Dim dHrs As Double, dLeave As Double
    Saved = False
    
    'On Error GoTo err
    Dim SQL As String, Cond As String, Deduct As Integer
            
    Dim iLoop As Integer
    Dim iTotal As Double
    Dim iTotalLimit As Double
    Dim iTotalTaken As Double
    Dim DT As Date
    Dim StartDT As Date
    Dim FinalDT As Date
        
    If Not cmbEmp.MatchFound Then
        MsgBox "Please Select Valid Employee From The List.", vbInformation
        Exit Function
    End If
    
    If DTPicker2.Value < DTPicker1.Value Then
        MsgBox "Invalid Date Range Selected.", vbInformation
        Exit Function
    End If
    
    If IsProbation(cmbEmp.ID, DTPicker1.Value) And OptType(1) Then
        MsgBox "Employee is still in his Probation. Can't Take Leave.", vbInformation
        Exit Function
    End If
    
    iLoop = DateDiff("d", DTPicker1, DTPicker2)
    iTotal = iLoop + 1
    
    Cond = " Where DT Between '" & DTPicker1 & "' AND '" & DTPicker2 & "' AND EmpID='" & cmbEmp.ID & "'"
      
    If OptType(1) Then  'Leave Selected
        If Not cmbEmp.MatchFound Then
                MsgBox "Please Select Valid Employee From The List.", vbInformation
                Exit Function
        ElseIf Not cmbAuth.MatchFound Then
            MsgBox "Authoriy Person That You Entered Does Not Exits !", vbInformation
            'cmbAuth.SetFocus
            Exit Function
        'ElseIf IsPrevTime And chkFull = vbUnchecked Then
        '    MsgBox "You Have Already Taken Leave For This Time !", vbInformation
        '    DTFrom.SetFocus
        '    Exit Function
        ElseIf cmbAuth.ID = cmbEmp.ID Then
            MsgBox "Employee And Authority Person Could Not Be Same !", vbInformation
            cmbAuth.SetFocus
            Exit Function
        ElseIf DtFrom.Value >= DtTo.Value And OptShortLeave.Value = True Then
            MsgBox "Out Value Could Not Be Greater/Equal In Value  !", vbInformation
            DtFrom.SetFocus
            Exit Function
        ElseIf cmbLeaveType.ListIndex < 0 Then
            MsgBox "Please Select Leave Type", vbInformation
            cmbLeaveType.SetFocus
            Exit Function
        End If

        If Month(DTPicker1) < 7 Then
            If strCompany = "Tecno" Then
                StartDT = DateSerial(year(DTPicker1), 1, 1)  ' '7/1/' + Cast(Year(@DT)-1 As Varchar)
            Else
                StartDT = DateSerial(year(DTPicker1) - 1, 7, 1) ' '7/1/' + Cast(Year(@DT)-1 As Varchar)
            End If
        Else
            If strCompany = "Tecno" Then
                StartDT = DateSerial(year(DTPicker1), 1, 1)
            Else
                StartDT = DateSerial(year(DTPicker1), 7, 1)
            End If
        End If
        If strCompany = "Tecno" Then
            FinalDT = DateSerial(year(StartDT), 12, 31)
        Else
            FinalDT = DateSerial(year(StartDT) + 1, 6, 30)
        End If
        
        Dim FieldName As String
              
'        If OptFullDay Then
'            FieldName = "," & cmbLeaveType.List(cmbLeaveType.ListIndex, 1)
'        Else
'            FieldName = ""
'        End If
'
        If OptShortLeave Then
            TFrom = Format(DtFrom, "HH:mm")
            TTo = Format(DtTo, "HH:mm")
            dHrs = Round(DateDiff("n", TFrom, TTo) / 60, 1)
        ElseIf OptHalfDay Then
            dHrs = 4
            Dim dAttPayableHrs As Double
            dAttPayableHrs = GetSingleLongValue("SUM(PayableHrs)", "EmpTimes", " WHERE EmpID='" & cmbEmp.ID & "' AND DT='" & DTPicker1.Value & "' AND OverTime=0")
            If dAttPayableHrs > 4 Then
                dHrs = 8 - dAttPayableHrs
            End If
        ElseIf OptFullDay Then
            dHrs = 8
        End If
        dLeave = dHrs / 8
        
        Dim dLeavestoTake As Double
        If OptFullDay.Value Then
            dLeavestoTake = iTotal
        Else
            dLeavestoTake = dLeave
        End If
        Dim strField As String
        strField = cmbLeaveType.List(cmbLeaveType.ListIndex, 1)
        FieldName = "," & cmbLeaveType.List(cmbLeaveType.ListIndex, 1)
        iTotalTaken = con.Execute("SELECT ISNULL(SUM(CASE WHEN " & strField & "<0.5 THEN 0.5 ELSE " & strField & " END),0) FROM Leaves WHERE DT BETWEEN '" & StartDT & "' AND '" & FinalDT & "' AND EmpID='" & cmbEmp.ID & "'").Fields(0).Value
        iTotalLimit = Val(cmbLeaveType.List(cmbLeaveType.ListIndex, 2))
        
        If iTotalLimit <> 0 And OptType(1) Then
                If strField = "TML" Then
                    If iTotalTaken > 0 Then
                        MsgBox "Already Taken Three Months Leave.", vbInformation
                        Exit Function
                    End If
                Else
                    If iTotalTaken + dLeavestoTake > iTotalLimit Then
                        MsgBox "The Date Range Exceeds The Limit Of " & cmbLeaveType.Text & " Leaves." & vbNewLine & "Can't Add Leaves.", vbInformation
                        Exit Function
                    End If
                End If
        End If
    End If
    
     
    Dim rs As New ADODB.Recordset
    'First = Val(con.Execute("Select COUNT(*) FROM AttendanceSheet " & Cond).Fields(0) & "")
    
        'Put Code Here If You Want To Confirm The Deletion Of Attendance
        If OptFullDay.Value Then
            With rs
                .Open "Select COUNT(*) FROM AttendanceSheet " & Cond, con, adOpenForwardOnly, adLockReadOnly
                First = .Fields(0)
                .Close
            End With
            If First > 0 Then
                If MsgBox("The Attendance of this employee is already taken for the selected Dates. Do you want to Delete ?", vbYesNo + vbQuestion) = vbNo Then
                    Exit Function
                Else
                    con.Execute "DELETE FROM AttendanceSheet " & Cond
                    con.Execute "DELETE FROM EmpTimes " & Cond
                    con.Execute "DELETE FROM Leaves " & Cond
                End If
            End If
        ElseIf OptShortLeave.Value Then
            Dim strIn As String, strOut As String
            strIn = Format(DtFrom, "HH:nn")
            strOut = Format(DtTo, "HH:nn")
            If strIn > "17:00" Then
                MsgBox "Invalid Out Time.", vbInformation
                Exit Function
            End If
             With rs
                '.Open "Select Count(*) From EmpTimes " & Cond & " AND ('" & strIn & "' Between InTime AND OutTime OR '" & strOut & _
                 "' Between InTime AND OutTime OR InTime>='" & strIn & "')", con, adOpenForwardOnly, adLockReadOnly
                .Open "Select Count(*) From EmpTimes " & Cond & " AND ('" & strIn & "' Between InTime AND OutTime OR '" & strOut & _
                 "' Between InTime AND OutTime)", con, adOpenForwardOnly, adLockReadOnly
                 
                First = .Fields(0)
                .Close
                If First > 0 Then
                    MsgBox "The Attendance of this employee is already taken for the selected Period.", vbInformation
                    Exit Function
                End If
                .Open "Select Count(*) From Leaves " & Cond, con, adOpenForwardOnly, adLockReadOnly
                First = .Fields(0)
                .Close
                If First > 0 Then
                    MsgBox "Leave is already Feeded.", vbInformation
                    Exit Function
                End If
            End With
        End If
   
    
    Set rs = Nothing
    
    Call StartTrans(con)
    
'    con.Execute "DELETE FROM AttendanceSheet " & Cond
'    con.Execute "DELETE FROM EmpTimes " & Cond
'    con.Execute "DELETE FROM Leaves " & Cond
    
    For i = 0 To iLoop
        DT = DateAdd("d", i, DTPicker1)
        If DateIsValid(DT) Then
            If OptType(0).Value Then
                con.Execute "INSERT INTO AttendanceSheet(EmpID,DT,Attendance) VALUES('" & cmbEmp.ID & "','" & DT & "',0)"
            Else
                On Error Resume Next    'Just Resume Next,
                If Opts(1).Value Then
                    con.Execute "INSERT INTO AttendanceSheet(EmpID,DT,Attendance) VALUES('" & cmbEmp.ID & "','" & DT & "',3)"
                Else
                    con.Execute "INSERT INTO AttendanceSheet(EmpID,DT,Attendance) VALUES('" & cmbEmp.ID & "','" & DT & "',2)"
                End If
                On Error GoTo 0
                '''''''''' WILL INSERT INTO LEAVES '''''''''''
                con.Execute "INSERT INTO Leaves(EmpID,DT,AuthPerson,Reason," & _
                 "[From],[To],[Full],Hrs,Deduct" & FieldName & ") VALUES('" & cmbEmp.ID _
                 & "','" & DT & "','" & cmbAuth.Text & "','" & txtDesc _
                 & "','" & TFrom & "','" & TTo & _
                 "','" & Full & "'," & dHrs & "," & Deduct & _
                 IIf(FieldName = "", ")", "," & dLeave & ")")
            End If
        End If
    Next
    
    con.CommitTrans
        
    Saved = True
    Unload Me
    SQL = ""
    
    Exit Function
    
err:
    con.RollbackTrans
    MsgBox err.Description
    
End Function

Private Sub TakeAtt(DT As Date)
    
    
     
End Sub

''''''''    Will Select The Attendance From AttendanceSheet '''''''''''

Private Function IsPrevLeave(Cond As String) As Boolean
    
    Dim First As Integer
    First = Val(con.Execute("Select COUNT(*) FROM AttendanceSheet " & Cond & "").Fields(0) & "")
    If First <> 0 Then
        If MsgBox("The Attendance For This Employee Is Already Taken,Do You Want To Delete That", vbInformation + vbYesNo) = vbYes Then
            con.Execute "DELETE FROM AttendanceSheet " & Cond
            con.Execute "DELETE FROM EmpTimes " & Cond
            con.Execute "DELETE FROM Leaves " & Cond & " AND Full=1"
            IsPrevLeave = True
        Else
            IsPrevLeave = False
        End If
    End If
    
End Function

'Private Function IsPrevEntry(Cond As String) As Boolean
'On Error Resume Next
'Dim Shift As String
'Shift = con.Execute("Select Shift FROM Leaves " & Cond & "").Fields(0)
'If Shift = "" Then
'    IsPrevEntry = False: Exit Function
'Else
'    If Shift = "Complete Day" Then
'        IsPrevEntry = True: Exit Function
'    End If
'    If Shift = cmbShift Then
'        IsPrevEntry = True: Exit Function
'    End If
'End If
'End Function

Private Function IsPrevTime() As Boolean


Dim Temp As New ADODB.Recordset
Dim TimeFrom  As Date, TimeTo As Date, CTimeFrom As Date, CTimeTo As Date

With Temp
    .CursorLocation = adUseClient
    .Open "SELECT * FROM Leaves WHERE DT='" & DTPicker1 & "' AND EmpID='" & cmbEmp.ID & "'", con, adOpenStatic, adLockReadOnly
    
    CTimeFrom = Format(DtFrom, "HH:mm")
    CTimeTo = Format(DtTo, "HH:mm")
            
    If .RecordCount > 0 Then
        For i = 1 To .RecordCount
            TimeFrom = !From & ""
            TimeTo = !to & ""
            
            If CTimeFrom > TimeFrom And CTimeTo < TimeTo Then
                IsPrevTime = True
                Exit For
            ElseIf CTimeFrom = TimeFrom Then
                IsPrevTime = True
                Exit For
            End If
            .MoveNEXT
        Next
    
    End If
    
End With

End Function

Private Sub cmbEmp_matched()
'    Dim Emptype As Long
'    Emptype = con.Execute("Select EmpType From Employees Where EmpID='" & cmbEmp.ID & "'").Fields(0)
    
    'cmbAuth.ClearVals
    'cmbAuth.ClearVals
    'cmbAuth.AddVals con, "PersonName", "LeaveAuthPersons", "EntryID", " Order By PersonName"
'    If Emptype = 2 Then
'        cmbAuth.AddItem "Ahsan Naeem", "0"
'        cmbAuth.AddItem "Mohsin Naeem", "1"
'        cmbAuth.AddItem "Aslam Pervaiz", "2"
''        cmbAuth.AddItem "Asifa Lone", "3"
''        cmbAuth.AddItem "Muhammad Wasiq", "4"
''        cmbAuth.AddItem "Abid Butt", "5"
'    Else
'        cmbAuth.AddItem "Ahsan Naeem", "0"
'        cmbAuth.AddItem "Mohsin Naeem", "1"
'        cmbAuth.AddItem "Aslam Pervaiz", "2"
''        cmbAuth.AddItem "Asifa Lone", "3"
''        cmbAuth.AddItem "Muhammad Wasiq", "4"
''        cmbAuth.AddItem "Abid Butt", "5"
'    End If
        
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "GetLeaveBalances"
        .Parameters("@EmpID").Value = cmbEmp.ID
        .Parameters("@DT").Value = DTPicker1.Value 'DateSerial(DTPicker1.Year, DTPicker1.Month, 1)
        Set rs = .Execute
    End With
    Dim LColor As Long
    With rs
        Set ITM = LVBalances.ListItems(2)
        ITM.ListSubItems.Clear
        ITM.SubItems(1) = IIf(Val(!CL_E & "") = 0, "-", Val(!CL_E & ""))
        ITM.SubItems(2) = IIf(Val(!SL_E & "") = 0, "-", Val(!SL_E & ""))
        ITM.SubItems(3) = IIf(Val(!AL_E & "") = 0, "-", Val(!AL_E & ""))
        ITM.SubItems(4) = IIf(Val(!SAL_E & "") = 0, "-", Val(!SAL_E & ""))
        ITM.SubItems(5) = IIf(Val(!LWP_E & "") = 0, "-", Val(!LWP_E & ""))
        ITM.SubItems(6) = IIf(Val(!HL_E & "") = 0, "-", Val(!HL_E & ""))
        
        Set ITM = LVBalances.ListItems(3)
        ITM.ListSubItems.Clear
        LColor = RGB(255, 188, 0)
        ITM.ListSubItems.add(, , IIf(Val(!CL & "") = 0, "-", Val(!CL & ""))).ForeColor = LColor
        ITM.ListSubItems.add(, , IIf(Val(!SL & "") = 0, "-", Val(!SL & ""))).ForeColor = LColor
        ITM.ListSubItems.add(, , IIf(Val(!AL & "") = 0, "-", Val(!AL & ""))).ForeColor = LColor
        ITM.ListSubItems.add(, , IIf(Val(!Sal & "") = 0, "-", Val(!Sal & ""))).ForeColor = LColor
        ITM.ListSubItems.add(, , IIf(Val(!LWP & "") = 0, "-", Val(!LWP & ""))).ForeColor = LColor
        ITM.ListSubItems.add(, , IIf(Val(!HL & "") = 0, "-", Val(!HL & ""))).ForeColor = LColor
        For i = 1 To ITM.ListSubItems.count
            ITM.ListSubItems(i).Bold = True
        Next
        Set ITM = LVBalances.ListItems(4)
        ITM.ListSubItems.Clear
        ITM.ListSubItems.add(, , Val(LVBalances.ListItems(1).SubItems(1)) - Val(LVBalances.ListItems(2).SubItems(1)) - Val(LVBalances.ListItems(3).SubItems(1))).ForeColor = vbBlue
        ITM.ListSubItems.add(, , Val(LVBalances.ListItems(1).SubItems(2)) - Val(LVBalances.ListItems(2).SubItems(2)) - Val(LVBalances.ListItems(3).SubItems(2))).ForeColor = vbBlue
        ITM.ListSubItems.add(, , Val(LVBalances.ListItems(1).SubItems(3)) - Val(LVBalances.ListItems(2).SubItems(3)) - Val(LVBalances.ListItems(3).SubItems(3))).ForeColor = vbBlue
        ITM.ListSubItems.add(, , Val(LVBalances.ListItems(1).SubItems(4)) - Val(LVBalances.ListItems(2).SubItems(4)) - Val(LVBalances.ListItems(3).SubItems(4))).ForeColor = vbBlue
        ITM.ListSubItems.add(, , Val(LVBalances.ListItems(1).SubItems(5)) - Val(LVBalances.ListItems(2).SubItems(5)) - Val(LVBalances.ListItems(3).SubItems(5))).ForeColor = vbBlue
        ITM.ListSubItems.add(, , Val(LVBalances.ListItems(1).SubItems(6)) - Val(LVBalances.ListItems(2).SubItems(6)) - Val(LVBalances.ListItems(3).SubItems(6))).ForeColor = vbBlue
    End With
    
End Sub

Private Sub cmbLeaveType_Change()
    Call cmbLeaveType_Click
End Sub

Private Sub cmbLeaveType_Click()

    If cmbLeaveType.ListIndex = -1 Then Exit Sub
    Dim strLeaveType As String
    strLeaveType = cmbLeaveType.List(cmbLeaveType.ListIndex, 1)
    If strLeaveType = "TML" Then        'Three Months Leave
        Dim DtFrom As Date, DtTo As Date
        
        DtFrom = DateSerial(DTPicker1.year, DTPicker1.Month, 1)
        DtTo = DateSerial(DTPicker1.year, DTPicker1.Month + 3, 0)
        DTPicker1.Value = DtFrom
        DTPicker2.Value = DtTo
        DTPicker1.Enabled = False
        DTPicker2.Enabled = False
    Else
        DTPicker1.Enabled = True
        DTPicker2.Enabled = True
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If Saved Then
        Unload Me
        frmLeaves.Show 1
    End If
End Sub

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub

Private Sub cmdPrintLeaves_Click()

    On Error GoTo err
    Dim cmd As New ADODB.Command
    With cmd
        .CommandTimeout = 120
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_EmpLeaveBalances"
        .Parameters("@DT").Value = DateSerial(DTPicker1.year, DTPicker1.Month, 1)
        .Execute
    End With
    
    Dim rpt As CRAXDDRT.Report
    If cmbEmp.MatchFound Then
        'Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpLeaveBalanceMonthlyDetail.rpt")
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpLeaveBalanceMonthly.rpt")
    Else
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpLeaveBalanceMonthly.rpt")
    End If
    rpt.FormulaFields(16).Text = "'" & cmbMonth & " - " & cmbYear & "'"
    Dim f As New frmPrevRpt
    Dim forSel As String
    
    If cmbEmp.MatchFound = False Then
        forSel = " "
    Else
        forSel = "{VEmpLeaveBalance.EmpID}='" & cmbEmp.ID & "' "
    End If
    
    Dim StartDT As Date
    Dim EndDT As Date
    
    If DTPicker1.Month < 7 Then
        StartDT = DateSerial(DTPicker1.year - 1, 7, 1)
    Else
        StartDT = DateSerial(DTPicker1.year, 7, 1)
    End If
    EndDT = DTPicker1
    
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
    
    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & strCompany & "'"
        End If
        If cmbEmp.MatchFound Then
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

Private Sub DTPicker1_Change()
    DTPicker2.Value = DTPicker1.Value
End Sub

Private Sub Form_Load()
    
    Dim ITM As ListItem
    Set ITM = LVBalances.ListItems.add(, , "Available")
    ITM.Bold = True
    ITM.ForeColor = vbRed
    If strCompany = "IAA" Or strCompany = "Kami" Then
        ITM.SubItems(1) = "12"
    Else
        ITM.SubItems(1) = "10"
    End If
    
    If strCompany = "Tecno" Then
        ITM.SubItems(2) = "-"
    Else
        ITM.SubItems(2) = "8"
    End If
    
    ITM.SubItems(3) = "14"
    ITM.SubItems(4) = "-"
    ITM.SubItems(5) = "-"
    ITM.SubItems(6) = "12"
    
    Set ITM = LVBalances.ListItems.add(, , "Used")
    ITM.Bold = True
    ITM.ForeColor = vbRed
    ITM.SubItems(1) = "-"
    ITM.SubItems(2) = "-"
    ITM.SubItems(3) = "-"
    ITM.SubItems(4) = "-"
    ITM.SubItems(5) = "-"
    ITM.SubItems(6) = "-"
    Set ITM = LVBalances.ListItems.add(, , "This Mon.")
    ITM.Bold = True
    ITM.ForeColor = vbRed
    ITM.SubItems(1) = "-"
    ITM.SubItems(2) = "-"
    ITM.SubItems(3) = "-"
    ITM.SubItems(4) = "-"
    ITM.SubItems(5) = "-"
    ITM.SubItems(6) = "-"
    Set ITM = LVBalances.ListItems.add(, , "Balance")
    ITM.Bold = True
    ITM.ForeColor = vbRed
    ITM.SubItems(1) = "-"
    ITM.SubItems(2) = "-"
    ITM.SubItems(3) = "-"
    ITM.SubItems(4) = "-"
    ITM.SubItems(5) = "-"
    ITM.SubItems(6) = "-"
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    cmbAuth.ListHeight = 600
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID"
    cmbDept.ID = "0"
    
    'cmbDept.Text = "<All Departments>"
    cmbAuth.AddVals con, "UserName", "Users", "UserName", "WHERE InActive=0"
    cmbAuth.ID = CurrentUserName
    
    DTPicker1 = Date
    DTPicker2 = Date
    DtFrom = Now
    DtTo = Now

    With cmbLeaveType
        .AddItem "Casual Leave"
        .List(0, 1) = "CL"
        If strCompany = "IAA" Or strCompany = "Kami" Then
            .List(0, 2) = "12"
        Else
            .List(0, 2) = "10"
        End If
        
        .AddItem "Sick Leave"
        .List(1, 1) = "SL"
        If strCompany = "Tecno" Then
            .List(1, 2) = "-1"
        Else
            .List(1, 2) = "8"
        End If
'        If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" OR strCompany = "Kami" Then
'            .List(1, 2) = "8"
'        Else
'            .List(1, 2) = "16"
'        End If
        
        .AddItem "Annual Leave"
        .List(2, 1) = "AL"
        .List(2, 2) = "14"
        .AddItem "Compensatory Leave"
        .List(3, 1) = "CPL"
        .List(3, 2) = "10"
        .AddItem "Unapproved Leave"
        .List(4, 1) = "LWP"
        .List(4, 2) = "0"
        .AddItem "Maternity Leave"
        .List(5, 1) = "ML"
        .List(5, 2) = "90"
        
        .AddItem "Special Approved Leave"
        .List(6, 1) = "SAL"
        .List(6, 2) = "0"
        
        If strCompany = "Tecno" Then
            .AddItem "Three Months Leave"
            .List(7, 1) = "TML"
            .List(7, 2) = "90"
        End If
        
    End With
        
End Sub

Private Sub txtAmt_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyNumber(KeyAscii.Value, True)
End Sub

Private Sub txtDeduct_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyNumber(KeyAscii.Value, True)
End Sub

Private Sub Fra1_DragDrop(Source As Control, x As Single, y As Single)

End Sub

Private Sub label1_Click(Index As Integer)
'    If Index = 1 Then
'        Dim f As New frmMiscS
'        f.ShowForm "LeaveAuthPersons", "PersonName", "Auth Person", "EntryID", "Leave Auth Persons"
'        Set f = Nothing
'        cmbAuth.ClearVals
'        cmbAuth.AddVals con, "PersonName", "LeaveAuthPersons", "EntryID", " Order By PersonName"
'    End If
End Sub

Private Sub OptFullDay_Click()
    FraShortLeave.Visible = False
    DTPicker2.Enabled = True
    chkPrintGatePass.Visible = False
End Sub

Private Sub OptHalfDay_Click()
    FraShortLeave.Visible = False
    DTPicker2.Value = DTPicker1.Value
End Sub

Private Sub OptShortLeave_Click()
    FraShortLeave.Visible = True
    DTPicker2.Value = DTPicker1.Value
    DTPicker2.Enabled = False
    chkPrintGatePass.Visible = True
End Sub

Private Sub OptType_Click(Index As Integer)
    Frame1.Enabled = OptType(1)
End Sub

Private Function DateIsValid(DT As Date) As Boolean

    If Format(DT, "ddd") = "Sun" Then Exit Function
    Dim rs As New ADODB.Recordset
    Dim bHoliday As Boolean
    
    With rs
        .Open "SELECT DT FROM Holidays WHERE DT='" & DT & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            bHoliday = False
        Else
            bHoliday = True
        End If
        .Close
    End With
    Set rs = Nothing
    
    If Not bHoliday Then
        DateIsValid = True
    End If
    
End Function

