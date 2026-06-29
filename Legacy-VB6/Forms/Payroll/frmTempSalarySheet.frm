VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmTempSalarySheet 
   BorderStyle     =   0  'None
   ClientHeight    =   7785
   ClientLeft      =   15
   ClientTop       =   195
   ClientWidth     =   12135
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleMode       =   0  'User
   ScaleWidth      =   10869.48
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3555
      Top             =   2010
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
            Picture         =   "frmTempSalarySheet.frx":0000
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTempSalarySheet.frx":0452
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkExternal 
      Caption         =   "External Salary Sheet"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9690
      TabIndex        =   8
      Top             =   375
      Visible         =   0   'False
      Width           =   2145
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   3510
      TabIndex        =   3
      Top             =   330
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   5730
      TabIndex        =   4
      Top             =   330
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin VB.Frame FLV 
      Height          =   6915
      Left            =   75
      TabIndex        =   0
      Top             =   690
      Width           =   11925
      Begin VB.Frame FButs2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   825
         Left            =   150
         TabIndex        =   17
         Top             =   5985
         Width           =   5040
         Begin VB.CommandButton Command1 
            Caption         =   "Pay All"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2355
            TabIndex        =   19
            Top             =   420
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Unpay All"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3660
            TabIndex        =   18
            Top             =   420
            Visible         =   0   'False
            Width           =   1275
         End
         Begin MSForms.CommandButton cmdUpdate 
            Height          =   360
            Left            =   45
            TabIndex        =   20
            Top             =   420
            Visible         =   0   'False
            Width           =   2265
            ForeColor       =   0
            Caption         =   "Save Changes"
            PicturePosition =   327683
            Size            =   "3995;635"
            Accelerator     =   80
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin Crystal.CrystalReport CR2 
         Left            =   3990
         Top             =   600
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
      Begin Crystal.CrystalReport cr1 
         Left            =   3285
         Top             =   600
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
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   3285
         TabIndex        =   7
         Top             =   1050
         Visible         =   0   'False
         Width           =   1155
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5820
         Left            =   105
         TabIndex        =   1
         Top             =   180
         Width           =   11730
         _ExtentX        =   20690
         _ExtentY        =   10266
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
         NumItems        =   28
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "Emp ID"
            Text            =   "Emp ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "Emp Name"
            Text            =   "Emp Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Object.Tag             =   "B. Sal"
            Text            =   "B. Sal"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Object.Tag             =   "Rate/Hr"
            Text            =   "Rate/Day"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Object.Tag             =   "A. Hrs"
            Text            =   "A.Days"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Tag             =   "A. Amt"
            Text            =   "A. Amt"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Object.Tag             =   "A.Allow"
            Text            =   "A.Allow"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Object.Tag             =   "Allow.Amt"
            Text            =   "Allow.Amt"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Object.Tag             =   "S. Hrs"
            Text            =   "S.Days"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Object.Tag             =   "S. Amt"
            Text            =   "S. Amt"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Object.Tag             =   "Leave Hrs"
            Text            =   "Leaves"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Object.Tag             =   "Lv. Amt."
            Text            =   "Lv. Amt."
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   12
            Object.Tag             =   "OT Hrs"
            Text            =   "OT Hrs"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   13
            Object.Tag             =   "OT Amt"
            Text            =   "OT Amt"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   14
            Object.Tag             =   "L.Hrs"
            Text            =   "L.Hrs"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   15
            Object.Tag             =   "L.Amt"
            Text            =   "L.Amt"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   16
            Object.Tag             =   "Total"
            Text            =   "Total"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   17
            Object.Tag             =   "T. Amt"
            Text            =   "T. Amt"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   18
            Object.Tag             =   "Net Ttl"
            Text            =   "Net Ttl"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   19
            Object.Tag             =   "ST Loan"
            Text            =   "ST Loan"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   20
            Object.Tag             =   "Loan Ded."
            Text            =   "Loan Ded."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   21
            Object.Tag             =   "Adv Sal"
            Text            =   "Adv Sal"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   22
            Object.Tag             =   "U. Fund"
            Text            =   "U. Fund"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   23
            Object.Tag             =   "Fine"
            Text            =   "Fine"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   24
            Object.Tag             =   "Scholorship"
            Text            =   "Scholorship"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   25
            Object.Tag             =   "Lunch Ded."
            Text            =   "Lunch Ded."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   26
            Object.Tag             =   "EOBI"
            Text            =   "EOBI"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   27
            Object.Tag             =   "Balance"
            Text            =   "Balance"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Frame fButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   915
         Left            =   8385
         TabIndex        =   10
         Top             =   5970
         Width           =   3450
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Total Salary Amount :"
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
            Left            =   120
            TabIndex        =   16
            Top             =   135
            Width           =   1875
         End
         Begin MSForms.CommandButton cmdCancel 
            Height          =   360
            Left            =   1995
            TabIndex        =   14
            Top             =   510
            Visible         =   0   'False
            Width           =   1425
            ForeColor       =   0
            Caption         =   " Cancel     "
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
         Begin MSForms.CommandButton cmdClose 
            Height          =   360
            Left            =   1995
            TabIndex        =   13
            Top             =   510
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
         Begin MSForms.CommandButton cmdPrint 
            Height          =   360
            Left            =   450
            TabIndex        =   12
            Top             =   510
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
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00 "
            Height          =   285
            Left            =   2025
            TabIndex        =   11
            Top             =   75
            Width           =   1425
         End
         Begin MSForms.CommandButton cmdSave 
            Height          =   360
            Left            =   450
            TabIndex        =   15
            Top             =   510
            Visible         =   0   'False
            Width           =   1425
            ForeColor       =   0
            Caption         =   "  Save       "
            PicturePosition =   327683
            Size            =   "2514;635"
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
   End
   Begin VB.TextBox txtTotalHrs 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2535
      TabIndex        =   22
      Top             =   330
      Width           =   960
   End
   Begin VB.TextBox txtTotal 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2535
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   330
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Year                  Month                Total Days     Department                         Employee"
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
      Left            =   60
      TabIndex        =   21
      Top             =   90
      Width           =   7920
   End
   Begin MSForms.ComboBox cmbMonth 
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Top             =   330
      Width           =   1350
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2381;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbYear 
      Height          =   285
      Left            =   60
      TabIndex        =   5
      Top             =   330
      Width           =   1140
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2011;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   8025
      TabIndex        =   2
      Top             =   270
      Width           =   1620
      Caption         =   "Refresh        "
      PicturePosition =   327683
      Size            =   "2857;635"
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
      Begin VB.Menu mnuSalSheet 
         Caption         =   "Print Salary Sheet (Detail)"
      End
      Begin VB.Menu PrintSalsummary 
         Caption         =   "Print Salary Sheet (Summary)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSalarysheetInternal 
         Caption         =   "Print Salary Sheet (Internal)"
      End
      Begin VB.Menu ln2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLateEmpList 
         Caption         =   "Print Employee List With Late Coming Deduction"
      End
      Begin VB.Menu mnuCurrency 
         Caption         =   "Print Currency Denomination"
         Visible         =   0   'False
      End
      Begin VB.Menu dash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPayslips 
         Caption         =   "Print PaySlips"
      End
      Begin VB.Menu mnuPayslipsInternal 
         Caption         =   "Print Payslips (Internal)"
      End
      Begin VB.Menu mnuPrintOTPaySheet 
         Caption         =   "Print Over Time Payment Sheet"
      End
   End
End
Attribute VB_Name = "frmTempSalarySheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

''''''''''Currency Denomination'''''''''''''
    Dim Thousands As Double, FHundreds As Double, Hundreds As Double, Fifties As Double, Tens As Double, Fives As Double, TotalSal As Double, Twos As Double, Ones As Double
'******************************************

Dim EditAbsents As Boolean
Dim MakeFinal As Integer
Dim rsLedger As New ADODB.Recordset
Dim MonthDays As Long
Dim IsInternalFinal As Boolean


Private Sub chkExternal_Click()
    
    With LV.ColumnHeaders
        For i = 1 To .Count
            .Item(i).Text = .Item(i).Tag
        Next
    End With
    
    If chkExternal Then
        txtTotalHrs.Visible = True
        txtTotal.Visible = False
    Else
        txtTotalHrs.Visible = False
        txtTotal.Visible = True
    End If
    
    Call cmbMonth_Click
    
    If Me.Visible Then Call cmdRefresh_Click
    
    
End Sub


Private Sub cmbMonth_Change()
    If cmbMonth.MatchFound Then Call cmbMonth_Click
End Sub

Private Sub cmbMonth_Click()

    txtTotal = TotalMonthDays
    
    Dim MonthDays As Long
    MonthDays = txtTotal - TotalSundays(cmbMonth.List(cmbMonth.ListIndex, 1), Val(cmbYear), Val(txtTotal))
    
    If chkExternal = vbChecked Then
        Hdays = Val(con.Execute("Select count(DT) from Holidays Where month(DT)=" & _
        cmbMonth.List(cmbMonth.ListIndex, 1) & " and Year(DT)=" & Val(cmbYear)).Fields(0))
        MonthDays = MonthDays '- Hdays
        txtTotalHrs = MonthDays * 8
    End If
    Call AddEmployees
    
End Sub

Private Sub cmbYear_Change()
    If cmbYear.MatchFound Then Call cmbMonth_Click
    txtTotal = TotalMonthDays
End Sub

Private Sub cmdCancel_Click()
    cmdRefresh.Visible = True
    cmdClose.Visible = True
    cmdPrint.Visible = True
    cmdSave.Visible = False
    cmdCancel.Visible = False
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub PrintReport(PaySlips As Boolean, Optional Crpt As CrystalReport = Nothing, Optional Summary As Boolean)


If LV.ListItems.Count = 0 Then Exit Sub
Call InsertRecords

Dim unloadflag As Boolean
'Sleep (1000)
'Dim TDate As Date
'Dim MonthDays As Integer
'TDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
'MonthDays = DateAdd("m", 1, TDate) - TDate

'With Crpt
'
'    If Not PaySlips Then
'        If Summary Then
'          .ReportFileName = RptPath & "\SalarySummary.rpt"
'          .Formulas(1) = ""
'          .Formulas(1) = "Company='" & company & "'"
'        Else
'          .ReportFileName = RptPath & "\SalarySheet.rpt"
'          .Formulas(1) = ""
'          .Formulas(1) = "Company='" & company & "'"
'        End If
'      .Formulas(0) = "Fromto='For " & cmbMonth.Tag & "-" & cmbYear.Tag & IIf(chkExternal = vbChecked, "", " ( For Factory Use Only )") & "'"
'      .Formulas(1) = "Company='" & company & "'"
'    Else
'        .ReportFileName = RptPath & "\PaySlip.rpt"
'
'        .Formulas(0) = "ForMonth='For " & cmbMonth.Tag & "-" & cmbYear.Tag & "'"
'        .Formulas(1) = ""
'    End If
'
'    .Action = 1
'    .PageZoomNext
'End With


Screen.MousePointer = vbHourglass

Dim f As New frmPrevRpt, rpt  As craxddrt.Report

If PaySlips Then
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptPayslip.rpt")
Else
    If Summary Then
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptSalarySummary.rpt")
    Else
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptSalarySheet.rpt")
    End If
End If

Dim FormulaFields As craxddrt.FormulaFieldDefinitions
Dim FormulaField As craxddrt.FormulaFieldDefinition

Set FormulaFields = rpt.FormulaFields
For Each FormulaField In FormulaFields
    If FormulaField.Name = "{@Company}" Then
        FormulaField.Text = "'" & company & "'"
    ElseIf FormulaField.Name = "{@FromTo}" Then
        FormulaField.Text = "' For " & cmbMonth.Tag & "-" & cmbYear.Tag & "'"
    ElseIf FormulaField.Name = "{@MonthDays}" Then
        FormulaField.Text = Val(txtTotalHrs)
    End If
Next
Cond = ""
rpt.SQLQueryString = Sql
f.ShowReport Cond & " ", rpt


Screen.MousePointer = vbDefault




End Sub

Private Sub PrintInternalReport(PaySlips As Boolean, Optional Crpt As CrystalReport = Nothing)


If LV.ListItems.Count = 0 Then Exit Sub

Call InsertRecords

Dim unloadflag As Boolean
'Sleep (1000)
'Dim TDate As Date
'Dim MonthDays As Integer
'TDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
'MonthDays = DateAdd("m", 1, TDate) - TDate

'With Crpt
'
'    If Not PaySlips Then
'        If Summary Then
'          .ReportFileName = RptPath & "\SalarySummary.rpt"
'          .Formulas(1) = ""
'          .Formulas(1) = "Company='" & company & "'"
'        Else
'          .ReportFileName = RptPath & "\SalarySheet.rpt"
'          .Formulas(1) = ""
'          .Formulas(1) = "Company='" & company & "'"
'        End If
'      .Formulas(0) = "Fromto='For " & cmbMonth.Tag & "-" & cmbYear.Tag & IIf(chkExternal = vbChecked, "", " ( For Factory Use Only )") & "'"
'      .Formulas(1) = "Company='" & company & "'"
'    Else
'        .ReportFileName = RptPath & "\PaySlip.rpt"
'
'        .Formulas(0) = "ForMonth='For " & cmbMonth.Tag & "-" & cmbYear.Tag & "'"
'        .Formulas(1) = ""
'    End If
'
'    .Action = 1
'    .PageZoomNext
'End With


Screen.MousePointer = vbHourglass

Dim f As New frmPrevRpt, rpt  As New craxddrt.Report

If PaySlips Then
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptPayslipInternal.rpt")
Else
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptSalarySheetInternal.rpt")
End If

Dim FormulaFields As craxddrt.FormulaFieldDefinitions
Dim FormulaField As craxddrt.FormulaFieldDefinition

Set FormulaFields = rpt.FormulaFields
For Each FormulaField In FormulaFields
    If FormulaField.Name = "{@Company}" Then
        FormulaField.Text = "'" & company & "'"
    ElseIf FormulaField.Name = "{@FromTo}" Then
        FormulaField.Text = "' For " & cmbMonth.Tag & "-" & cmbYear.Tag & "'"
    ElseIf FormulaField.Name = "{@MonthDays}" Then
        'FormulaField.Text = MonthDays
        FormulaField.Text = Val(txtTotal)
    End If
Next
Cond = ""
rpt.SQLQueryString = Sql
f.ShowReport Cond & " ", rpt


Screen.MousePointer = vbDefault




End Sub



Private Sub cmdPrint_Click()


    mnuSalSheet.Visible = (chkExternal.value = vbChecked)
    mnuPayslips.Visible = (chkExternal.value = vbChecked)
    
    mnuSalarysheetInternal.Visible = (chkExternal.value = vbUnchecked)
    mnuPayslipsInternal.Visible = (chkExternal.value = vbUnchecked)
    
    Me.PopUpMenu mnupop, 2, cmdPrint.Left + FLV.Left + fButs1.Left, fButs1.Top + cmdPrint.Top + FLV.Top + cmdPrint.Height

End Sub

Public Sub cmdRefresh_Click()

    Call ShowWait
    cmbMonth.Tag = cmbMonth
    cmbYear.Tag = cmbYear
    Call RefreshLedger
    
    MakeFinal = 0
    
    Command1.Visible = LV.Checkboxes
    Command2.Visible = LV.Checkboxes
 
    Call HideWait
    
End Sub



Private Sub cmdSave_Click()
    
    cmdRefresh.Visible = True
    cmdClose.Visible = True
    cmdPrint.Visible = True
    cmdSave.Visible = False
    cmdCancel.Visible = False
    
End Sub

Private Sub cmdUpdate_Click()
    
    If cmdUpdate.Caption = "Make This Sheet Final" Then
    
'        If chkExternal.value = vbChecked Then
'           If IsInternalFinal = False Then
'                MsgBox "You Must Finalize The Internal Salary First.", vbInformation
'                Exit Sub
'            End If
'        End If
        If getPasswordNew("wasiq") = False Then
            Exit Sub
        End If
        
        With rsLedger
            If .RecordCount = 0 Then Exit Sub Else .MoveFirst
   
            con.BeginTrans
            Dim i As Integer
            'For i = 1 To LV.ListItems.Count
                '   .Find "EMPID='" & LV.ListItems(i) & "'", 0, adSearchForward, 1
                Call SaveIt
            '.MoveNext
            'Next
            If AddSalaryVouchers Then
                con.CommitTrans
                Call cmdRefresh_Click
            Else
                MsgBox "Could Not Add Vouchers.", vbInformation
                Exit Sub
            End If
        End With
        
        MsgBox "Salary Sheet Saved Successfully", vbInformation
        
        cmdUpdate.Visible = False
        'LV.Checkboxes = True
        'Command1.Visible = True
        'Command2.Visible = True
        'cmdUpdate.Caption = "Save Changes"
    Else
   
        Dim setVal As Integer
        Dim EmpId As String
        Dim TheMonth As Byte
        Dim TheYear As Integer
        
'        con.BeginTrans
'
'        For i = 1 To LV.ListItems.Count
'            setVal = Abs(LV.ListItems(i).Checked)
'            EmpId = Left(LV.ListItems(i).key, 7)
'            TheMonth = Month(CDate(Right(LV.ListItems(i).key, 11)))
'            TheYear = year(CDate(Right(LV.ListItems(i).key, 11)))
'
'            con.Execute "Update MonthlySalaries Set Payed=" & setVal & ",NetSalary=" & Val(LV.ListItems(i).SubItems(20)) + Val(LV.ListItems(i).SubItems(18)) & _
'            ",OTAmt=" & Val(LV.ListItems(i).SubItems(9)) & ",StAmt=" & Val(LV.ListItems(i).SubItems(16)) & _
'            ",LTAmt=" & Val(LV.ListItems(i).SubItems(17)) & ",LateAmt=" & Val(LV.ListItems(i).SubItems(11)) & _
'            ",salAmt=" & Val(LV.ListItems(i).SubItems(7)) & _
'            " Where EmpID='" & EmpId & "' and month(dt)=" & TheMonth & " and Year(dt)=" & TheYear & " and SASalary=" & chkExternal
'
'            AmtClrd = Val(LV.ListItems(i).SubItems(20))
'
'            con.Execute "Update AmtCleared Set AmtClrd=" & AmtClrd & " Where EmpID='" & EmpId & "' and month(dt)=" & TheMonth & " and Year(dt)=" & TheYear
'
'            If LV.ListItems(i).Checked Then
'                con.Execute "Update Advances Set amountcleared=amount Where EmpId='" & EmpId & "' and type =0 and month(dt)=" & TheMonth & " and Year(dt)=" & TheYear
'            Else
'                con.Execute "Update Advances Set amountcleared=0 Where EmpId='" & EmpId & "' and type =0 and month(dt)=" & TheMonth & " and Year(dt)=" & TheYear
'            End If
'
'        Next i
'
'        con.CommitTrans

End If
       
End Sub

Private Function AddSalaryVouchers()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim i As Long, INextSno As Double, Bal As Double, Amt As Double
    Dim VchrNo As String, TheDesc As String, AccNo As String
    Dim Dated As Date
    Dim dDRTaxDiff As Double
    
    With rs
        .Open "Select Departments.Name,Departments.AccruedAccNo,Departments.SalaryAccNo,Departments.AccruedFoodAccNo,Departments.AccruedUnionAccNo,Departments.SchAccNo,Departments.AccruedEOBIAccNo,Departments.TaxAccNo, " & _
                "T1.DeptID,T1.TotalAmt,T1.SchAmt,T1.LunchAmt,T1.UnionAmt,T1.BalanceAmt,T1.ShortAmt,T1.LongAmt,T1.EOBIAmt,T1.TaxAmt " & _
                "From (Select Left(EmpID,4) As DeptID,Sum(SAmt) As TotalAmt,Sum(Bonus) As SchAmt,Sum(Lunch) As LunchAmt,Sum(UnionFund) As UnionAmt,Sum(Balance) As BalanceAmt,Sum(ShortTerm) As ShortAmt,Sum(LongTerm) As LongAmt,Sum(EOBI) As EOBIAmt,Sum(Tax) As TaxAmt From MonthlySalaries " & _
                "Where Month(DT)=" & cmbMonth.ListIndex + 1 & " and year(DT)=" & cmbYear & " " & _
                "Group By Left(EmpID,4)) T1 Inner JOIN Departments ON T1.DeptID=Departments.DeptID " & _
                "Where IsNull(Departments.AccruedAccNo,'')<>'' AND IsNull(Departments.SalaryAccNo,'')<>'' AND IsNull(Departments.AccruedFoodAccNo,'')<>'' AND IsNull(Departments.AccruedUnionAccNo,'')<>''  AND IsNull(Departments.SchAccNo,'')<>'' AND IsNull(Departments.AccruedEOBIAccNo,'')<>'' AND IsNull(Departments.TaxAccNo,'')<>''", con, adOpenForwardOnly, adLockReadOnly
                
        Do Until .EOF
        
            Dated = DateSerial(cmbYear, cmbMonth.ListIndex + 2, 0) 'Date
            
            If Format(Dated, "ddd") = "Sun" Then
                Dated = DateAdd("d", -1, Dated)
            End If
            
            INextSno = getNextSno(Dated)
            VchrNo = getNextVchrNo(Dated, "JV")
            
            '''''''''''''''''' T A X   E N T R I E S '''''''''''''
            If Val(!TaxAmt & "") > 0 Then
                TheDesc = "Tax Of [" & !DeptID & "] " & !Name
                AccNo = !TaxAccNo & ""
                Bal = getBalance(AccNo, Dated)
                Amt = Val(!TaxAmt & "")
                
                If Amt <> 0 Then
                    Bal = Bal - Amt
                    
                    con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
                     Dated & "','" & VchrNo & "','" & AccNo & "','" & TheDesc & _
                     "',0," & Amt & "," & Bal & ",'',0)"
                     
                    Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
                    INextSno = INextSno + 1
                End If
            End If
'            If !DeptID = "VDRS" Then
'                TheDesc = "Tax Of Director Salaries"
'
'                '''''''''''''''''' T A X   E N T R I E S '''''''''''''
'
'                AccNo = '"45-04001"
'                Bal = getBalance(AccNo, Dated)
'                Amt = Val(!TaxAmt & "")
'
'                If Amt <> 0 Then
'
'                    Bal = Bal - Amt
'
'                    con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
'                       Dated & "','" & VchrNo & "','" & AccNo & _
'                       "','" & TheDesc & "',0," & Amt & "," & Bal & ",'',0)"
'
'                    Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
'                    INextSno = INextSno + 1
'                End If
'
'            End If
            
            'Charge All Amount to Cash
            Amt = Val(!TotalAmt & "")
    
            If Amt = 0 Then Exit Function
            'Give The Description
            TheDesc = "Salary Of  [" & !DeptID & "] " & !Name
            
            '''''''''''''''''' D E B I T   E N T R I E S '''''''''''''
            '''''''''''''''''' Accrued Salary Account '''''''''''''
            AccNo = !SalaryAccNo
            Amt = Val(!TotalAmt & "")
            Bal = getBalance(AccNo, Dated)
      
            Amt = Amt
            Bal = Bal + Amt
            
            con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
                Dated & "','" & VchrNo & "','" & AccNo & _
                "','" & TheDesc & "'," & Amt & ",0," & Bal & ",'',0)"
        
            Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
            INextSno = INextSno + 1
            '''''''''''''''''' Scholorship Account '''''''''''''
            
            AccNo = !SchAccNo
            Amt = Val(!SchAmt & "")
            If Amt <> 0 Then
              Bal = getBalance(AccNo, Dated)
        
              Amt = Amt
              Bal = Bal + Amt
              
              con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
                  Dated & "','" & VchrNo & "','" & AccNo & _
                  "','" & TheDesc & "'," & Amt & ",0," & Bal & ",'',0)"
          
              Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
              INextSno = INextSno + 1
            End If
            '''''''''''''''''' C R E D I T   E N T R I E S '''''''''''''
            ''''''''''''''''Temp. Advance'''''''''''''
            
            AccNo = "15-015-20001"
            Amt = Val(!ShortAmt & "")
            If Amt <> 0 Then
                Bal = getBalance(AccNo, Dated)
                
                Bal = Bal - Amt
            
                con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
                    Dated & "','" & VchrNo & "','" & AccNo & _
                    "','" & TheDesc & "',0," & Amt & "," & Bal & ",'',0)"
        
                Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
         
                INextSno = INextSno + 1
            End If
            ''''''''''''''''Long Term Deduction'''''''''''''
            
            AccNo = "13-003-16001"
            Amt = Val(!LongAmt & "")
            If Amt <> 0 Then
                Bal = getBalance(AccNo, Dated)
    
                Bal = Bal - Amt
    
                con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
                    Dated & "','" & VchrNo & "','" & AccNo & _
                    "','" & TheDesc & "',0," & Amt & "," & Bal & ",'',0)"
    
                Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
    
                INextSno = INextSno + 1
            End If
            ''''''''''''''''Food Amount Deduction'''''''''''''
            
            AccNo = !AccruedFoodAccNo & ""
            Amt = Val(!LunchAmt & "")
            If Amt <> 0 Then
                Bal = getBalance(AccNo, Dated)
                
                Bal = Bal - Amt
            
                con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
                    Dated & "','" & VchrNo & "','" & AccNo & _
                    "','" & TheDesc & "',0," & Amt & "," & Bal & ",'',0)"
        
                Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
         
                INextSno = INextSno + 1
            End If
            ''''''''''''''''Food Amount Deduction'''''''''''''
            
            AccNo = !AccruedUnionAccNo & ""
            Amt = Val(!UnionAmt & "")
            If Amt <> 0 Then
                Bal = getBalance(AccNo, Dated)
                
                Bal = Bal - Amt
            
                con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
                    Dated & "','" & VchrNo & "','" & AccNo & _
                    "','" & TheDesc & "',0," & Amt & "," & Bal & ",'',0)"
        
                Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
         
                INextSno = INextSno + 1
            End If
            ''''''''''''''''Balance Amount Deduction'''''''''''''
            'If There's Ded. Of Tax. Then Adjust The Difference Of Rounding Of Tax Amount.
            
            AccNo = !AccruedAccNo & ""
            Amt = Val(!BalanceAmt & "")
            

            dDRTaxDiff = Val(!TotalAmt & "") + Val(!SchAmt & "") - Amt - Val(!LunchAmt & "") - Val(!UnionAmt & "") - Val(!ShortAmt & "") - Val(!LongAmt & "") - Val(!EOBIAmt & "") - Val(!TaxAmt & "")
            Amt = Amt + dDRTaxDiff

            
'            If !DeptID = "VDRS" Then
'                dDRTaxDiff = Val(!TotalAmt & "") - Amt - Val(!TaxAmt & "")
'                Amt = Amt + dDRTaxDiff
'            Else
'                dDRTaxDiff = Val(!TotalAmt & "") + Val(!SchAmt & "") - Amt - Val(!LunchAmt & "") - Val(!UnionAmt & "") - Val(!ShortAmt & "") - Val(!LongAmt & "") - Val(!EOBIAmt & "")
'                Amt = Amt + dDRTaxDiff
'            End If
            
            Bal = getBalance(AccNo, Dated)
            
            Bal = Bal - Amt
        
            con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
                Dated & "','" & VchrNo & "','" & AccNo & _
                "','" & TheDesc & "',0," & Amt & "," & Bal & ",'',0)"
    
            Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
     
            INextSno = INextSno + 1
            
            ''''''''''''''''EOBI Amount Deduction'''''''''''''
            
            AccNo = !AccruedEOBIAccNo & ""
            Amt = Val(!EOBIAmt & "")
            If Amt <> 0 Then
                Bal = getBalance(AccNo, Dated)
                
                Bal = Bal - Amt
            
                con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
                    Dated & "','" & VchrNo & "','" & AccNo & _
                    "','" & TheDesc & "',0," & Amt & "," & Bal & ",'',0)"
        
                Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
         
                INextSno = INextSno + 1
            End If
            
            con.Execute "Insert Into VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & VchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate & "')"
                
            .MoveNext
        Loop
        .Close
        
'        .Open "Select Sum(Tax) From MonthlySalaries Where Month(DT)=" & cmbMonth.ListIndex + 1 & " and year(DT)=" & cmbYear & " AND  Left(EmpID,4)='VDRS'", con, adOpenForwardOnly, adLockReadOnly
'
'        If Val(.Fields(0) & "") > 0 Then
'
'            Dated = DateAdd("d", 7, Dated)
'            If Format(Dated, "ddd") = "Sun" Then
'                Dated = DateAdd("d", 1, Dated)
'            End If
'
'            Amt = Val(.Fields(0) & "")
'
'            INextSno = getNextSno(Dated)
'            VchrNo = getNextVchrNo(Dated, "JV")
'
'            'Give The Description
'            TheDesc = "Tax Of Director Salaries"
'
'            '''''''''''''''''' T A X   E N T R I E S '''''''''''''
'
'            AccNo = "45-4001"
'            Bal = getBalance(AccNo, Dated)
'            Amt = Amt
'            Bal = Bal + Amt
'
'            con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
'                Dated & "','" & VchrNo & "','" & AccNo & _
'                "','" & TheDesc & "'," & Amt & ",0," & Bal & ",'',0)"
'
'            Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
'            INextSno = INextSno + 1
'            ''''''''''''''' C R E D I T     E N T R Y '''''''''''''''''
'            AccNo = CashInHand
'
'            Bal = getBalance(AccNo, Dated)
'
'            Bal = Bal - Amt
'
'            con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
'                Dated & "','" & VchrNo & "','" & AccNo & _
'                "','" & TheDesc & "',0," & Amt & "," & Bal & ",'',0)"
'
'            Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
'
'            con.Execute "Insert Into VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & VchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate & "')"
'        End If
'        .Close
    End With
    
    Set rs = Nothing
    AddSalaryVouchers = True
    
    Exit Function
err:
    MsgBox err.Description
    
End Function

Private Sub Command1_Click()
    For i = 1 To LV.ListItems.Count
        LV.ListItems(i).Checked = True
    Next
End Sub

Private Sub Command2_Click()
For i = 1 To LV.ListItems.Count
    LV.ListItems(i).Checked = False
Next
End Sub

Private Sub Form_Load()
    
    mnupop.Visible = False

    txtTotalHrs.Visible = False
    txtTotal.Visible = True

    'LV.ColumnHeaders(14).Width = 0
    'LV.ColumnHeaders(15).Width = 0

    DtTo = GetServerDate
    DtFrom = DateAdd("M", -1, GetServerDate)
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    
    
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
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID", " Where IsNull(TempDept,0)=0"
    cmbDept.Text = "<All Departments>"
    
    cmbEmp.Text = "<All Employees>"
    'Set LV.SmallIcons = ImageList1
    
    cmbMonth.ListIndex = Val(Format$(GetServerDate, "MM")) - 1
    cmbYear = Format$(GetServerDate, "yyyy")
    
 
    
    'Call cmdRefresh_Click
    
End Sub
Private Sub cmbdept_matched()
    
'    cmbEmp.ClearVals
'    cmbEmp.AddItem "<All Employees>", "0"
'    If cmbDept.Text = "<All Departments>" Then
'        cmbEmp.AddVals con, "Name", "Employees", "EmpID", "Where Active=1"
'    Else
'        cmbEmp.AddVals con, "Name", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "' and Active=1"
'    End If
'    cmbEmp.ID = 0
    Call AddEmployees
    
End Sub

Private Sub AddEmployees()
    
    cmbEmp.ClearVals
    If cmbYear.MatchFound = False Then
        'MsgBox "Select Year.", vbInformation
        Exit Sub
    End If

    If cmbMonth.MatchFound = False Then
        'MsgBox "Select Year.", vbInformation
        Exit Sub
    End If
    
    'cmbEmp.AddItem "<All Employees>", "0"
    Dim LastDT As Date
    Dim FirstDT As Date
    
    FirstDT = DateSerial(cmbYear, cmbMonth.ListIndex + 1, 1)
    LastDT = DateSerial(cmbYear, cmbMonth.ListIndex + 2, 0)
    
    cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.Text = "<All Departments>" Then
        cmbEmp.AddVals con, "Name", "VSalariedEmps", "EmpID", " Where JOINDate<='" & LastDT & "' AND (DiscontinuedOn>'" & FirstDT & "' OR DiscontinuedOn Is Null)"
    Else
        cmbEmp.AddVals con, "Name", "VSalariedEmps", "EmpID", "Where DeptID='" & cmbDept.ID & "' and JOINDate<='" & LastDT & "' AND (DiscontinuedOn>'" & FirstDT & "' OR DiscontinuedOn Is Null)"
    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    Dim TopMargin As Integer
    TopMargin = cmbEmp.Top + cmbEmp.Height + 30
    
    FLV.Move 50, TopMargin, Me.ScaleWidth - 100, Me.ScaleHeight - TopMargin
    
    LV.Move 100, 200, Me.Width - 400, FLV.Height - fButs1.Height - 300
    
    fButs1.Move LV.Left + LV.Width - fButs1.Width, LV.Top + LV.Height
    
    FButs2.Move LV.Left, LV.Top + LV.Height
 
End Sub



Private Sub RefreshLedger()
    
    Dim rs As New ADODB.Recordset
    Dim rsInternal As New ADODB.Recordset
    With rs
        .Open "Select Count(EmpID) From MonthlySalaries Where Month(DT)=" & cmbMonth.ListIndex + 1 & " AND Year(DT)=" & cmbYear.Text & " AND IsForSA=0", con, adOpenForwardOnly, adLockReadOnly
        If Val(.Fields(0) & "") = 0 Then
            IsInternalFinal = False
        Else
            IsInternalFinal = True
            
            'To Save Long term Loan Information....
            rsInternal.Open "Select EmpID,LongTerm,PrevLTLoan From MonthlySalaries Where Month(DT)=" & cmbMonth.ListIndex + 1 & " AND Year(DT)=" & cmbYear.Text & " AND IsForSA=0", con, adOpenStatic, adLockReadOnly
            
        End If
        .Close
    End With
    
    Dim TotalSalDays As Double
    Dim TDate As Date
    Dim SDate As Date 'For Saving
    Dim Hdays As Long
    
    SDate = DateSerial(cmbYear, cmbMonth.ListIndex + 1, 1)
    TDate = DateSerial(cmbYear, cmbMonth.ListIndex + 1, Day(GetServerDate))
    
    TotalSalDays = DateAdd("m", 1, SDate) - SDate

    MonthDays = TotalSalDays - TotalSundays(cmbMonth.List(cmbMonth.ListIndex, 1), cmbYear, TotalSalDays)
    
    If chkExternal = vbUnchecked Then
        Hdays = Val(con.Execute("Select count(DT) from Holidays Where month(DT)=" & _
        cmbMonth.List(cmbMonth.ListIndex, 1) & " and Year(DT)=" & Val(cmbYear)).Fields(0))
        MonthDays = MonthDays - Hdays
    End If
   
    If FinalizedSalarySheetLoaded Then Exit Sub
    
    txtTotalHrs.Tag = txtTotalHrs


    Dim UnionFund As Integer
    Dim GrandTotal As Double
    Dim EOBIAmt As Byte
    If cmbDept.MatchFound = False Or cmbEmp.MatchFound = False Then Exit Sub
    If cmbYear.MatchFound = False Or cmbMonth.MatchFound = False Then Exit Sub
    
    Dim SunDates() As Date
        
    Dim EmpId As String
    Dim EmpAge As Double
    
    Dim TempTotalSaldays As Double
    Dim LastDay As Integer
    
    Dim ITM As ListItem
    Dim Cond As String
    Dim Sql As String
    Dim cmd As New ADODB.command
    
    
    Dim SalHrs As Double
    Dim HrsMultiple As Double
    
    
    Dim Leaves As Double
    Dim ApprovedLeaves As Double
    Dim UnApprovedLeaves As Double
    
    If cmbEmp.ID = "0" Then
       EmpId = IIf(cmbDept.ID = "0", "", cmbDept.ID)
    Else
       EmpId = IIf(cmbEmp.ID = "0", "", cmbEmp.ID)
    End If
    
    Dim rsCheck As New ADODB.Recordset
        
    Dim LunchDays As Long
    
    
    Dim OTByQty As Boolean
    
    
    'TDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), Day(GetServerDate))
        
   
    ''''''''''''''''''''' This Will Refresh The LateHours '''''''
    '            con.Execute "EXEC SP_CalcOverTime '" & SDate & "'", "'" & cmbEmp.ID & "'"
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        
        
    LastDay = TotalSalDays
    
    Dim JDate As Date
    
    If TDate = GetServerDate Then
        TotalSalDays = Day(GetServerDate)
    ElseIf GetServerDate < TDate Then
        MsgBox "Can't Show Entries Of Future Months.", vbInformation
        Exit Sub
    End If
    
    TempTotalSaldays = TotalSalDays     'Store For Later Use
    
    JDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), TotalSalDays)
        
    'Sql = "Select * from VAdvanceSalary " & Cond & " JoinDate<='" & Format(JDate, "dd-MMM-yyyy") & "' Order By DeptID,EmpID"
    
    cmd.CommandType = adCmdStoredProc
    
    cmd.CommandText = "VAdvanceSalary"
    
    Set cmd.ActiveConnection = con
    cmd.Parameters.Refresh
    
    cmd.Parameters(1).Type = adDBTimeStamp
    cmd.Parameters(1).value = JDate

    cmd.Parameters(2).Type = adVarChar
    cmd.Parameters(2).value = EmpId
        
    If rsLedger.State = 1 Then rsLedger.Close
    rsLedger.CursorType = adOpenForwardOnly
    rsLedger.LockType = adLockReadOnly
    
    Set rsLedger = cmd.Execute
    
    With rsLedger
        
        '.Open Sql, Con, adOpenForwardOnly, adLockReadOnly
        Dim TotalSuns As Integer
        Dim SalDays As Double
        Dim Rate As Double
        Dim AbsAllowed As Integer
        Dim AbsentDays As Double
        
        Dim Total As Double
        Dim OverTimeHours As Double
        Dim OverTimeHoursSA As Double
        Dim LateHours As Double
        Dim TaxPer As Double
        Dim DedAmt As Double
        Dim Payed As Double
        Dim LongLoan As Double
        Dim TaxSal As Double
        Dim IsComputed As Boolean
        
        Dim NTime As Integer
        Dim LTime As Double
        Dim DT As String
        Dim IsPayed As Boolean
        Dim IsTransfered As Boolean
        Dim Holidays As Integer
        
        Dim EmpSalary As Double
        Dim EmpExtSal As Double
                
        LV.ListItems.Clear
        LV.Visible = False
        
        For i = 0 To .RecordCount - 1
                    
            'Reset Total Salary days For Each Employee
            
            TotalSalDays = DateAdd("m", 1, SDate) - SDate
                
            If TDate = GetServerDate Then
                TotalSalDays = Day(GetServerDate)
            End If
            
            'Check that Employee Discontinued In This Month
                        
            If Not IsNull(![DiscontinuedOn]) Then
                If Month(![DiscontinuedOn]) = Month(JDate) And ![DiscontinuedOn] <= JDate Then
                  TotalSalDays = Day(![DiscontinuedOn])
                End If
            Else
               TotalSalDays = TempTotalSaldays
            End If
            
            Leaves = Val(!Leaves & "")
                                   

           'if Joined This Month Calculate From Join Date
           
           If Format(![JoinDate], "MMM-yyyy") = Format(JDate, "MMM-yyyy") Then
               
               rs.Open "Select count(DT) from Holidays Where DT Between '" & ![JoinDate] & "' and '" & JDate & "'", con, adOpenForwardOnly, adLockReadOnly
               
                  If rs.RecordCount > 0 Then
                     Holidays = rs.Fields(0) & ""
                  End If
               rs.Close
               Set rs = Nothing
               
               SundayAbsents = getSundayAbsents(![EmpId] & "", Val(cmbMonth.List(cmbMonth.ListIndex, 1)), cmbYear, ![WeeklyAbs], Int(TotalSalDays))
               TotalSuns = TotalSundays(cmbMonth.List(cmbMonth.ListIndex, 1), cmbYear, TotalSalDays, Day(![JoinDate]))
               Holidays = TotalSuns + Holidays - SundayAbsents
               
            Else
            
                 rs.Open "Select count(DT) from Holidays Where convert(char(3),DT,7)='" & cmbMonth & "' and Year(DT)=" & cmbYear & " and day(dt)<" & TotalSalDays, con, adOpenForwardOnly, adLockReadOnly
                  
                  If rs.RecordCount > 0 Then
                     Holidays = rs.Fields(0) & ""
                  End If
                  
                 rs.Close
               Set rs = Nothing

               'SundayAbsents = getSundayAbsents(![EmpId], cmbMonth.List(cmbMonth.ListIndex, 1), cmbYear, Int(TotalSalDays), ![WeeklyAbs])
               
               TotalSuns = TotalSundays(cmbMonth.List(cmbMonth.ListIndex, 1), cmbYear, TotalSalDays)
                                              
               Holidays = TotalSuns + Holidays - SundayAbsents
               

            End If
                    
            If chkExternal = vbChecked Then
                'Exclude The Holidays From Calculation
                TotalSalDays = TotalSalDays - Holidays
            End If
            
            
            'Get Sundays Not To Be Paid

            Call Computed(![EmpId], IsComputed, TaxPer, AbsAllowed, OverTimeHours, NTime, LTime, LongLoan, DT, IsPayed, IsTransfered, JDate, EmpSalary, EmpExtSal, ![Active], HrsMultiple)
            
            If Not IsComputed Then
                'Tax %
                'If IsNull(![PerDeduct]) Then TaxPer = 0 Else TaxPer = Val(![PerDeduct]) & ""
                
                'Absent Days....
                AbsAllowed = Val(![AbsAllowed] & "")
    
                ' Reset Orver Time Hours
                OverTimeHours = 0
                OverTimeHoursSA = 0
                
                'Late Hours
                NTime = Val(![NTimes] & "")
                LTime = Val(![TimeAllowed] & "")
                
                'Long Term Loan Deduction
                Dim lDedAmt As Long, lClsAmt As Long, lTotAmt As Long
                If IsNull(![DedAmt]) Then lDedAmt = 0 Else lDedAmt = ![DedAmt]
                If IsNull(![TotAmt]) Then lTotAmt = 0 Else lTotAmt = ![TotAmt]
                If IsNull(![clsamt]) Then lClsAmt = 0 Else lClsAmt = ![clsamt]
                
                If lDedAmt > (lTotAmt - lClsAmt) Then
                    LongLoan = (lTotAmt - lClsAmt)
                Else
                    LongLoan = lDedAmt
                End If
                                      
                'Get The Salary Depending On Salary Sheet Type
                'EmpSalary = IIf(chkExternal = vbUnchecked, Val(![StartingSalary] & ""), Val(![Salary] & ""))
                
                 EmpSalary = Val(![StartingSalary] & "")
                 'EmpExtSal = Val(![Salary] & "")
                 EmpExtSal = EmpSalary
            End If
            
            
            
            'Check If Employee is paid Over time per piece
'            If Not IsNull(![OTByQty]) Then If ![OTByQty] Then OTByQty = True
                            
                If Not IsNull(![OTByQty]) Then
                    If ![OTByQty] Then               ' Calculate Over Time On the Basis of Qty Worked If Applicable
                        OTByQty = True
                        If Val(![QtyPerHour] & "") > 0 Then
                            'OTQty = Val(![TotalQtyWorked] & "") - (Val(![MinQty] & "") * SalDays)
                            OverTimeHours = Val(![OTQtyWorked] & "")
                            'OverTimeHours = OTQty / Val(![QtyPerHour] & "")
                        Else
                            OverTimeHours = 0
                        End If
                    Else
                        If IsNull(![OverTimeHours]) Then
                            OverTimeHours = 0
                            OverTimeHoursSA = 0
                        Else
                            OverTimeHours = ![OverTimeHours]
                            OverTimeHoursSA = ![OverTimeHours]
                        End If
                    End If
                    OTByQty = ![OTByQty]
                Else
                    
                    OTByQty = False
                    'OverTime Hours
                    
                    If IsNull(![OverTimeHours]) Then
                        OverTimeHours = 0
                        OverTimeHoursSA = 0
                    Else
                        OverTimeHours = ![OverTimeHours]
                        OverTimeHoursSA = ![OverTimeHours]
                    End If
                    
                End If
            '*********************
'            OverTimeHours = 0
'            OverTimeHoursSA = 0
            '*********************
            Dim theKey As String
            
            If IsComputed Then
                theKey = ![EmpId] & Format(SDate, "dd-MMM-yyyy")
            Else
                theKey = ![EmpId] & ""
            End If
            
            Set ITM = LV.ListItems.Add(, theKey, ![EmpId] & "")
            
            If IsComputed Then
                ITM.Checked = IsPayed
                ITM.Bold = IsTransfered
            End If
            
            
            ITM.Tag = Val(![Emptype] & "")
            
'            If IsInternalFinal And chkExternal.value = vbChecked Then
'                rsInternal.MoveFirst
'                rsInternal.Find "EmpID='" & ![EmpId] & "'", , adSearchForward, 0
'                If rsInternal.EOF Then
'                    ITM.ListSubItems.add(, , ![Name] & "").Tag = 0
'                Else
'                    ITM.ListSubItems.add(, , ![Name] & "").Tag = Val(rsInternal![PrevLTLoan] & "")
'                End If
'            Else
'                ITM.ListSubItems.add(, , ![Name] & "").Tag = Val(![TotAmt] & "") - Val(![clsamt] & "")
'            End If
            ITM.ListSubItems.Add(, , ![Name] & "").Tag = Val(![TotAmt] & "") - Val(![clsamt] & "")
            
            ITM.ListSubItems.Add , , IIf(chkExternal = vbChecked, EmpExtSal, EmpSalary & "")
            
            
            ITM.ListSubItems(2).Tag = ![CasualLeaves] & "," & ![SickLeaves] & "," & ![AnnualLeaves] & _
            "," & ![CompensatoryLeaves] & "," & ![WPLeaves] & "," & ![MaternityLeaves]
            
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '   Rate = Format(Val(EmpSalary) / LastDay, "0.00")
            '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '   Calculate the Rate According To Total No ........
            '   Of Working Days in This Month for External Salary
            '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            
            If chkExternal = vbChecked Then
                
                
                'External Salary Rate Calculated By Working Days
                'Rate = Format(Val(EmpSalary) / (LastDay - Holidays), "0.00")
                Rate = Format(Val(EmpSalary) / (LastDay - TotalSuns), "0.00")
                
                'Multiple Factor for Rate Per hour
                'HrsMultiple = 8
                
                'If Not IsComputed Then HrsMultiple = Val(![WorkingHrs] & "")
                
                
                ' If Working hours are not defined then consider it 8 Hours Per day.
                If HrsMultiple = 0 Then HrsMultiple = 8
            Else
                
                'Rate = Format(Val(EmpSalary) / TempTotalSaldays, "0.00")
                
                
                Rate = Format(Val(EmpSalary) / LastDay, "0.00")
                
                'Multiple Factor for Rate Per Day
                HrsMultiple = 1
                
            End If
            
            
            Rate = Rate / HrsMultiple
            
            
            
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '   Rate Formula For October Salary Sheet ...
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            If chkExternal Then Rate = EmpSalary / Val(txtTotalHrs)
                
            
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            
            
            'Store The Total Working hours for This Employee
            ITM.ListSubItems.Add(, , Format(Rate, "0.00")).Tag = HrsMultiple
            
            
            If IsNull(![presentDays]) And !DeptID <> "VDRS" Then
                SalDays = 0
                'IF No Presents Found Exclude This Employee
                LV.ListItems.Remove ITM.key
                GoTo NEXT_EMPLOYEE
            Else
                
                
                SalDays = ((Val(![presentDays] & "")) + Holidays) + AbsAllowed
                
                'Store for Later Calculation
                
                'LunchDays = Val(![presentDays])
                
                
                
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'Exclude the Holidays now (for Vision External Salary Sheet Only)
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                
                
                
                If chkExternal = vbChecked Then SalDays = ((Val(![presentDays] & ""))) + AbsAllowed
                
                
                'If Salary Days Exceed Total Salary Days Set Them To Salary Days
                If SalDays > TotalSalDays Then SalDays = TotalSalDays
                    
            End If
            
            Dim DaysToExclude As Integer
            
            If Format(![JoinDate], "MMM-yyyy") = Format(JDate, "MMM-yyyy") Then
               DaysToExclude = Day(![JoinDate]) - 1
            Else
               DaysToExclude = 0
            End If
            
            
            'Deduct Penalty Days
            
            'If chkExternal = vbUnchecked Then
            '    SalDays = SalDays - Val(![AbsentDays] & "")
            'End If
            
            If SalDays < 0 Then SalDays = 0
            
            
            
            If chkExternal Then
                Leaves = Val(!Leaves & "")
                UnApprovedLeaves = 0
            Else
                Leaves = Val(!ApprovedLeaves & "")
                UnApprovedLeaves = Val(!Leaves & "") - Val(!ApprovedLeaves & "")
            End If
            

            
            '''''''''''''''''''''''''''''''
            ' dont do anything with sunday/holiday related absents at the moment
            ''''''''''''''''''''''''''''''''
            ' SalDays = SalDays - CalcSpecialAbs(!EmpId, JDate, !BGHoliday, !AGHoliday, !BPHoliday, !APHoliday)
            '''''''''''''''''''''''''''''''
            
            SalDays = SalDays - UnApprovedLeaves
            
            If Abs(![DeductAbsents]) = 0 Then SalDays = TotalSalDays - DaysToExclude
            
            AbsentDays = TotalSalDays - SalDays - DaysToExclude
            
            
            ' Make Absents 0 if in Negative
            If AbsentDays < 0 Then AbsentDays = 0
            
            
            'If IsNull(![AbsentDays]) Then AbsentDays = 0 Else AbsentDays = Val(![AbsentDays]) / 2
            'AbsentDays = IIf(AbsentDays > 0, AbsentDays - AbsAllowed, AbsentDays)
            

            ' Convert To Hours if External Salary
            
            ITM.ListSubItems.Add , , AbsentDays * HrsMultiple
            
            '''''''''''''''''' Save the Salary Hours in it tag Later '''''''''''''''''''''''
            
            
            ITM.SubItems(5) = Round(AbsentDays * HrsMultiple * Rate)
            
            'If Val(!Leaves & "") > 0 Then
            '    If AbsentDays - AbsAllowed >= 0 Then
                
            '    itm.SubItems(11) = IIf((Exempt), "0", (Leaves * Rate))
                
            '    End If
            'Else
            '    itm.SubItems(11) = "0"
            'End If
            
            ITM.SubItems(11) = "0"
            
            ' Calculat Net Absents Allowed
            ITM.SubItems(6) = AbsAllowed
            Dim NetAbsAllowed As Integer
            
            ' Calculate The Allowed Amount That is Subtracted From The Absent Amount
            
            If AbsentDays = 0 Then
                NetAbsAllowed = 0
            Else
                NetAbsAllowed = IIf(AbsAllowed - AbsentDays <= 0, AbsAllowed, AbsAllowed - AbsentDays)
            End If
            
            'Absent Allowed Amount
            ITM.SubItems(7) = IIf((Exempt), "0", Round(Rate * NetAbsAllowed * HrsMultiple, 2))
            
            
            SalHrs = SalDays * HrsMultiple
            
            
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '                      Fix the Salary hours of Directors
            '               and Give Relaxation of 3 hours for october salary
            '                       Add the Leave hours as defined
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            
            
            If chkExternal Then
                If ![DeptID] = "VDRS" Then
                    SalHrs = Val(txtTotalHrs)
                Else
             
                    SalHrs = Val(![HrsWorked] & "") + ((Holidays - TotalSuns) * 8)
                    
                    ' Add All the Leaves to Worked hours
                    SalHrs = SalHrs + (Val(![Leaves] & "") * 8)
                    
                    'Remove Full Hours of Without Pay leaves
                    SalHrs = SalHrs - (Val(![WPLeaves] & "") * 8)
                    
                    ' Remove half Hours of Sick Leaves
                    SalHrs = SalHrs - (Val(![SickLeaves] & "") * 4)
                    
                    ' Check if Salary Hours are greater than Total Hours Due to relaxation
                    ' Fix them to maximum hours
                    
                    If SalHrs > Val(txtTotalHrs) Then SalHrs = Val(txtTotalHrs)
                End If
            Else
                If ![DeptID] = "VDRS" Then
                    SalHrs = Val(txtTotal)
                Else
                    
                    'Remove Full Hours of Without Pay leaves
                    SalHrs = SalHrs - Val(![WPLeaves] & "")
                    
                    'Remove half Hours of Sick Leaves
                    SalHrs = SalHrs - Val(![SickLeaves] & "") / 2
                    
                    ' Check if Salary Hours are greater than Total Hours Due to relaxation
                    ' Fix them to maximum hours
                    
                    If SalHrs > Val(txtTotal) Then SalHrs = Val(txtTotal)
                    
                End If
                
            End If
            
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '       Show in Time Format if external salary sheet
            '       and Show days in internal salarysheet instead
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            ITM.SubItems(8) = IIf(chkExternal, getHrsMin(SalHrs), SalHrs)
            
            ITM.SubItems(9) = Round((SalHrs * Rate) / 10, 0) * 10
            
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '       Save the Salary Hours in Listsubitmes(4).tag To Save
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ITM.ListSubItems(4).Tag = SalHrs
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            ITM.SubItems(10) = Leaves * HrsMultiple
            
            If cmbMonth = "Sep" And cmbYear = "2004" Then
                ' Salary Sheet For September 2004 Must Be shown with the leave deduction
                ' BCoz Its on SA-8000 Record Now ...
                ITM.SubItems(11) = Round(Val(!Leaves & "") * HrsMultiple * Rate, 2)
            
            Else        'This Bug is Corrected Afterwards as follows
            
                ' Do it only For Internal Salary Sheet
                If chkExternal = vbUnchecked Then
                    LeaveDedAmt = Val(![SickLeaves] & "") * HrsMultiple * Rate / 2
                    LeaveDedAmt = LeaveDedAmt + Val(![WPLeaves] & "") * HrsMultiple * Rate
                    ITM.SubItems(11) = Round(LeaveDedAmt, 2)
                End If
                    ITM.SubItems(11) = 0
            End If
            
            
            'If SalDays = 0 Then SalDays = 0 Else SalDays = (SalDays - AbsentDays)
            
            If OTByQty Then
                ITM.SubItems(12) = OverTimeHours
                'Over time By Quantity : Per Day Rate / Sheets Per Day * 1.4
                ITM.SubItems(13) = Round(OverTimeHours * (Rate * HrsMultiple / 56 * 1.4))
            Else
                
                ITM.SubItems(12) = IIf(chkExternal, MinsToTime(OverTimeHoursSA * 60), MinsToTime(OverTimeHours * 60))
                
                If chkExternal Then
                    ITM.SubItems(13) = Round(OverTimeHoursSA * (Rate * 2))
                Else
                    ITM.SubItems(13) = Round(OverTimeHours * (Rate / 8 * 1.3333))
                End If
                
            End If
            
            
            'Itm.SubItems(13) = Round(OverTimeHours * (Rate * 8 / Val(![OverTime] & "")))
            'Itm.SubItems(12) = Val(!Leaves & "") * 8
            'Itm.SubItems(13) = Val(!CurrentLong & "")
            'LateHours = GetLateHours(![EmpId], NTime, LTime, JDate)
            
            
'            If ![DeptID] = "VFNC" Or ![DeptID] = "VSLS" Or ![EmpId] = "VGLV-0001" Then
'                LateHours = Val(![Hours] & "") - LTAllowed
'            Else
'                LateHours = 0
'            End If
            
            LateHours = Val(![Hours] & "") - LTAllowed
            
            If LateHours < 0 Then LateHours = 0
            
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''
            '           Exempt For October 2004
            ''''''''''''''''''''''''''''''''''''''''''''''''''
            'Well It was just written as LateHours=0 And That was why it was'nt deducting LT.
            'So I Just Commented it Out.
            'LateHours = 0
            ''''''''''''''''''''''''''''''''''''''''''''''''''
                        
            ITM.SubItems(14) = MinsToTime(LateHours * 60)
            
           ' itm.SubItems(15) = IIf((Exempt), "0", Format((LateHours * (Rate / Val(!WorkingHrs & ""))), "0.00"))
            
            If chkExternal Then
                ITM.SubItems(15) = IIf((Exempt), "0", Format((LateHours * Rate), "0.00"))
            Else
                ITM.SubItems(15) = IIf((Exempt), "0", Format((LateHours * Rate / 8), "0.00"))
            End If
             
             
            'Total = Format(Val((ITM.SubItems(9)) + Val(ITM.ListSubItems(7))) - (Val(ITM.SubItems(15)) + Val(ITM.SubItems(11))), "0.00")
            Total = Format(Val((ITM.SubItems(9)) + Val(ITM.ListSubItems(7))) - (Val(ITM.SubItems(11))), "0.00")
            ITM.SubItems(16) = Total
            
            'TaxPer = Val(![PerDeduct])
            ITM.SubItems(17) = Val(!TaxAmt & "") 'Round(getTaxAmt((EmpSalary * 12) / 1.5) / 12, 0)
            'DedAmt = Round((Val(![Salary]) / TotalSalDays) * (TaxPer / 100))
            'TaxSal = Round((Val(EmpSalary) / LastDay) * TotalSalDays)
            DedAmt = 0 'Round(TaxSal * (TaxPer / 100))
            
            'Tax Percentag That Is Zero Length
            'itm.SubItems(18) = "0"
            
            'itm.SubItems(14) = DedAmt
            ITM.SubItems(18) = Total 'Total - Val(ITM.SubItems(17))
            If IsNull(![Payed]) Then Payed = 0 Else Payed = ![Payed]
            ITM.SubItems(19) = Round(Payed / 10) * 10
            
            If chkExternal.value = vbChecked And IsInternalFinal Then
                If rsInternal.EOF Then
                    ITM.SubItems(20) = 0
                Else
                    ITM.SubItems(20) = Val(rsInternal![LongTerm] & "")
                End If
            Else
                ITM.SubItems(20) = LongLoan
            End If
            
            ITM.SubItems(21) = Val(![AdvSalpaid] & "")
            
            
            'Dim SocialAmt As Double
            'SocialAmt = IIf((!Social = True), !SocialAmt, 0)
            
            
            ''''''''''''''''''''''''''''''''''''''''''
            '       Union Fund deduction
            ''''''''''''''''''''''''''''''''''''''''''
            
                If Not IsNull(![UnionFund]) Then
                    If ![UnionFund] Then UnionFund = 20 Else UnionFund = 0
                Else
                    UnionFund = 0
                End If
            
            ''''''''''''''''''''''''''''''''''''''''''
                Dim LunchDed As Double
                
                If Not IsNull(![Lunch]) Then
                    LunchDed = Val(![LunchAmt] & "")
                Else
                    LunchDed = 0
                End If
            ''''''''''''''''''''''''''''''''''''''''''
            
            ITM.SubItems(22) = UnionFund
            ITM.SubItems(23) = Val(!TotalFine & "")
            
            ' Use This Column For Scholorships Instead Of Bonus ...
            'ITM.SubItems(24) = Val(!TotalBonus & "")
            
            ITM.SubItems(24) = Val(![ScholorshipAmt] & "")
            ITM.SubItems(25) = LunchDed
            
            'If Not IsNull(![OldAge]) Then If ![OldAge] Then EOBIAmt = 20
            EmpAge = DateDiff("M", !DOB, JDate) / 12
            
            If ![DeptID] = "VDRS" Or EmpAge > 60 Or !ArmyRetired = True Then
                EOBIAmt = 0
            Else
                EOBIAmt = 30
            End If
            '/\/\//\/\ IF U CHANGE EOBI FORMULA HERE, REMEMBER TO CHANGE IN rptSalarySheet and rptPaySlip TOO./\/\/\/\/\
            
            ITM.SubItems(26) = EOBIAmt
            
            'Removed OverTime From Balance
            ' ITM.SubItems(27) = Round((Val(ITM.SubItems(18)) + Val(!ScholorshipAmt & "") _
            - (Payed + Val(!TotalFine & "") + UnionFund + LongLoan + LunchDed + Val(!AdvSalpaid & "") + Val(ITM.SubItems(26))) _
            + Val(ITM.SubItems(13))))
            
            ITM.SubItems(27) = Round((Val(ITM.SubItems(18)) + Val(!ScholorshipAmt & "") _
            - (Payed + Val(!TotalFine & "") + UnionFund + LongLoan + LunchDed + Val(!AdvSalpaid & "") + Val(ITM.SubItems(17)) + Val(ITM.SubItems(26)))) / 10, 0) * 10
            
            
            GrandTotal = GrandTotal + Val(ITM.SubItems(27))
            
            Dim LastDate As Date
            LastDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), LastDay)


            If (IsComputed) Then
                cmdUpdate.Visible = UserHasAccess("FinalSalSheet")          'And (chkExternal = vbUnchecked)
            ElseIf Not (IsComputed) And (GetServerDate > LastDate) Then
                cmdUpdate.Visible = UserHasAccess("FinalSalSheet")          'And (chkExternal = vbUnchecked)
            Else
                cmdUpdate.Visible = False
            End If
            
            
NEXT_EMPLOYEE:
            .MoveNext
        Next
        
        LV.Visible = True
        
        lblTotal = Format(GrandTotal, "#,##0.00")
        
    End With
    cmbMonth.Tag = cmbMonth.Text
    cmbYear.Tag = cmbYear.Text
    
    Sql = ""
    Cond = ""
    ICO = ""
    
    If cmbDept.ID = "0" And cmbEmp.ID = "0" Then
        cmdUpdate.Caption = "Make This Sheet Final"
        cmdUpdate.Visible = UserHasAccess("FinalSalSheet")
    Else
        cmdUpdate.Visible = False
    End If
    
End Sub
Private Function GetLateHours(EmpId As String, NTime As Integer, LTime As Double, ToDate As Date) As Double

    Dim LHours As Double
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.command
    
    With cmd
        '.CommandText = "Select LHrs from VLHours Where EmpID='" & EmpID & "'"
        .CommandText = "VLHours"
        .CommandType = adCmdStoredProc
         Set .ActiveConnection = con
        .Parameters.Refresh
        .Parameters(1).Type = adNumeric
        .Parameters(1).value = LTime
        
        .Parameters(2).Type = adDBTimeStamp
        .Parameters(2).value = ToDate
        
        .Parameters(3).Type = adNumeric
        .Parameters(3).value = NTime
        
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        Set rs = .Execute
        
    End With
    
    If rs.RecordCount > 0 Then
        If IsNull(rs.Fields(1)) Then LHours = 0 Else LHours = rs.Fields(1)
    Else
        LHours = 0
    End If
    GetLateHours = LHours
    
End Function
Private Sub Computed(Emp As String, ByRef IsComputed As Boolean, ByRef TaxPer As Double, ByRef AbsentDays As Integer, ByRef OverTimeHours As Double, ByRef NTime As Integer, ByRef LTime As Double, ByRef LongLoan As Double, ByRef DT As String, ByRef Payed As Boolean, ByRef Transfered As Boolean, ToDate As Date, ByRef EmpSal As Double, Optional EmpExtSal As Double, Optional IsActive As Boolean, Optional HrsPerDay As Double)

    
Exit Sub
    
    Dim rsCheck As New ADODB.Recordset
    Dim cmd As New ADODB.command
    
    Dim IsExernalSalary As Byte
    
    IsExernalSalary = chkExternal
    
    With cmd
        '.CommandText = "Select * from VSalarySettings where empid='" & Emp & "'"
        .CommandType = adCmdStoredProc
        .CommandText = "VSalarySettings"
        Set .ActiveConnection = con
        .Parameters.Refresh
        .Parameters(1).Type = adDBTimeStamp
        .Parameters(1).value = ToDate
        
        
        .Parameters(2).Type = adVarChar
        .Parameters(2).value = Emp
        
        .Parameters(3).Type = adBoolean
        .Parameters(3).value = IsExernalSalary
        
        rsCheck.CursorType = adOpenForwardOnly
        rsCheck.LockType = adLockReadOnly
        Set rsCheck = .Execute()
    End With
    
    'If Emp Like "VLMT*" Then Stop
    
    With rsCheck
        
        If .RecordCount > 0 Then
            IsComputed = True
            TaxPer = ![PerDeduct]
            AbsentDays = ![absentallowed]
            'OverTimeHours = ![OverTime]
            NTime = ![NTimes]
            LTime = ![TimeAllowed]
            DT = Format(![DT], "dd-MMM-yyyy") & ""
            ', AbsentDays, OverTimeHours, NTime, LTime)
            LV.Checkboxes = True
            cmdUpdate.Visible = True
            
            Payed = ![Payed] & ""
            Transfered = ![Transfered] & ""
            
            If IsNull(![LongLoan]) Then LongLoan = 0 Else LongLoan = ![LongLoan]
            
            HrsPerDay = Val(![HrsPerDay] & "")
            
            'Get The Salary Depending On Salary Sheet Type
            'EmpSal = IIf(chkExternal = vbUnchecked, Val(![Salary] & ""), Val(![ExtSalary] & ""))
            EmpSal = Val(![Salary] & "")
            'EmpExtSal = Val(![ExtSalary] & "")
            EmpExtSal = EmpSal
            cmdUpdate.Caption = "Save Changes"
            
        Else
            If IsActive Then
                IsComputed = False
                LV.Checkboxes = False
                'cmdUpdate.Visible = False
                cmdUpdate.Caption = "Make This Sheet Final"
            End If
        End If
        
        
    End With
    
End Sub


Private Sub SaveIt_OLD(rs As ADODB.Recordset, PerDeduct As Double, Optional i As Integer)


    Dim DT As Date
    Dim LongLoan As Double
    Dim lDedAmt As Long, lClsAmt1 As Long, lTotAmt As Long
    
    If IsNull(rs![DedAmt]) Then lDedAmt = 0 Else lDedAmt = rs![DedAmt]
    If IsNull(rs![TotAmt]) Then lTotAmt = 0 Else lTotAmt = rs![TotAmt]
    If IsNull(rs![clsamt]) Then lClsAmt1 = 0 Else lClsAmt1 = rs![clsamt]
                
    If lDedAmt > (lTotAmt - lClsAmt1) Then
        LongLoan = (lTotAmt - lClsAmt1)
    Else
        LongLoan = lDedAmt
    End If
    
    
    
    'If IsNull(rs![LongLoan]) Then LongLoan = 0 Else LongLoan = rs![LongLoan]
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    
    
    'Get The Last Day of Month
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), DateAdd("m", 1, DT) - DT)
    
    
    con.Execute "Insert into MonthlySalaries (EmpID,DT,Salary," & _
    "ExtSalary,PerDeduct,AbsentAllowed,TimeAllowed,NTimes,OverTime,Payed,ShortTermPer,LongLoan,Transfered," & _
    "SASalary,NetSalary,OTAmt,SalAmt,LtAmt,StAmt,LateAmt,SalDays,HrsPerDay)" & " Values('" & rs![EmpId] & "','" & DT & _
    "'," & rs![StartingSalary] & "," & rs![Salary] & "," & PerDeduct & "," & Abs(rs![AbsAllowed]) & _
    "," & rs![TimeAllowed] & "," & Val(rs![NTimes] & "") & "," & Val(rs![OverTimeHours] & "") & ",0," & _
    Val(rs![ShortTermPer] & "") & "," & LongLoan & ",0," & chkExternal.value & "," & Val(LV.ListItems(i).SubItems(27)) + Val(LV.ListItems(i).SubItems(21)) & _
    "," & Val(LV.ListItems(i).SubItems(13)) & "," & Val(LV.ListItems(i).SubItems(9)) & "," & Val(LV.ListItems(i).SubItems(20)) & _
    "," & Val(LV.ListItems(i).SubItems(19)) & "," & Val(LV.ListItems(i).SubItems(15)) & "," & _
    IIf(chkExternal, Val(LV.ListItems(i).SubItems(8)) / Val(LV.ListItems(i).ListSubItems(3).Tag), Val(LV.ListItems(i).SubItems(8))) & "," & Val(LV.ListItems(i).ListSubItems(3).Tag) & ")"
    

        
    'Update Long Term ...
    
    Dim LongAmt As Double
    LongAmt = Val(LV.ListItems(i).SubItems(20))
    
    
    con.Execute "Delete From AmtCleared Where empID='" & rs![EmpId] & "' and DT='" & Format(DT, "dd-MMM-yyyy") & "'"
    
    con.Execute "Insert Into AmtCleared (EmpID,DT,AmtClrd) Values('" & rs![EmpId] & "','" & DT & "'," & LongAmt & ")"
    
    
    If LV.Checkboxes = False Then
        LV.Checkboxes = True
        cmdUpdate.Visible = True
    End If
    
    
End Sub
Private Sub SaveIt()
    
Dim rs As New ADODB.Recordset
With rs
    .Open "Select * from MonthlySalaries", con, adOpenStatic, adLockOptimistic
    For i = 1 To LV.ListItems.Count
        .AddNew
        ![EmpId] = LV.ListItems(i).Text & ""
        
        DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
        DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), DateAdd("m", 1, DT) - DT)
        
        
       ' If LV.ListItems(i).Text & "" = "VGLV-0002" Then Stop

        ![DT] = DT & ""
        
        For j = 2 To 27
            If j = 8 And chkExternal Then
                .Fields(j + 2) = LV.ListItems(i).ListSubItems(4).Tag
            Else
                .Fields(j + 2) = LV.ListItems(i).SubItems(j) & ""
            End If
        Next j
        
        ![PrevLTLoan] = Val(LV.ListItems(i).ListSubItems(1).Tag)
                
                
        Dim ArrLeaves() As String
        
        ArrLeaves = Split(LV.ListItems(i).ListSubItems(2).Tag, ",")
        
        ![CasualLeaves] = Val(ArrLeaves(0))
        ![SickLeaves] = Val(ArrLeaves(1))
        ![AnnualLeaves] = Val(ArrLeaves(2))
        ![CompensatoryLeaves] = Val(ArrLeaves(3))
        ![WPLeaves] = Val(ArrLeaves(4))
        ![MaternityLeaves] = Val(ArrLeaves(5))
        ![HrsPerDay] = Val(LV.ListItems(i).ListSubItems(3).Tag)
        ![IsForSA] = chkExternal
        ![TotalMonthHrs] = Val(txtTotalHrs.Tag)
        ![DTFinal] = Date
        
        .Update
        
    'Update Long Term ...
    
    Dim LongAmt As Double
    LongAmt = Val(LV.ListItems(i).SubItems(20))
    
    con.Execute "Delete From AmtCleared Where empID='" & rs![EmpId] & "' and DT='" & Format(DT, "dd-MMM-yyyy") & "'"
    con.Execute "Insert Into AmtCleared (EmpID,DT,AmtClrd) Values('" & rs![EmpId] & "','" & DT & "'," & LongAmt & ")"
    
    Next i
    
    .Close
    
End With

    
    
Set rs = Nothing

End Sub

Private Function TotalMonthDays() As Integer

    Dim TDate As Date
    Dim Tot As Integer
    Dim year As Integer
    If cmbYear = "" Then
        year = 2002
    Else
        year = cmbYear
    End If
    
    If Month(GetServerDate) = cmbMonth.List(cmbMonth.ListIndex, 1) Then
        Tot = Day(GetServerDate)
    Else
        TDate = DateSerial(year, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
        Tot = DateAdd("m", 1, TDate) - TDate
    End If
    TotalMonthDays = Tot
    
End Function

Private Sub LV_DblClick()

    If LV.ListItems.Count = 0 Then Exit Sub
    If Val(LV.SelectedItem.SubItems(20)) = 0 Then Exit Sub
    If LV.SelectedItem.Bold Then Exit Sub
    
    If chkExternal.value = vbChecked Then Exit Sub
    
    Dim Amt As String
    Amt = InputBox("Please Enter Amount To Deduct For Long Term." & vbNewLine & "The Current Deduction Amount Is " & LV.SelectedItem.SubItems(20) & ".", , LV.SelectedItem.SubItems(20)) & ""
    
    If Not IsNumeric(Amt) Then
        MsgBox "Invalid Amount.", vbInformation
    Else
        
        ' Update The Balance
        
        LV.SelectedItem.SubItems(27) = Val(LV.SelectedItem.SubItems(27)) + Val(LV.SelectedItem.SubItems(20)) - Abs(Val(Amt))
        LV.SelectedItem.SubItems(20) = Abs(Val(Amt))
        
        ' Update The Saved Entries
        
        con.Execute "Update AmtCleared Set AmtClrd=" & Abs(Val(Amt)) & "  Where convert(char(3),dt,7)='" & cmbMonth & "' and EmpID='" & LV.SelectedItem & "'"
        con.Execute "Update MonthlySalaries Set LongTerm=" & Abs(Val(Amt)) & " Where convert(char(3),dt,7)='" & cmbMonth & "' and EmpID='" & LV.SelectedItem & "'"
        
    End If
    
    ' With LV.SelectedItem
    '    txtEdit.Move LV.Left + LV.ColumnHeaders(18).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(18).Width
    '    txtEdit = .SubItems(17)
    ' End With
    '
    ' With txtEdit
    '     .Visible = True
    '     .SelStart = 0
    '     .SelLength = Len(.Text)
    '     .SetFocus
    ' End With
    ' MsgBox LV.FindItem (

End Sub

Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If Item.Bold Then
      Item.Checked = Not Item.Checked
   MsgBox "Cannot Modify This Entry It Has Been Transfered To Accounts", vbInformation

End If
End Sub

Private Sub LV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If LV.ListItems.Count = 0 Then
    EditAbsents = False
    Exit Sub
End If
    If X >= LV.Left + LV.ColumnHeaders(5).Left And X < LV.Left + LV.ColumnHeaders(6).Left And _
        Y >= LV.SelectedItem.Top And Y < LV.SelectedItem.Top + LV.SelectedItem.Height Then
        EditAbsents = True
    Else
        EditAbsents = False
    End If
    
End Sub

Private Sub mnuLateEmpList_Click()
    Call PrintLateEmpList
End Sub

Private Sub mnuPaySlips_Click()
   Call PrintReport(True)
End Sub

Private Sub mnuPayslipsInternal_Click()
Call PrintInternalReport(True)
End Sub

Private Sub mnuPrintOTPaySheet_Click()

    If LV.ListItems.Count = 0 Then Exit Sub
    Call InsertRecords

    Dim unloadflag As Boolean

    Screen.MousePointer = vbHourglass

    Dim f As New frmPrevRpt, rpt  As craxddrt.Report

    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptOverTimePaymentSheet.rpt")
    
    Dim FormulaFields As craxddrt.FormulaFieldDefinitions
    Dim FormulaField As craxddrt.FormulaFieldDefinition

    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & company & "'"
        ElseIf FormulaField.Name = "{@FromTo}" Then
            FormulaField.Text = "' For " & cmbMonth.Tag & "-" & cmbYear.Tag & "'"
        ElseIf FormulaField.Name = "{@MonthDays}" Then
            FormulaField.Text = Val(txtTotalHrs)
        End If
    Next
    Cond = "{@OHAmt}>0"
    rpt.SQLQueryString = Sql
    f.ShowReport Cond & " ", rpt

    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuSalarysheetInternal_Click()
 Call PrintInternalReport(False)
End Sub

Private Sub mnuSalSheet_Click()
   Call PrintReport(False)
End Sub

Private Sub PrintSalsummary_Click()
 Call PrintReport(False, , True)
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LV.SelectedItem.SubItems(19) = txtEdit.Text
        txtEdit.Text = ""
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEdit.Text = ""
        txtEdit.Visible = False
    Else
        KeyAscii = OnlyNumber(KeyAscii)
    End If
End Sub






Private Function AddSalaryEntries() As Boolean

   Dim SSAdv As Double, SLadv As Double, SSal As Double, STax As Double     'Staff Member Vars
   Dim DSAdv As Double, DLadv As Double, DSal As Double, DTax As Double     'Directors Vars
   Dim PSAdv As Double, PLadv As Double, PSal As Double, PTax As Double     'Production Staff Vars
   
   
   
   Dim Dated As Date
   Dated = GetServerDate
   
   '''''''''''''''''' A D V A N C E   C L E A R E N C E ''''''''''''''''''
   ConAcc.BeginTrans
   
   Call GetTotal(SSAdv, DSAdv, PSAdv, "SADV")        'Short Term Advances
   Call AddAdvClearVoucher(SSAdv, DSAdv, PSAdv, Dated, False, cmbMonth & ", " & cmbYear)
   
   Call GetTotal(SLadv, DLadv, PLadv, "LADV")        'Long  Term Advances
   Call AddAdvClearVoucher(SLadv, DLadv, PLadv, Dated, True, cmbMonth & ", " & cmbYear)
   
   
   '''''''''''''''''' S A L R A R Y   &   T A X    E N T R Y ''''''''''''''''''
   
   Call GetTotal(SSal, DSal, PSal, "SAL")            'Salaries
   Call GetTotal(STax, DTax, PTax, "TAX")            'Taxes
   
   Call addSalaryVoucher(SSal, DSal, PSal, Dated, STax, DTax, PTax, cmbMonth & ", " & cmbYear)
   
   ConAcc.CommitTrans
   
   
   
    For i = 1 To LV.ListItems.Count
    
        con.Execute "Update MonthlySalaries Set Transfered=" & CBool(LV.ListItems(i).Checked Or LV.ListItems(i).Bold) & " Where EmpID+Format(DT,'dd-MMM-yyyy')='" & LV.ListItems(i).key & "'"
    
    Next i
      
   
   AddSalaryEntries = True
   Exit Function
err:
   AddSalaryEntries = False
   MsgBox "Cannot Add Salary Vouchers Due To Error in " & err.Source & " :" & vbNewLine & err.Description, vbCritical
   ConAcc.RollbackTrans
End Function

Private Sub GetTotal(Stotal As Double, DTotal As Double, PTotal As Double, TotalType As String)

   Dim StaffTotal As Double, ProductionTotal As Double, DirectorTotal As Double, Amount As Double



   For i = 1 To LV.ListItems.Count
      'if Paid or Transfered then
      If Not LV.ListItems(i).Checked Or LV.ListItems(i).Bold Then GoTo NEXT_ITEM
         
   
         Select Case TotalType
            Case "LADV"
               'Long Term Advances
               Amount = Val(LV.ListItems(i).SubItems(19))
            Case "SADV"
               'Short Tem Advances
               Amount = Val(LV.ListItems(i).SubItems(20))
            Case "SAL"
               'Salary amount
               Amount = Val(LV.ListItems(i).SubItems(16))
            Case "TAX"
               Amount = Val(LV.ListItems(i).SubItems(17))
         End Select
   
   
   
   
    Select Case LV.ListItems(i).Tag
      Case "0"
         StaffTotal = StaffTotal + Amount
      Case "1"
         DirectorTotal = DirectorTotal + Amount
      Case "2"
         ProductionTotal = ProductionTotal + Amount
    End Select


NEXT_ITEM:
   Next
   
   PTotal = ProductionTotal
   Stotal = StaffTotal
   DTotal = DirectorTotal
   
End Sub

Private Function GetTotalTax() As Double
Dim Total As Double
   For i = 1 To LV.ListItems.Count
    Total = Total + Val(LV.ListItems(i).SubItems(14))
   Next
GetTotalTax = Total
End Function


Public Sub ShowExternalSalSheet(forMonth As String, forYear As String, Dept As String, Crpt As CrystalReport)
   Me.Visible = False
   cmbMonth.Text = forMonth
   cmbYear.Text = forYear
   cmbDept.Text = Dept
   cmbEmp.ID = 0
   chkExternal = vbChecked
   Call PrintReport(False, Crpt)
End Sub

Private Sub mnuCurrency_Click()

    Thousands = 0
    FHundreds = 0
    Hundreds = 0
    Fifties = 0
    Tens = 0
    Fives = 0
    Twos = 0
    Ones = 0
    TotalSal = 0


    For i = 1 To LV.ListItems.Count
        If Val(LV.ListItems(i).ListSubItems(27)) > 0 Then
            CountRupee Val(LV.ListItems(i).ListSubItems(27))
            TotalSal = TotalSal + Val(LV.ListItems(i).ListSubItems(20))
        End If
    Next
    
'    With CR2
'        .ReportFileName = RptPath & "\RupeeStructure.rpt"
'
'            .Formulas(0) = "Thousands=" & Thousands
'            .Formulas(1) = "FHundreds=" & FHundreds
'            .Formulas(2) = "Hundreds=" & Hundreds
'            .Formulas(3) = "Fifties=" & Fifties
'            .Formulas(4) = "Tens=" & Tens
'            .Formulas(5) = "Fives=" & Fives
'            .Formulas(6) = "Tows=" & Twos
'            .Formulas(7) = "Ones=" & Ones
'            .Formulas(8) = "TotalSal='" & lblTotal & "'"
'            .Formulas(9) = "Company='" & company & "'"
'
'        .WindowState = crptMaximized
'        .WindowTitle = "Print Rupee Structure"
'        .Action = 1
'
'    End With

Screen.MousePointer = vbHourglass

Dim f As New frmPrevRpt, rpt As craxddrt.Report
Set rpt = rpt.OpenSubreport(rptPayrollPath & "\rptRupeeStructure.rpt")
Dim FormulaFields As craxddrt.FormulaFieldDefinitions
Dim FormulaField As craxddrt.FormulaFieldDefinition

Set FormulaFields = rpt.FormulaFields
For Each FormulaField In FormulaFields
    If FormulaField.Name = "{@Company}" Then
        FormulaField.Text = ""
    ElseIf FormulaField.Name = "{@Thousands}" Then
        FormulaField.Text = Thousands
    ElseIf FormulaField.Name = "{@FHundreds}" Then
        FormulaField.Text = FHundreds
    ElseIf FormulaField.Name = "{@Hundreds}" Then
        FormulaField.Text = Hundreds
    ElseIf FormulaField.Name = "{@Fifties}" Then
        FormulaField.Text = Fifties
    ElseIf FormulaField.Name = "{@Tens}" Then
        FormulaField.Text = Tens
    ElseIf FormulaField.Name = "{@Fives}" Then
        FormulaField.Text = Fives
    ElseIf FormulaField.Name = "{@Tows}" Then
        FormulaField.Text = Twos
    ElseIf FormulaField.Name = "{@Ones}" Then
        FormulaField.Text = Ones
    ElseIf FormulaField.Name = "{@TotalSal}" Then
        FormulaField.Text = Val(Replace(lblTotal, ",", ""))
    End If
Next

f.ShowReport " ", rpt


Screen.MousePointer = vbDefault


End Sub

Private Sub CountRupee(Salary As Double)

    Thousands = Thousands + Int(Salary / 1000)
    If Salary >= 500 Then FHundreds = FHundreds + Int(Val(Right(Salary, 3)) / 500)
    Hundreds = Hundreds + Int(Val(Right(Salary, 3) Mod 500) / 100)
    Fifties = Fifties + Int(Val(Right(Salary, 2)) / 50)
    Tens = Tens + Int(Val(Right(Salary, 3) Mod 50) / 10)
    Fives = Fives + Int(Val(Right(Salary, 1)) / 5)
    Twos = Twos + Int(Val(Right(Salary, 1) Mod 5) / 2)
    Ones = Ones + Int(Val(Right(Salary, 1) Mod 5) Mod 2)


End Sub

Private Sub InsertRecords()

Dim DT As Date
DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)

Dim i As Integer
Dim j As Integer

con.Execute "Delete from PrintSalary"

Dim rs As New ADODB.Recordset
With rs
    .Open "Select * from PrintSalary", con, adOpenStatic, adLockOptimistic
    For i = 1 To LV.ListItems.Count
        .AddNew
        ![EmpId] = LV.ListItems(i).Text & ""
        
        For j = 1 To 25
            If j = 25 Then
                .Fields(j) = LV.ListItems(i).SubItems(j + 2) & ""
            ElseIf j = 8 And chkExternal Then                           '''' For Saving in Decimal Format instead of Time
                .Fields(j) = LV.ListItems(i).ListSubItems(4).Tag
            'ElseIf .Fields(j).Name = "OHrs" And chkExternal Then
            '    .Fields(j) = LV.ListItems(i).ListSubItems(12).Tag
            Else
                .Fields(j) = LV.ListItems(i).SubItems(j) & ""
            
            End If
        Next j
        
        ![DT] = DT & ""
        ![EOBI] = Val(LV.ListItems(i).ListSubItems(26))
        ![PrevLTLoan] = Val(LV.ListItems(i).ListSubItems(1).Tag)
                
        Dim ArrLeaves() As String
        
        ArrLeaves = Split(LV.ListItems(i).ListSubItems(2).Tag, ",")
        
        ![CasualLeaves] = Val(ArrLeaves(0))
        ![SickLeaves] = Val(ArrLeaves(1))
        ![AnnualLeaves] = Val(ArrLeaves(2))
        ![CompensatoryLeaves] = Val(ArrLeaves(3))
        ![WPLeaves] = Val(ArrLeaves(4))
        ![MaternityLeaves] = Val(ArrLeaves(5))
        ![HrsPerDay] = Val(LV.ListItems(i).ListSubItems(3).Tag)
        
        .Update
    Next i
    
    .Close
    
End With

Set rs = Nothing

    Sql = "insert into EmpLeaveBalance SELECT  EmpID,Count(CL) as CasualLeaves,Count(SL) as SickLeaves," & _
    "count(AL) as AnnualLeaves" & _
    ",count(ML) as MaternityLeaves From VEmpLeaves WHERE DT <'" & DateSerial(cmbYear, cmbMonth.ListIndex + 1, 1) & "' Group By empID"
    
    con.Execute "Delete From EmpLeaveBalance"
    con.Execute Sql, a

End Sub

Private Function FinalizedSalarySheetLoaded() As Boolean

    Dim rs As New ADODB.Recordset
    Dim Cond As String
    Dim IsExernalSalary As Byte
    Dim GrandTotal As Double
    
    IsExernalSalary = chkExternal
    
    If cmbEmp.ID <> "0" Then
        Cond = "Where EmpID='" & cmbEmp.ID & "'"
    ElseIf cmbDept.ID <> "0" Then
        Cond = "Where DeptID='" & cmbDept.ID & "'"
    End If
    
    If Cond = "" Then
        Cond = "Where"
    Else
        Cond = Cond & " And "
    End If
    
    
    Cond = Cond & " Month(DT)=" & cmbMonth.List(cmbMonth.ListIndex, 1) & _
    " and Year(DT)=" & Val(cmbYear) & " and IsForSA=" & chkExternal.value
    
    
    With rs
        .Open "Select * From VMonthlySalaries " & Cond, con, adOpenForwardOnly, adLockReadOnly
        
        If .RecordCount = 0 Then
            FinalizedSalarySheetLoaded = False
            GoTo END_OF_FUNCTION
        End If
            
            LV.ListItems.Clear
            Do Until .EOF
        
                Set ITM = LV.ListItems.Add(, ![EmpId], ![EmpId] & "")
                
'                If IsComputed Then
'                    ITM.Checked = IsPayed
'                    ITM.Bold = IsTransfered
'                End If
                
                If Val(txtTotalHrs) <> Val(![TotalMonthHrs] & "") Then
                    txtTotalHrs = ![TotalMonthHrs] & ""
                    MsgBox "Salary Sheet Is Finalized According To " & txtTotalHrs & " Working Hours.", vbInformation
                End If
                
            
                ITM.Tag = Val(![Emptype] & "")
                
                ITM.ListSubItems.Add(, , ![Name] & "").Tag = Val(![PrevLTLoan] & "")
                ITM.ListSubItems.Add , , ![BSal] & ""
                
                ITM.ListSubItems(2).Tag = ![CasualLeaves] & "," & ![SickLeaves] & "," & ![CompensatoryLeaves] & _
                "," & ![AnnualLeaves] & "," & ![WPLeaves] & "," & ![MaternityLeaves]
                
                ITM.ListSubItems.Add(, , ![Rate] & "").Tag = ![HrsPerDay]
                ITM.ListSubItems.Add(, , ![ADays] & "").Tag = Val(![SDays] & "")
                ITM.ListSubItems.Add , , ![AAmt] & ""
                ITM.ListSubItems.Add , , ![AAllow] & ""
                ITM.ListSubItems.Add , , ![AAllowAmt] & ""
                ITM.ListSubItems.Add , , IIf(chkExternal = vbChecked, getHrsMin(![SDays]), ![SDays] & "")
                ITM.ListSubItems.Add , , ![SAmt] & ""
                ITM.ListSubItems.Add , , ![Leaves] & ""
                ITM.ListSubItems.Add , , ![LeaveAmt] & ""
                ITM.ListSubItems.Add , , ![OHrs] & ""
                ITM.ListSubItems.Add , , ![OAmt] & ""
                ITM.ListSubItems.Add , , ![lHrs] & ""
                ITM.ListSubItems.Add , , ![LAmt] & ""
                ITM.ListSubItems.Add , , ![Total] & ""
                ITM.ListSubItems.Add , , ![Tax] & ""
                ITM.ListSubItems.Add , , ![NetTtl] & ""
                ITM.ListSubItems.Add , , ![ShortTerm] & ""
                ITM.ListSubItems.Add , , ![LongTerm] & ""
                ITM.ListSubItems.Add , , ![AdvSal] & ""
                ITM.ListSubItems.Add , , ![UnionFund] & ""
                ITM.ListSubItems.Add , , ![Fine] & ""
                ITM.ListSubItems.Add , , ![Bonus] & ""
                ITM.ListSubItems.Add , , ![Lunch] & ""
                ITM.ListSubItems.Add , , ![EOBI] & ""
                ITM.ListSubItems.Add , , ![Balance] & ""
                
                GrandTotal = GrandTotal + Val(ITM.SubItems(27))
                
                .MoveNext
            Loop
        
        
    End With
    
    lblTotal = Format(GrandTotal, "#,##0.00")
    
    FinalizedSalarySheetLoaded = True
    cmdUpdate.Visible = False
END_OF_FUNCTION:

    If rs.State = 1 Then rs.Close
    Set rs = Nothing
    
End Function

Private Sub txtTotalHrs_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumber(KeyAscii, False)
End Sub

Private Sub PrintLateEmpList()
    Call InsertRecords
    
    Screen.MousePointer = vbHourglass

    Dim f As New frmPrevRpt, rpt  As craxddrt.Report


    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptLateEmpList.rpt")

    Cond = "{PrintSalary.LAmt}>0 "
    
    'rpt.RecordSelectionFormula = "{PrintSalary.LAmt}>0"
    f.ShowReport Cond & " ", rpt

    Screen.MousePointer = vbDefault

End Sub
