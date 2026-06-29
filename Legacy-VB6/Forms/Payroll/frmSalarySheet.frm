VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmSalarySheet 
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   7785
   ScaleMode       =   0  'User
   ScaleWidth      =   10748.56
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2505
      Top             =   1815
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
            Picture         =   "frmSalarySheet.frx":0000
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalarySheet.frx":0452
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
      Left            =   5715
      TabIndex        =   8
      Top             =   7200
      Visible         =   0   'False
      Width           =   2145
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   315
      Left            =   3510
      TabIndex        =   3
      Top             =   330
      Width           =   2220
      _ExtentX        =   3916
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
      Left            =   5730
      TabIndex        =   4
      Top             =   330
      Width           =   4935
      _ExtentX        =   8705
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
      Height          =   6915
      Left            =   75
      TabIndex        =   0
      Top             =   690
      Width           =   11865
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
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   5070
         TabIndex        =   7
         Top             =   690
         Visible         =   0   'False
         Width           =   1155
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5820
         Left            =   105
         TabIndex        =   1
         Top             =   180
         Width           =   11685
         _ExtentX        =   20611
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
         NumItems        =   31
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
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   28
            Text            =   "extra-1"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   29
            Text            =   "extra"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   30
            Text            =   "Salary Paid"
            Object.Width           =   0
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
      Begin VB.Label lblTransferNegative 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Transfer Negative"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6915
         MouseIcon       =   "frmSalarySheet.frx":0564
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   6585
         Visible         =   0   'False
         Width           =   1455
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
      Width           =   10575
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
      Left            =   10710
      TabIndex        =   2
      Top             =   255
      Width           =   1245
      Caption         =   "Refresh        "
      PicturePosition =   327683
      Size            =   "2196;635"
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
      Begin VB.Menu mnuSalSheetNegative 
         Caption         =   "Print Salary Sheet (&Negative)"
      End
      Begin VB.Menu mnuPrintNegativeEmpList 
         Caption         =   "Print Negative Emp. List"
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
         Visible         =   0   'False
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
      Begin VB.Menu mnuPayslipRule9 
         Caption         =   "Print Payslip (Rule 9)"
      End
      Begin VB.Menu mnuPayslipMini 
         Caption         =   "Print Payslip Mini"
      End
      Begin VB.Menu mnuPayslipFull 
         Caption         =   "Print Payslip (Full)"
      End
      Begin VB.Menu mnuPrintSalPacket 
         Caption         =   "Print Salary Packet"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPayslipsInternal 
         Caption         =   "Print Payslips (Internal)"
      End
      Begin VB.Menu mnuPrintOTPaySheet 
         Caption         =   "Print Over Time Payment Sheet"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRePostSalaryVouchers 
         Caption         =   "Re-Post Salary Vouchers"
      End
   End
End
Attribute VB_Name = "frmSalarySheet"
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
Dim SQL As String
Public m_bShowFake As Boolean
Dim bSalaryFinalyzed As Boolean

Private Sub chkExternal_Click()
    
    With LV.ColumnHeaders
        For i = 1 To .count
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
    MonthDays = txtTotal '- TotalSundays(cmbMonth.List(cmbMonth.ListIndex, 1), Val(cmbYear), Val(txtTotal))
    
    If chkExternal = vbChecked Then
        Hdays = Val(con.Execute("Select count(DT) from Holidays Where month(DT)=" & _
        cmbMonth.List(cmbMonth.ListIndex, 1) & " and Year(DT)=" & Val(cmbYear)).Fields(0))
        MonthDays = MonthDays '- Hdays
        txtTotalHrs = MonthDays * 8
        If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "Towne" Then
            txtTotal = "240"
            txtTotalHrs = "240"
            txtTotalHrs.Tag = "240"
        Else
            txtTotalHrs.Tag = MonthDays * 8
        End If
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

Private Sub PrintReport(PaySlips As Boolean, Optional Crpt As CrystalReport = Nothing, Optional Summary As Boolean, Optional bSalaryPacket As Boolean = False, Optional bNegative As Boolean = False, Optional iReportType As Integer = 0)

    If LV.ListItems.count = 0 Then Exit Sub
    Dim p_rptApp As New CRAXDDRT.Application
    Call InsertRecords
    
    Dim unloadflag As Boolean
    
    Cond = ""
    Screen.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt, rpt  As CRAXDDRT.Report
        
        
    If PaySlips Then
        If bSalaryPacket Then
            Set rpt = p_rptApp.OpenReport(rptPayrollPath & "\rptSalaryPacket.rpt")
            Cond = "{VPaySlip.BankPymt}=False"
        Else
            If iReportType = 0 Then
                Set rpt = p_rptApp.OpenReport(rptPayrollPath & "\rptPayslip.rpt")
            ElseIf iReportType = 1 Then
                Set rpt = p_rptApp.OpenReport(rptPayrollPath & "\rptPayslip_Rule9.rpt")
            ElseIf iReportType = 2 Then
                Set rpt = p_rptApp.OpenReport(rptPayrollPath & "\rptPayslip_Mini.rpt")
            ElseIf iReportType = 3 Then
                Set rpt = p_rptApp.OpenReport(rptPayrollPath & "\rptPayslip_Full.rpt")
            End If
        End If
    Else
        If Summary Then
            Set rpt = p_rptApp.OpenReport(rptPayrollPath & "\rptSalarySummary.rpt")
        Else
            If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                Set rpt = p_rptApp.OpenReport(rptPayrollPath & "\SalarySheet.rpt")
            ElseIf strCompany = "Tecno" Then
                Set rpt = p_rptApp.OpenReport(rptPayrollPath & "\rptSalarySheet.rpt")
            Else
                If cmdUpdate.Visible Then
                    Set rpt = p_rptApp.OpenReport(rptPayrollPath & "\rptSalarySheetUnFinal.rpt")
                Else
                    Set rpt = p_rptApp.OpenReport(rptPayrollPath & "\rptSalarySheet.rpt")
                End If
            End If
        End If
    End If
    
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
    
    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & strCompany & "'"
        ElseIf FormulaField.Name = "{@FromTo}" Then
            FormulaField.Text = "' For " & cmbMonth.Tag & "-" & cmbYear.Tag & "'"
        ElseIf FormulaField.Name = "{@MonthDays}" Then
            FormulaField.Text = Val(txtTotalHrs)
        End If
    Next
    If strCompany <> "Dr-Frgz" And strCompany <> "Banzai" Then
        If PaySlips = False Then
            Dim rptSub As CRAXDDRT.Report
            Set rptSub = rpt.OpenSubreport("rptSubSal.rpt")
            rptSub.FormulaFields.GetItemByName("MonthDays").Text = Val(txtTotalHrs)
        End If
    Else
            
        Dim DTStart As Date, DTFinish As Date
        Dim iTotalDays As Integer, iTotalSundays As Integer, iTotalHolidays As Integer, iNetDays As Integer
        DTStart = DateSerial(cmbYear, cmbMonth.ListIndex + 1, 1)
        DTFinish = DateSerial(cmbYear, cmbMonth.ListIndex + 2, 0)
        iTotalDays = Day(DTFinish)
        rpt.FormulaFields.GetItemByName("MonthDays").Text = iTotalDays
        iTotalSundays = TotalSundays(cmbMonth.ListIndex + 1, cmbYear.Text)
        
        iTotalHolidays = GetSingleLongValue("COUNT(DT)", "Holidays", " WHERE DT BETWEEN '" & DTStart & "' AND '" & DTFinish & "'")
        iNetDays = iTotalDays - iTotalSundays - iTotalHolidays
        
        rpt.FormulaFields.GetItemByName("FromTo").Text = "' For " & cmbMonth.Tag & "-" & cmbYear.Tag & "'"
        rpt.FormulaFields.GetItemByName("TotalMonthDays").Text = iTotalDays
        rpt.FormulaFields.GetItemByName("TotalSundays").Text = iTotalSundays
        rpt.FormulaFields.GetItemByName("TotalHolidays").Text = iTotalHolidays
        
    End If
    If bNegative Then
        Cond = "{@Balance}<0"
    End If
    f.ShowReport Cond & " ", rpt, ""

    Screen.MousePointer = vbDefault

End Sub

Private Sub PrintInternalReport(PaySlips As Boolean, Optional Crpt As CrystalReport = Nothing)

    If LV.ListItems.count = 0 Then Exit Sub
    
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
    '          .Formulas(1) = "Company='" & strCompany & "'"
    '        Else
    '          .ReportFileName = RptPath & "\SalarySheet.rpt"
    '          .Formulas(1) = ""
    '          .Formulas(1) = "Company='" & strCompany & "'"
    '        End If
    '      .Formulas(0) = "Fromto='For " & cmbMonth.Tag & "-" & cmbYear.Tag & IIf(chkExternal = vbChecked, "", " ( For Factory Use Only )") & "'"
    '      .Formulas(1) = "Company='" & strCompany & "'"
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

    Dim f As New frmPrevRpt, rpt  As New CRAXDDRT.Report
    
    If PaySlips Then
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptPayslipInternal.rpt")
    Else
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptSalarySheetInternal.rpt")
    End If
    
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
    
    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & strCompany & "'"
        ElseIf FormulaField.Name = "{@FromTo}" Then
            FormulaField.Text = "' For " & cmbMonth.Tag & "-" & cmbYear.Tag & "'"
        ElseIf FormulaField.Name = "{@MonthDays}" Then
            'FormulaField.Text = MonthDays
            FormulaField.Text = Val(txtTotal)
        End If
    Next
    Cond = ""
'    On Error Resume Next
'    SrcReport.Database.Tables(1).ConnectBufferString = "DSN=DSN_Vision;;User ID=sa;;Database=Vision;;Password=visionasim"
'    rpt.SQLQueryString = SQL
    f.ShowReport Cond & " ", rpt


    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdPrint_Click()


    mnuSalSheet.Visible = (chkExternal.Value = vbChecked)
    mnuPayslips.Visible = (chkExternal.Value = vbChecked)
    
    mnuSalarysheetInternal.Visible = (chkExternal.Value = vbUnchecked)
    mnuPayslipsInternal.Visible = (chkExternal.Value = vbUnchecked)
    
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
        'Check for Negative Salaries......
        If getDBPassword("MakeSalarySheetFinal") = False Then
            Exit Sub
        End If
        
        Dim i As Integer
        For i = 1 To LV.ListItems.count
            With LV.ListItems(i)
                If Val(LV.ListItems(i).SubItems(27)) < 0 Then
                    If MsgBox("Negative Salary of " & .Text & vbNewLine & " Do you want to Force Salary Finalyzation?", vbQuestion + vbYesNo) = vbYes Then
                        If getDBPassword("MakeSalarySheetFinal_Forced") = False Then
                            Exit Sub
                        Else
                            Exit For
                        End If
                    Else
                        Exit Sub
                    End If
                    
                End If
            End With
        Next
        
        With rsLedger
            If .RecordCount = 0 Then Exit Sub Else .MoveFirst
   
            con.BeginTrans
            
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
    Dim VchrNo As String, theDesc As String, AccNo As String
    Dim Dated As Date
    Dim dDRTaxDiff As Double
    
    With rs
        '.Open "Select Departments.Name,Departments.AccruedAccNo,Departments.SalaryAccNo,Departments.AccruedFoodAccNo,Departments.AccruedUnionAccNo,Departments.SchAccNo,Departments.AccruedEOBIAccNo,Departments.TaxAccNo, " & _
                "T1.DeptID,T1.TotalAmt,T1.SchAmt,T1.LunchAmt,T1.UnionAmt,T1.BalanceAmt,T1.ShortAmt,T1.LongAmt,T1.EOBIAmt,T1.TaxAmt " & _
                "From (Select DeptID,Sum(SAmt) As TotalAmt,Sum(Bonus) As SchAmt,Sum(Lunch) As LunchAmt,Sum(UnionFund) As UnionAmt,Sum(Balance) As BalanceAmt,Sum(ShortTerm) As ShortAmt,Sum(LongTerm) As LongAmt,Sum(EOBI) As EOBIAmt,Sum(Tax) As TaxAmt From VMonthlySalaries " & _
                "Where Month(DT)=" & cmbMonth.ListIndex + 1 & " and year(DT)=" & cmbYear & " " & _
                "Group By DeptID) T1 Inner JOIN Departments ON T1.DeptID=Departments.DeptID " & _
                "Where IsNull(Departments.AccruedAccNo,'')<>'' AND IsNull(Departments.SalaryAccNo,'')<>'' AND IsNull(Departments.AccruedFoodAccNo,'')<>'' AND IsNull(Departments.AccruedUnionAccNo,'')<>''  AND IsNull(Departments.SchAccNo,'')<>'' AND IsNull(Departments.AccruedEOBIAccNo,'')<>'' AND IsNull(Departments.TaxAccNo,'')<>''", con, adOpenForwardOnly, adLockReadOnly
        
        .Open "SELECT Departments.Name,Departments.AccruedAccNo,Departments.SalaryAccNo,Departments.AccruedFoodAccNo,Departments.AccruedUnionAccNo,Departments.SchAccNo,Departments.AccruedEOBIAccNo,Departments.TaxAccNo,Departments.OnePercentAccNo " & _
                ",T1.DeptID,T1.TotalAmt,T1.SchAmt,T1.LunchAmt,T1.UnionAmt,T1.BalanceAmt,T1.ShortAmt,T1.LongAmt,T1.EOBIAmt,T1.TaxAmt " & _
                ",T1.OTAmt,(DedOnePercentAmt) AS DedOnePercentAmt,T1.ZeroAbsentBonus,T1.GPHrsAmt From (SELECT DeptID,SUM(Total-OAmt-ZeroAbsentBonus-GPHrsAmt) AS TotalAmt,SUM(Bonus) AS SchAmt,SUM(Lunch) AS LunchAmt,SUM(UnionFund) AS UnionAmt,SUM(Balance) AS BalanceAmt,SUM(ShortTerm) AS ShortAmt,SUM(LongTerm) AS LongAmt,SUM(EOBI) AS EOBIAmt,SUM(Tax) AS TaxAmt,SUM(OAmt) AS OTAmt,SUM(Fine) AS DedOnePercentAmt,SUM(ZeroAbsentBonus) AS ZeroAbsentBonus,SUM(GPHrsAmt) AS GPHrsAmt FROM VMonthlySalaries " & _
                "WHERE Month(DT)=" & cmbMonth.ListIndex + 1 & " and year(DT)=" & cmbYear & " " & _
                "GROUP BY DeptID) T1 INNER JOIN Departments ON T1.DeptID=Departments.DeptID " & _
                "WHERE ISNULL(Departments.AccruedAccNo,'')<>'' AND ISNULL(Departments.SalaryAccNo,'')<>'' AND ISNULL(Departments.SchAccNo,'')<>'' AND ISNULL(Departments.OnePercentAccNo,'')<>''", con, adOpenForwardOnly, adLockReadOnly
                
        Do Until .EOF
        
            Dated = DateSerial(cmbYear, cmbMonth.ListIndex + 2, 0) 'Date
            
            If Format(Dated, "ddd") = "Sun" Then
                Dated = DateAdd("d", -1, Dated)
            End If
            
            INextSno = getNextSno(Dated)
            VchrNo = getNextVchrNo(Dated, "JV")
            
            '''''''''''''''''' T A X   E N T R I E S '''''''''''''
            If Val(!TaxAmt & "") > 0 Then
                theDesc = "Tax Of [" & !DeptID & "] " & !Name
                AccNo = !TaxAccNo & ""
                Bal = getBalance(AccNo, Dated)
                Amt = Val(!TaxAmt & "")
                
                If Amt <> 0 Then
                    Bal = Bal - Amt
                    
                    con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
                     Dated & "','" & VchrNo & "','" & AccNo & "','" & theDesc & _
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
    
            'If Amt = 0 Then Exit Function
            If Amt > 0 Then
              'Give The Description
              theDesc = "Salary Of  [" & !DeptID & "] " & !Name
              
              '''''''''''''''''' D E B I T   E N T R I E S '''''''''''''
              '''''''''''''''''' Salary Account '''''''''''''
              AccNo = !SalaryAccNo
              Amt = Val(!TotalAmt & "") '+ Val(!OTAmt & "")
              Bal = getBalance(AccNo, Dated)
        
              Amt = Amt
              Bal = Bal + Amt
              
              con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
                  Dated & "','" & VchrNo & "','" & AccNo & _
                  "','" & theDesc & "'," & Amt & ",0," & Bal & ",'',0)"
          
              Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
              INextSno = INextSno + 1
            End If
            ''''''''''''''''Over Time Account'''''''''''''
            ''DEBIT Entry
            AccNo = !AccruedUnionAccNo & ""
            Amt = Val(!OTAmt & "")
            If Amt <> 0 Then
                Bal = getBalance(AccNo, Dated)
    
                Bal = Bal + Amt
    
                con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
                    Dated & "','" & VchrNo & "','" & AccNo & _
                    "','" & theDesc & "'," & Amt & ",0," & Bal & ",'',0)"
    
                Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
    
                INextSno = INextSno + 1
            End If
'            '''''''''''''''''' Scholorship Account '''''''''''''
'
            AccNo = !SchAccNo
            Amt = Val(!ZeroAbsentBonus & "")
            If Amt <> 0 Then
              Bal = getBalance(AccNo, Dated)

              Amt = Amt
              Bal = Bal + Amt

              con.Execute "INSERT INTO Vouchers VALUES(" & INextSno & ",'" & _
                  Dated & "','" & VchrNo & "','" & AccNo & _
                  "','" & theDesc & "'," & Amt & ",0," & Bal & ",'',0)"

              Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
              INextSno = INextSno + 1
            End If
            '''''''''''''''''' C R E D I T   E N T R I E S '''''''''''''
            ''''''''''''''''Temp. Advance'''''''''''''
            
            AccNo = strEmpShortTermAccNo
            Amt = Val(!Shortamt & "")
            If Amt <> 0 Then
                Bal = getBalance(AccNo, Dated)
                
                Bal = Bal - Amt
            
                con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
                    Dated & "','" & VchrNo & "','" & AccNo & _
                    "','" & theDesc & "',0," & Amt & "," & Bal & ",'',0)"
        
                Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
         
                INextSno = INextSno + 1
            End If
            ''''''''''''''''Long Term Deduction'''''''''''''
            
            AccNo = strEmpLongTermAccNo '"13-003-12001" '"13-003-16001"
            Amt = Val(!LongAmt & "")
            If Amt <> 0 Then
                Bal = getBalance(AccNo, Dated)
    
                Bal = Bal - Amt
    
                con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
                    Dated & "','" & VchrNo & "','" & AccNo & _
                    "','" & theDesc & "',0," & Amt & "," & Bal & ",'',0)"
    
                Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
    
                INextSno = INextSno + 1
            End If
            
'            ''''''''''''''''Food Amount Deduction'''''''''''''
'
'            AccNo = !AccruedFoodAccNo & ""
'            Amt = Val(!LunchAmt & "")
'            If Amt <> 0 Then
'                Bal = getBalance(AccNo, Dated)
'
'                Bal = Bal - Amt
'
'                con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
'                    Dated & "','" & VchrNo & "','" & AccNo & _
'                    "','" & TheDesc & "',0," & Amt & "," & Bal & ",'',0)"
'
'                Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
'
'                INextSno = INextSno + 1
'            End If
'            ''''''''''''''''Food Amount Deduction'''''''''''''
'
'            AccNo = !AccruedUnionAccNo & ""
'            Amt = Val(!UnionAmt & "")
'            If Amt <> 0 Then
'                Bal = getBalance(AccNo, Dated)
'
'                Bal = Bal - Amt
'
'                con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
'                    Dated & "','" & VchrNo & "','" & AccNo & _
'                    "','" & TheDesc & "',0," & Amt & "," & Bal & ",'',0)"
'
'                Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
'
'                INextSno = INextSno + 1
'            End If
            ''''''''''''''''Balance Amount Deduction'''''''''''''
            'If There's Ded. Of Tax. Then Adjust The Difference Of Rounding Of Tax Amount.
            
            AccNo = !AccruedAccNo & ""
            Amt = Val(!BalanceAmt & "")
            

            dDRTaxDiff = Val(!TotalAmt & "") + Val(!OTAmt & "") + Val(!SchAmt & "") - Amt - Val(!LunchAmt & "") - Val(!UnionAmt & "") - Val(!Shortamt & "") - Val(!LongAmt & "") - Val(!EOBIAmt & "") - Val(!TaxAmt & "") + Val(!DedOnePercentAmt & "") + Val(!ZeroAbsentBonus & "")
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
        
            con.Execute "INSERT INTO Vouchers VALUES(" & INextSno & ",'" & _
                Dated & "','" & VchrNo & "','" & AccNo & _
                "','" & theDesc & "',0," & Amt & "," & Bal & ",'',0)"
    
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
                    "','" & theDesc & "',0," & Amt & "," & Bal & ",'',0)"
        
                Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
         
                INextSno = INextSno + 1
            End If
                        
            AccNo = !OnePercentAccNo & ""
            Amt = Val(!DedOnePercentAmt & "")
            If Amt < 0 Then 'Credit
            
                Amt = Abs(Amt)
                Bal = getBalance(AccNo, Dated)
                
                Bal = Bal - Amt
            
                con.Execute "INSERT INTO Vouchers VALUES(" & INextSno & ",'" & _
                    Dated & "','" & VchrNo & "','" & AccNo & _
                    "','" & theDesc & "',0," & Amt & "," & Bal & ",'',0)"
        
                Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
         
                INextSno = INextSno + 1
                        
            ElseIf Amt > 0 Then 'Debit
                Bal = getBalance(AccNo, Dated)
                
                Bal = Bal + Amt
            
                con.Execute "INSERT INTO Vouchers VALUES(" & INextSno & ",'" & _
                    Dated & "','" & VchrNo & "','" & AccNo & _
                    "','" & theDesc & "'," & Amt & ",0," & Bal & ",'',0)"
        
                Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
         
                INextSno = INextSno + 1
            End If
            
            con.Execute "INSERT INTO VoucherInfo(VchrNo,UserName,MachineName,DT) VALUES('" & VchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
                
            .MoveNEXT
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

    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = True
    Next
    
End Sub

Private Sub Command2_Click()
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = False
    Next
End Sub

Private Sub Form_Load()
            
    If strCompany = "Banzai" Then
        mnuPrintNegativeEmpList.Visible = True
    Else
        mnuPrintNegativeEmpList.Visible = False
    End If
    
    m_bShowFake = False
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
        cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' - ' + FName + ' {' + Designation + '}'", "VSalariedEmps", "EmpID", " Where JOINDate<='" & LastDT & "' AND (DiscontinuedOn>'" & FirstDT & "' OR DiscontinuedOn Is Null)"
    Else
        cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' - ' + FName + ' {' + Designation + '}'", "VSalariedEmps", "EmpID", "Where DeptID='" & cmbDept.ID & "' and JOINDate<='" & LastDT & "' AND (DiscontinuedOn>'" & FirstDT & "' OR DiscontinuedOn Is Null)"
    End If
    cmbEmp.ID = "0"
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    Dim TopMargin As Integer
    TopMargin = cmbEmp.Top + cmbEmp.Height + 30
    
    FLV.Move 50, TopMargin, Me.ScaleWidth - 100, Me.ScaleHeight - TopMargin
    
    LV.Move 100, 200, Me.Width - 400, FLV.Height - fButs1.Height - 300
    
    fButs1.Move LV.Left + LV.Width - fButs1.Width, LV.Top + LV.Height
    
    FButs2.Move LV.Left, LV.Top + LV.Height
    lblTransferNegative.Move fButs1.Left - lblTransferNegative.Width - 50, fButs1.Top + fButs1.Height - lblTransferNegative.Height
    
End Sub

Private Sub RefreshLedger()

    Dim iRounding As Integer
    iRounding = 1
    
    Dim bFoundNegatives As Boolean
    bFoundNegatives = False
    Dim lColor As Long
    
    Dim rs As New ADODB.Recordset
    Dim rsInternal As New ADODB.Recordset
    
    With rs
        .Open "SELECT COUNT(EmpID) FROM MonthlySalaries WHERE Month(DT)=" & cmbMonth.ListIndex + 1 & " AND Year(DT)=" & cmbYear.Text & " AND IsForSA=0", con, adOpenForwardOnly, adLockReadOnly
        If Val(.Fields(0) & "") = 0 Then
            IsInternalFinal = False
        Else
            IsInternalFinal = True
            'To Save Long term Loan Information....
            rsInternal.Open "SELECT EmpID,LongTerm,PrevLTLoan FROM MonthlySalaries WHERE Month(DT)=" & cmbMonth.ListIndex + 1 & " AND Year(DT)=" & cmbYear.Text & " AND IsForSA=0", con, adOpenStatic, adLockReadOnly
        End If
        .Close
    End With
    
    Dim TotalSalDays As Double
    Dim TDate As Date
    Dim SDate As Date 'For Saving
    Dim Hdays As Long
    
    SDate = DateSerial(cmbYear, cmbMonth.ListIndex + 1, 1)
    
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(DT)", "FullMonthSalaryCalculations", " WHERE MONTH(DT)=" & cmbMonth.ListIndex + 1 & " AND YEAR(DT)=" & cmbYear.Text)
    If lCount > 0 Then
        TDate = DateSerial(cmbYear, cmbMonth.ListIndex + 2, 0)
    Else
        TDate = DateSerial(cmbYear, cmbMonth.ListIndex + 1, Day(GetServerDate))
    End If
    
    TotalSalDays = DateAdd("m", 1, SDate) - SDate

    MonthDays = TotalSalDays - TotalSundays(cmbMonth.List(cmbMonth.ListIndex, 1), cmbYear, TotalSalDays)
    
    If chkExternal = vbUnchecked Then
        Hdays = Val(con.Execute("Select count(DT) from Holidays Where month(DT)=" & _
        cmbMonth.List(cmbMonth.ListIndex, 1) & " and Year(DT)=" & Val(cmbYear)).Fields(0))
        MonthDays = MonthDays - Hdays
    End If
   
    bSalaryFinalyzed = FinalizedSalarySheetLoaded
    If bSalaryFinalyzed Then
        Exit Sub
    Else
        If m_bShowFake Then
            MsgBox "Please finalize Original Salary Sheet.", vbInformation
            Exit Sub
        End If
    End If
    
    txtTotalHrs.Tag = txtTotalHrs

    Dim UnionFund As Integer
    Dim GrandTotal As Double
    Dim EOBIAmt As Byte
    If cmbDept.MatchFound = False Or cmbEmp.MatchFound = False Then Exit Sub
    If cmbYear.MatchFound = False Or cmbMonth.MatchFound = False Then Exit Sub
    
    Dim SunDates() As Date
        
    Dim EmpId As String, DeptID As String
    Dim EmpAge As Double
    
    Dim TempTotalSaldays As Double
    Dim LastDay As Integer
    
    Dim ITM As ListItem
    Dim Cond As String
    
    Dim cmd As New ADODB.Command
    
    Dim SalHrs As Double
    Dim HrsMultiple As Double
    
    Dim Leaves As Double
    Dim ApprovedLeaves As Double
    Dim UnApprovedLeaves As Double
    
    
'    If cmbEmp.ID = "0" Then
'       EmpId = IIf(cmbDept.ID = "0", "", cmbDept.ID)
'    Else
'       EmpId = IIf(cmbEmp.ID = "0", "", cmbEmp.ID)
'    End If
    
    EmpId = cmbEmp.ID
    DeptID = cmbDept.ID
    
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
        'MsgBox "Can't Show Entries Of Future Months.", vbInformation
        'Exit Sub
    End If
    
    TempTotalSaldays = TotalSalDays     'Store For Later Use
    
    JDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), TotalSalDays)
        
    'Sql = "Select * from VAdvanceSalary " & Cond & " JoinDate<='" & Format(JDate, "dd-MMM-yyyy") & "' Order By DeptID,EmpID"
    
    cmd.CommandType = adCmdStoredProc
    
    If DeptID <> "0" And EmpId = "0" Then
        cmd.CommandText = "VAdvanceSalaryDeptwise"
    Else
        cmd.CommandText = "VAdvanceSalary"
    End If
    
    Set cmd.ActiveConnection = con
    
    cmd.Parameters("@ParDate").Value = JDate
    cmd.CommandTimeout = 120
    If DeptID <> "0" And EmpId = "0" Then
        cmd.Parameters("@DeptID").Value = DeptID
    Else
        cmd.Parameters("@EmpID").Value = IIf(EmpId = "0", "%", EmpId)
    End If
    
    If rsLedger.State = 1 Then rsLedger.Close
    rsLedger.CursorType = adOpenForwardOnly
    rsLedger.LockType = adLockReadOnly
    
    Set rsLedger = cmd.Execute
    
    With rsLedger
        
        '.Open Sql, con, adOpenForwardOnly, adLockReadOnly
        Dim TotalSuns As Integer
        Dim SalDays As Double
        Dim Rate As Double, FakeRate As Double
        Dim AbsAllowed As Integer
        Dim AbsentDays As Double
        
        Dim total As Double
        Dim OverTimeHours As Double
        Dim OverTimeHoursSA As Double
        Dim OverTime_Food_Allowance As Long
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
        
        Dim dLateComingHrs As Double, dOTHrs As Double, dShortHrs As Double, dShortAmt As Double, dGPHrs As Double, dGPHrsAmt As Double
        Dim lOTDinnerCount As Long, lOTDinnerAmt As Long, dOnePercentDeduction As Double
        
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
               
               rs.Open "SELECT count(DT) FROM Holidays WHERE DT Between '" & ![JoinDate] & "' AND '" & JDate & "'", con, adOpenForwardOnly, adLockReadOnly
               
                  If rs.RecordCount > 0 Then
                     Holidays = rs.Fields(0) & ""
                  End If
               rs.Close
               Set rs = Nothing
               
               SundayAbsents = getSundayAbsents(![EmpId] & "", Val(cmbMonth.List(cmbMonth.ListIndex, 1)), cmbYear, Val(![WeeklyAbs] & ""), Int(TotalSalDays))
               TotalSuns = TotalSundays(cmbMonth.List(cmbMonth.ListIndex, 1), cmbYear, TotalSalDays, Day(![JoinDate]))
               Holidays = TotalSuns + Holidays - SundayAbsents
               
            Else
            
                 rs.Open "SELECT count(DT) FROM Holidays WHERE CONVERT(CHAR(3),DT,7)='" & cmbMonth & "' AND Year(DT)=" & cmbYear & " AND day(dt)<=" & TotalSalDays, con, adOpenForwardOnly, adLockReadOnly
                  
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
                If IsNull(![totamt]) Then lTotAmt = 0 Else lTotAmt = ![totamt]
                If IsNull(![ClsAmt]) Then lClsAmt = 0 Else lClsAmt = ![ClsAmt]
                
                
                If IsNull(!AmtToDeduct) Then
                    If lDedAmt > (lTotAmt - lClsAmt) Then
                        LongLoan = (lTotAmt - lClsAmt)
                    Else
                        LongLoan = lDedAmt
                    End If
                Else
                    LongLoan = Val(!AmtToDeduct & "")
                End If
                                      
                'Get The Salary Depending On Salary Sheet Type
                'EmpSalary = IIf(chkExternal = vbUnchecked, Val(![StartingSalary] & ""), Val(![Salary] & ""))
                
                 EmpSalary = IIf(m_bShowFake, Val(![EmpMinSalary] & ""), Val(![StartingSalary] & ""))
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
                        
            'OverTime_Food_Allowance = (Floor(OverTimeHours) * Val(!EmpLunchAmt & "")) 'Rs. (Rate from Profile) for each hour
            OverTime_Food_Allowance = Val(!Food_Allowance & "")
            If IsComputed Then
                theKey = ![EmpId] & Format(SDate, "dd-MMM-yyyy")
            Else
                theKey = ![EmpId] & ""
            End If
            
            Set ITM = LV.ListItems.add(, theKey, ![EmpId] & "")
            
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
            ITM.ListSubItems.add(, , ![Name] & "").Tag = Val(![totamt] & "") - Val(![ClsAmt] & "")
            
            ITM.ListSubItems.add , , IIf(chkExternal = vbChecked, EmpExtSal, EmpSalary & "")
            
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
                                
                'Rate = Format(Val(EmpSalary) / (LastDay - TotalSuns), "0.00")
                
                Rate = Format(Val(EmpSalary) / (30), "0.00")
                If HrsMultiple = 0 Then HrsMultiple = 8
            Else
                
                Rate = Format(Val(EmpSalary) / LastDay, "0.00")
                
                HrsMultiple = 1
                
            End If

            Rate = Rate / HrsMultiple
                       
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '   Rate Formula For October Salary Sheet ...
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            If chkExternal Then
                If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "Towne" Then
                    Rate = EmpSalary / 240
                Else
                    Rate = EmpSalary / 240 'Val(txtTotalHrs)
                End If
                FakeRate = Val(!EmpMinSalary & "") / Val(txtTotalHrs)
            End If
                       
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        
            'Store The Total Working hours for This Employee
            ITM.ListSubItems.add(, , Format(Rate, "0.00")).Tag = HrsMultiple
            
            If IsNull(![presentDays]) And !Pay_Full_Salary = False Then
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
               If Day(!JoinDate) = 5 Then
                    DaysToExclude = Day(![JoinDate]) - 1
                Else
                    DaysToExclude = Day(![JoinDate]) - 1
                End If
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
            
            ITM.ListSubItems.add , , AbsentDays * HrsMultiple
            
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
                        
            Dim iMonthDays As Integer
            iMonthDays = Day(DateSerial(cmbYear, cmbMonth.ListIndex + 2, 0))
            If chkExternal Then
'                If ![DeptID] = "VDRS" Then
                If !Pay_Full_Salary Then
                    SalHrs = Val(txtTotalHrs)
                Else
             
                    'SalHrs = Round(Val(![HrsWorked] & ""), 1) + ((Holidays - TotalSuns) * 8)
                    'SalHrs = Round(Val(![HrsWorked] & ""), 2) + ((Holidays) * 8)
                    SalHrs = Round(Val(![HrsWorked] & ""), 2) + ((Holidays) * 8)
                    ' Add All the Leaves to Worked hours
                    SalHrs = SalHrs + (Val(![Leaves] & "") * 8)
                    
                    'Remove Full Hours of Without Pay leaves
                    SalHrs = SalHrs - (Val(![WPLeaves] & "") * 8)
                    
                    ' Remove half Hours of Sick Leaves
                    'SalHrs = SalHrs - (Val(![SickLeaves] & "") * 4)
                    
                    ' Check if Salary Hours are greater than Total Hours Due to relaxation
                    ' Fix them to maximum hours
                    SalHrs = SalHrs - Val(!ExtraHrsToDeduct & "")
                    If strCompany <> "Dr-Frgz" Then
                        'If SalHrs > Val(txtTotalHrs) Then SalHrs = Val(txtTotalHrs)
                        If SalHrs > (iMonthDays * 8) Then SalHrs = (iMonthDays * 8)
                    End If
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
            
            If !DeductionSpecialAbsents Then
                SalHrs = SalHrs - (Val(!SPAbsentDays & "") * 8)
            End If
            If SalHrs < 0 Then SalHrs = 0
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '       Show in Time Format if external salary sheet
            '       and Show days in internal salarysheet instead
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            ITM.SubItems(8) = IIf(chkExternal, getHrsMin(SalHrs), SalHrs)
            
            dLateComingHrs = Round(Val(!LateComingHrs & "") - Val(!LateComingHrs_Adjusted & ""), 4)
            
            'SalHrs = SalHrs + Val(!LateComingHrs_Adjusted & "") - Val(!GPHrs & "")
            
            If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                'If !EmpId = "DFG-000014" Then Stop
                Dim dAbsentHrs As Double, dDailyRate As Double, dAbsentDays As Double, dPresentDays As Double, dLateHrs As Double, dLateAmt As Double, lZeroAbsentBonusAmt As Long
                dAbsentHrs = Val(txtTotalHrs.Tag) - SalHrs
                If strCompany = "Kami" Then
                    dDailyRate = EmpSalary / iMonthDays
                Else
                    dDailyRate = EmpSalary / 30
                End If
                dPresentDays = Val(!presentDays & "") + Holidays + Leaves - Val(![WPLeaves] & "") - Val(!HalfLeaves & "")
                
                dAbsentDays = iMonthDays - dPresentDays
                
'                If dAbsentDays <= 0 And Floor(Leaves) = 0 Then
'                    If IsProbation(!EmpId & "", JDate) = False Then
'                        lZeroAbsentBonusAmt = dDailyRate
'                    Else
'                        lZeroAbsentBonusAmt = 0
'                    End If
'                Else
'                    lZeroAbsentBonusAmt = 0
'                End If
                lZeroAbsentBonusAmt = 0
                Dim dAbsDaysForSalCalc As Double
                dAbsDaysForSalCalc = dAbsentDays
                If !DeductionSpecialAbsents Then
                    dPresentDays = dPresentDays - Val(!SPAbsentDays & "")
                    dAbsDaysForSalCalc = dAbsDaysForSalCalc + Val(!SPAbsentDays & "")
                    'SalHrs = SalHrs - (Val(!SPAbsentDays & "") * 8)
                End If
                
                If strCompany = "Kami" Then
                    ITM.SubItems(9) = Round(dPresentDays * dDailyRate) 'Round((dPresentDays * dDailyRate) / 10, 0) * 10
                Else
                    If !Pay_Full_Salary Then
                        ITM.SubItems(9) = EmpSalary
                    Else
                        If dPresentDays > dAbsDaysForSalCalc Then
                            ITM.SubItems(9) = Round((EmpSalary - (dDailyRate * dAbsDaysForSalCalc)) / iRounding, 0) * iRounding   'Round((SalHrs * Rate) / 10, 0) * 10
                            dPresentDays = 30 - dAbsDaysForSalCalc
                        Else
                            ITM.SubItems(9) = Round((dPresentDays * dDailyRate) / iRounding, 0) * iRounding
                        End If
                    End If
                End If
                dLateHrs = (Val(iMonthDays * 8)) - SalHrs - (Val(![WPLeaves] & "") * 8)
                
                dLateHrs = dLateHrs - (dAbsentDays * 8)
                dLateHrs = dLateHrs - Val(!LateComingHrs_Adjusted & "")
                If !DeductionSpecialAbsents Then
                    dLateHrs = dLateHrs - (Val(!SPAbsentDays & "") * 8)
                End If
                
                If strCompany = "Dr-Frgz" Then
                    dShortHrs = (dLateHrs - dLateComingHrs) + Val(!GPHrs & "")
                Else
                    dShortHrs = (dLateHrs - dLateComingHrs)
                End If
                If dShortHrs < 0 Then dShortHrs = 0
                'If strCompany = "Dr-Frgz" Then
                    'Deduct Overtime according to Late/Short
'                    If OverTimeHoursSA > 0 Then
'                        If OverTimeHoursSA > dLateComingHrs Then
'                            OverTimeHoursSA = OverTimeHoursSA - dLateComingHrs
'                            dLateComingHrs = 0
'                        Else
'                            dLateComingHrs = dLateComingHrs - OverTimeHoursSA
'                            OverTimeHoursSA = 0
'                        End If
'                    End If
                'End If
                'dLateAmt = (dShortHrs + dLateComingHrs) * (Rate * (Val(!DeductionPercentage & "") / 100))
                If !Pay_Full_Salary Then
                    dLateComingHrs = 0
                    dShortHrs = 0
                    dLateHrs = 0
                End If
                dLateAmt = (dLateComingHrs) * (Rate * (Val(!DeductionPercentage & "") / 100))
                'dLateAmt = Round(dLateAmt / 10, 0) * 10
                dShortAmt = (dShortHrs) * (Rate)
                dLateAmt = Round(dLateAmt)
                dShortAmt = Round(dShortAmt)
            Else
                ITM.SubItems(9) = Round((SalHrs * Rate) / iRounding, 0) * iRounding
            End If
            
            dGPHrs = Val(!GPHrs & "")
            dGPHrsAmt = dGPHrs * Rate
            dGPHrsAmt = Round(dGPHrsAmt)
'            If strCompany = "Tecno" Then
'                dGPHrs = Val(!GPHrs & "")
'                dGPHrsAmt = dGPHrs * Rate
'            Else
'                dGPHrs = 0
'                dGPHrsAmt = 0
'            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '       Save the Salary Hours in Listsubitmes(4).tag To Save
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ITM.ListSubItems(4).Tag = SalHrs
            ITM.ListSubItems(5).Tag = Val(![EmpMinSalary] & "")
            ITM.ListSubItems(6).Tag = FakeRate
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            ITM.SubItems(10) = Leaves * HrsMultiple
            
            
            ' Do it only For Internal Salary Sheet
            If chkExternal = vbUnchecked Then
                LeaveDedAmt = Val(![SickLeaves] & "") * HrsMultiple * Rate / 2
                LeaveDedAmt = LeaveDedAmt + Val(![WPLeaves] & "") * HrsMultiple * Rate
                ITM.SubItems(11) = Round(LeaveDedAmt, 2)
            End If
            ITM.SubItems(11) = 0
           
            
            'If SalDays = 0 Then SalDays = 0 Else SalDays = (SalDays - AbsentDays)
            
            If OTByQty Then
                ITM.SubItems(12) = OverTimeHours
                'Over time By Quantity : Per Day Rate / Sheets Per Day * 1.4
                ITM.SubItems(13) = Round(OverTimeHours * (Rate * HrsMultiple / 56 * 1.4))
            Else
                '
                ITM.SubItems(12) = IIf(chkExternal, MinsToTime(OverTimeHoursSA * 60), MinsToTime(OverTimeHours * 60))
                ITM.ListSubItems(12).Tag = Val(!SundayOTHrs & "")
                If chkExternal Then
                    If strCompany = "QEL" Then
                        'ITM.SubItems(13) = Round(OverTimeHoursSA * (Rate + (Rate * 62.5 / 100)))
                        ITM.SubItems(13) = Round(OverTimeHoursSA * ((EmpSalary / 240) * 1.333333))
                    ElseIf strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                        ITM.SubItems(13) = Round((OverTimeHoursSA * (Rate * Val(!OTRate & ""))) / iRounding, 0) * iRounding
                    ElseIf strCompany = "Tecno" Then
                        ITM.SubItems(13) = Round(OverTimeHoursSA * (Rate * 1.5))
                    Else
                        ITM.SubItems(13) = Round(OverTimeHoursSA * (Rate * 1.333333))
                    End If
                Else
                    ITM.SubItems(13) = Round(OverTimeHours * (Rate / 8 * 1.333333))
                End If
                If strCompany = "PAKSMITH GROUP" Then
                    ITM.ListSubItems(13).Tag = Val(!SundayOTHrs & "") * (Rate * 1.333333)
                ElseIf strCompany = "QEL" Then
                    '2.5 Hours=4;5=8; I think it's 62.5 % Increase over Standard Rate
                    'ITM.ListSubItems(13).Tag = Val(!SundayOTHrs & "") * (Rate + (Rate * 62.5 / 100))
                    ITM.ListSubItems(13).Tag = Val(!SundayOTHrs & "") * ((EmpSalary / 240) * 1.333333)
                Else
                    ITM.ListSubItems(13).Tag = Val(!SundayOTHrs & "") * (Rate * Val(!OTRate & ""))
                End If
            End If
            
            LateHours = Val(![HOURS] & "") - LTAllowed
            
            If LateHours < 0 Then LateHours = 0
            
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''
            '           Exempt For October 2004
            ''''''''''''''''''''''''''''''''''''''''''''''''''
            'Well It was just written as LateHours=0 And That was why it was'nt deducting LT.
            'So I Just Commented it Out.
            'LateHours = 0
            ''''''''''''''''''''''''''''''''''''''''''''''''''
                        
            If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                ITM.SubItems(14) = MinsToTime(dLateHrs * 60)
            Else
                ITM.SubItems(14) = MinsToTime(LateHours * 60)
            End If
            
            ITM.ListSubItems(14).Tag = Val(!FixAllowance & "")
           ' itm.SubItems(15) = IIf((Exempt), "0", Format((LateHours * (Rate / Val(!WorkingHrs & ""))), "0.00"))
            
            If chkExternal Then
                If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                    ITM.SubItems(15) = dLateAmt
                Else
                    ITM.SubItems(15) = IIf((Exempt), "0", Format((LateHours * Rate), "0.00"))
                End If
            Else
                ITM.SubItems(15) = IIf((Exempt), "0", Format((LateHours * Rate / 8), "0.00"))
            End If
            
            'Total = Format(Val((ITM.SubItems(9)) + Val(ITM.ListSubItems(7))) - (Val(ITM.SubItems(15)) + Val(ITM.SubItems(11))), "0.00")
            If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                total = Format(Val((ITM.SubItems(9)) + Val(ITM.ListSubItems(7)) + Val(ITM.ListSubItems(13)) + Val(ITM.ListSubItems(13).Tag) + Val(ITM.ListSubItems(14).Tag)) + lZeroAbsentBonusAmt - Val(ITM.SubItems(15)) - dShortAmt - (Val(ITM.SubItems(11))), "0.00")
                total = total + OverTime_Food_Allowance
                total = Round(total / iRounding) * iRounding
            Else
                total = Format(Val((ITM.SubItems(9)) + Val(ITM.ListSubItems(7)) + Val(ITM.ListSubItems(13)) + Val(ITM.ListSubItems(13).Tag) + Val(ITM.ListSubItems(14).Tag)) - (Val(ITM.SubItems(11))), "0.00")
            End If
            
            ITM.SubItems(16) = total
            
            'TaxPer = Val(![PerDeduct])
            ITM.SubItems(17) = Val(!TaxAmt & "") 'Round(getTaxAmt((EmpSalary * 12) / 1.5) / 12, 0)
            'DedAmt = Round((Val(![Salary]) / TotalSalDays) * (TaxPer / 100))
            'TaxSal = Round((Val(EmpSalary) / LastDay) * TotalSalDays)
            DedAmt = 0 'Round(TaxSal * (TaxPer / 100))
            
            'Tax Percentage That Is Zero Length
            'itm.SubItems(18) = "0"
            
            'itm.SubItems(14) = DedAmt
            ITM.SubItems(18) = total 'Total - Val(ITM.SubItems(17))
            If IsNull(![Payed]) Then Payed = 0 Else Payed = ![Payed]
            ITM.SubItems(19) = Payed 'Round(Payed / 10) * 10
            
            If chkExternal.Value = vbChecked And IsInternalFinal Then
                If rsInternal.EOF Then
                    ITM.SubItems(20) = 0
                Else
                    ITM.SubItems(20) = Val(rsInternal![LongTerm] & "")
                End If
            Else
                ITM.SubItems(20) = LongLoan
                ITM.ListSubItems(20).Tag = (lTotAmt - lClsAmt)
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
                    LunchDed = "0" 'Val(![LunchAmt] & "")
                Else
                    LunchDed = 0
                End If
            ''''''''''''''''''''''''''''''''''''''''''
'            If strCompany = "IAA" Then
'                dOnePercentDeduction = Round(Val(!StartingSalary & "") / 100)
'            Else
'                dOnePercentDeduction = 0
'            End If
            dOnePercentDeduction = 0
            ITM.SubItems(22) = UnionFund
            ITM.SubItems(23) = Val(!TotalFine & "")
            
            ' Use This Column For Scholorships Instead Of Bonus ...
            'ITM.SubItems(24) = Val(!TotalBonus & "")
            
            ITM.SubItems(24) = "0" 'Val(![ScholorshipAmt] & "")
            ITM.SubItems(25) = LunchDed
            
            'If Not IsNull(![OldAge]) Then If ![OldAge] Then EOBIAmt = 130 Else EOBIAmt = 0
            EOBIAmt = 0
'            EmpAge = DateDiff("M", !DOB, JDate) / 12
'
'            If ![DeptID] = "VDRS" Or EmpAge > 60 Or !ArmyRetired = True Then
'                EOBIAmt = 0
'            Else
'                EOBIAmt = 40
'            End If
            '/\/\//\/\ IF U CHANGE EOBI FORMULA HERE, REMEMBER TO CHANGE IN rptSalarySheet and rptPaySlip TOO./\/\/\/\/\
            
            ITM.SubItems(26) = EOBIAmt
            
            lOTDinnerCount = Val(!DinnerCount & "")
            
'            If strCompany = "QEL" Then
'                lOTDinnerAmt = lOTDinnerCount * 25
'            Else
'                lOTDinnerAmt = 0
'            End If
            lOTDinnerAmt = OverTime_Food_Allowance
            'Removed OverTime From Balance
            ' ITM.SubItems(27) = Round((Val(ITM.SubItems(18)) + Val(!ScholorshipAmt & "") _
            - (Payed + Val(!TotalFine & "") + UnionFund + LongLoan + LunchDed + Val(!AdvSalpaid & "") + Val(ITM.SubItems(26))) _
            + Val(ITM.SubItems(13))))
            
            If strCompany = "Dr-Frgz" Then
                ITM.SubItems(27) = Round((Val(ITM.SubItems(18)) _
                 - (Payed + UnionFund + LongLoan + LunchDed + Val(!AdvSalpaid & "") + Val(ITM.SubItems(17)) + Val(ITM.SubItems(26)))) / iRounding, 0) * iRounding
            Else
                ITM.SubItems(27) = Round((Val(ITM.SubItems(18)) + Val(ITM.SubItems(23)) _
                 - (dOnePercentDeduction + dGPHrsAmt + Payed + UnionFund + LongLoan + LunchDed + Val(!AdvSalpaid & "") + Val(ITM.SubItems(17)) + Val(ITM.SubItems(26)))) / iRounding, 0) * iRounding
            End If
            
            ITM.ListSubItems(7).Tag = lOTDinnerCount
            ITM.ListSubItems(8).Tag = lOTDinnerAmt
            
            ITM.ListSubItems(9).Tag = dGPHrs
            ITM.ListSubItems(10).Tag = dGPHrsAmt
            '--------------------------------------------------------
            ITM.ListSubItems(11).Tag = Val(!AmtPaid & "")
            ITM.ListSubItems(15).Tag = dLateComingHrs
            
            ITM.ListSubItems(16).Tag = dShortHrs
            ITM.ListSubItems(17).Tag = dShortAmt
                        
            ITM.ListSubItems(18).Tag = lZeroAbsentBonusAmt
            ITM.ListSubItems(19).Tag = Val(!HL & "")
            
            
            ITM.ListSubItems(23).Tag = dOnePercentDeduction
            ITM.ListSubItems(24).Tag = !DeptID & ""
            ITM.ListSubItems(25).Tag = Val(!Leaves & "") - Val(![WPLeaves] & "")
            ITM.ListSubItems(26).Tag = dPresentDays - Val(!Leaves & "")  'Val(!PresentDays & "") + Holidays
            
            ITM.SubItems(28) = Val(!OTHrs_Actual & "")
            ITM.ListSubItems(28).Tag = Val(!OverTimeHours & "")
            ITM.SubItems(29) = Val(!LateComingHrs_Actual & "")
            ITM.ListSubItems(29).Tag = Val(!LateComingHrs_Actual & "") - Val(!LateComingHrs_Adjusted & "")
            
            If Val(ITM.SubItems(27)) < 0 Then
                lColor = vbRed
                bFoundNegatives = True
            Else
                lColor = LV.ForeColor
            End If
            
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
            
            Dim iCounter As Integer
            
            For iCounter = 1 To ITM.ListSubItems.count
                ITM.ForeColor = lColor
                ITM.ListSubItems(iCounter).ForeColor = lColor
            Next
            ITM.ListSubItems(27).Tag = Val(!HoldAmt & "")
NEXT_EMPLOYEE:
            .MoveNEXT
        Next
        
        LV.Visible = True
        
        lblTotal = Format(GrandTotal, "#,##0.00")
        
    End With
    cmbMonth.Tag = cmbMonth.Text
    cmbYear.Tag = cmbYear.Text
    
    SQL = ""
    Cond = ""
    ICO = ""
    
    If cmbDept.ID = "0" And cmbEmp.ID = "0" Then
        cmdUpdate.Caption = "Make This Sheet Final"
        cmdUpdate.Visible = UserHasAccess("FinalSalSheet")
    Else
        cmdUpdate.Visible = False
    End If
    If bFoundNegatives Then
        lblTransferNegative.Visible = True
    Else
        lblTransferNegative.Visible = False
    End If
    
End Sub

Private Function GetLateHours(EmpId As String, NTime As Integer, LTime As Double, ToDate As Date) As Double

    Dim LHours As Double
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
    With cmd
        .CommandTimeout = 120
        '.CommandText = "Select LHrs from VLHours Where EmpID='" & EmpID & "'"
        .CommandText = "VLHours"
        .CommandType = adCmdStoredProc
         Set .ActiveConnection = con
        .Parameters.Refresh
        
        .Parameters(1).Type = adNumeric
        .Parameters(1).Value = LTime
        
        .Parameters(2).Type = adDBTimeStamp
        .Parameters(2).Value = ToDate
        
        .Parameters(3).Type = adNumeric
        .Parameters(3).Value = NTime
        
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
    Dim cmd As New ADODB.Command
    
    Dim IsExernalSalary As Byte
    
    IsExernalSalary = chkExternal
    
    With cmd
        .CommandTimeout = 120
        '.CommandText = "Select * from VSalarySettings where empid='" & Emp & "'"
        .CommandType = adCmdStoredProc
        .CommandText = "VSalarySettings"
        Set .ActiveConnection = con
        .Parameters.Refresh
        .Parameters(1).Type = adDBTimeStamp
        .Parameters(1).Value = ToDate
        
        
        .Parameters(2).Type = adVarChar
        .Parameters(2).Value = Emp
        
        .Parameters(3).Type = adBoolean
        .Parameters(3).Value = IsExernalSalary
        
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
    If IsNull(rs![totamt]) Then lTotAmt = 0 Else lTotAmt = rs![totamt]
    If IsNull(rs![ClsAmt]) Then lClsAmt1 = 0 Else lClsAmt1 = rs![ClsAmt]
                
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
    Val(rs![ShortTermPer] & "") & "," & LongLoan & ",0," & chkExternal.Value & "," & Val(LV.ListItems(i).SubItems(27)) + Val(LV.ListItems(i).SubItems(21)) & _
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
    
    Dim rs As New ADODB.Recordset, DTLastSunday As Date, DTFirstSunday As Date
    With rs
        .Open "SELECT * FROM MonthlySalaries", con, adOpenStatic, adLockOptimistic
        
        For i = 1 To LV.ListItems.count
        
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
            
            ![FakeWorkingHrs] = GetFakeWorkingHrs(LV.ListItems(i))
            ![FakeRate] = Val(LV.ListItems(i).ListSubItems(6).Tag)
            ![FakeSalary] = Val(LV.ListItems(i).ListSubItems(5).Tag)
            
            !OTDinnerCount = Val(LV.ListItems(i).ListSubItems(7).Tag)
            !OTDinnerAmount = Val(LV.ListItems(i).ListSubItems(8).Tag)
            
            !GPHrs = Val(LV.ListItems(i).ListSubItems(9).Tag)
            !GPHrsAmt = Val(LV.ListItems(i).ListSubItems(10).Tag)
            
            !AmtPaid = Val(LV.ListItems(i).ListSubItems(11).Tag)
            !SundayOTHrs = Val(LV.ListItems(i).ListSubItems(12).Tag)
            !SundayOTRate = Val(LV.ListItems(i).ListSubItems(13).Tag)
            !FixAllowance = Val(LV.ListItems(i).ListSubItems(14).Tag)
            
            !LateComingHrs = Val(LV.ListItems(i).ListSubItems(15).Tag)
            !ShortHrs = Val(LV.ListItems(i).ListSubItems(16).Tag)
            !ShortHrsAmt = Val(LV.ListItems(i).ListSubItems(17).Tag)
                        
            !ZeroAbsentBonus = Val(LV.ListItems(i).ListSubItems(18).Tag)
            
            !DedOnePercent = LV.ListItems(i).ListSubItems(23).Tag
            !DeptID = LV.ListItems(i).ListSubItems(24).Tag
            !LeaveDays = Val(LV.ListItems(i).ListSubItems(25).Tag)
            !presentDays = Val(LV.ListItems(i).ListSubItems(26).Tag)
            
            !HoldSalaryAmt = Val(LV.ListItems(i).ListSubItems(27).Tag)
            
            !OTHrs_Original = Val(LV.ListItems(i).SubItems(28))
            !OTHrs_Net = Val(LV.ListItems(i).ListSubItems(28).Tag)
            !LateHrs_Original = Val(LV.ListItems(i).SubItems(29))
            !LateHrs_Net = Val(LV.ListItems(i).ListSubItems(29).Tag)
            
            .Update
        
            'Update Long Term ...
        
            Dim LongAmt As Double
            LongAmt = Val(LV.ListItems(i).SubItems(20))
        
            con.Execute "DELETE FROM AmtCleared WHERE EmpID='" & rs![EmpId] & "' AND DT='" & Format(DT, "dd-MMM-yyyy") & "'"
            con.Execute "INSERT INTO AmtCleared (EmpID,DT,AmtClrd) VALUES('" & rs![EmpId] & "','" & DT & "'," & LongAmt & ")"
                    
            If Val(LV.ListItems(i).ListSubItems(18).Tag) > 0 Then   'If Zero-Absent, We are paying 1 Day Salary, So Deduct 1 Annual Leave.
                'Get Last Sunday of the Month, & Post Annual Leave...
                DTLastSunday = GetSingleDateValue("MAX(DateFull)", "DateLookup", " WHERE MonthName='" & Format(DT, "MMMM") & "' AND WeekDayName='Sunday' AND FullYear=" & Format(DT, "yyyy"))
                DTFirstSunday = GetSingleDateValue("MIN(DateFull)", "DateLookup", " WHERE MonthName='" & Format(DT, "MMMM") & "' AND WeekDayName='Sunday' AND FullYear=" & Format(DT, "yyyy"))
                con.Execute "INSERT INTO Leaves(EmpID,DT,AuthPerson,Reason,Hrs,CL) VALUES('" & !EmpId & "','" & DTLastSunday & "','Software','Against Bonus',8,1)"
                
                If Val(LV.ListItems(i).ListSubItems(20).Tag) = 0 Then
                    con.Execute "INSERT INTO Leaves(EmpID,DT,AuthPerson,Reason,Hrs,CL) VALUES('" & !EmpId & "','" & DTFirstSunday & "','Software','Auto',4,0.5)"
                End If
                
            End If
            
        Next i
    
        .Close
    
    End With

    Set rs = Nothing
    'Update Short Term...
    con.Execute "UPDATE Advances SET Cleared=1 WHERE MONTH(DeductionStartDT)=" & Month(DT) & " AND YEAR(DeductionStartDT)=" & year(DT) & " AND Type=0 AND Cleared=0"
    
End Sub

Private Function GetFakeWorkingHrs(ITM As ListItem) As Double
    
    Dim myRetVal As Double
    If Val(ITM.SubItems(2)) = Val(ITM.ListSubItems(4).Tag) Then
        myRetVal = Val(ITM.SubItems(8))
    Else
        myRetVal = Val(txtTotalHrs) 'Val(ITM.SubItems(9)) / Val(ITM.ListSubItems(6).Tag) Commented out for Division by Zero Error
    End If
    GetFakeWorkingHrs = myRetVal
    
End Function
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

Private Sub lblTransferNegative_Click()

    On Error GoTo err
    If Not (cmbDept.ID = "0" And cmbEmp.ID = "0") Then
        MsgBox "Please Select All Departments and Employees !!!", vbInformation
        Exit Sub
    End If
    Dim DT As Date
    DT = DateSerial(cmbYear.Text, cmbMonth.ListIndex + 2, 0)
    If DT >= Date Then
        MsgBox "This can only be allowed after month end.", vbInformation
        Exit Sub
    End If
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If Val(.SubItems(27)) < 0 Then  'Salary is in Negative Numbers
                If Val(.SubItems(19)) > 0 Then  'Short Term Loan is Greater than 0
                    If Val(.SubItems(19)) > Abs(Val(.SubItems(27))) Then
                        Call AdjustNegative(.Text, Abs(Val(.SubItems(27))), DT)
                    End If
                End If
            End If
        End With
    Next
    
    MsgBox "Adjusted...", vbInformation
    
    Call cmdRefresh_Click
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub AdjustNegative(strEmpID As String, lAmttoAdjust As Long, DT As Date)

    Dim rs As New ADODB.Recordset
    Dim lAmtLeft As Long
    lAmtLeft = lAmttoAdjust
    Dim lTempAmt As Long
    
    Dim myDT As Date
    myDT = DateSerial(year(DT), Month(DT) + 1, 1)
    
    Call StartTrans(con)
    With rs
        .Open "SELECT EntryID,Amount FROM Advances WHERE EmpID='" & strEmpID & "' AND MONTH(DeductionStartDT)=" & Month(DT) & " AND YEAR(DeductionStartDT)=" & year(DT) & " AND Type=0 ORDER BY Amount DESC", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            If Val(!Amount & "") >= lAmtLeft Then
                lTempAmt = lAmtLeft
                lAmtLeft = 0
            Else
                lTempAmt = Val(!Amount & "")
                lAmtLeft = lAmtLeft - lTempAmt
            End If
            con.Execute "UPDATE Advances SET Amount=" & Val(!Amount & "") - lTempAmt & ",SoftwareDescription='Amount Adjusted for Negative Salary, Orignal Amt=" & _
             Val(!Amount & "") & ", Amt Transferred=" & lTempAmt & " Remain=" & Val(!Amount & "") - lTempAmt & "' WHERE EntryID=" & Val(!EntryID & "")
            If lAmtLeft = 0 Then
                Exit Do
            End If
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    con.Execute "INSERT INTO Advances(EmpID,DT,Type,Description,Amount,DAmount,AmountCleared,AccVoucherNo,DeductionStartDT) VALUES('" & _
     strEmpID & "','" & myDT & "',0,'Negative Salary Adjustment Entry'," & lAmttoAdjust - lAmtLeft & ",0,0,'','" & myDT & "')"
     
    con.CommitTrans
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    If Val(LV.SelectedItem.SubItems(20)) = 0 Then Exit Sub
    If LV.SelectedItem.Bold Then Exit Sub
    
    If chkExternal.Value = vbChecked Then Exit Sub
    
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

Private Sub LV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If LV.ListItems.count = 0 Then
    EditAbsents = False
    Exit Sub
End If
    If x >= LV.Left + LV.ColumnHeaders(5).Left And x < LV.Left + LV.ColumnHeaders(6).Left And _
        y >= LV.SelectedItem.Top And y < LV.SelectedItem.Top + LV.SelectedItem.Height Then
        EditAbsents = True
    Else
        EditAbsents = False
    End If
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button <> 2 Then Exit Sub
    
    If LV.ListItems.count = 0 Then Exit Sub
    
    Dim strRet As String
    
    If bSalaryFinalyzed Then
        'Check If Salary Paid
        If Val(LV.SelectedItem.SubItems(30)) = 0 Then   'Salary Not Paid yet
            strRet = ShowPopUpMenu1("Mark Salary As Paid")
        Else
            strRet = ShowPopUpMenu1("Mark Salary As Un-Paid")
        End If
    Else
        strRet = ShowPopUpMenu1("Change Long Term Deduction", "-", "Stop All Long Term Deductions", "-", "Mark Salary As Paid")
    End If
    
    Dim i As Integer, lColor As Long
    If strRet = "Change Long Term Deduction" Then
        ChangeLTDeductionAmt
    ElseIf strRet = "Stop All Long Term Deductions" Then
        StopAllLTDeductions
    ElseIf strRet = "Mark Salary As Paid" Then
    
        con.Execute "UPDATE MonthlySalaries SET Salary_Paid=1,Salary_Paid_UserName='" & CurrentUserName & "',Salary_Paid_MachineName='" & strComputerName & "',Salary_Paid_DTEntry=getDATE() " & _
         " WHERE EmpID='" & LV.SelectedItem.Text & "' AND MONTH(DT)=" & cmbMonth.List(cmbMonth.ListIndex, 1) & " AND YEAR(DT)=" & cmbYear.Text
         
        
        LV.SelectedItem.ForeColor = vbBlue
        For i = 1 To LV.SelectedItem.ListSubItems.count
            LV.SelectedItem.ListSubItems(i).ForeColor = vbBlue
        Next
        'Call EnterSalaryAlreadyPaidAmount
    ElseIf strRet = "Mark Salary As Un-Paid" Then
    
        con.Execute "UPDATE MonthlySalaries SET Salary_Paid=0,Salary_Paid_UserName='" & CurrentUserName & "',Salary_Paid_MachineName='" & strComputerName & "',Salary_Paid_DTEntry=getDATE() " & _
         " WHERE EmpID='" & LV.SelectedItem.Text & "' AND MONTH(DT)=" & cmbMonth.List(cmbMonth.ListIndex, 1) & " AND YEAR(DT)=" & cmbYear.Text
        
        LV.SelectedItem.ForeColor = LV.ForeColor
        For i = 1 To LV.SelectedItem.ListSubItems.count
            LV.SelectedItem.ListSubItems(i).ForeColor = LV.ForeColor
        Next
    End If
    
End Sub

Private Sub ChangeLTDeductionAmt()

    If LV.ListItems.count = 0 Then Exit Sub
    If Val(LV.SelectedItem.ListSubItems(20).Tag) = 0 Then Exit Sub
    If LV.SelectedItem.Bold Then Exit Sub
    
    'If chkExternal.Value = vbChecked Then Exit Sub
    
    Dim Amt As String
    Amt = InputBox("Please Enter Amount To Deduct For Long Term." & vbNewLine & "The Current Deduction Amount Is " & LV.SelectedItem.SubItems(20) & ".", , LV.SelectedItem.SubItems(20)) & ""
        
    If Not IsNumeric(Amt) Then
        MsgBox "Invalid Amount.", vbInformation
    Else
        ' Update The Balance
        Dim dAmt As Double
        dAmt = Val(Amt)
        If dAmt > Val(LV.SelectedItem.ListSubItems(20).Tag) Or dAmt < 0 Then
            MsgBox "Invalid Amount.", vbInformation
            Exit Sub
        End If
        Dim iAffect As Integer, myDT As Date
        myDT = DateSerial(cmbYear.Text, cmbMonth.ListIndex + 2, 0)  'Get the Last date of the month.
                
        Call StartTrans(con)
        con.Execute "UPDATE LongTermDeductionsMonthly SET AmtToDeduct=" & dAmt & ",DTEntry=getDate(),UserName='" & _
         CurrentUserName & "',MachineName='" & strComputerName & "' WHERE EmpID='" & LV.SelectedItem.Text & _
          "' AND DedMonth='" & myDT & "'", iAffect
        
        If iAffect = 0 Then
            con.Execute "INSERT INTO LongTermDeductionsMonthly(DedMonth,EmpID,AmtToDeduct,UserName,MachineName)" & _
             " VALUES('" & myDT & "','" & LV.SelectedItem.Text & "'," & dAmt & ",'" & CurrentUserName & _
             "','" & strComputerName & "')"
        End If
        
        con.CommitTrans
        
        LV.SelectedItem.SubItems(27) = Val(LV.SelectedItem.SubItems(27)) + Val(LV.SelectedItem.SubItems(20)) - Abs(Val(Amt))
        LV.SelectedItem.SubItems(20) = Abs(Val(Amt))
        
    End If
    
End Sub

Private Sub EnterSalaryAlreadyPaidAmount()

    If LV.ListItems.count = 0 Then Exit Sub
    If Val(LV.SelectedItem.ListSubItems(27)) <= 0 Then Exit Sub
    If LV.SelectedItem.Bold Then Exit Sub
    
    'If chkExternal.Value = vbChecked Then Exit Sub
    
    Dim Amt As String
    Amt = InputBox("Please Enter Amount." & vbNewLine & "The Current Amount Is " & Val(LV.SelectedItem.ListSubItems(11).Tag) & ".", , LV.SelectedItem.SubItems(22)) & ""
        
    If Not IsNumeric(Amt) Then
        MsgBox "Invalid Amount.", vbInformation
    Else
        ' Update The Balance
        Dim dAmt As Double
        dAmt = Val(Amt)
        If dAmt > Val(LV.SelectedItem.ListSubItems(27)) Or dAmt < 0 Then
            MsgBox "Invalid Amount.", vbInformation
            Exit Sub
        End If
        Dim iAffect As Integer, myDT As Date
        myDT = DateSerial(cmbYear.Text, cmbMonth.ListIndex + 2, 0)  'Get the Last date of the month.
                
        Call StartTrans(con)
        con.Execute "UPDATE SalaryAlreadyPaid SET AmtPaid=" & dAmt & ",DTEntry=getDate(),UserName='" & _
         CurrentUserName & "',MachineName='" & strComputerName & "' WHERE EmpID='" & LV.SelectedItem.Text & _
          "' AND ForMonth='" & myDT & "'", iAffect
        
        If iAffect = 0 Then
            con.Execute "INSERT INTO SalaryAlreadyPaid(ForMonth,EmpID,AmtPaid,UserName,MachineName)" & _
             " VALUES('" & myDT & "','" & LV.SelectedItem.Text & "'," & dAmt & ",'" & CurrentUserName & _
             "','" & strComputerName & "')"
        End If
        
        con.CommitTrans
        
        LV.SelectedItem.ListSubItems(11).Tag = Abs(Val(Amt))
        
    End If
    
End Sub

Private Sub StopAllLTDeductions()

    If LV.ListItems.count = 0 Then Exit Sub
    
    
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        Dim dAmt As Double
        dAmt = 0 'Val(LV.ListItems(i).SubItems(20))
        
        Dim iAffect As Integer, myDT As Date
        myDT = DateSerial(cmbYear.Text, cmbMonth.ListIndex + 2, 0)  'Get the Last date of the month.
        
        Call StartTrans(con)
        
        con.Execute "UPDATE LongTermDeductionsMonthly SET AmtToDeduct=" & dAmt & ",DTEntry=getDate(),UserName='" & _
         CurrentUserName & "',MachineName='" & strComputerName & "' WHERE EmpID='" & LV.ListItems(i).Text & _
          "' AND DedMonth='" & myDT & "'", iAffect
        
        If iAffect = 0 Then
            con.Execute "INSERT INTO LongTermDeductionsMonthly(DedMonth,EmpID,AmtToDeduct,UserName,MachineName)" & _
             " VALUES('" & myDT & "','" & LV.ListItems(i).Text & "'," & dAmt & ",'" & CurrentUserName & _
             "','" & strComputerName & "')"
        End If
        
        con.CommitTrans
        LV.ListItems(i).SubItems(27) = Val(LV.ListItems(i).SubItems(27)) + Val(LV.ListItems(i).SubItems(20)) - dAmt
        LV.ListItems(i).SubItems(20) = dAmt
        
    Next
        
End Sub

Private Sub mnuLateEmpList_Click()
    Call PrintLateEmpList
End Sub


Private Sub mnuPayslipFull_Click()
    Call PrintReport(True, , , , , 3)
End Sub

Private Sub mnuPayslipMini_Click()
    Call PrintReport(True, , , , , 2)
End Sub

Private Sub mnuPayslipRule9_Click()
    Call PrintReport(True, , , , , 1)
End Sub

Private Sub mnuPaySlips_Click()
   Call PrintReport(True)
End Sub

Private Sub mnuPayslipsInternal_Click()
    Call PrintInternalReport(True)
End Sub

Private Sub mnuPrintNegativeEmpList_Click()
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\EmpListNegativeSalaries.rpt")
    rpt.FormulaFields.GetItemByName("Period").Text = "'" & cmbMonth.Text & " - " & cmbYear.Text & "'"
    Dim myDT As Date
    myDT = DateSerial(cmbYear, cmbMonth.ListIndex + 2, 1)
    Dim f As New frmPrevRpt
    f.ShowReport "{Advances.Description}='Negative Salary Adjustment Entry' AND {Advances.DT}=#" & myDT & "#", rpt

End Sub

Private Sub mnuPrintOTPaySheet_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    Call InsertRecords

    Dim unloadflag As Boolean

    Screen.MousePointer = vbHourglass

    Dim f As New frmPrevRpt, rpt  As CRAXDDRT.Report

    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptOverTimePaymentSheet.rpt")
    
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & strCompany & "'"
        ElseIf FormulaField.Name = "{@FromTo}" Then
            FormulaField.Text = "' For " & cmbMonth.Tag & "-" & cmbYear.Tag & "'"
        ElseIf FormulaField.Name = "{@MonthDays}" Then
            FormulaField.Text = Val(txtTotalHrs)
        End If
    Next
    Cond = "{@OHAmt}>0"
'    On Error Resume Next
'    SrcReport.Database.Tables(1).ConnectBufferString = "DSN=DSN_Vision;;User ID=sa;;Database=Vision;;Password=visionasim"
'    rpt.SQLQueryString = SQL
    f.ShowReport Cond & " ", rpt

    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuPrintSalPacket_Click()
   Call PrintReport(True, , , True)
End Sub

Private Sub mnuRePostSalaryVouchers_Click()
    
    If cmdUpdate.Visible Then Exit Sub
    If CurrentUserName <> "Administrator" Then Exit Sub
    If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Call AddSalaryVouchers
    MsgBox "Done", vbInformation
    
End Sub

Private Sub mnuSalarysheetInternal_Click()
 Call PrintInternalReport(False)
End Sub

Private Sub mnuSalSheet_Click()
   Call PrintReport(False)
End Sub

Private Sub mnuSalSheetNegative_Click()

    Call PrintReport(False, , , , True)
    
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
   
    For i = 1 To LV.ListItems.count
    
        con.Execute "Update MonthlySalaries Set Transfered=" & CBool(LV.ListItems(i).Checked Or LV.ListItems(i).Bold) & " Where EmpID+Format(DT,'dd-MMM-yyyy')='" & LV.ListItems(i).key & "'"
    
    Next i
      
   
   AddSalaryEntries = True
   Exit Function
err:
   AddSalaryEntries = False
   MsgBox "Cannot Add Salary Vouchers Due To Error in " & err.Source & " :" & vbNewLine & err.Description, vbCritical
   ConAcc.RollbackTrans
End Function

Private Sub GetTotal(Stotal As Double, dTotal As Double, PTotal As Double, TotalType As String)

   Dim StaffTotal As Double, ProductionTotal As Double, DirectorTotal As Double, Amount As Double



   For i = 1 To LV.ListItems.count
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
   dTotal = DirectorTotal
   
End Sub

Private Function GetTotalTax() As Double
    Dim total As Double
    For i = 1 To LV.ListItems.count
        total = total + Val(LV.ListItems(i).SubItems(14))
    Next
    GetTotalTax = total
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


    For i = 1 To LV.ListItems.count
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
'            .Formulas(9) = "Company='" & strCompany & "'"
'
'        .WindowState = crptMaximized
'        .WindowTitle = "Print Rupee Structure"
'        .Action = 1
'
'    End With

Screen.MousePointer = vbHourglass

Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
Set rpt = rpt.OpenSubreport(rptPayrollPath & "\rptRupeeStructure.rpt")
Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

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

    con.Execute "DELETE FROM PrintSalary"

    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM PrintSalary", con, adOpenStatic, adLockOptimistic
        For i = 1 To LV.ListItems.count
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
            
            
            !OTDinnerCount = Val(LV.ListItems(i).ListSubItems(7).Tag)
            !OTDinnerAmount = Val(LV.ListItems(i).ListSubItems(8).Tag)
            
            !GPHrs = Val(LV.ListItems(i).ListSubItems(9).Tag)
            !GPHrsAmt = Val(LV.ListItems(i).ListSubItems(10).Tag)
            
            !AmtPaid = Val(LV.ListItems(i).ListSubItems(11).Tag)
            !SundayOTHrs = Val(LV.ListItems(i).ListSubItems(12).Tag)
            !SundayOTRate = Val(LV.ListItems(i).ListSubItems(13).Tag)
            !FixAllowance = Val(LV.ListItems(i).ListSubItems(14).Tag)
            !LateComingHrs = Val(LV.ListItems(i).ListSubItems(15).Tag)
            !ShortHrs = Val(LV.ListItems(i).ListSubItems(16).Tag)
            !ShortHrsAmt = Val(LV.ListItems(i).ListSubItems(17).Tag)
                        
            !ZeroAbsentBonus = Val(LV.ListItems(i).ListSubItems(18).Tag)
            
            !DedOnePercent = LV.ListItems(i).ListSubItems(23).Tag
            !DeptID = LV.ListItems(i).ListSubItems(24).Tag
            !LeaveDays = Val(LV.ListItems(i).ListSubItems(25).Tag)
            !presentDays = Val(LV.ListItems(i).ListSubItems(26).Tag)
            
            !HoldSalaryAmt = Val(LV.ListItems(i).ListSubItems(27).Tag)
            
            !OTHrs_Original = Val(LV.ListItems(i).SubItems(28))
            !OTHrs_Net = Val(LV.ListItems(i).ListSubItems(28).Tag)
            !LateHrs_Original = Val(LV.ListItems(i).SubItems(29))
            !LateHrs_Net = Val(LV.ListItems(i).ListSubItems(29).Tag)
            
            .Update
        Next i
    
        .Close
    
    End With

    Set rs = Nothing

    SQL = "INSERT INTO EmpLeaveBalance SELECT  EmpID,SUM(CASE WHEN CL<0.5 THEN 0.5 ELSE CL END) AS CasualLeaves,Count(CASE WHEN SL<0.5 THEN 0.5 ELSE SL END) as SickLeaves," & _
     "count(CASE WHEN AL<0.5 THEN 0.5 ELSE AL END) AS AnnualLeaves" & _
     ",count(CASE WHEN ML<0.5 THEN 0.5 ELSE ML END) AS MaternityLeaves FROM VEmpLeaves WHERE DT <'" & DateSerial(cmbYear, cmbMonth.ListIndex + 1, 1) & "' Group By empID"
    
    con.Execute "Delete From EmpLeaveBalance"
    con.Execute SQL, a

End Sub

Private Function FinalizedSalarySheetLoaded() As Boolean

    Dim rs As New ADODB.Recordset, ITM As ListItem
    Dim Cond As String
    Dim IsExernalSalary As Byte
    Dim GrandTotal As Double
    
    Dim dSalary As Double
    Dim dRate As Double
    Dim dWorkingHrs As Double
    
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
    " and Year(DT)=" & Val(cmbYear) & " and IsForSA=" & chkExternal.Value
            
    Dim lForeColor As Long, i As Integer
    With rs
        .Open "SELECT * FROM VMonthlySalaries " & Cond, con, adOpenForwardOnly, adLockReadOnly
        
        If .RecordCount = 0 Then
            FinalizedSalarySheetLoaded = False
            GoTo END_OF_FUNCTION
        End If
            
            LV.ListItems.Clear
            Do Until .EOF
        
                Set ITM = LV.ListItems.add(, ![EmpId], ![EmpId] & "")
                
'                If IsComputed Then
'                    ITM.Checked = IsPayed
'                    ITM.Bold = IsTransfered
'                End If
                
                If Val(txtTotalHrs) <> Val(![TotalMonthHrs] & "") Then
                    txtTotalHrs = ![TotalMonthHrs] & ""
                    MsgBox "Salary Sheet Is Finalized According To " & txtTotalHrs & " Working Hours.", vbInformation
                End If
                
                If m_bShowFake Then
                    dSalary = Val(!FakeSalary & "")
                    dRate = Val(!FakeRate & "")
                    dWorkingHrs = Val(!FakeWorkingHrs & "")
                Else
                    dSalary = Val(!BSal & "")
                    dRate = Val(!Rate & "")
                    dWorkingHrs = Val(!SDays & "")
                End If
            
                ITM.Tag = Val(![Emptype] & "")
                
                ITM.ListSubItems.add(, , ![Name] & "").Tag = Val(![PrevLTLoan] & "")
                ITM.ListSubItems.add , , dSalary
                
                ITM.ListSubItems(2).Tag = ![CasualLeaves] & "," & ![SickLeaves] & "," & ![CompensatoryLeaves] & _
                "," & ![AnnualLeaves] & "," & ![WPLeaves] & "," & ![MaternityLeaves]
                
                ITM.ListSubItems.add(, , dRate).Tag = ![HrsPerDay]
                ITM.ListSubItems.add(, , ![ADays] & "").Tag = dWorkingHrs 'Val(![SDays] & "")
                ITM.ListSubItems.add , , ![AAmt] & ""
                ITM.ListSubItems.add , , ![AAllow] & ""
                ITM.ListSubItems.add , , ![AAllowAmt] & ""
                ITM.ListSubItems.add , , IIf(chkExternal = vbChecked, getHrsMin(dWorkingHrs), dWorkingHrs)
                ITM.ListSubItems.add , , ![SAmt] & ""
                ITM.ListSubItems.add , , ![Leaves] & ""
                ITM.ListSubItems.add , , ![LeaveAmt] & ""
                ITM.ListSubItems.add , , ![OHrs] & ""
                ITM.ListSubItems.add , , ![OAmt] & ""
                ITM.ListSubItems.add , , ![lHrs] & ""
                ITM.ListSubItems.add , , ![lAmt] & ""
                ITM.ListSubItems.add , , ![total] & ""
                ITM.ListSubItems.add , , ![Tax] & ""
                ITM.ListSubItems.add , , ![NetTtl] & ""
                ITM.ListSubItems.add , , ![ShortTerm] & ""
                ITM.ListSubItems.add , , ![LongTerm] & ""
                ITM.ListSubItems.add , , ![AdvSal] & ""
                ITM.ListSubItems.add , , ![UnionFund] & ""
                ITM.ListSubItems.add , , ![Fine] & ""
                ITM.ListSubItems.add , , ![Bonus] & ""
                ITM.ListSubItems.add , , ![Lunch] & ""
                ITM.ListSubItems.add , , ![EOBI] & ""
                ITM.ListSubItems.add , , ![Balance] & ""
               
                ITM.ListSubItems(7).Tag = Val(!OTDinnerCount & "")
                ITM.ListSubItems(8).Tag = Val(!OTDinnerAmount & "")
                
                ITM.ListSubItems(9).Tag = Val(!GPHrs & "")
                ITM.ListSubItems(10).Tag = Val(!GPHrsAmt & "")
                
                ITM.ListSubItems(11).Tag = Val(!AmtPaid & "")
                ITM.ListSubItems(12).Tag = Val(!SundayOTHrs & "")
                ITM.ListSubItems(13).Tag = Val(!SundayOTRate & "")
                ITM.ListSubItems(14).Tag = Val(!FixAllowance & "")
                
                ITM.ListSubItems(15).Tag = Val(!LateComingHrs & "")
                ITM.ListSubItems(16).Tag = Val(!ShortHrs & "")
                ITM.ListSubItems(17).Tag = Val(!ShortHrsAmt & "")
                                
                ITM.ListSubItems(18).Tag = Val(!ZeroAbsentBonus & "")
                
                ITM.ListSubItems(23).Tag = Val(!DedOnePercent & "")
                ITM.ListSubItems(24).Tag = !DeptID & ""
                ITM.ListSubItems(25).Tag = Val(!LeaveDays & "")
                ITM.ListSubItems(26).Tag = Val(!presentDays & "")
            
                ITM.ListSubItems(27).Tag = Val(!HoldSalaryAmt & "")
                        
                ITM.SubItems(28) = Val(!OTHrs_Original & "")
                ITM.ListSubItems(28).Tag = Val(!OTHrs_Net & "")
                
                ITM.SubItems(29) = Val(!LateHrs_Original & "")
                ITM.ListSubItems(29).Tag = Val(!LateHrs_Net & "")
                
                ITM.SubItems(30) = Abs(!Salary_Paid)
                GrandTotal = GrandTotal + Val(ITM.SubItems(27))
                            
                'If Val(!Salary_Already_Paid_Amt & "") > 0 Then
                If !Salary_Paid Then
                    lForeColor = vbBlue
                Else
                    lForeColor = LV.ForeColor
                End If
                
                ITM.ForeColor = lForeColor
                For i = 1 To ITM.ListSubItems.count
                    ITM.ListSubItems(i).ForeColor = lForeColor
                Next
                
                .MoveNEXT
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

    Dim f As New frmPrevRpt, rpt  As CRAXDDRT.Report


    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptLateEmpList.rpt")

    Cond = "{PrintSalary.LAmt}>0 "
    
    'rpt.RecordSelectionFormula = "{PrintSalary.LAmt}>0"
    f.ShowReport Cond & " ", rpt

    Screen.MousePointer = vbDefault

End Sub
