VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCashBookReport 
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleMode       =   0  'User
   ScaleWidth      =   10641.08
   WindowState     =   2  'Maximized
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   705
      Left            =   45
      TabIndex        =   17
      Top             =   0
      Width           =   11850
      Begin VB.CheckBox Chks1 
         Caption         =   "Cash Trans."
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
         Height          =   195
         Index           =   0
         Left            =   1755
         TabIndex        =   23
         Top             =   390
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.ComboBox cmbRange 
         Height          =   315
         ItemData        =   "frmCashBookReport.frx":0000
         Left            =   75
         List            =   "frmCashBookReport.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   315
         Width           =   2685
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   4485
         TabIndex        =   19
         Top             =   315
         Width           =   1560
         _ExtentX        =   2752
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
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   122093571
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   2775
         TabIndex        =   20
         Top             =   315
         Width           =   1710
         _ExtentX        =   3016
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
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   122093571
         CurrentDate     =   37055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   " Range                         From                            To"
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
         Height          =   255
         Index           =   0
         Left            =   15
         TabIndex        =   22
         Top             =   120
         Width           =   11130
      End
      Begin MSForms.CommandButton cmdRefresh 
         Default         =   -1  'True
         Height          =   360
         Left            =   6075
         TabIndex        =   21
         Top             =   270
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
   Begin VB.Frame FraSummary 
      Caption         =   "Summary :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   900
      Left            =   60
      TabIndex        =   7
      Top             =   6975
      Visible         =   0   'False
      Width           =   5580
      Begin VB.TextBox txtNetBalance 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   3945
         TabIndex        =   11
         Top             =   435
         Width           =   1530
      End
      Begin VB.TextBox txtLastBalance 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   105
         TabIndex        =   10
         Top             =   435
         Width           =   1530
      End
      Begin VB.TextBox txtUnAuthCredit 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   2805
         TabIndex        =   9
         Top             =   435
         Width           =   1125
      End
      Begin VB.TextBox txtUnAuthDebit 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1650
         TabIndex        =   8
         Top             =   435
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Net Balance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   3945
         TabIndex        =   15
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Balance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   135
         TabIndex        =   14
         Top             =   210
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Un-Auth Credit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   2805
         TabIndex        =   13
         Top             =   210
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Un-Auth Debit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   1650
         TabIndex        =   12
         Top             =   210
         Width           =   1020
      End
   End
   Begin VB.OptionButton Opts 
      Caption         =   "Cash Book"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   16
      Top             =   7410
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.OptionButton Opts 
      Caption         =   "Day Book"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   3210
      TabIndex        =   6
      Top             =   7605
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2490
      TabIndex        =   4
      Top             =   1635
      Visible         =   0   'False
      Width           =   3615
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2025
      Top             =   2100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashBookReport.frx":0055
            Key             =   "PND"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashBookReport.frx":04AD
            Key             =   "CLR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashBookReport.frx":0901
            Key             =   "DUE"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashBookReport.frx":0D59
            Key             =   "BNC"
         EndProperty
      EndProperty
   End
   Begin VB.Frame FDrag 
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   60
      MousePointer    =   7  'Size N S
      TabIndex        =   2
      Top             =   4155
      Width           =   11805
   End
   Begin MSComctlLib.ListView LV 
      Height          =   3375
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   8388608
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Dated"
         Object.Width           =   1969
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Voucher #"
         Object.Width           =   2515
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Account"
         Object.Width           =   3742
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Description"
         Object.Width           =   5317
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Debit"
         Object.Width           =   2166
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Credit"
         Object.Width           =   2166
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "User"
         Object.Width           =   1969
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Comp."
         Object.Width           =   2363
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Entry Date"
         Object.Width           =   2836
      EndProperty
   End
   Begin MSComctlLib.ListView LVCheque 
      Height          =   2610
      Left            =   60
      TabIndex        =   1
      Top             =   4320
      Visible         =   0   'False
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   4604
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   15330543
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Dated"
         Object.Width           =   2363
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cheque #"
         Object.Width           =   2068
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bank & Branch"
         Object.Width           =   3939
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Description"
         Object.Width           =   4529
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Amount"
         Object.Width           =   2363
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Type"
         Object.Width           =   1969
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Chq. Date"
         Object.Width           =   2068
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Status"
         Object.Width           =   1575
      EndProperty
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   375
      Left            =   9105
      TabIndex        =   5
      Top             =   7170
      Width           =   1335
      Caption         =   "Print        "
      PicturePosition =   327683
      Size            =   "2355;661"
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
      Left            =   10455
      TabIndex        =   3
      Top             =   7170
      Width           =   1335
      ForeColor       =   64
      Caption         =   "Close     "
      PicturePosition =   327683
      Size            =   "2355;661"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Menu mnupop 
      Caption         =   "mnuPop"
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete This Voucher"
      End
      Begin VB.Menu mnuEditDesc 
         Caption         =   "Edit Description"
      End
      Begin VB.Menu dash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditVoucher 
         Caption         =   "Edit Voucher"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditVoucherDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintVoucher 
         Caption         =   "Print This Voucher"
      End
      Begin VB.Menu mnuPreviewVoucher 
         Caption         =   "Preview This Voucher"
      End
      Begin VB.Menu mnuprintOnblank 
         Caption         =   "Print This Voucher (Blank Paper)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintCheque 
         Caption         =   "Print Cheque"
      End
      Begin VB.Menu mnuPrintMakerBillOld 
         Caption         =   "Print Maker Bill (Old)"
      End
      Begin VB.Menu mnuPrintMakerBill 
         Caption         =   "Print Maker Bill"
      End
      Begin VB.Menu mnuPrintVenderBill 
         Caption         =   "Print Vender Bill"
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   ""
      Begin VB.Menu mnuDailyTrans 
         Caption         =   "Print Daily Transaction Report"
      End
      Begin VB.Menu mnuCatTrans 
         Caption         =   "Print Catagorised Transaction Report"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintVoucher1 
         Caption         =   "Print This Voucher"
      End
      Begin VB.Menu mnuprintOnblank1 
         Caption         =   "Print This Voucher (Blank Paper)"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmCashBookReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReadyToDrag As Boolean
Dim BankAccNo As String
Public Loaded As Boolean
Dim PermissionDenied As Boolean

Private Sub Chks_Click(Index As Integer)
    Call cmdRefresh_Click
End Sub

Private Sub Chks1_Click(Index As Integer)
    Call cmdRefresh_Click
End Sub

Private Sub cmbRange_Click()

    DtTo = Date
    DtFrom.Enabled = False
    DtTo.Enabled = False
    Select Case cmbRange.ListIndex
    Case 0
        DtFrom = Date
    Case 1
        DtFrom = DateAdd("d", -7, Date)
    Case 2
        DtFrom = DateAdd("d", -15, Date)
    Case 3
        DtFrom = DateAdd("d", -30, Date)
    Case 4
        DtFrom.Enabled = True
        DtTo.Enabled = True
        'DtFrom.SetFocus
        Exit Sub
    End Select
    Call cmdRefresh_Click


End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    'PopUpMenu mnuPrint, 8, cmdPrint.Left, cmdPrint.Top + cmdPrint.Height
'    If strCompany = "Instrumed" Then
'        Call PrintCashBookReport(DtFrom.Value, DtTo.Value)
'    Else
'        Call PrintCashnBankTransaction
'    End If
    Call PrintCashBookReport(DtFrom.Value, DtTo.Value)
    
End Sub



Public Sub cmdRefresh_Click()

    If Not Loaded Then Exit Sub
    'If Opts(0) Then Call RefreshLedger Else Call RefreshCashLedger
    Call RefreshCashLedger
    'cmbAccNo.Tag = cmbAccNo
    'CmbAccTitle.Tag = CmbAccTitle
    Call Form_Resize
    
End Sub

Private Sub DtFrom_Change()
    Call setDateRange(DtFrom)
End Sub

Private Sub DtTo_Change()
    Call setDateRange(DtTo)
End Sub

Private Sub FDrag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReadyToDrag = True
    FDrag.BackColor = &H808080
    FDrag.ZOrder
End Sub

Private Sub FDrag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim TheTop As Long
TheTop = FDrag.Top + Y

    If ReadyToDrag Then
        If (TheTop > (FAcc.Top + FAcc.Height + 500)) And (TheTop < cmdClose.Top - 500) Then
            FDrag.Top = TheTop
        End If
    End If

    
End Sub

Private Sub FDrag_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If ReadyToDrag Then
    
        FDrag.BackColor = Me.BackColor
        LV.Height = FDrag.Top - LV.Top
        LVCheque.Move 50, FDrag.Top + FDrag.Height, LVCheque.Width, Me.ScaleHeight - (FAcc.Height + cmdClose.Height + LV.Height + FDrag.Height + 100)
        If LVCheque.Top > cmdClose.Top - 500 Then
            LVCheque.Top = cmdClose.Top - 500
        End If
        
        ReadyToDrag = False
        
        FDrag.ZOrder 1
    End If
    
End Sub

Private Sub Form_Activate()

'  Check For User Access
    mnupop.Visible = False
    mnuPrint.Visible = False
   If PermissionDenied Then
      Me.Hide
      MsgBox "You Donot Have Permissions To Use This Option. Please Contact Your Administrator", vbInformation
      Unload Me
      Exit Sub
   End If

   If Not Loaded Then
      Loaded = True
      Call cmdRefresh_Click
   End If

End Sub

Private Sub Form_GotFocus()
    Call Form_Activate
End Sub

Private Sub Form_Load()

    'PermissionDenied = Not UserHasAccess("TransactionReport")
    Loaded = False
    cmbRange.ListIndex = 0
        
    If strCompany = "Sunlike" Then
        FraSummary.Visible = True
    Else
        FraSummary.Visible = False
    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Dim TheHeight As Long
    LVCheque.Visible = LVCheque.ListItems.count > 0

    FAcc.Move (Me.ScaleWidth - FAcc.Width) / 2
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 100
    cmdPrint.Move cmdClose.Left - cmdPrint.Width - 50, cmdClose.Top
    
    FraSummary.Top = cmdPrint.Top - 900 + cmdPrint.Height
    
    FDrag.Move 50, Me.ScaleHeight * 3 / 5, Me.ScaleWidth - 100
 
    If LVCheque.Visible Then
        TheHeight = FDrag.Top
        FDrag.Visible = True
    Else
        'TheHeight = Me.ScaleHeight - cmdClose.Height - 200
        TheHeight = Me.ScaleHeight - FraSummary.Height - 200
        FDrag.Visible = False
    End If
 
    LV.Move 50, FAcc.Top + FAcc.Height, Me.ScaleWidth - 100, TheHeight - (FAcc.Top + FAcc.Height)
    LVCheque.Move 50, FDrag.Top + FDrag.Height, Me.ScaleWidth - 100, Me.ScaleHeight - (FAcc.Height + cmdClose.Height + LV.Height + 230)
    FDrag.ZOrder 1
    
End Sub


Private Sub RefreshLedger()

    Screen.MousePointer = vbHourglass
    Dim PrevDate As Date
    Dim PrevVoucher As String
    Dim rsLedger As New ADODB.Recordset
    
    
    Dim trans(0 To 4) As String

    trans(0) = "Journal Transactions"
    trans(1) = "Bank Payment Transactions"
    trans(2) = "Cash Payment Transactions"
    trans(3) = "Cash Receipt Transactions"
    trans(4) = "Bank Receipt Transactions"
    
    LV.ListItems.Clear
    LVCheque.ListItems.Clear
        
    With rsLedger
        .CursorLocation = adUseClient
       
    Dim ITM As ListItem, Balance As String
 
    For j = 0 To Chks.count - 1

    If cmbRange.ListIndex = 0 Then
        SQL = "SELECT * FROM VLedger1 where VDate = '" & DtFrom & "' AND VchrNo Like '%" & Chks(j).Tag & "%' Order By VDate,VchrNo,Debit DESC"
    ElseIf j = 4 Then
        SQL = "SELECT * FROM VLedger1 where (VDate between '" & DtFrom & "' AND '" & DtTo & "') AND (VchrNo Like '%" & Chks(j).Tag & "%' OR LEFT(VchrNo,2)='RV') Order By VDate,VchrNo,Debit DESC "
    Else
        SQL = "SELECT * FROM VLedger1 where (VDate between '" & DtFrom & "' AND '" & DtTo & "') AND (VchrNo Like '%" & Chks(j).Tag & "%') Order By VDate,VchrNo,Debit DESC"
    End If
   
    If .State = 1 Then .Close
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        ' Refresh Cheque List
  
        If j > 0 And j < 3 And Chks(j).Value = vbChecked Then Call RefreshChequeLedger(CBool(j = 2))
 
        If Chks(j).Value <> vbChecked Or .RecordCount = 0 Then GoTo EndOfLoop
     
        LV.ListItems.add
        With LV.ListItems.add().ListSubItems
            .add
            .add
            .add(3, , trans(j)).Bold = True
        End With
        LV.ListItems.add
 
        For i = 0 To .RecordCount - 1
            If ![VchrNo] <> PrevVoucher Then
                PrevVoucher = ![VchrNo]
                Set ITM = LV.ListItems.add
                ITM.SubItems(3) = "-------------------------------------------"
            End If
            Set ITM = LV.ListItems.add(, ![SNo] & "'", Format(![Vdate], "dd-MMM-yyyy"))
        
            If ![Vdate] <> PrevDate Then
                PrevDate = ![Vdate]
            Else
                ITM.ForeColor = RGB(125, 125, 225)
            End If
        
            ITM.SubItems(1) = ![VchrNo] & ""
            ITM.SubItems(2) = ![AccTitle] & ""
        
            Dim Dpst As String
        
            If Val(![DpstSlip] & "") = 0 Then Dpst = "" Else Dpst = " Dpst Slip:" & ![DpstSlip]
        
                ITM.SubItems(3) = ![Description] & Dpst
                ITM.SubItems(4) = Format(Val(![Debit] & ""), "###,##0.00")
                ITM.SubItems(5) = Format(Val(![Credit] & ""), "###,##0.00")
                ITM.Tag = Val(![CSNo] & "")
                ITM.SubItems(6) = ![UserName] & ""
                ITM.SubItems(7) = ![MachineName] & ""
                ITM.SubItems(8) = Format(!VchrEntryDate & "", "dd-MMM-yyyy")
                .MoveNEXT
         
            Next
      
EndOfLoop:

        Next


    End With
     
    '"(VDate" & Format(DtFrom.value, "(yyyy,MM,dd)") & " to Date" & Format(DtTo.value, "(yyyy,MM,dd)") & ") and ({ledgerrpt.acno}='" & Trim(Text16.Text) & "')"
    'Now Fill The Other List With Cheques
    'Call RefreshChequeLedger(True)
    Screen.MousePointer = vbDefault
    PrevVoucher = ""
End Sub

Private Sub RefreshChequeLedger(Payment As Boolean)
Screen.MousePointer = vbHourglass
If notpayment Then LVCheque.ListItems.Clear
    Set LVCheque.SmallIcons = ImageList1
    Dim ListColor As Long
    Dim CrsLedger As New ADODB.Recordset
    
    With CrsLedger
        .CursorLocation = adUseClient
        If cmbRange.ListIndex = 0 Then
            .Open "Select * from VChqLedger Where CDate ='" & DtFrom & "' AND  Payment=" & Abs(Payment) & _
               " AND Amount<>0 Order by CDate,Sno", con, 0, 1
        Else
            .Open "Select * from VChqLedger Where (CDate between '" & DtFrom & "' and '" & DtTo & "') AND Amount<>0 AND  Payment=" & Abs(Payment) & _
               " Order by CDate,Sno", con, 0, 1
        End If
            
        
        If .RecordCount > 0 Then
            
         
         
         If Payment Then TheStr = "Payment Cheques" Else TheStr = "Cheques Received"
            With LVCheque.ListItems.add().ListSubItems
               .add
               .add
               .add(, , TheStr).Bold = True
            End With
         
         
         
         
            Dim ITM As ListItem
            Do Until .EOF
                If ![Bounced] Then
                  ListColor = vbRed
                ElseIf (![ChequeDate] <= Date And Not ![Posted]) Then
                  ListColor = vbBlue
                ElseIf ![Posted] Then
                  ListColor = RGB(0, 125, 0)
                Else
                  ListColor = vbBlack
                End If
                
                Set ITM = LVCheque.ListItems.add(, ![SNo] & "'", Format(![CDate], "dd-MMM-yyyy"))
                ITM.ForeColor = ListColor
                
                
                
                For i = 1 To 7
                    With ITM.ListSubItems.add(, , "")
                     .ForeColor = ListColor
                    End With
                    
                Next i
                
                ITM.ListSubItems(1) = ![ChequeNo] & ""
                ITM.ListSubItems(2) = ![Bank] & Space$(3) & ![Branch]
                ITM.ListSubItems(3) = ![Description]
                ITM.ListSubItems(4) = Format(Val(![Amount] & ""), "###,##0.00")
                ITM.ListSubItems(5) = ![ChequeType] & ""
                ITM.ListSubItems(6) = Format(![ChequeDate], "dd-MMM-yyyy")
                
                'Check For The Chq Status
                Dim Status As String, img As String
                If ![Posted] Then
                  Status = "Cleared"
                  img = "CLR"
                ElseIf ![Bounced] Then
                  Status = "Bounced"
                  img = "BNC"
                ElseIf ITM.ForeColor = vbBlue Then
                  Status = "Due"
                  img = "DUE"
                Else
                  Status = "Pending"
                  img = "PND"
                End If
                 
                 ITM.ListSubItems(7) = Status
                 ITM.SmallIcon = img
                 
                'Set Flag That it Was Payment Or Recept
                If ![Payment] Then ITM.Tag = "P" Else ITM.Tag = "R"
                
                
                
                .MoveNEXT
                LVCheque.Visible = True
                FDrag.Visible = True
            Loop
            
        Else
            LVCheque.Visible = False
            FDrag.Visible = False
        End If
        .Close
    End With
    
   If LVCheque.ListItems.count > 0 Then
      Dim RSTOTAL As New ADODB.Recordset
      RSTOTAL.Open "Select Sum(amount) as TotalAmt from cheque where (CDate between '" & DtFrom & "' and '" & DtTo & "') and Payment=" & Abs(Payment), con
      If Val(RSTOTAL![TotalAmt] & "") > 0 Then
         With LVCheque.ListItems.add().ListSubItems
            .add
            .add
            .add(, , IIf(Payment, "                      Total Payment:", "                  Total Receipt:")).Bold = True
            .add(, , Format(Val(RSTOTAL![TotalAmt] & ""), "#,##0.00")).Bold = True
         End With
      End If

      RSTOTAL.Close
      Set RSTOTAL = Nothing
   End If
   
   
   Call Form_Resize
    Set CrsLedger = Nothing
    Screen.MousePointer = vbDefault
End Sub


Private Sub LV_DblClick()
    If LV.ListItems.count = 0 Or OutOfFinancialYear Then Exit Sub
    If InStr(LV.SelectedItem, "-") > 0 Then Call mnuEditDesc_Click
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)

    If LV.SelectedItem.SubItems(2) = "" Then Exit Sub
    Dim strAccNo As String, dBalance As Double
    strAccNo = LV.SelectedItem.ListSubItems(2).Tag
    dBalance = GetSingleDoubleValue("Balance", "Accounts", " WHERE AccNo='" & strAccNo & "'")
    If dBalance >= 0 Then
        txtLastBalance = Format(dBalance, "#,##0.00") & " DR"
    Else
        txtLastBalance = Format(Abs(dBalance), "#,##0.00") & " CR"
    End If
    
    txtUnAuthDebit = GetSingleLongValue("SUM(Debit)", "VouchersForAuthorization", " WHERE AccNo='" & strAccNo & "' AND Authorized=0")
    txtUnAuthCredit = GetSingleLongValue("SUM(Credit)", "VouchersForAuthorization", " WHERE AccNo='" & strAccNo & "' AND Authorized=0")
    
    Dim dNetBalance As Double
    dNetBalance = dBalance + Val(txtUnAuthDebit) - Val(txtUnAuthCredit)
    If dNetBalance >= 0 Then
        txtNetBalance = Format(dNetBalance, "#,##0.00") & " DR"
    Else
        txtNetBalance = Format(Abs(dNetBalance), "#,##0.00") & " CR"
    End If
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Exit Sub
    If LV.ListItems.count = 0 Or Button <> 2 Or OutOfFinancialYear Then Exit Sub
     
    'Check for Maker Bill....
    Dim strVchrNo As String, strVenderVchrNo As String
    'strVchrNo = GetSingleStringValue("VchrNo", "VendIssued", " WHERE VchrNo='" & LV.SelectedItem.SubItems(1) & "'")
    strVchrNo = GetSingleStringValue("VchrNo", "MakerPostedBills", " WHERE VchrNo='" & LV.SelectedItem.SubItems(1) & "'")
    
    strVenderVchrNo = GetSingleStringValue("VchrNo", "VenderPostedBills", " WHERE VchrNo='" & LV.SelectedItem.SubItems(1) & "'")
    
    mnuPrintVenderBill.Visible = False
    mnuPrintMakerBillOld.Visible = False
    If strVchrNo <> "" Then
        mnuPrintMakerBill.Visible = True
        If strCompany = "QEL" Then
            mnuPrintMakerBillOld.Visible = True
        End If
    Else
        mnuPrintMakerBill.Visible = False
    End If
    
    
    If strVenderVchrNo <> "" Then
        mnuPrintVenderBill.Visible = True
    Else
        mnuPrintVenderBill.Visible = False
    End If
    
    'if a Voucher Entry Show Delete Menu
    If InStr(LV.SelectedItem, "-") > 0 Then
        If Left(LV.SelectedItem.SubItems(1), 3) = "BPV" And Val(LV.SelectedItem.SubItems(4)) > 0 Then
            mnuPrintCheque.Visible = True
        Else
            mnuPrintCheque.Visible = False
        End If
        Me.PopUpMenu mnupop
    End If

End Sub

Private Sub LVCheque_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Item.Checked = Not Item.Checked
End Sub

Private Sub PrintLedger()
With cr1
        

End With
End Sub


Private Sub PrintCashnBankTransaction()

    Dim rpt As CRAXDDRT.Report, rptCashEntries As CRAXDDRT.Report
    Dim rptIOUEntries As CRAXDDRT.Report, rptCashBalances As CRAXDDRT.Report
    Dim rptBankEntries As CRAXDDRT.Report, rptBankBalances As CRAXDDRT.Report
    
    Dim rptApp_Local As New CRAXDDRT.Application
    
    Set rpt = rptApp_Local.OpenReport(RptPath & "\TransRptCashnBank.rpt")
    rpt.FormulaFields.GetItemByName("Company").Text = "'" & strCompany & "'"
    rpt.FormulaFields.GetItemByName("FromTo").Text = "'" & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
    rpt.FormulaFields.GetItemByName("Address").Text = "'" & COMPANYADDRESS & "'"
    rpt.FormulaFields.GetItemByName("DTFrom").Text = "#" & DtFrom.Value & "#"
    rpt.FormulaFields.GetItemByName("DTTo").Text = "#" & DtTo.Value & "#"
      
    '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\\/
    'Cash Entries
    Set rptCashEntries = rpt.OpenSubreport("CashEntries")
    rptCashEntries.RecordSelectionFormula = "{VLedger.VDate}=#" & DtFrom & "# TO #" & DtTo & "# AND LEFT({VLedger.Accno},6)='" & CashInHand & "'"
    rptCashEntries.DiscardSavedData
    '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\\/
    
    '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\\/
    'IOU Entries
    Set rptIOUEntries = rpt.OpenSubreport("IOUEntries")
    rptIOUEntries.RecordSelectionFormula = "{VLedger.VDate}=#" & DtFrom & "# TO #" & DtTo & "# AND {VLedger.SubAccOf}='15-015'"
    'rptCashEntries.DiscardSavedData
    '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\\/
    
    '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\\/
    'Cash Balances
    Set rptCashBalances = rpt.OpenSubreport("CashBalances")
    rptCashBalances.RecordSelectionFormula = "LEFT({Accounts.AccNo},7)='" & CashInHand & "-' OR {Accounts.SubAccOf}='15-015'"
    rptCashBalances.ParameterFields.GetItemByName("@DT").SetCurrentValue DtTo.Value
'    rptCashBalances.DiscardSavedData
    '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\\/
    
    '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\\/
    'Bank Entries
    Set rptBankEntries = rpt.OpenSubreport("BankEntries")
    rptBankEntries.RecordSelectionFormula = "{VLedger.VDate}=#" & DtFrom & "# TO #" & DtTo & "# AND {VLedger.SubAccOf}='15-002'"
    '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\\/
    
    '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\\/
    'Bank Balances
    Set rptBankBalances = rpt.OpenSubreport("BankBalances")
    rptBankBalances.RecordSelectionFormula = "{Accounts.SubAccOf}='15-002'"
    rptBankBalances.ParameterFields.GetItemByName("@DT").SetCurrentValue DtTo.Value
'    rptCashBalances.DiscardSavedData
    '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\\/
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, , False
End Sub
    

Private Sub mnuDelete_Click()
    
    
    Dim frmDelVchr As New frmDeleteVoucher
    Load frmDelVchr
    frmDelVchr.ShowMe LV.SelectedItem.SubItems(1)
    Exit Sub
    'Code Below was Used to Delete It
'    If Not UserHasAccess("DeleteVoucher") Then
'       MsgBox "You Don't Have Permissions To Delete Transaction", vbInformation
'       Exit Sub
'    End If
'    Dim getPassword As Boolean
'    getPassword = getDBPassword("DeleteVoucher")
'
'    If getPassword Then
'        If MsgBox("Are You Sure To Delete Voucher # : " & LV.SelectedItem.SubItems(1), vbQuestion + vbYesNo) = vbNo Then Exit Sub
'        'If Not getPassword Then Exit Sub
'
'        If DeleteVoucher(LV.SelectedItem.SubItems(1)) Then MsgBox "Voucher Deleted Successfully.", vbInformation
'        '/\\/\/\\/\\/\\/\/\/\/\/\\/\/\/\/\/\/\/\/\/\/\///////////////////////////
'        'Delete From Advances..If It's An Advance Payment Entry To Employee.
'        con.Execute "Delete From Advances Where AccVoucherNo='" & LV.SelectedItem.SubItems(1) & "'"
'        '/\\/\/\\/\\/\\/\/\/\/\/\\/\/\/\/\/\/\/\/\/\/\///////////////////////////
'        Call cmdRefresh_Click
'    End If
    
End Sub

Private Sub mnuEditDesc_Click()

    'Check For User Access
    If Not UserHasAccess("EditVoucherDesc") Then
        MsgBox "You Don't Have Permissions  To Edit Transaction", vbInformation
        Exit Sub
    End If

Dim NewDesc As String, OldDesc As String, DpstPos As Long, DpstSlip As String
DpstPos = InStr(LV.SelectedItem.SubItems(3), " Dpst Slip :")

 If DpstPos > 0 Then
    OldDesc = Left(LV.SelectedItem.SubItems(3), DpstPos - 1)
    DpstSlip = Right(LV.SelectedItem.SubItems(3), Len(LV.SelectedItem.SubItems(3)) - DpstPos - 1)
 Else
    OldDesc = LV.SelectedItem.SubItems(3)
    DpstSlip = ""
 End If
 
    'NewDesc = InputBox("Enter Description.", "Edit Description", OldDesc)
 
    With txtDesc
        .Move LV.Left + LV.ColumnHeaders(4).Left + 85, LV.Top + LV.SelectedItem.Top + 25, LV.ColumnHeaders(4).Width ', LV.SelectedItem.Height
        .Text = OldDesc
        Set .Font = LV.Font
        .Tag = DpstSlip
        .Visible = True
        .SetFocus
    End With

End Sub

Private Sub mnuPreviewVoucher_Click()
    
    If LV.ListItems.count = 0 Then Exit Sub
    Call PrintVoucher(MainForm.cr1, LV.SelectedItem.SubItems(1), , False)
    
End Sub

Private Sub mnuPrintCheque_Click()
    Dim f As New frmPrintCheque
    f.ShowMe (LV.SelectedItem.Tag)
End Sub

Private Sub mnuPrintMakerBill_Click()
    Call PrintMakerPostedBill(LV.SelectedItem.SubItems(1))
End Sub

Private Sub mnuPrintMakerBillOld_Click()
    Call PrintMakerPostedBill(LV.SelectedItem.SubItems(1), True)
End Sub

Private Sub mnuprintOnblank_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    Call PrintVoucher(MainForm.cr1, LV.SelectedItem.SubItems(1), , True)
 
End Sub

Private Sub mnuprintOnblank1_Click()
    Call mnuprintOnblank_Click
End Sub

Private Sub mnuPrintVenderBill_Click()
    Call PrintVenderPostedBill(LV.SelectedItem.SubItems(1))
End Sub

Private Sub mnuPrintVoucher_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    
    If LCase(strCompany) Like "helicon*" Then
        Call PrintVoucher(MainForm.cr1, LV.SelectedItem.SubItems(1), True, True)
    Else
        mnuprintOnblank_Click
    End If
    
End Sub

Private Sub mnuPrintVoucher1_Click()
    Call mnuPrintVoucher_Click
End Sub


Private Sub RefreshCashLedger()

    Screen.MousePointer = vbHourglass
    
    Dim cmd As New ADODB.Command, rs As New ADODB.Recordset
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "CashBookReport_SP"
        .Parameters("@DTFrom").Value = DtFrom.Value
        .Parameters("@DTTo").Value = DtTo.Value
        Set rs = .Execute
    End With
   
    Dim PrevDate As Date

    LV.ListItems.Clear
    LVCheque.ListItems.Clear

    Dim ITM As ListItem, Balance As String
    Dim VType As String
    Dim BankCondition As String
    Dim dLastCashInHandBalance As Double
    Dim dTotalDebit As Double, dTotalCredit As Double
    With rs
         ''''''''''''' Show Transactions'''''''''''
        If .EOF = False Then
            With LV.ListItems.add().ListSubItems
               .add
               .add
               .add(, , "Opening Cash in Hand").Bold = True
                If Val(rs!OpeningBalance & "") >= 0 Then
                    .add(, , Format(Val(rs!OpeningBalance & ""), "#,##0.00")).Bold = True
                    .add
                Else
                    .add
                    .add(, , Format(Val(rs!OpeningBalance & ""), "#,##0.00")).Bold = True
                End If
            End With
        End If
        
        For i = 0 To .RecordCount - 1
           
            Set ITM = LV.ListItems.add(, ![SNo] & "'" & Rnd, Format(![Vdate], "dd-MMM-yyyy"))
            If ![Vdate] <> PrevDate Then
               PrevDate = ![Vdate]
            Else
                ITM.ForeColor = RGB(125, 125, 225)
            End If
         
            ITM.ListSubItems.add , , ![VchrNo] & ""
            ITM.ListSubItems.add(, , ![AccTitle] & "").ToolTipText = ![AccTitle] & ""
            
            Dim Dpst As String
            
            If Val(![DpstSlip] & "") = 0 Then Dpst = "" Else Dpst = " Dpst Slip:" & ![DpstSlip]
        
            ITM.ListSubItems.add(, , ![Description] & Dpst).ToolTipText = ![Description] & Dpst
            ITM.SubItems(4) = Format(Val(![Debit] & ""), "###,##0.00")
            ITM.SubItems(5) = Format(Val(![Credit] & ""), "###,##0.00")
            ITM.Tag = Val(![CSNo] & "")
            dTotalDebit = dTotalDebit + Val(!Debit & "")
            dTotalCredit = dTotalCredit + Val(!Credit & "")
'            If !AccNo = "15-001-03001" Then
'                dLastCashInHandBalance = Val(!Balance & "")
'            End If
            dLastCashInHandBalance = Val(!CashBalance & "")
            .MoveNEXT

        Next
      
    End With
        
    With LV.ListItems.add().ListSubItems
       .add
       .add
       .add(, , TotalCaption).Bold = True
       .add(, , Format(dTotalDebit, "#,##0.00")).Bold = True
       .add(, , Format(dTotalCredit, "#,##0.00")).Bold = True
    End With
    
    'If strCompany = "Instrumed" Then
        With LV.ListItems.add().ListSubItems
           .add
           .add
           .add(, , "Closing Cash in Hand").Bold = True
            If dLastCashInHandBalance >= 0 Then
                .add(, , Format(dLastCashInHandBalance, "#,##0.00")).Bold = True
                .add
            Else
                .add
                .add(, , Format(dLastCashInHandBalance, "#,##0.00")).Bold = True
            End If
        End With
    'End If
   
    Screen.MousePointer = vbDefault

End Sub



Private Sub txtDesc_GotFocus()
    cmdRefresh.Default = False
    txtDesc.SelStart = 0
    txtDesc.SelLength = Len(txtDesc)
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
  
If KeyAscii = 13 Then
    
    If Trim(txtDesc = "") Then
        MsgBox "Invalid Description", vbInformation
        Exit Sub
    End If
    
   con.Execute "Update Vouchers Set [Description]='" & txtDesc & "' Where Sno=" & Val(LV.SelectedItem.key)
   LV.SelectedItem.ListSubItems(3) = txtDesc & txtDesc.Tag
   txtDesc.Visible = False
   cmdRefresh.Default = False
ElseIf KeyAscii = 27 Then
   txtDesc.Visible = False
   cmdRefresh.Default = False
End If

End Sub

Private Sub txtDesc_LostFocus()
    txtDesc.Visible = False
End Sub
