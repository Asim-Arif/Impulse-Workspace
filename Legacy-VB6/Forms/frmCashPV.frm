VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCashPV 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment Entry ..."
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10875
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   7485
   ScaleWidth      =   10875
   WindowState     =   2  'Maximized
   Begin VB.Frame FALL 
      Height          =   7350
      Left            =   105
      TabIndex        =   17
      Top             =   45
      Width           =   10695
      Begin VB.Frame Frame1 
         Height          =   645
         Left            =   0
         TabIndex        =   33
         Top             =   45
         Width           =   10695
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Voucher"
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
            Index           =   0
            Left            =   60
            TabIndex        =   34
            Top             =   165
            Width           =   10605
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Voucher"
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
            Height          =   465
            Index           =   1
            Left            =   75
            TabIndex        =   35
            Top             =   180
            Width           =   10590
         End
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1590
         Left            =   5490
         TabIndex        =   32
         Top             =   2490
         Visible         =   0   'False
         Width           =   4320
      End
      Begin VB.TextBox Tbox 
         Height          =   315
         Index           =   0
         Left            =   1665
         TabIndex        =   4
         Top             =   1470
         Width           =   4545
      End
      Begin VB.TextBox Tbox 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   9165
         TabIndex        =   8
         Top             =   1470
         Width           =   1335
      End
      Begin VB.Frame FraCash 
         BorderStyle     =   0  'None
         Height          =   4785
         Left            =   150
         TabIndex        =   22
         Top             =   2445
         Width           =   10395
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
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   60
            TabIndex        =   38
            Top             =   4485
            Value           =   1  'Checked
            Width           =   1830
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H80000004&
            Height          =   660
            Left            =   5010
            TabIndex        =   36
            Top             =   4080
            Width           =   5280
            Begin MSForms.CommandButton cmdSClose 
               Height          =   360
               Left            =   1800
               TabIndex        =   15
               Top             =   195
               Width           =   1665
               Caption         =   "Save & Close  "
               PicturePosition =   327683
               Size            =   "2937;635"
               Accelerator     =   83
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
               Left            =   3510
               TabIndex        =   16
               Top             =   195
               Width           =   1665
               Caption         =   "Cancel            "
               PicturePosition =   327683
               Size            =   "2937;635"
               Picture         =   "frmCashPV.frx":0000
               Accelerator     =   67
               FontName        =   "Tahoma"
               FontEffects     =   1073741825
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
               FontWeight      =   700
            End
            Begin MSForms.CommandButton cmdSClear 
               Height          =   360
               Left            =   75
               TabIndex        =   14
               Top             =   195
               Width           =   1665
               Caption         =   "Save & New   "
               PicturePosition =   327683
               Size            =   "2937;635"
               Picture         =   "frmCashPV.frx":0112
               Accelerator     =   78
               FontName        =   "Tahoma"
               FontEffects     =   1073741825
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
               FontWeight      =   700
            End
         End
         Begin MSComctlLib.ListView LVPymt 
            Height          =   3180
            Left            =   30
            TabIndex        =   12
            Top             =   0
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   5609
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   8388608
            BackColor       =   15199215
            BorderStyle     =   1
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
               Text            =   "Voucher #"
               Object.Width           =   3757
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description"
               Object.Width           =   5468
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Key             =   "Amt"
               Text            =   "Total Amout"
               Object.Width           =   2081
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Amount Due"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Applied Pymt"
               Object.Width           =   2028
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Status"
               Object.Width           =   2205
            EndProperty
         End
         Begin MSForms.CommandButton cmd 
            Height          =   360
            Left            =   45
            TabIndex        =   13
            Top             =   3315
            Width           =   2010
            VariousPropertyBits=   25
            Caption         =   "Auto Apply         "
            PicturePosition =   327683
            Size            =   "3545;635"
            Picture         =   "frmCashPV.frx":0224
            Accelerator     =   80
            FontName        =   "Tahoma"
            FontEffects     =   1073750017
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label lblToApply 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   8790
            TabIndex        =   26
            Top             =   3660
            Width           =   1515
         End
         Begin VB.Label lblApplied 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   8790
            TabIndex        =   25
            Top             =   3315
            Width           =   1515
         End
         Begin VB.Label Lbls 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Un Applied Amount :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   3
            Left            =   6720
            TabIndex        =   24
            Top             =   3690
            Width           =   1755
         End
         Begin VB.Label Lbls 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Amount Applied :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   2
            Left            =   6525
            TabIndex        =   23
            Top             =   3345
            Width           =   1950
         End
      End
      Begin VB.Frame FraBank 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   180
         TabIndex        =   18
         Top             =   1785
         Width           =   10380
         Begin VB.CheckBox ChkClear 
            Caption         =   "Clear Cheque "
            Height          =   225
            Left            =   8985
            TabIndex        =   37
            ToolTipText     =   "Clear This Cheque Automatically"
            Top             =   345
            Value           =   1  'Checked
            Width           =   1350
         End
         Begin VB.TextBox Tbox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   4545
            TabIndex        =   9
            Top             =   285
            Width           =   1335
         End
         Begin VB.CommandButton cmdBank 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4185
            Picture         =   "frmCashPV.frx":061E
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   285
            Width           =   360
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   7530
            TabIndex        =   11
            Top             =   285
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   47382531
            CurrentDate     =   37250
         End
         Begin MSForms.ComboBox cmbCheckType 
            Height          =   315
            Left            =   5880
            TabIndex        =   10
            Top             =   285
            Width           =   1650
            VariousPropertyBits=   679495707
            DisplayStyle    =   7
            Size            =   "2910;556"
            ListWidth       =   2822
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label2 
            Height          =   255
            Index           =   1
            Left            =   15
            TabIndex        =   20
            Top             =   30
            Width           =   10305
            ForeColor       =   4194304
            BackColor       =   12504280
            Caption         =   $"frmCashPV.frx":0990
            Size            =   "18177;450"
            BorderColor     =   14937593
            BorderStyle     =   1
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label lblBankAcc 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   15
            TabIndex        =   19
            Top             =   300
            Width           =   4170
         End
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   975
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dddd, dd/MM/yyyy"
         Format          =   47382531
         CurrentDate     =   37250
      End
      Begin VB.TextBox Tbox 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   6240
         TabIndex        =   5
         Top             =   1470
         Width           =   1335
      End
      Begin MSForms.ComboBox CmbAccTitle 
         Height          =   315
         Left            =   5760
         TabIndex        =   2
         Top             =   960
         Width           =   3375
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "5953;556"
         ListWidth       =   8819
         ListStyle       =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   9135
         TabIndex        =   27
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label lblBal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   7590
         TabIndex        =   6
         Top             =   1470
         Width           =   1560
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vocuher # "
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
         Index           =   3
         Left            =   240
         TabIndex        =   31
         Top             =   765
         Width           =   945
      End
      Begin VB.Label lblBalCap 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash in Hand  "
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
         Left            =   7605
         TabIndex        =   29
         Top             =   1275
         Width           =   1185
      End
      Begin MSForms.ComboBox cmbAccNo 
         Height          =   315
         Left            =   3840
         TabIndex        =   1
         Top             =   960
         Width           =   1920
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3387;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbPymtmode 
         Height          =   315
         Left            =   195
         TabIndex        =   3
         Top             =   1470
         Width           =   1470
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "2593;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   2
         Left            =   195
         TabIndex        =   30
         Top             =   750
         Width           =   10305
         ForeColor       =   4194304
         BackColor       =   12504280
         Caption         =   $"frmCashPV.frx":0A1D
         Size            =   "18177;397"
         BorderColor     =   14937593
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   4
         Left            =   210
         TabIndex        =   28
         Top             =   975
         Width           =   1455
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   0
         Left            =   195
         TabIndex        =   21
         Top             =   1260
         Width           =   10305
         ForeColor       =   4194304
         BackColor       =   12504280
         Caption         =   $"frmCashPV.frx":0AE1
         Size            =   "18177;397"
         BorderColor     =   14937593
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
End
Attribute VB_Name = "frmCashPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DontShowList As Boolean

Dim PermissionDenied As Boolean

Private Sub CmbAccTitle_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDown Then
   If List1.ListIndex < List1.ListCount - 1 Then List1.ListIndex = List1.ListIndex + 1
   KeyCode = 0
   DontShowList = True
ElseIf KeyCode = vbKeyUp Then
   If List1.ListIndex > 0 Then List1.ListIndex = List1.ListIndex - 1
   KeyCode = 0
   DontShowList = True
ElseIf KeyCode = 13 And List1.Visible Then
   CmbAccTitle = List1
   List1.Visible = False
   CmbAccTitle.SetFocus
End If

End Sub

Private Sub CmbAccTitle_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If DontShowList Then DontShowList = False Else ShowList List1, CmbAccTitle, CmbAccTitle
End Sub
Private Sub CmbAccTitle_LostFocus()
List1.Visible = False
End Sub


Private Sub DTPicker2_Change()
   Call setDateRange(DTPicker2)
   ChkClear.Enabled = (DTPicker1 = DTPicker2)
End Sub

Private Sub Form_Activate()
'Check For User Access
If PermissionDenied Then
   Me.Hide
   MsgBox "You Donot Have Permissions To Use This Option. Please Contact Your Administrator", vbInformation
   Unload Me
   Exit Sub
End If


End Sub

Private Sub List1_Click()
If UCase(Screen.ActiveControl.Name) <> "CMBACCTITLE" Then
   CmbAccTitle = List1
   List1.Visible = False
   CmbAccTitle.SetFocus
End If
End Sub
Private Sub cmbAccNo_Change()
    If cmbAccNo.MatchFound Then Call cmbAccNo_Click Else If Not cmbAccNo = "" Then cmbAccNo.DropDown
End Sub
Private Sub cmbAccNo_Click()
    If CmbAccTitle.ListIndex <> cmbAccNo.ListIndex Then CmbAccTitle.ListIndex = cmbAccNo.ListIndex
    lblBalance.Visible = cmbAccNo.MatchFound
    Call RefreshList
End Sub

Private Sub CmbAccTitle_Change()
    If CmbAccTitle.MatchFound Then Call cmbAccTitle_Click 'Else If CmbAccTitle <> "" Then CmbAccTitle.DropDown
End Sub
Private Sub cmbAccTitle_Click()
    If cmbAccNo.ListIndex <> CmbAccTitle.ListIndex Then cmbAccNo.ListIndex = CmbAccTitle.ListIndex
End Sub
Private Sub cmbAccNo_Validate(Cancel As Boolean)
    If Not cmbAccNo.MatchFound Then cmbAccNo.ListIndex = CmbAccTitle.ListIndex
End Sub
Private Sub CmbAccTitle_Validate(Cancel As Boolean)
   If Not CmbAccTitle.MatchFound Then CmbAccTitle.ListIndex = cmbAccNo.ListIndex
End Sub
Private Sub CmbAccTitle_DropButtonClick()
   List1.Visible = False
End Sub


Private Sub cmbPymtmode_Change()
    Call cmbPymtmode_Click
End Sub
Private Sub cmbPymtmode_Click()

    If cmbPymtmode.ListIndex = 0 Then
        
        FraBank.Visible = False
        FraCash.Top = 1920
        Label1(3).Visible = True
        Label1(4).Visible = True
        chkPrint.Visible = True
        TBox(0) = "Cash Paid "
        FALL.Height = 6780
        
    Else
    
        If Not AnyBankAccountDefined Then
           MsgBox "Payment Cannot Be Made By Cheque, No Bank Accounts Defined.", vbInformation
           cmbPymtmode.ListIndex = 0
           Exit Sub
        End If
    
        TBox(0) = "Cheque Paid"
        FraBank.Visible = True
        Label1(3).Visible = False
        Label1(4).Visible = False
        cmbCheckType.ListIndex = 0
        FraCash.Top = 2550
        FALL.Height = 7400
        'Get Default BankID
        Dim BankID As Double
         lblBankAcc = GetDefaultBankAcc(BankID)
         lblBankAcc.Tag = BankID
         'Get Next Chq #
        Dim ChqBookNo As Double
         TBox(2) = GetNextChqNo(getID(lblBankAcc), ChqBookNo)
         TBox(2).Tag = ChqBookNo
         DTPicker2 = DTPicker1
         'chkPrint.Visible = False
         
    End If
    Call ShowBals
    Call Form_Resize
End Sub

Public Sub cmd_Click()

With LVPymt
    If Trim(cmd.Caption) = "Auto Apply" Then
        For i = 1 To .ListItems.Count
            
            .ListItems(i).Checked = True
            Call LVPymt_ItemCheck(.ListItems(i))
            If Val(lblToApply) = 0 Then Exit For
        Next
        If i > 1 Then cmd.Caption = "Clear Payments     "
    Else
        For i = 1 To .ListItems.Count
            .ListItems(i).Checked = False
            Call LVPymt_ItemCheck(.ListItems(i))
        Next
        lblToApply = TBox(1)
        lblApplied = 0
        cmd.Caption = "Auto Apply         "
    End If
End With
End Sub

Private Sub cmdBank_Click()

    Dim BankID As Double
    lblBankAcc = SelectBankAcc(BankID)
    lblBankAcc.Tag = BankID
    
    'Get ChqNo
         
    Dim ChqBookNo As Double
    TBox(2) = GetNextChqNo(getID(lblBankAcc), ChqBookNo)
    TBox(2).Tag = ChqBookNo
        
End Sub

Private Sub cmdSClear_Click()

If Not AskForAutoApply Then Exit Sub

If cmbPymtmode.ListIndex = 0 Then
    If Not SaveTransForCash Then Exit Sub
Else
    If Not SaveTransForCheque Then Exit Sub
End If

Unload Me
Me.Show
Me.SetFocus

'Call ClearFields

End Sub

Private Sub cmdSClose_Click()
If Not AskForAutoApply Then Exit Sub


If cmbPymtmode.ListIndex = 0 Then
    If Not SaveTransForCash Then Exit Sub
Else
    If Not SaveTransForCheque Then Exit Sub
    
End If

Unload Me

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub DtPicker1_Change()
    
    Call setDateRange(DTPicker1)
    
    Label1(4).Caption = getNextVchrNo(DTPicker1.value, "PV")
    If cmbAccNo.MatchFound Then Call RefreshList
    
End Sub

Private Sub DTPicker1_Click()
    Call DtPicker1_Change
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
PermissionDenied = Not UserHasAccess("PaymentVoucher")
     
    Me.KeyPreview = True
    DTPicker1.value = Date
    DTPicker2.value = Date
    Call DTPicker1_Click
    Call FillAccCmbs
    
   chkPrint = IIf(frmMain.mnuPrintOnSave.Checked, vbChecked, vbUnchecked)
   
    
   ' CashInHand = getAccountNo("Cash In Hand")
    
   ' If CashInHand = "" Then
   '     MsgBox "No Cash in Hand Account is defined", vbInformation
   '     Exit Sub
   ' End If
    
End Sub

Private Sub FillAccCmbs()

    
        'Clear All The Combos
        cmbAccNo.Clear
        CmbAccTitle.Clear
        cmbPymtmode.Clear
'        cmbBank.Clear
'        cmbBankBranch.Clear
        cmbCheckType.Clear
        
        'Fill With Payment Mode
        cmbPymtmode.AddItem "Cash"
        cmbPymtmode.AddItem "Cheque"
        cmbPymtmode.ListIndex = 0
        
        'Fill With CheckType
        cmbCheckType.AddItem "Open"
        cmbCheckType.AddItem "Payees A/C Only"
        cmbCheckType.ListIndex = 0
        
        
    Call AddToCombo(cmbAccNo, "AccNo", "VactiveAccounts", "Where AccNo like '24%' and Not Parent Order By AccNo")
    Call AddToCombo(CmbAccTitle, "AccTitle", "VactiveAccounts", "Where AccNo like '24%' and Not Parent  Order By AccNo")
    'Call AddToCombo(cmbBank, "Distinct Bank", "BankList", "Order By Bank")
    
    
End Sub







Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If UnloadMode = 0 Then
'    If LVPymt.ListItems.Count > 0 Then
'        If MsgBox("There is a Transaction active." & vbNewLine & "Save or Cancel Transaction first.", vbInformation + vbOKOnly) Then
'            Cancel = True
'        End If
'    End If
'End If
End Sub










Private Sub Form_Resize()
   
   FALL.Move (Me.ScaleWidth - FALL.Width) / 2, (Me.ScaleHeight - FALL.Height) / 2
End Sub

Private Sub lblBankAcc_Change()
Call ShowBals
End Sub

Private Sub Tbox_Change(Index As Integer)
'    If Index = 1 Then lblToApply = Tbox(1)
End Sub

Private Sub Tbox_KeyPress(Index As Integer, KeyAscii As Integer)

If Index = 1 Then
    KeyAscii = OnlyNumber(KeyAscii, True)
End If

End Sub
Private Sub ClearLView()

Dim iTotal As Integer
iTotal = LView.ListItems.Count
    For i = 1 To iTotal
        LView.ListItems.Remove (1)
    Next i
    
TBox(0).Text = ""
TBox(1).Text = ""
cmdDelete.Visible = False
Call DTPicker1_Click

End Sub


Private Sub LVPymt_ItemCheck(ByVal Item As MSComctlLib.ListItem)

If Not Item.Checked Then
    lblApplied = Format(Val(lblApplied) - Val(Item.SubItems(4)), "0.00")
    lblToApply = Format(Val(lblToApply) + Val(Item.SubItems(4)), "0.00")
    Item.SubItems(4) = 0
    Item.SubItems(5) = "DUE"
Else
 
    If Val(lblToApply) = 0 Then
        MsgBox "No Amount To Apply", vbInformation
        Item.Checked = False
        
        Exit Sub
    End If
 
    If Item.SubItems(3) <= Val(lblToApply) Then
        Item.SubItems(4) = Format(Val(Item.SubItems(3)), "0.00")
        Item.SubItems(5) = "FULL PAID"
   
    Else
        Item.SubItems(4) = Format(Val(lblToApply), "0.00")
        Item.SubItems(5) = "PARTIAL PAID"
    End If
  
    lblApplied = Format(Val(lblApplied) + Val(Item.SubItems(4)), "0.00")
    lblToApply = Format(Val(lblToApply) - Val(Item.SubItems(4)), "0.00")
  
End If

   If lblToApply = TBox(1) Then
      cmd.Caption = "Auto Apply      "
   Else
      cmd.Caption = "Clear Payments    "
   End If

'If Val(lblapply) = 0 Then
' cmdSave.Enabled = True
'Else
' cmdSave.Enabled = False
'End If

End Sub



Private Sub RefreshList()

Screen.MousePointer = vbHourglass
Dim rsList As New ADODB.Recordset
 
rsList.CursorLocation = adUseClient
rsList.Open "Select Sno,VchrNo,Desc,TotalAmount,AmountLeft  from AccPay where (AccNo='" & cmbAccNo & "') and (VDate<=#" & Format(DTPicker1.value, "dd-MMM-yyyy") & "#) and ((AmountLeft>0) or (AmountLeft is null))", con, adOpenForwardOnly, adLockReadOnly
rsList.Sort = "Sno"
 
Call ShowAccBalance

 
With LVPymt.ListItems
    .Clear

cmd.Caption = "Auto Apply         "
    Dim itm As ListItem
    Do Until rsList.EOF
        Set itm = .Add(, rsList![Sno] & "?", rsList![VchrNo])
        itm.SubItems(1) = rsList![Desc] & ""
        itm.SubItems(2) = rsList![TotalAmount] & ""
        If IsNull(rsList![AmountLeft]) Then
            itm.SubItems(3) = rsList![TotalAmount]
        Else
            itm.SubItems(3) = rsList![AmountLeft] & ""
        End If
        
        itm.SubItems(4) = 0
        itm.SubItems(5) = "DUE"
        rsList.MoveNext
    Loop
    If .Count > 0 Then cmd.Enabled = True Else cmd.Enabled = False
End With

'Reset The Amount
Call Tbox_Validate(1, False)
Call ShowBals
Screen.MousePointer = vbDefault
End Sub

Private Function SaveTransForCash() As Boolean

If CheckData = False Then Exit Function

    
    Dim i As Integer
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim Bal As Double
    Dim RcptNo As String
    'Get Serial # and Balances
    INextSno = getNextSno(DTPicker1.value)
    
    RcptNo = IIf(Len(Trim(TBox(3))) = 0, "", "Rcpt # :" & TBox(3))
    
    Bal = getBalance(CASHINHAND, DTPicker1.value)
    Bal = Bal - Val(TBox(1))
    'Start New Transaction
    con.BeginTrans
    'Credit Entry
    con.Execute "insert into Vouchers Values(" & INextSno & ",#" & _
         Format(DTPicker1.value, "dd-MMM-yyyy") & "#,'" & Label1(4) & "','" & CASHINHAND & _
        "','Cash Paid To " & CmbAccTitle & "',0," & Val(TBox(1)) & "," & Bal & ",'',0)"
        
    
    
    Call adjustLedger(CASHINHAND, DTPicker1.value, INextSno, -Val(TBox(1)))
    
    INextSno = INextSno + 1
    SnoToUse = INextSno
    Bal = getBalance(cmbAccNo, DTPicker1.value)
    Bal = Bal + Val(TBox(1))
    'Debit Entry
    con.Execute "insert into Vouchers Values(" & INextSno & ",#" & _
         Format(DTPicker1.value, "dd-MMM-yyyy") & "#,'" & Label1(4) & "','" & cmbAccNo & _
        "','" & TBox(0) & RcptNo & "'," & TBox(1) & _
        ",0," & Bal & ",''," & Val(CSNo) & ")"
    
    
    Call adjustLedger(cmbAccNo, DTPicker1.value, INextSno, Val(TBox(1)))
    
    
    
    With LVPymt.ListItems
        For i = 1 To .Count
            If .Item(i).Checked Then con.Execute "Insert into VoucherDetail Values('" & Val(.Item(i).Key) & "','" & SnoToUse & "'," & Val(LVPymt.ListItems(i).SubItems(4)) & ")"
        Next
    End With
    
    'Commit Transaction
    con.CommitTrans
        
    Call PlaySaveSound
    
    'Print If To Be Printed
    
         If chkPrint = vbChecked Then Call PrintVoucher(frmMain.cr1, Label1(4), True)
      

    
    MsgBox "Transaction Saved Successfuly.", vbInformation
    
    'FillAccCmbs
    SaveTransForCash = True
    Exit Function
ERR:
    con.RollbackTrans
    MsgBox "Transaction Terminated Due to ERROR:" & ERR.Number & vbNewLine & ERR.Description, vbCritical
    SaveTransForCash = False
    
End Function


Private Function CheckData() As Boolean

On Error GoTo ERR

'If LVPymt.ListItems.Count = 0 Then
'    MsgBox "No Transaction to save!", vbInformation
'    Exit Function
'End If
    
If cmd.Caption = "Auto A&pply" And cmd.Enabled Then
    If MsgBox("No Amount is Applied. Do You Want To Apply Automatically.", vbQuestion + vbYesNo) = vbYes Then
        Call cmd_Click
    Else
        Exit Function
    End If
End If

If (cmbAccNo = "") Or (Not cmbAccNo.MatchFound) Then
    MsgBox "Invalid Account Selected...", vbInformation
    cmbAccNo.SetFocus
    Exit Function
ElseIf TBox(0) = "" Then
    MsgBox "No Description Given...", vbInformation
    TBox(0).SetFocus
    Exit Function
ElseIf Val(TBox(1) & "") = 0 Then
    MsgBox "No Amount Given...", vbInformation
    TBox(0).SetFocus
    Exit Function

End If


Dim Issued As Boolean, ChqBookNo As Double
If cmbPymtmode.ListIndex = 1 Then
    If lblBankAcc = "" Then
      MsgBox "Invalid Bank Selected.", vbInformation
      Exit Function
    ElseIf Not ValidChqNo(getID(lblBankAcc), TBox(2), False, ChqBookNo) Then
      If Issued Then
         MsgBox "The Cheque With Number '" & txtChqNo & "' Has Been Issued.", vbInformation
      Else
         MsgBox "Invalid Cheque No.", vbInformation
      End If
      TBox(2).SetFocus
      Exit Function
    End If
    TBox(2).Tag = ChqBookNo
End If



''''''''''''''''''''''Check For Voucher No. Used Or Not '''''''''''''''''

Dim VNo As String
VNo = getNextVchrNo(DTPicker1.value, "PV")
If Label1(4) <> VNo Then
   MsgBox "Voucher No. : " & Label1(4) & " Has Been Used By Other Transaction." & vbNewLine & _
       "Voucher No. : " & VNo & " Will Be Used For This Transaction.", vbInformation
       Label1(4) = VNo
End If



CheckData = True

Exit Function



ERR:
    MsgBox ERR.Description
    
End Function



Private Function SaveTransForCheque() As Boolean
On Error GoTo ERR
If CheckData = False Then Exit Function


    Dim i As Integer
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim Bal As Double
    Dim RcptNo As String
    
    RcptNo = IIf(Len(Trim(TBox(3))) = 0, "", "Rcpt # :" & TBox(3))
    
    'Get Serial # and Balances
    INextSno = getNextSno(DTPicker1.value, "Cheque")
    SnoToUse = INextSno
    
    'Start New Transaction
    con.BeginTrans
    bankDesc = Right(lblBankAcc, Len(lblBankAcc) - InStr(lblBankAcc, "]")) & " Chq # :" & TBox(2)
    con.Execute "insert into Cheque Values(" & INextSno & ",#" & Format(DTPicker1.value, "dd-MMM-yyyy") & "#,'" & _
    cmbAccNo & "','" & TBox(0) & RcptNo & " (" & bankDesc & ") '," & TBox(1) & "," & lblBankAcc.Tag & ",'" & _
    TBox(2) & "'," & TBox(2).Tag & ",'" & cmbCheckType & "',#" & DTPicker2 & "#,False,True,False)"
    
    'Remove the Cheque From Cheque book
    Call IssueChq(Val(TBox(2).Tag), TBox(2))
    
    With LVPymt.ListItems
        For i = 1 To .Count
            If .Item(i).Checked Then con.Execute "Insert into ChequeDetail Values('" & SnoToUse & "','" & Val(.Item(i).Key) & "'," & Val(LVPymt.ListItems(i).SubItems(4)) & ")"
        Next
    End With
    
    
    'Commit Transaction
    con.CommitTrans
    
    If ChkClear.Enabled And ChkClear = vbChecked Then
      If Not SavePaymentTrans(SnoToUse) Then GoTo ERR:
    ElseIf chkPrint = vbChecked Then
      Call PrintChq(frmMain.cr1, SnoToUse)
    End If

    
    Call PlaySaveSound
    MsgBox "Transaction Saved Successfuly.", vbInformation
    
    
    SaveTransForCheque = True
    Exit Function
ERR:
    MsgBox "Transaction Terminated Due to ERROR:" & ERR.Number & vbNewLine & ERR.Description, vbCritical
    SaveTransForCheque = False
    con.RollbackTrans
End Function


Private Sub ClearFields()
    Label1(3).Visible = True
    Label1(4).Visible = True
    Call FillAccCmbs
    cmbPymtmode.ListIndex = 0
    TBox(0) = "Cash Recieved"
    TBox(1) = ""
    TBox(2) = ""
End Sub


Private Function AskForAutoApply() As Boolean
AskForAutoApply = False
If Val(TBox(1)) > Val(lblBal) Then
   If Not AllowOD(getID(lblBankAcc)) Then
      MsgBox "Not Much Amount In " & Left(lblBalCap, Len(lblBalCap) - 2) & ".", vbInformation
      Exit Function
   End If
End If



If Trim(cmd.Caption) = "Auto Apply" And LVPymt.ListItems.Count > 0 Then
   
   If MsgBox("No Amounts Applied,Do You Want To Auto Apply", vbYesNo + vbQuestion) = vbYes Then
      Call cmd_Click
      AskForAutoApply = True
   End If
   
Else
      AskForAutoApply = True
End If

'If Val(lblToApply) <> 0 And LVPymt.ListItems.Count > 1 Then
'      MsgBox "There is Some Unapplied Amount, Cannot Save", vbInformation
'      AskForAutoApply = False
'End If

End Function




Private Sub ShowAccBalance()
'Show Previous Balance of Account
If lblBalance.Visible Then
    lblBalance.Caption = Format(getBalance(cmbAccNo, DTPicker1), "#,##0.00 DR;#,##0.00 CR;0.00 ")
    lblBalance.ToolTipText = "Balance of " & cmbAccNo & " as of " & Format(DTPicker1.value, "dd/MM/yyyy")
End If

End Sub

Private Sub Tbox_Validate(Index As Integer, Cancel As Boolean)
If Index <> 1 Then Exit Sub
 lblApplied = 0
 lblToApply = TBox(1)
 If Trim(cmd.Caption) = "Clear Payments" Then
   Call cmd_Click
   Call cmd_Click
 End If
End Sub


Private Sub ShowBals()
If cmbPymtmode.ListIndex = 1 Then
   lblBalCap = "Bank Balance  "
   lblBal = Format(getBalance(getID(lblBankAcc), DTPicker1), "0.00")
ElseIf cmbPymtmode.ListIndex = 0 Then
   lblBalCap = "Cash In Hand  "
   lblBal = Format(getBalance(CASHINHAND, DTPicker1), "0.00")
End If
End Sub



Private Function SavePaymentTrans(ChqSNo As Double) As Boolean

    Dim i As Integer
    Dim NextVchrNo As String
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim dAmount As Double
    Dim sScNo As String
    Dim sCIH As String
    Dim Bal As Double
    Dim Desc As String
    Dim AccNo As String
    Dim Dated As Date
    'Get Serial # and Balances
    Dated = DTPicker2
    
    NextVchrNo = getNextVchrNo(Dated, "PV")
    INextSno = getNextSno(Dated)
    
    SnoToUse = INextSno
    
    'Party Account's Balance
    Bal = getBalance(cmbAccNo, Dated)
    dAmount = TBox(1)
    
    sScNo = ChqSNo
    
    
    'Bank Account No.
    sCIH = getID(lblBankAcc)
    
       
    
    Bal = Bal + dAmount
    
    'Debit Entry
    bankDesc = Right(lblBankAcc, Len(lblBankAcc) - InStr(lblBankAcc, "]")) & " Chq # :" & TBox(2)
    con.Execute "insert into Vouchers Values(" & INextSno & ",#" & _
        Dated & "#,'" & NextVchrNo & "','" & cmbAccNo & _
        "','" & TBox(0) & bankDesc & "'," & dAmount & ",0," & Bal & ",''," & sScNo & ")"
    
    
    
    Call adjustLedger(cmbAccNo, Dated, INextSno, Val(dAmount))
    
    Dim ChqNo As String
    INextSno = INextSno + 1
    Bal = getBalance(sCIH, Dated)
    Bal = Bal - dAmount
    'Description
    
    
      Desc = "Cheque Paid To " & CmbAccTitle & " Chq# : " & TBox(2)
      ChqNo = TBox(2)
    
           
    'Credit Entry
    
    con.Execute "insert into Vouchers Values(" & INextSno & ",#" & _
        Dated & "#,'" & NextVchrNo & "','" & sCIH & _
        "','" & Desc & "',0," & dAmount & "," & Bal & ",''," & sScNo & ")"
        
    
    
    Call adjustLedger(sCIH, Dated, INextSno, -Val(dAmount))
    
    Dim rsCopy As New ADODB.Recordset
    With rsCopy
        .CursorLocation = adUseClient
        .Open "Select * from ChequeDetail Where (Sno=" & sScNo & ")", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            con.Execute "Insert into VoucherDetail Values('" & ![BalancedSno] & "','" & SnoToUse & "'," & ![Amount] & ")"
            .MoveNext
        Loop
        .Close
    End With
    Set rsCopy = Nothing
    con.Execute "Update Cheque Set Posted=True Where Sno=" & sScNo
    
    
   
    
    
    SavePaymentTrans = True
    Exit Function
ERR:
    'MsgBox "Transaction Terminated Due To Following Error;" & vbNewLine & err.Description
     SavePaymentTrans = False
End Function


