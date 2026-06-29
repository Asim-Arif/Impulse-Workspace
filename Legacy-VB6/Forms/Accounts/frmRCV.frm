VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmRCV 
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11910
   ControlBox      =   0   'False
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
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin VB.Frame FALL 
      Height          =   8325
      Left            =   60
      TabIndex        =   19
      Top             =   -30
      Width           =   11835
      Begin VB.TextBox txtRcvdFrom 
         Height          =   315
         Left            =   1500
         MaxLength       =   255
         TabIndex        =   43
         Top             =   6135
         Width           =   3720
      End
      Begin VB.TextBox txtVchrNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   1275
         TabIndex        =   41
         Top             =   1005
         Width           =   1575
      End
      Begin VB.TextBox Tbox 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   10365
         TabIndex        =   13
         Top             =   6765
         Visible         =   0   'False
         Width           =   1395
      End
      Begin ComboList.Usercontrol1 cmbOtherAcc 
         Height          =   285
         Left            =   9735
         TabIndex        =   12
         Top             =   6810
         Visible         =   0   'False
         Width           =   3735
         _ExtentX        =   6588
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
      Begin VB.CommandButton cmdNewAccount 
         Caption         =   "New Account"
         Height          =   255
         Left            =   4410
         TabIndex        =   38
         Top             =   1320
         Width           =   1185
      End
      Begin VB.TextBox Tbox 
         Height          =   870
         Index           =   4
         Left            =   135
         MultiLine       =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   7320
         Width           =   6240
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   0
         TabIndex        =   34
         Top             =   60
         Width           =   11850
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Receipt Voucher"
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
            Index           =   8
            Left            =   75
            TabIndex        =   35
            Top             =   165
            Width           =   11655
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Receipt Voucher"
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
            Index           =   9
            Left            =   90
            TabIndex        =   36
            Top             =   180
            Width           =   11640
         End
      End
      Begin VB.TextBox Tbox 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   8760
         TabIndex        =   4
         Top             =   1545
         Width           =   1080
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
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   6510
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   7305
         Width           =   1830
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1590
         Left            =   1785
         TabIndex        =   33
         Top             =   1875
         Visible         =   0   'False
         Width           =   4320
      End
      Begin VB.CommandButton cmdDelete 
         CausesValidation=   0   'False
         Height          =   285
         Left            =   11175
         Picture         =   "frmRCV.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Remove This Entry From Voucher"
         Top             =   1965
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Post"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   11220
         Picture         =   "frmRCV.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Add This Entry To Voucher"
         Top             =   1305
         UseMaskColor    =   -1  'True
         Width           =   525
      End
      Begin VB.TextBox Tbox 
         Height          =   315
         Index           =   0
         Left            =   5595
         MaxLength       =   255
         TabIndex        =   3
         Top             =   1545
         Width           =   3150
      End
      Begin VB.TextBox Tbox 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   9840
         TabIndex        =   5
         Top             =   1545
         Width           =   1335
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E3EDF9&
         CausesValidation=   0   'False
         Height          =   285
         Left            =   4230
         TabIndex        =   20
         Top             =   2115
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.CheckBox chkBank 
         Caption         =   "Receive In Bank"
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
         Left            =   135
         TabIndex        =   9
         Top             =   6495
         Width           =   1845
      End
      Begin MSComctlLib.ListView LView 
         Height          =   4245
         Left            =   105
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1905
         Width           =   11670
         _ExtentX        =   20585
         _ExtentY        =   7488
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         PictureAlignment=   5
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Acc #"
            Object.Width           =   2937
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Acc Title"
            Object.Width           =   4630
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   7488
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   2646
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   3855
         TabIndex        =   0
         Top             =   960
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   529
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
         CustomFormat    =   "dddd dd/MM/yyyy"
         Format          =   131006467
         CurrentDate     =   2958465
         MinDate         =   731
      End
      Begin VB.Frame fBankAcc 
         Height          =   660
         Left            =   90
         TabIndex        =   21
         Top             =   6465
         Visible         =   0   'False
         Width           =   6465
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
            Height          =   300
            Left            =   4515
            Picture         =   "frmRCV.frx":0A64
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   315
            Width           =   330
         End
         Begin VB.TextBox txtChqNo 
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
            Height          =   285
            Left            =   4830
            TabIndex        =   11
            Top             =   345
            Width           =   1575
         End
         Begin VB.Label lblBankAcc 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   30
            TabIndex        =   23
            Top             =   315
            Width           =   4500
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Desposit Slip No."
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
            Index           =   2
            Left            =   4860
            TabIndex        =   22
            Top             =   135
            Width           =   1380
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
         Height          =   660
         Left            =   6495
         TabIndex        =   30
         Top             =   7560
         Width           =   5280
         Begin MSForms.CommandButton cmdSClose 
            Height          =   360
            Left            =   1815
            TabIndex        =   17
            Top             =   210
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
            Left            =   3525
            TabIndex        =   18
            Top             =   210
            Width           =   1665
            Caption         =   "Cancel            "
            PicturePosition =   327683
            Size            =   "2937;635"
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
            Left            =   90
            TabIndex        =   16
            Top             =   210
            Width           =   1665
            Caption         =   "Save & New   "
            PicturePosition =   327683
            Size            =   "2937;635"
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
      Begin MSForms.ComboBox cmbCashInHand 
         Height          =   300
         Left            =   6765
         TabIndex        =   45
         Top             =   6165
         Width           =   2835
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5001;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Lbls 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receive Cash In :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Index           =   0
         Left            =   5280
         TabIndex        =   44
         Top             =   6195
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Received From :"
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
         Index           =   12
         Left            =   105
         TabIndex        =   42
         Top             =   6165
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Amount"
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
         Index           =   11
         Left            =   10395
         TabIndex        =   40
         Top             =   6555
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Other Account : (Debit)"
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
         Index           =   10
         Left            =   9750
         TabIndex        =   39
         Top             =   6585
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes :"
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
         Index           =   7
         Left            =   150
         TabIndex        =   37
         Top             =   7125
         Width           =   585
      End
      Begin VB.Label lblBalCap 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Cash in hand :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   6600
         TabIndex        =   32
         Top             =   6555
         Width           =   1395
      End
      Begin VB.Label lblBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   8010
         TabIndex        =   31
         Top             =   6525
         Width           =   1560
      End
      Begin MSForms.ComboBox CmbAccTitle 
         Height          =   315
         Left            =   1785
         TabIndex        =   2
         Top             =   1545
         Width           =   3810
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "6720;556"
         ListWidth       =   8819
         ListStyle       =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbAccNo 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1545
         Width           =   1665
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "2937;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   10365
         TabIndex        =   29
         Top             =   6180
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Total :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Index           =   6
         Left            =   9705
         TabIndex        =   28
         Top             =   6210
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Voucher # :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Index           =   3
         Left            =   105
         TabIndex        =   27
         Top             =   1020
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Date :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Index           =   5
         Left            =   3240
         TabIndex        =   26
         Top             =   990
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   2535
         TabIndex        =   25
         Top             =   750
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmRCV.frx":0DD6
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
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   11055
      End
   End
End
Attribute VB_Name = "frmRCV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dontShowList As Boolean

Private Sub chkBank_Click()


    If chkBank.Value = vbChecked Then

BankAccCheck:

   If Not AnyBankAccountDefined Then
      If MsgBox("Cannot Charge To Bank, No Bank Accounts Found" & vbNewLine & _
         "Do You Want To Define A Bank Account Now", vbInformation + vbYesNo) = vbYes Then
         With frmNewAccount
            .chkBankAcc = vbChecked
            .Show 1
         End With
         Screen.MousePointer = vbHourglass
         'Sleep (800)
         Screen.MousePointer = vbDefault
         GoTo BankAccCheck
         cmbCashInHand.Enabled = True
       
      Else
         chkBank.Value = vbUnchecked
         cmbCashInHand.ListIndex = -1
        cmbCashInHand.Enabled = False
         Exit Sub
      End If
   End If
End If


Dim BankID As Double
fBankAcc.Visible = chkBank.Value = vbChecked
      'get Default Account
        lblBankAcc = GetDefaultBankAcc(BankID)
        lblBankAcc.Tag = BankID
      'Get ChqNo
      Dim ChqBookNo As Double
      'txtChqNo = GetNextChqNo(GetID(lblBankAcc), ChqBookNo)
      'txtChqNo.Tag = ChqBookNo
      Call ShowBals
      Call DTPicker1_Change
 Exit Sub

BankErr:
   
   MsgBox "No Default Bank Account Found.", vbInformation
   chkBank.Value = vbUnchecked
   
End Sub

Private Sub cmbACCNo_Change()
    If cmbAccNo.MatchFound Then Call cmbACCNo_Click
End Sub

Private Sub cmbACCNo_Click()
    If CmbAccTitle.ListIndex <> cmbAccNo.ListIndex Then CmbAccTitle.ListIndex = cmbAccNo.ListIndex
 
End Sub

Private Sub cmbACCTitle_Change()
    If CmbAccTitle.MatchFound Then Call cmbACCTitle_Click
End Sub

Private Sub cmbACCTitle_Click()
    If cmbAccNo.ListIndex <> CmbAccTitle.ListIndex Then cmbAccNo.ListIndex = CmbAccTitle.ListIndex
    With Tbox(0)
        .Text = "Payment For " & CmbAccTitle
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
End Sub

Private Sub CmbAccTitle_DropButtonClick()
   List1.Visible = False
End Sub
Private Sub CmbAccTitle_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDown Then
   If List1.ListIndex < List1.ListCount - 1 Then List1.ListIndex = List1.ListIndex + 1
   KeyCode = 0
   dontShowList = True
ElseIf KeyCode = vbKeyUp Then
   If List1.ListIndex > 0 Then List1.ListIndex = List1.ListIndex - 1
   KeyCode = 0
   dontShowList = True
ElseIf KeyCode = 13 And List1.Visible Then
   CmbAccTitle = List1
   List1.Visible = False
   CmbAccTitle.SetFocus
End If

End Sub

Private Sub CmbAccTitle_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If dontShowList Then dontShowList = False Else ShowList List1, CmbAccTitle, CmbAccTitle
End Sub

Private Sub cmdBank_Click()
   Dim BankID As Double
   lblBankAcc = SelectBankAcc(BankID)
   lblBankAcc.Tag = BankID
   
'      'Get ChqNo
'      Dim ChqBookNo As Double
'
'      txtChqNo = GetNextChqNo(GetID(lblBankAcc), ChqBookNo)
'      txtChqNo.Tag = ChqBookNo
'
End Sub

Private Sub cmdCancel_Click()
    ClearLView
    Unload Me
End Sub

Private Sub cmdDelete_Click()


With LView.ListItems(LView.SelectedItem.Index)
    cmbAccNo.AddItem .Text
    CmbAccTitle.AddItem .ListSubItems(1).Text
    LView.ListItems.Remove (LView.SelectedItem.Index)
    cmdDelete.Visible = False
    If LView.ListItems.count > 0 Then
        Call GetTotal
    Else
        Tbox(1).Text = ""
        Label1(1).Caption = ""


        
    End If
End With

End Sub


Private Sub cmdNewAccount_Click()
    RecentAccNo = ""
    frmNewAccount.Show 1
    
    If RecentAccNo <> "" Then
        Call FillAccCmbs
        cmbAccNo = RecentAccNo
        CmbAccTitle.SetFocus
    End If

End Sub

Private Sub cmdSave_Click()
    Call AddToLView
End Sub



Private Sub cmdSClear_Click()
    If SaveTrans Then
        MsgBox "Transaction Saved successfuly.", vbInformation
        Dim myDT As Date
        myDT = DTPicker1.Value
        Unload Me
        
        Me.Show
        DTPicker1.Value = myDT
        Call DTPicker1_Change
        Me.SetFocus
    End If
End Sub



Private Sub cmdSClose_Click()
    If SaveTrans Then
        MsgBox "Transaction Saved successfuly.", vbInformation
        Unload Me
    End If
End Sub


Private Sub DTPicker1_Change()

    Call setDateRange(DTPicker1)
    txtVchrNo.Text = getNextVchrNo(DTPicker1.Value, IIf(chkBank = vbUnchecked, "CRV", "BRV"))
    Call ShowBals
    
End Sub

Private Sub DTPicker1_Click()
    Call DTPicker1_Change
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
       
    Me.KeyPreview = True
    DTPicker1.Value = Date
    Call DTPicker1_Click
'    If CurrentUserName = "Administrator" Then
'        DTPicker1.Enabled = True
'    Else
'        DTPicker1.Enabled = False
'    End If
    DTPicker1.Enabled = True
    Call FillAccCmbs
    'chkPrint = IIf(frmMain.mnuPrintOnSave.Checked, vbChecked, vbUnchecked)
End Sub

Private Sub FillAccCmbs()

    Dim AccRec As New ADODB.Recordset
    With AccRec
    Dim ExpAccNo As String
        .CursorLocation = adUseClient
        ExpAccNo = getAccountNo("Direct Expenses")
        '.Open "Select AccNo,AccTitle from Accounts Where Not Parent " & _
              " and (trim(AccNo) like '42%' OR trim(AccNo) like '43%' OR trim(AccNo) like '44%' Or Trim(AccNo) Like '" & ExpAccNo & "%') And Trim(AccNo) Not Like '42-001-%'", con, adOpenForwardOnly, adLockReadOnly
'         .Open "Select AccNo,AccTitle from Accounts Where Parent=0 " & _
'              " and (AccNo like '4%' Or AccNo Like '22%') And AccNo Not Like '41-001-%' ", con, adOpenForwardOnly, adLockReadOnly
            
        .Open "Select AccNo,AccTitle from VActiveAccounts Where Parent=0 AND Active=1 AND AccNo<>'" & CashInHand & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbAccNo.Clear
        CmbAccTitle.Clear
        
        Do Until .EOF
            cmbAccNo.AddItem ![AccNo]
            CmbAccTitle.AddItem ![AccTitle] & ""
            .MoveNEXT
        Loop
        
        .Close
        
    End With
    Set AccRec = Nothing
        
    Call AddToCombo(cmbCashInHand, "AccTitle", "VActiveAccounts", "WHERE SubAccOf='" & CashInHand & "'", , "AccNo")
    If cmbCashInHand.ListCount > 0 Then cmbCashInHand.ListIndex = cmbCashInHand.ListCount - 1
        
'    With cmbOtherAcc
'        .ListHeight = 1000
'        .ClearVals
'        .AddVals con, "AccNo + '{' + Acctitle + '}'", "VActiveAccounts", "AccNo"
'    End With
    
End Sub

Private Function AddToLView() As Boolean

    If cmbAccNo.Text = "" Then
        MsgBox "First select Account # from the list.", vbInformation
        cmbAccNo.SetFocus
        Exit Function
    ElseIf Val(Tbox(1).Text) = 0 Then
        MsgBox "First enter Some Amount.", vbInformation
        Tbox(1).SetFocus
        Exit Function
    End If
    
    Dim LItems As ListItem
    Set LItems = LView.ListItems.add(, cmbAccNo.Text & "", cmbAccNo.Text)
    
    LItems.ListSubItems.add 1, , CmbAccTitle.Text & "" 'UCase(CmbAccTitle.Text & "")
    LItems.ListSubItems.add 2, , (Tbox(0).Text & IIf(Len(Trim(Tbox(2))) = 0, "", " Rcpt # :" & Tbox(2))) 'UCase(Tbox(0).Text & IIf(Len(Trim(Tbox(2))) = 0, "", " Rcpt # :" & Tbox(2)))
      
    LItems.ListSubItems.add 3, , Val(Tbox(1).Text)
    LItems.ListSubItems.add 4, , 0
        
    cmbAccNo.RemoveItem cmbAccNo.ListIndex
    CmbAccTitle.RemoveItem CmbAccTitle.ListIndex
    cmbAccNo.Text = ""
    CmbAccTitle.Text = ""
    
    Tbox(0) = ""
    Tbox(1) = ""
    
    CmbAccTitle.SetFocus
    Call GetTotal
        
    AddToLView = True
    
End Function

Private Sub GetTotal()
Dim StrDescription As String
Dim i As Integer
Dim DebitTotal As Double
Dim CreditTotal As Double
StrDescription = ""
For i = 1 To LView.ListItems.count
    DebitTotal = DebitTotal + Val(LView.ListItems(i).ListSubItems(3).Text)
    StrDescription = StrDescription & IIf((i = LView.ListItems.count And i <> 1), " And " & LView.ListItems(i).SubItems(2), LView.ListItems(i).SubItems(2) & ", ")
    
Next i

Label1(1).Caption = DebitTotal
Tbox(4) = StrDescription


End Sub


Private Function SaveTrans() As Boolean
    
    On Error GoTo err
    SaveTrans = False

    If LView.ListItems.count = 0 Then
        MsgBox "No Transaction to save!", vbInformation
        Exit Function
    ElseIf Val(Label1(1).Caption) = 0 Then
        MsgBox "Invalid Entry.", vbInformation
        Exit Function
    ElseIf Val(Tbox(3)) > 0 Then
        If Not cmbOtherAcc.MatchFound Then
            MsgBox "Please Select Others Account", vbInformation
            cmbOtherAcc.SetFocus
            Exit Function
        End If
    End If
    If chkBank.Value = vbUnchecked Then
        If cmbCashInHand.MatchFound = False Then
            MsgBox "Please Select Cash In Hand Account.", vbInformation
            Exit Function
        End If
    End If
'
'    'Validate ChqNo
'    If chkBank.value = vbChecked Then
'        Dim ChqBookNo As Double
'        Dim Issued As Boolean
'        If Not ValidChqNo(GetID(lblBankAcc), txtChqNo, Issued, ChqBookNo) Then
'            If Issued Then
'                MsgBox "The Cheque With Number '" & txtChqNo & "' Has Been Issued.", vbInformation
'            Else
'                MsgBox "Invalid Cheque No", vbInformation
'            End If
'            txtChqNo.SetFocus
'            Exit Function
'        End If
'        txtChqNo.Tag = ChqBookNo
'    End If
'
    ''''''''''''''''''''''Check For Voucher No. Used Or Not '''''''''''''''''
    If VoucherExists(txtVchrNo.Text) Then
        Dim VNo As String
        VNo = getNextVchrNo(DTPicker1.Value, IIf(cmbPymtmode.ListIndex = 0, "CRV", "BRV"))
        MsgBox "Voucher No. : " & Label1(4) & " Has Been Used By Other Transaction." & vbNewLine & _
           "Voucher No. : " & VNo & " Will Be Used For This Transaction.", vbInformation
           txtVchrNo.Text = VNo
    End If
    
'    Dim VNo As String
'    VNo = getNextVchrNo(DTPicker1.value, IIf(chkBank = vbUnchecked, "CRV", "BRV"))
'    If Label1(4) <> VNo Then
'        MsgBox "Voucher No. : " & Label1(4) & " Has Been Used By Other Transaction." & vbNewLine & _
'         "Voucher No. : " & VNo & " Will Be Used For This Transaction.", vbInformation
'         Label1(4) = VNo
'    End If

    Dim i As Integer
    Dim INextSno As Double
    Dim Bal As Double
    
    Dim theDesc As String
    theDesc = "Payment For "
    
    
    Dim DebitAcc As String
      Dim SlipNo As String
      
      
    INextSno = getNextSno(DTPicker1.Value)
    
    SlipNo = 0 'INextSno
    Call StartTrans(con)
    
    For i = 1 To LView.ListItems.count
    
        Bal = getBalance(LView.ListItems(i).Text, DTPicker1.Value)
        Bal = Bal - Val(LView.ListItems(i).ListSubItems(3).Text) '- Val(LView.ListItems(i).ListSubItems(4).Text)
        
        con.Execute "INSERT INTO Vouchers Values(" & INextSno & ",'" & Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & _
         Label1(4) & "','" & LView.ListItems(i).Text & "','" & LView.ListItems(i).ListSubItems(2).Text & "',0," & _
         Val(LView.ListItems(i).ListSubItems(3).Text) & "," & _
         Bal & ",''," & SlipNo & ")"
         
         Call adjustLedger(LView.ListItems(i).Text, DTPicker1.Value, INextSno, -Val(LView.ListItems(i).ListSubItems(3).Text))
         
         INextSno = INextSno + 1
         
         If InStr(LView.ListItems(i).SubItems(2), "Rcpt # :") > 0 Then
            'Seperate Teh Rcpt No.
            RcptNo = " (" & Right(LView.ListItems(i).SubItems(2), Len(LView.ListItems(i).SubItems(2)) - InStr(LView.ListItems(i).SubItems(2), "Rcpt # :") + 2) & ")"
         Else
            RcptNo = ""
         End If
         'TheDesc = TheDesc & IIf((i = LView.ListItems.Count And i <> 1), " And " & LView.ListItems(i).SubItems(2), LView.ListItems(i).SubItems(2) & ", ")
    Next i
    
    'Generate The Desc
    theDesc = Tbox(4)

    
    
    Dim NetAmount As Double
    
    NetAmount = Val(Label1(1)) - Val(Tbox(3))

   '  Check Whether Received in Bank Or in Cash in Hand
   
    If chkBank.Value = vbChecked Then
        INextSno = getNextSno(DTPicker1.Value, "Cheque")
        DebitAcc = GetID(lblBankAcc)
      
        con.Execute "Insert Into Cheque ( SNo, CDate, AccNo, Description, Amount, BankID, chequeno, ChqBookNo, chequeType, chequeDate, Posted, Payment, Bounced,BankAccNo) Values(" & INextSno & ",'" & Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & _
         DebitAcc & "','" & Tbox(0) & "'," & NetAmount & "," & lblBankAcc.Tag & ",'" & _
         txtChqNo & "'," & Val(txtChqNo.Tag) & ",'Open','" & DTPicker1 & "',1,0,0,'" & GetID(lblBankAcc) & "')"
       
        'Call IssueChq(Val(txtChqNo.Tag), txtChqNo)
        INextSno = getNextSno(DTPicker1.Value)
    Else
        DebitAcc = cmbCashInHand.List(cmbCashInHand.ListIndex, 1) 'CashInHand
    End If
    
    ''''''''''''''''Charge To  Cash Entry'''''''''''''
    
    Bal = getBalance(DebitAcc, DTPicker1.Value)
    Bal = Bal + Val(Label1(1))
    
    
    'Debit Entry
    con.Execute "INSERT INTO Vouchers VALUES(" & INextSno & ",'" & _
        Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & Label1(4) & "','" & DebitAcc & _
        "','" & theDesc & "'," & NetAmount & ",0," & Bal & ",''," & SlipNo & ")"
    
    Call adjustLedger(DebitAcc, DTPicker1.Value, INextSno, NetAmount)
    
    
    'Other Debit Entry
    
    If Val(Tbox(3)) > 0 Then
    
        Bal = getBalance(cmbOtherAcc.ID, DTPicker1.Value)
        Bal = Bal + Val(Tbox(3))
        
        con.Execute "Insert into Vouchers Values(" & INextSno & ",'" & _
            Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & Label1(4) & "','" & cmbOtherAcc.ID & _
            "','" & theDesc & "'," & Val(Tbox(3)) & ",0," & Bal & ",''," & SlipNo & ")"
        
        Call adjustLedger(cmbOtherAcc.ID, DTPicker1.Value, INextSno, NetAmount)
    End If
    
    'NOTE: Payee is used for Received From Field for Reciept Vouchers.
    con.Execute "INSERT INTO VoucherInfo(VchrNo,UserName,MachineName,DT,Notes,Payee) VALUES('" & Label1(4) & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "','" & Tbox(4) & "','" & txtRcvdFrom & "')"
    
    con.CommitTrans
    
    'Print If To Be Printed
    If chkPrint = vbChecked Then Call PrintVoucher(MainForm.Cr1, Label1(4), True, Not bVoucherViewOnlyWhenSaved)
    
    Call PlaySaveSound
    SaveTrans = True
    Exit Function
    
err:
    MsgBox err.Description
     con.RollbackTrans
     
End Function

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    If LView.ListItems.count > 0 Then
        If MsgBox("There is a Transaction active." & vbNewLine & "Save or Cancel Transaction first.", vbInformation + vbOKOnly) Then
            cancel = True
        End If
    End If
End If
End Sub

Private Sub lblBankAcc_Change()
   Call ShowBals
End Sub

Private Sub List1_Click()
If UCase(Screen.ActiveControl.Name) <> "CMBACCTITLE" Then
   CmbAccTitle = List1
   List1.Visible = False
   CmbAccTitle.SetFocus
End If
End Sub
Private Sub LView_Click()
    If LView.ListItems.count > 0 Then
        With cmdDelete
            .Move .Left, LView.Top + (LView.SelectedItem.Top)
            .Visible = True
        End With
    End If
End Sub





Private Sub LView_DblClick()
    If LView.ListItems.count = 0 Then Exit Sub
    
        With LView.SelectedItem
            
                txtEdit.Width = LView.ColumnHeaders(4).Width + 30
                txtEdit.Move LView.ColumnHeaders(4).Left + LView.Left + 30, LView.Top + .Top
                txtEdit.Text = Val(.ListSubItems(3).Text)
                txtEdit.Tag = 3
            
        End With
        LView.Tag = LView.SelectedItem.Index
        txtEdit.Visible = True
        txtEdit.SelStart = 0
        txtEdit.SelLength = Len(txtEdit)
        txtEdit.SetFocus
    
End Sub

Private Sub LView_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo err
If LView.ListItems.count > 0 Then
    If KeyCode = 46 Then
        Call cmdDelete_Click
    End If
End If
Exit Sub
err:
    
    MsgBox err.Description
    

End Sub

Private Sub LView_Validate(cancel As Boolean)
On Error Resume Next
    cmdDelete.Visible = False
    txtEdit.Visible = False
End Sub

Private Sub TBox_GotFocus(Index As Integer)
'    If Index = 1 Then
'        Tbox(1).SelStart = 0
'        Tbox(1).SelLength = Len(Tbox(1).Text)
'    End If

With Tbox(Index)
    .SelStart = 0
    .SelLength = Len(Tbox(Index))
End With

End Sub

Private Sub Tbox_KeyPress(Index As Integer, KeyAscii As Integer)

If Index = 1 Or Index = 3 Then
    KeyAscii = OnlyNumber(KeyAscii, True)
End If

End Sub
Private Sub ClearLView()

Dim iTotal As Integer
iTotal = LView.ListItems.count
    For i = 1 To iTotal
        LView.ListItems.Remove (1)
    Next i
Tbox(0).Text = ""
Tbox(1).Text = ""
cmdDelete.Visible = False
Call DTPicker1_Click
End Sub


Private Sub TBox_LostFocus(Index As Integer)
    On Error Resume Next
    'Tbox(Index) = UCase(Tbox(Index))
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

With LView.ListItems(Val(LView.Tag))
    If KeyAscii = 13 Then
        .ListSubItems(Val(txtEdit.Tag)).Text = txtEdit
        Call GetTotal
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    Else
        KeyAscii = OnlyNumber(KeyAscii, True)
    End If
End With

End Sub

Private Sub Form_Resize()
    FALL.Move (Me.ScaleWidth - FALL.Width) / 2, (Me.ScaleHeight - FALL.Height) / 2
End Sub

Private Sub cmbCashInHand_Change()
    Call cmbCashInHand_Click
End Sub

Private Sub cmbCashInHand_Click()
    lblBal = Format(getBalance(cmbCashInHand.List(cmbCashInHand.ListIndex, 1), DTPicker1), "0.00")
End Sub


Private Sub ShowBals()

    If chkBank.Value = vbChecked Then
        lblBalCap = "Bank Balance :"
        If lblBankAcc.Caption = "" Then Exit Sub
        lblBal = Format(getBalance(GetID(lblBankAcc), DTPicker1), "0.00")
    Else
        lblBalCap = "Cash In Hand :"
        'lblBal = Format(getBalance(CashInHand, DTPicker1), "0.00")
        If cmbCashInHand.MatchFound Then
            lblBal = Format(getBalance(cmbCashInHand.List(cmbCashInHand.ListIndex, 1), DTPicker1), "0.00")
        End If
    End If
    
End Sub

Private Sub txtVchrNo_Change()
    Label1(4).Caption = txtVchrNo.Text
End Sub
