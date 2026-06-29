VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExpensesOld 
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
      TabIndex        =   15
      Top             =   -30
      Width           =   11835
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
         TabIndex        =   30
         Top             =   60
         Width           =   11850
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Expense Voucher"
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
            TabIndex        =   31
            Top             =   165
            Width           =   11655
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Expense Voucher"
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
            TabIndex        =   32
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
         Left            =   165
         TabIndex        =   29
         Top             =   7965
         Value           =   1  'Checked
         Width           =   1830
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1590
         Left            =   1785
         TabIndex        =   28
         Top             =   1875
         Visible         =   0   'False
         Width           =   4320
      End
      Begin VB.CommandButton cmdDelete 
         CausesValidation=   0   'False
         Height          =   285
         Left            =   11175
         Picture         =   "frmExpensesOld.frx":0000
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
         Picture         =   "frmExpensesOld.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Add This Entry To Voucher"
         Top             =   1320
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
         TabIndex        =   16
         Top             =   2115
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.CheckBox chkBank 
         Caption         =   "Charge To Bank"
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
         Left            =   150
         TabIndex        =   9
         Top             =   7050
         Width           =   1635
      End
      Begin MSComctlLib.ListView LView 
         Height          =   4995
         Left            =   105
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1905
         Width           =   11670
         _ExtentX        =   20585
         _ExtentY        =   8811
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
         Format          =   98959363
         CurrentDate     =   2958465
         MinDate         =   731
      End
      Begin VB.Frame fBankAcc 
         Height          =   660
         Left            =   120
         TabIndex        =   17
         Top             =   6930
         Visible         =   0   'False
         Width           =   7695
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
            Left            =   5715
            Picture         =   "frmExpensesOld.frx":0A64
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
            Left            =   6075
            TabIndex        =   11
            Top             =   330
            Width           =   1575
         End
         Begin VB.Label lblBankAcc 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   45
            TabIndex        =   19
            Top             =   315
            Width           =   5640
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            Caption         =   "Cheque #"
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
            Left            =   6105
            TabIndex        =   18
            Top             =   120
            Width           =   1170
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
         Height          =   660
         Left            =   6495
         TabIndex        =   26
         Top             =   7560
         Width           =   5280
         Begin MSForms.CommandButton cmdSClose 
            Height          =   360
            Left            =   1815
            TabIndex        =   13
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
            TabIndex        =   14
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
            TabIndex        =   12
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
      Begin VB.Label lblBalCap 
         Caption         =   "Balance"
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
         Left            =   9600
         TabIndex        =   35
         Top             =   705
         Width           =   1515
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Cash In Hand"
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
         Left            =   6870
         TabIndex        =   34
         Top             =   720
         Width           =   1305
      End
      Begin MSForms.ComboBox cmbCashInHand 
         Height          =   285
         Left            =   6840
         TabIndex        =   33
         Top             =   960
         Width           =   2745
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4842;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
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
         Left            =   9615
         TabIndex        =   27
         Top             =   960
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
         Left            =   10350
         TabIndex        =   25
         Top             =   6930
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
         Left            =   9690
         TabIndex        =   24
         Top             =   6960
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         Left            =   1290
         TabIndex        =   21
         Top             =   960
         Width           =   1560
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmExpensesOld.frx":0DD6
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
         TabIndex        =   20
         Top             =   1320
         Width           =   11055
      End
   End
End
Attribute VB_Name = "frmExpensesOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dontShowList As Boolean

Private Sub chkBank_Click()


    If chkBank.Value = vbChecked Then

BankAccCheck:
    cmbCashInHand.Enabled = False
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
         cmbCashInHand.Enabled = False
       
      Else
         chkBank.Value = vbUnchecked
         Exit Sub
      End If
   End If
   Else
    cmbCashInHand.Enabled = True
End If


Dim BankID As Double
fBankAcc.Visible = chkBank.Value = vbChecked
      'get Default Account
        lblBankAcc = GetDefaultBankAcc(BankID)
        lblBankAcc.Tag = BankID
      'Get ChqNo
      Dim ChqBookNo As Double
      txtChqNo = GetNextChqNo(GetID(lblBankAcc), ChqBookNo)
      txtChqNo.Tag = ChqBookNo
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
    With TBox(0)
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

Private Sub cmbCashInHand_Change()
    Call cmbCashInHand_Click
End Sub

Private Sub cmbCashInHand_Click()
    
    If chkBank.Value = vbChecked Then
        lblBalCap = "Bank Balance :"
        If lblBankAcc.Caption = "" Then Exit Sub
        lblBal = Format(getBalance(GetID(lblBankAcc), DTPicker1), "0.00")
    Else
        lblBalCap = "Balance :"
'        If cmbCashInHand.MatchFound = False Then
'            MsgBox "Please Select Cash In Hand Account", vbInformation
'            Exit Sub
'        End If
        If cmbCashInHand.MatchFound = False Then Exit Sub
        lblBal = Format(getBalance(cmbCashInHand.List(cmbCashInHand.ListIndex, 1), DTPicker1), "0.00")
    End If
End Sub

Private Sub cmdBank_Click()
   Dim BankID As Double
   lblBankAcc = SelectBankAcc(BankID)
   lblBankAcc.Tag = BankID
   
      'Get ChqNo
      Dim ChqBookNo As Double
   
      txtChqNo = GetNextChqNo(GetID(lblBankAcc), ChqBookNo)
      txtChqNo.Tag = ChqBookNo
        
End Sub

Private Sub cmdCancel_Click()
    ClearLView
    Unload Me
End Sub

Private Sub CmdDelete_Click()


With LView.ListItems(LView.SelectedItem.Index)
    cmbAccNo.AddItem .Text
    CmbAccTitle.AddItem .ListSubItems(1).Text
    LView.ListItems.Remove (LView.SelectedItem.Index)
    cmdDelete.Visible = False
    If LView.ListItems.count > 0 Then
        Call GetTotal
    Else
        TBox(1).Text = ""
        Label1(1).Caption = ""


        
    End If
End With

End Sub


Private Sub cmdSave_Click()
    Call AddToLView
End Sub



Private Sub cmdSClear_Click()
    If SaveTrans Then
        MsgBox "Transaction Saved successfuly.", vbInformation
        Dim myDT As Date, iCashAccountIndex As String
        myDT = DTPicker1.Value
        iCashAccountIndex = cmbCashInHand.ListIndex
        Unload Me
        Me.Show
        DTPicker1.Value = myDT
        cmbCashInHand.ListIndex = iCashAccountIndex
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
    Label1(4).Caption = getNextVchrNo(DTPicker1.Value, IIf(chkBank = vbUnchecked, "CPV", "BPV"))
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
    Call FillAccCmbs
    chkPrint = Abs(bPrintWhenSaved)
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
            
        .Open "Select AccNo,AccTitle from Accounts Where Parent=0 AND Active=1 AND AccNo<>'" & CashInHand & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbAccNo.Clear
        CmbAccTitle.Clear
        
        Do Until .EOF
            cmbAccNo.AddItem ![AccNo]
            CmbAccTitle.AddItem ![AccTitle] & ""
            .MoveNEXT
        Loop
        
        .Close
                
        .Open "SELECT AccNo,AccTitle FROM VActiveAccounts WHERE SubAccOf='" & CashInHand & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCashInHand.Clear
        Do Until .EOF
            cmbCashInHand.AddItem !AccTitle & ""
            cmbCashInHand.List(.AbsolutePosition - 1, 1) = !AccNo & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set AccRec = Nothing
    If cmbCashInHand.ListCount > 0 Then cmbCashInHand.ListIndex = i
    
    
    Set AccRec = Nothing
End Sub

Private Function AddToLView() As Boolean

    If cmbAccNo.Text = "" Then
        MsgBox "First select Account # from the list.", vbInformation
        cmbAccNo.SetFocus
        Exit Function
    ElseIf Val(TBox(1).Text) = 0 Then
        MsgBox "First enter Some Amount.", vbInformation
        TBox(1).SetFocus
        Exit Function
    End If
    
    Dim LItems As ListItem
    Set LItems = LView.ListItems.add(, cmbAccNo.Text & "", cmbAccNo.Text)
    
    LItems.ListSubItems.add 1, , CmbAccTitle.Text & ""
    LItems.ListSubItems.add 2, , TBox(0).Text & IIf(Len(Trim(TBox(2))) = 0, "", " Rcpt # :" & TBox(2))
      
    LItems.ListSubItems.add 3, , Val(TBox(1).Text)
    LItems.ListSubItems.add 4, , 0
        
    cmbAccNo.RemoveItem cmbAccNo.ListIndex
    CmbAccTitle.RemoveItem CmbAccTitle.ListIndex
    cmbAccNo.Text = ""
    CmbAccTitle.Text = ""
    CmbAccTitle.SetFocus
    Call GetTotal
        
    AddToLView = True
    
End Function

Private Sub GetTotal()

Dim i As Integer
Dim DebitTotal As Double
Dim CreditTotal As Double

For i = 1 To LView.ListItems.count
    DebitTotal = DebitTotal + Val(LView.ListItems(i).ListSubItems(3).Text)
Next i

Label1(1).Caption = DebitTotal

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
    ElseIf Val(Label1(1)) > Val(lblBal) Then
        If chkBank.Value = vbChecked Then
            If Not AllowOD(GetID(lblBankAcc)) Then
                MsgBox "Not Much Amount In " & Left(lblBalCap, Len(lblBalCap) - 2) & ".", vbInformation
                Exit Function
            End If
        Else
            If strCompany <> "Sunlike" Then
                MsgBox "Not Much Amount In " & Left(lblBalCap, Len(lblBalCap) - 2) & ".", vbInformation
                Exit Function
            End If
        End If
    End If
    
    'Validate ChqNo
    If chkBank.Value = vbChecked Then
        Dim ChqBookNo As Double
        Dim Issued As Boolean
        If Not ValidChqNo(GetID(lblBankAcc), txtChqNo, Issued, ChqBookNo) Then
            If Issued Then
                MsgBox "The Cheque With Number '" & txtChqNo & "' Has Been Issued.", vbInformation
            Else
                MsgBox "Invalid Cheque No", vbInformation
            End If
            txtChqNo.SetFocus
            Exit Function
        End If
        txtChqNo.Tag = ChqBookNo
    Else
        If cmbCashInHand.MatchFound = False Then
            MsgBox "Please Select Cash In Hand Account.", vbInformation
            Exit Function
        End If
    End If
 
    ''''''''''''''''''''''Check For Voucher No. Used Or Not '''''''''''''''''

    Dim VNo As String
    VNo = getNextVchrNo(DTPicker1.Value, IIf(chkBank = vbUnchecked, "CPV", "BPV"))
    If Label1(4) <> VNo Then
        MsgBox "Voucher No. : " & Label1(4) & " Has Been Used By Other Transaction." & vbNewLine & _
         "Voucher No. : " & VNo & " Will Be Used For This Transaction.", vbInformation
         Label1(4) = VNo
    End If

    Dim i As Integer
    Dim INextSno As Double
    Dim Bal As Double
    
    Dim theDesc As String
    theDesc = "Payment For "
    
    
    Dim CreditAcc As String
      Dim SlipNo As String
      
      
    INextSno = getNextSno(DTPicker1.Value)
    
    Dim bWaitForAuth As Boolean
    If chkBank.Value = vbChecked Then
        bWaitForAuth = VoucherRequiresAuthorization("BPV")
    Else
        bWaitForAuth = VoucherRequiresAuthorization("CPV")
    End If
    
    SlipNo = INextSno
    con.BeginTrans
    
    For i = 1 To LView.ListItems.count
        If bWaitForAuth = False Then
            Bal = getBalance(LView.ListItems(i).Text, DTPicker1.Value)
            Bal = Bal + Val(LView.ListItems(i).ListSubItems(3).Text) - Val(LView.ListItems(i).ListSubItems(4).Text)
        
            con.Execute "insert into Vouchers Values(" & INextSno & ",'" & Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & _
             Label1(4) & "','" & LView.ListItems(i).Text & "','" & LView.ListItems(i).ListSubItems(2).Text & "'," & _
             Val(LView.ListItems(i).ListSubItems(3).Text) & ",0," & _
             Bal & ",''," & SlipNo & ")"
         
            Call adjustLedger(LView.ListItems(i).Text, DTPicker1.Value, INextSno, Val(LView.ListItems(i).ListSubItems(3).Text))
         
            INextSno = INextSno + 1
        Else
            con.Execute "INSERT INTO VouchersForAuthorization(VDate,VchrNo,Accno,Description,Debit,Credit,balance,DpstSlip,CSNo) VALUES('" & Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & _
             Label1(4) & "','" & LView.ListItems(i).Text & "','" & LView.ListItems(i).ListSubItems(2).Text & "'," & _
             Val(LView.ListItems(i).ListSubItems(3).Text) & ",0,0,0," & SlipNo & ")"
        End If
        
        If InStr(LView.ListItems(i).SubItems(2), "Rcpt # :") > 0 Then
            'Seperate Teh Rcpt No.
            RcptNo = " (" & Right(LView.ListItems(i).SubItems(2), Len(LView.ListItems(i).SubItems(2)) - InStr(LView.ListItems(i).SubItems(2), "Rcpt # :") + 2) & ")"
        Else
            RcptNo = ""
        End If
        theDesc = theDesc & IIf((i = LView.ListItems.count And i <> 1), " And " & LView.ListItems(i).SubItems(2), LView.ListItems(i).SubItems(2) & ", ")
    Next i
    
    'Generate The Desc
    theDesc = theDesc

   '  Check Whether Expense is Charged To Bank Or To Cash in Hand
   
    If chkBank.Value = vbChecked Then
        INextSno = getNextSno(DTPicker1.Value, "Cheque")
        CreditAcc = GetID(lblBankAcc)
      
        con.Execute "INSERT INTO Cheque(SNo,CDate,AccNo,Description,Amount,BankID,chequeno,ChqBookNo,chequeType,chequeDate,Posted,Payment,Bounced) Values(" & INextSno & ",'" & Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & _
         CreditAcc & "','" & TBox(0) & "'," & Label1(1).Caption & "," & lblBankAcc.Tag & ",'" & _
         txtChqNo & "'," & txtChqNo.Tag & ",'Open','" & DTPicker1 & "',1,1,0)"
       
        Call IssueChq(Val(txtChqNo.Tag), txtChqNo)
        INextSno = getNextSno(DTPicker1.Value)
    Else
        CreditAcc = cmbCashInHand.List(cmbCashInHand.ListIndex, 1) 'CashInHand
    End If
    
    ''''''''''''''''Charge To  Cash Entry'''''''''''''
    If bWaitForAuth = False Then
        Bal = getBalance(CreditAcc, DTPicker1.Value)
        Bal = Bal - Val(Label1(1))
        
        
        'Credit Entry
        con.Execute "Insert into Vouchers Values(" & INextSno & ",'" & _
            Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & Label1(4) & "','" & CreditAcc & _
            "','" & theDesc & "',0," & Label1(1) & "," & Bal & ",''," & SlipNo & ")"
        
        Call adjustLedger(CreditAcc, DTPicker1.Value, INextSno, -Val(Label1(1)))
        
        con.Execute "Insert Into VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & Label1(4) & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
    Else
        con.Execute "INSERT INTO VouchersForAuthorization(VDate,VchrNo,Accno,Description,Debit,Credit,balance,DpstSlip,CSNo) VALUES('" & Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & _
             Label1(4) & "','" & CreditAcc & "','" & theDesc & "',0," & Label1(1) & ",0,''," & SlipNo & ")"
             
        con.Execute "INSERT INTO VouchersForAuthorizationInfo(VchrNo,UserName,MachineName,DT,PostedThroughJVForm) VALUES('" & Label1(4) & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "',0)"
    End If
    con.CommitTrans
    
    'Print If To Be Printed
    If chkPrint = vbChecked Then Call PrintVoucher(MainForm.cr1, Label1(4), True, , bWaitForAuth)
    
    Call PlaySaveSound
    SaveTrans = True
    Exit Function
err:
    MsgBox err.Description
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
        Call CmdDelete_Click
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
    If Index = 1 Then
        TBox(1).SelStart = 0
        TBox(1).SelLength = Len(TBox(1).Text)
    End If
End Sub

Private Sub Tbox_KeyPress(Index As Integer, KeyAscii As Integer)

If Index = 1 Then
    KeyAscii = OnlyNumber(KeyAscii, True)
End If

End Sub
Private Sub ClearLView()

Dim iTotal As Integer
iTotal = LView.ListItems.count
    For i = 1 To iTotal
        LView.ListItems.Remove (1)
    Next i
TBox(0).Text = ""
TBox(1).Text = ""
cmdDelete.Visible = False
Call DTPicker1_Click
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
    FAll.Move (Me.ScaleWidth - FAll.Width) / 2, (Me.ScaleHeight - FAll.Height) / 2
End Sub



Private Sub ShowBals()
    
    Call cmbCashInHand_Click
'    If chkBank.Value = vbChecked Then
'        'lblBalCap = "Bank Balance :"
'        If lblBankAcc.Caption = "" Then Exit Sub
'        lblBal = Format(getBalance(GetID(lblBankAcc), DTPicker1), "0.00")
'    Else
'        'lblBalCap = "Cash In Hand :"
'        lblBal = Format(getBalance(CashInHand, DTPicker1), "0.00")
'    End If
    
End Sub
