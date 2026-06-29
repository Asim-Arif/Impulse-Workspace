VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEmpAdv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Advances"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11040
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
   ScaleHeight     =   6930
   ScaleWidth      =   11040
   WindowState     =   2  'Maximized
   Begin VB.Frame FALL 
      Height          =   6795
      Left            =   75
      TabIndex        =   13
      Top             =   60
      Width           =   10920
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
         Left            =   180
         TabIndex        =   32
         Top             =   6450
         Value           =   1  'Checked
         Width           =   1830
      End
      Begin VB.Frame Frame1 
         Height          =   645
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   10920
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Advances"
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
            Index           =   7
            Left            =   60
            TabIndex        =   28
            Top             =   165
            Width           =   10605
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Advances"
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
            Index           =   8
            Left            =   75
            TabIndex        =   29
            Top             =   180
            Width           =   10590
         End
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1590
         Left            =   1845
         TabIndex        =   26
         Top             =   1755
         Visible         =   0   'False
         Width           =   4140
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E3EDF9&
         CausesValidation=   0   'False
         Height          =   285
         Left            =   5505
         TabIndex        =   25
         Top             =   2970
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.CommandButton cmdDelete 
         CausesValidation=   0   'False
         Height          =   285
         Left            =   10275
         Picture         =   "frmEmpAdv.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Remove This Entry From Voucher"
         Top             =   1815
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
         Left            =   10275
         Picture         =   "frmEmpAdv.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Add This Entry To Voucher"
         Top             =   1215
         UseMaskColor    =   -1  'True
         Width           =   525
      End
      Begin VB.TextBox Tbox 
         Height          =   315
         Index           =   0
         Left            =   4455
         MaxLength       =   255
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1440
         Width           =   4245
      End
      Begin VB.TextBox Tbox 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   8730
         TabIndex        =   4
         Top             =   1440
         Width           =   1500
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
         Left            =   195
         TabIndex        =   8
         Top             =   5550
         Width           =   1635
      End
      Begin MSComctlLib.ListView LView 
         Height          =   3720
         Left            =   150
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1770
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   6562
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
         Left            =   4185
         TabIndex        =   0
         Top             =   870
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
         Format          =   22609923
         CurrentDate     =   37250
      End
      Begin VB.Frame fBankAcc 
         Height          =   675
         Left            =   150
         TabIndex        =   14
         Top             =   5415
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
            Left            =   5595
            Picture         =   "frmEmpAdv.frx":0A64
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   330
            Width           =   315
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
            Left            =   5925
            TabIndex        =   10
            Top             =   345
            Width           =   1710
         End
         Begin VB.Label lblBankAcc 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   45
            TabIndex        =   16
            Top             =   345
            Width           =   5535
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
            Left            =   5955
            TabIndex        =   15
            Top             =   135
            Width           =   1170
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
         Height          =   660
         Left            =   4980
         TabIndex        =   30
         Top             =   6060
         Width           =   5280
         Begin MSForms.CommandButton cmdCancel 
            Height          =   360
            Left            =   3495
            TabIndex        =   31
            Top             =   195
            Width           =   1665
            Caption         =   "Cancel            "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmEmpAdv.frx":0DD6
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
            TabIndex        =   11
            Top             =   195
            Width           =   1665
            Caption         =   "Save & New   "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmEmpAdv.frx":0EE8
            Accelerator     =   78
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdSClose 
            Height          =   360
            Left            =   1785
            TabIndex        =   12
            Top             =   195
            Width           =   1665
            Caption         =   "Save & Close  "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmEmpAdv.frx":0FFA
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
      Begin VB.Label lblBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   8655
         TabIndex        =   24
         Top             =   870
         Width           =   1560
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
         Left            =   7245
         TabIndex        =   23
         Top             =   915
         Width           =   1395
      End
      Begin MSForms.ComboBox CmbAccTitle 
         Height          =   315
         Left            =   1845
         TabIndex        =   2
         Top             =   1440
         Width           =   2610
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "4604;556"
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
         Left            =   180
         TabIndex        =   1
         Top             =   1440
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
         Left            =   8865
         TabIndex        =   22
         Top             =   5700
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
         Left            =   8190
         TabIndex        =   21
         Top             =   5730
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
         Left            =   180
         TabIndex        =   20
         Top             =   900
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
         Left            =   3480
         TabIndex        =   19
         Top             =   900
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   1350
         TabIndex        =   18
         Top             =   870
         Width           =   1560
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmEmpAdv.frx":1158
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
         Left            =   180
         TabIndex        =   17
         Top             =   1215
         Width           =   10065
      End
   End
End
Attribute VB_Name = "frmEmpAdv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim DontShowList As Boolean
Dim Loaded As Boolean
Public LongTermAdvances As Boolean

Private Sub cmbAccNo_Validate(Cancel As Boolean)
    If Not cmbAccNo.MatchFound Then cmbAccNo.ListIndex = CmbAccTitle.ListIndex
End Sub

Private Sub CmbAccTitle_LostFocus()
List1.Visible = False
End Sub

Private Sub CmbAccTitle_Validate(Cancel As Boolean)
   If Not CmbAccTitle.MatchFound Then CmbAccTitle.ListIndex = cmbAccNo.ListIndex
End Sub

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
         Sleep (800)
         Screen.MousePointer = vbDefault
         GoTo BankAccCheck
         
         
      Else
         chkBank.Value = vbUnchecked
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
      txtChqNo = GetNextChqNo(GetID(lblBankAcc), ChqBookNo)
      txtChqNo.Tag = ChqBookNo
            Call ShowBals
 Exit Sub

BankErr:
   
   MsgBox "No Default Bank Account Found.", vbInformation
   chkBank.Value = vbUnchecked
   
End Sub

Private Sub cmbAccNo_Change()
    If cmbAccNo.MatchFound Then Call cmbAccNo_Click
End Sub

Private Sub cmbAccNo_Click()
    If CmbAccTitle.ListIndex <> cmbAccNo.ListIndex Then CmbAccTitle.ListIndex = cmbAccNo.ListIndex
End Sub

Private Sub CmbAccTitle_Change()
    If CmbAccTitle.MatchFound Then Call cmbAccTitle_Click
End Sub

Private Sub cmbAccTitle_Click()
    If cmbAccNo.ListIndex <> CmbAccTitle.ListIndex Then cmbAccNo.ListIndex = CmbAccTitle.ListIndex
    Tbox(0) = "Advance Payment To " & CmbAccTitle
End Sub




Private Sub cmbBankBranch_Change()
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
    If LView.ListItems.Count > 0 Then
        Call GetTotal
    Else
        Tbox(1).Text = ""
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
        Unload Me
        Me.Show
        Me.SetFocus
    End If
End Sub

Private Sub cmdSClose_Click()
    If SaveTrans Then
       
        MsgBox "Transaction Saved successfuly.", vbInformation
        Unload Me
    End If
End Sub

Private Sub ComboBox1_Change()

End Sub

Private Sub DTPicker1_Change()
       Call setDateRange(DTPicker1)
    Label1(4).Caption = getNextVchrNo(DTPicker1.Value, "PV")
        Call ShowBals
End Sub

Private Sub DTPicker1_Click()
    Label1(4).Caption = getNextVchrNo(DTPicker1.Value, "PV")
        Call ShowBals
End Sub

Public Sub Form_Activate()
If Not Loaded Then

      If LongTermAdvances Then
         If Not UserHasAccess("LongTermAdvance") Then GoTo NoPermission
      Else
         If Not UserHasAccess("ShortTermAdvance") Then GoTo NoPermission
      End If
         
         
         
      
      
      
      
   Call FillAccCmbs
   Loaded = True
   
End If

Exit Sub

NoPermission:
   Me.Hide
   MsgBox "You Donot Have Permissions To Use This Option. Please Contact Your Administrator", vbInformation
   Unload Me
   
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
   
    Me.KeyPreview = True
    DTPicker1.Value = Date
    Loaded = False
    Call DTPicker1_Click
chkPrint = IIf(frmMain.mnuPrintOnSave.Checked, vbChecked, vbUnchecked)
End Sub

Private Sub FillAccCmbs()

    Dim AccRec As New ADODB.Recordset
    With AccRec
        .CursorLocation = adUseClient
        
        If LongTermAdvances Then
            Me.Caption = "Long Term Advances"
            
            .Open "Select LAdvAccno as AccountNo,EmpName From vEmps Where Active=1 Order By Accno", con, adOpenForwardOnly, adLockReadOnly
        Else
            Me.Caption = "Short Term Advances"
            
            .Open "Select SAdvAccno as AccountNo,EmpName From vEmps Where Active=1 Order By Accno", con, adOpenForwardOnly, adLockReadOnly
        End If
        
        cmbAccNo.Clear
        CmbAccTitle.Clear
        
        Do Until .EOF
            cmbAccNo.AddItem ![AccountNo] & ""
            CmbAccTitle.AddItem ![EmpName] & ""
            .MoveNext
        Loop
        
        .Close
    End With
    Set AccRec = Nothing
    Label1(7) = Me.Caption
    Label1(8) = Me.Caption
    
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
    Set LItems = LView.ListItems.Add(, cmbAccNo.Text & "", cmbAccNo.Text)
    
    LItems.ListSubItems.Add 1, , CmbAccTitle.Text & ""
    LItems.ListSubItems.Add 2, , Tbox(0).Text & ""
      
        LItems.ListSubItems.Add 3, , Val(Tbox(1).Text)
        LItems.ListSubItems.Add 4, , 0
        
    cmbAccNo.RemoveItem cmbAccNo.ListIndex
    CmbAccTitle.RemoveItem CmbAccTitle.ListIndex
    cmbAccNo.Text = ""
    CmbAccTitle.Text = ""
    Call GetTotal
    
    
        
    AddToLView = True
End Function

Private Sub GetTotal()

Dim i As Integer
Dim DebitTotal As Double
Dim CreditTotal As Double

For i = 1 To LView.ListItems.Count
    DebitTotal = DebitTotal + Val(LView.ListItems(i).ListSubItems(3).Text)
Next i

Label1(1).Caption = DebitTotal

End Sub


Private Function SaveTrans() As Boolean
SaveTrans = False

If LView.ListItems.Count = 0 Then
    MsgBox "No Transaction to save!", vbInformation
    Exit Function
ElseIf Val(Label1(1).Caption) = 0 Then
   MsgBox "Invalid Entry.", vbInformation
   Exit Function
ElseIf Val(Label1(1)) > Val(lblBal) Then
   If Not AllowOD(GetID(lblBankAcc)) Then
      MsgBox "Not Much Amount In " & Left(lblBalCap, Len(lblBalCap) - 2) & ".", vbInformation
      Exit Function
   End If
End If
    
    'Validate ChqNo
 If chkBank.Value = vbChecked Then
  Dim Issued As Boolean, ChqBookNo As Double
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
 End If
 
 
    
    
    
    ''''''''''''''''''''''Check For Voucher No. Used Or Not '''''''''''''''''

Dim VNo As String
VNo = getNextVchrNo(DTPicker1.Value, "PV")
If Label1(4) <> VNo Then
   MsgBox "Voucher No. : " & Label1(4) & " Has Been Used By Other Transaction." & vbNewLine & _
       "Voucher No. : " & VNo & " Will Be Used For This Transaction.", vbInformation
       Label1(4) = VNo
End If
 
    
    
    
    
    
    Dim i As Integer
    Dim INextSno As Double
    Dim Bal As Double
    
    Dim CreditAcc As String
      Dim SlipNo As String
      
      INextSno = getNextSno(DTPicker1.Value, "Cheque")
      SlipNo = INextSno
     con.BeginTrans
     
     
   
   'Check Whether Expense is Charged To Bank Or To Cash in Hand
   
   If chkBank.Value = vbChecked Then
       CreditAcc = GetID(lblBankAcc)
      
      con.Execute "insert into Cheque Values(" & INextSno & ",'" & Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & _
      CreditAcc & "','" & Tbox(0) & "'," & Val(Label1(1)) & "," & lblBankAcc.Tag & ",'" & _
        txtChqNo & "'," & txtChqNo.Tag & ",'Open','" & DTPicker1 & "',1,1,0)"
       
      Call IssueChq(Val(txtChqNo.Tag), Val(txtChqNo))

   Else
      CreditAcc = CashInHand

   End If
    
    
    
    
    
    
    
    INextSno = getNextSno(DTPicker1.Value)
    
    
    
    
    For i = 1 To LView.ListItems.Count
        Bal = getBalance(LView.ListItems(i).Text, DTPicker1.Value)
        Bal = Bal + Val(LView.ListItems(i).ListSubItems(3).Text) - Val(LView.ListItems(i).ListSubItems(4).Text)
        
        con.Execute "insert into Vouchers Values(" & INextSno & ",'" & Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & _
         Label1(4) & "','" & LView.ListItems(i).Text & "','" & LView.ListItems(i).ListSubItems(2).Text & "'," & _
         Val(LView.ListItems(i).ListSubItems(3).Text) & ",0," & _
         Bal & ",''," & SlipNo & ")"
         
         
         Call adjustLedger(LView.ListItems(i).Text, DTPicker1.Value, INextSno, Val(LView.ListItems(i).ListSubItems(3).Text))
         INextSno = INextSno + 1
    Next i
    
    
    ''''''''''''''''Charge To  Cash Entry'''''''''''''
    
    
    
    Bal = getBalance(CreditAcc, DTPicker1.Value)
    Bal = Bal - Val(Tbox(1))
    Dim TheDesc As String
    If LView.ListItems.Count > 1 Then
     TheDesc = "Expense Details Annexed"
    Else
     TheDesc = LView.ListItems(1).SubItems(2)
    End If
    
    'Credit Entry
    con.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & Label1(4) & "','" & CreditAcc & _
        "','" & TheDesc & "',0," & Label1(1) & "," & Bal & ",''," & SlipNo & ")"
    
    
    
    Call adjustLedger(CreditAcc, DTPicker1.Value, INextSno, -Val(Label1(1)))
    
    con.CommitTrans
    
    
     'Print If To Be Printed
        If chkPrint = vbChecked Then Call PrintVoucher(frmMain.cr1, Label1(4), True)
        Call PlaySaveSound
    
    SaveTrans = True
End Function





Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    If LView.ListItems.Count > 0 Then
        If MsgBox("There is a Transaction active." & vbNewLine & "Save or Cancel Transaction first.", vbInformation + vbOKOnly) Then
            Cancel = True
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
    If LView.ListItems.Count > 0 Then
        With cmdDelete
            .Move .Left, LView.Top + (LView.SelectedItem.Top)
            .Visible = True
        End With
    End If
End Sub



Private Sub LView_DblClick()
    If LView.ListItems.Count > 0 Then
        With LView.SelectedItem
            If .ListSubItems(3).Text > 0 Then
                txtEdit.Width = LView.ColumnHeaders(4).Width + 30
                txtEdit.Move LView.ColumnHeaders(4).Left + LView.Left + 30, LView.Top + .Top
                txtEdit.Text = .ListSubItems(3).Text
                txtEdit.Tag = 3
            Else
                txtEdit.Width = LView.ColumnHeaders(5).Width + 30
                txtEdit.Move LView.ColumnHeaders(5).Left + LView.Left + 30, LView.Top + .Top
                txtEdit.Text = .ListSubItems(4).Text
                txtEdit.Tag = 4
            End If
        End With
        LView.Tag = LView.SelectedItem.Index
        txtEdit.Visible = True
        txtEdit.SelStart = 0
        txtEdit.SelLength = Len(txtEdit)
        txtEdit.SetFocus
    End If
End Sub

Private Sub LView_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ERR
If LView.ListItems.Count > 0 Then
    If KeyCode = 46 Then
        Call CmdDelete_Click
    End If
End If
Exit Sub
ERR:
    
    MsgBox ERR.Description
    

End Sub

Private Sub LView_Validate(Cancel As Boolean)
On Error Resume Next
    cmdDelete.Visible = False
    txtEdit.Visible = False
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
Tbox(0).Text = ""
Tbox(1).Text = ""
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
   FALL.Move (Me.ScaleWidth - FALL.Width) / 2, (Me.ScaleHeight - FALL.Height) / 2
End Sub


Private Sub ShowBals()
If chkBank.Value = vbChecked Then
   lblBalCap = "Bank Balance :"
   lblBal = Format(getBalance(GetID(lblBankAcc), DTPicker1), "0.00")
Else
   lblBalCap = "Cash In Hand :"
   lblBal = Format(getBalance(CashInHand, DTPicker1), "0.00")
End If
End Sub

