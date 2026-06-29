VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "combolist.ocx"
Begin VB.Form frmMakerTempLoan 
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleMode       =   0  'User
   ScaleWidth      =   10721.69
   WindowState     =   2  'Maximized
   Begin VB.Frame FAcc 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   8400
      Left            =   60
      TabIndex        =   3
      Top             =   -45
      Width           =   11865
      Begin VB.CheckBox chkPayByCheque 
         Caption         =   "Pay through Cheque"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   1815
         TabIndex        =   29
         Top             =   6825
         Width           =   1785
      End
      Begin VB.Frame FraBank 
         Caption         =   "Cheque Details :"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Left            =   60
         TabIndex        =   18
         Top             =   6810
         Width           =   11715
         Begin VB.CheckBox ChkClear 
            Caption         =   "Clear Cheque "
            Enabled         =   0   'False
            Height          =   225
            Left            =   10275
            TabIndex        =   21
            ToolTipText     =   "Clear This Cheque Automatically"
            Top             =   615
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
            Left            =   5670
            TabIndex        =   20
            Top             =   555
            Width           =   1485
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
            Left            =   5265
            Picture         =   "frmMakerTempLoan.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   555
            Width           =   360
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   8850
            TabIndex        =   22
            Top             =   555
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   122093571
            CurrentDate     =   37250
         End
         Begin MSForms.ComboBox cmbCheckType 
            Height          =   315
            Left            =   7185
            TabIndex        =   28
            Top             =   555
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
         Begin VB.Label lblBankAcc 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   45
            TabIndex        =   27
            Top             =   570
            Width           =   5190
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Bank Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   11
            Left            =   60
            TabIndex        =   26
            Top             =   330
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cheque No."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   12
            Left            =   5685
            TabIndex        =   25
            Top             =   315
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cheque Type"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   13
            Left            =   7200
            TabIndex        =   24
            Top             =   300
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   14
            Left            =   8835
            TabIndex        =   23
            Top             =   345
            Width           =   390
         End
      End
      Begin VB.TextBox txtTotalAdvance 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   7740
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   7950
         Width           =   1275
      End
      Begin VB.TextBox txtTotalValue 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   7935
         Width           =   1275
      End
      Begin VB.CommandButton cmdDistribute 
         Caption         =   "Distribute"
         Height          =   315
         Left            =   2055
         TabIndex        =   12
         Top             =   7950
         Width           =   1215
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   735
         TabIndex        =   10
         Top             =   7950
         Width           =   1275
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3945
         TabIndex        =   9
         Top             =   1800
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.ComboBox cmbRange 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMakerTempLoan.frx":0372
         Left            =   5970
         List            =   "frmMakerTempLoan.frx":0388
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   330
         Width           =   1650
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   9105
         TabIndex        =   1
         Top             =   330
         Width           =   1515
         _ExtentX        =   2672
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
         Format          =   122159107
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   7605
         TabIndex        =   0
         Top             =   330
         Width           =   1515
         _ExtentX        =   2672
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
         Format          =   122159107
         CurrentDate     =   37055
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   75
         TabIndex        =   5
         Top             =   345
         Width           =   5880
         _ExtentX        =   10372
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
         ListIndex       =   -1
         Appearance      =   0
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5685
         Left            =   60
         TabIndex        =   6
         Top             =   765
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   10028
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   15199215
         BorderStyle     =   1
         Appearance      =   0
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
            Text            =   "S. #"
            Object.Width           =   873
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PO #"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Item ID"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Issd Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Price"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Value"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Adv Taken"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "New Adv."
            Object.Width           =   1764
         EndProperty
      End
      Begin MSForms.ComboBox cmbCashInHand 
         Height          =   300
         Left            =   1185
         TabIndex        =   31
         Top             =   6465
         Width           =   3165
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5583;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cash Account :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   30
         Top             =   6510
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Advance :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   2
         Left            =   6585
         TabIndex        =   17
         Top             =   7995
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Value :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   4350
         TabIndex        =   15
         Top             =   7995
         Width           =   900
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         Caption         =   $"frmMakerTempLoan.frx":03D8
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
         Index           =   1
         Left            =   75
         TabIndex        =   13
         Top             =   135
         Width           =   10515
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Amount :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   11
         Top             =   7980
         Width           =   660
      End
      Begin MSForms.CommandButton cmdPost 
         Height          =   375
         Left            =   9060
         TabIndex        =   8
         Top             =   7935
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Post Selected"
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
         Left            =   10440
         TabIndex        =   7
         Top             =   7935
         Width           =   1335
         ForeColor       =   0
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
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   360
         Left            =   10635
         TabIndex        =   2
         Top             =   315
         Width           =   1125
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "1984;635"
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
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   10
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakerTempLoan.frx":0493
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakerTempLoan.frx":064F
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PopMenu"
   End
End
Attribute VB_Name = "frmMakerTempLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCourierInvoice As Boolean
Dim bShipping As Boolean
Dim bCustomInvoice As Boolean

Public Sub ShowMe(Optional p_bCustomInvoice As Boolean = True)
    bCustomInvoice = p_bCustomInvoice
    Me.Show
End Sub

Public Sub ShowForShipping()
    bShipping = True
    Me.Show
End Sub

Private Sub chkPayByCheque_Click()
    If chkPayByCheque.Value = vbChecked Then
        FraBank.Enabled = True
    Else
        FraBank.Enabled = False
    End If
End Sub

Private Sub cmbRange_Click()

    DtTo = GetServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    Select Case cmbRange.ListIndex
        Case 0
            DtFrom = Date
        Case 1
            DtFrom = DateAdd("d", -15, Date)
        Case 2
            DtFrom = DateAdd("d", -30, Date)
        Case 3
            DtFrom = DateAdd("d", -60, Date)
        Case 4
            DtFrom = DateAdd("d", -90, Date)
        Case 5
            DtFrom.Enabled = True
            DtTo.Enabled = True
            DtFrom.SetFocus
            Exit Sub
    End Select
    
End Sub

Private Sub cmdBank_Click()

    Dim BankID As Double
    lblBankAcc = SelectBankAcc(BankID)
    lblBankAcc.Tag = BankID
    
    'Get ChqNo
         
    Dim ChqBookNo As Double
    TBox(2) = GetNextChqNo(GetID(lblBankAcc), ChqBookNo)
    TBox(2).Tag = ChqBookNo
    
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdDistribute_Click()

    If Val(txtAmount) <= 0 Then
        MsgBox "Invalid Amount.", vbInformation
        Exit Sub
    End If
    Dim i As Integer, lTotalAmount As Long, lAmtToDist As Long
    lAmtToDist = Val(txtAmount)
    For i = 1 To LV.ListItems.count
        lTotalAmount = lTotalAmount + Val(LV.ListItems(i).SubItems(6))
    Next
    Dim lAmt As Long, dPer As Double
    For i = 1 To LV.ListItems.count
        lAmt = Val(LV.ListItems(i).SubItems(6))
        dPer = lAmt / lTotalAmount * 100
        LV.ListItems(i).SubItems(8) = Round((lAmtToDist / 100) * dPer)
        LV.ListItems(i).Checked = True
    Next
    Call CalculateTotals
End Sub

Private Sub cmdPost_Click()

    On Error GoTo err
    Dim i As Integer
    Dim bSelected As Boolean
    bSelected = False
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            bSelected = True
            Exit For
        End If
    Next
    If bSelected = False Then
        MsgBox "No Order Selected.", vbInformation
        Exit Sub
    End If
    
    If chkPayByCheque.Value = vbUnchecked Then
        If cmbCashInHand.MatchFound = False Then
            MsgBox "Please Select Cash Account.", vbInformation
            Exit Sub
        End If
    End If
    
    Dim dValue As Double
    dValue = 0
    For i = 1 To LV.ListItems.count
        dValue = dValue + Val(LV.ListItems(i).SubItems(8))
    Next
    Dim strVchrNo As String
    If chkPayByCheque.Value = vbUnchecked Then
        
        'Now Post Vouchers....
        'Services / Work Order AccNo: 41-003-19001
        Dim strDebitAccNo As String, strMakerAccNo As String
        strDebitAccNo = cmbCashInHand.List(cmbCashInHand.ListIndex, 1) '"15-001" '"41-003-19001"
        strMakerAccNo = GetSingleStringValue("AccNo", "Makers", " WHERE VendID=" & cmbCust.id)

        Dim objVchr As New clsVouchers
        objVchr.VoucherDate = GetServerDate
        objVchr.VoucherType = "CPV"
        
        objVchr.AddVoucherHead strMakerAccNo, dValue, "Temp. Advance for " & cmbCust.Text & " against Diff. Orders"
        objVchr.AddVoucherHead strDebitAccNo, -dValue, "Temp. Advance for " & cmbCust.Text & " against Diff. Orders"
        
        Call StartTrans(con)
        strVchrNo = objVchr.PostVoucher(, True)

    Else
        Dim ChqBookNo As Double
       If lblBankAcc = "" Then
            MsgBox "Invalid Bank Selected.", vbInformation
            Exit Sub
        ElseIf Not ValidChqNo(GetID(lblBankAcc), TBox(2), False, ChqBookNo) Then
            If Issued Then
                MsgBox "The Cheque With Number '" & txtChqNo & "' Has Been Issued.", vbInformation
            Else
                MsgBox "Invalid Cheque No.", vbInformation
            End If
            TBox(2).SetFocus
            Exit Sub
        End If
        TBox(2).Tag = ChqBookNo
        Call StartTrans(con)
        strVchrNo = SaveTransForCheque(dValue)
    End If
     
    Dim lGroupID As Long
    con.Execute "INSERT INTO MakerTemporaryLoansGroup(MakerID,DT,UserName,MachineName,Active) VALUES(" & cmbCust.id & _
     ",'" & GetServerDate & "','" & CurrentUserName & "','" & strComputerName & "',1)"
    
    lGroupID = GetSingleLongValue("MAX(EntryID)", "MakerTemporaryLoansGroup")
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                con.Execute "INSERT INTO MakerTemporaryLoans(Issuance_RefID,Amount,VchrNo,GroupID) VALUES(" & _
                 Val(.Tag) & "," & Val(.SubItems(8)) & ",'" & strVchrNo & "'," & lGroupID & ")"
            End If
        End With
    Next
    
    con.CommitTrans
    MsgBox "Posted Successfuly !!!", vbInformation
    Call PrintVoucher(MainForm.cr1, strVchrNo, False, True)
    
    Call RefreshLV
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Function SaveTransForCheque(dValue As Double) As String

    On Error GoTo err
    Dim i As Integer
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim Bal As Double
    Dim RcptNo As String
    
    RcptNo = "" 'IIf(Len(Trim(Tbox(3))) = 0, "", "Rcpt # :" & Tbox(3))
    
    'Get Serial # and Balances
    INextSno = getNextSno(Date, "Cheque")
    SnoToUse = INextSno
    Dim strMakerAccNo As String
    strMakerAccNo = GetSingleStringValue("AccNo", "Makers", " WHERE VendID=" & cmbCust.id)
    bankDesc = Right(lblBankAcc, Len(lblBankAcc) - InStr(lblBankAcc, "]")) & " Chq # :" & TBox(2)
    Dim ChqIsDue As Integer
    If DTPicker2.Value > Date Then
        ChqIsDue = 1
    Else
        ChqIsDue = 0
    End If
    
    con.Execute "INSERT INTO Cheque(SNo,CDate,AccNo,Description,Amount,BankID,chequeno,ChqBookNo,chequeType,chequeDate,Posted,Payment,Bounced,ChqIsDue,BankAccNo) VALUES(" & INextSno & ",'" & Format(Date, "dd-MMM-yyyy") & "','" & _
     strMakerAccNo & "','" & txtVchrDescription & " (" & bankDesc & ") '," & dValue & "," & lblBankAcc.Tag & ",'" & _
     TBox(2) & "'," & TBox(2).Tag & ",'" & cmbCheckType & "','" & DTPicker2 & "',0,1,0," & ChqIsDue & ",'" & GetID(lblBankAcc) & "')"
     
    'Remove the Cheque From Cheque book
    Call IssueChq(Val(TBox(2).Tag), TBox(2))
     
    SaveTransForCheque = SavePaymentTrans(SnoToUse, dValue)

    Call PlaySaveSound
    'MsgBox "Transaction Saved Successfuly.", vbInformation
   
    Exit Function
err:
    MsgBox "Transaction Terminated Due to ERROR:" & err.Number & vbNewLine & err.Description, vbCritical
    SaveTransForCheque = ""
    'con.RollbackTrans
    
End Function

Private Function SavePaymentTrans(ChqSNo As Double, dValue As Double) As String

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
    
    NextVchrNo = getNextVchrNo(Dated, "BPV")
    INextSno = getNextSno(Dated)
    
    SnoToUse = INextSno
    
    'Party Account's Balance
    Dim strDebitAccNo As String, strMakerAccNo As String
    strDebitAccNo = "24-005-19001" '"41-003-19001"
    strMakerAccNo = GetSingleStringValue("AccNo", "Makers", " WHERE VendID=" & cmbCust.id)
       
    Bal = getBalance(strMakerAccNo, Dated)
    dAmount = dValue
    
    sScNo = ChqSNo
    
    'Bank Account No.
    sCIH = GetID(lblBankAcc)
    
    Bal = Bal + dAmount
    
    'Debit Entry
    bankDesc = Right(lblBankAcc, Len(lblBankAcc) - InStr(lblBankAcc, "]")) & " Chq # :" & TBox(2) & " Against Diff. Orders"
    con.Execute "INSERT INTO Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & NextVchrNo & "','" & strMakerAccNo & _
        "','" & bankDesc & "'," & dAmount & ",0," & Bal & ",''," & sScNo & ")"
    
    Call adjustLedger(strMakerAccNo, Dated, INextSno, Val(dAmount))
    
    Dim ChqNo As String
    INextSno = INextSno + 1
    Bal = getBalance(sCIH, Dated)
    dAmount = dValue
    Bal = Bal - dAmount
    'Description
    
    Desc = "Cheque Paid To " & cmbCust.Text & " Chq# : " & TBox(2)
    ChqNo = TBox(2)
           
    'Credit Entry
    
    con.Execute "INSERT INTO Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & NextVchrNo & "','" & sCIH & _
        "','" & Desc & "',0," & dAmount & "," & Bal & ",''," & sScNo & ")"
    
    Call adjustLedger(sCIH, Dated, INextSno, -Val(dAmount))
    
    Dim rsCopy As New ADODB.Recordset
    With rsCopy
        .CursorLocation = adUseClient
        .Open "Select * from ChequeDetail Where (Sno=" & sScNo & ")", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            con.Execute "Insert into VoucherDetail Values('" & ![BalancedSno] & "','" & SnoToUse & "'," & ![Amount] & ")"
            .MoveNEXT
        Loop
        .Close
    End With
    Set rsCopy = Nothing
    con.Execute "Update Cheque Set Posted=1 Where Sno=" & sScNo
    
    con.Execute "Insert Into VoucherInfo(VchrNo,UserName,MachineName,DT,PostedThroughJVForm) Values('" & NextVchrNo & _
     "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "',0)"
     
    'SavePaymentTrans = True
    SavePaymentTrans = NextVchrNo
    Exit Function
err:
    'MsgBox "Transaction Terminated Due To Following Error;" & vbNewLine & err.Description
    SavePaymentTrans = ""
End Function

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub DtFrom_Change()
    'Call setDateRange(DtFrom)
End Sub

Private Sub DtTo_Change()
    'Call setDateRange(DtTo)
End Sub


Private Sub Form_Load()
    
    mnupop.Visible = False
    cmbCust.ListHeight = 3800
    'cmbCust.AddItem "<All Makers>", "0"
    Call cmbCust.AddVals(con, "'{ ' + VendID1 + ' } ' + VenderName ", "Makers", "VendID", " WHERE VendID<>" & iFactoryMakerID)
    cmbCust.id = "0"
    DtFrom = Date
    DtTo = Date
    cmbRange.ListIndex = 0
    DTPicker2.Value = Date
    cmbCheckType.AddItem "Open"
    cmbCheckType.AddItem "Payees A/C Only"
    cmbCheckType.ListIndex = 0
    
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT AccNo,AccTitle FROM VActiveAccounts WHERE SubAccOf='" & CashInHand & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCashInHand.Clear
        Do Until .EOF
            cmbCashInHand.AddItem !AccTitle & ""
            cmbCashInHand.List(.AbsolutePosition - 1, 1) = !AccNo & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    If cmbCashInHand.ListCount > 0 Then cmbCashInHand.ListIndex = i
    
End Sub

Private Sub RefreshLV()
    
    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "MakerOrdersForTempLoan_SP"
        .Parameters("@MakerID").Value = cmbCust.id
        .Parameters("@DTFrom").Value = DtFrom.Value
        .Parameters("@DTTo").Value = DtTo.Value + 1
        Set rs = .Execute
    End With
    
    Dim Cond As String
    Dim DTRange As String
    Dim ITM As ListItem
      
    Screen.MousePointer = vbHourglass
    
'    If cmbCust.ID <> "0" Then
'        Cond = " Where VendID=" & cmbCust.ID & " "
'    End If
   
    With rs
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, Val(![EntryID] & "") & "'", .AbsolutePosition)
            ITM.Tag = Val(![EntryID] & "")
            ITM.ListSubItems.add , , !RecieptID & ""
            ITM.ListSubItems.add , , Format(!DT, "dd-MM-yy")
            ITM.ListSubItems.add , , !ITemID & ""
            ITM.ListSubItems.add , , Val(!IssQty & "")
            ITM.ListSubItems.add , , Val(!Price & "")
            ITM.ListSubItems.add , , Val(!Value & "")
            ITM.ListSubItems.add , , Val(!TempAdvance & "")
            ITM.ListSubItems.add , , "0"
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Call CalculateTotals
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    FAcc.Top = (Me.ScaleHeight - FAcc.Height) / 2
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    Exit Sub
    If LV.SortKey + 1 = ColumnHeader.Index Then
        If LV.SortOrder = lvwAscending Then
            LV.SortOrder = lvwDescending
            ColumnHeader.Icon = "Down"
        Else
            LV.SortOrder = lvwAscending
            ColumnHeader.Icon = "Up"
        End If
    Else
        LV.ColumnHeaderIcons = Nothing
        Set LV.ColumnHeaderIcons = ImageList2
        LV.SortKey = ColumnHeader.Index - 1
        LV.SortOrder = lvwAscending
        ColumnHeader.Icon = "Down"
    End If
    Call RefreshSnos
    'If LV.Sorted = False Then LV.Sorted = True
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    If LV.SelectedItem.Checked = False Then Exit Sub
    
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(9).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(9).Width
        .Text = Val(LV.SelectedItem.SubItems(8))
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
    
End Sub

Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        Call LV_DblClick
    Else
        Item.SubItems(8) = "0"
    End If
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button <> 2 Then Exit Sub
    Exit Sub
    If LV.ListItems.count = 0 Then
        Call ShowMenus(False)
    Else
        Call ShowMenus(True)
    End If
'    If bShipping Then
'        ShowMenusForShipping
'    ElseIf bCourierInvoice Then
'        ShowMenusforCourier
'    Else
'        If LV.ListItems.Count = 0 Then
'            Call ShowMenus(False)
'        Else
'            Call ShowMenus(True)
'        End If
'    End If
    'If LV.ListItems.Count = 0 Or Button <> 2 Or OutOfFinancialYear Then Exit Sub

    'if a Voucher Entry Show Delete Menu
    'If InStr(LV.SelectedItem, "-") > 0 Then Me.PopupMenu mnuPop
    
End Sub

Private Sub ShowMenusForShipping()

    On Error Resume Next
    Dim c As Control
    For Each c In Controls
        If TypeOf c Is Menu Then
            If c.Name <> "mnuPrintPList" Then
                c.Visible = False
            End If
        End If
    Next
    
    On Error GoTo 0
    
    If LV.ListItems.count > 0 Then
        mnuPrintPList.Visible = True
        mnuCustomPList.Visible = True
        mnuCustomPListOther.Visible = True
        mnuCommercialPList.Visible = True
        mnuCommercialPListOther.Visible = True
        mnuDashPackingLabels.Visible = True
        mnuPrintPackingLabels.Visible = True
    End If
'    mnuPrintDecDocs.Visible = False
    Me.PopUpMenu mnupop
    
End Sub

Private Sub ShowMenus(vis As Boolean)
        
    
    Me.PopUpMenu mnupop
    
End Sub

Private Sub ShowMenusforCourier()
         
    On Error Resume Next
    Dim c As Control
    For Each c In Controls
        If TypeOf c Is Menu Then
            c.Visible = False
        End If
    Next
    
    On Error GoTo 0
    mnuAdd.Visible = True
    If LV.ListItems.count > 0 Then
        dd.Visible = True
        mnuEdit.Visible = True
        mnuComDocs.Visible = True
        mnuPrintComInv.Visible = True
    End If
    mnuPrintDecDocs.Visible = False
    Me.PopUpMenu mnupop
    
End Sub


Private Sub mnuAdd_Click()
    'Load frmNewCustomInvoice
    'frmNewCustomInvoice.Show
    'ShowForm frmNewCustomInvoice, MainForm, False, "New Custom Invoice"
    'frmNewCustomInvoice.AddNewProforma
    Load frmNewCustomInvoice
    frmNewCustomInvoice.AddNewInvoice True, bCourierInvoice
End Sub

Private Sub mnuAddShippInfo_Click()

    Load frmEditShippInfo
    'ShowForm f, MainForm, , "Shipp. Info"
    Call frmEditShippInfo.EditEntry(LV.SelectedItem.SubItems(3))
    
End Sub

Private Sub mnuBillOfExchange_Click()
    Call PrintBillOfExchange
End Sub

Private Sub PrintBillOfExchange(Optional Firstrpt As Boolean = True)

    On Error GoTo err
    Me.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptBillOfExchange.rpt")
    rpt.FormulaFields(3).Text = Firstrpt
    Dim frm As New frmPrevRpt
    frm.ShowReport "{CustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    Me.MousePointer = vbNormal
    
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description

End Sub

    
Private Sub mnuBillOfExchange2_Click()
    Call PrintBillOfExchange(False)
End Sub

Private Sub mnuCommercialPList_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3), False, True, LV.SelectedItem.SubItems(1) = "DECATHLON", LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2))
End Sub

Private Sub mnuCommercialPListOther_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3), True, True, LV.SelectedItem.SubItems(1) = "DECATHLON", LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2))
End Sub

Private Sub mnuCoveringLetter_Click()
    
    On Error GoTo err
    Me.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptComCoveringLetter.rpt")
    If LCase(CurrentUserName) = "mnaeem" Then
        rpt.FormulaFields(3).Text = "'Naeem Ahmed'"
    End If
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "{CustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Private Sub mnuCustomPList_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3))
End Sub

Private Sub mnuCustomPListOther_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3), True)
End Sub

Private Sub mnuDelPackingInfo_Click()
    
    If LV.ListItems.count = 0 Then Exit Sub
    Dim iAffect As Integer
    If MsgBox("Are you sure to Delete Packing List of Invoice # " & LV.SelectedItem.SubItems(3) & " ?", vbQuestion + vbYesNo) = vbYes Then
        Call StartTrans(con)
        con.Execute "Delete From CustomPList Where CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'", iAffect
        con.CommitTrans
        If iAffect > 0 Then
            LV.SelectedItem.ListSubItems(2).Tag = 0
        End If
    End If
    
End Sub

Private Sub mnuEdit_Click()
    Call LV_DblClick
End Sub

Private Sub mnuEForm_Click()

    Load frmPrintFormE
    If LV.ListItems.count > 0 Then
        frmPrintFormE.ShowForm LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
    Else
        frmPrintFormE.Show , MainForm
    End If
    
End Sub

Private Sub mnuGSP_Click()

'    Load frmGSP
'    If LV.ListItems.Count > 0 Then
'        frmGSP.ShowForm LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
'    Else
'        frmGSP.Show , MainForm
'    End If
    
End Sub

Private Sub mnuPacking_Click()

'    If Left(mnuPacking.Caption, 3) = "Add" Then
'
'        Load frmNewPList
'        frmNewPList.ShowMe LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
'    Else
'        Load frmNewPList
'        Call frmNewPList.EditPList(LV.SelectedItem.SubItems(3))
'        ShowForm frmNewPList, MainForm, , "Packing List"
'    End If
    
End Sub

Private Sub mnuPrint_Click()
    Call PrintCustomInvoice
End Sub

Private Sub PrintCustomInvoice(Optional bImpMaterialDetail As Boolean = False)

    On Error GoTo err

    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    If bCustomInvoice Then
        Set rpt = rptApp.OpenReport(RptPath & "\rptCustomInvoice.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\rptCommercialInvoice.rpt")
    End If
    'rpt.FormulaFields(15).Text = bImpMaterialDetail
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VrptCustomInvoiceDetail.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt
    
    Screen.MousePointer = vbDefault


    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub mnuPrint10_Click()
    Call PrintCommercialInvoice(True, False)
End Sub

Private Sub mnuPrint5_Click()
    Call PrintCommercialInvoice(True, True)
End Sub

Private Sub mnuPrintBillofExchange_Click()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptBillOfExchange.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VRptProformas.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintComInv_Click()
    Call PrintCommercialInvoice
End Sub

Private Sub PrintCommercialInvoice(Optional bPer As Boolean = False, Optional bFivePer As Boolean = True)
    
    On Error GoTo err

    Screen.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt
    Dim rpt As CRAXDDRT.Report
    
    If UCase(LV.SelectedItem.SubItems(1)) = "MIKASA" And LV.SelectedItem.SubItems(2) = "JAPAN" Then
        Set rpt = rptApp.OpenReport(rptexportpath & "\rptCommercialInvoiceMikasaJP.rpt")
    ElseIf bPer Then
        Set rpt = rptApp.OpenReport(rptexportpath & "\rptCommercialInvoiceFairTrade.rpt")
        rpt.FormulaFields(22).Text = bFivePer
    ElseIf LV.SelectedItem.SubItems(1) = "DECATHLON" Then
        Set rpt = rptApp.OpenReport(rptexportpath & "\rptDecathlonCommercialInvoice.rpt")
    Else
        Set rpt = rptApp.OpenReport(rptexportpath & "\rptCommercialInvoice.rpt")
    End If
    f.ShowReport "{VrptCustomInvoiceDetail.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt
    
    If Not bPer Then
        If LV.SelectedItem.SubItems(2) = "Canada" Then
            Dim rpt1 As CRAXDDRT.Report
            Set rpt1 = rptApp.OpenReport(rptexportpath & "\rptCommercialInvoice.rpt")
            Dim f1 As New frmPrevRpt
            Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
            Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
            
            Set FormulaFields = rpt1.FormulaFields
            
            For Each FormulaField In FormulaFields
                If FormulaField.Name = "{@forHeading}" Then
                    FormulaField.Text = "'C A N A D I A N   C U S T O M S   I N V O I C E'"
                    Exit For
                End If
            Next
            f1.ShowReport "{VrptCustomInvoiceDetail.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt1
        End If
    End If
    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub mnuPrintSummary_Click()

    On Error GoTo err
    Me.MousePointer = vbHourglass
    
    Call StartTrans(con)
    con.Execute "Delete From PrintFOrderList"
    
    For i = 1 To LV.ListItems.count
        con.Execute "Insert Into PrintFOrderList(Sno,Customer,Country," & _
         "OrderNo,OrderDT,OrderAmt,Status,Curr) Values(" & LV.ListItems(i) & _
         ",'" & LV.ListItems(i).SubItems(1) & "','" & LV.ListItems(i).SubItems(2) & _
         "','" & LV.ListItems(i).SubItems(3) & "','" & _
         LV.ListItems(i).SubItems(4) & "'," & Val(LV.ListItems(i).ListSubItems(5).Tag & "") & _
         ",'" & LV.ListItems(i).SubItems(6) & "','" & LV.ListItems(i).Tag & "')"
    Next i
    
    con.CommitTrans
    
    Load frmPrevRpt
    
    Dim strFromTo As String
    Dim iNumber As Integer
    If cmbRange.ListIndex = 0 Then
        iNumber = -15
    ElseIf cmbRange.ListIndex = 1 Then
        iNumber = -30
    ElseIf cmbRange.ListIndex = 2 Then
        iNumber = -60
    ElseIf cmbRange.ListIndex = 3 Then
        iNumber = -90
    End If
    
    If cmbRange.ListIndex = 4 Then
        strFromTo = Format(DtFrom, "dd-MMM-yyyy") & " To  " & Format(DtTo, "dd-MMM-yyyy")
    Else
        strFromTo = Format(DateAdd("d", iNumber, ServerDate), "dd-MMM-yyyy") & " To  " & Format(ServerDate, "dd-MMM-yyyy")
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptFProformaList.rpt")
    rpt.FormulaFields(3).Text = "'" & cmbCust.Text & "'"
    rpt.FormulaFields(4).Text = "'" & strFromTo & "'"
    
    frmPrevRpt.ShowReport "", rpt
    
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Private Sub mnuPrintCoverLetter_Click()

    On Error GoTo err

    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptDecathlonCoveringLetter.rpt")
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VCustomInvoice.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt

    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub mnuPrintDecDocs_Click()
    Load frmPrintDecathlonDocs
    frmPrintDecathlonDocs.ShowMe (LV.SelectedItem.SubItems(3))
End Sub

Private Sub mnuPrintInvWDetail_Click()

    Call PrintCustomInvoice(True)
    
End Sub

Private Sub mnuPrintPackingLabels_Click()

    On Error GoTo err
    
    If LV.SelectedItem.SubItems(1) <> "DECATHLON" Then
        MsgBox "Packing Labels are available only for DECATHLON at the moment.", vbInformation
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report, f As New frmPrevRpt
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptDecathlonPackingLabels.rpt")
    f.ShowReport "{VCustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    
    Me.MousePointer = vbNormal
    
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub mnuShippingInstructions_Click()
'    Load frmShippingInstructions
'    If LV.ListItems.Count = 0 Then
'        frmShippingInstructions.Show , MainForm
'    Else
'        frmShippingInstructions.ShowForm LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
'    End If
End Sub

Private Sub mnuUpdateCustomPrice_Click()
    If LV.ListItems.count = 0 Then Exit Sub
    If MsgBox("Are you Sure, You want to Update Custom Prices?", vbInformation + vbYesNo) = vbNo Then Exit Sub
    
    Dim dCustomPrice As Double
    dCustomPrice = Val(con.Execute("Select SUM(Qty*Price)/SUM(Qty) From CustomInvoiceItems Where CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'").Fields(0).Value & "")
    con.Execute "Update CustomInvoiceItems SET CustomPrice=" & dCustomPrice & " WHERE CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'"
    
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
Private Sub RefreshSnos()
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Text = i
    Next i
End Sub

Private Sub mnuPrintList_Click()

    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\IssList.rpt")
    Dim strList As String
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                strList = strList & .Tag & ","
            End If
        End With
    Next
    Dim f As New frmPrevRpt
    f.ShowReport "{VendIssued.EntryID} IN[" & Left(strList, Len(strList) - 1) & "]", rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintPurchaseOrder_Click()
    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\IssSlip.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VendIssued.EntryID}=" & LV.SelectedItem.Tag, rpt
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Val(txtEdit) <= 0 Then
            MsgBox "Invalid Amount!", vbInformation
            Exit Sub
        End If
        LV.SelectedItem.SubItems(8) = Val(txtEdit)
        txtEdit.Visible = False
        Call CalculateTotals
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
        If Val(LV.SelectedItem.SubItems(8)) = 0 Then
            LV.SelectedItem.Checked = False
        End If
    End If
    
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub

Private Sub CalculateTotals()
    Dim i As Integer
    Dim lTotalValue As Long, lTotalAdvance As Long
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            lTotalValue = lTotalValue + Val(.SubItems(6))
            lTotalAdvance = lTotalAdvance + Val(.SubItems(8))
        End With
    Next
    txtTotalValue = lTotalValue
    txtTotalAdvance = lTotalAdvance
End Sub
