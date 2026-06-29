VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJV 
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   12870
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
   ScaleHeight     =   8715
   ScaleWidth      =   12870
   WindowState     =   2  'Maximized
   Begin VB.Frame FALL 
      Height          =   8640
      Left            =   75
      TabIndex        =   11
      Top             =   0
      Width           =   12645
      Begin VB.CommandButton cmdAddNewMaker 
         Caption         =   "New Maker"
         Height          =   345
         Left            =   1575
         TabIndex        =   32
         Top             =   7425
         Width           =   1185
      End
      Begin VB.CommandButton cmdNewVender 
         Caption         =   "New Vender"
         Height          =   345
         Left            =   345
         TabIndex        =   31
         Top             =   7425
         Width           =   1185
      End
      Begin VB.CommandButton cmdNewAccount 
         Caption         =   "New Account"
         Height          =   255
         Left            =   4230
         TabIndex        =   30
         Top             =   1035
         Width           =   1185
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
         TabIndex        =   27
         Top             =   -30
         Width           =   12630
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Journal Voucher"
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
            TabIndex        =   28
            Top             =   165
            Width           =   12525
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Journal Voucher"
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
            TabIndex        =   29
            Top             =   180
            Width           =   12510
         End
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
         Left            =   360
         TabIndex        =   25
         Top             =   8220
         Value           =   1  'Checked
         Width           =   1830
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   2010
         TabIndex        =   23
         Top             =   1590
         Visible         =   0   'False
         Width           =   4320
      End
      Begin VB.TextBox Tbox 
         Height          =   315
         Index           =   0
         Left            =   5460
         TabIndex        =   3
         Top             =   1275
         Width           =   3600
      End
      Begin VB.TextBox Tbox 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   9090
         TabIndex        =   4
         Top             =   1275
         Width           =   1380
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
         Left            =   11730
         Picture         =   "frmJV.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Add This Entry To Voucher"
         Top             =   1050
         UseMaskColor    =   -1  'True
         Width           =   525
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         ItemData        =   "frmJV.frx":0532
         Left            =   10470
         List            =   "frmJV.frx":053C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1275
         Width           =   1230
      End
      Begin VB.CommandButton cmdDelete 
         CausesValidation=   0   'False
         Height          =   285
         Left            =   10620
         Picture         =   "frmJV.frx":054F
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Remove This Entry From Voucher"
         Top             =   1650
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E3EDF9&
         CausesValidation=   0   'False
         Height          =   285
         Left            =   5820
         TabIndex        =   12
         Top             =   1035
         Visible         =   0   'False
         Width           =   1110
      End
      Begin MSComctlLib.ListView LView 
         Height          =   5775
         Left            =   330
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1605
         Width           =   11955
         _ExtentX        =   21087
         _ExtentY        =   10186
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Acc #"
            Object.Width           =   2910
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Acc Title"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   6085
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Debit"
            Object.Width           =   2434
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Credit"
            Object.Width           =   2434
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   4605
         TabIndex        =   0
         Top             =   675
         Width           =   2550
         _ExtentX        =   4498
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
         Format          =   295370755
         CurrentDate     =   37250
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
         Height          =   660
         Left            =   7005
         TabIndex        =   22
         Top             =   7845
         Width           =   5280
         Begin MSForms.CommandButton cmdSClose 
            Height          =   375
            Left            =   1800
            TabIndex        =   10
            Top             =   195
            Width           =   1665
            Caption         =   "Save & Close  "
            PicturePosition =   327683
            Size            =   "2937;661"
            Accelerator     =   83
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdSClear 
            Height          =   375
            Left            =   90
            TabIndex        =   9
            Top             =   195
            Width           =   1665
            Caption         =   "Save & New   "
            PicturePosition =   327683
            Size            =   "2937;661"
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
            Height          =   375
            Left            =   3495
            TabIndex        =   24
            Top             =   195
            Width           =   1665
            Caption         =   "Cancel            "
            PicturePosition =   327683
            Size            =   "2937;661"
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
      Begin MSForms.CommandButton cmdFinClosing 
         Height          =   405
         Left            =   2580
         TabIndex        =   26
         Top             =   7980
         Visible         =   0   'False
         Width           =   2310
         Caption         =   "Close Financial Year"
         PicturePosition =   327683
         Size            =   "4075;714"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Last Balance :"
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
         Index           =   7
         Left            =   8055
         TabIndex        =   21
         Top             =   705
         Width           =   1395
      End
      Begin VB.Label lblBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   9525
         TabIndex        =   20
         Top             =   690
         Width           =   1695
      End
      Begin MSForms.ComboBox cmbAccNo 
         Height          =   315
         Left            =   345
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1275
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
      Begin MSForms.ComboBox CmbAccTitle 
         Height          =   315
         Left            =   2010
         TabIndex        =   1
         Top             =   1275
         Width           =   3435
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "6059;556"
         ListWidth       =   8819
         ListStyle       =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmJV.frx":0A81
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
         Left            =   345
         TabIndex        =   19
         Top             =   1035
         Width           =   11325
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   9480
         TabIndex        =   18
         Top             =   7455
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   2
         Left            =   10905
         TabIndex        =   17
         Top             =   7455
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Totals :"
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
         Left            =   8730
         TabIndex        =   16
         Top             =   7500
         Width           =   735
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
         Left            =   360
         TabIndex        =   15
         Top             =   735
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
         Left            =   3990
         TabIndex        =   14
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   1530
         TabIndex        =   13
         Top             =   735
         Width           =   1560
      End
   End
End
Attribute VB_Name = "frmJV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsList As New ADODB.Recordset
Dim dontShowList As Boolean
Dim PermissionDenied As Boolean
Dim strEditVchrNo As String
Dim bPosting As Boolean

Public Sub EditVoucher(strVchrNo As String)

    '/* Just Delete the Voucher and
    On Error GoTo err
    
    strEditVchrNo = strVchrNo
    Dim rs As New ADODB.Recordset
    With rs
        .Open "", con, adOpenForwardOnly, adLockReadOnly
        .Close
    End With
    
    Exit Sub
err:
    MsgBox err.Description & vbNewLine & "Error Occured In: EditVoucher of frmJV", vbCritical
    
End Sub

Private Sub cmbACCNo_Change()
    If cmbAccNo.MatchFound Then Call cmbACCNo_Click
End Sub

Private Sub cmbACCNo_Click()
    If CmbAccTitle.ListIndex <> cmbAccNo.ListIndex Then CmbAccTitle.ListIndex = cmbAccNo.ListIndex
    Call ShowAccBalance
End Sub

Private Sub cmbAccNo_Validate(cancel As Boolean)
    If Not cmbAccNo.MatchFound Then cmbAccNo.ListIndex = CmbAccTitle.ListIndex
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
Private Sub cmbACCTitle_Change()
    If CmbAccTitle.MatchFound Then Call cmbACCTitle_Click 'Else If CmbAccTitle <> "" Then CmbAccTitle.DropDown
End Sub
Private Sub cmbACCTitle_Click()
    If cmbAccNo.ListIndex <> CmbAccTitle.ListIndex Then cmbAccNo.ListIndex = CmbAccTitle.ListIndex
End Sub

Private Sub CmbAccTitle_LostFocus()
   List1.Visible = False
End Sub

Private Sub CmbAccTitle_Validate(cancel As Boolean)
    If Not CmbAccTitle.MatchFound Then CmbAccTitle.ListIndex = cmbAccNo.ListIndex
End Sub
Private Sub CmbAccTitle_DropButtonClick()
   List1.Visible = False
End Sub

Private Sub cmdAddNewMaker_Click()

    Dim f As New frmMaker
    With f
        .add = True
        .Show 1
    End With
    Call FillAccCmbs
    
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
            Label1(2).Caption = ""
            cmbType.ListIndex = 0
        End If
    End With

End Sub


Private Sub cmdFinClosing_Click()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim AccList() As String
    Dim AccBalances() As Double
    Dim i As Integer
    Dim strVchrNo As String
    With rs
        .Open "Select AccNo,Sum(Credit)-Sum(Debit) As Balance From Vouchers Where VDate Between '7/1/05' And '6/30/06' AND Cast(Left(AccNo,2) as Int)>=31 AND AccNo<>'46-16001' GROUP BY AccNo Having Sum(Credit)-Sum(Debit)<>0  Order By AccNo", con, adOpenForwardOnly, adLockReadOnly
        ReDim AccList(.RecordCount)
        ReDim AccBalances(.RecordCount)
        i = 0
        Do Until .EOF
            AccList(i) = !AccNo & ""
            AccBalances(i) = !Balance
            i = i + 1
            .MoveNEXT
        Loop

        .Close
        strVchrNo = getNextVchrNo(#6/30/2006#, "JV")
    End With
    
    Set rs = Nothing
    
    Dim INextSno As Double
    INextSno = getNextSno(#6/30/2006#)
    Call StartTrans(con)
    Dim dCredit As Double, dDebit As Double
    Dim dTotalCredit As Double, dTotalDebit As Double
    
    For i = 0 To UBound(AccList) - 1
        If AccBalances(i) < 0 Then  'Debit Is Larget
            dCredit = Abs(AccBalances(i))
            dDebit = 0
        Else
            dDebit = Abs(AccBalances(i))
            dCredit = 0
        End If
        con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & Format(#6/30/2006#, "dd-MMM-yyyy") & "','" & _
         strVchrNo & "','" & AccList(i) & "','Year (2005-2006) Closing Voucher'," & _
         dDebit & "," & dCredit & ",0,'',0)"
                         
        If dCredit > 0 Then
            Call adjustLedger(AccList(i), #6/30/2006#, INextSno, -dCredit)
        Else
            Call adjustLedger(AccList(i), #6/30/2006#, INextSno, dDebit)
        End If
         
        INextSno = INextSno + 1
        dTotalCredit = dTotalCredit + dCredit
        dTotalDebit = dTotalDebit + dDebit
        Debug.Print i
    Next
    
    Dim strAccNo As String, dAmt As Double
    Dim dCreditAmt As Double, dDebitAmt As Double
    If dTotalCredit > dTotalDebit Then
        dDebitAmt = dTotalCredit - dTotalDebit
        dCreditAmt = 0
    Else
        dCreditAmt = dTotalDebit - dTotalCredit
        dDebitAmt = 0
    End If
    
    strAccNo = "46-16001"
    con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & Format(#6/30/2006#, "dd-MMM-yyyy") & "','" & _
         strVchrNo & "','" & strAccNo & "','Year (2005-2006) Closing Voucher'," & _
         dDebitAmt & "," & dCreditAmt & "," & dDebitAmt - dCreditAmt & ",'',0)"
    Call adjustLedger(strAccNo, #6/30/2006#, INextSno, dDebitAmt - dCreditAmt)
     
    con.Execute "Insert Into VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & strVchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
    con.CommitTrans
    
    Exit Sub
err:
    MsgBox err.Description
    
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

Private Sub cmdNewVender_Click()

    Dim f As New frmVender
    With f
        .add = True
        .Show 1
    End With
    Call FillAccCmbs
    
    
End Sub

Private Sub cmdSave_Click()
    Call AddToLView
End Sub

Private Sub cmdSClear_Click()

    If bPosting Then Exit Sub
    bPosting = True
    If SaveTrans Then
        MsgBox "Transaction Saved successfuly.", vbInformation
        bPosting = False
        Dim myDT As Date
        myDT = DTPicker1.Value
        
        Unload Me
        
        Me.Show
        DTPicker1.Value = myDT
        Call DTPicker1_Change
        Me.SetFocus
    End If
    bPosting = False
    
End Sub



Private Sub cmdSClose_Click()
    On Error GoTo err
    If bPosting Then Exit Sub
    bPosting = True
    If SaveTrans Then
        MsgBox "Transaction Saved successfuly.", vbInformation
        bPosting = False
        Unload Me
    End If
    bPosting = False
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub DTPicker1_Change()
    Call setDateRange(DTPicker1)
    Label1(4).Caption = getNextVchrNo(DTPicker1.Value, "JV")
End Sub

Private Sub DTPicker1_Click()
    Label1(4).Caption = getNextVchrNo(DTPicker1.Value, "JV")
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    
    'PermissionDenied = Not UserHasAccess("JournalVoucher")
    bPosting = False
    strEditVchrNo = ""
    Me.KeyPreview = True
    cmbType.ListIndex = 0
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
        .CursorLocation = adUseClient
        
        '.Open "Select AccNo,AccTitle from Accounts as Acc1 Where not AccNo Like '" & ACCRECNO & "%' and AccNo Not Like '" & ACCPAYNO & "%' AND Parent=0 AND Active=1", con, adOpenForwardOnly, adLockReadOnly
        
        .Open "Select AccNo,AccTitle from Accounts as Acc1 Where Parent=0 AND Active=1 Order By AccNo", con, adOpenForwardOnly, adLockReadOnly
              
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
    LItems.ListSubItems.add 2, , TBox(0).Text & ""
    
    If cmbType.Text = "Debit" Then
        LItems.ListSubItems.add(3, , Format(Val(TBox(1).Text), "#,##")).Tag = Val(TBox(1).Text)
        LItems.ListSubItems.add 4, , ""
    Else
        LItems.ListSubItems.add 3, , ""
        LItems.ListSubItems.add(4, , Format(Val(TBox(1).Text), "#,##")).Tag = Val(TBox(1).Text)
    End If
    
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
        DebitTotal = DebitTotal + Val(LView.ListItems(i).ListSubItems(3).Tag)
        CreditTotal = CreditTotal + Val(LView.ListItems(i).ListSubItems(4).Tag)
    Next i

    Label1(1).Caption = DebitTotal
    Label1(2).Caption = CreditTotal

    If DebitTotal > CreditTotal Then
        cmbType.ListIndex = 1
        TBox(1).Text = DebitTotal - CreditTotal
    Else
        cmbType.ListIndex = 0
        TBox(1).Text = CreditTotal - DebitTotal
    End If

End Sub


Private Function SaveTrans() As Boolean
    
    On Error GoTo err
    If LView.ListItems.count = 0 Then
        MsgBox "No Transaction to save!", vbInformation
        Exit Function
    End If
    If Val(Label1(1).Caption) <> Val(Label1(2).Caption) Then
        MsgBox "Debit & Credit totals are not equal.", vbInformation
        Exit Function
    ElseIf Val(Label1(1).Caption) = 0 Then
        MsgBox "Invalid Entry.", vbInformation
        Exit Function
    End If
    
    ''''''''''''''''''''''Check For Voucher No. Used Or Not '''''''''''''''''

    Dim VNo As String
    VNo = getNextVchrNo(DTPicker1.Value, "JV")
    If Label1(4) <> VNo Then
       MsgBox "Voucher No. : " & Label1(4) & " Has Been Used By Other Transaction." & vbNewLine & _
              "Voucher No. : " & VNo & " Will Be Used For This Transaction.", vbInformation
              Label1(4) = VNo
    End If

    Dim i As Integer
    Dim INextSno As Double
    Dim Bal As Double
    INextSno = getNextSno(DTPicker1.Value)
    Dim bWaitForAuth As Boolean
    bWaitForAuth = VoucherRequiresAuthorization("JV")
    Call StartTrans(con)
    For i = 1 To LView.ListItems.count
        If bWaitForAuth = False Then
            Bal = getBalance(LView.ListItems(i).Text, DTPicker1.Value)
            Bal = Bal + Val(LView.ListItems(i).ListSubItems(3).Tag) - Val(LView.ListItems(i).ListSubItems(4).Tag)
        
            con.Execute "INSERT INTO Vouchers VALUES(" & INextSno & ",'" & Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & _
             Label1(4) & "','" & LView.ListItems(i).Text & "','" & LView.ListItems(i).ListSubItems(2).Text & "'," & _
             Val(LView.ListItems(i).ListSubItems(3).Tag) & "," & Val(LView.ListItems(i).ListSubItems(4).Tag) & "," & _
             Bal & ",'',0)"
             
             Dim Amt As Double
            Amt = Val(LView.ListItems(i).ListSubItems(3).Tag) - Val(LView.ListItems(i).ListSubItems(4).Tag)
         
            Call adjustLedger(LView.ListItems(i).Text, DTPicker1, INextSno, Amt)
        Else
            con.Execute "INSERT INTO VouchersForAuthorization(VDate,VchrNo,Accno,Description,Debit,Credit,balance,DpstSlip,CSNo) VALUES('" & Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & _
             Label1(4) & "','" & LView.ListItems(i).Text & "','" & LView.ListItems(i).ListSubItems(2).Text & "'," & _
             Val(LView.ListItems(i).ListSubItems(3).Tag) & "," & Val(LView.ListItems(i).ListSubItems(4).Tag) & ",0,'',0)"
        End If
          
         INextSno = INextSno + 1
    Next i
    
    If strVchrNo = "" Then
        If bWaitForAuth = False Then
            con.Execute "INSERT INTO VoucherInfo(VchrNo,UserName,MachineName,DT,PostedThroughJVForm) VALUES('" & Label1(4) & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "',1)"
        Else
            con.Execute "INSERT INTO VouchersForAuthorizationInfo(VchrNo,UserName,MachineName,DT,PostedThroughJVForm) VALUES('" & Label1(4) & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "',1)"
        End If
    Else
        con.Execute "INSERT INTO VouchersEditingInfo(VchrNo,UserName,MachineName,DT) VALUES('" & Label1(4) & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
    End If
    con.CommitTrans
    'Print If To Be Printed
    
    If chkPrint = vbChecked Then
        Call PrintVoucher(MainForm.cr1, Label1(4), True, Not bVoucherViewOnlyWhenSaved, bWaitForAuth)
    End If
    
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
    If LView.ListItems.count > 0 Then
        With LView.SelectedItem
            If Val(.ListSubItems(3).Text) > 0 Then
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

Private Sub RefreshList(IsVenderAccount As Boolean)

    Dim QryName As String
    
    If rsList.State = 1 Then rsList.Close
    
    If IsVenderAccount Then QryName = "AccRecAdv" Else QryName = "AccPayAdv"
    rsList.CursorLocation = adUseClient
    rsList.Open "Select Sno,VchrNo,Description,TotalAmount,AmountLeft from " & QryName & " where (AccNo='" & cmbAccNo & "') and (VDate<='" & DTPicker1 & "') and ((AmountLeft<0) or (AmountLeft is null))", con, adOpenForwardOnly, adLockReadOnly
    rsList.Sort = "Sno"
    'Set DataGrid1.DataSource = rsList
End Sub

Private Sub ShowAccBalance()
'Show Previous Balance of Account
If lblBalance.Visible Then
    lblBalance.Caption = Format(getBalance(cmbAccNo, DTPicker1), "#,##0.00 DR;#,##0.00 CR;0.00 ")
    lblBalance.ToolTipText = "Balance of " & cmbAccNo & " as of " & Format(DTPicker1.Value, "dd/MM/yyyy")
End If
End Sub

Private Sub txtEdit_LostFocus()
txtEdit.Visible = False
End Sub
