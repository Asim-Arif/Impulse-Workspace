VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAcrExpenses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accured Expenses"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
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
   ScaleHeight     =   6945
   ScaleWidth      =   10470
   WindowState     =   2  'Maximized
   Begin VB.Frame FALL 
      Height          =   6855
      Left            =   60
      TabIndex        =   9
      Top             =   60
      Width           =   10365
      Begin VB.Frame Frame1 
         Height          =   645
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   10365
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Accured Expenses"
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
            Index           =   9
            Left            =   75
            TabIndex        =   22
            Top             =   165
            Width           =   10290
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Accured Expenses"
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
            Index           =   10
            Left            =   90
            TabIndex        =   23
            Top             =   180
            Width           =   10275
         End
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E3EDF9&
         CausesValidation=   0   'False
         Height          =   255
         Left            =   4380
         TabIndex        =   1
         Top             =   2880
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
         Left            =   210
         TabIndex        =   3
         Top             =   5520
         Width           =   1635
      End
      Begin MSComctlLib.ListView LView 
         Height          =   4050
         Left            =   165
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1410
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   7144
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Acc #"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Account Title"
            Object.Width           =   7408
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Accured Amount"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amt To Clear"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   4065
         TabIndex        =   0
         Top             =   810
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
         Format          =   48168963
         CurrentDate     =   37250
      End
      Begin VB.Frame fBankAcc 
         Height          =   675
         Left            =   150
         TabIndex        =   10
         Top             =   5385
         Visible         =   0   'False
         Width           =   6570
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
            Left            =   4620
            Picture         =   "frmAcrExpenses.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   345
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
            Left            =   4950
            TabIndex        =   5
            Top             =   360
            Width           =   1560
         End
         Begin VB.Label lblBankAcc 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   60
            TabIndex        =   12
            Top             =   345
            Width           =   4530
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
            Left            =   4995
            TabIndex        =   11
            Top             =   150
            Width           =   1170
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
         Height          =   660
         Left            =   4980
         TabIndex        =   17
         Top             =   6045
         Width           =   5280
         Begin MSForms.CommandButton cmdSClear 
            Height          =   360
            Left            =   75
            TabIndex        =   6
            Top             =   195
            Width           =   1665
            Caption         =   "Save & New   "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmAcrExpenses.frx":0372
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
            Left            =   3510
            TabIndex        =   8
            Top             =   195
            Width           =   1665
            Caption         =   "Cancel            "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmAcrExpenses.frx":0484
            Accelerator     =   67
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
            Left            =   1800
            TabIndex        =   7
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
      End
      Begin VB.Label Label2 
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
         Left            =   7950
         TabIndex        =   24
         Top             =   5565
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   1
         Left            =   8670
         TabIndex        =   16
         Top             =   5520
         Width           =   1455
      End
      Begin VB.Label lblBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   8700
         TabIndex        =   20
         Top             =   810
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
         Left            =   7290
         TabIndex        =   19
         Top             =   855
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00BECCD8&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Accured Expenses To Clear"
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
         Index           =   7
         Left            =   195
         TabIndex        =   18
         Top             =   1170
         Width           =   10050
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
         TabIndex        =   15
         Top             =   855
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
         Left            =   3285
         TabIndex        =   14
         Top             =   840
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
         TabIndex        =   13
         Top             =   810
         Width           =   1560
      End
   End
End
Attribute VB_Name = "frmAcrExpenses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PermissionDenied As Boolean
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
    Call GetExpenseList
    Call ShowBals
End Sub

Private Sub DTPicker1_Click()
    'Label1(4).Caption = getNextVchrNo(DTPicker1.value, "PV")
    'Call GetEmpList
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

PermissionDenied = Not UserHasAccess("AccuredPrepaidVoucher")
  
   
    Me.KeyPreview = True
    DTPicker1.Value = Date
    Call DTPicker1_Change
End Sub

Private Sub GetExpenseList()
    
  
    Dim AccRec As New ADODB.Recordset
    
    Dim itm As ListItem
    
    
    With AccRec
    
         .Open "Select * From Accounts Where Accno in (Select AcrExpAccno From ExpenseAccounts Where Prepaid=0) and Balance<>0", con, adOpenForwardOnly, adLockReadOnly
         LView.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set itm = LView.ListItems.Add(, , ![AccNo])
            itm.SubItems(1) = ![AccTitle] & ""
            itm.SubItems(2) = Format(Abs(Val(![Balance] & "")), "0.00")
            itm.SubItems(3) = Format(Abs(Val(![Balance] & "")), "0.00")
           .MoveNext
        Next
        .Close
   End With
    Set AccRec = Nothing
End Sub



Private Function GetTotal() As Double

Dim i As Integer
Dim DebitTotal As Double
Dim CreditTotal As Double
Dim Amt As Double


 
For i = 1 To LView.ListItems.Count
   
      Amt = Val(LView.ListItems(i).SubItems(3))
   
    If LView.ListItems(i).Checked Then DebitTotal = DebitTotal + Amt
Next i

GetTotal = DebitTotal

End Function


Private Function SaveTrans() As Boolean
On Error GoTo err

SaveTrans = False
Dim TheDesc As String

     For i = 1 To LView.ListItems.Count
        TheDesc = TheDesc & IIf((i = LView.ListItems.Count) And (i > 1), " and " & LView.ListItems(i).SubItems(1), LView.ListItems(i).SubItems(1) & ",")
     Next
     
      
      
 TheDesc = TheDesc & " Cleared For " & Format(DTPicker1, "MMMM ,yyyy")
 
 
If LView.ListItems.Count = 0 Then
    MsgBox "No Transaction to save!", vbInformation
    Exit Function
ElseIf Val(Label1(1).Caption) = 0 Then
   MsgBox "Invalid Entry.", vbInformation
   Exit Function
End If



   If Val(Label1(1)) > Val(lblBal) Then
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
          MsgBox "Invalid Cheque No.", vbInformation
      End If
      txtChqNo.SetFocus
      Exit Function
    End If
    txtChqNo.Tag = ChqBookNo
 End If
 
 
 
    'Dim i As Integer
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
      CreditAcc & "','" & TheDesc & "'," & Val(Label1(1)) & "," & lblBankAcc.Tag & ",'" & _
        txtChqNo & "'," & txtChqNo.Tag & ",'Open','" & DTPicker1 & "',1,0,0)"
       
      Call IssueChq(Val(txtChqNo.Tag), Val(txtChqNo))

   Else
      CreditAcc = CashInHand

   End If
    
    
    
    
    
    
       
    
    INextSno = getNextSno(DTPicker1.Value)
    
    
    Dim ExpAmt As Double, Advance As Double, Loan As Double
    
    
    
    For i = 1 To LView.ListItems.Count
      'If Not Paid Then Go To Next One
      If Not LView.ListItems(i).Checked Then GoTo EndOfLoop
      
         '''''''''''  Accrued Expense Account Entry  ''''''''''''
         
            ExpAmt = Val(LView.ListItems(i).ListSubItems(3).Text)
         
            Bal = getBalance(LView.ListItems(i), DTPicker1.Value)
            Bal = Bal + ExpAmt
         
            con.Execute "insert into Vouchers Values(" & INextSno & ",'" & Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & _
            Label1(4) & "','" & LView.ListItems(i) & "','" & "Accured Expense Cleared For " & Format(DTPicker1, "MMMM ,yyyy") & "'," & _
            ExpAmt & ",0," & _
            Bal & ",''," & SlipNo & ")"
         
            Call adjustLedger(LView.ListItems(i), DTPicker1.Value, INextSno, ExpAmt)
            INextSno = INextSno + 1
         
         
            
            
EndOfLoop:
    Next i
    
    
    ''''''''''''''''Charge To  Cash Entry'''''''''''''
    
    INextSno = INextSno + 1
    Dim SalTotal As Double
    Bal = getBalance(CreditAcc, DTPicker1.Value)
    SalTotal = GetTotal()
    Bal = Bal - SalTotal

    
    'Credit Entry
    con.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & Label1(4) & "','" & CreditAcc & _
        "','" & TheDesc & "',0," & SalTotal & "," & Bal & ",''," & SlipNo & ")"

    Call adjustLedger(CreditAcc, DTPicker1.Value, INextSno, -SalTotal)
   
    con.CommitTrans
      SaveTrans = True
    Exit Function
err:
   MsgBox "Error:" & err.Number & vbNewLine & err.Description
   SaveTrans = False

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

Private Sub LView_DblClick()
If LView.ListItems.Count > 0 Then
   With LView.SelectedItem
      If Not .Checked Then Exit Sub

            
            txtEdit.Move LView.ColumnHeaders(4).Left + LView.Left + 30, LView.Top + .Top + 25, LView.ColumnHeaders(4).Width, .Height
            txtEdit.Text = .ListSubItems(3).Text
            txtEdit.Tag = 5
   End With
   LView.Tag = LView.SelectedItem.Index
   txtEdit.Visible = True
   txtEdit.SelStart = 0
   txtEdit.SelLength = Len(txtEdit)
   txtEdit.SetFocus
End If
End Sub

Private Sub LView_ItemCheck(ByVal ITEM As MSComctlLib.ListItem)
   Label1(1).Caption = Format(GetTotal, "0.00")
End Sub


Private Sub LView_Validate(Cancel As Boolean)
On Error Resume Next
'    cmdDelete.Visible = False
    txtEdit.Visible = False
End Sub


Private Sub ClearLView()

Dim iTotal As Integer
iTotal = LView.ListItems.Count
    For i = 1 To iTotal
        LView.ListItems.Remove (1)
    Next i
'Tbox(0).Text = ""
'Tbox(1).Text = ""
'cmdDelete.Visible = False
Call DTPicker1_Click
End Sub




Private Sub txtEdit_KeyPress(KeyAscii As Integer)

With LView.ListItems(Val(LView.Tag))
    If KeyAscii = 13 Then
         If Val(txtEdit) > .SubItems(2) Then
            MsgBox "Amount Cannot Be Greater Than Total Balance.", vbInformation
            Exit Sub
         End If
         
         .ListSubItems(3).Text = Format(Val(txtEdit), "0.00")
   Label1(1).Caption = Format(GetTotal, "0.00")
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    Else
        KeyAscii = OnlyNumber(KeyAscii, True)
    End If
End With
End Sub



Private Sub Form_Resize()
   fAll.Move (Me.ScaleWidth - fAll.Width) / 2, (Me.ScaleHeight - fAll.Height) / 2
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

Private Sub txtEdit_LostFocus()
   txtEdit.Visible = False
End Sub
