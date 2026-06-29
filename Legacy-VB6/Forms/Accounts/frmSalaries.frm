VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSalaries 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Salaries"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11265
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
   ScaleWidth      =   11265
   WindowState     =   2  'Maximized
   Begin VB.Frame FALL 
      Height          =   6855
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   11220
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
         Height          =   660
         Left            =   5730
         TabIndex        =   27
         Top             =   6060
         Width           =   5280
         Begin MSForms.CommandButton cmdSClear 
            Height          =   360
            Left            =   75
            TabIndex        =   5
            Top             =   195
            Width           =   1665
            Caption         =   "Save & New   "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmSalaries.frx":0000
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
            Left            =   3495
            TabIndex        =   7
            Top             =   180
            Width           =   1665
            Caption         =   "Cancel            "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmSalaries.frx":0112
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
            Left            =   1785
            TabIndex        =   6
            Top             =   195
            Width           =   1665
            Caption         =   "Save & Close  "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmSalaries.frx":0464
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
      Begin VB.Frame Frame1 
         Height          =   645
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   11190
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Salaries"
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
            TabIndex        =   24
            Top             =   165
            Width           =   11085
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Salaries"
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
            TabIndex        =   25
            Top             =   180
            Width           =   11070
         End
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E3EDF9&
         CausesValidation=   0   'False
         Height          =   255
         Left            =   4380
         TabIndex        =   9
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
         TabIndex        =   2
         Top             =   5520
         Width           =   1635
      End
      Begin MSComctlLib.ListView LView 
         Height          =   4050
         Left            =   165
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1410
         Width           =   10980
         _ExtentX        =   19368
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Type"
            Object.Width           =   2302
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   5980
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Salary"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "O.Time"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Advance"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Loan"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Loan Ded."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Net Payable"
            Object.Width           =   1879
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   4485
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
         Format          =   56164355
         CurrentDate     =   37250
      End
      Begin VB.Frame fBankAcc 
         Height          =   675
         Left            =   165
         TabIndex        =   10
         Top             =   5385
         Visible         =   0   'False
         Width           =   5715
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
            Left            =   4125
            Picture         =   "frmSalaries.frx":05C2
            Style           =   1  'Graphical
            TabIndex        =   3
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
            Left            =   4455
            TabIndex        =   4
            Top             =   360
            Width           =   1200
         End
         Begin VB.Label lblBankAcc 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   60
            TabIndex        =   12
            Top             =   345
            Width           =   4065
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
            Left            =   4455
            TabIndex        =   11
            Top             =   150
            Width           =   1170
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   11
         Left            =   7170
         TabIndex        =   26
         Top             =   5745
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   8
         Left            =   5910
         TabIndex        =   22
         Top             =   5745
         Width           =   1260
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   6
         Left            =   8445
         TabIndex        =   21
         Top             =   5745
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   1
         Left            =   9720
         TabIndex        =   16
         Top             =   5745
         Width           =   1275
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Salary     Total O.Time    Total Adv's       Total Payable"
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
         Height          =   270
         Index           =   0
         Left            =   5910
         TabIndex        =   20
         Top             =   5490
         Width           =   5100
      End
      Begin VB.Label lblBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   9270
         TabIndex        =   19
         Top             =   810
         Width           =   1845
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
         Left            =   7845
         TabIndex        =   18
         Top             =   855
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Select The Employess To Pay Salary"
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
         TabIndex        =   17
         Top             =   1170
         Width           =   10935
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
         Left            =   3705
         TabIndex        =   14
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
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
Attribute VB_Name = "frmSalaries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PermissionDenied As Boolean
Private Sub chkBank_Click()


If chkBank.value = vbChecked Then

BankAccCheck:

   If Not AnyBankAccountDefined Then
      If MsgBox("Cannot Charge To Bank, No Bank Accounts Found" & vbNewLine & _
         "Do You Want To Define A Bank Account Now", vbInformation + vbYesNo) = vbYes Then
         With frmNewAccount
            .chkBankAcc = vbChecked
            .Show 1
         End With
         
         'Screen.MousePointer = vbHourglass
         'Sleep (800)
         'Screen.MousePointer = vbDefault
         GoTo BankAccCheck
         
      Else
         chkBank.value = vbUnchecked
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
   End If

End If


Dim BankID As Double
fBankAcc.Visible = chkBank.value = vbChecked
      'get Default Account
        lblBankAcc = GetDefaultBankAcc(BankID)
        lblBankAcc.Tag = BankID
      'Get ChqNo
      Dim ChqBookNo As Double
      txtChqNo = GetNextChqNo(GetID(lblBankAcc), ChqBookNo)
      txtChqNo.Tag = ChqBookNo
        Call ShowBals
Screen.MousePointer = vbDefault
 Exit Sub

BankErr:
   
   MsgBox "No Default Bank Account Found.", vbInformation
   chkBank.value = vbUnchecked
   Screen.MousePointer = vbDefault
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

Private Sub DtPicker1_Change()
    Label1(4).Caption = getNextVchrNo(DTPicker1.value, "PV")
    Call GetEmpList
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
   
PermissionDenied = Not UserHasAccess("SalariesVoucher")
   
   
    Me.KeyPreview = True
    DTPicker1.value = Date
    Call DtPicker1_Change
End Sub

Private Sub GetEmpList()
    
  
    Dim AccRec As New ADODB.Recordset
    
    Dim itm As ListItem
    
    
    With AccRec
    
         .Open "Select * From VEmpSal Where TotalSalary >0 and Active=1 Order By EmpType,EmpName", Con, adOpenForwardOnly, adLockReadOnly
         LView.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set itm = LView.ListItems.Add(, ![LAdvAccno] & "", ![EmpType] & "")
            itm.Tag = ![SAdvAccno]
            itm.ListSubItems.Add(1, , ![EmpName] & "").Tag = ![AcrAccNo] & ""
            itm.ListSubItems.Add , , Format(Val(![TotalSalary] & ""), "0.00")
            
            'Store Over Time And OverTime Acc
            itm.ListSubItems.Add(, , Format(Val(![TotalOverTime] & ""), "0.00")).Tag = ![AcrOTAcc] & ""
            
            itm.SubItems(4) = Format(Val(![TotalSAdv] & ""), "0.00")
            itm.SubItems(5) = Format(Val(![TotalLAdv] & ""), "0.00")
            itm.SubItems(6) = "0.00"
            itm.SubItems(7) = Format(Val(![TotalSalary] & "") + Val(![TotalOverTime] & "") - Val(![TotalSAdv] & ""), "0.00")
           .MoveNext
        Next
        .Close
   End With
    Set AccRec = Nothing
End Sub



Private Function GetTotal(Optional TotalOf As String) As Double

Dim i As Integer
Dim DebitTotal As Double
Dim CreditTotal As Double
Dim Amt As Double


 
For i = 1 To LView.ListItems.Count
   If TotalOf = "ADV" Then
      Amt = Val(LView.ListItems(i).SubItems(4)) + Val(LView.ListItems(i).SubItems(6))
   ElseIf TotalOf = "SAL" Then
      Amt = Val(LView.ListItems(i).SubItems(2))
   ElseIf TotalOf = "OTM" Then
      Amt = Val(LView.ListItems(i).SubItems(3))
   Else
      Amt = Val(LView.ListItems(i).SubItems(7))
   End If
    If LView.ListItems(i).Checked Then DebitTotal = DebitTotal + Amt
Next i

GetTotal = DebitTotal

End Function


Private Function SaveTrans() As Boolean
On Error GoTo ERR

SaveTrans = False
Dim TheDesc As String

 
 TheDesc = "Employee Salaries Paid For " & Format(DTPicker1, "MMMM ,yyyy")
 
 
If LView.ListItems.Count = 0 Then
    MsgBox "No Transaction to save!", vbInformation
    Exit Function
ElseIf Val(Label1(1).Caption) = 0 Then
   MsgBox "Invalid Entry.", vbInformation
   Exit Function
End If


If chkBank.value = vbChecked Then
   If Val(Label1(1)) > Val(lblBal) Then
      
         MsgBox "Not Much Amount In " & Left(lblBalCap, Len(lblBalCap) - 2) & ".", vbInformation
         Exit Function
      
   End If
ElseIf chkBank.value = vbUnchecked Then
   If Val(Label1(8)) > Val(lblBal) Then
      If Not AllowOD(GetID(lblBankAcc)) Then
         MsgBox "Not Much Amount In " & Left(lblBalCap, Len(lblBalCap) - 2) & ".", vbInformation
         Exit Function
      End If
   End If
End If
    
    'Validate ChqNo
 If chkBank.value = vbChecked Then
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
 
 
 
 
 
 
 ''''''''''''''''''''''Check For Voucher No. Used Or Not '''''''''''''''''

Dim VNo As String
VNo = getNextVchrNo(DTPicker1.value, "PV")
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
      
      INextSno = getNextSno(DTPicker1.value, "Cheque")
      SlipNo = INextSno
     Con.BeginTrans
     
      
      
   'Check Whether Expense is Charged To Bank Or To Cash in Hand
   
   If chkBank.value = vbChecked Then
       CreditAcc = GetID(lblBankAcc)
      
      Con.Execute "insert into Cheque Values(" & INextSno & ",'" & Format(DTPicker1.value, "dd-MMM-yyyy") & "','" & _
      CreditAcc & "','" & TheDesc & "'," & Val(Label1(1)) & "," & lblBankAcc.Tag & ",'" & _
        txtChqNo & "'," & txtChqNo.Tag & ",'Open','" & DTPicker1 & "',1,1,0)"
       
      Call IssueChq(Val(txtChqNo.Tag), Val(txtChqNo))

   Else
      CreditAcc = CashInHand

   End If
    
    
    
    ' First Clear Advances then Pay Salary
    
    
    If Not AddAdvanceVchr Then
         Con.RollbackTrans
         SaveTrans = False
         Exit Function
    End If
    
    
    INextSno = getNextSno(DTPicker1.value)
    
    
    Dim Salary As Double, Advance As Double, Loan As Double
    
    
    
    For i = 1 To LView.ListItems.Count
      'If Not Paid Then Go To Next One
      If Not LView.ListItems(i).Checked Then GoTo EndOfLoop
      
         '''''''''''  Accrued Salary Entry  ''''''''''''
         
            Salary = Val(LView.ListItems(i).ListSubItems(2).Text)
         
            Bal = getBalance(LView.ListItems(i).ListSubItems(1).Tag, DTPicker1.value)
            Bal = Bal + Salary
         
            Con.Execute "insert into Vouchers Values(" & INextSno & ",'" & Format(DTPicker1.value, "dd-MMM-yyyy") & "','" & _
            Label1(4) & "','" & LView.ListItems(i).ListSubItems(1).Tag & "','" & "Salary Paid For " & Format(DTPicker1, "MMMM ,yyyy") & "'," & _
            Salary & ",0," & _
            Bal & ",''," & SlipNo & ")"
         
            Call adjustLedger(LView.ListItems(i).ListSubItems(1).Tag, DTPicker1.value, INextSno, Salary)
            INextSno = INextSno + 1
         
         '''''''''''  Accrued OverTime Entry  ''''''''''''
         
            Dim OverTime As Double, AcrOverTimeAcc As String
            
            OverTime = Val(LView.ListItems(i).SubItems(3))
            ' if Some Over Time Enter it in Transaction
            If OverTime = 0 Then GoTo EndOfLoop
            
            AcrOverTimeAcc = LView.ListItems(i).ListSubItems(3).Tag
            
            If OverTime = 0 Then GoTo EndOfLoop
            
            Bal = getBalance(AcrOverTimeAcc, DTPicker1.value)
            Bal = Bal + OverTime
         
            Con.Execute "insert into Vouchers Values(" & INextSno & ",'" & Format(DTPicker1.value, "dd-MMM-yyyy") & "','" & _
            Label1(4) & "','" & AcrOverTimeAcc & "','" & "Over Time Paid For " & Format(DTPicker1, "MMMM ,yyyy") & "'," & _
            OverTime & ",0," & Bal & ",''," & SlipNo & ")"
         
            Call adjustLedger(AcrOverTimeAcc, DTPicker1.value, INextSno, OverTime)
            INextSno = INextSno + 1
            
            
EndOfLoop:
    Next i
    
    
    ''''''''''''''''Charge To  Cash Entry'''''''''''''
    
    INextSno = INextSno + 1
    Dim SalTotal As Double
    Bal = getBalance(CreditAcc, DTPicker1.value)
    SalTotal = GetTotal("SAL") + GetTotal("OTM")
    Bal = Bal - SalTotal
    'Dim TheDesc As String
    
    
    
    
    'Credit Entry
    Con.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Format(DTPicker1.value, "dd-MMM-yyyy") & "','" & Label1(4) & "','" & CreditAcc & _
        "','" & TheDesc & "',0," & SalTotal & "," & Bal & ",''," & SlipNo & ")"

    Call adjustLedger(CreditAcc, DTPicker1.value, INextSno, -SalTotal)
   
    Con.CommitTrans
      SaveTrans = True
    Exit Function
ERR:
   MsgBox "Error:" & ERR.Number & vbNewLine & ERR.Description
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
          If Val(.SubItems(4)) = 0 Then
             MsgBox "There is No Loan For This Employee, Deduction Cannot Be Made.", vbInformation
             Exit Sub
          End If
            
            txtEdit.Move LView.ColumnHeaders(6).Left + LView.Left + 30, LView.Top + .Top + 25, LView.ColumnHeaders(6).Width, .Height
            txtEdit.Text = .ListSubItems(5).Text
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
If Val(ITEM.SubItems(2)) <= 0 And ITEM.Checked Then
   ITEM.Checked = False
   MsgBox "This Employee Has No Salary Payable", vbInformation
Else
   Label1(1).Caption = Format(GetTotal, "0.00")
   Label1(6).Caption = Format(GetTotal("ADV"), "0.00")
   Label1(8).Caption = Format(GetTotal("SAL"), "0.00")
   Label1(11).Caption = Format(GetTotal("OTM"), "0.00")
End If
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
         If Val(txtEdit) > .SubItems(5) Then
            MsgBox "Deduction Cannot Be Greater Than Loan", vbInformation
            Exit Sub
         End If
         
         .ListSubItems(5).Text = Format(Val(txtEdit), "0.00")
         .SubItems(6) = Val(.SubItems(2)) + Val(.SubItems(3)) - Val(.SubItems(4)) - Val(.SubItems(6))
         .SubItems(6) = Format(.SubItems(6), "0.00")
         Label1(1).Caption = Format(GetTotal, "0.00")
         Label1(6).Caption = Format(GetTotal("ADV"), "0.00")
         Label1(8).Caption = Format(GetTotal("SAL"), "0.00")
         Label1(11).Caption = Format(GetTotal("OTM"), "0.00")
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




Private Function AddAdvanceVchr() As Boolean
On Error GoTo ERR

AddAdvanceVchr = False
Dim TheDesc As String

 
 TheDesc = "Advance Adjustment For " & Format(DTPicker1, "MMMM ,yyyy")
 
 

    
    Dim i As Integer
    Dim INextSno As Double
    Dim Bal As Double
    Dim VchrNo As String
      

    ' Con.BeginTrans
    VchrNo = getNextVchrNo(DTPicker1, "RV")
    INextSno = getNextSno(DTPicker1.value)
   
    
   Dim Advance As Double
    
    For i = 1 To LView.ListItems.Count
      'If Not Paid Then Go To Next One
      If (Not LView.ListItems(i).Checked) Or Val(LView.ListItems(i).SubItems(4)) = 0 Then GoTo EndOfLoop
      
      ''''''''  Short Term Advances '''''''''
         
            Advance = Val(LView.ListItems(i).ListSubItems(3).Text)
            If Advance > 0 Then
               Bal = getBalance(LView.ListItems(i).Tag, DTPicker1.value)
               Bal = Bal - Advance
         
               Con.Execute "insert into Vouchers Values(" & INextSno & ",'" & Format(DTPicker1.value, "dd-MMM-yyyy") & "','" & _
               Label1(4) & "','" & LView.ListItems(i).Tag & "','" & "Advances Cleared For " & Format(DTPicker1, "MMMM ,yyyy") & "',0," & _
               Advance & "," & _
               Bal & ",'',0)"

               Call adjustLedger(LView.ListItems(i).Tag, DTPicker1.value, INextSno, -Advance)
               INextSno = INextSno + 1
            End If
            
            ''''''''  Long Term Advances (Loan) '''''''''
         
            Loan = Val(LView.ListItems(i).ListSubItems(6).Text)
            If Loan > 0 Then
               Bal = getBalance(LView.ListItems(i).key, DTPicker1.value)
               Bal = Bal - Loan
         
               Con.Execute "insert into Vouchers Values(" & INextSno & ",'" & Format(DTPicker1.value, "dd-MMM-yyyy") & "','" & _
               Label1(4) & "','" & LView.ListItems(i).key & "','" & "Loan Deduction For " & Format(DTPicker1, "MMMM ,yyyy") & "',0," & _
               Loan & "," & _
               Bal & ",'',0)"

               Call adjustLedger(LView.ListItems(i).key, DTPicker1.value, INextSno, -Loan)
               INextSno = INextSno + 1
            End If
         
         
EndOfLoop:
      Next i
    
    
    ''''''''''''''''Charge To  Cash Entry'''''''''''''
    
    INextSno = INextSno + 1
    
    Bal = getBalance(CashInHand, DTPicker1.value)
    
    'Dim TheDesc As String
    Dim AdvTotal As Double
    AdvTotal = GetTotal("ADV")
    Bal = Bal + AdvTotal
    If AdvTotal <> 0 Then
         'Debit Entry
         Con.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
            Format(DTPicker1.value, "dd-MMM-yyyy") & "','" & VchrNo & "','" & CashInHand & _
            "','" & TheDesc & "'," & AdvTotal & ",0," & Bal & ",'',0)"

         Call adjustLedger(CashInHand, DTPicker1.value, INextSno, AdvTotal)
    End If
    
    AddAdvanceVchr = True
    Exit Function
    
ERR:
   MsgBox "Error:" & ERR.Number & vbNewLine & ERR.Description
    AddAdvanceVchr = False
End Function



Private Sub ShowBals()
If chkBank.value = vbChecked Then
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
