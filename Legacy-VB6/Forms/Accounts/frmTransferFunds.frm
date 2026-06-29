VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransferFunds 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transfer Funds"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   105
      TabIndex        =   26
      Top             =   4215
      Value           =   1  'Checked
      Width           =   1800
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   3615
      TabIndex        =   23
      Top             =   90
      Width           =   2445
      _ExtentX        =   4313
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
      CustomFormat    =   "dddd dd/MM/yyyy"
      Format          =   52494339
      CurrentDate     =   37378
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
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
      Left            =   4590
      TabIndex        =   14
      Top             =   2115
      Width           =   1350
   End
   Begin VB.Frame fFrom 
      Caption         =   "Transfer Amount From :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1530
      Left            =   105
      TabIndex        =   0
      Top             =   495
      Width           =   5940
      Begin VB.Frame fBankAcc 
         Enabled         =   0   'False
         Height          =   660
         Left            =   60
         TabIndex        =   6
         Top             =   810
         Width           =   5805
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
            Left            =   4050
            Picture         =   "frmTransferFunds.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   315
            Width           =   360
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
            Left            =   4395
            TabIndex        =   7
            Top             =   330
            Width           =   1350
         End
         Begin VB.Label lblFromBankAcc 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   60
            TabIndex        =   9
            Top             =   330
            Width           =   4005
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Bank Account                                                                       Cheque #               "
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
            Height          =   225
            Index           =   2
            Left            =   60
            TabIndex        =   8
            Top             =   105
            Width           =   5700
         End
      End
      Begin VB.OptionButton OptFromBank 
         Caption         =   "Bank Account"
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
         Left            =   120
         TabIndex        =   3
         Top             =   615
         Width           =   1650
      End
      Begin VB.OptionButton OptFromCash 
         Caption         =   "Cash In Hand"
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
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   1650
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Account Balance :"
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
         Left            =   2400
         TabIndex        =   11
         Top             =   405
         Width           =   1665
      End
      Begin VB.Label lblFromBal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4200
         TabIndex        =   10
         Top             =   375
         Width           =   1635
      End
   End
   Begin VB.Frame fTo 
      Caption         =   "Transfer Amount To :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1530
      Left            =   105
      TabIndex        =   1
      Top             =   2490
      Width           =   5940
      Begin VB.Frame Frame3 
         Enabled         =   0   'False
         Height          =   660
         Left            =   60
         TabIndex        =   16
         Top             =   810
         Width           =   5805
         Begin VB.CommandButton cmdToBank 
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
            Left            =   4050
            Picture         =   "frmTransferFunds.frx":0372
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   330
            Width           =   360
         End
         Begin VB.TextBox txtSlipNo 
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
            Left            =   4410
            TabIndex        =   19
            Top             =   330
            Width           =   1350
         End
         Begin VB.Label lbltoBankAcc 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   60
            TabIndex        =   18
            Top             =   330
            Width           =   3990
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Bank Account                                                                       Deposit Slip #"
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
            Height          =   225
            Index           =   0
            Left            =   75
            TabIndex        =   17
            Top             =   105
            Width           =   5700
         End
      End
      Begin VB.OptionButton OptToBank 
         Caption         =   "Bank Account"
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
         Left            =   120
         TabIndex        =   5
         Top             =   630
         Width           =   1485
      End
      Begin VB.OptionButton OptToCash 
         Caption         =   "Cash In Hand"
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
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   1650
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Account Balance :"
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
         Left            =   2460
         TabIndex        =   13
         Top             =   390
         Width           =   1665
      End
      Begin VB.Label lblToBal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4185
         TabIndex        =   12
         Top             =   390
         Width           =   1635
      End
   End
   Begin MSForms.CommandButton cmdTransferNew 
      Height          =   360
      Left            =   1950
      TabIndex        =   27
      Top             =   4125
      Width           =   1485
      Caption         =   "Transfer & New"
      PicturePosition =   327683
      Size            =   "2619;635"
      Accelerator     =   84
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
      Left            =   4800
      TabIndex        =   25
      Top             =   4125
      Width           =   1290
      Caption         =   "Cancel       "
      PicturePosition =   327683
      Size            =   "2275;635"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Transfer Date :"
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
      Left            =   1905
      TabIndex        =   24
      Top             =   150
      Width           =   1665
   End
   Begin MSForms.CommandButton cmdOK 
      Height          =   360
      Left            =   3465
      TabIndex        =   20
      Top             =   4125
      Width           =   1290
      Caption         =   "Transfer     "
      PicturePosition =   327683
      Size            =   "2275;635"
      Accelerator     =   84
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Amount To Transfer : "
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
      Left            =   2610
      TabIndex        =   15
      Top             =   2145
      Width           =   1935
   End
End
Attribute VB_Name = "frmTransferFunds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PermissionDenied As Boolean

Private Sub cmdBank_Click()

   Dim BankID As Double
   lblFromBankAcc = SelectBankAcc(BankID)
   lblFromBankAcc.Tag = BankID
   
      'Get ChqNo
      Dim ChqBookNo As Double
   
      txtChqNo = GetNextChqNo(GetID(lblFromBankAcc), ChqBookNo)
      txtChqNo.Tag = ChqBookNo
   
   fFrom.Tag = GetID(lblFromBankAcc)
   Call getFromBal
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdClose_Click()

End Sub

Private Sub cmdOK_Click()
    Call Save
End Sub

Private Sub cmdToBank_Click()
   Dim BankID As Double
   lbltoBankAcc = SelectBankAcc(BankID)
   lbltoBankAcc.Tag = BankID
   
   fTo.Tag = GetID(lbltoBankAcc)
   Call getToBal
End Sub



Private Sub cmdTransferNew_Click()
    Call Save(True)
End Sub

Private Sub DTPicker1_Change()
   Call setDateRange(DTPicker1)
   'Refresh Balances
   Call getFromBal
   Call getToBal
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

    'PermissionDenied = Not UserHasAccess("TransferFunds")
    DTPicker1 = Date
    If CurrentUserName = "Administrator" Then
        DTPicker1.Enabled = True
    Else
        DTPicker1.Enabled = False
    End If
    Me.KeyPreview = True
    'chkPrint = IIf(frmMain.mnuPrintOnSave.Checked, vbChecked, vbUnchecked)
    
End Sub

Private Sub OptFromBank_Click()

    If Not OptFromBank Then Exit Sub

    If Not AnyBankAccountDefined Then
        MsgBox "Funds Cannot Be Transfered From Bank, No Bank Accounts Defined.", vbInformation
        OptFromBank.Value = False
        Exit Sub
    End If

    fFrom.Tag = GetID(lblFromBankAcc)
    Call getFromBal

End Sub

Private Sub OptFromCash_Click()
If Not OptFromCash Then Exit Sub
   fFrom.Tag = CashInHand
   Call getFromBal
End Sub

Private Sub OptToBank_Click()
   If Not OptToBank Then Exit Sub
   
   If Not AnyBankAccountDefined Then
      MsgBox "Funds Cannot Be Transfered To Bank, No Bank Accounts Defined.", vbInformation
      OptToBank.Value = False
      Exit Sub
   End If
   
   fTo.Tag = GetID(lbltoBankAcc)
   Call getToBal
End Sub


Private Sub getFromBal()
fBankAcc.Enabled = OptFromBank

If fFrom.Tag = "" Then
   lblFromBal = 0
   Exit Sub
End If


   If fFrom.Tag <> "" Then
      lblFromBal.Tag = getBalance(fFrom.Tag, DTPicker1)
      lblFromBal = IIf(Val(lblFromBal.Tag) > 0, Format(Val(Abs(lblFromBal.Tag)), "#,##0.00" & " DR"), Format(Val(Abs(lblFromBal.Tag)), "#,##0.00" & " CR"))
   End If
   
End Sub
Private Sub getToBal()

Frame3.Enabled = OptToBank
If fTo.Tag = "" Then
   lblToBal = 0
   Exit Sub
End If

      lblToBal.Tag = getBalance(fTo.Tag, DTPicker1)
      lblToBal = IIf(Val(lblToBal.Tag) > 0, Format(Val(Abs(lblToBal.Tag)), "#,##0.00" & " DR"), Format(Val(Abs(lblToBal.Tag)), "#,##0.00" & " CR"))
   
End Sub

Private Sub OptToCash_Click()

If Not OptToCash Then Exit Sub
   
   fTo.Tag = CashInHand
   Call getToBal
   
End Sub

Private Function FundsTransfered() As Boolean

    On Error GoTo err

    Dim INextSno As Double, VchrNo As String, DescFrom As String, DescTo As String
   
    con.BeginTrans
      
    VchrNo = getNextVchrNo(DTPicker1, IIf(OptFromCash, "CPV", "BPV"))
                  
  '''''''''Cheque Entry
        Dim CSNo As Double
        
    Dim SlipNo As String, ChqNo As String
      If OptToBank Then
         DescTo = Replace(lbltoBankAcc, "[" & GetID(lbltoBankAcc) & "]", "")
         SlipNo = txtSlipNo
      Else
         DescTo = "Cash In Hand"
         SlipNo = ""
      End If
      
      
      
      If OptFromBank Then
         
         DescFrom = Replace(lblFromBankAcc, "[" & GetID(lblFromBankAcc) & "]", "") & " Chq # :" & txtChqNo
         ChqNo = txtChqNo
      Else
         DescFrom = "Cash In Hand"
         ChqNo = ""
      End If
        
        
        
         If OptFromBank Then
            INextSno = getNextSno(DTPicker1, "Cheque")
            CSNo = INextSno
            con.Execute "Insert Into Cheque ( SNo, CDate, AccNo, Description, Amount, BankID, chequeno, ChqBookNo, chequeType, chequeDate, Posted, Payment, Bounced,BankAccNo) Values(" & INextSno & ",'" & DTPicker1 & "','" & _
            fFrom.Tag & "','Amount Transfered To " & DescTo & "'," & Val(txtAmt) & "," & lblFromBankAcc.Tag & ",'" & _
            txtChqNo & "'," & txtChqNo.Tag & ",'Open','" & DTPicker1 & "',1,1,0,'" & GetID(lblFromBankAcc) & "')"
       
            Call IssueChq(Val(txtChqNo.Tag), txtChqNo)
         End If
         
         
         
         
      Dim strChqNo As String
      strChqNo = " Chq #: " & txtChqNo
      
   '''''Credit Entry
   
   INextSno = getNextSno(DTPicker1)
   
   
      Bal = Val(lblFromBal.Tag)
      Bal = Bal - Val(txtAmt)
      

        
        con.Execute "insert into Vouchers Values(" & INextSno & ",'" & DTPicker1 & "','" & _
        VchrNo & "','" & fFrom.Tag & "','Amount Transfered To " & DescTo & strChqNo & "', 0," & Val(txtAmt) & "," & Bal & ",''," & CSNo & ")"
        
        
      
      Call adjustLedger(fFrom.Tag, DTPicker1, INextSno, -Val(txtAmt))
      
      ''''''''Debit Entry
      INextSno = INextSno + 1
      Bal = Val(lblToBal.Tag)
      Bal = Bal + Val(txtAmt)

        
        con.Execute "insert into Vouchers Values(" & INextSno & ",'" & DTPicker1 & "','" & _
         VchrNo & "','" & fTo.Tag & "','Amount Transfered From " & DescFrom & strChqNo & "', " & Val(txtAmt) & ",0," & Bal & ",'" & SlipNo & "'," & CSNo & ")"
        
        Call adjustLedger(fTo.Tag, DTPicker1, INextSno, Val(txtAmt))
        
        con.Execute "Insert Into VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & VchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
        
    con.CommitTrans
 
 'Print If To Be Printed
        If chkPrint = vbChecked Then Call PrintVoucher(MainForm.cr1, VchrNo, True, True)
         
 FundsTransfered = True
 
 Exit Function

err:
   MsgBox "ERROR:" & err.Number & vbNewLine & err.Description
      con.RollbackTrans
      FundsTransfered = False
End Function

Private Sub txtAmt_GotFocus()
   txtAmt.SelStart = 0
   txtAmt.SelLength = Len(txtAmt)
End Sub

Private Sub txtAmt_KeyPress(KeyAscii As Integer)
   KeyAscii = OnlyNumber(KeyAscii, True)
End Sub

Private Sub txtAmt_LostFocus()
txtAmt = Format(Val(txtAmt), "###0.00")
End Sub

Private Sub txtChqNo_KeyPress(KeyAscii As Integer)
   KeyAscii = OnlyNumber(KeyAscii)
End Sub

Private Sub txtSlipNo_KeyPress(KeyAscii As Integer)
KeyAscii = OnlyNumber(KeyAscii)
End Sub


Private Sub Save(Optional ShowAgain As Boolean)

    If fFrom.Tag = "" Then
        MsgBox "Select The Account From Which The Amount is To Tranfer.", vbInformation
        Exit Sub
    ElseIf fTo.Tag = "" Then
        MsgBox "Select The Account To Which The Amount is To Tranfer.", vbInformation
        Exit Sub
    ElseIf fFrom.Tag = fTo.Tag Then
        MsgBox "The Amount Cannot Be Transfered To Same Account.", vbInformation
        Exit Sub
    ElseIf Val(txtAmt) = 0 Then
        MsgBox "No Amount Given To Transfer", vbInformation
        txtAmt.SetFocus
        Exit Sub
    ElseIf Val(txtAmt) > Val(lblFromBal.Tag) Then
        If Not AllowOD(GetID(lblFromBankAcc)) Then
            MsgBox "Not Enough Amount To Transfer.", vbInformation
            txtAmt.SetFocus
            Exit Sub
        End If
    ElseIf OptToBank And Trim(txtSlipNo) = "" Then
        MsgBox "No Deposit Slip No. Given", vbInformation
        txtSlipNo.SetFocus
        Exit Sub
    End If

    'Validate chq no
    If OptFromBank Then
        Dim Issued As Boolean, ChqBookNo As Double
        If Not ValidChqNo(fFrom.Tag, txtChqNo, Issued, ChqBookNo) Then
            If Issued Then
                MsgBox "The Cheque With Number '" & txtChqNo & "' Has Been Issued.", vbInformation
            Else
                MsgBox "Invalid Cheque No", vbInformation
            End If
            txtChqNo.SetFocus
            Exit Sub
        End If
        txtChqNo.Tag = ChqBookNo
    End If

    If FundsTransfered Then
        Call PlaySaveSound
        MsgBox "Transaction Saved Successfully.", vbInformation
        Unload Me
        If ShowAgain Then Me.Show 1
    End If


End Sub
