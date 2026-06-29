VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDueChqs 
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   255
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6690
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   105
      Top             =   1305
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDueChqs.frx":0000
            Key             =   "PND"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDueChqs.frx":0458
            Key             =   "CAN"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDueChqs.frx":08AA
            Key             =   "CLR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDueChqs.frx":0CFE
            Key             =   "DUE"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDueChqs.frx":1156
            Key             =   "BNC"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LVCheque 
      Height          =   6015
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   10610
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cheque #"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Issued On"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Chq. Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Bank & Branch"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Paid To"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Type"
         Object.Width           =   1969
      EndProperty
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   390
      Left            =   8655
      TabIndex        =   2
      Top             =   6135
      Width           =   1560
      ForeColor       =   64
      Caption         =   "Print"
      PicturePosition =   327683
      Size            =   "2752;688"
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
      Height          =   390
      Left            =   10275
      TabIndex        =   1
      Top             =   6135
      Width           =   1560
      ForeColor       =   64
      Caption         =   "Close            "
      PicturePosition =   327683
      Size            =   "2752;688"
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
      Caption         =   "mnupop"
      Begin VB.Menu mnuToCash 
         Caption         =   "Transfer To Cash in Hand"
         Begin VB.Menu mnuOnChqDate 
            Caption         =   "On Cheque's Date"
         End
         Begin VB.Menu mnuOnToday 
            Caption         =   "On Today's Date"
         End
      End
      Begin VB.Menu dash12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToBank 
         Caption         =   "Transfer To Bank Account On Cheque's Date"
         Begin VB.Menu mnuBanks 
            Caption         =   "< No Bank Accounts Defined >"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuToBankToday 
         Caption         =   "Transfer To Bank Account On Today's Date"
         Begin VB.Menu mnuBanksToday 
            Caption         =   "< No Bank Accounts Defined >"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChqBounced 
         Caption         =   "Cheque &Bounced"
      End
   End
   Begin VB.Menu mnuClear 
      Caption         =   "mnu"
      Begin VB.Menu mnuChqCashed 
         Caption         =   "Clear This Cheque On Cheque's Date"
      End
      Begin VB.Menu mnuChqCashedToday 
         Caption         =   "Clear This Cheque On Today's Date"
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBounced 
         Caption         =   "Cheque &Bounced"
      End
      Begin VB.Menu dashprint 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintChq 
         Caption         =   "Print This Cheque"
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "mnuPrint"
      Begin VB.Menu mnuClearChq 
         Caption         =   "Clear this Cheque"
      End
   End
End
Attribute VB_Name = "frmDueChqs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PermissionDenied As Boolean

Private Sub RefreshChqList()

    Screen.MousePointer = vbHourglass
    Dim ListColor As Long
    Dim CrsLedger As New ADODB.Recordset
    With CrsLedger
        .CursorLocation = adUseClient
        .Open "SELECT * FROM VChqLedger WHERE ClearanceDT IS NULL ORDER BY CDate,Sno", con, 0, 1
            
            LVCheque.ListItems.Clear
        
            Dim ITM As ListItem
            Do Until .EOF
                ListColor = vbRed
                Set ITM = LVCheque.ListItems.add(, ![SNo] & "'", ![ChequeNo] & "")
                ITM.ForeColor = ListColor
                
                For i = 1 To 7
                    With ITM.ListSubItems.add(, , "")
                     .ForeColor = ListColor
                    End With
                    
                Next i
                
                ITM.ListSubItems(1) = Format(![CDate], "dd-MMM-yyyy")
                ITM.Tag = ![Description]
                
                ITM.ListSubItems(2) = Format(![ChequeDate], "dd-MMM-yyyy")
                
                ITM.ListSubItems(3) = ![Bank] & ", " & ![Branch]
                ITM.ListSubItems(3).ToolTipText = ![Bank] & ", " & ![Branch]
                
                ITM.ListSubItems(4) = ![AccTitle] & ""
                ITM.ListSubItems(4).ToolTipText = ![AccTitle] & ""
                ITM.ListSubItems(4).Tag = ![AccNo] & ""
                
                ITM.ListSubItems(5) = Format(Val(![Amount] & ""), "###,##0.00")
                ITM.ListSubItems(6) = ![ChequeType] & ""
                
               
                Dim Status As String, img As String
                Status = "Ready To Clear"
                img = "DUE"
                ITM.ToolTipText = ITM.Tag & " [" & Status & "]"

                
                .MoveNEXT
              
            Loop
            
        
        .Close
    End With
    
    Set CrsLedger = Nothing
    Screen.MousePointer = vbDefault
    
End Sub


Private Sub chkCleared_Click()
   Call cmdRefresh_Click
End Sub

Private Sub cmbRange_Click()
DtFrom.Enabled = False
DtTo.Enabled = False
Select Case cmbRange.ListIndex
   Case 0
      DtFrom = YearStartDate
      DtTo = Date
   Case 1
      DtFrom = YearStartDate
      DtTo = Date + 1
   Case 2
      DtFrom = YearStartDate
      DtTo = Date + 7
   Case 3
      DtFrom = YearStartDate
      DtTo = Date + 30
   Case 4
      DtFrom.Enabled = True
      DtTo.Enabled = True
      DtFrom.SetFocus
      Exit Sub
End Select

   Call cmdRefresh_Click
   
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdRefresh_Click()
   Call RefreshChqList
End Sub

Private Sub cmdPrint_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\DueCheques.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "ISNULL({VChqBookLedger.ClearanceDT})", rpt
    
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

Private Sub Form_Load()
 
    mnuPrint.Visible = False
    mnuClear.Visible = False
    mnuPop.Visible = False

    Call RefreshChqList
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Dim TitleBottom As Long
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    TitleBottom = 100 'FAcc.Top + FAcc.Height
    LVCheque.Move 50, TitleBottom, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + cmdClose.Height + 100)
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, LVCheque.Top + LVCheque.Height + 50
    cmdPrint.Move cmdClose.Left - cmdPrint.Width - 50, cmdClose.Top
    
End Sub

Private Sub mnuChqCashedToday_Click()
If MsgBox("Are You Sure To Make This Cheque Cleared.", vbInformation + vbYesNo) = vbNo Then Exit Sub

   Screen.MousePointer = vbHourglass
    
   If SavePaymentTrans(Date) Then
      Call PlaySaveSound
      MsgBox "Transaction Saved Successfully", vbInformation
      Call RefreshChqList
   End If
       
   Screen.MousePointer = vbDefault
End Sub

Private Sub mnuClearChq_Click()
    con.Execute "UPDATE Cheque SET ChqIsDue=0 WHERE Sno=" & Val(LVCheque.SelectedItem.key)
    Call RefreshChqList
End Sub

Private Sub mnuOnChqDate_Click()
If MsgBox("Are You Sure To Transfer This Cheque To Cash In Hand.", vbInformation + vbYesNo) = vbNo Then Exit Sub
   Screen.MousePointer = vbHourglass
   Dim Dated As Date
   Dated = CDate(LVCheque.SelectedItem.SubItems(2))
   If SaveReceiptTrans(CashInHand, Dated) Then
      Call PlaySaveSound
        
      MsgBox "Transaction Saved Successfully", vbInformation
       Call RefreshChqList
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub mnuOnToday_Click()
If MsgBox("Are You Sure To Transfer This Cheque To Cash In Hand.", vbInformation + vbYesNo) = vbNo Then Exit Sub
   Screen.MousePointer = vbHourglass
   If SaveReceiptTrans(CashInHand, Date) Then
      Call PlaySaveSound
      MsgBox "Transaction Saved Successfully", vbInformation
      Call RefreshChqList
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub mnuprintChq_Click()
'Call PrintChq(frmFlash.Cr1, Val(LVCheque.SelectedItem.key))
End Sub

Private Sub mnuPrintChq1_Click()
   Call mnuprintChq_Click
End Sub

Private Sub GetBankAccountsInMenu()
  Dim rs As New ADODB.Recordset
 With rs
  .Open "Select * from VBankAccounts Order By AccNo", con, adOpenForwardOnly, adLockReadOnly
  
  For i = 0 To .RecordCount - 1
      Load mnuBanks(i + 1)
      Load mnuBanksToday(i + 1)
      mnuBanks(i + 1).Caption = "[" & ![AccNo] & "] " & ![AccTitle]
      mnuBanksToday(i + 1).Caption = "[" & ![AccNo] & "] " & ![AccTitle]
      
      mnuBanks(i + 1).Enabled = True
      mnuBanksToday(i + 1).Enabled = True
      
      mnuBanks(i + 1).Tag = ![AccNo] & ""
      mnuBanksToday(i + 1).Tag = ![AccNo] & ""
      
      
   .MoveNEXT
  Next
   mnuBanks(0).Visible = .RecordCount = 0
   mnuBanksToday(0).Visible = .RecordCount = 0
  .Close
 End With
 Set rs = Nothing
  
End Sub


Private Sub mnuBanks_Click(Index As Integer)

If MsgBox("Are You Sure To Transfer This Cheque To " & mnuBanks(Index).Caption, vbInformation + vbYesNo) = vbNo Then Exit Sub
   Dim SlipNo As String
   SlipNo = InputBox("Enter Deposit Slip No.")
   
   If Trim(SlipNo) = "" Then
      MsgBox "Transaction Terminated Due To Invlalid Bank Deposit SlipNo", vbInformation
      Exit Sub
   End If
   
   Dim Dated As Date
   Dated = CDate(LVCheque.SelectedItem.SubItems(2))
   
   Screen.MousePointer = vbHourglass
   If SaveReceiptTrans(mnuBanks(Index).Tag, Dated, SlipNo) Then
      Call PlaySaveSound
      MsgBox "Transaction Saved Successfully", vbInformation
      Call RefreshChqList
   End If
   Screen.MousePointer = vbDefault
End Sub
Private Sub mnuBanksToday_Click(Index As Integer)

If MsgBox("Are You Sure To Transfer This Cheque To " & mnuBanks(Index).Caption, vbInformation + vbYesNo) = vbNo Then Exit Sub
   Dim SlipNo As String
   SlipNo = InputBox("Enter Deposit Slip No.")
   
   If Trim(SlipNo) = "" Then
      MsgBox "Transaction Terminated Due To Invlalid Bank Deposit SlipNo", vbInformation
      Exit Sub
   End If
   
   
   Screen.MousePointer = vbHourglass
   If SaveReceiptTrans(mnuBanks(Index).Tag, Date, SlipNo) Then
      Call PlaySaveSound
      MsgBox "Transaction Saved Successfully", vbInformation
      Call RefreshChqList
   End If
   Screen.MousePointer = vbDefault
End Sub
Private Sub mnuBounced_Click()
Call mnuChqBounced_Click
End Sub

Private Sub mnuChqBounced_Click()
   If MsgBox("Are You Sure To Make This Cheque Bounced.", vbInformation + vbYesNo) = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
      con.Execute "Update Cheque Set Bounced=1 Where Sno=" & Val(LVCheque.SelectedItem.key)
      'Sleep (1000)
      Call RefreshChqList
      Screen.MousePointer = vbDefault
End Sub

Private Sub mnuChqCashed_Click()
If MsgBox("Are You Sure To Make This Cheque Cleared.", vbInformation + vbYesNo) = vbNo Then Exit Sub

   Screen.MousePointer = vbHourglass
   Dim Dated As Date
   Dated = CDate(LVCheque.SelectedItem.SubItems(2))
   If SavePaymentTrans(Dated) Then
      Call PlaySaveSound
      MsgBox "Transaction Saved Successfully", vbInformation
      Call RefreshChqList
   End If
       
   Screen.MousePointer = vbDefault
   
   
End Sub

Private Function SavePaymentTrans(Dated As Date) As Boolean

    Dim i As Integer
    Dim NextVchrNo As String
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim dAmount As Double
    Dim sScNo As String
    Dim sCIH As String
    Dim Bal As Double
    Dim Desc As String
    Dim AccountNo As String
    
    
    
    'Account No To Which Cheque is Paid
    AccountNo = LVCheque.SelectedItem.ListSubItems(4).Tag
    
    'Get Serial # and Balances
    NextVchrNo = getNextVchrNo(Dated, "BPV")
    INextSno = getNextSno(Dated)
    SnoToUse = INextSno
    Bal = getBalance(AccountNo, Dated)
    dAmount = Format$(LVCheque.SelectedItem.ListSubItems(5).Text, "###")
    sScNo = Left$(LVCheque.SelectedItem.key, Len(LVCheque.SelectedItem.key) - 1)
    
    
    
    sCIH = con.Execute("Select AccNo From VPayChqs Where Sno=" & Val(LVCheque.SelectedItem.key)).Fields(0)
    
       
    
    
    
    Bal = Bal + dAmount
    'Start New Transaction
    con.BeginTrans
    'Debit Entry
    con.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & NextVchrNo & "','" & AccountNo & _
        "','" & LVCheque.SelectedItem.Tag & _
        "'," & dAmount & ",0," & Bal & ",''," & sScNo & ")"
    
    
    
    Call adjustLedger(AccountNo, Dated, INextSno, Val(dAmount))
    
    Dim ChqNo As String
    INextSno = INextSno + 1
    Bal = getBalance(sCIH, Dated)
    Bal = Bal - dAmount
    'Description
    Dim ChqRec As New ADODB.Recordset
    With ChqRec
        .CursorLocation = adUseClient
        .Open "Select ChequeNo From Cheque Where (Sno=" & sScNo & ")", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            Desc = "Cheque Paid To " & LVCheque.SelectedItem.ListSubItems(4) & " Chq# : " & _
                .Fields(0) & ""
             ChqNo = .Fields(0) & ""
        Else
            Desc = "Cheque Paid To " & LVCheque.SelectedItem.ListSubItems(4) & " Chq# : "
        End If
        
        .Close
    End With
    Set ChqRec = Nothing
    'Credit Entry
    
    con.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & NextVchrNo & "','" & sCIH & _
        "','" & Desc & _
        "',0," & dAmount & "," & Bal & ",''," & sScNo & ")"
        
    
    
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
    'Commit Transaction
    con.CommitTrans
   
    
    
    SavePaymentTrans = True
    Exit Function
err:
    MsgBox "Transaction Terminated Due To Following Error;" & vbNewLine & err.Description
    con.RollbackTrans
        SavePaymentTrans = False
End Function



Private Function SaveReceiptTrans(ReceiveInAccNo As String, Dated As Date, Optional SlipNo As String) As Boolean

    Dim i As Integer
    Dim NextVchrNo As String
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim dAmount As Double
    Dim sScNo As String
    Dim sCIH As String
    Dim Bal As Double
    Dim Desc As String
    Dim AccountNo As String
    
    'Set The Transaction Date
    
    
    
    'Get Serial ' and Balances
    
    AccountNo = LVCheque.SelectedItem.ListSubItems(4).Tag
    
    NextVchrNo = getNextVchrNo(Dated, "BRV")
    INextSno = getNextSno(Dated)
    SnoToUse = INextSno
    Bal = getBalance(AccountNo, Dated)
    dAmount = Format$(LVCheque.SelectedItem.ListSubItems(5).Text, "###")
    sScNo = Left$(LVCheque.SelectedItem.key, Len(LVCheque.SelectedItem.key) - 1)
    sCIH = ReceiveInAccNo
    Bal = Bal - dAmount
    'Start New Transaction
    con.BeginTrans
    
    
     
    
    'Credit Entry
    con.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & NextVchrNo & "','" & AccountNo & _
        "','" & LVCheque.SelectedItem.Tag & _
        "',0," & dAmount & "," & Bal & ",'" & SlipNo & "'," & sScNo & ")"
    
    
    
    Call adjustLedger(AccountNo, Dated, INextSno, -Val(dAmount))
    
    Dim ChqNo As String
    INextSno = INextSno + 1
    Bal = getBalance(sCIH, Dated)
    Bal = Bal + dAmount
    'Description
    Dim ChqRec As New ADODB.Recordset
    With ChqRec
        .CursorLocation = adUseClient
        .Open "Select ChequeNo From Cheque Where (Sno=" & sScNo & ")", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            Desc = "Cheque Recieved from " & LVCheque.SelectedItem.ListSubItems(4) & " Chq# : " & _
                .Fields(0) & ""
             ChqNo = .Fields(0) & ""
        Else
            Desc = "Cheque Recieved from " & LVCheque.SelectedItem.ListSubItems(4) & " Chq# : "
        End If
        
        .Close
    End With
    Set ChqRec = Nothing
    'Debit Entry
    
    con.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & NextVchrNo & "','" & sCIH & _
        "','" & Desc & _
        "'," & dAmount & ",0," & Bal & ",'" & SlipNo & "'," & sScNo & ")"
        
    
    
    Call adjustLedger(sCIH, Dated, INextSno, Val(dAmount))
    
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
    'Commit Transaction
    con.CommitTrans
    
    
   
    
    SaveReceiptTrans = True
    Exit Function

err:
    MsgBox "Transaction Terminated Due To Following Error;" & vbNewLine & err.Description
    con.RollbackTrans
        SaveReceiptTrans = False
End Function
