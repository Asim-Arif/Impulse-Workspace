VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmChqs 
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
            Picture         =   "frmChqs.frx":0000
            Key             =   "PND"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChqs.frx":0458
            Key             =   "CAN"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChqs.frx":08AA
            Key             =   "CLR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChqs.frx":0CFE
            Key             =   "DUE"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChqs.frx":1156
            Key             =   "BNC"
         EndProperty
      EndProperty
   End
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11835
      Begin VB.ComboBox cmbRange 
         Height          =   315
         ItemData        =   "frmChqs.frx":15AA
         Left            =   2625
         List            =   "frmChqs.frx":15BD
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
      Begin VB.PictureBox Picture1 
         Height          =   315
         Left            =   60
         ScaleHeight     =   255
         ScaleWidth      =   2460
         TabIndex        =   11
         Top             =   360
         Width           =   2520
         Begin VB.OptionButton Opts 
            Caption         =   "Payments"
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
            Index           =   0
            Left            =   15
            TabIndex        =   0
            Top             =   30
            Value           =   -1  'True
            Width           =   1230
         End
         Begin VB.OptionButton Opts 
            Caption         =   "Receipts"
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
            Index           =   1
            Left            =   1230
            TabIndex        =   1
            Top             =   30
            Width           =   1125
         End
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   6750
         TabIndex        =   4
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   16646147
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   5160
         TabIndex        =   3
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   16646147
         CurrentDate     =   37055
      End
      Begin MSForms.CheckBox CheckBox1 
         Height          =   300
         Left            =   8400
         TabIndex        =   5
         Top             =   375
         Width           =   1920
         BackColor       =   -2147483633
         ForeColor       =   8388608
         DisplayStyle    =   4
         Size            =   "3387;529"
         Value           =   "0"
         Caption         =   "Show Cleared Chqs"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdRefresh 
         Default         =   -1  'True
         Height          =   360
         Left            =   10440
         TabIndex        =   6
         Top             =   285
         Width           =   1305
         Caption         =   "Refresh  "
         PicturePosition =   327683
         Size            =   "2302;635"
         Accelerator     =   82
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00BECCD8&
         BackStyle       =   0  'Transparent
         Caption         =   "  Cheque Type                               Cheque  Range                             From                         To"
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
         Height          =   240
         Index           =   2
         Left            =   45
         TabIndex        =   8
         Top             =   135
         Width           =   11745
      End
   End
   Begin MSComctlLib.ListView LVCheque 
      Height          =   5280
      Left            =   30
      TabIndex        =   7
      Top             =   795
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   9313
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
   Begin MSForms.CommandButton cmdClose 
      Height          =   390
      Left            =   10275
      TabIndex        =   9
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
Attribute VB_Name = "frmChqs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PermissionDenied As Boolean
Private Sub RefreshChqList()
   Screen.MousePointer = vbhurglass
    Set LVCheque.SmallIcons = ImageList1
    Dim ListColor As Long
    Dim CrsLedger As New ADODB.Recordset
    Dim ClearedOption As String
    If chkCleared = vbUnchecked Then ClearedOption = " and Posted=0 "
    With CrsLedger
        .CursorLocation = adUseClient
        '.Open "Select * from VChqLedger Where (CDate between '" & DtFrom & "' and '" & DtTo & "') And Payment=" & Abs(opts(0)) & _
            ClearedOption & " and (ChequeType<>'Canceled') Order by CDate,Sno", con, 0, 1
        .Open "Select * from VChqLedger Where (ChequeDate between '" & DtFrom & "' and '" & DtTo & "') AND ChqIsDue=1 Order by CDate,Sno", con, 0, 1
            
            LVCheque.ListItems.Clear
        
            Dim ITM As ListItem
            Do Until .EOF
'                If ![Bounced] Or ![ChequeType] = "Canceled" Then
'                  ListColor = vbRed
'                ElseIf (![ChequeDate] <= Date And Not ![Posted]) Then
'                  ListColor = vbBlue
'                ElseIf ![Posted] Then
'                  ListColor = RGB(0, 125, 0)
'
'                Else
'                  ListColor = vbBlack
'                End If
                ListColor = vbRed
                Set ITM = LVCheque.ListItems.add(, ![Sno] & "'", ![ChequeNo] & "")
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
                
                
                'Check For The Chq Status
                Dim Status As String, img As String
'                If ![ChequeType] = "Canceled" Then
'                  Status = "Canceled"
'                  img = "CAN"
'                ElseIf ![Posted] Then
'                  Status = "Cleared"
'                  img = "CLR"
'                ElseIf ![Bounced] Then
'                  Status = "Bounced"
'                  img = "BNC"
'                ElseIf ITM.ForeColor = vbBlue Then
'                  Status = "Ready To Clear"
'                  img = "DUE"
'
'                Else
'                  Status = ""
'                  img = "PND"
'                End If
                Status = "Ready To Clear"
                img = "DUE"
                 ITM.ToolTipText = ITM.Tag & " [" & Status & "]"
                 ITM.SmallIcon = img
                 
                             
                
                
                
                .MoveNext
                'LVCheque.Visible = True
                'FDrag.Visible = True
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

Private Sub DtFrom_Change()
    Call setDateRange(DtFrom)
End Sub

Private Sub DtTo_Change()
    Call setDateRange(DtTo)
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

    'PermissionDenied = Not UserHasAccess("DueCheques")
    Call GetBankAccountsInMenu
    cmbRange.ListIndex = 0
    mnuPrint.Visible = False
    mnuClear.Visible = False
    mnupop.Visible = False
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
Dim TitleBottom As Long
FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
TitleBottom = FAcc.Top + FAcc.Height
LVCheque.Move 50, TitleBottom, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + cmdClose.Height + 100)
cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, LVCheque.Top + LVCheque.Height + 50

End Sub

Private Sub LVCheque_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If OutOfFinancialYear Then Exit Sub

    With LVCheque
    
      If .ListItems.count = 0 Or Button <> 2 Then Exit Sub
        
         If .SelectedItem.ForeColor <> vbBlue Then
            If Opts(0) Then Me.PopUpMenu mnuPrint ' Payment Chq
            Exit Sub
         End If
         
        
      If .SelectedItem.ListSubItems(6).Text = "Open" Then
         mnuToCash.Visible = True
      Else
         mnuToCash.Visible = False
      End If
        
      If Opts(1) Then             ' Receipt Chq
         Me.PopUpMenu mnupop
      ElseIf Opts(0) Then         ' Payment Chq
         Me.PopUpMenu mnuClear
      End If
    End With
    
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

Private Sub Opts_Click(Index As Integer)
If Opts(0) Then
   LVCheque.ColumnHeaders(5).Text = "Paid To"
   LVCheque.ColumnHeaders(2).Text = "Issued On"
Else
   LVCheque.ColumnHeaders(5).Text = "Received From"
   LVCheque.ColumnHeaders(2).Text = "Rcvd. On"
End If
   Call cmdRefresh_Click
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
      
      
   .MoveNext
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
            .MoveNext
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
            .MoveNext
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
