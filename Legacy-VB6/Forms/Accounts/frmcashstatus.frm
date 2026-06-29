VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCashStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash & Bank Status"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2280
      TabIndex        =   8
      Top             =   2670
      Visible         =   0   'False
      Width           =   1620
   End
   Begin MSComctlLib.ListView LV 
      Height          =   2190
      Left            =   45
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   795
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   3863
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
         Text            =   "Bank & Branch"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Bank Account #"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Balance"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Chq Format"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.ListView LVCash 
      Height          =   2190
      Left            =   45
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3255
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   3863
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Account #"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Account Title"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Balance"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cash Balance Details"
      Height          =   270
      Left            =   45
      TabIndex        =   10
      Top             =   3015
      Width           =   9255
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   360
      Left            =   6585
      TabIndex        =   7
      Top             =   5625
      Width           =   1320
      ForeColor       =   64
      Caption         =   "Print"
      PicturePosition =   327683
      Size            =   "2328;635"
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
      Height          =   360
      Left            =   7980
      TabIndex        =   6
      Top             =   5610
      Width           =   1320
      ForeColor       =   64
      Caption         =   "Close     "
      PicturePosition =   327683
      Size            =   "2328;635"
      Picture         =   "frmcashstatus.frx":0000
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bank Balance Details"
      Height          =   270
      Left            =   45
      TabIndex        =   5
      Top             =   555
      Width           =   9255
   End
   Begin VB.Label lblBank 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   6855
      TabIndex        =   4
      Top             =   195
      Width           =   2430
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total Bank Balance :"
      Height          =   225
      Left            =   5160
      TabIndex        =   3
      Top             =   225
      Width           =   1695
   End
   Begin VB.Label lblCash 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   1365
      TabIndex        =   2
      Top             =   195
      Width           =   2430
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cash In Hand :"
      Height          =   225
      Left            =   120
      TabIndex        =   1
      Top             =   225
      Width           =   1200
   End
End
Attribute VB_Name = "frmCashStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PermissionDenied As Boolean
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report, rptSUB As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\CashnBankStatus.rpt")
    Set rptSUB = rpt.OpenSubreport("CashInHandBalance")
    rptSUB.RecordSelectionFormula = "{Accounts.AccNo}='" & CashInHand & "'"
    Dim f As New frmPrevRpt
    Unload Me
    f.ShowReport "NOT ({VBankAccounts.Balance}=0 To 1000)", rpt
    
    Exit Sub
err:
    MsgBox err.Description
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

    'PermissionDenied = Not UserHasAccess("CashBankStatus")

    Dim Balance As Double

    'Balance = getBalance(CashInHand, #1/1/3000#)
    Balance = GetSingleLongValue("SUM(Balance)", "VActiveAccounts", " WHERE LEFT(AccNo,6)='" & CashInHand & "'")
    lblCash = Format(Balance, "#,##0.00 DR ;#,##0.00 CR ")
    Balance = Val(con.Execute("SELECT SUM(Balance) FROM VBankAccounts").Fields(0) & "")
    lblBank = Format(Balance, "#,##0.00 DR ;#,##0.00 CR ")

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim lColor As Long
    With rs
       LV.ListItems.Clear
       .CursorLocation = adUseClient
       .Open "SELECT * FROM VBankAccounts", con, adOpenForwardOnly, adLockReadOnly
       For i = 0 To .RecordCount - 1
          Set ITM = LV.ListItems.add(, !AccNo, ![Bank] & ", " & ![Branch])
          ITM.ListSubItems.add , , ![BankAccNo] & ""
          'Itm.SubItems(2) = ![AccNo] & ""
          ITM.ListSubItems.add , , Format(![Balance], "#,##0.00 DR ;#,##0.00 CR ")
          
          ITM.ListSubItems.add , , !ChqFormat & ""
            If Val(!Balance & "") < 0 Then
                lColor = vbRed
            Else
                lColor = LV.ForeColor
            End If
            ITM.ForeColor = lColor
            For j = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(j).ForeColor = lColor
            Next
          .MoveNEXT
       Next
       .Close
       
       .Open "SELECT * FROM VActiveAccounts WHERE LEFT(AccNo,6)='" & CashInHand & "'", con, adOpenForwardOnly, adLockReadOnly
        LVCash.ListItems.Clear
        Do Until .EOF
            Set ITM = LVCash.ListItems.add(, , !AccNo & "")
            ITM.ListSubItems.add , , !AccTitle & ""
            ITM.ListSubItems.add , , Format(![Balance], "#,##0.00 DR ;#,##0.00 CR ")
            If Val(!Balance & "") < 0 Then
                lColor = vbRed
            Else
                lColor = LV.ForeColor
            End If
            ITM.ForeColor = lColor
            For j = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(j).ForeColor = lColor
            Next
            .MoveNEXT
        Loop
       .Close
    End With
    Set rs = Nothing

End Sub



Private Sub LV_DblClick()
     
    If LV.ListItems.count = 0 Then Exit Sub
    With txtDesc
        .Move LV.Left + LV.ColumnHeaders(4).Left + 85, LV.Top + LV.SelectedItem.Top + 25, LV.ColumnHeaders(4).Width ', LV.SelectedItem.Height
        .Text = OldDesc
        Set .Font = LV.Font
        .Tag = DpstSlip
        .Visible = True
        .SetFocus
    End With
    
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    LV.ToolTipText = Item.Text
End Sub

Private Sub txtDesc_GotFocus()

    txtDesc.SelStart = 0
    txtDesc.SelLength = Len(txtDesc)
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
  
If KeyAscii = 13 Then
    
    If Trim(txtDesc = "") Then
        MsgBox "Invalid Cheque Format Name", vbInformation
        Exit Sub
    End If
    
   con.Execute "Update BankAccounts Set Chqformat='" & txtDesc & "' Where AccNo='" & LV.SelectedItem.key & "'"
   LV.SelectedItem.ListSubItems(3) = txtDesc
   txtDesc.Visible = False

ElseIf KeyAscii = 27 Then
   txtDesc.Visible = False

End If

End Sub

Private Sub txtDesc_LostFocus()
    txtDesc.Visible = False
End Sub
