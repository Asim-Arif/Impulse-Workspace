VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmChartOfAccounts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Balance Sheet..."
   ClientHeight    =   8625
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1275
      ScaleHeight     =   225
      ScaleWidth      =   5805
      TabIndex        =   3
      Top             =   1770
      Visible         =   0   'False
      Width           =   5835
      Begin VB.TextBox txtBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BorderStyle     =   0  'None
         DataField       =   "RptTime"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   3540
         MaxLength       =   12
         TabIndex        =   5
         Top             =   0
         Width           =   1680
      End
      Begin VB.TextBox txtAccTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BorderStyle     =   0  'None
         DataField       =   "RptTime"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   -15
         MaxLength       =   150
         TabIndex        =   4
         Top             =   0
         Width           =   3510
      End
      Begin VB.ComboBox cmbCrDr 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "frmBalanceSheet.frx":0000
         Left            =   5250
         List            =   "frmBalanceSheet.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   -45
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   405
      Left            =   10410
      TabIndex        =   2
      Top             =   8025
      Width           =   1395
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1695
      Top             =   3330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceSheet.frx":0016
            Key             =   "ACC"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceSheet.frx":046A
            Key             =   "PARIN"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceSheet.frx":057E
            Key             =   "PAR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceSheet.frx":0692
            Key             =   "ACCIN"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV 
      Height          =   7845
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   13838
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
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
         Text            =   "Account #"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Account Title"
         Object.Width           =   10583
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Opening Bal."
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Current Bal."
         Object.Width           =   3175
      EndProperty
   End
   Begin MSForms.CommandButton cmdOpt 
      Height          =   405
      Left            =   8895
      TabIndex        =   1
      Top             =   8025
      Width           =   1470
      Caption         =   "    Options"
      PicturePosition =   327683
      Size            =   "2593;714"
      Picture         =   "frmBalanceSheet.frx":0AE6
      Accelerator     =   79
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Menu mnupop 
      Caption         =   "mnupop"
      Visible         =   0   'False
      Begin VB.Menu mnurefresh 
         Caption         =   "Re&fresh"
      End
      Begin VB.Menu mnuShowAll 
         Caption         =   "Show All"
      End
      Begin VB.Menu mnuLedger 
         Caption         =   "View &Ledger For Last 30 Days"
      End
      Begin VB.Menu dash0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewAcc 
         Caption         =   "&New Account"
      End
      Begin VB.Menu mnuEditAcc 
         Caption         =   "&Edit Account"
      End
      Begin VB.Menu dash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActive 
         Caption         =   "Make This Account Ina&ctive"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTransfer 
         Caption         =   "&Transfer Funds"
      End
      Begin VB.Menu mnuJV 
         Caption         =   "Make &Journal Entry"
      End
      Begin VB.Menu dash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print Chart of Accounts"
      End
   End
End
Attribute VB_Name = "frmChartOfAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim rec As New ADODB.Recordset
Dim rsAccounts As New ADODB.Recordset
Dim i As Integer
Dim Bal As Double
Dim AccList As String
Dim SubAccList As String
Dim TotalBalance As Double

Dim ShowAll As Boolean







Private Sub Command2_Click()


Dim AccNo As String
Dim AccTitle As String
Dim Bal As Double

    For i = 1 To LV.ListItems.Count
        AccNo = LV.ListItems(i).Text & ""
        AccTitle = LV.ListItems(i).ListSubItems(1).Text
        Bal = Val(LV.ListItems(i).ListSubItems(2).Tag)
        Con.Execute "Delete * From TempChaccount"
        Con.Execute "Insert into TempChAccount Values('" & AccNo & "','" & AccTitle & "'," & Bal & ")"
    Next i
    
    
End Sub




Private Sub RefreshList()

LV.Sorted = False

Dim PrevIndex As Double

If LV.ListItems.Count > 0 Then PrevIndex = LV.SelectedItem.Index


LV.Visible = False
LV.ListItems.Clear
Dim i As Integer
Me.MousePointer = vbHourglass


With rsAccounts
      
    .MoveFirst
    For i = 0 To .RecordCount - 1
        TotalBalance = 0
        AccList = ""
        Call CalculateBalance(.Fields("AccNo"))
        Call AddToList(i, .Fields("SubaccOf") & "", .Fields("AccNo"), .Fields("AccTitle"), .Fields("OpenBal"), TotalBalance, .Fields("AccNo"), ![Parent], ![Active])
        .MoveNext
    Next i
End With
LV.Visible = True
LV.Sorted = True
If PrevIndex > 0 And PrevIndex <= LV.ListItems.Count Then LV.ListItems(PrevIndex).Selected = True
Me.MousePointer = vbNormal



End Sub

Private Sub cmbCrDr_Click()
 If cmbCrDr.Text = "CR" Then
  cmbCrDr.Tag = -Val(txtBal.Text)
 Else
  cmbCrDr.Tag = Val(txtBal.Text)
 End If
End Sub

Private Sub cmbCrDr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call UpdateAccount
End If
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdOpt_Click()
With cmdOpt
 Call CheckMenu
 Me.PopupMenu mnupop, , .Left, .Top
End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 And PicEdit.Visible Then
   PicEdit.Visible = False
ElseIf KeyAscii = 13 Then
 SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
Me.KeyPreview = True
Dim Condition As String
If ShowAll Then
   Condition = ""
Else
   Condition = "Where Active "
End If
   Set LV.SmallIcons = ImageList1
   rsAccounts.Open "Select * from Accounts " & Condition & " Order By Accno", Con, adOpenForwardOnly
   Call RefreshList
End Sub
Private Sub AddToList(No As Integer, Parent As String, Code As String, AccTitle As String, OpenBal, Bal As Double, key As String, ParentAccount As Boolean, Active As Boolean)
Dim LColor As Long
Dim NumSpaces As Integer
Dim TheTag As String

' Format Opening Balance

Dim OpenBalance As String
If OpenBal < 0 Then
   OpenBalance = Format(Abs(OpenBal), "#,##0.00") & "   CR"
ElseIf OpenBal > 0 Then
   OpenBalance = Format(Abs(OpenBal), "#,##0.00") & "   DR"
Else
   OpenBalance = Format(Abs(OpenBal), "#,##0.00") & "        "
End If

'Format current Balance


Dim CurrBalance As String
If Bal < 0 Then
   CurrBalance = Format(Abs(Bal), "#,##0.00") & "   CR"
ElseIf Bal > 0 Then
   CurrBalance = Format(Abs(Bal), "#,##0.00") & "   DR"
Else
      CurrBalance = Format(Abs(Bal), "#,##0.00") & "        "
End If


If Not Active Then
   LColor = &H808080
   TheTag = "Inactive"
Else
   LColor = &H800000
   TheTag = "Active"
End If

'For Image Selection
Dim img As String
If ParentAccount Then
   If Active Then
     img = "PAR"
   Else
     img = "PARIN"
   End If
Else
   If Active Then
     img = "ACC"
   Else
     img = "ACCIN"
   End If
End If
   
   
   
   

Dim i As Integer




For i = 1 To LV.ListItems.Count
    If Parent = Replace(LV.ListItems(i).key, "[", "") Then
    
        NumSpaces = Len(LV.ListItems(i).ListSubItems(1).Text) - Len(Trim(LV.ListItems(i).ListSubItems(1).Text))
         
         
        

        
        LV.ListItems.Add i + 1, "[" & key, Code, , img
        LV.ListItems(i + 1).Tag = TheTag
        LV.ListItems(i + 1).ForeColor = LColor
        LV.ListItems(i + 1).ListSubItems.Add , , Space$(5 + NumSpaces) & "»  " & AccTitle
        LV.ListItems(i + 1).ListSubItems(1).ForeColor = LColor
        LV.ListItems(i + 1).ListSubItems(1).Bold = ParentAccount
        LV.ListItems(i + 1).ListSubItems.Add , , OpenBalance
        LV.ListItems(i + 1).ListSubItems(2).Tag = OpenBal
        LV.ListItems(i + 1).ListSubItems(2).ForeColor = LColor
        LV.ListItems(i + 1).ListSubItems(2).Bold = ParentAccount
        LV.ListItems(i + 1).ListSubItems.Add , , CurrBalance
        LV.ListItems(i + 1).ListSubItems(3).ForeColor = LColor
        LV.ListItems(i + 1).ListSubItems(3).Tag = Bal
        LV.ListItems(i + 1).ListSubItems(3).Bold = ParentAccount
        Exit Sub
    End If
Next i


LV.ListItems.Add , "[" & key, Code, , img
LV.ListItems(No + 1).Tag = TheTag
LV.ListItems(No + 1).Bold = ParentAccount
LV.ListItems(No + 1).ForeColor = LColor
LV.ListItems(No + 1).ListSubItems.Add , , AccTitle
LV.ListItems(No + 1).ListSubItems(1).ForeColor = LColor
LV.ListItems(No + 1).ListSubItems(1).Bold = ParentAccount

LV.ListItems(No + 1).ListSubItems.Add , , OpenBalance
LV.ListItems(No + 1).ListSubItems(2).Bold = ParentAccount
LV.ListItems(No + 1).ListSubItems(2).ForeColor = LColor
LV.ListItems(No + 1).ListSubItems(2).Tag = OpenBal

LV.ListItems(No + 1).ListSubItems.Add , , CurrBalance
LV.ListItems(No + 1).ListSubItems(3).Bold = ParentAccount
LV.ListItems(No + 1).ListSubItems(3).ForeColor = LColor
LV.ListItems(No + 1).ListSubItems(3).Tag = Bal

End Sub

Private Sub CreateLists(PRec As ADODB.Recordset, AccNo As String)

'If Len(AccList) > 0 Then AccList = AccList & "'" & AccNo & "',"
AccList = AccList & "'" & AccNo & "',"
SubAccList = ""
Dim i As Integer
With PRec
    For i = 0 To .RecordCount - 1
        AccList = AccList & "'" & .Fields("AccNo") & "',"
        SubAccList = SubAccList & "'" & .Fields("AccNo") & "',"
        .MoveNext
    Next i
    
End With
If Right$(SubAccList, 1) = "," Then SubAccList = Left$(SubAccList, Len(SubAccList) - 1)

End Sub

Private Sub CalculateBalance(AccNo As String)

Dim i As Integer
With rec

    If .State = 1 Then .Close
    .CursorLocation = adUseClient
    .Open "Select AccNo from Accounts Where SubAccOf='" & AccNo & "'", Con, adOpenForwardOnly, adLockReadOnly
    Do While .RecordCount > 0
        Call CreateLists(rec, AccNo)
        If .State = 1 Then .Close
    .CursorLocation = adUseClient
    'If Right$(SubAccList, 1) = "," Then
    '    Useable = Left$(SubAccList, Len(SubAccList) - 1)
    'Else
    '    Useable = SubAccList
    'End If
    .Open "Select AccNo from Accounts Where SubAccOf in (" & SubAccList & ")", Con, adOpenForwardOnly, adLockReadOnly
    Loop
    If Len(AccList) > 0 Then
        AccList = Left$(AccList, Len(AccList) - 1)
    Else
        AccList = "'" & AccNo & "'"
    End If
    If .State = 1 Then .Close
    .CursorLocation = adUseClient
    .Open "Select Sum(Balance) from Accounts Where AccNo in (" & AccList & ")", Con, adOpenForwardOnly, adLockReadOnly
    TotalBalance = .Fields(0)
    
End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
rsAccounts.Close
Set rsAccounts = Nothing
End Sub

Private Sub LV_DblClick()



With LV

   If .ListItems.Count = 0 Then Exit Sub
   
   If .SelectedItem.Tag = "Inactive" Then
      MsgBox "Inactive Account Cannot Be Edited.", vbInformation
      Exit Sub
   End If
   
   txtBal.Enabled = .SelectedItem.SmallIcon <> "PAR"
   cmbCrDr.Enabled = .SelectedItem.SmallIcon <> "PAR"
   
   
   PicEdit.Move .Left + .ColumnHeaders(2).Left + 25, LV.Top + .SelectedItem.Top + 40, .ColumnHeaders(2).Width + .ColumnHeaders(3).Width
   txtAccTitle.Move 0, 0, .ColumnHeaders(2).Width
   txtBal.Move txtAccTitle.Width + 25, 0, .ColumnHeaders(3).Width - cmbCrDr.Width - 15
   cmbCrDr.Move PicEdit.Width - cmbCrDr.Width
   
   txtAccTitle = Trim(Replace(.SelectedItem.SubItems(1), "»", ""))
   
   txtAccTitle.FontBold = .SelectedItem.Bold
   
   txtBal = Val(Replace(.SelectedItem.SubItems(2), ",", ""))
   'FORMAT AND STORE IN TAG
   Call txtBal_LostFocus
   Call FindInCombo(cmbCrDr, Right(Trim(.SelectedItem.SubItems(2)), 2))
   'STORE PREVIOUS OPENING BALANCE
   PicEdit.Tag = Val(.SelectedItem.ListSubItems(2).Tag)
End With
PicEdit.Visible = True
PicEdit.ZOrder
txtAccTitle.SetFocus
End Sub

Private Sub LV_GotFocus()
If PicEdit.Visible Then Call Form_KeyPress(27)
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
 Call CheckMenu
 Me.PopupMenu mnupop, , (LV.Left + (LV.Width / 2)), (LV.Top + LV.SelectedItem.Top)
End If
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 2 Then Exit Sub
 Call CheckMenu
 Me.PopupMenu mnupop
 
End Sub

Private Sub mnuActive_Click()

   If mnuActive.Tag = "Active" Then
      If Not MakeAccountInActive Then Exit Sub
   ElseIf mnuActive.Tag = "Inactive" Then
      If Not MakeAccountActive Then Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
'   Sleep (500)
   Call mnurefresh_Click
   Screen.MousePointer = vbDefault
End Sub

Private Sub mnuEditAcc_Click()
   Call LV_DblClick
End Sub

Private Sub mnuJV_Click()
If LV.SelectedItem.Tag = "Inactive" Then
   MsgBox "No Transaction Can Be Made In An Inactive Account.", vbInformation
   Exit Sub
End If

If frmJV.Visible Then
 MsgBox "There Is Some Transaction Already Active"
Else
   
   With frmJV
      .Show
      
      Call FindInCombo(.cmbAccNo, LV.SelectedItem)
   End With
   
End If
End Sub

Private Sub mnuLedger_Click()
Dim f As New frmLedger
With f
   .Show
   Call FindInCombo(.cmbAccNo, LV.SelectedItem)
   Call .cmdRefresh_Click
End With
End Sub

Private Sub mnuNewAcc_Click()
   frmNewAccount.Show 1
   Call mnurefresh_Click
End Sub

Private Sub mnurefresh_Click()
   Screen.MousePointer = vbHourglass
   rsAccounts.Close
   Set rsAccounts = Nothing
   Call Form_Load
   Screen.MousePointer = vbDefault
End Sub


Private Sub CheckMenu()

   If LV.ListItems.Count = 0 Then
     mnuEditAcc.Enabled = False
     mnuTransfer.Enabled = False
     mnuJV.Enabled = False
     mnuLedger.Enabled = False
     mnuPrint.Enabled = False
        
   Else
     mnuEditAcc.Enabled = True
     mnuTransfer.Enabled = True
     mnuJV.Enabled = True
     mnuLedger.Enabled = True = Not (LV.SelectedItem.SmallIcon = "PAR" Or LV.SelectedItem.SmallIcon = "PARIN")
    
     If LV.SelectedItem.Tag = "Inactive" Then
       mnuActive.Caption = "Make This Account A&ctive."
     Else
       mnuActive.Caption = "Make This Account Ina&ctive."
     End If
     mnuActive.Tag = LV.SelectedItem.Tag
   End If



End Sub

Private Sub mnuShowAll_Click()
mnuShowAll.Checked = Not mnuShowAll.Checked
ShowAll = mnuShowAll.Checked
Call mnurefresh_Click
End Sub

Private Sub mnuTransfer_Click()
   frmTransferFunds.Show 1
End Sub



Private Sub txtAccTitle_GotFocus()
txtAccTitle.SelStart = 0
txtAccTitle.SelLength = Len(txtAccTitle)
End Sub

Private Sub txtAccTitle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then If Not txtBal.Enabled Then Call UpdateAccount

End Sub


Private Sub txtBal_GotFocus()
txtBal.SelStart = 0
txtBal.SelLength = Len(Bal)
End Sub

Private Sub txtBal_KeyPress(KeyAscii As Integer)
   KeyAscii = OnlyNumber(KeyAscii, True)
End Sub

Private Sub txtBal_LostFocus()
 txtBal = Format(Val(txtBal), "###0.00")
 txtBal.Tag = LV.SelectedItem.ListSubItems(2).Tag
 Call cmbCrDr_Click
End Sub



Private Function MakeAccountInActive() As Boolean
MakeAccountInActive = False
 With LV.SelectedItem
   
   If .ListSubItems(3).Tag <> 0 Then
        MsgBox "This Account Cannot Be Made Inactive." & vbNewLine & _
           "To Make An Account Inactive Balance Must Be Zero.", vbInformation
         Exit Function
   ElseIf .SmallIcon = "PAR" Then
       If MsgBox("This Is A Master Account, All Sub Accounts Will Also Become Inactive." & vbNewLine & _
           "Are You Sure to Make This Account Inactive", vbYesNo + vbQuestion) = vbNo Then Exit Function
   Else
      If MsgBox("Are You Sure to Make This Account Inactive", vbYesNo + vbQuestion) = vbNo Then Exit Function
   End If
   
   
   
   
   Con.Execute "Update Accounts Set Active=false Where Accno='" & Trim(.Text) & _
               "' OR SubAccOf Like '" & Trim(.Text) & "'"
  
   MakeAccountInActive = True
End With

End Function


Private Function MakeAccountActive() As Boolean

MakeAccountActive = False
Dim AlsoChild As String

 With LV.SelectedItem
   
   If .SmallIcon = "PARIN" Then
       If MsgBox("This Is A Parent Account, Do You Want All Child Accounts To Become Active." _
          , vbYesNo + vbQuestion) = vbYes Then
          AlsoChild = " OR SubAccOf Like '" & Trim(.Text) & "'"
       Else
           AlsoChild = ""
       End If
   End If
   
   
   
   
   Con.Execute "Update Accounts Set Active=True Where Accno='" & Trim(.Text) & "'" & AlsoChild
  
MakeAccountActive = True
End With




End Function



Private Function UpdateAccount() As Boolean
 If txtAccTitle.Text = "" Then
   MsgBox "Invalid Account Title", vbInformation
   Exit Function
 End If
 
Dim Diff As Double

Screen.MousePointer = vbHourglass
Diff = Val(cmbCrDr.Tag) - Val(txtBal.Tag)
   Con.Execute "Update Accounts Set Acctitle='" & txtAccTitle & "',OpenBal=" & Val(cmbCrDr.Tag) & _
               ", Balance=Balance+" & Diff & " Where AccNo='" & LV.SelectedItem.Text & "'"
               
   'Adjust The Ledger
   Call adjustLedger(LV.SelectedItem.Text, #1/1/1900#, 0, Diff)
   
   'Sleep (500)
   Call mnurefresh_Click
Screen.MousePointer = vbDefault
   
End Function
