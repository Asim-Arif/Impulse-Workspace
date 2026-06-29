VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmItems 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item List"
   ClientHeight    =   8595
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport cr1 
      Left            =   390
      Top             =   630
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowBorderStyle=   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.PictureBox PicEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1275
      ScaleHeight     =   225
      ScaleWidth      =   5805
      TabIndex        =   1
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
         Left            =   3525
         MaxLength       =   12
         TabIndex        =   3
         Top             =   -15
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
         TabIndex        =   2
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
         ItemData        =   "frmItems.frx":0000
         Left            =   5220
         List            =   "frmItems.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   -45
         Width           =   600
      End
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
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":0016
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":046A
            Key             =   "ORD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":08C2
            Key             =   "PARIN"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":0D16
            Key             =   "ITMIN"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":116A
            Key             =   "CAT"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":15BE
            Key             =   "OK"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":1A16
            Key             =   "TYPE"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV 
      Height          =   7845
      Left            =   30
      TabIndex        =   0
      Top             =   45
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item ID"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Item"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Unit"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Rate1"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Rate2"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Rate3"
         Object.Width           =   1411
      EndProperty
   End
   Begin MSForms.CommandButton cmdOpt 
      Height          =   390
      Left            =   8700
      TabIndex        =   6
      Top             =   8025
      Width           =   1545
      Caption         =   "Options        "
      PicturePosition =   327683
      Size            =   "2725;688"
      Picture         =   "frmItems.frx":1E6A
      Accelerator     =   79
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
      TabIndex        =   5
      Top             =   8025
      Width           =   1545
      ForeColor       =   64
      Caption         =   "Close          "
      PicturePosition =   327683
      Size            =   "2725;688"
      Picture         =   "frmItems.frx":21EC
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
      Visible         =   0   'False
      Begin VB.Menu mnurefresh 
         Caption         =   "Re&fresh"
      End
      Begin VB.Menu mnuShowAll 
         Caption         =   "&Show Inactive Items"
      End
      Begin VB.Menu mnuLedger 
         Caption         =   "View &Ledger For Last 30 Days"
         Visible         =   0   'False
      End
      Begin VB.Menu dash0 
         Caption         =   "-"
      End
      Begin VB.Menu mnunewcat 
         Caption         =   "Add New Catagory"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuNewAcc 
         Caption         =   "&Add New Item"
      End
      Begin VB.Menu mnuEditAcc 
         Caption         =   "&Edit Item"
      End
      Begin VB.Menu dash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActive 
         Caption         =   "Make This Item Ina&ctive"
      End
      Begin VB.Menu dash3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print Item List"
      End
   End
End
Attribute VB_Name = "frmItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim rec As New ADODB.Recordset
Dim rsAccounts As New ADODB.Recordset
'Dim i As Integer
Dim Bal As Double
Dim AccList As String
Dim SubAccList As String
Dim TotalBalance As Double
Dim PrevIndex As Long
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

Private Sub FillItems()
    
LV.Sorted = False

    Dim rs As New ADODB.Recordset
    Dim Cnd As String, itm As ListItem

    If Not ShowAll Then Cnd = " Where Active=1 " Else Cnd = ""
    
    With rs
        .CursorLocation = adUseClient
        .Open "Select * From Items " & Cnd, Con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            
            
            
            If Not ![Active] Then clr = &H707070 Else clr = LV.ForeColor
            Set itm = LV.ListItems.Add(, ![ItemID] & "'", ![ItemID])
                
                itm.ListSubItems.Add(, , GetlvlStr(![ItemID]) & ![ItemName] & "", , ![ItemName]).ForeColor = clr
                itm.ListSubItems.Add(, , ![Description] & "", , ![Description] & "").ForeColor = clr
                itm.ListSubItems.Add(, , ![Unit]).ForeColor = clr
                itm.ListSubItems.Add(, , ![Rate1]).ForeColor = clr
                itm.ListSubItems.Add(, , ![Rate2]).ForeColor = clr
                itm.ListSubItems.Add(, , ![Rate3]).ForeColor = clr
                itm.ListSubItems.Add
                'If Not IsNull(![Status]) Then
                '  itm.ListSubItems(7).ReportIcon = CStr(![Status])
                '  itm.ListSubItems(7).ToolTipText = "Insufficent Stock"
                'Else
                '  itm.ListSubItems(7).ReportIcon = "OK"
                '  itm.ListSubItems(7).ToolTipText = "Sufficent Stock"
                'End If
                If ![Active] Then itm.Tag = "Active" Else itm.Tag = "Inactive"
               
                .MoveNext
        Next
    .Close
End With

Set rs = Nothing
LV.Visible = True
LV.Sorted = True
If PrevIndex > 0 And PrevIndex <= LV.ListItems.Count Then LV.ListItems(PrevIndex).Selected = True
On Error Resume Next
LV.SelectedItem.EnsureVisible
Call HideWait

End Sub

Private Sub cmbCrDr_Click()
 If cmbCrDr.Text = "CR" Then
  cmbCrDr.Tag = -Val(txtBal.Text)
 Else
  cmbCrDr.Tag = Val(txtBal.Text)
 End If
End Sub

'Private Sub cmbCrDr_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
' Call UpdateAccount
'End If
'End Sub

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
 'SendKeys "{TAB}"
End If

End Sub

Private Sub Form_Load()

Me.KeyPreview = True

Dim Condition As String
If ShowAll Then
   Condition = ""
Else
   Condition = " Where Active=1 "
End If
   'Set LV.SmallIcons = ImageList1
   'rsAccounts.Open "Select * from Accounts " & Condition & " Order By Accno", Con, adOpenForwardOnly
   'Call RefreshList
   Call FillItems
   mnupop.Visible = False
End Sub

Private Sub Form_Resize()
On Error Resume Next

LV.Move 50, 50, Me.ScaleWidth - 100, Me.ScaleHeight - (cmdClose.Height + 100)
cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, LV.Top + LV.Height + 50
'cmdPrint.Move cmdClose.Left - cmdPrint.Width - 50, cmdClose.Top
cmdOpt.Move cmdClose.Left - cmdOpt.Width - 50, cmdClose.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
'rsAccounts.Close
Set rsAccounts = Nothing
End Sub

Private Sub LV_DblClick()
'Exit Sub
If LV.ListItems.Count = 0 Then Exit Sub
Dim f As Form


'If LV.SelectedItem.SmallIcon Like "ITM*" Then
    Set f = New frmNewItem
'Else
'    Set f = New frmCatagories
'End If

With f
    .Add = False
    .Tag = LV.SelectedItem
    .Edit = True
    .Show 1
End With
Set f = Nothing


Exit Sub

With LV

   If .ListItems.Count = 0 Then Exit Sub
   
   If .SelectedItem.Tag = "Inactive" Then
      MsgBox "Inactive Item Cannot Be Edited.", vbInformation
      Exit Sub
   ElseIf .SelectedItem.SmallIcon = "CAT" Then
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

Private Sub LV_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If LV.ListItems.Count = 0 Then Exit Sub
   If LV.SelectedItem.SmallIcon = "TYPE" Then Exit Sub
   Call CheckMenu
   Me.PopupMenu mnupop, , (LV.Left + (LV.Width / 2)), (LV.Top + LV.SelectedItem.Top)
End If
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 2 Or LV.ListItems.Count = 0 Then Exit Sub
If LV.SelectedItem.SmallIcon = "TYPE" Then Exit Sub

 Call CheckMenu
 Me.PopupMenu mnupop
 
End Sub

Private Sub mnuActive_Click()

   If mnuActive.Tag = "Active" Then
      If Not MakeAccountInActive Then Exit Sub
   ElseIf mnuActive.Tag = "Inactive" Then
      If Not MakeAccountActive Then Exit Sub
   End If
   
   Call ShowWait
'   Sleep (500)
   Call mnuRefresh_Click
   Call HideWait
End Sub

Private Sub mnuEditAcc_Click()
   Call LV_DblClick
End Sub




Private Sub mnuNewAcc_Click()
Dim CatID As String
'If LV.SelectedItem.SmallIcon Like "ITM*" Then CatID = Left(LV.SelectedItem, Len(LV.SelectedItem) - 4) Else CatID = LV.SelectedItem
   With frmNewItem
      .Add = True
      '.cmbCatID = CatID
      .Show 1
   End With
   Call mnuRefresh_Click
End Sub

'Private Sub mnunewcat_Click()
'Dim f As New frmCatagories
'   f.Add = True
'   f.Show 1
'   Call mnuRefresh_Click
'End Sub

Private Sub mnuParent_Click()
   If MsgBox("Are You Sure To Make This Account Parent", vbQuestion + vbYesNo) = vbNo Then Exit Sub
   Con.Execute "update accounts set Parent=True Where AccNo='" & LV.SelectedItem & "'"
   Call mnuRefresh_Click
End Sub

Private Sub mnuPrint_Click()

Dim Inactive As Boolean
Call ShowWait
Con.BeginTrans
Con.Execute "Delete From TempAccounts"

 For i = 1 To LV.ListItems.Count
    With LV.ListItems(i)
        If .SmallIcon = "PARIN" Or .SmallIcon = "ACCIN" Then Inactive = True Else Inactive = False
        Con.Execute "Insert into TempAccounts Values('" & .Text & "','" & .SubItems(1) & _
                    "','" & .SubItems(3) & "','" & .ListSubItems(4) & "','" & .SubItems(5) & "','" & .SubItems(6) & "'," & .Bold & ")"
                    
    End With
 Next

Con.CommitTrans
Dim Company As String
Company = getCompanyName
With cr1
    '.DataFiles(0) = DBName
    .Connect = ConnectStr
    .ReportFileName = App.Path & "\ItemList.rpt"
    .Formulas(0) = "Company='" & Company & "'"
    .Formulas(1) = "Caption='Inventory Item List'"
    '.Formulas(2) = "PrintBalance=False"
    .Action = 1
    .PageZoomNext
End With

Call HideWait

End Sub

Private Sub mnuPrintBal_Click()
Dim Inactive As Boolean
Call ShowWait
Con.BeginTrans
Con.Execute "Delete From TempAccounts"
 For i = 1 To LV.ListItems.Count
    With LV.ListItems(i)
        If .SmallIcon = "PARIN" Or .SmallIcon = "ACCIN" Then Inactive = True Else Inactive = False
        Con.Execute "Insert into TempAccounts Values('" & .Text & "','" & .SubItems(1) & _
                    "','" & Left(.Text, 2) & "','" & .ListSubItems(3) & "'," & .Bold & "," & Inactive & ")"
    End With
 Next

Con.CommitTrans

With cr1
    '.DataFiles(0) = DBName
    .Connect = ConnectStr
    .ReportFileName = RptPath & "\Accounts.rpt"
    .Formulas(0) = "Company='" & Company & "'"
    .Formulas(1) = "Caption='Current Balances of Accounts'"
    .Formulas(2) = "PrintBalance=True"
    .Action = 1
    .PageZoomNext
End With

Call HideWait
End Sub

Private Sub mnuRefresh_Click()
   Call ShowWait
   'rsAccounts.Close
   'Set rsAccounts = Nothing
   'Call Form_Load
   Call FillItems
   Call HideWait
End Sub


Private Sub CheckMenu()

    If LV.ListItems.Count = 0 Then
        mnuNewAcc.Enabled = True
        mnuEditAcc.Enabled = False
        'mnuTransfer.Enabled = False
        'mnuJV.Enabled = False
        mnuLedger.Enabled = False
        mnuPrint.Enabled = False
        mnuActive.Enabled = False
        'mnuParent.Enabled = False
    Else
        mnuNewAcc.Enabled = True
        mnuEditAcc.Enabled = True
        'mnuEditAcc.Caption = IIf(LV.SelectedItem.SmallIcon Like "ITM*", "Edit This Item", "Edit This Catagory")
        'mnuTransfer.Enabled = True
        'mnuJV.Enabled = True
        mnuLedger.Enabled = LV.SelectedItem.SmallIcon Like "ITM*"
        mnuActive.Enabled = True
        If LV.SelectedItem.Tag = "Inactive" Then
            mnuEditAcc.Enabled = False
            'mnuTransfer.Enabled = False
            'mnuJV.Enabled = False
            mnuLedger.Enabled = False
            mnuActive.Enabled = True
            'mnuParent.Enabled = False
            mnuActive.Caption = "Make This Item A&ctive."
        Else
            mnuActive.Caption = "Make This Item Ina&ctive."
        End If
        mnuActive.Tag = LV.SelectedItem.Tag
        mnuPrint.Enabled = True
        'mnuPrintBal.Enabled = True
        'mnuActive.Enabled = LV.SelectedItem.SmallIcon Like "ITM*"
    End If



End Sub

Private Sub mnuShowAll_Click()
mnuShowAll.Checked = Not mnuShowAll.Checked
ShowAll = mnuShowAll.Checked
Call mnuRefresh_Click
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
   
   'If Val(.ListSubItems(6)) > 0 Then
   '     MsgBox "This Item Cannot Be Made Inactive." & vbNewLine & _
   '        "To Make An Item Inactive Quantity In Hand Must Be Zero.", vbInformation
   '      Exit Function
   'ElseIf .SmallIcon = "PAR" Then
   '    If MsgBox("This Is A Master Account, All Sub Accounts Will Also Become Inactive." & vbNewLine & _
   '        "Are You Sure to Make This Account Inactive", vbYesNo + vbQuestion) = vbNo Then Exit Function
   'Else
      If MsgBox("Are You Sure to Make This Item Inactive", vbYesNo + vbQuestion) = vbNo Then Exit Function
  ' End If
   
   
   
   
   Con.Execute "Update Items Set Active=false Where ItemID='" & Trim(.Text) & "'"
  
   MakeAccountInActive = True
End With

End Function


Private Function MakeAccountActive() As Boolean

MakeAccountActive = False
Dim AlsoChild As String

 With LV.SelectedItem
   Con.Execute "Update Items Set Active=True Where ItemID='" & Trim(.Text) & "'"
   MakeAccountActive = True
End With




End Function



Private Function UpdateAccount() As Boolean

 If txtAccTitle.Text = "" Then
   MsgBox "Invalid Account Title", vbInformation
   Exit Function
 End If
 
 Dim Diff As Double

   Call ShowWait
   Diff = Val(cmbCrDr.Tag) - Val(txtBal.Tag)
   Con.Execute "Update Accounts Set Acctitle='" & txtAccTitle & "',OpenBal=" & Val(cmbCrDr.Tag) & _
               ",Balance=Balance+" & Diff & " Where AccNo='" & LV.SelectedItem.Text & "'"
               
   'Adjust The Ledger
   'Call adjustLedger(LV.SelectedItem.Text, #1/1/1900#, 0, Diff, True)
   'Sleep (500)
   
   Call mnuRefresh_Click
   
Call HideWait
   
End Function


Private Sub AddHeads()
Dim rs As New ADODB.Recordset
With rs
   '.CursorLocation = adUseClient
   
      .Open "Select * from Items", Con, adOpenForwardOnly, adLockReadOnly
   
Dim itm As ListItem, OpenBal As String, Balance As String
Set LV.SmallIcons = ImageList1
    
    If LV.ListItems.Count > 0 Then PrevIndex = LV.SelectedItem.Index
    
LV.ListItems.Clear
   For i = 0 To .RecordCount - 1
     
      Set itm = LV.ListItems.Add(, ![CatID] & "'", ![CatID], , "CAT")
      
      itm.Bold = True
      
      itm.ListSubItems.Add(, , GetlvlStr(![CatID]) & ![Catagory] & "").Bold = True
      itm.ListSubItems.Add(, , ![Description] & "").Bold = True
      itm.ListSubItems.Add(, , " ").Bold = True
      itm.ListSubItems.Add(, , " ").Bold = True
      itm.ListSubItems.Add(, , " ").Bold = True
      itm.ListSubItems.Add(, , Format(Val(![qty] & ""), "#0")).Bold = True
      .MoveNext
   Next
.Close
End With
Set rs = Nothing
Call FillItems
End Sub


Private Function GetlvlStr(CatID As String) As String
    GetlvlStr = Space(5 * SubStrCount(CatID, "-"))
    'If Len(GetlvlStr) > 0 Then GetlvlStr = GetlvlStr & "» "
    'GetlvlStr = CatID
End Function

