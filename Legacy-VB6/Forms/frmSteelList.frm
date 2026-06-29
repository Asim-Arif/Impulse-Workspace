VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSteelList 
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   -120
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8310
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
      TabIndex        =   5
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
         TabIndex        =   1
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
         TabIndex        =   0
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
         ItemData        =   "frmSteelList.frx":0000
         Left            =   5250
         List            =   "frmSteelList.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   2
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
            Picture         =   "frmSteelList.frx":0016
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSteelList.frx":046A
            Key             =   "ORD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSteelList.frx":08C2
            Key             =   "PARIN"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSteelList.frx":0D16
            Key             =   "ITMIN"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSteelList.frx":116A
            Key             =   "CAT"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSteelList.frx":15BE
            Key             =   "OK"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSteelList.frx":1A16
            Key             =   "TYPE"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV 
      Height          =   7845
      Left            =   60
      TabIndex        =   3
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   9701
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Unit"
         Object.Width           =   882
      EndProperty
   End
   Begin MSForms.CommandButton cmdOpt 
      Height          =   390
      Left            =   8700
      TabIndex        =   4
      Top             =   8025
      Width           =   1545
      Caption         =   "Options        "
      PicturePosition =   327683
      Size            =   "2725;688"
      Picture         =   "frmSteelList.frx":1E6A
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
      TabIndex        =   6
      Top             =   8025
      Width           =   1545
      ForeColor       =   64
      Caption         =   "Close          "
      PicturePosition =   327683
      Size            =   "2725;688"
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
         Visible         =   0   'False
      End
      Begin VB.Menu mnuStockHistory 
         Caption         =   "View &Stock History For Last 30 Days"
         Visible         =   0   'False
      End
      Begin VB.Menu dash0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewItem 
         Caption         =   "&Add New Guage"
      End
      Begin VB.Menu mnuNewCat 
         Caption         =   "Add New Steel"
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Edit Guage"
      End
      Begin VB.Menu dash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActive 
         Caption         =   "Make This Item Ina&ctive"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuParent 
         Caption         =   "Make This Account Paren&t"
         Visible         =   0   'False
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSales 
         Caption         =   "&Enter Sales"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPurchases 
         Caption         =   "Enter Purchases"
         Visible         =   0   'False
      End
      Begin VB.Menu dash3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print Item Price List"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintBal 
         Caption         =   "Print Current &Stock Status"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmSteelList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rec As New ADODB.Recordset
Dim rsAccounts As New ADODB.Recordset
'Dim i As Integer
Dim Bal As Double
Dim AccList As String
Dim SubAccList As String
Dim TotalBalance As Double
Dim PrevIndex As Long
Dim ShowAll As Boolean

Private Sub FillItems(ID As Integer)
    
    'LV.Sorted = False
    Set LV.SmallIcons = ImageList1
    
    Dim rs As New ADODB.Recordset
    Dim Cnd As String, ITM As ListItem

    'If Not ShowAll Then Cnd = " Where InActive=0" Else Cnd = ""
    'LV.ListItems.Clear
    
    With rs
        .CursorLocation = adUseClient
        .Open "Select * From SteelGages Where SteelID=" & ID & " Order By Cast(Substring(GageID,charindex('-',GageID)+1,50) as int)", con, adOpenForwardOnly, adLockReadOnly
        
        For i = 0 To .RecordCount - 1
            
            'Set ITM = LV.ListItems.Add(, ![GageID] & "g", ![GageID], , "ITM")
            Set ITM = LV.ListItems.add(, ![GageID] & "g", ![GageID], , "ITM")
            
            'itm.ListSubItems.Add(, , Space(5) & GetlvlStr(![CatID]) & ![ItemName] & "", , ![ItemName]).ForeColor = clr
            Call ITM.ListSubItems.add(, , "     »  " & ![Gage] & "", , ![Gage])
            'itm.ListSubItems.Add(, , ![ItemName] & "", , ![ItemName]).ForeColor = clr
            Call ITM.ListSubItems.add(, , ![GageUnit] & "")
                
            'itm.ListSubItems.Add(, , Format(Val(![FOB] & ""), "#0.00"), , "").ForeColor = clr
            'itm.ListSubItems.Add(, , Format(Val(![CIFSea] & ""), "#0.00"), , "").ForeColor = clr
            'itm.ListSubItems.Add(, , Format(Val(![CIFAir] & ""), "#0.00"), , "").ForeColor = clr
            'itm.ListSubItems.Add(, , Format(Val(![CnFSea] & ""), "#0.00"), , "").ForeColor = clr
            'itm.ListSubItems.Add(, , Format(Val(![CnFAir] & ""), "#0.00"), , "").ForeColor = clr
                
            'itm.ListSubItems.Add
                

                
                
            'If ![Inactive] = 0 Then itm.Tag = "Active" Else itm.Tag = "Inactive"
               
            .MoveNEXT
        Next
        .Close
    End With
    Set rs = Nothing
    LV.Visible = True
    'LV.Sorted = True
    If PrevIndex > 0 And PrevIndex <= LV.ListItems.count Then LV.ListItems(PrevIndex).Selected = True
    On Error Resume Next
   LV.SelectedItem.EnsureVisible
    Screen.MousePointer = vbDefault

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
 Me.PopUpMenu mnupop, 255, .Left, .Top + .Height
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
   Condition = " Where Active "
End If
   'Set LV.SmallIcons = ImageList1
   'rsAccounts.Open "Select * from Accounts " & Condition & " Order By Accno", Con, adOpenForwardOnly
   'Call RefreshList
   Call AddHeads
   
End Sub

Private Sub Form_Resize()
On Error Resume Next

LV.Move 50, 50, Me.ScaleWidth - 100, Me.ScaleHeight - (cmdClose.Height + 100)
cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, LV.Top + LV.Height + 50
'cmdPrint.Move cmdClose.Left - cmdPrint.Width - 50, cmdClose.Top
cmdOpt.Move cmdClose.Left - cmdOpt.Width - 50, cmdClose.Top
End Sub

Private Sub Form_Unload(cancel As Integer)
'rsAccounts.Close
Set rsAccounts = Nothing
End Sub

Private Sub LV_DblClick()
    
    If LV.SelectedItem.SmallIcon = "CAT" Then
        Load frmNewSteel
        Call frmNewSteel.EditSteel(Val(LV.SelectedItem.key))
    ElseIf LV.SelectedItem.SmallIcon = "ITM" Then
        Load frmGages
        Call frmGages.EditGage((LV.SelectedItem))
    End If
    
    
    Exit Sub

End Sub

Private Sub LV_GotFocus()
   If PicEdit.Visible Then Call Form_KeyPress(27)
End Sub

Private Sub LV_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If LV.ListItems.count = 0 Then Exit Sub
   If LV.SelectedItem.SmallIcon = "TYPE" Then Exit Sub
   Call CheckMenu
   Me.PopUpMenu mnupop, , (LV.Left + (LV.Width / 2)), (LV.Top + LV.SelectedItem.Top)
End If
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 2 Then Exit Sub
If LV.ListItems.count > 0 Then If LV.SelectedItem.SmallIcon = "TYPE" Then Exit Sub

 Call CheckMenu
 Me.PopUpMenu mnupop
 
End Sub

Private Sub mnuActive_Click()

   If mnuActive.Tag = "Active" Then
      If Not MakeAccountInActive Then Exit Sub
   ElseIf mnuActive.Tag = "Inactive" Then
      If Not MakeAccountActive Then Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
'   Sleep (500)
   Call mnuRefresh_Click
   Screen.MousePointer = vbDefault
End Sub

Private Sub MnuEditItem_Click()
   Call LV_DblClick
End Sub

Private Sub mnuPurchases_Click()
    frmItemPurchase.Show
    frmItemPurchase.SetFocus
End Sub

Private Sub mnuStockHistory_Click()
Exit Sub
'Dim f As New frmItemLedger
With f
   .Show
   Call FindInCombo(.cmbItemID, LV.SelectedItem)
   Call .cmdRefresh_Click
End With
End Sub

Private Sub mnuNewItem_Click()

    
    Load frmNewItem
    frmGages.Show 1
    
    Call mnuRefresh_Click
    
End Sub

Private Sub mnunewcat_Click()
    Load frmNewSteel
    frmNewSteel.Show 1
End Sub

Private Sub mnuParent_Click()
   If MsgBox("Are You Sure To Make This Account Parent", vbQuestion + vbYesNo) = vbNo Then Exit Sub
   con.Execute "update accounts set Parent=True Where AccNo='" & LV.SelectedItem & "'"
   Call mnuRefresh_Click
End Sub

Private Sub mnuPrint_Click()
    On Error GoTo err
    Dim Inactive As Boolean
    Screen.MousePointer = vbHourglass
    
    With cr1
        .Connect = ConnectStr
        .ReportFileName = RptPath & "\ItemList.rpt"
        .Formulas(0) = "Company='" & strCompany & "'"
        .Formulas(1) = "address='" & COMPANYADDRESS & "'"
        .Formulas(2) = "ReportType='Item Price List'"
        .SelectionFormula = IIf(mnuActive.Checked, "", "{Vitems.active}=True")
        .Action = 1
        .PageZoomNext
    End With
    
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    MsgBox err.Description
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuPrintBal_Click()

    Exit Sub
    Dim Inactive As Boolean
    Screen.MousePointer = vbHourglass
    
    
    With cr1
        .Connect = ConnectStr
        .ReportFileName = RptPath & "\ItemList.rpt"
        .Formulas(0) = "Company='" & strCompany & "'"
        .Formulas(1) = "address='" & COMPANYADDRESS & "'"
        .Formulas(2) = "ReportType='Stock Status'"
        .SelectionFormula = IIf(mnuActive.Checked, "", "{Vitems.active}=True")
        .Action = 1
        .PageZoomNext
    End With
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRefresh_Click()
   Screen.MousePointer = vbHourglass
   'rsAccounts.Close
   'Set rsAccounts = Nothing
   'Call Form_Load
   Call AddHeads
   'Call FillItems
   Screen.MousePointer = vbDefault
End Sub


Private Sub CheckMenu()

   If LV.ListItems.count = 0 Then
     mnuNewItem.Enabled = False
     mnuEditItem.Enabled = False
     mnuSales.Enabled = False
     mnuPurchases.Enabled = False
     mnuStockHistory.Enabled = False
     mnuPrint.Enabled = False
     mnuParent.Enabled = False
     mnuNewCat.Enabled = True
   Else
     mnuNewItem.Enabled = True
     mnuEditItem.Caption = IIf(LV.SelectedItem.SmallIcon Like "ITM*", "Edit This Guage", "Edit This Steel")
     mnuSales.Enabled = True
     mnuPurchases.Enabled = True
     mnuStockHistory.Enabled = LV.SelectedItem.SmallIcon Like "ITM*"
     
     
     
     If LV.SelectedItem.Tag = "Inactive" Then
         mnuEditItem.Enabled = False
         mnuSales.Enabled = False
         mnuPurchases.Enabled = False
         mnuStockHistory.Enabled = False
         mnuParent.Enabled = False
        'mnuActive.Caption = "Make This Item A&ctive."
     Else
       'mnuActive.Caption = "Make This Item Ina&ctive."
     End If
     'mnuActive.Tag = LV.SelectedItem.Tag
     mnuPrint.Enabled = True
     mnuPrintBal.Enabled = True
     'mnuActive.Enabled = LV.SelectedItem.SmallIcon Like "ITM*"
   End If



End Sub

Private Sub mnuShowAll_Click()
    mnuShowAll.Checked = Not mnuShowAll.Checked
    ShowAll = mnuShowAll.Checked
    Call mnuRefresh_Click
End Sub

Private Sub mnuSales_Click()
   frmItemSale.Show
   frmItemSale.SetFocus
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
   
   
   
   
   con.Execute "Update Items Set InActive=1 Where ItemID='" & Trim(.Text) & "'"
  
   MakeAccountInActive = True
End With

End Function


Private Function MakeAccountActive() As Boolean

MakeAccountActive = False
Dim AlsoChild As String

 With LV.SelectedItem
   con.Execute "Update Items Set InActive=0 Where ItemID='" & Trim(.Text) & "'"
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
   con.Execute "Update Accounts Set Acctitle='" & txtAccTitle & "',OpenBal=" & Val(cmbCrDr.Tag) & _
               ", Balance=Balance+" & Diff & " Where AccNo='" & LV.SelectedItem.Text & "'"
               
   'Adjust The Ledger
   Call adjustLedger(LV.SelectedItem.Text, #1/1/1900#, 0, Diff, True)
   
   'Sleep (500)
   Call mnuRefresh_Click
Screen.MousePointer = vbDefault
   
End Function



Private Sub AddHeads()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "Select * from SteelTypes Order By SteelID", con, adOpenForwardOnly, adLockReadOnly
        Set LV.SmallIcons = ImageList1
        If LV.ListItems.count > 0 Then PrevIndex = LV.SelectedItem.Index
    
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set ITM = LV.ListItems.add(, ![SteelID] & "'", ![SteelID], , "CAT")
            ITM.Bold = True
            ITM.ListSubItems.add(, , ![SteelType] & "").Bold = True
                       
            'ITM.ListSubItems.Add(, , Format(Val(![Qty] & ""), "#0")).Bold = True
            Call FillItems(![SteelID])
            .MoveNEXT
        Next
        .Close
    End With
    
    Set rs = Nothing
    
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Function GetlvlStr(CatID As String) As String
    GetlvlStr = Space(5 * SubStrCount(CatID, "-"))
    If Len(GetlvlStr) > 0 Then GetlvlStr = GetlvlStr & "» "
End Function
