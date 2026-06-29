VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmItems 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item List"
   ClientHeight    =   8595
   ClientLeft      =   150
   ClientTop       =   435
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
      TabIndex        =   4
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
         ItemData        =   "frmItems.frx":0000
         Left            =   5250
         List            =   "frmItems.frx":000A
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
   Begin MSFlexGridLib.MSFlexGrid FGrid 
      Height          =   6390
      Left            =   150
      TabIndex        =   6
      Top             =   150
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   11271
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColor       =   15333882
      BackColorFixed  =   13627635
      ForeColorFixed  =   2125999
      BackColorBkg    =   13627635
      GridColor       =   6244673
      GridColorFixed  =   6244673
      GridLines       =   0
      GridLinesFixed  =   1
      Appearance      =   0
      FormatString    =   "<  Item Code  |<  Description  |<  Unit  |>  FOB  |>  CIF(Air)  |>  CIF(Sea)  |>  C&F(Air)  |>  C&F(Sea) "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSForms.CommandButton cmdOpt 
      Height          =   390
      Left            =   8700
      TabIndex        =   3
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
      Begin VB.Menu mnurefresh 
         Caption         =   "Re&fresh"
      End
      Begin VB.Menu mnuShowAll 
         Caption         =   "&Show Inactive Items"
      End
      Begin VB.Menu mnuStockHistory 
         Caption         =   "View &Stock History For Last 30 Days"
      End
      Begin VB.Menu dash0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewItem 
         Caption         =   "&Add New Item"
      End
      Begin VB.Menu mnuNewCat 
         Caption         =   "Add New Catagory"
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Edit Item"
      End
      Begin VB.Menu dash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActive 
         Caption         =   "Make This Item Ina&ctive"
      End
      Begin VB.Menu mnuParent 
         Caption         =   "Make This Account Paren&t"
         Visible         =   0   'False
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSales 
         Caption         =   "&Enter Sales"
      End
      Begin VB.Menu mnuPurchases 
         Caption         =   "Enter Purchases"
      End
      Begin VB.Menu dash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print Report"
      End
      Begin VB.Menu mnuPrintBal 
         Caption         =   "Print Report(Internal)"
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
        con.Execute "Delete * From TempChaccount"
        con.Execute "Insert into TempChAccount Values('" & AccNo & "','" & AccTitle & "'," & Bal & ")"
    Next i
    
End Sub
Private Sub FillItems()
    
    LV.Sorted = False
    Set LV.SmallIcons = ImageList1
    
    Dim rs As New ADODB.Recordset
    Dim Cnd As String, itm As ListItem

    If Not ShowAll Then Cnd = " Where InActive=0" Else Cnd = ""
    'LV.ListItems.Clear
    
    With rs
        .CursorLocation = adUseClient
        .Open "Select * From Items " & Cnd & " Order By cast(CatID as numeric)", con, adOpenForwardOnly, adLockReadOnly
        
        For i = 0 To .RecordCount - 1
            
            If ![Inactive] Then
               clr = &H707070
            Else
               clr = LV.ForeColor
            End If
            
            
            Set itm = LV.ListItems.Add(, ![ITemID], ![ITemID], , IIf(![Inactive] = 0, "ITM", "ITMIN"))
                
            'itm.ListSubItems.Add(, , Space(5) & GetlvlStr(![CatID]) & ![ItemName] & "", , ![ItemName]).ForeColor = clr
            itm.ListSubItems.Add(, , "     »  " & ![ItemName] & "", , ![ItemName]).ForeColor = clr
            'itm.ListSubItems.Add(, , ![ItemName] & "", , ![ItemName]).ForeColor = clr
            itm.ListSubItems.Add(, , ![Unit] & "").ForeColor = clr
                
            itm.ListSubItems.Add(, , Format(Val(![FOB] & ""), "#0.00"), , "").ForeColor = clr
            itm.ListSubItems.Add(, , Format(Val(![CIFSea] & ""), "#0.00"), , "").ForeColor = clr
            itm.ListSubItems.Add(, , Format(Val(![CIFAir] & ""), "#0.00"), , "").ForeColor = clr
            itm.ListSubItems.Add(, , Format(Val(![CnFSea] & ""), "#0.00"), , "").ForeColor = clr
            itm.ListSubItems.Add(, , Format(Val(![CnFAir] & ""), "#0.00"), , "").ForeColor = clr
                
            'itm.ListSubItems.Add
                

                
                
            If ![Inactive] = 0 Then itm.Tag = "Active" Else itm.Tag = "Inactive"
               
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
 Me.PopupMenu mnupop, 255, .Left, .Top + .Height
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

Private Sub Form_Unload(Cancel As Integer)
'rsAccounts.Close
Set rsAccounts = Nothing
End Sub

Private Sub Lv_DblClick()
    
    If LV.SelectedItem.SmallIcon = "CAT" Then
        Load frmNewItemCat
        Call frmNewItemCat.EditCat(LV.SelectedItem)
    ElseIf LV.SelectedItem.SmallIcon = "ITM" Then
        Load frmNewItem
        Call frmNewItem.EditItem(LV.SelectedItem)
    End If
    
    
    Exit Sub

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

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button <> 2 Then Exit Sub
    If LV.ListItems.Count > 0 Then If LV.SelectedItem.SmallIcon = "TYPE" Then Exit Sub

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
   Call mnuRefresh_Click
   Screen.MousePointer = vbDefault
End Sub

Private Sub MnuEditItem_Click()
   Call Lv_DblClick
End Sub

Private Sub mnuPurchases_Click()
frmItemPurchase.Show
frmItemPurchase.SetFocus
End Sub

Private Sub mnuStockHistory_Click()
'Dim f As New frmItemLedger
With f
   .Show
   Call FindInCombo(.cmbItemID, LV.SelectedItem)
   Call .cmdRefresh_Click
End With
End Sub

Private Sub mnuNewItem_Click()

    With frmNewItem
        Load frmNewItem
        .Show 1
    End With
    Call mnuRefresh_Click
End Sub

Private Sub mnunewcat_Click()
    Load frmNewItemCat
    frmNewItemCat.Show 1
End Sub

Private Sub mnuParent_Click()
   If MsgBox("Are You Sure To Make This Account Parent", vbQuestion + vbYesNo) = vbNo Then Exit Sub
   con.Execute "update accounts set Parent=True Where AccNo='" & LV.SelectedItem & "'"
   Call mnuRefresh_Click
End Sub

Private Sub mnuPrint_Click()

    
    Screen.MousePointer = vbHourglass

    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VrptDMR.ItemID}='" & LV.SelectedItem & "'", rptDMR
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuPrintBal_Click()

    Screen.MousePointer = vbHourglass
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VrptDMR.ItemID}='" & LV.SelectedItem & "'", rptDMRInternal
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

    If LV.ListItems.Count = 0 Then
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
        mnuEditItem.Caption = IIf(LV.SelectedItem.SmallIcon Like "ITM*", "Edit This Item", "Edit This Catagory")
        mnuSales.Enabled = True
        mnuPurchases.Enabled = True
        mnuStockHistory.Enabled = LV.SelectedItem.SmallIcon Like "ITM*"
     
        If LV.SelectedItem.Tag = "Inactive" Then
            mnuEditItem.Enabled = False
            mnuSales.Enabled = False
            mnuPurchases.Enabled = False
            mnuStockHistory.Enabled = False
            mnuParent.Enabled = False
            mnuActive.Caption = "Make This Item A&ctive."
        Else
            mnuActive.Caption = "Make This Item Ina&ctive."
        End If
        mnuActive.Tag = LV.SelectedItem.Tag
        
        mnuPrint.Enabled = LV.SelectedItem.SmallIcon Like "ITM*"
        mnuPrintBal.Enabled = LV.SelectedItem.SmallIcon Like "ITM*"
        
        mnuActive.Enabled = LV.SelectedItem.SmallIcon Like "ITM*"
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

    On Error GoTo ERR
    Dim rs As New ADODB.Recordset
    Dim itm As ListItem
    
    With rs
        .Open "Select * from ItemCatagories Order By Cast(CatID as Numeric)", con, adOpenForwardOnly, adLockReadOnly
        Set LV.SmallIcons = ImageList1
        If LV.ListItems.Count > 0 Then PrevIndex = LV.SelectedItem.Index
    
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set itm = LV.ListItems.Add(, ![CatID] & "'", ![CatID], , "CAT")
            itm.Bold = True
            itm.ListSubItems.Add(, , ![Description] & "").Bold = True
                       
            'ITM.ListSubItems.Add(, , Format(Val(![Qty] & ""), "#0")).Bold = True
            .MoveNext
        Next
        .Close
    End With
    
    Set rs = Nothing
    Call FillItems
    
    Exit Sub
ERR:
    MsgBox ERR.Description
End Sub


Private Function GetlvlStr(CatID As String) As String
    GetlvlStr = Space(5 * SubStrCount(CatID, "-"))
    If Len(GetlvlStr) > 0 Then GetlvlStr = GetlvlStr & "» "
End Function
