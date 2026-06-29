VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmHICOR 
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin ComboList.Usercontrol1 cmbGroup 
      Height          =   285
      Left            =   75
      TabIndex        =   3
      Top             =   270
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
      Appearance      =   0
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
            Picture         =   "frmHICOR.frx":0000
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHICOR.frx":0454
            Key             =   "ORD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHICOR.frx":08AC
            Key             =   "CAT1"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHICOR.frx":0D00
            Key             =   "ITMIN"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHICOR.frx":1154
            Key             =   "CAT"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHICOR.frx":15A8
            Key             =   "OK"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHICOR.frx":1A00
            Key             =   "TYPE"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV 
      Height          =   7410
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   13070
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
      Appearance      =   0
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
         Text            =   "Device Code"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Device Description"
         Object.Width           =   10583
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Assigned To"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Rate"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFBDCB&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Group"
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
      Index           =   0
      Left            =   75
      TabIndex        =   4
      Top             =   30
      Width           =   1905
   End
   Begin MSForms.CommandButton cmdOpt 
      Height          =   390
      Left            =   8700
      TabIndex        =   1
      Top             =   8025
      Width           =   1545
      Caption         =   "Options        "
      PicturePosition =   327683
      Size            =   "2725;688"
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
      TabIndex        =   2
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
      Begin VB.Menu dash0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewItem 
         Caption         =   "&Add New Hico Resource"
      End
      Begin VB.Menu mnuNewCat 
         Caption         =   "Add New Hico Resource Group"
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Edit Hico Resource"
      End
      Begin VB.Menu dash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAssign 
         Caption         =   "Assign This Resource"
      End
      Begin VB.Menu d6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActive 
         Caption         =   "Make This HICO Resource Ina&ctive"
      End
      Begin VB.Menu dash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print Report"
      End
   End
End
Attribute VB_Name = "frmHICOR"
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
Private Sub FillItems(strOrderCond As String, Optional startIndx As Long)
    
    'LV.Sorted = False
    Set LV.SmallIcons = ImageList1
    'If startIndx = 0 Then
    '    startIndx = LV.ListItems.Count
    'End If
    startIndx = 0
    Dim rs As New ADODB.Recordset
    Dim Cnd As String, ITM As ListItem

    'If Not ShowAll Then Cnd = " Where InActive=0 and GroupID=" & CatID Else Cnd = " Where GroupID=" & CatID
    
    If cmbGroup.ID <> "0" Then
        Cnd = Cnd & " Where GroupID=" & cmbGroup.ID & ""
    End If
    
    If Not ShowAll Then
        If Cnd = "" Then
            Cnd = " Where InActive=0"
        Else
            Cnd = Cnd & " and InActive=0"
        End If
    End If
    
    'LV.ListItems.Clear
    If strOrderCond = "" Then
        strOrderCond = " Order By RID1"
    End If
    
    With rs
        .CursorLocation = adUseClient
        .Open "Select * From VHR " & Cnd & strOrderCond, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            
            If ![Inactive] Then
               clr = &H707070
            Else
               clr = ![GroupColor]
            End If
            
            Set ITM = LV.ListItems.Add(startIndx + 1, ![RID] & "HR'", ![RID1], , IIf(![Inactive] = 0, "ITM", "ITMIN"))
            
            ITM.ListSubItems.Add(, , ![RName]).ForeColor = clr
            
            If ![Inactive] = 0 Then ITM.Tag = "Active" Else ITM.Tag = "Inactive"
            
            If IsNull(![VendID]) Then
                ITM.SubItems(2) = "<UnAssigned>"
            Else
                ITM.ListSubItems.Add(, , ![VenderName] & "").ForeColor = vbBlack
                ITM.SubItems(3) = Val(![Rate] & "")
            End If
            
            .MoveNext
            startIndx = startIndx + 1
        Next
        .Close
    End With
    Set rs = Nothing
    LV.Visible = True
    'LV.Sorted = True
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

Private Sub cmbGroup_matched()
    Call AddHeads
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
    mnupop.Visible = False
    
    Dim Condition As String
    If ShowAll Then
        Condition = ""
    Else
        Condition = " Where Active=1 "
    End If
    'Set LV.SmallIcons = ImageList1
    'rsAccounts.Open "Select * from Accounts " & Condition & " Order By Accno", Con, adOpenForwardOnly
    'Call RefreshList
    cmbGroup.ListHeight = 2000
    Call cmbGroup.AddItem("<All Groups>", "0")
    Call cmbGroup.AddVals(con, "Description", "HRGroups", "ID")
    cmbGroup.ID = "0"
    
    'Call AddHeads
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    'LV.Move 50, 50, Me.ScaleWidth - 100, Me.ScaleHeight - (cmdClose.Height + 100)
    LV.Move 50, 600, Me.ScaleWidth - 100, Me.ScaleHeight - (cmdClose.Height + 100 + cmbGroup.Height + cmbGroup.Top)
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, LV.Top + LV.Height + 50
    'cmdPrint.Move cmdClose.Left - cmdPrint.Width - 50, cmdClose.Top
    cmdOpt.Move cmdClose.Left - cmdOpt.Width - 50, cmdClose.Top
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'rsAccounts.Close
Set rsAccounts = Nothing
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Select Case ColumnHeader.Index
        Case 1
            Call AddHeads
        Case 3
            Call AddHeads(" Order By Description")
    End Select
End Sub

Private Sub LV_DblClick()
    
    If LV.SelectedItem.SmallIcon = "CAT" Or LV.SelectedItem.SmallIcon = "CAT1" Then
        'Load frmNewItemCat
        'Call frmNewItemCat.EditCat(Val(LV.SelectedItem.Key))
        If LV.SelectedItem.SmallIcon = "CAT" Then
            LV.SelectedItem.SmallIcon = "CAT1"
            For i = LV.SelectedItem.Index + 1 To LV.ListItems.Count
                If Not LV.ListItems(LV.SelectedItem.Index + 1).SmallIcon Like "CAT*" Then
                    LV.ListItems.Remove LV.SelectedItem.Index + 1
                Else
                    Exit For
                End If
            Next i
        Else
            'LV.SelectedItem.SmallIcon = "CAT"
            'Call FillItems(Val(LV.SelectedItem.Key), " Order By RMName", LV.SelectedItem.Index)
        End If
        
    ElseIf LV.SelectedItem.SmallIcon = "ITM" Then
        Load frmNewHR
        Call frmNewHR.EditItem(Left(LV.SelectedItem.Key, Len(LV.SelectedItem.Key) - 3))
    End If
    
    Exit Sub

End Sub

Private Sub LV_GotFocus()
   'If PicEdit.Visible Then Call Form_KeyPress(27)
End Sub

Private Sub LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If LV.ListItems.Count = 0 Then Exit Sub
       If LV.SelectedItem.SmallIcon = "TYPE" Then Exit Sub
       Call CheckMenu
       Me.PopUpMenu mnupop, , (LV.Left + (LV.Width / 2)), (LV.Top + LV.SelectedItem.Top)
    End If
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button <> 2 Then Exit Sub
    If LV.ListItems.Count > 0 Then If LV.SelectedItem.SmallIcon = "TYPE" Then Exit Sub

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

Private Sub mnuAssign_Click()
    
    If mnuAssign.Caption = "Assign This Resource" Then
        Load frmVendAssHR
        frmVendAssHR.ShowForm (Val(LV.SelectedItem.Key))
    Else
        Call StartTrans(con)
        con.Execute "Update HICOResources Set InUse=0 Where RID=" & Val(LV.SelectedItem.Key)
        con.Execute "Update VendAssResources Set InUse=0,UnAssDT='" & ServerDate & _
         "' Where RID=" & Val(LV.SelectedItem.Key) & " AND InUse=1"
        con.CommitTrans
        
        
    End If
    
End Sub

Private Sub MnuEditItem_Click()
    If LV.SelectedItem.SmallIcon = "CAT" Or LV.SelectedItem.SmallIcon = "CAT1" Then
        Load frmNewRMCat
        Call frmNewRMCat.EditCat(Val(LV.SelectedItem.Key))
    ElseIf LV.SelectedItem.SmallIcon = "ITM" Then
        Load frmNewHR
        Call frmNewHR.EditItem(Left(LV.SelectedItem.Key, Len(LV.SelectedItem.Key) - 3))
    End If
End Sub

Private Sub mnuPurchases_Click()
    frmItemPurchase.Show
    frmItemPurchase.SetFocus
End Sub

Private Sub mnuStockHistory_Click()
    'Dim f As New frmItemLedger
    Exit Sub
    With F
       .Show
       Call FindInCombo(.cmbItemID, LV.SelectedItem)
       Call .cmdRefresh_Click
    End With
End Sub

Private Sub mnuNewItem_Click()

    With frmNewHR
        Load frmNewHR
        Call .AddNew
    End With
    Call mnuRefresh_Click
    
End Sub

Private Sub mnunewcat_Click()
    Dim F As New frmItemGroups
    Load F
    Call F.ShowForm("RMGroups", "RM Groups", "RM Name")
End Sub

Private Sub mnuParent_Click()
   If MsgBox("Are You Sure To Make This Account Parent", vbQuestion + vbYesNo) = vbNo Then Exit Sub
   con.Execute "update accounts set Parent=True Where AccNo='" & LV.SelectedItem & "'"
   Call mnuRefresh_Click
End Sub

Private Sub mnuPrint_Click()
    Call ShowHRrpt
End Sub

Private Sub mnuPrintBal_Click()

    Screen.MousePointer = vbHourglass
    Load frmPrevRpt
    Dim rpt As craxddrt.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptDMRInternal.rpt") '
    frmPrevRpt.ShowReport "{VrptDMR.RMID}='" & LV.SelectedItem.Key & "'", rpt
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
        'mnuNewItem.Enabled = False
        mnuEditItem.Enabled = False
        mnuPrint.Enabled = False
        mnuAssign.Enabled = False
        mnuNewCat.Enabled = True
    Else
        mnuNewItem.Enabled = True
        mnuEditItem.Caption = IIf(LV.SelectedItem.SmallIcon Like "ITM*", "Edit This Raw Material", "Edit This Catagory")
        mnuAssign.Enabled = True
        
        If LV.SelectedItem.Tag = "Inactive" Then
            mnuEditItem.Enabled = False
            mnuActive.Caption = "Make This Raw Material A&ctive."
        Else
            mnuActive.Caption = "Make This Raw Material Ina&ctive."
        End If
        mnuActive.Tag = LV.SelectedItem.Tag
        
        mnuPrint.Enabled = LV.SelectedItem.SmallIcon Like "ITM*"

        
        mnuActive.Enabled = LV.SelectedItem.SmallIcon Like "ITM*"
        If LV.SelectedItem.SubItems(2) = "<UnAssigned>" Or LV.SelectedItem.SubItems(2) = "" Then
            mnuAssign.Caption = "Assign This Resource"
        Else
            mnuAssign.Caption = "UnAssign This Resource"
        End If
    End If
    
End Sub

Private Sub mnuShowAll_Click()
    mnuShowAll.Checked = Not mnuShowAll.Checked
    ShowAll = mnuShowAll.Checked
    Call mnuRefresh_Click
End Sub

Private Sub mnuSales_Click()
   'frmItemSale.Show
   'frmItemSale.SetFocus
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
      If MsgBox("Are You Sure to Make This Raw Material Inactive", vbYesNo + vbQuestion) = vbNo Then Exit Function
  ' End If
   
   
   
   
   con.Execute "Update RM Set InActive=1 Where RMID='" & Trim(.Text) & "'"
  
   MakeAccountInActive = True
End With

End Function


Private Function MakeAccountActive() As Boolean

MakeAccountActive = False
Dim AlsoChild As String

 With LV.SelectedItem
   con.Execute "Update RM Set InActive=0 Where RMID='" & Trim(.Text) & "'"
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



Public Sub AddHeads(Optional strOrderCond As String)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim lColor As Long, lRed As Long, lGreen As Long, lBlue As Long
    
    Dim ItemOrdCond As String
    
    If strOrderCond = "" Then
        strOrderCond = " Order By CatID "
        ItemOrdCond = ""
    Else
        ItemOrdCond = " Order By RMName"
    End If
    
    Call FillItems(ItemOrdCond)
    Exit Sub
    With rs
        .Open "Select * from RMCatagories " & " Order By Cast(CatID as Numeric) ", con, adOpenForwardOnly, adLockReadOnly
        Set LV.SmallIcons = ImageList1
        If LV.ListItems.Count > 0 Then PrevIndex = LV.SelectedItem.Index
    
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            '
            'Randomize
            'lRed = Int(Rnd * (255 - (0 - 1)) + (0))
            'lGreen = Int(Rnd * (255 - (0 - 1)) + (0))
            'lBlue = Int(Rnd * (255 - (0 - 1)) + (0))
            'lColor = RGB(lRed, lGreen, lBlue)
            'itm.SubItems(6) = Int(Rnd * (iMax - (iMin - 1)) + (iMin))
            lColor = vbBlack
            Set ITM = LV.ListItems.Add(, ![CatID] & "'", ![Description], , "CAT")
            ITM.Bold = True
            ITM.ForeColor = lColor
            ITM.ListSubItems.Add
            'ITM.ListSubItems.Add(, , ![Description] & "").Bold = True
            'ITM.ListSubItems(1).ForeColor = lColor
            'ITM.ListSubItems.Add(, , Format(Val(![Qty] & ""), "#0")).Bold = True
            'Call FillItems(![CatID], ItemOrdCond)
            .MoveNext
        Next
        .Close
    End With
    
    Set rs = Nothing
    'Call FillItems
    Exit Sub
    
err:
    MsgBox err.Description
    
End Sub


Private Function GetlvlStr(CatID As String) As String
    GetlvlStr = Space(5 * SubStrCount(CatID, "-"))
    If Len(GetlvlStr) > 0 Then GetlvlStr = GetlvlStr & "» "
End Function
