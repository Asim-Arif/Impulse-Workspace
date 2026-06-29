VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmFCustOrdStatusComp 
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleMode       =   0  'User
   ScaleWidth      =   10641.08
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10560
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Order Status"
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
         Index           =   8
         Left            =   75
         TabIndex        =   5
         Top             =   165
         Width           =   10605
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Order Status"
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
         Index           =   9
         Left            =   90
         TabIndex        =   6
         Top             =   180
         Width           =   10590
      End
   End
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   6555
      Left            =   15
      TabIndex        =   0
      Top             =   630
      Width           =   11850
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   2820
         Top             =   2700
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   11
         ImageHeight     =   10
         MaskColor       =   255
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFCustOrdStatusComp.frx":0000
               Key             =   "Down"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFCustOrdStatusComp.frx":01BC
               Key             =   "Up"
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2475
         TabIndex        =   7
         Top             =   1605
         Visible         =   0   'False
         Width           =   3615
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   195
         TabIndex        =   2
         Top             =   345
         Width           =   2310
         _ExtentX        =   4075
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
         Left            =   2010
         Top             =   2070
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin Crystal.CrystalReport Cr1 
         Left            =   270
         Top             =   2040
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5355
         Left            =   195
         TabIndex        =   10
         Top             =   660
         Width           =   11565
         _ExtentX        =   20399
         _ExtentY        =   9446
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
            Text            =   "S. #"
            Object.Width           =   1014
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Order Qty"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSForms.ComboBox cmbOrders 
         Height          =   285
         Left            =   4425
         TabIndex        =   9
         Top             =   345
         Width           =   1860
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3281;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   375
         Left            =   10110
         TabIndex        =   8
         Top             =   6105
         Width           =   1650
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "2910;661"
         Picture         =   "frmFCustOrdStatusComp.frx":0378
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   2535
         TabIndex        =   3
         Top             =   345
         Width           =   1860
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3281;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         Caption         =   " Customer                                   Country                           Order No"
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
         Left            =   180
         TabIndex        =   1
         Top             =   120
         Width           =   6105
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PopMenu"
   End
End
Attribute VB_Name = "frmFCustOrdStatusComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()
    On Error GoTo err
    
    If cmbCust.ID <> "0" Then
        Cond = " Where CustCode='" & cmbCust.ID & "' "
    Else
        Cond = ""
    End If
    
    If cmbCountry.ListIndex <> 0 Then
        If Cond = "" Then
            Cond = " Where Country='" & cmbCountry & "'"
        Else
            Cond = Cond & " and Country='" & cmbCountry & "'"
        End If
    End If
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select OrderNo From FCustomerOrders " & Cond, con, adOpenForwardOnly, adLockReadOnly
        cmbOrders.Clear
        Do Until .EOF
            cmbOrders.AddItem .Fields(0) & ""
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
        
End Sub


Private Sub cmbOrders_Change()
    Call cmbOrders_Click
End Sub

Private Sub cmbOrders_Click()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim itm As ListItem
    With rs
        .Open "Select * From VrptOrders Where OrderNo='" & cmbOrders.Text & "' Order By ID", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set itm = LV.ListItems.Add(, ![CompItemID] & "'", LV.ListItems.Count + 1)
            Call itm.ListSubItems.Add(, , ![CompItemID] & "")
            Call itm.ListSubItems.Add(, , ![Description] & "")
            Call itm.ListSubItems.Add(, , ![Qty] & "")
            'ITM.SubItems(4) = ![] & ""
            For i = 4 To LV.ColumnHeaders.Count - 3
                Call itm.ListSubItems.Add(, , "N.A")
                itm.ListSubItems(i).ForeColor = RGB(255, 100, 100)
            Next i
            itm.ListSubItems.Add
            itm.ListSubItems.Add
            .MoveNext
        Loop
        .Close
        
        Dim iIndex As Integer
        .Open "Select * From VFOrdItemsStatusComp Where OrderNo='" & cmbOrders.Text & "' Order By CompItemCode", con, adOpenForwardOnly, adLockReadOnly
        'RGB(0, 100, 255)
        Do Until .EOF
            Set itm = LV.ListItems(![CompItemCode] & "'")
            iIndex = LV.ColumnHeaders(![ProcessID] & "'").Index - 1
            itm.ListSubItems(iIndex).Text = Val(![TotQty] & "")
            itm.ListSubItems(iIndex).ForeColor = RGB(0, 100, 255)
            'ITM.SubItems(LV.ColumnHeaders(![ProcessID] & "'").Index - 1) = Val(![TotQty] & "")
            itm.SubItems(LV.ColumnHeaders.Count - 2) = Val(![NotIssued] & "")
            itm.SubItems(LV.ColumnHeaders.Count - 1) = Val(![TotalWastage] & "")
            
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub cmbCust_matched()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    Sql = "Select Distinct Country From ForeignCustomers "
    If cmbCust.ID <> "0" Then
        Sql = Sql & " Where CustCode='" & cmbCust.ID & "'"
    End If
    
    With rs
        .Open Sql, con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
        cmbCountry.AddItem "<All Countries>"
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            'cmbCountry.List(cmbCountry.ListCount - 1, 1) = ![Curr] & ""
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdClose_Click()
    Me.Hide
    Unload Me
End Sub


Private Sub Form_Load()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim CHeader As ColumnHeader
    mnupop.Visible = False
    cmbCust.ListHeight = 2500
    cmbCust.AddItem "<All Customers>", "0"
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode")
    
    'Add Headers...
    '------------------------------------------------------------------------
    For i = 5 To LV.ColumnHeaders.Count
        LV.ColumnHeaders.Remove 5
    Next i
    
    With rs
        .Open "Select ProcessID,Description From Processes Order By SNo Desc", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            Set CHeader = LV.ColumnHeaders.Add(, ![ProcessID] & "'", ![Description] & "")
            CHeader.Width = 800
            CHeader.Alignment = lvwColumnCenter
            .MoveNext
        Loop
        Set CHeader = LV.ColumnHeaders.Add(, , "Not Issd")
        CHeader.Width = 800
        CHeader.Alignment = lvwColumnCenter
        
        Set CHeader = LV.ColumnHeaders.Add(, , "Wastage")
        CHeader.Width = 800
        CHeader.Alignment = lvwColumnCenter
        .Close
    End With
    Set rs = Nothing
    '-----------------------------------------------------------------------
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Dim TheHeight As Long
    Frame1.Left = (Me.ScaleWidth - Frame1.Width) / 2
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    
End Sub


Private Sub RefreshLV(Optional ItemCode As String)
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim Cond As String
    Dim DTRange As String
    Dim itm As ListItem
    
    Screen.MousePointer = vbHourglass
    
      
    With rs
        
        .Open "Select * From VFOrderStatus Where OrderNo='" & cmbOrders.Text & "' and ItemID='" & cmbItems.Text & "'  Order By SNo", con, adOpenForwardOnly, adLockReadOnly
        
        If .RecordCount > 0 Then
            Set itm = LV.ListItems.Add(, ![ITemID], .AbsolutePosition)
            itm.ListSubItems.Add , , ![ItemDescription] & ""
            itm.ListSubItems.Add , , "<All Processes>"
            itm.ListSubItems.Add(, , ![TotalInStock] & "").ForeColor = RGB(100, 100, 100)
            itm.ListSubItems.Add(, , ![TotalOutStock] & "").ForeColor = vbRed
            itm.ListSubItems.Add(, , Val(itm.SubItems(3) & "") + Val(itm.SubItems(4) & "")).ForeColor = vbBlue
        End If
        'LV.Sorted = False
        Do Until .EOF
        
            Set itm = LV.ListItems.Add(, ![ProcessID] & "P", .AbsolutePosition + 1)
            
            itm.ListSubItems.Add , , ![ItemDescription] & ""
            'ITM.SubItems(1) = ![CustCode] & ""
            itm.ListSubItems.Add , , ![Description] & ""
            itm.ListSubItems.Add(, , ![InStock] & "").ForeColor = RGB(100, 100, 100)
            itm.ListSubItems.Add(, , ![OutStock] & "").ForeColor = vbRed
            itm.ListSubItems.Add(, , Val(itm.SubItems(3) & "") + Val(itm.SubItems(4) & "")).ForeColor = vbBlue
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Exit Sub
    If LV.SortKey + 1 = ColumnHeader.Index Then
        If LV.SortOrder = lvwAscending Then
            LV.SortOrder = lvwDescending
            ColumnHeader.Icon = "Down"
        Else
            LV.SortOrder = lvwAscending
            ColumnHeader.Icon = "Up"
        End If
    Else
        LV.ColumnHeaderIcons = Nothing
        Set LV.ColumnHeaderIcons = ImageList2
        LV.SortKey = ColumnHeader.Index - 1
        LV.SortOrder = lvwAscending
        ColumnHeader.Icon = "Down"
    End If
    'Call RefreshSnos
    'If LV.Sorted = False Then LV.Sorted = True
End Sub

Private Sub LV_DblClick()
    'If LV.ListItems.Count = 0 Or OutOfFinancialYear Then Exit Sub
    'If InStr(LV.SelectedItem, "-") > 0 Then Call mnuEditDesc_Click
    Exit Sub
    If LV.ListItems.Count = 0 Then Exit Sub
    Load frmFNewCustOrder
    Call frmFNewCustOrder.EditOrder(Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1))
    
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)

    On Error GoTo err
    Exit Sub
    Dim rs As New ADODB.Recordset
    Dim Cond As String
    
    Dim itm As ListItem
    
    
    Dim Valid As Boolean
      
    With rs
        
        .Open "Select * From VFOrderStatus1 Where OrderNo='" & cmbOrders.Text & "' and CompItemCode='" & Item.Tag & "'  Order By SNo", con, adOpenForwardOnly, adLockReadOnly
        LV1.ListItems.Clear
        Do Until .EOF
            
            Set itm = LV1.ListItems.Add(, , LV1.ListItems.Count + 1)
            itm.Tag = ![ProcessID] & ""
            itm.ListSubItems.Add(, , ![Process] & "").Bold = True
            
            itm.ListSubItems.Add(, , Val(![NetQty] & "")).ForeColor = RGB(0, 100, 255)
            itm.ListSubItems.Add(, , Val(![TotIssable] & "")).ForeColor = RGB(255, 100, 100)
            itm.ListSubItems.Add(, , ![TotIssdQty] & "").ForeColor = RGB(100, 100, 200)
            itm.ListSubItems.Add(, , ![InStock] & "").ForeColor = RGB(50, 50, 50)
            itm.ListSubItems.Add(, , ![OutStock] & "").ForeColor = RGB(150, 50, 100)
            
            
            .MoveNext
        Loop
        .Close
    End With
    
    
    Set rs = Nothing
    
    Exit Sub
err:
    'Screen.MousePointer = vbDefault
    MsgBox err.Description
        
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    Exit Sub
    If LV.ListItems.Count = 0 Or Button <> 2 Then Exit Sub
    Me.PopUpMenu mnupop
    'If LV.ListItems.Count = 0 Or Button <> 2 Or OutOfFinancialYear Then Exit Sub

    'if a Voucher Entry Show Delete Menu
    'If InStr(LV.SelectedItem, "-") > 0 Then Me.PopupMenu mnuPop
    
End Sub





Private Sub LV1_DblClick()
    On Error GoTo err
    If LV1.ListItems.Count = 0 Then Exit Sub
    With LV1.SelectedItem
        If Val(.SubItems(4)) = 0 Then
            MsgBox "No Issue Detail To Show.", vbInformation
            Exit Sub
        End If
'        Dim F As New frmVendIssdDetail
'        Call F.ShowDetail(cmbOrders.Text, LV.SelectedItem.Tag, .Tag)
    End With
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub txtDesc_GotFocus()
    cmdRefresh.Default = False
    txtDesc.SelStart = 0
    txtDesc.SelLength = Len(txtDesc)
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
  
    If KeyAscii = 13 Then
    
        If Trim(txtDesc = "") Then
            MsgBox "Invalid Description", vbInformation
            Exit Sub
        End If
    
        con.Execute "Update Vouchers Set [Desc]='" & txtDesc & "' Where Sno=" & Val(LV.SelectedItem.key)
        LV.SelectedItem.ListSubItems(3) = txtDesc & txtDesc.Tag
        txtDesc.Visible = False
        cmdRefresh.Default = False
    ElseIf KeyAscii = 27 Then
        txtDesc.Visible = False
        cmdRefresh.Default = False
    End If

End Sub

Private Sub txtDesc_LostFocus()
    txtDesc.Visible = False
End Sub
