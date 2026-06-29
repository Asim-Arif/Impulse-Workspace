VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmFCustOrderListOLD 
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   9885
   ScaleMode       =   0  'User
   ScaleWidth      =   10641.08
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   10560
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Order List"
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
         TabIndex        =   9
         Top             =   165
         Width           =   10605
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Order List"
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
         TabIndex        =   10
         Top             =   180
         Width           =   10590
      End
   End
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   9225
      Left            =   15
      TabIndex        =   3
      Top             =   615
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
               Picture         =   "frmFCustOrderListOLD.frx":0000
               Key             =   "Down"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFCustOrderListOLD.frx":01BC
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
         TabIndex        =   11
         Top             =   1605
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ComboBox cmbRange 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmFCustOrderListOLD.frx":0378
         Left            =   4440
         List            =   "frmFCustOrderListOLD.frx":038B
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   345
         Width           =   2010
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   8190
         TabIndex        =   1
         Top             =   345
         Width           =   1560
         _ExtentX        =   2752
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
         Format          =   22740995
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   6465
         TabIndex        =   0
         Top             =   345
         Width           =   1710
         _ExtentX        =   3016
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
         Format          =   22740995
         CurrentDate     =   37055
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   195
         TabIndex        =   6
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
         Height          =   7965
         Left            =   45
         TabIndex        =   12
         Top             =   765
         Width           =   11715
         _ExtentX        =   20664
         _ExtentY        =   14049
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Country"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Order #"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Order DT"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Order Amt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Status"
            Object.Width           =   2646
         EndProperty
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   375
         Left            =   10425
         TabIndex        =   13
         Top             =   8775
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "2355;661"
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
         Left            =   2550
         TabIndex        =   7
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
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   360
         Left            =   9765
         TabIndex        =   2
         Top             =   300
         Width           =   1500
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "2646;635"
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
         BackColor       =   &H00AFBDCB&
         Caption         =   $"frmFCustOrderListOLD.frx":03D3
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
         Left            =   165
         TabIndex        =   4
         Top             =   120
         Width           =   11130
      End
   End
End
Attribute VB_Name = "frmFCustOrderListOLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub mnuProdPlan_Click()
    Load frmProdPlan
    Call frmProdPlan.AddPlanning(Left(LV.SelectedItem.Key, Len(LV.SelectedItem.Key) - 1))
End Sub
Private Sub cmbCust_matched()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    SQL = "Select Distinct Country From ForeignCustomers "
    If cmbCust.ID <> "0" Then
        SQL = SQL & " Where CustCode='" & cmbCust.ID & "'"
    End If
    
    With rs
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
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

Private Sub cmbRange_Click()

    DtTo = ServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    Select Case cmbRange.ListIndex
        Case 0
            DtFrom = ServerDate
        Case 1
            DtFrom = DateAdd("d", -7, ServerDate)
        Case 2
            DtFrom = DateAdd("d", -15, ServerDate)
        Case 3
            DtFrom = DateAdd("d", -30, ServerDate)
        Case 4
            DtFrom.Enabled = True
            DtTo.Enabled = True
            DtFrom.SetFocus
            Exit Sub
    End Select
    
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub



Private Sub DtFrom_Change()
    'Call setDateRange(DtFrom)
End Sub

Private Sub DtTo_Change()
    'Call setDateRange(DtTo)
End Sub


Private Sub Form_Load()
    
    'mnuPop.Visible = False
    cmbCust.ListHeight = 2500
    cmbCust.AddItem "<All Customers>", "0"
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where custCode<>'Hico'")
    
    cmbRange.ListIndex = 0
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Dim TheHeight As Long
    Frame1.Left = (Me.ScaleWidth - Frame1.Width) / 2
    
    
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    
End Sub


Private Sub RefreshLV()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim Cond As String
    Dim DTRange As String
    Dim ITM As ListItem
    
    Screen.MousePointer = vbHourglass
    
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
    
    If cmbRange.ListIndex = 0 Then
        DTRange = " DT Between '" & DateAdd("d", -15, ServerDate) & "' and '" & ServerDate & "'"
    ElseIf cmbRange.ListIndex = 1 Then
        DTRange = " DT Between '" & DateAdd("d", -30, ServerDate) & "' and '" & ServerDate & "'"
    ElseIf cmbRange.ListIndex = 2 Then
        DTRange = " DT Between '" & DateAdd("d", -60, ServerDate) & "' and '" & ServerDate & "'"
    ElseIf cmbRange.ListIndex = 3 Then
        DTRange = " DT Between '" & DateAdd("d", -90, ServerDate) & "' and '" & ServerDate & "'"
    ElseIf cmbRange.ListIndex = 4 Then
        DTRange = " DT Between '" & DtFrom & "' and '" & DtTo & "'"
    End If
    
    If Cond = "" Then
        Cond = " Where "
    Else
        Cond = Cond & " and "
    End If
    Cond = Cond & DTRange
    
    With rs
        
        .Open "Select * From VFOrderList " & Cond & " Order By DT", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        'LV.Sorted = False
        Do Until .EOF
            Set ITM = LV.ListItems.Add(, ![OrderNo] & "'", .AbsolutePosition)
            ITM.Tag = ![Curr] & ""
            ITM.ListSubItems.Add , , ![CustCode] & ""
            'ITM.SubItems(1) = ![CustCode] & ""
            ITM.ListSubItems.Add , , ![Country] & ""
            'ITM.SubItems(2) = ![Country] & ""
            ITM.ListSubItems.Add , , ![OrderNo] & ""
            'ITM.SubItems(3) = ![OrderNo] & ""
            ITM.ListSubItems.Add , , Format(![DT] & "", "dd-MMM-yyyy")
            'ITM.SubItems(4) = Format(![Dt] & "", "dd-MMM-yyyy")
            ITM.ListSubItems.Add(, , ![Curr] & " " & Val(![OrderAmt] & "")).Tag = Val(![OrderAmt] & "")
            'ITM.SubItems(5) = Val(![OrderAmt] & "")
            ITM.SubItems(6) = ""
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
    Call RefreshSnos
    'If LV.Sorted = False Then LV.Sorted = True
End Sub

Private Sub LV_DblClick()
    'If LV.ListItems.Count = 0 Or OutOfFinancialYear Then Exit Sub
    'If InStr(LV.SelectedItem, "-") > 0 Then Call mnuEditDesc_Click
    On Error GoTo err
    Dim OrderNo As String
    Dim iTemp As Integer
    
    If LV.ListItems.Count = 0 Then Exit Sub
    
    
    OrderNo = Left(LV.SelectedItem.Key, Len(LV.SelectedItem.Key) - 1)
    
    iTemp = Val(con.Execute("Select Count(OrderNo) From VFItemsQtyToPlan Where RemQty>0 and OrderNo='" & OrderNo & "'").Fields(0).value & "")
    
    If iTemp = 0 Then
        MsgBox "Can't Edit This Order.It Has Been Planned.", vbInformation
        Exit Sub
    End If
    
    If LV.SelectedItem.SubItems(1) = "HICO" Then
        Load frmFNewStockOrder
        Call frmFNewStockOrder.EditOrder(Left(LV.SelectedItem.Key, Len(LV.SelectedItem.Key) - 1))
    Else
        Load frmFNewCustOrder
        Call frmFNewCustOrder.EditOrder(Left(LV.SelectedItem.Key, Len(LV.SelectedItem.Key) - 1))
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If LV.ListItems.Count = 0 Or Button <> 2 Then Exit Sub
 
    Dim strRet As String
    If UserHasAccess("AddProdPlan") Then
        strRet = ShowPopUpMenu1("&Edit This Order", "-", "&Add Production Planning", "-", "History Of The Order (Summary)", "History Of The Order (Detail)", "-", "Print Order (&Summary)", "Print Order (&Detail)", "-", "&Print Order")
    Else
        strRet = ShowPopUpMenu("&Edit This Order", "-", "History Of The Order (Summary)", "History Of The Order (Detail)", "-", "Print Order (&Summary)", "Print Order (&Detail)", "-", "&Print Order")
    End If
    
    Select Case strRet
        Case "&Edit This Order"
            Call EditOrder
        Case "&Add Production Planning"
            Call AddProdPlanning
        Case "History Of The Order (Summary)"
            
        Case "History Of The Order (Detail)"
            
        Case "Print Order (&Summary)"
            Call PrintSummary
        Case "Print Order (&Detail)"
            
        Case "&Print Order"
            Call PrintOrder
    End Select
    'Me.PopupMenu mnuPop
    'If LV.ListItems.Count = 0 Or Button <> 2 Or OutOfFinancialYear Then Exit Sub

    'if a Voucher Entry Show Delete Menu
    'If InStr(LV.SelectedItem, "-") > 0 Then Me.PopupMenu mnuPop
    
End Sub
Private Sub AddProdPlanning()
    Load frmProdPlan
    If Not frmProdPlan.CheckForPlanning(Left(LV.SelectedItem.Key, Len(LV.SelectedItem.Key) - 1)) Then
        Unload frmProdPlan
        MsgBox "This Order Has Been Planned.", vbInformation
        Exit Sub
    End If
End Sub
Private Sub EditOrder()
    Call LV_DblClick
End Sub
Private Sub mnuEdit_Click()
    Call LV_DblClick
End Sub

Private Sub PrintOrder()
    On Error GoTo err
    
    Screen.MousePointer = vbHourglass

    Load frmPrevRpt
    frmPrevRpt.ShowReport "{CustOrders.OrderNo}='" & Left(LV.SelectedItem.Key & "'", Len(LV.SelectedItem.Key) - 1) & "' ", rptFCustOrders
    
    Screen.MousePointer = vbDefault

    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub PrintSummary()
    On Error GoTo err
    Me.MousePointer = vbHourglass
    
    Call StartTrans(con)
    con.Execute "Delete From PrintFOrderList"
    
    For i = 1 To LV.ListItems.Count
        con.Execute "Insert Into PrintFOrderList(Sno,Customer,Country," & _
         "OrderNo,OrderDT,OrderAmt,Status,Curr) Values(" & LV.ListItems(i) & _
         ",'" & LV.ListItems(i).SubItems(1) & "','" & LV.ListItems(i).SubItems(2) & _
         "','" & LV.ListItems(i).SubItems(3) & "','" & _
         LV.ListItems(i).SubItems(4) & "'," & Val(LV.ListItems(i).ListSubItems(5).Tag & "") & _
         ",'" & LV.ListItems(i).SubItems(6) & "','" & LV.ListItems(i).Tag & "')"
    Next i
    
    con.CommitTrans
    
    Load frmPrevRpt
    
    Dim strFromTo As String
    Dim iNumber As Integer
    If cmbRange.ListIndex = 0 Then
        iNumber = -15
    ElseIf cmbRange.ListIndex = 1 Then
        iNumber = -30
    ElseIf cmbRange.ListIndex = 2 Then
        iNumber = -60
    ElseIf cmbRange.ListIndex = 3 Then
        iNumber = -90
    End If
    
    If cmbRange.ListIndex = 4 Then
        strFromTo = Format(DtFrom, "dd-MMM-yyyy") & " To  " & Format(DtTo, "dd-MMM-yyyy")
    Else
        strFromTo = Format(DateAdd("d", iNumber, ServerDate), "dd-MMM-yyyy") & " To  " & Format(ServerDate, "dd-MMM-yyyy")
    End If
    
        
    rptFOrderList.FormulaFields(3).Text = "'" & cmbCust.Text & "'"
    rptFOrderList.FormulaFields(4).Text = "'" & strFromTo & "'"
    
    frmPrevRpt.ShowReport "", rptFOrderList
    
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Private Sub mnuSummary_Click()

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
    
   con.Execute "Update Vouchers Set [Desc]='" & txtDesc & "' Where Sno=" & Val(LV.SelectedItem.Key)
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
Private Sub RefreshSnos()
    For i = 1 To LV.ListItems.Count
        LV.ListItems(i).Text = i
    Next i
End Sub
