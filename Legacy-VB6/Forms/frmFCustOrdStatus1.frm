VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmFCustOrdStatus1 
   ClientHeight    =   9600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   9600
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
      Height          =   8925
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
               Picture         =   "frmFCustOrdStatus1.frx":0000
               Key             =   "Down"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFCustOrdStatus1.frx":01BC
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
         Height          =   4245
         Left            =   195
         TabIndex        =   10
         Top             =   660
         Width           =   11565
         _ExtentX        =   20399
         _ExtentY        =   7488
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   1014
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cat ID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item ID"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Description"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Unit"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Price"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Amount"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView LV1 
         Height          =   3465
         Left            =   195
         TabIndex        =   11
         Top             =   4950
         Width           =   11565
         _ExtentX        =   20399
         _ExtentY        =   6112
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         PictureAlignment=   5
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Manufacturing Process"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Order Qty"
            Object.Width           =   1614
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Issueable"
            Object.Width           =   1614
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Issued"
            Object.Width           =   1614
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Repair"
            Object.Width           =   1614
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Repair Iss"
            Object.Width           =   1614
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Completed"
            Object.Width           =   1693
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Wastage"
            Object.Width           =   1614
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "In Process"
            Object.Width           =   1640
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "Not Issd"
            Object.Width           =   1614
         EndProperty
         Picture         =   "frmFCustOrdStatus1.frx":0378
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   375
         Left            =   8385
         TabIndex        =   12
         Top             =   8475
         Width           =   1650
         ForeColor       =   0
         Caption         =   "Print     "
         PicturePosition =   327683
         Size            =   "2910;661"
         Picture         =   "frmFCustOrdStatus1.frx":07D2
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
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
         Top             =   8475
         Width           =   1650
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "2910;661"
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
Attribute VB_Name = "frmFCustOrdStatus1"
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
    
    'Now Get Rid Of HICO Order
    If Cond = "" Then
        Cond = " Where OrderNo<>'HICO Order'"
    Else
        Cond = Cond & " and OrderNo<>'HICO Order'"
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
    Call EditOrder(cmbOrders.Text)
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
    If cmbCountry.ListCount = 2 Then
        cmbCountry.ListIndex = 1
    Else
        cmbCountry.ListIndex = 0
    End If
    
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdClose_Click()
    Me.Hide
    Unload Me
End Sub


Private Sub cmdPrint_Click()
    
    Dim strRet As String
    strRet = ShowPopUpMenu1("&Order Status Report", "-", "Order Status Report With &Detail", "-", "Order Status Report With &Full Detail")
    
    If strRet = "&Order Status Report" Then
        Call PrintReport(0)
    ElseIf strRet = "Order Status Report With &Detail" Then
        Call PrintReport(1)
    ElseIf strRet = "Order Status Report With &Full Detail" Then
        Call PrintReport(2)
    End If
    
End Sub
Private Sub PrintReport(Param As Integer)

    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    
    Dim rpt As craxddrt.Report 'New rptProdOrderStatus
    
    Set rpt = rptApp.OpenReport(RptPath & "\rptProdOrderStatus.rpt") '
    Dim F As New frmPrevRpt
    Dim SelFor As String
    
    If cmbOrders.MatchFound Then
        SelFor = "{VFOrderStatus1.OrderNo}='" & cmbOrders.Text & "' "
    Else
        SelFor = " "
    End If
    
    If Param = 0 Then
        rpt.Section10.Suppress = True
    ElseIf Param = 1 Then
        
        rpt.Subreport1_Section4.Suppress = True
        rpt.Subreport1_Section5.Suppress = True
        rpt.Subreport1_Section6.Suppress = True

    End If
    
    F.ShowReport SelFor, rpt
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub
Private Sub Form_Load()
        
    mnupop.Visible = False
    cmbCust.ListHeight = 2500
    cmbCust.AddItem "<All Customers>", "0"
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where CustCode<>'HICO'")
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Dim TheHeight As Long
    Frame1.Left = (Me.ScaleWidth - Frame1.Width) / 2
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    
End Sub





Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim Cond As String
    
    Dim itm As ListItem
    
    
    Dim Valid As Boolean
      
    With rs
    
        .Open "Select * From VFOrderStatus1  Where OrderNo='" & cmbOrders.Text & "' and CompItemCode='" & Item.Tag & "'  Order By SNo", con, adOpenForwardOnly, adLockReadOnly
        '.Open "Select * From VFOrderStatus1 ", con, adOpenForwardOnly, adLockReadOnly
        LV1.ListItems.Clear
        
        Dim lCompQty As Long
        
        Do Until .EOF
            
            Set itm = LV1.ListItems.Add(, , LV1.ListItems.Count + 1)
            itm.Tag = ![ProcessID] & ""
            itm.ListSubItems.Add(, , ![Process] & "").Bold = True   'Process Name
            
            itm.ListSubItems.Add(, , Val(![NetQty] & "")).ForeColor = RGB(0, 100, 255)  'Order Qty + Free Stock
            
            'itm.ListSubItems.Add(, , Val(![NetQty] & "") - Val(![NetOrderQty] & "")).ForeColor = RGB(0, 100, 255)
            
            If LV1.ListItems.Count = 1 Then
                itm.ListSubItems.Add(, , Val(![NotIssued] & "")).ForeColor = RGB(255, 100, 100)     'Not Issued
            Else
                itm.ListSubItems.Add(, , Val(![TotIssable] & "")).ForeColor = RGB(255, 100, 100)    'Not Issued
            End If
            
            If Val(![TotTransferred] & "") <> 0 Then itm.ListSubItems(3).ToolTipText = "Transfer Qty :" & Val(![TotTransferred] & "")
            
            itm.ListSubItems.Add(, , Val(![TotIssdQty] & "")).ForeColor = RGB(100, 100, 200)    'Total
            
            itm.ListSubItems.Add(, , Val(![RcvRepairQty] & "")).ForeColor = RGB(100, 100, 200)
            itm.ListSubItems.Add(, , Val(![RepairIssued] & "")).ForeColor = RGB(100, 100, 200)  '
            
            'lCompQty = Val(![InStock] & "") - Val(![RepairIssued] & "") - Val(![RepairWastage] & "")
            lCompQty = Val(![TotIssdQty] & "") - Val(![OutStock] & "") - Val(![RepairIssued] & "") - Val(![RepairWastage] & "") - Val(![TotWastage] & "") - Val(![RepairWastage] & "") - Val(![RcvRepairQty] & "")  'Completed
            
            itm.ListSubItems.Add(, , lCompQty).ForeColor = RGB(50, 50, 50)
            
            itm.ListSubItems.Add(, , Val(![TotWastage] & "") + Val(![RepairWastage] & "")).ForeColor = RGB(200, 50, 150)
            
            itm.ListSubItems.Add(, , ![OutStock] & "").ForeColor = RGB(255, 100, 0)
            
            itm.ListSubItems.Add(, , Val(![NotIssued] & "")).ForeColor = RGB(255, 100, 255)
            
            
            .MoveNext
        Loop
        .Close
    End With
    
    
    Set rs = Nothing
    
    Exit Sub
err:
    'Screen.MousePointer = vbDefault
    MsgBox err.Description & vbNewLine & "Error # :" & err.Number & vbNewLine & "Source :" & err.Source
        
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
Public Sub EditOrder(TempOrdNo As String)
    
    On Error GoTo err
    strOrderNo = TempOrdNo
    Dim rs As New ADODB.Recordset
    Dim itm As ListItem
    LV.ListItems.Clear
    With rs
        .Open "Select * From VrptOrders Where OrderNo='" & strOrderNo & "'", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            Set itm = LV.ListItems.Add(, ![ITemID] & "'", .AbsolutePosition)
            itm.Tag = ![CompItemID] & "" 'cmbItems.List(cmbItems.ListIndex, 2)
            
            itm.ListSubItems.Add , , Left(![ITemID] & "", 2) '![CatID] & ""
            itm.ListSubItems.Add , , ![ITemID] & ""
            itm.ListSubItems.Add(, ![ItemName] & "", ![ItemName] & "").Tag = ![Description] & ""
            
            itm.SubItems(4) = txtUnit
            
            itm.SubItems(5) = Format(![Price] & "", "#0.00")
            itm.SubItems(6) = Val(![Qty] & "")
            itm.SubItems(7) = Val(![Qty] & "") * Val(![Price])
   
            .MoveNext
        Loop
        .Close
    End With
    'Call GetTotal

    txtQty = 0
    'cmbItems.SetFocus
    
    Set rs = Nothing
    'Me.Show
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

