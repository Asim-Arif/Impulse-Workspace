VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmIPOList 
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   360
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
      Left            =   75
      TabIndex        =   4
      Top             =   15
      Width           =   10560
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Internal Production Order"
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
         Caption         =   "Internal Production Order"
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
      Height          =   9225
      Left            =   15
      TabIndex        =   0
      Top             =   615
      Width           =   11850
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
         ItemData        =   "frmIPOList.frx":0000
         Left            =   5715
         List            =   "frmIPOList.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   360
         Width           =   2025
      End
      Begin VB.TextBox txtTotalAmt 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8595
         TabIndex        =   13
         Top             =   8805
         Width           =   1785
      End
      Begin ComboList.Usercontrol1 cmbCompany 
         Height          =   285
         Left            =   165
         TabIndex        =   11
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
               Picture         =   "frmIPOList.frx":0078
               Key             =   "Down"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmIPOList.frx":0234
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
         Left            =   2505
         TabIndex        =   2
         Top             =   345
         Width           =   1815
         _ExtentX        =   3201
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
         TabIndex        =   8
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Country"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "IPO #"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Target Date"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   9255
         TabIndex        =   14
         Top             =   360
         Width           =   1485
         _ExtentX        =   2619
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
         Format          =   122159107
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   7755
         TabIndex        =   15
         Top             =   360
         Width           =   1500
         _ExtentX        =   2646
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
         Format          =   122159107
         CurrentDate     =   37055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Amount (Rs.):"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6390
         TabIndex        =   12
         Top             =   8835
         Width           =   2130
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   360
         Left            =   10815
         TabIndex        =   10
         Top             =   270
         Width           =   945
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "1667;635"
         Accelerator     =   82
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   375
         Left            =   10425
         TabIndex        =   9
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
         Left            =   4320
         TabIndex        =   3
         Top             =   345
         Width           =   1365
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2408;503"
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
         Caption         =   $"frmIPOList.frx":03F0
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
         TabIndex        =   1
         Top             =   120
         Width           =   10560
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "Pop up"
      Begin VB.Menu mnuNew 
         Caption         =   "&New IPO"
      End
      Begin VB.Menu mnuPrintDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit IPO"
      End
      Begin VB.Menu mnuEditDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteIPO 
         Caption         =   "Delete IPO"
      End
      Begin VB.Menu mnuDashDeleteIPO 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintIPO 
         Caption         =   "Print IPO"
      End
      Begin VB.Menu mnuDashLots 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintSpecs 
         Caption         =   "Print Specification Sheets"
      End
      Begin VB.Menu mnuPrintContractReview 
         Caption         =   "Print Contract Review"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDashGL 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGenerateLots 
         Caption         =   "Generate Lots"
      End
   End
End
Attribute VB_Name = "frmIPOList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQLString As String
Dim bForProduction As Boolean

Private Sub mnuProdPlan_Click()
    Load frmProdPlan
    Call frmProdPlan.AddPlanning(Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1))
End Sub

Public Sub ShowMe(p_bForProduction As Boolean)
    bForProduction = p_bForProduction
    If bForProduction Then
        Label2.Visible = False
        txtTotalAmt.Visible = False
    End If
    Me.Show
    Call RefreshLV
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
            .MoveNEXT
        Loop
        .Close
    End With
    cmbCountry.ListIndex = 0
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
            DtFrom = DateAdd("d", -15, ServerDate)
        Case 2
            DtFrom = DateAdd("d", -30, ServerDate)
        Case 3
            DtFrom = DateAdd("d", -60, ServerDate)
        Case 4
            DtFrom = DateAdd("d", -90, ServerDate)
        Case 5
            DtFrom = DateAdd("d", -180, ServerDate)
        Case 6
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

Private Sub Form_Load()
    
    mnupop.Visible = False
    bForProduction = False
    cmbCompany.ListHeight = 4000
    cmbCompany.AddItem "< All Companies>", "0"
    cmbCompany.AddVals con, "CompanyName", "Companies", "EntryID", " Order By CompanyName"
    cmbCompany.ID = "0"
    
    cmbCust.ListHeight = 2500
    cmbCust.AddItem "<All Customers>", "0"
    'Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where custCode<>'Hico'")
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode")
    
    cmbCust.ID = "0"
    cmbCountry.ListIndex = 0
    Call RefreshLV
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
    
    Cond = " WHERE IPODT BETWEEN '" & DtFrom.Value & "' AND '" & DtTo.Value & "'"
    If cmbCust.ID <> "0" Then
        Cond = Cond & " AND CustCode='" & cmbCust.ID & "' "
    End If
    
    If cmbCountry.ListIndex <> 0 Then
        Cond = Cond & " and Country='" & cmbCountry & "'"
    End If
    
    SQLString = "SELECT * FROM VInternalProductionOrder " & Cond & " Order By IPODT DESC"
    Dim sStatus As String
    Dim iRecNo As Integer
    Dim dTotalAmt As Double, dTotalAmtPKR As Double
    dTotalAmt = 0
    With rs
        
        .Open SQLString, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        iRecNo = .RecordCount
        'LV.Sorted = False
        Do Until .EOF
            
            Set ITM = LV.ListItems.add(, ![IPONo] & "'", iRecNo - .AbsolutePosition + 1)
            ITM.Tag = Val(!EntryID & "")
            ITM.ListSubItems.add , , ![CustCode] & ""
            'ITM.SubItems(1) = ![CustCode] & ""
            ITM.ListSubItems.add , , ![Country] & ""
            'ITM.SubItems(2) = ![Country] & ""
            ITM.ListSubItems.add , , ![IPONo] & ""
            'ITM.SubItems(3) = ![OrderNo] & ""
            ITM.ListSubItems.add , , Format(![IPODT] & "", "dd-MMM-yyyy")
            'ITM.SubItems(4) = Format(![Dt] & "", "dd-MMM-yyyy")
            ITM.ListSubItems.add , , Format(![TargetDT] & "", "dd-MMM-yyyy")
            
            'dTotalAmt = dTotalAmt + Val(!OrderAmt & "")
            'dTotalAmtPKR = dTotalAmtPKR + (Val(!OrderAmt & "") * Val(!ExchRate & ""))
            .MoveNEXT
            
        Loop
        
        .Close
        
    End With
    
    Set rs = Nothing
    'txtTotalAmt = Format(dTotalAmtPKR, "###,##0.00")
    
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
    
    If LV.ListItems.count = 0 Then Exit Sub
    
    If LotsAlreadyGenerated() Then
        MsgBox "Lots Generated, Can't Edit Now.", vbInformation
        Exit Sub
    End If
    
    Dim f As New frmNewIPO
    Load f
    f.EditIPO Val(LV.SelectedItem.SubItems(3))
    
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 2 Then Exit Sub
    Me.PopUpMenu mnupop
    
End Sub

Private Sub PrintPTC()

    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\PTC.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{FCustomerOrders.OrderNo}='" & LV.SelectedItem.SubItems(4) & "'", rpt
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub PrintOrderBalanceReportOld()

    On Error GoTo err
    Me.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report, f As New frmPrevRpt
    Set rpt = rptApp.OpenReport(RptPath & "\rptPartiallyShippedOrders.rpt")
    'f.ShowReport "{FCustomerOrders.OrderNo}='" & LV.SelectedItem.SubItems(4) & "' AND {VFOrderItemswithShippedQty.Qty}>{VFOrderItemswithShippedQty.ShippedQty}", rpt
    f.ShowReport "{FCustomerOrders.OrderNo}='" & LV.SelectedItem.SubItems(4) & "'", rpt
    Me.MousePointer = vbNormal
    
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub PrintOrderStatus()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report 'New rptCustOrderStatus
    Set rpt = rptApp.OpenReport(RptPath & "\rptCustOrderStatus.rpt") '
    Dim f As New frmPrevRpt
    f.ShowReport "{VFCustOrders.OrderNo}='" & LV.SelectedItem.SubItems(3) & "' ", rpt
        
    Set rpt = Nothing
    Set f = Nothing
    
    Exit Sub
    
err:
    MsgBox err.Description
    
End Sub
Private Sub AddProdPlanning()
    
    Load frmProdPlan
    If Not frmProdPlan.CheckForPlanning(Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1)) Then
        Unload frmProdPlan
        MsgBox "This Order Has Been Planned.", vbInformation
        Exit Sub
    End If
    
End Sub

Private Sub EditOrder()
    Call LV_DblClick
End Sub

Private Sub mnuDeleteIPO_Click()
    
    If LV.ListItems.count = 0 Then Exit Sub
    
    If LotsAlreadyGenerated() Then
        MsgBox "Lots Generated, Can't Delete Now.", vbInformation
        Exit Sub
    End If
    
    If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbYes Then
    
        Dim lEntryID As Long
        lEntryID = Val(LV.SelectedItem.Tag)
        
        con.Execute "DELETE FROM InternalProductionOrder WHERE EntryID=" & lEntryID
        con.Execute "DELETE FROM InternalProductionOrderDetail WHERE RefID=" & lEntryID
    End If
    
End Sub

Private Sub mnuedit_Click()
    Call LV_DblClick
End Sub

Private Sub PrintOrder(Optional bPrintBalance As Boolean = False)

    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    Dim strSelection As String
    Dim rpt As CRAXDDRT.Report
    If Not bPrintBalance Then
        Set rpt = rptApp.OpenReport(RptPath & "\rptProductionOrder.rpt") '
        
        rpt.FormulaFields.GetItemByName("ProductionOrder").Text = bForProduction
        strSelection = "{CustOrders.OrderNo}='" & Left(LV.SelectedItem.key & "'", Len(LV.SelectedItem.key) - 1) & "' "
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\rptProductionOrderBalanceQty.rpt") '
        rpt.FormulaFields.GetItemByName("ProductionOrder").Text = bForProduction
        strSelection = "{CustOrders.OrderNo}='" & Left(LV.SelectedItem.key & "'", Len(LV.SelectedItem.key) - 1) & "' AND {@Balance}>0"
    End If
    rpt.FormulaFields.GetItemByName("UserName").Text = "'" & CurrentUserName & "'"
    Load frmPrevRpt
    frmPrevRpt.ShowReport strSelection, rpt
    
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
    
    For i = 1 To LV.ListItems.count
        con.Execute "Insert Into PrintFOrderList(Sno,Customer,Country," & _
         "OrderNo,OrderDT,OrderAmt,Status,Curr) Values(" & LV.ListItems(i) & _
         ",'" & LV.ListItems(i).SubItems(1) & "','" & LV.ListItems(i).SubItems(2) & _
         "','" & LV.ListItems(i).SubItems(3) & "','" & _
         LV.ListItems(i).SubItems(4) & "'," & Val(LV.ListItems(i).ListSubItems(5).Tag & "") & _
         ",'" & LV.ListItems(i).SubItems(6) & "','" & LV.ListItems(i).Tag & "')"
    Next i
    
    con.CommitTrans
    
    Dim f As New frmPrevRpt
    Dim rpt As CRAXDDRT.Report 'New rptFOrderList
    Set rpt = rptApp.OpenReport(RptPath & "\rptFOrderList.rpt") '
    rpt.SQLQueryString = SQLString
        
    rpt.FormulaFields(3).Text = "'" & cmbCust.Text & " " & cmbCountry.Text & "'"
    'rptFOrderList.FormulaFields(4).Text = "'" & strFromTo & "'"
    
    frmPrevRpt.ShowReport " ", rpt
    
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Private Sub mnuSummary_Click()

End Sub

Private Sub mnuGenerateLots_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    If LotsAlreadyGenerated() Then
        MsgBox "Lots already Generated.", vbInformation
        Exit Sub
    End If
    Dim f As New frmGenerateIPOLots
    Load f
    f.ShowMe (Val(LV.SelectedItem.SubItems(3)))
'    Dim rs As New ADODB.Recordset
'    With rs
'        .Open "SELECT T1.EntryID,T1.Qty,MaxLotSize FROM " & _
'         "InternalProductionOrderDetail T1 INNER JOIN InternalProductionOrder ON InternalProductionOrder.EntryID=T1.RefID " & _
'         "INNER JOIN Items ON T1.CompItemCode=Items.ItemID INNER JOIN ItemCatagories ON Items.CatID=ItemCatagories.CatID WHERE IPONo=" & _
'         Val(LV.SelectedItem.SubItems(3)) & " ORDER BY T1.EntryID", con, adOpenForwardOnly, adLockReadOnly
'        Call StartTrans(con)
'        Do Until .EOF
'            Call SaveLots(Val(!EntryID & ""), Val(!MaxLotSize & ""), Val(!Qty & ""))
'            .MoveNEXT
'        Loop
'        con.CommitTrans
'        .Close
'    End With
'
'    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Function LotsAlreadyGenerated() As Boolean

    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(IPONo)", "VInternalProductionOrderDetail", "WHERE IPONo=" & Val(LV.SelectedItem.SubItems(3)))
    If lCount > 0 Then
        LotsAlreadyGenerated = True
    Else
        LotsAlreadyGenerated = False
    End If
        
End Function

Private Sub SaveLots(lRefID As Long, lMaxLotQty As Long, lQty As Long)

    Dim lLots As Long, lExtraLotQty As Long
    lExtraLotQty = 0
    lLots = Floor(lQty / lMaxLotQty)
    If lQty > (lLots * lMaxLotQty) Then
        lExtraLotQty = lQty - (lLots * lMaxLotQty)
    End If
    
    Dim i As Integer
    For i = 1 To lLots
        con.Execute "INSERT INTO InternalProductionOrderDetailLots(IPOD_RefID,LotQty) VALUES(" & _
         lRefID & "," & lMaxLotQty & ")"
    Next
    
    If lExtraLotQty > 0 Then
        con.Execute "INSERT INTO InternalProductionOrderDetailLots(IPOD_RefID,LotQty) VALUES(" & _
         lRefID & "," & lExtraLotQty & ")"
    End If
    
End Sub

Private Sub mnuNew_Click()
    frmNewIPO.Show
End Sub

Private Sub mnuPrintContractReview_Click()
    
    If LV.ListItems.count = 0 Then Exit Sub
    
    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ContractReview.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{InternalProductionOrder.IPONo}=" & Val(LV.SelectedItem.SubItems(3)), rpt
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub mnuPrintIPO_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\IPO.rpt")
    Dim f As New frmPrevRpt
    Dim rptSUB As CRAXDDRT.Report
    Set rptSUB = rpt.OpenSubreport("IPOLots")
    rpt.EnableParameterPrompting = False
    rptSUB.ParameterFields.GetItemByName("@IPONo").SetCurrentValue CLng(LV.SelectedItem.SubItems(3))
    f.ShowReport "{InternalProductionOrder.IPONo}=" & LV.SelectedItem.SubItems(3), rpt, , False
    
End Sub

Private Sub mnuPrintSpecs_Click()

    Dim f As New frmPrintSpecsLotWise
    Load f
    f.ShowMe (Val(LV.SelectedItem.SubItems(3)))
    
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
Private Sub RefreshSnos()
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Text = i
    Next i
End Sub
