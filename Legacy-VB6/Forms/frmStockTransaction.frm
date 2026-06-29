VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmStockTransaction 
   ClientHeight    =   7695
   ClientLeft      =   75
   ClientTop       =   -30
   ClientWidth     =   12780
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   7695
   ScaleMode       =   0  'User
   ScaleWidth      =   11447.22
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Width           =   10560
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Transactions"
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
         TabIndex        =   2
         Top             =   165
         Width           =   10425
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Transactions"
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
         TabIndex        =   3
         Top             =   180
         Width           =   10410
      End
   End
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   7050
      Left            =   15
      TabIndex        =   0
      Top             =   615
      Width           =   12720
      Begin ComboList.Usercontrol1 cmbStore 
         Height          =   285
         Left            =   6075
         TabIndex        =   15
         Top             =   390
         Width           =   2505
         _ExtentX        =   4419
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
      End
      Begin ComboList.Usercontrol1 cmbItems 
         Height          =   285
         Left            =   60
         TabIndex        =   16
         Top             =   975
         Width           =   12585
         _ExtentX        =   22199
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
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Top             =   1065
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.CheckBox chkRcving 
         Caption         =   "Receiving Entries"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8610
         TabIndex        =   12
         Top             =   435
         Value           =   1  'Checked
         Width           =   1605
      End
      Begin VB.CheckBox chkIssuance 
         Caption         =   "Issuance Entries"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8610
         TabIndex        =   11
         Top             =   180
         Value           =   1  'Checked
         Width           =   1605
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
         ItemData        =   "frmStockTransaction.frx":0000
         Left            =   75
         List            =   "frmStockTransaction.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   375
         Width           =   2100
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5250
         Left            =   60
         TabIndex        =   4
         Top             =   1275
         Width           =   12585
         _ExtentX        =   22199
         _ExtentY        =   9260
         View            =   3
         LabelEdit       =   1
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   900
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Entry #"
            Object.Width           =   2558
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Rcvd By/Issd By"
            Object.Width           =   2461
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Rcvd From/Issd To"
            Object.Width           =   2884
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Order No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Item Code"
            Object.Width           =   1746
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Item Name"
            Object.Width           =   3122
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Qty"
            Object.Width           =   979
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Location"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   4155
         TabIndex        =   7
         Top             =   375
         Width           =   1920
         _ExtentX        =   3387
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
         Format          =   145358851
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   2205
         TabIndex        =   8
         Top             =   375
         Width           =   1920
         _ExtentX        =   3387
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
         Format          =   145358851
         CurrentDate     =   37055
      End
      Begin MSForms.CommandButton cmdPrintRackWiseStock 
         Height          =   360
         Left            =   5520
         TabIndex        =   20
         Top             =   6570
         Width           =   2655
         ForeColor       =   0
         Caption         =   "Print Rack Wise Stock"
         PicturePosition =   327683
         Size            =   "4683;635"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrintFinishedStock 
         Height          =   360
         Left            =   2775
         TabIndex        =   19
         Top             =   6570
         Width           =   2655
         ForeColor       =   0
         Caption         =   "Print Finished Stock"
         PicturePosition =   327683
         Size            =   "4683;635"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrintLocationWise 
         Height          =   360
         Left            =   60
         TabIndex        =   18
         Top             =   6570
         Width           =   2655
         ForeColor       =   0
         Caption         =   "Print Location Wise Stock"
         PicturePosition =   327683
         Size            =   "4683;635"
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
         Caption         =   " Item"
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
         Index           =   1
         Left            =   60
         TabIndex        =   17
         Top             =   720
         Width           =   12585
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   360
         Left            =   9315
         TabIndex        =   14
         Top             =   6600
         Width           =   1620
         ForeColor       =   0
         Caption         =   "Print"
         PicturePosition =   327683
         Size            =   "2857;635"
         Accelerator     =   80
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
         Caption         =   " Range                                   From                                 To                                       Store"
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
         Left            =   60
         TabIndex        =   10
         Top             =   150
         Width           =   8505
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   495
         Left            =   10230
         TabIndex        =   9
         Top             =   165
         Width           =   1470
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "2593;873"
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
         Height          =   360
         Left            =   11040
         TabIndex        =   5
         Top             =   6600
         Width           =   1620
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "2857;635"
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
End
Attribute VB_Name = "frmStockTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String
Dim bSample As Boolean

Private Sub cmbMaterial_matched()
    Call RefreshLV
End Sub

Private Sub cmbRange_Change()
    Call cmbRange_Click
End Sub

Private Sub cmbRange_Click()

    DtTo = GetServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    
    Select Case cmbRange.ListIndex
        Case 0
            DtFrom = GetServerDate
        Case 1
            DtFrom = DateAdd("d", -15, GetServerDate)
        Case 2
            DtFrom = DateAdd("d", -30, GetServerDate)
        Case 3
            DtFrom = DateAdd("d", -60, GetServerDate)
        Case 4
            DtFrom = DateAdd("d", -90, GetServerDate)
        Case 5
            DtFrom.Enabled = True
            DtTo.Enabled = True
            DtFrom.SetFocus
            Exit Sub
    End Select
    
End Sub


Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\StockTransactions.rpt")
    Dim f As New frmPrevRpt
    Dim strSelection As String
    strSelection = "{VStockTransactions.DT}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#"
        
    If chkIssuance.Value = vbUnchecked Then
        strSelection = strSelection & " AND {VStockTransactions.EntryType}<>0"
    End If
    
    If chkRcving.Value = vbUnchecked Then
        strSelection = strSelection & " AND {VStockTransactions.EntryType}<>1"
    End If
     
    If cmbItems.ID <> "0" Then
        strSelection = strSelection & " AND {VStockTransactions.ItemID}='" & cmbItems.ID & "'"
    End If
    
    strSelection = strSelection & " AND {VStockTransactions.SampleEntry}=" & Abs(bSample)
    
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub cmdPrintFinishedStock_Click()
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Items_FinishedStock_Detail.rpt")
    Dim f As New frmPrevRpt
    
    Dim strSelection As String
     If cmbItems.ID <> "0" Then
        strSelection = "{VItems_StockReport.ItemID}='" & cmbItems.ID & "'"
    End If
    
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub cmdPrintLocationWise_Click()
        
    Dim strStoreIDs As String
    Dim frm As New frmGetStores
    If frm.getData(strStoreIDs) = False Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\BinWiseStockStatusOfItems_Lot_Batch_Wise.rpt")
    
    Dim f As New frmPrevRpt
    Dim strSelection As String
    strSelection = "{VItemsWithShelfWiseStock.NetQty}>0"
    If strStoreIDs <> "All" Then
        strSelection = strSelection & " AND {VStoreShelfs.Store_RefID} IN[" & strStoreIDs & "]"
    End If
    
    If cmbItems.ID <> "0" Then
        strSelection = strSelection & " AND {VItems.ItemID}='" & cmbItems.ID & "'"
    End If
    
    f.ShowReport strSelection, rpt

End Sub

Private Sub cmdPrintRackWiseStock_Click()
        
    Dim strStoreIDs As String
    Dim frm As New frmGetStores
    If frm.getData(strStoreIDs) = False Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\RackWiseFinishStockStatus.rpt")
    
    Dim f As New frmPrevRpt
    Dim strSelection As String
    strSelection = "{VItemsWithShelfWiseStock.NetQty}>0"
    If strStoreIDs <> "All" Then
        strSelection = strSelection & " AND {VStoreShelfs.Store_RefID} IN[" & strStoreIDs & "]"
    End If
    
    If cmbItems.ID <> "0" Then
        strSelection = strSelection & " AND {VItems.ItemID}='" & cmbItems.ID & "'"
    End If
    
    f.ShowReport strSelection, rpt

End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub

Private Sub Form_Load()
    
    bSample = False
    cmbStore.ListHeight = 4000
    cmbStore.AddItem "<All Stores>", "0"
    
    If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
        cmbStore.AddVals con, "StoreName", "Stores", "EntryID"
        cmbStore.ID = "0"
    Else
        cmbStore.ID = "0"
        cmbStore.Enabled = False
    End If
    
    If strCompany = "Banzai" Then
        cmbStore.Enabled = False
    End If
    
    cmbItems.ListHeight = 4000
    cmbItems.DropDownOffLine = True
    cmbItems.AddItem "<All Articles>", "0"
    Call cmbItems.AddVals(con, "ItemID + ' ' + ItemName + ' { ' + ISNULL(Attributes,'') + ' }'", "Items", "ItemID", " WHERE InActive=0 ORDER BY ItemName")
    cmbItems.ID = "0"
    
    Call RefreshLV
    
    cmbRange.ListIndex = 0
    
    If strCompany = "Instrumed" Then
        chkIssuance.Visible = False
    End If
    
End Sub

Private Sub RefreshLV()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim DTRange As String
    
    Screen.MousePointer = vbHourglass
    
    DTRange = " WHERE DT BETWEEN '" & DtFrom & "' AND '" & DtTo & "'"
    
    If chkIssuance.Value = vbUnchecked Then
        DTRange = DTRange & " AND EntryType<>0"
    End If
    If chkRcving.Value = vbUnchecked Then
        DTRange = DTRange & " AND EntryType<>1"
    End If
    
    If cmbStore.ID <> "0" Then
        DTRange = DTRange & " AND Store_RefID=" & cmbStore.ID
    End If
    
    If cmbItems.ID <> "0" Then
        DTRange = DTRange & " AND ItemID='" & cmbItems.ID & "'"
    End If
    
    DTRange = DTRange & " AND SampleEntry=" & Abs(bSample)
    
    Dim strPrevEntryNo As String
    strPrevEntryNo = ""
    Dim LColor As Long
    With rs
        SQL = "SELECT * FROM VStockTransactions " & DTRange & " ORDER BY DT,EntryNo"
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        
        'LV.Sorted = False
        Do Until .EOF
        
            If strPrevEntryNo = "" Or strPrevEntryNo <> ![EntryNo] Then
                If !EntryType = 1 Then
                    LColor = vbRed
                Else
                    LColor = vbBlue
                End If
                Set ITM = LV.ListItems.add(, ![EntryNo] & "'", .AbsolutePosition)
                ITM.ForeColor = LColor
                ITM.ListSubItems.add , , !EntryNo
                ITM.ListSubItems.add , , Format(![DT] & "", "dd-MMM-yyyy")
            
                ITM.ListSubItems.add , , !RcvdBy
                ITM.ListSubItems.add , , !RcvdFrom
                For i = 1 To ITM.ListSubItems.count
                    ITM.ListSubItems(i).ForeColor = LColor
                Next
                
            End If
            
            Set ITM = LV.ListItems.add()
            ITM.Tag = !EntryNo & ""
            ITM.ListSubItems.add().Tag = Val(!ChildEntryID & "")
            ITM.ListSubItems.add().Tag = !EntryType
            ITM.ListSubItems.add
            ITM.ListSubItems.add
            ITM.ListSubItems.add , , !OrderNo
            ITM.ListSubItems.add , , !ITemID & ""
            ITM.ListSubItems.add , , !ItemName & ""
            ITM.ListSubItems.add , , IIf(!EntryType = 1, Val(!Qty & ""), -Val(!Qty & ""))
            ITM.ListSubItems.add , , !Location & ""
            
            For i = 1 To ITM.ListSubItems.count
                ITM.ToolTipText = ITM.Text
            Next
            strPrevEntryNo = !EntryNo
            
            .MoveNEXT
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

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    If LV.SelectedItem.ListSubItems(2).Tag = "" Then Exit Sub
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(10).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(10).Width
        .Text = LV.SelectedItem.SubItems(9)
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 2 Then Exit Sub
    If LV.ListItems.count = 0 Then Exit Sub
    If LV.SelectedItem.Text = "" Then Exit Sub
    
    Dim strRet As String
'    If LV.ListItems.count = 0 Then
'        strRet = ShowPopUpMenu1("&Issue New Material")
'    Else
'        strRet = ShowPopUpMenu1("&Issue New Material", "-", "&Print This Entry", "&Return Stock", "Print &List")
'    End If
    strRet = ShowPopUpMenu1("Print Slip")
    Select Case strRet
    
        Case "&Issue New Material"
        
            frmIssItemsSimple.Show
            
        Case "&Print This Entry"
        
'            Dim f As New frmPrevRpt
'            Dim rpt As CRAXDDRT.Report
'            Set rpt = rptApp.OpenReport(rptProductionPath & "\rptIssItemsSimple.rpt")
'            f.ShowReport "{IssItemsSimple.IssNo}='" & LV.SelectedItem.SubItems(1) & "' ", rpt
            
        Case "&Return Stock"
        
'            Dim frm As New frmReturnIssdItems
'            frm.ShowDetail (LV.SelectedItem.SubItems(1))
        
        Case "Print Slip"
                    
            Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
            If Left(LV.SelectedItem.SubItems(1), 3) = "RCV" Then
                Set rpt = rptApp.OpenReport(RptPath & "\Finished_Stock_Receiving_Slip.rpt")
            Else
                Set rpt = rptApp.OpenReport(RptPath & "\Finished_Stock_Issuance_Slip.rpt")
            End If
'            rpt.FormulaFields(1).Text = "'" & Format(DtFrom, "dd-MMM-yyyy") & " to " & Format(DtTo, "dd-MMM-yyyy") & "'"
            
            f.ShowReport "{VStockTransactions.EntryNo}='" & LV.SelectedItem.SubItems(1) & "'", rpt
        
            
    End Select
        
End Sub


Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If LV.SelectedItem.ListSubItems(2).Tag = 0 Then 'Issuance Entry
            con.Execute "UPDATE IssItemsSimpleDetail SET Location='" & txtEdit & "' WHERE EntryID=" & LV.SelectedItem.ListSubItems(1).Tag
            con.Execute "UPDATE ItemLedgerSimple SET ItemLedgerSimple.Location='" & txtEdit & "' FROM ItemLedgerSimple " & _
             "INNER JOIN IssItemsSimpleDetail ON ItemLedgerSimple.EntryID=IssItemsSimpleDetail.LedgerRefID WHERE IssItemsSimpleDetail.EntryID=" & LV.SelectedItem.ListSubItems(1).Tag
        ElseIf LV.SelectedItem.ListSubItems(2).Tag = 1 Then
            con.Execute "UPDATE RcvItemsSimpleDetail SET Location='" & txtEdit & "' WHERE EntryID=" & LV.SelectedItem.ListSubItems(1).Tag
            con.Execute "UPDATE ItemLedgerSimple SET ItemLedgerSimple.Location='" & txtEdit & "' FROM ItemLedgerSimple " & _
             "INNER JOIN RcvItemsSimpleDetail ON ItemLedgerSimple.EntryID=RcvItemsSimpleDetail.LedgerRefID WHERE RcvItemsSimpleDetail.EntryID=" & LV.SelectedItem.ListSubItems(1).Tag
        ElseIf LV.SelectedItem.ListSubItems(2).Tag = 2 Then
            'con.Execute "UPDATE StockOrderAdjustments SET Location='" & txtEdit & "' WHERE EntryID=" & LV.SelectedItem.ListSubItems(1).Tag
            con.Execute "UPDATE StockOrderAdjustmentsDetail SET Location='" & txtEdit & "' WHERE EntryID=" & LV.SelectedItem.ListSubItems(1).Tag
        End If
        LV.SelectedItem.ListSubItems(9) = txtEdit
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub

Public Sub ShowMe(Optional p_bSample As Boolean = False)
    bSample = p_bSample
    chkIssuance.Value = vbChecked
    chkRcving.Value = vbUnchecked
    chkIssuance.Visible = False
    chkRcving.Visible = False
    Me.Show
End Sub
