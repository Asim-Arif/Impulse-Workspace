VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmVendRcvList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendor Recieving List"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11910
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleMode       =   0  'User
   ScaleWidth      =   10667.95
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10560
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Recieving List"
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
         TabIndex        =   8
         Top             =   165
         Width           =   10605
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Recieving List"
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
         TabIndex        =   9
         Top             =   180
         Width           =   10590
      End
   End
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   6555
      Left            =   15
      TabIndex        =   3
      Top             =   615
      Width           =   11850
      Begin ComboList.Usercontrol1 cmbRMGroup 
         Height          =   285
         Left            =   2820
         TabIndex        =   15
         Top             =   360
         Width           =   1845
         _ExtentX        =   3254
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
      Begin VB.TextBox txtPONo 
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
         Left            =   9255
         TabIndex        =   14
         Top             =   345
         Width           =   1500
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
               Picture         =   "frmVendRcvList.frx":0000
               Key             =   "Down"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVendRcvList.frx":01BC
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
         TabIndex        =   10
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
         ItemData        =   "frmVendRcvList.frx":0378
         Left            =   4665
         List            =   "frmVendRcvList.frx":038E
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   345
         Width           =   1620
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   7770
         TabIndex        =   1
         Top             =   345
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
         Format          =   152502275
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   6270
         TabIndex        =   0
         Top             =   345
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
         Format          =   152502275
         CurrentDate     =   37055
      End
      Begin ComboList.Usercontrol1 cmbVender 
         Height          =   285
         Left            =   75
         TabIndex        =   6
         Top             =   360
         Width           =   2745
         _ExtentX        =   4842
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
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2010
         Top             =   2070
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5235
         Left            =   30
         TabIndex        =   11
         Top             =   765
         Width           =   11715
         _ExtentX        =   20664
         _ExtentY        =   9234
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
            Text            =   "P.O. No."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Recieving ID"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Dated"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Vendor"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Total Qty"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Rej. Qty"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSForms.CommandButton cmdUpdate 
         Height          =   360
         Left            =   8460
         TabIndex        =   13
         Top             =   6090
         Visible         =   0   'False
         Width           =   1620
         VariousPropertyBits=   25
         Caption         =   "Update Changes"
         PicturePosition =   327683
         Size            =   "2857;635"
         Accelerator     =   85
         FontName        =   "Tahoma"
         FontEffects     =   1073750017
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   360
         Left            =   10155
         TabIndex        =   12
         Top             =   6090
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
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   345
         Left            =   10785
         TabIndex        =   2
         Top             =   330
         Width           =   960
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "1693;609"
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
         Caption         =   $"frmVendRcvList.frx":03DD
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
         TabIndex        =   4
         Top             =   120
         Width           =   11670
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PopMenu"
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit This Order"
         Visible         =   0   'False
      End
      Begin VB.Menu d1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuProdPlan 
         Caption         =   "Add Production Planning"
         Visible         =   0   'False
      End
      Begin VB.Menu d6 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSummary 
         Caption         =   "&History Of The Order (Summary)"
         Visible         =   0   'False
      End
      Begin VB.Menu d2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDetail 
         Caption         =   "&History Of The Order (Detail)"
         Visible         =   0   'False
      End
      Begin VB.Menu d3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintSummary 
         Caption         =   "Print Order (Summary)"
         Visible         =   0   'False
      End
      Begin VB.Menu d4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintDetail 
         Caption         =   "Print Order (Detail)"
         Visible         =   0   'False
      End
      Begin VB.Menu d5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print Daily Receiving for Selected Date"
      End
      Begin VB.Menu mnuPrintSelected 
         Caption         =   "Print Selected GRN"
      End
      Begin VB.Menu mnuPrintMakerGRNSlip 
         Caption         =   "Print Maker GRN Slip"
      End
      Begin VB.Menu mnuDashMakerGRN 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintMatReq 
         Caption         =   "Print Order Material Requsition"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuProcessOutPut 
         Caption         =   "Print Order Estimated Process Output"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintMatReqStock 
         Caption         =   "Print Stock Status For This Order"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintCost 
         Caption         =   "Print Order Cost Sheet"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuInspectionEntry 
         Caption         =   "Inspection Entry"
      End
      Begin VB.Menu mnuRejectionEntry 
         Caption         =   "Rejection Entry"
      End
      Begin VB.Menu mnuDashDeleteStart 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Selected Entry"
      End
      Begin VB.Menu mnuDeletedDashEnd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintListWithValue 
         Caption         =   "Print List (With Value)"
      End
   End
End
Attribute VB_Name = "frmVendRcvList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String
Dim bSample As Boolean
Private Sub cmdUpdate_Click()
    If LV.ListItems.count = 0 Then Exit Sub
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            Load frmVendRcvd
            Call frmVendRcvd.EditVendRcv(Replace(LV.SelectedItem.key, "'", ""))
        End If
    Next
    cmdUpdate.Enabled = False
End Sub

Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Not CBool(Item.Tag) Then
        Item.Checked = Not Item.Checked
    'Else
    '    If Item.Checked Then
    '        Item.SubItems(5) = Format(GetServerDate, "dd-MMM-yyyy")
    '    Else
    '        Item.SubItems(5) = "< Washing Pending >"
    '    End If
     cmdUpdate.Enabled = True
    End If
            
   
    
End Sub

Private Sub mnuDelete_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    Dim strRcvID As String
    strRcvID = LV.SelectedItem.SubItems(2)
    
    '
    
    Dim lCount As Long, lCount_Issued As Long
    lCount = GetSingleLongValue("COUNT(*)", "VMaterialLocationwiseStatus", " WHERE RcvID='" & strRcvID & "'")
    If lCount > 0 Then  'If Material is placed, checked for issuance
        lCount_Issued = GetSingleLongValue("COUNT(*)", "VMaterialLocationwiseStatus", " WHERE RcvID='" & strRcvID & "' AND QtyIssued>0")
    End If
    
    If lCount_Issued > 0 Then
        MsgBox "Can't Delete, Material Issued.", vbCritical
        Exit Sub
    End If
    
    If getDBPassword("Delete_RM_Receiving") = False Then Exit Sub
    
    con.Execute "DELETE FROM VendRcvd WHERE RcvID='" & strRcvID & "'"
    Call RefreshLV
    
End Sub

Private Sub mnuInspectionEntry_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    
    Dim f As New frmVendRcvdPOInspection
    Load f
    f.ShowMe LV.SelectedItem.SubItems(2)
    
End Sub

Private Sub mnuPrintListWithValue_Click()

    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\RMReceivingListWValue.rpt")
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, SQL
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintMakerGRNSlip_Click()

    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\GRN_MKR.rpt")
    
    Dim f As New frmPrevRpt
    f.ShowReport "{VendRcvd.RcvID}='" & LV.SelectedItem.SubItems(2) & "'", rpt
    
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub mnuPrintSelected_Click()

    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\GRN.rpt")
    
    Dim f As New frmPrevRpt
    f.ShowReport "{VendRcvd.RcvID}='" & LV.SelectedItem.SubItems(2) & "'", rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

'Private Sub mnuPrintCost_Click()
'Dim f As New frmRpt
'
'With f
'    Call .ShowOrderCostSheet(Replace(LV.SelectedItem.Key, "'", ""))
'End With
'
'
'Exit Sub
''''''''''''''''''''''''''''''''''
'With Cr1
'    .Connect = ConnectStr
'    .ReportFileName = RptPath & "\FOrderCost.rpt"
'    .Formulas(1) = ""
'    .SelectionFormula = "{VFOrderList.OrderNo}='" & Replace(LV.SelectedItem.Key, "'", "") & "'"
'    .Action = 1
'    .PageZoomNext
'End With
'    Screen.MousePointer = vbDefault
'''''''''''''''''''''''''''''''''''
'    Exit Sub
'ERR:
'    MsgBox ERR.Description
'End Sub
'
'Private Sub mnuPrintMatReq_Click()
'
'Dim f As New frmRpt
'
'Call f.ShowOrdMatReq(Replace(LV.SelectedItem.Key, "'", ""), False)
'
''With Cr1
''    .Connect = ConnectStr
''    .ReportFileName = RptPath & "\FOrdMatReq.rpt"
''    .SelectionFormula = "{VFOrderList.OrderNo}='" & Replace(LV.SelectedItem.Key, "'", "") & "'"
''    .SubreportToChange = "MaterialTotals"
''    .Formulas(1) = "StockStatus=False"
''
''    .Action = 1
''    .PageZoomNext
''    .Formulas(1) = ""
''    .SubreportToChange = ""
''
''End With
'
'    Screen.MousePointer = vbDefault
'Set f = Nothing
'    Exit Sub
'ERR:
'    MsgBox ERR.Description
'End Sub
'Private Sub mnuPrintMatReqStock_Click()
'
'
'Dim f As New frmRpt
'
'Call f.ShowOrdMatReq(Replace(LV.SelectedItem.Key, "'", ""), True)
'
''Dim cry As CrystalReport
''Set cry = Cr1
''With cry
''    .Connect = ConnectStr
''    .ReportFileName = RptPath & "\FOrdMatReq.rpt"
''    .SelectionFormula = "{VFOrderList.OrderNo}='" & Replace(LV.SelectedItem.Key, "'", "") & "'"
''    .SubreportToChange = "MaterialTotals"
''    .Formulas(1) = "StockStatus=True"
''    .Action = 1
''    .PageZoomNext
''    .Formulas(1) = ""
''    .SubreportToChange = ""
''End With
'    Screen.MousePointer = vbDefault
'
'Set cry = Nothing
'    Exit Sub
'ERR:
'    MsgBox ERR.Description
'End Sub
'Private Sub mnuPrintDetail_Click()
'With Cr1
'    .Connect = ConnectStr
'    .ReportFileName = RptPath & "\FCustOrders.rpt"
'    '.Formulas(1) = "ForCustomer='" & cmbVender.Text & "'"
'    .SelectionFormula = "{VrptOrders.OrderNo}='" & Replace(LV.SelectedItem.Key, "'", "") & "'"
'    .Action = 1
'    .PageZoomNext
'End With
'    Screen.MousePointer = vbDefault
'
'
'    Exit Sub
'ERR:
'    MsgBox ERR.Description
'End Sub
'
'Private Sub mnuProcessOutPut_Click()
'Dim f As New frmRpt
'
'Call f.ShowOrdProcessOutPut(Replace(LV.SelectedItem.Key, "'", ""))
'
'    Screen.MousePointer = vbDefault
'Set f = Nothing
'    Exit Sub
'ERR:
'    MsgBox ERR.Description
'
'End Sub
'
Private Sub mnuProdPlan_Click()
'    Load frmProdPlan
'    Call frmProdPlan.AddPlanning(Left(LV.SelectedItem.Key, Len(LV.SelectedItem.Key) - 1))
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
    
    mnuPop.Visible = False
    cmbVender.ListHeight = 2500
    cmbVender.AddItem "<All Venders>", "0"
    Call cmbVender.AddVals(con, "AccTitle", "VVenders", "AccNo")
    cmbVender.ID = "0"
    cmbRange.ListIndex = 0
    
    If strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
        mnuRejectionEntry.Visible = True
    Else
        mnuRejectionEntry.Visible = False
    End If
        
    bSample = False
    
    cmbRMGroup.ListHeight = 3000
    cmbRMGroup.AddItem "<All Groups>", "0"
    cmbRMGroup.AddVals con, "Description", "RMGroups", "ID"
    cmbRMGroup.ID = "0"
    
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
    
    If cmbVender.ID <> "0" Then
        Cond = " WHERE (VendID='" & cmbVender.ID & "') AND "
    Else
        Cond = " WHERE "
    End If
    
    DTRange = " RcvDate BETWEEN '" & DtFrom & "' and '" & DtTo & "'"
    
    Cond = Cond & " (" & DTRange & ")"
    
    If Trim(txtPONo) <> "" Then
        Cond = " WHERE PORefNo LIKE '%" & Trim(txtPONo) & "%'"
    End If
    
    Cond = Cond & " AND SampleOrder=" & Abs(bSample)
    
    If cmbRMGroup.ID <> "0" Then
        Cond = Cond & " AND RcvID IN(SELECT RcvID FROM VVendRcvdDetailPO WHERE GroupID=" & cmbRMGroup.ID & ")"
    End If
    
    With rs
        SQL = "SELECT * FROM VVenderRcvs " & Cond & " ORDER BY RcvDate Asc"
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        
        'LV.Sorted = False
        Do Until .EOF
            
            Set ITM = LV.ListItems.add(, !RcvID & "'", .AbsolutePosition)
            
            ITM.ListSubItems.add , , !PORefNo & ""
            
            ITM.ListSubItems.add , , !RcvID & ""
            
            ITM.ListSubItems.add , , Format(!RcvDate, "dd-MMM-yyyy")
            
            ITM.ListSubItems.add , , !VenderName & ""
            
            ITM.ListSubItems.add , , Round(Val(!TotalQty & ""), 2)
            ITM.ListSubItems.add , , Round(Val(!TotalQtyRejected & ""), 2)
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

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

If ColumnHeader.Index = 1 Then Exit Sub

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
    Exit Sub
    If CBool(LV.SelectedItem.Tag) Then
        Load frmVendRcvd
        Call frmVendRcvd.EditVendRcv(Replace(LV.SelectedItem.key, "'", ""))
    'Else
    '    Load frmViewRcvDetail
    '    Call frmViewRcvDetail.ShowRcvDetail(LV.SelectedItem.SubItems(1))
    
    End If
    
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If LV.ListItems.count = 0 Or Button <> 2 Then Exit Sub
    If bSample Then
        mnuInspectionEntry.Visible = False
        mnuRejectionEntry.Caption = "Accept/Reject Samples"
    End If
    
    Me.PopUpMenu mnuPop
    
    
    'If LV.ListItems.Count = 0 Or Button <> 2 Or OutOfFinancialYear Then Exit Sub

    'if a Voucher Entry Show Delete Menu
    'If InStr(LV.SelectedItem, "-") > 0 Then Me.PopupMenu mnuPop
    
End Sub

Private Sub mnuedit_Click()
    Call LV_DblClick
End Sub

Private Sub mnuPrint_Click()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\DailyRcving.rpt")
    Dim f As New frmPrevRpt
    Dim strSelection As String
    strSelection = "{VVendRcvd.RcvDate}=#" & DtFrom.Value & "# to #" & DtTo.Value & "#"
    strSelection = strSelection & " AND {VVendRcvd.SampleOrder}=" & bSample
    If cmbVender.ID <> "0" Then
        strSelection = strSelection & " AND {VVendRcvd.VendID}='" & cmbVender.ID & "'"
    End If
    
    f.ShowReport strSelection, rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintSummary_Click()
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
        strFromTo = Format(DateAdd("d", iNumber, GetServerDate), "dd-MMM-yyyy") & " To  " & Format(GetServerDate, "dd-MMM-yyyy")
    End If
    
    Dim rptFOrderList As CRAXDDRT.Report
    rptFOrderList.FormulaFields(3).Text = "'" & cmbVender.Text & "'"
    rptFOrderList.FormulaFields(4).Text = "'" & strFromTo & "'"
    
    frmPrevRpt.ShowReport "", rptFOrderList
    
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Private Sub mnuRejectionEntry_Click()
     
    Dim f As Form
    
    If bSample Then
        Set f = New frmVendRcvdPO_Sample
        Load f
        f.ShowMe LV.SelectedItem.SubItems(2)
    Else
        Set f = New frmVendRcvdPO_Rejection
        Load f
        f.ShowMe LV.SelectedItem.SubItems(2)
    End If
    
End Sub

Private Sub txtDesc_GotFocus()
    cmdRefresh.Default = False
    txtDesc.SelStart = 0
    txtDesc.SelLength = Len(txtDesc)
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
  
'If KeyAscii = 13 Then
'
'    If Trim(txtDesc = "") Then
'        MsgBox "Invalid Description", vbInformation
'        Exit Sub
'    End If
'
'   con.Execute "Update Vouchers Set [Desc]='" & txtDesc & "' Where Sno=" & Val(LV.SelectedItem.Key)
'   LV.SelectedItem.ListSubItems(3) = txtDesc & txtDesc.Tag
'   txtDesc.Visible = False
'   cmdRefresh.Default = False
'ElseIf KeyAscii = 27 Then
'   txtDesc.Visible = False
'   cmdRefresh.Default = False
'End If

End Sub

Private Sub txtDesc_LostFocus()
    txtDesc.Visible = False
End Sub
Private Sub RefreshSnos()
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Text = i
    Next i
End Sub


Public Sub ShowMe(Optional p_bSample As Boolean = False)
    bSample = p_bSample
    If bSample Then
        Label1(8).Caption = "Samples Receive List"
        Label1(9).Caption = "Samples Receive List"
    End If
    Me.Show
End Sub
