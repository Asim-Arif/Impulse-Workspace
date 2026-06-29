VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmRMPOList 
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   12060
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   8340
   ScaleMode       =   0  'User
   ScaleWidth      =   10802.31
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   915
      TabIndex        =   7
      Top             =   0
      Width           =   10560
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PO List"
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
         Caption         =   "PO List"
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
      Height          =   7665
      Left            =   15
      TabIndex        =   3
      Top             =   615
      Width           =   12030
      Begin VB.CheckBox chkOpenPOs 
         Alignment       =   1  'Right Justify
         Caption         =   "Open PO Only"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   9315
         TabIndex        =   22
         Top             =   1260
         Width           =   1635
      End
      Begin ComboList.Usercontrol1 cmbMaterial 
         Height          =   285
         Left            =   5445
         TabIndex        =   17
         Top             =   360
         Width           =   6495
         _ExtentX        =   11456
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
      Begin ComboList.Usercontrol1 cmbRMGroup 
         Height          =   285
         Left            =   7470
         TabIndex        =   21
         Top             =   915
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
         Left            =   9780
         TabIndex        =   20
         Top             =   915
         Width           =   1170
      End
      Begin ComboList.Usercontrol1 cmbVender 
         Height          =   285
         Left            =   75
         TabIndex        =   6
         Top             =   360
         Width           =   5370
         _ExtentX        =   9472
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
      Begin VB.ComboBox cmbPOType 
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
         ItemData        =   "frmRMPOList.frx":0000
         Left            =   60
         List            =   "frmRMPOList.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   915
         Width           =   2115
      End
      Begin VB.CheckBox chkLate 
         Caption         =   "Show Late Orders"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9345
         TabIndex        =   15
         Top             =   945
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.CheckBox chkRunning 
         Caption         =   "Open POs"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9345
         TabIndex        =   13
         Top             =   945
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
         ItemData        =   "frmRMPOList.frx":0040
         Left            =   2175
         List            =   "frmRMPOList.frx":0056
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   915
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   5925
         TabIndex        =   1
         Top             =   915
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
         Format          =   148242435
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   4365
         TabIndex        =   0
         Top             =   915
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   148242435
         CurrentDate     =   37055
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5610
         Left            =   30
         TabIndex        =   10
         Top             =   1500
         Width           =   11910
         _ExtentX        =   21008
         _ExtentY        =   9895
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   900
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PO No."
            Object.Width           =   2434
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Dated"
            Object.Width           =   2037
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Vendor"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Delivery Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Type"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "PDF"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Status"
            Object.Width           =   2469
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   2445
         Top             =   2985
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
               Picture         =   "frmRMPOList.frx":00AB
               Key             =   "Down"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRMPOList.frx":0267
               Key             =   "Up"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         Caption         =   $"frmRMPOList.frx":0423
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
         TabIndex        =   18
         Top             =   675
         Width           =   10890
      End
      Begin MSForms.CommandButton cmdEndPO 
         Height          =   360
         Left            =   6945
         TabIndex        =   16
         Top             =   7215
         Visible         =   0   'False
         Width           =   1665
         Caption         =   "End This PO"
         PicturePosition =   327683
         Size            =   "2937;635"
         Accelerator     =   69
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdOptions 
         Height          =   360
         Left            =   8640
         TabIndex        =   14
         Top             =   7215
         Width           =   1620
         ForeColor       =   0
         Caption         =   "Options"
         PicturePosition =   327683
         Size            =   "2857;635"
         Accelerator     =   79
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdUpdate 
         Height          =   360
         Left            =   5280
         TabIndex        =   12
         Top             =   7215
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
         Left            =   10335
         TabIndex        =   11
         Top             =   7200
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
         Height          =   525
         Left            =   10995
         TabIndex        =   2
         Top             =   660
         Width           =   930
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "1640;926"
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
         Caption         =   $"frmRMPOList.frx":04F1
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
         Width           =   11850
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   -30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Please Select Picture"
      Filter          =   "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PopMenu"
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit This Order"
      End
      Begin VB.Menu d3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Order"
      End
      Begin VB.Menu mnuDashDelete 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print This Order"
      End
      Begin VB.Menu mnuPrintThisOrderStatus 
         Caption         =   "Print This Order (Status)"
      End
      Begin VB.Menu mnuPrintWithZeroRate 
         Caption         =   "Print This Order (Zero Rate)"
      End
      Begin VB.Menu mnuPrintOfficeCopy 
         Caption         =   "Print This Order (Office Copy)"
      End
      Begin VB.Menu mnuPrintVenderCopy 
         Caption         =   "Print This Order (Vender Copy)"
      End
      Begin VB.Menu mnuPrintAccountsCopy 
         Caption         =   "Print This Order (Accounts Copy)"
      End
      Begin VB.Menu mnuEndPO 
         Caption         =   "Close This &PO"
      End
      Begin VB.Menu mnuPrintWithReceiving 
         Caption         =   "Print Order with Receiving"
         Visible         =   0   'False
      End
      Begin VB.Menu mnudash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintListSimple 
         Caption         =   "Print Order List (&Simple)"
      End
      Begin VB.Menu mnuDashVenderwiseHistory 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAttachPDF 
         Caption         =   "Attach PDF File"
      End
   End
End
Attribute VB_Name = "frmRMPOList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rptCond As String, iColNo As Long
Dim bSample As Boolean

Private Sub cmdEndPO_Click()
    If MsgBox("Are you sure to end this purchase order?", vbQuestion + vbYesNo) = vbYes Then
        con.Execute "Update VendOrders Set Final=1 Where POrdNo='" & LV.SelectedItem.SubItems(1) & "'"
        Call RefreshLV
    End If
End Sub

Private Sub cmdOptions_Click()
    Me.PopUpMenu mnuPop, 255, FAcc.Left + cmdOptions.Left, cmdOptions.Top + FAcc.Top
End Sub

Private Sub cmdUpdate_Click()
    If LV.ListItems.count = 0 Then Exit Sub
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            Load frmVendRcvd
            Call frmVendRcvd.EditVendRcv(LV.SelectedItem.key)
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



Private Sub mnuProdPlan_Click()
    Load frmProdPlan
    Call frmProdPlan.AddPlanning(Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1))
End Sub
Private Sub cmbRange_Click()

    DtTo = GetServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    Select Case cmbRange.ListIndex
        Case 0
            DtFrom = DateAdd("d", -15, GetServerDate)
        Case 1
            DtFrom = DateAdd("d", -30, GetServerDate)
        Case 2
            DtFrom = DateAdd("d", -60, GetServerDate)
        Case 3
            DtFrom = DateAdd("d", -90, GetServerDate)
        Case 4
            DtFrom = DateAdd("yyyy", -1, GetServerDate)
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
    cmbVender.ListHeight = 6000
    cmbVender.AddItem "<All Venders>", "0"
    Call cmbVender.AddVals(con, "AccTitle", "VVenders", "AccNo")
    'Call cmbVender.AddVals(con, "AccTitle", "VVendersANDMakers", "AccNo")
    cmbVender.ID = "0"
    '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
    cmbMaterial.ListHeight = 6000
    cmbMaterial.AddItem "<All Materials>", "0"
    Call cmbMaterial.AddVals(con, "'{' + RMID1 + '} ' + RMName ", "RM", "RMID1", " ORDER BY RMName")
    cmbMaterial.ID = "0"
    '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
    
    cmbRMGroup.ListHeight = 3000
    cmbRMGroup.AddItem "<All Groups>", "0"
    cmbRMGroup.AddVals con, "Description", "RMGroups", "ID"
    cmbRMGroup.ID = "0"
    
    cmbRange.ListIndex = 0
        
    cmbPOType.ListIndex = 3
    bSample = False
    
End Sub

Private Sub Form_Resize()
    
    Exit Sub
    On Error Resume Next
    Dim TheHeight As Long
    Frame1.Left = (Me.ScaleWidth - Frame1.Width) / 2
    
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    
End Sub

Private Sub RefreshLV()
     
    If cmbVender.MatchFound = False Then
        MsgBox "Please Select Vender From List.", vbInformation
        Exit Sub
    ElseIf cmbMaterial.MatchFound = False Then
        MsgBox "Please Select Material.", vbInformation
        Exit Sub
    End If
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim Cond As String
    Dim DTRange As String
    Dim ITM As ListItem
    
    Screen.MousePointer = vbHourglass
    rptCond = ""
    If cmbVender.ID <> "0" Then
        Cond = " WHERE (VendID='" & cmbVender.ID & "') AND "
        rptCond = "{VVendOrders.VendID}='" & cmbVender.ID & "' AND "
    Else
        Cond = " WHERE "
    End If
        
    If cmbRange.ListIndex = 0 Then
        DTRange = " DT Between '" & DateAdd("d", -15, GetServerDate) & "' AND '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 1 Then
        DTRange = " DT Between '" & DateAdd("d", -30, GetServerDate) & "' AND '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 2 Then
        DTRange = " DT Between '" & DateAdd("d", -60, GetServerDate) & "' AND '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 3 Then
        DTRange = " DT Between '" & DateAdd("d", -90, GetServerDate) & "' AND '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 4 Then
        DTRange = " DT Between '" & DateAdd("yyyy", -1, GetServerDate) & "' AND '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 5 Then
        DTRange = " DT Between '" & DtFrom & "' AND '" & DtTo & "'"
    End If
        
    Cond = Cond & " (" & DTRange & ")"
    rptCond = rptCond & " ({VVendOrders.DT}=#" & Format(DtFrom.Value, "dd-MMM-yyyy") & "# to #" & Format(DtTo.Value, "dd-MMM-yyyy") & "#)"
    
    If chkRunning.Value = vbChecked Then
        Cond = Cond & " AND (OrderNo IN(SELECT OrderNo FROM VVendOrdersToRcv))"
        rptCond = rptCond & " AND {VVendOrdersToRcv.OrderNo}={VVendOrders.OrderNo} AND {@BalanceQty}>0"
    End If
    
    If chkLate.Value = vbChecked Then
        Cond = Cond & " AND DeliveryDT<=getDate()"
        rptCond = rptCond & " AND {VVendOrdersRcvd.DeliveryDT}<=CurrentDate"
    End If
    
    If cmbMaterial.ID <> "0" Then
        Cond = Cond & " AND OrderNo IN(SELECT OrderNo FROM VendOrderDetail WHERE RMID='" & cmbMaterial.ID & "')"
    End If
    
    If cmbRMGroup.ID <> "0" Then
        Cond = Cond & " AND OrderNo IN(SELECT OrderNo FROM VVendOrderDetail WHERE GroupID=" & cmbRMGroup.ID & ")"
    End If
    
    If cmbPOType.ListIndex < 3 Then
        Cond = Cond & " AND IssuanceType=" & cmbPOType.ListIndex
        rptCond = rptCond & " AND {VVendOrders.IssuanceType}=" & cmbPOType.ListIndex
    End If
    
    If Trim(txtPONo) <> "" Then
        Cond = " WHERE OrderNo='" & Trim(txtPONo) & "'"
        rptCond = rptCond & " AND {VVendOrders.OrderNo}='" & Trim(txtPONo) & "'"
    End If
        
    If bSample Then
        Cond = Cond & " AND SampleOrder=1"
    End If
    
    If chkOpenPOs.Value = vbChecked Then
        Cond = Cond & " AND RcvableOrderNo IS NOT NULL"
    End If
    
    Dim strPOType As String, strStatus As String
    With rs
        SQL = "SELECT * FROM VVendOrdersList " & Cond & " ORDER BY DT ASC"
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        'LV.Sorted = False
        Do Until .EOF
            
            Set ITM = LV.ListItems.add(, ![OrderNo], .AbsolutePosition)
            
            ITM.ListSubItems.add(, , ![OrderNo] & " (" & Val(!RunningPONo & "") & ")").Tag = !OrderNo & ""
            
            ITM.ListSubItems.add(, , Format(![DT], "dd-MMM-yyyy")).Tag = Val(!RunningPONo & "")
            
            ITM.ListSubItems.add , , ![AccTitle] & ""
            ITM.ListSubItems.add , , Format(![DeliveryDT], "dd-MMM-yyyy")
            
            If !IssuanceType = 0 Then
                strPOType = "Regular"
            ElseIf !IssuanceType = 1 Then
                strPOType = "Annealing"
            ElseIf !IssuanceType = 2 Then
                strPOType = "Machining"
            End If
            ITM.ListSubItems.add , , strPOType
            
            If Not IsNull(!PDFAttachment) Then
                ITM.ListSubItems.add , , "Attached"
            Else
                ITM.ListSubItems.add , , "-"
            End If
            
            If !Final Then
                strStatus = "Closed"
            ElseIf Val(!TotalQtyRcvd & "") = 0 Then
                strStatus = "Pending"
            ElseIf Val(!TotalQtyRcvd & "") < Val(!TotalQtyOrdered & "") Then
                strStatus = "Partially"
            Else
                strStatus = "Received"
            End If
            ITM.ListSubItems.add , , strStatus
            .MoveNEXT
        Loop
        .Close
    End With
'    If LV.ListItems.Count > 0 Then
'        LV.SetFocus
'        cmdEndPO.Visible = True
'    Else
'        cmdEndPO.Visible = True
'    End If
    
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
        LV.Sorted = True
    Else
        LV.ColumnHeaderIcons = Nothing
        Set LV.ColumnHeaderIcons = ImageList2
        LV.SortKey = ColumnHeader.Index - 1
        LV.SortOrder = lvwAscending
        ColumnHeader.Icon = "Down"
        LV.Sorted = True
    End If
    Call RefreshSnos
    'If LV.Sorted = False Then LV.Sorted = True
End Sub

Private Sub LV_DblClick()
        
    If LV.ListItems.count = 0 Then Exit Sub
    If getDBPassword("Edit_Purchase_Order") = False Then Exit Sub
    If iColNo = 7 Then
        If LV.SelectedItem.SubItems(6) = "Attached" Then
            Call ShowTechDrawing("VendOrders", "PDFAttachment", " WHERE OrderNo='" & LV.SelectedItem.ListSubItems(1).Tag & "'", LV.SelectedItem.ListSubItems(2).Tag, CD1)
        Else
            MsgBox "No PDF attached.", vbInformation
            Exit Sub
        End If
    Else
        If LV.SelectedItem.SubItems(7) = "Closed" Then
            MsgBox "PO is closed, Can't Edit.", vbInformation
            Exit Sub
        End If
        Dim f As Form
        If strCompany = "Tecno" Then
            Set f = New frmNewRMPO_Adv
        Else
            Set f = New frmNewRMPO
        End If
        Load f
        f.EditOrder (LV.SelectedItem.ListSubItems(1).Tag)
    End If
    
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        If x > Val(LV.ColumnHeaders(i).Left) And x < (Val(LV.ColumnHeaders(i).Left) + Val(LV.ColumnHeaders(i).Width)) Then
            iColNo = i
            Exit For
        End If
    Next
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If LV.ListItems.count = 0 Or Button <> 2 Then Exit Sub
    Me.PopUpMenu mnuPop
    
End Sub


Private Sub mnuAttachPDF_Click()

    On Error Resume Next
    
    Dim strPDFPath As String
    strPDFPath = ""
    CD1.Filter = "PDF Files Only|*.pdf"
    CD1.ShowSave
    If Not (err.Number = 32755) Then
        strPDFPath = CD1.FileName
    Else
        strPDFPath = ""
    End If
    If strPDFPath <> "" Then
        Dim Rec As New ADODB.Recordset
        Dim rsPic As New ADODB.Stream
        Rec.Open "SELECT OrderNo,PDFAttachment FROM VendOrders WHERE OrderNo='" & LV.SelectedItem.ListSubItems(1).Tag & "'", con, adOpenStatic, adLockOptimistic
        If Rec.RecordCount > 0 Then
            rsPic.Open
            rsPic.Type = adTypeBinary
            rsPic.LoadFromFile strPDFPath
            Rec.Fields(1) = rsPic.Read
            rsPic.Close
            Rec.Update
        End If
        
        Rec.Close
        
        Set Rec = Nothing
        Set rsPic = Nothing
        
        LV.SelectedItem.SubItems(6) = "Attached"
    End If
    
End Sub

Private Sub mnuDelete_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    If LV.SelectedItem.SubItems(7) = "Closed" Then
        MsgBox "PO is closed, Can't Delete.", vbInformation
        Exit Sub
    End If
    If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    If getDBPassword("Delete_Purchase_Order") = False Then Exit Sub
    Dim strPONo As String
    strPONo = LV.SelectedItem.ListSubItems(1).Tag
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(*)", "VendRcvdDetailPO", " WHERE PORefNo='" & strPONo & "'")
    If lCount > 0 Then
        MsgBox "Can't Delete, Already Received.", vbInformation
        Exit Sub
    End If
    con.Execute "DELETE FROM VendOrders WHERE OrderNo='" & strPONo & "'"
    MsgBox "Deleted.", vbInformation
    Call RefreshLV
    
End Sub

Private Sub mnuedit_Click()
    Call LV_DblClick
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
    
        
    rptFOrderList.FormulaFields(3).Text = "'" & cmbVender.Text & "'"
    rptFOrderList.FormulaFields(4).Text = "'" & strFromTo & "'"
    
    'frmPrevRpt.ShowReport "", rptFOrderList
    
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
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
    
   con.Execute "Update Vouchers Set [Description]='" & txtDesc & "' Where Sno=" & Val(LV.SelectedItem.key)
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
Private Sub mnuEndPO_Click()
    If MsgBox("Are you sure to end this purchase order?", vbQuestion + vbYesNo) = vbYes Then
        con.Execute "UPDATE VendOrders SET Final=1 WHERE OrderNo='" & LV.SelectedItem.key & "'"
        Call RefreshLV
    End If
End Sub

Private Sub mnuPrint_Click()
    Call PrintRMPO(False, "", 0)
End Sub

Private Sub PrintRMPO(Optional bZeroRate As Boolean, Optional strCopy As String, Optional iReportType As Integer = 0)

    Screen.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report
    
'    If strCopy = "Office Copy" Or strCopy = "Accounts Copy" Then
'        'Set rpt = rptApp.OpenReport(RptPath & "\RMPOAccountsCopy.rpt")
'        Set rpt = rptApp.OpenReport(RptPath & "\RMPO.rpt")
'    Else
'        Set rpt = rptApp.OpenReport(RptPath & "\RMPO.rpt")
'    End If
    If iReportType = 0 Then
        Set rpt = rptApp.OpenReport(RptPath & "\RMPO.rpt")
    ElseIf iReportType = 1 Then
        Set rpt = rptApp.OpenReport(RptPath & "\RMPOWP.rpt")
    ElseIf iReportType = 2 Then
        Set rpt = rptApp.OpenReport(RptPath & "\RMPOOfficeCopy.rpt")
    ElseIf iReportType = 3 Then
        Set rpt = rptApp.OpenReport(RptPath & "\RMPOVendorCopy.rpt")
    ElseIf iReportType = 4 Then
        Set rpt = rptApp.OpenReport(RptPath & "\RMPOAccountCopy.rpt")
    ElseIf iReportType = 5 Then
        Set rpt = rptApp.OpenReport(RptPath & "\RMPOWithReceiving.rpt")
    End If
    
    rpt.FormulaFields.GetItemByName("HideRate").Text = bZeroRate
    rpt.FormulaFields.GetItemByName("Copy").Text = "'" & strCopy & "'"
    Dim f As New frmPrevRpt
    f.ShowReport "{VendOrders.OrderNo}='" & LV.SelectedItem.ListSubItems(1).Tag & "' ", rpt

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub mnuPrintAccountsCopy_Click()
    Call PrintRMPO(False, "Accounts Copy", 4)
End Sub

Private Sub mnuPrintList_Click()
    Call PrintOrderList(True, True, True, False, False)
End Sub

Private Sub mnuPrintListImportedForeignCurr_Click()
    Call PrintOrderList(True, False, False, False, False, True)
End Sub

Private Sub mnuPrintListPO_Click()
    Call PrintOrderList(True, True, False, False, False)
End Sub

Private Sub PrintOrderList(Optional ImportedMaterial As Boolean = False, Optional LocalMaterial As Boolean = False, Optional SortyByDelivery As Boolean = False, Optional bSimple As Boolean = False, Optional bConsolidate As Boolean = False, Optional bForeignCurr As Boolean = False)
    
    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    Dim SelFor As String, rptHeading As String
    Dim rpt As New CRAXDDRT.Report, f As New frmPrevRpt

    If SortyByDelivery Then
        Set rpt = rptApp.OpenReport(RptPath & "\rptVendPOList.rpt")
    Else
        If cmbMaterial.ID <> "0" Then
            If bSimple = False Then
                rptCond = rptCond & " AND {Material.MaterialID}='" & cmbMaterial.ID & "'"
            Else
                rptCond = rptCond & " AND {Material.RMID1}='" & cmbMaterial.ID & "'"
            End If
        End If
        If bSimple Then
            'Set rpt = rptApp.OpenReport(RptPath & "\rptVendPOListSimple.rpt")
            Set rpt = rptApp.OpenReport(RptPath & "\rptVendPOList.rpt")
        ElseIf bConsolidate Then
            Set rpt = rptApp.OpenReport(RptPath & "\rptVendPOListItemWiseCon.rpt")
        Else
            Set rpt = rptApp.OpenReport(RptPath & "\rptVendPOListDateWise.rpt")
            If bForeignCurr Then
                rpt.FormulaFields.GetItemByName("ForeignCurr").Text = True
            End If
        End If
    End If
    
    If ImportedMaterial And LocalMaterial Then
         SelFor = rptCond
         rptHeading = " - Items"
    ElseIf ImportedMaterial Then
        SelFor = rptCond & " AND {VVendOrders.ImportVender}=True"
        rptHeading = " - Import"
    ElseIf LocalMaterial Then
        SelFor = rptCond & " AND {VVendOrders.ImportVender}=False"
        rptHeading = " - Local"
    Else
         'rptCond = "({VVendOrders.DT}=#" & Format(DtFrom.Value, "dd-MMM-yyyy") & "# to #" & Format(DtTo.Value, "dd-MMM-yyyy") & "#)"
         SelFor = rptCond
         rptHeading = " - Items"
    End If
    
    If chkRunning.Value Then
        rpt.FormulaFields(1).Text = "'Pending" & rptHeading & "'"
    Else
        rpt.FormulaFields(1).Text = "'All" & rptHeading & "'"
    End If
    
    rptCond = rptCond & " AND {VVendOrders.SampleOrder}=" & bSample
    f.ShowReport rptCond, rpt
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Private Sub mnuPrintListImported_Click()
    Call PrintOrderList(True, False, False, False, False)
End Sub

Private Sub mnuPrintListImportedSortedByDelivery_Click()
    Call PrintOrderList(True, False, True, False, False)
End Sub

Private Sub mnuPrintListLocal_Click()
    Call PrintOrderList(False, True, False, False, False)
End Sub

Private Sub mnuPrintListLocalSortedByDelivery_Click()
    Call PrintOrderList(False, True, True, False, False)
End Sub

Private Sub mnuPrintListSimple_Click()
    Call PrintOrderList(False, False, False, True, False)
End Sub

Private Sub mnuConsolidateLocal_Click()
    Call PrintOrderList(False, True, False, False, True)
End Sub

Private Sub mnuConsolidateImported_Click()
    Call PrintOrderList(True, False, False, False, True)
End Sub

Private Sub mnuPrintOfficeCopy_Click()
    Call PrintRMPO(False, "Office Copy", 2)
End Sub

Private Sub mnuPrintThisOrderStatus_Click()
    Call PrintRMPO(False, "", 5)
End Sub

Private Sub mnuPrintVenderCopy_Click()
    Call PrintRMPO(False, "Vender Copy", 3)
End Sub

Private Sub mnuPrintVenderwiseMaterialHistory_Click()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptProductionPath & "\VenderMaterialHistory.rpt")
    Dim strSelection As String
    strSelection = "{VendOrders.DT}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#"
    If cmbVender.ID <> "0" Then
        strSelection = strSelection & " AND {VVenders.VendID}=" & cmbVender.ID
    End If
    If cmbMaterial.ID <> "0" Then
        strSelection = strSelection & " AND {Material.MaterialID}='" & cmbMaterial.ID & "'"
    End If
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub mnuPrintWithReceiving_Click()

    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptProductionPath & "\PODetailWithRcving.rpt")
    
    Dim f As New frmPrevRpt
    f.ShowReport "{VendOrders.POrdNo}='" & LV.SelectedItem.SubItems(1) & "'", rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintWithZeroRate_Click()
    PrintRMPO True, "", 1
End Sub

Public Sub ShowMe(Optional p_bSample As Boolean = False)
    bSample = p_bSample
    If p_bSample Then
        Label1(8).Caption = "Samples PO List"
        Label1(9).Caption = "Samples PO List"
    End If
    Me.Show
End Sub
