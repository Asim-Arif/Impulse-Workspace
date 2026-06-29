VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmVendor_Order_List 
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   13905
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   9885
   ScaleMode       =   0  'User
   ScaleWidth      =   12454.9
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   15
      TabIndex        =   3
      Top             =   0
      Width           =   13890
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Order List"
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
         TabIndex        =   4
         Top             =   165
         Width           =   13770
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Order List"
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
         TabIndex        =   5
         Top             =   180
         Width           =   13755
      End
   End
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   9225
      Left            =   15
      TabIndex        =   0
      Top             =   615
      Width           =   13890
      Begin VB.CheckBox chkFilterByDeliveryDT 
         Caption         =   "Filter By Delivery Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   10035
         TabIndex        =   25
         Top             =   690
         Width           =   2415
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   45
         TabIndex        =   17
         Top             =   8670
         Width           =   3960
         Begin VB.TextBox txtFind 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1965
            TabIndex        =   20
            Top             =   135
            Width           =   1320
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3285
            TabIndex        =   19
            Top             =   135
            Width           =   615
         End
         Begin VB.ComboBox cmbfield 
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
            ItemData        =   "frmVendor_Order_List.frx":0000
            Left            =   60
            List            =   "frmVendor_Order_List.frx":000D
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   135
            Width           =   1875
         End
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
         ItemData        =   "frmVendor_Order_List.frx":002F
         Left            =   8460
         List            =   "frmVendor_Order_List.frx":0048
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   330
         Width           =   1530
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   11490
         TabIndex        =   15
         Top             =   330
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
         Format          =   171048963
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   9990
         TabIndex        =   16
         Top             =   330
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
         Format          =   171048963
         CurrentDate     =   37055
      End
      Begin VB.TextBox txtTotalAmt 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5760
         TabIndex        =   13
         Top             =   8775
         Width           =   1425
      End
      Begin ComboList.Usercontrol1 cmbCompany 
         Height          =   285
         Left            =   45
         TabIndex        =   11
         Top             =   345
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
               Picture         =   "frmVendor_Order_List.frx":00A7
               Key             =   "Down"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVendor_Order_List.frx":0263
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
         TabIndex        =   6
         Top             =   1605
         Visible         =   0   'False
         Width           =   3615
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   1905
         TabIndex        =   2
         Top             =   345
         Width           =   2940
         _ExtentX        =   5186
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
         Left            =   3375
         Top             =   2700
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin Crystal.CrystalReport Cr1 
         Left            =   3165
         Top             =   3270
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
         Height          =   7725
         Left            =   45
         TabIndex        =   7
         Top             =   990
         Width           =   13755
         _ExtentX        =   24262
         _ExtentY        =   13626
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
         NumItems        =   15
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Company"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Customer"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Country"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Order #"
            Object.Width           =   1746
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Int. Ord"
            Object.Width           =   1482
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Order DT"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Delivery DT"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Order Amt"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "Bal. Ord. Amt"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Status"
            Object.Width           =   1879
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   11
            Text            =   "Weight"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   12
            Text            =   "Qty"
            Object.Width           =   1138
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   13
            Text            =   "Articles"
            Object.Width           =   1453
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "App. Status"
            Object.Width           =   2469
         EndProperty
      End
      Begin MSForms.CommandButton cmdBackOrderUrgent 
         Height          =   375
         Left            =   7515
         TabIndex        =   24
         Top             =   8775
         Visible         =   0   'False
         Width           =   1845
         ForeColor       =   0
         Caption         =   "Backorder (Urgent)"
         PicturePosition =   327683
         Size            =   "3254;661"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbOrderType 
         Height          =   285
         Left            =   6390
         TabIndex        =   23
         Top             =   345
         Width           =   2055
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3625;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CommandButton cmdOrderStatus1 
         Height          =   375
         Left            =   9405
         TabIndex        =   22
         Top             =   8775
         Visible         =   0   'False
         Width           =   1800
         ForeColor       =   0
         Caption         =   "Backorder Report"
         PicturePosition =   327683
         Size            =   "3175;661"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrintOrderStatus 
         Height          =   375
         Left            =   11235
         TabIndex        =   21
         Top             =   8775
         Visible         =   0   'False
         Width           =   1530
         ForeColor       =   0
         Caption         =   "Order Status"
         PicturePosition =   327683
         Size            =   "2699;661"
         Accelerator     =   83
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Amount (Rs.):"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4050
         TabIndex        =   12
         Top             =   8835
         Width           =   1695
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   390
         Left            =   12975
         TabIndex        =   10
         Top             =   270
         Width           =   840
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "1482;688"
         Accelerator     =   82
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbStatus 
         Height          =   285
         Left            =   4860
         TabIndex        =   9
         Top             =   345
         Width           =   1515
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2672;503"
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
         Left            =   12795
         TabIndex        =   8
         Top             =   8775
         Width           =   1005
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "1773;661"
         Accelerator     =   67
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
         Caption         =   $"frmVendor_Order_List.frx":041F
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
         Left            =   45
         TabIndex        =   1
         Top             =   120
         Width           =   12915
      End
   End
End
Attribute VB_Name = "frmVendor_Order_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQLString As String
Dim bForProduction As Boolean, bSample As Boolean

Private Sub mnuProdPlan_Click()
    Load frmProdPlan
    Call frmProdPlan.AddPlanning(Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1))
End Sub

Public Sub ShowMe(p_bForProduction As Boolean, Optional p_bSample As Boolean = False)

    bSample = p_bSample
    
    If bSample Then
        cmbOrderType.ListIndex = 2
        cmbOrderType.Enabled = False
    End If
    
    bForProduction = p_bForProduction
    If bForProduction Then
        Label2.Visible = False
        txtTotalAmt.Visible = False
        cmdOrderStatus1.Visible = False
    End If
     
    Me.Show
    Call RefreshLV
    
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

Private Sub cmdBackOrderUrgent_Click()
    Call PrintOrderStatusReportProcessWise(True, True)
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdFind_Click()
    
    Dim StartIdx As Integer
    
    If LV.SelectedItem Is Nothing Then
       StartIdx = 1
    Else
       If LV.SelectedItem.Index = LV.ListItems.count Then
          StartIdx = 1
       Else
          StartIdx = LV.SelectedItem.Index
       End If
    End If

    Dim fldValue As String

    For i = StartIdx To LV.ListItems.count
       
        fldValue = LV.ListItems(i).SubItems(cmbfield.ListIndex + 2)
       
        If LCase(fldValue) Like "*" & LCase(Trim(txtFind)) & "*" Then
            LV.ListItems(i).Selected = True
            LV.ListItems(i).EnsureVisible
            If i <> StartIdx Then Exit For
        End If
       
    Next
    Dim rpt As New CRAXDDRT.Report
    
    If i = LV.ListItems.count + 1 Then
        MsgBox "Record Not Found.", vbInformation
        Set LV.SelectedItem = Nothing
    End If

End Sub

Private Sub cmdOrderStatus1_Click()
    PrintOrderStatusReportProcessWise (True)
End Sub

Private Sub cmdPrintOrderStatus_Click()
    Call PrintOrderStatusReportProcessWise
End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub


Private Sub Form_Load()
    
    'mnuPop.Visible = False
    bSample = False
    bForProduction = False
    cmbCompany.ListHeight = 4000
    cmbCompany.AddItem "< All Companies>", "0"
    cmbCompany.AddVals con, "CompanyName", "Companies", "EntryID", " Order By CompanyName"
    cmbCompany.ID = "0"
    
    cmbCust.ListHeight = 2500
    cmbCust.AddItem "<All Vendors>", "0"
    'Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where custCode<>'Hico'")
    Call cmbCust.AddVals(con, "AccTitle + ' {'+AccNo+'}'", "VVenders", "AccNo")
    With cmbStatus
        .AddItem "<Unshipped>"
        .AddItem "<Shipped>"
        .AddItem "<Cancelled>"
        .AddItem "<Hold>"
        .AddItem "<All Orders>"
    End With
    
    With cmbOrderType
        .AddItem "<All Types>"
        .AddItem "Stock Order"
        .AddItem "Sample Order"
        .AddItem "Customer Order"
        .AddItem "Replacement Order"
        .ListIndex = 0
    End With
    
    cmbStatus.ListIndex = 0
    
    
    cmbCust.ID = "0"
    
    Call RefreshLV
    
    cmbRange.ListIndex = 5
    cmbfield.ListIndex = 2
    
    
    
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
    
    If chkFilterByDeliveryDT.Value = vbChecked Then
        Cond = " WHERE DeliveryDT BETWEEN '" & DtFrom.Value & "' AND '" & DtTo.Value & "'"
    Else
        Cond = " WHERE DT BETWEEN '" & DtFrom.Value & "' AND '" & DtTo.Value & "'"
    End If
    
    If cmbCust.ID <> "0" Then
        Cond = Cond & " AND Vend_AccNo='" & cmbCust.ID & "' "
    End If
    
     If cmbCompany.ID <> "0" Then
        Cond = Cond & " AND CompanyRefID=" & cmbCompany.ID
    End If
     
    If cmbOrderType.ListIndex > 0 Then
        Cond = Cond & " AND OrderType='" & cmbOrderType.Text & "'"
    End If
    
    Dim sTmp As String
'    If cmbStatus.ListIndex = 0 Then                 'UnShipped.
'        'sTmp = " (TotalShippedQty<TotalOrderQty AND ISNULL(Cancelled,0)=0)"
'        sTmp = " (UnShippedOrderNo IS NOT NULL)"
'    ElseIf cmbStatus.ListIndex = 1 Then
'        'sTmp = " (TotalShippedQty=TotalOrderQty OR Cancelled=0)"
'        sTmp = " (UnShippedOrderNo IS NULL)"
'    ElseIf cmbStatus.ListIndex = 2 Then             'Cancelled
'        sTmp = " Cancelled=1"
'    ElseIf cmbStatus.ListIndex = 3 Then             'Hold
'        sTmp = " Cancelled=2"
'    ElseIf cmbStatus.ListIndex = 4 Then
'        sTmp = ""
'    End If
        
    If sTmp <> "" Then Cond = Cond & " And "
    Cond = Cond & sTmp
    
    SQLString = "SELECT * FROM VVendor_Orders_List " & Cond
    Dim sStatus As String
    Dim iRecNo As Integer
    Dim dTotalAmt As Double, dTotalAmtPKR As Double
    dTotalAmt = 0
    With rs
        
        .Open SQLString & " ORDER BY DT DESC", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        iRecNo = .RecordCount
        'LV.Sorted = False
        Dim strPurchasePlanStatus As String
        Do Until .EOF
            
            Set ITM = LV.ListItems.add(, ![OrderNo] & "'", iRecNo - .AbsolutePosition + 1)
            ITM.Tag = ![Curr] & ""
            ITM.ListSubItems.add , , ![CompanyName] & ""
            ITM.ListSubItems.add , , ![AccTitle] & ""
            'ITM.SubItems(1) = ![CustCode] & ""
            ITM.ListSubItems.add , , ![Country] & ""
            'ITM.SubItems(2) = ![Country] & ""
            ITM.ListSubItems.add , , ![OrderNo] & ""
            ITM.ListSubItems.add , , ![InternalRefNo] & ""
            'ITM.SubItems(3) = ![OrderNo] & ""
            ITM.ListSubItems.add , , Format(![DT] & "", "dd-MMM-yyyy")
            'ITM.SubItems(4) = Format(![Dt] & "", "dd-MMM-yyyy")
            ITM.ListSubItems.add , , Format(![DeliveryDT] & "", "dd-MMM-yyyy")
            
            If bForProduction Then
                ITM.ListSubItems.add(, , "-").Tag = Val(![OrderAmt] & "")
                ITM.ListSubItems.add(, , "-").Tag = Val(![OrderAmt] & "") - Val(!TotalShippedAmt & "")
            Else
                ITM.ListSubItems.add(, , ![Curr] & " " & Val(![OrderAmt] & "")).Tag = Val(![OrderAmt] & "")
                ITM.ListSubItems.add(, , Val(![OrderAmt] & "") - Val(!TotalShippedAmt & "")).Tag = Val(![OrderAmt] & "") - Val(!TotalShippedAmt & "")
            End If
            'ITM.SubItems(5) = Val(![OrderAmt] & "")
            
            If Val(![TotalShippedQty] & "") = 0 Then
                sStatus = "UnShipped"
            ElseIf Val(![TotalOrderQty] & "") > Val(![TotalShippedQty] & "") Then
                sStatus = "Partialy Shipped"
            Else
                sStatus = "Shipped"
            End If
            
            If IsNull(!Cancelled) = False Then
                If Val(!Cancelled & "") = 1 Then
                    sStatus = "Cancelled"
                ElseIf Val(!Cancelled & "") = 2 Then
                    sStatus = "Hold"
                Else
                    If IsNull(!UnShippedOrderNo) Then
                        sStatus = "Shipped"
                    Else
                        sStatus = "UnShipped"
                    End If
                End If
            End If
            
            ITM.ListSubItems.add(, , sStatus).ToolTipText = !Remarks & ""
            ITM.ListSubItems.add(, , Val(!TotalWeight & "")).ForeColor = vbRed
            
            ITM.ListSubItems.add(, , Val(!TotalOrderQty & "")).ForeColor = vbRed
            ITM.ListSubItems.add(, , Val(!TotalArticles & "")).ForeColor = vbRed
            
            dTotalAmt = dTotalAmt + Val(!OrderAmt & "")
            dTotalAmtPKR = dTotalAmtPKR + (Val(!OrderAmt & "") * Val(!ExchRate & ""))
            
            If !OrderPlanApproved Then
                strPurchasePlanStatus = "Approved"
            ElseIf Val(!TotalPlannedQty & "") > 0 Then
                strPurchasePlanStatus = "Planned"
            Else
                strPurchasePlanStatus = "-"
            End If
            ITM.ListSubItems.add , , strPurchasePlanStatus
            
            .MoveNEXT
            
        Loop
        
        .Close
        
    End With
    
    Set rs = Nothing
    txtTotalAmt = Format(dTotalAmtPKR, "###,##0.00")
    
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If LV.SortKey + 1 = ColumnHeader.Index Then
        LV.ColumnHeaderIcons = Nothing
        Set LV.ColumnHeaderIcons = ImageList2
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
    If LV.Sorted = False Then LV.Sorted = True
    
End Sub

Private Sub LV_DblClick()
    
    'If LV.ListItems.Count = 0 Or OutOfFinancialYear Then Exit Sub
    'If InStr(LV.SelectedItem, "-") > 0 Then Call mnuEditDesc_Click
    If bForProduction And strCompany = "IAA" Then Exit Sub
    
'    If bForProduction And Not (strCompany = "Weldon Industries" Or strCompany = "All Seas" Or strCompany = "Tecno" Or strCompany = "Dr-Frgz" Or strCompany = "Towne") Then
'        MsgBox "You don't have permissions.", vbInformation
'        Exit Sub
'    End If
    On Error GoTo err
    Dim OrderNo As String
    Dim iTemp As Integer
    
    If LV.ListItems.count = 0 Then Exit Sub
    
    
    OrderNo = Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1)
    
'    iTemp = Val(con.Execute("Select Count(OrderNo) From VFItemsQtyToPlan Where RemQty>0 and OrderNo='" & OrderNo & "'").Fields(0).value & "")
'
'    If iTemp = 0 Then
'        MsgBox "Can't Edit This Order.It Has Been Planned.", vbInformation
'        Exit Sub
'    End If
      
    If strCompany = "IAA" Or strCompany = "Kami" Then
        If getDBPassword("EditExportDocs") = False Then Exit Sub
    End If
    
    Load frmVendor_Order_AE
    Call frmVendor_Order_AE.EditOrder(Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1))
   
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
            
    If LV.ListItems.count = 0 Or Button <> 2 Then Exit Sub
 
    Dim strRet As String
'    If UserHasAccess("AddProdPlan") Then
'        strRet = ShowPopUpMenu1("&Edit This Order", "-", "&Add Production Planning", "-", "History Of The Order (Summary)", "History Of The Order (Detail)", "-", "Print Order (&Summary)", "Print Order (&Detail)", "-", "&Print Order", "-", "Print Order &Status")
'    Else
'        strRet = ShowPopUpMenu("&Edit This Order", "-", "History Of The Order (Summary)", "History Of The Order (Detail)", "-", "Print Order (&Summary)", "Print Order (&Detail)", "-", "&Print Order", "-", "Print Order &Status")
'    End If
    
    Dim strPrintOrderListCaption As String
    If strCompany = "IAA" Or strCompany = "Kami" Then
        If bForProduction Then
            strPrintOrderListCaption = "-"
        Else
            strPrintOrderListCaption = "Print Order Balance List"
        End If
    Else
        strPrintOrderListCaption = "Print Order List"
    End If
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(OrderNo)", "FCustomerFinalOrders", " WHERE OrderNo='" & LV.SelectedItem.SubItems(4) & "'")
    
    
    If strCompany = "IAA" Or strCompany = "Kami" Then
    
        If lCount > 0 Then
            'strRet = ShowPopUpMenu1("&Edit This Order", "-", "&Print Order", "Print Order (Balance)", "Print Order (With Value)", "Print Order (Groupwise)", "Print Order (Summary Stock)", "Print Order (Marking Plan)", "Print Production Order", "-", strPrintOrderListCaption, "-", "Print Order (Balance Quantities)", "-", "Print Balance Order", "-", "Material Consumption Sheet", "Material Consumption Sheet (Articlewise)", "Print Customer Order", "Order Planning", "Print Order Planning Report", "Print Quotatoin Planning Report", "-", "Stock Order Adjustment", "Un-Cancel/Un-Hold Order", "Combined Order Report (Item wise)", "Combined Order Report (Item Groupwise)", "Combined Order Report (Order wise)", "Delivery Analysis", "Order Status Log", "Order Status Log (New)", "Planning Purchase Log Sheet", "Production Timeline Report", "Production Lots", IIf(bForProduction, "-", "Shipping Dates"), "Urgent Items", "Semi Finish Stock", "Finish Stock", "Urgent Request", "Score Board", "Orders Log Sheet")
            strRet = ShowPopUpMenu1("Copy Order", "Delete Order", "&Edit This Order", "-", "&Print Order", "Print Order Status", "Order Balance Report", "Print Order (Balance)", "-", strPrintOrderListCaption, "-", "Print Order (Balance Quantities)", "-", "Print Balance Order", "Purchase Order Planning", "Print Order Planning Report", "Hub Wise Status", "Unplanned Items", "-", "Stock Order Adjustment", "Un-Cancel/Un-Hold Order")
        Else
            'strRet = ShowPopUpMenu1("&Edit This Order", "-", "&Print Order", "Print Order (Balance)", "Print Order (With Value)", "Print Order (Groupwise)", "Print Order (Summary Stock)" _
             , "Print Order (Marking Plan)", "Print Production Order", "-", strPrintOrderListCaption, "-", "Print Order (Balance Quantities)", "-", "Print Balance Order", "-", "Material Consumption Sheet", "Material Consumption Sheet (Articlewise)", "Print Customer Order", "Order Planning", "Print Order Planning Report", "Print Quotatoin Planning Report", "-", "Stock Order Adjustment", IIf(bForProduction, "-", "Mark As Shipped"), IIf(bForProduction, "-", "Cancel Order"), IIf(bForProduction, "-", "Hold Order"), "Combined Order Report (Item wise)", "Combined Order Report (Item Groupwise)", "Combined Order Report (Order wise)", "Delivery Analysis", "Order Status Log", "Order Status Log (New)", "Planning Purchase Log Sheet", "Production Timeline Report", "Production Lots", IIf(bForProduction, "-", "Shipping Dates"), "Urgent Items", "Semi Finish Stock", "Finish Stock", "Urgent Request", "Score Board", "Orders Log Sheet")
            strRet = ShowPopUpMenu1("Copy Order", "Delete Order", "&Edit This Order", "-", "&Print Order", "Print Order Status", "Print Order (Balance)" _
             , strPrintOrderListCaption, "-", "Print Order (Balance Quantities)", "-", "Order Balance Report", "Print Balance Order", "-", "Purchase Order Planning", "Print Order Planning Report", "Hub Wise Status", "Unplanned Items", "-", "Stock Order Adjustment", IIf(bForProduction, "-", "Mark As Shipped"), IIf(bForProduction, "-", "Cancel Order"), IIf(bForProduction, "-", "Hold Order"))
        End If
        
    End If
    
    Select Case strRet
        Case "Copy Order"
            Call CopyOrder(LV.SelectedItem.SubItems(4))
        Case "Delete Order"
            If DeleteOrder(LV.SelectedItem.SubItems(4)) Then
                Call RefreshLV
            End If
        Case "&Edit This Order"
            Call EditOrder
'        Case "Delete Order"
'            Call DeleteOrder
        Case "&Add Production Planning"
            Call AddProdPlanning
        Case "History Of The Order (Summary)"
            
        Case "History Of The Order (Detail)"
            
        Case "Print Order (&Summary)"
            Call PrintSummary
        Case "Print Order List"
            Call PrintOrderList
        Case "Print Order Balance List"
            Call PrintOrderList
        Case "Print Order (&Detail)"
        Case "Print Order (Balance Quantities)"
            Call PrintOrderBalanceReport(LV.SelectedItem.SubItems(4))
        Case "&Print Order"
            Call PrintOrder
        Case "Print Order Status"
            Call PrintOrder(False, False, False, False, False, False, False, True)
        Case "Print Order (Balance)"
            Call PrintOrder(True)
        Case "Print Order (With Value)"
            Call PrintOrder(, , , , , True)
        Case "Print Order (Groupwise)"
            Call PrintOrder(False, False, False, True)
        Case "Print Order (Summary Stock)"
            Call PrintOrder(False, False, True)
        Case "Print Order (Marking Plan)"
            Call PrintOrder(False, False, False, False, True)
        Case "Print Production Order"
            Call PrintOrder(False, False, False, False, False, False, True)
        Case "Print Order &Makerwise"
            Call PrintOrder(False, True)
        Case "Print Balance Order"
            Call PrintOrder(True)
        Case "Print PTC"
            Call PrintPTC
        Case "Material Consumption Sheet"
            Call PrintMaterialConsumptionSheet(LV.SelectedItem.SubItems(4))
        Case "Material Consumption Sheet (Articlewise)"
            Call PrintMaterialConsumptionSheet(LV.SelectedItem.SubItems(4), True)
        Case "Print Customer Order"
            Call PrintCustomerOrder
        Case "Stock Order Adjustment"
            If strCompany = "IAA" Or strCompany = "Kami" Then
                If getDBPassword("StockOrderAdjustment") = False Then Exit Sub
            End If
            Set f = New frmStockOrderAdjustment
            Load f
            f.ShowMe (LV.SelectedItem.SubItems(4))
        Case "Mark As Shipped", "Cancel Order", "Hold Order", "Un-Cancel/Un-Hold Order"

            If MsgBox("Are you sure ?", vbQuestion + vbYesNo) = vbYes Then
                
                If strRet = "Un-Cancel/Un-Hold Order" Then
                    con.Execute "DELETE FROM FCustomerFinalOrders WHERE OrderNo='" & LV.SelectedItem.SubItems(4) & "'"
                Else
                    Dim iCancelled As Integer
                    If strRet = "Cancel Order" Then
                        iCancelled = 1
                        If getDBPassword("DeleteExportEntries") = False Then Exit Sub
                    ElseIf strRet = "Hold Order" Then
                        iCancelled = 2
                        Dim strHoldRemarks As String, frmGet As New frmGetSingleTextValue
                        If frmGet.getTextValue(strHoldRemarks) = False Then Exit Sub
                    Else
                        iCancelled = 0
                    End If
                    con.Execute "INSERT INTO FCustomerFinalOrders(CustCode,Country,OrderNo,UserName,MachineName,Cancelled,Remarks) VALUES('" & _
                     LV.SelectedItem.SubItems(2) & "','" & LV.SelectedItem.SubItems(3) & "','" & LV.SelectedItem.SubItems(4) & _
                     "','" & CurrentUserName & "','" & strComputerName & "'," & iCancelled & ",'" & strHoldRemarks & "')"
                End If
            End If
        Case "Purchase Order Planning"
            If LV.SelectedItem.SubItems(14) = "Planned" Then
                If getDBPassword("Edit_Order_Planning") = False Then Exit Sub
            End If
            Set f = New frmOrderPlanning
            Load f
            f.ShowMe (LV.SelectedItem.SubItems(4))
        Case "Print Order Planning Report"
            Call PrintOrderPlanningReport(LV.SelectedItem.SubItems(4))
        Case "Print Quotatoin Planning Report"
            Call PrintOrderPlanningReport(LV.SelectedItem.SubItems(4), True)
        Case "Copy Customer Order"
            Call CopyCustomerOrder
        Case "Combined Order Report (Item wise)"
            Call PrintCombinedOrderReportNew
        Case "Combined Order Report (Item Groupwise)"
            Call PrintCombinedOrderReport(True)
        Case "Combined Order Report (Order wise)"
            Call PrintCombinedOrderReportAll
        Case "Print Production Export Cost"
            Call PrintProductionExportCost
        Case "Delivery Analysis"
            Call PrintDeliveryAnalysis(LV.SelectedItem.SubItems(4))
        Case "Print All PTC"
            Call PrintAllPTC(LV.SelectedItem.SubItems(4))
        Case "R1-Order Planning"
            Call PrintOrderPlanningReportR1(LV.SelectedItem.SubItems(4), True)
        Case "R2-Order Planning"
            Call PrintOrderPlanningReportR1(LV.SelectedItem.SubItems(4), False)
        Case "Print Delivery Schedule"
            Call PrintOrderItemList(LV.SelectedItem.SubItems(4))
        Case "Print Contract Review"
            Call PrintContractReview(LV.SelectedItem.SubItems(4))
        Case "Order Status Log"
            Call PrintOrderStatusLog
        Case "Order Status Log (New)"
            Call PrintOrderStatusLogNew
        Case "Planning Purchase Log Sheet"
            Call PrintPlanningPurchaseLog
        Case "Production Timeline Report"
            Call PrintProductionTimelineReport
        Case "Shipping Dates"
            Set f = New frmFCustomerOrders_ShippingDates
            Load f
            Call f.ShowMe(LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(4))
        Case "Location-wise Stock Report"
            Call PrintStockOutReport(LV.SelectedItem.SubItems(4))
        Case "Update Rate for Invoices"
            Call UpdateRatesInProforma(LV.SelectedItem.SubItems(4))
        Case "Production Lots"
            Set f = New frmMakerRcvList
            Load f
            f.ShowForOrder (LV.SelectedItem.SubItems(4))
        Case "Urgent Items"
            Call PrintUrgetItemsRequestForm
        Case "Semi Finish Stock"
            Call PrintOrderSemiFinishStock
        Case "Finish Stock"
            Call PrintOrderFinishStock
        Case "Urgent Request"
            Call PrintUrgentRequest
        Case "Score Board"
            Call PrintScoreBoardReport
        Case "Orders Log Sheet"
            Call PrintOrdersLogSheet
        Case "Order Balance Report"
            Call PrintOrderBalanceReport_CT
        Case "Hub Wise Status"
            Call PrintHubWiseOrderStatusReport(LV.SelectedItem.SubItems(4))
        Case "Unplanned Items"
            Call PrintUnPlannedItems(LV.SelectedItem.SubItems(4))
    End Select
    'Me.PopupMenu mnuPop
    'If LV.ListItems.Count = 0 Or Button <> 2 Or OutOfFinancialYear Then Exit Sub

    'if a Voucher Entry Show Delete Menu
    'If InStr(LV.SelectedItem, "-") > 0 Then Me.PopupMenu mnuPop
    
End Sub


Private Sub PrintUrgentRequest()

    Dim strCondition As String
    strCondition = "{VrptOrders.DT}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#"
    strCondition = strCondition & " AND {VrptOrders.DeliveryStatus}=2"
    If cmbCust.ID <> "0" Then
        strCondition = strCondition & " AND {VrptOrders.CustCode}='" & cmbCust.ID & "'"
    End If
    If cmbOrderType.ListIndex > 0 Then
        strCondition = strCondition & " AND {VrptOrders.OrderType}='" & cmbOrderType.Text & "'"
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\UrgentRequest.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport strCondition, rpt
    
End Sub

Private Sub UpdateRatesInProforma(strOrderNo As String)

    If MsgBox("This will update Prices in Proforma & Future Invoices will be created with updated Prices" & vbNewLine & "Are you sure?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
    con.Execute "UPDATE FProformaOrders SET FProformaOrders.Price=FOrderItems.Price," & _
     "FProformaOrders.CustomPrice=FOrderItems.CustomPrice FROM FProformaOrders INNER JOIN FOrderItems ON " & _
     "FProformaOrders.OrderEntryID=FOrderItems.ID WHERE FOrderItems.OrderNo='" & strOrderNo & "'"
    
End Sub

Private Sub PrintOrderStatusLog()
    'PrintOrderStatusLogNew
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\OrderStatusLog.rpt")
   
    
    If cmbCust.ID <> "0" Then
        rpt.ParameterFields.GetItemByName("@CustCode").SetCurrentValue cmbCust.ID
    Else
        rpt.ParameterFields.GetItemByName("@CustCode").SetCurrentValue ""
    End If
    If cmbCountry.ListIndex > 0 Then
        rpt.ParameterFields.GetItemByName("@Country").SetCurrentValue cmbCountry.Text
    Else
        rpt.ParameterFields.GetItemByName("@Country").SetCurrentValue ""
    End If
        
    rpt.ParameterFields.GetItemByName("@ShipStatus").SetCurrentValue cmbStatus.ListIndex
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DtFrom.Value
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DtTo.Value
      
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, , False
    
End Sub

Private Sub PrintOrderStatusLogNew()
    'PrintOrderStatusLogNew
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\OrderStatusLog_Ex.rpt")
   
    If cmbCust.ID <> "0" Then
        rpt.ParameterFields.GetItemByName("@CustCode").SetCurrentValue cmbCust.ID
    Else
        rpt.ParameterFields.GetItemByName("@CustCode").SetCurrentValue ""
    End If
    If cmbCountry.ListIndex > 0 Then
        rpt.ParameterFields.GetItemByName("@Country").SetCurrentValue cmbCountry.Text
    Else
        rpt.ParameterFields.GetItemByName("@Country").SetCurrentValue ""
    End If
    rpt.ParameterFields.GetItemByName("@DeliveryDTFilter").SetCurrentValue CBool(chkFilterByDeliveryDT.Value)
    rpt.ParameterFields.GetItemByName("@ShipStatus").SetCurrentValue cmbStatus.ListIndex
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DtFrom.Value
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DtTo.Value
      
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, , False
    
End Sub

Private Sub PrintPlanningPurchaseLog()
    'PrintProductionTimelineReport
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\OrderPurchasePlanningLog.rpt")
    
    If cmbCust.ID <> "0" Then
        rpt.ParameterFields.GetItemByName("@CustCode").SetCurrentValue cmbCust.ID
    Else
        rpt.ParameterFields.GetItemByName("@CustCode").SetCurrentValue ""
    End If
    
    If cmbCountry.ListIndex > 0 Then
        rpt.ParameterFields.GetItemByName("@Country").SetCurrentValue cmbCountry.Text
    Else
        rpt.ParameterFields.GetItemByName("@Country").SetCurrentValue ""
    End If
        
    rpt.ParameterFields.GetItemByName("@ShipStatus").SetCurrentValue cmbStatus.ListIndex
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DtFrom.Value
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DtTo.Value
      
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, , False
    
End Sub

Private Sub PrintCombinedOrderReport(Optional bGroupwise As Boolean)
    
    If bGroupwise Then
    
        Dim strGroupID As String
        Dim frm As New frmGetSingleComboValue
        
        frm.cmbCombo.AddItem "<All Groups>", "0"
        frm.cmbCombo.AddVals con, "Description", "ItemGroups", "ID"
        
        Load frm
        If frm.getComboValue(strGroupID, "Group", "Group") = False Then Exit Sub
        
    End If
    
    Dim strSelection As String
    strSelection = "{FCustomerOrders.DT}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#"
    
'    If cmbStatus.ListIndex <> 2 Then
'        If cmbStatus.ListIndex = 0 Then
'            strSelection = strSelection & " AND {VFOrderItemswithShippedQty.Qty}>{VFOrderItemswithShippedQty.ShippedQty}"
'        ElseIf cmbStatus.ListIndex = 1 Then
'            strSelection = strSelection & " AND {VFOrderItemswithShippedQty.Qty}<{VFOrderItemswithShippedQty.ShippedQty}"
'        End If
'    End If
    
    If cmbCust.ID <> "0" Then
        strSelection = strSelection & " AND {FCustomerOrders.CustCode}='" & cmbCust.ID & "'"
    End If
    
    If cmbCountry.ListIndex <> 0 Then
        strSelection = strSelection & " AND {FCustomerOrders.Country}='" & cmbCountry.Text & "'"
    End If
    
    If Not (strGroupID = "0" Or strGroupID = "") Then
        strSelection = strSelection & " AND {Items.GroupID}=" & strGroupID
    End If
    
    Dim rpt As CRAXDDRT.Report
    
    Set rpt = rptApp.OpenReport(RptPath & "\rptPartiallyShippedOrdersItemwise.rpt")
    
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt

End Sub

Private Sub PrintCombinedOrderReportNew(Optional bGroupwise As Boolean)
    
    If bGroupwise Then
    
        Dim strGroupID As String
        Dim frm As New frmGetSingleComboValue
        
        frm.cmbCombo.AddItem "<All Groups>", "0"
        frm.cmbCombo.AddVals con, "Description", "ItemGroups", "ID"
        
        Load frm
        If frm.getComboValue(strGroupID, "Group", "Group") = False Then Exit Sub
        
    End If
    
    Dim strSelection As String
    strSelection = "{VCombinedOrderReportItemwise.DT}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#"
    
'    If cmbStatus.ListIndex <> 2 Then
'        If cmbStatus.ListIndex = 0 Then
'            strSelection = strSelection & " AND {VFOrderItemswithShippedQty.Qty}>{VFOrderItemswithShippedQty.ShippedQty}"
'        ElseIf cmbStatus.ListIndex = 1 Then
'            strSelection = strSelection & " AND {VFOrderItemswithShippedQty.Qty}<{VFOrderItemswithShippedQty.ShippedQty}"
'        End If
'    End If
    
    If cmbCust.ID <> "0" Then
        strSelection = strSelection & " AND {VCombinedOrderReportItemwise.CustCode}='" & cmbCust.ID & "'"
    End If
    
    If cmbCountry.ListIndex <> 0 Then
        strSelection = strSelection & " AND {VCombinedOrderReportItemwise.Country}='" & cmbCountry.Text & "'"
    End If
    
    If Not (strGroupID = "0" Or strGroupID = "") Then
        strSelection = strSelection & " AND {VCombinedOrderReportItemwise.GroupID}=" & strGroupID
    End If
    
    Dim rpt As CRAXDDRT.Report
    
    Set rpt = rptApp.OpenReport(RptPath & "\CombinedOrderReportItemwise.rpt")
    
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt

End Sub

Private Sub PrintCombinedOrderReportAll()

    Dim strSelection As String
    strSelection = "{FCustomerOrders.DT}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#"
    
    If cmbStatus.ListIndex <> 2 Then
        If cmbStatus.ListIndex = 0 Then
            strSelection = strSelection & " AND {VFOrderItemswithShippedQty.Qty}>{VFOrderItemswithShippedQty.ShippedQty}"
        ElseIf cmbStatus.ListIndex = 1 Then
            strSelection = strSelection & " AND {VFOrderItemswithShippedQty.Qty}<{VFOrderItemswithShippedQty.ShippedQty}"
        End If
    End If
    
    If cmbCust.ID <> "0" Then
        strSelection = strSelection & " AND {FCustomerOrders.CustCode}='" & cmbCust.ID & "'"
    End If
    
    If cmbCountry.ListIndex <> 0 Then
        strSelection = strSelection & " AND {FCustomerOrders.Country}='" & cmbCountry.Text & "'"
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptPartiallyShippedOrders.rpt")
    
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt

End Sub

Private Sub PrintProductionExportCost()
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\OrderwiseProductionExportCost.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VProductionExportCost.OrderNo}='" & LV.SelectedItem.SubItems(4) & "'", rpt
End Sub

Private Sub CopyCustomerOrder()

    If LV.ListItems.count = 0 Then Exit Sub
'    Dim f As New frmVendor_Order_AE
'    Load f
'    f.CopyCustomerOrder (LV.SelectedItem.SubItems(4))
    
End Sub

Private Sub PrintOrderList()

    On Error GoTo err
    Me.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptFOrderList.rpt")
        
    
    rpt.FormulaFields.GetItemByName("ForCustomer").Text = "'" & cmbCust.Text & " " & cmbCountry.Text & "'"
    rpt.FormulaFields.GetItemByName("ForFromTo").Text = "'" & Format(DtFrom, "dd-MMM-yyyy") & " to " & Format(DtTo.Value, "dd-MMM-yyyy") & "'"
    rpt.FormulaFields.GetItemByName("OrderType").Text = "'" & cmbOrderType.Text & "'"
    
    frmPrevRpt.ShowReport " ", rpt, SQLString & " ORDER BY CustCode"
    
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
End Sub
Private Sub PrintPTC()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Dim f As New frmPrevRpt, strSelection As String
    If strCompany = "QEL" Then
        Set rpt = rptApp.OpenReport(RptPath & "\PTCQEL.rpt")
        strSelection = "{VFOrderItemPTC.OrderNo}='" & LV.SelectedItem.SubItems(4) & "'"
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\PTC.rpt")
        strSelection = "{FCustomerOrders.OrderNo}='" & LV.SelectedItem.SubItems(4) & "'"
    End If
    f.ShowReport strSelection, rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Public Sub PrintOrderBalanceReportOld()


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

Private Sub mnuEdit_Click()
    Call LV_DblClick
End Sub

Private Sub PrintOrder(Optional bPrintBalance As Boolean = False, Optional bMakerWise As Boolean = False, Optional bStockSummary As Boolean = False, Optional bGroupwise As Boolean = False, Optional bMarkingPlan As Boolean = False, Optional bWithValue As Boolean = False, Optional bProductionOrder As Boolean = False, Optional bOrderStatus As Boolean = False)

    On Error GoTo err
    
    If strCompany = "IAA" Or strCompany = "Kami" Then
        If bMakerWise = False Then
             Call PrintOrderWrangler(bPrintBalance, bMakerWise, bStockSummary, bGroupwise, bMarkingPlan, bWithValue, bProductionOrder, bOrderStatus)
             Exit Sub
        End If
    End If
    
    Dim strSelection As String
    Dim rpt As CRAXDDRT.Report
    
    If Not bPrintBalance Then
        If bMakerWise Then
            Set rpt = rptApp.OpenReport(RptPath & "\rptProductionOrderMakerwise.rpt") '
        Else
            If strCompany = "IAA" Or strCompany = "Kami" Then
                If bGroupwise Then
                    Dim strGroupIDs As String
                    Dim f As New frmGetSingleComboValue
                    Load f
                    f.cmbCombo.AddItem "<All Groups>", "0"
                    f.cmbCombo.Functionality = 1
                    f.cmbCombo.AddVals con, "Description", "ItemGroups", "ID"
                    f.cmbCombo.ID = "0"
                    If f.getComboValue(strGroupIDs, "Group", "Group") = False Then Exit Sub
                    Set rpt = rptApp.OpenReport(RptPath & "\rptProductionOrderWranglerGroupWise.rpt") '
                Else
                    If bProductionOrder Then
                        Set rpt = rptApp.OpenReport(RptPath & "\rptProductionOrderWrangler_1.rpt") '
                    Else
                        Set rpt = rptApp.OpenReport(RptPath & "\rptProductionOrderWrangler.rpt") '
                    End If
                End If
            Else
                Set rpt = rptApp.OpenReport(RptPath & "\rptProductionOrder.rpt") '
            End If
        End If
        rpt.FormulaFields.GetItemByName("ProductionOrder").Text = bForProduction
        strSelection = "{CustOrders.OrderNo}='" & Left(LV.SelectedItem.key & "'", Len(LV.SelectedItem.key) - 1) & "' "
        If bStockSummary Then
            strSelection = strSelection & " AND ({CustOrders.InHand}+{CustOrders.AdjQty})>0"
        End If
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\rptProductionOrderBalanceQty.rpt") '
        rpt.FormulaFields.GetItemByName("ProductionOrder").Text = bForProduction
        strSelection = "{CustOrders.OrderNo}='" & Left(LV.SelectedItem.key & "'", Len(LV.SelectedItem.key) - 1) & "' AND {@Balance}>0"
    End If
    
    rpt.FormulaFields.GetItemByName("UserName").Text = "'" & CurrentUserName & "'"
    Load frmPrevRpt
    
    If Not (strGroupIDs = "0" Or strGroupIDs = "") Then
        strSelection = strSelection & " AND {CustOrders.GroupID} IN[" & Replace(strGroupIDs, "'", "") & "]"
    End If
    
    frmPrevRpt.ShowReport strSelection, rpt
   
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub PrintOrderWrangler(Optional bPrintBalance As Boolean = False, Optional bMakerWise As Boolean = False, Optional bStockSummary As Boolean = False, Optional bGroupwise As Boolean = False, Optional bMarkingPlan As Boolean = False, Optional bWithValue As Boolean = False, Optional bProductionOrder As Boolean = False, Optional bOrderStatus As Boolean = False)

    On Error GoTo err
    
   
    Dim strSelection As String
    Dim rpt As CRAXDDRT.Report
    
    If bGroupwise Then
        Dim strGroupIDs As String
        Dim f As New frmGetSingleComboValue
        Load f
        f.cmbCombo.AddItem "<All Groups>", "0"
        f.cmbCombo.Functionality = 1
        f.cmbCombo.AddVals con, "Description", "ItemGroups", "ID"
        f.cmbCombo.ID = "0"
        If f.getComboValue(strGroupIDs, "Group", "Group") = False Then Exit Sub
        Set rpt = rptApp.OpenReport(RptPath & "\rptProductionOrderWranglerGroupWise.rpt") '
    Else
        If bMarkingPlan Then
            Set rpt = rptApp.OpenReport(RptPath & "\rptProductionOrderWranglerMarkingPlan.rpt") '
        ElseIf bWithValue Then
            Set rpt = rptApp.OpenReport(RptPath & "\rptProductionOrderWValueWrangler.rpt") '
        ElseIf bProductionOrder Then
            Set rpt = rptApp.OpenReport(RptPath & "\rptProductionOrderWrangler_1.rpt")
        ElseIf bOrderStatus Then
            Set rpt = rptApp.OpenReport(RptPath & "\rptProductionOrder_Status.rpt")
        Else
            Set rpt = rptApp.OpenReport(RptPath & "\rptProductionOrderWrangler.rpt") '
        End If
    End If
       
    rpt.FormulaFields.GetItemByName("ProductionOrder").Text = bForProduction
    strSelection = "{CustOrders.OrderNo}='" & Left(LV.SelectedItem.key & "'", Len(LV.SelectedItem.key) - 1) & "' "
    If bStockSummary Then
        strSelection = strSelection & " AND ({CustOrders.InHand}+{CustOrders.AdjQty})>0"
    End If
        
    rpt.ParameterFields.GetItemByName("@OrderNo").SetCurrentValue LV.SelectedItem.SubItems(4)
     
    rpt.FormulaFields.GetItemByName("UserName").Text = "'" & CurrentUserName & "'"
    Load frmPrevRpt
    
    If Not (strGroupIDs = "0" Or strGroupIDs = "") Then
        strSelection = strSelection & " AND {CustOrders.GroupID} IN[" & Replace(strGroupIDs, "'", "") & "]"
    End If
    If bPrintBalance Then
        strSelection = strSelection & " AND {CustOrders.Qty}>{CustOrders.ShippedQty}"
    End If
    rpt.EnableParameterPrompting = False
    frmPrevRpt.ShowReport strSelection, rpt, , False
   
    
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

Private Sub PrintCustomerOrder()
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\CustomerPO.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{FCustomerOrders.OrderNo}='" & LV.SelectedItem.SubItems(4) & "'", rpt
End Sub

Private Sub txtFind_Change()
    cmdFind.Enabled = Len(Trim(txtFind)) > 0
End Sub

Private Sub txtFind_GotFocus()
   txtFind.SelStart = 0
   txtFind.SelLength = Len(txtFind)
   cmdFind.Default = True
End Sub

Private Sub txtFind_LostFocus()
   cmdFind.Default = False
End Sub

Private Sub PrintOrderItemList(strOrderNo As String)
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\OrderItemList.rpt")
    
    Dim strFilters As String
     
    strFilters = "Order No.:" & strOrderNo
   
    rpt.FormulaFields.GetItemByName("Filters").Text = "'" & strFilters & "'"
     
    Dim strReportSQL As String
    strReportSQL = "SELECT * FROM VFOrderItems INNER JOIN VFOrderItemsWithShippedQty ON VFOrderItems.ID=VFOrderItemsWithShippedQty.ID LEFT OUTER JOIN FOrderItems_DeliveryRevisions ON VFOrderItems.ID=FOrderItems_DeliveryRevisions.FOI_RefID WHERE VFOrderItems.OrderNo='" & strOrderNo & "' AND VFOrderItemsWithShippedQty.Qty>VFOrderItemsWithShippedQty.ShippedQty ORDER BY VFOrderItems.DeliveryDT DESC"
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, strReportSQL

End Sub


Private Sub PrintProductionTimelineReport()
    '
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ProductionTimeLineReport.rpt")
    
    If cmbCust.ID <> "0" Then
        rpt.ParameterFields.GetItemByName("@CustCode").SetCurrentValue cmbCust.ID
    Else
        rpt.ParameterFields.GetItemByName("@CustCode").SetCurrentValue ""
    End If
    
    If cmbCountry.ListIndex > 0 Then
        rpt.ParameterFields.GetItemByName("@Country").SetCurrentValue cmbCountry.Text
    Else
        rpt.ParameterFields.GetItemByName("@Country").SetCurrentValue ""
    End If
        
    rpt.ParameterFields.GetItemByName("@ShipStatus").SetCurrentValue cmbStatus.ListIndex
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DtFrom.Value
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DtTo.Value
      
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, , False
    
End Sub

Private Sub PrintUrgetItemsRequestForm()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\UrgetItemsRequestForm.rpt")
    
    Dim strSelection As String
    strSelection = "{VFOrderItems.DT}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "# AND {VFOrderItems.DeliveryStatus}=2"
    
    If cmbCust.ID <> "0" Then
        strSelection = strSelection & " AND {VFOrderItems.CustCode}='" & cmbCust.ID & "'"
    End If
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub PrintOrderSemiFinishStock()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Order_SemiFinishStock.rpt")
    
    Dim strSelection As String
    strSelection = "{FOrderItems.OrderNo}='" & LV.SelectedItem.SubItems(4) & "' AND {VItems_StockReport.SFStock}>0"
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub PrintOrderFinishStock()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Order_FinishStock.rpt")
    
    Dim strSelection As String
    strSelection = "{FOrderItems.OrderNo}='" & LV.SelectedItem.SubItems(4) & "' AND {VItems_StockReport.InHand}>0"
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub PrintScoreBoardReport()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ScoreBoard.rpt")
    
    rpt.ParameterFields.GetItemByName("@OrderNo").SetCurrentValue LV.SelectedItem.SubItems(4)
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, , False
    
End Sub

Private Sub PrintOrdersLogSheet()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Orders_Log_Sheet.rpt")
   
    If cmbCust.ID <> "0" Then
        rpt.ParameterFields.GetItemByName("@CustCode").SetCurrentValue cmbCust.ID
    Else
        rpt.ParameterFields.GetItemByName("@CustCode").SetCurrentValue ""
    End If
    If cmbCountry.ListIndex > 0 Then
        rpt.ParameterFields.GetItemByName("@Country").SetCurrentValue cmbCountry.Text
    Else
        rpt.ParameterFields.GetItemByName("@Country").SetCurrentValue ""
    End If
    rpt.ParameterFields.GetItemByName("@DeliveryDTFilter").SetCurrentValue CBool(chkFilterByDeliveryDT.Value)
    rpt.ParameterFields.GetItemByName("@ShipStatus").SetCurrentValue cmbStatus.ListIndex
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DtFrom.Value
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DtTo.Value
      
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, , False
    
End Sub

Private Function DeleteOrder(p_strOrderNo As String) As Boolean

'    If getDBPassword("Delete_Customer_Order") = False Then Exit Function
'    Dim lCount As Long
'    lCount = GetSingleLongValue("COUNT(*)", "FProformaOrders", " WHERE OrderEntryID IN(SELECT ID FROM FOrderItems WHERE OrderNo='" & p_strOrderNo & "'")
'    If lCount > 0 Then
'        MsgBox "Can't Delete,Proforma Exist.", vbInformation
'        Exit Function
'    End If
    
    con.Execute "DELETE FROM Vendor_Orders WHERE OrderNo='" & p_strOrderNo & "'"
    DeleteOrder = True

End Function

Private Sub CopyOrder(p_strOrderNo As String)
    Dim f As New frmFNewCustOrder_Frgz
    Load f
    Call f.EditOrder(p_strOrderNo, True)
End Sub

Private Sub PrintOrderBalanceReport_CT()
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Balance_Orders_CT.rpt")
    
    Dim strSelection As String
    strSelection = "{VFOrderItemswithShippedQty.DT}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#"
    If cmbCust.ID <> "0" Then
        strSelection = strSelection & " AND {VFOrderItemswithShippedQty.CustCode}='" & cmbCust.ID & "'"
    End If
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
End Sub
