VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmVendor_Catalog 
   ClientHeight    =   6570
   ClientLeft      =   1050
   ClientTop       =   0
   ClientWidth     =   12735
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   12735
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkShowTop 
      Caption         =   "Top Quality"
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
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   6210
      Visible         =   0   'False
      Width           =   1560
   End
   Begin ComboList.Usercontrol1 cmbItem 
      Height          =   285
      Left            =   3870
      TabIndex        =   1
      Top             =   855
      Width           =   6045
      _ExtentX        =   10663
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
   Begin ComboList.Usercontrol1 cmbCust 
      Height          =   285
      Left            =   75
      TabIndex        =   0
      Top             =   855
      Width           =   3780
      _ExtentX        =   6668
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
   Begin VB.TextBox txtMaxDiscount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   75
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   1365
      Width           =   1350
   End
   Begin VB.PictureBox PicProduct 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2865
      Left            =   7245
      ScaleHeight     =   2835
      ScaleWidth      =   2760
      TabIndex        =   19
      Top             =   1950
      Visible         =   0   'False
      Width           =   2790
      Begin VB.Image ImgProduct 
         Height          =   2835
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2760
      End
   End
   Begin VB.CheckBox chkShowPicture 
      Caption         =   "Show Picture"
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
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   5985
      Visible         =   0   'False
      Width           =   1560
   End
   Begin ComboList.Usercontrol1 cmbItemSearch 
      Height          =   285
      Left            =   1425
      TabIndex        =   3
      Top             =   1365
      Width           =   8490
      _ExtentX        =   14975
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
   Begin VB.Frame fupdate 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   390
      Left            =   1725
      TabIndex        =   15
      Top             =   6090
      Visible         =   0   'False
      Width           =   2775
      Begin MSForms.CommandButton cmdCheckAll 
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   1365
         ForeColor       =   0
         Caption         =   " Assign All"
         PicturePosition =   327683
         Size            =   "2408;661"
         Accelerator     =   104
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdClearAll 
         Height          =   375
         Left            =   1395
         TabIndex        =   8
         Top             =   0
         Width           =   1365
         ForeColor       =   0
         Caption         =   "Unassign All"
         PicturePosition =   327683
         Size            =   "2408;661"
         Accelerator     =   108
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin VB.CommandButton cmdUnAssign 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11100
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "UnAssign Selected Item..."
      Top             =   1230
      UseMaskColor    =   -1  'True
      Width           =   945
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   570
      Top             =   0
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
            Picture         =   "frmVendor_Catalog.frx":0000
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVendor_Catalog.frx":01BC
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVendor_Catalog.frx":0378
            Key             =   "Active"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVendor_Catalog.frx":07CC
            Key             =   "Inactive"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fTitle 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10365
      Begin ComboList.Usercontrol1 cmbCountry1 
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   1530
         _ExtentX        =   2699
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Catalog"
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
         Index           =   0
         Left            =   3840
         TabIndex        =   10
         Top             =   120
         Width           =   2445
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Catalog"
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
         Height          =   390
         Index           =   1
         Left            =   3855
         TabIndex        =   11
         Top             =   135
         Width           =   2445
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4005
      Left            =   120
      TabIndex        =   12
      Top             =   1950
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   7064
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
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
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Vend. Ref. No."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Item No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Unit"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "FOB KHI"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "FOB SKT"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "C&I Sea"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "C&I Air"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "CIF (Sea)"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "CIF (Air)"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "C&F Sea"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "C&F Air"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "PU"
         Object.Width           =   1499
      EndProperty
   End
   Begin MSForms.CommandButton cmdPrintCatalog 
      Height          =   375
      Left            =   4080
      TabIndex        =   23
      Top             =   6090
      Visible         =   0   'False
      Width           =   1425
      ForeColor       =   0
      Caption         =   "Print Catalog"
      PicturePosition =   327683
      Size            =   "2514;661"
      Accelerator     =   103
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdPrintPriceListTop 
      Height          =   375
      Left            =   5550
      TabIndex        =   22
      Top             =   6090
      Visible         =   0   'False
      Width           =   1950
      ForeColor       =   0
      Caption         =   "Print Price List (Top)"
      PicturePosition =   327683
      Size            =   "3440;661"
      Accelerator     =   84
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   375
      Left            =   9945
      TabIndex        =   4
      Top             =   1230
      Width           =   1125
      ForeColor       =   0
      Caption         =   "Refresh"
      PicturePosition =   327683
      Size            =   "1984;661"
      Accelerator     =   82
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdPrintPriceList 
      Height          =   375
      Left            =   7545
      TabIndex        =   5
      Top             =   6090
      Visible         =   0   'False
      Width           =   1470
      ForeColor       =   0
      Caption         =   "Print Price List"
      PicturePosition =   327683
      Size            =   "2593;661"
      Accelerator     =   80
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   1
      Left            =   75
      TabIndex        =   17
      Top             =   1140
      Width           =   9840
      BackColor       =   11517387
      Caption         =   " Max. Discount   Search Item No."
      Size            =   "17357;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdAssign 
      Height          =   375
      Left            =   9945
      TabIndex        =   2
      Top             =   765
      Width           =   1125
      ForeColor       =   0
      Caption         =   "Assign   "
      PicturePosition =   327683
      Size            =   "1984;661"
      Picture         =   "frmVendor_Catalog.frx":0C20
      Accelerator     =   65
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   0
      Left            =   75
      TabIndex        =   13
      Top             =   630
      Width           =   9840
      BackColor       =   11517387
      Caption         =   "  Vendor                                                                     Item No."
      Size            =   "17357;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   375
      Left            =   9045
      TabIndex        =   6
      Top             =   6090
      Width           =   990
      ForeColor       =   0
      Caption         =   "Close          "
      PicturePosition =   327683
      Size            =   "1746;661"
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
Attribute VB_Name = "frmVendor_Catalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbCountry_matched()
    cmbItem.ListHeight = 2600
    cmbItem.ClearVals
    cmbItem.AddVals con, "ItemID", "Items", "ItemID", " Where ItemID Not In(Select CompItemID From FCustomerCatalog Where CustCode='" & cmbCust.ID & "' and Country='" & cmbCountry & "')"
    
    'cmdUnAssign.Visible = False
    Call RefreshLV

End Sub
Private Sub cmbCountry_UnMatched()
    cmbItem.ClearVals
    LV.ListItems.Clear
End Sub

Private Sub chkShowPicture_Click()
    If chkShowPicture.Value = vbChecked Then
        Call LV_ItemClick(LV.SelectedItem)
    Else
        PicProduct.Visible = False
    End If
End Sub

Private Sub chkShowTop_Click()
    Call RefreshLV
End Sub

Private Sub cmbCountry_Change()
    'Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()
    
    
End Sub

Private Sub cmbCust_matched()

    Dim f As New frmWaitAnimation
    f.Show
    DoEvents
    
    cmbItem.ListHeight = 2600
    cmbItem.ClearVals
    cmbItem.DropDownOffLine = True
    cmbItem.AddVals con, "'{' + ItemID + '} ' + ItemName", "Items", "ItemID", " WHERE ItemID NOT IN(SELECT ItemID FROM Vendor_Catalog WHERE Vend_AccNo='" & cmbCust.ID & "' AND InActive=0)"
    
    
    cmbItemSearch.ClearVals
    cmbItemSearch.DropDownOffLine = True
    cmbItemSearch.AddVals con, "ItemID_Vendor + ' ' + Description + '{ ' + ItemID + ' }'", "Vendor_Catalog", "ItemID", "WHERE Vend_AccNo='" & cmbCust.ID & "'"
    
    Call RefreshLV
    
    'txtMaxDiscount = GetSingleDoubleValue("MaxDiscount", "ForeignCustomers", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'")
    
    Unload f
End Sub

Private Sub cmbCust_UnMatched()
    LV.ListItems.Clear
End Sub

Private Sub cmbItemName_matched()
    cmbItem.ID = cmbItemName.ID
End Sub

Private Sub cmdAssign_Click()
    
    
    If cmbCust.MatchFound = False Then Exit Sub
    If cmbItem.MatchFound = False Then Exit Sub
    
    Dim Saved As Integer
    Dim cmd As New ADODB.Command
    
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "Vendor_AssItems_SP"
        .Parameters("@Vend_AccNo").Value = cmbCust.ID
        .Parameters("@ItemID").Value = cmbItem.ID
        .Execute Saved
    End With
    
    Dim strItemID As String
    strItemID = cmbItem.ID
    
    If Saved = 1 Then
        Call RefreshLV
        'Call CmbCust's Matched Event To Remove Current Item
        'From The cmbItem...
        
        'Call cmbCountry_Click
        
        LV.ListItems(cmbCust.ID & strItemID).Selected = True
        LV.ListItems(cmbCust.ID & strItemID).EnsureVisible
        
        cmbItem.ID = ""
        cmbItem.SetFocus
       
    End If
    
End Sub



Private Sub cmdCheckAll_Click()

    If cmbCust.MatchFound = False Then Exit Sub
    If cmbCountry.MatchFound = False Then Exit Sub
    'If cmbItem.MatchFound = False Then Exit Sub
    
    If MsgBox("Are you sure, You want to Assign all ?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    Dim Saved As Integer
    Dim cmd As New ADODB.Command
    
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "Vendor_AssItems_SP"
        .Parameters("@Vend_AccNo").Value = cmbCust.ID
        .Parameters("@ItemID").Value = "%" 'Assign All...
        .Execute Saved
    End With
    
    
    If Saved > 0 Then
        Call RefreshLV
        'Call CmbCust's Matched Event To Remove Current Item
        'From The cmbItem...
        Call cmbCountry_matched
        'cmdAssign.SetFocus
    End If

End Sub

Private Sub cmdClearAll_Click()
    On Error GoTo err
    Dim iAffect As Integer
    If MsgBox("Are You Sure To UnAssign All Items?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    con.Execute "DELETE FROM FCustomerCatalog WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "' AND CompItemID NOT IN(SELECT CompItemCode FROM FOrderItems INNER JOIN FCustomerOrders ON FCustomerOrders.OrderNo=FOrderItems.OrderNo WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "')", iAffect
    If iAffect > 0 Then
        Call RefreshLV
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrintCatalog_Click()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\CustomerCatalog.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{FCustomerCatalog.CustCode}='" & cmbCust.ID & "'", rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdPrintPriceList_Click()
    Call PrintPriceList
End Sub

Private Sub PrintPriceList(Optional bTop As Boolean = False)

    On Error GoTo err
    
    Dim myDT As Date
    myDT = InputBox("Please Enter Validity Date (dd/MM/yyyy)", "Validity Date", Format(DateAdd("M", 1, Date), "dd/MM/yyyy"))
    
    If cmbCust.MatchFound = False Then Exit Sub
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptCustomerPriceList.rpt")
    
    rpt.FormulaFields.GetItemByName("TopPrices").Text = bTop
    Dim f As New frmPrevRpt
    rpt.FormulaFields.GetItemByName("ValidityDT").Text = "#" & myDT & "#"
    f.ShowReport "{ForeignCustomers.CustCode}='" & cmbCust.ID & "'", rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
    
Private Sub cmdPrintPriceListTop_Click()
    PrintPriceList True
End Sub

Private Sub cmdRefresh_Click()
    Call cmbCountry_Click
End Sub

Private Sub cmdUnAssign_Click()

    On Error GoTo err
    If LV.ListItems.count = 0 Then Exit Sub
    
'    If ItemIsInOrder(LV.SelectedItem.SubItems(1), cmbCust.ID, cmbCountry.Text) <> "" Then
'        MsgBox "Item is in use in Order.", vbInformation
'        Exit Sub
'    End If
    
    Dim iAffect As Integer
    
    con.Execute "DELETE FROM Vendor_Catalog WHERE Vend_AccNo+ItemID='" & LV.SelectedItem.key & "'", iAffect
    If iAffect > 0 Then
        Call RefreshLV
        'cmdUnAssign.Visible = False
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Public Function ItemIsInOrder(strItemCode, Optional strCustomer As String = "", Optional strCountry As String = "") As String

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim strOrderNo As String
    With rs
        If strCustomer = "" Then
            .Open "SELECT TOP 1 OrderNo FROM FOrderItems WHERE CompItemCode='" & strItemCode & "'", con, adOpenForwardOnly, adLockReadOnly
        Else
            .Open "SELECT TOP 1 OrderNo FROM VrptOrders WHERE CustCode='" & strCustomer & "' AND Country='" & strCountry & "' AND CompItemID='" & strItemCode & "'", con, adOpenForwardOnly, adLockReadOnly
        End If
        If .EOF = False Then
            strOrderNo = .Fields(0) & ""
        Else
            strOrderNo = ""
        End If
        .Close
    End With
    Set rs = Nothing
    ItemIsInOrder = strOrderNo
    
    Exit Function
err:
    MsgBox err.Description
End Function

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Exit Sub
    LV.Sorted = True
    If LV.SortKey = ColumnHeader.Index - 1 Then
        If LV.SortOrder = lvwAscending Then
            LV.SortOrder = lvwDescending
            ColumnHeader.Icon = "Up"
        Else
            LV.SortOrder = lvwAscending
            ColumnHeader.Icon = "Down"
        End If
    Else
        LV.ColumnHeaderIcons = Nothing
        Set LV.ColumnHeaderIcons = ImageList2
        LV.SortKey = ColumnHeader.Index - 1
        LV.SortOrder = lvwAscending
        ColumnHeader.Icon = "Down"
    End If

    On Error Resume Next
    LV.SelectedItem.EnsureVisible

End Sub

Private Sub LV_DblClick()
    
    If LV.ListItems.count = 0 Then Exit Sub
    
    
'    If strCompany = "IAA" OR strCompany = "Kami" Then
'        If getDBPassword("EditExportDocs") = False Then Exit Sub
'    End If

    Load frmEditItem_VC
    Call frmEditItem_VC.EditItem(LV.SelectedItem.key)
    'Call RefreshLV
    
End Sub


Private Sub Form_Load()

    cmbCust.ListHeight = 2600
    cmbCust.ClearVals
    cmbCust.AddVals con, "AccTitle + ' {'+AccNo+'}'", "VVenders", "AccNo", " WHERE Active=1"
    
    cmbItemSearch.ListHeight = 2600
    
End Sub

Public Sub RefreshLV()

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem, LColor As Long, i As Integer
    
    With rs
        .Open "SELECT Vend_AccNo,ItemID,ItemID_Vendor,Description,Unit,FOB,ExWorks,CnISea,CnIAir,CIFSea,CIFAir,CnFSea,CnFAir,PackingMode,ILO,OCR FROM Vendor_Catalog WHERE Vend_AccNo='" & cmbCust.ID & "' ORDER BY ItemID", con, adOpenForwardOnly, adLockReadOnly
         
        LV.ListItems.Clear
        Do Until .EOF
            
            Set ITM = LV.ListItems.add(, ![Vend_AccNo] & ![ITemID], ![ItemID_Vendor] & "")
            If LV.ListItems.count Mod 2 = 1 Then
                LColor = vbBlue
            Else
                LColor = vbBlack
            End If
            ITM.ForeColor = LColor
            ITM.ListSubItems.add , , !ITemID & ""
            ITM.ListSubItems.add , , ![Description] & ""
'            ITM.ListSubItems.add , , ![ILO] & ""
'            ITM.ListSubItems.add , , ![OCR] & ""
            ITM.ListSubItems.add , , ![Unit] & ""
            ITM.ListSubItems.add , , ![FOB] & ""
            ITM.ListSubItems.add , , ![ExWorks] & ""
            ITM.ListSubItems.add , , ![CnISea] & ""
            ITM.ListSubItems.add , , ![CnIAir] & ""
            ITM.ListSubItems.add , , ![CIFSea] & ""
            ITM.ListSubItems.add , , ![CIFAir] & ""
            ITM.ListSubItems.add , , ![CnFSea] & ""
            ITM.ListSubItems.add , , ![CnFAir] & ""
            
            ITM.ListSubItems.add , , ![PackingMode] & ""
            For i = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(i).ForeColor = LColor
            Next
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
End Sub
Private Sub Form_Resize()

    On Error Resume Next
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
     
    Dim TitleBottom As Long
    
    TitleBottom = cmbItemSearch.Top + cmbItemSearch.Height 'fTitle.Top + fTitle.Height
     
    LV.Move 50, TitleBottom, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + cmdClose.Height + 300)
     
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 50
    cmdPrintPriceList.Move cmdClose.Left - cmdPrintPriceList.Width - 50, cmdClose.Top
    cmdPrintPriceListTop.Move cmdPrintPriceList.Left - cmdPrintPriceListTop.Width - 50, cmdPrintPriceList.Top
    
    cmdPrintCatalog.Move cmdPrintPriceListTop.Left - cmdPrintCatalog.Width - 50, cmdPrintPriceListTop.Top
    
    chkShowTop.Move LV.Left, LV.Top + LV.Height + 50
    chkShowPicture.Move LV.Left, LV.Top + chkShowTop.Height + LV.Height + 50
    fupdate.Move chkShowTop.Left + chkShowTop.Width + 50, chkShowTop.Top
    'picproduct.Move lv.Left+lv.Width-picproduct.
    'cmdCheckAll.Move LV.Left, LV.Top + LV.Height + 50
 
 
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)

    
    If chkShowPicture.Value = vbChecked Then
        Dim rs As New ADODB.Recordset
        With rs
            .Open "SELECT ItemPic FROM Items WHERE ItemID='" & Item.SubItems(1) & "'", con, adOpenForwardOnly, adLockReadOnly
            If Not IsNull(![ItemPic]) Then
                Set ImgProduct.DataSource = rs
                ImgProduct.DataField = "ItemPic"
            Else
                ImgProduct.Picture = LoadPicture("")
            End If
            .Close
        End With
        Set rs = Nothing
        Dim lTop As Long
        lTop = LV.Top + LV.SelectedItem.Top
        If (lTop + PicProduct.Height) > LV.Top + LV.Height Then
            lTop = LV.Top + LV.Height - PicProduct.Height
        End If
        PicProduct.Move LV.Left + LV.Width - PicProduct.Width, lTop
        PicProduct.Visible = True
    Else
'        If PicProduct.Visible = True Then PicProduct.Visible = False
    End If
    
End Sub
Private Sub cmbItemSearch_matched()
    Call GoToSelected
End Sub

Private Sub cmbDescSearch_matched()
    cmbItemSearch.ID = cmbDescSearch.ID
End Sub

Private Sub GoToSelected()

    On Error GoTo err
    Dim ITemID As String
    ITemID = cmbCust.ID & cmbCountry.Text & cmbItemSearch.ID
    LV.ListItems(ITemID).EnsureVisible
    'LV.FindItem "     »  " & cmbitem.ID
    LV.ListItems(ITemID).Selected = True
    LV.SetFocus
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If LV.ListItems.count = 0 Then Exit Sub
    If Button <> 2 Then Exit Sub
    
    Dim strRet As String
    
    strRet = ShowPopUpMenu1("Increase Rate", "-", "Decrease Rate")
    
    Dim iOperation As Integer
    If strRet = "Increase Rate" Then
        iOperation = 0
    ElseIf strRet = "Decrease Rate" Then
        iOperation = 1
    Else
        Exit Sub
    End If
     
    Dim strValue As String, dValue As Double
    Dim f As New frmGetSingleTextValue
    If f.getTextValue(strValue, False, "Percentage", "Percentage") = False Then Exit Sub
    dValue = Val(strValue)
      
    If iOperation = 0 Then
        con.Execute "UPDATE FcustomerCatalog SET FOB=ROUND(FOB+((FOB/100)*" & dValue & "),2) WHERE CustCode='" & cmbCust.ID & "'"
    ElseIf iOperation = 1 Then
        con.Execute "UPDATE FcustomerCatalog SET FOB=ROUND(FOB-((FOB/100)*" & dValue & "),2) WHERE CustCode='" & cmbCust.ID & "'"
    End If
        
    Call RefreshLV
    
End Sub
