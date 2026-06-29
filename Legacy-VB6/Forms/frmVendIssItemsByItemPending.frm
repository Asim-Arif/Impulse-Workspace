VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmVendIssItemsByItemPending 
   ClientHeight    =   9810
   ClientLeft      =   1110
   ClientTop       =   60
   ClientWidth     =   11805
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin ComboList.Usercontrol1 cmbVender 
      Height          =   285
      Left            =   420
      TabIndex        =   19
      Top             =   5085
      Width           =   3435
      _ExtentX        =   6059
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
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3870
      TabIndex        =   21
      Top             =   5085
      Width           =   1110
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5745
      TabIndex        =   12
      Top             =   9120
      Width           =   5280
      Begin MSForms.CommandButton cmdSave 
         Height          =   360
         Left            =   75
         TabIndex        =   15
         Top             =   165
         Width           =   1665
         Caption         =   "Save & New   "
         PicturePosition =   327683
         Size            =   "2937;635"
         Accelerator     =   110
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdCancel 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   3510
         TabIndex        =   14
         Top             =   165
         Width           =   1665
         Caption         =   "Cancel            "
         PicturePosition =   327683
         Size            =   "2937;635"
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdSaveClose 
         Height          =   360
         Left            =   1785
         TabIndex        =   13
         Top             =   165
         Width           =   1665
         Caption         =   "Save & Close  "
         PicturePosition =   327683
         Size            =   "2937;635"
         Accelerator     =   83
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin ComboList.Usercontrol1 cmbProcess 
      Height          =   285
      Left            =   795
      TabIndex        =   10
      Top             =   810
      Width           =   2355
      _ExtentX        =   4154
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
      Enabled         =   0   'False
      Appearance      =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3630
      Left            =   405
      TabIndex        =   7
      Top             =   5430
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   6403
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Vender History"
      TabPicture(0)   =   "frmVendIssItemsByItemPending.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtRate"
      Tab(0).Control(1)=   "LVMaker"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Issued Items"
      TabPicture(1)   =   "frmVendIssItemsByItemPending.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "LVIssued"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtRate1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.TextBox txtRate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   16
         Top             =   1590
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -70050
         TabIndex        =   11
         Top             =   645
         Visible         =   0   'False
         Width           =   1350
      End
      Begin MSComctlLib.ListView LVMaker 
         Height          =   3150
         Left            =   -74910
         TabIndex        =   8
         Top             =   390
         Width           =   10770
         _ExtentX        =   18997
         _ExtentY        =   5556
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S No."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item Description"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Qty"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView LVIssued 
         Height          =   3180
         Left            =   75
         TabIndex        =   9
         Top             =   375
         Width           =   10800
         _ExtentX        =   19050
         _ExtentY        =   5609
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S No."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Vender"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Iss ID"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Order No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Item Code"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Description"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Rate"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Issued Qty"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   3165
      TabIndex        =   6
      Top             =   810
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   503
      _Version        =   393216
      CheckBox        =   -1  'True
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   48562179
      CurrentDate     =   37711
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "POST"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6495
      Picture         =   "frmVendIssItemsByItemPending.frx":0038
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4860
      UseMaskColor    =   -1  'True
      Width           =   600
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   30
      Top             =   360
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
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   555
      Top             =   -195
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
            Picture         =   "frmVendIssItemsByItemPending.frx":013A
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVendIssItemsByItemPending.frx":02F6
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -15
      Top             =   -195
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
            Picture         =   "frmVendIssItemsByItemPending.frx":04B2
            Key             =   "Active"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVendIssItemsByItemPending.frx":0906
            Key             =   "Inactive"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV 
      Height          =   3630
      Left            =   420
      TabIndex        =   5
      Top             =   1215
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   6403
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "S No."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Customer"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Order No"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Item Code"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Description"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Issueable Qty"
         Object.Width           =   2822
      EndProperty
   End
   Begin VB.Frame fTitle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   795
      TabIndex        =   0
      Top             =   0
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Items"
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
         Left            =   4440
         TabIndex        =   1
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Items"
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
         Left            =   4455
         TabIndex        =   2
         Top             =   135
         Width           =   1695
      End
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   285
      Left            =   4980
      TabIndex        =   22
      Top             =   5070
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   48562179
      CurrentDate     =   37711
   End
   Begin MSForms.Label Label3 
      Height          =   240
      Left            =   420
      TabIndex        =   20
      Top             =   4860
      Width           =   5970
      BackColor       =   11517387
      Caption         =   "  Venders                                                            Quantity         Date"
      Size            =   "10530;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cmbDesc 
      Height          =   285
      Left            =   5130
      TabIndex        =   18
      Top             =   810
      Width           =   3300
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5821;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbItems 
      Height          =   285
      Left            =   3165
      TabIndex        =   17
      Top             =   810
      Width           =   1935
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3413;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   795
      TabIndex        =   3
      Top             =   585
      Width           =   7635
      BackColor       =   11517387
      Caption         =   "  Process                                    Item Code                        Item Description"
      Size            =   "13467;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmVendIssItemsByItemPending"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyCol As New Collection
Dim MyCol1 As New Collection
Dim IsItemsFirst As Boolean
Dim IssIDs As New Collection
Dim rsMain As New ADODB.Recordset
Dim IsLotProcess As Boolean
Dim GotIssID As Long
Dim RcvingLot As Boolean
Dim UserReqAuth As Integer 'Whether Or Not User Required Authorization On Issue For This Process.

Private Sub AddItems()
    
    On Error GoTo err
    Exit Sub
    Dim OrCond As String
    'Create In List
    
    'For i = 1 To MyCol.Count
    '    OrCond = OrCond & "'" & MyCol(i).CompItemCode & "',"
    'Next i
    
    If Len(InList) > 0 Then
        InList = Mid$(InList, 1, Len(InList) - 1)
    End If
    'InList = InList & ")"
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select * From VVendAssItems", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            For i = 1 To MyCol.Count
                If MyCol(i).CompItemCode = ![ITemID] Then
                    cmbItems.AddItem ![ITemID] 'MyCol(i).ItemCode
                    cmbItems.List(cmbItems.ListCount - 1, 1) = MyCol(i).ItemCode
                    cmbItems.List(cmbItems.ListCount - 1, 2) = MyCol(i).lQty
                    cmbItems.List(cmbItems.ListCount - 1, 3) = ![Rate]
                    cmbItemDesc.AddItem MyCol(i).ItemDescription
                    Exit For
                End If
            Next i
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub cmbItem_matched()
    cmbItemName.ID = cmbItem.ID
End Sub

Private Sub cmbItemName_matched()
    cmbItem.ID = cmbItemName.ID
End Sub

Private Sub cmbDesc_Change()
    Call cmbDesc_Click
End Sub

Private Sub cmbDesc_Click()
    cmbItems.ListIndex = cmbDesc.ListIndex
End Sub

Private Sub cmbItems_Change()
    Call cmbItems_Click
End Sub

Private Sub cmbItems_Click()

    If Not cmbItems.MatchFound Then Exit Sub
    cmbDesc.ListIndex = cmbItems.ListIndex
    'Call FillList
    Call FillVendList
    
End Sub
Private Sub FillList()

    On Error GoTo err
    
    LV.ListItems.Clear
    For i = 1 To MyCol.Count
        MyCol.Remove 1
    Next i
    
    Dim itm As ListItem
    
    With rsMain
        .Filter = adFilterNone
        If RcvingLot Then
            .Filter = "RecieptID='" & cmbItems.Text & "'"
        Else
            .Filter = "ItemCode='" & cmbItems.Text & "'"
        End If
        LV.ListItems.Clear
            
        Do Until .EOF
            If RcvingLot Then
                Set itm = LV.ListItems.Add(, ![LotNo] & "'", LV.ListItems.Count + 1)
                itm.Tag = ![RecieptID] & ""
            Else
                Set itm = LV.ListItems.Add(, ![OrderNo] & "|" & ![CompItemCode], LV.ListItems.Count + 1)
            End If
            
            'ITM.Tag = ![EntryID]
            itm.ListSubItems.Add , , ![CustCode] & ""
            If RcvingLot Then
                itm.ListSubItems.Add(, , ![LotNo] & "/" & ![OrderNo] & "").Tag = Val(![LotNo] & "")
            Else
                itm.ListSubItems.Add , , ![OrderNo] & ""
            End If
            
            itm.ListSubItems.Add , , ![ItemCode] & ""
            itm.ListSubItems.Add , , ![ItemDescription] & ""
            itm.ListSubItems.Add(, , Val(![IssQty] & "")).Tag = Val(![IssQty] & "")
            
            Dim Obj As New IssueItems
            
            Obj.LVIndex = LV.ListItems.Count + 1
                
            Obj.OrderQty = Val(![Qty] & "")
            Obj.FreeStock = Val(![FreeStock] & "")
            Obj.IssueableQty = Val(![IssQty] & "")
            If RcvingLot Then
                MyCol.Add Obj, ![LotNo] & "'"
            Else
                MyCol.Add Obj, ![OrderNo] & "|" & ![CompItemCode]
            End If
            Set Obj = Nothing
            
            .MoveNext
        Loop
        '.Close
    End With
    
    'Set rs = Nothing
    
    'Call AddItems
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub cmbProcess_matched()

    On Error GoTo err
    LV.ListItems.Clear
    LVIssued.ListItems.Clear
    LVMaker.ListItems.Clear
    
    Dim cmd As New ADODB.Command
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "SP_ReqAuthForIss"
        Set .ActiveConnection = con
        .Parameters("@UserID").value = UserID
        .Parameters("@ProcessID").value = cmbProcess.ID
        .Execute
        UserReqAuth = Val(.Parameters("@ReqAuth").value & "")
    End With
    
    IsLotProcess = LotProcess(cmbProcess.ID)
    
    If IsLotProcess And RcvLot(cmbProcess.ID) Then
        RcvingLot = True
        cmbItems.Visible = False
        DTPicker1.Visible = True
        Label2.Caption = "  Process                                    Date                                      Lot No"
        LV.Checkboxes = True
    Else
        RcvingLot = False
        cmbItems.Visible = True
        DTPicker1.Visible = False
        Label2.Caption = "  Process                                    Item Code                        Item Description"
        LV.Checkboxes = False
    End If
    
    

    With cmd
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        
        
        If IsLotProcess Then
            
            If RcvingLot Then
                LV.ColumnHeaders(3).Text = "Lot No/Order No"
                .CommandText = "SP_IssueItemsLots"
                txtQty.Locked = True
                txtRate1.Locked = True
            Else
                LV.ColumnHeaders(3).Text = "Order No"
                .CommandText = "SP_IssueItems"
                txtQty.Locked = False
                txtRate1.Locked = False
            End If
            cmbProcess.Tag = "Lot No"
        Else
            LV.ColumnHeaders(3).Text = "Order No"
            .CommandText = "SP_IssueItems"
            cmbProcess.Tag = ""
        End If
        
        If rsMain.State = 1 Then rsMain.Close
        rsMain.CursorType = adOpenForwardOnly
        rsMain.LockType = adLockReadOnly
        
        .Parameters("@ProcessID").value = cmbProcess.ID
            
        Set rsMain = .Execute
    End With
    
    Dim isAdded As Boolean
    
    With rsMain
        cmbItems.Clear
        cmbDesc.Clear
        
        If RcvingLot Then
            Do Until .EOF
                isAdded = False
                For i = 0 To cmbItems.ListCount - 1
                    If Trim(cmbItems.List(i)) = Trim(![RecieptID] & "") Then
                        isAdded = True
                        Exit For
                    End If
                Next i
                
                If Not isAdded Then
                    cmbItems.AddItem ![RecieptID] & ""
                    cmbDesc.AddItem ![RecieptID] & ""
                End If
                    
                
                Set itm = LV.ListItems.Add(, ![LotNo] & "'", LV.ListItems.Count + 1)
                itm.Tag = ![RecieptID] & ""
                
                'ITM.Tag = ![EntryID]
                itm.ListSubItems.Add , , ![CustCode] & ""
                
                itm.ListSubItems.Add(, , ![LotNo] & "/" & ![OrderNo] & "").Tag = Val(![LotNo] & "")
                
                itm.ListSubItems.Add , , ![ItemCode] & ""
                itm.ListSubItems.Add , , ![ItemDescription] & ""
                itm.ListSubItems.Add(, , Val(![IssQty] & "")).Tag = Val(![IssQty] & "")
                    
                .MoveNext
            Loop
        Else
            Do Until .EOF
                isAdded = False
                For i = 0 To cmbItems.ListCount - 1
                    If Trim(cmbItems.List(i)) = Trim(![ItemCode] & "") Then
                        isAdded = True
                        Exit For
                    End If
                Next i
                
                If Not isAdded Then
                    cmbItems.AddItem ![ItemCode] & ""
                    cmbDesc.AddItem ![ItemDescription] & ""
                End If
                
                
            
                Set itm = LV.ListItems.Add(, ![OrderNo] & "|" & ![CompItemCode], LV.ListItems.Count + 1)
            
            
                'ITM.Tag = ![EntryID]
                itm.ListSubItems.Add , , ![CustCode] & ""
            
                itm.ListSubItems.Add , , ![OrderNo] & ""
            
                
                itm.ListSubItems.Add , , ![ItemCode] & ""
                itm.ListSubItems.Add , , ![ItemDescription] & ""
                itm.ListSubItems.Add(, , Val(![IssQty] & "")).Tag = Val(![IssQty] & "")
                
                .MoveNext
            Loop
        End If
    End With
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbVender_matched()

    On Error GoTo err
    
    cmbVender.Tag = Val(con.Execute("Select Rate From VVendAssItems Where VendID=" & cmbVender.ID & " and ItemID='" & LV.SelectedItem.SubItems(3) & "'").Fields(0).value & "")
    
    Call ShowVenderHistory
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub ShowVenderHistory()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim itm As ListItem
    With rs
        .Open "Select Top 10 * From VVendIssItems Where VendID=" & cmbVender.ID & " and ProcessID=" & cmbProcess.ID, con, adOpenForwardOnly, adLockReadOnly
        
        LVMaker.ListItems.Clear
        Do Until .EOF
            Set itm = LVMaker.ListItems.Add(, , LVMaker.ListItems.Count + 1)
            itm.SubItems(1) = ![ItemCode] & ""
            itm.SubItems(2) = ![ItemDescription] & ""
            itm.SubItems(3) = Format(![DT] & "", "dd-MMM-yyyy")
            itm.SubItems(4) = Val(![IssQty] & "")
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPost_Click()

    On Error GoTo err
    
    If Not cmbVender.MatchFound Then Exit Sub
    If cmbVender.ID = "" Or cmbVender.Text = "" Then Exit Sub
    
    If RcvingLot Then
        PostForLots
        Exit Sub
    End If
    
    If Val(txtQty) < 1 Then Exit Sub
    
    If Val(txtQty) > Val(LV.SelectedItem.SubItems(5)) Then
        MsgBox "Invalid Qty", vbInformation
        Exit Sub
    End If
    
    If GotIssID = 0 Then
        'Dim IssID As String
        'For i = 1 To LVIssued.ListItems.Count
        '    If Val(cmbVender.ID) = Val(LVIssued.ListItems(i).ListSubItems(2).Key) Then
        '        IssID = LVIssued.ListItems(i).ListSubItems(2)
        '        GoTo GotID:
        '    End If
        'Next i
    
        GotIssID = Val(con.Execute("Select Max(right(RecieptID,Len(RecieptID)-10)) From VVendIssItems Where Convert(varchar,IssDT,6)='" & Format(DTPicker2, "dd MMM yy") & "' and Left(RecieptID,3)<>'HSR'").Fields(0).value & "")
        GotIssID = GotIssID + 1
    Else
        GotIssID = GotIssID + 1
    End If
    'IssID = "ISS-" & Val(LVMaker.SelectedItem.ListSubItems(1).Key) & "-" & _
     Format(ServerDate, "ddMMyy") & "-" & Format(Val(IssID) + 1, "000")
    IssID = "ISU-" & _
     Format(DTPicker2, "ddMMyy") & GotIssID
    
GotID:
        
    Dim itm As ListItem
    With itm
        'Set ITM = LVIssued.ListItems.Add(, Val(LV.SelectedItem.Key) & Val(LVMaker.SelectedItem.Key) & "'", LVIssued.ListItems.Count + 1)
        'Set ITM = LVIssued.ListItems.Add(, LV.SelectedItem.Key & "|" & LVIssued.ListItems.Count + 1, LVIssued.ListItems.Count + 1)
        Set itm = LVIssued.ListItems.Add(, , LVIssued.ListItems.Count + 1)
        itm.Tag = LV.SelectedItem.key
        itm.ListSubItems.Add , cmbVender.ID & "'", cmbVender.Text
        itm.ListSubItems.Add(, , IssID & "").Tag = DTPicker2
        itm.ListSubItems.Add , , LV.SelectedItem.SubItems(2)
        itm.ListSubItems.Add , , LV.SelectedItem.SubItems(3)
        itm.ListSubItems.Add , , LV.SelectedItem.SubItems(4)
        itm.ListSubItems.Add , , Val(cmbVender.Tag)
        itm.ListSubItems.Add , , Val(txtQty)
        If cmbProcess.Tag = "Lot No" Then
            itm.ListSubItems(6).Tag = LV.SelectedItem.SubItems(2) & ""
        End If
        
    End With
    
    'MyCol(LV.SelectedItem.Index).IssueableQty = MyCol(LV.SelectedItem.Index).IssueableQty - Val(txtQty)
    LV.SelectedItem.SubItems(5) = Val(LV.SelectedItem.SubItems(5)) - Val(txtQty)
    txtRate.Visible = False
    
    'Removing Item From Combos
    If LV.SelectedItem.SubItems(5) = 0 Then
        Dim ItemIndex As Long
        ItemIndex = cmbItems.ListIndex
        If ItemIndex <> -1 Then
            
            cmbItems.RemoveItem ItemIndex
            cmbDesc.RemoveItem ItemIndex
            cmbItems.Text = ""
            cmbDesc.Text = ""
            If cmbItems.ListCount >= ItemIndex And ItemIndex > 0 Then
                cmbItems.ListIndex = ItemIndex
            ElseIf cmbItems.ListCount > 0 Then
                cmbItems.ListIndex = 0
            End If
        End If
    End If
    'if cmbitems.
    Call LV_ItemClick(LV.SelectedItem)
    
    Call FillVendList
    txtQty = ""
    Exit Sub
err:
    MsgBox err.Description
        'End If
End Sub

Private Sub cmdSave_Click()

    If Saved Then
        FormProcessID = cmbProcess.ID
        Unload Me
        'Me.Show
        'cmbProcess.ID = cmbProcess.Tag
        Call ShowIssForm(FormProcessID)
    End If
    
End Sub

Private Sub cmdSaveClose_Click()
    
    If Saved Then
        Unload Me
    End If
End Sub

Private Sub cmdUnAssign_Click()

    On Error GoTo err
    If LV.ListItems.Count = 0 Then Exit Sub
    Dim iAffect As Integer
    
    con.Execute "Delete From VendAssItems Where EntryID=" & Val(LV.SelectedItem.key), iAffect
    If iAffect > 0 Then
        'Call cmbNick_Click
        cmdUnAssign.Visible = False
    End If
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub DTPicker1_Change()
    Call DTPicker1_Click
End Sub

Private Sub DTPicker1_Click()
    On Error GoTo err
    
    
    With rsMain
    
        If DTPicker1.value & "" = "" Then
            .Filter = adFilterNone
        Else
            .Filter = "RecieptID Like 'HS-" & Format(DTPicker1, "ddMMyy") & "-*'"
            '.Filter = "Left(RecieptID,9) ='HS-" & Format(DTPicker1, "ddMMyy") & "'"
        End If
                
        Dim isAdded As Boolean
        
        cmbItems.Clear
        cmbDesc.Clear
        LV.ListItems.Clear
        Do Until .EOF
            isAdded = False
            For i = 0 To cmbItems.ListCount - 1
                If Trim(cmbItems.List(i)) = Trim(![RecieptID] & "") Then
                    isAdded = True
                    Exit For
                End If
            Next i
            
            If Not isAdded Then
                cmbItems.AddItem ![RecieptID] & ""
                cmbDesc.AddItem ![RecieptID] & ""
            End If
            .MoveNext
        Loop
    End With
    
    
        
    
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Exit Sub 'For Now
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


Private Sub LVIssued_DblClick()
    
    
    If LVIssued.ListItems.Count = 0 Then Exit Sub
    'If cmbCust.MatchFound = False Then Exit Sub
    If LV.ColumnHeaders(3).Text = "Lot No" And RcvingLot Then Exit Sub
    
    With txtRate1
        .Move LVIssued.Left + LVIssued.ColumnHeaders(8).Left, LVIssued.Top + LVIssued.SelectedItem.Top, LVIssued.ColumnHeaders(8).Width
        .Text = LVIssued.SelectedItem.SubItems(7) 'LVMaker.SelectedItem.SubItems(3)
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
End Sub

Private Sub LVIssued_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo err
    
    If KeyCode = 46 Then
        MyCol(LVIssued.SelectedItem.Tag).IssueableQty = MyCol(LVIssued.SelectedItem.Tag).IssueableQty + Val(LVIssued.SelectedItem.SubItems(7))
        LV.ListItems(LVIssued.SelectedItem.Tag).SubItems(5) = MyCol(LVIssued.SelectedItem.Tag).IssueableQty
        LVIssued.ListItems.Remove LVIssued.SelectedItem.Index
        Call LV_ItemClick(LV.SelectedItem)
        
        'Reset The IssID's
        Dim TempID As Long
        TempID = 0
        If LVIssued.ListItems.Count > 0 Then
            For i = LVIssued.SelectedItem.Index To LVIssued.ListItems.Count
                TempID = Val(Right(LVIssued.ListItems(i).SubItems(2), 3)) - 1
                If TempID > 1 Then
                    LVIssued.ListItems(i).SubItems(2) = "ISU-" & _
                     Format(ServerDate, "ddMMyy") & "-" & Format(TempID, "000")
                End If
            Next i
        Else
            GotIssID = 0
            
        End If
        If GotIssID > 0 Then GotIssID = GotIssID - 1
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub LvMaker_DblClick()
    
    Exit Sub
    If LVMaker.ListItems.Count = 0 Then Exit Sub
    'If cmbCust.MatchFound = False Then Exit Sub

    
    With txtRate
        .Move LVMaker.Left + LVMaker.ColumnHeaders(4).Left, LVMaker.Top + LVMaker.SelectedItem.Top, LVMaker.ColumnHeaders(4).Width
        '.Text = MyCol(LV.SelectedItem.SubItems(2) & "|" & LV.SelectedItem.SubItems(3)).IssueableQty 'LVMaker.SelectedItem.Tag 'LVMaker.SelectedItem.SubItems(3)
        .Text = MyCol(LV.SelectedItem.key).IssueableQty 'LVMaker.SelectedItem.Tag 'LVMaker.SelectedItem.SubItems(3)
        .Visible = True
        
        If LV.ColumnHeaders(3).Text = "Lot No" Then
            If RcvingLot Then
                .Locked = True
            Else
                .Locked = False
            End If
        Else
            .Locked = False
        End If
        
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With

    

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    
    On Error GoTo err
    
    'cmbCust.ListHeight = 2600
    'cmbCustName.ListHeight = 2600
    'cmbOrders.ListHeight = 2600
    
    'cmbProcess.Enabled = True
    cmbProcess.ListHeight = 2600
    cmbProcess.AddVals con, "Description", "Processes", "ProcessID", " Where ProcessID<>1 Order By SNo"
    
    DTPicker2.value = ServerDate
    DTPicker1.value = ServerDate
    DTPicker1.value = ""
    SSTab1.Tab = 0
    IsLotProcess = False
    cmbVender.ListHeight = 2000
    GotIssID = 0
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Public Sub LockProc(ProcID As Long)
    
    
    cmbProcess.Tag = ProcID
    
    'cmbProcess.Enabled = False
    Me.Refresh
    Me.Show
    
    
    cmbProcess.ID = ProcID
    
End Sub
Public Sub RefreshLV()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim itm As ListItem
    
    With rs
        .Open "Select *  from FCustomerCatalog Where CustCode='" & cmbCust.ID & "' and Country='" & cmbCountry.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set itm = LV.ListItems.Add(, ![CustCode] & ![Country] & ![ITemID], ![ITemID] & "")
            itm.SubItems(1) = ![Description] & ""
            itm.SubItems(2) = ![Unit] & ""
            itm.SubItems(3) = ![FOB] & ""
            itm.SubItems(4) = ![CIFSea] & ""
            itm.SubItems(5) = ![CIFAir] & ""
            itm.SubItems(6) = ![CnFSea] & ""
            itm.SubItems(7) = ![CnFAir] & ""
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub Form_Resize()

    On Error Resume Next
    
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
    
    Dim TitleBottom As Long

    TitleBottom = DTPicker1.Top + DTPicker1.Height + 50 '+ fTitle.Height
 
    'LV.Move 50, TitleBottom, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + Frame6.Height + 100)
    'LV.Move 50, TitleBottom ', Me.ScaleWidth - 100 , Me.ScaleHeight - (TitleBottom + Frame6.Height + 100)
    LV.Move Label2.Left, TitleBottom
    
    Label3.Move Label2.Left, LV.Top + LV.Height + 50
    cmbVender.Move Label3.Left, Label3.Top + Label3.Height - 15
    txtQty.Move cmbVender.Left + cmbVender.Width + 10, Label3.Top + Label3.Height - 15
    
    DTPicker2.Move txtQty.Left + txtQty.Width + 20, Label3.Top + Label3.Height - 15
    
    cmdPost.Move DTPicker2.Left + DTPicker2.Width + 20, Label3.Top
    
    SSTab1.Move LV.Left, cmbVender.Top + cmbVender.Height + 50
    
    'Picture1.Move LV.Left + LV.Width + 40, TitleBottom - 50
    
    'Frame6.Move Me.ScaleWidth - Frame6.Width - 50, Me.ScaleHeight - Frame6.Height - 50
    
    
    
End Sub
Private Sub FillVendList()

    On Error GoTo err
        
    Dim VendIDList As String
    VendIDList = ""
    For i = 1 To LVIssued.ListItems.Count
        With LVIssued.ListItems(i)
            If LV.SelectedItem.key = .Tag Then
                VendIDList = VendIDList & Val(.ListSubItems(1).key) & ","
            End If
        End With
    Next i
    
    If VendIDList <> "" Then
        VendIDList = "(" & Left(VendIDList, Len(VendIDList) - 1) & ")"
    Else
        VendIDList = "(0)"
    End If
    
    'With rs
    '    .Open "Select VendID,Rate,VenderName From VVendAssItems Where ItemID='" & LV.SelectedItem.SubItems(3) & "' and PhaseID=" & cmbProcess.ID & " and VendID Not In" & VendIDList, con, adOpenForwardOnly, adLockReadOnly
    '    LVMaker.ListItems.Clear
    '    Do Until .EOF
    '
    '        'Set ITM = LVMaker.ListItems.Add(, ![VendID] & "'", LVMaker.ListItems.Count + 1)
    '        Set ITM = LVMaker.ListItems.Add(, , LVMaker.ListItems.Count + 1)
    '        ITM.Tag = Val(MyCol(Item.Index).IssueableQty & "")
    '        ITM.ListSubItems.Add , ![VendID] & "'", .Fields(2) & ""
    '        ITM.ListSubItems.Add , , Val(![Rate] & "")
    '
    '        ITM.ListSubItems.Add(, , 0).Tag = LV.SelectedItem.SubItems(2) & ""
    '        .MoveNext
    '    Loop
    '    .Close
    'End With
    Call cmbVender.ClearVals
    'If RcvingLot Then
    '    Call cmbVender.AddVals(con, "VenderName", "VVendAssItems", "VendID", " Where ItemID='" & LV.SelectedItem.SubItems(3) & "' and (PhaseID=" & cmbProcess.ID & " Or PhaseID=-1) ")
    'Else
    If HicoVisible And RcvingLot Then
        Call cmbVender.AddVals(con, "Distinct VenderName", "VVendAssItems", "VendID", " Where (ItemID='" & LV.SelectedItem.SubItems(3) & "' and PhaseID=" & cmbProcess.ID & "  and VendID Not In" & VendIDList & " and VendType=1 Or (PhaseID=-1))")
    Else
        Call cmbVender.AddVals(con, "Distinct VenderName", "VVendAssItems", "VendID", " Where ItemID='" & LV.SelectedItem.SubItems(3) & "' and (PhaseID=" & cmbProcess.ID & ") and VendID Not In" & VendIDList & " and VendType=1")
    End If
    'End If
    
    
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)

    On Error GoTo err
    Call FillVendList
    
    'Dim rs As New ADODB.Recordset
    'Dim ITM As ListItem
    
    Dim VendIDList As String
    VendIDList = ""
    For i = 1 To LVIssued.ListItems.Count
        With LVIssued.ListItems(i)
            If LV.SelectedItem.key = .Tag Then
                VendIDList = VendIDList & Val(.ListSubItems(1).key) & ","
            End If
        End With
    Next i
    
    If VendIDList <> "" Then
        VendIDList = "(" & Left(VendIDList, Len(VendIDList) - 1) & ")"
    Else
        VendIDList = "(0)"
    End If
    
    'With rs
    '    .Open "Select VendID,Rate,VenderName From VVendAssItems Where ItemID='" & LV.SelectedItem.SubItems(3) & "' and PhaseID=" & cmbProcess.ID & " and VendID Not In" & VendIDList, con, adOpenForwardOnly, adLockReadOnly
    '    LVMaker.ListItems.Clear
    '    Do Until .EOF
    '        'Set ITM = LVMaker.ListItems.Add(, ![VendID] & "'", LVMaker.ListItems.Count + 1)
    '        Set ITM = LVMaker.ListItems.Add(, , LVMaker.ListItems.Count + 1)
    '        ITM.Tag = Val(MyCol(Item.Index).IssueableQty & "")
    '        ITM.ListSubItems.Add , ![VendID] & "'", .Fields(2) & ""
    '        ITM.ListSubItems.Add , , Val(![Rate] & "")
    '
    '        ITM.ListSubItems.Add(, , 0).Tag = LV.SelectedItem.SubItems(2) & ""
    '        .MoveNext
    '    Loop
    '    .Close
    'End With
    'Set rs = Nothing
    
    'txtQty = Val(MyCol(Item.Index).IssueableQty & "")
    txtQty = LV.SelectedItem.ListSubItems(5)
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub



Private Sub txtRate_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    
    If KeyAscii = 13 Then
        'LVMaker.SelectedItem.SubItems(3) = Val(txtRate)
        
        Dim IssID As String
        For i = 1 To LVIssued.ListItems.Count
            If Val(LVMaker.SelectedItem.ListSubItems(1).key) = Val(LVIssued.ListItems(i).ListSubItems(2).key) Then
                IssID = LVIssued.ListItems(i).ListSubItems(2)
                GoTo GotID:
            End If
        Next i
        
        IssID = Val(con.Execute("Select Max(right(RecieptID,3)) From VendIssued Where DT='" & Format(ServerDate, "dd-MMM-yyyy") & "'").Fields(0).value & "")
        
        'IssID = "ISS-" & Val(LVMaker.SelectedItem.ListSubItems(1).Key) & "-" & _
         Format(ServerDate, "ddMMyy") & "-" & Format(Val(IssID) + 1, "000")
         IssID = "ISU-" & _
         Format(ServerDate, "ddMMyy") & "-" & Format(Val(IssID) + 1, "000")
GotID:
        
        Dim itm As ListItem
        With itm
            'Set ITM = LVIssued.ListItems.Add(, Val(LV.SelectedItem.Key) & Val(LVMaker.SelectedItem.Key) & "'", LVIssued.ListItems.Count + 1)
            'Set ITM = LVIssued.ListItems.Add(, LV.SelectedItem.Key & "|" & LVIssued.ListItems.Count + 1, LVIssued.ListItems.Count + 1)
            Set itm = LVIssued.ListItems.Add(, , LVIssued.ListItems.Count + 1)
            itm.Tag = LV.SelectedItem.key
            itm.ListSubItems.Add , LVMaker.SelectedItem.ListSubItems(1).key, LVMaker.SelectedItem.SubItems(1)
            itm.ListSubItems.Add , , IssID & ""
            itm.ListSubItems.Add , , LV.SelectedItem.SubItems(2)
            itm.ListSubItems.Add , , LV.SelectedItem.SubItems(3)
            itm.ListSubItems.Add , , LV.SelectedItem.SubItems(4)
            itm.ListSubItems.Add , , LVMaker.SelectedItem.ListSubItems(2)
            itm.ListSubItems.Add , , Val(txtRate)
            If cmbProcess.Tag = "Lot No" Then
                itm.ListSubItems(6).Tag = LV.SelectedItem.SubItems(2) & ""
            End If
        End With
        
        'MyCol(LV.SelectedItem.Index).IssueableQty = MyCol(LV.SelectedItem.Index).IssueableQty - Val(txtRate)
        'LV.SelectedItem.ListSubItems(5) = Val(LV.SelectedItem.ListSubItems(5)) - Val(txtRate)
        LV.SelectedItem.SubItems(5) = Val(LV.SelectedItem.SubItems(5)) - Val(txtRate)
        txtRate.Visible = False
        Call LV_ItemClick(LV.SelectedItem)
        'End If
    ElseIf KeyAscii = 27 Then
        txtRate.Visible = False
    End If
    Call OnlyNums(KeyAscii)
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Function Saved() As Boolean

    On Error GoTo err
    Dim SaveIt As Boolean
    
    Dim VendID As New Collection
    Dim VIssID As New Collection
    
    Dim cmd As New ADODB.Command
    Dim TempIndex As Integer
    TempIndex = 1
    
    Dim lRefID As Long
    
    Dim IssLot As Boolean
    Dim iAuth As Integer 'Authorize
    
    IssLot = RcvingLot 'RcvLot(cmbProcess.ID)
    Call StartTrans(con)
    
    For i = 1 To LVIssued.ListItems.Count
        With LVIssued.ListItems(i)
            SaveIt = True
            For j = 1 To VendID.Count
                If Val(LVIssued.ListItems(i).ListSubItems(1).key) = VendID(j) Then
                    'It Means That It's Been Saved Already.
                    SaveIt = False
                    Exit For
                End If
            Next j
        
            
            If SaveIt Then
                
                con.Execute "Insert Into VendIssued(DT,VendID) Values('" & _
                 ServerDate & " " & ServerTime & "'," & Val(LVIssued.ListItems(i).ListSubItems(1).key) & ")"
                 
                lRefID = Val(con.Execute("Select Max(EntryID) From VendIssued").Fields(0).value & "")
        
                VendID.Add Val(LVIssued.ListItems(i).ListSubItems(1).key)
                VIssID.Add lRefID
            Else
                lRefID = VIssID(j)
            End If
        
            '1. Insert Into VendIssuedItems
            
            If cmbProcess.Tag = "Lot No" And IssLot Then
                With cmd
                    .ActiveConnection = con
                    .CommandType = adCmdStoredProc
                    .CommandText = "SP_UpdateForIARLots"
                    
                    .Parameters("@ProcessID").value = cmbProcess.ID
                    .Parameters("@Qty").value = Val(LVIssued.ListItems(i).SubItems(7))
                    .Parameters("@LotNo").value = Val(LVIssued.ListItems(i).ListSubItems(6).Tag)
                    .Parameters("@RefiD").value = lRefID
                    .Parameters("@Rate").value = Val(LVIssued.ListItems(i).SubItems(6))
                    .Parameters("@RecID").value = LVIssued.ListItems(i).SubItems(2)
                    .Parameters("@ReqAuth").value = UserReqAuth
                    .Execute
                End With
                
            Else
                con.Execute "Insert Into VendIssdDetail(RefID,IssDT,RecieptID,OrderNo,ItemCode" & _
                 ",Rate,ProcessID,IssQty,ReqAuth) Values('" & lRefID & "','" & _
                 .ListSubItems(2).Tag & "','" & .SubItems(2) & _
                 "','" & .SubItems(3) & "','" & .SubItems(4) & "'," & _
                 Val(.SubItems(6)) & "," & cmbProcess.ID & "," & _
                 Val(.SubItems(7)) & "," & UserReqAuth & ")"
                 
                 'Do Not Execute Stored Procedure
                 'If This User Requires Authorization...
                 If UserReqAuth = 0 Then
                    With cmd
                        .ActiveConnection = con
                        .CommandType = adCmdStoredProc
                        .CommandText = "SP_UpdateForIAR"
                        
                        .Parameters("@OrderNo").value = LVIssued.ListItems(i).SubItems(3)
                        .Parameters("@ItemCode").value = LVIssued.ListItems(i).SubItems(4)
                        .Parameters("@ProcessID").value = cmbProcess.ID
                        .Parameters("@Qty").value = Val(LVIssued.ListItems(i).SubItems(7))
                        .Execute
                    End With
                    
                End If
            End If
            'Authentication Info...
            If UserReqAuth = 0 Then
                Dim DetailRefID As Long
                DetailRefID = Val(con.Execute("Select Max(EntryID) From VendIssdDetail").Fields(0).value & "")
                con.Execute "Insert Into VendIssAuthDetail(UserID,IssID,DT) Values(" & _
                 UserID & "," & DetailRefID & ",'" & ServerDate & " " & ServerTime & "')"
            End If
            
        End With
        
    Next i
                       
    
    '2. Now If Issued After Received Then
    'Update The IssQty In VendRecievedItems....
    '-----------------------------------------------------------------
    'If IssLot Then
    '    For i = 1 To LV.ListItems.Count
    '        Dim lQty As Long
    '        lQty = Val(LV.ListItems(i).ListSubItems(5).Tag) - Val(LV.ListItems(i).ListSubItems(5))
    '        If Val(LV.ListItems(i).ListSubItems(5)) <> Val(LV.ListItems(i).ListSubItems(5).Tag) Then
    '            If UserReqAuth = 0 Then
    '                con.Execute "Update VendRcvdDetail Set IssQty=RcvdQty Where LotNo=" & LV.ListItems(i).ListSubItems(2).Tag
    '            End If
    '        End If
    '    Next
    'End If
    
    If IssLot Then
        For i = 1 To LVIssued.ListItems.Count
            'Dim lQty As Long
            'lQty = Val(LV.ListItems(i).ListSubItems(5).Tag) - Val(LV.ListItems(i).ListSubItems(5))
            If Left(LVIssued.ListItems(i).SubItems(2), 2) = "HS" Then
                If UserReqAuth = 0 Then
                    con.Execute "Update VendRcvdDetail Set IssQty=RcvdQty Where RecieptID='" & LVIssued.ListItems(i).SubItems(2) & "'"
                End If
            End If
        Next
    End If
    
    '-----------------------------------------------------------------
    
    con.CommitTrans
    
    Saved = True
    Exit Function
err:
    MsgBox err.Description
    
End Function

Private Sub txtRate1_KeyPress(KeyAscii As Integer)

    On Error GoTo err
    
    If KeyAscii = 13 Then
        'LVMaker.SelectedItem.SubItems(3) = Val(txtRate)
        Dim itm As ListItem
        
        'MyCol(LVIssued.SelectedItem.Tag).IssueableQty = MyCol(LVIssued.SelectedItem.Tag).IssueableQty + Val(LVIssued.SelectedItem.SubItems(7)) - Val(txtRate1)
        LV.ListItems(LVIssued.SelectedItem.Tag).SubItems(5) = MyCol(LVIssued.SelectedItem.Tag).IssueableQty
        LVIssued.SelectedItem.SubItems(7) = Val(txtRate1)
        txtRate1.Visible = False
        Call LV_ItemClick(LV.SelectedItem)
        'End If
        If Val(txtRate1) = 0 Then
            Call LVIssued_KeyDown(46, 0)
        End If
        
    ElseIf KeyAscii = 27 Then
        txtRate1.Visible = False
    End If
    Call OnlyNums(KeyAscii)
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub txtRate1_LostFocus()
    txtRate1.Visible = False
End Sub
Private Sub txtRate_LostFocus()
    txtRate.Visible = False
End Sub
Private Sub PostForLots()

    On Error GoTo err
    Dim itm As ListItem

    For i = 1 To LV.ListItems.Count
        If LV.ListItems(i).Checked Then
            IssID = LV.ListItems(i).Tag
            With itm
                Set itm = LVIssued.ListItems.Add(, , LVIssued.ListItems.Count + 1)
                itm.Tag = LV.ListItems(i).key
                itm.ListSubItems.Add , cmbVender.ID & "'", cmbVender.Text
                itm.ListSubItems.Add(, , IssID & "").Tag = DTPicker2
                itm.ListSubItems.Add , , LV.ListItems(i).SubItems(2)
                itm.ListSubItems.Add , , LV.ListItems(i).SubItems(3)
                itm.ListSubItems.Add , , LV.ListItems(i).SubItems(4)
                itm.ListSubItems.Add , , Val(cmbVender.Tag)
                itm.ListSubItems.Add , , Val(LV.ListItems(i).SubItems(5))
                If cmbProcess.Tag = "Lot No" Then
                    itm.ListSubItems(6).Tag = LV.ListItems(i).SubItems(2) & ""
                End If
            End With
            LV.ListItems(i).SubItems(5) = 0
        End If
    Next
    
    
    'if cmbitems.
    'Call LV_ItemClick(LV.ListItems(i))
    
    Call FillVendList
    txtQty = ""
    Exit Sub
err:
    MsgBox err.Description
        
End Sub
