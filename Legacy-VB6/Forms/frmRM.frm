VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmRM 
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   15675
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8310
   ScaleWidth      =   15675
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   14535
      TabIndex        =   17
      Top             =   105
      Width           =   1110
   End
   Begin VB.PictureBox PicProduct 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2865
      Left            =   12855
      ScaleHeight     =   2835
      ScaleWidth      =   2760
      TabIndex        =   16
      Top             =   600
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
      Left            =   60
      TabIndex        =   15
      Top             =   8040
      Width           =   1560
   End
   Begin ComboList.Usercontrol1 cmbUsedIn 
      Height          =   285
      Left            =   9945
      TabIndex        =   14
      Top             =   330
      Width           =   4560
      _ExtentX        =   8043
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
   Begin VB.TextBox txtTotalValue 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   5520
      TabIndex        =   12
      Top             =   8040
      Width           =   1365
   End
   Begin VB.TextBox txtActiveArticles 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2985
      TabIndex        =   10
      Top             =   8040
      Width           =   1365
   End
   Begin ComboList.Usercontrol1 cmbItem 
      Height          =   285
      Left            =   4515
      TabIndex        =   8
      Top             =   330
      Width           =   5430
      _ExtentX        =   9578
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
   Begin ComboList.Usercontrol1 cmbGroup 
      Height          =   285
      Left            =   60
      TabIndex        =   7
      Top             =   330
      Width           =   4455
      _ExtentX        =   7858
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
   Begin Crystal.CrystalReport cr1 
      Left            =   390
      Top             =   630
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowBorderStyle=   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.PictureBox PicEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1275
      ScaleHeight     =   225
      ScaleWidth      =   5805
      TabIndex        =   5
      Top             =   1770
      Visible         =   0   'False
      Width           =   5835
      Begin VB.TextBox txtBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BorderStyle     =   0  'None
         DataField       =   "RptTime"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   3540
         MaxLength       =   12
         TabIndex        =   1
         Top             =   0
         Width           =   1680
      End
      Begin VB.TextBox txtAccTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BorderStyle     =   0  'None
         DataField       =   "RptTime"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   -15
         MaxLength       =   150
         TabIndex        =   0
         Top             =   0
         Width           =   3510
      End
      Begin VB.ComboBox cmbCrDr 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "frmRM.frx":0000
         Left            =   5250
         List            =   "frmRM.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   -45
         Width           =   600
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1695
      Top             =   3330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRM.frx":0016
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRM.frx":046A
            Key             =   "ORD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRM.frx":08C2
            Key             =   "CAT1"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRM.frx":0D16
            Key             =   "ITMIN"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRM.frx":116A
            Key             =   "CAT"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRM.frx":15BE
            Key             =   "OK"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRM.frx":1A16
            Key             =   "TYPE"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV 
      Height          =   7410
      Left            =   60
      TabIndex        =   3
      Top             =   600
      Width           =   15585
      _ExtentX        =   27490
      _ExtentY        =   13070
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Device Code"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Device Description"
         Object.Width           =   10583
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Stock"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Rate"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Group"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Unit"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "W / L"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Re-Ord"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Open Pos"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "Min. Stock"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "Max. Stock"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Label lblTotalValue 
      AutoSize        =   -1  'True
      Caption         =   "Total Value :"
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
      Left            =   4440
      TabIndex        =   13
      Top             =   8085
      Width           =   1035
   End
   Begin VB.Label lblActiveArticles 
      AutoSize        =   -1  'True
      Caption         =   "Total No. :"
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
      Left            =   2115
      TabIndex        =   11
      Top             =   8070
      Width           =   825
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   1
      Left            =   60
      TabIndex        =   9
      Top             =   105
      Width           =   14445
      BackColor       =   11517387
      Caption         =   $"frmRM.frx":1E6A
      Size            =   "25479;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdOpt 
      Height          =   390
      Left            =   8700
      TabIndex        =   4
      Top             =   8025
      Width           =   1545
      Caption         =   "Options        "
      PicturePosition =   327683
      Size            =   "2725;688"
      Accelerator     =   79
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   390
      Left            =   10275
      TabIndex        =   6
      Top             =   8025
      Width           =   1545
      ForeColor       =   64
      Caption         =   "Close          "
      PicturePosition =   327683
      Size            =   "2725;688"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Menu mnupop 
      Caption         =   "mnupop"
      Begin VB.Menu mnurefresh 
         Caption         =   "Re&fresh"
      End
      Begin VB.Menu mnuShowAll 
         Caption         =   "&Show Inactive Raw Material"
      End
      Begin VB.Menu mnuStockHistory 
         Caption         =   "View &Stock History For Last 30 Days"
      End
      Begin VB.Menu dash0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewItem 
         Caption         =   "&Add New Raw Material"
      End
      Begin VB.Menu mnuAddNewRMCopy 
         Caption         =   "Add New Raw Material (Copy Selected)"
      End
      Begin VB.Menu mnuNewCat 
         Caption         =   "Add New RM Group"
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Edit Raw Material"
      End
      Begin VB.Menu mnuDeleteDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteRawMaterial 
         Caption         =   "Delete Raw Material"
      End
      Begin VB.Menu dash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActive 
         Caption         =   "Make This Raw Material Ina&ctive"
      End
      Begin VB.Menu mnuParent 
         Caption         =   "Make This Account Paren&t"
         Visible         =   0   'False
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintRMSpecification 
         Caption         =   "Print RM Sepcification"
      End
      Begin VB.Menu mnuPrintRMDetails 
         Caption         =   "Print RM Details"
      End
      Begin VB.Menu mnuPrintRMLable 
         Caption         =   "Print RM Lable"
      End
      Begin VB.Menu mnuPrintRMRLable 
         Caption         =   "Print RM R Lable"
      End
      Begin VB.Menu mnuPrintRMBarcode 
         Caption         =   "Print RM Barcode"
      End
      Begin VB.Menu mnuPOrd 
         Caption         =   "Place Purchase Order"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSales 
         Caption         =   "&Enter Sales"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPurchases 
         Caption         =   "Enter Purchases"
         Visible         =   0   'False
      End
      Begin VB.Menu dash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print Raw Material List"
      End
      Begin VB.Menu mnuPrintRMRackwise 
         Caption         =   "Raw Material List (Rackwise)"
      End
      Begin VB.Menu mnuPrintRMRackwiseGrouped 
         Caption         =   "Raw Material List (Groupwise)"
      End
      Begin VB.Menu mnuPrintNonZero 
         Caption         =   "Print Raw Material List (Non-Zero Qty)"
      End
      Begin VB.Menu mnuPrintRMListWithValue 
         Caption         =   "Print Raw Material List (With Value)"
      End
      Begin VB.Menu mnuPrintRMListWithUsedIn 
         Caption         =   "Print Raw Material List (With Used-In)"
      End
      Begin VB.Menu mnuPrintRMListCrossReference 
         Caption         =   "Print Material Lists (Cross Reference)"
      End
      Begin VB.Menu mnuPrintRMListCrossNo 
         Caption         =   "Print Material Lists (Cross No.)"
      End
      Begin VB.Menu mnuPrintRateHistory 
         Caption         =   "Print Rate History"
      End
      Begin VB.Menu mnuPrintBal 
         Caption         =   "Print Report(Internal)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintStockReport 
         Caption         =   "Print Stock Report (Date Range)"
      End
      Begin VB.Menu mnuPrintGeneralHistory 
         Caption         =   "Print General Material Purchase History"
      End
      Begin VB.Menu mnuBelowMinLevel 
         Caption         =   "Below Min. Level"
      End
      Begin VB.Menu mnuAboveMaxLevel 
         Caption         =   "Above Max. Level"
      End
   End
End
Attribute VB_Name = "frmRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rec As New ADODB.Recordset
Dim rsAccounts As New ADODB.Recordset
'Dim i As Integer
Dim Bal As Double
Dim AccList As String
Dim SubAccList As String
Dim TotalBalance As Double
Dim PrevIndex As Long
Dim ShowAll As Boolean
Dim bSampling As Boolean
Dim strSQL_Reports As String

Private Sub Command2_Click()

    Dim AccNo As String
    Dim AccTitle As String
    Dim Bal As Double

    For i = 1 To LV.ListItems.count
        AccNo = LV.ListItems(i).Text & ""
        AccTitle = LV.ListItems(i).ListSubItems(1).Text
        Bal = Val(LV.ListItems(i).ListSubItems(2).Tag)
        con.Execute "Delete * From TempChaccount"
        con.Execute "Insert into TempChAccount Values('" & AccNo & "','" & AccTitle & "'," & Bal & ")"
    Next i
    
End Sub

Private Sub FillItems(strOrderCond As String, Optional startIndx As Long)
    
    'LV.Sorted = False
    Set LV.SmallIcons = ImageList1
    'If startIndx = 0 Then
    '    startIndx = LV.ListItems.Count
    'End If
    Dim dValue As Double
    startIndx = 0
    Dim rs As New ADODB.Recordset
    Dim Cnd As String, ITM As ListItem

    'If Not ShowAll Then Cnd = " Where InActive=0 and GroupID=" & CatID Else Cnd = " Where GroupID=" & CatID
    Cnd = " WHERE Sampling=" & Abs(bSampling)
    If cmbGroup.ID <> "0" And cmbGroup.ID <> "" Then
        Cnd = Cnd & " AND GroupID=" & cmbGroup.ID & ""
    End If
    
    If Not ShowAll Then
        Cnd = Cnd & " AND InActive=0"
    End If
    
    'LV.ListItems.Clear
    
    
    If cmbUsedIn.ID <> "0" Then
        If cmbUsedIn.Functionality = 1 Then
            Cnd = Cnd & " AND VRM.RMID IN(SELECT RMID FROM ItemsRMComp WHERE ItemID IN(" & cmbUsedIn.ID & "))"
        Else
            Cnd = Cnd & " AND VRM.RMID IN(SELECT RMID FROM ItemsRMComp WHERE ItemID IN('" & cmbUsedIn.ID & "'))"
        End If
    End If
    If strOrderCond = "" Then
        strOrderCond = " ORDER BY RMID1"
    End If
        
    strSQL_Reports = "SELECT * FROM VRM LEFT JOIN ItemsRMComp ON VRM.RMID=ItemsRMComp.RMID" & Cnd
    dValue = 0
    With rs
        .CursorLocation = adUseClient
        .Open "SELECT * FROM VRMWithRMOpenPOs VRM " & Cnd & strOrderCond, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            
            If ![Inactive] Then
               clr = &H707070
            Else
               clr = Val(![GroupColor] & "")
            End If
            
            Set ITM = LV.ListItems.add(startIndx + 1, ![RMID] & "RM'", ![RMID1], , IIf(![Inactive] = 0, "ITM", "ITMIN"))
            
            ITM.ListSubItems.add(, , "     »  " & ![RMName] & " " & ![RMSize] & " " & ![SizeUnit]).ForeColor = clr
            ITM.ListSubItems.add(, , Round(Val(!QtyInStock & ""), 2)).ForeColor = clr
            ITM.ListSubItems.add , , Val(![Rate] & "")
            ITM.ListSubItems.add , , !GroupName & ""
            
            ITM.ListSubItems.add(, , ![Unit] & "").ForeColor = clr
            
            ITM.ListSubItems.add(, , ![Weight_Length] & "").ForeColor = clr
            
            If Val(![ReorderPoint] & "") > Val(![QtyInStock] & "") Then
                clr = vbRed
            End If
            
            ITM.ListSubItems.add(, , Val(![ReorderPoint] & "")).ForeColor = clr
            ITM.ListSubItems.add(, , Val(!OpenPOsQty & "")).ForeColor = clr
            
            ITM.ListSubItems.add(, , Val(!MinLevel & "")).ForeColor = clr
            ITM.ListSubItems.add(, , Val(!MaxLevel & "")).ForeColor = clr
            
            If ![Inactive] = 0 Then ITM.Tag = "Active" Else ITM.Tag = "Inactive"
            dValue = dValue + (Val(!QtyInStock & "") * Val(!Rate & ""))
            .MoveNEXT
            startIndx = startIndx + 1
        Next
        .Close
    End With
    txtActiveArticles.Text = LV.ListItems.count 'lCount
    txtTotalValue = Format(dValue, "#,##0.00")
    
    Set rs = Nothing
    LV.Visible = True
    'LV.Sorted = True
    If PrevIndex > 0 And PrevIndex <= LV.ListItems.count Then LV.ListItems(PrevIndex).Selected = True
    On Error Resume Next
    LV.SelectedItem.EnsureVisible
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmbCrDr_Click()
 If cmbCrDr.Text = "CR" Then
  cmbCrDr.Tag = -Val(txtBal.Text)
 Else
  cmbCrDr.Tag = Val(txtBal.Text)
 End If
End Sub

Private Sub cmbCrDr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call UpdateAccount
End If
End Sub

Private Sub cmbGroup_matched()
    Call AddHeads
End Sub

Private Sub cmbItem_matched()
    Call GoToSelected
End Sub

Private Sub GoToSelected()
    
    LV.ListItems(cmbItem.ID & "RM'").EnsureVisible
    'LV.FindItem "     »  " & cmbitem.ID
    LV.ListItems(cmbItem.ID & "RM'").Selected = True
    LV.SetFocus
    
End Sub



Private Sub cmbUsedIn_matched()
    Call AddHeads
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdOpt_Click()
With cmdOpt
 Call CheckMenu
 Me.PopUpMenu mnuPop, 255, .Left, .Top + .Height
End With
End Sub



Private Sub cmdRefresh_Click()
    Call mnuRefresh_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 And PicEdit.Visible Then
       PicEdit.Visible = False
    ElseIf KeyAscii = 13 Then
     'SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()

    Me.KeyPreview = True
    mnuPop.Visible = False
    
    Dim Condition As String
    If ShowAll Then
        Condition = ""
    Else
        Condition = " WHERE Active=1 "
    End If
    'Set LV.SmallIcons = ImageList1
    'rsAccounts.Open "Select * from Accounts " & Condition & " Order By Accno", Con, adOpenForwardOnly
    'Call RefreshList
    cmbGroup.ListHeight = 2000
    Call cmbGroup.AddItem("<All Groups>", "0")
    Call cmbGroup.AddVals(con, "Description", "RMGroups", "ID", " ORDER BY Description")
    
    cmbUsedIn.ListHeight = 3000
    cmbUsedIn.EnableFunctionality = True
    cmbUsedIn.AddItem "<All>", "0"
    cmbUsedIn.AddVals con, "DISTINCT ItemID", "ItemsRMComp", "ItemID"
    cmbUsedIn.ID = "0"
    
    cmbItem.DropDownOffLine = False
    cmbItem.ListHeight = 6000
    
    If strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
        mnuPrintStockReport.Visible = True
    Else
        mnuPrintStockReport.Visible = False
    End If
    'Call AddHeads
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    'LV.Move 50, 50, Me.ScaleWidth - 100, Me.ScaleHeight - (cmdClose.Height + 100)
    LV.Move 50, 600, Me.ScaleWidth - 100, Me.ScaleHeight - (cmdClose.Height + 100 + cmbGroup.Height + cmbGroup.Top)
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, LV.Top + LV.Height + 50
    chkShowPicture.Move LV.Left, LV.Top + LV.Height + 50
    
    'cmdPrint.Move cmdClose.Left - cmdPrint.Width - 50, cmdClose.Top
    cmdOpt.Move cmdClose.Left - cmdOpt.Width - 50, cmdClose.Top
    
    lblActiveArticles.Top = cmdClose.Top + 50
    txtActiveArticles.Top = cmdClose.Top
    
    lblTotalValue.Top = lblActiveArticles.Top
    txtTotalValue.Top = txtActiveArticles.Top
    
End Sub

Private Sub Form_Unload(cancel As Integer)
'rsAccounts.Close
Set rsAccounts = Nothing
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Select Case ColumnHeader.Index
        Case 1
            Call AddHeads
        Case 3
            Call AddHeads(" Order By Description")
    End Select
End Sub

Private Sub LV_DblClick()
    
    If LV.SelectedItem.SmallIcon = "CAT" Or LV.SelectedItem.SmallIcon = "CAT1" Then
        'Load frmNewItemCat
        'Call frmNewItemCat.EditCat(Val(LV.SelectedItem.Key))
        If LV.SelectedItem.SmallIcon = "CAT" Then
            LV.SelectedItem.SmallIcon = "CAT1"
            For i = LV.SelectedItem.Index + 1 To LV.ListItems.count
                If Not LV.ListItems(LV.SelectedItem.Index + 1).SmallIcon Like "CAT*" Then
                    LV.ListItems.Remove LV.SelectedItem.Index + 1
                Else
                    Exit For
                End If
            Next i
        Else
            'LV.SelectedItem.SmallIcon = "CAT"
            'Call FillItems(Val(LV.SelectedItem.Key), " Order By RMName", LV.SelectedItem.Index)
        End If
        
    ElseIf LV.SelectedItem.SmallIcon = "ITM" Then
        Load frmNewRM
        Call frmNewRM.EditItem(Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 3))
    End If
    
    Exit Sub

End Sub

Private Sub LV_GotFocus()
   If PicEdit.Visible Then Call Form_KeyPress(27)
End Sub

Private Sub LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If LV.ListItems.count = 0 Then Exit Sub
       If LV.SelectedItem.SmallIcon = "TYPE" Then Exit Sub
       Call CheckMenu
       Me.PopUpMenu mnuPop, , (LV.Left + (LV.Width / 2)), (LV.Top + LV.SelectedItem.Top)
    End If
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button <> 2 Then Exit Sub
    If LV.ListItems.count > 0 Then If LV.SelectedItem.SmallIcon = "TYPE" Then Exit Sub

    Call CheckMenu
    Me.PopUpMenu mnuPop
 
End Sub

Private Sub mnuAboveMaxLevel_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\RMList_Max.rpt")
    
    Dim f As New frmPrevRpt
    f.ShowReport "{VRM.InActive}=FALSE AND {VRM.QtyInStock}>{VRM.MaxLevel} AND {VRM.MaxLevel}>0", rpt
    
End Sub

Private Sub mnuActive_Click()

   If mnuActive.Tag = "Active" Then
      If Not MakeAccountInActive Then Exit Sub
   ElseIf mnuActive.Tag = "Inactive" Then
      If Not MakeAccountActive Then Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
'   Sleep (500)
   Call mnuRefresh_Click
   Screen.MousePointer = vbDefault
End Sub

Private Sub mnuAddNewRMCopy_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    If LV.SelectedItem.SmallIcon <> "ITM" Then Exit Sub
        
    Dim lRMID_For_Copy As Long
    lRMID_For_Copy = Val(LV.SelectedItem.key)
    With frmNewRM
        Load frmNewRM
        Call .AddNew(bSampling, True, lRMID_For_Copy)
    End With
    Call mnuRefresh_Click
        
        
End Sub

Private Sub mnuBelowMinLevel_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\RMList_Min.rpt")
    
    Dim f As New frmPrevRpt
    f.ShowReport "{VRM.InActive}=FALSE AND {VRM.QtyInStock}<{VRM.MinLevel}", rpt

End Sub

Private Sub mnuDeleteRawMaterial_Click()
    If LV.ListItems.count = 0 Then Exit Sub
    If LV.SelectedItem.SmallIcon = "CAT" Or LV.SelectedItem.SmallIcon = "CAT1" Then Exit Sub
    If MsgBox("Are you sure to delete?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    'Now First Check If it's Assigned to Any Item...
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(RMID)", "ItemsRMComp", " WHERE RMID=" & Val(LV.SelectedItem.key))
    If lCount > 0 Then
        MsgBox "This Raw material is Assigned to Item(s).", vbInformation
        Exit Sub
    End If
    lCount = GetSingleLongValue("COUNT(RMID)", "VendOrderDetail", " WHERE RMID='" & LV.SelectedItem.Text & "'")
    If lCount > 0 Then
        MsgBox "This Raw material appears in PO(s).", vbInformation
        Exit Sub
    End If
    
    con.Execute "DELETE FROM RM WHERE RMID=" & Val(LV.SelectedItem.key)
    Call AddHeads
    
End Sub

Private Sub MnuEditItem_Click()
    If LV.SelectedItem.SmallIcon = "CAT" Or LV.SelectedItem.SmallIcon = "CAT1" Then
        Load frmNewRMCat
        Call frmNewRMCat.EditCat(Val(LV.SelectedItem.key))
    ElseIf LV.SelectedItem.SmallIcon = "ITM" Then
        Load frmNewRM
        Call frmNewRM.EditItem(Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 3))
    End If
End Sub

Private Sub mnuPOrd_Click()

    On Error GoTo err
    
    Dim f As New frmRMPurchOrd
    Load f
    Call f.AddPurchOrd(Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 3), 100)
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintGeneralHistory_Click()
    Dim myDTFrom As Date, myDTTo As Date
    Dim f As New frmGetDateDouble
    If f.getDate(myDTFrom, myDTTo) = False Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\GeneralMaterialPurchaseHistory.rpt")
    rpt.FormulaFields.GetItemByName("DateRange").Text = "'" & Format(myDTFrom, "dd-MMM-yyyy") & " to " & Format(myDTTo, "dd-MMM-yyyy") & "'"
    Dim frm As New frmPrevRpt
    frm.ShowReport "{VendOrders.DT}=#" & myDTFrom & "# TO #" & myDTTo & "#", rpt
    
End Sub

Private Sub mnuPrintNonZero_Click()
        
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\RMList_NonZero.rpt")
    Dim strSelection As String
    strSelection = "{VRM.QtyInStock}>0"
    If cmbGroup.ID <> "0" Then
        strSelection = strSelection & "  AND {Material.GroupID}=" & cmbGroup.ID
    End If
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub mnuPrintRateHistory_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\RMRateHistory.rpt")
    Dim strSelection As String
    If cmbGroup.ID <> "0" Then
        strSelection = "{Material.GroupID}=" & cmbGroup.ID
    Else
        strSelection = ""
    End If
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
     
End Sub

Private Sub mnuPrintRMBarcode_Click()
    Call PrintRMReports(4)
End Sub

Private Sub mnuPrintRMDetails_Click()
    Call PrintRMReports(1)
End Sub

Private Sub mnuPrintRMLable_Click()
    Call PrintRMReports(2)
End Sub

Private Sub mnuPrintRMListCrossNo_Click()
    
    Dim rpt As CRAXDDRT.Report 'New rptRawMaterialList
    Set rpt = rptApp.OpenReport(RptPath & "\RMList_CrossNo.rpt") '
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, strSQL_Reports
    
End Sub

Private Sub mnuPrintRMListCrossReference_Click()

    Dim rpt As CRAXDDRT.Report 'New rptRawMaterialList
    Set rpt = rptApp.OpenReport(RptPath & "\RMList_CrossReference.rpt") '
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, strSQL_Reports
    
End Sub

Private Sub mnuPrintRMListWithUsedIn_Click()

    Dim rpt As CRAXDDRT.Report 'New rptRawMaterialList
    Set rpt = rptApp.OpenReport(RptPath & "\RMList_UsedIn.rpt") '
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, strSQL_Reports
    
End Sub

Private Sub mnuPrintRMListWithValue_Click()

    Dim strSelection As String
    strSelection = "{VRM.InActive}=False"
    If cmbGroup.ID <> "0" Then
        strSelection = strSelection & " AND {VRM.GroupID}=" & cmbGroup.ID
    End If
    Dim rpt As CRAXDDRT.Report 'New rptRawMaterialList
    Set rpt = rptApp.OpenReport(RptPath & "\RMList_Value.rpt") '
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt

End Sub

Private Sub mnuPrintRMRackwise_Click()
            
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\RMListRackwise.rpt")
    Dim strSelection As String
    strSelection = "{VRM.QtyInStock}>0"
    If cmbGroup.ID <> "0" Then
        strSelection = strSelection & "  AND {Material.GroupID}=" & cmbGroup.ID
    End If
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub mnuPrintRMRackwiseGrouped_Click()
            
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\RMListRackwiseGrouped.rpt")
    Dim strSelection As String
    strSelection = "{VRM.QtyInStock}>0"
    If cmbGroup.ID <> "0" Then
        strSelection = strSelection & "  AND {Material.GroupID}=" & cmbGroup.ID
    End If
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt

End Sub

Private Sub mnuPrintRMRLable_Click()
    Call PrintRMReports(3)
End Sub

Private Sub mnuPrintRMSpecification_Click()
    Call PrintRMReports(0)
End Sub

Private Sub mnuPrintStockReport_Click()

    Dim fDT As New frmGetDateDouble
    Dim myDTFrom As Date, myDTTo As Date
    
    If fDT.getDate(myDTFrom, myDTTo) = False Then Exit Sub
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\MaterialStatusPeriodwise.rpt")
    
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue myDTFrom
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue myDTTo
    rpt.ParameterFields.GetItemByName("@GroupID").SetCurrentValue CLng(cmbGroup.ID)
    rpt.FormulaFields.GetItemByName("DateRange").Text = "'" & Format(myDTFrom, "dd-MMM-yyyy") & " to " & Format(myDTTo, "dd-MMM-yyyy") & "'"
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, , False
    
End Sub

Private Sub mnuPurchases_Click()
    frmItemPurchase.Show
    frmItemPurchase.SetFocus
End Sub

Private Sub mnuStockHistory_Click()

    Load frmStockLedger
    frmStockLedger.cmbCat.ID = "0"
    frmStockLedger.cmbMaterial.ID = LV.SelectedItem.Text
    frmStockLedger.cmbRange.ListIndex = 3
    
    frmStockLedger.Show
    
End Sub

Private Sub mnuNewItem_Click()

    With frmNewRM
        Load frmNewRM
        Call .AddNew(bSampling)
    End With
    Call mnuRefresh_Click
    
End Sub

Private Sub mnunewcat_Click()
    Dim f As New frmItemGroups
    Load f
    Call f.ShowForm("RMGroups", "RM Groups", "RM Name")
End Sub

Private Sub mnuParent_Click()
   If MsgBox("Are You Sure To Make This Account Parent", vbQuestion + vbYesNo) = vbNo Then Exit Sub
   con.Execute "update accounts set Parent=True Where AccNo='" & LV.SelectedItem & "'"
   Call mnuRefresh_Click
End Sub

Private Sub mnuPrint_Click()

    Dim strSelection As String
    strSelection = "{VRM.InActive}=False"
    If cmbGroup.ID <> "0" Then
        strSelection = strSelection & " AND {VRM.GroupID}=" & cmbGroup.ID
    End If
    
    Call ShowRMListrpt(strSelection)
    
End Sub

Private Sub mnuPrintBal_Click()

    Screen.MousePointer = vbHourglass
    Load frmPrevRpt
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptDMRInternal.rpt") '
    frmPrevRpt.ShowReport "{VrptDMR.RMID}='" & LV.SelectedItem.key & "'", rpt
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub mnuRefresh_Click()
   Screen.MousePointer = vbHourglass
   'rsAccounts.Close
   'Set rsAccounts = Nothing
   'Call Form_Load
   Call AddHeads
   'Call FillItems
   Screen.MousePointer = vbDefault
End Sub


Private Sub CheckMenu()

    If LV.ListItems.count = 0 Then
        mnuNewItem.Enabled = True
        mnuEditItem.Enabled = False
        mnuSales.Enabled = False
        mnuPurchases.Enabled = False
        mnuStockHistory.Enabled = False
        mnuPrint.Enabled = False
        mnuParent.Enabled = False
        mnuNewCat.Enabled = True
    Else
        mnuNewItem.Enabled = True
        mnuEditItem.Caption = IIf(LV.SelectedItem.SmallIcon Like "ITM*", "Edit This Raw Material", "Edit This Catagory")
        mnuEditItem.Enabled = True
        mnuSales.Enabled = True
        mnuPurchases.Enabled = True
        mnuStockHistory.Enabled = LV.SelectedItem.SmallIcon Like "ITM*"
     
        If LV.SelectedItem.Tag = "Inactive" Then
            mnuEditItem.Enabled = False
            mnuSales.Enabled = False
            mnuPurchases.Enabled = False
            mnuStockHistory.Enabled = False
            mnuParent.Enabled = False
            mnuActive.Caption = "Make This Raw Material A&ctive."
        Else
            mnuActive.Caption = "Make This Raw Material Ina&ctive."
        End If
        mnuActive.Tag = LV.SelectedItem.Tag
        
        mnuPrint.Enabled = LV.SelectedItem.SmallIcon Like "ITM*"
        mnuPrintBal.Enabled = LV.SelectedItem.SmallIcon Like "ITM*"
        
        mnuActive.Enabled = LV.SelectedItem.SmallIcon Like "ITM*"
    End If



End Sub

Private Sub mnuShowAll_Click()
    mnuShowAll.Checked = Not mnuShowAll.Checked
    ShowAll = mnuShowAll.Checked
    Call mnuRefresh_Click
End Sub

Private Sub mnuSales_Click()
   'frmItemSale.Show
   'frmItemSale.SetFocus
End Sub



Private Sub txtAccTitle_GotFocus()
txtAccTitle.SelStart = 0
txtAccTitle.SelLength = Len(txtAccTitle)
End Sub

Private Sub txtAccTitle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then If Not txtBal.Enabled Then Call UpdateAccount

End Sub


Private Sub txtBal_GotFocus()
txtBal.SelStart = 0
txtBal.SelLength = Len(Bal)
End Sub

Private Sub txtBal_KeyPress(KeyAscii As Integer)
   KeyAscii = OnlyNumber(KeyAscii, True)
End Sub

Private Sub txtBal_LostFocus()
 txtBal = Format(Val(txtBal), "###0.00")
 txtBal.Tag = LV.SelectedItem.ListSubItems(2).Tag
 Call cmbCrDr_Click
End Sub



Private Function MakeAccountInActive() As Boolean
MakeAccountInActive = False
 With LV.SelectedItem
   
   'If Val(.ListSubItems(6)) > 0 Then
   '     MsgBox "This Item Cannot Be Made Inactive." & vbNewLine & _
   '        "To Make An Item Inactive Quantity In Hand Must Be Zero.", vbInformation
   '      Exit Function
   'ElseIf .SmallIcon = "PAR" Then
   '    If MsgBox("This Is A Master Account, All Sub Accounts Will Also Become Inactive." & vbNewLine & _
   '        "Are You Sure to Make This Account Inactive", vbYesNo + vbQuestion) = vbNo Then Exit Function
   'Else
      If MsgBox("Are You Sure to Make This Raw Material Inactive", vbYesNo + vbQuestion) = vbNo Then Exit Function
  ' End If
   
   con.Execute "Update RM Set InActive=1 Where RMID=" & Val(.key)
  
   MakeAccountInActive = True
End With

End Function


Private Function MakeAccountActive() As Boolean

MakeAccountActive = False
Dim AlsoChild As String

 With LV.SelectedItem
   con.Execute "Update RM Set InActive=0 Where RMID=" & Val(.key)
   MakeAccountActive = True
End With




End Function



Private Function UpdateAccount() As Boolean
 If txtAccTitle.Text = "" Then
   MsgBox "Invalid Account Title", vbInformation
   Exit Function
 End If
 
Dim Diff As Double

Screen.MousePointer = vbHourglass
Diff = Val(cmbCrDr.Tag) - Val(txtBal.Tag)
   con.Execute "Update Accounts Set Acctitle='" & txtAccTitle & "',OpenBal=" & Val(cmbCrDr.Tag) & _
               ", Balance=Balance+" & Diff & " Where AccNo='" & LV.SelectedItem.Text & "'"
               
   'Adjust The Ledger
   Call adjustLedger(LV.SelectedItem.Text, #1/1/1900#, 0, Diff, True)
   
   'Sleep (500)
   Call mnuRefresh_Click
Screen.MousePointer = vbDefault
   
End Function



Public Sub AddHeads(Optional strOrderCond As String)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim LColor As Long, lRed As Long, lGreen As Long, lBlue As Long
    
    Dim ItemOrdCond As String
    
    If strOrderCond = "" Then
        strOrderCond = " ORDER BY CatID "
        ItemOrdCond = ""
    Else
        ItemOrdCond = " ORDER BY RMName"
    End If
     
    cmbItem.DropDownOffLine = True
    cmbItem.ClearVals
    If cmbGroup.ID = "0" Then
        cmbItem.AddVals con, "'{' + RMID1 + '} ' + RMName + ' {' + CAST(ISNULL(RMSize,'') AS VARCHAR(50)) + ' '+ ISNULL(SizeUnit,'') + '}'", "RM", "RMID", " WHERE Sampling=" & Abs(bSampling) & " ORDER BY RMID1"
    Else
        cmbItem.AddVals con, "'{' + RMID1 + '} ' + RMName + ' {' + CAST(ISNULL(RMSize,'') AS VARCHAR(50)) + ' '+ ISNULL(SizeUnit,'') + '}'", "RM", "RMID", " WHERE GroupID=" & cmbGroup.ID & " AND Sampling=" & Abs(bSampling) & " ORDER BY RMID1"
    End If
    
    Call FillItems(ItemOrdCond)
    Exit Sub
    With rs
        .Open "SELECT * FROM RMCatagories " & " ORDER BY CAST(CatID AS NUMERIC) ", con, adOpenForwardOnly, adLockReadOnly
        Set LV.SmallIcons = ImageList1
        If LV.ListItems.count > 0 Then PrevIndex = LV.SelectedItem.Index
    
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            '
            'Randomize
            'lRed = Int(Rnd * (255 - (0 - 1)) + (0))
            'lGreen = Int(Rnd * (255 - (0 - 1)) + (0))
            'lBlue = Int(Rnd * (255 - (0 - 1)) + (0))
            'lColor = RGB(lRed, lGreen, lBlue)
            'itm.SubItems(6) = Int(Rnd * (iMax - (iMin - 1)) + (iMin))
            LColor = vbBlack
            Set ITM = LV.ListItems.add(, ![CatID] & "'", ![Description], , "CAT")
            ITM.Bold = True
            ITM.ForeColor = LColor
            ITM.ListSubItems.add
            'ITM.ListSubItems.Add(, , ![Description] & "").Bold = True
            'ITM.ListSubItems(1).ForeColor = lColor
            'ITM.ListSubItems.Add(, , Format(Val(![Qty] & ""), "#0")).Bold = True
            'Call FillItems(![CatID], ItemOrdCond)
            .MoveNEXT
        Next
        .Close
    End With
    
    Set rs = Nothing
    'Call FillItems
    Exit Sub
    
err:
    MsgBox err.Description
    
End Sub


Private Function GetlvlStr(CatID As String) As String
    GetlvlStr = Space(5 * SubStrCount(CatID, "-"))
    If Len(GetlvlStr) > 0 Then GetlvlStr = GetlvlStr & "» "
End Function

Public Sub ShowMe(p_bSampling As Boolean)
    bSampling = p_bSampling
    cmbGroup.ID = "0"
    Me.Show
End Sub

Private Sub chkShowPicture_Click()

    If chkShowPicture.Value = vbChecked Then
        Call LV_ItemClick(LV.SelectedItem)
    Else
        PicProduct.Visible = False
    End If
    
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)

    '''
    If LV.SelectedItem.SmallIcon <> "ITM" Then
        PicProduct.Visible = False
        Exit Sub
    End If
    If chkShowPicture.Value = vbChecked Then
        Dim rs As New ADODB.Recordset
        With rs
            .Open "SELECT Pic FROM RM WHERE RMID=" & Val(LV.SelectedItem.key), con, adOpenForwardOnly, adLockReadOnly
            If Not IsNull(![Pic]) Then
                Set ImgProduct.DataSource = rs
                ImgProduct.DataField = "Pic"
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

Private Sub PrintRMReports(iReportType As Integer)

    Dim rpt As CRAXDDRT.Report
    
    If iReportType = 0 Then
        Set rpt = rptApp.OpenReport(RptPath & "\RM_Specification.rpt")
    ElseIf iReportType = 1 Then
        Set rpt = rptApp.OpenReport(RptPath & "\RM_Details.rpt")
    ElseIf iReportType = 2 Then
        Set rpt = rptApp.OpenReport(RptPath & "\RM_Lable.rpt")
    ElseIf iReportType = 3 Then
        Set rpt = rptApp.OpenReport(RptPath & "\RM_R_Lable.rpt")
    ElseIf iReportType = 4 Then
        Set rpt = rptApp.OpenReport(RptPath & "\RM_Barcode.rpt")
    End If
    
    Dim f As New frmPrevRpt
    Dim strRMID As String
    strRMID = Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 3)
    
    Dim strSelection As String
    strSelection = "{VRM.RMID}=" & strRMID
    f.ShowReport strSelection, rpt
    
End Sub
