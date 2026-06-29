VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmItems 
   AutoRedraw      =   -1  'True
   ClientHeight    =   9495
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   14430
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9495
   ScaleWidth      =   14430
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtTotalStockQty 
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
      Left            =   6555
      TabIndex        =   27
      Top             =   8670
      Width           =   1755
   End
   Begin ComboList.Usercontrol1 cmbGroup 
      Height          =   285
      Left            =   75
      TabIndex        =   7
      Top             =   270
      Width           =   4305
      _ExtentX        =   7594
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
   Begin ComboList.Usercontrol1 cmbItemCatagory 
      Height          =   285
      Left            =   4395
      TabIndex        =   23
      Top             =   270
      Width           =   4980
      _ExtentX        =   8784
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
   Begin ComboList.Usercontrol1 cmbItem 
      Height          =   285
      Left            =   4395
      TabIndex        =   9
      Top             =   825
      Width           =   4980
      _ExtentX        =   8784
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
   Begin ComboList.Usercontrol1 cmbAdditionalGroup 
      Height          =   285
      Left            =   75
      TabIndex        =   25
      Top             =   840
      Width           =   4305
      _ExtentX        =   7594
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
   Begin VB.Frame FraCatalog 
      Height          =   480
      Left            =   10620
      TabIndex        =   20
      Top             =   75
      Width           =   3750
      Begin VB.TextBox txtRef 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2310
         TabIndex        =   22
         Top             =   135
         Width           =   1365
      End
      Begin MSForms.ComboBox cmbCatalog 
         Height          =   300
         Left            =   75
         TabIndex        =   21
         Top             =   135
         Width           =   2205
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3889;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2175
      Top             =   3495
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":0000
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":0454
            Key             =   "ORD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":08AC
            Key             =   "UNLOCK"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":1636
            Key             =   "CAT1"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":1A8A
            Key             =   "ITMIN"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":1EDE
            Key             =   "LOCK"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":2C68
            Key             =   "CAT"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":30BC
            Key             =   "OK"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItems.frx":3514
            Key             =   "TYPE"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox ImgLoading 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3600
      Left            =   1830
      ScaleHeight     =   3540
      ScaleWidth      =   7845
      TabIndex        =   18
      Top             =   2040
      Visible         =   0   'False
      Width           =   7905
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Please wait While Loading Data"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1860
         Left            =   1350
         TabIndex        =   19
         Top             =   960
         Width           =   5310
      End
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
      Left            =   3135
      TabIndex        =   16
      Top             =   8010
      Width           =   1755
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
      Top             =   8025
      Width           =   1560
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
      Left            =   8940
      ScaleHeight     =   2835
      ScaleWidth      =   2760
      TabIndex        =   14
      Top             =   1590
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
      Left            =   6870
      TabIndex        =   13
      Top             =   8010
      Width           =   1755
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   480
      Top             =   2145
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
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "frmItems.frx":3968
         Left            =   5250
         List            =   "frmItems.frx":3972
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   -45
         Width           =   600
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   6765
      Left            =   60
      TabIndex        =   3
      Top             =   1155
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   11933
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
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item No."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Item Group"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   4604
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Size"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Fin. Weight"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Stock"
         Object.Width           =   1931
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Unit"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Cost"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Value"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "FOB SKT"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "FOB SKT Top"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Type"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Tip Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Fin. Quality"
         Object.Width           =   1879
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "ReOrder"
         Object.Width           =   1508
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "PU"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Label lblTotalStockQty 
      AutoSize        =   -1  'True
      Caption         =   "Total Stock (Pcs) :"
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
      Left            =   4950
      TabIndex        =   28
      Top             =   8700
      Width           =   1530
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFBDCB&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Additional Group"
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
      Index           =   2
      Left            =   75
      TabIndex        =   26
      Top             =   600
      Width           =   4305
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFBDCB&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Item Catagory"
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
      Left            =   4395
      TabIndex        =   24
      Top             =   30
      Width           =   4980
   End
   Begin VB.Label lblActiveArticles 
      AutoSize        =   -1  'True
      Caption         =   "Active Articles :"
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
      Left            =   1770
      TabIndex        =   17
      Top             =   8040
      Width           =   1320
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
      Left            =   5790
      TabIndex        =   12
      Top             =   8040
      Width           =   1035
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   840
      Left            =   9495
      TabIndex        =   11
      Top             =   120
      Width           =   975
      Caption         =   "Refresh"
      PicturePosition =   327683
      Size            =   "1720;1482"
      Accelerator     =   82
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
      Left            =   4395
      TabIndex        =   10
      Top             =   600
      Width           =   4980
      BackColor       =   11517387
      Caption         =   " Item Code"
      Size            =   "8784;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFBDCB&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Group"
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
      Left            =   75
      TabIndex        =   8
      Top             =   30
      Width           =   4305
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
         Caption         =   "&Show Inactive Items"
      End
      Begin VB.Menu mnuStockHistory 
         Caption         =   "View &Stock History For Last 30 Days"
      End
      Begin VB.Menu dash0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewItem 
         Caption         =   "&Add New Item"
      End
      Begin VB.Menu mnuNewCat 
         Caption         =   "Add New Catagory"
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Edit Item"
      End
      Begin VB.Menu mnuDeleteItem 
         Caption         =   "Delete Selected Item"
      End
      Begin VB.Menu dash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyItem 
         Caption         =   "&Copy Item"
      End
      Begin VB.Menu mnucopydash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActive 
         Caption         =   "Make This Item Ina&ctive"
      End
      Begin VB.Menu mnuParent 
         Caption         =   "Make This Account Paren&t"
         Visible         =   0   'False
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAssignAllProcesses 
         Caption         =   "Assign All Processes"
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
         Caption         =   "&Print Report"
      End
      Begin VB.Menu mnuPrintTechDrawing 
         Caption         =   "Print Tech. Drawing"
      End
      Begin VB.Menu mnuPrintBal 
         Caption         =   "Print Report(Internal)"
      End
      Begin VB.Menu mnuPrintdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintCatalog 
         Caption         =   "Print Catalog"
      End
      Begin VB.Menu mnuPrintSpecSheet 
         Caption         =   "Print Specification Sheet"
      End
      Begin VB.Menu mnuPrintStorewiseStatus 
         Caption         =   "Print Storewise Status"
      End
      Begin VB.Menu mnuLockThisItem 
         Caption         =   "Lock This Item"
      End
      Begin VB.Menu mnuResetProcessGroups 
         Caption         =   "Reset/Re-Assign Process Groups"
      End
      Begin VB.Menu mnuItemsWithMissingParameters 
         Caption         =   "Print Items With Missing Parameters"
      End
      Begin VB.Menu mnuPrintNewItems 
         Caption         =   "Print New Items"
      End
      Begin VB.Menu mnuItemsNoUsed 
         Caption         =   "Items Not Used in Customer Orders"
      End
      Begin VB.Menu mnuPrintECommerceStockReport 
         Caption         =   "Print Stock Report (E-Commerce)"
      End
      Begin VB.Menu mnuPrintItemsBelowMinLevel 
         Caption         =   "Below Min. Level"
      End
      Begin VB.Menu mnuPrintItemsAboveMaxLevel 
         Caption         =   "Above Max. Level"
      End
      Begin VB.Menu mnuPrintCombineStockReport 
         Caption         =   "Print Combine Stock Report"
      End
      Begin VB.Menu mnuPrintCombineReportWithLocation 
         Caption         =   "Print Combine Report With Location"
      End
      Begin VB.Menu mnuPrintSummaryStock 
         Caption         =   "Print Summary Stock"
      End
      Begin VB.Menu mnuPrintSFStock 
         Caption         =   "Print S.F Stock"
      End
   End
End
Attribute VB_Name = "frmItems"
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
Dim iCatalogType As Integer
Dim bSample As Boolean
Dim bShow_Values_Company_Catalog As Boolean

Public Sub ShowMe(Optional p_iCatalogType As Long = 0, Optional m_bSample As Boolean = False)

    'Zero for all other customers than QEL, which means show all
    '1 show only Surgical
    '2 show only Dental
    '3 show all other than surgical & Dental
    bSample = m_bSample
    iCatalogType = p_iCatalogType
    
    
    Call cmbItem.AddVals(con, "ItemID + ' ' + ItemName", "Items", "ItemID", " WHERE InActive=0 AND Sample=" & Abs(bSample) & " ORDER BY ItemName")
    
    Me.Show
 
    cmbGroup.ID = "0"
         
End Sub

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

Private Sub FillItems(CatID As Integer, strOrderCond As String, Optional startIndx As Long)
    
    'LV.Sorted = False
    Set LV.SmallIcons = ImageList1
    If startIndx = 0 Then
        startIndx = LV.ListItems.count
    End If
    Dim rs As New ADODB.Recordset
    Dim Cnd As String, ITM As ListItem
    
    If bSample Then
        Cnd = " WHERE Sample=1"
    Else
        Cnd = " WHERE Sample=0"
    End If
    
    If Not ShowAll Then
        Cnd = Cnd & " AND InActive=0 AND CatID=" & CatID
    Else
        Cnd = Cnd & " AND CatID=" & CatID
    End If
    
    If cmbGroup.ID <> "0" Then
        Cnd = Cnd & " AND GroupID=" & cmbGroup.ID & ""
    End If
    If cmbAdditionalGroup.ID <> "0" Then
        Cnd = Cnd & " AND MainGroupID=" & cmbAdditionalGroup.ID & ""
    End If
    'LV.ListItems.Clear
    
    
    If strOrderCond = "" Then
        If strCompany = "QEL" Or strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
            strOrderCond = " Order By ItemID"  '" Order By cast(CatID as numeric)"
        Else
            strOrderCond = " Order By Category,ItemGroup"  '" Order By cast(CatID as numeric)"
        End If
    End If
    
    With rs
        .CursorLocation = adUseClient
        .Open "SELECT GrpColor,ItemID,InActive,ItemGroup,ItemName,ItemSize,SizeUnit,Unit,FinQuality,FinishedWeight,ReorderPoint,FixedPackingUnit,InHand,PriceForCost,Type,TipSize,ExWorks,ExWorksTop FROM VItems " & Cnd & strOrderCond, con_ServerSide, adOpenForwardOnly, adLockReadOnly
        
        For i = 0 To .RecordCount - 1
            
            If ![Inactive] Then
               clr = &H707070
            Else
               clr = Val(![GrpColor] & "")
            End If
             
            'Set itm = LV.ListItems.Add(, ![ITemID], ![ITemID], , IIf(![Inactive] = 0, "ITM", "ITMIN"))
            
            Set ITM = LV.ListItems.add(startIndx + 1, ![ITemID] & "'", "     »  " & ![ITemID], , IIf(![Inactive] = 0, "ITM", "ITMIN"))
            
            'itm.ListSubItems.Add(, , Space(5) & GetlvlStr(![CatID]) & ![ItemName] & "", , ![ItemName]).ForeColor = clr
            ITM.ListSubItems.add(, , ![ItemGroup] & "", , ![ItemGroup] & "").ForeColor = Val(![GrpColor] & "")
            ITM.ListSubItems(1).Bold = True
            'itm.ListSubItems.Add(, , "     »  " & ![ItemName] & " " & ![ItemSize] & " " & ![SizeUnit] & " " & IIf(![Type] = "N.A", "", ![Type]) & "", , ![ItemName]).ForeColor = clr
            ITM.ListSubItems.add(, , "»  " & ![ItemName] & "", , ![ItemName] & "").ForeColor = clr
            ITM.ListSubItems(2).ReportIcon = "UNLOCK"
            ITM.ListSubItems.add(, , ![ItemSize] & " " & ![SizeUnit]).ForeColor = clr
            
            
            ITM.ListSubItems.add(, , Val(![FinishedWeight] & "")).ForeColor = clr
            
                
            ITM.ListSubItems.add(, , Val(!InHand & "")).ForeColor = clr
            ITM.ListSubItems.add(, , ![Unit] & "").ForeColor = clr
            ITM.ListSubItems.add(, , Val(!PriceForCost & "")).ForeColor = clr
            ITM.ListSubItems.add(, , Val(!InHand & "") * Val(!PriceForCost & "")).ForeColor = clr
            
            ITM.ListSubItems.add(, , Val(!ExWorks & "")).ForeColor = clr
            ITM.ListSubItems.add(, , Val(!ExWorksTop & "")).ForeColor = clr
            
            
            ITM.ListSubItems.add(, , ![Type] & "").ForeColor = clr
            ITM.ListSubItems.add(, , ![TipSize] & "").ForeColor = clr
            ITM.ListSubItems.add(, , ![FinQuality] & "").ForeColor = clr
            ITM.ListSubItems.add(, , ![ReorderPoint] & "").ForeColor = clr
            ITM.ListSubItems.add(, , Val(!FixedPackingUnit & "")).ForeColor = clr
            
            If ![Inactive] = 0 Then ITM.Tag = "Active" Else ITM.Tag = "Inactive"
               
            .MoveNEXT
            startIndx = startIndx + 1
        Next
        .Close
             
        On Error Resume Next
        .Open "SELECT ItemCode FROM ItemsLockedForEditing", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            LV.ListItems(!ItemCode & "'").ListSubItems(2).ReportIcon = "LOCK"
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    LV.Visible = True
    'LV.Sorted = True
    If PrevIndex > 0 And PrevIndex <= LV.ListItems.count Then LV.ListItems(PrevIndex).Selected = True
    
    LV.SelectedItem.EnsureVisible
    Screen.MousePointer = vbDefault

End Sub

Private Sub chkShowPicture_Click()

    If chkShowPicture.Value = vbChecked Then
        Call LV_ItemClick(LV.SelectedItem)
    Else
        PicProduct.Visible = False
    End If
    
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

    cmbItem.ClearVals
    Dim Cnd As String
    If bSample Then
        Cnd = " WHERE Sample=1"
    Else
        Cnd = " WHERE Sample=0"
    End If
    
'    If Not ShowAll Then
'        Cnd = Cnd & " AND InActive=0 AND CatID=" & CatID
'    Else
'        Cnd = Cnd & " AND CatID=" & CatID
'    End If
    
    If cmbGroup.ID <> "0" Then
        Cnd = Cnd & " AND GroupID=" & cmbGroup.ID & ""
    End If
    
    If mnuShowAll.Checked = False Then
        Cnd = Cnd & " AND InActive=0"
    End If
    
    Call cmbItem.AddVals(con, "ItemID + ' ' + ItemName", "Items", "ItemID", Cnd & " ORDER BY ItemName")
    Call AddHeads
End Sub

Private Sub cmbItem_matched()
    'cmbItemName.ID = cmbItem.ID
    Call GoToSelected
End Sub

Private Sub cmbItemName_matched()
'    cmbItem.ID = cmbItemName.ID
End Sub

Private Sub GoToSelected()
    
    LV.ListItems(cmbItem.ID & "'").EnsureVisible
    'LV.FindItem "     »  " & cmbitem.ID
    LV.ListItems(cmbItem.ID & "'").Selected = True
    LV.SetFocus
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
'    Debug.Print Now
    Screen.MousePointer = vbHourglass
   'rsAccounts.Close
   'Set rsAccounts = Nothing
   'Call Form_Load
   Call AddHeads
   'Call FillItems
   Screen.MousePointer = vbDefault
'   Debug.Print Now
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
        
    If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
        mnuPrintStorewiseStatus.Visible = True
    Else
        mnuPrintStorewiseStatus.Visible = False
    End If
    
    If strCompany = "Dr-Frgz" Then
        mnuPrintECommerceStockReport.Visible = True
    Else
        mnuPrintECommerceStockReport.Visible = False
    End If
    
    Dim Condition As String
    If ShowAll Then
        Condition = ""
    Else
        Condition = " Where Active "
    End If
    'Set LV.SmallIcons = ImageList1
    'rsAccounts.Open "Select * from Accounts " & Condition & " Order By Accno", Con, adOpenForwardOnly
    'Call RefreshList
    cmbGroup.ListHeight = 4000
    Call cmbGroup.AddItem("<All Groups>", "0")
    Call cmbGroup.AddVals(con, "Description", "ItemGroups", "ID")
    
    cmbItemCatagory.ListHeight = 4000
    cmbItemCatagory.AddItem "<All>", "0"
    cmbItemCatagory.AddVals con, "Description", "ItemCatagories", "CatID", " ORDER BY ItemCatagories.Description"
    cmbItemCatagory.ID = "0"
    
    cmbAdditionalGroup.ListHeight = 4000
    cmbAdditionalGroup.AddItem "<All>", "0"
    cmbAdditionalGroup.AddVals con, "MainGroupName", "ItemGroupsMain", "MainGroupID", " ORDER BY ItemGroupsMain.MainGroupName"
    cmbAdditionalGroup.ID = "0"
    
    cmbItem.DropDownOffLine = True
    cmbItem.ListHeight = 2200
    
'    Call cmbItemName.AddVals(con, "IsNull(ItemName,'') + ' ' + IsNull(ItemSize,'') + ' ' + IsNull(SizeUnit,'') + ' ' + IsNull(Type1,'') + ' ' + IsNull(FinQuality,'') ", "VItems", "ItemID", " WHERE InActive=0 Order By ItemName")
    
'    Call AddHeads
    If strCompany = "Tecno" Or strCompany = "Dr-Frgz" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
        mnuPrintSpecSheet.Visible = True
    Else
        mnuPrintSpecSheet.Visible = False
    End If
    
    Call AddToCombo(cmbCatalog, "CatalogName", "ForeignCatalog", " ORDER BY CatalogName", , "CatalID")
    bSample = False
    
    bShow_Values_Company_Catalog = GetSingleBooleanValue("Show_Values_Company_Catalog", "Users", " WHERE UserName='" & CurrentUserName & "'")
    
    If bShow_Values_Company_Catalog = False Then
        txtTotalValue.Visible = False
        lblTotalValue.Visible = False
        LV.ColumnHeaders(10).Width = 0
        LV.ColumnHeaders(13).Width = 0
    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    'LV.Move 50, 50, Me.ScaleWidth - 100, Me.ScaleHeight - (cmdClose.Height + 100)
    LV.Move 50, 1200, Me.ScaleWidth - 100, Me.ScaleHeight - (cmdClose.Height + 100 + cmbAdditionalGroup.Height + cmbAdditionalGroup.Top)
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, LV.Top + LV.Height + 50
    'cmdPrint.Move cmdClose.Left - cmdPrint.Width - 50, cmdClose.Top
    cmdOpt.Move cmdClose.Left - cmdOpt.Width - 50, cmdClose.Top
    chkShowPicture.Move LV.Left, cmdClose.Top
     
    
    
    txtTotalValue.Move cmdOpt.Left - txtTotalValue.Width - 50, cmdClose.Top + 20
    lblTotalValue.Move txtTotalValue.Left - lblTotalValue.Width - 50, cmdClose.Top + 30
    
    txtTotalStockQty.Move lblTotalValue.Left - txtTotalStockQty.Width - 50, txtTotalValue.Top
    lblTotalStockQty.Move txtTotalStockQty.Left - lblTotalStockQty.Width - 50, lblTotalValue.Top
    
    lblActiveArticles.Move chkShowPicture.Left + chkShowPicture.Width + 50, lblTotalStockQty.Top
    txtActiveArticles.Move lblActiveArticles.Left + lblActiveArticles.Width, txtTotalValue.Top
    
    ImgLoading.Move (Me.ScaleWidth - ImgLoading.Width) / 2, (Me.ScaleHeight - ImgLoading.Height) / 2
    
End Sub

Private Sub Form_Unload(cancel As Integer)
    'rsAccounts.Close
    Set rsAccounts = Nothing
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Select Case ColumnHeader.Index
        Case 1
            Call AddHeads(" Order By ItemID")
        Case 2
            Call AddHeads(" Order By ItemGroup")
        Case 3
            Call AddHeads(" Order By Description")
        Case 5
            Call AddHeads(" Order By FinQuality")
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
            LV.SelectedItem.SmallIcon = "CAT"
            Call FillItems(Val(LV.SelectedItem.key), " Order By ItemName", LV.SelectedItem.Index)
'            mnuRefresh_Click
        End If
        
    ElseIf LV.SelectedItem.SmallIcon = "ITM" Then
'        If GetSingleLongValue("COUNT(ItemCode)", "ItemsLockedForEditing", " WHERE ItemCode='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "'") > 0 Then   'Item Is Locked...
'            If ValidateAdminPassword = False Then
'                Exit Sub
'            End If
'        End If
        Load frmNewItem
        Call frmNewItem.EditItem(LV.SelectedItem.key)
'        Call mnuRefresh_Click
    End If
     
    Exit Sub

End Sub

Private Sub LV_GotFocus()
   If PicEdit.Visible Then Call Form_KeyPress(27)
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
            .Open "SELECT ItemPic FROM Items WHERE ItemID='" & Left(Item.key, Len(Item.key) - 1) & "'", con, adOpenForwardOnly, adLockReadOnly
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

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)

    If LV.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If LV.SelectedItem.SmallIcon <> "ITM" Then Exit Sub
        If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo) = vbYes Then
            Dim lOrderCount As Long
            lOrderCount = GetSingleLongValue("COUNT(CompItemCode)", "FOrderItems", " WHERE CompItemCode='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "'")
            If lOrderCount > 0 Then
                MsgBox "Can't Delete this Item, Orders are booked against this.", vbInformation
                Exit Sub
            Else
                If GetSingleLongValue("COUNT(ItemCode)", "ItemsLockedForEditing", " WHERE ItemCode='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "'") > 0 Then
                    If getDBPassword("LockedItem") = False Then
                        Exit Sub
                    End If
                End If
                con.Execute "DELETE FROM Items WHERE ItemID='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "'"
            End If
        End If
    End If
    
End Sub

Private Sub LV_KeyPress(KeyAscii As Integer)

    If LV.ListItems.count = 0 Then Exit Sub
    If KeyAscii = 13 Then
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

Private Sub mnuActive_Click()
    
    If getDBPassword("InactiveItem") = False Then Exit Sub
     
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

Private Sub mnuAssignAllProcesses_Click()

    On Error GoTo err
    If MsgBox("Are you sure to assign all Processes?", vbInformation) = vbNo Then Exit Sub
    If LV.SelectedItem.SmallIcon <> "ITM" Then Exit Sub
    
    Call StartTrans(con)
    
    con.Execute "DELETE FROM ItemProcesses WHERE ItemID='" & LV.SelectedItem.key
    con.Execute "INSERT INTO ItemProcesses SELECT ProcessID,'" & LV.SelectedItem.key & ",0,ProcessID FROM Processes"
    
    con.CommitTrans
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuCopyItem_Click()
    If LV.SelectedItem.SmallIcon = "CAT" Or LV.SelectedItem.SmallIcon = "CAT1" Then
        Exit Sub
     ElseIf LV.SelectedItem.SmallIcon = "ITM" Then
        Load frmNewItem
        Call frmNewItem.CopyItem(LV.SelectedItem.key)
        'Call mnuRefresh_Click
    End If
End Sub

Private Sub mnuDeleteItem_Click()

    On Error GoTo err
    
    If MsgBox("Are you sure to delete this Article & All Transactions of it.", vbQuestion + vbYesNo) = vbYes Then
    
        If GetSingleLongValue("COUNT(ItemCode)", "ItemsLockedForEditing", " WHERE ItemCode='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "'") > 0 Then
            If getDBPassword("LockedItem") = False Then
                Exit Sub
            End If
        End If
        
        If ValidateAdminPassword Then
            Call StartTrans(con)
            Dim strItemID As String
            strItemID = Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1)
            con.Execute "DELETE FROM FProformaOrders WHERE OrderEntryID IN(SELECT ID FROM FOrderItems WHERE CompItemCode='" & strItemID & "')"
            con.Execute "DELETE FROM Items WHERE ItemID='" & strItemID & "'"
            con.CommitTrans
        End If
        
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub MnuEditItem_Click()
    If LV.SelectedItem.SmallIcon = "CAT" Or LV.SelectedItem.SmallIcon = "CAT1" Then
        Load frmNewItemCat
        Call frmNewItemCat.EditCat(Val(LV.SelectedItem.key))
'        Call mnuRefresh_Click
     ElseIf LV.SelectedItem.SmallIcon = "ITM" Then
        Load frmNewItem
        Call frmNewItem.EditItem(LV.SelectedItem.key)
'        Call mnuRefresh_Click
    End If
End Sub

Private Sub mnuItemsNoUsed_Click()
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\CompanyCatalogItemsNotUsed.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VItems.InActive}=FALSE", rpt
End Sub

Private Sub mnuItemsWithMissingParameters_Click()
    frmPrintMissingParametersReport.Show 1
End Sub

Private Sub mnuLockThisItem_Click()

    If mnuLockThisItem.Tag = "0" Then
        con.Execute "INSERT INTO ItemsLockedForEditing(ItemCode) VALUES('" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "')"
    Else
        If getDBPassword("LockedItem") Then
            con.Execute "DELETE FROM ItemsLockedForEditing WHERE ItemCode='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "'"
        End If
    End If
    
End Sub

Private Sub mnuPrintCatalog_Click()
    'Screen.MousePointer = vbHourglass
    'Load frmPrevRpt
    'frmPrevRpt.ShowReport "", rptCompanyCatalog
    'Screen.MousePointer = vbDefault
    If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
        frmPrintCatalog_Frgz.Show 1
    Else
        frmPrintCatalog.Show 1
    End If
End Sub

Private Sub mnuPrintCombineReportWithLocation_Click()
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\CombineStockReport_WithLocation.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VItems_StockReport.InActive}=FALSE", rpt
End Sub

Private Sub mnuPrintCombineStockReport_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\CombineStockReport.rpt")
    Dim f As New frmPrevRpt
    
    Dim strSelection As String
    strSelection = "{VItems_StockReport.InActive}=FALSE"
        
    If Not (LV.SelectedItem Is Nothing) Then
        If LV.SelectedItem.Bold Then
            strSelection = strSelection & " AND {VItems_StockReport.CatID}=" & Val(LV.SelectedItem.key)
        End If
    End If
    
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub mnuPrintECommerceStockReport_Click()
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ECommerceStockReport.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VItemWithStocks.AvailableForECommerce}=TRUE", rpt
End Sub

Private Sub mnuPrintItemsAboveMaxLevel_Click()
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\CompanyCatalog_MaxLevel.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VItems.InActive}=FALSE AND {VItems.InHand}>{VItems.MaxLevel} AND {VItems.MaxLevel}>0", rpt
End Sub

Private Sub mnuPrintItemsBelowMinLevel_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\CompanyCatalog_MinLevel.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VItems.InActive}=FALSE AND {VItems.InHand}<{VItems.MinLevel}", rpt
    
End Sub

Private Sub mnuPrintNewItems_Click()

    Dim myDTFrom As Date, myDTTo As Date
    Dim f As New frmGetDateDouble
    Load f
    If f.getDate(myDTFrom, myDTTo) = False Then Exit Sub
       
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptCompanyCatalog.rpt")
    rpt.FormulaFields.GetItemByName("ShowLocationWiseStock").Text = False
    
    Dim strSelection As String
    strSelection = "{VItems.CreateDT}=#" & myDTFrom & "# TO #" & myDTTo & "#"
    
    Dim frm As New frmPrevRpt
    frm.ShowReport strSelection, rpt
    
End Sub

Private Sub mnuPrintSFStock_Click()
        
    Dim strStoreIDs As String
    Dim frm As New frmGetStores
    If frm.getData(strStoreIDs) = False Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\BinWiseSemiFinishStockStatus.rpt")
    
    Dim f As New frmPrevRpt
    Dim strSelection As String
    strSelection = "{VItemsWithShelfWiseStock.Qty}>0"
    If strStoreIDs <> "All" Then
        strSelection = strSelection & " AND {VStoreShelfs.Store_RefID} IN[" & strStoreIDs & "]"
    End If
    
'    If cmbItems.ID <> "0" Then
'        strSelection = strSelection & " AND {VItems.ItemID}='" & cmbItems.ID & "'"
'    End If
    
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub mnuPrintSpecSheet_Click()

    On Error GoTo err
    
    If LV.ListItems.count = 0 Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ItemSpecification.rpt")
    Dim f As New frmPrevRpt, strCondition As String
    strCondition = "{Items.ItemID}='" & LV.SelectedItem.key
    f.ShowReport strCondition, rpt
   
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintStorewiseStatus_Click()
    
    Dim strStoreIDs As String
    Dim frm As New frmGetStores
    If frm.getData(strStoreIDs) = False Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\BinWiseStockStatusOfItems.rpt")
    Dim f As New frmPrevRpt
    Dim strSelection As String
    strSelection = "{VItemsWithShelfWiseStock.NetQty}>0"
    If strStoreIDs <> "All" Then
        strSelection = strSelection & " AND {VStoreShelfs.Store_RefID} IN[" & strStoreIDs & "]"
    End If
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub mnuPrintSummaryStock_Click()
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Complete_Stock_Summary.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VItems_StockReport.InActive}=FALSE", rpt
End Sub

Private Sub mnuPrintTechDrawing_Click()
    Dim strItemID As String
    strItemID = LV.SelectedItem.key
    strItemID = Left(strItemID, Len(strItemID) - 1)
    Call ShowTechDrawing("Items", "TechnicalDrawing", " WHERE ItemID='" & strItemID & "'", strItemID, MainForm.CD1)
End Sub

Private Sub mnuPurchases_Click()
    frmItemPurchase.Show
    frmItemPurchase.SetFocus
End Sub

Private Sub mnuResetProcessGroups_Click()

    If MsgBox("Are you sure to Do this? This might take Several Minutes.", vbQuestion + vbYesNo) = vbYes Then
        Debug.Print Now
        Dim i As Integer, j As Integer
        Dim strItems() As String, lItemCount As Long
        lItemCount = GetSingleLongValue("COUNT(DISTINCT ItemID)", "ItemProcesses")
        ReDim strItems(lItemCount, 1)
        Dim iItemProcesses() As Integer, iFirstProcessID As Integer, iLastProcessID As Integer, iTotalProcesses As Integer
        Dim iPossiblePGs() As Integer
        
        Dim rs As New ADODB.Recordset
        With rs
            .Open "SELECT DISTINCT ItemID FROM ItemProcesses", con, adOpenForwardOnly, adLockReadOnly
            Do Until .EOF
                strItems(.AbsolutePosition - 1, 0) = !ITemID & ""
                .MoveNEXT
            Loop
            .Close
        End With
        
        For i = 0 To UBound(strItems) - 1
            With rs
                If .State = 1 Then .Close
                .Open "SELECT ProcessID FROM ItemProcesses WHERE ItemID='" & strItems(i, 0) & "' ORDER BY SNo"
                iTotalProcesses = .RecordCount
                If iTotalProcesses = 0 Then GoTo SKIPITEM
                ReDim iItemProcesses(iTotalProcesses - 1)
                
                Do Until .EOF
                    iItemProcesses(.AbsolutePosition - 1) = !ProcessID
                    .MoveNEXT
                Loop
                .Close
                
                iFirstProcessID = iItemProcesses(0)
                iLastProcessID = iItemProcesses(iTotalProcesses - 1)
                
                'Now We'll Check this Process List against Process Groups...
                'If It's Exactly Matches with some Process Group We'll Assign it to Item...
                .Open "SELECT * FROM ProcessGroupsProcesses WHERE SeqNo=1 AND Process_RefID=" & iFirstProcessID & _
                 " AND Group_RefID IN(SELECT Group_RefID FROM ProcessGroupsProcesses GROUP BY Group_RefID HAVING COUNT(*)=" & _
                 iTotalProcesses & ")", con, adOpenForwardOnly, adLockReadOnly
                 
                If .RecordCount = 0 Then GoTo SKIPITEM
                ReDim iPossiblePGs(.RecordCount - 1)
                Do Until .EOF
                    iPossiblePGs(.AbsolutePosition - 1) = !Group_RefID
                    .MoveNEXT
                Loop
                .Close
                
                Dim iIndex As Integer, iProcessGroup As Integer
                For j = 0 To UBound(iPossiblePGs) - 1
                    .Open "SELECT Process_RefID FROM ProcessGroupsProcesses WHERE Group_RefID=" & iPossiblePGs(j) & " ORDER BY SeqNo", con, adOpenForwardOnly, adLockReadOnly
                    iIndex = 0
                    Do Until .EOF
                        If !Process_RefID = iItemProcesses(iIndex) Then
                            'Current Process Matched
                            iProcessGroup = iPossiblePGs(j)
                        Else
                            'Current Process did not Match
                            iProcessGroup = 0
                            GoTo ProcessDidNotMatch
                        End If
                        iIndex = iIndex + 1
                        .MoveNEXT
                    Loop
ProcessDidNotMatch:
                    .Close
                Next
                If iProcessGroup <> 0 Then
                    'It Means Processes Matched with some group's, Now Update it.
                    con.Execute "DELETE FROM ItemProcessGroups WHERE ItemID='" & strItems(i, 0) & "'"
                    con.Execute "INSERT INTO ItemProcessGroups(ItemID,PG_RefID) VALUES('" & strItems(i, 0) & "'," & iProcessGroup & ")"
                End If
            End With
SKIPITEM:
        Next
    End If
    Debug.Print Now
    MsgBox "Done.", vbInformation
End Sub

Private Sub mnuStockHistory_Click()
'Dim f As New frmItemLedger
Exit Sub
With f
   .Show
   Call FindInCombo(.cmbItemID, LV.SelectedItem)
   Call .cmdRefresh_Click
End With
End Sub

Private Sub mnuNewItem_Click()
    
    Dim f As New frmNewItem
    Load f
    
    Dim strNewItemCode As String
    strNewItemCode = f.AddNew(bSample)
    Call mnuRefresh_Click
    
    LV.ListItems(strNewItemCode & "'").Selected = True
    LV.ListItems(strNewItemCode & "'").EnsureVisible
    
End Sub

Private Sub mnunewcat_Click()
    Load frmNewItemCat
    frmNewItemCat.Show 1
'    Call mnuRefresh_Click
End Sub

Private Sub mnuParent_Click()
   If MsgBox("Are You Sure To Make This Account Parent", vbQuestion + vbYesNo) = vbNo Then Exit Sub
   con.Execute "update accounts set Parent=True Where AccNo='" & LV.SelectedItem & "'"
   Call mnuRefresh_Click
End Sub

Private Sub mnuPrint_Click()
    
    Screen.MousePointer = vbHourglass

    Load frmPrevRpt
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptDMR.rpt") '
    frmPrevRpt.ShowReport "{VrptDMR.ItemID}='" & LV.SelectedItem.key, rpt
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuPrintBal_Click()

    Dim f As New frmPrevRpt
    
    Dim rpt As CRAXDDRT.Report, rptApp_Temp As New CRAXDDRT.Application
    
    Set rpt = rptApp_Temp.OpenReport(RptPath & "\rptDMRInternal.rpt") '
    f.ShowReport "{VrptDMR.ItemID}='" & LV.SelectedItem.key, rpt
    
End Sub

Private Sub mnuRefresh_Click()
   Screen.MousePointer = vbHourglass
   'rsAccounts.Close
   'Set rsAccounts = Nothing
   'Call Form_Load
    cmbItem.ClearVals
    Dim strCondition As String
    strCondition = " WHERE InActive=0"
    
    Dim strCatCondition As String
    If iCatalogType = 1 Then    'Surgical
        strCatCondition = " AND CatID=51"
    ElseIf iCatalogType = 2 Then    'Dental
        strCatCondition = " AND CatID=52"
    ElseIf iCatalogType = 3 Then
        strCatCondition = " AND CatID NOT IN(51,52)"
    Else
        strCatCondition = ""
    End If
    strCatCondition = strCondition & strCatCondition
    
    Dim strConditionEx As String
    strConditionEx = " WHERE Sample=" & Abs(bSample)
    If mnuShowAll.Checked = False Then
        strConditionEx = strCondition & " AND InActive=0"
    End If
    Call cmbItem.AddVals(con, "ItemID + ' ' + ItemName", "Items", "ItemID", strConditionEx & " ORDER BY ItemName")
    
   Call AddHeads
   'Call FillItems
   Screen.MousePointer = vbDefault
   
End Sub


Private Sub CheckMenu()
    
    If strCompany = "Sunlike" Then
        mnuDeleteItem.Visible = True
    Else
        mnuDeleteItem.Visible = False
    End If
    
    mnuLockThisItem.Visible = False
    If LV.ListItems.count = 0 Then
        mnuNewItem.Enabled = False
        mnuEditItem.Enabled = False
        mnuSales.Enabled = False
        mnuPurchases.Enabled = False
        mnuStockHistory.Enabled = False
        mnuPrint.Enabled = False
        mnuParent.Enabled = False
        mnuNewCat.Enabled = True
        mnuDeleteItem.Enabled = False
    Else
        mnuNewItem.Enabled = True
        mnuEditItem.Caption = IIf(LV.SelectedItem.SmallIcon Like "ITM*", "Edit This Item", "Edit This Catagory")
        mnuSales.Enabled = True
        mnuPurchases.Enabled = True
        mnuStockHistory.Enabled = LV.SelectedItem.SmallIcon Like "ITM*"
        If LV.SelectedItem.SmallIcon Like "ITM*" Then
            mnuLockThisItem.Visible = True
            If GetSingleLongValue("COUNT(ItemCode)", "ItemsLockedForEditing", " WHERE ItemCode='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "'") > 0 Then
                mnuLockThisItem.Caption = "Un-Lock This Item"
                mnuLockThisItem.Tag = "1"
            Else
                mnuLockThisItem.Caption = "Lock This Item"
                mnuLockThisItem.Tag = "0"
            End If
            mnuDeleteItem.Enabled = True
        End If
        If LV.SelectedItem.Tag = "Inactive" Then
            mnuEditItem.Enabled = False
            mnuSales.Enabled = False
            mnuPurchases.Enabled = False
            mnuStockHistory.Enabled = False
            mnuParent.Enabled = False
            mnuActive.Caption = "Make This Item A&ctive."
        Else
            mnuActive.Caption = "Make This Item Ina&ctive."
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
   frmItemSale.Show
   frmItemSale.SetFocus
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
    Dim lCount As Long
    
    With LV.SelectedItem
        lCount = GetSingleLongValue("COUNT(CompItemCode)", "FOrderItems", " WHERE CompItemCode='" & .key)
        If lCount > 0 Then
            MsgBox "Can't Make Inactive, Item is in Orders.", vbInformation
            Exit Function
        End If
        If MsgBox("Are You Sure to Make This Item Inactive", vbYesNo + vbQuestion) = vbNo Then Exit Function
        con.Execute "Update Items Set InActive=1 Where ItemID='" & .key
        MakeAccountInActive = True
    End With

End Function

Private Function MakeAccountActive() As Boolean

    MakeAccountActive = False
    Dim AlsoChild As String

    With LV.SelectedItem
        con.Execute "Update Items Set InActive=0 Where ItemID='" & .key
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
    
    'ImgLoading.Visible = True
    Dim f As New frmWaitAnimation
    f.Show
    
    DoEvents
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim LColor As Long, lRed As Long, lGreen As Long, lBlue As Long
    Dim dTotalValue As Double
    Dim ItemOrdCond As String
    
    If strOrderCond = "" Then
        strOrderCond = " Order By CatID "
        ItemOrdCond = ""
    Else
        If strOrderCond = " Order By Description" Then
            ItemOrdCond = " Order By ItemName"
        Else
            ItemOrdCond = strOrderCond
        End If
    End If
    
    Dim strCatCondition As String
    If iCatalogType = 1 Then    'Surgical
        strCatCondition = " WHERE CatID=51"
    ElseIf iCatalogType = 2 Then    'Dental
        strCatCondition = " WHERE CatID=52"
    ElseIf iCatalogType = 3 Then
        strCatCondition = " WHERE CatID NOT IN(51,52)"
    Else
        strCatCondition = " WHERE 1=1"
    End If
    
    If cmbItemCatagory.ID <> "0" Then
        strCatCondition = strCatCondition & " AND CatID='" & cmbItemCatagory.ID & "'"
    End If
    
    With rs
        .Open "SELECT * FROM VItemCatagories " & strCatCondition & " ORDER BY Description ", con, adOpenForwardOnly, adLockReadOnly
        Set LV.SmallIcons = ImageList1
        If LV.ListItems.count > 0 Then PrevIndex = LV.SelectedItem.Index
    
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            LColor = vbBlack
            Set ITM = LV.ListItems.add(, ![CatID] & "'CAT", "", , "CAT")
            ITM.Bold = True
            ITM.ForeColor = LColor
            ITM.ListSubItems.add
            ITM.ListSubItems.add(, , ![Description] & "").Bold = True
            ITM.ListSubItems.add
            ITM.ListSubItems.add
            ITM.ListSubItems.add
            ITM.ListSubItems.add
            ITM.ListSubItems.add
            dTotalValue = dTotalValue + Val(!TotalValue & "")
            ITM.ListSubItems.add , , Val(!TotalValue & "")
            ITM.ListSubItems.add
            
            Call FillItems(![CatID], ItemOrdCond)
            .MoveNEXT
        Next
        .Close
    End With
    txtTotalValue = Format(dTotalValue, "#,##")
    'Format(OpenBal, "#,##0.00 DR;#,##0.00 CR;#,##0.00      ")
    Dim strCondition_Summary As String
    strCondition_Summary = strCatCondition
    If cmbGroup.ID <> "0" Then
        strCondition_Summary = strCondition_Summary & " AND GroupID=" & cmbGroup.ID
    End If
    If cmbAdditionalGroup.ID <> "0" Then
        strCondition_Summary = strCondition_Summary & " AND MainGroupID=" & cmbAdditionalGroup.ID
    End If
    txtActiveArticles = Format(GetSingleLongValue("COUNT(ItemID)", "Items", strCondition_Summary), "#,##")
    txtTotalStockQty = Format(GetSingleLongValue("SUM(InHand)", "VItems", strCondition_Summary), "#,##")
    
    Set rs = Nothing
'    Call FillItems
        
    'ImgLoading.Visible = False
    Unload f
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Function GetlvlStr(CatID As String) As String
    GetlvlStr = Space(5 * SubStrCount(CatID, "-"))
    If Len(GetlvlStr) > 0 Then GetlvlStr = GetlvlStr & "» "
End Function

Private Sub txtRef_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        'Try to Find Article...
        Dim strItemID As String, strSelItemID As String, strCondition As String
        strCondition = " WHERE ItemID IN(SELECT ItemID FROM ItemsCatalogReferences WHERE CatalogName='" & cmbCatalog.Text & "' AND Reference='" & txtRef & "')"
        
'        If LV.SelectedItem Is Nothing Then
'            strCondition = strCondition & " ORDER BY ItemID"
'        Else
'            If Right(LV.SelectedItem.key, 3) <> "CAT" Then
'                strCondition = strCondition & " AND ItemID>'" & LV.SelectedItem.key & " ORDER BY ItemID"
'            Else
'                strCondition = strCondition & " ORDER BY ItemID"
'            End If
'        End If
        strCondition = strCondition & " ORDER BY ItemID"
        strItemID = GetSingleStringValue("TOP 1 ItemID", "Items", strCondition)
        
        
        If strItemID = "" Then
            MsgBox "No Mathes Found.", vbInformation
            Set LV.SelectedItem = Nothing
        Else
            LV.ListItems(strItemID & "'").Selected = True
            LV.ListItems(strItemID & "'").EnsureVisible
        End If
    
    End If
    
End Sub
