VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmVendIssItemsByMakerOriginal 
   ClientHeight    =   10125
   ClientLeft      =   1110
   ClientTop       =   360
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
   MDIChild        =   -1  'True
   ScaleHeight     =   10125
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin VB.Frame FraSkip 
      Caption         =   "Skip to :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3270
      Left            =   645
      TabIndex        =   63
      Top             =   1065
      Visible         =   0   'False
      Width           =   5040
      Begin VB.CommandButton cmdSkipOk 
         Caption         =   "Ok"
         Height          =   435
         Left            =   1170
         TabIndex        =   66
         Top             =   2760
         Width           =   1320
      End
      Begin VB.CommandButton cmdSkipCancel 
         Caption         =   "Cancel"
         Height          =   435
         Left            =   2535
         TabIndex        =   65
         Top             =   2760
         Width           =   1320
      End
      Begin ComboList.Usercontrol1 cmbSkipProcess 
         Height          =   285
         Left            =   1005
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   240
         Width           =   3915
         _ExtentX        =   6906
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
      Begin VB.Label Label6A 
         AutoSize        =   -1  'True
         Caption         =   "Process :"
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
         Height          =   195
         Left            =   225
         TabIndex        =   62
         Top             =   270
         Width           =   750
      End
      Begin VB.Label lblSkipDesc 
         Caption         =   "This will make the selected item to skip the current process && make it Issuable to the selected Process."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   645
         Left            =   210
         TabIndex        =   67
         Top             =   1365
         Width           =   4710
      End
   End
   Begin VB.CommandButton cmdUpdateRate 
      Cancel          =   -1  'True
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2265
      TabIndex        =   52
      Top             =   5040
      Width           =   600
   End
   Begin VB.Frame FraStockOrder 
      Caption         =   "Article for Stock Order Issuance :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3330
      Left            =   645
      TabIndex        =   45
      Top             =   1065
      Visible         =   0   'False
      Width           =   10485
      Begin VB.CommandButton cmdSOCancel 
         Caption         =   "Cancel"
         Height          =   435
         Left            =   9060
         TabIndex        =   49
         Top             =   840
         Width           =   1320
      End
      Begin VB.CommandButton cmdSOOK 
         Caption         =   "Ok"
         Height          =   435
         Left            =   7680
         TabIndex        =   48
         Top             =   840
         Width           =   1320
      End
      Begin VB.TextBox txtSOQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9615
         TabIndex        =   46
         Top             =   465
         Width           =   765
      End
      Begin ComboList.Usercontrol1 cmbSOItems 
         Height          =   285
         Left            =   75
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   465
         Width           =   9510
         _ExtentX        =   16775
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Article :"
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
         Height          =   195
         Left            =   75
         TabIndex        =   51
         Top             =   255
         Width           =   645
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Qty :"
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
         Height          =   195
         Left            =   9600
         TabIndex        =   50
         Top             =   255
         Width           =   390
      End
   End
   Begin VB.TextBox txtExcess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataField       =   "Fax2"
      DataSource      =   "datPrimaryRS"
      Height          =   330
      Left            =   10170
      MaxLength       =   100
      TabIndex        =   40
      Tag             =   "ExcessQtyPercentage"
      Text            =   "100"
      Top             =   4650
      Width           =   1110
   End
   Begin VB.TextBox txtRcvings 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataField       =   "Fax2"
      DataSource      =   "datPrimaryRS"
      Height          =   330
      Left            =   9015
      MaxLength       =   100
      TabIndex        =   39
      Tag             =   "MaximumRcvingsAgainstPO"
      Text            =   "50"
      Top             =   4650
      Width           =   1110
   End
   Begin VB.TextBox txtAssignedUnit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3660
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   5235
      Width           =   810
   End
   Begin ComboList.Usercontrol1 cmbProcess 
      Height          =   285
      Left            =   420
      TabIndex        =   8
      TabStop         =   0   'False
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
      Appearance      =   0
   End
   Begin MSComCtl2.DTPicker DTReturn 
      Height          =   285
      Left            =   9285
      TabIndex        =   31
      Top             =   5235
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   299827203
      CurrentDate     =   37711
   End
   Begin VB.TextBox txtVendRate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1815
      TabIndex        =   30
      Top             =   5235
      Width           =   1110
   End
   Begin ComboList.Usercontrol1 cmbVender 
      Height          =   285
      Left            =   2790
      TabIndex        =   12
      Top             =   810
      Width           =   6450
      _ExtentX        =   11377
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
      Left            =   2910
      TabIndex        =   14
      Top             =   5235
      Width           =   765
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5745
      TabIndex        =   9
      Top             =   9495
      Width           =   5280
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3480
         TabIndex        =   18
         Top             =   165
         Width           =   1665
      End
      Begin VB.CommandButton cmdSaveClose 
         Caption         =   "&Save && Print"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1770
         TabIndex        =   17
         Top             =   165
         Width           =   1665
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save && &New   "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   60
         TabIndex        =   16
         Top             =   165
         Width           =   1665
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3915
      Left            =   405
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5550
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   6906
      _Version        =   393216
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Vender Balances"
      TabPicture(0)   =   "frmVendIssItemsByMakerOriginal.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "LVBal"
      Tab(0).Control(1)=   "txtTotal"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Vender History"
      TabPicture(1)   =   "frmVendIssItemsByMakerOriginal.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LVMaker"
      Tab(1).Control(1)=   "txtRate"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Issued Items"
      TabPicture(2)   =   "frmVendIssItemsByMakerOriginal.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label5(0)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label5(1)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "LVIssued"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "txtRate1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "txtTotalValue"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "txtTotalQty"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Other Vender's History"
      TabPicture(3)   =   "frmVendIssItemsByMakerOriginal.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "LVMakerOther"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "RM. Issuance"
      TabPicture(4)   =   "frmVendIssItemsByMakerOriginal.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label6"
      Tab(4).Control(1)=   "cmdAdd"
      Tab(4).Control(2)=   "LVRMIss"
      Tab(4).Control(3)=   "cmbMaterial"
      Tab(4).Control(4)=   "cmbStore"
      Tab(4).Control(5)=   "txtPrice"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "txtInHand"
      Tab(4).Control(7)=   "txtRMQty"
      Tab(4).Control(8)=   "txtUnit"
      Tab(4).ControlCount=   9
      Begin VB.TextBox txtUnit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -67215
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   690
         Width           =   615
      End
      Begin VB.TextBox txtRMQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -65595
         TabIndex        =   57
         Top             =   690
         Width           =   825
      End
      Begin VB.TextBox txtInHand 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -66585
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   690
         Width           =   960
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -69405
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   690
         Width           =   795
      End
      Begin VB.TextBox txtTotalQty 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7290
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   3435
         Width           =   1275
      End
      Begin VB.TextBox txtTotalValue 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9585
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   3435
         Width           =   1275
      End
      Begin VB.TextBox txtTotal 
         Height          =   285
         Left            =   -65475
         TabIndex        =   26
         Top             =   3540
         Width           =   1335
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -70065
         TabIndex        =   23
         Top             =   630
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.TextBox txtRate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2145
         TabIndex        =   21
         Top             =   1575
         Visible         =   0   'False
         Width           =   1350
      End
      Begin MSComctlLib.ListView LVIssued 
         Height          =   3030
         Left            =   75
         TabIndex        =   22
         Top             =   360
         Width           =   10800
         _ExtentX        =   19050
         _ExtentY        =   5345
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S No."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Vender"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Return Process"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Buyer"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Order No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Item Code"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Return Date"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Rate"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Issued Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Value"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Priority"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView LVMaker 
         Height          =   3450
         Left            =   -74925
         TabIndex        =   24
         Top             =   375
         Width           =   10770
         _ExtentX        =   18997
         _ExtentY        =   6085
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S No."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Process"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item Code"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Item Description"
            Object.Width           =   4586
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Qty"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Rate"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Value"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Del. Date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Priority"
            Object.Width           =   1940
         EndProperty
      End
      Begin MSComctlLib.ListView LVBal 
         Height          =   3075
         Left            =   -74940
         TabIndex        =   25
         Top             =   360
         Width           =   10785
         _ExtentX        =   19024
         _ExtentY        =   5424
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
            Text            =   "Process"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Item Description"
            Object.Width           =   7761
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Balance"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView LVMakerOther 
         Height          =   3450
         Left            =   -74925
         TabIndex        =   32
         Top             =   375
         Width           =   10770
         _ExtentX        =   18997
         _ExtentY        =   6085
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S No."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Maker"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item Code"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Item Description"
            Object.Width           =   4586
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Qty"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Rate"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Value"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Del. Date"
            Object.Width           =   1940
         EndProperty
      End
      Begin ComboList.Usercontrol1 cmbStore 
         Height          =   285
         Left            =   -68580
         TabIndex        =   54
         Top             =   690
         Width           =   1350
         _ExtentX        =   2381
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
      Begin ComboList.Usercontrol1 cmbMaterial 
         Height          =   285
         Left            =   -74925
         TabIndex        =   56
         Top             =   690
         Width           =   5505
         _ExtentX        =   9710
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
      Begin MSComctlLib.ListView LVRMIss 
         Height          =   2790
         Left            =   -74925
         TabIndex        =   61
         Top             =   1035
         Width           =   10740
         _ExtentX        =   18944
         _ExtentY        =   4921
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S No."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Material ID"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Material Name"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Store"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Unit"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Rate"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSForms.CommandButton cmdAdd 
         Height          =   345
         Left            =   -64740
         TabIndex        =   60
         Top             =   690
         Width           =   555
         Caption         =   "Add"
         PicturePosition =   2
         Size            =   "979;609"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.Label Label6 
         Height          =   240
         Left            =   -74925
         TabIndex        =   59
         Top             =   435
         Width           =   10740
         BackColor       =   11517387
         Caption         =   $"frmVendIssItemsByMakerOriginal.frx":008C
         Size            =   "18944;423"
         BorderColor     =   16512
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total Qty :"
         Height          =   195
         Index           =   1
         Left            =   6360
         TabIndex        =   44
         Top             =   3495
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total Value :"
         Height          =   195
         Index           =   0
         Left            =   8655
         TabIndex        =   36
         Top             =   3495
         Width           =   900
      End
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
      Left            =   10680
      Picture         =   "frmVendIssItemsByMakerOriginal.frx":014D
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5010
      UseMaskColor    =   -1  'True
      Width           =   600
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
            Picture         =   "frmVendIssItemsByMakerOriginal.frx":024F
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVendIssItemsByMakerOriginal.frx":040B
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
            Picture         =   "frmVendIssItemsByMakerOriginal.frx":05C7
            Key             =   "Active"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVendIssItemsByMakerOriginal.frx":0A1B
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
      Left            =   420
      TabIndex        =   15
      Top             =   5250
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   198901763
      CurrentDate     =   37711
   End
   Begin MSComctlLib.ListView LV 
      Height          =   3255
      Left            =   435
      TabIndex        =   5
      Top             =   1110
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   5741
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "S No."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Order No"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Item Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Description"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Issueable Qty"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Int. Ref."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Order Qty"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   165
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3375
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   503
      _Version        =   393216
      CheckBox        =   -1  'True
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   198901763
      CurrentDate     =   37711
   End
   Begin VB.Frame FraLoad 
      Height          =   3255
      Left            =   420
      TabIndex        =   19
      Top             =   1140
      Visible         =   0   'False
      Width           =   10860
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Loading Data"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2295
         Left            =   90
         TabIndex        =   20
         Top             =   780
         Width           =   10605
      End
   End
   Begin VB.CommandButton cmdCheckAll 
      Caption         =   "Check &All"
      Height          =   255
      Left            =   30
      TabIndex        =   27
      Top             =   6660
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Height          =   255
      Left            =   1110
      TabIndex        =   28
      Top             =   6660
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   420
      TabIndex        =   33
      Top             =   4350
      Width           =   8565
      Begin VB.TextBox txtSpecialInstructions 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   60
         TabIndex        =   34
         Top             =   330
         Width           =   4215
      End
      Begin VB.Label lblSteel 
         AutoSize        =   -1  'True
         Caption         =   "Steel :"
         Height          =   195
         Left            =   4305
         TabIndex        =   70
         Top             =   120
         Width           =   465
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Special Instructions :"
         Height          =   195
         Left            =   105
         TabIndex        =   69
         Top             =   120
         Width           =   1500
      End
      Begin MSForms.ComboBox cmbSteel 
         Height          =   315
         Left            =   4290
         TabIndex        =   68
         Tag             =   "SteelUsed"
         Top             =   330
         Width           =   4185
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "7382;556"
         ListWidth       =   8819
         MatchEntry      =   1
         ListStyle       =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Excess Qty %"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   10185
      TabIndex        =   42
      Top             =   4440
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Max. Rcvings"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   4
      Left            =   9015
      TabIndex        =   41
      Top             =   4440
      Width           =   960
   End
   Begin MSForms.ComboBox cmbReturnProcess 
      Height          =   285
      Left            =   4455
      TabIndex        =   37
      Top             =   5235
      Width           =   4815
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "8493;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbOrderNo 
      Height          =   285
      Left            =   165
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   3930
      Visible         =   0   'False
      Width           =   1965
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3466;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label3 
      Height          =   240
      Left            =   420
      TabIndex        =   13
      Top             =   5010
      Width           =   10230
      BackColor       =   11517387
      Caption         =   $"frmVendIssItemsByMakerOriginal.frx":0E6F
      Size            =   "18045;423"
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
      Left            =   165
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3645
      Visible         =   0   'False
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
      Left            =   9240
      TabIndex        =   10
      TabStop         =   0   'False
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
      Left            =   420
      TabIndex        =   3
      Top             =   585
      Width           =   10770
      BackColor       =   11517387
      Caption         =   $"frmVendIssItemsByMakerOriginal.frx":0F1F
      Size            =   "18997;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Menu mnuPop 
      Caption         =   "Pop up Menu"
      Begin VB.Menu mnuSkipProcess 
         Caption         =   "Skip Process"
      End
      Begin VB.Menu mnuDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStockOrderIssuance 
         Caption         =   "Stock Order Issuance"
      End
   End
End
Attribute VB_Name = "frmVendIssItemsByMakerOriginal"
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

Dim lVendBal As Long
Dim iColNo As Integer
Dim bFromMaker As Boolean

Dim bReIssuance As Boolean
Dim bDontUpdateFromProcess As Boolean
Dim bDontUpdateFromMaker As Boolean
Dim bCanChangeRate As Boolean

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
            For i = 1 To MyCol.count
                If MyCol(i).CompItemCode = ![ITemID] Then
                    cmbItems.AddItem ![ITemID] 'MyCol(i).ItemCode
                    cmbItems.List(cmbItems.ListCount - 1, 1) = MyCol(i).ItemCode
                    cmbItems.List(cmbItems.ListCount - 1, 2) = MyCol(i).lQty
                    cmbItems.List(cmbItems.ListCount - 1, 3) = ![Rate]
                    cmbItemDesc.AddItem MyCol(i).ItemDescription
                    Exit For
                End If
            Next i
            .MoveNEXT
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

    If cmbDesc.MatchFound = False Then Exit Sub
    cmbItems.ListIndex = cmbDesc.ListIndex
    
    
    If RcvingLot Then
        LV.ListItems(cmbDesc.ListIndex + 1).Selected = True
    End If
    
End Sub

Private Sub cmbItems_Change()
    Call cmbItems_Click
End Sub

Private Sub cmbItems_Click()

    If Not cmbItems.MatchFound Then Exit Sub
    'cmbDesc.ListIndex = cmbItems.ListIndex
    'txtVendRate = Val(cmbItems.List(cmbItems.ListIndex, 1))
    If bReIssuance Then
        Call FillList(cmbProcess.ID, cmbVender.ID, cmbItems.Text)
    Else
        Call FillList
    End If
    '/\/\/\/\/\/\/\/\/\/
    Dim rs As New ADODB.Recordset, lSNo As Long
    With rs
        .Open "SELECT SNO FROM ItemProcesses WHERE ProcessID=" & cmbProcess.ID & " AND ItemID='" & cmbItems.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            lSNo = Val(!SNo & "")
        End If
        .Close
    End With
    
    
'    cmbReturnProcess.ClearVals
'    cmbReturnProcess.AddVals con, "Description", "VItemProcesses", "ProcessID", " WHERE IsExist='" & _
'     cmbItems.Text & "' AND ProcessID IN(SELECT ProcessID FROM MakerProcesses WHERE VendID=" & _
'     cmbVender.ID & ") AND ItemSNO>=" & lSNo & " ORDER BY ItemSNo"
'
'    cmbReturnProcess.ID = cmbProcess.ID
'Or cmbProcess.ID = iSemiFinishProcessID

    If (cmbProcess.ID = iReadyFinishProcessID) Then
        rs.Open "SELECT Rate,Unit FROM VendAssItems WHERE VendID=" & cmbVender.ID & " AND ProcessID=" & cmbProcess.ID & " AND ItemID='" & cmbItems.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        With cmbReturnProcess
            .Clear
            .AddItem cmbProcess.Text '"Ready Finish"
            .List(0, 1) = cmbProcess.ID
            .List(0, 2) = Val(rs!Rate & "")
            .List(0, 3) = rs!Unit & ""
        End With
        cmbReturnProcess.ListIndex = 0
    Else
        Set rs = con.Execute("exec NextProcessesForIssuance_SP " & cmbProcess.ID & "," & cmbVender.ID & ",'" & cmbItems.Text & "'")
        Dim i As Integer
        cmbReturnProcess.Clear
        Do Until rs.EOF
            cmbReturnProcess.AddItem rs!Description & ""
            cmbReturnProcess.List(cmbReturnProcess.ListCount - 1, 1) = Val(rs!ProcessID & "")
            cmbReturnProcess.List(cmbReturnProcess.ListCount - 1, 2) = Val(rs!Rate & "")
            cmbReturnProcess.List(cmbReturnProcess.ListCount - 1, 3) = rs!Unit & ""
            rs.MoveNEXT
        Loop
        If cmbReturnProcess.ListCount > 0 Then cmbReturnProcess.ListIndex = 0
    End If
    rs.Close
    Set rsz = Nothing
    '/\/\/\/\/\/\/\/\/\/
    'Call FillVendList
    Call ShowOtherVendersHistory
End Sub

Private Sub ShowOtherVendersHistory()

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "Select Top 10 * From VMakerIssItems Where VendID<>" & cmbVender.ID & " AND ProcessID=" & cmbProcess.ID & " AND ItemCode='" & cmbItems & "' Order By DT Desc,EntryID Desc", con, adOpenForwardOnly, adLockReadOnly
        
        LVMakerOther.ListItems.Clear
        'LVBal.ListItems.Clear
        lVendBal = 0
        Do Until .EOF
            Set ITM = LVMakerOther.ListItems.add(, , LVMakerOther.ListItems.count + 1)
            ITM.SubItems(1) = ![VendID1] & ""
            ITM.SubItems(2) = ![ItemCode] & ""
            ITM.SubItems(3) = ![ItemName] & ""
            ITM.SubItems(4) = Format(![DT] & "", "dd-MM-yy")
            ITM.SubItems(5) = Val(![IssQty] & "")
            ITM.SubItems(6) = Val(![Rate] & "")
            ITM.SubItems(7) = Val(!IssValue & "") 'Round(Val(![IssQty] & "") * Val(![Rate] & ""))
            ITM.SubItems(8) = Format(![ReturnDT] & "", "dd-MM-yy")
            
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub FillList(Optional iProcessID As Integer = 0, Optional iMakerID As Integer = 0, Optional strItemID As String = "")

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
    If bReIssuance = False Then
        If cmbItems.MatchFound = False Then Exit Sub
        With cmd
            Set .ActiveConnection = con
            .CommandType = adCmdStoredProc
            .CommandText = "IssueItems_SP"
            .Parameters("@ProcessID").Value = cmbProcess.ID
            .Parameters("@ItemID").Value = cmbItems.Text
            Set rs = .Execute
        End With
    Else
        With cmd
            Set .ActiveConnection = con
            .CommandType = adCmdStoredProc
            .CommandText = "ReIssueItems_SP"
            .Parameters("@ProcessID").Value = iProcessID
            .Parameters("@MakerID").Value = iMakerID
            .Parameters("@ItemID").Value = strItemID
            Set rs = .Execute
        End With
    End If
    
    Dim lQty As Long
    
    
    With rs
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, !OrderNo & !CompITemID & "'", LV.ListItems.count + 1)
            ITM.ListSubItems.add , , ![OrderNo] & ""
            ITM.ListSubItems.add , , ![CompITemID] & ""
            ITM.ListSubItems.add , , ![ItemName] & ""
            If cmbProcess.ID = !InitialProcessID Or (cmbProcess.ID = iReadyFinishProcessID Or cmbProcess.ID = iSemiFinishProcessID Or cmbProcess.ID = iFilingProcessID Or cmbProcess.ID = iFiling2ProcessID) Then
                lQty = Val(!OrderQty & "") - Val(!TotIssdProc & "")
            Else
                lQty = Val(!IssQty & "")
            End If
            ITM.ListSubItems.add(, , lQty).Tag = lQty
            ITM.ListSubItems.add , , !InternalRefNo & ""
            
            ITM.ListSubItems.add(, , Val(!OrderQty & "")).Tag = Val(!TotIssdProc & "")
            
            LColor = RGB(150, 55, 150)
            ITM.ForeColor = LColor
            For i = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(i).ForeColor = LColor
            Next
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub



Private Sub cmbProcessOld_matched()
    
    On Error GoTo err
'    If cmbProcess.ID = 9 Then
'        cmdCheckAll.Visible = True
'        cmdClearAll.Visible = True
'    Else
'        cmdCheckAll.Visible = False
'        cmdClearAll.Visible = False
'    End If
    
    Dim cmd As New ADODB.Command
    
    LV.ListItems.Clear
    LVIssued.ListItems.Clear
    LVMaker.ListItems.Clear
    
    IsLotProcess = LotProcess(cmbProcess.ID)
    
    If IsLotProcess And RcvLot(cmbProcess.ID) Then
        RcvingLot = True
        cmbItems.Visible = False
        DTPicker1.Visible = True
        Label2.Caption = "  Process                                    Date                                      Lot No"
        LV.Checkboxes = True
        
        txtQty.Visible = False
        Label3.Caption = "  Venders                                                                Date"
        Label3.Width = 4830
        cmdPost.Left = DTPicker2.Left + DTPicker2.Width + 20
    Else
        RcvingLot = False
        cmbItems.Visible = True
        DTPicker1.Visible = False
        Label2.Caption = "  Process                                    Item Code                        Item Description"
        LV.Checkboxes = False
        
        txtQty.Visible = True
        Label3.Caption = "  Venders                                                                Date                      Quantity"
        Label3.Width = 5970
        cmdPost.Left = txtQty.Left + txtQty.Width + 20
        Me.Refresh
    End If
    
    Dim LColor As Long
    Dim ITM As ListItem
    
    With cmd
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        
        
        If IsLotProcess Then
            
            If RcvingLot Then
                LV.ColumnHeaders.add 2, , "Lot No.", 1200
                'LV.ColumnHeaders(2).Text = "Lot No/Order No"
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
        
        .Parameters("@ProcessID").Value = cmbProcess.ID
            
        Set rsMain = .Execute
    End With
    
    Dim isAdded As Boolean
    
    With rsMain
        cmbItems.Clear
        cmbDesc.Clear
        
        If RcvingLot Then
'            DTPicker1.value = ServerDate
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
                    
                
                Set ITM = LV.ListItems.add(, ![LotNo] & "''", LV.ListItems.count + 1)
                ITM.Tag = ![RecieptID] & ""
                'Debug.Print LV.ListItems.Count
                'ITM.Tag = ![EntryID]
                'itm.ListSubItems.Add , , ![CustCode] & ""
                ITM.ListSubItems.add(, , ![RecieptID] & "").Tag = ![LotNo] & ""
                
                ITM.ListSubItems.add(, , ![ItemCode] & "").Tag = ![ItemCode] & ""
                
                'itm.ListSubItems.Add , , ![ItemCode] & ""
                ITM.ListSubItems.add , , ![ItemDescription] & ""
                ITM.ListSubItems.add(, , Val(![IssQty] & "")).Tag = Val(![IssQty] & "")
                
                If IsNull(![LockedItemCode]) Then
                    LColor = LV.ForeColor
                Else
                    LColor = vbWhite
                End If
                
                ITM.ForeColor = LColor
                For i = 1 To ITM.ListSubItems.count
                    ITM.ListSubItems(i).ForeColor = LColor
                Next
                
                .MoveNEXT
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
                
                
            
                Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            
            
                'ITM.Tag = ![EntryID]
                
                ITM.ListSubItems.add , , ![ItemCode] & ""
                ITM.ListSubItems.add , , ![ItemDescription] & ""
                ITM.ListSubItems.add(, , Val(![IssQty] & "")).Tag = Val(![IssQty] & "")
                
                If IsNull(![LockedItemCode]) Then
                    LColor = LV.ForeColor
                Else
                    LColor = RGB(150, 55, 150)
                End If
                
                ITM.ForeColor = LColor
                For i = 1 To ITM.ListSubItems.count
                    ITM.ListSubItems(i).ForeColor = LColor
                Next
                .MoveNEXT
            Loop
        End If
    End With
    
    If LV.ListItems.count > 0 Then
        LV.ListItems(1).Selected = True
        Call LV_ItemClick(LV.ListItems(1))
    End If
    
    Me.Show
    cmbVender.SetFocus
    SendKeys "a"
    SendKeys Chr(8)
    
    FraLoad.Visible = False
    LV.Visible = True
    'Debug.Print Time
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbMaterial_matched()
        
    cmbStore.ClearVals
    cmbStore.AddVals con, "StoreName", "VRMStatusStorewise", "StoreName", " WHERE RMID=" & cmbMaterial.ID & " AND NetQty>0"
    
    Dim rs As New ADODB.Recordset
    
    With rs
        
        .Open "SELECT Unit,QtyInStock,Rate FROM RM WHERE RMID='" & cmbMaterial.ID & "'", con, adOpenForwardOnly, adLockReadOnly
         
        txtPrice = Val(!Rate & "")
        txtUnit = !Unit & ""
'        txtUnit.Tag = Val(!QtyInStock & "")
        txtInHand.Tag = Val(!QtyInStock & "")   'Total Qty from All Stores....
        
        .Close
        
    End With
    Set rs = Nothing
    
    Exit Sub
    
err:
    MsgBox err.Description

End Sub

Private Sub cmbProcess_matched()
    
    Dim iVendID As Integer
    
    If cmbVender.MatchFound Then
        iVendID = cmbVender.ID
    End If
    
    cmbVender.ClearVals
    If cmbProcess.ID <> "0" Then
        Dim iOperation As Integer
        iOperation = GetSingleLongValue("Operation", "Processes", " WHERE ProcessID=" & cmbProcess.ID)
        
        If iOperation = 0 Then
            cmbVender.AddVals con, "'{ ' + VendID1 + ' } ' + VenderName", "VMakers", "VendID", " WHERE VendID=" & iFactoryMakerID & " AND VendID IN(SELECT VendID FROM MakerProcesses WHERE ProcessID=" & cmbProcess.ID & ")"
        ElseIf iOperation = 1 Then
            cmbVender.AddVals con, "'{ ' + VendID1 + ' } ' + VenderName", "VMakers", "VendID", " WHERE VendID<>" & iFactoryMakerID & " AND VendID IN(SELECT VendID FROM MakerProcesses WHERE ProcessID=" & cmbProcess.ID & ")"
        ElseIf iOperation = 2 Then
            cmbVender.AddVals con, "'{ ' + VendID1 + ' } ' + VenderName", "VMakers", "VendID", " WHERE VendID IN(SELECT VendID FROM MakerProcesses WHERE ProcessID=" & cmbProcess.ID & ")"
        End If
    Else
        cmbVender.AddVals con, "'{ ' + VendID1 + ' } ' + VenderName", "VMakers", "VendID"
    End If
    
'    If cmbProcess.ID = "0" Then
'        cmbVender.AddVals con, "'{ ' + VendID1 + ' } ' + VenderName", "VMakers", "VendID", " "
'    Else
'        cmbVender.AddVals con, "'{ ' + VendID1 + ' } ' + VenderName", "VMakers", "VendID", " WHERE VendID IN(SELECT VendID FROM MakerProcesses WHERE ProcessID=" & cmbProcess.ID & ")"
'    End If
    
    If iVendID <> 0 Then
        cmbProcess.Tag = "From Process"
        bDontUpdateFromMaker = True
        cmbVender.ID = iVendID
        bDontUpdateFromMaker = False
    Else
        cmbProcess.Tag = ""
    End If
    
    If bDontUpdateFromProcess = False Then
        If bReIssuance Then
            'Fill Issuable Items...
            Call FillList(cmbProcess.ID)
        End If
    End If
    
End Sub

Private Sub cmbReturnProcess_Change()
    Call cmbReturnProcess_Click
End Sub


Private Sub cmbReturnProcess_Click()

    If cmbReturnProcess.MatchFound = False Then Exit Sub
    Dim iProcessNo As Integer, dRate As Double
    dRate = 0
    iProcessNo = cmbReturnProcess.List(cmbReturnProcess.ListIndex, 1)
    If (iProcessNo = iReadyFinishProcessID) Or (iProcessNo = iSemiFinishProcessID) Or (iProcessNo = iFilingProcessID) Or (iProcessNo = iFiling2ProcessID) Then 'Ready Finish
        dRate = Val(cmbReturnProcess.List(cmbReturnProcess.ListIndex, 2))
    Else
        Dim i As Integer
        For i = 0 To cmbReturnProcess.ListIndex
            dRate = dRate + Val(cmbReturnProcess.List(i, 2))
        Next
    End If
    txtAssignedUnit = cmbReturnProcess.List(cmbReturnProcess.ListIndex, 3)
    txtVendRate = dRate
    
End Sub

Private Sub cmbStore_matched()
    
    Dim dStoreQty As Double
    dStoreQty = GetSingleDoubleValue("NetQty", "VRMStatusStorewise", " WHERE RMID=" & cmbMaterial.ID & " AND StoreName='" & cmbStore.Text & "'")
    txtUnit.Tag = dStoreQty
    txtInHand = dStoreQty
    
End Sub

Private Sub cmbVender_matched()

    On Error GoTo err
'    If cmbVender.Text = "HICO" Then Exit Sub
'    If RcvingLot Then
'        cmbVender.Tag = Val(con.Execute("Select Rate From VMakerAssItems Where VendID=" & cmbVender.ID & " and ItemID='" & LV.SelectedItem.SubItems(2) & "'").Fields(0).value & "")
'    Else
'        cmbVender.Tag = Val(con.Execute("Select Rate From VMakerAssItems Where VendID=" & cmbVender.ID & " and ItemID='" & LV.SelectedItem.SubItems(1) & "'").Fields(0).value & "")
'    End If
    
    Dim iProcessID As Integer
    
    If cmbProcess.MatchFound Then
        iProcessID = cmbProcess.ID
    End If
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT VendID1 FROM Makers WHERE VendID=" & cmbVender.ID, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False And cmbProcess.ID = "0" Then
            cmbVender.Tag = !VendID1 & ""
            'cmbProcess.ID = Val(!PhaseID & "")
        End If
        .Close
        
        If cmbProcess.Tag = "" Then
            cmbProcess.ClearVals
            cmbProcess.AddVals con, "Description", "VMakerAssignedProcesses", "ProcessID", " WHERE VendID=" & cmbVender.ID
            
            .Open "SELECT * FROM VMakerInitialProcess WHERE VendID=" & cmbVender.ID, con, adOpenForwardOnly, adLockReadOnly
            If iProcessID <> 0 Then
                bDontUpdateFromProcess = True
                cmbProcess.ID = iProcessID
                bDontUpdateFromProcess = False
            ElseIf .EOF = False Then
                cmbProcess.ID = Val(!FirstProcessID & "")
            End If
            .Close
        End If
        
        .Open "SELECT * FROM VMakerAssItems WHERE VendID=" & cmbVender.ID & " AND ProcessID=" & cmbProcess.ID, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        cmbItems.Clear
        'cmbDesc.Clear
        Do Until .EOF
            cmbItems.AddItem !ITemID & ""
            cmbItems.List(.AbsolutePosition - 1, 1) = Val(!Rate & "")
            'cmbDesc.AddItem !Description & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    If bDontUpdateFromMaker = False Then
        If bReIssuance Then
            'Fill Issuable Items...
            Call FillList(cmbProcess.ID, cmbVender.ID)
        End If
    End If
    
    Call ShowVenderHistory
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub ShowVenderHistory()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim DTLast As Date
    DTLast = DateAdd("m", -3, Date)
    With rs
        .Open "SELECT * FROM VMakerIssItems WHERE VendID=" & cmbVender.ID & " AND DT>='" & DTLast & "' ORDER BY DT Desc,EntryID Desc", con, adOpenForwardOnly, adLockReadOnly
        
        LVMaker.ListItems.Clear
        LVBal.ListItems.Clear
        lVendBal = 0
        Do Until .EOF
            Set ITM = LVMaker.ListItems.add(, , LVMaker.ListItems.count + 1)
            ITM.SubItems(1) = !ProcessName & ""
            ITM.SubItems(2) = ![ItemCode] & ""
            ITM.SubItems(3) = ![ItemName] & ""
            ITM.SubItems(4) = Format(![DT] & "", "dd-MM-yy")
            ITM.SubItems(5) = Val(![IssQty] & "")
            ITM.SubItems(6) = Val(![Rate] & "")
            ITM.SubItems(7) = Val(!IssValue & "") 'Round(Val(![IssQty] & "") * Val(![Rate] & ""))
            ITM.SubItems(8) = Format(![ReturnDT] & "", "dd-MM-yy")
            ITM.SubItems(9) = Val(!Priority & "")
            .MoveNEXT
        Loop
        .Close
        
        'If cmbVender.ID <> 129 Then
            .Open "Select ProcessName,ItemCode,ItemName,Sum(RemQty) As RemQty From VMakerIssItems Where VendID=" & cmbVender.ID & "  AND RemQty>0 Group By ProcessName,ItemCode,ItemName Having Sum(RemQty)>0", con, adOpenForwardOnly, adLockReadOnly
            
            Do Until .EOF
                Set ITM = LVBal.ListItems.add(, , LVBal.ListItems.count + 1)
                ITM.SubItems(1) = !ProcessName & ""
                ITM.SubItems(2) = ![ItemCode] & ""
                ITM.SubItems(3) = ![ItemName] & ""
                ITM.SubItems(4) = Val(![RemQty] & "")
                lVendBal = lVendBal + Val(![RemQty] & "")
                .MoveNEXT
            Loop
            .Close
            txtTotal = lVendBal
        'End If
        
        
        
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbVender_Validate(cancel As Boolean)
    cmdPost.Default = True
End Sub

Private Sub cmdAdd_Click()
    
    On Error GoTo err
    
    If cmbMaterial.MatchFound = False Then Exit Sub
    If Val(txtRMQty) <= 0 Then Exit Sub
    
    Dim lSF_RefID As Long
    lSF_RefID = 0 'GetSingleLongValue("SF_RefID", "RM", " WHERE RMID=" & cmbMaterial.ID)
    
    If lSF_RefID = 0 Then   'If Regual Material then Apply check, No Check for SemiFinish materials...
        If Val(txtRMQty) > Val(txtUnit.Tag) Then
            MsgBox "Can't Issue This Quantity." & vbNewLine & "Only " & txtUnit.Tag & " " & txtUnit & " Is Available In The Store.", vbInformation
            Exit Sub
        End If
    End If
    
    Dim ITM As ListItem
    Set ITM = LVRMIss.ListItems.add(, cmbMaterial.ID & "'", LVRMIss.ListItems.count + 1)
    ITM.Tag = Val(txtInHand.Tag)
    ITM.ListSubItems.add(, , cmbMaterial.ID).Tag = cmbMaterial.Tag ''Material Type
    ITM.ListSubItems.add(, , cmbMaterial.Text).Tag = ""
    ITM.SubItems(3) = cmbStore.Text
    ITM.SubItems(4) = txtUnit
    ITM.SubItems(5) = Val(txtRMQty)
    ITM.SubItems(6) = Val(txtPrice)
    
    cmbStore.ClearVals
    cmbMaterial.ID = ""
    txtPrice = ""
    txtUnit = ""
    txtQty = ""
    txtInHand = ""
    txtInHand.Tag = ""
    cmbMaterial.SetFocus
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    con.Execute "Delete From PrintSession"
End Sub

Private Sub cmdCheckAll_Click()
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmdClearAll_Click()
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmdPost_Click()

    On Error GoTo err
'    If LV.SelectedItem.ForeColor = 9844630 Then
'        MsgBox "Can't Issue. Item Is Locked", vbInformation
'        Exit Sub
'    End If
        
    If Not cmbVender.MatchFound Then Exit Sub
    If cmbVender.ID = "" Or cmbVender.Text = "" Then Exit Sub
    If Val(txtVendRate) <= 0 And cmbVender.ID <> iFactoryMakerID Then
        MsgBox "Rate Can't be Zero!", vbInformation
        Exit Sub
    End If
    If cmbReturnProcess.MatchFound = False Then Exit Sub
    'cmbVender.Enabled = False
    
    'Check For Max Balance Limit For This Vender.......
    If cmbVender.ID <> 129 Then
        Dim rs As New ADODB.Recordset
        With rs
'            .Open "Select MaxLimit From Venders Where VendID=" & cmbVender.ID, con, adOpenForwardOnly, adLockReadOnly
'            If lVendBal >= Val(![MaxLimit] & "") Then
'                MsgBox "Sorry, The Selected Vender Has Reached His Max. Issuence Limit." & _
'                 vbNewLine & "Can't Post It For This Vender." & vbNewLine & "Vender Maximum Limit : " & Val(![MaxLimit] & "") & vbNewLine & "Vender Curr. Balance  : " & lVendBal, vbInformation
'
'                .Close
'                Exit Sub
'            End If
'            .Close
        End With
        Set rs = Nothing
    End If
    
    If RcvingLot Then
        PostForLots
        Exit Sub
    End If
    
    If Val(txtQty) < 1 Then Exit Sub
    
    If bReIssuance Then
        If Val(txtQty) > Val(LV.SelectedItem.SubItems(4)) Then
            MsgBox "Invalid Qty", vbInformation
            Exit Sub
        End If
    End If
    
    If bReIssuance = False Then 'Issue / Purchase--Check for 20% of Excess issuance for Wrangler.
        If strCompany = "IAA" Or strCompany = "Kami" Then
            Dim lQtyWrangler As Long, lOrderQtyWrangler As Long, lIssdQtyWrangler As Long, lExtraQtyWrangler As Long
            lQtyWrangler = Val(txtQty)
            lOrderQtyWrangler = Val(LV.SelectedItem.SubItems(6))
            lIssdQtyWrangler = Val(LV.SelectedItem.ListSubItems(6).Tag)
            
            lExtraQtyWrangler = (lOrderQtyWrangler / 100) * 20
            
            If lIssdQtyWrangler + lQtyWrangler > lOrderQtyWrangler + lExtraQtyWrangler Then
                MsgBox "Invalid Qty." & vbNewLine & "Can't Issue more than 20% of the Customer Order.", vbInformation
                Exit Sub
            End If
        End If
    End If
    
    If GotIssID = 0 Then
        'Dim IssID As String
        'For i = 1 To LVIssued.ListItems.Count
        '    If Val(cmbVender.ID) = Val(LVIssued.ListItems(i).ListSubItems(2).Key) Then
        '        IssID = LVIssued.ListItems(i).ListSubItems(2)
        '        GoTo GotID:
        '    End If
        'Next i
    
        GotIssID = Val(con.Execute("Select Max(Cast(right(RecieptID,Len(RecieptID)-10) As Int)) From VMakerIssItems Where Convert(varchar,DT,6)='" & Format(DTPicker2, "dd MMM yy") & "' and Left(RecieptID,3)<>'HSR'").Fields(0).Value & "")
        GotIssID = GotIssID + 1
    Else
        GotIssID = GotIssID + 1
    End If
    'IssID = "ISS-" & Val(LVMaker.SelectedItem.ListSubItems(1).Key) & "-" & _
     Format(ServerDate, "ddMMyy") & "-" & Format(Val(IssID) + 1, "000")
    IssID = "ISU-" & _
     Format(DTPicker2, "ddMMyy") & GotIssID
    
    Dim strBuyerCode As String
    strBuyerCode = con.Execute("SELECT CustCode FROM FCUstomerOrders WHERE OrderNo='" & LV.SelectedItem.SubItems(1) & "'").Fields(0).Value & ""
    
GotID:
        
    Dim ITM As ListItem
    With ITM
        'Set ITM = LVIssued.ListItems.Add(, Val(LV.SelectedItem.Key) & Val(LVMaker.SelectedItem.Key) & "'", LVIssued.ListItems.Count + 1)
        'Set ITM = LVIssued.ListItems.Add(, LV.SelectedItem.Key & "|" & LVIssued.ListItems.Count + 1, LVIssued.ListItems.Count + 1)
        Set ITM = LVIssued.ListItems.add(, , LVIssued.ListItems.count + 1)
        ITM.Tag = LV.SelectedItem.key
        'itm.ListSubItems.Add , cmbVender.ID & "'", cmbVender.Text
        ITM.ListSubItems.add(, , cmbVender.Tag).Tag = cmbVender.ID
        ITM.ListSubItems.add(, , cmbReturnProcess.Text).Tag = IssID & ""  'DTPicker2
        ITM.ListSubItems.add(, , strBuyerCode).Tag = cmbReturnProcess.List(cmbReturnProcess.ListIndex, 1) 'DTPicker2
        ITM.ListSubItems.add(, , LV.SelectedItem.SubItems(1)).Tag = cmbReturnProcess.List(cmbReturnProcess.ListIndex, 1)
        ITM.ListSubItems.add(, , LV.SelectedItem.SubItems(2)).Tag = txtSpecialInstructions.Text
        ITM.ListSubItems.add , , Format(DTReturn.Value, "dd-MMM-yyyy") 'LV.SelectedItem.SubItems(3)
        ITM.ListSubItems.add(, , Val(txtVendRate.Text)).Tag = cmbProcess.ID
        ITM.ListSubItems.add , , Val(txtQty) 'LV.SelectedItem.SubItems(3)
        ITM.ListSubItems.add , , GetValueofItem(cmbItems.Text)     'Val(txtQty) * Val(txtVendRate.Text)
        ITM.ListSubItems.add , , LVIssued.ListItems.count
        'ITM.ListSubItems.Add , , Val(txtQty)
        If cmbProcess.Tag = "Lot No" Then
            ITM.ListSubItems(6).Tag = LV.SelectedItem.SubItems(2) & ""
        End If
    End With
    Call getTotalValue
    
    'MyCol(LV.SelectedItem.Index).IssueableQty = MyCol(LV.SelectedItem.Index).IssueableQty - Val(txtQty)
    LV.SelectedItem.SubItems(4) = Val(LV.SelectedItem.SubItems(4)) - Val(txtQty)
     
    txtRate.Visible = False
    Dim i As Integer
    For i = 1 To LVIssued.ListItems.count
        If LVIssued.ListItems(i).SubItems(5) = LV.SelectedItem.SubItems(2) Then
            LVIssued.ListItems(i).ListSubItems(5).Tag = txtSpecialInstructions.Text
        End If
    Next
    'Removing Item From Combos
'    If LV.SelectedItem.SubItems(4) = 0 Then
'        Dim ItemIndex As Long
'        ItemIndex = cmbItems.ListIndex
'        If ItemIndex <> -1 Then
'
'            cmbItems.RemoveItem ItemIndex
'            cmbDesc.RemoveItem ItemIndex
'            cmbItems.Text = ""
'            cmbDesc.Text = ""
'            If cmbItems.ListCount >= ItemIndex And ItemIndex > 0 Then
'                cmbItems.ListIndex = ItemIndex
'            ElseIf cmbItems.ListCount > 0 Then
'                cmbItems.ListIndex = 0
'            End If
'        End If
'    End If
    'if cmbitems.
    'Call LV_ItemClick(LV.SelectedItem)
    
    txtQty = ""
    
    'Select The Next Item Automaticaly....
    If LV.SelectedItem.Index < LV.ListItems.count Then
        If Val(LV.SelectedItem.SubItems(3)) = 0 Then
            LV.ListItems(LV.SelectedItem.Index + 1).Selected = True
        End If
        Call LV_ItemClick(LV.ListItems(LV.SelectedItem.Index))
    End If
    
    'cmbVender.SetFocus
    
    Exit Sub
err:
    MsgBox err.Description
        'End If
End Sub

Private Function GetValueofItem(strItemID As String) As Double

    On Error GoTo err
    If txtAssignedUnit = "Kgs" Then
        Dim dWeight As Double
        dWeight = GetSingleDoubleValue("ForgingWeight", "Items", " WHERE ItemID='" & strItemID & "'")
        GetValueofItem = ((dWeight * Val(txtQty)) / 1000) * Val(txtVendRate.Text)
    Else
        GetValueofItem = Val(txtQty) * Val(txtVendRate.Text)
    End If
    
    Exit Function
err:
    MsgBox err.Description
End Function

Private Sub getTotalValue()

    Dim i As Integer, lTotal As Long, lQty As Long
    For i = 1 To LVIssued.ListItems.count
        lTotal = lTotal + Val(LVIssued.ListItems(i).SubItems(9))
        lQty = lQty + Val(LVIssued.ListItems(i).SubItems(8))
    Next
    
    txtTotalValue = lTotal
    txtTotalQty = lQty
    
End Sub

Private Sub cmdSave_Click()

    If Saved Then
        FormProcessID = cmbProcess.ID
        'Unload Me
        'Me.Show
        'cmbProcess.ID = cmbProcess.Tag
        'Call ShowIssForm(FormProcessID)
        'frmVendIssItemsByMaker.Show
        LVIssued.ListItems.Clear
        con.Execute "Delete From PrintSession"
        Call cmbItems_Click
    End If
    
End Sub

Private Sub cmdSaveClose_Click()
    
    If Saved Then
        Call PrintIssuanceSlips
        DoEvents
'        If cmbProcess.ID <= 7 Then
'            If Not UserReqAuth Then Call PrintIssRcvRpts("Original", True)
'        End If
        Unload Me
        con.Execute "Delete From PrintSession"
    End If
End Sub

Private Sub PrintIssuanceSlips()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report, strSelection As String
    If bReIssuance Then
        Set rpt = rptApp.OpenReport(RptPath & "\IssSlipGP.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\IssSlip.rpt")
    End If
    strSelection = "ISNULL({PrintSession.RecieptNo})=False"
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdUnAssign_Click()

    On Error GoTo err
    If LV.ListItems.count = 0 Then Exit Sub
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


Private Sub cmdSkipCancel_Click()
    FraSkip.Visible = False
End Sub

Private Sub cmdSkipOk_Click()

    If cmbSkipProcess.MatchFound = False Then
        MsgBox "Please Select Process.", vbInformation
        Exit Sub
    End If
    
    Dim strItemID As String
    
    strItemID = LV.SelectedItem.SubItems(2)
    
    con.Execute "UPDATE VendRcvdDetail SET NextProcessID=" & cmbSkipProcess.ID & " WHERE ItemCode='" & _
     strItemID & "' AND RcvdQty>IssQty AND NextProcessID=" & cmbProcess.ID
     
    Call cmbItems_Click
    FraSkip.Visible = False
End Sub

Private Sub cmdSOCancel_Click()
    FraStockOrder.Visible = False
End Sub

Private Sub cmdSOOK_Click()
    
    If cmbSOItems.MatchFound = False Then
        MsgBox "Please Select Article.", vbInformation
        Exit Sub
    ElseIf Val(txtSOQty) <= 0 Then
        MsgBox "Invalid Qty.", vbInformation
        Exit Sub
    End If
    
    '-----------------------------------------
    'Check If Opening for Selected Article for Selected Process is feeded ?
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(ItemID)", "StockOrderOpening", " WHERE ItemID='" & cmbSOItems.ID & "' AND ProcessID=" & cmbProcess.ID)
'    If lCount > 0 Then
'        MsgBox "This Item has already Opening feeded at this Process.", vbInformation
'        Exit Sub
'    End If
    
    Call StartTrans(con)
    
    'Check Customer...
    lCount = 0
    lCount = GetSingleLongValue("COUNT(CustCode)", "ForeignCustomers", " WHERE CustCode='Stock' AND Country='PK'")
    If lCount = 0 Then
        con.Execute "INSERT INTO ForeignCustomers(Custcode,Country) VALUES('Stock','PK')"
    End If
    
    'Now Check to see if Stock Order is feeded.
    lCount = 0
    lCount = GetSingleLongValue("COUNT(OrderNo)", "FCustomerOrders", " WHERE OrderNo='Stock-Order'")
    If lCount = 0 Then
        con.Execute "INSERT INTO FCustomerOrders(CustCode,Country,OrderNo,DT) VALUES('Stock','PK','Stock-Order','" & Date & "')"
    End If
    
    'Now Check If Article is Assigned to Customer
    lCount = 0
    lCount = GetSingleLongValue("COUNT(CompItemID)", "FCustomerCatalog", " WHERE Custcode='Stock' AND Country='PK' AND CompItemID='" & cmbSOItems.ID & "'")
    If lCount = 0 Then
        con.Execute "INSERT INTO FCustomerCatalog(Custcode,Country,ItemID,CompItemID) VALUES('Stock','PK','" & cmbSOItems.ID & "','" & cmbSOItems.ID & "')"
    End If
    'Now Check If Article is feeded against this Article
    lCount = 0
    lCount = GetSingleLongValue("COUNT(OrderNo)", "FOrderItems", " WHERE OrderNo='Stock-Order' AND CompItemCode='" & cmbSOItems.ID & "'")
    If lCount = 0 Then
        con.Execute "INSERT INTO FOrderItems(OrderNo,ItemCode,CompItemCode,Qty) VALUES('Stock-Order','" & cmbSOItems.ID & _
         "','" & cmbSOItems.ID & "'," & Val(txtSOQty) & ")"
    Else
        con.Execute "UPDATE FOrderItems SET Qty=Qty+" & Val(txtSOQty) & " WHERE OrderNo='Stock-Order' AND CompItemCode='" & cmbSOItems.ID & "'"
    End If
    
    'Now Check If Article is feeded in Stock Order
    lCount = 0
    lCount = GetSingleLongValue("MAX(EntryID)", "StockOrderOpening", " WHERE ItemID='" & cmbSOItems.ID & "' AND ProcessID=" & Val(cmbProcess.ID))
    If lCount = 0 Then
        Dim strLocation As String, lQty As Long, lEntryID As Long
        strLocation = "": lQty = Val(txtSOQty)
        con.Execute "INSERT INTO StockOrderOpening(ItemID,ProcessID,Qty,Location,UserName,MachineName) VALUES('" & cmbSOItems.ID & _
         "'," & cmbProcess.ID & "," & lQty & ",'" & strLocation & "','" & CurrentUserName & "','" & strComputerName & "')"
         
        lEntryID = GetSingleLongValue("MAX(EntryID)", "StockOrderOpening")
    Else
        lEntryID = lCount
        'con.Execute "UPDATE StockOrderOpening SET Qty=Qty+" & Val(txtQty)
    End If
    
    Dim lRcvEntryID As Long, lPrevProcessID As Long
    
    
    con.Execute "INSERT INTO VendReceived(VendID,DT,RecieptID,UserID,ProcessID,Issuance_RefID) VALUES(" & _
     cmbVender.ID & ",'" & Date & "','',0," & cmbProcess.ID & ",0)"
    
    lRcvEntryID = GetSingleLongValue("MAX(EntryID)", "VendReceived")
    
    lPrevProcessID = getPrevProcessID(cmbSOItems.ID, cmbProcess.ID)
    
    con.Execute "INSERT INTO VendRcvdDetail(RefID,ItemCode,RecieptID,RcvdQty,Rate,LotNo,NextProcessID,OrderNo" & _
     ",CountedBy,Issue_RefID,ProcessID,RcvdWeight,Opening_RefID) VALUES(" & lRcvEntryID & ",'" & cmbSOItems.ID & _
     "',''," & lQty & ",0,'0'," & cmbProcess.ID & ",'Stock-Order','',0," & lPrevProcessID & _
     ",0," & lEntryID & ")"
     
    con.CommitTrans
    '-----------------------------------------
    
    FraStockOrder.Visible = False
    Call cmbVender_matched
    
    Exit Sub
    
    
err:
    MsgBox err.Description
End Sub

Private Sub cmdUpdateRate_Click()
    If cmbVender.MatchFound = False Then
        MsgBox "No Maker Selected.", vbInformation
        Exit Sub
    ElseIf LV.SelectedItem Is Nothing Then
        MsgBox "No Item Selected.", vbInformation
        Exit Sub
    ElseIf cmbProcess.MatchFound = False Or cmbProcess.ID = "0" Then
        MsgBox "No Process Selected.", vbInformation
        Exit Sub
    End If
    con.Execute "UPDATE VendAssItems SET Rate=" & Val(txtVendRate) & " WHERE VendID=" & cmbVender.ID & " AND ItemID='" & _
     LV.SelectedItem.SubItems(2) & "' AND ProcessID=" & cmbProcess.ID
End Sub

Private Sub DTPicker1_Change()
    Call DTPicker1_Click
End Sub

Private Sub DTPicker1_Click()
    On Error GoTo err
        If cmbVender.MatchFound = False Then
        MsgBox "No Maker Selected.", vbInformation
        Exit Sub
    ElseIf LV.SelectedItem Is Nothing Then
        MsgBox "No Item Selected.", vbInformation
        Exit Sub
    ElseIf cmbProcess.MatchFound = False Or cmbProcess.ID = "0" Then
        MsgBox "No Process Selected.", vbInformation
        Exit Sub
    End If
    con.Execute "UPDATE VendAssItems SET Rate=" & Val(txtVendRate) & " WHERE VendID=" & cmbVender.ID & " AND ItemID='" & _
     LV.SelectedItem.SubItems(2) & "' AND ProcessID=" & cmbProcess.ID
    If cmbVender.MatchFound = False Then
        MsgBox "No Maker Selected.", vbInformation
        Exit Sub
    ElseIf LV.SelectedItem Is Nothing Then
        MsgBox "No Item Selected.", vbInformation
        Exit Sub
    ElseIf cmbProcess.MatchFound = False Or cmbProcess.ID = "0" Then
        MsgBox "No Process Selected.", vbInformation
        Exit Sub
    End If
    con.Execute "UPDATE VendAssItems SET Rate=" & Val(txtVendRate) & " WHERE VendID=" & cmbVender.ID & " AND ItemID='" & _
     LV.SelectedItem.SubItems(2) & "' AND ProcessID=" & cmbProcess.ID

    
    With rsMain
    
        If DTPicker1.Value & "" = "" Then
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
            .MoveNEXT
        Loop
    End With
    
    
        
    
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub DTPicker2_Change()
    DTReturn.Value = DateAdd("d", 30, DTPicker2.Value)
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


Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If bReIssuance Then
        If strCompany = "Sunlike" Then
            Me.PopUpMenu mnuPop
        End If
    End If
End Sub

Private Sub LVIssued_DblClick()
    
    
    If LVIssued.ListItems.count = 0 Then Exit Sub
    'If cmbCust.MatchFound = False Then Exit Sub
    If LV.ColumnHeaders(3).Text = "Lot No/Order No" And RcvingLot Then Exit Sub
     
    With txtRate1
        cmdPost.Default = False
        .Move LVIssued.Left + LVIssued.ColumnHeaders(iColNo).Left, LVIssued.Top + LVIssued.SelectedItem.Top, LVIssued.ColumnHeaders(iColNo).Width
        .Text = LVIssued.SelectedItem.SubItems(iColNo - 1) 'LVMaker.SelectedItem.SubItems(3)
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
        
    End With

End Sub

Private Sub LVIssued_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo err
    
    If KeyCode = 46 Then
    
'        If RcvingLot Then
'            LVIssued.ListItems.Remove LVIssued.SelectedItem.Index
'            Exit Sub
'        End If
        
        'MyCol(LVIssued.SelectedItem.Tag).IssueableQty = MyCol(LVIssued.SelectedItem.Tag).IssueableQty + Val(LVIssued.SelectedItem.SubItems(7))
        LV.ListItems(LVIssued.SelectedItem.Tag).ListSubItems(4) = LV.ListItems(LVIssued.SelectedItem.Tag).ListSubItems(4).Tag
        'LV.ListItems(LVIssued.SelectedItem.Tag).SubItems(5) = MyCol(LVIssued.SelectedItem.Tag).IssueableQty
        LVIssued.ListItems.Remove LVIssued.SelectedItem.Index
        Call LV_ItemClick(LV.SelectedItem)
        
        'Reset The IssID's
        Dim TempID As Long
        TempID = 0
        If LVIssued.ListItems.count > 0 Then
            For i = LVIssued.SelectedItem.Index To LVIssued.ListItems.count
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
        Call getTotalValue
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub LVIssued_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If x > LVIssued.ColumnHeaders(10).Left Then
        iColNo = 10
    ElseIf x > LVIssued.ColumnHeaders(9).Left Then
        iColNo = 9
    Else
        iColNo = 8
    End If
End Sub

Private Sub LvMaker_DblClick()
    
    Exit Sub
    If LVMaker.ListItems.count = 0 Then Exit Sub
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
'   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    
    On Error GoTo err
    SSTab1.TabVisible(4) = False
    'Call AddToCombo(cmbOrderNo, "OrderNo", "FCustomerOrders", " Order By OrderNo")
    mnuPop.Visible = False
    bReIssuance = False
    
    'cmbCust.ListHeight = 2600
    'cmbCustName.ListHeight = 2600
    'cmbOrders.ListHeight = 2600
    
    'cmbProcess.Enabled = True
    bFromMaker = False
    cmbProcess.ListHeight = 6000
     
    DTPicker2.Value = ServerDate
    Call DTPicker2_Change
    DTPicker1.Value = ServerDate
    DTPicker1.Value = ""
    DTReturn.Value = DateAdd("d", 30, DTPicker2.Value)
    SSTab1.Tab = 2
    IsLotProcess = False
    cmbVender.ListHeight = 2000
    GotIssID = 0
    
    If CurrentUserName <> "Administrator" Then
        bCanChangeRate = GetSingleBooleanValue("ChangeRateonIssuance", "Users", " WHERE UserID=" & UserID)
    Else
        bCanChangeRate = True
    End If
    
    If bCanChangeRate Then
        txtRate.Locked = False
        cmdUpdateRate.Enabled = True
    Else
        txtRate.Locked = True
        cmdUpdateRate.Enabled = False
    End If
        
    cmbMaterial.ListHeight = 2000
    cmbMaterial.AddVals con, "RMName + ' {' + RMID1 + '}'", "RM", "RMID", " ORDER BY RMID"
    
    Me.Hide
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Public Sub LockProc(ProcID As Long)
    
    
    cmbProcess.Tag = ProcID
    
    'cmbProcess.Enabled = False
    'Me.Refresh
    
    cmbProcess.ID = ProcID
    Me.Show
    
End Sub
Public Sub RefreshLV()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "Select *  from FCustomerCatalog Where CustCode='" & cmbCust.ID & "' and Country='" & cmbCountry.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, ![CustCode] & ![Country] & ![ITemID], ![ITemID] & "")
            ITM.SubItems(1) = ![Description] & ""
            ITM.SubItems(2) = ![Unit] & ""
            ITM.SubItems(3) = ![FOB] & ""
            ITM.SubItems(4) = ![CIFSea] & ""
            ITM.SubItems(5) = ![CIFAir] & ""
            ITM.SubItems(6) = ![CnFSea] & ""
            ITM.SubItems(7) = ![CnFAir] & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub Form_Resize()
    
    Exit Sub
    On Error Resume Next
    
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
    
    Dim TitleBottom As Long

    TitleBottom = DTPicker1.Top + DTPicker1.Height + 50 '+ fTitle.Height
 
    
    LV.Move Label2.Left, TitleBottom
    
    Label3.Move Label2.Left, LV.Top + LV.Height + 50
'    cmbVender.Move Label3.Left, Label3.Top + Label3.Height - 15
    
    
    DTPicker2.Move cmbVender.Left + cmbVender.Width + 20, Label3.Top + Label3.Height - 15
    txtQty.Move DTPicker2.Left + DTPicker2.Width + 10, Label3.Top + Label3.Height - 15
    cmdPost.Move txtQty.Left + txtQty.Width + 10, Label3.Top
    SSTab1.Move LV.Left, cmbVender.Top + cmbVender.Height + 50
    
End Sub

Private Sub FillVendList()

    On Error GoTo err
    
    If LV.ListItems.count = 0 Then Exit Sub
    Dim VendIDList As String
    VendIDList = ""
    For i = 1 To LVIssued.ListItems.count
        With LVIssued.ListItems(i)
            'If LV.SelectedItem.key = .Tag Then
            '    VendIDList = VendIDList & Val(.ListSubItems(1).key) & ","
            'End If
            If LV.SelectedItem.ListSubItems(1) = .SubItems(3) Then
                VendIDList = VendIDList & Val(.ListSubItems(1).Tag) & ","
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
        Call cmbVender.AddVals(con, "Distinct VenderName", "VMakerAssItems", "VendID", " Where (ItemID='" & LV.SelectedItem.SubItems(2) & "' and PhaseID=" & cmbProcess.ID & "  AND Active=1 and VendID Not In" & VendIDList & " and VendType=1 Or (PhaseID=-1))")
    Else
        Call cmbVender.AddVals(con, "Distinct VenderName", "VMakerAssItems", "VendID", " Where (ItemID='" & LV.SelectedItem.SubItems(1) & "' AND Active=1 and (PhaseID=" & cmbProcess.ID & ") and VendID Not In" & VendIDList & " ) OR VendID=129 ")
    End If
    'End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)

'    On Error GoTo err
'    Call FillVendList
    
    'Dim rs As New ADODB.Recordset
    'Dim ITM As ListItem
    
'    Dim VendIDList As String
'    VendIDList = ""
'    For i = 1 To LVIssued.ListItems.count
'        With LVIssued.ListItems(i)
'            If LV.SelectedItem.key = .Tag Then
'                VendIDList = VendIDList & Val(.ListSubItems(1).key) & ","
'            End If
'        End With
'    Next i
'
'    If VendIDList <> "" Then
'        VendIDList = "(" & Left(VendIDList, Len(VendIDList) - 1) & ")"
'    Else
'        VendIDList = "(0)"
'    End If
    
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
    txtQty = LV.SelectedItem.ListSubItems(4)
        
    Dim i As Integer
    For i = 1 To LVIssued.ListItems.count
        If LV.SelectedItem.SubItems(2) = LVIssued.ListItems(i).SubItems(5) Then
            txtSpecialInstructions = LVIssued.ListItems(i).ListSubItems(5).Tag
            Exit For
        End If
    Next
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub



Private Sub mnuSkipProcess_Click()
    ''
    If LV.ListItems.count = 0 Then Exit Sub
    Dim lSNo As Long, strItemID As String
    strItemID = LV.SelectedItem.SubItems(2)
    lSNo = GetSingleLongValue("SNo", "ItemProcesses", " WHERE ItemID='" & strItemID & "' AND ProcessID=" & cmbProcess.ID)
    
    cmbSkipProcess.ClearVals
    cmbSkipProcess.AddVals con, "Description", "Processes", "ProcessID", " WHERE ProcessID IN(SELECT ProcessID FROM ItemProcesses WHERE ItemID='" & strItemID & "' AND SNo>" & lSNo & ") Order By SNo"
    
    FraSkip.Visible = True

End Sub

Private Sub mnuStockOrderIssuance_Click()
    cmbSOItems.ClearVals
    Dim strCondition As String
    strCondition = " WHERE InActive=0 AND ItemID IN(SELECT ItemID FROM VendAssItems WHERE VendID=" & Val(cmbVender.ID) & " AND ProcessID=" & Val(cmbProcess.ID) & ") ORDER BY ItemName"
    Call cmbSOItems.AddVals(con, "ItemID + ' ' + ItemName + ' { ' + ISNULL(Attributes,'') + ' }'", "Items", "ItemID", strCondition)
    
    FraStockOrder.Visible = True
End Sub

Private Sub txtQty_GotFocus()


    With txtQty
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    

End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    
    If KeyAscii = 13 Then
        'LVMaker.SelectedItem.SubItems(3) = Val(txtRate)
        
        Dim IssID As String
        For i = 1 To LVIssued.ListItems.count
            If Val(LVMaker.SelectedItem.ListSubItems(1).key) = Val(LVIssued.ListItems(i).ListSubItems(2).key) Then
                IssID = LVIssued.ListItems(i).ListSubItems(2)
                GoTo GotID:
            End If
        Next i
        
        IssID = Val(con.Execute("Select Max(right(RecieptID,3)) From VendIssued Where DT='" & Format(ServerDate, "dd-MMM-yyyy") & "'").Fields(0).Value & "")
        
        'IssID = "ISS-" & Val(LVMaker.SelectedItem.ListSubItems(1).Key) & "-" & _
         Format(ServerDate, "ddMMyy") & "-" & Format(Val(IssID) + 1, "000")
         IssID = "ISU-" & _
         Format(ServerDate, "ddMMyy") & "-" & Format(Val(IssID) + 1, "000")
GotID:
        
        Dim ITM As ListItem
        With ITM
            'Set ITM = LVIssued.ListItems.Add(, Val(LV.SelectedItem.Key) & Val(LVMaker.SelectedItem.Key) & "'", LVIssued.ListItems.Count + 1)
            'Set ITM = LVIssued.ListItems.Add(, LV.SelectedItem.Key & "|" & LVIssued.ListItems.Count + 1, LVIssued.ListItems.Count + 1)
            Set ITM = LVIssued.ListItems.add(, , LVIssued.ListItems.count + 1)
            ITM.Tag = LV.SelectedItem.key
            ITM.ListSubItems.add , LVMaker.SelectedItem.ListSubItems(1).key, LVMaker.SelectedItem.SubItems(1)
            ITM.ListSubItems.add , , IssID & ""
            ITM.ListSubItems.add , , LV.SelectedItem.SubItems(2)
            ITM.ListSubItems.add , , LV.SelectedItem.SubItems(3)
            ITM.ListSubItems.add , , LV.SelectedItem.SubItems(4)
            ITM.ListSubItems.add , , LVMaker.SelectedItem.ListSubItems(2)
            ITM.ListSubItems.add , , Val(txtRate)
            If cmbProcess.Tag = "Lot No" Then
                ITM.ListSubItems(6).Tag = LV.SelectedItem.SubItems(2) & ""
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
    If Val(txtRcvings) <= 0 Then
        MsgBox "Please Enter Max. Rcvings.", vbInformation
        Exit Function
    End If
    
    Dim SaveIt As Boolean
    
    Dim VendID As New Collection
    Dim VIssID As New Collection
     
'    Dim strVendnItem() As String
'    Dim iVendnItemCount As Integer
    
    Dim myPOSavedCol As New Collection
    
    Dim cmd As New ADODB.Command
    Dim TempIndex As Integer
    TempIndex = 1
    
    Dim lRefID As Long
    
    Dim IssLot As Boolean
    Dim iAuth As Integer 'Authorize
    
    IssLot = RcvingLot 'RcvLot(cmbProcess.ID)
    
'    With cmd
'        .CommandType = adCmdStoredProc
'        .CommandText = "SP_ReqAuthForIss"
'        Set .ActiveConnection = con
'        .Parameters("@UserID").value = UserID
'        .Parameters("@ProcessID").value = cmbProcess.ID
'        .Execute
'        UserReqAuth = Val(.Parameters("@ReqAuth").value & "")
'    End With
    UserReqAuth = 0
    Dim iSteelType As Integer
    If cmbSteel.MatchFound = False Then
        iSteelType = 0
    Else
        iSteelType = cmbSteel.List(cmbSteel.ListIndex, 1)
    End If
    
    Call StartTrans(con)
    
    For i = 1 To LVIssued.ListItems.count
        With LVIssued.ListItems(i)
            SaveIt = True
'            For j = 1 To VendID.count
'                If Val(LVIssued.ListItems(i).ListSubItems(1).Tag) = VendID(j) Then  'Check Vender
'                    'It Means That It's Been Saved Already.
'                        SaveIt = False
'                        Exit For
'                End If
'            Next j
'            If i = 1 Then
'                iVendnItemCount = 0
'            Else
'                iVendnItemCount = UBound(strVendnItem)
'            End If
'            For j = 0 To iVendnItemCount - 1
'                If strVendnItem(j, 0) = Val(LVIssued.ListItems(i).ListSubItems(1).Tag) And strVendnItem(j, 1) = .SubItems(5) Then
'                    SaveIt = False
'                    Exit For
'                End If
'            Next
            Dim myPOObjTemp As clsPOEntry
            For j = 1 To myPOSavedCol.count
                Set myPOObjTemp = myPOSavedCol(j)
                If myPOObjTemp.iIssueProcess = Val(.ListSubItems(7).Tag) And myPOObjTemp.iReturnProcess = Val(.ListSubItems(3).Tag) And myPOObjTemp.iVendID = Val(.ListSubItems(1).Tag) And myPOObjTemp.strItemCode = .SubItems(5) Then
                    SaveIt = False
                    lRefID = myPOObjTemp.lRefID
                    Set myPOObjTemp = Nothing
                    Exit For
                End If
                Set myPOObjTemp = Nothing
            Next
            
            If SaveIt Then
                Dim lMakerIssNo As Long, strMakerIssNo As String
                lMakerIssNo = Val(con.Execute("Select Max(Cast(right(RecieptID,Len(RecieptID)-12) As Int)) From VendIssued Where Convert(varchar,DT,6)='" & Format(DTPicker2, "dd MMM yy") & "'").Fields(0).Value & "")
                lMakerIssNo = lMakerIssNo + 1
                IssID = "ISU-" & Format(DTPicker2, "ddMMyy") & GotIssID
                strMakerIssNo = "M-ISU-" & Format(DTPicker2, "ddMMyy") & lMakerIssNo
                
                con.Execute "INSERT INTO VendIssued(VendID,DT,RecieptID,UserID,ProcessID,ItemID," & _
                "UserName,MachineName,SpecialInstructions,ExcessQtyPercentage," & _
                "MaximumRcvingsAgainstPO,SteelType_RefID) VALUES(" & Val(LVIssued.ListItems(i).ListSubItems(1).Tag) & ",'" & _
                 DTPicker2.Value & "','" & strMakerIssNo & "'," & UserID & "," & _
                 cmbProcess.ID & ",'" & .SubItems(5) & "','" & CurrentUserName & "','" & _
                 strComputerName & "','" & .ListSubItems(5).Tag & "'," & Val(txtExcess) & "," & _
                 Val(txtRcvings) & "," & iSteelType & ")"
                 
                lRefID = Val(con.Execute("SELECT MAX(EntryID) FROM VendIssued").Fields(0).Value & "")
                
                'Add to Collection After Assigning to Object
                Dim myPOObj As New clsPOEntry
                myPOObj.iIssueProcess = Val(.ListSubItems(7).Tag)
                myPOObj.iReturnProcess = Val(.ListSubItems(3).Tag)
                myPOObj.iVendID = Val(.ListSubItems(1).Tag)
                myPOObj.strItemCode = .SubItems(5)
                myPOObj.lRefID = lRefID
                myPOSavedCol.add myPOObj
                Set myPOObj = Nothing
                
'                ReDim Preserve strVendnItem(iVendnItemCount + 1, 3)
'                ReDim Preserve strVendnItem(iVendnItemCount + 1)
'                ReDim Preserve strVendnItem(iVendnItemCount)
'                strVendnItem(iVendnItemCount, 0) = Val(LVIssued.ListItems(i).ListSubItems(1).Tag)
'                strVendnItem(iVendnItemCount, 1) = .SubItems(5)
'                strVendnItem(iVendnItemCount, 2) = lRefID
                
                VendID.add Val(LVIssued.ListItems(i).ListSubItems(1).Tag)
                VIssID.add lRefID
            Else
                'lRefID = VIssID(j)
                'lRefID = 'strVendnItem(j, 2)
            End If
        
            '1. Insert Into VendIssuedItems
            
            If cmbProcess.Tag = "Lot No" And IssLot Then
                With cmd
                    .ActiveConnection = con
                    .CommandType = adCmdStoredProc
                    .CommandText = "SP_UpdateForIARLots"
                    
                    .Parameters("@ProcessID").Value = cmbProcess.ID
                    .Parameters("@Qty").Value = Val(LVIssued.ListItems(i).SubItems(6))
                    .Parameters("@LotNo").Value = LVIssued.ListItems(i).ListSubItems(6).Tag
                    .Parameters("@RefiD").Value = lRefID
                    .Parameters("@Rate").Value = Val(LVIssued.ListItems(i).SubItems(5))
                    .Parameters("@RecID").Value = LVIssued.ListItems(i).SubItems(2)
                    .Parameters("@ReqAuth").Value = UserReqAuth
                    
                    .Execute
                End With
                
            Else
                con.Execute "INSERT INTO VendIssdDetail(RefID,RecieptID,ItemCode" & _
                 ",Rate,IssQty,ReqAuth,OrderNo,RcvProcessID,ReturnDT,Priority) Values(" & lRefID & ",'" & _
                 .ListSubItems(2).Tag & "','" & .SubItems(5) & "'," & _
                 Val(.SubItems(7)) & "," & Val(.SubItems(8)) & "," & UserReqAuth & ",'" & _
                 .SubItems(4) & "'," & .ListSubItems(3).Tag & ",'" & .SubItems(6) & "'," & .SubItems(10) & ")"
                 
                 'Do Not Execute Stored Procedure
                 'If This User Requires Authorization...
                 If UserReqAuth = 0 Then
                    With cmd
                        .ActiveConnection = con
                        .CommandType = adCmdStoredProc
                        .CommandText = "SP_UpdateForIAR"    'Issue After Receive
                        
                        .Parameters("@ItemCode").Value = LVIssued.ListItems(i).SubItems(5)
                        .Parameters("@ProcessID").Value = cmbProcess.ID
                        .Parameters("@Qty").Value = Val(LVIssued.ListItems(i).SubItems(8))
                        .Parameters("@OrderNo").Value = LVIssued.ListItems(i).SubItems(4)
                        
                        .Execute
                    End With
                    
                End If
            End If
            'Authentication Info...
            If UserReqAuth = 0 Then
                Dim DetailRefID As Long
                DetailRefID = Val(con.Execute("SELECT MAX(EntryID) FROM VendIssdDetail").Fields(0).Value & "")
                con.Execute "INSERT INTO VendIssAuthDetail(UserID,IssID,DT) Values(" & _
                 UserID & "," & DetailRefID & ",'" & ServerDate & " " & ServerTime & "')"
            End If
            
        End With
        con.Execute "Insert Into PrintSession(RecieptNo) Values('" & LVIssued.ListItems(i).ListSubItems(2).Tag & "')"
        'Debug.Print i & " of " & LVIssued.ListItems.count
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
        For i = 1 To LVIssued.ListItems.count
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
    
    'Now Save RM. Issuance
    'Call SaveRMIssuance
    con.CommitTrans
    
    Saved = True
    Exit Function
err:
    MsgBox err.Description
    
End Function

Private Sub SaveRMIssuance()
    
    Dim rs As New ADODB.Recordset, strIssNo As String
    With rs
        .Open "Select Max(Cast(Right(IssNo,3) As Int)) From IssItemsSimple Where Convert(Char(8),DT,1)='" & Format(Date, "MM/dd/yy") & "'", con, adOpenForwardOnly, adLockReadOnly
        
        strIssNo = "ISS-" & Format(Date, "yyMMdd-") & Format(Val(.Fields(0) & "") + 1, "000")
        
        .Close
    End With
    Set rs = Nothing
    
    con.Execute "INSERT INTO RawMaterialIssuance(IssNo,DT,UserName,ComputerName) VALUES('" & _
         strIssNo & "','" & DT & "','" & CurrentUserName & "','" & strComputerName & "')"
      
    Dim strMaterialID As String, dRate As Double
    For i = 1 To LVRMIss.ListItems.count
        With LVRMIss.ListItems(i)
            strMaterialID = GetSingleStringValue("RMID1", "RM", " WHERE RMID=" & Val(.key))
            dQty = Val(.SubItems(5))
            If Val(.ListSubItems(3).Tag) = 0 Then
                dRate = GetSingleDoubleValue("Rate", "RM", " WHERE RMID=" & Val(.key))
                '-------------------------------------------------------------------------
                con.Execute "INSERT INTO RawMaterialIssuanceDetail(IssNo,RMID1,Qty,InHand,StoreName,Price) VALUES('" & _
                 strIssNo & "','" & strMaterialID & "'," & dQty & "," & Val(.Tag) - dQty & ",'" & .SubItems(3) & _
                 "'," & dRate & ")"
                 
                con.Execute "UPDATE RM SET QtyInStock=QtyInStock-" & dQty & " WHERE RMID=" & Val(.key)
            End If
        End With
    Next
    
End Sub
Private Sub txtRate1_KeyDown(KeyCode As Integer, Shift As Integer)
'    MsgBox KeyCode
End Sub

Private Sub txtRate1_KeyPress(KeyAscii As Integer)

    On Error GoTo err
    Dim iQty  As Long
    If KeyAscii = 13 Then
        'LVMaker.SelectedItem.SubItems(3) = Val(txtRate)
        If Val(txtRate1) <= 0 Then
            MsgBox "Invalid Data.", vbInformation
            Exit Sub
        End If
        Dim ITM As ListItem
        If iColNo = 9 Then
            iQty = LV.ListItems(LVIssued.SelectedItem.Tag).ListSubItems(4).Tag - Val(txtRate1)
            If iQty < 0 Then
                MsgBox "Invalid Qty!", vbInformation
                Exit Sub
            End If
            LV.ListItems(LVIssued.SelectedItem.Tag).SubItems(4) = iQty
            LVIssued.SelectedItem.SubItems(iColNo - 1) = Val(txtRate1)
        Else
            LVIssued.SelectedItem.SubItems(iColNo - 1) = Val(txtRate1)
        End If
        
        txtRate1.Visible = False
        Call LV_ItemClick(LV.SelectedItem)
        'End If
'        If Val(txtRate1) = 0 Then
'            Call LVIssued_KeyDown(46, 0)
'        End If
        
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
    'Me.KeyPreview = True
    cmdPost.Default = True
End Sub
Private Sub txtRate_LostFocus()
    txtRate.Visible = False
End Sub
Private Sub PostForLots()

    On Error GoTo err
    Dim ITM As ListItem

    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            IssID = LV.ListItems(i).Tag
            With ITM
                Set ITM = LVIssued.ListItems.add(, , LVIssued.ListItems.count + 1)
                ITM.Tag = LV.ListItems(i).key
                ITM.ListSubItems.add(, cmbVender.ID & "'", cmbVender.Text).Tag = cmbVender.ID
                ITM.ListSubItems.add(, , IssID & "").Tag = DTPicker2
                ITM.ListSubItems.add , , LV.ListItems(i).SubItems(2)
                ITM.ListSubItems.add , , LV.ListItems(i).SubItems(3)
                ITM.ListSubItems.add , , Val(cmbVender.Tag)
                ITM.ListSubItems.add , , LV.ListItems(i).SubItems(4)
                'itm.ListSubItems.Add , , Val(LV.ListItems(i).SubItems(5))
                If cmbProcess.Tag = "Lot No" Then
                    ITM.ListSubItems(6).Tag = LV.ListItems(i).ListSubItems(1).Tag
                End If
            End With
            LV.ListItems(i).SubItems(4) = 0
            
        End If
    Next
    
    
    'Remove Posted Entries From the List.
    i = 1
    Do While i <= LV.ListItems.count
        If LV.ListItems(i).Checked Then
            LV.ListItems.Remove i
            On Error Resume Next
            cmbDesc.RemoveItem i - 1
            If err.Number = -2147024809 Then Resume Next
                
            i = i - 1
        End If
        i = i + 1
    Loop
    
    'if cmbitems.
    'Call LV_ItemClick(LV.ListItems(i))
    
    Call FillVendList
    
    txtQty = ""
    Exit Sub
err:
    MsgBox err.Description

End Sub

Public Sub ShowMe(m_bReIssuance As Boolean)

    bReIssuance = m_bReIssuance
    
    If bReIssuance = False Then
        cmbProcess.AddItem "<All Processes>", "0"
        txtSpecialInstructions.Width = 8415
        cmbSteel.Visible = False
        lblSteel.Visible = False
    Else
        label1(2).Visible = False
        txtExcess.Text = "0"
        txtExcess.Visible = False
        txtSpecialInstructions.Width = 4215
        lblSteel.Visible = True
        cmbSteel.Visible = True
    End If
    
    cmbProcess.AddVals con, "Description", "Processes", "ProcessID", " Order By SNo"
    If bReIssuance = False Then cmbProcess.ID = "0"
   
    Me.Show
    
End Sub

