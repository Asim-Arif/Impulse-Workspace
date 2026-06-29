VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmMakerBillingRcvingWise 
   ClientHeight    =   10650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   10650
   ScaleMode       =   0  'User
   ScaleWidth      =   10721.69
   WindowState     =   2  'Maximized
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   60
      TabIndex        =   46
      Top             =   -90
      Width           =   11865
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Maker Billing"
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
         Index           =   3
         Left            =   4785
         TabIndex        =   47
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Maker Billing"
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
         Index           =   4
         Left            =   4800
         TabIndex        =   48
         Top             =   135
         Width           =   2175
      End
   End
   Begin VB.Frame FAcc 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   10140
      Left            =   60
      TabIndex        =   8
      Top             =   450
      Width           =   11865
      Begin VB.TextBox txtLTBalance 
         Alignment       =   2  'Center
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   9060
         Width           =   1455
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
         Left            =   9810
         TabIndex        =   69
         Top             =   9195
         Width           =   1560
      End
      Begin VB.TextBox txtTotalAudited 
         Alignment       =   2  'Center
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3180
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   9060
         Width           =   1455
      End
      Begin VB.TextBox txtTotalSelected 
         Alignment       =   2  'Center
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   9060
         Width           =   1455
      End
      Begin VB.TextBox txtTotalLots 
         Alignment       =   2  'Center
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   195
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   9045
         Width           =   1455
      End
      Begin ComboList.Usercontrol1 cmbCats 
         Height          =   285
         Left            =   6105
         TabIndex        =   5
         Top             =   855
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   503
         BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextBackColor   =   16777215
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListIndex       =   -1
         Appearance      =   0
      End
      Begin VB.TextBox txtMasterPONo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9840
         TabIndex        =   57
         Top             =   855
         Width           =   960
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   195
         TabIndex        =   0
         Top             =   345
         Width           =   5880
         _ExtentX        =   10372
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
      Begin ComboList.Usercontrol1 cmbProcess 
         Height          =   285
         Left            =   195
         TabIndex        =   4
         Top             =   855
         Width           =   5880
         _ExtentX        =   10372
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
      Begin VB.CheckBox chkRejectionRelaxation 
         Caption         =   "Allow 2 % Rejection Relaxation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   195
         TabIndex        =   53
         Top             =   9645
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtValue 
         Alignment       =   2  'Center
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   7950
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   8430
         Width           =   1455
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   6780
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   8430
         Width           =   1155
      End
      Begin ComboList.Usercontrol1 cmbDebitHead 
         Height          =   285
         Left            =   195
         TabIndex        =   7
         Top             =   1395
         Width           =   11520
         _ExtentX        =   20320
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
      Begin VB.TextBox txtManualDed 
         Alignment       =   2  'Center
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   5565
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   8430
         Width           =   1200
      End
      Begin VB.TextBox txtMaterialDed 
         Alignment       =   2  'Center
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   4275
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   8430
         Width           =   1275
      End
      Begin VB.TextBox txtRepairAmt 
         Alignment       =   2  'Center
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2895
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   8430
         Width           =   1365
      End
      Begin VB.TextBox txtFixLoan 
         Alignment       =   2  'Center
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1665
         TabIndex        =   29
         Top             =   8430
         Width           =   1215
      End
      Begin VB.TextBox txtDeductShort 
         Alignment       =   2  'Center
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   195
         TabIndex        =   27
         Top             =   8430
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPosting 
         Height          =   315
         Left            =   9420
         TabIndex        =   24
         Top             =   8430
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
         Format          =   159973379
         CurrentDate     =   37055
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
         ItemData        =   "frmMakerBillingRcvingWise.frx":0000
         Left            =   6105
         List            =   "frmMakerBillingRcvingWise.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   345
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   9345
         TabIndex        =   3
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
         Format          =   159973379
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   7815
         TabIndex        =   2
         Top             =   345
         Width           =   1530
         _ExtentX        =   2699
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
         Format          =   159973379
         CurrentDate     =   37055
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6420
         Left            =   210
         TabIndex        =   13
         Top             =   1755
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   11324
         _Version        =   393216
         Style           =   1
         Tabs            =   6
         TabsPerRow      =   6
         TabHeight       =   520
         TabCaption(0)   =   "Maker Billing"
         TabPicture(0)   =   "frmMakerBillingRcvingWise.frx":0066
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label2(2)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "LV"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "txtEdit"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "cmdSelectAll"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cmdUnSelectAll"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "txtLotNo"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "cmdMarkAsAudited"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "cmdUnmarkAudited"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "cmdSelectAudited"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "PicProduct"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).ControlCount=   10
         TabCaption(1)   =   "Repair Wastage Details"
         TabPicture(1)   =   "frmMakerBillingRcvingWise.frx":0082
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdPostRepair"
         Tab(1).Control(1)=   "LVRepair"
         Tab(1).Control(2)=   "ChkAll_RWD"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "Material Deductions"
         TabPicture(2)   =   "frmMakerBillingRcvingWise.frx":009E
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "LVMaterial"
         Tab(2).Control(1)=   "txtEditMaterialRate"
         Tab(2).Control(2)=   "cmdPrintMaterialDeductions"
         Tab(2).ControlCount=   3
         TabCaption(3)   =   "Manual Deductions"
         TabPicture(3)   =   "frmMakerBillingRcvingWise.frx":00BA
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label1(1)"
         Tab(3).Control(1)=   "LVManualMaterial"
         Tab(3).Control(2)=   "txtManualMatDesc"
         Tab(3).Control(3)=   "txtManualMatRate"
         Tab(3).ControlCount=   4
         TabCaption(4)   =   "Hidden Lots"
         TabPicture(4)   =   "frmMakerBillingRcvingWise.frx":00D6
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "LVHL"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "RM Hidden"
         TabPicture(5)   =   "frmMakerBillingRcvingWise.frx":00F2
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "LVRMHidden"
         Tab(5).ControlCount=   1
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
            Left            =   8640
            ScaleHeight     =   2835
            ScaleWidth      =   2760
            TabIndex        =   70
            Top             =   390
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
         Begin VB.CommandButton cmdSelectAudited 
            Caption         =   "Select Audited"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5985
            TabIndex        =   61
            Top             =   6045
            Width           =   1680
         End
         Begin VB.CommandButton cmdUnmarkAudited 
            Caption         =   "Un-Mark Audited"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4260
            TabIndex        =   59
            Top             =   6045
            Width           =   1680
         End
         Begin VB.CheckBox ChkAll_RWD 
            Height          =   195
            Left            =   -74940
            TabIndex        =   56
            Top             =   570
            Width           =   300
         End
         Begin VB.CommandButton cmdMarkAsAudited 
            Caption         =   "Mark Audited"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2520
            TabIndex        =   52
            Top             =   6045
            Width           =   1680
         End
         Begin VB.TextBox txtLotNo 
            Alignment       =   2  'Center
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   9975
            TabIndex        =   51
            Top             =   6030
            Width           =   1455
         End
         Begin VB.CommandButton cmdPrintMaterialDeductions 
            Caption         =   "Print"
            Height          =   330
            Left            =   -64710
            TabIndex        =   41
            Top             =   6015
            Width           =   1125
         End
         Begin VB.TextBox txtEditMaterialRate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -73185
            TabIndex        =   39
            Top             =   1695
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.CommandButton cmdUnSelectAll 
            Caption         =   "Clear All"
            Height          =   300
            Left            =   1290
            TabIndex        =   38
            Top             =   6045
            Width           =   1170
         End
         Begin VB.CommandButton cmdSelectAll 
            Caption         =   "Select All"
            Height          =   300
            Left            =   90
            TabIndex        =   37
            Top             =   6045
            Width           =   1170
         End
         Begin VB.TextBox txtManualMatRate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   -64585
            TabIndex        =   20
            Top             =   630
            Width           =   1000
         End
         Begin VB.TextBox txtManualMatDesc 
            Appearance      =   0  'Flat
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   -74925
            TabIndex        =   19
            Top             =   630
            Width           =   10350
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2325
            TabIndex        =   14
            Top             =   1485
            Visible         =   0   'False
            Width           =   690
         End
         Begin MSComctlLib.ListView LV 
            Height          =   5565
            Left            =   90
            TabIndex        =   15
            Top             =   390
            Width           =   11340
            _ExtentX        =   20003
            _ExtentY        =   9816
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   8388608
            BackColor       =   15199215
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   19
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S. #"
               Object.Width           =   873
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Rcv #"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Date"
               Object.Width           =   1455
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Order #"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Item ID"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Description"
               Object.Width           =   1879
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "Issd Qty"
               Object.Width           =   1349
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "Price"
               Object.Width           =   1111
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   8
               Text            =   "Value"
               Object.Width           =   1085
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   9
               Text            =   "Temp. Adv"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   10
               Text            =   "Repair"
               Object.Width           =   1109
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   11
               Text            =   "Rcvd. Value"
               Object.Width           =   1667
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   12
               Text            =   "Rcvd Qty"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   13
               Text            =   "Un. Auth."
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   14
               Text            =   "Bal Value"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   15
               Text            =   "% Rcvd"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   16
               Text            =   "Lot #"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   17
               Text            =   "Process"
               Object.Width           =   2778
            EndProperty
            BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   18
               Text            =   "Cust"
               Object.Width           =   1499
            EndProperty
         End
         Begin MSComctlLib.ListView LVRepair 
            Height          =   5445
            Left            =   -74940
            TabIndex        =   16
            Top             =   885
            Width           =   11340
            _ExtentX        =   20003
            _ExtentY        =   9604
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
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
               Object.Width           =   873
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "PO #"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Date"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Item ID"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Qty"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Rate"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Repair Amt"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "Wastage Amt"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Lot No."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Process"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView LVMaterial 
            Height          =   5550
            Left            =   -74925
            TabIndex        =   18
            Top             =   405
            Width           =   11340
            _ExtentX        =   20003
            _ExtentY        =   9790
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
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
               Object.Width           =   873
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Iss #"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Date"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Raw Material"
               Object.Width           =   4586
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Item Code"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Rate"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "Qty"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "Amount"
               Object.Width           =   1764
            EndProperty
         End
         Begin MSComctlLib.ListView LVManualMaterial 
            Height          =   5430
            Left            =   -74925
            TabIndex        =   21
            Top             =   900
            Width           =   11340
            _ExtentX        =   20003
            _ExtentY        =   9578
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S. #"
               Object.Width           =   873
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Amount"
               Object.Width           =   1764
            EndProperty
         End
         Begin MSComctlLib.ListView LVHL 
            Height          =   5700
            Left            =   -74895
            TabIndex        =   60
            Top             =   570
            Width           =   11280
            _ExtentX        =   19897
            _ExtentY        =   10054
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
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Lot No."
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Qty"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Item Code"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Item Name"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Order No."
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Customer"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Process"
               Object.Width           =   3175
            EndProperty
         End
         Begin MSComctlLib.ListView LVRMHidden 
            Height          =   5550
            Left            =   -74940
            TabIndex        =   68
            Top             =   615
            Width           =   11340
            _ExtentX        =   20003
            _ExtentY        =   9790
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
               Object.Width           =   873
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Iss #"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Date"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Raw Material"
               Object.Width           =   4586
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Item Code"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Rate"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "Qty"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "Amount"
               Object.Width           =   1764
            EndProperty
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Lot No. :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   195
            Index           =   2
            Left            =   9255
            TabIndex        =   50
            Top             =   6090
            Width           =   660
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"frmMakerBillingRcvingWise.frx":010E
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
            Left            =   -74925
            TabIndex        =   22
            Top             =   390
            Width           =   11340
         End
         Begin MSForms.CommandButton cmdPostRepair 
            Height          =   375
            Left            =   -74925
            TabIndex        =   17
            Top             =   5310
            Visible         =   0   'False
            Width           =   2310
            ForeColor       =   0
            Caption         =   "Post Repair / Wastage"
            PicturePosition =   327683
            Size            =   "4075;661"
            Accelerator     =   82
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "L.T Balance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Index           =   6
         Left            =   4815
         TabIndex        =   72
         Top             =   8850
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Audited Lots"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Index           =   5
         Left            =   3315
         TabIndex        =   67
         Top             =   8850
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Selected Lots"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Index           =   4
         Left            =   1785
         TabIndex        =   65
         Top             =   8850
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Lots"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Index           =   3
         Left            =   405
         TabIndex        =   63
         Top             =   8835
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Master PO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   8
         Left            =   9840
         TabIndex        =   58
         Top             =   660
         Width           =   870
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Catagory"
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
         Height          =   195
         Index           =   6
         Left            =   6105
         TabIndex        =   55
         Top             =   675
         Width           =   780
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Process"
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
         Height          =   195
         Index           =   5
         Left            =   210
         TabIndex        =   54
         Top             =   660
         Width           =   660
      End
      Begin MSForms.CommandButton cmdPrintAllMakers 
         Height          =   375
         Left            =   4170
         TabIndex        =   49
         Top             =   9660
         Width           =   1770
         ForeColor       =   0
         Caption         =   "Print All Makers"
         PicturePosition =   327683
         Size            =   "3122;661"
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
         Caption         =   "Value"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Index           =   1
         Left            =   8475
         TabIndex        =   45
         Top             =   8235
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Index           =   0
         Left            =   7125
         TabIndex        =   43
         Top             =   8235
         Width           =   300
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Debit Head"
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
         Left            =   210
         TabIndex        =   40
         Top             =   1200
         Width           =   11130
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Manual Ded."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Index           =   16
         Left            =   5580
         TabIndex        =   36
         Top             =   8235
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Material Ded."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Index           =   15
         Left            =   4290
         TabIndex        =   34
         Top             =   8235
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Repair Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Index           =   14
         Left            =   2910
         TabIndex        =   32
         Top             =   8220
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "L.T Ded."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Index           =   13
         Left            =   1740
         TabIndex        =   30
         Top             =   8220
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Short Term Loan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Index           =   12
         Left            =   210
         TabIndex        =   28
         Top             =   8220
         Width           =   1410
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMakerBillingRcvingWise.frx":01F6
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
         Left            =   150
         TabIndex        =   26
         Top             =   120
         Width           =   11130
      End
      Begin MSForms.CommandButton cmdPrintMakerPostedBills 
         Height          =   375
         Left            =   5985
         TabIndex        =   25
         Top             =   9660
         Width           =   1770
         ForeColor       =   0
         Caption         =   "Maker Posted Bills"
         PicturePosition =   327683
         Size            =   "3122;661"
         Accelerator     =   76
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
         Caption         =   "Posting Date"
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
         Height          =   195
         Index           =   11
         Left            =   9480
         TabIndex        =   23
         Top             =   8205
         Width           =   1080
      End
      Begin MSForms.CommandButton cmdPrintZero 
         Height          =   375
         Left            =   2310
         TabIndex        =   12
         Top             =   9660
         Width           =   1830
         ForeColor       =   0
         Caption         =   "Print Balance Order"
         PicturePosition =   327683
         Size            =   "3228;661"
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrintUnPosted 
         Height          =   375
         Left            =   7785
         TabIndex        =   11
         Top             =   9660
         Width           =   1500
         ForeColor       =   0
         Caption         =   "Print Un-Posted"
         PicturePosition =   327683
         Size            =   "2646;661"
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPost 
         Height          =   375
         Left            =   9300
         TabIndex        =   10
         Top             =   9660
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Post Selected"
         PicturePosition =   327683
         Size            =   "2355;661"
         Accelerator     =   80
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
         Left            =   10665
         TabIndex        =   9
         Top             =   9660
         Width           =   1125
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "1984;661"
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
         Height          =   840
         Left            =   10860
         TabIndex        =   6
         Top             =   315
         Width           =   855
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "1508;1482"
         Accelerator     =   82
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
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
            Picture         =   "frmMakerBillingRcvingWise.frx":02B1
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakerBillingRcvingWise.frx":046D
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PopMenu"
   End
End
Attribute VB_Name = "frmMakerBillingRcvingWise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCourierInvoice As Boolean
Dim bShipping As Boolean
Dim bCustomInvoice As Boolean
Dim strAdvanceEntryList As String
Dim lTotalRepair As Long
Dim strMaterialDedSQL As String
Dim bEditRate As Boolean

Public Sub ShowMe(Optional p_bCustomInvoice As Boolean = True)
    bCustomInvoice = p_bCustomInvoice
    Me.Show
End Sub

Public Sub ShowForShipping()
    bShipping = True
    Me.Show
End Sub

Private Sub ChkAll_RWD_Click()

    Dim i As Integer
    Dim bValue As Boolean
    bValue = IIf(ChkAll_RWD.Value = vbChecked, True, False)
    For i = 1 To LVRepair.ListItems.count
        LVRepair.ListItems(i).Checked = bValue
    Next
    Call GetRepairTotals
    
End Sub


Private Sub cmbCust_matched()

    cmbCust.Tag = GetSingleStringValue("AccNo", "Makers", "WHERE VendID=" & cmbCust.ID)
    Call RefreshLV
    
End Sub

Private Sub cmbRange_Click()

    DtTo = GetServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    Select Case cmbRange.ListIndex
        Case 0
            DtFrom = Date
        Case 1
            DtFrom = DateAdd("d", -15, Date)
        Case 2
            DtFrom = DateAdd("d", -30, Date)
        Case 3
            DtFrom = DateAdd("d", -60, Date)
        Case 4
            DtFrom = DateAdd("d", -90, Date)
        Case 5
            DtFrom.Enabled = True
            DtTo.Enabled = True
            DtFrom.SetFocus
            Exit Sub
    End Select
    
End Sub

Private Sub cmdCats_Click()

    If LVCats.Visible Then
    
        LVCats.Visible = False
        Dim i As Integer
        If LVCats.ListItems(1).Checked Then
            txtCats.Text = "<All Catagories>"
            txtCats.Tag = "0"
        Else
            Dim strGroups As String, strGroupIDs As String
            For i = 1 To LVCats.ListItems.count
                If LVCats.ListItems(i).Checked Then
                    strGroups = strGroups & LVCats.ListItems(i).Text & ","
                    strGroupIDs = strGroupIDs & "'" & LVCats.ListItems(i).Tag & "',"
                End If
            Next
            If strGroups <> "" Then
                strGroups = Left(strGroups, Len(strGroups) - 1)
                strGroupIDs = Left(strGroupIDs, Len(strGroupIDs) - 1)
            End If
            txtCats.Text = strGroups
            txtCats.Tag = strGroupIDs
        End If
        
    Else
        LVCats.Visible = True
    End If
    
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdMarkAsAudited_Click()
    Dim i As Integer, j As Integer
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                con.Execute "UPDATE VendRcvdDetail SET Audited=1 WHERE EntryID=" & Val(.key)
                .ForeColor = vbBlue
                For j = 1 To .ListSubItems.count
                    .ListSubItems(j).ForeColor = vbBlue
                Next
                .ListSubItems(15).Tag = True
            End If
        End With
    Next
    Call LV.Refresh
    
End Sub

Private Sub cmdPost_Click()

    On Error GoTo err
    Dim bAccess As Boolean
    bAccess = GetSingleBooleanValue("PostMakerBill", "Users", " WHERE UserName='" & CurrentUserName & "'")
    If bAccess = False Then
        MsgBox "You don't have permission.", vbInformation
        Exit Sub
    End If
    
    If cmbDebitHead.MatchFound = False Then
        MsgBox "Please Select Debit Head.", vbInformation
        Exit Sub
    End If
    
    Dim i As Integer
    Dim bSelected As Boolean
    bSelected = False
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            bSelected = True
            Exit For
        End If
    Next
    If bSelected = False Then
        MsgBox "No Order Selected.", vbInformation
        Exit Sub
    End If
    Dim dValue As Double
    dValue = 0
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            dValue = dValue + Val(LV.ListItems(i).SubItems(11))
        End If
    Next
    dValue = Round(dValue)
'    'Check for Un-Auth...
'    For i = 1 To LV.ListItems.count
'        With LV.ListItems(i)
'            If .Checked Then
'                If Val(.SubItems(13)) > 0 Then
'                    MsgBox "Some Selected Orders have Un-Authorized Entries, Can't Continue.", vbInformation
'                    Exit Sub
'                End If
'            End If
'        End With
'    Next
    'Get Bill No...
    Dim lBillNo As Long, strMakerID As String, strDesc As String
    strMakerID = GetSingleStringValue("VendID1", "Makers", " WHERE VendID=" & cmbCust.ID)
    lBillNo = GetSingleLongValue("MAX(BillNo)", "MakerPostedBills", " WHERE BillNo IS NOT NULL AND VendID=" & cmbCust.ID)
    lBillNo = lBillNo + 1
   
    'Now Post Vouchers....
    Dim strVchrNo As String
    'Services / Work Order AccNo: 41-003-19001
    Dim strDebitAccNo As String, strMakerAccNo As String, dManualDedAmt As Double
    
    For i = 1 To LVManualMaterial.ListItems.count
        dManualDedAmt = dManualDedAmt + Val(LVManualMaterial.ListItems(i).SubItems(2))
    Next
    
'    If strCompany = "Sunlike" Then
'        strDebitAccNo = "41-001-16005"
'    Else
'        strDebitAccNo = "41-003-19001"
'    End If

    If cmbDebitHead.MatchFound Then
        strDebitAccNo = cmbDebitHead.ID
        strMakerAccNo = GetSingleStringValue("AccNo", "Makers", " WHERE VendID=" & cmbCust.ID)
         
        Dim objVchr As New clsVouchers
        objVchr.VoucherDate = DTPosting 'GetServerDate
        objVchr.VoucherType = "JV"
        
        Dim dAmtToPost As Double
        'dAmtToPost = Val(txts1(4)) - Val(txtMakerAdvancesOther) - dManualDedAmt
        'dAmtToPost = Val(txts(4)) - Val(txtDeductShort) - dManualDedAmt - Val(txtFixLoan) - Val(txtMaterialDed)
        Dim lRepairAmt As Long, lRepairRelax As Long, lRepairAmtToDeduct As Long
        
        If chkRejectionRelaxation.Value = vbChecked Then
            lRepairAmt = Val(txtRepairAmt)
            lRepairRelax = (dValue / 100) * 2
            lRepairAmtToDeduct = lRepairAmt - lRepairRelax
            
            If lRepairAmtToDeduct < 0 Then
                lRepairAmtToDeduct = 0
            End If
        Else
            lRepairAmtToDeduct = Val(txtRepairAmt)
        End If
        
        dAmtToPost = dValue - dManualDedAmt - Val(txtMaterialDed) - lRepairAmtToDeduct
        
        If dAmtToPost <= 0 Then
            MsgBox "Invalid Amount to Post.", vbInformation
            Exit Sub
        End If
        
        strDesc = "Computer Bill No. " & strMakerID & " / " & lBillNo & " - " & year(DTPosting)
        objVchr.AddVoucherHead strDebitAccNo, dAmtToPost, strDesc
        
        objVchr.AddVoucherHead strMakerAccNo, -(dAmtToPost), strDesc
        Call StartTrans(con)
        strVchrNo = objVchr.PostVoucher(, True)
    Else
        Call StartTrans(con)
    End If
    
'    If Val(txtRepairAmt) <> 0 Then  'If Repair....
'        objVchr.AddVoucherHead "24-005-19001", -Val(txts(3)), strDesc
'    End If
     
    
    
        
'    Dim dRepairAmt As Double, dWastageAmt As Double, strRepairVchrNo As String
    'Call PostRepairVchr(strRepairVchrNo, dRepairAmt, dWastageAmt)
    Dim dRepairAmt As Double, dWastageAmt As Double
'    For i = 1 To LVRepair.ListItems.count
'        With LVRepair.ListItems(i)
'            dRepairAmt = dRepairAmt + Val(.SubItems(6))
'            dWastageAmt = dWastageAmt + Val(.SubItems(7))
'        End With
'    Next
    dRepairAmt = 0
    dWastageAmt = Val(txtRepairAmt)
    
    con.Execute "INSERT INTO MakerPostedBills(VendID,DT,VchrNo,UserName,MachineName,BillNo,TempAdvance,BalanceOrderValue," & _
    "UnAuthValue,RepairVchrNo,RepairAmt,WastageAmt,ManualDedAmt,STDeduction,LTDeduction,MaterialDedAmt,DTFrom,DTTo,WastageAmtRelaxation,TotalWastageAmt) VALUES(" & cmbCust.ID & ",'" & DTPosting & "','" & strVchrNo & "','" & CurrentUserName & "','" & _
     strComputerName & "'," & lBillNo & "," & Val(txtMakerAdvancesOther) & ",0" & ",0,'" & _
     strRepairVchrNo & "'," & dRepairAmt & "," & lRepairAmtToDeduct & "," & dManualDedAmt & "," & Val(txtDeductShort) & "," & _
     Val(txtFixLoan) & "," & Val(txtMaterialDed) & ",'" & DtFrom.Value & "','" & DtTo.Value & "'," & lRepairRelax & "," & lRepairAmt & ")"
    
    Dim lEntryID As Long
    lEntryID = GetSingleLongValue("MAX(EntryID)", "MakerPostedBills")
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                con.Execute "INSERT INTO MakerPostedBillsDetail(RefID,Iss_RefID,TempAdvance," & _
                 "RepairAmt,RcvdValue,RcvdQty,UnAuthValue,BalanceValue,RcvdPercentage,IssDetail_RefID) VALUES(" & lEntryID & "," & Val(.Tag) & "," & _
                 Val(.SubItems(9)) & "," & Val(.SubItems(10)) & "," & Val(.SubItems(11)) & "," & Val(.SubItems(12)) & _
                 "," & Val(.SubItems(13)) & "," & Val(.SubItems(14)) & "," & Val(.SubItems(15)) & "," & Val(.ListSubItems(1).Tag) & ")"
             
                'con.Execute "UPDATE VendIssued SET VchrNo='" & strVchrNo & "' WHERE EntryID=" & .Tag
                
                con.Execute "INSERT INTO MakerPostedBillsDetail_Receivings(MPB_RefID,VRD_RefID,Qty,Rate,Main_VI_EntryID) VALUES(" & lEntryID & "," & Val(.key) & "," & Val(.SubItems(12)) & "," & Val(.SubItems(7)) & "," & Val(.ListSubItems(16).Tag) & ")"
            End If
        End With
    Next
        
'    If strCompany = "PAKSMITH GROUP" Or strCompany = "QEL" Then
'        'Now Update TemporaryLoanGroups
'        If strAdvanceEntryList <> "" Then
'            con.Execute "UPDATE MakerTemporaryLoansGroup SET Active=0 WHERE EntryID IN(SELECT GroupID FROM MakerTemporaryLoans WHERE Issuance_RefID IN(" & strAdvanceEntryList & "))"
'        End If
'    End If
        
    'Now Save Manual Material Ded. Entries...
    For i = 1 To LVManualMaterial.ListItems.count
        With LVManualMaterial.ListItems(i)
            con.Execute "INSERT INTO MakerPostedBills_ManualDeductions(Description,Amount,MPB_RefID) VALUES('" & _
             .SubItems(1) & "'," & Val(.SubItems(2)) & "," & lEntryID & ")"
        End With
    Next
    
    'Now Save Material Ded. Entries...
    For i = 1 To LVMaterial.ListItems.count
        With LVMaterial.ListItems(i)
            If .Checked Then
                con.Execute "INSERT INTO MakerPostedBills_MaterialDeductions(MPB_RefID,RMI_RefID,RMID_RefID) VALUES(" & _
                 lEntryID & "," & Val(.Tag) & "," & Val(.key) & ")"
            End If
        End With
    Next
    
    'Now Repair
    For i = 1 To LVRepair.ListItems.count
        With LVRepair.ListItems(i)
            If .Checked Then
                If Val(.ListSubItems(8).Tag) = 0 Then
                    con.Execute "UPDATE MakerRepair SET VchrNo='" & strVchrNo & "' WHERE EntryID=" & Val(.key)
                Else
                    con.Execute "INSERT INTO MakerPostedBills_Wastage(MPB_RefID,MBD_RefID) VALUES(" & _
                     lEntryID & "," & Val(.key) & ")"
                End If
            End If
        End With
    Next
    
    If Val(txtFixLoan) > 0 Then
        con.Execute "INSERT INTO MakerAmtCleared(AccNo,DT,AmtClrd,VchrNo) VALUES('" & cmbCust.Tag & "','" & _
         DTPosting & "'," & Val(txtFixLoan) & ",'" & strVchrNo & "')"
    End If
        
    If Val(txtDeductShort) > 0 Then 'Clear All Short Term... Deducting All....
        'con.Execute "INSERT INTO MakerPostedBills_STDeductions(MAS_RefID,MPB_RefID,Amount) SELECT EntryID," & lEntryID & ",Amount-AmountCleared FROM VMakerAdvancesShort WHERE AccNo='" & cmbCust.Tag & "' AND Amount>AmountCleared"
        'con.Execute "UPDATE MakerAdvancesShort SET AmountCleared=Amount WHERE AccNo='" & cmbCust.Tag & "'"
        con.Execute "INSERT INTO MakerPostedBills_STDeductions(MPB_RefID,Amount,AccNo) VALUES(" & lEntryID & "," & Val(txtDeductShort) & ",'" & cmbCust.Tag & "')"
    End If
    
    con.CommitTrans

    LVManualMaterial.ListItems.Clear
    
    MsgBox "Posted Successfuly !!!", vbInformation
    Call PrintMakerPostedBill(strVchrNo, False, lEntryID)
    
    Call RefreshLV
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub PostRepairVchr(ByRef p_strVchrNo As String, ByRef p_dRepairAmt As Double, ByRef p_dWastageAmt As Double)

    Dim i As Integer, bYouCanPost As Boolean
    bYouCanPost = False
'    For i = 1 To LVRepair.ListItems.count
'        If LVRepair.ListItems(i).Checked Then
'            bYouCanPost = True
'            Exit For
'        End If
'    Next
    If LVRepair.ListItems.count > 0 Then
        bYouCanPost = True
    End If
    
    If bYouCanPost = False Then Exit Sub
    
    Dim dRepairAmt As Double, dWastageAmt As Double
    For i = 1 To LVRepair.ListItems.count
        With LVRepair.ListItems(i)
            If .Checked Then
                dRepairAmt = dRepairAmt + Val(.SubItems(6))
                dWastageAmt = dWastageAmt + Val(.SubItems(7))
            End If
        End With
    Next
    'Now Post...
    Dim dAmt As Double, strRepairAccNo As String, strVchrNo As String, strWastageAccNo As String, strMakerAccNo As String
    Dim strDescription As String
    strRepairAccNo = "41-004-13001"
    strWastageAccNo = "41-003-19001"
    strMakerAccNo = GetSingleStringValue("AccNo", "Makers", " WHERE VendID=" & cmbCust.ID)
    strDescription = "Repair / Wastage Charges to " & cmbCust.Text
    
    dAmt = dWastageAmt + dRepairAmt
    Dim objVchr As New clsVouchers
    objVchr.VoucherType = "JV"
    objVchr.VoucherDate = DTPosting
    objVchr.AddVoucherHead strMakerAccNo, dAmt, strDescription
    If dRepairAmt > 0 Then
        objVchr.AddVoucherHead strRepairAccNo, -dRepairAmt, strDescription
    End If
    If dWastageAmt > 0 Then
        objVchr.AddVoucherHead strWastageAccNo, -dWastageAmt, strDescription
    End If
    
    strVchrNo = objVchr.PostVoucher(False, True)
    
    For i = 1 To LVRepair.ListItems.count
        With LVRepair.ListItems(i)
            If .Checked Then
                con.Execute "UPDATE MakerRepair SET VchrNo='" & strVchrNo & "' WHERE EntryID=" & Val(.key)
            End If
        End With
    Next
         
    p_strVchrNo = strVchrNo
    p_dRepairAmt = dRepairAmt
    p_dWastageAmt = dWastageAmt
    
    'Call AddRepairDetail
    
End Sub

Private Sub cmdPrintAllMakers_Click()
    Call PrintUnPostedBills(0, False)
End Sub

Private Sub cmdPrintMakerPostedBills_Click()

    If cmbCust.MatchFound = False Then
        MsgBox "Please Select Maker.", vbInformation
        Exit Sub
    End If

    Dim myDTFrom As Date, myDTTo As Date
    Dim f As New frmGetDateDouble
    If f.getDate(myDTFrom, myDTTo) = False Then Exit Sub

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\MakerPostedBillsListIssWise.rpt")

    Dim frm As New frmPrevRpt
    frm.ShowReport "{MakerPostedBills.VendID}=" & cmbCust.ID & " AND {MakerPostedBills.DT}=#" & myDTFrom & "# To #" & myDTTo & "#", rpt
        
End Sub

Private Sub cmdPrintMaterialDeductions_Click()
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\MaterialDeductions.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, strMaterialDedSQL
End Sub

Private Sub cmdPrintUnPosted_Click()
    Call PrintUnPostedBills(cmbCust.ID, False)
End Sub

Private Sub PrintUnPostedBills(lMakerID As Long, Optional bShowZero As Boolean = False)
    
    On Error GoTo err
    Dim rptApp_Temp As New CRAXDDRT.Application
    Dim rpt As CRAXDDRT.Report
    
    Dim strEntryIDs As String, i As Integer
    strEntryIDs = ""
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                strEntryIDs = strEntryIDs & Val(.key) & ","
            End If
        End With
    Next

    If lMakerID <> 0 Then
        If strEntryIDs = "" Then
            MsgBox "Please Select Entries.", vbInformation
            Exit Sub
        End If
        strEntryIDs = Left(strEntryIDs, Len(strEntryIDs) - 1)
    End If

    Dim strSelection As String
    If lMakerID <> 0 Then
        strSelection = "{MakerPostedBillsDetail.RcvDetailEntryID} IN[" & strEntryIDs & "]"
    End If
    
    If lMakerID = 0 Then
        Set rpt = rptApp_Temp.OpenReport(RptPath & "\AllMakersBilling.rpt")
    Else
        Set rpt = rptApp_Temp.OpenReport(RptPath & "\MakerUnPostedBills.rpt")
    End If
    
    rpt.ParameterFields.GetItemByName("@MakerID").SetCurrentValue CInt(cmbCust.ID)
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DtFrom.Value
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DtTo.Value
    
    rpt.ParameterFields.GetItemByName("@CatID").SetCurrentValue cmbCats.ID
   
    rpt.ParameterFields.GetItemByName("@ProcessID").SetCurrentValue Val(cmbProcess.ID)
    rpt.ParameterFields.GetItemByName("@MasterPONo").SetCurrentValue txtMasterPONo.Text
    'rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DtTo.Value
    
    rpt.FormulaFields.GetItemByName("ShowMaker").Text = True
    rpt.FormulaFields.GetItemByName("UserName").Text = "'" & CurrentUserName & "'"
    rpt.FormulaFields.GetItemByName("ComputerName").Text = "'" & strComputerName & "'"
    
    rpt.FormulaFields.GetItemByName("ShortTermLoan").Text = Val(txtDeductShort)
    rpt.FormulaFields.GetItemByName("LongTermDeduction").Text = Val(txtFixLoan)
    
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    
'    If bShowZero Then
'        strSelection = "{@RcvdQty}=0"
'    Else
'        strSelection = "{@RcvdQty}>0"
'    End If
    
    rpt.FormulaFields.GetItemByName("STDeduction").Text = Val(txtDeductShort)
    rpt.FormulaFields.GetItemByName("LTDeduction").Text = Val(txtFixLoan)
    rpt.FormulaFields.GetItemByName("OtherRejectionAmt").Text = Val(txtRepairAmt)
    rpt.FormulaFields.GetItemByName("MaterialDeduction").Text = Val(txtMaterialDed)
    rpt.FormulaFields.GetItemByName("ManualDeduction").Text = Val(txtManualDed)
    rpt.FormulaFields.GetItemByName("LT_Balance").Text = Val(txtFixLoan.Tag)
    
    If lMakerID <> 0 Then
        Dim rptSub As CRAXDDRT.Report
        Set rptSub = rpt.OpenSubreport("RepairWastageDetail")
        
        Dim strEntryIDs_Repair As String, strSelection_Repair As String
        strEntryIDs_Repair = ""
        For i = 1 To LVRepair.ListItems.count
            With LVRepair.ListItems(i)
                If .Checked Then
                    strEntryIDs_Repair = strEntryIDs_Repair & Val(.key) & ","
                End If
            End With
        Next
        
        If strEntryIDs_Repair <> "" Then
            strEntryIDs_Repair = Left(strEntryIDs_Repair, Len(strEntryIDs_Repair) - 1)
            strSelection_Repair = "{VMakerRepair.EntryID} IN[" & strEntryIDs_Repair & "]"
        Else
            strSelection_Repair = "0=1"
        End If
        'rptSUB.RecordSelectionFormula = "{VMakerRepair.VendID}=" & cmbCust.ID & " AND {VMakerRepair.DT}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "# " & strCondition
        rptSub.RecordSelectionFormula = strSelection_Repair
        
        Dim dTotalRepair As Double
        For i = 1 To LVRepair.ListItems.count
            If LVRepair.ListItems(i).Checked Then
                dTotalRepair = dTotalRepair + Val(LVRepair.ListItems(i).SubItems(7))
            End If
        Next
        rpt.FormulaFields.GetItemByName("RepairTotal").Text = dTotalRepair
    End If
    
    f.ShowReport strSelection, rpt, , False
    
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub cmdPrintZero_Click()
    Call PrintUnPostedBills(cmbCust.ID, True)
End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub

Private Sub cmdSelectAll_Click()
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = True
    Next
    Call LV_ItemCheck(LV.ListItems(1))
End Sub

Private Sub cmdSelectAudited_Click()

    Dim i As Long, lIndex As Long
    For i = 1 To LV.ListItems.count
        If CBool(LV.ListItems(i).ListSubItems(15).Tag) Then
            LV.ListItems(i).Checked = True
            lIndex = i
        Else
            LV.ListItems(i).Checked = False
        End If
    Next
    If lIndex > 0 Then Call LV_ItemCheck(LV.ListItems(lIndex))
    
End Sub

Private Sub cmdUnmarkAudited_Click()

    Dim i As Integer, j As Integer
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                con.Execute "UPDATE VendRcvdDetail SET Audited=0 WHERE EntryID=" & Val(.key)
                .ForeColor = vbBlack
                For j = 1 To .ListSubItems.count
                    .ListSubItems(j).ForeColor = vbBlack
                Next
                .ListSubItems(15).Tag = False
            End If
        End With
    Next
    Call LV.Refresh
End Sub

Private Sub cmdUnSelectAll_Click()
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = False
    Next
    Call LV_ItemCheck(LV.ListItems(1))
End Sub

Private Sub DtFrom_Change()
    'Call setDateRange(DtFrom)
End Sub

Private Sub DtTo_Change()
    'Call setDateRange(DtTo)
End Sub


Private Sub Form_Load()
    
    mnuPop.Visible = False
    cmbCust.ListHeight = 3800
    
    bEditRate = GetSingleBooleanValue("MakerBill_EditRate", "Users", "WHERE UserName='" & CurrentUserName & "'")
    
    'cmbCust.AddItem "<All Makers>", "0"
    Call cmbCust.AddVals(con, "'{ ' + VendID1 + ' } ' + VenderName ", "Makers", "VendID", " WHERE VendID<>" & iFactoryMakerID)
    cmbCust.ID = "0"
    DtFrom = Date
    DtTo = Date
    cmbRange.ListIndex = 1
    If strCompany = "Weldon Instruments" Or strCompany = "Sunlike" Or strCompany = "QEL" Or (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
        SSTab1.TabVisible(1) = True
    Else
        SSTab1.TabVisible(1) = False
    End If
    
    If strCompany <> "Dr-Frgz" Then
        cmdPrintZero.Visible = False
    End If
    
'    With cmbItemCatagory
'        .ListHeight = 4000
'        .AddItem "<All Catagories>", "0"
'        .AddVals con, "Description", "ItemCatagories", "CatID"
'        .ID = "0"
'    End With
    
    SSTab1.Tab = 0
    
    DTPosting.Value = Date
    
    With cmbDebitHead
        .ListHeight = 4000
        .AddVals con, "AccTitle + '{ ' + AccNo + ' }'", "VActiveAccounts", "AccNo"
    End With
     
    cmbCats.ListHeight = 4000
    cmbCats.EnableFunctionality = False
    cmbCats.Functionality = 1
    cmbCats.AddItem "<All Catagories>", "0"
    cmbCats.AddVals con, "Description", "ItemCatagories", "CatID"
    cmbCats.ID = "0"
    
    cmbProcess.ListHeight = 4000
    cmbProcess.AddItem "<All Processes>", "0"
    cmbProcess.AddVals con, "Description", "Processes", "ProcessID", " ORDER BY SNo"
    cmbProcess.ID = "0"
'    Dim rs As New ADODB.Recordset, ITM As ListItem
'    With rs
'        .Open "SELECT * FROM ItemCatagories", con, adOpenForwardOnly, adLockReadOnly
'        LVCats.ListItems.Clear
'        LVCats.ListItems.add(, , "<All Catagories>").Checked = True
'        txtCats.Text = "<All Catagories>"
'        txtCats.Tag = "0"
'        Do Until .EOF
'            Set ITM = LVCats.ListItems.add(, !CatID & "'", !Description & "")
'            ITM.Tag = !CatID & ""
'            .MoveNEXT
'        Loop
'        .Close
'    End With
'    Set rs = Nothing
    
End Sub

Private Sub RefreshLV()
        
    On Error GoTo err
    If cmbCust.MatchFound = False Then
        MsgBox "Please Select Maker.", vbInformation
        Exit Sub
    End If
    Dim lTotal_Audited As Long
    
    strAdvanceEntryList = ""
    txtMakerAdvancesOther = ""
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "MakerBilling_RcvingWise_SP"
        .Parameters("@MakerID").Value = cmbCust.ID
        .Parameters("@DTFrom").Value = DtFrom.Value
        .Parameters("@DTTo").Value = DtTo.Value
        .Parameters("@CatID").Value = cmbCats.ID 'txtCats.Tag 'Val(cmbItemCatagory.ID)
        .Parameters("@ProcessID").Value = Val(cmbProcess.ID)
        .Parameters("@MasterPONo").Value = Trim(txtMasterPONo)
        
        Set rs = .Execute
    End With
    
    Dim Cond As String
    Dim DTRange As String
    Dim ITM As ListItem
      
    Screen.MousePointer = vbHourglass
    
'    If cmbCust.ID <> "0" Then
'        Cond = " Where VendID=" & cmbCust.ID & " "
'    End If
    Dim i As Integer
    Dim lForeColor As Long
    With rs
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, Val(!RcvDetailEntryID & "") & "'", .AbsolutePosition)
            ITM.Tag = Val(![EntryID] & "")
            ITM.ListSubItems.add(, , !RcvRecieptID & "").ToolTipText = !ItemName & ""
            ITM.ListSubItems.add(, , Format(!RcvDT, "dd-MMM-yy")).ToolTipText = !ItemName & ""
            ITM.ListSubItems.add(, , !OrderNo & "").ToolTipText = !ItemName & ""
            ITM.ListSubItems.add(, , !ITemID & "").ToolTipText = !ItemName & ""
            ITM.ListSubItems.add(, , !ItemName & "").ToolTipText = !ItemName & ""
            ITM.ListSubItems.add(, , Val(!IssQty & "")).ToolTipText = !ItemName & ""
            ITM.ListSubItems.add(, , Val(!Rate & "")).ToolTipText = !ItemName & ""
            ITM.ListSubItems.add(, , Val(!Rate & "") * Val(!IssQty & "")).ToolTipText = !ItemName & ""
            ITM.ListSubItems.add(, , Val(!TempAdvance & "")).ToolTipText = !ItemName & ""
            ITM.ListSubItems.add(, , Val(!RepairAmt & "")).ToolTipText = !ItemName & ""
            'ITM.ListSubItems.add(, , Round(Val(!RcvdValue & ""))).Tag = Val(!RcvdQty & "") 'Round(Val(!RcvdQty & "") * Val(!Price & ""), 2)
            ITM.ListSubItems.add(, , Round(Val(!BillingProcAmount & ""), 4)).Tag = Val(!RcvdQty & "")  'Round(Val(!RcvdQty & "") * Val(!Price & ""), 2)
            ITM.ListSubItems.add(, , Val(!BillingProcRcvdQty & "")).Tag = Val(!RcvdQtyUnAuth & "")
            ITM.ListSubItems.add(, , Round(Val(!RcvdValueUnAuth & ""))).Tag = Val(!RcvdQtyUnAuth & "") 'Round(Val(!RcvdQtyUnAuth & "") * Val(!Price & ""), 2)
            ITM.ListSubItems.add , , Round(Val(!RcvdValue & "") - Val(!RcvdValueUnAuth & "") - Val(!RepairAmt & "")) 'Round(Val(!value & "") - Val(ITM.SubItems(8)) - Val(ITM.SubItems(9)), 2)
            ITM.ListSubItems.add , , Round(Val(!PercentRcvd & "")) & "%"
            ITM.ListSubItems.add , , !LotNo & ""
            ITM.ListSubItems.add , , !Description & ""
            ITM.ListSubItems.add , , !CustCode & ""
            ITM.ListSubItems(1).Tag = Val(!Issue_RefID & "")
            
            ITM.ListSubItems(15).Tag = !Audited
            
            ITM.ListSubItems(16).Tag = Val(!MainIssuanceEntryID & "")
             
            If !Audited Then
                lForeColor = vbBlue
                lTotal_Audited = lTotal_Audited + 1
            Else
                lForeColor = vbBlack
            End If
            
            ITM.ForeColor = lForeColor
            For j = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(j).ForeColor = lForeColor
            Next
            .MoveNEXT
        Loop
        .Close
            
            
        .Open "SELECT TotalAmountTaken,TotalAmountCleared,ActualDeductionAmount FROM VMakerAdvancesDeduction WHERE AccNo='" & cmbCust.Tag & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            txtFixLoan.Text = "0" 'Val(!ActualDeductionAmount & "")
            txtFixLoan.Tag = "0" 'Val(!TotalAmountTaken & "") - Val(!TotalAmountCleared & "") - Val(!ActualDeductionAmount & "")
        Else
            txtFixLoan.Text = Val(!ActualDeductionAmount & "")
            txtFixLoan.Tag = Val(!TotalAmountTaken & "") - Val(!TotalAmountCleared & "") '- Val(!ActualDeductionAmount & "")
            txtLTBalance = Val(!TotalAmountTaken & "") - Val(!TotalAmountCleared & "")
        End If
        .Close
        
        'Get Short Term Deduction ....
        '.Open "SELECT SUM(Amount-AmountCleared) AS TotalDeductionAmt FROM VMakerAdvancesShort WHERE AccNo='" & cmbCust.Tag & "'", con, adOpenForwardOnly, adLockReadOnly
        .Open "SELECT * FROM VMakerAdvancesShortDeduction WHERE AccNo='" & cmbCust.Tag & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            txtDeductShort.Text = Val(!Amount & "") - Val(!AmountCleared & "")
            txtDeductShort.Tag = Val(!Amount & "") - Val(!AmountCleared & "")
        Else
            txtDeductShort.Text = 0
            txtDeductShort.Tag = "0"
        End If
        .Close
        
    End With
    
    txtTotalLots = LV.ListItems.count
    txtTotalAudited = lTotal_Audited
    txtTotalSelected = 0
    
    Set rs = Nothing
            
'    If strCompany = "Weldon Instruments" Or strCompany = "Sunlike" Or strCompany = "QEL" Or (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Banzai" Then
'        Call AddRepairDetail
'    Else
'        txtRepairAmt = "0"
'    End If
    Call AddRepairDetail
    
'    If strCompany = "QEL" Or (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Banzai" Or strCompany = "Towne" Then
        Call AddMaterialIssuanceDetail
'    Else
'        txtMaterialDed = "0"
'    End If
        
    
'    For i = 1 To LV.ListItems.count
'        LV.ListItems(i).Checked = True
'    Next
    Call ShowHiddenLots
    Call ShowHiddenRM
    LVManualMaterial.ListItems.Clear
    
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Private Sub AddRepairDetail()

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem, dRate As Double
    Dim strCondition As String
    
    If cmbCats.ID <> "0" Then
        strCondition = " AND CatID IN(" & cmbCats.ID & ")"
    End If
    'Dim lTotal As Long
    With rs
        .Open "SELECT * FROM VMakerRepair WHERE VendID=" & cmbCust.ID & " AND VchrNo IS NULL AND (ShowInBilling=0 OR (ShowInBilling=1 AND DeductFromBilling=1)) AND DT BETWEEN '" & DtFrom.Value & "' AND '" & DtTo.Value & "' " & strCondition, con, adOpenForwardOnly, adLockReadOnly
        LVRepair.ListItems.Clear
        Do Until .EOF
            Set ITM = LVRepair.ListItems.add(, !EntryID & "'", LVRepair.ListItems.count + 1)
            ITM.ListSubItems.add , , !RecieptID & ""
            ITM.ListSubItems.add , , Format(!DT, "dd-MMM-yyyy")
            ITM.ListSubItems.add , , !ItemCode & ""
            ITM.ListSubItems.add , , Val(!Qty & "")
'            If IsNull(!Wastage_Originator) Then
'                dRate = Val(!VendIssdDetailRate & "")
'            ElseIf !Wastage_Originator Then
'                dRate = Val(!Rate & "")
'            Else
'                dRate = Val(!VendIssdDetailRate & "")
'            End If
            dRate = Val(!Rate & "")
            ITM.ListSubItems.add , , dRate
            ITM.ListSubItems.add , , "0"    'Repair Amt....
            ITM.ListSubItems.add , , Val(!Qty & "") * dRate
            ITM.ListSubItems.add , , !LotNo & ""
            ITM.ListSubItems.add , , !Description & ""
            'lTotal = lTotal + (Val(!Qty & "") * dRate) + Val(!RepairAmt & "")
            .MoveNEXT
        Loop
        .Close
        
        .Open "SELECT * FROM VMakerBilling_WastageToDeduct WHERE VendID=" & cmbCust.ID & " AND DT BETWEEN '" & DtFrom.Value & "' AND '" & DtTo.Value & "' " & strCondition, con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            Set ITM = LVRepair.ListItems.add(, !EntryID & "'", LVRepair.ListItems.count + 1)
            ITM.ListSubItems.add , , !RecieptID & ""
            ITM.ListSubItems.add , , Format(!DT, "dd-MMM-yyyy")
            ITM.ListSubItems.add , , !ItemCode & ""
            ITM.ListSubItems.add , , Val(!Wastage & "")
            
            ITM.ListSubItems.add , , Val(!IssRate & "")
            ITM.ListSubItems.add , , "0"    'Repair Amt....
            ITM.ListSubItems.add , , Val(!Wastage & "") * Val(!IssRate & "")
            ITM.ListSubItems.add(, , !LotNo & "").Tag = "1"
            ITM.ForeColor = vbBlue
            For i = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(i).ForeColor = vbBlue
            Next
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    'txtRepairAmt = lTotal
    
End Sub

Private Sub GetRepairTotals()

    Dim i As Integer, lTotal As Long
    For i = 1 To LVRepair.ListItems.count
        With LVRepair.ListItems(i)
            If .Checked Then
                lTotal = lTotal + Val(.SubItems(7))
            End If
        End With
    Next
    txtRepairAmt = lTotal
    
End Sub

Private Sub AddMaterialIssuanceDetail()

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim lTotal As Long
    With rs
        Dim strCondition As String
        strCondition = " WHERE VendID=" & cmbCust.ID & " AND MasterEntryID NOT IN(SELECT RMI_RefID FROM MakerPostedBills_MaterialDeductions) AND EntryID NOT IN(SELECT RMID_RefID FROM MakerPostedBills_MaterialDeductions)"
        strCondition = strCondition & " AND EntryID NOT IN(SELECT RMID_RefID FROM RawMaterialIssuanceDetail_Hidden_From_Maker_Billing)"
        If Trim(txtMasterPONo) <> "" Then
            strCondition = strCondition & " AND MasterPONo='" & Trim(txtMasterPONo) & "'"
        End If
        strMaterialDedSQL = "SELECT * FROM VRawMaterialIssuance " & strCondition
        .Open strMaterialDedSQL, con, adOpenForwardOnly, adLockReadOnly
        LVMaterial.ListItems.Clear
        Do Until .EOF
            Set ITM = LVMaterial.ListItems.add(, !EntryID & "'", LVMaterial.ListItems.count + 1)
            ITM.Tag = Val(!MasterEntryID & "")
            ITM.ListSubItems.add , , !IssNo & ""
            ITM.ListSubItems.add , , Format(!DT, "dd-MMM-yyyy")
            ITM.ListSubItems.add , , "(" & !RMID1 & ") " & !RMName
            ITM.ListSubItems.add , , !ITemID & ""
            ITM.ListSubItems.add , , Val(!Price & "")
            ITM.ListSubItems.add , , Val(!Qty & "")
            ITM.ListSubItems.add , , Round(Val(!Qty & "") * Val(!Price & ""))
            lTotal = lTotal + (Round(Val(!Qty & "") * Val(!Price & "")))
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    txtMaterialDed = lTotal
    
End Sub

Private Sub getMaterialDedTotal()
    Dim i As Integer, lTotal As Long
    For i = 1 To LVMaterial.ListItems.count
        With LVMaterial.ListItems(i)
            If .Checked Then
                lTotal = lTotal + Val(.SubItems(7))
            End If
        End With
    Next
    txtMaterialDed = lTotal
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    fTitle.Left = (Me.ScaleWidth - fTitle.Width) / 2
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    'FAcc.Top = (Me.ScaleHeight - FAcc.Height) / 2
End Sub


Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    Dim myLV As ListView
    Set myLV = LV
    
    If myLV.SortKey + 1 = ColumnHeader.Index Then
        If myLV.SortOrder = lvwAscending Then
            myLV.SortOrder = lvwDescending
        Else
            myLV.SortOrder = lvwAscending
        End If
    Else
        myLV.SortKey = ColumnHeader.Index - 1
        myLV.SortOrder = lvwAscending
    End If
    If myLV.Sorted = False Then myLV.Sorted = True

End Sub

Private Sub LV_DblClick()
    'Edit Rate
    If LV.ListItems.count = 0 Then Exit Sub
    
    If bEditRate = False Then Exit Sub
    
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(8).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(8).Width
        .Text = Val(LV.SelectedItem.SubItems(7))
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
End Sub

Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    Dim i As Integer
    Dim lTotalQty As Long, lTotalSelectedLots As Long
    Dim lTotalValue As Double
    strAdvanceEntryList = ""
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                
                lTotalQty = lTotalQty + Val(.SubItems(12)) 'RcvdValue
                lTotalValue = lTotalValue + Val(.ListSubItems(11))  'RcvdQty
                lTotalSelectedLots = lTotalSelectedLots + 1
                
            End If
        End With
    Next
    
    txtQty = lTotalQty
    txtValue = Round(lTotalValue)
    
    txtTotalSelected = lTotalSelectedLots
    
'    If strCompany = "PAKSMITH GROUP" Or strCompany = "QEL" Then
'        If strAdvanceEntryList <> "" Then
'            strAdvanceEntryList = Left(strAdvanceEntryList, Len(strAdvanceEntryList) - 1)
'            txtMakerAdvancesOther = GetSingleLongValue("SUM(Amount)", "MakerTemporaryLoans", " WHERE GroupID IN(SELECT GroupID FROM MakerTemporaryLoans WHERE Issuance_RefID IN(" & strAdvanceEntryList & ")) AND GroupID NOT IN(SELECT EntryID FROM MakerTemporaryLoansGroup WHERE Active=0)")
'        Else
'            txtMakerAdvancesOther = ""
'        End If
'    End If
    
    
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)

    If LV.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure to Delete ? Entry Will be hidden from Billing.", vbQuestion + vbYesNo) = vbYes Then
            con.Execute "INSERT INTO MakerBilling_Hidden(Iss_RefID) VALUES(" & Val(LV.SelectedItem.Tag) & ")"
            Call RefreshLV
        End If
    End If
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If LV.ListItems.count = 0 Then Exit Sub
    If Button <> 2 Then Exit Sub
    Dim strRet As String
    strRet = ShowPopUpMenu1("Hide This Lot from Billing", "Print PTC")
    
    If strRet = "Hide This Lot from Billing" Then
        con.Execute "INSERT INTO MakerBilling_Receivings_Hidden(VRD_RefID,UserName,MachineName) VALUES(" & Val(LV.SelectedItem.key) & ",'" & CurrentUserName & "','" & strComputerName & "')"
        LV.ListItems.Remove LV.SelectedItem.Index
    ElseIf strRet = "Print PTC" Then
        Call PrintPTC(LV.SelectedItem.SubItems(16))
   End If
   
End Sub

Private Sub ShowMenusForShipping()

    On Error Resume Next
    Dim c As Control
    For Each c In Controls
        If TypeOf c Is Menu Then
            If c.Name <> "mnuPrintPList" Then
                c.Visible = False
            End If
        End If
    Next
    
    On Error GoTo 0
    
    If LV.ListItems.count > 0 Then
        mnuPrintPList.Visible = True
        mnuCustomPList.Visible = True
        mnuCustomPListOther.Visible = True
        mnuCommercialPList.Visible = True
        mnuCommercialPListOther.Visible = True
        mnuDashPackingLabels.Visible = True
        mnuPrintPackingLabels.Visible = True
    End If
'    mnuPrintDecDocs.Visible = False
    Me.PopUpMenu mnuPop
    
End Sub

Private Sub ShowMenus(vis As Boolean)
        
    
    Me.PopUpMenu mnuPop
    
End Sub

Private Sub ShowMenusforCourier()
         
    On Error Resume Next
    Dim c As Control
    For Each c In Controls
        If TypeOf c Is Menu Then
            c.Visible = False
        End If
    Next
    
    On Error GoTo 0
    mnuAdd.Visible = True
    If LV.ListItems.count > 0 Then
        dd.Visible = True
        mnuEdit.Visible = True
        mnuComDocs.Visible = True
        mnuPrintComInv.Visible = True
    End If
    mnuPrintDecDocs.Visible = False
    Me.PopUpMenu mnuPop
    
End Sub


Private Sub mnuAdd_Click()
    'Load frmNewCustomInvoice
    'frmNewCustomInvoice.Show
    'ShowForm frmNewCustomInvoice, MainForm, False, "New Custom Invoice"
    'frmNewCustomInvoice.AddNewProforma
    Load frmNewCustomInvoice
    frmNewCustomInvoice.AddNewInvoice True, bCourierInvoice
End Sub

Private Sub mnuAddShippInfo_Click()

    Load frmEditShippInfo
    'ShowForm f, MainForm, , "Shipp. Info"
    Call frmEditShippInfo.EditEntry(LV.SelectedItem.SubItems(3))
    
End Sub

Private Sub mnuBillOfExchange_Click()
    Call PrintBillOfExchange
End Sub

Private Sub PrintBillOfExchange(Optional Firstrpt As Boolean = True)

    On Error GoTo err
    Me.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptBillOfExchange.rpt")
    rpt.FormulaFields(3).Text = Firstrpt
    Dim frm As New frmPrevRpt
    frm.ShowReport "{CustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    Me.MousePointer = vbNormal
    
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description

End Sub

    
Private Sub mnuBillOfExchange2_Click()
    Call PrintBillOfExchange(False)
End Sub

Private Sub mnuCommercialPList_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3), False, True, LV.SelectedItem.SubItems(1) = "DECATHLON", LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2))
End Sub

Private Sub mnuCommercialPListOther_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3), True, True, LV.SelectedItem.SubItems(1) = "DECATHLON", LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2))
End Sub

Private Sub mnuCoveringLetter_Click()
    
    On Error GoTo err
    Me.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptComCoveringLetter.rpt")
    If LCase(CurrentUserName) = "mnaeem" Then
        rpt.FormulaFields(3).Text = "'Naeem Ahmed'"
    End If
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "{CustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Private Sub mnuCustomPList_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3))
End Sub

Private Sub mnuCustomPListOther_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3), True)
End Sub

Private Sub mnuDelPackingInfo_Click()
    
    If LV.ListItems.count = 0 Then Exit Sub
    Dim iAffect As Integer
    If MsgBox("Are you sure to Delete Packing List of Invoice # " & LV.SelectedItem.SubItems(3) & " ?", vbQuestion + vbYesNo) = vbYes Then
        Call StartTrans(con)
        con.Execute "Delete From CustomPList Where CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'", iAffect
        con.CommitTrans
        If iAffect > 0 Then
            LV.SelectedItem.ListSubItems(2).Tag = 0
        End If
    End If
    
End Sub

Private Sub mnuedit_Click()
    Call LV_DblClick
End Sub

Private Sub mnuEForm_Click()

    Load frmPrintFormE
    If LV.ListItems.count > 0 Then
        frmPrintFormE.ShowForm LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
    Else
        frmPrintFormE.Show , MainForm
    End If
    
End Sub

Private Sub mnuGSP_Click()

'    Load frmGSP
'    If LV.ListItems.Count > 0 Then
'        frmGSP.ShowForm LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
'    Else
'        frmGSP.Show , MainForm
'    End If
    
End Sub

Private Sub mnuPacking_Click()

'    If Left(mnuPacking.Caption, 3) = "Add" Then
'
'        Load frmNewPList
'        frmNewPList.ShowMe LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
'    Else
'        Load frmNewPList
'        Call frmNewPList.EditPList(LV.SelectedItem.SubItems(3))
'        ShowForm frmNewPList, MainForm, , "Packing List"
'    End If
    
End Sub

Private Sub mnuPrint_Click()
    Call PrintCustomInvoice
End Sub

Private Sub PrintCustomInvoice(Optional bImpMaterialDetail As Boolean = False)

    On Error GoTo err

    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    If bCustomInvoice Then
        Set rpt = rptApp.OpenReport(RptPath & "\rptCustomInvoice.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\rptCommercialInvoice.rpt")
    End If
    'rpt.FormulaFields(15).Text = bImpMaterialDetail
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VrptCustomInvoiceDetail.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt
    
    Screen.MousePointer = vbDefault


    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub mnuPrint10_Click()
    Call PrintCommercialInvoice(True, False)
End Sub

Private Sub mnuPrint5_Click()
    Call PrintCommercialInvoice(True, True)
End Sub

Private Sub mnuPrintBillofExchange_Click()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptBillOfExchange.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VRptProformas.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintComInv_Click()
    Call PrintCommercialInvoice
End Sub

Private Sub PrintCommercialInvoice(Optional bPer As Boolean = False, Optional bFivePer As Boolean = True)
    
    On Error GoTo err

    Screen.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt
    Dim rpt As CRAXDDRT.Report
    
    If UCase(LV.SelectedItem.SubItems(1)) = "MIKASA" And LV.SelectedItem.SubItems(2) = "JAPAN" Then
        Set rpt = rptApp.OpenReport(rptexportpath & "\rptCommercialInvoiceMikasaJP.rpt")
    ElseIf bPer Then
        Set rpt = rptApp.OpenReport(rptexportpath & "\rptCommercialInvoiceFairTrade.rpt")
        rpt.FormulaFields(22).Text = bFivePer
    ElseIf LV.SelectedItem.SubItems(1) = "DECATHLON" Then
        Set rpt = rptApp.OpenReport(rptexportpath & "\rptDecathlonCommercialInvoice.rpt")
    Else
        Set rpt = rptApp.OpenReport(rptexportpath & "\rptCommercialInvoice.rpt")
    End If
    f.ShowReport "{VrptCustomInvoiceDetail.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt
    
    If Not bPer Then
        If LV.SelectedItem.SubItems(2) = "Canada" Then
            Dim rpt1 As CRAXDDRT.Report
            Set rpt1 = rptApp.OpenReport(rptexportpath & "\rptCommercialInvoice.rpt")
            Dim f1 As New frmPrevRpt
            Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
            Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
            
            Set FormulaFields = rpt1.FormulaFields
            
            For Each FormulaField In FormulaFields
                If FormulaField.Name = "{@forHeading}" Then
                    FormulaField.Text = "'C A N A D I A N   C U S T O M S   I N V O I C E'"
                    Exit For
                End If
            Next
            f1.ShowReport "{VrptCustomInvoiceDetail.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt1
        End If
    End If
    Screen.MousePointer = vbDefault

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
        strFromTo = Format(DateAdd("d", iNumber, ServerDate), "dd-MMM-yyyy") & " To  " & Format(ServerDate, "dd-MMM-yyyy")
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptFProformaList.rpt")
    rpt.FormulaFields(3).Text = "'" & cmbCust.Text & "'"
    rpt.FormulaFields(4).Text = "'" & strFromTo & "'"
    
    frmPrevRpt.ShowReport "", rpt
    
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Private Sub mnuPrintCoverLetter_Click()

    On Error GoTo err

    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptDecathlonCoveringLetter.rpt")
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VCustomInvoice.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt

    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub mnuPrintDecDocs_Click()
    Load frmPrintDecathlonDocs
    frmPrintDecathlonDocs.ShowMe (LV.SelectedItem.SubItems(3))
End Sub

Private Sub mnuPrintInvWDetail_Click()

    Call PrintCustomInvoice(True)
    
End Sub

Private Sub mnuPrintPackingLabels_Click()

    On Error GoTo err
    
    If LV.SelectedItem.SubItems(1) <> "DECATHLON" Then
        MsgBox "Packing Labels are available only for DECATHLON at the moment.", vbInformation
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report, f As New frmPrevRpt
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptDecathlonPackingLabels.rpt")
    f.ShowReport "{VCustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    
    Me.MousePointer = vbNormal
    
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub mnuShippingInstructions_Click()
'    Load frmShippingInstructions
'    If LV.ListItems.Count = 0 Then
'        frmShippingInstructions.Show , MainForm
'    Else
'        frmShippingInstructions.ShowForm LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
'    End If
End Sub

Private Sub mnuUpdateCustomPrice_Click()
    If LV.ListItems.count = 0 Then Exit Sub
    If MsgBox("Are you Sure, You want to Update Custom Prices?", vbInformation + vbYesNo) = vbNo Then Exit Sub
    
    Dim dCustomPrice As Double
    dCustomPrice = Val(con.Execute("Select SUM(Qty*Price)/SUM(Qty) From CustomInvoiceItems Where CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'").Fields(0).Value & "")
    con.Execute "Update CustomInvoiceItems SET CustomPrice=" & dCustomPrice & " WHERE CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'"
    
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

Private Sub mnuPrintList_Click()

    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\IssList.rpt")
    Dim strList As String
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                strList = strList & .Tag & ","
            End If
        End With
    Next
    Dim f As New frmPrevRpt
    f.ShowReport "{VendIssued.EntryID} IN[" & Left(strList, Len(strList) - 1) & "]", rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintPurchaseOrder_Click()
    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\IssSlip.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VendIssued.EntryID}=" & LV.SelectedItem.Tag, rpt
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub LVCats_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    Dim i As Integer
    If Item.Checked Then
        If Item.Index = 1 Then
            For i = 2 To LVCats.ListItems.count
                LVCats.ListItems(i).Checked = False
            Next
        Else
            LVCats.ListItems(1).Checked = False
        End If
    Else
'        If Item.Index > 1 Then
'            LVCats.ListItems(1).Checked = False
'        End If
    End If
    
End Sub

Private Sub LVManualMaterial_KeyDown(KeyCode As Integer, Shift As Integer)

    If LVManualMaterial.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo) = vbYes Then
            LVManualMaterial.ListItems.Remove LVManualMaterial.SelectedItem.Index
        End If
    End If
    
    Call getManualDedTotal
End Sub

Private Sub getManualDedTotal()
    Dim lTotal As Long
    For i = 1 To LVManualMaterial.ListItems.count
        lTotal = lTotal + Val(LVManualMaterial.ListItems(i).SubItems(2))
    Next
    txtManualDed = lTotal
End Sub

Private Sub LVMaterial_DblClick()
    If LVMaterial.ListItems.count = 0 Then Exit Sub
    With txtEditMaterialRate
        .Move LVMaterial.Left + LVMaterial.ColumnHeaders(6).Left, LVMaterial.Top + LVMaterial.SelectedItem.Top, LV.ColumnHeaders(6).Width
        .Text = Val(LVMaterial.SelectedItem.SubItems(5))
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub LVMaterial_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    getMaterialDedTotal
    
End Sub

Private Sub LVMaterial_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If LVMaterial.ListItems.count = 0 Then Exit Sub
    If Button <> 2 Then Exit Sub
    Dim strRet As String
    strRet = ShowPopUpMenu1("Hide from Billing")
    
    If strRet = "Hide from Billing" Then
        con.Execute "INSERT INTO RawMaterialIssuanceDetail_Hidden_From_Maker_Billing(RMID_RefID,UserName,MachineName) VALUES(" & Val(LVMaterial.SelectedItem.key) & ",'" & CurrentUserName & "','" & strComputerName & "')"
        LVMaterial.ListItems.Remove LVMaterial.SelectedItem.Index
   End If
    
End Sub

Private Sub LVRepair_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call GetRepairTotals
End Sub

Private Sub txtDeductShort_Validate(cancel As Boolean)
    If Val(txtDeductShort.Text) > Val(txtDeductShort.Tag) Then
        MsgBox "Invalid Amount, Max. Deduction Amount can be " & Val(txtDeductShort.Tag)
        cancel = True
    End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Val(txtEdit) <= 0 Then
            MsgBox "Invalid Rate.", vbInformation
            Exit Sub
        End If
        Dim iSelectedID As Long
        iSelectedID = Val(LV.SelectedItem.Tag)
        con.Execute "UPDATE VendIssdDetail SET Rate=" & Val(txtEdit) & " WHERE RefID=" & Val(LV.SelectedItem.Tag)
        txtEdit.Visible = False
        Call RefreshLV
        'Now Try to Re-Select the Row
        Dim i As Integer
        For i = 1 To LV.ListItems.count
            If Val(LV.ListItems(i).Tag) = iSelectedID Then
                LV.ListItems(i).Selected = True
                LV.ListItems(i).EnsureVisible
                LV.SetFocus
            End If
        Next
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub

Private Sub txtEditMaterialRate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LVMaterial.SelectedItem.SubItems(5) = Val(txtEditMaterialRate)
        LVMaterial.SelectedItem.SubItems(7) = Val(txtEditMaterialRate) * Val(LVMaterial.SelectedItem.SubItems(6))
        con.Execute "UPDATE RawMaterialIssuanceDetail SET Price=" & Val(txtEditMaterialRate) & " WHERE EntryID=" & Val(LVMaterial.SelectedItem.key)
        Call getMaterialDedTotal
        txtEditMaterialRate.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEditMaterialRate.Visible = False
    End If
End Sub

Private Sub txtEditMaterialRate_LostFocus()
    txtEditMaterialRate.Visible = False
End Sub

Private Sub txtFixLoan_Validate(cancel As Boolean)
    If Val(txtFixLoan.Text) > Val(txtFixLoan.Tag) Then
        MsgBox "Invalid Amount, Max. Deduction Amount can be " & Val(txtFixLoan.Tag)
        cancel = True
    End If
End Sub

Private Sub txtLotNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim i As Integer
        For i = 1 To LV.ListItems.count
            If LV.ListItems(i).SubItems(16) = txtLotNo Then
                LV.ListItems(i).Selected = True
                LV.ListItems(i).EnsureVisible
                LV.SetFocus
                LV_DblClick
            End If
        Next
    End If
End Sub

Private Sub txtManualMatRate_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Dim lTotal As Long
        If txtManualMatDesc = "" Then
            MsgBox "Please Enter Description.", vbInformation
            Exit Sub
        ElseIf Val(txtManualMatRate) <= 0 Then
            If getDBPassword("ManualAddInMakerBilling") = False Then
                Exit Sub
            End If
        End If
        Dim ITM As ListItem
        Set ITM = LVManualMaterial.ListItems.add(, , LVManualMaterial.ListItems.count + 1)
        ITM.ListSubItems.add , , txtManualMatDesc.Text
        ITM.ListSubItems.add , , Val(txtManualMatRate)
    End If
    getManualDedTotal
    
End Sub

Private Sub LVHL_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If LVHL.ListItems.count = 0 Then Exit Sub
    If Button <> 2 Then Exit Sub
    Dim strRet As String
    strRet = ShowPopUpMenu1("Show in Billing", "Print PTC")
    
    If strRet = "Show in Billing" Then
        con.Execute "DELETE FROM MakerBilling_Receivings_Hidden WHERE EntryID=" & Val(LVHL.SelectedItem.Tag)
        LVHL.ListItems.Remove LVHL.SelectedItem.Index
    ElseIf strRet = "Print PTC" Then
        Call PrintPTC(LVHL.SelectedItem.Text)
   End If
   
End Sub


Private Sub ShowHiddenLots()

    Dim rs As New ADODB.Recordset, ITM As ListItem
    
    With rs
        .Open "SELECT * FROM VMakerBilling_Receivings_Hidden WHERE VendID=" & cmbCust.ID & " ORDER BY OnlyDT DESC", con, adOpenForwardOnly, adLockReadOnly
        LVHL.ListItems.Clear
        Do Until .EOF
            Set ITM = LVHL.ListItems.add(, , !LotNo & "")
            ITM.Tag = Val(!EntryID & "")
            ITM.ListSubItems.add , , Val(!RcvdQty & "")
            ITM.ListSubItems.add , , !ItemCode & ""
            ITM.ListSubItems.add , , !ItemName & ""
            ITM.ListSubItems.add , , !OrderNo & ""
            ITM.ListSubItems.add , , !CustCode & ""
            ITM.ListSubItems.add , , !Description & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub


Private Sub ShowHiddenRM()
    
    Dim rs As New ADODB.Recordset, ITM As ListItem
    
    With rs
        .Open "SELECT * FROM VRawMaterialIssuance WHERE VendID=" & cmbCust.ID & " AND EntryID IN(SELECT RMID_RefID FROM RawMaterialIssuanceDetail_Hidden_From_Maker_Billing) ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LVRMHidden.ListItems.Clear
        Do Until .EOF
            Set ITM = LVRMHidden.ListItems.add(, !EntryID & "'", LVRMHidden.ListItems.count + 1)
            ITM.Tag = Val(!MasterEntryID & "")
            ITM.ListSubItems.add , , !IssNo & ""
            ITM.ListSubItems.add , , Format(!DT, "dd-MMM-yyyy")
            ITM.ListSubItems.add , , "(" & !RMID1 & ") " & !RMName
            ITM.ListSubItems.add , , !ITemID & ""
            ITM.ListSubItems.add , , Val(!Price & "")
            ITM.ListSubItems.add , , Val(!Qty & "")
            ITM.ListSubItems.add , , Round(Val(!Qty & "") * Val(!Price & ""))
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
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
    
    If chkShowPicture.Value = vbChecked Then
        Dim rs As New ADODB.Recordset
        With rs
            .Open "SELECT ItemPic FROM Items WHERE ItemID='" & (LV.SelectedItem.SubItems(4)) & "'", con, adOpenForwardOnly, adLockReadOnly
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


