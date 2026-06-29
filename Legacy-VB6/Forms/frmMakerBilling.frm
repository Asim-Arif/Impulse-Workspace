VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "combolist.ocx"
Begin VB.Form frmMakerBilling 
   ClientHeight    =   8385
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleMode       =   0  'User
   ScaleWidth      =   10721.69
   WindowState     =   2  'Maximized
   Begin VB.Frame FAcc 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   8400
      Left            =   60
      TabIndex        =   4
      Top             =   -45
      Width           =   11865
      Begin MSComCtl2.DTPicker DTPosting 
         Height          =   315
         Left            =   2580
         TabIndex        =   50
         Top             =   7965
         Width           =   1710
         _ExtentX        =   3016
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
         Format          =   188153859
         CurrentDate     =   37055
      End
      Begin VB.TextBox txtMakerAdvancesOther 
         Alignment       =   2  'Center
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   6645
         Width           =   1455
      End
      Begin VB.TextBox txts1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   7
         Left            =   10665
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   7545
         Width           =   1125
      End
      Begin VB.TextBox txts1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   6
         Left            =   9510
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   7545
         Width           =   1125
      End
      Begin VB.TextBox txts1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   5
         Left            =   8355
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   7545
         Width           =   1125
      End
      Begin VB.TextBox txts1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   4
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   7545
         Width           =   1125
      End
      Begin VB.TextBox txts1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   3
         Left            =   6045
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   7545
         Width           =   1125
      End
      Begin VB.TextBox txts1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   2
         Left            =   4890
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   7545
         Width           =   1125
      End
      Begin VB.TextBox txts1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   1
         Left            =   3735
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   7545
         Width           =   1125
      End
      Begin VB.TextBox txts1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   0
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   7545
         Width           =   1125
      End
      Begin VB.TextBox txts 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   7
         Left            =   10665
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   7185
         Width           =   1125
      End
      Begin VB.TextBox txts 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   6
         Left            =   9510
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   7185
         Width           =   1125
      End
      Begin VB.TextBox txts 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   5
         Left            =   8355
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   7185
         Width           =   1125
      End
      Begin VB.TextBox txts 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   4
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   7185
         Width           =   1125
      End
      Begin VB.TextBox txts 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   3
         Left            =   6045
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   7185
         Width           =   1125
      End
      Begin VB.TextBox txts 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   2
         Left            =   4890
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   7185
         Width           =   1125
      End
      Begin VB.TextBox txts 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   1
         Left            =   3735
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   7185
         Width           =   1125
      End
      Begin VB.TextBox txts 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   0
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   7185
         Width           =   1125
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
         ItemData        =   "frmMakerBilling.frx":0000
         Left            =   4335
         List            =   "frmMakerBilling.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2850
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   7770
         TabIndex        =   2
         Top             =   2850
         Visible         =   0   'False
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
         Format          =   116719619
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   6045
         TabIndex        =   1
         Top             =   2850
         Visible         =   0   'False
         Width           =   1710
         _ExtentX        =   3016
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
         Format          =   116719619
         CurrentDate     =   37055
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   195
         TabIndex        =   0
         Top             =   345
         Width           =   4995
         _ExtentX        =   8811
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
      Begin TabDlg.SSTab SSTab1 
         Height          =   5775
         Left            =   210
         TabIndex        =   39
         Top             =   690
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   10186
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Maker Billing"
         TabPicture(0)   =   "frmMakerBilling.frx":0066
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "LV"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtEdit"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Repair Wastage Details"
         TabPicture(1)   =   "frmMakerBilling.frx":0082
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdPostRepair"
         Tab(1).Control(1)=   "LVRepair"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Material Deductions"
         TabPicture(2)   =   "frmMakerBilling.frx":009E
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "LVMaterial"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Manual Deductions"
         TabPicture(3)   =   "frmMakerBilling.frx":00BA
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label1(1)"
         Tab(3).Control(1)=   "LVManualMaterial"
         Tab(3).Control(2)=   "txtManualMatDesc"
         Tab(3).Control(3)=   "txtManualMatRate"
         Tab(3).ControlCount=   4
         Begin VB.TextBox txtManualMatRate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   -64585
            TabIndex        =   46
            Top             =   630
            Width           =   1000
         End
         Begin VB.TextBox txtManualMatDesc 
            Appearance      =   0  'Flat
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   -74925
            TabIndex        =   45
            Top             =   630
            Width           =   10350
         End
         Begin VB.TextBox txtEdit 
            Height          =   285
            Left            =   2325
            TabIndex        =   40
            Top             =   1485
            Visible         =   0   'False
            Width           =   690
         End
         Begin MSComctlLib.ListView LV 
            Height          =   5310
            Left            =   90
            TabIndex        =   41
            Top             =   405
            Width           =   11340
            _ExtentX        =   20003
            _ExtentY        =   9366
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
            NumItems        =   16
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S. #"
               Object.Width           =   873
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Rcv #"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Date"
               Object.Width           =   0
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
               Object.Width           =   1587
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
               Object.Width           =   1376
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   14
               Text            =   "Bal Value"
               Object.Width           =   1508
            EndProperty
            BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   15
               Text            =   "% Rcvd"
               Object.Width           =   1499
            EndProperty
         End
         Begin MSComctlLib.ListView LVRepair 
            Height          =   5280
            Left            =   -74940
            TabIndex        =   42
            Top             =   405
            Width           =   11340
            _ExtentX        =   20003
            _ExtentY        =   9313
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
               Text            =   "Rate"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Repair Amt"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "Wastage Amt"
               Object.Width           =   2469
            EndProperty
         End
         Begin MSComctlLib.ListView LVMaterial 
            Height          =   5280
            Left            =   -74925
            TabIndex        =   44
            Top             =   420
            Width           =   11340
            _ExtentX        =   20003
            _ExtentY        =   9313
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
               Text            =   "Rate"
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
               Text            =   "Amount"
               Object.Width           =   1764
            EndProperty
         End
         Begin MSComctlLib.ListView LVManualMaterial 
            Height          =   4785
            Left            =   -74925
            TabIndex        =   47
            Top             =   900
            Width           =   11340
            _ExtentX        =   20003
            _ExtentY        =   8440
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
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"frmMakerBilling.frx":00D6
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
            TabIndex        =   48
            Top             =   390
            Width           =   11340
         End
         Begin MSForms.CommandButton cmdPostRepair 
            Height          =   375
            Left            =   -74925
            TabIndex        =   43
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
         Caption         =   "Posting Date :"
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
         Left            =   1335
         TabIndex        =   49
         Top             =   8025
         Width           =   1170
      End
      Begin MSForms.CommandButton cmdPrintZero 
         Height          =   375
         Left            =   5160
         TabIndex        =   38
         Top             =   7935
         Width           =   2025
         ForeColor       =   0
         Caption         =   "Print Balance Order"
         PicturePosition =   327683
         Size            =   "3572;661"
         Accelerator     =   80
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
         Caption         =   "Deductable Advance Total :"
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
         Index           =   9
         Left            =   195
         TabIndex        =   36
         Top             =   6690
         Width           =   2310
      End
      Begin MSForms.CommandButton cmdPrintUnPosted 
         Height          =   375
         Left            =   7230
         TabIndex        =   35
         Top             =   7935
         Width           =   1800
         ForeColor       =   0
         Caption         =   "Print Un-Posted"
         PicturePosition =   327683
         Size            =   "3175;661"
         Accelerator     =   80
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
         Caption         =   "Total Finalized Bills :"
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
         Index           =   10
         Left            =   825
         TabIndex        =   27
         Top             =   7590
         Width           =   1680
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "% Rcvd."
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   8
         Left            =   10875
         TabIndex        =   25
         Top             =   6975
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Bal. Value"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   7
         Left            =   9660
         TabIndex        =   23
         Top             =   6975
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Un. Auth"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   6
         Left            =   8505
         TabIndex        =   21
         Top             =   6975
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Rcvd. Value"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   5
         Left            =   7305
         TabIndex        =   19
         Top             =   6975
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Repair"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   4
         Left            =   6345
         TabIndex        =   17
         Top             =   6975
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Temp. Adv."
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   3
         Left            =   4980
         TabIndex        =   15
         Top             =   6975
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Value"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   2
         Left            =   4065
         TabIndex        =   13
         Top             =   6975
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Maker's Total Performance :"
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
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   7230
         Width           =   2385
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Issd Qty"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   2850
         TabIndex        =   10
         Top             =   6975
         Width           =   615
      End
      Begin MSForms.CommandButton cmdPost 
         Height          =   375
         Left            =   9060
         TabIndex        =   8
         Top             =   7935
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
         Left            =   10440
         TabIndex        =   7
         Top             =   7935
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "2355;661"
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
         Height          =   360
         Left            =   10260
         TabIndex        =   3
         Top             =   315
         Width           =   1500
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "2646;635"
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
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMakerBilling.frx":01BE
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
         Left            =   165
         TabIndex        =   5
         Top             =   120
         Width           =   11130
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
            Picture         =   "frmMakerBilling.frx":0268
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakerBilling.frx":0424
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PopMenu"
   End
End
Attribute VB_Name = "frmMakerBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCourierInvoice As Boolean
Dim bShipping As Boolean
Dim bCustomInvoice As Boolean
Dim strAdvanceEntryList As String
Public Sub ShowMe(Optional p_bCustomInvoice As Boolean = True)
    bCustomInvoice = p_bCustomInvoice
    Me.Show
End Sub

Public Sub ShowForShipping()
    bShipping = True
    Me.Show
End Sub

Private Sub cmbCust_matched()
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

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdPost_Click()

    On Error GoTo err
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
        dValue = dValue + Val(LV.ListItems(i).SubItems(10))
    Next
    'Check for Un-Auth...
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                If Val(.SubItems(13)) > 0 Then
                    MsgBox "Some Selected Orders have Un-Authorized Entries, Can't Continue.", vbInformation
                    Exit Sub
                End If
            End If
        End With
    Next
    'Get Bill No...
    Dim lBillNo As Long, strMakerID As String, strDesc As String
    strMakerID = GetSingleStringValue("VendID1", "Makers", " WHERE VendID=" & cmbCust.id)
    lBillNo = GetSingleLongValue("MAX(BillNo)", "MakerPostedBills", " WHERE BillNo IS NOT NULL AND VendID=" & cmbCust.id)
    lBillNo = lBillNo + 1
   
    'Now Post Vouchers....
    Dim strVchrNo As String
    'Services / Work Order AccNo: 41-003-19001
    Dim strDebitAccNo As String, strMakerAccNo As String, dManualDedAmt As Double
    
    For i = 1 To LVManualMaterial.ListItems.count
        dManualDedAmt = dManualDedAmt + Val(LVManualMaterial.ListItems(i).SubItems(2))
    Next
    
    If strCompany = "Sunlike" Then
        strDebitAccNo = "41-001-16005"
    Else
        strDebitAccNo = "41-003-19001"
    End If
    strMakerAccNo = GetSingleStringValue("AccNo", "Makers", " WHERE VendID=" & cmbCust.id)
     
    Dim objVchr As New clsVouchers
    objVchr.VoucherDate = DTPosting 'GetServerDate
    objVchr.VoucherType = "JV"
    
    Dim dAmtToPost As Double
    dAmtToPost = Val(txts1(4)) - Val(txtMakerAdvancesOther) - dManualDedAmt
    
    If dAmtToPost <= 0 Then
        MsgBox "Invalid Amount to Post.", vbInformation
        Exit Sub
    End If
    
    strDesc = "Computer Bill No. " & strMakerID & " / " & lBillNo & " - " & year(DTPosting)
    objVchr.AddVoucherHead strDebitAccNo, dAmtToPost, strDesc
    
    objVchr.AddVoucherHead strMakerAccNo, -(dAmtToPost - Val(txts1(3))), strDesc
    
    If Val(txts1(3)) <> 0 Then  'If Repair....
        objVchr.AddVoucherHead "24-005-19001", -Val(txts1(3)), strDesc
    End If
     
    Call StartTrans(con)
    strVchrNo = objVchr.PostVoucher(, True)
        
    Dim dRepairAmt As Double, dWastageAmt As Double, strRepairVchrNo As String
    Call PostRepairVchr(strRepairVchrNo, dRepairAmt, dWastageAmt)
    
    con.Execute "INSERT INTO MakerPostedBills(VendID,DT,VchrNo,UserName,MachineName,BillNo,TempAdvance,BalanceOrderValue," & _
    "UnAuthValue,RepairVchrNo,RepairAmt,WastageAmt,ManualDedAmt) VALUES(" & cmbCust.id & ",'" & DTPosting & "','" & strVchrNo & "','" & CurrentUserName & "','" & _
     strComputerName & "'," & lBillNo & "," & Val(txtMakerAdvancesOther) & "," & Val(txts(6)) - Val(txts1(6)) & _
     "," & Val(txts(5)) - Val(txts1(5)) & ",'" & strRepairVchrNo & "'," & dRepairAmt & "," & dWastageAmt & "," & dManualDedAmt & ")"
    
    Dim lEntryID As Long
    lEntryID = GetSingleLongValue("Max(EntryID)", "MakerPostedBills")
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                con.Execute "INSERT INTO MakerPostedBillsDetail(RefID,Iss_RefID,TempAdvance," & _
                 "RepairAmt,RcvdValue,RcvdQty,UnAuthValue,BalanceValue,RcvdPercentage,IssDetail_RefID) VALUES(" & lEntryID & "," & Val(.Tag) & "," & _
                 Val(.SubItems(9)) & "," & Val(.SubItems(10)) & "," & Val(.SubItems(11)) & "," & Val(.SubItems(12)) & _
                 "," & Val(.SubItems(13)) & "," & Val(.SubItems(14)) & "," & Val(.SubItems(15)) & "," & Val(.key) & ")"
             
                con.Execute "UPDATE VendIssued SET VchrNo='" & strVchrNo & "' WHERE EntryID=" & .Tag
            End If
        End With
    Next
        
    If strCompany = "PAKSMITH GROUP" Or strCompany = "QEL" Then
        'Now Update TemporaryLoanGroups
        con.Execute "UPDATE MakerTemporaryLoansGroup SET Active=0 WHERE EntryID IN(SELECT GroupID FROM MakerTemporaryLoans WHERE Issuance_RefID IN(" & strAdvanceEntryList & "))"
    End If
        
    'Now Save Manual Material Ded. Entries...
    For i = 1 To LVManualMaterial.ListItems.count
        With LVManualMaterial.ListItems(i)
            con.Execute "INSERT INTO MakerPostedBills_ManualDeductions(Description,Amount,MPB_RefID) VALUES('" & _
             .SubItems(1) & "'," & Val(.SubItems(2)) & "," & lEntryID & ")"
        End With
    Next
    
    con.CommitTrans
        
    LVManualMaterial.ListItems.Clear
    MsgBox "Posted Successfuly !!!", vbInformation
    Call PrintMakerPostedBill(strVchrNo)
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
                dRepairAmt = dRepairAmt + Val(.SubItems(5))
                dWastageAmt = dWastageAmt + Val(.SubItems(6))
            End If
        End With
    Next
    'Now Post...
    Dim dAmt As Double, strRepairAccNo As String, strVchrNo As String, strWastageAccNo As String, strMakerAccNo As String
    Dim strDescription As String
    strRepairAccNo = "41-004-13001"
    strWastageAccNo = "41-003-19001"
    strMakerAccNo = GetSingleStringValue("AccNo", "Makers", " WHERE VendID=" & cmbCust.id)
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

Private Sub cmdPrintUnPosted_Click()
    Call PrintUnPostedBills(cmbCust.id, False)
End Sub

Private Sub PrintUnPostedBills(lMakerID As Long, Optional bShowZero As Boolean = False)
    
    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    
    Set rpt = rptApp.OpenReport(RptPath & "\MakerUnPostedBills.rpt")
    rpt.ParameterFields.GetItemByName("@MakerID").SetCurrentValue CInt(cmbCust.id)
    rpt.FormulaFields.GetItemByName("ShowMaker").Text = True
    rpt.FormulaFields.GetItemByName("UserName").Text = "'" & CurrentUserName & "'"
    rpt.FormulaFields.GetItemByName("ComputerName").Text = "'" & strComputerName & "'"
    
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    Dim strSelection As String
    If bShowZero Then
        strSelection = "{@RcvdQty}=0"
    Else
        strSelection = "{@RcvdQty}>0"
    End If
    
    f.ShowReport strSelection, rpt, , False
    
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub cmdPrintZero_Click()
    Call PrintUnPostedBills(cmbCust.id, True)
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
    
    mnupop.Visible = False
    cmbCust.ListHeight = 3800
    'cmbCust.AddItem "<All Makers>", "0"
    Call cmbCust.AddVals(con, "'{ ' + VendID1 + ' } ' + VenderName ", "Makers", "VendID", " WHERE VendID<>" & iFactoryMakerID)
    cmbCust.id = "0"
    DtFrom = Date
    DtTo = Date
    cmbRange.ListIndex = 1
    If strCompany = "Weldon Instruments" Or strCompany = "Sunlike" Or strCompany = "QEL" Or (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Banzai" Or strCompany = "Towne" Then
        SSTab1.TabVisible(1) = True
    Else
        SSTab1.TabVisible(1) = False
    End If
    
    SSTab1.Tab = 0
    
    DTPosting.Value = Date
    
End Sub

Private Sub RefreshLV()
        
    On Error GoTo err
    If cmbCust.MatchFound = False Then
        MsgBox "Please Select Maker.", vbInformation
        Exit Sub
    End If
    strAdvanceEntryList = ""
    txtMakerAdvancesOther = ""
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "MakerBilling_SP"
        .Parameters("@MakerID").Value = cmbCust.id
        
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
    For i = 0 To 7
        txts(i) = ""
        txts(i).Tag = ""
        txts1(i) = ""
        txts1(i) = ""
    Next
    With rs
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, Val(!Issue_RefID & "") & "'", .AbsolutePosition)
            ITM.Tag = Val(![EntryID] & "")
            ITM.ListSubItems.add(, , !RecieptID & "").ToolTipText = !ItemName & ""
            ITM.ListSubItems.add(, , Format(!DT, "dd-MM-yy")).ToolTipText = !ItemName & ""
            ITM.ListSubItems.add(, , !OrderNo & "").ToolTipText = !ItemName & ""
            ITM.ListSubItems.add(, , !ITemID & "").ToolTipText = !ItemName & ""
            ITM.ListSubItems.add(, , !ItemName & "").ToolTipText = !ItemName & ""
            ITM.ListSubItems.add(, , Val(!IssQty & "")).ToolTipText = !ItemName & ""
            ITM.ListSubItems.add(, , Val(!Rate & "")).ToolTipText = !ItemName & ""
            ITM.ListSubItems.add(, , Val(!Value & "")).ToolTipText = !ItemName & ""
            ITM.ListSubItems.add(, , Val(!TempAdvance & "")).ToolTipText = !ItemName & ""
            ITM.ListSubItems.add(, , Val(!RepairAmt & "")).ToolTipText = !ItemName & ""
            ITM.ListSubItems.add(, , Round(Val(!RcvdValue & ""))).Tag = Val(!RcvdQty & "") 'Round(Val(!RcvdQty & "") * Val(!Price & ""), 2)
            ITM.ListSubItems.add(, , Val(!RcvdQty & "")).Tag = Val(!RcvdQtyUnAuth & "")
            ITM.ListSubItems.add(, , Round(Val(!RcvdValueUnAuth & ""))).Tag = Val(!RcvdQtyUnAuth & "") 'Round(Val(!RcvdQtyUnAuth & "") * Val(!Price & ""), 2)
            ITM.ListSubItems.add , , Round(Val(!Value & "") - Val(!RcvdValue & "") - Val(!RcvdValueUnAuth & "") - Val(!RepairAmt & "")) 'Round(Val(!value & "") - Val(ITM.SubItems(8)) - Val(ITM.SubItems(9)), 2)
            ITM.ListSubItems.add , , Round(Val(!PercentRcvd & "")) & "%"
            
            
            txts(0) = Val(txts(0)) + Val(ITM.SubItems(6))   'IssQty
            txts(1) = Val(txts(1)) + Val(ITM.SubItems(8))   'Value
            txts(2) = Val(txts(2)) + Val(ITM.SubItems(9))   'TempAdvance
            txts(3) = Val(txts(3)) + Val(ITM.SubItems(10))   'RepairAmt
            txts(4) = Val(txts(4)) + Val(ITM.SubItems(11))   'RcvdValue
            txts(4).Tag = Val(txts(4).Tag) + Val(ITM.ListSubItems(11).Tag)   'RcvdQty
            txts(5) = Val(txts(5)) + Val(ITM.SubItems(13))  'RcvdValueUnAuth
            txts(5).Tag = Val(txts(5).Tag) + Val(ITM.ListSubItems(12).Tag)   'RcvdQtyUnAuth
            txts(6) = Val(txts(6)) + Val(ITM.SubItems(14))  'BalValue
            .MoveNEXT
        Loop
        .Close
    End With
    If Val(txts(0)) > 0 Then txts(7) = Round(((Val(txts(4).Tag) + Val(txts(5).Tag)) / Val(txts(0))) * 100) & "%"
    Set rs = Nothing
            
    If strCompany = "Weldon Instruments" Or strCompany = "Sunlike" Or strCompany = "QEL" Or (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Banzai" Or strCompany = "Towne" Then
        Call AddRepairDetail
    End If
    If strCompany = "QEL" Or (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Banzai" Or strCompany = "Towne" Then
        Call AddMaterialIssuanceDetail
    End If
    
    LVManualMaterial.ListItems.Clear
    
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Private Sub AddRepairDetail()

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "SELECT * FROM VMakerRepair WHERE VendID=" & cmbCust.id & " AND VchrNo IS NULL", con, adOpenForwardOnly, adLockReadOnly
        LVRepair.ListItems.Clear
        Do Until .EOF
            Set ITM = LVRepair.ListItems.add(, !EntryID & "'", LVRepair.ListItems.count + 1)
            ITM.ListSubItems.add , , !RecieptID & ""
            ITM.ListSubItems.add , , Format(!DT, "dd-MMM-yyyy")
            ITM.ListSubItems.add , , !ItemCode & ""
            ITM.ListSubItems.add , , Val(!Rate & "")
            ITM.ListSubItems.add , , Val(!RepairAmt & "")
            ITM.ListSubItems.add , , Val(!Qty & "") * Val(!Rate & "")
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub


Private Sub AddMaterialIssuanceDetail()

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "SELECT * FROM VRawMaterialIssuance WHERE VendID=" & cmbCust.id & " AND VchrNo IS NULL", con, adOpenForwardOnly, adLockReadOnly
        LVMaterial.ListItems.Clear
        Do Until .EOF
            Set ITM = LVMaterial.ListItems.add(, !EntryID & "'", LVMaterial.ListItems.count + 1)
            ITM.ListSubItems.add , , !IssNo & ""
            ITM.ListSubItems.add , , Format(!DT, "dd-MMM-yyyy")
            ITM.ListSubItems.add , , "(" & !RMID1 & ") " & !RMName
            ITM.ListSubItems.add , , Val(!Price & "")
            ITM.ListSubItems.add , , Val(!Qty & "")
            ITM.ListSubItems.add , , Val(!Qty & "") * Val(!Price & "")
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    FAcc.Top = (Me.ScaleHeight - FAcc.Height) / 2
End Sub


Private Sub LV_DblClick()
    'Edit Rate
    If LV.ListItems.count = 0 Then Exit Sub
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
    For i = 0 To 7
        txts1(i) = ""
        txts1(i).Tag = ""
    Next
    
    strAdvanceEntryList = ""
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                txts1(0) = Val(txts1(0)) + Val(.SubItems(6))   'IssQty
                txts1(1) = Val(txts1(1)) + Val(.SubItems(8))   'Value
                txts1(2) = Val(txts1(2)) + Val(.SubItems(9))   'TempAdvance
                txts1(3) = Val(txts1(3)) + Val(.SubItems(10))   'RepairAmt
                txts1(4) = Val(txts1(4)) + Val(.SubItems(11))   'RcvdValue
                txts1(4).Tag = Val(txts1(4).Tag) + Val(.ListSubItems(12).Tag)   'RcvdQty
                txts1(5) = Val(txts1(5)) + Val(.SubItems(12))  'RcvdValueUnAuth
                txts1(5).Tag = Val(txts1(5).Tag) + Val(.ListSubItems(12).Tag)   'RcvdQtyUnAuth
                txts1(6) = Val(txts1(6)) + Val(.SubItems(13))  'BalValue
                '
                strAdvanceEntryList = strAdvanceEntryList & Val(.Tag) & ","
            End If
        End With
    Next
     
    If strCompany = "PAKSMITH GROUP" Or strCompany = "QEL" Then
        If strAdvanceEntryList <> "" Then
            strAdvanceEntryList = Left(strAdvanceEntryList, Len(strAdvanceEntryList) - 1)
            txtMakerAdvancesOther = GetSingleLongValue("SUM(Amount)", "MakerTemporaryLoans", " WHERE GroupID IN(SELECT GroupID FROM MakerTemporaryLoans WHERE Issuance_RefID IN(" & strAdvanceEntryList & ")) AND GroupID NOT IN(SELECT EntryID FROM MakerTemporaryLoansGroup WHERE Active=0)")
        Else
            txtMakerAdvancesOther = ""
        End If
    End If
    
    If Val(txts1(0)) = 0 Then Exit Sub
    txts1(7) = Round(((Val(txts1(4).Tag) + Val(txts1(5).Tag)) / Val(txts1(0))) * 100) & "%"
    
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

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button <> 2 Then Exit Sub
    Exit Sub
    If LV.ListItems.count = 0 Then
        Call ShowMenus(False)
    Else
        Call ShowMenus(True)
    End If
'    If bShipping Then
'        ShowMenusForShipping
'    ElseIf bCourierInvoice Then
'        ShowMenusforCourier
'    Else
'        If LV.ListItems.Count = 0 Then
'            Call ShowMenus(False)
'        Else
'            Call ShowMenus(True)
'        End If
'    End If
    'If LV.ListItems.Count = 0 Or Button <> 2 Or OutOfFinancialYear Then Exit Sub

    'if a Voucher Entry Show Delete Menu
    'If InStr(LV.SelectedItem, "-") > 0 Then Me.PopupMenu mnuPop
    
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
    Me.PopUpMenu mnupop
    
End Sub

Private Sub ShowMenus(vis As Boolean)
        
    
    Me.PopUpMenu mnupop
    
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
    Me.PopUpMenu mnupop
    
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

Private Sub mnuEdit_Click()
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


Private Sub LVManualMaterial_KeyDown(KeyCode As Integer, Shift As Integer)

    If LVManualMaterial.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo) = vbYes Then
            LVManualMaterial.ListItems.Remove LVManualMaterial.SelectedItem.Index
        End If
    End If

End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Val(txtEdit) <= 0 Then
            MsgBox "Invalid Rate.", vbInformation
            Exit Sub
        End If
        Dim iSelectedID As Integer
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

Private Sub txtManualMatRate_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If txtManualMatDesc = "" Then
            MsgBox "Please Enter Description.", vbInformation
            Exit Sub
        ElseIf Val(txtManualMatRate) <= 0 Then
            MsgBox "Invalid Amount.", vbInformation
            Exit Sub
        End If
        Dim ITM As ListItem
        Set ITM = LVManualMaterial.ListItems.add(, , LVManualMaterial.ListItems.count + 1)
        ITM.ListSubItems.add , , txtManualMatDesc.Text
        ITM.ListSubItems.add , , Val(txtManualMatRate)
    End If
    
End Sub
