VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmUserSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Access Rights"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9165
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserSettings.frx":0000
   LinkTopic       =   "Form24"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   9165
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUnCheckAll 
      Caption         =   "Clear All"
      Height          =   330
      Left            =   1125
      TabIndex        =   48
      Top             =   8595
      Width           =   885
   End
   Begin VB.CommandButton cmdCheckAll 
      Caption         =   "Select All"
      Height          =   330
      Left            =   210
      TabIndex        =   47
      Top             =   8595
      Width           =   885
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5040
      Left            =   210
      TabIndex        =   19
      Top             =   3525
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   8890
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      Tab             =   1
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "Basic Settings"
      TabPicture(0)   =   "frmUserSettings.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LV"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmbModule"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Advance Settings"
      TabPicture(1)   =   "frmUserSettings.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "chkAdv(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "chkAdv(9)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "chkAdv(10)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "chkAdv(11)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "chkAdv(12)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "chkAdv(13)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "chkAdv(14)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "chkAdv(15)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "chkAdv(16)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "chkAdv(17)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "chkAdv(19)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "chkAdv(20)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "chkAdv(18)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "chkAdv(21)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "chkAdv(22)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "chkAdv(23)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "chkAdv(24)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "chkAdv(25)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "chkAdv(26)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "chkAdv(27)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "RM Groups"
      TabPicture(2)   =   "frmUserSettings.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LVRMGroups"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Customers"
      TabPicture(3)   =   "frmUserSettings.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "LVCustomers"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Item Groups"
      TabPicture(4)   =   "frmUserSettings.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "LVItemGroups"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Stores"
      TabPicture(5)   =   "frmUserSettings.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "LVStores"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Security"
      TabPicture(6)   =   "frmUserSettings.frx":03B2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "LVSecurity"
      Tab(6).Control(1)=   "chkShowPasswords"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "Allowed Computers"
      TabPicture(7)   =   "frmUserSettings.frx":03CE
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label4"
      Tab(7).Control(1)=   "LVAC"
      Tab(7).Control(2)=   "txtComputerName"
      Tab(7).ControlCount=   3
      Begin VB.CheckBox chkAdv 
         Caption         =   "Show Values Company Catalog"
         DataField       =   "Show_Values_Company_Catalog"
         DataSource      =   "Adodc1"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   27
         Left            =   3675
         TabIndex        =   64
         Top             =   2520
         Width           =   3060
      End
      Begin VB.CheckBox chkAdv 
         Caption         =   "Show Balance on Vendor List"
         DataField       =   "Show_Balance_Vendor_List"
         DataSource      =   "Adodc1"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   26
         Left            =   3675
         TabIndex        =   63
         Top             =   2160
         Width           =   3060
      End
      Begin VB.CheckBox chkAdv 
         Caption         =   "Maker Issuance Report (Valuewise)"
         DataField       =   "MIL_Print_Maker_Issuance_Report_Valuewise"
         DataSource      =   "Adodc1"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   25
         Left            =   3675
         TabIndex        =   62
         Top             =   1800
         Width           =   3060
      End
      Begin VB.CheckBox chkAdv 
         Caption         =   "Print Maker PO"
         DataField       =   "MIL_Print_Master_PO"
         DataSource      =   "Adodc1"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   24
         Left            =   3675
         TabIndex        =   61
         Top             =   1425
         Width           =   3060
      End
      Begin VB.CheckBox chkAdv 
         Caption         =   "Hide Rate (Maker Assigned Items)"
         DataField       =   "HideRateMakerAssign"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   23
         Left            =   3675
         TabIndex        =   60
         Top             =   3825
         Width           =   3060
      End
      Begin VB.CheckBox chkAdv 
         Caption         =   "Print Duplicate PTC"
         DataField       =   "PrintDuplicate_PTC"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   22
         Left            =   3675
         TabIndex        =   59
         Top             =   1020
         Width           =   2940
      End
      Begin VB.CheckBox chkAdv 
         Caption         =   "Maker Billing (Edit Rate)"
         DataField       =   "MakerBill_EditRate"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   21
         Left            =   3675
         TabIndex        =   58
         Top             =   4110
         Width           =   3060
      End
      Begin VB.CheckBox chkAdv 
         Caption         =   "MR"
         DataField       =   "CloseCAPA"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   18
         Left            =   3675
         TabIndex        =   57
         Top             =   4395
         Width           =   2580
      End
      Begin VB.TextBox txtComputerName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -74910
         TabIndex        =   54
         Top             =   690
         Width           =   8655
      End
      Begin VB.CheckBox chkShowPasswords 
         Caption         =   "Show Passwords"
         Height          =   255
         Left            =   -74880
         TabIndex        =   53
         Top             =   4695
         Width           =   1695
      End
      Begin VB.CheckBox chkAdv 
         Caption         =   "Restrict Company Catalog Editing"
         DataField       =   "RestrictCompanyCatalogEditing"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   20
         Left            =   3675
         TabIndex        =   51
         Top             =   600
         Width           =   2940
      End
      Begin VB.CheckBox chkAdv 
         Caption         =   "User Can Change Password"
         DataField       =   "ChangePassword"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   19
         Left            =   3675
         TabIndex        =   49
         Top             =   4680
         Width           =   2580
      End
      Begin VB.CheckBox chkAdv 
         Caption         =   "Edit Follow-up"
         DataField       =   "EditFollowup"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   17
         Left            =   210
         TabIndex        =   46
         Top             =   4395
         Width           =   3060
      End
      Begin VB.CheckBox chkAdv 
         Caption         =   "Open Command Center at Login"
         DataField       =   "OpenCommandCenter"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   16
         Left            =   210
         TabIndex        =   45
         Top             =   4680
         Width           =   3060
      End
      Begin VB.CheckBox chkAdv 
         Caption         =   "Add/Edit Customer Complaint"
         DataField       =   "AddEditCustomerComplaint"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   15
         Left            =   210
         TabIndex        =   41
         Top             =   4110
         Width           =   3060
      End
      Begin VB.CheckBox chkAdv 
         Caption         =   "Authorize Issuance"
         DataField       =   "AuthorizeIssuance"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   14
         Left            =   210
         TabIndex        =   39
         Top             =   3825
         Width           =   3060
      End
      Begin MSComctlLib.ListView LVRMGroups 
         Height          =   4530
         Left            =   -74895
         TabIndex        =   38
         Top             =   435
         Width           =   8685
         _ExtentX        =   15319
         _ExtentY        =   7990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S #"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Group"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.CheckBox chkAdv 
         Caption         =   "Maker Billing (Post Bill)"
         DataField       =   "PostMakerBill"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   13
         Left            =   210
         TabIndex        =   37
         Top             =   3540
         Width           =   3060
      End
      Begin VB.CheckBox chkAdv 
         Caption         =   "Customer Catalog Editing"
         DataField       =   "EditCustCatalog"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   12
         Left            =   210
         TabIndex        =   36
         Top             =   3255
         Width           =   3060
      End
      Begin VB.CheckBox chkAdv 
         Caption         =   "Authorize Vouchers"
         DataField       =   "AuthorizeVouchers"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   11
         Left            =   210
         TabIndex        =   35
         Top             =   2970
         Width           =   3060
      End
      Begin VB.CheckBox chkAdv 
         Caption         =   "Change Rate Maker Assigned Items"
         DataField       =   "ChangeRateMakerAssign"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   10
         Left            =   210
         TabIndex        =   33
         Top             =   2685
         Width           =   3060
      End
      Begin VB.CheckBox chkAdv 
         Caption         =   "Change Rate on Issuance"
         DataField       =   "RestrictedItemProfile"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   9
         Left            =   210
         TabIndex        =   32
         Top             =   2400
         Width           =   2670
      End
      Begin VB.Frame Frame1 
         Caption         =   "Item Profile"
         ForeColor       =   &H00404080&
         Height          =   1380
         Left            =   165
         TabIndex        =   34
         Top             =   870
         Width           =   3225
         Begin VB.CheckBox chkAdv 
            Caption         =   "Ship. Info."
            DataField       =   "ShipInfoItemProfile"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   1755
            TabIndex        =   31
            Top             =   1080
            Width           =   1440
         End
         Begin VB.CheckBox chkAdv 
            Caption         =   "Raw Materials"
            DataField       =   "RMItemProfile"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   60
            TabIndex        =   30
            Top             =   1080
            Width           =   1440
         End
         Begin VB.CheckBox chkAdv 
            Caption         =   "Picture"
            DataField       =   "PictureItemProfile"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   1755
            TabIndex        =   29
            Top             =   815
            Width           =   1440
         End
         Begin VB.CheckBox chkAdv 
            Caption         =   "References"
            DataField       =   "ReferencesItemProfile"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   60
            TabIndex        =   28
            Top             =   815
            Width           =   1440
         End
         Begin VB.CheckBox chkAdv 
            Caption         =   "Price"
            DataField       =   "PriceItemProfile"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   1755
            TabIndex        =   27
            Top             =   550
            Width           =   1440
         End
         Begin VB.CheckBox chkAdv 
            Caption         =   "Weight"
            DataField       =   "WeightItemProfile"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   60
            TabIndex        =   26
            Top             =   550
            Width           =   1440
         End
         Begin VB.CheckBox chkAdv 
            Caption         =   "Processes"
            DataField       =   "ProcessesItemProfile"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   1755
            TabIndex        =   25
            Top             =   285
            Width           =   1440
         End
         Begin VB.CheckBox chkAdv 
            Caption         =   "General Info"
            DataField       =   "GeneralInfoItemProfile"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   24
            Top             =   285
            Width           =   1440
         End
      End
      Begin VB.CheckBox chkAdv 
         Caption         =   "Restricted Item Profile Access"
         DataField       =   "RestrictedItemProfile"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   23
         Top             =   600
         Width           =   2670
      End
      Begin VB.ComboBox cmbModule 
         Height          =   315
         Left            =   -74940
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   570
         Width           =   8715
      End
      Begin MSComctlLib.ListView LV 
         Height          =   4005
         Left            =   -74925
         TabIndex        =   21
         Top             =   900
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   7064
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15199215
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Option"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Module"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView LVCustomers 
         Height          =   4530
         Left            =   -74880
         TabIndex        =   40
         Top             =   435
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   7990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S #"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer"
            Object.Width           =   7056
         EndProperty
      End
      Begin MSComctlLib.ListView LVItemGroups 
         Height          =   4530
         Left            =   -74895
         TabIndex        =   43
         Top             =   405
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   7990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S #"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item Groups"
            Object.Width           =   7056
         EndProperty
      End
      Begin MSComctlLib.ListView LVStores 
         Height          =   4530
         Left            =   -74880
         TabIndex        =   44
         Top             =   405
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   7990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S #"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Store Name"
            Object.Width           =   7056
         EndProperty
      End
      Begin MSComctlLib.ListView LVSecurity 
         Height          =   4275
         Left            =   -74880
         TabIndex        =   52
         Top             =   390
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   7541
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S #"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Security"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Password"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView LVAC 
         Height          =   3840
         Left            =   -74910
         TabIndex        =   55
         Top             =   1080
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   6773
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S #"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer"
            Object.Width           =   10583
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Computer Name"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74910
         TabIndex        =   56
         Top             =   375
         Width           =   1680
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Module :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   -74925
         TabIndex        =   22
         Top             =   360
         Width           =   705
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Main Links"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2295
      Left            =   615
      TabIndex        =   11
      Top             =   1140
      Width           =   2955
      Begin VB.CheckBox Check1 
         Caption         =   "Sampling"
         DataField       =   "SamplingMainLink"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   10
         Left            =   150
         TabIndex        =   50
         Top             =   1980
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Fixed Assets"
         DataField       =   "FixedAssetsMainLink"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   9
         Left            =   1545
         TabIndex        =   10
         Top             =   1650
         Width           =   1305
      End
      Begin VB.CheckBox Check1 
         Caption         =   "QMS"
         DataField       =   "QMSMainLink"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   150
         TabIndex        =   9
         Top             =   1650
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Help"
         DataField       =   "HelpMainLink"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   1545
         TabIndex        =   8
         Top             =   1299
         Width           =   1305
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Dash Board"
         DataField       =   "DashBoardMainLink"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   165
         TabIndex        =   7
         Top             =   1299
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Production"
         DataField       =   "ProductionMainLink"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   1545
         TabIndex        =   6
         Top             =   951
         Width           =   1305
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Payroll"
         DataField       =   "PayrollMainLink"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   1545
         TabIndex        =   4
         Top             =   603
         Width           =   1305
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Export"
         DataField       =   "ExportMainLink"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   3
         Top             =   603
         Width           =   1080
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Financial"
         DataField       =   "FinancialMainLink"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   5
         Top             =   951
         Width           =   1080
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Stock"
         DataField       =   "StockMainLink"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   1545
         TabIndex        =   2
         Top             =   255
         Width           =   1305
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Company"
         DataField       =   "CompanyMainLink"
         DataSource      =   "Adodc1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   1
         Top             =   255
         Width           =   1080
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   45
      Top             =   1995
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtUserName 
      DataField       =   "FullUserName"
      DataSource      =   "Adodc1"
      Height          =   300
      Left            =   615
      TabIndex        =   0
      Top             =   810
      Width           =   2955
   End
   Begin VB.Label lblSig 
      Alignment       =   2  'Center
      Caption         =   "Double Click To Add Picture."
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   6600
      TabIndex        =   42
      Top             =   2730
      Width           =   1095
   End
   Begin VB.Image ImgSignature 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   900
      Left            =   5385
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   3600
   End
   Begin VB.Label lblPic 
      Alignment       =   2  'Center
      Caption         =   "Double Click To Add Picture."
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   6675
      TabIndex        =   18
      Top             =   1140
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User Picture :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5385
      TabIndex        =   17
      Top             =   60
      Width           =   1125
   End
   Begin VB.Image ImgUser 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2250
      Left            =   5385
      Stretch         =   -1  'True
      Top             =   270
      Width           =   3600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "User Complete Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   645
      TabIndex        =   16
      Top             =   600
      Width           =   1770
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   375
      Left            =   7650
      TabIndex        =   14
      Top             =   8655
      Width           =   1425
      Caption         =   " Close          "
      PicturePosition =   327683
      Size            =   "2514;661"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdSave 
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   8655
      Width           =   1425
      Caption         =   "OK           "
      PicturePosition =   327683
      Size            =   "2514;661"
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label USer 
      AutoSize        =   -1  'True
      Caption         =   "User Login Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   675
      TabIndex        =   15
      Top             =   60
      Width           =   1410
   End
   Begin VB.Label lblUserName 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "UserName"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   615
      TabIndex        =   13
      Top             =   285
      Width           =   2955
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   90
      Picture         =   "frmUserSettings.frx":03EA
      Top             =   -30
      Width           =   540
   End
End
Attribute VB_Name = "frmUserSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsUser As New ADODB.Recordset
Dim Loaded As Boolean
Dim ColIndex As Integer
Dim PicFileName As String, SigFileName As String
Dim IntUserID As Long
Dim strDelACList As String

Private Sub AddSignature()

    On Error Resume Next
    Dim Picfile As String
    CD1.ShowSave
    If Not (err.Number = 32755) Then
        Picfile = CD1.FileName
        img.Picture = LoadPicture(Picfile)
        PicFileName = Picfile
        Label3.Visible = False
    End If
    
End Sub

Private Sub chkShowPasswords_Click()

    If chkShowPasswords.Value = vbChecked Then
    
        Dim strAdminPWD As String
        strAdminPWD = GetSingleStringValue("Password", "Users", " WHERE UserName='Administrator'")
        If getPasswordNew(strAdminPWD) Then
            Dim i As Integer, strTemp As String
            For i = 1 To LVSecurity.ListItems.count
                LVSecurity.ListItems(i).SubItems(2) = LVSecurity.ListItems(i).ListSubItems(2).Tag
            Next
            chkShowPasswords.Visible = False
        End If
        
    End If
    
End Sub

Private Sub cmbModule_Click()
    Call RefreshList
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo err
    rsUser.CancelUpdate
    Unload Me
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdCheckAll_Click()
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmdSave_Click()
    
    On Error GoTo err
'
'    Call StartTrans(con)
'    If rsUser![UserID] = 1 Then GoTo DoItForAdmin
'
'    rsUser.Update
'
'
'    con.Execute "Delete From UserAssProcesses Where UserID=" & rsUser![UserID]
'
'    For i = 1 To LV.ListItems.count
'        If LV.ListItems(i).Checked Then
'            con.Execute "Insert Into UserAssProcesses(UserID,ProcessID,ReqIssAuth,ReqRcvAuth) Values(" & _
'             rsUser![UserID] & "," & Val(LV.ListItems(i).key) & "," & IIf(LV.ListItems(i).SubItems(1) = "Yes", 1, 0) & "," & IIf(LV.ListItems(i).SubItems(1) = "Yes", 1, 0) & ")"
'        End If
'    Next i
'
'DoItForAdmin:
'    If (Image2.Picture = 0 Or Trim(PicFileName) = "") Then GoTo TheEnd
'
    con.Execute "UPDATE Users SET ExportMainLink=" & Check1(1).Value & ",FinancialMainLink=" & Check1(2).Value & _
     ",PayrollMainLink=" & Check1(4).Value & ",ProductionMainLink=" & Check1(5).Value & ",StockMainLink=" & _
     Check1(3).Value & ",CompanyMainLink=" & Check1(0).Value & ",RestrictedItemProfile=" & chkAdv(0).Value & _
     ",GeneralInfoItemProfile=" & chkAdv(1).Value & ",ProcessesItemProfile=" & chkAdv(2).Value & ",WeightItemProfile=" & _
     chkAdv(3).Value & ",PriceItemProfile=" & chkAdv(4).Value & ",ReferencesItemProfile=" & chkAdv(5).Value & _
     ",PictureItemProfile=" & chkAdv(6).Value & ",RMItemProfile=" & chkAdv(7).Value & ",ShipInfoItemProfile=" & _
     chkAdv(8).Value & ",ChangeRateonIssuance=" & chkAdv(9).Value & ",ChangeRateMakerAssign=" & chkAdv(10).Value & _
     ",AuthorizeVouchers=" & chkAdv(11).Value & ",EditCustCatalog=" & chkAdv(12).Value & ",DashBoardMainLink=" & _
     Check1(6).Value & ",PostMakerBill=" & chkAdv(13).Value & ",AuthorizeIssuance=" & chkAdv(14).Value & ",HelpMainLink=" & _
     Check1(7).Value & ",FixedAssetsMainLink=" & Check1(9).Value & ",QMSMainLink=" & Check1(8).Value & ",OpenCommandCenter=" & _
     chkAdv(16).Value & ",EditFollowup=" & chkAdv(17).Value & ",CloseCAPA=" & chkAdv(18).Value & ",ChangePassword=" & chkAdv(19).Value & ",SamplingMainLink=" & Check1(10).Value & ",RestrictCompanyCatalogEditing=" & chkAdv(20).Value & ",MakerBill_EditRate=" & chkAdv(21).Value & ",PrintDuplicate_PTC=" & chkAdv(22).Value & ",HideRateMakerAssign=" & chkAdv(23).Value & ",MIL_Print_Master_PO=" & chkAdv(24).Value & ",MIL_Print_Maker_Issuance_Report_Valuewise=" & chkAdv(25).Value & ",Show_Balance_Vendor_List=" & chkAdv(26).Value & ",Show_Values_Company_Catalog=" & chkAdv(27).Value & " WHERE UserID=" & IntUserID
     
    Dim Rec As New ADODB.Recordset
'
    Dim rsThumb As New ADODB.Stream

    Rec.Open "SELECT UserID,UserPic FROM Users WHERE UserID=" & IntUserID, con, adOpenStatic, adLockOptimistic

    'Save User Pic
    If ImgUser.Picture <> 0 And Trim(PicFileName) <> "" Then
        rsThumb.Open
        rsThumb.Type = adTypeBinary
        rsThumb.LoadFromFile PicFileName
        Rec.Fields("UserPic") = rsThumb.Read
        rsThumb.Close
    End If

    Rec.Update
    Rec.Close
    
    Rec.Open "SELECT UserID,UserSignature FROM Users WHERE UserID=" & IntUserID, con, adOpenStatic, adLockOptimistic

    'Save User Pic
    If ImgSignature.Picture <> 0 And Trim(SigFileName) <> "" Then
        rsThumb.Open
        rsThumb.Type = adTypeBinary
        rsThumb.LoadFromFile SigFileName
        Rec.Fields("UserSignature") = rsThumb.Read
        rsThumb.Close
    End If

    Rec.Update
    Rec.Close
    
    Set Rec = Nothing
    Set rsThumb = Nothing
    
    con.Execute "DELETE FROM Users_RMGroups WHERE UserID=" & IntUserID
    For i = 1 To LVRMGroups.ListItems.count
        If LVRMGroups.ListItems(i).Checked Then
            con.Execute "INSERT INTO Users_RMGroups(UserID,RMGroupID) VALUES(" & IntUserID & "," & Val(LVRMGroups.ListItems(i).key) & ")"
        End If
    Next
    
    con.Execute "DELETE FROM Users_Customers WHERE UserID=" & IntUserID
    For i = 1 To LVCustomers.ListItems.count
        If LVCustomers.ListItems(i).Checked Then
            con.Execute "INSERT INTO Users_Customers(UserID,CustCode) VALUES(" & IntUserID & ",'" & LVCustomers.ListItems(i).Tag & "')"
        End If
    Next
    
    con.Execute "DELETE FROM Users_ItemGroups WHERE UserID=" & IntUserID
    For i = 1 To LVItemGroups.ListItems.count
        If LVItemGroups.ListItems(i).Checked Then
            con.Execute "INSERT INTO Users_ItemGroups(UserID,ItemGroups_ID) VALUES(" & IntUserID & "," & Val(LVItemGroups.ListItems(i).Tag) & ")"
        End If
    Next
    
    con.Execute "DELETE FROM Users_Stores WHERE UserID=" & IntUserID
    For i = 1 To LVStores.ListItems.count
        If LVStores.ListItems(i).Checked Then
            con.Execute "INSERT INTO Users_Stores(UserID,Store_RefID) VALUES(" & IntUserID & "," & Val(LVStores.ListItems(i).Tag) & ")"
        End If
    Next
    
    'Allowed Computers
    If strDelACList <> "" Then
        strDelACList = Left(strDelACList, Len(strDelACList) - 1)
        con.Execute "DELETE FROM Users_AllowedComputers WHERE EntryID IN(" & strDelACList & ")"
    End If
    
    For i = 1 To LVAC.ListItems.count
        With LVAC.ListItems(i)
            If Val(.ListSubItems(1).Tag) = 0 Then
                con.Execute "INSERT INTO Users_AllowedComputers(UserID,MachineName) VALUES(" & IntUserID & ",'" & .SubItems(1) & "')"
            End If
        End With
    Next
    
'
'TheEnd:
'
'    con.CommitTrans
'
    Unload Me

    Exit Sub
err:
    MsgBox err.Description

End Sub

Public Sub ShowUserSettings(iUserID As Integer)
    
    
    Me.Hide
    
    IntUserID = iUserID
    
    On Error Resume Next
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        .Open "SELECT * FROM Users_RMGroups WHERE UserId=" & iUserID, con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            LVRMGroups.ListItems(Val(!RMGROUPID & "") & "'").Checked = True
            .MoveNEXT
        Loop
        .Close
        
        .Open "SELECT * FROM Users_Customers WHERE UserId=" & iUserID, con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            LVCustomers.ListItems(!CustCode & "'").Checked = True
            .MoveNEXT
        Loop
        .Close
        
        .Open "SELECT * FROM Users_ItemGroups WHERE UserId=" & iUserID, con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            LVItemGroups.ListItems(!ItemGroups_ID & "'").Checked = True
            .MoveNEXT
        Loop
        .Close
        
        .Open "SELECT * FROM Users_Stores WHERE UserId=" & iUserID, con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            LVStores.ListItems(!Store_RefID & "'").Checked = True
            .MoveNEXT
        Loop
        .Close
                
        .Open "SELECT * FROM Users_AllowedComputers WHERE UserId=" & iUserID & " ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            Set ITM = LVAC.ListItems.add(, , LVAC.ListItems.count + 1)
            ITM.ListSubItems.add(, , !MachineName & "").Tag = Val(!EntryID & "")
            .MoveNEXT
        Loop
        .Close
        
        .Open "SELECT * FROM Security", con, adOpenForwardOnly, adLockReadOnly
'        LVSecurity.ListItems.Clear
        Do Until .EOF
            Set ITM = LVSecurity.ListItems.add(, , LVSecurity.ListItems.count + 1)
            ITM.ListSubItems.add , , !SettingName & ""
            ITM.ListSubItems.add(, , "**********").Tag = !Password & ""
            .MoveNEXT
        Loop
        .Close
        
    End With
    Set rs = Nothing
    On Error GoTo 0
    
    With rsUser
        .Open "SELECT * FROM Users WHERE UserID=" & iUserID, con, adOpenDynamic, adLockOptimistic
        'On Error GoTo err
        
        For i = 0 To Check1.count - 1
            Set Check1(i).DataSource = rsUser
        Next
        For i = 0 To chkAdv.count - 1
            Set chkAdv(i).DataSource = rsUser
        Next
'        Set txtUserName.DataSource = rsUser
'
'        If Not IsNull(![Signature]) Then
'            Set Image2.DataSource = rsUser
'            Image2.DataField = "Signature"
'            Label3.Visible = False
'        Else
'            Label3.Visible = True
'        End If
        
        lblUserName.Caption = " " & ![UserName] & "" 'UserName & ""
        txtUserName = ![FullUserName] & ""
        
        If Not IsNull(![UserPic]) Then
            Set ImgUser.DataSource = rsUser
            ImgUser.DataField = "UserPic"
            lblPic.Visible = False
        Else
            lblPic.Visible = True
        End If
            
        If Not IsNull(![UserSignature]) Then
            Set ImgSignature.DataSource = rsUser
            ImgSignature.DataField = "UserSignature"
            lblSig.Visible = False
        Else
            lblSig.Visible = True
        End If
    End With
    
    Set rs = Nothing
    Me.Show 1
    
End Sub

Private Sub cmdUnCheckAll_Click()
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = False
        Call LV_ItemCheck(LV.ListItems(i))
    Next
End Sub

Private Sub Form_Load()

    On Error GoTo err
    
    Loaded = False
    
    Call AddToCombo(cmbModule, "Distinct ModuleName", "MenuOptions")
    SSTab1.Tab = 0
    
    Dim i As Integer
    Dim rs As New ADODB.Recordset, ITM As ListItem
    
    With rs
        .Open "SELECT ID,Description FROM RMGroups ORDER BY Description", con, adOpenForwardOnly, adLockReadOnly
        LVRMGroups.ListItems.Clear
        Do Until .EOF
            Set ITM = LVRMGroups.ListItems.add(, Val(!ID & "") & "'")
            ITM.ListSubItems.add , , !Description & ""
            .MoveNEXT
        Loop
        .Close
        
        .Open "SELECT CustCode FROM ForeignCustomers ORDER BY CustCode", con, adOpenForwardOnly, adLockReadOnly
        LVCustomers.ListItems.Clear
        Do Until .EOF
            Set ITM = LVCustomers.ListItems.add(, !CustCode & "'")
            ITM.Tag = !CustCode & ""
            ITM.ListSubItems.add , , !CustCode & ""
            .MoveNEXT
        Loop
        .Close
        
        .Open "SELECT * FROM ItemGroups ORDER BY ID", con, adOpenForwardOnly, adLockReadOnly
        LVItemGroups.ListItems.Clear
        Do Until .EOF
            Set ITM = LVItemGroups.ListItems.add(, !ID & "'")
            ITM.Tag = !ID & ""
            ITM.ListSubItems.add , , !Description & ""
            .MoveNEXT
        Loop
        .Close
        
        .Open "SELECT * FROM Stores ORDER BY StoreName", con, adOpenForwardOnly, adLockReadOnly
        LVStores.ListItems.Clear
        Do Until .EOF
            Set ITM = LVStores.ListItems.add(, !EntryID & "'")
            ITM.Tag = !EntryID & ""
            ITM.ListSubItems.add , , !StoreName & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub Form_Unload(cancel As Integer)
    On Error Resume Next
    If rsUser.State = 1 Then rsUser.Close
    Set rsUser = Nothing
    OurOwnForm = True
End Sub


Private Sub Image2_DblClick()
    Call AddSignature
End Sub

Private Sub AddUserSignature()
    On Error Resume Next
    Dim Picfile As String
    CD1.ShowSave
    If Not (err.Number = 32755) Then
        Picfile = CD1.FileName
        ImgSignature.Picture = LoadPicture(Picfile)
        SigFileName = Picfile
        lblSig.Visible = False
    End If
    
End Sub

Private Sub ImgSignature_DblClick()
    Call AddUserSignature
End Sub

Private Sub ImgUser_DblClick()
    AddUserPic
End Sub

Private Sub Label3_DblClick()
    Call AddSignature
End Sub

Private Sub AddUserPic()
    On Error Resume Next
    Dim Picfile As String
    CD1.ShowSave
    If Not (err.Number = 32755) Then
        Picfile = CD1.FileName
        ImgUser.Picture = LoadPicture(Picfile)
        PicFileName = Picfile
        lblPic.Visible = False
    End If
    
End Sub

Private Sub lblPic_DblClick()
    AddUserPic
End Sub

Private Sub lblSig_DblClick()
    Call AddUserSignature
End Sub

Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    If Item.Checked Then
       con.Execute "INSERT INTO UserMenuOptions VALUES(" & IntUserID & ",'" & Item.key & "')"
    Else
       con.Execute "DELETE FROM UserMenuOptions WHERE UserID=" & IntUserID & " and OptionID='" & Item.key & "'"
    End If
    
End Sub

Private Sub RefreshList()

    Dim ITM As ListItem
    Dim rs As New ADODB.Recordset

    LV.ListItems.Clear

    Set rs = con.Execute("exec UserSettings_SP " & IntUserID & ",'" & cmbModule & "'")

    With rs
    
        Do Until .EOF
            Set ITM = LV.ListItems.add(, ![OptionID] & "", ![OptionName] & "")
            ITM.Checked = CBool(Val(![UserID] & "") <> 0)
            .MoveNEXT
        Loop
    
        .Close
    
    End With
    Set rs = Nothing

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        cmdCheckAll.Visible = True
        cmdUnCheckAll.Visible = True
    Else
        cmdCheckAll.Visible = False
        cmdUnCheckAll.Visible = False
    End If
End Sub

Private Sub txtComputerName_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        Dim strComputerName As String, ITM As ListItem
        strComputerName = Trim(txtComputerName)
        If strComputerName = "" Then
            MsgBox "Invalid Computer Name.", vbInformation
            Exit Sub
        End If
        Set ITM = LVAC.ListItems.add(, , LVAC.ListItems.count + 1)
        ITM.ListSubItems.add , , strComputerName
        txtComputerName = ""
        txtComputerName.SetFocus
    End If
    
End Sub

Private Sub LVAC_KeyDown(KeyCode As Integer, Shift As Integer)
    If LVAC.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure to Remove?", vbQuestion + vbYesNo) = vbYes Then
            If Val(LVAC.SelectedItem.ListSubItems(1).Tag) > 0 Then
                strDelACList = strDelACList & Val(LVAC.SelectedItem.ListSubItems(1).Tag) & ","
            End If
            LVAC.ListItems.Remove LVAC.SelectedItem.Index
        End If
    End If
End Sub

