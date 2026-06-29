VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmMakerPOList 
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13860
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
   ScaleHeight     =   8835
   ScaleMode       =   0  'User
   ScaleWidth      =   12414.59
   WindowState     =   2  'Maximized
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
      Height          =   8925
      Left            =   60
      TabIndex        =   3
      Top             =   -90
      Width           =   13755
      Begin ComboList.Usercontrol1 cmbMaker 
         Height          =   285
         Left            =   60
         TabIndex        =   6
         Top             =   855
         Width           =   4320
         _ExtentX        =   7620
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
      Begin ComboList.Usercontrol1 cmbItemCat 
         Height          =   285
         Left            =   45
         TabIndex        =   20
         Top             =   1380
         Width           =   1740
         _ExtentX        =   3069
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
      Begin ComboList.Usercontrol1 cmbItemGroup 
         Height          =   285
         Left            =   1815
         TabIndex        =   32
         Top             =   1380
         Width           =   1950
         _ExtentX        =   3440
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
      Begin ComboList.Usercontrol1 cmbItems 
         Height          =   285
         Left            =   3780
         TabIndex        =   12
         Top             =   1380
         Width           =   9900
         _ExtentX        =   17463
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
      Begin ComboList.Usercontrol1 cmbPurchaser 
         Height          =   285
         Left            =   4740
         TabIndex        =   41
         Top             =   1935
         Width           =   8925
         _ExtentX        =   15743
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
         Left            =   45
         TabIndex        =   25
         Top             =   1935
         Width           =   4665
         _ExtentX        =   8229
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
      Begin ComboList.Usercontrol1 cmbCustomer 
         Height          =   285
         Left            =   3420
         TabIndex        =   35
         Top             =   2460
         Width           =   2175
         _ExtentX        =   3836
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
      Begin VB.CheckBox chkRepairIssuance 
         Caption         =   "Repair Issuance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10590
         TabIndex        =   38
         Top             =   2520
         Width           =   1770
      End
      Begin VB.CheckBox chkBookmarks 
         Caption         =   "Bookmarks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9060
         TabIndex        =   37
         Top             =   2520
         Width           =   1320
      End
      Begin VB.CheckBox chkRegularLotsOnly 
         Caption         =   "Regular Lots Only"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9060
         TabIndex        =   34
         Top             =   2265
         Width           =   1860
      End
      Begin VB.CheckBox chkOpenLotsOnly 
         Caption         =   "Open Lots Only"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7335
         TabIndex        =   33
         Top             =   2520
         Width           =   1680
      End
      Begin VB.TextBox txtMasterPONo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   31
         Top             =   2460
         Width           =   1245
      End
      Begin VB.CheckBox chkMasterPOOpen 
         Caption         =   "Open Master PO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5625
         TabIndex        =   29
         Top             =   2520
         Width           =   1665
      End
      Begin VB.TextBox txtOrderNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   900
         TabIndex        =   28
         Top             =   2460
         Width           =   1245
      End
      Begin VB.TextBox txtLotNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   60
         TabIndex        =   27
         Top             =   2460
         Width           =   825
      End
      Begin VB.CheckBox chkRepairLots 
         Caption         =   "Repair Lots Only"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7335
         TabIndex        =   26
         Top             =   2265
         Width           =   1680
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
         TabIndex        =   22
         Top             =   -15
         Width           =   13740
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Issuance List"
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
            Index           =   2
            Left            =   0
            TabIndex        =   23
            Top             =   135
            Width           =   13710
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Issuance List"
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
            Index           =   3
            Left            =   0
            TabIndex        =   24
            Top             =   150
            Width           =   13710
         End
      End
      Begin VB.CheckBox chkShowMasterPOOnly 
         Caption         =   "Master PO Only"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   21
         Top             =   2265
         Width           =   1665
      End
      Begin VB.TextBox txtDueDays 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   10875
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   855
         Width           =   885
      End
      Begin VB.ComboBox cmbDueDays 
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
         ItemData        =   "frmMakerPOList.frx":0000
         Left            =   9165
         List            =   "frmMakerPOList.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   855
         Width           =   1710
      End
      Begin VB.Frame Frame1 
         Height          =   480
         Left            =   4185
         TabIndex        =   13
         Top             =   8400
         Width           =   5475
         Begin VB.ComboBox cmbfield 
            Height          =   315
            ItemData        =   "frmMakerPOList.frx":0048
            Left            =   60
            List            =   "frmMakerPOList.frx":0061
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   135
            Width           =   1875
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4395
            TabIndex        =   15
            Top             =   135
            Width           =   1020
         End
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   1965
            TabIndex        =   14
            Top             =   135
            Width           =   2430
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
         ItemData        =   "frmMakerPOList.frx":00A8
         Left            =   4425
         List            =   "frmMakerPOList.frx":00BE
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   870
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   7665
         TabIndex        =   1
         Top             =   870
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
         Format          =   173735939
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   6135
         TabIndex        =   0
         Top             =   870
         Width           =   1515
         _ExtentX        =   2672
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
         Format          =   173735939
         CurrentDate     =   37055
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5625
         Left            =   60
         TabIndex        =   7
         Top             =   2775
         Width           =   13605
         _ExtentX        =   23998
         _ExtentY        =   9922
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Maker"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "PO #"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Process"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Item ID"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Value"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Status"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "(Qty) Promises"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "Lot #"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Master PO No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Cust. PO #"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   12
            Text            =   "Internal Ref #"
            Object.Width           =   2117
         EndProperty
      End
      Begin ComboList.Usercontrol1 Usercontrol11 
         Height          =   285
         Left            =   3720
         TabIndex        =   39
         Top             =   210
         Width           =   5175
         _ExtentX        =   9128
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
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   " Process                                                                                         Purchaser"
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
         Index           =   6
         Left            =   60
         TabIndex        =   42
         Top             =   1710
         Width           =   11640
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMakerPOList.frx":010E
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
         Index           =   5
         Left            =   0
         TabIndex        =   40
         Top             =   0
         Width           =   11640
      End
      Begin MSForms.CommandButton cmdPurchaseCalendar 
         Height          =   375
         Left            =   10305
         TabIndex        =   36
         Top             =   8460
         Width           =   1290
         ForeColor       =   0
         Caption         =   "Calendar"
         PicturePosition =   327683
         Size            =   "2275;661"
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
         Caption         =   " Lot #               Order #        Master PO #              Customer "
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
         Index           =   4
         Left            =   60
         TabIndex        =   30
         Top             =   2265
         Width           =   5505
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMakerPOList.frx":01CB
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
         TabIndex        =   17
         Top             =   1170
         Width           =   13620
      End
      Begin MSForms.CommandButton cmdAuth 
         Height          =   375
         Left            =   11655
         TabIndex        =   11
         Top             =   8460
         Width           =   1035
         ForeColor       =   0
         Caption         =   "Authorize"
         PicturePosition =   327683
         Size            =   "1826;661"
         Accelerator     =   65
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdCheckAll 
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   10
         Top             =   8445
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Un-Check All"
         PicturePosition =   327683
         Size            =   "2355;661"
         Accelerator     =   85
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdCheckAll 
         Height          =   375
         Index           =   0
         Left            =   60
         TabIndex        =   9
         Top             =   8445
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Check All"
         PicturePosition =   327683
         Size            =   "2355;661"
         Accelerator     =   104
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
         Left            =   12750
         TabIndex        =   8
         Top             =   8460
         Width           =   915
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "1614;661"
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
         Height          =   405
         Left            =   12780
         TabIndex        =   2
         Top             =   2340
         Width           =   885
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "1561;714"
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
         Caption         =   $"frmMakerPOList.frx":0281
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
         Top             =   645
         Width           =   11700
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
            Picture         =   "frmMakerPOList.frx":035E
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakerPOList.frx":051A
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PopMenu"
      Begin VB.Menu mnuPrintPurchaseOrder 
         Caption         =   "Print IPO/CTC"
      End
      Begin VB.Menu mnuPrintInternalMasterPO 
         Caption         =   "Print Internal Master PO"
      End
      Begin VB.Menu mnuPrintIPOCTCMin 
         Caption         =   "Print Gate Pass"
      End
      Begin VB.Menu mnuPrintTag 
         Caption         =   "Print Tag"
      End
      Begin VB.Menu mnuPrintPTC 
         Caption         =   "Print PTC"
      End
      Begin VB.Menu mnuPrintPTCMini 
         Caption         =   "Print PTC Mini"
      End
      Begin VB.Menu mnuPrintPTCWithPrice 
         Caption         =   "Print PTC (Cost)"
      End
      Begin VB.Menu mnuPrintGatePass 
         Caption         =   "Print Gate-Pass"
      End
      Begin VB.Menu mnuPrintList 
         Caption         =   "Print Master PO"
      End
      Begin VB.Menu mnuPrintMasterPOAccountsCopy 
         Caption         =   "Print Master PO (Accounts Copy)"
      End
      Begin VB.Menu mnuPrintMasterPOMakerCopy 
         Caption         =   "Print Master PO (Maker Copy)"
      End
      Begin VB.Menu mnuPrintItemPicture 
         Caption         =   "Print Item Picture"
      End
      Begin VB.Menu mnuPrintDrawings 
         Caption         =   "Print Item Drawings"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintMasterPOHideRate 
         Caption         =   "Print Master PO (Hide Rate)"
      End
      Begin VB.Menu mnuPrintMasterPOStatus 
         Caption         =   "Print Master PO Status"
      End
      Begin VB.Menu mnuDashOutlookReport 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintR2OutlookReport 
         Caption         =   "R2-Outlook Report"
      End
      Begin VB.Menu mnuDashDelete 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Selected Order(s)"
      End
      Begin VB.Menu mnuDashClosePO 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCloseMakerPO 
         Caption         =   "Close PO"
      End
      Begin VB.Menu mnuEditDatesDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPromises 
         Caption         =   "Edit Return Dates"
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddBookmarks 
         Caption         =   "Add Bookmark for Receiving"
      End
      Begin VB.Menu mnuAddFollowupEntries 
         Caption         =   "Add Followup Entries"
      End
      Begin VB.Menu mnuPrintFollowupReport 
         Caption         =   "Print Followup Report"
      End
      Begin VB.Menu mnuDashBookmark 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintThisList 
         Caption         =   "Print This List"
      End
      Begin VB.Menu mnuPrintListBatchWise 
         Caption         =   "Print List (Batchwise)"
      End
      Begin VB.Menu mnuPrintMakerList 
         Caption         =   "Print Maker List"
      End
      Begin VB.Menu mnuDashSFStock1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintSFStock 
         Caption         =   "Print S/F Stock"
      End
      Begin VB.Menu mnuDashPrintBalance 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMakerBalanceReport 
         Caption         =   "Print Maker Balance Report"
      End
      Begin VB.Menu mnuPrintMakerBalanceReportSummary 
         Caption         =   "Print Maker Balance Report (Summary)"
      End
      Begin VB.Menu mnuPrintMakerIssuanceReport 
         Caption         =   "Maker Issuance Report (Valuewise)"
      End
      Begin VB.Menu mnuPrintPurchasePlan 
         Caption         =   "Print Purchase Plan"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintOrdersOfPurchaser 
         Caption         =   "Print Orders of "
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintVendorPurchaseStatusLog 
         Caption         =   "Vendor Purchase Status Log"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintVendorPurchaseStatusLogBalanceOnly 
         Caption         =   "Vendor Purchase Status Log (Balance)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintVendorPurchaseStatusLogNillOnly 
         Caption         =   "Vendor Purchase Status Log (Nil)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDashNew 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDeliveryPerformanceReportPuchaserwise 
         Caption         =   "Report No:. PLN/009/00"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDeliveryPerformanceReportSummary 
         Caption         =   "Report No:. PLN/008/00"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDeliveryPerformanceReportPuchaserwiseAnnualSummary 
         Caption         =   "Report No:. PLN/007/00"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDeliveryPerformanceReportPuchaserwiseAnnualSummaryMaker 
         Caption         =   "Report No:. PLN/006/00"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPurchaserwisePurchaseplan 
         Caption         =   "Report No:. PLN/002/00"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMakerPOList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCourierInvoice As Boolean
Dim bShipping As Boolean
Dim bCustomInvoice As Boolean
Dim strReportSQL As String

Public Sub ShowMe(Optional p_bCustomInvoice As Boolean = True)
    bCustomInvoice = p_bCustomInvoice
    Me.Show
End Sub

Public Sub ShowForShipping()
    bShipping = True
    Me.Show
End Sub

Private Sub chkMasterPOOpen_Click()
    If chkMasterPOOpen.Value = vbChecked Then
        chkShowMasterPOOnly.Value = vbUnchecked
        chkRepairLots.Value = vbUnchecked
    End If
End Sub

Private Sub chkRegularLotsOnly_Click()
    If chkRegularLotsOnly.Value = vbChecked Then
        chkRepairLots.Value = vbUnchecked
    End If
End Sub

Private Sub chkRepairLots_Click()

    If chkRepairLots.Value = vbChecked Then
        chkShowMasterPOOnly.Value = vbUnchecked
        chkMasterPOOpen.Value = vbUnchecked
        chkRegularLotsOnly.Value = vbUnchecked
    End If
    
End Sub

Private Sub chkShowMasterPOOnly_Click()

    If chkShowMasterPOOnly.Value = vbChecked Then
        chkMasterPOOpen.Value = vbUnchecked
        chkRepairLots.Value = vbUnchecked
    End If
    
End Sub

Private Sub cmbDueDays_Change()
    Call cmbDueDays_Click
End Sub

Private Sub cmbDueDays_Click()

    txtDueDays.Locked = True
    If cmbDueDays.ListIndex = 0 Then
        txtDueDays.Text = ""
    ElseIf cmbDueDays.ListIndex = 1 Then
        txtDueDays.Text = "3"
    ElseIf cmbDueDays.ListIndex = 2 Then
        txtDueDays.Text = "7"
    Else
        txtDueDays.Text = ""
        txtDueDays.Locked = False
    End If
    
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
            If DtFrom.Visible Then
                DtFrom.SetFocus
            End If
            Exit Sub
    End Select
    
End Sub

Private Sub cmdAuth_Click()

    Dim i As Integer
    If GetSingleBooleanValue("AuthorizeIssuance", "Users", " WHERE UserName='" & CurrentUserName & "'") = False Then
        MsgBox "Access Denied.", vbInformation
        Exit Sub
    End If
    
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked And CBool(.ListSubItems(7).Tag) = False Then
                con.Execute "UPDATE VendIssued SET Authorized=1,AuthUserName='" & CurrentUserName & "',AuthMachineName='" & strComputerName & "',AuthEntryDT=getDate() WHERE EntryID=" & .Tag
                .ListSubItems(6).Text = "Authorized"
                .ListSubItems(6).Tag = True
            End If
        End With
    Next
    Call RefreshLV
    
End Sub

Private Sub cmdCheckAll_Click(Index As Integer)
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = (Index = 0)
    Next
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
       
       
       If cmbfield.ListIndex = 0 Then
          fldValue = Trim(Replace(LV.ListItems(i), "»", ""))
       Else
          fldValue = Trim(Replace(LV.ListItems(i).SubItems(cmbfield.ListIndex), "»", ""))
       End If
    
       If LCase(fldValue) Like "*" & LCase(Trim(txtFind)) & "*" Then
          LV.ListItems(i).Selected = True
          LV.ListItems(i).EnsureVisible
          If i <> StartIdx Then Exit For
       End If
    Next
    
    If i = LV.ListItems.count + 1 Then
        MsgBox "No Mathes Found.", vbInformation
        Set LV.SelectedItem = Nothing
    End If

End Sub

Private Sub cmdPurchaseCalendar_Click()

    Dim rpt As CRAXDDRT.Report, strSelectionMain As String
    Set rpt = rptApp.OpenReport(RptPath & "\PurchaseCalendar.rpt")
    Dim f As New frmPrevRpt
    strSelectionMain = "{DateLookup.DateFull}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#"
    
    rpt.FormulaFields.GetItemByName("Maker").Text = cmbMaker.ID
    rpt.FormulaFields.GetItemByName("ItemGroup").Text = cmbItemGroup.ID
    rpt.FormulaFields.GetItemByName("Customer").Text = "'" & cmbCustomer.ID & "'"
    
    f.ShowReport strSelectionMain, rpt
    
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
        
    If strCompany = "IAA" Or strCompany = "Kami" Then
        mnuPrintMasterPOHideRate.Visible = True
        mnuPrintItemPicture.Visible = True
        cmbMaker.EnableFunctionality = True
    Else
        mnuPrintMasterPOHideRate.Visible = False
        mnuPrintItemPicture.Visible = False
        cmbMaker.EnableFunctionality = False
    End If
    
    If strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
        mnuPrintR2OutlookReport.Visible = True
        mnuDashOutlookReport.Visible = True
        mnuPrintList.Caption = "Print Master PO (Office Copy)"
        
    Else
        mnuPrintR2OutlookReport.Visible = False
        mnuDashOutlookReport.Visible = False
        mnuPrintMasterPOAccountsCopy.Visible = False
        mnuPrintMasterPOMakerCopy.Visible = False
    End If
    
    
    mnuPop.Visible = False
    
    cmbItems.ListHeight = 6000
    
    cmbMaker.ListHeight = 3800
    cmbMaker.AddItem "<All Makers>", "0"
    Call cmbMaker.AddVals(con, "'{ ' + VendID1 + ' } ' + VenderName ", "Makers", "VendID")
    
    cmbCustomer.ListHeight = 4000
    cmbCustomer.EnableFunctionality = True
    cmbCustomer.AddItem "<All Customers>", "0"
    
    Call cmbCustomer.AddVals(con, "CustCode", "ForeignCustomers", "CustCode")
    
    cmbMaker.ID = "0"
    cmbCustomer.ID = "0"
    DtFrom = Date
    DtTo = Date
    cmbRange.ListIndex = 0
     
    cmbItemCat.ListHeight = 4000
    cmbItemCat.EnableFunctionality = True
    cmbItemCat.AddItem "<All Catagories>", "0"
    cmbItemCat.AddVals con, "Description", "ItemCatagories", "CatID"
    cmbItemCat.ID = "0"
    
    cmbItemGroup.ListHeight = 4000
    cmbItemGroup.EnableFunctionality = True
    cmbItemGroup.AddItem "<All Groups>", "0"
    cmbItemGroup.AddVals con, "Description", "ItemGroups", "ID"
    cmbItemGroup.ID = "0"
    
    cmbItems.DropDownOffLine = True
    cmbItems.AddItem "<All Articles>", "0"
    Call cmbItems.AddVals(con, "ItemID + ' ' + ItemName + ' { ' + ISNULL(Attributes,'') + ' }'", "Items", "ItemID", " WHERE InActive=0 ORDER BY ItemName")
    cmbItems.ID = "0"
     
    cmbProcess.ListHeight = 4000
    cmbProcess.EnableFunctionality = True
    
    cmbProcess.AddItem "<All Processes>", "0"
    cmbProcess.AddVals con, "Description", "Processes", "ProcessID", " ORDER BY Description"
    cmbProcess.ID = "0"
    
    cmbPurchaser.AddItem "<All Purchasers>", "0"
    cmbPurchaser.AddVals con, "'{'+EmpID+'} '+Name", "Employees", "EmpID", " ORDER BY EmpID"
    cmbPurchaser.ID = "0"
    
    With cmbfield
        .Clear
        For i = 1 To LV.ColumnHeaders.count
            cmbfield.AddItem LV.ColumnHeaders(i).Text
        Next
    End With
        
    cmbDueDays.ListIndex = 0
    
End Sub

Private Sub RefreshLV()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim Cond As String
    Dim DTRange As String
    Dim ITM As ListItem
    
    Screen.MousePointer = vbHourglass
    
    If cmbDueDays.ListIndex > 0 Then
    
        Dim myTargetDT As Date
        myTargetDT = DateAdd("d", Val(txtDueDays), Date)
        Dim myDTFrom As Date, myDTTo As Date
        If myTargetDT < Date Then
            myDTFrom = myTargetDT
            myDTTo = Date
        Else
            myDTFrom = Date
            myDTTo = myTargetDT
        End If
        Cond = " WHERE (VVendIssued.EntryID IN(SELECT RefID FROM VendIssdDetail INNER JOIN VendIssdDetail_ReturnDTs ON " & _
         "VendIssdDetail.EntryID=VendIssdDetail_ReturnDTs.VIS_RefID WHERE VendIssdDetail_ReturnDTs.ReturnDT BETWEEN '" & myDTFrom & "' AND '" & myDTTo & "' AND VendIssdDetail_ReturnDTs.IssQty>VendIssdDetail_ReturnDTs.RcvQty))"
         
    Else
        If cmbRange.ListIndex = 0 Then
            DTRange = " OnlyDT Between '" & DateAdd("d", 0, GetServerDate) & "' and '" & GetServerDate & "'"
        ElseIf cmbRange.ListIndex = 1 Then
            DTRange = " OnlyDT Between '" & DateAdd("d", -15, GetServerDate) & "' and '" & GetServerDate & "'"
        ElseIf cmbRange.ListIndex = 2 Then
            DTRange = " OnlyDT Between '" & DateAdd("d", -30, GetServerDate) & "' and '" & GetServerDate & "'"
        ElseIf cmbRange.ListIndex = 3 Then
            DTRange = " OnlyDT Between '" & DateAdd("d", -60, GetServerDate) & "' and '" & GetServerDate & "'"
        ElseIf cmbRange.ListIndex = 4 Then
            DTRange = " OnlyDT Between '" & DateAdd("d", -90, GetServerDate) & "' and '" & GetServerDate & "'"
        ElseIf cmbRange.ListIndex = 5 Then
            DTRange = " OnlyDT Between '" & DtFrom & "' and '" & DtTo & "'"
        End If
        
        If Trim(txtLotNo) <> "" Or Trim(txtOrderNo) <> "" Or Trim(txtMasterPONo) <> "" Then
            Cond = " WHERE OnlyDT BETWEEN '1/1/1' AND '1/1/2090'"
        Else
            Cond = " WHERE " & DTRange
        End If
        
    End If
    
    If cmbMaker.ID <> "0" Then
        Cond = Cond & " AND VendID IN(" & cmbMaker.ID & ") "
    End If
    
    If cmbItems.ID <> "0" Then
        Cond = Cond & " AND VVendIssued.ItemID='" & cmbItems.ID & "' "
    End If
    
    If cmbItemCat.ID <> "0" Then
        If cmbItemCat.Functionality = 1 Then
            Cond = Cond & " AND CatID IN(" & cmbItemCat.ID & ")"
        Else
            Cond = Cond & " AND CatID='" & cmbItemCat.ID & "'"
        End If
    End If
    
    If cmbItemGroup.ID <> "0" Then
        Cond = Cond & " AND GroupID IN(" & cmbItemGroup.ID & ")"
    End If
    
    If cmbProcess.ID <> "0" Then
        Cond = Cond & " AND ProcessID IN(" & cmbProcess.ID & ")"
    End If
      
    If Trim(txtLotNo) <> "" Then
        Cond = " WHERE LotNo='" & txtLotNo & "' "
    End If
    
    If Trim(txtOrderNo) <> "" Then
        Cond = Cond & " AND VVendIssued.EntryID IN(SELECT RefID FROM VendIssdDetail WHERE OrderNo='" & txtOrderNo & "')"
    End If
    
    If Trim(txtMasterPONo) <> "" Then
        Cond = Cond & " AND MasterPONo='" & txtMasterPONo & "'"
    End If
    
    If chkShowMasterPOOnly.Value = vbChecked Then
        Cond = Cond & " AND VVendIssued.MasterPONo IS NOT NULL"
    End If
    
    If chkMasterPOOpen.Value = vbChecked Then
'        Cond = Cond & " AND TotalIssQty+ISNULL(WastageQty,0)>ISNULL(RcvdQty,0) AND MasterPONo IS NOT NULL"
'        Cond = Cond & " AND Closed=0"
        Cond = Cond & " AND DetailEntryID IN(SELECT VID_EntryID FROM VItemOpenPOs)"
    End If
    
    If chkRepairLots.Value = vbChecked Then
        Cond = Cond & " AND ReWorkLot=1"
    End If
    
    If chkRegularLotsOnly.Value = vbChecked Then
        Cond = Cond & " AND ReWorkLot=0"
    End If
    
    If chkBookmarks.Value = vbChecked Then
        Cond = Cond & " AND BookMarkEntryID IS NOT NULL"
    End If
    
    If chkOpenLotsOnly.Value = vbChecked Then
        Cond = Cond & " AND (LotNo<>'0' AND VVendIssued.EntryID NOT IN(SELECT Issuance_RefID FROM VendReceived) AND LotNo IN(SELECT LotNo FROM VRunningLots_Simple) AND LotNo NOT IN(SELECT LotNo FROM Lots_Closed))"
    End If
    
    If cmbCustomer.ID <> "0" Then
        If cmbCustomer.Functionality = 1 Then
            Cond = Cond & " AND CustCode IN(" & cmbCustomer.ID & ")"
        Else
            Cond = Cond & " AND CustCode='" & cmbCustomer.ID & "'"
        End If
    End If
    
    If chkRepairIssuance.Value = vbChecked Then
        Cond = Cond & " AND VVendIssued.DetailEntryID IN(SELECT VID_RefID FROM VendReWorkIssuanceDetail)"
    End If
    
    If cmbPurchaser.ID <> "0" Then
        Cond = Cond & " AND VVendIssued.EntryID IN(SELECT RefID FROM VendIssdDetail WHERE VID_EmpID='" & cmbPurchaser.ID & "')"
    End If
    Dim lForeColor As Long
    
    With rs
        strReportSQL = "SELECT VVendIssued.* FROM VVendIssued LEFT OUTER JOIN VItemOpenPOs ON VVendIssued.DetailEntryID=VItemOpenPOs.VID_EntryID " & Cond & " ORDER BY DT"
        .Open "SELECT VVendIssued.*,VItems_Complaints.ItemID AS ComplaintItemID  FROM VVendIssued  LEFT OUTER JOIN VItems_Complaints ON VVendIssued.ItemID=VItems_Complaints.ItemID" & Cond & " ORDER BY DT", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        'LV.Sorted = False
        Do Until .EOF
            If IsNull(!ReWorklot) Then
                lForeColor = LV.ForeColor
            ElseIf !ReWorklot Then
                lForeColor = lForeColor_RepairLot
            ElseIf Val(!BookmarkEntryID & "") > 0 Then
                lForeColor = RGB(255, 0, 255)
            ElseIf !Closed Then
                lForeColor = RGB(59, 59, 59)
            ElseIf !ComplaintItemID & "" <> "" Then
                lForeColor = RGB(200, 200, 0)
            Else
                lForeColor = LV.ForeColor
            End If
            
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            ITM.Tag = Val(![EntryID] & "")
            ITM.ListSubItems.add(, , "{ " & !VendID1 & " } " & !VenderName).Tag = "" '![ImpMaterial] & ""
            ITM.ListSubItems.add(, , ![RecieptID] & "").Tag = Val(!DetailEntryID & "")
            ITM.ListSubItems.add(, , Format(![DT] & "", "dd-MMM-yyyy")).Tag = ""
          
            ITM.ListSubItems.add , , !Description & ""
            ITM.ListSubItems.add , , ![ITemID] & " " & !ItemName & " " & !TipSize & " " & !ItemSize & " " & !SizeUnit & ""
            
            ITM.ListSubItems.add , , Val(!TotalValue & "")
            ITM.ListSubItems.add(, , IIf(!Authorized, "Authorized", "Un-Auth.")).Tag = !Authorized
            
            ITM.ListSubItems.add(, , !MakerPOReturnDTs & "").ForeColor = vbRed
            
            ITM.ListSubItems.add(, , ![LotNo] & "").Tag = !MasterPONo & ""
            ITM.ListSubItems.add , , !MasterPONo & ""
            ITM.ListSubItems.add , , !OrderNo & ""
            ITM.ListSubItems.add , , !InternalRefNo & ""
            
            ITM.ForeColor = lForeColor
            For i = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(i).ForeColor = lForeColor
            Next
            
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing

    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    FAcc.Top = (Me.ScaleHeight - FAcc.Height) / 2
    
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If LV.SortKey + 1 = ColumnHeader.Index Then
        If LV.SortOrder = lvwAscending Then
            LV.SortOrder = lvwDescending
            'ColumnHeader.Icon = "Down"
        Else
            LV.SortOrder = lvwAscending
            'ColumnHeader.Icon = "Up"
        End If
    Else
        LV.ColumnHeaderIcons = Nothing
        Set LV.ColumnHeaderIcons = ImageList2
        LV.SortKey = ColumnHeader.Index - 1
        LV.SortOrder = lvwAscending
        ColumnHeader.Icon = "Down"
    End If
    Call RefreshSnos
    'If LV.Sorted = False Then LV.Sorted = True
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 2 Then Exit Sub
    
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
    Me.PopUpMenu mnuPop
    
End Sub

Private Sub ShowMenus(vis As Boolean)
            
    If LV.ListItems.count = 0 Then
        mnuEditDatesDash.Visible = False
        mnuEditPromises.Visible = False
    Else
        If LV.SelectedItem.ListSubItems(9).Tag <> "" Then
            mnuEditDatesDash.Visible = True
            mnuEditPromises.Visible = True
        Else
            mnuEditDatesDash.Visible = False
            mnuEditPromises.Visible = False
        End If
    End If
    
    If strCompany = "QEL" Then
        mnuPrintGatePass.Visible = True
    Else
        mnuPrintGatePass.Visible = False
    End If
    
    If strCompany <> "IAA" Then
        mnuAddFollowupEntries.Visible = False
    Else
        If LV.ListItems.count > 0 Then
            If LV.SelectedItem.SubItems(10) <> "" Then
                mnuAddFollowupEntries.Visible = True
            Else
                mnuAddFollowupEntries.Visible = False
            End If
        End If
    End If
    
    If cmbPurchaser.ID = "0" Then
        mnuPrintOrdersOfPurchaser.Visible = False
    Else
        mnuPrintOrdersOfPurchaser.Caption = "Print Orders of " & cmbPurchaser.Text
        mnuPrintOrdersOfPurchaser.Visible = True
    End If
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
    rpt.FormulaFields(3).Text = "'" & cmbMaker.Text & "'"
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

Private Sub mnuAddBookmarks_Click()
     
    Dim f As New frmAddBookmarkforReceiving
    Load f
    f.AddBookmark (Val(LV.SelectedItem.ListSubItems(2).Tag))
    
End Sub

Private Sub mnuAddFollowupEntries_Click()

    Dim lReturnDTsEntryID As Long
    lReturnDTsEntryID = GetSingleLongValue("MAX(EntryID)", "VendIssdDetail_ReturnDTs", " WHERE VIS_RefID=" & Val(LV.SelectedItem.ListSubItems(2).Tag))
    Dim f As New frmAddFollowupData
    Load f
    f.AddFollowup lReturnDTsEntryID, "(F1)"
    
End Sub

Private Sub mnuCloseMakerPO_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    If Val(LV.SelectedItem.SubItems(9)) > 0 Then Exit Sub
    
    If getDBPassword("CloseMakerPO") Then
        con.Execute "UPDATE VendIssued SET Closed=1 WHERE EntryID=" & Val(LV.SelectedItem.Tag)
    End If
    
End Sub

Private Sub mnuDelete_Click()

    On Error GoTo err
    If LV.ListItems.count = 0 Then Exit Sub
    Dim i As Integer, strDelList As String, iCount As Integer, iSkipCount As Integer
    strDelList = ""
    iCount = 0
    iSkipCount = 0
    
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                strDelList = strDelList & Val(.Tag) & ","
                iCount = iCount + 1
            End If
        End With
    Next
    If strDelList = "" Then
        MsgBox "No Order Selected.", vbInformation
        Exit Sub
    Else
        strDelList = Left(strDelList, Len(strDelList) - 1)
    End If
    
    If MsgBox("Are you sure to Delete " & iCount & " Orders ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    If getDBPassword("DeleteProdIss") = False Then Exit Sub
     
    Dim lCount_Pay As Long
    
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            
            If .Checked Then
                If .SubItems(10) <> "" Then 'If Master PO
                    lCount_Pay = GetSingleLongValue("COUNT(*)", "MakerAdvancesShort", " WHERE MasterPONo='" & .SubItems(10) & "'")
                    If lCount_Pay > 0 Then
                        MsgBox "Can't Delete, Short Term Loan is issued.", vbInformation
                        Exit Sub
                    End If
                    lCount_Pay = GetSingleLongValue("COUNT(*)", "MakerAdvances", " WHERE MasterPONo='" & .SubItems(10) & "'")
                    If lCount_Pay > 0 Then
                        MsgBox "Can't Delete, Long Term Loan is issued.", vbInformation
                        Exit Sub
                    End If
                End If
                If .SubItems(7) = "Un-Auth." Then
                    Call DeleteIssuance(Val(.Tag))
                    con.Execute "DELETE FROM VendIssued WHERE RecieptID='" & .SubItems(2) & "'"
                Else
                    'Check Rcving Against this Order.
                    Dim lCount As Long
                    lCount = GetSingleLongValue("COUNT(Issuance_RefID)", "VendReceived", "WHERE Issuance_RefID=" & Val(.Tag))
                    If lCount > 0 Then
                        iSkipCount = iSkipCount + 1
                    Else
                        Call DeleteIssuance(Val(.Tag))
                    End If
                End If
            End If
        End With
    Next
    
    MsgBox "Total:" & iCount & " Orders were selected, out of which " & (iCount - iSkipCount) & " were deleted, " & iSkipCount & " were skipped.", vbInformation
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub DeleteIssuance(lEntryID As Long)

    Dim lRcvd_RefID As Long, lDetailEntryID As Long
    lDetailEntryID = GetSingleLongValue("EntryID", "VendIssdDetail", " WHERE RefID=" & lEntryID)
    lRcvd_RefID = GetSingleLongValue("Rcvd_RefID", "VendIssdDetail", " WHERE RefID=" & lEntryID)
        
    Dim bRepairLot As Boolean
    bRepairLot = GetSingleBooleanValue("ReWorkLot", "VendIssdDetail", " WHERE RefID=" & lEntryID)
    
    Call StartTrans(con)
    
    If bRepairLot And lRcvd_RefID = 0 Then
        'It's Repair Lot & It's not Issued after Receiving, It was issued from Re-Work Issuance....
        con.Execute "UPDATE VendRcvdDetailReWorkDetail SET VendRcvdDetailReWorkDetail.IssQty=VendRcvdDetailReWorkDetail.IssQty-VendReWorkIssuanceDetail.Qty FROM VendRcvdDetailReWorkDetail INNER JOIN " & _
         "VendReWorkIssuanceDetail ON VendRcvdDetailReWorkDetail.EntryID=VRDRWD_RefID INNER JOIN VendIssdDetail ON " & _
         "VendReWorkIssuanceDetail.VID_RefID=VendIssdDetail.EntryID WHERE VendIssdDetail.EntryID=" & lDetailEntryID
        
        con.Execute "DELETE FROM VendReWorkIssuanceDetail WHERE VID_RefID=" & lDetailEntryID
        
    End If
    
    'Update Entries that were SF Issuance...
    
    con.Execute "UPDATE VendRcvdDetail SET VendRcvdDetail.IssQty=VendRcvdDetail.IssQty-ISNULL(T1.QtyIssd,0) FROM VendRcvdDetail " & _
     "INNER JOIN (SELECT SOO_RefID,SUM(QtyIssd) AS QtyIssd FROM StockOrderOpening_Issuance WHERE VID_RefID=" & lDetailEntryID & " GROUP BY SOO_RefID)  T1 " & _
     "ON VendRcvdDetail.Opening_RefID=T1.SOO_RefID"
    
    con.Execute "UPDATE VendRcvdDetail SET VendRcvdDetail.IssQty=VendRcvdDetail.IssQty-VendIssdDetail.IssQty FROM " & _
     "VendRcvdDetail INNER JOIN VendIssdDetail ON VendRcvdDetail.EntryID=VendIssdDetail.Rcvd_RefID " & _
     "WHERE VendIssdDetail.RefID=" & lEntryID
    'con.Execute "UPDATE VendRcvdDetail SET IssQty=IssQty-" & val( WHERE"
    con.Execute "DELETE FROM VendIssued WHERE EntryID=" & lEntryID
    'con.Execute "DELETE FROM VendIssued WHERE RecieptID='" & LV.SelectedItem.SubItems(2) & "'"
                
    con.CommitTrans
    
End Sub

Private Sub mnuDeliveryPerformanceReportPuchaserwise_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Delivery_Performance_Report_Purchaserwise.rpt")
    
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DtFrom.Value
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DtTo.Value
    rpt.ParameterFields.GetItemByName("@Purchaser").SetCurrentValue IIf(cmbPurchaser.ID = "0", "", cmbPurchaser.ID)
    rpt.ParameterFields.GetItemByName("@VendID").SetCurrentValue Val(cmbMaker.ID)
    
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, , False
    
End Sub

Private Sub mnuDeliveryPerformanceReportPuchaserwiseAnnualSummary_Click()
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Delivery_Performance_Report_Purchaserwise_Monthwise.rpt")
    
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DtFrom.Value
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DtTo.Value
    rpt.ParameterFields.GetItemByName("@Purchaser").SetCurrentValue IIf(cmbPurchaser.ID = "0", "", cmbPurchaser.ID)
    rpt.ParameterFields.GetItemByName("@VendID").SetCurrentValue Val(cmbMaker.ID)
    
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, , False
    
End Sub

Private Sub mnuDeliveryPerformanceReportPuchaserwiseAnnualSummaryMaker_Click()
        
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Delivery_Performance_Report_Makerwise_Monthwise.rpt")
    
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DtFrom.Value
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DtTo.Value
    rpt.ParameterFields.GetItemByName("@Purchaser").SetCurrentValue IIf(cmbPurchaser.ID = "0", "", cmbPurchaser.ID)
    rpt.ParameterFields.GetItemByName("@VendID").SetCurrentValue Val(cmbMaker.ID)
    
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, , False

End Sub

Private Sub mnuDeliveryPerformanceReportSummary_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Delivery_Performance_Report_Purchaserwise_Summary.rpt")
    
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DtFrom.Value
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DtTo.Value
    rpt.ParameterFields.GetItemByName("@Purchaser").SetCurrentValue IIf(cmbPurchaser.ID = "0", "", cmbPurchaser.ID)
    rpt.ParameterFields.GetItemByName("@VendID").SetCurrentValue Val(cmbMaker.ID)
    
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, , False
    
End Sub

Private Sub mnuEditPromises_Click()
    Dim f As New frmEditMakerPOReturnDates
    Load f
    f.ShowMe Val(LV.SelectedItem.Tag)
End Sub

Private Sub mnuMakerBalanceReport_Click()

    Dim rpt As CRAXDDRT.Report, strSelection As String
    Set rpt = rptApp.OpenReport(RptPath & "\MakerBalanceReport.rpt")
    
    strSelection = " (NOT ({Makers.VendID} IN[" & strFactoryMakers & "]))"
    
    strSelection = strSelection & " AND {VendIssdDetail.RcvdQty}<{VendIssdDetail.IssQty} AND {VendIssued.Closed}=FALSE"
    
    If cmbProcess.ID <> "0" Then
        strSelection = strSelection & " AND {VendIssued.ProcessID} IN[" & Replace(cmbProcess.ID, "'", "") & "]"
    End If
    
    If cmbMaker.ID <> "0" Then
        strSelection = strSelection & " AND {VendIssued.VendID}=" & cmbMaker.ID
    End If
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub mnuPrintFollowupReport_Click()

    Dim rpt As CRAXDDRT.Report
    
    Set rpt = rptApp.OpenReport(RptPath & "\Followup_Report.rpt")
    rpt.FormulaFields.GetItemByName("DateRange").Text = "'" & Format(DtFrom, "dd-MMM-yyyy") & " to " & Format(DtTo, "dd-MMM-yyyy") & "'"
    
    Dim f As New frmPrevRpt
    
    Dim i As Long, strEntries
    For i = 1 To LV.ListItems.count
        strEntries = strEntries & Val(LV.ListItems(i).Tag) & ","
    Next
    
    If strEntries <> "" Then
        strEntries = Left(strEntries, Len(strEntries) - 1)
    End If
    f.ShowReport "{VendIssued.EntryID} IN[" & strEntries & "]", rpt
    'f.ShowReport "{FollowUpData.DT}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#", rpt
    
End Sub

Private Sub mnuPrintGatePass_Click()
    Call PrintIssSlip(True)
End Sub

Private Sub mnuPrintInternalMasterPO_Click()
    Call PrintIssSlip(False, False, 3)
End Sub

Private Sub mnuPrintIPOCTCMin_Click()
    Call PrintIssSlip(False, False, 1)
End Sub

Private Sub mnuPrintItemPicture_Click()
    Call PrintMasterPOPictures(False)
End Sub

Private Sub PrintMasterPOPictures(Optional bPrintArtwork As Boolean = False)

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\IssListArticlePic.rpt")
    
    Dim f As New frmPrevRpt
    
    Dim strSelection As String

    strSelection = "{VendIssued.MasterPONo}='" & LV.SelectedItem.ListSubItems(9).Tag & "'"
    
    f.ShowReport strSelection, rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintList_Click()

    If GetUserRightAccess("MIL_Print_Master_PO") = False Then Exit Sub
    Call PrintMasterPO(False, "OFFICE COPY")
    
End Sub

Private Sub PrintMasterPO(Optional bHideRate As Boolean, Optional strCopy As String)

    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\IssList.rpt")
    
    If (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Towne" Then
        rpt.FormulaFields.GetItemByName("Copy").Text = "'" & strCopy & "'"
    End If
    
    If strCompany = "IAA" Or strCompany = "Kami" Then
        rpt.FormulaFields.GetItemByName("HideRate").Text = bHideRate
    End If
'    Dim strList As String
'    Dim i As Integer
'    For i = 1 To LV.ListItems.count
'        With LV.ListItems(i)
'            If .Checked Then
'                strList = strList & .Tag & ","
'            End If
'        End With
'    Next
'    strList = Left(strList, Len(strList) - 1)
    Dim f As New frmPrevRpt
    
    Dim strSelection As String
'    strSelection = "{VendIssued.EntryID} IN[" & strList & "]"
    strSelection = "{VendIssued.MasterPONo}='" & LV.SelectedItem.ListSubItems(9).Tag & "'"
'    Dim rs As New ADODB.Recordset
'    Dim strOrders As String
'    strOrders = ""
'    With rs
'        .Open "SELECT DISTINCT OrderNo FROM VendIssdDetail WHERE RefID IN(" & strList & ")", con, adOpenForwardOnly, adLockReadOnly
'        Do Until .EOF
'            strOrders = strOrders & !OrderNo & ","
'            .MoveNEXT
'        Loop
'        .Close
'    End With
'    Set rs = Nothing
'    strOrders = Left(strOrders, Len(strOrders) - 1)
'
'    rpt.FormulaFields.GetItemByName("Orders").Text = "'" & strOrders & "'"
    
    f.ShowReport strSelection, rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
    
Private Sub mnuPrintListBatchWise_Click()

    Dim strFilters As String
    strFilters = "Maker:" & cmbMaker.Text & " | Process:" & Replace(cmbProcess.Text, "'", "") & " | Catagory:" & Replace(cmbItemCat.Text, "'", "") & " | Article:" & cmbItems.Text & " | Customer:" & Replace(cmbCustomer.ID, "'", "")
    
    If Trim(txtLotNo) <> "" Then
        strFilters = strFilters & " | Lot No.:" & Trim(txtLotNo)
    End If
    
    If Trim(txtOrderNo) <> "" Then
        strFilters = strFilters & " | Order No.:" & Trim(txtOrderNo)
    End If
    
    If Trim(txtMasterPONo) <> "" Then
        strFilters = strFilters & " | Master PO:" & Trim(txtMasterPONo)
    End If
    
    If Trim(txtInActiveDays) <> "" Then
        strFilters = strFilters & " | In-Active Days:" & Trim(txtInActiveDays)
    End If
        
    If chkShowMasterPOOnly.Value = vbChecked Then
        strFilters = strFilters & " | Master POs Only"
    End If
    
    If chkMasterPOOpen.Value = vbChecked Then
        strFilters = strFilters & " | Open Master POs Only"
    End If
    
    If chkRepairLots.Value = vbChecked Then
        strFilters = strFilters & " | Repair Lots Only"
    End If
    
    Dim strDateRange As String
    strDateRange = Format(DtFrom.Value, "dd-MMM-yyyy") & " to " & Format(DtTo.Value, "dd-MMM-yyyy")
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\IssuanceList_BatchNoWise.rpt")
    rpt.FormulaFields.GetItemByName("Filters").Text = "'" & strFilters & "'"
    rpt.FormulaFields.GetItemByName("DateRange").Text = "'" & strDateRange & "'"
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, strReportSQL
    
End Sub

Private Sub mnuPrintMakerBalanceReportSummary_Click()
    
    Dim rpt As CRAXDDRT.Report, strSelection As String
    Set rpt = rptApp.OpenReport(RptPath & "\MakerBalanceReport_Summary.rpt")
    
    strSelection = " (NOT ({Makers.VendID} IN[" & strFactoryMakers & "]))"
    
    strSelection = strSelection & " AND {VendIssdDetail.RcvdQty}<{VendIssdDetail.IssQty} AND {VendIssued.Closed}=FALSE"
    
    If cmbProcess.ID <> "0" Then
        strSelection = strSelection & " AND {VendIssued.ProcessID} IN[" & Replace(cmbProcess.ID, "'", "") & "]"
    End If
    
    If cmbMaker.ID <> "0" Then
        strSelection = strSelection & " AND {VendIssued.VendID}=" & cmbMaker.ID
    End If
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt

End Sub

Private Sub mnuPrintMakerIssuanceReport_Click()
        
    If GetUserRightAccess("MIL_Print_Maker_Issuance_Report_Valuewise") = False Then Exit Sub
    
    Dim strFilters As String
    strFilters = "Maker:" & cmbMaker.Text & " | Process:" & Replace(cmbProcess.Text, "'", "") & " | Catagory:" & Replace(cmbItemCat.Text, "'", "") & " | Article:" & cmbItems.Text & " | Customer:" & Replace(cmbCustomer.ID, "'", "")
    
    If Trim(txtLotNo) <> "" Then
        strFilters = strFilters & " | Lot No.:" & Trim(txtLotNo)
    End If
    
    If Trim(txtOrderNo) <> "" Then
        strFilters = strFilters & " | Order No.:" & Trim(txtOrderNo)
    End If
    
    If Trim(txtMasterPONo) <> "" Then
        strFilters = strFilters & " | Master PO:" & Trim(txtMasterPONo)
    End If
    
    If Trim(txtInActiveDays) <> "" Then
        strFilters = strFilters & " | In-Active Days:" & Trim(txtInActiveDays)
    End If
        
    If chkShowMasterPOOnly.Value = vbChecked Then
        strFilters = strFilters & " | Master POs Only"
    End If
    
    If chkMasterPOOpen.Value = vbChecked Then
        strFilters = strFilters & " | Open Master POs Only"
    End If
    
    If chkRepairLots.Value = vbChecked Then
        strFilters = strFilters & " | Repair Lots Only"
    End If
    
    Dim strDateRange As String
    strDateRange = Format(DtFrom.Value, "dd-MMM-yyyy") & " to " & Format(DtTo.Value, "dd-MMM-yyyy")
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\MakerIssuanceReportValuewise.rpt")
    rpt.FormulaFields.GetItemByName("Filters").Text = "'" & strFilters & "'"
    rpt.FormulaFields.GetItemByName("DateRange").Text = "'" & strDateRange & "'"
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, strReportSQL
    
End Sub

Private Sub mnuPrintMakerList_Click()

    Dim strFilters As String
    strFilters = "Maker:" & cmbMaker.Text & " | Process:" & Replace(cmbProcess.Text, "'", "") & " | Catagory:" & Replace(cmbItemCat.Text, "'", "") & " | Article:" & cmbItems.Text & " | Customer:" & Replace(cmbCustomer.ID, "'", "")
    
    If Trim(txtLotNo) <> "" Then
        strFilters = strFilters & " | Lot No.:" & Trim(txtLotNo)
    End If
    
    If Trim(txtOrderNo) <> "" Then
        strFilters = strFilters & " | Order No.:" & Trim(txtOrderNo)
    End If
    
    If Trim(txtMasterPONo) <> "" Then
        strFilters = strFilters & " | Master PO:" & Trim(txtMasterPONo)
    End If
    
    If Trim(txtInActiveDays) <> "" Then
        strFilters = strFilters & " | In-Active Days:" & Trim(txtInActiveDays)
    End If
        
    If chkShowMasterPOOnly.Value = vbChecked Then
        strFilters = strFilters & " | Master POs Only"
    End If
    
    If chkMasterPOOpen.Value = vbChecked Then
        strFilters = strFilters & " | Open Master POs Only"
    End If
    
    If chkRepairLots.Value = vbChecked Then
        strFilters = strFilters & " | Repair Lots Only"
    End If
    
    Dim strDateRange As String
    strDateRange = Format(DtFrom.Value, "dd-MMM-yyyy") & " to " & Format(DtTo.Value, "dd-MMM-yyyy")
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\MakerList_Issuance.rpt")
    rpt.FormulaFields.GetItemByName("Filters").Text = "'" & strFilters & "'"
    rpt.FormulaFields.GetItemByName("DateRange").Text = "'" & strDateRange & "'"
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, strReportSQL
    
End Sub

Private Sub mnuPrintMasterPOAccountsCopy_Click()
    Call PrintMasterPO(False, "ACCOUNTS COPY")
End Sub

Private Sub mnuPrintMasterPOHideRate_Click()
    Call PrintMasterPO(True)
End Sub

Private Sub mnuPrintMasterPOMakerCopy_Click()
    Call PrintMasterPO(False, "MAKER COPY")
End Sub

Private Sub mnuPrintMasterPOStatus_Click()
    Call PrintMasterPOStatus(LV.SelectedItem.SubItems(10))
End Sub

Private Sub mnuPrintOrdersOfPurchaser_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\PurchaserwiseMonthlyReport.rpt")
    Dim f As New frmPrevRpt
    Dim strSelection As String
    strSelection = "{VVendIssdDetail_Simple.VID_EmpID}='" & cmbPurchaser.ID & "'"
    strSelection = strSelection & " AND {VVendIssdDetail_Simple.DT}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#"
    If Trim(txtOrderNo) <> "" Then
        strSelection = strSelection & " AND {FCustomerOrders.OrderNo}='" & Trim(txtOrderNo) & "'"
    End If
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub mnuPrintPTC_Click()

    If LV.SelectedItem.ListSubItems(9) = "0" Then
        MsgBox "This is not a Lot.", vbInformation
        Exit Sub
    End If
    
    Call PrintPTC(LV.SelectedItem.ListSubItems(9), False)
    
End Sub

Private Sub mnuPrintPTCMini_Click()

    If LV.SelectedItem.ListSubItems(9) = "0" Then
        MsgBox "This is not a Lot.", vbInformation
        Exit Sub
    End If
    
    Call PrintPTC(LV.SelectedItem.ListSubItems(9), False, 1)
    
End Sub

Private Sub mnuPrintPTCWithPrice_Click()

    If LV.SelectedItem.ListSubItems(9) = "0" Then
        MsgBox "This is not a Lot.", vbInformation
        Exit Sub
    End If
    
    Call PrintPTC(LV.SelectedItem.ListSubItems(9), True)
    
End Sub

Private Sub mnuPrintPurchaseOrder_Click()
    Call PrintIssSlip
End Sub

Private Sub PrintIssSlip(Optional bPrintGP As Boolean = False, Optional bHideCustomer As Boolean = False, Optional iReportType = 0)

    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    If bPrintGP Then
        Set rpt = rptApp.OpenReport(RptPath & "\IssSlipGP.rpt")
    Else
        If iReportType = 0 Then
            Set rpt = rptApp.OpenReport(RptPath & "\IssSlip.rpt")
        ElseIf iReportType = 1 Then
            Set rpt = rptApp.OpenReport(RptPath & "\IssSlipMin.rpt")
        ElseIf iReportType = 2 Then
            Set rpt = rptApp.OpenReport(RptPath & "\IssSlipTag.rpt")
        ElseIf iReportType = 3 Then
            Set rpt = rptApp.OpenReport(RptPath & "\IssSlipInternal.rpt")
        End If
    End If
    
    
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
    If strList = "" Then
        f.ShowReport "{VendIssued.EntryID} IN[" & LV.SelectedItem.Tag & "]", rpt
    Else
        f.ShowReport "{VendIssued.EntryID} IN[" & Left(strList, Len(strList) - 1) & "]", rpt
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub mnuPrintPurchaseOrderWithoutRate_Click()
    Call PrintIssSlip(False, True)
End Sub

Private Sub mnuPrintPurchasePlan_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\PlanforPurchase.rpt")
    Dim strSelection As String
    
    
    strSelection = "{VVendIssdDetail_Simple.DT}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#"
    If cmbMaker.ID <> "0" Then
        strSelection = strSelection & " AND {VVendIssdDetail_Simple.VendID}=" & cmbMaker.ID
    End If
    
    If cmbPurchaser.ID <> "0" Then
        strSelection = strSelection & " AND {VVendIssdDetail_Simple.VID_EmpID}='" & cmbPurchaser.ID & "'"
    End If
    
    If Trim(txtOrderNo) <> "" Then
        strSelection = strSelection & " AND {VrptOrders_ForProduction.OrderNo}='" & Trim(txtOrderNo) & "'"
    End If
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub mnuPrintR2OutlookReport_Click()

    Dim rpt As CRAXDDRT.Report, rptApp_Temp As New CRAXDDRT.Application
    Set rpt = rptApp_Temp.OpenReport(RptPath & "\R2OutLookReport.rpt")
     
    Dim strItemGroups As String, strItemCats As String, strCusts As String
    
    If cmbItemCat.ID <> "0" Then
        If cmbItemCat.Functionality = 1 Then
            strItemCats = cmbItemCat.ID
        Else
            strItemCats = "'" & cmbItemCat.ID & "'"
        End If
    Else
        strItemCats = ""
    End If
     
    If cmbItemGroup.ID <> "0" Then
        If cmbItemGroup.Functionality = 1 Then
            strItemGroups = cmbItemGroup.ID
        Else
            strItemGroups = "'" & cmbItemGroup.ID & "'"
        End If
    Else
        strItemGroups = ""
    End If
     
    If cmbCustomer.ID <> "0" Then
        If cmbCustomer.Functionality = 1 Then
            strCusts = cmbCustomer.ID
        Else
            strCusts = "'" & cmbCustomer.ID & "'"
        End If
    Else
        strCusts = ""
    End If
     
    rpt.ParameterFields.GetItemByName("@CustCode").SetCurrentValue strCusts
    rpt.ParameterFields.GetItemByName("@ItemGroup").SetCurrentValue strItemGroups
    rpt.ParameterFields.GetItemByName("@ItemCats").SetCurrentValue strItemCats
    rpt.ParameterFields.GetItemByName("@Type").SetCurrentValue ""
    
    rpt.EnableParameterPrompting = False
    
    rpt.FormulaFields.GetItemByName("CustomerFilter").Text = "'" & Replace(cmbCustomer.Text, "'", "") & "'"
    rpt.FormulaFields.GetItemByName("GroupFilter").Text = "'" & Replace(cmbItemGroup.Text, "'", "") & "'"
    rpt.FormulaFields.GetItemByName("CatagoryFilter").Text = "'" & Replace(cmbItemCat.Text, "'", "") & "'"
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, , False
    
End Sub

Private Sub mnuPrintSFStock_Click()

    Dim strCond As String
    strCond = "{@NetQty}>0"
    If cmbItems.ID <> "0" Then
        strCond = strCond & " AND {StockOrderOpening.ItemID}='" & cmbItems.ID & "'"
    End If
    If cmbProcess.ID <> "0" Then
        strCond = strCond & " AND {StockOrderOpening.ProcessID}=" & cmbProcess.ID
    End If
    
    Dim strStores As String
    Dim frm As New frmGetStores
    Load frm
    If frm.getData(strStores) = False Then Exit Sub
    
    If strStores <> "All" Then
        strCond = strCond & " AND {StockOrderOpening.Store_RefID} IN[" & Replace(strStores, "'", "") & "]"
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\SFStockReportStorewise.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport strCond, rpt
    
End Sub

Private Sub mnuPrintTag_Click()
    Call PrintIssSlip(False, False, 2)
End Sub

Private Sub mnuPrintThisList_Click()

    Dim strFilters As String
    strFilters = "Maker:" & cmbMaker.Text & " | Process:" & Replace(cmbProcess.Text, "'", "") & " | Catagory:" & Replace(cmbItemCat.Text, "'", "") & " | Article:" & cmbItems.Text & " | Customer:" & Replace(cmbCustomer.ID, "'", "")
    
    If Trim(txtLotNo) <> "" Then
        strFilters = strFilters & " | Lot No.:" & Trim(txtLotNo)
    End If
    
    If Trim(txtOrderNo) <> "" Then
        strFilters = strFilters & " | Order No.:" & Trim(txtOrderNo)
    End If
    
    If Trim(txtMasterPONo) <> "" Then
        strFilters = strFilters & " | Master PO:" & Trim(txtMasterPONo)
    End If
    
    If Trim(txtInActiveDays) <> "" Then
        strFilters = strFilters & " | In-Active Days:" & Trim(txtInActiveDays)
    End If
        
    If chkShowMasterPOOnly.Value = vbChecked Then
        strFilters = strFilters & " | Master POs Only"
    End If
    
    If chkMasterPOOpen.Value = vbChecked Then
        strFilters = strFilters & " | Open Master POs Only"
    End If
    
    If chkRepairLots.Value = vbChecked Then
        strFilters = strFilters & " | Repair Lots Only"
    End If
    
    Dim strDateRange As String
    strDateRange = Format(DtFrom.Value, "dd-MMM-yyyy") & " to " & Format(DtTo.Value, "dd-MMM-yyyy")
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\IssuanceList.rpt")
    rpt.FormulaFields.GetItemByName("Filters").Text = "'" & strFilters & "'"
    rpt.FormulaFields.GetItemByName("DateRange").Text = "'" & strDateRange & "'"
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, strReportSQL
    
End Sub

Private Sub mnuPurchaseRegister_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\PurchaseCalendar.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{DateLookup.DateFull}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#", rpt
     
    
End Sub

Private Sub mnuPrintVendorPurchaseStatusLog_Click()
    Call PrintVendorPurchaseStatusLog(0)
End Sub

Private Sub mnuPrintVendorPurchaseStatusLogBalanceOnly_Click()
    Call PrintVendorPurchaseStatusLog(1)
End Sub

Private Sub mnuPrintVendorPurchaseStatusLogNillOnly_Click()
    Call PrintVendorPurchaseStatusLog(2)
End Sub

Private Sub mnuPurchaserwisePurchaseplan_Click()
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Purchaserwise_PurchasePlan.rpt")
    
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DtFrom.Value
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DtTo.Value
    rpt.ParameterFields.GetItemByName("@PurchaserID").SetCurrentValue IIf(cmbPurchaser.ID = "0", "", cmbPurchaser.ID)
    
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, , False
    
End Sub

Private Sub txtFind_Change()
    cmdFind.Enabled = Len(Trim(txtFind)) > 0
End Sub

Public Sub ShowIssuanceList(p_lMakerID As Long, p_lProcessID As Long, p_DtFrom As Date, p_DtTo As Date)

    cmbMaker.ID = p_lMakerID
    cmbProcess.ID = p_lProcessID
    cmbRange.ListIndex = cmbRange.ListCount - 1
    DtFrom.Value = p_DtFrom
    DtTo.Value = p_DtTo
    Call RefreshLV
    
    Me.Show
    
End Sub

Private Sub PrintVendorPurchaseStatusLog(Optional iReportType As Integer = 0)

    Dim strSelection As String
    strSelection = "{VendIssued.DT}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "# AND {VendIssued.MasterPONo}<>''"
    If cmbMaker.ID <> "0" Then
        strSelection = strSelection & " AND {VendIssued.VendID}=" & cmbMaker.ID
    End If
    If cmbCustomer.ID <> "0" Then
        strSelection = strSelection & " AND {FCustomerOrders.CustCode}='" & cmbCustomer.ID & "'"
    End If
    
    Dim srSelection_Group As String
    If iReportType = 1 Then 'Balance
        'strSelection = strSelection & " AND {VendIssdDetail.RcvdQty}<{VendIssdDetail.IssQty}"
        srSelection_Group = "SUM({@BalQty}, {VendIssdDetail.OrderNo})>0"
        strSelection = strSelection & " AND {VendIssued.Closed}=FALSE"
    ElseIf iReportType = 2 Then  'Nill
        'strSelection = strSelection & " AND {VendIssdDetail.RcvdQty}>={VendIssdDetail.IssQty}"
        srSelection_Group = "SUM({@BalQty}, {VendIssdDetail.OrderNo})=0"
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\VendorPurchaseStatusLog.rpt")
    rpt.GroupSelectionFormula = srSelection_Group
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
End Sub
