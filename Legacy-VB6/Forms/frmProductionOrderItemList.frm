VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmProductionOrderItemList 
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   14385
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
   ScaleHeight     =   9885
   ScaleMode       =   0  'User
   ScaleWidth      =   12884.84
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   15
      TabIndex        =   14
      Top             =   0
      Width           =   14325
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Production Item List"
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
         TabIndex        =   15
         Top             =   165
         Width           =   14175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Production Item List"
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
         TabIndex        =   16
         Top             =   180
         Width           =   14160
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
      Height          =   9225
      Left            =   15
      TabIndex        =   12
      Top             =   615
      Width           =   14325
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   45
         TabIndex        =   0
         Top             =   345
         Width           =   1440
         _ExtentX        =   2540
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
         Left            =   7965
         TabIndex        =   5
         Top             =   330
         Width           =   1650
         _ExtentX        =   2910
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
      Begin ComboList.Usercontrol1 cmbOrderNo 
         Height          =   285
         Left            =   9630
         TabIndex        =   6
         Top             =   330
         Width           =   2025
         _ExtentX        =   3572
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
         Left            =   7005
         TabIndex        =   10
         Top             =   900
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
      Begin ComboList.Usercontrol1 cmbMaker 
         Height          =   285
         Left            =   3525
         TabIndex        =   9
         Top             =   900
         Width           =   3465
         _ExtentX        =   6112
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
         TabIndex        =   8
         Top             =   900
         Width           =   3465
         _ExtentX        =   6112
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
      Begin ComboList.Usercontrol1 cmbProcessEx 
         Height          =   285
         Left            =   45
         TabIndex        =   58
         Top             =   1455
         Width           =   3465
         _ExtentX        =   6112
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
      Begin VB.Frame FHidden 
         Caption         =   "This is Hidden"
         Height          =   2910
         Left            =   5280
         TabIndex        =   54
         Top             =   4035
         Visible         =   0   'False
         Width           =   4155
         Begin VB.OptionButton OptR1 
            Caption         =   "R1"
            Height          =   270
            Index           =   1
            Left            =   1680
            TabIndex        =   57
            Top             =   1200
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.OptionButton OptR1 
            Caption         =   "R2"
            Height          =   270
            Index           =   0
            Left            =   3075
            TabIndex        =   56
            Tag             =   "2"
            Top             =   1200
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.OptionButton OptR1 
            Caption         =   "Both"
            Height          =   270
            Index           =   2
            Left            =   285
            TabIndex        =   55
            Top             =   1200
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   960
         End
      End
      Begin VB.CheckBox chkIssued 
         Caption         =   "Issuance"
         ForeColor       =   &H00375DEB&
         Height          =   255
         Left            =   9795
         TabIndex        =   53
         Top             =   1500
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.CheckBox chkReceived 
         Caption         =   "Receiving"
         ForeColor       =   &H00375DEB&
         Height          =   255
         Left            =   8460
         TabIndex        =   52
         Top             =   1500
         Value           =   1  'Checked
         Width           =   1080
      End
      Begin VB.ComboBox cmbHubs 
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
         ItemData        =   "frmProductionOrderItemList.frx":0000
         Left            =   11685
         List            =   "frmProductionOrderItemList.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   885
         Width           =   1575
      End
      Begin VB.CheckBox chkNotYetStarted 
         Caption         =   "Not Yet Started"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   6690
         TabIndex        =   48
         Top             =   1500
         Value           =   1  'Checked
         Width           =   1545
      End
      Begin VB.CheckBox chkLotsOnly 
         Caption         =   "Lots Only"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5355
         TabIndex        =   47
         Top             =   1500
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.CheckBox chkRegularLots 
         Caption         =   "Regular Lots"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   12720
         TabIndex        =   46
         Top             =   1485
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkRepairLots 
         Caption         =   "Repair Lots"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   11130
         TabIndex        =   45
         Top             =   1485
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox txtLotNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   11670
         TabIndex        =   7
         Top             =   330
         Width           =   1575
      End
      Begin VB.CheckBox chkUseOrderDates 
         Caption         =   "Use Order Dates"
         Height          =   195
         Left            =   10230
         TabIndex        =   40
         Top             =   4395
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.CheckBox chkUrgent 
         Caption         =   "Urgent"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   10425
         TabIndex        =   31
         Top             =   4605
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CheckBox chkBoth 
         Caption         =   "Both"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   9750
         TabIndex        =   27
         Top             =   4605
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.CheckBox chkNoDelivery 
         Caption         =   "No Delivery Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   9750
         TabIndex        =   26
         Top             =   4350
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox txtEdit 
         Height          =   285
         Left            =   4050
         TabIndex        =   43
         Top             =   2490
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtTotalAmt 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   12345
         TabIndex        =   42
         Top             =   8310
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Frame FraSp 
         Height          =   480
         Left            =   5550
         TabIndex        =   36
         Top             =   8265
         Width           =   2835
         Begin VB.CommandButton cmdLoadList 
            Caption         =   "Load Selection"
            Height          =   270
            Left            =   1425
            TabIndex        =   38
            Top             =   150
            Width           =   1305
         End
         Begin VB.CommandButton cmdSaveList 
            Caption         =   "Save Selection"
            Height          =   270
            Left            =   75
            TabIndex        =   37
            Top             =   150
            Width           =   1305
         End
      End
      Begin MSComCtl2.DTPicker DTDelivery 
         Height          =   300
         Left            =   3615
         TabIndex        =   28
         Top             =   1980
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yy"
         Format          =   155058179
         CurrentDate     =   40642
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
         TabIndex        =   20
         Top             =   8265
         Width           =   5475
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
            TabIndex        =   23
            Top             =   135
            Width           =   2430
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
            Left            =   4395
            TabIndex        =   22
            Top             =   135
            Width           =   1020
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
            ItemData        =   "frmProductionOrderItemList.frx":0004
            Left            =   60
            List            =   "frmProductionOrderItemList.frx":0011
            Style           =   2  'Dropdown List
            TabIndex        =   21
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
         ItemData        =   "frmProductionOrderItemList.frx":0034
         Left            =   3150
         List            =   "frmProductionOrderItemList.frx":0047
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   330
         Width           =   1725
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   6435
         TabIndex        =   4
         Top             =   330
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
         Format          =   155058179
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   4875
         TabIndex        =   3
         Top             =   330
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
         Format          =   155058179
         CurrentDate     =   37055
      End
      Begin MSComctlLib.ListView LV 
         Height          =   6495
         Left            =   45
         TabIndex        =   17
         Top             =   1770
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   11456
         View            =   3
         LabelEdit       =   1
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
         NumItems        =   16
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Start "
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Lot #"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Item Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Item Description"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Status"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Last Days"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Remarks"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "TL Running/Orders"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "Semi"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   11
            Text            =   "Pack"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Order #"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Status"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   14
            Text            =   "Price"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   15
            Text            =   "Hub"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         Caption         =   "  Next Process"
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
         Left            =   45
         TabIndex        =   59
         Top             =   1200
         Width           =   3465
      End
      Begin MSForms.CommandButton cmdPrintHubwise 
         Height          =   375
         Left            =   8985
         TabIndex        =   51
         Top             =   8775
         Width           =   1230
         ForeColor       =   0
         Caption         =   "Hubwise"
         PicturePosition =   327683
         Size            =   "2170;661"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrintValue 
         Height          =   375
         Left            =   10260
         TabIndex        =   49
         Top             =   8775
         Width           =   1500
         ForeColor       =   0
         Caption         =   "Print (Value)"
         PicturePosition =   327683
         Size            =   "2646;661"
         Accelerator     =   86
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   1485
         TabIndex        =   1
         Top             =   345
         Width           =   1650
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2910;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         Caption         =   $"frmProductionOrderItemList.frx":008C
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
         Left            =   45
         TabIndex        =   44
         Top             =   645
         Width           =   13200
      End
      Begin MSForms.CommandButton cmdPrintPriceDiff 
         Height          =   375
         Left            =   855
         TabIndex        =   41
         Top             =   8775
         Visible         =   0   'False
         Width           =   1200
         ForeColor       =   0
         Caption         =   "PD"
         PicturePosition =   327683
         Size            =   "2117;661"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdSPPlanning 
         Height          =   375
         Left            =   7440
         TabIndex        =   39
         Top             =   8775
         Visible         =   0   'False
         Width           =   1230
         ForeColor       =   0
         Caption         =   "SP Report"
         PicturePosition =   327683
         Size            =   "2170;661"
         Accelerator     =   83
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrintListWithValue 
         Height          =   375
         Left            =   2100
         TabIndex        =   35
         Top             =   8775
         Visible         =   0   'False
         Width           =   1200
         ForeColor       =   0
         Caption         =   "P Value"
         PicturePosition =   327683
         Size            =   "2117;661"
         Accelerator     =   86
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrintR2 
         Height          =   375
         Left            =   3330
         TabIndex        =   34
         Top             =   8775
         Visible         =   0   'False
         Width           =   945
         ForeColor       =   0
         Caption         =   "R2"
         PicturePosition =   327683
         Size            =   "1667;661"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrintR1 
         Height          =   375
         Left            =   4320
         TabIndex        =   33
         Top             =   8775
         Visible         =   0   'False
         Width           =   945
         ForeColor       =   0
         Caption         =   "R1"
         PicturePosition =   327683
         Size            =   "1667;661"
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
         Left            =   5295
         TabIndex        =   32
         Top             =   8775
         Visible         =   0   'False
         Width           =   1980
         ForeColor       =   0
         Caption         =   "Order Status Report"
         PicturePosition =   327683
         Size            =   "3492;661"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrintItemBalance 
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   8775
         Visible         =   0   'False
         Width           =   1575
         ForeColor       =   0
         Caption         =   "Balance Report"
         PicturePosition =   327683
         Size            =   "2778;661"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPPRSelected 
         Height          =   375
         Left            =   7305
         TabIndex        =   29
         Top             =   8775
         Visible         =   0   'False
         Width           =   1575
         ForeColor       =   0
         Caption         =   "Selected PPR"
         PicturePosition =   327683
         Size            =   "2778;661"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrintPlanningReport 
         Height          =   375
         Left            =   7185
         TabIndex        =   25
         Top             =   8775
         Visible         =   0   'False
         Width           =   1575
         ForeColor       =   0
         Caption         =   "Planning Report"
         PicturePosition =   327683
         Size            =   "2778;661"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   375
         Left            =   11805
         TabIndex        =   24
         Top             =   8775
         Width           =   1200
         ForeColor       =   0
         Caption         =   "Print"
         PicturePosition =   327683
         Size            =   "2117;661"
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
         Caption         =   "Total Amount (Rs.):"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -30
         TabIndex        =   19
         Top             =   8790
         Visible         =   0   'False
         Width           =   2130
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   1290
         Left            =   13320
         TabIndex        =   11
         Top             =   150
         Width           =   915
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "1614;2275"
         Accelerator     =   82
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
         Left            =   13050
         TabIndex        =   18
         Top             =   8775
         Width           =   1200
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "2117;661"
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
         Caption         =   $"frmProductionOrderItemList.frx":0183
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
         TabIndex        =   13
         Top             =   120
         Width           =   13200
      End
   End
End
Attribute VB_Name = "frmProductionOrderItemList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQLString As String, SQLStringForR1R2 As String
Dim bForProduction As Boolean
Dim strReportSQL As String
Dim bAuthorized As Boolean
Dim lListEntryID As Long
Dim iColNo As Integer

Private Sub mnuProdPlan_Click()
    Load frmProdPlan
    Call frmProdPlan.AddPlanning(Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1))
End Sub

Public Sub ShowMe(p_bForProduction As Boolean)
    bForProduction = p_bForProduction
    If bForProduction Then
        Label2.Visible = False
        txtTotalAmt.Visible = False
    End If
    Me.Show
    Call RefreshLV
End Sub

Private Sub chkUseOrderDates_Click()

    If chkUseOrderDates.Value = vbChecked Then
        cmdPrintPlanningReport.Enabled = False
    Else
        cmdPrintPlanningReport.Enabled = True
    End If
    
End Sub

Private Sub cmbCust_matched()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    SQL = "Select Distinct Country From ForeignCustomers "
    If cmbCust.ID <> "0" Then
        SQL = SQL & " Where CustCode='" & cmbCust.ID & "'"
    End If
    
    With rs
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
        cmbCountry.AddItem "<All Countries>"
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            'cmbCountry.List(cmbCountry.ListCount - 1, 1) = ![Curr] & ""
            .MoveNEXT
        Loop
        .Close
    End With
    cmbCountry.ListIndex = 0
    
    Set rs = Nothing
    
    cmbOrderNo.ClearVals
    cmbOrderNo.AddItem "<All Orders>", "0"
    If cmbCust.ID = "0" Then
        cmbOrderNo.AddVals con, "OrderNo + ' {' + InternalRefNo + '}'", "FCustomerOrders", "OrderNo"
    Else
        cmbOrderNo.AddVals con, "OrderNo + ' {' + InternalRefNo + '}'", "FCustomerOrders", "OrderNo", " WHERE CustCode='" & cmbCust.ID & "'"
    End If
    cmbOrderNo.ID = "0"
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbRange_Click()

    DtTo = ServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    Select Case cmbRange.ListIndex
        Case 1
            DtTo = DateAdd("d", 15, ServerDate)
        Case 2
            DtTo = DateAdd("d", 30, ServerDate)
        Case 3
            DtTo = DateAdd("d", 45, ServerDate)
        Case 4
            DtFrom.Enabled = True
            DtTo.Enabled = True
            DtFrom.SetFocus
            Exit Sub
    End Select
    
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
        
        If cmbfield.ListIndex = 0 Then  'Lot No.
            fldValue = LV.ListItems(i).SubItems(2)
        ElseIf cmbfield.ListIndex = 1 Then  'Item Code
            fldValue = LV.ListItems(i).SubItems(3)
        ElseIf cmbfield.ListIndex = 2 Then  'Order No.
            fldValue = LV.ListItems(i).SubItems(12)
        End If
        
        If LCase(fldValue) Like "*" & LCase(Trim(txtFind)) & "*" Then
            LV.ListItems(i).Selected = True
            LV.ListItems(i).EnsureVisible
            If i <> StartIdx Then Exit For
        End If
       
    Next
    
    If i = LV.ListItems.count + 1 Then
        MsgBox "Record Not Found.", vbInformation
        Set LV.SelectedItem = Nothing
    End If

End Sub

Private Sub cmdLoadList_Click()

    Dim lEntryID As Long, strEntryID As String
    Dim f As New frmGetSingleComboValue_Ex
    Load f
    
    If f.getComboValue(strEntryID, "List Name", "List", False) = False Then Exit Sub
    
    Call RefreshLV(Val(strEntryID))
    
End Sub

Private Sub cmdPPRSelected_Click()
    
    Dim strItems As String
    Dim i As Integer
    strItems = ""
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            strItems = strItems & "'" & LV.ListItems(i).SubItems(5) & "',"
        End If
    Next
    
    If strItems <> "" Then
        strItems = Left(strItems, Len(strItems) - 1)
    Else
        MsgBox "No Item Selected.", vbInformation
        Exit Sub
    End If
    
    If LV.SelectedItem Is Nothing Then
        MsgBox "Please Select Item.", vbInformation
        Exit Sub
    End If
    
    Dim rpt As CRAXDDRT.Report
    
    Set rpt = rptApp.OpenReport(RptPath & "\OrderPlanningReport.rpt")
    
    'rpt.ParameterFields.GetItemByName("@OrderNo").SetCurrentValue strOrderNo
    rpt.ParameterFields.GetItemByName("@FromItems").SetCurrentValue True
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DtFrom.Value
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DtTo.Value
    rpt.ParameterFields.GetItemByName("@NoDelivery").SetCurrentValue IIf(chkNoDelivery.Value = vbChecked, True, False)
    rpt.ParameterFields.GetItemByName("@Both").SetCurrentValue IIf(chkBoth.Value = vbChecked, True, False)
    
    If cmbCust.ID <> "0" Then
        rpt.ParameterFields.GetItemByName("@CustCode").SetCurrentValue cmbCust.ID
    End If
    
    If cmbCountry.ListIndex > 0 Then
        rpt.ParameterFields.GetItemByName("@Country").SetCurrentValue cmbCountry.Text
    End If
    
    rpt.ParameterFields.GetItemByName("@ItemCode").SetCurrentValue strItems
    rpt.ParameterFields.GetItemByName("@TempList_RefID").SetCurrentValue 0
    
    rpt.EnableParameterPrompting = False
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, , False
    
End Sub

Private Sub cmdPrint_Click()

    Call PrintList
    
End Sub

Private Sub PrintList(Optional bWithValue As Boolean = False, Optional bSPPlanning As Boolean = False, Optional bHubwise As Boolean = False)

    Dim rpt As CRAXDDRT.Report
    
    If bWithValue Then
        Set rpt = rptApp.OpenReport(RptPath & "\ProductionItemList_With_Value.rpt")
    ElseIf bHubwise Then
        Set rpt = rptApp.OpenReport(RptPath & "\ProductionItemList_Hubwise.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\ProductionItemList.rpt")
    End If
    
    Dim strFilters As String
    
'    strFilters = Format(DtFrom.Value, "dd-MMM-yyyy") & " to " & Format(DtTo.Value, "dd-MMM-yyyy")
'
'    strFilters = strFilters & " | Customer:" & cmbCust.Text
'    strFilters = strFilters & " | Country:" & cmbCountry.Text
     
'    If chkUrgent.Value = vbChecked Then
'         strFilters = strFilters & " | Urgent"
'    Else
'        If chkNoDelivery.Value = vbChecked Then
'            strFilters = strFilters & " | No Delivery Date"
'        End If
'
'        If chkBoth.Value = vbChecked Then
'            strFilters = strFilters & " | Both"
'        End If
'    End If
    'rpt.FormulaFields.GetItemByName("Filters").Text = "'" & strFilters & "'"
    If bWithValue Then
        Dim rptSub As CRAXDDRT.Report
        Set rptSub = rpt.OpenSubreport("Summary")
        rptSub.SQLQueryString = SQLString
    End If
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, SQLString
    
End Sub

Private Sub cmdPrintHubwise_Click()
    Call PrintList(False, False, True)
End Sub

Private Sub cmdPrintItemBalance_Click()
    
    Dim strItems As String
    Dim i As Integer
    strItems = ""
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            strItems = strItems & "'" & LV.ListItems(i).SubItems(5) & "',"
        End If
    Next
    If strItems <> "" Then
        strItems = Left(strItems, Len(strItems) - 1)
    Else
        MsgBox "No Item Selected.", vbInformation
        Exit Sub
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptPartiallyShippedOrdersItemwise.rpt")
    Dim strSelection As String
    strSelection = "{FOrderItems.CompItemCode} IN[" & strItems & "] AND {VFOrderItemswithShippedQty.Qty}>{VFOrderItemswithShippedQty.ShippedQty}"
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub cmdPrintListWithValue_Click()
    If getDBPassword("ViewPValueReport") = False Then
        Exit Sub
    End If
    Call PrintList(True)
End Sub

Private Sub cmdPrintOrderStatus_Click()
        
    Dim strRet As String
    strRet = ShowPopUpMenu1("Complete", "R1", "R2")
    
    Select Case strRet
    
        Case "Complete"
            Call PrintOrderStatusReport(False, True)
        Case "R1"
            Call PrintOrderStatusReport(True)
        Case "R2"
            Call PrintOrderStatusReport(False)
            
    End Select
    
End Sub

Private Sub PrintOrderStatusReport(Optional bR1 As Boolean = False, Optional bShowComplete As Boolean = False)

    Dim i As Integer, strIDs As String
    For i = 1 To LV.ListItems.count
        strIDs = strIDs & Val(LV.ListItems(i).key) & ","
    Next
    
    If strIDs <> "" Then
        strIDs = Left(strIDs, Len(strIDs) - 1)
    End If
    
    Dim rpt As CRAXDDRT.Report
    'Set rpt = rptApp.OpenReport(RptPath & "\OrderStatusReport.rpt")
    Set rpt = rptApp.OpenReport(RptPath & "\OrderStatusReport_New.rpt")
    rpt.ParameterFields.GetItemByName("@FOI_IDs").SetCurrentValue strIDs
    
    rpt.EnableParameterPrompting = False
    
    Dim strSelection As String
    If bShowComplete = False Then
        If bR1 Then
            strSelection = "{OrderStatusReport_SP;1.ForgingQty}>0 OR {OrderStatusReport_SP;1.MachiningQty}>0 OR {OrderStatusReport_SP;1.FilingQty}>0 OR {OrderStatusReport_SP;1.TemperQty}>0 OR {OrderStatusReport_SP;1.PolishQty}>0 OR {OrderStatusReport_SP;1.QAQty}>0"
        Else
            strSelection = "{OrderStatusReport_SP;1.QAQty}>0 OR {OrderStatusReport_SP;1.SettingQty}>0 OR {OrderStatusReport_SP;1.BoilQty}>0 OR {OrderStatusReport_SP;1.FinalPolQty}>0 OR {OrderStatusReport_SP;1.FinalInspQty}>0"
        End If
    End If
    
    Dim frm As New frmPrevRpt
    frm.ShowReport strSelection, rpt, "", False
    
End Sub


Private Sub cmdPrintPlanningReport_Click()

    Dim rpt As CRAXDDRT.Report
    
    Set rpt = rptApp.OpenReport(RptPath & "\OrderPlanningReport.rpt")
    
    'rpt.ParameterFields.GetItemByName("@OrderNo").SetCurrentValue strOrderNo
    rpt.ParameterFields.GetItemByName("@FromItems").SetCurrentValue True
    
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DtFrom.Value
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DtTo.Value
    
    rpt.ParameterFields.GetItemByName("@NoDelivery").SetCurrentValue IIf(chkNoDelivery.Value = vbChecked, True, False)
    rpt.ParameterFields.GetItemByName("@Both").SetCurrentValue IIf(chkBoth.Value = vbChecked, True, False)
    rpt.ParameterFields.GetItemByName("@Urgent").SetCurrentValue IIf(chkUrgent.Value = vbChecked, True, False)
    rpt.ParameterFields.GetItemByName("@ItemCode").SetCurrentValue "0"
    
    rpt.ParameterFields.GetItemByName("@GroupID").SetCurrentValue Val(cmbItemGroup.ID)
    
    If cmbCust.ID <> "0" Then
        rpt.ParameterFields.GetItemByName("@CustCode").SetCurrentValue cmbCust.ID
    End If
    
    If cmbCountry.ListIndex > 0 Then
        rpt.ParameterFields.GetItemByName("@Country").SetCurrentValue cmbCountry.Text
    End If
    
    rpt.ParameterFields.GetItemByName("@TempList_RefID").SetCurrentValue lListEntryID
    
    If cmbOrderNo.MatchFound And cmbOrderNo.ID <> "0" Then
        rpt.ParameterFields.GetItemByName("@OrderNo").SetCurrentValue cmbOrderNo.ID
    End If
    
    rpt.EnableParameterPrompting = False
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, , False
    
End Sub
    
Private Sub PrintR1R2PlanningReport(Optional bR1 As Boolean)

    On Error GoTo err
    'Prepare Data for the Report...
    
    Dim strCondition As String
    strCondition = getCondition()
    If LV.SelectedItem Is Nothing Then
        MsgBox "Please Select Item.", vbInformation
        Exit Sub
    End If
     
    Dim cmd As New ADODB.Command
    With cmd
        Set .ActiveConnection = con_ServerSide
        .CommandType = adCmdStoredProc
        .CommandTimeout = 300
        .CommandText = "PreparePlanningReport_SP"
        .Parameters("@OrderNo").Value = cmbOrderNo.ID
        .Parameters("@R1").Value = bR1
        .Parameters("@ItemCodes").Value = strCondition 'strItems
        .Execute
    End With
     
    Dim rpt As CRAXDDRT.Report
    If bR1 Then
        Set rpt = rptApp.OpenReport(RptPath & "\OrderPlanningR1.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\OrderPlanningR2.rpt")
    End If
    rpt.FormulaFields.GetItemByName("OrderWise").Text = False
    
    rpt.FormulaFields.GetItemByName("OrderNo").Text = "'" & cmbOrderNo.ID & "'"
    Dim f As New frmPrevRpt
    Dim strConditionReport As String
    strConditionReport = getConditionitionForReport
    f.ShowReport strConditionReport, rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdPrintPriceDiff_Click()

    If getDBPassword("Export_Order_CC_RateDiff") = False Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\OrderItems_CustomerCatalog_RateDiff.rpt")
    Dim f As New frmPrevRpt
    Dim strSQLEdit As String
    strSQLEdit = Left(strReportSQL, InStr(strReportSQL, "ORDER BY") - 1)
    strSQLEdit = strSQLEdit & " AND ROUND(VFOrderItems.Price,3)<>ROUND(FCustomerCatalog.ExWorks,3)"
    f.ShowReport "", rpt, strSQLEdit
    
End Sub

Private Sub cmdPrintR1_Click()
    Call PrintR1R2PlanningReport(True)
End Sub

Private Sub cmdPrintR2_Click()
    Call PrintR1R2PlanningReport(False)
End Sub

Private Sub cmdPrintValue_Click()
    Call PrintList(True)
End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub

Private Sub cmdSaveList_Click()

    On Error GoTo err
    
    Dim strName As String, strRemarks As String
    Dim f As New frmGetDoubleTextValues
    If f.getTextValue(strName, strRemarks, False, "List Name", "List", "Remarks") = False Then Exit Sub
     
    Dim i As Integer, lEntryID As Long
    Call StartTrans(con)
    
    con.Execute "INSERT INTO POI_List(ListName,Remarks,UserName,MachineName) VALUES('" & strName & "','" & strRemarks & _
     "','" & CurrentUserName & "','" & strComputerName & "')"
    
    lEntryID = GetSingleLongValue("MAX(EntryID)", "POI_List", " WHERE MachineName='" & strComputerName & "'")
    
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                con.Execute "INSERT INTO POI_List_Detail(POI_List_RefID,EntryType,LotNo,OrderNo,ItemCode) VALUES(" & lEntryID & _
                 "," & Val(.ListSubItems(1).Tag) & ",'" & .SubItems(2) & "','" & .SubItems(12) & "','" & .SubItems(3) & "')"
            End If
        End With
    Next
    
    con.CommitTrans
    
    Exit Sub
err:
    MsgBox err.Description
End Sub




Private Sub cmdSPPlanning_Click()
    Call PrintList(False, True)
End Sub



Private Sub DTDelivery_KeyDown(KeyCode As Integer, Shift As Integer)

    
    If KeyCode = 13 Then
         
        Dim lID As Long, iStatus As Integer
        lID = Val(LV.SelectedItem.key)
        
        
        
        If iColNo = 8 Then
            
            
            
            iStatus = Val(LV.SelectedItem.ListSubItems(7).Tag)
            
            Call StartTrans(con)
            'con.Execute "UPDATE FOrderItems SET DeliveryStatus=1,DeliveryDT='" & DTDelivery.Value & "' WHERE ID=" & lID
            con.Execute "UPDATE FOrderItems SET DeliveryDT='" & DTDelivery.Value & "',DeliveryStatus=CASE WHEN DeliveryStatus=0 THEN 1 ELSE DeliveryStatus END WHERE ID=" & lID
            con.Execute "INSERT INTO FOrderItems_DeliveryRevisions(FOI_RefID,DeliveryStatus,DeliveryDT,UserName,MachineName)" & _
             "VALUES(" & lID & "," & iStatus & ",'" & DTDelivery.Value & "','" & CurrentUserName & "','" & strComputerName & _
             "')"
            
            If (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Towne" Or strCompany = "Tecnimen" Then
                con.Execute "INSERT INTO FOrderItems_DeliveryDates_PTW(FOI_RefID) VALUES(" & lID & ")"
            End If
            con.CommitTrans
            
            LV.SelectedItem.SubItems(7) = Format(DTDelivery.Value, "dd-MMM-yyyy")
            
            If LV.SelectedItem.Index < LV.ListItems.count Then
                LV.ListItems(LV.SelectedItem.Index + 1).Selected = True
                Call LV_DblClick
                Exit Sub
            End If
            
        Else
            Dim iRowAffect As Integer
            Call StartTrans(con)
            con.Execute "UPDATE FOrderItems_MoreDetails SET R2TransferDT='" & DTDelivery.Value & "' WHERE FOI_RefID=" & lID, iRowAffect
            If iRowAffect = 0 Then
                con.Execute "INSERT INTO FOrderItems_MoreDetails(FOI_RefID,R2TransferDT) VALUES(" & lID & ",'" & DTDelivery.Value & "')"
            End If
            LV.SelectedItem.SubItems(iColNo - 1) = Format(DTDelivery, "dd-MMM-yy")
            con.CommitTrans
        End If
        
        
        
        DTDelivery.Visible = False
        'Call RefreshLV
        
    ElseIf KeyCode = 27 Then
        DTDelivery.Visible = False
    End If
    
End Sub

Private Sub DTDelivery_LostFocus()
    DTDelivery.Visible = False
End Sub

Private Sub Form_Load()
    
    'mnuPop.Visible = False
    bForProduction = False
    
    cmbCust.ListHeight = 2500
    cmbCust.AddItem "<All Customers>", "0"
    'Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where custCode<>'Hico'")
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode")
    
    cmbCust.ID = "0"
    cmbCountry.ListIndex = 0
    
    cmbItemGroup.AddItem "<All Groups>", "0"
    cmbItemGroup.AddVals con, "Description", "ItemGroups", "ID", " ORDER BY Description"
    cmbItemGroup.ID = "0"
    
    'Call RefreshLV
    
    DtFrom.Value = Date
    DtTo.Value = Date
    
    cmbRange.ListIndex = 0
    cmbfield.ListIndex = 2
     
    bAuthorized = False
    
    If Not (strCompany = "Dr-Frgz" Or strCompany = "Towne" Or strCompany = "Tecnimen") Then
        cmdPrintR1.Visible = False
        cmdPrintR2.Visible = False
        cmdPrintOrderStatus.Visible = False
    End If
        
    
    cmbProcess.AddItem "<All Processes>", "0"
    cmbProcess.AddVals con, "Description", "Processes", "ProcessID", " ORDER BY SNo"
    cmbProcess.ID = "0"
    
    cmbProcessEx.AddItem "<All Processes>", "0"
    cmbProcessEx.AddVals con, "Description", "Processes", "ProcessID", " ORDER BY SNo"
    cmbProcessEx.ID = "0"
    
    cmbMaker.AddItem "<All Makers>", "0"
    cmbMaker.AddVals con, "VendID1+' '+VenderName", "Makers", "VendID", " ORDER BY VendID"
    cmbMaker.ID = "0"
    
    cmbItem.DropDownOffLine = True
    cmbItem.AddItem "<All Items>", "0"
    cmbItem.AddVals con, "ItemID+' '+ItemName", "VItems", "ItemID", " ORDER BY ItemID"
    cmbItem.ID = "0"
        
    cmbHubs.AddItem "<All>"
    Call AddToCombo(cmbHubs, "DISTINCT ProcessGroupsProcesses.Hub_Name", "ProcessGroupsProcesses", " ORDER BY ProcessGroupsProcesses.Hub_Name", True)
    cmbHubs.ListIndex = 0
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Dim TheHeight As Long
    Frame1.Left = (Me.ScaleWidth - Frame1.Width) / 2
    
    
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    
End Sub

Private Sub RefreshLV(Optional p_lListEntryID As Long = 0)
    
    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim Cond As String
    Dim DTRange As String
    Dim ITM As ListItem
    
    Screen.MousePointer = vbHourglass
        
    lListEntryID = p_lListEntryID
    
    Dim strCondition As String
    strCondition = getCondition()
    
'    Dim sTmp As String
'    If cmbStatus.ListIndex = 0 Then
'        sTmp = " TotalShippedQty<TotalOrderQty"
'    ElseIf cmbStatus.ListIndex = 1 Then
'        sTmp = " TotalShippedQty=TotalOrderQty"
'    ElseIf cmbStatus.ListIndex = 2 Then
'        sTmp = ""
'    End If
    
    SQLStringForR1R2 = "SELECT DISTINCT ItemCode FROM VFOrderItems " & strCondition
    
    SQLString = "SELECT * FROM VRunningLots_POI WITH (NoLock) " & strCondition
        
    strReportSQL = "SELECT VFOrderItems.*,FCustomerCatalog.FOB,RunningPlanned,POPlanned,SemiPlanned,FinishPlanned,ForgingPlanned FROM VFOrderItems LEFT OUTER JOIN FCustomerCatalog ON VFOrderItems.CustCode=FCustomerCatalog.CustCode AND VFOrderItems.Country=FCustomerCatalog.Country AND VFOrderItems.CompItemCode=FCustomerCatalog.CompItemID LEFT OUTER JOIN FOrderItems_DeliveryRevisions ON VFOrderItems.ID=FOrderItems_DeliveryRevisions.FOI_RefID LEFT OUTER JOIN FOrderItems_InitialPlanningMemo ON VFOrderItems.ID=FOrderItems_InitialPlanningMemo.FOI_RefID " & strCondition & " ORDER BY VFOrderItems.DeliveryDT DESC"
    
    Dim strCondition_RL As String
    strCondition_RL = " WHERE 1=1"
    If chkReceived.Value = vbUnchecked Then
        strCondition_RL = strCondition_RL & " AND NOT (VRD_EntryID>0)"
    End If
    If chkIssued.Value = vbUnchecked Then
        strCondition_RL = strCondition_RL & " AND NOT (VRD_EntryID=0)"
    End If
    
    If cmbProcessEx.ID <> "0" Then
        strCondition_RL = strCondition_RL & " AND (Lot_Available_On_ProcessID=TPP.Previous_ProcessID)"
    End If
    
    con_ServerSide.Execute "DELETE FROM RunningLots_Temp"
    con_ServerSide.Execute "INSERT INTO RunningLots_Temp(ItemCode,LotNo,OrderNo,Description,Qty,ProcessID,VendID,ReWorkLot,R2InDT,PolishingItemSNO,SNo,DT,MaxSNo,TrayNo) " & _
     "SELECT ItemCode,LotNo,OrderNo,Description,Qty,ProcessID,VendID,ReWorkLot,R2InDT,PolishingItemSNO,SNo,DT,MaxSNo,TrayNo FROM VRunningLots_Simple_WPP VRL WITH (NoLock) " & _
     "LEFT JOIN VendRcvdDetail_QC_Report ON VRL.VRD_EntryID=VendRcvdDetail_QC_Report.VRD_RefID " & _
     "LEFT JOIN dbo.GetItem_Previous_Next_ProcessID_TblFn(" & cmbProcessEx.ID & ") TPP ON VRL.ItemCode=TPP.ItemID " & strCondition_RL
    
    Dim sStatus As String
    Dim iRecNo As Integer
    Dim dTotalAmt As Double, dTotalAmtPKR As Double
    Dim strRemarks As String
    dTotalAmt = 0
    With rs
        
        .Open SQLString, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        iRecNo = .RecordCount
        'LV.Sorted = False
        Do Until .EOF
            
            Set ITM = LV.ListItems.add(, , iRecNo - .AbsolutePosition + 1)
            ITM.Tag = !LotNo & ""
            
            ITM.ListSubItems.add , , Format(![StartDT] & "", "dd-MMM-yyyy")
            ITM.ListSubItems(1).Tag = Val(!EntryType & "")
            
            ITM.ListSubItems.add , , !LotNo & ""
            
            ITM.ListSubItems.add , , ![ItemCode] & ""
            ITM.ListSubItems.add , , ![ItemName] & ""
            ITM.ListSubItems.add , , Val(!Qty & "")
            ITM.ListSubItems.add , , !Description & ""
            ITM.ListSubItems.add , , Val(!LastProcessDays & "")
            ITM.ListSubItems.add , , IIf(!POIRemarks & "" = "", !Remarks & "", !POIRemarks & "")
            ITM.ListSubItems.add , , Val(!TLRunning & "") & "/" & Val(!TLUnShippedQty & "")
            ITM.ListSubItems.add , , Val(!R1 & "") + Val(!R2 & "") 'Val(!R1 & "") & "/" & Val(!R2 & "")
            ITM.ListSubItems.add , , Val(!PackStock & "")
            ITM.ListSubItems.add , , !OrderNo & ""
            If !UnShipped_OrderNo & "" = "" Then
                strRemarks = "Shipped"
            Else
                strRemarks = "Un-Shipped"
            End If
            If Val(!OpenPOQty & "") > 0 Then
                strRemarks = strRemarks & " Open POs Qty : " & Val(!OpenPOQty & "")
            End If
            ITM.ListSubItems.add(, , strRemarks).Tag = Val(!OpenPOQty & "")
            If LCase(!CustCode & "") = "stock" Then
                dRate = Val(!PriceForCost & "")
            Else
                dRate = Val(!Price & "")
            End If
            ITM.ListSubItems.add , , dRate
            ITM.ListSubItems.add , , !Hub_Name & ""
            .MoveNEXT
            
        Loop
        
        .Close
        
    End With
    
    Set rs = Nothing
    'txtTotalAmt = Format(dTotalAmtPKR, "###,##0.00")
    
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Private Function getCondition() As String

    Dim Cond As String
    
    Dim strDTFieldName As String
'    If chkUseOrderDates.Value = vbChecked Then
'        strDTFieldName = "VFOrderItems.DT"
'    Else
'        strDTFieldName = "VFOrderItems.DeliveryDT"
'    End If
    strDTFieldName = "DT"
'    If chkUrgent.Value = vbChecked Then
'        Cond = " WHERE ((VFOrderItems.DeliveryStatus=2 OR Remarks<>'') AND " & strDTFieldName & " BETWEEN '" & DtFrom.Value & "' AND '" & DtTo.Value & "')"
'    Else
'        If chkBoth.Value = vbChecked Then
'            'Both Means everyting in Between Date Range & those Items which have no Date Range.
'            Cond = " WHERE ((VFOrderItems.DeliveryStatus IN(0) OR VFOrderItems.DeliveryDT='1/1/1900') OR (" & strDTFieldName & " BETWEEN '" & DtFrom.Value & "' AND '" & DtTo.Value & "'))"
'            'Cond = " WHERE ((VFOrderItems.DeliveryStatus IN(0) OR VFOrderItems.DeliveryDT='1/1/1900') OR ((Remarks<>'' OR DeliveryStatus=2) AND (VFOrderItems.DeliveryDT BETWEEN '" & DtFrom.Value & "' AND '" & DtTo.Value & "')))"
'            'Cond = " WHERE (VFOrderItems.DeliveryStatus=0 OR VFOrderItems.DeliveryStatus=2)"
'        Else
'            If chkNoDelivery.Value = vbChecked Then
'                Cond = " WHERE (VFOrderItems.DeliveryStatus=0 OR VFOrderItems.DeliveryDT='1/1/1900')"
'            Else
'                Cond = " WHERE " & strDTFieldName & " BETWEEN '" & DtFrom.Value & "' AND '" & DtTo.Value & "'"
'            End If
'        End If
'    End If
    If cmbRange.ListIndex = 0 Then
        Cond = " WHERE 1=1"
    Else
        Cond = " WHERE " & strDTFieldName & " BETWEEN '" & DtFrom.Value & "' AND '" & DtTo.Value & "'"
    End If
    
    If cmbCust.ID <> "0" Then
        Cond = Cond & " AND (CustCode='" & cmbCust.ID & "') "
    End If
    
    If cmbCountry.ListIndex <> 0 Then
        Cond = Cond & " AND (Country='" & cmbCountry & "')"
    End If
    
    If cmbItemGroup.ID <> "0" Then
        Cond = Cond & " AND (GroupID=" & cmbItemGroup.ID & ")"
    End If
    
'    Cond = Cond & " AND (Qty>ShippedQty)"
'
'    Cond = Cond & " AND VFOrderItems.OrderNo IN (SELECT OrderNo FROM VUnshippedOrderList)"
    
'    Cond = Cond & " AND VFOrderItems.CustCode NOT IN('Stock')"
    
    If cmbOrderNo.MatchFound And cmbOrderNo.ID <> "0" Then
        Cond = Cond & " AND OrderNo='" & cmbOrderNo.ID & "'"
    End If
    
    If cmbProcess.ID <> "0" Then
        Cond = Cond & " AND ProcessID=" & Val(cmbProcess.ID)
    End If
    
    If cmbMaker.ID <> "0" Then
        Cond = Cond & " AND VendID=" & Val(cmbMaker.ID)
    End If
    
    If cmbItem.ID <> "0" Then
        Cond = Cond & " AND ItemCode='" & cmbItem.ID & "'"
    End If
    
    If chkRepairLots.Value = vbUnchecked Then
        Cond = Cond & " AND ReWorkLot=0"
    End If
    
    If chkRegularLots.Value = vbUnchecked Then
        Cond = Cond & " AND ReWorkLot=1"
    End If
    
    If Val(OptR1(0).Tag) < 2 Then
        Cond = Cond & " AND (EntryType=0 AND RunningLots_R1=" & Val(OptR1(0).Tag) & ")"
    End If
    
    If chkLotsOnly.Value = vbUnchecked Then
        Cond = Cond & " AND EntryType<>0"
    End If
    If chkNotYetStarted.Value = vbUnchecked Then
        Cond = Cond & " AND EntryType<>1"
    End If
    
    If cmbHubs.ListIndex > 0 Then
        Cond = Cond & " AND Hub_Name='" & cmbHubs.Text & "'"
    End If
    
    Cond = Cond & " AND Qty>0"
    If txtLotNo <> "" Then
        Cond = " WHERE LotNo='" & txtLotNo & "'"
    End If
        
    
    getCondition = Cond
    
    If lListEntryID <> 0 Then
        getCondition = " WHERE Qty>0 AND ((EntryType=0 AND LotNo IN(SELECT LotNo FROM POI_List_Detail WHERE POI_List_RefID=" & lListEntryID & ")) OR (EntryType=1 AND OrderNo+ItemCode IN(SELECT OrderNo+ItemCode FROM POI_List_Detail WHERE POI_List_RefID=" & lListEntryID & ")))"
    Else
        getCondition = Cond
    End If
    
End Function

Private Function getConditionitionForReport() As String

    Dim strstrConditionition As String
        
    If chkUseOrderDates.Value = vbChecked Then
        strDTFieldName = "{VFOrderItems.DT}"
    Else
        strDTFieldName = "{VFOrderItems.DeliveryDT}"
    End If
    
    If chkUrgent.Value = vbChecked Then
        strCondition = " (({VFOrderItems.DeliveryStatus}=2 OR {VFOrderItems.Remarks}<>'') AND " & strdtfilename & "=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#)"
        'Cond = "((VFOrderItems.DeliveryStatus=2 OR Remarks<>'') AND " & strDTFieldName & " BETWEEN '" & DtFrom.Value & "' AND '" & DtTo.Value & "')"
    Else
        If chkBoth.Value = vbChecked Then
            'strCondition = " ({VFOrderItems.DeliveryStatus} IN [0,2] OR ({VFOrderItems.DeliveryDT}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#))"
            strCondition = " (({VFOrderItems.DeliveryStatus} IN [0] OR {VFOrderItems.DeliveryDT}=#1/1/1900#) OR (" & strDTFieldName & "=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#))"
            'Cond = " WHERE ((VFOrderItems.DeliveryStatus IN(0) OR VFOrderItems.DeliveryDT='1/1/1900') OR (" & strDTFieldName & " BETWEEN '" & DtFrom.Value & "' AND '" & DtTo.Value & "'))"
        Else
            If chkNoDelivery.Value = vbChecked Then
                'strCondition = " {VFOrderItems.DeliveryStatus}=0"
                strCondition = " ({VFOrderItems.DeliveryStatus}=0 OR {VFOrderItems.DeliveryDT}=#1/1/1900#)"
                'Cond = " WHERE (VFOrderItems.DeliveryStatus=0 OR VFOrderItems.DeliveryDT='1/1/1900')"
            Else
                strCondition = strDTFieldName & "=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#"
            End If
        End If
    End If
    
    If cmbCust.ID <> "0" Then
        strCondition = strCondition & " AND ({VFOrderItems.CustCode}='" & cmbCust.ID & "') "
    End If
    
    If cmbCountry.ListIndex <> 0 Then
        strCondition = strCondition & " AND ({VFOrderItems.Country}='" & cmbCountry & "')"
    End If
    
    If cmbItemGroup.ID <> "0" Then
        strCondition = strCondition & " AND ({VFOrderItems.GroupID}=" & cmbItemGroup.ID & ")"
    End If
     
    If lListEntryID <> 0 Then
        Dim i As Integer, strIDs As String
        For i = 1 To LV.ListItems.count
            strIDs = strIDs & Val(LV.ListItems(i).key) & ","
        Next
        
        If strIDs <> "" Then
            strIDs = Left(strIDs, Len(strIDs) - 1)
        End If
        
        strCondition = "{VFOrderItems.ID} IN[" & strIDs & "]"
    End If
    
    If cmbOrderNo.ID <> "0" Then
        strCondition = strCondition & " AND ({VFOrderItems.OrderNo}='" & cmbOrderNo.ID & "')"
    End If
    
    strCondition = strCondition & " AND ({VFOrderItems.Qty}>{VFOrderItems.ShippedQty})"
    
    
    
    'strCondition = strCondition & " AND {VFOrderItems.OrderNo} IN (SELECT OrderNo FROM VUnshippedOrderList)"
    
    getConditionitionForReport = strCondition
    
End Function


Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If LV.SortKey + 1 = ColumnHeader.Index Then
        LV.ColumnHeaderIcons = Nothing
        'Set LV.ColumnHeaderIcons = ImageList2
        If LV.SortOrder = lvwAscending Then
            LV.SortOrder = lvwDescending
            'ColumnHeader.Icon = "Down"
        Else
            LV.SortOrder = lvwAscending
            'ColumnHeader.Icon = "Up"
        End If
    Else
        LV.ColumnHeaderIcons = Nothing
        'Set LV.ColumnHeaderIcons = ImageList2
        LV.SortKey = ColumnHeader.Index - 1
        LV.SortOrder = lvwAscending
        'ColumnHeader.Icon = "Down"
    End If
    Call RefreshSnos
    If LV.Sorted = False Then LV.Sorted = True
    
End Sub

Private Sub LV_DblClick()
    
    If LV.ListItems.count = 0 Then Exit Sub
     
    'If Not (iColNo = 14 Or iColNo = 11) Then iColNo = 8
    iColNo = 9
    
    If iColNo = 9 Then
        With txtEdit
            .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(iColNo).Width
            .Text = LV.SelectedItem.SubItems(iColNo - 1)
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    End If
    Exit Sub
    If iColNo = 8 Then
        If bAuthorized = False Then
            If getDBPassword("EditDeliveryDT") = False Then
                Exit Sub
            Else
                bAuthorized = True
            End If
        End If
        
        If Val(LV.SelectedItem.ListSubItems(7).Tag) = 0 Then 'No Delivery
            DTDelivery.Value = Date
        Else
            DTDelivery.Value = LV.SelectedItem.ListSubItems(7).Text
        End If
    ElseIf iColNo = 14 Then
        If LV.SelectedItem.SubItems(iColNo - 1) = "" Then
            DTDelivery.Value = Date
        Else
            DTDelivery.Value = LV.SelectedItem.SubItems(iColNo - 1)
        End If
    ElseIf iColNo = 11 Then             'Remarks
        With txtEdit
            .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(iColNo).Width
            .Text = LV.SelectedItem.SubItems(iColNo - 1)
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
        Exit Sub
    End If
    
    
    If iColNo = 11 Then
        
    Else
        With DTDelivery
            .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top
            .Visible = True
            .SetFocus
        End With
    End If
    
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        With LV.ColumnHeaders(i)
            If x > Val(.Left) And x < (Val(.Left) + Val(.Width)) Then
                iColNo = i
                Exit For
            End If
        End With
    Next
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If LV.ListItems.count = 0 Or Button <> 2 Then Exit Sub
    Dim strRet As String
    strRet = ShowPopUpMenu1("Print PTC")
    If strRet = "Print PTC" Then
        Call PrintPTC(LV.SelectedItem.Tag)
    End If
    
End Sub

Private Sub PrintOrderList()

    On Error GoTo err
    Me.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptFOrderList.rpt")
        
    rpt.FormulaFields(3).Text = "'" & cmbCust.Text & " " & cmbCountry.Text & "'"
    
    frmPrevRpt.ShowReport " ", rpt, SQLString
    
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub PrintOrderBalanceReportOld()

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

Private Sub mnuedit_Click()
    Call LV_DblClick
End Sub

Private Sub PrintOrder(Optional bPrintBalance As Boolean = False, Optional bMakerWise As Boolean = False)

    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    Dim strSelection As String
    Dim rpt As CRAXDDRT.Report
    If Not bPrintBalance Then
        If bMakerWise Then
            Set rpt = rptApp.OpenReport(RptPath & "\rptProductionOrderMakerwise.rpt") '
        Else
            Set rpt = rptApp.OpenReport(RptPath & "\rptProductionOrder.rpt") '
        End If
        rpt.FormulaFields.GetItemByName("ProductionOrder").Text = bForProduction
        strSelection = "{CustOrders.OrderNo}='" & Left(LV.SelectedItem.key & "'", Len(LV.SelectedItem.key) - 1) & "' "
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\rptProductionOrderBalanceQty.rpt") '
        rpt.FormulaFields.GetItemByName("ProductionOrder").Text = bForProduction
        strSelection = "{CustOrders.OrderNo}='" & Left(LV.SelectedItem.key & "'", Len(LV.SelectedItem.key) - 1) & "' AND {@Balance}>0"
    End If
    rpt.FormulaFields.GetItemByName("UserName").Text = "'" & CurrentUserName & "'"
    Load frmPrevRpt
    frmPrevRpt.ShowReport strSelection, rpt
    
    Screen.MousePointer = vbDefault

    
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

Private Sub OptR1_Click(Index As Integer)
    OptR1(0).Tag = Index
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Dim iRecAffect As Integer
        If Val(LV.SelectedItem.ListSubItems(1).Tag) = 0 Then    'Lot
            con.Execute "UPDATE POI_Remarks SET Remarks='" & txtEdit & "' WHERE LotNo='" & LV.SelectedItem.Tag & "'", iRecAffect
            If iRecAffect = 0 Then
                con.Execute "INSERT INTO POI_Remarks(EntryType,LotNo,Remarks) VALUES(0,'" & LV.SelectedItem.Tag & "','" & txtEdit & "')"
            End If
        Else
            con.Execute "UPDATE POI_Remarks SET Remarks='" & txtEdit & "' WHERE ItemCode='" & LV.SelectedItem.SubItems(3) & "' AND OrderNo='" & LV.SelectedItem.SubItems(12) & "'", iRecAffect
            If iRecAffect = 0 Then
                con.Execute "INSERT INTO POI_Remarks(EntryType,ItemCode,OrderNo,Remarks) VALUES(1,'" & LV.SelectedItem.SubItems(3) & "','" & LV.SelectedItem.SubItems(12) & "','" & txtEdit & "')"
            End If
        End If
        
        LV.SelectedItem.SubItems(iColNo - 1) = txtEdit
        If LV.SelectedItem.Index < LV.ListItems.count Then
            LV.ListItems(LV.SelectedItem.Index + 1).Selected = True
            Call LV_DblClick
            Exit Sub
        End If
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
    
End Sub

Private Sub txtEdit_LostFocus()
'    txtEdit.Visible = False
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

