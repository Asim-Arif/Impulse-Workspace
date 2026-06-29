VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmTrafficDashBoard 
   ClientHeight    =   10335
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
   ScaleHeight     =   10335
   ScaleMode       =   0  'User
   ScaleWidth      =   12884.84
   WindowState     =   2  'Maximized
   Begin VB.Frame FraDetail 
      Height          =   6870
      Left            =   90
      TabIndex        =   29
      Top             =   2265
      Visible         =   0   'False
      Width           =   14175
      Begin VB.CommandButton cmdHide 
         Caption         =   "&Hide"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   12915
         TabIndex        =   30
         Top             =   6285
         Width           =   1140
      End
      Begin MSComctlLib.ListView LVOpenPOs 
         Height          =   2655
         Left            =   105
         TabIndex        =   31
         Top             =   3540
         Width           =   13950
         _ExtentX        =   24606
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Master PO No."
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Qty"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Maker"
            Object.Width           =   4233
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Ret. Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView LVLots 
         Height          =   2655
         Left            =   105
         TabIndex        =   32
         Top             =   570
         Width           =   13950
         _ExtentX        =   24606
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Lot #"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Qty"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Process"
            Object.Width           =   4233
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Maker"
            Object.Width           =   4233
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Open POs :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   135
         TabIndex        =   34
         Top             =   3255
         Width           =   1125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Running Lots :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   135
         TabIndex        =   33
         Top             =   285
         Width           =   1455
      End
   End
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
      TabIndex        =   4
      Top             =   0
      Width           =   14325
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Traffic Dashboard"
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
         TabIndex        =   5
         Top             =   165
         Width           =   14175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Traffic Dashboard"
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
         TabIndex        =   6
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
      Height          =   9675
      Left            =   15
      TabIndex        =   0
      Top             =   615
      Width           =   14325
      Begin MSComctlLib.ListView LVWhite 
         Height          =   1800
         Left            =   45
         TabIndex        =   27
         Top             =   7260
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   3175
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Order #"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Item Group"
            Object.Width           =   3572
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Item Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Item Name"
            Object.Width           =   11024
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Qty"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Criteria"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.OptionButton Opts 
         Caption         =   "R2"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   2
         Left            =   7155
         TabIndex        =   26
         Top             =   9255
         Width           =   870
      End
      Begin VB.OptionButton Opts 
         Caption         =   "R1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   1
         Left            =   6255
         TabIndex        =   25
         Top             =   9255
         Width           =   870
      End
      Begin VB.OptionButton Opts 
         Caption         =   "Both"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   0
         Left            =   5265
         TabIndex        =   24
         Tag             =   "0"
         Top             =   9255
         Value           =   -1  'True
         Width           =   870
      End
      Begin ComboList.Usercontrol1 cmbOrderNo 
         Height          =   285
         Left            =   10260
         TabIndex        =   23
         Top             =   330
         Width           =   2160
         _ExtentX        =   3810
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
         Left            =   9195
         TabIndex        =   22
         Top             =   9075
         Visible         =   0   'False
         Width           =   1515
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
         Left            =   10110
         TabIndex        =   21
         Top             =   9330
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   11700
         TabIndex        =   20
         Top             =   9165
         Width           =   1350
      End
      Begin ComboList.Usercontrol1 cmbItemGroup 
         Height          =   285
         Left            =   7965
         TabIndex        =   13
         Top             =   330
         Width           =   2295
         _ExtentX        =   4048
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
         Left            =   12450
         TabIndex        =   12
         Top             =   375
         Width           =   810
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
         ItemData        =   "frmTrafficDashBoard.frx":0000
         Left            =   3135
         List            =   "frmTrafficDashBoard.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   330
         Width           =   1725
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   6420
         TabIndex        =   10
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
         Format          =   198311939
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   4860
         TabIndex        =   11
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
         Format          =   198311939
         CurrentDate     =   37055
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   45
         TabIndex        =   2
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
      Begin MSComctlLib.ListView LVRed 
         Height          =   1980
         Left            =   45
         TabIndex        =   14
         Top             =   705
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   3493
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   16777215
         BackColor       =   152
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Order #"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Item Group"
            Object.Width           =   3572
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Item Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Item Name"
            Object.Width           =   11024
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Qty"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Criteria"
            Object.Width           =   1411
         EndProperty
      End
      Begin MSComctlLib.ListView LVYellow 
         Height          =   1980
         Left            =   45
         TabIndex        =   15
         Top             =   2850
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   3493
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   1885936
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Order #"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Item Group"
            Object.Width           =   3572
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Item Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Item Name"
            Object.Width           =   11024
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Qty"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Criteria"
            Object.Width           =   1411
         EndProperty
      End
      Begin MSComctlLib.ListView LVGreen 
         Height          =   2025
         Left            =   45
         TabIndex        =   16
         Top             =   5010
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   3572
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   16777215
         BackColor       =   2658830
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Order #"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Item Group"
            Object.Width           =   3572
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Item Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Item Name"
            Object.Width           =   11024
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Qty"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Criteria"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "1-2=Maker PO is Delayed by 2 or More Days... | 2-3=Maker PO's Delivery Date is Yet to Come"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000098&
         Height          =   285
         Index           =   3
         Left            =   45
         TabIndex        =   28
         Top             =   9060
         Width           =   14175
      End
      Begin VB.Label Label2 
         Caption         =   "3= All of the Rest"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0028920E&
         Height          =   285
         Index           =   2
         Left            =   45
         TabIndex        =   19
         Top             =   7050
         Width           =   11340
      End
      Begin VB.Label Label2 
         Caption         =   "2-1 Lot is stuck for 2-4 days | 2-2=Repair Not Issued"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   285
         Index           =   1
         Left            =   45
         TabIndex        =   18
         Top             =   4815
         Width           =   14175
      End
      Begin VB.Label Label2 
         Caption         =   $"frmTrafficDashBoard.frx":004A
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000098&
         Height          =   285
         Index           =   0
         Left            =   45
         TabIndex        =   17
         Top             =   2670
         Width           =   14175
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   360
         Left            =   13305
         TabIndex        =   8
         Top             =   240
         Width           =   915
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "1614;635"
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
         Height          =   435
         Left            =   13080
         TabIndex        =   7
         Top             =   9165
         Width           =   1140
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "2011;767"
         Accelerator     =   67
         FontName        =   "Calibri"
         FontEffects     =   1073741825
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   1485
         TabIndex        =   3
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
         Caption         =   $"frmTrafficDashBoard.frx":00D6
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
         TabIndex        =   1
         Top             =   120
         Width           =   12375
      End
   End
End
Attribute VB_Name = "frmTrafficDashBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQLString As String, SQLStringForR1R2 As String
Dim bForProduction As Boolean
Dim strReportSQL As String
Dim bAuthorized As Boolean

Private Sub mnuProdPlan_Click()
    Load frmProdPlan
    Call frmProdPlan.AddPlanning(Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1))
End Sub


Private Sub cmbCust_matched()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    SQL = "Select Distinct Country From ForeignCustomers "
    
    If cmbCust.ID <> "0" Then
        If Left(cmbCust.ID, 1) <> "'" Then
            SQL = SQL & " WHERE CustCode IN('" & cmbCust.ID & "')"
        Else
            SQL = SQL & " WHERE CustCode IN(" & cmbCust.ID & ")"
        End If
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
    
    cmbOrderNo.ClearVals
    cmbOrderNo.AddItem "<All Orders>", "0"
    If cmbCust.ID = "0" Then
        cmbOrderNo.AddVals con, "OrderNo + ' {' + InternalRefNo + '}'", "FCustomerOrders", "OrderNo"
    Else
        Dim strCondition As String
        If Left(cmbCust.ID, 1) <> "'" Then
            strCondition = " WHERE CustCode IN('" & cmbCust.ID & "')"
        Else
            strCondition = " WHERE CustCode IN(" & cmbCust.ID & ")"
        End If
        
        
        cmbOrderNo.AddVals con, "OrderNo + ' {' + InternalRefNo + '}'", "FCustomerOrders", "OrderNo", strCondition
        
    End If
    cmbOrderNo.ID = "0"
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbRange_Click()
    
    DtTo = ServerDate
    DtFrom = ServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    
    Select Case cmbRange.ListIndex
        Case 0
            
            DtTo = DateAdd("d", 15, ServerDate)
        Case 1
            DtTo = DateAdd("d", 30, ServerDate)
        Case 2
            DtTo = DateAdd("d", 45, ServerDate)
        Case 3
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
       
        fldValue = LV.ListItems(i).SubItems(cmbfield.ListIndex + 1)
       
        If LCase(fldValue) Like "*" & LCase(Trim(txtFind)) & "*" Then
            LV.ListItems(i).Selected = True
            LV.ListItems(i).EnsureVisible
            If i <> StartIdx Then Exit For
        End If
       
    Next
    Dim rpt As New CRAXDDRT.Report
    
    If i = LV.ListItems.count + 1 Then
        MsgBox "Record Not Found.", vbInformation
        Set LV.SelectedItem = Nothing
    End If

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
    rpt.EnableParameterPrompting = False
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, , False
    
End Sub

Private Sub cmdPrint_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\TrafficDashBoard.rpt")
    
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DtFrom.Value
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DtTo.Value
    
    rpt.ParameterFields.GetItemByName("@CustCode").SetCurrentValue IIf(cmbCust.ID = "0", "", cmbCust.ID)
    rpt.ParameterFields.GetItemByName("@Country").SetCurrentValue IIf(cmbCountry.ListIndex = 0, "", cmbCountry.Text)
    rpt.ParameterFields.GetItemByName("@OrderNo").SetCurrentValue IIf(cmbOrderNo.ID = "0", "", cmbOrderNo.ID)
    rpt.ParameterFields.GetItemByName("@GroupID").SetCurrentValue CLng(IIf(cmbItemGroup.ID = "0", 0, cmbItemGroup.ID))
    
    rpt.ParameterFields.GetItemByName("@NoDelivery").SetCurrentValue CBool(chkNoDelivery.Value)
    rpt.ParameterFields.GetItemByName("@Urgent").SetCurrentValue CBool(chkUrgent.Value)
    rpt.ParameterFields.GetItemByName("@Both").SetCurrentValue CBool(chkBoth.Value)
    rpt.ParameterFields.GetItemByName("@R1R2").SetCurrentValue Val(Opts(0).Tag)
    
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, "", False
    
End Sub

Private Sub PrintList(Optional bWithValue As Boolean = False)

    Dim rpt As CRAXDDRT.Report
    If bWithValue Then
        Set rpt = rptApp.OpenReport(RptPath & "\OrderItemListWithValue.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\OrderItemList.rpt")
    End If
    
    Dim strFilters As String
    
    strFilters = Format(DtFrom.Value, "dd-MMM-yyyy") & " to " & Format(DtTo.Value, "dd-MMM-yyyy")
    
    strFilters = strFilters & " | Customer:" & cmbCust.Text
    strFilters = strFilters & " | Country:" & cmbCountry.Text
     
    If chkUrgent.Value = vbChecked Then
         strFilters = strFilters & " | Urgent"
    Else
        If chkNoDelivery.Value = vbChecked Then
            strFilters = strFilters & " | No Delivery Date"
        End If
        
        If chkBoth.Value = vbChecked Then
            strFilters = strFilters & " | Both"
        End If
    End If
    rpt.FormulaFields.GetItemByName("Filters").Text = "'" & strFilters & "'"
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, strReportSQL
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
    
    Dim i As Integer, strIDs As String
    For i = 1 To LV.ListItems.count
        strIDs = strIDs & Val(LV.ListItems(i).key) & ","
    Next
    
    If strIDs <> "" Then
        strIDs = Left(strIDs, Len(strIDs) - 1)
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\OrderStatusReport.rpt")
    rpt.ParameterFields.GetItemByName("@FOI_IDs").SetCurrentValue strIDs
    
    rpt.EnableParameterPrompting = False
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, "", False
    
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
        .Parameters("@OrderNo").Value = "0"
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
    rpt.FormulaFields.GetItemByName("OrderNo").Text = "'0'"
    Dim f As New frmPrevRpt
    Dim strConditionReport As String
    strConditionReport = getConditionitionForReport
    f.ShowReport strConditionReport, rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdPrintR1_Click()
    Call PrintR1R2PlanningReport(True)
End Sub

Private Sub cmdPrintR2_Click()
    Call PrintR1R2PlanningReport(False)
End Sub

Private Sub cmdHide_Click()
    FraDetail.Visible = False
End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub


Private Sub DTDelivery_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
    
        Dim lID As Long, iStatus As Integer
        lID = Val(LV.SelectedItem.key)
        iStatus = Val(LV.SelectedItem.ListSubItems(7).Tag)
        Call StartTrans(con)
        
        con.Execute "UPDATE FOrderItems SET DeliveryStatus=1,DeliveryDT='" & DTDelivery.Value & "' WHERE ID=" & lID
        con.Execute "INSERT INTO FOrderItems_DeliveryRevisions(FOI_RefID,DeliveryStatus,DeliveryDT,UserName,MachineName)" & _
         "VALUES(" & lID & "," & iStatus & ",'" & DTDelivery.Value & "','" & CurrentUserName & "','" & strComputerName & _
         "')"
        
        con.CommitTrans
        
        DTDelivery.Visible = False
        Call RefreshLV
        
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
    
    cmbCust.Functionality = 1
    cmbCust.ListHeight = 2500
    cmbCust.AddItem "<All Customers>", "0"
    'Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where custCode<>'Hico'")
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode")
    
    cmbCust.ID = "0"
    cmbCountry.ListIndex = 0
    
    cmbItemGroup.AddItem "<All Groups>", "0"
    cmbItemGroup.AddVals con, "Description", "ItemGroups", "ID", " ORDER BY Description"
    cmbItemGroup.ID = "0"
   
    
    DtFrom.Value = Date
    DtTo.Value = Date
    cmbRange.ListIndex = 0
    
    cmbOrderNo.ListHeight = 3000
    
    Dim i As Integer
    If strCompany <> "Dr-Frgz" Then
        For i = 0 To 2
            Opts(i).Visible = False
        Next
    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Dim TheHeight As Long
    Frame1.Left = (Me.ScaleWidth - Frame1.Width) / 2
    
    
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    
End Sub


Private Sub RefreshLV()
    On Error GoTo err
    
    Dim cmd As New ADODB.Command, rs As New ADODB.Recordset
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "TrafficDashBoard_SP"
        .CommandTimeout = 120
        .Parameters("@DTFrom").Value = DtFrom.Value
        .Parameters("@DTTo").Value = DtTo.Value
        .Parameters("@CustCode").Value = IIf(cmbCust.ID = "0", "", cmbCust.ID)
        .Parameters("@Country").Value = IIf(cmbCountry.ListIndex = 0, "", cmbCountry.Text)
        .Parameters("@OrderNo").Value = IIf(cmbOrderNo.ID = "0", "", cmbOrderNo.ID)
        .Parameters("@GroupID").Value = IIf(cmbItemGroup.ID = "0", 0, cmbItemGroup.ID)
        .Parameters("@NoDelivery").Value = Abs(chkNoDelivery.Value)
        .Parameters("@Both").Value = Abs(chkBoth.Value)
        .Parameters("@Urgent").Value = Abs(chkUrgent.Value)
        .Parameters("@R1R2").Value = Val(Opts(0).Tag)
        Set rs = .Execute
    End With
    
    Dim myLV As ListView, ITM As ListItem
    With rs
        LVRed.ListItems.Clear
        LVYellow.ListItems.Clear
        LVGreen.ListItems.Clear
        Do Until .EOF
            If !Color & "" = "Red" Then
                Set myLV = LVRed
            ElseIf !Color & "" = "Yellow" Then
                Set myLV = LVYellow
            ElseIf !Color & "" = "Green" Then
                Set myLV = LVGreen
            ElseIf !Color & "" = "White" Then
                Set myLV = LVWhite
            End If
            Set ITM = myLV.ListItems.add(, , myLV.ListItems.count + 1)
            ITM.ListSubItems.add , , !CustCode & ""
            ITM.ListSubItems.add , , IIf(strCompany = "IAA" Or strCompany = "Kami", !InternalRefNo & "", !OrderNo & "")
            ITM.ListSubItems.add , , !Description & ""
            ITM.ListSubItems.add , , !CompItemCode & ""
            ITM.ListSubItems.add , , !ItemName & ""
            ITM.ListSubItems.add , , Val(!Qty & "")
            ITM.ListSubItems.add , , !Criteria & ""
            
            .MoveNEXT
        Loop
    End With
    Set rs = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Private Function getCondition() As String

    Dim Cond As String
    
    If chkUrgent.Value = vbChecked Then
        Cond = " WHERE (VFOrderItems.DeliveryStatus=2 OR Remarks<>'')"
    Else
        If chkBoth.Value = vbChecked Then
            Cond = " WHERE (VFOrderItems.DeliveryStatus=0 OR (VFOrderItems.DeliveryDT BETWEEN '" & DtFrom.Value & "' AND '" & DtTo.Value & "'))"
        Else
            If chkNoDelivery.Value = vbChecked Then
                Cond = " WHERE VFOrderItems.DeliveryStatus=0"
            Else
                Cond = " WHERE VFOrderItems.DeliveryDT BETWEEN '" & DtFrom.Value & "' AND '" & DtTo.Value & "'"
            End If
        End If
    End If
    
    If cmbCust.ID <> "0" Then
        Cond = Cond & " AND (VFOrderItems.CustCode='" & cmbCust.ID & "') "
    End If
    
    If cmbCountry.ListIndex <> 0 Then
        Cond = Cond & " AND (VFOrderItems.Country='" & cmbCountry & "')"
    End If
    
    If cmbItemGroup.ID <> "0" Then
        Cond = Cond & " AND (GroupID=" & cmbItemGroup.ID & ")"
    End If
    
    Cond = Cond & " AND (Qty>ShippedQty)"
    
    Cond = Cond & " AND VFOrderItems.OrderNo IN (SELECT OrderNo FROM VUnshippedOrderList)"
    
    Cond = Cond & " AND VFOrderItems.CustCode NOT IN('Stock')"
    
    getCondition = Cond
    
End Function

Private Function getConditionitionForReport() As String

    Dim strstrConditionition As String
    
    If chkUrgent.Value = vbChecked Then
        strCondition = " {VFOrderItems.DeliveryStatus}=2 OR {VFOrderItems.Remarks}<>'')"
    Else
        If chkBoth.Value = vbChecked Then
            strCondition = " ({VFOrderItems.DeliveryStatus}=0 OR ({VFOrderItems.DeliveryDT}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#))"
        Else
            If chkNoDelivery.Value = vbChecked Then
                strCondition = " {VFOrderItems.DeliveryStatus}=0"
            Else
                strCondition = " {VFOrderItems.DeliveryDT}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#"
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
        strCondition = strCondition & " AND (VFOrderItems.GroupID}=" & cmbItemGroup.ID & ")"
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
    
    With DTDelivery
        .Move LV.Left + LV.ColumnHeaders(8).Left, LV.Top + LV.SelectedItem.Top
        .Visible = True
        .SetFocus
    End With
    
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

Private Sub LVGreen_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    Dim myLV As ListView
    Set myLV = LVGreen
    
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

Private Sub LVGreen_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If LVGreen.ListItems.count = 0 Then Exit Sub
        
    '''''''''''''''''''
    Dim strRet As String
    strRet = ShowPopUpMenu1("Show Planning Dashboard")
    
    If strRet = "Show Planning Dashboard" Then
        Call ShowProductionPlanningDashBoard(LVGreen.SelectedItem.SubItems(4))
    End If
End Sub

Private Sub LVLots_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Dim myLV As ListView
    Set myLV = LVLots
    
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

Private Sub LVLots_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button <> 2 Then Exit Sub
    If LVLots.ListItems.count = 0 Then Exit Sub
    
    Dim strRet As String
    strRet = ShowPopUpMenu1("Print PTC")
    
    If strRet = "Print PTC" Then
        Call PrintPTC(LVLots.SelectedItem.SubItems(1), False)
    End If
    
End Sub

Private Sub LVOpenPOs_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Dim myLV As ListView
    Set myLV = LVOpenPOs
    
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

Private Sub LVOpenPOs_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button <> 2 Then Exit Sub
    If LVOpenPOs.ListItems.count = 0 Then Exit Sub
    
    Dim strRet As String
    strRet = ShowPopUpMenu1("Print Master PO Status")
    
    If strRet = "Print Master PO Status" Then
        Dim rpt As CRAXDDRT.Report
        Set rpt = rptApp.OpenReport(RptPath & "\MasterPOStatus.rpt")
        
        Dim f As New frmPrevRpt
        f.ShowReport "{VendIssued.MasterPONo}='" & LVOpenPOs.SelectedItem.SubItems(1) & "'", rpt
    End If
    
End Sub

Private Sub LVRed_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Dim myLV As ListView
    Set myLV = LVRed
    
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

Private Sub LVRed_DblClick()
    
    If LVRed.ListItems.count = 0 Then Exit Sub
    Call ShowDetails(LVRed, "Red")
        
End Sub

Private Sub ShowDetails(LV As ListView, strLight As String)

    Dim strOrderNo As String, strItemCode As String, strCriteria As String
    
    strOrderNo = LV.SelectedItem.SubItems(2)
    strItemCode = LV.SelectedItem.SubItems(4)
    strCriteria = LV.SelectedItem.SubItems(7)
     
    Dim strConditionRL As String, strConditionOP As String, strConditionRepair As String
    If strLight = "Red" Then
        If strCriteria = "1-1" Then
            strConditionRL = "WHERE ItemCode='" & strItemCode & "' AND OrderNo='" & strOrderNo & "'"
            strConditionOP = "WHERE ItemCode='" & strItemCode & "' AND OrderNo='" & strOrderNo & "'"
            strConditionRepair = ""
        ElseIf strCriteria = "1-2" Then
            strConditionRL = ""
            strConditionOP = "WHERE ItemCode='" & strItemCode & "' AND OrderNo='" & strOrderNo & "' AND DATEADD(day,2,ReturnDT)<GETDATE()"
            strConditionRepair = ""
        ElseIf strCriteria = "1-3" Then
            strConditionRL = "WHERE ItemCode='" & strItemCode & "' AND OrderNo='" & strOrderNo & "' AND DateAdd(Day, 5, DT) < getDate()"
            strConditionOP = ""
            strConditionRepair = ""
        End If
    ElseIf strLight = "Yellow" Then
        If strCriteria = "2-1" Then
            strConditionRL = "WHERE ItemCode='" & strItemCode & "' AND OrderNo='" & strOrderNo & "' AND DATEDIFF(day,DT,GetDate()) BETWEEN 2 AND 4"
            strConditionOP = ""
            strConditionRepair = ""
        ElseIf strCriteria = "2-2" Then
            strConditionRL = ""
            strConditionOP = ""
            strConditionRepair = "WHERE OrderNo='" & strOrderNo & "' AND ItemCode='" & strItemCode & "' AND Qty<IssQty"
        ElseIf strCriteria = "2-3" Then
            strConditionRL = ""
            strConditionOP = "WHERE OrderNo='" & strOrderNo & "' AND ItemCode='" & strItemCode & "' AND ReturnDT>GETDATE()"
            strConditionRepair = ""
        End If
    End If
    
    If Opts(2).Value Then
        If strConditionRL <> "" Then
            strConditionRL = strConditionRL & " AND SNo>PolishingItemSNo"
        End If
    End If
    
    LVOpenPOs.ListItems.Clear
    LVLots.ListItems.Clear
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        If strConditionRL <> "" Then
            .Open "SELECT LotNo,Qty,Description,Makers.VenderName,DT FROM VRunningLots INNER JOIN Makers ON VRunningLots.VendID=Makers.VendID " & strConditionRL, con, adOpenForwardOnly, adLockReadOnly
            
            Do Until .EOF
                Set ITM = LVLots.ListItems.add(, , LVLots.ListItems.count + 1)
                ITM.ListSubItems.add , , !LotNo & ""
                ITM.ListSubItems.add , , Val(!Qty & "")
                ITM.ListSubItems.add , , !Description & ""
                ITM.ListSubItems.add , , !VenderName & ""
                ITM.ListSubItems.add , , Format(!DT, "dd-MMM-yyyy")
                .MoveNEXT
            Loop
            .Close
        End If
        If strConditionOP <> "" Then
            .Open "SELECT * FROM VItemOpenPOs " & strConditionOP, con, adOpenForwardOnly, adLockReadOnly
            
            Do Until .EOF
                Set ITM = LVOpenPOs.ListItems.add(, , LVOpenPOs.ListItems.count + 1)
                ITM.ListSubItems.add , , !MasterPONo & ""
                ITM.ListSubItems.add , , Val(!Qty & "")
                ITM.ListSubItems.add , , !VenderName & ""
                ITM.ListSubItems.add , , Format(!ReturnDT, "dd-MMM-yyyy")
                .MoveNEXT
            Loop
            .Close
        End If
    End With
    Set rs = Nothing
    
    FraDetail.Visible = True
    
End Sub

Private Sub LVRed_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button <> 2 Then Exit Sub
    If LVRed.ListItems.count = 0 Then Exit Sub
        
    '''''''''''''''''''
    Dim strRet As String
    strRet = ShowPopUpMenu1("Show Planning Dashboard")
    
    If strRet = "Show Planning Dashboard" Then
        Call ShowProductionPlanningDashBoard(LVRed.SelectedItem.SubItems(4))
    End If
    
End Sub

Private Sub ShowProductionPlanningDashBoard(strItemID As String)

    Dim f As New frmProductionPlanningDashBoard
    
    Load f
    f.ShowMe (strItemID)
    
End Sub

Private Sub LVWhite_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Dim myLV As ListView
    Set myLV = LVWhite
    
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

Private Sub LVYellow_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Dim myLV As ListView
    Set myLV = LVYellow
    
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

Private Sub LVYellow_DblClick()
    If LVYellow.ListItems.count = 0 Then Exit Sub
    Call ShowDetails(LVYellow, "Yellow")
End Sub

Private Sub LVYellow_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If LVYellow.ListItems.count = 0 Then Exit Sub
        
    '''''''''''''''''''
    Dim strRet As String
    strRet = ShowPopUpMenu1("Show Planning Dashboard")
    
    If strRet = "Show Planning Dashboard" Then
        Call ShowProductionPlanningDashBoard(LVYellow.SelectedItem.SubItems(4))
    End If
End Sub

Private Sub Opts_Click(Index As Integer)
    Opts(0).Tag = Index
End Sub
