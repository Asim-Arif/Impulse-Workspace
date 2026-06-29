VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmOrderPlanning 
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   45
   ClientWidth     =   14250
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
   ScaleHeight     =   8910
   ScaleMode       =   0  'User
   ScaleWidth      =   12763.92
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
      Height          =   9030
      Left            =   0
      TabIndex        =   0
      Top             =   -105
      Width           =   14265
      Begin VB.TextBox txtCustomer 
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
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   6390
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   795
         Width           =   3135
      End
      Begin VB.TextBox txtInternalOrderNo 
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
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   3210
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   795
         Width           =   3135
      End
      Begin VB.PictureBox PicEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5355
         Left            =   9075
         ScaleHeight     =   5325
         ScaleWidth      =   5040
         TabIndex        =   4
         Top             =   2985
         Visible         =   0   'False
         Width           =   5070
         Begin ComboList.Usercontrol1 cmbProcess 
            Height          =   285
            Left            =   15
            TabIndex        =   29
            Top             =   195
            Width           =   5010
            _ExtentX        =   8837
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
            TextBackColor   =   12640511
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
         Begin ComboList.Usercontrol1 cmbMaker 
            Height          =   285
            Left            =   15
            TabIndex        =   30
            Top             =   720
            Width           =   5010
            _ExtentX        =   8837
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
            TextBackColor   =   12640511
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
         Begin VB.TextBox txtRate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1815
            TabIndex        =   25
            Top             =   1320
            Width           =   870
         End
         Begin VB.TextBox txtBalance 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   915
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   1320
            Width           =   870
         End
         Begin VB.TextBox txtCapacity 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   15
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   1320
            Width           =   870
         End
         Begin VB.TextBox txtTotalAdj 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   3870
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   4920
            Width           =   1140
         End
         Begin MSComctlLib.ListView LVAdjAll 
            Height          =   2790
            Left            =   60
            TabIndex        =   16
            Top             =   1920
            Visible         =   0   'False
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   4921
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Process"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Maker"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "del"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Qty"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Rate"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView LVAdj 
            Height          =   3210
            Left            =   15
            TabIndex        =   15
            Top             =   1665
            Width           =   5010
            _ExtentX        =   8837
            _ExtentY        =   5662
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
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Process"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Maker"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "D. Date"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Qty"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Rate"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4155
            TabIndex        =   5
            Top             =   1320
            Width           =   870
         End
         Begin MSComCtl2.DTPicker DTDelivery 
            Height          =   300
            Left            =   2715
            TabIndex        =   27
            Top             =   1320
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   203882499
            CurrentDate     =   40493
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Delivery Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   2730
            TabIndex        =   28
            Top             =   1095
            Width           =   1155
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rate"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   1995
            TabIndex        =   26
            Top             =   1095
            Width           =   405
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Balance"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   975
            TabIndex        =   24
            Top             =   1095
            Width           =   660
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Capacity"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   15
            TabIndex        =   22
            Top             =   1095
            Width           =   735
         End
         Begin MSForms.CommandButton cmdCancel 
            Height          =   375
            Left            =   1095
            TabIndex        =   14
            Top             =   4905
            Width           =   1050
            ForeColor       =   0
            BackColor       =   16777215
            Caption         =   "Cancel"
            PicturePosition =   327683
            Size            =   "1852;661"
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdOK 
            Height          =   375
            Left            =   15
            TabIndex        =   13
            Top             =   4905
            Width           =   1050
            ForeColor       =   0
            BackColor       =   16777215
            Caption         =   "OK"
            PicturePosition =   327683
            Size            =   "1852;661"
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
            BackStyle       =   0  'Transparent
            Caption         =   "Process"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   15
            TabIndex        =   12
            Top             =   0
            Width           =   660
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Index           =   1
            Left            =   4380
            TabIndex        =   11
            Top             =   1095
            Width           =   300
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Maker"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   15
            TabIndex        =   10
            Top             =   510
            Width           =   540
         End
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
         TabIndex        =   18
         Top             =   0
         Width           =   14250
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Order Planning"
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
            Left            =   0
            TabIndex        =   19
            Top             =   135
            Width           =   14205
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Order Planning"
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
            Left            =   0
            TabIndex        =   20
            Top             =   150
            Width           =   14205
         End
      End
      Begin MSComCtl2.DTPicker DT 
         Height          =   300
         Left            =   7560
         TabIndex        =   8
         Top             =   8460
         Visible         =   0   'False
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   192020483
         CurrentDate     =   40493
      End
      Begin VB.TextBox txtOrderNo 
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
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   45
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   810
         Width           =   3135
      End
      Begin MSComctlLib.ListView LV 
         Height          =   7140
         Left            =   60
         TabIndex        =   7
         Top             =   1200
         Width           =   14085
         _ExtentX        =   24844
         _ExtentY        =   12594
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
            Text            =   "S#"
            Object.Width           =   661
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item Name"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Type"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "History/Qty"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Pack"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Semi"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Forg."
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "F. PO"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "H1"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "H2"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "Open/Stock PO"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   11
            Text            =   "Total C."
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   12
            Text            =   "To Pur."
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   13
            Text            =   "Adj"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   14
            Text            =   "Plan"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   15
            Text            =   "Disp."
            Object.Width           =   1499
         EndProperty
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   375
         Left            =   1470
         TabIndex        =   36
         Top             =   8385
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Print"
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
      Begin VB.Label lbl 
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   6390
         TabIndex        =   35
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Internal Order No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   3210
         TabIndex        =   33
         Top             =   600
         Width           =   1515
      End
      Begin MSForms.CommandButton cmdApprove 
         Height          =   375
         Left            =   60
         TabIndex        =   31
         Top             =   8385
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Approve"
         PicturePosition =   327683
         Size            =   "2355;661"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   6990
         TabIndex        =   9
         Top             =   8475
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSForms.CommandButton cmdSave 
         Height          =   375
         Left            =   11415
         TabIndex        =   6
         Top             =   8385
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Save"
         PicturePosition =   327683
         Size            =   "2355;661"
         Accelerator     =   83
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label lbl 
         Caption         =   "Order No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   45
         TabIndex        =   2
         Top             =   615
         Width           =   1155
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   375
         Left            =   12810
         TabIndex        =   1
         Top             =   8385
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
   End
End
Attribute VB_Name = "frmOrderPlanning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iColNo As Integer
Dim strOrderNo As String
Dim strDelList As String, strDelList_Temp As String
Dim lProcessID_Selected As Long, lVendID_Selected As Long

Private Sub cmbmaker_matched()

    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM VMakerBalances WHERE VendID=" & cmbMaker.ID, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            txtCapacity = Val(!MakerCapacity & "")
            txtBalance = Val(!Balance & "")
        End If
        .Close
        Dim dRate As Double
        dRate = GetSingleDoubleValue("Rate", "VendAssItems", " WHERE VendID=" & cmbMaker.ID & " AND ProcessID=" & cmbProcess.ID)
        If dRate = 0 Then dRate = 1
        txtRate = dRate
    End With
    
End Sub

Private Sub cmbProcess_matched()
    cmbMaker.ClearVals
    'cmbMaker.AddVals con, "VendID1+' {'+VenderName+'}'", "VMakers", "VendID", " WHERE VendID IN(SELECT VendID FROM VendAssItems WHERE ProcessID=" & cmbProcess.ID & " AND ItemID='" & LV.SelectedItem.ListSubItems(1).Tag & "')"
    cmbMaker.AddVals con, "VendID1+' {'+VenderName+'}'", "VMakers", "VendID", " WHERE VendID IN(SELECT VendID FROM MakerProcesses WHERE ProcessID=" & cmbProcess.ID & ")"
    cmbMaker.ID = lVendID_Selected
End Sub

Private Sub cmdApprove_Click()
    
    If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbYes Then
        Dim lCount As Long
        lCount = GetSingleLongValue("COUNT(*)", "OrderPlanningDetails", " WHERE OrderNo='" & strOrderNo & "'")
        If lCount = 0 Then
            MsgBox "No Plan !!!"
            Exit Sub
        End If
        con.Execute "UPDATE FCustomerOrders SET OrderPlanApproved=1 WHERE OrderNo='" & strOrderNo & "'"
        cmdApprove.Enabled = False
    End If
    
End Sub

Private Sub cmdCancel_Click()
    PicEdit.Visible = False
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
    
    Dim i As Integer
    For i = LVAdjAll.ListItems.count To 1 Step -1
        If Val(LVAdjAll.ListItems(i).Tag) = Val(LV.SelectedItem.Index) Then
            LVAdjAll.ListItems.Remove i
        End If
    Next
    
    'Now Add these.
    Dim ITM As ListItem
    For i = 1 To LVAdj.ListItems.count
        Set ITM = LVAdjAll.ListItems.add(, , LVAdj.ListItems(i).Text)
        ITM.Tag = LV.SelectedItem.Index
        ITM.ListSubItems.add(, , LVAdj.ListItems(i).SubItems(1)).Tag = LVAdj.ListItems(i).ListSubItems(1).Tag
        ITM.ListSubItems.add(, , LVAdj.ListItems(i).SubItems(2)).Tag = LVAdj.ListItems(i).ListSubItems(2).Tag
        ITM.ListSubItems.add(, , LVAdj.ListItems(i).SubItems(3)).Tag = LVAdj.ListItems(i).ListSubItems(3).Tag
        ITM.ListSubItems.add(, , LVAdj.ListItems(i).SubItems(4)).Tag = LVAdj.ListItems(i).ListSubItems(4).Tag
        'ITM.ListSubItems.add(, , "").Tag = LVAdj.ListItems(i).key
    Next
    LV.SelectedItem.SubItems(14) = Val(txtTotalAdj)
    PicEdit.Visible = False
    strDelList = strDelList & strDelList_Temp
    
    lProcessID_Selected = cmbProcess.ID
    lVendID_Selected = cmbMaker.ID
    
End Sub

Private Sub cmdPrint_Click()
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\OrderPlanningDetails.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{FCustomerOrders.OrderNo}='" & strOrderNo & "'", rpt
End Sub

Private Sub cmdSave_Click()
    If Saved Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()

    iColNo = 0
    DT.Value = Date
    DTDelivery.Value = Date
    strDelList_Temp = ""
    strDelList = ""
    
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
        LV.SortKey = ColumnHeader.Index - 1
        LV.SortOrder = lvwAscending
        'ColumnHeader.Icon = "Down"
    End If
    If LV.Sorted = False Then LV.Sorted = True
End Sub

Private Sub LV_DblClick()
    
    Dim i As Integer
        
    If LV.ListItems.count = 0 Then Exit Sub
    If cmdApprove.Enabled = False Then Exit Sub
    
    Dim lSNo As Long, strItemID As String
    strItemID = LV.SelectedItem.ListSubItems(1).Tag
    
    cmbProcess.ClearVals
    cmbProcess.AddVals con, "Description", "VItemProcesses", "ProcessID", " WHERE IsExist='" & strItemID & "' ORDER BY ItemSNo"
    cmbProcess.ID = lProcessID_Selected
    
'    If Val(LV.SelectedItem.ListSubItems(7).Tag) > 0 Then
'        For i = 0 To cmbProcess.ListCount - 1
'            If Val(cmbProcess.List(i, 1)) = Val(LV.SelectedItem.ListSubItems(7).Tag) Then
'                cmbProcess.ListIndex = i
'                Exit For
'            End If
'        Next
'    End If
    
    '----------------------------------------------
    LVAdj.ListItems.Clear
    txtEdit = ""
    strDelList_Temp = ""
    'Call Already Added.
    Dim ITM As ListItem
    For i = 1 To LVAdjAll.ListItems.count
        If Val(LVAdjAll.ListItems(i).Tag) = Val(LV.SelectedItem.Index) Then
            'Add These...
            Set ITM = LVAdj.ListItems.add(, , LVAdjAll.ListItems(i).Text)
            ITM.Tag = LVAdjAll.ListItems(i).Tag
            ITM.ListSubItems.add(, , LVAdjAll.ListItems(i).SubItems(1)).Tag = LVAdjAll.ListItems(i).ListSubItems(1).Tag
            ITM.ListSubItems.add(, , LVAdjAll.ListItems(i).SubItems(2)).Tag = LVAdjAll.ListItems(i).ListSubItems(2).Tag
            ITM.ListSubItems.add(, , LVAdjAll.ListItems(i).SubItems(3)).Tag = LVAdjAll.ListItems(i).ListSubItems(3).Tag
            ITM.ListSubItems.add(, , LVAdjAll.ListItems(i).SubItems(4)).Tag = LVAdjAll.ListItems(i).ListSubItems(4).Tag
        End If
    Next
    Call GetAdjTotalQty
    
    With PicEdit
        .Move (LV.Width + LV.Left) - PicEdit.Width, LV.Top + LV.SelectedItem.Top
        If PicEdit.Top > 2520 Then PicEdit.Top = 2520
'            PicEdit.Top = LV.Top + LV.Height - PicEdit.Height
'        End If
        .Visible = True
        txtEdit.Text = Val(LV.SelectedItem.SubItems(6))
        'cmbLocation_QEL.SetFocus
        cmbProcess.SetFocus
    End With
    
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    PicEdit.Visible = False
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        If x > LV.ColumnHeaders(1).Left And x < LV.ColumnHeaders(i).Left + LV.ColumnHeaders(i).Width Then
            iColNo = i
            Exit For
        End If
    Next
    
End Sub

Public Sub ShowMe(p_strOrderNo As String)
    
    strOrderNo = p_strOrderNo
    txtOrderNo = strOrderNo
    On Error GoTo err
    
    Dim rs As New ADODB.Recordset, ITM As ListItem, cmd As New ADODB.Command
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "OrderPlanningReport_SP"
        .Parameters("@OrderNo").Value = strOrderNo
        .Parameters("@FromItems").Value = False
        .Parameters("@ItemCode").Value = "0"
        .Parameters("@TempList_RefID").Value = 0
        Set rs = .Execute
    End With
    
    Dim lQtyToPurchase As Long, strItemID As String, lPlanned As Long, lQtyAdjusted As Long
    Dim rsPlan As New ADODB.Recordset, ITMPlan As ListItem, lTotalC As Long
    With rs
        
        LV.ListItems.Clear
        LVAdjAll.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            ITM.Tag = ""
            ITM.ListSubItems.add(, , !ITemID & " - " & !ItemName & "").Tag = !ITemID & ""
            ITM.ListSubItems.add , , !ItemQuality & ""
            ITM.ListSubItems.add , , Val(!AllHits & "") & "/" & Val(!AllQty & "")
            ITM.ListSubItems.add , , Val(!InHand & "")
            ITM.ListSubItems.add , , Val(!SFQty & "")
            ITM.ListSubItems.add , , Val(!ForgingStock & "")
            ITM.ListSubItems.add , , Val(!QtyToRcv & "")
            ITM.ListSubItems.add , , Val(!InProcessLotsQtyB4 & "") & "/" & Val(!InProcessLotsQtyB4Stock & "")
            ITM.ListSubItems.add , , Val(!InProcessLotsQty & "") & "/" & Val(!InProcessLotsQtyStock & "")
            ITM.ListSubItems.add , , Val(!OpenPOQty & "") & "/" & Val(!OpenPOQtyStock & "")
            lTotalC = Val(!Qty & "") - Val(!InProcessLotsQtyB4 & "") - Val(!InProcessLotsQty & "") - Val(!OpenPOQty & "")
            ITM.ListSubItems.add , , Val(!Qty & "") & "/" & Val(!TotalOrdersQty & "")
            'ITM.ListSubItems.add , , Val(!Qty & "") & "/" & Val(!TotalOrdersQty & "")
            'lQtyToPurchase = Val(!TotalOrdersQty & "") - Val(!InHand & "") - Val(!SFQty & "") - Val(!ForgingStock & "") - Val(!QtyToRcv & "") - Val(!InProcessLotsQty & "") - Val(!InProcessLotsQtyB4 & "") - Val(!OpenPOQty & "")
            
            strItemID = !ITemID & ""
            lPlanned = 0
            With rsPlan
                .Open "SELECT * FROM VOrderPlanningDetails WHERE OrderNo='" & strOrderNo & "' AND ItemID='" & strItemID & "' ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
                Do Until .EOF
                    Set ITMPlan = LVAdjAll.ListItems.add(, , !Description & "")
                    ITMPlan.Tag = ITM.Index
                    ITMPlan.ListSubItems.add(, , !VenderName & "").Tag = Val(!VendID & "")
                    ITMPlan.ListSubItems.add(, , Format(!DeliveryDT, "dd-MM-yy")).Tag = Format(!DeliveryDT, "dd-MMM-yyyy")
                    ITMPlan.ListSubItems.add(, , Val(!Qty & "")).Tag = Val(!ProcessID & "")
                    ITMPlan.ListSubItems.add(, , Val(!Rate & "")).Tag = Val(!EntryID & "")
                    lPlanned = lPlanned + Val(!Qty & "")
                    .MoveNEXT
                Loop
                .Close
                
                .Open "SELECT InternalRefNo,CustCode FROM FCustomerOrders WHERE OrderNo='" & strOrderNo & "'", con, adOpenForwardOnly, adLockReadOnly
                txtInternalOrderNo = !InternalRefNo & ""
                txtCustomer = !CustCode & ""
                .Close
                
            End With
            
            lQtyAdjusted = GetSingleLongValue("SUM(Qty)", "StockOrderAdjustments", " WHERE OrderNo='" & strOrderNo & "' AND ItemID='" & !ITemID & "'")
            'lQtyToPurchase = Val(!TotalOrdersQty & "") - Val(!InHand & "") - Val(!ForgingStock & "") - Val(!QtyToRcv & "") - Val(!InProcessLotsQty & "") - Val(!InProcessLotsQtyB4 & "") - Val(!OpenPOQty & "") - lQtyAdjusted
            lQtyToPurchase = Val(!Qty & "") - Val(!InProcessLotsQty & "") - Val(!InProcessLotsQtyB4 & "") - Val(!InHand & "") - Val(!DispatchedQty & "")
            ITM.ListSubItems.add , , lQtyToPurchase
            ITM.ListSubItems.add , , lQtyAdjusted
            
            ITM.ListSubItems.add , , lPlanned
            ITM.ListSubItems.add , , Val(!DispatchedQty & "")
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Dim bApproved As Boolean
    bApproved = GetSingleBooleanValue("OrderPlanApproved", "FCustomerOrders", " WHERE OrderNo='" & strOrderNo & "'")
    If bApproved Then
        cmdApprove.Enabled = False
    Else
        cmdApprove.Enabled = True
    End If
            
    Me.Show
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub LVAdj_KeyDown(KeyCode As Integer, Shift As Integer)

    If LVAdj.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbYes Then
            If Val(LVAdj.SelectedItem.ListSubItems(4).Tag) > 0 Then
                strDelList_Temp = strDelList_Temp & Val(LVAdj.SelectedItem.ListSubItems(4).Tag) & ","
            End If
            LVAdj.ListItems.Remove LVAdj.SelectedItem.Index
        End If
    End If
    
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        If cmbProcess.MatchFound = False Then
            MsgBox "Please Select Process.", vbInformation
            Exit Sub
        ElseIf cmbMaker.MatchFound = False Then
            MsgBox "Please Select Maker.", vbInformation
            Exit Sub
        ElseIf Val(txtEdit) < 0 Then
            MsgBox "Invalid Qty.", vbInformation
            Exit Sub
        End If
        
'        If Val(txtEdit) > Val(LV.SelectedItem.SubItems(5)) Or Val(txtEdit) <= 0 Or Val(txtEdit) > Val(cmbLocation_QEL.List(cmbLocation_QEL.ListIndex, 2)) Then
'            MsgBox "Invalid Qty."
'            Exit Sub
'        End If
        
        Dim ITM As ListItem
        Set ITM = LVAdj.ListItems.add(, , cmbProcess.Text)
        ITM.ListSubItems.add(, , cmbMaker.Text).Tag = cmbMaker.ID
        ITM.ListSubItems.add(, , Format(DTDelivery, "dd-MM-yy")).Tag = Format(DTDelivery, "dd-MMM-yyyy")
        ITM.ListSubItems.add(, , Val(txtEdit)).Tag = cmbProcess.ID
        ITM.ListSubItems.add(, , Val(txtRate)).Tag = "0"
        
        Call GetAdjTotalQty
    ElseIf KeyAscii = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub

Private Function GetAdjTotalQty() As Long

    Dim i As Integer, lTotal As Long
    For i = 1 To LVAdj.ListItems.count
        lTotal = lTotal + Val(LVAdj.ListItems(i).SubItems(3))
    Next
    GetAdjTotalQty = lTotal
    txtTotalAdj = GetAdjTotalQty
    
End Function

Private Function Saved() As Boolean
    
    On Error GoTo err
                   
    Call StartTrans(con)
    
    If strDelList <> "" Then
        strDelList = Left(strDelList, Len(strDelList) - 1)
        con.Execute "DELETE FROM OrderPlanningDetails WHERE EntryID IN(" & strDelList & ")"
    End If
    
    Dim i As Integer, j As Integer
    Dim strItemID As String
    Dim lVendID As Long, lProcessID As Long, dRate As Double
    Dim lCount_MAI As Long
    
    Dim lVendors() As Long, bVendorFound As Boolean
    Dim iVend_Index As Integer
    iVend_Index = 0
    ReDim lVendors(LVAdjAll.ListItems.count)
    Dim k As Integer
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            strItemID = .ListSubItems(1).Tag
            
            
            For j = 1 To LVAdjAll.ListItems.count
                With LVAdjAll.ListItems(j)
                    bVendorFound = False
                    lVendID = Val(.ListSubItems(1).Tag)
                    lProcessID = Val(.ListSubItems(3).Tag)
                    dRate = Val(.SubItems(4))
                    If Val(.Tag) = Val(LV.ListItems(i).Index) And (Val(.ListSubItems(4).Tag) = 0) Then
                        
                        con.Execute "INSERT INTO OrderPlanningDetails(OrderNo,ItemID,ProcessID,VendID,DeliveryDT,Qty,Rate) VALUES('" & _
                         strOrderNo & "','" & strItemID & "'," & lProcessID & "," & lVendID & _
                         ",'" & .ListSubItems(2).Tag & "'," & Val(.SubItems(3)) & "," & dRate & ")"
                                                
                        lCount_MAI = GetSingleLongValue("COUNT(*)", "VendAssItems", " WHERE VendID=" & lVendID & " AND ProcessID=" & lProcessID & " AND ItemID='" & strItemID & "'")
                        If lCount_MAI = 0 Then
                            con.Execute "INSERT INTO VendAssItems(VendID,ProcessID,ItemID,Rate,Unit,Remarks) VALUES(" & lVendID & "," & lProcessID & ",'" & strItemID & "'," & dRate & ",'Pcs','Auto')"
                        End If
                        
                        For k = 0 To UBound(lVendors)
                            If lVendID = lVendors(k) Then
                                bVendorFound = True
                                Exit For
                            End If
                        Next
                        If bVendorFound = False Then    'Add Vendor to Array
                            lVendors(iVend_Index) = lVendID
                            iVend_Index = iVend_Index + 1
                        End If
                        
                    End If
                End With
            Next
        End With
    Next
    
    'Now Generate Purchase Orders
    Dim myDT As Date
    myDT = Date
    Dim strMasterPONo As String, lMakerIssNo As Long, IssID As String, strMakerIssNo As String
    Dim lVI_EntryID As Long, lQty As Long, lVID_EntryID As Long, strIssNo As String
    Dim lVendID_Previous As Long
    lVendID_Previous = 0
    For k = 0 To iVend_Index - 1
        For j = 1 To LVAdjAll.ListItems.count
            With LVAdjAll.ListItems(j)
                If (Val(.ListSubItems(4).Tag) = 0) Then
                    lVendID = Val(.ListSubItems(1).Tag)
                    lProcessID = Val(.ListSubItems(3).Tag)
                    dRate = Val(.SubItems(4))
                    lQty = Val(.SubItems(3))
                    If lVendID = lVendors(k) Then
                        If lVendID_Previous <> lVendID Then 'Generate Master PO No.
                            strMasterPONo = GetNextMasterPONo
                        End If
                        strItemID = LV.ListItems(Val(.Tag)).ListSubItems(1).Tag
                        
                        lMakerIssNo = Val(con.Execute("SELECT MAX(CAST(RIGHT(RecieptID,Len(RecieptID)-12) AS INT)) FROM VendIssued WHERE Convert(varchar,DT,6)='" & Format(myDT, "dd MMM yy") & "'").Fields(0).Value & "")
                        lMakerIssNo = lMakerIssNo + 1
                        IssID = "ISU-" & Format(myDT, "ddMMyy") & lMakerIssNo
                        strMakerIssNo = "M-ISU-" & Format(myDT, "ddMMyy") & lMakerIssNo
                    
                        con.Execute "INSERT INTO VendIssued(VendID,DT,RecieptID,UserID,ProcessID,ItemID,UserName,MachineName,SpecialInstructions,ExcessQtyPercentage" & _
                        ",MaximumRcvingsAgainstPO,MasterPONo,SampleProvided,DrawingProvided,ForgingProvided,SteelProvided,Authorized,IssEmpID," & _
                        "SteelType_RefID) VALUES(" & lVendID & ",'" & myDT & "','" & strMakerIssNo & "'," & _
                        UserID & "," & lProcessID & ",'" & strItemID & "','" & CurrentUserName & "','" & _
                        strComputerName & "','',0,100,'" & strMasterPONo & "',0,0,0,0,1,'',0)"
    
                        lVI_EntryID = GetSingleLongValue("MAX(EntryID)", "VendIssued", " WHERE MachineName='" & strComputerName & "'")
                        
                        con.Execute "INSERT INTO VendIssdDetail(RefID,RecieptID,ItemCode,Rate,IssQty,ReqAuth,OrderNo,RcvProcessID," & _
                            "ReturnDT,Priority,Rcvd_RefID,LotNo,ReWorkLot,Repair_RefID) VALUES(" & lVI_EntryID & ",'" & _
                            strIssNo & "','" & strItemID & "'," & dRate & "," & lQty & ",0,'" & strOrderNo & "'," & lProcessID & ",'" & _
                            myDT & "',0,0,'0',0,0)"
                                            
                        lVID_EntryID = GetSingleLongValue("MAX(EntryID)", "VendIssdDetail")
                    End If
                    lVendID_Previous = lVendID
                End If
            End With
            
        Next
    Next
    
    con.CommitTrans
    Saved = True
    
    Exit Function
err:
    MsgBox err.Description
End Function
 
Private Function GetNextMasterPONo() As String

    Dim Prefix As String
    'ORD-DDMM-001
    Dim myDT As Date
    myDT = Date
    Dim lOrderNo As Long
    lOrderNo = GetSingleLongValue("MAX(CAST(LEFT(MasterPONo,CHARINDEX('/',MasterPONo)-1) AS INT))", "VendIssued", " WHERE RIGHT(MasterPONo,4)='" & Format(myDT, "MMyy") & "' AND MasterPONo IS NOT NULL")
    lOrderNo = lOrderNo + 1
    GetNextMasterPONo = lOrderNo & Format(myDT, "/MMyy")
    
End Function

