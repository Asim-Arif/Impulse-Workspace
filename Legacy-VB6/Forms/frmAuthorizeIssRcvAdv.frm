VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmAuthorizeIssRcvAdv 
   AutoRedraw      =   -1  'True
   ClientHeight    =   9525
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   12360
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
   ScaleHeight     =   9525
   ScaleMode       =   0  'User
   ScaleWidth      =   11071.02
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
      Left            =   615
      TabIndex        =   1
      Top             =   0
      Width           =   10560
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Authorization"
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
         Left            =   105
         TabIndex        =   2
         Top             =   165
         Width           =   10605
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Authorization"
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
         Left            =   120
         TabIndex        =   3
         Top             =   180
         Width           =   10590
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
      Height          =   8880
      Left            =   15
      TabIndex        =   0
      Top             =   615
      Width           =   12345
      Begin ComboList.Usercontrol1 cmbEmpID 
         Height          =   285
         Left            =   7080
         TabIndex        =   77
         Top             =   7395
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
      Begin VB.Frame Frame2 
         Caption         =   "Search Box :"
         Height          =   690
         Left            =   2175
         TabIndex        =   18
         Top             =   7200
         Width           =   4860
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   1950
            TabIndex        =   21
            Top             =   240
            Width           =   2025
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Enabled         =   0   'False
            Height          =   300
            Left            =   3975
            TabIndex        =   20
            Top             =   240
            Width           =   765
         End
         Begin VB.ComboBox cmbfield 
            Height          =   315
            ItemData        =   "frmAuthorizeIssRcvAdv.frx":0000
            Left            =   60
            List            =   "frmAuthorizeIssRcvAdv.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   240
            Width           =   1875
         End
      End
      Begin VB.TextBox txtBarcode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   9705
         TabIndex        =   15
         Top             =   7530
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.CommandButton cmdUnchkAll 
         Caption         =   "Uncheck All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9585
         TabIndex        =   12
         Top             =   8385
         Width           =   1305
      End
      Begin VB.CommandButton cmdChkAll 
         Caption         =   "Check All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8220
         TabIndex        =   11
         Top             =   8385
         Width           =   1305
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6870
         Left            =   75
         TabIndex        =   4
         Top             =   240
         Width           =   12195
         _ExtentX        =   21511
         _ExtentY        =   12118
         _Version        =   393216
         Tab             =   1
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Authorize Issued"
         TabPicture(0)   =   "frmAuthorizeIssRcvAdv.frx":0004
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "LVIss"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Authorize Received"
         TabPicture(1)   =   "frmAuthorizeIssRcvAdv.frx":0020
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "LVRcv"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Pic"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "txtEdit"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "FraWastageOwn"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "FraEdit"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "FraAddRework"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "FraAddWastage"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).ControlCount=   7
         TabCaption(2)   =   "Authorize Raw Material"
         TabPicture(2)   =   "frmAuthorizeIssRcvAdv.frx":003C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "LVRM"
         Tab(2).ControlCount=   1
         Begin VB.Frame FraAddWastage 
            BackColor       =   &H00E7EBEF&
            Caption         =   "Add Wastage / Rejection"
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
            Height          =   6105
            Left            =   7215
            TabIndex        =   61
            Top             =   435
            Visible         =   0   'False
            Width           =   4710
            Begin VB.CommandButton cmdCanelWastageX 
               Caption         =   "Cancel"
               Height          =   345
               Left            =   1035
               TabIndex        =   68
               Top             =   5580
               Width           =   870
            End
            Begin VB.TextBox txtTotalQtyWastage 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2925
               Locked          =   -1  'True
               TabIndex        =   67
               Top             =   5580
               Width           =   960
            End
            Begin VB.CommandButton cmdOKWastageX 
               Caption         =   "&Save"
               Height          =   345
               Left            =   105
               TabIndex        =   66
               Top             =   5580
               Width           =   870
            End
            Begin VB.CommandButton cmdAddWastage 
               Caption         =   "&Add"
               Height          =   525
               Left            =   3705
               TabIndex        =   72
               Top             =   1305
               Width           =   870
            End
            Begin VB.TextBox txtQtyToWastage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2715
               TabIndex        =   71
               Top             =   1545
               Width           =   960
            End
            Begin VB.TextBox txtProcessWastage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00D3E7EF&
               Height          =   285
               Left            =   90
               Locked          =   -1  'True
               TabIndex        =   65
               TabStop         =   0   'False
               Top             =   1005
               Width           =   3750
            End
            Begin VB.TextBox txtQtyWastage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00D3E7EF&
               Height          =   285
               Left            =   3825
               Locked          =   -1  'True
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   1005
               Width           =   765
            End
            Begin VB.TextBox txtItemNameWastage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00D3E7EF&
               Height          =   285
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   63
               TabStop         =   0   'False
               Top             =   510
               Width           =   3030
            End
            Begin VB.TextBox txtItemCodeWastage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00D3E7EF&
               Height          =   285
               Left            =   90
               Locked          =   -1  'True
               TabIndex        =   62
               TabStop         =   0   'False
               Top             =   510
               Width           =   1485
            End
            Begin MSComctlLib.ListView LVWastageAll 
               Height          =   1650
               Left            =   510
               TabIndex        =   69
               Top             =   2490
               Visible         =   0   'False
               Width           =   3510
               _ExtentX        =   6191
               _ExtentY        =   2910
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
               NumItems        =   4
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Process"
                  Object.Width           =   10583
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Object.Width           =   1693
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   2
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Object.Width           =   0
               EndProperty
            End
            Begin ComboList.Usercontrol1 cmbWastage 
               Height          =   285
               Left            =   90
               TabIndex        =   70
               Top             =   1545
               Width           =   2610
               _ExtentX        =   4604
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
            Begin MSComctlLib.ListView LVWastage 
               Height          =   3675
               Left            =   105
               TabIndex        =   73
               Top             =   1860
               Width           =   4500
               _ExtentX        =   7938
               _ExtentY        =   6482
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
               NumItems        =   4
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Process"
                  Object.Width           =   4657
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   1
                  Text            =   "Qty"
                  Object.Width           =   1693
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   2
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Object.Width           =   0
               EndProperty
            End
            Begin MSForms.Label Label7 
               Height          =   240
               Left            =   90
               TabIndex        =   76
               Top             =   780
               Width           =   4500
               BackColor       =   11517387
               Caption         =   " Process                                                                   Qty"
               Size            =   "7937;423"
               BorderColor     =   16512
               BorderStyle     =   1
               FontName        =   "Tahoma"
               FontEffects     =   1073741825
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               FontWeight      =   700
            End
            Begin MSForms.Label Label8 
               Height          =   240
               Left            =   90
               TabIndex        =   75
               Top             =   1320
               Width           =   3585
               BackColor       =   11517387
               Caption         =   " Wastage Type                                     Qty"
               Size            =   "6324;423"
               BorderColor     =   16512
               BorderStyle     =   1
               FontName        =   "Tahoma"
               FontEffects     =   1073741825
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               FontWeight      =   700
            End
            Begin MSForms.Label Label9 
               Height          =   240
               Left            =   90
               TabIndex        =   74
               Top             =   285
               Width           =   4500
               BackColor       =   11517387
               Caption         =   $"frmAuthorizeIssRcvAdv.frx":0058
               Size            =   "7937;423"
               BorderColor     =   16512
               BorderStyle     =   1
               FontName        =   "Tahoma"
               FontEffects     =   1073741825
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               FontWeight      =   700
            End
         End
         Begin VB.Frame FraAddRework 
            BackColor       =   &H00E7EBEF&
            Caption         =   "Add Re-Work"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   6105
            Left            =   4905
            TabIndex        =   45
            Top             =   435
            Visible         =   0   'False
            Width           =   7020
            Begin VB.CommandButton cmdCanelRW 
               Caption         =   "Cancel"
               Height          =   345
               Left            =   1035
               TabIndex        =   60
               Top             =   5580
               Width           =   870
            End
            Begin MSComctlLib.ListView LVRWAll 
               Height          =   1650
               Left            =   510
               TabIndex        =   57
               Top             =   2490
               Visible         =   0   'False
               Width           =   3510
               _ExtentX        =   6191
               _ExtentY        =   2910
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
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Process"
                  Object.Width           =   10583
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Object.Width           =   1693
               EndProperty
            End
            Begin VB.TextBox txtTotalQtyRW 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2925
               Locked          =   -1  'True
               TabIndex        =   58
               Top             =   5580
               Width           =   960
            End
            Begin VB.CommandButton cmdOKRW 
               Caption         =   "&Save"
               Height          =   345
               Left            =   105
               TabIndex        =   56
               Top             =   5580
               Width           =   870
            End
            Begin VB.CommandButton cmdAddRW 
               Caption         =   "&Add"
               Height          =   525
               Left            =   6285
               TabIndex        =   53
               Top             =   1305
               Width           =   660
            End
            Begin VB.TextBox txtQtyToRW 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5580
               TabIndex        =   52
               Top             =   1545
               Width           =   660
            End
            Begin VB.TextBox txtProcessRW 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00D3E7EF&
               Height          =   285
               Left            =   90
               Locked          =   -1  'True
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   1005
               Width           =   6210
            End
            Begin VB.TextBox txtQtyRW 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00D3E7EF&
               Height          =   285
               Left            =   6285
               Locked          =   -1  'True
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   1005
               Width           =   660
            End
            Begin VB.TextBox txtItemNameRW 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00D3E7EF&
               Height          =   285
               Left            =   2130
               Locked          =   -1  'True
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   510
               Width           =   4815
            End
            Begin VB.TextBox txtItemCodeRW 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00D3E7EF&
               Height          =   285
               Left            =   90
               Locked          =   -1  'True
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   510
               Width           =   2055
            End
            Begin ComboList.Usercontrol1 cmbRepairs 
               Height          =   285
               Left            =   90
               TabIndex        =   51
               Top             =   1545
               Width           =   5460
               _ExtentX        =   9631
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
            Begin MSComctlLib.ListView LV 
               Height          =   3675
               Left            =   105
               TabIndex        =   55
               Top             =   1860
               Width           =   6840
               _ExtentX        =   12065
               _ExtentY        =   6482
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
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Repair"
                  Object.Width           =   8819
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   1
                  Text            =   "Qty"
                  Object.Width           =   1693
               EndProperty
            End
            Begin MSForms.Label Label6 
               Height          =   240
               Left            =   105
               TabIndex        =   59
               Top             =   780
               Width           =   6840
               BackColor       =   11517387
               Caption         =   $"frmAuthorizeIssRcvAdv.frx":0122
               Size            =   "12065;423"
               BorderColor     =   16512
               BorderStyle     =   1
               FontName        =   "Tahoma"
               FontEffects     =   1073741825
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               FontWeight      =   700
            End
            Begin MSForms.Label Label5 
               Height          =   240
               Left            =   90
               TabIndex        =   54
               Top             =   1320
               Width           =   6150
               BackColor       =   11517387
               Caption         =   " Repair Type                                                                                                     Qty"
               Size            =   "10848;423"
               BorderColor     =   16512
               BorderStyle     =   1
               FontName        =   "Tahoma"
               FontEffects     =   1073741825
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               FontWeight      =   700
            End
            Begin MSForms.Label Label4 
               Height          =   240
               Left            =   105
               TabIndex        =   50
               Top             =   285
               Width           =   6840
               BackColor       =   11517387
               Caption         =   $"frmAuthorizeIssRcvAdv.frx":01AD
               Size            =   "12065;423"
               BorderColor     =   16512
               BorderStyle     =   1
               FontName        =   "Tahoma"
               FontEffects     =   1073741825
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               FontWeight      =   700
            End
         End
         Begin VB.Frame FraEdit 
            BackColor       =   &H00E7EBEF&
            Height          =   2820
            Left            =   960
            TabIndex        =   22
            Top             =   2775
            Visible         =   0   'False
            Width           =   10965
            Begin VB.TextBox txtItemCode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00D3E7EF&
               Height          =   285
               Left            =   60
               Locked          =   -1  'True
               TabIndex        =   36
               Top             =   1350
               Width           =   1470
            End
            Begin VB.TextBox txtItemName 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00D3E7EF&
               Height          =   285
               Left            =   1530
               Locked          =   -1  'True
               TabIndex        =   35
               Top             =   1350
               Width           =   5085
            End
            Begin VB.TextBox txtIssdQty 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00D3E7EF&
               Height          =   285
               Left            =   10125
               Locked          =   -1  'True
               TabIndex        =   34
               Top             =   1350
               Width           =   765
            End
            Begin VB.TextBox txtProcess 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00D3E7EF&
               Height          =   285
               Left            =   6615
               Locked          =   -1  'True
               TabIndex        =   33
               Top             =   1350
               Width           =   3525
            End
            Begin VB.CommandButton cmdOK 
               Caption         =   "Ok"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   8175
               TabIndex        =   32
               Top             =   1920
               Width           =   1320
            End
            Begin VB.CommandButton cmdCancel 
               Caption         =   "Cancel"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   9570
               TabIndex        =   31
               Top             =   1920
               Width           =   1320
            End
            Begin VB.TextBox txtWastage 
               Alignment       =   2  'Center
               Height          =   300
               Left            =   75
               TabIndex        =   24
               Top             =   450
               Width           =   840
            End
            Begin MSForms.Label lblmine 
               Height          =   240
               Left            =   60
               TabIndex        =   37
               Top             =   1125
               Width           =   10830
               BackColor       =   11517387
               Caption         =   $"frmAuthorizeIssRcvAdv.frx":0277
               Size            =   "19103;423"
               BorderColor     =   16512
               BorderStyle     =   1
               FontName        =   "Tahoma"
               FontEffects     =   1073741825
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               FontWeight      =   700
            End
            Begin MSForms.ComboBox cmbRcvID 
               Height          =   300
               Left            =   8655
               TabIndex        =   30
               Top             =   450
               Width           =   2235
               VariousPropertyBits=   746604571
               DisplayStyle    =   3
               Size            =   "3942;529"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   178
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Iss ID"
               Height          =   195
               Index           =   3
               Left            =   8640
               TabIndex        =   29
               Top             =   225
               Width           =   420
            End
            Begin MSForms.ComboBox cmbProcess 
               Height          =   300
               Left            =   5115
               TabIndex        =   28
               Top             =   450
               Width           =   3525
               VariousPropertyBits=   746604571
               DisplayStyle    =   3
               Size            =   "6218;529"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   178
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Process"
               Height          =   195
               Index           =   2
               Left            =   5145
               TabIndex        =   27
               Top             =   225
               Width           =   555
            End
            Begin MSForms.ComboBox cmbMaker 
               Height          =   300
               Left            =   930
               TabIndex        =   26
               Top             =   450
               Width           =   4170
               VariousPropertyBits=   746604571
               DisplayStyle    =   3
               Size            =   "7355;529"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   178
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Maker"
               Height          =   195
               Index           =   1
               Left            =   990
               TabIndex        =   25
               Top             =   225
               Width           =   435
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Wastage"
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   23
               Top             =   225
               Width           =   645
            End
         End
         Begin VB.Frame FraWastageOwn 
            BackColor       =   &H00E7EBEF&
            Height          =   1515
            Left            =   5925
            TabIndex        =   38
            Top             =   840
            Visible         =   0   'False
            Width           =   6000
            Begin VB.TextBox txtReasonWastageOwn 
               Alignment       =   2  'Center
               Height          =   300
               Left            =   975
               TabIndex        =   43
               Top             =   450
               Width           =   4815
            End
            Begin VB.TextBox txtWastageOwn 
               Alignment       =   2  'Center
               Height          =   300
               Left            =   120
               TabIndex        =   41
               Top             =   450
               Width           =   840
            End
            Begin VB.CommandButton cmdCancelWastage 
               Caption         =   "Cancel"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   4485
               TabIndex        =   40
               Top             =   885
               Width           =   1320
            End
            Begin VB.CommandButton cmdOKWastage 
               Caption         =   "Ok"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   3090
               TabIndex        =   39
               Top             =   885
               Width           =   1320
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Reason"
               Height          =   195
               Index           =   5
               Left            =   990
               TabIndex        =   44
               Top             =   225
               Width           =   540
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Wastage"
               Height          =   195
               Index           =   4
               Left            =   135
               TabIndex        =   42
               Top             =   225
               Width           =   645
            End
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2430
            TabIndex        =   17
            Top             =   1800
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.PictureBox Pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            CausesValidation=   0   'False
            ClipControls    =   0   'False
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
            Left            =   165
            ScaleHeight     =   255
            ScaleWidth      =   945
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   810
            Visible         =   0   'False
            Width           =   945
            Begin VB.TextBox txtQty 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   8
               Top             =   0
               Width           =   950
            End
         End
         Begin MSComctlLib.ListView LVIss 
            Height          =   7050
            Left            =   -74910
            TabIndex        =   5
            Top             =   390
            Width           =   11280
            _ExtentX        =   19897
            _ExtentY        =   12435
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
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Reciept No"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Vend ID"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Vender"
               Object.Width           =   3881
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Item Code"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Item Description"
               Object.Width           =   3881
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Process"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Qty"
               Object.Width           =   1764
            EndProperty
         End
         Begin MSComctlLib.ListView LVRcv 
            Height          =   6375
            Left            =   90
            TabIndex        =   6
            Top             =   375
            Width           =   12015
            _ExtentX        =   21193
            _ExtentY        =   11245
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
            NumItems        =   13
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Reciept No"
               Object.Width           =   2408
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Vend ID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Vender"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Order No"
               Object.Width           =   1508
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Item Code"
               Object.Width           =   1508
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Lot No."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Item Description"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Process"
               Object.Width           =   1958
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   8
               Text            =   "Qty"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   9
               Text            =   "Repair Amt"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   10
               Text            =   "Waste"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   11
               Text            =   "Lost"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   12
               Text            =   "Re-Work"
               Object.Width           =   1411
            EndProperty
         End
         Begin MSComctlLib.ListView LVRM 
            Height          =   7065
            Left            =   -74910
            TabIndex        =   9
            Top             =   390
            Width           =   11310
            _ExtentX        =   19950
            _ExtentY        =   12462
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
               Text            =   "Date"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Vend ID"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Vender Name"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "RM ID"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "RM Name"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Rate"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Qty"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Amount"
               Object.Width           =   1764
            EndProperty
         End
      End
      Begin VB.Label Label10 
         Caption         =   "Inspector"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7080
         TabIndex        =   78
         Top             =   7170
         Width           =   1335
      End
      Begin VB.Image imgItem 
         BorderStyle     =   1  'Fixed Single
         Height          =   1650
         Left            =   180
         Stretch         =   -1  'True
         Top             =   7140
         Width           =   1950
      End
      Begin VB.Label Label2 
         Caption         =   "Purchase Order"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   10305
         TabIndex        =   16
         Top             =   7290
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   405
         Left            =   10950
         TabIndex        =   14
         Top             =   8385
         Width           =   1305
         ForeColor       =   0
         Caption         =   "Close      "
         PicturePosition =   327683
         Size            =   "2302;714"
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdDelete 
         Height          =   405
         Left            =   6855
         TabIndex        =   13
         Top             =   8385
         Width           =   1305
         ForeColor       =   0
         Caption         =   "Decline        "
         PicturePosition =   327683
         Size            =   "2302;714"
         Accelerator     =   68
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdUpdate 
         Height          =   405
         Left            =   5490
         TabIndex        =   10
         Top             =   8385
         Width           =   1305
         ForeColor       =   0
         Caption         =   "Authorize        "
         PicturePosition =   327683
         Size            =   "2302;714"
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "mnuPop"
      Begin VB.Menu mnuPostWastageToOtherMakers 
         Caption         =   "Post Wastage To Other Makers"
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReWork 
         Caption         =   "Add Re-Work"
      End
      Begin VB.Menu mnuDaSh2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddWastage 
         Caption         =   "Add Wastage"
      End
   End
End
Attribute VB_Name = "frmAuthorizeIssRcvAdv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iColNo As Integer
Dim rptSQL As String
Dim TableName As String
Dim lRcvEntryID As Long

Private Sub DeleteIssd()

    For i = 1 To LVIss.ListItems.count
        If LVIss.ListItems(i).Checked Then
            con.Execute "Delete From VendIssdDetail Where EntryID=" & Val(LVIss.ListItems(i).key)
        End If
    Next i
    'Now Clean VendIssued
    con.Execute "Delete From VendIssued Where EntryID Not In(Select RefID From VendIssdDetail)"
    
End Sub
Private Sub DeleteRcvd()

    For i = 1 To LVRcv.ListItems.count
        If LVRcv.ListItems(i).Checked Then
            con.Execute "Delete From VendRcvdDetail Where EntryID=" & Val(LVRcv.ListItems(i).key)
        End If
    Next i
    'Now Clean VendReceived
    con.Execute "Delete From VendReceived Where EntryID Not In(Select RefID From VendRcvdDetail)"
    
End Sub
Private Sub cmbRange_Click()

    DtTo = ServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    Select Case cmbRange.ListIndex
        Case 1
            DtFrom = ServerDate
        Case 2
            DtFrom = DateAdd("d", -7, ServerDate)
        Case 3
            DtFrom = DateAdd("d", -15, ServerDate)
        Case 4
            DtFrom = DateAdd("d", -30, ServerDate)
        Case 5
            DtFrom.Enabled = True
            DtTo.Enabled = True
            DtFrom.SetFocus
            Exit Sub
    End Select
    
End Sub

Private Sub cmbMaker_Change()
    Call cmbMaker_Click
End Sub

Private Sub cmbMaker_Click()
    If cmbMaker.MatchFound = False Then Exit Sub
    Call AddToCombo(cmbProcess, "Description", "VMakerAssignedProcesses", " WHERE VendID=" & cmbMaker.List(cmbMaker.ListIndex, 1), , "ProcessID")
End Sub

Private Sub cmbProcess_Change()
    Call cmbProcess_Click
End Sub

Private Sub cmbProcess_Click()

    If cmbProcess.MatchFound = False Then Exit Sub
    Call AddToCombo(cmbRcvID, "RecieptID", "VVendIssdDetail", " WHERE VendID=" & _
        cmbMaker.List(cmbMaker.ListIndex, 1) & " AND ProcessID=" & _
        cmbProcess.List(cmbProcess.ListIndex, 1) & " AND ItemCode='" & LVRcv.SelectedItem.SubItems(4) & "'", , "EntryID")
        
End Sub

Private Sub cmbRcvID_Change()
    Call cmbRcvID_Click
End Sub

Private Sub cmbRcvID_Click()
    
    On Error GoTo err
    
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM VVendIssued WHERE EntryID=" & cmbRcvID.List(cmbRcvID.ListIndex, 1), con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            txtItemCode = !ITemID & ""
            txtItemName = !ItemName & ""
            txtProcess = !Description & ""
            txtIssdQty = Val(!TotalIssQty & "")
        Else
            txtItemCode = "-"
            txtItemName = "-"
            txtProcess = "-"
            txtIssdQty = "-"
        End If
        .Close
    End With
    
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub cmdAddRW_Click()

    Dim i As Integer
    If cmbRepairs.MatchFound = False Then
        MsgBox "Please Select Repair Type.", vbInformation
        Exit Sub
    End If
    
    If Val(txtTotalQtyRW) + Val(txtQtyToRW) > Val(txtQtyRW) Then
        MsgBox "Invalid Qty.", vbInformation
        Exit Sub
    End If
    
    For i = 1 To LV.ListItems.count
        If Val(LV.ListItems(i).key) = Val(cmbRepairs.ID) Then
            MsgBox "Already Added.", vbInformation
            Exit Sub
        End If
    Next
    
    Dim ITM As ListItem
    Set ITM = LV.ListItems.add(, , cmbRepairs.Text)
    ITM.Tag = cmbRepairs.ID
    ITM.ListSubItems.add(, , Val(txtQtyToRW)).Tag = Val(LVRcv.SelectedItem.key)
    
    Call GetTotalRWQty
    txtQtyToRW = ""
    cmbRepairs.SetFocus
    
End Sub

Private Sub GetTotalRWQty()

    Dim i As Integer, lTotalRW As Long
    For i = 1 To LV.ListItems.count
        lTotalRW = lTotalRW + Val(LV.ListItems(i).SubItems(1))
    Next
    txtTotalQtyRW = lTotalRW
    
End Sub

Private Sub GetTotalWastageQty()

    Dim i As Integer, lTotalWastage As Long
    For i = 1 To LVWastage.ListItems.count
        lTotalWastage = lTotalWastage + Val(LVWastage.ListItems(i).SubItems(1))
    Next
    txtTotalQtyWastage = lTotalWastage
    
End Sub


Private Sub cmdAddWastage_Click()

    Dim i As Integer
    If cmbWastage.MatchFound = False Then
        MsgBox "Please Select Wastage Type.", vbInformation
        Exit Sub
    End If
    
    If Val(txtTotalQtyWastage) + Val(txtQtyToWastage) > Val(txtQtyWastage) Then
        MsgBox "Invalid Qty.", vbInformation
        Exit Sub
    End If
    
    For i = 1 To LVWastage.ListItems.count
        If Val(LVWastage.ListItems(i).Tag) = Val(cmbWastage.ID) Then
            MsgBox "Already Added.", vbInformation
            Exit Sub
        End If
    Next
    
    Dim iWastageType As Integer
    iWastageType = GetSingleLongValue("WastageType", "WastageTypes", " WHERE EntryID=" & cmbWastage.ID)
    
    If iWastageType = 3 Then    'Wastage Type 3 Means It should ask about Return Process...Just Like Return to Orig. but It'll be returned to the Selected Process.
        Dim f As New frmGetReturnProcessForWastage, lEntryID As Long, lVendID As Long, strEmpID As String
        lEntryID = Val(LVRcv.SelectedItem.key)
        Load f
        If f.getData(LVRcv.SelectedItem.SubItems(5), lEntryID, lVendID, strEmpID) = False Then Exit Sub
    End If
    
    Dim ITM As ListItem
    Set ITM = LVWastage.ListItems.add(, , cmbWastage.Text)
    ITM.Tag = cmbWastage.ID
    ITM.ListSubItems.add(, , Val(txtQtyToWastage)).Tag = Val(LVRcv.SelectedItem.key)
    
    If iWastageType = 3 Then
        ITM.ListSubItems.add(, , lEntryID).Tag = lVendID
        ITM.ListSubItems.add(, , strEmpID).Tag = lVendID
    Else
        ITM.ListSubItems.add(, , "").Tag = ""
        ITM.ListSubItems.add(, , "").Tag = ""
    End If
    
    Call GetTotalWastageQty
    txtQtyToWastage = ""
    cmbWastage.SetFocus
    
End Sub

Private Sub cmdCancel_Click()
    FraEdit.Visible = False
End Sub

Private Sub cmdCancelWastage_Click()
    FraWastageOwn.Visible = False
End Sub

Private Sub cmdCanelRW_Click()
    FraAddRework.Visible = False
End Sub

Private Sub cmdCanelWastageX_Click()
    FraAddWastage.Visible = False
End Sub

Private Sub cmdChkAll_Click()

    If SSTab1.Tab = 0 Then
        For i = 1 To LVIss.ListItems.count
            LVIss.ListItems(i).Checked = True
        Next i
    ElseIf SSTab1.Tab = 1 Then
        For i = 1 To LVRcv.ListItems.count
            LVRcv.ListItems(i).Checked = True
        Next i
    ElseIf SSTab1.Tab = 2 Then
        For i = 1 To LVRM.ListItems.count
            LVRM.ListItems(i).Checked = True
        Next i
    End If

End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub


Private Sub cmdDelete_Click()

    On Error GoTo err
        
'    Dim F As New frmAuthPwd
'    F.Show 1
'    If Not F.LoginSucceeded Then
'        Exit Sub
'    End If
    
    Call StartTrans(con)
    
'    Call DeleteIssd
    
    Call DeleteRcvd
    
'    Call DeleteRM
    
    con.CommitTrans
        
    Call RefreshLV
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub



Private Sub cmdOK_Click()

    If cmbRcvID.MatchFound = False Then Exit Sub
    LVRcv.SelectedItem.SubItems(10) = Val(txtWastage)
    LVRcv.SelectedItem.ListSubItems(9).Tag = cmbRcvID.List(cmbRcvID.ListIndex, 1)
    LVRcv.SelectedItem.ListSubItems(10).Tag = cmbMaker.List(cmbMaker.ListIndex, 1)
    FraEdit.Visible = False
    
End Sub

Private Sub cmdOKRW_Click()

    Dim i As Integer
    For i = LVRWAll.ListItems.count To 1 Step -1
        If LVRWAll.ListItems(i).ListSubItems(1).Tag = Val(LVRcv.SelectedItem.key) Then
            LVRWAll.ListItems.Remove i
        End If
    Next
    
    'Now Add these.
    Dim ITM As ListItem
    For i = 1 To LV.ListItems.count
        Set ITM = LVRWAll.ListItems.add(, , LV.ListItems(i).Text)
        ITM.Tag = LV.ListItems(i).Tag
        ITM.ListSubItems.add(, , LV.ListItems(i).SubItems(1)).Tag = LV.ListItems(i).ListSubItems(1).Tag
    Next
    
    FraAddRework.Visible = False
    LVRcv.SelectedItem.SubItems(12) = Val(txtTotalQtyRW)
    
End Sub

Private Sub cmdOKWastage_Click()

    LVRcv.SelectedItem.SubItems(9) = Val(txtWastageOwn)
    LVRcv.SelectedItem.ListSubItems(9).Tag = txtReasonWastageOwn
    LVRcv.SelectedItem.ListSubItems(10).Tag = "WastageOwn"
    FraWastageOwn.Visible = False
    
End Sub

Private Sub cmdOKWastageX_Click()

    Dim i As Integer
    For i = LVWastageAll.ListItems.count To 1 Step -1
        If LVWastageAll.ListItems(i).ListSubItems(1).Tag = Val(LVRcv.SelectedItem.key) Then
            LVWastageAll.ListItems.Remove i
        End If
    Next
    
    'Now Add these.
    Dim ITM As ListItem
    For i = 1 To LVWastage.ListItems.count
        Set ITM = LVWastageAll.ListItems.add(, , LVWastage.ListItems(i).Text)
        ITM.Tag = LVWastage.ListItems(i).Tag
        ITM.ListSubItems.add(, , LVWastage.ListItems(i).SubItems(1)).Tag = LVWastage.ListItems(i).ListSubItems(1).Tag
        ITM.ListSubItems.add(, , LVWastage.ListItems(i).SubItems(2)).Tag = LVWastage.ListItems(i).ListSubItems(2).Tag
        ITM.ListSubItems.add(, , LVWastage.ListItems(i).SubItems(3)).Tag = LVWastage.ListItems(i).ListSubItems(3).Tag
    Next
    
    FraAddWastage.Visible = False
    LVRcv.SelectedItem.SubItems(10) = Val(txtTotalQtyWastage)
    
End Sub

Private Sub cmdUnchkAll_Click()

    If SSTab1.Tab = 0 Then
        For i = 1 To LVIss.ListItems.count
            LVIss.ListItems(i).Checked = False
        Next i
    ElseIf SSTab1.Tab = 1 Then
        For i = 1 To LVRcv.ListItems.count
            LVRcv.ListItems(i).Checked = False
        Next i
    ElseIf SSTab1.Tab = 2 Then
        For i = 1 To LVRM.ListItems.count
            LVRM.ListItems(i).Checked = False
        Next i
    End If

End Sub

Private Sub cmdUpdate_Click()

    On Error GoTo err

'    Dim F As New frmAuthPwd
'    F.Show 1
'    If Not F.LoginSucceeded Then
'        Exit Sub
'    End If
    If cmbEmpID.MatchFound = False Then
        MsgBox "Please enter Inspector Name", vbCritical
        Exit Sub
    End If
    
    Dim cmd As New ADODB.Command
    
    Call StartTrans(con)
    
    'Call UpdateIssEntries
    
    Call UpdateRcvEntries
    
    'Call UpdateRMEntries

    
    con.CommitTrans
    
'    Call PrintRpts("Original")
    Call RefreshLV
    
    
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub UpdateRMEntries()

    For i = 1 To LVRM.ListItems.count
        With LVRM.ListItems(i)
            If .Checked Then
                con.Execute "Update RMRcvdItems Set ReqAuth=1 Where ID=" & Val(.key)
            End If
        End With
    Next
    
End Sub

Private Sub PrintRpts(strOriginal As String)

    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    
    Dim rs As New ADODB.Recordset
    
    Call StartTrans(con)
    'Insert Into PrintIssRcvRptsTemp.
    
    con.Execute "Delete From PrintAuthorized"
    con.Execute "Delete From PrintIssRcvRptsTemp"
    
    For i = 1 To LVIss.ListItems.count
        With LVIss.ListItems(i)
            If .Checked Then
                con.Execute "Insert Into PrintAuthorized(RecieptNo) Values('" & .Text & "')"
                con.Execute "Insert Into PrintIssRcvRptsTemp(RecieptID) Values('" & .Text & "')"
            End If
        End With
    Next
    
    For i = 1 To LVRcv.ListItems.count
        With LVRcv.ListItems(i)
            If .Checked Then
                con.Execute "Insert Into PrintAuthorized(RecieptNo) Values('" & .Text & "')"
                con.Execute "Insert Into PrintIssRcvRptsTemp(RecieptID) Values('" & .Text & "')"
            End If
        End With
    Next
    
    con.Execute "Insert Into PrintIssRcvRpts Select CompItemDesc,DT,RecieptID,ItemCode,VenderName" & _
                ",VendID1,ProcessID,Process,RcvdQty,Wastage,LostQty,UserID,AuthUserID,AuthDT,0 As IssEntry " & _
                "From VVendRcvItemsrpt Where RecieptID In(Select RecieptNo From PrintAuthorized)"
                
    con.Execute "Insert Into PrintIssRcvRpts Select CompItemDesc,DT,RecieptID,ItemCode,VenderName," & _
                "VendID1,ProcessID,Process,IssQty,0 As Wastage,0 As LostQty,UserID,AuthUserID,AuthDT,1 As IssEntry " & _
                "From VVendIssItemsrpt Where RecieptID In(Select RecieptNo From PrintAuthorized)"
    
    con.CommitTrans
    
    Dim rpt As CRAXDDRT.Report  'New rptVendIssDetail
    Set rpt = rptApp.OpenReport(RptPath & "\rptVendIssDetail.rpt") '
    Dim ForTxt As String
    If TableName = "VVendIssItemsrpt" Then
        ForTxt = "GATEPASS ISSUED"
    Else
        ForTxt = "GOODS RECEIVED"
    End If
    
    Dim FFDefs As FormulaFieldDefinitions
    Dim FFDef As CRAXDDRT.FormulaFieldDefinition
    Set FFDefs = rpt.FormulaFields
    
    For Each FFDef In FFDefs
'        If FFDef.Name = "{@Heading}" Then
'            FFDef.Text = "'" & ForTxt & "'"
        If FFDef.Name = "{@ForOriginal}" Then
            FFDef.Text = "'" & strOriginal & "'"
        End If
    Next
        
    Load frmPrevRpt
    frmPrevRpt.ShowReport " ", rpt
    
    Screen.MousePointer = vbDefault

    
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
End Sub

Private Sub UpdateRcvEntries()

    Dim cmd As New ADODB.Command
    Dim ProcID As Long
    
    Dim lWastage As Long, lWastageCh2Vender As Long, lWastageType As Long, lWastageCh2Originator As Long, lWastageCh2Selection As Long
    Dim lTempEntryID As Long, lmyVendID As Long, strmyEmpID As String, lChargeToEntryID As Long
    Dim lRepair As Long
    Dim lLost As Long
    Dim lReWork As Long, j As Integer
    Dim lReWorkEntryID As Long
    
    Dim strEmpID As String
    If cmbEmpID.MatchFound Then
        strEmpID = cmbEmpID.ID
    Else
        strEmpID = ""
    End If
    For i = 1 To LVRcv.ListItems.count
        If LVRcv.ListItems(i).Checked Then
        
            ProcID = LVRcv.ListItems(i).ListSubItems(3).Tag
            lRepair = Val(LVRcv.ListItems(i).SubItems(9))
            lWastage = Val(LVRcv.ListItems(i).SubItems(10))
            lLost = Val(LVRcv.ListItems(i).SubItems(11))
            lReWork = Val(LVRcv.ListItems(i).SubItems(12))
            
            con.Execute "UPDATE VendRcvdDetail SET ReqAuth=0,Wastage=" & lWastage & ",RepairAmt=" & _
             lRepair & ",LostQty=" & lLost & ",ReWorkQty=" & lReWork & ",Insp_EmpID='" & cmbEmpID.ID & _
             "' WHERE EntryID=" & Val(LVRcv.ListItems(i).key)
            
            If lReWork > 0 Then
                'Save Re-Work Status...
                For j = 1 To LVRWAll.ListItems.count
                    With LVRWAll.ListItems(j)
                        If Val(.ListSubItems(1).Tag) = Val(LVRcv.ListItems(i).key) Then
                        
                            con.Execute "INSERT INTO VendRcvdDetailReWorkDetail(VRD_RefID,Repair_RefID,Qty) VALUES(" & _
                             Val(LVRcv.ListItems(i).key) & "," & Val(.Tag) & "," & Val(.SubItems(1)) & ")"
                             
                            lReWorkEntryID = GetSingleLongValue("MAX(EntryID)", "VendRcvdDetailReWorkDetail")
                            
                            con.Execute "INSERT INTO VendRcvdDetailReWorkDetailProcesses(VRDRWD_RefID,ProcessID) " & _
                             "SELECT " & lReWorkEntryID & ",ProcessID FROM RepairTypeProcesses WHERE Repair_RefID=" & Val(.Tag)
                        End If
                        
                    End With
                Next
            End If
            
            If lWastage > 0 Then
                'Save Wastage Status....
                lWastageCh2Vender = 0
                lWastageCh2Originator = 0
                lWastageCh2Selection = 0
                Dim lReturnToEntryID As Long
                
                For j = 1 To LVWastageAll.ListItems.count
                    
                    With LVWastageAll.ListItems(j)
                    
                        If Val(.ListSubItems(1).Tag) = Val(LVRcv.ListItems(i).key) Then
                            
                            lWastageType = GetSingleLongValue("WastageType", "WastageTypes", " WHERE EntryID=" & Val(.Tag))
                            lReturnToEntryID = 0
                            If lWastageType = 1 Then
                                lWastageCh2Vender = lWastageCh2Vender + Val(.SubItems(1))
                                lmyVendID = Val(LVRcv.ListItems(i).Tag)
                            ElseIf lWastageType = 2 Then
                                lWastageCh2Originator = lWastageCh2Originator + Val(.SubItems(1))
                                lTempEntryID = GetSingleLongValue("MIN(EntryID)", "VendRcvdDetail", " WHERE LotNo='" & LVRcv.ListItems(i).SubItems(5) & "'")
                                lmyVendID = GetSingleLongValue("VendID", "VendReceived INNER JOIN VendRcvdDetail ON VendReceived.EntryID=VendRcvdDetail.RefID", " WHERE VendRcvdDetail.EntryID=" & lTempEntryID)
                            ElseIf lWastageType = 3 Then
                                lWastageCh2Selection = lWastageCh2Selection + Val(.SubItems(1))
                                lReturnToEntryID = Val(.SubItems(2))
                                lmyVendID = Val(.ListSubItems(2).Tag)
                            End If
                            
                            con.Execute "INSERT INTO VendRcvdDetailWastageDetail(VRD_RefID,Wastage_RefID,Qty,WastageType,ReturnTo_VRD_RefID,VendID,EmpID) VALUES(" & _
                             Val(LVRcv.ListItems(i).key) & "," & Val(.Tag) & "," & Val(.SubItems(1)) & "," & _
                             lWastageType & "," & lReturnToEntryID & "," & lmyVendID & ",'" & .SubItems(3) & "')"
                            
                        End If
                        
                    End With
                Next
            End If
            
            con.Execute "INSERT INTO VendRcvAuthDetail(UserID,RcvID,DT) Values(" & _
                UserID & "," & Val(LVRcv.ListItems(i).key) & ",'" & ServerDate & " " & ServerTime & "')"
                
            If LVRcv.ListItems(i).ListSubItems(10).Tag = "WastageOwn" Then
                con.Execute "INSERT INTO VendRcvdDetail_AuthDetails(VRD_RefID,WastageReason) VALUES(" & Val(LVRcv.ListItems(i).key) & _
                 ",'" & LVRcv.ListItems(i).ListSubItems(9).Tag & "')"
            End If
            
            With cmd
                .ActiveConnection = con
                .CommandType = adCmdStoredProc
                
'                If IsInitLotProcess(ProcID) Then
'                    .CommandText = "SP_InsertVendRcvdILP"
'                    '.Parameters("@OrderNo").value = LVRcv.ListItems(i).ListSubItems(1).Tag
'                    .Parameters("@ItemCode").Value = LVRcv.ListItems(i).SubItems(3)
'                    .Parameters("@ProcessID").Value = ProcID
'                    If Left(LVRcv.ListItems(i).Text, 3) = "WST" Then
'                        .Parameters("@Qty").Value = Val(LVRcv.ListItems(i).SubItems(8))
'                        .Parameters("@AddLedger").Value = 0
'                    Else
'                        .Parameters("@Qty").Value = Val(LVRcv.ListItems(i).SubItems(6))
'                    End If
'
'                    .Parameters("@RefID").Value = LVRcv.ListItems(i).ListSubItems(5).Tag
'                    .Parameters("@VendID").Value = LVRcv.ListItems(i).ListSubItems(2).Tag
'                    .Parameters("@LotNo").Value = Val(LVRcv.ListItems(i).ListSubItems(4).Tag & "")
'                    .Parameters("@RecID").Value = ""
'                    .Parameters("@ReqAuth").Value = 0   'Has To Pass Zero.
'                    .Parameters("@AuthorizeEntry").Value = 1
'                    .Execute
'
'                ElseIf RcvLot(ProcID) Then
'
'                    .CommandText = "SP_InsertVendRcvdLots"
'
'                    .Parameters("@ProcessID").Value = ProcID
'                    .Parameters("@Qty").Value = Val(LVRcv.ListItems(i).SubItems(6))
'
'                    .Parameters("@RefID").Value = LVRcv.ListItems(i).ListSubItems(5).Tag
'                    .Parameters("@VendID").Value = LVRcv.ListItems(i).ListSubItems(2).Tag
'                    .Parameters("@LotNo").Value = LVRcv.ListItems(i).ListSubItems(4).Tag
'                    .Parameters("@RecID").Value = ""
'                    .Parameters("@RcvDT").Value = ServerDate    'Just 4 Passing Not Used In The Procedure.
'                    .Parameters("@ReqAuth").Value = 0   'Has To Pass Zero.
'                    .Parameters("@AuthorizeEntry").Value = 1
'
'                    .Execute
                    
'                ElseIf ProcID = 4 Then
'                    .CommandText = "SP_AuthorizeFilling"
'                    .Parameters("@ItemCode").Value = LVRcv.ListItems(i).SubItems(3)
'                    .Parameters("@ProcessID").Value = ProcID
'                    .Parameters("@Qty").Value = Val(LVRcv.ListItems(i).SubItems(6))
'                    .Parameters("@Repair").Value = Val(LVRcv.ListItems(i).SubItems(7))
'                    .Parameters("@WasteQty").Value = Val(LVRcv.ListItems(i).SubItems(8) & "")
'                    .Parameters("@iRefID").Value = LVRcv.ListItems(i).ListSubItems(5).Tag
'                    .Parameters("@VendID").Value = Val(LVRcv.ListItems(i).ListSubItems(2).Tag & "")
'                    .Parameters("@LotNo").Value = 0
'                    .Parameters("@RecID").Value = ""
'                    .Parameters("@ReqAuth").Value = 0   'Has To Pass Zero.
'                    .Parameters("@AuthorizeEntry").Value = 1
    '                    .Parameters("@UserID").Value = UserID
'                    .Parameters("@UpdateEntryID").Value = Val(LVRcv.ListItems(i).key)
'                    .Parameters("@RcvDT").Value = ServerDate 'This Won't be Used It's Just 4 Passing
'                    .Execute
'                Else
                    '.CommandText = "SP_InsertVendRcvd"
                    .CommandText = "SP_InsertVendRcvdNew"
                    .Parameters("@OrderNo").Value = LVRcv.ListItems(i).SubItems(3)
                    .Parameters("@ItemCode").Value = LVRcv.ListItems(i).SubItems(4)
                    .Parameters("@ProcessID").Value = ProcID 'LVRcv.ListItems(i).ListSubItems(6).Tag    'RcvProcessID
                    .Parameters("@Qty").Value = Val(LVRcv.ListItems(i).SubItems(8))
                    .Parameters("@WasteQty").Value = lWastage 'Val(LVRcv.ListItems(i).SubItems(9) & "")
                    .Parameters("@RepairQty").Value = lRepair
                    .Parameters("@LostQty").Value = lLost
                    .Parameters("@ReWorkQty").Value = lReWork
                    .Parameters("@iRefID").Value = LVRcv.ListItems(i).ListSubItems(5).Tag
                    .Parameters("@VendID").Value = LVRcv.ListItems(i).Tag 'Val(LVRcv.ListItems(i).ListSubItems(2).Tag & "")
                    .Parameters("@LotNo").Value = Val(LVRcv.ListItems(i).ListSubItems(4).Tag)
                    .Parameters("@RecID").Value = ""
                    .Parameters("@ReqAuth").Value = 0   'Has To Pass Zero.
                    .Parameters("@AuthorizeEntry").Value = 1
                    .Parameters("@UserID").Value = UserID
                    .Parameters("@RcvDT").Value = ServerDate 'This Won't be Used It's Just 4 Passing
                    .Parameters("@DeliveryRefID").Value = Val(LVRcv.ListItems(i).ListSubItems(2).Tag)
                    .Parameters("@CountedBy").Value = "" 'This Won't be Used It's Just 4 Passing
                    .Parameters("@Issue_RefID").Value = Val(LVRcv.ListItems(i).ListSubItems(1).Tag)
                    .Parameters("@Weight").Value = Val(LVRcv.ListItems(i).ListSubItems(6).Tag)
                    .Parameters("@VRD_EntryID").Value = Val(LVRcv.ListItems(i).key)
                    .Execute
'                End If
            End With
            'Post Wastage to Provious Processes...
            If strCompany = "Weldon Instruments" Then
                If lWastage > 0 Or lRepair > 0 Then '
                
                    Dim lToProcessID As Long
                    lToProcessID = 0
                    If ProcID = 12 Then 'Setting
                        lToProcessID = iSemiFinishProcessID
                    ElseIf ProcID = 16 Then
                        lToProcessID = iReadyFinishProcessID
                    Else
                        lToProcessID = GetSingleLongValue("WastageToProcessID", "ProcessesWastage", " WHERE ProcessID=" & ProcID)
                    End If
                    If lToProcessID > 0 Then
                        Dim rs As New ADODB.Recordset
                        With rs
                            .Open "SELECT EntryID,VendID,Rate,AccNo FROM VVendIssdDetail WHERE ProcessID=" & lToProcessID & " AND LotNo=" & Val(LVRcv.ListItems(i).ListSubItems(4).Tag), con, adOpenForwardOnly, adLockReadOnly
                            If .EOF = False Then
                                'Post Voucher
                                If Val(!VendID & "") <> iFactoryMakerID Then
                                    Dim dAmt As Double, strRepairAccNo As String, strVchrNo As String
                                    strRepairAccNo = "41-004-13001"
                                    dAmt = lWastage * Val(!Rate & "")
'                                    Dim objVchr As New clsVouchers
'                                    objVchr.VoucherType = "JV"
'                                    objVchr.VoucherDate = GetServerDate(False)
'                                    objVchr.AddVoucherHead !AccNo, dAmt, "Repair Charges: Lot No:" & Val(LVRcv.ListItems(i).ListSubItems(4).Tag)
'                                    objVchr.AddVoucherHead strRepairAccNo, -dAmt, "Repair Charges: Lot No:" & Val(LVRcv.ListItems(i).ListSubItems(4).Tag)
'                                    strVchrNo = objVchr.PostVoucher(False, True)
                                
                                    con.Execute "INSERT INTO MakerRepair(VendID,DT,Rcvd_RefID,Qty,Rate,UserName,MachineName,VendIssdDetail_RefID,RepairAmt) VALUES(" & _
                                     Val(!VendID & "") & ",'" & GetServerDate(False) & "'," & Val(LVRcv.ListItems(i).key) & "," & lWastage & _
                                     "," & Val(!Rate & "") & ",'" & CurrentUserName & "','" & strComputerName & "'," & Val(!EntryID & "") & "," & lRepair & ")"
                                End If
                            End If
                            .Close
                        End With
                    End If
                End If
            ElseIf strCompany = "QEL" Or strCompany = "Dr-Frgz" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                Dim lmyEntryID As Long, dmyRate As Double
                lTempEntryID = Val(LVRcv.ListItems(i).key)
                lmyEntryID = GetSingleLongValue("Issue_RefID", "VendRcvdDetail", " WHERE EntryID=" & lTempEntryID)
                If (lWastageCh2Vender > 0 Or lRepair > 0) And Val(LVRcv.ListItems(i).key) > 0 Then  '
                    lmyVendID = Val(LVRcv.ListItems(i).Tag)
                    dmyRate = GetSingleDoubleValue("Rate", "VVendIssdDetail", " WHERE EntryID=" & lmyEntryID) 'Doubtfull Line
                    con.Execute "INSERT INTO MakerRepair(VendID,DT,Rcvd_RefID,Qty,Rate,UserName,MachineName,VendIssdDetail_RefID,RepairAmt,VRD_RefID_ChargeTo) VALUES(" & _
                     lmyVendID & ",'" & GetServerDate(False) & "'," & Val(LVRcv.ListItems(i).key) & "," & lWastageCh2Vender & _
                     "," & dmyRate & ",'" & CurrentUserName & "','" & strComputerName & "'," & lmyEntryID & "," & lRepair & "," & lTempEntryID & ")"
                End If
                
                If lWastageCh2Originator > 0 Or lWastageCh2Selection > 0 Then
                    Dim lVendIssdDetail As Long, iShowInBilling As Integer
                    Dim lMakerRepair_EntryID As Long, lRcvMoreDetails_EntryID As Long
                    Dim lIssRcvBalance As Long
                End If
                
                If lWastageCh2Originator > 0 Then
                    
                    lTempEntryID = GetSingleLongValue("MIN(EntryID)", "VendRcvdDetail", " WHERE LotNo='" & LVRcv.ListItems(i).SubItems(5) & "'")
                    lmyVendID = 0 'How to Find Originator... got it
                    lmyVendID = GetSingleLongValue("VendID", "VendReceived INNER JOIN VendRcvdDetail ON VendReceived.EntryID=VendRcvdDetail.RefID", " WHERE VendRcvdDetail.EntryID=" & lTempEntryID)
                    lVendIssdDetail = GetSingleLongValue("Issue_RefID", "VendReceived INNER JOIN VendRcvdDetail ON VendReceived.EntryID=VendRcvdDetail.RefID", " WHERE VendRcvdDetail.EntryID=" & lTempEntryID)
                    dmyRate = GetSingleDoubleValue("Rate", "VendReceived INNER JOIN VendRcvdDetail ON VendReceived.EntryID=VendRcvdDetail.RefID", " WHERE VendRcvdDetail.EntryID=" & lTempEntryID)

                    'Check If this Rcving Entry has been posted in bill or not...
                    
                    If GetSingleLongValue("COUNT(VRD_RefID)", "MakerPostedBillsDetail_Receivings", " WHERE VRD_RefID=" & lTempEntryID) > 0 Then
                        iShowInBilling = 1
                    Else
                        iShowInBilling = 0
                    End If
                    
                    con.Execute "INSERT INTO MakerRepair(VendID,DT,Rcvd_RefID,Qty,Rate,UserName,MachineName,VendIssdDetail_RefID,RepairAmt,Wastage_Originator,ShowInBilling,VRD_RefID_ChargeTo) VALUES(" & _
                     lmyVendID & ",'" & GetServerDate(False) & "'," & Val(LVRcv.ListItems(i).key) & "," & lWastageCh2Originator & _
                     "," & dmyRate & ",'" & CurrentUserName & "','" & strComputerName & "'," & lmyEntryID & ",0,1," & iShowInBilling & "," & lTempEntryID & ")"
                     
                    
                    lMakerRepair_EntryID = GetSingleLongValue("MAX(EntryID)", "MakerRepair")
                    'Now Stand this Qty for rcving...
                    
                    lRcvMoreDetails_EntryID = GetSingleLongValue("EntryID", "VendIssdDetail_MoreDetails", " WHERE VID_RefID=" & lVendIssdDetail)
                    
                    lIssRcvBalance = GetSingleLongValue("IssQty-RcvdQty-ISNULL(IssRcvBalance,0)-UnAuthRcvdQty+ISNULL(Wastage_Return,0)", "VMakerIssItems", " WHERE EntryID=" & lVendIssdDetail)
                    
                    If lRcvMoreDetails_EntryID = 0 Then
                        con.Execute "INSERT INTO VendIssdDetail_MoreDetails(VID_RefID,EmpID,Wastage_Return,IssRcvBalance) VALUES(" & _
                         lVendIssdDetail & ",''," & lWastageCh2Originator & "," & lIssRcvBalance & ")"
                         
                        lRcvMoreDetails_EntryID = GetSingleLongValue("MAX(EntryID)", "VendIssdDetail_MoreDetails")
                    Else
                        con.Execute "UPDATE VendIssdDetail_MoreDetails SET Wastage_Return=ISNULL(Wastage_Return,0)+" & lWastageCh2Originator & ",IssRcvBalance=" & lIssRcvBalance & " WHERE EntryID=" & lRcvMoreDetails_EntryID
                    End If
                    
                    con.Execute "INSERT INTO MakerRepair_VID_MD_Details(MR_RefID,VIS_MD_RefID,Qty) VALUES(" & lMakerRepair_EntryID & "," & lRcvMoreDetails_EntryID & "," & lWastageCh2Originator & ")"
                End If
                If lWastageCh2Selection > 0 Then
                
                    For j = 1 To LVWastageAll.ListItems.count
                    
                        With LVWastageAll.ListItems(j)
                        
                            If Val(.ListSubItems(1).Tag) = Val(LVRcv.ListItems(i).key) Then
                                lWastageType = GetSingleLongValue("WastageType", "WastageTypes", " WHERE EntryID=" & Val(.Tag))
                                If lWastageType = 3 Then
                                     
                                    lTempEntryID = Val(.SubItems(2))            'Rcving Entry
                                    lmyVendID = 0 'How to Find Originator... got it
                                    lmyVendID = Val(.ListSubItems(2).Tag)
                                    strmyEmpID = .SubItems(3)
                                                                                                             
                                    'Mark Rcvable---------------------------------------------------------
                                    lVendIssdDetail = GetSingleLongValue("Issue_RefID", "VendReceived INNER JOIN VendRcvdDetail ON VendReceived.EntryID=VendRcvdDetail.RefID", " WHERE VendRcvdDetail.EntryID=" & lTempEntryID)
                                    
                                    lRcvMoreDetails_EntryID = GetSingleLongValue("EntryID", "VendIssdDetail_MoreDetails", " WHERE VID_RefID=" & lVendIssdDetail)
                                    
                                    lIssRcvBalance = GetSingleLongValue("IssQty-RcvdQty-ISNULL(IssRcvBalance,0)-UnAuthRcvdQty+ISNULL(Wastage_Return,0)", "VMakerIssItems", " WHERE EntryID=" & lVendIssdDetail)
                                    
                                    If lRcvMoreDetails_EntryID = 0 Then
                                        
                                        If lVendIssdDetail > 0 Then
                                            con.Execute "INSERT INTO VendIssdDetail_MoreDetails(VID_RefID,EmpID,Wastage_Return,IssRcvBalance) VALUES(" & _
                                             lVendIssdDetail & ",''," & Val(.SubItems(1)) & "," & lIssRcvBalance & ")"
                                         
                                            lRcvMoreDetails_EntryID = GetSingleLongValue("MAX(EntryID)", "VendIssdDetail_MoreDetails")
                                        End If
                                    Else
                                        con.Execute "UPDATE VendIssdDetail_MoreDetails SET Wastage_Return=ISNULL(Wastage_Return,0)+" & Val(.SubItems(1)) & ",IssRcvBalance=" & lIssRcvBalance & " WHERE EntryID=" & lRcvMoreDetails_EntryID
                                    End If
                                    'Following Line is commented out, It's now after the following Block....
                                    'con.Execute "INSERT INTO MakerRepair_VID_MD_Details(MR_RefID,VIS_MD_RefID,Qty) VALUES(" & lMakerRepair_EntryID & "," & lRcvMoreDetails_EntryID & "," & Val(.SubItems(1)) & ")"
                                    'Mark Rcvable---------------------------------------------------------
                                    
                                    With rs
                                        .Open "SELECT * FROM VVendRcvdDetail WHERE LotNo='" & LVRcv.ListItems(i).SubItems(5) & "' AND EntryID<=" & Val(LVRcv.ListItems(i).key) & " ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
                                        Do Until .EOF
                                        
                                            lVendIssdDetail = GetSingleLongValue("Issue_RefID", "VendReceived INNER JOIN VendRcvdDetail ON VendReceived.EntryID=VendRcvdDetail.RefID", " WHERE VendRcvdDetail.EntryID=" & Val(!EntryID & ""))
                                            
                                            dmyRate = GetSingleLongValue("Rate", "VendReceived INNER JOIN VendRcvdDetail ON VendReceived.EntryID=VendRcvdDetail.RefID", " WHERE VendRcvdDetail.EntryID=" & Val(!EntryID & ""))
                                                     
                                            If GetSingleLongValue("COUNT(VRD_RefID)", "MakerPostedBillsDetail_Receivings", " WHERE VRD_RefID=" & Val(!EntryID & "")) > 0 Then
                                                iShowInBilling = 1
                                            Else
                                                iShowInBilling = 0
                                            End If
                                            
                                            If Val(!VendID & "") = lmyVendID Then
                                                lChargeToEntryID = Val(!EntryID & "")
                                            Else
                                                lChargeToEntryID = 0
                                                iShowInBilling = 1
                                            End If
                                            
                                            If strmyEmpID <> "" Then
                                                lChargeToEntryID = 0
                                            End If
                                            
                                            con.Execute "INSERT INTO MakerRepair(VendID,DT,Rcvd_RefID,Qty,Rate,UserName,MachineName,VendIssdDetail_RefID,RepairAmt,Wastage_Originator,ShowInBilling,VRD_RefID_ChargeTo,EmpID,VRD_RefID) VALUES(" & _
                                             lmyVendID & ",'" & GetServerDate(False) & "'," & Val(LVRcv.ListItems(i).key) & "," & Val(LVWastageAll.ListItems(j).SubItems(1)) & _
                                             "," & dmyRate & ",'" & CurrentUserName & "','" & strComputerName & "'," & lmyEntryID & ",0,0," & iShowInBilling & "," & lChargeToEntryID & ",'" & strmyEmpID & "'," & Val(!EntryID & "") & ")"
                                              
                                            If Val(!EntryID & "") = lTempEntryID Then
                                                lMakerRepair_EntryID = GetSingleLongValue("MAX(EntryID)", "MakerRepair")
                                            End If
                                            
                                            .MoveNEXT
                                        Loop
                                        .Close
                                    End With
                                    If lMakerRepair_EntryID > 0 Then
                                        con.Execute "INSERT INTO MakerRepair_VID_MD_Details(MR_RefID,VIS_MD_RefID,Qty) VALUES(" & lMakerRepair_EntryID & "," & lRcvMoreDetails_EntryID & "," & Val(.SubItems(1)) & ")"
                                    End If
                                End If
                                
                            End If
                            
                        End With
                    Next

                End If
                
            End If
        End If
        
    Next i
    
    If lRcvEntryID <> 0 Then
        Unload Me
    End If
    
End Sub

Private Sub UpdateIssEntries()

    Dim cmd As New ADODB.Command
    Dim IssLot As Boolean
    
    With cmd
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_UpdateForIAR"
        For i = 1 To LVIss.ListItems.count    'Update Rcv Quantities.
        
            If LVIss.ListItems(i).Checked Then
            
                IssLot = RcvLot(LVIss.ListItems(i).ListSubItems(3).Tag) 'Whether Or Not Issuing Lot.
                
                con.Execute "Update VendIssdDetail Set ReqAuth=0 Where EntryID=" & Val(LVIss.ListItems(i).key)
                'Insert Into Iss Auth Detail Table
                con.Execute "Insert Into VendIssAuthDetail(UserID,IssID,DT) Values(" & _
                UserID & "," & Val(LVIss.ListItems(i).key) & ",'" & ServerDate & " " & ServerTime & "')"
                
                If Not IssLot Then
                    '.Parameters("@OrderNo").value = LVIss.ListItems(i).ListSubItems(1).Tag
                    .Parameters("@ItemCode").Value = LVIss.ListItems(i).SubItems(3)
                    .Parameters("@ProcessID").Value = LVIss.ListItems(i).ListSubItems(3).Tag
                    .Parameters("@Qty").Value = Val(LVIss.ListItems(i).SubItems(6))
                    .Execute
                Else
                    con.Execute "Update VendRcvdDetail Set IssQty=RcvdQty Where LotNo=" & LVIss.ListItems(i).ListSubItems(4).Tag & " and NextProcessID=" & LVIss.ListItems(i).ListSubItems(3).Tag
                End If
                
            End If
        Next i
    End With
    
End Sub
Private Sub DtFrom_Change()
    'Call setDateRange(DtFrom)
End Sub

Private Sub DtTo_Change()
    'Call setDateRange(DtTo)
End Sub


Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
        
    mnuPop.Visible = False
    SSTab1.TabVisible(0) = False
    SSTab1.TabVisible(2) = False
    SSTab1.Tab = 1
    iColNo = 0
    
    
    For i = 1 To LVRcv.ColumnHeaders.count
        cmbfield.AddItem LVRcv.ColumnHeaders(i).Text
    Next
    
    
    cmbRepairs.AddVals con, "'{' + ISNULL(RepairCode,'') + '} ' + RepairType + ' (' + ISNULL(DefectType,'') + ')'", "RepairTypes", "EntryID"
    cmbWastage.AddVals con, "WastageName", "WastageTypes", "EntryID", " WHERE Closed=0"
        
    lRcvEntryID = 0
    
    With cmbEmpID
        .ListHeight = 1600
        .AddVals con, "'{' + EmpID + '} ' + Name + ' {' + Designation + '}'", "Employees", "EmpID"
    End With
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Dim TheHeight As Long
    Frame1.Left = (Me.ScaleWidth - Frame1.Width) / 2
    
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    
End Sub


Public Sub ShowForIss(Optional ForRcv As Boolean)
    If ForRcv Then
        TableName = "VVendRcvItemsrpt"
    Else
        TableName = "VVendIssItemsrpt"
    End If
    Me.Show
End Sub

Private Sub RefreshLV()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim Cond As String
    Dim DTRange As String
    Dim ITM As ListItem
    
    Screen.MousePointer = vbHourglass
    Dim lColor As Long
        
    With rs
'        .Open "Select * From VVendIssItemsrpt Where (ReqAuth=1)  Order By ProcessID Desc,DT Asc", con, adOpenForwardOnly, adLockReadOnly
'        LVIss.ListItems.Clear
'
'        Do Until .EOF
'            If Left(![RecieptID] & "", 2) = "HS" Then lColor = LVIss.ForeColor Else lColor = vbRed
'            Set ITM = LVIss.ListItems.add(, ![EntryID] & "'", ![RecieptID] & "")
'            ITM.ForeColor = lColor
'            ITM.ListSubItems.add(, , ![VendID1] & "").ToolTipText = ![VendID1] & ""
'            ITM.ListSubItems(1).ForeColor = lColor
'            ITM.ListSubItems.add(, , ![VenderName] & "").ToolTipText = ![VenderName] & ""
'            ITM.ListSubItems(2).Tag = ![VendID] & ""
'            ITM.ListSubItems(2).ForeColor = lColor
'            ITM.ListSubItems.add(, , ![ItemCode] & "").ToolTipText = ![ItemCode] & ""
'            ITM.ListSubItems(3).Tag = ![ProcessID] & ""
'            ITM.ListSubItems(3).ForeColor = lColor
'            ITM.ListSubItems.add(, , ![ItemDescription] & "").ToolTipText = ![ItemDescription] & ""
'            ITM.ListSubItems(4).Tag = ![LotNo] & ""
'            ITM.ListSubItems(4).ForeColor = lColor
'            ITM.ListSubItems.add(, , ![Process] & "").ToolTipText = ![Process] & ""
'            ITM.ListSubItems(5).ForeColor = lColor
'            ITM.ListSubItems.add(, , ![IssQty] & "").ForeColor = lColor
'            .MoveNext
'        Loop
'        .Close
        
        '.Open "Select * From VVendRcvItemsrpt Where (ReqAuth=1)  and VendID Not In(Select VendID From Venders Where VendType=0)  Order By ProcessID Desc,DT Asc", con, adOpenForwardOnly, adLockReadOnly
        If lRcvEntryID = 0 Then
            .Open "SELECT * FROM VVendRcvItemsrpt WHERE (ReqAuth=1) ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        Else
            '.Open "SELECT * FROM VVendRcvItemsrpt WHERE EntryID=" & lRcvEntryID & " AND (ReqAuth=1) ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
            .Open "SELECT * FROM VVendRcvItemsrpt WHERE RefID=" & lRcvEntryID & " AND (ReqAuth=1) ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        End If
        LVRcv.ListItems.Clear
        Do Until .EOF
            If Left(![RecieptID] & "", 2) = "HS" Then lColor = LVRcv.ForeColor Else lColor = vbRed
            Set ITM = LVRcv.ListItems.add(, ![EntryID] & "'", ![RecieptID] & "")
            ITM.ForeColor = lColor
            ITM.ListSubItems.add(, , ![VendID1] & "").ToolTipText = ![VendID1] & ""
            If ![VenderName] & "" = "" Then
                ITM.ListSubItems.add(, , ![Supervisor] & "").ToolTipText = ![Supervisor] & ""
            Else
                ITM.ListSubItems.add(, , ![VendID1] & "").ToolTipText = ![VenderName] & ""
                ITM.Tag = ![VendID] & ""
            End If
            ITM.ListSubItems.add(, , ![OrderNo] & "").ToolTipText = ![OrderNo] & ""
            ITM.ListSubItems(1).Tag = Val(!Issue_RefID & "") 'VendIssdDetail.Issue_RefID
            ITM.ListSubItems.add(, , ![ItemCode] & "").ToolTipText = ![ItemCode] & ""
            ITM.ListSubItems(2).Tag = Val(!Issuance_RefID & "") 'VendIssued.Issuance_RefID
            ITM.ListSubItems(3).Tag = ![ProcessID] & ""
            ITM.ListSubItems.add(, , ![LotNo] & "").ToolTipText = ![LotNo] & ""
            ITM.ListSubItems.add(, , ![ItemDescription] & "").ToolTipText = ![ItemDescription] & ""
            ITM.ListSubItems(4).Tag = ![LotNo] & ""
            ITM.ListSubItems.add(, , ![Process] & "").ToolTipText = ![Process] & ""
            ITM.ListSubItems(5).Tag = ![RefID] & ""
            ITM.ListSubItems.add , , ![RcvdQty] & ""
            'ITM.ListSubItems(6).Tag = ![RcvProcessID] & ""
            ITM.ListSubItems.add , , "0"
            ITM.ListSubItems.add , , ![Wastage] & ""
            ITM.ListSubItems.add , , ![LostQty] & ""
            ITM.ListSubItems.add , , Val(!ReWorkQty & "")    'Re-Work
            ITM.ListSubItems(6).Tag = ![RcvdWeight] & ""
            ITM.ListSubItems(7).Tag = ![Unit] & ""
            ITM.ListSubItems(8).Tag = ![IssRecieptID] & ""
            
            For i = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(i).ForeColor = lColor
            Next
            .MoveNEXT
        Loop
        .Close
        
'        .Open "Select * From VRMRcvdDetail Where IsNull(ReqAuth,0)=1", con, adOpenForwardOnly, adLockReadOnly
'        LVRM.ListItems.Clear
'        Do Until .EOF
'            Set ITM = LVRM.ListItems.add(, !EntryID & "'", Format(!DT, "dd-MMM-yyyy"))
'            ITM.SubItems(1) = !VendID1 & ""
'            ITM.SubItems(2) = !VenderName & ""
'            ITM.SubItems(3) = !RMID1 & ""
'            ITM.SubItems(4) = !RMName & ""
'            ITM.SubItems(5) = Val(!Rate & "")
'            ITM.SubItems(6) = Val(!Qty & "")
'            ITM.SubItems(7) = Val(!Rate & "") * Val(!Qty & "")
'            .MoveNext
'        Loop
'        .Close
    End With
    
    Set rs = Nothing
'    Call RefreshPO
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Private Sub RefreshPO()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim Cond As String
    Dim DTRange As String
    Dim ITM As ListItem
    
    Screen.MousePointer = vbHourglass
    
    
    With rs
        
        '.Open "Select * From VVendRcvItemsrpt Where (ReqAuth=1) and (ProcessID=7 OR ProcessID=12) and (Left(RecieptID,5)='HSPOL' OR Left(RecieptID,5)='HSPAC') Order By DT", con, adOpenForwardOnly, adLockReadOnly
        .Open "Select * From VVendRcvItemsrpt Where (ReqAuth=1) and VendID In(Select VendID From Venders Where VendType=0) Order By DT", con, adOpenForwardOnly, adLockReadOnly
        LVPO.ListItems.Clear
        'LV.Sorted = False
        Do Until .EOF
            Set ITM = LVPO.ListItems.add(, ![EntryID] & "'", ![RecieptID] & "")
            ITM.ListSubItems.add(, , ![VendID1] & "").ToolTipText = ![VendID1] & ""
            'itm.ListSubItems(1).Tag = ![OrderNo] & ""
            If ![VenderName] & "" = "" Then
                ITM.ListSubItems.add(, , ![Supervisor] & "").ToolTipText = ![Supervisor] & ""
            Else
                ITM.ListSubItems.add(, , ![VenderName] & "").ToolTipText = ![VenderName] & ""
                ITM.ListSubItems(2).Tag = ![VendID] & ""
            End If
            ITM.ListSubItems.add(, , ![ItemCode] & "").ToolTipText = ![ItemCode] & ""
            ITM.ListSubItems(3).Tag = ![ProcessID] & ""
            ITM.ListSubItems.add(, , ![ItemDescription] & "").ToolTipText = ![ItemDescription] & ""
            ITM.ListSubItems(4).Tag = ![LotNo] & ""
            ITM.ListSubItems.add(, , ![Process] & "").ToolTipText = ![Process] & ""
            ITM.ListSubItems(5).Tag = ![RefID] & ""
            ITM.SubItems(6) = ![RcvdQty] & ""
            ITM.SubItems(7) = ![Wastage] & ""
            ITM.SubItems(8) = ![LostQty] & ""
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

Private Sub RefreshSnos()
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Text = i
    Next i
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)

    If LV.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo) = vbYes Then
            LV.ListItems.Remove LV.SelectedItem.Index
        End If
    End If
    
End Sub

Private Sub LVIss_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    'Do Nothing....
End Sub

Private Sub LVRcv_DblClick()
    
    If iColNo < 10 Then Exit Sub
    
    If LVRcv.ListItems.count = 0 Then Exit Sub
    'If Val(LVRcv.SelectedItem.ListSubItems(3).Tag) <> 4 Then Exit Sub
    
    If iColNo = 11 Then
'        txtWastageOwn = Val(LVRcv.SelectedItem.SubItems(9))
'        txtReasonWastageOwn.Text = LVRcv.SelectedItem.ListSubItems(9).Tag
'        FraWastageOwn.Visible = True
'        Exit Sub
        Call mnuAddWastage_Click
        Exit Sub
    ElseIf iColNo = 13 Then
        Call mnuReWork_Click
        Exit Sub
    End If
    
    With txtEdit
        .Move LVRcv.Left + LVRcv.ColumnHeaders(iColNo).Left, LVRcv.SelectedItem.Top + LVRcv.Top, LVRcv.ColumnHeaders(iColNo).Width
        .Visible = True
    End With

    With txtEdit
        .Text = Val(LVRcv.SelectedItem.SubItems(iColNo - 1))
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
   
    
End Sub


Private Sub LVRcv_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'
'    If Item.Checked = False Then Exit Sub
'    Item.Selected = True
'    If Val(LVRcv.SelectedItem.ListSubItems(3).Tag) <> 4 Then Exit Sub
'
'    With Pic
'        .Move LVRcv.Left + 10 + LVRcv.ColumnHeaders(8).Left, LVRcv.SelectedItem.Top + LVRcv.Top
'        .Visible = True
'    End With
'
'    With txtQty
'
'        .Text = Val(LVRcv.SelectedItem.SubItems(7))
'
'        .SelStart = 0
'        .SelLength = Len(.Text)
'        .SetFocus
'    End With
    
End Sub

Private Sub LVRcv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim strItemCode As String
    strItemCode = Item.SubItems(4)
    Dim rsPic As New ADODB.Recordset
    rsPic.Open "SELECT ItemID,ItemPic FROM Items WHERE ItemID='" & strItemCode & "'", con, adOpenForwardOnly, adLockReadOnly
    If Not IsNull(rsPic!ItemPic) Then
        Set imgItem.DataSource = rsPic
        imgItem.DataField = "ItemPic"
    Else
        imgItem.Picture = LoadPicture("")
    End If
    rsPic.Close
    
    FraEdit.Visible = False
    FraAddRework.Visible = False
    
End Sub

Private Sub LVRcv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Integer
    For i = 1 To LVRcv.ColumnHeaders.count
        If X > LVRcv.ColumnHeaders(i).Left And X < LVRcv.ColumnHeaders(i).Left + LVRcv.ColumnHeaders(i).Width Then
            iColNo = i
            Exit For
        End If
    Next
    
End Sub

Private Sub LVRcv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If strCompany <> "Sunlike" Then
        mnuPostWastageToOtherMakers.Visible = False
    Else
        mnuPostWastageToOtherMakers.Visible = True
    End If
    
    If Button <> 2 Then Exit Sub
    If LVRcv.ListItems.count = 0 Then Exit Sub
    
    Me.PopUpMenu mnuPop
    
End Sub

Private Sub mnuAddWastage_Click()

    txtItemCodeWastage = LVRcv.SelectedItem.SubItems(4)
    txtItemNameWastage = LVRcv.SelectedItem.SubItems(6)
    txtProcessWastage = LVRcv.SelectedItem.SubItems(7)
    txtQtyWastage = Val(LVRcv.SelectedItem.SubItems(8))
    '----------------------------------------------
    LVWastage.ListItems.Clear
    txtQtyToWastage = ""
    txtTotalQtyWastage = ""
    
    'Call Already Added.
    Dim i As Integer, ITM As ListItem
    For i = 1 To LVWastageAll.ListItems.count
        If Val(LVWastageAll.ListItems(i).ListSubItems(1).Tag) = Val(LVRcv.SelectedItem.key) Then
            'Add These...
            Set ITM = LVWastage.ListItems.add(, , LVWastageAll.ListItems(i).Text)
            ITM.Tag = LVWastageAll.ListItems(i).Tag
            ITM.ListSubItems.add(, , LVWastageAll.ListItems(i).SubItems(1)).Tag = LVWastageAll.ListItems(i).ListSubItems(1).Tag
            ITM.ListSubItems.add(, , LVWastageAll.ListItems(i).SubItems(2)).Tag = LVWastageAll.ListItems(i).ListSubItems(2).Tag
        End If
    Next
    Call GetTotalWastageQty
    FraAddWastage.Visible = True
    cmbWastage.SetFocus

End Sub

Private Sub mnuPostWastageToOtherMakers_Click()

    With FraEdit
        FraEdit.Visible = True
        If cmbMaker.ListCount = 0 Then
            Call AddToCombo(cmbMaker, "VenderName + ' ' + VendID1", "Makers", , , "VendID")
        End If
        txtWastage.SetFocus
    End With
    
End Sub

Private Sub mnuReWork_Click()
'    FraWastageOwn.Visible = True
'    txtWastageOwn.SetFocus

    txtItemCodeRW = LVRcv.SelectedItem.SubItems(4)
    txtItemNameRW = LVRcv.SelectedItem.SubItems(6)
    txtProcessRW = LVRcv.SelectedItem.SubItems(7)
    txtQtyRW = Val(LVRcv.SelectedItem.SubItems(8))
    '----------------------------------------------
    LV.ListItems.Clear
    txtQtyToRW = ""
    txtTotalQtyRW = ""
    
    'Call Already Added.
    Dim i As Integer, ITM As ListItem
    For i = 1 To LVRWAll.ListItems.count
        If Val(LVRWAll.ListItems(i).ListSubItems(1).Tag) = Val(LVRcv.SelectedItem.key) Then
            'Add These...
            Set ITM = LV.ListItems.add(, , LVRWAll.ListItems(i).Text)
            ITM.Tag = LVRWAll.ListItems(i).Tag
            ITM.ListSubItems.add(, , LVRWAll.ListItems(i).SubItems(1)).Tag = LVRWAll.ListItems(i).ListSubItems(1).Tag
        End If
    Next
    Call GetTotalRWQty
    FraAddRework.Visible = True
    cmbRepairs.SetFocus
End Sub

Private Sub txtBarcode_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
'        Dim i As Integer
'        For i = 1 To LVIss.ListItems.count
'            If txtBarcode.Text = LVIss.ListItems(i).Text Then
'                LVIss.ListItems(i).Checked = True
'                LVIss.ListItems(i).Selected = True
'                SSTab1.Tab = 0
'                LVIss.ListItems(i).EnsureVisible
'                Call LVIss_ItemCheck(LVIss.ListItems(i))
'                With txtBarcode
'                    .SelStart = 0
'                    .SelLength = Len(.Text)
'                    .SetFocus
'                End With
'                Exit Sub
'            End If
'        Next
        Dim iCounter As Integer, iLastIndex As Integer
        iCounter = 1
        iLastIndex = LVRcv.ListItems.count
        Do While iCounter <= iLastIndex
            i = i + 1
            If txtBarcode.Text <> Mid(LVRcv.ListItems(i).ListSubItems(8).Tag, 7) Then
                LVRcv.ListItems.Remove i
                i = i - 1
            End If
            iCounter = iCounter + 1
        Loop
        
'        For i = 1 To LVRcv.ListItems.count
'            If txtBarcode.Text <> LVRcv.ListItems(i).ListSubItems(8).Tag Then
'                'LVRcv.ListItems(i).Checked = True
'                'SSTab1.Tab = 1
'                'LVRcv.ListItems(i).EnsureVisible
'                'Call LVRcv_ItemCheck(LVRcv.ListItems(i))
''                With txtBarcode
''                    .SelStart = 0
''                    .SelLength = Len(.Text)
''                    .SetFocus
''                End With
''                Exit Sub
'                LVRcv.ListItems.Remove i
'                i = i - 1
'            End If
'        Next
'        MsgBox "Not Found.", vbInformation
'        With txtBarcode
'            .SelStart = 0
'            .SelLength = Len(.Text)
'            .SetFocus
'        End With
    End If
    
    
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(txtEdit) < 0 Then
            MsgBox "Invalid Qty.", vbInformation
            Exit Sub
        End If
        LVRcv.SelectedItem.ListSubItems(iColNo - 1).Text = Val(txtEdit)
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
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

Private Sub cmdFind_Click()

    Dim StartIdx As Integer
    Dim bStartFromTop As Boolean
    
    If LVRcv.SelectedItem Is Nothing Then
       StartIdx = 1
       bStartFromTop = True
    Else
       If LVRcv.SelectedItem.Index = LVRcv.ListItems.count Then
          StartIdx = 1
          bStartFromTop = True
       Else
          StartIdx = LVRcv.SelectedItem.Index
          bStartFromTop = False
       End If
    End If

    Dim fldValue As String

    For i = StartIdx To LVRcv.ListItems.count
       
       
       If cmbfield.ListIndex = 0 Then
          fldValue = Trim(Replace(LVRcv.ListItems(i), "»", ""))
       Else
          fldValue = Trim(Replace(LVRcv.ListItems(i).SubItems(cmbfield.ListIndex), "»", ""))
       End If
    
    
       If LCase(fldValue) Like "*" & LCase(Trim(txtFind)) & "*" Then
          LVRcv.ListItems(i).Selected = True
          LVRcv.ListItems(i).EnsureVisible
          LVRcv.SetFocus
          If i <> StartIdx Or bStartFromTop Then Exit For
       End If
       
    Next

    If i = LVRcv.ListItems.count + 1 Then
        MsgBox "Item Not Found.", vbInformation
        Set LVRcv.SelectedItem = Nothing
    End If
    
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)

    On Error GoTo err
    If KeyAscii = 13 Then
        If Val(txtQty) > Val(LVRcv.SelectedItem.SubItems(7)) + Val(LVRcv.SelectedItem.SubItems(9)) + Val(LVRcv.SelectedItem.SubItems(10)) Or Val(txtQty) < 0 Then
            MsgBox "Invalid Quantities!", vbInformation
            Exit Sub
        End If
        
        
        LVRcv.SelectedItem.SubItems(8) = Val(txtQty)
            
        Pic.Visible = False
                
    ElseIf KeyAscii = 27 Then
        Pic.Visible = False
    End If
    
    Call OnlyNums(KeyAscii, False)
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub DeletePORcvd()

    For i = 1 To LVPO.ListItems.count
    
        If LVPO.ListItems(i).Checked Then
        
            con.Execute "Delete From VendRcvdDetail Where EntryID=" & Val(LVPO.ListItems(i).key)
            
            'Update The Order's Rcvd Quantity.....
            con.Execute "Update  VendPurchOrdDetail Set VendPurchOrdDetail.QtyRcvd=VendPurchOrdDetail.QtyRcvd-" & Val(LVPO.ListItems(i).SubItems(6)) & _
             " From VendPurchOrdDetail Inner JOIN VendPORcvdDetail ON VendPORcvdDetail.OrderRefID" & _
             "=VendPurchOrdDetail.ID Inner JOIN VendRcvdDetail ON VendPORcvdDetail.RcvdRefID = " & _
             "VendRcvdDetail.EntryID Where VendRcvdDetail.EntryID=" & Val(LVPO.ListItems(i).key)
            'Update Rcving Detail Entry....
            con.Execute "Delete From VendPORcvdDetail Where RcvdRefID=" & Val(LVPO.ListItems(i).key)
'            con.Execute "Update VendPORcvdDetail Set VendPORcvdDetail.Qty=VendPORcvdDetail.Qty-" & Val(LVPO.ListItems(i).SubItems(7)) & _
'            " From VendPORcvdDetail Inner JOIN VendRcvdDetail ON VendPORcvdDetail.RcvdRefID=VendRcvdDetail.EntryID " & _
'            " Where VendRcvdDetail.EntryID=" & Val(LVPO.ListItems(i).Key)
'
        End If
        
    Next i
    
    'Now Clean VendReceived
    con.Execute "Delete From VendReceived Where EntryID Not In(Select RefID From VendRcvdDetail)"
    
End Sub

Private Sub UpdatePORcvd()

    Dim cmd As New ADODB.Command
    Dim ProcID As Long
    
    For i = 1 To LVPO.ListItems.count
        If LVPO.ListItems(i).Checked Then
            ProcID = LVPO.ListItems(i).ListSubItems(3).Tag
            
            con.Execute "Update VendRcvdDetail Set ReqAuth=0 Where EntryID=" & Val(LVPO.ListItems(i).key)
            
            'Update The Order's Rcvd Quantity.....
            con.Execute "Update  VendPurchOrdDetail Set VendPurchOrdDetail.QtyRcvd=VendPurchOrdDetail.QtyRcvd-" & Val(LVPO.ListItems(i).SubItems(7)) & _
             " From VendPurchOrdDetail Inner JOIN VendPORcvdDetail ON VendPORcvdDetail.OrderRefID" & _
             "=VendPurchOrdDetail.ID Inner JOIN VendRcvdDetail ON VendPORcvdDetail.RcvdRefID = " & _
             "VendRcvdDetail.EntryID Where VendRcvdDetail.EntryID=" & Val(LVPO.ListItems(i).key)
            'Update Rcving Detail Entry....
            con.Execute "Update VendPORcvdDetail Set VendPORcvdDetail.Qty=VendPORcvdDetail.Qty-" & Val(LVPO.ListItems(i).SubItems(7)) & _
            " From VendPORcvdDetail Inner JOIN VendRcvdDetail ON VendPORcvdDetail.RcvdRefID=VendRcvdDetail.EntryID " & _
            " Where VendRcvdDetail.EntryID=" & Val(LVPO.ListItems(i).key)
            
            con.Execute "Insert Into VendRcvAuthDetail(UserID,RcvID,DT) Values(" & _
                UserID & "," & Val(LVPO.ListItems(i).key) & ",'" & ServerDate & " " & ServerTime & "')"
                
            With cmd
                .ActiveConnection = con
                .CommandType = adCmdStoredProc
                
                .CommandText = "SP_AuthManRcving"
                .Parameters("@ItemCode").Value = LVPO.ListItems(i).SubItems(3)
                .Parameters("@ProcessID").Value = ProcID
                .Parameters("@Qty").Value = Val(LVPO.ListItems(i).SubItems(6))
                .Parameters("@Repair").Value = Val(LVPO.ListItems(i).SubItems(7))
                .Parameters("@WasteQty").Value = 0 'Val(LVPO.ListItems(i).SubItems(5) & "")
                .Parameters("@iRefID").Value = LVPO.ListItems(i).ListSubItems(5).Tag
                .Parameters("@VendID").Value = Val(LVPO.ListItems(i).ListSubItems(2).Tag & "")
                .Parameters("@LotNo").Value = 0
                .Parameters("@RecID").Value = ""
                .Parameters("@ReqAuth").Value = 0   'Has To Pass Zero.
                .Parameters("@AuthorizeEntry").Value = 1
                .Parameters("@UserID").Value = UserID
                .Parameters("@UpdateEntryID").Value = Val(LVPO.ListItems(i).key)
                .Parameters("@RcvDT").Value = ServerDate 'This Won't be Used It's Just 4 Passing
                .Execute
                
            End With
            
        End If
        
    Next i
    
End Sub


Private Sub DeleteRM()

    For i = 1 To LVRM.ListItems.count
        With LVRM.ListItems(i)
            If .Checked Then
            
                con.Execute "DELETE FROM RMRcvdItems WHERE ID=" & Val(.key)
'
            End If
        End With
    Next i
    
    'Now Clean VendReceived
    con.Execute "Delete From RMRcvd Where ID Not In(Select RefID From RMRcvdItems)"
    
End Sub

Public Function ShowMe(Optional p_lRcvEntryID As Long = 0) As Boolean
    lRcvEntryID = p_lRcvEntryID
    Call RefreshLV
    Me.Show
    Me.SetFocus
End Function
