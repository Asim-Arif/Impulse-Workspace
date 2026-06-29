VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmCreateDispatchList_Adv 
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
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleMode       =   0  'User
   ScaleWidth      =   10721.69
   WindowState     =   2  'Maximized
   Begin VB.Frame FraPrintIL 
      BackColor       =   &H00E7EBEF&
      Caption         =   "Print Inner Labels"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   5325
      Left            =   4710
      TabIndex        =   41
      Top             =   1875
      Visible         =   0   'False
      Width           =   3840
      Begin VB.TextBox txtTotalIL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   1005
         Width           =   1200
      End
      Begin VB.TextBox txtLabelsIL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1290
         TabIndex        =   51
         Top             =   1005
         Width           =   1200
      End
      Begin VB.TextBox txtQtyIL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   75
         TabIndex        =   50
         Top             =   1005
         Width           =   1200
      End
      Begin VB.TextBox txtBalanceIL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   450
         Width           =   1800
      End
      Begin VB.TextBox txtTotalQtyIL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   450
         Width           =   1800
      End
      Begin MSComctlLib.ListView LVIL 
         Height          =   3375
         Left            =   75
         TabIndex        =   42
         Top             =   1380
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   5953
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Qty"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Qty"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Labels"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Total"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Qty"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   2760
         TabIndex        =   53
         Top             =   810
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Labels"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   1590
         TabIndex        =   52
         Top             =   810
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Qty"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   525
         TabIndex        =   49
         Top             =   765
         Width           =   270
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Balance"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   2520
         TabIndex        =   47
         Top             =   210
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Qty"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   585
         TabIndex        =   45
         Top             =   210
         Width           =   675
      End
      Begin MSForms.CommandButton cmdPrintIL 
         Height          =   375
         Left            =   990
         TabIndex        =   44
         Top             =   4815
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Print"
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
      Begin MSForms.CommandButton cmdCancel 
         Height          =   375
         Left            =   2370
         TabIndex        =   43
         Top             =   4815
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Cancel"
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
      Height          =   8400
      Left            =   60
      TabIndex        =   0
      Top             =   -45
      Width           =   11865
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   75
         TabIndex        =   8
         Top             =   435
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
      Begin ComboList.Usercontrol1 cmbItems 
         Height          =   285
         Left            =   75
         TabIndex        =   3
         Top             =   945
         Width           =   8970
         _ExtentX        =   15822
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
      Begin VB.TextBox txtOrderNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9075
         TabIndex        =   32
         Top             =   945
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9090
         TabIndex        =   31
         Top             =   7935
         Visible         =   0   'False
         Width           =   1335
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6600
         Left            =   75
         TabIndex        =   16
         Top             =   1290
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   11642
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Lots Finalyzed"
         TabPicture(0)   =   "frmCreateDispatchList_Adv.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "LV"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtEdit"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "PicAddCartons"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "LVLots"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "LVLotsMain"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cmdUnSelectAll"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "cmdSelectAll"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "Dispatch List"
         TabPicture(1)   =   "frmCreateDispatchList_Adv.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label2(0)"
         Tab(1).Control(1)=   "Label1(2)"
         Tab(1).Control(2)=   "lblTotalCartons"
         Tab(1).Control(3)=   "Label1(3)"
         Tab(1).Control(4)=   "lblTotalPcs"
         Tab(1).Control(5)=   "cmbCountry"
         Tab(1).Control(6)=   "Label2(1)"
         Tab(1).Control(7)=   "Label2(2)"
         Tab(1).Control(8)=   "lblTotal"
         Tab(1).Control(9)=   "LVPList"
         Tab(1).Control(10)=   "txtDispatchListNo"
         Tab(1).Control(11)=   "cmbCustomer"
         Tab(1).Control(12)=   "cmdSelectAll1"
         Tab(1).Control(13)=   "cmdUnSelectAll1"
         Tab(1).ControlCount=   14
         Begin VB.CommandButton cmdUnSelectAll1 
            Caption         =   "Clear All"
            Height          =   255
            Left            =   -73710
            TabIndex        =   56
            Top             =   6285
            Width           =   1170
         End
         Begin VB.CommandButton cmdSelectAll1 
            Caption         =   "Select All"
            Height          =   255
            Left            =   -74895
            TabIndex        =   55
            Top             =   6285
            Width           =   1170
         End
         Begin ComboList.Usercontrol1 cmbCustomer 
            Height          =   285
            Left            =   -70620
            TabIndex        =   37
            Top             =   615
            Width           =   2310
            _ExtentX        =   4075
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
         Begin VB.CommandButton cmdSelectAll 
            Caption         =   "Select All"
            Height          =   300
            Left            =   105
            TabIndex        =   36
            Top             =   6000
            Width           =   1170
         End
         Begin VB.CommandButton cmdUnSelectAll 
            Caption         =   "Clear All"
            Height          =   300
            Left            =   1290
            TabIndex        =   35
            Top             =   6000
            Width           =   1170
         End
         Begin MSComctlLib.ListView LVLotsMain 
            Height          =   2085
            Left            =   2745
            TabIndex        =   34
            Top             =   2505
            Visible         =   0   'False
            Width           =   6030
            _ExtentX        =   10636
            _ExtentY        =   3678
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ItemCode"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Qty"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Process"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView LVLots 
            Height          =   2085
            Left            =   2745
            TabIndex        =   33
            Top             =   2085
            Visible         =   0   'False
            Width           =   6030
            _ExtentX        =   10636
            _ExtentY        =   3678
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "RcvingID"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "DispListIndex"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Qty"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "MainIndex"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.TextBox txtDispatchListNo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   -74910
            TabIndex        =   26
            Top             =   585
            Width           =   4245
         End
         Begin VB.PictureBox PicAddCartons 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   525
            Left            =   9270
            ScaleHeight     =   495
            ScaleWidth      =   2340
            TabIndex        =   19
            Top             =   5955
            Width           =   2370
            Begin VB.TextBox txtTotalQty 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1155
               Locked          =   -1  'True
               TabIndex        =   22
               Top             =   225
               Width           =   1200
            End
            Begin VB.TextBox txtCartonQty 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   -30
               TabIndex        =   21
               Top             =   225
               Width           =   1200
            End
            Begin VB.TextBox txtCartonFrom 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   -30
               Locked          =   -1  'True
               TabIndex        =   20
               Top             =   225
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "    Qty / Box        Total Qty"
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
               Index           =   30
               Left            =   -30
               TabIndex        =   23
               Top             =   -15
               Width           =   2385
            End
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   3120
            TabIndex        =   17
            Top             =   1275
            Visible         =   0   'False
            Width           =   780
         End
         Begin MSComctlLib.ListView LVPList 
            Height          =   5355
            Left            =   -74910
            TabIndex        =   24
            Top             =   930
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   9446
            SortKey         =   1
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
               Text            =   "#"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "To"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Qty"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Item ID"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Item Name"
               Object.Width           =   8290
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Carton No."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Lot"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "Internal Ord."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Process"
               Object.Width           =   2469
            EndProperty
         End
         Begin MSComctlLib.ListView LV 
            Height          =   5535
            Left            =   90
            TabIndex        =   18
            Top             =   420
            Width           =   11550
            _ExtentX        =   20373
            _ExtentY        =   9763
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
            NumItems        =   10
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S. #"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Maker"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Batch #"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Date"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Process"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Item ID"
               Object.Width           =   2117
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
               Text            =   "Done"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   8
               Text            =   "Available"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   9
               Text            =   "Lot #"
               Object.Width           =   1764
            EndProperty
         End
         Begin MSForms.TextBox lblTotal 
            Height          =   270
            Left            =   -71580
            TabIndex        =   57
            Top             =   6255
            Width           =   1515
            VariousPropertyBits=   746604575
            BackColor       =   15199215
            ForeColor       =   12582912
            BorderStyle     =   1
            Size            =   "2672;476"
            SpecialEffect   =   0
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
            Caption         =   "Country"
            Height          =   195
            Index           =   2
            Left            =   -68190
            TabIndex        =   40
            Top             =   375
            Width           =   585
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Customer"
            Height          =   195
            Index           =   1
            Left            =   -70620
            TabIndex        =   39
            Top             =   375
            Width           =   690
         End
         Begin MSForms.ComboBox cmbCountry 
            Height          =   285
            Left            =   -68265
            TabIndex        =   38
            Top             =   615
            Width           =   1860
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "3281;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblTotalPcs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E7EBEF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   -64920
            TabIndex        =   30
            Top             =   6255
            Width           =   1515
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Total Pcs :"
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
            Height          =   270
            Index           =   3
            Left            =   -66255
            TabIndex        =   29
            Top             =   6255
            Width           =   1350
         End
         Begin MSForms.TextBox lblTotalCartons 
            Height          =   270
            Left            =   -67755
            TabIndex        =   28
            Top             =   6255
            Width           =   1515
            VariousPropertyBits=   746604575
            BackColor       =   15199215
            BorderStyle     =   1
            Size            =   "2672;476"
            SpecialEffect   =   0
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
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Total Cartons :"
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
            Height          =   270
            Index           =   2
            Left            =   -69090
            TabIndex        =   27
            Top             =   6255
            Width           =   1350
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Dispatch List No. :"
            Height          =   195
            Index           =   0
            Left            =   -74895
            TabIndex        =   25
            Top             =   375
            Width           =   1305
         End
      End
      Begin VB.TextBox txtLotNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10065
         TabIndex        =   4
         Top             =   945
         Width           =   825
      End
      Begin VB.Frame Frame1 
         Height          =   480
         Left            =   2640
         TabIndex        =   11
         Top             =   7890
         Width           =   5475
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   1965
            TabIndex        =   14
            Top             =   135
            Width           =   2430
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4395
            TabIndex        =   13
            Top             =   135
            Width           =   1020
         End
         Begin VB.ComboBox cmbfield 
            Height          =   315
            ItemData        =   "frmCreateDispatchList_Adv.frx":0038
            Left            =   60
            List            =   "frmCreateDispatchList_Adv.frx":0051
            Style           =   2  'Dropdown List
            TabIndex        =   12
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
         ItemData        =   "frmCreateDispatchList_Adv.frx":0098
         Left            =   5265
         List            =   "frmCreateDispatchList_Adv.frx":00AE
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   435
         Width           =   1995
      End
      Begin VB.CommandButton cmdTransfer 
         Caption         =   "Save"
         Height          =   345
         Left            =   60
         TabIndex        =   2
         Top             =   7950
         Width           =   2205
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   9075
         TabIndex        =   6
         Top             =   435
         Width           =   1800
         _ExtentX        =   3175
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
         Format          =   152174595
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   7260
         TabIndex        =   7
         Top             =   435
         Width           =   1815
         _ExtentX        =   3201
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
         Format          =   152174595
         CurrentDate     =   37055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmCreateDispatchList_Adv.frx":00FE
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
         TabIndex        =   15
         Top             =   735
         Width           =   10785
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmCreateDispatchList_Adv.frx":01DD
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
         TabIndex        =   10
         Top             =   210
         Width           =   11130
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   810
         Left            =   10905
         TabIndex        =   9
         Top             =   435
         Width           =   840
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "1482;1429"
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
         Left            =   10455
         TabIndex        =   1
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
            Picture         =   "frmCreateDispatchList_Adv.frx":0298
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateDispatchList_Adv.frx":0454
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PopMenu"
      Begin VB.Menu mnuPrintInnerLabels 
         Caption         =   "Print Inner Labels"
      End
      Begin VB.Menu mnuPrintInnerLabelsII 
         Caption         =   "Print Inner Labels - II"
      End
      Begin VB.Menu mnuPrintInnerLabelsManualQty 
         Caption         =   "Print Inner Labels - Manual Qty"
      End
      Begin VB.Menu mnuPrintInnerLabelsManualQtyII 
         Caption         =   "Print Inner Labels - II Manual Qty"
      End
      Begin VB.Menu mnuAddToCarton 
         Caption         =   "Add to Carton"
      End
      Begin VB.Menu mnuDashDelete 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmCreateDispatchList_Adv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lDPListEntryID As Long
Dim strDelList As String

Public Sub ShowMe(Optional p_bCustomInvoice As Boolean = True)
    bCustomInvoice = p_bCustomInvoice
    Me.Show
End Sub

Public Sub ShowForShipping()
    bShipping = True
    Me.Show
End Sub

Private Sub cmbCustomer_matched()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    SQL = "Select Distinct Country From ForeignCustomers "
    
    'If cmbCust.ID <> "0" Then
    SQL = SQL & " Where CustCode='" & cmbCustomer.ID & "'"
    'End If
    
    With rs
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        
        cmbCountry.Clear
        'cmbCountry.AddItem "<All Countries>"
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            'cmbCountry.List(cmbCountry.ListCount - 1, 1) = ![Curr] & ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    If cmbCountry.ListCount > 0 Then cmbCountry.ListIndex = 0
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
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

Private Sub cmdCancel_Click()
    FraPrintIL.Visible = False
    FAcc.Enabled = True
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdFind_Click()
    
    Dim myLV As ListView, iItemIDIndex As Integer
    
    If SSTab1.Tab = 0 Then
        Set myLV = LV
        iItemIDIndex = 5
    ElseIf SSTab1.Tab = 1 Then
        Set myLV = LVPList
        iItemIDIndex = 3
    End If
    
    Dim StartIdx As Integer

    If myLV.SelectedItem Is Nothing Then
       StartIdx = 1
    Else
       If myLV.SelectedItem.Index = myLV.ListItems.count Then
          StartIdx = 1
       Else
          StartIdx = myLV.SelectedItem.Index
       End If
    End If

    Dim fldValue As String

    For i = StartIdx To myLV.ListItems.count
       
        If cmbField.ListIndex = 0 Then
            fldValue = Trim(myLV.ListItems(i))
        Else
            If cmbField.ListIndex = 5 Then  'Item
                fldValue = Trim(myLV.ListItems(i).SubItems(iItemIDIndex))
            Else
                fldValue = Trim(myLV.ListItems(i).SubItems(cmbField.ListIndex))
            End If
        End If
    
       If LCase(fldValue) Like "*" & LCase(Trim(txtFind)) & "*" Then
          myLV.ListItems(i).Selected = True
          myLV.ListItems(i).EnsureVisible
          If i <> StartIdx Then Exit For
       End If
       
    Next
    
    If i = myLV.ListItems.count + 1 Then
        MsgBox "No Mathes Found.", vbInformation
        Set myLV.SelectedItem = Nothing
    End If
End Sub


Private Sub cmdPrint_Click()

    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\DispatchList.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdPrintIL_Click()

    Call StartTrans(con)
    
    con.Execute "DELETE FROM DispatchListInnerLabelPrinting WHERE DLDA_RefID=" & Val(txtBalanceIL.Tag)
    Dim i As Integer
    For i = 1 To LVIL.ListItems.count
        With LVIL.ListItems(i)
            con.Execute "INSERT INTO DispatchListInnerLabelPrinting(Labels,Qty,DLDA_RefID,ItemID) VALUES(" & _
             Val(.SubItems(2)) & "," & Val(.SubItems(1)) & "," & Val(txtBalanceIL.Tag) & ",'" & txtTotalQtyIL.Tag & "')"
        End With
    Next
    con.CommitTrans
    
    ''''Now Printing...
    Dim rpt As CRAXDDRT.Report
    Dim strReportName As String
    strReportName = GetSingleStringValue("InnerPackingLabel", "ForeignCustomers", " WHERE CustCode='" & cmbCustomer.ID & "' AND Country='" & cmbCountry.Text & "'")
    
    Set rpt = rptApp.OpenReport(RptPath & "\" & strReportName)
    rpt.ParameterFields.GetItemByName("@DLDA_RefID").SetCurrentValue Val(txtBalanceIL.Tag)
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, "", False
    
    FraPrintIL.Visible = False
    FAcc.Enabled = True
    
End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub



Private Sub cmdSelectAll_Click()

    Dim i As Integer
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = True
        LV.ListItems(i).ListSubItems(8).Tag = Val(LV.ListItems(i).SubItems(8))
        LV.ListItems(i).ListSubItems(7).Tag = Val(LV.ListItems(i).SubItems(8))
    Next
    GetCheckItemsTotal

End Sub

Private Sub cmdSelectAll1_Click()
    Dim i As Integer
    For i = 1 To LVPList.ListItems.count
        LVPList.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmdTransfer_Click()
    
    'If LVPList.ListItems.count = 0 Then Exit Sub
    
    If txtDispatchListNo = "" Then
        MsgBox "Please Enter Dispatch List No.", vbInformation
        Exit Sub
    End If
    '-----------------------------------------------
    Dim i As Integer, lEntryID As Long, lQty As Long, j As Integer, lEntryIDCartons As Long, k As Long, l As Long
    Call StartTrans(con)
    If lDPListEntryID = 0 Then
        con.Execute "INSERT INTO DispatchList(UserName,MachineName,DT,DispatchListNo,CustCode,Country,NewFormat) VALUES('" & CurrentUserName & _
         "','" & strComputerName & "','" & GetServerDate & "','" & txtDispatchListNo & "','" & cmbCustomer.Text & "','" & cmbCountry.Text & "',1)"
     
        lEntryID = GetSingleLongValue("MAX(EntryID)", "DispatchList", "WHERE MachineName='" & strComputerName & "'")
    Else
        con.Execute "UPDATE DispatchList SET DispatchListNo='" & txtDispatchListNo & "',CustCode='" & cmbCustomer.ID & "',Country='" & cmbCountry.Text & "' WHERE EntryID=" & lDPListEntryID
        lEntryID = lDPListEntryID
        If strDelList <> "" Then
            strDelList = Left(strDelList, Len(strDelList) - 1)
            Dim rs As New ADODB.Recordset
            Dim l_Temp_Qty As Long
            With rs
                .Open "SELECT VRD_RefID,InnerNo,DispatchListDetail_VRD.Qty FROM DispatchListDetail_VRD INNER JOIN DispatchListDetails_Adv ON DispatchListDetails_Adv.EntryID=DispatchListDetail_VRD.DLD_RefID INNER JOIN DispatchListDetail_Inners ON DLDC_RefID=DispatchListDetail_Inners.EntryID INNER JOIN DispatchList ON DispatchList.EntryID=DispatchListDetails_Adv.RefID WHERE DispatchListDetail_Inners.EntryID IN(" & strDelList & ") AND NewFormat=1", con, adOpenForwardOnly, adLockReadOnly
                Do Until .EOF
                    l_Temp_Qty = Val(!Qty & "") '(Val(!CartonTo & "") - Val(!CartonFrom & "") + 1) * Val(!Qty & "")
                    con.Execute "UPDATE VendRcvdDetail SET IssQty=IssQty-" & l_Temp_Qty & " WHERE EntryID=" & Val(!VRD_RefID & "")
                    .MoveNEXT
                Loop
                .Close
            End With
            Set rs = Nothing
            'con.Execute "DELETE FROM DispatchListDetail_Cartons WHERE EntryID IN(" & strDelList & ")", iRecAffect
            con.Execute "DELETE FROM DispatchListDetail_Inners WHERE EntryID IN(" & strDelList & ")", iRecAffect
            con.Execute "DELETE FROM DispatchListDetail_VRD WHERE DLDC_RefID IN(" & strDelList & ")", iRecAffect
            con.Execute "DELETE FROM DispatchListDetails_Adv WHERE EntryID NOT IN(SELECT RefID FROM DispatchListDetail_Cartons) AND EntryID NOT IN(SELECT RefID FROM DispatchListDetail_Inners)", iRecAffect
            
        End If
    End If
    
    Dim lEntryIDDetail As Long, bMasterPosted As Boolean
    
        For j = 1 To LVPList.ListItems.count
            With LVPList.ListItems(j)
                If Val(.ListSubItems(2).Tag) = 0 Then
                    For k = 1 To LVLots.ListItems.count
                        With LVLots.ListItems(k)
                            If Val(.SubItems(1)) = j Then   'Same Row...
                                For l = 1 To LVLotsMain.ListItems.count
                                    If Val(.SubItems(3)) = l Then
                                        con.Execute "INSERT INTO DispatchListDetails_Adv(RefID,ItemCode,Qty,ProcessID) VALUES(" & _
                                         lEntryID & ",'" & LVLotsMain.ListItems(l).Text & "'," & Val(LVLotsMain.ListItems(l).SubItems(1)) & "," & Val(LVLotsMain.ListItems(l).ListSubItems(2).Tag) & ")"
                                             
                                        lEntryIDDetail = GetSingleLongValue("MAX(EntryID)", "DispatchListDetails_Adv", " WHERE RefID=" & lEntryID)
                                        con.Execute "UPDATE Items SET InHand=ISNULL(InHand,0)+" & Val(LVLotsMain.ListItems(l).SubItems(1)) & " WHERE ItemID='" & Val(LVLotsMain.ListItems(l).Text) & "'"
                                        GoTo GotDLD_RefID
                                    End If
                                Next
                            End If
                        End With
                    Next
GotDLD_RefID:
                
                    'con.Execute "INSERT INTO DispatchListDetail_Cartons(RefID,CartonFrom,CartonTo,Qty) VALUES(" & _
                     lEntryIDDetail & "," & Val(.Text) & "," & Val(.SubItems(1)) & "," & Val(.SubItems(2)) & ")"
                    
                    con.Execute "INSERT INTO DispatchListDetail_Inners(RefID,InnerNo,Qty) VALUES(" & _
                     lEntryIDDetail & "," & Val(.Text) & "," & Val(.SubItems(2)) & ")"
                     
                    lEntryIDCartons = GetSingleLongValue("MAX(EntryID)", "DispatchListDetail_Inners")
                    
                    For k = 1 To LVLots.ListItems.count
                        With LVLots.ListItems(k)
                            If Val(.SubItems(1)) = j Then     'Same Row...
                                con.Execute "INSERT INTO DispatchListDetail_VRD(DLD_RefID,DLDC_RefID,VRD_RefID,Qty) VALUES(" & _
                                 lEntryIDDetail & "," & lEntryIDCartons & "," & Val(.Text) & "," & Val(.SubItems(2)) & ")"
                            
                                con.Execute "UPDATE VendRcvdDetail SET IssQty=IssQty+" & Val(.SubItems(2)) & " WHERE EntryID=" & Val(.Text)
                            End If
                        End With
                    Next
                    '-------------------------------------
                End If
            End With
        Next
    
        con.CommitTrans
        
        Call EditDispatchList(lEntryID)
        'Unload Me
        
    'Call RefreshLV
    
End Sub

Private Sub cmdUnSelectAll_Click()

    Dim i As Integer
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = False
        LV.ListItems(i).ListSubItems(8).Tag = "0"
    Next
    GetCheckItemsTotal

End Sub

Private Sub cmdUnSelectAll1_Click()
    Dim i As Integer
    For i = 1 To LVPList.ListItems.count
        LVPList.ListItems(i).Checked = False
    Next
End Sub

Private Sub DtFrom_Change()
    'Call setDateRange(DtFrom)
End Sub

Private Sub DtTo_Change()
    'Call setDateRange(DtTo)
End Sub


Private Sub Form_Load()
        
    lDPListEntryID = 0
    mnuPop.Visible = False
    
    If Not (strCompany = "QEL" Or (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Banzai" Or strCompany = "Towne") Then
        'mnuPrintPTC.Visible = False
    End If
    
    cmbCust.ListHeight = 3800
    cmbCust.AddItem "<All Makers>", "0"
    Call cmbCust.AddVals(con, "'{ ' + VendID1 + ' } ' + VenderName ", "Makers", "VendID")
    cmbCust.ID = "0"
    DtFrom = Date
    DtTo = Date
    cmbRange.ListIndex = 0

    cmbItems.DropDownOffLine = True
    'cmbItems.AddItem "<All Articles>", "0"
    Call cmbItems.AddVals(con, "ItemID + ' ' + ItemName + ' { ' + ISNULL(Attributes,'') + ' }'", "Items", "ItemID", " WHERE InActive=0 ORDER BY ItemName")
    'cmbItems.ID = "0"
    
    With cmbField
        .Clear
        For i = 1 To LV.ColumnHeaders.count
            cmbField.AddItem LV.ColumnHeaders(i).Text
        Next
    End With
    
    strDelList = ""
    
    
    cmbCustomer.ListHeight = 2500
    Call cmbCustomer.AddVals(con, "Distinct CustCode", "ForeignCustomers", "CustCode")
    
End Sub

Private Sub RefreshLV()
    
    On Error GoTo err
    cmbItems.Tag = cmbItems.ID
    If cmbCustomer.MatchFound = False Then
        MsgBox "Please Select Customer.", vbInformation
        Exit Sub
    End If
'    If cmbItems.MatchFound = False Then
'        MsgBox "Please Select Article.", vbInformation
'        Exit Sub
'    End If
    
    Dim rs As New ADODB.Recordset
    Dim Cond As String
    Dim DTRange As String
    Dim ITM As ListItem
        
    Screen.MousePointer = vbHourglass
    
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
        DTRange = " OnlyDT BETWEEN '" & DtFrom & "' AND '" & DtTo & "'"
    End If
       
    Cond = " WHERE (" & DTRange
    If cmbCust.ID <> "0" Then
        Cond = Cond & " AND VendID=" & cmbCust.ID & " "
    End If
    'If cmbItems.ID <> "0" Then
    Cond = Cond & " AND ItemCode='" & cmbItems.ID & "' "
    'End If
    
    If Trim(txtOrderNo) <> "" Then
        Cond = Cond & " AND OrderNo='" & Trim(txtOrderNo) & "' "
    End If
    
    If Trim(txtLotNo) <> "" Then
        'Cond = Cond & " AND LotNo='" & Trim(txtLotNo) & "' "
        Cond = " WHERE (LotNo='" & Trim(txtLotNo) & "'"
    End If
    
    'AND NextProcessID IS NULL
    Cond = Cond & " AND (RcvdQty-ISNULL(Wastage,0)-ISNULL(LostQty,0)-ISNULL(ReWorkQty,0))>IssQty AND ReqAuth=0"
    Cond = Cond & " AND CustCode='" & cmbCustomer.ID & "')"
    
    If cmbCustomer.ID = "137" Then
        Cond = Cond & " AND (NextProcessID IS NULL OR ProcessID IN(" & iTemper_Inspection_Process_ID & "," & iFirst_Inspection_Process_ID & "))"
    Else
        Cond = Cond & " AND (NextProcessID IS NULL)"
    End If
    
    'Cond = Cond & " AND Left(CustomInvoice,4)='VTL '"
    With rs
        
        .Open "SELECT * FROM VVendReceivingList " & Cond & " ORDER BY DT", con, adOpenForwardOnly, adLockReadOnly
'        LVPList.ListItems.Clear
        LV.ListItems.Clear
        'LV.Sorted = False
        Do Until .EOF
'            If strCompany = "Tecno" Or (strCompany = "Dr-Frgz" And strSubCompany = "Electro") Then  'Check If this is the Process Just Before Scanning If Yes, THen Show otherwise Skip.
'                If IsNextProcessScanning(Val(!ProcessID & ""), !ItemCode) = False Then
'                    GoTo SKIPEntry
'                End If
'            End If
            Set ITM = LV.ListItems.add(, , .AbsolutePosition)
            ITM.Tag = Val(![VRD_EntryID] & "")
            ITM.ListSubItems.add(, , "{ " & !VendID1 & " } " & !VenderName).Tag = "" '![ImpMaterial] & ""
            ITM.ListSubItems.add , , !InternalRefNo & "" '!OrderNo & ""
            ITM.ListSubItems.add(, , Format(![DT] & "", "dd-MMM-yyyy")).Tag = ""
            ITM.ListSubItems.add(, , !Description & "").Tag = Val(!ProcessID & "")
            ITM.ListSubItems.add(, , ![ItemCode] & "").Tag = Val(!Issuance_RefID & "")
            
            ITM.ListSubItems.add , , Val(!RcvdQty & "") - (Val(!Wastage & "") + Val(!LostQty & "")) - Val(!IssQty & "") - Val(!ReWorkQty & "")
            ITM.ListSubItems.add , , "0"
            ITM.ListSubItems.add , , Val(!RcvdQty & "") - (Val(!Wastage & "") + Val(!LostQty & "")) - Val(!IssQty & "") - Val(!ReWorkQty & "")
'            ITM.ListSubItems.add , , Val(!Wastage & "") + Val(!LostQty & "")
'            ITM.ListSubItems.add(, , IIf(!ReqAuth, "Under Auth.", "Authorized")).Tag = !LotNo & ""
            ITM.ListSubItems.add , , (!LotNo & "")
            ITM.ListSubItems(3).Tag = !OrderNo & ""
SKIPEntry:
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

Private Function IsNextProcessScanning(lProcessID As Long, strItemID As String) As Boolean
    Dim iSNo As Integer
    iSNo = GetSingleLongValue("SNo", "ItemProcesses", " WHERE ItemID='" & strItemID & "' AND ProcessID=" & lProcessID)
    'Now Get The Next Process
    Dim bScanning As Boolean
    bScanning = GetSingleBooleanValue("TOP 1 Scanning", "ItemProcesses", " WHERE ItemID='" & strItemID & "' AND SNo>" & iSNo & " ORDER BY SNo")
    IsNextProcessScanning = bScanning
End Function

Private Sub Form_Resize()
    On Error Resume Next
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    FAcc.Top = (Me.ScaleHeight - FAcc.Height) / 2
    
    
    FraPrintIL.Move (Me.ScaleWidth - FraPrintIL.Width) / 2, (Me.ScaleHeight - FraPrintIL.Height) / 2
End Sub




Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    
    If Item.Checked Then
             
        'Get First Item
        Dim i As Integer, iProcessID As Integer
        For i = 1 To LV.ListItems.count
            If LV.ListItems(i).Checked And i <> Item.Index Then
                iProcessID = Val(LV.ListItems(i).ListSubItems(4).Tag)
                Exit For
            End If
        Next
        If Val(Item.ListSubItems(4).Tag) <> iProcessID And iProcessID <> 0 Then
            Item.Checked = False
            Exit Sub
        End If
        With txtEdit
            Item.Selected = True
            .Move LV.Left + LV.ColumnHeaders(9).Left, LV.Top + Item.Top, LV.ColumnHeaders(9).Width
            .Text = Val(Item.SubItems(8))
            .SelStart = 0
            .SelLength = Len(.Text)
            .Visible = True
            .SetFocus
        End With
    Else
        Item.ListSubItems(8).Tag = "0"
        Item.ListSubItems(7).Tag = "0"
        GetCheckItemsTotal
    End If
    
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
    If LV.ListItems.count > 0 Then
    End If
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Exit Sub
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



Private Sub LVPList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Dim myLV As ListView
    Set myLV = LVPList
    
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

Private Sub LVPList_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Val(Item.ListSubItems(2).Tag) = 0 Then
        Item.Checked = False
    End If
    Call getCheckedTotals
End Sub

Private Sub getCheckedTotals()

    Dim i As Integer, lTotal As Long
    For i = 1 To LVPList.ListItems.count
        If LVPList.ListItems(i).Checked Then
            lTotal = lTotal + Val(LVPList.ListItems(i).SubItems(2))
        End If
    Next
    lblTotal = lTotal
    
End Sub

Private Sub LVPList_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If LVPList.ListItems.count = 0 Then Exit Sub
'    If KeyCode = 46 Then
''        Dim lSelectedCount As Long
''        lSelectedCount = 0
''        For i = 1 To LVPList.ListItems.count
''            If LVPList.ListItems(i).Checked Then
''                lSelectedCount = lSelectedCount + 1
''            End If
''        Next
''        If lSelectedCount > 1 Then
''        End If
'
'        If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo) = vbYes Then
'            If Val(LVPList.SelectedItem.ListSubItems(2).Tag) > 0 Then
'                strDelList = strDelList & Val(LVPList.SelectedItem.ListSubItems(2).Tag) & ","
'            End If
'            LVPList.ListItems.Remove LVPList.SelectedItem.Index
'            Call cmdTransfer_Click
'        End If
'    End If
    
End Sub

Private Sub LVPList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If LVPList.ListItems.count = 0 Then Exit Sub
    Me.PopUpMenu mnuPop
End Sub

Private Sub mnuAddToCarton_Click()

    Dim strCartonNo As String
    Dim f As New frmGetSingleTextValue
    Load f
    If f.getTextValue(strCartonNo, False, "Carton No.", "Carton No.") = False Then Exit Sub
    
    Dim i As Integer
    For i = 1 To LVPList.ListItems.count
        With LVPList.ListItems(i)
            If .Checked Then
                .SubItems(5) = Val(strCartonNo)
                con.Execute "UPDATE DispatchListDetail_Inners SET CartonNo=" & Val(strCartonNo) & " WHERE EntryID=" & .ListSubItems(2).Tag
            End If
        End With
    Next
    Call getCheckedTotals
    
End Sub

Private Sub mnuPrintInnerLabelsold_Click()
    
    LVIL.ListItems.Clear
    FraPrintIL.Visible = True
    Dim lTotalQty As Long
    txtTotalQtyIL.Tag = LVPList.SelectedItem.SubItems(3)
    lTotalQty = (Val(LVPList.SelectedItem.SubItems(1)) - Val(LVPList.SelectedItem.Text) + 1) * Val(LVPList.SelectedItem.SubItems(2))
    
    txtTotalQtyIL = lTotalQty
    txtBalanceIL = lTotalQty
    
    txtBalanceIL.Tag = Val(LVPList.SelectedItem.ListSubItems(2).Tag)
    
    txtQtyIL.SetFocus
    
    FAcc.Enabled = False
    
    
    '''''''''''
    Dim rs As New ADODB.Recordset, ITM As ListItem, lTotal As Long, lGTotal As Long
    Dim i As Integer
    With rs
        .Open "SELECT * FROM DispatchListInnerLabelPrinting WHERE DLDA_RefID=" & Val(txtBalanceIL.Tag) & " ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            
            Set ITM = LVIL.ListItems.add()
            ITM.ListSubItems.add , , Val(!Qty & "")
            ITM.ListSubItems.add , , Val(!Labels & "")
            
            lTotal = Val(!Qty & "") * Val(!Labels & "")
            lGTotal = lGTotal + lTotal
            ITM.ListSubItems.add , , lTotal
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    txtBalanceIL = txtBalanceIL - lGTotal
    
End Sub

Private Sub mnuDelete_Click()
    
    If LVPList.ListItems.count = 0 Then Exit Sub
    
'        Dim lSelectedCount As Long
'        lSelectedCount = 0
'        For i = 1 To LVPList.ListItems.count
'            If LVPList.ListItems(i).Checked Then
'                lSelectedCount = lSelectedCount + 1
'            End If
'        Next
'        If lSelectedCount > 1 Then
'        End If
        
        If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo) = vbYes Then
            For i = LVPList.ListItems.count To 1 Step -1
                With LVPList.ListItems(i)
                    If .Checked Then
                        If Val(.ListSubItems(2).Tag) > 0 Then
                            strDelList = strDelList & Val(.ListSubItems(2).Tag) & ","
                        End If
                        LVPList.ListItems.Remove .Index
                    End If
                End With
            Next
            'Call cmdTransfer_Click
        End If
        
    

End Sub

Private Sub mnuPrintInnerLabels_Click()
    Call PrintInnerLabels
End Sub

Private Sub PrintInnerLabels(Optional bOtherLabel As Boolean = False, Optional bWithManualQty As Boolean = False)
    
    On Error GoTo err
    Dim strEntries As String, l As Long
    
    For l = 1 To LVPList.ListItems.count
        With LVPList.ListItems(l)
            If .Checked Then
                strEntries = strEntries & Val(.ListSubItems(2).Tag) & ","
            End If
        End With
    Next
    
    If strEntries <> "" Then
        strEntries = Left(strEntries, Len(strEntries) - 1)
    Else
        Exit Sub
    End If
    
    If bWithManualQty Then
        Dim f As New frmGetSingleTextValue
        Load f
        Dim strQty As String, lQty As Long
        If f.getTextValue(strQty, False, "Qty", "Qty") = False Then Exit Sub
        lQty = strQty
    End If
    
    Dim rpt As CRAXDDRT.Report
    Dim strReportName As String
    If bWithManualQty Then
        'strReportName = "PackingLabelInner_MQ.rpt"
        If bOtherLabel Then
            strReportName = GetSingleStringValue("Inner_Label_Manual_II", "ForeignCustomers", " WHERE CustCode='" & cmbCustomer.ID & "' AND Country='" & cmbCountry.Text & "'")
        Else
            strReportName = GetSingleStringValue("Inner_Label_Manual_I", "ForeignCustomers", " WHERE CustCode='" & cmbCustomer.ID & "' AND Country='" & cmbCountry.Text & "'")
        End If
    Else
        strReportName = GetSingleStringValue("InnerPackingLabel", "ForeignCustomers", " WHERE CustCode='" & cmbCustomer.ID & "' AND Country='" & cmbCountry.Text & "'")
        If bOtherLabel Then
            strReportName = Left(strReportName, Len(strReportName) - 4) & "_II.rpt"
        End If
    End If
    
    Set rpt = rptApp.OpenReport(RptPath & "\" & strReportName)
    If bWithManualQty Then
        rpt.FormulaFields.GetItemByName("QtyToShow").Text = lQty
    End If
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "{VPrintInnerLabels.EntryID} IN[" & strEntries & "]", rpt
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub mnuPrintInnerLabelsII_Click()
    Call PrintInnerLabels(True)
End Sub

Private Sub mnuPrintInnerLabelsManualQty_Click()
    Call PrintInnerLabels(False, True)
End Sub

Private Sub mnuPrintInnerLabelsManualQtyII_Click()
    Call PrintInnerLabels(True, True)
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Val(txtEdit) < 0 Or Val(txtEdit) > Val(LV.SelectedItem.SubItems(8)) Then
            MsgBox "Invalid Qty.", vbInformation
            Exit Sub
        End If
        LV.SelectedItem.ListSubItems(8).Tag = Val(txtEdit)
        LV.SelectedItem.ListSubItems(7).Tag = Val(txtEdit)
        txtEdit.Visible = False
        GetCheckItemsTotal
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
    
End Sub

Private Sub GetCheckItemsTotal()

    Dim i As Integer, lTotal As Long
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                lTotal = lTotal + Val(.ListSubItems(8).Tag)
            End If
        End With
    Next
    txtTotalQty = lTotal
    
End Sub

Private Sub txtEdit_LostFocus()
'    txtEdit.Visible = False
End Sub

Private Sub txtFind_Change()
    cmdFind.Enabled = Len(Trim(txtFind)) > 0
End Sub


Private Sub txtLabelsIL_Change()
    txtTotalIL = Val(txtQtyIL) * Val(txtLabelsIL)
End Sub

Private Sub txtLabelsIL_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        If Val(txtTotalIL) > Val(txtBalanceIL) Then
            MsgBox "Invalid Qty.", vbInformation
            Exit Sub
        End If
        
        Dim ITM As ListItem
        Set ITM = LVIL.ListItems.add()
        ITM.ListSubItems.add , , Val(txtQtyIL)
        ITM.ListSubItems.add , , Val(txtLabelsIL)
        ITM.ListSubItems.add , , Val(txtTotalIL)
        
        txtBalanceIL = Val(txtBalanceIL) - Val(txtTotalIL)
                
        txtQtyIL = ""
        txtLabelsIL = ""
        txtQtyIL.SetFocus
        
    End If
        
End Sub

Private Sub txtLotNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call RefreshLV
    End If
End Sub

Private Sub txtQtyIL_Change()
    txtTotalIL = Val(txtQtyIL) * Val(txtLabelsIL)
End Sub

Private Sub txtTotalQty_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        If strCompany = "IAA" Or strCompany = "Kami" Then
            'Check the Qty/Order, It should not increase the actual Order Qty....
            If CheckExcessOrderQty = False Then
                'MsgBox "Invalid Qty.", vbInformation
                Exit Sub
            End If
        End If
        
        Dim dResult As Double, i As Integer, ITM As ListItem, iCartonTo As Long, iCartonFrom As Long, iQty As Long
        Dim strArticleNo As String, strArticleName As String
        Dim iProcessID As Integer, strProcessName As String
        
        strArticleNo = LV.SelectedItem.SubItems(5)
        strArticleName = GetSingleStringValue("ItemName", "Items", " WHERE ItemID='" & LV.SelectedItem.SubItems(5) & "'")
        
        iProcessID = Val(LV.SelectedItem.ListSubItems(4).Tag)
        strProcessName = LV.SelectedItem.SubItems(4)
        
        Dim ITM1 As ListItem, ITMMain As ListItem
        Dim lTotalQty As Long, lBalanceQty As Long, lQtyToUse As Long
        dResult = Val(txtTotalQty) / Val(txtCartonQty)
        
        Dim lCrtns As Long
        iQty = Val(txtCartonQty)
        iTotalCartons = Floor(Val(txtTotalQty) / Val(txtCartonQty))

        iCartonTo = Val(txtCartonFrom) + iTotalCartons - 1
        lCrtns = iTotalCartons 'Val(txtCrtnTo) - Val(txtCrtnFrom) + 1
        Dim lInnerNo As Long
        Dim iCounter As Integer
        
        Set ITMMain = LVLotsMain.ListItems.add(, , cmbItems.Tag)
        ITMMain.ListSubItems.add , , iQty
        ITMMain.ListSubItems.add(, , strProcessName).Tag = iProcessID
        'lInnerNo = LVPList.ListItems.count
        lInnerNo = 0
        For i = 1 To LVPList.ListItems.count
            If Val(LVPList.ListItems(i).Text) > lInnerNo Then
                lInnerNo = Val(LVPList.ListItems(i).Text)
            End If
        Next
        
        For iCounter = 0 To iTotalCartons - 1
            lInnerNo = lInnerNo + 1
            Set ITM = LVPList.ListItems.add(, , lInnerNo)
            ITM.Tag = LV.SelectedItem.Tag
            ITM.ListSubItems.add , , lInnerNo
            ITM.ListSubItems.add , , iQty
        
            ITM.ListSubItems.add , , LV.SelectedItem.SubItems(5) 'Article No.
            ITM.ListSubItems.add , , GetSingleStringValue("ItemName", "Items", " WHERE ItemID='" & LV.SelectedItem.SubItems(5) & "'") 'LV.SelectedItem.SubItems(3) 'Article Name.
            
            '-------------
            lBalanceQty = iQty
            For i = 1 To LV.ListItems.count
                With LV.ListItems(i)
                    If .Checked Then
                        Set ITM1 = LVLots.ListItems.add(, , Val(.Tag))
                        ITM1.ListSubItems.add , , ITM.Index
                        If Val(.ListSubItems(7).Tag) >= lBalanceQty Then
                            lQtyToUse = lBalanceQty
                        Else
                            lQtyToUse = Val(.ListSubItems(7).Tag)
                        End If
                        .ListSubItems(7).Tag = Val(.ListSubItems(7).Tag) - lQtyToUse
                        
                        ITM1.ListSubItems.add , , lQtyToUse 'Val(.ListSubItems(8).Tag)
                        ITM1.ListSubItems.add , , Val(ITMMain.Index)
                        
                        
                        .SubItems(7) = Val(.SubItems(7)) + lQtyToUse 'Val(.ListSubItems(8).Tag)
                        .SubItems(8) = Val(.SubItems(8)) - lQtyToUse '"0" 'LV.SelectedItem.SubItems(8) - (iQty * lCrtns)
                        lBalanceQty = lBalanceQty - lQtyToUse
                        
                    End If
                    If lBalanceQty <= 0 Then Exit For
                End With
            Next
        Next
        
        lTotalQty = lCrtns * iQty
        lBalanceQty = lTotalQty
        
'        For i = 1 To LV.ListItems.count
'            With LV.ListItems(i)
'                If .Checked Then
'                    Set ITM1 = LVLots.ListItems.add(, , Val(.Tag))
'                    ITM1.ListSubItems.add , , ITM.Index
'                    If Val(.ListSubItems(7).Tag) >= lBalanceQty Then
'                        lQtyToUse = lBalanceQty
'                    Else
'                        lQtyToUse = Val(.ListSubItems(7).Tag)
'                    End If
'                    .ListSubItems(7).Tag = Val(.ListSubItems(7).Tag) - lQtyToUse
'
'                    ITM1.ListSubItems.add , , lQtyToUse 'Val(.ListSubItems(8).Tag)
'                    ITM1.ListSubItems.add , , Val(ITMMain.Index)
'                    .SubItems(7) = Val(.SubItems(7)) + lQtyToUse 'Val(.ListSubItems(8).Tag)
'                    .SubItems(8) = Val(.SubItems(8)) - lQtyToUse '"0" 'LV.SelectedItem.SubItems(8) - (iQty * lCrtns)
'                    lBalanceQty = lBalanceQty - lQtyToUse
'
'                End If
'            End With
'        Next
        
        
        If Val(txtTotalQty) Mod Val(txtCartonQty) > 0 Then
            iCartonFrom = iCartonTo + 1
            iCartonTo = iCartonFrom
            lCrtns = 1
            iQty = Val(txtTotalQty) Mod Val(txtCartonQty)
            lInnerNo = lInnerNo + 1
            Set ITM = LVPList.ListItems.add(, , lInnerNo)
            ITM.Tag = LV.SelectedItem.Tag
            ITM.ListSubItems.add , , lInnerNo
            ITM.ListSubItems.add , , iQty
    
            
            ITM.ListSubItems.add , , strArticleNo
            ITM.ListSubItems.add , , strArticleName
            
            lTotalQty = iQty
            lBalanceQty = lTotalQty
            
            For i = 1 To LV.ListItems.count
                With LV.ListItems(i)
                    If .Checked And Val(.ListSubItems(7).Tag) > 0 Then
                    
                        Set ITM1 = LVLots.ListItems.add(, , Val(.Tag))
                        ITM1.ListSubItems.add , , ITM.Index
                        If Val(.ListSubItems(7).Tag) >= lBalanceQty Then
                            lQtyToUse = lBalanceQty
                        Else
                            lQtyToUse = Val(.ListSubItems(7).Tag)
                        End If
                        .ListSubItems(7).Tag = Val(.ListSubItems(7).Tag) - lQtyToUse
                        ITM1.Tag = "Sub"
                        ITM1.ListSubItems.add , , lQtyToUse 'Val(.ListSubItems(8).Tag)
                        ITM1.ListSubItems.add , , Val(ITMMain.Index)
                        .SubItems(7) = Val(.SubItems(7)) + lQtyToUse 'Val(.ListSubItems(8).Tag)
                        .SubItems(8) = Val(.SubItems(8)) - lQtyToUse '"0" 'LV.SelectedItem.SubItems(8) - (iQty * lCrtns)
                        lBalanceQty = lBalanceQty - lQtyToUse
                        
                    End If
                End With
            Next
        End If
        For i = 1 To LV.ListItems.count
            LV.ListItems(i).Checked = False
        Next
        Call CalculateTotals
         
        'Call cmdTransfer_Click
        txtCartonFrom = lInnerNo + 1
        txtTotalQty = "0"
        
    ElseIf KeyAscii = 27 Then
        'PicAddCartons.Visible = False
    End If

End Sub

Private Sub CalculateTotals()


    Dim dTotal As Double
    Dim DPcsTotal As Long
    Dim dblTotalCartons As Long
    Dim Cartons As Long
    Dim PrevCartonNo As Long
    Dim dCurrCartons As Long
    
    With LVPList
        For i = 1 To .ListItems.count
            'dQty = dQty + Val(.TextMatrix(i, 4))
            
            If PrevCartonNo < Val(.ListItems(i).SubItems(5)) Then
                PrevCartonNo = Val(.ListItems(i).SubItems(5))
            End If
            'dCurrCartons = Val(.ListItems(i).SubItems(1) & "") - Val(.ListItems(i)) + 1
            
            Pcs = Pcs + Val(.ListItems(i).SubItems(2) & "")
            
            dblTotalCartons = PrevCartonNo 'dblTotalCartons + Cartons
            DPcsTotal = Pcs 'DPcsTotal + (Pcs * dCurrCartons)
            
            dTotal = dTotal + (Val(Pcs * dCurrCartons)) '* Val(.ListItems(i).ListSubItems(4).Tag))
            'Pcs = 0
            'Cartons = 0
            'PrevCartonNo = Val(.ListItems(i).SubItems(1) & "")
        Next i
        
    End With
    
    'txttQty = dQty
    'lblTotalAmount = DTotal
    'lblTotWords = ConvertInWords(CCur(DTotal), txtCurrency.Text)
    
    lblTotalPcs = DPcsTotal
    lblTotalCartons = PrevCartonNo 'dblTotalCartons
    
End Sub

Public Sub EditDispatchList(p_lEntryID As Long)

    strDelList = ""
    lDPListEntryID = p_lEntryID
    txtDispatchListNo = GetSingleStringValue("DispatchListNo", "DispatchList", " WHERE EntryID=" & lDPListEntryID)
    LVPList.Checkboxes = True
    
    LVLots.ListItems.Clear
    LVLotsMain.ListItems.Clear
    LVIL.ListItems.Clear
    
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        '.Open "SELECT CustCode,Country,CartonFrom,CartonTo,DispatchListDetail_Cartons.Qty,DispatchListDetail_Cartons.EntryID,DispatchListDetail_Cartons.RefID," & _
         "ItemID,ItemName FROM DispatchList INNER JOIN DispatchListDetails_Adv ON DispatchList.EntryID=DispatchListDetails_Adv.RefID " & _
         "INNER JOIN DispatchListDetail_Cartons ON DispatchListDetails_Adv.EntryID=DispatchListDetail_Cartons.RefID INNER JOIN Items " & _
         "ON DispatchListDetails_Adv.ItemCode=Items.ItemID WHERE DispatchList.EntryID=" & lDPListEntryID & _
         " ORDER BY DispatchListDetails_Adv.EntryID,DispatchListDetail_Cartons.EntryID", con, adOpenForwardOnly, adLockReadOnly
        
        .Open "SELECT CustCode,Country,InnerNo,DispatchListDetail_Inners.Qty,DispatchListDetail_Inners.EntryID,DispatchListDetail_Inners.RefID," & _
         "ItemID,ItemName,DispatchListDetail_Inners.CartonNo,DispatchListDetail_Inners.InternalNos,Processes.Description,DispatchListDetail_Inners.Lots FROM DispatchList INNER JOIN DispatchListDetails_Adv ON DispatchList.EntryID=DispatchListDetails_Adv.RefID " & _
         "INNER JOIN VDispatchListDetail_Inners DispatchListDetail_Inners ON DispatchListDetails_Adv.EntryID=DispatchListDetail_Inners.RefID INNER JOIN Items " & _
         "ON DispatchListDetails_Adv.ItemCode=Items.ItemID " & _
         "LEFT JOIN Processes ON DispatchListDetails_Adv.ProcessID=Processes.ProcessID " & _
         "WHERE DispatchList.EntryID=" & lDPListEntryID & _
         " ORDER BY DispatchListDetails_Adv.EntryID,DispatchListDetail_Inners.EntryID", con, adOpenForwardOnly, adLockReadOnly
         
        If .EOF = False Then
            cmbCustomer.ID = !CustCode & ""
            cmbCountry.Text = !Country & ""
        End If
        
        LVPList.ListItems.Clear
        Do Until .EOF
            Set ITM = LVPList.ListItems.add(, , Val(!InnerNo & ""))
            ITM.Tag = "0"
            ITM.ListSubItems.add(, , Val(!InnerNo & "")).Tag = Val(!RefID & "")
            ITM.ListSubItems.add(, , Val(!Qty & "")).Tag = Val(!EntryID & "")
            ITM.ListSubItems.add , , !ITemID & ""
            ITM.ListSubItems.add , , !ItemName & ""
            ITM.ListSubItems.add , , Val(!CartonNo & "")
            ITM.ListSubItems.add , , !Lots & ""
            ITM.ListSubItems.add , , !InternalNos & ""
            ITM.ListSubItems.add , , !Description & ""
            .MoveNEXT
        Loop
         .Close
    End With
    Set rs = Nothing
        
    Call CalculateTotals
    Me.Show
    
End Sub


Private Function CheckExcessOrderQty() As Boolean

    Dim i As Integer, iSelected As Integer
    Dim strOrders() As String
    iSelected = 0
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                iSelected = iSelected + 1
            End If
        End With
    Next
    
    ReDim strOrders(iSelected, 5)   '0=Order;1=Item;2=Qty;3=Already in Dispatch Qty;4=Order Balance Qty
    Dim iIndex As Integer, j As Integer
    Dim strItemID As String, strOrderNo As String, bFound As Boolean, lQty As Long
    iIndex = 0
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                bFound = False
                strOrderNo = .ListSubItems(3).Tag '.SubItems(2)
                strItemID = .SubItems(5)
                lQty = Val(.ListSubItems(8).Tag)
                
                For j = 0 To iIndex - 1
                    If strOrders(j, 0) = strOrderNo And strOrders(j, 1) = strItemID Then
                        strOrders(j, 2) = Val(strOrders(j, 2)) + lQty
                        bFound = True
                    End If
                Next
                If bFound = False Then
                    strOrders(iIndex, 0) = strOrderNo
                    strOrders(iIndex, 1) = strItemID
                    strOrders(iIndex, 2) = lQty
                    iIndex = iIndex + 1
                End If
            End If
        End With
    Next
    
    'Now we have an array which holds Order+Item+Qty...Now...
    'Calculate already Dispatched Qty & Balance Order Qty for Comparison.
    For i = 0 To iIndex - 1
        strOrders(i, 3) = GetSingleLongValue("SUM(Qty)", "VDispatchListDetail_Ex_Orderwise", " WHERE OrderNo='" & strOrders(i, 0) & "' AND ItemCode='" & strOrders(i, 1) & "'")
        strOrders(i, 4) = GetSingleLongValue("SUM(Qty)", "FOrderItems", " WHERE OrderNo='" & strOrders(i, 0) & "' AND CompItemCode='" & strOrders(i, 1) & "'")
         
        If (Val(strOrders(i, 2)) + Val(strOrders(i, 3))) > (Val(strOrders(i, 4))) Then
'            MsgBox "Following will exceed order qty." & vbNewLine & "Order No.:" & strOrders(i, 0) & ", Item Code:" & strOrders(i, 1) & ", Order Qty=" & Val(strOrders(i, 4)) & ", Already Dispatch Qty=" & Val(strOrders(i, 3)) & ", Current Qty=" & Val(strOrders(i, 2))
'            CheckExcessOrderQty = False
'            Exit Function
        End If
    Next
     
    CheckExcessOrderQty = True
End Function
