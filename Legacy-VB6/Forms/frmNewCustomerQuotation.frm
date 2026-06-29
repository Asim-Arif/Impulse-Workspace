VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmNewCustomerQuotation 
   ClientHeight    =   9735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11925
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9735
   ScaleWidth      =   11925
   WindowState     =   2  'Maximized
   Begin VB.Frame FAll 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9675
      Left            =   45
      TabIndex        =   23
      Top             =   30
      Width           =   11880
      Begin VB.TextBox txtDestination 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   8070
         MaxLength       =   150
         TabIndex        =   69
         Top             =   8670
         Width           =   3705
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   75
         TabIndex        =   15
         Top             =   3015
         Width           =   2250
         _ExtentX        =   3969
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
         Height          =   4770
         Left            =   90
         TabIndex        =   40
         Top             =   3345
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   8414
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Items"
         TabPicture(0)   =   "frmNewCustomerQuotation.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblGoTo(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "cmbItemSearch"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "FNew"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "LV"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "DTEdit"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "txtEditQuality"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "txtEditPrice"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txtEditQty"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "cmdDelete"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).ControlCount=   9
         TabCaption(1)   =   "Other Charges"
         TabPicture(1)   =   "frmNewCustomerQuotation.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtOtherCharges"
         Tab(1).Control(1)=   "cmdAddOtherCharges"
         Tab(1).Control(2)=   "txtOtherChargeAmt"
         Tab(1).Control(3)=   "txtOtherChargesSNo"
         Tab(1).Control(4)=   "LVOtherCharges"
         Tab(1).Control(5)=   "lblCap(2)"
         Tab(1).Control(6)=   "lblCap(1)"
         Tab(1).Control(7)=   "lblCap(0)"
         Tab(1).ControlCount=   8
         Begin VB.TextBox txtOtherCharges 
            Height          =   315
            Left            =   -73710
            TabIndex        =   62
            Top             =   630
            Width           =   7005
         End
         Begin VB.CommandButton cmdAddOtherCharges 
            Caption         =   "Add"
            Height          =   315
            Left            =   -64635
            TabIndex        =   61
            Top             =   630
            Width           =   1245
         End
         Begin VB.TextBox txtOtherChargeAmt 
            Height          =   315
            Left            =   -66675
            TabIndex        =   60
            Top             =   630
            Width           =   2000
         End
         Begin VB.TextBox txtOtherChargesSNo 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   -74940
            TabIndex        =   59
            Top             =   630
            Width           =   1200
         End
         Begin VB.CommandButton cmdDelete 
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   11100
            Picture         =   "frmNewCustomerQuotation.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Remove This Entry From Voucher"
            Top             =   1215
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox txtEditQty 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5190
            TabIndex        =   44
            Top             =   1275
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtEditPrice 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3945
            TabIndex        =   43
            Top             =   1290
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtEditQuality 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6345
            TabIndex        =   41
            Top             =   1260
            Visible         =   0   'False
            Width           =   1260
         End
         Begin MSComCtl2.DTPicker DTEdit 
            Height          =   300
            Left            =   5505
            TabIndex        =   42
            Top             =   2505
            Visible         =   0   'False
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   529
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
            CustomFormat    =   "dd-MMM-yy"
            Format          =   162922499
            CurrentDate     =   37250
         End
         Begin MSComctlLib.ListView LV 
            Height          =   3465
            Left            =   75
            TabIndex        =   57
            Top             =   900
            Width           =   10980
            _ExtentX        =   19368
            _ExtentY        =   6112
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
            NumItems        =   10
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S No."
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Cust ID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Item ID"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Description"
               Object.Width           =   6985
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "P. Unit"
               Object.Width           =   1147
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Delivery"
               Object.Width           =   1693
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Remarks"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "Price"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   8
               Text            =   "Qty"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   9
               Text            =   "Amount"
               Object.Width           =   1764
            EndProperty
         End
         Begin VB.Frame FNew 
            BorderStyle     =   0  'None
            Height          =   570
            Left            =   75
            TabIndex        =   45
            Top             =   330
            Width           =   11535
            Begin VB.TextBox txtQty 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               DataField       =   "RptTime"
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   9300
               MaxLength       =   150
               TabIndex        =   50
               Top             =   255
               Width           =   660
            End
            Begin VB.TextBox txtPrice 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               DataField       =   "RptTime"
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   8670
               MaxLength       =   150
               TabIndex        =   49
               Top             =   255
               Width           =   645
            End
            Begin VB.TextBox txtUnit 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               DataField       =   "RptTime"
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   5220
               Locked          =   -1  'True
               MaxLength       =   150
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   255
               Width           =   720
            End
            Begin VB.CommandButton cmdAddItem 
               Caption         =   "&Add"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   10980
               Picture         =   "frmNewCustomerQuotation.frx":056A
               Style           =   1  'Graphical
               TabIndex        =   47
               ToolTipText     =   "Add This Entry To Voucher"
               Top             =   15
               UseMaskColor    =   -1  'True
               Width           =   540
            End
            Begin VB.TextBox txtQuality 
               Appearance      =   0  'Flat
               DataField       =   "RptTime"
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   7260
               MaxLength       =   150
               TabIndex        =   46
               Top             =   255
               Width           =   1425
            End
            Begin MSComCtl2.DTPicker DTDelivery 
               Height          =   300
               Left            =   5955
               TabIndex        =   58
               Top             =   255
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   529
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
               CustomFormat    =   "dd-MMM-yy"
               Format          =   162922499
               CurrentDate     =   37250
            End
            Begin MSForms.ComboBox cmbDesc 
               Height          =   285
               Left            =   1785
               TabIndex        =   55
               Top             =   255
               Width           =   3435
               VariousPropertyBits=   746604575
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "6059;503"
               ListWidth       =   10583
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
               Left            =   0
               TabIndex        =   54
               Top             =   255
               Width           =   1785
               VariousPropertyBits=   746604575
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "3149;503"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbCatID 
               Height          =   285
               Left            =   0
               TabIndex        =   53
               Top             =   600
               Visible         =   0   'False
               Width           =   1545
               VariousPropertyBits=   746604575
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "2725;503"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label lblAmt 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   9945
               TabIndex        =   52
               Top             =   255
               Width           =   1020
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BorderStyle     =   1  'Fixed Single
               Caption         =   $"frmNewCustomerQuotation.frx":0A9C
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
               Left            =   0
               TabIndex        =   51
               Top             =   15
               Width           =   10965
            End
         End
         Begin MSComctlLib.ListView LVOtherCharges 
            Height          =   3135
            Left            =   -74940
            TabIndex        =   63
            Top             =   990
            Width           =   11550
            _ExtentX        =   20373
            _ExtentY        =   5530
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
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
               Text            =   "S No."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description"
               Object.Width           =   10583
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Amount"
               Object.Width           =   3528
            EndProperty
         End
         Begin MSForms.ComboBox cmbItemSearch 
            Height          =   285
            Left            =   1335
            TabIndex        =   68
            Top             =   4395
            Width           =   9720
            VariousPropertyBits=   746604571
            BackColor       =   16777215
            ForeColor       =   6244673
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "17145;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label lblGoTo 
            AutoSize        =   -1  'True
            Caption         =   "Search Item : "
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
            Index           =   0
            Left            =   135
            TabIndex        =   67
            Top             =   4425
            Width           =   1185
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            Caption         =   "Amount"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   -66660
            TabIndex        =   66
            Top             =   390
            Width           =   555
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            Caption         =   "Charges Description :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   -73710
            TabIndex        =   65
            Top             =   405
            Width           =   1545
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            Caption         =   "S/No. :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   -74895
            TabIndex        =   64
            Top             =   405
            Width           =   510
         End
      End
      Begin VB.TextBox txtNoOfPackages 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   9675
         MaxLength       =   150
         TabIndex        =   14
         Top             =   2505
         Width           =   2115
      End
      Begin VB.TextBox txtNetWeight 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   7770
         MaxLength       =   150
         TabIndex        =   13
         Top             =   2505
         Width           =   1920
      End
      Begin VB.TextBox txtGrossWeight 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5865
         MaxLength       =   150
         TabIndex        =   12
         Top             =   2505
         Width           =   1935
      End
      Begin ComboList.Usercontrol1 cmbCompany 
         Height          =   285
         Left            =   75
         TabIndex        =   0
         Top             =   975
         Width           =   5265
         _ExtentX        =   9287
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
         TextForeColor   =   16711680
         ListIndex       =   -1
         Appearance      =   0
      End
      Begin VB.TextBox txtSwiftCode 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   9675
         MaxLength       =   150
         TabIndex        =   10
         Top             =   1995
         Width           =   2115
      End
      Begin VB.TextBox txtDepartment 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   9915
         MaxLength       =   150
         TabIndex        =   3
         Top             =   975
         Width           =   1875
      End
      Begin VB.TextBox txtAttention 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   9480
         MaxLength       =   150
         TabIndex        =   7
         Top             =   1485
         Width           =   2310
      End
      Begin VB.TextBox txtCustRef 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4395
         MaxLength       =   150
         TabIndex        =   38
         Top             =   8670
         Width           =   3690
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   915
         Left            =   75
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   8670
         Width           =   4335
      End
      Begin VB.TextBox txtDeliveryTime 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   75
         MaxLength       =   150
         TabIndex        =   4
         Top             =   1485
         Width           =   5265
      End
      Begin VB.TextBox txtPacking 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   75
         MaxLength       =   150
         TabIndex        =   11
         Top             =   2505
         Width           =   5805
      End
      Begin VB.TextBox txtValidity 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5340
         MaxLength       =   150
         TabIndex        =   5
         Top             =   1485
         Width           =   2370
      End
      Begin VB.TextBox txtCurrency 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5475
         MaxLength       =   150
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3030
         Width           =   2085
      End
      Begin VB.TextBox txtOrderNo 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   7545
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   18
         Top             =   3030
         Width           =   2280
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
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   11880
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "New Quotation"
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
            TabIndex        =   28
            Top             =   135
            Width           =   11700
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "New Quotation"
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
            TabIndex        =   29
            Top             =   150
            Width           =   11685
         End
      End
      Begin MSComCtl2.DTPicker OrderDT 
         Height          =   300
         Left            =   9825
         TabIndex        =   19
         Top             =   3015
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
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
         Format          =   162070531
         CurrentDate     =   37250
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
         Height          =   660
         Left            =   5190
         TabIndex        =   24
         Top             =   8940
         Width           =   6600
         Begin MSForms.CommandButton cmdSavenStay 
            Height          =   360
            Left            =   3300
            TabIndex        =   70
            Top             =   195
            Width           =   1590
            Caption         =   "Save & Stay"
            PicturePosition =   327683
            Size            =   "2805;635"
            Accelerator     =   116
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdCancel 
            Height          =   360
            Left            =   4920
            TabIndex        =   22
            Top             =   195
            Width           =   1590
            Caption         =   "Cancel            "
            PicturePosition =   327683
            Size            =   "2805;635"
            Accelerator     =   67
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdSClear 
            Height          =   360
            Left            =   60
            TabIndex        =   20
            Top             =   195
            Width           =   1590
            Caption         =   "Save & New   "
            PicturePosition =   327683
            Size            =   "2805;635"
            Accelerator     =   78
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdSClose 
            Height          =   360
            Left            =   1680
            TabIndex        =   21
            Top             =   195
            Width           =   1590
            Caption         =   "Save & Close  "
            PicturePosition =   327683
            Size            =   "2805;635"
            Accelerator     =   83
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmNewCustomerQuotation.frx":0B53
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
         Index           =   10
         Left            =   75
         TabIndex        =   39
         Top             =   2265
         Width           =   11715
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmNewCustomerQuotation.frx":0C20
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
         Left            =   75
         TabIndex        =   36
         Top             =   8430
         Width           =   11700
      End
      Begin MSForms.ComboBox cmbBankAccNo 
         Height          =   285
         Left            =   4755
         TabIndex        =   9
         Top             =   1995
         Width           =   4920
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "8678;503"
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
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmNewCustomerQuotation.frx":0CC2
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
         Left            =   75
         TabIndex        =   35
         Top             =   1755
         Width           =   11715
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmNewCustomerQuotation.frx":0D90
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
         Left            =   75
         TabIndex        =   34
         Top             =   1245
         Width           =   11715
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmNewCustomerQuotation.frx":0E55
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
         Index           =   3
         Left            =   75
         TabIndex        =   33
         Top             =   735
         Width           =   11715
      End
      Begin MSForms.ComboBox cmbBank 
         Height          =   285
         Left            =   75
         TabIndex        =   8
         Top             =   1995
         Width           =   4680
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "8255;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbShipVia 
         Height          =   285
         Left            =   7695
         TabIndex        =   6
         Top             =   1485
         Width           =   1800
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3175;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbPaymentTerms 
         Height          =   285
         Left            =   5340
         TabIndex        =   1
         Top             =   975
         Width           =   2370
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4180;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbTradeTerms 
         Height          =   285
         Left            =   7695
         TabIndex        =   2
         Top             =   975
         Width           =   2220
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3916;503"
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
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Amount :"
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
         Index           =   1
         Left            =   8760
         TabIndex        =   32
         Top             =   8175
         Width           =   885
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Amount (In Words) :"
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
         Left            =   75
         TabIndex        =   31
         Top             =   8175
         Width           =   1830
      End
      Begin VB.Label lblTotWords 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1890
         TabIndex        =   30
         Top             =   8175
         Width           =   6885
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmNewCustomerQuotation.frx":0F0E
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
         Index           =   7
         Left            =   75
         TabIndex        =   26
         Top             =   2775
         Width           =   11715
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   9630
         TabIndex        =   25
         Top             =   8175
         Width           =   2160
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   2355
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3015
         Width           =   3120
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5503;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frmNewCustomerQuotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lQuotNo As Long
Dim strDelList As String
Dim iColNo As Integer

Public Sub EditQuotation(p_lQuotNo As Long)
    
    On Error GoTo err
    lQuotNo = p_lQuotNo
    strDelList = ""
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "SELECT * FROM VCustomerQuotations WHERE QuotationNo=" & lQuotNo, con, adOpenForwardOnly, adLockReadOnly
        
        cmbCompany.ID = !CompanyRefID
        cmbCust.Text = ![CustCode] & ""
        cmbCountry = ![Country] & ""
        
        txtOrderNo = Val(![QuotationNo] & "")
        OrderDT = ![DT] & ""
        cmbTradeTerms = ![TradeTerms] & ""
        cmbShipVia = !ShipVia & ""
        cmbPaymentTerms = ![PaymentTerms] & ""

        txtPacking = ![Packing] & ""
        txtDeliveryTime = !DeliveryTime & ""
        txtValidity = !Validity & ""
        cmbBank = !Bank & ""
        txtRemarks = !Remarks & ""
        txtCustRef = !CustomerRef & ""
        
        txtDepartment = !Department & ""
        txtAttention = !AttentionPerson & ""
        txtSwiftCode = !SwiftCode & ""
        txtGrossWeight = !GrossWeight & ""
        txtNetWeight = !NetWeight & ""
        txtNoOfPackages = Val(!NoOfPackages & "")
        txtDestination = !Destination & ""
        cmbItemSearch.Clear
        LV.ListItems.Clear
        Do Until .EOF
        
            Set ITM = LV.ListItems.add(, !ItemCode & "'" & !QualityType & "", .AbsolutePosition)
            
            ITM.Tag = !ItemCode & "" '![CompITemID] & "" 'cmbItems.List(cmbItems.ListIndex, 2)
            
            ITM.ListSubItems.add , , !ItemCode & "" '![CatID] & ""
            ITM.ListSubItems.add , , ![ItemCode] & ""
            ITM.ListSubItems.add(, , ![ItemName] & "").Tag = "" '![Description] & ""
            
            ITM.ListSubItems.add , , !Unit & ""
            ITM.ListSubItems.add , , Format(!DeliveryDT, "dd-MMM-yy")
            ITM.ListSubItems.add , , !QualityType & ""
            
            ITM.ListSubItems.add(, , Format(![Price] & "", "#0.00")).Tag = 0 '![CustomPrice] & ""
            ITM.ListSubItems.add(, , Val(![Qty] & "")).Tag = 0
            
            ITM.ListSubItems.add(, , Val(![Qty] & "") * Val(![Price])).Tag = ![EntryID] & ""
            
            cmbItemSearch.AddItem !ItemCode
            cmbItemSearch.List(.AbsolutePosition - 1, 1) = !ItemCode & "'" & !QualityType & ""
            
            .MoveNEXT
        Loop
        .Close
        
        .Open "SELECT * FROM CustomerQuotations_OtherCharges WHERE QuotationNo=" & lQuotNo & " ORDER BY SNo", con_ServerSide, adOpenForwardOnly, adLockReadOnly
        LVOtherCharges.ListItems.Clear
        Do Until .EOF
            Set ITM = LVOtherCharges.ListItems.add(, , Val(!SNo & ""))
            ITM.SubItems(1) = !ChargesDetail & ""
            ITM.SubItems(2) = Val(!ChargesAmt & "")
            .MoveNEXT
        Loop
        .Close
        txtOtherChargesSNo.Text = LVOtherCharges.ListItems.count + 1
        
    End With
    
    Call GetTotal

    txtQty = 0
'    cmbItems.SetFocus
    
    Set rs = Nothing
    Me.Show
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub chkCust_Click()

    If cmbCatID.MatchFound Then
        Call FillItems
        For i = 1 To LV.ListItems.count
            If chkCust Then
                LV.ListItems(i).SubItems(2) = LV.ListItems(i).key
                LV.ListItems(i).ListSubItems(3) = LV.ListItems(i).ListSubItems(3).key
            Else
                LV.ListItems(i).SubItems(2) = LV.ListItems(i).Tag
                LV.ListItems(i).ListSubItems(3) = LV.ListItems(i).ListSubItems(3).Tag
            End If
        Next i
    End If
    
End Sub

Private Sub cmbBank_Change()
    Call cmbBank_Click
End Sub

Private Sub cmbBank_Click()
    cmbBankAccNo.Clear
    Call AddToCombo(cmbBankAccNo, "AccNo + ' ( ' + Branch + ' )'", "BankList", " WHERE CompanyRefID=" & cmbCompany.ID & " AND Bank='" & cmbBank.Text & "'", , "BankID")
    If cmbBankAccNo.ListCount > 0 Then cmbBankAccNo.ListIndex = 0
End Sub

Private Sub cmbCatID_Change()
    Call cmbCatID_Click
End Sub

Private Sub cmbCatID_Click()
    If cmbCatID.MatchFound = False Then Exit Sub
    'Call FillItems
End Sub

Private Sub cmbCompany_matched()
    
    On Error GoTo err
    
    If lQuotNo <> 0 Then Exit Sub
    Dim InvoiecNo As String
    InvoiceNo = GetSingleLongValue("MAX(QuotationNo)", "CustomerQuotations") + 1
    
    
    txtOrderNo = InvoiceNo
    
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select DISTINCT Bank From BankList WHERE CompanyRefID=" & cmbCompany.ID, con, adOpenForwardOnly, adLockReadOnly
        cmbBank.Clear
        Do Until .EOF
            cmbBank.AddItem .Fields(0) & ""
            'cmbBank.List(.AbsolutePosition - 1, 1) = .Fields(1) & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()
    
    On Error GoTo err
    If cmbCountry.MatchFound = False Then Exit Sub
    Dim rs As New ADODB.Recordset
    txtCurrency = cmbCountry.List(cmbCountry.ListIndex, 1)
    With rs
        '.Open "Select Distinct CatID,CatDesc From VFCustomerCatalog Where CustCode+Country='" & cmbCust.Text & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCatID.Clear
        cmbCatID.AddItem "<All Catagories>"
        If cmbCountry.Tag <> cmbCountry.Text Then
            cmbCountry.Tag = cmbCountry.Text
            Call FillItems
        End If
        'Do Until .EOF
       '     cmbCatID.AddItem .Fields(1) & ""
      '      cmbCatID.List(cmbCatID.ListCount - 1, 1) = .Fields(0) & ""
     '       .MoveNext
    '    Loop
    '    .Close
        .Open "SELECT TradeTerms,PaymentTerms,Cont1Name FROM ForeignCustomers WHERE CustCode+Country='" & cmbCust.Text & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbTradeTerms.Text = !TradeTerms & ""
        txtAttention = !Cont1Name & ""
        'cmbPayment.Text = !PaymentTerms & ""
        .Close
    End With
    
    Set rs = Nothing
    If cmbCatID.ListCount > 0 Then cmbCatID.ListIndex = 0
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbCust_matched()
    
    On Error GoTo err
    
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Country,Curr From ForeignCustomers Where CustCode='" & cmbCust.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            cmbCountry.List(cmbCountry.ListCount - 1, 1) = ![Curr] & ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    If cmbCountry.ListCount > 0 Then cmbCountry.ListIndex = 0
    
    Set rs = Nothing
'    If lQuotNo = "" Then
'        If Len(txtOrderNo) > 8 Then
'            txtOrderNo = Left(txtOrderNo, 8) & Right(cmbCust.ID, 4)
'        Else
'            txtOrderNo = Left(txtOrderNo, 8) & Right(cmbCust.ID, 4)
'        End If
'    End If
     
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbCust_UnMatched()
    'cmbCountry.ClearVals
End Sub

Private Sub cmbCustItemID_Change()
    Call cmbCustItemID_Click
End Sub

Private Sub cmbCustItemID_Click()
    cmbItems.ListIndex = cmbCustItemID.ListIndex
End Sub

Private Sub cmbCustItemID_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode.Value = 13 Then
        txtQty.SetFocus
    End If
End Sub

Private Sub cmbDesc_Change()
    Call cmbDesc_Click
End Sub

Private Sub cmbDesc_Click()
    If cmbDesc.MatchFound Then
        cmbItems.ListIndex = cmbDesc.ListIndex
    End If
    
End Sub

Private Sub cmbDesc_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode.Value = 13 Then
        txtQty.SetFocus
    End If
End Sub

Private Sub cmbItems_Change()
    Call cmbItems_Click
End Sub

Private Sub cmbItems_Click()

    If cmbItems.MatchFound Then
        'cmbCustItemID.ListIndex = cmbItems.ListIndex
        cmbDesc.ListIndex = cmbItems.ListIndex
        txtUnit = cmbItems.List(cmbItems.ListIndex, 1)
        txtPrice = Val(cmbItems.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 2) & "")
        txtPrice.Tag = Val(cmbDesc.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 2) & "")
    Else
        txtUnit = ""
        txtPrice = ""
        txtPrice.Tag = ""
    End If
    
End Sub

Private Sub cmbItems_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode.Value = 13 Then
        txtQty.SetFocus
    End If
End Sub

Private Sub cmbItemSearch_Change()
    Call cmbItemSearch_Click
End Sub

Private Sub cmbItemSearch_Click()

    If cmbItemSearch.MatchFound = False Then Exit Sub
    
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).key = cmbItemSearch.List(cmbItemSearch.ListIndex, 1) Then
            LV.ListItems(i).Selected = True
            LV.ListItems(i).EnsureVisible
            LV.SetFocus
            Exit For
        End If
    Next
    
End Sub

Private Sub cmbTradeTerms_Change()
    Call cmbTradeTerms_Click
End Sub

Private Sub cmbTradeTerms_Click()

    If cmbTradeTerms.MatchFound Then
        cmbCatID.Locked = False
        cmbItems.Locked = False
        'cmbCustItemID.Locked = False
        cmbDesc.Locked = False
        If cmbItems.MatchFound Then
            txtPrice = cmbItems.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 2)
        End If
    Else
        cmbCatID.Locked = True
        cmbItems.Locked = True
        'cmbCustItemID.Locked = True
        cmbDesc.Locked = True
    End If
    
End Sub


Private Sub cmdSavenStay_Click()
    If Saved Then
        Dim p_lQuotNo As Long
        p_lQuotNo = txtOrderNo
        Call Me.EditQuotation(p_lQuotNo)
    End If
End Sub

Private Sub DTEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        LV.SelectedItem.SubItems(iColNo - 1) = Format(DTEdit, "dd-MMM-yy")
        DTEdit.Visible = False
    ElseIf KeyCode = 27 Then
        DTEdit.Visible = False
    End If
End Sub

Private Sub DTEdit_LostFocus()
    DTEdit.Visible = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub cmdAddItem_Click()
    
    On Error GoTo err
    
    If Not cmbItems.MatchFound Or cmbItems = "" Then
        MsgBox "Invalid Item Selected", vbInformation
        Exit Sub
    ElseIf Val(txtQty) = 0 Then
        MsgBox "Quantity To Order Cannot Be Zero", vbInformation
        Exit Sub
    End If

    Dim ITM As ListItem
    Dim SNo As Integer
    If LV.ListItems.count > 0 Then
        SNo = Val(LV.ListItems(LV.ListItems.count).Text) + 1
    Else
        SNo = 1
    End If
    
    Set ITM = LV.ListItems.add(, cmbItems.Text & "'" & txtQuality.Text, SNo)
    ITM.Tag = cmbItems.Text 'cmbItems.List(cmbItems.ListIndex, 2)
    
    ITM.ListSubItems.add , , "" 'cmbCustItemID.Text
    ITM.ListSubItems.add , , cmbItems
    ITM.ListSubItems.add(, , cmbDesc).Tag = "" 'cmbCustItemID.Text
    
    ITM.ListSubItems.add , , txtUnit
    ITM.ListSubItems.add , , Format(DTDelivery, "dd-MMM-yy")
    ITM.ListSubItems.add , , txtQuality.Text
    
    ITM.ListSubItems.add(, , Format(txtPrice, "#0.0000")).Tag = Val(txtPrice.Tag)
    ITM.ListSubItems.add(, , txtQty).Tag = "New"
    ITM.ListSubItems.add , , lblAmt
    'ITM.ListSubItems.add , , Sno
   
    'cmbItemID.RemoveItem cmbItemID.ListIndex
    'cmbItem.RemoveItem cmbItem.ListIndex
    
    Dim iIndex As Integer
    
    cmbItemSearch.AddItem cmbItems.Text
    iIndex = cmbItemSearch.ListCount - 1
    cmbItemSearch.List(iIndex, 1) = cmbItems.Text & "'" & txtQuality.Text
    
    Call GetTotal

    txtQty = 0
    'cmbCustItemID.SetFocus
    
    
    ITM.EnsureVisible
    ITM.Selected = True
   
    cmbItems.SetFocus
    
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    
    If LV.SelectedItem.ListSubItems(6).Tag <> "New" Then
        'If Val(LV.SelectedItem.ListSubItems(6).Tag) <> Val(LV.SelectedItem.ListSubItems(6)) Then
        If Val(LV.SelectedItem.ListSubItems(6).Tag) <> 0 Then
            MsgBox "Can't Remove This Item.This Item Is In Use In Some Proforma.", vbInformation
            Exit Sub
        End If
    End If
    
    If Val(LV.SelectedItem.ListSubItems(9).Tag) <> 0 Then
        strDelList = strDelList & Val(LV.SelectedItem.ListSubItems(9).Tag) & ","
    End If
    
    LV.ListItems.Remove LV.SelectedItem.Index
    
End Sub

Private Sub cmdSClear_Click()
    If Saved Then
        Unload Me
        Me.Show
    End If
End Sub

Private Sub cmdSClose_Click()
    If Saved Then
        Unload Me
    End If
End Sub


Private Sub Form_Load()
   
    On Error GoTo err
    lQuotNo = 0
    
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where Active=1")
    OrderDT = ServerDate
    
    
    Me.KeyPreview = True
    cmbCust.ListHeight = 2500
    Call FillCmbs
    DTDelivery = Date
       
    If strCompany = "Tecno" Then
        txtOrderNo.Locked = False
    End If
    SSTab1.Tab = 0
    
    Exit Sub
err:
    MsgBox err.Description
   
End Sub

Public Sub AddNew()
    txtInternalRefNo = GetInternalRefNo
    Me.Show
End Sub

Private Function GetInternalRefNo() As Long
 
    On Error GoTo err
    
    Dim lRefNo As Long
    lRefNo = GetSingleLongValue("MAX(InternalRefNo)", "FCustomerOrders")
     
    If lRefNo = 0 Then
        lRefNo = Val(InputBox("Please Enter the Starting Number for Internal Ref. #", "Internal Reference No.", "596"))
    Else
        lRefNo = lRefNo + 1
    End If
    
    GetInternalRefNo = lRefNo
    
    Exit Function
err:
    MsgBox err.Description

    
End Function

Private Sub Form_Resize()
On Error Resume Next
   FAll.Move ((Me.ScaleWidth - FAll.Width) / 2), ((Me.ScaleHeight - FAll.Height) / 2)
End Sub

Private Sub lblPymtTerms_Click()
    Dim f As New frmMiscS
    Load f
    Call f.ShowForm("PaymentTerms", "Payment Terms", "Payment Terms")
End Sub

Private Sub LV_Click()
    If LV.ListItems.count > 0 Then
        Call LV_ItemClick(LV.SelectedItem)
    End If
End Sub

Private Sub LV_DblClick()
    
    If iColNo = 6 Then
        With DTEdit
            If LV.SelectedItem.SubItems(iColNo - 1) = "" Then
                DTEdit.Value = Date
            Else
                DTEdit.Value = LV.SelectedItem.SubItems(iColNo - 1)
            End If
            .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.SelectedItem.Top + LV.Top
            .Visible = True
            .SetFocus
        End With
    ElseIf iColNo = 8 Then
'        Dim strInvoiceNo As String
'        strInvoiceNo = UsedInPInvoice()
'        If strInvoiceNo <> "" Then
'            MsgBox "Article has been used in " & strInvoiceNo & ".Can't Edit Price Now.'", vbInformation
'            Exit Sub
'        End If
        'Editing of Price
        With txtEditPrice
            .Alignment = 1
            .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.SelectedItem.Top + LV.Top, LV.ColumnHeaders(iColNo).Width
            .Text = LV.SelectedItem.SubItems(iColNo - 1)
            .SelStart = 0
            .SelLength = Len(.Text)
            .Visible = True
            .SetFocus
        End With
    ElseIf iColNo = 9 Then
        With txtEditQty
            .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.SelectedItem.Top + LV.Top, LV.ColumnHeaders(iColNo).Width
            .Text = LV.SelectedItem.SubItems(iColNo - 1)
            .SelStart = 0
            .SelLength = Len(.Text)
            .Visible = True
            .SetFocus
        End With
    ElseIf iColNo = 7 Then
        With txtEditQuality
            .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.SelectedItem.Top + LV.Top, LV.ColumnHeaders(iColNo).Width
            .Text = LV.SelectedItem.SubItems(iColNo - 1)
            .SelStart = 0
            .SelLength = Len(.Text)
            .Visible = True
            .SetFocus
        End With
    Else
        With txtEditQty
            .Move LV.Left + LV.ColumnHeaders(8).Left, LV.SelectedItem.Top + LV.Top, LV.ColumnHeaders(8).Width
            .Text = LV.SelectedItem.SubItems(7)
            .SelStart = 0
            .SelLength = Len(.Text)
            .Visible = True
            .SetFocus
        End With
    End If
    
End Sub

Private Function UsedInPInvoice() As String

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim lRefID As Long
    lRefID = Val(LV.SelectedItem.ListSubItems(9).Tag)
    With rs
        .Open "Select Top 1 CustomInvoice From CustomInvoiceItems Where RefID In(Select EntryID From FProformaOrders Where OrderEntryID=" & lRefID & ")", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            UsedInPInvoice = ""
        Else
            UsedInPInvoice = .Fields(0)
        End If
        .Close
    End With
    
    Exit Function
    
err:
    MsgBox err.Description
    
End Function

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With cmdDelete
        .Move .Left, LV.Top + Item.Top
        .Visible = True
    End With
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        If x > LV.ColumnHeaders(i).Left And x < LV.ColumnHeaders(i).Left + LV.ColumnHeaders(i).Width Then
            iColNo = i
            Exit For
        End If
    Next
    
End Sub

Private Sub LVOtherCharges_KeyDown(KeyCode As Integer, Shift As Integer)

    If LVOtherCharges.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbYes Then
            LVOtherCharges.ListItems.Remove LVOtherCharges.SelectedItem.Index
        End If
    End If
    
End Sub

Private Sub OrderDT_Change()
    'DeliveryDT.Value = DateAdd("m", 2, OrderDT.Value)
End Sub

Private Sub txtEditPrice_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        txtEditPrice.Visible = False
    ElseIf KeyAscii = 13 Then
        LV.SelectedItem.SubItems(iColNo - 1) = Val(txtEditPrice)
        LV.SelectedItem.SubItems(9) = Val(txtEditPrice) * Val(LV.SelectedItem.SubItems(8))
        txtEditPrice.Visible = False
    End If

End Sub

Private Sub txtEditQty_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        txtEditQty.Visible = False
    ElseIf KeyAscii = 13 Then
        If iColNo = 9 Then
'            If Val(txtEditQty) <= 0 Then
'                MsgBox "Invalid Sort No.", vbInformation
'                txtEditQty.SetFocus
'                Exit Sub
'            End If
            LV.SelectedItem.SubItems(iColNo - 1) = Val(txtEditQty)
            If LV.SelectedItem.Index < LV.ListItems.count Then
                Dim iIndex As Integer
                iIndex = LV.SelectedItem.Index + 1
                LV.ListItems(iIndex).Selected = True
                With txtEditQty
                    .Top = LV.SelectedItem.Top + LV.Top
                    .Text = LV.SelectedItem.SubItems(iColNo - 1)
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
                End With
            Else
                txtEditQty.Visible = False
            End If
            'txtEditQty.Visible = False
        Else
            If Val(txtEditQty) <= 0 Or Val(txtEditQty) < Val(LV.SelectedItem.ListSubItems(6).Tag) Then
                MsgBox "Invalid Qty.", vbInformation
                txtEditQty.SetFocus
                Exit Sub
            End If
            LV.SelectedItem.SubItems(7) = Val(txtEditQty)
            LV.SelectedItem.SubItems(8) = Val(txtEditQty) * Val(LV.SelectedItem.SubItems(6))
            txtEditQty.Visible = False
        End If
    End If
    
End Sub

Private Sub txtEditQuality_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        txtEditQuality.Visible = False
    ElseIf KeyAscii = 13 Then
        LV.SelectedItem.SubItems(iColNo - 1) = txtEditQuality
        txtEditQuality.Visible = False
    End If
   
    
End Sub

Private Sub txtEditQuality_LostFocus()
    txtEditQuality.Visible = False
End Sub

Private Sub txtPrice_Change()
   lblAmt = Format(Val(txtQty) * Val(txtPrice), "#0.00")
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
   KeyAscii = OnlyNumber(KeyAscii, True)
End Sub

Private Sub txtQty_Change()
   lblAmt = Format(Val(txtQty) * Val(txtPrice), "#0.00")
End Sub

Private Sub txtQty_GotFocus()
    cmdAddItem.Default = True
    With txtQty
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub txtQty_LostFocus()
    cmdAddItem.Default = False
End Sub
Private Sub txtQty_KeyPress(KeyAscii As Integer)
   KeyAscii = OnlyNumber(KeyAscii)
End Sub


Private Sub GetTotal()

    Dim i As Integer
    Dim total As Double
    total = 0

    For i = 1 To LV.ListItems.count
        total = total + Val(LV.ListItems(i).ListSubItems(9).Text)
    Next i

    lblTotal = Format(total, "#0.00")
    lblTotWords = ConvertInWords(lblTotal, txtCurrency.Text)
    
End Sub



Private Function Saved() As Boolean
            
    On Error GoTo err
        
    Saved = False
    If cmbCompany.MatchFound = False Then
        MsgBox "Please Select Company Name.", vbInformation
        Exit Function
    ElseIf txtOrderNo = "" Then
        MsgBox "No Invoice No.", vbInformation
        txtOrderNo.SetFocus
        Exit Function
'    ElseIf cmbBankAccNo.MatchFound = False Then
'        MsgBox "No Bank Account Selected.", vbInformation
'        Exit Function
    End If
    Dim iBankID As Integer
    If cmbBankAccNo.MatchFound = False Then
        iBankID = -1
    Else
        iBankID = cmbBankAccNo.List(cmbBankAccNo.ListIndex, 1)
    End If
    If LV.ListItems.count = 0 Then
        MsgBox "No Items Added to This Sample Invoice", vbInformation
        Exit Function
    End If
    Dim iRecAffect_CC As Integer
    Call StartTrans(con)
    
    
    If lQuotNo <> 0 Then
        con.Execute "UPDATE CustomerQuotations SET DT='" & OrderDT & "',TradeTerms='" & cmbTradeTerms & "',CompanyRefID=" & _
        cmbCompany.ID & ",Validity='" & txtValidity & "',PaymentTerms='" & cmbPaymentTerms & "',ShipVia='" & cmbShipVia & _
        "',Packing='" & txtPacking & "',DeliveryTime='" & txtDeliveryTime & _
        "',BankID=" & iBankID & ",Remarks='" & txtRemarks & "',CustomerRef='" & txtCustRef & "',Department='" & txtDepartment & _
        "',AttentionPerson='" & txtAttention & "',SwiftCode='" & txtSwiftCode & "',GrossWeight='" & txtGrossWeight & _
        "',NetWeight='" & txtNetWeight & "',NoOfPackages=" & Val(txtNoOfPackages) & ",Destination='" & txtDestination & "' WHERE QuotationNo=" & lQuotNo
             
        If strDelList <> "" Then
            strDelList = Left(strDelList, Len(strDelList) - 1)
            con.Execute "DELETE FROM CustomerQuotationsItems WHERE EntryID In(" & strDelList & ")"
        End If
        
    Else
        con.Execute "INSERT INTO CustomerQuotations(CustCode,Country,QuotationNo,DT,TradeTerms,CompanyRefID,Validity" & _
        ",PaymentTerms,ShipVia,Packing,DeliveryTime,BankID,UserName,MachineName,Remarks,CustomerRef,Department,AttentionPerson," & _
        "SwiftCode,GrossWeight,NetWeight,NoOfPackages,Destination) VALUES('" & cmbCust.ID & "','" & cmbCountry & "'," & txtOrderNo & ",'" & _
        OrderDT & "','" & cmbTradeTerms & "'," & cmbCompany.ID & ",'" & txtValidity & "','" & cmbPaymentTerms & "','" & cmbShipVia & _
        "','" & txtPacking & "','" & txtDeliveryTime & "'," & iBankID & ",'" & CurrentUserName & "','" & strComputerName & "','" & _
        txtRemarks & "','" & txtCustRef & "','" & txtDepartment & "','" & txtAttention & "','" & txtSwiftCode & "','" & _
        txtGrossWeight & "','" & txtNetWeight & "'," & Val(txtNoOfPackages) & ",'" & txtDestination & "')"
    End If
   
    'Now Save Items
    Dim strItemCode As String
    Dim cmd As New ADODB.Command
    For i = 1 To LV.ListItems.count
        strItemCode = Left(LV.ListItems(i).key, Len(LV.ListItems(i).key) - 1)
SaveNewItem:
'            If strCompany <> "All Seas" Then
'                'Assign to Customer....
'                With cmd
'                    Set .ActiveConnection = con
'                    .CommandType = adCmdStoredProc
'                    .CommandText = "SP_AssItem"
'                    .Parameters("@CustCode").Value = cmbCust.ID
'                    .Parameters("@CustCountry").Value = cmbCountry
'                    .Parameters("@ItemID").Value = LV.ListItems(i).Tag
'                    .Execute Saved
'                End With
'
'            End If
            If Val(LV.ListItems(i).ListSubItems(9).Tag) = 0 Then
                con.Execute "INSERT INTO CustomerQuotationsItems(QuotationNo,ItemCode,DeliveryDT,QualityType," & _
                 "Price,Qty) VALUES(" & txtOrderNo & ",'" & LV.ListItems(i).Tag & "','" & LV.ListItems(i).SubItems(5) & "','" & _
                 LV.ListItems(i).SubItems(6) & "'," & Val(LV.ListItems(i).SubItems(7)) & "," & Val(LV.ListItems(i).SubItems(8)) & ")"
                'Now Check If Item Present in Customer Catalog...
                con.Execute "UPDATE FCustomerCatalog SET ExWorks=" & Val(LV.ListItems(i).SubItems(7)) & " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "' AND CompItemID='" & LV.ListItems(i).Tag & "'", iRecAffect_CC
                If iRecAffect_CC = 0 Then
                    con.Execute "INSERT INTO FCustomerCatalog(CustCode,Country,CompItemID,ItemID,Description,ExWorks) VALUES('" & cmbCust.ID & "','" & cmbCountry.Text & _
                     "','" & LV.ListItems(i).Tag & "','" & LV.ListItems(i).Tag & "','" & LV.ListItems(i).SubItems(3) & "'," & Val(LV.ListItems(i).SubItems(7)) & ")"
                End If
            Else
                con.Execute "UPDATE CustomerQuotationsItems SET Qty=" & Val(LV.ListItems(i).SubItems(8)) & _
                 ",Price=" & Val(LV.ListItems(i).SubItems(7)) & ",DeliveryDT='" & LV.ListItems(i).SubItems(5) & _
                 "',QualityType='" & LV.ListItems(i).SubItems(6) & "' WHERE EntryID=" & Val(LV.ListItems(i).ListSubItems(9).Tag)
            End If
    Next
    
    
    'Other Charges....
    If lQuotNo <> 0 Then
        con.Execute "DELETE FROM CustomerQuotations_OtherCharges WHERE QuotationNo=" & lQuotNo
    End If
    
    For i = 1 To LVOtherCharges.ListItems.count
        With LVOtherCharges.ListItems(i)
            con.Execute "INSERT INTO CustomerQuotations_OtherCharges(QuotationNo,ChargesDetail,ChargesAmt,SNo)" & _
             " VALUES(" & txtOrderNo & ",'" & .SubItems(1) & "'," & Val(.SubItems(2)) & "," & Val(.Index) & ")"
        End With
    Next
    
    con.CommitTrans
    
    Saved = True
     
    Exit Function

err:
    MsgBox err.Description
    
End Function

Private Sub PrintOrder(lQuotNo As String)

    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptFCustOrders.rpt") '
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{CustOrders.OrderNo}='" & lQuotNo & "' ", rpt
    
    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub FillCmbs()
    
    Call AddToCombo(cmbTradeTerms, "Description", "TradeTerms", " Order By ID")
    Call AddToCombo(cmbPaymentTerms, "Description", "PaymentTerms", " Order By ID")
    
    Call AddToCombo(cmbShipVia, "Description", "ShippingMethods", " Order By EntryID")
    
'    With cmbShipVia
'        .AddItem "By Sea"
'        .AddItem "By Air"
'    End With
    
    cmbCompany.ListHeight = 6000
    cmbCompany.AddVals con, "CompanyName", "Companies", "EntryID", " Order By CompanyName"
    cmbCompany.ID = 1
    
End Sub

Private Sub FillItems()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    
    With rs
        
        If 0 = 0 Then
'            If strCompany = "All Seas" Then
                .Open "SELECT ItemID,ItemID AS CompItemID,ItemName AS Description,'' AS PackingMode,FOB,ExWorks" & _
                 ",CIFAir,CIFSea,CnFAir,CnFSea,CnIAir,CnISea,FOB AS CustomFOB,ExWorks AS CustomExWorks,CIFAir AS CustomCIFAir,CIFSea AS CustomCIFSea,CnFAir AS CustomCnFAir,CnFSea AS CustomCnFSea,CnIAir AS CustomCnIAir,CnISea AS CustomCnISea FROM Items", con_ServerSide, adOpenForwardOnly, adLockReadOnly
'            Else
'                .Open "SELECT * FROM VFCustomerCatalog WHERE CustCode+Country='" & cmbCust.Text & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
'            End If
        Else
'            If strCompany = "All Seas" Then
                .Open "SELECT ItemID,ItemID AS CompItemID,ItemName AS Description,Unit AS PackingMode,1 AS FOB,1 AS ExWorks" & _
                 ",1 AS CIFAir,1 AS CIFSea,1 AS CnFAir,1 AS CnFSea,1 AS CnIAir,1 AS CnISea,1 AS CustomFOB,1 AS CustomExWorks,1 AS CustomCIFAir,1 AS CustomCIFSea,1 AS CustomCnFAir,1 AS CustomCnFSea,1 AS CustomCnIAir,1 AS CustomCnISea FROM Items WHERE CatID=" & cmbCatID.List(cmbCatID.ListIndex, 1), con, adOpenForwardOnly, adLockReadOnly
'            Else
'                .Open "SELECT * FROM VFCustomerCatalog WHERE CatID=" & cmbCatID.List(cmbCatID.ListIndex, 1) & " AND CustCode+Country='" & cmbCust.Text & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
'            End If
        End If
        cmbItems.Clear
        cmbDesc.Clear
        'cmbCustItemID.Clear
        Do Until .EOF
            'cmbCustItemID.AddItem !ITemID & ""
            If chkCust Then
                cmbItems.AddItem ![ITemID] & ""
                cmbDesc.AddItem ![Description] & ""
            Else
                cmbItems.AddItem ![CompITemID] & ""
                'cmbDesc.AddItem ![CompDesc] & ""
                cmbDesc.AddItem ![Description] & ""
            End If
            
            'cmbItems.List(cmbItems.ListCount - 1, 1) = ![Unit] & ""
            cmbItems.List(cmbItems.ListCount - 1, 1) = Val(![PackingMode] & "")
            
            cmbItems.List(cmbItems.ListCount - 1, 2) = Val(![FOB] & "")
            cmbItems.List(cmbItems.ListCount - 1, 3) = Val(![ExWorks] & "")
            
            cmbItems.List(cmbItems.ListCount - 1, 4) = Val(![CIFAir] & "")
            cmbItems.List(cmbItems.ListCount - 1, 5) = Val(![CIFSea] & "")
            
            cmbItems.List(cmbItems.ListCount - 1, 6) = Val(![CnFAir] & "")
            cmbItems.List(cmbItems.ListCount - 1, 7) = Val(![CnFSea] & "")
            
            cmbItems.List(cmbItems.ListCount - 1, 8) = Val(![CnIAir] & "")
            cmbItems.List(cmbItems.ListCount - 1, 9) = Val(![CnISea] & "")
            
            'cmbDesc.List(cmbDesc.ListCount - 1, 1) = ![Description] & ""
            'cmbDesc.List(cmbDesc.ListCount - 1, 2) = ![CompDesc] & ""
            
            cmbDesc.List(cmbDesc.ListCount - 1, 2) = Val(![CustomFOB] & "")
            cmbDesc.List(cmbDesc.ListCount - 1, 3) = Val(![CustomExWorks] & "")
            cmbDesc.List(cmbDesc.ListCount - 1, 4) = Val(![CustomCIFAir] & "")
            cmbDesc.List(cmbDesc.ListCount - 1, 5) = Val(![CustomCIFSea] & "")
            
            cmbDesc.List(cmbDesc.ListCount - 1, 6) = Val(![CustomCnFAir] & "")
            cmbDesc.List(cmbDesc.ListCount - 1, 7) = Val(![CustomCnFSea] & "")
            
            cmbDesc.List(cmbDesc.ListCount - 1, 8) = Val(![CustomCnIAir] & "")
            cmbDesc.List(cmbDesc.ListCount - 1, 9) = Val(![CustomCnISea] & "")
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub cmdAddOtherCharges_Click()
    
    On Error GoTo err
    
    If Val(txtOtherChargesSNo) > LVOtherCharges.ListItems.count + 1 Then
        MsgBox "Invalid S/No.", vbCritical
        txtOtherChargesSNo.Text = LVOtherCharges.ListItems.count + 1
        txtOtherChargesSNo.SetFocus
        Exit Sub
    End If
    
    If txtOtherCharges.Text = "" Then
        MsgBox "Please Enter Other Charges Description.", vbCritical
        txtOtherCharges.SetFocus
    End If
    
    If IsNumeric(txtOtherChargeAmt.Text) = False Then
        MsgBox "Invalid Amount.", vbCritical
        txtOtherChargeAmt.SelStart = 0
        txtOtherChargeAmt.SelLength = Len(txtOtherChargeAmt.Text)
        txtOtherChargeAmt.SetFocus
    End If
    
    Dim ITM As ListItem
    Set ITM = LVOtherCharges.ListItems.add(Val(txtOtherChargesSNo), , Val(txtOtherChargesSNo))
    ITM.SubItems(1) = txtOtherCharges.Text & ""
    ITM.SubItems(2) = Val(txtOtherChargeAmt)
    
    Call PrepareForNext
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub PrepareForNext()

    Dim i As Integer
    For i = 1 To LVOtherCharges.ListItems.count
        LVOtherCharges.ListItems(i).Text = i
    Next
    txtOtherChargesSNo.Text = LVOtherCharges.ListItems.count + 1
    txtOtherCharges.Text = ""
    txtOtherChargeAmt.Text = ""
    txtOtherCharges.SetFocus
    
End Sub

