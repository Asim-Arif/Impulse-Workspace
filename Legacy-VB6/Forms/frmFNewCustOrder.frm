VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmFNewCustOrder 
   ClientHeight    =   8625
   ClientLeft      =   -120
   ClientTop       =   -120
   ClientWidth     =   11700
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
   ScaleHeight     =   8625
   ScaleWidth      =   11700
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
      Height          =   8520
      Left            =   45
      TabIndex        =   18
      Top             =   30
      Width           =   11640
      Begin ComboList.Usercontrol1 cmbCompany 
         Height          =   285
         Left            =   285
         TabIndex        =   28
         Top             =   945
         Width           =   5160
         _ExtentX        =   9102
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
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   285
         TabIndex        =   0
         Top             =   1455
         Width           =   1770
         _ExtentX        =   3122
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
         Height          =   5460
         Left            =   255
         TabIndex        =   30
         Top             =   1770
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   9631
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Article Detail"
         TabPicture(0)   =   "frmFNewCustOrder.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1(1)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label1(2)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblTotWords"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lblTotal"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "LV"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "PicProduct"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "txtEditPrice"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txtEditQty"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "FNew"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "cmdDelete"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "cmdAddFromQuotation"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).ControlCount=   11
         TabCaption(1)   =   "Other Detail"
         TabPicture(1)   =   "frmFNewCustOrder.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtRevisionHistory"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "txtPackaging"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label4(1)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label4(0)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).ControlCount=   4
         Begin VB.TextBox txtRevisionHistory 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   1680
            Left            =   -74775
            MultiLine       =   -1  'True
            TabIndex        =   56
            Top             =   3150
            Width           =   10635
         End
         Begin VB.CommandButton cmdAddFromQuotation 
            Caption         =   "Add Items from Quotation"
            Height          =   375
            Left            =   75
            TabIndex        =   54
            Top             =   990
            Width           =   2220
         End
         Begin VB.TextBox txtPackaging 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   1680
            Left            =   -74775
            MultiLine       =   -1  'True
            TabIndex        =   50
            Top             =   1065
            Width           =   10635
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
            Left            =   10650
            Picture         =   "frmFNewCustOrder.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Remove This Entry From Voucher"
            Top             =   990
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.Frame FNew 
            BorderStyle     =   0  'None
            Height          =   570
            Left            =   75
            TabIndex        =   34
            Top             =   375
            Width           =   11100
            Begin VB.TextBox txtQty 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               DataField       =   "RptTime"
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   8370
               MaxLength       =   150
               TabIndex        =   38
               Top             =   255
               Width           =   930
            End
            Begin VB.TextBox txtPrice 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               DataField       =   "RptTime"
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   7410
               MaxLength       =   150
               TabIndex        =   37
               Top             =   255
               Width           =   930
            End
            Begin VB.TextBox txtUnit 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               DataField       =   "RptTime"
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   5280
               Locked          =   -1  'True
               MaxLength       =   150
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   255
               Width           =   930
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
               Left            =   10560
               Picture         =   "frmFNewCustOrder.frx":056A
               Style           =   1  'Graphical
               TabIndex        =   35
               ToolTipText     =   "Add This Entry To Voucher"
               Top             =   15
               UseMaskColor    =   -1  'True
               Width           =   525
            End
            Begin MSForms.ComboBox cmbDesc 
               Height          =   285
               Left            =   2820
               TabIndex        =   43
               Top             =   240
               Width           =   2415
               VariousPropertyBits=   746604575
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "4260;503"
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
               Left            =   1590
               TabIndex        =   42
               Top             =   240
               Width           =   1215
               VariousPropertyBits=   746604575
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "2143;503"
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
               Left            =   9330
               TabIndex        =   41
               Top             =   255
               Width           =   1200
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BorderStyle     =   1  'Fixed Single
               Caption         =   $"frmFNewCustOrder.frx":0A9C
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
               TabIndex        =   40
               Top             =   0
               Width           =   10530
            End
            Begin MSForms.ComboBox cmbCustItemID 
               Height          =   285
               Left            =   0
               TabIndex        =   7
               Top             =   255
               Width           =   1575
               VariousPropertyBits=   746604575
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "2778;503"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbQuality 
               Height          =   285
               Left            =   6210
               TabIndex        =   39
               Top             =   255
               Width           =   1185
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "2090;503"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin VB.TextBox txtEditQty 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5175
            TabIndex        =   33
            Top             =   1740
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtEditPrice 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3945
            TabIndex        =   32
            Top             =   1740
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.PictureBox PicProduct 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   2865
            Left            =   75
            ScaleHeight     =   2835
            ScaleWidth      =   2760
            TabIndex        =   31
            Top             =   1425
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
         Begin MSComctlLib.ListView LV 
            Height          =   3720
            Left            =   75
            TabIndex        =   45
            Top             =   1425
            Width           =   10545
            _ExtentX        =   18600
            _ExtentY        =   6562
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
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S No."
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Cust ID"
               Object.Width           =   1147
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Item ID"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Description"
               Object.Width           =   5997
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "P. Unit"
               Object.Width           =   1147
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Price"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "Qty"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "Amount"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   8
               Text            =   "Sort"
               Object.Width           =   1764
            EndProperty
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Revision History Box :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   -74760
            TabIndex        =   57
            Top             =   2910
            Width           =   1830
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Special Instructions :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   -74760
            TabIndex        =   51
            Top             =   825
            Width           =   1770
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E7EBEF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   9315
            TabIndex        =   49
            Top             =   5130
            Width           =   1305
         End
         Begin VB.Label lblTotWords 
            Appearance      =   0  'Flat
            BackColor       =   &H00E7EBEF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1890
            TabIndex        =   48
            Top             =   5130
            Width           =   6570
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
            TabIndex        =   47
            Top             =   5130
            Width           =   1830
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
            Left            =   8445
            TabIndex        =   46
            Top             =   5130
            Width           =   885
         End
      End
      Begin VB.TextBox txtOrderRcvdvia 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   10080
         MaxLength       =   150
         TabIndex        =   6
         Top             =   945
         Width           =   1470
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
         Left            =   270
         TabIndex        =   29
         Top             =   8100
         Width           =   1560
      End
      Begin VB.TextBox txtStamps 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   5430
         TabIndex        =   13
         Top             =   945
         Width           =   3210
      End
      Begin VB.TextBox txtInternalRefNo 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   8625
         MaxLength       =   150
         TabIndex        =   15
         Top             =   945
         Width           =   1470
      End
      Begin VB.TextBox txtQuality 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   5430
         TabIndex        =   12
         Top             =   945
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.TextBox txtCurrency 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3810
         MaxLength       =   150
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1455
         Width           =   1005
      End
      Begin VB.TextBox txtOrderNo 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4845
         MaxLength       =   150
         TabIndex        =   3
         Top             =   1455
         Width           =   1470
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
         TabIndex        =   21
         Top             =   0
         Width           =   11640
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "New Order"
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
            TabIndex        =   22
            Top             =   165
            Width           =   11505
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "New Order"
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
            TabIndex        =   23
            Top             =   180
            Width           =   11490
         End
      End
      Begin MSComCtl2.DTPicker OrderDT 
         Height          =   300
         Left            =   6345
         TabIndex        =   4
         Top             =   1455
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   197459971
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
         Left            =   4470
         TabIndex        =   19
         Top             =   7785
         Width           =   7020
         Begin MSForms.CommandButton cmdSavenStay 
            Height          =   360
            Left            =   3495
            TabIndex        =   55
            Top             =   195
            Width           =   1665
            Caption         =   "Save & Stay"
            PicturePosition =   327683
            Size            =   "2937;635"
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
            Left            =   5205
            TabIndex        =   17
            Top             =   195
            Width           =   1665
            Caption         =   "Cancel            "
            PicturePosition =   327683
            Size            =   "2937;635"
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
            Left            =   90
            TabIndex        =   14
            Top             =   195
            Width           =   1665
            Caption         =   "Save & New   "
            PicturePosition =   327683
            Size            =   "2937;635"
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
            Left            =   1800
            TabIndex        =   16
            Top             =   195
            Width           =   1665
            Caption         =   "Save & Close  "
            PicturePosition =   327683
            Size            =   "2937;635"
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
      Begin MSComCtl2.DTPicker DeliveryDT 
         Height          =   285
         Left            =   9480
         TabIndex        =   11
         Top             =   7485
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   503
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
         CheckBox        =   -1  'True
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   197459971
         CurrentDate     =   37250
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmFNewCustOrder.frx":0B44
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
         Left            =   285
         TabIndex        =   53
         Top             =   705
         Width           =   11265
      End
      Begin MSForms.ComboBox cmbCatID 
         Height          =   285
         Left            =   0
         TabIndex        =   52
         Top             =   345
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
      Begin VB.Label lblPymtTerms 
         Alignment       =   2  'Center
         Caption         =   "This Is Not Visible"
         Height          =   165
         Left            =   345
         MouseIcon       =   "frmFNewCustOrder.frx":0BFD
         MousePointer    =   99  'Custom
         TabIndex        =   27
         Top             =   7290
         Width           =   1335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmFNewCustOrder.frx":0D4F
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
         Left            =   285
         TabIndex        =   26
         Top             =   7245
         Width           =   11205
      End
      Begin MSForms.CheckBox chkPrint 
         Height          =   345
         Left            =   270
         TabIndex        =   25
         Top             =   7830
         Width           =   1545
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   4
         Size            =   "2725;609"
         Value           =   "0"
         Caption         =   "Print On Save"
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbPayment 
         Height          =   285
         Left            =   285
         TabIndex        =   8
         Top             =   7500
         Width           =   3045
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5371;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbTrans 
         Height          =   285
         Left            =   3360
         TabIndex        =   9
         Top             =   7500
         Width           =   3045
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5371;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkCust 
         Height          =   360
         Left            =   270
         TabIndex        =   24
         Top             =   8085
         Visible         =   0   'False
         Width           =   2685
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   4
         Size            =   "4736;635"
         Value           =   "0"
         Caption         =   "Show Customer Item Codes"
         Accelerator     =   73
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbTradeTerms 
         Height          =   285
         Left            =   7920
         TabIndex        =   5
         Top             =   1455
         Width           =   3630
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "6403;503"
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
         Caption         =   $"frmFNewCustOrder.frx":0DFC
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
         Left            =   285
         TabIndex        =   20
         Top             =   1215
         Width           =   11265
      End
      Begin MSForms.ComboBox cmbPartial 
         Height          =   285
         Left            =   6435
         TabIndex        =   10
         Top             =   7500
         Width           =   3045
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5371;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   2070
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1455
         Width           =   1725
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3043;503"
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
Attribute VB_Name = "frmFNewCustOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrderNo As String
Dim strDelList As String
Dim iColNo As Integer
Dim bCopy As Boolean
Dim strCopyFromOrderNo As String

Public Sub EditOrder(TempOrdNo As String)
    
    On Error GoTo err
    strOrderNo = TempOrdNo
    strDelList = ""
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim strOrderBy As String
    If strCompany = "QEL" Then
        strOrderBy = "ORDER BY CompItemID"
    Else
        strOrderBy = ""
    End If
    With rs
        
        .Open "SELECT * FROM VrptOrders WHERE OrderNo='" & strOrderNo & "' " & strOrderBy, con, adOpenForwardOnly, adLockReadOnly
        
        cmbCompany.ID = !CompanyRefID
        cmbCust.Text = ![CustCode] & ""
        cmbCountry = ![Country] & ""
        txtOrderNo = ![OrderNo] & ""
        OrderDT = ![DT] & ""
        cmbTradeTerms = ![TradeTerms] & ""
        
        cmbPayment = ![PaymentTerms] & ""
        cmbTrans = ![TransShipment] & ""
        cmbPartial = ![PartialShipment] & ""
        txtPackaging = ![Packaging] & ""
        DeliveryDT = ![DeliveryDT] & ""
        txtQuality = !Quality & ""
        txtInternalRefNo = Val(!InternalRefNo & "")
        txtStamps = !StampDT & ""
        txtOrderRcvdvia = !OrderRcvdvia & ""
        txtRevisionHistory = !RevisionHistoryBox & ""
        
'        If IsNull(!StampDT) = False Then
'            DTStamp = !StampDT
'        End If
        LV.ListItems.Clear
        Do Until .EOF
        
            Set ITM = LV.ListItems.add(, ![ITemID] & "'", .AbsolutePosition)
            
            ITM.Tag = ![CompITemID] & "" 'cmbItems.List(cmbItems.ListIndex, 2)
            
            ITM.ListSubItems.add , , !ITemID & "" '![CatID] & ""
            ITM.ListSubItems.add , , ![CompITemID] & ""
            ITM.ListSubItems.add(, , ![ItemName] & "").Tag = ![Description] & ""
            
            ITM.ListSubItems.add , , !Unit & ""
            
            ITM.ListSubItems.add(, , Format(![Price] & "", "#0.0000")).Tag = ![CustomPrice] & ""
            ITM.ListSubItems.add(, , Val(![Qty] & "")).Tag = Val(![Qty] & "") - Val(![InvQty] & "")
            
            ITM.ListSubItems.add(, , Val(![Qty] & "") * Val(![Price])).Tag = ![ID] & ""
            
            ITM.ListSubItems.add , , Val(![SortNo] & "")
            
            If bCopy Then
                ITM.ListSubItems(6).Tag = ""
                ITM.ListSubItems(7).Tag = ""
            End If
            
            .MoveNEXT
        Loop
        .Close
        .Open "Select Count(*) From FProformaOrders Where OrderEntryID In(Select ID From FOrderItems Where OrderNo='" & strOrderNo & "')", con, adOpenForwardOnly, adLockReadOnly
        If .Fields(0).Value = 0 Then
            cmbCompany.Enabled = True
        Else
            cmbCompany.Enabled = False
        End If
        .Close
    End With
    
    Call GetTotal
        
    If bCopy Then
        strOrderNo = ""
        txtOrderNo = ""
        OrderDT = Date
        Call OrderDT_Change
    End If
    
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

Private Sub chkShowPicture_Click()
    If chkShowPicture.Value = vbChecked Then
        Call cmbItems_Change
    Else
        PicProduct.Visible = False
    End If
End Sub

'Private Sub cmbCatID_Change()
'    Call cmbCatID_Click
'End Sub
'
'Private Sub cmbCatID_Click()
'    If cmbCatID.MatchFound = False Then Exit Sub
'    Call FillItems
'End Sub

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
        Call FillItems
        'Do Until .EOF
       '     cmbCatID.AddItem .Fields(1) & ""
      '      cmbCatID.List(cmbCatID.ListCount - 1, 1) = .Fields(0) & ""
     '       .MoveNext
    '    Loop
    '    .Close
        .Open "Select TradeTerms,PaymentTerms From ForeignCustomers Where CustCode+Country='" & cmbCust.Text & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbTradeTerms.Text = !TradeTerms & ""
        cmbPayment.Text = !PaymentTerms & ""
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
    If cmbCust.Text = "HICO" Then
        txtOrderNo = "HICO Order"
        txtOrderNo.Locked = True
    Else
        txtOrderNo = ""
        txtOrderNo.Locked = False
    End If
    
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
    ''''''''Show Picture
    If chkShowPicture.Value = vbChecked Then
        Dim rs As New ADODB.Recordset
        With rs
            .Open "SELECT ItemPic FROM Items WHERE ItemID='" & cmbItems.Text & "'", con, adOpenForwardOnly, adLockReadOnly
            If Not IsNull(![ItemPic]) Then
                Set ImgProduct.DataSource = rs
                ImgProduct.DataField = "ItemPic"
            Else
                ImgProduct.Picture = LoadPicture("")
            End If
            .Close
        End With
        Set rs = Nothing
        PicProduct.Visible = True
    End If
End Sub

Private Sub cmbItems_Click()

    If cmbItems.MatchFound Then
        
        cmbCustItemID.ListIndex = cmbItems.ListIndex
        cmbDesc.ListIndex = cmbItems.ListIndex
        txtUnit = cmbItems.List(cmbItems.ListIndex, 1)
        If cmbQuality.ListIndex <= 0 Then
            txtPrice = Val(cmbItems.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 2) & "")
        Else    'Top Quality
            txtPrice = Val(cmbDesc.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 2) & "") 'Val(cmbDesc.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 2) & "")
        End If
        
    Else
        txtUnit = ""
        txtPrice = ""
        txtPrice.Tag = ""
    End If
    
End Sub

Private Sub cmbItems_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode.Value = 13 Then
        txtPrice.SetFocus
    End If
End Sub

Private Sub cmbQuality_Change()
    Call cmbQuality_Click
End Sub

Private Sub cmbQuality_Click()
    If cmbItems.MatchFound = False Then Exit Sub
    If cmbQuality.ListIndex <= 0 Then
        txtPrice = Val(cmbItems.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 2) & "")
    Else    'Top Quality
        txtPrice = Val(cmbDesc.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 2) & "") 'Val(cmbDesc.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 2) & "")
    End If
End Sub

Private Sub cmbTradeTerms_Change()
    Call cmbTradeTerms_Click
End Sub

Private Sub cmbTradeTerms_Click()

    If cmbTradeTerms.MatchFound Then
        cmbCatID.Locked = False
        cmbItems.Locked = False
        cmbCustItemID.Locked = False
        cmbDesc.Locked = False
        If cmbItems.MatchFound Then
            txtPrice = cmbItems.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 2)
        End If
    Else
        cmbCatID.Locked = True
        cmbItems.Locked = True
        cmbCustItemID.Locked = True
        cmbDesc.Locked = True
    End If
    
End Sub

Private Sub cmdAddFromQuotation_Click()

    Dim f As New frmGetSingleTextValue
    Dim strQuotationNo As String
    Call AddToCombo(f.cmbComments, "QuotationNo", "CustomerQuotations", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'")
    If f.getTextValue(strQuotationNo, True, "QuotationNo", "Quotations") = False Then Exit Sub
        
        
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim SNo As Integer
    
    With rs
        .Open "SELECT * FROM VCustomerQuotations WHERE QuotationNo=" & Val(strQuotationNo), con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            If LV.ListItems.count > 0 Then
                SNo = Val(LV.ListItems(LV.ListItems.count).Text) + 1
            Else
                SNo = 1
            End If
            Set ITM = LV.ListItems.add(, !ItemCode & "'", SNo)
            ITM.Tag = !ItemCode & ""
            ITM.ListSubItems.add , , !ITemID & ""
            ITM.ListSubItems.add , , !ItemCode & ""
            ITM.ListSubItems.add(, , !ItemName & "").Tag = !ItemName & ""
            ITM.ListSubItems.add , , !Unit & ""
            ITM.ListSubItems.add(, , Format(Val(!Price & ""), "#0.0000")).Tag = Val(!Price & "")
            ITM.ListSubItems.add(, , Val(!Qty & "")).Tag = "New"
            ITM.ListSubItems.add , , Val(!Qty & "") * Val(!Price & "")
            ITM.ListSubItems.add , , SNo
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub cmdSavenStay_Click()

    If Saved Then
        Dim strOrderNo As String
        strOrderNo = txtOrderNo
        Call EditOrder(strOrderNo)
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub cmdAddItem_Click()
    
    On Error GoTo err
    
    If Not cmbItems.MatchFound Or cmbItems = "" Then
        MsgBox "Invalid Item Selected", vbInformation
        Exit Sub
    'ElseIf Val(txtPrice) = 0 Then
    '    MsgBox "Price Cannot Be Zero", vbInformation
    '    Exit Sub
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
    
    Set ITM = LV.ListItems.add(, cmbCustItemID.Text & "'", SNo)
    ITM.Tag = cmbItems.Text 'cmbItems.List(cmbItems.ListIndex, 2)
    
    ITM.ListSubItems.add , , cmbCustItemID.Text
    ITM.ListSubItems.add , , cmbItems
    ITM.ListSubItems.add(, , cmbDesc).Tag = cmbCustItemID.Text
    
    ITM.ListSubItems.add , , txtUnit
    
    ITM.ListSubItems.add(, , Format(txtPrice, "#0.0000")).Tag = Val(txtPrice.Tag)
    ITM.ListSubItems.add(, , txtQty).Tag = "New"
    ITM.ListSubItems.add , , lblAmt
    ITM.ListSubItems.add , , SNo
   
    'cmbItemID.RemoveItem cmbItemID.ListIndex
    'cmbItem.RemoveItem cmbItem.ListIndex
    
    Call GetTotal

    txtQty = 0
    cmbCustItemID.SetFocus
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    
'    If LV.SelectedItem.ListSubItems(6).Tag <> "New" Then
'        If Val(LV.SelectedItem.ListSubItems(6).Tag) <> 0 Then
'            MsgBox "Can't Remove This Item.This Item Is In Use In Some Proforma.", vbInformation
'            Exit Sub
'        End If
'    End If
    
    If Val(LV.SelectedItem.ListSubItems(7).Tag) <> 0 Then
        strDelList = strDelList & Val(LV.SelectedItem.ListSubItems(7).Tag) & ","
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
    strOrderNo = ""
    strCopyFromOrderNo = ""
    bCopy = False
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where Active=1")
    OrderDT = ServerDate
    DeliveryDT = DateAdd("m", 2, ServerDate)
    
    If strCompany = "IAA" Or strCompany = "Kami" Then
        DeliveryDT.Value = Null
    End If
    
    DTStamp = ServerDate
    Me.KeyPreview = True
    cmbCust.ListHeight = 2500
    Call FillCmbs
    
    With lblPymtTerms
        .Caption = ""
        .BackStyle = 0
    End With
    
    strDelList = ""
     
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
'        Dim strInvoiceNo As String
'        strInvoiceNo = UsedInPInvoice()
'        If strInvoiceNo <> "" Then
'            MsgBox "Article has been used in " & strInvoiceNo & ".Can't Edit Price Now.'", vbInformation
'            Exit Sub
'        End If
        'Editing of Price
        With txtEditPrice
            .Alignment = 1
            .Move LV.Left + LV.ColumnHeaders(6).Left, LV.SelectedItem.Top + LV.Top, LV.ColumnHeaders(6).Width
            .Text = LV.SelectedItem.SubItems(5)
            .SelStart = 0
            .SelLength = Len(.Text)
            .Visible = True
            .SetFocus
        End With
    ElseIf iColNo = 9 Then
        With txtEditQty
            .Move LV.Left + LV.ColumnHeaders(9).Left, LV.SelectedItem.Top + LV.Top, LV.ColumnHeaders(9).Width
            .Text = LV.SelectedItem.SubItems(8)
            .SelStart = 0
            .SelLength = Len(.Text)
            .Visible = True
            .SetFocus
        End With
    Else
        With txtEditQty
            .Move LV.Left + LV.ColumnHeaders(7).Left, LV.SelectedItem.Top + LV.Top, LV.ColumnHeaders(7).Width
            .Text = LV.SelectedItem.SubItems(6)
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
    lRefID = Val(LV.SelectedItem.ListSubItems(7).Tag)
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

Private Sub OrderDT_Change()
    DeliveryDT.Value = DateAdd("m", 2, OrderDT.Value)
End Sub

Private Sub txtEditPrice_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        txtEditPrice.Visible = False
    ElseIf KeyAscii = 13 Then
        LV.SelectedItem.SubItems(5) = Val(txtEditPrice)
        LV.SelectedItem.SubItems(7) = Val(txtEditPrice) * Val(LV.SelectedItem.SubItems(6))
        txtEditPrice.Visible = False
    End If

End Sub

Private Sub txtEditQty_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        txtEditQty.Visible = False
    ElseIf KeyAscii = 13 Then
        If iColNo = 9 Then
            If Val(txtEditQty) <= 0 Then
                MsgBox "Invalid Sort No.", vbInformation
                txtEditQty.SetFocus
                Exit Sub
            End If
            LV.SelectedItem.SubItems(8) = Val(txtEditQty)
            If LV.SelectedItem.Index < LV.ListItems.count Then
                Dim iIndex As Integer
                iIndex = LV.SelectedItem.Index + 1
                LV.ListItems(iIndex).Selected = True
                With txtEditQty
                    .Top = LV.SelectedItem.Top + LV.Top
                    .Text = LV.SelectedItem.SubItems(8)
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
                End With
            Else
                txtEditQty.Visible = False
            End If
            'txtEditQty.Visible = False
        Else
'            If Val(txtEditQty) <= 0 Or Val(txtEditQty) < Val(LV.SelectedItem.ListSubItems(6).Tag) Then
'                MsgBox "Invalid Qty.", vbInformation
'                txtEditQty.SetFocus
'                Exit Sub
'            End If
            LV.SelectedItem.SubItems(6) = Val(txtEditQty)
            LV.SelectedItem.SubItems(7) = Val(txtEditQty) * Val(LV.SelectedItem.SubItems(5))
            txtEditQty.Visible = False
        End If
    End If
    
End Sub

Private Sub txtPrice_Change()
   lblAmt = Format(Val(txtQty) * Val(txtPrice), "#0.00")
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtQty.SetFocus
    End If
   
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
    If KeyAscii = 13 Then
        Call cmdAddItem_Click
    End If
   KeyAscii = OnlyNumber(KeyAscii)
End Sub


Private Sub GetTotal()

    Dim i As Integer
    Dim total As Double
    total = 0

    For i = 1 To LV.ListItems.count
        total = total + Val(LV.ListItems(i).ListSubItems(7).Text)
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
        MsgBox "Please Enter Order No.", vbInformation
        txtOrderNo.SetFocus
        Exit Function
    End If
    
    If LV.ListItems.count = 0 Then
        MsgBox "No Items Added to This Order", vbInformation
        Exit Function
    End If
    
    
    Dim lProformaRefID As Long
    Call StartTrans(con)
    '----------------------------------------------------------------
    'If strOrderNo <> "" Then
    '    con.Execute "Delete from FCustomerOrders Where OrderNo='" & _
         strOrderNo & "'"
    '     'Now Deleting is giving a problem...
    'End If
    
    'Save Order
    If Not cmbCust.Text = "HICO" Then
        If strOrderNo <> "" Then
            con.Execute "UPDATE FCustomerOrders SET OrderNo='" & txtOrderNo & "',DT='" & _
             OrderDT & "',TradeTerms='" & cmbTradeTerms & "',PartialShipment='" & _
             cmbPartial & "',PaymentTerms='" & cmbPayment & "',TransShipment='" & _
             cmbTrans & "',Packaging='" & txtPackaging & "',DeliveryDT='" & _
             DeliveryDT & "',CompanyRefID=" & cmbCompany.ID & ",StampDT='" & txtStamps & _
             "',Quality='" & txtQuality & "',InternalRefNo=" & Val(txtInternalRefNo) & _
             ",OrderRcvdVia='" & txtOrderRcvdvia & "',RevisionHistoryBox='" & txtRevisionHistory & "' WHERE OrderNo='" & strOrderNo & "'"
             
             If strDelList <> "" Then
                strDelList = Left(strDelList, Len(strDelList) - 1)
             
                con.Execute "DELETE FROM FOrderItems Where ID In(" & strDelList & ")"
                
                If bAutoProforma Then
                    con.Execute "DELETE FROM FProformaOrders WHERE OrderEntryID In(" & strDelList & ")"
                End If
                
            End If
            
            If bAutoProforma Then
                con.Execute "UPDATE FPInvoice SET PInvoice='" & txtOrderNo & "' WHERE PInvoice='" & strOrderNo & "'"
            End If
        Else
        
            con.Execute "INSERT INTO FCustomerOrders (CustCode,Country," & _
             "OrderNo,DT,TradeTerms,PartialShipment,PaymentTerms," & _
             "TransShipment,Packaging,DeliveryDT,CompanyRefID,StampDT,Quality,InternalRefNo,OrderRcvdVia,RevisionHistoryBox) VALUES('" & _
             cmbCust.ID & "','" & cmbCountry & "','" & txtOrderNo & _
             "','" & OrderDT & "','" & cmbTradeTerms & "','" & cmbPartial & _
             "','" & cmbPayment & "','" & cmbTrans & "','" & txtPackaging & _
             "','" & DeliveryDT & "'," & cmbCompany.ID & ",'" & txtStamps & "','" & _
             txtQuality & "'," & Val(txtInternalRefNo) & ",'" & txtOrderRcvdvia & "','" & txtRevisionHistory & "')"
             
            If bAutoProforma Then
                con.Execute "INSERT INTO FPInvoice(CustCode,Country,PInvoice,DT,SMethod,TradeTerms,PaymentTerms,CompanyRefID,DeliveryDT) " & _
                "VALUES('" & cmbCust.ID & "','" & cmbCountry & "','" & txtOrderNo & "','','" & OrderDT & "','" & cmbTradeTerms & _
                "','" & cmbPayment & "'," & cmbCompany.ID & ",'" & DeliveryDT & "')"
            End If
            
        End If
    End If
     
    If IsOrderQtyChanges(strOrderNo) Then
        'do Something.
        con.Execute "INSERT INTO FCustomerOrderRevisions(OrderNo,UserName,MachineName) VALUES('" & txtOrderNo & "','" & CurrentUserName & "','" & strComputerName & "')"
        Dim lRevEntryID As Long
        lRevEntryID = GetSingleLongValue("MAX(EntryID)", "FCustomerOrderRevisions", " WHERE OrderNo='" & txtOrderNo & "' AND MachineName='" & strComputerName & "'")
        con.Execute "INSERT INTO FCustomerOrderRevisionsDetail(FCOR_RefID,CompItemCode,Price,Qty) SELECT " & lRevEntryID & ",CompItemCode,Price,Qty FROM FOrderItems WHERE OrderNo='" & txtOrderNo & "'"
    End If
    
    'Now Save Items
    Dim strItemCode As String
    For i = 1 To LV.ListItems.count
        strItemCode = Left(LV.ListItems(i).key, Len(LV.ListItems(i).key) - 1)
        
SaveNewItem:
            If Val(LV.ListItems(i).ListSubItems(7).Tag) = 0 Then
                
                con.Execute "INSERT INTO FOrderItems (OrderNo,ItemCode,CompItemCode," & _
                 "Price,Qty,InvQty,CustomPrice,SortNo) VALUES('" & txtOrderNo & "','" & _
                 Left(LV.ListItems(i).key, Len(LV.ListItems(i).key) - 1) & _
                 "','" & LV.ListItems(i).Tag & "'," & _
                 Val(LV.ListItems(i).SubItems(5)) & _
                 "," & Val(LV.ListItems(i).SubItems(6)) & "," & _
                 Val(LV.ListItems(i).SubItems(6)) & "," & Val(LV.ListItems(i).ListSubItems(5).Tag) & "," & Val(LV.ListItems(i).SubItems(8)) & ")"
                 
                If bAutoProforma Then
                
                    
                    lProformaRefID = GetSingleLongValue("MAX(ID)", "FOrderItems")
                    
                    con.Execute "INSERT INTO FProformaOrders(PInvoice,OrderEntryID,Qty,InvQty,Price,CustomPrice) VALUES('" & _
                     txtOrderNo & "'," & lProformaRefID & "," & Val(LV.ListItems(i).SubItems(6)) & "," & _
                     Val(LV.ListItems(i).SubItems(6)) & "," & Val(LV.ListItems(i).SubItems(5)) & "," & _
                     Val(LV.ListItems(i).SubItems(5)) & ")"
                     
                    con.Execute "UPDATE FOrderItems SET InvQty=0 WHERE ID=" & lProformaRefID
                End If
                 'INSERT INTO FCustomerCatalog for All Seas, If Not Already ...
                 If strCompany = "All Seas" Then
                    If GetSingleLongValue("COUNT(ItemID)", "FCustomerCatalog", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "' AND ItemID='" & strItemCode & "'") = 0 Then
                        con.Execute "INSERT INTO FCustomerCatalog(CustCode,Country,ItemID,CompItemID) VALUES('" & _
                         cmbCust.Text & "','" & cmbCountry.Text & "','" & strItemCode & "','" & strItemCode & "')"
                    End If
                 End If
            Else
                con.Execute "UPDATE FOrderItems SET Qty=" & Val(LV.ListItems(i).SubItems(6)) & _
                 ",InvQty=" & Val(LV.ListItems(i).SubItems(6)) - Val(LV.ListItems(i).ListSubItems(6).Tag) & _
                 ",Price=" & Val(LV.ListItems(i).SubItems(5)) & ",CustomPrice=" & Val(LV.ListItems(i).SubItems(5)) & _
                 ",SortNo=" & Val(LV.ListItems(i).SubItems(8)) & " WHERE ID=" & Val(LV.ListItems(i).ListSubItems(7).Tag)
                
                If bAutoProforma Then
                    con.Execute "UPDATE FProformaOrders SET Qty=" & Val(LV.ListItems(i).SubItems(6)) & _
                    ",InvQty=InvQty-(Qty-" & Val(LV.ListItems(i).SubItems(6)) & ") WHERE OrderEntryID=" & _
                    Val(LV.ListItems(i).ListSubItems(7).Tag)
                End If
                
            End If
    Next
      
    con.CommitTrans
    
    Saved = True
    If chkPrint.Value = True Then
        Call PrintOrder(txtOrderNo)
    End If
    
    Exit Function

err:
    MsgBox err.Description
    
End Function

Private Function IsOrderQtyChanges(p_strOrderNo As String) As Boolean

    Dim rs As New ADODB.Recordset
    Dim strItemCode As String
    Dim myCollection As New Collection
    Dim myObj As clsOrderDetails
    Dim bOrderUpdated As Boolean
    bOrderUpdated = False
    With rs
        .Open "SELECT * FROM FOrderItems WHERE OrderNo='" & p_strOrderNo & "'", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            Set myObj = New clsOrderDetails
            myObj.strItemCode = !CompItemCode & ""
            myObj.lQty = Val(!Qty & "")
            myObj.dPrice = Val(!Price & "")
            
            myCollection.add myObj, !CompItemCode & ""
            Set myObj = Nothing
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    'Now Compare this Collection with Now Present Items & Their Qty.
    Dim i As Integer, j As Integer, bFound As Boolean
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            'Find it in the Collection & Compare Qty...
            bFound = False
            For j = 1 To myCollection.count
                If .SubItems(2) = myCollection.Item(j).strItemCode Then
                    bFound = True
                    Set myObj = myCollection.Item(.SubItems(2))
                    Exit For
                End If
            Next
            If bFound = False Then  'Not Present in the Collection, New Item...
                'New Item(s) Added....
                bOrderUpdated = True
                Exit For
            End If
            If myObj.lQty <> Val(.SubItems(6)) Then
                'Qty Changed....
                bOrderUpdated = True
                Exit For
            End If
        End With
    Next
    
    If bOrderUpdated Then
        IsOrderQtyChanges = True
    Else
        IsOrderQtyChanges = False
    End If
    
End Function

Private Sub PrintOrder(strOrderNo As String)

    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptFCustOrders.rpt") '
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{CustOrders.OrderNo}='" & strOrderNo & "' ", rpt
    
    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub FillCmbs()
    
    Call AddToCombo(cmbTradeTerms, "Description", "TradeTerms", " Order By ID")
'    With cmbTradeTerms
'        .Clear
'        .AddItem "FOB"
'        .AddItem "CIF (Air)"
'        .AddItem "CIF (Sea)"
'        .AddItem "C&F (Air)"
'        .AddItem "C&F (Sea)"
'        .ListIndex = 0
'    End With
    
    With cmbPartial
        .AddItem "Allowed"
        .AddItem "Not Allowed"
        .ListIndex = 0
    End With
    
    With cmbTrans
        .AddItem "Allowed"
        .AddItem "Not Allowed"
        .ListIndex = 0
    End With
    
    Call AddToCombo(cmbPayment, "Description", "PaymentTerms")
    
    cmbCompany.ListHeight = 6000
    cmbCompany.AddVals con, "CompanyName", "Companies", "EntryID", " Order By CompanyName"
    cmbCompany.ID = 1
    
    With cmbQuality
        .AddItem "Regular"
        .AddItem "Top"
        .ListIndex = 0
    End With
    
End Sub

Private Sub FillItems()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    
    With rs
        
        If 0 = 0 Then
            If strCompany = "All Seas" Then
                .Open "SELECT ItemID,ItemID AS CompItemID,ItemName AS Description,0 AS PackingMode,1 AS FOB,1 AS ExWorks" & _
                 ",1 AS CIFAir,1 AS CIFSea,1 AS CnFAir,1 AS CnFSea,1 AS CnIAir,1 AS CnISea,1 AS CustomFOB,1 AS CustomExWorks,1 AS CustomCIFAir,1 AS CustomCIFSea,1 AS CustomCnFAir,1 AS CustomCnFSea,1 AS CustomCnIAir,1 AS CustomCnISea FROM Items", con, adOpenForwardOnly, adLockReadOnly
            Else
                .Open "SELECT * FROM VFCustomerCatalog WHERE CustCode+Country='" & cmbCust.Text & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
            End If
        Else
            If strCompany = "All Seas" Then
                .Open "SELECT ItemID,ItemID AS CompItemID,ItemName AS Description,Unit AS PackingMode,1 AS FOB,1 AS ExWorks" & _
                 ",1 AS CIFAir,1 AS CIFSea,1 AS CnFAir,1 AS CnFSea,1 AS CnIAir,1 AS CnISea,1 AS CustomFOB,1 AS CustomExWorks,1 AS CustomCIFAir,1 AS CustomCIFSea,1 AS CustomCnFAir,1 AS CustomCnFSea,1 AS CustomCnIAir,1 AS CustomCnISea FROM Items WHERE CatID=" & cmbCatID.List(cmbCatID.ListIndex, 1), con, adOpenForwardOnly, adLockReadOnly
            Else
                .Open "SELECT * FROM VFCustomerCatalog WHERE CatID=" & cmbCatID.List(cmbCatID.ListIndex, 1) & " AND CustCode+Country='" & cmbCust.Text & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
            End If
        End If
        cmbItems.Clear
        cmbDesc.Clear
        cmbCustItemID.Clear
        Do Until .EOF
            cmbCustItemID.AddItem !ITemID & ""
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
            
            cmbDesc.List(cmbDesc.ListCount - 1, 2) = Val(![FOBTop] & "")
            cmbDesc.List(cmbDesc.ListCount - 1, 3) = Val(![ExWorksTop] & "")
            cmbDesc.List(cmbDesc.ListCount - 1, 4) = Val(![CIFAirTop] & "")
            cmbDesc.List(cmbDesc.ListCount - 1, 5) = Val(![CIFSeaTop] & "")
            
            cmbDesc.List(cmbDesc.ListCount - 1, 6) = Val(![CnFAirTop] & "")
            cmbDesc.List(cmbDesc.ListCount - 1, 7) = Val(![CnFSeaTop] & "")
            
            cmbDesc.List(cmbDesc.ListCount - 1, 8) = Val(![CnIAirTop] & "")
            cmbDesc.List(cmbDesc.ListCount - 1, 9) = Val(![CnISeaTop] & "")
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Public Sub CopyCustomerOrder(p_strOrderNo As String)
    bCopy = True
    strCopyFromOrderNo = p_strOrderNo
    Call EditOrder(strCopyFromOrderNo)
End Sub
