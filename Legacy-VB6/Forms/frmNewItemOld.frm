VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "combolist.ocx"
Begin VB.Form frmNewItemOld 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DMR (Device Master Record)"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9690
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   525
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   646
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2640
      Top             =   7170
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Please Select Employee's Picture"
      Filter          =   "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2100
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewItemOld.frx":0000
            Key             =   "0'"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewItemOld.frx":0892
            Key             =   "7'"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewItemOld.frx":0CE4
            Key             =   "5'"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewItemOld.frx":1136
            Key             =   "6'"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewItemOld.frx":3470
            Key             =   "4'"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewItemOld.frx":414A
            Key             =   "2'"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewItemOld.frx":49DC
            Key             =   "3'"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewItemOld.frx":718E
            Key             =   "1'"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Device Info."
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
      Height          =   7065
      Left            =   90
      TabIndex        =   3
      Top             =   15
      Width           =   9480
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   120
         TabIndex        =   73
         Top             =   270
         Width           =   9120
         Begin MSForms.ComboBox cmbItemType 
            Height          =   315
            Left            =   6990
            TabIndex        =   174
            Tag             =   "ItemType"
            Top             =   210
            Width           =   2040
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "3598;556"
            MatchEntry      =   1
            ListStyle       =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Type :"
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
            Index           =   65
            Left            =   6420
            TabIndex        =   173
            Top             =   240
            Width           =   510
         End
         Begin MSForms.ComboBox cmbItemGroup 
            Height          =   315
            Left            =   975
            TabIndex        =   83
            Tag             =   "GroupID"
            Top             =   570
            Width           =   5385
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "9499;556"
            MatchEntry      =   1
            ListStyle       =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Group :"
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
            Index           =   26
            Left            =   345
            TabIndex        =   82
            Top             =   615
            Width           =   600
         End
         Begin MSForms.ComboBox cmbCatID 
            Height          =   315
            Left            =   975
            TabIndex        =   77
            Tag             =   "CatID"
            Top             =   210
            Width           =   5385
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "9499;556"
            MatchEntry      =   1
            ListStyle       =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Category :"
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
            Index           =   16
            Left            =   75
            TabIndex        =   76
            Top             =   270
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Code :"
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
            Index           =   4
            Left            =   6435
            TabIndex        =   75
            Top             =   615
            Width           =   510
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   3
            Left            =   6990
            TabIndex        =   74
            Tag             =   "ItemID"
            Top             =   570
            Width           =   2040
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "3598;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5595
         Left            =   120
         TabIndex        =   4
         Top             =   1350
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   9869
         _Version        =   393216
         Style           =   1
         Tabs            =   10
         Tab             =   1
         TabsPerRow      =   10
         TabHeight       =   609
         ShowFocusRect   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "General Info."
         TabPicture(0)   =   "frmNewItemOld.frx":7A20
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame4"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Processes"
         TabPicture(1)   =   "frmNewItemOld.frx":7A3C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame11"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Components"
         TabPicture(2)   =   "frmNewItemOld.frx":7A58
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmdComp"
         Tab(2).Control(1)=   "Label1(31)"
         Tab(2).Control(2)=   "NBox(19)"
         Tab(2).Control(3)=   "cmdAdd"
         Tab(2).Control(4)=   "cmbComp"
         Tab(2).Control(5)=   "Label1(29)"
         Tab(2).Control(6)=   "LV1"
         Tab(2).Control(7)=   "cmdDelete"
         Tab(2).Control(8)=   "cmbCompProc"
         Tab(2).ControlCount=   9
         TabCaption(3)   =   "Weight"
         TabPicture(3)   =   "frmNewItemOld.frx":7A74
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label1(47)"
         Tab(3).Control(1)=   "TBox(11)"
         Tab(3).Control(2)=   "Label1(48)"
         Tab(3).Control(3)=   "TBox(12)"
         Tab(3).Control(4)=   "Frame3"
         Tab(3).Control(5)=   "Frame6"
         Tab(3).Control(6)=   "Frame7"
         Tab(3).ControlCount=   7
         TabCaption(4)   =   "Account Info."
         TabPicture(4)   =   "frmNewItemOld.frx":7A90
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Label1(13)"
         Tab(4).Control(1)=   "Label1(12)"
         Tab(4).Control(2)=   "Label1(14)"
         Tab(4).Control(3)=   "cmbAssetAcc"
         Tab(4).Control(4)=   "cmbSaleRAcc"
         Tab(4).Control(5)=   "cmbSaleAcc"
         Tab(4).ControlCount=   6
         TabCaption(5)   =   "Price/Costing"
         TabPicture(5)   =   "frmNewItemOld.frx":7AAC
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "chkStamper"
         Tab(5).Control(1)=   "chkSnaffer"
         Tab(5).Control(2)=   "chkPlanter"
         Tab(5).Control(3)=   "NBox(1)"
         Tab(5).Control(4)=   "NBox(2)"
         Tab(5).Control(5)=   "NBox(3)"
         Tab(5).Control(6)=   "Frame5"
         Tab(5).Control(7)=   "Frame8"
         Tab(5).Control(8)=   "Frame10"
         Tab(5).ControlCount=   9
         TabCaption(6)   =   "F. Catalog Ref."
         TabPicture(6)   =   "frmNewItemOld.frx":7AC8
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "LV3"
         Tab(6).Control(1)=   "txtRefID"
         Tab(6).ControlCount=   2
         TabCaption(7)   =   "Picture"
         TabPicture(7)   =   "frmNewItemOld.frx":7AE4
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "lblPic"
         Tab(7).Control(1)=   "Label1(32)"
         Tab(7).Control(2)=   "Image1"
         Tab(7).Control(3)=   "lblTechDrawing"
         Tab(7).Control(4)=   "Label1(60)"
         Tab(7).Control(5)=   "ImgTechDrawing"
         Tab(7).Control(6)=   "cmdViewTechDrawing"
         Tab(7).Control(7)=   "cmdViewPicture"
         Tab(7).ControlCount=   8
         TabCaption(8)   =   "RM Components"
         TabPicture(8)   =   "frmNewItemOld.frx":7B00
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "cmbRawMaterial"
         Tab(8).Control(1)=   "cmdDeleteRM"
         Tab(8).Control(2)=   "cmbProcRM"
         Tab(8).Control(3)=   "LVRM"
         Tab(8).Control(4)=   "Label1(34)"
         Tab(8).Control(5)=   "Label1(33)"
         Tab(8).Control(6)=   "NBox(13)"
         Tab(8).Control(7)=   "cmdAddRM"
         Tab(8).Control(8)=   "cmbRM"
         Tab(8).Control(9)=   "Label1(30)"
         Tab(8).ControlCount=   10
         TabCaption(9)   =   "Ship. Info."
         TabPicture(9)   =   "frmNewItemOld.frx":7B1C
         Tab(9).ControlEnabled=   0   'False
         Tab(9).Control(0)=   "Frame9"
         Tab(9).ControlCount=   1
         Begin VB.Frame Frame11 
            Height          =   5115
            Left            =   60
            TabIndex        =   175
            Top             =   360
            Width           =   8955
            Begin VB.TextBox txtProcRate 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3060
               TabIndex        =   177
               Top             =   1695
               Visible         =   0   'False
               Width           =   2730
            End
            Begin VB.TextBox txtSortNo 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3135
               TabIndex        =   176
               Top             =   1185
               Visible         =   0   'False
               Width           =   1515
            End
            Begin MSComctlLib.ListView LV2 
               Height          =   4830
               Left            =   45
               TabIndex        =   178
               Top             =   165
               Width           =   8835
               _ExtentX        =   15584
               _ExtentY        =   8520
               View            =   3
               Arrange         =   2
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               AllowReorder    =   -1  'True
               Checkboxes      =   -1  'True
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
                  Text            =   "Sort No."
                  Object.Width           =   2822
               EndProperty
            End
         End
         Begin VB.Frame Frame10 
            Height          =   780
            Left            =   -73140
            TabIndex        =   164
            Top             =   3930
            Width           =   5040
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Re-Order Level"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   64
               Left            =   3735
               TabIndex        =   172
               Top             =   165
               Width           =   1095
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   36
               Left            =   3735
               TabIndex        =   168
               Tag             =   "ReOrderLevel"
               Top             =   375
               Width           =   1185
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2090;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Max. Level"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   63
               Left            =   2535
               TabIndex        =   171
               Top             =   165
               Width           =   780
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   35
               Left            =   2520
               TabIndex        =   167
               Tag             =   "MaxLevel"
               Top             =   375
               Width           =   1185
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2090;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Min. Level"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   62
               Left            =   1335
               TabIndex        =   170
               Top             =   165
               Width           =   720
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   34
               Left            =   1305
               TabIndex        =   166
               Tag             =   "MinLevel"
               Top             =   375
               Width           =   1185
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2090;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Price For Cost"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   50
               Left            =   120
               TabIndex        =   169
               Top             =   165
               Width           =   1005
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   25
               Left            =   90
               TabIndex        =   165
               Tag             =   "PriceForCost"
               Top             =   375
               Width           =   1185
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2090;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
         End
         Begin VB.CommandButton cmdViewPicture 
            Caption         =   "View Actual Size"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   -67440
            TabIndex        =   160
            Top             =   4800
            Width           =   1200
         End
         Begin VB.CommandButton cmdViewTechDrawing 
            Caption         =   "View Actual Size"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   -74865
            TabIndex        =   159
            Top             =   4785
            Width           =   1200
         End
         Begin VB.Frame Frame9 
            Caption         =   "Shipment Hisotry :"
            Height          =   4605
            Left            =   -74820
            TabIndex        =   153
            Top             =   840
            Width           =   8685
            Begin VB.TextBox txtTotalShipValue 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   5310
               Locked          =   -1  'True
               TabIndex        =   156
               Top             =   4125
               Width           =   1230
            End
            Begin VB.TextBox txtTotalShipQty 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2925
               Locked          =   -1  'True
               TabIndex        =   155
               Top             =   4125
               Width           =   1230
            End
            Begin MSComctlLib.ListView LVShipInfo 
               Height          =   3765
               Left            =   120
               TabIndex        =   154
               Top             =   255
               Width           =   8430
               _ExtentX        =   14870
               _ExtentY        =   6641
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
               NumItems        =   5
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Sr #"
                  Object.Width           =   1411
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Invoice No."
                  Object.Width           =   3175
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   2
                  Text            =   "Qty"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   3
                  Text            =   "Price"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   4
                  Text            =   "Amount"
                  Object.Width           =   2117
               EndProperty
            End
         End
         Begin ComboList.Usercontrol1 cmbRawMaterial 
            Height          =   285
            Left            =   -71370
            TabIndex        =   152
            Top             =   1335
            Width           =   2775
            _ExtentX        =   4895
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
         Begin VB.Frame Frame8 
            Caption         =   "Conversion :"
            Height          =   1530
            Left            =   -74085
            TabIndex        =   132
            Top             =   2355
            Width           =   7020
            Begin MSForms.TextBox txtConversion 
               Height          =   315
               Left            =   90
               TabIndex        =   150
               Top             =   450
               Width           =   1335
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2355;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Euro to US$"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   59
               Left            =   150
               TabIndex        =   149
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "C & F (Air)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   58
               Left            =   4200
               TabIndex        =   148
               Top             =   825
               UseMnemonic     =   0   'False
               Width           =   750
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   33
               Left            =   4185
               TabIndex        =   147
               Top             =   1050
               Width           =   1335
               VariousPropertyBits=   679495711
               BorderStyle     =   1
               Size            =   "2355;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "C & F (Sea)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   57
               Left            =   2850
               TabIndex        =   146
               Top             =   825
               UseMnemonic     =   0   'False
               Width           =   825
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   32
               Left            =   2820
               TabIndex        =   145
               Top             =   1050
               Width           =   1335
               VariousPropertyBits=   679495711
               BorderStyle     =   1
               Size            =   "2355;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "CIF (Air)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   56
               Left            =   1500
               TabIndex        =   144
               Top             =   825
               Width           =   615
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   31
               Left            =   1455
               TabIndex        =   143
               Top             =   1050
               Width           =   1335
               VariousPropertyBits=   679495711
               BorderStyle     =   1
               Size            =   "2355;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "CIF (Sea)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   55
               Left            =   150
               TabIndex        =   142
               Top             =   825
               Width           =   690
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   30
               Left            =   90
               TabIndex        =   141
               Top             =   1050
               Width           =   1335
               VariousPropertyBits=   679495711
               BorderStyle     =   1
               Size            =   "2355;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   29
               Left            =   1470
               TabIndex        =   140
               Top             =   450
               Width           =   1335
               VariousPropertyBits=   679495711
               BorderStyle     =   1
               Size            =   "2355;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "FOB Karachi"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   54
               Left            =   1545
               TabIndex        =   139
               Top             =   225
               Width           =   870
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "FOB Sialkot"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   53
               Left            =   2910
               TabIndex        =   138
               Top             =   225
               Width           =   810
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   28
               Left            =   2835
               TabIndex        =   137
               Top             =   450
               Width           =   1335
               VariousPropertyBits=   679495711
               BorderStyle     =   1
               Size            =   "2355;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   27
               Left            =   4200
               TabIndex        =   136
               Top             =   450
               Width           =   1335
               VariousPropertyBits=   679495711
               BorderStyle     =   1
               Size            =   "2355;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "C && I (Sea)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   52
               Left            =   4260
               TabIndex        =   135
               Top             =   225
               Width           =   795
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   26
               Left            =   5565
               TabIndex        =   134
               Top             =   450
               Width           =   1335
               VariousPropertyBits=   679495711
               BorderStyle     =   1
               Size            =   "2355;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "C && I (Air)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   51
               Left            =   5610
               TabIndex        =   133
               Top             =   225
               Width           =   720
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Small Box :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   945
            Left            =   -74310
            TabIndex        =   119
            Top             =   3975
            Width           =   5325
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Length (cm)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   46
               Left            =   105
               TabIndex        =   125
               Top             =   255
               Width           =   855
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   23
               Left            =   75
               TabIndex        =   124
               Tag             =   "SmallCartonL"
               Top             =   480
               Width           =   1740
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "3069;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   22
               Left            =   1830
               TabIndex        =   123
               Tag             =   "SmallCartonH"
               Top             =   480
               Width           =   1665
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2937;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Height (cm)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   44
               Left            =   1860
               TabIndex        =   122
               Top             =   255
               Width           =   825
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   20
               Left            =   3525
               TabIndex        =   121
               Tag             =   "SmallCartonW"
               Top             =   480
               Width           =   1665
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2937;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Width (cm)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   41
               Left            =   3570
               TabIndex        =   120
               Top             =   255
               Width           =   780
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Master Carton :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   945
            Left            =   -74310
            TabIndex        =   112
            Top             =   3015
            Width           =   5325
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Width (cm)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   45
               Left            =   3570
               TabIndex        =   118
               Top             =   255
               Width           =   780
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   21
               Left            =   3525
               TabIndex        =   117
               Tag             =   "MasterCartonW"
               Top             =   480
               Width           =   1665
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2937;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Height (cm)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   43
               Left            =   1860
               TabIndex        =   116
               Top             =   255
               Width           =   825
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   18
               Left            =   1830
               TabIndex        =   115
               Tag             =   "MasterCartonH"
               Top             =   480
               Width           =   1665
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2937;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   17
               Left            =   75
               TabIndex        =   114
               Tag             =   "MasterCartonL"
               Top             =   480
               Width           =   1740
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "3069;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Length (cm)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   42
               Left            =   105
               TabIndex        =   113
               Top             =   255
               Width           =   855
            End
         End
         Begin VB.CommandButton cmdDeleteRM 
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
            Left            =   -67260
            Picture         =   "frmNewItemOld.frx":7B38
            Style           =   1  'Graphical
            TabIndex        =   94
            ToolTipText     =   "Remove This Entry From Voucher"
            Top             =   1680
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   540
         End
         Begin ComboList.Usercontrol1 cmbCompProc 
            Height          =   285
            Left            =   -73875
            TabIndex        =   92
            Top             =   1575
            Width           =   2730
            _ExtentX        =   4815
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
            Left            =   -67005
            Picture         =   "frmNewItemOld.frx":806A
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Remove This Entry From Voucher"
            Top             =   1920
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox txtRefID 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -73290
            TabIndex        =   56
            Top             =   1500
            Visible         =   0   'False
            Width           =   2730
         End
         Begin VB.Frame Frame3 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2160
            Left            =   -74310
            TabIndex        =   32
            Top             =   795
            Width           =   3630
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Fixed Packing Unit"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   49
               Left            =   1935
               TabIndex        =   131
               Top             =   1455
               Width           =   1305
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   24
               Left            =   1890
               TabIndex        =   130
               Tag             =   "FixedPackingUnit"
               Top             =   1665
               Width           =   1635
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2884;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin MSForms.ComboBox cmbWeightUnit 
               Height          =   315
               Left            =   75
               TabIndex        =   55
               Tag             =   "UnitWeight"
               Top             =   435
               Width           =   1740
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "3069;556"
               ListWidth       =   8819
               MatchEntry      =   1
               ListStyle       =   1
               ShowDropButtonWhen=   1
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Weighing Unit"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   18
               Left            =   105
               TabIndex        =   54
               Top             =   225
               Width           =   990
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Forging Weight"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   10
               Left            =   1935
               TabIndex        =   50
               Top             =   225
               Width           =   1095
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   4
               Left            =   1890
               TabIndex        =   48
               Tag             =   "ForgingWeight"
               Top             =   435
               Width           =   1635
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2884;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   5
               Left            =   75
               TabIndex        =   46
               Tag             =   "WasteVisible"
               Top             =   1020
               Width           =   1740
               VariousPropertyBits=   679495711
               BorderStyle     =   1
               Size            =   "3069;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Wastage (Visible)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   9
               Left            =   90
               TabIndex        =   44
               Top             =   780
               Width           =   1245
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   6
               Left            =   1890
               TabIndex        =   41
               Top             =   1020
               Width           =   1635
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2884;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Wastage (Invisible)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   8
               Left            =   1935
               TabIndex        =   39
               Top             =   810
               Width           =   1395
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   7
               Left            =   75
               TabIndex        =   37
               Tag             =   "FinishedWeight"
               Top             =   1665
               Width           =   1740
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "3069;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Finished Weight"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   7
               Left            =   120
               TabIndex        =   35
               Top             =   1455
               Width           =   1140
            End
         End
         Begin VB.Frame Frame4 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4440
            Left            =   -73845
            TabIndex        =   18
            Top             =   660
            Width           =   6765
            Begin VB.TextBox txtUrduFields 
               Alignment       =   1  'Right Justify
               DataSource      =   "datPrimaryRS"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   0
               Left            =   105
               TabIndex        =   53
               Top             =   3900
               Width           =   6585
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Type"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   11
               Left            =   4785
               TabIndex        =   163
               Top             =   1875
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Color"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   40
               Left            =   3150
               TabIndex        =   162
               Top             =   1875
               Width           =   375
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Tip Size"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   61
               Left            =   3630
               TabIndex        =   161
               Top             =   1335
               Width           =   540
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   13
               Left            =   3600
               TabIndex        =   34
               Tag             =   "TipSize"
               Top             =   1545
               Width           =   1110
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "1958;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Urdu Name :"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   19
               Left            =   150
               TabIndex        =   151
               Top             =   3675
               Width           =   900
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   10
               Left            =   3120
               TabIndex        =   43
               Tag             =   "ItemColor"
               Top             =   2085
               Width           =   1590
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2805;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   9
               Left            =   105
               TabIndex        =   31
               Tag             =   "CustomDescription"
               Top             =   960
               Width           =   6585
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "11615;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Custom Description"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   36
               Left            =   135
               TabIndex        =   103
               Top             =   750
               Width           =   1380
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   8
               Left            =   1875
               TabIndex        =   42
               Tag             =   "EAN128"
               Top             =   2085
               Width           =   1200
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2117;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Catalog Page #"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   35
               Left            =   1890
               TabIndex        =   102
               Top             =   1875
               Width           =   1125
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Finishing Quality"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   27
               Left            =   5055
               TabIndex        =   84
               Top             =   3000
               Width           =   1170
            End
            Begin MSForms.ComboBox cmbFinQuality 
               Height          =   315
               Left            =   5025
               TabIndex        =   52
               Tag             =   "FinQuality"
               Top             =   3225
               Width           =   1650
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "2910;556"
               ListWidth       =   3527
               MatchEntry      =   1
               ListStyle       =   1
               ShowDropButtonWhen=   1
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Actual Description"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   1
               Left            =   135
               TabIndex        =   29
               Top             =   180
               Width           =   1290
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   0
               Left            =   105
               TabIndex        =   28
               Tag             =   "ItemName"
               Top             =   390
               Width           =   5610
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "9895;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbType 
               Height          =   315
               Left            =   4755
               TabIndex        =   45
               Tag             =   "Type"
               Top             =   2085
               Width           =   1920
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "3387;556"
               ListWidth       =   3527
               MatchEntry      =   1
               ListStyle       =   1
               ShowDropButtonWhen=   1
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   7
               Left            =   4755
               TabIndex        =   36
               Tag             =   "ItemSize"
               Top             =   1545
               Width           =   960
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "1693;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Size"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   15
               Left            =   4785
               TabIndex        =   27
               Top             =   1335
               Width           =   285
            End
            Begin MSForms.ComboBox cmbUnit 
               Height          =   315
               Left            =   5745
               TabIndex        =   30
               Tag             =   "Unit"
               Top             =   390
               Width           =   930
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "1640;556"
               ListWidth       =   3527
               MatchEntry      =   1
               ListStyle       =   1
               ShowDropButtonWhen=   1
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   5
               Left            =   105
               TabIndex        =   33
               Tag             =   "ItemUsage"
               Top             =   1545
               Width           =   3465
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "6112;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Special Instructions"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   5
               Left            =   135
               TabIndex        =   26
               Top             =   3015
               Width           =   1395
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   2
               Left            =   105
               TabIndex        =   51
               Tag             =   "PacknLabel"
               Top             =   3225
               Width           =   4875
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "8599;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Steel Used"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   3
               Left            =   135
               TabIndex        =   25
               Top             =   2445
               Width           =   765
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Internal Catalog Ref. #"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   2
               Left            =   135
               TabIndex        =   24
               Top             =   1875
               Width           =   1695
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   1
               Left            =   120
               TabIndex        =   40
               Tag             =   "Attributes"
               Top             =   2085
               Width           =   1725
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "3043;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Item Usage"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   0
               Left            =   135
               TabIndex        =   23
               Top             =   1335
               Width           =   825
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Unit"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   6
               Left            =   5790
               TabIndex        =   22
               Top             =   180
               Width           =   285
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Unit"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   24
               Left            =   5790
               TabIndex        =   21
               Top             =   1335
               Width           =   285
            End
            Begin MSForms.ComboBox cmbSizeUnit 
               Height          =   315
               Left            =   5745
               TabIndex        =   38
               Tag             =   "SizeUnit"
               Top             =   1545
               Width           =   930
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "1640;556"
               ListWidth       =   3527
               MatchEntry      =   1
               ListStyle       =   1
               ShowDropButtonWhen=   1
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbSteel 
               Height          =   315
               Left            =   105
               TabIndex        =   47
               Tag             =   "SteelUsed"
               Top             =   2640
               Width           =   3270
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "5768;556"
               ListWidth       =   8819
               MatchEntry      =   1
               ListStyle       =   1
               ShowDropButtonWhen=   1
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Gauge"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   25
               Left            =   3525
               TabIndex        =   20
               Top             =   2445
               Width           =   465
            End
            Begin MSForms.ComboBox cmbGuage 
               Height          =   315
               Left            =   3405
               TabIndex        =   49
               Tag             =   "Gage"
               Top             =   2640
               Width           =   3270
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "5768;556"
               ListWidth       =   8819
               MatchEntry      =   1
               ListStyle       =   1
               ShowDropButtonWhen=   1
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   6
               Left            =   2850
               TabIndex        =   19
               Tag             =   "CreateDT"
               Top             =   3360
               Visible         =   0   'False
               Width           =   1305
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2302;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Pricing :"
            Height          =   1530
            Left            =   -74055
            TabIndex        =   5
            Top             =   840
            Width           =   7020
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   12
               Left            =   5550
               TabIndex        =   111
               Tag             =   "ReorderPoint"
               Top             =   1050
               Width           =   1335
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2355;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Reorder Point"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   28
               Left            =   5595
               TabIndex        =   110
               Top             =   855
               Width           =   990
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "C && I (Air)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   39
               Left            =   5610
               TabIndex        =   109
               Top             =   225
               Width           =   720
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   16
               Left            =   5565
               TabIndex        =   108
               Tag             =   "CnIAir"
               Top             =   450
               Width           =   1335
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2355;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "C && I (Sea)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   38
               Left            =   4260
               TabIndex        =   107
               Top             =   225
               Width           =   795
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   15
               Left            =   4200
               TabIndex        =   106
               Tag             =   "CnISea"
               Top             =   450
               Width           =   1335
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2355;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   14
               Left            =   2835
               TabIndex        =   105
               Tag             =   "ExWorks"
               Top             =   450
               Width           =   1335
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2355;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "FOB Sialkot"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   37
               Left            =   2910
               TabIndex        =   104
               Top             =   225
               Width           =   810
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "FOB Karachi"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   17
               Left            =   1545
               TabIndex        =   17
               Top             =   225
               Width           =   870
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   0
               Left            =   1470
               TabIndex        =   16
               Tag             =   "FOB"
               Top             =   450
               Width           =   1335
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2355;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   8
               Left            =   105
               TabIndex        =   15
               Tag             =   "CIFSea"
               Top             =   1050
               Width           =   1335
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2355;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "CIF (Sea)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   20
               Left            =   165
               TabIndex        =   14
               Top             =   825
               Width           =   690
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   9
               Left            =   1470
               TabIndex        =   13
               Tag             =   "CIFAir"
               Top             =   1050
               Width           =   1335
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2355;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "CIF (Air)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   21
               Left            =   1515
               TabIndex        =   12
               Top             =   825
               Width           =   615
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   10
               Left            =   2835
               TabIndex        =   11
               Tag             =   "CnFSea"
               Top             =   1050
               Width           =   1335
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2355;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "C & F (Sea)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   22
               Left            =   2865
               TabIndex        =   10
               Top             =   825
               UseMnemonic     =   0   'False
               Width           =   825
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   11
               Left            =   4200
               TabIndex        =   9
               Tag             =   "CnFAir"
               Top             =   1050
               Width           =   1335
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2355;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "C & F (Air)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   23
               Left            =   4215
               TabIndex        =   8
               Top             =   825
               UseMnemonic     =   0   'False
               Width           =   750
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Currency"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   19
               Left            =   150
               TabIndex        =   7
               Top             =   240
               Width           =   660
            End
            Begin MSForms.ComboBox cmbCurr 
               Height          =   315
               Left            =   105
               TabIndex        =   6
               Tag             =   "ItemCurr"
               Top             =   450
               Width           =   1335
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "2355;556"
               ListWidth       =   3527
               MatchEntry      =   1
               ListStyle       =   1
               ShowDropButtonWhen=   1
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin ComboList.Usercontrol1 cmbSaleAcc 
            Height          =   285
            Left            =   -74835
            TabIndex        =   57
            Tag             =   "SaleAccNo"
            Top             =   1260
            Width           =   3315
            _ExtentX        =   5847
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
         Begin ComboList.Usercontrol1 cmbSaleRAcc 
            Height          =   285
            Left            =   -74835
            TabIndex        =   58
            Tag             =   "SRTAccNo"
            Top             =   1800
            Width           =   3315
            _ExtentX        =   5847
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
         Begin ComboList.Usercontrol1 cmbAssetAcc 
            Height          =   285
            Left            =   -74835
            TabIndex        =   59
            Tag             =   "AssetAccNo"
            Top             =   2370
            Width           =   3315
            _ExtentX        =   5847
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
         Begin MSComctlLib.ListView LV1 
            Height          =   2370
            Left            =   -73890
            TabIndex        =   61
            Top             =   1905
            Width           =   6825
            _ExtentX        =   12039
            _ExtentY        =   4180
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Process"
               Object.Width           =   4939
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Component"
               Object.Width           =   4939
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Quantity"
               Object.Width           =   1764
            EndProperty
         End
         Begin MSComctlLib.ListView LV3 
            Height          =   3075
            Left            =   -74910
            TabIndex        =   62
            Top             =   915
            Width           =   7350
            _ExtentX        =   12965
            _ExtentY        =   5424
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
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
               Text            =   "Catalog"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Item Ref ID"
               Object.Width           =   2540
            EndProperty
         End
         Begin ComboList.Usercontrol1 cmbProcRM 
            Height          =   285
            Left            =   -74115
            TabIndex        =   93
            Top             =   1335
            Width           =   2730
            _ExtentX        =   4815
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
         Begin MSComctlLib.ListView LVRM 
            Height          =   2370
            Left            =   -74130
            TabIndex        =   95
            Top             =   1650
            Width           =   6825
            _ExtentX        =   12039
            _ExtentY        =   4180
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Process"
               Object.Width           =   4939
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Component"
               Object.Width           =   4939
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Quantity"
               Object.Width           =   1764
            EndProperty
         End
         Begin VB.Image ImgTechDrawing 
            BorderStyle     =   1  'Fixed Single
            Height          =   3525
            Left            =   -74850
            Stretch         =   -1  'True
            Top             =   1230
            Width           =   4290
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tech. Drawing"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   60
            Left            =   -74820
            TabIndex        =   158
            Top             =   1035
            Width           =   1035
         End
         Begin VB.Label lblTechDrawing 
            Alignment       =   2  'Center
            Caption         =   "Double Click to Add"
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   -73410
            TabIndex        =   157
            Top             =   2790
            Width           =   1125
         End
         Begin MSForms.TextBox TBox 
            Height          =   630
            Index           =   12
            Left            =   -70605
            TabIndex        =   129
            Tag             =   "Polybag"
            Top             =   2325
            Width           =   4050
            VariousPropertyBits=   -1467987941
            BorderStyle     =   1
            Size            =   "7144;1111"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Polybag :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   48
            Left            =   -70575
            TabIndex        =   128
            Top             =   2100
            Width           =   675
         End
         Begin MSForms.TextBox TBox 
            Height          =   1110
            Index           =   11
            Left            =   -70605
            TabIndex        =   127
            Tag             =   "PackingInstructions"
            Top             =   945
            Width           =   4050
            VariousPropertyBits=   -1467987941
            BorderStyle     =   1
            Size            =   "7144;1958"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Packing Instructions"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   47
            Left            =   -70575
            TabIndex        =   126
            Top             =   720
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "RM Components"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   34
            Left            =   -71325
            TabIndex        =   101
            Top             =   1140
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Quanity"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   33
            Left            =   -68550
            TabIndex        =   100
            Top             =   1125
            Width           =   570
         End
         Begin MSForms.TextBox NBox 
            Height          =   285
            Index           =   13
            Left            =   -68580
            TabIndex        =   99
            Top             =   1335
            Width           =   1020
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "1799;503"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.CommandButton cmdAddRM 
            Height          =   345
            Left            =   -67530
            TabIndex        =   98
            Top             =   1275
            Width           =   795
            Caption         =   "Add"
            PicturePosition =   2
            Size            =   "1402;609"
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.ComboBox cmbRM 
            Height          =   285
            Left            =   -72450
            TabIndex        =   97
            Top             =   4515
            Visible         =   0   'False
            Width           =   2730
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4815;503"
            ListWidth       =   8819
            MatchEntry      =   1
            ListStyle       =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Process"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   30
            Left            =   -74100
            TabIndex        =   96
            Top             =   1125
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Process"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   29
            Left            =   -73860
            TabIndex        =   91
            Top             =   1365
            Width           =   555
         End
         Begin MSForms.TextBox NBox 
            Height          =   315
            Index           =   3
            Left            =   -72600
            TabIndex        =   90
            Tag             =   "SnaffRate"
            Top             =   5160
            Visible         =   0   'False
            Width           =   1335
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "2355;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.TextBox NBox 
            Height          =   315
            Index           =   2
            Left            =   -72600
            TabIndex        =   89
            Tag             =   "StampRate"
            Top             =   4695
            Visible         =   0   'False
            Width           =   1335
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "2355;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.TextBox NBox 
            Height          =   315
            Index           =   1
            Left            =   -69615
            TabIndex        =   88
            Tag             =   "PlantRate"
            Top             =   4710
            Visible         =   0   'False
            Width           =   1335
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "2355;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.CheckBox chkPlanter 
            Height          =   345
            Left            =   -71040
            TabIndex        =   87
            Top             =   4695
            Visible         =   0   'False
            Width           =   1365
            VariousPropertyBits=   1015023643
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2408;609"
            Value           =   "0"
            Caption         =   "Planting Rate"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkSnaffer 
            Height          =   345
            Left            =   -74025
            TabIndex        =   86
            Top             =   5145
            Visible         =   0   'False
            Width           =   1395
            VariousPropertyBits=   1015023643
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2461;609"
            Value           =   "0"
            Caption         =   "Snaffing Rate"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkStamper 
            Height          =   345
            Left            =   -74025
            TabIndex        =   85
            Top             =   4650
            Visible         =   0   'False
            Width           =   1455
            VariousPropertyBits=   1015023643
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2566;609"
            Value           =   "0"
            Caption         =   "Stamping Rate"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Asset Account"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   14
            Left            =   -74820
            TabIndex        =   72
            Top             =   2160
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sale Return Account"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   12
            Left            =   -74805
            TabIndex        =   71
            Top             =   1590
            Width           =   1470
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sale Account"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   13
            Left            =   -74805
            TabIndex        =   70
            Top             =   1050
            Width           =   930
         End
         Begin MSForms.ComboBox cmbComp 
            Height          =   285
            Left            =   -71115
            TabIndex        =   69
            Top             =   1575
            Width           =   2730
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4815;503"
            ListWidth       =   8819
            MatchEntry      =   1
            ListStyle       =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CommandButton cmdAdd 
            Height          =   345
            Left            =   -67290
            TabIndex        =   68
            Top             =   1515
            Width           =   795
            Caption         =   "Add"
            PicturePosition =   2
            Size            =   "1402;609"
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.TextBox NBox 
            Height          =   285
            Index           =   19
            Left            =   -68340
            TabIndex        =   67
            Top             =   1575
            Width           =   1020
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "1799;503"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Quanity"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   31
            Left            =   -68310
            TabIndex        =   66
            Top             =   1365
            Width           =   570
         End
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   3525
            Left            =   -70530
            Stretch         =   -1  'True
            Top             =   1230
            Width           =   4290
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Picture"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   32
            Left            =   -70530
            TabIndex        =   65
            Top             =   1050
            Width           =   495
         End
         Begin VB.Label lblPic 
            Alignment       =   2  'Center
            Caption         =   "Double click to Add"
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   -68970
            TabIndex        =   64
            Top             =   2790
            Width           =   1125
         End
         Begin MSForms.CommandButton cmdComp 
            Height          =   390
            Left            =   -71115
            TabIndex        =   63
            Top             =   1155
            Width           =   1470
            Caption         =   "Component"
            PicturePosition =   2
            Size            =   "2593;688"
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   4
         Left            =   90
         TabIndex        =   78
         Top             =   1335
         Visible         =   0   'False
         Width           =   1725
         VariousPropertyBits=   679495707
         MaxLength       =   9
         Size            =   "3043;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame FRev 
      Caption         =   "Revision ID :"
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
      Height          =   615
      Left            =   1170
      TabIndex        =   80
      Top             =   7200
      Visible         =   0   'False
      Width           =   2055
      Begin VB.Label lblRevID 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1005
         TabIndex        =   81
         Tag             =   "RevID"
         Top             =   240
         Width           =   60
      End
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   375
      Left            =   8010
      TabIndex        =   79
      Top             =   7365
      Width           =   1575
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "2778;661"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CheckBox chkInactive 
      Height          =   345
      Left            =   105
      TabIndex        =   2
      Tag             =   "InActive"
      Top             =   7365
      Width           =   1170
      BackColor       =   -2147483633
      ForeColor       =   8388608
      DisplayStyle    =   4
      Size            =   "2064;609"
      Value           =   "0"
      Caption         =   "Inactive"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   375
      Left            =   6405
      TabIndex        =   1
      Top             =   7365
      Width           =   1575
      Caption         =   " Save & Close"
      PicturePosition =   327683
      Size            =   "2778;661"
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdNext 
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   7365
      Width           =   1575
      Caption         =   "  Save & New"
      PicturePosition =   327683
      Size            =   "2778;661"
      Accelerator     =   78
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmNewItemOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ItemCode As String
'Dim CopyItemCode As String
Dim PicFileName As String, TechDrawingFileName As String
Dim SelItm As ListItem
Dim RevEntry As New DMRRev '
Dim DoPrint As Boolean

Private Sub AddTechDrawing()

    On Error Resume Next
    
    TechDrawingFileName = ""
    CD1.ShowSave
    If Not (err.Number = 32755) Then
        TechDrawingFileName = CD1.FileName
        ImgTechDrawing.Picture = LoadPicture(TechDrawingFileName)
        lblTechDrawing.Visible = False
    Else
        TechDrawingFileName = ""
        imgttechdrawing.Picture = LoadPicture(TechDrawingFileName)
        lblTechDrawing.Visible = True
    End If
    
End Sub


Private Sub AddPic()
    On Error Resume Next
    
    PicFileName = ""
    CD1.ShowSave
    If Not (err.Number = 32755) Then
        PicFileName = CD1.FileName
        Image1.Picture = LoadPicture(PicFileName)
        lblPic.Visible = False
    Else
        PicFileName = ""
        Image1.Picture = LoadPicture(PicFileName)
        lblPic.Visible = True
    End If
    
End Sub

Public Sub EditItem(strItemCode As String)

    On Error GoTo err
    'Dim rs As New ADODB.Recordset
    ItemCode = ""
    ItemCode = Left(strItemCode, Len(strItemCode) - 1)
    'cmbCatID.Locked = True
    If strCompany = "Sunlike" Or strCompany = "QEL" Then
        Tbox(3).Locked = False
    Else
        Tbox(3).Locked = True
    End If
    Call ShowData(ItemCode)
    Me.Show 1
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub AddProcesses()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        '.Open "Select * From Processes Where ProcessID Not In(2,3,5,8) Order By Sno", con, adOpenForwardOnly, adLockReadOnly
        .Open "Select * From Processes Order By Sno", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            Call LV2.ListItems.add(, ![ProcessID] & "'", ![SNo] & ". " & ![Description])
            .MoveNEXT
        Loop
        .Close
        
'        .Open "Select * From FDAProcesses Order By Sno", con, adOpenForwardOnly, adLockReadOnly
'        Do Until .EOF
'            Call LV4.ListItems.add(, ![ProcessID] & "'", ![Sno] & ". " & ![Description])
'
'            .MoveNext
'        Loop
'        .Close
        
    End With
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub AddCatals()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "Select * From ForeignCatalog Order By CatalogName", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            Set ITM = LV3.ListItems.add(, ![CatalID] & "'", ![CatalogName])
            ITM.SubItems(1) = ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub FillCmbs()
        
    'Call AddToCombo(cmbPchAcc, "AccTitle", "VActiveAccounts", " Where AccNo Like '" & PchAcc & "%' AND Parent=0 Order By ACCTitle")
    Call cmbSaleAcc.AddVals(con, "AccTitle", "VActiveAccounts", "AccNo", " Where AccNo Like '" & SALEACC & "%' AND Parent=0 Order By ACCTitle")
    cmbSaleAcc.id = "31-001-001"
    'Call AddToCombo(cmbPRTAcc, "AccTitle", "VActiveAccounts", " Where AccNo Like '" & PRTAcc & "%' AND Parent=0 Order By ACCTitle")
    Call cmbSaleRAcc.AddVals(con, "AccTitle", "VActiveAccounts", "AccNo", " Where AccNo Like '" & SRTACC & "%' AND Parent=0 Order By ACCTitle")
    cmbSaleRAcc.id = "31-003-003"
    
    Call cmbAssetAcc.AddVals(con, "AccTitle", "VActiveAccounts", "AccNo", " Where AccNo Like '" & STOCKASSETACC & "%' AND Parent=0 Order By ACCTitle")
    cmbAssetAcc.id = "15-005-001"
    'Call AddToCombo(cmbAssetAcc, "AccTitle", "VActiveAccounts", " Where AccNo Like '" & StockAssetAcc & "%' AND Parent=0 Order By ACCTitle")
    
    'Call AddToCombo(cmbUnit, "Distinct Unit", "Items", "Order By Unit")
    'Call AddToCombo(cmbWeightUnit, "Distinct UnitWeight", "Items", "Order By UnitWeight")
    
    Call AddToCombo(cmbComp, "Description", "Components", , , "CompID")
    
    cmbCompProc.ListHeight = 2400
    Call cmbCompProc.ClearVals
    Call cmbCompProc.AddVals(con, "Description", "Processes", "ProcessID")
    
    
    'Call AddToCombo(cmbRM, "RMName + '( ' + Unit + ' )'", "RM", , , "RMID,RMName,Unit")
    cmbRawMaterial.AddVals con, "RMName + '( ' + Unit + ' )'", "RM", "RMID"
    
    cmbProcRM.ListHeight = 2400
    cmbRawMaterial.ListHeight = 2400
'    Call cmbProcRM.ClearVals
'    Call cmbProcRM.AddVals(con, "Description", "Processes", "ProcessID")
    
    Call AddToCombo(cmbCatID, "Description", "ItemCatagories", " Order By Description", , "CatID")
    
    Call AddToCombo(cmbSteel, "SteelType", "SteelTypes", , , "SteelID")
    If cmbSteel.ListCount > 0 Then cmbSteel.ListIndex = 0
    
    Call AddToCombo(cmbItemGroup, "Description", "ItemGroups", " Order By Description", , "ID")
    If cmbItemGroup.ListCount > 0 Then cmbItemGroup.ListIndex = 0
    
    Call AddToCombo(cmbFinQuality, "Description", "ItemsFinQuality", , , "Description")
    Call AddToCombo(cmbType, "Description", "ItemTypes", , , "Description")
    '----------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select UnitName From Units Order By UnitName", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            cmbUnit.AddItem .Fields(0) & ""
            cmbSizeUnit.AddItem .Fields(0) & ""
            'cmbWeightUnit.AddItem .Fields(0) & ""
            .MoveNEXT
        Loop
        .Close
    End With
    If cmbUnit.ListCount > 0 Then cmbUnit.Text = "Pcs"
    
    Set rs = Nothing
    
    With cmbWeightUnit
        .AddItem "Kgs"
        .AddItem "Grams"
    End With
    '----------------------------------------------------------------
    cmbCurr.AddItem "EUR"
    cmbCurr.ListIndex = 0
    cmbCurr.Enabled = False
    
    With cmbItemType
        .AddItem "Finished"
        .List(0, 1) = 0
        If strCompany = "Tecno" Or strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
            .AddItem "Semi Finished"
            .List(1, 1) = 1
        End If
        .ListIndex = 0
    End With
    
End Sub

Private Sub chkPlanter_Click()
    NBox(1).Visible = chkPlanter.Value
    If Not chkPlanter.Value Then NBox(1) = ""
End Sub

Private Sub chkSnaffer_Click()
    NBox(3).Visible = chkSnaffer.Value
    If Not chkSnaffer.Value Then NBox(3) = ""
End Sub

Private Sub chkStamper_Click()
    NBox(2).Visible = chkStamper.Value
    If Not chkStamper.Value Then NBox(2) = ""
End Sub

Private Sub cmbCatID_Change()
    Call cmbCatID_Click
End Sub

Private Sub cmbCatID_Click()
    If Not cmbCatID.MatchFound Then Exit Sub
End Sub

Private Sub cmbSteel_Change()
    Call cmbSteel_Click
End Sub

Private Sub cmbSteel_Click()

    If cmbSteel.MatchFound Then
        Call AddToCombo(cmbGuage, "Gage", "SteelGages", " Where SteelID=" & cmbSteel.List(cmbSteel.ListIndex, 1), , "GageID")
    End If
    If cmbGuage.ListCount > 0 Then cmbGuage.ListIndex = 0
    
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo err
    If Not cmbComp.MatchFound Then Exit Sub
    Dim ITM As ListItem
    
    For i = 1 To LV1.ListItems.count
        If cmbComp.List(cmbComp.ListIndex, 1) & "'" = LV1.ListItems(i).key Then
            MsgBox "This Component Is Already Added.", vbInformation
            Exit Sub
        End If
    Next i
    
    
    Set ITM = LV1.ListItems.add(, cmbComp.List(cmbComp.ListIndex, 1) & "'", cmbCompProc.Text)
    ITM.Tag = cmbCompProc.id
    ITM.ListSubItems.add , , cmbComp.Text
    ITM.ListSubItems.add , , Val(NBox(19))
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdAddRM_Click()

    On Error GoTo err
    If Not cmbRawMaterial.MatchFound Then Exit Sub
    Dim ITM As ListItem
    
    For i = 1 To LVRM.ListItems.count
        If cmbRawMaterial.id & "'" = LVRM.ListItems(i).key Then
            MsgBox "This Component Is Already Added.", vbInformation
            Exit Sub
        End If
    Next i
    
    Dim strUnit As String, strMaterialName As String
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT RMName,Unit FROM RM WHERE RMID=" & cmbRawMaterial.id, con, adOpenForwardOnly, adLockReadOnly
        strMaterialName = !RMName & ""
        strUnit = !Unit & ""
        .Close
    End With
    Set rs = Nothing
    
    Set ITM = LVRM.ListItems.add(, cmbRawMaterial.id & "'", cmbProcRM.Text)
    ITM.Tag = cmbProcRM.id
    ITM.ListSubItems.add , , strMaterialName
    ITM.ListSubItems.add , , Val(NBox(13))
    
    Exit Sub
err:
    MsgBox err.Description

End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdDelete_Click()
    LV1.ListItems.Remove (LV1.SelectedItem.key)
End Sub

Private Sub cmdComp_Click()
    frmComponents.Show 1
End Sub

Private Sub cmdDeleteRM_Click()
    LVRM.ListItems.Remove (LVRM.SelectedItem.key)
End Sub

Private Sub cmdNext_Click()
    If SaveItem Then
        Unload Me
        
        Me.Show 1
    End If
End Sub

Private Sub cmdOK_Click()
    If SaveItem Then Unload Me
End Sub


Private Sub cmdViewPicture_Click()

    If Image1.Picture = 0 Then Exit Sub
    Dim f As New frmViewPictureFSExt
    Load f
    f.ShowFromPic Image1.Picture

End Sub

Private Sub cmdViewTechDrawing_Click()
    If ImgTechDrawing.Picture = 0 Then Exit Sub
    Dim f As New frmViewPictureFSExt
    Load f
    f.ShowFromPic ImgTechDrawing.Picture
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    PicFileName = ""
    'For i = 0 To SSTab1.Tabs - 1
    '    SSTab1.Tab = i
    '    SSTab1.Picture = ImageList1.ListImages(i & "'").ExtractIcon
    'Next
       
    SSTab1.TabVisible(1) = True
    SSTab1.TabVisible(2) = False
    SSTab1.TabVisible(4) = False
    SSTab1.TabVisible(6) = False
    SSTab1.TabVisible(8) = True
    
    If CurrentUserName <> "Administrator" Then
        SSTab1.TabVisible(5) = False
        NBox(24).Locked = True
    End If
    
    If LCase(CurrentUserName) = "production" Then
        SSTab1.TabVisible(5) = False
    End If
    SSTab1.Tab = 0
    
    SSTab2.TabVisible(1) = False
    
    Me.KeyPreview = True
    
    'DTPicker1.value = ServerDate
    
    cmbSaleAcc.ListHeight = 1500
    cmbSaleRAcc.ListHeight = 1500
    cmbAssetAcc.ListHeight = 1500
    
    cmbCatID.Locked = False
    Tbox(3).Locked = False
    Tbox(6).Text = ServerDate
    
    Call AddProcesses
    Call AddCatals
    Call FillCmbs
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select EuroToUS From EuroToUS", con
        If .EOF = False Then
            txtConversion = Val(.Fields(0).Value & "")
        End If
        .Close
    End With
    Set rs = Nothing
    
    FRev.Visible = False
   
    'CopyItemCode = ""
End Sub

Private Function SaveItem() As Boolean

    On Error GoTo err

    Dim OpBalance As Double
    
    If Trim(Tbox(0)) = "" Then
        MsgBox "Invalid Item Name", vbInformation
        SaveItem = False
        Exit Function
    ElseIf Trim(cmbUnit) = "" Then
        MsgBox "No Unit Selected.", vbInformation
        SaveItem = False
        Exit Function
'    ElseIf cmbAssetAcc.Enabled And Not cmbAssetAcc.MatchFound Then
'        MsgBox "No Asset Account Selected", vbInformation
'        SaveItem = False
'        Exit Function
    ElseIf cmbItemType.MatchFound = False Then
        If strCompany = "Tecno" Or strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
            MsgBox "Please Select Item Type.", vbInformation
            Exit Function
        Else
            cmbItemType.ListIndex = 0
        End If
    End If
    
    If cmbSteel.MatchFound = False Then
        MsgBox "Please Select Steel.", vbInformation
        Exit Function
    End If
    
    If cmbGuage.MatchFound = False Then
        MsgBox "Please Select Guage.", vbInformation
        Exit Function
    End If
    
    Dim PCHACC As String, ItmPRTAcc As String, AssetAcc As String

    Dim Sql As String
    
    Call StartTrans(con)

    If ItemCode = "" Then
        Sql = CreateInsertSQL("Items")
        'con.Execute "Update Accounts Set OpenBal=OpenBal+" & TotalCost & " Where AccNo='" & AssetAcc & "'"
        'Call adjustLedger(AssetAcc, "1/1/1900", 0, TotalCost)
    Else
        Sql = CreateUpdateSQL("Items")
        Sql = Sql & " Where ItemID='" & ItemCode & "'"
    End If
    
    con.Execute Sql
    If bUrduSupported Then
        con.Execute "UPDATE Items SET ItemNameUrdu=N'" & txtUrduFields(0).Text & "' WHERE ItemID='" & Tbox(3).Text & "'", i
    End If
    'If New Item Then Assign This New Item To HICO Vender.
'    If ItemCode = "" Then
'        con.Execute "Insert Into VendAssItems(VendID,ItemID,Rate) Values(" & _
'         "129,'" & TBox(3) & "',0)"
'    End If
         
'    'If New Item Then Assign This New Item To HICO Customer.
'    If ItemCode = "" Then
'        con.Execute "Insert Into FCustomerCatalog(CustCode,Country,ItemID,Unit,CompItemID,Description) Values(" & _
'         "'HICO','Pakistan','" & TBox(3) & "','" & cmbUnit & "','" & TBox(3) & "','" & TBox(0) & " " & TBox(7) & " " & cmbSizeUnit & " " & cmbType & " " & cmbFinQuality & "')"
'    End If
    
    If strCompany = "Weldon Industries" And ItemCode = "" Then  'Assign All Processes to new item
        For i = 1 To LV2.ListItems.count
            LV2.ListItems(i).Checked = True
        Next
    End If
    
    Call SaveComponents
    Call SaveComponentsRM
    Call SaveProcesses
    Call SaveForeignCatal
    
    If strCompany = "Weldon Instruments" Then
        'Assign All Processes to Factory Maker....
        If ItemCode <> "" Then
            con.Execute "DELETE FROM VendAssItems WHERE ItemID='" & Tbox(3) & "' AND VendID=15"
        End If
        con.Execute "INSERT INTO VendAssItems(VendID,ProcessID,ItemID) SELECT DISTINCT 15,ProcessID,ItemID FROM ItemProcesses " & _
         "WHERE ProcessID IN(SELECT ProcessID FROM Processes WHERE Operation IN(0,2)) AND ItemID='" & Tbox(3) & "' ORDER BY ItemID,ProcessID"
    End If
    'If ItemCode <> "" Then Call SaveRevision
    'If CopyItemCode <> "" Then
    '    con.Execute "Update Items Set ItemPic=(Select ItemPic From Items Where ItemID='" & CopyItemCode & "') Where ItemID='" & TBox(3) & "'"
    'End If

    con.CommitTrans
    
    '---------------------------------------------------------
    'If DoPrint Then
    '    If MsgBox("Item Has Changed." & vbNewLine & "Do You Want To Print This Item?", vbQuestion + vbYesNo) = vbYes Then
    '        Load frmPrevRpt
    '        frmPrevRpt.ShowReport "{VrptDMR.ItemID}='" & TBox(4) & "'", rptDMR
    '    End If
    'End If
    '---------------------------------------------------------
    ItemCode = ""
    SaveItem = True
    
SavePic:

    If Image1.Picture <> 0 And Trim(PicFileName) <> "" Then
        Dim Rec As New ADODB.Recordset
        Dim rsPic As New ADODB.Stream
        'rec.Close
        Rec.Open "Select ItemID,ItemPic from Items Where ItemID='" & Tbox(3) & "'", con, adOpenStatic, adLockOptimistic
        If Rec.RecordCount > 0 Then
            rsPic.Open
            rsPic.Type = adTypeBinary
            rsPic.LoadFromFile PicFileName
            Rec.Fields(1) = rsPic.Read
            rsPic.Close
            Rec.Update
        End If
        
        Rec.Close
        
        Set Rec = Nothing
        Set rsPic = Nothing
    End If
    
    If ImgTechDrawing.Picture <> 0 And Trim(TechDrawingFileName) <> "" Then
        
        'rec.Close
        Rec.Open "SELECT ItemID,TechnicalDrawing FROM Items WHERE ItemID='" & Tbox(3) & "'", con, adOpenStatic, adLockOptimistic
        If Rec.RecordCount > 0 Then
            rsPic.Open
            rsPic.Type = adTypeBinary
            rsPic.LoadFromFile TechDrawingFileName
            Rec.Fields(1) = rsPic.Read
            rsPic.Close
            Rec.Update
        End If
        
        Rec.Close
        
        Set Rec = Nothing
        Set rsPic = Nothing
    End If
    
    Exit Function
err:
    MsgBox err.Description
    
End Function

Private Sub SaveRevision()
    con.Execute "Insert Into ItemRevisions(ItemID,RevDTTime) Values('" & Tbox(3) & "','" & Now & "')"
End Sub




Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
    Set RevEntry = Nothing
End Sub

Private Sub Form_Unload(cancel As Integer)
    OurOwnForm = True
End Sub

Private Sub Image1_DblClick()
    Call AddPic
End Sub



Private Sub ImgTechDrawing_DblClick()
    Call AddTechDrawing
End Sub

Private Sub lblPic_DblClick()
    Call AddPic
End Sub


Private Sub lblTechDrawing_DblClick()
    Call AddTechDrawing
End Sub

Private Sub LV1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With cmdDelete
        
        If LV1.ListItems.count > 0 Then
            .Move .Left, LV1.Top + Item.Top
            .Visible = True
        Else
            .Visible = False
        End If
    End With
End Sub



Private Sub LV2_DblClick()
    
    
    If LV2.ListItems.count = 0 Then Exit Sub
    If LV2.SelectedItem.Checked = False Then Exit Sub
    Set SelItm = LV2.SelectedItem
    
'    With SelItm
'        txtProcRate.Move SSTab2.Left + LV2.ColumnHeaders(2).Left, .Top + LV2.Top, LV2.ColumnHeaders(2).Width
'        txtProcRate.Visible = True
'        txtProcRate = Val(.SubItems(1) & "")
'        txtProcRate.SetFocus
'        txtProcRate.Tag = "Edit"
'    End With
    With txtSortNo
        .Move LV2.Left + LV2.ColumnHeaders(2).Left, LV2.Top + SelItm.Top, LV2.ColumnHeaders(2).Width
        .Text = Val(SelItm.SubItems(1))
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
    
End Sub

Private Sub LV2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    
    Dim i As Integer, j As Integer, striProcID As String
    j = 1
    striProcID = ""
    For i = 1 To LV2.ListItems.count
        With LV2.ListItems(i)
            If .Checked Then
                striProcID = striProcID & Val(.key) & ","
                .SubItems(1) = j
                j = j + 1
            End If
        End With
    Next
    cmbProcRM.ClearVals
    If striProcID <> "" Then
        striProcID = Left(striProcID, Len(striProcID) - 1)
    End If
    cmbProcRM.AddVals con, "Description", "Processes", "ProcessID", " WHERE ProcessID IN(" & striProcID & ") ORDER BY SNO"
    
    Exit Sub
    With Item
        If .Checked Then
            txtProcRate.Move SSTab2.Left + LV2.ColumnHeaders(2).Left - 30, .Top + LV2.Top, LV2.ColumnHeaders(2).Width
            txtProcRate.Visible = True
            txtProcRate = ""
            txtProcRate.SetFocus
            txtProcRate.Tag = "Add"
        Else
            .SubItems(1) = ""
        End If
    End With
    Set SelItm = Item
End Sub

Private Sub LV3_DblClick()
    Exit Sub
    With txtRefID
        If LV3.SelectedItem.Checked Then
            .Move LV3.ColumnHeaders(2).Left + LV3.Left, LV3.Top + LV3.SelectedItem.Top, LV3.ColumnHeaders(2).Width
            .Text = LV3.SelectedItem.SubItems(1)
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End If
    End With
End Sub

Private Sub LV3_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Exit Sub
    With txtRefID
        If Item.Checked Then
            .Move LV3.ColumnHeaders(2).Left + LV3.Left, LV3.Top + Item.Top, LV3.ColumnHeaders(2).Width
            .Text = Item.SubItems(1)
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        Else
            Item.SubItems(1) = ""
        End If
    End With
End Sub

Private Sub LVRM_KeyDown(KeyCode As Integer, Shift As Integer)
    If LVRM.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you Sure to Remove?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
        LVRM.ListItems.Remove LVRM.SelectedItem.Index
    End If
End Sub

Private Sub NBox_Change(Index As Integer)
    Select Case Index
        Case 4, 5, 7
            NBox(6) = Val(NBox(4)) - Val(NBox(5)) - Val(NBox(7))
        Case 0, 14, 15, 16, 8, 9, 10, 11
            Call txtConversion_Change
    End Select
    
End Sub



Private Sub Tbox_Change(Index As Integer)
    Exit Sub
    Select Case Index
    
    Case 8, 3
        Tbox(4) = Tbox(8) & Tbox(3)
    End Select
    
End Sub

Private Sub Tbox_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    Exit Sub
    Select Case Index
    Case 2, 3
        KeyAscii = OnlyNumber(Val(KeyAscii))
    Case 5
        KeyAscii = OnlyNumber(Val(KeyAscii), True)
    End Select
    
End Sub

Private Sub TBox_LostFocus(Index As Integer)
    If strCompany = "QEL" Then
        If Index = 3 Then
            '0000-000-000
            'First Check If there is 2 Dashes...
            Dim iPos1 As Integer, iPos2 As Integer, iCount As Integer
            iPos1 = InStr(Tbox(3), "-")
            If iPos1 > 0 Then
                iPos2 = InStr(iPos1 + 1, Tbox(3), "-")
            End If
            If iPos1 = 0 Or iPos2 = 0 Then
                MsgBox "Can't Format the Item Code.", vbInformation
                Exit Sub
            End If
            'Now Format it...
            Dim strFirst As String, strSecond As String, strThird As String
            strFirst = Left(Tbox(3), iPos1 - 1)
            strSecond = Mid(Tbox(3), iPos1 + 1, iPos2 - (iPos1 + 1))
            strThird = Mid(Tbox(3), iPos2 + 1)
            
            Tbox(3) = Format(strFirst, "0000") & "-" & Format(strSecond, "000") & "-" & Format(strThird, "000")
        End If
    End If
End Sub

Private Sub TBox_Validate(Index As Integer, cancel As Boolean)
    Exit Sub
    If Index = 5 Then
        Tbox(Index) = Format(Val(Tbox(Index)), "#0.00")
    ElseIf Index = 2 Or Index = 3 Then
        Tbox(Index) = Format(Val(Tbox(Index)), "#0")
    End If
End Sub

Private Sub SaveComponents()

    con.Execute "Delete From ItemsComp Where ItemID='" & ItemCode & "'"

    For i = 1 To LV1.ListItems.count
        con.Execute "Insert Into ItemsComp(ItemID,CompID,Qty,ProcessID) Values('" & Tbox(3) & "'," & Val(LV1.ListItems(i).key) & "," & LV1.ListItems(i).SubItems(2) & "," & Val(LV1.ListItems(i).Tag) & ")"
    Next i
    
End Sub
Private Sub SaveComponentsRM()

    con.Execute "DELETE FROM ItemsRMComp WHERE ItemID='" & ItemCode & "'"

    For i = 1 To LVRM.ListItems.count
        con.Execute "INSERT INTO ItemsRMComp(ItemID,RMID,Qty,ProcessID) VALUES('" & Tbox(3) & "'," & Val(LVRM.ListItems(i).key) & "," & Val(LVRM.ListItems(i).SubItems(2)) & "," & Val(LVRM.ListItems(i).Tag) & ")"
    Next i
    
End Sub

Private Sub SaveProcesses()
    
    'Save Item's Production Processes...
    '------------------------------------------------------------------------
    con.Execute "DELETE FROM ItemProcesses WHERE ItemID='" & ItemCode & "'"
    For i = 1 To LV2.ListItems.count
        If LV2.ListItems(i).Checked Then
            'con.Execute "Insert Into ItemProcesses(ItemID,ProcessID,Rate) " & _
             " Values('" & TBox(4) & "'," & Val(LV2.ListItems(i).Key) & "," & _
             Val(LV2.ListItems(i).SubItems(1) & "") & ")"
             con.Execute "INSERT INTO ItemProcesses(ItemID,ProcessID,SNO) " & _
              " VALUES('" & Tbox(3) & "'," & Val(LV2.ListItems(i).key) & "," & Val(LV2.ListItems(i).SubItems(1)) & ")"
        End If
    Next i
    '------------------------------------------------------------------------
    
    'Now Save Item's FDA Processes
    '------------------------------------------------------------------------
'    con.Execute "Delete From ItemProcessesFDA Where ItemID='" & ItemCode & "'"
'    For i = 1 To LV4.ListItems.count
'        If LV4.ListItems(i).Checked Then
'            con.Execute "Insert Into ItemProcessesFDA(ItemID,ProcessID) " & _
'             " Values('" & TBox(3) & "'," & Val(LV4.ListItems(i).key) & ")"
'        End If
'    Next i
    '------------------------------------------------------------------------
    
End Sub
Private Sub SaveForeignCatal()

    con.Execute "Delete From ItemCatalRef Where ItemID='" & ItemCode & "'"
    For i = 1 To LV3.ListItems.count
        If LV3.ListItems(i).Checked Then
            If LV3.ListItems(i).SubItems(1) <> "" Then
                con.Execute "Insert Into ItemCatalRef(CatalID,ItemID,ItemRefID) Values(" & Val(LV3.ListItems(i).key) & ",'" & Tbox(3) & "','" & LV3.ListItems(i).SubItems(1) & "')"
            End If
        End If
    Next i
    
End Sub
Private Function CreateInsertSQL(strTable As String) As String
    
    Dim c As Control
    Dim RetString As String
    
    
    'Create The First Part Of Insert Query...
    RetString = "Insert Into " & strTable & "("
    '---------------------------------------------------------------------------
    For Each c In Controls
        
        If (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is MSForms.ComboBox) Or (TypeOf c Is MSForms.CheckBox) Or (TypeOf c Is ComboList.Usercontrol1) Then
            If c.Tag <> "" Then
                RetString = RetString & c.Tag & ","
            End If
        End If
    Next
    
    'Get For ListView
    'For i = 1 To LV2.ListItems.Count
    '    RetString = RetString & LV2.ListItems(i).Key & ","
    'Next i
    '---------------------------------------------------------------------------
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    RetString = RetString & ") Values("
    
    
    'Now  Get The Second Part...
    '---------------------------------------------------------------------------
    For Each c In Controls
        
        
        If c.Tag <> "" Then
            If (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is MSForms.ComboBox) Then
            
                If c.Name = "cmbCatID" Or c.Name = "cmbItemType" Then
                    RetString = RetString & Val(c.List(c.ListIndex, 1)) & ","
                ElseIf c.Name = "NBox" Then
                    RetString = RetString & Val(c.Text & "") & ","
                ElseIf c.Name = "cmbSteel" Or c.Name = "cmbItemGroup" Then
                    RetString = RetString & c.List(c.ListIndex, 1) & ","
                ElseIf c.Name = "cmbGuage" Then
                    RetString = RetString & "'" & c.List(c.ListIndex, 1) & "',"
                Else
                    RetString = RetString & "'" & c.Text & "',"
                End If
                
            ElseIf (TypeOf c Is MSForms.CheckBox) Then
                RetString = RetString & Abs(c.Value) & ","
            ElseIf (TypeOf c Is ComboList.Usercontrol1) Then
                'If C.Name = cmbSteel Or C.Name = cmbGuage Then
                '    RetString = RetString & C.ID & ","
                RetString = RetString & "'" & c.id & "',"
            End If
        End If
    Next
    
    
    'Get For ListView
    'For i = 1 To LV2.ListItems.Count
    '    RetString = RetString & Abs(LV2.ListItems(i).Checked) & ","
    'Next i
    
    '---------------------------------------------------------------------------
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    RetString = RetString & ")"
    CreateInsertSQL = RetString
    
End Function
Private Function CreateUpdateSQL(strTable As String) As String
    
    Dim c As Control
    Dim RetString As String
    
    'Create The First Part Of Insert Query...
    RetString = "Update " & strTable & " Set "
    '---------------------------------------------------------------------------
    For Each c In Controls
    
        If c.Tag <> "" Then
            If (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is MSForms.ComboBox) Or (TypeOf c Is MSForms.ComboBox) Then '
                If c.Name = "cmbCatID" Or c.Name = "cmbItemType" Then
                    RetString = RetString & c.Tag & "=" & Val(c.List(c.ListIndex, 1)) & ","
                ElseIf c.Name = "NBox" Then
                    RetString = RetString & c.Tag & "=" & Val(c.Text) & ","
                ElseIf c.Name = "cmbSteel" Or c.Name = "cmbItemGroup" Then
                    RetString = RetString & c.Tag & "=" & c.List(c.ListIndex, 1) & ","
                ElseIf c.Name = "cmbGuage" Then
                    RetString = RetString & c.Tag & "='" & c.List(c.ListIndex, 1) & "',"
                Else
                    RetString = RetString & c.Tag & "='" & c.Text & "',"
                End If
            ElseIf (TypeOf c Is MSForms.CheckBox) Then
                RetString = RetString & c.Tag & "=" & Abs(c.Value) & ","
            ElseIf (TypeOf c Is ComboList.Usercontrol1) Then
                RetString = RetString & c.Tag & "='" & c.id & "',"
            End If
        End If
    Next
    
    If NotEdited = False Then
        
        'Get Rid of Zeros.
        Dim iZeros As Integer
        iZeros = 1
        Do While Mid(Tbox(3), iZeros, 1) = "0"
            iZeros = iZeros + 1
        Loop
        
        Dim strRevID As String
        'strRevID = Format(Now, "ddMMyyyyHHmm")
        strRevID = Format(Now, "ddMMyyyy") & Mid(Tbox(3), iZeros)
        
        RetString = RetString & "RevID='" & strRevID & "',"
        DoPrint = True
    Else
        DoPrint = False
    End If
    '---------------------------------------------------------------------------
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    
    CreateUpdateSQL = RetString
    
End Function

Private Sub txtConversion_Change()
    Dim dConversionRate As Double
    dConversionRate = Val(txtConversion)
    con.Execute "Delete From EuroToUS"
    con.Execute "Insert Into EuroToUS(EuroToUS) Values(" & Val(txtConversion) & ")"
    NBox(29) = Val(NBox(0)) * dConversionRate    'FOB Karachi
    NBox(28) = Val(NBox(14)) * dConversionRate    'Exworks
    NBox(27) = Val(NBox(15)) * dConversionRate    'C&I Sea
    NBox(26) = Val(NBox(16)) * dConversionRate    'C&I Air
    NBox(30) = Val(NBox(8)) * dConversionRate    'CIF Sea
    NBox(31) = Val(NBox(9)) * dConversionRate    'CIF Air
    NBox(32) = Val(NBox(10)) * dConversionRate    'C&F Sea
    NBox(33) = Val(NBox(11)) * dConversionRate    'C&F Air
End Sub

Private Sub txtProcRate_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Val(txtProcRate) <> 0 Then
            SelItm.SubItems(1) = Val(txtProcRate)
            txtProcRate.Visible = False
        End If
    ElseIf KeyAscii = 27 Then
        If txtProcRate.Tag = "Add" Then
            SelItm.Checked = False
        End If
        
        txtProcRate.Visible = False
    End If
    
End Sub

Private Sub txtProcRate_Validate(cancel As Boolean)
    If Val(LV2.SelectedItem.SubItems(1)) = 0 Then
        LV2.SelectedItem.Checked = False
    End If
    txtProcRate.Visible = False
End Sub

Private Sub txtRefID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LV3.SelectedItem.SubItems(1) = txtRefID.Text
        txtRefID.Visible = False
    ElseIf KeyAscii = 27 Then
        txtRefID.Visible = False
    End If
End Sub

Private Sub txtRefID_Validate(cancel As Boolean)
    txtRefID.Visible = False
End Sub
Private Sub AddRevEntries(rs As ADODB.Recordset)
    
    On Error GoTo err
    
    'For GInfo,Weight,Pricing.
    With RevEntry
        'For GInfo Tab
        .GDevDesc = rs![ItemName] & ""
        .GUnit = rs![Unit] & ""
        .GUsage = rs![ItemUsage] & ""
        .GType = rs![Type] & ""
        .GAttributes = rs![Attributes] & ""
        .GSize = rs![ItemSize] & ""
        .GSizeUnit = rs![SizeUnit] & ""
        .GSteelUsed = rs![SteelUsed] & ""
        .GGuageaType = rs![Gage] & ""
        .GPacknLabel = rs![PacknLabel] & ""
            
        'For Weight
        .WFinWeight = Val(rs![FinishedWeight] & "")
        .WForgeWeight = Val(rs![ForgingWeight] & "")
        .WWasteVis = Val(rs![WasteVisible] & "")
        .WWeightUnit = rs![UnitWeight] & ""
            
        'For Pricing
        .PCIFAir = Val(rs![CIFAir] & "")
        .PCIFSea = Val(rs![CIFSea] & "")
        .PCnFAir = Val(rs![CnFAir] & "")
        .PCnFSea = Val(rs![CnFSea] & "")
        .PFOB = Val(rs![FOB] & "")
        .PCurrency = rs![ItemCurr] & ""
        
    End With
        
    
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Function NotEdited() As Boolean

    On Error GoTo err
    
    With RevEntry
        'Check GInfo Tab First For Changes.
        If Tbox(0) <> .GDevDesc Or cmbUnit <> .GUnit Or Tbox(5) <> .GUsage Or cmbType <> .GType Or Tbox(1) <> .GAttributes Or Tbox(7) <> .GSize Or cmbSizeUnit <> .GSizeUnit Or cmbSteel <> .GSteelUsed Or cmbGuage <> .GGuageaType Or Tbox(2) <> .GPacknLabel Then
            Exit Function
        End If
        
        'Check Weight Tab For Changes.
        If cmbWeightUnit <> .WWeightUnit Or Val(NBox(4)) <> .WForgeWeight Or Val(NBox(5)) <> .WWasteVis Or Val(NBox(7)) <> .WFinWeight Then
            Exit Function
        End If
        
        'Check Pricing Tab For Changes.
        If cmbCurr <> .PCurrency Or Val(NBox(0)) <> .PFOB Or Val(NBox(8)) <> .PCIFSea Or Val(NBox(9)) <> .PCIFAir Or Val(NBox(10)) <> .PCnFSea Or Val(NBox(11)) <> .PCnFAir Then
            Exit Function
        End If
        
        'Check For The Prod.Processes For Changes.
        If Not CheckForLVs(LV2, .PrdPrIDs) Then
            Exit Function
        End If
        
        'Check For The FDAProcesses For Changes.
'        If Not CheckForLVs(LV4, .FDAPrIDs) Then
'            Exit Function
'        End If
        
        'Check For The FDAProcesses For Changes.
        If Not CheckForLVs(LV1, .CompIDs) Then
            Exit Function
        End If
        
    End With
    
    
    NotEdited = True
    Exit Function
    
err:
    MsgBox err.Description
    
End Function
Private Function CheckForLVs(LView As ListView, MyCol As Collection) As Boolean

    Dim iNo As Integer, iTotal As Integer
        
    'Check For The Prod.Processes For Changes.
    iNo = 0
    iTotal = 0
    
    With MyCol
        For i = 1 To LView.ListItems.count
            If LView.ListItems(i).Checked Then
                iNo = iNo + 1
                iTotal = iTotal + 1
                If iNo > .count Then
                    Exit Function
                End If
                
                If LView.Name = "LV2" Then 'Check PrcID & Rate For LV2
                    If Val(LView.ListItems(i).key) <> .Item(iNo).PrcID Then
                        Exit Function
                    Else
                        'If Val(LView.ListItems(i).SubItems(1)) <> .Item(iNo).PrcRate Then
                        '    Exit Function
                        'End If
                    End If
                Else 'Just Check The ID For Other LVs.
                    If Val(LView.ListItems(i).key) <> .Item(iNo) Then
                        Exit Function
                    End If
                End If
            End If
        Next i
        
        If iTotal <> .count Then
            Exit Function
        End If
    End With
    
    CheckForLVs = True
    
End Function
Private Sub PrintDMR()
    
End Sub

Public Sub CopyItem(strItemCode As String)
    ItemCode = ""
    
    strItemCode = Left(strItemCode, Len(strItemCode) - 1)
    'CopyItemCode = strItemCode
    Call ShowData(strItemCode, True)
    Tbox(3) = ""
    
    Me.Show 1
End Sub

Private Sub ShowData(strItemCode As String, Optional bCopy As Boolean = False)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    
    
    With rs
        .Open "Select * From VItems Where ItemID='" & strItemCode & "'", con, adOpenForwardOnly, adLockReadOnly
        
        For Each c In Controls
            If c.Tag <> "" Then
                If TypeOf c Is MSForms.CheckBox Then
                    c.Value = .Fields(c.Tag)
                ElseIf TypeOf c Is ComboList.Usercontrol1 Then
                    'C.Text = .Fields(C.Tag & "Text") & ""
                    c.id = .Fields(c.Tag) & ""
                ElseIf TypeOf c Is VB.Label Then
                    If ItemCode <> "" Then c.Caption = .Fields(c.Tag) & ""
                ElseIf c.Name = "cmbItemType" Then
                    c.ListIndex = Val(.Fields(c.Tag) & "")
                Else
                    If c.Name = "cmbCatID" Then
                        c.Text = .Fields("Category")
                    Else
                        c.Text = .Fields(c.Tag) & ""
                    End If
                End If
            End If
        Next
        cmbCurr.Text = "Eur"
        If ![PlantRate] <> 0 Then
            chkPlanter.Value = True
            NBox(1).Visible = True
        End If
        If ![StampRate] <> 0 Then
            chkStamper.Value = True
            NBox(2).Visible = True
        End If
        If ![SnaffRate] <> 0 Then
            chkSnaffer.Value = True
            NBox(3).Visible = True
        End If
        
        cmbSteel = ![SteelType] & ""
        cmbGuage = ![SteelGage] & ""
        cmbItemGroup = ![ItemGroup] & ""
        'TBox(8) = ![CatID] & "-"
        Tbox(3) = ![ITemID] & ""
        
        If Not IsNull(![ItemPic]) Then
            Set Image1.DataSource = rs
            Image1.DataField = "ItemPic"
            lblPic.Visible = False
        Else
            lblPic.Visible = True
        End If
        
        If Not IsNull(![TechnicalDrawing]) Then
            Set ImgTechDrawing.DataSource = rs
            ImgTechDrawing.DataField = "TechnicalDrawing"
            lblTechDrawing.Visible = False
        Else
            lblTechDrawing.Visible = True
        End If
        
        If ![RevID] & "" = "" Then
            FRev.Visible = False
        Else
            If ItemCode <> "" Then FRev.Visible = True
        End If
         
        txtUrduFields(0) = !ItemNameUrdu & ""
        
        '/\/\/\/\\/\/\/\/\/\/\
        If strCompany = "Sunlike" Then
            If bCopy Then
                Tbox(1) = ""
            End If
        End If
        'Add Entries To DMRRev Object For GInfo,Weight and Pricing Tabs...
        '------------------------------------------------------------------
        Call AddRevEntries(rs)
        '------------------------------------------------------------------
        .Close
        
        'Select Item's Production Processes
        '--------------------------------------------------------------------------------------------------
        .Open "Select ProcessID,Rate,SNo From ItemProcesses Where ItemID='" & ItemCode & "'", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            If Not IsNull(![ProcessID]) Then
                LV2.ListItems(![ProcessID] & "'").Checked = True
                LV2.ListItems(!ProcessID & "'").SubItems(1) = !SNo & ""
                'LV2.ListItems(![ProcessID] & "'").SubItems(1) = Val(![Rate] & "")
                '----------------------------------
                Dim Obj As New PrRate
                Obj.PrcID = ![ProcessID]
                Obj.PrcRate = Val(![Rate] & "")
                RevEntry.PrdPrIDs.add Obj '.Fields(0).value '![ProcessID]
                '----------------------------------
                Set Obj = Nothing
            End If
            .MoveNEXT
        Loop
        .Close
        cmbProcRM.ClearVals
        cmbProcRM.AddVals con, "Description", "VItemProcesses", "ProcessID", " WHERE IsExist='" & ItemCode & "'"
        '--------------------------------------------------------------------------------------------------
        
        'Select Item's Production Processes
        '--------------------------------------------------------------------------------------------------
        .Open "Select ProcessID From ItemProcessesFDA Where ItemID='" & ItemCode & "'", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            If Not IsNull(![ProcessID]) Then
                LV4.ListItems(![ProcessID] & "'").Checked = True
                '----------------------------------
                RevEntry.FDAPrIDs.add .Fields(0).Value
                '----------------------------------
            End If
            .MoveNEXT
        Loop
        .Close
        '--------------------------------------------------------------------------------------------------
        
        .Open "Select CatalID,ItemRefID from ItemCatalRef Where ItemID='" & ItemCode & "'", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            If Not IsNull(![CatalID]) Then
                LV3.ListItems(![CatalID] & "'").Checked = True
                LV3.ListItems(![CatalID] & "'").SubItems(1) = ![ItemRefID] & ""
            End If
            .MoveNEXT
        Loop
        .Close
        
        .Open "Select * From VItemsComp Where ItemID='" & ItemCode & "'", con, adOpenForwardOnly, adLockReadOnly
        
        Dim ITM As ListItem
        
        Do Until .EOF
            If Not IsNull(![CompID]) Then
                Set ITM = LV1.ListItems.add(, ![CompID] & "'", ![ProcDesc] & "")
                ITM.Tag = ![ProcessID]
                ITM.SubItems(1) = ![Description] & ""
                ITM.SubItems(2) = ![Qty] & ""
                '----------------------------------
                RevEntry.CompIDs.add .Fields("CompID").Value
                '----------------------------------
            End If
            .MoveNEXT
        Loop
        
        .Close
        
        .Open "Select * From VItemsRMComp Where ItemID='" & ItemCode & "'", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            If Not IsNull(![RMID]) Then
                Set ITM = LVRM.ListItems.add(, ![RMID] & "'", ![ProcDesc] & "")
                ITM.Tag = ![ProcessID]
                ITM.SubItems(1) = ![Description] & ""
                ITM.SubItems(2) = ![Qty] & ""
                '----------------------------------
                'RevEntry.CompIDs.Add .Fields("CompID").value
                '----------------------------------
            End If
            .MoveNEXT
        Loop
        
        .Close
    End With
    
    Set rs = Nothing
    
    Call ShowShipmentInfo(ItemCode)
    
    If cmbSteel.Text = "" Then If cmbSteel.ListCount > 0 Then cmbSteel.ListIndex = 0
    If cmbItemGroup.Text = "" Then If cmbItemGroup.ListCount > 0 Then cmbItemGroup.ListIndex = 0
    If cmbUnit.Text = "" Then If cmbUnit.ListCount > 0 Then cmbUnit.Text = "Pcs"
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub ShowShipmentInfo(strItemCode As String)

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim lQtyTotal As Long, dAmtTotal As Double
    Dim bRestrictedAccess As Boolean
    If CurrentUserName <> "Administrator" Then
        bRestrictedAccess = GetSingleBooleanValue("RestrictedItemProfile", "Users", " WHERE UserName='" & CurrentUserName & "'")
    Else
        bRestrictedAccess = False
    End If
    Dim dPrice As Double, dAmt As Double
    With rs
        .Open "SELECT CustomInvoice,Qty,Price FROM VCustomInvoice WHERE CompItemCode='" & strItemCode & "'", con, adOpenForwardOnly, adLockReadOnly
        LVShipInfo.ListItems.Clear
        Do Until .EOF
            Set ITM = LVShipInfo.ListItems.add(, , LVShipInfo.ListItems.count + 1)
            ITM.SubItems(1) = !CustomInvoice & ""
            ITM.SubItems(2) = Val(!Qty & "")
            If bRestrictedAccess Then
                dPrice = 0
                dAmt = 0
            Else
                dPrice = Round(Val(!Price & ""), 4)
                dAmt = Round(Val(!Price & "") * Val(!Qty & ""), 4)
            End If
            ITM.SubItems(3) = dPrice
            
            ITM.SubItems(4) = dAmt
            lQtyTotal = lQtyTotal + Val(!Qty & "")
            dAmtTotal = dAmtTotal + dAmt
            .MoveNEXT
        Loop
        .Close
    End With
    txtTotalShipQty = lQtyTotal
    txtTotalShipValue = Format(dAmtTotal, "#,##0.00")
    
    Set rs = Nothing
    
End Sub

Private Sub txtSortNo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        LV2.SelectedItem.SubItems(1) = Val(txtSortNo)
        txtSortNo.Visible = False
    ElseIf KeyAscii = 27 Then
        txtSortNo.Visible = False
    End If
    
End Sub

Private Sub txtSortNo_LostFocus()
    txtSortNo.Visible = False
End Sub

Public Sub AddNew()
    
    If strCompany = "Weldon Instruments" Then
        Dim i As Integer
        For i = 1 To LV2.ListItems.count
            LV2.ListItems(i).Checked = True
            LV2.ListItems(i).SubItems(1) = i
        Next
        cmbCatID.Text = "Surgical Instruments"
    End If
    Me.Show 1
    
End Sub

Private Sub txtUrduFields_GotFocus(Index As Integer)
    Call SetUrduLanguage
End Sub

Private Sub txtUrduFields_LostFocus(Index As Integer)
    Call SetEnglishLanguage
End Sub
