VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmNewItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DMR (Device Master Record)"
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10710
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
   ScaleHeight     =   616
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   714
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3090
      Top             =   8490
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Please Select Picture"
      Filter          =   "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2865
      Top             =   8415
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
            Picture         =   "frmNewItem.frx":0000
            Key             =   "0'"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewItem.frx":0892
            Key             =   "7'"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewItem.frx":0CE4
            Key             =   "5'"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewItem.frx":1136
            Key             =   "6'"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewItem.frx":3470
            Key             =   "4'"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewItem.frx":414A
            Key             =   "2'"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewItem.frx":49DC
            Key             =   "3'"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewItem.frx":718E
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
      Height          =   8775
      Left            =   90
      TabIndex        =   3
      Top             =   15
      Width           =   10575
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
         Height          =   1350
         Left            =   120
         TabIndex        =   67
         Top             =   120
         Width           =   10365
         Begin MSForms.ComboBox cmbItemGroupMain 
            Height          =   315
            Left            =   1620
            TabIndex        =   243
            Tag             =   "MainGroupID"
            Top             =   945
            Width           =   8625
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "15214;556"
            MatchEntry      =   1
            ListStyle       =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Additional Group :"
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
            Index           =   84
            Left            =   60
            TabIndex        =   242
            Top             =   990
            Width           =   1500
         End
         Begin MSForms.ComboBox cmbItemType 
            Height          =   315
            Left            =   6990
            TabIndex        =   123
            Tag             =   "ItemType"
            Top             =   210
            Width           =   3255
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5741;556"
            MatchEntry      =   1
            ListStyle       =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
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
            TabIndex        =   122
            Top             =   240
            Width           =   510
         End
         Begin MSForms.ComboBox cmbItemGroup 
            Height          =   315
            Left            =   1185
            TabIndex        =   77
            Tag             =   "GroupID"
            Top             =   570
            Width           =   5175
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "9128;556"
            MatchEntry      =   1
            ListStyle       =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Item Group :"
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
            Left            =   45
            TabIndex        =   76
            Top             =   615
            Width           =   1065
         End
         Begin MSForms.ComboBox cmbCatID 
            Height          =   315
            Left            =   1185
            TabIndex        =   71
            Tag             =   "CatID"
            Top             =   210
            Width           =   5175
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "9128;556"
            MatchEntry      =   1
            ListStyle       =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
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
            Left            =   240
            TabIndex        =   70
            Top             =   270
            Width           =   870
         End
         Begin VB.Label label1 
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
            TabIndex        =   69
            Top             =   615
            Width           =   510
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   3
            Left            =   6990
            TabIndex        =   68
            Tag             =   "ItemID"
            Top             =   570
            Width           =   3255
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "5741;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   7185
         Left            =   120
         TabIndex        =   4
         Top             =   1515
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   12674
         _Version        =   393216
         Style           =   1
         Tabs            =   13
         Tab             =   7
         TabsPerRow      =   13
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
         TabPicture(0)   =   "frmNewItem.frx":7A20
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame14"
         Tab(0).Control(1)=   "Frame10"
         Tab(0).Control(2)=   "Frame4"
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Processes"
         TabPicture(1)   =   "frmNewItem.frx":7A3C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame11"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Components"
         TabPicture(2)   =   "frmNewItem.frx":7A58
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmbCompProc"
         Tab(2).Control(1)=   "cmdDelete"
         Tab(2).Control(2)=   "LV1"
         Tab(2).Control(3)=   "label1(29)"
         Tab(2).Control(4)=   "cmbComp"
         Tab(2).Control(5)=   "cmdAdd"
         Tab(2).Control(6)=   "NBox(19)"
         Tab(2).Control(7)=   "label1(31)"
         Tab(2).Control(8)=   "cmdComp"
         Tab(2).ControlCount=   9
         TabCaption(3)   =   "Weight"
         TabPicture(3)   =   "frmNewItem.frx":7A74
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame7"
         Tab(3).Control(1)=   "Frame6"
         Tab(3).Control(2)=   "Frame3"
         Tab(3).Control(3)=   "TBox(28)"
         Tab(3).Control(4)=   "label1(101)"
         Tab(3).Control(5)=   "label1(90)"
         Tab(3).Control(6)=   "TBox(22)"
         Tab(3).Control(7)=   "label1(79)"
         Tab(3).Control(8)=   "label1(80)"
         Tab(3).Control(9)=   "NBox(52)"
         Tab(3).Control(10)=   "NBox(53)"
         Tab(3).Control(11)=   "TBox(12)"
         Tab(3).Control(12)=   "label1(48)"
         Tab(3).Control(13)=   "TBox(11)"
         Tab(3).Control(14)=   "label1(47)"
         Tab(3).ControlCount=   15
         TabCaption(4)   =   "Account Info."
         TabPicture(4)   =   "frmNewItem.frx":7A90
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "label1(13)"
         Tab(4).Control(1)=   "label1(12)"
         Tab(4).Control(2)=   "label1(14)"
         Tab(4).Control(3)=   "cmbAssetAcc"
         Tab(4).Control(4)=   "cmbSaleRAcc"
         Tab(4).Control(5)=   "cmbSaleAcc"
         Tab(4).ControlCount=   6
         TabCaption(5)   =   "Price/Costing"
         TabPicture(5)   =   "frmNewItem.frx":7AAC
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "chkStamper"
         Tab(5).Control(1)=   "chkSnaffer"
         Tab(5).Control(2)=   "chkPlanter"
         Tab(5).Control(3)=   "NBox(1)"
         Tab(5).Control(4)=   "NBox(2)"
         Tab(5).Control(5)=   "NBox(3)"
         Tab(5).Control(6)=   "Frame12"
         Tab(5).ControlCount=   7
         TabCaption(6)   =   "Reference"
         TabPicture(6)   =   "frmNewItem.frx":7AC8
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "label1(64)"
         Tab(6).Control(1)=   "label1(69)"
         Tab(6).Control(2)=   "cmbForeignCatalogName"
         Tab(6).Control(3)=   "LVRef"
         Tab(6).Control(4)=   "txtRef"
         Tab(6).Control(5)=   "txtCatalogName"
         Tab(6).ControlCount=   6
         TabCaption(7)   =   "Picture"
         TabPicture(7)   =   "frmNewItem.frx":7AE4
         Tab(7).ControlEnabled=   -1  'True
         Tab(7).Control(0)=   "lblPic"
         Tab(7).Control(0).Enabled=   0   'False
         Tab(7).Control(1)=   "label1(32)"
         Tab(7).Control(1).Enabled=   0   'False
         Tab(7).Control(2)=   "Image1"
         Tab(7).Control(2).Enabled=   0   'False
         Tab(7).Control(3)=   "label1(60)"
         Tab(7).Control(3).Enabled=   0   'False
         Tab(7).Control(4)=   "ImgTechDrawing"
         Tab(7).Control(4).Enabled=   0   'False
         Tab(7).Control(5)=   "lblTechDrawing"
         Tab(7).Control(5).Enabled=   0   'False
         Tab(7).Control(6)=   "cmdViewTechDrawing"
         Tab(7).Control(6).Enabled=   0   'False
         Tab(7).Control(7)=   "cmdViewPicture"
         Tab(7).Control(7).Enabled=   0   'False
         Tab(7).Control(8)=   "cmdChangeTechDrawing"
         Tab(7).Control(8).Enabled=   0   'False
         Tab(7).Control(9)=   "cmdRemoveTechnicalDrawing"
         Tab(7).Control(9).Enabled=   0   'False
         Tab(7).ControlCount=   10
         TabCaption(8)   =   "RM Components"
         TabPicture(8)   =   "frmNewItem.frx":7B00
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "Frame13"
         Tab(8).ControlCount=   1
         TabCaption(9)   =   "Ship. Info."
         TabPicture(9)   =   "frmNewItem.frx":7B1C
         Tab(9).ControlEnabled=   0   'False
         Tab(9).Control(0)=   "Frame9"
         Tab(9).ControlCount=   1
         TabCaption(10)  =   "Set Detail"
         TabPicture(10)  =   "frmNewItem.frx":7B38
         Tab(10).ControlEnabled=   0   'False
         Tab(10).Control(0)=   "label1(70)"
         Tab(10).Control(1)=   "label1(71)"
         Tab(10).Control(2)=   "LVSet"
         Tab(10).Control(3)=   "txtSetQty"
         Tab(10).Control(4)=   "txtEditSetQty"
         Tab(10).Control(5)=   "cmbItemSet"
         Tab(10).Control(6)=   "txtTotalSetAmt"
         Tab(10).ControlCount=   7
         TabCaption(11)  =   "Look alike"
         TabPicture(11)  =   "frmNewItem.frx":7B54
         Tab(11).ControlEnabled=   0   'False
         Tab(11).Control(0)=   "Label2"
         Tab(11).Control(1)=   "Label3"
         Tab(11).Control(2)=   "LVLookalike"
         Tab(11).Control(3)=   "cmdAddLookalike"
         Tab(11).Control(4)=   "cmbItemLookalike"
         Tab(11).Control(5)=   "txtFinishedStock"
         Tab(11).ControlCount=   6
         TabCaption(12)  =   "E-Commerce"
         TabPicture(12)  =   "frmNewItem.frx":7B70
         Tab(12).ControlEnabled=   0   'False
         Tab(12).Control(0)=   "chkShowInECommerce"
         Tab(12).Control(1)=   "FraECommerce"
         Tab(12).ControlCount=   2
         Begin VB.Frame Frame14 
            Height          =   780
            Left            =   -74955
            TabIndex        =   284
            Top             =   6270
            Width           =   10200
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Scale"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   100
               Left            =   8385
               TabIndex        =   293
               Top             =   120
               Width           =   375
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   27
               Left            =   8355
               TabIndex        =   292
               Tag             =   "Scale"
               Top             =   345
               Width           =   1710
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "3016;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Fig No."
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   99
               Left            =   6735
               TabIndex        =   291
               Top             =   120
               Width           =   510
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   26
               Left            =   6705
               TabIndex        =   290
               Tag             =   "Fig_No"
               Top             =   345
               Width           =   1635
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2884;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   25
               Left            =   4845
               TabIndex        =   269
               Tag             =   "SFDA_Listing"
               Top             =   345
               Width           =   1845
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "3254;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Product Listing No."
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   96
               Left            =   4875
               TabIndex        =   287
               Top             =   120
               Width           =   1350
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   24
               Left            =   3105
               TabIndex        =   268
               Tag             =   "SFDA_No"
               Top             =   345
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
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "SFDA No."
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   95
               Left            =   3135
               TabIndex        =   286
               Top             =   135
               Width           =   690
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   23
               Left            =   75
               TabIndex        =   267
               Tag             =   "SFDA_Name"
               Top             =   345
               Width           =   3015
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "5318;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Product Name"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   94
               Left            =   120
               TabIndex        =   285
               Top             =   135
               Width           =   1005
            End
         End
         Begin VB.CommandButton cmdRemoveTechnicalDrawing 
            Caption         =   "Remove Drawing"
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
            Left            =   2010
            TabIndex        =   283
            Top             =   4800
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Frame FraECommerce 
            Enabled         =   0   'False
            Height          =   4620
            Left            =   -74910
            TabIndex        =   270
            Top             =   1080
            Width           =   10095
            Begin VB.CommandButton cmdViewActualSizeEC 
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
               Left            =   8760
               TabIndex        =   276
               Top             =   3990
               Width           =   1200
            End
            Begin VB.Label lblECommerceGroup 
               AutoSize        =   -1  'True
               Caption         =   "E-Commerce Group"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   165
               MouseIcon       =   "frmNewItem.frx":7B8C
               MousePointer    =   99  'Custom
               TabIndex        =   279
               Top             =   180
               Width           =   1380
            End
            Begin VB.Label lblPicEC 
               Alignment       =   2  'Center
               Caption         =   "Double click to Add"
               ForeColor       =   &H00FF0000&
               Height          =   435
               Left            =   7230
               TabIndex        =   278
               Top             =   1965
               Width           =   1125
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Picture"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   93
               Left            =   5670
               TabIndex        =   277
               Top             =   165
               Width           =   495
            End
            Begin VB.Image ImgEC 
               BorderStyle     =   1  'Fixed Single
               Height          =   3525
               Left            =   5670
               Stretch         =   -1  'True
               Top             =   405
               Width           =   4290
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Price"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   92
               Left            =   4650
               TabIndex        =   275
               Top             =   165
               Width           =   345
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   58
               Left            =   4650
               TabIndex        =   274
               Tag             =   "PriceForECommerce"
               Top             =   405
               Width           =   930
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "1640;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   21
               Left            =   150
               TabIndex        =   273
               Tag             =   "DescriptionForECommerce"
               Top             =   1185
               Width           =   5445
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "9604;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Description"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   91
               Left            =   180
               TabIndex        =   272
               Top             =   975
               Width           =   795
            End
            Begin MSForms.ComboBox cmbItemGroup_ECommerce 
               Height          =   315
               Left            =   150
               TabIndex        =   271
               Tag             =   "GroupIDForECommerce"
               Top             =   405
               Width           =   4485
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "7911;556"
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
         Begin VB.Frame Frame10 
            Height          =   780
            Left            =   -74130
            TabIndex        =   251
            Top             =   5475
            Width           =   8700
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Price For Cost"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   50
               Left            =   7260
               TabIndex        =   280
               Top             =   180
               Width           =   1005
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   25
               Left            =   7125
               TabIndex        =   266
               Tag             =   "PriceForCost"
               Top             =   375
               Width           =   1350
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2381;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   34
               Left            =   225
               TabIndex        =   257
               Tag             =   "MinLevel"
               Top             =   375
               Width           =   1350
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2381;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Min. Level"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   62
               Left            =   540
               TabIndex        =   256
               Top             =   165
               Width           =   720
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   35
               Left            =   1605
               TabIndex        =   259
               Tag             =   "MaxLevel"
               Top             =   375
               Width           =   1350
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2381;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Max. Level"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   63
               Left            =   1875
               TabIndex        =   255
               Top             =   165
               Width           =   780
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   12
               Left            =   2985
               TabIndex        =   261
               Tag             =   "ReorderPoint"
               Top             =   375
               Width           =   1350
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2381;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Reorder Point"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   28
               Left            =   3120
               TabIndex        =   254
               Top             =   180
               Width           =   990
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Ready Finish"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   82
               Left            =   4545
               TabIndex        =   253
               Top             =   165
               Width           =   915
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   54
               Left            =   4365
               TabIndex        =   263
               Tag             =   "ReadyFinishPrice"
               Top             =   375
               Width           =   1350
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2381;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "S.F Cost"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   83
               Left            =   6030
               TabIndex        =   252
               Top             =   165
               Width           =   615
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   55
               Left            =   5745
               TabIndex        =   265
               Tag             =   "FillingPrice"
               Top             =   375
               Width           =   1350
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2381;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
            End
         End
         Begin VB.TextBox txtTotalSetAmt 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   -66240
            Locked          =   -1  'True
            TabIndex        =   244
            Top             =   5700
            Width           =   1395
         End
         Begin VB.TextBox txtFinishedStock 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -67560
            Locked          =   -1  'True
            TabIndex        =   239
            Top             =   660
            Width           =   885
         End
         Begin ComboList.Usercontrol1 cmbItemLookalike 
            Height          =   285
            Left            =   -74850
            TabIndex        =   237
            Top             =   660
            Width           =   7275
            _ExtentX        =   12832
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
            TextBackColor   =   12640511
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
         Begin VB.CommandButton cmdAddLookalike 
            Caption         =   "Add"
            Height          =   330
            Left            =   -66660
            TabIndex        =   238
            Top             =   600
            Width           =   630
         End
         Begin VB.CommandButton cmdChangeTechDrawing 
            Caption         =   "Update Drawing"
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
            Left            =   3255
            TabIndex        =   216
            Top             =   4800
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Frame Frame13 
            Height          =   4635
            Left            =   -73740
            TabIndex        =   207
            Top             =   945
            Width           =   7770
            Begin ComboList.Usercontrol1 cmbRawMaterial 
               Height          =   285
               Left            =   2940
               TabIndex        =   208
               Top             =   420
               Width           =   3975
               _ExtentX        =   7011
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
            Begin ComboList.Usercontrol1 cmbProcRM 
               Height          =   285
               Left            =   180
               TabIndex        =   209
               Top             =   420
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
               Height          =   3810
               Left            =   165
               TabIndex        =   210
               Top             =   735
               Width           =   7470
               _ExtentX        =   13176
               _ExtentY        =   6720
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
                  Object.Width           =   4939
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Material"
                  Object.Width           =   4939
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Quantity"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Extra"
                  Object.Width           =   0
               EndProperty
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Process"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   30
               Left            =   195
               TabIndex        =   214
               Top             =   210
               Width           =   555
            End
            Begin MSForms.TextBox NBox 
               Height          =   285
               Index           =   13
               Left            =   6915
               TabIndex        =   213
               Top             =   420
               Width           =   690
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "1217;503"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Quanity"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   33
               Left            =   6930
               TabIndex        =   212
               Top             =   210
               Width           =   570
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Material"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   34
               Left            =   2940
               TabIndex        =   211
               Top             =   225
               Width           =   570
            End
         End
         Begin ComboList.Usercontrol1 cmbItemSet 
            Height          =   285
            Left            =   -74880
            TabIndex        =   200
            Top             =   660
            Width           =   9210
            _ExtentX        =   16245
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
         Begin VB.TextBox txtEditSetQty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -68850
            TabIndex        =   202
            Top             =   1920
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.TextBox txtSetQty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -65655
            TabIndex        =   201
            Top             =   660
            Width           =   825
         End
         Begin VB.TextBox txtCatalogName 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -74130
            TabIndex        =   198
            Top             =   360
            Visible         =   0   'False
            Width           =   5880
         End
         Begin VB.TextBox txtRef 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -68160
            TabIndex        =   199
            Top             =   855
            Width           =   1215
         End
         Begin VB.Frame Frame12 
            Caption         =   "Pricing :"
            Height          =   4110
            Left            =   -74940
            TabIndex        =   135
            Top             =   375
            Width           =   8940
            Begin VB.Frame Frame8 
               Caption         =   "Top :"
               Height          =   1560
               Left            =   75
               TabIndex        =   167
               Top             =   2430
               Width           =   8790
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "C && I (Air)"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   58
                  Left            =   3960
                  TabIndex        =   193
                  Top             =   225
                  Width           =   720
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   51
                  Left            =   3945
                  TabIndex        =   192
                  Tag             =   "CnIAirTop"
                  Top             =   450
                  Width           =   930
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "C && I (Sea)"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   57
                  Left            =   3000
                  TabIndex        =   191
                  Top             =   225
                  Width           =   795
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   50
                  Left            =   2985
                  TabIndex        =   190
                  Tag             =   "CnISeaTop"
                  Top             =   450
                  Width           =   930
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   49
                  Left            =   2025
                  TabIndex        =   189
                  Tag             =   "ExWorksTop"
                  Top             =   450
                  Width           =   930
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "FOB Sialkot"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   56
                  Left            =   2055
                  TabIndex        =   188
                  Top             =   225
                  Width           =   810
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "FOB Karachi"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   55
                  Left            =   1065
                  TabIndex        =   187
                  Top             =   225
                  Width           =   870
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   48
                  Left            =   1065
                  TabIndex        =   186
                  Tag             =   "FOBTop"
                  Top             =   450
                  Width           =   930
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   47
                  Left            =   4905
                  TabIndex        =   185
                  Tag             =   "CIFSeaTop"
                  Top             =   450
                  Width           =   930
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "CIF (Sea)"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   54
                  Left            =   4905
                  TabIndex        =   184
                  Top             =   240
                  Width           =   690
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   46
                  Left            =   5865
                  TabIndex        =   183
                  Tag             =   "CIFAirTop"
                  Top             =   450
                  Width           =   930
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "CIF (Air)"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   53
                  Left            =   5865
                  TabIndex        =   182
                  Top             =   240
                  Width           =   615
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   45
                  Left            =   6825
                  TabIndex        =   181
                  Tag             =   "CnFSeaTop"
                  Top             =   450
                  Width           =   930
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "C & F (Sea)"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   52
                  Left            =   6840
                  TabIndex        =   180
                  Top             =   240
                  UseMnemonic     =   0   'False
                  Width           =   825
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   44
                  Left            =   7785
                  TabIndex        =   179
                  Tag             =   "CnFAirTop"
                  Top             =   450
                  Width           =   930
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "C & F (Air)"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   51
                  Left            =   7800
                  TabIndex        =   178
                  Top             =   255
                  UseMnemonic     =   0   'False
                  Width           =   750
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   43
                  Left            =   2025
                  TabIndex        =   177
                  Top             =   1020
                  Width           =   930
                  VariousPropertyBits=   679495711
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   42
                  Left            =   2985
                  TabIndex        =   176
                  Top             =   1020
                  Width           =   930
                  VariousPropertyBits=   679495711
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   41
                  Left            =   3945
                  TabIndex        =   175
                  Top             =   1020
                  Width           =   930
                  VariousPropertyBits=   679495711
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   40
                  Left            =   7785
                  TabIndex        =   174
                  Top             =   1020
                  Width           =   930
                  VariousPropertyBits=   679495711
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   39
                  Left            =   6825
                  TabIndex        =   173
                  Top             =   1020
                  Width           =   930
                  VariousPropertyBits=   679495711
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   38
                  Left            =   5865
                  TabIndex        =   172
                  Top             =   1020
                  Width           =   930
                  VariousPropertyBits=   679495711
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   37
                  Left            =   4905
                  TabIndex        =   171
                  Top             =   1020
                  Width           =   930
                  VariousPropertyBits=   679495711
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   36
                  Left            =   1065
                  TabIndex        =   170
                  Top             =   1020
                  Width           =   930
                  VariousPropertyBits=   679495711
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin VB.Label lblMainCurr 
                  AutoSize        =   -1  'True
                  Caption         =   "Euro"
                  BeginProperty Font 
                     Name            =   "Trebuchet MS"
                     Size            =   15.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   405
                  Index           =   1
                  Left            =   240
                  TabIndex        =   169
                  Top             =   390
                  Width           =   675
               End
               Begin VB.Label lblSubCurr 
                  AutoSize        =   -1  'True
                  Caption         =   "US $"
                  BeginProperty Font 
                     Name            =   "Trebuchet MS"
                     Size            =   15.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   405
                  Index           =   1
                  Left            =   240
                  TabIndex        =   168
                  Top             =   960
                  Width           =   645
               End
            End
            Begin VB.Frame Frame5 
               Caption         =   "Regular :"
               Height          =   1560
               Left            =   75
               TabIndex        =   136
               Top             =   795
               Width           =   8790
               Begin VB.Label lblSubCurr 
                  AutoSize        =   -1  'True
                  Caption         =   "US $"
                  BeginProperty Font 
                     Name            =   "Trebuchet MS"
                     Size            =   15.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   405
                  Index           =   0
                  Left            =   240
                  TabIndex        =   166
                  Top             =   960
                  Width           =   645
               End
               Begin VB.Label lblMainCurr 
                  AutoSize        =   -1  'True
                  Caption         =   "Euro"
                  BeginProperty Font 
                     Name            =   "Trebuchet MS"
                     Size            =   15.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   405
                  Index           =   0
                  Left            =   240
                  TabIndex        =   165
                  Top             =   375
                  Width           =   675
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   29
                  Left            =   1065
                  TabIndex        =   160
                  Top             =   1020
                  Width           =   930
                  VariousPropertyBits=   679495711
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   30
                  Left            =   4905
                  TabIndex        =   159
                  Top             =   1020
                  Width           =   930
                  VariousPropertyBits=   679495711
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   31
                  Left            =   5865
                  TabIndex        =   158
                  Top             =   1020
                  Width           =   930
                  VariousPropertyBits=   679495711
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   32
                  Left            =   6825
                  TabIndex        =   157
                  Top             =   1020
                  Width           =   930
                  VariousPropertyBits=   679495711
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   33
                  Left            =   7785
                  TabIndex        =   156
                  Top             =   1020
                  Width           =   930
                  VariousPropertyBits=   679495711
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   26
                  Left            =   3945
                  TabIndex        =   155
                  Top             =   1020
                  Width           =   930
                  VariousPropertyBits=   679495711
                  BorderStyle     =   1
                  Size            =   "1640;556"
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
                  Left            =   2985
                  TabIndex        =   154
                  Top             =   1020
                  Width           =   930
                  VariousPropertyBits=   679495711
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   28
                  Left            =   2025
                  TabIndex        =   153
                  Top             =   1020
                  Width           =   930
                  VariousPropertyBits=   679495711
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "C & F (Air)"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   23
                  Left            =   7800
                  TabIndex        =   152
                  Top             =   255
                  UseMnemonic     =   0   'False
                  Width           =   750
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   11
                  Left            =   7785
                  TabIndex        =   151
                  Tag             =   "CnFAir"
                  Top             =   450
                  Width           =   930
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "C & F (Sea)"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   22
                  Left            =   6840
                  TabIndex        =   150
                  Top             =   240
                  UseMnemonic     =   0   'False
                  Width           =   825
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   10
                  Left            =   6825
                  TabIndex        =   149
                  Tag             =   "CnFSea"
                  Top             =   450
                  Width           =   930
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "CIF (Air)"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   21
                  Left            =   5865
                  TabIndex        =   148
                  Top             =   240
                  Width           =   615
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   9
                  Left            =   5865
                  TabIndex        =   147
                  Tag             =   "CIFAir"
                  Top             =   450
                  Width           =   930
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "CIF (Sea)"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   20
                  Left            =   4905
                  TabIndex        =   146
                  Top             =   240
                  Width           =   690
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   8
                  Left            =   4905
                  TabIndex        =   145
                  Tag             =   "CIFSea"
                  Top             =   450
                  Width           =   930
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   0
                  Left            =   1065
                  TabIndex        =   144
                  Tag             =   "FOB"
                  Top             =   450
                  Width           =   930
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "FOB Karachi"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   17
                  Left            =   1065
                  TabIndex        =   143
                  Top             =   225
                  Width           =   870
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "FOB Sialkot"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   37
                  Left            =   2055
                  TabIndex        =   142
                  Top             =   225
                  Width           =   810
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   14
                  Left            =   2040
                  TabIndex        =   141
                  Tag             =   "ExWorks"
                  Top             =   450
                  Width           =   930
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   15
                  Left            =   2985
                  TabIndex        =   140
                  Tag             =   "CnISea"
                  Top             =   450
                  Width           =   930
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "C && I (Sea)"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   38
                  Left            =   3000
                  TabIndex        =   139
                  Top             =   225
                  Width           =   795
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   16
                  Left            =   3945
                  TabIndex        =   138
                  Tag             =   "CnIAir"
                  Top             =   450
                  Width           =   930
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1640;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "C && I (Air)"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   39
                  Left            =   3960
                  TabIndex        =   137
                  Top             =   225
                  Width           =   720
               End
            End
            Begin VB.Label lblConversion 
               AutoSize        =   -1  'True
               Caption         =   "Euro to US$"
               ForeColor       =   &H00400000&
               Height          =   195
               Left            =   1065
               TabIndex        =   164
               Top             =   270
               Width           =   855
            End
            Begin MSForms.TextBox txtConversion 
               Height          =   315
               Left            =   1035
               TabIndex        =   163
               Top             =   480
               Width           =   930
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "1640;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin MSForms.ComboBox cmbCurr 
               Height          =   315
               Left            =   75
               TabIndex        =   162
               Tag             =   "ItemCurr"
               Top             =   480
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
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Currency"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   19
               Left            =   105
               TabIndex        =   161
               Top             =   270
               Width           =   660
            End
         End
         Begin VB.Frame Frame11 
            Height          =   5895
            Left            =   -74940
            TabIndex        =   124
            Top             =   360
            Width           =   10230
            Begin VB.CommandButton cmdUpdateForAllItems 
               Caption         =   "Update All Items"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1365
               TabIndex        =   262
               ToolTipText     =   "Update All Items of the Selected Catagory with selected Processes"
               Top             =   5460
               Visible         =   0   'False
               Width           =   1905
            End
            Begin VB.TextBox txtReOrdLvl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6405
               TabIndex        =   247
               Top             =   4290
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.TextBox txtMaxLvl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5715
               TabIndex        =   246
               Top             =   4290
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.TextBox txtMinLvl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5025
               TabIndex        =   245
               Top             =   4290
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.CommandButton cmdDeleteAll 
               Caption         =   "Delete All"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   45
               TabIndex        =   235
               Top             =   5460
               Width           =   1230
            End
            Begin VB.CommandButton cmdBreakLink 
               Caption         =   "Break Link"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   8910
               TabIndex        =   234
               Top             =   5430
               Width           =   1230
            End
            Begin VB.TextBox txtProcessFamily 
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   4800
               Locked          =   -1  'True
               TabIndex        =   233
               Top             =   5430
               Width           =   4035
            End
            Begin MSComctlLib.ListView LV2Hidden 
               Height          =   3735
               Left            =   420
               TabIndex        =   226
               Top             =   1485
               Visible         =   0   'False
               Width           =   9390
               _ExtentX        =   16563
               _ExtentY        =   6588
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
                  Object.Width           =   10583
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   1
                  Text            =   "Scanning"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Sort No."
                  Object.Width           =   2822
               EndProperty
            End
            Begin VB.TextBox txtStRate 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   4140
               TabIndex        =   231
               Top             =   1170
               Visible         =   0   'False
               Width           =   1000
            End
            Begin VB.ComboBox cmbScanningEdit 
               Height          =   315
               Left            =   3285
               TabIndex        =   217
               Text            =   "Combo1"
               Top             =   2865
               Visible         =   0   'False
               Width           =   1875
            End
            Begin ComboList.Usercontrol1 cmbProcess 
               Height          =   285
               Left            =   4050
               TabIndex        =   128
               Top             =   390
               Width           =   5400
               _ExtentX        =   9525
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
            Begin ComboList.Usercontrol1 cmbProcessGroup 
               Height          =   285
               Left            =   45
               TabIndex        =   127
               Top             =   390
               Width           =   3990
               _ExtentX        =   7038
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
            Begin VB.TextBox txtProcessSortNo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   9465
               TabIndex        =   130
               Top             =   390
               Width           =   675
            End
            Begin VB.TextBox txtProcRate 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3060
               TabIndex        =   126
               Top             =   1695
               Visible         =   0   'False
               Width           =   2730
            End
            Begin VB.TextBox txtSortNo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3120
               TabIndex        =   125
               Top             =   1170
               Visible         =   0   'False
               Width           =   1000
            End
            Begin MSComctlLib.ListView LV2 
               Height          =   4680
               Left            =   45
               TabIndex        =   131
               Top             =   690
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   8255
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
               NumItems        =   7
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Process"
                  Object.Width           =   15293
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   1
                  Text            =   "Scanning"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   2
                  Text            =   "Sort No."
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   3
                  Text            =   "St. Rate"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   4
                  Text            =   "Min."
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   5
                  Text            =   "Max."
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   6
                  Text            =   "Re-Ord."
                  Object.Width           =   0
               EndProperty
            End
            Begin VB.Label label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Re-Ord."
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   87
               Left            =   6420
               TabIndex        =   250
               Top             =   4095
               Width           =   600
            End
            Begin VB.Label label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Max."
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   86
               Left            =   5820
               TabIndex        =   249
               Top             =   4095
               Width           =   360
            End
            Begin VB.Label label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Min."
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   85
               Left            =   5160
               TabIndex        =   248
               Top             =   4095
               Width           =   300
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Process Family :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   81
               Left            =   3345
               TabIndex        =   232
               Top             =   5505
               Width           =   1350
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Scanning :"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   72
               Left            =   3465
               TabIndex        =   206
               Top             =   4095
               Width           =   750
            End
            Begin MSForms.ComboBox cmbScanning 
               Height          =   285
               Left            =   3435
               TabIndex        =   129
               Top             =   4290
               Width           =   870
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "1535;503"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   178
               FontPitchAndFamily=   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Seq. No."
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   68
               Left            =   9465
               TabIndex        =   134
               Top             =   195
               Width           =   630
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Process Group :"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   67
               Left            =   45
               TabIndex        =   133
               Top             =   195
               Width           =   1140
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Process :"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   66
               Left            =   4065
               TabIndex        =   132
               Top             =   195
               Width           =   660
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
            Left            =   7560
            TabIndex        =   118
            Top             =   4815
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
            Left            =   135
            TabIndex        =   117
            Top             =   4785
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Frame Frame9 
            Caption         =   "Shipment Hisotry :"
            Height          =   4605
            Left            =   -74175
            TabIndex        =   112
            Top             =   840
            Width           =   8685
            Begin VB.TextBox txtTotalShipValue 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   5310
               Locked          =   -1  'True
               TabIndex        =   115
               Top             =   4125
               Width           =   1230
            End
            Begin VB.TextBox txtTotalShipQty 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2925
               Locked          =   -1  'True
               TabIndex        =   114
               Top             =   4125
               Width           =   1230
            End
            Begin MSComctlLib.ListView LVShipInfo 
               Height          =   3765
               Left            =   120
               TabIndex        =   113
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
            TabIndex        =   101
            Top             =   3975
            Width           =   7020
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "No. of Pcs"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   89
               Left            =   5235
               TabIndex        =   260
               Top             =   240
               Width           =   735
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   57
               Left            =   5220
               TabIndex        =   98
               Tag             =   "SmallBoxPcs"
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
               ParagraphAlign  =   3
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Length (cm)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   46
               Left            =   105
               TabIndex        =   104
               Top             =   255
               Width           =   855
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   23
               Left            =   75
               TabIndex        =   95
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
               ParagraphAlign  =   3
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   22
               Left            =   1830
               TabIndex        =   96
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
               ParagraphAlign  =   3
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Height (cm)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   44
               Left            =   1860
               TabIndex        =   103
               Top             =   255
               Width           =   825
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   20
               Left            =   3525
               TabIndex        =   97
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
               ParagraphAlign  =   3
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Width (cm)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   41
               Left            =   3570
               TabIndex        =   102
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
            TabIndex        =   89
            Top             =   3015
            Width           =   7020
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   56
               Left            =   5220
               TabIndex        =   94
               Tag             =   "MasterCartonSmallBoxes"
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
               ParagraphAlign  =   3
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "No. of Small Boxes"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   88
               Left            =   5265
               TabIndex        =   258
               Top             =   255
               Width           =   1335
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Width (cm)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   45
               Left            =   3570
               TabIndex        =   100
               Top             =   255
               Width           =   780
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   21
               Left            =   3525
               TabIndex        =   93
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
               ParagraphAlign  =   3
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Height (cm)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   43
               Left            =   1860
               TabIndex        =   99
               Top             =   255
               Width           =   825
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   18
               Left            =   1830
               TabIndex        =   92
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
               ParagraphAlign  =   3
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   17
               Left            =   75
               TabIndex        =   91
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
               ParagraphAlign  =   3
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Length (cm)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   42
               Left            =   105
               TabIndex        =   90
               Top             =   255
               Width           =   855
            End
         End
         Begin ComboList.Usercontrol1 cmbCompProc 
            Height          =   285
            Left            =   -73875
            TabIndex        =   86
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
            Picture         =   "frmNewItem.frx":7CDE
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Remove This Entry From Voucher"
            Top             =   1920
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   540
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
            TabIndex        =   21
            Top             =   795
            Width           =   3630
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Fixed Packing Unit"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   49
               Left            =   1935
               TabIndex        =   110
               Top             =   1455
               Width           =   1305
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   24
               Left            =   1890
               TabIndex        =   109
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
               ParagraphAlign  =   3
            End
            Begin MSForms.ComboBox cmbWeightUnit 
               Height          =   315
               Left            =   75
               TabIndex        =   51
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
               ParagraphAlign  =   3
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Weighing Unit"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   18
               Left            =   105
               TabIndex        =   50
               Top             =   225
               Width           =   990
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Forging Weight"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   10
               Left            =   1935
               TabIndex        =   39
               Top             =   225
               Width           =   1095
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   4
               Left            =   1890
               TabIndex        =   37
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
               ParagraphAlign  =   3
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   5
               Left            =   75
               TabIndex        =   35
               Tag             =   "WasteVisible"
               Top             =   1020
               Width           =   1740
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "3069;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Wastage (Visible)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   9
               Left            =   90
               TabIndex        =   33
               Top             =   780
               Width           =   1245
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   6
               Left            =   1890
               TabIndex        =   30
               Top             =   1020
               Width           =   1635
               VariousPropertyBits=   679495711
               BorderStyle     =   1
               Size            =   "2884;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Wastage (Invisible)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   8
               Left            =   1935
               TabIndex        =   28
               Top             =   810
               Width           =   1395
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   7
               Left            =   75
               TabIndex        =   26
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
               ParagraphAlign  =   3
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Finished Weight"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   7
               Left            =   120
               TabIndex        =   24
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
            Height          =   5055
            Left            =   -74130
            TabIndex        =   5
            Top             =   435
            Width           =   8700
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
               TabIndex        =   44
               Top             =   3885
               Width           =   8475
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Maker Description"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   102
               Left            =   135
               TabIndex        =   296
               Top             =   750
               Width           =   1275
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   29
               Left            =   105
               TabIndex        =   17
               Tag             =   "MakerDescription"
               Top             =   960
               Width           =   4245
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "7488;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   60
               Left            =   5430
               TabIndex        =   42
               Tag             =   "HRC_To"
               Top             =   3225
               Width           =   1065
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "1879;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   59
               Left            =   4350
               TabIndex        =   41
               Tag             =   "HRC_From"
               Top             =   3225
               Width           =   1065
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "1879;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "HRC To"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   98
               Left            =   5400
               TabIndex        =   289
               Top             =   3015
               Width           =   540
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "HRC From"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   97
               Left            =   4350
               TabIndex        =   288
               Top             =   3015
               Width           =   720
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "FDA510(K)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   59
               Left            =   6900
               TabIndex        =   225
               Top             =   4380
               Width           =   780
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   20
               Left            =   6900
               TabIndex        =   49
               Tag             =   "FDA510K"
               Top             =   4590
               Width           =   1680
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2963;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Euro. Reg. No."
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   78
               Left            =   5190
               TabIndex        =   224
               Top             =   4395
               Width           =   1080
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   19
               Left            =   5190
               TabIndex        =   48
               Tag             =   "EuropeanRegNo"
               Top             =   4590
               Width           =   1650
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2910;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "FDA Prod. Code"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   77
               Left            =   3480
               TabIndex        =   223
               Top             =   4395
               Width           =   1155
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   18
               Left            =   3480
               TabIndex        =   47
               Tag             =   "FDAProductCode"
               Top             =   4590
               Width           =   1650
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2910;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "FDA Listing No."
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   76
               Left            =   1785
               TabIndex        =   222
               Top             =   4380
               Width           =   1095
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   17
               Left            =   1785
               TabIndex        =   46
               Tag             =   "FDAListingNo"
               Top             =   4590
               Width           =   1650
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2910;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "UMDNS Code"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   75
               Left            =   105
               TabIndex        =   221
               Top             =   4380
               Width           =   945
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   16
               Left            =   75
               TabIndex        =   45
               Tag             =   "UMDNSCode"
               Top             =   4590
               Width           =   1650
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2910;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Custom Description"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   74
               Left            =   1125
               TabIndex        =   220
               Top             =   3630
               Visible         =   0   'False
               Width           =   1380
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   15
               Left            =   105
               TabIndex        =   219
               Tag             =   "Description"
               Top             =   3840
               Visible         =   0   'False
               Width           =   8475
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "14949;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   14
               Left            =   5850
               TabIndex        =   20
               Tag             =   "GTINBarcodeNo"
               Top             =   960
               Width           =   2730
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "4815;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "GTIN Barcode No."
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   73
               Left            =   5880
               TabIndex        =   218
               Top             =   735
               Width           =   1290
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Type"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   11
               Left            =   6510
               TabIndex        =   121
               Top             =   1875
               Width           =   360
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Color"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   40
               Left            =   4320
               TabIndex        =   120
               Top             =   1875
               Width           =   375
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Tip Size"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   61
               Left            =   4095
               TabIndex        =   119
               Top             =   1335
               Width           =   540
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   13
               Left            =   4095
               TabIndex        =   23
               Tag             =   "TipSize"
               Top             =   1545
               Width           =   1620
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2857;556"
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
               TabIndex        =   111
               Top             =   3660
               Width           =   900
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   10
               Left            =   4350
               TabIndex        =   32
               Tag             =   "ItemColor"
               Top             =   2085
               Width           =   2130
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "3757;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   9
               Left            =   4365
               TabIndex        =   16
               Tag             =   "CustomDescription"
               Top             =   390
               Width           =   4215
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "7435;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Custom Description"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   36
               Left            =   4365
               TabIndex        =   88
               Top             =   210
               Width           =   1380
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   8
               Left            =   2220
               TabIndex        =   31
               Tag             =   "EAN128"
               Top             =   2085
               Width           =   2115
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "3731;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Catalog Page #"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   35
               Left            =   2220
               TabIndex        =   87
               Top             =   1875
               Width           =   1125
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Finishing Quality"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   27
               Left            =   6510
               TabIndex        =   78
               Top             =   3000
               Width           =   1170
            End
            Begin MSForms.ComboBox cmbFinQuality 
               Height          =   315
               Left            =   6510
               TabIndex        =   43
               Tag             =   "FinQuality"
               Top             =   3225
               Width           =   2070
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "3651;556"
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
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Actual Description"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   1
               Left            =   135
               TabIndex        =   19
               Top             =   180
               Width           =   1290
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   0
               Left            =   105
               TabIndex        =   15
               Tag             =   "ItemName"
               Top             =   390
               Width           =   4245
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "7488;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbType 
               Height          =   315
               Left            =   6510
               TabIndex        =   34
               Tag             =   "Type"
               Top             =   2085
               Width           =   2070
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "3651;556"
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
               Left            =   5730
               TabIndex        =   25
               Tag             =   "ItemSize"
               Top             =   1545
               Width           =   1470
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2593;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Size"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   15
               Left            =   5730
               TabIndex        =   14
               Top             =   1335
               Width           =   285
            End
            Begin MSForms.ComboBox cmbUnit 
               Height          =   315
               Left            =   4365
               TabIndex        =   18
               Tag             =   "Unit"
               Top             =   960
               Width           =   1470
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "2593;556"
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
               TabIndex        =   22
               Tag             =   "ItemUsage"
               Top             =   1545
               Width           =   3975
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "7011;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Special Instructions"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   5
               Left            =   135
               TabIndex        =   13
               Top             =   3015
               Width           =   1395
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   2
               Left            =   105
               TabIndex        =   40
               Tag             =   "PacknLabel"
               Top             =   3225
               Width           =   4230
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "7461;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Steel / Material Used"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   3
               Left            =   135
               TabIndex        =   12
               Top             =   2445
               Width           =   1485
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Internal Catalog Ref."
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   2
               Left            =   135
               TabIndex        =   11
               Top             =   1875
               Width           =   1530
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   1
               Left            =   120
               TabIndex        =   29
               Tag             =   "Attributes"
               Top             =   2085
               Width           =   2070
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "3651;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Item Usage"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   0
               Left            =   135
               TabIndex        =   10
               Top             =   1335
               Width           =   825
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Unit"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   6
               Left            =   4410
               TabIndex        =   9
               Top             =   750
               Width           =   285
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Size Unit"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   24
               Left            =   7215
               TabIndex        =   8
               Top             =   1335
               Width           =   615
            End
            Begin MSForms.ComboBox cmbSizeUnit 
               Height          =   315
               Left            =   7215
               TabIndex        =   27
               Tag             =   "SizeUnit"
               Top             =   1545
               Width           =   1365
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "2408;556"
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
               TabIndex        =   36
               Tag             =   "SteelUsed"
               Top             =   2640
               Width           =   4230
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "7461;556"
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
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Gauge"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   25
               Left            =   4350
               TabIndex        =   7
               Top             =   2445
               Width           =   465
            End
            Begin MSForms.ComboBox cmbGuage 
               Height          =   315
               Left            =   4350
               TabIndex        =   38
               Tag             =   "Gage"
               Top             =   2640
               Width           =   4230
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "7461;556"
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
               TabIndex        =   6
               Tag             =   "CreateDT"
               Top             =   3495
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
         Begin ComboList.Usercontrol1 cmbSaleAcc 
            Height          =   285
            Left            =   -74835
            TabIndex        =   52
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
            TabIndex        =   53
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
            TabIndex        =   54
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
            TabIndex        =   56
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
         Begin MSComctlLib.ListView LVRef 
            Height          =   4110
            Left            =   -74070
            TabIndex        =   194
            Top             =   1170
            Width           =   7125
            _ExtentX        =   12568
            _ExtentY        =   7250
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
            BackColor       =   16777215
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
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Catalog Name"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Reference"
               Object.Width           =   2117
            EndProperty
         End
         Begin MSComctlLib.ListView LVSet 
            Height          =   4755
            Left            =   -74880
            TabIndex        =   203
            Top             =   960
            Width           =   10035
            _ExtentX        =   17701
            _ExtentY        =   8387
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
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S. #"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Item Code"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Item Name"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Size"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Qty"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Price"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "Amount"
               Object.Width           =   1764
            EndProperty
         End
         Begin MSComctlLib.ListView LVLookalike 
            Height          =   5250
            Left            =   -74850
            TabIndex        =   236
            Top             =   930
            Width           =   8835
            _ExtentX        =   15584
            _ExtentY        =   9260
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
            BackColor       =   16777215
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
               Object.Width           =   1147
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Item"
               Object.Width           =   12700
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Stock"
               Object.Width           =   2117
            EndProperty
         End
         Begin MSForms.TextBox TBox 
            Height          =   510
            Index           =   28
            Left            =   -74280
            TabIndex        =   295
            Tag             =   "Additional_Detail"
            Top             =   6555
            Width           =   7005
            VariousPropertyBits=   -1467987941
            BorderStyle     =   1
            Size            =   "12356;900"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Additional Detail"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   101
            Left            =   -74265
            TabIndex        =   294
            Top             =   6315
            Width           =   1155
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "PO Instructions :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   90
            Left            =   -72570
            TabIndex        =   282
            Top             =   5025
            Width           =   1215
         End
         Begin MSForms.TextBox TBox 
            Height          =   975
            Index           =   22
            Left            =   -72585
            TabIndex        =   281
            Tag             =   "POInustructions"
            Top             =   5265
            Width           =   5280
            VariousPropertyBits=   -1467987941
            BorderStyle     =   1
            Size            =   "9313;1720"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkShowInECommerce 
            Height          =   345
            Left            =   -74730
            TabIndex        =   264
            Tag             =   "AvailableForECommerce"
            Top             =   795
            Width           =   2625
            BackColor       =   -2147483633
            ForeColor       =   8388608
            DisplayStyle    =   4
            Size            =   "4630;609"
            Value           =   "0"
            Caption         =   "Available for E-Commerce"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Stock"
            Height          =   195
            Left            =   -67560
            TabIndex        =   241
            Top             =   465
            Width           =   390
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Item"
            Height          =   195
            Left            =   -74850
            TabIndex        =   240
            Top             =   435
            Width           =   330
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Max. Lot Size"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   79
            Left            =   -74250
            TabIndex        =   230
            Top             =   5055
            Width           =   960
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Lot Size (Buffer)"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   80
            Left            =   -74250
            TabIndex        =   229
            Top             =   5715
            Width           =   1170
         End
         Begin MSForms.TextBox NBox 
            Height          =   315
            Index           =   52
            Left            =   -74295
            TabIndex        =   228
            Tag             =   "ItemMaxLotSize"
            Top             =   5265
            Width           =   1650
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "2910;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.TextBox NBox 
            Height          =   315
            Index           =   53
            Left            =   -74295
            TabIndex        =   227
            Tag             =   "ItemLotSizeBuffer"
            Top             =   5910
            Width           =   1650
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "2910;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.ComboBox cmbForeignCatalogName 
            Height          =   285
            Left            =   -74070
            TabIndex        =   197
            Top             =   855
            Width           =   5880
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "10372;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblTechDrawing 
            Alignment       =   2  'Center
            Caption         =   "Click to Add PDF"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   285
            MouseIcon       =   "frmNewItem.frx":8210
            MousePointer    =   99  'Custom
            TabIndex        =   215
            Top             =   2955
            Width           =   4005
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Item"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   71
            Left            =   -74880
            TabIndex        =   205
            Top             =   465
            Width           =   330
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Qty"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   70
            Left            =   -65655
            TabIndex        =   204
            Top             =   465
            Width           =   270
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Catalog Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   69
            Left            =   -74070
            MouseIcon       =   "frmNewItem.frx":8362
            MousePointer    =   99  'Custom
            TabIndex        =   196
            Top             =   630
            Width           =   1005
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Reference"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   64
            Left            =   -68160
            TabIndex        =   195
            Top             =   660
            Width           =   750
         End
         Begin VB.Image ImgTechDrawing 
            BorderStyle     =   1  'Fixed Single
            Height          =   3525
            Left            =   150
            Stretch         =   -1  'True
            Top             =   1230
            Width           =   4290
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Tech. Drawing"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   60
            Left            =   180
            TabIndex        =   116
            Top             =   1035
            Width           =   1035
         End
         Begin MSForms.TextBox TBox 
            Height          =   630
            Index           =   12
            Left            =   -70605
            TabIndex        =   108
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
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Polybag :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   48
            Left            =   -70590
            TabIndex        =   107
            Top             =   2085
            Width           =   675
         End
         Begin MSForms.TextBox TBox 
            Height          =   1095
            Index           =   11
            Left            =   -70605
            TabIndex        =   106
            Tag             =   "PackingInstructions"
            Top             =   945
            Width           =   4050
            VariousPropertyBits=   -1467987941
            BorderStyle     =   1
            Size            =   "7144;1931"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Packing Instructions"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   47
            Left            =   -70575
            TabIndex        =   105
            Top             =   720
            Width           =   1440
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Process"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   29
            Left            =   -73860
            TabIndex        =   85
            Top             =   1365
            Width           =   555
         End
         Begin MSForms.TextBox NBox 
            Height          =   315
            Index           =   3
            Left            =   -71340
            TabIndex        =   84
            Tag             =   "SnaffRate"
            Top             =   945
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
            Left            =   -71340
            TabIndex        =   83
            Tag             =   "StampRate"
            Top             =   480
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
            Left            =   -68355
            TabIndex        =   82
            Tag             =   "PlantRate"
            Top             =   495
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
            Left            =   -69780
            TabIndex        =   81
            Top             =   480
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
            Left            =   -72765
            TabIndex        =   80
            Top             =   930
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
            Left            =   -72765
            TabIndex        =   79
            Top             =   435
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
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Asset Account"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   14
            Left            =   -74820
            TabIndex        =   66
            Top             =   2160
            Width           =   1035
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Sale Return Account"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   12
            Left            =   -74805
            TabIndex        =   65
            Top             =   1590
            Width           =   1470
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Sale Account"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   13
            Left            =   -74805
            TabIndex        =   64
            Top             =   1050
            Width           =   930
         End
         Begin MSForms.ComboBox cmbComp 
            Height          =   285
            Left            =   -71115
            TabIndex        =   63
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
            TabIndex        =   62
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
            TabIndex        =   61
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
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Quanity"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   31
            Left            =   -68310
            TabIndex        =   60
            Top             =   1365
            Width           =   570
         End
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   3525
            Left            =   4470
            Stretch         =   -1  'True
            Top             =   1230
            Width           =   4290
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Picture"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   32
            Left            =   4470
            TabIndex        =   59
            Top             =   1050
            Width           =   495
         End
         Begin VB.Label lblPic 
            Alignment       =   2  'Center
            Caption         =   "Double click to Add"
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   6030
            TabIndex        =   58
            Top             =   2790
            Width           =   1125
         End
         Begin MSForms.CommandButton cmdComp 
            Height          =   390
            Left            =   -71115
            TabIndex        =   57
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
         TabIndex        =   72
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
      Height          =   465
      Left            =   2745
      TabIndex        =   74
      Top             =   8730
      Visible         =   0   'False
      Width           =   3015
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
         TabIndex        =   75
         Tag             =   "RevID"
         Top             =   240
         Width           =   60
      End
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   420
      Left            =   9045
      TabIndex        =   73
      Top             =   8790
      Width           =   1575
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "2778;741"
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
      Top             =   8850
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
      Height          =   420
      Left            =   7440
      TabIndex        =   1
      Top             =   8790
      Width           =   1575
      Caption         =   " Save & Close"
      PicturePosition =   327683
      Size            =   "2778;741"
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
      Height          =   420
      Left            =   5835
      TabIndex        =   0
      Top             =   8790
      Width           =   1575
      Caption         =   "  Save & New"
      PicturePosition =   327683
      Size            =   "2778;741"
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
Attribute VB_Name = "frmNewItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ItemCode As String, strNewItemCode As String
'Dim CopyItemCode As String
Dim PicFileName As String, TechDrawingFileName As String, strPicECFileName As String
Dim SelItm As ListItem
Dim RevEntry As New DMRRev '
Dim DoPrint As Boolean
Dim strDelListSet As String
Dim strDelListRM As String
Dim strDelListSF As String
Dim iColNoProcess As Integer
Dim strDelListItemProcs As String
Dim strCopyFrom As String
Dim strDelListLookAlikes As String
Dim bSample As Boolean
Dim bDataLoading As Boolean
Dim bDelAllProesses As Boolean

Private Sub AddTechDrawing()

    On Error Resume Next
    
    TechDrawingFileName = ""
    CD1.Filter = "PDF Files Only|*.pdf"
    CD1.ShowSave
    If Not (err.Number = 32755) Then
        TechDrawingFileName = CD1.FileName
        'ImgTechDrawing.Picture = LoadPicture(TechDrawingFileName)
        lblTechDrawing.Caption = CD1.FileTitle
        lblTechDrawing.Tag = CD1.FileName
    Else
        TechDrawingFileName = ""
        imgttechdrawing.Picture = LoadPicture(TechDrawingFileName)
        lblTechDrawing.Visible = True
    End If
    
End Sub


Private Sub AddPic()
    On Error Resume Next
    
    PicFileName = ""
    CD1.Filter = "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
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

Private Sub AddPicEC()
    On Error Resume Next
    
    strPicECFileName = ""
    CD1.Filter = "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
    CD1.ShowSave
    If Not (err.Number = 32755) Then
        strPicECFileName = CD1.FileName
        ImgEC.Picture = LoadPicture(strPicECFileName)
        lblPicEC.Visible = False
    Else
        strPicECFileName = ""
        ImgEC.Picture = LoadPicture(strPicECFileName)
        lblPicEC.Visible = True
    End If
    
End Sub


Public Sub EditItem(strItemCode As String)

    On Error GoTo err
    'Dim rs As New ADODB.Recordset
    ItemCode = ""
    ItemCode = Left(strItemCode, Len(strItemCode) - 1)
   
    Dim bRestrictEditing As Boolean
    If CurrentUserName <> "Administrator" Then
        bRestrictEditing = GetSingleBooleanValue("RestrictCompanyCatalogEditing", "Users", " WHERE UserName='" & CurrentUserName & "'")
    Else
        bRestrictEditing = False
    End If
    If bRestrictEditing Then
        cmdOk.Enabled = False
        cmdNext.Enabled = False
    End If
    
    'cmbCatID.Locked = True
'    If strCompany = "Sunlike" Or strCompany = "QEL" Or strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
'        TBox(3).Locked = False
'    Else
'        TBox(3).Locked = True
'    End If
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
'    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
'    With rs
'        .Open "Select * From ForeignCatalog Order By CatalogName", con, adOpenForwardOnly, adLockReadOnly
'        Do Until .EOF
'            Set ITM = LV3.ListItems.add(, ![CatalID] & "'", ![CatalogName])
'            ITM.SubItems(1) = ""
'            .MoveNEXT
'        Loop
'        .Close
'    End With
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub FillCmbs()

    On Error GoTo 0
    'Call AddToCombo(cmbPchAcc, "AccTitle", "VActiveAccounts", " Where AccNo Like '" & PchAcc & "%' AND Parent=0 Order By ACCTitle")
    Call cmbSaleAcc.AddVals(con, "AccTitle", "VActiveAccounts", "AccNo", " Where AccNo Like '" & SALEACC & "%' AND Parent=0 Order By ACCTitle")
    cmbSaleAcc.ID = "31-001-001"
    'Call AddToCombo(cmbPRTAcc, "AccTitle", "VActiveAccounts", " Where AccNo Like '" & PRTAcc & "%' AND Parent=0 Order By ACCTitle")
    Call cmbSaleRAcc.AddVals(con, "AccTitle", "VActiveAccounts", "AccNo", " Where AccNo Like '" & SRTACC & "%' AND Parent=0 Order By ACCTitle")
    cmbSaleRAcc.ID = "31-003-003"
    
    Call cmbAssetAcc.AddVals(con, "AccTitle", "VActiveAccounts", "AccNo", " Where AccNo Like '" & STOCKASSETACC & "%' AND Parent=0 Order By ACCTitle")
    cmbAssetAcc.ID = "15-005-001"
    'Call AddToCombo(cmbAssetAcc, "AccTitle", "VActiveAccounts", " Where AccNo Like '" & StockAssetAcc & "%' AND Parent=0 Order By ACCTitle")
    
    'Call AddToCombo(cmbUnit, "Distinct Unit", "Items", "Order By Unit")
    'Call AddToCombo(cmbWeightUnit, "Distinct UnitWeight", "Items", "Order By UnitWeight")
    
    Call AddToCombo(cmbComp, "Description", "Components", , , "CompID")
    
    cmbCompProc.ListHeight = 2400
    Call cmbCompProc.ClearVals
    Call cmbCompProc.AddVals(con, "Description", "Processes", "ProcessID")
    
    
    'Call AddToCombo(cmbRM, "RMName + '( ' + Unit + ' )'", "RM", , , "RMID,RMName,Unit")
    'cmbRawMaterial.AddVals con, "RMName + '( ' + Unit + ' )'", "RM", "RMID"
    cmbRawMaterial.AddVals con, "RMID1 + ' ('+RMName+')'+RawMaterial", "VRMForItems", "RMID1"
    
    cmbProcRM.ListHeight = 2400
    cmbRawMaterial.ListHeight = 2400
'    Call cmbProcRM.ClearVals
'    Call cmbProcRM.AddVals(con, "Description", "Processes", "ProcessID")
    
    Call AddToCombo(cmbCatID, "Description", "ItemCatagories", " Order By Description", , "CatID")
    
    Call AddToCombo(cmbSteel, "SteelType", "SteelTypes", , , "SteelID")
    If cmbSteel.ListCount > 0 Then cmbSteel.ListIndex = 0
    
    Call AddToCombo(cmbItemGroup, "Description", "ItemGroups", " Order By Description", , "ID")
    If cmbItemGroup.ListCount > 0 Then cmbItemGroup.ListIndex = 0
    
    Call AddToCombo(cmbItemGroup_ECommerce, "ECommerceGroupName", "ItemGroups_ECommerce", , , "ID")
    
    Call AddToCombo(cmbItemGroupMain, "MainGroupName", "ItemGroupsMain", " ORDER BY MainGroupName", , "MainGroupID")
    If cmbItemGroupMain.ListCount > 0 Then cmbItemGroupMain.ListIndex = 0
    
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
    
    With cmbItemType
        .AddItem "Finished"
        .List(0, 1) = 0
        If strCompany = "Tecno" Or strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
            .AddItem "Semi Finished"
            .List(1, 1) = 1
        End If
        If strCompany = "IAA" Or strCompany = "Kami" Then
            .AddItem "Set"
            .List(2, 1) = 2
        End If
        .ListIndex = 0
    End With
        
    With cmbProcessGroup
        .ListHeight = 4000
        .AddItem "<All Groups>", "0"
        .AddVals con, "GroupName", "ProcessGroups", "EntryID", " ORDER BY GroupName"
    End With
    cmbProcess.ListHeight = 4000
        
    With cmbScanning
        .AddItem "No"
        .AddItem "Yes"
    End With
    With cmbScanningEdit
        .AddItem "No"
        .AddItem "Yes"
    End With
    
    AddToCombo cmbForeignCatalogName, "CatalogName", "ForeignCatalog", , , "CatalID"
    
End Sub

Private Sub chkPlanter_Click()
    NBox(1).Visible = chkPlanter.Value
    If Not chkPlanter.Value Then NBox(1) = ""
End Sub

Private Sub chkShowInECommerce_Click()
    
    
    If chkShowInECommerce.Value Then
        FraECommerce.Enabled = True
        If bDataLoading = False Then
            TBox(21) = TBox(0)
        End If
    Else
        FraECommerce.Enabled = False
        If bDataLoading = False Then
            TBox(21) = ""
        End If
    End If
    
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

Private Sub cmbItemLookalike_matched()
    txtFinishedStock = GetSingleLongValue("InHand", "VItems", " WHERE ItemID='" & cmbItemLookalike.ID & "'")
End Sub

Private Sub cmbItemType_Change()
    Call cmbItemType_Click
End Sub

Private Sub cmbItemType_Click()
    If cmbItemType.MatchFound = False Then Exit Sub
    If Val(cmbItemType.List(cmbItemType.ListIndex, 1)) = 2 Then
        SSTab1.TabVisible(10) = True
        cmbItemSet.ClearVals
        Call cmbItemSet.AddVals(con, "ItemID + ' ' + ItemName + ' { ' + ISNULL(ItemSize,'') + ' }'", "Items", "ItemID", " WHERE InActive=0 AND ItemID<>'" & ItemCode & "' ORDER BY ItemName")
    Else
        SSTab1.TabVisible(10) = False
    End If
End Sub

Private Sub cmbProcessGroup_matched()

    With cmbProcess
        .ClearVals
        .AddItem "<All Processes>", "0"
        If cmbProcessGroup.ID = "0" Then
            .AddVals con, "Description + ' (' + ISNULL(Code,'') + ')'", "Processes", "ProcessID", " ORDER BY SNo"
        Else
            .AddVals con, "Description + ' (' + ISNULL(Code,'') + ')'", "ProcessGroupsProcesses PGP INNER JOIN Processes P ON PGP.Process_RefID=P.ProcessID", "P.ProcessID", " WHERE PGP.Group_RefID=" & cmbProcessGroup.ID & " ORDER BY PGP.SeqNo"
        End If
    End With
    
End Sub

Private Sub cmbScanningEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LV2.SelectedItem.SubItems(1) = cmbScanningEdit.Text
        LV2.SelectedItem.ListSubItems(1).Tag = cmbScanningEdit.ListIndex
        cmbScanningEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        cmbScanningEdit.Visible = False
    End If
End Sub

Private Sub cmbScanningEdit_LostFocus()
    cmbScanningEdit.Visible = False
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
    ITM.Tag = cmbCompProc.ID
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
        If cmbRawMaterial.ID & "'" = LVRM.ListItems(i).key Then
            MsgBox "This Component Is Already Added.", vbInformation
            Exit Sub
        End If
    Next i
    
    Dim strUnit As String, strMaterialName As String
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT RMName,Unit FROM RM WHERE RMID=" & cmbRawMaterial.ID, con, adOpenForwardOnly, adLockReadOnly
        strMaterialName = !RMName & ""
        strUnit = !Unit & ""
        .Close
    End With
    Set rs = Nothing
    
    Set ITM = LVRM.ListItems.add(, cmbRawMaterial.ID & "'", cmbProcRM.Text)
    ITM.Tag = cmbProcRM.ID
    ITM.ListSubItems.add , , strMaterialName
    ITM.ListSubItems.add , , Val(NBox(13))
    
    Exit Sub
err:
    MsgBox err.Description

End Sub


Private Sub cmdAddLookalike_Click()
    
    Dim i As Integer
    For i = 1 To LVLookalike.ListItems.count
        If cmbItemLookalike.ID = LVLookalike.ListItems(i).ListSubItems(1).Tag Then
            MsgBox "Already Added.", vbInformation
            Exit Sub
        End If
    Next
    Dim ITM As ListItem
    Set ITM = LVLookalike.ListItems.add(, , LVLookalike.ListItems.count + 1)
    ITM.Tag = "0"
    ITM.ListSubItems.add(, , cmbItemLookalike.Text).Tag = cmbItemLookalike.ID
    ITM.ListSubItems.add , , Val(txtFinishedStock)
    
End Sub

Private Sub cmdBreakLink_Click()

    txtProcessFamily.Tag = "Del"
    txtProcessFamily = ""
    cmdBreakLink.Tag = ""
    cmdBreakLink.Enabled = False
    cmdDeleteAll.Enabled = True
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChangeTechDrawing_Click()
    AddTechDrawing
End Sub

Private Sub cmdDelete_Click()
    LV1.ListItems.Remove (LV1.SelectedItem.key)
End Sub

Private Sub cmdComp_Click()
    frmComponents.Show 1
End Sub

Private Sub cmdDeleteRM_Click()
    LVRM.ListItems.Remove (LVRM.SelectedItem.key)
End Sub

Private Sub cmdDeleteAll_Click()
    If LV2.ListItems.count > 0 Then
        If MsgBox("Are you sure to Delete All?", vbQuestion + vbYesNo) = vbYes Then
            LV2.ListItems.Clear
            txtProcessSortNo.Text = "1"
            bDelAllProesses = True
        End If
    End If
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


Private Sub cmdRemoveTechnicalDrawing_Click()

    lblTechDrawing.Caption = "Click here to Add."
    TechDrawingFileName = "Remove"
    cmdChangeTechDrawing.Visible = False
    cmdRemoveTechnicalDrawing.Visible = False
    
End Sub

Private Sub cmdUpdateForAllItems_Click()
    
    Exit Sub
    If cmbCatID.MatchFound = False Then
        MsgBox "Please Select Catagory.", vbInformation
        Exit Sub
    End If
    
    If MsgBox("Are you sure?" & vbNewLine & "this will update Processes for All the Items of selected Catagory." & vbNewLine & "Please note that this action is irreversible.", vbQuestion + vbYesNo) = vbYes Then
        Call SaveProcessesForAllItemsOfSameCatagory
    End If
    
End Sub

Private Sub cmdViewActualSizeEC_Click()
    
    If ImgEC.Picture = 0 Then Exit Sub
    Dim f As New frmViewPictureFSExt
    Load f
    f.ShowFromPic ImgEC.Picture
    
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
    'If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    bDataLoading = False
    bSample = False
    PicFileName = ""
    strCopyFrom = ""
    strPicECFileName = ""
    Call getNextProcessSeqNo
    'For i = 0 To SSTab1.Tabs - 1
    '    SSTab1.Tab = i
    '    SSTab1.Picture = ImageList1.ListImages(i & "'").ExtractIcon
    'Next

    SSTab1.TabVisible(1) = True
    SSTab1.TabVisible(2) = False
    SSTab1.TabVisible(4) = False
    SSTab1.TabVisible(10) = False
    'SSTab1.TabVisible(6) = False
    SSTab1.TabVisible(8) = True
        
    If strCompany = "Dr-Frgz" Then
        SSTab1.TabVisible(12) = True
    Else
        SSTab1.TabVisible(12) = False
    End If
    
    If CurrentUserName <> "Administrator" Then
        SSTab1.TabVisible(5) = False
        NBox(24).Locked = True
    End If
    
    If LCase(CurrentUserName) = "production" Then
        SSTab1.TabVisible(5) = False
    End If
    SSTab1.Tab = 0
    
    'SSTab2.TabVisible(1) = False
    
    Me.KeyPreview = True
    
    'DTPicker1.value = ServerDate
    
    cmbSaleAcc.ListHeight = 1500
    cmbSaleRAcc.ListHeight = 1500
    cmbAssetAcc.ListHeight = 1500
    
    cmbCatID.Locked = False
    TBox(3).Locked = False
    TBox(6).Text = ServerDate
    
'    Call AddProcesses
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
    cmbItemSet.ListHeight = 4000
    cmbItemSet.DropDownOffLine = True
    
    strDelListSet = ""
    strDelListRM = ""
    strDelListSF = ""
    strDelListLookAlikes = ""
    'CopyItemCode = ""
    
    
    If strCompany = "Tecno" Or strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
        lblMainCurr(0) = "US $"
        lblMainCurr(1) = "US $"
        lblSubCurr(0) = "Euro"
        lblSubCurr(1) = "Euro"
        lblConversion.Caption = "US$ to Euro"
        
        cmbCurr.AddItem "US$"
        cmbCurr.ListIndex = 0
        cmbCurr.Enabled = False
    Else
        cmbCurr.AddItem "EUR"
        cmbCurr.ListIndex = 0
        cmbCurr.Enabled = False
    End If
    
    cmbItemLookalike.ListHeight = 3000
    cmbItemLookalike.DropDownOffLine = True
    cmbItemLookalike.AddVals con, "'{' + ItemID + '} ' + ItemName", "Items", "ItemID"
    strDelListItemProcs = ""
    bDelAllProesses = False
        
    'Now Check for Access of Specific Tabs....
    
    With rs
        .Open "SELECT GeneralInfoItemProfile,ProcessesItemProfile,WeightItemProfile,PriceItemProfile," & _
         "ReferencesItemProfile,PictureItemProfile,RMItemProfile,ShipInfoItemProfile FROM Users WHERE UserID=" & UserID, con, adOpenForwardOnly, adLockReadOnly
        
        SSTab1.TabVisible(0) = !GeneralInfoItemProfile
        SSTab1.TabVisible(1) = !ProcessesItemProfile
        SSTab1.TabVisible(3) = !WeightItemProfile
        SSTab1.TabVisible(5) = !PriceItemProfile
        SSTab1.TabVisible(6) = !ReferencesItemProfile
        SSTab1.TabVisible(7) = !PictureItemProfile
        SSTab1.TabVisible(8) = !RMItemProfile
        SSTab1.TabVisible(9) = !ShipInfoItemProfile
        
        .Close
    End With
    
End Sub

Private Function SaveItem() As Boolean

    On Error GoTo err
    
    If ItemCode <> "" Then
        If GetSingleLongValue("COUNT(ItemCode)", "ItemsLockedForEditing", " WHERE ItemCode='" & ItemCode & "'") > 0 Then   'Item Is Locked...
            If getDBPassword("LockedItem") = False Then
                Exit Function
            End If
        End If
    End If
    If ItemCode = "" Then
        lCount = GetSingleLongValue("COUNT(*)", "Items", " WHERE ItemID='" & TBox(3) & "'")
        If lCount > 0 Then
            MsgBox "Item Code already exists.", vbInformation
            Exit Function
        End If
    End If
    Dim OpBalance As Double
    
    If Trim(TBox(0)) = "" Then
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
        If strCompany = "Tecno" Or strCompany = "Dr-Frgz" Then
            MsgBox "Please Select Item Type.", vbInformation
            Exit Function
        Else
            cmbItemType.ListIndex = 0
        End If
    ElseIf cmbItemGroup.MatchFound = False Then
        MsgBox "Please Select Group.", vbInformation
        Exit Function
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

    Dim SQL As String
    
    Call StartTrans(con)

    If ItemCode = "" Then
        SQL = CreateInsertSQL("Items")
        'con.Execute "Update Accounts Set OpenBal=OpenBal+" & TotalCost & " Where AccNo='" & AssetAcc & "'"
        'Call adjustLedger(AssetAcc, "1/1/1900", 0, TotalCost)
    Else
        SQL = CreateUpdateSQL("Items")
        SQL = SQL & " WHERE ItemID='" & ItemCode & "'"
        'Update ItemSets
        con.Execute "UPDATE ItemsSets SET Set_ItemID='" & TBox(3) & "' WHERE Set_ItemID='" & ItemCode & "'"
    End If
    
    con.Execute SQL
    If ItemCode = "" Then
        If bSample Then
            con.Execute "UPDATE Items SET Sample=1 WHERE ItemID='" & TBox(3).Text & "'"
        End If
    End If
    
    If bUrduSupported Then
        con.Execute "UPDATE Items SET ItemNameUrdu=N'" & txtUrduFields(0).Text & "' WHERE ItemID='" & TBox(3).Text & "'", i
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
    
    If ItemCode <> "" Then
        If TBox(3) <> ItemCode Then
            con.Execute "UPDATE FOrderItems SET CompItemCode='" & TBox(3) & "' WHERE CompItemCode='" & ItemCode & "'"
            con.Execute "UPDATE VendIssdDetail SET ItemCode='" & TBox(3) & "' WHERE ItemCode='" & ItemCode & "'"
            con.Execute "UPDATE VendRcvdDetail SET ItemCode='" & TBox(3) & "' WHERE ItemCode='" & ItemCode & "'"
            con.Execute "UPDATE VendIssued SET ItemID='" & TBox(3) & "' WHERE ItemID='" & ItemCode & "'"
            con.Execute "UPDATE StockOrderOpening SET ItemID='" & TBox(3) & "' WHERE ItemID='" & ItemCode & "'"
            con.Execute "UPDATE StockOrderAdjustments SET ItemID='" & TBox(3) & "' WHERE ItemID='" & ItemCode & "'"
            con.Execute "UPDATE ItemLookalikes SET ItemID_LookAlike='" & TBox(3) & "' WHERE ItemID_LookAlike='" & ItemCode & "'"
            con.Execute "UPDATE EmpDailyTargets SET ItemID='" & TBox(3) & "' WHERE ItemID='" & ItemCode & "'"
        End If
    End If
    
    If bDelAllProesses Then
        con.Execute "DELETE FROM ItemProcesses WHERE ItemID='" & ItemCode & "'"
    End If
    
    If (strCompany = "Weldon Industries" Or strCompany = "Weldon Instruments") And ItemCode = "" Then  'Assign All Processes to new item
        con.Execute "INSERT INTO ItemProcesses(ProcessID,ItemID,Rate,SNo)" & _
         "SELECT ProcessID,'" & TBox(3) & "',0,SNo FROM Processes ORDER BY SNo"
    Else
        Call SaveProcesses
    End If
    
    Call SaveComponents
    Call SaveComponentsRM
    Call SaveForeignCatal
    Call SaveSetDetail
    Call SaveItemLookAlikes
    If strCompany = "Weldon Instruments" Then
        'Assign All Processes to Factory Maker....
        If ItemCode <> "" Then
            con.Execute "DELETE FROM VendAssItems WHERE ItemID IN ('" & ItemCode & "','" & TBox(3) & "') AND VendID=15"
        End If
        con.Execute "INSERT INTO VendAssItems(VendID,ProcessID,ItemID) SELECT DISTINCT 15,ProcessID,ItemID FROM ItemProcesses " & _
         "WHERE ProcessID IN(SELECT ProcessID FROM Processes WHERE Operation IN(0,2)) AND ItemID='" & TBox(3) & "' ORDER BY ItemID,ProcessID"
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
        Rec.Open "Select ItemID,ItemPic from Items Where ItemID='" & TBox(3) & "'", con, adOpenStatic, adLockOptimistic
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
    Else
        If strCopyFrom <> "" Then
            con.Execute "UPDATE Items SET Items.ItemPic=T2.ItemPic FROM Items," & _
             "(SELECT ItemPic FROM Items WHERE ItemID='" & strCopyFrom & "') T2 WHERE ItemID='" & TBox(3) & "'"
        End If
    End If
    
    'If ImgTechDrawing.Picture <> 0 And Trim(TechDrawingFileName) <> "" Then
    If Trim(TechDrawingFileName) <> "" Then
        'rec.Close
        If TechDrawingFileName = "Remove" Then
            con.Execute "UPDATE Items SET TechnicalDrawing=NULL WHERE ItemID='" & TBox(3) & "'"
        Else
            Rec.Open "SELECT ItemID,TechnicalDrawing FROM Items WHERE ItemID='" & TBox(3) & "'", con, adOpenStatic, adLockOptimistic
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
    End If
    
    If Trim(strPicECFileName) <> "" Then
        'rec.Close
        Rec.Open "SELECT ItemID,PicForECommerce FROM Items WHERE ItemID='" & TBox(3) & "'", con, adOpenStatic, adLockOptimistic
        If Rec.RecordCount > 0 Then
            rsPic.Open
            rsPic.Type = adTypeBinary
            rsPic.LoadFromFile strPicECFileName
            Rec.Fields(1) = rsPic.Read
            rsPic.Close
            Rec.Update
        End If
        
        Rec.Close
        
        Set Rec = Nothing
        Set rsPic = Nothing
    Else
        If lblPic.Visible = True Then
            con.Execute "UPDATE Items SET ItemPic=NULL WHERE ItemID='" & TBox(3) & "'"
        End If
    End If
    strNewItemCode = TBox(3)
    
    Exit Function
err:
    MsgBox err.Description
    
End Function

Private Sub SaveItemLookAlikes()

    If strDelListLookAlikes <> "" Then
        strDelListLookAlikes = Left(strDelListLookAlikes, Len(strDelListLookAlikes) - 1)
        con.Execute "DELETE FROM ItemLookalikes WHERE EntryID IN(" & strDelListLookAlikes & ")"
    End If
    
    Dim i As Integer
    For i = 1 To LVLookalike.ListItems.count
        With LVLookalike.ListItems(i)
            If Val(.Tag) = 0 Then
                con.Execute "INSERT INTO ItemLookalikes(ItemID,ItemID_LookAlike) VALUES('" & TBox(3) & "','" & _
                 .ListSubItems(1).Tag & "')"
            End If
        End With
    Next
    
End Sub

Private Sub SaveSetDetail()

    Dim i As Integer
    
    If strDelListSet <> "" Then
        strDelListSet = Left(strDelListSet, Len(strDelListSet) - 1)
        con.Execute "DELETE FROM ItemsSets WHERE EntryID IN(" & strDelListSet & ")"
    End If
    
    For i = 1 To LVSet.ListItems.count
        With LVSet.ListItems(i)
            If Val(.Tag) = 0 Then
                con.Execute "INSERT INTO ItemsSets(ItemID,Set_ItemID,Qty) VALUES('" & TBox(3) & _
                 "','" & .ListSubItems(1).Tag & "'," & Val(.SubItems(4)) & ")"
            Else
                con.Execute "UPDATE ItemsSets SET Qty=" & Val(.SubItems(4)) & " WHERE EntryID=" & Val(.Tag)
            End If
        End With
    Next
    
End Sub
Private Sub SaveRevision()
    con.Execute "Insert Into ItemRevisions(ItemID,RevDTTime) Values('" & TBox(3) & "','" & Now & "')"
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



Private Sub ImgEC_DblClick()
    Call AddPicEC
End Sub

Private Sub ImgTechDrawing_DblClick()
'    Call AddTechDrawing
End Sub

Private Sub label1_Click(Index As Integer)
    If Index = 69 Then
        frmForeignCatalog.Show 1
        AddToCombo cmbForeignCatalogName, "CatalogName", "ForeignCatalog", , , "CatalID"
    End If
End Sub

Private Sub lblECommerceGroup_Click()
        
    Dim f As New frmItemGroups_ECommerce
    Load f
    f.Show 1
    
    Call AddToCombo(cmbItemGroup_ECommerce, "ECommerceGroupName", "ItemGroups_ECommerce", , , "ID")
    
End Sub

Private Sub lblPic_DblClick()
    Call AddPic
End Sub


Private Sub lblPicEC_DblClick()
    Call AddPicEC
End Sub

Private Sub lblTechDrawing_Click()
    If lblTechDrawing.Caption = "Click here to Open Tech. Drawing." Then
        Call ShowTechDrawing("Items", "TechnicalDrawing", " WHERE ItemID='" & ItemCode & "'", ItemCode, CD1)
    Else
        Call AddTechDrawing
    End If
End Sub

Private Sub ShowTechDrawingPrivate()

    On Error GoTo err
     
    Dim strSavePath As String, strProgramPath As String
    strSavePath = GetTemporaryDirectory '"\\Database\Software Attachments"
    
    Dim rs As New ADODB.Recordset
    Dim binObj As New ADODB.Stream
    rs.Open "SELECT TechnicalDrawing FROM Items WHERE ItemID='" & ItemCode & "'", con, adOpenForwardOnly, adLockReadOnly
    
    binObj.Type = adTypeBinary
    binObj.Open
    binObj.Write rs!TechnicalDrawing
    
    strSavePath = strSavePath & ItemCode & ".pdf"
    binObj.SaveToFile strSavePath, adSaveCreateOverWrite
    binObj.Close
    Set binObj = Nothing
         
    rs.Close
    Set rs = Nothing
     
    strProgramPath = GetSetting(App.Title, "SettingsEx", "PDF Path", "C:\Program Files\Adobe\Acrobat 7.0\Reader\AcroRd32.exe")
    Dim oFso As New FileSystemObject
    If oFso.FileExists(strProgramPath) = False Then
        On Error Resume Next
        CD1.DialogTitle = "Please Select Adobe Acrobat Executeable"
        CD1.Filter = "Acrobat Executeable|*.exe"
        CD1.ShowOpen
        If err.Number = 32755 Then Exit Sub
        If CD1.FileName <> "" Then
            strProgramPath = CD1.FileName
            SaveSetting App.Title, "SettingsEx", "Acrobat Path", strProgramPath
        End If
    End If
    'strProgramPath = "C:\Program Files\Adobe\Acrobat 7.0\Reader\AcroRd32.exe"
    Shell """" & strProgramPath & """" & " " & """" & strSavePath & """", vbNormalFocus
    
    Exit Sub
err:
    MsgBox err.Description
    
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
    'If LV2.SelectedItem.Checked = False Then Exit Sub
    Set SelItm = LV2.SelectedItem
    
'    With SelItm
'        txtProcRate.Move SSTab2.Left + LV2.ColumnHeaders(2).Left, .Top + LV2.Top, LV2.ColumnHeaders(2).Width
'        txtProcRate.Visible = True
'        txtProcRate = Val(.SubItems(1) & "")
'        txtProcRate.SetFocus
'        txtProcRate.Tag = "Edit"
'    End With
    If iColNoProcess = 2 Or iColNoProcess = 3 Then
        If cmdBreakLink.Enabled Then
            MsgBox "Can't Edit, Process Family is attached.", vbInformation
            Exit Sub
        End If
    End If
    If iColNoProcess = 3 Then
        With txtSortNo
            .Move LV2.Left + LV2.ColumnHeaders(3).Left, LV2.Top + SelItm.Top, LV2.ColumnHeaders(3).Width
            .Text = Val(SelItm.SubItems(2))
            .SelStart = 0
            .SelLength = Len(.Text)
            .Visible = True
            .SetFocus
        End With
    ElseIf iColNoProcess = 2 Then
        With cmbScanningEdit
            .Move LV2.Left + LV2.ColumnHeaders(2).Left, LV2.Top + SelItm.Top, LV2.ColumnHeaders(2).Width
            .ListIndex = SelItm.ListSubItems(1).Tag
            .Visible = True
            .SetFocus
        End With
    ElseIf iColNoProcess = 4 Then
        With txtStRate
            .Move LV2.Left + LV2.ColumnHeaders(iColNoProcess).Left, LV2.Top + SelItm.Top, LV2.ColumnHeaders(iColNoProcess).Width
            .Text = SelItm.SubItems(iColNoProcess - 1)
            .SelStart = 0
            .SelLength = Len(.Text)
            .Visible = True
            .SetFocus
        End With
    ElseIf iColNoProcess = 5 Or iColNoProcess = 6 Or iColNoProcess = 7 Then
        With txtStRate
            .Move LV2.Left + LV2.ColumnHeaders(iColNoProcess).Left, LV2.Top + SelItm.Top, LV2.ColumnHeaders(iColNoProcess).Width
            .Text = Val(SelItm.SubItems(iColNoProcess - 1))
            .SelStart = 0
            .SelLength = Len(.Text)
            .Visible = True
            .SetFocus
        End With
    End If
    
    
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

Private Sub LV2_KeyDown(KeyCode As Integer, Shift As Integer)

    If LV2.ListItems.count = 0 Then Exit Sub
    
    If KeyCode = 46 Then
        If cmdBreakLink.Enabled Then
            MsgBox "Can't Delete, Process Family is attached !!!", vbInformation
            Exit Sub
        End If
        If MsgBox("Are you sure to Delete ?", vbQuestion + vbYesNo) = vbYes Then
            strDelListItemProcs = strDelListItemProcs & LV2.SelectedItem.ListSubItems(2).Tag & ","
            LV2.ListItems.Remove LV2.SelectedItem.Index
            Call ReIndexProcesses
        End If
    End If
    
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

Private Sub LV2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim i As Integer
    For i = 1 To LV2.ColumnHeaders.count
        If X > LV2.ColumnHeaders(i).Left And X < (LV2.ColumnHeaders(i).Left + LV2.ColumnHeaders(i).Width) Then
            iColNoProcess = i
            Exit For
        End If
    Next
    
End Sub

Private Sub LVLookalike_KeyDown(KeyCode As Integer, Shift As Integer)

    If LVLookalike.ListItems.count = 0 Then Exit Sub
    
    If KeyCode = 46 Then
        If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo) = vbYes Then
            strDelListLookAlikes = strDelListLookAlikes & Val(LVLookalike.SelectedItem.Tag) & ","
            LVLookalike.ListItems.Remove LVLookalike.SelectedItem.Index
        End If
    End If
    
End Sub

Private Sub LVRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If LVRef.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure to Remove?", vbYesNo + vbQuestion) = vbYes Then
            LVRef.ListItems.Remove LVRef.SelectedItem.Index
        End If
    End If
End Sub

Private Sub LVRM_KeyDown(KeyCode As Integer, Shift As Integer)

    If LVRM.ListItems.count = 0 Then Exit Sub
    
    If KeyCode = 46 Then
        If MsgBox("Are you Sure to Remove?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
        If Val(LVRM.SelectedItem.ListSubItems(2).Tag) > 0 Then
            If LVRM.SelectedItem.ListSubItems(3).Tag = "1" Then
                strDelListRM = strDelListRM & Val(LVRM.SelectedItem.ListSubItems(2).Tag) & ","
            Else
                strDelListSF = strDelListSF & Val(LVRM.SelectedItem.ListSubItems(2).Tag) & ","
            End If
        End If
        LVRM.ListItems.Remove LVRM.SelectedItem.Index
    End If
    
End Sub

Private Sub LVSet_DblClick()
    If LVSet.ListItems.count = 0 Then Exit Sub
    With txtEditSetQty
        .Move LVSet.Left + LVSet.ColumnHeaders(5).Left, LVSet.Top + LVSet.SelectedItem.Top, LVSet.ColumnHeaders(5).Width
        .Text = Val(LVSet.SelectedItem.SubItems(4))
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub LVSet_KeyDown(KeyCode As Integer, Shift As Integer)

    If LVSet.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure to Remove this ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        If Val(LVSet.SelectedItem.Tag) > 0 Then
            strDelListSet = strDelListSet & Val(LVSet.SelectedItem.Tag) & ","
        End If
        LVSet.ListItems.Remove LVSet.SelectedItem.Index
    End If
    
End Sub

Private Sub NBox_Change(Index As Integer)

    Select Case Index
        Case 4, 5, 7
            NBox(6) = Val(NBox(4)) - Val(NBox(5)) - Val(NBox(7))
        Case 0, 14, 15, 16, 8, 9, 10, 11, 44, 45, 46, 47, 48, 49, 50, 51
            Call txtConversion_Change
        Case 25
            NBox(55) = Round((Val(NBox(25)) * 0.3), 2)
    End Select
    
End Sub

Private Sub AddToRMGrid()

    If cmbProcRM.MatchFound = False Then
        MsgBox "Please Select Process.", vbInformation
        Exit Sub
    ElseIf cmbRawMaterial.MatchFound = False Then
        MsgBox "Please Select Material.", vbInformation
        Exit Sub
    ElseIf Val(NBox(13)) <= 0 Then
        MsgBox "Invalid Qty.", vbInformation
        Exit Sub
    End If
    
    Dim ITM As ListItem
    Dim i As Integer
    For i = 1 To LVRM.ListItems.count
        If LVRM.ListItems(i).ListSubItems(1).Tag = cmbRawMaterial.ID Then
            MsgBox "Already Added.", vbInformation
            Exit Sub
        End If
    Next
    
    Set ITM = LVRM.ListItems.add(, , cmbProcRM.Text)
    ITM.Tag = cmbProcRM.ID
    ITM.ListSubItems.add(, , cmbRawMaterial.Text).Tag = cmbRawMaterial.ID
    ITM.ListSubItems.add , , Val(NBox(13))
    
    ITM.ListSubItems.add().Tag = IIf(Right(cmbRawMaterial.Text, 2) = "RM", 1, 0)    '(3).tag=1 means RM
    
End Sub

Private Sub NBox_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If Index = 13 Then
        If KeyCode.Value = 13 Then
            Call AddToRMGrid
        End If
    End If
End Sub

Private Sub Tbox_Change(Index As Integer)
    Exit Sub
    Select Case Index
    
    Case 8, 3
        TBox(4) = TBox(8) & TBox(3)
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
            iPos1 = InStr(TBox(3), "-")
            If iPos1 > 0 Then
                iPos2 = InStr(iPos1 + 1, TBox(3), "-")
            End If
            If iPos1 = 0 Or iPos2 = 0 Then
                MsgBox "Can't Format the Item Code.", vbInformation
                Exit Sub
            End If
            'Now Format it...
            Dim strFirst As String, strSecond As String, strThird As String
            strFirst = Left(TBox(3), iPos1 - 1)
            strSecond = Mid(TBox(3), iPos1 + 1, iPos2 - (iPos1 + 1))
            strThird = Mid(TBox(3), iPos2 + 1)
            
            If Left(strFirst, 1) = "D" Then
                strFirst = "D" & Format(Right(strFirst, Len(strFirst) - 1), "0000")
            End If
            TBox(3) = strFirst & "-" & Format(strSecond, "000") & "-" & Format(strThird, "000")
        End If
    End If
End Sub

Private Sub TBox_Validate(Index As Integer, cancel As Boolean)
    Exit Sub
    If Index = 5 Then
        TBox(Index) = Format(Val(TBox(Index)), "#0.00")
    ElseIf Index = 2 Or Index = 3 Then
        TBox(Index) = Format(Val(TBox(Index)), "#0")
    End If
End Sub

Private Sub SaveComponents()

    con.Execute "Delete From ItemsComp Where ItemID='" & ItemCode & "'"

    For i = 1 To LV1.ListItems.count
        con.Execute "Insert Into ItemsComp(ItemID,CompID,Qty,ProcessID) Values('" & TBox(3) & "'," & Val(LV1.ListItems(i).key) & "," & LV1.ListItems(i).SubItems(2) & "," & Val(LV1.ListItems(i).Tag) & ")"
    Next i
    
End Sub

Private Sub SaveComponentsRM()

    'con.Execute "DELETE FROM ItemsRMComp WHERE ItemID='" & ItemCode & "'"
    If strDelListRM <> "" Then
        strDelListRM = Left(strDelListRM, Len(strDelListRM) - 1)
        con.Execute "DELETE FROM ItemsRMComp WHERE EntryID IN(" & strDelListRM & ")"
        'con.Execute "DELETE FROM ItemsSFMaterial WHERE EntryID IN(" & strDelListRM & ")"
    End If
    If strDelListSF <> "" Then
        strDelListSF = Left(strDelListSF, Len(strDelListSF) - 1)
        con.Execute "DELETE FROM ItemsSFMaterial WHERE EntryID IN(" & strDelListSF & ")"
    End If
    
    Dim strMaterialID As String
    For i = 1 To LVRM.ListItems.count
        With LVRM.ListItems(i)
            If .ListSubItems(3).Tag = "1" Then
                strMaterialID = GetSingleLongValue("RMID", "RM", " WHERE RMID1='" & .ListSubItems(1).Tag & "'")
            Else
                strMaterialID = .ListSubItems(1).Tag
            End If
            If Val(.ListSubItems(2).Tag) > 0 Then
                'Do Nothing; Editing is not allowed.
            Else
                If .ListSubItems(3).Tag = "1" Then 'Raw Material
                    con.Execute "INSERT INTO ItemsRMComp(ItemID,RMID,Qty,ProcessID) VALUES('" & TBox(3) & "'," & strMaterialID & "," & Val(LVRM.ListItems(i).SubItems(2)) & "," & Val(LVRM.ListItems(i).Tag) & ")"
                Else
                    con.Execute "INSERT INTO ItemsSFMaterial(ItemID,ItemID_SF,Qty,ProcessID) VALUES('" & TBox(3) & "','" & strMaterialID & "'," & Val(LVRM.ListItems(i).SubItems(2)) & "," & Val(LVRM.ListItems(i).Tag) & ")"
                End If
            End If
        End With
    Next i
    
End Sub

Private Sub SaveProcesses()
    
    'Save Item's Production Processes...
    
    '------------------------------------------------------------------------
    If strDelListItemProcs <> "" Then
        strDelListItemProcs = Left(strDelListItemProcs, Len(strDelListItemProcs) - 1)
        con.Execute "INSERT INTO ItemProcesses_Deletions(ProcessID,ItemID,Rate,SNO,Scanning,MinLevel,MaxLevel,ReOrderLevel," & _
         "UserName,MachineName) SELECT ProcessID,ItemID,Rate,SNO,Scanning,MinLevel,MaxLevel,ReOrderLevel,'" & _
         CurrentUserName & "','" & strComputerName & "' FROM ItemProcesses WHERE EntryID IN(" & strDelListItemProcs & ")"
         
        con.Execute "DELETE FROM ItemProcesses WHERE EntryID IN (" & strDelListItemProcs & ")"
    End If
    
    'con.Execute "DELETE FROM ItemProcesses WHERE ItemID IN ('" & ItemCode & "','" & TBox(3) & "')"
    
    For i = 1 To LV2.ListItems.count
        'If LV2.ListItems(i).Checked Then
            'con.Execute "Insert Into ItemProcesses(ItemID,ProcessID,Rate) " & _
             " Values('" & TBox(4) & "'," & Val(LV2.ListItems(i).Key) & "," & _
             Val(LV2.ListItems(i).SubItems(1) & "") & ")"
        With LV2.ListItems(i)
        If Val(.ListSubItems(2).Tag) = 0 Then
            con.Execute "INSERT INTO ItemProcesses(ItemID,ProcessID,Scanning,SNO,Rate,MinLevel,MaxLevel,ReOrderLevel) " & _
             " VALUES('" & TBox(3) & "'," & Val(LV2.ListItems(i).key) & "," & Val(LV2.ListItems(i).ListSubItems(1).Tag) & _
             "," & Val(LV2.ListItems(i).SubItems(2)) & "," & Val(LV2.ListItems(i).SubItems(3)) & "," & Val(LV2.ListItems(i).SubItems(4)) & _
             "," & Val(LV2.ListItems(i).SubItems(5)) & "," & Val(LV2.ListItems(i).SubItems(6)) & ")"
        Else
            con.Execute "UPDATE ItemProcesses SET SNo=" & Val(.SubItems(2)) & ",Rate=" & Val(.SubItems(3)) & ",MinLevel=" & Val(.SubItems(4)) & _
                 ",MaxLevel=" & Val(.SubItems(5)) & ",ReOrderLevel=" & Val(.SubItems(6)) & " WHERE EntryID=" & Val(.ListSubItems(2).Tag)
        End If
        End With
        'End If
    Next i
    
    'Save Process Group Details....
    If txtProcessFamily.Tag = "New" Then
        con.Execute "INSERT INTO ItemProcessGroups_Deletions(ItemID,PG_RefID,UserName,MachineName) VALUES('" & TBox(3) & "'," & Val(cmdBreakLink.Tag) & ",'" & CurrentUserName & "','" & strComputerName & "')"
        con.Execute "DELETE FROM ItemProcessGroups WHERE ItemID='" & ItemCode & "'"
        con.Execute "INSERT INTO ItemProcessGroups(ItemID,PG_RefID) VALUES('" & TBox(3) & "'," & Val(cmdBreakLink.Tag) & ")"
    ElseIf txtProcessFamily.Tag = "Del" Then
        con.Execute "INSERT INTO ItemProcessGroups_Deletions(ItemID,PG_RefID,UserName,MachineName) VALUES('" & TBox(3) & "'," & Val(cmdBreakLink.Tag) & ",'" & CurrentUserName & "','" & strComputerName & "')"
        con.Execute "DELETE FROM ItemProcessGroups WHERE ItemID='" & ItemCode & "'"
    End If
    
End Sub

Private Sub SaveProcessesForAllItemsOfSameCatagory()
    
    'Save Item's Production Processes...
    Dim iCatID As Integer
    iCatID = Val(cmbCatID.List(cmbCatID.ListIndex, 1))
    
    Call StartTrans(con)
    '------------------------------------------------------------------------
    con.Execute "DELETE FROM ItemProcesses WHERE ItemID IN (SELECT ItemID FROM Items WHERE CatID=" & iCatID & ")"
     
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT ItemID FROM Items WHERE CatID=" & iCatID & " ORDER BY ItemID", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            For i = 1 To LV2.ListItems.count
                With LV2.ListItems(i)
                    con.Execute "INSERT INTO ItemProcesses(ItemID,ProcessID,Scanning,SNO,Rate,MinLevel,MaxLevel,ReOrderLevel) " & _
                     " VALUES('" & rs!ITemID & "'," & Val(.key) & "," & Val(.ListSubItems(1).Tag) & _
                     "," & Val(.SubItems(2)) & "," & Val(.SubItems(3)) & "," & Val(.SubItems(4)) & _
                     "," & Val(.SubItems(5)) & "," & Val(.SubItems(6)) & ")"
                End With
            Next
            .MoveNEXT
        Loop
    End With
    Set rs = Nothing
    
    'Save Process Group Details....
    If txtProcessFamily.Tag = "New" Then
        con.Execute "DELETE FROM ItemProcessGroups WHERE ItemID IN(SELECT ItemID FROM Items WHERE CatID=" & iCatID & ")"
        con.Execute "INSERT INTO ItemProcessGroups(ItemID,PG_RefID) SELECT ItemID," & Val(cmdBreakLink.Tag) & " FROM Items WHERE CatID=" & iCatID
    ElseIf txtProcessFamily.Tag = "Del" Then
        con.Execute "DELETE FROM ItemProcessGroups WHERE ItemID IN(SELECT ItemID FROM Items WHERE CatID=" & iCatID & ")"
    End If
    
    con.CommitTrans
    
    MsgBox "Done.", vbInformation
End Sub


Private Sub SaveForeignCatal()

    con.Execute "DELETE FROM ItemsCatalogReferences WHERE ItemID IN ('" & ItemCode & "','" & TBox(3) & "')"
    For i = 1 To LVRef.ListItems.count
        With LVRef.ListItems(i)
            con.Execute "INSERT INTO ItemsCatalogReferences(ItemID,CatalogName,Reference) VALUES('" & _
             TBox(3) & "','" & .SubItems(1) & "','" & .SubItems(2) & "')"
        End With
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
                ElseIf c.Name = "cmbSteel" Then
                    RetString = RetString & c.List(c.ListIndex, 1) & ","
                ElseIf c.Name = "cmbItemGroupMain" Or c.Name = "cmbItemGroup" Or c.Name = "cmbItemGroup_ECommerce" Then
                    If c.MatchFound Then
                        RetString = RetString & c.List(c.ListIndex, 1) & ","
                    Else
                        RetString = RetString & "0,"
                    End If
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
                RetString = RetString & "'" & c.ID & "',"
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
                ElseIf c.Name = "cmbSteel" Then
                    RetString = RetString & c.Tag & "=" & c.List(c.ListIndex, 1) & ","
                ElseIf c.Name = "cmbItemGroupMain" Or c.Name = "cmbItemGroup" Or c.Name = "cmbItemGroup_ECommerce" Then
                    If c.MatchFound Then
                        RetString = RetString & c.Tag & "=" & c.List(c.ListIndex, 1) & ","
                    Else
                        RetString = RetString & c.Tag & "=0,"
                    End If
                ElseIf c.Name = "cmbGuage" Then
                    RetString = RetString & c.Tag & "='" & c.List(c.ListIndex, 1) & "',"
                Else
                    RetString = RetString & c.Tag & "='" & c.Text & "',"
                End If
            ElseIf (TypeOf c Is MSForms.CheckBox) Then
                RetString = RetString & c.Tag & "=" & Abs(c.Value) & ","
            ElseIf (TypeOf c Is ComboList.Usercontrol1) Then
                RetString = RetString & c.Tag & "='" & c.ID & "',"
            End If
        End If
    Next
    
    If NotEdited = False Then
        
        'Get Rid of Zeros.
        Dim iZeros As Integer
        iZeros = 1
        Do While Mid(TBox(3), iZeros, 1) = "0"
            iZeros = iZeros + 1
        Loop
        
        Dim strRevID As String
        'strRevID = Format(Now, "ddMMyyyyHHmm")
        strRevID = Format(Now, "ddMMyyyy") & Mid(TBox(3), iZeros)
        
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
    con.Execute "DELETE FROM EuroToUS"
    con.Execute "INSERT INTO EuroToUS(EuroToUS) VALUES(" & Val(txtConversion) & ")"
    NBox(29) = Val(NBox(0)) * dConversionRate    'FOB Karachi
    NBox(28) = Val(NBox(14)) * dConversionRate    'Exworks
    NBox(27) = Val(NBox(15)) * dConversionRate    'C&I Sea
    NBox(26) = Val(NBox(16)) * dConversionRate    'C&I Air
    NBox(30) = Val(NBox(8)) * dConversionRate    'CIF Sea
    NBox(31) = Val(NBox(9)) * dConversionRate    'CIF Air
    NBox(32) = Val(NBox(10)) * dConversionRate    'C&F Sea
    NBox(33) = Val(NBox(11)) * dConversionRate    'C&F Air
    '------------------------------------------------------
    NBox(36) = Val(NBox(48)) * dConversionRate
    NBox(43) = Val(NBox(49)) * dConversionRate
    NBox(42) = Val(NBox(50)) * dConversionRate
    NBox(41) = Val(NBox(51)) * dConversionRate
    NBox(37) = Val(NBox(47)) * dConversionRate
    NBox(38) = Val(NBox(46)) * dConversionRate
    NBox(39) = Val(NBox(45)) * dConversionRate
    NBox(40) = Val(NBox(44)) * dConversionRate
    
End Sub

Private Sub txtEditSetQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(txtEditSetQty) <= 0 Then
            MsgBox "Invalid Qty.", vbInformation
            Exit Sub
        End If
        LVSet.SelectedItem.SubItems(4) = Val(txtEditSetQty)
        LVSet.SelectedItem.SubItems(6) = Val(txtEditSetQty) * Val(LVSet.SelectedItem.SubItems(5))
        txtEditSetQty.Visible = False
        Call getTotalSetAmount
    ElseIf KeyAscii = 27 Then
        txtEditSetQty.Visible = False
    End If
End Sub

Private Sub txtEditSetQty_LostFocus()
    txtEditSetQty.Visible = False
End Sub

Private Sub txtProcessSortNo_KeyPress(KeyAscii As Integer)
    Call txtReOrdLvl_KeyPress(13)
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
        If TBox(0) <> .GDevDesc Or cmbUnit <> .GUnit Or TBox(5) <> .GUsage Or cmbType <> .GType Or TBox(1) <> .GAttributes Or TBox(7) <> .GSize Or cmbSizeUnit <> .GSizeUnit Or cmbSteel <> .GSteelUsed Or cmbGuage <> .GGuageaType Or TBox(2) <> .GPacknLabel Then
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
'
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
    strCopyFrom = strItemCode
    'CopyItemCode = strItemCode
    Call ShowData(strItemCode, True)
    TBox(3) = ""
    
    Me.Show 1
    
End Sub

Private Sub ShowData(strItemCode As String, Optional bCopy As Boolean = False)
    
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim strObject As String
    'MsgBox "Starting Show Data"
    bDataLoading = True
    With rs
        .Open "SELECT * FROM VItems WHERE ItemID='" & strItemCode & "'", con, adOpenForwardOnly, adLockReadOnly
        
        For Each c In Controls
            If c.Tag <> "" Then
                strObject = c.Name
                If TypeOf c Is MSForms.CheckBox Then
                    c.Value = .Fields(c.Tag)
                ElseIf TypeOf c Is ComboList.Usercontrol1 Then
                    'C.Text = .Fields(C.Tag & "Text") & ""
                    c.ID = .Fields(c.Tag) & ""
                ElseIf TypeOf c Is VB.Label Then
                    If ItemCode <> "" Then c.Caption = .Fields(c.Tag) & ""
                ElseIf c.Name = "cmbItemType" Then
                    For i = 0 To cmbItemType.ListCount - 1
                        If cmbItemType.List(i, 1) = Val(.Fields(c.Tag) & "") Then
                            c.ListIndex = i
                            Exit For
                        End If
                    Next
                    
                Else
                    If c.Name = "cmbCatID" Then
                        c.Text = .Fields("Category") & ""
                    Else
                        c.Text = .Fields(c.Tag) & ""
                    End If
                End If
                
            End If
        Next
        'MsgBox "AFter  Controls Loop"
'        If strCompany = "Dr-Frgz" Or strCompany = "Tecno" Or strCompany = "Banzai" Then
'            cmbCurr.Text = "US$"
'        Else
'            cmbCurr.Text = "Eur"
'        End If
        'MsgBox "AFter  Currency"
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
        'MsgBox "before Steel"
        cmbSteel = ![SteelType] & ""
        cmbGuage = ![SteelGage] & ""
        cmbItemGroup = ![ItemGroup] & ""
        cmbItemGroup_ECommerce = !ECommerceGroupName & ""
        cmbItemGroupMain = ![MainGroupName] & ""
        'MsgBox "After ItemGroup"
        'TBox(8) = ![CatID] & "-"
        TBox(3) = ![ITemID] & ""
        
        If bCopy = False Then
            If Not IsNull(![ItemPic]) Then
                Set Image1.DataSource = rs
                Image1.DataField = "ItemPic"
                lblPic.Visible = False
            Else
                lblPic.Visible = True
            End If
        End If
        
        If Not IsNull(![TechnicalDrawing]) Then
'            Set ImgTechDrawing.DataSource = rs
'            ImgTechDrawing.DataField = "TechnicalDrawing"
            lblTechDrawing.Caption = "Click here to Open Tech. Drawing."
            lblTechDrawing.Tag = ""
            cmdChangeTechDrawing.Visible = True
            cmdRemoveTechnicalDrawing.Visible = True
        Else
            lblTechDrawing.Caption = "Click here to Add."
            cmdChangeTechDrawing.Visible = False
            cmdRemoveTechnicalDrawing.Visible = False
        End If
                
        If Not IsNull(!PicForECommerce) Then
            Set ImgEC.DataSource = rs
            ImgEC.DataField = "PicForECommerce"
            lblPicEC.Visible = False
        Else
            lblPicEC.Visible = True
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
                TBox(1) = ""
            End If
        End If
        'Add Entries To DMRRev Object For GInfo,Weight and Pricing Tabs...
        '------------------------------------------------------------------
        Call AddRevEntries(rs)
        '------------------------------------------------------------------
        .Close
        
        'Select Item's Production Processes
        '--------------------------------------------------------------------------------------------------
        .Open "SELECT ProcessID,SNo,Description,ItemSNo,Scanning,Code,Rate,MinLevel,MaxLevel,ReOrderLevel,EntryID FROM VItemProcesses WHERE IsExist='" & strItemCode & "' ORDER BY ItemSNo", con, adOpenForwardOnly, adLockReadOnly
        LV2.ListItems.Clear
        Do Until .EOF
            If Not IsNull(![ProcessID]) Then
                Set ITM = LV2.ListItems.add(, !ProcessID & "'", !Description & " (" & !code & ")")
                ITM.Tag = !ProcessID
                ITM.ListSubItems.add(, , IIf(!Scanning, "Yes", "No")).Tag = Abs(!Scanning)
                ITM.ListSubItems.add(, , Val(!ItemSNo & "")).Tag = Val(!EntryID & "")
                ITM.ListSubItems.add , , Val(!Rate & "")
                
                ITM.ListSubItems.add , , Val(!MinLevel & "")
                ITM.ListSubItems.add , , Val(!MaxLevel & "")
                ITM.ListSubItems.add , , Val(!ReOrderLevel & "")
'                LV2.ListItems(![ProcessID] & "'").Checked = True
'                LV2.ListItems(!ProcessID & "'").SubItems(1) = !Sno & ""
                'LV2.ListItems(![ProcessID] & "'").SubItems(1) = Val(![Rate] & "")
                '----------------------------------
'                Dim Obj As New PrRate
'                Obj.PrcID = ![ProcessID]
'                Obj.PrcRate = Val(![Rate] & "")
'                RevEntry.PrdPrIDs.add Obj '.Fields(0).value '![ProcessID]
                '----------------------------------
                Set Obj = Nothing
            End If
            .MoveNEXT
        Loop
        .Close
        Call getNextProcessSeqNo
        
        'Check for Process Groups...
        .Open "SELECT * FROM VItemProcessGroups WHERE ItemID='" & strItemCode & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            cmdBreakLink.Tag = Val(!PG_RefID & "")
            cmdBreakLink.Enabled = True
            txtProcessFamily.Text = !GroupName & ""
            txtProcessFamily.Tag = "Edit"
            cmdDeleteAll.Enabled = False
        End If
        .Close
        '------------------------------------
        cmbProcRM.ClearVals
        cmbProcRM.AddVals con, "Description", "VItemProcesses", "ProcessID", " WHERE IsExist='" & strItemCode & "'"
        '--------------------------------------------------------------------------------------------------
        
        'Select Item's Production Processes
        '--------------------------------------------------------------------------------------------------
        .Open "Select ProcessID From ItemProcessesFDA Where ItemID='" & strItemCode & "'", con, adOpenForwardOnly, adLockReadOnly
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
        
        .Open "Select CatalID,ItemRefID from ItemCatalRef Where ItemID='" & strItemCode & "'", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            If Not IsNull(![CatalID]) Then
                LV3.ListItems(![CatalID] & "'").Checked = True
                LV3.ListItems(![CatalID] & "'").SubItems(1) = ![ItemRefID] & ""
            End If
            .MoveNEXT
        Loop
        .Close
        
        .Open "Select * From VItemsComp Where ItemID='" & strItemCode & "'", con, adOpenForwardOnly, adLockReadOnly
         
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
        
        .Open "Select * From VItemsRMSFUnioned WHERE ItemID='" & strItemCode & "'", con, adOpenForwardOnly, adLockReadOnly
        LVRM.ListItems.Clear
        Do Until .EOF
            
            Set ITM = LVRM.ListItems.add(, , !ProcDesc & "")
            ITM.Tag = ![ProcessID]
            ITM.ListSubItems.add(, , !RMID1 & " (" & !RMName & ")").Tag = !RMID1 & ""
            ITM.ListSubItems.add(, , Val(![Qty] & "")).Tag = IIf(bCopy, 0, Val(!EntryID & ""))
            ITM.ListSubItems.add().Tag = Val(!RM & "")
            .MoveNEXT
        Loop
        
        .Close
        
        .Open "SELECT * FROM ItemsCatalogReferences WHERE ItemID='" & strItemCode & "'", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
           
            Set ITM = LVRef.ListItems.add(, , .AbsolutePosition)
            ITM.Tag = !EntryID
            ITM.ListSubItems.add , , !CatalogName & ""
            ITM.ListSubItems.add , , !Reference & ""
            
            .MoveNEXT
        Loop
        
        .Close
    End With
    
    Set rs = Nothing
    
    Call ShowShipmentInfo(ItemCode)
    Call ShowSetDetail(strItemCode, bCopy)
    Call ShowItemLookAlikes(strItemCode, bCopy)
    
    If cmbSteel.Text = "" Then If cmbSteel.ListCount > 0 Then cmbSteel.ListIndex = 0
    If cmbItemGroup.Text = "" Then If cmbItemGroup.ListCount > 0 Then cmbItemGroup.ListIndex = 0
    If cmbUnit.Text = "" Then If cmbUnit.ListCount > 0 Then cmbUnit.Text = "Pcs"
        
    bDataLoading = False
    Exit Sub
err:
    MsgBox err.Description & vbNewLine '& " Object Name:" & strObject
    bDataLoading = False
End Sub
Private Sub ShowItemLookAlikes(strItemCode As String, Optional bCopy As Boolean = False)

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "SELECT * FROM VItemLookalikes WHERE ItemID='" & strItemCode & "' ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LVLookalike.ListItems.Clear
        Do Until .EOF
            Set ITM = LVLookalike.ListItems.add(, , .AbsolutePosition)
            ITM.Tag = IIf(bCopy, 0, Val(!EntryID & ""))
            ITM.ListSubItems.add(, , "{" & !ItemID_LookAlike & "} " & !ItemName & "").Tag = !ItemID_LookAlike & ""
            ITM.ListSubItems.add , , Val(!InHand & "")
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub


Private Sub ShowSetDetail(strItemCode As String, Optional bCopy As Boolean = False)

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "SELECT * FROM VItemsSets WHERE ItemID='" & strItemCode & "' ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LVSet.ListItems.Clear
        Do Until .EOF
            Set ITM = LVSet.ListItems.add(, , .AbsolutePosition)
            ITM.Tag = IIf(bCopy, 0, Val(!EntryID & ""))
            ITM.ListSubItems.add(, , !Set_ItemID & "").Tag = !Set_ItemID & ""
            ITM.ListSubItems.add , , !SetItemName & ""
            ITM.ListSubItems.add , , !ItemSize & ""
            ITM.ListSubItems.add , , Val(!Qty & "")
            ITM.ListSubItems.add , , Val(!ExWorks & "")
            ITM.ListSubItems.add , , Val(!Qty & "") * Val(!ExWorks & "")
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    Call getTotalSetAmount
    
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

Private Sub txtRef_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If cmbForeignCatalogName.MatchFound = False Then
            MsgBox "Please Select Catalog.", vbInformation
            Exit Sub
        End If
        'Check for Duplicate...
        Dim lCount As Long
        lCount = GetSingleLongValue("COUNT(*)", "ItemsCatalogReferences", " WHERE CatalogName='" & cmbForeignCatalogName.Text & "' AND Reference='" & txtRef & "'")
        
        If lCount > 0 Then
            MsgBox "Already Exists.", vbInformation
            Exit Sub
        End If
        
        Dim ITM As ListItem
        Set ITM = LVRef.ListItems.add(, , LVRef.ListItems.count + 1)
        ITM.ListSubItems.add , , cmbForeignCatalogName.Text
        ITM.ListSubItems.add , , txtRef
        
        cmbForeignCatalogName = ""
        txtRef = ""
        cmbForeignCatalogName.SetFocus
    End If
    
End Sub

Private Sub txtReOrdLvl_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        If cmdBreakLink.Enabled Then
            MsgBox "Can't Add/Edit Processes, Process Family is attached.", vbInformation
            Exit Sub
        End If
        'Add it in the Box Below
        Dim i As Integer, ITM As ListItem
        If cmbProcess.ID <> "0" Then
'            If cmbScanning.MatchFound = False Then
'                MsgBox "Please Select Scanning.", vbInformation
'                Exit Sub
'            End If
            For i = 1 To LV2.ListItems.count
                If Val(LV2.ListItems(i).Tag) = Val(cmbProcess.ID) Then
                    MsgBox "This Process is already assigned.", vbInformation
                    Exit Sub
                End If
            Next
            Dim iPlacementIndex As Integer
            If Val(txtProcessSortNo) > LV2.ListItems.count Then
                iPlacementIndex = LV2.ListItems.count + 1
            Else
                iPlacementIndex = Val(txtProcessSortNo)
            End If
            Set ITM = LV2.ListItems.add(iPlacementIndex, cmbProcess.ID & "'", cmbProcess.Text)
            ITM.Tag = cmbProcess.ID
            ITM.ListSubItems.add(, , "").Tag = 0
            ITM.ListSubItems.add , , (Val(txtProcessSortNo))
'            ITM.ListSubItems.add , , ""
'            ITM.ListSubItems.add , , (Val(txtMinLvl))
'            ITM.ListSubItems.add , , (Val(txtMaxLvl))
'            ITM.ListSubItems.add , , (Val(txtReOrdLvl))
            Dim iNextIndex As Integer
            iNextIndex = Val(txtProcessSortNo) + 1
            For i = (iPlacementIndex + 1) To LV2.ListItems.count
                If Val(LV2.ListItems(i).SubItems(2)) >= Val(txtProcessSortNo) Then
                    LV2.ListItems(i).SubItems(2) = iNextIndex 'Val(LV2.ListItems(i).SubItems(2)) + 1
                    iNextIndex = iNextIndex + 1
                End If
            Next
            
        Else
            'Add Group..............
            If cmdBreakLink.Enabled = False Then 'No Process Group Assigned.
                If cmbProcessGroup.ID <> "0" And cmbProcess.ID = "0" And LV2.ListItems.count = 0 Then
                    'Adding Process Group....
                    cmdBreakLink.Tag = cmbProcessGroup.ID
                    cmdBreakLink.Enabled = True
                    txtProcessFamily.Text = cmbProcessGroup.Text
                    txtProcessFamily.Tag = "New"
                    cmdDeleteAll.Enabled = False
                End If
            End If
            '----------------------------------------------------------------------------------------
            Dim strAssigned As String
            For i = 1 To LV2.ListItems.count
                strAssigned = strAssigned & Val(LV2.ListItems(i).Tag) & ","
            Next
            If strAssigned <> "" Then
                strAssigned = Left(strAssigned, Len(strAssigned) - 1)
            End If
            Dim rs As New ADODB.Recordset
            With rs
                If strAssigned = "" Then
                    .Open "SELECT * FROM VProcessGroups WHERE Group_RefID=" & cmbProcessGroup.ID & " ORDER BY SeqNo", con, adOpenForwardOnly, adLockReadOnly
                Else
                    .Open "SELECT * FROM Processes WHERE ProcessID NOT IN(" & strAssigned & ") AND ProcessID IN(SELECT Process_RefID FROM ProcessGroupsProcesses WHERE Group_RefID=" & cmbProcessGroup.ID & ")", con, adOpenForwardOnly, adLockReadOnly
                End If
                Do Until .EOF
                    Set ITM = LV2.ListItems.add(, !Process_RefID & "'", !Description & "")
                    ITM.Tag = cmbProcess.ID
                    ITM.ListSubItems.add(, , IIf(!Scanning, "Yes", "No")).Tag = Abs(!Scanning)
                    ITM.ListSubItems.add , , (Val(txtProcessSortNo)) + Val(.AbsolutePosition) - 1
'                    ITM.ListSubItems.add , , ""
'                    ITM.ListSubItems.add , , Val(txtMinLvl)
'                    ITM.ListSubItems.add , , Val(txtMaxLvl)
'                    ITM.ListSubItems.add , , Val(txtReOrdLvl)
                    .MoveNEXT
                Loop
                .Close
            End With
            
        End If
        Call getNextProcessSeqNo
        cmbProcess.ID = ""
        cmbProcess.SetFocus
        'cmbProcessGroup.SetFocus
        
    End If
End Sub

Private Sub txtSetQty_KeyPress(KeyAscii As Integer)
     
    If KeyAscii <> 13 Then Exit Sub
    If cmbItemSet.MatchFound = False Then
        MsgBox "Please Select Item.", vbInformation
        Exit Sub
    End If
    
    If Val(txtSetQty) <= 0 Then
        MsgBox "Invalid Qty.", vbInformation
        Exit Sub
    End If
    
    Dim i As Integer
    For i = 1 To LVSet.ListItems.count
        If LVSet.ListItems(i).ListSubItems(1).Tag = cmbItemSet.ID Then
            MsgBox "Already Added.", vbInformation
            Exit Sub
        End If
    Next
     
    Dim ITM As ListItem
    Set ITM = LVSet.ListItems.add(, , LVSet.ListItems.count + 1)
    ITM.Tag = "0"
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT ItemID,ItemName,ItemSize,ExWorks FROM Items WHERE ItemID='" & cmbItemSet.ID & "'", con, adOpenForwardOnly, adLockReadOnly
     
    ITM.ListSubItems.add(, , cmbItemSet.ID).Tag = cmbItemSet.ID
    ITM.ListSubItems.add , , rs!ItemName & ""
    ITM.ListSubItems.add , , rs!ItemSize & ""
    ITM.ListSubItems.add , , Val(txtSetQty)
    ITM.ListSubItems.add , , Val(rs!ExWorks & "")
    ITM.ListSubItems.add , , Val(rs!ExWorks & "") * Val(txtSetQty)
    '---------------------------------------------------------------
    txtSetQty = ""
    cmbItemSet.ID = ""
    cmbItemSet.SetFocus
    
    Call getTotalSetAmount
End Sub

Private Sub getTotalSetAmount()

    Dim i As Integer, dAmt As Double
    For i = 1 To LVSet.ListItems.count
        dAmt = dAmt + Val(LVSet.ListItems(i).SubItems(6))
    Next
    txtTotalSetAmt = dAmt
    
End Sub
Private Sub txtSortNo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        
        If Val(txtSortNo) > LV2.ListItems.count Then
            MsgBox "Invalid Sort No.", vbInformation
            Exit Sub
        End If
        Dim bPlus As Boolean
        If Val(txtSortNo) < LV2.SelectedItem.SubItems(2) Then
            bPlus = True
        Else
            bPlus = False
        End If
        txtSortNo.Visible = False
        LV2.SelectedItem.SubItems(2) = Val(txtSortNo)
                
        Dim ITM As ListItem, myITM As ListItem
        Set ITM = LV2.SelectedItem
        Dim iNewIndex As Integer
        If bPlus Then
            iNewIndex = Val(txtSortNo)
        Else
            iNewIndex = Val(txtSortNo) + 1
        End If
        Set myITM = LV2.ListItems.add(iNewIndex, , ITM.Text)
        myITM.Tag = ITM.Tag
        myITM.ListSubItems.add(, , ITM.SubItems(1)).Tag = ITM.ListSubItems(1).Tag
        myITM.ListSubItems.add(, , ITM.SubItems(2)).Tag = ITM.ListSubItems(2).Tag

        LV2.ListItems.Remove LV2.SelectedItem.Index
        myITM.key = ITM.key
'
'        Dim i As Integer, iNextIndex As Integer
'        iNextIndex = myITM.Index + 1
'
'        Dim iStartIndex As Integer, iLastIndex As Integer
'
'        If bPlus Then
'            For i = iNextIndex To LV2.ListItems.count
'                If Val(LV2.ListItems(i).SubItems(2)) >= Val(txtSortNo) And i <> myITM.Index Then
'                    LV2.ListItems(i).SubItems(2) = i 'Val(LV2.ListItems(i).SubItems(2)) + 1
'                End If
'            Next
'        Else
'            For i = LV2.ListItems.count To iNextIndex Step -1
'                If Val(LV2.ListItems(i).SubItems(2)) >= Val(txtSortNo) And i <> myITM.Index Then
'                    LV2.ListItems(i).SubItems(2) = i - 1 'Val(LV2.ListItems(i).SubItems(2)) + 1
'                End If
'            Next
'        End If
        Call ReIndexProcesses
        
        
    ElseIf KeyAscii = 27 Then
        txtSortNo.Visible = False
    End If
    
End Sub

Private Sub ReIndexProcesses()

    'First Move All to LV2Hidden
    LV2.Visible = False
    Dim i As Integer
    Dim FromITM As ListItem, ToITM As ListItem
    For i = 1 To LV2.ListItems.count
        LV2.ListItems(i).SubItems(2) = i
        'Set FromITM = LV2.ListItems(i)
        'set toitm=lv2hidden.ListItems.Add(
    Next
    LV2.Visible = True
End Sub
Private Sub txtSortNo_LostFocus()
    txtSortNo.Visible = False
End Sub

Public Function AddNew(Optional m_bSample As Boolean = False) As String
    
    bSample = m_bSample
    If strCompany = "Weldon Instruments" Then
        Dim i As Integer
        For i = 1 To LV2.ListItems.count
            LV2.ListItems(i).Checked = True
            LV2.ListItems(i).SubItems(1) = i
        Next
        cmbCatID.Text = "Surgical Instruments"
    End If
    Me.Show 1
        
    AddNew = strNewItemCode
    
End Function

Private Sub txtStRate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LV2.SelectedItem.SubItems(iColNoProcess - 1) = Val(txtStRate)
        txtStRate.Visible = False
    ElseIf KeyAscii = 27 Then
        txtStRate.Visible = False
    End If
End Sub

Private Sub txtStRate_LostFocus()
    txtStRate.Visible = False
End Sub

Private Sub txtUrduFields_GotFocus(Index As Integer)
    Call SetUrduLanguage
End Sub

Private Sub txtUrduFields_LostFocus(Index As Integer)
    Call SetEnglishLanguage
End Sub

Private Sub getNextProcessSeqNo()
    If LV2.ListItems.count > 0 Then
        txtProcessSortNo = Val(LV2.ListItems(LV2.ListItems.count).SubItems(2)) + 1
    Else
        txtProcessSortNo = 1
    End If
End Sub
