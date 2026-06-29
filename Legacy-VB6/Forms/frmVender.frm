VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmVender 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vender Info."
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9075
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fTitle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   150
      TabIndex        =   8
      Top             =   0
      Width           =   8700
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vender Info."
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
         Index           =   0
         Left            =   60
         TabIndex        =   9
         Top             =   120
         Width           =   8580
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vender Info."
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
         Index           =   1
         Left            =   60
         TabIndex        =   10
         Top             =   135
         Width           =   8580
      End
   End
   Begin VB.Frame FAll 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6225
      Left            =   60
      TabIndex        =   4
      Top             =   600
      Width           =   8925
      Begin VB.CheckBox chkOwnRep 
         Caption         =   "Own Repair"
         DataField       =   "Inactive"
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
         Height          =   270
         Left            =   1845
         TabIndex        =   72
         Tag             =   "OwnRepair"
         Top             =   5190
         Width           =   1950
      End
      Begin VB.CheckBox chkRepVend 
         Caption         =   "Repair Vender"
         DataField       =   "Inactive"
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
         Height          =   270
         Left            =   180
         TabIndex        =   71
         Tag             =   "RepairVend"
         Top             =   5190
         Width           =   1575
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Index           =   12
         Left            =   3900
         MaxLength       =   255
         TabIndex        =   56
         Tag             =   "VendID1"
         Top             =   435
         Width           =   1830
      End
      Begin VB.CheckBox chkInactive 
         Caption         =   "Vender is inactive"
         DataField       =   "Inactive"
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
         Height          =   270
         Left            =   180
         TabIndex        =   49
         Top             =   5505
         Width           =   2040
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Index           =   1
         Left            =   180
         MaxLength       =   255
         TabIndex        =   0
         Tag             =   "VenderName"
         Top             =   435
         Width           =   3690
      End
      Begin VB.TextBox txtfields 
         DataField       =   "VendID"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   195
         TabIndex        =   5
         Top             =   435
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   3495
         TabIndex        =   11
         Top             =   5505
         Width           =   5280
         Begin MSForms.CommandButton cmdSaveClose 
            Height          =   360
            Left            =   1800
            TabIndex        =   2
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
         Begin MSForms.CommandButton cmdCancel 
            CausesValidation=   0   'False
            Height          =   360
            Left            =   3510
            TabIndex        =   3
            Top             =   195
            Width           =   1665
            Caption         =   "Cancel            "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmVender.frx":0000
            Accelerator     =   67
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdSave 
            Height          =   360
            Left            =   75
            TabIndex        =   1
            Top             =   195
            Width           =   1665
            Caption         =   "Save & New   "
            PicturePosition =   327683
            Size            =   "2937;635"
            Accelerator     =   110
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4260
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   8640
         _ExtentX        =   15240
         _ExtentY        =   7514
         _Version        =   393216
         Tabs            =   4
         Tab             =   1
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "General Info."
         TabPicture(0)   =   "frmVender.frx":0112
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame1"
         Tab(0).Control(1)=   "Image2"
         Tab(0).Control(2)=   "Label2"
         Tab(0).Control(3)=   "Label1(9)"
         Tab(0).Control(4)=   "Image1"
         Tab(0).Control(5)=   "Label1(32)"
         Tab(0).Control(6)=   "lblPic"
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "Detail"
         TabPicture(1)   =   "frmVender.frx":012E
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "lblLabels(15)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "cmbType"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "lblLabels(16)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "cmbProcess"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "cmbProcess1"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "lblLabels(18)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "chkComp2"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "chkComp1"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "lblLabels(19)"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "lblItems"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "cmdAssign"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "LV"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "txtfields(14)"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "txtfields(16)"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "txtfields(17)"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "txtfields(18)"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "cmbItem"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).Control(17)=   "cmbItemName"
         Tab(1).Control(17).Enabled=   0   'False
         Tab(1).Control(18)=   "txtfields(0)"
         Tab(1).Control(18).Enabled=   0   'False
         Tab(1).ControlCount=   19
         TabCaption(2)   =   "Referal Info"
         TabPicture(2)   =   "frmVender.frx":014A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "FRefInfo"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Sub Vendors"
         TabPicture(3)   =   "frmVender.frx":0166
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "chkPlanter"
         Tab(3).Control(1)=   "chkSnaffer"
         Tab(3).Control(2)=   "chkStamper"
         Tab(3).Control(3)=   "cmbSnaffer"
         Tab(3).Control(4)=   "cmbStamper"
         Tab(3).Control(5)=   "cmbPlanter"
         Tab(3).ControlCount=   6
         Begin VB.TextBox txtfields 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   6525
            MaxLength       =   10
            TabIndex        =   80
            Top             =   1830
            Width           =   1125
         End
         Begin ComboList.Usercontrol1 cmbItemName 
            Height          =   285
            Left            =   1905
            TabIndex        =   75
            Top             =   1830
            Width           =   4650
            _ExtentX        =   8202
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
            Left            =   150
            TabIndex        =   76
            Top             =   1830
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
         Begin VB.TextBox txtfields 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Index           =   18
            Left            =   6825
            MaxLength       =   10
            TabIndex        =   73
            Tag             =   "RepairDedRate"
            Top             =   675
            Width           =   1590
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "Phone1"
            Height          =   315
            Index           =   17
            Left            =   5100
            MaxLength       =   20
            TabIndex        =   70
            Tag             =   "OtherCompany2"
            Top             =   1245
            Visible         =   0   'False
            Width           =   3315
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "Phone1"
            Height          =   315
            Index           =   16
            Left            =   1770
            MaxLength       =   20
            TabIndex        =   69
            Tag             =   "OtherCompany1"
            Top             =   1245
            Visible         =   0   'False
            Width           =   3300
         End
         Begin VB.TextBox txtfields 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Index           =   14
            Left            =   150
            MaxLength       =   10
            TabIndex        =   65
            Tag             =   "Experience"
            Top             =   1245
            Width           =   1590
         End
         Begin VB.Frame FRefInfo 
            Height          =   2370
            Left            =   -73755
            TabIndex        =   38
            Top             =   825
            Width           =   5850
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Address"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   28
               Left            =   2370
               TabIndex        =   48
               Top             =   945
               Width           =   585
            End
            Begin MSForms.TextBox TBox 
               Height          =   990
               Index           =   21
               Left            =   2340
               TabIndex        =   47
               Tag             =   "RefByAddress"
               Top             =   1170
               Width           =   3045
               VariousPropertyBits=   -1467987941
               BorderStyle     =   1
               ScrollBars      =   2
               Size            =   "5371;1746"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Phone 2"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   27
               Left            =   495
               TabIndex        =   46
               Top             =   1635
               Width           =   585
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   20
               Left            =   465
               TabIndex        =   45
               Tag             =   "RefByPhone2"
               Top             =   1845
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
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Phone 1"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   26
               Left            =   510
               TabIndex        =   44
               Top             =   960
               Width           =   585
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   19
               Left            =   465
               TabIndex        =   43
               Tag             =   "RefByPhone1"
               Top             =   1170
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
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Father's Name"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   25
               Left            =   2940
               TabIndex        =   42
               Top             =   300
               Width           =   1035
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   18
               Left            =   2940
               TabIndex        =   41
               Tag             =   "RefByFName"
               Top             =   540
               Width           =   2445
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "4313;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   10
               Left            =   465
               TabIndex        =   40
               Tag             =   "RefBy"
               Top             =   540
               Width           =   2445
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "4313;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Refered By (Name)"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   15
               Left            =   510
               TabIndex        =   39
               Top             =   315
               Width           =   1380
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Contact Information :"
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
            Height          =   3720
            Left            =   -74910
            TabIndex        =   13
            Top             =   420
            Width           =   4920
            Begin VB.TextBox txtfields 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               DataField       =   "Fax1"
               DataSource      =   "datPrimaryRS"
               Height          =   315
               Index           =   9
               Left            =   2205
               MaxLength       =   100
               TabIndex        =   53
               Tag             =   "MaxLimit"
               Top             =   1695
               Width           =   1305
            End
            Begin VB.TextBox txtfields 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               DataField       =   "Fax2"
               DataSource      =   "datPrimaryRS"
               Height          =   315
               Index           =   5
               Left            =   3540
               MaxLength       =   100
               TabIndex        =   52
               Tag             =   "TimeLimit"
               Top             =   1695
               Width           =   1290
            End
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "Mobile"
               Height          =   315
               Index           =   13
               Left            =   90
               MaxLength       =   20
               TabIndex        =   32
               Tag             =   "Mobile"
               Top             =   1065
               Width           =   1560
            End
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "Fax1"
               DataSource      =   "datPrimaryRS"
               Height          =   315
               Index           =   15
               Left            =   90
               MaxLength       =   100
               TabIndex        =   30
               Tag             =   "NICNo"
               Top             =   1695
               Width           =   2085
            End
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "Phone1"
               Height          =   315
               Index           =   3
               Left            =   90
               MaxLength       =   20
               TabIndex        =   19
               Tag             =   "Phone1"
               Top             =   450
               Width           =   1560
            End
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "Phone2"
               DataSource      =   "datPrimaryRS"
               Height          =   315
               Index           =   4
               Left            =   1680
               MaxLength       =   20
               TabIndex        =   18
               Tag             =   "Phone2"
               Top             =   450
               Width           =   1560
            End
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "Phone3"
               DataSource      =   "datPrimaryRS"
               Height          =   315
               Index           =   6
               Left            =   3270
               MaxLength       =   20
               TabIndex        =   17
               Tag             =   "Phone3"
               Top             =   450
               Width           =   1560
            End
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "Fax2"
               DataSource      =   "datPrimaryRS"
               Height          =   315
               Index           =   7
               Left            =   3270
               MaxLength       =   100
               TabIndex        =   16
               Tag             =   "Fax2"
               Top             =   1065
               Width           =   1560
            End
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "Address"
               DataSource      =   "datPrimaryRS"
               Height          =   1320
               Index           =   8
               Left            =   90
               MaxLength       =   255
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   15
               Tag             =   "Address"
               Top             =   2310
               Width           =   4740
            End
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "Fax1"
               DataSource      =   "datPrimaryRS"
               Height          =   315
               Index           =   2
               Left            =   1680
               MaxLength       =   100
               TabIndex        =   14
               Tag             =   "Fax1"
               Top             =   1065
               Width           =   1560
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Time Limit (days)"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   10
               Left            =   3540
               TabIndex        =   55
               Top             =   1470
               Width           =   1200
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Max. Limit"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   5
               Left            =   2235
               TabIndex        =   54
               Top             =   1470
               Width           =   720
            End
            Begin MSForms.OptionButton NIC 
               Height          =   285
               Index           =   1
               Left            =   570
               TabIndex        =   51
               Top             =   1410
               Width           =   600
               BackColor       =   -2147483633
               ForeColor       =   8388608
               DisplayStyle    =   5
               Size            =   "1058;503"
               Value           =   "0"
               Caption         =   "Old"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.OptionButton NIC 
               Height          =   285
               Index           =   0
               Left            =   1200
               TabIndex        =   50
               Top             =   1410
               Width           =   870
               BackColor       =   -2147483633
               ForeColor       =   8388608
               DisplayStyle    =   5
               Size            =   "1535;503"
               Value           =   "0"
               Caption         =   "New"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mobile"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   13
               Left            =   90
               TabIndex        =   33
               Top             =   825
               Width           =   450
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "NIC #"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   14
               Left            =   90
               TabIndex        =   31
               Top             =   1440
               Width           =   435
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fax 1"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   2
               Left            =   1710
               TabIndex        =   25
               Top             =   840
               Width           =   405
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Phone1"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   3
               Left            =   90
               TabIndex        =   24
               Top             =   225
               Width           =   540
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Phone2"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   4
               Left            =   1710
               TabIndex        =   23
               Top             =   240
               Width           =   540
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Phone 3"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   6
               Left            =   3285
               TabIndex        =   22
               Top             =   255
               Width           =   585
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fax 2"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   8
               Left            =   3285
               TabIndex        =   21
               Top             =   840
               Width           =   405
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   9
               Left            =   90
               TabIndex        =   20
               Top             =   2070
               Width           =   585
            End
         End
         Begin ComboList.Usercontrol1 cmbPlanter 
            Height          =   285
            Left            =   -73740
            TabIndex        =   62
            Tag             =   "Planter"
            Top             =   720
            Visible         =   0   'False
            Width           =   3120
            _ExtentX        =   5503
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
         Begin ComboList.Usercontrol1 cmbStamper 
            Height          =   285
            Left            =   -73740
            TabIndex        =   63
            Tag             =   "Stamper"
            Top             =   1185
            Visible         =   0   'False
            Width           =   3120
            _ExtentX        =   5503
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
         Begin ComboList.Usercontrol1 cmbSnaffer 
            Height          =   285
            Left            =   -73725
            TabIndex        =   61
            Tag             =   "Snaffer"
            Top             =   1635
            Visible         =   0   'False
            Width           =   3120
            _ExtentX        =   5503
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
            Height          =   2010
            Left            =   150
            TabIndex        =   79
            Top             =   2130
            Width           =   8265
            _ExtentX        =   14579
            _ExtentY        =   3545
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
               Text            =   "S No."
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Item Code"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Description"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Rate"
               Object.Width           =   1764
            EndProperty
         End
         Begin MSForms.CommandButton cmdAssign 
            Height          =   525
            Left            =   7665
            TabIndex        =   77
            Top             =   1590
            Width           =   750
            ForeColor       =   6244673
            Caption         =   "Add"
            PicturePosition =   327683
            Size            =   "1323;926"
            Picture         =   "frmVender.frx":0182
            Accelerator     =   65
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.Label lblItems 
            Height          =   240
            Left            =   150
            TabIndex        =   78
            Top             =   1605
            Width           =   7500
            BackColor       =   11517387
            Caption         =   " Device Code               Device Description                                                                        Rate"
            Size            =   "13229;423"
            BorderColor     =   16512
            BorderStyle     =   1
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Repair Ded. Rate"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   19
            Left            =   6825
            TabIndex        =   74
            Top             =   450
            Width           =   1245
         End
         Begin MSForms.CheckBox chkComp1 
            Height          =   345
            Left            =   1770
            TabIndex        =   68
            Top             =   975
            Width           =   1815
            VariousPropertyBits=   1015023643
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3201;609"
            Value           =   "0"
            Caption         =   "1st Other Company"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkComp2 
            Height          =   345
            Left            =   5085
            TabIndex        =   67
            Top             =   975
            Width           =   1860
            VariousPropertyBits=   1015023643
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3281;609"
            Value           =   "0"
            Caption         =   "2nd Other Company"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Experience (In Years)"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   18
            Left            =   165
            TabIndex        =   66
            Top             =   1020
            Width           =   1560
         End
         Begin MSForms.ComboBox cmbProcess1 
            Height          =   300
            Left            =   3045
            TabIndex        =   64
            Tag             =   "SubVendType"
            Top             =   675
            Visible         =   0   'False
            Width           =   3750
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "6615;529"
            ListWidth       =   8819
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkStamper 
            Height          =   315
            Left            =   -74670
            TabIndex        =   60
            Top             =   1140
            Width           =   1005
            VariousPropertyBits=   1015023643
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "1773;556"
            Value           =   "0"
            Caption         =   "Stamper"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkSnaffer 
            Height          =   315
            Left            =   -74670
            TabIndex        =   59
            Top             =   1635
            Width           =   945
            VariousPropertyBits=   1015023643
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "1667;556"
            Value           =   "0"
            Caption         =   "Snaffer"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkPlanter 
            Height          =   315
            Left            =   -74670
            TabIndex        =   58
            Top             =   675
            Width           =   915
            VariousPropertyBits=   1015023643
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "1614;556"
            Value           =   "0"
            Caption         =   "Planter"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   1260
            Left            =   -69870
            Stretch         =   -1  'True
            Top             =   2880
            Width           =   1545
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Double Click To Add Thumb Exp. /Signature"
            ForeColor       =   &H00FF0000&
            Height          =   615
            Left            =   -69705
            TabIndex        =   37
            Top             =   3255
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Thumb Exp/Signature"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   -69870
            TabIndex        =   36
            Top             =   2655
            Width           =   1545
         End
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   1740
            Left            =   -69870
            Stretch         =   -1  'True
            Top             =   630
            Width           =   1545
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Picture"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   32
            Left            =   -69855
            TabIndex        =   35
            Top             =   405
            Width           =   495
         End
         Begin VB.Label lblPic 
            Alignment       =   2  'Center
            Caption         =   "Double Click To Add Picture."
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   -69645
            TabIndex        =   34
            Top             =   1290
            Width           =   1095
         End
         Begin MSForms.ComboBox cmbProcess 
            Height          =   300
            Left            =   3045
            TabIndex        =   29
            Tag             =   "PhaseID"
            Top             =   675
            Width           =   3750
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "6615;529"
            ListWidth       =   8819
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Process"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   16
            Left            =   3090
            TabIndex        =   28
            Top             =   465
            Width           =   555
         End
         Begin MSForms.ComboBox cmbType 
            Height          =   300
            Left            =   150
            TabIndex        =   27
            Tag             =   "VendType"
            Top             =   675
            Width           =   2880
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5080;529"
            ListWidth       =   5291
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vender Type"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   15
            Left            =   165
            TabIndex        =   26
            Top             =   465
            Width           =   915
         End
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   5805
         Top             =   255
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DialogTitle     =   "Please Select Employee's Picture"
         Filter          =   "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor ID"
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
         Height          =   195
         Index           =   17
         Left            =   3885
         TabIndex        =   57
         Top             =   225
         Width           =   840
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Name"
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
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   225
         Width           =   1125
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vender ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmVender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim rsVend As New ADODB.Recordset
Public Add As Boolean, Edit As Boolean, View
Dim Acno As String
Dim iTotNo As Integer
Dim VendID As Integer
Dim VendID1 As String
Dim OpBal As Double
Dim PicFileName As String
Dim ThumbFileName As String
Dim ShowingData As Boolean
Dim CopingVender As Boolean
Dim FromVendID As Integer
Public Sub EditVender(TempVendID As Integer)
    On Error GoTo err
    
    ShowingData = True
    
    
    VendID = TempVendID
    Call showdata
    Call ShowItems
    ShowingData = False
    Me.Show 1
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub ShowItems()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "Select * From VVendAssItemsOthers Where VendID=" & VendID, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.Add(, !ItemCode & "'", LV.ListItems.Count + 1)
            ITM.SubItems(1) = !ItemCode & ""
            ITM.SubItems(2) = !CompleteItemName & ""
            ITM.SubItems(3) = !ItemRate & ""
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
Public Sub CopyVender(TempVendID As Integer)

    On Error GoTo err
    
    CopingVender = True
    ShowingData = True
        
    
    VendID = TempVendID
    FromVendID = TempVendID
    Call showdata
    VendID = 0
    
    ShowingData = False
    Call GetVendID
    Me.Show 1
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub chkComp1_Click()
    txtfields(16).Visible = chkComp1.value
End Sub

Private Sub chkComp2_Click()
    txtfields(17).Visible = chkComp2.value
End Sub

Private Sub chkInactive_Click()

    If chkinactive = vbChecked And Val(lblBalance) <> 0 Then
        MsgBox "This Vender Cannot Be Made Inactive," & vbNewLine & _
         "To Make A Vender Inactive Balance Must Be Zero", vbInformation
        chkinactive = vbUnchecked
    End If
    
End Sub


Private Sub chkOwnRep_Click()
    If chkOwnRep.value = vbChecked Then
        txtfields(18).Text = 0
        txtfields(18).Locked = True
    Else
        txtfields(18).Locked = False
    End If
End Sub

Private Sub chkPlanter_Click()
    
    cmbPlanter.Visible = chkPlanter.value
    If Not chkPlanter.value Then cmbPlanter.ID = 0
    
End Sub

Private Sub chkRepVend_Click()
    If chkRepVend.value = vbChecked Then
        cmbType.ListIndex = 1
        cmbType.Enabled = False
        cmbProcess.Enabled = False
        chkOwnRep.value = vbUnchecked
        chkOwnRep.Enabled = False
        txtfields(18).Visible = False
        lblLabels(19).Visible = False
        Call GetRepairVendID
    Else
        cmbType.Enabled = True
        chkOwnRep.Enabled = True
        cmbProcess.Enabled = True
        txtfields(18).Visible = True
        lblLabels(19).Visible = True
    End If
End Sub

Private Sub chkSnaffer_Click()
    cmbSnaffer.Visible = chkSnaffer.value
    If Not chkSnaffer.value Then cmbSnaffer.ID = 0
End Sub

Private Sub chkStamper_Click()
    cmbStamper.Visible = chkStamper.value
    If Not chkStamper.value Then cmbStamper.ID = 0
End Sub

Private Sub cmbDrCr_Click()

    If cmbDrCr.ListIndex = 1 Then
        'txtfields(11).Tag = -Val(txtfields(11))
        OpBal = -Val(txtfields(11))
    Else
        'txtfields(11).Tag = Val(txtfields(11))
        OpBal = Val(txtfields(11))
    End If
    
End Sub

Private Sub cmdAddItem_Click()

    On Error GoTo err
    Dim ITM As ListItem
    
    'Set ITM = LV.ListItems.Add(, txtfields(0) & "", LV.ListItems.Count + 1)
    Set ITM = LV.ListItems.Add(, cmbType & cmbProcess.List(cmbProcess.ListIndex, 1), LV.ListItems.Count + 1)
    ITM.ListSubItems.Add , , txtfields(14) & ""
    ITM.ListSubItems.Add , , cmbType.Text & ""
    ITM.ListSubItems.Add , cmbProcess.List(cmbProcess.ListIndex, 1) & "'", cmbProcess.Text & ""
    ITM.ListSubItems.Add , , txtfields(0) & ""
    ITM.ListSubItems.Add(, , OpBal).Tag = Format(DtOpenDate, "dd-MMM-yyyy")
    
    iTotNo = iTotNo + 1
    txtfields(11) = ""
    'cmbDrCr = ""
    cmbType.Text = ""
    cmbProcess.Text = ""
    
    Call CmbType_Click
    Exit Sub
err:
    If err.Number = 35602 Then
        MsgBox "Invalid Entry!", vbInformation
        Exit Sub
    Else
        MsgBox err.Description
    End If
    
    
End Sub

Private Sub cmbItem_matched()
    cmbItemName.ID = cmbItem.ID
End Sub

Private Sub cmbItemName_matched()
    cmbItem.ID = cmbItemName.ID
    'txtfields(0).SetFocus
End Sub

Private Sub cmbProcess_Change()
    Call cmbProcess_Click
End Sub

Private Sub cmbProcess_Click()

    Call GetVendID
    
'    cmbItem.ClearVals
'    cmbItemName.ClearVals
'
'    If Val(cmbProcess.List(cmbProcess.ListIndex, 1) & "") > 0 Then
'        cmbItem.AddVals con, "ItemID", "VItems1", "ItemID", " Where ItemID In(Select ItemID From ItemProcesses Where ProcessID=" & cmbProcess.List(cmbProcess.ListIndex, 1) & ")"
'        cmbItemName.AddVals con, "CompleteItemName", "VItems1", "ItemID", " Where ItemID In(Select ItemID From ItemProcesses Where ProcessID=" & cmbProcess.List(cmbProcess.ListIndex, 1) & ")"
'    End If
    
End Sub
Private Sub GetVendID()

    On Error GoTo err
    
    If ShowingData Then Exit Sub
    
    Dim strVendID As String
    If Not cmbType.MatchFound Then Exit Sub
    If cmbProcess.Visible Then
        If Not cmbProcess.MatchFound Then Exit Sub
    End If
    If cmbProcess1.Visible Then
        If Not cmbProcess1.MatchFound Then Exit Sub
    End If
    
    'strVendID = Val(con.Execute("Select Max(Cast(SubString(VendID1,5,4) as int)) From Venders Where UPPER(Left(VendID1,3))=" & UCase(Left(cmbProcess.Text, 3))).Fields(0).value & "")
    
    If VendID <> 0 Then
        If UCase(Left(VendID1, 3)) = UCase(Left(cmbProcess, 3)) Then
            txtfields(12) = VendID1
            Exit Sub
        End If
    End If
    
    If cmbType.ListIndex = "0" Or cmbType.ListIndex = "1" Then
        strVendID = Val(con.Execute("Select Max(Cast(SubString(VendID1,5,4) as int)) From Venders Where PhaseID=" & cmbProcess.List(cmbProcess.ListIndex, 1) & " and VendType='" & cmbType.ListIndex & "'").Fields(0).value & "")
    ElseIf cmbType.ListIndex = "2" Then
        strVendID = Val(con.Execute("Select Max(Cast(SubString(VendID1,5,4) as int)) From Venders Where VendType='" & cmbType.ListIndex & "'").Fields(0).value & "")
    End If
    
    If strVendID = 0 Then
        If cmbType.ListIndex = "0" Or cmbType.ListIndex = "2" Then
            strVendID = 1000
        ElseIf cmbType.ListIndex = "1" Then
            strVendID = 5000
        End If
    End If
    
    strVendID = strVendID + 1
    If cmbType.ListIndex = "0" Or cmbType.ListIndex = "1" Then
        txtfields(12) = UCase(Left(cmbProcess.Text, 3)) & Format(strVendID, "-0000")
    ElseIf cmbType.ListIndex = "2" Then
        txtfields(12) = "SUB" & Format(strVendID, "-0000") 'UCase(Left(cmbProcess.Text, 3)) & Format(strVendID, "-1000")
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmbProcess1_Change()
    Call cmbProcess1_Click
End Sub

Private Sub cmbProcess1_Click()

    Call GetVendID
    
End Sub

Private Sub cmdAssign_Click()

    On Error Resume Next
    
    Dim ITM As ListItem
    
    Set ITM = LV.ListItems.Add(, cmbItem.ID & "'", LV.ListItems.Count + 1)
    
    ITM.SubItems(1) = cmbItem.Text
    ITM.SubItems(2) = cmbItemName.Text
    ITM.SubItems(3) = Val(txtfields(0))
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    On Error GoTo err

    If Trim(txtfields(1)) = "" Then
        MsgBox "Invalid Vender Name", vbInformation
        Exit Sub
    ElseIf cmbType.Enabled And Not cmbType.MatchFound Then
        MsgBox "Please Select Vender Type", vbInformation
        cmbType.SetFocus
        Exit Sub
    End If

    If Add Then
        Call AddAccountInfo
    ElseIf Edit Then
        Call UpdateAccinfo
    End If


    rsVend.Update
    
    Me.Add = True
    Call Form_QueryUnload(0, 0)
    Call ClearFields
    'Call Form_Activate
    txtfields(1).SetFocus
    If cmbType.Enabled Then cmbType = ""
    Exit Sub

err:
    MsgBox "ERROR : # " & err.Number & vbNewLine & err.Description, vbInformation
 
End Sub
Private Function SaveVendor() As Boolean
    
    On Error GoTo err
    Dim rec As New ADODB.Recordset
    
    If Trim(txtfields(1)) = "" Then
        MsgBox "Invalid Vender Name", vbInformation
        Exit Function
    End If
    
    If Not cmbType.MatchFound Then
        MsgBox "No Type Is Selected!", vbInformation
        SSTab1.Tab = 1
        cmbType.SetFocus
        Exit Function
    End If
    
    If chkRepVend.value = vbUnchecked Then
        If cmbProcess.Visible Then
            If Not cmbProcess.MatchFound Then
                MsgBox "No Process Is Selected!", vbInformation
                SSTab1.Tab = 1
                cmbProcess.SetFocus
                Exit Function
            End If
        End If
    
        If cmbProcess1.Visible Then
            If Not cmbProcess1.MatchFound Then
                MsgBox "No Process Is Selected!", vbInformation
                SSTab1.Tab = 1
                cmbProcess1.SetFocus
                Exit Function
            End If
        End If
    End If
    Call StartTrans(con)
    
    Me.MousePointer = vbHourglass
    Dim Sql As String
    
    If VendID = 0 Then
        Sql = CreateInsertSQL("Venders")
    Else
        Sql = CreateUpdateSQL("Venders")
        Sql = Sql & " Where VendID=" & VendID
    End If
    
    
    con.Execute Sql
    If VendID = 0 Then
        rec.Open "Select Max(VendID) From Venders", con, adOpenForwardOnly, adLockReadOnly
        VendID = rec.Fields(0) & ""
        rec.Close
    End If
    Call SaveItems(VendID)
    'If Add Then
    '    Call AddAccountInfo
    'ElseIf Edit Then
    '    Call UpdateAccinfo
    'End If
    If CopingVender Then
        Dim NewVendID As Long
        NewVendID = Val(con.Execute("Select Max(VendID) From Venders").Fields(0).value & "")
        
        Dim cmd As New ADODB.Command
        With cmd
            Set .ActiveConnection = con
            .CommandType = adCmdStoredProc
            .CommandText = "SP_CopyAssItems"
            .Parameters("@VendID").value = NewVendID
            .Parameters("@FromVendID").value = FromVendID
            .Execute
        End With
    End If
    
    con.CommitTrans
    
SavePic:
    If (Image1.Picture = 0 Or Trim(PicFileName) = "") And (Image2.Picture = 0 Or Trim(ThumbFileName) = "") Then GoTo TheEnd

    
    Dim rsPic As New ADODB.Stream
    Dim rsThumb As New ADODB.Stream
    
    
    rec.Open "Select VendID,VendPic,VendThumb from Venders  Where VendID=" & VendID, con, adOpenStatic, adLockOptimistic
    
    If Image1.Picture <> 0 And Trim(PicFileName) <> "" Then
      'Save Picture
        rsPic.Open
        rsPic.Type = adTypeBinary
        rsPic.LoadFromFile PicFileName
        rec.Fields(1) = rsPic.Read
        rsPic.Close
    End If
    
    'Save Thumb Expression
    If Image2.Picture <> 0 And Trim(ThumbFileName) <> "" Then
        rsThumb.Open
        rsThumb.Type = adTypeBinary
        rsThumb.LoadFromFile ThumbFileName
        rec.Fields("VendThumb") = rsThumb.Read
        rsThumb.Close
    End If
    
    rec.Update
    rec.Close
    Set rec = Nothing
    Set rsPic = Nothing
    Set rsThumb = Nothing

TheEnd:
    VendID = 0
    SaveVendor = True
    Me.MousePointer = vbDefault
    Exit Function
err:
    Me.MousePointer = vbDefault
    MsgBox err.Description
    
End Function
Private Sub SaveItems(p_iVendID As Integer)

    con.Execute "Delete From VendAssItemsOthers Where VendID=" & VendID
    For i = 1 To LV.ListItems.Count
        con.Execute "Insert Into VendAssItemsOthers (VendID,ItemCode,ItemRate) Values(" & VendID & ",'" & LV.ListItems(i).SubItems(1) & "'," & Val(LV.ListItems(i).SubItems(3)) & ")"
    Next
    
End Sub
Private Sub cmdSaveClose_Click()

    On Error GoTo err

    If SaveVendor Then
        'rsVend.Update
    
        Unload Me
    End If
    
    Exit Sub

err:
    MsgBox "ERROR : # " & err.Number & vbNewLine & err.Description, vbInformation
    
End Sub

Private Sub cmbType_Change()
    If cmbType.MatchFound Then Call CmbType_Click
End Sub
Private Sub SHForSub(bVal As Boolean)
    SSTab1.TabVisible(3) = bVal
    chkPlanter.value = Not bVal
    chkSnaffer.value = Not bVal
    chkStamper.value = Not bVal
    chkPlanter.Visible = bVal
    chkSnaffer.Visible = bVal
    chkPlanter.Visible = bVal
    cmbProcess1.Visible = Not bVal
    cmbProcess.Visible = bVal
End Sub
    
Private Sub CmbType_Click()
    Dim AccType As String, NewVal As String
    If cmbType = "" Or Not cmbType.Enabled Then Exit Sub
    
    If cmbType.ListIndex = 2 Then
        Call SHForSub(False)
    Else
        Call SHForSub(True)
    End If
    
    AccType = getAccountNo(cmbType)
    'NewVal = GetMax("right(AccNo,3)", "Accounts", "Where SubAccOf='" & AccType & "'")
    'txtfields(0) = AccType & "-" & Format(Val(NewVal) + 1, "000")
    
    NewVal = GetMax("right(AccNo,3)", "Accounts", "Where SubAccOf='" & ACCPAYNO & "'")
 
    'txtfields(0) = ACCPAYNO & "-" & Format(Val(NewVal) + 1 + iTotNo, "000")
        
    Call GetVendID
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    
    
    
    CopingVender = False
    iTotNo = 0
    Me.KeyPreview = True
    If Not AccDefined Then
        MsgBox "Account "" " & ACCRECTITLE & " "" is Not Defined, Please Define The Account First", vbInformation
    End If

    'Call AddToCombo(cmbType, "AccTitle", "Accounts", "Where SubAccof Like '" & AcNo & "%' And Parent=1")
    'cmbType.Enabled = cmbType.ListCount > 0

    DtOpenDate = ServerDate
    
    cmbType.AddItem "Manufacturer"
    cmbType.AddItem "Service Providers"
    cmbType.AddItem "Sub Venders"
    
    Call AddToCombo(cmbProcess, "Description", "Processes", " Order By SNo ", , "ProcessID")
    
    
    cmbProcess1.AddItem "Planter"
    cmbProcess1.AddItem "Stamper"
    cmbProcess1.AddItem "Snaffer"
    
    Call AddSubVenders
    SSTab1.Tab = 0
    
    
    cmbItem.ListHeight = 2000
    cmbItemName.ListHeight = 2000
    cmbItem.AddVals con, "ItemID", "VItems1", "ItemID", " Order By ItemID"
    cmbItemName.AddVals con, "CompleteItemName", "VItems1", "ItemID", " Order By ItemID"
    
End Sub
Private Sub AddSubVenders()
    On Error GoTo err
    
    cmbPlanter.ListHeight = 2200
    cmbSnaffer.ListHeight = 2200
    cmbStamper.ListHeight = 2200
    
    Call cmbPlanter.AddVals(con, "VenderName", "Venders", "VendID", " Where SubVendType='0'")
    Call cmbStamper.AddVals(con, "VenderName", "Venders", "VendID", " Where SubVendType='1'")
    Call cmbSnaffer.AddVals(con, "VenderName", "Venders", "VendID", " Where SubVendType='2'")
    
    
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Set rsVend = Nothing
End Sub


Private Sub GetAccInfo(AccNo As String)

    Dim rsAcc As New ADODB.Recordset

    With rsAcc
        .Open "Select AccNo,AccTitle,SubaccOf,OpenDate,OpenBal,Balance,Active from Accounts Where AccNo='" & AccNo & "'", con, adOpenForwardOnly
        AccNo = ![AccNo]
        DtOpenDate = ![OpenDate]
        txtfields(11) = Abs(![OpenBal])
        txtfields(1) = ![AccTitle]
        If Not ![Active] Then chkinactive = vbChecked Else chkinactive = vbUnchecked
        If ![SubAccOf] <> Acno Then cmbType = con.Execute("Select AccTitle From Accounts Where AccNo='" & ![SubAccOf] & "'").Fields(0)
   
        If Val(![OpenBal]) >= 0 Then
            cmbDrCr.ListIndex = 0
        Else
            cmbDrCr.ListIndex = 1
        End If

        'Remember Previous Opening Balance
        cmbDrCr.Tag = Val(![OpenBal] & "")

        Dim Balance As Double
        Balance = Val(![Balance] & "")
        lblBalance = IIf(Balance > 0, Abs(Balance) & " DR", IIf(Balance < 0, Abs(Balance) & " CR", sBalance))
        .Close
    End With

   
End Sub


Private Sub AddAccountInfo()

'    Call cmbDrCr_Click
'    Dim Active As Boolean
'    Active = Active = chkInactive <> vbChecked
'
'    Call NewAccPay(txtfields(1), DtOpenDate, Val(txtfields(11).Tag), Active, Left(txtfields(0), Len(txtfields(0)) - 4))
    
End Sub

Private Function AccDefined() As Boolean

    'Get Accounts Receiveable AccountNo
    Acno = ACCPAYNO
    If Acno = "" Then
        AccDefined = False
    Else
        AccDefined = True
    End If
    
End Function

Private Sub UpdateAccinfo()
 
'    Dim Diff As Double
'    Dim Active As Boolean
'    Call cmbDrCr_Click
'
'    Diff = Val(txtfields(11).Tag) - Val(cmbDrCr.Tag)
'    Active = chkInactive <> vbChecked
'
'    con.Execute "Update Accounts Set accTitle='" & txtfields(1) & "',OpenBal=" & Val(txtfields(11).Tag) & ",Balance=balance + " & Diff & ",Active=" & Active & " Where Accno='" & txtfields(0) & "'"
'
'    Call adjustLedger(txtfields(0), DtOpenDate, 0, Diff, True)
'
'    If Diff <> 0 Then Call UpdateAdvEntry(txtfields(0), Val(txtfields(11).Tag))
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmVender = Nothing
    OurOwnForm = True
End Sub



Private Sub Image1_DblClick()
    AddVendPic
End Sub

Private Sub Image2_Click()
    Call AddVendPic(True)
End Sub

Private Sub Label2_Click()
    Call AddVendPic(True)
End Sub

Private Sub lblPic_DblClick()
    AddVendPic
End Sub

  

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        LV.ListItems.Remove LV.SelectedItem.Index
    End If
End Sub

Private Sub NIC_Click(Index As Integer)
    If NIC(0).value Then
        txtfields(15).MaxLength = 13
        txtfields(15).Text = ""
    ElseIf NIC(1).value Then
        txtfields(15).MaxLength = 11
        txtfields(15).Text = ""
    End If
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Select Case Index
        Case 11
            KeyAscii = OnlyNumber(KeyAscii, True)
        Case 2, 3, 4, 6, 7, 12, 13
            If Not Chr(KeyAscii) = "-" Then KeyAscii = OnlyNumber(KeyAscii)
        Case 15, 9, 5
            Call OnlyNums(KeyAscii, False)
    End Select
    
End Sub

Private Sub ClearFields()
    For i = 0 To txtfields.Count - 1
        txtfields(i) = ""
    Next
End Sub

Private Function CreateInsertSQL(strTable As String) As String
    
    Dim C As Control
    Dim RetString As String
    
    
    'Create The First Part Of Insert Query...
    RetString = "Insert Into " & strTable & "("
    '---------------------------------------------------------------------------
    For Each C In Controls
        If C.Tag <> "" Then
            If C.Visible Then
                If (TypeOf C Is MSForms.TextBox) Or (TypeOf C Is MSForms.ComboBox) Or (TypeOf C Is ComboList.Usercontrol1) Or (TypeOf C Is TextBox) Or (TypeOf C Is CheckBox) Then
                    RetString = RetString & C.Tag & ","
                End If
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
    For Each C In Controls
        
        
        If C.Tag <> "" Then
            If C.Visible Then
                If (TypeOf C Is MSForms.TextBox) Or (TypeOf C Is TextBox) Or (TypeOf C Is MSForms.ComboBox) Then
                
                    If C.Name = "NBox" Then
                        RetString = RetString & Val(C.Text & "") & ","
                    ElseIf C.Name = "cmbProcess" Then
                        If chkRepVend.value = vbUnchecked Then
                            RetString = RetString & C.List(C.ListIndex, 1) & ","
                        Else
                            RetString = RetString & "0,"
                        End If
                    ElseIf C.Name = "cmbProcess1" Then
                        RetString = RetString & "'" & C.ListIndex & "',"
                    ElseIf C.Name = "cmbType" Then
                        RetString = RetString & "'" & C.ListIndex & "',"
                    Else
                        RetString = RetString & "'" & C.Text & "',"
                    End If
                ElseIf (TypeOf C Is CheckBox) Then
                    RetString = RetString & Abs(C.value) & ","
                ElseIf (TypeOf C Is ComboList.Usercontrol1) Then
                    'If C.Name = cmbSteel Or C.Name = cmbGuage Then
                        RetString = RetString & Val(C.ID & "") & ","
                    'RetString = RetString & "'" & C.ID & "',"
                End If
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
    
    Dim C As Control
    Dim RetString As String
    
    'Create The First Part Of Insert Query...
    RetString = "Update " & strTable & " Set "
    '---------------------------------------------------------------------------
    For Each C In Controls
    
        If C.Tag <> "" Then
            'If C.Visible Then
                If (TypeOf C Is MSForms.TextBox) Or (TypeOf C Is MSForms.ComboBox) Or (TypeOf C Is MSForms.ComboBox) Or (TypeOf C Is TextBox) Then '
                    If C.Name = "NBox" Then
                        RetString = RetString & C.Tag & "=" & C.Text & ","
                    ElseIf C.Name = "cmbSteel" Or C.Name = "cmbGuage" Or C.Name = "cmbItemGroup" Or C.Name = "cmbProcess" Then
                        If cmbType.ListIndex = 2 And C.Name = "cmbProcess" Then
                            RetString = RetString & C.Tag & "=-1,"
                        ElseIf C.Name = "cmbProcess" And Left(txtfields(12), 3) = "REP" Then
                            RetString = RetString & C.Tag & "=0,"
                        Else
                            RetString = RetString & C.Tag & "=" & C.List(C.ListIndex, 1) & ","
                        
                        End If
                    ElseIf C.Name = "cmbProcess1" Then
                        RetString = RetString & "SubVendType='" & C.ListIndex & "',"
                    ElseIf C.Name = "cmbType" Then
                        RetString = RetString & C.Tag & "='" & C.ListIndex & "',"
                    Else
                        RetString = RetString & C.Tag & "='" & C.Text & "',"
                    End If
                ElseIf (TypeOf C Is CheckBox) Then
                    RetString = RetString & C.Tag & "=" & Abs(C.value) & ","
                ElseIf (TypeOf C Is ComboList.Usercontrol1) Then
                        RetString = RetString & C.Tag & "=" & C.ID & ","
                End If
            'End If
        End If
    Next
    '---------------------------------------------------------------------------
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    
    CreateUpdateSQL = RetString
    
End Function


Private Sub AddVendPic(Optional ThumbExp As Boolean)
      On Error Resume Next
    Dim Picfile As String
    CD1.ShowSave
    If Not (err.Number = 32755) Then
        Picfile = CD1.FileName
        If ThumbExp Then
            Image2.Picture = LoadPicture(Picfile)
            ThumbFileName = Picfile
            Label2.Visible = False
        Else
            Image1.Picture = LoadPicture(Picfile)
            PicFileName = Picfile
            lblPic.Visible = False
        End If
        
    End If
End Sub

Private Sub ConvertToID(NewID As Boolean)

    'if User does not enter any value then
    If Len(Trim(txtfields(15).Text)) = 0 Then MsgBox "Please Enter valid NIC No.", vbInformation, "Information": Exit Sub
    Dim Str As String, iLen As Integer
    Str = Trim(txtfields(15))
    iLen = Len(Str)
    'if user select option for New NIC No.
    If NewID = True Then
        If iLen <> 13 Then
        
            If iLen = 15 Then
                If Mid$(Str, 6, 1) = "-" And Mid$(Str, 14, 1) = "-" Then
                    Exit Sub
                End If
            End If
            
            MsgBox "13 digits should be entered!", vbInformation
            txtfields(15).SelStart = 0
            txtfields(15).SelLength = Len(txtfields(15))
            txtfields(15).SetFocus
            Exit Sub
        End If
    Else
        'if user selects option for Old NIC
        If iLen <> 11 Then
            
            If iLen = 13 Then
                If Mid$(Str, 4, 1) = "-" And Mid$(Str, 7, 1) = "-" Then
                    Exit Sub
                End If
            End If
            
            MsgBox "11 digits should be entered!", vbInformation
            txtfields(15).SelStart = 0
            txtfields(15).SelLength = Len(txtfields(15))
            txtfields(15).SetFocus
            Exit Sub
        End If
    End If
    
    If NIC(0) Then
        txtfields(15).MaxLength = 15
        txtfields(15).Text = Format(txtfields(15), "#####-#######-#")
    Else
        txtfields(15).MaxLength = 13
        txtfields(15).Text = Format(txtfields(15), "###-##-######")
    End If
    
End Sub


Private Sub txtfields_Validate(Index As Integer, Cancel As Boolean)
    
    If Index = 15 Then
        If NIC(0).value Then
            Call ConvertToID(True)
        Else
            Call ConvertToID(False)
        End If
    End If
    
End Sub
Private Sub showdata()

    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select * From VVenders Where VendID=" & VendID & " Order By VendID", con, adOpenForwardOnly, adLockReadOnly
        VendID1 = ![VendID1] & ""
        cmbProcess = ![Description] & ""
        If Len(![NICNo]) = 13 Then
            NIC(1).value = True
        Else
            NIC(0).value = True
        End If
                
        For Each C In Controls
            If C.Tag <> "" Then
                'If TypeOf C Is CheckBox Then
                '    C.value = .Fields(C.Tag)
                'If TypeOf C Is ComboList.Usercontrol1 Then
                    'C.Text = .Fields(C.Tag & "Text") & ""
                '    C.ID = .Fields(C.Tag) & ""
                'If TypeOf C Is VB.Label Then
                '    C.Caption = .Fields(C.Tag) & ""
                If TypeOf C Is MSForms.ComboBox Then
                    If C.Name = "cmbType" Then
                        C.ListIndex = ![VendType]
                    ElseIf C.Name = "cmbProcess" Then
                        C.Text = ![Description] & ""
                    ElseIf C.Name = "cmbProcess1" Then
                        C.ListIndex = Val(![SubVendType] & "")
                    End If
                ElseIf (TypeOf C Is CheckBox) Then
                    C.value = Abs(.Fields(C.Tag))
                ElseIf TypeOf C Is TextBox Or (TypeOf C Is MSForms.TextBox) Then
                    C.Text = .Fields(C.Tag) & ""
                ElseIf TypeOf C Is ComboList.Usercontrol1 Then
                    C.ID = Val(.Fields(C.Tag) & "")
                
                End If
            End If
        Next
        
        If ![Planter] <> 0 Then
            chkPlanter.value = True
            cmbPlanter.Visible = True
        End If
        If ![Stamper] <> 0 Then
            chkStamper.value = True
            cmbStamper.Visible = True
        End If
        If ![Snaffer] <> 0 Then
            chkSnaffer.value = True
            cmbSnaffer.Visible = True
        End If
        
        If ![OtherCompany1] & "" <> "" Then
            chkComp1.value = True
            txtfields(16).Visible = True
        End If
        
        If ![OtherCompany2] & "" <> "" Then
            chkComp2.value = True
            txtfields(17).Visible = True
        End If
        
        .Close
        
        .Open "Select VendPic,VendThumb From Venders Where VendID=" & VendID, con, adOpenForwardOnly, adLockReadOnly
        
        If Not IsNull(![VendPic]) Then
            Set Image1.DataSource = rs
            Image1.DataField = "VendPic"
            lblPic.Visible = False
        Else
            lblPic.Visible = True
        End If
        
        If Not IsNull(![VendThumb]) Then
            Set Image2.DataSource = rs
            Image2.DataField = "VendThumb"
            Label2.Visible = False
        Else
            Label2.Visible = True
        End If
        
    End With
    
    Set rs = Nothing
    
    
End Sub
Private Sub GetRepairVendID()

    On Error GoTo err
    
    If ShowingData Then Exit Sub
    
    Dim strVendID As String
    
    If VendID <> 0 Then
        If UCase(Left(VendID1, 3)) = UCase(Left(cmbProcess, 3)) Then
            txtfields(12) = VendID1
            Exit Sub
        End If
    End If
    
    
    strVendID = Val(con.Execute("Select Max(Cast(SubString(VendID1,5,4) as int)) From Venders Where Left(VendID1,3)='REP'").Fields(0).value & "")
    
    If strVendID = 0 Then
        strVendID = 5000
    End If
    
    strVendID = strVendID + 1
    
    txtfields(12) = "REP-" & Format(strVendID, "0000")
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
