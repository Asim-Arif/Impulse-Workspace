VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmNewMaker_NewComing 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Coming Maker Info."
   ClientHeight    =   7320
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
   ScaleHeight     =   7320
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   150
      TabIndex        =   8
      Top             =   0
      Width           =   8700
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Coming Maker INfo."
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
         Left            =   2370
         TabIndex        =   9
         Top             =   120
         Width           =   3960
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Coming Maker INfo."
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
         Left            =   2370
         TabIndex        =   10
         Top             =   135
         Width           =   3960
      End
   End
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
      Height          =   6600
      Left            =   60
      TabIndex        =   4
      Top             =   600
      Width           =   8925
      Begin VB.TextBox txtfields 
         DataField       =   "AccNo"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   5775
         TabIndex        =   83
         Tag             =   "AccNo"
         Top             =   420
         Width           =   2775
      End
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
         TabIndex        =   70
         Top             =   5835
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.CheckBox chkRepVend 
         Caption         =   "Repair Maker"
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
         TabIndex        =   69
         Top             =   5835
         Visible         =   0   'False
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
         TabIndex        =   54
         Tag             =   "MakerID1"
         Top             =   435
         Width           =   1830
      End
      Begin VB.CheckBox chkInactive 
         Caption         =   "Maker is inactive"
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
         TabIndex        =   47
         Top             =   6150
         Visible         =   0   'False
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
         Tag             =   "MakerName"
         Top             =   435
         Width           =   3690
      End
      Begin VB.TextBox txtfields 
         DataField       =   "VendID"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   3495
         TabIndex        =   11
         Top             =   5835
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
         Height          =   4860
         Left            =   105
         TabIndex        =   12
         Top             =   840
         Width           =   8640
         _ExtentX        =   15240
         _ExtentY        =   8573
         _Version        =   393216
         Tabs            =   5
         Tab             =   4
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "General Info."
         TabPicture(0)   =   "frmNewMaker_NewComing.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "lblPic"
         Tab(0).Control(1)=   "Label1(32)"
         Tab(0).Control(2)=   "Image1"
         Tab(0).Control(3)=   "Label1(9)"
         Tab(0).Control(4)=   "Label2"
         Tab(0).Control(5)=   "Image2"
         Tab(0).Control(6)=   "lblLabels(11)"
         Tab(0).Control(7)=   "cmbAuthReq"
         Tab(0).Control(8)=   "Label1(3)"
         Tab(0).Control(9)=   "Label1(2)"
         Tab(0).Control(10)=   "Label1(4)"
         Tab(0).Control(11)=   "Label1(5)"
         Tab(0).Control(12)=   "cmbShowRate"
         Tab(0).Control(13)=   "Label1(6)"
         Tab(0).Control(14)=   "cmbPaymentTerms"
         Tab(0).Control(15)=   "Frame1"
         Tab(0).Control(16)=   "txtUrduFields(0)"
         Tab(0).Control(17)=   "txtfields(19)"
         Tab(0).Control(18)=   "txtfields(20)"
         Tab(0).ControlCount=   19
         TabCaption(1)   =   "Detail"
         TabPicture(1)   =   "frmNewMaker_NewComing.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lblLabels(15)"
         Tab(1).Control(1)=   "cmbType"
         Tab(1).Control(2)=   "lblLabels(16)"
         Tab(1).Control(3)=   "cmbProcess"
         Tab(1).Control(4)=   "cmbProcess1"
         Tab(1).Control(5)=   "lblLabels(18)"
         Tab(1).Control(6)=   "chkComp2"
         Tab(1).Control(7)=   "chkComp1"
         Tab(1).Control(8)=   "lblLabels(19)"
         Tab(1).Control(9)=   "lblItems"
         Tab(1).Control(10)=   "cmdAssign"
         Tab(1).Control(11)=   "LV"
         Tab(1).Control(12)=   "txtfields(14)"
         Tab(1).Control(13)=   "txtfields(16)"
         Tab(1).Control(14)=   "txtfields(17)"
         Tab(1).Control(15)=   "txtfields(18)"
         Tab(1).Control(16)=   "cmbItem"
         Tab(1).Control(17)=   "cmbItemName"
         Tab(1).Control(18)=   "txtfields(0)"
         Tab(1).ControlCount=   19
         TabCaption(2)   =   "Referal Info"
         TabPicture(2)   =   "frmNewMaker_NewComing.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "FRefInfo"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Sub Makers"
         TabPicture(3)   =   "frmNewMaker_NewComing.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "chkPlanter"
         Tab(3).Control(1)=   "chkSnaffer"
         Tab(3).Control(2)=   "chkStamper"
         Tab(3).Control(3)=   "cmbSnaffer"
         Tab(3).Control(4)=   "cmbStamper"
         Tab(3).Control(5)=   "cmbPlanter"
         Tab(3).ControlCount=   6
         TabCaption(4)   =   "Processes"
         TabPicture(4)   =   "frmNewMaker_NewComing.frx":0070
         Tab(4).ControlEnabled=   -1  'True
         Tab(4).Control(0)=   "LV2"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).ControlCount=   1
         Begin VB.TextBox txtfields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "Fax2"
            DataSource      =   "datPrimaryRS"
            Height          =   330
            Index           =   20
            Left            =   -69930
            MaxLength       =   100
            TabIndex        =   91
            Tag             =   "MaximumRcvingsAgainstPO"
            Top             =   2040
            Width           =   1815
         End
         Begin VB.TextBox txtfields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "Fax2"
            DataSource      =   "datPrimaryRS"
            Height          =   330
            Index           =   19
            Left            =   -69930
            MaxLength       =   100
            TabIndex        =   89
            Tag             =   "ExcessQtyPercentage"
            Top             =   1290
            Width           =   1815
         End
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
            Left            =   -69870
            TabIndex        =   79
            Top             =   4365
            Width           =   3345
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
            Height          =   285
            Index           =   0
            Left            =   -68475
            MaxLength       =   10
            TabIndex        =   78
            Top             =   1830
            Visible         =   0   'False
            Width           =   1125
         End
         Begin ComboList.Usercontrol1 cmbItemName 
            Height          =   285
            Left            =   -73095
            TabIndex        =   73
            Top             =   1830
            Visible         =   0   'False
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
            Left            =   -74850
            TabIndex        =   74
            Top             =   1830
            Visible         =   0   'False
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
            Left            =   -68175
            MaxLength       =   10
            TabIndex        =   71
            Top             =   675
            Width           =   1590
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "Phone1"
            Height          =   315
            Index           =   17
            Left            =   -69900
            MaxLength       =   20
            TabIndex        =   68
            Top             =   1245
            Visible         =   0   'False
            Width           =   3315
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "Phone1"
            Height          =   315
            Index           =   16
            Left            =   -73230
            MaxLength       =   20
            TabIndex        =   67
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
            Left            =   -74850
            MaxLength       =   10
            TabIndex        =   63
            Top             =   1245
            Width           =   1590
         End
         Begin VB.Frame FRefInfo 
            Height          =   2370
            Left            =   -73755
            TabIndex        =   36
            Top             =   825
            Width           =   5850
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Address"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   28
               Left            =   2370
               TabIndex        =   46
               Top             =   945
               Width           =   585
            End
            Begin MSForms.TextBox TBox 
               Height          =   990
               Index           =   21
               Left            =   2340
               TabIndex        =   45
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
               TabIndex        =   44
               Top             =   1635
               Width           =   585
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   20
               Left            =   465
               TabIndex        =   43
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
               TabIndex        =   42
               Top             =   960
               Width           =   585
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   19
               Left            =   465
               TabIndex        =   41
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
               TabIndex        =   40
               Top             =   300
               Width           =   1035
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   18
               Left            =   2940
               TabIndex        =   39
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
               TabIndex        =   38
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
               TabIndex        =   37
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
            Height          =   4335
            Left            =   -74910
            TabIndex        =   13
            Top             =   420
            Width           =   4920
            Begin VB.TextBox txtfields 
               DataField       =   "BankAccNo"
               DataSource      =   "datPrimaryRS"
               Height          =   285
               Index           =   11
               Left            =   105
               MaxLength       =   100
               TabIndex        =   81
               Tag             =   "BankAccNo"
               Top             =   3420
               Width           =   4740
            End
            Begin VB.TextBox txtfields 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               DataField       =   "Fax1"
               DataSource      =   "datPrimaryRS"
               Height          =   315
               Index           =   9
               Left            =   2205
               MaxLength       =   100
               TabIndex        =   51
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
               TabIndex        =   50
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
               TabIndex        =   30
               Tag             =   "Mobile"
               Top             =   1065
               Width           =   3150
            End
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "Fax1"
               DataSource      =   "datPrimaryRS"
               Height          =   315
               Index           =   15
               Left            =   90
               MaxLength       =   100
               TabIndex        =   28
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
               TabIndex        =   18
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
               TabIndex        =   17
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
               TabIndex        =   16
               Tag             =   "Fax1"
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
               TabIndex        =   15
               Tag             =   "NTNNo"
               Top             =   1065
               Width           =   1560
            End
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "Address"
               DataSource      =   "datPrimaryRS"
               Height          =   885
               Index           =   8
               Left            =   90
               MaxLength       =   255
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   14
               Tag             =   "Address"
               Top             =   2280
               Width           =   4740
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Maker Catagory :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   7
               Left            =   135
               TabIndex        =   97
               Top             =   3735
               Width           =   1260
            End
            Begin MSForms.ComboBox cmbMasterCat 
               Height          =   300
               Left            =   105
               TabIndex        =   96
               Top             =   3945
               Width           =   4740
               VariousPropertyBits=   746604571
               DisplayStyle    =   3
               Size            =   "8361;529"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   178
               FontPitchAndFamily=   2
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Bank Account No. :"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   7
               Left            =   135
               TabIndex        =   82
               Top             =   3195
               Width           =   1380
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Time Limit (days)"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   10
               Left            =   3540
               TabIndex        =   53
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
               TabIndex        =   52
               Top             =   1470
               Width           =   720
            End
            Begin MSForms.OptionButton NIC 
               Height          =   285
               Index           =   1
               Left            =   570
               TabIndex        =   49
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
               TabIndex        =   48
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
               TabIndex        =   31
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
               TabIndex        =   29
               Top             =   1440
               Width           =   435
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Phone1"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   3
               Left            =   90
               TabIndex        =   23
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
               TabIndex        =   22
               Top             =   240
               Width           =   540
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fax"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   6
               Left            =   3285
               TabIndex        =   21
               Top             =   255
               Width           =   270
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "NTN No"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   8
               Left            =   3285
               TabIndex        =   20
               Top             =   840
               Width           =   540
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   9
               Left            =   90
               TabIndex        =   19
               Top             =   2070
               Width           =   585
            End
         End
         Begin ComboList.Usercontrol1 cmbPlanter 
            Height          =   285
            Left            =   -73740
            TabIndex        =   60
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
            TabIndex        =   61
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
            TabIndex        =   59
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
            Height          =   2625
            Left            =   -74850
            TabIndex        =   77
            Top             =   2130
            Visible         =   0   'False
            Width           =   8265
            _ExtentX        =   14579
            _ExtentY        =   4630
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
         Begin MSComctlLib.ListView LV2 
            Height          =   4095
            Left            =   450
            TabIndex        =   85
            Top             =   540
            Width           =   7740
            _ExtentX        =   13653
            _ExtentY        =   7223
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Process"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Code"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Sort No."
               Object.Width           =   1764
            EndProperty
         End
         Begin MSForms.ComboBox cmbPaymentTerms 
            Height          =   300
            Left            =   -69930
            TabIndex        =   95
            Tag             =   "PaymentTerms"
            Top             =   3390
            Width           =   1815
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "3201;529"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Payment Terms"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   -69930
            TabIndex        =   94
            Top             =   3195
            Width           =   1110
         End
         Begin MSForms.ComboBox cmbShowRate 
            Height          =   300
            Left            =   -69930
            TabIndex        =   93
            Tag             =   "ShowRateOnPO"
            Top             =   2715
            Width           =   1455
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "2566;529"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Show Rate"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   -69930
            TabIndex        =   92
            Top             =   2520
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Max. Rcvings Against PO"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   -69930
            TabIndex        =   90
            Top             =   1830
            Width           =   1800
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Excess Qty Percentage"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   -69915
            TabIndex        =   88
            Top             =   1080
            Width           =   1680
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Authorization Req."
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   -69915
            TabIndex        =   87
            Top             =   420
            Width           =   1350
         End
         Begin MSForms.ComboBox cmbAuthReq 
            Height          =   300
            Left            =   -69930
            TabIndex        =   86
            Tag             =   "AuthRequired"
            Top             =   615
            Width           =   1455
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "2566;529"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Urdu Name :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   11
            Left            =   -69840
            TabIndex        =   80
            Top             =   4140
            Width           =   900
         End
         Begin MSForms.CommandButton cmdAssign 
            Height          =   525
            Left            =   -67335
            TabIndex        =   75
            Top             =   1590
            Visible         =   0   'False
            Width           =   750
            ForeColor       =   6244673
            Caption         =   "Add"
            PicturePosition =   327683
            Size            =   "1323;926"
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
            Left            =   -74850
            TabIndex        =   76
            Top             =   1605
            Visible         =   0   'False
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
            Left            =   -68175
            TabIndex        =   72
            Top             =   450
            Width           =   1245
         End
         Begin MSForms.CheckBox chkComp1 
            Height          =   345
            Left            =   -73230
            TabIndex        =   66
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
            Left            =   -69915
            TabIndex        =   65
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
            Left            =   -74835
            TabIndex        =   64
            Top             =   1020
            Width           =   1560
         End
         Begin MSForms.ComboBox cmbProcess1 
            Height          =   300
            Left            =   -71955
            TabIndex        =   62
            Top             =   765
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
            TabIndex        =   58
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
            TabIndex        =   57
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
            TabIndex        =   56
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
            Left            =   -68085
            Stretch         =   -1  'True
            Top             =   2820
            Width           =   1545
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Double Click To Add Thumb Exp. /Signature"
            ForeColor       =   &H00FF0000&
            Height          =   615
            Left            =   -67920
            TabIndex        =   35
            Top             =   3195
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Thumb Exp/Signature"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   -68085
            TabIndex        =   34
            Top             =   2595
            Width           =   1545
         End
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   1740
            Left            =   -68085
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
            Left            =   -68070
            TabIndex        =   33
            Top             =   405
            Width           =   495
         End
         Begin VB.Label lblPic 
            Alignment       =   2  'Center
            Caption         =   "Double Click To Add Picture."
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   -67860
            TabIndex        =   32
            Top             =   1290
            Width           =   1095
         End
         Begin MSForms.ComboBox cmbProcess 
            Height          =   300
            Left            =   -71955
            TabIndex        =   27
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
            Left            =   -71910
            TabIndex        =   26
            Top             =   465
            Width           =   555
         End
         Begin MSForms.ComboBox cmbType 
            Height          =   300
            Left            =   -74850
            TabIndex        =   25
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
            Caption         =   "Maker Type"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   15
            Left            =   -74835
            TabIndex        =   24
            Top             =   465
            Width           =   840
         End
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   8115
         Top             =   1395
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
         Caption         =   "Acc # :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   2
         Left            =   5775
         TabIndex        =   84
         Top             =   195
         Width           =   645
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Maker Code"
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
         TabIndex        =   55
         Top             =   225
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Maker Name"
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
         Width           =   1065
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vender ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
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
Attribute VB_Name = "frmNewMaker_NewComing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim rsVend As New ADODB.Recordset
Public add As Boolean, edit As Boolean, View
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
    Call ShowData
    Call ShowItems
    ShowingData = False
    Me.Show 1
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub ShowItems()

    On Error GoTo err
    
'    Dim rs As New ADODB.Recordset
'    Dim ITM As ListItem
'    With rs
'        .Open "Select * From VMakerAssItems Where VendID=" & VendID, con, adOpenForwardOnly, adLockReadOnly
'        LV.ListItems.Clear
'        Do Until .EOF
'            Set ITM = LV.ListItems.add(, !ITemID & "'", LV.ListItems.count + 1)
'            ITM.SubItems(1) = !ITemID & ""
'            ITM.SubItems(2) = !ItemName & ""
'            ITM.SubItems(3) = !Rate & ""
'            .MoveNext
'        Loop
'        .Close
'    End With
'
'    Set rs = Nothing
    
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
    Call ShowData
    VendID = 0
    
    ShowingData = False
    Call GetVendID
    Me.Show 1
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub chkComp1_Click()
    txtfields(16).Visible = chkComp1.Value
End Sub

Private Sub chkComp2_Click()
    txtfields(17).Visible = chkComp2.Value
End Sub

Private Sub chkInactive_Click()

    If chkInactive = vbChecked And Val(lblBalance) <> 0 Then
        MsgBox "This Vender Cannot Be Made Inactive," & vbNewLine & _
         "To Make A Vender Inactive Balance Must Be Zero", vbInformation
        chkInactive = vbUnchecked
    End If
    
End Sub


Private Sub chkOwnRep_Click()
    If chkOwnRep.Value = vbChecked Then
        txtfields(18).Text = 0
        txtfields(18).Locked = True
    Else
        txtfields(18).Locked = False
    End If
End Sub

Private Sub chkPlanter_Click()
    
    cmbPlanter.Visible = chkPlanter.Value
    If Not chkPlanter.Value Then cmbPlanter.ID = 0
    
End Sub

Private Sub chkRepVend_Click()
    If chkRepVend.Value = vbChecked Then
        cmbType.ListIndex = 1
        cmbType.Enabled = False
        cmbProcess.Enabled = False
        chkOwnRep.Value = vbUnchecked
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
    cmbSnaffer.Visible = chkSnaffer.Value
    If Not chkSnaffer.Value Then cmbSnaffer.ID = 0
End Sub

Private Sub chkStamper_Click()
    cmbStamper.Visible = chkStamper.Value
    If Not chkStamper.Value Then cmbStamper.ID = 0
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
    Set ITM = LV.ListItems.add(, cmbType & cmbProcess.List(cmbProcess.ListIndex, 1), LV.ListItems.count + 1)
    ITM.ListSubItems.add , , txtfields(14) & ""
    ITM.ListSubItems.add , , cmbType.Text & ""
    ITM.ListSubItems.add , cmbProcess.List(cmbProcess.ListIndex, 1) & "'", cmbProcess.Text & ""
    ITM.ListSubItems.add , , txtfields(0) & ""
    ITM.ListSubItems.add(, , OpBal).Tag = Format(DtOpenDate, "dd-MMM-yyyy")
    
    iTotNo = iTotNo + 1
    txtfields(11) = ""
    'cmbDrCr = ""
    cmbType.Text = ""
    cmbProcess.Text = ""
    
    Call cmbType_Click
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
    
'    If cmbType.ListIndex = "0" Or cmbType.ListIndex = "1" Then
'        strVendID = Val(con.Execute("Select Max(Cast(SubString(VendID1,5,4) as int)) From Venders Where PhaseID=" & cmbProcess.List(cmbProcess.ListIndex, 1) & " and VendType='" & cmbType.ListIndex & "'").Fields(0).value & "")
'    ElseIf cmbType.ListIndex = "2" Then
'        strVendID = Val(con.Execute("Select Max(Cast(SubString(VendID1,5,4) as int)) From Venders Where VendType='" & cmbType.ListIndex & "'").Fields(0).value & "")
'    End If
    
'    If strVendID = 0 Then
'        If cmbType.ListIndex = "0" Or cmbType.ListIndex = "2" Then
'            strVendID = 1000
'        ElseIf cmbType.ListIndex = "1" Then
'            strVendID = 5000
'        End If
'    End If
'
'    strVendID = strVendID + 1
'    If cmbType.ListIndex = "0" Or cmbType.ListIndex = "1" Then
'        txtfields(12) = UCase(Left(cmbProcess.Text, 3)) & Format(strVendID, "-0000")
'    ElseIf cmbType.ListIndex = "2" Then
'        txtfields(12) = "SUB" & Format(strVendID, "-0000") 'UCase(Left(cmbProcess.Text, 3)) & Format(strVendID, "-1000")
'    End If
    
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
    
    Set ITM = LV.ListItems.add(, cmbItem.ID & "'", LV.ListItems.count + 1)
    
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
    
    If SaveVendor Then
        Unload Me
        Me.Show 1
    End If
    
    
    Exit Sub
    If Trim(txtfields(1)) = "" Then
        MsgBox "Invalid Vender Name", vbInformation
        Exit Sub
    ElseIf cmbType.Enabled And Not cmbType.MatchFound Then
        MsgBox "Please Select Vender Type", vbInformation
        cmbType.SetFocus
        Exit Sub
    End If

    If add Then
        Call AddAccountInfo
    ElseIf edit Then
        Call UpdateAccinfo
    End If


    rsVend.Update
    
    Me.add = True
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
    Dim Rec As New ADODB.Recordset
    
'    If cmbProcess.MatchFound = False Then
'        MsgBox "Please Select Process.", vbInformation
'        Exit Function
    
    If Trim(txtfields(1)) = "" Then
        MsgBox "Invalid Maker Name", vbInformation
        Exit Function
'    ElseIf cmbMasterCat.MatchFound = False And VendID <> iFactoryMakerID And VendID = 0 Then
'        MsgBox "Please Select Maker Catagory.", vbInformation
'        Exit Function
'    ElseIf cmbAuthReq.MatchFound = False Then
'        MsgBox "Please Select Authorization.", vbInformation
'        Exit Function
'    ElseIf txtfields(15) = "" Then
'        MsgBox "Please Enter NIC No.", vbInformation
'        Exit Function
    End If
    
    
    Dim lCount As Long
    If VendID = 0 Then
        lCount = GetSingleLongValue("COUNT(NICNo)", "Makers_NewComing", " WHERE NicNo='" & txtfields(15) & "'")
    Else
        lCount = GetSingleLongValue("COUNT(NICNo)", "Makers_NewComing", " WHERE MakerID<>" & VendID & " AND NicNo='" & txtfields(15) & "'")
    End If
    If lCount > 0 Then
        MsgBox "NIC No. Already Exists.", vbInformation
        Exit Function
    End If
'    If Not cmbType.MatchFound Then
'        MsgBox "No Type Is Selected!", vbInformation
'        SSTab1.Tab = 1
'        cmbType.SetFocus
'        Exit Function
'    End If
    
    If chkRepVend.Value = vbUnchecked Then
'        If cmbProcess.Visible Then
'            If Not cmbProcess.MatchFound Then
'                MsgBox "No Process Is Selected!", vbInformation
'                SSTab1.Tab = 1
'                cmbProcess.SetFocus
'                Exit Function
'            End If
'        End If
    
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
    Dim SQL As String
    
    'AccType = getAccountNo(cmbType)
    
    
    If VendID = 0 Then
          
        SQL = CreateInsertSQL("Makers_NewComing")
        
    Else
        SQL = CreateUpdateSQL("Makers_NewComing")
        SQL = SQL & " Where MakerID=" & VendID
    
    End If
    
    con.Execute SQL
    If VendID = 0 Then
        Rec.Open "SELECT Max(MakerID) From Makers_NewComing", con, adOpenForwardOnly, adLockReadOnly
        VendID = Rec.Fields(0) & ""
        Rec.Close
    End If
    Call SaveItems(VendID)
    If bUrduSupported Then
        con.Execute "UPDATE Makers_NewComing SET MakerNameUrdu=N'" & txtUrduFields(0).Text & "' WHERE MakerID=" & VendID
    End If
    
    'If Add Then
    '    Call AddAccountInfo
    'ElseIf Edit Then
    '    Call UpdateAccinfo
    'End If
'    If CopingVender Then
'        Dim NewVendID As Long
'        NewVendID = Val(con.Execute("Select Max(MakerID) From Makers_NewComing").Fields(0).Value & "")
'
'        Dim cmd As New ADODB.Command
'        With cmd
'            Set .ActiveConnection = con
'            .CommandType = adCmdStoredProc
'            .CommandText = "SP_CopyAssItems"
'            .Parameters("@VendID").Value = NewVendID
'            .Parameters("@FromVendID").Value = FromVendID
'            .Execute
'        End With
'    End If
    
    'Save Processes....
    'Call SaveProcesses(CLng(VendID))
    con.CommitTrans
    
SavePic:
    If (Image1.Picture = 0 Or Trim(PicFileName) = "") And (Image2.Picture = 0 Or Trim(ThumbFileName) = "") Then GoTo TheEnd

    
    Dim rsPic As New ADODB.Stream
    Dim rsThumb As New ADODB.Stream
    
    
    Rec.Open "Select VendID,VendPic,VendThumb from Makers_NewComing Where VendID=" & VendID, con, adOpenStatic, adLockOptimistic
    
    If Image1.Picture <> 0 And Trim(PicFileName) <> "" Then
      'Save Picture
        rsPic.Open
        rsPic.Type = adTypeBinary
        rsPic.LoadFromFile PicFileName
        Rec.Fields(1) = rsPic.Read
        rsPic.Close
    End If
    
    'Save Thumb Expression
    If Image2.Picture <> 0 And Trim(ThumbFileName) <> "" Then
        rsThumb.Open
        rsThumb.Type = adTypeBinary
        rsThumb.LoadFromFile ThumbFileName
        Rec.Fields("VendThumb") = rsThumb.Read
        rsThumb.Close
    End If
    
    Rec.Update
    Rec.Close
    Set Rec = Nothing
    Set rsPic = Nothing
    Set rsThumb = Nothing

TheEnd:
    VendID = 0
    SaveVendor = True
    Me.MousePointer = vbDefault
    Exit Function
err:
    Me.MousePointer = vbDefault
    MsgBox err.Description & vbNewLine & SQL
    
End Function

Private Sub SaveProcesses(lVendID As Long)
    
    'Save Item's Production Processes...
    '------------------------------------------------------------------------
    con.Execute "Delete From MakerProcesses Where VendID=" & lVendID
    For i = 1 To LV2.ListItems.count
        If LV2.ListItems(i).Checked Then
             con.Execute "Insert Into MakerProcesses(VendID,ProcessID) " & _
             " Values(" & lVendID & "," & Val(LV2.ListItems(i).key) & ")"
             
        End If
    Next i
    '------------------------------------------------------------------------
    
End Sub

Private Sub SaveItems(p_iVendID As Integer)

'    con.Execute "Delete From VendAssItemsOthers Where VendID=" & VendID
'    For i = 1 To LV.ListItems.count
'        con.Execute "Insert Into VendAssItemsOthers (VendID,ItemCode,ItemRate) Values(" & VendID & ",'" & LV.ListItems(i).SubItems(1) & "'," & Val(LV.ListItems(i).SubItems(3)) & ")"
'    Next
    
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

    If cmbType.MatchFound Then Call cmbType_Click
End Sub

Private Sub SHForSub(bVal As Boolean)
    'SSTab1.TabVisible(3) = bVal
    chkPlanter.Value = Not bVal
    chkSnaffer.Value = Not bVal
    chkStamper.Value = Not bVal
    chkPlanter.Visible = bVal
    chkSnaffer.Visible = bVal
    chkPlanter.Visible = bVal
    cmbProcess1.Visible = Not bVal
    cmbProcess.Visible = bVal
End Sub
    
Private Sub cmbType_Click()
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
    
    Acno = ACCPAYNO
    
    CopingVender = False
    iTotNo = 0
    Me.KeyPreview = True
'    If Not AccDefined Then
'        MsgBox "Account "" " & ACCRECTITLE & " "" is Not Defined, Please Define The Account First", vbInformation
'    End If
    
    'Call AddToCombo(cmbType, "AccTitle", "Accounts", "Where SubAccof Like '" & AcNo & "%' And Parent=1")
    'cmbType.Enabled = cmbType.ListCount > 0

    DtOpenDate = ServerDate
    
'    cmbType.AddItem "Manufacturer"
'    cmbType.AddItem "Service Providers"
'    cmbType.AddItem "Sub Venders"
    Call AddToCombo(cmbType, "DISTINCT MakerType", "Makers_NewComing", " WHERE MakerType IS NOT NULL")
    
    Call AddToCombo(cmbProcess, "Description", "Processes", " Order By SNo ", , "ProcessID")
    
    Call AddToCombo(cmbMasterCat, "AccTitle", "Accounts", " WHERE Parent=1", , "AccNo")
    
    cmbProcess1.AddItem "Planter"
    cmbProcess1.AddItem "Stamper"
    cmbProcess1.AddItem "Snaffer"
    With cmbAuthReq
        .AddItem "No"
        .AddItem "Yes"
    End With
    
    With cmbShowRate
        .AddItem "No"
        .AddItem "Yes"
    End With
    Call AddToCombo(cmbPaymentTerms, "DISTINCT PaymentTerms", "Makers_NewComing")
    Call AddSubVenders
    SSTab1.TabVisible(3) = False
    SSTab1.TabVisible(1) = False
    SSTab1.Tab = 0
    
    
'    cmbItem.ListHeight = 2000
'    cmbItemName.ListHeight = 2000
'    cmbItem.AddVals con, "ItemID", "VItems1", "ItemID", " Order By ItemID"
'    cmbItemName.AddVals con, "CompleteItemName", "VItems1", "ItemID", " Order By ItemID"
    
    Call AddProcesses
    
    If strCompany = "QEL" Then
        txtfields(12) = GetSingleLongValue("MAX(CAST(VendID1 AS INT))", "Makers_NewComing", "WHERE ISNUMERIC(VendID1)=1") + 1
    ElseIf strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
        Dim strVendID As String
        
        'strVendID = (GetSingleLongValue("COUNT(VendID)", "Makers_NewComing") + 1)
        strVendID = GetSingleLongValue("MAX(CAST(RIGHT(MakerID1,4) AS INT))", "Makers_NewComing", "WHERE ISNUMERIC(RIGHT(MakerID1,4))=1") + 1
        strVendID = Format(strVendID, "0000")
        strVendID = "NMKR-" & strVendID
        txtfields(12) = strVendID
    End If
    
End Sub
Private Sub AddSubVenders()
    On Error GoTo err
    
    cmbPlanter.ListHeight = 2200
    cmbSnaffer.ListHeight = 2200
    cmbStamper.ListHeight = 2200
    
'    Call cmbPlanter.AddVals(con, "VenderName", "Venders", "VendID", " Where SubVendType='0'")
'    Call cmbStamper.AddVals(con, "VenderName", "Venders", "VendID", " Where SubVendType='1'")
'    Call cmbSnaffer.AddVals(con, "VenderName", "Venders", "VendID", " Where SubVendType='2'")
    
    
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)

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
        If Not ![Active] Then chkInactive = vbChecked Else chkInactive = vbUnchecked
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

Private Sub Form_Unload(cancel As Integer)
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

Private Sub LV2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked = False Then
        'Check for Assigned Items....
    End If
End Sub

Private Sub NIC_Click(Index As Integer)
    If NIC(0).Value Then
        txtfields(15).MaxLength = 13
        txtfields(15).Text = ""
    ElseIf NIC(1).Value Then
        txtfields(15).MaxLength = 11
        txtfields(15).Text = ""
    End If
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Select Case Index
        Case 11
            KeyAscii = OnlyNumber(KeyAscii, True)
        Case 2, 3, 4, 6, 7, 13
            If Not Chr(KeyAscii) = "-" Then KeyAscii = OnlyNumber(KeyAscii)
        Case 15, 9, 5
            Call OnlyNums(KeyAscii, False)
    End Select
    
End Sub

Private Sub ClearFields()
    For i = 0 To txtfields.count - 1
        txtfields(i) = ""
    Next
End Sub

Private Function CreateInsertSQL(strTable As String) As String
    
    Dim c As Control
    Dim RetString As String
    
    
    'Create The First Part Of Insert Query...
    RetString = "Insert Into " & strTable & "("
    '---------------------------------------------------------------------------
    For Each c In Controls
        If c.Tag <> "" Then
            If c.Visible Then
                If (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is MSForms.ComboBox) Or (TypeOf c Is ComboList.Usercontrol1) Or (TypeOf c Is TextBox) Or (TypeOf c Is CheckBox) Then
                    RetString = RetString & c.Tag & ","
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
    For Each c In Controls
        
        
        If c.Tag <> "" Then
            If c.Visible Then
                If (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is TextBox) Or (TypeOf c Is MSForms.ComboBox) Then
                
                    If c.Name = "NBox" Then
                        RetString = RetString & Val(c.Text & "") & ","
                    ElseIf c.Name = "cmbProcess" Then
                        If chkRepVend.Value = vbUnchecked Then
                            RetString = RetString & c.List(c.ListIndex, 1) & ","
                        Else
                            RetString = RetString & "0,"
                        End If
                    ElseIf c.Name = "cmbProcess1" Then
                        RetString = RetString & "'" & c.ListIndex & "',"
                    ElseIf c.Name = "cmbType" Then
                        RetString = RetString & "'" & c.ListIndex & "',"
                    ElseIf c.Name = "cmbAuthReq" Or c.Name = "cmbShowRate" Then
                        RetString = RetString & c.ListIndex & ","
                    Else
                        RetString = RetString & "'" & c.Text & "',"
                    End If
                ElseIf (TypeOf c Is CheckBox) Then
                    RetString = RetString & Abs(c.Value) & ","
                ElseIf (TypeOf c Is ComboList.Usercontrol1) Then
                    'If C.Name = cmbSteel Or C.Name = cmbGuage Then
                        RetString = RetString & Val(c.ID & "") & ","
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
    
    Dim c As Control
    Dim RetString As String
    
    'Create The First Part Of Insert Query...
    RetString = "Update " & strTable & " Set "
    '---------------------------------------------------------------------------
    For Each c In Controls
    
        If c.Tag <> "" Then
            'If C.Visible Then
                If (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is MSForms.ComboBox) Or (TypeOf c Is MSForms.ComboBox) Or (TypeOf c Is TextBox) Then '
                    If c.Name = "NBox" Then
                        RetString = RetString & c.Tag & "=" & c.Text & ","
                    ElseIf c.Name = "cmbSteel" Or c.Name = "cmbGuage" Or c.Name = "cmbItemGroup" Or c.Name = "cmbProcess" Then
                        If cmbType.ListIndex = 2 And c.Name = "cmbProcess" Then
                            RetString = RetString & c.Tag & "=-1,"
                        ElseIf c.Name = "cmbProcess" And Left(txtfields(12), 3) = "REP" Then
                            RetString = RetString & c.Tag & "=0,"
                        Else
                            RetString = RetString & c.Tag & "=" & c.List(c.ListIndex, 1) & ","
                        End If
                    ElseIf c.Name = "cmbProcess1" Then
                        RetString = RetString & "SubVendType='" & c.ListIndex & "',"
                    ElseIf c.Name = "cmbType" Then
                        RetString = RetString & c.Tag & "='" & c.Text & "',"
                    ElseIf c.Name = "cmbAuthReq" Or c.Name = "cmbShowRate" Then
                        RetString = RetString & c.Tag & "=" & c.ListIndex & ","
                    Else
                        RetString = RetString & c.Tag & "='" & c.Text & "',"
                    End If
                ElseIf (TypeOf c Is CheckBox) Then
                    RetString = RetString & c.Tag & "=" & Abs(c.Value) & ","
                ElseIf (TypeOf c Is ComboList.Usercontrol1) Then
                        RetString = RetString & c.Tag & "=" & c.ID & ","
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
    Dim str As String, iLen As Integer
    str = Trim(txtfields(15))
    iLen = Len(str)
    'if user select option for New NIC No.
    If NewID = True Then
        If iLen <> 13 Then
        
            If iLen = 15 Then
                If Mid$(str, 6, 1) = "-" And Mid$(str, 14, 1) = "-" Then
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
                If Mid$(str, 4, 1) = "-" And Mid$(str, 7, 1) = "-" Then
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


Private Sub txtfields_Validate(Index As Integer, cancel As Boolean)
    
    If Index = 15 Then
        If NIC(0).Value Then
            Call ConvertToID(True)
        Else
            Call ConvertToID(False)
        End If
    End If
    
End Sub

Private Sub ShowData()

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "Select * From VMakers_NewComing Where MakerID=" & VendID & " Order By MakerID", con, adOpenForwardOnly, adLockReadOnly
        VendID1 = ![MakerID1] & ""
'        cmbProcess = ![Description] & ""
        If Len(![NICNo]) = 13 Then
            NIC(1).Value = True
        Else
            NIC(0).Value = True
        End If
                
        For Each c In Controls
            If c.Tag <> "" Then
                'If TypeOf C Is CheckBox Then
                '    C.value = .Fields(C.Tag)
                'If TypeOf C Is ComboList.Usercontrol1 Then
                    'C.Text = .Fields(C.Tag & "Text") & ""
                '    C.ID = .Fields(C.Tag) & ""
                'If TypeOf C Is VB.Label Then
                '    C.Caption = .Fields(C.Tag) & ""
                If TypeOf c Is MSForms.ComboBox Then
                    If c.Name = "cmbType" Then
                        c.Text = ![MakerType] & ""
                    ElseIf c.Name = "cmbProcess" Then
                        c.Text = ![Description] & ""
                    ElseIf c.Name = "cmbProcess1" Then
                        c.ListIndex = Val(![SubVendType] & "")
                    ElseIf c.Name = "cmbAuthReq" Then
                        If IsNull(!AuthRequired) Then
                            c.ListIndex = 0
                        Else
                            c.ListIndex = Abs(!AuthRequired)
                        End If
                    ElseIf c.Name = "cmbPaymentTerms" Then
                        c.Text = !PaymentTerms & ""
                    ElseIf c.Name = "cmbShowRate" Then
                        If IsNull(!ShowRateOnPO) Then
                            c.ListIndex = 0
                        Else
                            c.ListIndex = Abs(!ShowRateOnPO)
                        End If
                    End If
                ElseIf (TypeOf c Is CheckBox) Then
                    c.Value = Abs(.Fields(c.Tag))
                ElseIf TypeOf c Is TextBox Or (TypeOf c Is MSForms.TextBox) Then
                    c.Text = .Fields(c.Tag) & ""
                ElseIf TypeOf c Is ComboList.Usercontrol1 Then
                    c.ID = Val(.Fields(c.Tag) & "")
                
                End If
            End If
        Next
        
        If ![Planter] <> 0 Then
            chkPlanter.Value = True
            cmbPlanter.Visible = True
        End If
        If ![Stamper] <> 0 Then
            chkStamper.Value = True
            cmbStamper.Visible = True
        End If
        If ![Snaffer] <> 0 Then
            chkSnaffer.Value = True
            cmbSnaffer.Visible = True
        End If
        
        If ![OtherCompany1] & "" <> "" Then
            chkComp1.Value = True
            txtfields(16).Visible = True
        End If
        
        If ![OtherCompany2] & "" <> "" Then
            chkComp2.Value = True
            txtfields(17).Visible = True
        End If
        txtUrduFields(0) = !MakerNameUrdu & ""
        
'        cmbMasterCat.Enabled = False
'        For i = 0 To cmbMasterCat.ListCount - 1
'            If cmbMasterCat.List(i, 1) = !SubAccOf & "" Then
'                cmbMasterCat.ListIndex = i
'                Exit For
'            End If
'        Next
        .Close
        
        .Open "SELECT VendPic,VendThumb FROM Makers_NewComing WHERE MakerID=" & VendID, con, adOpenForwardOnly, adLockReadOnly
        
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
                
        .Close
       
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
    
    
    strVendID = Val(con.Execute("Select Max(Cast(SubString(VendID1,5,4) as int)) From Venders Where Left(VendID1,3)='REP'").Fields(0).Value & "")
    
    If strVendID = 0 Then
        strVendID = 5000
    End If
    
    strVendID = strVendID + 1
    
    txtfields(12) = "REP-" & Format(strVendID, "0000")
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub AddProcesses()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "SELECT * FROM Processes ORDER BY Sno", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            Set ITM = LV2.ListItems.add(, ![ProcessID] & "'", ![SNo] & ". " & ![Description])
            ITM.ListSubItems.add , , !code & ""
            ITM.ListSubItems.add
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

