VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmRMVender 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vender Info."
   ClientHeight    =   6480
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
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
      Height          =   5805
      Left            =   60
      TabIndex        =   4
      Top             =   600
      Width           =   8925
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Index           =   12
         Left            =   3900
         MaxLength       =   255
         TabIndex        =   60
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
         Left            =   195
         TabIndex        =   53
         Tag             =   "Inactive"
         Top             =   5265
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
         Top             =   5100
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
            Picture         =   "frmRMVender.frx":0000
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
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "General Info."
         TabPicture(0)   =   "frmRMVender.frx":0112
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
         TabPicture(1)   =   "frmRMVender.frx":012E
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "lblLabels(18)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "chkComp2"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "chkComp1"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "FAcc"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "txtfields(14)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "txtfields(16)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "txtfields(17)"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).ControlCount=   7
         TabCaption(2)   =   "Referal Info"
         TabPicture(2)   =   "frmRMVender.frx":014A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "FRefInfo"
         Tab(2).ControlCount=   1
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "Phone1"
            Height          =   315
            Index           =   17
            Left            =   1935
            MaxLength       =   20
            TabIndex        =   67
            Tag             =   "OtherCompany2"
            Top             =   2835
            Visible         =   0   'False
            Width           =   4350
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "Phone1"
            Height          =   315
            Index           =   16
            Left            =   1935
            MaxLength       =   20
            TabIndex        =   66
            Tag             =   "OtherCompany1"
            Top             =   2475
            Visible         =   0   'False
            Width           =   4350
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
            TabIndex        =   62
            Tag             =   "Experience"
            Top             =   1005
            Width           =   1590
         End
         Begin VB.Frame FRefInfo 
            Height          =   2370
            Left            =   -73755
            TabIndex        =   42
            Top             =   525
            Width           =   5850
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Address"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   28
               Left            =   2370
               TabIndex        =   52
               Top             =   945
               Width           =   585
            End
            Begin MSForms.TextBox TBox 
               Height          =   990
               Index           =   21
               Left            =   2340
               TabIndex        =   51
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
               TabIndex        =   50
               Top             =   1635
               Width           =   585
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   20
               Left            =   465
               TabIndex        =   49
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
               TabIndex        =   48
               Top             =   960
               Width           =   585
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   19
               Left            =   465
               TabIndex        =   47
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
               TabIndex        =   46
               Top             =   300
               Width           =   1035
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   18
               Left            =   2940
               TabIndex        =   45
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
               TabIndex        =   44
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
               TabIndex        =   43
               Top             =   315
               Width           =   1380
            End
         End
         Begin VB.Frame FAcc 
            Caption         =   "Account Information :"
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
            Height          =   870
            Left            =   150
            TabIndex        =   26
            Top             =   1410
            Width           =   7530
            Begin VB.ComboBox cmbDrCr 
               Appearance      =   0  'Flat
               Height          =   315
               ItemData        =   "frmRMVender.frx":0166
               Left            =   6060
               List            =   "frmRMVender.frx":0170
               Style           =   2  'Dropdown List
               TabIndex        =   29
               Top             =   450
               Width           =   1380
            End
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "AccNo"
               DataSource      =   "datPrimaryRS"
               Enabled         =   0   'False
               Height          =   315
               Index           =   0
               Left            =   135
               TabIndex        =   28
               Tag             =   "AccNo"
               Top             =   450
               Width           =   2250
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
               DataSource      =   "datPrimaryRS"
               Height          =   315
               Index           =   11
               Left            =   4680
               MaxLength       =   10
               TabIndex        =   27
               Top             =   450
               Width           =   1350
            End
            Begin MSComCtl2.DTPicker DtOpenDate 
               Height          =   315
               Left            =   2400
               TabIndex        =   30
               Top             =   450
               Width           =   2250
               _ExtentX        =   3969
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
               Format          =   54460419
               CurrentDate     =   37055
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Acc # :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   7
               Left            =   135
               TabIndex        =   33
               Top             =   225
               Width           =   525
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Opening Bal."
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   12
               Left            =   4710
               TabIndex        =   32
               Top             =   225
               Width           =   915
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Opening Date :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   11
               Left            =   2430
               TabIndex        =   31
               Top             =   225
               Width           =   1095
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
            Top             =   390
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
               TabIndex        =   57
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
               TabIndex        =   56
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
               TabIndex        =   36
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
               TabIndex        =   34
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
               TabIndex        =   59
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
               TabIndex        =   58
               Top             =   1470
               Width           =   720
            End
            Begin MSForms.OptionButton NIC 
               Height          =   285
               Index           =   1
               Left            =   570
               TabIndex        =   55
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
               TabIndex        =   54
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
               TabIndex        =   37
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
               TabIndex        =   35
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
         Begin MSForms.CheckBox chkComp1 
            Height          =   345
            Left            =   150
            TabIndex        =   65
            Top             =   2445
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
            Left            =   150
            TabIndex        =   64
            Top             =   2805
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
            TabIndex        =   63
            Top             =   780
            Width           =   1560
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   1260
            Left            =   -69870
            Stretch         =   -1  'True
            Top             =   2850
            Width           =   1545
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Double Click To Add Thumb Exp. /Signature"
            ForeColor       =   &H00FF0000&
            Height          =   615
            Left            =   -69705
            TabIndex        =   41
            Top             =   3225
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Thumb Exp/Signature"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   -69870
            TabIndex        =   40
            Top             =   2610
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
            TabIndex        =   39
            Top             =   420
            Width           =   495
         End
         Begin VB.Label lblPic 
            Alignment       =   2  'Center
            Caption         =   "Double Click To Add Picture."
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   -69645
            TabIndex        =   38
            Top             =   1290
            Width           =   1095
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
         TabIndex        =   61
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
Attribute VB_Name = "frmRMVender"
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
Public Sub EditVender(TempVendID As Integer)
    On Error GoTo err
    
    ShowingData = True
    Dim rs As New ADODB.Recordset
    
    VendID = TempVendID
    With rs
        .Open "Select * From RawMaterialVenders Where VendID=" & VendID & " Order By VendID", con, adOpenForwardOnly, adLockReadOnly
        VendID1 = ![VendID1] & ""
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
                ElseIf TypeOf C Is TextBox Or (TypeOf C Is MSForms.TextBox) Then
                    C.Text = .Fields(C.Tag) & ""
                ElseIf TypeOf C Is ComboList.Usercontrol1 Then
                    C.ID = Val(.Fields(C.Tag) & "")
                End If
            End If
        Next
        
        
        If ![OtherCompany1] & "" <> "" Then
            chkComp1.value = True
            txtfields(16).Visible = True
        End If
        
        If ![OtherCompany2] & "" <> "" Then
            chkComp2.value = True
            txtfields(17).Visible = True
        End If
        
        .Close
        
        .Open "Select VendPic,VendThumb From RawMaterialVenders Where VendID=" & VendID, con, adOpenForwardOnly, adLockReadOnly
        
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
    ShowingData = False
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

    If chkInactive = vbChecked And Val(lblBalance) <> 0 Then
        MsgBox "This Vender Cannot Be Made Inactive," & vbNewLine & _
         "To Make A Vender Inactive Balance Must Be Zero", vbInformation
        chkInactive = vbUnchecked
    End If
    
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


Private Sub GetVendID()
    On Error GoTo err
    If ShowingData Then Exit Sub
    Dim strVendID As String
    
    
    'strVendID = Val(con.Execute("Select Max(Cast(SubString(VendID1,5,4) as int)) From Venders Where UPPER(Left(VendID1,3))=" & UCase(Left(cmbProcess.Text, 3))).Fields(0).value & "")
    
    If VendID <> 0 Then
        
        txtfields(12) = VendID1
        Exit Sub
        
    End If
    
    
    strVendID = Val(con.Execute("Select Max(Cast(SubString(VendID1,5,4) as int)) From RawMaterialVenders").Fields(0).value & "")
    
    If strVendID = 0 Then
        strVendID = 1000
    End If
    
    strVendID = strVendID + 1
    
    txtfields(12) = "RMV" & Format(strVendID, "-0000")
    
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
    
    If Trim(txtfields(1)) = "" Then
        MsgBox "Invalid Vender Name", vbInformation
        Exit Function
    End If
    
    Call StartTrans(con)
    
    Me.MousePointer = vbHourglass
    Dim SQL As String
    
    If VendID = 0 Then
        SQL = CreateInsertSQL("RawMaterialVenders")
        'con.Execute "Update Accounts Set OpenBal=OpenBal+" & TotalCost & " Where AccNo='" & AssetAcc & "'"
        'Call adjustLedger(AssetAcc, "1/1/1900", 0, TotalCost)
    Else
        SQL = CreateUpdateSQL("RawMaterialVenders")
        SQL = SQL & " Where VendID=" & VendID
    End If
    
    con.Execute SQL
    
    
    'If Add Then
    '    Call AddAccountInfo
    'ElseIf Edit Then
    '    Call UpdateAccinfo
    'End If
    
    con.CommitTrans
    
SavePic:
    If (Image1.Picture = 0 Or Trim(PicFileName) = "") And (Image2.Picture = 0 Or Trim(ThumbFileName) = "") Then GoTo TheEnd

    Dim rec As New ADODB.Recordset
    Dim rsPic As New ADODB.Stream
    Dim rsThumb As New ADODB.Stream
    
    If VendID = 0 Then
        rec.Open "Select Max(VendID) From Venders", con, adOpenForwardOnly, adLockReadOnly
        VendID = rec.Fields(0) & ""
        rec.Close
    End If
    
    
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
    
Private Sub DtOpenDate_Change()
    Call setDateRange(DtOpenDate)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()

    iTotNo = 0
    Me.KeyPreview = True
    If Not AccDefined Then
        MsgBox "Account "" " & ACCRECTITLE & " "" is Not Defined, Please Define The Account First", vbInformation
    End If

    'Call AddToCombo(cmbType, "AccTitle", "Accounts", "Where SubAccof Like '" & AcNo & "%' And Parent=1")
    'cmbType.Enabled = cmbType.ListCount > 0

    DtOpenDate = ServerDate
    
    SSTab1.Tab = 0
    Call GetVendID
    NewVal = GetMax("right(AccNo,3)", "Accounts", "Where SubAccOf='" & ACCPAYNO & "'")
 
    txtfields(0) = ACCPAYNO & "-" & Format(Val(NewVal) + 1 + iTotNo, "000")
    
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

    Call cmbDrCr_Click
    Dim Active As Boolean
    Active = Active = chkInactive <> vbChecked
 
    Call NewAccPay(txtfields(1), DtOpenDate, Val(txtfields(11).Tag), Active, Left(txtfields(0), Len(txtfields(0)) - 4))
    
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
 
    Dim Diff As Double
    Dim Active As Boolean
    Call cmbDrCr_Click
  
    Diff = Val(txtfields(11).Tag) - Val(cmbDrCr.Tag)
    Active = chkInactive <> vbChecked
         
    con.Execute "Update Accounts Set accTitle='" & txtfields(1) & "',OpenBal=" & Val(txtfields(11).Tag) & ",Balance=balance + " & Diff & ",Active=" & Active & " Where Accno='" & txtfields(0) & "'"
      
    Call adjustLedger(txtfields(0), DtOpenDate, 0, Diff, True)

    If Diff <> 0 Then Call UpdateAdvEntry(txtfields(0), Val(txtfields(11).Tag))
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
    OurOwnForm = True
    Set frmVender = Nothing
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
                If (TypeOf C Is MSForms.TextBox) Or (TypeOf C Is MSForms.ComboBox) Or (TypeOf C Is ComboList.Usercontrol1) Or (TypeOf C Is TextBox) Then
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
                    ElseIf C.Name = "cmbSteel" Or C.Name = "cmbGuage" Or C.Name = "cmbItemGroup" Or C.Name = "cmbProcess" Then
                        RetString = RetString & C.List(C.ListIndex, 1) & ","
                    ElseIf C.Name = "cmbProcess1" Then
                        RetString = RetString & "'" & C.ListIndex & "',"
                    ElseIf C.Name = "cmbType" Then
                        RetString = RetString & "'" & C.ListIndex & "',"
                    Else
                        RetString = RetString & "'" & C.Text & "',"
                    End If
                    
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
            If C.Visible Then
                If (TypeOf C Is MSForms.TextBox) Or (TypeOf C Is MSForms.ComboBox) Or (TypeOf C Is MSForms.ComboBox) Or (TypeOf C Is TextBox) Then '
                    If C.Name = "NBox" Then
                        RetString = RetString & C.Tag & "=" & C.Text & ","
                    ElseIf C.Name = "cmbSteel" Or C.Name = "cmbGuage" Or C.Name = "cmbItemGroup" Or C.Name = "cmbProcess" Then
                        RetString = RetString & C.Tag & "=" & C.List(C.ListIndex, 1) & ","
                    ElseIf C.Name = "cmbProcess1" Then
                        RetString = RetString & "SubVendType='" & C.ListIndex & "',"
                    ElseIf C.Name = "cmbType" Then
                        RetString = RetString & C.Tag & "='" & C.ListIndex & "',"
                    Else
                        RetString = RetString & C.Tag & "='" & C.Text & "',"
                    End If
                'ElseIf (TypeOf C Is CheckBox) Then
                '    RetString = RetString & C.Tag & "=" & Abs(C.value) & ","
                ElseIf (TypeOf C Is ComboList.Usercontrol1) Then
                    RetString = RetString & C.Tag & "=" & C.ID & ","
                End If
            End If
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
