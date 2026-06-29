VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCompanyAE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Company Profile"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10020
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
   ScaleHeight     =   8625
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   45
      TabIndex        =   23
      Top             =   0
      Width           =   9930
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Profile"
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
         Left            =   3315
         TabIndex        =   24
         Top             =   120
         Width           =   2610
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Profile"
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
         Left            =   3330
         TabIndex        =   25
         Top             =   135
         Width           =   2610
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
      Height          =   8010
      Left            =   45
      TabIndex        =   22
      Top             =   570
      Width           =   9930
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
         Height          =   630
         Left            =   4560
         TabIndex        =   26
         Top             =   7275
         Width           =   5280
         Begin MSForms.CommandButton cmdSave 
            Height          =   360
            Left            =   75
            TabIndex        =   19
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
         Begin MSForms.CommandButton cmdCancel 
            Height          =   360
            Left            =   3510
            TabIndex        =   21
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
         Begin MSForms.CommandButton cmdSaveClose 
            Height          =   360
            Left            =   1815
            TabIndex        =   20
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
      Begin TabDlg.SSTab SSTab1 
         Height          =   7020
         Left            =   150
         TabIndex        =   27
         Top             =   240
         Width           =   9690
         _ExtentX        =   17092
         _ExtentY        =   12383
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   617
         TabCaption(0)   =   " Company Info."
         TabPicture(0)   =   "frmCompanyAE.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblLabels(1)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblLabels(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblLabels(2)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lblLabels(3)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lblLabels(5)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "lblLabels(8)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "lblLabels(9)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "lblLabels(10)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "lblLabels(11)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "lblLabels(12)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "lblLabels(13)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "lblLabels(14)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "lblLabels(15)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "lblLabels(16)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "lblLabels(17)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "lblLabels(18)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "lblLabels(19)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "DTPicker2"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "txtName"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Text1"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "Text2"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "DTPicker1"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "Text3"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "Text4"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "Text5"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "Text6"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "Text7"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "Text8"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "Text9"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "Text10"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "Text11"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "Text12"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "Text13"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "Text14"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).ControlCount=   34
         TabCaption(1)   =   "Reporting"
         TabPicture(1)   =   "frmCompanyAE.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtfields(3)"
         Tab(1).Control(1)=   "txtfields(0)"
         Tab(1).Control(2)=   "cmdViewPicture"
         Tab(1).Control(3)=   "CD1"
         Tab(1).Control(4)=   "lblLabels(7)"
         Tab(1).Control(5)=   "lblLabels(6)"
         Tab(1).Control(6)=   "lblLabels(4)"
         Tab(1).Control(7)=   "lblPic"
         Tab(1).Control(8)=   "Image1"
         Tab(1).ControlCount=   9
         Begin VB.TextBox Text14 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   330
            TabIndex        =   16
            Tag             =   "Internet"
            Top             =   6375
            Width           =   5235
         End
         Begin VB.TextBox Text13 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   7650
            TabIndex        =   15
            Tag             =   "SalesTaxNo"
            Top             =   5445
            Width           =   1710
         End
         Begin VB.TextBox Text12 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   330
            MultiLine       =   -1  'True
            TabIndex        =   5
            Tag             =   "OriginalAddress"
            Top             =   3420
            Width           =   9030
         End
         Begin VB.TextBox Text11 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5910
            TabIndex        =   14
            Tag             =   "Email1"
            Top             =   5445
            Width           =   1710
         End
         Begin VB.TextBox Text10 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3810
            TabIndex        =   13
            Tag             =   "Email"
            Top             =   5445
            Width           =   2085
         End
         Begin VB.TextBox Text9 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2070
            TabIndex        =   12
            Tag             =   "EFax"
            Top             =   5445
            Width           =   1710
         End
         Begin VB.TextBox Text8 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   330
            TabIndex        =   11
            Tag             =   "FaxNo"
            Top             =   5445
            Width           =   1710
         End
         Begin VB.TextBox Text7 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   7650
            TabIndex        =   10
            Tag             =   "PhoneNo1"
            Top             =   4560
            Width           =   1710
         End
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5910
            TabIndex        =   9
            Tag             =   "PhoneNo"
            Top             =   4560
            Width           =   1710
         End
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2070
            TabIndex        =   7
            Tag             =   "ExportRegNo"
            Top             =   4560
            Width           =   1710
         End
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   330
            TabIndex        =   6
            Tag             =   "NTNNo"
            Top             =   4560
            Width           =   1710
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   330
            MultiLine       =   -1  'True
            TabIndex        =   4
            Tag             =   "Address"
            Top             =   2295
            Width           =   9030
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   405
            Left            =   7350
            TabIndex        =   3
            Tag             =   "IssDT"
            Top             =   1575
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   714
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CheckBox        =   -1  'True
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   116588547
            CurrentDate     =   41071
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2085
            TabIndex        =   2
            Tag             =   "NICNo"
            Top             =   1575
            Width           =   5235
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   330
            TabIndex        =   1
            Tag             =   "InvoicePrefix"
            Top             =   1575
            Width           =   1740
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "CEmail"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Index           =   3
            Left            =   -74520
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   18
            Tag             =   "ReportSubHead"
            Top             =   5775
            Width           =   8520
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "CEmail"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   0
            Left            =   -74520
            MaxLength       =   100
            TabIndex        =   17
            Tag             =   "ReportMainHead"
            Top             =   5085
            Width           =   8520
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
            Left            =   -72555
            TabIndex        =   29
            Top             =   4395
            Width           =   4290
         End
         Begin VB.TextBox txtName 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   330
            TabIndex        =   0
            Tag             =   "CompanyName"
            Top             =   825
            Width           =   9030
         End
         Begin MSComDlg.CommonDialog CD1 
            Left            =   -66570
            Top             =   360
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            CancelError     =   -1  'True
            DialogTitle     =   "Please Select Picture"
            Filter          =   "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   405
            Left            =   3810
            TabIndex        =   8
            Tag             =   "ExportRegNoDT"
            Top             =   4560
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   714
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CheckBox        =   -1  'True
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   116588547
            CurrentDate     =   41071
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Web Address"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   19
            Left            =   360
            TabIndex        =   49
            Top             =   6120
            Width           =   1350
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Tax No."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   18
            Left            =   7620
            TabIndex        =   48
            Top             =   5175
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Original Address"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   17
            Left            =   360
            TabIndex        =   47
            Top             =   3135
            Width           =   1680
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email 2"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   16
            Left            =   5880
            TabIndex        =   46
            Top             =   5175
            Width           =   735
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email 1"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   15
            Left            =   3855
            TabIndex        =   45
            Top             =   5175
            Width           =   735
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E Fax"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   14
            Left            =   2100
            TabIndex        =   44
            Top             =   5175
            Width           =   510
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax No."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   13
            Left            =   360
            TabIndex        =   43
            Top             =   5175
            Width           =   750
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone 2"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   12
            Left            =   7620
            TabIndex        =   42
            Top             =   4260
            Width           =   840
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone 1"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   11
            Left            =   5880
            TabIndex        =   41
            Top             =   4290
            Width           =   840
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   10
            Left            =   3840
            TabIndex        =   40
            Top             =   4290
            Width           =   480
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Export Reg. No."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   9
            Left            =   2100
            TabIndex        =   39
            Top             =   4290
            Width           =   1590
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NTN No."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   8
            Left            =   360
            TabIndex        =   38
            Top             =   4290
            Width           =   870
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   5
            Left            =   360
            TabIndex        =   37
            Top             =   2040
            Width           =   810
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   3
            Left            =   7350
            TabIndex        =   36
            Top             =   1305
            Width           =   480
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CNIC #"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   2
            Left            =   2085
            TabIndex        =   35
            Top             =   1320
            Width           =   645
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice Prefix"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   0
            Left            =   360
            TabIndex        =   34
            Top             =   1305
            Width           =   1380
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   7
            Left            =   -74520
            TabIndex        =   33
            Top             =   5505
            Width           =   810
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   6
            Left            =   -74490
            TabIndex        =   32
            Top             =   4830
            Width           =   600
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Logo"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   4
            Left            =   -72540
            TabIndex        =   31
            Top             =   495
            Width           =   495
         End
         Begin VB.Label lblPic 
            Alignment       =   2  'Center
            Caption         =   "Double click to Add"
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   -70800
            TabIndex        =   30
            Top             =   2430
            Width           =   1125
         End
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   3525
            Left            =   -72555
            Stretch         =   -1  'True
            Top             =   825
            Width           =   4290
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   1
            Left            =   360
            TabIndex        =   28
            Top             =   555
            Width           =   600
         End
      End
   End
End
Attribute VB_Name = "frmCompanyAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lEntryID As Long

Dim strPicFileName As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    On Error GoTo err
    If Trim(txtName) = "" Then
        MsgBox "Invalid Name.", vbInformation
        Exit Sub
    End If

    If Saved Then
        Unload Me
        Dim f As New frmCompanyAE
        Load f
        f.AddNew
    End If
    
    Exit Sub

err:
    MsgBox "ERROR : # " & err.Number & vbNewLine & err.Description, vbInformation
 
End Sub


Private Sub cmdSaveClose_Click()

    On Error GoTo err
    
    If Trim(txtName) = "" Then
        MsgBox "Invalid Name.", vbInformation
        Exit Sub
    End If

    If Saved Then Unload Me

    Exit Sub
err:
    MsgBox "ERROR : # " & err.Number & vbNewLine & err.Description, vbInformation
 
End Sub

Private Sub cmdViewPicture_Click()
    
    If Image1.Picture = 0 Then Exit Sub
    Dim f As New frmViewPictureFSExt
    Load f
    f.ShowFromPic Image1.Picture

End Sub


Private Sub Form_Load()
    
    SSTab1.Tab = 0
     
    strPicFileName = ""
    
    Me.KeyPreview = True
    lEntryID = 0
     
End Sub

Private Sub Image1_DblClick()
    Call AddPic
End Sub



Private Sub lblPic_DblClick()
    Call AddPic
End Sub


Private Function CreateInsertSQL(strTable As String) As String
    
    Dim c As Control
    Dim RetString As String
    
    
    'Create The Field Name Portion of Qry...
    RetString = "INSERT INTO " & strTable & "("
    '---------------------------------------------------------------------------
    For Each c In Controls
        
        If (TypeOf c Is VB.TextBox) Or (TypeOf c Is VB.ComboBox) Or (TypeOf c Is MSForms.ComboBox) Or (TypeOf c Is MSForms.CheckBox) Or (TypeOf c Is ComboList.Usercontrol1) Then
            If c.Tag <> "" Then
                RetString = RetString & c.Tag & ","
            End If
        End If
      
    Next
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    RetString = RetString & ") VALUES("
    
    
    ' Create The Values Portion of Qry....
    '---------------------------------------------------------------------------
    For Each c In Controls
        
        
        If c.Tag <> "" Then
            If (TypeOf c Is VB.TextBox) Or (TypeOf c Is VB.ComboBox) Or (TypeOf c Is MSForms.ComboBox) Then
            
                If c.Name = "NBox" Then
                    RetString = RetString & Val(c.Text & "") & ","
                ElseIf c.Name = "cmbDisPort" Then
                    RetString = RetString & iDisID & ","
                Else
                    RetString = RetString & "'" & c.Text & "',"
                End If
            ElseIf (TypeOf c Is MSForms.CheckBox) Then
                If c.Name = "chkinactive" Then
                    RetString = RetString & Abs(Not c.Value) & ","
                Else
                    RetString = RetString & Abs(c.Value) & ","
                End If
            
            ElseIf (TypeOf c Is ComboList.Usercontrol1) Then
                RetString = RetString & "'" & c.Text & "',"
            
            End If
        End If
       
    Next
    
      
    'Finalaize the Qry
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
            If (TypeOf c Is VB.TextBox) Or (TypeOf c Is VB.ComboBox) Or (TypeOf c Is MSForms.ComboBox) Then '
                If c.Name = "cmbDisPort" Then
                    RetString = RetString & c.Tag & "=" & iDisID & ","
                ElseIf c.Name = "NBox" Then
                    RetString = RetString & c.Tag & "=" & c.Text & ","
                Else
                    RetString = RetString & c.Tag & "='" & c.Text & "',"
                End If
            ElseIf (TypeOf c Is MSForms.CheckBox) Then
                If c.Name = "chkinactive" Then
                    RetString = RetString & c.Tag & "=" & Abs(Not c.Value) & ","
                Else
                    RetString = RetString & c.Tag & "=" & Abs(c.Value) & ","
                End If
            ElseIf (TypeOf c Is ComboList.Usercontrol1) Then
                RetString = RetString & c.Tag & "='" & c.Text & "',"
            End If
        End If
    Next
    
  
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    
    CreateUpdateSQL = RetString
    
End Function


Public Sub EditCompany(p_lEntryID As Long)

    On Error GoTo err
    Dim rs As New ADODB.Recordset, ITM As ListItem
    
    lEntryID = p_lEntryID
    
    With rs
        .Open "SELECT * FROM Companies WHERE EntryID=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
        
        For Each c In Controls
            If c.Tag <> "" Then
                If TypeOf c Is MSForms.CheckBox Then
                    If c.Name = "chkinactive" Then
                        c.Value = Not .Fields(c.Tag)
                    Else
                        c.Value = .Fields(c.Tag)
                    End If
                ElseIf TypeOf c Is DTPicker Then
                    c.Value = .Fields(c.Tag) & ""
                ElseIf TypeOf c Is ComboList.Usercontrol1 Then
                    c.Text = .Fields(c.Tag) & ""
                Else
                    c.Text = .Fields(c.Tag) & ""
                End If
            End If
        Next
        
        If Not IsNull(![LogoForReports]) Then
            Set Image1.DataSource = rs
            Image1.DataField = "LogoForReports"
            lblPic.Visible = False
        Else
            lblPic.Visible = True
        End If
            
        .Close
    End With
   
    
   Set rs = Nothing
   Set Bnk = Nothing
    
   Me.Show 1
   
    
   Exit Sub
err:
    MsgBox err.Description
End Sub

Private Function Saved() As Boolean

   Dim SQL As String
    
    
    If lEntryID = 0 Then
        SQL = CreateInsertSQL("Companies")
    Else
       SQL = CreateUpdateSQL("Companies")
       SQL = SQL & " WHERE EntryID=" & lEntryID
    End If
     
    Dim lmyEntryID As Long
    If lEntryID = 0 Then
        lmyEntryID = GetSingleLongValue("MAX(EntryID)", "Companies")
    Else
        lmyEntryID = lEntryID
    End If
    
    Call StartTrans(con)
   
    con.Execute SQL, a
    Call SaveLogo(lmyEntryID)
   
    con.CommitTrans
    
    Saved = True
End Function

Private Sub SaveLogo(lmyEntryID As Long)

    If Image1.Picture <> 0 And Trim(strPicFileName) <> "" Then
        Dim Rec As New ADODB.Recordset
        Dim rsPic As New ADODB.Stream
        Rec.Open "SELECT EntryID,LogoForReports FROM Companies WHERE EntryID=" & lmyEntryID, con, adOpenStatic, adLockOptimistic
        If Rec.RecordCount > 0 Then
            rsPic.Open
            rsPic.Type = adTypeBinary
            rsPic.LoadFromFile strPicFileName
            Rec.Fields(1) = rsPic.Read
            rsPic.Close
            Rec.Update
        End If
        
        Rec.Close
        
        Set Rec = Nothing
        Set rsPic = Nothing
    End If
    
End Sub

Public Sub AddNew()
    
    Me.Show 1
    
End Sub

Private Sub AddPic()

    On Error Resume Next
    strPicFileName = ""
    CD1.Filter = "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
    CD1.ShowSave
    If Not (err.Number = 32755) Then
        strPicFileName = CD1.FileName
        Image1.Picture = LoadPicture(strPicFileName)
        lblPic.Visible = False
    Else
        strPicFileName = ""
        Image1.Picture = LoadPicture(strPicFileName)
        lblPic.Visible = True
    End If
    
End Sub
