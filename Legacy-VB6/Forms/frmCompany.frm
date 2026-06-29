VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCompany 
   ClientHeight    =   7395
   ClientLeft      =   -180
   ClientTop       =   -180
   ClientWidth     =   8085
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
   ScaleHeight     =   7395
   ScaleWidth      =   8085
   WindowState     =   2  'Maximized
   Begin VB.Frame FAll 
      Height          =   7275
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   7920
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   7920
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Company Info."
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
            Left            =   60
            TabIndex        =   2
            Top             =   165
            Width           =   7710
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Company Info."
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
            Left            =   75
            TabIndex        =   3
            Top             =   180
            Width           =   7695
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5805
         Left            =   90
         TabIndex        =   4
         Top             =   750
         Width           =   7680
         _ExtentX        =   13547
         _ExtentY        =   10239
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
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
         TabCaption(0)   =   "General Info."
         TabPicture(0)   =   "frmCompany.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label2(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label2(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "TBox(0)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "TBox(1)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label2(2)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "TBox(2)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label2(3)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "TBox(3)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label2(4)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "TBox(4)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label2(5)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "TBox(5)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Label2(6)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "TBox(6)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Label2(7)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "TBox(7)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Label2(8)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "TBox(8)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).ControlCount=   18
         TabCaption(1)   =   "Other Info."
         TabPicture(1)   =   "frmCompany.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "DTPicker1"
         Tab(1).Control(1)=   "DTPicker2"
         Tab(1).Control(2)=   "DTPicker3"
         Tab(1).Control(3)=   "DTPicker4"
         Tab(1).Control(4)=   "DTPicker5"
         Tab(1).Control(5)=   "DTPicker6"
         Tab(1).Control(6)=   "DTPicker7"
         Tab(1).Control(7)=   "DTPicker8"
         Tab(1).Control(8)=   "DTPicker9"
         Tab(1).Control(9)=   "DTPicker10"
         Tab(1).Control(10)=   "DTPicker11"
         Tab(1).Control(11)=   "DTPicker12"
         Tab(1).Control(12)=   "Label1(36)"
         Tab(1).Control(13)=   "Label1(35)"
         Tab(1).Control(14)=   "Label1(34)"
         Tab(1).Control(15)=   "Label1(33)"
         Tab(1).Control(16)=   "Label1(32)"
         Tab(1).Control(17)=   "Label1(31)"
         Tab(1).Control(18)=   "Label1(30)"
         Tab(1).Control(19)=   "Label1(3)"
         Tab(1).Control(20)=   "TBox(17)"
         Tab(1).Control(21)=   "Label1(2)"
         Tab(1).Control(22)=   "TBox(16)"
         Tab(1).Control(23)=   "Label1(1)"
         Tab(1).Control(24)=   "TBox(15)"
         Tab(1).Control(25)=   "TBox(14)"
         Tab(1).Control(26)=   "TBox(13)"
         Tab(1).Control(27)=   "TBox(12)"
         Tab(1).Control(28)=   "TBox(11)"
         Tab(1).Control(29)=   "TBox(10)"
         Tab(1).Control(30)=   "TBox(9)"
         Tab(1).Control(31)=   "Label1(12)"
         Tab(1).Control(32)=   "Label1(10)"
         Tab(1).Control(33)=   "Label1(0)"
         Tab(1).Control(34)=   "Label1(7)"
         Tab(1).Control(35)=   "Label1(6)"
         Tab(1).Control(36)=   "Label1(5)"
         Tab(1).Control(37)=   "Label1(15)"
         Tab(1).Control(38)=   "Label1(16)"
         Tab(1).Control(39)=   "Label1(17)"
         Tab(1).Control(40)=   "Label1(28)"
         Tab(1).Control(41)=   "Label1(29)"
         Tab(1).ControlCount=   42
         TabCaption(2)   =   "Certifications"
         TabPicture(2)   =   "frmCompany.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "CDBarcode"
         Tab(2).Control(1)=   "LV"
         Tab(2).Control(2)=   "cmdBrowseBarcode"
         Tab(2).Control(3)=   "DTCertification"
         Tab(2).Control(4)=   "DTExpiryCertification"
         Tab(2).Control(5)=   "Label1(14)"
         Tab(2).Control(6)=   "txtScanPath"
         Tab(2).Control(7)=   "Label1(13)"
         Tab(2).Control(8)=   "Label1(11)"
         Tab(2).Control(9)=   "Label1(4)"
         Tab(2).Control(10)=   "txtCertification"
         Tab(2).ControlCount=   11
         TabCaption(3)   =   "Auth. Reps."
         TabPicture(3)   =   "frmCompany.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame3"
         Tab(3).Control(1)=   "Frame2"
         Tab(3).ControlCount=   2
         Begin VB.Frame Frame3 
            Caption         =   "Authorized Representative in EC :"
            Height          =   2520
            Left            =   -74775
            TabIndex        =   79
            Top             =   3120
            Width           =   7245
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
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
               Index           =   27
               Left            =   150
               TabIndex        =   84
               Top             =   315
               Width           =   480
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   27
               Left            =   150
               TabIndex        =   74
               Tag             =   "RepNameEC"
               Top             =   510
               Width           =   6840
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               Size            =   "12065;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
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
               Index           =   26
               Left            =   165
               TabIndex        =   83
               Top             =   870
               Width           =   690
            End
            Begin MSForms.TextBox TBox 
               Height          =   690
               Index           =   26
               Left            =   150
               TabIndex        =   75
               Tag             =   "RepAddressEC"
               Top             =   1080
               Width           =   6840
               VariousPropertyBits=   -1400879077
               BorderStyle     =   1
               Size            =   "12065;1217"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Phone"
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
               Index           =   25
               Left            =   180
               TabIndex        =   82
               Top             =   1845
               Width           =   525
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   25
               Left            =   150
               TabIndex        =   76
               Tag             =   "RepPhoneEC"
               Top             =   2040
               Width           =   1545
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               Size            =   "2725;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fax"
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
               Index           =   24
               Left            =   1755
               TabIndex        =   81
               Top             =   1860
               Width           =   300
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   24
               Left            =   1710
               TabIndex        =   77
               Tag             =   "RepFaxEC"
               Top             =   2040
               Width           =   1605
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               Size            =   "2831;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Contact Person"
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
               Index           =   23
               Left            =   3330
               TabIndex        =   80
               Top             =   1845
               Width           =   1290
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   23
               Left            =   3345
               TabIndex        =   78
               Tag             =   "RepContactEC"
               Top             =   2040
               Width           =   3645
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               Size            =   "6429;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Authorized Representative in US :"
            Height          =   2520
            Left            =   -74775
            TabIndex        =   63
            Top             =   525
            Width           =   7245
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   22
               Left            =   3345
               TabIndex        =   73
               Tag             =   "RepContactUS"
               Top             =   2040
               Width           =   3645
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               Size            =   "6429;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Contact Person"
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
               Index           =   22
               Left            =   3330
               TabIndex        =   72
               Top             =   1845
               Width           =   1290
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   21
               Left            =   1710
               TabIndex        =   71
               Tag             =   "RepFaxUS"
               Top             =   2040
               Width           =   1605
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               Size            =   "2831;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fax"
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
               Index           =   21
               Left            =   1755
               TabIndex        =   70
               Top             =   1860
               Width           =   300
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   20
               Left            =   150
               TabIndex        =   69
               Tag             =   "RepPhoneUS"
               Top             =   2040
               Width           =   1545
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               Size            =   "2725;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Phone"
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
               Index           =   20
               Left            =   180
               TabIndex        =   68
               Top             =   1845
               Width           =   525
            End
            Begin MSForms.TextBox TBox 
               Height          =   690
               Index           =   19
               Left            =   150
               TabIndex        =   67
               Tag             =   "RepAddressUS"
               Top             =   1080
               Width           =   6840
               VariousPropertyBits=   -1400879077
               BorderStyle     =   1
               Size            =   "12065;1217"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
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
               Index           =   19
               Left            =   165
               TabIndex        =   66
               Top             =   870
               Width           =   690
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   18
               Left            =   150
               TabIndex        =   65
               Tag             =   "RepNameUS"
               Top             =   510
               Width           =   6840
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               Size            =   "12065;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
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
               Index           =   18
               Left            =   150
               TabIndex        =   64
               Top             =   315
               Width           =   480
            End
         End
         Begin MSComDlg.CommonDialog CDBarcode 
            Left            =   -74970
            Top             =   5190
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            CancelError     =   -1  'True
         End
         Begin MSComctlLib.ListView LV 
            Height          =   4545
            Left            =   -74880
            TabIndex        =   62
            Top             =   1125
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   8017
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Certification"
               Object.Width           =   3942
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Date"
               Object.Width           =   2381
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Expiry Date"
               Object.Width           =   2381
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Scanned Image"
               Object.Width           =   2381
            EndProperty
         End
         Begin VB.CommandButton cmdBrowseBarcode 
            Caption         =   "..."
            Height          =   330
            Left            =   -67845
            TabIndex        =   61
            Top             =   765
            Width           =   405
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   -71085
            TabIndex        =   41
            Tag             =   "NTNDT"
            Top             =   1320
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   115605507
            CurrentDate     =   37865
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   -71085
            TabIndex        =   42
            Tag             =   "STaxDT"
            Top             =   1950
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   115605507
            CurrentDate     =   37865
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   315
            Left            =   -71085
            TabIndex        =   43
            Tag             =   "ChamberDT"
            Top             =   2595
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   115605507
            CurrentDate     =   37865
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   315
            Left            =   -71085
            TabIndex        =   44
            Tag             =   "ExportDT"
            Top             =   3225
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   115605507
            CurrentDate     =   37865
         End
         Begin MSComCtl2.DTPicker DTPicker5 
            Height          =   315
            Left            =   -71085
            TabIndex        =   45
            Tag             =   "ImportDT"
            Top             =   3870
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   115605507
            CurrentDate     =   37865
         End
         Begin MSComCtl2.DTPicker DTCertification 
            Height          =   315
            Left            =   -72615
            TabIndex        =   55
            Top             =   780
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   115605507
            CurrentDate     =   37865
         End
         Begin MSComCtl2.DTPicker DTExpiryCertification 
            Height          =   315
            Left            =   -71235
            TabIndex        =   57
            Top             =   780
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   115605507
            CurrentDate     =   37865
         End
         Begin MSComCtl2.DTPicker DTPicker6 
            Height          =   315
            Left            =   -69300
            TabIndex        =   85
            Tag             =   "NTNDTExpiry"
            Top             =   1320
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   115605507
            CurrentDate     =   37865
         End
         Begin MSComCtl2.DTPicker DTPicker7 
            Height          =   315
            Left            =   -69300
            TabIndex        =   87
            Tag             =   "STaxDTExpiry"
            Top             =   1950
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   115605507
            CurrentDate     =   37865
         End
         Begin MSComCtl2.DTPicker DTPicker8 
            Height          =   315
            Left            =   -69300
            TabIndex        =   88
            Tag             =   "ChamberDTExpiry"
            Top             =   2595
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   115605507
            CurrentDate     =   37865
         End
         Begin MSComCtl2.DTPicker DTPicker9 
            Height          =   315
            Left            =   -69300
            TabIndex        =   89
            Tag             =   "ExportDTExpiry"
            Top             =   3225
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   115605507
            CurrentDate     =   37865
         End
         Begin MSComCtl2.DTPicker DTPicker10 
            Height          =   315
            Left            =   -69300
            TabIndex        =   90
            Tag             =   "ImportDTExpiry"
            Top             =   3870
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   115605507
            CurrentDate     =   37865
         End
         Begin MSComCtl2.DTPicker DTPicker11 
            Height          =   315
            Left            =   -71085
            TabIndex        =   95
            Tag             =   "SurgAssDT"
            Top             =   4530
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   115605507
            CurrentDate     =   37865
         End
         Begin MSComCtl2.DTPicker DTPicker12 
            Height          =   315
            Left            =   -69300
            TabIndex        =   97
            Tag             =   "SurgAssDTExpiry"
            Top             =   4530
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   115605507
            CurrentDate     =   37865
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Expiry Date"
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
            Index           =   36
            Left            =   -69285
            TabIndex        =   98
            Top             =   4290
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reg Date"
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
            Index           =   35
            Left            =   -71085
            TabIndex        =   96
            Top             =   4290
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Expiry Date"
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
            Index           =   34
            Left            =   -69255
            TabIndex        =   94
            Top             =   3600
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Expiry Date"
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
            Index           =   33
            Left            =   -69255
            TabIndex        =   93
            Top             =   2370
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Expiry Date"
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
            Index           =   32
            Left            =   -69255
            TabIndex        =   92
            Top             =   2955
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Expiry Date"
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
            Index           =   31
            Left            =   -69255
            TabIndex        =   91
            Top             =   1725
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Expiry Date"
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
            Index           =   30
            Left            =   -69270
            TabIndex        =   86
            Top             =   1095
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Scanned Image's Path"
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
            Index           =   14
            Left            =   -69840
            TabIndex        =   60
            Top             =   570
            Width           =   1890
         End
         Begin MSForms.TextBox txtScanPath 
            Height          =   315
            Left            =   -69870
            TabIndex        =   59
            Top             =   765
            Width           =   2010
            VariousPropertyBits=   746604571
            ForeColor       =   4194304
            BorderStyle     =   1
            Size            =   "3545;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Expiry Date"
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
            Index           =   13
            Left            =   -71220
            TabIndex        =   58
            Top             =   570
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
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
            Index           =   11
            Left            =   -72600
            TabIndex        =   56
            Top             =   570
            Width           =   405
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Certification"
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
            Index           =   4
            Left            =   -74850
            TabIndex        =   54
            Top             =   570
            Width           =   1035
         End
         Begin MSForms.TextBox txtCertification 
            Height          =   315
            Left            =   -74880
            TabIndex        =   53
            Top             =   780
            Width           =   2235
            VariousPropertyBits=   746604571
            ForeColor       =   4194304
            BorderStyle     =   1
            Size            =   "3942;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DHL #"
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
            Index           =   3
            Left            =   -74700
            TabIndex        =   52
            Top             =   4935
            Width           =   510
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   17
            Left            =   -74730
            TabIndex        =   51
            Tag             =   "DHLNo"
            Top             =   5175
            Width           =   2385
            VariousPropertyBits=   746604571
            ForeColor       =   4194304
            BorderStyle     =   1
            Size            =   "4207;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FDA Registration #"
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
            Index           =   2
            Left            =   -72330
            TabIndex        =   50
            Top             =   4965
            Width           =   1605
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   16
            Left            =   -72315
            TabIndex        =   49
            Tag             =   "FDANo"
            Top             =   5175
            Width           =   2355
            VariousPropertyBits=   746604571
            ForeColor       =   4194304
            BorderStyle     =   1
            Size            =   "4154;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Surgical Association Registration #"
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
            Left            =   -74685
            TabIndex        =   48
            Top             =   4290
            Width           =   2970
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   15
            Left            =   -74730
            TabIndex        =   47
            Tag             =   "SurgicalAssociationNo"
            Top             =   4530
            Width           =   3510
            VariousPropertyBits=   746604571
            ForeColor       =   4194304
            BorderStyle     =   1
            Size            =   "6191;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   14
            Left            =   -74730
            TabIndex        =   40
            Tag             =   "ImportNo"
            Top             =   3864
            Width           =   3510
            VariousPropertyBits=   746604571
            ForeColor       =   4194304
            BorderStyle     =   1
            Size            =   "6191;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   13
            Left            =   -69930
            TabIndex        =   39
            Tag             =   "InsAgent"
            Top             =   5175
            Width           =   2385
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "4207;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   12
            Left            =   -74730
            TabIndex        =   38
            Tag             =   "Chamber"
            Top             =   2592
            Width           =   3510
            VariousPropertyBits=   746604571
            ForeColor       =   4194304
            BorderStyle     =   1
            Size            =   "6191;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   11
            Left            =   -74730
            TabIndex        =   37
            Tag             =   "ExportNo"
            Top             =   3228
            Width           =   3510
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "6191;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   10
            Left            =   -74730
            TabIndex        =   36
            Tag             =   "NTN"
            Top             =   1320
            Width           =   3510
            VariousPropertyBits=   746604571
            ForeColor       =   4194304
            BorderStyle     =   1
            Size            =   "6191;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   9
            Left            =   -74730
            TabIndex        =   35
            Tag             =   "STax"
            Top             =   1956
            Width           =   3510
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "6191;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Chamber #"
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
            Index           =   12
            Left            =   -74700
            TabIndex        =   34
            Top             =   2385
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Insurance Agent"
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
            Index           =   10
            Left            =   -69930
            TabIndex        =   33
            Top             =   4965
            Width           =   1410
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Export Reg #"
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
            Index           =   0
            Left            =   -74700
            TabIndex        =   32
            Top             =   2985
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reg Date"
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
            Index           =   7
            Left            =   -71055
            TabIndex        =   31
            Top             =   1740
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Tax #"
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
            Index           =   6
            Left            =   -74700
            TabIndex        =   30
            Top             =   1755
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N.T.N #"
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
            Index           =   5
            Left            =   -74700
            TabIndex        =   29
            Top             =   1095
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reg Date"
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
            Index           =   15
            Left            =   -71055
            TabIndex        =   28
            Top             =   1095
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reg Date"
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
            Index           =   16
            Left            =   -71055
            TabIndex        =   27
            Top             =   2970
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reg Date"
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
            Left            =   -71055
            TabIndex        =   26
            Top             =   2385
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reg Date"
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
            Index           =   28
            Left            =   -71055
            TabIndex        =   25
            Top             =   3615
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Import Reg #"
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
            Index           =   29
            Left            =   -74700
            TabIndex        =   24
            Top             =   3630
            Width           =   1155
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   8
            Left            =   2640
            TabIndex        =   22
            Tag             =   "Web"
            Top             =   5295
            Width           =   4680
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "8255;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Web Address"
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
            Index           =   8
            Left            =   2745
            TabIndex        =   21
            Top             =   5085
            Width           =   1110
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   7
            Left            =   285
            TabIndex        =   20
            Tag             =   "EMail"
            Top             =   5295
            Width           =   2310
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "4075;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "E-Mail Address"
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
            Index           =   7
            Left            =   270
            TabIndex        =   19
            Top             =   5085
            Width           =   1245
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   6
            Left            =   3810
            TabIndex        =   18
            Tag             =   "Fax2"
            Top             =   4710
            Width           =   3510
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "6191;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fax 2"
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
            Index           =   6
            Left            =   3840
            TabIndex        =   17
            Top             =   4500
            Width           =   450
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   5
            Left            =   270
            TabIndex        =   16
            Tag             =   "Fax1"
            Top             =   4710
            Width           =   3510
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "6191;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fax 1"
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
            Index           =   5
            Left            =   285
            TabIndex        =   15
            Top             =   4500
            Width           =   450
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   4
            Left            =   5010
            TabIndex        =   14
            Tag             =   "Phone3"
            Top             =   4140
            Width           =   2310
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "4075;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Phone 3"
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
            Index           =   4
            Left            =   5025
            TabIndex        =   13
            Top             =   3930
            Width           =   675
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   3
            Left            =   2640
            TabIndex        =   12
            Tag             =   "Phone2"
            Top             =   4140
            Width           =   2325
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "4101;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Phone 2"
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
            Index           =   3
            Left            =   2685
            TabIndex        =   11
            Top             =   3930
            Width           =   675
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   2
            Left            =   285
            TabIndex        =   10
            Tag             =   "Phone1"
            Top             =   4140
            Width           =   2310
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "4075;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Phone 1"
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
            Index           =   2
            Left            =   300
            TabIndex        =   9
            Top             =   3930
            Width           =   675
         End
         Begin MSForms.TextBox TBox 
            Height          =   2280
            Index           =   1
            Left            =   270
            TabIndex        =   8
            Tag             =   "Address"
            Top             =   1545
            Width           =   7050
            VariousPropertyBits=   -1400879077
            BorderStyle     =   1
            ScrollBars      =   3
            Size            =   "12435;4022"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   0
            Left            =   270
            TabIndex        =   7
            Tag             =   "CompanyName"
            Top             =   885
            Width           =   7050
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "12435;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Company Address"
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
            Left            =   300
            TabIndex        =   6
            Top             =   1290
            Width           =   1530
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Company Name"
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
            Index           =   0
            Left            =   285
            TabIndex        =   5
            Top             =   675
            Width           =   1320
         End
      End
      Begin MSForms.CommandButton cmdSClose 
         Height          =   390
         Left            =   4140
         TabIndex        =   46
         Top             =   6705
         Width           =   1800
         Caption         =   "Save & Close  "
         PicturePosition =   327683
         Size            =   "3175;688"
         Picture         =   "frmCompany.frx":0070
         Accelerator     =   83
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   405
         Left            =   5970
         TabIndex        =   23
         Top             =   6690
         Width           =   1800
         Caption         =   "Close            "
         PicturePosition =   327683
         Size            =   "3175;714"
         Picture         =   "frmCompany.frx":04BE
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
End
Attribute VB_Name = "frmCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBrowseBarcode_Click()

    On Error GoTo err
    CDBarcode.Filter = "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
    On Error Resume Next
    CDBarcode.ShowOpen
    If err.Number = 32755 Then Exit Sub
    
    txtScanPath = CDBarcode.FileName
    
    
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSClose_Click()

    On Error GoTo err
    
    Dim str As String
    str = CreateUpdateSQL("Company")
    
    Call StartTrans(con)
    
    con.Execute str
    
    Dim i As Integer
    Dim lEntryID As Long
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If Val(.Tag) = "0" Then
                con.Execute "INSERT INTO CompanyCertifications(Certification,DT,ExpiryDT) VALUES('" & .Text & _
                 "','" & .SubItems(1) & "','" & .SubItems(2) & "')"
                         
                lEntryID = GetSingleLongValue("MAX(EntryID)", "CompanyCertifications")
                
                Dim Rec As New ADODB.Recordset
                Dim rsPic As New ADODB.Stream
                    
                Rec.Open "SELECT EntryID,ScannedCopy FROM CompanyCertifications WHERE EntryID=" & lEntryID, con, adOpenStatic, adLockOptimistic
                If .SubItems(3) <> "" Then
                    rsPic.Open
                    rsPic.Type = adTypeBinary
                    rsPic.LoadFromFile .SubItems(3)
                    Rec.Fields(1) = rsPic.Read
                    rsPic.Close
                End If
                Rec.Update
                
            End If
        End With
    Next
    
    con.CommitTrans
    
    Unload Me
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    SSTab1.Tab = 0
    Call EditCompany
    
    DTCertification = Date
    DTExpiryCertification = Date
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    FAll.Move ((Me.ScaleWidth - FAll.Width) / 2), ((Me.ScaleHeight - FAll.Height) / 2)
End Sub
Private Function CreateUpdateSQL(strTable As String) As String
    
    Dim c As Control
    Dim RetString As String
    
    'Create The First Part Of Insert Query...
    RetString = "Update " & strTable & " Set "
    '---------------------------------------------------------------------------
    For Each c In Controls
    
        If c.Tag <> "" Then
            If (TypeOf c Is MSForms.TextBox) Then '
                RetString = RetString & c.Tag & "='" & c.Text & "',"
            ElseIf (TypeOf c Is DTPicker) Then
                RetString = RetString & c.Tag & "='" & (c.Value) & "',"
            
            End If
        End If
    Next
    
  
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    
    CreateUpdateSQL = RetString
    
End Function


Private Sub EditCompany()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "Select * From Company ", con, adOpenForwardOnly, adLockReadOnly
        
            For Each c In Controls
                If c.Tag <> "" Then
                    If TypeOf c Is DTPicker Then
                        c.Value = .Fields(c.Tag) & ""
                    Else
                        c.Text = .Fields(c.Tag) & ""
                    End If
                End If
            Next
            
        .Close
        
        .Open "SELECT EntryID,Certification,DT,ExpiryDT FROM CompanyCertifications ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, !EntryID & "'", !Certification & "")
            ITM.Tag = !EntryID
            ITM.ListSubItems.add , , Format(!DT, "dd-MMM-yyyy")
            ITM.ListSubItems.add , , Format(!ExpiryDT, "dd-MMM-yyyy")
            ITM.ListSubItems.add , , "<<Dbl Click to View>>"
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Me.Show
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    If LV.SelectedItem.SubItems(3) = "" Then Exit Sub
    '---------------------------------------------------
'    Dim myPic As StdPicture
'    Dim rs As New ADODB.Recordset
'    With rs
'        .Open "SELECT ScannedCopy FROM CompanyCertifications WHERE EntryID=" & Val(LV.SelectedItem.key), con, adOpenForwardOnly, adLockReadOnly
'        If .EOF = False Then
'            Set Image1.DataSource = rs
'            Image1.DataField = "ScannedCopy"
'        End If
'        .Close
'    End With
'    Set rs = Nothing
'
'    If Image1.Picture <> 0 Then
'        Dim f As New frmViewPictureFSExt
'        Load f
'        f.ShowFromPic Image1.Picture
'        f.ShowFromDBField
'    End If
    
    Dim f As New frmViewPictureFSExt
    Load f

    f.ShowFromDBField "ScannedCopy", "CompanyCertifications", " WHERE EntryID=" & Val(LV.SelectedItem.key)
        
    'Call ShowTechDrawing("CompanyCertifications", "ScannedCopy", " WHERE EntryID=" & Val(LV.SelectedItem.key), "", CDBarcode)
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If LV.ListItems.count = 0 Then Exit Sub
    If Button <> 2 Then Exit Sub
    Dim strRet As String
    strRet = ShowPopUpMenu1("Print Certificate")
    If strRet = "Print Certificate" Then
        Dim rpt As CRAXDDRT.Report
        Set rpt = rptApp.OpenReport(RptPath & "\CompanyCertificate.rpt")
        Dim f As New frmPrevRpt
        f.ShowReport "{CompanyCertifications.EntryID}=" & Val(LV.SelectedItem.key), rpt, , , , Me
    End If
End Sub

Private Sub txtScanPath_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    If KeyCode.Value = 13 Then
        If txtCertification = "" Then
            MsgBox "Please Enter Certification.", vbInformation
            Exit Sub
        End If
        Dim ITM As ListItem
        Set ITM = LV.ListItems.add(, , txtCertification)
        ITM.Tag = "0"
        ITM.ListSubItems.add , , Format(DTCertification, "dd-MMM-yyyy")
        ITM.ListSubItems.add , , Format(DTExpiryCertification, "dd-MMM-yyyy")
        ITM.ListSubItems.add , , txtScanPath
    End If
    
End Sub

