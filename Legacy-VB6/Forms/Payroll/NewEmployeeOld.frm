VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form NewEmployeeOld 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Employee's Info..."
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
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
   ScaleHeight     =   6675
   ScaleWidth      =   9480
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fCat 
      Caption         =   "Department's Info :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   555
      TabIndex        =   73
      Top             =   150
      Width           =   8070
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Department Name :"
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
         Index           =   0
         Left            =   3165
         TabIndex        =   77
         Top             =   330
         Width           =   1650
      End
      Begin MSForms.ComboBox cmbDept 
         Height          =   300
         Left            =   4860
         TabIndex        =   76
         Top             =   285
         Width           =   3090
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "5450;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbDeptID 
         Height          =   300
         Left            =   1500
         TabIndex        =   75
         Top             =   285
         Width           =   1575
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "2778;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Department ID :"
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
         Index           =   11
         Left            =   90
         TabIndex        =   74
         Top             =   345
         Width           =   1365
      End
   End
   Begin TabDlg.SSTab EmpTab 
      Height          =   4980
      Left            =   525
      TabIndex        =   4
      Top             =   1005
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   8784
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Employee Info"
      TabPicture(0)   =   "NewEmployeeOld.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FEmpInfo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Referential Info"
      TabPicture(1)   =   "NewEmployeeOld.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FRefInfo"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Previous History"
      TabPicture(2)   =   "NewEmployeeOld.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FHistory"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "General Info"
      TabPicture(3)   =   "NewEmployeeOld.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FGInfo"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame FRefInfo 
         Height          =   2370
         Left            =   1260
         TabIndex        =   53
         Top             =   495
         Width           =   5850
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Refered By (Name)"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   15
            Left            =   510
            TabIndex        =   65
            Top             =   315
            Width           =   1380
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   10
            Left            =   465
            TabIndex        =   64
            Top             =   540
            Width           =   1845
            VariousPropertyBits=   679495707
            Size            =   "3254;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   18
            Left            =   2340
            TabIndex        =   63
            Top             =   540
            Width           =   1845
            VariousPropertyBits=   679495707
            Size            =   "3254;556"
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
            Left            =   2385
            TabIndex        =   62
            Top             =   300
            Width           =   1035
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   19
            Left            =   465
            TabIndex        =   61
            Top             =   1170
            Width           =   1845
            VariousPropertyBits=   679495707
            Size            =   "3254;556"
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
            TabIndex        =   60
            Top             =   960
            Width           =   585
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   20
            Left            =   465
            TabIndex        =   59
            Top             =   1845
            Width           =   1845
            VariousPropertyBits=   679495707
            Size            =   "3254;556"
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
            TabIndex        =   58
            Top             =   1635
            Width           =   585
         End
         Begin MSForms.TextBox TBox 
            Height          =   990
            Index           =   21
            Left            =   2340
            TabIndex        =   57
            Top             =   1170
            Width           =   3045
            VariousPropertyBits=   -1467987941
            ScrollBars      =   2
            Size            =   "5371;1746"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Address"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   28
            Left            =   2370
            TabIndex        =   56
            Top             =   945
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Religion"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   30
            Left            =   4245
            TabIndex        =   55
            Top             =   315
            Width           =   555
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   23
            Left            =   4215
            TabIndex        =   54
            Top             =   540
            Width           =   1185
            VariousPropertyBits=   679495707
            Size            =   "2090;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame FEmpInfo 
         Height          =   4155
         Left            =   -74310
         TabIndex        =   26
         Top             =   495
         Width           =   6795
         Begin MSComCtl2.DTPicker DOB 
            Height          =   315
            Left            =   1710
            TabIndex        =   27
            Top             =   1350
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   22740995
            CurrentDate     =   36891
            MaxDate         =   36891
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   1
            Left            =   3945
            TabIndex        =   52
            Top             =   630
            Width           =   2205
            VariousPropertyBits=   679495707
            Size            =   "3889;556"
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
            Index           =   2
            Left            =   3990
            TabIndex        =   51
            Top             =   420
            Width           =   1035
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   0
            Left            =   1695
            TabIndex        =   50
            Top             =   630
            Width           =   2205
            VariousPropertyBits=   679495707
            Size            =   "3889;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Name"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   1740
            TabIndex        =   49
            Top             =   420
            Width           =   405
         End
         Begin VB.Label lblEmpID 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   600
            TabIndex        =   48
            Top             =   630
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Employee ID"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   615
            TabIndex        =   47
            Top             =   420
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Date Of Birth"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   1740
            TabIndex        =   46
            Top             =   1125
            Width           =   945
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   2
            Left            =   2715
            TabIndex        =   45
            Top             =   2085
            Width           =   1185
            VariousPropertyBits=   679495707
            Size            =   "2090;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Religion"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   2745
            TabIndex        =   44
            Top             =   1875
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Identity Mark"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   16
            Left            =   3990
            TabIndex        =   43
            Top             =   1125
            Width           =   960
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   11
            Left            =   3945
            TabIndex        =   42
            Top             =   1350
            Width           =   2205
            VariousPropertyBits=   679495707
            Size            =   "3889;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Address"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   13
            Left            =   2595
            TabIndex        =   41
            Top             =   2595
            Width           =   585
         End
         Begin MSForms.TextBox TBox 
            Height          =   990
            Index           =   8
            Left            =   2535
            TabIndex        =   40
            Top             =   2820
            Width           =   3600
            VariousPropertyBits=   -1467987941
            ScrollBars      =   2
            Size            =   "6350;1746"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "NIC #"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   615
            TabIndex        =   39
            Top             =   1875
            Width           =   435
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   4
            Left            =   600
            TabIndex        =   38
            Top             =   2085
            Width           =   2070
            VariousPropertyBits=   679495707
            Size            =   "3651;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbSex 
            Height          =   315
            Left            =   600
            TabIndex        =   37
            Top             =   1350
            Width           =   1065
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "1879;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sex"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   615
            TabIndex        =   36
            Top             =   1125
            Width           =   270
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   3
            Left            =   3210
            TabIndex        =   35
            Top             =   1350
            Width           =   690
            VariousPropertyBits=   679495711
            Size            =   "1217;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Age"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   3240
            TabIndex        =   34
            Top             =   1125
            Width           =   285
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Phone 2"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   12
            Left            =   615
            TabIndex        =   33
            Top             =   3285
            Width           =   585
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   7
            Left            =   600
            TabIndex        =   32
            Top             =   3495
            Width           =   1875
            VariousPropertyBits=   679495707
            Size            =   "3307;556"
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
            Index           =   10
            Left            =   615
            TabIndex        =   31
            Top             =   2610
            Width           =   585
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   6
            Left            =   600
            TabIndex        =   30
            Top             =   2820
            Width           =   1875
            VariousPropertyBits=   679495707
            Size            =   "3307;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "NTN"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   3990
            TabIndex        =   29
            Top             =   1875
            Width           =   300
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   5
            Left            =   3945
            TabIndex        =   28
            Top             =   2085
            Width           =   2205
            VariousPropertyBits=   679495707
            Size            =   "3889;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame FHistory 
         Height          =   1785
         Left            =   -74160
         TabIndex        =   19
         Top             =   495
         Width           =   6270
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Reason"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   19
            Left            =   510
            TabIndex        =   25
            Top             =   885
            Width           =   540
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   14
            Left            =   480
            TabIndex        =   24
            Top             =   1095
            Width           =   5355
            VariousPropertyBits=   679495707
            Size            =   "9446;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Last Designation"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   18
            Left            =   3210
            TabIndex        =   23
            Top             =   285
            Width           =   1185
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   13
            Left            =   3180
            TabIndex        =   22
            Top             =   480
            Width           =   2655
            VariousPropertyBits=   679495707
            Size            =   "4683;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Last Firm"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   17
            Left            =   525
            TabIndex        =   21
            Top             =   270
            Width           =   645
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   12
            Left            =   480
            TabIndex        =   20
            Top             =   480
            Width           =   2655
            VariousPropertyBits=   679495707
            Size            =   "4683;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame FGInfo 
         Height          =   3960
         Left            =   -74250
         TabIndex        =   5
         Top             =   495
         Width           =   6690
         Begin MSComCtl2.DTPicker JoinDate 
            Height          =   315
            Left            =   2415
            TabIndex        =   6
            Top             =   465
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "ddd  dd-MM-yyyy"
            Format          =   22740995
            CurrentDate     =   37383
         End
         Begin MSComCtl2.DTPicker RetireDate 
            Height          =   315
            Left            =   4305
            TabIndex        =   7
            Top             =   465
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "ddd  dd-MM-yyyy"
            Format          =   22740995
            CurrentDate     =   37383
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   24
            Left            =   2055
            TabIndex        =   71
            Top             =   1080
            Width           =   2040
            VariousPropertyBits=   679495707
            Size            =   "3598;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Socail Security No."
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   31
            Left            =   2085
            TabIndex        =   70
            Top             =   885
            Width           =   1335
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   22
            Left            =   525
            TabIndex        =   69
            Top             =   1080
            Width           =   1500
            VariousPropertyBits=   679495707
            Size            =   "2646;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Status"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   29
            Left            =   540
            TabIndex        =   68
            Top             =   870
            Width           =   465
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   9
            Left            =   525
            TabIndex        =   67
            Top             =   465
            Width           =   1860
            VariousPropertyBits=   679495707
            Size            =   "3281;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Designation"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   14
            Left            =   525
            TabIndex        =   66
            Top             =   255
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Prod. Profit"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   24
            Left            =   5160
            TabIndex        =   18
            Top             =   855
            Width           =   825
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   17
            Left            =   5160
            TabIndex        =   17
            Top             =   1080
            Width           =   990
            VariousPropertyBits=   679495707
            Size            =   "1746;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkOldAge 
            Height          =   345
            Left            =   4275
            TabIndex        =   16
            Top             =   3420
            Width           =   1905
            VariousPropertyBits=   1015023643
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3360;609"
            Value           =   "0"
            Caption         =   "Old Age Benifit Holder"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkSocial 
            Height          =   345
            Left            =   1650
            TabIndex        =   15
            Top             =   3420
            Width           =   2445
            VariousPropertyBits=   1015023643
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "4313;609"
            Value           =   "0"
            Caption         =   "Social Security Benifit Holder"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkTaxable 
            Height          =   345
            Left            =   525
            TabIndex        =   14
            Top             =   3420
            Width           =   915
            VariousPropertyBits=   1015023643
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "1614;609"
            Value           =   "0"
            Caption         =   "Taxable"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   1560
            Index           =   16
            Left            =   525
            TabIndex        =   13
            Top             =   1680
            Width           =   5595
            VariousPropertyBits=   -1467987941
            ScrollBars      =   2
            Size            =   "9869;2752"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Remarks"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   23
            Left            =   555
            TabIndex        =   12
            Top             =   1470
            Width           =   615
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   15
            Left            =   4125
            TabIndex        =   11
            Top             =   1080
            Width           =   990
            VariousPropertyBits=   679495707
            Size            =   "1746;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Salary"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   22
            Left            =   4125
            TabIndex        =   10
            Top             =   855
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Retirement Date"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   21
            Left            =   4335
            TabIndex        =   9
            Top             =   255
            Width           =   1185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Join Date"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   20
            Left            =   2505
            TabIndex        =   8
            Top             =   255
            Width           =   675
         End
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   4905
      TabIndex        =   72
      Top             =   990
      Visible         =   0   'False
      Width           =   4320
   End
   Begin MSForms.CheckBox chkSpecial 
      Height          =   345
      Left            =   540
      TabIndex        =   3
      Top             =   6090
      Visible         =   0   'False
      Width           =   1305
      VariousPropertyBits=   1015023643
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2302;609"
      Value           =   "0"
      Caption         =   "Special Case"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   405
      Left            =   7125
      TabIndex        =   2
      Top             =   6120
      Width           =   1620
      Caption         =   "Cancel        "
      PicturePosition =   327683
      Size            =   "2857;714"
      Picture         =   "NewEmployeeOld.frx":0070
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdNext 
      Height          =   405
      Left            =   3765
      TabIndex        =   0
      Top             =   6120
      Width           =   1620
      Caption         =   "Save & New"
      PicturePosition =   327683
      Size            =   "2857;714"
      Picture         =   "NewEmployeeOld.frx":0182
      Accelerator     =   78
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   405
      Left            =   5445
      TabIndex        =   1
      Top             =   6120
      Width           =   1620
      Caption         =   " Save & Close"
      PicturePosition =   327683
      Size            =   "2857;714"
      Picture         =   "NewEmployeeOld.frx":0294
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
Attribute VB_Name = "NewEmployeeOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Add As Boolean, Edit As Boolean
Dim Loaded As Boolean
Dim AccType As String
Dim AccountNo As String
Dim SubAccNo As String
Dim dontShowList As Boolean
Public SqlSettings As String

Private Sub chkParent_Click()

If chkParent = vbChecked Then
   TBox(1) = "0.00"
   TBox(1).Enabled = False
Else
   TBox(1).Enabled = True
End If
   
End Sub

Private Sub FillCmbs()
    cmbSex.AddItem "Male"
    cmbSex.AddItem "Female"
    cmbSex.ListIndex = 0
    Call AddToCombo(cmbDeptID, "deptID", "Departments", "Order By DeptID")
    Call AddToCombo(cmbDept, "Name", "Departments", " Order By DeptID")
    'Call AddToCombo(cmbSaleAcc, "AccTitle", "VActiveAccounts", " Where AccNo Like '31%' AND Not Parent Order By ACCTitle")
    'Call AddToCombo(cmbPchAcc, "AccTitle", "VActiveAccounts", " Where AccNo Like '41%' AND Not Parent Order By ACCTitle")
    'Call AddToCombo(cmbUnit, "Distinct Unit", "Items", "Order By Unit")
    
End Sub
Private Sub chkSpecial_Click()
    If Not Loaded Then Exit Sub
    If lblEmpID = "" Then chkSpecial.value = False
    If chkSpecial Then
        If SqlSettings = "" Then
            
            Load frmSettings
            frmSettings.lblEmpID = lblEmpID
            frmSettings.IsNew = True
            
            Dim rs As New ADODB.Recordset
            rs.Open "Select ID from Generalsettings Where ID='" & lblEmpID & "'", Con, adOpenForwardOnly, adLockReadOnly
            If rs.RecordCount > 0 Then
                frmSettings.Add = False
                Call frmSettings.Editrec
            Else
                frmSettings.Add = True
            End If
            rs.Close
            Set rs = Nothing
            
            frmSettings.Show 1
        Else
            frmSettings.Show 1
        End If
    End If
End Sub

Private Sub cmbdept_DropButtonClick()
    List1.Visible = False
End Sub

Private Sub cmbdept_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

If KeyCode = vbKeyDown Then
   If List1.ListIndex < List1.ListCount - 1 Then List1.ListIndex = List1.ListIndex + 1
   KeyCode = 0
   dontShowList = True
ElseIf KeyCode = vbKeyUp Then
   If List1.ListIndex > 0 Then List1.ListIndex = List1.ListIndex - 1
   KeyCode = 0
   dontShowList = True
ElseIf KeyCode = 13 And List1.Visible Then
   cmbDept = List1
   List1.Visible = False
   cmbDept.SetFocus
End If

End Sub

Private Sub cmbdept_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If dontShowList Then dontShowList = False Else ShowList List1, cmbDept, cmbDept, , fCat.Top + 110
End Sub

Private Sub cmbdept_LostFocus()
    List1.Visible = False
End Sub


Private Sub cmbdeptID_Validate(Cancel As Boolean)
    If Not cmbDeptID.MatchFound Then cmbDeptID.ListIndex = cmbDept.ListIndex
End Sub
Private Sub cmbdept_Change()
    If cmbDept.MatchFound Then Call cmbdept_Click 'Else If cmbdept <> "" Then cmbdept.DropDown
End Sub
Private Sub cmbdept_Click()
    If cmbDeptID.ListIndex <> cmbDept.ListIndex Then cmbDeptID.ListIndex = cmbDept.ListIndex
End Sub
Private Sub cmbdept_Validate(Cancel As Boolean)
   If Not cmbDept.MatchFound Then cmbDept.ListIndex = cmbDeptID.ListIndex
End Sub
Private Sub cmbsaleAcc_Validate(Cancel As Boolean)
    If Not cmbSaleAcc.MatchFound Then cmbSaleAcc = ""
End Sub

Private Sub cmbPchAcc_Validate(Cancel As Boolean)
    If Not cmbPchAcc.MatchFound Then cmbPchAcc = ""
End Sub

Private Sub cmbDeptID_Change()
    If cmbDeptID.MatchFound Then Call cmbdeptid_Click Else If Not cmbDeptID = "" Then cmbDeptID.DropDown
End Sub
Private Sub cmbdeptid_Click()
    If cmbDept.ListIndex <> cmbDeptID.ListIndex Then cmbDept.ListIndex = cmbDeptID.ListIndex
    If Add Then lblEmpID = GetNextEmpID(cmbDeptID)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    If SqlSettings <> "" Then
        Unload frmSettings
    End If
End Sub

Private Sub cmdNext_Click()
    If saveEmp Then
        Unload Me
        Me.SqlSettings = ""
        Me.Add = True
        Me.Edit = False
        Me.Show 1
    End If
End Sub

Private Sub cmdOK_Click()
    If saveEmp Then
        Unload Me
        If SqlSettings <> "" Then Unload frmSettings
    End If
End Sub
'Private Sub Form_Activate()
'   If Edit And Not Loaded Then Call showdata
'End Sub




Private Sub DOB_Change()
    Dim Months As Double
    Dim Years As Integer
    Dim Totdays As Integer
    Totdays = DateDiff("d", Format(DOB.value, "MMM-dd-yyyy"), Format(Date, "MMM-dd-yyyy"))
    Years = Totdays / 365
    'Months = Totdays Mod 365
    'Months = Round((Tot - Years) * 12)
    'TBox(3) = Round((Date - DOB.value) / (365))
    TBox(3) = Years
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
   Me.KeyPreview = True
'    DTPicker1.value = Date
    Call FillCmbs
    SqlSettings = ""
End Sub

Private Function saveEmp() As Boolean

On Error GoTo err

Dim OpBalance As Double
Dim rec As New ADODB.Recordset




If Not cmbDeptID.MatchFound Then
    MsgBox "Invalid Department Name.", vbInformation
    saveEmp = False
    Exit Function
ElseIf Trim(TBox(0)) = "" Then
    MsgBox "Invalid Employee Name.", vbInformation
    saveEmp = False
    Exit Function
'ElseIf Not cmbPchAcc.MatchFound Then
'    MsgBox "No Purchase Account Selected", vbInformation
'    saveEmp = False
'    Exit Function
'ElseIf Not cmbSaleAcc.MatchFound Then
'    MsgBox "No Sales Account Selected", vbInformation
'    saveEmp = False
'    Exit Function
'ElseIf Val(TBox(4)) = 0 Then
'    MsgBox "Invalid Purchase Price", vbInformation
'    saveEmp = False
'    Exit Function
'ElseIf Val(TBox(5)) = 0 Then
'    MsgBox "Invalid Sale Price", vbInformation
'    saveEmp = False
'    Exit Function
'ElseIf Not Val(TBox(4)) >= Val(TBox(5)) = 0 Then
'    If MsgBox("Sale Price Should Be More Than Purchase Price." & vbNewLine & "Proceed Anyway ?", vbInformation + vbYesNo) = vbNo Then
'        saveEmp = False
'        Exit Function
'    End If
End If


'pDim SaleAcc As String, PchAcc As String

'SaleAcc = getAccountNo(cmbSaleAcc)
'PchAcc = getAccountNo(cmbPchAcc)
Dim OldAge As Boolean, Social As Boolean, TaxPayee As Boolean
Dim Special As String
'If chkOldAge.value = vbChecked Then OldAge = True Else OldAge = False
'If chkSocial.value = vbChecked Then Social = True Else Social = False
'If chkTaxable.value = vbChecked Then TaxPayee = True Else TaxPayee = False
OldAge = chkOldAge.value
Social = chkSocial.value
TaxPayee = chkTaxable.value
If chkSpecial.value Then Special = lblEmpID Else Special = "GENERAL"
Con.BeginTrans
If Special <> "GENERAL" Then Con.Execute SqlSettings
If Add Then
    Con.Execute "Insert into Employees(DeptID,EmpID,Name,fname," & _
     "dob,religion,age,sex,address,NIC,NTN,Phone1,Phone2,designation" & _
     ",reference,IdentityMark,lastEmpFirm,lastEmpDesig,lastEmpReason" & _
     ",JoinDate,Retirement,Salary,PProfit,StartingSalary,oldage,Social," & _
     "taxpayee,remarks,GInfo,RefFName,RefPhone1,RefPhone2,RefReligion," & _
     "RefAddress,SocailNo,Status) Values('" & cmbDeptID & "','" & lblEmpID & _
     "','" & TBox(0) & "','" & TBox(1) & "',#" & DOB & "#,'" & TBox(2) & _
     "'," & Val(TBox(3)) & ",'" & cmbSex & "','" & TBox(8) & "','" & TBox(4) & _
     "','" & TBox(5) & "','" & TBox(6) & "','" & TBox(7) & "','" & _
     TBox(9) & "','" & TBox(10) & "','" & TBox(11) & "','" & TBox(12) & _
     "','" & TBox(13) & "','" & TBox(14) & "',#" & JoinDate & "#,#" & _
     RetireDate & "#," & TBox(15) & "," & TBox(17) & "," & _
     Val(TBox(15)) + Val(TBox(17)) & "," & OldAge & "," & Social & "," & _
     TaxPayee & ",'" & TBox(16) & "','" & Special & "','" & TBox(18) & _
     "','" & TBox(19) & "','" & TBox(20) & "','" & TBox(23) & "','" & _
     TBox(21) & "','" & TBox(24) & "','" & TBox(22) & "')"
Else
    Con.Execute "Update Employees Set Name='" & TBox(0) & "',fName='" & _
     TBox(1) & "',DOB=#" & DOB & "#,Religion='" & TBox(2) & "',age=" & _
     TBox(3) & ",Sex='" & cmbSex & "',Address='" & TBox(8) & "',NIC='" & _
     TBox(4) & "',NTN='" & TBox(5) & "',Phone1='" & TBox(6) & "',Phone2='" & _
     TBox(7) & "',designation='" & TBox(9) & "',Reference='" & TBox(10) & _
     "',IdentityMark='" & TBox(11) & "',lastEmpFirm='" & TBox(12) & _
     "',lastEmpDesig='" & TBox(13) & "',lastEmpReason='" & TBox(14) & _
     "',JoinDate=#" & JoinDate & "#,Retirement=#" & RetireDate & "#," & _
     "StartingSalary=" & Val(TBox(15)) + Val(TBox(17)) & ",PProfit=" & _
     TBox(17) & "," & "Salary=" & TBox(15) & ",OldAge=" & OldAge & _
     ",Social=" & Social & ",TaxPayee=" & TaxPayee & ",Remarks='" & _
     TBox(16) & "',GInfo='" & Special & ",RefFName='" & TBox(18) & _
     "',RefPhone1='" & TBox(19) & "',RefPhone2='" & TBox(20) & _
     ",RefReligion='" & TBox(23) & "',RefAddress='" & TBox(21) & _
     "',SocialNo='" & TBox(24) & "',Status='" & TBox(22) & _
     "' Where EmpID='" & Me.Tag & "'"
End If
Con.CommitTrans
'If Special = "GENERAL" Then Con.Execute "delete from GeneralSettings Where ID='" & lblEmpID & "'"
    
 saveEmp = True
Exit Function
err:
    MsgBox err.Description
End Function


Private Sub List1_Click()

    If UCase(Screen.ActiveControl.Name) <> "CMBDEPT" Then
        cmbDept = List1
        List1.Visible = False
        cmbDept.SetFocus
    End If

End Sub

Private Sub Tbox_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    Select Case Index
    Case 3
        KeyAscii = OnlyNumber(Val(KeyAscii))
    Case 15
        KeyAscii = OnlyNumber(Val(KeyAscii), True)
    End Select
End Sub



Private Function GetNextItemID(CatID As String) As String
Dim NextVal As String
 NextVal = GetMax("val(Right(ItemID,2))", "Items", "Where CatID ='" & CatID & "'")
 GetNextItemID = CatID & "-" & Format(Val(NextVal) + 1, "000")
End Function

Private Sub TBox_Validate(Index As Integer, Cancel As Boolean)
    'If Index = 4 Or Index = 5 Then
    '    TBox(Index) = Format(Val(TBox(Index)), "#0.00")
    'ElseIf Index = 2 Or Index = 3 Then
    '    TBox(Index) = Format(Val(TBox(Index)), "#0")
    'End If
End Sub


Public Sub Editrec()
    Loaded = False
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select * From Employees Where EmpID='" & Me.Tag & "'", Con, adOpenForwardOnly, adLockReadOnly
        lblEmpID = ![EmpID] & ""
        cmbDeptID = ![DeptID] & ""
        TBox(0) = ![Name] & ""
        TBox(1) = ![FName] & ""
        DOB = ![DOB] & ""
        TBox(2) = ![Religion] & ""
        TBox(3) = ![age] & ""
        cmbSex = ![Sex] & ""
        TBox(4) = ![NIC] & ""
        TBox(5) = ![NTN] & ""
        TBox(6) = ![Phone1] & ""
        TBox(7) = ![Phone2] & ""
        TBox(8) = ![Address] & ""
        TBox(9) = ![Designation] & ""
        TBox(10) = ![Reference] & ""
        TBox(11) = ![IdentityMark] & ""
        TBox(12) = ![lastEmpFirm] & ""
        TBox(13) = ![LastEmpDesig] & ""
        TBox(14) = ![LastEmpReason] & ""
        TBox(15) = ![Salary] & ""
        TBox(17) = ![PProfit] & ""
        TBox(16) = ![Remarks] & ""
        JoinDate = ![JoinDate] & ""
        RetireDate = ![Retirement] & ""
        chkOldAge.value = ![OldAge] & ""
        chkSocial.value = ![Social] & ""
        chkTaxable.value = ![TaxPayee] & ""
        
        If ![GInfo] = "GENERAL" Then
            chkSpecial.value = False
        Else
            chkSpecial.value = True
        End If
        .Close
    End With

    Set rs = Nothing

    fCat.Enabled = False
    Loaded = True
    
    'TBox(2).Enabled = False
    'Loaded = True
    
End Sub
