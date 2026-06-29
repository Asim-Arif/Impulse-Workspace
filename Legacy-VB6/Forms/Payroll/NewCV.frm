VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form NewCV 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Employee's Info..."
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8160
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
   ScaleHeight     =   6165
   ScaleWidth      =   8160
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   60
      TabIndex        =   51
      Top             =   0
      Width           =   7995
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Save CV Info For Jobs"
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
         Index           =   2
         Left            =   2085
         TabIndex        =   52
         Top             =   150
         Width           =   3450
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Save CV Info For Jobs"
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
         Index           =   9
         Left            =   2115
         TabIndex        =   53
         Top             =   135
         Width           =   3450
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   -225
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Please Select Employee's Picture"
      Filter          =   "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
   End
   Begin TabDlg.SSTab EmpTab 
      Height          =   4980
      Left            =   60
      TabIndex        =   8
      Top             =   690
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   8784
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Personal Info"
      TabPicture(0)   =   "NewCV.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "General Info"
      TabPicture(1)   =   "NewCV.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "FHistory"
      Tab(1).Control(2)=   "FRefInfo"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Qualification"
      TabPicture(2)   =   "NewCV.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "Previous Job History"
         Height          =   885
         Left            =   -74220
         TabIndex        =   54
         Top             =   510
         Width           =   6360
         Begin MSComCtl2.DTPicker DTApply 
            Height          =   315
            Left            =   4335
            TabIndex        =   62
            Top             =   420
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "ddd dd-MMM-yyyy"
            Format          =   70647811
            CurrentDate     =   37744
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Apply Date"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   11
            Left            =   4350
            TabIndex        =   57
            Top             =   225
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Job Description"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   495
            TabIndex        =   56
            Top             =   210
            Width           =   1095
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   2
            Left            =   480
            TabIndex        =   55
            Top             =   420
            Width           =   3855
            VariousPropertyBits=   679495707
            Size            =   "6800;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Write Personal Info"
         Height          =   4095
         Left            =   240
         TabIndex        =   32
         Top             =   600
         Width           =   7605
         Begin MSComCtl2.DTPicker DOB 
            Height          =   315
            Left            =   1710
            TabIndex        =   33
            Top             =   1410
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "ddd dd-MMM-yyyy"
            Format          =   70647811
            CurrentDate     =   36891
            MaxDate         =   36891
         End
         Begin VB.Label Label1 
            BackColor       =   &H009C9C9C&
            BackStyle       =   0  'Transparent
            Caption         =   "Permanent Address :                                                      Temporary Address :"
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   3
            Left            =   300
            TabIndex        =   61
            Top             =   2640
            Width           =   6960
         End
         Begin VB.Label Label1 
            BackColor       =   &H009C9C9C&
            BackStyle       =   0  'Transparent
            Caption         =   "NIC # :                                 Relegion :                   Phone1 :                     Phone2 :"
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   1
            Left            =   300
            TabIndex        =   60
            Top             =   1890
            Width           =   6990
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   5
            Left            =   300
            TabIndex        =   58
            Top             =   660
            Width           =   615
            VariousPropertyBits=   679495711
            Size            =   "1085;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbMaritalStatus 
            Height          =   315
            Left            =   4800
            TabIndex        =   34
            Top             =   1410
            Width           =   2475
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "4366;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CommandButton cmdCopy 
            Height          =   675
            Left            =   3600
            TabIndex        =   50
            Top             =   3060
            Width           =   540
            Caption         =   "Copy"
            PicturePosition =   65543
            Size            =   "952;1191"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.OptionButton SO 
            Height          =   345
            Index           =   0
            Left            =   3840
            TabIndex        =   49
            Top             =   360
            Width           =   675
            VariousPropertyBits=   1015023635
            BackColor       =   -2147483633
            ForeColor       =   8388608
            DisplayStyle    =   5
            Size            =   "1191;609"
            Value           =   "1"
            Caption         =   "S/O"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.OptionButton SO 
            Height          =   345
            Index           =   1
            Left            =   4545
            TabIndex        =   48
            Top             =   360
            Width           =   690
            VariousPropertyBits=   1015023635
            BackColor       =   -2147483633
            ForeColor       =   8388608
            DisplayStyle    =   5
            Size            =   "1217;609"
            Value           =   "0"
            Caption         =   "D/O"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.OptionButton SO 
            Height          =   345
            Index           =   2
            Left            =   5310
            TabIndex        =   47
            Top             =   360
            Width           =   735
            VariousPropertyBits=   1015023635
            BackColor       =   -2147483633
            ForeColor       =   8388608
            DisplayStyle    =   5
            Size            =   "1296;609"
            Value           =   "0"
            Caption         =   "W/O"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbReligion 
            Height          =   315
            Left            =   2295
            TabIndex        =   46
            Top             =   2145
            Width           =   1605
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "2831;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   990
            Index           =   8
            Left            =   300
            TabIndex        =   45
            Top             =   2895
            Width           =   3240
            VariousPropertyBits=   -1467987941
            ScrollBars      =   2
            Size            =   "5715;1746"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   1
            Left            =   3870
            TabIndex        =   44
            Tag             =   "S/O"
            Top             =   660
            Width           =   3375
            VariousPropertyBits=   679495707
            Size            =   "5953;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   0
            Left            =   930
            TabIndex        =   43
            Top             =   660
            Width           =   2925
            VariousPropertyBits=   679495707
            Size            =   "5159;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   11
            Left            =   7170
            TabIndex        =   42
            Top             =   1410
            Visible         =   0   'False
            Width           =   2205
            VariousPropertyBits=   679495707
            Size            =   "3889;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   990
            Index           =   23
            Left            =   4230
            TabIndex        =   41
            Top             =   2895
            Width           =   3030
            VariousPropertyBits=   -1467987941
            ScrollBars      =   2
            Size            =   "5345;1746"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   4
            Left            =   300
            TabIndex        =   40
            Top             =   2145
            Width           =   1965
            VariousPropertyBits=   679495707
            Size            =   "3466;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbSex 
            Height          =   315
            Left            =   300
            TabIndex        =   39
            Top             =   1410
            Width           =   1395
            VariousPropertyBits=   679495711
            DisplayStyle    =   3
            Size            =   "2461;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            BackColor       =   &H009C9C9C&
            BackStyle       =   0  'Transparent
            Caption         =   "Sex:                       Date Of Birth :                         Age:             Marital Status "
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   7
            Left            =   315
            TabIndex        =   38
            Top             =   1155
            Width           =   6930
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   3
            Left            =   3870
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   1410
            Width           =   900
            VariousPropertyBits=   679495711
            Size            =   "1587;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   7
            Left            =   5535
            TabIndex        =   36
            Top             =   2145
            Width           =   1740
            VariousPropertyBits=   679495707
            Size            =   "3069;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   6
            Left            =   3900
            TabIndex        =   35
            Top             =   2145
            Width           =   1620
            VariousPropertyBits=   679495707
            Size            =   "2857;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            BackColor       =   &H009C9C9C&
            BackStyle       =   0  'Transparent
            Caption         =   "ID :         Name :"
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   0
            Left            =   300
            TabIndex        =   59
            Top             =   420
            Width           =   6960
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Define Employee Duties Below"
         Height          =   4365
         Left            =   -74910
         TabIndex        =   22
         Top             =   420
         Width           =   7755
         Begin VB.TextBox txts 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   27
            Text            =   "1"
            Top             =   510
            Width           =   465
         End
         Begin VB.TextBox txts 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   1
            Left            =   570
            TabIndex        =   26
            Top             =   510
            Width           =   2295
         End
         Begin VB.TextBox txts 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   2
            Left            =   2880
            TabIndex        =   25
            Top             =   510
            Width           =   1755
         End
         Begin VB.TextBox txts 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   3
            Left            =   5910
            TabIndex        =   24
            Top             =   510
            Width           =   675
         End
         Begin VB.TextBox txts 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   4
            Left            =   6600
            TabIndex        =   23
            Top             =   510
            Width           =   675
         End
         Begin MSComCtl2.DTPicker DT 
            Height          =   315
            Left            =   4650
            TabIndex        =   28
            Top             =   510
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy"
            Format          =   70647811
            UpDown          =   -1  'True
            CurrentDate     =   37683
         End
         Begin MSComctlLib.ListView LVQ 
            Height          =   3345
            Left            =   90
            TabIndex        =   29
            Top             =   930
            Width           =   7590
            _ExtentX        =   13388
            _ExtentY        =   5900
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   15199215
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Sr No"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Diploma/Degree/Certificate"
               Object.Width           =   4128
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Institute"
               Object.Width           =   3120
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Passing Year"
               Object.Width           =   2194
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Grade"
               Object.Width           =   1270
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Division"
               Object.Width           =   1270
            EndProperty
         End
         Begin VB.Label Label1 
            Caption         =   "Sr No:  Diploma/Degree/Certificate :     Name Of Institute :       Passing Year :      Grade :   Division :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   30
            Left            =   90
            TabIndex        =   31
            Top             =   270
            Width           =   7305
         End
         Begin MSForms.CommandButton cmdPost 
            Height          =   315
            Left            =   7260
            TabIndex        =   30
            Top             =   510
            Width           =   420
            PicturePosition =   327683
            Size            =   "741;556"
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
      Begin VB.Frame FHistory 
         Caption         =   "Previous Job History"
         Height          =   1365
         Left            =   -74250
         TabIndex        =   18
         Top             =   3420
         Width           =   6420
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   12
            Left            =   480
            TabIndex        =   5
            Top             =   420
            Width           =   3135
            VariousPropertyBits=   679495707
            Size            =   "5530;556"
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
            Top             =   210
            Width           =   645
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   13
            Left            =   3630
            TabIndex        =   6
            Top             =   420
            Width           =   2595
            VariousPropertyBits=   679495707
            Size            =   "4577;556"
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
            Left            =   3570
            TabIndex        =   20
            Top             =   225
            Width           =   1185
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   14
            Left            =   480
            TabIndex        =   7
            Top             =   945
            Width           =   5745
            VariousPropertyBits=   679495707
            Size            =   "10134;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Reason"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   19
            Left            =   510
            TabIndex        =   19
            Top             =   735
            Width           =   540
         End
      End
      Begin VB.Frame FRefInfo 
         Caption         =   "Referal Information"
         Height          =   2010
         Left            =   -74220
         TabIndex        =   12
         Top             =   1395
         Width           =   6360
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Refered By (Name)"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   15
            Left            =   510
            TabIndex        =   17
            Top             =   315
            Width           =   1380
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   10
            Left            =   465
            TabIndex        =   0
            Top             =   540
            Width           =   3105
            VariousPropertyBits=   679495707
            Size            =   "5477;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   18
            Left            =   3600
            TabIndex        =   1
            Top             =   540
            Width           =   2625
            VariousPropertyBits=   679495707
            Size            =   "4630;556"
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
            Left            =   3600
            TabIndex        =   16
            Top             =   300
            Width           =   1035
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   19
            Left            =   465
            TabIndex        =   2
            Top             =   1080
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
            TabIndex        =   15
            Top             =   870
            Width           =   585
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   20
            Left            =   465
            TabIndex        =   3
            Top             =   1605
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
            TabIndex        =   14
            Top             =   1395
            Width           =   585
         End
         Begin MSForms.TextBox TBox 
            Height          =   840
            Index           =   21
            Left            =   2340
            TabIndex        =   4
            Top             =   1080
            Width           =   3885
            VariousPropertyBits=   -1467987941
            ScrollBars      =   2
            Size            =   "6853;1482"
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
            TabIndex        =   13
            Top             =   855
            Width           =   585
         End
      End
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   405
      Left            =   6315
      TabIndex        =   11
      Top             =   5700
      Width           =   1620
      Caption         =   "Cancel        "
      PicturePosition =   327683
      Size            =   "2857;714"
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
      Left            =   2955
      TabIndex        =   9
      Top             =   5700
      Width           =   1620
      Caption         =   "Save & New"
      PicturePosition =   327683
      Size            =   "2857;714"
      Picture         =   "NewCV.frx":0054
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
      Left            =   4635
      TabIndex        =   10
      Top             =   5700
      Width           =   1620
      Caption         =   " Save & Close"
      PicturePosition =   327683
      Size            =   "2857;714"
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Menu mnupopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete This Job"
      End
   End
   Begin VB.Menu mnupopup1 
      Caption         =   "popup1"
      Visible         =   0   'False
      Begin VB.Menu mnuDeleteQual 
         Caption         =   "Delete This Entry"
      End
   End
End
Attribute VB_Name = "NewCV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SelOption As Integer
Dim LastPoint As Integer
Public add As Boolean, edit As Boolean
Dim Loaded As Boolean
Dim AccType As String
Dim AccountNo As String
Dim SubAccNo As String
Dim dontShowList As Boolean
Public SqlSettings As String
Dim PicFileName As String
Dim ThumbFileName As String
Public TableName As String
Dim DeleteJob As Boolean

Private Sub FillCmbs()
    cmbSex.AddItem "Male"
    cmbSex.AddItem "Female"
    cmbSex.ListIndex = 0
    Dim DeptTable As String
    
    If TableName = "Employees" Then
        DeptTable = "Departments"
    Else
        DeptTable = "ContDepartments"
    End If
    
    'Call AddToCombo(cmbSaleAcc, "AccTitle", "VActiveAccounts", " Where AccNo Like '31%' AND Not Parent Order By ACCTitle")
    'Call AddToCombo(cmbPchAcc, "AccTitle", "VActiveAccounts", " Where AccNo Like '41%' AND Not Parent Order By ACCTitle")
    'Call AddToCombo(cmbUnit, "Distinct Unit", "Items", "Order By Unit")
    
End Sub


Private Sub cmbMaritalStatus_Click()
TBox(11) = cmbMaritalStatus
End Sub

Private Sub cmdCancel_Click()

    Unload Me
    If SqlSettings <> "" Then
        Unload frmSettings
    End If
    
    
End Sub

Private Sub cmdCopy_Click()
    TBox(23) = TBox(8).Text & ""
End Sub

Private Sub cmdNext_Click()
    If saveEmp Then
        Unload Me
        Me.SqlSettings = ""
        Me.add = True
        Me.edit = False
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









Private Sub cmdPost_Click()
Dim ITM As ListItem
With LVQ
    Set ITM = .ListItems.add(, , txts(0))
    ITM.ListSubItems.add , , txts(1)
    ITM.ListSubItems.add , , txts(2)
    ITM.ListSubItems.add , , DT.year
    ITM.ListSubItems.add , , txts(3)
    ITM.ListSubItems.add , , txts(4)
End With

For i = 0 To txts.Count - 1
    txts(i) = ""
Next

NextQual

End Sub

Private Sub RefreshQual(ID As String)
Dim Rec As New ADODB.Recordset, ITM As ListItem
With Rec
    .CursorLocation = adUseClient
    .Open "SELECT * FROM CVQualification WHERE ID='" & ID & "'", con, adOpenKeyset, adLockReadOnly
    If .RecordCount > 0 Then
        For i = 1 To .RecordCount
            Set ITM = LVQ.ListItems.add(, , !SrNo)
            ITM.ListSubItems.add , , !Diploma & ""
            ITM.ListSubItems.add , , !Institute & ""
            ITM.ListSubItems.add , , ![year] & ""
            ITM.ListSubItems.add , , !Grad & ""
            ITM.ListSubItems.add , , !Division & ""
            .MoveNext
        Next
    End If
End With
End Sub

Private Sub SaveQual(ID As String)

con.Execute "DELETE FROM CVQualification WHERE ID='" & ID & "'"

For i = 1 To LVQ.ListItems.Count
    With LVQ.ListItems(i)
        con.Execute "INSERT INTO CVQualification VALUES('" & ID _
            & "'," & .Text & ",'" & .SubItems(1) & "','" & .SubItems(2) & "'," & .SubItems(3) _
            & ",'" & .SubItems(4) & "','" & .SubItems(5) & "')"
    End With
Next

End Sub

Private Sub DOB_Change()
    Dim Months As Double
    Dim Years As Integer
    Dim TotDays As Integer
    TotDays = DateDiff("d", Format(DOB.value, "MMM-dd-yyyy"), Format(Date, "MMM-dd-yyyy"))
    Years = TotDays / 365
    'Months = Totdays Mod 365
    'Months = Round((Tot - Years) * 12)
    'TBox(3) = Round((GetServerDate - DOB.value) / (365))
    TBox(3) = Years
End Sub

Private Sub Form_Activate()

If edit = True Then
    Editrec
Else
    TBox(5) = GetNextPriority
End If

'DTApply.MaxDate = Date + 10
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
   Me.KeyPreview = True
    DOB.MaxDate = Date
    Call FillCmbs
    SqlSettings = ""
    cmbReligion.AddItem "Islam"
    cmbReligion.AddItem "Chirstian"
    cmbReligion.AddItem "Other"
    
    cmbMaritalStatus.AddItem "Married"
    cmbMaritalStatus.AddItem "Single"
    
    JoinDate = GetServerDate
    RetireDate = DateAdd("yyyy", 25, GetServerDate)
    EmpTab.Tab = 0
    
    Loaded = True
End Sub

Private Function saveEmp() As Boolean

'On Error GoTo ERR

Dim OpBalance As Double

If Trim(TBox(0)) = "" Then
    MsgBox "Invalid Employee Name.", vbInformation
    saveEmp = False
    Exit Function
End If

    If Val(con.Execute("Select count(*) From Employees where upper(name)='" & Trim(UCase(TBox(0))) & "' and empid<>'" & lblEmpID & "'").Fields(0) & "") > 0 Then
        MsgBox "An Employee With The Name """ & TBox(0) & """ Already Exists.", vbInformation
         Exit Function
    End If

Dim OldAge As Byte, Social As Byte, TaxPayee As Byte, Bonus As Byte

Dim Special As String


If edit = False Then
    
    con.Execute "Insert into CVs(ID,Name,Rel,fname" & _
     ",dob,religion,sex,address,TempAddress,NIC,Phone1,Phone2" & _
     ",lastEmpFirm,lastEmpDesig,lastEmpReason" & _
     ",refName,RefFName,RefPhone1,RefPhone2" & _
     ",RefAddress,Status,Job,AppDate) Values('" & TBox(5) & "','" & TBox(0) & "','" & TBox(1).Tag & _
     "','" & TBox(1) & "','" & DOB & "','" & cmbReligion & _
     "','" & cmbSex & "','" & TBox(8) & "','" & TBox(23) & "','" & TBox(4) & _
     "','" & TBox(6) & "','" & TBox(7) & "','" & TBox(12) & "','" & TBox(13) & _
     "','" & TBox(14) & "','" & TBox(10) & "','" & TBox(18) & _
     "','" & TBox(19) & "','" & TBox(20) & "','" & TBox(21) & "','" & cmbStatus _
     & "','" & TBox(2) & "','" & DTApply.value & "')"

Else
    con.Execute "Update CVs Set Name='" & TBox(0) & "',Rel='" & TBox(1).Tag & "',fName='" & _
     TBox(1) & "',DOB='" & DOB & "',Religion='" & cmbReligion & "',Sex='" & _
     cmbSex & "',Address='" & TBox(8) & "',NIC='" & TBox(4) & _
     "',Phone1='" & TBox(6) & "',Phone2='" & _
     TBox(7) & "',lastEmpFirm='" & TBox(12) & "',lastEmpDesig='" & TBox(13) & "',lastEmpReason='" & _
     TBox(14) & "',RefName='" & TBox(10) & "',RefFName='" & TBox(18) & _
     "',RefPhone1='" & TBox(19) & "',RefPhone2='" & TBox(20) & "',RefAddress= '" & TBox(21) & _
     "',Status='" & cmbMaritalStatus & "', Job='" & TBox(2) & "',TempAddress='" & TBox(23) & "',Appdate='" & DTApply.value & _
     "' Where ID='" & Me.Tag & "'"
End If

SaveQual (TBox(5))

saveEmp = True
 
Exit Function

err:
    MsgBox err.Description
End Function



Private Sub AddEmpPic(Optional ThumbExp As Boolean)
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

Private Sub Image1_DblClick()
    Call AddEmpPic
End Sub




Private Sub Image2_dblClick()
Call Label2_dblClick
End Sub

Private Sub Label2_dblClick()
    Call AddEmpPic(True)
End Sub

Private Sub lblPic_DblClick()
    AddEmpPic
End Sub




Private Sub Form_Unload(cancel As Integer)
    edit = False
End Sub

Private Sub LVQ_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If LVQ.ListItems.Count > 0 Then
    If Button = vbRightButton Then
        PopUpMenu mnupopup1
    End If
End If

End Sub

Private Sub mnuDeleteQual_Click()

Dim SrNo As Integer
SrNo = Val(LVQ.SelectedItem)
If SrNo <> 0 Then
If LVQ.ListItems.Count > 0 Then
    LVQ.ListItems.Remove (LVQ.SelectedItem.Index)
    For i = SrNo To LVQ.ListItems.Count
        LVQ.ListItems(i).Text = Val(LVQ.ListItems(i)) - 1
    Next
End If
End If
    NextQual

End Sub

Private Sub SO_Click(Index As Integer)

    On Error GoTo err
    'If Not Loaded Then Exit Sub
    Dim sSex As String
    If SO(0) Then
        TBox(1).Tag = "S/O"
        sSex = "Male"
        TBox(1).SetFocus
        
    ElseIf SO(1) Then
        TBox(1).Tag = "D/O"
        sSex = "Female"
        TBox(1).SetFocus
    ElseIf SO(2) Then
        TBox(1).Tag = "W/O"
        sSex = "Female"
        TBox(1).SetFocus
    Else
        TBox(1).Tag = ""
    End If
    cmbSex.Text = sSex
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub SO_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    On Error GoTo err
    If KeyAscii = 13 Then
        If SO(Index).value = False Then
            SO(Index).value = True
        Else
            TBox(1).SetFocus
        End If
    End If
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub TBox_GotFocus(Index As Integer)
    
    Exit Sub
    
    If Not Loaded Then Exit Sub
    Select Case Index
        Case 10, 18, 19, 20, 21
            EmpTab.Tab = 1
        Case 12 To 14
            EmpTab.Tab = 1
        Case 15, 16, 17, 9, 5, 24, 2, 22, 25
            EmpTab.Tab = 2
        Case Else
            EmpTab.Tab = 0
    End Select
    
End Sub

Private Sub Tbox_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    Select Case Index
    Case 3
        KeyAscii = OnlyNumber(Val(KeyAscii))
    Case 15
        KeyAscii = OnlyNumber(Val(KeyAscii), True)
    End Select
End Sub



Private Function GetNextPriority() As String

Dim NextVal As String
 NextVal = GetMax("ID", "CVs")
 GetNextPriority = Format(NextVal + 1, "000")
 
End Function


Public Sub Editrec()
    'On Error GoTo ERR
    
    Loaded = False
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select * From CVs Where ID='" & Me.Tag & "'", con, adOpenForwardOnly, adLockReadOnly
        TBox(5) = ![ID] & ""
        TBox(0) = ![Name] & ""
        
        If UCase$(![Rel]) = "S/O" Then
            SO(0).value = True
            TBox(1).Tag = "S/O"
        ElseIf UCase$(![Rel]) = "D/O" Then
            SO(1).value = True
            TBox(1).Tag = "D/O"
        Else
            SO(2).value = True
            TBox(1).Tag = "W/O"
        End If
        
        TBox(1) = ![FName] & ""
        DOB = ![DOB] & ""
        Call DOB_Change
        cmbReligion = ![Religion] & ""
        'TBox(3) = ![age] & ""
        cmbSex = ![Sex] & ""
        TBox(4) = ![NIC] & ""
        TBox(6) = ![Phone1] & ""
        TBox(7) = ![Phone2] & ""
        TBox(8) = ![Address] & ""
        TBox(10) = ![RefName] & ""
        TBox(18) = ![RefFName] & ""
        TBox(19) = ![RefPhone1] & ""
        TBox(20) = ![RefPhone2] & ""
        TBox(23) = ![TempAddress] & ""
        TBox(21) = ![RefAddress] & ""
        TBox(12) = ![lastEmpFirm] & ""
        TBox(13) = ![LastEmpDesig] & ""
        TBox(14) = ![LastEmpReason] & ""
        cmbMaritalStatus = ![Status] & ""
        TBox(2) = !Job & ""
        DTApply.value = Format(!AppDate, "dd-mm-yyyy")
        .Close
    End With

    Set rs = Nothing

    If TBox(11) = "" Then
      cmbMaritalStatus.ListIndex = -1
    Else
      cmbMaritalStatus = TBox(11)
    End If
      RefreshQual (TBox(5))
    
    NextQual
    Exit Sub
err:
    MsgBox err.Description
    Resume Next
End Sub

Private Sub NextQual()

If LVQ.ListItems.Count > 0 Then
    txts(0) = Val(LVQ.ListItems(LVQ.ListItems.Count).Text) + 1
Else
    txts(0) = "1"
End If

End Sub
