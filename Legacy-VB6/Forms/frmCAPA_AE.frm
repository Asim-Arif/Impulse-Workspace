VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmCAPA_AE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CAPA"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11895
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
   ScaleHeight     =   9165
   ScaleWidth      =   11895
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   4650
      Left            =   120
      TabIndex        =   41
      Top             =   3870
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8202
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Step 1"
      TabPicture(0)   =   "frmCAPA_AE.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "label1(11)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "label1(10)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "label1(9)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "label1(16)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DTResult"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtResult"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtTeam"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtNonConfDescription"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "FraSystem"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "FraReason"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Step 2"
      TabPicture(1)   =   "frmCAPA_AE.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtPreventive"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtRemarks"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtCorrective"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "DTCorrectiveAction"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "DTPreventiveAction"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "label1(18)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "label1(17)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "label1(14)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "label1(13)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "label1(12)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Step 3"
      TabPicture(2)   =   "frmCAPA_AE.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txtFollowup"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "DTCorrectiveFollowup"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "DTPreventiveFollowup"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "DTCAPAClose"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "label1(21)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "label1(20)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "label1(19)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "label1(15)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      Begin VB.Frame Frame4 
         Caption         =   "Evaluation of Response :"
         Height          =   1845
         Left            =   -74880
         TabIndex        =   68
         Top             =   645
         Width           =   5610
         Begin VB.OptionButton OptEval 
            Caption         =   "Verification audit is required"
            Height          =   390
            Index           =   0
            Left            =   165
            TabIndex        =   71
            Tag             =   "-1"
            Top             =   270
            Width           =   2475
         End
         Begin VB.OptionButton OptEval 
            Caption         =   "Verification of submitted documentation or self declaration prior to closure"
            Height          =   390
            Index           =   1
            Left            =   165
            TabIndex        =   70
            Top             =   840
            Width           =   5085
         End
         Begin VB.OptionButton OptEval 
            Caption         =   "Proposal of corrective actioni acceptable. Verify effectiveneess."
            Height          =   390
            Index           =   2
            Left            =   165
            TabIndex        =   69
            Top             =   1425
            Width           =   5190
         End
      End
      Begin VB.TextBox txtFollowup 
         Appearance      =   0  'Flat
         Height          =   1755
         Left            =   -69195
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   67
         Top             =   735
         Width           =   5625
      End
      Begin VB.TextBox txtPreventive 
         Appearance      =   0  'Flat
         Height          =   1065
         Left            =   -69225
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   65
         Top             =   2445
         Width           =   5610
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         Height          =   1065
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   63
         Top             =   2445
         Width           =   5610
      End
      Begin VB.Frame Frame3 
         Caption         =   "Corrective Action Reviewed By MR:"
         Height          =   1500
         Left            =   -69195
         TabIndex        =   59
         Top             =   750
         Width           =   5610
         Begin VB.OptionButton OPTCAR 
            Caption         =   "Change in Follow-up Date required"
            Height          =   315
            Index           =   2
            Left            =   165
            TabIndex        =   62
            Top             =   1095
            Width           =   3615
         End
         Begin VB.OptionButton OPTCAR 
            Caption         =   "Change in Implementation Date required"
            Height          =   315
            Index           =   1
            Left            =   165
            TabIndex        =   61
            Top             =   690
            Width           =   3615
         End
         Begin VB.OptionButton OPTCAR 
            Caption         =   "Change in Corrective Action required"
            Height          =   315
            Index           =   0
            Left            =   165
            TabIndex        =   60
            Tag             =   "-1"
            Top             =   270
            Width           =   3615
         End
      End
      Begin VB.TextBox txtCorrective 
         Appearance      =   0  'Flat
         Height          =   1065
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   57
         Top             =   1170
         Width           =   5610
      End
      Begin VB.Frame FraReason 
         Caption         =   "Reason :"
         Height          =   900
         Left            =   105
         TabIndex        =   49
         Top             =   525
         Width           =   5805
         Begin VB.OptionButton OptReason 
            Caption         =   "Product"
            Height          =   300
            Index           =   0
            Left            =   165
            TabIndex        =   53
            Tag             =   "-1"
            Top             =   375
            Width           =   945
         End
         Begin VB.OptionButton OptReason 
            Caption         =   "Process"
            Height          =   300
            Index           =   1
            Left            =   1650
            TabIndex        =   52
            Top             =   375
            Width           =   945
         End
         Begin VB.OptionButton OptReason 
            Caption         =   "System"
            Height          =   300
            Index           =   2
            Left            =   3135
            TabIndex        =   51
            Top             =   375
            Width           =   945
         End
         Begin VB.OptionButton OptReason 
            Caption         =   "Service"
            Height          =   300
            Index           =   3
            Left            =   4620
            TabIndex        =   50
            Top             =   375
            Width           =   945
         End
      End
      Begin VB.Frame FraSystem 
         Caption         =   "System :"
         Height          =   1455
         Left            =   105
         TabIndex        =   48
         Top             =   1425
         Width           =   5805
         Begin VB.CheckBox chkSystem 
            Caption         =   "Medical Device Directive 93/42/EEC"
            Height          =   345
            Index           =   5
            Left            =   2655
            TabIndex        =   78
            Top             =   1035
            Width           =   3000
         End
         Begin VB.CheckBox chkSystem 
            Caption         =   "IT Security Management"
            Height          =   345
            Index           =   4
            Left            =   2655
            TabIndex        =   77
            Top             =   645
            Width           =   3000
         End
         Begin VB.CheckBox chkSystem 
            Caption         =   "Occup. Health && Safety Management"
            Height          =   345
            Index           =   3
            Left            =   2655
            TabIndex        =   76
            Top             =   255
            Width           =   3000
         End
         Begin VB.CheckBox chkSystem 
            Caption         =   "cGMP / CFR 21 820"
            Height          =   345
            Index           =   2
            Left            =   120
            TabIndex        =   75
            Top             =   1035
            Width           =   2205
         End
         Begin VB.CheckBox chkSystem 
            Caption         =   "Environment Management"
            Height          =   345
            Index           =   1
            Left            =   120
            TabIndex        =   74
            Top             =   645
            Width           =   2205
         End
         Begin VB.CheckBox chkSystem 
            Caption         =   "Quality Management"
            Height          =   345
            Index           =   0
            Left            =   120
            TabIndex        =   73
            Top             =   255
            Width           =   2205
         End
      End
      Begin VB.TextBox txtNonConfDescription 
         Appearance      =   0  'Flat
         Height          =   1065
         Left            =   5940
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   47
         Top             =   630
         Width           =   5610
      End
      Begin VB.Frame Frame2 
         Height          =   630
         Left            =   105
         TabIndex        =   44
         Top             =   2880
         Width           =   5805
         Begin VB.OptionButton OptAcceptance 
            Caption         =   "Acceptance for further Action"
            Height          =   270
            Index           =   0
            Left            =   135
            TabIndex        =   46
            Tag             =   "-1"
            Top             =   210
            Width           =   2580
         End
         Begin VB.OptionButton OptAcceptance 
            Caption         =   "Hold for further action"
            Height          =   270
            Index           =   1
            Left            =   2835
            TabIndex        =   45
            Top             =   210
            Width           =   2580
         End
      End
      Begin VB.TextBox txtTeam 
         Appearance      =   0  'Flat
         Height          =   1065
         Left            =   5940
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   43
         Top             =   1905
         Width           =   5610
      End
      Begin VB.TextBox txtResult 
         Appearance      =   0  'Flat
         Height          =   1065
         Left            =   5940
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   42
         Top             =   3255
         Width           =   5610
      End
      Begin MSComCtl2.DTPicker DTResult 
         Height          =   330
         Left            =   120
         TabIndex        =   79
         Tag             =   "DT"
         Top             =   3750
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   108068867
         CurrentDate     =   40698
      End
      Begin MSComCtl2.DTPicker DTCorrectiveAction 
         Height          =   330
         Left            =   -74880
         TabIndex        =   81
         Tag             =   "DT"
         Top             =   3945
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   108068867
         CurrentDate     =   40698
      End
      Begin MSComCtl2.DTPicker DTPreventiveAction 
         Height          =   330
         Left            =   -72885
         TabIndex        =   83
         Tag             =   "DT"
         Top             =   3945
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   108068867
         CurrentDate     =   40698
      End
      Begin MSComCtl2.DTPicker DTCorrectiveFollowup 
         Height          =   330
         Left            =   -74880
         TabIndex        =   85
         Tag             =   "DT"
         Top             =   2835
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   108068867
         CurrentDate     =   40698
      End
      Begin MSComCtl2.DTPicker DTPreventiveFollowup 
         Height          =   330
         Left            =   -72840
         TabIndex        =   87
         Tag             =   "DT"
         Top             =   2835
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   108068867
         CurrentDate     =   40698
      End
      Begin MSComCtl2.DTPicker DTCAPAClose 
         Height          =   330
         Left            =   -65595
         TabIndex        =   89
         Tag             =   "DT"
         Top             =   2835
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   108068867
         CurrentDate     =   40698
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "CAPA Close Date"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   21
         Left            =   -65595
         TabIndex        =   90
         Top             =   2625
         Width           =   1230
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Preventive Action Followup"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   20
         Left            =   -72825
         TabIndex        =   88
         Top             =   2625
         Width           =   1950
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Corrective Action Followup"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   19
         Left            =   -74880
         TabIndex        =   86
         Top             =   2625
         Width           =   1920
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Preventive Action Date"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   18
         Left            =   -72885
         TabIndex        =   84
         Top             =   3735
         Width           =   1665
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Corrective Action Date"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   17
         Left            =   -74880
         TabIndex        =   82
         Top             =   3735
         Width           =   1635
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Result Investigation Date"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   16
         Left            =   150
         TabIndex        =   80
         Top             =   3540
         Width           =   1830
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Follow up :"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   15
         Left            =   -69150
         TabIndex        =   72
         Top             =   480
         Width           =   780
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Preventive Action :"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   14
         Left            =   -69195
         TabIndex        =   66
         Top             =   2235
         Width           =   1380
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Remarks :"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   13
         Left            =   -74850
         TabIndex        =   64
         Top             =   2235
         Width           =   720
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Corrective Action :"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   12
         Left            =   -74865
         TabIndex        =   58
         Top             =   975
         Width           =   1350
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Non-Conformity Description :"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   9
         Left            =   5970
         TabIndex        =   56
         Top             =   420
         Width           =   2085
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Team Nominated :"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   10
         Left            =   5970
         TabIndex        =   55
         Top             =   1695
         Width           =   1305
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Result of Investigation :"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   11
         Left            =   5970
         TabIndex        =   54
         Top             =   3030
         Width           =   1740
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "NC Details :"
      Height          =   1560
      Left            =   1605
      TabIndex        =   24
      Top             =   1590
      Width           =   8655
      Begin MSComCtl2.DTPicker DTNCRONC 
         Height          =   330
         Left            =   2445
         TabIndex        =   25
         Tag             =   "DT"
         Top             =   435
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   108068867
         CurrentDate     =   40698
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "NC Type"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   0
         Left            =   3930
         TabIndex        =   36
         Top             =   225
         Width           =   615
      End
      Begin MSForms.ComboBox cmbNCTypeRONC 
         Height          =   315
         Left            =   3930
         TabIndex        =   35
         Tag             =   "NCType"
         Top             =   435
         Width           =   2430
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4286;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtNCRONC 
         Height          =   315
         Index           =   0
         Left            =   75
         TabIndex        =   34
         Tag             =   "NCNo"
         Top             =   435
         Width           =   2340
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "4128;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "NC No."
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   33
         Top             =   225
         Width           =   510
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   2
         Left            =   2475
         TabIndex        =   32
         Top             =   225
         Width           =   345
      End
      Begin MSForms.TextBox txtNCRONC 
         Height          =   315
         Index           =   1
         Left            =   6375
         TabIndex        =   31
         Tag             =   "Source"
         Top             =   435
         Width           =   2160
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "3810;556"
         Value           =   "Customer Complaint"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Source"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   3
         Left            =   6345
         TabIndex        =   30
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   29
         Top             =   855
         Width           =   855
      End
      Begin MSForms.TextBox txtNCRONC 
         Height          =   315
         Index           =   2
         Left            =   75
         TabIndex        =   28
         Tag             =   "Department"
         Top             =   1065
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
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   1
         Left            =   1965
         TabIndex        =   27
         Top             =   855
         Width           =   795
      End
      Begin MSForms.TextBox txtNCRONC 
         Height          =   315
         Index           =   3
         Left            =   1935
         TabIndex        =   26
         Tag             =   "Description"
         Top             =   1065
         Width           =   6615
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "11668;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame FraReadOnly 
      Caption         =   "Customer Complaint Details :"
      Enabled         =   0   'False
      Height          =   1500
      Left            =   1605
      TabIndex        =   3
      Top             =   75
      Width           =   8655
      Begin ComboList.Usercontrol1 cmbCustRO 
         Height          =   285
         Left            =   75
         TabIndex        =   4
         Tag             =   "CustCode"
         Top             =   450
         Width           =   2310
         _ExtentX        =   4075
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
      Begin MSComCtl2.DTPicker DTPicker2RO 
         Height          =   330
         Left            =   6840
         TabIndex        =   5
         Tag             =   "DT"
         Top             =   450
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   108068867
         CurrentDate     =   40698
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Complaint Date"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   5
         Left            =   6885
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin MSForms.TextBox TBoxRO 
         Height          =   315
         Index           =   14
         Left            =   4290
         TabIndex        =   22
         Tag             =   "ComplaintNo"
         Top             =   450
         Width           =   2520
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "4445;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Complaint No."
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   73
         Left            =   4275
         TabIndex        =   21
         Top             =   240
         Width           =   1005
      End
      Begin MSForms.ComboBox cmbCountryRO 
         Height          =   315
         Left            =   2415
         TabIndex        =   20
         Tag             =   "Country"
         Top             =   450
         Width           =   1845
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3254;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Customer"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   19
         Top             =   240
         Width           =   690
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Country"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   7
         Left            =   2430
         TabIndex        =   18
         Top             =   240
         Width           =   585
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Complaint By"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   17
         Top             =   810
         Width           =   930
      End
      Begin MSForms.TextBox TBoxRO 
         Height          =   315
         Index           =   2
         Left            =   75
         TabIndex        =   16
         Tag             =   "ComplaintBy"
         Top             =   1035
         Width           =   2310
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "4075;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   3
         Left            =   2430
         TabIndex        =   15
         Top             =   810
         Width           =   840
      End
      Begin MSForms.TextBox TBoxRO 
         Height          =   315
         Index           =   3
         Left            =   2415
         TabIndex        =   14
         Tag             =   "Designation"
         Top             =   1035
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
      Begin MSForms.ComboBox cmbPenaltyRO 
         Height          =   315
         Left            =   4290
         TabIndex        =   13
         Tag             =   "Penalty"
         Top             =   1035
         Width           =   1695
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2990;556"
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
         Caption         =   "Penalty"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   4
         Left            =   4275
         TabIndex        =   12
         Top             =   810
         Width           =   540
      End
      Begin MSForms.TextBox TBoxRO 
         Height          =   315
         Index           =   4
         Left            =   6015
         TabIndex        =   11
         Tag             =   "PenaltyQty"
         Top             =   1035
         Width           =   810
         VariousPropertyBits=   679495705
         BorderStyle     =   1
         Size            =   "1429;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.TextBox TBoxRO 
         Height          =   315
         Index           =   5
         Left            =   6840
         TabIndex        =   10
         Tag             =   "PenaltyAmount"
         Top             =   1035
         Width           =   810
         VariousPropertyBits=   679495705
         BorderStyle     =   1
         Size            =   "1429;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.TextBox TBoxRO 
         Height          =   315
         Index           =   6
         Left            =   7665
         TabIndex        =   9
         Tag             =   "PenaltyCurrency"
         Top             =   1035
         Width           =   810
         VariousPropertyBits=   679495705
         BorderStyle     =   1
         Size            =   "1429;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   5
         Left            =   6030
         TabIndex        =   8
         Top             =   810
         Width           =   270
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   6
         Left            =   6915
         TabIndex        =   7
         Top             =   810
         Width           =   555
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Currency"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   7
         Left            =   7695
         TabIndex        =   6
         Top             =   810
         Width           =   660
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   495
      Top             =   645
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "PDF Files Only|*.pdf"
   End
   Begin MSComCtl2.DTPicker DTCAPA 
      Height          =   330
      Left            =   2580
      TabIndex        =   39
      Tag             =   "DT"
      Top             =   3465
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   108068867
      CurrentDate     =   40698
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   8
      Left            =   2625
      TabIndex        =   40
      Top             =   3255
      Width           =   345
   End
   Begin MSForms.TextBox txtCAPA 
      Height          =   315
      Index           =   0
      Left            =   210
      TabIndex        =   38
      Tag             =   "NCNo"
      Top             =   3450
      Width           =   2340
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "4128;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "CAPA No."
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   37
      Top             =   3240
      Width           =   705
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   420
      Left            =   10200
      TabIndex        =   2
      Top             =   8655
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
   Begin MSForms.CommandButton cmdOk 
      Height          =   420
      Left            =   8565
      TabIndex        =   1
      Top             =   8655
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
      Left            =   6930
      TabIndex        =   0
      Top             =   8655
      Visible         =   0   'False
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
Attribute VB_Name = "frmCAPA_AE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lEntryID As Long
Dim lCCEntryID As Long
Dim lNCEntryID As Long

Public Sub EditItem(p_lEntryID As Long)

    On Error GoTo err
    lEntryID = p_lEntryID
    
    Call ShowData(lEntryID)
    
    Me.Show 1
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub FillCmbs()

    On Error GoTo 0
    
    cmbCustRO.ListHeight = 2500
    Call cmbCustRO.AddVals(con, "DISTINCT CustCode", "ForeignCustomers", "CustCode")
    
    
    
    
    With cmbPenaltyRO
        .AddItem "None"
        .AddItem "Free Placement"
        .AddItem "Credit Note"
    End With
    
    With cmbNCTypeRONC
        .AddItem "Major"
        .AddItem "Minor"
        .AddItem "Observation"
        .AddItem "Critical"
    End With
End Sub

Private Sub cmbCust_matched()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    SQL = "SELECT DISTINCT Country FROM ForeignCustomers "
    
    If cmbCustRO.ID <> "0" Then
        SQL = SQL & " WHERE CustCode='" & cmbCustRO.ID & "'"
    End If
    
    With rs
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        cmbCountryRO.Clear
        Do Until .EOF
            cmbCountryRO.AddItem ![Country] & ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    If cmbCountryRO.ListCount > 0 Then cmbCountryRO.ListIndex = 0
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub cmbPenalty_Change()
    Call cmbPenalty_Click
End Sub

Private Sub cmbPenalty_Click()

    TBoxRO(4).Enabled = False
    TBoxRO(5).Enabled = False
    TBoxRO(6).Enabled = False
    If cmbPenaltyRO.ListIndex = 1 Then
        TBoxRO(4).Enabled = True
    ElseIf cmbPenaltyRO.ListIndex = 2 Then
        TBoxRO(5).Enabled = True
        TBoxRO(6).Enabled = True
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
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

Private Sub DTCAPA_Change()
    If lEntryID = 0 Then
        txtCAPA(0) = GetNewCAPANo(DTCAPA.Value)
    End If
End Sub

Private Sub DTNC_Change()
    
End Sub

Private Sub Form_Load()
    
    Me.KeyPreview = True
    DTPicker2RO.Value = Date
    
    DTNCRONC.Value = Date
    DTCAPA.Value = Date
    
    
    DTCAPAClose.Value = Date
    DTCAPAClose.Value = Null
    
    DTResult.Value = Date
    DTResult.Value = Null
    
    DTCorrectiveAction.Value = Date
    DTCorrectiveAction.Value = Null
    
    DTPreventiveAction.Value = Date
    DTPreventiveAction.Value = Null
    
    DTCorrectiveFollowup.Value = Date
    DTCorrectiveFollowup.Value = Null
    
    DTPreventiveFollowup.Value = Date
    DTPreventiveFollowup.Value = Null
    
    Call DTCAPA_Change
    Call FillCmbs
        
    DTCAPAClose.Enabled = UserHasAccess("CloseCAPA")
   
End Sub

Private Function SaveItem() As Boolean

    On Error GoTo err
    
    Dim SQL As String
    
    Dim strResultDT As String, strCorrectiveDT As String, strPreventiveDT As String
    Dim strCorrectiveFollowUpDT As String, strPreventiveFollowUpDT As String, strCapaCloseDT As String
    
    If IsNull(DTResult.Value) Then
        strResultDT = "NULL"
    Else
        strResultDT = "'" & DTResult.Value & "'"
    End If
    
    If IsNull(DTCorrectiveAction.Value) Then
        strCorrectiveDT = "NULL"
    Else
        strCorrectiveDT = "'" & DTCorrectiveAction.Value & "'"
    End If
    
    If IsNull(DTPreventiveAction.Value) Then
        strPreventiveDT = "NULL"
    Else
        strPreventiveDT = "'" & DTPreventiveAction.Value & "'"
    End If
    
    If IsNull(DTCorrectiveFollowup.Value) Then
        strCorrectiveFollowUpDT = "NULL"
    Else
        strCorrectiveFollowUpDT = "'" & DTCorrectiveFollowup.Value & "'"
    End If
    
    If IsNull(DTPreventiveFollowup.Value) Then
        strPreventiveFollowUpDT = "NULL"
    Else
        strPreventiveFollowUpDT = "'" & DTPreventiveFollowup.Value & "'"
    End If
    
    If IsNull(DTCAPAClose.Value) Then
        strCapaCloseDT = "NULL"
    Else
        strCapaCloseDT = "'" & DTCAPAClose.Value & "'"
    End If
    
    Call StartTrans(con)
    
    Dim lmyEntryID As Long
    If lEntryID = 0 Then
        con.Execute "INSERT INTO CAPA(CAPANo,DT,Reason,Description,Acceptance,TeamNominated," & _
         "ResultOfInvestigation,CorrectiveAction,CAReview,Remarks,PreventiveActions,ResponseEvaluation,FollowUp," & _
         "System_QM,System_EM,System_cGMP,System_OHSM,System_ITSM,System_MDD,ResultDT,CorrectiveActionDT,PreventiveActionDT" & _
         ",CorrectiveActionFollowupDT,PreventiveActionFollowUpDT,CAPACloseDT)" & " VALUES('" & txtCAPA(0) & "','" & DTCAPA & _
         "'," & Val(OptReason(0).Tag) & ",'" & txtNonConfDescription & "'," & Val(OptAcceptance(0).Tag) & ",'" & txtTeam & _
         "','" & txtResult & "','" & txtCorrective & "'," & Val(OPTCAR(0).Tag) & ",'" & txtRemarks & "','" & txtPreventive & _
         "'," & Val(OptEval(0).Tag) & ",'" & txtFollowup & "'," & chkSystem(0).Value & "," & chkSystem(1).Value & "," & _
         chkSystem(2).Value & "," & chkSystem(3).Value & "," & chkSystem(4).Value & "," & chkSystem(5).Value & "," & strResultDT & _
         "," & strCorrectiveDT & "," & strPreventiveDT & "," & strCorrectiveFollowUpDT & "," & strPreventiveFollowUpDT & "," & strCapaCloseDT & ")"
    Else
        con.Execute "UPDATE CAPA SET Reason='" & Val(OptReason(0).Tag) & "',Description='" & _
         txtNonConfDescription & "',Acceptance=" & Val(OptAcceptance(0).Tag) & ",TeamNominated='" & txtTeam & _
         "',ResultOfInvestigation='" & txtResult & "',CorrectiveAction='" & txtCorrective & "',CAReview=" & _
         Val(OPTCAR(0).Tag) & ",Remarks='" & txtRemarks & "',PreventiveActions='" & txtPreventive & _
         "',ResponseEvaluation='" & Val(OptEval(0).Tag) & "',FollowUp='" & txtFollowup & "',System_QM=" & _
         chkSystem(0).Value & ",System_EM=" & chkSystem(1).Value & ",System_cGMP=" & chkSystem(2).Value & _
         ",System_OHSM=" & chkSystem(3).Value & ",System_ITSM=" & chkSystem(4).Value & ",System_MDD=" & _
         chkSystem(5).Value & ",ResultDT=" & strResultDT & ",CorrectiveActionDT=" & strCorrectiveDT & _
         ",PreventiveActionDT=" & strPreventiveDT & ",CorrectiveActionFollowupDT=" & strCorrectiveFollowUpDT & _
         ",PreventiveActionFollowUpDT=" & strPreventiveFollowUpDT & ",CAPACloseDT=" & strCapaCloseDT & " WHERE EntryID=" & lEntryID
    End If
    
    'con.Execute Sql
    
    If lEntryID = 0 Then
        lmyEntryID = GetSingleLongValue("MAX(EntryID)", "CustomerComplaints")
        con.Execute "UPDATE CustomerComplaints SET CAPANo='" & txtCAPA(0).Text & "' WHERE EntryID=" & lCCEntryID
    Else
        lmyEntryID = lEntryID
    End If
    
    con.CommitTrans
   
    lEntryID = 0
    SaveItem = True
    
    
    Exit Function
err:
    MsgBox err.Description
    
End Function



Private Function CreateInsertSQL(strTable As String) As String
    
    Dim c As Control
    Dim RetString As String
    
    'Create The First Part Of Insert Query...
    RetString = "INSERT INTO " & strTable & "("
    '---------------------------------------------------------------------------
    For Each c In Controls
        If Not (Right(c.Name, 2) = "RO" Or Right(c.Name, 2) = "RONC") Then
            If (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is MSForms.ComboBox) Or (TypeOf c Is MSForms.CheckBox) Or (TypeOf c Is ComboList.Usercontrol1) Or (TypeOf c Is DTPicker) Then
                If c.Tag <> "" Then
                    RetString = RetString & c.Tag & ","
                End If
            End If
        End If
    Next
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    RetString = RetString & ") Values("
    
    
    'Now  Get The Second Part...
    '---------------------------------------------------------------------------
    For Each c In Controls
        If Not (Right(c.Name, 2) = "RO" Or Right(c.Name, 2) = "RONC") Then
            If c.Tag <> "" Then
                If (TypeOf c Is DTPicker) Then
                    RetString = RetString & "'" & c.Value & "',"
                Else
                    RetString = RetString & "'" & c.Text & "',"
                End If
            End If
        End If
    Next
    
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
        If Not (Right(c.Name, 2) = "RO" Or Right(c.Name, 2) = "RONC") Then
            If c.Tag <> "" Then
                If (TypeOf c Is DTPicker) Then
                    RetString = RetString & c.Tag & "='" & c.Value & "',"
                Else
                    RetString = RetString & c.Tag & "='" & c.Text & "',"
                End If
            End If
        End If
    Next
    
    
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    
    CreateUpdateSQL = RetString
    
End Function



Private Sub ShowCCData(p_lEntryID As Long)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim c As Control
    With rs
        .Open "SELECT * FROM CustomerComplaints WHERE EntryID=" & p_lEntryID, con, adOpenForwardOnly, adLockReadOnly
        
        For Each c In Controls
            If Right(c.Name, 2) = "RO" Then
                If c.Tag <> "" Then
                    If TypeOf c Is MSForms.CheckBox Then
                        c.Value = .Fields(c.Tag)
                    ElseIf TypeOf c Is ComboList.Usercontrol1 Then
                        'C.Text = .Fields(C.Tag & "Text") & ""
                        c.ID = .Fields(c.Tag) & ""
                    ElseIf TypeOf c Is DTPicker Then
                        If Not IsNull(.Fields(c.Tag)) Then
                            c.Value = .Fields(c.Tag)
                        End If
                    ElseIf c.Name = "cmbPenaltyRO" Then
                        c.ListIndex = Val(.Fields(c.Tag) & "")
                    Else
                        c.Text = .Fields(c.Tag) & ""
                    End If
                End If
            End If
        Next
        txtNonConfDescription = !Reason & ""
        .Close
        
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub ShowNCData(p_lEntryID As Long)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim c As Control
    With rs
        .Open "SELECT * FROM NonCompliance WHERE EntryID=" & p_lEntryID, con, adOpenForwardOnly, adLockReadOnly
        
        For Each c In Controls
            If Right(c.Name, 4) = "RONC" Then
                If c.Tag <> "" Then
                    If TypeOf c Is MSForms.CheckBox Then
                        c.Value = .Fields(c.Tag)
                    ElseIf TypeOf c Is ComboList.Usercontrol1 Then
                        'C.Text = .Fields(C.Tag & "Text") & ""
                        c.ID = .Fields(c.Tag) & ""
                    ElseIf TypeOf c Is DTPicker Then
                        If Not IsNull(.Fields(c.Tag)) Then
                            c.Value = .Fields(c.Tag)
                        End If
                    ElseIf c.Name = "cmbPenaltyRO" Then
                        c.ListIndex = Val(.Fields(c.Tag) & "")
                    Else
                        c.Text = .Fields(c.Tag) & ""
                    End If
                End If
            End If
        Next
        
        
        .Close
        
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Public Sub AddNew(p_ComplaintNo As Long, p_NCEntryID As Long)

    lCCEntryID = p_ComplaintNo
    
    lNCEntryID = p_NCEntryID
    
    Call ShowCCData(lCCEntryID)
    Call ShowNCData(lNCEntryID)
    
    txtCAPA(0) = GetNewCAPANo(DTCAPA.Value)
    Me.Show 1
    
End Sub

Private Function GetNewCAPANo(myDT As Date) As String
    'CC-1/2011
    Dim lCAPANo As Long
    lCAPANo = GetSingleLongValue("MAX(CAST(SUBSTRING(CAPANo,6,CHARINDEX('/',CAPANo)-6) AS INT))", "CAPA")
    lCAPANo = lCAPANo + 1
    
    GetNewCAPANo = "CAPA-" & lCAPANo & "/" & year(myDT)
    
End Function

Private Sub ShowData(p_lEntryID As Long)

    On Error GoTo err
    
    
    Dim rs As New ADODB.Recordset
    Dim c As Control
    With rs
         
        .Open "SELECT * FROM CAPA WHERE EntryID=" & p_lEntryID, con, adOpenForwardOnly, adLockReadOnly
        
'        con.Execute "INSERT INTO CAPA(CAPANo,DT,Reason,Description,Acceptance,TeamNominated," & _
'         "ResultOfInvestigation,CorrectiveAction,CAReview,Remarks,PreventiveActions,ResponseEvaluation,FollowUp," & _
'         "System_QM,System_EM,System_cGMP,System_OHSM,System_ITSM,System_MDD)" & _
'         " VALUES('" & txtCAPA(0) & "','" & DTCAPA & "'," & Val(OptReason(0).Tag) & _
'         ",'" & txtNonConfDescription & "'," & Val(OptAcceptance(0).Tag) & ",'" & txtTeam & "','" & txtResult & _
'         "','" & txtCorrective & "'," & Val(OPTCAR(0).Tag) & ",'" & txtRemarks & "','" & txtPreventive & "'," & _
'         Val(OptEval(0).Tag) & ",'" & txtFollowup & "'," & chkSystem(0).Value & "," & chkSystem(1).Value & "," & _
'         chkSystem(2).Value & "," & chkSystem(3).Value & "," & chkSystem(4).Value & "," & chkSystem(5).Value & ")"
         
         
        txtCAPA(0) = !CapaNo & ""
        DTCAPA = !DT
        
        If Val(!Reason & "") > -1 Then
            OptReason(Val(!Reason & "")).Value = True
        End If
         
        txtNonConfDescription = !Description & ""
        
        If Val(!Acceptance & "") > -1 Then
            OptAcceptance(Val(!Acceptance & "")).Value = True
        End If
        
        txtTeam = !TeamNominated & ""
        txtResult = !ResultOfInvestigation & ""
        txtCorrective = !CorrectiveAction & ""
        
        If Val(!CAReview & "") > -1 Then
            OPTCAR(Val(!CAReview & "")).Value = True
        End If
        
        txtRemarks = !Remarks & ""
        txtPreventive = !PreventiveActions & ""
        
        If Val(!ResponseEvaluation & "") > -1 Then
            OptEval(Val(!ResponseEvaluation & "")).Value = True
        End If
        
        txtFollowup = !FollowUp & ""
        
        chkSystem(0).Value = Abs(!System_QM)
        chkSystem(1).Value = Abs(!System_EM)
        chkSystem(2).Value = Abs(!System_cGMP)
        chkSystem(3).Value = Abs(!System_OHSM)
        chkSystem(4).Value = Abs(!System_ITSM)
        chkSystem(5).Value = Abs(!System_MDD)
        
         
        DTResult.Value = !ResultDT & ""
        
        DTCorrectiveAction.Value = !CorrectiveActionDT & ""
        DTPreventiveAction.Value = !PreventiveActionDT & ""
        
        DTCorrectiveFollowup.Value = !CorrectiveActionFollowupDT & ""
        DTPreventiveFollowup.Value = !PreventiveActionFollowUpDT & ""
        DTCAPAClose.Value = !CAPACloseDT & ""
        
        .Close
        
        .Open "SELECT CustomerComplaints.EntryID AS lCCEntryID,NonCompliance.EntryID AS lNCEntryID FROM CustomerComplaints INNER JOIN NonCompliance ON CustomerComplaints.NCNo=NonCompliance.NCNo WHERE CAPANo='" & txtCAPA(0).Text & "'", con, adOpenForwardOnly, adLockReadOnly
        
        If .EOF = False Then
            Call ShowCCData(Val(!lCCEntryID & ""))
            Call ShowNCData(Val(!lNCEntryID & ""))
        End If
        
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub OptAcceptance_Click(Index As Integer)
    OptAcceptance(0).Tag = Index
End Sub

Private Sub OPTCAR_Click(Index As Integer)
    OPTCAR(0).Tag = Index
End Sub

Private Sub OptEval_Click(Index As Integer)
    OptEval(0).Tag = Index
End Sub

Private Sub OptReason_Click(Index As Integer)
    OptReason(0).Tag = Index
End Sub
