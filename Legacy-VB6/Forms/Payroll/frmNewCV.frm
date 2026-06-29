VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "combolist.ocx"
Begin VB.Form frmNewCV 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CVs"
   ClientHeight    =   7410
   ClientLeft      =   -18000
   ClientTop       =   1395
   ClientWidth     =   10230
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
   ScaleHeight     =   7410
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicSearch 
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   2700
      ScaleHeight     =   600
      ScaleWidth      =   7125
      TabIndex        =   16
      Top             =   6735
      Visible         =   0   'False
      Width           =   7125
      Begin VB.ComboBox cmbField 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmNewCV.frx":0000
         Left            =   30
         List            =   "frmNewCV.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   210
         Width           =   2835
      End
      Begin VB.TextBox txtValue 
         Height          =   315
         Left            =   2880
         TabIndex        =   17
         Top             =   210
         Width           =   2745
      End
      Begin VB.Label lblSearch 
         BackStyle       =   0  'Transparent
         Caption         =   "Search By :                                           Search For:"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   45
         TabIndex        =   20
         Top             =   0
         Width           =   5235
      End
      Begin MSForms.CommandButton cmdSearch 
         Height          =   360
         Left            =   5670
         TabIndex        =   19
         Top             =   180
         Width           =   1380
         Caption         =   "Search      "
         PicturePosition =   327683
         Size            =   "2434;635"
         Accelerator     =   83
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   45
      TabIndex        =   21
      Top             =   1005
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Basic Info"
      TabPicture(0)   =   "frmNewCV.frx":0062
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FOptions(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Qualifications / Computer Literacy"
      TabPicture(1)   =   "frmNewCV.frx":007E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FOptions(2)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Experience / Reference"
      TabPicture(2)   =   "frmNewCV.frx":009A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FOptions(3)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Remarks"
      TabPicture(3)   =   "frmNewCV.frx":00B6
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label1(10)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label1(11)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label1(12)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "lblDesig"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "DTPicker1"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "OptSelection(0)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "OptSelection(1)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "OptSelection(2)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "txtInterviewBy"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "txtInterviewRemarks"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "LVDesig"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "cmbDesig"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).ControlCount=   12
      Begin ComboList.Usercontrol1 cmbDesig 
         Height          =   285
         Left            =   5610
         TabIndex        =   73
         Top             =   630
         Visible         =   0   'False
         Width           =   4065
         _ExtentX        =   7170
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
         TextBackColor   =   16777215
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
      Begin MSComctlLib.ListView LVDesig 
         Height          =   2040
         Left            =   5610
         TabIndex        =   72
         Top             =   900
         Visible         =   0   'False
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   3598
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Designation"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.TextBox txtInterviewRemarks 
         Height          =   1890
         Left            =   255
         MultiLine       =   -1  'True
         TabIndex        =   70
         Tag             =   "Remarks"
         Top             =   3570
         Width           =   9435
      End
      Begin VB.TextBox txtInterviewBy 
         Height          =   315
         Left            =   2415
         TabIndex        =   68
         Tag             =   "InterviewBy"
         Top             =   2985
         Width           =   7275
      End
      Begin VB.OptionButton OptSelection 
         Caption         =   "Future Short List"
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
         Height          =   375
         Index           =   2
         Left            =   3795
         TabIndex        =   65
         Top             =   615
         Width           =   1740
      End
      Begin VB.OptionButton OptSelection 
         Caption         =   "Short List"
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
         Height          =   375
         Index           =   1
         Left            =   2025
         TabIndex        =   64
         Top             =   615
         Width           =   1740
      End
      Begin VB.OptionButton OptSelection 
         Caption         =   "Reject"
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
         Height          =   375
         Index           =   0
         Left            =   255
         TabIndex        =   63
         Top             =   615
         Width           =   1740
      End
      Begin VB.Frame FOptions 
         Height          =   5115
         Index           =   2
         Left            =   -74715
         TabIndex        =   46
         Top             =   390
         Width           =   9150
         Begin VB.Frame Frame3 
            Caption         =   "Academic / Professional Qualification :"
            Height          =   2460
            Left            =   75
            TabIndex        =   50
            Top             =   120
            Width           =   9000
            Begin VB.CommandButton cmdAddAcaQual 
               Caption         =   "Add"
               Height          =   255
               Left            =   7995
               TabIndex        =   51
               Top             =   150
               Width           =   930
            End
            Begin MSComctlLib.ListView LVQ 
               Height          =   1725
               Left            =   75
               TabIndex        =   52
               Top             =   465
               Width           =   8850
               _ExtentX        =   15610
               _ExtentY        =   3043
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
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
               NumItems        =   7
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Sr No"
                  Object.Width           =   1058
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Diploma/Degree/Certificate"
                  Object.Width           =   5997
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Passing Year"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Name of Institute"
                  Object.Width           =   5997
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Marks Obtained"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Total Marks"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "Percentage"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Computer Literacy :"
            Height          =   2460
            Left            =   75
            TabIndex        =   47
            Top             =   2580
            Width           =   9000
            Begin VB.CommandButton cmdAddComputerLiteracy 
               Caption         =   "Add"
               Height          =   255
               Left            =   7980
               TabIndex        =   48
               Top             =   240
               Width           =   930
            End
            Begin MSComctlLib.ListView LVComputer 
               Height          =   1725
               Left            =   75
               TabIndex        =   49
               Top             =   570
               Width           =   8850
               _ExtentX        =   15610
               _ExtentY        =   3043
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
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
               NumItems        =   7
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Sr No"
                  Object.Width           =   1058
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Course Title"
                  Object.Width           =   5997
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Name of Institute"
                  Object.Width           =   5997
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Period"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Marks"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Tot. Marks"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "Percentage"
                  Object.Width           =   2117
               EndProperty
            End
         End
      End
      Begin VB.Frame FOptions 
         Height          =   5115
         Index           =   3
         Left            =   -74640
         TabIndex        =   39
         Top             =   375
         Width           =   9150
         Begin VB.Frame Frame14 
            Caption         =   "References :"
            Height          =   2460
            Left            =   75
            TabIndex        =   43
            Top             =   2580
            Width           =   9000
            Begin VB.CommandButton cmdAddRef 
               Caption         =   "Add"
               Height          =   255
               Left            =   7980
               TabIndex        =   44
               Top             =   240
               Width           =   930
            End
            Begin MSComctlLib.ListView LVRef 
               Height          =   1725
               Left            =   75
               TabIndex        =   45
               Top             =   570
               Width           =   8850
               _ExtentX        =   15610
               _ExtentY        =   3043
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
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
               NumItems        =   5
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Sr No"
                  Object.Width           =   1058
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Name"
                  Object.Width           =   2822
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Relation"
                  Object.Width           =   2469
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Organization"
                  Object.Width           =   3528
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Contact No."
                  Object.Width           =   2469
               EndProperty
            End
         End
         Begin VB.Frame Frame15 
            Caption         =   "Experience :"
            Height          =   2460
            Left            =   75
            TabIndex        =   40
            Top             =   120
            Width           =   9000
            Begin VB.CommandButton cmdAddExp 
               Caption         =   "Add"
               Height          =   255
               Left            =   7995
               TabIndex        =   41
               Top             =   150
               Width           =   930
            End
            Begin MSComctlLib.ListView LVExp 
               Height          =   1725
               Left            =   75
               TabIndex        =   42
               Top             =   465
               Width           =   8850
               _ExtentX        =   15610
               _ExtentY        =   3043
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
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
               NumItems        =   6
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Sr No"
                  Object.Width           =   1058
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Designation"
                  Object.Width           =   2822
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Name of Institute"
                  Object.Width           =   4233
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Period"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   4
                  Text            =   "Salary"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Why Left"
                  Object.Width           =   2469
               EndProperty
            End
         End
      End
      Begin VB.Frame FOptions 
         Height          =   4635
         Index           =   0
         Left            =   -74640
         TabIndex        =   22
         Top             =   600
         Width           =   9150
         Begin VB.Frame Frame7 
            Caption         =   "PERMANENT ADDRESS :"
            Height          =   2415
            Left            =   4680
            TabIndex        =   26
            Top             =   2055
            Width           =   4095
            Begin MSForms.TextBox TBox 
               Height          =   2070
               Index           =   40
               Left            =   60
               TabIndex        =   27
               Tag             =   "PerAddress"
               Top             =   240
               Width           =   3975
               VariousPropertyBits=   -1467987941
               ScrollBars      =   2
               Size            =   "7011;3651"
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin VB.CommandButton cmdCopyAddress 
            Caption         =   ">>"
            Height          =   285
            Left            =   4230
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   3180
            Width           =   375
         End
         Begin VB.Frame Frame6 
            Caption         =   "CURRENT ADDRESS :"
            Height          =   2415
            Left            =   75
            TabIndex        =   23
            Top             =   2055
            Width           =   4095
            Begin MSForms.TextBox TBox 
               Height          =   2085
               Index           =   8
               Left            =   60
               TabIndex        =   24
               Tag             =   "Address"
               Top             =   240
               Width           =   3975
               VariousPropertyBits=   -1467987941
               ScrollBars      =   2
               Size            =   "7011;3678"
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin MSComCtl2.DTPicker DOB 
            Height          =   315
            Left            =   75
            TabIndex        =   28
            Tag             =   "dob"
            Top             =   420
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "ddd dd-MMM-yyyy"
            Format          =   94437379
            CurrentDate     =   36891
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Area :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   4440
            TabIndex        =   61
            Top             =   1425
            Width           =   450
         End
         Begin MSForms.ComboBox cmbArea 
            Height          =   315
            Left            =   4425
            TabIndex        =   60
            Tag             =   "Area"
            Top             =   1635
            Width           =   2145
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "3784;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   6
            Left            =   2250
            TabIndex        =   59
            Tag             =   "Cell"
            Top             =   1635
            Width           =   2145
            VariousPropertyBits=   679495707
            Size            =   "3784;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cell :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   2265
            TabIndex        =   58
            Top             =   1425
            Width           =   360
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   5
            Left            =   75
            TabIndex        =   57
            Tag             =   "Phone1"
            Top             =   1635
            Width           =   2145
            VariousPropertyBits=   679495707
            Size            =   "3784;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Phone :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   56
            Top             =   1425
            Width           =   555
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   4
            Left            =   75
            TabIndex        =   55
            Tag             =   "NIC"
            Top             =   1035
            Width           =   2145
            VariousPropertyBits=   679495707
            Size            =   "3784;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Picture"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   32
            Left            =   7320
            TabIndex        =   54
            Top             =   180
            Width           =   495
         End
         Begin VB.Label lblPic 
            Alignment       =   2  'Center
            Caption         =   "Double Click To Add Picture."
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   7545
            TabIndex        =   53
            Top             =   1035
            Width           =   1095
         End
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   1590
            Left            =   7305
            Stretch         =   -1  'True
            Top             =   390
            Width           =   1500
         End
         Begin MSForms.ComboBox cmbMaritalStatus 
            Height          =   315
            Left            =   4425
            TabIndex        =   38
            Tag             =   "MaritalStatus"
            Top             =   1035
            Width           =   2145
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "3784;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbReligion 
            Height          =   315
            Left            =   4425
            TabIndex        =   37
            Tag             =   "Religion"
            Top             =   420
            Width           =   2145
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "3784;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Date Of Birth"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   135
            TabIndex        =   36
            Top             =   195
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Religion"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   4485
            TabIndex        =   35
            Top             =   210
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Marital Status "
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   16
            Left            =   4425
            TabIndex        =   34
            Top             =   795
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N.I.C. # : (Nadra Format)"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   33
            Top             =   795
            Width           =   1875
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sex"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   2310
            TabIndex        =   32
            Top             =   795
            Width           =   270
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   3
            Left            =   2250
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   420
            Width           =   2145
            VariousPropertyBits=   679495711
            Size            =   "3784;556"
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
            Left            =   2310
            TabIndex        =   30
            Top             =   195
            Width           =   285
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   15
            Left            =   2250
            TabIndex        =   29
            TabStop         =   0   'False
            Tag             =   "Sex"
            Top             =   1035
            Width           =   2145
            VariousPropertyBits=   679495711
            Size            =   "3784;556"
            Value           =   "Male"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   255
         TabIndex        =   62
         Tag             =   "InterviewDT"
         Top             =   2985
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "ddd dd-MMM-yyyy"
         Format          =   94437379
         CurrentDate     =   36891
      End
      Begin VB.Label lblDesig 
         AutoSize        =   -1  'True
         Caption         =   "Designations"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5610
         TabIndex        =   71
         Top             =   420
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Remarks"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   12
         Left            =   255
         TabIndex        =   69
         Top             =   3345
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Interview By"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   11
         Left            =   2460
         TabIndex        =   67
         Top             =   2775
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Interview Date"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   10
         Left            =   255
         TabIndex        =   66
         Top             =   2775
         Width           =   1080
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4305
      Top             =   6135
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Please Select Employee's Picture"
      Filter          =   "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
   End
   Begin VB.Frame fraNav 
      Caption         =   " Showing :   0 / 0 "
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
      Height          =   540
      Left            =   135
      TabIndex        =   11
      Top             =   6750
      Visible         =   0   'False
      Width           =   2505
      Begin VB.CommandButton cmdNav 
         Caption         =   "|<"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   200
         Width           =   495
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   705
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   200
         Width           =   495
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   1260
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   200
         Width           =   495
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   ">|"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   1815
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   200
         Width           =   495
      End
   End
   Begin MSForms.CommandButton cmdSave 
      Height          =   405
      Left            =   6675
      TabIndex        =   10
      Top             =   6825
      Width           =   1620
      Caption         =   "Save"
      PicturePosition =   327683
      Size            =   "2857;714"
      Accelerator     =   118
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox TBox 
      Height          =   390
      Index           =   2
      Left            =   5400
      TabIndex        =   9
      Tag             =   "Rel"
      Top             =   450
      Visible         =   0   'False
      Width           =   4485
      VariousPropertyBits=   679495707
      Size            =   "7911;688"
      Value           =   "S/O"
      FontName        =   "Tahoma"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.OptionButton SO 
      Height          =   345
      Index           =   0
      Left            =   5370
      TabIndex        =   8
      Top             =   120
      Width           =   630
      VariousPropertyBits=   1015023635
      BackColor       =   -2147483633
      ForeColor       =   8388608
      DisplayStyle    =   5
      Size            =   "1111;609"
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
      Left            =   5925
      TabIndex        =   7
      Top             =   120
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
      Left            =   6525
      TabIndex        =   6
      Top             =   120
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
   Begin MSForms.TextBox TBox 
      Height          =   390
      Index           =   1
      Left            =   5400
      TabIndex        =   1
      Tag             =   "FName"
      Top             =   450
      Width           =   4485
      VariousPropertyBits=   679495707
      Size            =   "7911;688"
      FontName        =   "Tahoma"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TBox 
      Height          =   390
      Index           =   0
      Left            =   1365
      TabIndex        =   0
      Tag             =   "name"
      Top             =   450
      Width           =   4020
      VariousPropertyBits=   679495707
      Size            =   "7091;688"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name :"
      ForeColor       =   &H00800000&
      Height          =   345
      Index           =   1
      Left            =   1380
      TabIndex        =   5
      Top             =   225
      Width           =   510
   End
   Begin VB.Label lblEmpID 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   390
      Left            =   60
      TabIndex        =   2
      Tag             =   "CVID"
      Top             =   450
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CV # :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   4
      Left            =   60
      TabIndex        =   4
      Top             =   225
      Width           =   465
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   405
      Left            =   8370
      TabIndex        =   3
      Top             =   6825
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
Attribute VB_Name = "frmNewCV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SelOption As Integer
Dim LastPoint As Integer

Dim Loaded As Boolean
Dim AccType As String
Dim AccountNo As String
Dim SubAccNo As String
Dim dontShowList As Boolean

Public SqlSettings As String

Dim PicFileName As String
Dim ThumbFileName As String
Dim SigFileName As String

Public TableName As String
Dim EditEmp As Boolean

Dim EmpId As String
Dim WithEvents rsView As ADODB.Recordset
Attribute rsView.VB_VarHelpID = -1
Dim strName As String, Relation As String, NICNo As String, Birth As Date, strAge As String
Dim MaritalStatus As String, Education As String, ClassName As String, Institute As String, Occupation As String

Dim myVal As Integer
Dim LvChildIndex As Integer

Public Sub ShowMe(strMainUnit As String, strDeptID As String)

    cmbMainUnit.ID = strMainUnit
    cmbDepartment.ID = strDeptID
    Me.Show 1
    
End Sub

Public Sub ViewRecords(Optional bShowInactive As Boolean = False)

    On Error GoTo err
   
    Call RefreshRecordSet(bShowInactive)
    
    cmdSave.Visible = False
    cmdOk.Visible = False
    cmdNext.Visible = False
    fraNav.Visible = True
    PicSearch.Visible = True
    
    Me.Show 1
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub RefreshRecordSet(bShowInactive As Boolean)

    Set rsView = New ADODB.Recordset
    
    If bShowInactive Then
        rsView.Open "Select CVID,Name,FName,Religion,Sex From CVs Order By CVID", con, adOpenForwardOnly, adLockReadOnly
    Else
        rsView.Open "Select CVID,Name,FName,Religion,Sex From CVs Order By EmpID", con, adOpenForwardOnly, adLockReadOnly
    End If
    If rsView.EOF = False Then rsView.MoveFirst
    
End Sub

Private Sub FillCmbs()
    
    With cmbMaritalStatus
        .AddItem "Single"
        .AddItem "Married"
        .AddItem "Widow"
        .AddItem "Widower"
        .AddItem "Divorced"
        .ListIndex = 0
    End With

'    With cmbBloodGroup
'        .AddItem "--"
'        .AddItem "A+"
'        .AddItem "B+"
'        .AddItem "AB+"
'        .AddItem "O+"
'        .AddItem "A-"
'        .AddItem "B-"
'        .AddItem "AB-"
'        .AddItem "O-"
'    End With
'
    With cmbReligion
        .AddItem "Islam"
        .AddItem "Chirstian"
        .AddItem "Other"
        .ListIndex = 0
    End With
    
    Call AddToCombo(cmbArea, "DISTINCT Area", "CVs")
        
    With cmbDesig
        .ListHeight = 3000
        .AddVals con, "Designation", "Designations", "Designation", " ORDER BY Designation"
    End With
End Sub


Private Sub chkShowInactive_Click()

    
    
   
    
End Sub

Private Sub cmbAKIWCML_Change()
    Call cmbAKIWCML_Click
End Sub

Private Sub cmbAKIWCML_Click()
    If cmbAKIWCML.ListIndex = 0 Then
        TBox(59) = ""
        TBox(59).Locked = True
    Else
        TBox(59) = ""
        TBox(59).Locked = False
    End If
End Sub

Private Sub cmbAllergy_Change()
    Call cmbAllergy_Click
End Sub

Private Sub cmbAllergy_Click()

    If cmbAllergy.ListIndex = 0 Then
        Label1(101).Visible = False
        TBox(17).Visible = False
        TBox(17) = ""
    Else
        Label1(101).Visible = True
        TBox(17).Visible = True
    End If
    
End Sub

Private Sub cmbDesignation_Change()
    Call cmbDesignation_Click
End Sub

Private Sub cmbDesignation_Click()

'    If cmbDesignation = "Driver" Then
'        TBox(12).Locked = False
'        TBox(13).Locked = False
'        DTLicenseUpto.Enabled = True
'        cmbLicenseTypes.Locked = False
'    Else
'        TBox(12).Locked = True
'        TBox(13).Locked = True
'        DTLicenseUpto.Enabled = False
'        cmbLicenseTypes.Locked = True
'    End If
'
'    TBox(12) = ""
'    TBox(13) = ""
'    DTLicenseUpto.value = Null
'    cmbLicenseTypes.ListIndex = -1
    
End Sub

Private Sub cmbEOBIL_Change()
    Call cmbEOBIL_Click
End Sub

Private Sub cmbEOBIL_Click()

    TBox(52) = ""
    TBox(10) = ""
    DTEOBIIss = Null
    DTMaturity = Null
    
    If cmbEOBIL.ListIndex = 0 Then
        cmbPensionableL.Visible = True
        cmbPensionableL.ListIndex = 0
        'TBox(52).Locked = True
        TBox(52).Left = DTEOBIIss.Left
        Label1(93).Caption = "Status"
        Label1(97).Caption = "EOBI #"
        Label1(98).Caption = "Card #"
        DTMaturity.Enabled = False
        DTEOBIIss.Enabled = False
        DTMaturity.Visible = False
        DTEOBIIss.Visible = False
        TBox(10).Visible = True
        If cmbPensionableL.Visible Then cmbPensionableL.SetFocus
    Else
        cmbPensionableL.Visible = False
        'TBox(52).Locked = False
        TBox(52).Left = cmbPensionableL.Left
        Label1(93).Caption = "EOBI #"
        Label1(97).Caption = "Issue Date"
        Label1(98).Caption = "Date Of Maturity"
        Call DOB_Change
        DTMaturity.Enabled = True
        DTEOBIIss.Enabled = True
        DTMaturity.Visible = True
        DTEOBIIss.Visible = True
        TBox(10).Visible = False
        If TBox(52).Visible Then TBox(52).SetFocus
    End If
    
End Sub

Private Sub cmbFlatL_Change()
    Call cmbFlatL_Click
End Sub

Private Sub cmbFlatL_Click()

    Dim bLocked As Boolean
    Dim StrText As String
    If cmbFlatL.ListIndex = 0 Then
        StrText = "N/A"
        bLocked = True
    Else
        StrText = ""
        bLocked = False
    End If
    
    For i = 60 To 62
        TBox(i) = StrText
        TBox(i).Locked = bLocked
    Next
    TBox(60) = 0
    
End Sub

Private Sub cmbGroupInsuranceL_Change()
    Call cmbGroupInsuranceL_Click
End Sub

Private Sub cmbGroupInsuranceL_Click()
    
    Dim bLocked As Boolean
    If cmbGroupInsuranceL.ListIndex = 0 Then
        bLocked = True
    Else
        bLocked = False
    End If
    For i = 56 To 56
        TBox(i).Locked = bLocked
        TBox(i) = ""
    Next
    lblNominee.Enabled = Not bLocked
    cmbNominees.Locked = bLocked
    cmbNominees.Clear
    TBox(11) = ""
    TBox(14) = ""
   
End Sub

Private Sub cmbMainUnit_matched()
    cmbDepartment.ClearVals
    cmbDepartment.AddVals con, "Name", "Departments", "DeptID", " Where SubDeptOf='" & cmbMainUnit.ID & "'"
End Sub

Private Sub cmbNominees_Change()
    Call cmbNominees_Click
End Sub

Private Sub cmbNominees_Click()
    If cmbNominees.ListIndex <> -1 Then
        TBox(14) = cmbNominees.List(cmbNominees.ListIndex, 1)
        TBox(11) = cmbNominees.List(cmbNominees.ListIndex, 2)
    End If
End Sub

Private Sub cmbPensionableL_Change()
    Call cmbPensionableL_Click
End Sub

Private Sub cmbPensionableL_Click()
    If cmbPensionableL.ListIndex = 0 Then
        Label1(98).Caption = "Remarks"
    Else
        Label1(98).Caption = "Card #"
    End If
End Sub

Private Sub cmbPymtmode_Change()
    Call cmbPymtmode_Click
End Sub

Private Sub cmbPymtmode_Click()

    If cmbPymtmode.ListIndex = 2 Then
        TBox(47) = ""
        TBox(5) = ""
        TBox(47).Locked = False
        TBox(5).Locked = False
    Else
        TBox(47) = "N/A"
        TBox(5) = "N/A"
        TBox(47).Locked = True
        TBox(5).Locked = True
    End If
    
End Sub

Private Sub cmbShare1_Change()
    cmbShare1_Click
End Sub

Private Sub cmbShare1_Click()
    If cmbShare1.ListIndex <> -1 Then
        cmbShare2.ListIndex = 100 - cmbShare1.ListIndex
    End If
End Sub

Private Sub cmbShare2_Change()
    Call cmbShare2_Click
End Sub

Private Sub cmbShare2_Click()
    If cmbShare2.ListIndex <> -1 Then
        cmbShare1.ListIndex = 100 - cmbShare2.ListIndex
    End If
End Sub

Private Sub cmbSSecurityL_Change()
    Call cmbSSecurityL_Click
End Sub

Private Sub cmbSSecurityL_Click()
    
    Dim StrText As String
    DTSSIss = Null
    If cmbSSecurityL.ListIndex = 0 Then
        StrText = "N/A"
        TBox(51).Locked = True
        TBox(53).Locked = True
        DTSSIss.Enabled = False
    Else
        StrText = ""
        TBox(51).Locked = False
        TBox(53).Locked = False
        DTSSIss.Enabled = True
    End If
    TBox(51) = StrText
    TBox(53) = StrText
    
    
End Sub

Private Sub cmbTaxDeductionL_Change()
    Call cmbTaxDeductionL_Click
End Sub

Private Sub cmbTaxDeductionL_Click()

    Dim bLocked As Boolean
    Dim StrText As String
    If cmbTaxDeductionL.ListIndex = 0 Then
        StrText = "N/A"
        bLocked = True
    Else
        StrText = ""
        bLocked = False
    End If
    
    For i = 48 To 49
        TBox(i) = StrText
        TBox(i).Locked = bLocked
    Next
    TBox(16) = 0
    TBox(16).Locked = bLocked
    
End Sub

Private Sub cmbDesig_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
        Dim i As Integer
        If cmbDesig.MatchFound = False Then
            MsgBox "Invalid Designation.", vbInformation
            Exit Sub
        End If
        For i = 1 To LVDesig.ListItems.count
            If cmbDesig.Text = LVDesig.ListItems(i).Text Then
                MsgBox "Already Added.", vbInformation
                Exit Sub
            End If
        Next
    
        LVDesig.ListItems.add , , cmbDesig.Text
    End If
    
End Sub

Private Sub cmdAddAcaQual_Click()

    Dim f As New frmAddQualAca
    Load f
    Dim SNo As String, Diploma As String, Institute As String, strYear As String
    Dim TotalMarks As Integer, MarksObtained As Integer, Percentage As Integer
    
    SNo = LVQ.ListItems.count + 1
    Dim myVal As Integer
    myVal = f.ShowMe(SNo, Diploma, Institute, strYear, TotalMarks, MarksObtained, Percentage)
    If myVal = 1 Or myVal = 2 Then
        'Code To Add It In The List...
        Dim ITM As ListItem
        
        Set ITM = LVQ.ListItems.add(, , SNo)
        ITM.SubItems(1) = Diploma
        ITM.SubItems(2) = strYear
        ITM.SubItems(3) = Institute
        ITM.SubItems(4) = MarksObtained
        ITM.SubItems(5) = TotalMarks
        ITM.SubItems(6) = Percentage & " %"
    End If
    If myVal = 2 Then cmdAddAcaQual_Click
    
End Sub

Private Sub cmdAddComputerLiteracy_Click()

    Dim f As New frmAddComputerLit
    Load f
    Dim SNo As String, Diploma As String, Institute As String, strYear As String
    Dim TotalMarks As Integer, MarksObtained As Integer, Percentage As Integer
    
    SNo = LVComputer.ListItems.count + 1
    Dim myVal As Integer
    myVal = f.ShowMe(SNo, Diploma, Institute, strYear, TotalMarks, MarksObtained, Percentage)
    If myVal = 1 Or myVal = 2 Then
        'Code To Add It In The List...
        Dim ITM As ListItem
        
        Set ITM = LVComputer.ListItems.add(, , SNo)
        ITM.SubItems(1) = Diploma
        ITM.SubItems(2) = Institute
        ITM.SubItems(3) = strYear
        ITM.SubItems(4) = MarksObtained
        ITM.SubItems(5) = TotalMarks
        ITM.SubItems(6) = Percentage & " %"
    End If
    If myVal = 2 Then cmdAddComputerLiteracy_Click
    
End Sub

Private Sub cmdAddExp_Click()

    Dim strInstitute As String, strDesignation As String, strPeriod As String, iSalary As Integer, strWhyLeft As String
    Dim f As New frmAddExperience
    If f.ShowMe(strInstitute, strDesignation, strPeriod, iSalary, strWhyLeft) = False Then Exit Sub
    
    Set myLV = LVExp
    Set ITM = LVExp.ListItems.add(, , LVExp.ListItems.count + 1)
    ITM.ListSubItems.add , , strDesignation
    ITM.ListSubItems.add , , strInstitute
    ITM.ListSubItems.add , , strPeriod
    ITM.ListSubItems.add , , iSalary
    ITM.ListSubItems.add , , strWhyLeft
    
End Sub



Private Sub cmdAddRef_Click()

    Dim strName As String, strOrganization As String, strRelation As String, strContactNo As String, iGuarantorID As Integer
    Dim f As New frmAddReference
    If f.ShowMe(strOrganization, strRelation, strName, strContactNo, iGuarantorID) = False Then Exit Sub
   
    Dim ITM As ListItem
    Set ITM = LVRef.ListItems.add(, , LVRef.ListItems.count + 1)
    ITM.ListSubItems.add , , strName
    ITM.ListSubItems.add , , strRelation
    ITM.ListSubItems.add , , strOrganization
    ITM.ListSubItems.add , , strContactNo
    ITM.ListSubItems.add(, , GetSingleStringValue("strFullName", "GuarantorList_Imported", " WHERE intGuarantorId=" & iGuarantorID)).Tag = iGuarantorID
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
        
End Sub

Private Sub cmdCopyAddress_Click()
    
    TBox(40) = TBox(8)
         
End Sub

Private Sub cmdNav_Click(Index As Integer)

    With rsView
        Select Case Index
            Case 0          '''Move to First Record
                .MoveFirst
            Case 1          '''Move to Previous Record
                .MovePrevious
                If .BOF Then .MoveFirst
            Case 2   '''Move to Next Record
                If .EOF Then .MoveLast
                .MoveNEXT
               
            Case 3          '''Move to Last Record
                .MoveLast
        End Select
    End With
    
End Sub

Private Sub cmdNext_Click()

    If saveCV Then
        'StrTemp = cmbDept.Text
        Unload Me
        Load frmNewCV
        frmNewCV.AddNewCV
    End If
    
End Sub

Private Sub cmdOK_Click()
    If saveCV Then
        Unload Me
        If SqlSettings <> "" Then Unload frmSettings
    End If
End Sub

Private Sub cmdSave_Click()

    If saveCV Then
        EmpId = lblEmpID
    End If

End Sub

Private Sub cmdSearch_Click()

    If cmbField.ListIndex = -1 Then Exit Sub
    If txtValue = "" Then Exit Sub
    
    Dim iCurrPos As Integer
    
    With rsView
        If .EOF Or .BOF Then
            
            .Find cmbField.Text & " like '%" & txtValue & "%'", , adSearchForward, 1
        Else
            iCurrPos = .AbsolutePosition
            .Find cmbField.Text & " like '%" & txtValue & "%'", iCurrPos, adSearchForward, 1
        End If
    End With
    
End Sub

Private Sub DOB_Change()
    TBox(3) = DateDiffInText(DOB.Value)
    If SO(0).Value = True Then  'Male
        DTMaturity = DateAdd("yyyy", 60, DOB)
    Else
        DTMaturity = DateAdd("yyyy", 55, DOB)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()

    Me.KeyPreview = True
    DTPicker1 = #12/31/2007#
    DOB.MaxDate = Date
    DOB = DateAdd("yyyy", -18, Date)
    Call DOB_Change
    Call FillCmbs
    EmpId = ""
    SSTab1.TabVisible(0) = True
   
    TBox(2).Text = "S/O"
    
    bAddNew = False
    SSTab1.Tab = 0
End Sub

Private Function saveCV() As Boolean

    On Error GoTo err

    Dim OpBalance As Double
    
    If Trim(TBox(0)) = "" Then
        MsgBox "Invalid Employee Name.", vbInformation
        saveCV = False
        Exit Function
    ElseIf TBox(4) = "" Then
        MsgBox "Invalid NIC #", vbInformation
        saveCV = False
        Exit Function
    End If
    
    Call StartTrans(con)
    Dim strSQL As String
    If EmpId = "" Then
        strSQL = CreateInsertSQL("CVs")
    Else
        strSQL = CreateUpdateSQL("CVs") & " WHERE CVID='" & EmpId & "'"
    End If
       
    con.Execute strSQL
    
    'Selection
    For i = 0 To OptSelection.count - 1
        If OptSelection(i).Value Then
            con.Execute "UPDATE CVs SET SelectionType=" & i + 1 & " WHERE CVID='" & lblEmpID & "'"
        End If
    Next
    
    'con.Execute "UPDATE Employees SET"
    Call SaveQualifications(lblEmpID)
    
    Call SaveExpnReference(lblEmpID)
   
    Call SaveDesignations(lblEmpID)
SavePic:

    If (Image1.Picture = 0 Or Trim(PicFileName) = "") Then GoTo TheEnd

    Dim Rec As New ADODB.Recordset
    Dim rsPic As New ADODB.Stream
    
    Rec.Open "Select CVID,Pic FROM CVs WHERE CVID='" & lblEmpID & "'", con, adOpenStatic, adLockOptimistic
    
    If Image1.Picture <> 0 And Trim(PicFileName) <> "" Then
      'Save Picture
        rsPic.Open
        rsPic.Type = adTypeBinary
        rsPic.LoadFromFile PicFileName
        Rec.Fields(1) = rsPic.Read
        rsPic.Close
    End If
   
    Rec.Update
    Rec.Close
    Set Rec = Nothing
    Set rsPic = Nothing
     
'If Special = "GENERAL" Then Con.Execute "delete from GeneralSettings Where ID='" & lblEmpID & "'"
TheEnd:
    saveCV = True
    con.CommitTrans
    Exit Function
err:
    'If con.BeginTrans Then con.CommitTrans
    con.RollbackTrans
    MsgBox err.Description
    
End Function

Private Sub AddEmpPic(Optional iOption As Integer)

    On Error Resume Next
    Dim Picfile As String
    CD1.ShowSave
    If Not (err.Number = 32755) Then
        Picfile = CD1.FileName
        If iOption = 0 Then
            Image1.Picture = LoadPicture(Picfile)
            PicFileName = Picfile
            lblPic.Visible = False
        End If
    End If
    
End Sub

Private Sub Image1_DblClick()
    Call AddEmpPic
End Sub

Private Sub JoinDate_Change()
    TBox(9) = DateDiffInText(JoinDate.Value)
End Sub


Private Sub lblPic_DblClick()
    AddEmpPic
End Sub

Private Sub LVComputer_KeyDown(KeyCode As Integer, Shift As Integer)
    If LVComputer.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure to Delete?", vbYesNo + vbQuestion) = vbYes Then
            LVComputer.ListItems.Remove LVComputer.SelectedItem.Index
        End If
    End If
End Sub

Private Sub LVDesig_KeyDown(KeyCode As Integer, Shift As Integer)

    If LVDesig.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        LVDesig.ListItems.Remove LVDesig.SelectedItem.Index
    End If
    
End Sub

Private Sub LVExp_KeyDown(KeyCode As Integer, Shift As Integer)
    If LVExp.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure to Delete?", vbYesNo + vbQuestion) = vbYes Then
            LVExp.ListItems.Remove LVExp.SelectedItem.Index
        End If
    End If
End Sub

Private Sub LVQ_KeyDown(KeyCode As Integer, Shift As Integer)

    If LVQ.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure to Delete?", vbYesNo + vbQuestion) = vbYes Then
            LVQ.ListItems.Remove LVQ.SelectedItem.Index
        End If
    End If

End Sub

Private Sub LVRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If LVRef.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure to Delete?", vbYesNo + vbQuestion) = vbYes Then
            LVRef.ListItems.Remove LVRef.SelectedItem.Index
        End If
    End If
End Sub

Private Sub OptSelection_Click(Index As Integer)
    If Index = 0 Then
        lblDesig.Visible = False
        cmbDesig.Visible = False
        LVDesig.Visible = False
    Else
        lblDesig.Visible = True
        cmbDesig.Visible = True
        LVDesig.Visible = True
    End If
End Sub

Private Sub rsView_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    
    With rsView
        If .EOF = False And .BOF = False Then
            fraNav.Caption = " Showing :   " & CStr(.AbsolutePosition) & " / " & CStr(.RecordCount) & " "
            EmpId = !CVID & ""
            Call ShowCVData
        End If
    End With
    
End Sub

Private Sub SO_Click(Index As Integer)

    On Error GoTo err
    Dim sSex As String
    If Index = 0 Then
        
        TBox(15) = "Male"
        TBox(2) = "S/O"
        TBox(1).SetFocus
        
    ElseIf Index = 1 Then
        
        TBox(15) = "Female"
        TBox(1).SetFocus
        TBox(2) = "D/O"
    ElseIf Index = 2 Then
        
        TBox(15) = "Female"
        TBox(1).SetFocus
        TBox(2) = "W/O"

    End If
    Call DOB_Change
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub SO_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)

    On Error GoTo err
    If KeyAscii = 13 Then
        If SO(Index).Value = False Then
            SO(Index).Value = True
        Else
            TBox(1).SetFocus
        End If
    End If
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub TBox_GotFocus(Index As Integer)
    
    TBox(Index).SelStart = 0
    TBox(Index).SelLength = Len(TBox(Index))
    
End Sub

Private Sub Tbox_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)

    Select Case Index
    Case 3, 16, 41, 42, 59, 56, 60
        KeyAscii = OnlyNumber(Val(KeyAscii))
    
    End Select
    
End Sub

Private Function GetNextItemID(CatID As String) As String
    Dim NextVal As String
    NextVal = GetMax("cast(Right(ItemID,2) as int)", "Items", "Where CatID ='" & CatID & "'")
    GetNextItemID = CatID & "-" & Format(Val(NextVal) + 1, "000")
End Function

Private Sub TBox_LostFocus(Index As Integer)
    If Index = 26 Then
        TBox(Index) = UCase(TBox(26))
    Else
        TBox(Index) = StrConv(TBox(Index), vbProperCase)
    End If
End Sub

Private Sub TBox_Validate(Index As Integer, cancel As Boolean)
    If Index = 41 Then
        If Val(TBox(41)) < 5000 Then
            cmbSSecurityL.ListIndex = 1
        Else
            cmbSSecurityL.ListIndex = 0
        End If
    End If
End Sub



Public Sub EditCV(strEmpID As String, Optional bShowInactive As Boolean = False)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim iCurPos As Integer
    
    EditEmp = True
   
    Set rsView = New ADODB.Recordset
    
'    If bShowInactive Then
'        chkShowInactive.Value = vbChecked
'        'rsView.Open "SELECT EmpID,PrevCode,Name,FName,Designation,Religion,Sex FROM Employees ORDER BY EmpID", con, adOpenForwardOnly, adLockReadOnly
'    Else
'        chkShowInactive.Value = vbUnchecked
'        'rsView.Open "SELECT EmpID,PrevCode,Name,FName,Designation,Religion,Sex FROM Employees WHERE Active=1 ORDER BY EmpID", con, adOpenForwardOnly, adLockReadOnly
'    End If
    rsView.Open "SELECT CVID,Name,FName,Religion,Sex FROM CVs ORDER BY CVID", con, adOpenForwardOnly, adLockReadOnly
    Call RefreshRecordSet(True)
    
    'If rsView.EOF = False Then rsView.MoveFirst   '
    iCurPos = rsView.AbsolutePosition
    EmpId = strEmpID
    rsView.Find "CVID ='" & EmpId & "'", iCurPos, adSearchForward, 1
    lblEmpID.Caption = EmpId
    'Call ShowEmpData
    fraNav.Visible = True
    'cmdNext.Visible = False
    'cmdOK.Visible = False
    cmdSave.Move cmdCancel.Left - 2000, cmdCancel.Top
    Me.Show 1
   
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub ShowCVData()
    
    On Error Resume Next
    Image1.Picture = LoadPicture("")
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM CVs WHERE CVID='" & EmpId & "'", con, adOpenForwardOnly, adLockReadOnly
        
        For Each c In Controls
            If c.Tag <> "" Then
                'Debug.Assert C.Tag <> "CustCode"
                If TypeOf c Is MSForms.CheckBox Then
                    c.Value = Not .Fields(c.Tag)
                ElseIf TypeOf c Is ComboList.Usercontrol1 Then
                    c.ID = .Fields(c.Tag) & ""
              
                    'C.ID = .Fields(C.Tag) & ""
                ElseIf TypeOf c Is MSForms.Label Or TypeOf c Is VB.Label Then
                    c.Caption = .Fields(c.Tag) & ""
                ElseIf TypeOf c Is DTPicker Then
                    c.Value = .Fields(c.Tag) & ""
                ElseIf (TypeOf c Is MSForms.ComboBox) Or (TypeOf c Is VB.ComboBox) Then
                
                    If Right(c.Name, 1) = "L" Then
                        If Not IsNull(.Fields(c.Tag)) Then
                            c.ListIndex = Abs(.Fields(c.Tag))
                        Else
                            c.ListIndex = -1
                        End If
                    Else
                        If c.Name = "cmbShare1" Then
                            c.Text = .Fields(c.Tag) & "%"
                        ElseIf c.Name = "cmbArea" Then
                            c.Text = .Fields(c.Tag) & ""
                        Else
                            'C.Text = .Fields(C.Tag) & ""
                            For i = 0 To c.ListCount - 1
                                If c.List(i) = .Fields(c.Tag) & "" Then
                                    c.ListIndex = i
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                ElseIf (TypeOf c Is MSForms.OptionButton) Then
                    ''''
                ElseIf (TypeOf c Is VB.CheckBox) Then
                    c.Value = Abs(.Fields(c.Tag))
                Else
                    c.Text = .Fields(c.Tag) & ""
                End If
            End If
        Next
        TBox(2).Text = "S/O"
        If Not IsNull(![Pic]) Then
            Set Image1.DataSource = rs
            Image1.DataField = "Pic"
            lblPic.Visible = False
        Else
            lblPic.Visible = True
        End If
        
        If Val(!SelectionType & "") > 0 Then
            OptSelection(Val(!SelectionType & "") - 1).Value = True
        End If
        .Close
       
    
    End With
    
    Set rs = Nothing
    
    Call DOB_Change
    
    Call LoadQualifications(EmpId)
    Call LoadExpnReference(EmpId)
    Call LoadDesignations(EmpId)
    
End Sub



Private Function CreateInsertSQL(strTable As String, Optional With_ As Boolean = False) As String
    
    Dim c As Control
    Dim RetString As String
    
    'Create The First Part Of Insert Query...
    RetString = "Insert Into " & strTable & "("
    '---------------------------------------------------------------------------
    
    For Each c In Controls
        If c.Tag <> "" Then
            'If (TypeOf C Is MSForms.TextBox) Or (TypeOf C Is MSForms.ComboBox) Or _
            (TypeOf C Is MSForms.CheckBox) Or (TypeOf C Is ComboList.Usercontrol1) Or _
            (TypeOf C Is VB.TextBox) Or (TypeOf C Is VB.ComboBox) Then
            
            RetString = RetString & c.Tag & ","
            
            'End If
        End If
    Next
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    RetString = RetString & ") Values("
    
    
    'Now The Second Part...
    '---------------------------------------------------------------------------
    
    For Each c In Controls
        
        If c.Tag <> "" Then
            If (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is VB.TextBox) Then
                RetString = RetString & "'" & c.Text & "',"
            ElseIf (TypeOf c Is MSForms.CheckBox) Or (TypeOf c Is VB.CheckBox) Then
                RetString = RetString & Abs(c.Value) & ","
            ElseIf (TypeOf c Is ComboList.Usercontrol1) Then
                RetString = RetString & "'" & c.ID & "',"
            ElseIf (TypeOf c Is MSForms.ComboBox) Or (TypeOf c Is VB.ComboBox) Then
                If Right(c.Name, 1) = "L" Then
                    RetString = RetString & c.ListIndex & ","
                Else
                    If c.Name = "cmbShare1" Then
                        RetString = RetString & Val(c.Text) & ","
                    Else
                        RetString = RetString & "'" & c.Text & "',"
                    End If
                End If
            ElseIf (TypeOf c Is MSForms.OptionButton) Then
                If c.Value = True Then
                    RetString = RetString & "'" & c.Caption & "',"
                End If
            ElseIf (TypeOf c Is DTPicker) Then
                If c.Value & "" = "" Then
                    RetString = RetString & "Null,"
                Else
                    RetString = RetString & "'" & Format(c.Value, "dd-MMM-yyyy") & "',"
                End If
            Else
                RetString = RetString & "'" & c & "',"
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

Private Function CreateUpdateSQL(strTable As String, Optional With_ As Boolean = False) As String
    
    Dim c As Control
    Dim RetString As String
    
    'Create The First Part Of Insert Query...
    RetString = "Update " & strTable & " Set "
    '---------------------------------------------------------------------------
    For Each c In Controls
    
        If c.Tag <> "" And c.Name <> "lblEmpID" Then
            If (TypeOf c Is MSForms.TextBox) Then   '
                If c.Tag = "name" Then
                    RetString = RetString & "[Name]" & "='" & c.Text & "',"
                ElseIf LCase(c.Name) Like "nbox*" Then
                         RetString = RetString & c.Tag & "=" & c.Text & ","
                Else
                     RetString = RetString & c.Tag & "='" & c.Text & "',"
                End If
                
            ElseIf (TypeOf c Is MSForms.CheckBox) Or (TypeOf c Is VB.CheckBox) Then
                RetString = RetString & c.Tag & "=" & Abs(c.Value) & ","
            ElseIf (TypeOf c Is ComboList.Usercontrol1) Then
                RetString = RetString & c.Tag & "='" & c.ID & "',"
            ElseIf (TypeOf c Is MSForms.ComboBox) Or (TypeOf c Is VB.ComboBox) Then
                If Right(c.Name, 1) = "L" Then
                    RetString = RetString & c.Tag & "=" & c.ListIndex & ","
                Else
                    If c.Name = "cmbShare1" Then
                        RetString = RetString & c.Tag & "=" & Val(c) & ","
                    Else
                        RetString = RetString & c.Tag & "='" & c & "',"
                    End If
                End If
            ElseIf (TypeOf c Is DTPicker) Then
                If c.Value & "" = "" Then
                    RetString = RetString & c.Tag & "=Null,"
                Else
                    RetString = RetString & c.Tag & "='" & Format(c.Value, "dd-MMM-yyyy") & "',"
                End If
     
            Else
                RetString = RetString & c.Tag & "='" & c & "',"
            End If
        End If
    Next
    
    'Get For ListView
    'For i = 1 To LV2.ListItems.Count
    '    RetString = RetString & LV2.ListItems(i).Key & "=" & Abs(LV2.ListItems(i).Checked) & ","
    'Next i
    '---------------------------------------------------------------------------
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    
    CreateUpdateSQL = RetString
    
End Function


Private Sub SaveExpnReference(EmpId As String)
    'SaveDesignations
    con.Execute "DELETE FROM CVExperience WHERE CVID='" & EmpId & "'"
    
    For i = 1 To LVExp.ListItems.count
        With LVExp.ListItems(i)
            con.Execute "INSERT INTO CVExperience(CVID,SrNo,Designation,Institute,Experience,Salary,WhyLeft) VALUES('" & _
             EmpId & "'," & Val(.Text) & ",'" & .SubItems(1) & "','" & .SubItems(2) & "'," & _
             Val(.SubItems(3)) & "," & Val(.SubItems(4)) & ",'" & .SubItems(5) & "')"
        End With
    Next
    
    con.Execute "DELETE FROM CVReference WHERE CVID='" & EmpId & "'"
    For i = 1 To LVRef.ListItems.count
        With LVRef.ListItems(i)
            con.Execute "INSERT INTO CVReference(CVID,SrNo,Name,Relation,Organization,ContactNo,Guarantor_RefID) VALUES('" & _
             EmpId & "'," & Val(.Text) & ",'" & .SubItems(1) & "','" & .SubItems(2) & "','" & .SubItems(3) & "','" & _
             .SubItems(4) & "'," & Val(.ListSubItems(5).Tag) & ")"
        End With
    Next
    
End Sub

Private Sub SaveDesignations(EmpId As String)
    '
    con.Execute "DELETE FROM CVDesignations WHERE CVID='" & EmpId & "'"
    
    For i = 1 To LVDesig.ListItems.count
        With LVDesig.ListItems(i)
            con.Execute "INSERT INTO CVDesignations(CVID,Designation) VALUES('" & _
             EmpId & "','" & .Text & "')"
        End With
    Next
    
End Sub


Private Sub SaveQualifications(EmpId As String)

    con.Execute "DELETE FROM CVAcademicQualifications WHERE CVID='" & EmpId & "'"

    For i = 1 To LVQ.ListItems.count
        With LVQ.ListItems(i)
            con.Execute "INSERT INTO CVAcademicQualifications (CVID,SrNo,Diploma,Institute," & _
             "PassingYear,MarksObtained,TotalMarks)VALUES('" & EmpId & "'," & .Text & ",'" & .SubItems(1) & _
             "','" & .SubItems(3) & "'," & .SubItems(2) & "," & Val(.SubItems(4)) & "," & Val(.SubItems(5)) & ")"
        End With
    Next
    
    con.Execute "DELETE FROM CVComputerLiteracy WHERE CVID='" & EmpId & "'"

    For i = 1 To LVComputer.ListItems.count
        With LVComputer.ListItems(i)
            con.Execute "INSERT INTO CVComputerLiteracy(CVID,SrNo,Diploma,Institute," & _
             "Period,MarksObtained,TotalMarks)VALUES('" & EmpId & "'," & .Text & ",'" & .SubItems(1) & _
             "','" & .SubItems(3) & "','" & .SubItems(2) & "'," & Val(.SubItems(4)) & "," & Val(.SubItems(5)) & ")"
        End With
    Next
    
End Sub

Private Sub LoadQualifications(EnrollNo As String)

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "Select * From CVAcademicQualifications Where CVID='" & EmpId & "' Order By SrNo", con, adOpenForwardOnly, adLockReadOnly
        LVQ.ListItems.Clear
        Do Until .EOF
            Set ITM = LVQ.ListItems.add(, , ![SrNo] & "")
            ITM.SubItems(1) = !Diploma & ""
            ITM.SubItems(2) = !PassingYear & ""
            ITM.SubItems(3) = !Institute & ""
            
            ITM.SubItems(4) = Val(!MarksObtained & "")
            ITM.SubItems(5) = Val(!TotalMarks & "")
            ITM.SubItems(6) = Round(((Val(ITM.SubItems(4)) / Val(ITM.SubItems(5))) * 100), 2) & " %"
            .MoveNEXT
        Loop
        .Close
        
        .Open "Select * From CVComputerLiteracy Where CVID='" & EmpId & "' Order By SrNo", con, adOpenForwardOnly, adLockReadOnly
        LVComputer.ListItems.Clear
        Do Until .EOF
            Set ITM = LVComputer.ListItems.add(, , ![SrNo] & "")
            ITM.SubItems(1) = !Diploma & ""
            ITM.SubItems(2) = !Period & ""
            ITM.SubItems(3) = !Institute & ""
            
            ITM.SubItems(4) = Val(!MarksObtained & "")
            ITM.SubItems(5) = Val(!TotalMarks & "")
            If Val(ITM.SubItems(5)) = 0 Then
                ITM.SubItems(6) = "0"
            Else
                ITM.SubItems(6) = Round(((Val(ITM.SubItems(4)) / Val(ITM.SubItems(5))) * 100), 2) & " %"
            End If
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub


Private Sub LoadExpnReference(EmpId As String)
    'LoadDesignations
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "SELECT * FROM CVExperience WHERE CVID='" & EmpId & "' ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LVExp.ListItems.Clear
        Do Until .EOF
           Set ITM = LVExp.ListItems.add(, , LVExp.ListItems.count + 1)
           ITM.ListSubItems.add , , !Designation & ""
           ITM.ListSubItems.add , , !Institute & ""
           ITM.ListSubItems.add , , Val(!Experience & "")
           ITM.ListSubItems.add , , Val(!Salary & "")
           ITM.ListSubItems.add , , !WhyLeft & ""
           .MoveNEXT
        Loop
        .Close
         
        .Open "SELECT * FROM VCVReference WHERE CVID='" & EmpId & "' ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LVRef.ListItems.Clear
        Do Until .EOF
            Set ITM = LVRef.ListItems.add(, , LVRef.ListItems.count + 1)
            ITM.ListSubItems.add , , !Name & ""
            ITM.ListSubItems.add , , !Relation & ""
            ITM.ListSubItems.add , , !Organization & ""
            ITM.ListSubItems.add , , !ContactNo & ""
            
            ITM.ListSubItems.add(, , !strFullName & "").Tag = Val(!Guarantor_RefID & "")
           .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
End Sub

Private Sub LoadDesignations(EmpId As String)
    '
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "SELECT * FROM CVDesignations WHERE CVID='" & EmpId & "' ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LVDesig.ListItems.Clear
        Do Until .EOF
           Set ITM = LVDesig.ListItems.add(, , !Designation & "")
           .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
End Sub

Public Function AddNewCV()
    lblEmpID = GetNextCVNo
    Me.Show
End Function

Private Function GetNextCVNo()

   
    
    Dim NextVal As String
    
    NextVal = GetMax("CAST(RIGHT(CVID,LEN(CVID)-3) AS INT)", "CVs", "WHERE LEFT(CVID,3)='CV-'")
    GetNextCVNo = "CV-" & Format(Val(NextVal) + 1, "000")
   
End Function
