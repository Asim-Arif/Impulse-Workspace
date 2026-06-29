VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTHTC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "THTC Income Percentage"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   5610
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   180
      TabIndex        =   56
      Top             =   -15
      Width           =   5205
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "THTC Income Percentage"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   360
         Index           =   21
         Left            =   810
         TabIndex        =   57
         Top             =   150
         Width           =   3525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "THTC Income Percentage"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   22
         Left            =   825
         TabIndex        =   58
         Top             =   165
         Width           =   3525
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3345
      Left            =   150
      TabIndex        =   3
      Top             =   645
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   5900
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "ECG"
      TabPicture(0)   =   "frmTHTC.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ETT"
      TabPicture(1)   =   "frmTHTC.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "ECHO"
      TabPicture(2)   =   "frmTHTC.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Cardiac T"
      TabPicture(3)   =   "frmTHTC.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame4 
         Height          =   2625
         Left            =   -74650
         TabIndex        =   43
         Top             =   465
         Width           =   4455
         Begin VB.TextBox txtCardiacT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   4
            Left            =   3000
            TabIndex        =   62
            Top             =   1470
            Width           =   900
         End
         Begin VB.TextBox txtCardiacT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   3
            Left            =   2235
            TabIndex        =   47
            Top             =   2070
            Width           =   435
         End
         Begin VB.TextBox txtCardiacT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   0
            Left            =   2640
            TabIndex        =   46
            Top             =   465
            Width           =   1260
         End
         Begin VB.TextBox txtCardiacT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   1
            Left            =   2640
            TabIndex        =   45
            Top             =   825
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtCardiacT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   2
            Left            =   2235
            TabIndex        =   44
            Top             =   1470
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total Income :"
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
            Index           =   18
            Left            =   300
            TabIndex        =   55
            Top             =   1238
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblCardiacT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   0
            Left            =   2640
            TabIndex        =   54
            Top             =   1185
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total CardiacT Charges Rs :"
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
            Index           =   20
            Left            =   300
            TabIndex        =   53
            Top             =   510
            Width           =   2295
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cardiac T Cost Rs. :"
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
            Index           =   19
            Left            =   300
            TabIndex        =   52
            Top             =   878
            Visible         =   0   'False
            Width           =   1590
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Consultant Charges @           % :"
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
            Index           =   17
            Left            =   360
            TabIndex        =   51
            Top             =   1530
            Width           =   2625
         End
         Begin VB.Label lblCardiacT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   1
            Left            =   735
            TabIndex        =   50
            Top             =   1785
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   7
            X1              =   285
            X2              =   3910
            Y1              =   1125
            Y2              =   1125
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   6
            X1              =   285
            X2              =   3910
            Y1              =   1110
            Y2              =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Heart Care Society  @           % :"
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
            Left            =   360
            TabIndex        =   49
            Top             =   2130
            Width           =   2610
         End
         Begin VB.Label lblCardiacT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   2
            Left            =   3000
            TabIndex        =   48
            Top             =   2070
            Width           =   900
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2625
         Left            =   -74650
         TabIndex        =   30
         Top             =   465
         Width           =   4455
         Begin VB.TextBox txtEcho 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   4
            Left            =   3000
            TabIndex        =   61
            Top             =   1395
            Width           =   900
         End
         Begin VB.TextBox txtEcho 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   3
            Left            =   2235
            TabIndex        =   34
            Top             =   2070
            Width           =   435
         End
         Begin VB.TextBox txtEcho 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   0
            Left            =   2490
            TabIndex        =   33
            Top             =   465
            Width           =   1425
         End
         Begin VB.TextBox txtEcho 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   1
            Left            =   2475
            TabIndex        =   32
            Top             =   855
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.TextBox txtEcho 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   2
            Left            =   2235
            TabIndex        =   31
            Top             =   1395
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total ECHO Charges  Rs. :"
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
            Index           =   14
            Left            =   360
            TabIndex        =   42
            Top             =   510
            Width           =   2085
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "ECHO Cost Rs. :"
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
            Index           =   13
            Left            =   345
            TabIndex        =   41
            Top             =   855
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total Income :"
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
            Index           =   12
            Left            =   345
            TabIndex        =   40
            Top             =   855
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblECHO 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   0
            Left            =   2475
            TabIndex        =   39
            Top             =   855
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Consultant Charges @           % :"
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
            Left            =   360
            TabIndex        =   38
            Top             =   1455
            Width           =   2625
         End
         Begin VB.Label lblECHO 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   1
            Left            =   315
            TabIndex        =   37
            Top             =   1740
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   5
            X1              =   285
            X2              =   3910
            Y1              =   1140
            Y2              =   1140
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   4
            X1              =   285
            X2              =   3910
            Y1              =   1125
            Y2              =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Heart Care Society  @           % :"
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
            Index           =   10
            Left            =   360
            TabIndex        =   36
            Top             =   2130
            Width           =   2610
         End
         Begin VB.Label lblECHO 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   2
            Left            =   3000
            TabIndex        =   35
            Top             =   2070
            Width           =   900
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2625
         Left            =   -74650
         TabIndex        =   17
         Top             =   465
         Width           =   4455
         Begin VB.TextBox txtETT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   4
            Left            =   3015
            TabIndex        =   60
            Top             =   1470
            Width           =   900
         End
         Begin VB.TextBox txtETT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   2
            Left            =   2235
            TabIndex        =   21
            Top             =   1455
            Width           =   435
         End
         Begin VB.TextBox txtETT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   1
            Left            =   2490
            TabIndex        =   20
            Top             =   1080
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.TextBox txtETT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   0
            Left            =   2475
            TabIndex        =   19
            Top             =   465
            Width           =   1425
         End
         Begin VB.TextBox txtETT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   3
            Left            =   2235
            TabIndex        =   18
            Top             =   2070
            Width           =   435
         End
         Begin VB.Label lblETT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   2
            Left            =   3015
            TabIndex        =   29
            Top             =   2070
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Heart Care Society  @           % :"
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
            Index           =   9
            Left            =   360
            TabIndex        =   28
            Top             =   2130
            Width           =   2610
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   3
            X1              =   285
            X2              =   3910
            Y1              =   1110
            Y2              =   1110
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   2
            X1              =   285
            X2              =   3910
            Y1              =   1125
            Y2              =   1125
         End
         Begin VB.Label lblETT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   1
            Left            =   240
            TabIndex        =   27
            Top             =   1740
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Consultant Charges @           % :"
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
            Index           =   8
            Left            =   360
            TabIndex        =   26
            Top             =   1515
            Width           =   2625
         End
         Begin VB.Label lblETT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   0
            Left            =   2490
            TabIndex        =   25
            Top             =   795
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total Income :"
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
            Index           =   7
            Left            =   375
            TabIndex        =   24
            Top             =   960
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "ETT Cost Rs. :"
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
            Index           =   6
            Left            =   375
            TabIndex        =   23
            Top             =   795
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total ETT Charges  Rs.    :"
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
            Left            =   360
            TabIndex        =   22
            Top             =   510
            Width           =   2085
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2625
         Left            =   350
         TabIndex        =   4
         Top             =   465
         Width           =   4455
         Begin VB.TextBox txtECG 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   4
            Left            =   3000
            TabIndex        =   59
            Top             =   1455
            Width           =   900
         End
         Begin VB.TextBox txtECG 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   3
            Left            =   2235
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   2070
            Width           =   435
         End
         Begin VB.TextBox txtECG 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   0
            Left            =   2475
            TabIndex        =   7
            Top             =   465
            Width           =   1425
         End
         Begin VB.TextBox txtECG 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   1
            Left            =   2475
            TabIndex        =   6
            Top             =   825
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.TextBox txtECG 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   2
            Left            =   2235
            TabIndex        =   5
            Top             =   1455
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total ECG Charges  Rs.    :"
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
            Left            =   360
            TabIndex        =   15
            Top             =   510
            Width           =   2100
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "ECG Cost Rs. :"
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
            Index           =   1
            Left            =   360
            TabIndex        =   14
            Top             =   795
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total Income :"
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
            Index           =   2
            Left            =   360
            TabIndex        =   13
            Top             =   1170
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblECG 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   0
            Left            =   2475
            TabIndex        =   12
            Top             =   1200
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Consultant Charges @           % :"
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
            Index           =   3
            Left            =   360
            TabIndex        =   11
            Top             =   1515
            Width           =   2625
         End
         Begin VB.Label lblECG 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   1
            Left            =   270
            TabIndex        =   10
            Top             =   1725
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   1
            X1              =   285
            X2              =   3910
            Y1              =   1095
            Y2              =   1095
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   0
            X1              =   285
            X2              =   3910
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Heart Care Society  @           % :"
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
            Index           =   5
            Left            =   360
            TabIndex        =   9
            Top             =   2130
            Width           =   2610
         End
         Begin VB.Label lblECG 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   2
            Left            =   3000
            TabIndex        =   8
            Top             =   2070
            Width           =   900
         End
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Serial No's :"
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
      Index           =   15
      Left            =   7890
      TabIndex        =   2
      Top             =   15
      Width           =   960
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   405
      Left            =   1965
      TabIndex        =   0
      Top             =   4095
      Width           =   1665
      Caption         =   " Save & Close"
      PicturePosition =   327683
      Size            =   "2937;714"
      Picture         =   "frmTHTC.frx":0070
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
      Height          =   405
      Left            =   3750
      TabIndex        =   1
      Top             =   4095
      Width           =   1665
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "2937;714"
      Picture         =   "frmTHTC.frx":01CE
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
Attribute VB_Name = "frmTHTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    ElseIf TypeOf Screen.ActiveControl Is TextBox Then
        KeyAscii = OnlyNumber(KeyAscii, True)
    End If
End Sub

Private Sub Form_Load()
    Me.KeyPreview = True
   
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select * FRom THTCPercentages", con, adOpenKeyset, adLockReadOnly
        
        '''''''''''' ECG ''''''''''''''''''
        .MoveFirst
        .Find "IncomeSource='ECG'"
        txtECG(0) = Format(Val(![Price] & ""), "0.00")
        txtECG(1) = Format(Val(![Cost] & ""), "0.00")
        txtECG(2) = Format(Val(![Consaltant] & ""), "0.00")
        txtECG(3) = Format(Val(![HCS] & ""), "0.00")
        
        
        '''''''''''' ETT ''''''''''''''''''
        .MoveFirst
        .Find "IncomeSource='ETT'"
        txtETT(0) = Format(Val(![Price] & ""), "0.00")
        txtETT(1) = Format(Val(![Cost] & ""), "0.00")
        txtETT(2) = Format(Val(![Consaltant] & ""), "0.00")
        txtETT(3) = Format(Val(![HCS] & ""), "0.00")
        
        
        '''''''''''' ECHO ''''''''''''''''''
        .MoveFirst
        .Find "IncomeSource='ECHO'"
        txtEcho(0) = Format(Val(![Price] & ""), "0.00")
        txtEcho(1) = Format(Val(![Cost] & ""), "0.00")
        txtEcho(2) = Format(Val(![Consaltant] & ""), "0.00")
        txtEcho(3) = Format(Val(![HCS] & ""), "0.00")
        
        
        
        
        '''''''''''' CardiacT ''''''''''''''''''
        .MoveFirst
        .Find "IncomeSource='CardiacT'"
        txtCardiacT(0) = Format(Val(![Price] & ""), "0.00")
        txtCardiacT(1) = Format(Val(![Cost] & ""), "0.00")
        txtCardiacT(2) = Format(Val(![Consaltant] & ""), "0.00")
        txtCardiacT(3) = Format(Val(![HCS] & ""), "0.00")
        .Close
    End With
    Set rs = Nothing
End Sub

Private Function Saved() As Boolean

On Error GoTo err




   con.BeginTrans

   con.Execute "Update THTCPercentages Set Price=" & Val(txtECG(0)) & ",Cost=" & Val(txtECG(1)) & _
   ",Consaltant=" & Val(txtECG(2)) & ",HCS=" & Val(txtECG(3)) & " Where IncomeSource='ECG'"
   
   
   con.Execute "Update THTCPercentages Set Price=" & Val(txtETT(0)) & ",Cost=" & Val(txtETT(1)) & _
   ",Consaltant=" & Val(txtETT(2)) & ",HCS=" & Val(txtETT(3)) & " Where IncomeSource='ETT'"
   
   con.Execute "Update THTCPercentages Set Price=" & Val(txtEcho(0)) & ",Cost=" & Val(txtEcho(1)) & _
   ",Consaltant=" & Val(txtEcho(2)) & ",HCS=" & Val(txtEcho(3)) & " Where IncomeSource='ECHO'"
   
   con.Execute "Update THTCPercentages Set Price=" & Val(txtCardiacT(0)) & ",Cost=" & Val(txtCardiacT(1)) & _
   ",Consaltant=" & Val(txtCardiacT(2)) & ",HCS=" & Val(txtCardiacT(3)) & " Where IncomeSource='CardiacT'"
   
   con.CommitTrans

 Saved = True
Exit Function

err:
    MsgBox err.Description
End Function



Private Sub txtECG_Change(Index As Integer)
    If Index <> 4 Then
        lblECG(0) = Format(Val(txtECG(0)) - Val(txtECG(1)), "0.00 ")
        txtECG(4) = Format(Val(lblECG(0)) * Val(txtECG(2)) / 100, "0.00 ")
        txtECG(3) = Format(100 - Val(txtECG(2)), "0.00")
        lblECG(2) = Format(Val(lblECG(0)) * Val(txtECG(3)) / 100, "0.00 ")
    End If
End Sub

Private Sub txtECG_GotFocus(Index As Integer)
    txtECG(Index).SelStart = 0
    txtECG(Index).SelLength = Len(txtECG(Index))
End Sub

Private Sub txtECG_LostFocus(Index As Integer)
    txtECG(2) = Format(Val(txtECG(4)) / Val(lblECG(0)) * 100, "0.00")
End Sub
Private Sub txtECHO_LostFocus(Index As Integer)
      txtEcho(2) = Format(Val(txtEcho(4)) / Val(lblECHO(0)) * 100, "0.00")
End Sub
Private Sub txtETT_LostFocus(Index As Integer)
      txtETT(2) = Format(Val(txtETT(4)) / Val(lblETT(0)) * 100, "0.00")
End Sub
Private Sub txtCardiacT_LostFocus(Index As Integer)
  txtCardiacT(2) = Format(Val(txtCardiacT(4)) / Val(lblCardiacT(0)) * 100, "0.00")
End Sub

Private Sub txtETT_GotFocus(Index As Integer)
    txtETT(Index).SelStart = 0
    txtETT(Index).SelLength = Len(txtETT(Index))
End Sub

Private Sub txtECHO_GotFocus(Index As Integer)
    txtEcho(Index).SelStart = 0
    txtEcho(Index).SelLength = Len(txtEcho(Index))
End Sub

Private Sub txtCardiacT_GotFocus(Index As Integer)
    txtCardiacT(Index).SelStart = 0
    txtCardiacT(Index).SelLength = Len(txtCardiacT(Index))
End Sub

Private Sub txtETT_Change(Index As Integer)
If Index <> 4 Then
    lblETT(0) = Format(Val(txtETT(0)) - Val(txtETT(1)), "0.00 ")
    txtETT(4) = Format(Val(lblETT(0)) * Val(txtETT(2)) / 100, "0.00 ")
    txtETT(3) = 100 - Val(txtETT(2))
    lblETT(2) = Format(Val(lblETT(0)) * Val(txtETT(3)) / 100, "0.00 ")
End If
End Sub



Private Sub txtECHO_Change(Index As Integer)
If Index <> 4 Then
    lblECHO(0) = Format(Val(txtEcho(0)) - Val(txtEcho(1)), "0.00 ")
    txtEcho(4) = Format(Val(lblECHO(0)) * Val(txtEcho(2)) / 100, "0.00 ")
    txtEcho(3) = 100 - Val(txtEcho(2))
    lblECHO(2) = Format(Val(lblECHO(0)) * Val(txtEcho(3)) / 100, "0.00 ")
End If
End Sub

Private Sub txtCardiacT_Change(Index As Integer)
If Index <> 4 Then
    lblCardiacT(0) = Format(Val(txtCardiacT(0)) - Val(txtCardiacT(1)), "0.00 ")
    txtCardiacT(4) = Format(Val(lblCardiacT(0)) * Val(txtCardiacT(2)) / 100, "0.00 ")
    txtCardiacT(3) = 100 - Val(txtCardiacT(2))
    lblCardiacT(2) = Format(Val(lblCardiacT(0)) * Val(txtCardiacT(3)) / 100, "0.00 ")
End If
End Sub






