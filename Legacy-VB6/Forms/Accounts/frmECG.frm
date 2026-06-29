VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmECG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Detail"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   7875
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   3645
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4890
      _ExtentX        =   8625
      _ExtentY        =   6429
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "ECG"
      TabPicture(0)   =   "frmECG.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ETT"
      TabPicture(1)   =   "frmECG.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "ECHO"
      TabPicture(2)   =   "frmECG.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Cardiac T"
      TabPicture(3)   =   "frmECG.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame4 
         Height          =   2985
         Left            =   -74775
         TabIndex        =   47
         Top             =   480
         Width           =   4455
         Begin VB.TextBox txtCardiacTPer2 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2235
            TabIndex        =   51
            Top             =   2070
            Width           =   435
         End
         Begin VB.TextBox txtCardiacTPrice 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2640
            TabIndex        =   50
            Top             =   465
            Width           =   1260
         End
         Begin VB.TextBox txtCardiacTCost 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2640
            TabIndex        =   49
            Top             =   825
            Width           =   1260
         End
         Begin VB.TextBox txtCardiacTPer1 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2235
            TabIndex        =   48
            Top             =   1680
            Width           =   435
         End
         Begin VB.Label label1 
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
            TabIndex        =   59
            Top             =   1238
            Width           =   1215
         End
         Begin VB.Label lblCardiacTIncome 
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
            Left            =   2640
            TabIndex        =   58
            Top             =   1185
            Width           =   1260
         End
         Begin VB.Label label1 
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
            TabIndex        =   57
            Top             =   510
            Width           =   2295
         End
         Begin VB.Label label1 
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
            TabIndex        =   56
            Top             =   878
            Width           =   1590
         End
         Begin VB.Label label1 
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
            TabIndex        =   55
            Top             =   1740
            Width           =   2625
         End
         Begin VB.Label lblCardiacTConsaltant 
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
            Left            =   3000
            TabIndex        =   54
            Top             =   1680
            Width           =   900
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   7
            X1              =   285
            X2              =   3910
            Y1              =   1575
            Y2              =   1575
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   6
            X1              =   285
            X2              =   3910
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Label label1 
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
            TabIndex        =   53
            Top             =   2130
            Width           =   2610
         End
         Begin VB.Label lblCardiacTHCS 
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
            Left            =   3000
            TabIndex        =   52
            Top             =   2070
            Width           =   900
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2985
         Left            =   -74775
         TabIndex        =   34
         Top             =   465
         Width           =   4455
         Begin VB.TextBox txtECHOPer2 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2235
            TabIndex        =   38
            Top             =   2070
            Width           =   435
         End
         Begin VB.TextBox txtECHOPrice 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2475
            TabIndex        =   37
            Top             =   465
            Width           =   1425
         End
         Begin VB.TextBox txtEchoCost 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2475
            TabIndex        =   36
            Top             =   825
            Width           =   1425
         End
         Begin VB.TextBox txtECHOPer1 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2235
            TabIndex        =   35
            Top             =   1680
            Width           =   435
         End
         Begin VB.Label label1 
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
            TabIndex        =   46
            Top             =   510
            Width           =   2085
         End
         Begin VB.Label label1 
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
            Left            =   360
            TabIndex        =   45
            Top             =   878
            Width           =   1245
         End
         Begin VB.Label label1 
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
            Left            =   360
            TabIndex        =   44
            Top             =   1170
            Width           =   1215
         End
         Begin VB.Label lblECHOIncome 
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
            Left            =   2490
            TabIndex        =   43
            Top             =   1185
            Width           =   1425
         End
         Begin VB.Label label1 
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
            TabIndex        =   42
            Top             =   1740
            Width           =   2625
         End
         Begin VB.Label lblECHOConsaltant 
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
            Left            =   3000
            TabIndex        =   41
            Top             =   1680
            Width           =   900
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   5
            X1              =   285
            X2              =   3910
            Y1              =   1575
            Y2              =   1575
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   4
            X1              =   285
            X2              =   3910
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Label label1 
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
            TabIndex        =   40
            Top             =   2130
            Width           =   2610
         End
         Begin VB.Label lblECHOHCS 
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
            Left            =   3000
            TabIndex        =   39
            Top             =   2070
            Width           =   900
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2985
         Left            =   -74775
         TabIndex        =   21
         Top             =   465
         Width           =   4455
         Begin VB.TextBox txtETTper1 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2235
            TabIndex        =   25
            Top             =   1680
            Width           =   435
         End
         Begin VB.TextBox txtETTCost 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2475
            TabIndex        =   24
            Top             =   825
            Width           =   1425
         End
         Begin VB.TextBox txtETTPrice 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2475
            TabIndex        =   23
            Top             =   465
            Width           =   1425
         End
         Begin VB.TextBox txtETTper2 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2235
            TabIndex        =   22
            Top             =   2070
            Width           =   435
         End
         Begin VB.Label lblETTHCS 
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
            Left            =   3000
            TabIndex        =   33
            Top             =   2070
            Width           =   900
         End
         Begin VB.Label label1 
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
            TabIndex        =   32
            Top             =   2130
            Width           =   2610
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   3
            X1              =   285
            X2              =   3910
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   2
            X1              =   285
            X2              =   3910
            Y1              =   1575
            Y2              =   1575
         End
         Begin VB.Label lblETTConsaltant 
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
            Left            =   3000
            TabIndex        =   31
            Top             =   1695
            Width           =   900
         End
         Begin VB.Label label1 
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
            TabIndex        =   30
            Top             =   1740
            Width           =   2625
         End
         Begin VB.Label lblETTIncome 
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
            Left            =   2475
            TabIndex        =   29
            Top             =   1185
            Width           =   1425
         End
         Begin VB.Label label1 
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
            Left            =   360
            TabIndex        =   28
            Top             =   1170
            Width           =   1215
         End
         Begin VB.Label label1 
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
            Left            =   360
            TabIndex        =   27
            Top             =   795
            Width           =   1110
         End
         Begin VB.Label label1 
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
            TabIndex        =   26
            Top             =   510
            Width           =   2085
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2985
         Left            =   225
         TabIndex        =   8
         Top             =   465
         Width           =   4455
         Begin VB.TextBox txtECGPer2 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2235
            TabIndex        =   20
            Top             =   2070
            Width           =   435
         End
         Begin VB.TextBox txtECGPrice 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2475
            TabIndex        =   11
            Top             =   465
            Width           =   1425
         End
         Begin VB.TextBox txtECGCost 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2475
            TabIndex        =   10
            Top             =   825
            Width           =   1425
         End
         Begin VB.TextBox txtECGPer 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2235
            TabIndex        =   9
            Top             =   1680
            Width           =   435
         End
         Begin VB.Label label1 
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
            TabIndex        =   19
            Top             =   510
            Width           =   2100
         End
         Begin VB.Label label1 
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
            TabIndex        =   18
            Top             =   795
            Width           =   1125
         End
         Begin VB.Label label1 
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
            TabIndex        =   17
            Top             =   1170
            Width           =   1215
         End
         Begin VB.Label lblIncome 
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
            Left            =   2475
            TabIndex        =   16
            Top             =   1185
            Width           =   1425
         End
         Begin VB.Label label1 
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
            TabIndex        =   15
            Top             =   1740
            Width           =   2625
         End
         Begin VB.Label lblECGConsaltant 
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
            Left            =   3000
            TabIndex        =   14
            Top             =   1680
            Width           =   900
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   1
            X1              =   285
            X2              =   3910
            Y1              =   1575
            Y2              =   1575
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   0
            X1              =   285
            X2              =   3910
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Label label1 
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
            TabIndex        =   13
            Top             =   2130
            Width           =   2610
         End
         Begin VB.Label lblECGHCS 
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
            Left            =   3000
            TabIndex        =   12
            Top             =   2070
            Width           =   900
         End
      End
   End
   Begin VB.ComboBox cmbSnos 
      Height          =   4275
      Left            =   7875
      Style           =   1  'Simple Combo
      TabIndex        =   5
      Top             =   210
      Width           =   1995
   End
   Begin VB.CheckBox chkHasSno 
      Caption         =   "Item Has Serial No's"
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
      Height          =   195
      Left            =   150
      TabIndex        =   4
      Tag             =   "False"
      Top             =   4605
      Width           =   2115
   End
   Begin VB.CheckBox chkInactive 
      Caption         =   "Item is Inactive"
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
      Height          =   195
      Left            =   150
      TabIndex        =   3
      Tag             =   "False"
      Top             =   4875
      Width           =   2115
   End
   Begin VB.Label label1 
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
      TabIndex        =   6
      Top             =   15
      Width           =   960
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   405
      Left            =   4395
      TabIndex        =   1
      Top             =   4710
      Width           =   1665
      Caption         =   " Save & Close"
      PicturePosition =   327683
      Size            =   "2937;714"
      Picture         =   "frmECG.frx":0070
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
      Height          =   405
      Left            =   2670
      TabIndex        =   0
      Top             =   4695
      Width           =   1665
      Caption         =   "  Save & New"
      PicturePosition =   327683
      Size            =   "2937;714"
      Accelerator     =   78
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
      Left            =   6120
      TabIndex        =   2
      Top             =   4710
      Width           =   1665
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "2937;714"
      Picture         =   "frmECG.frx":01CE
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
Attribute VB_Name = "frmECG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Add As Boolean, Edit As Boolean
Dim Loaded As Boolean
Dim AccType As String
Dim AccountNo As String
Dim SubAccNo As String

Private Sub FillCmbs()

    Call AddToCombo(cmbCatID, "CatID", "ItemCatagories", "Order By CATID")
    
    Call AddToCombo(cmbCat, "Catagory", "ItemCatagories", " Order By CATID")
    
    Call AddToCombo(cmbSaleAcc, "AccTitle", "VActiveAccounts", " Where Trim(AccNo) Like '" & SaleAcc & "%'  AND Not Parent Order By ACCTitle")
    
    Call AddToCombo(cmbPchAcc, "AccTitle", "VActiveAccounts", " Where Trim(AccNo) Like '" & PchAcc & "%' AND Not Parent Order By ACCTitle")
    
    Call AddToCombo(cmbSRTAcc, "AccTitle", "VActiveAccounts", " Where Trim(AccNo) Like '" & SRTACC & "%' AND Not Parent Order By ACCTitle")
    
    Call AddToCombo(cmbPRTAcc, "AccTitle", "VActiveAccounts", " Where Trim(AccNo) Like '" & PRTACC & "%' AND Not Parent Order By ACCTitle")
    
    Call AddToCombo(cmbAssetAcc, "AccTitle", "VActiveAccounts", " Where Trim(AccNo) Like '" & STOCKASSETACC & "%' AND Not Parent Order By ACCTitle")
    
    Call AddToCombo(cmbUnit, "Distinct Unit", "Items", "Order By Unit")
    
End Sub



Private Sub chkHasSno_Click()
   Me.Width = IIf(chkHasSno = vbChecked, 10000, 8000)
   cmbSnos.Visible = chkHasSno = vbChecked
   Me.Visible = False
   Me.Show 1
End Sub

Private Sub cmbAssetAcc_Validate(Cancel As Boolean)
   If Not cmbAssetAcc.MatchFound Then cmbAssetAcc = ""
End Sub

Private Sub cmbCatID_Change()
    If cmbCatID.MatchFound Then Call cmbCatID_Click Else If Not cmbCatID = "" Then cmbCatID.DropDown
End Sub
Private Sub cmbCatID_Click()
    If cmbCat.ListIndex <> cmbCatID.ListIndex Then cmbCat.ListIndex = cmbCatID.ListIndex
    If Add Then lblItemID = GetNextItemID(cmbCatID)
End Sub
Private Sub cmbCatID_Validate(Cancel As Boolean)
    If Not cmbCatID.MatchFound Then cmbCatID.ListIndex = cmbCat.ListIndex
End Sub
Private Sub cmbCat_Change()
    If cmbCat.MatchFound Then Call cmbCat_Click Else If cmbCat <> "" Then cmbCat.DropDown
End Sub
Private Sub cmbCat_Click()
    If cmbCatID.ListIndex <> cmbCat.ListIndex Then cmbCatID.ListIndex = cmbCat.ListIndex
End Sub
Private Sub cmbCat_Validate(Cancel As Boolean)
   If Not cmbCat.MatchFound Then cmbCat.ListIndex = cmbCatID.ListIndex
End Sub

Private Sub cmbITemType_Click()

cmbPchAcc.Enabled = (cmbITemType.ListIndex <> 1)
cmbPRTAcc.Enabled = (cmbITemType.ListIndex = 0)
cmbSRTAcc.Enabled = (cmbITemType.ListIndex = 0)
cmbAssetAcc.Enabled = (cmbITemType.ListIndex <> 1)
TBox(3).Enabled = (cmbITemType.ListIndex <> 1)
TBox(4).Enabled = (cmbITemType.ListIndex <> 1)
TBox(5).Enabled = (cmbITemType.ListIndex <> 1)

chkHasSno.Visible = (cmbITemType.ListIndex = 0)

If cmbITemType.ListIndex <> 0 And chkHasSno = vbChecked Then chkHasSno = vbUnchecked

End Sub

Private Sub cmbsaleAcc_Validate(Cancel As Boolean)
    If Not cmbSaleAcc.MatchFound Then cmbSaleAcc = ""
End Sub

Private Sub cmbPchAcc_Validate(Cancel As Boolean)
    If Not cmbPchAcc.MatchFound Then cmbPchAcc = ""
End Sub



Private Sub cmbSnos_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cmbSnos_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And cmbSnos <> "" Then
   If FindInCombo(cmbSnos, cmbSnos.Text, True) Then
      MsgBox "This Serial No Already Added...", vbInformation
         cmbSnos.SelStart = 0
         cmbSnos.SelLength = Len(cmbSnos)
   Else
      cmbSnos.AddItem cmbSnos.Text
      cmbSnos = ""
   End If
   
ElseIf KeyCode = 46 Then
   If FindInCombo(cmbSnos, cmbSnos.Text, True) Then
      If MsgBox("Remove Sno : " & cmbSnos, vbQuestion + vbYesNo) = vbNo Then Exit Sub
      cmbSnos.RemoveItem cmbSnos.ListIndex
   End If
End If
End Sub

Private Sub cmbSnos_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If SaveItem Then
    Unload Me
    Me.Add = True
    Me.Edit = False
    Me.Show 1
    End If
End Sub

Private Sub cmdOK_Click()
    If SaveItem Then Unload Me
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub Form_Activate()
   If Edit And Not Loaded Then Call showdata
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
   Me.KeyPreview = True
'    DTPicker1.value = Date
    Call FillCmbs
End Sub

Private Function SaveItem() As Boolean

On Error GoTo err

Dim OpBalance As Double
Dim rec As New ADODB.Recordset

If Not cmbCatID.MatchFound Then
    MsgBox "Invalid Item Catagory", vbInformation
    SaveItem = False
    Exit Function
ElseIf Trim(TBox(0)) = "" Then
    MsgBox "Invalid Item Name", vbInformation
    SaveItem = False
    Exit Function
ElseIf Trim(cmbUnit) = "" And cmbPchAcc.Enabled Then
    MsgBox "No Unit Selected.", vbInformation
    SaveItem = False
    Exit Function
ElseIf cmbPchAcc.Enabled And Not cmbPchAcc.MatchFound Then
    MsgBox "No Purchase Account Selected", vbInformation
    SaveItem = False
    Exit Function
ElseIf cmbAssetAcc.Enabled And Not cmbAssetAcc.MatchFound Then
    MsgBox "No Asset Account Selected", vbInformation
    SaveItem = False
    Exit Function
ElseIf cmbSaleAcc.Enabled And Not cmbSaleAcc.MatchFound Then
    MsgBox "No Sales Account Selected", vbInformation
    SaveItem = False
    Exit Function
ElseIf cmbSRTAcc.Enabled And Not cmbSRTAcc.MatchFound Then
    MsgBox "No Sale Return Account Selected", vbInformation
    SaveItem = False
    Exit Function
ElseIf cmbPRTAcc.Enabled And Not cmbPRTAcc.MatchFound Then
    MsgBox "No Purchase Return Account Selected", vbInformation
    SaveItem = False
    Exit Function
ElseIf TBox(4).Enabled And Val(TBox(4)) = 0 Then
    MsgBox "Invalid Purchase Price", vbInformation
    SaveItem = False
    Exit Function
ElseIf TBox(5).Enabled And Val(TBox(5)) = 0 Then
    MsgBox "Invalid Sale Price", vbInformation
    SaveItem = False
    Exit Function
ElseIf Val(TBox(4)) >= Val(TBox(5)) And Val(TBox(4)) > 0 Then
    If MsgBox("Sale Price Should Be More Than Purchase Price." & vbNewLine & "Proceed Anyway ?", vbInformation + vbYesNo) = vbNo Then
        SaveItem = False
        Exit Function
    End If
ElseIf chkHasSno = vbChecked Then
   If cmbSnos.ListCount <> Val(TBox(2)) Then
      MsgBox "List Of Serial No.'s is Not Equal To The Quantity.", vbInformation
      saveitems = False
      Exit Function
   End If
End If


Dim SaleAcc As String, PchAcc As String, ItmPRTAcc As String, itmSRTAcc As String, AssetAcc As String

   con.BeginTrans


If Add Then

If cmbITemType.ListIndex <> 1 Then

   Dim TotalCost As Double

   AssetAcc = getAccountNo(cmbAssetAcc)      'Item's Asset Account
   
   PchAcc = getAccountNo(cmbPchAcc)          'Item's Purchase Account
   
   If cmbITemType.ListIndex = 0 Then
      ItmPRTAcc = getAccountNo(cmbPRTAcc)       'Item's Purchase Return Account
      itmSRTAcc = getAccountNo(cmbSRTAcc)       'Item's Sale Return Account
   End If
   
   TotalCost = Val(TBox(4)) * Val(TBox(2))
   
   con.Execute "Update Accounts Set OpenBal=OpenBal+" & TotalCost & " Where AccNo='" & AssetAcc & "'"
   Call adjustLedger(AssetAcc, "1/1/1900", 0, TotalCost)

End If

   SaleAcc = getAccountNo(cmbSaleAcc)        'Item's Sale Account

   
   ''''  Save Items
   
   con.Execute "Insert into Items(ItemID,ItemName,CatID,Unit,Description,AssetAccNo,SaleAccNo,SRTAccNo,PchAccno,PRTAccNo,SalePrice,PchPrice,OpenStock,InHand,Reorder,Active,HAsSno,ItemType) " & _
        "Values('" & lblItemID & "','" & TBox(0).Text & "','" & cmbCatID & "','" & cmbUnit & "','" & TBox(1) & "','" & AssetAcc & "','" & SaleAcc & "','" & itmSRTAcc & "','" & _
        PchAcc & "','" & ItmPRTAcc & "'," & Val(TBox(5)) & "," & Val(TBox(4)) & "," & Val(TBox(2)) & "," & Val(TBox(2)) & "," & Val(TBox(3)) & "," & CBool(chkInactive = vbUnchecked) & "," & CBool(chkHasSno = vbChecked) & "," & cmbITemType.ListIndex & ")"
        
Else

   con.Execute "Update ITems Set ItemName='" & TBox(0) & "',Unit='" & cmbUnit & "',Description='" & TBox(1) & "',PchPrice=" & Val(TBox(4)) & ",SalePrice=" & Val(TBox(5)) & ",Reorder =" & Val(TBox(3)) & ",Active=" & CBool(chkInactive <> vbChecked) & ",HasSno=" & CBool(chkHasSno = vbChecked) & " Where ItemID='" & Me.Tag & "'"
   
End If

   ''''  Save Serial Numbers if Item Has Serial No's Applied
   If chkHasSno = vbChecked Then
         con.Execute "Delete From ItemSNo where ItemID='" & lblItemID & "' and not out "
         For i = 0 To cmbSnos.ListCount - 1
            con.Execute "Insert into ItemSno(ITemID,Sno) Values('" & lblItemID & "','" & cmbSnos.List(i) & "')"
         Next
   End If
con.CommitTrans

 SaveItem = True
Exit Function

err:
    MsgBox err.Description
End Function




Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Tbox_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    Select Case Index
    Case 2, 3
        KeyAscii = OnlyNumber(Val(KeyAscii))
    Case 4, 5
        KeyAscii = OnlyNumber(Val(KeyAscii), True)
    End Select
End Sub



Private Function GetNextItemID(CatID As String) As String
Dim NextVal As String
 NextVal = GetMax("val(Right(ItemID,2))", "Items", "Where CatID ='" & CatID & "'")
 GetNextItemID = CatID & "-" & Format(Val(NextVal) + 1, "000")
End Function

Private Sub Tbox_Validate(Index As Integer, Cancel As Boolean)
    If Index = 4 Or Index = 5 Then
        TBox(Index) = Format(Val(TBox(Index)), "#0.00")
    ElseIf Index = 2 Or Index = 3 Then
        TBox(Index) = Format(Val(TBox(Index)), "#0")
    End If
End Sub


Private Sub showdata()
Dim rs As New ADODB.Recordset


With rs
    .Open "Select * From VItems Where ITemID='" & Me.Tag & "'", con, adOpenForwardOnly, adLockReadOnly
    lblItemID = ![ItemID] & ""
    cmbCatID = ![CatID] & ""
    TBox(0) = ![ItemName] & ""
    TBox(1) = ![Description] & ""
    TBox(2) = Val(![inhand] & "")
    TBox(3) = Val(![Reorder] & "")
    TBox(4) = Format(Val(![PchPRice] & ""), "#0.00")
    TBox(5) = Format(Val(![SalePrice] & ""), "#0.00")
    
    Call FindInCombo(cmbSaleAcc, ![SaleAccount] & "")
    
    If (![PurchaseAccount] & "") <> "" Then Call FindInCombo(cmbPchAcc, ![PurchaseAccount] & "") Else cmbPchAcc = ""
    
    'If ![NonInventory] Then cmbITemType.ListIndex = 1 Else cmbITemType.ListIndex = 0
    cmbITemType.ListIndex = Val(![ItemType] & "")
    
    If (![SRTAccount] & "") <> "" Then Call FindInCombo(cmbSRTAcc, ![SRTAccount] & "") Else cmbSrtAccno = ""
    
    If (![PRTAccount] & "") <> "" Then Call FindInCombo(cmbPRTAcc, ![PRTAccount] & "") Else cmbprtaccno = ""
    
    
    If (![AssetAccount] & "") <> "" Then Call FindInCombo(cmbAssetAcc, ![AssetAccount] & "") Else cmbassetaccno = ""
    
    Call FindInCombo(cmbUnit, ![Unit] & "")
    chkInactive = IIf(![Active], vbUnchecked, vbChecked)
    chkHasSno = IIf(![HasSno], vbChecked, vbUnchecked)
    
    '
    If ![HasSno] Then Call AddToCombo(cmbSnos, "Sno", "ItemSno", "Where ITemID='" & lblItemID & "' and not out")
    
    .Close
End With

Set rs = Nothing

fCat.Enabled = False

cmbSaleAcc.Enabled = False
cmbPchAcc.Enabled = False

cmbSRTAcc.Enabled = False
cmbPRTAcc.Enabled = False

cmbAssetAcc.Enabled = False
cmbITemType.Enabled = False
TBox(2).Enabled = False
Loaded = True

End Sub
