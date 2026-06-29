VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPercentages 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Income Percentages"
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
      TabIndex        =   24
      Top             =   -15
      Width           =   5205
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Income Percentages"
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
         Left            =   1320
         TabIndex        =   25
         Top             =   150
         Width           =   2715
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Income Percentages"
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
         Left            =   1335
         TabIndex        =   26
         Top             =   165
         Width           =   2715
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
      TabHeight       =   520
      TabCaption(0)   =   "OPD"
      TabPicture(0)   =   "frmPercentages.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Private Rooms"
      TabPicture(1)   =   "frmPercentages.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(7)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(16)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtPR(4)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtPR(5)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "CCU"
      TabPicture(2)   =   "frmPercentages.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   1755
         Left            =   -74745
         TabIndex        =   34
         Top             =   600
         Width           =   4455
         Begin VB.TextBox txtCCU 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   5
            Left            =   3090
            TabIndex        =   53
            Top             =   1035
            Width           =   900
         End
         Begin VB.TextBox txtCCU 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   4
            Left            =   3090
            TabIndex        =   52
            Top             =   690
            Width           =   900
         End
         Begin VB.TextBox txtCCU 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   3
            Left            =   2265
            TabIndex        =   38
            Top             =   1365
            Width           =   480
         End
         Begin VB.TextBox txtCCU 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   0
            Left            =   2760
            TabIndex        =   37
            Top             =   180
            Width           =   1215
         End
         Begin VB.TextBox txtCCU 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   2265
            TabIndex        =   36
            Top             =   1035
            Width           =   480
         End
         Begin VB.TextBox txtCCU 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   1
            Left            =   2265
            TabIndex        =   35
            Top             =   690
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total CCU Charges  Rs.    :"
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
            Left            =   570
            TabIndex        =   45
            Top             =   210
            Width           =   2115
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Alkhidmat                  @           % :"
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
            Left            =   435
            TabIndex        =   44
            Top             =   1410
            Width           =   2610
         End
         Begin VB.Label lblCCU 
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
            Height          =   120
            Index           =   1
            Left            =   90
            TabIndex        =   43
            Top             =   570
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   5
            X1              =   360
            X2              =   3985
            Y1              =   585
            Y2              =   585
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   4
            X1              =   360
            X2              =   3985
            Y1              =   570
            Y2              =   570
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
            Index           =   11
            Left            =   435
            TabIndex        =   42
            Top             =   1065
            Width           =   2610
         End
         Begin VB.Label lblCCU 
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
            Left            =   3090
            TabIndex        =   41
            Top             =   1365
            Width           =   900
         End
         Begin VB.Label lblCCU 
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
            Height          =   135
            Index           =   0
            Left            =   90
            TabIndex        =   40
            Top             =   435
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
            Index           =   10
            Left            =   435
            TabIndex        =   39
            Top             =   735
            Width           =   2625
         End
      End
      Begin VB.TextBox txtPR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   5
         Left            =   -71880
         TabIndex        =   32
         Top             =   2595
         Width           =   1215
      End
      Begin VB.TextBox txtPR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   4
         Left            =   -71880
         TabIndex        =   30
         Top             =   2190
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Height          =   1755
         Left            =   -74640
         TabIndex        =   15
         Top             =   375
         Width           =   4455
         Begin VB.TextBox txtPR 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   7
            Left            =   3075
            TabIndex        =   51
            Top             =   1020
            Width           =   900
         End
         Begin VB.TextBox txtPR 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   6
            Left            =   3075
            TabIndex        =   50
            Top             =   675
            Width           =   900
         End
         Begin VB.TextBox txtPR 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   1
            Left            =   2250
            TabIndex        =   27
            Top             =   675
            Width           =   525
         End
         Begin VB.TextBox txtPR 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   2
            Left            =   2250
            TabIndex        =   18
            Top             =   1020
            Width           =   525
         End
         Begin VB.TextBox txtPR 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   0
            Left            =   2760
            TabIndex        =   17
            Top             =   195
            Width           =   1215
         End
         Begin VB.TextBox txtPR 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   3
            Left            =   2250
            TabIndex        =   16
            Top             =   1365
            Width           =   525
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
            Index           =   6
            Left            =   435
            TabIndex        =   29
            Top             =   735
            Width           =   2625
         End
         Begin VB.Label lblPR 
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
            Height          =   270
            Index           =   0
            Left            =   195
            TabIndex        =   28
            Top             =   1020
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label lblPR 
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
            Left            =   3075
            TabIndex        =   23
            Top             =   1365
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
            Left            =   435
            TabIndex        =   22
            Top             =   1065
            Width           =   2610
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   3
            X1              =   360
            X2              =   3985
            Y1              =   570
            Y2              =   570
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   2
            X1              =   360
            X2              =   3985
            Y1              =   585
            Y2              =   585
         End
         Begin VB.Label lblPR 
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
            Height          =   270
            Index           =   1
            Left            =   195
            TabIndex        =   21
            Top             =   1275
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Alkhidmat                  @           % :"
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
            Left            =   435
            TabIndex        =   20
            Top             =   1395
            Width           =   2610
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total Pvt. Room Chgs  Rs.    :"
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
            Left            =   345
            TabIndex        =   19
            Top             =   225
            Width           =   2370
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2640
         Left            =   360
         TabIndex        =   4
         Top             =   465
         Width           =   4455
         Begin VB.TextBox txtOPD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   6
            Left            =   3015
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   2175
            Width           =   900
         End
         Begin VB.TextBox txtOPD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   5
            Left            =   3015
            TabIndex        =   48
            Top             =   1785
            Width           =   900
         End
         Begin VB.TextBox txtOPD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   4
            Left            =   2580
            TabIndex        =   46
            Top             =   225
            Width           =   1350
         End
         Begin VB.TextBox txtOPD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   3
            Left            =   2205
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   2175
            Width           =   465
         End
         Begin VB.TextBox txtOPD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   0
            Left            =   2580
            TabIndex        =   7
            Top             =   570
            Width           =   1350
         End
         Begin VB.TextBox txtOPD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   1
            Left            =   2580
            TabIndex        =   6
            Top             =   930
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.TextBox txtOPD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   2
            Left            =   2205
            TabIndex        =   5
            Top             =   1785
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "OPD Consultancy Chgs Rs. :"
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
            Left            =   300
            TabIndex        =   47
            Top             =   270
            Width           =   2265
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total OPD ECG Chgs  Rs.     :"
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
            Left            =   315
            TabIndex        =   13
            Top             =   600
            Width           =   2250
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "OPD Cost Rs. :"
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
            Left            =   330
            TabIndex        =   12
            Top             =   960
            Visible         =   0   'False
            Width           =   1155
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
            Left            =   315
            TabIndex        =   11
            Top             =   1305
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblOPD 
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
            Left            =   2580
            TabIndex        =   10
            Top             =   1305
            Visible         =   0   'False
            Width           =   1350
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
            TabIndex        =   9
            Top             =   1830
            Width           =   2625
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
            TabIndex        =   8
            Top             =   2220
            Width           =   2610
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Air Conditioner Charges  Rs. :"
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
         Left            =   -74370
         TabIndex        =   33
         Top             =   2640
         Width           =   2430
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Room Heater Charges  Rs.    :"
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
         Left            =   -74385
         TabIndex        =   31
         Top             =   2235
         Width           =   2430
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
      Picture         =   "frmPercentages.frx":0054
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
      Picture         =   "frmPercentages.frx":01B2
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
Attribute VB_Name = "frmPercentages"
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
        
        '''''''''''' OPD ''''''''''''''''''
        .MoveFirst
        .Find "IncomeSource='OPD'"
        txtOPD(0) = Format(Val(![Price] & ""), "0.00")
        txtOPD(1) = Format(Val(![Cost] & ""), "0.00")
        txtOPD(2) = Format(Val(![Consaltant] & ""), "0.00")
        txtOPD(3) = Format(Val(![HCS] & ""), "0.00")
        txtOPD(4) = Format(Val(![Price1] & ""), "0.00")
        
        
        '''''''''''' PR  ''''''''''''''''''
        
        .MoveFirst
        .Find "IncomeSource='PR'"
        txtPR(0) = Format(Val(![Price] & ""), "0.00")
        txtPR(1) = Format(Val(![Consaltant] & ""), "0.00")
        txtPR(2) = Format(Val(![HCS] & ""), "0.00")
        txtPR(3) = Format(Val(![Alkhidmat] & ""), "0.00")
        txtPR(4) = Format(Val(![Heater] & ""), "0.00")
        txtPR(5) = Format(Val(![AC] & ""), "0.00")
        
        '''''''''''' CCU  ''''''''''''''''''
        
        .MoveFirst
        .Find "IncomeSource='CCU'"
        txtCCU(0) = Format(Val(![Price] & ""), "0.00")
        'txtCCU(1) = Format(Val(![Cost] & ""), "0.00")
        txtCCU(1) = Format(Val(![Consaltant] & ""), "0.00")
        txtCCU(2) = Format(Val(![HCS] & ""), "0.00")
        
        
        
        
      
        .Close
    End With
    Set rs = Nothing
End Sub

Private Function Saved() As Boolean

On Error GoTo err
   
     con.BeginTrans
     
     con.Execute "Update THTCPercentages Set Price=" & Val(txtOPD(0)) & ",Cost=" & Val(txtOPD(1)) & _
     ",Consaltant=" & Val(txtOPD(2)) & ",HCS=" & Val(txtOPD(3)) & " Where IncomeSource='OPD'"
     
     con.Execute "Update THTCPercentages Set Price=" & Val(txtPR(0)) & _
     ",Consaltant=" & Val(txtPR(1)) & ",HCS=" & Val(txtPR(2)) & _
     ",Alkhidmat=" & Val(txtPR(3)) & ",Heater=" & Val(txtPR(4)) & _
     ",AC=" & Val(txtPR(5)) & " Where IncomeSource='PR'"
     
     con.Execute "Update THTCPercentages Set Price=" & Val(txtCCU(0)) & _
     ",Consaltant=" & Val(txtCCU(1)) & ",HCS=" & Val(txtCCU(2)) & _
     ",Alkhidmat=" & Val(txtCCU(3)) & " Where IncomeSource='CCU'"
     
    
     con.CommitTrans

     Saved = True
     
Exit Function

err:
    MsgBox err.Description
End Function



Private Sub txtCCU_LostFocus(Index As Integer)
    txtCCU(1).Text = Format(Val(txtCCU(4)) / Val(txtCCU(0)) * 100, "0.00")
    txtCCU(2).Text = Format(Val(txtCCU(5)) / Val(txtCCU(0)) * 100, "0.00")
End Sub

Private Sub txtOPD_Change(Index As Integer)
    
    If Index <> 5 Then
        lblOPD(0) = Format(Val(txtOPD(0)) - Val(txtOPD(1)), "0.00 ")
        txtOPD(5) = Format(Val(lblOPD(0)) * Val(txtOPD(2)) / 100, "0.00 ")
        txtOPD(3) = Format(100 - Val(txtOPD(2)), "0.00")
        txtOPD(6) = Format(Val(lblOPD(0)) * Val(txtOPD(3)) / 100, "0.00 ")
    End If
    
End Sub

Private Sub txtOPD_GotFocus(Index As Integer)
    txtOPD(Index).SelStart = 0
    txtOPD(Index).SelLength = Len(txtOPD(Index))
End Sub

Private Sub txtOPD_LostFocus(Index As Integer)
    txtOPD(2) = Format(Val(txtOPD(5)) / Val(lblOPD(0)) * 100, "0.00")
End Sub

Private Sub txtPR_GotFocus(Index As Integer)
    txtPR(Index).SelStart = 0
    txtPR(Index).SelLength = Len(txtPR(Index))
End Sub

Private Sub txtCCU_GotFocus(Index As Integer)
    txtCCU(Index).SelStart = 0
    txtCCU(Index).SelLength = Len(txtCCU(Index))
End Sub

Private Sub txtPR_Change(Index As Integer)
    If Index <> 6 And Index <> 7 Then
        txtPR(6) = Format(Val(txtPR(0)) * Val(txtPR(1)) / 100, "0.00 ")
        txtPR(7) = Format(Val(txtPR(0)) * Val(txtPR(2)) / 100, "0.00 ")
        txtPR(3) = Format(100 - Val(txtPR(2)) - Val(txtPR(1)), "0.00")
        lblPR(2) = Format(Val(txtPR(0)) * Val(txtPR(3)) / 100, "0.00 ")
    End If
End Sub

Private Sub txtCCU_Change(Index As Integer)
    If Index <> 4 And Index <> 5 Then
        txtCCU(4) = Format(Val(txtCCU(0)) * Val(txtCCU(1)) / 100, "0.00 ")
        txtCCU(5) = Format(Val(txtCCU(0)) * Val(txtCCU(2)) / 100, "0.00 ")
    End If
    
    txtCCU(3) = Format(100 - Val(txtCCU(2)) - Val(txtCCU(1)), "0.00")
    lblCCU(2) = Format(Val(txtCCU(0)) * Val(txtCCU(3)) / 100, "0.00 ")
    
    
End Sub




Private Sub txtPR_LostFocus(Index As Integer)
    txtPR(1).Text = Format(Val(txtPR(6)) / Val(txtPR(0)) * 100, "0.00")
    txtPR(2).Text = Format(Val(txtPR(7)) / Val(txtPR(0)) * 100, "0.00")
End Sub


