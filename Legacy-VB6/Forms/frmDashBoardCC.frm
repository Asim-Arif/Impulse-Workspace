VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmDashBoardCC 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Delete Voucher"
   ClientHeight    =   10065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15105
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDashBoardCC.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10065
   ScaleWidth      =   15105
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00005601&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005601&
      Height          =   4485
      Index           =   4
      Left            =   5130
      ScaleHeight     =   4485
      ScaleWidth      =   9870
      TabIndex        =   71
      TabStop         =   0   'False
      Tag             =   "PRODUCTION"
      Top             =   5010
      Width           =   9870
      Begin VB.Frame FLoading 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3705
         Index           =   4
         Left            =   480
         TabIndex        =   72
         Top             =   360
         Width           =   9210
         Begin MSComCtl2.Animation Animation 
            Height          =   2205
            Index           =   4
            Left            =   3015
            TabIndex        =   73
            Top             =   765
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   3889
            _Version        =   393216
            Center          =   -1  'True
            BackColor       =   12632256
            FullWidth       =   220
            FullHeight      =   147
         End
      End
      Begin VB.TextBox txtLotsAtPolishingQty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   450
         Left            =   6570
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   2568
         Width           =   3120
      End
      Begin VB.TextBox txtLotsAtTemperQty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   450
         Left            =   6570
         Locked          =   -1  'True
         TabIndex        =   86
         Top             =   2076
         Width           =   3120
      End
      Begin VB.TextBox txtRunningLotsQty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   450
         Left            =   6570
         Locked          =   -1  'True
         TabIndex        =   85
         Top             =   1584
         Width           =   3120
      End
      Begin VB.TextBox txtMakerOrdersQty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   450
         Left            =   6570
         Locked          =   -1  'True
         TabIndex        =   84
         Top             =   1092
         Width           =   3120
      End
      Begin VB.TextBox txtForgingOrdersQty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   450
         Left            =   6570
         Locked          =   -1  'True
         TabIndex        =   83
         Top             =   600
         Width           =   3120
      End
      Begin VB.TextBox txtLotsAtQCQty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   450
         Left            =   6570
         TabIndex        =   82
         Top             =   3060
         Width           =   3120
      End
      Begin VB.TextBox txtLotsAtPackedQty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   450
         Left            =   6570
         TabIndex        =   81
         Top             =   3555
         Width           =   3120
      End
      Begin VB.TextBox txtLotsAtPackedNos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   450
         Left            =   3450
         TabIndex        =   80
         Top             =   3555
         Width           =   3120
      End
      Begin VB.TextBox txtLotsAtQCNos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   450
         Left            =   3450
         TabIndex        =   79
         Top             =   3060
         Width           =   3120
      End
      Begin VB.TextBox txtForgingOrdersNos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   450
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   78
         Top             =   600
         Width           =   3120
      End
      Begin VB.TextBox txtMakerOrdersNos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   450
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   77
         Top             =   1092
         Width           =   3120
      End
      Begin VB.TextBox txtRunningLotsNos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   450
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   76
         Top             =   1584
         Width           =   3120
      End
      Begin VB.TextBox txtLotsAtTemperNos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   450
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   2076
         Width           =   3120
      End
      Begin VB.TextBox txtLotsAtPolishingNos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   450
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   2568
         Width           =   3120
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lots at Polishing :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   24
         Left            =   1155
         MouseIcon       =   "frmDashBoardCC.frx":014A
         MousePointer    =   99  'Custom
         TabIndex        =   96
         Top             =   2580
         Width           =   2025
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lots at Temper :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   25
         Left            =   1290
         MouseIcon       =   "frmDashBoardCC.frx":029C
         MousePointer    =   99  'Custom
         TabIndex        =   95
         Top             =   2115
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Running Lots :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   26
         Left            =   1560
         MouseIcon       =   "frmDashBoardCC.frx":03EE
         MousePointer    =   99  'Custom
         TabIndex        =   94
         Top             =   1635
         Width           =   1620
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Open Maker Orders :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   27
         Left            =   855
         MouseIcon       =   "frmDashBoardCC.frx":0540
         MousePointer    =   99  'Custom
         TabIndex        =   93
         Top             =   1140
         Width           =   2325
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Open Forging Orders :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   28
         Left            =   675
         MouseIcon       =   "frmDashBoardCC.frx":0692
         MousePointer    =   99  'Custom
         TabIndex        =   92
         Top             =   630
         Width           =   2505
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lots at QC 100% :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   29
         Left            =   1200
         MouseIcon       =   "frmDashBoardCC.frx":07E4
         MousePointer    =   99  'Custom
         TabIndex        =   91
         Top             =   3090
         Width           =   1980
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lots at Packed today :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   13
         Left            =   645
         MouseIcon       =   "frmDashBoardCC.frx":0936
         MousePointer    =   99  'Custom
         TabIndex        =   90
         Top             =   3585
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QTY"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   30
         Left            =   7905
         TabIndex        =   89
         Top             =   270
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nos."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   31
         Left            =   4920
         TabIndex        =   88
         Top             =   270
         Width           =   510
      End
   End
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00005601&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005601&
      Height          =   4305
      Index           =   3
      Left            =   10110
      ScaleHeight     =   4305
      ScaleWidth      =   4875
      TabIndex        =   32
      TabStop         =   0   'False
      Tag             =   "FINANCIALS"
      Top             =   570
      Width           =   4875
      Begin VB.Frame FLoading 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3780
         Index           =   3
         Left            =   120
         TabIndex        =   66
         Top             =   345
         Width           =   4605
         Begin MSComCtl2.Animation Animation 
            Height          =   2205
            Index           =   3
            Left            =   675
            TabIndex        =   67
            Top             =   765
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   3889
            _Version        =   393216
            Center          =   -1  'True
            BackColor       =   12632256
            FullWidth       =   220
            FullHeight      =   147
         End
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cash Book Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Index           =   1
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   3435
         Width           =   2205
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Six Months Expense"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Index           =   0
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   3435
         Width           =   2205
      End
      Begin VB.TextBox txtFixedAssets 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   3135
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   2715
         Width           =   1590
      End
      Begin VB.TextBox txtAccRcv 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   3135
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   2190
         Width           =   1590
      End
      Begin VB.TextBox txtAccPay 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   3135
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1650
         Width           =   1590
      End
      Begin VB.TextBox txtBankBal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   3135
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1125
         Width           =   1590
      End
      Begin VB.TextBox txtCashAcc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   3135
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   600
         Width           =   1590
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fixed Assets :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   22
         Left            =   1530
         TabIndex        =   42
         Top             =   2775
         Width           =   1545
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Accounts Receivable :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   21
         Left            =   570
         MouseIcon       =   "frmDashBoardCC.frx":0A88
         TabIndex        =   41
         Top             =   2250
         Width           =   2505
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Accounts Payable :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   20
         Left            =   930
         MouseIcon       =   "frmDashBoardCC.frx":0BDA
         TabIndex        =   40
         Top             =   1680
         Width           =   2145
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Balances :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   19
         Left            =   1290
         MouseIcon       =   "frmDashBoardCC.frx":0D2C
         MousePointer    =   99  'Custom
         TabIndex        =   39
         Top             =   1140
         Width           =   1785
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Account :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   18
         Left            =   1365
         MouseIcon       =   "frmDashBoardCC.frx":0E7E
         MousePointer    =   99  'Custom
         TabIndex        =   38
         Top             =   615
         Width           =   1710
      End
   End
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00005601&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005601&
      Height          =   4485
      Index           =   2
      Left            =   180
      ScaleHeight     =   4485
      ScaleWidth      =   4875
      TabIndex        =   23
      TabStop         =   0   'False
      Tag             =   "MERCHANDISNG & EXPORT"
      Top             =   5010
      Width           =   4875
      Begin VB.Frame FLoading 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3705
         Index           =   2
         Left            =   90
         TabIndex        =   64
         Top             =   345
         Width           =   4650
         Begin MSComCtl2.Animation Animation 
            Height          =   2205
            Index           =   2
            Left            =   675
            TabIndex        =   65
            Top             =   765
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   3889
            _Version        =   393216
            Center          =   -1  'True
            BackColor       =   12632256
            FullWidth       =   220
            FullHeight      =   147
         End
      End
      Begin VB.TextBox txtTotalSalesFY 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   2814
         Width           =   1590
      End
      Begin VB.TextBox txtTotalSales 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   3135
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   2076
         Width           =   1590
      End
      Begin VB.TextBox txtInHandValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   3135
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1338
         Width           =   1590
      End
      Begin VB.TextBox txtInHandQty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   3135
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   600
         Width           =   1590
      End
      Begin VB.TextBox txtRebateRcvable 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   3135
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   3555
         Width           =   1590
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Sales of Current FY :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   12
         Left            =   75
         MouseIcon       =   "frmDashBoardCC.frx":0FD0
         MousePointer    =   99  'Custom
         TabIndex        =   52
         Top             =   2865
         Width           =   2985
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Sales :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   16
         Left            =   1695
         MouseIcon       =   "frmDashBoardCC.frx":1122
         TabIndex        =   31
         Top             =   2085
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Orders In Hand - Value :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   15
         Left            =   390
         MouseIcon       =   "frmDashBoardCC.frx":1274
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   1365
         Width           =   2685
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Orders In Hand - Qty :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   14
         Left            =   630
         MouseIcon       =   "frmDashBoardCC.frx":13C6
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   615
         Width           =   2445
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rebate Receivable :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   11
         Left            =   840
         TabIndex        =   28
         Top             =   3585
         Width           =   2235
      End
   End
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00005601&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005601&
      Height          =   4305
      Index           =   1
      Left            =   180
      ScaleHeight     =   4305
      ScaleWidth      =   4875
      TabIndex        =   10
      TabStop         =   0   'False
      Tag             =   "HR & PAYROLL"
      Top             =   570
      Width           =   4875
      Begin VB.Frame FLoading 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3675
         Index           =   1
         Left            =   150
         TabIndex        =   60
         Top             =   345
         Width           =   4590
         Begin MSComCtl2.Animation Animation 
            Height          =   2205
            Index           =   1
            Left            =   675
            TabIndex        =   61
            Top             =   765
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   3889
            _Version        =   393216
            Center          =   -1  'True
            BackColor       =   12632256
            FullWidth       =   220
            FullHeight      =   147
         End
      End
      Begin VB.TextBox txtShortTerm 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   3135
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   3255
         Width           =   1590
      End
      Begin VB.TextBox txtNoOfEmployees 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   3135
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   600
         Width           =   1590
      End
      Begin VB.TextBox txtPresentEmployees 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   3135
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1125
         Width           =   1590
      End
      Begin VB.TextBox txtAbsentEmployees 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   3135
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1650
         Width           =   1590
      End
      Begin VB.TextBox txtAvgMonthlySalary 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   3135
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2190
         Width           =   1590
      End
      Begin VB.TextBox txtLongTerm 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   3135
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2715
         Width           =   1590
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Short Term Advances :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   10
         Left            =   495
         MouseIcon       =   "frmDashBoardCC.frx":1518
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   3285
         Width           =   2580
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total No. of Employees :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   9
         Left            =   270
         TabIndex        =   20
         Top             =   630
         Width           =   2805
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Present Employees :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   8
         Left            =   810
         TabIndex        =   19
         Top             =   1140
         Width           =   2265
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Absent Employees :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   7
         Left            =   885
         MouseIcon       =   "frmDashBoardCC.frx":166A
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   1680
         Width           =   2190
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Average Monthly Salary :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   5
         Left            =   285
         TabIndex        =   17
         Top             =   2250
         Width           =   2790
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Long Term Advances :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   4
         Left            =   540
         MouseIcon       =   "frmDashBoardCC.frx":17BC
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   2775
         Width           =   2535
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   255
      Top             =   1740
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13770
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9585
      Width           =   1215
   End
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00005601&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005601&
      Height          =   4305
      Index           =   0
      Left            =   5145
      ScaleHeight     =   4305
      ScaleWidth      =   4875
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "STOCK & PROCURMENT"
      Top             =   570
      Width           =   4875
      Begin VB.Frame FLoading 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3675
         Index           =   0
         Left            =   150
         TabIndex        =   62
         Top             =   345
         Width           =   4605
         Begin MSComCtl2.Animation Animation 
            Height          =   2205
            Index           =   0
            Left            =   675
            TabIndex        =   63
            Top             =   765
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   3889
            _Version        =   393216
            Center          =   -1  'True
            BackColor       =   12632256
            FullWidth       =   220
            FullHeight      =   147
         End
      End
      Begin VB.TextBox txtReworkLotsQty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   3345
         Width           =   1155
      End
      Begin VB.TextBox txtReWorkLots 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   3345
         Width           =   1155
      End
      Begin VB.TextBox txtFinishedStockValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   600
         Width           =   1155
      End
      Begin VB.TextBox txtSFStockValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   1286
         Width           =   1155
      End
      Begin VB.TextBox txtForgingStockValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   1972
         Width           =   1155
      End
      Begin VB.TextBox txtOtherInventoryValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   2658
         Width           =   1155
      End
      Begin VB.TextBox txtOtherInventoryQty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2658
         Width           =   1155
      End
      Begin VB.TextBox txtForgingStockQty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1972
         Width           =   1155
      End
      Begin VB.TextBox txtSFStockQty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1286
         Width           =   1155
      End
      Begin VB.TextBox txtFinishedStockQty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rework Lots :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   17
         Left            =   735
         MouseIcon       =   "frmDashBoardCC.frx":190E
         MousePointer    =   99  'Custom
         TabIndex        =   68
         Top             =   3390
         Width           =   1605
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VALUE"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   23
         Left            =   3705
         TabIndex        =   50
         Top             =   285
         Width           =   750
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QTY"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   3
         Left            =   2715
         TabIndex        =   49
         Top             =   270
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Other Inventory :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   2
         Left            =   435
         MouseIcon       =   "frmDashBoardCC.frx":1A60
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   2730
         Width           =   1920
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forging Stock :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   1
         Left            =   600
         MouseIcon       =   "frmDashBoardCC.frx":1BB2
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   1995
         Width           =   1740
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Semi Finish Stock :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   0
         Left            =   210
         MouseIcon       =   "frmDashBoardCC.frx":1D04
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   1305
         Width           =   2145
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Finished Stock :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   6
         Left            =   555
         MouseIcon       =   "frmDashBoardCC.frx":1E56
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   630
         Width           =   1800
      End
   End
   Begin VB.Frame FraSummaryOfStock 
      BackColor       =   &H00404040&
      Caption         =   "Summary of Stock"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   7425
      Left            =   135
      TabIndex        =   53
      Top             =   2100
      Visible         =   0   'False
      Width           =   14835
      Begin VB.CommandButton cmdHideSummaryOfStock 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Hide"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13470
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   6975
         Width           =   1215
      End
      Begin MSComctlLib.ListView LVSOS 
         Height          =   6555
         Left            =   165
         TabIndex        =   54
         Top             =   375
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   11562
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   16520707
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S #"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Group"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Qty"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Value"
            Object.Width           =   2469
         EndProperty
      End
   End
   Begin VB.Frame FraDOS 
      BackColor       =   &H00404040&
      Caption         =   "Detail of Stock :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   7425
      Left            =   165
      TabIndex        =   56
      Top             =   2115
      Visible         =   0   'False
      Width           =   14835
      Begin VB.CommandButton cmdHideDOS 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Hide"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13470
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   6975
         Width           =   1215
      End
      Begin MSComctlLib.ListView LVDOS 
         Height          =   6555
         Left            =   165
         TabIndex        =   58
         Top             =   375
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   11562
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   16520707
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S #"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item"
            Object.Width           =   17639
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Qty"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Value"
            Object.Width           =   2469
         EndProperty
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "COMMAND CENTER"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   135
      TabIndex        =   43
      Top             =   105
      Width           =   14850
   End
End
Attribute VB_Name = "frmDashBoardCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bMouseDown As Boolean
Const lCaptionHeight = 600
Dim lDownX As Long, lDownY As Long
Dim iColNo As Integer
Dim iColNoTemp As Integer
Dim bSaved As Boolean
Dim lEntryID As Long

Private Sub DrawBorder(Pic As PictureBox)

    Dim lBorderColor As Long
    lBorderColor = RGB(0, 200, 0)    'RGB(69, 140, 201) 'RGB(69, 140, 201)
    Pic.Line (0, 0)-(Pic.Width, 0), lBorderColor 'Top
    Pic.Line (0, Pic.Height - 85)-(Pic.Width, Pic.Height - 85), lBorderColor 'Bottom
    Pic.Line (0, 0)-(0, Pic.Height), lBorderColor    'Left
    Pic.Line (Pic.Width - 85, 0)-(Pic.Width - 85, Pic.Height), lBorderColor  'Right
    
End Sub

Private Sub DrawPicCaptions()

    Dim i As Integer
    Dim lBorderColor As Long, lForeColor As Long
    For i = 0 To PicWC.count - 1
'        If i = 1 Then
'            lBorderColor = RGB(153, 204, 255)
'            lForeColor = 0
'        ElseIf i = 0 Then
'            lBorderColor = RGB(255, 204, 153)
'            lForeColor = 0
'        ElseIf i = 3 Then
'            lBorderColor = RGB(255, 153, 204)
'            lForeColor = 0
'        ElseIf i = 2 Then
'            lBorderColor = RGB(0, 255, 0)
'            lForeColor = 0
'        ElseIf i = 4 Then
'            lBorderColor = RGB(200, 200, 192)
'            lForeColor = 0
'        End If
        lBorderColor = &HFFC0C0
        Call DrawCaptions(PicWC(i), lBorderColor, lForeColor)
    Next
    
End Sub

Private Sub DrawCaptions(Pic As PictureBox, lBorderColor As Long, lForeColor As Long)

    Dim lFillColor
    'lBorderColor = 0 'RGB(226, 103, 161) 'RGB(70, 70, 181)
    lFillColor = lBorderColor 'RGB(226, 103, 161) 'RGB(70, 70, 181)
    Pic.Line (0, 20)-(Pic.Width, 20), lBorderColor
    
    If Pic.Tag <> "" Then
        Pic.Line (0, 150)-(Pic.Width, 150), lBorderColor
        
        Pic.Line (0, 50)-(Pic.Width, 300), lFillColor, BF
        Pic.ForeColor = lForeColor 'vbGreen
        Pic.CurrentX = (Pic.Width - Pic.TextWidth(Pic.Tag)) / 2
        Pic.CurrentY = 0
        Pic.Print Pic.Tag
    End If
        
    'Pic.Line (0, Pic.Height - 30)-(Pic.Width, Pic.Height - 30), lBorderColor 'Bottom
    Pic.Line (0, Pic.Height - 50)-(Pic.Width, Pic.Height), lFillColor, BF
    
    Pic.Line (20, 0)-(20, Pic.Height), lBorderColor                          'Left
    Pic.Line (Pic.Width - 20, 0)-(Pic.Width - 30, Pic.Height), lBorderColor  'Right
    
    'Pic.Line (0, Pic.Height - 500)-(50, Pic.Height), lBorderColor
'    Me.BackColor = lBorderColor
End Sub


Private Sub cmdButton_Click(Index As Integer)
    If Index = 0 Then
        Load frmExpenseReport
        frmExpenseReport.ShowMe (True)
    ElseIf Index = 1 Then
        Call PrintCashBookReport(Date, Date)
        'frmPrintWOStatus.Show , MainForm
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub cmdHideDOS_Click()
    FraDOS.Visible = False
End Sub

Private Sub cmdHideSummaryOfStock_Click()
    FraSummaryOfStock.Visible = False
End Sub

Private Sub Form_Load()

    Call DrawPicCaptions
    bMouseDown = False
    bSaved = False
    lEntryID = 0
    
    Dim myCTL As Control
    For Each myCTL In Controls
        If TypeOf myCTL Is VB.TextBox Then
            myCTL.Text = "-"
        End If
    Next
     
    
    
End Sub

Private Sub label1_Click(Index As Integer)
    
   
    If Index = 7 Then
        Call PrintAbsentSheet
    ElseIf Index = 6 Then
        Call ShowSummaryOfStock(0)
    ElseIf Index = 0 Then
        Call ShowSummaryOfStock(1)
    ElseIf Index = 1 Then   'Forging
        Call ShowSummaryOfStock(2)
    ElseIf Index = 2 Then   'Other Inventory
        Call ShowSummaryOfStock(3)
    ElseIf Index = 4 Or Index = 10 Then
        Call PrintLoanBalance
    ElseIf Index = 14 Or Index = 15 Then
        Call PrintUnshippedOrderList
    ElseIf Index = 17 Then
        Call PrintReWorkLots
    ElseIf Index = 18 Or Index = 19 Then
        frmCashStatus.Show 1
    ElseIf Index = 23 Then
'        Call PrintPendingPurchaseOrdersList
    ElseIf Index = 20 Then
        'Call PrintVenderPaymentSheet(True)
    ElseIf Index = 21 Then
        'Call PrintReceivableAgingReport
    ElseIf Index = 12 Then
        Call PrintTotalExportCurrentFY
    ElseIf Index = 26 Then
        Call PrintLots
    ElseIf Index = 25 Then      'Lots @ Temper
        Call PrintLots(txtLotsAtTemperNos.Tag)
    ElseIf Index = 24 Then      'Lots @ Polishing
        Call PrintLots(txtLotsAtPolishingNos.Tag)
    ElseIf Index = 27 Then  'Open Maker Orders
        Call PrintOpenMakerOrders
    ElseIf Index = 28 Then  'Open Forging Orders
        Call PrintForgingOpenOrders
    ElseIf Index = 29 Then      'Lots @ QC 100%
        Call PrintLots(txtLotsAtQCNos.Tag)  'Lots @ Packed
    ElseIf Index = 13 Then
        Call PrintLots(txtLotsAtPackedNos.Tag)
    End If
    
End Sub

Private Sub ShowSummaryOfStock(Optional iType As Integer = 0)

    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        If iType = 0 Then
            .Open "SELECT Items.GroupID,ItemGroups.Description,SUM(NetQty*PriceForCost) AS TotalValue,SUM(NetQty) AS TotalStock FROM VItemsWithShelfWiseStock INNER JOIN Items ON VItemsWithShelfWiseStock.ItemID=Items.ItemID INNER JOIN ItemGroups ON Items.GroupID=ItemGroups.ID GROUP BY Items.GroupID,ItemGroups.Description", con, adOpenForwardOnly, adLockReadOnly
        ElseIf iType = 1 Then
            .Open "SELECT Items.GroupID,ItemGroups.Description,SUM((Qty-IssdQty)*PriceForCost) AS TotalValue,SUM((Qty-IssdQty)) AS TotalStock FROM VStockOrderOpening INNER JOIN (SELECT Opening_RefID FROM VendRcvdDetail GROUP BY Opening_RefID) T1 ON T1.Opening_RefID=VStockOrderOpening.EntryID INNER JOIN Items ON VStockOrderOpening.ItemID=Items.ItemID INNER JOIN ItemGroups ON Items.GroupID=ItemGroups.ID WHERE Qty>IssdQty GROUP BY Items.GroupID,ItemGroups.Description", con, adOpenForwardOnly, adLockReadOnly
        ElseIf iType = 2 Then   'Forging
            .Open "SELECT VRM.GroupID,VRM.GroupName AS Description,SUM(QtyInStock) AS TotalStock,SUM(QtyInStock*Rate) AS TotalValue FROM VRM WHERE GroupID IN(SELECT Group_ID FROM RMGroupIDsForForging) AND QtyInStock>0 GROUP BY VRM.GroupID,VRM.GroupName", con, adOpenForwardOnly, adLockReadOnly
        ElseIf iType = 3 Then   'Other Inventory
            .Open "SELECT VRM.GroupID,VRM.GroupName AS Description,SUM(QtyInStock) AS TotalStock,SUM(QtyInStock*Rate) AS TotalValue FROM VRM WHERE GroupID NOT IN(SELECT Group_ID FROM RMGroupIDsForForging) AND QtyInStock>0 GROUP BY VRM.GroupID,VRM.GroupName", con, adOpenForwardOnly, adLockReadOnly
        End If
        LVSOS.ListItems.Clear
        Do Until .EOF
            Set ITM = LVSOS.ListItems.add(, , LVSOS.ListItems.count + 1)
            ITM.Tag = Val(!GroupID & "")
            ITM.ListSubItems.add , , !Description & ""
            ITM.ListSubItems.add , , Val(!TotalStock & "")
            ITM.ListSubItems.add , , Val(!TotalValue & "")
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    LVSOS.Tag = iType
    
    FraSummaryOfStock.Visible = True
    FraSummaryOfStock.ZOrder 0
    
End Sub

Private Sub PrintTotalExportCurrentFY()


    On Error GoTo err
        
    Dim FYDTFrom As Date, FYDTTo As Date
    If Month(Date) >= 6 Then
        FYDTFrom = DateSerial(year(Date), 7, 1)
        FYDTTo = DateSerial(year(Date) + 1, 6, 30)
    Else
        FYDTFrom = DateSerial(year(Date) - 1, 7, 1)
        FYDTTo = DateSerial(year(Date), 6, 30)
    End If
    
    Dim rpt As CRAXDDRT.Report 'New rptTotalExport
    Set rpt = rptApp.OpenReport(RptPath & "\TotalExport.rpt") '
    Dim f As New frmPrevRpt
    Dim Cond As String
    
    rpt.FormulaFields.GetItemByName("CustomAmt").Text = False 'chkCustom.Value
    rpt.FormulaFields.GetItemByName("SubHeading").Text = "'" & Format(FYDTFrom, "dd-MMM-yyyy") & " to " & Format(FYDTTo, "dd-MMM-yyyy") & "'"
    
    Dim strSelection As String
    strSelection = "{VTotalExport.DT}=#" & FYDTFrom & "# To #" & FYDTTo & "#"
    
    
    Dim rptSub As CRAXDDRT.Report
    Set rptSub = rpt.OpenSubreport("currencytotalexport")
    rptSub.RecordSelectionFormula = strSelection
    
    f.ShowReport strSelection, rpt

    Set rpt = Nothing
    Set f = Nothing
    
    Exit Sub
    
err:
    MsgBox err.Description
End Sub

Private Sub PrintAbsentSheet()

    On Error GoTo err
    Dim cmd As New ADODB.Command
    With cmd
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_PrintEmpAbsentList"
        .Parameters("@DTFrom").Value = Date
        .Parameters("@DTTo").Value = Date
        .Execute
    End With
    
    Me.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report, FRM As New frmPrevRpt
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpAbsentSheetSummary.rpt")
    rpt.FormulaFields(1).Text = "'" & Format(Date, "dd-MMM-yyyy") & "'"
    FRM.ShowReport "", rpt
    Me.MousePointer = vbNormal
    
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
End Sub


Private Sub LVSOS_DblClick()

    If LVSOS.ListItems.count = 0 Then Exit Sub
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        If LVSOS.Tag = "1" Then 'SF
            .Open "SELECT Items.ItemID,Items.ItemName,SUM((Qty-IssdQty)*PriceForCost) AS TotalValue,SUM((Qty-IssdQty)) AS TotalStock FROM VStockOrderOpening INNER JOIN (SELECT Opening_RefID FROM VendRcvdDetail GROUP BY Opening_RefID) T1 ON T1.Opening_RefID=VStockOrderOpening.EntryID INNER JOIN Items ON VStockOrderOpening.ItemID=Items.ItemID WHERE GroupID=" & Val(LVSOS.SelectedItem.Tag) & " AND Qty>IssdQty GROUP BY Items.ItemID,Items.ItemName", con, adOpenForwardOnly, adLockReadOnly
        ElseIf LVSOS.Tag = "2" Then 'Forging
            .Open "SELECT VRM.RMID1 AS ItemID,VRM.RMName AS ItemName,SUM(QtyInStock*Rate) AS TotalValue,SUM(QtyInStock) AS TotalStock FROM VRM WHERE GroupID=" & Val(LVSOS.SelectedItem.Tag) & " AND QtyInStock>0 GROUP BY VRM.RMID1,VRM.RMName", con, adOpenForwardOnly, adLockReadOnly
        Else                        'Finished
            .Open "SELECT Items.ItemID,Items.ItemName,SUM(NetQty*PriceForCost) AS TotalValue,SUM(NetQty) AS TotalStock FROM VItemsWithShelfWiseStock INNER JOIN Items ON VItemsWithShelfWiseStock.ItemID=Items.ItemID WHERE GroupID=" & Val(LVSOS.SelectedItem.Tag) & " GROUP BY Items.ItemID,Items.ItemName", con, adOpenForwardOnly, adLockReadOnly
        End If
        LVDOS.ListItems.Clear
        Do Until .EOF
            Set ITM = LVDOS.ListItems.add(, , LVDOS.ListItems.count + 1)
            ITM.Tag = !ITemID & ""
            ITM.ListSubItems.add , , !ITemID & " - " & !ItemName & ""
            ITM.ListSubItems.add , , Val(!TotalStock & "")
            ITM.ListSubItems.add , , Val(!TotalValue & "")
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    FraDOS.Caption = "Detail of " & LVSOS.SelectedItem.SubItems(1) & " :"
    FraDOS.Visible = True
    FraDOS.ZOrder 0
    
End Sub

Private Sub PicWC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If y < 600 Then
        If Button = 1 Then
            If bMouseDown = False Then
                bMouseDown = True
                lDownX = x
                lDownY = y
            End If
        End If
    End If
End Sub

Private Sub PicWC_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If bMouseDown Then
        Me.Move Me.Left + ((x - lDownX) / 15), Me.Top + ((y - lDownY) / 15)
    End If
End Sub

Private Sub PicWC_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    bMouseDown = False
End Sub

Public Function ShowMe(Optional bStockProduction As Boolean = False) As Boolean
    
    'Will hold the handle to the new region.
    Dim hRegion As Long
    '
    'This API can create a Round or Oval Shaped Region.
    hRegion = CreateRoundRectRgn(0, 0, Me.Width / 15, Me.Height / 15, 20, 20)
    '
    'Change the shape of the Form to the newly created Region.
    SetWindowRgn Me.hwnd, hRegion, True
    '
    'Show the new form with the specified shape.
    'frmRegion.Show

    Me.Show
    
    DoEvents
    
    If bStockProduction Then
        PicWC(1).Visible = False
        PicWC(2).Visible = False
        PicWC(3).Visible = False
        PicWC(0).Left = 180
        PicWC(4).Top = PicWC(0).Top
    End If
    
    Call UpdateFields
    
    ShowMe = bSaved
    
End Function

Private Sub Timer1_Timer()
    Call UpdateFields
End Sub

Private Sub UpdateFields()
    
    On Error Resume Next
    Animation(0).Open App.Path & "\LoadingData.avi"
    Animation(1).Open App.Path & "\LoadingData.avi"
    Animation(2).Open App.Path & "\LoadingData.avi"
    Animation(3).Open App.Path & "\LoadingData.avi"
    Animation(4).Open App.Path & "\LoadingData.avi"
            
    Animation(0).Play
    Animation(1).Play
    Animation(2).Play
    Animation(3).Play
    Animation(4).Play
    Me.Refresh
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim myDT As Date
    myDT = DateAdd("d", -1, Date)
    If Format(myDT, "ddd") = "Sun" Then
        myDT = DateAdd("d", -1, myDT)
    End If
    
    Dim FYDTFrom As Date, FYDTTo As Date
    If Month(Date) >= 6 Then
        FYDTFrom = DateSerial(year(Date), 7, 1)
        FYDTTo = DateSerial(year(Date) + 1, 6, 30)
    Else
        FYDTFrom = DateSerial(year(Date) - 1, 7, 1)
        FYDTTo = DateSerial(year(Date), 6, 30)
    End If
    
    With rs
        DoEvents
        '/\/\/\/\/\/\/HR
        .Open "SELECT COUNT(EmpID) AS TotalActiveEmps,SUM(StartingSalary) AS TotalMonthlySalary FROM Employees WHERE Active=1", con, adOpenForwardOnly, adLockReadOnly
        txtNoOfEmployees = Val(!TotalActiveEmps & "")
        txtAvgMonthlySalary = Format(Val(!TotalMonthlySalary & ""), "#,##0")
        .Close
        DoEvents
        .Open "SELECT COUNT(EmpID) AS TotalPresentEmps FROM AttendanceSheet WHERE Attendance=1 AND DT='" & Date & "'", con, adOpenForwardOnly, adLockReadOnly
        txtPresentEmployees = Val(!TotalPresentEmps & "")
        .Close
        txtAbsentEmployees = Val(txtNoOfEmployees) - Val(txtPresentEmployees)
        DoEvents
        .Open "SELECT SUM(LongTermBalance) AS LongTermBalance FROM VLongTermBalance", con, adOpenForwardOnly, adLockReadOnly
        txtLongTerm = Format(Val(!LongTermBalance & ""), "#,##0")
        .Close
        DoEvents
        '.Open "SELECT Sum(Amount) AS STBalance FROM Advances WHERE CONVERT(VARCHAR(5),DT,11) NOT IN(SELECT CONVERT(VARCHAR(5),DT,11) FROM MonthlySalaries)", con, adOpenForwardOnly, adLockReadOnly
        .Open "SELECT Sum(Amount) AS STBalance FROM Advances WHERE Type=0 AND Cleared=0", con, adOpenForwardOnly, adLockReadOnly
        txtShortTerm = Format(Val(!STBalance & ""), "#,##0")
        .Close
        DoEvents
        FLoading(1).Visible = False
        Animation(1).Stop
        DoEvents
        '\/\/\/\
         
         '/////////////////////STOCK///////////////////////
        .Open "SELECT SUM(NetQty*PriceForCost) AS TotalValue,SUM(NetQty) AS TotalStock FROM VItemsWithShelfWiseStock INNER JOIN Items ON VItemsWithShelfWiseStock.ItemID=Items.ItemID", con, adOpenForwardOnly, adLockReadOnly
        txtFinishedStockQty = Format(Val(!TotalStock & ""), "#,##0")
        txtFinishedStockValue = Format(Val(!TotalValue & ""), "#,##0")
        .Close
        
        
        .Open "SELECT SUM((RcvdQty-IssQty)) AS TotalStock,SUM((RcvdQty-IssQty)*PriceForCost) AS TotalValue FROM VendRcvdDetail INNER JOIN VStockOrderOpening ON VendRcvdDetail.Opening_RefID=VStockOrderOpening.EntryID" & _
            " INNER JOIN Items ON VendRcvdDetail.ItemCode=Items.ITemID WHERE (RcvdQty-IssQty)>0", con, adOpenForwardOnly, adLockReadOnly    'SemiFinish
            
        txtSFStockQty = Format(Val(!TotalStock & ""), "#,##0")
        txtSFStockValue = Format(Val(!TotalValue & ""), "#,##0")
        
        .Close
        
        .Open "SELECT SUM(QtyInStock) AS TotalStock,SUM(QtyInStock*Rate) AS TotalValue FROM VRM WHERE GroupID IN(SELECT Group_ID FROM RMGroupIDsForForging)", con, adOpenForwardOnly, adLockReadOnly
        txtForgingStockQty = Format(Val(!TotalStock & ""), "#,##0")
        txtForgingStockValue = Format(Val(!TotalValue & ""), "#,##0")
        .Close
        
        .Open "SELECT SUM(QtyInStock) AS TotalStock,SUM(QtyInStock*Rate) AS TotalValue FROM VRM WHERE GroupID NOT IN(SELECT Group_ID FROM RMGroupIDsForForging)", con, adOpenForwardOnly, adLockReadOnly
        txtOtherInventoryQty = Format(Val(!TotalStock & ""), "#,##0")
        txtOtherInventoryValue = Format(Val(!TotalValue & ""), "#,##0")
        .Close
        
        .Open "SELECT COUNT(*) AS TotalReWorkLots,SUM(Qty-IssQty) AS TotalReWorkQty FROM VVendRcvdDetailReWorkDetail WHERE Qty>IssQty", con, adOpenForwardOnly, adLockReadOnly
        txtReWorkLots = Format(Val(!TotalReWorkLots & ""), "#,##0")
        txtReworkLotsQty = Format(Val(!TotalReWorkQty & ""), "#,##0")
        .Close
        
        DoEvents
        FLoading(0).Visible = False
        Animation(0).Stop
        DoEvents
         '/////////////////////STOCK///////////////////////
         
        '/\/\/\/\/\/\/Financials/\/\/\/\/\/\/\/\/
        .Open "SELECT SUM(Balance) FROM Accounts WHERE SubAccOf='" & CashInHand & "'", con, adOpenForwardOnly, adLockReadOnly
        txtCashAcc = Format(Val(.Fields(0).Value & ""), "#,##0")
        .Close
        .Open "SELECT SUM(Balance) FROM Accounts WHERE SubAccOf='15-002'", con, adOpenForwardOnly, adLockReadOnly
        txtBankBal = Format(Val(.Fields(0).Value & ""), "#,##0")
        .Close
        .Open "SELECT SUM((Balance)) FROM Accounts WHERE LEFT(AccNo,6)='24-004'", con, adOpenForwardOnly, adLockReadOnly
        txtAccPay = Format(Val(.Fields(0).Value & ""), "#,##0")
        .Close
        .Open "SELECT SUM(Balance) FROM Accounts WHERE SubAccOf='15-003'", con, adOpenForwardOnly, adLockReadOnly
        txtAccRcv = Format(Val(.Fields(0).Value & ""), "#,##0")
        .Close
        
        .Open "SELECT SUM(AcquisitionCost) FROM FixedAssets", con, adOpenForwardOnly, adLockReadOnly
        txtFixedAssets = Format(Val(.Fields(0).Value & ""), "#,##0")
        .Close
        
        DoEvents
        FLoading(3).Visible = False
        Animation(3).Stop
        DoEvents
'        .Open "SELECT SUM(Balance) FROM Accounts WHERE LEFT(AccNo,6)='15-007'", con, adOpenForwardOnly, adLockReadOnly
'        txtAdvances = "-" 'Format(Val(.Fields(0).Value & ""), "#,##0")
'        .Close
        '\/\/\/\
          
        
         
        
        '/\/\/\Export\/\/\/\/\
        .Open "SELECT SUM(Qty-ShippedQty) AS InHandQty,SUM((Qty-ShippedQty)*(Price*ExchRate)) AS InHandAmt FROM VFOrderItemsWithShippedQty WHERE OrderNo IN(SELECT OrderNo FROM VUnShippedOrderList) AND Qty>ShippedQty AND CustCode<>'Stock'", con, adOpenForwardOnly, adLockReadOnly
        txtInHandQty = Format(Val(!InHandQty & ""), "#,##0")
        txtInHandValue = Format(Val(!InHandAmt & ""), "#,##0")
        .Close
        .Open "SELECT SUM(Balance) AS TotalSales FROM Accounts WHERE SubAccOf='31-001'", con, adOpenForwardOnly, adLockReadOnly
        txtTotalSales = Format(Abs(Val(!TotalSales & "")), "#,##0")
        .Close
        .Open "SELECT SUM(Credit-Debit) AS TotalSales FROM Vouchers WHERE LEFT(AccNo,6)='31-001' AND VDate BETWEEN '" & FYDTFrom & "' AND '" & FYDTTo & "'", con, adOpenForwardOnly, adLockReadOnly
        txtTotalSalesFY = Format(Val(!TotalSales & ""), "#,##0")
        .Close
        txtRebateRcvable = "-"
        
        DoEvents
        FLoading(2).Visible = False
        Animation(2).Stop
        DoEvents
        '/\/\/\/\/\/\/\/\/\/\/
        
        Call LoadProductionData
        
      End With
       
    
    Exit Sub
err:
    MsgBox err.Description
    Exit Sub

End Sub

Private Sub PrintUnshippedOrderList()

    On Error GoTo err
    Me.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptFOrderList.rpt")
        
    Dim strSQLString As String
    strSQLString = "SELECT * FROM VFOrderList  WHERE TotalShippedQty<TotalOrderQty AND ISNULL(Cancelled,0)=0 AND CustCode<>'Stock' ORDER BY CustCode"
    
    rpt.FormulaFields.GetItemByName("ForCustomer").Text = "'<All Customers>'"
    rpt.FormulaFields.GetItemByName("ForFromTo").Text = "'All Pending'"
    
    frmPrevRpt.ShowReport " ", rpt, strSQLString
    
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Private Sub PrintLots(Optional lProcessID As Long = 0)

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\RunningLots.rpt")
    Dim strSelection As String
    If lProcessID = 0 Then
        strSelection = ""
    Else
        strSelection = "{VRunningLots.ProcessID}=" & lProcessID
    End If
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub LoadProductionData()
    
    Dim rs As New ADODB.Recordset
    With rs
        '/////////////////////PRODUCTION/////////////////////
        .Open "SELECT COUNT(*) AS TotalCount,SUM(QtyToRcv) AS TotalQty FROM VVendOrdersToRcv WHERE GroupID IN(SELECT Group_ID FROM RMGroupIDsForForging)", con, adOpenForwardOnly, adLockReadOnly
        txtForgingOrdersNos = Format(Val(!TotalCount & ""), "#,##0")
        txtForgingOrdersQty = Format(Val(!TotalQty & ""), "#,##0")
        .Close
         
        .Open "SELECT COUNT(DISTINCT MasterPONo) AS TotalCount,SUM(VendIssdDetail_ReturnDTs.IssQty-VendIssdDetail_ReturnDTs.RcvQty+ISNULL(WastageQty,0)) AS TotalQty " & _
          "FROM VendIssued INNER JOIN VendIssdDetail ON VendIssued.EntryID=VendIssdDetail.RefID " & _
          "INNER JOIN VendIssdDetail_ReturnDTs ON VendIssdDetail.EntryID=VendIssdDetail_ReturnDTs.VIS_RefID " & _
          "INNER JOIN Makers ON VendIssued.VendID=Makers.VendID LEFT OUTER JOIN VendIssdDetail_MoreDetails ON VendIssdDetail.EntryID=VendIssdDetail_MoreDetails.VID_RefID " & _
          "LEFT OUTER JOIN VIssuanceEntryWithWastage ON VendIssdDetail.EntryID=VIssuanceEntryWithWastage.EntryID " & _
          "WHERE Closed=0 AND LotNo='0' AND VendIssdDetail_ReturnDTs.IssQty+CASE WHEN ISNULL(IssRcvBalance,0)<0 THEN ABS(ISNULL(IssRcvBalance,0)) ELSE 0 END+ISNULL(Wastage_Return,0)>VendIssdDetail_ReturnDTs.RcvQty", con, adOpenForwardOnly, adLockReadOnly
          
        txtMakerOrdersNos = Format(Val(!TotalCount & ""), "#,##0")
        txtMakerOrdersQty = Format(Val(!TotalQty & ""), "#,##0")
        .Close
        
        .Open "SELECT COUNT(*) AS TotalCount,SUM(Qty) AS TotalQty FROM VRunningLots_Simple", con, adOpenForwardOnly, adLockReadOnly
        txtRunningLotsNos = Format(Val(!TotalCount & ""), "#,##0")
        txtRunningLotsQty = Format(Val(!TotalQty & ""), "#,##0")
        .Close
        
        Dim lProcessID As Long
        lProcessID = Val(getGeneralDataValue("HeatTreatmentProcessID"))
        .Open "SELECT COUNT(*) AS TotalCount,SUM(Qty) AS TotalQty FROM VRunningLots_Simple WHERE ProcessID=" & lProcessID, con, adOpenForwardOnly, adLockReadOnly
        txtLotsAtTemperNos = Format(Val(!TotalCount & ""), "#,##0")
        txtLotsAtTemperNos.Tag = lProcessID
        txtLotsAtTemperQty = Format(Val(!TotalQty & ""), "#,##0")
        .Close
        
        lProcessID = Val(getGeneralDataValue("PolishingProcessID"))
'        If strCompany = "IAA" Or strCompany = "Kami" Then
'            .Open "SELECT COUNT(*) AS TotalCount,SUM(Qty) AS TotalQty FROM VRunningLots_Simple WHERE ProcessID IN(182,185,186,177,180,199,210)", con, adOpenForwardOnly, adLockReadOnly
'        Else
            .Open "SELECT COUNT(*) AS TotalCount,SUM(Qty) AS TotalQty FROM VRunningLots_Simple WHERE ProcessID IN(" & lProcessID & ")", con, adOpenForwardOnly, adLockReadOnly
'        End If
        txtLotsAtPolishingNos = Format(Val(!TotalCount & ""), "#,##0")
        txtLotsAtPolishingNos.Tag = lProcessID
        txtLotsAtPolishingQty = Format(Val(!TotalQty & ""), "#,##0")
        .Close
        
        lProcessID = Val(getGeneralDataValue("Q.CInspection100%"))
        .Open "SELECT COUNT(*) AS TotalCount,SUM(Qty) AS TotalQty FROM VRunningLots_Simple WHERE ProcessID=" & lProcessID, con, adOpenForwardOnly, adLockReadOnly
        txtLotsAtQCNos = Format(Val(!TotalCount & ""), "#,##0")
        txtLotsAtQCNos.Tag = lProcessID
        txtLotsAtQCQty = Format(Val(!TotalQty & ""), "#,##0")
        .Close
        
        lProcessID = Val(getGeneralDataValue("PackingProcessID"))
        .Open "SELECT COUNT(*) AS TotalCount,SUM(Qty) AS TotalQty FROM VRunningLots_Simple WHERE ProcessID=" & lProcessID, con, adOpenForwardOnly, adLockReadOnly
        txtLotsAtPackedNos = Format(Val(!TotalCount & ""), "#,##0")
        txtLotsAtPackedNos.Tag = lProcessID
        txtLotsAtPackedQty = Format(Val(!TotalQty & ""), "#,##0")
        .Close
        
        DoEvents
        FLoading(4).Visible = False
        Animation(4).Stop
        DoEvents
        
        '/\/\//\/\/Production/////////////////////
    End With
    Set rs = Nothing
    
End Sub

Private Sub PrintReWorkLots()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ReWorkLots.rpt")
    
    Dim strSelection As String
    strSelection = "{VVendRcvdDetailReWorkDetail.Qty}>{VVendRcvdDetailReWorkDetail.IssQty}"
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub PrintForgingOpenOrders()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\RM_PO_List.rpt")
    rpt.FormulaFields.GetItemByName("Heading").Text = "'Forging Open Orders'"
    
    Dim strSelection As String
    strSelection = "{VVendOrdersToRcv.OrderNo}={VVendOrders.OrderNo} AND {@BalanceQty}>0"
    strSelection = strSelection & " AND {Material.ForgingGroupID}<>0"
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
End Sub

