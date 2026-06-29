VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{77EBD0B1-871A-4AD1-951A-26AEFE783111}#2.1#0"; "VBalExpBar6.ocx"
Begin VB.MDIForm MainForm 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   11055
   ClientLeft      =   165
   ClientTop       =   315
   ClientWidth     =   17505
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CDExcel 
      Left            =   4740
      Top             =   2580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Excel Files Only|*.xls"
   End
   Begin VB.PictureBox PicFavourites 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9810
      Left            =   14190
      ScaleHeight     =   9780
      ScaleWidth      =   3030
      TabIndex        =   21
      Top             =   855
      Visible         =   0   'False
      Width           =   3060
      Begin VB.CommandButton cmdHideFavourites 
         Appearance      =   0  'Flat
         BackColor       =   &H00E3FBFB&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   10155
         Left            =   -30
         Picture         =   "MainForm.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         Width           =   270
      End
      Begin vbalExplorerBarLib6.vbalExplorerBarCtl ExpBarFavourites 
         Height          =   10260
         Left            =   255
         TabIndex        =   23
         Top             =   0
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   18098
         BackColorEnd    =   16777215
         BackColorStart  =   16777215
      End
   End
   Begin VB.PictureBox PicShowFavourites 
      Align           =   4  'Align Right
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      Height          =   9810
      Left            =   17250
      ScaleHeight     =   9810
      ScaleWidth      =   255
      TabIndex        =   19
      Top             =   855
      Width           =   255
      Begin VB.CommandButton cmdShowFavourites 
         Appearance      =   0  'Flat
         BackColor       =   &H00E3FBFB&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   10185
         Left            =   0
         Picture         =   "MainForm.frx":0C02
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         Width           =   270
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   9810
      Left            =   255
      ScaleHeight     =   9780
      ScaleWidth      =   3030
      TabIndex        =   8
      Top             =   855
      Visible         =   0   'False
      Width           =   3060
      Begin VB.CommandButton cmdHideExpBar 
         Appearance      =   0  'Flat
         BackColor       =   &H00E3FBFB&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9465
         Left            =   2760
         Picture         =   "MainForm.frx":14D0
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   270
      End
      Begin vbalExplorerBarLib6.vbalExplorerBarCtl ExpBar 
         Height          =   10260
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   18098
         BackColorEnd    =   16777215
         BackColorStart  =   16777215
      End
   End
   Begin VB.PictureBox PicShowExpBar 
      Align           =   3  'Align Left
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      Height          =   9810
      Left            =   0
      ScaleHeight     =   9810
      ScaleWidth      =   255
      TabIndex        =   6
      Top             =   855
      Visible         =   0   'False
      Width           =   255
      Begin VB.CommandButton cmdShowExpBar 
         Appearance      =   0  'Flat
         BackColor       =   &H00E3FBFB&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9465
         Left            =   0
         Picture         =   "MainForm.frx":1686
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   270
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4725
      Top             =   885
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Please Select Employee's Picture"
      Filter          =   "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
   End
   Begin VB.Timer TmrAlerts 
      Interval        =   60000
      Left            =   4785
      Top             =   2100
   End
   Begin VB.PictureBox PicAlerts 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00000098&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   17505
      TabIndex        =   2
      Top             =   10665
      Visible         =   0   'False
      Width           =   17505
      Begin MSComCtl2.Animation Animation2 
         Height          =   705
         Left            =   0
         TabIndex        =   3
         Top             =   5010
         Visible         =   0   'False
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1244
         _Version        =   393216
         Center          =   -1  'True
         BackStyle       =   1
         BackColor       =   0
         FullWidth       =   43
         FullHeight      =   47
      End
      Begin VB.Label lblCustomerEvents 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Events"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   12405
         MouseIcon       =   "MainForm.frx":19B7
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   3420
      End
      Begin VB.Label lblLateOrderAlert 
         BackStyle       =   0  'Transparent
         Caption         =   "Late Orders"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   7650
         MouseIcon       =   "MainForm.frx":1B09
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   0
         Width           =   3420
      End
      Begin VB.Label lblAlertsHeasing 
         Alignment       =   2  'Center
         BackColor       =   &H00000098&
         Caption         =   "Alerts :"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   75
         TabIndex        =   16
         Top             =   15
         Width           =   1140
      End
      Begin VB.Label lblAlerts 
         BackStyle       =   0  'Transparent
         Caption         =   "lblAlerts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   1275
         MouseIcon       =   "MainForm.frx":1C5B
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   0
         Width           =   6180
      End
   End
   Begin MSComctlLib.ImageList ImageList5 
      Left            =   3330
      Top             =   2625
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":1DAD
            Key             =   "Payroll"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":2352
            Key             =   "Import"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":2AAC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":3120
            Key             =   "Sampling"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":37FA
            Key             =   "Dashboard"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":3DB7
            Key             =   "FixedAssets"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":42E9
            Key             =   "QMS"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":4893
            Key             =   "ChangePassword"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":4DE1
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":52C6
            Key             =   "Production"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":58E4
            Key             =   "Stock"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":5EEC
            Key             =   "Company"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":6451
            Key             =   "Financial"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":69B0
            Key             =   "Export"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer3 
      Interval        =   60000
      Left            =   4155
      Top             =   1080
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3345
      Top             =   1095
   End
   Begin VB.PictureBox PicTitle 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   17475
      TabIndex        =   0
      Top             =   0
      Width           =   17505
      Begin VB.Frame FraWelcome 
         BorderStyle     =   0  'None
         Height          =   915
         Left            =   13170
         TabIndex        =   11
         Top             =   -90
         Width           =   4950
         Begin VB.Label lblTime 
            BackColor       =   &H00000098&
            BackStyle       =   0  'Transparent
            Caption         =   "[Date && Time]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00333333&
            Height          =   240
            Left            =   45
            TabIndex        =   15
            Top             =   660
            Width           =   2865
         End
         Begin VB.Label lblWelcomeMsg 
            BackColor       =   &H00000098&
            BackStyle       =   0  'Transparent
            Caption         =   "Asalam-O-Alikum, MARHABA!"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00333333&
            Height          =   240
            Left            =   45
            TabIndex        =   14
            Top             =   450
            Width           =   3180
         End
         Begin VB.Label lblUserName 
            BackColor       =   &H00000098&
            BackStyle       =   0  'Transparent
            Caption         =   "[User Name]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000098&
            Height          =   285
            Left            =   1905
            TabIndex        =   13
            Top             =   135
            Width           =   2865
         End
         Begin VB.Label lblWelcome 
            Alignment       =   2  'Center
            BackColor       =   &H00000098&
            Caption         =   "Welcome"
            BeginProperty Font 
               Name            =   "Impact"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   30
            TabIndex        =   12
            Top             =   120
            Width           =   1800
         End
      End
      Begin VB.CommandButton cmdMin 
         Height          =   240
         Left            =   12810
         Picture         =   "MainForm.frx":6F78
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   210
         Visible         =   0   'False
         Width           =   285
      End
      Begin MSComctlLib.Toolbar MainToolbar 
         Height          =   570
         Left            =   0
         TabIndex        =   1
         Top             =   -15
         Width           =   16440
         _ExtentX        =   28998
         _ExtentY        =   1005
         ButtonWidth     =   1746
         ButtonHeight    =   1005
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   27
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Company"
               Key             =   "Company"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Export"
               Key             =   "Export"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "Import"
               Key             =   "Import"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "Sampling"
               Key             =   "Sampling"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Financial"
               Key             =   "Financial"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "HR"
               Key             =   "Payroll"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Stock"
               Key             =   "Stock"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Production"
               Key             =   "Production"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "QMS"
               Key             =   "QMS"
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "Fixed Assets"
               Key             =   "FixedAssets"
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Dashboard"
               Key             =   "Dashboard"
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "DashboardDash"
               Style           =   3
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "Help"
               Key             =   "Help"
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "HelpDash"
               Style           =   3
            EndProperty
            BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Exit"
               Key             =   "Exit"
            EndProperty
            BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "PlaceHolder1"
               Style           =   4
               Object.Width           =   500
            EndProperty
         EndProperty
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3765
      Top             =   1095
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   4800
      Top             =   1695
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3930
      Top             =   1500
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":733E
            Key             =   "Cont"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":7840
            Key             =   "PaySheet"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":78DF
            Key             =   "PayEntry"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":79BD
            Key             =   "Issue"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":7A92
            Key             =   "Recieve"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":7B34
            Key             =   "Items"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":7C17
            Key             =   "ContLedger"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   3375
      Top             =   1470
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   108
      ImageHeight     =   17
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":81F9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   3945
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   150
      ImageHeight     =   17
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":8366
            Key             =   "Accounts"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":95F0
            Key             =   "test"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":9891
            Key             =   "Company"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":A2F8
            Key             =   "Glossary"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":ACCB
            Key             =   "Developers"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":B895
            Key             =   "Export"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":CA37
            Key             =   "Management"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":D7D2
            Key             =   "Orders"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":E8E9
            Key             =   "Payroll"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":FBE5
            Key             =   "Policies"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":10CEC
            Key             =   "Production"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":11EA7
            Key             =   "Purchases"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":12956
            Key             =   "Stats"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":13373
            Key             =   "Inventory"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":1449D
            Key             =   "BusPart"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3315
      Top             =   2085
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   130
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":15579
            Key             =   "test"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":16C4B
            Key             =   "Company"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":176B2
            Key             =   "Glossary"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":18085
            Key             =   "Management"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":18E20
            Key             =   "Policies"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":19F27
            Key             =   "Purchases"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":1A9D6
            Key             =   "BusPart"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":1BAB2
            Key             =   "Export"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":1CC54
            Key             =   "Developers"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":1D81E
            Key             =   "Inventory"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":1E948
            Key             =   "Orders"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":1FA5F
            Key             =   "Payroll"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":20D5B
            Key             =   "Production"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":21F16
            Key             =   "Stats"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":22933
            Key             =   "Accounts"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuSubOptions 
         Caption         =   "Use Calculator"
         Index           =   1
      End
      Begin VB.Menu mnuSubOptions 
         Caption         =   "User Manager"
         Index           =   2
      End
      Begin VB.Menu mnuSubOptions 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuSubOptions 
         Caption         =   "Change Password"
         Index           =   4
      End
      Begin VB.Menu mnuSubOptions 
         Caption         =   "Log Off"
         Index           =   5
      End
      Begin VB.Menu mnuSubOptions 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuSubOptions 
         Caption         =   "E&xit"
         Index           =   7
      End
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "General Settings"
      Visible         =   0   'False
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Set Initial Lot Process"
         Index           =   1
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Percentage Ranges"
         Index           =   3
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "&Components"
         Index           =   5
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "&Manufacturing Processes"
         Index           =   7
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "&FDA Processes"
         Index           =   8
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Steel Vendors"
         Index           =   10
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Test Labs"
         Index           =   11
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Measuring &Units"
         Index           =   13
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Compositions"
         Index           =   14
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Cities && Ports"
         Index           =   16
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Bank List"
         Index           =   18
         Shortcut        =   ^B
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Payment Terms"
         Index           =   19
      End
   End
   Begin VB.Menu mnupopdept 
      Caption         =   "popMenuDept"
      Visible         =   0   'False
      Begin VB.Menu mnuDeptRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu popdeptdash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowDept 
         Caption         =   "Show Inactive Departments"
      End
      Begin VB.Menu mnuDeptdash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeptInact 
         Caption         =   "Make This Department Inactive"
      End
      Begin VB.Menu mnuDeptdash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeptnew 
         Caption         =   "Add &New Department"
      End
      Begin VB.Menu mnuDepdash4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeptEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuDeptdash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeptPrint 
         Caption         =   "&Print Department List"
      End
   End
   Begin VB.Menu mnuVenders 
      Caption         =   "Venders"
      Visible         =   0   'False
      Begin VB.Menu mnuVendList 
         Caption         =   "&Vender List"
      End
      Begin VB.Menu vendlistdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVendAssItems 
         Caption         =   "Assigned Items"
      End
      Begin VB.Menu VendAssItemsdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVendPymt 
         Caption         =   "Vender Advance Payment"
      End
      Begin VB.Menu mnuVendPaySheetdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVendPaySheet 
         Caption         =   "Vender Payment Sheet"
      End
      Begin VB.Menu mnuAdvdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVendLoans 
         Caption         =   "Loans"
         Begin VB.Menu mnuVendAdvShort 
            Caption         =   "New Short Term Advance"
         End
         Begin VB.Menu mnuVendAdvLong 
            Caption         =   "New Long Term Loan"
         End
         Begin VB.Menu mnuVendLoanLedgersdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVendShortLoanLedger 
            Caption         =   "Short Term Loan Ledger"
         End
         Begin VB.Menu mnuVendLongLoanLedger 
            Caption         =   "Long Term Loan Ledger"
         End
      End
      Begin VB.Menu VendLedgerdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVendLedger 
         Caption         =   "Vender Ledger"
      End
   End
   Begin VB.Menu mnuRMVend 
      Caption         =   "Raw Material"
      Visible         =   0   'False
      Begin VB.Menu mnuRMGroups 
         Caption         =   "RM Groups"
      End
      Begin VB.Menu mnuRM 
         Caption         =   "Raw Material"
      End
      Begin VB.Menu RMPOdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRMPO 
         Caption         =   "New Purchase Order"
      End
      Begin VB.Menu mnuPOrderList 
         Caption         =   "Purchase Order List"
      End
      Begin VB.Menu RRMdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecRM 
         Caption         =   "Receive Raw Material"
      End
      Begin VB.Menu mnuIssRM 
         Caption         =   "Issue Raw Material"
      End
   End
   Begin VB.Menu mnuHICORes 
      Caption         =   "Hico Resources"
      Visible         =   0   'False
      Begin VB.Menu mnuHRGroups 
         Caption         =   "Hico Resources Groups"
      End
   End
   Begin VB.Menu mnuCompanya 
      Caption         =   "Company"
      Visible         =   0   'False
      Begin VB.Menu mnuItems 
         Caption         =   "Items"
         Begin VB.Menu mnuItmGrps 
            Caption         =   "Item Groups"
         End
         Begin VB.Menu mnuItemsFinQuality 
            Caption         =   "Item's Finishing Quality"
         End
         Begin VB.Menu mnuItemTypes 
            Caption         =   "Item Types"
         End
      End
   End
   Begin VB.Menu mnuCustomers 
      Caption         =   "Customers"
      Visible         =   0   'False
      Begin VB.Menu mnuNewOrd 
         Caption         =   "New Order"
      End
      Begin VB.Menu mnuNewStockOrder 
         Caption         =   "New Stock Order"
      End
      Begin VB.Menu mnuCustOrdList 
         Caption         =   "Customer Order List"
      End
      Begin VB.Menu custdash4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrdStatus 
         Caption         =   "Order Status"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOrdStatus1 
         Caption         =   "Order Status(Other Format)"
      End
      Begin VB.Menu mnuOrdStatusComp 
         Caption         =   "Order Status Complete"
         Visible         =   0   'False
      End
      Begin VB.Menu custdash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewProforma 
         Caption         =   "New Proforma"
      End
      Begin VB.Menu mnuCustProformaList 
         Caption         =   "Customer Proforma List"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Visible         =   0   'False
      Begin VB.Menu mnuDailyRpts 
         Caption         =   "Daily Reports"
      End
      Begin VB.Menu mnuDailyRptsdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIssRpts 
         Caption         =   "Issue Reports"
      End
      Begin VB.Menu mnuRcvRpts 
         Caption         =   "Receive Reports"
      End
      Begin VB.Menu mnuRcvRptsdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRMListrpt 
         Caption         =   "Raw Material List"
      End
      Begin VB.Menu mnuRMVendListrpt 
         Caption         =   "Raw Material Vender List"
      End
      Begin VB.Menu mnuRMIssRpts 
         Caption         =   "Raw Material Issue Reports"
      End
      Begin VB.Menu mnuRMRcvRpts 
         Caption         =   "Raw Material ReceiveReports"
      End
      Begin VB.Menu mnuRMRcvRptdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHRrpt 
         Caption         =   "Hico Resources"
      End
      Begin VB.Menu mnuAssHRrpt 
         Caption         =   "Assigned Hico Resources"
      End
      Begin VB.Menu mnuCustListrptdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCustListrpt 
         Caption         =   "Customer List"
      End
      Begin VB.Menu mnuReqrptdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReqrpt 
         Caption         =   "Requsition Report"
      End
      Begin VB.Menu mnuPORpt 
         Caption         =   "P.O. Report"
      End
      Begin VB.Menu mnuPORptdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProdPlanrpt 
         Caption         =   "Production Planning Report"
      End
      Begin VB.Menu mnuPPlanrptdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVendListrpt 
         Caption         =   "Vender List"
      End
   End
   Begin VB.Menu mnuProdPhases 
      Caption         =   "Production Phases"
      Visible         =   0   'False
      Begin VB.Menu mnuReq 
         Caption         =   "Requsition"
         Visible         =   0   'False
      End
      Begin VB.Menu prodplandash 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuProdPlan 
         Caption         =   "Prdoduction Planning"
      End
      Begin VB.Menu TestDash 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAuth 
         Caption         =   "Authorize"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAuthDash 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "FORGING"
         Index           =   1
         Begin VB.Menu mnuIssForg 
            Caption         =   "Issue Items"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecVForg 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSForg 
            Caption         =   "Receive Items From Supervisor"
         End
         Begin VB.Menu mnudForg 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastForg 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuForgSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuForgSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "MACHINING"
         Index           =   2
         Begin VB.Menu mnuIssMach 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVMach 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSMach 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudMach 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastMach 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuMachSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMachSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "DRILLING"
         Index           =   3
         Begin VB.Menu mnuIssDril 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVDril 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSDril 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudDril 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastDril 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuDrilSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDrilSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "FILLING"
         Index           =   4
         Begin VB.Menu mnuIssFill 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVFill 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSFill 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudFill 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastFill 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuFillSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFillSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "FUNTIONAL INSPECTION UN-FINISH"
         Index           =   5
         Begin VB.Menu mnuIssFunc 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVFunc 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSFunc 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudFunc 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastFunc 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuFuncSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFuncSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "HEAT TREATMENT"
         Index           =   6
         Begin VB.Menu mnuIssHeat 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVHeat 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSHeat 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudHeat 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastHeat 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuHeatSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHeatSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "POLISHING"
         Index           =   7
         Begin VB.Menu mnuIssPoli 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVPoli 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSPoli 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecWastPoli 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnudPoli 
            Caption         =   "-"
         End
         Begin VB.Menu mnuIssRepairPoli 
            Caption         =   "Issue Repair"
         End
         Begin VB.Menu mnuRcvRepairPoli 
            Caption         =   "Receive Repair"
         End
         Begin VB.Menu mnuPoliSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPoliSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "PASSIVATION"
         Index           =   8
         Begin VB.Menu mnuIsspass 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVPass 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSPass 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudPass 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastPass 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuPassSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPassSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "FUNTIONAL INSPECTION"
         Index           =   9
         Begin VB.Menu mnuIssFunt 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVFunt 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSFunt 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudFunt 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastFunt 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuFuntdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFuntSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "POLISHING && RE-BUFFING"
         Index           =   10
         Begin VB.Menu mnuIssPoRe 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVPoRe 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSPoRe 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudPoRe 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastPoRe 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuPORESSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPORESS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "ULTRASONIC CLEAING"
         Index           =   11
         Begin VB.Menu mnuIssUltr 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVUltr 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSUltr 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudUltr 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastUltr 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuUltrSSDash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuUltrSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "PACKING && LABELING"
         Index           =   12
         Begin VB.Menu mnuIssPack 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVPack 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSPack 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudPack 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastPack 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuPackSSDash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPackSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "HOLDING && DISTRIBUTION"
         Index           =   13
         Begin VB.Menu mnuIssHold 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVHold 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSHold 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudHold 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastHold 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuHoldSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHoldSS 
            Caption         =   "Stock Status"
         End
      End
   End
   Begin VB.Menu mnumainSteel 
      Caption         =   "Steel"
      Visible         =   0   'False
      Begin VB.Menu mnuSteel 
         Caption         =   "&Steel List"
      End
      Begin VB.Menu mnuGuage 
         Caption         =   "Guage List"
      End
   End
   Begin VB.Menu mnuPopCust 
      Caption         =   "Customers"
      Visible         =   0   'False
      Begin VB.Menu PopSubCust 
         Caption         =   "Add New Customer"
         Index           =   0
      End
      Begin VB.Menu PopSubCust 
         Caption         =   "Customer List"
         Index           =   1
      End
      Begin VB.Menu PopSubCust 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu PopSubCust 
         Caption         =   "Customer Assigned Items"
         Index           =   3
      End
      Begin VB.Menu PopSubCust 
         Caption         =   "Customer Ledger"
         Index           =   4
      End
   End
   Begin VB.Menu mnuPopProdVend 
      Caption         =   "Venders"
      Visible         =   0   'False
      Begin VB.Menu PopSubProdVend 
         Caption         =   "Add New Vender"
         Index           =   0
      End
      Begin VB.Menu PopSubProdVend 
         Caption         =   "Vender List"
         Index           =   1
      End
      Begin VB.Menu PopSubProdVend 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu PopSubProdVend 
         Caption         =   "Vender Assigned Items"
         Index           =   3
      End
      Begin VB.Menu PopSubProdVend 
         Caption         =   "Vender Ledger"
         Index           =   4
      End
   End
   Begin VB.Menu mnuPopRMVend 
      Caption         =   "R.M. Venders"
      Visible         =   0   'False
      Begin VB.Menu PopSubRMVend 
         Caption         =   "Add New R.M. Vender"
         Index           =   0
      End
      Begin VB.Menu PopSubRMVend 
         Caption         =   "R.M. Vender List"
         Index           =   1
      End
      Begin VB.Menu PopSubRMVend 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu PopSubRMVend 
         Caption         =   "Assigned Items"
         Index           =   3
      End
      Begin VB.Menu PopSubRMVend 
         Caption         =   "R.M. Vender Ledger"
         Index           =   4
      End
   End
   Begin VB.Menu mnuPopSteelVend 
      Caption         =   "Steel Venders"
      Visible         =   0   'False
      Begin VB.Menu PopSubSteelVend 
         Caption         =   "Add New Steel Vender"
         Index           =   0
      End
      Begin VB.Menu PopSubSteelVend 
         Caption         =   "Steel Vender List"
         Index           =   1
      End
   End
   Begin VB.Menu mnupopemp 
      Caption         =   "PopMenuEmp"
      Visible         =   0   'False
      Begin VB.Menu mnuempRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuempdash0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowInactive 
         Caption         =   "Show Inactive Employees/Departments"
      End
      Begin VB.Menu mnuempdash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMake 
         Caption         =   "Make This Employee Inactive"
      End
      Begin VB.Menu dd1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuempNewDept 
         Caption         =   "Add New &Department"
      End
      Begin VB.Menu mnuEmpNewEmp 
         Caption         =   "Add New &Employee"
      End
      Begin VB.Menu mnuempdash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmpEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuEmpdash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAssItems 
         Caption         =   "Show &Assigned Items"
      End
      Begin VB.Menu mnuEmpPrint 
         Caption         =   "&Print Employee List (Internal Salary)"
      End
      Begin VB.Menu mnuPrintEmpExt 
         Caption         =   "Print Employee List (E&xternal Salary)"
      End
      Begin VB.Menu mnuPrintEmpTrans 
         Caption         =   "Print Employee List (Company Transport)"
      End
      Begin VB.Menu dashemps 
         Caption         =   "-"
      End
      Begin VB.Menu mnuappForm 
         Caption         =   "Print Application Form"
      End
      Begin VB.Menu mnuPrintCard 
         Caption         =   "Print Employee Card"
         Begin VB.Menu mnuPrintEmpCard 
            Caption         =   "This Employee Only"
         End
         Begin VB.Menu mnuPrintAllCard 
            Caption         =   "For All Employees"
         End
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempForeWin As Long
Dim ForeWin As Long
Dim Ready As Boolean
'********************************************
Dim strAlerts As String
Dim iAlerts As Integer
Dim strAlertCusts() As String
Dim strAlertOrders() As String
Dim iUpper As Integer, iIndex As Integer
Dim bMain As Boolean, iDelayCounter As Integer
'********************************************
Dim bResizing As Boolean
Dim DTLastAlert As Date

Private Sub dashBankLIst_Click()
    frmBankList.Show
End Sub

Private Sub BusPartToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.key
        Case "Customers"
            Me.PopUpMenu mnuPopCust, 512, CoolBar1.Width
        Case "ProdVend"
            Me.PopUpMenu mnuPopProdVend, 512, CoolBar1.Width
        Case "RMVend"
            Me.PopUpMenu mnuPopRMVend, 512, CoolBar1.Width
    End Select
    
End Sub

Private Sub cmdHideFavourites_Click()
    PicFavourites.Visible = False
    PicShowFavourites.Visible = True
End Sub

Private Sub cmdShowFavourites_Click()
    PicShowFavourites.Visible = False
    PicFavourites.Visible = True
End Sub

Private Sub Command1_Click()

    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    With rs
        .Open "SELECT EntryID,CompItemID,PackingMode FROM FCustomerCatalog WHERE CustCode='EMD'", con, adOpenDynamic, adLockOptimistic
        rs1.Open "SELECT EntryID,CompItemID,PackingMode FROM FCustomerCatalog WHERE CustCode='SSI'", con, adOpenDynamic, adLockOptimistic
        Do Until .EOF
            rs1.Find "CompItemID='" & !CompITemID & "'", , adSearchForward, 1
            If rs1.EOF = False Then
                ![PackingMode] = rs1![PackingMode]
            End If
            .MoveNEXT
        Loop
        .Close
        rs1.Close
    End With
    Set rs = Nothing
    Set rs1 = Nothing
    
End Sub

Private Sub CoolBar1_Resize()
    For i = 1 To CoolBar1.Bands.count
        If CoolBar1.Bands(i).NewRow Then CoolBar1.Bands(i).NewRow = False
    Next i
End Sub

Private Sub cmdMin_Click()
    Me.WindowState = vbMinimized
End Sub

Public Sub ExpBar_BarClick(bar As vbalExplorerBarLib6.cExplorerBar)
    
    If bar.Items.count > 0 Then Exit Sub
    
    If Left(bar.key, 3) = "Acc" Then
        Call FillAccountsItems(bar.key)
        Exit Sub
    ElseIf Left(bar.key, 2) = "FA" Then
        Call FillFixedAssetsItems(bar.key)
        Exit Sub
    ElseIf Left(bar.key, 3) = "Prl" Then
        Call FillPayrollItems(bar.key)
        Exit Sub
    ElseIf Left(bar.key, 3) = "Stk" Then
        Call FillStockItems(bar.key)
        Exit Sub
    ElseIf Left(bar.key, 3) = "Prd" Then
        Call FillProductionItems(bar.key)
        Exit Sub
'    ElseIf Left(bar.key, 2) = "QA" Then
'        Call FillQAItems(bar.key)
'        Exit Sub
    ElseIf Left(bar.key, 3) = "Das" Then
        Call FillDashBoardItems(bar.key)
        Exit Sub
    ElseIf Left(bar.key, 3) = "Hlp" Then
        Call FillHelpItems(bar.key)
        Exit Sub
    ElseIf Left(bar.key, 3) = "QMS" Then
        Call FillQMSItems(bar.key)
        Exit Sub
    ElseIf Left(bar.key, 3) = "SAM" Then
        Call fillSamplingItems(bar.key)
        Exit Sub
    ElseIf Left(bar.key, 3) = "Imp" Then
        Call FillImportItems(bar.key)
        Exit Sub
    End If
    
    
    
    Select Case bar.key
        Case "Company"
'            If strCompany = "Dr-Frgz" And strSubCompany = "Intl" Then
'                Call bar.Items.add(, "PushToWeb", "Push To Web")
'            End If
            Call bar.Items.add(, "CompanyInfo", "Company Info") '.CanClick = UserHasAccess("ExpCompInfo")
'            If strCompany = "QEL" Then
'                Call bar.Items.add(, "CompanyCatalogSurgical", "Surgical Catalog") '.CanClick = UserHasAccess("ExpCompCatalog")
'                Call bar.Items.add(, "CompanyCatalogDental", "Dental Catalog") '.CanClick = UserHasAccess("ExpCompCatalog")
'                Call bar.Items.add(, "CompanyCatalogOthers", "Other Products") '.CanClick = UserHasAccess("ExpCompCatalog")
'            Else
                Call bar.Items.add(, "CompanyCatalog", "Company Catalog") '.CanClick = UserHasAccess("ExpCompCatalog")
                Call bar.Items.add(, "CompanyImportCatalog", "Import Catalog")
''                If strCompany = "Dr-Frgz" Then
''                    Call bar.Items.add(, "SampleCatalog", "Sample Catalog") '.CanClick = UserHasAccess("ExpCompCatalog")
''                End If
'            End If
'            If strCompany <> "Dr-Frgz" Then
'                Call bar.Items.add(, "ResetAllProcesses", "Reset && Re-Assign Processes") '.CanClick = UserHasAccess("ExpGSPCodes")
'            End If
            Call bar.Items.add(, "GroupList", "Group List")
            Call bar.Items.add(, "AdditionalGroupList", "Additional Groups")
            Call bar.Items.add(, "ItemTypes", "Item Types")
            Call bar.Items.add(, "ItemFinishedQuality", "Finished Quality")
'            Call bar.Items.add(, "PaymentTerms", "Payment Terms")
'            Call bar.Items.add(, "TradeTerms", "Trade Terms")
'            If strCompany <> "All Seas" Then
                Call bar.Items.add(, "SteelList", "SteelList") '.CanClick = UserHasAccess("ExpSteelList")
'                Call bar.Items.add(, "SROList", "SRO List") '.CanClick = UserHasAccess("ExpSROList")
'                'call bar.Items.add( , "DefineBox", "Define Box Size" ) '.CanClick = UserHasAccess("ExpBoxSize")
'                'call bar.Items.add( , "SteelVenders", "Steel Venders" ) '.CanClick = UserHasAccess("ExpSteelVends")
'                'call bar.Items.add( , "FDAProcesses", "FDA Processes" ) '.CanClick = UserHasAccess("ExpFDAProcs")
                Call bar.Items.add(, "Ports", "Ports") '.CanClick = UserHasAccess("ExpPorts")
'
'                Call bar.Items.add(, "GSPCodes", "GSP Codes") '.CanClick = UserHasAccess("ExpGSPCodes")
'                Call bar.Items.add(, "MeasuringUnits", "Measuring Units")
                Call bar.Items.add(, "ExchangeRates", "Exchange Rates") '.CanClick = UserHasAccess("ExpGSPCodes")
'                'Call bar.Items.add(, "Stores", "Stores")
'
'                If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                    Call bar.Items.add(, "AddStoresRacksBins", "Add Racks/Bins")
'                End If
'
'            End If
'
'            If strCompany = "Dr-Frgz" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
'                Call bar.Items.add(, "ImportPictures", "Import Pictures")
'            End If
'
'            Call bar.Items.add(, "Security", "Security")
            
        Case "Customers"
            Call bar.Items.add(, "FCustomers", "Foreign Customers") '.CanClick = UserHasAccess("ExpForeignCusts")
            If strCompany = "QEL" Then
                Call bar.Items.add(, "MarketingCustomers", "Marketing Customers")
            End If
            If strCompany <> "All Seas" Then
                Call bar.Items.add(, "CustomerCatalog", "Customer Catalog") '.CanClick = UserHasAccess("ExpCustCatalog")
            End If
            If strCompany = "IAA" Or strCompany = "Kami" Then
                Call bar.Items.add(, "CustomerComparison", "Customer Comp.")
            End If
        
        Case "Orders"
        
            Call bar.Items.add(, "QuotationList", "Quotation List") '.CanClick = UserHasAccess("ExpOrderList")
            
            
            Call bar.Items.add(, "NewOrder", "New Order") '.CanClick = UserHasAccess("ExpNewOrder")
            Call bar.Items.add(, "OrderList", "Order List") '.CanClick = UserHasAccess("ExpOrderList")
            Call bar.Items.add(, "OrderListCustomers", "Cust Order List")
            Call bar.Items.add(, "OrderListStock", "Stock && Parts Order List")
            
            Call bar.Items.add(, "PrintLabels", "Print Labels")
            Call bar.Items.add(, "AdvancePayments", "Advance Payments") '.CanClick = UserHasAccess("ExpOrderList")
            If Not (strCompany = "Banzai" Or strCompany = "QEL" Or strCompany = "Instrumed") Then
                Call bar.Items.add(, "OrderItemList", "Order Item List")
            End If
            If strCompany = "IAA" Or strCompany = "Kami" Then
                Call bar.Items.add(, "OrderItemBalances", "Order Item Balances")
            End If
                'call bar.Items.add( , "NewStockOrder", "New Stock Order" ) '.CanClick = UserHasAccess("ExpNewStockOrder")
'            If Not (strCompany = "Instrumed") Then
'                Call bar.Items.add(, "OrderStatus", "Order Status")
'            End If
            Call bar.Items.add(, "ArticleWiseStatus", "Articlewise Status")
            
        Case "ProformaInvoice"
            Call bar.Items.add(, "NewProforma", "New Proforma") '.CanClick = UserHasAccess("ExpNewPInvoice")
            Call bar.Items.add(, "PInvoiceList", "Proforma Invoice List") '.CanClick = UserHasAccess("ExpPInvoiceList")
            
        Case "Agents"
            Call bar.Items.add(, "ShippingAgents", "Shipping Agents") '.CanClick = UserHasAccess("ExpShippingAgents")
            Call bar.Items.add(, "InsuranceAgents", "Insurance Agents") '.CanClick = UserHasAccess("ExpInsuranceAgents")
        Case "CustomDocuments"
            'call bar.Items.add( , "GenFormE", "Generated Form E Nos." ) '.CanClick = UserHasAccess("ExpCustomGenFormE")
            Call bar.Items.add(, "NewCustomInvoice", "New Custom Invoice") '.CanClick = UserHasAccess("ExpCustomInvoice")
            'Call bar.Items.add(, "NewCustomInvoiceEx", "New Custom Invoice (Ex)") '.CanClick = UserHasAccess("ExpCustomInvoice")
            Call bar.Items.add(, "CustomInvoice", "Custom Invoice List") '.CanClick = UserHasAccess("ExpCustomInvoice")
            
            Call bar.Items.add(, "SampleInvoiceList", "Sample Invoice List")
            
            If Not (strCompany = "Instrumed") Then
                Call bar.Items.add(, "NewPackingListM", "New Packing List (Manual)")
                Call bar.Items.add(, "CustomPackingList", "Packing List")
            End If
            
            Call bar.Items.add(, "CustomLabels", "Packing Outer Labels") '.CanClick = UserHasAccess("ExpPackingLabels")
            Call bar.Items.add(, "PrintInnerLabels", "Print Inner Labels") '.CanClick = UserHasAccess("ExpPackingLabels")
            Call bar.Items.add(, "CustomShipping", "Shipping Instruction") '.CanClick = UserHasAccess("ExpCustomShipping")
            Call bar.Items.add(, "PrintValuationForm", "Print Valuation Form") '.CanClick = UserHasAccess("ExpValuationForm")
            'call bar.Items.add( , "CustomCovering", "Covering Letter" ) '.CanClick = UserHasAccess("ExpCustomCovering")
            Call bar.Items.add(, "FormE", "Form E") '.CanClick = UserHasAccess("ExpEForm")
            'call bar.Items.add( , "HidePackage", "Hide Package" ) '.CanClick = UserHasAccess("ExpHidePackage")
            Call bar.Items.add(, "GSP", "GSP")
            Call bar.Items.add(, "CertofOrigin", "Certificate of Origin")
        Case "ComDocuments"
            
            Call bar.Items.add(, "ComInvoice", "Commercial Invoice") '.CanClick = UserHasAccess("ExpComInvoice")
            If Not (strCompany = "Instrumed") Then
                Call bar.Items.add(, "ComPackingList", "Packing List")
            End If
            'Call bar.Items.add(, "GSP", "G S P") '.CanClick = UserHasAccess("ExpGSP")
            Call bar.Items.add(, "ComCovering", "Covering Letter") '.CanClick = UserHasAccess("ExpComCovering")
'            call bar.Items.add( , "FreightRate", "Enter Freight Rate" ) '.CanClick = UserHasAccess("ExpEnterRate")
            'call bar.Items.add( , "CustomerLetter", "Customer Letter" ) '.CanClick = UserHasAccess("ExpCustomerLedger")
        Case "cGMP"
            Call bar.Items.add(, "DHRWiz", "DHR Wizard") '.CanClick = UserHasAccess("ExpDHRWiz")
            Call bar.Items.add(, "DHR", "D H R") '.CanClick = UserHasAccess("ExpDHR")
            Call bar.Items.add(, "SteelReq", "Steel Requisition") '.CanClick = UserHasAccess("ExpSteelReq")
            Call bar.Items.add(, "CalculationSheet", "Calculation Sheet") '.CanClick = UserHasAccess("ExpCalcSheet")
        Case "BankDocuments"
            Call bar.Items.add(, "PrintInvoice", "Bank Invoice") '.CanClick = UserHasAccess("ExpPrintInvoice")
            Call bar.Items.add(, "PrintPList", "Bank Packing List") '.CanClick = UserHasAccess("ExpPrintInvoice")
            Call bar.Items.add(, "LetterToBank", "Letter To Bank") '.CanClick = UserHasAccess("ExpLetterToBank")
            Call bar.Items.add(, "LetterToBankSimple", "Letter To Bank (Simple)")
            Call bar.Items.add(, "BankSalesContract", "Bank Sales Contract")
            Call bar.Items.add(, "BIReasonForDelay", "Reason For Delay")
            Call bar.Items.add(, "BIReasonForDelaySBP", "Reason For Delay (SBP)")
            Call bar.Items.add(, "BILabReport", "Lab Report")
        Case "PaymentStatus"
            Call bar.Items.add(, "CustPayStatus", "Payment Status")
            Call bar.Items.add(, "RecCustPay", "Receive Payments")
            If Not (strCompany = "Banzai" Or strCompany = "QEL" Or strCompany = "Instrumed") Then
                Call bar.Items.add(, "AgingReport", "Aging Report")
            End If
'            If strCompany <> "Banzai" Then
'                Call bar.Items.add(, "CustomerPayStatus", "Customer Payment Status")
'                Call bar.Items.add(, "RecCustomerPay", "Rec. Customer Payments")
'            End If
        Case "Rebate"
            Call bar.Items.add(, "RebateDocuments", "Rebate Documents") '.CanClick = UserHasAccess("ExpRebateDocs")
            Call bar.Items.add(, "RebateStatus", "Rebate Status") '.CanClick = UserHasAccess("ExpRebateStatus")
        
        Case "Satistics"
        
            Call bar.Items.add(, "StatTotalExport", "1. Total Export")
            Call bar.Items.add(, "StatTotalDemand", "1. Total Demand")
            Call bar.Items.add(, "StatArticleWiesSales", "2. Articlewise Sales")
            Call bar.Items.add(, "StatArticleWiesOrders", "3. Articlewise Demand")
               
            If (strCompany = "IAA" Or strCompany = "Kami") Then
                Call bar.Items.add(, "StatExportPerformanceReport", "4. Itemwise Export Performance")
                bar.Items.add , "PrdYearlyDemand", "5. Customer Demand Report Current Year (Qty)"
                bar.Items.add , "PrdYearlyDemandAmt", "6. Customer Demand Report Current Year (Amt)"
                bar.Items.add , "PrdYearlyGroupwiseDemand", "7. Current Year monthwise demand"
                bar.Items.add , "PrdYearlyDemandSupply", "8. Yearly Demand/Supply"
                bar.Items.add , "PrdYearlyDemandSupplyGraph", "9. Demand/Supply Comparison Graph Qty wise"
                bar.Items.add , "PrdYearlyDemandSupplyGraphAt", "10. Demand/Supply Comparison Graph Amt wise"
                'bar.Items.add , "PrdMainGroupDemandSupply", "11. Main Group Demand/Supply"
                'bar.Items.add , "PrdSubGroupDemandSupply", "12. Sub-Group Demand/Supply"
                bar.Items.add , "PrdItemDemandSupply", "13. Main Group/Sub-Group Item Level D&&S"
                'bar.Items.add , "PrdCustomerWiseDemandSupply", "14. Customerwise Demands"
                bar.Items.add , "PrdTotalExportValueShare", "15. Total Export Value Share"
                'bar.Items.add , "PrdPeriodwiseDemand", "16. Quarterly Demand"
                bar.Items.add , "PrdDeliveryPerformanceReport", "17. Delivery Performance Report"
                bar.Items.add , "PrdDeliveryPerformanceDetailReport", "18. Delivery Performance Detail"
                
                bar.Items.add , "PrdCustomerPurchaseHistory", "21. Customer Purchase History"
                bar.Items.add , "PrdItemwiseYearlyDemand", "22. Itemwise Yearly Demand"
                bar.Items.add , "PrdItemwiseYearlyDemandX2", "23. Itemwise Yearly Demand (X2)"
                bar.Items.add , "PrdItemwiseYearlyDemand_Yearwise", "24. Itemwise Yearwise Demand"
                bar.Items.add , "PrdYearlyCustomerWiseDemands", "25. Yearly Customerwise Demands"
                
                bar.Items.add , "PrdProductLineDemand", "26. Product Line Demand"
                bar.Items.add , "PrdProductGroupwiseDemand", "27. Product Groupwise Demand"
                bar.Items.add , "PrdProductLineSupply", "28. Product Line Supply"
                bar.Items.add , "PrdProductGroupwiseSupply", "29. Product Groupwise Supply"
                bar.Items.add , "PrdTotalDemandValueShare", "30. Total Demand Value Share"
                bar.Items.add , "PrdAdvancePlan", "31. Advance Plan"
            End If
            If strCompany = "Dr-Frgz" Then
                bar.Items.add , "PrdCustomerStatusValueWise", "Customer Analysis"
            End If
            'call bar.Items.add( , "StatPaymentLedger", "Payment Ledger"
'            call bar.Items.add( , "StatExportStats", "Export Statistics"
        Case "EEStatement"
'            call bar.Items.add( , "EETotalExport", "Total Export"
'            call bar.Items.add( , "AddBankNos", "Add Bank Nos."
            Call bar.Items.add(, "EE", "Form EE") '.CanClick = UserHasAccess("EE")
        Case "Others"
            Call bar.Items.add(, "MiscSentItems", "Misc. Sent Items")
        Case "UserManagement"
            Call bar.Items.add(, "UserManagement", "User Management")
            Call bar.Items.add(, "UserActivity", "User Activity")
            
        
    End Select
    'If bar.Items.Count > 5 Then bar.WatermarkPicture = HICOLogo.Picture
    
    For i = 1 To bar.Items.count
        bar.Items(i).CanClick = UserHasAccessEX(bar.Items(i).key)
    Next
    
    
End Sub

Private Sub ExpBar_ItemClick(ITM As vbalExplorerBarLib6.cExplorerBarItem)
    Call ShowLinks(ITM.key)
End Sub

Private Sub HICOLogo_Click()
    'Show About Form
End Sub

Private Sub ImgTitle_Click()
    Me.WindowState = vbMinimized
End Sub


Private Sub Label1_Change()
    'Call MDIForm_Resize
End Sub

Private Sub label1_Click()
    'show about form
End Sub

Private Sub ExpBar_ItemRightClick(ITM As vbalExplorerBarLib6.cExplorerBarItem)

    Dim strRet As String
    strRet = ShowPopUpMenu1("Add to Favourites")
    
    If strRet = "Add to Favourites" Then
        Dim lCount As Long
        lCount = GetSingleLongValue("COUNT(*)", "User_FavouriteLinks", " WHERE UserID=" & UserID & " AND OptionID='" & ITM.key & "'")
        If lCount > 0 Then
            MsgBox "Already Added.", vbInformation
            Exit Sub
        Else
            con.Execute "INSERT INTO User_FavouriteLinks(UserID,OptionID) VALUES(" & UserID & ",'" & ITM.key & "')"
            Call CreateExplorerBarForFavourites
        End If
    End If
    
End Sub

Private Sub ExpBarFavourites_ItemClick(ITM As vbalExplorerBarLib6.cExplorerBarItem)
    Call ShowLinks(ITM.key)
End Sub

Private Sub ExpBarFavourites_ItemRightClick(ITM As vbalExplorerBarLib6.cExplorerBarItem)
    
    Dim strRet As String
    strRet = ShowPopUpMenu1("Remove from Favourites")
    
    If strRet = "Remove from Favourites" Then
        con.Execute "DELETE FROM User_FavouriteLinks WHERE UserID=" & UserID & " AND OptionID='" & ITM.key & "'"
        Call CreateExplorerBarForFavourites
    End If
    
End Sub

Private Sub lblAlerts_Click()
    
    If lblAlerts.Caption = "[Click to Load]" Then
        Call ShowPurchaseCalendarAlerts(True)
        Exit Sub
    End If
    
    If strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
        Load frmAlertsDashBoard
        frmAlertsDashBoard.ShowMe
    Else
        Dim f As New frmAlertsDetails
        Load f
        Call f.ShowAlerts(0)
    End If
    
End Sub

Private Sub Logo_DblClick()

    On Error GoTo err
    
    If CurrentUserName <> "Administrator" Then Exit Sub
    On Error Resume Next
    Dim strPicfile As String
    CD1.ShowSave
    
    If Not (err.Number = 32755) Then
        strPicfile = CD1.FileName
        On Error GoTo 0
        Logo.Picture = LoadPicture(strPicfile)
        
        Dim Rec As New ADODB.Recordset
        Dim rsPic As New ADODB.Stream
        
        Rec.Open "SELECT EntryID,CompanyLogo FROM Company", con, adOpenStatic, adLockOptimistic
    
        If Logo.Picture <> 0 And Trim(strPicfile) <> "" Then
            rsPic.Open
            rsPic.Type = adTypeBinary
            rsPic.LoadFromFile strPicfile
            Rec.Fields(1) = rsPic.Read
            Rec.Update
            rsPic.Close
        End If
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub LVFavourites_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call ShowLinks(Item.Tag)
End Sub

Private Sub lblCustomerEvents_Click()
    Dim f As New frmAlertsDetails
    Load f
    If lblCustomerEvents.Tag = "Planning" Then
        Call f.ShowAlerts(3)
    Else
        Call f.ShowAlerts(2)
    End If
End Sub

Private Sub lblLateOrderAlert_Click()
    Dim f As New frmAlertsDetails
    Load f
    Call f.ShowAlerts(1)
End Sub

Public Sub MainToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim i As Integer
    For i = 1 To MainToolbar.Buttons.count
        If MainToolbar.Buttons(i).Style = tbrCheck Then
            If MainToolbar.Buttons(i).key <> Button.key Then
                MainToolbar.Buttons(i).Value = tbrUnpressed
            End If
        End If
    Next
    
'    For i = 0 To Forms.count
'        On Error Resume Next
'        If Forms(i).Name = "frmProductionLinks" Then
'            Unload Forms(i)
'            Exit For
'        End If
'    Next
    Select Case Button.key
        Case "Company"
            Call CreateExplorerBarForCompany
        Case "Export"
            Call CreateExplorerBarForExport
        Case "Import"
            Call CreateExplorerBarForImport
        Case "Financial"
            Call CreateExplorerBarForAccounts
        Case "Payroll"
            Call CreateExplorerBarForPayroll
        Case "Stock"
            Call CreateExplorerBarForStock
        Case "Production"
            Call CreateExplorerBarForProduction
        Case "QMS"
            Call CreateExplorerBarForQMS
        Case "FixedAssets"
            Call CreateExplorerBarForFixedAssets
        Case "Sampling"
            Call CreateExplorerBarForSampling
'        Case "QMS"
'            Call CreateExplorerBarForQMS
'            Dim f As New frmProductionLinks
'            Load f
'
'            f.ShowMe
        Case "ChangePassword"
            With frmUserInfo
                .Show_Option = Opt_ChangeUserPassword
                .Tag = CurrentUserName
                .Show 1
            End With
            Exit Sub
        Case "Dashboard"
            Call CreateExplorerBarForDashBoard
        Case "Help"
            Call CreateExplorerBarForHelp
        Case "Exit"
            If MsgBox("Are you sure to Quit ?", vbQuestion + vbYesNo) = vbYes Then
                On Error Resume Next
                If con.State = 1 Then con.Close
                Set con = Nothing
                End
            End If
            Exit Sub
    End Select
    Picture1.Visible = True
    PicShowExpBar.Visible = False
End Sub

Private Sub MDIForm_Activate()
    Call Timer3_Timer
End Sub

Private Sub MDIForm_DblClick()
    
'    Dim lTBBackColor As Long
'    lTBBackColor = vbWhite 'CreateSolidBrush(RGB(240, 120, 120))        'Creates the background from a Color (Long)
'    ChangeTBBack MainToolbar, lTBBackColor, enuTB_STANDARD
'    InvalidateRect 0&, 0&, False
'
'    Exit Sub
    On Error Resume Next
    
    If CurrentUserName <> "Administrator" Then Exit Sub
        Dim strPicfile As String
        CD1.ShowSave
        
        If Not (err.Number = 32755) Then
            strPicfile = CD1.FileName
        frmMDIBG.Hide
        On Error GoTo 0
        frmMDIBG.Image1.Picture = LoadPicture(strPicfile)
        
        Dim Rec As New ADODB.Recordset
        Dim rsPic As New ADODB.Stream
        
        Rec.Open "SELECT EntryID,SoftBG1 FROM Company", con, adOpenStatic, adLockOptimistic
    
        If frmMDIBG.Image1.Picture <> 0 And Trim(strPicfile) <> "" Then
            rsPic.Open
            rsPic.Type = adTypeBinary
            rsPic.LoadFromFile strPicfile
            Rec.Fields(1) = rsPic.Read
            Rec.Update
            rsPic.Close
            'Now Refresh the form...
            Call frmMDIBG.Form_Resize
            frmMDIBG.Move (Me.ScaleWidth - frmMDIBG.ScaleWidth) / 2, (Me.ScaleHeight - frmMDIBG.ScaleHeight) / 2
            frmMDIBG.Visible = True
        End If
    
    Exit Sub
err:
    MsgBox err.Description
    End If
End Sub

Private Sub MDIForm_Load()
    
    Dim bVisible As Boolean
    If strCompany = "Dr-Frgz" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
        bVisible = True
    Else
        bVisible = False
    End If
    
'    Dim iHelpIndex As Integer, iQMSIndex As Integer
'    iHelpIndex = MainToolbar.Buttons("Help").Index
'    MainToolbar.Buttons(iHelpIndex).Visible = bVisible
'    MainToolbar.Buttons(iHelpIndex + 1).Visible = bVisible
    
    iQMSIndex = MainToolbar.Buttons("QMS").Index
    If strCompany = "QEL" Or strCompany = "Banzai" Then
        MainToolbar.Buttons(iQMSIndex).Visible = False
        MainToolbar.Buttons(iQMSIndex + 1).Visible = False
    End If
    
'    If strCompany <> "QEL" Then
'        'Do Nothing
'    Else
'        Call ShowLateOrderAlerts
'    End If
    
    'Background
    Load frmMDIBG
    frmMDIBG.Hide
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT SoftBG1 FROM Company", con, adOpenForwardOnly, adLockReadOnly
        If Not IsNull(!SoftBG1) Then
            Set frmMDIBG.Image1.DataSource = rs
            frmMDIBG.Image1.DataField = "SoftBG1"
        End If
        .Close
    End With
    Set rs = Nothing
    
    If frmMDIBG.Image1.Picture <> 0 Then
        frmMDIBG.Visible = True
        frmMDIBG.Move (Me.ScaleWidth - frmMDIBG.ScaleWidth) / 2, (Me.ScaleHeight - frmMDIBG.ScaleHeight) / 2
    End If
    'Background
   
    On Error GoTo err
    'LoadProductionProcesses
    'mnuPrdMain(0).Visible = False
    Dim i As Integer
    Set MainToolbar.ImageList = ImageList5
    For i = 1 To MainToolbar.Buttons.count
        With MainToolbar.Buttons(i)
            If .Style = tbrDefault Then
                .Image = ImageList5.ListImages(.key).Index
            End If
        End With
    Next
    
    Export_Package = True
    
    Call GetWinVersionInfo
    
    gHW = Me.hwnd
    'Call Hook
    
    On Error Resume Next
    
    
    
    With rs
'        .Open "SELECT CompanyLogo FROM Company", con, adOpenForwardOnly, adLockReadOnly
'        If Not IsNull(!CompanyLogo) Then
'            Set Logo.DataSource = rs
'            Logo.DataField = "CompanyLogo"
'        End If
'        .Close
    End With
    Set rs = Nothing
    
    'Dim ctr As Control
    'For Each ctr In Controls
    '    If TypeOf ctr Is Menu Then
    '        ctr.Visible = False
    '    End If
    'Next ctr

    On Error GoTo err
    
    'Hide The Title Bar  Of Main Form.
    SetWindowLong Me.hwnd, -16, WS_DLGFRAME
    
    'PicBar.Visible = False
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub LoadProductionProcesses()
    
'    On Error GoTo err
'    Dim rs As New ADODB.Recordset, i As Long, iNewIndex As Integer
'    i = 0
'    iNewIndex = 0
'    With rs
'        .Open "SELECT * FROM Processes ORDER BY ProcessID", con, adOpenForwardOnly, adLockReadOnly
'        Do Until .EOF
'            If i > 0 Then
'                iNewIndex = i + 1
'                Load mnuPrdMain(iNewIndex)
'                mnuPrdMain(iNewIndex).Visible = True
'            End If
'            mnuPrdMain(iNewIndex).Caption = !Description & ""
'            'Now Add Subs..0 for Issuance & 1 for Receiving
'            i = i + 1
'            .MoveNext
'        Loop
'        .Close
'    End With
'    Set rs = Nothing
'    Exit Sub
'err:
'    MsgBox err.Description
End Sub

Private Sub MDIForm_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If CurrentUserName <> "Administrator" Then Exit Sub
    If Button <> 2 Then Exit Sub
    
    If y <= 180 And x >= (Me.ScaleWidth - PicFavourites.Width) Then
        If getPasswordNew("awM") Then
            Call ExpBar.Bars("Company").Items.add(, "CompaniesList", "Companies")
        End If
    End If
   
End Sub

Private Sub MDIForm_QueryUnload(cancel As Integer, UnloadMode As Integer)
    
    On Error Resume Next
    'rsUserSettings.Close
    Set rsUserSettings = Nothing
    Set con = Nothing
    
End Sub

Private Sub MDIForm_Resize()

    On Error Resume Next
    'ImgAlerts.Move PicAlerts.Width - ImgAlerts.Width
    
    'lblAlertsHead.Move ImgAlerts.Left - lblAlertsHead.Width
    
    'lblAlerts.Width = PicAlerts.Width - lblAlerts.Left - lblAlertsHead.Width - 200
    
    frmMDIBG.Move (Me.ScaleWidth - frmMDIBG.ScaleWidth) / 2, (Me.ScaleHeight - frmMDIBG.ScaleHeight) / 2
   
    'Label1.Top = Picture1.Height - 300
    'Logo.Visible = False
    'Logo.Top = Label1.Top - Logo.Height - 30
    'ExpBar.Height = Picture1.Height - Logo.Height - Label1.Height - 50
     
    cmdShowExpBar.Height = PicShowExpBar.Height
    cmdHideExpBar.Height = PicShowExpBar.Height
    ExpBar.Height = PicShowExpBar.Height
    
    
    cmdShowFavourites.Height = PicShowExpBar.Height
    cmdHideFavourites.Height = PicShowExpBar.Height
    ExpBarFavourites.Height = PicShowExpBar.Height
    
    'SWF.Move 0, 0, PicTitle.Width, PicTitle.Height
End Sub

Private Sub MDIForm_Unload(cancel As Integer)
    On Error Resume Next
   ' Con.RollbackTrans
    'If Con.State = 1 Then Con.Close
    Call Unhook
    Set con = Nothing
End Sub

Private Sub mnuAssItems_Click()
    Call ContractorList.ShowItems
End Sub

Private Sub mnuAuth_Click()
    frmAuthorizeIssRcv.Show
End Sub

Private Sub mnuCalc_Click()

    On Error GoTo err
    Call Shell("calc.exe")
    Exit Sub
err:
    MsgBox "Cannot Run Calculater Due To Error: " & vbNewLine & err.Description, vbInformation

End Sub

Private Sub mnuComp_Click()
    frmComponents.Show 1
End Sub

Private Sub mnuCompCatal_Click()
    frmItems.Show
End Sub

Private Sub mnuCompositions_Click()
    frmCompositions.Show
End Sub

Private Sub mnuCustCatal_Click()
    frmCustCatal.Show
End Sub

Private Sub mnuCustList_Click()
    frmFCustomerList.Show
End Sub

Private Sub mnuCustListrpt_Click()
    Call ShowCustListrpt
End Sub

Private Sub mnuCustOrdList_Click()
    frmFCustOrderList.Show
End Sub

Private Sub mnuCustProformaList_Click()
    frmFProformaList.Show
End Sub

Private Sub mnuDailyRpts_Click()
    frmDailyIssRcvRpt.Show
End Sub


Private Sub mnuDeptInact_Click()
    If mnupopdept.Tag = "EmpDept" Then
        Call DeptList.MakeInactive
    Else
        Call ContDeptList.MakeInactive
    End If
End Sub

Private Sub mnuDeptList_Click()
    DeptList.Show
End Sub

Private Sub mnuDeptnew_Click()
    If mnupopdept.Tag = "EmpDept" Then
        Call DeptList.mnuAddDept
    Else
        Call ContDeptList.mnuAddDept
    End If
End Sub

Private Sub mnuDeptPrint_Click()
    If mnupopdept.Tag = "EmpDept" Then
        Call DeptList.PrintDeptList
    Else
        'Call ContDeptList.PrintDeptList
    End If
End Sub

Private Sub mnuDeptRefresh_Click()
    If mnupopdept.Tag = "EmpDept" Then
        Call DeptList.RefreshIt
    Else
        Call ContDeptList.RefreshIt
    End If
End Sub


Private Sub mnuEmpPrint_Click()
    If mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.PrintEmpList
    Else
          'Call ContractorList.PrintEmpList
    End If
End Sub

Private Sub mnuempRefresh_Click()
    If mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.RefreshIt
    Else
        Call ContractorList.RefreshIt
    End If
End Sub
Private Sub mnuExit_Click()
    If con.State = 1 Then con.Close
    Set con = Nothing
    End
End Sub

Private Sub mnuFDAProcess_Click()
    frmFDAProcesses.Show 1
End Sub


Private Sub mnuForeignCatal_Click()
    frmForeignCatalog.Show 1
End Sub

Private Sub mnuGuage_Click()
    frmGages.Show 1
End Sub


Private Sub mnuHICOResources_Click()
    frmHICOR.Show
End Sub

Private Sub mnuHolidays_Click()
    Load Holidays
    'Holidays.Move 3000, 500
    With Holidays
        .Move (Me.Width - .Width) / 2, Me.Top + 500
        .Show
    End With
End Sub

Private Sub mnuHRGroups_Click()
    Dim f As New frmItemGroups
    Load f
    Call f.ShowForm("HRGroups", "Hico Resources Groups", "Hico Resource Name")
End Sub



Private Sub mnuIssHold_Click()
    frmVendRcvItemsHnD.Show
End Sub


Private Sub mnuIssRepair_Click()
    frmVendIssRepair.Show
End Sub

Private Sub mnuIssRepairPoli_Click()
    frmVendIssRepair.Show
End Sub

Private Sub mnuIssRM_Click()
    frmRMIss.Show
End Sub



Private Sub mnuItemsFinQuality_Click()
    Dim f As New frmMiscS
    Load f
    Call f.ShowForm("ItemsFinQuality", "Items Fin. Quality", "Finishing Quality")
End Sub

Private Sub mnuItemTypes_Click()
    Dim f As New frmMiscS
    Load f
    Call f.ShowForm("ItemTypes", "Item Types", "Item Types")
End Sub

Private Sub mnuItmGrps_Click()
    Dim f As New frmItemGroups
    Load f
    Call f.ShowForm("ItemGroups", "Item Groups", "Group Name")
End Sub

Private Sub mnuLabs_Click()
    frmTestLabs.Show
End Sub

Private Sub mnuLogoff_Click()

    If MsgBox("Are You Sure To Log Off ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Unload MainForm
    LoggingOff = True
    Call Main
    
End Sub

Private Sub mnuLotProc_Click()
    frmInitialLotProcess.Show
End Sub


Private Sub mnuMake_Click()

    If mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.MakeInactive
    Else
        Call ContractorList.MakeInactive
    End If
    
End Sub

Private Sub mnuMatList_Click()
    frmMaterial.Show
End Sub

Private Sub mnuNewOrd_Click()
    frmFNewCustOrder.Show
End Sub

Private Sub mnuNewProforma_Click()
    frmFNewProforma.Show
End Sub

Private Sub mnuNewStockOrder_Click()
    frmFNewStockOrder.Show
End Sub

Private Sub mnuOrdStatus_Click()
    frmFCustOrdStatus.Show
End Sub

Private Sub mnuOrdStatus1_Click()
    frmFCustOrdStatus1.Show
End Sub

Private Sub mnuOrdStatusComp_Click()
    frmFCustOrdStatusComp.Show
End Sub


Private Sub mnuPassword_Click()
    
    With frmUserInfo
        .Show_Option = Opt_ChangeUserPassword
        .Tag = CurrentUserName
        .Show 1
    End With
    
End Sub

Private Sub mnuPer_Click()
    frmPercentRanges.Show
End Sub

Private Sub mnuPOrderList_Click()
    frmRMPOrdList.Show
End Sub


Private Sub mnuPorts_Click()
    frmPorts.Show
End Sub

Private Sub mnuProcess_Click()
    frmProcesses.Show 1
End Sub


Private Sub mnuReqForm_Click()
    'frmRequisition.Show
    frmGMPs.Show
End Sub

Private Sub mnuProdPlan_Click()
    frmProdPlan.Show
End Sub

Private Sub mnuPymtTerms_Click()
    Dim f As New frmMiscS
    Load f
    Call f.ShowForm("PaymentTerms", "Payment Terms", "Payment Terms")
End Sub

Private Sub mnuRcvRepair_Click()
    frmVendRcvRepair.Show
End Sub

Private Sub mnuRcvRepairPoli_Click()
    frmVendRcvRepair.Show
End Sub

Private Sub mnuRcvRepairWastage_Click()
    frmVendRcvRepairWaste.Show
End Sub


Private Sub mnuRecRM_Click()
    frmRMRcvItems.Show
End Sub


Private Sub mnuReq_Click()
    'frmRequisition.Show
    frmGMPs.Show
End Sub

Private Sub mnuRM_Click()
    frmRM.Show
End Sub

Private Sub mnuRMAssItems_Click()
    frmRMVendAssItems.Show
End Sub

Private Sub mnuRMGroups_Click()
    Dim f As New frmItemGroups
    Load f
    Call f.ShowForm("RMGroups", "RM Groups", "RM Name")
End Sub


Private Sub mnuRMPO_Click()
    frmRMPurchOrd.Show
End Sub

Private Sub mnuRMVendList_Click()
    frmRMVenderList.Show
End Sub


Private Sub mnuShowDept_Click()
    If mnupopdept.Tag = "EmpDept" Then
        Call DeptList.ShowInactive(mnushowDept)
    Else
        Call ContDeptList.ShowInactive(mnushowDept)
    End If
End Sub

Private Sub mnuShowInactive_Click()
    If mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.ShowInactive(mnuShowInactive)
    Else
        Call ContractorList.ShowInactive(mnuShowInactive)
    End If
End Sub

Private Sub mnuSteel_Click()
    frmSteelList.Show
End Sub

Private Sub mnuSteelVends_Click()
    frmSteelVends.Show
End Sub

Private Sub mnuSubGeneral_Click(Index As Integer)

    Select Case Index
        Case 1
            frmInitialLotProcess.Show
        Case 3
            frmPercentRanges.Show
        Case 5
            frmComponents.Show 1
        Case 7
            frmProcesses.Show 1
        Case 8
            frmFDAProcesses.Show 1
        Case 10
            frmSteelVends.Show
        Case 11
            frmTestLabs.Show
        Case 13
            frmUnits.Show
        Case 14
            frmCompositions.Show
        Case 16
            frmPorts.Show
        Case 19
            Call mnuPymtTerms_Click
    End Select
    
End Sub

Private Sub mnuSubOptions_Click(Index As Integer)

    Select Case Index
        Case 1
            Call mnuCalc_Click
        Case 2
            frmUsers.Show 1
        Case 4
            Call mnuPassword_Click
        Case 5
            Call mnuLogoff_Click
        Case 7
            Call mnuExit_Click
    End Select
    
End Sub

Private Sub mnuTaxRanges_Click()
    Load TaxRanges
    'Holidays.Move 3000, 500
    With TaxRanges
        .Move (Me.Width - .Width) / 2, Me.Top + 500
        .Show
    End With
End Sub


Private Sub mnuUnits_Click()
    frmUnits.Show
End Sub

Private Sub mnuUserManager_Click()
    frmUsers.Show 1
End Sub

Public Sub ApplySecuritySettings(UserID As Long)
   
    On Error GoTo err
    
    'General Settings...
    mnuSubGeneral(1).Enabled = UserHasAccess("SetInitLotProc", UserID)
    mnuSubGeneral(3).Enabled = UserHasAccess("PercentRanges")
    mnuSubGeneral(5).Enabled = UserHasAccess("Components")
    mnuSubGeneral(7).Enabled = UserHasAccess("Processes")
    mnuSubGeneral(8).Enabled = UserHasAccess("FDAProcesses")
    mnuSubGeneral(10).Enabled = UserHasAccess("SteelVenders")
    mnuSubGeneral(11).Enabled = UserHasAccess("TestLabs")
    mnuSubGeneral(13).Enabled = UserHasAccess("MeasuringUnits")
    mnuSubGeneral(14).Enabled = UserHasAccess("Compositions")
    mnuSubGeneral(16).Enabled = UserHasAccess("CitiesNPorts")
    mnuSubGeneral(18).Enabled = UserHasAccess("PymtTerms")
    
    'Venders
    'mnuVendList.Enabled = UserHasAccess("VenderList")
    'mnuVendAssItems.Enabled = UserHasAccess("VendAssItems")
    mnuVendPymt.Enabled = UserHasAccess("VendAdvPymt")
    mnuVendPaySheet.Enabled = UserHasAccess("VendPymtSheet")
    mnuVendAdvShort.Enabled = UserHasAccess("NewShortTerm")
    mnuVendAdvLong.Enabled = UserHasAccess("NewLongTerm")
    mnuVendShortLoanLedger.Enabled = UserHasAccess("ShortTermLedger")
    mnuVendLongLoanLedger.Enabled = UserHasAccess("LongTermLedger")
    'mnuVendLedger.Enabled = UserHasAccess("VendLedger")
    
    'Raw Material
    mnuRMGroups.Enabled = UserHasAccess("RMGroups")
    mnuRM.Enabled = UserHasAccess("RawMaterial")
    'mnuRMVendList.Enabled = UserHasAccess("RMVendList")
    'mnuRMAssItems.Enabled = UserHasAccess("RMVAssItems")
    mnuRMPO.Enabled = UserHasAccess("RMVNewPO")
    mnuPOrderList.Enabled = UserHasAccess("RMVPOList")
    mnuRecRM.Enabled = UserHasAccess("RMRcvRM")
    mnuIssRM.Enabled = UserHasAccess("RMIssRM")
    
    'HICO Resources
    mnuHRGroups.Enabled = UserHasAccess("HICOResGroups")
    'mnuHICOResources.Enabled = UserHasAccess("HICOResources")
    'mnuVendAssR.Enabled = UserHasAccess("AssHICORes")
        
    'Company
    'mnuCompDeatil.Enabled = UserHasAccess("CompanyDetail")
    'mnuCompCatal.Enabled = UserHasAccess("CompanyCatalog")
    mnuItmGrps.Enabled = UserHasAccess("ItemGroups")
    mnuItemsFinQuality.Enabled = UserHasAccess("ItemFinQuality")
    mnuItemTypes.Enabled = UserHasAccess("ItemTypes")
    'mnuForeignCatal.Enabled = UserHasAccess("ForeignCatalog")
    'mnuCustList.Enabled = UserHasAccess("CustomerList")
    'mnuCustCatal.Enabled = UserHasAccess("CustomerCatalog")
    mnuNewOrd.Enabled = UserHasAccess("CustNewOrder")
    mnuNewStockOrder.Enabled = UserHasAccess("StockNewOrder")
    mnuCustOrdList.Enabled = UserHasAccess("CustOrdList")
    mnuOrdStatus1.Enabled = UserHasAccess("CustOrdStatus")
    mnuNewProforma.Enabled = UserHasAccess("CustNewProforma")
    mnuCustProformaList.Enabled = UserHasAccess("CustProList")
        
    mnuReq.Enabled = UserHasAccess("Requsition")
    mnuProdPlan.Enabled = UserHasAccess("AddProdPlan")
    mnuAuth.Enabled = UserHasAccess("Authorize")
    
    'Reports
    
    'Users
    mnuSubOptions(2).Enabled = UserHasAccess("UserManagement")
    mnuSubOptions(4).Enabled = UserHasAccess("ChangePassword")
    
    
    HicoVisible = UserHasAccess("HicoVisible")
    'If UserID = 1 Then mnuAuth.Enabled = True Else mnuAuth.Enabled = False
    
    If UserID <> 1 Then Call ApplySecurityForIssRcv(UserID)
    '
    'With Toolbar1.Buttons
    '   '.Item("TakeAtt").Enabled = mnutake.Enabled
    '   '.Item("NewShort").Enabled = mnuShort.Enabled
    '   '.Item("NewLong").Enabled = mnuLong.Enabled
    '   '.Item("EmpLedger").Enabled = mnuEmpLedger.Enabled
    '   '.Item("SalSheet").Enabled = mnuSalSheet.Enabled
    'End With
    Call MDIForm_Resize
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub ApplySecurityForIssRcv(UserID As Long)
    
    For i = 1 To mnuIssRcvProcs.count
        mnuIssRcvProcs(i).Enabled = False
    Next i
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select ProcessID From UserAssProcesses Where UserID=" & UserID, con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            mnuIssRcvProcs(![ProcessID]).Enabled = True
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub
Public Sub mnuVendAdvLong_Click()

    With frmVendNewAdv
        .LongTerm = True
        .Show 1
    End With
    
End Sub

Public Sub mnuVendAdvShort_Click()
    
    With frmVendNewAdv
        .LongTerm = False
        .Show 1
    End With
    
End Sub

Private Sub mnuVendAssItems_Click()
    frmVendAssItems.Show
End Sub

Private Sub mnuVendAssR_Click()
    frmVendAssHR1.Show
End Sub

Private Sub mnuVendIssItems_Click()
    frmVendIssItems.Show
End Sub

Private Sub mnuVendIssItemsByItems_Click()
    frmVendIssItemsByItem.Show
End Sub

Private Sub mnuVendLedger_Click()
    frmVendLedger.Show
End Sub

Private Sub mnuVendList_Click()
    frmVenderList.Show
End Sub


Private Sub mnuVendLongLoanLedger_Click()
    frmVendLoanLedger.Show
End Sub

Private Sub mnuVendPaySheet_Click()
    frmVendSS.Show    'frmVendPaySheet.Show
End Sub

Private Sub mnuVendPymt_Click()
    frmVendAdvSalary.Show
End Sub

Private Sub mnuVendRcvItems_Click()
    frmVendRcvItems.Show
End Sub

Private Sub mnuVendRcvItemsFromSupervisor_Click()
    frmVendRcvItems1.Show
End Sub

Private Sub mnuVendRcvWst_Click()
    frmVendRcvWastage.Show
End Sub

Private Sub mnuVendShortLoanLedger_Click()
    frmVendShortAdvLedger.Show
End Sub

'-----------------------------------------------------------
'Code Of Pop Menu Ports.
Private Sub NewCity_Click()
    Call frmPorts.NewCity_Click
End Sub


Private Sub Rename_Click()
    Call frmPorts.Rename_Click
End Sub

Private Sub del_Click()
    Call frmPorts.del_Click
End Sub

Private Sub PicBusPart_Resize()

    Dim Temp As Long
    PicBusPart.Left = 0
    PicBusPart.Width = 1950
    
    'If PicBusPart.Height = 1995 Then 'Means Its Open
    Temp = Round(PicBusPart.Height / 350)
     'MsgBox "Total Buttons=" & Round(PicBusPart.Height / 350)
    'End If
    
    VSBusPart.Move PicBusPart.Width - VSBusPart.Width - 50, VSBusPart.Top, VSBusPart.Width, PicBusPart.Height - 15
    VSBusPart.Max = (BusPartToolBar.Buttons.count - Temp) + 1
    'VSBusPart.Left = PicBusPart.Width - VSBusPart.Width - 50
    
    'PicBusPart.Height = 1665
End Sub


Private Sub PicToolBar_Paint(Index As Integer)
    'If Index = 1 Then Stop
    'If PicToolBar(Index).Height > 15 Then
    '    ToolBar(Index).Visible = False
    '    ToolBar(Index).Visible = True
    'End If
    
End Sub



Private Sub PicTitle_DblClick()
    Call AddAllOptions
End Sub


Private Sub PopSubCust_Click(Index As Integer)
    Select Case Index
        Case 0
            frmFCustomer.Show 1
        Case 1
            frmFCustomerList.Show
        Case 3
            frmCustCatal.Show
        Case 4
    End Select
End Sub



Private Sub ProdToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 4 Then
        Me.PopUpMenu mnuProdPhases, 512, CoolBar1.Width
    End If
End Sub



'-----------------------------------------------------------
Private Sub Timer1_Timer()
On Error Resume Next


'<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
    '*************************************************
        'This Is Special Code...
        'It Will Check If User Is Back From Some Other Form.
        'If So Then It Will Do Something To Get The Buttons Back
        'Of The Toolbar.
    '*************************************************
    Dim a As String
    Dim Temp As Long
    a = Space$(255)
    Temp = GetClassName(GetForegroundWindow, a, 255)
    a = Left(a, Temp)
    
    TempForeWin = GetForegroundWindow
    If TempForeWin <> Me.hwnd And a <> "msvb_lib_monthview" Then
        'Now Check If Screen.ActiveForm.Hwnd Not In The Loaded forms Collection...
        
        On Error Resume Next
        If Not InLoadedForms(Screen.ActiveForm.hwnd) Then
            If AwayTime = "12:00:00 AM" Then
                AwayTime = ServerTime
            End If
            ForeWin = TempForeWin
        End If
        On Error GoTo 0
    Else
        If ForeWin <> Me.hwnd Then
            
            ForeWin = Me.hwnd
            If Not OurOwnForm Then
                'FrmUnlock.Show 1 'Ask For Pass
                Load frmAway
                frmAway.ShowForm
                AwayTime = "12:00:00 AM"
            End If
            
            OurOwnForm = False
            
            'Sleep (5000)
            'Call ABC
            Ready = True
        End If
    End If
    '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
    
End Sub
Private Function InLoadedForms(lHWnd As Long) As Boolean

    Dim f As Form
    InLoadedForms = False
    'Debug.Print lHwnd
    For Each f In Forms
        If f.hwnd = lHWnd Then
            InLoadedForms = True
            Exit Function
        End If
    Next
    
End Function

Private Sub Timer2_Timer()
On Error Resume Next

AwaySeconds = AwaySeconds + 1
End Sub

Private Sub Timer3_Timer()

    On Error Resume Next
    Timer3.Interval = 1000
    Dim ServerTime1 As Date
    ServerTime1 = con.Execute("Select GetDate()").Fields(0)
    ServerTime = Format(ServerTime1, "hh:mm:ss AMPM")
    If Format(ServerDate, "dd-MMM-yyyy") <> Format(ServerTime1, "dd-MMM-yyyy") Then ServerDate = Format(ServerTime1, "dd-MMM-yyyy")
    'lblTime.Caption = Now
    'SWF.SetVariable "lblTime", Format(ServerTime1, "dddd, MMMM dd, yyyy (HH:nn AMPM)")
    lblTime.Caption = Format(ServerTime1, "dd-MMM-yy hh:mm AM/PM")
    
End Sub


'Private Sub company(ByVal Button As MSComctlLib.Button, MenuTop As Long)
'
'    'Error Traping Of Toolbar_ButtonClick Will be Used.
'
'    Dim strRet As String
'
'    Select Case Button.key
'        Case "CompDetail"
'            frmCompany.Show
'        Case "CompCatalog"
'            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "&Group List", "-", "&Item List", "Add &New Item")
'            Select Case strRet
'                Case "&Group List"
'                    Dim f As New frmItemGroups
'                    Load f
'                    Call f.ShowForm("ItemGroups", "Item Groups", "Group Name")
'                Case "&Item List"
'                    frmItems.Show
'
'                Case "Add &New Item"
'                    With frmNewItem
'                        Load frmNewItem
'                        .Show 1
'                    End With
'            End Select
'        Case "RawMaterial"
'
'            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "Raw Material &Groups", "&Raw Material", "-", "&New Purchase Order", "&Purchase Order List", "-", "&Receive Raw Material", "&Issue Raw Material")
'            Select Case strRet
'                Case "Raw Material &Groups"
'
'                    Load frmItemGroups
'                    Call frmItemGroups.ShowForm("RMGroups", "RM Groups", "RM Name")
'                Case "&Raw Material"
'                    frmRM.Show
'                Case "&New Purchase Order"
'                    frmRMPurchOrd.Show
'                Case "&Purchase Order List"
'                    frmRMPOrdList.Show
'                Case "&Receive Raw Material"
'                    frmRMRcvItems.Show
'                Case "&Issue Raw Material"
'                    frmRMIss.Show
'            End Select
'
'        Case "SteelList"
'            frmSteelList.Show
'        Case "ForCatalog"
'            frmForeignCatalog.Show 1
'        Case "CompDept"
'
'        Case "CompResources"
'            frmHICOR.Show
'
'        Case "Banks"
'            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "&Bank List", "&Add New Bank", "-", "Add &New Cheque Book")
'
'        Case "Processes"
'            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "&Process List", "-", "&FDA Process List")
'            Select Case strRet
'                Case "&Process List"
'                    frmProcesses.Show 1
'                Case "&FDA Process List"
'                    frmFDAProcesses.Show 1
'            End Select
'
'        Case "GeneralOptions"
'            'strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "&Process List", "-", "&FDA Process List")
'            Me.PopUpMenu mnuGeneral, 512, CoolBar1.Width, MenuTop
'
'    End Select
'
'
'
'End Sub
Private Sub BusPart(ByVal Button As MSComctlLib.Button, MenuTop As Long)

    'Error Traping Of Toolbar_ButtonClick Will be Used.
    Dim strRet As String
    Select Case Button.key
        Case "FCustomers"
        
            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "&Customer List", "&Add New Customer", "-", "Customer Catalog", "-", "Customer Ledger", "-", "Customer Mailing Labels")
            Select Case strRet
                Case "&Customer List"
                    frmFCustomerList.Show
                Case "&Add New Customer"
                    frmFCustomer.Show 1
                Case "Customer Ledger"
                
                Case "Customer Catalog"
                    frmCustCatal.Show
                Case "Customer Mailing Labels"
            End Select
            
        Case "Contractors"
            Me.PopUpMenu mnuVenders, 512, CoolBar1.Width, MenuTop
            Exit Sub
            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, Button.Description & " &List", "-", "&New " & Button.Description, "-", Button.Description & " &Assigned Items", "-", Button.Description & " Ledger", "-", "Vender Advance Payment", "-", "Vender Payment Sheet")
            
            If Right(strRet, 4) = "List" Then
                Call ShowBusPartList(Button.Description)
            ElseIf Left(strRet, 4) = "&New" Then
                Call AddNewBusPart(Button.Description)
            ElseIf Right(strRet, 5) = "Items" Then
                Call ShowBusPartAssItems(Button.Description)
            ElseIf Right(strRet, 6) = "Ledger" Then
                Call ShowBusPartLedger(Button.Description)
            ElseIf strRet = "Vender Advance Payment" Then
                frmVendAdvSalary.Show
            ElseIf strRet = "Vender Payment Sheet" Then
                frmVendSS.Show
            End If
            
        Case "RMVenders"
            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, Button.Description & " &List", "-", "&New " & Button.Description, "-", Button.Description & " &Assigned Items", "-", Button.Description & " Ledger")
            If Right(strRet, 4) = "List" Then
                Call ShowBusPartList(Button.Description)
            ElseIf Left(strRet, 4) = "&New" Then
                Call AddNewBusPart(Button.Description)
            ElseIf Right(strRet, 5) = "Items" Then
                Call ShowBusPartAssItems(Button.Description)
            ElseIf Right(strRet, 6) = "Ledger" Then
                Call ShowBusPartLedger(Button.Description)
            End If
        Case Else
            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, Button.Description & " &List", "-", "&New " & Button.Description)
            If Right(strRet, 4) = "List" Then
                Call ShowBusPartList(Button.Description)
            ElseIf Left(strRet, 4) = "&New" Then
                Call AddNewBusPart(Button.Description)
            End If
        End Select
        
    
    
End Sub
Private Sub Orders(ByVal Button As MSComctlLib.Button, MenuTop As Long)
    'Error Traping Of Toolbar_ButtonClick Will be Used.
    'Toolbar Index 4.
    
    Select Case Button.key
        Case "NewOrder"
            frmFNewCustOrder.Show
        Case "NewStockOrder"
            frmFNewStockOrder.Show
        Case "Batches"
            frmBatches.Show 1
        Case "OrderList"
            frmFCustOrderList.Show
        Case "OrderStatus"
            frmFCustOrdStatus1.Show
        Case "OrderPlanning"
            frmProdPlan.Show
    End Select
    
End Sub
Private Sub Production(ByVal Button As MSComctlLib.Button, MenuTop As Long)
    'Error Traping Of Toolbar_ButtonClick Will be Used.
    
    
    Dim strRet As String
    
    Select Case Button.key
        Case "Requsition"
            frmGMPs.Show
        Case "PPlan"
            frmProdPlan.Show
        Case "Authorise"
            frmAuthorizeIssRcv.Show
        Case "Processes"
            'mnuReq.Visible = False
            'mnuProdPlan.Visible = False
            'mnuAuth.Visible = False
            Me.PopUpMenu mnuProdPhases, 512, CoolBar1.Width
    End Select
        
    
End Sub
Private Sub Payroll(ByVal Button As MSComctlLib.Button, MenuTop As Long)
    'Error Traping Of Toolbar_ButtonClick Will be Used.
    
    
    Dim strRet As String
    
    Select Case Button.key
    
        Case "Departments"
            DeptList.Show
        Case "Employees"
            EmployeeList.Show
        Case "Attendance"
            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "&Take Attendance", "Attendance &Sheet", "-", "&Over Time Record", "-", "Employee &Penalty Days")
            Select Case strRet
                Case "&Take Attendance"
                    Load Attend
                    Attend.Show
                Case "Attendance &Sheet"
                    AttSheet.Show
                Case "&Over Time Record"
                    OverTime.Show
                Case "Employee &Penalty Days"
                    EmpAbsents.Show
            End Select
        Case "AbsentSheet"
            AbsentSheet.Show
        Case "OverTime"
            OverTime.Show
        Case "ShortTerm"
            AdvanceShort.Show
        Case "LongTerm"
            LoanLedger.Show
        Case "Outstanding"
            PrintOutStanding.Show 1
        Case "AdvSalSheet"
            AdvSalary.Show
        Case "Deductions"
            
        Case "Benefits"
            
        Case "Bonuses"
            BonusSheet.Show
        Case "SalSheet"
            SalarySheet.Show
        Case "PaySlips"
            Load Payslip
            'Holidays.Move 3000, 500
            With Payslip
                .Move (Me.Width - .Width) / 2, Me.Top + 500
                .Show
            End With
        Case "EmpLedger"
            SalaryLedger.Show
        Case "SalHistory"
            SalaryHistory.Show
        Case "JoinedDis"
            PrintEmpJoinLeave.Show 1
        Case "Assets"
            
        Case "OldAge"
            OldAge.Show
        Case "SSecurity"
            SocialSecurity.Show
    End Select
        
    
End Sub
Private Sub Export(ByVal Button As MSComctlLib.Button, MenuTop As Long)
    'Error Traping Of Toolbar_ButtonClick Will be Used.
    
    Dim rpt As CRAXDDRT.Report
    Dim strRet As String
    Dim Concat As String
    Concat = "|||"
    Select Case Button.key
        Case "Company"
            strRet = ShowPopUpMenuAdv(CoolBar1.Width, MenuTop, "Company &Info" & Concat & "ExpCompInfo", "-", "&Bank List" & Concat & "ExpBankList", "-", "Company &Catalog" & Concat & "ExpCompCatalog", "-", "&Group List" & Concat & "ExpGroupList", "-", "&SRO List" & Concat & "ExpSROList", "-", "Define &Box Size" & Concat & "ExpBoxSize", "-", "&Steel Venders" & Concat & "ExpSteelVends", "-", "&FDA Processes" & Concat & "ExpFDAProcs", "-", "Ports" & Concat & "ExpPorts", "-", "Steel List" & Concat & "ExpSteelList")
            Select Case strRet
                Case "Company &Info"
                    frmCompany.Show
                Case "&Bank List"
                    frmBankList.Show 1
                Case "Company &Catalog"
                    frmItems.Show
                Case "&Group List"
                    Dim f As New frmItemGroups
                    Load f
                    Call f.ShowForm("ItemGroups", "Item Groups", "Group Name")
                Case "&SRO List"
                    frmSROList.Show
                Case "Define &Box Size"
                    frmBoxSize.Show 1
                Case "&Steel Venders"
                    frmSteelVends.Show
                Case "&FDA Processes"
                    frmFDAProcesses.Show
                Case "Ports"
                    frmPorts.Show 1
                Case "Steel List"
                    frmSteelList.Show
            End Select
            
        Case "Customers"
            strRet = ShowPopUpMenuAdv(CoolBar1.Width, MenuTop, "&Foreign Customers" & Concat & "ExpForeignCusts", "-", "Customer Catalog" & Concat & "ExpCustCatalog")
            
            Select Case strRet
                Case "&Foreign Customers"
                    frmFCustomerList.Show
                Case "&Add New Customer"
                    frmFCustomer.Show 1
                Case "Customer Catalog"
                    frmCustCatal.Show
                
            End Select
            
        Case "ProformaOrders"
            strRet = ShowPopUpMenuAdv(CoolBar1.Width, MenuTop, "Order &List" & Concat & "ExpOrderList", "-", "&New Order" & Concat & "ExpNewOrder", "-", "New &Stock Order" & Concat & "ExpNewStockOrder", "-", "&Order Status" & Concat & "ExpOrderStatus")
            
            Select Case strRet
                Case "&New Order"
                    frmFNewCustOrder.Show
                Case "New &Stock Order"
                    frmFNewStockOrder.Show
                
                Case "Order &List"
                    frmFCustOrderList.Show
                Case "&Order Status"
                    frmFCustOrderStatus.Show
                
            End Select
            
            
        Case "Agents"
            strRet = ShowPopUpMenuAdv(CoolBar1.Width, MenuTop, "&Shipping Agents" & Concat & "ExpShippingAgents", "-", "&Insurance Agents" & Concat & "ExpInsuranceAgents")
            Select Case strRet
                Case "&Shipping Agents"
                    frmShippingAgents.Show
                Case "&Insurance Agents"
                    frmInsuranceAgents.Show
            End Select
            
        Case "PInvoice"
            strRet = ShowPopUpMenuAdv(CoolBar1.Width, MenuTop, "&Proforma Invoice List" & Concat & "ExpPInvoiceList", "-", "&New Proforma Invoice" & Concat & "ExpNewPInvoice")
            Select Case strRet
                Case "&Proforma Invoice List"
                    frmFProformaList.Show
                Case "&New Proforma Invoice"
                    'frmFNewProforma.Show
                    Load frmFNewProforma
                    frmFNewProforma.Show
                    frmFNewProforma.AddNewProforma
            End Select
            
        Case "CustomDocs"
            strRet = ShowPopUpMenuAdv(CoolBar1.Width, MenuTop, "Generated Form 'E' Nos." & Concat & "ExpCustomGenFormE", "-", "Covering &Letter" & Concat & "ExpCustomCovering", "-", "&Custom Invoice" & Concat & "ExpCustomInvoice", "-", "&Packing List" & Concat & "ExpCustomPacking", "-", "&Shipping Instruction" & Concat & "ExpCustomShipping", "-", "Packing &Labels" & Concat & "ExpPackingLabels", "-", "&E Form" & Concat & "ExpEForm", "-", "Print &Valuation Form" & Concat & "ExpValuationForm", "-", "&Hide Package" & Concat & "ExpHidePackage")
            
            Select Case strRet
                Case "Generated Form 'E' Nos."
                    frmGenFormE.Show 1
                Case "Covering &Letter"
                    frmCustomCovering.Show
                Case "&Custom Invoice"
                    frmCustomInvoice.Show
                Case "&Packing List"
                    Load frmCustomPList
                    frmCustomPList.ShowForm (True)
                Case "&E Form"
                    frmEForm.Show
                Case "&Shipping Instruction"
                    frmShippingInstructions.Show
                Case "Packing &Labels"
                    frmPackingLabels.Show 1
                Case "Print &Valuation Form"
                    frmPrintValuationForm.Show , MainForm
                Case "&Hide Package"
                    frmHidePackage.Show 1
            End Select
            
        Case "ComDocs"
            strRet = ShowPopUpMenuAdv(CoolBar1.Width, MenuTop, "Covering &Letter" & Concat & "ExpComCovering", "-", "&Commercial Invoice" & Concat & "ExpComInvoice", "-", "&Packing List" & Concat & "ExpComPacking", "-", "&Inspection Certificate" & Concat & "ExpInspection", "-", "&GSP" & Concat & "ExpGSP", "-", "Enter &Freight Rate" & Concat & "ExpEnterRate", "-", "Customer Letter" & Concat & "ExpCustomerLedger")
            
            Select Case strRet
                Case "Covering &Letter"
                    frmCommercialCovering.Show
                Case "&Commercial Invoice"
                    'frmComInvList.Show
                    frmCommercialInvoice.Show
                Case "&Packing List"
                    Load frmCustomPList
                    frmCustomPList.ShowForm (False)
                Case "&Inspection Certificate"
                    frmInspection.Show
                Case "&GSP"
                    Dim FGSP As New frmPrint
                    With FGSP
                        Set .rpt = rptApp.OpenReport(RptPath & "\rptGSP.rpt")  ''New rptGSP
                        .cmbCust.ClearVals
                        '.cmbCust.AddVals con, "Distinct OrderNo", "VFItemsQtyToPlan", "OrderNo", " Where RemQty>0"
                        .cmbCust.AddItem "<All Customers>", "0"
                        .cmbCust.AddVals con, "Distinct CustCode", "ForeignCustomers", "CustCode"
                        .SelFor = "{VCustomInvoice.PInvoice}='"
                        .JobCondField = " CustCode "
                        .JobField = " CustomInvoice "
                        .JobIDField = " CustomInvoice "
                        .JobTable = " VCustomInvoice1 "
                        .TheEnd = "'"
                        .Show , MainForm
                    End With
                Case "Enter &Freight Rate"
                    frmCustomFreight.Show 1
                Case "Customer Letter"
                    Dim FCLetter As New frmPrint
                    With FCLetter
                        Set .rpt = rptApp.OpenReport(RptPath & "\rptWireTransfer.rpt")  'New rptWireTransfer
                        .cmbCust.ClearVals
                        '.cmbCust.AddVals con, "Distinct OrderNo", "VFItemsQtyToPlan", "OrderNo", " Where RemQty>0"
                        .cmbCust.AddItem "<All Customers>", "0"
                        .cmbCust.AddVals con, "Distinct CustCode", "ForeignCustomers", "CustCode"
                        .SelFor = "{VrptProformas.PInvoice}='"
                        .JobCondField = " CustCode "
                        .JobField = " CustomInvoice "
                        .JobIDField = " CustomInvoice "
                        .JobTable = " VCustomInvoice1 "
                        .TheEnd = "'"
                        .Show , MainForm
                    End With
            End Select
            
        Case "cGMP"
            strRet = ShowPopUpMenuAdv(CoolBar1.Width, MenuTop, "&Create DHR Wizard" & Concat & "ExpDHRWiz", "-", "&DHR" & Concat & "ExpDHR", "-", "&Steel Requsition" & Concat & "ExpSteelReq", "-", "&Calculation Sheet" & Concat & "ExpCalcSheet")
            
            Select Case strRet
                Case "&DHR"
                    frmDHR.Show
                Case "&Steel Requsition"
                    'frmRequisition.Show
                    frmGMPs.Show
                Case "&Create DHR Wizard"
                    frmWizDHR.Show 1
                Case "&Calculation Sheet"
                    Dim MyPrintForm As New frmPrint
                    With MyPrintForm
                        Set .rpt = rptApp.OpenReport(RptPath & "\rptSteelCalculation.rpt")  ''New rptSteelCalculation
                        .cmbCust.ClearVals
                        '.cmbCust.AddVals con, "Distinct OrderNo", "VFItemsQtyToPlan", "OrderNo", " Where RemQty>0"
                        .cmbCust.AddItem "<All Customers>", "0"
                        .cmbCust.AddVals con, "Distinct CustCode", "ForeignCustomers", "CustCode"
                        .SelFor = "{VrptSteelCalculation.PInvoice}='"
                        .JobCondField = " CustCode "
                        .JobField = " PInvoice "
                        .JobIDField = " PInvoice "
                        .JobTable = " FPInvoice "
                        .TheEnd = "'"
                        .Show , MainForm
                    End With
            End Select
            
        Case "BankDocs"
            strRet = ShowPopUpMenuAdv(CoolBar1.Width, MenuTop, "&Print Invoice" & Concat & "ExpPrintInvoice", "-", "&Letter To Bank" & Concat & "ExpLetterToBank")
            
            Select Case strRet
                Case "&Print Invoice"
                    Set rpt = rptApp.OpenReport(RptPath & "\rptCustomInvoice.rpt") '
                    Call PrintInvoiceForBank(rpt)
                Case "&Letter To Bank"
                    Set rpt = rptApp.OpenReport(RptPath & "\rptLetterToBank.rpt") '
                    Call PrintInvoice(rpt)
                
            End Select
            
        Case "PaymentStatus"
            strRet = ShowPopUpMenuAdv(CoolBar1.Width, MenuTop, "&Custom Payment Status" & Concat & "ExpCustPayStatus", "-", "Assign &FDBC No." & Concat & "ExpRecCustPay", "-", "Receive Custom &Payments" & Concat & "ExpRecCustPay", "-", "Customer Payment &Status" & Concat & "ExpCustomerPayStat", "-", "Receive Customer &Payments" & Concat & "ExpRecCustomerPay")
            Select Case strRet
                Case "&Custom Payment Status"
                    frmCustomPayments.Show
                Case "Assign &FDBC No."
                    frmFCustFDBCNo.Show 1
                Case "Receive Custom &Payments"
                    Load frmFCustRcvPymnts
                    frmFCustRcvPymnts.ShowForm "", "", True
                Case "Customer Payment &Status"
                    frmFCustPayments.Show
                Case "Receive Customer &Payments"
                    Load frmFCustRcvPymnts
                    frmFCustRcvPymnts.ShowForm "", "", False
            End Select
            
        Case "RebateDocs"
        
            strRet = ShowPopUpMenuAdv(CoolBar1.Width, MenuTop, "Rebate &Documents" & Concat & "ExpRebateDocs", "-", "&Rebate Status" & Concat & "ExpRebateStatus")
            Select Case strRet
                Case "Rebate &Documents"
                    frmRebateA.Show
                Case "&Rebate Status"
                    frmRebateStatus.Show
            End Select
            
        Case "Statistics"
        
            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "&Total Export", "-", "&Payment Ledger")
            Select Case strRet
                Case "&Total Export"
                    frmTotalExport.Show
                Case "&Payment Ledger"
            End Select
            
        Case "EEStatement"
            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "&Total Export", "-", "&Payment Ledger", "-", "&Add Invoices For EE", "-", "Form &EE")
            Select Case strRet
                Case "&Total Export"
                    frmTotalExport.Show
                Case "&Payment Ledger"
                
                Case "&Add Invoices For EE"
                    frmManInvsForEE.Show
                Case "Form &EE"
                    frmEE.Show
                
            End Select
        
        Case "UserManagement"
            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "User &Management")
            Select Case strRet
                Case "User &Management"
                    frmUsers.Show 1
            End Select
        
        Case "Quit"
            If con.State = 1 Then con.Close
            Set con = Nothing
            End
            
    End Select
    
        
    
End Sub

Sub FillPayrollItems(BarKey As String)
    
    
    Dim Rec As New ADODB.Recordset
    Rec.Open "SELECT * FROM Users WHERE UserID=" & UserID & "", con, adOpenStatic, adLockReadOnly

    With ExpBar
        Select Case Right(BarKey, Len(BarKey) - 3)
            Case "CVs"
                .Bars(BarKey).Items.add , "PrlCVsDetail", "CVs Detail"
                .Bars(BarKey).Items.add , "PrlNewCV", "Add New CV"
            Case "Employees"
                .Bars(BarKey).Items.add , "PrlDeptList", "Department List"
                .Bars(BarKey).Items.add , "PrlNewDept", "Add New Department"
                
                .Bars(BarKey).Items.add , "PrlEmpList", "Employee List"
'                If (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Banzai" Or strCompany = "Tecno" Or strCompany = "Towne" Then
'                    .Bars(BarKey).Items.add , "PrlTempEmpList", "Temp. Employee List"
'                End If
                .Bars(BarKey).Items.add , "PrlNewEmp", "Add New Employee"
                .Bars(BarKey).Items.add , "PrlEmpSalHistory", "Employee Salary History"
''                .Bars(BarKey).Items.add , "PrlEmpBonus", "Employee Bonus"
'                'Hide at the moment
'                '.Bars(BarKey).Items.Add , "PrlEmpAssets", "Employee Assets"
'                '.Bars(BarKey).Items.Add , "PrlAssignAssets", "Assign Assets"
'                '.Bars(BarKey).Items.Add , "PrlEmpSalLedger", "Employee Salary Ledger"
                .Bars(BarKey).Items.add , "PrlTransferEmp", "Transfer Employee"
                .Bars(BarKey).Items.add , "PrlGuarantorList", "Guarantor List"
                .Bars(BarKey).Items.add , "PrlPrintApplicationForm", "Application Form"
                .Bars(BarKey).Items.add , "PrlPrintEmpDetails", "Emp. Details"
'
'                If strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
'                    .Bars(BarKey).Items.add , "PrlEmpDC", "Discontinue Employees"
'                    .Bars(BarKey).Items.add , "PrlPrintEmpDC", "Print Final Settlement"
'                End If
                 
            Case "Attendance"
'                .Bars(BarKey).Items.add , "PrlTakeAttendance", "Take Attendance (Employees)"
                '.Bars(BarKey).Items.add , "PrlReadFromMachine", "Read Machine Data"
                .Bars(BarKey).Items.add , "PrlTakeAttendanceEx", "Manual Attendance"
                If strCompany <> "Dr-Frgz" Then
                    .Bars(BarKey).Items.add , "PrlOverTimeEntry", "Over Time Entry"
                End If
                .Bars(BarKey).Items.add , "PrlReCalculateAttendance", "Re-Calcualte Attendance"
'                .Bars(BarKey).Items.add , "PrlTakeAttendanceCont", "Take Attendance (Contractors)"
                .Bars(BarKey).Items.add , "PrlTakeAttendanceAuto", "Take Attendance (Auto)"
'                .Bars(BarKey).Items.add , "PrlTakeAttSGuards", "Take Attendance Of Guards"
                .Bars(BarKey).Items.add , "PrlDailyActivitySheet", "Daily Activity Sheet"
                .Bars(BarKey).Items.add , "PrlEmpLedger", "Emp. Time Ledger"
                .Bars(BarKey).Items.add , "PrlLeaves", "Take Leaves/Absents"
                .Bars(BarKey).Items.add , "PrlEditAttendance", "Edit Attendance"
'                .Bars(BarKey).Items.add , "PrlEmpOnLeaves", "Employees On Leave"
'                .Bars(BarKey).Items.add , "PrlAuthorizeOvertime", "OverTime Authorization"
'                .Bars(BarKey).Items.add , "PrlOverTimeExcess", "Excess OverTime"
'                .Bars(BarKey).Items.add , "PrlPrintOTSheet", "Print OverTime Sheet"
'                .Bars(BarKey).Items.add , "PrlScholorships", "Student Scholorships"
'                .Bars(BarKey).Items.add , "PrlEmpLunch", "Employees Monthly Lunch Detail"
'                .Bars(BarKey).Items.add , "PrlPrintEmpLedger", "Print Employee Ledger"
                'If strCompany = "Tecno" Or strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
                .Bars(BarKey).Items.add , "PrlGatePassEntries", "Gate Pass Entries"
                'End If
                If strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
                    .Bars(BarKey).Items.add , "PrlPerformanceDeductionAmt", "Performance Ded."
                    .Bars(BarKey).Items.add , "PrlRejectionCharging", "Rejection Charging"
                End If
            Case "Advances"
                'If strCompany = "Kami" Then
                    .Bars(BarKey).Items.add , "PrlLongTermLoan", "New Long Term Loan"
                    .Bars(BarKey).Items.add , "PrlShortTermLoan", "New Short Term Loan"
                'End If
                .Bars(BarKey).Items.add , "PrlClearLoan", "Clear Long Term Loan"
                .Bars(BarKey).Items.add , "PrlClearShortTermLoan", "Clear Short Term Loan"
                .Bars(BarKey).Items.add , "PrlLTEmpToEmpTransfer", "Emp. to Emp. Transfer"
                .Bars(BarKey).Items.add , "PrlAdjDeduction", "Change Deduction Amount"
                .Bars(BarKey).Items.add , "PrlShortTermLoanLedger", "Short Term Loan Ledger"
                .Bars(BarKey).Items.add , "PrlLongTermLoanLedger", "Long Term Loan Ledger"
                .Bars(BarKey).Items.add , "PrlLoanBalance", "Employees Loan Balance"
                .Bars(BarKey).Items.add , "PrlLoanBalanceDateRange", "Employees Loan Balance (Date)"
                .Bars(BarKey).Items.add , "PrlLoanBalanceTerminated", "Loan Balance (Terminated)"
                If strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
                    .Bars(BarKey).Items.add , "PrlAdvanceSalaryReport", "Advance Salary Report"
                End If
'                If strCompany = "IAA" Or strCompany = "Kami" Then
                    .Bars(BarKey).Items.add , "PrlEmpNewFine", "Post Bonus/Fine"
                    .Bars(BarKey).Items.add , "PrlEmpFineLedger", "Bonus/Fine Ledger"
'                End If
            Case "EmployeePayments"
                .Bars(BarKey).Items.add , "PrlSalarySheetExt", "Salary Sheet"   '"Salary Sheet (SA-2000 Format)"
                .Bars(BarKey).Items.add , "PrlHoldSalary", "Hold Salary"
                .Bars(BarKey).Items.add , "PrlGratuity", "Gratuity"
'            Case "Contractors"
'                .Bars(BarKey).Items.add , "PrlContractorList", "Contractor's List"
'                .Bars(BarKey).Items.add , "PrlContractorLedger", "Contractor's Ledger"
'                .Bars(BarKey).Items.add , "PrlNewContractorJob", "New Work Done"
'                .Bars(BarKey).Items.add , "PrlContNewShortTerm", "New Short Term Loan"
'                .Bars(BarKey).Items.add , "PrlContNewLongTerm", "New Long Term Loan"
'                .Bars(BarKey).Items.add , "PrlPaymentSheet", "Payment Sheet"
'                .Bars(BarKey).Items.add , "PrlContEOBI", "EOBI List"
'                .Bars(BarKey).Items.add , "PrlContSS", "Social Security List"
            Case "Reports"
'                .Bars(BarKey).Items.add , "PrlApplicationForm", "Application Form"
                .Bars(BarKey).Items.add , "PrlEmpListWithSal", "Employee List(With Salary)"
                .Bars(BarKey).Items.add , "PrlEmpListWithoutSal", "Employee List(Without Salary)"
'                .Bars(BarKey).Items.add , "PrlEmpCards", "Employee Cards"
'                .Bars(BarKey).Items.add , "PrlDailyOTSheet", "Daily Over Time Sheet"
'                .Bars(BarKey).Items.add , "PrlDailyLeavesReport", "Daily Leaves Report"
                .Bars(BarKey).Items.add , "PrlAbsentSheet", "Absent Sheet"
'                .Bars(BarKey).Items.add , "PrlJoinDisjoined", "Employee Join && Disjoined"
'                .Bars(BarKey).Items.add , "PrlLeaveEncashment", "Annual Leave Encashment"
'                .Bars(BarKey).Items.add , "PrlMonthlySalComparison", "Montly Salary Comparison"
'                .Bars(BarKey).Items.add , "PrlOutstandingLoans", "Outstanding Loans(Discontinued Employees)"
'                .Bars(BarKey).Items.add , "PrlEmpLoanBalance", "Employee Loan Balances"
                .Bars(BarKey).Items.add , "PrlSocialSecurity", "Social Security Sheet"
                .Bars(BarKey).Items.add , "PrlEOBI", "EOBI Sheet"
                .Bars(BarKey).Items.add , "PrlEmpSalaryLedger", "Salary Ledger"
                .Bars(BarKey).Items.add , "PrlAttendanceStatus", "Attendance Status"
                '.Bars(BarKey).Items.add , "PrlBonusCalculation", "Bonus Calc."
                If strCompany <> "Banzai" Then
                    .Bars(BarKey).Items.add , "PrlEmpWiseLeaves", "Employee Wise Leaves"
                    .Bars(BarKey).Items.add , "PrlMissingOutTime", "Missing OutTime"
                    .Bars(BarKey).Items.add , "PrlTerminatedEmpReport", "Terminated Employees Report"
                    .Bars(BarKey).Items.add , "PrlNewHiredReport", "New Hired Employees Report"
                    .Bars(BarKey).Items.add , "PrlLateAttendanceReport", "Late Attendance Report"
                    
'                    .Bars(BarKey).Items.add , "PrlOTPaymentSheet", "OverTime Payment Sheet"
'                    .Bars(BarKey).Items.add , "PrlOTIndividualReport", "OverTime Individual Report"
'                    .Bars(BarKey).Items.add , "PrlOTMonthlyRegister", "Monthly OverTime Register"
'                    .Bars(BarKey).Items.add , "PrlOTAnnualReportHrs", "OverTime Annual (Hrs)"
'                    .Bars(BarKey).Items.add , "PrlOTAnnualReportAmt", "OverTime Annual (Amt)"
'                    .Bars(BarKey).Items.add , "PrlSalaryRevisions", "Salary Revisions"
'                    .Bars(BarKey).Items.add , "PrlProbationEndCurrentMonth", "Probation (Current Month)"
                End If
                If strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
                    .Bars(BarKey).Items.add , "PrlEmpListGuarantorWise", "Employee List (Guarantorwise)"
                    .Bars(BarKey).Items.add , "PrlManualAttendance", "Manual Attendance"
                End If
                .Bars(BarKey).Items.add , "PrlEmployeeTargets", "Employee Targets"
'                If strCompany = "IAA" Or strCompany = "Kami" Then
'                    .Bars(BarKey).Items.add , "PrlSecurityGuardsTiming", "S-Guard Timings"
'                End If
'                .Bars(BarKey).Items.add , "PrlBankAccList", "Bank Salary List"
'                .Bars(BarKey).Items.add , "PrlTempAdvAppSheet", "Temporary Advance Approval Sheet"
'                .Bars(BarKey).Items.add , "PrlEmpYearlyDaysReport", "Days Report"
            Case "Settings"
'                .Bars(BarKey).Items.add , "PrlBarcodeSettings", "Barcode Settings"
                .Bars(BarKey).Items.add , "PrlDesignationsList", "Designations"
                .Bars(BarKey).Items.add , "PrlDepartmentSettings", "Department Settings"
                .Bars(BarKey).Items.add , "PrlPayrollPolicies", "Payroll Policies"
                .Bars(BarKey).Items.add , "PrlHolidays", "Define Holidays"
'                .Bars(BarKey).Items.add , "PrlMachineSettings", "Machine Settings"
            Case "UserManagement"
                .Bars(BarKey).Items.add , "PrlUserManager", "User Manager"
        End Select
         
        'Dim Rec  As New ADODB.Recordset, Item As String
        'Rec.Open "SELECT * FROM Users WHERE UserName='" & CurrentUserName & "'", con, adOpenStatic, adLockReadOnly
        
        On Error GoTo err
        
        For i = 1 To .Bars(BarKey).Items.count
            .Bars(BarKey).Items(i).CanClick = UserHasAccessEX(.Bars(BarKey).Items(i).key)
        Next
        
'        For i = 1 To .Bars.count - 2
'            For k = 1 To .Bars(i).Items.count
'                Item = Right(.Bars(i).Items(k).key, Len(.Bars(i).Items(k).key) - 3)
'                .Bars(i).Items(k).CanClick = rec.Fields(Item)
'            Next
'        Next
    End With
    
    
    
    
Exit Sub
err:
    If err.Number = 438 Or err.Number = 3265 Or err.Number = 3021 Then          'Field not found in User settings
        Resume Next
    Else
        MsgBox err.Description
    End If
    
End Sub

Private Sub FillStockItems(BarKey As String)

    With ExpBar
        Select Case Right(BarKey, Len(BarKey) - 3)
            Case "RawMaterial"
'                If strCompany = "Kami" Then
'                    .Bars(BarKey).Items.add , "AccVenderList", "Vender List"
'                    .Bars(BarKey).Items.add , "AccVenderAssRM", "Vender Assign RM"
'                End If
                .Bars(BarKey).Items.add , "AccVenderList", "Vender List"
                .Bars(BarKey).Items.add , "AccVenderAssRM", "Vender Assign RM"
                
                .Bars(BarKey).Items.add , "StkMaterialGroup", "Raw Material Groups"
                .Bars(BarKey).Items.add , "StkMaterialList", "Raw Material List"
                .Bars(BarKey).Items.add , "StkAddNewRM", "Add New RM"
                
                If strCompany = "Dr-Frgz" Then
                    .Bars(BarKey).Items.add , "StkAssignRM", "Assign RM"
                End If
                
                
                .Bars(BarKey).Items.add , "StkPO", "Purchase Order"
                .Bars(BarKey).Items.add , "StkPOList", "PO List"
                .Bars(BarKey).Items.add , "StkOpenRMPOList", "Open RM PO List"
                
                .Bars(BarKey).Items.add , "StkPORcv", "Receive PO Items"
                .Bars(BarKey).Items.add , "StkRcvList", "Receive List"
                
                If strCompany <> "Banzai" Then
                    .Bars(BarKey).Items.add , "StkVenderBilling", "Vendor Billing"
                    .Bars(BarKey).Items.add , "StkVenderBillingList", "Vendor Billing List"
                End If
                
                If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                    .Bars(BarKey).Items.add , "StkMaterialPlacement", "Material Placement"
                    .Bars(BarKey).Items.add , "StkMaterialPlacementList", "Material Placement List"
                    .Bars(BarKey).Items.add , "StkChangeMaterialPlacement", "Change Material Location"
                End If
                .Bars(BarKey).Items.add , "StkChangeBatches", "Change Batches"
                .Bars(BarKey).Items.add , "StkChangeBatchesLot", "Change Batches Lot"
                .Bars(BarKey).Items.add , "StkRMIssuance", "RM Issuance"
                .Bars(BarKey).Items.add , "StkRMIssuanceList", "RM Issuance List"
                .Bars(BarKey).Items.add , "StkRMLedger", "RM Ledger"
                .Bars(BarKey).Items.add , "StkPrintRMStockList", "RM Stock List"
                
            Case "RawMaterialSampling"
                
                .Bars(BarKey).Items.add , "StkMaterialGroupSampling", "Raw Material Groups"
                .Bars(BarKey).Items.add , "StkMaterialListSampling", "Raw Material List"
                
                .Bars(BarKey).Items.add , "StkPOSampling", "Purchase Order"
                .Bars(BarKey).Items.add , "StkPOListSampling", "PO List"
                
                .Bars(BarKey).Items.add , "StkPORcvSampling", "Receive PO Items"
                .Bars(BarKey).Items.add , "StkRcvListSampling", "Receive List"
                
                
                .Bars(BarKey).Items.add , "StkVenderBillingSampling", "Vendor Billing"
                .Bars(BarKey).Items.add , "StkVenderBillingListSampling", "Vendor Billing List"
                
                
                .Bars(BarKey).Items.add , "StkMaterialPlacementSampling", "Material Placement"
                .Bars(BarKey).Items.add , "StkChangeMaterialPlacementSampling", "Change Material Location"
                
'                .Bars(BarKey).Items.add , "StkRMIssuanceSampling", "RM Issuance"
'                .Bars(BarKey).Items.add , "StkRMIssuanceListSampling", "RM Issuance List"
'                .Bars(BarKey).Items.add , "StkRMLedgerSampling", "RM Ledger"
                
            Case "SemiFinishStock"
                .Bars(BarKey).Items.add , "PrdOpeningStocks", "S/F Receive Items"
                If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                    .Bars(BarKey).Items.add , "StkChangeLocationsSF", "Change Locations"
                End If
                .Bars(BarKey).Items.add , "PrdSFTransaction", "S/F Transaction"
                .Bars(BarKey).Items.add , "PrdSFIssuance", "S/F Issue"
            Case "Stock"
            
                .Bars(BarKey).Items.add , "StkIssItems", "Issue Items"
                
                '.Bars(BarKey).Items.add , "StkIssItemsList", "Issue Items List"
                .Bars(BarKey).Items.add , "StkRcvItems", "Receive Items"
                .Bars(BarKey).Items.add , "StkStockTransaction", "Stock Transaction"
                
                .Bars(BarKey).Items.add , "StkItemLedger", "Item Ledger"
                
'                If strCompany = "QEL" Or (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Then
'                    .Bars(BarKey).Items.add , "PrdOpeningStocks", "S/Finish Opening Stocks"
'                End If
                
                If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                    .Bars(BarKey).Items.add , "StkChangeLocations", "Change Locations"
                End If
                If strCompany = "Dr-Frgz" Then
                    .Bars(BarKey).Items.add , "StkDeadStock", "Dead Stock"
                End If
                .Bars(BarKey).Items.add , "StkBelowMinLevel", "Below Min. Level"
                .Bars(BarKey).Items.add , "StkAboveMinLevel", "Above Min. Level"
            Case "GateEntries"
                .Bars(BarKey).Items.add , "StkGateIn", "Gate-In Entries"
                .Bars(BarKey).Items.add , "StkGateOut", "Gate-Out Entries"
            Case "RejectionStore"
                .Bars(BarKey).Items.add , "StkReceiveRejection", "Receive Rejection"
                .Bars(BarKey).Items.add , "StkRejectionTransaction", "Rejection Transaction"
                .Bars(BarKey).Items.add , "StkRejectionIssuance", "Rejection Issuance"
                .Bars(BarKey).Items.add , "StkRejectionAlerts", "Alerts"
        End Select
        
        On Error GoTo err
        For i = 1 To .Bars(BarKey).Items.count
            .Bars(BarKey).Items(i).CanClick = UserHasAccessEX(.Bars(BarKey).Items(i).key)
        Next
    End With
    Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub FillProductionItems(BarKey As String)

    With ExpBar
        Select Case Right(BarKey, Len(BarKey) - 3)
            Case "Orders"
                .Bars(BarKey).Items.add , "PrdNewOrder", "New Order"
                .Bars(BarKey).Items.add , "PrdProductionOrder", "Customer Orders"
            Case "PurchaseANDProcessing"
                If Not (strCompany = "Tecno" Or (strCompany = "Dr-Frgz" And strSubCompany = "Electro")) Then
                    .Bars(BarKey).Items.add , "PrdIssuance", "Issue/Purchasing"
                    .Bars(BarKey).Items.add , "PrdSFIssuance", "S/F Issue"
                    .Bars(BarKey).Items.add , "PrdProcessIssuance", "Process Issuance"
                End If
                .Bars(BarKey).Items.add , "PrdIssuanceList", "Issuance List"
                .Bars(BarKey).Items.add , "PrdOpenMasterPOs", "Open Master POs"
                .Bars(BarKey).Items.add , "PrdOpenLots", "Open Lots"
                If Not (strCompany = "Tecno" Or (strCompany = "Dr-Frgz" And strSubCompany = "Electro")) Then
                    .Bars(BarKey).Items.add , "PrdReceivingAgainstPO", "Receive against PO"
                    .Bars(BarKey).Items.add , "PrdReceiving", "Receive Lots"
                    .Bars(BarKey).Items.add , "PrdReceivingList", "Receiving List"
                End If
                
                If strCompany = "QEL" Or (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Then
                    .Bars(BarKey).Items.add , "PrdReWorkIssuance", "Re-Work Issuance"
                    .Bars(BarKey).Items.add , "PrdTransferToReadyFinish", "Transfer to Ready Finish"
                End If
                
            Case "Production"
            
                If strCompany = "Weldon Industries" Or strCompany = "All Seas" Or strCompany = "Weldon Instruments" Or strCompany = "Kami" Then
                    .Bars(BarKey).Items.add , "PrdNewOrder", "New Order"
                End If
                If strCompany <> "Tecno" Then
                    .Bars(BarKey).Items.add , "PrdProductionOrder", "Customer Orders"
                    .Bars(BarKey).Items.add , "PrdProductionOrderCust", "Cust Order List"
                    .Bars(BarKey).Items.add , "PrdProductionOrderStocksAndParts", "Stock && Parts Order List"
                    .Bars(BarKey).Items.add , "PrdViewOrderItemPDFs", "Marking Plan"
                End If
                If strCompany = "Tecno" Or (strCompany = "Dr-Frgz" And strSubCompany = "Electro") Then
                    .Bars(BarKey).Items.add , "PrdIPOList", "IPO List"
                End If
                
                If (strCompany = "Dr-Frgz" And strSubCompany = "Electro") Then
                    .Bars(BarKey).Items.add , "PrdReWorkIssuance", "Re-Work Issuance"
                    .Bars(BarKey).Items.add , "PrdIssuanceList", "Issuance List"
                    .Bars(BarKey).Items.add , "PrdReceivingList", "Receiving List"
                End If
                
                If strCompany <> "All Seas" Then
                    
                    If Not (strCompany = "Tecno" Or (strCompany = "Dr-Frgz" And strSubCompany = "Electro")) Then
                        .Bars(BarKey).Items.add , "PrdAuthorization", "Authorization"
                        .Bars(BarKey).Items.add , "PrdAuthorizationRejection", "Auth. Rejection"
                    End If
                    If strCompany <> "Tecno" Then
                        .Bars(BarKey).Items.add , "PrdArticlewiseWIP", "Articlewise (WIP)"
                        .Bars(BarKey).Items.add , "PrdCreateDispatchList", "Create Dispatch List"
                        .Bars(BarKey).Items.add , "PrdDispatchList", "Dispatch List"
                    End If
                    .Bars(BarKey).Items.add , "PrdMakerBilling", "Maker Billing"
                    .Bars(BarKey).Items.add , "PrdMakerBillingList", "Maker Billing List"
                    
                    
'                    If strCompany = "Sunlike" Or strCompany = "QEL" Or (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Then
'                        .Bars(BarKey).Items.add , "PrdOpeningStocks", "S/Finish Opening Stocks"
'                    End If
'                    If strCompany = "QEL" Or (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Then
'                        .Bars(BarKey).Items.add , "PrdReWorkIssuance", "Re-Work Issuance"
'                    End If
                    '\/\/\/\/\/
'                    If strCompany <> "Tecno" Then
'                        .Bars(BarKey).Items.add , "PrdRMPO", "Raw Material PO"
'                    End If
                    If strCompany = "Tecno" Or (strCompany = "Dr-Frgz" And strSubCompany = "Electro") Then
                        .Bars(BarKey).Items.add , "PrdMakerNewShortTermLoan", "New Short Term Loan"
                        .Bars(BarKey).Items.add , "PrdMakerNewLongTermLoan", "New Long Term Loan"
                    End If
                    .Bars(BarKey).Items.add , "PrdProcessWiseStockReport", "Stock Report"
                    If strCompany = "Dr-Frgz" Or strCompany = "IAA" Or strCompany = "Kami" Then
                        If strCompany = "Dr-Frgz" Then
                            .Bars(BarKey).Items.add , "PrdPieciesReport", "Group Load"
                            .Bars(BarKey).Items.add , "PrdDailyPacking", "Daily Packing"
                            .Bars(BarKey).Items.add , "PrdPlanningReport", "Inventory Planning Report"
                        End If
'                        If strCompany = "IAA" Or strCompany = "Kami" Then
'                            .Bars(BarKey).Items.add , "PrdProductoinTimeLineReport", "Prod. Timeline Report"
'                        End If
                    End If
                    
                End If
            Case "Statistics"
                .Bars(BarKey).Items.add , "PrlProductivityEntries", "Emp. Productivity"
                .Bars(BarKey).Items.add , "PrlProductivityMonthly", "Emp. Prod. (Monthly)"
                .Bars(BarKey).Items.add , "PrlProductivityEntriesDaily", "Dept. Productivity"
                '.Bars(BarKey).Items.add , "PrlProductivityEntriesOT", "Emp. Productivity (OT)"
                '.Bars(BarKey).Items.add , "PrlProductivityEntriesDailyOT", "Dept. Productivity (OT)"
                .Bars(BarKey).Items.add , "PrlEmpProdLineChart", "Productivity Chart"
'                If strCompany = "Dr-Frgz" Or strCompany = "IAA" Or strCompany = "Kami" Then
'                    .Bars(BarKey).Items.add , "PrlProcessWiseReport", "Processwise Report"
'                End If
'                If strCompany = "IAA" Or strCompany = "Kami" Then
'                    .Bars(BarKey).Items.add , "PrlMakerPerformanceReport", "Maker Performance"
'                End If
'                .Bars(BarKey).Items.add , "PrdDemandAndSupplyGraph", "D&&S Graph"
'                .Bars(BarKey).Items.add , "PrdDeliveryPerformanceGraph", "Delivery Performance Graph"
'                .Bars(BarKey).Items.add , "PrdDeliveryPerformanceDetailReportWOValue", "Delivery Performance Report"
                .Bars(BarKey).Items.add , "PrdProductionItemList", "Production Item List"
                'bar.Items.add , "PrdDemandAndSupplyGraph", "19. D&&S Graph"
                'bar.Items.add , "PrdDeliveryPerformanceGraph", "20. Delivery Performance Graph"
                .Bars(BarKey).Items.add , "PrdHubwiseRunningLots", "Running Lots Hubwise"
                .Bars(BarKey).Items.add , "PrdRunningLotsQty", "Running Lots"
                .Bars(BarKey).Items.add , "PrdRunningLotsValue", "Running Lots (Value)"
                .Bars(BarKey).Items.add , "PrdRunningLotsStock", "Running Lots (Stock Order)"
                .Bars(BarKey).Items.add , "PrdLotsPacked", "Lots Packed"
                
            Case "QA"
                If strCompany = "Tecno" Or (strCompany = "Dr-Frgz" And strSubCompany = "Electro") Then
                    .Bars(BarKey).Items.add , "PrdReceiving", "Receiving"
                End If
            Case "Store"
                If strCompany = "Tecno" Or (strCompany = "Dr-Frgz" And strSubCompany = "Electro") Then
                    .Bars(BarKey).Items.add , "PrdStoreIssuanceList", "Issuance List"
                    .Bars(BarKey).Items.add , "PrdReceivingList", "Receiving List"
                End If
            Case "ManagementReports"
                .Bars(BarKey).Items.add , "PrdMonthlyRejectionReport", "Monthly Rejection Report"
            Case "GeneralLists"
'                .Bars(BarKey).Items.add , "PrdInspectionProcesses", "Inspection Processes"
                .Bars(BarKey).Items.add , "PrdProcesses", "Processes"
                .Bars(BarKey).Items.add , "PrdProcessGroups", "Process Groups"
                .Bars(BarKey).Items.add , "PrdRepairTypes", "Repair Types"
                .Bars(BarKey).Items.add , "PrdWastageTypes", "Wastage Types"
            
                
        End Select
        On Error GoTo err
        For i = 1 To .Bars(BarKey).Items.count
            .Bars(BarKey).Items(i).CanClick = UserHasAccessEX(.Bars(BarKey).Items(i).key)
        Next
    End With
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub FillQMSItems(BarKey As String)
    'FillImportItems
    With ExpBar
        Select Case Right(BarKey, Len(BarKey) - 3)
            Case "QA"
                If strCompany = "Tecno" Or strCompany = "Dr-Frgz" Then
                    .Bars(BarKey).Items.add , "QAReceiving", "Receiving"
                Else
                    .Bars(BarKey).Items.add , "QACustomerComplainsts", "Complaints/Feedback/Audit"
                End If
            Case "QMS"
                .Bars(BarKey).Items.add , "QMSCustomerComplaints", "QMS Command Center"
                .Bars(BarKey).Items.add , "AccCalibrations", "Calibrations"
                .Bars(BarKey).Items.add , "QMSDocumentTypes", "Document Types"
                .Bars(BarKey).Items.add , "QMSDocuments", "Document Control"
                
                .Bars(BarKey).Items.add , "QMSMeetings", "Meetings"
                
                .Bars(BarKey).Items.add , "QMSTrackBackItems", "Track back Items"
                
                .Bars(BarKey).Items.add , "QMSInitialEvaluation", "Supplier Evaluation"
                .Bars(BarKey).Items.add , "QMSSupplierEvaluationList", "Supplier Evaluation List"
                .Bars(BarKey).Items.add , "QMSSupplierReEvaluationList", "Supplier Re-Evaluation List"
                .Bars(BarKey).Items.add , "QMSReasons", "Reasons"
                
                .Bars(BarKey).Items.add , "QMSReasonwiseGraph", "Reasonwise Graph"
                .Bars(BarKey).Items.add , "QMSGroupwiseGraph", "Groupwise Graph"
                .Bars(BarKey).Items.add , "QMSClaimsStatistics", "Claims Statistics"
            Case "Documents"
                .Bars(BarKey).Items.add , "QMSStandards", "Standards"
                .Bars(BarKey).Items.add , "QMSManuals", "Manuals"
                .Bars(BarKey).Items.add , "QMSProcedures", "Procedures"
                .Bars(BarKey).Items.add , "QMSWorkInstructions", "Work Instructions"
                .Bars(BarKey).Items.add , "QMSTechnicalFiles", "Technical Files"
                .Bars(BarKey).Items.add , "QMSLeanManagement", "Lean Management"
        End Select
        On Error GoTo err
        For i = 1 To .Bars(BarKey).Items.count
            .Bars(BarKey).Items(i).CanClick = UserHasAccessEX(.Bars(BarKey).Items(i).key)
        Next
    End With
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub fillSamplingItems(BarKey As String)
    '
    With ExpBar
        Select Case Right(BarKey, Len(BarKey) - 3)
            Case "Sampling"
                .Bars(BarKey).Items.add , "SAMSampleCatalog", "Sample Catalogue"
'                .Bars(BarKey).Items.add , "SAMNewSampleOrder", "New Sample Order"
'                .Bars(BarKey).Items.add , "SAMSampleOrderList", "Sample Order List"
'                .Bars(BarKey).Items.add , "SAMPurchaseOrder", "Samples PO"
'                .Bars(BarKey).Items.add , "SAMPurchaseOrderList", "Samples PO List"
'                .Bars(BarKey).Items.add , "SAMReceiveItems", "Receive Samples"
'                .Bars(BarKey).Items.add , "SAMReceiveList", "Samples Receive List"
                                
                .Bars(BarKey).Items.add , "SAMOpeningStock", "Opening Stock"
                .Bars(BarKey).Items.add , "SAMIssuance", "Sample Issuance"
                .Bars(BarKey).Items.add , "SAMIssuanceList", "Issuance List"
                .Bars(BarKey).Items.add , "SAMReceiving", "Receiving"
                .Bars(BarKey).Items.add , "SAMSamplingStock", "Sampling Stock"
                
'                .Bars(BarKey).Items.add , "SAMSampleInvoiceList", "Sample Invoice List"
'
'                .Bars(BarKey).Items.add , "SAMSampleBilling", "Sample Billing"
'                .Bars(BarKey).Items.add , "SAMSampleBillingList", "Sample Billing List"
                
        End Select
        On Error GoTo err
        For i = 1 To .Bars(BarKey).Items.count
            .Bars(BarKey).Items(i).CanClick = UserHasAccessEX(.Bars(BarKey).Items(i).key)
        Next
    End With
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub FillDashBoardItems(BarKey As String)
    '
    With ExpBar
        Select Case Right(BarKey, Len(BarKey) - 3)
            Case "DashBoard"
                '.Bars(BarKey).Items.add , "DasIssueReceive", "Issue / Receive"
                .Bars(BarKey).Items.add , "DasCommandCenter", "Command Center"
                '.Bars(BarKey).Items.add , "DasCommandCenterProductionStock", "Production / Stock"
                '.Bars(BarKey).Items.add , "DasJobOrderCosting", "Job Order Costing"
                .Bars(BarKey).Items.add , "DasProductionPlanning", "Production Planning"
'                .Bars(BarKey).Items.add , "DasEmployeeDashBoard", "Employee Dashboard"
'                .Bars(BarKey).Items.add , "DasMakerDashBoard", "Maker Dashboard"
''                .Bars(BarKey).Items.add , "DasMakerLoadDashBoard", "Maker Load Dashboard"
''                If (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Towne" Then
''                    .Bars(BarKey).Items.add , "DasAlertsDashBoard", "Alerts Dashboard"
''                End If
''
''                If strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
'                    .Bars(BarKey).Items.add , "DasTrafficLights", "Traffic Dashboard"
''
''                End If
''                If strCompany = "IAA" Or strCompany = "Kami" Then
''                    .Bars(BarKey).Items.add , "PrlMonthlyReport", "Monthly Report"
''                    .Bars(BarKey).Items.add , "DasProductionTimeLineDashBoard", "Production Timeline"
'                    .Bars(BarKey).Items.add , "DasSalesDashBoard", "Sales Dashboard"
''                End If
''                'If strCompany = "IAA" OR strCompany = "Kami" Then
''                    .Bars(BarKey).Items.add , "DasProcessDashboard", "Process Dashboard"
'                'End If
            Case "Statistics"
                .Bars(BarKey).Items.add , "PrlMonthlyReport", "Monthly Report"
                .Bars(BarKey).Items.add , "DasProductionStatusLog", "Production Status Log"
                .Bars(BarKey).Items.add , "DasMakerPerfromanceReport", "Maker Performance"
        End Select
        On Error GoTo err
        For i = 1 To .Bars(BarKey).Items.count
            .Bars(BarKey).Items(i).CanClick = UserHasAccessEX(.Bars(BarKey).Items(i).key)
        Next
    End With
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub FillHelpItems(BarKey As String)
    '
    With ExpBar
        Select Case Right(BarKey, Len(BarKey) - 3)
            Case "Help"
                .Bars(BarKey).Items.add , "HlpHelp", "Help"
                .Bars(BarKey).Items.add , "HlpTips", "Tips"
                .Bars(BarKey).Items.add , "HlpTipViewer", "Tip Viewer"
        End Select
        On Error GoTo err
        For i = 1 To .Bars(BarKey).Items.count
            .Bars(BarKey).Items(i).CanClick = UserHasAccessEX(.Bars(BarKey).Items(i).key)
        Next
    End With
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub ShowProductionForm(ITemKey As String)
    'ShowQAForm
    Dim f As Form
    Select Case Right(ITemKey, Len(ITemKey) - 3)
        Case "NewOrder"
            frmFNewCustOrder_Frgz.Show
            'frmFNewCustOrder.Show
        Case "ProductionOrder"
            Set f = New frmFCustOrderList
            Load f
            f.ShowMe True
        Case "ProductionOrderCust"
            Set f = New frmFCustOrderList
            Load f
            Call f.ShowMe(True, False, 1)
        Case "ProductionOrderStocksAndParts"
            Set f = New frmFCustOrderList
            Load f
            Call f.ShowMe(True, False, 2)
        Case "ViewOrderItemPDFs"
            frmViewOrderItemPDFs.Show
        Case "IPOList"
            frmIPOList.Show
        Case "Issuance"
            'frmVendIssItemsByMaker.Show
            If strCompany = "Weldon Instruments" Or (strCompany = "QEL" Or (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Banzai") Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                Set f = New frmVendIssItemsByMaker
            Else
                Set f = New frmVendIssItemsByMakerOriginal
            End If
            Load f
            f.ShowMe (False)
        Case "IssuanceList"
            frmMakerPOList.Show
        Case "OpenMasterPOs"
            Call PrintOpenMakerOrders
        Case "OpenLots"
            Call PrintOpenLots
'            If strCompany = "Tecno" Or (strCompany = "Dr-Frgz" And strSubCompany = "Electro") Then
'                frmMakerPOListTecno.Show
'            Else
'                frmMakerPOList.Show
'            End If
        Case "StoreIssuanceList"
            frmMakerPOListTecno.Show
        Case "ReceivingAgainstPO"
            frmVendRcvItems_PO.Show
        Case "Receiving"
            frmVendRcvItems_Lots.Show
        Case "ReceivingManual"
            frmVendRcvItemsManual.Show
        Case "ReceivingList"
'            If strCompany = "Tecno" Or (strCompany = "Dr-Frgz" And strSubCompany = "Electro") Then
'                frmMakerRcvListTecno.Show
'            Else
'                frmMakerRcvList.Show
'            End If
            frmMakerRcvList.Show
        Case "TransferToReadyFinish"
            frmTransferToReadyFinishStock.Show
        Case "CreateDispatchList"
'            If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Then
'                frmCreateDispatchList_Adv.Show
'            Else
'                frmCreateDispatchList.Show
'            End If
            frmCreateDispatchList_Adv.Show
'            frmCreateDispatchList.Show
        Case "DispatchList"
            frmDispatchList.Show
        Case "SFIssuance"
            Set f = New frmMakerIssuanceFromSF
            Load f
            f.ShowMe False
        Case "ProcessIssuance"
            'frmVendIssItemsByItem.Show
'            If strCompany = "Weldon Instruments" Or (strCompany = "QEL" Or (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne") Then
'                Set f = New frmVendIssItemsByMaker
'            Else
'                Set f = New frmVendIssItemsByMakerOriginal
'            End If
            Set f = New frmProcessIssuance
            Load f
            f.ShowMe True
        Case "IPOProcessInspection"
            frmIPOLotsInspection.Show 1
        Case "Authorization"
            If strCompany = "QEL" Or (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                Set f = New frmAuthorizeIssRcvAdv
                Load f
                f.ShowMe
            Else
                frmAuthorizeIssRcv.Show
            End If
        Case "AuthorizationRejection"
            frmMakerBilling_Rejections.Show
        Case "MakerTempLoan"
            frmMakerTempLoan.Show
        Case "MakerBilling"
            If strCompany = "Tecno" Or (strCompany = "Dr-Frgz" And strSubCompany = "Electro") Then
                frmMakerBillingTecno.Show
            Else
                frmMakerBillingRcvingWise.Show
            End If
        Case "MakerAdvanceClearance"
            Set f = New frmMakerLoanClearance
            Load f
            Call f.ShowMe(0)
        Case "MakerLoanTypeTransfer"
            Set f = New frmMakerLoanType_Transfer
            Load f
            Call f.ShowMe(0)
        Case "MakerSTLedger"
            frmMakerShortLoanLedger.Show
        Case "MakerLTLedger"
            frmMakerLongLoanLedger.Show
        Case "MakerLTBalances"
            Call PrintVendorBalances
        Case "MakerBillingList"
            frmMakerBillingList.Show
        Case "ArticlewiseWIP"
            frmArticlewiseWIP.Show
        Case "OpeningStocks"
            If iProductionType = 0 Then
                frmProductionOpeningStocks_Frgz.Show
            Else
                frmProductionOpeningStocks.Show
            End If
        Case "SFTransaction"
            frmSFStockTransaction.Show
        Case "SFLedger"
            'Call PrintSFLedger
        Case "ReWorkIssuance"
            If strCompany = "Dr-Frgz" And strSubCompany = "Electro" Then
                frmReWorkIssuance_IPOLots.Show
            Else
                frmReWorkIssuance.Show
            End If
        Case "RMPO"
            frmVendNewPORM.Show
        Case "MakerNewShortTermLoan"
            Set f = New frmNewShortTermMaker
            Load f
            f.ShowMe
        Case "MakerNewLongTermLoan"
            Set f = New frmNewLongTermMaker
            Load f
            f.ShowMe
        Case "ProcessWiseStockReport"
            frmPrintProcessWiseStockReport.Show 1
        Case "PieciesReport"
            Call PrintPieciesReport
        Case "DailyPacking"
            Call PrintDailyPackingReport
        Case "PlanningReport"
            Call PrintPlanningReport
        Case "ProductoinTimeLineReport"
            Call PrintMonthlyProductionTimeLineReport
        Case "Processes"
            'frmProcesses.Show
            Set f = New frmProcesses
            Load f
            f.ShowMe
        Case "ProcessGroups"
            frmProcessGroups.Show
        Case "RepairTypes"
            frmRepairTypes.Show 1
        Case "WastageTypes"
            frmWastageTypes.Show 1
        Case "InspectionProcesses"
            'frmInspectionProcesses.Show 1
            Set f = New frmProcesses
            Load f
            f.ShowMe True
        Case "RMGroups"
            Set f = New frmItemGroups
            Load f
            Call f.ShowForm("RMGroups", "RM Groups", "RM Name")
        Case "RM"
            frmRM.Show
        Case "MakerAssItems"
            If strCompany = "Tecno" Or (strCompany = "Dr-Frgz" And strSubCompany = "Electro") Or strCompany = "Instrumed" Then
                frmVendAssItemsTecno.Show   'Without Process No.
            Else
                frmVendAssItems.Show
            End If
        Case "ProductionProcess"
        'Tecno Management Reports.
        Case "MonthlyRejectionReport"
            Call PrintMonthlyRejectionReport
        'Statistics
        Case "YearlyDemand"
            Call PrintYearlyDemandReport
        Case "YearlyDemandAmt"
            Call PrintYearlyDemandReport(False, True)
        Case "YearlyDemandSupply"
            Call PrintYearlyDemandSupply
        Case "YearlyDemandSupplyGraph"
            Call PrintYearlyDemandSupply(True)
        Case "YearlyDemandSupplyGraphAt"
            Call PrintYearlyDemandSupply(True, True)
        Case "YearlyGroupwiseDemand"
            Call PrintYearlyDemandReport(True)
        Case "MainGroupDemandSupply"
            Call ItemWiseDemandSupply(2)
        Case "SubGroupDemandSupply"
            Call ItemWiseDemandSupply(1)
        Case "ItemDemandSupply"
            Call ItemWiseDemandSupply(0)
        Case "CustomerWiseDemandSupply"
            Call ItemWiseDemandSupply(True, 0)
        Case "TotalExportValueShare"
            Call ItemWiseDemandSupply(True, 3)
        Case "PeriodwiseDemand"
            Call PrintPeriodWiseDemand
        Case "DeliveryPerformanceReport"
            Call ItemWiseDemandSupply(False, 4)
        Case "DeliveryPerformanceDetailReport"
            Call ItemWiseDemandSupply(False, 5)
        Case "DeliveryPerformanceDetailReportWOValue"
            Call ItemWiseDemandSupply(False, 5, True)
        Case "ProductionItemList"
            frmProductionOrderItemList.Show
        Case "DemandAndSupplyGraph"
            Call PrepareDemandAndSupplyGraph
        Case "DeliveryPerformanceGraph"
            Call PrintDeliveryPerformanceGraph
        Case "HubwiseRunningLots"
            Call PrintRunningLots_Hubwise
        Case "RunningLotsQty"
            Call PrintRunningLots(False)
        Case "RunningLotsValue"
            Call PrintRunningLots(True)
        Case "RunningLotsStock"
            Call PrintRunningLots(False, 1)
        Case "LotsPacked"
            Call PrintRunningLots(False, 2)
        Case "CustomerPurchaseHistory"
            Call CustomerPurchaseHistory_Orderwise
        Case "CustomerStatusValueWise"
            Call PrintCustomerStatus_Valuewise
        Case "ItemwiseYearlyDemand"
            Call PrintYearlyDemandReport_ItemWise
        Case "ItemwiseYearlyDemandX2"
            Call PrintYearlyDemandReport_ItemWise(False, False, False, True)
        Case "ItemwiseYearlyDemand_Yearwise"
            Call PrintYearlyDemandReport_ItemWise(, , True)
        Case "YearlyCustomerWiseDemands"
            Call PrintCustomerWiseDemands_Yearwise
        Case "ProductLineDemand"
            Call PrintProductLineDemand
        Case "ProductGroupwiseDemand"
            Call PrintProductLineDemand(True)
        Case "ProductLineSupply"
            Call PrintProductLineSupply
        Case "ProductGroupwiseSupply"
            Call PrintProductLineSupply(True)
        Case "TotalDemandValueShare"
            Call ItemWiseDemandSupply(True, 6)
        Case "AdvancePlan"
            Call PrintAdvancePlanceReport
    End Select
    
End Sub

Private Sub ShowQAForm(ITemKey As String)
    'ShowDashBoardForm
    Dim f As Form
    Select Case Right(ITemKey, Len(ITemKey) - 2)
        
        Case "Receiving"
            If strCompany = "Tecno" Or strCompany = "Dr-Frgz" Then
                frmRcvIPOLots.Show 1
            Else
                frmVendRcvItems.Show
            End If
        Case "CustomerComplaints"
            
    End Select
    
End Sub

Private Sub ShowDashBoardForm(ITemKey As String)
    '
    Dim f As Form
    Select Case Right(ITemKey, Len(ITemKey) - 3)
        
        Case "IssueReceive"
            Load frmDashBoard
            frmDashBoard.ShowMe
        Case "JobOrderCosting"
            Load frmJobOrderCosting
            frmJobOrderCosting.ShowMe
        Case "ProductionPlanning"
            Load frmProductionPlanningDashBoard
            frmProductionPlanningDashBoard.ShowMe
        Case "EmployeeDashBoard"
            Load frmEmpDashBoard
            frmEmpDashBoard.ShowMe
        Case "MakerDashBoard"
            Load frmMakerDashBoard
            frmMakerDashBoard.ShowMe
        Case "MakerLoadDashBoard"
            Load frmMakerLoadDashBoard
            frmMakerLoadDashBoard.ShowMe
        Case "CommandCenter"
            Load frmDashBoardCC
            frmDashBoardCC.ShowMe
        Case "CommandCenterProductionStock"
            Set f = New frmDashBoardCC
            Load f
            f.ShowMe (True)
        Case "AlertsDashBoard"
            Load frmAlertsDashBoard
            frmAlertsDashBoard.ShowMe
        Case "TrafficLights"
            frmTrafficDashBoard.Show
        Case "ProcessDashboard"
            Load frmProcessDashBoard
            frmProcessDashBoard.ShowMe
        Case "ProductionTimeLineDashBoard"
            frmProductionTimeLineDashBoard.Show
        Case "ProductionStatusLog"
            frmFCustOrderList.Show
        Case "MakerPerfromanceReport"
            frmMakerRcvList.Show
        Case "SalesDashBoard"
            Load frmDashBoardExport
            frmDashBoardExport.ShowMe
    End Select
    
End Sub

Private Sub ShowHelpForm(ITemKey As String)
    '
    Dim f As Form
    Select Case Right(ITemKey, Len(ITemKey) - 3)
        
        Case "Help"
            frmHelpList.Show
        Case "Tips"
            frmTipList.Show
        Case "TipViewer"
            frmTipViewer.Show 1
    End Select
    
End Sub


Private Sub ShowQMSForm(ITemKey As String)
    'ShowSamplingForm
    Dim f As Form
    Select Case Right(ITemKey, Len(ITemKey) - 3)
        
        Case "CustomerComplaints"
            frmCustomerComplaints.Show
        Case "Documents"
            frmQMSDocuments.Show
        Case "DocumentTypes"
            Call ShowQMSDocumentTypes
        Case "Meetings"
            frmQMSMeetings.Show
        Case "TrackBackItems"
            frmTrackBackItems.Show
        Case "InitialEvaluation"
            frmSupplierEvaluationAE.Show 1
        Case "SupplierEvaluationList"
            frmSupplierEvaluationList.Show
        Case "SupplierReEvaluationList"
            frmSupplierReEvaluationList.Show
        Case "Reasons"
            Set f = New frmMiscS_New
            Load f
            Call f.ShowForm("QMS_Reasons_List", "Reason_QMS", "Reasons", "EntryID", "QMS Reasons")
        Case "ReasonwiseGraph"
            Call PrintQMSGraph(0)
        Case "GroupwiseGraph"
            Call PrintQMSGraph(1)
        Case "ClaimsStatistics"
            Call PrintQMSGraph(2)
        Case "Standards"
            Set f = New frmQMSDocuments_Ex
            Load f
            f.ShowMe "Standards"
        Case "Manuals"
            Set f = New frmQMSDocuments_Ex
            Load f
            f.ShowMe "Manuals"
        Case "Procedures"
            Set f = New frmQMSDocuments_Ex
            Load f
            f.ShowMe "Procedures"
        Case "WorkInstructions"
            Set f = New frmQMSDocuments_Ex
            Load f
            f.ShowMe "Work Instructions"
        Case "TechnicalFiles"
            Set f = New frmQMSDocuments_Ex
            Load f
            f.ShowMe "Technical Files"
        Case "LeanManagement"
            Set f = New frmQMSDocuments_Ex
            Load f
            f.ShowMe "Lean Management"
                
    End Select
    
End Sub

Private Sub ShowSamplingForm(ITemKey As String)
                
    Dim f As Form
    Select Case Right(ITemKey, Len(ITemKey) - 3)
        
        Case "SampleCatalog"
            Set f = New frmItems
            Load f
            f.ShowMe 0, True
        Case "NewSampleOrder"
            Set f = New frmFNewCustOrder_Frgz
            Load f
            Call f.AddNew(True)
        Case "SampleOrderList"
            Set f = New frmFCustOrderList
            Call f.ShowMe(False, True)
        Case "PurchaseOrder"
            Set f = New frmNewRMPO
            Load f
            Call f.ShowMe(True)
        Case "PurchaseOrderList"
            Set f = New frmRMPOList
            Load f
            Call f.ShowMe(True)
        Case "ReceiveItems"
            Set f = New frmVendGateRcvd
            Load f
            Call f.ShowMe(True)
        Case "ReceiveList"
            Set f = New frmVendRcvList
            Load f
            f.ShowMe (True)
        Case "OpeningStock"
            frmItems_Opening_Locationwise.Show
        Case "Issuance"
            Set f = New frmIssItemsSimpleLocWise_Sampling
            Load f
            f.ShowMe (True)
        Case "IssuanceList"
            Set f = New frmStockTransaction
            Load f
            f.ShowMe (True)
        Case "Receiving"
            frmItems_Rcv_Sampling.Show
            
        Case "SamplingStock"
            Call PrintSamplingStockReport
            
        Case "SampleInvoiceList"
            Set f = New frmCustomInvoiceList
            Load f
            Call f.ShowMe(False, True)
            
        Case "SampleBilling"
            Set f = New frmVenderBilling
            Load f
            f.ShowMe (True)
        Case "SampleBillingList"
            Set f = New frmVenderBillingList
            Load f
            f.ShowMe (True)
    End Select
    
End Sub


Private Sub PrintPListForBank()

    Dim frmPrint As New frmPrint
    Dim rpt As CRAXDDRT.Report 'New rptCustomPList
    Set rpt = rptApp.OpenReport(RptPath & "\Bank_Packing_List.rpt") '
    
    rpt.FormulaFields(4).Text = True
    With frmPrint
        Set .rpt = rpt
        .cmbCust.ClearVals
        '.cmbCust.AddVals con, "Distinct OrderNo", "VFItemsQtyToPlan", "OrderNo", " Where RemQty>0"
        .cmbCust.AddItem "<All Customers>", "0"
        .cmbCust.AddVals con, "Distinct CustCode", "ForeignCustomers", "CustCode", " Where CustCode<>'Hico'"
        .SelFor = "{VCustomPList.CustomInvoice}='"
        .JobCondField = " CustCode "
        .JobField = " CustomInvoice "
        .JobIDField = " CustomInvoice "
        .JobTable = " CustomInvoice "
        .TheEnd = "'"
        .Show , MainForm
    End With
    
End Sub

Sub FillFixedAssetsItems(BarKey As String)

    On Error GoTo err
    With ExpBar
        Select Case Right(BarKey, Len(BarKey) - 2)
            Case "FixedAssets"
                .Bars(BarKey).Items.add , "AccFixedAssets", "Fixed Assets"
                .Bars(BarKey).Items.add , "AccFixedAssetsChangeHeads", "Change Heads"
                .Bars(BarKey).Items.add , "AccRepairLog", "Repair Log"
                .Bars(BarKey).Items.add , "AccFuelSlips", "Fuel Slips"
        End Select
        On Error GoTo err
        For i = 1 To .Bars(BarKey).Items.count
            .Bars(BarKey).Items(i).CanClick = UserHasAccessEX(.Bars(BarKey).Items(i).key)
        Next
    End With
    
    Exit Sub
err:
    MsgBox err.Description
    Resume Next
End Sub


Sub FillQMSItems_Bad(BarKey As String)

    On Error GoTo err
    With ExpBar
        Select Case Right(BarKey, Len(BarKey) - 3)
            Case "QMS"
                .Bars(BarKey).Items.add , "QMSCustomerComplaints", "Customer Complaints"
        End Select
        On Error GoTo err
        For i = 1 To .Bars(BarKey).Items.count
            .Bars(BarKey).Items(i).CanClick = UserHasAccessEX(.Bars(BarKey).Items(i).key)
        Next
    End With
    
    Exit Sub
err:
    MsgBox err.Description
    Resume Next
End Sub

Sub FillAccountsItems(BarKey As String)

    With ExpBar
        Select Case Right(BarKey, Len(BarKey) - 3)
            Case "Heads"
                .Bars(BarKey).Items.add , "AccNewAccount", "Define New Account"
                '.Bars(BarKey).Items.add , "AccNewExpAccount", "Define New Expense Account"
                .Bars(BarKey).Items.add , "AccPayables", "Accounts Payable"
'                If Not (strCompany = "Tecno" Or strCompany = "Dr-Frgz" Or strCompany = "Instrumed") Then
'                    .Bars(BarKey).Items.add , "AccVenderList", "Vender List"
'                End If
                
'                If Not (strCompany = "Instrumed" Or strCompany = "Scard") Then
'                    .Bars(BarKey).Items.add , "AccVenderAssRM", "Vender Assign RM"
'                End If
                
                .Bars(BarKey).Items.add , "AccReceivables", "Accounts Receivable"
                .Bars(BarKey).Items.add , "AccChartOfAccounts", "Chart Of Accounts"
                .Bars(BarKey).Items.add , "AccRe-Index", "Financial Re-Indexing"
                .Bars(BarKey).Items.add , "AccFinancialClosing", "Financial Closing"
                If strCompany = "Sunlike" Then
                    .Bars(BarKey).Items.add , "AccVouchersForAuthorization", "Vouchers for Authorization"
                End If
'                .Bars(BarKey).Items.Add , "AccSetDateRange", "Set Date Range"
                .Bars(BarKey).Items.add , "AccExpenseGroups", "Groups Accounts"
                .Bars(BarKey).Items.add , "AccChangeAccHeads", "Change Account Heads"
            Case "FinancialBudgeting"
                .Bars(BarKey).Items.add , "AccMonthlyBudget", "Monthly Budget"
                .Bars(BarKey).Items.add , "AccPrintMonthlyBudgetReport", "Print Report"
            Case "Makers"
                .Bars(BarKey).Items.add , "AccMakerList", "Maker List"
                .Bars(BarKey).Items.add , "PrdMakerAssItems", "Maker Assigned Items"
                
                .Bars(BarKey).Items.add , "AccMakerListNew", "Maker List (New)"
                If strCompany = "Tecno" Or strCompany = "Dr-Frgz" Or strCompany = "Instrumed" Then
                    .Bars(BarKey).Items.add , "AccVenderList", "Vender List"
                End If
            Case "Banks"
'                .Bars(BarKey).Items.add , "AccBankProfile", "Bank and Branch Profile"
'                .Bars(BarKey).Items.add , "AccNewBankAccount", "Define New Bank Account"
                .Bars(BarKey).Items.add , "BankList", "Bank List"
                .Bars(BarKey).Items.add , "AccChqBookDetail", "Bank Accounts && Chq. Books Detail"
                 
            Case "Vouchers"
                .Bars(BarKey).Items.add , "AccRV", "Receipt Voucher"
                .Bars(BarKey).Items.add , "AccPV", "Bank Payment Voucher"
                .Bars(BarKey).Items.add , "AccJV", "Journal Voucher"
                .Bars(BarKey).Items.add , "AccEV", "Cash Payment Voucher"
'                .Bars(BarKey).Items.Add , "PrlLongTermLoan", "New Long Term Loan"
'                .Bars(BarKey).Items.Add , "PrlShortTermLoan", "New Short Term Loan"
'                .Bars(BarKey).Items.Add , "AccClearAcrExpenses", "Clear Accrued Expenses"
'                .Bars(BarKey).Items.Add , "AccMonthlyAccured", "Monthly Accrued/Prepaid Expenses"
'                If Not (strCompany = "Instrumed") Then
'                    .Bars(BarKey).Items.add , "AccTransferFunds", "Transfer Funds"
'                End If
                
                If strCompany <> "Scard" Then
                    .Bars(BarKey).Items.add , "AccCustomInvoiceAuth", "Auth. Custom Invoice"
                End If
                
                If strCompany = "Banzai" Then
                    .Bars(BarKey).Items.add , "PrlLongTermLoan", "New Long Term Loan"
                    .Bars(BarKey).Items.add , "PrlLongTermLoanCheque", "New Long Term Loan (Cheque)"
                    .Bars(BarKey).Items.add , "PrlShortTermLoan", "New Short Term Loan"
                    .Bars(BarKey).Items.add , "PrlShortTermLoanCheque", "New Short Term Loan (Cheque)"
                    
                    .Bars(BarKey).Items.add , "AccMakerShortTerm", "Maker Short Term"
                    .Bars(BarKey).Items.add , "AccMakerLongTerm", "Maker Long Term"
                    .Bars(BarKey).Items.add , "PrdMakerLTLedger", "Maker Long Term Ledger"
                    
                    .Bars(BarKey).Items.add , "AccVenderShortTerm", "Vender Short Term"
                    .Bars(BarKey).Items.add , "AccVenderLongTerm", "Vender Long Term"
                End If
                
            Case "EmployeeAdvances"
                .Bars(BarKey).Items.add , "PrlShortTermSheet", "Short Term Sheet"
                .Bars(BarKey).Items.add , "PrlLongTermLoan", "New Long Term Loan"
                .Bars(BarKey).Items.add , "PrlLongTermLoanCheque", "New Long Term Loan (Cheque)"
                .Bars(BarKey).Items.add , "PrlShortTermLoan", "New Short Term Loan"
                .Bars(BarKey).Items.add , "PrlShortTermLoanCheque", "New Short Term Loan (Cheque)"
                
            Case "MakerAdvances"
                .Bars(BarKey).Items.add , "AccMakerShortTerm", "Vendor Short Term"
                .Bars(BarKey).Items.add , "AccMakerLongTerm", "Vendor Long Term"
                '.Bars(BarKey).Items.add , "PrdMakerLTLedger", "Vendor Long Term Ledger"
                .Bars(BarKey).Items.add , "PrdMakerSTLedger", "Vendor S.T Ledger"
                .Bars(BarKey).Items.add , "PrdMakerLTLedger", "Vendor L.T Ledger"
                .Bars(BarKey).Items.add , "PrdMakerLTBalances", "Vendor L.T Balances"
'                .Bars(BarKey).Items.add , "AccVenderShortTerm", "Vender Short Term"
'                .Bars(BarKey).Items.add , "AccVenderLongTerm", "Vender Long Term"
                
                .Bars(BarKey).Items.add , "PrdMakerAdvanceClearance", "Maker Adv. Clearance"
                .Bars(BarKey).Items.add , "PrdMakerLoanTypeTransfer", "Maker Loan Type Transfer"
            Case "Reports"
                .Bars(BarKey).Items.add , "AccLedger", "Detailed Account Ledger"
                .Bars(BarKey).Items.add , "AccCashBankStatus", "Cash && Bank Status"
                If (strCompany = "Instrumed") Then
                    '.Bars(BarKey).Items.add , "AccDueChqs", "Due Cheques"
                End If
               
               
                .Bars(BarKey).Items.add , "AccIncomeStatement", "Income Statement"
                .Bars(BarKey).Items.add , "AccIncomeStatementSummary", "Income Statement (Summary)"
                .Bars(BarKey).Items.add , "AccBalanceSheet", "Balance Sheet"
                .Bars(BarKey).Items.add , "AccBalanceSheetSummary", "Balance Sheet (Summary)"
                .Bars(BarKey).Items.add , "AccTBSummary", "Trial Balance (Summary)"
                .Bars(BarKey).Items.add , "AccTBDetail", "Trial Balance (Detail)"
                
                .Bars(BarKey).Items.add , "AccCashBook", "Cash Book Report"
                .Bars(BarKey).Items.add , "AccTransRpt", "Transaction Report"
                
'                If Not (strCompany = "Banzai" Or strCompany = "QEL" Or strCompany = "Instrumed" Or strCompany = "Scard") Then
'                    .Bars(BarKey).Items.add , "AccMakerPaymentSheet", "Maker Payment Sheet"
'                End If
                
                If Not (strCompany = "Instrumed" Or strCompany = "Banzai" Or strCompany = "Scard") Then
                    .Bars(BarKey).Items.add , "AccExpenseReport", "Expense Report"
                End If
                '.Bars(BarKey).Items.add , "AccVendorSummaryReport", "Vendor Payments Report"
                .Bars(BarKey).Items.add , "AccMakerPayableReport", "Maker Payable Report"
            Case "Authorizations"
                '.Bars(BarKey).Items.add , "AccCustomInvoiceAuth", "Authorize Custom Invoice"
'                '.Bars(BarKey).Items.add , "AccVendRcvAuth", "Vender Receving Authorization"
'                .Bars(BarKey).Items.Add , "AccVendORderFinal", "Make Orders Final"
'                .Bars(BarKey).Items.Add , "AccImpBillAuth", "Import Bill Authorization"
        End Select
        
        On Error GoTo err
        For i = 1 To .Bars(BarKey).Items.count
            .Bars(BarKey).Items(i).CanClick = UserHasAccessEX(.Bars(BarKey).Items(i).key)
        Next
    End With
'    Dim i As Integer
'    For i = 0 To ExpBar.Bars.count
'        ExpBar.Bars (i)
'    Next
Exit Sub
err:
MsgBox err.Description
Resume Next
    
End Sub

Private Sub ShowAccountsForm(ITemKey As String)

    Dim f As Form
    Select Case Right(ITemKey, Len(ITemKey) - 3)
        Case "NewAccount"
            frmNewAccount.Show 1
        Case "NewExpAccount"
            Load frmNewAccount
            With frmNewAccount
               .ExpAccount = True
               .Show 1
            End With
        Case "Payables"
            Dim fAccPay As New frmAccPayRcv
            With fAccPay
               .Tag = ACCPAYNO
               fAccPay.Show
            End With
        Case "VenderList"
            frmVenderList.Show
        Case "VenderAssRM"
            frmVendRMAssItems.Show
        Case "MakerList"
            frmMakerList.Show
        Case "MakerListNew"
            frmMakerList_NewComing.Show
        Case "Receivables"
            Dim fAccRcv As New frmAccPayRcv
            With fAccRcv
               .Tag = ACCRECNO
               fAccRcv.Show
            End With
        Case "ChartOfAccounts"
            'If getDBPassword("ChartOfAccounts") Then
                Load frmChartOfAccounts
                frmChartOfAccounts.Show
            'End If
'        Case "SetDateRange"
'            If getDBPassword("SetDateRange") Then
'                frmSetDateRange.Show 1
'            End If
        Case "Re-Index"
            frmCorrectBals.Show 1
        Case "FinancialClosing"
            frmFinancialYearClosing.Show 1
        Case "ExpenseGroups"
            frmAccountGroups.Show 1
        Case "ChangeAccHeads"
            frmChangeAccountCatagory.Show 1
        Case "MonthlyBudget"
            frmFinancialBudgeting_GroupWise.Show
        Case "PrintMonthlyBudgetReport"
            Call PrintMonthlyBudgetReport
        Case "FixedAssets"
            frmFixedAssets.Show
        Case "FixedAssetsChangeHeads"
            Set f = New frmChangeFixedAssetCatagory
            Load f
            f.ShowMe
        Case "RepairLog"
            frmFixedAsset_RepairLogs_List.Show
        Case "FuelSlips"
            frmFixedAsset_FuelSlips_List.Show
        Case "Calibrations"
            frmFixedAsset_Calibration_List.Show
        Case "VouchersForAuthorization"
            Dim fVForAuth As New frmVouchersForAuthorization
            With fVForAuth
               .Opts(0).Value = True
               fVForAuth.Show
            End With
        Case "MakerShortTerm"
            Set f = New frmMakerNewShortTermLoan
            Load f
            f.ShowMe
             
        Case "MakerLongTerm"
            Set f = New frmMakerNewLongTermLoan
            Load f
            f.ShowMe
        Case "VenderShortTerm"
            Set f = New frmVenderNewShortTermLoan
            Load f
            f.ShowMe
        Case "VenderLongTerm"
            Set f = New frmVenderNewLongTermLoan
            Load f
            f.ShowMe
        Case "BankProfile"
            frmBankList.Show 1
        Case "NewBankAccount"
            With frmNewAccount
                .chkBankAcc = vbChecked
                .Show 1
            End With
        Case "ChqBookDetail"
            frmChqBooks.Show
        Case "RV"
            'frmCashRV.Show
            frmRCV.Show
        Case "PV"
            frmCashPV.Show
        Case "JV"
            frmJV.Show
        Case "EV"
            frmExpenses.Show
'        Case "LongTermLoan"
'            Call ShowLongTermForm(False)
'        Case "ShortTermLoan"
'            Call ShowShortTermForm(False)
'        Case "ClearAcrExpenses"
'            ShowForm frmAcrExpenses, MainForm, , "Accured Expenses"
'        Case "MonthlyAccured"
'            ShowForm frmMonthlyAccured, MainForm, , "Accured Expenses"
        Case "TransferFunds"
            frmTransferFunds.Show 1
'        Case "DelVchr"
'            frmDeleteVoucher.Show 1
        Case "Ledger"
            frmLedger.Show
        Case "CompLedger"
            frmMonthLedger.Show
        Case "CashBankStatus"
            frmCashStatus.Show 1
        Case "DueChqs"
            frmDueChqs.Show
        Case "DueRcpts"
            Set f = New frmDueRcptsPymts
            With f
               Call .FillCmbs(False)
               f.Show
               .SetFocus
            End With
        Case "DuePymts"
            Dim f1 As New frmDueRcptsPymts
            With f1
               Call .FillCmbs(True)
               f1.Show
               .SetFocus
            End With
        Case "MonthTrans"
            frmComparison.Show
        Case "IncomeStatement"
            Call PrintIncomeStatement(True)
        Case "IncomeStatementSummary"
            Call PrintIncomeStatement(False)
        Case "BalanceSheet"
            Call PrintBalanceSheet(True)
        Case "BalanceSheetSummary"
            Call PrintBalanceSheet(False)
        Case "TBSummary"
            Dim f2 As New frmTrialBalance
            f2.Show
        Case "TBDetail"
            Dim f3 As New frmTrialBalance
            f3.chkDetail.Value = vbChecked
            f3.Show
        Case "PLSStandered"
            Dim f4 As New frmPLS
            With f4
                .Summary = False
                .Standered = True
                f4.Show
            End With
        Case "PLSSummary"
            Dim f5 As New frmPLS
            With f5
                .Summary = True
                .Standered = False
                f5.Show
            End With
        Case "PLSDetail"
            Dim f6 As New frmPLS
            With f6
                .Summary = False
                .Standered = False
                f6.Show
            End With
        Case "BalSheet"
            frmBalSheet.Show
        Case "CashBook"
            Dim f7 As New frmCashBookReport
            With f7
               .Opts(0).Value = True
               f7.Show
            End With
        Case "TransRpt"
            Dim f8 As New frmTransaction
            With f8
               .Opts(0).Value = True
               f8.Show
            End With
        Case "MakerPaymentSheet"
            Call PrintMakerPaymentSheet
        Case "ExpenseReport"
            Load frmExpenseReport
            frmExpenseReport.ShowMe (True)
        Case "VendorSummaryReport"
            Call PrintVendorSummaryReport
        Case "MakerPayableReport"
            Call PrintMakerPayableReport
'        Case "ForeignSales"
'            frmForeignSales.Show , MainForm
'        Case "PaymentSheet"
'            Call PrintPaymentSheet
'        '*******A U T H O R I Z A T I O N S***********
        Case "CustomInvoiceAuth"
            If getDBPassword("CustomInvoiceAuth") Then
                frmCustomInvoiceAuth.Show
            End If
'        Case "VendRcvAuth"
'            If getDBPassword("Others") Then
'                ShowForm frmVendRcvsToAuthorize, MainForm, , "Vender Receiving Authorization"
'            End If
'        Case "VendORderFinal"
'            ShowForm frmVendPOFinal, MainForm, , "Make Final"
'        Case "ImpBillAuth"
'            frmImpCost.Show 1
        End Select
    
End Sub

Private Sub ShowPayrollForm(ITemKey As String)

    Dim f As Form
    Select Case Right(ITemKey, Len(ITemKey) - 3)
        Case "CVsDetail"
            frmCVList.Show
        Case "NewCV"
            Set f = frmNewCV
            Load f
            f.AddNewCV
        Case "DeptList"
            DeptList.Show
        Case "NewDept"
            Set f = NewDepartment
            f.TableName = "Departments"
            f.lblCatID = GetNextDeptID("Departments")
            f.add = True
            f.Show 1
        Case "EmpList"
'            ShowForm EmployeeList, frmFlash, , "Employee List"
            Set f = New EmployeeList
            Load f
            f.ShowMe (False)
        Case "TempEmpList"
            Set f = New EmployeeList
            Load f
            f.ShowMe (True)
        Case "NewEmp"
            Set f = NewEmployee
            f.TableName = "Employees"
            f.add = True
            f.Show 1
        Case "EmpAssets"
            EmpAssets.Show
        Case "AssignAssets"
            AssignAssets.Show 1
        Case "EmpSalHistory"
            SalaryHistory.Show
        Case "EmpSalLedger"
            SalaryLedger.Show
'        Case "EmpBonus"
'            ShowForm frmEmpBonus, frmFlash, , "Emp. Bonus"
        Case "TransferEmp"
            frmEmpTransfer.Show 1
        Case "GuarantorList"
            frmGuarantorList.Show
        Case "PrintApplicationForm"
            Call PrintAppForm
        Case "PrintEmpDetails"
            Call PrintEmpDetails
        Case "PrintEmpDC"
            frmDCEmpList.Show
        '*************************************************************
        Case "EmpDC"
            frmDCEmp.Show
        '*************************************************************
'        Case "ProductivityEntries"
'            frmEmpProductivityEntries.Show
        Case "ProductivityEntries"
            Call PrintEmpProductivityReport
        Case "ProductivityMonthly"
            Call PrintEmpProductivityReportMonthly
        Case "ProductivityEntriesDaily"
            Call PrintEmpProductivityReport(True)
        Case "ProductivityEntriesOT"
            Call PrintEmpProductivityReport(, , True)
        Case "ProductivityEntriesDailyOT"
            Call PrintEmpProductivityReport(True, False, True)
        Case "EmpProdLineChart"
            Call PrintEmpProductivityReport(False, True)
        Case "ProcessWiseReport"
            Call PrintProcessWiseProductionReport
        Case "MakerPerformanceReport"
            Call PrintMakerPerformanceReportWrangler
        Case "MonthlyReport"
'            Call PrintMonthlyDateWiseProductionReport
            frmMonthlyProductionDashBoard.Show
        Case "ReadFromMachine"
            frmReadData.Show
        Case "TakeAttendance"
'            If getDBPassword("TakeAtt") Then
                Set f = New frmAtt
                Load f
                f.ShowMe True, False
'            ElseIf getDBPassword("TakeAttOT") Then
'                Set f = New frmAtt
'                Load f
'                f.ShowMe True, True
'            End If
        Case "TakeAttendanceEx"
'            If getDBPassword("TakeAtt") Then
            If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                Set f = New frmAttEx_Frgz
            Else
                Set f = New frmAttEx
            End If
                Load f
                f.ShowMe True, False
        Case "ReCalculateAttendance"
            Call ReCalculateAttendance
        Case "TakeAttendanceAuto"
            frmTakeAtt.Show 1
        Case "OverTimeEntry"
            frmOverTimeAuth.Show
'            End If
'        Case "TakeAttendanceCont"
'            If getDBPassword("TakeAttCont") Then
'                Set f = New frmAtt
'                Load f
'                f.ShowMe False
'            End If
'        Case "TakeAttSGuards"
'            If getDBPassword("TakeAtt") Then
'                frmTakeAttSGuards.Show 1
'            End If
'        Case "TakeAttendanceAuto"
'            If getDBPassword("TakeAtt") Then
'                frmTakeAtt.Show 1
'            End If
        Case "DailyActivitySheet"
            If strCompany = "Dr-Frgz" And strSubCompany = "Electro" Then
                frmEmpLedger_BioMetrics_Electro.Show
            ElseIf strCompany = "QEL" Or strCompany = "Dr-Frgz" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                frmEmpLedger_BioMetrics.Show
            Else
                frmEmpLedger.Show
            End If
        Case "EmpLedger"
            EmpLedger.Show
'        Case "EmpLedgerFake"
'            Load EmpLedger
'            EmpLedger.m_bShowFake = True
'            ShowForm EmpLedger, frmFlash, , "Employee Ledger"
        Case "Leaves"
            frmLeaves.Show
        Case "EditAttendance"
            frmAtt_Edit.Show
        Case "GatePassEntries"
'            If strCompany = "Tecno" Then
                frmGatePassEntry.Show
'            Else
'                frmShortTimeEntry.Show
'            End If
        Case "PerformanceDeductionAmt"
            frmEmpMonthlyPerformanceDeductionAmt.Show
        Case "RejectionCharging"
            frmEmployee_RejectionCharged.Show
        Case "EmpOnLeaves"
            frmEmpOnLeaves.Show
        Case "LeavesDetail"
            frmLeavesDetail.Show
'        Case "AuthorizeOvertime"
'            ShowForm frmOverTimeAuth, frmFlash, , "Authorize Over Time"
'        Case "OverTimeExcess"
'            ShowForm frmOverTimeExcess, frmFlash
'        Case "PrintOTSheet"
'            Call PrintOTSheet
'        Case "EmpWork"
'            frmEmpWorkDone.Show 1
'        Case "EmpLunch"
'            ShowForm frmEmpLunch, frmFlash, , "Employee Monthly Lunch Deduction"
        Case "ShortTermSheet"
            frmEmployees_ST_Sheet.Show
        Case "LongTermLoan"
            Call ShowLongTermForm(False)
        Case "LongTermLoanCheque"
            Call ShowLongTermForm(False, True)
        Case "ShortTermLoan"
            Call ShowShortTermForm(False)
        Case "ShortTermLoanCheque"
            Call ShowShortTermForm(False, True)
        Case "LoanBalance"
            Call PrintLoanBalance
        Case "LoanBalanceDateRange"
            Call PrintLoanBalance(False, True)
        Case "LoanBalanceTerminated"
            Call PrintLoanBalance(True)
        Case "ClearLoan"
            frmClearAmt.Show 1
        Case "ClearShortTermLoan"
            frmClearShortTerm.Show 1
        Case "LTEmpToEmpTransfer"
            Set f = frmTransferLTAdvance
            Load f
            f.ShowMe False
        Case "AdjDeduction"
            frmAdjustDedAmt.Show 1
        Case "ShortTermLoanLedger"
            frmShortTermLedger.Show
        Case "LongTermLoanLedger"
            frmLongLoanLedger.Show
        Case "AdvanceSalaryReport"
            frmSalarySheetAdvance.Show 1
        Case "EmpNewFine"
            Set f = New frmNewEmpFine
            Load f
            f.ShowMe (False)
        Case "EmpFineLedger"
            frmEmpFineLedger.Show
'        Case "FineBonus"
'            ShowForm frmFine, frmFlash, , "Fine/Bonus"
'        Case "Scholorships"
'            frmScholorShips.Show
'        Case "AdvanceSalary"
'            ShowForm frmAdvSalary, frmFlash, , "Advance Salary"
        Case "SalarySheetExt"
            frmSalarySheet.chkExternal = vbChecked
            frmSalarySheet.Show
        Case "HoldSalary"
            frmEmpHoldSalary.Show 1
        Case "Gratuity"
            frmGratuity_Calculation.Show
'        Case "SalarySheetFake"
'            frmSalarySheet.chkExternal = vbChecked
'            frmSalarySheet.m_bShowFake = True
'            ShowForm frmSalarySheet, frmFlash, , "Salary Sheet"
'        Case "OverTimePayments"
'            ShowForm frmPayOverTime, frmFlash, , "Overtime Payments"
'        Case "OverTimeSheet"
'            ShowForm frmOvertimeSheet, frmFlash, , "Overtime Sheet"
'        Case "ApplicationForm"
'            Screen.MousePointer = vbHourglass
'            If Not IsLoaded("EmployeeList") Then
'                EmployeeList.Visible = False
'            End If
'            EmployeeList.PrintAppForm
'            Screen.MousePointer = vbDefault
        Case "EmpListWithSal"
            If Not IsLoaded("EmployeeList") Then
                EmployeeList.Visible = False
            End If
            EmployeeList.PrintEmpList
        Case "EmpListWithoutSal"
            If Not IsLoaded("EmployeeList") Then
                EmployeeList.Visible = False
            End If
            Load EmployeeList
            EmployeeList.PrintEmpListWithoutSal
        Case "EmpCards"
            If Not IsLoaded("EmployeeList") Then
                EmployeeList.Visible = False
            End If
            EmployeeList.PrintEmpCard
        Case "AbsentSheet"
            AbsentSheet.Show
'        Case "DailyOTSheet"
'            Call PrintDailyOTSheet
'        Case "DailyLeavesReport"
'            Call PrintDailyLeaves
'        Case "JoinDisjoined"
'            PrintEmpJoinLeave.Show 1
'        Case "LeaveEncashment"
'            frmPrintLeaveEncashment.Show , frmFlash
'        Case "AttRegister"
'            PrintAttReg.Show 1
'        Case "MonthlySalComparison"
'            PrintSalaryComparison.Show 1
'        Case "OutstandingLoans"
'            PrintOutStanding.Show 1
'        Case "EmpLoanBalance"
'            Call PrintLoanBalance
        Case "SocialSecurity"
            SocialSecurity.Show
        Case "EOBI"
            EOBI.Show
        Case "EmpSalaryLedger"
            PrintEmpSalaryLedger
        Case "AttendanceStatus"
            Call PrintMonthlyAttendanceStatus
        Case "BonusCalculation"
            Call PrintMonthlyAttendanceStatus(True)
        Case "EmpWiseLeaves"
            Call PrintLeaveList
        Case "MissingOutTime"
            'Call PrintMissingAttendance
            frmPrintMissingOutTimes.Show 1
        Case "TerminatedEmpReport"
            Call PrintTerminatedEmpReport
        Case "NewHiredReport"
            Call PrintNewHiredReport
        Case "LateAttendanceReport"
            Call PrintLateAttendanceReport
        Case "OTPaymentSheet"
            Call PrintOverTimeSheet(1)
        Case "OTIndividualReport"
            Call PrintOverTimeIndividualReport
        Case "OTMonthlyRegister"
            Call PrintOTRegister
        Case "OTAnnualReportHrs"
            PrintOverTimeReportMonthwise (True)
        Case "OTAnnualReportAmt"
            PrintOverTimeReportMonthwise (False)
        Case "SalaryRevisions"
            Call PrintEmpListSalaryRevisions
        Case "ProbationEndCurrentMonth"
            Call PrintEmpListProbationEnd
        Case "EmpListGuarantorWise"
            Call PrintEmpListGuarantorWise
        Case "ManualAttendance"
            Call PrintManualAttendance
        Case "SecurityGuardsTiming"
            Call PrintSecuirtyGuardsTimings
        Case "EmpYearlyDaysReport"
            Call PrintEmpYearlyDaysReport
        Case "EmployeeTargets"
            Call PrintEmployeeTargets
'        Case "BankAccList"
'            ShowForm frmSalariedEmps, frmFlash, , "Salaried Employees"
'        Case "TempAdvAppSheet"
'            frmTempAdvSheet.chkExternal = vbChecked
'            ShowForm frmTempAdvSheet, frmFlash, , "Temporary  Advance Approval Sheet"
'        Case "BarcodeSettings"
'            frmTimeSettings.Show 1
        Case "DepartmentSettings"
            frmSettings.Show 1
        Case "PayrollPolicies"
            frmPayrollPolicies.Show 1
        Case "DesignationsList"
            'Call ShowDesignationList
            frmDeisgnationsList.Show
        Case "Holidays"
            Holidays.Show 1
'        Case "MachineSettings"
'            If VBA.Command$ = "FromVBIDE" Then
'                frmReadData.Show 1
'            Else
'                MsgBox "You don't have permission for it.", vbInformation
'                Exit Sub
'            End If
'        Case "Tax"
'            TaxRanges.Show 1
'        Case "UserManager"
'            frmUsers.Show 1
'        Case "PrintEmpLedger"
'            frmPrintEmpLedger.Show , frmFlash
            
        '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
        '''''''''''''''Contractors''''''''''''''''''''
        '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'        Case "ContractorList"
'            Set f = New EmployeeList
'            Load f
'            f.ShowMe (True)
'        Case "ContractorLedger"
'            ShowForm EmpPieceRateLedger, frmFlash, , "Contractor's Ledger"
'        Case "NewContractorJob"
'            frmContEmpWorkDone.Show 1
'        Case "PaymentSheet"
'            frmTempPaymentSheet.chkExternal = vbChecked
'            ShowForm frmTempPaymentSheet, frmFlash, , "Payment Sheet"
'        Case "ContNewShortTerm"
'            Call ShowShortTermForm(True)
'        Case "ContNewLongTerm"
'            Call ShowLongTermForm(True)
'        Case "ContEOBI"
'            ShowForm frmTempEOBI, frmFlash, , "Old Age"
'        Case "ContSS"
'            ShowForm frmTempSocialSecurity, frmFlash, , "Social Security"
    End Select
    
End Sub

Private Sub ShowStockForm(ITemKey As String)

    Dim f As Form
    Select Case Right(ITemKey, Len(ITemKey) - 3)
        Case "IssItems"
            If strCompany = "Dr-Frgz" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Banzai" Or strCompany = "Towne" Then
                frmIssItemsSimpleLocWise_Frgz.Show
            Else
                frmIssItemsSimple.Show
            End If
        Case "IssItemsList"
            frmIssdItemList.Show
        Case "RcvItems"
        
            If strCompany = "Tecno" Then
                frmRcvItemsSimpleTecno.Show
            ElseIf strCompany = "Dr-Frgz" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Banzai" Or strCompany = "Towne" Then
                frmRcvItemsSimple_Frgz.Show
            Else
                frmRcvItemsSimple.Show
            End If
            
        Case "RcvItemsList"
            frmRcvdItemList.Show
        Case "MaterialPlacement"
            Set f = New frmMaterialPlacement
            Load f
            f.ShowMe (False)
        Case "MaterialPlacementList"
            frmMaterialPlacementList.Show
        Case "ChangeMaterialPlacement"
            frmRMMovement.Show 1
        Case "StockTransaction"
            frmStockTransaction.Show
        Case "ItemLedger"
            If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Instrumed" Or strCompany = "Towne" Then
                Set f = New frmItemLedger_Frgz
            Else
                Set f = New frmItemLedger
            End If
            Load f
            Call f.ShowMe
            'frmItemLedger.Show
        Case "ChangeLocations"
            frmMaterialMovement.Show 1
        Case "MaterialGroup"
            Set f = New frmItemGroups
            Load f
            Call f.ShowForm("RMGroups", "RM Groups", "RM Name")
        Case "MaterialList"
            Set f = New frmRM
            Load f
            Call f.ShowMe(False)
        Case "AddNewRM"
            With frmNewRM
                Load frmNewRM
                Call .AddNew(False)
            End With
        Case "MaterialListSampling"
            Set f = New frmRM
            Load f
            f.ShowMe (True)
        Case "AssignRM"
            If getDBPassword("AssignRM") = False Then Exit Sub
            frmRMAssign.Show
        
        Case "PO"
            'frmVendNewPO.Show
            If strCompany = "Tecno" Then
                frmNewRMPO_Adv.Show
            Else
                frmNewRMPO.Show
            End If
        Case "POList"
            Set f = New frmRMPOList
            Load f
            f.ShowMe (bSampling)
        Case "OpenRMPOList"
            Call PrintOpenVendorOrders
        Case "PORcv"
            frmVendGateRcvd.Show
        Case "RcvList"
            frmVendRcvList.Show
        Case "VenderBilling"
            Set f = New frmVenderBilling
            Load f
            frmVenderBilling.ShowMe (False)
        Case "VenderBillingList"
            Set f = New frmVenderBillingList
            Load f
            f.ShowMe (False)
        Case "ChangeBatches"
            frmChangeBatchNo_RM.Show
        Case "ChangeBatchesLot"
            frmChangeBatchNo_Lot.Show
        Case "RMIssuance"
            If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                frmRMIssuanceLocWise.Show
            Else
                frmRMIssuance.Show
            End If
        Case "RMIssuanceList"
            frmRMIssList.Show
        Case "RMLedger"
            frmStockLedger.Show
        Case "PrintRMStockList"
            frmPrintRMList.Show 1
        Case "ChangeLocationsSF"
            frmMaterialMovementSF.Show 1
            
        Case "DeadStock"
            Call PrintDeadStockReport
            
        Case "Below"
            
        Case "GateIn"
            frmMiscGateInList.Show
            
        Case "BelowMinLevel"
            Call PrintFinishStockBelowMinLevelReport
        Case "AboveMinLevel"
            Call PrintFinishStockAboveMinLevelReport
            
        Case "ReceiveRejection"
            frmRejection_Receive.Show
        Case "RejectionTransaction"
            frmRejectionTransaction.Show
        Case "RejectionIssuance"
            frmRejection_Issuance.Show
        Case "RejectionAlerts"
            frmRejection_Store_Alerts.Show
    End Select
    
End Sub



Private Sub ResetnReAssignProcesses()
    
    If MsgBox("Are you Sure to Reset?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    
    If ValidateAdminPassword Then
    
        Call StartTrans(con)
        
        con.Execute "DELETE FROM ItemProcesses"
        
        con.Execute "INSERT INTO ItemProcesses(ProcessID,ItemID,Rate,SNo) " & _
            "SELECT ProcessID,ItemID,0,SNo FROM Items CROSS JOIN Processes ORDER BY ItemID,SNo"
         
         
         
        con.Execute "DELETE FROM VendAssItems WHERE VendID=" & iFactoryMakerID
    
        con.Execute "INSERT INTO VendAssItems(VendID,ProcessID,ItemID) " & _
        "SELECT DISTINCT " & iFactoryMakerID & ",ProcessID,ItemID FROM ItemProcesses " & _
        "WHERE ProcessID IN(SELECT ProcessID FROM Processes WHERE Operation IN(0,2)) " & _
        "ORDER BY ItemID,ProcessID"
        
        con.CommitTrans
        
        MsgBox "Done.", vbInformation
    End If
    
End Sub

Private Sub ShowDesignationList()
    Dim f As New frmMiscS_New
    Load f
    f.ShowForm "Designations", "Designation", "Designation", "EntryID", "Designations"
End Sub


Private Sub ShowLateOrderAlerts()
    'ShowPurchaseCalendarAlerts
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    With rs
        'AND LateOrderAlerts=1
        .Open "SELECT DISTINCT FCustomerOrders.Custcode,FCustomerOrders.Country,DeliveryDT,FCustomerOrders.OrderNo FROM FCustomerOrders INNER JOIN " & _
         "VUnshippedOrderList ON FCustomerOrders.OrderNo=VUnshippedOrderList.OrderNo INNER JOIN " & _
         "ForeignCustomers ON FCustomerOrders.CustCode=ForeignCustomers.CustCode AND " & _
         "FCustomerOrders.Country=ForeignCustomers.Country WHERE DeliveryDT<'" & Date & _
         "' AND LateOrderAlerts=1 ORDER BY FCustomerOrders.CustCode,FCustomerOrders.Country,DeliveryDT DESC", con, adOpenForwardOnly, adLockReadOnly

        i = .RecordCount
        If i = 0 Then
            Exit Sub
        End If
        
        iUpper = i

        strAlerts = ""
        i = 0
        Do Until .EOF
            If !CustCode <> strAlertCusts(i) Then
                If .AbsolutePosition > 1 Then
                    strAlertCusts(i) = strAlertCusts(i)
                    strAlertOrders(i) = Left(strAlertOrders(i), Len(strAlertOrders(i)) - 1)
 
                    i = i + 1
                End If
                
                strAlertCusts(i) = !CustCode
            End If
            strAlertOrders(i) = strAlertOrders(i) & !OrderNo & ","
            
            strAlerts = strAlerts & !CustCode & "(" & !OrderNo & "),"
            .MoveNEXT
        Loop
        
        .Close
    End With
    Set rs = Nothing
    strAlertOrders(i) = Left(strAlertOrders(i), Len(strAlertOrders(i)) - 1)
    
    strAlerts = Space(100) & Left(strAlerts, Len(strAlerts) - 1)
    lblAlerts.Caption = ""
    bMain = True
    iAlerts = 1
    iIndex = 0
    lblAlerts.Left = 1000
    
    PicAlerts.Visible = True
    TmrAlerts.Enabled = True
    
    
End Sub

Private Function GetLateOrders() As String
    'ShowPurchaseCalendarAlerts
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Dim strOrders As String, lCount As Long
    With rs
        'AND LateOrderAlerts=1
        .Open "SELECT DISTINCT FCustomerOrders.OrderNo FROM FCustomerOrders INNER JOIN " & _
         "VUnshippedOrderList ON FCustomerOrders.OrderNo=VUnshippedOrderList.OrderNo INNER JOIN " & _
         "ForeignCustomers ON FCustomerOrders.CustCode=ForeignCustomers.CustCode AND " & _
         "FCustomerOrders.Country=ForeignCustomers.Country WHERE DeliveryDT<'" & Date & _
         "' AND ForeignCustomers.CustCode NOT IN('STOCK','PARTS') AND LateOrderAlerts=1", con, adOpenForwardOnly, adLockReadOnly

        i = .RecordCount
        lCount = .RecordCount
        If i = 0 Then
            Exit Function
        End If
        
        strOrders = ""
        i = 0
'        Do Until .EOF
'            strOrders = strOrders & !OrderNo & ","
'            .MoveNEXT
'        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    If strOrders <> "" Then
        strOrders = Left(strOrders, Len(strOrders) - 1)
        strOrders = "Late Orders [" & lCount & "] " & "Order No. " & strOrders
        'strOrders = strOrders & "Order No. " & strOrders
    End If
    
    GetLateOrders = "Late Orders [" & lCount & "]" 'strOrders
    
End Function


Private Sub TmrAlerts_Timer()
    
    If DateDiff("n", DTLastAlert, Now) >= 30 Then 'Every 30 mins. or So, Update the Alerts
        Call ShowPurchaseCalendarAlerts
    End If
    
    Exit Sub
    
    If iDelayCounter > 10 Then
        
        If bMain Then
            lblAlerts.Caption = ""
            lblAlerts.Left = 690
            If strAlertCusts(iIndex) = "" Then
                Call ShowLateOrderAlerts
            End If
            lblAlertsMain.Caption = Mid$(strAlertCusts(iIndex), 1, iAlerts)
            iAlerts = iAlerts + 1
            If iAlerts > Len(strAlertCusts(iIndex)) Then
                bMain = False
                iAlerts = 1
            End If
        Else
            lblAlerts.Caption = Mid$(strAlertOrders(iIndex), 1, iAlerts)
            iAlerts = iAlerts + 1
            If iAlerts > 150 Then
                lblAlerts.Left = lblAlerts.Left - 130
                lblAlerts.Width = lblAlerts.Width + 130
            End If
            If iAlerts > Len(strAlertOrders(iIndex)) Then
                bMain = True
                iAlerts = 1
                
                If iIndex = iUpper Then
                    iIndex = 0
                Else
                    iIndex = iIndex + 1
                End If
                iDelayCounter = 0
            End If
        End If
    Else
        iDelayCounter = iDelayCounter + 1
    End If
    'lblAlerts.Caption = Mid$(strAlerts, iAlerts, 100)
    ' iAlerts = iAlerts + 1
     'lblAlerts.Left - 500
    
End Sub


Private Sub cmdHideExpBar_Click()
    Picture1.Visible = False
    PicShowExpBar.Visible = True
End Sub

Private Sub cmdShowExpBar_Click()
    PicShowExpBar.Visible = False
    Picture1.Visible = True
End Sub


Public Sub ShowPurchaseCalendarAlerts(Optional bLoadPurchaseCalendarAlerts As Boolean = False)
    '
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    
    If bLoadPurchaseCalendarAlerts Then
        With rs
            
            .Open "SELECT COUNT(*) FROM VPurchaseCalendar WHERE IssQty>RcvQty AND MasterPONo IS NOT NULL AND VendID<>207 " & _
             "AND ReturnDT='" & Date & "' AND CustCode IN(SELECT CustCode FROM Users_Customers WHERE UserID=" & UserID & ")", con, adOpenForwardOnly, adLockReadOnly
    
            i = .Fields(0).Value
            
            .Close
        End With
        strAlerts = "Purchase Calendar Alerts [" & i & "]"
    Else
        strAlerts = "[Click to Load]"
    End If
    
'    If i > 0 Then
'        ReDim strAlertCusts(1)
'        ReDim strAlertOrders(1)
'        iUpper = 1
'
'        strAlerts = ""
'
'        strAlerts = "Purchase Calendar Alerts [" & i & "]"
'    End If
    '
    
    
    
    Dim strLateOrders As String
    strLateOrders = GetLateOrders
    lblLateOrderAlert = strLateOrders
    'strAlerts = strAlerts & " " & strLateOrders
    
    
    If strCompany = "IAA" Or strCompany = "Kami" Then
        Dim strEventAlerts As String
        strEventAlerts = getEventAlerts
        lblCustomerEvents = strEventAlerts
        lblCustomerEvents.Visible = True
    ElseIf (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Towne" Then
        Dim strPlanningAlert As String
        strPlanningAlert = getInitialPlanningAlerts
        lblCustomerEvents = strPlanningAlert
        lblCustomerEvents.Visible = True
        lblCustomerEvents.Tag = "Planning"
    End If
    'strAlerts = strAlerts & " " & strEventAlerts
    
    PicAlerts.Visible = True
    
'    If Trim(strAlerts) = "" Then
'        Exit Sub
'    Else
    lblAlerts.Caption = strAlerts
'        bMain = True
'        iAlerts = 1
'        iIndex = 0
'    End If
    
    PicAlerts.Visible = True
        
    DTLastAlert = Now
    
End Sub

Private Function getEventAlerts() As String
   
    Dim rs As New ADODB.Recordset, strEvents As String
    Dim i As Integer, lCount As Long
    With rs
  
        .Open "SELECT Description FROM ForeignCustomers_Events WHERE EventDT='" & Date & "' AND RepeatedType=0" & _
        "UNION ALL SELECT Description FROM ForeignCustomers_Events WHERE CONVERT(Varchar(50),EventDT,11)='" & Format(Date, "yy/mm") & "' AND RepeatedType=1 " & _
        "UNION ALL SELECT Description FROM VCompanyCertificationExpiries WHERE DT<='" & Date & "'", con, adOpenForwardOnly, adLockReadOnly
       
        lCount = .RecordCount
        strEvents = ""
        
        Do Until .EOF
            strEvents = strEvents & !Description & ","
            .MoveNEXT
        Loop
        
        .Close
    End With
    
    Set rs = Nothing
    If strEvents <> "" Then
         strEvents = Left(strEvents, Len(strEvents) - 1)
         strEvents = "Events [" & lCount & "] " & strEvents
    End If
    getEventAlerts = "Events [" & lCount & "]" 'strEvents
    
End Function

Private Sub ShowLinks(strKey As String)
    
    If Left(strKey, 3) = "Acc" Then
        Call ShowAccountsForm(strKey)
        Exit Sub
    ElseIf Left(strKey, 3) = "Prl" Then
        Call ShowPayrollForm(strKey)
        Exit Sub
    ElseIf Left(strKey, 3) = "Stk" Then
        Call ShowStockForm(strKey)
        Exit Sub
   ElseIf Left(strKey, 3) = "Prd" Then
        Call ShowProductionForm(strKey)
        Exit Sub
    ElseIf Left(strKey, 2) = "QA" Then
        Call ShowQAForm(strKey)
        Exit Sub
    ElseIf Left(strKey, 3) = "Das" Then
        Call ShowDashBoardForm(strKey)
        Exit Sub
    ElseIf Left(strKey, 3) = "Hlp" Then
        Call ShowHelpForm(strKey)
        Exit Sub
    ElseIf Left(strKey, 3) = "QMS" Then
        Call ShowQMSForm(strKey)
        Exit Sub
    ElseIf Left(strKey, 3) = "SAM" Then
        Call ShowSamplingForm(strKey)
        Exit Sub
    ElseIf Left(strKey, 3) = "Imp" Then
        Call ShowImportForm(strKey)
        Exit Sub
    End If
    
    Dim rpt As CRAXDDRT.Report
    Dim frm As Form
    If LCase(CurrentUserName) = "production" Then
        If strKey <> "CompanyCatalog" Then
            Exit Sub
        End If
    End If
    Select Case strKey
        Case "PushToWeb"
            frmPushDataTomySQL.Show 1
        Case "CompanyInfo"
            frmCompany.Show
        Case "CompaniesList"
            frmCompanies.Show
        Case "BankList"
            frmBankList.Show 1
        Case "CompanyCatalog"
            If strCompany = "Tecno" Or strCompany = "Dr-Frgz" Then
                If getDBPassword("Customer List") = False Then
                    Exit Sub
                End If
            End If
            Set f = New frmItems
            Load f
            f.ShowMe 0
        Case "SampleCatalog"
            Set f = New frmItems
            Load f
            f.ShowMe 0, True
        Case "CompanyCatalogSurgical"
            Set f = New frmItems
            Load f
            f.ShowMe 1
        Case "CompanyCatalogDental"
            Set f = New frmItems
            Load f
            f.ShowMe 2
        Case "CompanyCatalogOthers"
            Set f = New frmItems
            Load f
            f.ShowMe 3
        Case "CompanyImportCatalog"
            Call ImportCompanyCatalog
        Case "ResetAllProcesses"
            Call ResetnReAssignProcesses
        Case "GroupList"
            
            Set f = New frmItemGroups_Ex
            Load f
            Call f.ShowForm("ItemGroups", "Item Groups", "Group Name")
        Case "AdditionalGroupList"
            Set f = New frmMiscS_New
            Load f
            Call f.ShowForm("ItemGroupsMain", "MainGroupName", "Group Name", "MainGroupID", "Additional Groups")
        Case "ItemTypes"
            Set f = New frmMiscS_New
            Load f
            Call f.ShowForm("ItemTypes", "Description", "Type", "ID", "Item Types")
        Case "ItemFinishedQuality"
            Set f = New frmMiscS_New
            Load f
            Call f.ShowForm("ItemsFinQuality", "Description", "Finished Quality", "ID", "Item Finished Quality")
        Case "PaymentTerms"
            Set f = New frmMiscS_New
            Load f
            Call f.ShowForm("PaymentTerms", "Description", "Payment Terms", "ID", "Payment Terms")
        Case "TradeTerms"
            Set f = New frmMiscS_New
            Load f
            Call f.ShowForm("TradeTerms", "Description", "Trade Terms", "ID", "Trade Terms")
'            frmItemGroups_Ex.Show 1
        Case "SROList"
            frmSROList.Show
        Case "DefineBox"
            frmBoxSize.Show 1
        Case "SteelVenders"
            frmSteelVends.Show
        Case "FDAProcesses"
            frmFDAProcesses.Show
        Case "Ports"
            frmPorts.Show 1
        Case "SteelList"
            frmSteelList.Show
        Case "GSPCodes"
            frmGSPCodes.Show 1
        Case "MeasuringUnits"
            Call ShowMeasuringUnits
        Case "ExchangeRates"
            frmCurrencyExchRates.Show 1
        Case "Stores"
            If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                Set f = New frmStoreRacks
                f.ShowMe
            Else
                frmStores.Show 1
            End If
        Case "AddStoresRacksBins"
            frmAddStoreRackBin.Show 1
        Case "ImportPictures"
            frmImportPictures.Show 1
        'Foreign Customers
        Case "FCustomers"
            If strCompany = "Tecno" Or strCompany = "Dr-Frgz" Then
                If getDBPassword("SpecialLinks") = False Then
                    Exit Sub
                End If
            End If
            frmFCustomerList.Show
        Case "MarketingCustomers"
            frmMarketingCustomerList.Show
        Case "AddNewCustomer"
            frmFCustomer.Show 1
        Case "CustomerCatalog"
            frmCustCatal.Show
        Case "CustomerComparison"
            frmCustomerComparison.Show 1
        Case "Letters"
            frmLettersList.Show
        'Orders
        Case "NewOrder"
            If (strCompany = "Tecno") Then
                Load frmFNewCustOrder
                frmFNewCustOrder.AddNew
            Else
                Load frmFNewCustOrder_Frgz
                frmFNewCustOrder_Frgz.AddNew
            End If
        Case "QuotationList"
            frmCustomerQuotationList.Show
        Case "NewStockOrder"
            frmFNewStockOrder.Show
        Case "OrderList"
            'frmFCustOrderList.Show
            Set f = New frmFCustOrderList
            Load f
            Call f.ShowMe(False, False, 0)
        Case "OrderListCustomers"
            Set f = New frmFCustOrderList
            Load f
            Call f.ShowMe(False, False, 1)
        Case "OrderListStock"
            Set f = New frmFCustOrderList
            Load f
            Call f.ShowMe(False, False, 2)
        Case "PrintLabels"
            frmPrintItemLabels.Show
        Case "OrderItemList"
            frmFCustOrderItemList.Show
        Case "OrderItemBalances"
            frmCustomerItemBalances.Show
        Case "AdvancePayments"
            frmAdvancePaymentList.Show
        Case "OrderStatus"
            frmFCustOrderStatus.Show
        Case "ArticleWiseStatus"
            Set f = New frmArticlewiseShippedStatus
            Load f
            f.ShowMe ("")
        
            'frmArticlewiseShippedStatus.Show
        'Proforma Invoice
        Case "PInvoiceList"
            frmFProformaList.Show
        Case "NewProforma"
            Load frmFNewProforma
            frmFNewProforma.Show
            frmFNewProforma.AddNewProforma
        'Agents
        Case "ShippingAgents"
            frmShippingAgents.Show
        Case "InsuranceAgents"
            frmInsuranceAgents.Show
            
        'Custom Documents
        
        Case "GenFormE"
            frmGenFormE.Show 1
        Case "CustomCovering"
            frmCustomCovering.Show
        Case "NewCustomInvoice"
            Set frm = New frmNewCustomInvoice_Adv
            frm.AddNewInvoice
        Case "NewCustomInvoiceEx"
            Set frm = New frmNewCustomInvoice_Adv
            frm.AddNewInvoice
        Case "CustomInvoice"
            Set frm = New frmCustomInvoiceList
            Load frm
            Call frm.ShowMe(True, False, 0)
        Case "SampleInvoiceList"
            'frmSampleInvoiceList.Show
            Set f = New frmCustomInvoiceList
            Load f
            Call f.ShowMe(False, True, 0)
        Case "NewPackingListM"
            frmNewCustomPListM_New.Show
        Case "NewPackingListA"
            frmNewCustomPListA.Show 1
        Case "CustomPackingList"
            Load frmCustomPList
            frmCustomPList.ShowForm (True)
        Case "CustomShipping"
            frmShippingInstructions.Show
        Case "CustomLabels"
            frmPackingLabels.Show 1
        Case "PrintInnerLabels"
            frmPrintInnerLabels.Show 1
        Case "FormE"
            frmPrintFormE.Show
        Case "PrintValuationForm"
            frmPrintValuationForm.Show , MainForm
        Case "GSP"
            frmGSP.Show , MainForm
        Case "CertofOrigin"
            frmCertOrigin.Show , MainForm
        Case "HidePackage"
            frmHidePackage.Show
            
        'Comm. Documents
        Case "ComCovering"
            frmCommercialCovering.Show
        Case "ComInvoice"
            Set frm = New frmCustomInvoiceList
            Load frm
            Call frm.ShowMe(False, False, 1)
        Case "ComPackingList"
            Load frmCustomPList
            frmCustomPList.ShowForm (False)
        Case "GSP"
            Dim FGSP As New frmPrint
            With FGSP
                Set .rpt = rptApp.OpenReport(RptPath & "\rptGSP.rpt") 'New rptGSP
                
                .cmbCust.ClearVals
                '.cmbCust.AddVals con, "Distinct OrderNo", "VFItemsQtyToPlan", "OrderNo", " Where RemQty>0"
                .cmbCust.AddItem "<All Customers>", "0"
                .cmbCust.AddVals con, "Distinct CustCode", "ForeignCustomers", "CustCode"
                .SelFor = "{VCustomInvoice.CustomInvoice}='"
                .JobCondField = " CustCode "
                .JobField = " CustomInvoice "
                .JobIDField = " CustomInvoice "
                .JobTable = " CustomInvoice "
                .TheEnd = "'"
                .Show , MainForm
            End With
        Case "FreightRate"
            frmCustomFreight.Show 1
        Case "CustomerLetter"
            Dim FCLetter As New frmPrint
            With FCLetter
                Set .rpt = rptApp.OpenReport(RptPath & "\rptWireTransfer.rpt") ''New rptWireTransfer
                .cmbCust.ClearVals
                '.cmbCust.AddVals con, "Distinct OrderNo", "VFItemsQtyToPlan", "OrderNo", " Where RemQty>0"
                .cmbCust.AddItem "<All Customers>", "0"
                .cmbCust.AddVals con, "Distinct CustCode", "ForeignCustomers", "CustCode"
                .SelFor = "{VrptProformas.PInvoice}='"
                .JobCondField = " CustCode "
                .JobField = " CustomInvoice "
                .JobIDField = " CustomInvoice "
                .JobTable = " CustomInvoice "
                .TheEnd = "'"
                .Show , MainForm
            End With
        'cGMP
            Case "DHRWiz"
                frmWizDHR.Show 1
            Case "DHR"
                frmDHR.Show
            Case "SteelReq"
                frmGMPs.Show
            Case "CalculationSheet"
                Dim MyPrintForm As New frmPrint
                With MyPrintForm
                    Set .rpt = rptApp.OpenReport(RptPath & "\rptSteelCalculation.rpt")    'New rptSteelCalculation
                    .cmbCust.ClearVals
                    '.cmbCust.AddVals con, "Distinct OrderNo", "VFItemsQtyToPlan", "OrderNo", " Where RemQty>0"
                    .cmbCust.AddItem "<All Customers>", "0"
                    .cmbCust.AddVals con, "Distinct CustCode", "ForeignCustomers", "CustCode"
                    .SelFor = "{VrptSteelCalculation.PInvoice}='"
                    .JobCondField = " CustCode "
                    .JobField = " PInvoice "
                    .JobIDField = " PInvoice "
                    .JobTable = " FPInvoice "
                    .TheEnd = "'"
                    .Show , MainForm
                End With
            'Bank Documents
            Case "PrintInvoice"
                Set frm = New frmCustomInvoiceList
                Load frm
                Call frm.ShowMe(False, False, 2)
'                Set rpt = rptApp.OpenReport(RptPath & "\Bank_Invoice.rpt")  '
'                Call PrintInvoiceForBank(rpt)
            Case "PrintPList"
                Call PrintPListForBank
            Case "LetterToBank"
                Set rpt = rptApp.OpenReport(RptPath & "\rptLetterToBank.rpt")  '
                Call PrintInvoice(rpt)
            Case "LetterToBankSimple"
                Set rpt = rptApp.OpenReport(RptPath & "\rptLetterToBankSimple.rpt")  '
                Call PrintInvoice(rpt)
            Case "BankSalesContract"
                Set rpt = rptApp.OpenReport(RptPath & "\Bank_Sales_Contract.rpt")  '
                Call PrintInvoiceForBank(rpt)
            Case "BIReasonForDelay"
                Set rpt = rptApp.OpenReport(RptPath & "\Bank_Reason_For_Delay.rpt")
                Call PrintInvoiceForBank(rpt)
            Case "BIReasonForDelaySBP"
                Set rpt = rptApp.OpenReport(RptPath & "\Bank_Reason_For_Delay_SBP.rpt")
                Call PrintInvoiceForBank(rpt)
            Case "BILabReport"
                Set rpt = rptApp.OpenReport(RptPath & "\Bank_Lab_Report.rpt")
                Call PrintInvoiceForBank(rpt)
            'Payment Status
            Case "CustPayStatus"
                frmCustomPayments.Show
            Case "AgingReport"
                Call PrintReceivableAgingReport
            Case "AssignFDBCNo"
                frmFCustFDBCNo.Show 1
            Case "RecCustPay"
                Load frmFCustRcvPymnts
                frmFCustRcvPymnts.ShowForm "", "", True
            Case "CustomerPayStatus"
                frmFCustPayments.Show
            Case "RecCustomerPay"
                Load frmFCustRcvPymnts
                frmFCustRcvPymnts.ShowForm "", "", False
            'Rebate
            Case "RebateDocuments"
                frmRebateA.Show
            Case "RebateStatus"
                frmRebateStatus.Show
            'Statistics
            Case "StatTotalExport"
                frmTotalExport.Show
            Case "StatTotalDemand"
                frmTotalDemand.Show
            Case "StatArticleWiesSales"
                frmArticlewiseSale.Show
            Case "StatArticleWiesOrders"
                frmArticlewiseOrders.Show
            Case "StatShippedUnShippedSummary"
                Call PrintShippedUnShippedSummary
            Case "StatExportPerformanceReport"
                frmPrintExportPerformanceReport.Show 1
            Case "StatPaymentLedger"
            'EE
            Case "StatExportStats"
                'frmExportStats.Show
            Case "EETotalExport"
                frmTotalExport.Show
                
            Case "EEPaymentLedger"
                
            Case "AddForEE"
                frmManInvsForEE.Show
            Case "AddBankNos"
                frmAddBankNo.Show 1
            Case "EE"
                frmEE.Show
            Case "MiscSentItems"
                frmMiscIssuanceList.Show
            'User Management
            Case "UserManagement"
                frmUsers.Show 1
            Case "UserActivity"
                Call Print_User_Activity_Report
            Case "Security"
                Dim strAdminPWD As String
                strAdminPWD = GetSingleStringValue("Password", "Users", " WHERE UserName='Administrator'")
                If getPasswordNew(strAdminPWD) Then
                    frmSecurityData.Show 1
                End If
            Case "Quit"
            On Error Resume Next
            If con.State = 1 Then con.Close
            Set con = Nothing
            End
    End Select
End Sub


Private Function getInitialPlanningAlerts() As String
   
    Dim rs As New ADODB.Recordset, strEvents As String
    Dim i As Integer, lCount As Long
    With rs
  
        .Open "SELECT * FROM FOrderItems_InitialPlanningMemo_Users WHERE UserName='" & CurrentUserName & "' AND Acknowledged=0", con, adOpenForwardOnly, adLockReadOnly
       
        lCount = .RecordCount
        strEvents = ""
        
'        Do Until .EOF
'            strEvents = strEvents & !Description & ","
'            .MoveNEXT
'        Loop
        
        .Close
    End With
    
    Set rs = Nothing
'    If strEvents <> "" Then
'         strEvents = Left(strEvents, Len(strEvents) - 1)
'         strEvents = "Planning [" & lCount & "] " & strEvents
'    End If
    getInitialPlanningAlerts = "Planning [" & lCount & "]" 'strEvents
    
End Function

Private Sub ShowSecurityForm()
    Dim strAdminPWD As String
    strAdminPWD = GetSingleStringValue("Password", "Users", " WHERE UserName='Administrator'")
    strAdminPWD = CryptRC4(FromHexDump(strAdminPWD), "AwmWrangler")
    If getPasswordNew(strAdminPWD) Then
        frmSecurityData.Show 1
    End If
End Sub

Private Sub FillImportItems(BarKey As String)
    'FillImportItems
    With ExpBar
        Select Case Right(BarKey, Len(BarKey) - 3)
            
            Case "Import"
            
                .Bars(BarKey).Items.add , "ImpAssignVendorItems", "Assign Items"
                .Bars(BarKey).Items.add , "ImpVendorOrder", "Vendor Order"
                .Bars(BarKey).Items.add , "ImpVendorOrderList", "Vendor Order List"
                .Bars(BarKey).Items.add , "ImpVendorOrderRcv", "Vendor Order Rcv"
                
        End Select
        On Error GoTo err
        For i = 1 To .Bars(BarKey).Items.count
            .Bars(BarKey).Items(i).CanClick = UserHasAccessEX(.Bars(BarKey).Items(i).key)
        Next
    End With
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub ShowImportForm(ITemKey As String)
    'ShowSamplingForm
    Dim f As Form
    Select Case Right(ITemKey, Len(ITemKey) - 3)
        
        Case "AssignVendorItems"
            frmVendor_Catalog.Show
        Case "VendorOrder"
            frmVendor_Order_AE.AddNew
        Case "VendorOrderList"
            frmVendor_Order_List.Show
        Case "VendorOrderRcv"
            frmVendor_Order_Receiving.Show
                
    End Select
    
End Sub

