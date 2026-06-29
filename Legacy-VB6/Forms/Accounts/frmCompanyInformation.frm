VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCompanyInformation 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stock Management Registeration"
   ClientHeight    =   4170
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6270
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Height          =   330
      Left            =   2550
      Top             =   4380
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   30
      Left            =   15
      TabIndex        =   12
      Top             =   15
      Width           =   15
   End
   Begin VB.Frame Ffull 
      BorderStyle     =   0  'None
      Height          =   4200
      Left            =   -15
      TabIndex        =   10
      Top             =   -45
      Width           =   6255
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next -->"
         Height          =   330
         Left            =   1335
         TabIndex        =   29
         Top             =   3810
         Width           =   1260
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Close"
         Height          =   330
         Left            =   4980
         TabIndex        =   28
         Top             =   3810
         Width           =   1260
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<--  Back"
         Height          =   330
         Left            =   30
         TabIndex        =   11
         Top             =   3810
         Width           =   1260
      End
      Begin VB.Frame f4 
         Height          =   3735
         Left            =   30
         TabIndex        =   31
         Top             =   15
         Visible         =   0   'False
         Width           =   6225
         Begin VB.PictureBox PicPass 
            Height          =   2280
            Left            =   180
            ScaleHeight     =   2220
            ScaleWidth      =   5835
            TabIndex        =   46
            Top             =   480
            Width           =   5895
            Begin VB.Frame Frame4 
               Height          =   1215
               Left            =   495
               TabIndex        =   47
               Top             =   810
               Width           =   4950
               Begin VB.TextBox txtpass 
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Index           =   0
                  Left            =   1980
                  MaxLength       =   50
                  PasswordChar    =   "*"
                  TabIndex        =   49
                  Top             =   270
                  Width           =   2670
               End
               Begin VB.TextBox txtpass 
                  Height          =   300
                  IMEMode         =   3  'DISABLE
                  Index           =   1
                  Left            =   1980
                  MaxLength       =   50
                  PasswordChar    =   "*"
                  TabIndex        =   48
                  Top             =   720
                  Width           =   2670
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  Caption         =   "Confirm Password"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
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
                  Left            =   240
                  TabIndex        =   51
                  Top             =   795
                  Width           =   1515
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  Caption         =   "Enter Password"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
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
                  Left            =   420
                  TabIndex        =   50
                  Top             =   315
                  Width           =   1335
               End
            End
            Begin VB.Image Image3 
               Height          =   480
               Left            =   195
               Picture         =   "frmCompanyInformation.frx":0000
               Top             =   210
               Width           =   480
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "This Password is given for Administrator Account. This Password will be used when you Log in first time."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   495
               Index           =   2
               Left            =   900
               TabIndex        =   52
               Top             =   180
               Width           =   4470
            End
         End
         Begin VB.OptionButton OptVF 
            Caption         =   "Monthly Voucher Format"
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
            Height          =   210
            Index           =   1
            Left            =   300
            TabIndex        =   44
            Tag             =   "The Voucher # is Generated On Monthly  Bases e.g. On  1st -Jul-2001 the First Voucher # will Be  RV-07-0001"
            Top             =   3240
            Width           =   2400
         End
         Begin VB.OptionButton OptVF 
            Caption         =   "Daily Voucher Format"
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
            Height          =   210
            Index           =   0
            Left            =   315
            TabIndex        =   43
            Tag             =   "The Voucher # is Generated On Daily Bases e.g. On  1st -Jul-2001 the First Voucher # will Be  RV-0701-001"
            Top             =   2925
            Value           =   -1  'True
            Width           =   2370
         End
         Begin VB.Frame Frame3 
            Height          =   2595
            Left            =   255
            TabIndex        =   32
            Top             =   240
            Visible         =   0   'False
            Width           =   5715
            Begin VB.TextBox txts 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   3
               Left            =   2040
               TabIndex        =   41
               Text            =   "Accounts Payable"
               Top             =   2055
               Width           =   3435
            End
            Begin VB.TextBox txts 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   2
               Left            =   2040
               TabIndex        =   35
               Text            =   "Bank Balances"
               Top             =   1740
               Width           =   3435
            End
            Begin VB.TextBox txts 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   1
               Left            =   2040
               TabIndex        =   34
               Text            =   "Accounts Receivable"
               Top             =   1425
               Width           =   3435
            End
            Begin VB.TextBox txts 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   2040
               TabIndex        =   33
               Text            =   "Cash In Hand"
               Top             =   1110
               Width           =   3435
            End
            Begin VB.Label lblAccno 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E9ECEF&
               Caption         =   "23-001"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   3
               Left            =   225
               TabIndex        =   42
               Top             =   2055
               Width           =   1800
            End
            Begin VB.Label lblAccno 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E9ECEF&
               Caption         =   "13-004"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   2
               Left            =   225
               TabIndex        =   39
               Top             =   1740
               Width           =   1800
            End
            Begin VB.Label lblAccno 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E9ECEF&
               Caption         =   "13-002"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   1
               Left            =   225
               TabIndex        =   38
               Top             =   1425
               Width           =   1800
            End
            Begin VB.Label lblAccno 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E9ECEF&
               Caption         =   "13-001"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   0
               Left            =   225
               TabIndex        =   37
               Top             =   1110
               Width           =   1800
            End
            Begin VB.Label Label3 
               Caption         =   "You Can Use The Custom Names for The Following Accounts According to Your Convention. "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   495
               Index           =   3
               Left            =   180
               TabIndex        =   36
               Top             =   240
               Width           =   5160
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00AFBDCB&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "          Account #                                  Account Title"
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
               Height          =   1560
               Index           =   2
               Left            =   210
               TabIndex        =   40
               Top             =   840
               Width           =   5310
            End
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E8EB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "The Voucher # is Generated On Daily Bases e.g. On  1st -Jul-2001 the First Voucher # will Be  RV-0701-001"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   810
            Index           =   4
            Left            =   2730
            TabIndex        =   45
            Top             =   2880
            Width           =   3255
         End
      End
      Begin VB.Frame F2 
         BackColor       =   &H8000000B&
         Height          =   3735
         Left            =   30
         TabIndex        =   13
         Top             =   15
         Visible         =   0   'False
         Width           =   6225
         Begin VB.TextBox txtFields 
            DataField       =   "Vformat"
            DataSource      =   "datPrimaryRS"
            Height          =   270
            Index           =   11
            Left            =   1485
            TabIndex        =   30
            Top             =   4710
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00FFFFFF&
            DataField       =   "CompanyName"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   0
            Left            =   195
            TabIndex        =   0
            Top             =   450
            Width           =   3375
         End
         Begin VB.TextBox txtFields 
            DataField       =   "City"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   2
            Left            =   165
            TabIndex        =   3
            Top             =   2190
            Width           =   2850
         End
         Begin VB.TextBox txtFields 
            DataField       =   "PhoneNo"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   3
            Left            =   165
            TabIndex        =   5
            Top             =   2700
            Width           =   1905
         End
         Begin VB.TextBox txtFields 
            DataField       =   "AltPhoneNo"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   4
            Left            =   2115
            TabIndex        =   6
            Top             =   2700
            Width           =   1905
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Faxno"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   5
            Left            =   4065
            TabIndex        =   7
            Top             =   2700
            Width           =   1905
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Address"
            DataSource      =   "datPrimaryRS"
            Height          =   900
            Index           =   6
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   990
            Width           =   5775
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Contperson"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   7
            Left            =   3060
            TabIndex        =   4
            Top             =   2190
            Width           =   2910
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Email"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   8
            Left            =   165
            TabIndex        =   8
            Top             =   3240
            Width           =   2895
         End
         Begin VB.TextBox txtFields 
            DataField       =   "url"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   9
            Left            =   3105
            TabIndex        =   9
            Top             =   3240
            Width           =   2880
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmCompanyInformation.frx":0442
            Left            =   3630
            List            =   "frmCompanyInformation.frx":0452
            TabIndex        =   1
            Top             =   435
            Width           =   2340
         End
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   1
            Left            =   3645
            TabIndex        =   14
            Top             =   450
            Visible         =   0   'False
            Width           =   2340
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Company Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   240
            TabIndex        =   24
            Top             =   225
            Width           =   1320
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Company Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   3690
            TabIndex        =   23
            Top             =   210
            Width           =   1260
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   210
            TabIndex        =   22
            Top             =   750
            Width           =   690
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "City"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   180
            TabIndex        =   21
            Top             =   1980
            Width           =   330
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone 1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   165
            TabIndex        =   20
            Top             =   2505
            Width           =   720
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone 2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   2100
            TabIndex        =   19
            Top             =   2505
            Width           =   720
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   4065
            TabIndex        =   18
            Top             =   2490
            Width           =   675
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Person"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   3045
            TabIndex        =   17
            Top             =   1980
            Width           =   1320
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   165
            TabIndex        =   16
            Top             =   3030
            Width           =   465
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "URL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   3075
            TabIndex        =   15
            Top             =   3030
            Width           =   390
         End
      End
      Begin VB.Frame F1 
         BackColor       =   &H00FFFFFF&
         Height          =   3780
         Left            =   15
         TabIndex        =   25
         Top             =   -30
         Width           =   6225
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Click Next To Proceed"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   60
            TabIndex        =   27
            Top             =   3435
            Width           =   1605
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   525
            Index           =   0
            Left            =   180
            TabIndex        =   26
            Top             =   2880
            Width           =   5940
         End
         Begin VB.Image Image2 
            Height          =   2385
            Left            =   15
            Picture         =   "frmCompanyInformation.frx":0499
            Stretch         =   -1  'True
            Top             =   120
            Width           =   6135
         End
      End
   End
End
Attribute VB_Name = "frmCompanyInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public FirstTime As Boolean

Private Sub cmdNext_Click()

Command1.Enabled = True
Dim RegNo As String

If F2.Visible Then
   'Combo1.ListIndex = 0
   'txtfields(0).SetFocus
   If Not Save Then Exit Sub
   
   F1.Visible = False
   F2.Visible = False
   f4.Visible = True
   
   PicPass.Visible = True
   cmdNext.Caption = "Finish"

ElseIf F1.Visible Then
   If datPrimaryRS.Recordset.RecordCount = 0 Then
      datPrimaryRS.Recordset.AddNew
   Else
      datPrimaryRS.Recordset.MoveFirst
   End If
 
   F1.Visible = False
   F2.Visible = True
   f4.Visible = False
   
 
ElseIf f4.Visible Then



   If PicPass.Visible Then
      If Not getAdminPassword Then Exit Sub
      Con.Execute "Update Users set [Password]='" & txtpass(0) & "" & "' Where UserName='Administrator'"
   End If
    Con.Execute "Update RegInfo set MonthlyVF=" & Abs(OptVF(1).value)
    Unload Me
End If


End Sub

Private Sub Combo1_Click()
txtfields(1) = Combo1.Text
End Sub

Private Sub Command1_Click()

If F2.Visible Then
 datPrimaryRS.Recordset.CancelUpdate
 F1.Visible = True
 F2.Visible = False
 f4.Visible = False
 Command1.Enabled = False
ElseIf f4.Visible Then
 F1.Visible = False
 F2.Visible = True
 f4.Visible = False
' OldF4.Visible = False
'ElseIf OldF4.Visible Then
' F1.Visible = False
' F2.Visible = False
' f4.Visible = False
 'OldF4.Visible = True
cmdNext.Caption = "Next -->"
End If
End Sub



Private Sub Command2_Click()

If MsgBox("Do You Want To Exit", vbQuestion + vbYesNo, "Registration") = vbYes Then
   If FirstTime Then Call CleanUp Else Unload Me
End If

End Sub

Private Sub Form_Activate()
   f4.Enabled = FirstTime
End Sub

Private Sub Form_Load()



txtfields(11).Text = ""
 With datPrimaryRS
 .ConnectionString = Con
 .RecordSource = "select * from regInfo"
 .Refresh

If .Recordset.RecordCount > 0 Then
   txtfields(0).Enabled = Not Registered(.Recordset![CompanyName], .Recordset![RegKey])
End If
End With
Label1(0).Caption = "Welcome To CTW Accounts Package Company Information..." & vbNewLine & _
                    "Please Fill You Company Information Which Will Be Used In CTW Accounting System. "

Me.Refresh
End Sub

Private Function Save() As Boolean
On Error GoTo ERR
Dim RegNo As String

If Len(txtfields(0)) < 4 Then
 MsgBox "Invalid Company Name...", vbInformation
 Save = False
 Exit Function
End If






If Trim(txtfields(0)) = "" Then
 MsgBox "You Must Enter Company Name.", vbCritical, "Registration"
 Save = False
 txtfields(0).SetFocus
 Exit Function
End If






For i = 1 To txtfields.Count - 3
 
 If Trim(txtfields(i)) = "" Then
  If MsgBox("You Have Not Filled All Information,However Do You Still Wish To Continue ?", vbYesNo + vbQuestion, "Registration") = vbNo Then
   Save = False
   'txtfields(i).SetFocus
   Exit Function
  Else
   Exit For
  End If
 
 End If
  
Next
datPrimaryRS.Recordset![Type] = Combo1.Text
datPrimaryRS.Recordset![RegKey] = RegNo
datPrimaryRS.Recordset.Update
company = datPrimaryRS.Recordset![CompanyName] & ""
COMPANYADDRESS = Replace(datPrimaryRS.Recordset![Address] & "", vbCrLf, " ")
Save = True

Exit Function
ERR:
   MsgBox "ERROR:" & ERR.Description, vbInformation
   Save = False

End Function

Private Sub Option1_Click()

txtfields(11).Text = Option1.Caption
Text1.Visible = True
Text2.Visible = False


End Sub

Private Sub Option2_Click()
txtfields(11).Text = Option2.Caption
Text1.Visible = False
Text2.Visible = True
End Sub





Private Sub OptVF_Click(Index As Integer)
Label3(4) = IIf(OptVF(0), OptVF(0).Tag, OptVF(1).Tag)
End Sub






Private Function getAdminPassword() As Boolean
If Not FirstTime Then GoTo OK
If Trim(txtpass(0)) = "" Then
 If MsgBox("No Password Given For Adminstrator Account. Do You Wish To Continue ?", vbQuestion + vbYesNo) = vbNo Then
  getAdminPassword = False
  Exit Function
 End If
ElseIf txtpass(0) <> txtpass(1) Then
 MsgBox "Password Not Confirmed", vbInformation
 txtpass(1).SelStart = 0
 txtpass(1).SelLength = Len(txtpass(1))
  txtpass(1).SetFocus
 getAdminPassword = False
End If
OK:
getAdminPassword = True
End Function
