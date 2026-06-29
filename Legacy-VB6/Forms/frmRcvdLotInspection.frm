VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRcvdLotInspection 
   BorderStyle     =   0  'None
   Caption         =   "Receiving Lots"
   ClientHeight    =   8895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14580
   DrawWidth       =   2
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   14580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00ECE0E5&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8895
      Index           =   0
      Left            =   0
      ScaleHeight     =   8895
      ScaleWidth      =   14565
      TabIndex        =   1
      Tag             =   "LOTS INSPECTION"
      Top             =   0
      Width           =   14565
      Begin VB.PictureBox PicTabs 
         Appearance      =   0  'Flat
         BackColor       =   &H00ECE0E5&
         ForeColor       =   &H80000008&
         Height          =   5085
         Index           =   0
         Left            =   75
         ScaleHeight     =   5055
         ScaleWidth      =   14190
         TabIndex        =   15
         Top             =   3720
         Width           =   14220
         Begin VB.CommandButton cmdClose 
            BackColor       =   &H00ECE0E5&
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   12750
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   4200
            Width           =   1410
         End
         Begin VB.TextBox txtTemperSNo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   390
            Left            =   11175
            TabIndex        =   36
            Top             =   420
            Width           =   990
         End
         Begin VB.TextBox txtTemperValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   390
            Left            =   12195
            TabIndex        =   35
            Top             =   420
            Width           =   1950
         End
         Begin VB.CommandButton cmdPrint 
            BackColor       =   &H00ECE0E5&
            Caption         =   "&Print"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   11190
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   4650
            Width           =   2970
         End
         Begin VB.TextBox txtRemarks 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   390
            Left            =   2475
            TabIndex        =   21
            Top             =   4455
            Width           =   4095
         End
         Begin VB.TextBox txtDisposation 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   390
            Left            =   30
            TabIndex        =   20
            Top             =   4455
            Width           =   2415
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7110
            TabIndex        =   19
            Top             =   885
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.CommandButton cmdSaveInspection 
            BackColor       =   &H00ECE0E5&
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   11205
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   4200
            Width           =   1470
         End
         Begin VB.OptionButton OptInspection 
            BackColor       =   &H00ECE0E5&
            Caption         =   "Pass"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   8340
            TabIndex        =   17
            Tag             =   "-1"
            Top             =   4485
            Width           =   1035
         End
         Begin VB.OptionButton OptInspection 
            BackColor       =   &H00ECE0E5&
            Caption         =   "Fail"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   9585
            TabIndex        =   16
            Top             =   4485
            Width           =   1200
         End
         Begin MSComCtl2.DTPicker DTInspection 
            Height          =   390
            Left            =   6630
            TabIndex        =   23
            Top             =   4455
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   688
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   306053123
            CurrentDate     =   40357
         End
         Begin MSComctlLib.ListView LV1 
            Height          =   3300
            Left            =   30
            TabIndex        =   24
            Top             =   840
            Width           =   11115
            _ExtentX        =   19606
            _ExtentY        =   5821
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Attributes Description"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "AQL"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Sample Size"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Reject On"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Test Spec. No."
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Act. Rej."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "Status"
               Object.Width           =   1764
            EndProperty
         End
         Begin MSComctlLib.ListView LVTemper 
            Height          =   3300
            Left            =   11190
            TabIndex        =   37
            Top             =   840
            Width           =   2970
            _ExtentX        =   5239
            _ExtentY        =   5821
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Sr #"
               Object.Width           =   1746
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Value"
               Object.Width           =   3440
            EndProperty
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sr #"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   16
            Left            =   11205
            TabIndex        =   39
            Top             =   105
            Width           =   435
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Value"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   17
            Left            =   12195
            TabIndex        =   38
            Top             =   105
            Width           =   615
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Inspected"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   7
            Left            =   6615
            TabIndex        =   29
            Top             =   4125
            Width           =   1650
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "General Remarks"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   6
            Left            =   2490
            TabIndex        =   28
            Top             =   4125
            Width           =   1920
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Material Disposation"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   8
            Left            =   15
            TabIndex        =   27
            Top             =   4125
            Width           =   2310
         End
         Begin MSForms.ComboBox cmbProcessNo 
            Height          =   390
            Left            =   30
            TabIndex        =   26
            Top             =   420
            Width           =   11115
            VariousPropertyBits=   746604569
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "19606;688"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontEffects     =   1073750016
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Process"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   9
            Left            =   15
            TabIndex        =   25
            Top             =   90
            Width           =   870
         End
      End
      Begin VB.TextBox txtTipSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   10170
         TabIndex        =   33
         Top             =   1860
         Width           =   1920
      End
      Begin VB.TextBox txtSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   8220
         TabIndex        =   31
         Top             =   1860
         Width           =   1920
      End
      Begin VB.OptionButton OptTabs 
         Appearance      =   0  'Flat
         BackColor       =   &H00ECE0E5&
         Caption         =   "Inspection"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   0
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3150
         Width           =   2355
      End
      Begin VB.TextBox txtLotBarcode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   510
         Left            =   5730
         TabIndex        =   0
         Top             =   1035
         Width           =   2010
      End
      Begin VB.TextBox txtItemCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   75
         TabIndex        =   7
         Top             =   1860
         Width           =   1950
      End
      Begin VB.TextBox txtItemName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   2040
         TabIndex        =   6
         Top             =   1860
         Width           =   6165
      End
      Begin VB.TextBox txtLotQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   75
         TabIndex        =   5
         Top             =   2700
         Width           =   1950
      End
      Begin VB.TextBox txtProcess 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   2040
         TabIndex        =   4
         Top             =   2700
         Width           =   9285
      End
      Begin VB.TextBox txtIssuanceDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   11340
         TabIndex        =   3
         Top             =   2700
         Width           =   2775
      End
      Begin VB.TextBox txtOrderNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   12105
         TabIndex        =   2
         Top             =   1860
         Width           =   2010
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tip Size"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   22
         Left            =   10170
         TabIndex        =   34
         Top             =   1485
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   18
         Left            =   8235
         TabIndex        =   32
         Top             =   1530
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lot No."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   6180
         TabIndex        =   14
         Top             =   630
         Width           =   1020
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Code"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   0
         Left            =   75
         TabIndex        =   13
         Top             =   1515
         Width           =   2010
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   1
         Left            =   2040
         TabIndex        =   12
         Top             =   1515
         Width           =   1230
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Lot Qty"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   2
         Left            =   75
         TabIndex        =   11
         Top             =   2355
         Width           =   2010
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Process"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   3
         Left            =   2025
         TabIndex        =   10
         Top             =   2355
         Width           =   870
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Receiving Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   4
         Left            =   11340
         TabIndex        =   9
         Top             =   2370
         Width           =   2010
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Order"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   5
         Left            =   12105
         TabIndex        =   8
         Top             =   1515
         Width           =   2010
      End
   End
End
Attribute VB_Name = "frmRcvdLotInspection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const lCaptionHeight = 600
Dim bMouseDown As Boolean
Dim lDownX As Long, lDownY As Long

Dim iColNo As Integer
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
    For i = 0 To PicWC.count - 1
        Call DrawCaptions(PicWC(i))
    Next
    
End Sub

Private Sub DrawCaptions(Pic As PictureBox)

    Dim lBorderColor As Long, lFillColor
    lBorderColor = RGB(128, 100, 162)
    lFillColor = RGB(128, 100, 162)
    Pic.Line (0, 20)-(Pic.Width, 20), lBorderColor
    
    If Pic.Tag <> "" Then
        Pic.Line (0, 350)-(Pic.Width, 350), lBorderColor
        
        Pic.Line (0, 50)-(Pic.Width, 600), lFillColor, BF
        Pic.ForeColor = vbWhite
        Pic.CurrentX = (Pic.Width - Pic.TextWidth(Pic.Tag)) / 2
        Pic.CurrentY = 180
        Pic.Print Pic.Tag
    End If
        
    Pic.Line (0, Pic.Height - 20)-(Pic.Width, Pic.Height - 20), lBorderColor 'Bottom

    Pic.Line (20, 0)-(20, Pic.Height), lBorderColor    'Left
    Pic.Line (Pic.Width - 20, 0)-(Pic.Width - 20, Pic.Height), lBorderColor  'Right

End Sub


Private Sub cmbProcessNo_Change()
    Call cmbProcessNo_Click
End Sub

Private Sub cmbProcessNo_Click()
    
    If cmbProcessNo.MatchFound = False Then Exit Sub
    On Error GoTo err
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        .Open "SELECT * FROM ProcessInspectionParameters WHERE ProcessID=" & cmbProcessNo.List(cmbProcessNo.ListIndex, 1) & " ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LV1.ListItems.Clear
        Do Until .EOF
            Set ITM = LV1.ListItems.add(, !EntryID & "'", !parameterName & "")
            ITM.ListSubItems.add , , !AQL & ""
            ITM.ListSubItems.add , , !SampleSize & ""
            ITM.ListSubItems.add , , !RejectOn & ""
            ITM.ListSubItems.add , , !TechSpecNo & "" 'cmbProcessNo.List(cmbProcessNo.ListIndex, 2)
            ITM.ListSubItems.add , , "0"
            ITM.ListSubItems.add , , "OK"
            .MoveNEXT
        Loop
        .Close
        
        .Open "SELECT * FROM VendRcvdDetailInspection WHERE VRD_RefID=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            LV1.Tag = Val(!EntryID & "")
            txtDisposation = !Disposation & ""
            txtRemarks = !Comments & ""
            DTInspection = !DT
            OptInspection(Abs(Not !LotStatus)).Value = True
            '------------------------------------
        Else
            LV1.Tag = "0"
            txtDisposation = ""
            txtRemarks = ""
            DTInspection = Date
        End If
        .Close
        
        If Val(LV1.Tag) > 0 Then
            .Open "SELECT * FROM VendRcvdDetailInspectionDetail WHERE VRDI_RefID=" & Val(LV1.Tag), con, adOpenForwardOnly, adLockReadOnly
            Do Until .EOF
                Set ITM = LV1.ListItems(!PIP_RefID & "'")
                ITM.SubItems(1) = !AQL & ""
                ITM.SubItems(2) = !SampleSize & ""
                ITM.SubItems(3) = !RejectOn & ""
                ITM.SubItems(4) = !TestSpecificationNo & ""
                ITM.SubItems(5) = !ActualRejection & ""
                ITM.SubItems(6) = !Status & ""
                .MoveNEXT
            Loop
            .Close
        End If
             
        .Open "SELECT * FROM VendRcvdDetailTemperValues WHERE VRD_RefID=" & lEntryID & " ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LVTemper.ListItems.Clear
        Do Until .EOF
            Set ITM = LVTemper.ListItems.add(, !EntryID & "'", Val(!SNo & ""))
            ITM.ListSubItems.add , , !TempValue & ""
            .MoveNEXT
        Loop
        .Close
        
    End With
    Set rs = Nothing
'    If Val(LV1.Tag) > 0 Then
'        cmdPrint.Visible = True
'    Else
'        cmdPrint.Visible = False
'    End If
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    
    On Error GoTo err
    'If Val(LV1.Tag) <= 0 Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\VRDInspection.rpt")
    Dim f As New frmPrevRpt
    Me.Visible = False
    f.ShowReport "{ProcessInspection.VRD_RefID}=" & lEntryID, rpt
    Unload Me
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub cmdSaveInspection_Click()

    On Error GoTo err
    If OptInspection(0).Tag = "-1" Then
        MsgBox "Please Select Pass / Fail", vbInformation
        Exit Sub
    End If
    
    Dim i As Integer
    
    Dim myDT As Date
    myDT = Date
    Call StartTrans(con)
     
    
    Dim lVRDI_EntryID As Long
    If Val(LV1.Tag) > 0 Then
        con.Execute "DELETE FROM VendRcvdDetailInspection WHERE EntryID=" & Val(LV1.Tag)
    End If
     
    con.Execute "INSERT INTO VendRcvdDetailInspection(VRD_RefID,Disposation,Comments,DT,LotStatus,UserName,MachineName) " & _
     "VALUES(" & lEntryID & ",'" & txtDisposation & "','" & txtRemarks & "','" & DTInspection & "'," & _
     IIf(OptInspection(0).Tag = "1", 0, 1) & ",'" & CurrentUserName & "','" & strComputerName & "')"
      
    lVRDI_EntryID = GetSingleLongValue("MAX(EntryID)", "VendRcvdDetailInspection", " WHERE MachineName='" & strComputerName & "'")
    For i = 1 To LV1.ListItems.count
        With LV1.ListItems(i)
            con.Execute "INSERT INTO VendRcvdDetailInspectionDetail(VRDI_RefID,PIP_RefID,AQL,SampleSize,RejectOn," & _
             "TestSpecificationNo,ActualRejection,Status) VALUES(" & lVRDI_EntryID & "," & Val(.key) & ",'" & _
             .SubItems(1) & "','" & .SubItems(2) & "','" & .SubItems(3) & "','" & .SubItems(4) & "','" & _
             .SubItems(5) & "','" & .SubItems(6) & "')"
        End With
    Next
    
    'Now Save Temper Values....
    con.Execute "DELETE FROM VendRcvdDetailTemperValues WHERE VRD_RefID=" & lEntryID
    For i = 1 To LVTemper.ListItems.count
        con.Execute "INSERT INTO VendRcvdDetailTemperValues(VRD_RefID,SNo,TempValue) VALUES(" & lEntryID & _
         "," & i & ",'" & LVTemper.ListItems(i).SubItems(1) & "')"
    Next
    
    con.CommitTrans
    LV1.Tag = lEntryID
    
    'Unload Me
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    Call DrawPicCaptions
    bMouseDown = False
End Sub


Private Sub LV1_DblClick()
    If LV1.ListItems.count = 0 Then Exit Sub
    If iColNo = 1 Then Exit Sub
    With txtEdit
        .Move LV1.Left + LV1.ColumnHeaders(iColNo).Left, LV1.Top + LV1.SelectedItem.Top, LV1.ColumnHeaders(iColNo).Width
        .Text = LV1.SelectedItem.SubItems(iColNo - 1)
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
    
End Sub

Private Sub LV1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim i As Integer
    For i = 1 To LV1.ColumnHeaders.count
        If x > Val(LV1.ColumnHeaders(i).Left) And x < Val(LV1.ColumnHeaders(i).Left) + Val(LV1.ColumnHeaders(i).Width) Then
            iColNo = i
            Exit For
        End If
    Next
    
End Sub


Private Sub LVTemper_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        LVTemper.ListItems.Remove LVTemper.SelectedItem.Index
    End If
End Sub

Private Sub OptInspection_Click(Index As Integer)
    OptInspection(0).Tag = Index
End Sub


Private Sub OptTabs_Click(Index As Integer)

    Dim i As Integer
    For i = 0 To PicTabs.count - 1
        PicTabs(i).Visible = False
    Next
    PicTabs(Index).Visible = True
    
End Sub


Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        LV1.SelectedItem.SubItems(iColNo - 1) = txtEdit
        If LV1.SelectedItem.Index < LV1.ListItems.count Then
            LV1.ListItems(LV1.SelectedItem.Index + 1).Selected = True
            LV1_DblClick
            Exit Sub
        ElseIf iColNo < LV1.ColumnHeaders.count Then
            iColNo = iColNo + 1
            LV1.ListItems(1).Selected = True
            Call LV1_DblClick
            Exit Sub
        End If
        
        txtEdit.Visible = False
        'Call txtLotBarcode_KeyPress(13)
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
    
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
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

Private Sub txtTemperValue_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
'        If cmbProcessNo.List(cmbProcessNo.ListIndex, 1) <> 168 Then
'            MsgBox "Not Available for this Process.", vbInformation
'            Exit Sub
'        End If
        If txtTemperValue = "" Then
            MsgBox "Invalid Value.", vbInformation
            Exit Sub
        End If
        
        Dim ITM As ListItem
        Set ITM = LVTemper.ListItems.add(, , LVTemper.ListItems.count + 1)
        ITM.ListSubItems.add , , txtTemperValue
        txtTemperSNo.Text = LVTemper.ListItems.count + 1
        With txtTemperValue
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    End If
    
End Sub

Public Sub ShowMe(p_lEntryID As Long)
    
    Call DrawPicCaptions
    
    lEntryID = p_lEntryID
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim strStatus As String, iStatus As Integer
    Dim i As Integer
    Dim lForeColor As Long
    Dim lNextToProcessSNo As Long
    
    With rs
        
        .Open "SELECT * FROM VVendRcvdDetail WHERE EntryID=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
        
        txtItemCode = !ItemCode & ""
        txtLotBarcode = !LotNo & ""
        txtItemCode.Tag = !OrderNo & ""
        txtItemName = !ItemName & ""
        txtLotQty = Val(!RcvdQty & "")
        txtProcess = !Description & ""
        txtProcess.Tag = Val(!ProcessID & "")
        txtIssuanceDate = Format(!DT, "dd-MMM-yyyy")
        txtLotBarcode.Tag = Val(!EntryID & "")
        txtOrderNo = !OrderNo & ""
        txtSize = !ItemSize & " " & !SizeUnit
        txtTipSize = !TipSize & ""
        .Close
            
        .Open "SELECT * FROM VItemProcesses WHERE IsExist='" & txtItemCode & "' AND ProcessID=" & Val(txtProcess.Tag), con, adOpenForwardOnly, adLockReadOnly
        cmbProcessNo.Clear
        Do Until .EOF
            cmbProcessNo.AddItem !code & " - " & !Description
            cmbProcessNo.List(cmbProcessNo.ListCount - 1, 1) = !ProcessID
            cmbProcessNo.List(cmbProcessNo.ListCount - 1, 2) = !code
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    If cmbProcessNo.ListCount > 0 Then cmbProcessNo.ListIndex = 0
    
    Me.Show 1
    
End Sub
