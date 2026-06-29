VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmRcvIPOLots 
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
      Tag             =   "LOTS RECEIVING"
      Top             =   0
      Width           =   14565
      Begin VB.PictureBox PicTabs 
         Appearance      =   0  'Flat
         BackColor       =   &H00ECE0E5&
         ForeColor       =   &H80000008&
         Height          =   5055
         Index           =   2
         Left            =   75
         ScaleHeight     =   5025
         ScaleWidth      =   14190
         TabIndex        =   16
         Top             =   3720
         Visible         =   0   'False
         Width           =   14220
         Begin VB.CommandButton cmdGenerateLots 
            BackColor       =   &H00ECE0E5&
            Caption         =   "Generate Lots"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   10485
            Style           =   1  'Graphical
            TabIndex        =   83
            Top             =   4260
            Width           =   2130
         End
         Begin VB.CommandButton cmdPrintRework 
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
            Height          =   600
            Left            =   12690
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   4260
            Width           =   1470
         End
         Begin VB.TextBox txtRejQtyRW 
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
            Left            =   3765
            TabIndex        =   41
            Top             =   4575
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.TextBox txtAccQtyRW 
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
            Left            =   2700
            TabIndex        =   40
            Top             =   4575
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.TextBox txtRWDone 
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
            Left            =   45
            TabIndex        =   39
            Top             =   4575
            Visible         =   0   'False
            Width           =   2640
         End
         Begin VB.TextBox txtRWReason 
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
            Left            =   4260
            TabIndex        =   47
            Top             =   495
            Width           =   4245
         End
         Begin VB.TextBox txtRWQty 
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
            Left            =   12765
            TabIndex        =   50
            Top             =   495
            Width           =   1395
         End
         Begin MSComctlLib.ListView LVRW 
            Height          =   3300
            Left            =   30
            TabIndex        =   43
            Top             =   900
            Width           =   14130
            _ExtentX        =   24924
            _ExtentY        =   5821
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   7461
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Object.Width           =   7488
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Width           =   7461
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Object.Width           =   1764
            EndProperty
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Repair Family"
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
            Index           =   28
            Left            =   8520
            TabIndex        =   82
            Top             =   165
            Width           =   1545
         End
         Begin MSForms.ComboBox cmbRepairFamily 
            Height          =   390
            Left            =   8520
            TabIndex        =   48
            Top             =   495
            Width           =   4230
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "7461;688"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qty Rej."
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
            Index           =   15
            Left            =   3780
            TabIndex        =   53
            Top             =   4260
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qty Acc."
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
            Index           =   14
            Left            =   2715
            TabIndex        =   52
            Top             =   4260
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rework Done by"
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
            Index           =   13
            Left            =   45
            TabIndex        =   51
            Top             =   4260
            Visible         =   0   'False
            Width           =   1860
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reason For Rework"
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
            Index           =   12
            Left            =   4275
            TabIndex        =   49
            Top             =   165
            Width           =   2235
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rework Qty"
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
            Index           =   11
            Left            =   12765
            TabIndex        =   46
            Top             =   180
            Width           =   1335
         End
         Begin MSForms.ComboBox cmbProcessRW 
            Height          =   390
            Left            =   15
            TabIndex        =   45
            Top             =   495
            Width           =   4230
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "7461;688"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
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
            Index           =   10
            Left            =   15
            TabIndex        =   44
            Top             =   150
            Width           =   870
         End
      End
      Begin VB.PictureBox PicTabs 
         Appearance      =   0  'Flat
         BackColor       =   &H00ECE0E5&
         ForeColor       =   &H80000008&
         Height          =   5085
         Index           =   0
         Left            =   75
         ScaleHeight     =   5055
         ScaleWidth      =   14190
         TabIndex        =   32
         Top             =   3720
         Width           =   14220
         Begin VB.TextBox txtRate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7080
            TabIndex        =   35
            Top             =   735
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.CommandButton cmdReceive 
            Caption         =   "&Receive"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   7080
            TabIndex        =   34
            Top             =   4320
            Width           =   1470
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "&Close"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   8640
            TabIndex        =   33
            Top             =   4320
            Width           =   1470
         End
         Begin ComboList.Usercontrol1 cmbMaker 
            Height          =   285
            Left            =   3300
            TabIndex        =   36
            Tag             =   "Planter"
            Top             =   840
            Visible         =   0   'False
            Width           =   3120
            _ExtentX        =   5503
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
         Begin MSComctlLib.ListView LV 
            Height          =   4095
            Left            =   60
            TabIndex        =   37
            Top             =   150
            Width           =   10035
            _ExtentX        =   17701
            _ExtentY        =   7223
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
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Process"
               Object.Width           =   4586
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Maker"
               Object.Width           =   6703
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Rate"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Proc."
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Hidden"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   4740
            Left            =   10110
            Stretch         =   -1  'True
            Top             =   165
            Width           =   4035
         End
      End
      Begin VB.PictureBox PicTabs 
         Appearance      =   0  'Flat
         BackColor       =   &H00ECE0E5&
         ForeColor       =   &H80000008&
         Height          =   5085
         Index           =   3
         Left            =   75
         ScaleHeight     =   5055
         ScaleWidth      =   14190
         TabIndex        =   38
         Top             =   3720
         Visible         =   0   'False
         Width           =   14220
         Begin VB.CommandButton cmdPrintRej 
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
            Height          =   600
            Left            =   12690
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   4200
            Width           =   1470
         End
         Begin VB.TextBox txtRejQty 
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
            Left            =   2880
            TabIndex        =   55
            Top             =   420
            Width           =   1395
         End
         Begin VB.TextBox txtReasonRJ 
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
            Left            =   4290
            TabIndex        =   54
            Top             =   420
            Width           =   9870
         End
         Begin MSComctlLib.ListView LVRJ 
            Height          =   3300
            Left            =   30
            TabIndex        =   57
            Top             =   840
            Width           =   14130
            _ExtentX        =   24924
            _ExtentY        =   5821
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   4974
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Object.Width           =   2461
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Width           =   17410
            EndProperty
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
            Index           =   21
            Left            =   15
            TabIndex        =   61
            Top             =   90
            Width           =   870
         End
         Begin MSForms.ComboBox cmbProcessRJ 
            Height          =   390
            Left            =   30
            TabIndex        =   60
            Top             =   420
            Width           =   2820
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4974;688"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rej. Qty"
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
            Index           =   20
            Left            =   2865
            TabIndex        =   59
            Top             =   105
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reason For Rejection"
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
            Index           =   19
            Left            =   4275
            TabIndex        =   58
            Top             =   105
            Width           =   2445
         End
      End
      Begin VB.PictureBox PicTabs 
         Appearance      =   0  'Flat
         BackColor       =   &H00ECE0E5&
         ForeColor       =   &H80000008&
         Height          =   5085
         Index           =   1
         Left            =   75
         ScaleHeight     =   5055
         ScaleWidth      =   14190
         TabIndex        =   17
         Top             =   3720
         Visible         =   0   'False
         Width           =   14220
         Begin VB.TextBox txtInsulationSNo 
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
            Left            =   11775
            TabIndex        =   78
            Top             =   885
            Width           =   570
         End
         Begin VB.TextBox txtInsulationValue 
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
            Left            =   12360
            TabIndex        =   80
            Top             =   885
            Width           =   1770
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
            Left            =   9960
            TabIndex        =   69
            Top             =   885
            Width           =   1785
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
            Left            =   9375
            TabIndex        =   68
            Top             =   885
            Width           =   570
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
            Height          =   600
            Left            =   12675
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   4245
            Width           =   1470
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
            TabIndex        =   23
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
            TabIndex        =   22
            Top             =   4455
            Width           =   2415
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7110
            TabIndex        =   21
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
            Height          =   600
            Left            =   11115
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   4245
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
            TabIndex        =   19
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
            TabIndex        =   18
            Top             =   4485
            Width           =   1200
         End
         Begin MSComCtl2.DTPicker DTInspection 
            Height          =   390
            Left            =   6630
            TabIndex        =   25
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
            Format          =   151257091
            CurrentDate     =   40357
         End
         Begin MSComctlLib.ListView LV1 
            Height          =   2790
            Left            =   30
            TabIndex        =   26
            Top             =   1305
            Width           =   9330
            _ExtentX        =   16457
            _ExtentY        =   4921
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
            Height          =   2790
            Left            =   9375
            TabIndex        =   65
            Top             =   1305
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   4921
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
               Object.Width           =   1005
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Value"
               Object.Width           =   2646
            EndProperty
         End
         Begin MSComctlLib.ListView LVIL 
            Height          =   2790
            Left            =   11775
            TabIndex        =   74
            Top             =   1305
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   4921
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
               Object.Width           =   1005
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Value"
               Object.Width           =   2646
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
            Index           =   27
            Left            =   11805
            TabIndex        =   81
            Top             =   570
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
            Index           =   26
            Left            =   12795
            TabIndex        =   79
            Top             =   570
            Width           =   615
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "INSULATION VALUES"
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
            Index           =   25
            Left            =   11775
            TabIndex        =   77
            Top             =   120
            Width           =   2370
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TEMPER VALUES"
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
            Index           =   24
            Left            =   9375
            TabIndex        =   76
            Top             =   120
            Width           =   2370
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "INSPECTION PARAMETERS"
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
            Index           =   23
            Left            =   30
            TabIndex        =   75
            Top             =   120
            Width           =   9330
         End
         Begin VB.Shape Shape3 
            Height          =   465
            Left            =   30
            Shape           =   4  'Rounded Rectangle
            Top             =   75
            Width           =   9330
         End
         Begin VB.Shape Shape2 
            Height          =   465
            Left            =   11775
            Shape           =   4  'Rounded Rectangle
            Top             =   75
            Width           =   2370
         End
         Begin VB.Shape Shape1 
            Height          =   465
            Left            =   9375
            Shape           =   4  'Rounded Rectangle
            Top             =   75
            Width           =   2370
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
            Left            =   10395
            TabIndex        =   67
            Top             =   570
            Width           =   615
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
            Left            =   9405
            TabIndex        =   66
            Top             =   570
            Width           =   435
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
            TabIndex        =   31
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
            TabIndex        =   30
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
            TabIndex        =   29
            Top             =   4125
            Width           =   2310
         End
         Begin MSForms.ComboBox cmbProcessNo 
            Height          =   390
            Left            =   30
            TabIndex        =   28
            Top             =   885
            Width           =   9330
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "16457;688"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
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
            TabIndex        =   27
            Top             =   555
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
         Locked          =   -1  'True
         TabIndex        =   72
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
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   1860
         Width           =   1920
      End
      Begin VB.OptionButton OptTabs 
         Appearance      =   0  'Flat
         BackColor       =   &H00ECE0E5&
         Caption         =   "Rejection Record"
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
         Index           =   3
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   3150
         Width           =   2355
      End
      Begin VB.OptionButton OptTabs 
         Appearance      =   0  'Flat
         BackColor       =   &H00ECE0E5&
         Caption         =   "Rework Record"
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
         Index           =   2
         Left            =   4905
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   3150
         Width           =   2355
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
         Index           =   1
         Left            =   2490
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   3150
         Width           =   2355
      End
      Begin VB.OptionButton OptTabs 
         Appearance      =   0  'Flat
         BackColor       =   &H00ECE0E5&
         Caption         =   "Issue / Receiving"
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
         TabIndex        =   15
         Top             =   3150
         Value           =   -1  'True
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
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
         TabIndex        =   73
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
         TabIndex        =   71
         Top             =   1530
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IPO Lot"
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
         Left            =   6135
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
         Caption         =   "Issuance Date"
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
         Top             =   2355
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
Attribute VB_Name = "frmRcvIPOLots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const lCaptionHeight = 600
Dim bMouseDown As Boolean
Dim lDownX As Long, lDownY As Long
Dim bRepairLot As Boolean, lRepairEntryID As Long
Dim iColNo As Integer

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

Private Sub cmbMaker_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmbMaker.MatchFound Then
            con.Execute "UPDATE Processes SET Fix_Maker_RefID=" & cmbMaker.ID & " WHERE ProcessID=" & LV.SelectedItem.Tag
            LV.SelectedItem.ListSubItems(1).Text = cmbMaker.Text
            LV.SelectedItem.ListSubItems(1).Tag = cmbMaker.ID
            cmbMaker.Visible = False
        End If
        Call txtLotBarcode_KeyPress(13)
    ElseIf KeyAscii = 27 Then
        cmbMaker.Visible = False
    End If
End Sub

Private Sub cmbMaker_LostFocus()
    cmbMaker.Visible = False
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
        
        .Open "SELECT * FROM ProcessInspection WHERE IPO_RefID=" & Val(txtLotBarcode) & " AND ProcessID=" & cmbProcessNo.List(cmbProcessNo.ListIndex, 1), con, adOpenForwardOnly, adLockReadOnly
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
            .Open "SELECT * FROM ProcessInspectionDetail WHERE PI_RefID=" & Val(LV1.Tag), con, adOpenForwardOnly, adLockReadOnly
            Do Until .EOF
                On Error Resume Next
                Set ITM = LV1.ListItems(!PIP_RefID & "'")
                If err.Number = 35601 Then  'Element Not Found
                    GoTo SKIPEntry:
                End If
                On Error GoTo 0
                ITM.SubItems(1) = !AQL & ""
                ITM.SubItems(2) = !SampleSize & ""
                ITM.SubItems(3) = !RejectOn & ""
                ITM.SubItems(4) = !TestSpecificationNo & ""
                ITM.SubItems(5) = !ActualRejection & ""
                ITM.SubItems(6) = !Status & ""
SKIPEntry:
                .MoveNEXT
            Loop
            .Close
        End If
             
        .Open "SELECT * FROM IPOTemperValues WHERE ProcessID=" & cmbProcessNo.List(cmbProcessNo.ListIndex, 1) & " AND IPO_RefID=" & txtLotBarcode & " ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
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

Private Sub cmdGenerateLots_Click()

    Dim i As Integer
    Dim lEntryID As Long
    Dim strLotNos As String
    Call StartTrans(con)
    For i = 1 To LVRW.ListItems.count
        With LVRW.ListItems(i)
            If Val(.ListSubItems(3).Tag) = 0 Then   'Lot Not Generated...
                con.Execute "INSERT INTO InternalProductionOrderDetailLots(IPOD_RefID,LotQty,RepairLot,IPO_Rework_RefID) " & _
                 "VALUES(" & Val(cmdGenerateLots.Tag) & "," & Val(.SubItems(3)) & ",1," & Val(.key) & ")"
                
                lEntryID = GetSingleLongValue("MAX(EntryID)", "InternalProductionOrderDetailLots")
                con.Execute "UPDATE IPOReworkRecord SET IPO_Repair_RefID=" & lEntryID & " WHERE EntryID=" & Val(.key)
                strLotNos = strLotNos & lEntryID & ","
            End If
        End With
    Next
    con.CommitTrans
    
    If strLotNos <> "" Then
        strLotNos = Left(strLotNos, Len(strLotNos) - 1)
        MsgBox "Following Lots have been generated." & vbNewLine & strLotNos, vbInformation
    End If
    
End Sub

Private Sub cmdPrint_Click()
    
    On Error GoTo err
    'If Val(LV1.Tag) <= 0 Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\QAIR.rpt")
    Dim f As New frmPrevRpt
    Me.Visible = False
    f.ShowReport "{ProcessInspection.IPO_RefID}=" & Val(txtLotBarcode), rpt, , , , Me
    
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub cmdPrintRej_Click()
    Call PrintReworkAndRejection(Val(txtLotBarcode))
End Sub

Private Sub cmdPrintRework_Click()
    Call PrintReworkAndRejection(Val(txtLotBarcode))
End Sub

Private Sub PrintReworkAndRejection(lIPO As Long)
        
    On Error GoTo err
    'If Val(LV1.Tag) <= 0 Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report, rptSub As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ReworkAndRejection.rpt")
    Set rptSub = rpt.OpenSubreport("ReworkAndRejectionSub")
    rptSub.RecordSelectionFormula = "{VInternalProductionOrderDetail.EntryID}=" & Val(txtLotBarcode)
    
    Dim f As New frmPrevRpt
    Me.Hide
    f.ShowReport "{VInternalProductionOrderDetail.EntryID}=" & Val(txtLotBarcode), rpt
    
    Unload Me
    
    Exit Sub
err:
    MsgBox err.Description

End Sub
Private Sub cmdReceive_Click()

    On Error GoTo err
'    If cmbProcessNo.ListCount > 0 Then
'        If OptInspection(0).Tag = "-1" Or LV1.Tag = "0" Then
'            MsgBox "Please Enter Inspection Details.", vbInformation
'            Exit Sub
'        End If
'    End If
    
    Dim i As Integer, lCount As Long
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .ListSubItems(3).Tag <> 1 Then 'Means It 0 or 2, Rcving Or Iss/Rcving.
                If CBool(.ListSubItems(4).Tag) Then 'Inspection Process, Check for Inspection Feeding.
                    lCount = GetSingleLongValue("COUNT(*)", "ProcessInspection", " WHERE ProcessID=" & Val(.Tag) & " AND IPO_RefID=" & Val(txtLotBarcode.Text))
                    If lCount = 0 Then
                        MsgBox "Inspection Not feeded for " & .SubItems(4) & " " & .Text
                        Exit Sub
                    End If
                End If
            End If
        End With
    Next
    
    Dim bDataNotComplete As Boolean
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .ForeColor = vbBlue Then
                If Val(.ListSubItems(1).Tag) = 0 Then
                    MsgBox "No Maker for " & .Text & " Process."
                    bDataNotComplete = True
                    Exit For
                ElseIf Val(.ListSubItems(2)) <= 0 Then
                    If MsgBox("No Rates for " & .ListSubItems(1).Text & vbNewLine & "Do you want to continue?", vbYesNo + vbQuestion) = vbNo Then
                        bDataNotComplete = True
                        Exit For
                    End If
                End If
            End If
        End With
    Next
    If bDataNotComplete Then
        Exit Sub
    End If
    
'    If Val(txtRepair) + Val(txtReject) > Val(txtLotQty) Then
'        MsgBox "Invalid Reject & Repair Qty.", vbInformation
'        Exit Sub
'    End If
    
    Dim myDT As Date
    myDT = Date
    Dim lRepairQty As Long, lRejectQty As Long
    Call StartTrans(con)
    For i = 1 To LV.ListItems.count
    
'        If i = LV.ListItems.count Then
'            lRepairQty = Val(txtRepair)
'            lRejectQty = Val(txtReject)
'        Else
'            lRepairQty = 0
'            lRejectQty = 0
'        End If
        lRepairQty = 0
        lRejectQty = 0
        With LV.ListItems(i)
            If Val(.ListSubItems(3).Tag) = 1 Then 'Issuance Only
                Call IssueLot(Val(txtLotBarcode.Text), txtItemCode.Text, Val(.Tag), myDT, Val(.ListSubItems(1).Tag), txtItemCode.Tag, Val(.SubItems(2)))
            ElseIf Val(.ListSubItems(3).Tag) = 2 Then 'Receiving Only
                Call ReceiveLot(Val(txtLotBarcode.Text), txtItemCode.Text, Val(.Tag), myDT, lRejectQty, lRepairQty)
            ElseIf Val(.ListSubItems(3).Tag) = 3 Then 'Issuance & Receiving
                Call IssueLot(Val(txtLotBarcode.Text), txtItemCode.Text, Val(.Tag), myDT, Val(.ListSubItems(1).Tag), txtItemCode.Tag, Val(.SubItems(2)))
                Call ReceiveLot(Val(txtLotBarcode.Text), txtItemCode.Text, Val(.Tag), myDT, lRejectQty, lRepairQty)
            End If
        End With
    Next
    con.CommitTrans

    cmdReceive.Enabled = False
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub IssueLot(lIPOLotNo As Long, strItemID As String, lProcessID As Long, myDT As Date, lMakerID As Long, strOrderNo As String, dRate As Double)

    Dim strSpecialInstructions As String, lExcess As Long, lRcvings As Long
    Dim iPriority As Integer
    
    lExcess = 100
    lRcvings = 50
    iPriority = 0

    iPriority = iPriority + 1
    
    Dim lMakerIssNo As Long, strMakerIssNo As String
    lMakerIssNo = Val(con.Execute("SELECT MAX(CAST(RIGHT(RecieptID,Len(RecieptID)-12) AS INT)) FROM VendIssued WHERE CONVERT(varchar,DT,6)='" & Format(myDT, "dd MMM yy") & "'").Fields(0).Value & "")
    lMakerIssNo = lMakerIssNo + 1
                
    strMakerIssNo = "M-ISU-" & Format(myDT, "ddMMyy") & lMakerIssNo
    
    If lMakerID = 0 Then lMakerID = iFactoryMakerID
     
    Dim strProcessCode As String, lProcessOrderNo As Long
    strProcessCode = GetSingleStringValue("Code", "Processes", " WHERE ProcessID=" & lProcessID)
    If Left(strProcessCode, 2) = "SC" Then
        lProcessOrderNo = GetSingleLongValue("MAX(ProcessOrderNo)", "VendIssued")
        lProcessOrderNo = lProcessOrderNo + 1
    Else
        lProcessOrderNo = 0
    End If
    con.Execute "INSERT INTO VendIssued(VendID,DT,RecieptID,UserID,ProcessID,ItemID," & _
     "UserName,MachineName,SpecialInstructions,ExcessQtyPercentage," & _
     "MaximumRcvingsAgainstPO,ProcessOrderNo) VALUES(" & lMakerID & ",'" & _
     myDT & "','" & strMakerIssNo & "'," & UserID & "," & _
     lProcessID & ",'" & strItemID & "','" & CurrentUserName & "','" & _
     strComputerName & "','" & strSpecialInstructions & "'," & lExcess & "," & _
     lRcvings & "," & lProcessOrderNo & ")"
          
    Dim lRefID As Long
    lRefID = Val(con.Execute("SELECT MAX(EntryID) FROM VendIssued").Fields(0).Value & "")
         
    Dim lLotNo As String
    lLotNo = GetSingleLongValue("LotNo", "VendIssdDetail", "WHERE IPODL_RefID=" & lIPOLotNo)

    Dim UserReqAuth As Integer
    UserReqAuth = 0
    con.Execute "INSERT INTO VendIssdDetail(RefID,RecieptID,ItemCode" & _
     ",Rate,IssQty,ReqAuth,OrderNo,RcvProcessID,ReturnDT,Priority,IPODL_RefID,LotNo) VALUES(" & lRefID & ",'" & _
     GetIssuanceIDDetail(myDT) & "','" & strItemID & "'," & dRate & "," & _
     Val(txtLotQty) & "," & UserReqAuth & ",'" & strOrderNo & "'," & lProcessID & _
     ",'" & DateAdd("d", 7, myDT) & "'," & iPriority & "," & lIPOLotNo & "," & lLotNo & ")"
             
    Dim cmd As New ADODB.Command
    With cmd
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_UpdateForIAR"    'Issue After Receive
        
        .Parameters("@ItemCode").Value = strItemID
        .Parameters("@ProcessID").Value = lProcessID
        .Parameters("@Qty").Value = Val(txtLotQty)
        .Parameters("@OrderNo").Value = strOrderNo
        
        .Execute
    End With
    'Authentication Info...
    If UserReqAuth = 0 Then
        Dim DetailRefID As Long
        DetailRefID = Val(con.Execute("SELECT MAX(EntryID) FROM VendIssdDetail").Fields(0).Value & "")
        con.Execute "INSERT INTO VendIssAuthDetail(UserID,IssID,DT) VALUES(" & _
         UserID & "," & DetailRefID & ",'" & ServerDate & " " & ServerTime & "')"
    End If
    
End Sub

Private Sub ReceiveLot(lIPOLotNo As Long, strItemID As String, lProcessID As Long, myDT As Date, Optional lRejectQty As Long = 0, Optional lRepairQty As Long = 0)
    
    Dim rs As New ADODB.Recordset
    Dim lMakerID As Long, lIssMainEntryID As Long, lIssEntryID As Long, lQty As Long, lLotNo As String
    Dim strOrderNo As String
    With rs
        .Open "SELECT * FROM VVendIssuanceIPOLots WHERE ItemCode='" & strItemID & "' AND ProcessID=" & lProcessID & " AND IPODL_RefID=" & lIPOLotNo, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            lMakerID = Val(!VendID & "")
            lIssMainEntryID = Val(!MainEntryID & "")
            lIssEntryID = Val(!EntryID & "")
            lQty = Val(!IssQty & "") - Val(!TotalReWorkQty & "") - Val(!TotalRejectQty & "")
            lLotNo = Val(!LotNo & "")
            strOrderNo = !OrderNo & ""
        End If
        .Close
    End With
    If lMakerID = 0 Then lMakerID = iFactoryMakerID
    
    con.Execute "INSERT INTO VendReceived(VendID,DT,RecieptID,UserID,ProcessID,Issuance_RefID) Values(" & lMakerID & _
     ",'" & myDT & " " & ServerTime & "',''," & UserID & "," & lProcessID & "," & lIssMainEntryID & ")"
    
    Dim iRefID As Integer
    iRefID = Val(con.Execute("SELECT MAX(EntryID) FROM VendReceived").Fields(0).Value & "")
    
    'Get The Next Process ID For Current Item...
    Dim lEntryID As Long
    lEntryID = iRefID
   
    Dim strRcvID As String
    Dim lRcvID As Long
    lRcvID = GetSingleLongValue("MAX(CAST(RIGHT(RecieptID,Len(RecieptID)-10) As Int))", "VVendRcvdItems", " WHERE YEAR(DT)=" & year(myDT) & " AND MONTH(DT)=" & Month(myDT) & " AND DAY(DT)=" & Day(myDT) & " AND LEFT(RecieptID,3)='RCV'")
    lRcvID = lRcvID + 1
    strRcvID = "RCV-" & Format(myDT, "ddMMyy") & lRcvID
     
    Dim cmd As New ADODB.Command
    If lQty > 0 Then
        '2. Now If Rcvd After Issued Then
        'Update The RcvdQty In VendIssdDetails....
        '-----------------------------------------------------------------
        With cmd
            .ActiveConnection = con
            .CommandType = adCmdStoredProc
        
            .CommandText = "SP_InsertVendRcvdNew"
            .Parameters("@ItemCode").Value = strItemID
            .Parameters("@ProcessID").Value = lProcessID
            .Parameters("@Qty").Value = lQty
            .Parameters("@WasteQty").Value = lRejectQty 'Val(LV.ListItems(i).SubItems(5) & "")
            .Parameters("@iRefID").Value = iRefID
            .Parameters("@VendID").Value = lMakerID
            .Parameters("@LotNo").Value = lLotNo
            .Parameters("@RcvDT").Value = myDT 'LV.ListItems(i).SubItems(7)
            
            .Parameters("@RecID").Value = strRcvID
            .Parameters("@OrderNo").Value = strOrderNo
            .Parameters("@DeliveryRefID").Value = lIssMainEntryID
            .Parameters("@CountedBy").Value = ""
            .Parameters("@Issue_RefID").Value = lIssEntryID
            .Parameters("@Weight").Value = 0 'Val(LV.ListItems(i).ListSubItems(7))
            .Parameters("@ReqAuth").Value = 0 'iAuthRequired 'UserReqAuth
            .Parameters("@UserID").Value = UserID
            .Parameters("@UserID").Value = lRepairQty
            .Execute
        End With
        
    End If
   
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
     
    Dim lEntryID As Long
    
    If Val(LV1.Tag) > 0 Then
        con.Execute "DELETE FROM ProcessInspection WHERE EntryID=" & Val(LV1.Tag)
    End If
     
    con.Execute "INSERT INTO ProcessInspection(IPO_RefID,ProcessID,Disposation,Comments,DT,LotStatus,UserName,MachineName) " & _
     "VALUES(" & txtLotBarcode.Text & "," & cmbProcessNo.List(cmbProcessNo.ListIndex, 1) & ",'" & txtDisposation & _
     "','" & txtRemarks & "','" & DTInspection & "'," & IIf(OptInspection(0).Tag = "1", 0, 1) & ",'" & CurrentUserName & "','" & strComputerName & "')"
      
    lEntryID = GetSingleLongValue("MAX(EntryID)", "ProcessInspection", " WHERE MachineName='" & strComputerName & "'")
    For i = 1 To LV1.ListItems.count
        With LV1.ListItems(i)
            con.Execute "INSERT INTO ProcessInspectionDetail(PI_RefID,PIP_RefID,AQL,SampleSize,RejectOn," & _
             "TestSpecificationNo,ActualRejection,Status) VALUES(" & lEntryID & "," & Val(.key) & ",'" & _
             .SubItems(1) & "','" & .SubItems(2) & "','" & .SubItems(3) & "','" & .SubItems(4) & "','" & _
             .SubItems(5) & "','" & .SubItems(6) & "')"
        End With
    Next
    
    'Now Save Temper Values....
    con.Execute "DELETE FROM IPOTemperValues WHERE IPO_RefID=" & txtLotBarcode.Text & " AND ProcessID=" & cmbProcessNo.List(cmbProcessNo.ListIndex, 1)
    For i = 1 To LVTemper.ListItems.count
        con.Execute "INSERT INTO IPOTemperValues(IPO_RefID,ProcessID,SNo,TempValue) VALUES(" & txtLotBarcode.Text & _
         "," & cmbProcessNo.List(cmbProcessNo.ListIndex, 1) & "," & i & ",'" & LVTemper.ListItems(i).SubItems(1) & "')"
    Next
    
    'Now Save Insulation Values....
    con.Execute "DELETE FROM IPOInsulationLeakageValues WHERE IPO_RefID=" & txtLotBarcode.Text & " AND ProcessID=" & cmbProcessNo.List(cmbProcessNo.ListIndex, 1)
    For i = 1 To LVIL.ListItems.count
        con.Execute "INSERT INTO IPOInsulationLeakageValues(IPO_RefID,ProcessID,SNo,LeakageValue) VALUES(" & txtLotBarcode.Text & _
         "," & cmbProcessNo.List(cmbProcessNo.ListIndex, 1) & "," & i & ",'" & LVIL.ListItems(i).SubItems(1) & "')"
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
    
    Call AddToCombo(cmbRepairFamily, "RepairType", "RepairTypes", "", False, "EntryID")
End Sub

Private Sub LV_DblClick()
    If LV.ListItems.count = 0 Then Exit Sub
    If iColNo = 2 Then  'Change Fix Maker for Process
        cmbMaker.Move LV.Left + LV.ColumnHeaders(2).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(2).Width
        cmbMaker.ClearVals
        cmbMaker.AddVals con, "VenderName", "Makers", "VendID", " WHERE VendID IN(SELECT VendID FROM MakerProcesses WHERE ProcessID=" & LV.SelectedItem.Tag & ")"
        cmbMaker.Visible = True
        cmbMaker.SetFocus
    ElseIf iColNo = 3 Then  'Change Rate of Maker.. Assign if not already...
        With txtRate
            .Move LV.Left + LV.ColumnHeaders(3).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(3).Width
            .Text = LV.SelectedItem.SubItems(2)
            .SelStart = 0
            .SelLength = Len(.Text)
            .Visible = True
            .SetFocus
        End With
    End If
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        If x > Val(LV.ColumnHeaders(i).Left) And x < Val(LV.ColumnHeaders(i).Left) + Val(LV.ColumnHeaders(i).Width) Then
            iColNo = i
            Exit For
        End If
    Next
    
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


Private Sub LVRJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If LVRJ.ListItems.count = 0 Then Exit Sub
        If MsgBox("Are you sure to delete?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        con.Execute "DELETE FROM IPORejectionRecord WHERE EntryID=" & Val(LVRJ.SelectedItem.key)
        LVRJ.ListItems.Remove LVRJ.SelectedItem.Index
    End If
End Sub

Private Sub LVRW_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If LVRW.ListItems.count = 0 Then Exit Sub
        If MsgBox("Are you sure to delete?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        con.Execute "DELETE FROM IPOReworkRecord WHERE EntryID=" & Val(LVRW.SelectedItem.key)
        LVRW.ListItems.Remove LVRW.SelectedItem.Index
    End If
End Sub

Private Sub LVTemper_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        LVTemper.ListItems.Remove LVTemper.SelectedItem.Index
    End If
End Sub

Private Sub OptInspection_Click(Index As Integer)
    OptInspection(0).Tag = Index
End Sub

Private Sub Text5_Change()

End Sub

Private Sub OptTabs_Click(Index As Integer)

    Dim i As Integer
    For i = 0 To PicTabs.count - 1
        PicTabs(i).Visible = False
    Next
    PicTabs(Index).Visible = True
    
End Sub

Private Sub Text2_Change()

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


Private Sub txtInsulationValue_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If cmbProcessNo.List(cmbProcessNo.ListIndex, 1) <> iInsulationValuesProcessID Then
            MsgBox "Not Available for this Process.", vbInformation
            Exit Sub
        End If
        If txtInsulationValue = "" Then
            MsgBox "Invalid Value.", vbInformation
            Exit Sub
        End If
        
        Dim ITM As ListItem
        Set ITM = LVIL.ListItems.add(, , LVIL.ListItems.count + 1)
        ITM.ListSubItems.add , , txtInsulationValue
        txtInsulationSNo.Text = LVIL.ListItems.count + 1
        With txtInsulationValue
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    End If
    
End Sub

Private Sub txtLotBarcode_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        Dim rs As New ADODB.Recordset
        Dim ITM As ListItem
        Dim strStatus As String, iStatus As Integer
        Dim i As Integer
        Dim lForeColor As Long
        Dim lNextToProcessSNo As Long
        
        With rs
            
            .Open "SELECT * FROM VVendIssuanceIPOLots WHERE IPODL_RefID=" & Val(txtLotBarcode) & " AND EntryID=(SELECT MAX(EntryID) FROM VendIssdDetail WHERE IPODL_RefID=" & Val(txtLotBarcode) & ")", con, adOpenForwardOnly, adLockReadOnly
            If .EOF Then
                MsgBox "Invalid Barcode.", vbInformation
                Exit Sub
            Else
                txtItemCode = !ItemCode & ""
                txtItemCode.Tag = !OrderNo & ""
                txtItemName = !ItemName & ""
                txtLotQty = Val(!IssQty & "") - Val(!TotalReWorkQty & "") - Val(!TotalRejectQty & "")
                txtProcess = !Description & ""
                txtProcess.Tag = Val(!ProcessID & "")
                txtIssuanceDate = Format(!DT, "dd-MMM-yyyy")
                txtLotBarcode.Tag = Val(!EntryID & "")
                txtOrderNo = !OrderNo & ""
                txtSize = !ItemSize & " " & !SizeUnit
                txtTipSize = !TipSize & ""
                cmdGenerateLots.Tag = Val(!IPOD_RefID & "")
                bRepairLot = !ReWorklot
                lRepairEntryID = Val(!Repair_RefID & "")
                .Close
                
                .Open "SELECT COUNT(Issue_RefID) AS Issue_RefID FROM VendRcvdDetail WHERE Issue_RefID=" & Val(txtLotBarcode.Tag), con, adOpenForwardOnly, adLockReadOnly
                If Val(!Issue_RefID & "") > 0 Then
                    MsgBox "Already Received.", vbInformation
                    'Exit Sub
                    cmdReceive.Enabled = False
                Else
                    cmdReceive.Enabled = True
                End If
                
                .Close
        
                'Now Fill Next Processes in ListView upto Next Scanning.
                Dim cmd As New ADODB.Command
                With cmd
                    Set .ActiveConnection = con
                    .CommandType = adCmdStoredProc
                    .CommandText = "IPOLotsReceiving_SP"
                    .Parameters("@ItemCode").Value = txtItemCode.Text
                    .Parameters("@ProcessID").Value = Val(txtProcess.Tag)
                    If bRepairLot Then
                        .Parameters("@RepairType").Value = lRepairEntryID
                    Else
                        .Parameters("@RepairType").Value = 0
                    End If
                    Dim rs1 As New ADODB.Recordset
                    Set rs1 = .Execute
                End With
                
                LV.ListItems.Clear
                cmbProcessNo.Clear
                Do Until rs1.EOF
                    Set ITM = LV.ListItems.add(, , rs1!Description & "")
                    ITM.Tag = Val(rs1!ProcessID & "")
                    If Val(rs1!Operation & "") = 1 Then
                        ITM.ListSubItems.add(, , rs1!VenderName & "").Tag = Val(rs1!VendID & "")
                        lForeColor = vbBlue
                    Else
                        ITM.ListSubItems.add(, , "Employee").Tag = "0"
                        lForeColor = vbBlack
                    End If
                    ITM.ListSubItems.add(, , Val(rs1!Rate & "")).Tag = ""
                    If Val(rs1!ProcessID & "") = Val(txtProcess.Tag) Then
                        strStatus = "Receiving"
                        iStatus = 2
                    ElseIf Val(rs1!NextToProcessSNo & "") = Val(rs1!ItemSNo & "") Then
                        strStatus = "Issuance"
                        iStatus = 1
                    Else
                        strStatus = "Issuance / Receiving"
                        iStatus = 3
                    End If
                    ITM.ListSubItems.add(, , strStatus).Tag = iStatus
                    ITM.ListSubItems.add(, , rs1!code & "").Tag = rs1!InspectionProcess
                    ITM.ForeColor = lForeColor
                    For i = 1 To ITM.ListSubItems.count
                        ITM.ListSubItems(i).ForeColor = lForeColor
                    Next
                    
                    '------------------------------------------------------------------------------
                    If Val(rs1!NextToProcessSNo & "") = 0 Then
                        lNextToProcessSNo = rs1!ProcessID
                    Else
                        lNextToProcessSNo = Val(rs1!NextToProcessSNo & "")
                    End If
                    rs1.MoveNEXT
                Loop
                
            End If
            .Open "SELECT ItemPic FROM Items WHERE ItemID='" & txtItemCode & "'", con, adOpenForwardOnly, adLockReadOnly
            If Not IsNull(![ItemPic]) Then
                Set Image1.DataSource = rs
                Image1.DataField = "ItemPic"
            End If
            .Close
            
            .Open "SELECT * FROM VItemProcesses WHERE IsExist='" & txtItemCode & "' ORDER BY SNo", con, adOpenForwardOnly, adLockReadOnly
            cmbProcessRW.Clear
            Do Until .EOF
                cmbProcessRW.AddItem !code & " - " & !Description
                cmbProcessRW.List(cmbProcessRW.ListCount - 1, 1) = Val(!ProcessID & "")
                cmbProcessRW.List(cmbProcessRW.ListCount - 1, 2) = !code & ""
                cmbProcessRJ.AddItem !code & " - " & !Description
                cmbProcessRJ.List(cmbProcessRJ.ListCount - 1, 1) = Val(!ProcessID & "")
                cmbProcessRJ.List(cmbProcessRJ.ListCount - 1, 2) = !code & ""
                .MoveNEXT
            Loop
            .Close
            .Open "SELECT * FROM VItemProcesses WHERE IsExist='" & txtItemCode & "' AND ItemSNo<=" & lNextToProcessSNo & " AND InspectionProcess=1 ORDER BY ItemSNo", con, adOpenForwardOnly, adLockReadOnly
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
        
        Call FillRework(Val(txtLotBarcode))
        Call FillRejection(Val(txtLotBarcode))
    End If
    
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Val(txtRate) < 0 Then
            MsgBox "Please Enter Rate.", vbInformation
            Exit Sub
        End If
        'Now Check If this Item is Assigned at this Process...
        Dim iAffect As Integer
        iAffect = 0
        
        If Val(txtRate) > 0 Then
            con.Execute "UPDATE VendAssItems SET Rate=" & Val(txtRate) & " WHERE VendID=" & _
             LV.SelectedItem.ListSubItems(1).Tag & " AND ItemID='" & txtItemCode & "'", iAffect
             
            If iAffect = 0 Then 'Not Assigned
                con.Execute "INSERT INTO VendAssItems(VendID,ProcessID,ItemID,Rate,PlantRate,StampRate,SnaffRate,Unit)" & _
                 "VALUES(" & LV.SelectedItem.ListSubItems(1).Tag & ",0,'" & txtItemCode & _
                 "'," & Val(txtRate) & ",0,0,0,'" & _
                 GetSingleStringValue("Unit", "Items", " WHERE ItemID='" & txtItemCode & "'") & "')"
            End If
        End If
        
        LV.SelectedItem.SubItems(2) = Val(txtRate)
        
        txtRate.Visible = False
        'Call txtLotBarcode_KeyPress(13)
    ElseIf KeyAscii = 27 Then
        txtRate.Visible = False
    End If
    
End Sub

Private Sub txtRate_LostFocus()
    txtRate.Visible = False
End Sub

Private Sub txtReasonRJ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'Save & Add to List
        If cmbProcessRJ.MatchFound = False Then
            MsgBox "Please Select Process.", vbInformation
            Exit Sub
        ElseIf Val(txtRejQty) <= 0 Then
            MsgBox "Please Enter Rej. Qty", vbInformation
            Exit Sub
        ElseIf txtReasonRJ = "" Then
            MsgBox "Please Enter Reason", vbInformation
            Exit Sub
        End If
        con.Execute "INSERT INTO IPORejectionRecord(IPO_RefID,ProcessNo,RejectQty,Reason) VALUES(" & _
         Val(txtLotBarcode) & "," & cmbProcessRJ.List(cmbProcessRJ.ListIndex, 1) & "," & Val(txtRejQty) & ",'" & _
         txtReasonRJ & "')"
         
        Call FillRejection(Val(txtLotBarcode))
    End If

End Sub

Private Sub txtRejQtyRW_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        'Save & Add to List
        If cmbProcessRW.MatchFound = False Then
            MsgBox "Please Select Process.", vbInformation
            Exit Sub
        ElseIf txtRWDone = "" Then
            MsgBox "Please Enter Done By.", vbInformation
            Exit Sub
        ElseIf txtRWReason = "" Then
            MsgBox "Please Enter Reason", vbInformation
            Exit Sub
        End If
        con.Execute "INSERT INTO IPOReworkRecord(IPO_RefID,ProcessNo,ReworkQty,Reason,DoneBy,AccQty,RejQty) VALUES(" & _
         Val(txtLotBarcode) & "," & cmbProcessRW.List(cmbProcessRW.ListIndex, 1) & "," & Val(txtRWQty) & ",'" & _
         txtRWReason & "','" & txtRWDone & "'," & Val(txtAccQtyRW) & "," & Val(txtRejQtyRW) & ")"
         
        Call FillRework(Val(txtLotBarcode))
    End If
    
End Sub

Private Sub FillRework(lIPO As Long)

    On Error GoTo err
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        .Open "SELECT * FROM VIPOReworkRecord WHERE IPO_RefID=" & lIPO & " ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LVRW.ListItems.Clear
        Do Until .EOF
            Set ITM = LVRW.ListItems.add(, !EntryID & "'", !code & " " & !Description)
            'ITM.ListSubItems.add , , Val(!ReWorkQty & "")
            ITM.ListSubItems.add , , !Reason & ""
            ITM.ListSubItems.add , , !RepairType & ""
            ITM.ListSubItems.add(, , Val(!ReWorkQty & "")).Tag = Val(!IPO_Repair_RefID & "")
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub FillRejection(lIPO As Long)

    On Error GoTo err
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        .Open "SELECT * FROM VIPORejectionRecord WHERE IPO_RefID=" & lIPO & " ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LVRJ.ListItems.Clear
        Do Until .EOF
            Set ITM = LVRJ.ListItems.add(, !EntryID & "'", !code & " " & !Description)
            ITM.ListSubItems.add , , Val(!RejectQty & "")
            ITM.ListSubItems.add , , !Reason & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
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

Private Sub txtRWQty_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        'Save & Add to List
        If cmbProcessRW.MatchFound = False Then
            MsgBox "Please Select Process.", vbInformation
            Exit Sub
        ElseIf txtRWReason = "" Then
            MsgBox "Please Enter Reason", vbInformation
            Exit Sub
        ElseIf cmbRepairFamily.MatchFound = False Then
            MsgBox "Please Select Repair Family.", vbInformation
            Exit Sub
        End If
        con.Execute "INSERT INTO IPOReworkRecord(IPO_RefID,ProcessNo,ReworkQty,Reason,DoneBy,AccQty,RejQty,RepairType_RefID,DT) VALUES(" & _
         Val(txtLotBarcode) & "," & cmbProcessRW.List(cmbProcessRW.ListIndex, 1) & "," & Val(txtRWQty) & ",'" & _
         txtRWReason & "','" & txtRWDone & "'," & Val(txtAccQtyRW) & "," & Val(txtRejQtyRW) & "," & cmbRepairFamily.List(cmbRepairFamily.ListIndex, 1) & ",'" & Date & "')"
         
        Call FillRework(Val(txtLotBarcode))
    End If
    
End Sub

Private Sub txtTemperValue_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If cmbProcessNo.List(cmbProcessNo.ListIndex, 1) <> 168 Then
            MsgBox "Not Available for this Process.", vbInformation
            Exit Sub
        End If
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
