VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDailyIncome 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daily Income Sheet"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleMode       =   0  'User
   ScaleWidth      =   8357.006
   WindowState     =   2  'Maximized
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   105
      TabIndex        =   4
      Top             =   -45
      Width           =   8700
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Daily Income Sheet"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   390
         Index           =   0
         Left            =   60
         TabIndex        =   5
         Top             =   120
         Width           =   8475
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Daily Income Sheet"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Index           =   1
         Left            =   60
         TabIndex        =   6
         Top             =   135
         Width           =   8475
      End
   End
   Begin VB.Frame FLV 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   105
      TabIndex        =   0
      Top             =   465
      Width           =   9000
      Begin TabDlg.SSTab SSTab1 
         Height          =   5805
         Left            =   315
         TabIndex        =   7
         Top             =   195
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   10239
         _Version        =   393216
         Tabs            =   4
         Tab             =   3
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "The Heart Test Centre"
         TabPicture(0)   =   "frmDailyIncome.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "SSTab2"
         Tab(0).Control(1)=   "cmdAdd"
         Tab(0).Control(2)=   "txtfrom"
         Tab(0).Control(3)=   "txtto"
         Tab(0).Control(4)=   "Pic"
         Tab(0).Control(5)=   "LV"
         Tab(0).Control(6)=   "DTPicker1"
         Tab(0).Control(7)=   "Label2"
         Tab(0).ControlCount=   8
         TabCaption(1)   =   "CCU"
         TabPicture(1)   =   "frmDailyIncome.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame4"
         Tab(1).Control(1)=   "PicCCU"
         Tab(1).Control(2)=   "txtToCCU"
         Tab(1).Control(3)=   "txtFromCCU"
         Tab(1).Control(4)=   "cmdAddCCU"
         Tab(1).Control(5)=   "LVCCU"
         Tab(1).Control(6)=   "DTPicker2"
         Tab(1).Control(7)=   "Label3"
         Tab(1).ControlCount=   8
         TabCaption(2)   =   "Private Rooms"
         TabPicture(2)   =   "frmDailyIncome.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label4"
         Tab(2).Control(1)=   "DTPicker3"
         Tab(2).Control(2)=   "LVPR"
         Tab(2).Control(3)=   "Frame5"
         Tab(2).Control(4)=   "PicPR"
         Tab(2).Control(5)=   "txtToPR"
         Tab(2).Control(6)=   "txtFromPR"
         Tab(2).Control(7)=   "cmdAddPR"
         Tab(2).ControlCount=   8
         TabCaption(3)   =   "OPD"
         TabPicture(3)   =   "frmDailyIncome.frx":0054
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "Frame7"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "SSTab3"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).ControlCount=   2
         Begin TabDlg.SSTab SSTab3 
            Height          =   3570
            Left            =   90
            TabIndex        =   175
            Top             =   480
            Width           =   8205
            _ExtentX        =   14473
            _ExtentY        =   6297
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "OPD Consualtancy"
            TabPicture(0)   =   "frmDailyIncome.frx":0070
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label5"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "DTPicker4"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "LVOPD"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "cmdAddOPD"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "txtFromOPD"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "txtToOPD"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).ControlCount=   6
            TabCaption(1)   =   "ECG"
            TabPicture(1)   =   "frmDailyIncome.frx":008C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Label6"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "DTPicker5"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "LVOPDECG"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "cmdAddOPDECg"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).Control(4)=   "txtOECGfrom"
            Tab(1).Control(4).Enabled=   0   'False
            Tab(1).Control(5)=   "txtOECGTo"
            Tab(1).Control(5).Enabled=   0   'False
            Tab(1).Control(6)=   "txtsOPD(0)"
            Tab(1).Control(6).Enabled=   0   'False
            Tab(1).ControlCount=   7
            Begin VB.TextBox txtsOPD 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   225
               Index           =   0
               Left            =   -71745
               TabIndex        =   188
               Top             =   2520
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.TextBox txtOECGTo 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   -70785
               TabIndex        =   184
               Top             =   645
               Width           =   1200
            End
            Begin VB.TextBox txtOECGfrom 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   -72015
               TabIndex        =   183
               Top             =   645
               Width           =   1200
            End
            Begin VB.CommandButton cmdAddOPDECg 
               Caption         =   "Add"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   -69525
               Picture         =   "frmDailyIncome.frx":00A8
               Style           =   1  'Graphical
               TabIndex        =   182
               ToolTipText     =   "Add This Entry To Voucher"
               Top             =   390
               UseMaskColor    =   -1  'True
               Width           =   690
            End
            Begin VB.TextBox txtToOPD 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   4215
               TabIndex        =   178
               Top             =   645
               Width           =   1200
            End
            Begin VB.TextBox txtFromOPD 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2985
               TabIndex        =   177
               Top             =   645
               Width           =   1200
            End
            Begin VB.CommandButton cmdAddOPD 
               Caption         =   "Add"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   5475
               Picture         =   "frmDailyIncome.frx":05DA
               Style           =   1  'Graphical
               TabIndex        =   176
               ToolTipText     =   "Add This Entry To Voucher"
               Top             =   390
               UseMaskColor    =   -1  'True
               Width           =   690
            End
            Begin MSComctlLib.ListView LVOPD 
               Height          =   2460
               Left            =   105
               TabIndex        =   179
               Top             =   1005
               Width           =   7995
               _ExtentX        =   14102
               _ExtentY        =   4339
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   15199215
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "SNo."
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Fee"
                  Object.Width           =   2117
               EndProperty
            End
            Begin MSComCtl2.DTPicker DTPicker4 
               Height          =   285
               Left            =   90
               TabIndex        =   180
               Top             =   645
               Width           =   2850
               _ExtentX        =   5027
               _ExtentY        =   503
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "dddd dd-MMM-yyyy"
               Format          =   55902211
               CurrentDate     =   37384
            End
            Begin MSComctlLib.ListView LVOPDECG 
               Height          =   2460
               Left            =   -74895
               TabIndex        =   185
               Top             =   1005
               Width           =   7995
               _ExtentX        =   14102
               _ExtentY        =   4339
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   15199215
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "SNo."
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "ECG"
                  Object.Width           =   2117
               EndProperty
            End
            Begin MSComCtl2.DTPicker DTPicker5 
               Height          =   285
               Left            =   -74895
               TabIndex        =   186
               Top             =   645
               Width           =   2850
               _ExtentX        =   5027
               _ExtentY        =   503
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "dddd dd-MMM-yyyy"
               Format          =   55902211
               CurrentDate     =   37384
            End
            Begin MSForms.Label Label6 
               Height          =   240
               Left            =   -74895
               TabIndex        =   187
               Top             =   390
               Width           =   5325
               BackColor       =   11517387
               Caption         =   " Date                                                       SNo From          SNo To"
               Size            =   "9393;423"
               BorderColor     =   16512
               BorderStyle     =   1
               FontName        =   "Tahoma"
               FontEffects     =   1073741825
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               FontWeight      =   700
            End
            Begin MSForms.Label Label5 
               Height          =   240
               Left            =   105
               TabIndex        =   181
               Top             =   390
               Width           =   5325
               BackColor       =   11517387
               Caption         =   " Date                                                       SNo From          SNo To"
               Size            =   "9393;423"
               BorderColor     =   16512
               BorderStyle     =   1
               FontName        =   "Tahoma"
               FontEffects     =   1073741825
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               FontWeight      =   700
            End
         End
         Begin VB.Frame Frame7 
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
            Height          =   1650
            Left            =   525
            TabIndex        =   160
            Top             =   4050
            Width           =   7635
            Begin VB.Label lblOPDRate 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   1
               Left            =   5415
               TabIndex        =   174
               Top             =   623
               Width           =   375
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Total ECG Amount @         / ECG:"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   49
               Left            =   3915
               TabIndex        =   173
               Top             =   623
               Width           =   2370
            End
            Begin VB.Label lblOPD 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   3
               Left            =   6345
               TabIndex        =   172
               Top             =   600
               Width           =   1005
            End
            Begin VB.Label lblOPD 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   2
               Left            =   6345
               TabIndex        =   170
               Top             =   307
               Width           =   1005
            End
            Begin VB.Label lblOPDRate 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BorderStyle     =   1  'Fixed Single
               Caption         =   "100"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   0
               Left            =   1500
               TabIndex        =   168
               Top             =   630
               Width           =   600
            End
            Begin VB.Label lblOPD 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   5
               Left            =   6345
               TabIndex        =   167
               Top             =   1305
               Width           =   1005
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Net H.C.S. Income (OPD)         :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   48
               Left            =   3990
               TabIndex        =   166
               Top             =   1305
               Width           =   2295
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00808080&
               Index           =   9
               X1              =   3975
               X2              =   7300
               Y1              =   1245
               Y2              =   1245
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00FFFFFF&
               Index           =   8
               X1              =   3975
               X2              =   7300
               Y1              =   1260
               Y2              =   1260
            End
            Begin VB.Label lblOPD 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   4
               Left            =   6345
               TabIndex        =   165
               Top             =   960
               Width           =   1005
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Discount                                     :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   47
               Left            =   3945
               TabIndex        =   164
               Top             =   960
               Width           =   2340
            End
            Begin VB.Label lblOPD 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   225
               Index           =   1
               Left            =   2610
               TabIndex        =   163
               Top             =   608
               Width           =   1005
            End
            Begin VB.Label lblOPD 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   0
               Left            =   2610
               TabIndex        =   162
               Top             =   307
               Width           =   1005
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Total Charges  @              /OPD :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   52
               Left            =   225
               TabIndex        =   169
               Top             =   630
               Width           =   2355
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Total Number Of OPD's              :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   46
               Left            =   240
               TabIndex        =   161
               Top             =   330
               Width           =   2340
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Total Number Of ECG's               :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   51
               Left            =   3915
               TabIndex        =   171
               Top             =   330
               Width           =   2370
            End
         End
         Begin TabDlg.SSTab SSTab2 
            Height          =   1695
            Left            =   -74865
            TabIndex        =   83
            Top             =   4050
            Width           =   7950
            _ExtentX        =   14023
            _ExtentY        =   2990
            _Version        =   393216
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   441
            TabCaption(0)   =   "ECG"
            TabPicture(0)   =   "frmDailyIncome.frx":0B0C
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Frame1"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "ECHO"
            TabPicture(1)   =   "frmDailyIncome.frx":0B28
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame3"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "ETT"
            TabPicture(2)   =   "frmDailyIncome.frx":0B44
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Frame2"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "CardiacT"
            TabPicture(3)   =   "frmDailyIncome.frx":0B60
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "Frame6"
            Tab(3).ControlCount=   1
            Begin VB.Frame Frame6 
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
               Height          =   1300
               Left            =   -74850
               TabIndex        =   141
               Top             =   270
               Width           =   7635
               Begin VB.Label lblCardiacT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   6
                  Left            =   6405
                  TabIndex        =   159
                  Top             =   945
                  Width           =   1005
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Net H.C.S. Income (Cardiac T) :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   45
                  Left            =   4035
                  TabIndex        =   158
                  Top             =   945
                  Width           =   2295
               End
               Begin VB.Line Line1 
                  BorderColor     =   &H00808080&
                  Index           =   1
                  X1              =   4035
                  X2              =   7360
                  Y1              =   900
                  Y2              =   900
               End
               Begin VB.Line Line1 
                  BorderColor     =   &H00FFFFFF&
                  Index           =   0
                  X1              =   4035
                  X2              =   7360
                  Y1              =   915
                  Y2              =   915
               End
               Begin VB.Label lblCardiacT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   5
                  Left            =   6405
                  TabIndex        =   157
                  Top             =   645
                  Width           =   1005
               End
               Begin VB.Label lblCardiacT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   4
                  Left            =   6405
                  TabIndex        =   155
                  Top             =   420
                  Width           =   1005
               End
               Begin VB.Label lblCardiacT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   225
                  Index           =   1
                  Left            =   2685
                  TabIndex        =   154
                  Top             =   405
                  Width           =   1005
               End
               Begin VB.Label lblCardiacT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   225
                  Index           =   0
                  Left            =   2685
                  TabIndex        =   152
                  Top             =   195
                  Width           =   1005
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Total Number Of Cardiac T's       :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   42
                  Left            =   285
                  TabIndex        =   151
                  Top             =   225
                  Width           =   2385
               End
               Begin VB.Label lblCardiacT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   3
                  Left            =   6405
                  TabIndex        =   150
                  Top             =   195
                  Width           =   1005
               End
               Begin VB.Label lblCardiacT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   2
                  Left            =   2685
                  TabIndex        =   148
                  Top             =   615
                  Visible         =   0   'False
                  Width           =   1005
               End
               Begin VB.Label lblCardiacTRate 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "100"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   0
                  Left            =   1545
                  TabIndex        =   146
                  Top             =   435
                  Width           =   600
               End
               Begin VB.Label lblCardiacTRate 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "100"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   1
                  Left            =   1560
                  TabIndex        =   145
                  Top             =   645
                  Visible         =   0   'False
                  Width           =   600
               End
               Begin VB.Label lblCardiacTRate 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "100"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   2
                  Left            =   5520
                  TabIndex        =   144
                  Top             =   210
                  Width           =   525
               End
               Begin VB.Label lblCardiacTRate 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "100"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   3
                  Left            =   5520
                  TabIndex        =   142
                  Top             =   420
                  Width           =   525
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Total Charges  @              /Cd.T :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   43
                  Left            =   285
                  TabIndex        =   153
                  Top             =   420
                  Width           =   2385
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Total Cost        @              /Cd.T :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   40
                  Left            =   285
                  TabIndex        =   147
                  Top             =   630
                  Visible         =   0   'False
                  Width           =   2385
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Discount                                    :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   44
                  Left            =   4050
                  TabIndex        =   156
                  Top             =   660
                  Width           =   2295
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Consaltant Chgs. @             % :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   41
                  Left            =   4050
                  TabIndex        =   149
                  Top             =   195
                  Width           =   2295
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "H.C.S. Income     @             % :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   39
                  Left            =   4065
                  TabIndex        =   143
                  Top             =   420
                  Width           =   2280
               End
            End
            Begin VB.Frame Frame2 
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
               Height          =   1300
               Left            =   -74850
               TabIndex        =   122
               Top             =   270
               Width           =   7635
               Begin VB.Label lblETT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   6
                  Left            =   6345
                  TabIndex        =   140
                  Top             =   945
                  Width           =   1005
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Net H.C.S. Income (ETT)          :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   9
                  Left            =   3990
                  TabIndex        =   139
                  Top             =   945
                  Width           =   2295
               End
               Begin VB.Line Line1 
                  BorderColor     =   &H00808080&
                  Index           =   4
                  X1              =   3975
                  X2              =   7300
                  Y1              =   900
                  Y2              =   900
               End
               Begin VB.Line Line1 
                  BorderColor     =   &H00FFFFFF&
                  Index           =   5
                  X1              =   3975
                  X2              =   7300
                  Y1              =   915
                  Y2              =   915
               End
               Begin VB.Label lblETT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   5
                  Left            =   6345
                  TabIndex        =   138
                  Top             =   645
                  Width           =   1005
               End
               Begin VB.Label lblETT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   4
                  Left            =   6345
                  TabIndex        =   136
                  Top             =   420
                  Width           =   1005
               End
               Begin VB.Label lblETT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   225
                  Index           =   1
                  Left            =   2625
                  TabIndex        =   135
                  Top             =   405
                  Width           =   1005
               End
               Begin VB.Label lblETT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   225
                  Index           =   0
                  Left            =   2625
                  TabIndex        =   133
                  Top             =   195
                  Width           =   1005
               End
               Begin VB.Label lblETT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   3
                  Left            =   6345
                  TabIndex        =   131
                  Top             =   195
                  Width           =   1005
               End
               Begin VB.Label lblETT 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   2
                  Left            =   2625
                  TabIndex        =   129
                  Top             =   615
                  Visible         =   0   'False
                  Width           =   1005
               End
               Begin VB.Label lblETTRate 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "100"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   0
                  Left            =   1515
                  TabIndex        =   127
                  Top             =   435
                  Width           =   600
               End
               Begin VB.Label lblETTRate 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "100"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   1
                  Left            =   1530
                  TabIndex        =   126
                  Top             =   645
                  Visible         =   0   'False
                  Width           =   600
               End
               Begin VB.Label lblETTRate 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "100"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   2
                  Left            =   5460
                  TabIndex        =   125
                  Top             =   210
                  Width           =   525
               End
               Begin VB.Label lblETTRate 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "100"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   3
                  Left            =   5460
                  TabIndex        =   123
                  Top             =   420
                  Width           =   525
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Total Charges  @              /ETT :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   12
                  Left            =   255
                  TabIndex        =   134
                  Top             =   420
                  Width           =   2310
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Total Number Of ETT's              :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   13
                  Left            =   270
                  TabIndex        =   132
                  Top             =   225
                  Width           =   2295
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Total Cost        @              /ETT :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   15
                  Left            =   255
                  TabIndex        =   128
                  Top             =   630
                  Visible         =   0   'False
                  Width           =   2310
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Discount                                    :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   10
                  Left            =   3990
                  TabIndex        =   137
                  Top             =   660
                  Width           =   2295
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Consaltant Chgs. @             % :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   14
                  Left            =   3990
                  TabIndex        =   130
                  Top             =   195
                  Width           =   2295
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "H.C.S. Income     @             % :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   11
                  Left            =   4005
                  TabIndex        =   124
                  Top             =   420
                  Width           =   2280
               End
            End
            Begin VB.Frame Frame3 
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
               Height          =   1300
               Left            =   -74850
               TabIndex        =   103
               Top             =   270
               Width           =   7635
               Begin VB.Label lblECHO 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   2
                  Left            =   2655
                  TabIndex        =   120
                  Top             =   645
                  Visible         =   0   'False
                  Width           =   1005
               End
               Begin VB.Label lblECHO 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   3
                  Left            =   6390
                  TabIndex        =   118
                  Top             =   210
                  Width           =   1005
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Total Number Of ECHO's             :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   18
                  Left            =   240
                  TabIndex        =   117
                  Top             =   210
                  Width           =   2400
               End
               Begin VB.Label lblECHO 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   225
                  Index           =   0
                  Left            =   2655
                  TabIndex        =   116
                  Top             =   225
                  Width           =   1005
               End
               Begin VB.Label lblECHO 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   225
                  Index           =   1
                  Left            =   2655
                  TabIndex        =   114
                  Top             =   435
                  Width           =   1005
               End
               Begin VB.Label lblECHO 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   4
                  Left            =   6390
                  TabIndex        =   113
                  Top             =   435
                  Width           =   1005
               End
               Begin VB.Label lblECHO 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   5
                  Left            =   6390
                  TabIndex        =   111
                  Top             =   660
                  Width           =   1005
               End
               Begin VB.Line Line1 
                  BorderColor     =   &H00FFFFFF&
                  Index           =   10
                  X1              =   4020
                  X2              =   7345
                  Y1              =   930
                  Y2              =   930
               End
               Begin VB.Line Line1 
                  BorderColor     =   &H00808080&
                  Index           =   11
                  X1              =   4020
                  X2              =   7345
                  Y1              =   915
                  Y2              =   915
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Net H.C.S. Income (ECHO)       :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   22
                  Left            =   4050
                  TabIndex        =   110
                  Top             =   960
                  Width           =   2310
               End
               Begin VB.Label lblECHO 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   6
                  Left            =   6390
                  TabIndex        =   109
                  Top             =   960
                  Width           =   1005
               End
               Begin VB.Label lblECHORate 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "100"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   0
                  Left            =   1485
                  TabIndex        =   108
                  Top             =   435
                  Width           =   600
               End
               Begin VB.Label lblECHORate 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "100"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   1
                  Left            =   1485
                  TabIndex        =   107
                  Top             =   645
                  Visible         =   0   'False
                  Width           =   600
               End
               Begin VB.Label lblECHORate 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "100"
                  ForeColor       =   &H00800000&
                  Height          =   225
                  Index           =   2
                  Left            =   5550
                  TabIndex        =   106
                  Top             =   225
                  Width           =   495
               End
               Begin VB.Label lblECHORate 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "100"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   3
                  Left            =   5520
                  TabIndex        =   104
                  Top             =   480
                  Width           =   525
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Total Cost       @              /ECHO:"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   16
                  Left            =   255
                  TabIndex        =   121
                  Top             =   645
                  Visible         =   0   'False
                  Width           =   2370
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Total Charges @              /ECHO:"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   19
                  Left            =   255
                  TabIndex        =   115
                  Top             =   435
                  Width           =   2370
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Consaltant Chgs. @             % :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   17
                  Left            =   4065
                  TabIndex        =   119
                  Top             =   210
                  Width           =   2295
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Discount                                    :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   21
                  Left            =   4065
                  TabIndex        =   112
                  Top             =   675
                  Width           =   2295
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "H.C.S. Income     @             % :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   20
                  Left            =   4080
                  TabIndex        =   105
                  Top             =   450
                  Width           =   2280
               End
            End
            Begin VB.Frame Frame1 
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
               Height          =   1300
               Left            =   150
               TabIndex        =   84
               Top             =   270
               Width           =   7635
               Begin VB.Label lblECG 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   2
                  Left            =   2610
                  TabIndex        =   101
                  Top             =   630
                  Width           =   1005
               End
               Begin VB.Label lblECG 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   3
                  Left            =   6345
                  TabIndex        =   99
                  Top             =   180
                  Width           =   1005
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Total Number Of ECG's              :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   7
                  Left            =   240
                  TabIndex        =   98
                  Top             =   210
                  Width           =   2325
               End
               Begin VB.Label lblECG 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   0
                  Left            =   2610
                  TabIndex        =   97
                  Top             =   195
                  Width           =   1005
               End
               Begin VB.Label lblECG 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   225
                  Index           =   1
                  Left            =   2610
                  TabIndex        =   95
                  Top             =   420
                  Width           =   1005
               End
               Begin VB.Label lblECG 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   4
                  Left            =   6345
                  TabIndex        =   93
                  Top             =   405
                  Width           =   1005
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Discount                                    :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   6
                  Left            =   4005
                  TabIndex        =   92
                  Top             =   660
                  Width           =   2295
               End
               Begin VB.Label lblECG 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   5
                  Left            =   6345
                  TabIndex        =   91
                  Top             =   630
                  Width           =   1005
               End
               Begin VB.Line Line1 
                  BorderColor     =   &H00FFFFFF&
                  Index           =   2
                  X1              =   3975
                  X2              =   7300
                  Y1              =   915
                  Y2              =   915
               End
               Begin VB.Line Line1 
                  BorderColor     =   &H00808080&
                  Index           =   3
                  X1              =   3975
                  X2              =   7300
                  Y1              =   900
                  Y2              =   900
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Net H.C.S. Income (ECG)          :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   8
                  Left            =   3990
                  TabIndex        =   90
                  Top             =   960
                  Width           =   2325
               End
               Begin VB.Label lblECG 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   6
                  Left            =   6345
                  TabIndex        =   89
                  Top             =   960
                  Width           =   1005
               End
               Begin VB.Label lblECGRate 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "100"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   0
                  Left            =   1500
                  TabIndex        =   88
                  Top             =   435
                  Width           =   600
               End
               Begin VB.Label lblECGRate 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "100"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   1
                  Left            =   1500
                  TabIndex        =   87
                  Top             =   660
                  Width           =   600
               End
               Begin VB.Label lblECGRate 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "100"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   2
                  Left            =   5460
                  TabIndex        =   86
                  Top             =   210
                  Width           =   510
               End
               Begin VB.Label lblECGRate 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "100"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   3
                  Left            =   5475
                  TabIndex        =   85
                  Top             =   435
                  Width           =   510
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Consaltant Chgs. @             % :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   4
                  Left            =   3990
                  TabIndex        =   100
                  Top             =   210
                  Width           =   2295
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "H.C.S. Income     @             % :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   5
                  Left            =   4005
                  TabIndex        =   94
                  Top             =   435
                  Width           =   2280
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Total Cost        @              /ECG :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   3
                  Left            =   225
                  TabIndex        =   102
                  Top             =   660
                  Width           =   2340
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Total Charges  @              /ECG :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   2
                  Left            =   225
                  TabIndex        =   96
                  Top             =   435
                  Width           =   2340
               End
            End
         End
         Begin VB.CommandButton cmdAddPR 
            Caption         =   "Add"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   -69510
            Picture         =   "frmDailyIncome.frx":0B7C
            Style           =   1  'Graphical
            TabIndex        =   75
            ToolTipText     =   "Add This Entry To Voucher"
            Top             =   420
            UseMaskColor    =   -1  'True
            Width           =   690
         End
         Begin VB.TextBox txtFromPR 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -72000
            TabIndex        =   74
            Top             =   690
            Width           =   1200
         End
         Begin VB.TextBox txtToPR 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -70770
            TabIndex        =   68
            Top             =   690
            Width           =   1200
         End
         Begin VB.PictureBox PicPR 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   -74655
            ScaleHeight     =   195
            ScaleWidth      =   5235
            TabIndex        =   67
            Top             =   1620
            Visible         =   0   'False
            Width           =   5265
            Begin VB.TextBox txtsPR 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   4
               Left            =   4110
               TabIndex        =   73
               Top             =   0
               Width           =   1005
            End
            Begin VB.TextBox txtsPR 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   3
               Left            =   3075
               TabIndex        =   72
               Top             =   0
               Width           =   1005
            End
            Begin VB.TextBox txtsPR 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   0
               Left            =   30
               TabIndex        =   69
               Top             =   -30
               Width           =   930
            End
            Begin VB.TextBox txtsPR 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   1
               Left            =   990
               TabIndex        =   70
               Top             =   -30
               Width           =   1005
            End
            Begin VB.TextBox txtsPR 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   2
               Left            =   2025
               TabIndex        =   71
               Top             =   -30
               Width           =   1005
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   " Private Rooms "
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
            Height          =   1545
            Left            =   -74640
            TabIndex        =   48
            Top             =   4050
            Width           =   7440
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Air Conditioner Using Charges  :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   38
               Left            =   3870
               TabIndex        =   82
               Top             =   750
               Width           =   2295
            End
            Begin VB.Label lblPR 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   7
               Left            =   6210
               TabIndex        =   81
               Top             =   750
               Width           =   1005
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00808080&
               Index           =   21
               X1              =   3840
               X2              =   7165
               Y1              =   705
               Y2              =   705
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00FFFFFF&
               Index           =   20
               X1              =   3840
               X2              =   7165
               Y1              =   720
               Y2              =   720
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Heater Using Charges               :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   37
               Left            =   3855
               TabIndex        =   80
               Top             =   975
               Width           =   2310
            End
            Begin VB.Label lblPR 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   8
               Left            =   6210
               TabIndex        =   79
               Top             =   975
               Width           =   1005
            End
            Begin VB.Label lblPRRate 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BorderStyle     =   1  'Fixed Single
               Caption         =   "100"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   3
               Left            =   1695
               TabIndex        =   63
               Top             =   1230
               Width           =   510
            End
            Begin VB.Label lblPRRate 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BorderStyle     =   1  'Fixed Single
               Caption         =   "100"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   2
               Left            =   1695
               TabIndex        =   61
               Top             =   990
               Width           =   510
            End
            Begin VB.Label lblPRRate 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BorderStyle     =   1  'Fixed Single
               Caption         =   "100"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   0
               Left            =   1470
               TabIndex        =   60
               Top             =   435
               Width           =   600
            End
            Begin VB.Label lblPR 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   225
               Index           =   6
               Left            =   6210
               TabIndex        =   59
               Top             =   465
               Width           =   1005
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Net H.C.S. Income                    :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   32
               Left            =   3855
               TabIndex        =   58
               Top             =   465
               Width           =   2310
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00808080&
               Index           =   19
               X1              =   3840
               X2              =   7165
               Y1              =   420
               Y2              =   420
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00FFFFFF&
               Index           =   18
               X1              =   3840
               X2              =   7165
               Y1              =   435
               Y2              =   435
            End
            Begin VB.Label lblPR 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   5
               Left            =   6210
               TabIndex        =   57
               Top             =   165
               Width           =   1005
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Discount                                    :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   31
               Left            =   3870
               TabIndex        =   56
               Top             =   195
               Width           =   2295
            End
            Begin VB.Label lblPR 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   4
               Left            =   2580
               TabIndex        =   55
               Top             =   1200
               Width           =   1005
            End
            Begin VB.Label lblPR 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   1
               Left            =   2580
               TabIndex        =   54
               Top             =   435
               Width           =   1005
            End
            Begin VB.Label lblPR 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   2580
               TabIndex        =   53
               Top             =   195
               Width           =   1005
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Total Number Of Room/Days     :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   30
               Left            =   210
               TabIndex        =   52
               Top             =   210
               Width           =   2340
            End
            Begin VB.Label lblPR 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   3
               Left            =   2580
               TabIndex        =   51
               Top             =   975
               Width           =   1005
            End
            Begin VB.Label lblPR 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   2
               Left            =   2580
               TabIndex        =   50
               Top             =   750
               Width           =   1005
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00808080&
               Index           =   17
               X1              =   210
               X2              =   3535
               Y1              =   690
               Y2              =   690
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00FFFFFF&
               Index           =   16
               X1              =   210
               X2              =   3535
               Y1              =   705
               Y2              =   705
            End
            Begin VB.Label lblPRRate 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BorderStyle     =   1  'Fixed Single
               Caption         =   "100"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   1
               Left            =   1695
               TabIndex        =   49
               Top             =   765
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Consaltant Chgs. @             % :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   36
               Left            =   240
               TabIndex        =   66
               Top             =   750
               Width           =   2295
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "H.C.S. Income     @             % :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   34
               Left            =   255
               TabIndex        =   64
               Top             =   1200
               Width           =   2280
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "A.K.C         Chgs. @             % :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   33
               Left            =   240
               TabIndex        =   62
               Top             =   975
               Width           =   2295
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Total Charges  @              /Day :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   35
               Left            =   210
               TabIndex        =   65
               Top             =   435
               Width           =   2325
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   " CCU "
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
            Height          =   1380
            Left            =   -74565
            TabIndex        =   29
            Top             =   4125
            Width           =   7065
            Begin VB.Line Line1 
               BorderColor     =   &H00808080&
               Index           =   7
               X1              =   3540
               X2              =   6865
               Y1              =   675
               Y2              =   675
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00FFFFFF&
               Index           =   6
               X1              =   3540
               X2              =   6865
               Y1              =   690
               Y2              =   690
            End
            Begin VB.Label lblCCURate 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BorderStyle     =   1  'Fixed Single
               Caption         =   "100"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   1
               Left            =   1560
               TabIndex        =   46
               Top             =   743
               Width           =   510
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00FFFFFF&
               Index           =   15
               X1              =   75
               X2              =   3400
               Y1              =   675
               Y2              =   675
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00808080&
               Index           =   14
               X1              =   75
               X2              =   3400
               Y1              =   660
               Y2              =   660
            End
            Begin VB.Label lblCCU 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   2
               Left            =   2445
               TabIndex        =   45
               Top             =   720
               Width           =   1005
            End
            Begin VB.Label lblCCU 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   3
               Left            =   5880
               TabIndex        =   43
               Top             =   180
               Width           =   1005
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Total Number Of CCU Days       :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   27
               Left            =   75
               TabIndex        =   42
               Top             =   195
               Width           =   2325
            End
            Begin VB.Label lblCCU 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   225
               Index           =   0
               Left            =   2445
               TabIndex        =   41
               Top             =   195
               Width           =   1005
            End
            Begin VB.Label lblCCU 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   225
               Index           =   1
               Left            =   2445
               TabIndex        =   39
               Top             =   405
               Width           =   1005
            End
            Begin VB.Label lblCCU 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   4
               Left            =   5880
               TabIndex        =   37
               Top             =   405
               Width           =   1005
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Discount                                    :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   24
               Left            =   3540
               TabIndex        =   36
               Top             =   743
               Width           =   2295
            End
            Begin VB.Label lblCCU 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   5
               Left            =   5880
               TabIndex        =   35
               Top             =   720
               Width           =   1005
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00FFFFFF&
               Index           =   13
               X1              =   3510
               X2              =   6835
               Y1              =   990
               Y2              =   990
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00808080&
               Index           =   12
               X1              =   3510
               X2              =   6835
               Y1              =   975
               Y2              =   975
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Net H.C.S. Income                    :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   23
               Left            =   3525
               TabIndex        =   34
               Top             =   1020
               Width           =   2310
            End
            Begin VB.Label lblCCU 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   6
               Left            =   5880
               TabIndex        =   33
               Top             =   1020
               Width           =   1005
            End
            Begin VB.Label lblCCURate 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BorderStyle     =   1  'Fixed Single
               Caption         =   "100"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   0
               Left            =   1335
               TabIndex        =   32
               Top             =   420
               Width           =   600
            End
            Begin VB.Label lblCCURate 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BorderStyle     =   1  'Fixed Single
               Caption         =   "100"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   2
               Left            =   4995
               TabIndex        =   31
               Top             =   195
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Alkhidmat Chgs     @            % :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   28
               Left            =   3555
               TabIndex        =   44
               Top             =   180
               Width           =   2280
            End
            Begin VB.Label lblCCURate 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BorderStyle     =   1  'Fixed Single
               Caption         =   "100"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   3
               Left            =   4995
               TabIndex        =   30
               Top             =   435
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "H.C.S. Income     @             % :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   25
               Left            =   3555
               TabIndex        =   38
               Top             =   405
               Width           =   2280
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Total Charges  @              /Day :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   26
               Left            =   75
               TabIndex        =   40
               Top             =   420
               Width           =   2325
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Consaltant Chgs. @             % :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   29
               Left            =   105
               TabIndex        =   47
               Top             =   743
               Width           =   2295
            End
         End
         Begin VB.PictureBox PicCCU 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   -74715
            ScaleHeight     =   195
            ScaleWidth      =   4050
            TabIndex        =   25
            Top             =   1635
            Visible         =   0   'False
            Width           =   4080
            Begin VB.TextBox txtsCCU 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   2
               Left            =   2580
               TabIndex        =   28
               Top             =   -30
               Width           =   1275
            End
            Begin VB.TextBox txtsCCU 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   1
               Left            =   1275
               TabIndex        =   27
               Top             =   -30
               Width           =   1275
            End
            Begin VB.TextBox txtsCCU 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   0
               Left            =   30
               TabIndex        =   26
               Top             =   -30
               Width           =   1200
            End
         End
         Begin VB.TextBox txtToCCU 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -70770
            TabIndex        =   22
            Top             =   675
            Width           =   1200
         End
         Begin VB.TextBox txtFromCCU 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -72000
            TabIndex        =   21
            Top             =   675
            Width           =   1200
         End
         Begin VB.CommandButton cmdAddCCU 
            Caption         =   "Add"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   -69510
            Picture         =   "frmDailyIncome.frx":10AE
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Add This Entry To Voucher"
            Top             =   420
            UseMaskColor    =   -1  'True
            Width           =   690
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   -69480
            Picture         =   "frmDailyIncome.frx":15E0
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Add This Entry To Voucher"
            Top             =   435
            UseMaskColor    =   -1  'True
            Width           =   690
         End
         Begin VB.TextBox txtfrom 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   -71970
            TabIndex        =   15
            Top             =   690
            Width           =   1200
         End
         Begin VB.TextBox txtto 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   -70725
            TabIndex        =   14
            Top             =   690
            Width           =   1200
         End
         Begin VB.PictureBox Pic 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   -73470
            ScaleHeight     =   195
            ScaleWidth      =   5175
            TabIndex        =   8
            Top             =   1605
            Visible         =   0   'False
            Width           =   5205
            Begin VB.TextBox Txts 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   0
               Left            =   30
               TabIndex        =   9
               Top             =   -30
               Width           =   1200
            End
            Begin VB.TextBox Txts 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   1
               Left            =   1275
               TabIndex        =   10
               Top             =   -30
               Width           =   1275
            End
            Begin VB.TextBox Txts 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   3
               Left            =   3825
               TabIndex        =   12
               Top             =   -15
               Width           =   1275
            End
            Begin VB.TextBox Txts 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   2
               Left            =   2595
               TabIndex        =   11
               Top             =   -15
               Width           =   1200
            End
         End
         Begin MSComctlLib.ListView LV 
            Height          =   2910
            Left            =   -74865
            TabIndex        =   13
            Top             =   1005
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   5133
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   15199215
            Appearance      =   1
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
               Text            =   "S.No."
               Object.Width           =   2081
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   1
               Text            =   "ECG"
               Object.Width           =   2081
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "ECHO"
               Object.Width           =   2081
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "ETT"
               Object.Width           =   2081
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Cardiac T"
               Object.Width           =   2081
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Total Amt"
               Object.Width           =   2081
            EndProperty
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   300
            Left            =   -74850
            TabIndex        =   16
            Top             =   690
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dddd dd-MMM-yyyy"
            Format          =   55902211
            CurrentDate     =   37384
         End
         Begin MSComctlLib.ListView LVCCU 
            Height          =   2910
            Left            =   -74880
            TabIndex        =   19
            Top             =   1035
            Width           =   7995
            _ExtentX        =   14102
            _ExtentY        =   5133
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   15199215
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "SNo."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Bed No."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Days"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Amount"
               Object.Width           =   2117
            EndProperty
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   285
            Left            =   -74880
            TabIndex        =   23
            Top             =   675
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dddd dd-MMM-yyyy"
            Format          =   55902211
            CurrentDate     =   37384
         End
         Begin MSComctlLib.ListView LVPR 
            Height          =   2910
            Left            =   -74880
            TabIndex        =   76
            Top             =   1005
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   5133
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   15199215
            Appearance      =   1
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
               Text            =   "SNo."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Room No."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Days"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Amount"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Heater"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "A.C."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Net Amt."
               Object.Width           =   1764
            EndProperty
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   285
            Left            =   -74880
            TabIndex        =   77
            Top             =   690
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dddd dd-MMM-yyyy"
            Format          =   55902211
            CurrentDate     =   37384
         End
         Begin MSForms.Label Label4 
            Height          =   240
            Left            =   -74880
            TabIndex        =   78
            Top             =   420
            Width           =   5325
            BackColor       =   11517387
            Caption         =   " Date                                                       SNo From          SNo To"
            Size            =   "9393;423"
            BorderColor     =   16512
            BorderStyle     =   1
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.Label Label3 
            Height          =   240
            Left            =   -74880
            TabIndex        =   24
            Top             =   420
            Width           =   5325
            BackColor       =   11517387
            Caption         =   " Date                                                       SNo From          SNo To"
            Size            =   "9393;423"
            BorderColor     =   16512
            BorderStyle     =   1
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.Label Label2 
            Height          =   240
            Left            =   -74850
            TabIndex        =   17
            Top             =   420
            Width           =   5325
            BackColor       =   11517387
            Caption         =   " Date                                                       SNo From          SNo To"
            Size            =   "9393;423"
            BorderColor     =   16512
            BorderStyle     =   1
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
      End
      Begin Crystal.CrystalReport cr1 
         Left            =   90
         Top             =   420
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin MSForms.CommandButton cmdOk 
         Height          =   360
         Left            =   5535
         TabIndex        =   3
         Top             =   6120
         Width           =   1620
         Caption         =   " Save & Close"
         PicturePosition =   327683
         Size            =   "2857;635"
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
         Height          =   360
         Left            =   3870
         TabIndex        =   2
         Top             =   6120
         Width           =   1620
         Caption         =   "Save & New  "
         PicturePosition =   327683
         Size            =   "2857;635"
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
         Height          =   360
         Left            =   7215
         TabIndex        =   1
         Top             =   6120
         Width           =   1620
         Caption         =   "Cancel        "
         PicturePosition =   327683
         Size            =   "2857;635"
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9225
      Top             =   795
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyIncome.frx":1B12
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyIncome.frx":1F64
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDailyIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbdept_matched()
    Call FillList
End Sub

Private Function Saved() As Boolean
    On Error GoTo err
    
    Dim Sql As String
    con.BeginTrans
    con.Execute "Delete From DailyIncomeSheet Where DT='" & Format(DTPicker1, "dd-MMM-yyyy") & "'"
    For i = 1 To LV.ListItems.Count
        With LV.ListItems(i)
            If .SubItems(3) <> "" Then
                Sql = "Insert into DailyIncomeSheet (SNo,DT,ECG," & _
                 "ETT,ECHO,CardiacT) " & _
                 "Values('" & .Text & "','" & Format(DTPicker1, "dd-MMM-yyyy") & "'," & _
                 IIf(Val(.SubItems(1)) = 0, "NULL", Val(.SubItems(1))) & _
                 "," & IIf(.SubItems(2) = "- - -", "NULL", Val(.SubItems(2))) & _
                 "," & IIf(.SubItems(3) = "- - -", "NULL", Val(.SubItems(3))) & _
                 "," & IIf(.SubItems(4) = "- - -", "NULL", Val(.SubItems(4))) & ")"
                con.Execute Sql, a
            End If
        End With
    
    Next i
    
    Dim Desc As String
    Desc = lblECG(0) & " ECGs Sno=" & txtfrom & " TO " & txtto
    
    Call AddTHTCVoucher(DTPicker1, "ECG", Val(lblECG(1)) - Val(lblECG(2)), _
    Val(lblECG(2)), Val(lblECG(3)), Val(lblECG(1)) - Val(lblECG(5)), Val(lblECG(5)), Desc)
    
    Desc = lblECHO(0) & " ECHOs Sno=" & txtfrom & " TO " & txtto
    
    Call AddTHTCVoucher(DTPicker1, "ECHO", Val(lblECHO(1)) - Val(lblECHO(2)), _
    Val(lblECHO(2)), Val(lblECHO(3)), Val(lblECHO(1)) - Val(lblECHO(5)), Val(lblECHO(5)), Desc)
    
    Desc = lblETT(0) & " ETTs Sno=" & txtfrom & " TO " & txtto
    
    Call AddTHTCVoucher(DTPicker1, "ETT", Val(lblETT(1)) - Val(lblETT(2)), _
    Val(lblETT(2)), Val(lblETT(3)), Val(lblETT(1)) - Val(lblETT(5)), Val(lblETT(5)), Desc)
    
    Desc = lblCardiacT(0) & " CardiacTs Sno=" & txtfrom & " TO " & txtto
    
    Call AddTHTCVoucher(DTPicker1, "CardiacT", Val(lblCardiacT(1)) - Val(lblCardiacT(2)), _
    Val(lblCardiacT(2)), Val(lblCardiacT(3)), Val(lblCardiacT(1)) - Val(lblCardiacT(5)), Val(lblCardiacT(5)), Desc)
    
    
    
    ''''''''''''' CCU Saving ''''''''''''
    
    con.Execute "Delete From CCUDailyIncome Where DT='" & Format(DTPicker2, "dd-MMM-yyyy") & "'"
    For i = 1 To LVCCU.ListItems.Count
        With LVCCU.ListItems(i)
            If .SubItems(3) <> "" Then
                Sql = "Insert into CCUDailyIncome(SNo,DT,BedNo,Days,Amt) " & _
                 "Values('" & .Text & "','" & Format(DTPicker2, "dd-MMM-yyyy") & "'," & _
                 .SubItems(1) & "," & Val(.SubItems(2)) & _
                 "," & Val(.SubItems(3)) & ")"
                con.Execute Sql, a
            End If
        End With
    
    
    Next i
    
    Desc = lblCCU(0) & " CCUs Sno=" & txtFromCCU & " TO " & txtToCCU
    
    Call AddCCUVoucher(DTPicker2, Val(lblCCU(1)), Val(lblCCU(2)), Val(lblCCU(3)), _
    Val(lblCCU(1)) - Val(lblCCU(5)), Val(lblCCU(5)), Desc)
    
    
    ''''''''''''' Private Room Saving ''''''''''''
    
    con.Execute "Delete From PRDailyIncome Where DT='" & Format(DTPicker3, "dd-MMM-yyyy") & "'"
    For i = 1 To LVPR.ListItems.Count
        With LVPR.ListItems(i)
            If .SubItems(3) <> "" Then
                Sql = "Insert into PRDailyIncome(SNo,DT,RoomNo,Days,Amt,Heater,AC) " & _
                 "Values('" & .Text & "','" & Format(DTPicker3, "dd-MMM-yyyy") & "'," & _
                 .SubItems(1) & "," & Val(.SubItems(2)) & _
                 "," & Val(.SubItems(3)) & "," & Val(.SubItems(4)) & "," & _
                 Val(.SubItems(5)) & ")"
                con.Execute Sql, a
            End If
        End With
    
    Next i
    
    Desc = lblPR(0) & " Private Room Days Sno=" & txtFromPR & " TO " & txtToPR
    
    Call AddPRVoucher(DTPicker3, Val(lblPR(1)), Val(lblPR(2)), Val(lblPR(3)), _
    Val(lblPR(1)) - Val(lblPR(5)), Val(lblPR(5)), Val(lblPR(6)), Val(lblPR(7)), Desc)
    
    
    
    
    ''''''''''''' OPD Saving ''''''''''''
    
    con.Execute "Delete From OPDDailyIncome Where DT='" & Format(DTPicker4, "dd-MMM-yyyy") & "'"
    
    For i = 1 To LVOPD.ListItems.Count
        With LVOPD.ListItems(i)
            If .SubItems(1) <> "" Then
                Sql = "Insert into OPDDailyIncome(SNo,DT,Amt) " & _
                 "Values('" & .Text & "','" & Format(DTPicker4, "dd-MMM-yyyy") & "'," & _
                 .SubItems(1) & ")"
                con.Execute Sql, a
            End If
        End With
    
    Next i
    
    
     For i = 1 To LVOPDECG.ListItems.Count
        With LVOPDECG.ListItems(i)
            If .SubItems(1) <> "" Then
                Sql = "Insert into OPDECGDailyIncome(SNo,DT,ECG) " & _
                 "Values('" & .Text & "','" & Format(DTPicker5, "dd-MMM-yyyy") & "'," & _
                 .SubItems(1) & ")"
                con.Execute Sql, a
            End If
        End With
    
    Next i
    
    
    
    Desc = lblOPD(0) & " OPDs Sno=" & txtFromOPD & " TO " & txtToOPD
    
    
    Call AddOPDVoucher(DTPicker4, Val(lblOPD(1)) + Val(lblOPD(3)), Val(lblOPD(1)) + Val(lblOPD(3)) - Val(lblOPD(4)), Val(lblOPD(4)), Desc)
    
    
    con.CommitTrans
    Saved = True
    
    
    


    
    
    Sql = ""
    Exit Function
err:
    MsgBox err.Description
End Function




Private Sub cmdAdd_Click()
Dim itm As ListItem
With LV.ListItems
    .Clear
    For i = Val(txtfrom) To Val(txtto)
        Set itm = .Add(, , i)
        itm.SubItems(1) = "- - -"
        itm.SubItems(2) = "- - -"
        itm.SubItems(3) = "- - -"
        itm.SubItems(4) = "- - -"
    Next
End With
End Sub

Private Sub cmdAddOPDECg_Click()
Dim itm As ListItem
With LVOPDECG.ListItems
    .Clear
    For i = Val(txtOECGfrom) To Val(txtOECGTo)
        Set itm = .Add(, , i)
        itm.SubItems(1) = "50"
    Next
End With

End Sub

Private Sub cmdAddPR_Click()
    Dim itm As ListItem
    With LVPR.ListItems
        .Clear
        For i = Val(txtFromPR) To Val(txtToPR)
            Set itm = .Add(, , i)
            itm.SubItems(1) = "- - -"
            itm.SubItems(2) = "- - -"
            itm.SubItems(3) = "- - -"
            itm.SubItems(4) = "- - -"
            itm.SubItems(5) = "- - -"
            
        Next
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If Saved Then
        Unload Me
        frmDailyIncome.Show
    End If
End Sub

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub

Private Sub cmdSave_Click()
    
End Sub


Private Sub cmdAddCCU_Click()

Dim itm As ListItem
With LVCCU.ListItems
    .Clear
    For i = Val(txtFromCCU) To Val(txtToCCU)
        Set itm = .Add(, , i)
        itm.SubItems(1) = "- - -"
        itm.SubItems(2) = "- - -"
    Next
End With

End Sub

Private Sub cmdAddOPD_Click()

Dim itm As ListItem
With LVOPD.ListItems
    .Clear
    For i = Val(txtFromOPD) To Val(txtToOPD)
        Set itm = .Add(, , i)
        itm.SubItems(1) = "10.00"
    Next
End With

End Sub

Private Sub Command1_Click()

End Sub

Private Sub DtPicker1_Change()
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select * from Dailyincomesheet Where DT ='" & Format(DTPicker1, "dd-MMM-yyyy") & "' Order By SNo", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set itm = LV.ListItems.Add(, , ![Sno] & "")
            itm.SubItems(1) = IIf(IsNull(![ECG]), "- - -", Format(![ECG] & "", "0.00"))
            itm.SubItems(2) = IIf(IsNull(![ETT]), "- - -", Format(![ETT] & "", "0.00"))
            itm.SubItems(3) = IIf(IsNull(![ECHO]), "- - -", Format(![ECHO] & "", "0.00"))
            itm.SubItems(4) = IIf(IsNull(![CardiacT]), "- - -", Format(![CardiacT] & "", "0.00"))
            itm.SubItems(5) = Format(Val(![ECG] & "") + Val(![ETT] & "") + Val(![ECHO] & "") + Val(![CardiacT] & ""), "0.00")
            
            .MoveNext
        Loop
    End With
    
    If LV.ListItems.Count > 0 Then
        txtfrom = LV.ListItems(1).Text
        txtto = LV.ListItems(LV.ListItems.Count).Text
        txtfrom.Enabled = False
        txtto.Enabled = False
    Else
        txtfrom = ""
        txtto = ""
        txtfrom.Enabled = True
        txtto.Enabled = True
    End If
    Call GetTotal
    
End Sub

Private Sub DTPicker3_Change()



    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select * from PRDailyIncome Where DT ='" & Format(DTPicker1, "dd-MMM-yyyy") & "' Order By SNo", con, adOpenForwardOnly, adLockReadOnly
        LVPR.ListItems.Clear
        Do Until .EOF
            Set itm = LVPR.ListItems.Add(, , ![Sno] & "")
            itm.SubItems(1) = ![RoomNo] & ""
            itm.SubItems(2) = Val(![Days] & "")
            itm.SubItems(3) = Format(Val(![Amt] & ""), "0.00")
            itm.SubItems(4) = IIf(IsNull(![Heater]), "- - -", Format(![Heater] & "", "0.00"))
            itm.SubItems(5) = IIf(IsNull(![AC]), "- - -", Format(![AC] & "", "0.00"))
            itm.SubItems(6) = Format(Val(itm.SubItems(3)) + Val(itm.SubItems(4)) + Val(itm.SubItems(5)), "0.00")
            
            .MoveNext
        Loop
    End With
    
    If LVPR.ListItems.Count > 0 Then
        txtFromPR = LV.ListItems(1).Text
        txtToPR = LV.ListItems(LVPR.ListItems.Count).Text
        txtFromPR.Enabled = False
        txtToPR.Enabled = False
    Else
        txtFromPR = ""
        txtToPR = ""
        txtFromPR.Enabled = True
        txtToPR.Enabled = True
    End If
    
    Call GetPRTotal
    
End Sub


Private Sub DTPicker5_Change()
Dim rs As New ADODB.Recordset
    With rs
        .Open "Select * from OPDECGECGDailyIncome Where DT ='" & Format(DTPicker1, "dd-MMM-yyyy") & "' Order By SNo", con, adOpenForwardOnly, adLockReadOnly
        LVOPDECG.ListItems.Clear
        
        Do Until .EOF
            Set itm = LVOPDECG.ListItems.Add(, , ![Sno] & "")
            itm.SubItems(1) = Format(Val(![Amt] & ""), "0.00")
            .MoveNext
        Loop
        
    End With
    
    If LVOPDECG.ListItems.Count > 0 Then
        txtFromOPDECG = LV.ListItems(1).Text
        txtToOPDECG = LV.ListItems(LVOPDECG.ListItems.Count).Text
        txtFromOPDECG.Enabled = False
        txtToOPDECG.Enabled = False
    Else
        txtFromOPDECG = ""
        txtToOPDECG = ""
        txtFromOPDECG.Enabled = True
        txtToOPDECG.Enabled = True
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()



    For Each LBL In Controls
        If TypeOf LBL Is Label And LCase(LBL.Name) Like "*rate" Then
            LBL.BorderStyle = 0
        End If
    Next

   

    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select * FRom VTHTCPerAmts", con, adOpenKeyset, adLockReadOnly
        
        '''''''''''' ECG ''''''''''''''
        .MoveFirst
        .Find "IncomeSource='ECG'"
        lblECG(1).Tag = Format(Val(![Price] & ""), "0.00")
        lblECG(2).Tag = Format(Val(![Cost] & ""), "0.00")
        lblECG(3).Tag = Format(Val(![Consaltant] & ""), "0.00")
        lblECG(4).Tag = Format(Val(![HCS] & ""), "0.00")
        
        lblECGRate(0) = Format(Val(![Price] & ""), "0.00")
        lblECGRate(1) = Format(Val(![Cost] & ""), "0.00")
        lblECGRate(2) = Format(Val(![ConsaltantPer] & ""), "0.00")
        lblECGRate(3) = Format(Val(![HCSPer] & ""), "0.00")
        
        '''''''''''' ETT ''''''''''''''
        .MoveFirst
        .Find "IncomeSource='ETT'"
        lblETT(1).Tag = Format(Val(![Price] & ""), "0.00")
        lblETT(2).Tag = Format(Val(![Cost] & ""), "0.00")
        lblETT(3).Tag = Format(Val(![Consaltant] & ""), "0.00")
        lblETT(4).Tag = Format(Val(![HCS] & ""), "0.00")
        
        lblETTRate(0) = Format(Val(![Price] & ""), "0.00")
        lblETTRate(1) = Format(Val(![Cost] & ""), "0.00")
        lblETTRate(2) = Format(Val(![ConsaltantPer] & ""), "0.00")
        lblETTRate(3) = Format(Val(![HCSPer] & ""), "0.00")
        
         '''''''''''' ECHO ''''''''''''
         .MoveFirst
         .Find "IncomeSource='ECHO'"
         lblECHO(1).Tag = Format(Val(![Price] & ""), "0.00")
         lblECHO(2).Tag = Format(Val(![Cost] & ""), "0.00")
         lblECHO(3).Tag = Format(Val(![Consaltant] & ""), "0.00")
         lblECHO(4).Tag = Format(Val(![HCS] & ""), "0.00")
         
         lblECHORate(0) = Format(Val(![Price] & ""), "0.00")
         lblECHORate(1) = Format(Val(![Cost] & ""), "0.00")
         lblECHORate(2) = Format(Val(![ConsaltantPer] & ""), "0.00")
         lblECHORate(3) = Format(Val(![HCSPer] & ""), "0.00")
         
         '''''''''''' CardiacT ''''''''
         .MoveFirst
         .Find "IncomeSource='CardiacT'"
         lblCardiacT(1).Tag = Format(Val(![Price] & ""), "0.00")
         lblCardiacT(2).Tag = Format(Val(![Cost] & ""), "0.00")
         lblCardiacT(3).Tag = Format(Val(![Consaltant] & ""), "0.00")
         lblCardiacT(4).Tag = Format(Val(![HCS] & ""), "0.00")
        
        
         lblCardiacTRate(0) = Format(Val(![Price] & ""), "0.00")
         lblCardiacTRate(1) = Format(Val(![Cost] & ""), "0.00")
         lblCardiacTRate(2) = Format(Val(![ConsaltantPer] & ""), "0.00")
         lblCardiacTRate(3) = Format(Val(![HCSPer] & ""), "0.00")
        
        
         .MoveFirst
         .Find "IncomeSource='CCU'"
         lblCCU(1).Tag = Format(Val(![Price] & ""), "0.00")
         lblCCU(2).Tag = Format(Val(![Consaltant] & ""), "0.00")
         lblCCU(3).Tag = Format(Val(![Alkhidmat] & ""), "0.00")
         lblCCU(4).Tag = Format(Val(![HCS] & ""), "0.00")
         
         
         lblCCURate(0) = Format(Val(![Price] & ""), "0.00")
         lblCCURate(1) = Format(Val(![ConsaltantPer] & ""), "0.00")
         lblCCURate(2) = Format(Val(![AlkhidmatPer] & ""), "0.00")
         lblCCURate(3) = Format(Val(![HCSPer] & ""), "0.00")
         
         
         
         
         .MoveFirst
         .Find "IncomeSource='PR'"
         lblPR(1).Tag = Format(Val(![Price] & ""), "0.00")
         lblPR(2).Tag = Format(Val(![Consaltant] & ""), "0.00")
         lblPR(3).Tag = Format(Val(![Alkhidmat] & ""), "0.00")
         lblPR(4).Tag = Format(Val(![HCS] & ""), "0.00")
         
         
         lblPRRate(0) = Format(Val(![Price] & ""), "0.00")
         lblPRRate(1) = Format(Val(![ConsaltantPer] & ""), "0.00")
         lblPRRate(2) = Format(Val(![AlkhidmatPer] & ""), "0.00")
         lblPRRate(3) = Format(Val(![HCSPer] & ""), "0.00")
         
         
         
         
         
         
         .MoveFirst
         .Find "IncomeSource='OPD'"
         
         lblOPD(1).Tag = Format(Val(![Price1] & ""), "0.00")
         lblOPD(3).Tag = Format(Val(![Price] & ""), "0.00")
         
         lblOPD(4).Tag = Format(Val(![HCS] & ""), "0.00")
         
         lblOPDRate(0) = Format(Val(![Price1] & ""), "0.00")
         lblOPDRate(1) = Format(Val(![Price] & ""), "0.00")
         
         
        .Close
    End With
    Set rs = Nothing
    
    
    DTPicker1 = GetServerDate
    DTPicker2 = GetServerDate
    DTPicker3 = GetServerDate
    DTPicker4 = GetServerDate
        
        
    Call DtPicker1_Change
    Call DTPicker2_Change
    Call DTPicker3_Change
    Call DTPicker4_Change
    
End Sub
Private Sub FillList()
      
    If Not cmbDept.MatchFound Then Exit Sub
    
    Dim Sql As String
    Dim LItem As ListItem
    Dim rs As New ADODB.Recordset
    Dim theKey As String
    Dim IfBold As Boolean
    Dim theName As String
    
    If cmbDept.ID = "0" Then
        Sql = "Select * from VEmp"
        
    Else
        Sql = "Select * from Employees Where DeptID='" & cmbDept.ID & "'"
        
    End If
    LV.ListItems.Clear
    With rs
        .Open Sql, con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            If IsNull(![EmpID]) Then
                theKey = ![DeptID] & ""
                IfBold = True
            Else
                theKey = ![EmpID] & ""
                IfBold = False
            End If
            
            Set LItem = LV.ListItems.Add(, theKey & "'", theKey)
            LItem.Tag = Val(![EmpType] & "")
            LItem.Bold = IfBold
            LItem.SubItems(1) = .Fields("Name") & ""
            LItem.SubItems(2) = ![FName] & ""
            LItem.SubItems(3) = ""
            LItem.SubItems(4) = ""
            'LItem.SubItems(3) = ![] & ""
            .MoveNext
        Loop
        .Close
    End With
    Sql = ""
    theKey = ""
    theName = ""
    
    Set rs = Nothing
    
End Sub



Private Sub Form_Resize()

On Error Resume Next
Dim TheHeight As Long


 'FAcc.Move (Me.ScaleWidth - FAcc.Width) / 2
 
 FLV.Move (Me.ScaleWidth - FLV.Width) / 2
 fTitle.Move FLV.Left
'cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 100
'cmdPrint.Move cmdClose.Left - cmdPrint.Width - 50, cmdClose.Top

'FDrag.Move 50, Me.ScaleHeight * 3 / 5, Me.ScaleWidth - 100
 
 'If LVItems.Visible Then
 '  TheHeight = FDrag.Top
 '  FDrag.Visible = True
 'Else
 '  TheHeight = Me.ScaleHeight - cmdClose.Height - 200
 '  FDrag.Visible = False
 'End If
 
 'LV.Move 50, FAcc.Top + FAcc.Height, Me.ScaleWidth - 100, TheHeight - (FAcc.Top + FAcc.Height)
 'LVItems.Move 50, FDrag.Top + FDrag.Height, Me.ScaleWidth - 100, Me.ScaleHeight - (FAcc.Height + cmdClose.Height + LV.Height + 250)
 'FDrag.ZOrder 1
 
End Sub

Private Sub Lv_DblClick()

    If LV.ListItems.Count = 0 Then Exit Sub
    With LV.SelectedItem
    
        Pic.Move LV.ColumnHeaders(2).Left + LV.Left, LV.Top + LV.SelectedItem.Top, _
        LV.ColumnHeaders(5).Left + LV.ColumnHeaders(5).Width - LV.ColumnHeaders(2).Left
        
        
        For i = 0 To Txts.Count - 1
            Txts(i).Move LV.ColumnHeaders(i + 2).Left - LV.ColumnHeaders(1).Width + 10, 0, LV.ColumnHeaders(i + 2).Width - 20
            Txts(i) = .SubItems(i + 1)
        Next
        
        Pic.Visible = True
        Txts(0).SetFocus
    End With
    
End Sub
Private Sub LVCCU_DblClick()

    If LVCCU.ListItems.Count = 0 Then Exit Sub
    With LVCCU.SelectedItem
    
        PicCCU.Move LVCCU.ColumnHeaders(2).Left + LVCCU.Left, LVCCU.Top + LVCCU.SelectedItem.Top, _
        LVCCU.ColumnHeaders(4).Left + LVCCU.ColumnHeaders(4).Width - LVCCU.ColumnHeaders(2).Left
        
        
        For i = 0 To txtsCCU.Count - 1
            txtsCCU(i).Move LVCCU.ColumnHeaders(i + 2).Left - LVCCU.ColumnHeaders(1).Width + 10, 0, LVCCU.ColumnHeaders(i + 2).Width - 20
            txtsCCU(i) = .SubItems(i + 1)
        Next
        
        PicCCU.Visible = True
        txtsCCU(0).SetFocus
    End With
    
End Sub
Private Sub LV_GotFocus()
Pic.Visible = False
End Sub





Private Sub GetTotal()

   Dim ECGAmt As Double, ETTAmt As Double, ECHOAmt As Double, CardiacTAmt As Double
   Dim ECG As Integer, ETT As Integer, ECHO As Integer, CardiacT As Integer

    
    For i = 1 To LV.ListItems.Count
        With LV.ListItems(i)
            If .SubItems(1) <> "- - -" Then
                ECG = ECG + 1
                ECGAmt = ECGAmt + Val(.SubItems(1))
            End If
            
            If .SubItems(2) <> "- - -" Then
                ECHO = ECHO + 1
                ECHOAmt = ECHOAmt + Val(.SubItems(2))
            End If
            
            If .SubItems(3) <> "- - -" Then
                ETT = ETT + 1
                ETTAmt = ETTAmt + Val(.SubItems(3))
            End If
            
            If .SubItems(4) <> "- - -" Then
                CardiacT = CardiacT + 1
                CardiacTAmt = CardiacTAmt + Val(.SubItems(4))
            End If
        End With
    Next
   
   
   
   '''''''' ECG Totals '''''''''
   
   lblECG(0) = ECG
   lblECG(1) = Format(ECG * Val(lblECG(1).Tag), "0.00")
   lblECG(2) = Format(ECG * Val(lblECG(2).Tag), "0.00")
   lblECG(3) = Format(ECG * Val(lblECG(3).Tag), "0.00")
   lblECG(4) = Format(ECG * Val(lblECG(4).Tag), "0.00")
   lblECG(5) = Format(Val(lblECG(1)) - ECGAmt, "0.00")
   lblECG(6) = Format(Val(lblECG(4)) - Val(lblECG(5)), "0.00")
   
  
   
   '''''''' ECHO Totals '''''''''
   
   lblECHO(0) = ECHO
   lblECHO(1) = Format(ECHO * Val(lblECHO(1).Tag), "0.00")
   lblECHO(2) = Format(ECHO * Val(lblECHO(2).Tag), "0.00")
   lblECHO(3) = Format(ECHO * Val(lblECHO(3).Tag), "0.00")
   lblECHO(4) = Format(ECHO * Val(lblECHO(4).Tag), "0.00")
   lblECHO(5) = Format(Val(lblECHO(1)) - ECHOAmt, "0.00")
   lblECHO(6) = Format(Val(lblECHO(4)) - Val(lblECHO(5)), "0.00")
   
    '''''''' ETT Totals '''''''''
   
   lblETT(0) = ETT
   lblETT(1) = Format(ETT * Val(lblETT(1).Tag), "0.00")
   lblETT(2) = Format(ETT * Val(lblETT(2).Tag), "0.00")
   lblETT(3) = Format(ETT * Val(lblETT(3).Tag), "0.00")
   lblETT(4) = Format(ETT * Val(lblETT(4).Tag), "0.00")
   lblETT(5) = Format(Val(lblETT(1)) - ETTAmt, "0.00")
   lblETT(6) = Format(Val(lblETT(4)) - Val(lblETT(5)), "0.00")
   
   
   
    '''''''' CardiacT Totals '''''''''
   
   lblCardiacT(0) = CardiacT
   lblCardiacT(1) = Format(CardiacT * Val(lblCardiacT(1).Tag), "0.00")
   lblCardiacT(2) = Format(CardiacT * Val(lblCardiacT(2).Tag), "0.00")
   lblCardiacT(3) = Format(CardiacT * Val(lblCardiacT(3).Tag), "0.00")
   lblCardiacT(4) = Format(CardiacT * Val(lblCardiacT(4).Tag), "0.00")
   lblCardiacT(5) = Format(Val(lblCardiacT(1)) - CardiacTAmt, "0.00")
   lblCardiacT(6) = Format(Val(lblCardiacT(4)) - Val(lblCardiacT(5)), "0.00")
   
   
End Sub




Private Sub LVOPDECG_DblClick()

    If LVOPDECG.ListItems.Count = 0 Then Exit Sub
    With LVOPDECG.SelectedItem
    
        txtsOPD(0).Move LVOPDECG.ColumnHeaders(2).Left, LVOPDECG.Top + .Top, LVOPDECG.ColumnHeaders(2).Width
        txtsOPD(0).Visible = True
        txtsOPD(0).Text = .SubItems(1)
        txtsOPD(0).SetFocus
    End With
    
End Sub

Private Sub txtfrom_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumber(KeyAscii)
End Sub

Private Sub txts_GotFocus(Index As Integer)
    Txts(Index).SelStart = 0
    Txts(Index).SelLength = Len(Txts(Index))
End Sub

Private Sub txts_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 3 And KeyAscii = 13 Then
        Dim TotalAmt As Double
        TotalAmt = 0
        With LV.SelectedItem
        
            For i = 0 To 3
                .SubItems(i + 1) = Format(Txts(i), "0.00")
                TotalAmt = TotalAmt + Val(Txts(i))
            Next
            .SubItems(5) = Format(TotalAmt, "0.00")
            
        End With
        Call GetTotal
        LV.SetFocus
    ElseIf KeyAscii = 27 Then
        LV.SetFocus
    Else
        KeyAscii = OnlyNumber(KeyAscii)
    End If
    
End Sub

Private Sub Txts_LostFocus(Index As Integer)
    If Trim(Txts(Index)) = "" Then Txts(Index) = "- - -"
    If LCase(Screen.ActiveControl.Name) <> "txts" Then Pic.Visible = False
End Sub

Private Sub txtto_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumber(KeyAscii)
End Sub

Private Sub txtsCCU_Change(Index As Integer)
    If Index = 1 Then txtsCCU(2) = Val(txtsCCU(1)) * Val(lblCCU(1).Tag)
End Sub



Private Sub txtsCCU_GotFocus(Index As Integer)
    txtsCCU(Index).SelStart = 0
    txtsCCU(Index).SelLength = Len(txtsCCU(Index))
End Sub

Private Sub txtsCCU_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If Index = 2 And KeyAscii = 13 Then
        
        With LVCCU.SelectedItem
        
            For i = 0 To 2
                .SubItems(i + 1) = Format(txtsCCU(i), "0")
            Next
            
        End With
        Call GetCCUTotal
        LVCCU.SetFocus
    ElseIf KeyAscii = 27 Then
        LVCCU.SetFocus
    Else
        KeyAscii = OnlyNumber(KeyAscii)
    End If
    
End Sub

Private Sub txtsCCU_LostFocus(Index As Integer)
    If Trim(txtsCCU(Index)) = "" Then txtsCCU(Index) = "- - -"
    If LCase(Screen.ActiveControl.Name) <> "txtsccu" Then PicCCU.Visible = False
End Sub



Private Sub GetCCUTotal()


   Dim CCU As Long, CCUAmt As Double

    
    For i = 1 To LVCCU.ListItems.Count
        With LVCCU.ListItems(i)
            If .SubItems(3) <> "- - -" Then
                CCU = CCU + Val(.SubItems(2))
                CCUAmt = CCUAmt + Val(.SubItems(3))
            End If
        End With
    Next
   
   
   
   '''''''' CCU Totals '''''''''
   
   lblCCU(0) = CCU
   lblCCU(1) = Format(CCU * Val(lblCCU(1).Tag), "0.00")
   lblCCU(2) = Format(CCU * Val(lblCCU(2).Tag), "0.00")
   lblCCU(3) = Format(CCU * Val(lblCCU(3).Tag), "0.00")
   lblCCU(4) = Format(CCU * Val(lblCCU(4).Tag), "0.00")
   lblCCU(5) = Format(Val(lblCCU(1)) - CCUAmt, "0.00")
   lblCCU(6) = Format(Val(lblCCU(4)) - Val(lblCCU(5)), "0.00")
   
  

End Sub


Private Sub GetPRTotal()


   Dim PR As Long, PRAmt As Double
    Dim HeaterAmt As Double, ACAmt As Double
    
    For i = 1 To LVPR.ListItems.Count
        With LVPR.ListItems(i)
            If .SubItems(3) <> "- - -" Then
                PR = PR + Val(.SubItems(2))
                PRAmt = PRAmt + Val(.SubItems(3))
            End If
            
            If .SubItems(4) <> "- - -" Then
                HeaterAmt = HeaterAmt + Val(.SubItems(4))
            End If
            
            If .SubItems(5) <> "- - -" Then
                ACAmt = ACAmt + Val(.SubItems(5))
            End If
            
        End With
    Next
   
   
   
   '''''''' PR Totals '''''''''
   
   lblPR(0) = PR
   lblPR(1) = Format(PR * Val(lblPR(1).Tag), "0.00")
   lblPR(2) = Format(PR * Val(lblPR(2).Tag), "0.00")
   lblPR(3) = Format(PR * Val(lblPR(3).Tag), "0.00")
   lblPR(4) = Format(PR * Val(lblPR(4).Tag), "0.00")
   lblPR(5) = Format(Val(lblPR(1)) - PRAmt, "0.00")
   lblPR(6) = Format(Val(lblPR(4)) - Val(lblPR(5)), "0.00")
   lblPR(7) = Format(ACAmt, "0.00")
   lblPR(8) = Format(HeaterAmt, "0.00")

End Sub

Private Sub LVPR_DblClick()

    If LVPR.ListItems.Count = 0 Then Exit Sub
    With LVPR.SelectedItem
    
        PicPR.Move LVPR.ColumnHeaders(2).Left + LVPR.Left, LVPR.Top + LVPR.SelectedItem.Top, _
        LVPR.ColumnHeaders(6).Left + LVPR.ColumnHeaders(6).Width - LVPR.ColumnHeaders(2).Left
        
        
        For i = 0 To txtsPR.Count - 1
            txtsPR(i).Move LVPR.ColumnHeaders(i + 2).Left - LVPR.ColumnHeaders(1).Width + 10, 0, LVPR.ColumnHeaders(i + 2).Width - 20
            txtsPR(i) = .SubItems(i + 1)
        Next
        
        PicPR.Visible = True
        txtsPR(0).SetFocus
    End With
    
End Sub


Private Sub txtspr_Change(Index As Integer)
    If Index = 1 Then txtsPR(2) = Val(txtsPR(1)) * Val(lblPR(1).Tag)
End Sub

Private Sub txtspr_GotFocus(Index As Integer)
    txtsPR(Index).SelStart = 0
    txtsPR(Index).SelLength = Len(txtsPR(Index))
End Sub

Private Sub txtspr_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 4 And KeyAscii = 13 Then
        
        With LVPR.SelectedItem
        
            For i = 0 To 4
                .SubItems(i + 1) = Format(txtsPR(i), "0")
            Next
            .SubItems(6) = Format(Val(.SubItems(3)) + Val(.SubItems(4)) + Val(.SubItems(5)), "0.00")
        End With
        
        Call GetPRTotal
        LVPR.SetFocus
        
    ElseIf KeyAscii = 27 Then
        LVPR.SetFocus
    Else
        KeyAscii = OnlyNumber(KeyAscii)
    End If
    
End Sub

Private Sub txtspr_LostFocus(Index As Integer)
    If Trim(txtsPR(Index)) = "" Then txtsPR(Index) = "- - -"
    If LCase(Screen.ActiveControl.Name) <> "txtspr" Then PicPR.Visible = False
End Sub


Private Sub DTPicker2_Change()



    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select * from CCUDailyIncome Where DT ='" & Format(DTPicker1, "dd-MMM-yyyy") & "' Order By SNo", con, adOpenForwardOnly, adLockReadOnly
        LVCCU.ListItems.Clear
        Do Until .EOF
            Set itm = LVCCU.ListItems.Add(, , ![Sno] & "")
            itm.SubItems(1) = ![BedNo] & ""
            itm.SubItems(2) = Val(![Days] & "")
            itm.SubItems(3) = Format(Val(![Amt] & ""), "0.00")
            .MoveNext
        Loop
    End With
    
    If LVCCU.ListItems.Count > 0 Then
        txtFromCCU = LV.ListItems(1).Text
        txtToCCU = LV.ListItems(LVCCU.ListItems.Count).Text
        txtFromCCU.Enabled = False
        txtToCCU.Enabled = False
    Else
        txtFromCCU = ""
        txtToCCU = ""
        txtFromCCU.Enabled = True
        txtToCCU.Enabled = True
    End If
    
    Call GetCCUTotal
    
End Sub


Private Sub txtsOPD_Change(Index As Integer)
'    If Index = 1 Then txtsOPD(2) = Val(txtsOPD(1)) * Val(lblOPD(1).Tag)
End Sub



Private Sub txtsOPD_GotFocus(Index As Integer)
    txtsOPD(Index).SelStart = 0
    txtsOPD(Index).SelLength = Len(txtsOPD(Index))
End Sub

Private Sub txtsOPD_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If Index = 0 And KeyAscii = 13 Then
        
        With LVOPDECG.SelectedItem
        
            For i = 0 To 0
                .SubItems(i + 1) = Format(txtsOPD(i), "0.00")
            Next
            '.SubItems(3) = Format(Val(.SubItems(1)) + Val(.SubItems(2)), "0.00")
        End With
        Call GetOPDTotal
        LVOPDECG.SetFocus
    ElseIf KeyAscii = 27 Then
        LVOPDECG.SetFocus
    Else
        KeyAscii = OnlyNumber(KeyAscii)
    End If
    
End Sub

Private Sub txtsOPD_LostFocus(Index As Integer)
    If Trim(txtsOPD(Index)) = "" Then txtsOPD(Index) = "- - -"
    If Index = 0 Then
        If LCase(Screen.ActiveControl.Name) <> "txtsopd" Then txtsOPD(0).Visible = False
    End If
End Sub



Private Sub GetOPDTotal()


   Dim OPD As Long, OPDAmt As Double, ECG As Long, OPDECGAmt As Double

    
    For i = 1 To LVOPD.ListItems.Count
        With LVOPD.ListItems(i)
            If .SubItems(1) <> "- - -" Then
                OPD = OPD + 1
                OPDAmt = OPDAmt + Val(.SubItems(1))
                
            End If
        End With
    Next
   
   
    For i = 1 To LVOPDECG.ListItems.Count
        With LVOPDECG.ListItems(i)
            If .SubItems(1) <> "- - -" Then
               ECG = ECG + 1
               OPDECGAmt = OPDECGAmt + Val(.SubItems(1))
            End If
        End With
    Next
   
   
   
   '''''''' OPD Totals '''''''''
   
   lblOPD(0) = OPD
   lblOPD(1) = Format(OPD * Val(lblOPD(1).Tag), "0.00")
   lblOPD(2) = ECG
   lblOPD(3) = Format(ECG * Val(lblOPD(3).Tag), "0.00")
   lblOPD(4) = Format(Val(lblOPD(3)) - OPDECGAmt, "0.00")
   lblOPD(5) = Format(OPDECGAmt + Val(lblOPD(1)), "0.00")
   
  

End Sub



Private Sub DTPicker4_Change()



    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select * from OPDDailyIncome Where DT ='" & Format(DTPicker1, "dd-MMM-yyyy") & "' Order By SNo", con, adOpenForwardOnly, adLockReadOnly
        LVOPD.ListItems.Clear
        
        Do Until .EOF
            Set itm = LVOPD.ListItems.Add(, , ![Sno] & "")
            itm.SubItems(1) = Format(Val(![Amt] & ""), "0.00")
            
            .MoveNext
        Loop
        
    End With
    
    If LVOPD.ListItems.Count > 0 Then
        txtFromOPD = LV.ListItems(1).Text
        txtToOPD = LV.ListItems(LVOPD.ListItems.Count).Text
        txtFromOPD.Enabled = False
        txtToOPD.Enabled = False
    Else
        txtFromOPD = ""
        txtToOPD = ""
        txtFromOPD.Enabled = True
        txtToOPD.Enabled = True
    End If
    
    Call GetOPDTotal
    
End Sub
