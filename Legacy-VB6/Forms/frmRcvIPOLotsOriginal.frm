VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmRcvIPOLotsOriginal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Receiving Lots"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14580
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
   ScaleHeight     =   8070
   ScaleWidth      =   14580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5280
      Left            =   270
      TabIndex        =   22
      Top             =   2700
      Width           =   14250
      _ExtentX        =   25135
      _ExtentY        =   9313
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Issuance / Receiving"
      TabPicture(0)   =   "frmRcvIPOLotsOriginal.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdClose"
      Tab(0).Control(1)=   "cmdReceive"
      Tab(0).Control(2)=   "txtRate"
      Tab(0).Control(3)=   "cmbMaker"
      Tab(0).Control(4)=   "LV"
      Tab(0).Control(5)=   "Image1"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Inspection"
      TabPicture(1)   =   "frmRcvIPOLotsOriginal.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2(7)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2(6)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label2(8)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmbProcessNo"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label2(9)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "LV1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "DTInspection"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdPrint"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtRemarks"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtDisposation"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtEdit"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cmdSaveInspection"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "OptInspection(0)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "OptInspection(1)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "Rework Record"
      TabPicture(2)   =   "frmRcvIPOLotsOriginal.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdPrintRework"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txtRejQtyRW"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txtAccQtyRW"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "txtRWDone"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "txtRWReason"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "txtRWQty"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "LVRW"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label2(15)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label2(14)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label2(13)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label2(12)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label2(11)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "cmbProcessRW"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label2(10)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "Rejection Record"
      TabPicture(3)   =   "frmRcvIPOLotsOriginal.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label2(19)"
      Tab(3).Control(1)=   "Label2(20)"
      Tab(3).Control(2)=   "cmbProcessRJ"
      Tab(3).Control(3)=   "Label2(21)"
      Tab(3).Control(4)=   "LVRJ"
      Tab(3).Control(5)=   "txtReasonRJ"
      Tab(3).Control(6)=   "txtRejQty"
      Tab(3).Control(7)=   "cmdPrintRej"
      Tab(3).ControlCount=   8
      Begin VB.CommandButton cmdPrintRej 
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
         Left            =   -62295
         TabIndex        =   54
         Top             =   4545
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
         Left            =   -72105
         TabIndex        =   53
         Top             =   765
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
         Left            =   -70695
         TabIndex        =   55
         Top             =   765
         Width           =   9870
      End
      Begin VB.CommandButton cmdPrintRework 
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
         Left            =   -62280
         TabIndex        =   7
         Top             =   4575
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
         Left            =   -61845
         TabIndex        =   5
         Top             =   795
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
         Left            =   -62910
         TabIndex        =   4
         Top             =   795
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
         Left            =   -65565
         TabIndex        =   3
         Top             =   795
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
         Left            =   -70695
         TabIndex        =   2
         Top             =   795
         Width           =   5115
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
         Left            =   -72105
         TabIndex        =   1
         Top             =   795
         Width           =   1395
      End
      Begin VB.OptionButton OptInspection 
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
         Left            =   9630
         TabIndex        =   41
         Top             =   4830
         Width           =   1200
      End
      Begin VB.OptionButton OptInspection 
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
         Left            =   8385
         TabIndex        =   40
         Tag             =   "-1"
         Top             =   4830
         Width           =   1035
      End
      Begin VB.CommandButton cmdSaveInspection 
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
         Left            =   11160
         TabIndex        =   39
         Top             =   4590
         Width           =   1470
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7155
         TabIndex        =   32
         Top             =   1230
         Visible         =   0   'False
         Width           =   840
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
         Left            =   60
         TabIndex        =   31
         Top             =   4800
         Width           =   2415
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
         Left            =   2520
         TabIndex        =   30
         Top             =   4800
         Width           =   4095
      End
      Begin VB.CommandButton cmdPrint 
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
         Left            =   12720
         TabIndex        =   28
         Top             =   4590
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
         Left            =   -66330
         TabIndex        =   26
         Top             =   4575
         Width           =   1470
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
         Left            =   -67890
         TabIndex        =   25
         Top             =   4575
         Width           =   1470
      End
      Begin VB.TextBox txtRate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -67890
         TabIndex        =   23
         Top             =   990
         Visible         =   0   'False
         Width           =   840
      End
      Begin ComboList.Usercontrol1 cmbMaker 
         Height          =   285
         Left            =   -71670
         TabIndex        =   24
         Tag             =   "Planter"
         Top             =   975
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
         Left            =   -74910
         TabIndex        =   27
         Top             =   405
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
      Begin MSComCtl2.DTPicker DTInspection 
         Height          =   390
         Left            =   6675
         TabIndex        =   29
         Top             =   4800
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
         Format          =   151584771
         CurrentDate     =   40357
      End
      Begin MSComctlLib.ListView LV1 
         Height          =   3300
         Left            =   60
         TabIndex        =   33
         Top             =   1185
         Width           =   14130
         _ExtentX        =   24924
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
      Begin MSComctlLib.ListView LVRW 
         Height          =   3300
         Left            =   -74940
         TabIndex        =   6
         Top             =   1215
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
         NumItems        =   6
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
            Object.Width           =   9022
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   4657
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Width           =   1852
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Width           =   1852
         EndProperty
      End
      Begin MSComctlLib.ListView LVRJ 
         Height          =   3300
         Left            =   -74955
         TabIndex        =   48
         Top             =   1185
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   21
         Left            =   -74970
         TabIndex        =   52
         Top             =   435
         Width           =   870
      End
      Begin MSForms.ComboBox cmbProcessRJ 
         Height          =   390
         Left            =   -74955
         TabIndex        =   51
         Top             =   765
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   20
         Left            =   -72120
         TabIndex        =   50
         Top             =   450
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   19
         Left            =   -70710
         TabIndex        =   49
         Top             =   450
         Width           =   2445
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   15
         Left            =   -61830
         TabIndex        =   47
         Top             =   480
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   14
         Left            =   -62895
         TabIndex        =   46
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   13
         Left            =   -65565
         TabIndex        =   45
         Top             =   480
         Width           =   1860
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   12
         Left            =   -70695
         TabIndex        =   44
         Top             =   480
         Width           =   2235
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   11
         Left            =   -72105
         TabIndex        =   43
         Top             =   480
         Width           =   1335
      End
      Begin MSForms.ComboBox cmbProcessRW 
         Height          =   390
         Left            =   -74940
         TabIndex        =   0
         Top             =   795
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   10
         Left            =   -74955
         TabIndex        =   42
         Top             =   465
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   9
         Left            =   45
         TabIndex        =   38
         Top             =   435
         Width           =   870
      End
      Begin MSForms.ComboBox cmbProcessNo 
         Height          =   390
         Left            =   60
         TabIndex        =   37
         Top             =   765
         Width           =   14130
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "24924;688"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   8
         Left            =   60
         TabIndex        =   36
         Top             =   4440
         Width           =   2310
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   6
         Left            =   2535
         TabIndex        =   35
         Top             =   4440
         Width           =   1920
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   7
         Left            =   6615
         TabIndex        =   34
         Top             =   4440
         Width           =   1650
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   4740
         Left            =   -64860
         Stretch         =   -1  'True
         Top             =   420
         Width           =   4035
      End
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
      Left            =   12360
      TabIndex        =   21
      Top             =   1440
      Width           =   2010
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
      Left            =   11595
      TabIndex        =   20
      Top             =   2280
      Width           =   2775
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
      Left            =   2295
      TabIndex        =   17
      Top             =   2280
      Width           =   9285
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
      Left            =   270
      TabIndex        =   15
      Top             =   2280
      Width           =   2010
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
      Left            =   2295
      TabIndex        =   13
      Top             =   1440
      Width           =   10035
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
      Left            =   270
      TabIndex        =   11
      Top             =   1440
      Width           =   2010
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
      Left            =   5985
      TabIndex        =   9
      Top             =   540
      Width           =   2010
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   5
      Left            =   12360
      TabIndex        =   19
      Top             =   1035
      Width           =   2010
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   4
      Left            =   12075
      TabIndex        =   18
      Top             =   1890
      Width           =   2010
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   3
      Left            =   2280
      TabIndex        =   16
      Top             =   1905
      Width           =   870
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   2
      Left            =   255
      TabIndex        =   14
      Top             =   1905
      Width           =   2010
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   1
      Left            =   2295
      TabIndex        =   12
      Top             =   1035
      Width           =   1230
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   0
      Left            =   270
      TabIndex        =   10
      Top             =   1035
      Width           =   2010
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   6390
      TabIndex        =   8
      Top             =   135
      Width           =   1020
   End
End
Attribute VB_Name = "frmRcvIPOLotsOriginal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iColNo As Integer

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
            OptInspection(Abs(!LotStatus)).Value = True
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
    Set rpt = rptApp.OpenReport(RptPath & "\QAIR.rpt")
    Dim f As New frmPrevRpt
    Me.Hide
    f.ShowReport "{ProcessInspection.IPO_RefID}=" & Val(txtLotBarcode), rpt
    
    Unload Me
    
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
                    MsgBox "No Rates for " & .ListSubItems(1).Text
                    bDataNotComplete = True
                    Exit For
                End If
            End If
        End With
    Next
    If bDataNotComplete Then
        Exit Sub
    End If
    
    Dim myDT As Date
    myDT = Date
    Call StartTrans(con)
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If Val(.ListSubItems(3).Tag) = 1 Then 'Issuance Only
                Call IssueLot(Val(txtLotBarcode.Text), txtItemCode.Text, Val(.Tag), myDT, Val(.ListSubItems(1).Tag), txtItemCode.Tag, Val(.SubItems(2)))
            ElseIf Val(.ListSubItems(3).Tag) = 2 Then 'Receiving Only
                Call ReceiveLot(Val(txtLotBarcode.Text), txtItemCode.Text, Val(.Tag), myDT)
            ElseIf Val(.ListSubItems(3).Tag) = 3 Then 'Issuance & Receiving
                Call IssueLot(Val(txtLotBarcode.Text), txtItemCode.Text, Val(.Tag), myDT, Val(.ListSubItems(1).Tag), txtItemCode.Tag, Val(.SubItems(2)))
                Call ReceiveLot(Val(txtLotBarcode.Text), txtItemCode.Text, Val(.Tag), myDT)
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

Private Sub ReceiveLot(lIPOLotNo As Long, strItemID As String, lProcessID As Long, myDT As Date)
    
    Dim rs As New ADODB.Recordset
    Dim lMakerID As Long, lIssMainEntryID As Long, lIssEntryID As Long, lQty As Long, lLotNo As String
    Dim strOrderNo As String
    With rs
        .Open "SELECT * FROM VVendIssuanceIPOLots WHERE ItemCode='" & strItemID & "' AND ProcessID=" & lProcessID & " AND IPODL_RefID=" & lIPOLotNo, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            lMakerID = Val(!VendID & "")
            lIssMainEntryID = Val(!MainEntryID & "")
            lIssEntryID = Val(!EntryID & "")
            lQty = Val(!IssQty & "")
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
            .Parameters("@WasteQty").Value = 0 'Val(LV.ListItems(i).SubItems(5) & "")
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
    
    con.CommitTrans
    LV1.Tag = lEntryID
    
    'Unload Me
    Exit Sub
err:
    MsgBox err.Description
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

Private Sub OptInspection_Click(Index As Integer)
    OptInspection(0).Tag = Index
End Sub

Private Sub Text5_Change()

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
                txtLotQty = Val(!IssQty & "")
                txtProcess = !Description & ""
                txtProcess.Tag = Val(!ProcessID & "")
                txtIssuanceDate = Format(!DT, "dd-MMM-yyyy")
                txtLotBarcode.Tag = Val(!EntryID & "")
                txtOrderNo = !OrderNo & ""
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
        If Val(txtRate) <= 0 Then
            MsgBox "Please Enter Rate.", vbInformation
            Exit Sub
        End If
        'Now Check If this Item is Assigned at this Process...
        Dim iAffect As Integer
        iAffect = 0
        con.Execute "UPDATE VendAssItems SET Rate=" & Val(txtRate) & " WHERE VendID=" & _
         LV.SelectedItem.ListSubItems(1).Tag & " AND ItemID='" & txtItemCode & "'", iAffect
        
        If iAffect = 0 Then 'Not Assigned
            con.Execute "INSERT INTO VendAssItems(VendID,ProcessID,ItemID,Rate,PlantRate,StampRate,SnaffRate,Unit)" & _
             "VALUES(" & LV.SelectedItem.ListSubItems(1).Tag & ",0,'" & txtItemCode & _
             "'," & Val(txtRate) & ",0,0,0,'" & _
             GetSingleStringValue("Unit", "Items", " WHERE ItemID='" & txtItemCode & "'") & "')"
        End If
        LV.SelectedItem.SubItems(2) = Val(txtRate)
        
        txtRate.Visible = False
        Call txtLotBarcode_KeyPress(13)
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
            'ITM.ListSubItems.add , , !Description & ""
            ITM.ListSubItems.add , , Val(!ReWorkQty & "")
            ITM.ListSubItems.add , , !Reason & ""
            ITM.ListSubItems.add , , !DoneBy & ""
            ITM.ListSubItems.add , , Val(!AccQty & "")
            ITM.ListSubItems.add , , Val(!RejQty & "")
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

