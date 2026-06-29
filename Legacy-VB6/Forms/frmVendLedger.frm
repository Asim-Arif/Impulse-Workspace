VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmVendLedger 
   ClientHeight    =   10485
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   10485
   ScaleMode       =   0  'User
   ScaleWidth      =   10641.08
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTMonth 
      Height          =   300
      Left            =   6675
      TabIndex        =   34
      Top             =   915
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "MMMM, yyyy"
      Format          =   108527619
      CurrentDate     =   38334
   End
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   11520
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vender Ledger"
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
         Index           =   1
         Left            =   4440
         TabIndex        =   32
         Top             =   120
         Width           =   2340
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vender Ledger"
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
         Index           =   2
         Left            =   4455
         TabIndex        =   33
         Top             =   135
         Width           =   2340
      End
   End
   Begin VB.Frame FLV 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   60
      TabIndex        =   6
      Top             =   8610
      Width           =   9555
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Rep. Deduction"
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
         Left            =   5865
         TabIndex        =   41
         Top             =   780
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblRepairDed 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   5880
         TabIndex        =   40
         Top             =   1005
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Label lblChequeNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   7515
         TabIndex        =   38
         Top             =   1005
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cheque No."
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
         Left            =   7545
         TabIndex        =   37
         Top             =   780
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lblAdvSal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   2040
         TabIndex        =   36
         Top             =   390
         Width           =   1890
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Advance Salary"
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
         Left            =   2085
         TabIndex        =   35
         Top             =   165
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Balance"
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
         Left            =   7485
         TabIndex        =   20
         Top             =   180
         Width           =   660
      End
      Begin VB.Label lblDone 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   105
         TabIndex        =   19
         Top             =   390
         Width           =   1905
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Total Work Done"
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
         Left            =   135
         TabIndex        =   18
         Top             =   165
         Width           =   1410
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Long Loan Ded."
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
         Left            =   5910
         TabIndex        =   17
         Top             =   165
         Width           =   1275
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Snaffing Bill"
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
         Left            =   2085
         TabIndex        =   16
         Top             =   780
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Plant Bill"
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
         Left            =   135
         TabIndex        =   15
         Top             =   780
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Res. Ded."
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
         Left            =   4035
         TabIndex        =   14
         Top             =   780
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblShort 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   3960
         TabIndex        =   13
         Top             =   390
         Width           =   1890
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Short Term Loan"
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
         Left            =   4035
         TabIndex        =   12
         Top             =   165
         Width           =   1410
      End
      Begin VB.Label lblResDed 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   3960
         TabIndex        =   11
         Top             =   1005
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label lblPlant 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   105
         TabIndex        =   10
         Top             =   1005
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label lblSnaff 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   2040
         TabIndex        =   9
         Top             =   1005
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label lblLong 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   5880
         TabIndex        =   8
         Top             =   390
         Width           =   1575
      End
      Begin VB.Label lblBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   7485
         TabIndex        =   7
         Top             =   390
         Width           =   1575
      End
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   1365
      TabIndex        =   2
      Top             =   915
      Width           =   2220
      _ExtentX        =   3916
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
   End
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   3570
      TabIndex        =   3
      Top             =   915
      Width           =   3090
      _ExtentX        =   5450
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
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10875
      Top             =   7740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   19
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVendLedger.frx":0000
            Key             =   "Issue"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7320
      Left            =   60
      TabIndex        =   21
      Top             =   1290
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   12912
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Issue Detail"
      TabPicture(0)   =   "frmVendLedger.frx":00D5
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LV"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtEdit"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Recieve Detail"
      TabPicture(1)   =   "frmVendLedger.frx":00F1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1"
      Tab(1).Control(1)=   "LV4"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Payments"
      TabPicture(2)   =   "frmVendLedger.frx":010D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LVPymt"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Loans"
      TabPicture(3)   =   "frmVendLedger.frx":0129
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).Control(1)=   "Frame1"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame2 
         Caption         =   "LongTerm Loan :"
         Height          =   6885
         Left            =   -69135
         TabIndex        =   28
         Top             =   420
         Width           =   5715
         Begin MSComctlLib.ListView LV3 
            Height          =   6585
            Left            =   60
            TabIndex        =   29
            Top             =   210
            Width           =   5580
            _ExtentX        =   9843
            _ExtentY        =   11615
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Date"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   1
               Text            =   "Payment"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Description"
               Object.Width           =   4410
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Short Term Loan :"
         Height          =   6885
         Left            =   -74880
         TabIndex        =   26
         Top             =   420
         Width           =   5715
         Begin MSComctlLib.ListView LV2 
            Height          =   6555
            Left            =   60
            TabIndex        =   27
            Top             =   210
            Width           =   5580
            _ExtentX        =   9843
            _ExtentY        =   11562
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Date"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   1
               Text            =   "Payment"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Description"
               Object.Width           =   4410
            EndProperty
         End
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6465
         TabIndex        =   23
         Top             =   1755
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68685
         TabIndex        =   22
         Top             =   2730
         Visible         =   0   'False
         Width           =   915
      End
      Begin MSComctlLib.ListView LV 
         Height          =   6795
         Left            =   105
         TabIndex        =   24
         Top             =   405
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   11986
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
            Text            =   "Order No."
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item ID"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item Name"
            Object.Width           =   6350
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Iss Qty"
            Object.Width           =   2822
         EndProperty
      End
      Begin MSComctlLib.ListView LV4 
         Height          =   6795
         Left            =   -74895
         TabIndex        =   25
         Top             =   405
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   11986
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
            Text            =   "Order  No"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item ID"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item Name"
            Object.Width           =   6350
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Qty"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Rate"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Amount"
            Object.Width           =   2822
         EndProperty
      End
      Begin MSComctlLib.ListView LVPymt 
         Height          =   6795
         Left            =   -74910
         TabIndex        =   30
         Top             =   420
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   11986
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
            Text            =   "Pymt ID"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin MSForms.CommandButton cmdFinal 
      Height          =   360
      Left            =   5550
      TabIndex        =   39
      Top             =   10095
      Visible         =   0   'False
      Width           =   2265
      ForeColor       =   0
      Caption         =   "Make Final"
      PicturePosition =   327683
      Size            =   "3995;635"
      Accelerator     =   70
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   360
      Left            =   7890
      TabIndex        =   5
      Top             =   10095
      Visible         =   0   'False
      Width           =   1425
      ForeColor       =   0
      Caption         =   "Print         "
      PicturePosition =   327683
      Size            =   "2514;635"
      Accelerator     =   80
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   360
      Left            =   9405
      TabIndex        =   4
      Top             =   10095
      Width           =   1425
      ForeColor       =   0
      Caption         =   " Close      "
      PicturePosition =   327683
      Size            =   "2514;635"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   1380
      TabIndex        =   1
      Top             =   645
      Width           =   8415
      BackColor       =   11517387
      Caption         =   "  Process                                Venders                                                     Month"
      Size            =   "14843;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   8490
      TabIndex        =   0
      Top             =   900
      Width           =   1305
      Caption         =   "Refresh        "
      PicturePosition =   327683
      Size            =   "2302;635"
      Accelerator     =   82
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmVendLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub PrintBalRpt()
    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    
    Load frmPrevRpt
    
    'rptPurchOrdForSteel.FormulaFields(5).Text = Val(LV.SelectedItem.SubItems(5) & "")
    'rptPurchOrdForSteel.FormulaFields(2).Text = "'" & strFromTo & "'"
    Dim rpt As CRAXDDRT.Report 'New rptVendLedger
    Set rpt = rptApp.OpenReport(RptPath & "\rptVendLedger.rpt") '
    'rpt.FormulaFields(4).Text = "' From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
    Dim FFDefs As FormulaFieldDefinitions
    Dim FFDef As CRAXDDRT.FormulaFieldDefinition
    Set FFDefs = rpt.FormulaFields
    
    For Each FFDef In FFDefs
        If FFDef.Name = "{@FromTo}" Then
            FFDef.Text = "' From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
            Exit For
        End If
    Next
    
    'rpt.FormulaFields("@FromTo").Text = "' From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
    'rpt.SetReportVariableValue
    frmPrevRpt.ShowReport "{VLedger.VendID}=" & cmbEmp.ID & " and ({VLedger.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & "))" & " ", rpt
    
    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub PrintIssRcvStatus()
    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    
    Load frmPrevRpt
    
    
    Dim rpt As CRAXDDRT.Report 'New rptVendIssRcvStatus
    Set rpt = rptApp.OpenReport(RptPath & "\rptVendIssRcvStatus.rpt") '
    
    Dim FFDefs As FormulaFieldDefinitions
    Dim FFDef As CRAXDDRT.FormulaFieldDefinition
    Set FFDefs = rpt.FormulaFields
    
    For Each FFDef In FFDefs
        If FFDef.Name = "{@FromTo}" Then
            FFDef.Text = "' From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
        ElseIf FFDef.Name = "{@EmpName}" Then
            FFDef.Text = "'" & cmbEmp.Text & "'"
        End If
    Next
    
    Dim subRpt As Report
    Set subRpt = rpt.OpenSubreport("ContRcvdItems")
    subRpt.RecordSelectionFormula = "{VLedger.VendID}=" & cmbEmp.ID & " and ({VLedger.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & "))" & " "
    
    frmPrevRpt.ShowReport "{VLedger.VendID}=" & cmbEmp.ID & " and ({VLedger.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & "))" & " ", rpt
    
    Screen.MousePointer = vbDefault

    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Private Sub cmdFinal_Click()
    
    On Error GoTo err
     'Check If salary Has been Saved or Not
    If cmbEmp.MatchFound = False Then Exit Sub
    
    
    
    
    If cmdFinal.Caption = "Delete Finalized" Then
        If MsgBox("Are You Sure To Delete This Saved Salary.", vbQuestion + vbYesNo) = vbYes Then
            con.Execute "Delete From VendMonthlyPayments Where Month(DT)=" & DTMonth.Month & " AND Year(DT)=" & DTMonth.year & " AND VendID=" & cmbEmp.ID & " AND AdvSalary=0"
            con.Execute "Delete From VendDeductions Where Month(DT)=" & DTMonth.Month & " AND Year(DT)=" & DTMonth.year & " AND VendID=" & cmbEmp.ID
            MsgBox "Salary Has Been Deleted. Now Click On Refresh Button To Calculate.", vbInformation
            cmdFinal.Caption = "Make Final"
        End If
    Else
        Dim rs As New ADODB.Recordset
        With rs
            .Open "Select Count(*) From VVendRcvdItems Where Month(DT)=" & DTMonth.Month & " AND Year(DT)=" & DTMonth.year & " AND VendID=" & cmbEmp.ID & " And ReqAuth=1", con, adOpenForwardOnly, adLockReadOnly
            If .Fields(0).Value > 0 Then
                MsgBox "There Are Some Entries Waiting For Authorization." & vbNewLine & "You Should Do Something With Them Before Continuing.", vbInformation
                .Close
                Exit Sub
            Else
                Dim DT As Date
                DT = DateSerial(DTMonth.year, DTMonth.Month + 1, 0)
                con.Execute "Insert Into VendMonthlyPayments(VendID,DT,WorkDone,Payed,ShortTermPer," & _
                 "LongLoan,STAmt,LTAmt,ResDed,AdvSalary,AdvAmt,PaidAmt,CashPaid,ChequeNo,PlantBill," & _
                 "SnaffBill,RepairDeduction) Values(" & cmbEmp.ID & ",'" & DT & "'," & Val(lblDone.Tag) & _
                 ",0,0," & Val(lblLong.Tag) & "," & Val(lblShort) & "," & Val(lblLong) & "," & _
                 Val(lblResDed) & ",0," & Val(lblAdvSal) & "," & lblBalance & ",1,''," & Val(lblPlant) & "," & Val(lblSnaff) & "," & Val(lblRepairDed) & ")"
                 
                con.Execute "Insert Into VendDeductions(VendID,DT,DedAmt) Values(" & cmbEmp.ID & _
                ",'" & DT & "'," & Val(lblLong) & ")"
                cmdFinal.Caption = "Delete Finalized"
            End If
            .Close
        End With
    End If
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdPrint_Click()
    
    Dim lTotalRcvd As Long
    Dim lTotalWastage As Long
    Dim lTotalRcvd1 As Long
    Dim lTotalRepair As Long
    
    Dim LastDT As Date
    Dim rpt As CRAXDDRT.Report
    Dim rs As New ADODB.Recordset
    Dim iVendType As Integer
    With rs
        .Open "Select VendType From Venders Where VendID=" & cmbEmp.ID, con, adOpenForwardOnly, adLockReadOnly
        iVendType = .Fields(0)
        .Close
    End With
    Set rs = Nothing
    LastDT = DateSerial(DTMonth.year, DTMonth.Month + 1, 0)
    
    con.Execute "Delete From PrintVendBalLedgerMonthly"
    con.Execute "Delete From PrintVendMonthlyIssdDetail"
     Dim FirstDT As Date, ItemCodes As String
    ItemCodes = ""
    FirstDT = DateSerial(DTMonth.year, DTMonth.Month, 1)
        
    'Insert Brought FOrward entries....
    con.Execute "Insert Into PrintVendMonthlyIssdDetail Select VendBalIssd.ItemCode,'Brought Forward' as RecieptID,'" & FirstDT & "' as DT,IsNull(VendBalIssd.IssdQty,0)-IsNull(VendBalRcvd.RcvdQty,0) as qty,VendBalIssd.VendID From " & _
     " (Select VendID,ItemCode,Sum(IssQty) As IssdQty From VVendIssItems1 Where DT<'" & FirstDT & "' and vendid=" & cmbEmp.ID & _
     " Group By VendID,ItemCode) VendBalIssd Left OUTER JOIN (Select VendID,ItemCode,Sum(RcvdQty+Wastage+LostQty) As RcvdQty " & _
     " From VVendRcvdDetail Where DT<'" & FirstDT & "' " & _
     " and vendid=" & cmbEmp.ID & " AND ReqAuth=0 Group By VendID,ItemCode) VendBalRcvd ON VendBalIssd.VendID=VendBalRcvd.VendID " & _
     " AND VendBalIssd.ItemCode=VendBalRcvd.ItemCode Where VendBalIssd.IssdQty-IsNull(VendBalRcvd.RcvdQty,0)<>0"
     
    'Now Insert This Month's Issuence
    con.Execute "Insert Into PrintVendMonthlyIssdDetail Select ItemCode,RecieptID,DT,IssQty  as qty,vendid from vvendissddetail where month(DT)=" & Month(FirstDT) & " and year(DT)=" & year(FirstDT) & " and vendid=" & cmbEmp.ID
 
    'Insert For Balances From This Months Issue And Receive
    con.Execute "Insert Into PrintVendBalLedgerMonthly Select " & cmbEmp.ID & ",IsNull(VendBalIssd.ItemCode,VendBalRcvd.ItemCode),0 As BFQty , VendBalIssd.IssdQty,VendBalRcvd.RcvdQty From " & _
     "(Select VendID,ItemCode,Sum(IssQty) As IssdQty From VVendIssItems1 Where Month(DT)=" & DTMonth.Month & " AND Year(DT)=" & DTMonth.year & _
     " AND VendID=" & cmbEmp.ID & " Group By VendID,ItemCode) VendBalIssd FULL OUTER JOIN (Select VendID,ItemCode,Sum(RcvdQty+Wastage+LostQty) As RcvdQty " & _
     "From VVendRcvdDetail Where Month(DT)=" & DTMonth.Month & " AND Year(DT)=" & DTMonth.year & _
     " AND VendID=" & cmbEmp.ID & " Group By VendID,ItemCode) VendBalRcvd ON VendBalIssd.VendID=VendBalRcvd.VendID " & _
     "AND VendBalIssd.ItemCode=VendBalRcvd.ItemCode Where IsNull(VendBalIssd.IssdQty,0)-IsNull(VendBalRcvd.RcvdQty,0)<>0", a
     
    'Now Update BFQty For Those Items Which Has Been Issd And Rcvd this Month
    con.Execute "Update PrintVendBalLedgerMonthly Set BFQty=t2.Qty From PrintVendBalLedgerMonthly t1 Inner JOIN PrintVendMonthlyIssdDetail t2 On t2.ItemCode = t1.ItemCode And t2.VendID = t1.VendID and t2.RecieptID='Brought Forward'", a
    'Now Insert Those Items Who ...
    'VendID,ItemCode,BFQty,0,0
    con.Execute "Insert Into PrintVendBalLedgerMonthly Select " & cmbEmp.ID & " As VendID,ItemCode,Qty,0,0 From PrintVendMonthlyIssdDetail Where RecieptID='Brought Forward' AND ItemCode Not In(Select ItemCode From PrintVendBalLedgerMonthly Where VendID=" & cmbEmp.ID & ")"
    
    With rs
        .Open "Select VendID,Sum(RcvdQty),Sum(Wastage) From VVendRcvdDetail Where Month(DT)=" & Month(DTMonth) & " and year(DT)=" & year(DTMonth) & " AND VendID=" & cmbEmp.ID & " Group By VendID", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            lTotalRcvd = Val(.Fields(1) & "")
            lTotalWastage = Val(.Fields(2) & "")
        Else
            lTotalRcvd = 0
            lTotalWastage = 0
        End If
        .Close
        .Open "Select Sum(T1.RcvdQty) As RcvdQty,Sum(IsNull(T2.RcvRepairQty,0)) As RepairQty From (Select RecieptID,ItemCode,DT,RcvdQty From VVendRcvdDetail Where ProcessID=7 AND VendID=" & cmbEmp.ID & ") T1 " & _
                "LEFT OUTER JOIN (Select RecieptID,DT As RepairRcvDT,RcvdQty,RcvRepairQty From VVendRcvdDetail Where ProcessID=12 ) T2 ON " & _
                "T1.RecieptID=T2.RecieptID INNER JOIN VItems1 On T1.ItemCode = VItems1.ItemID " & _
                "Where Month(IsNull(T2.RepairRcvDT, t1.DT))=" & Month(DTMonth) & " And year(IsNull(T2.RepairRcvDT, t1.DT))=" & year(DTMonth) & " AND Left(T1.RecieptID,2)='HS'", con, adOpenForwardOnly, adLockReadOnly
                
        If .EOF = False Then
            lTotalRcvd1 = Val(.Fields(0) & "")
            lTotalRepair = Val(.Fields(1) & "")
        Else
            lTotalRcvd1 = 0
            lTotalRepair = 0
        End If
        .Close
    End With
    Set rs = Nothing
    
    If iVendType = 2 Then
        Set rpt = rptApp.OpenReport(RptPath & "\rptVendLedgerSubFull.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\rptVendLedgerFull.rpt")
        rpt.FormulaFields(6).Text = lTotalRcvd
        rpt.FormulaFields(7).Text = lTotalWastage
        rpt.FormulaFields(9).Text = lTotalRcvd1
        rpt.FormulaFields(10).Text = lTotalRepair
    End If
    Dim f As New frmPrevRpt
    rpt.FormulaFields(3).Text = "#" & DateSerial(DTMonth.year, DTMonth.Month + 1, 0) & "#"
    f.ShowReport " Month({VendMonthlyPayments.DT})=Month({@ForDT}) AND Year({VendMonthlyPayments.DT})=Year({@ForDT}) AND {VendMonthlyPayments.AdvSalary}=False AND {VendMonthlyPayments.VendID}=" & cmbEmp.ID & " ", rpt
    
    Exit Sub
    Dim strRet As String
    strRet = ShowPopUpMenu1("Print &Balance Report", "-", "Print &Issue Receive Status")
    
    Select Case strRet
        Case "Print &Balance Report"
            Call PrintBalRpt
        Case "Print &Issue Receive Status"
            Call PrintIssRcvStatus
    End Select
    
End Sub

Public Sub cmdRefresh_Click()

    Me.MousePointer = vbHourglass
   
    Call RefreshLedger
    Me.MousePointer = vbNormal
    'Call Form_Resize
    
End Sub


Private Sub Form_Load()
    
    On Error GoTo err
    
    DTMonth = ServerDate
    
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1600
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Processes>", "0"
    cmbDept.AddVals con, "Description", "Processes", "ProcessID"
    cmbDept.ID = "0"
    'cmbEmp.Text = "<All Contractors>"
    'Set LV.SmallIcons = ImageList1
    SSTab1.TabVisible(2) = False
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub cmbdept_matched()
    
    On Error GoTo err
    cmbEmp.ClearVals
    'cmbEmp.AddItem "<All Contractors>", "0"
    'cmbEmp.AddItem "<All Contractors>", "0"
    If cmbDept.ID = "0" Then
        'cmbEmp.AddVals con, "'{' +  EmpID + '}  ' + Name", "Contractors", "EmpID"
        Call cmbEmp.AddVals(con, "'{' + VendID1 + '} ' + MakerName", "VMakers", "VendID")
    Else
        Call cmbEmp.AddVals(con, "MakerName", "VMakers", "VendID", " Where PhaseID=" & cmbDept.ID)
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Dim TheHeight As Long

    SSTab1.Move (Me.ScaleWidth - SSTab1.Width) / 2
    FLV.Move (Me.ScaleWidth - FLV.Width) / 2
 
 
    
End Sub

Private Sub RefreshLedger()
    
    On Error GoTo err
    Call GetBillingDetails
    
    If cmbDept.MatchFound = False Or cmbEmp.MatchFound = False Then Exit Sub
    Dim rsLedger As New ADODB.Recordset
    Dim ITM As ListItem
    Dim Cond As String
    Dim SQL As String
    
    
    Dim dWork As Double
    Dim dShort As Double
    Dim dLong As Double
    
    dWork = 0
    dShort = 0
    dLong = 0
    
    If cmbEmp.ID = "0" Then
        'If cmbDept.ID = "0" Then
            Cond = " Where "
        'Else
        '    Cond = " Where DeptID='" & cmbDept.ID & "' and"
        'End If
    Else
        Cond = " Where VendID=" & cmbEmp.ID & " and"
    End If
    
    Cond = Cond & "  Month(DT)=" & Month(DTMonth) & " AND Year(DT)=" & year(DTMonth)
    
    'Sql = "select * From VVendIssItems " & Cond & " Order By VendID,DT,ItemCode Asc"
    
    With rsLedger
        
            .Open "Select OrderNo,ItemCode,ItemDescription,Sum(IssQty) As Qty From VVendIssItems Where VendID =" & cmbEmp.ID & _
                "  AND Month(DT)=" & DTMonth.Month & " AND Year(DT)=" & DTMonth.year & " Group By OrderNo,ItemCode,ItemDescription", con, adOpenForwardOnly, adLockReadOnly
                
        LV.Visible = False
        LV2.Visible = False
        LV3.Visible = False
        LV4.Visible = False
        LVPymt.Visible = False
        
        LV.ListItems.Clear
        LV2.ListItems.Clear
        LV3.ListItems.Clear
        LV4.ListItems.Clear
        LVPymt.ListItems.Clear
        
        For i = 0 To .RecordCount - 1
            Set ITM = LV.ListItems.add(, , !OrderNo & "")
            ITM.ListSubItems.add , , !ItemCode & ""
            'itm.SmallIcon = "Issue"
            'ITM.Tag = ![VendID] & Format(![DT], "dd-MMM-yyyy") & ![CompItemCode] & ""
            ITM.ListSubItems.add , , !ItemDescription
            ITM.ListSubItems.add , , Val(!Qty & "")
            
            .MoveNEXT
        Next
        .Close
        
        Dim TotalRate As Double
        Dim LColor As Long
        '.Open "Select * From VVendRcvdItems  " & Cond & " and TransEntry Is Null and RcvdQty>0 Order By DT", con, adOpenForwardOnly, adLockReadOnly
        '.Open "Select * From VVenderLedger " & Cond & " Order By DT ", con, adOpenForwardOnly, adLockReadOnly
        '.Open "Select * From VVendRcvdItems " & Cond & " Order By DT ", con, adOpenForwardOnly, adLockReadOnly
        .Open "Select OrderNo,ItemCode,ItemDescription,Sum(RcvdQty) As Qty,Rate From VVendRcvdItems Where VendID =" & cmbEmp.ID & _
                " AND (Left(RecieptID,3)<>'RCV' OR Left(RecieptID,3)<>'HS') AND Month(DT)=" & DTMonth.Month & " AND Year(DT)=" & DTMonth.year & " AND RcvdQty>0 Group By OrderNo,ItemCode,ItemDescription,Rate"
                
        For i = 0 To .RecordCount - 1
            Set ITM = LV4.ListItems.add(, , !OrderNo)
            ITM.ListSubItems.add , , !ItemCode
            ITM.ListSubItems.add , , !ItemDescription
            ITM.ListSubItems.add , , Val(!Qty & "")
            ITM.ListSubItems.add , , Val(!Rate & "")
            ITM.ListSubItems.add , , Round(Val(!Qty & "") * Val(!Rate & ""))
            dWork = dWork + Val(ITM.SubItems(4))
            
            .MoveNEXT
        Next
        .Close
        
        '*********Wastage***********
        .Open "Select OrderNo,ItemCode,ItemDescription,Sum(Wastage) As Qty,0 As Rate From VVendRcvdItems Where VendID =" & cmbEmp.ID & _
                " AND (Left(RecieptID,3)<>'RCV' OR Left(RecieptID,3)<>'HS') AND Month(DT)=" & DTMonth.Month & " AND Year(DT)=" & DTMonth.year & " AND Wastage>0 Group By OrderNo,ItemCode,ItemDescription,Rate"
                
        For i = 0 To .RecordCount - 1
            If .AbsolutePosition = 1 Then
                Set ITM = LV4.ListItems.add()
                
                ITM.ListSubItems.add , , "********************W A S T A G E******************"
                ITM.ListSubItems.add , , ""
                ITM.ListSubItems.add , , ""
                ITM.ListSubItems.add , , ""
            End If
            Set ITM = LV4.ListItems.add(, , !ItemCode)
            ITM.ListSubItems.add , , !ItemDescription
            ITM.ListSubItems.add , , Val(!Qty & "")
            ITM.ListSubItems.add , , Val(!Rate & "")
            ITM.ListSubItems.add , , Round(Val(!Qty & "") * Val(!Rate & ""))
            dWork = dWork + Val(ITM.SubItems(4))
            
            .MoveNEXT
        Next
        .Close
        '*********Wastage***********
        .Open "Select OrderNo,ItemCode,ItemDescription,Sum(LostQty) As Qty,Rate From VVendRcvdItems Where VendID =" & cmbEmp.ID & _
                " AND (Left(RecieptID,3)<>'RCV' OR Left(RecieptID,3)<>'HS') AND Month(DT)=" & DTMonth.Month & " AND Year(DT)=" & DTMonth.year & " AND LostQty>0 Group By OrderNo,ItemCode,ItemDescription,Rate"
                
        For i = 0 To .RecordCount - 1
            If .AbsolutePosition = 1 Then
                Set ITM = LV4.ListItems.add()
                ITM.ListSubItems.add , , "********************L O S T******************"
                ITM.ListSubItems.add , , ""
                ITM.ListSubItems.add , , ""
                ITM.ListSubItems.add , , ""
            End If
            Set ITM = LV4.ListItems.add(, , !ItemCode)
            ITM.ListSubItems.add , , !ItemDescription
            ITM.ListSubItems.add , , Val(!Qty & "")
            ITM.ListSubItems.add , , Val(!Rate & "")
            ITM.ListSubItems.add , , Round(Val(!Qty & "") * Val(!Rate & ""))
            dWork = dWork + Val(ITM.SubItems(4))
            
            .MoveNEXT
        Next
        .Close
        'This Is For Payments.......
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '.Open "Select * From VenderLedger " & Cond & " and Amt>0 Order By DT", con, adOpenForwardOnly, adLockReadOnly
        
        'For i = 0 To .RecordCount - 1
            
            
        '    Set ITM = LVPymt.ListItems.Add(, , ![PymtID] & "")
        '    ITM.Tag = ![EntryID] & "'"
            
                
        '    ITM.SubItems(1) = Format(![Dt], "dd-MMM-yyyy")
        '    If Left(![PymtID], 1) = "C" Then
        '        ITM.SubItems(2) = "Cash Payed."
        '    ElseIf Left(![PymtID], 1) = "B" Then
        '        ITM.SubItems(2) = "Payment By Cheque # :" & ![ChqNo] & ""
        '    End If
            
        '    ITM.SubItems(3) = Val(![Amt] & "")
        '
        '    .MoveNext
        'Next
        '.Close
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        LV.Visible = True
        LV2.Visible = True
        LV3.Visible = True
        LV4.Visible = True
        LVPymt.Visible = True
        Set rsLedger = Nothing
        'Exit Sub
        
        'lblDone = dWork
        
        
        '.Open "Select * from ContPayment " & Cond & " and Type=0 Order by EmpID,DT", con, adOpenForwardOnly, adLockReadOnly
        .Open "Select * from VenderAdvances " & Cond & " and Type=0 Order by VendID,DT", con, adOpenForwardOnly, adLockReadOnly
        LV2.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set ITM = LV2.ListItems.add(, , Format(![DT], "dd-MMM-yyyy"))
            ITM.SubItems(1) = ![Amount] & ""
            ITM.SubItems(2) = ![Description] & ""
            dShort = dShort + Val(ITM.SubItems(1))
            .MoveNEXT
        Next i
        .Close
        
        .Open "Select * from VenderAdvances " & Cond & " and type=1 Order by VendID,DT", con, adOpenForwardOnly, adLockReadOnly
        LV3.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set ITM = LV3.ListItems.add(, , Format(![DT], "dd-MMM-yyyy"))
            ITM.SubItems(1) = ![Amount] & ""
            ITM.SubItems(2) = ![Description] & ""
            dLong = dLong + Val(ITM.SubItems(1))
            .MoveNEXT
        Next i
        .Close
        
    End With
    
'    lblShort = dShort
'    lblLong = dLong
    
    Exit Sub
    
    Dim cmd As New ADODB.Command
    
    
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdText
        If cmbEmp.ID = "0" Then
            .CommandText = "Select Sum(NetOpBal) From VContNetOp"
        Else
            .CommandText = "Select NetOpBal From VContNetOp Where EmpID='" & cmbEmp.ID & "'"
        End If
        .Parameters.Refresh
        .Parameters(0).Type = adDate
        .Parameters(0).Value = DtFrom
        '.Parameters(1).Type = adDate
        '.Parameters(1).value = DtFrom
        
        Set rsLedger = .Execute
        
        If rsLedger.RecordCount > 0 Then
            lblOpening = Val(rsLedger.Fields(0) & "")
        Else
            lblOpening = 0
        End If
        rsLedger.Close
        
        If cmbEmp.ID = "0" Then
            .CommandText = "Select Sum(NetShort) From VContNetOpShort"
        Else
            .CommandText = "Select NetShort From VContNetOpShort Where EmpID='" & cmbEmp.ID & "'"
        End If
        
        .Parameters.Refresh
        .Parameters(0).Type = adDate
        .Parameters(0).Value = DtFrom
       '.Parameters(1).Type = adDate
       '.Parameters(1).value = DtFrom
        
        Set rsLedger = .Execute
        
        If rsLedger.RecordCount > 0 Then
            lblOpShort = Val(rsLedger.Fields(0) & "")
        Else
            lblOpShort = 0
        End If
        
    End With
        
    
    lblShort = dShort
    
    lblTotalShort = dShort + Val(lblOpShort)
    lblLong = dLong
    lblTotLong = Val(lblLong) + Val(lblOpening)
    lblTotalLoan = Val(lblTotalShort) + Val(lblTotLong)
    lblBalance = Val(lblDone) - Val(lblTotalLoan)
    
'    If Val(lblBalance) > 0 Then
'        Label1.Caption = "Payable"
'    Else
'        Label1.Caption = "Recievable"
'    End If
    
    LVPymt.Visible = False
    
    
    LV.Visible = True
    LV2.Visible = True
    LV3.Visible = True
    LV4.Visible = True
    
    Set rsLedger = Nothing
    SQL = ""
    Cond = ""
    
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub LV_DblClick()
    Exit Sub
    'If LV.ListItems.Count = 0 Or UserHasAccess("EditOverTime") = False Then Exit Sub
    
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(5).Left + 35, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(5).Width
        .Text = LV.SelectedItem.SubItems(4)
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With

End Sub



Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    If KeyAscii = 13 Then
        con.Execute "Update ContLedger Set Rate=" & Val(txtEdit) & " Where Empid+format(DT,'dd-MMM-yyyy')+ItemId='" & LV.SelectedItem.Tag & "'"
        LV.SelectedItem.SubItems(4) = Val(txtEdit)
        LV.SetFocus
    ElseIf KeyAscii = 27 Then
        LV.SetFocus
    Else
        KeyAscii = OnlyNumber(KeyAscii, True)
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub

Private Sub GetBillingDetails()

    'Clear All************************************************************************
    lblDone = ""
    lblDone.Tag = ""
    lblAdvSal = ""
    lblShort = ""
    lblLong = ""
    lblResDed = ""
    lblBalance = ""
    lblChequeNo = ""
    lblPlant = ""
    lblSnaff = ""
    lblRepairDed = ""
    '*******************************************
    
    Dim GrandTotal As Double
    Dim TotalDebit As Double
    Dim TotalCredit As Double
    If cmbDept.MatchFound = False Or cmbEmp.MatchFound = False Then Exit Sub
    
    Dim VendID As Long
    Dim ProcessID As Long
    
    Dim ITM As ListItem
    Dim Cond As String
    Dim SQL As String
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    'Dim rsLedger As New ADODB.Recordset
    
    
    VendID = cmbEmp.ID
    
    
    
    Dim TDate As Date
    Dim SDate As Date 'For Saving
    'TDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), Day(ServerDate))
    
    SDate = DateSerial(DTMonth.year, DTMonth.Month, 1)
    TDate = DateSerial(DTMonth.year, DTMonth.Month + 1, 0)
    If TDate > ServerDate Then
        TDate = DateSerial(DTMonth.year, DTMonth.Month, Day(ServerDate))
    End If
    
    LastDay = DateAdd("m", 1, SDate) - SDate
    
    Dim JDate As Date
    If TDate <= ServerDate Then
        TotalSalDays = Day(TDate)
    ElseIf Month(ServerDate) < Month(TDate) Then
        MsgBox "Can't Show Entries Of Future Months.", vbInformation
        Exit Sub
    End If
    
    TempTotalSaldays = TotalSalDays    'Store For Later Use
    
    JDate = DateSerial(DTMonth.year, DTMonth.Month, TotalSalDays)
        
    'Sql = "Select * from VAdvanceSalary " & Cond & " JoinDate<='" & Format(JDate, "dd-MMM-yyyy") & "' Order By DeptID,EmpID"
    
    'cmd.CommandText = Sql
    
    Dim bSaved As Boolean
    With rs
        .Open "Select * From VendMonthlyPayments Where VendID=" & cmbEmp.ID & " AND Month(DT)=" & DTMonth.Month & " AND Year(DT)=" & DTMonth.year & " AND AdvSalary=0", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            cmdFinal.Caption = "Make Final"
            bSaved = False
        Else
            bSaved = True
            cmdFinal.Caption = "Delete Finalized"
            lblDone = Round(Val(!WorkDone & "") + Val(!PlantBill & "") + Val(!SnaffBill & ""))
            lblDone.Tag = Round(Val(!WorkDone & ""))
            lblAdvSal = Round(Val(!AdvAmt & ""))
            lblShort = Round(Val(!STAmt & ""))
            lblLong = Round(Val(!LTAmt & ""))
            lblResDed = Round(Val(!ResDed & ""))
            lblBalance = Round(Val(!PaidAmt & ""))
            lblChequeNo = !ChequeNo & ""
            lblPlant = Round(Val(!PlantBill & ""))
            lblSnaff = Round(Val(!SnaffBill & ""))
            lblRepairDed = Round(Val(!RepairDeduction & ""))
        End If
        .Close
    End With
    
    If bSaved Then
        Set rs = Nothing
        Exit Sub
    End If
    
    With cmd
        .CommandType = adCmdStoredProc
    
        .CommandText = "SP_VendSalary"
    
        Set .ActiveConnection = con
    
        .Parameters("@ParDate").Value = JDate
        .Parameters("@VendID").Value = VendID
        .Parameters("@ProcessID").Value = ProcessID
        .Parameters("@TotMonDays").Value = TotalSalDays
        
        rs.CursorType = adOpenKeyset
        rs.LockType = adLockReadOnly
        Set rs = .Execute
    End With
    
    
    
    With rs
        
        '.Open Sql, Con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount = 0 Then
            lblDone = 0
            lblAdvSal = 0
            lblShort = 0
            lblLong = 0
            lblResDed = 0
            lblBalance = 0
            lblChequeNo = 0
            lblPlant = 0
            lblSnaff = 0
            lblRepairDed = 0
            Exit Sub
        End If
        
        Dim total As Double
        Dim Payed As Double
        Dim STLoan As Double
        Dim LongLoan As Double
        Dim VendWorkDone As Double
        
        Dim CashPaid As Boolean
        Dim ChequeNo As String
        
        Dim IsComputed As Boolean
        
        Dim DT As String
        Dim IsPayed As Boolean
        Dim IsTransfered As Boolean
        
        
        'Long Term Loan Deduction
        Dim lDedAmt As Long, lClsAmt As Long, lTotAmt As Long
        If IsNull(![DedAmt]) Then lDedAmt = 0 Else lDedAmt = ![DedAmt]
        If IsNull(![totamt]) Then lTotAmt = 0 Else lTotAmt = ![totamt]
        If IsNull(![ClsAmt]) Then lClsAmt = 0 Else lClsAmt = ![ClsAmt]
        If IsNull(![total]) Then VendWorkDone = 0 Else VendWorkDone = Val(![total] & "")
        
        If lDedAmt > (lTotAmt - lClsAmt) Then
            LongLoan = (lTotAmt - lClsAmt)
        Else
            LongLoan = lDedAmt
        End If
            
            
        
        'lblDone = Round(Val(VendWorkDone & "") + Val(!PlantBill & "") + Val(!SnaffBill & ""))
        lblDone = Round(Val(VendWorkDone & ""))
        lblDone.Tag = Round(Val(VendWorkDone & ""))
        lblShort = Round(Val(![Payed] & ""))
        lblLong = Round(LongLoan)
        lblLong.Tag = Round(Val(!totamt & "") - Val(!ClsAmt & "")) 'Round(Val(!totamt & ""))
        lblResDed = Round(Val(![TotalResDeduction] & ""))
        'itm.SubItems(5) = Val(itm.SubItems(2)) - Val(itm.SubItems(3)) '- LongLoan
        lblAdvSal = Round(Val(![AdvSalpaid] & ""))
        lblRepairDed = Round(Val(!RepairDeduction & ""))
        'itm.SubItems(6) =
        'lblPlant = Round(Val(!PlantBill & ""))
        'lblSnaff = Round(Val(!SnaffBill & ""))
        
        Dim NetPayable As Long
            
            'NetPayable = Round(Val(itm.SubItems(5)) * Val(![AdvSalPer] & "") / 10000, 0) * 100
        NetPayable = Round((Val(lblDone.Tag) - Val(lblShort) - LongLoan - Val(lblResDed) - Val(lblAdvSal) - Val(lblRepairDed)) / 10, 0) * 10
        
        lblBalance = Abs(NetPayable)
        
'        If NetPayable = 0 Then
'            ITM.SubItems(8) = ""
'        ElseIf NetPayable > 0 Then
'            ITM.SubItems(8) = "DB"
'            TotalDebit = TotalDebit + NetPayable
'        Else
'            ITM.SubItems(8) = "CR"
'            TotalCredit = TotalCredit - NetPayable
'        End If
            
        GrandTotal = GrandTotal + NetPayable
    
    End With
    DTMonth.Tag = DTMonth.Value
    
    
End Sub
