VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmVenderBilling 
   ClientHeight    =   9405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   9405
   ScaleMode       =   0  'User
   ScaleWidth      =   10721.69
   WindowState     =   2  'Maximized
   Begin VB.Frame FAcc 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   9255
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   11865
      Begin VB.Frame fTitle 
         Height          =   570
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Width           =   11865
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor Billing"
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
            Index           =   3
            Left            =   0
            TabIndex        =   40
            Top             =   120
            Width           =   11820
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor Billing"
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
            Index           =   4
            Left            =   0
            TabIndex        =   41
            Top             =   135
            Width           =   11820
         End
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   195
         TabIndex        =   0
         Top             =   870
         Width           =   4995
         _ExtentX        =   8811
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
      Begin ComboList.Usercontrol1 cmbDebitHead 
         Height          =   285
         Left            =   195
         TabIndex        =   35
         Top             =   1380
         Width           =   11520
         _ExtentX        =   20320
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
      Begin VB.TextBox txtDeductShort 
         Alignment       =   2  'Center
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   5820
         TabIndex        =   30
         Top             =   7740
         Width           =   1455
      End
      Begin VB.TextBox txtFixLoan 
         Alignment       =   2  'Center
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   7290
         TabIndex        =   29
         Top             =   7740
         Width           =   1455
      End
      Begin VB.TextBox txtMaterialDed 
         Alignment       =   2  'Center
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   8760
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   7740
         Width           =   1455
      End
      Begin VB.TextBox txtManualDed 
         Alignment       =   2  'Center
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   10230
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   7740
         Width           =   1455
      End
      Begin VB.TextBox txts1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   1
         Left            =   3735
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   8085
         Width           =   1125
      End
      Begin VB.TextBox txts1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   0
         Left            =   2565
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   8085
         Width           =   1125
      End
      Begin VB.TextBox txts 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   1
         Left            =   3735
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   7725
         Width           =   1125
      End
      Begin VB.TextBox txts 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   0
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   7725
         Width           =   1125
      End
      Begin VB.ComboBox cmbRange 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmVenderBilling.frx":0000
         Left            =   5220
         List            =   "frmVenderBilling.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   855
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   8655
         TabIndex        =   2
         Top             =   855
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   556
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
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   147980291
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   6930
         TabIndex        =   1
         Top             =   855
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   556
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
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   147980291
         CurrentDate     =   37055
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5775
         Left            =   210
         TabIndex        =   19
         Top             =   1695
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   10186
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         TabCaption(0)   =   "Vender Billing"
         TabPicture(0)   =   "frmVenderBilling.frx":0066
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "LV"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtEdit"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "FraFind"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Material Deductions"
         TabPicture(1)   =   "frmVenderBilling.frx":0082
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "LVMaterial"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Manual Deductions"
         TabPicture(2)   =   "frmVenderBilling.frx":009E
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label1(1)"
         Tab(2).Control(1)=   "LVManualMaterial"
         Tab(2).Control(2)=   "txtManualMatRate"
         Tab(2).Control(3)=   "txtManualMatDesc"
         Tab(2).ControlCount=   4
         Begin VB.Frame FraFind 
            Height          =   495
            Left            =   105
            TabIndex        =   42
            Top             =   5250
            Width           =   7005
            Begin VB.ComboBox cmbfield 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               ItemData        =   "frmVenderBilling.frx":00BA
               Left            =   60
               List            =   "frmVenderBilling.frx":00D0
               Style           =   2  'Dropdown List
               TabIndex        =   45
               Top             =   135
               Width           =   3420
            End
            Begin VB.CommandButton cmdFind 
               Caption         =   "&Find"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   5925
               TabIndex        =   44
               Top             =   120
               Width           =   1020
            End
            Begin VB.TextBox txtFind 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   3480
               TabIndex        =   43
               Top             =   135
               Width           =   2430
            End
         End
         Begin VB.TextBox txtManualMatDesc 
            Appearance      =   0  'Flat
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   -74910
            TabIndex        =   24
            Top             =   660
            Width           =   10350
         End
         Begin VB.TextBox txtManualMatRate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   -64575
            TabIndex        =   23
            Top             =   660
            Width           =   1000
         End
         Begin VB.TextBox txtEdit 
            Height          =   285
            Left            =   2325
            TabIndex        =   20
            Top             =   1485
            Visible         =   0   'False
            Width           =   690
         End
         Begin MSComctlLib.ListView LV 
            Height          =   4830
            Left            =   105
            TabIndex        =   21
            Top             =   420
            Width           =   11340
            _ExtentX        =   20003
            _ExtentY        =   8520
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   8388608
            BackColor       =   15199215
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S. #"
               Object.Width           =   873
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Order #"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Date"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Order Type"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Material"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Qty"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "Rate"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "Value"
               Object.Width           =   1764
            EndProperty
         End
         Begin MSComctlLib.ListView LVMaterial 
            Height          =   5280
            Left            =   -74925
            TabIndex        =   22
            Top             =   390
            Width           =   11340
            _ExtentX        =   20003
            _ExtentY        =   9313
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   8388608
            BackColor       =   15199215
            BorderStyle     =   1
            Appearance      =   0
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
               Text            =   "S. #"
               Object.Width           =   873
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Iss #"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Date"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Raw Material"
               Object.Width           =   4586
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Rate"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Qty"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "Amount"
               Object.Width           =   1764
            EndProperty
         End
         Begin MSComctlLib.ListView LVManualMaterial 
            Height          =   4785
            Left            =   -74910
            TabIndex        =   25
            Top             =   930
            Width           =   11340
            _ExtentX        =   20003
            _ExtentY        =   8440
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   8388608
            BackColor       =   15199215
            BorderStyle     =   1
            Appearance      =   0
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
               Text            =   "S. #"
               Object.Width           =   873
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Amount"
               Object.Width           =   1764
            EndProperty
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"frmVenderBilling.frx":0108
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
            Height          =   255
            Index           =   1
            Left            =   -74910
            TabIndex        =   26
            Top             =   420
            Width           =   11340
         End
      End
      Begin MSComCtl2.DTPicker DTPosting 
         Height          =   315
         Left            =   3540
         TabIndex        =   37
         Top             =   8430
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
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
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   148307971
         CurrentDate     =   37055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Posting Date :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   11
         Left            =   2295
         TabIndex        =   38
         Top             =   8490
         Width           =   1170
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Debit Head"
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
         Height          =   255
         Index           =   2
         Left            =   195
         TabIndex        =   36
         Top             =   1185
         Width           =   11130
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Short Term Loan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Index           =   12
         Left            =   5835
         TabIndex        =   34
         Top             =   7530
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Long Term"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Index           =   13
         Left            =   7365
         TabIndex        =   33
         Top             =   7530
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Material Ded."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Index           =   15
         Left            =   8775
         TabIndex        =   32
         Top             =   7545
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Manual Ded."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Index           =   16
         Left            =   10245
         TabIndex        =   31
         Top             =   7545
         Width           =   1035
      End
      Begin MSForms.CommandButton cmdPrintZero 
         Height          =   375
         Left            =   5160
         TabIndex        =   18
         Top             =   8355
         Visible         =   0   'False
         Width           =   2025
         ForeColor       =   0
         Caption         =   "Print Balance Order"
         PicturePosition =   327683
         Size            =   "3572;661"
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrintUnPosted 
         Height          =   375
         Left            =   7230
         TabIndex        =   17
         Top             =   8370
         Width           =   1800
         ForeColor       =   0
         Caption         =   "Print Un-Posted"
         PicturePosition =   327683
         Size            =   "3175;661"
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Finalized Bills :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   10
         Left            =   825
         TabIndex        =   15
         Top             =   8130
         Width           =   1680
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Value"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   2
         Left            =   4065
         TabIndex        =   13
         Top             =   7515
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Vender's Total Performance :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   7770
         Width           =   2445
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Qty"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   2985
         TabIndex        =   10
         Top             =   7515
         Width           =   270
      End
      Begin MSForms.CommandButton cmdPost 
         Height          =   375
         Left            =   9060
         TabIndex        =   8
         Top             =   8355
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Post Selected"
         PicturePosition =   327683
         Size            =   "2355;661"
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
         Height          =   375
         Left            =   10440
         TabIndex        =   7
         Top             =   8355
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "2355;661"
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   360
         Left            =   10260
         TabIndex        =   3
         Top             =   840
         Width           =   1500
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "2646;635"
         Accelerator     =   82
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmVenderBilling.frx":01F0
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
         Height          =   255
         Index           =   0
         Left            =   165
         TabIndex        =   5
         Top             =   645
         Width           =   11130
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   10
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVenderBilling.frx":029A
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVenderBilling.frx":0456
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PopMenu"
   End
End
Attribute VB_Name = "frmVenderBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCourierInvoice As Boolean
Dim bShipping As Boolean
Dim bCustomInvoice As Boolean
Dim strAdvanceEntryList As String
Dim strSQLRpt As String
Dim bSampling As Boolean

Public Sub ShowMe(Optional p_bSampling As Boolean = True)
    bSampling = p_bSampling
    Me.Show
End Sub

Public Sub ShowForShipping()
    bShipping = True
    Me.Show
End Sub

Private Sub cmbCust_matched()
    Call RefreshLV
End Sub

Private Sub cmbRange_Click()

    DtTo = GetServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    Select Case cmbRange.ListIndex
        Case 0
            DtFrom = Date
        Case 1
            DtFrom = DateAdd("d", -15, Date)
        Case 2
            DtFrom = DateAdd("d", -30, Date)
        Case 3
            DtFrom = DateAdd("d", -60, Date)
        Case 4
            DtFrom = DateAdd("d", -90, Date)
        Case 5
            DtFrom.Enabled = True
            DtTo.Enabled = True
            DtFrom.SetFocus
            Exit Sub
    End Select
    
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdPost_Click()

    On Error GoTo err
    Dim i As Integer
    Dim bSelected As Boolean
    bSelected = False
    
    If cmbDebitHead.MatchFound = False Then
        MsgBox "Please Select Debit Head.", vbInformation
        Exit Sub
    End If
    
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            bSelected = True
            Exit For
        End If
    Next
    If bSelected = False Then
        MsgBox "No Order Selected.", vbInformation
        Exit Sub
    End If
    Dim dValue As Double
    dValue = 0
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            dValue = dValue + Val(LV.ListItems(i).SubItems(7))
        End If
    Next
    dValue = Round(dValue)
    'Get Bill No...
    Dim lBillNo As Long, strMakerID As String, strDesc As String, strVchrNo As String, strDebitAccNo As String, strMakerAccNo As String
    'strMakerID = GetSingleStringValue("VendID1", "Makers", " WHERE VendID=" & cmbCust.ID)
    lBillNo = GetSingleLongValue("MAX(BillNo)", "VenderPostedBills", " WHERE BillNo IS NOT NULL AND AccNo='" & cmbCust.ID & "'")
    lBillNo = lBillNo + 1
   
    strDebitAccNo = cmbDebitHead.ID
    strMakerAccNo = cmbCust.ID
     
    Dim objVchr As New clsVouchers
    objVchr.VoucherDate = DTPosting 'GetServerDate
    objVchr.VoucherType = "JV"
    
    Dim dAmtToPost As Double
    
    dAmtToPost = dValue - Val(txtManualDed) - Val(txtMaterialDed)
    
    If dAmtToPost <= 0 Then
        MsgBox "Invalid Amount to Post.", vbInformation
        Exit Sub
    End If
    
    strDesc = "Computer Bill No. " & cmbCust.Text & " / " & lBillNo & " - " & year(DTPosting)
    objVchr.AddVoucherHead strDebitAccNo, dAmtToPost, strDesc

    objVchr.AddVoucherHead strMakerAccNo, -(dAmtToPost), strDesc
     
    Call StartTrans(con)
    strVchrNo = objVchr.PostVoucher(, True)
    
    con.Execute "INSERT INTO VenderPostedBills(AccNo,DT,UserName,MachineName,BillNo,ManualDedAmt,STDeduction,LTDeduction,MaterialDedAmt," & _
     "VchrNo,DTFrom,DTTo,SampleBilling) VALUES('" & cmbCust.ID & "','" & GetServerDate & "','" & CurrentUserName & "','" & strComputerName & "'," & lBillNo & _
      "," & Val(txtManualDed) & "," & Val(txtDeductShort) & "," & Val(txtFixLoan) & "," & Val(txtMaterialDed) & ",'" & strVchrNo & "','" & DtFrom.Value & "','" & DtTo.Value & "'," & Abs(bSample) & ")"
    
    Dim lEntryID As Long
    lEntryID = GetSingleLongValue("Max(EntryID)", "VenderPostedBills")
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                con.Execute "INSERT INTO VenderPostedBillsDetail(RefID,VRDPO_RefID) VALUES(" & lEntryID & "," & Val(.Tag) & ")"
            End If
        End With
    Next
    
    For i = 1 To LVMaterial.ListItems.count
        With LVMaterial.ListItems(i)
            con.Execute "INSERT INTO VenderPostedBillsIssuanceDetails(RefID,VRMID_RefID) VALUES(" & lEntryID & "," & Val(.Tag) & ")"
        End With
    Next
    
    'Now Save Manual Material Ded. Entries...
    For i = 1 To LVManualMaterial.ListItems.count
        With LVManualMaterial.ListItems(i)
            con.Execute "INSERT INTO VenderPostedBills_ManualDeductions(Description,Amount,VPB_RefID) VALUES('" & _
             .SubItems(1) & "'," & Val(.SubItems(2)) & "," & lEntryID & ")"
        End With
    Next
        
    If Val(txtFixLoan) > 0 Then
        con.Execute "INSERT INTO MakerAmtCleared(AccNo,DT,AmtClrd,VchrNo) VALUES('" & cmbCust.ID & "','" & _
         DTPosting & "'," & Val(txtFixLoan) & ",'" & strVchrNo & "')"
    End If
        
    If Val(txtDeductShort) > 0 Then 'Clear All Short Term... Deducting All....
        con.Execute "INSERT INTO VenderPostedBills_STDeductions(VPB_RefID,Amount) VALUES(" & lEntryID & "," & Val(txtDeductShort) & ")"
    End If
    
    con.CommitTrans
    
    LVManualMaterial.ListItems.Clear
    
    MsgBox "Posted Successfuly !!!", vbInformation
    Call PrintVenderPostedBill(strVchrNo)
    Call RefreshLV
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub PostRepairVchr(ByRef p_strVchrNo As String, ByRef p_dRepairAmt As Double, ByRef p_dWastageAmt As Double)

    Dim i As Integer, bYouCanPost As Boolean
    bYouCanPost = False
'    For i = 1 To LVRepair.ListItems.count
'        If LVRepair.ListItems(i).Checked Then
'            bYouCanPost = True
'            Exit For
'        End If
'    Next
    If LVRepair.ListItems.count > 0 Then
        bYouCanPost = True
    End If
    
    If bYouCanPost = False Then Exit Sub
    
    Dim dRepairAmt As Double, dWastageAmt As Double
    For i = 1 To LVRepair.ListItems.count
        With LVRepair.ListItems(i)
            If .Checked Then
                dRepairAmt = dRepairAmt + Val(.SubItems(5))
                dWastageAmt = dWastageAmt + Val(.SubItems(6))
            End If
        End With
    Next
    'Now Post...
    Dim dAmt As Double, strRepairAccNo As String, strVchrNo As String, strWastageAccNo As String, strMakerAccNo As String
    Dim strDescription As String
    strRepairAccNo = "41-004-13001"
    strWastageAccNo = "41-003-19001"
    strMakerAccNo = GetSingleStringValue("AccNo", "Makers", " WHERE VendID=" & cmbCust.ID)
    strDescription = "Repair / Wastage Charges to " & cmbCust.Text
    
    dAmt = dWastageAmt + dRepairAmt
    Dim objVchr As New clsVouchers
    objVchr.VoucherType = "JV"
    objVchr.VoucherDate = GetServerDate(False)
    objVchr.AddVoucherHead strMakerAccNo, dAmt, strDescription
    If dRepairAmt > 0 Then
        objVchr.AddVoucherHead strRepairAccNo, -dRepairAmt, strDescription
    End If
    If dWastageAmt > 0 Then
        objVchr.AddVoucherHead strWastageAccNo, -dWastageAmt, strDescription
    End If
    
    strVchrNo = objVchr.PostVoucher(False, True)
    
    For i = 1 To LVRepair.ListItems.count
        With LVRepair.ListItems(i)
            If .Checked Then
                con.Execute "UPDATE MakerRepair SET VchrNo='" & strVchrNo & "' WHERE EntryID=" & Val(.key)
            End If
        End With
    Next
         
    p_strVchrNo = strVchrNo
    p_dRepairAmt = dRepairAmt
    p_dWastageAmt = dWastageAmt
    
    'Call AddRepairDetail
    
End Sub

Private Sub cmdPrintUnPosted_Click()
    Call PrintUnPostedBills(cmbCust.ID)
End Sub

Private Sub PrintUnPostedBills(strAccNo As String)
    
    On Error GoTo err
     
    Dim i As Integer, strEntryIDs As String
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                strEntryIDs = strEntryIDs & Val(.Tag) & ","
            End If
        End With
    Next
    If strEntryIDs <> "" Then
        strEntryIDs = Left(strEntryIDs, Len(strEntryIDs) - 1)
    Else
        MsgBox "Please Select Entries.", vbInformation
        Exit Sub
    End If
    Dim strSelection As String
    strSelection = "{VVendRcvdDetailPO.EntryID} IN[" & strEntryIDs & "]"
    Dim rpt As CRAXDDRT.Report
    
    Set rpt = rptApp.OpenReport(RptPath & "\VenderUnPostedBill.rpt")
    
    rpt.FormulaFields.GetItemByName("UserName").Text = "'" & CurrentUserName & "'"
    rpt.FormulaFields.GetItemByName("ComputerName").Text = "'" & strComputerName & "'"
    
    
    
    Dim f As New frmPrevRpt
        
    f.ShowReport strSelection, rpt
    'f.ShowReport "", rpt, strSQLRpt
    
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub cmdPrintZero_Click()
    Call PrintUnPostedBills(cmbCust.ID)
End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub



Private Sub DtFrom_Change()
    'Call setDateRange(DtFrom)
End Sub

Private Sub DtTo_Change()
    'Call setDateRange(DtTo)
End Sub


Private Sub Form_Load()
    
    mnuPop.Visible = False
    cmbCust.ListHeight = 3800
    'cmbCust.AddItem "<All Makers>", "0"
    'cmbCust.AddVals con, "AccTitle", "VVendersANDMakers", "AccNo", " WHERE Active = 1  ORDER BY AccTitle"
    cmbCust.AddVals con, "AccTitle", "VVenders", "AccNo", " WHERE Active = 1  ORDER BY AccTitle"
    
    With cmbDebitHead
        .ListHeight = 4000
        .AddVals con, "AccTitle + '{ ' + AccNo + ' }'", "VActiveAccounts", "AccNo"
    End With
    
'    DtFrom = Date
'    DtTo = Date
    cmbRange.ListIndex = 1
'    SSTab1.TabVisible(2) = False
    DTPosting.Value = Date
    SSTab1.Tab = 0
        
    Dim i As Integer
    cmbfield.Clear
    For i = 1 To LV.ColumnHeaders.count
        cmbfield.AddItem LV.ColumnHeaders(i)
    Next
    cmbfield.ListIndex = 1
        
    bSampling = False
End Sub

Private Sub RefreshLV()
        
    On Error GoTo err
    If cmbCust.MatchFound = False Then
        MsgBox "Please Select Maker.", vbInformation
        Exit Sub
    End If
    
    Dim Cond As String
    
    Dim ITM As ListItem
      
    Screen.MousePointer = vbHourglass
    
    Dim i As Integer
    For i = 0 To 1
        txts(i) = ""
        txts(i).Tag = ""
        txts1(i) = ""
        txts1(i) = ""
    Next
    Dim rs As New ADODB.Recordset, strOrderType As String, dNetQty As Double
    Dim strCondition As String
    strCondition = " WHERE RcvDate BETWEEN '" & DtFrom & "' AND '" & DtTo & "' AND EntryID NOT IN(SELECT VRDPO_RefID FROM VenderPostedBillsDetail) AND VendID='" & cmbCust.ID & "'"
    strCondition = strCondition & " AND EntryID NOT IN(SELECT VRDPO_RefID FROM VendRcvdDetailPO_Hidden_Billing)"
    If bSample Then
        strCondition = strCondition & " AND SampleOrder=1"
    Else
        strCondition = strCondition & " AND SampleOrder=0"
    End If
    
    With rs
        
        strSQLRpt = "SELECT * FROM VVendRcvdDetailPO " & strCondition & " ORDER BY RcvDate"
        
        .Open strSQLRpt, con, adOpenForwardOnly, adLockReadOnly
        
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, Val(!EntryID & "") & "'", .AbsolutePosition)
            ITM.Tag = Val(![EntryID] & "")
            ITM.ListSubItems.add(, , !PORefNo & "").ToolTipText = !RMName & ""
            ITM.ListSubItems.add(, , Format(!RcvDate, "dd-MM-yy")).ToolTipText = !RMName & ""
            
            If Val(!IssuanceType & "") = 1 Then
                strOrderType = "Annealing"
            ElseIf Val(!IssuanceType & "") = 2 Then
                strOrderType = "Machining"
            Else
                strOrderType = "Regular"
            End If
            
            ITM.ListSubItems.add(, , strOrderType).ToolTipText = !RMName & ""
            ITM.ListSubItems.add , , "(" & !MaterialID & ") " & !RMName & ""
            dNetQty = Val(!QtyRcvd & "") + Val(!ExtraQtyRcvd & "") - Val(!QtyRejected & "")
            ITM.ListSubItems.add , , dNetQty
            ITM.ListSubItems.add , , Val(!Rate & "")
            ITM.ListSubItems.add , , (dNetQty) * Val(!Rate & "")
            
            txts(0) = Val(txts(0)) + Val(ITM.SubItems(5))   'IssQty
            txts(1) = Val(txts(1)) + Val(ITM.SubItems(7))   'Value
            
            .MoveNEXT
        Loop
        .Close
        
        .Open "SELECT TotalAmountTaken,TotalAmountCleared,ActualDeductionAmount FROM VMakerAdvancesDeduction WHERE AccNo='" & cmbCust.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            txtFixLoan.Text = "0" 'Val(!ActualDeductionAmount & "")
            txtFixLoan.Tag = "0" 'Val(!TotalAmountTaken & "") - Val(!TotalAmountCleared & "") - Val(!ActualDeductionAmount & "")
        Else
            txtFixLoan.Text = Val(!ActualDeductionAmount & "")
            txtFixLoan.Tag = Val(!TotalAmountTaken & "") - Val(!TotalAmountCleared & "") - Val(!ActualDeductionAmount & "")
        End If
        .Close
        
        'Get Short Term Deduction ....
        '.Open "SELECT SUM(Amount-AmountCleared) AS TotalDeductionAmt FROM VMakerAdvancesShort WHERE AccNo='" & cmbCust.Tag & "'", con, adOpenForwardOnly, adLockReadOnly
        .Open "SELECT * FROM VVenderAdvancesShortDeduction WHERE AccNo='" & cmbCust.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            txtDeductShort.Text = Val(!Amount & "") - Val(!AmountCleared & "")
            txtDeductShort.Tag = Val(!Amount & "") - Val(!AmountCleared & "")
        Else
            txtDeductShort.Text = 0
            txtDeductShort.Tag = "0"
        End If
        .Close
    End With
    
    Set rs = Nothing
    
    Call AddMaterialIssuanceDetail
        
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Private Sub AddRepairDetail()

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "SELECT * FROM VMakerRepair WHERE VendID=" & cmbCust.ID & " AND VchrNo IS NULL", con, adOpenForwardOnly, adLockReadOnly
        LVRepair.ListItems.Clear
        Do Until .EOF
            Set ITM = LVRepair.ListItems.add(, !EntryID & "'", LVRepair.ListItems.count + 1)
            ITM.ListSubItems.add , , !RecieptID & ""
            ITM.ListSubItems.add , , Format(!DT, "dd-MMM-yyyy")
            ITM.ListSubItems.add , , !ItemCode & ""
            ITM.ListSubItems.add , , Val(!Rate & "")
            ITM.ListSubItems.add , , Val(!RepairAmt & "")
            ITM.ListSubItems.add , , Val(!Qty & "") * Val(!Rate & "")
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub


Private Sub AddMaterialIssuanceDetail()

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "SELECT * FROM VRawMaterialIssuance WHERE AccNo='" & cmbCust.ID & "' AND VchrNo IS NULL ORDER BY DT", con, adOpenForwardOnly, adLockReadOnly
        LVMaterial.ListItems.Clear
        Do Until .EOF
            Set ITM = LVMaterial.ListItems.add(, !EntryID & "'", LVMaterial.ListItems.count + 1)
            ITM.Tag = Val(!EntryID & "")
            ITM.ListSubItems.add , , !IssNo & ""
            ITM.ListSubItems.add , , Format(!DT, "dd-MMM-yyyy")
            ITM.ListSubItems.add , , "(" & !RMID1 & ") " & !RMName
            ITM.ListSubItems.add , , Val(!Price & "")
            ITM.ListSubItems.add , , Val(!Qty & "")
            ITM.ListSubItems.add , , Val(!Qty & "") * Val(!Price & "")
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Call getMaterailIssuanceTotal
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    FAcc.Top = (Me.ScaleHeight - FAcc.Height) / 2
End Sub


Private Sub LV_DblClick()
    'Edit Rate
    If LV.ListItems.count = 0 Then Exit Sub
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(7).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(7).Width
        .Text = Val(LV.SelectedItem.SubItems(6))
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
End Sub

Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    Dim i As Integer
    For i = 0 To 1
        txts1(i) = ""
        txts1(i).Tag = ""
    Next
    
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                txts1(0) = Val(txts1(0)) + Val(.SubItems(5))   'IssQty
                txts1(1) = Val(txts1(1)) + Val(.SubItems(7))   'Value
            End If
        End With
    Next
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If LV.ListItems.count = 0 Then Exit Sub
    If Button <> 2 Then Exit Sub
    Dim strRet As String
    strRet = ShowPopUpMenu1("Hide This Lot from Billing")
    
    If strRet = "Hide This Lot from Billing" Then
        con.Execute "INSERT INTO VendRcvdDetailPO_Hidden_Billing(VRDPO_RefID,UserName,MachineName) VALUES(" & Val(LV.SelectedItem.Tag) & ",'" & CurrentUserName & "','" & strComputerName & "')"
        LV.ListItems.Remove LV.SelectedItem.Index
   End If
   
End Sub

Private Sub ShowMenusForShipping()

    On Error Resume Next
    Dim c As Control
    For Each c In Controls
        If TypeOf c Is Menu Then
            If c.Name <> "mnuPrintPList" Then
                c.Visible = False
            End If
        End If
    Next
    
    On Error GoTo 0
    
    If LV.ListItems.count > 0 Then
        mnuPrintPList.Visible = True
        mnuCustomPList.Visible = True
        mnuCustomPListOther.Visible = True
        mnuCommercialPList.Visible = True
        mnuCommercialPListOther.Visible = True
        mnuDashPackingLabels.Visible = True
        mnuPrintPackingLabels.Visible = True
    End If
'    mnuPrintDecDocs.Visible = False
    Me.PopUpMenu mnuPop
    
End Sub

Private Sub ShowMenus(vis As Boolean)
        
    
    Me.PopUpMenu mnuPop
    
End Sub

Private Sub ShowMenusforCourier()
         
    On Error Resume Next
    Dim c As Control
    For Each c In Controls
        If TypeOf c Is Menu Then
            c.Visible = False
        End If
    Next
    
    On Error GoTo 0
    mnuAdd.Visible = True
    If LV.ListItems.count > 0 Then
        dd.Visible = True
        mnuEdit.Visible = True
        mnuComDocs.Visible = True
        mnuPrintComInv.Visible = True
    End If
    mnuPrintDecDocs.Visible = False
    Me.PopUpMenu mnuPop
    
End Sub


Private Sub mnuAdd_Click()
    'Load frmNewCustomInvoice
    'frmNewCustomInvoice.Show
    'ShowForm frmNewCustomInvoice, MainForm, False, "New Custom Invoice"
    'frmNewCustomInvoice.AddNewProforma
    Load frmNewCustomInvoice
    frmNewCustomInvoice.AddNewInvoice True, bCourierInvoice
End Sub



Private Sub LVManualMaterial_KeyDown(KeyCode As Integer, Shift As Integer)

    If LVManualMaterial.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo) = vbYes Then
            LVManualMaterial.ListItems.Remove LVManualMaterial.SelectedItem.Index
        End If
    End If

End Sub


Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Val(txtEdit) <= 0 Then
            MsgBox "Invalid Rate.", vbInformation
            Exit Sub
        End If
        Dim iSelectedID As Integer
        iSelectedID = Val(LV.SelectedItem.Tag)
        con.Execute "UPDATE VendRcvdDetailPO SET Rate=" & Val(txtEdit) & " WHERE EntryID=" & Val(LV.SelectedItem.Tag)
        txtEdit.Visible = False
        Call RefreshLV
        'Now Try to Re-Select the Row
        Dim i As Integer
        For i = 1 To LV.ListItems.count
            If Val(LV.ListItems(i).Tag) = iSelectedID Then
                LV.ListItems(i).Selected = True
                LV.ListItems(i).EnsureVisible
                LV.SetFocus
            End If
        Next
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub

Private Sub txtManualMatRate_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If txtManualMatDesc = "" Then
            MsgBox "Please Enter Description.", vbInformation
            Exit Sub
'        ElseIf Val(txtManualMatRate) <= 0 Then
'            MsgBox "Invalid Amount.", vbInformation
'            Exit Sub
        End If
        Dim ITM As ListItem
        Set ITM = LVManualMaterial.ListItems.add(, , LVManualMaterial.ListItems.count + 1)
        ITM.ListSubItems.add , , txtManualMatDesc.Text
        ITM.ListSubItems.add , , Val(txtManualMatRate)
    End If
    
    Call getManualDedTotal
End Sub

Private Sub getManualDedTotal()

    Dim lTotal As Long
    For i = 1 To LVManualMaterial.ListItems.count
        lTotal = lTotal + Val(LVManualMaterial.ListItems(i).SubItems(2))
    Next
    txtManualDed = lTotal
    
End Sub

Private Sub getMaterailIssuanceTotal()

    Dim lTotal As Long
    For i = 1 To LVMaterial.ListItems.count
        lTotal = lTotal + Val(LVMaterial.ListItems(i).SubItems(6))
    Next
    txtMaterialDed = lTotal
    
End Sub

Private Sub txtFind_Change()
    cmdFind.Enabled = Len(Trim(txtFind)) > 0
End Sub

Private Sub txtFind_GotFocus()
   txtFind.SelStart = 0
   txtFind.SelLength = Len(txtFind)
   cmdFind.Default = True
End Sub

Private Sub cmdFind_Click()

    Dim StartIdx As Integer
    
    If LV.SelectedItem Is Nothing Then
       StartIdx = 1
    Else
       If LV.SelectedItem.Index = LV.ListItems.count Then
          StartIdx = 1
       Else
          StartIdx = LV.SelectedItem.Index
       End If
    End If

    Dim fldValue As String
    Dim strTemp As String
    For i = StartIdx To LV.ListItems.count
       
        If cmbfield.ListIndex = 0 Then
            fldValue = LV.ListItems(i).Text
        Else
            fldValue = LV.ListItems(i).SubItems(cmbfield.ListIndex)
        End If
    
    
        If LCase(fldValue) Like "*" & LCase(Trim(txtFind)) & "*" Then
            LV.ListItems(i).Selected = True
            LV.ListItems(i).EnsureVisible
            LV.SetFocus
            If i <> StartIdx Then Exit For
        End If
       
    Next

    If i = LV.ListItems.count + 1 Then
        MsgBox "Item Not Found.", vbInformation
        Set LV.SelectedItem = Nothing
    End If

End Sub
