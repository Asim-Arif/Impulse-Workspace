VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmCashPV 
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   8805
   ScaleWidth      =   12135
   WindowState     =   2  'Maximized
   Begin VB.Frame FALL 
      Height          =   8670
      Left            =   105
      TabIndex        =   21
      Top             =   45
      Width           =   11985
      Begin VB.TextBox txtVchrNo 
         Alignment       =   2  'Center
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   195
         TabIndex        =   52
         Top             =   1005
         Width           =   1470
      End
      Begin VB.CommandButton cmdNewAccount 
         Caption         =   "New Account"
         Height          =   255
         Left            =   9105
         TabIndex        =   45
         Top             =   765
         Width           =   1185
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Width           =   11970
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Payment Voucher"
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
            Index           =   8
            Left            =   75
            TabIndex        =   40
            Top             =   165
            Width           =   12525
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Payment Voucher"
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
            Height          =   465
            Index           =   9
            Left            =   90
            TabIndex        =   41
            Top             =   180
            Width           =   12510
         End
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1590
         Left            =   5490
         TabIndex        =   32
         Top             =   3090
         Visible         =   0   'False
         Width           =   4320
      End
      Begin MSComctlLib.ListView LVPymt 
         Height          =   3525
         Left            =   165
         TabIndex        =   36
         Top             =   2460
         Width           =   11760
         _ExtentX        =   20743
         _ExtentY        =   6218
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         PictureAlignment=   5
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   15199215
         BorderStyle     =   1
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
            Text            =   "Voucher #"
            Object.Width           =   3757
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Key             =   "Amt"
            Text            =   "Total Amout"
            Object.Width           =   2081
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amount Due"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Applied Pymt"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Status"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Due Date"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.TextBox Tbox 
         Height          =   315
         Index           =   0
         Left            =   1665
         TabIndex        =   4
         Top             =   1500
         Width           =   5100
      End
      Begin VB.TextBox Tbox 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   8745
         TabIndex        =   6
         Top             =   1515
         Width           =   1485
      End
      Begin VB.Frame FraCash 
         BorderStyle     =   0  'None
         Height          =   5520
         Left            =   165
         TabIndex        =   26
         Top             =   3030
         Width           =   12315
         Begin VB.TextBox Tbox 
            Height          =   315
            Index           =   7
            Left            =   30
            TabIndex        =   12
            Top             =   4200
            Width           =   6045
         End
         Begin VB.CommandButton cmdNewVender 
            Caption         =   "New Vender"
            Height          =   345
            Left            =   2085
            TabIndex        =   54
            Top             =   2970
            Width           =   1185
         End
         Begin VB.CommandButton cmdAddNewMaker 
            Caption         =   "New Maker"
            Height          =   345
            Left            =   3315
            TabIndex        =   53
            Top             =   2970
            Width           =   1185
         End
         Begin ComboList.Usercontrol1 cmbOtherAcc 
            Height          =   285
            Left            =   6405
            TabIndex        =   14
            Top             =   3570
            Visible         =   0   'False
            Width           =   3825
            _ExtentX        =   6747
            _ExtentY        =   556
            BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ListIndex       =   -1
         End
         Begin VB.TextBox Tbox 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   10245
            TabIndex        =   15
            Top             =   3555
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox Tbox 
            Height          =   315
            Index           =   5
            Left            =   30
            TabIndex        =   11
            Top             =   3600
            Width           =   6045
         End
         Begin VB.TextBox Tbox 
            Height          =   675
            Index           =   4
            Left            =   30
            MultiLine       =   -1  'True
            TabIndex        =   13
            Top             =   4770
            Width           =   6045
         End
         Begin VB.ComboBox cmbApprovedBy 
            Height          =   315
            ItemData        =   "frmCashPV.frx":0000
            Left            =   3465
            List            =   "frmCashPV.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   3555
            Visible         =   0   'False
            Width           =   2625
         End
         Begin VB.CheckBox chkPrint 
            Caption         =   "Print When Saved"
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
            Height          =   225
            Left            =   6420
            TabIndex        =   35
            Top             =   4140
            Width           =   1830
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H80000004&
            Height          =   900
            Left            =   6405
            TabIndex        =   33
            Top             =   4530
            Width           =   5295
            Begin MSForms.CommandButton cmdSClear 
               Height          =   480
               Left            =   75
               TabIndex        =   16
               Top             =   255
               Width           =   1665
               Caption         =   "Save & New   "
               PicturePosition =   327683
               Size            =   "2937;847"
               Accelerator     =   78
               FontName        =   "Segoe UI"
               FontEffects     =   1073741825
               FontHeight      =   195
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
               FontWeight      =   700
            End
            Begin MSForms.CommandButton cmdCancel 
               Height          =   480
               Left            =   3525
               TabIndex        =   18
               Top             =   255
               Width           =   1665
               Caption         =   "Cancel            "
               PicturePosition =   327683
               Size            =   "2937;847"
               Accelerator     =   67
               FontName        =   "Segoe UI"
               FontEffects     =   1073741825
               FontHeight      =   195
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
               FontWeight      =   700
            End
            Begin MSForms.CommandButton cmdSClose 
               Height          =   480
               Left            =   1800
               TabIndex        =   17
               Top             =   255
               Width           =   1665
               Caption         =   "Save & Close  "
               PicturePosition =   327683
               Size            =   "2937;847"
               Accelerator     =   83
               FontName        =   "Segoe UI"
               FontEffects     =   1073741825
               FontHeight      =   195
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
               FontWeight      =   700
            End
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Handed over to"
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
            Height          =   195
            Index           =   5
            Left            =   60
            TabIndex        =   55
            Top             =   3975
            Width           =   1335
         End
         Begin VB.Label Lbls 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Amount Applied :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   2
            Left            =   8325
            TabIndex        =   51
            Top             =   3015
            Width           =   1950
         End
         Begin VB.Label Lbls 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Un Applied Amount :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   3
            Left            =   4575
            TabIndex        =   50
            Top             =   3015
            Width           =   1755
         End
         Begin VB.Label lblApplied 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   10245
            TabIndex        =   49
            Top             =   2955
            Width           =   1515
         End
         Begin VB.Label lblToApply 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   6405
            TabIndex        =   48
            Top             =   2955
            Width           =   1515
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Bank Charges Account :"
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
            Height          =   195
            Index           =   10
            Left            =   6420
            TabIndex        =   47
            Top             =   3375
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Amount"
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
            Height          =   195
            Index           =   11
            Left            =   10275
            TabIndex        =   46
            Top             =   3360
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Payee "
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
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   44
            Top             =   3375
            Width           =   585
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Notes"
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
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   43
            Top             =   4575
            Width           =   495
         End
         Begin VB.Label Lbls 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Approved"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   0
            Left            =   3480
            TabIndex        =   37
            Top             =   3300
            Visible         =   0   'False
            Width           =   825
         End
         Begin MSForms.CommandButton cmd 
            Height          =   360
            Left            =   15
            TabIndex        =   10
            Top             =   2970
            Width           =   1740
            VariousPropertyBits=   25
            Caption         =   "Auto Apply         "
            PicturePosition =   327683
            Size            =   "3069;635"
            Accelerator     =   80
            FontName        =   "Tahoma"
            FontEffects     =   1073750017
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.Frame FraBank 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   180
         TabIndex        =   22
         Top             =   1830
         Width           =   11730
         Begin VB.CheckBox ChkClear 
            Caption         =   "Clear Cheque "
            Height          =   225
            Left            =   10305
            TabIndex        =   34
            ToolTipText     =   "Clear This Cheque Automatically"
            Top             =   345
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.TextBox Tbox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   5430
            TabIndex        =   7
            Top             =   285
            Width           =   1725
         End
         Begin VB.CommandButton cmdBank 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5040
            Picture         =   "frmCashPV.frx":0004
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   285
            Width           =   360
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   8850
            TabIndex        =   9
            Top             =   270
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   156106755
            CurrentDate     =   37250
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   0
            Left            =   5475
            TabIndex        =   42
            Top             =   75
            Width           =   945
         End
         Begin MSForms.ComboBox cmbCheckType 
            Height          =   315
            Left            =   7200
            TabIndex        =   8
            Top             =   285
            Width           =   1650
            VariousPropertyBits=   679495707
            DisplayStyle    =   7
            Size            =   "2910;556"
            ListWidth       =   2822
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label2 
            Height          =   255
            Index           =   1
            Left            =   15
            TabIndex        =   24
            Top             =   30
            Width           =   10620
            ForeColor       =   4194304
            BackColor       =   12504280
            VariousPropertyBits=   8388627
            Caption         =   $"frmCashPV.frx":0376
            Size            =   "18732;450"
            BorderColor     =   14937593
            BorderStyle     =   1
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label lblBankAcc 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   15
            TabIndex        =   23
            Top             =   300
            Width           =   5010
         End
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   1020
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dddd, dd/MM/yyyy"
         Format          =   156106755
         CurrentDate     =   37250
      End
      Begin VB.TextBox Tbox 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   6795
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1515
         Width           =   1920
      End
      Begin MSForms.ComboBox CmbAccTitle 
         Height          =   315
         Left            =   5760
         TabIndex        =   2
         Top             =   1005
         Width           =   4545
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "8017;556"
         ListWidth       =   8819
         ListStyle       =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   10335
         TabIndex        =   27
         Top             =   1005
         Width           =   1500
      End
      Begin VB.Label lblBal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   10260
         TabIndex        =   19
         Top             =   1515
         Width           =   1620
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vocuher # "
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
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   31
         Top             =   810
         Width           =   945
      End
      Begin VB.Label lblBalCap 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash in Hand  "
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
         Height          =   195
         Left            =   10290
         TabIndex        =   29
         Top             =   1320
         Width           =   1185
      End
      Begin MSForms.ComboBox cmbAccNo 
         Height          =   315
         Left            =   3840
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1005
         Width           =   1920
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3387;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbPymtmode 
         Height          =   315
         Left            =   195
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1500
         Width           =   1470
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "2593;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   2
         Left            =   195
         TabIndex        =   30
         Top             =   795
         Width           =   11655
         ForeColor       =   4194304
         BackColor       =   12504280
         VariousPropertyBits=   8388627
         Caption         =   $"frmCashPV.frx":042A
         Size            =   "20558;397"
         BorderColor     =   14937593
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   4
         Left            =   2580
         TabIndex        =   28
         Top             =   645
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   0
         Left            =   195
         TabIndex        =   25
         Top             =   1305
         Width           =   11700
         ForeColor       =   4194304
         BackColor       =   12504280
         VariousPropertyBits=   8388627
         Caption         =   $"frmCashPV.frx":0508
         Size            =   "20637;397"
         BorderColor     =   14937593
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
End
Attribute VB_Name = "frmCashPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dontShowList As Boolean
Dim PermissionDenied As Boolean
Dim bPosting As Boolean


Private Sub CmbAccTitle_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    If KeyCode = vbKeyDown Then
       If List1.ListIndex < List1.ListCount - 1 Then List1.ListIndex = List1.ListIndex + 1
       KeyCode = 0
       dontShowList = True
    ElseIf KeyCode = vbKeyUp Then
       If List1.ListIndex > 0 Then List1.ListIndex = List1.ListIndex - 1
       KeyCode = 0
       dontShowList = True
    ElseIf KeyCode = 13 And List1.Visible Then
       CmbAccTitle = List1
       List1.Visible = False
       CmbAccTitle.SetFocus
    End If

End Sub

Private Sub CmbAccTitle_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If dontShowList Then dontShowList = False Else ShowList List1, CmbAccTitle, CmbAccTitle
End Sub

Private Sub CmbAccTitle_LostFocus()
    List1.Visible = False
End Sub

Private Sub cmdAddNewMaker_Click()

    Dim f As New frmMaker
    With f
        .add = True
        .Show 1
    End With
    Call FillAccCmbs

End Sub

Private Sub cmdNewAccount_Click()

    RecentAccNo = ""
    frmNewAccount.Show 1
    
    If RecentAccNo <> "" Then
        Call FillAccCmbs
        cmbAccNo = RecentAccNo
        CmbAccTitle.SetFocus
    End If

End Sub

Private Sub cmdNewVender_Click()

    Dim f As New frmVender
    With f
        .add = True
        .Show 1
    End With
    Call FillAccCmbs

End Sub

Private Sub DTPicker2_Change()
   Call setDateRange(DTPicker2)
   'ChkClear.Enabled = (DTPicker1 = DTPicker2)
End Sub

Private Sub Form_Activate()

    'Check For User Access
    If PermissionDenied Then
       Me.Hide
       MsgBox "You Donot Have Permissions To Use This Option. Please Contact Your Administrator", vbInformation
       Unload Me
       Exit Sub
    End If


End Sub

Private Sub List1_Click()
    If UCase(Screen.ActiveControl.Name) <> "CMBACCTITLE" Then
       CmbAccTitle = List1
       List1.Visible = False
       CmbAccTitle.SetFocus
    End If
End Sub
Private Sub cmbACCNo_Change()
    If cmbAccNo.MatchFound Then Call cmbACCNo_Click Else If Not cmbAccNo = "" Then cmbAccNo.DropDown
End Sub
Private Sub cmbACCNo_Click()
    If CmbAccTitle.ListIndex <> cmbAccNo.ListIndex Then CmbAccTitle.ListIndex = cmbAccNo.ListIndex
    lblBalance.Visible = cmbAccNo.MatchFound
    If cmbAccNo Like "24*" Then Call RefreshList Else Call RefreshList(True)
End Sub

Private Sub cmbACCTitle_Change()
    If CmbAccTitle.MatchFound Then Call cmbACCTitle_Click 'Else If CmbAccTitle <> "" Then CmbAccTitle.DropDown
End Sub

Private Sub cmbACCTitle_Click()
    If cmbAccNo.ListIndex <> CmbAccTitle.ListIndex Then cmbAccNo.ListIndex = CmbAccTitle.ListIndex
    If cmbPymtmode.ListIndex = 0 Then
'        TBox(0).Text = "Cash Paid to " & CmbAccTitle.Text
    Else
        'Tbox(0).Text = "Cheque Paid to " & CmbAccTitle.Text
'        TBox(0).Text = "Payment Through " & cmbPymtmode.Text & " To " & CmbAccTitle '&  "  " & Label1(0) & " : " & Tbox(2)
    End If
    TBox(5) = CmbAccTitle
End Sub

Private Sub cmbAccNo_Validate(cancel As Boolean)
    If Not cmbAccNo.MatchFound Then cmbAccNo.ListIndex = CmbAccTitle.ListIndex
End Sub

Private Sub CmbAccTitle_Validate(cancel As Boolean)
   If Not CmbAccTitle.MatchFound Then CmbAccTitle.ListIndex = cmbAccNo.ListIndex
End Sub

Private Sub CmbAccTitle_DropButtonClick()
   List1.Visible = False
End Sub


Private Sub cmbPymtmode_Change()
    Call cmbPymtmode_Click
    Call DTPicker1_Change
End Sub

Private Sub cmbPymtmode_Click()


    If cmbPymtmode.ListIndex = 0 Then
        FraBank.Visible = False
        Label1(3).Visible = True
        'Label1(4).Visible = True
        chkPrint.Visible = True
        If CmbAccTitle.MatchFound = True Then
            TBox(0) = "Cash Paid to " & CmbAccTitle.Text
        End If
        LVPymt.Move LVPymt.Left, 1890, LVPymt.Width, 4100
        cmbOtherAcc.Visible = False
        Label1(10).Visible = False
        TBox(6).Visible = False
        Label1(11).Visible = False
        
    Else
        
        'Get Default BankID
        Dim BankID As Double
        lblBankAcc = GetDefaultBankAcc(BankID)
        lblBankAcc.Tag = BankID
        
        If cmbPymtmode.ListIndex = 1 Then
            If Not AnyBankAccountDefined Then
                MsgBox "Payment Cannot Be Made By Cheque, No Bank Accounts Defined.", vbInformation
                 cmbPymtmode.ListIndex = 0
                Exit Sub
            End If
            
            If CmbAccTitle.MatchFound = True Then
                'TBox(0) = "Cheque Paid to " & CmbAccTitle.Text
                TBox(0) = "Cheque Paid"
            End If
                
            cmbCheckType.ListIndex = 0
            
            'Label1(4).Visible = False

             'Get Next Chq #
            'If strCompany <> "Banzai" Then
            Dim ChqBookNo As Double
            TBox(2) = GetNextChqNo(GetID(lblBankAcc), ChqBookNo)
            TBox(2).Tag = ChqBookNo
            'End If
            DTPicker2 = DTPicker1
            cmbCheckType.Enabled = True
            ChkClear.Enabled = True
            DTPicker2.Enabled = True
        Else
                
            TBox(2) = ""
            TBox(2).Tag = 0
                 
            cmbCheckType.Enabled = False
            ChkClear.Enabled = False
            DTPicker2.Enabled = False
                
            cmbOtherAcc.Visible = True
            Label1(10).Visible = True
            TBox(6).Visible = True
            Label1(11).Visible = True
                 
            If cmbPymtmode.ListIndex = 2 Then
                Label1(0).Caption = "Draft No."
            ElseIf cmbPymtmode.ListIndex = 3 Then
                Label1(0).Caption = "Reciept No."
            End If
            
            
        End If
        
        'chkPrint.Visible = False
        LVPymt.Move LVPymt.Left, 2460, LVPymt.Width, 3500
        FraBank.Visible = True
        Label1(3).Visible = False
        
    End If
    Call ShowBals
    Call Form_Resize
    
End Sub

Public Sub cmd_Click()

With LVPymt
    If Trim(cmd.Caption) = "Auto Apply" Then
        For i = 1 To .ListItems.count
            
            .ListItems(i).Checked = True
            Call LVPymt_ItemCheck(.ListItems(i))
            If Val(lblToApply) = 0 Then Exit For
        Next
        If i > 1 Then cmd.Caption = "Clear Payments     "
    Else
        For i = 1 To .ListItems.count
            .ListItems(i).Checked = False
            Call LVPymt_ItemCheck(.ListItems(i))
        Next
        lblToApply = TBox(1)
        lblApplied = 0
        cmd.Caption = "Auto Apply         "
    End If
End With
End Sub

Private Sub cmdBank_Click()

    Dim BankID As Double
    lblBankAcc = SelectBankAcc(BankID)
    lblBankAcc.Tag = BankID
    
    'Get ChqNo
         
    Dim ChqBookNo As Double
    TBox(2) = GetNextChqNo(GetID(lblBankAcc), ChqBookNo)
    TBox(2).Tag = ChqBookNo
        
End Sub

Private Sub cmdSClear_Click()
    
    On Error GoTo err
    Dim bError As Boolean
    bError = False
    If bPosting Then Exit Sub
    bPosting = True
        
    If Val(TBox(1)) > Val(lblBal) Then
        If Not AllowOD(GetID(lblBankAcc)) Then
            MsgBox "Not Much Amount In " & Left(lblBalCap, Len(lblBalCap) - 2) & ".", vbInformation
            Exit Sub
        End If
    End If
    'If Not AskForAutoApply Then bError = True
    Call AskForAutoApply
'    If cmbApprovedBy.ListIndex = -1 Then
'        MsgBox "Please Select Voucher Approving Authority.", vbInformation
'        Exit Sub
'    End If
    
    If cmbPymtmode.ListIndex = 0 Then
        If Not SaveTransForCash Then bError = True
    Else
        If Not SaveTransForCheque Then bError = True
    End If
    
    bPosting = False
    If bError Then
        Exit Sub
    End If
    
    Dim myDT As Date
    myDT = DTPicker1.Value
    Unload Me
    Me.Show
    DTPicker1.Value = myDT
    Call DTPicker1_Change
    Me.SetFocus
   
    'Call ClearFields
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdSClose_Click()
        
    On Error GoTo err
    Dim bError As Boolean
    bError = False
    If bPosting Then Exit Sub
    
    Dim dODAmt As Double
    If Val(TBox(1)) > Val(lblBal) Then
        If Not AllowOD(GetID(lblBankAcc)) Then
            MsgBox "Not Much Amount In " & Left(lblBalCap, Len(lblBalCap) - 2) & ".", vbInformation
            Exit Sub
        End If
'        dODAmt = AllowODAmt(GetID(lblBankAcc))
'        If Val(Tbox(1)) Then
'            If Abs(lblBal) + Val(Tbox(1)) > dODAmt Then
'                MsgBox "This Amount Is Not Allowed.", vbInformation
'                Exit Sub
'            End If
'        End If
    End If
    
    bPosting = True
    
    If Not AskForAutoApply Then bError = True
     
    If cmbPymtmode.ListIndex = 0 Then
        If Not SaveTransForCash Then bError = True
    Else
        If Not SaveTransForCheque Then bError = True
    End If
    
    bPosting = False
    
    If bError Then
        Exit Sub
    End If
    
    Unload Me
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub DTPicker1_Change()
    
    Call setDateRange(DTPicker1)
    
    txtVchrNo = getNextVchrNo(DTPicker1.Value, IIf(cmbPymtmode.ListIndex = 0, "CPV", "BPV"))
    If cmbAccNo.MatchFound Then Call RefreshList
    
    DTPicker2 = DTPicker1
    
End Sub

Private Sub DTPicker1_Click()
    Call DTPicker1_Change
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    
    'PermissionDenied = Not UserHasAccess("PaymentVoucher")
    bPosting = False
    Me.KeyPreview = True
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    Call DTPicker1_Click
    Call FillAccCmbs
    
'    If CurrentUserName = "Administrator" Then
'        DTPicker1.Enabled = True
'    Else
'        DTPicker1.Enabled = False
'    End If
   DTPicker1.Enabled = True
   chkPrint = Abs(bPrintWhenSaved)
   'chkPrint = IIf(MainForm.mnuPrintOnSave.Checked, vbChecked, vbUnchecked)
   
    
   ' CashInHand = getAccountNo("Cash In Hand")
    
   ' If CashInHand = "" Then
   '     MsgBox "No Cash in Hand Account is defined", vbInformation
   '     Exit Sub
   ' End If
    
End Sub

Private Sub FillAccCmbs()

    
        'Clear All The Combos
        cmbAccNo.Clear
        CmbAccTitle.Clear
        cmbPymtmode.Clear
'        cmbBank.Clear
'        cmbBankBranch.Clear
        cmbCheckType.Clear
        
        'Fill With CheckType
        cmbCheckType.AddItem "Open"
        cmbCheckType.AddItem "Payees A/C Only"
        cmbCheckType.ListIndex = 0
        
        
        'Fill With Payment Mode
        cmbPymtmode.AddItem "Cash"
        cmbPymtmode.AddItem "Cheque"
        cmbPymtmode.AddItem "D.D."
        cmbPymtmode.AddItem "Online"
        cmbPymtmode.ListIndex = 1
        cmbPymtmode.Enabled = False
'        Call AddToCombo(cmbAccNo, "AccNo", "VActiveAccounts", "Where AccNo like '24%' and Parent=0 Order By AccNo")
'        Call AddToCombo(CmbAccTitle, "AccTitle", "VActiveAccounts", "Where AccNo like '24%' and Parent=0 Order By AccNo")
        
        Call AddToCombo(cmbAccNo, "AccNo", "VActiveAccounts", " Order By AccNo")
        Call AddToCombo(CmbAccTitle, "AccTitle", "VActiveAccounts", " Order By AccNo")
        
        'Call AddToCombo(cmbBank, "Distinct Bank", "BankList", "Order By Bank")

    With cmbOtherAcc
        .ListHeight = 1000
        .ClearVals
        .AddVals con, "AccNo + '{' + Acctitle + '}'", "VActiveAccounts", "AccNo"
    End With
    
End Sub


Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)

'If UnloadMode = 0 Then
'    If LVPymt.ListItems.Count > 0 Then
'        If MsgBox("There is a Transaction active." & vbNewLine & "Save or Cancel Transaction first.", vbInformation + vbOKOnly) Then
'            Cancel = True
'        End If
'    End If
'End If

End Sub










Private Sub Form_Resize()
   
   FALL.Move (Me.ScaleWidth - FALL.Width) / 2, (Me.ScaleHeight - FALL.Height) / 2
   
End Sub

Private Sub lblBankAcc_Change()
    Call ShowBals
End Sub

Private Sub Tbox_Change(Index As Integer)
'    If Index = 1 Then lblToApply = Tbox(1)
If Index = 0 Then TBox(4) = TBox(0)
End Sub

Private Sub TBox_GotFocus(Index As Integer)
With TBox(Index)
    .SelStart = 0
    .SelLength = Len(TBox(Index))
End With
End Sub

Private Sub Tbox_KeyPress(Index As Integer, KeyAscii As Integer)

If Index = 1 Or Index = 6 Then
    KeyAscii = OnlyNumber(KeyAscii, True)
End If

End Sub
Private Sub ClearLView()

    Dim iTotal As Integer
    iTotal = LView.ListItems.count
    For i = 1 To iTotal
        LView.ListItems.Remove (1)
    Next i
    
TBox(0).Text = ""
TBox(1).Text = ""
cmdDelete.Visible = False
Call DTPicker1_Click

End Sub


Private Sub LVPymt_ItemCheck(ByVal Item As MSComctlLib.ListItem)

If Not Item.Checked Then
    lblApplied = Format(Val(lblApplied) - Val(Item.SubItems(4)), "0.00")
    lblToApply = Format(Val(lblToApply) + Val(Item.SubItems(4)), "0.00")
    Item.SubItems(4) = 0
    Item.SubItems(5) = "DUE"
Else
 
    If Val(lblToApply) = 0 Then
        MsgBox "No Amount To Apply", vbInformation
        Item.Checked = False
        
        Exit Sub
    End If
 
    If Item.SubItems(3) <= Val(lblToApply) Then
        Item.SubItems(4) = Format(Val(Item.SubItems(3)), "0.00")
        Item.SubItems(5) = "FULL PAID"
   
    Else
        Item.SubItems(4) = Format(Val(lblToApply), "0.00")
        Item.SubItems(5) = "PARTIAL PAID"
    End If
  
    lblApplied = Format(Val(lblApplied) + Val(Item.SubItems(4)), "0.00")
    lblToApply = Format(Val(lblToApply) - Val(Item.SubItems(4)), "0.00")
  
End If

   If lblToApply = TBox(1) Then
      cmd.Caption = "Auto Apply      "
   Else
      cmd.Caption = "Clear Payments    "
   End If

'If Val(lblapply) = 0 Then
' cmdSave.Enabled = True
'Else
' cmdSave.Enabled = False
'End If

End Sub

Private Sub RefreshList(Optional DontShowBills As Boolean = False)

    Screen.MousePointer = vbHourglass
    
    Call ShowAccBalance
    
    If DontShowBills Then GoTo AfterBills
    
    Dim rsList As New ADODB.Recordset
 
    rsList.CursorLocation = adUseClient
    rsList.Open "SELECT TOP 50 Sno,VchrNo,Description,TotalAmount,AmountLeft,DueDate FROM VAccPay WHERE (AccNo='" & cmbAccNo & "') AND (VDate<='" & DTPicker1 & "') AND ((AmountLeft>0) or (AmountLeft is null)) AND VAccPay.VchrNo NOT IN(SELECT VchrNo FROM VMakerPostedBillsList WHERE ISNULL(TotalPaidAmount,0)>0) ORDER BY AccNo", con, adOpenForwardOnly, adLockReadOnly
    rsList.Sort = "Sno"
 
    With LVPymt.ListItems
        .Clear

        cmd.Caption = "Auto Apply         "
        Dim ITM As ListItem
        Do Until rsList.EOF
            Set ITM = .add(, rsList![SNo] & "?", rsList![VchrNo])
            ITM.SubItems(1) = rsList![Description] & ""
            ITM.SubItems(2) = rsList![TotalAmount] & ""
            If IsNull(rsList![AmountLeft]) Then
                ITM.SubItems(3) = rsList![TotalAmount]
            Else
                ITM.SubItems(3) = rsList![AmountLeft] & ""
            End If
            
            ITM.SubItems(4) = 0
            ITM.SubItems(5) = "DUE"
            ITM.SubItems(6) = Format(rsList!DueDate, "dd-MMM-yyyy")
            rsList.MoveNEXT
        Loop
        If .count > 0 Then cmd.Enabled = True Else cmd.Enabled = False
    End With


AfterBills:
    'Reset The Amount
    Call TBox_Validate(1, False)
    Call ShowBals
    Screen.MousePointer = vbDefault
    
End Sub

Private Function SaveTransForCash() As Boolean

    If CheckData = False Then Exit Function
    
    Dim i As Integer
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim Bal As Double
    Dim RcptNo As String
    
    Dim NetAmount As Double
    
    If TBox(6).Visible And Val(TBox(6)) > 0 Then
        NetAmount = Val(TBox(1).Text) + Val(TBox(6))
    Else
        NetAmount = Val(TBox(1).Text)
    End If
    
    
    'Get Serial # and Balances
    INextSno = getNextSno(DTPicker1.Value)
    
    Dim bWaitForAuth As Boolean
    bWaitForAuth = VoucherRequiresAuthorization("CPV")
    
    RcptNo = IIf(Len(Trim(TBox(3))) = 0, "", "Rcpt # :" & TBox(3))
    If bWaitForAuth = False Then
        Bal = getBalance(CashInHand, DTPicker1.Value)
        Bal = Bal - Val(NetAmount)
    End If
    
    'Start New Transaction
    Call StartTrans(con)
    'Credit Entry
    If bWaitForAuth = False Then
        con.Execute "INSERT INTO Vouchers VALUES(" & INextSno & ",'" & _
         Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & Label1(4) & "','" & CashInHand & _
         "','" & TBox(0) & RcptNo & "',0," & Val(NetAmount) & "," & Bal & ",'',0)"
        
        Call adjustLedger(CashInHand, DTPicker1.Value, INextSno, -Val(NetAmount))
    
        INextSno = INextSno + 1
        SnoToUse = INextSno
        Bal = getBalance(cmbAccNo, DTPicker1.Value)
        Bal = Bal + Val(TBox(1))
    
        'Debit Entry
        con.Execute "INSERT INTO Vouchers Values(" & INextSno & ",'" & _
         Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & Label1(4) & "','" & cmbAccNo & _
         "','" & TBox(0) & RcptNo & "'," & Val(TBox(1)) & _
         ",0," & Bal & ",''," & Val(CSNo) & ")"
        
        Call adjustLedger(cmbAccNo, DTPicker1.Value, INextSno, Val(TBox(1)))
    
        With LVPymt.ListItems
            For i = 1 To .count
                If .Item(i).Checked Then con.Execute "INSERT INTO VoucherDetail Values('" & Val(.Item(i).key) & "','" & SnoToUse & "'," & Val(LVPymt.ListItems(i).SubItems(4)) & ")"
            Next
        End With
    
        'Debit Entry (Bank Charges)
    
        If TBox(6).Visible And Val(TBox(6)) > 0 Then
            INextSno = INextSno + 1

            Bal = getBalance(cmbOtherAcc.ID, DTPicker1.Value)
            Bal = Bal + Val(TBox(6))
        
            con.Execute "INSERT INTO Vouchers Values(" & INextSno & ",'" & _
             Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & Label1(4) & "','" & cmbOtherAcc.ID & _
             "','" & TBox(0) & RcptNo & "'," & Val(TBox(6)) & _
             ",0," & Bal & ",'',0)"
             
            Call adjustLedger(cmbOtherAcc.ID, DTPicker1.Value, INextSno, Val(TBox(6)))
        
        End If
    
        con.Execute "INSERT INTO VoucherInfo(VchrNo,UserName,MachineName,DT,ApprovedBy,Notes,Payee,Handed_Over_To) VALUES('" & Label1(4) & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "','" & cmbApprovedBy.Text & "','" & TBox(4) & "','" & TBox(5) & "','" & TBox(7) & "')"
    Else
        con.Execute "INSERT INTO VouchersForAuthorization(VDate,VchrNo,Accno,Description,Debit,Credit,balance,DpstSlip,CSNo) VALUES('" & Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & _
             Label1(4) & "','" & CashInHand & "','" & TBox(0) & RcptNo & "',0," & Val(TBox(1)) & ",0,'',0)"
        
        con.Execute "INSERT INTO VouchersForAuthorization(VDate,VchrNo,Accno,Description,Debit,Credit,balance,DpstSlip,CSNo) VALUES('" & Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & _
             Label1(4) & "','" & cmbAccNo & "','" & TBox(0) & RcptNo & "'," & Val(TBox(1)) & ",0,0,''," & Val(CSNo) & ")"
             
        con.Execute "INSERT INTO VouchersForAuthorizationInfo(VchrNo,UserName,MachineName,DT,PostedThroughJVForm,Notes,Payee) VALUES('" & Label1(4) & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "',0,'" & TBox(4) & "','" & TBox(5) & "')"
    End If
    
    'Commit Transaction
    con.CommitTrans
    SaveTransForCash = True
    
    Call PlaySaveSound
    
    'Print If To Be Printed
    
    If chkPrint = vbChecked Then Call PrintVoucher(MainForm.Cr1, Label1(4), True, Not bVoucherViewOnlyWhenSaved, bWaitForAuth)
      
    MsgBox "Transaction Saved Successfuly.", vbInformation
    
    'FillAccCmbs
    
    Exit Function
err:
    con.RollbackTrans
    MsgBox "Transaction Terminated Due to ERROR:" & err.Number & vbNewLine & err.Description, vbCritical
    SaveTransForCash = False
    
End Function


Private Function CheckData() As Boolean

On Error GoTo err

'If LVPymt.ListItems.Count = 0 Then
'    MsgBox "No Transaction to save!", vbInformation
'    Exit Function
'End If
    
If cmd.Caption = "Auto A&pply" And cmd.Enabled Then
    If MsgBox("No Amount is Applied. Do You Want To Apply Automatically.", vbQuestion + vbYesNo) = vbYes Then
        Call cmd_Click
    Else
        Exit Function
    End If
End If

If (cmbAccNo = "") Or (Not cmbAccNo.MatchFound) Then
    MsgBox "Invalid Account Selected...", vbInformation
    cmbAccNo.SetFocus
    Exit Function
ElseIf TBox(0) = "" Then
    MsgBox "No Description Given...", vbInformation
    TBox(0).SetFocus
    Exit Function
ElseIf Val(TBox(1) & "") = 0 Then
    MsgBox "No Amount Given...", vbInformation
    TBox(0).SetFocus
    Exit Function

End If


Dim Issued As Boolean, ChqBookNo As Double
If cmbPymtmode.ListIndex = 1 Then
    If lblBankAcc = "" Then
      MsgBox "Invalid Bank Selected.", vbInformation
      Exit Function
    ElseIf Not ValidChqNo(GetID(lblBankAcc), TBox(2), False, ChqBookNo) Then
      If Issued Then
         MsgBox "The Cheque With Number '" & txtChqNo & "' Has Been Issued.", vbInformation
      Else
         MsgBox "Invalid Cheque No.", vbInformation
      End If
      TBox(2).SetFocus
      Exit Function
    End If
    TBox(2).Tag = ChqBookNo
End If

''''''''''''''''''''''Check For Voucher No. Used Or Not '''''''''''''''''
    If VoucherExists(txtVchrNo.Text) Then
        Dim VNo As String
        VNo = getNextVchrNo(DTPicker1.Value, IIf(cmbPymtmode.ListIndex = 0, "CPV", "BPV"))
        MsgBox "Voucher No. : " & Label1(4) & " Has Been Used By Other Transaction." & vbNewLine & _
           "Voucher No. : " & VNo & " Will Be Used For This Transaction.", vbInformation
           txtVchrNo.Text = VNo
    End If
'    Dim VNo As String
'    VNo = getNextVchrNo(DTPicker1.value, IIf(cmbPymtmode.ListIndex = 0, "CPV", "BPV"))
'    If Label1(4) <> VNo Then
'       MsgBox "Voucher No. : " & Label1(4) & " Has Been Used By Other Transaction." & vbNewLine & _
'           "Voucher No. : " & VNo & " Will Be Used For This Transaction.", vbInformation
'           Label1(4) = VNo
'    End If



CheckData = True

Exit Function

err:
    MsgBox err.Description
    
End Function

Private Function SaveTransForCheque() As Boolean

    On Error GoTo err
    If CheckData = False Then Exit Function


    Dim i As Integer
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim Bal As Double
    Dim RcptNo As String
    Dim ChqType As String
    
    If cmbPymtmode.ListIndex > 0 Then
        RcptNo = IIf(Len(Trim(TBox(3))) = 0, "", "Rcpt # :" & TBox(3))
        
        'Get Serial # and Balances
        INextSno = getNextSno(DTPicker1.Value, "Cheque")
        SnoToUse = INextSno
        
        'Start New Transaction
        Call StartTrans(con)
        
        'bankDesc = Right(lblBankAcc, Len(lblBankAcc) - InStr(lblBankAcc, "]")) & " Chq # :" & Tbox(2)
        
        
        If cmbPymtmode.ListIndex = 1 Then
            ChqType = cmbCheckType
            bankDesc = "Cheque No.:" & TBox(2) & " " & TBox(0)
        Else
            ChqType = cmbPymtmode
            bankDesc = TBox(0)
        End If
                
        con.Execute "Insert Into Cheque ( SNo, CDate, AccNo, Description, Amount, BankID, chequeno, ChqBookNo, chequeType, chequeDate, Posted, Payment, Bounced,Payee,BankAccNo) Values(" & INextSno & ",'" & Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & _
        cmbAccNo & "','" & TBox(0) & RcptNo & " (" & bankDesc & ") '," & TBox(1) & "," & lblBankAcc.Tag & ",'" & _
        TBox(2) & "'," & TBox(2).Tag & ",'" & ChqType & "','" & DTPicker2 & "',0,1,0,'" & TBox(5) & "','" & GetID(lblBankAcc) & "')"
        
          Dim dChqSNo As Double
        dChqSNo = INextSno
        
        'Remove the Cheque From Cheque book
        If cmbPymtmode.ListIndex = 1 Then Call IssueChq(Val(TBox(2).Tag), TBox(2))
        
        With LVPymt.ListItems
            For i = 1 To .count
                If .Item(i).Checked Then con.Execute "Insert into ChequeDetail Values('" & SnoToUse & "','" & Val(.Item(i).key) & "'," & Val(LVPymt.ListItems(i).SubItems(4)) & ")"
            Next
        End With
    
    'Commit Transaction
    con.CommitTrans
    SaveTransForCheque = True
    
'    If ChkClear.Enabled And ChkClear = vbChecked Then
'        If Not SavePaymentTrans(SnoToUse) Then GoTo err:
'    ElseIf chkPrint = vbChecked Then
'        Call PrintChq(MainForm.cr1, SnoToUse)
'    End If
    End If
    
    If Not SavePaymentTrans(SnoToUse) Then GoTo err:
    
    
    If cmbCheckType.ListIndex = 0 Then  'Open Cheque
        If Left(cmbAccNo, Len(CashInHand)) <> CashInHand Then
'            Call AddWHTaxOnCashWithdrawalVoucher
        End If
    End If
    
'    If Not (strCompany = "Weldon Instruments" Or strCompany = "Sunlike") Then
'        If cmbCheckType.ListIndex = 0 And Val(Tbox(1)) >= 25000 Then
'            Call AddTwoPercentVoucher
'        End If
'    End If
    SaveTransForCheque = True
'    If cmbCheckType.ListIndex = 0 And Val(Tbox(1)) >= 25000 Then
'        Call AddTwoPercentVoucher
'    End If
    
    Call PlaySaveSound
    MsgBox "Transaction Saved Successfuly.", vbInformation
   
   
'   If MsgBox("Transaction Saved Successfuly." & vbNewLine & "Do you want to Print Cheque?", vbYesNo + vbQuestion) = vbYes Then
'        Dim f As New frmPrintCheque
'        f.ShowMe (dChqSNo)
'    End If
    Exit Function
err:
    MsgBox "Transaction Terminated Due to ERROR:" & err.Number & vbNewLine & err.Description, vbCritical
    SaveTransForCheque = False
    con.RollbackTrans
End Function

Private Function AddTwoPercentVoucher() As String
    
    'Store Null To Return on an Unsuccessful Entry
    Dim VNo As String, m_DTVchr As Date
    m_DTVchr = DTPicker2.Value
    VNo = getNextVchrNo(m_DTVchr, "JV")
    AddTwoPercentVoucher = ""
    Dim i As Integer
    Dim dAmt As Double, strBills As String, Entries As String, strPOs As String, dDiscount As Double, dFreight As Double
    Dim lRefID As Long
    Dim strTaxAcc As String
    Dim dNetAmt As Double
    dNetAmt = (Val(TBox(1)) * 0.3 / 100)
    
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim Bal As Double, Desc As String
    
    INextSno = getNextSno(m_DTVchr)
    'Remember Sno For Voucher Details
    SnoToUse = INextSno
    
    Desc = "W/H Tax On Cash Chq # " & TBox(2)
    
    'Start New Transaction
    'Call StartTrans(con)
      
    'Debit Entry....        '<<<<Change Order No.>>>> & Description etc.
    strTaxAcc = "45-23006"
    If dNetAmt > 0 Then
        Bal = getBalance(strTaxAcc, m_DTVchr)
                
        Bal = Bal + dNetAmt
        con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & m_DTVchr & "','" & _
        VNo & "','" & strTaxAcc & "','" & Desc & "'," & _
        dNetAmt & ",0," & Bal & ",'',0)"
        
        Call adjustLedger(strTaxAcc, m_DTVchr, INextSno, dNetAmt)
        
    End If
    Dim strBankAccNo As String
    strBankAccNo = GetID(lblBankAcc)
    INextSno = INextSno + 1
    'Credit Entry
    Bal = getBalance(strBankAccNo, m_DTVchr)
    Bal = Bal - dNetAmt
   
    con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
         Format(m_DTVchr, "dd-MMM-yyyy") & "','" & VNo & "','" & strBankAccNo & _
        "','" & Desc & "',0," & dNetAmt & "," & Bal & ",'',0)"
          
    Call adjustLedger(strBankAccNo, m_DTVchr, INextSno, -dNetAmt)
    
    con.Execute "Insert Into VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & VNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
    
    
    '''Second Voucher
    'Debit Entry....        '<<<<Change Order No.>>>> & Description etc.
    'Debit Entry
    VNo = getNextVchrNo(m_DTVchr, "JV")
    Dim strPartyAccNo As String
    strPartyAccNo = cmbAccNo.Text
    If dNetAmt > 0 Then
        INextSno = INextSno + 1
        Bal = getBalance(strPartyAccNo, m_DTVchr)
        Bal = Bal + dNetAmt
        con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & m_DTVchr & "','" & _
        VNo & "','" & strPartyAccNo & "','" & Desc & "'," & _
        dNetAmt & ",0," & Bal & ",'',0)"
        
        Call adjustLedger(strPartyAccNo, m_DTVchr, INextSno, dNetAmt)
        
    End If
    
    INextSno = INextSno + 1
    'Credit Entry
    Bal = getBalance(strTaxAcc, m_DTVchr)
    Bal = Bal - dNetAmt
   
    con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
         Format(m_DTVchr, "dd-MMM-yyyy") & "','" & VNo & "','" & strTaxAcc & _
        "','" & Desc & "',0," & dNetAmt & "," & Bal & ",'',0)"
          
    Call adjustLedger(strTaxAcc, m_DTVchr, INextSno, -dNetAmt)
    
    con.Execute "Insert Into VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & VNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
    
    AddTwoPercentVoucher = VNo
    
    Exit Function
err:
    AddTwoPercentVoucher = False
    MsgBox "Error In AddVoucher :" & err.Number & vbNewLine & err.Description
End Function

Private Sub ClearFields()
    Label1(3).Visible = True
    Label1(4).Visible = True
    Call FillAccCmbs
    cmbPymtmode.ListIndex = 0
    TBox(0) = "Cash Recieved"
    TBox(1) = ""
    TBox(2) = ""
End Sub


Private Function AskForAutoApply() As Boolean

    AskForAutoApply = False
    


    If Trim(cmd.Caption) = "Auto Apply" And LVPymt.ListItems.count > 0 Then
        If MsgBox("No Amounts Applied,Do You Want To Auto Apply", vbYesNo + vbQuestion) = vbYes Then
            Call cmd_Click
            AskForAutoApply = True
        End If
    Else
        AskForAutoApply = True
    End If

'If Val(lblToApply) <> 0 And LVPymt.ListItems.Count > 1 Then
'      MsgBox "There is Some Unapplied Amount, Cannot Save", vbInformation
'      AskForAutoApply = False
'End If

End Function




Private Sub ShowAccBalance()
'Show Previous Balance of Account
If lblBalance.Visible Then
    lblBalance.Caption = Format(getBalance(cmbAccNo, DTPicker1), "#,##0.00 DR;#,##0.00 CR;0.00 ")
    If InStr(lblBalance.Caption, "CR") > 0 Then
        lblBalance.ForeColor = vbBlue
    Else
        lblBalance.ForeColor = vbRed
    End If
    lblBalance.ToolTipText = "Balance of " & cmbAccNo & " as of " & Format(DTPicker1.Value, "dd/MM/yyyy")
End If

End Sub

Private Sub TBox_LostFocus(Index As Integer)
    On Error Resume Next
    'TBox(Index) = UCase(TBox(Index))
End Sub

Private Sub TBox_Validate(Index As Integer, cancel As Boolean)

    If Index <> 1 Then Exit Sub
    lblApplied = 0
    lblToApply = TBox(1)
    If Trim(cmd.Caption) = "Clear Payments" Then
        Call cmd_Click
        Call cmd_Click
    End If
    
End Sub


Private Sub ShowBals()
    If cmbPymtmode.ListIndex = 1 Then
        lblBalCap = "Bank Balance  "
        If lblBankAcc = "" Then Exit Sub
        lblBal = Format(getBalance(GetID(lblBankAcc), DTPicker1), "0.00")
    ElseIf cmbPymtmode.ListIndex = 0 Then
        lblBalCap = "Cash In Hand  "
        lblBal = Format(getBalance(CashInHand, DTPicker1), "0.00")
    End If
    
End Sub

Private Function SavePaymentTrans(ChqSNo As Double) As Boolean
    
    
    Dim i As Integer
    Dim NextVchrNo As String
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim dAmount As Double
    Dim sScNo As String
    Dim sCIH As String
    Dim Bal As Double
    Dim Desc As String
    Dim AccNo As String
    Dim Dated As Date
    'Get Serial # and Balances
    If strCompany = "Banzai" Then
        Dated = DTPicker1
    Else
        Dated = DTPicker2
    End If
    
    NextVchrNo = txtVchrNo.Text 'getNextVchrNo(Dated, IIf(cmbPymtmode.ListIndex = 0, "CPV", "BPV"))
    INextSno = getNextSno(Dated)
    
    
    SnoToUse = INextSno
    
    'Party Account's Balance
    Bal = getBalance(cmbAccNo, Dated)
    
    dAmount = Val(TBox(1).Text)
    
    sScNo = ChqSNo
    
    
    'Bank Account No.
    sCIH = GetID(lblBankAcc)
    
    Bal = Bal + dAmount
    
    Dim bWaitForAuth As Boolean
    bWaitForAuth = VoucherRequiresAuthorization("BPV")
    
    'Debit Entry
    bankDesc = "Chq #:" & TBox(2) & " " & TBox(0) 'Right(lblBankAcc, Len(lblBankAcc) - InStr(lblBankAcc, "]")) & " Chq # :" & TBox(2)
    Dim ChqNo As String
    Desc = "Cheque Paid To " & CmbAccTitle & " Chq# : " & TBox(2)
    ChqNo = TBox(2)
    
    If bWaitForAuth = False Then
        con.Execute "Insert into Vouchers Values(" & INextSno & ",'" & _
         Dated & "','" & NextVchrNo & "','" & cmbAccNo & _
         "','" & bankDesc & "'," & dAmount & ",0," & Bal & ",''," & sScNo & ")"
    
        Call adjustLedger(cmbAccNo, Dated, INextSno, Val(dAmount))
    
        INextSno = INextSno + 1
        Bal = getBalance(sCIH, Dated)
        Bal = Bal - dAmount
        'Description
    
        'Desc = "Payment Through " & cmbPymtmode.Text & " To " & CmbAccTitle & "  " & Label1(0) & " : " & Tbox(2)
      
        'Debit Entry (Bank Charges)
    
        If TBox(6).Visible And Val(TBox(6)) > 0 Then
            INextSno = INextSno + 1
    
            Bal = getBalance(cmbOtherAcc.ID, DTPicker1.Value)
            Bal = Bal + Val(TBox(6))
            
            con.Execute "INSERT INTO Vouchers VALUES(" & INextSno & ",'" & _
                Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & Label1(4) & "','" & cmbOtherAcc.ID & _
                "','" & bankDesc & "'," & Val(TBox(6)) & _
                ",0," & Bal & ",'',0)"
            Call adjustLedger(cmbOtherAcc.ID, DTPicker1.Value, INextSno, Val(TBox(6)))
        
        End If
    
    
        If TBox(6).Visible And Val(TBox(6)) > 0 Then
            dAmount = Val(TBox(1).Text) + Val(TBox(6))
        Else
            dAmount = Val(TBox(1).Text)
        End If
    
        INextSno = INextSno + 1
     
        'Credit Entry
    
        con.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
            Dated & "','" & NextVchrNo & "','" & sCIH & _
            "','" & bankDesc & "',0," & dAmount & "," & Bal & ",''," & sScNo & ")"
        
        Call adjustLedger(sCIH, Dated, INextSno, -Val(dAmount))
    Else
        con.Execute "INSERT INTO VouchersForAuthorization(VDate,VchrNo,Accno,Description,Debit,Credit,balance,DpstSlip,CSNo) VALUES('" & _
         Dated & "','" & NextVchrNo & "','" & cmbAccNo & _
         "','" & TBox(0) & bankDesc & "'," & dAmount & ",0," & Bal & ",''," & sScNo & ")"
        
        If TBox(6).Visible And Val(TBox(6)) > 0 Then
            INextSno = INextSno + 1
    
            Bal = getBalance(cmbOtherAcc.ID, DTPicker1.Value)
            Bal = Bal + Val(TBox(6))
            
            con.Execute "Insert into VouchersForAuthorization(VDate,VchrNo,Accno,Description,Debit,Credit,balance,DpstSlip,CSNo) VALUES('" & _
                Dated & "','" & NextVchrNo & "','" & cmbOtherAcc.ID & _
                "','" & TBox(0) & bankDesc & "'," & Val(TBox(6)) & _
                ",0," & Bal & ",''," & sScNo & ")"
            Call adjustLedger(cmbOtherAcc.ID, DTPicker1.Value, INextSno, Val(TBox(6)))
            
        End If
        
        If TBox(6).Visible And Val(TBox(6)) > 0 Then
            dAmount = Val(TBox(1).Text) + Val(TBox(6))
        Else
            dAmount = Val(TBox(1).Text)
        End If
        con.Execute "INSERT INTO VouchersForAuthorization(VDate,VchrNo,Accno,Description,Debit,Credit,balance,DpstSlip,CSNo) VALUES('" & _
         Dated & "','" & NextVchrNo & "','" & sCIH & _
         "','" & Desc & "',0," & dAmount & "," & Bal & ",''," & sScNo & ")"
    End If
    
    Dim rsCopy As New ADODB.Recordset
    With rsCopy
        .CursorLocation = adUseClient
        .Open "SELECT * FROM ChequeDetail WHERE (Sno=" & sScNo & ")", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            con.Execute "INSERT INTO VoucherDetail VALUES('" & ![BalancedSno] & "','" & SnoToUse & "'," & ![Amount] & ")"
            .MoveNEXT
        Loop
        .Close
    End With
    Set rsCopy = Nothing
    con.Execute "UPDATE Cheque SET Posted=1 WHERE Sno=" & sScNo
    If bWaitForAuth = False Then
        con.Execute "INSERT INTO VoucherInfo(VchrNo,UserName,MachineName,DT,ApprovedBy,Notes,Payee,Handed_Over_To) VALUES('" & NextVchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "','" & cmbApprovedBy.Text & "','" & TBox(4) & "','" & TBox(5) & "','" & TBox(7) & "')"
    Else
        con.Execute "INSERT INTO VouchersForAuthorizationInfo(VchrNo,UserName,MachineName,DT,ApprovedBy,Notes,Payee) VALUES('" & NextVchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "','" & cmbApprovedBy.Text & "','" & TBox(4) & "','" & TBox(5) & "')"
    End If
    
    If chkPrint = vbChecked Then Call PrintVoucher(MainForm.Cr1, NextVchrNo, True, Not bVoucherViewOnlyWhenSaved, bWaitForAuth)
    
    SavePaymentTrans = True
    Exit Function
err:
    'MsgBox "Transaction Terminated Due To Following Error;" & vbNewLine & err.Description
    SavePaymentTrans = False
    con.RollbackTrans
End Function

Private Sub txtVchrNo_Change()
    Label1(4).Caption = txtVchrNo.Text
End Sub


Private Function AddWHTaxOnCashWithdrawalVoucher() As String
    
    'Store Null To Return on an Unsuccessful Entry
    Dim VNo As String, m_DTVchr As Date
    If strCompany = "Banzai" Then
        m_DTVchr = DTPicker1.Value
    Else
        m_DTVchr = DTPicker2.Value
    End If
    
    VNo = getNextVchrNo(m_DTVchr, "JV")
    AddWHTaxOnCashWithdrawalVoucher = ""
    Dim i As Integer
    Dim dAmt As Double, strBills As String, Entries As String, strPOs As String, dDiscount As Double, dFreight As Double
    Dim lRefID As Long
    Dim strTaxAcc As String, dTaxRate As Double
    
    strTaxAcc = getGeneralDataValue("CashWithdrawalWHTaxAccNo")
    dTaxRate = Val(getGeneralDataValue("CashWithdrawalWHTaxRate"))
    If strTaxAcc = "" Then
        MsgBox "Can't Post Withholding Tax on Cash Withdrawal, Account No. not Found.", vbInformation
        Exit Function
    ElseIf dTaxRate = 0 Then
        MsgBox "Can't Post Withholding Tax on Cash Withdrawal, Rate not Found.", vbInformation
        Exit Function
    End If
    
    Dim dNetAmt As Double
    dNetAmt = (Val(TBox(1)) * dTaxRate / 100)
    
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim Bal As Double, Desc As String
    
    INextSno = getNextSno(m_DTVchr)
    'Remember Sno For Voucher Details
    SnoToUse = INextSno
    
    Desc = "W/H Tax On Cash Chq # " & TBox(2)
    
    'Start New Transaction
    'Call StartTrans(con)
      
    'Debit Entry....        '<<<<Change Order No.>>>> & Description etc.
    'strTaxAcc = "45-23006"
    If dNetAmt > 0 Then
        Bal = getBalance(strTaxAcc, m_DTVchr)
                
        Bal = Bal + dNetAmt
        con.Execute "INSERT INTO Vouchers Values(" & INextSno & ",'" & m_DTVchr & "','" & _
        VNo & "','" & strTaxAcc & "','" & Desc & "'," & _
        dNetAmt & ",0," & Bal & ",'',0)"
        
        Call adjustLedger(strTaxAcc, m_DTVchr, INextSno, dNetAmt)
        
    End If
    Dim strBankAccNo As String
    strBankAccNo = GetID(lblBankAcc)
    INextSno = INextSno + 1
    'Credit Entry
    Bal = getBalance(strBankAccNo, m_DTVchr)
    Bal = Bal - dNetAmt
   
    con.Execute "INSERT INTO Vouchers Values(" & INextSno & ",'" & _
         Format(m_DTVchr, "dd-MMM-yyyy") & "','" & VNo & "','" & strBankAccNo & _
        "','" & Desc & "',0," & dNetAmt & "," & Bal & ",'',0)"
          
    Call adjustLedger(strBankAccNo, m_DTVchr, INextSno, -dNetAmt)
    
    con.Execute "INSERT INTO VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & VNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
    
    
    '''Second Voucher
    'Debit Entry....        '<<<<Change Order No.>>>> & Description etc.
    'Debit Entry
    VNo = getNextVchrNo(m_DTVchr, "JV")
    Dim strPartyAccNo As String
    strPartyAccNo = cmbAccNo.Text
    If dNetAmt > 0 Then
        INextSno = INextSno + 1
        Bal = getBalance(strPartyAccNo, m_DTVchr)
        Bal = Bal + dNetAmt
        con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & m_DTVchr & "','" & _
        VNo & "','" & strPartyAccNo & "','" & Desc & "'," & _
        dNetAmt & ",0," & Bal & ",'',0)"
        
        Call adjustLedger(strPartyAccNo, m_DTVchr, INextSno, dNetAmt)
        
    End If
    
    INextSno = INextSno + 1
    'Credit Entry
    Bal = getBalance(strTaxAcc, m_DTVchr)
    Bal = Bal - dNetAmt
   
    con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
         Format(m_DTVchr, "dd-MMM-yyyy") & "','" & VNo & "','" & strTaxAcc & _
        "','" & Desc & "',0," & dNetAmt & "," & Bal & ",'',0)"
          
    Call adjustLedger(strTaxAcc, m_DTVchr, INextSno, -dNetAmt)
    
    con.Execute "Insert Into VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & VNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
    
    AddWHTaxOnCashWithdrawalVoucher = VNo
    
    Exit Function
err:
    AddWHTaxOnCashWithdrawalVoucher = False
    MsgBox "Error In AddVoucher :" & err.Number & vbNewLine & err.Description
End Function


