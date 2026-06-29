VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmRebateA 
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12165
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   12165
   WindowState     =   2  'Maximized
   Begin VB.Frame FAll 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   45
      TabIndex        =   4
      Top             =   30
      Width           =   12120
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   285
         TabIndex        =   0
         Top             =   1035
         Width           =   1740
         _ExtentX        =   3069
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
      Begin VB.Frame Frame3 
         Height          =   1335
         Left            =   240
         TabIndex        =   26
         Top             =   1440
         Width           =   11655
         Begin VB.TextBox txtReturnGoods 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   4950
            TabIndex        =   38
            Top             =   975
            Width           =   1035
         End
         Begin VB.TextBox txtDTREExp 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   3930
            TabIndex        =   37
            Top             =   975
            Width           =   1035
         End
         Begin VB.TextBox txtBGExp 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   2910
            TabIndex        =   36
            Top             =   975
            Width           =   1035
         End
         Begin VB.TextBox txtFCharges 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   8610
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   375
            Width           =   1260
         End
         Begin VB.TextBox txtCustomNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   5970
            TabIndex        =   39
            Top             =   975
            Width           =   1605
         End
         Begin VB.TextBox txtFOB 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   10425
            TabIndex        =   43
            Top             =   975
            Width           =   1170
         End
         Begin VB.TextBox txtInPak 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   9270
            TabIndex        =   41
            Top             =   975
            Width           =   1170
         End
         Begin VB.TextBox txtInsurance 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1890
            TabIndex        =   35
            Top             =   975
            Width           =   1035
         End
         Begin VB.TextBox txtExch 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   60
            TabIndex        =   33
            Top             =   975
            Width           =   825
         End
         Begin VB.TextBox txtAmtRealized 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   9855
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   375
            Width           =   1725
         End
         Begin VB.TextBox txtFreight 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   870
            TabIndex        =   34
            Top             =   975
            Width           =   1035
         End
         Begin VB.TextBox txtValue 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   6900
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   375
            Width           =   1725
         End
         Begin VB.TextBox txtPayTerms 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   5190
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   375
            Width           =   1725
         End
         Begin VB.TextBox txtBillNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1770
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   375
            Width           =   1725
         End
         Begin VB.TextBox txtFormE 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   375
            Width           =   1725
         End
         Begin VB.TextBox txtCommodity 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   375
            Width           =   1725
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   285
            Left            =   7590
            TabIndex        =   40
            Top             =   975
            Width           =   1695
            _ExtentX        =   2990
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
            CheckBox        =   -1  'True
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   100859907
            CurrentDate     =   37250
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"frmRebateA.frx":0000
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
            Left            =   60
            TabIndex        =   44
            Top             =   735
            Width           =   11535
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"frmRebateA.frx":00A4
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
            Left            =   60
            TabIndex        =   42
            Top             =   135
            Width           =   11520
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Annex I && II"
         ForeColor       =   &H00C00000&
         Height          =   1515
         Left            =   240
         TabIndex        =   20
         Top             =   2835
         Width           =   11670
         Begin VB.TextBox txtVessel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   8850
            TabIndex        =   24
            Top             =   1110
            Width           =   2745
         End
         Begin VB.TextBox txtRebate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   60
            TabIndex        =   23
            Top             =   1110
            Width           =   8805
         End
         Begin VB.TextBox txtGoodsVal 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   7680
            TabIndex        =   22
            Top             =   510
            Width           =   2175
         End
         Begin VB.TextBox txtDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   75
            TabIndex        =   21
            Top             =   510
            Width           =   7620
         End
         Begin MSComCtl2.DTPicker DTRec 
            Height          =   315
            Left            =   9870
            TabIndex        =   25
            Top             =   495
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   100859907
            CurrentDate     =   37770
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"frmRebateA.frx":0157
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
            Index           =   3
            Left            =   60
            TabIndex        =   46
            Top             =   870
            Width           =   11535
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"frmRebateA.frx":0219
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
            Left            =   75
            TabIndex        =   45
            Top             =   270
            Width           =   11535
         End
      End
      Begin VB.TextBox txtCurrency 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3810
         MaxLength       =   150
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1050
         Width           =   1005
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
         TabIndex        =   6
         Top             =   15
         Width           =   12120
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Rebate Documents"
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
            Left            =   60
            TabIndex        =   7
            Top             =   150
            Width           =   12000
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Rebate Documents"
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
            TabIndex        =   8
            Top             =   165
            Width           =   11985
         End
      End
      Begin MSComCtl2.DTPicker OrderDT 
         Height          =   300
         Left            =   6795
         TabIndex        =   3
         Top             =   1035
         Width           =   1785
         _ExtentX        =   3149
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
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   100859907
         CurrentDate     =   37250
      End
      Begin VB.Frame FraAdd 
         BackColor       =   &H80000004&
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
         Left            =   5085
         TabIndex        =   10
         Top             =   8280
         Width           =   6945
         Begin MSForms.CommandButton cmdPrint 
            Height          =   360
            Left            =   3480
            TabIndex        =   14
            Top             =   180
            Width           =   1665
            Caption         =   "Print             "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmRebateA.frx":02DE
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
            Left            =   5175
            TabIndex        =   13
            Top             =   180
            Width           =   1665
            Caption         =   "Close            "
            PicturePosition =   327683
            Size            =   "2937;635"
            Accelerator     =   67
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAdd 
            Height          =   360
            Left            =   75
            TabIndex        =   12
            Top             =   180
            Width           =   1665
            Caption         =   "Add New     "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmRebateA.frx":03F0
            Accelerator     =   78
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdEdit 
            Height          =   360
            Left            =   1785
            TabIndex        =   11
            Top             =   180
            Width           =   1665
            Caption         =   "Edit   "
            PicturePosition =   327683
            Size            =   "2937;635"
            Accelerator     =   69
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.Frame FraSave 
         BackColor       =   &H80000004&
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
         Left            =   8460
         TabIndex        =   15
         Top             =   8280
         Visible         =   0   'False
         Width           =   3555
         Begin MSForms.CommandButton cmdCancel 
            Height          =   360
            Left            =   1800
            TabIndex        =   17
            Top             =   180
            Width           =   1665
            Caption         =   "Cancel            "
            PicturePosition =   327683
            Size            =   "2937;635"
            Accelerator     =   67
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdSave 
            Height          =   360
            Left            =   90
            TabIndex        =   16
            Top             =   180
            Width           =   1665
            Caption         =   "Save        "
            PicturePosition =   327683
            Size            =   "2937;635"
            Accelerator     =   83
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   3450
         Left            =   240
         TabIndex        =   19
         Top             =   4815
         Width           =   11670
         _ExtentX        =   20585
         _ExtentY        =   6085
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         TextBackground  =   -1  'True
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
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "SNo"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Freight"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Insurance"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "BG Exp."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "DTRE Exp."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Return"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Net Total"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "SRO #"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "SRO Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Rate"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "S.Tax"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "E.Duty"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   12
            Text            =   "Rebate"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSForms.CommandButton cmdCalc 
         Height          =   360
         Left            =   240
         TabIndex        =   47
         Top             =   4395
         Visible         =   0   'False
         Width           =   2130
         Caption         =   "Calculate Rebate"
         PicturePosition =   327683
         Size            =   "3757;635"
         Accelerator     =   82
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbInvNoNew 
         Height          =   285
         Left            =   4845
         TabIndex        =   18
         Tag             =   "AutoCatID"
         Top             =   1050
         Visible         =   0   'False
         Width           =   1920
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3387;503"
         MatchEntry      =   1
         ListStyle       =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbInvNo 
         Height          =   285
         Left            =   4845
         TabIndex        =   9
         Top             =   1050
         Width           =   1920
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3387;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Customer                     Country                       Currency       Invoice No                       Rebate Date"
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
         Index           =   7
         Left            =   285
         TabIndex        =   5
         Top             =   795
         Width           =   8265
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Top             =   1050
         Width           =   1725
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3043;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frmRebateA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PInvoice As String
Dim Editing As Boolean
Dim strBank As String
Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()
    
    On Error GoTo err
    If cmbCountry.MatchFound = False Then Exit Sub
    
    'cmbCust.Tag = cmbCust.Text & cmbCountry
    cmbCust.Tag = cmbCust.Text & cmbCountry
    txtCurrency = cmbCountry.List(cmbCountry.ListIndex, 1)
    
    Dim rs As New ADODB.Recordset
    
    With rs
         .Open "Select Distinct CustomInvoice From VRebate Where CustCode+Country='" & cmbCust.Tag & "' Order By CustomInvoice", con, adOpenForwardOnly, adLockReadOnly
            cmbInvNo.Clear
            Do Until .EOF
                cmbInvNo.AddItem .Fields(0) & ""
               .MoveNext
            Loop
            .Close
            
            .Open "Select CustomInvoice from VrptCustomInvoice Where CustCode+Country='" & cmbCust.Tag & "' and CustomInvoice not in (Select CustomInvoice from Rebate) And TotalCustomAmtToRcv<=0 ORDER BY CustomInvoice", con, adOpenForwardOnly, adLockReadOnly
            cmbInvNoNew.Clear
            Do Until .EOF
                cmbInvNoNew.AddItem .Fields(0) & ""
                .MoveNext
            Loop
            .Close
        
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub cmbCust_matched()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Country,Curr From ForeignCustomers Where CustCode='" & cmbCust.id & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            cmbCountry.List(cmbCountry.ListCount - 1, 1) = ![Curr] & ""
            .MoveNext
        Loop
        .Close
    End With
    
    If cmbCountry.ListCount > 0 Then cmbCountry.ListIndex = 0
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbInvNoNew_Change()
    Call cmbInvNoNew_Click
End Sub

Private Sub cmbInvNoNew_Click()

    On Error GoTo err
    If Not Editing Then Call ClearTexts
    If cmbInvNoNew.MatchFound = False Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    
    With rs
        
        .Open "Select * From VrptCustomInvoice Where CustomInvoice='" & cmbInvNoNew.Text & "' AND TotalCustomAmtToRcv<=0", con, adOpenForwardOnly, adLockReadOnly
        
'        If Val(![TotalCustomAmtToRcv]) <> 0 Then
'            MsgBox "Amount To Rcv Is Not Null"
'            Exit Sub
'        End If
        
        DT = ![DT] & ""
        txtFormE = ![FormE] & ""
        txtBillNo = ![AWBNo] & ""
        DTBill = CheckForNullDT(![AWBNoDT] & "")
        txtPayTerms = ![PaymentTerms] & ""
        
        txtVessel = ![Vessel] & ""
       
        txtValue = ![TotalCustomAmt] & " " & txtCurrency.Text
       
        txtAmtRealized = ![TotalCustomAmtRcvd] & " " & txtCurrency.Text
        
        .Close
        .Open "SELECT * FROM VTotalCartons WHERE CustomInvoice='" & cmbInvNoNew.Text & "'", con, adOpenForwardOnly, adLockReadOnly
         
        If .EOF = False Then
            txtCommodity = ![TotalQty] & "-Pcs"
            TotCrtns = ![TotalCartons] & ""
            TotQty = ![TotalQty] & ""
        End If
        txtDesc = "Surgical Instruments"
        .Close
        
    End With
    
    Set rs = Nothing
    txtFCharges = 0
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmbinvno_Change()
    Call cmbinvno_click
End Sub

Private Sub cmbinvno_click()

    On Error GoTo err
    
    Call ClearTexts
    If cmbInvNo.MatchFound = False Then Exit Sub
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
    txtCurr = cmbCountry.Tag & ""
    With rs
        
        .Open "Select * From VRebate Where CustomInvoice='" & cmbInvNo.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        
        OrderDT.Value = Format(![DT], "dd-MM-yyyy")
        txtFormE = ![FormE] & ""
        txtBillNo = ![AWBNo] & ""
        DTBill = Format(![AWBNoDT], "dd-MM-yyyy")
        
        txtCommodity = ![Commodity] & ""
        txtDesc = ![Commodity] & ""
        
        txtPayTerms = ![PayTerms] & ""
        
        txtValue = ![TotalCustomAmt] & " US$"
        txtFCharges = Val(![FBankCharges] & "") & " US$"
        txtAmtRealized = Val(![AmtRealized] & "") & " US$"
        
        txtExch = Val(![ExchRate] & "")
        txtFreight = Val(![LessFreight] & "")
        txtInsurance = Val(![LessInsurance] & "")
        
        txtBGExp = Val(!BGExpenses & "")
        txtDTREExp = Val(!DTREExpenses & "")
        txtReturnGoods = Val(!ReturnGoodsValue & "")
        
        txtInPak = Val(![InPak] & "")
        txtFOB = Val(![NetFOB] & "")
        DTRec = ![AppDT] & ""
        
        

        txtVessel = ![Vessel] & ""

        txtCustomNo = ![CustomNo] & ""
        DTPicker1.Value = ![CustomDT] & ""
        
        Dim ITM As ListItem
        
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            
            ITM.Tag = ![Percentage] & ""
            ITM.ListSubItems.add(, , ![SROFreight] & "").Tag = ![SRONo] & ""
            
            ITM.ListSubItems.add , , ![SROInsurance] & ""
            ITM.ListSubItems.add , , Val(!BGExp & "")
            ITM.ListSubItems.add , , Val(!DTREExp & "")
            ITM.ListSubItems.add , , Val(!SROReturnGoodsValue & "")
            ITM.ListSubItems.add , , ![SROTotal] & ""
            ITM.ListSubItems.add , , ![CompleteNo] & ""
            ITM.ListSubItems.add , , Format(![SRODT] & "", "dd-MMM-yyyy")
            ITM.ListSubItems.add , , ![CRebate] & ""
            ITM.ListSubItems.add , , ![STax] & ""
            ITM.ListSubItems.add , , ![EDuty] & ""
            
            ITM.ListSubItems.add , , Round(Val(![TotalRebate] & "") - Val(![TotalSTax] & "") - Val(![TotalEDuty] & ""), 2)
'            ''''
'            dPer = Round((Val(![TotalCustomAmount] & "") / TotalCustomAmt) * 100, 4)
'            dLessFreight = Round((Val(txtFreight) / 100) * dPer, 4)
'            dLessInsurance = Round((Val(txtInsurance) / 100) * dPer, 4)
'            DTotal = Round((Val(txtFOB) / 100) * dPer, 4)
'
'
'            Set itm = LV.ListItems.Add(, , LV.ListItems.Count + 1)
'            itm.Tag = dPer
'            itm.ListSubItems.Add(, , dLessFreight).Tag = ![SRONo] & ""
'            itm.SubItems(1) = dLessFreight
'            itm.SubItems(2) = dLessInsurance
'            itm.SubItems(3) = DTotal
'            itm.SubItems(4) = ![CompleteNo] & ""
'            itm.SubItems(5) = Format(![SRODT] & "", "dd-MMM-yyyy")
'            itm.SubItems(6) = ![CRebate] & ""
'            itm.SubItems(7) = (DTotal / 100) * Val(![CRebate] & "")
'            '''
            .MoveNext
        Loop
        
        .Close
        
    End With
    'txtVal = ConvertInWords(TotalVal, cmbCountry.Tag)
    'txtOrders = strOrders
    Set rs = Nothing
    Exit Sub
    
err:
    MsgBox err.Description
End Sub

Private Sub cmdAdd_Click()
    'cmbInvNoNew.Visible = True
    Editing = False
    cmbInvNoNew.Locked = False
    ShowButts (False)
    
    Dim CtrlTxt As Control
    
    Call ClearTexts
    
End Sub

Private Sub cmdCalc_Click()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    If Val(txtExch) <= 0 Then
        MsgBox "Invalid Exchange Rate.", vbInformation
        txtExch.SetFocus
        Exit Sub
    ElseIf cmbInvNoNew.MatchFound = False And Not Editing Then
        MsgBox "Invalid Invoice No.", vbInformation
        cmbInvNoNew.SetFocus
        Exit Sub
    End If
    
    Dim strInvNo As String
    If Editing Then
        strInvNo = cmbInvNoNew
    Else
        strInvNo = cmbInvNo
    End If
    
    Dim TotalCustomAmt As Double
    'Dim Arr() As Variant
    'Arr = rsExecute("Select TotalCustomAmount From VFProformaAmt Where PInvoice='" & cmbInvNoNew & "'")
    rs.Open "SELECT * FROM VrptCustomInvoice WHERE CustomInvoice='" & cmbInvNoNew.Text & "'", con, adOpenForwardOnly, adLockReadOnly
    TotalCustomAmt = Round(Val(rs![TotalCustomAmtRcvd] & ""), 4)
    rs.Close
    
    Dim dPer As Double
    Dim dLessFreight As Double
    Dim dLessInsurance As Double
    Dim dBGExp As Double
    Dim dDTREExp As Double
    Dim dReturnGoods As Double
    Dim dTotal As Double
     
    Dim ITM As ListItem
    LV.ListItems.Clear
    
    With rs
        .Open "SELECT * FROM VCalcRebate WHERE CustomInvoice='" & cmbInvNoNew & "'", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
        
            dPer = Round((Val(![TotalCustomAmount] & "") / TotalCustomAmt) * 100, 1)
            dLessFreight = Round((Val(txtFreight) / 100) * dPer, 4)
            dLessInsurance = Round((Val(txtInsurance) / 100) * dPer, 4)
            dBGExp = Round((Val(txtBGExp) / 100) * dPer, 4)
            dDTREExp = Round((Val(txtDTREExp) / 100) * dPer, 4)
            dReturnGoods = Round((Val(txtReturnGoods) / 100) * dPer, 4)
            
            dTotal = Round((Val(txtFOB) / 100) * dPer, 4)
             
             
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            ITM.Tag = dPer
            ITM.ListSubItems.add(, , Round(dLessFreight, 2)).Tag = ![SRONo] & ""
            ITM.ListSubItems.add , , Round(dLessInsurance, 2)
            
            ITM.ListSubItems.add , , Round(dBGExp, 2)
            ITM.ListSubItems.add , , Round(dDTREExp, 2)
            ITM.ListSubItems.add , , Round(dReturnGoods, 2)
            
            ITM.ListSubItems.add , , Round(dTotal, 2)
            ITM.ListSubItems.add , , !CompleteNo & ""
            ITM.ListSubItems.add , , Format(![SRODT] & "", "dd-MMM-yyyy")
            ITM.ListSubItems.add , , !SRORate & ""
            
            ITM.ListSubItems.add , , !STax & ""
            ITM.ListSubItems.add , , !EDuty & ""
            
            ITM.ListSubItems.add , , Round((dTotal / 100) * Val(![SRORate] & "") - (dTotal / 100) * Val(![STax] & "") - (dTotal / 100) * Val(![EDuty] & ""), 2)
             
             
            .MoveNext
        Loop
        .Close
    End With
        
    Set rs = Nothing
    
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdCancel_Click()

    Editing = False
    Call ShowButts(True)
    Call cmbinvno_click
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo err
    
    If cmbInvNo.MatchFound = False Then
        MsgBox "Please Select Invoice #.", vbInformation
        Exit Sub
    End If
    Editing = True
    
    
    'Call StartTrans(con)
    'con.Execute "Delete from EForm Where CustomInvoice='" & cmbInvNo & "'"
    
    Call ShowButts(False)
    cmbInvNoNew.Text = cmbInvNo.Text
    'cmbInvNo.Tag = cmbInvNo.Text
    cmbInvNoNew.Locked = True
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
    
    If cmbInvNo.MatchFound = False Then
        MsgBox "Please Select Invoice No.", vbInformation
        Exit Sub
    End If
    
    Dim strRet As String
    strRet = ShowPopUpMenu(850 + FAll.Left + FraAdd.Left + cmdPrint.Left, FraAdd.Top + cmdPrint.Top - 800, "&Application", "-", "Document&s", "-", "&Calculation", "-", "Annexure &I", "A&nnexure II", "-", "&Declaration", "Rebate Payment Order")
    
    Dim rpt As Report
    Dim ShowRpt As Boolean
    ShowRpt = False
    Select Case strRet
        Case "&Application"
            Set rpt = rptApp.OpenReport(RptPath & "\rptApplication.rpt") 'New rptApplication
            ShowRpt = True
        Case "Document&s"
            Set rpt = rptApp.OpenReport(RptPath & "\rptRebateDocs.rpt") 'New rptRebateDocs
            ShowRpt = True
        Case "&Calculation"
            Set rpt = rptApp.OpenReport(RptPath & "\rptCalculation.rpt") 'New rptCalculation
            ShowRpt = True
        Case "Annexure &I"
            Set rpt = rptApp.OpenReport(RptPath & "\rptAnnex1.rpt") 'New rptAnnex1
            ShowRpt = True
        Case "A&nnexure II"
            Set rpt = rptApp.OpenReport(RptPath & "\rptAnnex2.rpt") 'New rptAnnex2
            ShowRpt = True
        Case "&Declaration"
            Set rpt = rptApp.OpenReport(RptPath & "\rptDeclaration.rpt") 'New rptDeclaration
            ShowRpt = True
        Case "Rebate Payment Order"
            Set rpt = rptApp.OpenReport(RptPath & "\RPO.rpt") 'New rptDeclaration
            ShowRpt = True
    End Select
    
    If ShowRpt Then
    
        Dim f As New frmPrevRpt
        f.ShowReport "{reb.CustomInvoice}='" & cmbInvNo.Text & "' ", rpt
        
        Set rpt = Nothing
        Set f = Nothing
        
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Public Sub PrintRebReports(rptName As String)

    On Error GoTo err
    
    With cr1
        .Connect = "DSN=DSN_MotoExport"
        .ReportFileName = App.Path & "\" & rptName
        .SelectionFormula = "{VRebate.CustomInvoice}='" & cmbInvNo.Text & "'"
        .Action = 1
    End With
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdSave_Click()

    On Error GoTo err
    
    If txtCustomNo = "" Then
        MsgBox "Enter Custom No.", vbInformation
        txtCustomNo.SetFocus
        Exit Sub
    ElseIf Val(txtExch) <= 0 Then
        MsgBox "Invalid Exchange Rate.", vbInformation
        txtExch.SetFocus
        Exit Sub
    ElseIf cmbInvNoNew.MatchFound = False And Not Editing Then
        MsgBox "Invalid Invoice No.", vbInformation
        cmbInvNoNew.SetFocus
        Exit Sub
    End If
        
    Dim strInvNo As String
    If Editing Then
        strInvNo = cmbInvNo
    Else
        strInvNo = cmbInvNoNew
    End If
      
    
    
    Call StartTrans(con)
    If Editing Then con.Execute "Delete from Rebate Where CustomInvoice='" & strInvNo & "'"
    
    con.Execute "INSERT INTO Rebate(CustomInvoice,DT,Bill,BillDT,Vessel,EForm,InvValue,PayTerms,Commodity," & _
     "ExchRate,AmtRealized,InPak,LessFreight,LessInsurance,BGExpenses,DTREExpenses,ReturnGoodsValue,NetFOB," & _
     "NetRebate,RecDT,Paid,Applied,AppDT,CustomNo,CustomDT,FBankCharges) VALUES('" & strInvNo & "','" & OrderDT & "','" & _
     txtBillNo & "','" & DTBill & "','" & txtVessel & "','" & txtFormE & "'," & Val(txtValue) & ",'" & txtPayTerms & "','" & _
     txtDesc & "','" & txtExch & "','" & Val(txtAmtRealized) & "','" & txtInPak & "','" & Val(txtFreight) & "','" & Val(txtInsurance) & _
     "'," & Val(txtBGExp) & "," & Val(txtDTREExp) & "," & Val(txtReturnGoods) & ",'" & txtFOB & "','" & NetRebate & "','" & DTRec.Value & "',0,0,'" & DTRec.Value & "','" & txtCustomNo & _
     "','" & DTPicker1.Value & "'," & Val(txtFCharges) & ")"
    
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            con.Execute "INSERT INTO RebateSro(CustomInvoice,Percentage,LessFreight,LessInsurance,BGExp,DTREExp,ReturnGoodsValue," & _
            "NetTotal,SRONo) VALUES('" & strInvNo & "'," & .Tag & "," & Val(.SubItems(1)) & "," & Val(.SubItems(2)) & "," & _
            Val(.SubItems(3)) & "," & Val(.SubItems(4)) & "," & Val(.SubItems(5)) & "," & Val(.SubItems(6)) & ",'" & .ListSubItems(1).Tag & "')"
        End With
    Next i
    'If Editing Then con.CommitTrans
    con.CommitTrans
    Call ShowButts(True)
    Call cmbCountry_Click
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub Form_Resize()

    On Error Resume Next
    FAll.Move ((Me.ScaleWidth - FAll.Width) / 2), ((Me.ScaleHeight - FAll.Height) / 2)

End Sub

Private Sub Form_Load()

    Dim CtrlTxt As Control
    
    For Each CtrlTxt In Controls
        If TypeOf CtrlTxt Is TextBox Then CtrlTxt.Locked = True
    Next CtrlTxt
    
    cmbCust.ListHeight = 3000
    Call cmbCust.ClearVals
    Call cmbCust.AddVals(con, "Distinct CustCode", "ForeignCustomers", "CustCode", "  Where CustCode<>'Hico'")
    Editing = False
    
    OrderDT = Date
    DTRec = Date
    DTPicker1.Value = Date
    
End Sub
Private Sub ShowButts(vis As Boolean)
    
    FraAdd.Visible = vis
    FraSave.Visible = Not vis
    
    cmbInvNo.Visible = vis
    cmbInvNoNew.Visible = Not vis
    
    cmdCalc.Visible = Not vis
    ''''
    
    Dim CtrlTxt As Control
    
    For Each CtrlTxt In Controls
        If TypeOf CtrlTxt Is TextBox Then
            If CtrlTxt.Name <> "txtCurrency" Then
                CtrlTxt.Locked = vis
            End If
        End If
    Next CtrlTxt
 
End Sub
Private Function GetAddress(str As String) As String
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Address from Customers Where CustCode+Country='" & cmbCust.Text & str & "'", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            GetAddress = .Fields(0) & ""
        End If
        .Close
    End With
    Set rs = Nothing
    Exit Function
err:
    MsgBox err.Description
End Function
Private Sub ClearTexts()
    Dim ctrl As Control
    For Each ctrl In Controls
        If TypeOf ctrl Is TextBox Then
            If ctrl.Name <> "txtCurrency" Then
                ctrl.Text = ""
            End If
        End If
        
    Next ctrl
    
End Sub





Private Sub EditRebate(iRow As Integer)

    With FRebate
    
        If .TextMatrix(iRow, 0) <> "" Then
            txtEAmt = .TextMatrix(iRow, 0) & ""
            txtEFreight = .TextMatrix(iRow, 1) & ""
            txtEInsurance = .TextMatrix(iRow, 2) & ""
            txtENet = .TextMatrix(iRow, 3) & ""
            cmbESRONo = .TextMatrix(iRow, 4) & ""
            txtESRODT = .TextMatrix(iRow, 5) & ""
            cmbERate = .TextMatrix(iRow, 6) & ""
            txtERebate = .TextMatrix(iRow, 7) & ""
        Else
            Dim dAmt As Double, dFreight As Double, dInsurance As Double
            For i = 1 To .Rows - 1
                dAmt = dAmt + Val(.TextMatrix(i, 0))
                dFreight = dFreight + Val(.TextMatrix(i, 1))
                dInsurance = dInsurance + Val(.TextMatrix(i, 2))
            Next i
            
            txtEAmt = Val(txtFOB) - dAmt
            txtEFreight = Val(txtFreight) - dFreight
            txtEInsurance = Val(dInsurance) - dInsurance
            txtENet = Val(txtEAmt) - Val(txtEFreight) - Val(txtEInsurance)
            
        End If
        
        .Row = iRow
        .Col = 0
        
        PicRebate.Move .Left, .Top + .CellTop
        PicRebate.Visible = True
    End With
    
End Sub

Private Sub txtAmtRealized_Change()
    Call ReCalculate
End Sub

Private Sub txtAmtRealized_KeyPress(KeyAscii As Integer)
    Call OnlyNums(KeyAscii)
End Sub

Private Sub CheckEsc(KeyAscii As Integer)
    If KeyAscii = 27 Then
        PicRebate.Visible = False
    End If
End Sub


Private Sub txtBGExp_Change()
    Call ReCalculate
End Sub

Private Sub txtDTREExp_Change()
    Call ReCalculate
End Sub

Private Sub txtExch_Change()
    Call ReCalculate
End Sub
Private Sub ReCalculate()
    txtInPak = Val(txtExch) * Val(txtAmtRealized)
    txtFOB = Val(txtInPak) - Val(txtFreight) - Val(txtInsurance) - Val(txtBGExp) - Val(txtDTREExp) - Val(txtReturnGoods)
End Sub

Private Sub ReCalculateE()
    txtENet = Val(txtEAmt) - Val(txtEFreight) - Val(txtEInsurance)
End Sub
Private Sub txtExch_KeyPress(KeyAscii As Integer)
    Call OnlyNums(KeyAscii)
End Sub


Private Sub txtFCharges_Change()
    txtAmtRealized = Val(txtValue) - Val(txtFCharges) & " " & txtCurrency.Text
End Sub

Private Sub txtFreight_Change()
    Call ReCalculate
End Sub

Private Sub txtFreight_KeyPress(KeyAscii As Integer)
    Call OnlyNums(KeyAscii)
End Sub

Private Sub txtInsurance_Change()
    Call ReCalculate
End Sub

Private Sub txtInsurance_KeyPress(KeyAscii As Integer)
    Call OnlyNums(KeyAscii)
End Sub


Private Sub txtReturnGoods_Change()
    Call ReCalculate
End Sub

Private Sub txtvalue_Change()
    Call ReCalculate
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
    Call OnlyNums(KeyAscii)
End Sub



