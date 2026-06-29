VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmRebate 
   Appearance      =   0  'Flat
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12465
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
   Icon            =   "frmRebate.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   12465
   WindowState     =   2  'Maximized
   Begin VB.Frame FALL 
      Height          =   9615
      Left            =   60
      TabIndex        =   0
      Top             =   150
      Width           =   12375
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   210
         TabIndex        =   1
         Top             =   945
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   503
         ListBackColor   =   16777215
         BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListForeColor   =   6244673
         TextBackColor   =   16777215
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextForeColor   =   6244673
         ListIndex       =   -1
         Appearance      =   0
      End
      Begin MSComCtl2.DTPicker DTBill 
         Height          =   300
         Left            =   4980
         TabIndex        =   70
         Top             =   1515
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   50855939
         CurrentDate     =   37770
      End
      Begin VB.Frame Frame2 
         Height          =   645
         Left            =   0
         TabIndex        =   59
         Top             =   0
         Width           =   12375
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
            Height          =   465
            Index           =   8
            Left            =   75
            TabIndex        =   61
            Top             =   135
            Width           =   12180
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
            TabIndex        =   60
            Top             =   150
            Width           =   12180
         End
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
         Left            =   6540
         TabIndex        =   36
         Top             =   1515
         Width           =   2685
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
         Left            =   210
         TabIndex        =   35
         Top             =   1515
         Width           =   2355
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
         Left            =   2595
         TabIndex        =   34
         Top             =   1515
         Width           =   2355
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
         Left            =   9255
         TabIndex        =   33
         Top             =   1515
         Width           =   2880
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
         Left            =   195
         TabIndex        =   32
         Top             =   2070
         Width           =   2970
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
         Left            =   9195
         TabIndex        =   31
         Top             =   2070
         Width           =   2970
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
         Left            =   3195
         TabIndex        =   30
         Top             =   2070
         Width           =   2970
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
         Left            =   6195
         TabIndex        =   29
         Top             =   2070
         Width           =   2970
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
         Left            =   195
         TabIndex        =   28
         Top             =   2610
         Width           =   2970
      End
      Begin VB.TextBox txtCurr 
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
         Left            =   3195
         TabIndex        =   27
         Top             =   2610
         Width           =   2970
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
         Left            =   6195
         TabIndex        =   26
         Top             =   2610
         Width           =   2970
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
         Left            =   9195
         TabIndex        =   25
         Top             =   2610
         Width           =   2970
      End
      Begin VB.Frame Frame1 
         Caption         =   "Annex I && II"
         ForeColor       =   &H00C00000&
         Height          =   1680
         Left            =   195
         TabIndex        =   11
         Top             =   2940
         Width           =   11970
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
            Left            =   315
            TabIndex        =   15
            Top             =   510
            Width           =   3855
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
            Left            =   4200
            TabIndex        =   14
            Top             =   510
            Width           =   2175
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
            Left            =   315
            TabIndex        =   13
            Top             =   1110
            Width           =   6060
         End
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
            Left            =   8145
            TabIndex        =   12
            Top             =   1110
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTRec 
            Height          =   300
            Left            =   6420
            TabIndex        =   71
            Top             =   510
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   50855939
            CurrentDate     =   37770
         End
         Begin MSComCtl2.DTPicker DTShip 
            Height          =   300
            Left            =   8145
            TabIndex        =   72
            Top             =   510
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   50855939
            CurrentDate     =   37770
         End
         Begin MSComCtl2.DTPicker DTClaim 
            Height          =   300
            Left            =   9870
            TabIndex        =   73
            Top             =   510
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   50855939
            CurrentDate     =   37770
         End
         Begin MSComCtl2.DTPicker DTExp 
            Height          =   300
            Left            =   6420
            TabIndex        =   74
            Top             =   1110
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   50855939
            CurrentDate     =   37770
         End
         Begin MSComCtl2.DTPicker DTVessel 
            Height          =   300
            Left            =   9870
            TabIndex        =   75
            Top             =   1125
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   50855939
            CurrentDate     =   37770
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description Of Goods"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   19
            Left            =   360
            TabIndex        =   24
            Top             =   285
            Width           =   1515
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Value Of Goods"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   20
            Left            =   4200
            TabIndex        =   23
            Top             =   285
            Width           =   1110
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reciept Date"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   21
            Left            =   6450
            TabIndex        =   22
            Top             =   285
            Width           =   930
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ship. Date"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   22
            Left            =   7965
            TabIndex        =   21
            Top             =   300
            Width           =   750
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Claim Date"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   23
            Left            =   9540
            TabIndex        =   20
            Top             =   285
            Width           =   765
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rebate In Words"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   24
            Left            =   375
            TabIndex        =   19
            Top             =   900
            Width           =   1230
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Export Date"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   25
            Left            =   6465
            TabIndex        =   18
            Top             =   900
            Width           =   870
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vessel Name"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   26
            Left            =   7965
            TabIndex        =   17
            Top             =   900
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vessel Date"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   27
            Left            =   9495
            TabIndex        =   16
            Top             =   915
            Width           =   840
         End
      End
      Begin VB.PictureBox PicRebate 
         Appearance      =   0  'Flat
         BackColor       =   &H00CFF0F3&
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   180
         ScaleHeight     =   270
         ScaleWidth      =   11625
         TabIndex        =   2
         Top             =   5250
         Visible         =   0   'False
         Width           =   11655
         Begin VB.TextBox txtEAmt 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   -15
            TabIndex        =   8
            Top             =   -15
            Width           =   1290
         End
         Begin VB.TextBox txtEFreight 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   1260
            TabIndex        =   7
            Top             =   -15
            Width           =   1245
         End
         Begin VB.TextBox txtEInsurance 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   2490
            TabIndex        =   6
            Top             =   -15
            Width           =   1455
         End
         Begin VB.TextBox txtENet 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   3930
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   -15
            Width           =   1305
         End
         Begin VB.TextBox txtESRODT 
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
            Height          =   285
            Left            =   6735
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   -15
            Width           =   1890
         End
         Begin VB.TextBox txtERebate 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   10035
            TabIndex        =   3
            Top             =   -15
            Width           =   1605
         End
         Begin MSForms.ComboBox cmbESRONo 
            Height          =   285
            Left            =   5145
            TabIndex        =   10
            Top             =   -15
            Width           =   1590
            VariousPropertyBits=   746604571
            BackColor       =   16777215
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "2805;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbERate 
            Height          =   285
            Left            =   8610
            TabIndex        =   9
            Top             =   -15
            Width           =   1425
            VariousPropertyBits=   746604571
            BackColor       =   16777215
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "2514;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin Crystal.CrystalReport Cr1 
         Left            =   195
         Top             =   150
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSFlexGridLib.MSFlexGrid FRebate 
         Height          =   4140
         Left            =   165
         TabIndex        =   37
         Top             =   4710
         Width           =   11985
         _ExtentX        =   21140
         _ExtentY        =   7303
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   0
         ForeColorFixed  =   8388608
         BackColorBkg    =   -2147483633
         GridColor       =   6244673
         GridColorFixed  =   6244673
         MergeCells      =   4
         AllowUserResizing=   3
         Appearance      =   0
         FormatString    =   $"frmRebate.frx":0442
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DT 
         Height          =   300
         Left            =   10170
         TabIndex        =   76
         Top             =   945
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   50855939
         CurrentDate     =   37770
      End
      Begin VB.Frame FraAdd 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5385
         TabIndex        =   62
         Top             =   8910
         Width           =   6900
         Begin MSForms.CommandButton cmdPrint 
            Height          =   360
            Left            =   3450
            TabIndex        =   66
            Top             =   165
            Width           =   1665
            Caption         =   "Print "
            PicturePosition =   327683
            Size            =   "2937;635"
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
            Left            =   5145
            TabIndex        =   65
            Top             =   165
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
            Left            =   60
            TabIndex        =   64
            Top             =   165
            Width           =   1665
            Caption         =   "New "
            PicturePosition =   327683
            Size            =   "2937;635"
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
            Left            =   1755
            TabIndex        =   63
            Top             =   165
            Width           =   1665
            Caption         =   "Edit "
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8760
         TabIndex        =   67
         Top             =   8910
         Visible         =   0   'False
         Width           =   3525
         Begin MSForms.CommandButton cmdSave 
            Height          =   360
            Left            =   75
            TabIndex        =   69
            Top             =   180
            Width           =   1665
            Caption         =   " Save"
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
         Begin MSForms.CommandButton cmdCancel 
            Height          =   360
            Left            =   1770
            TabIndex        =   68
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
      End
      Begin MSForms.ComboBox cmbInvoice 
         Height          =   300
         Left            =   7335
         TabIndex        =   58
         Top             =   945
         Visible         =   0   'False
         Width           =   2790
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4921;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbPList 
         Height          =   300
         Left            =   7335
         TabIndex        =   57
         Top             =   945
         Width           =   2790
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4921;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   7470
         TabIndex        =   56
         Top             =   750
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   55
         Top             =   750
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   4170
         TabIndex        =   54
         Top             =   750
         Width           =   585
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   300
         Left            =   4140
         TabIndex        =   53
         Top             =   945
         Width           =   3165
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5583;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbCriteria 
         Height          =   300
         Left            =   210
         TabIndex        =   52
         Top             =   945
         Visible         =   0   'False
         Width           =   1740
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         DisplayStyle    =   3
         Size            =   "3069;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Commodity"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   6540
         TabIndex        =   51
         Top             =   1305
         Width           =   1395
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   10170
         TabIndex        =   50
         Top             =   750
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Form E #"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   255
         TabIndex        =   49
         Top             =   1305
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shipping Bill #"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   2595
         TabIndex        =   48
         Top             =   1305
         Width           =   1590
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bill Date"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   9
         Left            =   4980
         TabIndex        =   47
         Top             =   1305
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terms"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   10
         Left            =   9255
         TabIndex        =   46
         Top             =   1305
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   11
         Left            =   255
         TabIndex        =   45
         Top             =   1860
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Less Freight"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   12
         Left            =   9195
         TabIndex        =   44
         Top             =   1875
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amt Realized"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   13
         Left            =   3195
         TabIndex        =   43
         Top             =   1875
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exch. Rate"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   14
         Left            =   6195
         TabIndex        =   42
         Top             =   1875
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Less Insurance"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   15
         Left            =   255
         TabIndex        =   41
         Top             =   2415
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Currency"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   16
         Left            =   3195
         TabIndex        =   40
         Top             =   2415
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "In Pak"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   17
         Left            =   6195
         TabIndex        =   39
         Top             =   2415
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FOB Value"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   18
         Left            =   9195
         TabIndex        =   38
         Top             =   2415
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmRebate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Editing As Boolean
Dim strBank As String
Private Sub cmbCust_matched()
    'cmbCountry.ClearVals
    
    Call FillCountries(cmbCust.Text, cmbCountry)
    'Call FillCountries(cmbCust.Text, cmbNotify)
    
End Sub

Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()

    On Error GoTo err
    If cmbCountry.MatchFound Then
        cmbCust.Tag = cmbCust.Text & cmbCountry
        Dim rs As New ADODB.Recordset
        With rs
            .Open "Select Distinct CustomInvoice from VRebate Where CustCode+Country='" & cmbCust.Tag & "'", con, adOpenForwardOnly, adLockReadOnly
            cmbPList.Clear
            Do Until .EOF
                cmbPList.AddItem .Fields(0) & ""
               .MoveNext
            Loop
            .Close
            
            .Open "Select Distinct CustomInvoice from VCustomInvoice Where CustCode+Country='" & cmbCust.Tag & "' and CustomInvoice not in (Select CustomInvoice from Rebate)", con, adOpenForwardOnly, adLockReadOnly
            cmbInvoice.Clear
            Do Until .EOF
                cmbInvoice.AddItem .Fields(0) & ""
                .MoveNext
            Loop
            .Close
            
            .Open "Select Curr From ForeignCustomers Where CustCode+Country='" & cmbCust.Tag & "'", con, adOpenForwardOnly, adLockReadOnly
            cmbCountry.Tag = .Fields(0) & ""
            txtCurr = .Fields(0) & ""
            .Close
        End With
    End If
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmbERate_Change()
    Call cmbERate_Click
End Sub

Private Sub cmbERate_Click()
    txtERebate = (Val(cmbERate) / 100) * Val(txtENet)
End Sub

Private Sub cmbERate_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    Select Case KeyCode.value
        Case 13 'Enter Key
            With FRebate
                .TextMatrix(.Tag, 0) = txtEAmt
                .TextMatrix(.Tag, 1) = txtEFreight
                .TextMatrix(.Tag, 2) = txtEInsurance
                .TextMatrix(.Tag, 3) = txtENet
                .TextMatrix(.Tag, 4) = cmbESRONo
                .TextMatrix(.Tag, 5) = txtESRODT
                .TextMatrix(.Tag, 6) = cmbERate
                .TextMatrix(.Tag, 7) = txtERebate
                .TextMatrix(.Tag, 8) = cmbERate.List(cmbERate.ListIndex, 1)
                'Now Check To See If Amount Is Left Or Not....
                Dim DTotal As Double
                DTotal = 0
                
                For i = 1 To .Rows - 1
                    DTotal = DTotal + Val(.TextMatrix(i, 0))
                Next i
                If DTotal < Val(txtFOB) Then
                    .AddItem ""
                End If
                PicRebate.Visible = False
            End With
            
            Call SetBackColor(FRebate, FRebate.Tag)
            
        Case 27 'Escacpe Key
            PicRebate.Visible = False
    End Select
    
End Sub

Private Sub cmbESRONo_Change()
    Call cmbESRONo_Click
End Sub

Private Sub cmbESRONo_Click()

    With cmbESRONo
        If .MatchFound Then
            txtESRODT = .List(.ListIndex, 1)
            cmbERate.Clear
            
            cmbERate.AddItem .List(.ListIndex, 2)
            cmbERate.List(cmbERate.ListCount - 1, 1) = .List(.ListIndex, 3)
            cmbERate.ListIndex = 0
            
        End If
    End With
    
End Sub

Private Sub cmbESRONo_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Call CheckEsc(KeyAscii.value)
End Sub

Private Sub cmbInvoice_Change()
    Call cmbInvoice_Click
End Sub

Private Sub cmbInvoice_Click()
    On Error GoTo err
    If Not Editing Then Call ClearTexts
    If cmbInvoice.MatchFound = False Then Exit Sub
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select * From VCustomInvoice Where CustomInvoice='" & cmbInvoice.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        
        TotCrtns = ![TotalCartons] & ""
        TotQty = ![TotalQty] & ""
        txtValue = ![TotalCustomAmt] & ""
        txtAmtRealized = ![TotalCustomAmt] & ""
        txtCurr = cmbCountry.Tag & ""
    
        DT = ![DT] & ""
        txtFormE = ![FormE] & ""
        txtBillNo = ![AWBNo] & ""
        DTBill = CheckForNullDT(![AWBNoDT] & "")
        txtPayTerms = ![PaymentTerms] & ""
        
        txtCommodity = ![TotalQty] & "-Pcs"
        txtDesc = ![TotalQty] & "-Pcs"
        
        .Close
        
    End With
    
    'txttQty = UnitTotals
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmbPList_Change()
    Call cmbPList_Click
End Sub

Private Sub cmbPList_Click()
    On Error GoTo err
    
    Call ClearTexts
    If cmbPList.MatchFound = False Then Exit Sub
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
    txtCurr = cmbCountry.Tag & ""
    With rs
        
        .Open "Select * From VRebate Where CustomInvoice='" & cmbPList.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        
        DT = Format(![DT], "dd-MM-yyyy")
        txtFormE = ![EForm] & ""
        txtBillNo = ![Bill] & ""
        DTBill = Format(![BillDT], "dd-MM-yyyy")
        
        txtCommodity = ![Commodity] & ""
        txtDesc = ![Commodity] & ""
        
        txtPayTerms = ![PayTerms] & ""
        txtValue = ![InvValue] & ""
        txtAmtRealized = Val(![AmtRealized] & "")
        txtExch = Val(![ExchRate] & "")
        txtFreight = Val(![LessFreight] & "")
        txtInsurance = Val(![LessInsurance] & "")
        txtInPak = Val(![InPak] & "")
        txtFOB = Val(![NetFOB] & "")
        DTRec = Format(![RecDT], "dd-MM-yyyy")
        DTClaim = Format(![ClaimDT], "dd-MM-yyyy")
        DTShip = Format(![ShipDT], "dd-MM-yyyy")
        DTExp = Format(![ExpDT], "dd-MM-yyyy")
        txtVessel = ![Vessel] & ""
        DTVessel = Format(![VesselDT], "dd-MM-yyyy")
        
        Dim iRow As Integer
        iRow = 0
        Call ClearRows(FRebate)
        Do Until .EOF
            iRow = .AbsolutePosition
            If iRow > 1 Then FRebate.AddItem ""
            FRebate.TextMatrix(iRow, 0) = Val(![SROAmt])
            FRebate.TextMatrix(iRow, 1) = Val(![SROFreight])
            FRebate.TextMatrix(iRow, 2) = Val(![SROInsurance])
            FRebate.TextMatrix(iRow, 3) = Val(![SROTotal])
            FRebate.TextMatrix(iRow, 4) = Val(![CompleteNo])
            FRebate.TextMatrix(iRow, 5) = Format(![SRODT], "dd-MM-yyyy")
            FRebate.TextMatrix(iRow, 6) = Val(![CRebate])
            FRebate.TextMatrix(iRow, 7) = Val(![Rebate])
            FRebate.TextMatrix(iRow, 8) = Val(![SROID])
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
    'cmbInvoice.Visible = True
    Editing = False
    cmbInvoice.Locked = False
    ShowButts (False)
    
    Dim CtrlTxt As Control
    
    Call ClearTexts
    
End Sub

Private Sub cmdCancel_Click()
    If Editing Then con.RollbackTrans
    Call ShowButts(True)
    Call cmbPList_Click
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdedit_Click()
    On Error GoTo err
    
    If cmbPList.MatchFound = False Then
        MsgBox "Please Select Invoice #.", vbInformation
        Exit Sub
    End If
    Editing = True
    
    
    'Call StartTrans(con)
    'con.Execute "Delete from EForm Where CustomInvoice='" & cmbPList & "'"
    
    Call ShowButts(False)
    cmbInvoice.Text = cmbPList.Text
    'cmbPList.Tag = cmbPList.Text
    cmbInvoice.Locked = True
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
    
    If cmbPList.MatchFound = False Then
        MsgBox "Please Select Invoice No.", vbInformation
        Exit Sub
    End If
    
    Dim strRet As String
    strRet = ShowPopUpMenu(FraSave.Left + cmdPrint.Left - 500, FraSave.Top + cmdPrint.Top + 500, "&Application", "-", "&Calculation", "-", "Annexure &I", "A&nnexure II", "-", "&Declaration")
    
    Dim rpt As Report
    Dim ShowRpt As Boolean
    ShowRpt = False
    Select Case strRet
        Case "&Application"
            'Set Rpt = New rptApplication
            ShowRpt = True
        Case "&Calculation"
            'Set Rpt = New rptCalculation
            ShowRpt = True
        Case "Annexure &I"
            'Set Rpt = New rptAnnex1
            ShowRpt = True
        Case "A&nnexure II"
            'Set Rpt = New rptAnnex2
            ShowRpt = True
        Case "&Declaration"
            'Set Rpt = New rptDeclaration
            ShowRpt = True
    End Select
    
    If ShowRpt Then
    
        Dim F As New frmPrevRpt
        F.ShowReport "{VRebate.CustomInvoice}='" & cmbPList.Text & "' ", rpt
        
        Set rpt = Nothing
        Set F = Nothing
        
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Public Sub PrintRebReports(rptName As String)

    On Error GoTo err
    
    With Cr1
        .Connect = "DSN=DSN_MotoExport"
        .ReportFileName = App.Path & "\" & rptName
        .SelectionFormula = "{VRebate.CustomInvoice}='" & cmbPList.Text & "'"
        .Action = 1
    End With
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdSave_Click()
    On Error GoTo err
    
    
    If cmbInvoice.MatchFound = False And Editing = False Then Exit Sub
    
    Dim DTotal As Double
    With FRebate
        For i = 1 To .Rows - 1
            DTotal = DTotal + Val(.TextMatrix(i, 0))
        Next i
    End With
    
    If DTotal <> Val(txtFOB) Then
        MsgBox "Totals Are Not Same.", vbInformation
        Exit Sub
    End If
    
    If Editing And cmbInvoice.Text = "" Then Exit Sub
     
    Call StartTrans(con)
    If Editing Then con.Execute "Delete from Rebate Where CustomInvoice='" & cmbPList & "'"
    
    con.Execute "Insert Into Rebate(CustomInvoice,DT,Bill,BillDT," & _
     "ExpDT,Vessel,VesselDT,EForm,InvValue,PayTerms,Commodity," & _
     "ExchRate,AmtRealized,InPak,LessFreight,LessInsurance," & _
     "NetFOB,NetRebate,ClaimDT,RecDT,ShipDT,Paid,Applied,AppDT)" & _
     "Values('" & cmbInvoice.Text & "','" & DT & "','" & _
     txtBillNo & "','" & DTBill & "','" & DTExp & _
     "','" & txtVessel & "'," & DTVessel & ",'" & txtFormE & _
     "'," & Val(txtValue) & ",'" & txtPayTerms & "','" & _
     txtCommodity & "','" & txtExch & "','" & txtAmtRealized & _
     "','" & txtInPak & "','" & txtFreight & "','" & txtInsurance & _
     "','" & txtFOB & "','" & NetRebate & "','" & DTClaim & "','" & _
     DTRec & "','" & DTShip & "',0,0,null)"
    
    With FRebate
        For i = 1 To .Rows - 1
            con.Execute "Insert Into RebateSro(CustomInvoice,Amount," & _
             "LessFreight,LessInsurance,NetTotal,SROID) Values('" & _
             cmbInvoice.Text & "'," & Val(.TextMatrix(i, 0)) & _
             "," & Val(.TextMatrix(i, 1)) & "," & Val(.TextMatrix(i, 2)) & _
             "," & Val(.TextMatrix(i, 3)) & "," & Val(.TextMatrix(i, 8)) & _
             ")"
        Next i
    End With
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
    Call cmbCust.AddVals(con, "Distinct CustCode", "ForeignCustomers")
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select EntryID,CompleteNo,SRODT,CRebate from SROs", con, adOpenForwardOnly, adLockReadOnly
        cmbESRONo.Clear
        Do Until .EOF
            cmbESRONo.AddItem ![CompleteNo] & ""
            cmbESRONo.List(.AbsolutePosition - 1, 1) = Format(![SRODT] & "", "dd-MM-yyyy")
            cmbESRONo.List(.AbsolutePosition - 1, 2) = Val(![CRebate] & "")
            cmbESRONo.List(.AbsolutePosition - 1, 3) = Val(![EntryID] & "")
            .MoveNext
        Loop
    End With
    FRebate.ColWidth(8) = 0
    'FRebate.TextMatrix(1, 0) = "Double Click" ' To Add"
End Sub
Private Sub ShowButts(vis As Boolean)
    
    FraAdd.Visible = vis
    FraSave.Visible = Not vis
    
    cmbPList.Visible = vis
    cmbInvoice.Visible = Not vis
    ''''
    
    Dim CtrlTxt As Control
    
    For Each CtrlTxt In Controls
        If TypeOf CtrlTxt Is TextBox Then CtrlTxt.Locked = vis
    Next CtrlTxt
 
End Sub
Private Function GetAddress(Str As String) As String
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Address from Customers Where CustCode+Country='" & cmbCust.Text & Str & "'", con, adOpenForwardOnly, adLockReadOnly
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
        If TypeOf ctrl Is TextBox Then ctrl.Text = ""
    Next ctrl
    Call ClearRows(FRebate)
End Sub




Private Sub FRebate_DblClick()
    Dim iRow As Integer
    If Val(txtExch) = 0 Then
        MsgBox "Please Enter Exchange Rate.", vbInformation
        Exit Sub
    End If
    With FRebate
        iRow = .MouseRow
        .Tag = iRow
        If iRow > 0 Then Call EditRebate(iRow)
    End With
    
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

Private Sub txteAmt_Change()
    Call ReCalculateE
End Sub

Private Sub txtEAmt_KeyPress(KeyAscii As Integer)
    Call CheckEsc(KeyAscii)
    Call OnlyNums(KeyAscii)
    cmbESRONo.SetFocus
End Sub
Private Sub CheckEsc(KeyAscii As Integer)
    If KeyAscii = 27 Then
        PicRebate.Visible = False
    End If
End Sub
Private Sub txtEFreight_Change()
    Call ReCalculateE
End Sub

Private Sub txtEFreight_KeyPress(KeyAscii As Integer)
    Call CheckEsc(KeyAscii)
    Call OnlyNums(KeyAscii)
End Sub

Private Sub txtEInsurance_Change()
    Call ReCalculateE
End Sub

Private Sub txtEInsurance_KeyPress(KeyAscii As Integer)
    Call CheckEsc(KeyAscii)
    Call OnlyNums(KeyAscii)
End Sub

Private Sub txtENet_KeyPress(KeyAscii As Integer)
    Call CheckEsc(KeyAscii)
End Sub

Private Sub txtERebate_KeyPress(KeyAscii As Integer)
    Call CheckEsc(KeyAscii)
End Sub

Private Sub txtESRODT_KeyPress(KeyAscii As Integer)
    Call CheckEsc(KeyAscii)
End Sub

Private Sub txtExch_Change()
    Call ReCalculate
End Sub
Private Sub ReCalculate()
    txtInPak = Val(txtExch) * Val(txtAmtRealized)
    txtFOB = Val(txtInPak) - Val(txtFreight) - Val(txtInsurance)
End Sub

Private Sub ReCalculateE()
    txtENet = Val(txtEAmt) - Val(txtEFreight) - Val(txtEInsurance)
End Sub
Private Sub txtExch_KeyPress(KeyAscii As Integer)
    Call OnlyNums(KeyAscii)
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


Private Sub txtValue_Change()
    Call ReCalculate
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
    Call OnlyNums(KeyAscii)
End Sub

