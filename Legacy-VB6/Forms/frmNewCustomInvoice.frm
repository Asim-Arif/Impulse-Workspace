VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmNewCustomInvoice 
   ClientHeight    =   9570
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12390
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
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   12390
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
      Height          =   9630
      Left            =   90
      TabIndex        =   9
      Top             =   -75
      Width           =   12270
      Begin VB.Frame Frame9 
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
         TabIndex        =   150
         Top             =   0
         Width           =   12255
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "New Custom Invoice"
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
            Index           =   33
            Left            =   75
            TabIndex        =   151
            Top             =   135
            Width           =   12075
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "New Custom Invoice"
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
            Index           =   34
            Left            =   90
            TabIndex        =   152
            Top             =   150
            Width           =   12060
         End
      End
      Begin ComboList.Usercontrol1 cmbItemSearchInvoice 
         Height          =   285
         Left            =   1740
         TabIndex        =   146
         Top             =   8040
         Width           =   6405
         _ExtentX        =   11298
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
         TextForeColor   =   16711680
         ListIndex       =   -1
         Appearance      =   0
      End
      Begin ComboList.Usercontrol1 cmbItemSearch 
         Height          =   285
         Left            =   1740
         TabIndex        =   139
         Top             =   8145
         Visible         =   0   'False
         Width           =   6405
         _ExtentX        =   11298
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
         TextForeColor   =   16711680
         ListIndex       =   -1
         Appearance      =   0
      End
      Begin ComboList.Usercontrol1 cmbCompany 
         Height          =   285
         Left            =   1515
         TabIndex        =   0
         Top             =   825
         Width           =   4035
         _ExtentX        =   7117
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
         TextForeColor   =   16711680
         ListIndex       =   -1
         Appearance      =   0
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   90
         TabIndex        =   91
         Top             =   1395
         Width           =   2310
         _ExtentX        =   4075
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
      Begin VB.TextBox txtBatchNo 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   10410
         MaxLength       =   150
         TabIndex        =   24
         Top             =   1410
         Width           =   1620
      End
      Begin VB.TextBox txtCurrency 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3060
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   9150
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.TextBox txtPInvoice 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4575
         MaxLength       =   150
         TabIndex        =   3
         Top             =   1410
         Width           =   2700
      End
      Begin MSComCtl2.DTPicker OrderDT 
         Height          =   300
         Left            =   7290
         TabIndex        =   4
         Top             =   1395
         Width           =   1605
         _ExtentX        =   2831
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
         Format          =   300548099
         CurrentDate     =   37250
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6765
         Left            =   120
         TabIndex        =   17
         Top             =   1740
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   11933
         _Version        =   393216
         Tabs            =   5
         Tab             =   1
         TabsPerRow      =   5
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         TabCaption(0)   =   "Invoice Items"
         TabPicture(0)   =   "frmNewCustomInvoice.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "cmbOrders1"
         Tab(0).Control(1)=   "Label1(0)"
         Tab(0).Control(2)=   "cmdAddItem1"
         Tab(0).Control(3)=   "Label1(16)"
         Tab(0).Control(4)=   "lblTotPcs"
         Tab(0).Control(5)=   "lblGoTo(1)"
         Tab(0).Control(6)=   "LV"
         Tab(0).Control(7)=   "cmdDelete"
         Tab(0).Control(8)=   "txtQty"
         Tab(0).Control(9)=   "txtPrice"
         Tab(0).Control(10)=   "txtDTRENo"
         Tab(0).ControlCount=   11
         TabCaption(1)   =   "Other Info."
         TabPicture(1)   =   "frmNewCustomInvoice.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "cmbSMethod"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label1(21)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label1(23)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "cmbBank"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label1(24)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "txtInvoiceof"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Label1(30)"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "cmbBankAccNo"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "Label1(22)"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "Label1(31)"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "Label1(32)"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "txtDTREDesc"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "DTGatePass"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "DTLC"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "Frame5"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "Frame3"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "Frame4"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).Control(17)=   "Frame2"
         Tab(1).Control(17).Enabled=   0   'False
         Tab(1).Control(18)=   "txtLC"
         Tab(1).Control(18).Enabled=   0   'False
         Tab(1).Control(19)=   "txtGatePassNo"
         Tab(1).Control(19).Enabled=   0   'False
         Tab(1).Control(20)=   "chkNoSRO"
         Tab(1).Control(20).Enabled=   0   'False
         Tab(1).Control(21)=   "chkCIFTerms"
         Tab(1).Control(21).Enabled=   0   'False
         Tab(1).ControlCount=   22
         TabCaption(2)   =   "Proforma Items"
         TabPicture(2)   =   "frmNewCustomInvoice.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmdAddItem"
         Tab(2).Control(1)=   "cmbProforma"
         Tab(2).Control(2)=   "Label1(8)"
         Tab(2).Control(3)=   "cmdAddAll"
         Tab(2).Control(4)=   "lblGoTo(0)"
         Tab(2).Control(5)=   "LV2"
         Tab(2).Control(6)=   "txtQty1"
         Tab(2).Control(7)=   "Command1"
         Tab(2).Control(8)=   "cmbProformas"
         Tab(2).ControlCount=   9
         TabCaption(3)   =   "Commercial Invoice"
         TabPicture(3)   =   "frmNewCustomInvoice.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame7"
         Tab(3).Control(1)=   "Frame1"
         Tab(3).ControlCount=   2
         TabCaption(4)   =   "Other Charges"
         TabPicture(4)   =   "frmNewCustomInvoice.frx":0070
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "lblCap(2)"
         Tab(4).Control(1)=   "lblCap(1)"
         Tab(4).Control(2)=   "lblCap(0)"
         Tab(4).Control(3)=   "LVOtherCharges"
         Tab(4).Control(4)=   "txtOtherCharges"
         Tab(4).Control(5)=   "cmdAddOtherCharges"
         Tab(4).Control(6)=   "txtOtherChargeAmt"
         Tab(4).Control(7)=   "txtOtherChargesSNo"
         Tab(4).ControlCount=   8
         Begin VB.CheckBox chkCIFTerms 
            Caption         =   "Show CIF"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9660
            TabIndex        =   149
            Top             =   6255
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.CheckBox chkNoSRO 
            Caption         =   "No SRO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10935
            TabIndex        =   148
            Top             =   6255
            Width           =   1020
         End
         Begin ComboList.Usercontrol1 cmbProformas 
            Height          =   285
            Left            =   -74700
            TabIndex        =   143
            Top             =   690
            Width           =   5055
            _ExtentX        =   8916
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
            TextForeColor   =   16711680
            ListIndex       =   -1
            Appearance      =   0
         End
         Begin VB.TextBox txtDTRENo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   -70350
            TabIndex        =   140
            Top             =   2205
            Visible         =   0   'False
            Width           =   1400
         End
         Begin VB.TextBox txtGatePassNo 
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   75
            MaxLength       =   150
            TabIndex        =   136
            Top             =   5325
            Width           =   2265
         End
         Begin VB.TextBox txtPrice 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   -72090
            TabIndex        =   126
            Top             =   795
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtOtherChargesSNo 
            Height          =   315
            Left            =   -74895
            TabIndex        =   101
            Top             =   1020
            Width           =   1200
         End
         Begin VB.TextBox txtOtherChargeAmt 
            Height          =   315
            Left            =   -66630
            TabIndex        =   100
            Top             =   1020
            Width           =   2000
         End
         Begin VB.CommandButton cmdAddOtherCharges 
            Caption         =   "Add"
            Height          =   315
            Left            =   -64595
            TabIndex        =   99
            Top             =   1020
            Width           =   1560
         End
         Begin VB.TextBox txtOtherCharges 
            Height          =   315
            Left            =   -73665
            TabIndex        =   98
            Top             =   1020
            Width           =   7005
         End
         Begin VB.Frame Frame1 
            Height          =   2670
            Left            =   -74565
            TabIndex        =   69
            Top             =   960
            Width           =   11475
            Begin VB.TextBox txtSealNo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   7365
               TabIndex        =   124
               Top             =   405
               Width           =   1620
            End
            Begin VB.TextBox txtContainerNo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   5730
               TabIndex        =   122
               Top             =   405
               Width           =   1620
            End
            Begin VB.TextBox txtComPayTerms 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   5775
               TabIndex        =   79
               Top             =   960
               Visible         =   0   'False
               Width           =   2820
            End
            Begin VB.TextBox txtVessel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   2895
               TabIndex        =   77
               Top             =   405
               Width           =   2820
            End
            Begin MSForms.ComboBox cmbPayment 
               Height          =   285
               Left            =   90
               TabIndex        =   131
               Top             =   405
               Width           =   2805
               VariousPropertyBits=   746604571
               BackColor       =   16777215
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "4948;503"
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
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Seal No."
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
               Index           =   29
               Left            =   7365
               TabIndex        =   125
               Top             =   165
               Width           =   1620
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Container No."
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
               Index           =   27
               Left            =   5730
               TabIndex        =   123
               Top             =   165
               Width           =   1620
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Drawn Under"
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
               Left            =   60
               MouseIcon       =   "frmNewCustomInvoice.frx":008C
               TabIndex        =   88
               Top             =   1650
               Width           =   5730
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Against"
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
               Left            =   60
               MouseIcon       =   "frmNewCustomInvoice.frx":01DE
               TabIndex        =   87
               Top             =   720
               Width           =   5730
            End
            Begin MSForms.TextBox txtComAgainst 
               Height          =   705
               Left            =   60
               TabIndex        =   86
               Tag             =   "DeclarationsGarments"
               Top             =   960
               Width           =   5730
               VariousPropertyBits=   -1400879077
               BorderStyle     =   1
               Size            =   "10107;1244"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox txtComDrawn 
               Height          =   720
               Left            =   60
               TabIndex        =   85
               Top             =   1890
               Width           =   5730
               VariousPropertyBits=   -1400879077
               BorderStyle     =   1
               Size            =   "10107;1270"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Payment Terms"
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
               Index           =   19
               Left            =   90
               TabIndex        =   80
               Top             =   165
               Width           =   2820
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Vessel"
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
               Index           =   18
               Left            =   2895
               TabIndex        =   78
               Top             =   165
               Width           =   2820
            End
         End
         Begin VB.TextBox txtLC 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   75
            Locked          =   -1  'True
            TabIndex        =   64
            Top             =   5835
            Width           =   1740
         End
         Begin VB.Frame Frame2 
            Caption         =   "Port Of Loading :"
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
            Height          =   795
            Left            =   6045
            TabIndex        =   57
            Top             =   810
            Width           =   5865
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Country                         City                                  Port"
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
               Index           =   9
               Left            =   75
               TabIndex        =   61
               Top             =   210
               Width           =   5745
            End
            Begin MSForms.ComboBox cmbLoadCountry 
               Height          =   285
               Left            =   75
               TabIndex        =   60
               Top             =   450
               Width           =   1800
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "3175;503"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbLoadCity 
               Height          =   285
               Left            =   1905
               TabIndex        =   59
               Top             =   450
               Width           =   1800
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "3175;503"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbLoadPort 
               Height          =   285
               Left            =   3735
               TabIndex        =   58
               Top             =   450
               Width           =   2085
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "3678;503"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Port Of Discharge :"
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
            Height          =   1455
            Left            =   6030
            TabIndex        =   52
            Top             =   1590
            Width           =   5880
            Begin VB.TextBox txtGate 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   3090
               TabIndex        =   95
               Top             =   960
               Width           =   1155
            End
            Begin VB.TextBox txtWarehouse 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1905
               TabIndex        =   94
               Top             =   960
               Width           =   1155
            End
            Begin VB.TextBox txtManDischarge 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   75
               TabIndex        =   92
               Top             =   960
               Width           =   1800
            End
            Begin MSComCtl2.DTPicker DTETA 
               Height          =   285
               Left            =   4260
               TabIndex        =   96
               Top             =   960
               Width           =   1590
               _ExtentX        =   2805
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
               CustomFormat    =   "dd-MM-yyyy"
               Format          =   300613635
               CurrentDate     =   37250
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Discharge Port            Warehouse     Gate                ETA"
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
               Index           =   28
               Left            =   75
               TabIndex        =   93
               Top             =   720
               Width           =   5745
            End
            Begin MSForms.ComboBox cmbDisPort 
               Height          =   285
               Left            =   3735
               TabIndex        =   56
               Top             =   450
               Width           =   2085
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "3678;503"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbDisCity 
               Height          =   285
               Left            =   1905
               TabIndex        =   55
               Top             =   450
               Width           =   1800
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "3175;503"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbDisCountry 
               Height          =   285
               Left            =   75
               TabIndex        =   54
               Top             =   450
               Width           =   1800
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "3175;503"
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
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Country                         City                                  Port"
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
               Index           =   6
               Left            =   75
               TabIndex        =   53
               Top             =   210
               Width           =   5745
            End
         End
         Begin VB.Frame Frame3 
            Height          =   1485
            Left            =   180
            TabIndex        =   41
            Top             =   810
            Width           =   5805
            Begin VB.TextBox txtAWBNo 
               Appearance      =   0  'Flat
               DataField       =   "RptTime"
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   75
               MaxLength       =   150
               TabIndex        =   44
               Top             =   450
               Width           =   1830
            End
            Begin VB.TextBox txtFormE 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   75
               TabIndex        =   43
               Top             =   1125
               Width           =   1380
            End
            Begin VB.TextBox txtGWeight 
               Appearance      =   0  'Flat
               DataField       =   "RptTime"
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   3900
               MaxLength       =   150
               TabIndex        =   42
               Top             =   450
               Width           =   1830
            End
            Begin MSComCtl2.DTPicker AWBDT 
               Height          =   300
               Left            =   1935
               TabIndex        =   45
               Top             =   450
               Width           =   1935
               _ExtentX        =   3413
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
               CheckBox        =   -1  'True
               CustomFormat    =   "dd-MMM-yyyy"
               Format          =   300613635
               CurrentDate     =   37250
            End
            Begin MSComCtl2.DTPicker FormEDT 
               Height          =   300
               Left            =   1455
               TabIndex        =   46
               Top             =   1125
               Width           =   1680
               _ExtentX        =   2963
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
               CheckBox        =   -1  'True
               CustomFormat    =   "dd-MMM-yyyy"
               Format          =   300613635
               CurrentDate     =   37250
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Partial               Trans"
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
               Index           =   25
               Left            =   3135
               TabIndex        =   68
               Top             =   870
               Width           =   2580
            End
            Begin MSForms.ComboBox cmbTrans 
               Height          =   285
               Left            =   4425
               TabIndex        =   67
               Top             =   1110
               Width           =   1290
               VariousPropertyBits=   746604571
               BackColor       =   16777215
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "2275;503"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbPartial 
               Height          =   285
               Left            =   3135
               TabIndex        =   66
               Top             =   1110
               Width           =   1290
               VariousPropertyBits=   746604571
               BackColor       =   16777215
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "2275;503"
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
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " AWB No."
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
               Index           =   10
               Left            =   75
               TabIndex        =   51
               Top             =   195
               Width           =   1830
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " AWB Date"
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
               Index           =   11
               Left            =   1935
               TabIndex        =   50
               Top             =   195
               Width           =   1935
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Form 'E'"
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
               Index           =   5
               Left            =   75
               TabIndex        =   49
               Top             =   870
               Width           =   1380
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Form 'E' Date"
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
               Index           =   12
               Left            =   1470
               TabIndex        =   48
               Top             =   870
               Width           =   1680
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Gross Weight"
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
               Index           =   20
               Left            =   3900
               TabIndex        =   47
               Top             =   195
               Width           =   1830
            End
         End
         Begin VB.Frame Frame5 
            Height          =   2460
            Left            =   6030
            TabIndex        =   32
            Top             =   3090
            Width           =   5910
            Begin ComboList.Usercontrol1 cmbInsurAgent 
               Height          =   285
               Left            =   3015
               TabIndex        =   33
               Top             =   420
               Width           =   2835
               _ExtentX        =   5001
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
            Begin ComboList.Usercontrol1 cmbShipAgent 
               Height          =   285
               Left            =   75
               TabIndex        =   34
               Top             =   420
               Width           =   2910
               _ExtentX        =   5133
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
            Begin MSForms.ComboBox cmbConsignee 
               Height          =   285
               Left            =   75
               TabIndex        =   84
               Top             =   960
               Width           =   2910
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "5133;503"
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
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " To"
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
               Index           =   26
               Left            =   75
               TabIndex        =   83
               Top             =   720
               Width           =   2910
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Insurance Agent"
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
               Index           =   13
               Left            =   3015
               TabIndex        =   40
               Top             =   180
               Width           =   2835
            End
            Begin MSForms.ComboBox cmbThrough 
               Height          =   285
               Left            =   3015
               TabIndex        =   39
               Top             =   960
               Width           =   2835
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "5001;503"
               ListWidth       =   7055
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
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Through"
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
               Index           =   4
               Left            =   3015
               TabIndex        =   38
               Top             =   720
               Width           =   2835
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Consignee"
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
               Index           =   17
               Left            =   75
               TabIndex        =   37
               Top             =   1260
               Width           =   5775
            End
            Begin MSForms.TextBox txtBankAddress 
               Height          =   855
               Left            =   75
               TabIndex        =   36
               Top             =   1500
               Width           =   5775
               VariousPropertyBits=   -1400879077
               BorderStyle     =   1
               Size            =   "10186;1508"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Shipping Agent"
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
               Left            =   75
               TabIndex        =   35
               Top             =   180
               Width           =   2910
            End
         End
         Begin VB.CommandButton Command1 
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   -62895
            Picture         =   "frmNewCustomInvoice.frx":0330
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Remove This Entry From Voucher"
            Top             =   1245
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.TextBox txtQty1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   -68130
            TabIndex        =   26
            Top             =   285
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtQty 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   -70800
            TabIndex        =   25
            Top             =   795
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.CommandButton cmdDelete 
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   -63345
            Picture         =   "frmNewCustomInvoice.frx":0862
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Remove This Entry From Voucher"
            Top             =   1305
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   345
         End
         Begin MSComctlLib.ListView LV 
            Height          =   5640
            Left            =   -74700
            TabIndex        =   18
            Top             =   720
            Width           =   11340
            _ExtentX        =   20003
            _ExtentY        =   9948
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
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
            NumItems        =   14
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S No."
               Object.Width           =   1014
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Cust. Ref. No."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Order #"
               Object.Width           =   1720
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Size"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Color"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Article #"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Article Name"
               Object.Width           =   3941
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Item No."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Text            =   "Price"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   9
               Text            =   "Qty"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   10
               Text            =   "Amount"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   11
               Text            =   "DTRE #"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   12
               Text            =   "Batch #"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Text            =   "Delivery"
               Object.Width           =   2364
            EndProperty
         End
         Begin MSComctlLib.ListView LV2 
            Height          =   5415
            Left            =   -74700
            TabIndex        =   138
            Top             =   975
            Width           =   11685
            _ExtentX        =   20611
            _ExtentY        =   9551
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
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
            NumItems        =   12
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S No."
               Object.Width           =   1014
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Cust Ref No."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Order #"
               Object.Width           =   1720
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Size"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Color"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Article #"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Article Name"
               Object.Width           =   4586
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Item No."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Text            =   "Price"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   9
               Text            =   "Qty"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   10
               Text            =   "Amount"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "Delivery"
               Object.Width           =   2364
            EndProperty
         End
         Begin VB.Frame Frame7 
            Height          =   2505
            Left            =   -74760
            TabIndex        =   70
            Top             =   3615
            Width           =   11685
            Begin VB.CommandButton cmdSwap 
               Caption         =   "<< Swap>>"
               Height          =   240
               Left            =   2955
               TabIndex        =   76
               Top             =   150
               Width           =   1815
            End
            Begin VB.Label lblSwapAddresses 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Other Address"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   195
               Left            =   75
               MouseIcon       =   "frmNewCustomInvoice.frx":0D94
               MousePointer    =   99  'Custom
               TabIndex        =   97
               Top             =   2175
               Width           =   1065
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Speical Notes"
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
               Left            =   7380
               MouseIcon       =   "frmNewCustomInvoice.frx":0EE6
               TabIndex        =   90
               Top             =   150
               Width           =   4215
            End
            Begin MSForms.TextBox txtComSpecial 
               Height          =   1680
               Left            =   7380
               TabIndex        =   89
               Top             =   390
               Width           =   4215
               VariousPropertyBits=   -1400879077
               BorderStyle     =   1
               Size            =   "7435;2963"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox txtComNotify 
               Height          =   1680
               Left            =   3720
               TabIndex        =   75
               Top             =   390
               Width           =   3675
               VariousPropertyBits=   -1400879077
               BorderStyle     =   1
               Size            =   "6482;2963"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Notify"
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
               Index           =   14
               Left            =   4770
               TabIndex        =   74
               Top             =   150
               Width           =   2625
            End
            Begin MSForms.TextBox txtComConsignee 
               Height          =   1680
               Left            =   45
               TabIndex        =   73
               Top             =   390
               Width           =   3690
               VariousPropertyBits=   -1400879077
               BorderStyle     =   1
               Size            =   "6509;2963"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Consignee"
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
               Index           =   15
               Left            =   45
               TabIndex        =   72
               Top             =   150
               Width           =   1170
            End
            Begin MSForms.ComboBox cmbComConsignee 
               Height          =   255
               Left            =   1200
               TabIndex        =   71
               Top             =   150
               Width           =   1755
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "3096;450"
               ListWidth       =   7055
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin MSComctlLib.ListView LVOtherCharges 
            Height          =   5280
            Left            =   -74895
            TabIndex        =   102
            Top             =   1380
            Width           =   11865
            _ExtentX        =   20929
            _ExtentY        =   9313
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
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
               Text            =   "S No."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description"
               Object.Width           =   10583
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Amount"
               Object.Width           =   3528
            EndProperty
         End
         Begin MSComCtl2.DTPicker DTLC 
            Height          =   300
            Left            =   1845
            TabIndex        =   130
            Top             =   5835
            Width           =   1770
            _ExtentX        =   3122
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
            CheckBox        =   -1  'True
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   300679171
            CurrentDate     =   37250
         End
         Begin MSComCtl2.DTPicker DTGatePass 
            Height          =   300
            Left            =   2340
            TabIndex        =   137
            Top             =   5325
            Width           =   1770
            _ExtentX        =   3122
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
            CheckBox        =   -1  'True
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   300679171
            CurrentDate     =   37250
         End
         Begin VB.Label lblGoTo 
            AutoSize        =   -1  'True
            Caption         =   "Search Item : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   1
            Left            =   -74670
            TabIndex        =   145
            Top             =   6420
            Width           =   1185
         End
         Begin MSForms.TextBox txtDTREDesc 
            Height          =   1335
            Left            =   135
            TabIndex        =   142
            Top             =   3765
            Width           =   5865
            VariousPropertyBits=   -1400879077
            BorderStyle     =   1
            Size            =   "10345;2355"
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
            Caption         =   " DTRE Description"
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
            Index           =   32
            Left            =   135
            TabIndex        =   141
            Top             =   3525
            Width           =   5865
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Gate-Pass #                               Date"
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
            Index           =   31
            Left            =   75
            TabIndex        =   135
            Top             =   5085
            Width           =   4005
         End
         Begin VB.Label lblGoTo 
            AutoSize        =   -1  'True
            Caption         =   "Search Item : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   0
            Left            =   -74700
            TabIndex        =   134
            Top             =   6435
            Width           =   1185
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Bank Account No."
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
            Index           =   22
            Left            =   8235
            TabIndex        =   133
            Top             =   5595
            Width           =   3720
         End
         Begin MSForms.ComboBox cmbBankAccNo 
            Height          =   285
            Left            =   8235
            TabIndex        =   132
            Top             =   5835
            Width           =   3720
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "6562;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CommandButton cmdAddAll 
            Height          =   345
            Left            =   -64530
            TabIndex        =   129
            Top             =   630
            Width           =   1530
            Caption         =   "   Add All"
            PicturePosition =   327683
            Size            =   "2699;609"
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
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Invoice Of"
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
            Index           =   30
            Left            =   135
            TabIndex        =   128
            Top             =   2325
            Width           =   5865
         End
         Begin MSForms.TextBox txtInvoiceof 
            Height          =   975
            Left            =   135
            TabIndex        =   127
            Top             =   2565
            Width           =   5865
            VariousPropertyBits=   -1400879077
            BorderStyle     =   1
            Size            =   "10345;1720"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            Caption         =   "S/No. :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   -74850
            TabIndex        =   105
            Top             =   795
            Width           =   510
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            Caption         =   "Charges Description :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   -73665
            TabIndex        =   104
            Top             =   795
            Width           =   1545
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            Caption         =   "Amount"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   -66615
            TabIndex        =   103
            Top             =   780
            Width           =   555
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Our Bank"
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
            Index           =   24
            Left            =   5415
            TabIndex        =   82
            Top             =   5595
            Width           =   2835
         End
         Begin MSForms.ComboBox cmbBank 
            Height          =   285
            Left            =   5415
            TabIndex        =   81
            Top             =   5835
            Width           =   2835
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5001;503"
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
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " LC #                               LC Date"
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
            Index           =   23
            Left            =   75
            TabIndex        =   65
            Top             =   5595
            Width           =   3570
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Shipping Method"
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
            Index           =   21
            Left            =   3630
            TabIndex        =   63
            Top             =   5595
            Width           =   1800
         End
         Begin MSForms.ComboBox cmbSMethod 
            Height          =   285
            Left            =   3630
            TabIndex        =   62
            Top             =   5835
            Width           =   1800
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "3175;503"
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
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Proforma No. && Orders"
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
            Index           =   8
            Left            =   -74700
            TabIndex        =   31
            Top             =   450
            Width           =   5055
         End
         Begin MSForms.ComboBox cmbProforma 
            Height          =   285
            Left            =   -74700
            TabIndex        =   30
            Top             =   690
            Visible         =   0   'False
            Width           =   5055
            VariousPropertyBits=   746604571
            BackColor       =   16777215
            ForeColor       =   6244673
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "8916;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label lblTotPcs 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E7EBEF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   -65160
            TabIndex        =   29
            Top             =   6345
            Width           =   1800
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Total Pcs :"
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
            Height          =   270
            Index           =   16
            Left            =   -66150
            TabIndex        =   28
            Top             =   6345
            Width           =   1005
         End
         Begin MSForms.CommandButton cmdAddItem 
            Height          =   345
            Left            =   -69600
            TabIndex        =   144
            Top             =   630
            Width           =   1515
            Caption         =   "    Add"
            PicturePosition =   327683
            Size            =   "2672;609"
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAddItem1 
            Height          =   345
            Left            =   -72855
            TabIndex        =   22
            Top             =   2535
            Visible         =   0   'False
            Width           =   1515
            Caption         =   "    Add"
            PicturePosition =   327683
            Size            =   "2672;609"
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
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Order No."
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
            Left            =   -74640
            TabIndex        =   21
            Top             =   2370
            Visible         =   0   'False
            Width           =   1755
         End
         Begin MSForms.ComboBox cmbOrders1 
            Height          =   285
            Left            =   -74640
            TabIndex        =   20
            Top             =   2595
            Visible         =   0   'False
            Width           =   1755
            VariousPropertyBits=   746604571
            BackColor       =   16777215
            ForeColor       =   6244673
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "3096;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
      End
      Begin VB.Frame Frame6 
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
         Height          =   660
         Left            =   5250
         TabIndex        =   10
         Top             =   8880
         Width           =   6960
         Begin MSForms.CommandButton cmdSavenStay 
            Height          =   360
            Left            =   3495
            TabIndex        =   147
            Top             =   195
            Width           =   1665
            Caption         =   "Save & Stay"
            PicturePosition =   327683
            Size            =   "2937;635"
            Accelerator     =   116
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
            Left            =   5190
            TabIndex        =   8
            Top             =   195
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
         Begin MSForms.CommandButton cmdSClear 
            Height          =   360
            Left            =   90
            TabIndex        =   6
            Top             =   195
            Width           =   1665
            Caption         =   "Save & New   "
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
         Begin MSForms.CommandButton cmdSClose 
            Height          =   360
            Left            =   1800
            TabIndex        =   7
            Top             =   195
            Width           =   1665
            Caption         =   "Save & Close  "
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Company Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   106
         Top             =   840
         Width           =   1410
      End
      Begin MSForms.TextBox lblTotWords 
         Height          =   270
         Left            =   1890
         TabIndex        =   23
         Top             =   8550
         Width           =   8070
         VariousPropertyBits=   746604575
         BackColor       =   15199215
         BorderStyle     =   1
         Size            =   "14235;476"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Amount :"
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
         Height          =   270
         Index           =   1
         Left            =   9945
         TabIndex        =   16
         Top             =   8550
         Width           =   885
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Amount (In Words) :"
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
         Height          =   270
         Index           =   2
         Left            =   90
         TabIndex        =   15
         Top             =   8550
         Width           =   1830
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   10815
         TabIndex        =   14
         Top             =   8550
         Width           =   1380
      End
      Begin MSForms.CheckBox chkPrint 
         Height          =   345
         Left            =   90
         TabIndex        =   13
         Top             =   8910
         Width           =   1545
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   4
         Size            =   "2725;609"
         Value           =   "0"
         Caption         =   "Print On Save"
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.CheckBox chkCust 
         Height          =   360
         Left            =   90
         TabIndex        =   12
         Top             =   9180
         Visible         =   0   'False
         Width           =   2685
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   4
         Size            =   "4736;635"
         Value           =   "1"
         Caption         =   "Show Customer Item Codes"
         Accelerator     =   73
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbTradeTerms 
         Height          =   285
         Left            =   8895
         TabIndex        =   5
         Top             =   1410
         Width           =   1500
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2646;503"
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
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmNewCustomInvoice.frx":1038
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
         Left            =   90
         TabIndex        =   11
         Top             =   1155
         Width           =   11940
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   2415
         TabIndex        =   1
         Top             =   1410
         Width           =   2145
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3784;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Hidden"
      Height          =   8055
      Left            =   3600
      TabIndex        =   107
      Top             =   705
      Visible         =   0   'False
      Width           =   6285
      Begin VB.Label lblGlovesComDecs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Declarations For Gloves"
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
         Left            =   315
         MouseIcon       =   "frmNewCustomInvoice.frx":10F4
         MousePointer    =   99  'Custom
         TabIndex        =   121
         Top             =   5670
         Width           =   5730
      End
      Begin VB.Label lblGarmentsComDecs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Declarations For Garments"
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
         Left            =   315
         MouseIcon       =   "frmNewCustomInvoice.frx":1246
         MousePointer    =   99  'Custom
         TabIndex        =   120
         Top             =   4860
         Width           =   5730
      End
      Begin VB.Label lblBallComDecs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Declarations For Footballs"
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
         Left            =   315
         MouseIcon       =   "frmNewCustomInvoice.frx":1398
         MousePointer    =   99  'Custom
         TabIndex        =   119
         Top             =   4050
         Width           =   5730
      End
      Begin MSForms.TextBox txtComDeclarationGarments 
         Height          =   585
         Left            =   315
         TabIndex        =   118
         Tag             =   "DeclarationsGarments"
         Top             =   5100
         Width           =   5730
         VariousPropertyBits=   -1400879077
         BorderStyle     =   1
         Size            =   "10107;1032"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtComDeclarationBalls 
         Height          =   585
         Left            =   315
         TabIndex        =   117
         Top             =   4290
         Width           =   5730
         VariousPropertyBits=   -1400879077
         BorderStyle     =   1
         Size            =   "10107;1032"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtComDeclarationGloves 
         Height          =   585
         Left            =   315
         TabIndex        =   116
         Top             =   5910
         Width           =   5730
         VariousPropertyBits=   -1400879077
         BorderStyle     =   1
         Size            =   "10107;1032"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblGlovesDecs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Declarations For Gloves"
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
         Left            =   45
         MouseIcon       =   "frmNewCustomInvoice.frx":14EA
         MousePointer    =   99  'Custom
         TabIndex        =   115
         Top             =   1815
         Width           =   5805
      End
      Begin VB.Label lblGarmentsDecs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Declarations For Garments"
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
         Left            =   45
         MouseIcon       =   "frmNewCustomInvoice.frx":163C
         MousePointer    =   99  'Custom
         TabIndex        =   114
         Top             =   1050
         Width           =   5805
      End
      Begin VB.Label lblBallDecs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Declarations For Footballs"
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
         Left            =   45
         MouseIcon       =   "frmNewCustomInvoice.frx":178E
         MousePointer    =   99  'Custom
         TabIndex        =   113
         Top             =   285
         Width           =   5805
      End
      Begin MSForms.TextBox txtDeclarationGarments 
         Height          =   540
         Left            =   45
         TabIndex        =   112
         Tag             =   "DeclarationsGarments"
         Top             =   1290
         Width           =   5805
         VariousPropertyBits=   -1400879077
         BorderStyle     =   1
         Size            =   "10239;952"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtDeclarationBalls 
         Height          =   540
         Left            =   45
         TabIndex        =   111
         Top             =   525
         Width           =   5805
         VariousPropertyBits=   -1400879077
         BorderStyle     =   1
         Size            =   "10239;952"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtDeclarationGloves 
         Height          =   540
         Left            =   45
         TabIndex        =   110
         Top             =   2055
         Width           =   5805
         VariousPropertyBits=   -1400879077
         BorderStyle     =   1
         Size            =   "10239;952"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Manual Declaration :"
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
         Left            =   120
         MouseIcon       =   "frmNewCustomInvoice.frx":18E0
         MousePointer    =   99  'Custom
         TabIndex        =   109
         Top             =   2595
         Width           =   1875
      End
      Begin MSForms.TextBox txtManDeclaration 
         Height          =   915
         Left            =   1980
         TabIndex        =   108
         Tag             =   "ManDeclaration"
         Top             =   2595
         Width           =   3945
         VariousPropertyBits=   -1400879077
         BorderStyle     =   1
         Size            =   "6959;1614"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frmNewCustomInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Orders As New Collection
Dim strProforma As String
Dim bCustomInvoice As Boolean
Dim strDelList As String
Dim bCourierInvoice As Boolean
Dim iColNo As Integer
Dim bAcceptsExtraQty As Boolean

Public Sub EditProforma(TempProforma As String, Optional p_bCustomInvoice As Boolean = True)
         
    On Error GoTo err
    
    bCustomInvoice = p_bCustomInvoice
    
    If bCustomInvoice Then
        label1(33).Caption = "New Custom Invoice"
        label1(34).Caption = "New Custom Invoice"
    Else
        label1(33).Caption = "New Commercial Invoice"
        label1(34).Caption = "New Commercial Invoice"
    End If
    
    Me.Hide
    
    strProforma = TempProforma
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim dPrice As Double
    With rs
        .Open "SELECT * FROM VCustomInvoice WHERE CustomInvoice='" & strProforma & "'", con_ServerSide, adOpenForwardOnly, adLockReadOnly
        cmbCompany.ID = !CompanyRefID
        cmbCust.Text = ![CustCode] & ""
        cmbCountry = ![Country] & ""
        OrderDT = ![DT] & ""
        txtPInvoice = ![CustomInvoice] & ""
        
        cmbTradeTerms = ![TradeTerms] & ""
        
        cmbPayment = ![PaymentTerms] & ""
        
        txtLC = ![LCNo] & ""
        If IsNull(!LCDate) = False Then
            If !LCDate <> #1/1/1900# Then
                DTLC = !LCDate
            End If
        End If
         
        If Val(!OurBankID & "") > 0 Then
            
            cmbBank = !Bank & ""
            For i = 0 To cmbBankAccNo.ListCount - 1
                If cmbBankAccNo.List(i, 1) = !OurBankID & "" Then
                    cmbBankAccNo.ListIndex = i
                    Exit For
                End If
            Next
        End If
        
        cmbTrans = ![TransShipment] & ""
        cmbPartial = ![PartialShipment] & ""
        txtPackaging = ![Packaging] & ""
        cmbSMethod = ![ShippingMode] & ""
        txtAWBNo = ![AWBNo] & ""
        AWBDT = ![AWBNoDT]
        txtFormE = ![FormE] & ""
        FormEDT = ![FormEDT] & ""
        
        
        cmbShipAgent.ID = ![ShippingAgentID] & ""
        cmbInsurAgent.ID = ![InsuranceAgentID] & ""
        
        If IsNull(![LoadPortID]) = False Then
            cmbLoadCountry = ![LoadCountry] & ""
            cmbLoadCity = ![LoadCity] & ""
            cmbLoadPort = ![LoadPort] & ""
        Else
            cmbLoadCountry.Text = "Pakistan"
        End If
        
        If IsNull(![DisPortID]) = False Then
            cmbDisCountry = ![DisCountry] & ""
            cmbDisCity = ![DisCity] & ""
            cmbDisPort = ![DisPort] & ""
        End If
                
        txtBankAddress = ![Consignee] & ""
        txtDeclarationBalls = ![DeclarationsBalls] & ""
        txtDeclarationGarments = ![DeclarationsGarments] & ""
        txtDeclarationGloves = ![DeclarationsGloves] & ""
        
        txtManDischarge = ![ManDischarge] & ""
        txtGate = !Gate & ""
        txtWarehouse = !WareHouse & ""
        txtContainerNo = !ContainerNo & ""
        txtSealNo = !SealNo & ""
        
        If IsNull(!DTETA) = False Then
            If !DTETA <> #1/1/1900# Then
                DTETA = !DTETA
            End If
        End If
         
        'Com Entries
        txtComDeclarationBalls = ![ComDeclarationsBalls] & ""
        txtComDeclarationGarments = ![ComDeclarationsGarments] & ""
        txtComDeclarationGloves = ![ComDeclarationsGloves] & ""
        txtVessel = ![Vessel] & ""
        txtComConsignee = ![ComConsignee] & ""
        txtComNotify = ![ComNotify] & ""
        txtComPayTerms = ![ComPayTerms] & ""
        txtComAgainst = ![ComAgainst] & ""
        txtComDrawn = ![ComDrawnUnder] & ""
        txtComSpecial = ![ComSpecial] & ""
        txtManDeclaration = ![ManDeclaration] & ""
        txtInvoiceof = !InvoiceOf & ""
        
        txtGatePassNo = !GatePassNo & ""
        If IsNull(!GatePassDT) = False Then
            DTGatePass = !GatePassDT
        End If
        
        txtDTREDesc = !DTREDescription & ""
        txtGWeight = Val(!GrossWeight & "")
        
        chkNoSRO = Abs(!NoSRO)
        
        .Close
        
        chkCIFTerms = Abs(GetSingleBooleanValue("CIFTradeTerms", "CustomInvoice", " WHERE CustomInvoice='" & strProforma & "'"))
        
        Dim strOrderBy As String
        If strCompany = "QEL" Then
            strOrderBy = "ORDER BY CompItemCode"
        Else
            strOrderBy = "ORDER BY EntryID"
        End If
        
        .Open "SELECT * FROM VCustomInvoiceItems WHERE CustomInvoice='" & strProforma & "' " & strOrderBy, con_ServerSide, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, ![RefID] & "'", LV.ListItems.count + 1)
            ITM.Tag = ![ITemID] & "" 'cmbItems.List(cmbItems.ListIndex, 2)
            
            ITM.ListSubItems.add(, , ![ITemID]).Tag = Val(![EntryID] & "") '![CatID] & ""
            ITM.ListSubItems.add , , ![OrderNo] '![CatID] & ""
            ITM.ListSubItems.add , , ![ItemSize] & "" '![ItemID] & ""
            ITM.ListSubItems.add , , "" '![Color] & ""
            ITM.ListSubItems.add(, , "").Tag = "" '![ArticleNo] & ""
            If bCustomInvoice Then
                ITM.ListSubItems.add(, ![ItemCustomName] & "'", ![ItemCustomName] & "").Tag = ![ItemCustomName] & ""
                dPrice = Val(!CustomPrice & "")
            Else
                ITM.ListSubItems.add(, ![ItemName] & "'", ![ItemName] & "").Tag = ![ItemName] & ""
                dPrice = Val(!Price & "")
            End If
            
            ITM.ListSubItems.add , , !CompItemCode & "" '![Unit] & ""
            ITM.ListSubItems.add , , Format(dPrice, "#0.0000")
            ITM.ListSubItems.add , , Val(![Qty] & "")
            
            ITM.ListSubItems(8).Tag = Val(![Qty] & "") + Val(![InvQty] & "")
            
            ITM.SubItems(10) = Val(![Qty] & "") * dPrice
            
            ITM.SubItems(11) = !DTRENo & ""
            ITM.ListSubItems(11).Tag = !ItemForLabTest
            
            ITM.SubItems(12) = !BatchNo & ""
            
            If strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
                ITM.SubItems(13) = Format(!DeliveryDTItem, "dd-MMM-yyyy")
                ITM.ListSubItems(13).Tag = Format(!DeliveryDTItem, "dd/MM/yy")
            End If
            
            .MoveNEXT
        Loop
        .Close
        
        '.Open "Select * From CustomInvoiceOtherCharges WHERE CustomInvoice='" & strProforma & "' AND Custom=" & Abs(bCustomInvoice) & "ORDER BY SNo", con_ServerSide, adOpenForwardOnly, adLockReadOnly
        .Open "Select * From CustomInvoiceOtherCharges WHERE CustomInvoice='" & strProforma & "' ORDER BY SNo", con_ServerSide, adOpenForwardOnly, adLockReadOnly
        LVOtherCharges.ListItems.Clear
        Do Until .EOF
            Set ITM = LVOtherCharges.ListItems.add(, , Val(!SNo & ""))
            ITM.SubItems(1) = !ChargesDetail & ""
            ITM.SubItems(2) = Val(!ChargesAmt & "")
            .MoveNEXT
        Loop
        .Close
        txtOtherChargesSNo.Text = LVOtherCharges.ListItems.count + 1
        
    End With
    Call CalculateTotals
    'Call UpdateItemSearchInvoice
    cmbItemSearchInvoice.DropDownOffLine = True
    'cmbItemSearchInvoice.ClearVals
    If strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
        cmbItemSearchInvoice.AddVals con_ServerSide, "ItemID + ' {' + OrderNo + '} {' + CONVERT(varchar(50),DeliveryDTItem,3) + '}'", "VCustomInvoiceItems", "ItemID+'|'+OrderNo+'|'+CONVERT(varchar(50),DeliveryDTItem,3)", "WHERE CustomInvoice='" & strProforma & "' " & strOrderBy
    Else
        cmbItemSearchInvoice.AddVals con_ServerSide, "ItemID + ' {' + OrderNo + '}'", "VCustomInvoiceItems", "ItemID+'|'+OrderNo", "WHERE CustomInvoice='" & strProforma & "' " & strOrderBy
    End If
    'txtQty = 0
    'cmbItems.SetFocus
    
    Set rs = Nothing
    Me.Show
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub AWBDT_Change()
    If cmbCust.ID = "DECATHLON" Then
        DTETA = DateAdd("d", 20, AWBDT.Value)
    End If
End Sub

Private Sub chkCust_Click()
    
    Exit Sub
    'If cmbCatID.MatchFound Then
    '    Call FillItems
    '    For i = 1 To LV.ListItems.Count
    '        If chkCust Then
    '            LV.ListItems(i).SubItems(2) = LV.ListItems(i).Key
    '            LV.ListItems(i).ListSubItems(3) = LV.ListItems(i).ListSubItems(3).Key
    '        Else
    '            LV.ListItems(i).SubItems(2) = LV.ListItems(i).Tag
    '            LV.ListItems(i).ListSubItems(3) = LV.ListItems(i).ListSubItems(3).Tag
    '        End If
    '    Next i
    'End If
    
End Sub

Private Sub cmbBank_Change()
    Call cmbBank_Click
End Sub

Private Sub cmbBank_Click()
    cmbBankAccNo.Clear
    Call AddToCombo(cmbBankAccNo, "AccNo + ' ( ' + Branch + ' )'", "BankList", " WHERE CompanyRefID=" & cmbCompany.ID & " AND Bank='" & cmbBank.Text & "'", , "BankID")
    If cmbBankAccNo.ListCount > 0 Then cmbBankAccNo.ListIndex = 0
End Sub

Private Sub cmbComConsignee_Change()
    Call cmbComConsignee_Click
End Sub

Private Sub cmbComConsignee_Click()
    If cmbComConsignee.MatchFound = False Then Exit Sub
    txtComConsignee = cmbComConsignee.List(cmbComConsignee.ListIndex, 1)
End Sub

Private Sub cmbCompany_matched()

    If strProforma <> "" Then Exit Sub
    If strCompany = "Banzai" Or strCompany = "Instrumed" Then
        Call OrderDT_Change
    ElseIf Not (strCompany = "QEL" Or strCompany = "Tecno") Then
        Dim InvoiecNo As String
        InvoiceNo = GetInvoiceNo(cmbCompany.ID)
        
        If bCourierInvoice Then
            txtPInvoice = strCourierInvoicePrefix & " " & Format(InvoiceNo, "000")
        Else
            txtPInvoice.Tag = InvoiceNo
            txtPInvoice = txtPInvoice.Tag
        End If
    End If
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT DISTINCT Bank FROM BankList WHERE CompanyRefID=" & cmbCompany.ID, con, adOpenForwardOnly, adLockReadOnly
        cmbBank.Clear
        Do Until .EOF
            cmbBank.AddItem .Fields(0) & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbConsignee_Change()
    Call cmbConsignee_Click
End Sub

Private Sub cmbConsignee_Click()

    If cmbConsignee = "Customer Address" Then
        txtBankAddress = txtBankAddress.Tag
        cmbThrough.Text = ""
        cmbThrough.Enabled = False
    ElseIf cmbConsignee = "Bank Address" Then
        cmbThrough.Enabled = True
        If cmbThrough.ListCount > 0 Then cmbThrough.ListIndex = 0
    End If
    
End Sub

Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()
    
    On Error GoTo err
    If cmbCust.MatchFound = False Then Exit Sub
    If cmbCountry.MatchFound = False Then Exit Sub
    
    txtCurrency = cmbCountry.List(cmbCountry.ListIndex, 1)
    Call GetOrders
    Call GetBanks
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select FakeAddress From ForeignCustomers Where CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry & "'", con, adOpenForwardOnly, adLockReadOnly
        txtBankAddress.Tag = ![FakeAddress] & ""
        .Close
        .Open "SELECT AcceptsExtraQty FROM ForeignCustomers WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        bAcceptsExtraQty = !AcceptsExtraQty
        .Close
    End With
    cmbDisCountry = cmbCountry
    Set rs = Nothing
   
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub GetBanks()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select BankID,BankName,Address2 From FCustBanks Where CustCode+Country='" & cmbCust.Text & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbThrough.Clear
        Do Until .EOF
            cmbThrough.AddItem ![BankName] & ""
            cmbThrough.List(cmbThrough.ListCount - 1, 1) = ![BankID]
            cmbThrough.List(cmbThrough.ListCount - 1, 2) = ![Address2] & ""
            .MoveNEXT
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
        .Open "Select Country,Curr,Address From ForeignCustomers Where CustCode='" & cmbCust.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
        cmbComConsignee.Clear
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            cmbComConsignee.AddItem ![Country] & ""
            cmbComConsignee.List(.AbsolutePosition - 1, 1) = ![Address] & ""
            cmbCountry.List(cmbCountry.ListCount - 1, 1) = ![Curr] & ""
            'cmbDisCountry.AddItem ![Country] & ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    Call AddToCombo(cmbDisCountry, "CountryName", "Countries")
    
    If cmbCountry.ListCount > 0 Then cmbCountry.ListIndex = 0
    Set rs = Nothing
    
    If strProforma = "" Then
        If strCompany = "Sunlike" Then
            txtPInvoice = Left(txtPInvoice, 8) & Mid(cmbCust.ID, 4)
        ElseIf strCompany = "Dr-Frgz" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
            txtPInvoice = txtPInvoice.Tag & "/" & cmbCust.ID
        ElseIf strCompany = "QEL" Or strCompany = "Tecno" Then
            'Do Nothing,It's not Customer Specific..
        ElseIf strCompany <> "Banzai" And strCompany <> "Instrumed" Then
            txtPInvoice = Left(txtPInvoice, 8) & Right(cmbCust.ID, 4)
        End If
'        If Len(txtPInvoice) > 8 Then
'            txtPInvoice = Left(txtPInvoice, 8) & Right(cmbCust.ID, 4)
'        Else
'            txtPInvoice = Left(txtPInvoice, 8) & Right(cmbCust.ID, 4)
'        End If
    End If
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbCust_UnMatched()
    'cmbCountry.ClearVals
End Sub


Private Sub cmbItemSearch_matched()
    Dim strOrderNo As String, strItemCode As String
'    strOrderNo = cmbItemSearch.List(cmbItemSearch.ListIndex, 2)
'    strItemCode = cmbItemSearch.List(cmbItemSearch.ListIndex, 1)
    Dim i As Integer, strSearch As String
    For i = 1 To LV2.ListItems.count
        If strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
            strSearch = LV2.ListItems(i).SubItems(1) & "|" & LV2.ListItems(i).SubItems(2) & "|" & LV2.ListItems(i).SubItems(7) & "|" & LV2.ListItems(i).ListSubItems(11).Tag
        Else
            strSearch = LV2.ListItems(i).SubItems(1) & "|" & LV2.ListItems(i).SubItems(2) & "|" & LV2.ListItems(i).SubItems(7)
        End If
        If strSearch = cmbItemSearch.ID Then
            LV2.ListItems(i).Selected = True
            LV2.ListItems(i).EnsureVisible
            LV2.SetFocus
            Call LV2_DblClick
            Exit For
        End If
    Next
End Sub

Private Sub cmbItemSearchInvoice_GotFocus()
    cmbItemSearchInvoice.ID = ""
End Sub

Private Sub cmbItemSearchInvoice_matched()

    Dim strOrderNo As String, strItemCode As String
    Dim i As Integer, strSearch As String
    For i = 1 To LV.ListItems.count
        If strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
            strSearch = LV.ListItems(i).SubItems(1) & "|" & LV.ListItems(i).SubItems(2) & "|" & LV.ListItems(i).ListSubItems(13).Tag
        Else
            strSearch = LV.ListItems(i).SubItems(1) & "|" & LV.ListItems(i).SubItems(2)
        End If
        If strSearch = cmbItemSearchInvoice.ID Then
            LV.ListItems(i).Selected = True
            LV.ListItems(i).EnsureVisible
            LV.SetFocus
            Exit For
        End If
    Next
    
End Sub

Private Sub cmbPayment_Change()
    Call cmbPayment_Click
End Sub

Private Sub cmbPayment_Click()

    If LCase(cmbPayment.Text) = "letter of credit" Then
        txtLC.Locked = False
    Else
        txtLC.Locked = True
        txtLC.Text = ""
    End If
    
End Sub

Private Sub cmbSMethod_Change()
    Call cmbSMethod_Click
End Sub

Private Sub cmbSMethod_Click()
    If cmbSMethod.ListIndex = 0 Then
        'Label1(12).Caption = "AWB No."
    Else
        'Label1(12).Caption = "BL No."
    End If
End Sub

Private Sub cmbTradeTerms_Change()
    Call cmbTradeTerms_Click
End Sub

Private Sub cmbTradeTerms_Click()
    
    
    If cmbTradeTerms.ListIndex = 1 Or cmbTradeTerms.ListIndex = 2 Then
        cmbInsurAgent.Enabled = True
    Else
        cmbInsurAgent.Enabled = False
    End If
    
    If cmbTradeTerms.Text Like "*Air*" Then
        cmbSMethod.ListIndex = 0
    ElseIf cmbTradeTerms.Text Like "*Sea*" Then
        cmbSMethod.ListIndex = 1
    End If
    
    Exit Sub
    
    If cmbTradeTerms.MatchFound Then
    
        cmbCatID.Locked = False
        cmbItems.Locked = False
        cmbDesc.Locked = False
        
        If cmbItems.MatchFound Then
            txtPrice = cmbItems.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 3)
        End If
        
        
        
    Else
        cmbCatID.Locked = True
        cmbItems.Locked = True
        cmbDesc.Locked = True
    End If
    
End Sub


Private Sub cmdAddAll_Click()

    If LV2.ListItems.count = 0 Then Exit Sub
    Dim i As Integer
    For i = 1 To LV2.ListItems.count
        LV2.ListItems(i).Selected = True
        txtQty1 = Val(LV2.ListItems(i).SubItems(9))
        Call AddToInvoice
    Next
    
End Sub

Private Sub AddToInvoice()

    Call MoveToFirstLV
    
    LV2.SelectedItem.SubItems(9) = LV2.SelectedItem.SubItems(9) - Val(txtQty1)
    LV2.SelectedItem.SubItems(10) = Val(txtQty1) * LV2.SelectedItem.SubItems(8)
    
    txtQty1.Visible = False
    
    Call CalculateTotals
    LV2.SetFocus
    
End Sub

Private Sub cmdAddOtherCharges_Click()
    
    On Error GoTo err
    
    If Val(txtOtherChargesSNo) > LVOtherCharges.ListItems.count + 1 Then
        MsgBox "Invalid S/No.", vbCritical
        txtOtherChargesSNo.Text = LVOtherCharges.ListItems.count + 1
        txtOtherChargesSNo.SetFocus
        Exit Sub
    End If
    
    If txtOtherCharges.Text = "" Then
        MsgBox "Please Enter Other Charges Description.", vbCritical
        txtOtherCharges.SetFocus
    End If
    
    If IsNumeric(txtOtherChargeAmt.Text) = False Then
        MsgBox "Invalid Amount.", vbCritical
        txtOtherChargeAmt.SelStart = 0
        txtOtherChargeAmt.SelLength = Len(txtOtherChargeAmt.Text)
        txtOtherChargeAmt.SetFocus
    End If
    
    Dim ITM As ListItem
    Set ITM = LVOtherCharges.ListItems.add(Val(txtOtherChargesSNo), , Val(txtOtherChargesSNo))
    ITM.SubItems(1) = txtOtherCharges.Text & ""
    ITM.SubItems(2) = Val(txtOtherChargeAmt)
    
    Call PrepareForNext
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub PrepareForNext()

    Dim i As Integer
    For i = 1 To LVOtherCharges.ListItems.count
        LVOtherCharges.ListItems(i).Text = i
    Next
    txtOtherChargesSNo.Text = LVOtherCharges.ListItems.count + 1
    txtOtherCharges.Text = ""
    txtOtherChargeAmt.Text = ""
    txtOtherCharges.SetFocus
    
End Sub

Private Sub cmdSavenStay_Click()

    If Saved Then
        Dim strInvNo As String
        strInvNo = txtPInvoice
        Call EditProforma(strInvNo, bCustomInvoice)
    End If
    
End Sub

Private Sub cmdSwap_Click()
    Dim Temp As String
    Temp = txtComConsignee
    txtComConsignee = txtComNotify
    txtComNotify = Temp
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub cmdAddItem_Click()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset

    Dim strMsg As String
    strMsg = ""
    If Not (cmbCust.MatchFound And cmbCountry.MatchFound And cmbProformas.MatchFound) Then Exit Sub
    If cmbCust.Text = "" Or cmbCountry.Text = "" Or cmbProformas.MatchFound = False Then Exit Sub
    
    Dim cmd As New ADODB.Command
    'Now Check If Same Item With The Same Size Already Exists Or Not.
    
    For i = 1 To LV2.ListItems.count
        If LV2.ListItems(i).SubItems(1) = cmbProformas.Text Then
            MsgBox "Proforma Already Added.", vbInformation
            Exit Sub
        End If
    Next i
    
    
    strMsg = ""
    
    With rs
        .Open "SELECT * FROM VFProformas WHERE PInvoice='" & cmbProformas.ID & "' And InvQty>0 ORDER BY CompItemCode", con, adOpenForwardOnly, adLockReadOnly
        
        If LV2.ListItems.count > 0 Then
            
            If cmbTrade <> ![TradeTerms] & "" Then strMsg = ",Trade Terms"
            If cmbPartial <> ![PartialShipment] & "" Then strMsg = strMsg & ",Partial Shipment"
            If cmbTrans <> ![TransShipment] & "" Then strMsg = strMsg & ",Trans Shipment"
            If cmbPayment <> ![PaymentTerms] & "" Then strMsg = strMsg & ",Payment Terms"
            If txtPackaging <> ![Packaging] & "" Then strMsg = strMsg & ",Packaging"
            If DeliveryDT <> ![DeliveryDT] & "" Then strMsg = strMsg & ",Delivery Date"
            
        End If
        
        If strMsg <> "" Then
            strMsg = "Following Fields Are Different From The Previous Order." & vbNewLine & Right$(strMsg, Len(strMsg) - 1) & "." & vbNewLine & "Do You Want Add This Order?"
            If MsgBox(strMsg, vbQuestion + vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
        
        cmbTradeTerms = ![TradeTerms] & ""
        cmbPartial = ![PartialShipment] & ""
        cmbTrans = ![TransShipment] & ""
        cmbPayment = ![PaymentTerms] & ""
        txtPackaging = ![Packaging] & ""
        'cmbShipAgent.ID = !ShippingAgent
        cmbSMethod = !SMethod & ""
        txtDeliveryDT = Format(![DeliveryDT], "dd-MM-yyyy")
        txtInvoiceof = !InvoiceOf & ""
        cmbShipAgent.ID = Val(!ShippingAgent & "")
'        For i = 0 To cmbBank.ListCount - 1
'            If Val(cmbBank.List(i, 1)) = Val(!OurBankID & "") Then
'                cmbBank.ListIndex = i
'                Exit For
'            End If
'        Next
        cmbBank.Text = !Bank & ""
        If OrderDT.Tag = "" Then
            'OrderDT = ![dt] & ""
            OrderDT.Tag = "One Added"
        Else
            If DateDiff("d", OrderDT, ![DT]) > 0 Then
                'OrderDT = ![dt] & ""
            End If
        End If
        
        Dim ITM As ListItem
        Dim iSNo As Integer
        
        LV2.ListItems.Clear
        
        Do Until .EOF
            
            iSNo = LV2.ListItems.count + 1
            
        
            'Set ITM = lv2.ListItems.Add(, ![OrderNo] & ![ItemCode] & "'", iSno)
            Set ITM = LV2.ListItems.add(, ![EntryID] & "'", iSNo)
            ITM.ListSubItems.add , , !ITemID & "" '![PInvoice] & ""
            ITM.ListSubItems.add(, , ![OrderNo] & "").Tag = 0
            ITM.ListSubItems.add , , ![ItemSize] & "" '![ItemID] & ""
            ITM.ListSubItems.add , , ""
            ITM.ListSubItems.add , , ""
            ITM.ListSubItems.add , , ![ItemName] & ""
            ITM.ListSubItems.add , , !CompItemCode & "" '![Unit] & ""
            ITM.ListSubItems.add , , Val(![Price] & "")
            ITM.ListSubItems(7).Tag = Val(![Price] & "") 'Val(![CustomPrice] & "")
            ITM.ListSubItems.add , , Val(![InvQty] & "")
            ITM.ListSubItems(8).Tag = Val(![InvQty] & "")
            ITM.ListSubItems.add , , Val(![Unit] & "")
            ITM.SubItems(10) = Val(![Price] & "") * Val(![InvQty] & "")
            ITM.ListSubItems(9).Tag = Val(![RebateRate] & "")
            ITM.ListSubItems(5).Tag = 0 'Val(![PackNo] & "")
            
            If strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
                ITM.ListSubItems.add(, , Format(!DeliveryDTItem, "dd-MMM-yyyy")).Tag = Format(!DeliveryDTItem, "dd/MM/yy")
            End If
            'Set itm = LV2.ListItems.Add
            'itm.SubItems(3) = "Batch # " & txtBatchNo & "   Lot # " & txtBatchNo & "/" & ![ItemCode] & ""
            
'            cmbItemSearch.AddItem !ITemID & " (" & !OrderNo & ")"
'            cmbItemSearch.List(iSNo - 1, 1) = !ITemID & ""
'            cmbItemSearch.List(iSNo - 1, 2) = !OrderNo & ""
            
            .MoveNEXT
            
        Loop
        .Close
        
        cmbItemSearch.ClearVals
        
        If strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
            cmbItemSearch.AddVals con, "ItemID + ' {' + OrderNo + '}' + ' (' + CompItemCode + ') {' + CONVERT(varchar(50),DeliveryDTItem,3) + '}'", "VFProformas", "ItemID+'|'+OrderNo+'|'+CompItemCode+'|'+CONVERT(varchar(50),DeliveryDTItem,3)", "WHERE PInvoice='" & cmbProformas.ID & "' And InvQty>0"
        Else
            cmbItemSearch.AddVals con, "ItemID + ' {' + OrderNo + '}' + ' (' + CompItemCode + ')'", "VFProformas", "ItemID+'|'+OrderNo+'|'+CompItemCode", "WHERE PInvoice='" & cmbProformas.ID & "' And InvQty>0"
        End If
        
        .Open "SELECT * FROM FPInvoiceOtherCharges WHERE PInvoice='" & cmbProformas.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            If MsgBox("This Proforma has Other Charges, Do you want to Add ?", vbQuestion + vbYesNo) = vbYes Then
                Do Until .EOF
                    iSNo = LVOtherCharges.ListItems.count + 1
                    Set ITM = LVOtherCharges.ListItems.add(, , iSNo)
                    ITM.SubItems(1) = !ChargesDetail & ""
                    ITM.SubItems(2) = Val(!ChargesAmt & "")
                    Call PrepareForNext
                    .MoveNEXT
                Loop
            End If
        End If
        .Close
    End With
    
    'Dim iUp As Integer
    'iUp = UBound(Orders)
    'Orders(iUp) = cmbOrders
    'ReDim Preserve Orders(iUp + 1)
    
    Orders.add cmbProformas.ID, cmbProformas.ID
     
     
    Set rs = Nothing
    
    Call CalculateTotals
    cmbItemSearch.SetFocus
    
    Exit Sub
    
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()

    'LV.ListItems.Remove LV.SelectedItem.Index
    'Dim strOrder As String
    'strOrder = LV.SelectedItem.SubItems(1)
    If LV.ListItems.count = 0 Then Exit Sub
    If Val(LV.SelectedItem.ListSubItems(1).Tag) > 0 Then
        strDelList = strDelList & Val(LV.SelectedItem.ListSubItems(1).Tag) & ","
    End If
    
    
    
    LV.ListItems.Remove LV.SelectedItem.Index
    'For i = 1 To LV.ListItems.Count
    '    If i > LV.ListItems.Count Then Exit For
    '    If LV.ListItems(i).SubItems(1) = strOrder Then
    '        LV.ListItems.Remove i
    '    End If
    'Next i
    
    'Orders.Remove strOrder
    cmdDelete.Visible = False
    
    Call CalculateTotals
    
End Sub

Private Sub cmdSClear_Click()
    If Saved Then
        Unload Me
        Load frmNewCustomInvoice
        frmNewCustomInvoice.AddNewInvoice True, bCourierInvoice
    End If
End Sub

Private Sub cmdSClose_Click()
    If Saved Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
   
    On Error GoTo err
    
    If strCompany = "Tecno" Then
        txtPInvoice = GetNewTecnoInvoiceNo(OrderDT)
    End If
    
    bCourierInvoice = False
    strDelList = ""
    strProforma = ""
    'SSTab1.TabVisible(4) = False
    SSTab1.Tab = 0
    
    Call SaveLV(LV)
    
    Call cmbCust.AddVals(con, "Distinct CustCode", "ForeignCustomers", "CustCode", " WHERE Active=1")
    OrderDT = Date
    Call OrderDT_Change
    
    DTETA = Date
    DTETA = Null
    AWBDT = Date
    AWBDT = Null
    FormEDT = Date
    FormEDT = Null
    
    DTLC = Date
    DTLC = Null
    
    DTGatePass = Date
    DTGatePass = Null
    
    Me.KeyPreview = True
    cmbCust.ListHeight = 2500

    cmbShipAgent.ListHeight = 2000
    cmbInsurAgent.ListHeight = 2000
    
    Call FillCmbs
        
        
    txtDTREDesc = "Hereby it is certified that 49332 Kgs Stainless Steel received vide DTRE approval number " & _
     "RTOS/2589/29102008 & Bill of Entry Number 678 Dated 22.10.2008 under SRO No. 563(I)/2005 dated 06.06.2008 " & _
     "& against postdated cheque no 3179423 dated 28.10.2008." & vbNewLine & vbNewLine & _
     "Quantity Imported                                                     =   49332   Kgs" & vbNewLine & _
     "Quantity Already Re-Exported                      =   268056  Pcs     =   9784    Kgs" & vbNewLine & _
     "Quantity Re-Exported In this Invoice              =   5633    Pcs     =   206     Kgs" & vbNewLine & _
     "Total Quantity Re-Exported Including this Invoice =   273689  Pcs     =   9990    Kgs" & vbNewLine & _
     "Quantity Balance                                  =   1079811 Pcs     =   39413   Kgs"
     
         
    cmbProformas.ListHeight = 4000
    cmbItemSearch.ListHeight = 1000
    
    If strCompany = "IAA" Or strCompany = "Kami" Then
        chkCIFTerms.Visible = True
    Else
        chkCIFTerms.Visible = False
    End If
    
    Exit Sub
err:
    MsgBox err.Description
   
End Sub

Public Sub AddNewInvoice(Optional p_bCustomInvoice As Boolean = True, Optional bCourierInvoice As Boolean = False)
    bCustomInvoice = p_bCustomInvoice
    Me.Show
End Sub
Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)

    For i = 1 To Orders.count
        Orders.Remove 1
    Next i
    Set Orders = Nothing
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    FAll.Left = (Me.ScaleWidth - FAll.Width) / 2
    FAll.Top = (Me.ScaleHeight - FAll.Height) / 2
End Sub

Private Sub lblBallComDecs_Click()
    If txtComDeclarationBalls = "" Then
        txtComDeclarationBalls = con.Execute("Select Declaration From ItemCatagories Where CatID='01'").Fields(0).Value & ""
    End If
End Sub

Private Sub lblBallDecs_Click()
    If txtDeclarationBalls = "" Then
        txtDeclarationBalls = con.Execute("Select Declaration From ItemCatagories Where CatID='01'").Fields(0).Value & ""
    End If
End Sub

Private Sub lblGarmentsComDecs_Click()
    If txtComDeclarationGarments = "" Then
       txtComDeclarationGarments = con.Execute("Select Declaration From ItemCatagories Where CatID='02'").Fields(0).Value & ""
    End If
End Sub

Private Sub lblGarmentsDecs_Click()
    If txtDeclarationGarments = "" Then
       txtDeclarationGarments = con.Execute("Select Declaration From ItemCatagories Where CatID='02'").Fields(0).Value & ""
    End If
End Sub

Private Sub lblGlovesComDecs_Click()
    If txtComDeclarationGloves = "" Then
       txtComDeclarationGloves = con.Execute("Select Declaration From ItemCatagories Where CatID='03'").Fields(0).Value & ""
    End If
End Sub

Private Sub lblGlovesDecs_Click()
    If txtDeclarationGloves = "" Then
       txtDeclarationGloves = con.Execute("Select Declaration From ItemCatagories Where CatID='03'").Fields(0).Value & ""
    End If
End Sub

Private Sub lblSwapAddresses_Click()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    
    With rs
        If lblSwapAddresses.Tag = "" Then
            .Open "Select Address,FakeAddress From ForeignCustomers Where CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
            lblSwapAddresses.Tag = "0"
        Else
            .Open "Select OtherAddress,OtherFakeAddress From ForeignCustomers Where CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
            lblSwapAddresses.Tag = ""
        End If
        txtComConsignee.Text = .Fields(0) & ""
        txtComNotify.Text = .Fields(1) & ""
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub LV_Click()
    If LV.ListItems.count > 0 Then
        Call LV_ItemClick(LV.SelectedItem)
    End If
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    'If LV.SelectedItem.Index Mod 2 = 0 Then Exit Sub
    If iColNo = 9 Then
        With txtPrice
            .Move LV.ColumnHeaders(9).Left + LV.Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(9).Width
            .Text = Val(LV.SelectedItem.SubItems(8))
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    ElseIf iColNo = 12 Or iColNo = 13 Then
        With txtDTRENo
            .Move LV.ColumnHeaders(iColNo).Left + LV.Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(iColNo).Width
            .Text = LV.SelectedItem.SubItems(iColNo - 1)
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    Else
        With txtQty
            .Move LV.ColumnHeaders(10).Left + LV.Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(10).Width
            .Text = Val(LV.SelectedItem.SubItems(9))
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    End If
    
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    
    With cmdDelete
        .Move .Left, LV.Top + Item.Top
        .Visible = True
    End With
    txtQty.Visible = False
    
End Sub

Private Sub CalculateTotals()


    Dim dTotal As Double
    Dim DPcsTotal As Long
    With LV
        For i = 1 To .ListItems.count
            'dQty = dQty + Val(.TextMatrix(i, 4))
            dTotal = dTotal + Val(.ListItems(i).SubItems(10) & "")
            DPcsTotal = DPcsTotal + Val(.ListItems(i).SubItems(9) & "")
        Next i
    End With
    'txttQty = dQty
    lblTotal = dTotal
    lblTotWords = ConvertInWords(CCur(dTotal), txtCurrency.Text)
    lblTotPcs = DPcsTotal
    
End Sub

Private Function Saved() As Boolean
    
    On Error GoTo err
    
    Dim iLoadID As Integer
    Dim iDisID As Integer
    Dim iBankID1 As Integer
    Saved = False
    If txtPInvoice = "" Then
        MsgBox "Please Enter Proforma Invoice No.", vbInformation
        txtPInvoice.SetFocus
        Exit Function
'    ElseIf Trim(txtFormE) = "" And bCourierInvoice = False Then
'        MsgBox "Please Enter Form E #"
'        Exit Function
    End If
    'Check FormE # Duplication
'    If bCourierInvoice = False Then
'        If FormENoExists(strProforma, txtFormE.Text) Then
'            MsgBox "Form E No. Exists.", vbInformation
'            Exit Function
'        End If
'    End If
    If LV.ListItems.count = 0 Then
        MsgBox "No Orders Added To This Proforma.", vbInformation
        Exit Function
    End If
    
    
'    If cmbBank.MatchFound = False Then
'        MsgBox "Invalid Bank Selected.", vbInformation
'        SSTab1.Tab = 1
'        cmbBank.SetFocus
'        Exit Function
'    End If
    
    If LCase(cmbPayment) = "letter of credit" Then
        If txtLC = "" Then
            MsgBox "You Must Enter LC # If Payment Terms Is 'Letter Of Credit.", vbInformation
            Exit Function
        End If
    End If
    
    If cmbBankAccNo.ListIndex = -1 Then
        'iBankID = 4
        iBankID1 = -1
    Else
        iBankID1 = cmbBankAccNo.List(cmbBankAccNo.ListIndex, 1)
    End If
    
    If cmbLoadPort.ListIndex = -1 Then
        'iLoadID = 4
        iLoadID = -1
    Else
        iLoadID = cmbLoadPort.List(cmbLoadPort.ListIndex, 1)
    End If
    
    If cmbDisPort.ListIndex = -1 Then
        'idisid=4
        iDisID = -1
    Else
        iDisID = cmbDisPort.List(cmbDisPort.ListIndex, 1)
    End If
    
    Dim strSRONo  As String
    If strProforma = "" Then
        'Get Latest SRONo...
        strSRONo = con.Execute("SELECT SRONo FROM SROs WHERE EntryID=(SELECT MAX(EntryID) FROM SROs)").Fields(0).Value & ""
    End If
    
    Call StartTrans(con)
    '----------------------------------------------------------------
    If strProforma <> "" Then
    
        con.Execute "UPDATE FProformaOrders SET InvQty=InvQty+IsNull((Select  " & _
         "Qty From CustomInvoiceItems  Where CustomInvoice='" & strProforma & _
         "' And RefID=FProformaOrders.EntryID),0) WHERE EntryID In(Select RefID" & _
         " From CustomInvoiceItems Where CustomInvoice='" & strProforma & "')"
         
        'con.Execute "Delete from CustomInvoice Where CustomInvoice='" & _
         strProforma & "'"
         
    End If
        
    If strProforma = "" Then
        Dim ComDecBalls As String, ComDecGarments As String, ComDecGloves As String
        Dim ComNotify As String, ComConsignee As String
        Dim ComPayTerms As String
        
'        If cmbConsignee = "Bank Address" Then
'            ComNotify = txtBankAddress
'        Else
'            ComNotify = con.Execute("Select Address1 From FCustBanks Where CustCode='" & cmbCust.ID & "' And Country='" & cmbCountry.Text & "' AND DefaultBank=1").Fields(0).value & ""
'        End If
        Dim rs As New ADODB.Recordset
        With rs
            .Open "Select Address1 From FCustBanks Where CustCode='" & cmbCust.ID & "' And Country='" & cmbCountry.Text & "' AND DefaultBank=1", con, adOpenForwardOnly, adLockReadOnly
            If .RecordCount > 0 Then
                ComNotify = .Fields(0) & ""
            End If
            .Close
            .Open "Select Address From ForeignCustomers Where CustCode='" & cmbCust.ID & "' And Country='" & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
            If .RecordCount > 0 Then
                ComConsignee = .Fields(0) & ""
            End If
        End With
        
        ComDecBalls = txtDeclarationBalls
        ComDecGarments = txtDeclarationGarments
        ComDecGloves = txtDeclarationGloves
        ComPayTerms = cmbPayment
    Else
        ComNotify = txtComNotify
        ComConsignee = txtComConsignee
        
        ComDecBalls = txtComDeclarationBalls
        ComDecGarments = txtComDeclarationGarments
        ComDecGloves = txtComDeclarationGloves
        ComPayTerms = txtComPayTerms
        
    End If
    
    Dim strDTGatePass As String
    If IsNull(DTGatePass) Then
        strDTGatePass = "NULL"
    Else
        strDTGatePass = "'" & DTGatePass & "'"
    End If
    If strProforma = "" Then

        con.Execute "INSERT INTO CustomInvoice (CustomInvoice,DT,CustCode,Country,ShippingMode," & _
         "LoadPortID,DisPortID,TradeTerms,PaymentTerms,LCNo,PartialShipment," & _
         "TransShipment,ShippingAgentID,InsuranceAgentID,OurBankID," & _
         "AWBNo,AWBNoDT,FormE,FormEDT,Vessel,GrossWeight,DeclarationsBalls,DeclarationsGarments," & _
         "DeclarationsGloves,Consignee,ComDeclarationsBalls,ComDeclarationsGarments," & _
         "ComDeclarationsGloves,ComNotify,ComConsignee,ComPayTerms,ComAgainst,ComDrawnUnder," & _
         "ComSpecial,ManDischarge,ManDeclaration,Warehouse,Gate,ContainerNo,DTETA,SRONo,CustCurrency,CompanyRefID,SealNo,InvoiceOf,LCDate,GatePassNo,GatePassDT,DTREDescription,NoSRO,CIFTradeTerms) VALUES('" & txtPInvoice & _
         "','" & OrderDT.Value & "','" & cmbCust.Text & "','" & cmbCountry.Text & "','" & _
         cmbSMethod.Text & "'," & iLoadID & "," & iDisID & ",'" & cmbTradeTerms.Text & _
         "','" & cmbPayment & "','" & txtLC & "','" & cmbPartial & "','" & _
         cmbTrans & "'," & Val(cmbShipAgent.ID) & "," & Val(cmbInsurAgent.ID) & "," & _
         iBankID1 & ",'" & txtAWBNo & "','" & AWBDT & "','" & txtFormE & "','" & FormEDT & _
         "','" & txtVessel & "'," & Val(txtGWeight) & ",'" & _
         txtDeclarationBalls & "','" & txtDeclarationGarments & "','" & txtDeclarationGloves & _
         "','" & txtBankAddress & "','" & ComDecBalls & "','" & _
         ComDecGarments & "','" & ComDecGloves & "','" & ComNotify & _
         "','" & ComConsignee & "','" & ComPayTerms & "','" & txtComAgainst & "','" & _
         txtComDrawn & "','" & txtComSpecial & "','" & txtManDischarge & "','" & txtManDeclaration & "','" & _
         txtWarehouse & "','" & txtGate & "','" & txtContainerNo & "','" & DTETA & "','" & strSRONo & "','" & _
         txtCurrency & "'," & cmbCompany.ID & ",'" & txtSealNo & "','" & txtInvoiceof & "','" & DTLC & "','" & _
         txtGatePassNo & "'," & strDTGatePass & ",'" & txtDTREDesc & "'," & chkNoSRO.Value & "," & chkCIFTerms.Value & ")"
    Else
        con.Execute "UPDATE CustomInvoice SET DT='" & OrderDT & "',ShippingMode='" & cmbSMethod & _
         "',LoadPortID=" & iLoadID & ",DisPortID=" & iDisID & ",TradeTerms='" & cmbTradeTerms.Text & _
         "',PaymentTerms='" & cmbPayment & "',LCNo='" & txtLC & "',PartialShipment='" & _
         cmbPartial & "',TransShipment='" & cmbTrans & "',ShippingAgentID=" & Val(cmbShipAgent.ID) & _
         ",InsuranceAgentID=" & Val(cmbInsurAgent.ID) & ",OurBankID=" & iBankID1 & ",AWBNo='" & _
         txtAWBNo & "',AWBNoDT='" & AWBDT & "',FormE='" & txtFormE & "',FormEDT='" & FormEDT & _
         "',Vessel='" & txtVessel & "',GrossWeight=" & Val(txtGWeight) & ",DeclarationsBalls='" & _
         txtDeclarationBalls & "',DeclarationsGarments='" & txtDeclarationGarments & _
         "',DeclarationsGloves='" & txtDeclarationGloves & "',Consignee='" & txtBankAddress & _
         "',ComDeclarationsBalls='" & ComDecBalls & "',ComDeclarationsGarments='" & _
         ComDecGarments & "',ComDeclarationsGloves='" & ComDecGloves & "',ComNotify='" & _
         ComNotify & "',ComConsignee='" & ComConsignee & "',ComPayTerms='" & ComPayTerms & _
         "',ComAgainst='" & txtComAgainst & "',ComDrawnUnder='" & txtComDrawn & "',ComSpecial='" & _
         txtComSpecial & "',ManDischarge='" & txtManDischarge & "',ManDeclaration='" & _
         txtManDeclaration & "',Warehouse='" & txtWarehouse & "',Gate='" & txtGate & "',ContainerNo='" & _
         txtContainerNo & "',DTETA='" & DTETA & "',SealNo='" & txtSealNo & "',Invoiceof='" & txtInvoiceof & _
         "',LCDate='" & DTLC & "',GatePassNo='" & txtGatePassNo & "',GatePassDT=" & strDTGatePass & ",DTREDescription='" & txtDTREDesc & "',NoSRO=" & chkNoSRO.Value & ",CIFTradeTerms=" & chkCIFTerms.Value & " WHERE CustomInvoice='" & strProforma & "'"
         
        If strDelList <> "" Then
            strDelList = Left(strDelList, Len(strDelList) - 1)
             con.Execute "DELETE FROM CustomInvoiceItems WHERE EntryID IN(" & strDelList & ")"
        End If
    End If
    
    Dim dCustomPrice As Double, dPrice As Double
    For i = 1 To LV.ListItems.count
        
        With LV.ListItems(i)
            dPrice = Val(.SubItems(8))
            dCustomPrice = Val(.ListSubItems(7).Tag)
            If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                dCustomPrice = dPrice
            End If
            If Val(.ListSubItems(1).Tag) = 0 Then
                con.Execute "INSERT INTO CustomInvoiceItems(CustomInvoice,RefID,Qty,Price," & _
                "CustomPrice,RebateRate,DTRENo,ItemForLabTest,BatchNo) VALUES('" & txtPInvoice & "'," & Val(.key) & "," & _
                 Val(.SubItems(9)) & "," & dPrice & "," & _
                 dCustomPrice & "," & Val(.ListSubItems(9).Tag) & ",'" & .SubItems(11) & "'," & Abs(CBool(.ListSubItems(11).Tag)) & ",'" & .SubItems(12) & "')"
            Else
                If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                    con.Execute "UPDATE CustomInvoiceItems SET Qty=" & Val(.SubItems(9)) & _
                         ",CustomPrice=" & dPrice & ",Price=" & dPrice & ",DTRENo='" & .SubItems(11) & "',ItemForLabTest=" & Abs(CBool(.ListSubItems(11).Tag)) & ",BatchNo='" & .SubItems(12) & "' WHERE EntryID=" & Val(.ListSubItems(1).Tag)
                Else
                    If bCustomInvoice Then
                        con.Execute "UPDATE CustomInvoiceItems SET Qty=" & Val(.SubItems(9)) & _
                         ",CustomPrice=" & dPrice & ",DTRENo='" & .SubItems(11) & "',ItemForLabTest=" & Abs(CBool(.ListSubItems(11).Tag)) & ",BatchNo='" & .SubItems(12) & "' WHERE EntryID=" & Val(.ListSubItems(1).Tag)
                    Else
                        con.Execute "UPDATE CustomInvoiceItems SET Qty=" & Val(.SubItems(9)) & _
                         ",Price=" & dPrice & ",DTRENo='" & .SubItems(11) & "',ItemForLabTest=" & Abs(CBool(.ListSubItems(11).Tag)) & ",BatchNo='" & .SubItems(12) & "' WHERE EntryID=" & Val(.ListSubItems(1).Tag)
                    End If
                End If
            End If
            
            con.Execute "UPDATE FProformaOrders SET InvQty=InvQty-" & _
             Val(.SubItems(9)) & " WHERE EntryID=" & Val(.key)
             
         End With
         
    Next i
   
    ''Add Custom Invoice Rates...
    If strCompany = "PAKSMITH GROUP" Then
        If strProforma = "" Then
            dCustomPrice = Val(con.Execute("SELECT SUM(Qty*Price)/SUM(Qty) FROM CustomInvoiceItems WHERE CustomInvoice='" & txtPInvoice & "'").Fields(0).Value & "")
            con.Execute "UPDATE CustomInvoiceItems SET CustomPrice=" & dCustomPrice & " WHERE CustomInvoice='" & txtPInvoice.Text & "'"
        End If
    End If
    ''
    If strProforma <> "" Then
        'con.Execute "Delete From CustomInvoiceOtherCharges Where CustomInvoice='" & strProforma & "' AND Custom=" & Abs(bCustomInvoice)
        con.Execute "Delete From CustomInvoiceOtherCharges Where CustomInvoice='" & strProforma & "'"
    End If
    For i = 1 To LVOtherCharges.ListItems.count
        With LVOtherCharges.ListItems(i)
            con.Execute "Insert Into CustomInvoiceOtherCharges(CustomInvoice,ChargesDetail,ChargesAmt,SNo,Custom)" & _
             " Values('" & IIf(strProforma = "", txtPInvoice.Text, strProforma) & "','" & _
             .SubItems(1) & "'," & Val(.SubItems(2)) & "," & Val(.Index) & "," & Abs(bCustomInvoice) & ")"
             
             If strProforma = "" Then   'iF nEW Invoice THen save them as commercial....
                con.Execute "INSERT INTO CustomInvoiceOtherCharges(CustomInvoice,ChargesDetail,ChargesAmt,SNo,Custom)" & _
                 " Values('" & IIf(strProforma = "", txtPInvoice.Text, strProforma) & "','" & _
                 .SubItems(1) & "'," & Val(.SubItems(2)) & "," & Val(.Index) & ",0)"
             
             End If
        End With
    Next
    
    con.CommitTrans
        
    If chkPrint Then
        Call PrintCustomInvoice(txtPInvoice.Text, True)
    End If
    
    Saved = True
    
    Exit Function

err:
    MsgBox err.Description
    
End Function

Private Sub PrintCustomInvoice(p_strCustomInvoice As String, Optional bCustomInvoice As Boolean = True)

    On Error GoTo err

    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    If bCustomInvoice Then
        Set rpt = rptApp.OpenReport(RptPath & "\rptCustomInvoice.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\rptCommercialInvoice.rpt")
    End If
    'rpt.FormulaFields(15).Text = bImpMaterialDetail
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VrptCustomInvoiceDetail.CustomInvoice}='" & p_strCustomInvoice & "' ", rpt
    
    Screen.MousePointer = vbDefault


    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub FillCmbs()

    Call AddToCombo(cmbTradeTerms, "Description", "TradeTerms")
    
    With cmbPartial
        .AddItem "Allowed"
        .AddItem "Not Allowed"
    End With
    
    With cmbTrans
        .AddItem "Allowed"
        .AddItem "Not Allowed"
    End With
    
    With cmbConsignee
        .AddItem "Bank Address"
        .AddItem "Customer Address"
    End With
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Description From PaymentTerms", con, adOpenForwardOnly, adLockReadOnly
        cmbPayment.Clear
        Do Until .EOF
            cmbPayment.AddItem .Fields(0) & ""
            .MoveNEXT
        Loop
        .Close
        
    End With
    
    Set rs = Nothing
        
    With cmbSMethod
        .AddItem "By Air"
        .AddItem "By Sea"
        .AddItem "By Courier"
    End With
    
    cmbLoadCountry.AddItem "Pakistan"
    cmbLoadCountry.ListIndex = 0
    
    cmbShipAgent.ClearVals
    Call cmbShipAgent.AddVals(con, "Name", "ShippingAgents", "EntryID", " Where Active=1")
    
    cmbInsurAgent.ClearVals
    Call cmbInsurAgent.AddVals(con, "Name", "InsuranceAgents", "EntryID", "")
        
    cmbCompany.ListHeight = 6000
    cmbCompany.AddVals con, "CompanyName", "Companies", "EntryID", " Order By CompanyName"
    
    Dim iIndex As Integer
    iIndex = GetSingleLongValue("MIN(EntryID)", "Companies")
    cmbCompany.ID = iIndex
    
End Sub
Private Sub cmbDisCity_Change()
    Call cmbDisCity_Click
End Sub

Private Sub cmbDisCity_Click()
    If cmbDisCity.MatchFound = False Then Exit Sub
    Call AddPorts(cmbDisCity.List(cmbDisCity.ListIndex, 1), cmbDisPort)
End Sub

Private Sub cmbDisCountry_Change()
    Call cmbDisCountry_Click
End Sub

Private Sub cmbDisCountry_Click()
    Call AddCities(cmbDisCountry.Text, cmbDisCity)
    'txtShipment.Text = "From " & cmbLoadCountry.Text & " To " & cmbDisCountry.Text
End Sub

Private Sub cmbLoadCity_Change()
    Call cmbLoadCity_Click
End Sub

Private Sub cmbLoadCity_Click()
    If cmbLoadCity.MatchFound = False Then Exit Sub
    Call AddPorts(cmbLoadCity.List(cmbLoadCity.ListIndex, 1), cmbLoadPort)
End Sub

Private Sub cmbLoadCountry_Change()
    Call cmbLoadCountry_Click
End Sub

Private Sub cmbLoadCountry_Click()
    Call AddCities(cmbLoadCountry.Text, cmbLoadCity)
End Sub

Private Sub cmbThrough_Change()
    Call cmbThrough_Click
End Sub

Private Sub cmbThrough_Click()
    If cmbThrough.MatchFound = False Then Exit Sub
    txtBankAddress = cmbThrough.List(cmbThrough.ListIndex, 2)
End Sub
 
Private Sub GetOrders()

    On Error GoTo err
    
    Dim strTTerms As String
    
    With cmbCountry
        If Not .MatchFound Then Exit Sub
    End With
    
    Dim rs As New ADODB.Recordset
    
'    With rs
'        .Open "Select Distinct VFProformaList.PInvoice,VFProformaList.Orders From VFProformas INNER JOIN VFProformaList ON VFProformas.PInvoice=VFProformaList.PInvoice Where VFProformas.CustCode+VFProformas.Country='" & cmbCust.ID & cmbCountry & _
'         "' AND VFProformaList.CompanyRefID=" & Val(cmbCompany.ID) & " AND InvQty>0", con, adOpenForwardOnly, adLockReadOnly
'
'        cmbProforma.Clear
'        Do Until .EOF
'            cmbProforma.AddItem .Fields(0) & " (" & !Orders & ")"
'            cmbProforma.List(.AbsolutePosition - 1, 1) = .Fields(0)
'            .MoveNEXT
'        Loop
'        .Close
'    End With
    cmbProformas.ClearVals
    cmbProformas.AddVals con, "DISTINCT VFProformaList.PInvoice + ' {' + VFProformaList.Orders + '}'", "VFProformas INNER JOIN VFProformaList ON VFProformas.PInvoice=VFProformaList.PInvoice", "VFProformaList.PInvoice", "WHERE VFProformas.CustCode+VFProformas.Country='" & cmbCust.ID & cmbCountry & "' AND VFProformaList.CompanyRefID=" & Val(cmbCompany.ID) & " AND InvQty>0"
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub



Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        If x > LV.ColumnHeaders(i).Left And x < LV.ColumnHeaders(i).Left + LV.ColumnHeaders(i).Width Then
            iColNo = i
            Exit For
        End If
    Next
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button <> 2 Then Exit Sub
    If LV.ListItems.count = 0 Then Exit Sub
    Dim strRet As String
    strRet = ShowPopUpMenu1("Mark As Lab Test Item")
    'Remove Tag from All Others...
    
    If strRet = "Mark As Lab Test Item" Then
        Dim i As Integer
        For i = 1 To LV.ListItems.count
            LV.ListItems(i).ListSubItems(11).Tag = "FALSE"
        Next
        LV.SelectedItem.ListSubItems(11).Tag = "TRUE"
    End If
    
End Sub

Private Sub LVOtherCharges_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If MsgBox("Do you want to Delete it ?", vbQuestion + vbYesNo) = vbYes Then
            LVOtherCharges.ListItems.Remove LVOtherCharges.SelectedItem.Index
            Call PrepareForNext
        End If
    End If
End Sub

Private Sub OrderDT_Change()

    If strCompany = "QEL" Then
        txtPInvoice = GetNewQELInvoiceNo(OrderDT)
    ElseIf strCompany = "Banzai" Then
        If cmbCompany.MatchFound Then
            txtPInvoice = GetInvoiceNo(Val(cmbCompany.ID), False, OrderDT)
        End If
    ElseIf strCompany = "Instrumed" Then
        If cmbCompany.MatchFound Then
            txtPInvoice = GetInvoiceNo(Val(cmbCompany.ID), False, OrderDT)
        End If
    End If
    
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

    If SSTab1.Tab = 2 Then
        cmbItemSearch.Visible = True
    Else
        cmbItemSearch.Visible = False
    End If
    
    If SSTab1.Tab = 0 Then
        cmbItemSearchInvoice.Visible = True
    Else
        cmbItemSearchInvoice.Visible = False
    End If
    
End Sub

Private Sub txtDTRENo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LV.SelectedItem.SubItems(iColNo - 1) = txtDTRENo.Text
        txtDTRENo.Visible = False
    ElseIf KeyAscii = 27 Then
        txtDTRENo.Visible = False
    End If
End Sub

Private Sub txtDTRENo_LostFocus()
    txtDTRENo.Visible = False
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
        
    On Error GoTo err
    If KeyAscii = 13 Then
        If Val(txtPrice) <= 0 Then
            MsgBox "Invalid Price.", vbInformation
            Exit Sub
        End If
        LV.SelectedItem.SubItems(8) = Val(txtPrice)
        LV.SelectedItem.SubItems(10) = Val(txtPrice) * LV.SelectedItem.SubItems(9)
        
'        On Error Resume Next
'        LV2.ListItems(LV.SelectedItem.Key).SubItems(8) = Val(LV.SelectedItem.ListSubItems(8).Tag) - Val(txtQty)
'        If err.Number = 35601 Then
'            Resume Next
'        End If
        
        txtPrice.Visible = False
        
        Call CalculateTotals
    ElseIf KeyAscii = 27 Then
        txtPrice.Visible = False
    End If
    
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    If KeyAscii = 13 Then
    
        If Val(txtQty) > Val(LV.SelectedItem.ListSubItems(8).Tag) And bAcceptsExtraQty = False Then
            MsgBox "Invalid Qty." & vbNewLine & "You can only add " & Val(LV.SelectedItem.ListSubItems(8).Tag) & " Qty.", vbInformation
            Exit Sub
        End If
        
        LV.SelectedItem.SubItems(9) = Val(txtQty)
        LV.SelectedItem.SubItems(10) = Val(txtQty) * LV.SelectedItem.SubItems(8)
        
        On Error Resume Next
        LV2.ListItems(LV.SelectedItem.key).SubItems(8) = Val(LV.SelectedItem.ListSubItems(8).Tag) - Val(txtQty)
        If err.Number = 35601 Then
            Resume Next
        End If
        
        txtQty.Visible = False
        
        Call CalculateTotals
    ElseIf KeyAscii = 27 Then
        txtQty.Visible = False
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lv2_Click()
    If LV2.ListItems.count > 0 Then
        Call lv2_ItemClick(LV2.SelectedItem)
    End If
End Sub

Private Sub LV2_DblClick()

    If LV2.ListItems.count = 0 Then Exit Sub
    
    With txtQty1
        .Move LV2.ColumnHeaders(10).Left + LV2.Left, LV2.Top + LV2.SelectedItem.Top, LV2.ColumnHeaders(10).Width
        .Text = Val(LV2.SelectedItem.SubItems(9))
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
End Sub

Private Sub lv2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    If Item.Index Mod 2 = 0 Then
        cmdDelete.Visible = False
        Exit Sub
    End If
    With cmdDelete
        .Move .Left, LV2.Top + Item.Top
        .Visible = True
    End With
    txtQty1.Visible = False
    
End Sub
Private Sub txtqty1_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    Dim ITM As ListItem
    
    If KeyAscii = 13 Then
        If Val(txtQty1) > Val(LV2.SelectedItem.ListSubItems(9)) And bAcceptsExtraQty = False Then
            MsgBox "Invalid Qty." & vbNewLine & "You can only add " & Val(LV2.SelectedItem.ListSubItems(9)) & " Qty.", vbInformation
            Exit Sub
        End If
        If Val(txtQty1) <= 0 Then
            MsgBox "Invalid Qty.", vbInformation
            Exit Sub
        End If
        
        Call MoveToFirstLV
        
        LV2.SelectedItem.SubItems(9) = LV2.SelectedItem.SubItems(9) - Val(txtQty1)
        LV2.SelectedItem.SubItems(10) = Val(txtQty1) * Val(LV2.SelectedItem.SubItems(8) & "")
        
        txtQty1.Visible = False
        If Val(LV2.SelectedItem.ListSubItems(2).Tag) > 1 Then
            Call AddPacksOtherArticles(LV2.SelectedItem.ListSubItems(5).Tag, LV2.SelectedItem.Index, LV2.SelectedItem.SubItems(2))
        End If
'        Call CalculateTotals
        
        'LV2.SetFocus
        cmbItemSearch.SetFocus
        cmbItemSearch.ID = ""
    ElseIf KeyAscii = 27 Then
        txtQty1.Visible = False
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub AddPacksOtherArticles(iPackNo As Integer, iIndex As Integer, strOrderNo As String)
    
    For i = 1 To LV2.ListItems.count
        
        If LV2.ListItems(i).ListSubItems(5).Tag = iPackNo And strOrderNo = LV2.ListItems(i).SubItems(2) And i <> iIndex Then
            LV2.ListItems(i).Selected = True
            Call MoveToFirstLV
                
            LV2.SelectedItem.SubItems(9) = LV2.SelectedItem.SubItems(9) - Val(txtQty1)
            LV2.SelectedItem.SubItems(10) = Val(txtQty1) * LV2.SelectedItem.SubItems(8)
            txtQty1.Visible = False
        End If
    Next
    Call CalculateTotals
    
End Sub

Private Sub MoveToFirstLV()
    
    On Error GoTo err
    Dim iSNo As Integer
    Dim ITM As ListItem
     
    iSNo = LV.ListItems.count + 1
    
    Set ITM = LV.ListItems.add(, LV2.SelectedItem.key, iSNo)
    'Set ITM = LV2.SelectedItem
ComeHereForEdit:
    For i = 1 To LV2.SelectedItem.ListSubItems.count
        ITM.ListSubItems.add , , LV2.SelectedItem.ListSubItems(i).Text
        ITM.ListSubItems(i).Tag = LV2.SelectedItem.ListSubItems(i).Tag
    Next i
    ITM.ListSubItems(9).Text = txtQty1
    'ITM.ListSubItems(6).Tag = txtQty1
    ITM.SubItems(10) = Val(LV2.SelectedItem.SubItems(8) & "") * Val(txtQty1)
    ITM.SubItems(11) = ""
    ITM.ListSubItems(11).Tag = "False"
    
'    ITM.ListSubItems.Add , , ![OrderNo] & ""
'    ITM.ListSubItems.Add , , ![ItemCode] & ""
'    ITM.ListSubItems.Add , , ![Description] & ""
'    ITM.ListSubItems.Add , , ![Unit] & ""
'    ITM.ListSubItems.Add , , Val(![Price] & "")
'    ITM.ListSubItems.Add , , Val(![InvQty] & "")
'    ITM.ListSubItems(6).Tag = Val(![InvQty] & "")
'    ITM.SubItems(7) = Val(![Price] & "") * Val(![InvQty] & "")

    'Dim iIndex As Integer
    'iIndex = LV2.SelectedItem.Index
    'Set itm = LV.ListItems.Add
    'itm.SubItems(3) = LV2.ListItems(iIndex).SubItems(3)
    
    Exit Sub
err:
    If err.Description = "Key is not unique in collection" And err.Number = 35602 Then
        Set ITM = LV.ListItems(LV2.SelectedItem.key)
        ITM.SubItems(9) = Val(ITM.SubItems(9)) + Val(txtQty1)
        ITM.SubItems(10) = Val(ITM.SubItems(9)) * Val(ITM.SubItems(8))
        
    Else
        MsgBox err.Description
    End If
End Sub

Public Function FormENoExists(p_strInvNo As String, p_strFormENo As String) As Boolean
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim bReturn As Boolean
    If p_strInvNo = "" Then
        strSQL = "Select Count(FormE) From CustomInvoice Where FormE='" & p_strFormENo & "'"
    Else
        strSQL = "Select Count(FormE) From CustomInvoice Where FormE='" & p_strFormENo & "' AND CustomInvoice<>'" & p_strInvNo & "'"
    End If
    
    With rs
        .Open strSQL, con, adOpenForwardOnly, adLockReadOnly
        If .Fields(0) = 0 Then
            bReturn = False
        Else
            bReturn = True
        End If
        .Close
    End With
    Set rs = Nothing
    FormENoExists = bReturn
    
    Exit Function
err:
    MsgBox err.Description
    
End Function

