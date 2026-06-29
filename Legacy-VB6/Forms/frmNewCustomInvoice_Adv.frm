VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmNewCustomInvoice_Adv 
   AutoRedraw      =   -1  'True
   ClientHeight    =   10560
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
   ScaleHeight     =   10560
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
      Height          =   10620
      Left            =   90
      TabIndex        =   8
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
         TabIndex        =   133
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
            TabIndex        =   134
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
            TabIndex        =   135
            Top             =   150
            Width           =   12060
         End
      End
      Begin ComboList.Usercontrol1 cmbItemSearchInvoice 
         Height          =   285
         Left            =   1740
         TabIndex        =   129
         Top             =   9090
         Width           =   6150
         _ExtentX        =   10848
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
         TabIndex        =   118
         Top             =   9210
         Visible         =   0   'False
         Width           =   6150
         _ExtentX        =   10848
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
         TabIndex        =   72
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
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   21
         Top             =   1410
         Width           =   1620
      End
      Begin VB.TextBox txtPInvoice 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4575
         MaxLength       =   150
         TabIndex        =   2
         Top             =   1410
         Width           =   2700
      End
      Begin MSComCtl2.DTPicker OrderDT 
         Height          =   300
         Left            =   7290
         TabIndex        =   3
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
         Format          =   154271747
         CurrentDate     =   37250
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   7830
         Left            =   120
         TabIndex        =   14
         Top             =   1740
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   13811
         _Version        =   393216
         Tabs            =   7
         Tab             =   6
         TabsPerRow      =   7
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         TabCaption(0)   =   "Invoice Items"
         TabPicture(0)   =   "frmNewCustomInvoice_Adv.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "txtDTRENo"
         Tab(0).Control(1)=   "txtPrice"
         Tab(0).Control(2)=   "cmdDelete"
         Tab(0).Control(3)=   "LV"
         Tab(0).Control(4)=   "lblGoTo(1)"
         Tab(0).Control(5)=   "lblTotPcs"
         Tab(0).Control(6)=   "Label1(16)"
         Tab(0).Control(7)=   "cmdAddItem1"
         Tab(0).Control(8)=   "Label1(0)"
         Tab(0).Control(9)=   "cmbOrders1"
         Tab(0).ControlCount=   10
         TabCaption(1)   =   "Other Info."
         TabPicture(1)   =   "frmNewCustomInvoice_Adv.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtGDNo"
         Tab(1).Control(1)=   "chkFreeSamples"
         Tab(1).Control(2)=   "chkSampleInvoice"
         Tab(1).Control(3)=   "chkCIFTerms"
         Tab(1).Control(4)=   "chkNoSRO"
         Tab(1).Control(5)=   "txtGatePassNo"
         Tab(1).Control(6)=   "txtLC"
         Tab(1).Control(7)=   "Frame2"
         Tab(1).Control(8)=   "Frame4"
         Tab(1).Control(9)=   "Frame3"
         Tab(1).Control(10)=   "Frame5"
         Tab(1).Control(11)=   "DTLC"
         Tab(1).Control(12)=   "DTGatePass"
         Tab(1).Control(13)=   "DTCustom"
         Tab(1).Control(14)=   "Label1(42)"
         Tab(1).Control(15)=   "Label1(40)"
         Tab(1).Control(16)=   "Label1(39)"
         Tab(1).Control(17)=   "txtInvoiceDescription"
         Tab(1).Control(18)=   "txtDTREDesc"
         Tab(1).Control(19)=   "Label1(32)"
         Tab(1).Control(20)=   "Label1(31)"
         Tab(1).Control(21)=   "Label1(22)"
         Tab(1).Control(22)=   "cmbBankAccNo"
         Tab(1).Control(23)=   "Label1(30)"
         Tab(1).Control(24)=   "txtInvoiceof"
         Tab(1).Control(25)=   "Label1(24)"
         Tab(1).Control(26)=   "cmbBank"
         Tab(1).Control(27)=   "Label1(23)"
         Tab(1).Control(28)=   "Label1(21)"
         Tab(1).Control(29)=   "cmbSMethod"
         Tab(1).ControlCount=   30
         TabCaption(2)   =   "Proforma Items"
         TabPicture(2)   =   "frmNewCustomInvoice_Adv.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtCartonTo"
         Tab(2).Control(1)=   "txtQtyToAdd"
         Tab(2).Control(2)=   "txtCartonQty"
         Tab(2).Control(3)=   "cmbItems"
         Tab(2).Control(4)=   "cmbProformas"
         Tab(2).Control(5)=   "chkShowAllItems"
         Tab(2).Control(6)=   "SSTab2"
         Tab(2).Control(7)=   "txtOrderQty"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "txtBalance"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).Control(9)=   "txtUsed"
         Tab(2).Control(9).Enabled=   0   'False
         Tab(2).Control(10)=   "txtAvailable"
         Tab(2).Control(10).Enabled=   0   'False
         Tab(2).Control(11)=   "txtCartonNo"
         Tab(2).Control(12)=   "txtLotNo"
         Tab(2).Control(13)=   "cmdCallDispatch"
         Tab(2).Control(14)=   "txtEditPLQty"
         Tab(2).Control(15)=   "Command1"
         Tab(2).Control(16)=   "txtQty1"
         Tab(2).Control(17)=   "Frame10"
         Tab(2).Control(18)=   "cmdAdd"
         Tab(2).Control(19)=   "Label1(41)"
         Tab(2).Control(20)=   "lblTotalPcsPI"
         Tab(2).Control(21)=   "Label1(35)"
         Tab(2).Control(22)=   "Label1(8)"
         Tab(2).Control(23)=   "lblGoTo(0)"
         Tab(2).Control(24)=   "cmdAddItem"
         Tab(2).ControlCount=   25
         TabCaption(3)   =   "Commercial Invoice"
         TabPicture(3)   =   "frmNewCustomInvoice_Adv.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame1"
         Tab(3).Control(1)=   "Frame7"
         Tab(3).ControlCount=   2
         TabCaption(4)   =   "Other Charges"
         TabPicture(4)   =   "frmNewCustomInvoice_Adv.frx":0070
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "txtOtherChargesSNo"
         Tab(4).Control(1)=   "txtOtherChargeAmt"
         Tab(4).Control(2)=   "cmdAddOtherCharges"
         Tab(4).Control(3)=   "txtOtherCharges"
         Tab(4).Control(4)=   "LVOtherCharges"
         Tab(4).Control(5)=   "lblCap(0)"
         Tab(4).Control(6)=   "lblCap(1)"
         Tab(4).Control(7)=   "lblCap(2)"
         Tab(4).ControlCount=   8
         TabCaption(5)   =   "Advance Payments"
         TabPicture(5)   =   "frmNewCustomInvoice_Adv.frx":008C
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "txtAdvAmt"
         Tab(5).Control(1)=   "LVAdvance"
         Tab(5).ControlCount=   2
         TabCaption(6)   =   "Addresses"
         TabPicture(6)   =   "frmNewCustomInvoice_Adv.frx":00A8
         Tab(6).ControlEnabled=   -1  'True
         Tab(6).Control(0)=   "Label1(38)"
         Tab(6).Control(0).Enabled=   0   'False
         Tab(6).Control(1)=   "txtShippingAddress"
         Tab(6).Control(1).Enabled=   0   'False
         Tab(6).Control(2)=   "cmbShippingAddresses"
         Tab(6).Control(2).Enabled=   0   'False
         Tab(6).Control(3)=   "Label1(17)"
         Tab(6).Control(3).Enabled=   0   'False
         Tab(6).Control(4)=   "txtBankAddress"
         Tab(6).Control(4).Enabled=   0   'False
         Tab(6).Control(5)=   "txtInvoiceTo"
         Tab(6).Control(5).Enabled=   0   'False
         Tab(6).Control(6)=   "Label1(43)"
         Tab(6).Control(6).Enabled=   0   'False
         Tab(6).Control(7)=   "cmbInvoiceTos"
         Tab(6).Control(7).Enabled=   0   'False
         Tab(6).ControlCount=   8
         Begin VB.TextBox txtGDNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   -73020
            TabIndex        =   194
            Top             =   7485
            Width           =   1695
         End
         Begin VB.TextBox txtCartonTo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   -67245
            TabIndex        =   125
            Top             =   1200
            Width           =   1200
         End
         Begin VB.TextBox txtQtyToAdd 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   -64965
            Locked          =   -1  'True
            TabIndex        =   190
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox txtCartonQty 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   -66060
            TabIndex        =   126
            Top             =   1200
            Width           =   1110
         End
         Begin VB.TextBox txtAdvAmt 
            Height          =   315
            Left            =   -71745
            TabIndex        =   178
            Top             =   1755
            Visible         =   0   'False
            Width           =   2000
         End
         Begin VB.CheckBox chkFreeSamples 
            Caption         =   "Free Samples"
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
            Left            =   -67470
            TabIndex        =   175
            Top             =   7395
            Width           =   1455
         End
         Begin VB.Frame Frame1 
            Height          =   5985
            Left            =   -73485
            TabIndex        =   60
            Top             =   630
            Width           =   9075
            Begin VB.TextBox txtSealNo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   7365
               TabIndex        =   105
               Top             =   405
               Width           =   1620
            End
            Begin VB.TextBox txtContainerNo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   5730
               TabIndex        =   103
               Top             =   405
               Width           =   1620
            End
            Begin VB.TextBox txtVessel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   2895
               TabIndex        =   67
               Top             =   405
               Width           =   2820
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
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
               Left            =   90
               MouseIcon       =   "frmNewCustomInvoice_Adv.frx":00C4
               TabIndex        =   186
               Top             =   4605
               Width           =   8895
            End
            Begin MSForms.TextBox txtComNotify 
               Height          =   1065
               Left            =   90
               TabIndex        =   185
               Top             =   4845
               Width           =   8895
               VariousPropertyBits=   -1400879077
               BorderStyle     =   1
               Size            =   "15690;1879"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox txtComDrawn 
               Height          =   1740
               Left            =   90
               TabIndex        =   183
               Top             =   2880
               Width           =   8895
               VariousPropertyBits=   -1400879077
               BorderStyle     =   1
               Size            =   "15690;3069"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Notes II"
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
               Left            =   90
               MouseIcon       =   "frmNewCustomInvoice_Adv.frx":0216
               TabIndex        =   181
               Top             =   2640
               Width           =   8895
            End
            Begin MSForms.TextBox txtComSpecial 
               Height          =   1740
               Left            =   90
               TabIndex        =   168
               Top             =   915
               Width           =   8895
               VariousPropertyBits=   -1400879077
               BorderStyle     =   1
               Size            =   "15690;3069"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
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
               Left            =   90
               MouseIcon       =   "frmNewCustomInvoice_Adv.frx":0368
               TabIndex        =   167
               Top             =   675
               Width           =   8895
            End
            Begin MSForms.ComboBox cmbPayment 
               Height          =   285
               Left            =   90
               TabIndex        =   111
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
               TabIndex        =   106
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
               TabIndex        =   104
               Top             =   165
               Width           =   1620
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
               TabIndex        =   69
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
               TabIndex        =   68
               Top             =   165
               Width           =   2820
            End
         End
         Begin VB.CheckBox chkSampleInvoice 
            Caption         =   "Sample Invoice"
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
            Left            =   -65820
            TabIndex        =   165
            Top             =   7395
            Width           =   1710
         End
         Begin ComboList.Usercontrol1 cmbItems 
            Height          =   285
            Left            =   -74880
            TabIndex        =   138
            Top             =   690
            Width           =   6750
            _ExtentX        =   11906
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
         Begin ComboList.Usercontrol1 cmbProformas 
            Height          =   285
            Left            =   -68160
            TabIndex        =   122
            Top             =   690
            Width           =   4290
            _ExtentX        =   7567
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
         Begin VB.CheckBox chkShowAllItems 
            Caption         =   "Show All Items"
            Height          =   195
            Left            =   -69795
            TabIndex        =   164
            Top             =   480
            Visible         =   0   'False
            Width           =   1365
         End
         Begin TabDlg.SSTab SSTab2 
            Height          =   5790
            Left            =   -74880
            TabIndex        =   161
            Top             =   1515
            Width           =   11880
            _ExtentX        =   20955
            _ExtentY        =   10213
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Packing List Cartons"
            TabPicture(0)   =   "frmNewCustomInvoice_Adv.frx":04BA
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "LV2"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "txtEditLV2"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "cmdRemoveAll"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).ControlCount=   3
            TabCaption(1)   =   "Import Errors"
            TabPicture(1)   =   "frmNewCustomInvoice_Adv.frx":04D6
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "LVNotMatched"
            Tab(1).ControlCount=   1
            Begin VB.CommandButton cmdRemoveAll 
               Caption         =   "Remove All Items"
               Height          =   360
               Left            =   10320
               TabIndex        =   184
               Top             =   5355
               Width           =   1455
            End
            Begin VB.TextBox txtEditLV2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   915
               TabIndex        =   180
               Top             =   960
               Visible         =   0   'False
               Width           =   705
            End
            Begin MSComctlLib.ListView LV2 
               Height          =   4935
               Left            =   60
               TabIndex        =   162
               Top             =   390
               Width           =   11715
               _ExtentX        =   20664
               _ExtentY        =   8705
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
               NumItems        =   11
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "C. From"
                  Object.Width           =   1270
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   1
                  Text            =   "C. To"
                  Object.Width           =   1270
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   2
                  Text            =   "C. Qty"
                  Object.Width           =   1270
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Proforma No."
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Order No."
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Item"
                  Object.Width           =   10292
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   6
                  Text            =   "Qty"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   7
                  Text            =   "Lot No."
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   8
                  Text            =   "DispatchCarton"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   9
                  Text            =   "ItemCode"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   10
                  Text            =   "FromOrderNo"
                  Object.Width           =   0
               EndProperty
            End
            Begin MSComctlLib.ListView LVNotMatched 
               Height          =   5310
               Left            =   -74925
               TabIndex        =   163
               Top             =   405
               Width           =   11715
               _ExtentX        =   20664
               _ExtentY        =   9366
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
               NumItems        =   5
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Carton #"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Order No."
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Item"
                  Object.Width           =   10292
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   3
                  Text            =   "Qty"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   4
                  Text            =   "Balance"
                  Object.Width           =   1764
               EndProperty
            End
         End
         Begin VB.TextBox txtOrderQty 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   -74880
            Locked          =   -1  'True
            TabIndex        =   142
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1410
         End
         Begin VB.TextBox txtBalance 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   -73485
            Locked          =   -1  'True
            TabIndex        =   141
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1410
         End
         Begin VB.TextBox txtUsed 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   -72090
            Locked          =   -1  'True
            TabIndex        =   140
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1410
         End
         Begin VB.TextBox txtAvailable 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   -70695
            Locked          =   -1  'True
            TabIndex        =   139
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1410
         End
         Begin VB.TextBox txtCartonNo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   -68415
            TabIndex        =   124
            Top             =   1200
            Width           =   1200
         End
         Begin VB.TextBox txtLotNo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   -69300
            TabIndex        =   123
            Top             =   1200
            Width           =   900
         End
         Begin VB.CommandButton cmdCallDispatch 
            Caption         =   "Call from Dispatch"
            Height          =   1065
            Left            =   -63840
            TabIndex        =   137
            Top             =   450
            Width           =   810
         End
         Begin VB.TextBox txtEditPLQty 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   -69240
            TabIndex        =   136
            Top             =   3885
            Visible         =   0   'False
            Width           =   1200
         End
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
            Left            =   -68640
            TabIndex        =   132
            Top             =   7395
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
            Left            =   -64065
            TabIndex        =   131
            Top             =   7395
            Width           =   1020
         End
         Begin VB.TextBox txtDTRENo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   -70350
            TabIndex        =   119
            Top             =   2205
            Visible         =   0   'False
            Width           =   1400
         End
         Begin VB.TextBox txtGatePassNo 
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   -74865
            MaxLength       =   150
            TabIndex        =   116
            Top             =   6435
            Width           =   4095
         End
         Begin VB.TextBox txtPrice 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   -72090
            TabIndex        =   107
            Top             =   795
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtOtherChargesSNo 
            Height          =   315
            Left            =   -74895
            TabIndex        =   82
            Top             =   1020
            Width           =   1200
         End
         Begin VB.TextBox txtOtherChargeAmt 
            Height          =   315
            Left            =   -66630
            TabIndex        =   81
            Top             =   1020
            Width           =   2000
         End
         Begin VB.CommandButton cmdAddOtherCharges 
            Caption         =   "Add"
            Height          =   315
            Left            =   -64595
            TabIndex        =   80
            Top             =   1020
            Width           =   1560
         End
         Begin VB.TextBox txtOtherCharges 
            Height          =   315
            Left            =   -73665
            TabIndex        =   79
            Top             =   1020
            Width           =   7005
         End
         Begin VB.TextBox txtLC 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   -74865
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   6975
            Width           =   1695
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
            Left            =   -68955
            TabIndex        =   48
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
               TabIndex        =   52
               Top             =   210
               Width           =   5745
            End
            Begin MSForms.ComboBox cmbLoadCountry 
               Height          =   285
               Left            =   75
               TabIndex        =   51
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
               TabIndex        =   50
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
               TabIndex        =   49
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
            Left            =   -68970
            TabIndex        =   43
            Top             =   1590
            Width           =   5880
            Begin VB.TextBox txtGate 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   3090
               TabIndex        =   76
               Top             =   960
               Width           =   1155
            End
            Begin VB.TextBox txtWarehouse 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1905
               TabIndex        =   75
               Top             =   960
               Width           =   1155
            End
            Begin VB.TextBox txtManDischarge 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   75
               TabIndex        =   73
               Top             =   960
               Width           =   1800
            End
            Begin MSComCtl2.DTPicker DTETA 
               Height          =   285
               Left            =   4260
               TabIndex        =   77
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
               Format          =   155123715
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
               TabIndex        =   74
               Top             =   720
               Width           =   5745
            End
            Begin MSForms.ComboBox cmbDisPort 
               Height          =   285
               Left            =   3735
               TabIndex        =   47
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
               TabIndex        =   46
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
               TabIndex        =   45
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
               TabIndex        =   44
               Top             =   210
               Width           =   5745
            End
         End
         Begin VB.Frame Frame3 
            Height          =   1485
            Left            =   -74820
            TabIndex        =   32
            Top             =   810
            Width           =   5805
            Begin VB.TextBox txtAWBNo 
               Appearance      =   0  'Flat
               DataField       =   "RptTime"
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   75
               MaxLength       =   150
               TabIndex        =   35
               Top             =   450
               Width           =   1830
            End
            Begin VB.TextBox txtFormE 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   75
               TabIndex        =   34
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
               TabIndex        =   33
               Top             =   450
               Width           =   1830
            End
            Begin MSComCtl2.DTPicker AWBDT 
               Height          =   300
               Left            =   1935
               TabIndex        =   36
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
               Format          =   155123715
               CurrentDate     =   37250
            End
            Begin MSComCtl2.DTPicker FormEDT 
               Height          =   300
               Left            =   1455
               TabIndex        =   37
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
               Format          =   155123715
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
               TabIndex        =   59
               Top             =   870
               Width           =   2580
            End
            Begin MSForms.ComboBox cmbTrans 
               Height          =   285
               Left            =   4425
               TabIndex        =   58
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
               TabIndex        =   57
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
               TabIndex        =   42
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
               TabIndex        =   41
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
               TabIndex        =   40
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
               TabIndex        =   39
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
               TabIndex        =   38
               Top             =   195
               Width           =   1830
            End
         End
         Begin VB.Frame Frame5 
            Height          =   3630
            Left            =   -68970
            TabIndex        =   27
            Top             =   3090
            Width           =   5910
            Begin ComboList.Usercontrol1 cmbInsurAgent 
               Height          =   285
               Left            =   3015
               TabIndex        =   28
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
               TabIndex        =   29
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
               TabIndex        =   31
               Top             =   180
               Width           =   2835
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
               TabIndex        =   30
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
            Picture         =   "frmNewCustomInvoice_Adv.frx":04F2
            Style           =   1  'Graphical
            TabIndex        =   24
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
            Left            =   -69345
            TabIndex        =   23
            Top             =   120
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtQty 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   -70800
            TabIndex        =   22
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
            Picture         =   "frmNewCustomInvoice_Adv.frx":0A24
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Remove This Entry From Voucher"
            Top             =   1305
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   345
         End
         Begin MSComctlLib.ListView LV 
            Height          =   6630
            Left            =   -74685
            TabIndex        =   15
            Top             =   720
            Width           =   11340
            _ExtentX        =   20003
            _ExtentY        =   11695
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
         Begin MSComctlLib.ListView LVOtherCharges 
            Height          =   5280
            Left            =   -74895
            TabIndex        =   83
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
            Left            =   -73155
            TabIndex        =   110
            Top             =   6975
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
            Format          =   155123715
            CurrentDate     =   37250
         End
         Begin MSComCtl2.DTPicker DTGatePass 
            Height          =   300
            Left            =   -70770
            TabIndex        =   117
            Top             =   6435
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
            Format          =   155123715
            CurrentDate     =   37250
         End
         Begin VB.Frame Frame10 
            Caption         =   "Dispatch List Detail :"
            Height          =   825
            Left            =   -74880
            TabIndex        =   153
            Top             =   2415
            Visible         =   0   'False
            Width           =   11835
            Begin VB.CommandButton cmdBreakLink 
               Caption         =   "Break Link"
               Height          =   645
               Left            =   10770
               TabIndex        =   160
               Top             =   135
               Width           =   975
            End
            Begin VB.TextBox txtDispatchListQtyBalance 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   9285
               TabIndex        =   159
               Top             =   435
               Width           =   1440
            End
            Begin VB.TextBox txtDispatchListQtyUsed 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   7860
               TabIndex        =   158
               Top             =   435
               Width           =   1440
            End
            Begin VB.TextBox txtDispatchListQty 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   6435
               TabIndex        =   157
               Top             =   435
               Width           =   1440
            End
            Begin VB.TextBox txtDispatchListCartonNo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   5010
               TabIndex        =   156
               Top             =   435
               Width           =   1440
            End
            Begin VB.TextBox txtDispatchListNo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   90
               TabIndex        =   155
               Text            =   "Not Linked with any Dispatch List"
               Top             =   435
               Width           =   4935
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   $"frmNewCustomInvoice_Adv.frx":0F56
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
               Index           =   37
               Left            =   90
               TabIndex        =   154
               Top             =   195
               Width           =   10635
            End
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
            Height          =   1065
            Left            =   -63840
            TabIndex        =   166
            Top             =   435
            Width           =   810
         End
         Begin MSComctlLib.ListView LVAdvance 
            Height          =   7320
            Left            =   -74940
            TabIndex        =   179
            Top             =   405
            Width           =   11940
            _ExtentX        =   21061
            _ExtentY        =   12912
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
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S No."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Advice #"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Amount"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Already Used"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "This Inv."
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Balance"
               Object.Width           =   3528
            EndProperty
         End
         Begin VB.Frame Frame7 
            Height          =   2190
            Left            =   -74160
            TabIndex        =   61
            Top             =   3720
            Visible         =   0   'False
            Width           =   8655
            Begin VB.TextBox txtComPayTerms 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1410
               TabIndex        =   170
               Top             =   1065
               Visible         =   0   'False
               Width           =   2820
            End
            Begin VB.CommandButton cmdSwap 
               Caption         =   "<< Swap>>"
               Height          =   240
               Left            =   2955
               TabIndex        =   66
               Top             =   150
               Width           =   1815
            End
            Begin MSForms.TextBox txtComAgainst 
               Height          =   1740
               Left            =   360
               TabIndex        =   182
               Tag             =   "DeclarationsGarments"
               Top             =   1920
               Width           =   8895
               VariousPropertyBits=   -1400879077
               BorderStyle     =   1
               Size            =   "15690;3069"
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
               Left            =   7440
               TabIndex        =   174
               Top             =   720
               Width           =   2835
            End
            Begin MSForms.ComboBox cmbThrough 
               Height          =   285
               Left            =   7440
               TabIndex        =   173
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
               Left            =   7305
               TabIndex        =   172
               Top             =   225
               Width           =   2910
            End
            Begin MSForms.ComboBox cmbConsignee 
               Height          =   285
               Left            =   7305
               TabIndex        =   171
               Top             =   465
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
               Left            =   855
               MouseIcon       =   "frmNewCustomInvoice_Adv.frx":1016
               TabIndex        =   169
               Top             =   3090
               Visible         =   0   'False
               Width           =   5730
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
               MouseIcon       =   "frmNewCustomInvoice_Adv.frx":1168
               MousePointer    =   99  'Custom
               TabIndex        =   78
               Top             =   2175
               Width           =   1065
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
               TabIndex        =   65
               Top             =   150
               Width           =   2625
            End
            Begin MSForms.TextBox txtComConsignee 
               Height          =   1680
               Left            =   630
               TabIndex        =   64
               Top             =   1050
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
               TabIndex        =   63
               Top             =   150
               Width           =   1170
            End
            Begin MSForms.ComboBox cmbComConsignee 
               Height          =   255
               Left            =   1200
               TabIndex        =   62
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
         Begin MSComCtl2.DTPicker DTCustom 
            Height          =   300
            Left            =   -74865
            TabIndex        =   187
            Top             =   7485
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
            Format          =   155123715
            CurrentDate     =   37250
         End
         Begin MSForms.ComboBox cmbInvoiceTos 
            Height          =   285
            Left            =   375
            TabIndex        =   202
            Top             =   4995
            Width           =   11325
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "19976;503"
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
            Caption         =   " Invoice To"
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
            Index           =   43
            Left            =   375
            TabIndex        =   201
            Top             =   4755
            Width           =   11325
         End
         Begin MSForms.TextBox txtInvoiceTo 
            Height          =   2340
            Left            =   375
            TabIndex        =   200
            Top             =   5265
            Width           =   11325
            VariousPropertyBits=   -1400879077
            BorderStyle     =   1
            Size            =   "19976;4128"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBankAddress 
            Height          =   1065
            Left            =   375
            TabIndex        =   199
            Top             =   870
            Width           =   11325
            VariousPropertyBits=   -1400879077
            BorderStyle     =   1
            Size            =   "19976;1879"
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
            Index           =   17
            Left            =   375
            TabIndex        =   198
            Top             =   630
            Width           =   11325
         End
         Begin MSForms.ComboBox cmbShippingAddresses 
            Height          =   285
            Left            =   375
            TabIndex        =   197
            Top             =   2160
            Width           =   11325
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "19976;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtShippingAddress 
            Height          =   2340
            Left            =   375
            TabIndex        =   196
            Top             =   2430
            Width           =   11325
            VariousPropertyBits=   -1400879077
            BorderStyle     =   1
            Size            =   "19976;4128"
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
            Caption         =   " Shipping Address"
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
            Index           =   38
            Left            =   375
            TabIndex        =   195
            Top             =   1920
            Width           =   11325
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "GD No."
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
            Index           =   42
            Left            =   -73020
            TabIndex        =   193
            Top             =   7245
            Width           =   1770
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
            Index           =   41
            Left            =   -65805
            TabIndex        =   192
            Top             =   7410
            Width           =   1005
         End
         Begin VB.Label lblTotalPcsPI 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E7EBEF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   -64815
            TabIndex        =   191
            Top             =   7410
            Width           =   1800
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"frmNewCustomInvoice_Adv.frx":12BA
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
            Index           =   35
            Left            =   -74880
            TabIndex        =   189
            Top             =   960
            Width           =   11010
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " GD/Custom Date"
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
            Index           =   40
            Left            =   -74865
            TabIndex        =   188
            Top             =   7245
            Width           =   1770
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Invoice Description"
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
            Index           =   39
            Left            =   -74865
            TabIndex        =   177
            Top             =   5130
            Width           =   5865
         End
         Begin MSForms.TextBox txtInvoiceDescription 
            Height          =   840
            Left            =   -74865
            TabIndex        =   176
            Top             =   5370
            Width           =   5865
            VariousPropertyBits=   -1400879077
            BorderStyle     =   1
            Size            =   "10345;1482"
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
            Caption         =   $"frmNewCustomInvoice_Adv.frx":1361
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
            Left            =   -74880
            TabIndex        =   143
            Top             =   450
            Width           =   11010
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
            TabIndex        =   128
            Top             =   7500
            Width           =   1185
         End
         Begin MSForms.TextBox txtDTREDesc 
            Height          =   1395
            Left            =   -74865
            TabIndex        =   121
            Top             =   3765
            Width           =   5865
            VariousPropertyBits=   -1400879077
            BorderStyle     =   1
            Size            =   "10345;2461"
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
            Left            =   -74865
            TabIndex        =   120
            Top             =   3525
            Width           =   5865
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Gate-Pass #                                                                   Date"
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
            Left            =   -74865
            TabIndex        =   115
            Top             =   6195
            Width           =   5865
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
            TabIndex        =   114
            Top             =   7350
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
            Left            =   -66765
            TabIndex        =   113
            Top             =   6735
            Width           =   3720
         End
         Begin MSForms.ComboBox cmbBankAccNo 
            Height          =   285
            Left            =   -66765
            TabIndex        =   112
            Top             =   6975
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
            Left            =   -74865
            TabIndex        =   109
            Top             =   2325
            Width           =   5865
         End
         Begin MSForms.TextBox txtInvoiceof 
            Height          =   975
            Left            =   -74865
            TabIndex        =   108
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
            TabIndex        =   86
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
            TabIndex        =   85
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
            TabIndex        =   84
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
            Left            =   -69585
            TabIndex        =   71
            Top             =   6735
            Width           =   2835
         End
         Begin MSForms.ComboBox cmbBank 
            Height          =   285
            Left            =   -69585
            TabIndex        =   70
            Top             =   6975
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
            Left            =   -74865
            TabIndex        =   56
            Top             =   6735
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
            Left            =   -71370
            TabIndex        =   54
            Top             =   6735
            Width           =   1800
         End
         Begin MSForms.ComboBox cmbSMethod 
            Height          =   285
            Left            =   -71370
            TabIndex        =   53
            Top             =   6975
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
         Begin VB.Label lblTotPcs 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E7EBEF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   -65160
            TabIndex        =   26
            Top             =   7455
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
            TabIndex        =   25
            Top             =   7455
            Width           =   1005
         End
         Begin MSForms.CommandButton cmdAddItem 
            Height          =   345
            Left            =   -71190
            TabIndex        =   127
            Top             =   345
            Visible         =   0   'False
            Width           =   930
            Caption         =   "    Add"
            PicturePosition =   327683
            Size            =   "1640;609"
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
            TabIndex        =   19
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
            TabIndex        =   18
            Top             =   2370
            Visible         =   0   'False
            Width           =   1755
         End
         Begin MSForms.ComboBox cmbOrders1 
            Height          =   285
            Left            =   -74640
            TabIndex        =   17
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
         TabIndex        =   9
         Top             =   9885
         Width           =   6960
         Begin MSForms.CommandButton cmdSavenStay 
            Height          =   360
            Left            =   3495
            TabIndex        =   130
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
            TabIndex        =   7
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
            TabIndex        =   5
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
            TabIndex        =   6
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
         TabIndex        =   87
         Top             =   840
         Width           =   1410
      End
      Begin MSForms.TextBox lblTotWords 
         Height          =   270
         Left            =   1890
         TabIndex        =   20
         Top             =   9600
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
         TabIndex        =   13
         Top             =   9600
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
         TabIndex        =   12
         Top             =   9600
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
         TabIndex        =   11
         Top             =   9600
         Width           =   1380
      End
      Begin MSForms.ComboBox cmbTradeTerms 
         Height          =   285
         Left            =   8895
         TabIndex        =   4
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
         Caption         =   $"frmNewCustomInvoice_Adv.frx":140B
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
         TabIndex        =   10
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
      Height          =   9045
      Left            =   3600
      TabIndex        =   88
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
         MouseIcon       =   "frmNewCustomInvoice_Adv.frx":14C7
         MousePointer    =   99  'Custom
         TabIndex        =   102
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
         MouseIcon       =   "frmNewCustomInvoice_Adv.frx":1619
         MousePointer    =   99  'Custom
         TabIndex        =   101
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
         MouseIcon       =   "frmNewCustomInvoice_Adv.frx":176B
         MousePointer    =   99  'Custom
         TabIndex        =   100
         Top             =   4050
         Width           =   5730
      End
      Begin MSForms.TextBox txtComDeclarationGarments 
         Height          =   585
         Left            =   315
         TabIndex        =   99
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
         TabIndex        =   98
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
         TabIndex        =   97
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
         MouseIcon       =   "frmNewCustomInvoice_Adv.frx":18BD
         MousePointer    =   99  'Custom
         TabIndex        =   96
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
         MouseIcon       =   "frmNewCustomInvoice_Adv.frx":1A0F
         MousePointer    =   99  'Custom
         TabIndex        =   95
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
         MouseIcon       =   "frmNewCustomInvoice_Adv.frx":1B61
         MousePointer    =   99  'Custom
         TabIndex        =   94
         Top             =   285
         Width           =   5805
      End
      Begin MSForms.TextBox txtDeclarationGarments 
         Height          =   540
         Left            =   45
         TabIndex        =   93
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
         TabIndex        =   92
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
         TabIndex        =   91
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
         MouseIcon       =   "frmNewCustomInvoice_Adv.frx":1CB3
         MousePointer    =   99  'Custom
         TabIndex        =   90
         Top             =   2595
         Width           =   1875
      End
      Begin MSForms.TextBox txtManDeclaration 
         Height          =   915
         Left            =   1980
         TabIndex        =   89
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
   Begin VB.Frame FraDispatch 
      BackColor       =   &H00E7EBEF&
      Caption         =   "Dispatch List"
      Height          =   4290
      Left            =   2025
      TabIndex        =   144
      Top             =   4230
      Visible         =   0   'False
      Width           =   8820
      Begin ComboList.Usercontrol1 cmbDispatchNo 
         Height          =   285
         Left            =   60
         TabIndex        =   145
         Top             =   480
         Width           =   2520
         _ExtentX        =   4445
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
      Begin VB.TextBox txtDispQtyTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3660
         Locked          =   -1  'True
         TabIndex        =   152
         Top             =   3900
         Width           =   1230
      End
      Begin VB.CommandButton cmdOKDispatch 
         Caption         =   "OK"
         Height          =   360
         Left            =   7680
         TabIndex        =   151
         Top             =   3855
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancelDispatch 
         Caption         =   "Cancel"
         Height          =   360
         Left            =   6525
         TabIndex        =   150
         Top             =   3855
         Width           =   1050
      End
      Begin MSComctlLib.ListView LVDL 
         Height          =   3105
         Left            =   60
         TabIndex        =   149
         Top             =   750
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   5477
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15199215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S #"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Inner #"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Order #"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Qty"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.TextBox txtItemCartonQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4005
         Locked          =   -1  'True
         TabIndex        =   148
         Top             =   480
         Width           =   1230
      End
      Begin MSForms.ComboBox cmbCartonNo 
         Height          =   285
         Left            =   2580
         TabIndex        =   147
         Top             =   480
         Width           =   1440
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2540;503"
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
         Caption         =   " Dispatch List No.                        Carton No.                    Qty"
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
         Index           =   36
         Left            =   60
         TabIndex        =   146
         Top             =   240
         Width           =   8670
      End
   End
End
Attribute VB_Name = "frmNewCustomInvoice_Adv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Orders As New Collection
Dim strProforma As String
Dim bCustomInvoice As Boolean
Dim strDelList As String
Dim bCourierInvoice As Boolean
Dim iColNo As Integer, iColNoLV2 As Integer
Dim bAcceptsExtraQty As Boolean
Dim strDelListPL, strEditListPL As String
Dim bItemSelectionFromDispatch As Boolean

Public Sub EditProforma(TempProforma As String, Optional p_bCustomInvoice As Boolean = True)
         
    On Error GoTo err
    
    bCustomInvoice = p_bCustomInvoice
    
    If bCustomInvoice Then
        Label1(33).Caption = "New Custom Invoice"
        Label1(34).Caption = "New Custom Invoice"
    Else
        Label1(33).Caption = "New Commercial Invoice"
        Label1(34).Caption = "New Commercial Invoice"
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
        txtShippingAddress = !ComAgainst & ""
        txtComDrawn = ![ComDrawnUnder] & ""
        txtComSpecial = ![ComSpecial] & ""
        txtManDeclaration = ![ManDeclaration] & ""
        txtInvoiceof = !InvoiceOf & ""
        txtInvoiceDescription = !InvoiceDescription & ""
        
        txtGatePassNo = !GatePassNo & ""
        If IsNull(!GatePassDT) = False Then
            DTGatePass = !GatePassDT
        End If
        
        If IsNull(!CustomDT) = False Then
            DTCustom = !CustomDT
        End If
        
        txtDTREDesc = !DTREDescription & ""
        txtGWeight = Val(!GrossWeight & "")
        
        chkNoSRO = Abs(!NoSRO)
        
        txtBatchNo = !TotalCartons & ""
        
        chkCIFTerms = Abs(!CIFTradeTerms)
        chkSampleInvoice = Abs(!SampleInvoice)
        chkFreeSamples = Abs(!FreeSamples)
        
        txtGDNo = !GDNo & ""
        txtInvoiceTo = !InvoiceTo & ""
        
        .Close
        
'        chkCIFTerms = Abs(GetSingleBooleanValue("CIFTradeTerms", "CustomInvoice", " WHERE CustomInvoice='" & strProforma & "'"))
'        chkSampleInvoice = Abs(GetSingleBooleanValue("SampleInvoice", "CustomInvoice", " WHERE CustomInvoice='" & strProforma & "'"))
'
        
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
            
            ITM.ListSubItems.add(, , !CustItemID & "").Tag = Val(![EntryID] & "") '![CatID] & ""
            ITM.ListSubItems.add , , ![OrderNo] '![CatID] & ""
            ITM.ListSubItems.add , , ![ItemSize] & "" '![ItemID] & ""
            ITM.ListSubItems.add , , "" '![Color] & ""
            ITM.ListSubItems.add(, , "").Tag = "" '![ArticleNo] & ""
            If bCustomInvoice Then
                'ITM.ListSubItems.add(, ![ItemCustomName] & "'", ![ItemCustomName] & "").Tag = ![ItemCustomName] & ""
                ITM.ListSubItems.add(, ![ItemDescription] & "'", ![ItemDescription] & "").Tag = ![ItemDescription] & ""
                dPrice = Val(!CustomPrice & "")
            Else
                ITM.ListSubItems.add(, ![ItemDescription] & "'", ![ItemDescription] & "").Tag = ![ItemDescription] & ""
                dPrice = Val(!Price & "")
            End If
            
            ITM.ListSubItems.add , , !ITemID & "" '![Unit] & ""
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
         
        .Open "SELECT * FROM VCustomPList WHERE CustomInvoice='" & strProforma & "' ORDER BY ID", con, adOpenForwardOnly, adLockReadOnly
        LV2.ListItems.Clear
        Dim lQty As Long
        Do Until .EOF
            Set ITM = LV2.ListItems.add(, , Val(!CartonFrom & ""))
            ITM.Tag = Val(!ID & "")
            ITM.ListSubItems.add , , Val(!CartonTo & "")
            ITM.ListSubItems.add , , Val(!Qty & "")
            ITM.ListSubItems.add , , !PInvoice & ""
            ITM.ListSubItems.add , , !OrderNo & ""
            ITM.ListSubItems.add , , !ITemID & " - " & !ItemName
            lQty = (Val(!CartonTo & "") - Val(!CartonFrom & "") + 1) * Val(!Qty & "")
            ITM.ListSubItems.add(, , lQty).Tag = lQty
            ITM.ListSubItems.add , , !LotNo & ""
            ITM.ListSubItems.add(, , Val(!CartonNo & "")).Tag = Val(!DP_RefID & "")
            ITM.ListSubItems.add , , !ItemCode & ""
            ITM.ListSubItems.add , , !FromOrderNo & ""
            .MoveNEXT
        Loop
        .Close
        
        .Open "Select * From CustomInvoiceOtherCharges WHERE CustomInvoice='" & strProforma & "' ORDER BY SNo", con_ServerSide, adOpenForwardOnly, adLockReadOnly
        LVOtherCharges.ListItems.Clear
        Do Until .EOF
            Set ITM = LVOtherCharges.ListItems.add(, , Val(!SNo & ""))
            ITM.SubItems(1) = !ChargesDetail & ""
            ITM.SubItems(2) = Val(!ChargesAmt & "")
            .MoveNEXT
        Loop
        .Close
        
        'advance payments
        .Open "Select * From VFCustAdvancePaymentsCustomInvoice Where CustomInvoice='" & strProforma & "' Order By EntryID", con, adOpenForwardOnly
        
        Do Until .EOF
            On Error Resume Next
            Set ITM = LVAdvance.ListItems.add(, !EntryID & "'")
            If err.Number = 35602 Then
                Set ITM = LVAdvance.ListItems(!EntryID & "'")
                ITM.SubItems(4) = Val(!AmountUsed & "") 'This Invoice
                ITM.SubItems(3) = Val(ITM.SubItems(3)) - Val(!AmountUsed & "")  'Used
                ITM.SubItems(5) = Val(ITM.SubItems(2)) - Val(ITM.SubItems(3)) - Val(ITM.SubItems(4)) 'Balance
            Else
                Set ITM = LVAdvance.ListItems.add(, !EntryID & "'", LVAdvance.ListItems.count + 1)
                ITM.ListSubItems.add , , !AdviceNo & ""
                ITM.ListSubItems.add(, , Val(!Amount & "")).Tag = Val(!Amount & "")
                ITM.ListSubItems.add(, , Val(!AdjustedAmount & "")).Tag = Val(!AdjustedAmount & "")
                ITM.ListSubItems.add(, , Val(!AmountUsed & "")).Tag = Val(!AmountUsed & "")
                ITM.ListSubItems.add , , 0
            End If
            
            .MoveNEXT
        Loop
        .Close
        
        txtOtherChargesSNo.Text = LVOtherCharges.ListItems.count + 1
    End With
    Call CalculateTotals
   
    cmbItemSearchInvoice.DropDownOffLine = True
  
    If strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
        cmbItemSearchInvoice.AddVals con_ServerSide, "ItemID + ' {' + OrderNo + '} {' + CONVERT(varchar(50),DeliveryDTItem,3) + '}'", "VCustomInvoiceItems", "ItemID+'|'+OrderNo+'|'+CONVERT(varchar(50),DeliveryDTItem,3)", "WHERE CustomInvoice='" & strProforma & "' " & strOrderBy
    Else
        cmbItemSearchInvoice.AddVals con_ServerSide, "ItemID + ' {' + OrderNo + '}'", "VCustomInvoiceItems", "ItemID+'|'+OrderNo", "WHERE CustomInvoice='" & strProforma & "' " & strOrderBy
    End If
     
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

Private Sub chkShowAllItems_Click()
    If chkShowAllItems.Value = vbChecked Then
        cmbItems.AddVals con_ServerSide, "'{'+ItemID+'} '+ItemName", "Items", "ItemID"
        chkShowAllItems.Visible = False
    End If
End Sub

Private Sub cmbBank_Change()
    Call cmbBank_Click
End Sub

Private Sub cmbBank_Click()
    cmbBankAccNo.Clear
    Call AddToCombo(cmbBankAccNo, "AccNo + ' ( ' + Branch + ' )'", "BankList", " WHERE CompanyRefID=" & cmbCompany.ID & " AND Bank='" & cmbBank.Text & "'", , "BankID")
    If cmbBankAccNo.ListCount > 0 Then cmbBankAccNo.ListIndex = 0
End Sub

Private Sub cmbCartonNo_Change()
    Call cmbCartonNo_Click
End Sub

Private Sub cmbCartonNo_Click()
    
    If cmbCartonNo.MatchFound = False Then Exit Sub
    txtItemCartonQty = GetSingleLongValue("SUM(InnerQty)", "VDispatchListDetail_Ex", " WHERE CartonNo=" & cmbCartonNo.Text & " AND DispatchListEntryID=" & cmbDispatchNo.ID & " AND ItemCode='" & cmbItems.ID & "'")
    
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        .Open "SELECT * FROM VDispatchListDetail_Ex WHERE DispatchListEntryID=" & cmbDispatchNo.ID & " AND CartonNo=" & cmbCartonNo.Text & " AND ItemCode='" & cmbItems.ID & "' ORDER BY InnerEntryID", con, adOpenForwardOnly, adLockReadOnly
        LVDL.ListItems.Clear
        Do Until .EOF
            Set ITM = LVDL.ListItems.add(, , LVDL.ListItems.count + 1)
            ITM.Checked = True
            ITM.Tag = !InnerEntryID
            ITM.ListSubItems.add , , !InnerNo & ""
            ITM.ListSubItems.add , , !Orders & ""
            ITM.ListSubItems.add , , Val(!InnerQty & "")
            
            
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
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
'        Dim InvoiecNo As String
'        InvoiceNo = GetInvoiceNo(cmbCompany.ID, False, OrderDT.Value)
'
'        If bCourierInvoice Then
'            txtPInvoice = strCourierInvoicePrefix & " " & Format(InvoiceNo, "000")
'        Else
'            txtPInvoice.Tag = InvoiceNo
'            txtPInvoice = txtPInvoice.Tag
'        End If
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

'    If cmbConsignee = "Customer Address" Then
'        txtBankAddress = txtBankAddress.Tag
'        cmbThrough.Text = ""
'        cmbThrough.Enabled = False
'    ElseIf cmbConsignee = "Bank Address" Then
'        cmbThrough.Enabled = True
'        If cmbThrough.ListCount > 0 Then cmbThrough.ListIndex = 0
'    End If
    
End Sub

Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()
    
    On Error GoTo err
    If cmbCust.MatchFound = False Then Exit Sub
    If cmbCountry.MatchFound = False Then Exit Sub
    
    txtCurrency = cmbCountry.List(cmbCountry.ListIndex, 1)
'    Call GetOrders
    Call GetBanks
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select FakeAddress From ForeignCustomers Where CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry & "'", con, adOpenForwardOnly, adLockReadOnly
        txtBankAddress.Tag = ![FakeAddress] & ""
        .Close
        .Open "SELECT AcceptsExtraQty,DefaultShipMethod,Address,PaymentTerms FROM ForeignCustomers WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        bAcceptsExtraQty = !AcceptsExtraQty
        cmbSMethod = !DefaultShipMethod & ""
        txtBankAddress = !Address & ""
        txtComNotify = !Address & ""
        cmbPayment = !PaymentTerms & ""
        .Close
    End With
    cmbDisCountry = cmbCountry
    Set rs = Nothing
   
    Call AddToCombo(cmbShippingAddresses, "ShippingAddress", "ForeignCustomersShippingAddresses", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'", , "EntryID")
   
    cmbItems.ClearVals
    cmbItems.AddVals con_ServerSide, "'{'+CompItemID+'} {'+ItemID+'} '+CompDesc", "VFCustomerCatalog", "CompItemID", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'"
    
    'Add Advance Payments....
    Call FillAdvancePayments
        
    Call AddToCombo(cmbInvoiceTos, "InvoiceTo", "ForeignCustomers_InvoiceTo", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'", , "EntryID")
    
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
    
'    If strProforma = "" Then
'        If strCompany = "Sunlike" Then
'            txtPInvoice = Left(txtPInvoice, 8) & Mid(cmbCust.ID, 4)
'        ElseIf strCompany = "Dr-Frgz" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
'            txtPInvoice = txtPInvoice.Tag & "/" & cmbCust.ID
'        ElseIf strCompany = "QEL" Or strCompany = "Tecno" Then
'            'Do Nothing,It's not Customer Specific..
'        ElseIf strCompany <> "Banzai" And strCompany <> "Instrumed" Then
'            txtPInvoice = Left(txtPInvoice, 8) & Right(cmbCust.ID, 4)
'        End If
''        If Len(txtPInvoice) > 8 Then
''            txtPInvoice = Left(txtPInvoice, 8) & Right(cmbCust.ID, 4)
''        Else
''            txtPInvoice = Left(txtPInvoice, 8) & Right(cmbCust.ID, 4)
''        End If
'    End If
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbCust_UnMatched()
    'cmbCountry.ClearVals
End Sub


Private Sub cmbDispatchNo_matched()

    cmbCartonNo.Clear
    Call AddToCombo(cmbCartonNo, "DISTINCT CartonNo", "VDispatchListDetail_Ex", " WHERE DispatchListEntryID=" & cmbDispatchNo.ID & " AND ItemCode='" & cmbItems.ID & "' ORDER BY CartonNo")
    
End Sub

Private Sub cmbItems_matched()

    On Error GoTo err
        
    If bItemSelectionFromDispatch = False Then
        Call cmdBreakLink_Click
    End If
    
    cmbProformas.ClearVals
    
    If strCompany = "Dr-Frgz" And chkSampleInvoice.Enabled = False Then
        cmbProformas.AddVals con_ServerSide, "PInvoice+' {'+OrderNo+'}-{'+CAST(InternalRefNo AS VARCHAR(255))+'} {'+CAST(InvQty AS VARCHAR(50))+'}'", "VFProformas", "EntryID", " WHERE ItemID='" & cmbItems.ID & "' AND CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "' AND OrderType='Sample Order'"
    Else
        cmbProformas.AddVals con_ServerSide, "PInvoice+' {'+OrderNo+'}-{'+CAST(InternalRefNo AS VARCHAR(255))+'} {'+CAST(InvQty AS VARCHAR(50))+'}'", "VFProformas", "EntryID", " WHERE ItemID='" & cmbItems.ID & "' AND CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "' AND InvQty>0"
    End If
    
    
    
    Exit Sub
err:
    MsgBox err.Description
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
            strSearch = LV.ListItems(i).SubItems(7) & "|" & LV.ListItems(i).SubItems(2)
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

Private Sub cmbProformas_matched()

    Dim lQtyUsed As Long
    Dim i As Integer
    For i = 1 To LV2.ListItems.count
        If Val(LV2.ListItems(i).ListSubItems(1).Tag) = Val(cmbProformas.ID) Then
            lQtyUsed = lQtyUsed + Val(LV2.ListItems(i).SubItems(4))
        End If
    Next
       
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT Qty,InvQty FROM VFProformas WHERE EntryID=" & cmbProformas.ID, con, adOpenForwardOnly, adLockReadOnly
        txtOrderQty = Val(!Qty & "")
        txtBalance = Val(!InvQty & "")
        txtUsed = lQtyUsed
        txtAvailable = Val(txtBalance) - Val(txtUsed)
        .Close
    End With
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbShippingAddresses_Change()
    Call cmbShippingAddresses_Click
End Sub

Private Sub cmbShippingAddresses_Click()

    If cmbShippingAddresses.MatchFound = False Then
        txtShippingAddress = ""
    Else
        txtShippingAddress = GetSingleStringValue("ShippingAddress", "ForeignCustomersShippingAddresses", " WHERE EntryID=" & cmbShippingAddresses.List(cmbShippingAddresses.ListIndex, 1))
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
        
        cmbDesc.Locked = False
        
         
        
        
    Else
        cmbCatID.Locked = True
        
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

Private Sub cmdAdd_Click()
    Call txtQtyToAdd_KeyPress(13)
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

Private Sub cmdBreakLink_Click()

    txtDispatchListNo.Text = "Not Linked with any Dispatch List"
    txtDispatchListNo.Tag = "0"
    txtDispatchListCartonNo = ""
    txtDispatchListQty = ""
    txtDispatchListQtyUsed = ""
    txtDispatchListQtyBalance = ""
    
End Sub

Private Sub cmdCallDispatch_Click()
'    If cmbItems.MatchFound = False Then
'        MsgBox "Please Select Item.", vbInformation
'        Exit Sub
'    Else
'        If cmbDispatchNo.MatchFound = False Then
'            cmbDispatchNo.ClearVals
'            cmbDispatchNo.AddVals con, "DISTINCT DispatchListNo", "VDispatchListDetail_Ex", "DispatchListEntryID", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "' AND ItemCode='" & cmbItems.ID & "'"
'        End If
'        FAll.Enabled = False
'        FraDispatch.Visible = True
'    End If
    Dim f As New frmGetSingleComboValue, strDispatchListNo As String
    Load f
    Call f.cmbCombo.AddVals(con, "DispatchListNo+'{'+CAST(EntryID AS VARCHAR(10))+'}'", "DispatchList", "EntryID", " WHERE CustCode='" & cmbCust.ID & "' AND Finalyzed=1 AND EntryID NOT IN(SELECT DP_RefID FROM CustomPList_DispatchListDetail)")
    If f.getComboValue(strDispatchListNo, "Dispatch List", "Dispatch List") = False Then Exit Sub
    
    Dim lEntryID As Long, ITM As ListItem, strLotNo As String
    lEntryID = Val(strDispatchListNo)
    Dim lTotalQty_PI As Long
    Dim rs As New ADODB.Recordset, rsCartons As New ADODB.Recordset, rsProformas As New ADODB.Recordset, lAvailableQty As Long, lQtyToUse As Long, lTempUsedQty As Long
    With rs
        .Open "SELECT * FROM VDispatchListToExport WHERE DispatchListNo=" & lEntryID & " AND DispatchQty<=InvQty", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            With rsCartons
                .Open "SELECT CartonNo,OrderNo,ItemCode,LotNo,SUM(Qty) AS CartonQty FROM VDispatchListDetail_Ex_OrderWise " & _
                    "WHERE DispatchListEntryID=" & lEntryID & " AND OrderNo='" & rs!OrderNo & "' AND ItemCode='" & rs!ItemCode & "' GROUP BY CartonNo,OrderNo,ItemCode,LotNo ORDER BY CartonNo", con, adOpenForwardOnly, adLockReadOnly
                
                Do Until .EOF
                    With rsProformas
                        .Open "SELECT EntryID,OrderNo,PInvoice,ItemID,ItemName,Qty,InvQty,Price,CustomPrice " & _
                         "FROM VFProformas WHERE OrderNo='" & rsCartons!OrderNo & "' AND ItemID='" & rsCartons!ItemCode & "' AND InvQty>0", con, adOpenForwardOnly, adLockReadOnly
                        lAvailableQty = Val(rsCartons!CartonQty & "")
                         
                        
                        
                        Do Until (.EOF Or (lAvailableQty <= 0))
                                                            
                            '--------------------------------------------------
                            lTempUsedQty = getTempUsedQty(Val(!EntryID & ""))
                            If (Val(!InvQty & "") - lTempUsedQty) <= 0 Then GoTo MoveToNext
                            '--------------------------------------------------
                            Set ITM = LV2.ListItems.add(, , Val(rsCartons!CartonNo & ""))
                            ITM.ListSubItems.add(, , Val(rsCartons!CartonNo & "")).Tag = Val(!EntryID & "")
                            
                            If (Val(!InvQty & "") - lTempUsedQty) > lAvailableQty Then
                                lQtyToUse = lAvailableQty
                            Else
                                lQtyToUse = Val(!InvQty & "") - lTempUsedQty
                            End If
                            lAvailableQty = lAvailableQty - lQtyToUse
                            
                            ITM.ListSubItems.add(, , lQtyToUse).Tag = Val(!Price & "")
                            
                            ITM.ListSubItems.add(, , !PInvoice & "").Tag = Val(!CustomPrice & "")
                            ITM.ListSubItems.add(, , !OrderNo & "").Tag = ""
                            ITM.ListSubItems.add(, , !ITemID & " - " & !ItemName).Tag = ""
                            
                            ITM.ListSubItems.add(, , lQtyToUse).Tag = lEntryID
'                            If strCompany = "IAA" Or strCompany = "Kami" Then
'                                strLotNo = GetSingleStringValue("LotNo", "VDispatchListDetail_Adv", " WHERE ItemCOde='" & rs!ItemCode & "' AND OrderNo='" & rs!OrderNo & "' AND CartonNo='" & rsCartons!CartonNo & "' ORDER BY Qty DESC")
'                            Else
'                                strLotNo = ""
'                            End If
                            strLotNo = rsCartons!LotNo & ""
                            ITM.ListSubItems.add , , strLotNo
                            
                            'Following 2 Lines are for Dispatch...
                            ITM.ListSubItems.add(, , Val(rsCartons!CartonNo & "")).Tag = lEntryID
                            ITM.ListSubItems.add(, , !ITemID & "").Tag = ""
                            lTotalQty_PI = lTotalQty_PI + lQtyToUse
MoveToNext:
                            .MoveNEXT
                        Loop
                        .Close
                    End With
                    .MoveNEXT
                Loop
                .Close
            End With
            .MoveNEXT
        Loop
        .Close
        
        lblTotalPcsPI.Caption = lTotalQty_PI
        
        .Open "SELECT * FROM VDispatchListToExport WHERE DispatchListNo=" & lEntryID & " AND DispatchQty>InvQty", con, adOpenForwardOnly, adLockReadOnly
        LVNotMatched.ListItems.Clear
        Do Until .EOF
            With rsCartons
                .Open "SELECT CartonNo,OrderNo,ItemCode,SUM(Qty) AS CartonQty FROM VDispatchListDetail_Ex_OrderWise " & _
                    "WHERE DispatchListEntryID=" & lEntryID & " AND OrderNo='" & rs!OrderNo & "' AND ItemCode='" & rs!ItemCode & "' GROUP BY CartonNo,OrderNo,ItemCode ORDER BY CartonNo", con, adOpenForwardOnly, adLockReadOnly
                    
                
                Do Until .EOF
                             
                    Set ITM = LVNotMatched.ListItems.add(, , Val(rsCartons!CartonNo & ""))
                            
                    ITM.ListSubItems.add(, , !OrderNo & "").Tag = ""
                    ITM.ListSubItems.add(, , !ItemCode & "").Tag = ""
                    ITM.ListSubItems.add , , Val(!CartonQty & "")
                    ITM.ListSubItems.add(, , Val(!CartonQty & "")).Tag = lEntryID
                    
                    
                    .MoveNEXT
                Loop
                    .Close
            End With
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Set rsCartons = Nothing
    Set rsProformas = Nothing
    
End Sub

Private Function getTempUsedQty(lEntryID As Long) As Long

    Dim i As Integer, lQty As Long
    For i = 1 To LV2.ListItems.count
        With LV2.ListItems(i)
            If Val(.ListSubItems(1).Tag) = lEntryID Then
                lQty = lQty + Val(.SubItems(6))
            End If
        End With
    Next
    getTempUsedQty = lQty
    
End Function

Private Sub cmdCancelDispatch_Click()

    FraDispatch.Visible = False
    FAll.Enabled = True
    
End Sub

Private Sub cmdOKDispatch_Click()

    txtCartonNo.Text = cmbCartonNo

    txtDispatchListNo = cmbDispatchNo.Text
    txtDispatchListNo.Tag = cmbDispatchNo.ID
    
    txtDispatchListCartonNo = cmbCartonNo
    
    txtDispatchListQty = Val(txtItemCartonQty)
    
    Call UpdateDispatchListQtyUsed
    
    
    FAll.Enabled = True
    FraDispatch.Visible = False
     
    With txtCartonNo
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
End Sub

Private Sub cmdRemoveAll_Click()

    If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Dim i As Integer
    For i = LV2.ListItems.count To 1 Step -1
        With LV2.ListItems(i)
            If Val(.Tag) > 0 Then
                strDelListPL = strDelListPL & Val(.Tag) & ","
            End If
            LV2.ListItems.Remove .Index
        End With
    Next
    
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
    
    Exit Sub
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
    
    If strCompany = "Dr-Frgz" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
        cmdCallDispatch.Visible = True
        cmdAdd.Visible = False
        SSTab2.TabVisible(1) = True
    Else
        cmdCallDispatch.Visible = False
        cmdAdd.Visible = True
        SSTab2.TabVisible(1) = False
    End If
    
    bCourierInvoice = False
    strDelList = ""
    strProforma = ""
    'SSTab1.TabVisible(4) = False
    SSTab1.Tab = 0
    SSTab2.Tab = 0
    
    Call SaveLV(LV)
    
    Call cmbCust.AddVals(con, "DISTINCT CustCode", "ForeignCustomers", "CustCode", " WHERE Active=1")
    OrderDT = Date
    
    
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
    
    DTCustom = Date
    DTCustom = Null
    
    Me.KeyPreview = True
    cmbCust.ListHeight = 2500

    cmbShipAgent.ListHeight = 2000
    cmbInsurAgent.ListHeight = 2000
    
    Call FillCmbs
        
    If strCompany = "IAA" Or strCompany = "Kami" Then
        txtDTREDesc = ""
    Else
        txtDTREDesc = "Hereby it is certified that 49332 Kgs Stainless Steel received vide DTRE approval number " & _
     "RTOS/2589/29102008 & Bill of Entry Number 678 Dated 22.10.2008 under SRO No. 563(I)/2005 dated 06.06.2008 " & _
     "& against postdated cheque no 3179423 dated 28.10.2008." & vbNewLine & vbNewLine & _
     "Quantity Imported                                                     =   49332   Kgs" & vbNewLine & _
     "Quantity Already Re-Exported                      =   268056  Pcs     =   9784    Kgs" & vbNewLine & _
     "Quantity Re-Exported In this Invoice              =   5633    Pcs     =   206     Kgs" & vbNewLine & _
     "Total Quantity Re-Exported Including this Invoice =   273689  Pcs     =   9990    Kgs" & vbNewLine & _
     "Quantity Balance                                  =   1079811 Pcs     =   39413   Kgs"
    End If
    
     
         
    cmbProformas.ListHeight = 4000
    cmbItemSearch.ListHeight = 1000
    
    If strCompany = "IAA" Or strCompany = "Kami" Then
        chkCIFTerms.Visible = True
    Else
        chkCIFTerms.Visible = False
    End If
        
    cmbItems.ListHeight = 3400
    cmbItems.DropDownOffLine = True
    
    'If Not (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Then
    
    'End If
    
    strDelListPL = ""
    
'    If strCompany = "Dr-Frgz" And strSubCompany = "Intl" Then
'        cmbItems.Enabled = False
'    Else
'        cmbItems.Enabled = True
'    End If
     
    bItemSelectionFromDispatch = False
    Exit Sub
err:
    MsgBox err.Description
   
End Sub

Public Sub AddNewInvoice(Optional p_bCustomInvoice As Boolean = True, Optional bCourierInvoice As Boolean = False, Optional p_bSampleInvoice As Boolean = False)

    chkSampleInvoice = Abs(p_bSampleInvoice)
    If p_bSampleInvoice Then
        chkSampleInvoice.Enabled = False
    End If
    
    bCustomInvoice = p_bCustomInvoice
    Call OrderDT_Change
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
    
    FraDispatch.Left = (Me.ScaleWidth - FraDispatch.Width) / 2
    FraDispatch.Top = (Me.ScaleHeight - FraDispatch.Height) / 2
    
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
'        With txtQty
'            .Move LV.ColumnHeaders(10).Left + LV.Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(10).Width
'            .Text = Val(LV.SelectedItem.SubItems(9))
'            .Visible = True
'            .SelStart = 0
'            .SelLength = Len(.Text)
'            .SetFocus
'        End With
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
    lblTotWords = ConvertInWords(CCur(dTotal), "")
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
    End If
    
    If LV2.ListItems.count = 0 Then
        MsgBox "No Items Added!!!", vbInformation
        Exit Function
    End If
    
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
'    If strProforma <> "" Then
'
'        con.Execute "UPDATE FProformaOrders SET InvQty=InvQty+IsNull((Select  " & _
'         "Qty From CustomInvoiceItems  Where CustomInvoice='" & strProforma & _
'         "' And RefID=FProformaOrders.EntryID),0) WHERE EntryID In(Select RefID" & _
'         " From CustomInvoiceItems Where CustomInvoice='" & strProforma & "')"
'
'    End If
        
    If strProforma = "" Then
        Dim ComDecBalls As String, ComDecGarments As String, ComDecGloves As String
        Dim ComNotify As String, ComConsignee As String
        Dim ComPayTerms As String
        
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
    
    Dim strDTCustom As String
    If IsNull(DTCustom) Then
        strDTCustom = "NULL"
    Else
        strDTCustom = "'" & DTCustom & "'"
    End If
    
    If strProforma = "" Then

        con.Execute "INSERT INTO CustomInvoice (CustomInvoice,DT,CustCode,Country,ShippingMode," & _
         "LoadPortID,DisPortID,TradeTerms,PaymentTerms,LCNo,PartialShipment," & _
         "TransShipment,ShippingAgentID,InsuranceAgentID,OurBankID," & _
         "AWBNo,AWBNoDT,FormE,FormEDT,Vessel,GrossWeight,DeclarationsBalls,DeclarationsGarments," & _
         "DeclarationsGloves,Consignee,ComDeclarationsBalls,ComDeclarationsGarments," & _
         "ComDeclarationsGloves,ComNotify,ComConsignee,ComPayTerms,ComAgainst,ComDrawnUnder," & _
         "ComSpecial,ManDischarge,ManDeclaration,Warehouse,Gate,ContainerNo,DTETA,SRONo,CustCurrency,CompanyRefID,SealNo,InvoiceOf,LCDate,GatePassNo,GatePassDT,DTREDescription,NoSRO,CIFTradeTerms,SampleInvoice,FreeSamples,InvoiceDescription,CustomDT,GDNo,InvoiceTo) VALUES('" & txtPInvoice & _
         "','" & OrderDT.Value & "','" & cmbCust.Text & "','" & cmbCountry.Text & "','" & _
         cmbSMethod.Text & "'," & iLoadID & "," & iDisID & ",'" & cmbTradeTerms.Text & _
         "','" & cmbPayment & "','" & txtLC & "','" & cmbPartial & "','" & _
         cmbTrans & "'," & Val(cmbShipAgent.ID) & "," & Val(cmbInsurAgent.ID) & "," & _
         iBankID1 & ",'" & txtAWBNo & "','" & AWBDT & "','" & txtFormE & "','" & FormEDT & _
         "','" & txtVessel & "'," & Val(txtGWeight) & ",'" & _
         txtDeclarationBalls & "','" & txtDeclarationGarments & "','" & txtDeclarationGloves & _
         "','" & txtBankAddress & "','" & ComDecBalls & "','" & _
         ComDecGarments & "','" & ComDecGloves & "','" & ComNotify & _
         "','" & ComConsignee & "','" & ComPayTerms & "','" & txtShippingAddress & "','" & _
         txtComDrawn & "','" & txtComSpecial & "','" & txtManDischarge & "','" & txtManDeclaration & "','" & _
         txtWarehouse & "','" & txtGate & "','" & txtContainerNo & "','" & DTETA & "','" & strSRONo & "','" & _
         txtCurrency & "'," & cmbCompany.ID & ",'" & txtSealNo & "','" & txtInvoiceof & "','" & DTLC & "','" & _
         txtGatePassNo & "'," & strDTGatePass & ",'" & txtDTREDesc & "'," & chkNoSRO.Value & "," & chkCIFTerms.Value & "," & chkSampleInvoice.Value & "," & chkFreeSamples.Value & ",'" & txtInvoiceDescription & "'," & strDTCustom & ",'" & txtGDNo & "','" & txtInvoiceTo & "')"
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
         "',ComAgainst='" & txtShippingAddress & "',ComDrawnUnder='" & txtComDrawn & "',ComSpecial='" & _
         txtComSpecial & "',ManDischarge='" & txtManDischarge & "',ManDeclaration='" & _
         txtManDeclaration & "',Warehouse='" & txtWarehouse & "',Gate='" & txtGate & "',ContainerNo='" & _
         txtContainerNo & "',DTETA='" & DTETA & "',SealNo='" & txtSealNo & "',Invoiceof='" & txtInvoiceof & _
         "',LCDate='" & DTLC & "',GatePassNo='" & txtGatePassNo & "',GatePassDT=" & strDTGatePass & ",DTREDescription='" & txtDTREDesc & "',NoSRO=" & chkNoSRO.Value & ",CIFTradeTerms=" & chkCIFTerms.Value & ",SampleInvoice=" & chkSampleInvoice.Value & ",FreeSamples=" & chkFreeSamples.Value & ",InvoiceDescription='" & txtInvoiceDescription & "',CustomDT=" & strDTCustom & ",GDNo='" & txtGDNo & "',InvoiceTo='" & txtInvoiceTo & "' WHERE CustomInvoice='" & strProforma & "'"
         
'        If strDelList <> "" Then
'            strDelList = Left(strDelList, Len(strDelList) - 1)
'             con.Execute "DELETE FROM CustomInvoiceItems WHERE EntryID IN(" & strDelList & ")"
'        End If
        
    End If
   
    Dim lEntryIDs() As Double
    Dim i As Integer, iCount As Integer, j As Integer, bFound As Boolean
    
    ReDim lEntryIDs(LV2.ListItems.count, 4)
    iCount = -1
    bFound = False
    Dim iIndex As Integer, lQty As Long, dPrice As Double, dCustomPrice As Double
    
    For i = 1 To LV2.ListItems.count
        With LV2.ListItems(i)
            If Val(.Tag) = 0 Then
                bFound = False
                For j = 0 To iCount
                    If lEntryIDs(j, 0) = Val(.ListSubItems(1).Tag) Then
                        bFound = True
                        Exit For
                    End If
                Next
                If bFound Then
                    iIndex = j
                    lQty = lEntryIDs(ilindex, 1) + Val(.SubItems(6))
                    dPrice = Val(.ListSubItems(2).Tag)
                    dCustomPrice = Val(.ListSubItems(3).Tag)
                Else
                    iIndex = iCount + 1
                    lQty = Val(.SubItems(6))
                    dPrice = Val(.ListSubItems(2).Tag)
                    dCustomPrice = Val(.ListSubItems(3).Tag)
                    iCount = iCount + 1
                End If
                If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                    dCustomPrice = dPrice
                End If
                lEntryIDs(iIndex, 0) = Val(.ListSubItems(1).Tag)
                lEntryIDs(iIndex, 1) = lQty
                lEntryIDs(iIndex, 2) = dPrice
                lEntryIDs(iIndex, 3) = dCustomPrice
            End If
        End With
    Next
    
     
    If strProforma <> "" Then
        Dim lQtyDiff As Long
        For i = 1 To LV2.ListItems.count
            With LV2.ListItems(i)
                If Val(.Tag) > 0 Then
                    If Val(.SubItems(2)) <> Val(.ListSubItems(2).Tag) Then
                        lQtyDiff = Val(.SubItems(2)) - Val(.ListSubItems(2).Tag)
                        con.Execute "UPDATE CustomInvoiceItems SET CustomInvoiceItems.Qty=CustomInvoiceItems.Qty+" & _
                         lQtyDiff & " FROM CustomInvoiceitems INNER JOIN CustomPList ON CustomInvoiceItems.EntryID" & _
                         "=CustomPlist.OrderItemID WHERE CustomPList.ID=" & Val(.Tag)
                         
                        con.Execute "UPDATE CustomPList SET Qty=" & Val(.SubItems(2)) & " WHERE ID=" & Val(.Tag)
                    End If
                End If
            End With
        Next
    End If
    
    If strDelListPL <> "" Then
    
        strDelListPL = Left(strDelListPL, Len(strDelListPL) - 1)
        
        con.Execute "UPDATE CustomInvoiceItems SET CustomInvoiceItems.Qty=CustomInvoiceItems.Qty-T1.Qty FROM " & _
         "CustomInvoiceItems INNER JOIN (SELECT OrderItemID,SUM((CartonTo-CartonFrom+1)*Qty) AS Qty FROM CustomPList WHERE ID IN(" & strDelListPL & ") GROUP BY OrderItemID) T1 ON CustomInvoiceItems.EntryID=T1.OrderItemID"
         
        con.Execute "DELETE FROM CustomPList WHERE ID IN(" & strDelListPL & ")"
        
        con.Execute "DELETE FROM CustomInvoiceItems WHERE Qty=0"
    End If
     
    
    
    Dim lEntryID As Long, iRowAffect As Integer
    Dim lCII_EntryID As Long, lCustomPList_RefID As Long
    
    For i = 0 To iCount
    
        lEntryID = lEntryIDs(i, 0)
        dQty = lEntryIDs(i, 1)
        dPrice = lEntryIDs(i, 2)
        dCustomPrice = lEntryIDs(i, 3)
        
        con.Execute "UPDATE CustomInvoiceItems SET  Qty=" & dQty & ",Price=" & dPrice & ",CustomPrice=" & _
         dCustomPrice & " WHERE RefID=" & lEntryID & " AND CustomInvoice='" & txtPInvoice.Text & "'", iRowAffect
        
        If iRowAffect = 0 Then
            con.Execute "INSERT INTO CustomInvoiceItems(CustomInvoice,RefID,Qty,Price,CustomPrice) VALUES('" & _
             txtPInvoice.Text & "'," & lEntryID & "," & dQty & "," & dPrice & "," & dCustomPrice & ")"
        End If
               
        
        lCII_EntryID = GetSingleLongValue("EntryID", "CustomInvoiceItems", " WHERE CustomInvoice='" & txtPInvoice & "' AND RefID=" & lEntryID)
             
        For j = 1 To LV2.ListItems.count
            With LV2.ListItems(j)
                If lEntryID = Val(.ListSubItems(1).Tag) And Val(.Tag) = 0 Then
                    con.Execute "INSERT INTO CustomPlist(CustomInvoice,OrderItemID,CartonFrom,CartonTo,Qty,LotNo) VALUES('" & _
                     txtPInvoice.Text & "'," & lCII_EntryID & "," & Val(.Text) & "," & Val(.SubItems(1)) & "," & _
                     Val(.SubItems(2)) & ",'" & .SubItems(7) & "')"
                
                End If
                If Val(.ListSubItems(6).Tag) > 0 And Val(.Tag) = 0 And lEntryID = Val(.ListSubItems(1).Tag) Then
                    CustomPList_DispatchListDetail = GetSingleLongValue("MAX(ID)", "CustomPList")
                    con.Execute "INSERT INTO CustomPList_DispatchListDetail(CustomPlist_RefID,ItemCode,CartonNo,DP_RefID,FromOrderNo) VALUES(" & CustomPList_DispatchListDetail & _
                     ",'" & .SubItems(7) & "'," & Val(.SubItems(6)) & "," & Val(.ListSubItems(6).Tag) & ",'" & .SubItems(8) & "')"
                End If
            End With
        Next
        
    Next
    
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            'If CBool(.ListSubItems(11).Tag) Then
            dPrice = Val(.SubItems(8))
            dCustomPrice = Val(.ListSubItems(7).Tag)
            dCustomPrice = dPrice

            con.Execute "UPDATE CustomInvoiceItems SET ItemForLabTest=" & Abs(CBool(.ListSubItems(11).Tag)) & ",Price=" & dPrice & ",CustomPrice=" & dCustomPrice & " WHERE EntryID=" & LV.ListItems(i).ListSubItems(1).Tag
            'End If
        End With
    Next
    
    For i = 1 To LV2.ListItems.count
        With LV2.ListItems(i)
            If Val(.Tag) > 0 Then
                con.Execute "UPDATE CustomPList SET CartonFrom=" & Val(.Text) & ",CartonTo=" & Val(.SubItems(1)) & ",LotNo='" & .SubItems(7) & "' WHERE ID=" & Val(.Tag)
            End If
        End With
    Next
'    For i = 1 To LV.ListItems.count
'
'        With LV.ListItems(i)
'            dPrice = Val(.SubItems(8))
'            dCustomPrice = Val(.ListSubItems(7).Tag)
'            If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" OR strCompany = "Kami" Then
'                dCustomPrice = dPrice
'            End If
'            If Val(.ListSubItems(1).Tag) = 0 Then
'                con.Execute "INSERT INTO CustomInvoiceItems(CustomInvoice,RefID,Qty,Price," & _
'                "CustomPrice,RebateRate,DTRENo,ItemForLabTest,BatchNo) VALUES('" & txtPInvoice & "'," & Val(.key) & "," & _
'                 Val(.SubItems(9)) & "," & dPrice & "," & _
'                 dCustomPrice & "," & Val(.ListSubItems(9).Tag) & ",'" & .SubItems(11) & "'," & Abs(CBool(.ListSubItems(11).Tag)) & ",'" & .SubItems(12) & "')"
'            Else
'                If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" OR strCompany = "Kami" Then
'                    con.Execute "UPDATE CustomInvoiceItems SET Qty=" & Val(.SubItems(9)) & _
'                         ",CustomPrice=" & dPrice & ",Price=" & dPrice & ",DTRENo='" & .SubItems(11) & "',ItemForLabTest=" & Abs(CBool(.ListSubItems(11).Tag)) & ",BatchNo='" & .SubItems(12) & "' WHERE EntryID=" & Val(.ListSubItems(1).Tag)
'                Else
'                    If bCustomInvoice Then
'                        con.Execute "UPDATE CustomInvoiceItems SET Qty=" & Val(.SubItems(9)) & _
'                         ",CustomPrice=" & dPrice & ",DTRENo='" & .SubItems(11) & "',ItemForLabTest=" & Abs(CBool(.ListSubItems(11).Tag)) & ",BatchNo='" & .SubItems(12) & "' WHERE EntryID=" & Val(.ListSubItems(1).Tag)
'                    Else
'                        con.Execute "UPDATE CustomInvoiceItems SET Qty=" & Val(.SubItems(9)) & _
'                         ",Price=" & dPrice & ",DTRENo='" & .SubItems(11) & "',ItemForLabTest=" & Abs(CBool(.ListSubItems(11).Tag)) & ",BatchNo='" & .SubItems(12) & "' WHERE EntryID=" & Val(.ListSubItems(1).Tag)
'                    End If
'                End If
'            End If
'
'            con.Execute "UPDATE FProformaOrders SET InvQty=InvQty-" & _
'             Val(.SubItems(9)) & " WHERE EntryID=" & Val(.key)
'
'         End With
'
'    Next i
        
    'Packing List Info to be saved here.....
'    For i = 1 To LV2.ListItems.count
'        With LV2.ListItems(i)
'            If Val(.Tag) = 0 Then
'                con.Execute "INSERT INTO CustomPList(CustomInvoice,OrderItemID,CartonFrom,CartonTo,Qty,LotNo) " & _
'                 "VALUES('" & txtPInvoice & "'," & Val(.ListSubItems(1).Tag) & "," & Val(.Text) & "," & Val(.Text) & _
'                 "," & Val(.SubItems(4)) & ",'" & .SubItems(5) & "')"
'            End If
'        End With
'    Next
    
    Call SaveOtherCharges
    
    'Use the Following Queries Until the exact Problem is found & Resolved...
    con.Execute "UPDATE CustomInvoiceItems SET CustomInvoiceItems.Qty=T1.Qty FROM " & _
     "CustomInvoiceItems INNER JOIN (SELECT OrderItemID,SUM(((CartonTo-CartonFrom)+1)*Qty) AS Qty FROM CustomPList GROUP BY OrderItemID) T1" & _
     " ON CustomInvoiceItems.EntryID=T1.OrderItemID WHERE CustomInvoiceItems.CustomInvoice='" & txtPInvoice & "'"
     
    con.Execute "DELETE FROM CustomInvoiceItems WHERE Qty<0 AND CustomInvoice='" & txtPInvoice & "'"
     
    'Advance Payments
    If strProforma <> "" Then
        con.Execute "DELETE FROM FCustAdvancePaymentsAdjusted WHERE CustomInvoice='" & strProforma & "'"
    End If
    
    For i = 1 To LVAdvance.ListItems.count
        With LVAdvance.ListItems(i)
            If Val(.SubItems(4)) > 0 Then
                con.Execute "INSERT INTO FCustAdvancePaymentsAdjusted(CustomInvoice,RefID,Amount)" & _
                 " Values('" & IIf(strProforma = "", txtPInvoice.Text, strProforma) & "'," & _
                 Val(.key) & "," & Val(.SubItems(4)) & ")"
            End If
        End With
    Next
    con.CommitTrans
    
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
    
    Call AddToCombo(cmbSMethod, "Description", "ShippingMethods", " ORDER BY EntryID")
'    With cmbSMethod
'        .AddItem "By Air"
'        .AddItem "By Sea"
'        .AddItem "By Courier"
'    End With
    
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
    'cmbCompany.ID = 5 'iIndex
    
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
'    If cmbThrough.MatchFound = False Then Exit Sub
'    txtBankAddress = cmbThrough.List(cmbThrough.ListIndex, 2)
End Sub
 
Private Sub GetOrders()

    On Error GoTo err
    
    Dim strTTerms As String
    
    With cmbCountry
        If Not .MatchFound Then Exit Sub
    End With
    
    
    
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

Private Sub LV2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    Dim myLV As ListView
    Set myLV = LV2
    
    If myLV.SortKey + 1 = ColumnHeader.Index Then
        If myLV.SortOrder = lvwAscending Then
            myLV.SortOrder = lvwDescending

        Else
            myLV.SortOrder = lvwAscending

        End If
    Else
        myLV.SortKey = ColumnHeader.Index - 1
        myLV.SortOrder = lvwAscending

    End If
    If myLV.Sorted = False Then myLV.Sorted = True

End Sub

Private Sub LV2_KeyDown(KeyCode As Integer, Shift As Integer)

    If LV2.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbYes Then
            If Val(LV2.SelectedItem.Tag) > 0 Then
                strDelListPL = strDelListPL & Val(LV2.SelectedItem.Tag) & ","
            End If
            LV2.ListItems.Remove LV2.SelectedItem.Index
        End If
    End If
    
End Sub

Private Sub LV2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim i As Integer
    For i = 1 To LV2.ColumnHeaders.count
        With LV2.ColumnHeaders(i)
            If x > Val(.Left) And x < (Val(.Left) + Val(.Width)) Then
                iColNoLV2 = i
                Exit For
            End If
        End With
    Next
    
End Sub

Private Sub LVAdvance_DblClick()

    If LVAdvance.ListItems.count = 0 Then Exit Sub
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(*)", "CustomInvoiceAuthorization", " WHERE CustomInvoice='" & txtpinovice & "'")
    If lCount > 0 Then
        MsgBox "Can't Change Authorization has been done.", vbInformation
        Exit Sub
    End If
    
    With txtAdvAmt
        .Move LVAdvance.Left + LVAdvance.ColumnHeaders(5).Left, LVAdvance.Top + LVAdvance.SelectedItem.Top, LVAdvance.ColumnHeaders(5).Width
        .Text = Val(LVAdvance.SelectedItem.SubItems(4))
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
    
End Sub

Private Sub LVDL_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer, dTotal As Double
    For i = 1 To LVDL.ListItems.count
        If LVDL.ListItems(i).Checked Then
            dTotal = dTotal + Val(LVDL.ListItems(i).SubItems(3))
        End If
    Next
    txtDispQtyTotal = dTotal
End Sub

Private Sub LVNotMatched_DblClick()

    On Error GoTo err
    bItemSelectionFromDispatch = True
         
    cmbItems.DropDownOffLine = False
    cmbItems.ClearVals
    cmbItems.AddVals con, "'{'+ItemID+'} '+ItemName", "Items", "ItemID", " WHERE ItemID='" & LVNotMatched.SelectedItem.SubItems(2) & "'"
    cmbItems.ID = LVNotMatched.SelectedItem.SubItems(2)
    
    chkShowAllItems.Value = vbUnchecked
    chkShowAllItems.Visible = True
    
    bItemSelectionFromDispatch = False
        
    txtCartonNo.Text = LVNotMatched.SelectedItem.Text

    txtDispatchListNo.Tag = Val(LVNotMatched.SelectedItem.ListSubItems(4).Tag)
    
    txtDispatchListCartonNo = txtCartonNo.Text
    
    txtDispatchListQty = Val(LVNotMatched.SelectedItem.SubItems(4))
    txtOrderQty.Tag = LVNotMatched.SelectedItem.SubItems(1)
    
    Call UpdateDispatchListQtyUsed(txtOrderQty.Tag, True)
    
    FAll.Enabled = True
    FraDispatch.Visible = False
     
    With txtCartonNo
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
    Exit Sub
err:
    MsgBox err.Description
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
    
    If strProforma <> "" Then Exit Sub
    
    txtPInvoice = GetInvoiceNo(Val(cmbCompany.ID), False, OrderDT)
    
    
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

Private Sub txtAdvAmt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Round(Val(txtAdvAmt), 2) > Round(Val(LVAdvance.SelectedItem.SubItems(2)) - (Val(LVAdvance.SelectedItem.SubItems(3)) - Val(LVAdvance.SelectedItem.SubItems(4))), 2) Then
            MsgBox "Invalid Amount.", vbInformation
            Exit Sub
        ElseIf Val(txtAdvAmt) <= 0 Then
            MsgBox "Invalid Amount.", vbInformation
            Exit Sub
        End If
        LVAdvance.SelectedItem.SubItems(4) = Val(txtAdvAmt)
        LVAdvance.SelectedItem.SubItems(5) = Round(Val(LVAdvance.SelectedItem.SubItems(2)) - Val(LVAdvance.SelectedItem.SubItems(3)) - Val(LVAdvance.SelectedItem.SubItems(4)), 2)
        txtAdvAmt.Visible = False
    ElseIf KeyAscii = 27 Then
        txtAdvAmt.Visible = False
    End If
End Sub

Private Sub txtAdvAmt_LostFocus()
    txtAdvAmt.Visible = False
End Sub

Private Sub txtCartonQty_Change()
    Call GetTotalQty
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

Private Sub txtEditLV2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If iColNoLV2 = 1 Then
            LV2.SelectedItem.Text = Val(txtEditLV2.Text)
        Else
            LV2.SelectedItem.SubItems(iColNoLV2 - 1) = txtEditLV2.Text
        End If
        txtEditLV2.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEditLV2.Visible = False
    End If
End Sub

Private Sub txtEditLV2_LostFocus()
    txtEditLV2.Visible = False
End Sub

Private Sub txtEditPLQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(txtEditPLQty) <= 0 Then
            MsgBox "Invalid Qty.", vbInformation
            Exit Sub
        End If
        LV2.SelectedItem.SubItems(4) = Val(txtEditPLQty)
    ElseIf KeyAscii = 27 Then
        txtEditPLQty.Visible = False
    End If
End Sub

Private Sub txtEditPLQty_LostFocus()
    txtEditPLQty.Visible = False
End Sub

Private Sub UpdateDispatchListQtyUsed(Optional strOrderNo As String, Optional bImportErrDblClick As Boolean = False)

    Dim lQtyUsed As Long
    lQtyUsed = 0
    
    If Val(txtDispatchListNo.Tag) > 0 Then
        For i = 1 To LV2.ListItems.count
            With LV2.ListItems(i)
                If .SubItems(7) = cmbItems.ID And Val(.SubItems(6)) = Val(txtDispatchListCartonNo) And Val(.ListSubItems(6).Tag) = Val(txtDispatchListNo.Tag) Then
                    If strOrderNo <> "" Then
                        If strOrderNo = .SubItems(8) Then
                            lQtyUsed = lQtyUsed + Val(LV2.ListItems(i).SubItems(4))
                        End If
                    Else
                        lQtyUsed = lQtyUsed + Val(LV2.ListItems(i).SubItems(4))
                    End If
                    
                End If
            End With
        Next
        txtDispatchListQtyUsed = lQtyUsed
        'txtDispatchListQtyBalance = Val(txtDispatchListQty) - lQtyUsed
        
        LVNotMatched.SelectedItem.SubItems(4) = Val(LVNotMatched.SelectedItem.SubItems(3)) - lQtyUsed
        txtDispatchListQtyBalance = Val(LVNotMatched.SelectedItem.SubItems(3)) - lQtyUsed
    End If
    
    If Val(txtDispatchListNo.Tag) > 0 And bImportErrDblClick = False Then
        Call cmdSavenStay_Click
        Call LVNotMatched_DblClick
    End If
    
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
        
    On Error GoTo err
    If KeyAscii = 13 Then
'        If Val(txtPrice) <= 0 Then
'            MsgBox "Invalid Price.", vbInformation
'            Exit Sub
'        End If
        LV.SelectedItem.SubItems(8) = Val(txtPrice)
        LV.SelectedItem.SubItems(10) = Val(txtPrice) * LV.SelectedItem.SubItems(9)
        
'        On Error Resume Next
'        LV2.ListItems(LV.SelectedItem.Key).SubItems(8) = Val(LV.SelectedItem.ListSubItems(8).Tag) - Val(txtQty)
'        If err.Number = 35601 Then
'            Resume Next
'        End If
        
        Dim dPrice As Double, dCustomPrice As Double
        
        dPrice = Val(LV.SelectedItem.SubItems(8))
        dCustomPrice = Val(LV.SelectedItem.ListSubItems(7).Tag)
                
        If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Tecno" Or strCompany = "Towne" Then
            dCustomPrice = dPrice
        End If

        con.Execute "UPDATE CustomInvoiceItems SET Price=" & dPrice & ",CustomPrice=" & dCustomPrice & " WHERE EntryID=" & Val(LV.SelectedItem.ListSubItems(1).Tag)
        
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

Private Sub LV2_DblClick()
    
    If LV2.ListItems.count = 0 Then Exit Sub
    If iColNoLV2 <> 8 Then iColNoLV2 = 1
    With txtEditLV2
        .Move LV2.Left + LV2.ColumnHeaders(iColNoLV2).Left, LV2.Top + LV2.SelectedItem.Top, LV2.ColumnHeaders(iColNoLV2).Width
        If iColNoLV2 = 1 Then
            .Text = Val(LV2.SelectedItem.Text)
        Else
            .Text = LV2.SelectedItem.SubItems(iColNoLV2 - 1)
        End If
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
    Exit Sub
    With txtEditPLQty
        .Move LV2.ColumnHeaders(5).Left + LV2.Left, LV2.Top + LV2.SelectedItem.Top
        .Text = Val(LV2.SelectedItem.SubItems(4))
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
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

Private Sub SaveOtherCharges()

    If strProforma <> "" Then
        'con.Execute "Delete From CustomInvoiceOtherCharges Where CustomInvoice='" & strProforma & "' AND Custom=" & Abs(bCustomInvoice)
        con.Execute "Delete From CustomInvoiceOtherCharges Where CustomInvoice='" & strProforma & "'"
    End If
    
    For i = 1 To LVOtherCharges.ListItems.count
        With LVOtherCharges.ListItems(i)
            con.Execute "INSERT INTO CustomInvoiceOtherCharges(CustomInvoice,ChargesDetail,ChargesAmt,SNo,Custom)" & _
             " VALUES('" & IIf(strProforma = "", txtPInvoice.Text, strProforma) & "','" & _
             .SubItems(1) & "'," & Val(.SubItems(2)) & "," & Val(.Index) & "," & Abs(bCustomInvoice) & ")"
             
             If strProforma = "" Then   'iF nEW Invoice THen save them as commercial....
                con.Execute "INSERT INTO CustomInvoiceOtherCharges(CustomInvoice,ChargesDetail,ChargesAmt,SNo,Custom)" & _
                 " VALUES('" & IIf(strProforma = "", txtPInvoice.Text, strProforma) & "','" & _
                 .SubItems(1) & "'," & Val(.SubItems(2)) & "," & Val(.Index) & ",0)"
             
             End If
        End With
    Next
    
End Sub

Private Sub txtQtyToAdd_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        If cmbItems.MatchFound = False Then
            MsgBox "Please Select Item.", vbInformation
            Exit Sub
        End If
        
        If Val(txtQtyToAdd) > Val(txtAvailable) Then
            If Not (strCompany = "Dr-Frgz" And chkSampleInvoice.Enabled = False) Then
                MsgBox "Invalid Qty.", vbInformation
                Exit Sub
            End If
        ElseIf Val(txtQtyToAdd) > Val(txtDispatchListQtyBalance) And Val(txtDispatchListNo.Tag) > 0 Then
            MsgBox "Invalid Qty. in Dispatch List.", vbInformation
            Exit Sub
        End If
                
        If Val(txtCartonNo) <= 0 Then
            MsgBox "Invalid Carton No.", vbInformation
            Exit Sub
        End If
        
        Dim rs As New ADODB.Recordset
        rs.Open "SELECT EntryID,OrderNo,PInvoice,ItemID,ItemName,Qty,InvQty,Price,CustomPrice,TradeTerms FROM VFProformas WHERE EntryID=" & cmbProformas.ID, con, adOpenForwardOnly, adLockReadOnly
            
        Dim ITM As ListItem
        Set ITM = LV2.ListItems.add(, , Val(txtCartonNo))
        ITM.ListSubItems.add(, , Val(txtCartonTo)).Tag = Val(rs!EntryID & "")
        ITM.ListSubItems.add(, , Val(txtCartonQty)).Tag = Val(rs!Price & "")
        
        ITM.ListSubItems.add(, , rs!PInvoice & "").Tag = Val(rs!CustomPrice & "")
        ITM.ListSubItems.add(, , rs!OrderNo & "").Tag = ""
        ITM.ListSubItems.add(, , rs!ITemID & " - " & rs!ItemName).Tag = Val(rs!CustomPrice & "")
        ITM.ListSubItems.add(, , Val(txtQtyToAdd)).Tag = Val(txtDispatchListNo.Tag)
        ITM.ListSubItems.add , , txtLotNo & ""
         
        'Following 3 Lines are for Dispatch...
        ITM.ListSubItems.add(, , Val(txtDispatchListCartonNo)).Tag = Val(txtDispatchListNo.Tag)
        ITM.ListSubItems.add , , cmbItems.ID
        ITM.ListSubItems.add , , txtOrderQty.Tag
        
        If cmbTradeTerms.Text = "" Then
            cmbTradeTerms = rs!TradeTerms & ""
        End If
        
        rs.Close
        Set rs = Nothing
        
        Dim lQtyUsed As Long
        Dim i As Integer
        For i = 1 To LV2.ListItems.count
            If Val(LV2.ListItems(i).ListSubItems(1).Tag) = Val(cmbProformas.ID) Then
                lQtyUsed = lQtyUsed + Val(LV2.ListItems(i).SubItems(4))
            End If
        Next
        txtUsed = lQtyUsed
        txtAvailable = Val(txtBalance) - Val(txtUsed)
        
        Call UpdateDispatchListQtyUsed
        
        txtQtyToAdd = ""
        txtQtyToAdd.SetFocus
    End If
    
End Sub


Private Sub FillAdvancePayments()
    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "SELECT EntryID,AdviceNo,Amount,AdjustedAmount FROM VFCustAdvancePayments Where CustCode='" & cmbCust.ID & "' AND Amount>AdjustedAmount ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LVAdvance.ListItems.Clear
        Do Until .EOF
            Set ITM = LVAdvance.ListItems.add(, !EntryID & "'", LVAdvance.ListItems.count + 1)
            ITM.ListSubItems.add , , !AdviceNo & ""
            ITM.ListSubItems.add(, , Val(!Amount & "")).Tag = Val(!Amount & "")
            ITM.ListSubItems.add(, , Val(!AdjustedAmount & "")).Tag = Val(!AdjustedAmount & "")
            ITM.ListSubItems.add(, , 0).Tag = 0
            ITM.ListSubItems.add , , Round(Val(!Amount & "") - Val(!AdjustedAmount & ""), 2)
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub txtCartonQty_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Call txtQtyToAdd_KeyPress(13)
    End If
    
End Sub

Private Sub GetTotalQty()

    Dim lTotalQty As Long
    lTotalQty = (Val(txtCartonTo) - Val(txtCartonNo) + 1) * Val(txtCartonQty)
    txtQtyToAdd = lTotalQty
    
End Sub

Private Sub txtCartonTo_Change()
    Call GetTotalQty
End Sub

Private Sub cmbInvoiceTos_Change()
    Call cmbInvoiceTos_Click
End Sub

Private Sub cmbInvoiceTos_Click()

    If cmbInvoiceTos.MatchFound = False Then
        txtInvoiceTo = ""
    Else
        txtInvoiceTo = GetSingleStringValue("InvoiceTo", "ForeignCustomers_InvoiceTo", " WHERE EntryID=" & cmbInvoiceTos.List(cmbInvoiceTos.ListIndex, 1))
    End If
    
End Sub


