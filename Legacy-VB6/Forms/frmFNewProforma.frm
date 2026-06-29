VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmFNewProforma 
   ClientHeight    =   10050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12195
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
   ScaleHeight     =   10050
   ScaleWidth      =   12195
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
      Height          =   9990
      Left            =   45
      TabIndex        =   15
      Top             =   30
      Width           =   12120
      Begin ComboList.Usercontrol1 cmbItemSearchProforma 
         Height          =   285
         Left            =   900
         TabIndex        =   93
         Top             =   7890
         Width           =   4455
         _ExtentX        =   7858
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
         Left            =   1575
         TabIndex        =   0
         Top             =   720
         Width           =   3585
         _ExtentX        =   6324
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
      Begin VB.TextBox txtBatchNo 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   10410
         MaxLength       =   150
         TabIndex        =   52
         Top             =   1335
         Width           =   1620
      End
      Begin VB.TextBox txtLC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2190
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   8925
         Width           =   1830
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   90
         TabIndex        =   1
         Top             =   1335
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
      Begin VB.TextBox txtPackaging 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   8235
         TabIndex        =   10
         Top             =   8925
         Width           =   1995
      End
      Begin VB.TextBox txtCurrency 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4590
         MaxLength       =   150
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1335
         Width           =   960
      End
      Begin VB.TextBox txtPInvoice 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5565
         MaxLength       =   150
         TabIndex        =   4
         Top             =   1335
         Width           =   1710
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
         TabIndex        =   18
         Top             =   -15
         Width           =   12120
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "New Proforma"
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
            TabIndex        =   19
            Top             =   165
            Width           =   12015
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "New Proforma"
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
            TabIndex        =   20
            Top             =   180
            Width           =   12000
         End
      End
      Begin MSComCtl2.DTPicker OrderDT 
         Height          =   300
         Left            =   7290
         TabIndex        =   5
         Top             =   1320
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
         Format          =   153747459
         CurrentDate     =   37250
      End
      Begin MSComCtl2.DTPicker DeliveryDT 
         Height          =   285
         Left            =   10245
         TabIndex        =   11
         Top             =   8925
         Width           =   1755
         _ExtentX        =   3096
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
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   153681923
         CurrentDate     =   37250
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6660
         Left            =   75
         TabIndex        =   27
         Top             =   1650
         Width           =   11940
         _ExtentX        =   21061
         _ExtentY        =   11748
         _Version        =   393216
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         TabCaption(0)   =   "Details"
         TabPicture(0)   =   "frmFNewProforma.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "cmbOrders1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label1(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdAddItem1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label1(16)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lblTotPcs"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label3"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "lblTotalAmt"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label1(24)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "lblGTAmt"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "LV"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "cmdDelete"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txtQty"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "txtUpdatePrice"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "cmdUpdatePrices"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "cmdRemoveAll"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "txtDiscountPer"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "txtDiscountAmt"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "txtInsuranceAmt"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "txtInsurance_Description"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "txtDiscount_Description"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "txtInsurance_Description_II"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "txtInsuranceAmt_II"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "txtCIF_Description"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "txtTotal_Description"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).ControlCount=   24
         TabCaption(1)   =   "Other Info."
         TabPicture(1)   =   "frmFNewProforma.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdCalcWeights"
         Tab(1).Control(1)=   "txtPackages"
         Tab(1).Control(2)=   "txtNetWeight"
         Tab(1).Control(3)=   "txtGrossWeight"
         Tab(1).Control(4)=   "txtDelivery"
         Tab(1).Control(5)=   "txtSteelWeight"
         Tab(1).Control(6)=   "Frame5"
         Tab(1).Control(7)=   "Frame3"
         Tab(1).Control(8)=   "Frame4"
         Tab(1).Control(9)=   "Frame2"
         Tab(1).Control(10)=   "txtInvoiceTo"
         Tab(1).Control(11)=   "Label1(23)"
         Tab(1).Control(12)=   "cmbInvoiceTos"
         Tab(1).Control(13)=   "cmbShippingAddresses"
         Tab(1).Control(14)=   "Label1(22)"
         Tab(1).Control(15)=   "txtShippingAddress"
         Tab(1).Control(16)=   "Label1(21)"
         Tab(1).Control(17)=   "Label1(20)"
         Tab(1).Control(18)=   "txtInvoiceof"
         Tab(1).Control(19)=   "Label1(18)"
         Tab(1).Control(20)=   "Label1(17)"
         Tab(1).ControlCount=   21
         TabCaption(2)   =   "New Items"
         TabPicture(2)   =   "frmFNewProforma.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Command1"
         Tab(2).Control(1)=   "txtQty1"
         Tab(2).Control(2)=   "LV2"
         Tab(2).Control(3)=   "cmbItemSearch"
         Tab(2).Control(4)=   "lblGoTo"
         Tab(2).Control(5)=   "cmdAddAll"
         Tab(2).Control(6)=   "cmbOrders"
         Tab(2).Control(7)=   "Label1(14)"
         Tab(2).Control(8)=   "cmdAddItem"
         Tab(2).ControlCount=   9
         TabCaption(3)   =   "Other Charges"
         TabPicture(3)   =   "frmFNewProforma.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "txtOtherChargesSNo"
         Tab(3).Control(1)=   "txtOtherChargeAmt"
         Tab(3).Control(2)=   "cmdAddOtherCharges"
         Tab(3).Control(3)=   "txtOtherCharges"
         Tab(3).Control(4)=   "LVOtherCharges"
         Tab(3).Control(5)=   "lblCap(0)"
         Tab(3).Control(6)=   "lblCap(1)"
         Tab(3).Control(7)=   "lblCap(2)"
         Tab(3).ControlCount=   8
         TabCaption(4)   =   "Notes"
         TabPicture(4)   =   "frmFNewProforma.frx":0070
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Frame7"
         Tab(4).ControlCount=   1
         Begin VB.TextBox txtTotal_Description 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8040
            TabIndex        =   120
            Text            =   "TOTAL"
            Top             =   5175
            Width           =   1650
         End
         Begin VB.TextBox txtCIF_Description 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6135
            TabIndex        =   118
            Text            =   "GRAND TOTAL"
            Top             =   6255
            Width           =   3555
         End
         Begin VB.TextBox txtInsuranceAmt_II 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   9675
            TabIndex        =   117
            Top             =   5985
            Width           =   1770
         End
         Begin VB.TextBox txtInsurance_Description_II 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6135
            TabIndex        =   116
            Text            =   "OTHER"
            Top             =   5985
            Width           =   3555
         End
         Begin VB.TextBox txtDiscount_Description 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8040
            TabIndex        =   115
            Text            =   "DISCOUNT"
            Top             =   5445
            Width           =   1650
         End
         Begin VB.TextBox txtInsurance_Description 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6135
            TabIndex        =   114
            Text            =   "FREIGHT"
            Top             =   5715
            Width           =   3555
         End
         Begin VB.TextBox txtInsuranceAmt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   9675
            TabIndex        =   113
            Top             =   5715
            Width           =   1770
         End
         Begin VB.TextBox txtDiscountAmt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   9675
            Locked          =   -1  'True
            TabIndex        =   112
            Top             =   5445
            Width           =   1770
         End
         Begin VB.TextBox txtDiscountPer 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7080
            TabIndex        =   111
            Top             =   5445
            Width           =   975
         End
         Begin VB.CommandButton cmdRemoveAll 
            Caption         =   "Remove All Items"
            Height          =   360
            Left            =   405
            TabIndex        =   108
            Top             =   5835
            Width           =   1455
         End
         Begin VB.Frame Frame7 
            Height          =   4170
            Left            =   -73290
            TabIndex        =   103
            Top             =   1095
            Width           =   9075
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Notes I"
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
               MouseIcon       =   "frmFNewProforma.frx":008C
               TabIndex        =   107
               Top             =   150
               Width           =   8895
            End
            Begin MSForms.TextBox txtNotes1 
               Height          =   1740
               Left            =   90
               TabIndex        =   106
               Top             =   390
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
               MouseIcon       =   "frmFNewProforma.frx":01DE
               TabIndex        =   105
               Top             =   2115
               Width           =   8895
            End
            Begin MSForms.TextBox txtNotes2 
               Height          =   1740
               Left            =   90
               TabIndex        =   104
               Top             =   2355
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
         End
         Begin VB.CommandButton cmdCalcWeights 
            Caption         =   "Calc."
            Height          =   285
            Left            =   -65235
            TabIndex        =   98
            Top             =   4050
            Width           =   780
         End
         Begin VB.TextBox txtPackages 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   -64455
            MaxLength       =   150
            TabIndex        =   97
            Top             =   4050
            Width           =   1155
         End
         Begin VB.TextBox txtNetWeight 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   -67125
            MaxLength       =   150
            TabIndex        =   96
            Top             =   4050
            Width           =   1890
         End
         Begin VB.TextBox txtGrossWeight 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   -69000
            MaxLength       =   150
            TabIndex        =   95
            Top             =   4050
            Width           =   1890
         End
         Begin VB.TextBox txtDelivery 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   -74880
            TabIndex        =   92
            Top             =   6225
            Width           =   5700
         End
         Begin VB.CommandButton cmdUpdatePrices 
            Caption         =   "Update Prices"
            Height          =   300
            Left            =   2775
            TabIndex        =   90
            Top             =   5445
            Width           =   1335
         End
         Begin VB.TextBox txtUpdatePrice 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1950
            TabIndex        =   88
            Top             =   5460
            Width           =   780
         End
         Begin VB.TextBox txtOtherChargesSNo 
            Height          =   315
            Left            =   -74925
            TabIndex        =   79
            Top             =   915
            Width           =   1200
         End
         Begin VB.TextBox txtOtherChargeAmt 
            Height          =   315
            Left            =   -66615
            TabIndex        =   78
            Top             =   915
            Width           =   2000
         End
         Begin VB.CommandButton cmdAddOtherCharges 
            Caption         =   "Add"
            Height          =   315
            Left            =   -64560
            TabIndex        =   77
            Top             =   915
            Width           =   1380
         End
         Begin VB.TextBox txtOtherCharges 
            Height          =   315
            Left            =   -73695
            TabIndex        =   76
            Top             =   915
            Width           =   7050
         End
         Begin VB.TextBox txtSteelWeight 
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   -69000
            MaxLength       =   150
            TabIndex        =   71
            Top             =   3015
            Width           =   5700
         End
         Begin VB.Frame Frame5 
            Height          =   3300
            Left            =   -74895
            TabIndex        =   62
            Top             =   2685
            Width           =   5850
            Begin ComboList.Usercontrol1 cmbInsurAgent 
               Height          =   285
               Left            =   2955
               TabIndex        =   63
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
               Left            =   45
               TabIndex        =   64
               Top             =   420
               Width           =   2865
               _ExtentX        =   5054
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
               Index           =   12
               Left            =   45
               TabIndex        =   68
               Top             =   180
               Width           =   2865
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
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
               Left            =   2955
               TabIndex        =   67
               Top             =   180
               Width           =   2835
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Declarations"
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
               Left            =   75
               TabIndex        =   66
               Top             =   885
               Width           =   5700
            End
            Begin MSForms.TextBox txtDeclaration 
               Height          =   2085
               Left            =   75
               TabIndex        =   65
               Top             =   1125
               Width           =   5700
               VariousPropertyBits=   -1400879077
               BorderStyle     =   1
               Size            =   "10054;3678"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
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
            Left            =   -63495
            Picture         =   "frmFNewProforma.frx":0330
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Remove This Entry From Voucher"
            Top             =   975
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.TextBox txtQty1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -65730
            TabIndex        =   54
            Top             =   1785
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtQty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4230
            TabIndex        =   53
            Top             =   495
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.Frame Frame3 
            Height          =   2325
            Left            =   -74895
            TabIndex        =   43
            Top             =   345
            Width           =   5850
            Begin VB.TextBox txtBankAddress 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   750
               Left            =   1980
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   44
               Top             =   1470
               Width           =   3765
            End
            Begin MSForms.ComboBox cmbBankAccNo 
               Height          =   285
               Left            =   60
               TabIndex        =   85
               Top             =   960
               Width           =   5685
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "10028;503"
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
               Index           =   19
               Left            =   60
               TabIndex        =   84
               Top             =   720
               Width           =   5685
            End
            Begin MSForms.ComboBox cmbBank 
               Height          =   285
               Left            =   1980
               TabIndex        =   69
               Top             =   450
               Width           =   3765
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "6641;503"
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
               Caption         =   " Shipping Method            Bank"
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
               Left            =   60
               TabIndex        =   49
               Top             =   210
               Width           =   5685
            End
            Begin MSForms.ComboBox cmbSMethod 
               Height          =   285
               Left            =   60
               TabIndex        =   48
               Top             =   450
               Width           =   1905
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "3360;503"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
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
               Index           =   10
               Left            =   60
               TabIndex        =   47
               Top             =   1230
               Width           =   1920
            End
            Begin MSForms.ComboBox cmbThrough 
               Height          =   285
               Left            =   60
               TabIndex        =   46
               Top             =   1470
               Width           =   1935
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "3413;503"
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
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Bank Address"
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
               Left            =   1695
               TabIndex        =   45
               Top             =   1230
               Visible         =   0   'False
               Width           =   4050
            End
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
            Left            =   11490
            Picture         =   "frmFNewProforma.frx":0862
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Remove This Entry From Voucher"
            Top             =   975
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   345
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
            Height          =   1050
            Left            =   -69000
            TabIndex        =   34
            Top             =   1620
            Width           =   5835
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Country                          City                                   Port"
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
               Left            =   105
               TabIndex        =   38
               Top             =   315
               Width           =   5640
            End
            Begin MSForms.ComboBox cmbDisCountry 
               Height          =   285
               Left            =   105
               TabIndex        =   37
               Top             =   555
               Width           =   1860
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "3281;503"
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
               Left            =   1995
               TabIndex        =   36
               Top             =   555
               Width           =   1860
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "3281;503"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbDisPort 
               Height          =   285
               Left            =   3885
               TabIndex        =   35
               Top             =   555
               Width           =   1860
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "3281;503"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
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
            Height          =   1050
            Left            =   -69000
            TabIndex        =   29
            Top             =   345
            Width           =   5835
            Begin MSForms.ComboBox cmbLoadPort 
               Height          =   285
               Left            =   3885
               TabIndex        =   33
               Top             =   555
               Width           =   1860
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "3281;503"
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
               Left            =   1995
               TabIndex        =   32
               Top             =   555
               Width           =   1860
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "3281;503"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbLoadCountry 
               Height          =   285
               Left            =   105
               TabIndex        =   31
               Top             =   555
               Width           =   1860
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "3281;503"
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
               Caption         =   " Country                          City                                   Port"
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
               Left            =   105
               TabIndex        =   30
               Top             =   315
               Width           =   5640
            End
         End
         Begin MSComctlLib.ListView LV 
            Height          =   4800
            Left            =   300
            TabIndex        =   28
            Top             =   390
            Width           =   11145
            _ExtentX        =   19659
            _ExtentY        =   8467
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
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S No."
               Object.Width           =   1014
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Order #"
               Object.Width           =   1720
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Item ID"
               Object.Width           =   1692
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Description"
               Object.Width           =   8652
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Unit"
               Object.Width           =   1147
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Price"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "Qty"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "Amount"
               Object.Width           =   1852
            EndProperty
         End
         Begin MSComctlLib.ListView LV2 
            Height          =   5205
            Left            =   -74685
            TabIndex        =   56
            Top             =   990
            Width           =   11130
            _ExtentX        =   19632
            _ExtentY        =   9181
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
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S No."
               Object.Width           =   1014
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Order #"
               Object.Width           =   1720
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Item ID"
               Object.Width           =   1692
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Description"
               Object.Width           =   8652
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Unit"
               Object.Width           =   1147
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Price"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "Qty"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "Amount"
               Object.Width           =   1852
            EndProperty
         End
         Begin MSComctlLib.ListView LVOtherCharges 
            Height          =   5250
            Left            =   -74925
            TabIndex        =   80
            Top             =   1260
            Width           =   11760
            _ExtentX        =   20743
            _ExtentY        =   9260
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
         Begin MSForms.TextBox txtInvoiceTo 
            Height          =   435
            Left            =   -69000
            TabIndex        =   123
            Top             =   6135
            Width           =   5700
            VariousPropertyBits=   -1400879077
            BorderStyle     =   1
            Size            =   "10054;767"
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
            Index           =   23
            Left            =   -69000
            TabIndex        =   122
            Top             =   5625
            Width           =   5700
         End
         Begin MSForms.ComboBox cmbInvoiceTos 
            Height          =   285
            Left            =   -69000
            TabIndex        =   121
            Top             =   5865
            Width           =   5700
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "10054;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblGTAmt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E7EBEF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   9675
            TabIndex        =   119
            Top             =   6255
            Width           =   1770
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Dis (%) : "
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
            Height          =   285
            Index           =   24
            Left            =   6135
            TabIndex        =   110
            Top             =   5445
            Width           =   960
         End
         Begin VB.Label lblTotalAmt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E7EBEF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   9675
            TabIndex        =   109
            Top             =   5175
            Width           =   1770
         End
         Begin MSForms.ComboBox cmbShippingAddresses 
            Height          =   285
            Left            =   -69000
            TabIndex        =   101
            Top             =   4560
            Width           =   5700
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "10054;503"
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
            Index           =   22
            Left            =   -69000
            TabIndex        =   100
            Top             =   4320
            Width           =   5700
         End
         Begin MSForms.TextBox txtShippingAddress 
            Height          =   810
            Left            =   -69000
            TabIndex        =   99
            Top             =   4830
            Width           =   5700
            VariousPropertyBits=   -1400879077
            BorderStyle     =   1
            Size            =   "10054;1429"
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
            Caption         =   "        Gross Weight                  Net Weight                               Packages"
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
            Left            =   -69000
            TabIndex        =   94
            Top             =   3810
            Width           =   5700
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Delivery"
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
            Left            =   -74880
            TabIndex        =   91
            Top             =   5985
            Width           =   5700
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Price Update (%) :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   330
            TabIndex        =   89
            Top             =   5475
            Width           =   1560
         End
         Begin MSForms.ComboBox cmbItemSearch 
            Height          =   285
            Left            =   -73455
            TabIndex        =   87
            Top             =   6240
            Width           =   4575
            VariousPropertyBits=   746604571
            BackColor       =   16777215
            ForeColor       =   6244673
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "8070;503"
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
            Left            =   -74655
            TabIndex        =   86
            Top             =   6255
            Width           =   1185
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            Caption         =   "S/No. :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   -74880
            TabIndex        =   83
            Top             =   690
            Width           =   510
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            Caption         =   "Charges Description :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   -73695
            TabIndex        =   82
            Top             =   690
            Width           =   1545
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            Caption         =   "Amount"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   -66600
            TabIndex        =   81
            Top             =   675
            Width           =   555
         End
         Begin MSForms.TextBox txtInvoiceof 
            Height          =   300
            Left            =   -69000
            TabIndex        =   75
            Top             =   3525
            Width           =   5700
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "10054;529"
            Value           =   "Surgical Instruments"
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
            Index           =   18
            Left            =   -69000
            TabIndex        =   74
            Top             =   3285
            Width           =   5700
         End
         Begin MSForms.CommandButton cmdAddAll 
            Height          =   345
            Left            =   -65100
            TabIndex        =   73
            Top             =   600
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
            Caption         =   " Steel Weight"
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
            Left            =   -69000
            TabIndex        =   70
            Top             =   2775
            Width           =   5700
         End
         Begin VB.Label lblTotPcs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E7EBEF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   7080
            TabIndex        =   61
            Top             =   5175
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total Pcs : "
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
            Height          =   285
            Index           =   16
            Left            =   6135
            TabIndex        =   60
            Top             =   5175
            Width           =   960
         End
         Begin MSForms.ComboBox cmbOrders 
            Height          =   285
            Left            =   -74685
            TabIndex        =   59
            Top             =   675
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
            Index           =   14
            Left            =   -74685
            TabIndex        =   58
            Top             =   450
            Width           =   1755
         End
         Begin MSForms.CommandButton cmdAddItem 
            Height          =   348
            Left            =   -72840
            TabIndex        =   57
            Top             =   612
            Width           =   1512
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
            Left            =   2145
            TabIndex        =   42
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
            Left            =   360
            TabIndex        =   41
            Top             =   2370
            Visible         =   0   'False
            Width           =   1755
         End
         Begin MSForms.ComboBox cmbOrders1 
            Height          =   285
            Left            =   360
            TabIndex        =   40
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
         Left            =   5115
         TabIndex        =   16
         Top             =   9225
         Width           =   6900
         Begin MSForms.CommandButton cmdSavenStay 
            Height          =   360
            Left            =   3510
            TabIndex        =   102
            Top             =   195
            Width           =   1590
            Caption         =   "Save & Stay"
            PicturePosition =   327683
            Size            =   "2805;635"
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
            Left            =   5130
            TabIndex        =   14
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
            TabIndex        =   12
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
            TabIndex        =   13
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
      Begin VB.Label Label2 
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
         Left            =   120
         TabIndex        =   72
         Top             =   735
         Width           =   1410
      End
      Begin MSForms.ComboBox cmbPartial 
         Height          =   285
         Left            =   6135
         TabIndex        =   9
         Top             =   8925
         Width           =   2070
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3651;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox lblTotWords 
         Height          =   270
         Left            =   1890
         TabIndex        =   50
         Top             =   8325
         Width           =   7455
         VariousPropertyBits=   746604575
         BackColor       =   15199215
         BorderStyle     =   1
         Size            =   "13150;476"
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
         Left            =   9330
         TabIndex        =   26
         Top             =   8325
         Width           =   885
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
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
         TabIndex        =   25
         Top             =   8325
         Width           =   1830
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   10200
         TabIndex        =   24
         Top             =   8325
         Width           =   1800
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmFNewProforma.frx":0D94
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
         Left            =   90
         TabIndex        =   23
         Top             =   8685
         Width           =   11910
      End
      Begin MSForms.CheckBox chkPrint 
         Height          =   345
         Left            =   90
         TabIndex        =   22
         Top             =   9285
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
      Begin MSForms.ComboBox cmbPayment 
         Height          =   285
         Left            =   90
         TabIndex        =   7
         Top             =   8925
         Width           =   2070
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3651;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbTrans 
         Height          =   285
         Left            =   4035
         TabIndex        =   8
         Top             =   8925
         Width           =   2070
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3651;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkCust 
         Height          =   360
         Left            =   90
         TabIndex        =   21
         Top             =   9540
         Visible         =   0   'False
         Width           =   2685
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   4
         Size            =   "4736;635"
         Value           =   "0"
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
         TabIndex        =   6
         Top             =   1335
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
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmFNewProforma.frx":0E41
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
         TabIndex        =   17
         Top             =   1080
         Width           =   11940
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   2415
         TabIndex        =   2
         Top             =   1305
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
End
Attribute VB_Name = "frmFNewProforma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Orders As New Collection
Dim strProforma As String
Dim strDelList As String
Dim bAcceptsExtraQty As Boolean
Dim iColNo As Integer

Public Sub EditProforma(TempProforma As String)
    
    On Error GoTo err
    
    
    Me.Hide
    cmbCompany.Enabled = False
    strProforma = TempProforma
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "SELECT * FROM VFPInvoice WHERE PInvoice='" & strProforma & "'", con, adOpenForwardOnly, adLockReadOnly
                  
        cmbCompany.ID = !CompanyRefID
        cmbCust.Text = ![CustCode] & ""
        cmbCountry = ![Country] & ""
        txtPInvoice = ![PInvoice] & ""
        OrderDT = ![DT] & ""
        cmbTradeTerms = ![TradeTerms] & ""
        
        cmbPayment = ![PaymentTerms] & ""
        
        txtLC = ![LCNo] & ""
        
        cmbTrans = ![TransShipment] & ""
        cmbPartial = ![PartialShipment] & ""
        txtPackaging = ![Packaging] & ""
        DeliveryDT = ![ProformaDeliveryDT] & ""
        
        txtGrossWeight = Val(!GrossW & "")
        txtNetWeight = Val(!NetW & "")
        txtPackages = !Packages & ""
        
        cmbSMethod = ![SMethod] & ""
        'txtBL = ![BL] & ""
        'BLDT = ![BLDT] & ""
        
        'txtFormE = ![FormENo] & ""
        'If ![LoadPortID] <> 4 Then
        '    cmbLoadCountry = ![LoadCountry] & ""
        '    cmbLoadCity = ![LoadCity] & ""
        '    cmbLoadPort = ![LoadPort] & ""
        'Else
        '    cmbLoadCountry.Text = "Pakistan"
        'End If
        
        cmbShipAgent.ID = ![ShippingAgentID] & ""
        cmbInsurAgent.ID = ![InsuranceAgentID] & ""
        txtShippingAddress = !ShippingAddress & ""
        txtInvoiceTo = !InvoiceTo & ""
        
        If IsNull(![LoadPortID]) = False Then
            cmbLoadCountry = ![LoadCountry] & ""
            cmbLoadCity = ![LoadCity] & ""
            cmbLoadPort = ![LoadPort] & ""
        Else
            cmbLoadCountry.Text = "Pakistan"
        End If
        
        'On Error Resume Next
        'If ![DisPortID] <> 4 Then
        '    cmbDisCountry = ![DisCountry] & ""
        '    cmbDisCity = ![DisCity] & ""
        '    cmbDisPort = ![DisPort] & ""
        'End If
        
        If IsNull(![DisPortID]) = False Then
            cmbDisCountry = ![DisCountry] & ""
            cmbDisCity = ![DisCity] & ""
            cmbDisPort = ![DisPort] & ""
        End If
        
        cmbThrough = ![BankName] & ""
        txtBankAddress = ![Address1] & ""
        'txtBatchNo = ![BatchNo] & ""
        txtDeclaration = ![Declarations] & ""
        cmbBank = ![Bank] & ""
        
        txtInvoiceof = !InvoiceOf & ""
        txtDelivery = !DeliveryDetail & ""
        
        txtNotes1 = !Notes1 & ""
        txtNotes2 = !Notes2 & ""
        
        txtTotal_Description = !Total_Description & ""
        
        txtDiscountPer = Val(!DiscountPer & "")
        txtDiscountAmt = Val(!DiscountAmount & "")
        txtInsuranceAmt = Val(!Insurance_Freight_Amount & "")
        txtInsuranceAmt_II = Val(!Insurance_Freight_Amount_II & "")
        
        txtDiscount_Description = !Discount_Description & ""
        txtInsurance_Description = !Insurance_Description & ""
        txtInsurance_Description_II = !Insurance_Description_II & ""
        txtCIF_Description = !CIF_Description & ""
                
        For i = 0 To cmbBankAccNo.ListCount - 1
            If Val(cmbBankAccNo.List(i, 1)) = Val(!OurBankID & "") Then
                cmbBankAccNo.ListIndex = i
                Exit For
            End If
        Next
        .Close
        
        If strCompany = "QEL" Then
            strOrderBy = "ORDER BY CompItemID"
        Else
            strOrderBy = ""
        End If
        .Open "SELECT * FROM VProformaDetail WHERE PInvoice='" & strProforma & "' " & strOrderBy, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, ![ID] & "'", .AbsolutePosition)
            ITM.Tag = ![CompITemID] & "" 'cmbItems.List(cmbItems.ListIndex, 2)
            
            ITM.ListSubItems.add(, , ![OrderNo]).Tag = Val(!EntryID & "") 'ProformaEntryID
            If strCompany = "QEL" Then
                ITM.ListSubItems.add , , !CompITemID & " (" & ![ITemID] & ")"
            Else
                ITM.ListSubItems.add , , ![ITemID] & ""
            End If
            
            ITM.ListSubItems.add(, ![ItemDescription] & "", ![ItemDescription] & "").Tag = ![Description] & ""
            
            ITM.ListSubItems.add , , !Unit & ""
            
            ITM.ListSubItems.add(, , Format(![Price] & "", "#0.0000")).Tag = Val(![CustomPrice] & "")
            ITM.ListSubItems.add , , Val(![ProformaQty] & "")
            
            ITM.ListSubItems(6).Tag = Val(![ProformaQty] & "") + Val(![InvQty] & "")
            ITM.ListSubItems.add(, , Val(![ProformaQty] & "") * Val(![Price])).Tag = Val(![ProformaQty] & "") - Val(![CustomInvQty] & "")
            'ITM.SubItems(7) = Val(![ProformaQty] & "") * Val(![Price])
            ITM.ListSubItems(4).Tag = !Item_Highlight
            Set ITM = LV.ListItems.add
            If iProductionType = 0 Then
                'ITM.SubItems(3) = "Batch # " & txtBatchNo & "   Lot # " & txtBatchNo & "/" & ![ItemCode] & ""
                ITM.SubItems(3) = "Delivery:" & Format(!DeliveryDTItem, "dd-MMM-yyyy") & " | " & "Fin:" & !QualityItem & " | " & "Stamps:" & !StampsItem & ""
            Else
                ITM.SubItems(3) = "Batch # " & txtBatchNo & "   Lot # " & txtBatchNo & "/" & ![ItemCode] & ""
            End If
            
            
            .MoveNEXT
        Loop
        .Close
        
        .Open "SELECT SteelWeight From FPInvoiceSteelWeight Where PInvoice='" & strProforma & "'", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            txtSteelWeight = Round(Val(![SteelWeight] & ""), 2)
        End If
        .Close
        
        .Open "SELECT * FROM FPInvoiceOtherCharges WHERE PInvoice='" & strProforma & "' Order By SNo", con, adOpenForwardOnly
        LVOtherCharges.ListItems.Clear
        Do Until .EOF
            Set ITM = LVOtherCharges.ListItems.add(, , Val(!SNo & ""))
            ITM.SubItems(1) = !ChargesDetail & ""
            ITM.SubItems(2) = Val(!ChargesAmt & "")
            .MoveNEXT
        Loop
        .Close
        txtOtherChargesSNo.Text = LVOtherCharges.ListItems.count + 1
        'Now Load Orders In Orders Collection
        '.Open "Select OrderNo from FProformaOrders Where PInvoice='" & strProforma & "'", con, adOpenForwardOnly, adLockReadOnly
        'ReDim Orders(.RecordCount)
        'For i = 0 To .RecordCount - 1
        '    Orders.Add .Fields(0).value, .Fields(0).value
        '    .MoveNext
        'Next i
        '.Close
        
    End With
    Call CalculateTotals
    
    cmbItemSearchProforma.ClearVals
    cmbItemSearchProforma.AddVals con, "ItemID + ' {' + CompItemID + '} {' + OrderNo + '}'", "VProformaDetail", "ItemID+'|'+OrderNo", "WHERE PInvoice='" & strProforma & "' " & strOrderBy
    
    txtBatchNo = GetSingleStringValue("OrderRcvdVia", "FCustomerOrders", " WHERE OrderNo='" & strProforma & "'")
    'txtQty = 0
    'cmbItems.SetFocus
    
    Set rs = Nothing
    Me.Show
    Exit Sub
err:
    MsgBox err.Description
    
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

Private Sub cmbCompany_matched()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select DISTINCT Bank From BankList Where CompanyRefID=" & cmbCompany.ID, con, adOpenForwardOnly, adLockReadOnly
        cmbBank.Clear
        Do Until .EOF
            cmbBank.AddItem .Fields(0) & ""
            'cmbBank.List(.AbsolutePosition - 1, 1) = .Fields(1) & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()
    
    On Error GoTo err
    If cmbCountry.MatchFound = False Then Exit Sub
    If cmbCountry.MatchFound = False Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Max(DT) From FPInvoice Where CustCode='" & cmbCust.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount = 0 Then
            OrderDT = ServerDate
        ElseIf IsNull(.Fields(0)) Then
            OrderDT = Date
        Else
            OrderDT = DateAdd("M", 1, .Fields(0))
        End If
        .Close
        .Open "SELECT AcceptsExtraQty,DefaultShipMethod,PaymentTerms FROM ForeignCustomers WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        bAcceptsExtraQty = !AcceptsExtraQty
        cmbSMethod = !DefaultShipMethod & ""
        cmbPayment = !PaymentTerms & ""
        .Close
    End With
    Set rs = Nothing
    cmbDisCountry.Text = cmbCountry
    txtCurrency = cmbCountry.List(cmbCountry.ListIndex, 1)
    Call GetOrders
    Call GetBanks
    
    
    Call AddToCombo(cmbShippingAddresses, "ShippingAddress", "ForeignCustomersShippingAddresses", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'", , "EntryID")
    Call AddToCombo(cmbInvoiceTos, "InvoiceTo", "ForeignCustomers_InvoiceTo", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'", , "EntryID")
    
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub GetBanks()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select BankID,BankName,Address1 From FCustBanks Where CustCode+Country='" & cmbCust.Text & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbThrough.Clear
        Do Until .EOF
            cmbThrough.AddItem ![BankName] & ""
            cmbThrough.List(cmbThrough.ListCount - 1, 1) = ![BankID]
            cmbThrough.List(cmbThrough.ListCount - 1, 2) = ![Address1] & ""
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
        .Open "Select Country,Curr From ForeignCustomers Where CustCode='" & cmbCust.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            cmbCountry.List(cmbCountry.ListCount - 1, 1) = ![Curr] & ""
            cmbDisCountry.AddItem ![Country] & ""
            .MoveNEXT
        Loop
        .Close
    End With
    If cmbCountry.ListCount > 0 Then cmbCountry.ListIndex = 0
    Set rs = Nothing
    If strProforma = "" Then
        If strCompany = "Sunlike" Then
            txtPInvoice = Left(txtPInvoice, 6) & Mid(cmbCust.ID, 4)
        ElseIf iProductionType = 0 Then
            'txtPInvoice = "DFG"
        Else
            txtPInvoice = Left(txtPInvoice, 6) & Right(cmbCust.ID, 4)
        End If
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbCust_UnMatched()
    'cmbCountry.ClearVals
End Sub

Private Sub cmbItemSearch_Change()
    Call cmbItemSearch_Click
End Sub

Private Sub cmbItemSearch_Click()

    If cmbItemSearch.MatchFound = False Then Exit Sub
    Dim strOrderNo As String, strItemCode As String
    strOrderNo = cmbItemSearch.List(cmbItemSearch.ListIndex, 2)
    strItemCode = cmbItemSearch.List(cmbItemSearch.ListIndex, 1)
    Dim i As Integer
    For i = 1 To LV2.ListItems.count
        If LV2.ListItems(i).SubItems(1) = strOrderNo And LV2.ListItems(i).SubItems(2) = strItemCode Then
            LV2.ListItems(i).Selected = True
            LV2.ListItems(i).EnsureVisible
            LV2.SetFocus
            Exit For
        End If
    Next
    
    
End Sub

Private Sub cmbItemSearchProforma_matched()

    Dim strOrderNo As String, strItemCode As String
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).SubItems(2) & "|" & LV.ListItems(i).SubItems(1) = cmbItemSearchProforma.ID Then
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

Private Sub CommandButton1_Click()

End Sub

Private Sub cmdAddAll_Click()
    
    If LV2.ListItems.count = 0 Then Exit Sub
    Dim i As Integer
    For i = 1 To LV2.ListItems.count Step 2
        
        LV2.ListItems(i).Selected = True
        txtQty1 = Val(LV2.ListItems(i).SubItems(6))
        Call AddtoProforma
    Next
    
'    With rs
'        .Open "Select * from VFCustOrders Where OrderNo='" & cmbOrders.Text & "' And InvQty>0", con, adOpenForwardOnly, adLockReadOnly
'
'        cmbTradeTerms = ![TradeTerms] & ""
'        cmbPartial = ![PartialShipment] & ""
'        cmbTrans = ![TransShipment] & ""
'        cmbPayment = ![PaymentTerms] & ""
'        txtPackaging = ![Packaging] & ""
'
'        DeliveryDT = Format(![DeliveryDT], "dd-MMM-yyyy")
'    End With
        
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

Private Sub cmdCalcWeights_Click()
    Dim dWeight As Double
    dWeight = GetSingleDoubleValue("SUM(Qty*FinishedWeight)", "VrptProformas", " WHERE PInvoice='" & txtPInvoice & "'")
    txtGrossWeight = Round(dWeight / 1000, 2)
    txtNetWeight = Round(dWeight / 1000, 2)
End Sub

Private Sub cmdRemoveAll_Click()

    If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Dim i As Integer, iIndex As Integer
    For i = LV.ListItems.count - 1 To 1 Step -2
        With LV.ListItems(i)
            If Val(.ListSubItems(7).Tag) > 0 Then
                'MsgBox "Can't Remove This Item." & vbNewLine & "It's Being Used In Custom Invoice.", vbInformation
            Else
                If Val(.ListSubItems(1).Tag) > 0 Then
                    strDelList = strDelList & Val(.ListSubItems(1).Tag) & ","
                End If
                iIndex = .Index
                LV.ListItems.Remove iIndex
                LV.ListItems.Remove iIndex
            End If
        End With
    Next
    
End Sub

Private Sub cmdSavenStay_Click()
    If Saved Then
        Call EditProforma(txtPInvoice.Text)
    End If
End Sub

Private Sub cmdUpdatePrices_Click()

    If MsgBox("Are you sure to Update All Prices in this Proforma by " & Val(txtUpatePrice) & " %", vbYesNo + vbQuestion) = vbYes Then
        For i = 1 To LV.ListItems.count Step 2
            With LV.ListItems(i)
                .SubItems(5) = Val(.SubItems(5)) + ((Val(.SubItems(5)) / 100) * Val(txtUpdatePrice))
                .SubItems(7) = Val(.SubItems(5)) * Val(.SubItems(6))
            End With
        Next
    End If
    
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub cmdAddItem_Click()
    
    On Error GoTo err
    
    If Not (cmbCust.MatchFound And cmbCountry.MatchFound And cmbOrders.MatchFound) Then Exit Sub
    If cmbCust.Text = "" Or cmbCountry.Text = "" Or cmbOrders.Text = "" Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    
    Dim cmd As New ADODB.Command
    'Now Check If Same Item With The Same Size Already Exists Or Not.
    
    For i = 1 To LV2.ListItems.count
        If LV2.ListItems(i).SubItems(1) = cmbOrders.Text Then
            MsgBox "Order Already Added.", vbInformation
            Exit Sub
        End If
    Next i
    
    Dim strMsg As String
    strMsg = ""
    
    With rs
        .Open "SELECT * FROM VFCustOrders WHERE OrderNo='" & cmbOrders.Text & "' And InvQty>0 ORDER BY CompItemCode", con, adOpenForwardOnly, adLockReadOnly
        
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
        
        DeliveryDT = Format(![DeliveryDT], "dd-MMM-yyyy")
        
        If OrderDT.Tag = "" Then
            OrderDT = ![DT] & ""
            OrderDT.Tag = "One Added"
        Else
            If DateDiff("d", OrderDT, ![DT]) > 0 Then
                OrderDT = ![DT] & ""
            End If
        End If
        
        Dim ITM As ListItem
        Dim iSNo As Integer
        
        LV2.ListItems.Clear
        cmbItemSearch.Clear
        Do Until .EOF
            
            If LV2.ListItems.count = 0 Then
                iSNo = 1
            Else
                iSNo = Val(LV2.ListItems(LV2.ListItems.count - 1).Text + 1)
            End If
        
            'Set ITM = lv2.ListItems.Add(, ![OrderNo] & ![ItemCode] & "'", iSno)
            Set ITM = LV2.ListItems.add(, ![OrderItemID] & "'", iSNo)
            ITM.ListSubItems.add , , ![OrderNo] & ""
            If strCompany = "QEL" Then
                ITM.ListSubItems.add , , !CompItemCode & "(" & ![ItemCode] & ")"
            Else
                ITM.ListSubItems.add , , ![ItemCode] & ""
            End If
            ITM.ListSubItems.add , , ![Description] & ""
            ITM.ListSubItems.add , , ![Unit] & ""
            ITM.ListSubItems.add(, , Val(![Price] & "")).Tag = Val(![CustomPrice] & "")
            ITM.ListSubItems.add , , Val(![InvQty] & "")
            ITM.ListSubItems(6).Tag = Val(![InvQty] & "")
            ITM.SubItems(7) = Val(![Price] & "") * Val(![InvQty] & "")
            
            Set ITM = LV2.ListItems.add
            If iProductionType = 0 Then
                ITM.SubItems(3) = "Delivery:" & Format(!DeliveryDTItem, "dd-MMM-yyyy") & " | " & "Fin:" & !QualityItem & " | " & "Stamps:" & !StampsItem & ""
            Else
                ITM.SubItems(3) = "Batch # " & txtBatchNo & "   Lot # " & txtBatchNo & "/" & ![ItemCode] & ""
            End If
            
            cmbItemSearch.AddItem !ItemCode & " {" & !CompItemCode & "} (" & !OrderNo & ")"
            cmbItemSearch.List(iSNo - 1, 1) = !ItemCode & ""
            cmbItemSearch.List(iSNo - 1, 2) = !OrderNo & ""
            
            .MoveNEXT
            
        Loop
        .Close
    End With
    
    'Dim iUp As Integer
    'iUp = UBound(Orders)
    'Orders(iUp) = cmbOrders
    'ReDim Preserve Orders(iUp + 1)
    
    Orders.add cmbOrders, cmbOrders
    
    Set rs = Nothing
    
    Call CalculateTotals
    
    
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
    If Val(LV.SelectedItem.ListSubItems(7).Tag) > 0 Then
        MsgBox "Can't Remove This Item." & vbNewLine & "It's Being Used In Custom Invoice.", vbInformation
        Exit Sub
    End If
    
    Dim Temp As Integer
    Temp = LV.SelectedItem.Index
    
    If Val(LV.SelectedItem.ListSubItems(1).Tag) > 0 Then
        strDelList = strDelList & Val(LV.SelectedItem.ListSubItems(1).Tag) & ","
    End If
    
    LV.ListItems.Remove Temp
    LV.ListItems.Remove Temp
    
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
        Me.Show
    End If
End Sub

Private Sub cmdSClose_Click()
    If Saved Then
        Unload Me
    End If
End Sub


Private Sub Form_Load()
   
    On Error GoTo err
    strDelList = ""
    SSTab1.Tab = 0
    OrderDT.Tag = ""
    
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where Active=1")
    
    OrderDT = ServerDate
    
    Me.KeyPreview = True
    cmbCust.ListHeight = 2500
    
    cmbShipAgent.ListHeight = 2000
    cmbInsurAgent.ListHeight = 2000
    
    Call FillCmbs
    
    iColNo = 0
    strProforma = ""
    Exit Sub
err:
    MsgBox err.Description
   
End Sub

Public Sub AddNewProforma()

    Dim PNo As Long
    Dim BatchNo As Long
    
    PNo = GetProformaNo
    BatchNo = GetBatchNo
    
    If PNo = 0 Then
        Unload Me
        Exit Sub
    ElseIf BatchNo = 0 Then
        Unload Me
        Exit Sub
    End If
    
'    If iProductionType = 0 Then
'        txtPInvoice = "PI-" & PNo '"PSPL-" & Format(PNo, "0000") & Format(ServerDate, "-MM-yy")
'    Else
'        txtPInvoice = Format(PNo, "00000/") '"PSPL-" & Format(PNo, "0000") & Format(ServerDate, "-MM-yy")
'    End If
    If strCompany = "Tecno" Then
        txtPInvoice = Format(PNo, "00000/") '"PSPL-" & Format(PNo, "0000") & Format(ServerDate, "-MM-yy")
    Else
        txtPInvoice = "PI-" & PNo '"PSPL-" & Format(PNo, "0000") & Format(ServerDate, "-MM-yy")
    End If
    
    'txtBatchNo = BatchNo
    
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)

    For i = 1 To Orders.count
        Orders.Remove 1
    Next i
    Set Orders = Nothing
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    FALL.Move ((Me.ScaleWidth - FALL.Width) / 2), ((Me.ScaleHeight - FALL.Height) / 2)
End Sub

Private Sub lblTotalAmt_Change()
    Call Calculate_Grand_Total
End Sub

Private Sub LV_Click()
    If LV.ListItems.count > 0 Then
        Call LV_ItemClick(LV.SelectedItem)
    End If
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    If LV.SelectedItem.Index Mod 2 = 0 Then Exit Sub
    
    If Not (iColNo = 6 Or iColNo = 4) Then iColNo = 7
    With txtQty
        .Move LV.ColumnHeaders(iColNo).Left + LV.Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(iColNo).Width
        
        .Text = (LV.SelectedItem.SubItems(iColNo - 1))
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    If Item.Index Mod 2 = 0 Then
        cmdDelete.Visible = False
        Exit Sub
    End If
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
            dTotal = dTotal + Val(.ListItems(i).SubItems(7) & "")
            DPcsTotal = DPcsTotal + Val(.ListItems(i).SubItems(6) & "")
        Next i
    End With
    'txttQty = dQty
    lblTotal = dTotal
    lblTotalAmt = dTotal
    lblTotWords = ConvertInWords(CCur(dTotal), txtCurrency.Text)
    lblTotPcs = DPcsTotal
    
End Sub

Private Function Saved() As Boolean
    
    On Error GoTo err
    
    Dim iBankID As Integer
    Dim iLoadID As Integer
    Dim iDisID As Integer
    
    Saved = False
    If cmbCompany.MatchFound = False Then
        MsgBox "Please Select Company.", vbInformation
        Exit Function
    ElseIf txtPInvoice = "" Then
        MsgBox "Please Enter Proforma Invoice No.", vbInformation
        txtPInvoice.SetFocus
        Exit Function
    End If
    
    If LV.ListItems.count = 0 Then
        MsgBox "No Orders Added To This Proforma.", vbInformation
        Exit Function
    End If
    
'    If cmbBank.MatchFound = False Then
'        MsgBox "Invalid Bank Selected.", vbInformation
'        SSTab1.Tab = 1
'        cmbBank.SetFocus
'        Exit Function
'    ElseIf cmbBankAccNo.MatchFound = False Then
'        MsgBox "Invalid Bank Account.", vbInformation
'        SSTab1.Tab = 1
'        cmbBankAccNo.SetFocus
'        Exit Function
'    End If
    
    If LCase(cmbPayment) = "letter of credit" Then
        If txtLC = "" Then
            MsgBox "You Must Enter LC # If Payment Terms Is 'Letter Of Credit.", vbInformation
            Exit Function
        End If
    End If
    
    If cmbThrough.ListIndex = -1 Then
        'iBankID = 4
        iBankID = -1
    Else
        iBankID = cmbThrough.List(cmbThrough.ListIndex, 1)
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
    
    Dim iOurBankID As Integer
    If cmbBankAccNo.MatchFound Then
        iOurBankID = cmbBankAccNo.List(cmbBankAccNo.ListIndex, 1)
    Else
        iOurBankID = -1
    End If
    
    Call StartTrans(con)
    '----------------------------------------------------------------
    If strProforma <> "" Then
    
        con.Execute "UPDATE FOrderItems SET InvQty=InvQty+IsNull((Select " & _
        " Qty FROM FProformaOrders WHERE PInvoice='" & strProforma & "' And OrderEntryID=ID),0) " ')"
         
        'con.Execute "Delete from FPInvoice Where PInvoice='" & _
         strProforma & "'"
        con.Execute "UPDATE FPInvoice SET PInvoice='" & txtPInvoice & "',DT='" & OrderDT & "',SMethod='" & cmbSMethod.Text & "',LoadPort=" & iLoadID & _
        ",DischargePort=" & iDisID & ",Shipment='" & txtShipment & "',TradeTerms='" & cmbTradeTerms.Text & "',PaymentTerms='" & _
        cmbPayment & "',LCNo='" & txtLC & "',BankID=" & iBankID & ",Packaging='" & txtPackaging & "',Carton='" & txtCartons & _
        "',PartialShipment='" & cmbPartial & "',TransShipment='" & cmbTrans & "',DeliveryDT='" & DeliveryDT & "',ShippingAgent=" & _
        Val(cmbShipAgent.ID) & ",InsuranceAgent=" & Val(cmbInsurAgent.ID) & ",BatchNo='" & txtBatchNo & "',Declarations='" & _
        txtDeclaration & "',OurBankID=" & iOurBankID & ",InvoiceOf='" & txtInvoiceof & "',DeliveryDetail='" & txtDelivery & _
        "',GrossW=" & Val(txtGrossWeight) & ",NetW=" & Val(txtNetWeight) & ",Packages='" & txtPackages & "',ShippingAddress='" & _
        txtShippingAddress & "',Notes1='" & txtNotes1 & "',Notes2='" & txtNotes2 & "',DiscountPer=" & Val(txtDiscountPer) & _
        ",DiscountAmount=" & Val(txtDiscountAmt) & ",Insurance_Freight_Amount=" & Val(txtInsuranceAmt) & _
        ",Insurance_Freight_Amount_II=" & Val(txtInsuranceAmt_II) & ",Discount_Description='" & txtDiscount_Description & _
        "',Insurance_Description='" & txtInsurance_Description & "',Insurance_Description_II='" & txtInsurance_Description_II & _
        "',CIF_Description='" & txtCIF_Description & "',Total_Description='" & txtTotal_Description & "',InvoiceTo='" & txtInvoiceTo & "' WHERE PInvoice='" & strProforma & "'"
        
        con.Execute "UPDATE FPInvoiceSteelWeight SET SteelWeight=" & Val(txtSteelWeight) & " Where PInvoice='" & txtPInvoice & "'"
        If strDelList <> "" Then
            strDelList = Left(strDelList, Len(strDelList) - 1)
            con.Execute "Delete From FProformaOrders Where EntryID In(" & strDelList & ")"
        End If
        
    Else
    
        con.Execute "INSERT INTO FPInvoice(PInvoice,DT,CustCode,Country,SMethod," & _
        "LoadPort,DischargePort,Shipment,TradeTerms,PaymentTerms,LCNo,BankID,Packaging," & _
        "Carton,PartialShipment,TransShipment,DeliveryDT,ShippingAgent,InsuranceAgent," & _
        "BatchNo,Declarations,OurBankID,CompanyRefID,InvoiceOf,DeliveryDetail,GrossW,NetW," & _
        "Packages,ShippingAddress,Notes1,Notes2,DiscountPer,DiscountAmount,Insurance_Freight_Amount,Total_Description,InvoiceTo) VALUES('" & _
        txtPInvoice & "','" & OrderDT & "','" & cmbCust.Text & "','" & cmbCountry.Text & "','" & cmbSMethod.Text & "'," & iLoadID & _
        "," & iDisID & ",'" & txtShipment & "','" & cmbTradeTerms.Text & "','" & cmbPayment & "','" & txtLC & "'," & _
        iBankID & ",'" & txtPackaging & "','" & txtCarton & "','" & cmbPartial & "','" & cmbTrans & "','" & _
        DeliveryDT & "'," & Val(cmbShipAgent.ID) & "," & Val(cmbInsurAgent.ID) & ",'" & txtBatchNo & "','" & _
        txtDeclaration & "'," & iOurBankID & "," & cmbCompany.ID & ",'" & txtInvoiceof & "','" & _
        txtDelivery & "'," & Val(txtGrossWeight) & "," & Val(txtNetWeight) & ",'" & txtPackages & "','" & txtShippingAddress & _
        "','" & txtNotes1 & "','" & txtNotes2 & "'," & Val(txtDiscountPer) & "," & Val(txtDiscountAmt) & "," & Val(txtInsuranceAmt) & ",'" & txtTotal_Description & "','" & txtInvoiceTo & "')"
    
        con.Execute "INSERT INTO FPInvoiceSteelWeight(PInvoice,SteelWeight) VALUES('" & txtPInvoice & "'," & Val(txtSteelWeight) & ")"
    End If
    
    'Save Order
    'con.Execute "Insert Into FCustomerOrders (CustCode,Country," & _
     "OrderNo,DT,TradeTerms,PartialShipment,PaymentTerms," & _
     "TransShipment,Packaging,DeliveryDT) values('" & _
     cmbCust.ID & "','" & cmbCountry & "','" & txtOrderNo & _
     "','" & OrderDT & "','" & cmbTradeTerms & "','" & cmbPartial & _
     "','" & cmbPayment & "','" & cmbTrans & "','" & txtPackaging & _
     "','" & DeliveryDT & "')"
  
    '***********************************
    '*Now Save Orders With Proforma No.*
    '***********************************
    'Dim iUp As Integer
    'iUp = UBound(Orders)
    'Dim UpLimit As Integer
    'UpLimit = Orders.Count
    
    'For i = 1 To UpLimit
    '    con.Execute "Insert into FProformaOrders(PInvoice,OrderNo)" & _
    '    "Values('" & txtPInvoice & "','" & Orders(i) & "')"
    'Next i
    For i = 1 To LV.ListItems.count Step 2
        With LV.ListItems(i)
            If Val(.ListSubItems(1).Tag) <> 0 Then
                con.Execute "UPDATE FProformaOrders SET Qty=" & Val(.SubItems(6)) & _
                 ",InvQty=" & Val(.SubItems(6)) - Val(.ListSubItems(7).Tag) & ",Price=" & Val(.SubItems(5)) & ",CustomPrice=" & Val(.SubItems(5)) & _
                 ",Item_Highlight=" & Abs(CBool(.ListSubItems(4).Tag)) & ",ItemDescription='" & .SubItems(3) & "' WHERE EntryID=" & Val(.ListSubItems(1).Tag)
            Else
                con.Execute "INSERT INTO FProformaOrders(PInvoice,OrderEntryID,Qty,InvQty,Price,CustomPrice,Item_Highlight,ItemDescription) " & _
                 "VALUES('" & txtPInvoice & "'," & Val(LV.ListItems(i).key) & "," & Val(.SubItems(6)) & "," & Val(.SubItems(6)) & _
                 "," & Val(.ListSubItems(5)) & "," & Val(.ListSubItems(5).Tag) & "," & Abs(CBool(.ListSubItems(4).Tag)) & ",'" & .SubItems(3) & "')"
            End If
            con.Execute "UPDATE FOrderItems SET InvQty=InvQty-" & _
             Val(LV.ListItems(i).SubItems(6)) & " WHERE ID=" & Val(LV.ListItems(i).key)
         End With
    Next i
    
    
    If strProforma <> "" Then
        con.Execute "DELETE FROM FPInvoiceOtherCharges WHERE PInvoice='" & strProforma & "'"
    End If
    
    For i = 1 To LVOtherCharges.ListItems.count
        With LVOtherCharges.ListItems(i)
            con.Execute "INSERT INTO FPInvoiceOtherCharges(PInvoice,ChargesDetail,ChargesAmt,SNo)" & _
             " VALUES('" & IIf(strProforma = "", txtPInvoice.Text, strProforma) & "','" & _
             .SubItems(1) & "'," & Val(.SubItems(2)) & "," & Val(.Index) & ")"
        End With
    Next
    
    con.CommitTrans
    
    Saved = True
    If chkPrint.Value Then
        Call PrintProforma(txtPInvoice.Text)
    End If
    Exit Function

err:
    MsgBox err.Description
    
End Function

Private Sub PrintProforma(p_strPInvoice As String)
    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptFProforma.rpt") '
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VrptProforma.PInvoice}='" & p_strPInvoice & "' ", rpt
    
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
        .ListIndex = 0
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
    
    
'    With cmbSMethod
'        .AddItem "By Air"
'        .AddItem "By Sea"
'        .AddItem "By Courier"
'        .AddItem "By EMS"
'        .ListIndex = 0
'    End With
    Call AddToCombo(cmbSMethod, "Description", "ShippingMethods", " ORDER BY EntryID")
    
    cmbLoadCountry.AddItem "Pakistan"
    cmbLoadCountry.ListIndex = 0
    
    cmbShipAgent.ClearVals
    Call cmbShipAgent.AddVals(con, "Name", "ShippingAgents", "EntryID", " Where Active=1")
    
    cmbInsurAgent.ClearVals
    Call cmbInsurAgent.AddVals(con, "Name", "InsuranceAgents", "EntryID", " Where Active=1")
    
    cmbCompany.ListHeight = 6000
    cmbCompany.AddVals con, "CompanyName", "Companies", "EntryID", " Order By CompanyName"
    '----------------------------------------------------------------------------------------
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
    If cmbCompany.MatchFound = False Then
        MsgBox "Select Company.", vbInformation
        Exit Sub
    End If
    
    With cmbCountry
        If Not .MatchFound Then Exit Sub
    End With
    
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "SP_OrdersForProforma"
        Set .ActiveConnection = con
        '.Parameters.Refresh
        .Parameters("@CompanyID").Value = cmbCompany.ID
        .Parameters("@CustCode").Value = cmbCust.Text
        .Parameters("@Country").Value = cmbCountry.Text
        Set rs = .Execute
    End With
    
    With rs
        cmbOrders.Clear
        Do Until .EOF
            cmbOrders.AddItem .Fields(0) & ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        If X > LV.ColumnHeaders(i).Left And X < LV.ColumnHeaders(i).Left + LV.ColumnHeaders(i).Width Then
            iColNo = i
            Exit For
        End If
    Next
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If LV.ListItems.count = 0 Then Exit Sub
    If Button <> 2 Then Exit Sub
    Dim strRet As String
    If LV.SelectedItem.ListSubItems(4).Tag = "" Then
        strRet = ShowPopUpMenu1("Highlight Item")
    Else
        If CBool(LV.SelectedItem.ListSubItems(4).Tag) Then
            strRet = ShowPopUpMenu1("Un-Highlight Item")
        Else
            strRet = ShowPopUpMenu1("Highlight Item")
        End If
    End If
    
    If strRet = "Highlight Item" Then
        LV.SelectedItem.ListSubItems(4).Tag = True
    ElseIf strRet = "Un-Highlight Item" Then
        LV.SelectedItem.ListSubItems(4).Tag = False
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


Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        cmbItemSearchProforma.Visible = True
    Else
        cmbItemSearchProforma.Visible = False
    End If
End Sub

Private Sub txtDiscountAmt_Change()
    Call Calculate_Grand_Total
End Sub

Private Sub txtDiscountPer_Change()

    Dim dPer As Double, dDisAmt As Double
    dPer = Val(txtDiscountPer)
    dDisAmt = (Val(lblTotalAmt) / 100) * dPer
    dDisAmt = Round(dDisAmt, 2)
    txtDiscountAmt = dDisAmt
    
End Sub

Private Sub txtInsuranceAmt_Change()
    Call Calculate_Grand_Total
End Sub

Private Sub txtInsuranceAmt_II_Change()
    Call Calculate_Grand_Total
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        If iColNo = 6 Then
            If Val(txtQty) < 0 Then
                MsgBox "Invalid Price.", vbInformation
                Exit Sub
            End If
            
            LV.SelectedItem.SubItems(iColNo - 1) = Val(txtQty)
            LV.SelectedItem.SubItems(7) = Val(txtQty) * LV.SelectedItem.SubItems(5)
        ElseIf iColNo = 4 Then
            LV.SelectedItem.SubItems(iColNo - 1) = txtQty
        Else
            If Val(txtQty) > Val(LV.SelectedItem.ListSubItems(6).Tag) Then
                MsgBox "Invalid Qty." & vbNewLine & "You can only add " & Val(LV.SelectedItem.ListSubItems(6).Tag) & " Qty.", vbInformation
                Exit Sub
            ElseIf Val(txtQty) < Val(LV.SelectedItem.ListSubItems(7).Tag) And Val(LV.SelectedItem.ListSubItems(7).Tag) > 0 Then
                MsgBox "Invalid Qty." & vbNewLine & "Minimum Qty Is " & Val(LV.SelectedItem.ListSubItems(7).Tag), vbInformation
                Exit Sub
            End If
            LV.SelectedItem.SubItems(6) = Val(txtQty)
            LV.SelectedItem.SubItems(7) = Val(txtQty) * LV.SelectedItem.SubItems(5)
        End If
        txtQty.Visible = False
        Call CalculateTotals
    ElseIf KeyAscii = 27 Then
        txtQty.Visible = False
    End If
    
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lv2_Click()
    If LV2.ListItems.count > 0 Then
        Call lv2_ItemClick(LV2.SelectedItem)
    End If
End Sub

Private Sub LV2_DblClick()

    If LV2.ListItems.count = 0 Then Exit Sub
    If LV2.SelectedItem.Index Mod 2 = 0 Then Exit Sub
    With txtQty1
        .Move LV2.ColumnHeaders(7).Left + LV2.Left, LV2.Top + LV2.SelectedItem.Top, LV2.ColumnHeaders(7).Width
        .Text = Val(LV2.SelectedItem.SubItems(6))
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
        
        If Val(txtQty1) > Val(LV2.SelectedItem.ListSubItems(6)) Then
            MsgBox "Invalid Qty." & vbNewLine & "You can only add " & Val(LV2.SelectedItem.ListSubItems(6)) & " Qty.", vbInformation
            Exit Sub
        ElseIf Val(txtQty1) <= 0 Then
            MsgBox "Invalid Qty.", vbInformation
            Exit Sub
        End If
        Call AddtoProforma
    ElseIf KeyAscii = 27 Then
        txtQty1.Visible = False
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub AddtoProforma()

    Call MoveToFirstLV
        
    LV2.SelectedItem.SubItems(6) = LV2.SelectedItem.SubItems(6) - Val(txtQty1)
    LV2.SelectedItem.SubItems(7) = Val(txtQty1) * LV2.SelectedItem.SubItems(5)
    txtQty1.Visible = False
    
    Call CalculateTotals
    LV2.SetFocus
    
End Sub

Private Sub MoveToFirstLV()
    
    On Error GoTo err
    Dim iSNo As Integer
    Dim ITM As ListItem
    
    
    If LV.ListItems.count = 0 Then
        iSNo = 1
    Else
        iSNo = Val(LV.ListItems(LV.ListItems.count - 1).Text + 1)
    End If
    
    
    Set ITM = LV.ListItems.add(, LV2.SelectedItem.key, iSNo)
    
ComeHereForEdit:
    For i = 1 To LV2.SelectedItem.ListSubItems.count
        ITM.ListSubItems.add , , LV2.SelectedItem.ListSubItems(i).Text
        ITM.ListSubItems(i).Tag = LV2.SelectedItem.ListSubItems(i).Tag
    Next i
    ITM.ListSubItems(6).Text = txtQty1
    ITM.ListSubItems(4).Tag = 0 'Item_Highlight
    ITM.SubItems(7) = Val(LV2.SelectedItem.SubItems(5) & "") * Val(txtQty1)

    Dim iIndex As Integer
    iIndex = LV2.SelectedItem.Index
    Set ITM = LV.ListItems.add
    ITM.SubItems(3) = LV2.ListItems(iIndex + 1).SubItems(3)
        
    Exit Sub
err:
    If err.Description = "Key is not unique in collection" And err.Number = 35602 Then
        Set ITM = LV.ListItems(LV2.SelectedItem.key)
        ITM.SubItems(6) = Val(ITM.SubItems(6)) + Val(txtQty1)
        ITM.SubItems(7) = Val(ITM.SubItems(6)) * Val(ITM.SubItems(5))
        
    Else
        MsgBox err.Description
    End If
End Sub

Private Sub Calculate_Grand_Total()
    lblGTAmt = Val(lblTotalAmt) - Val(txtDiscountAmt) + Val(txtInsuranceAmt) + Val(txtInsuranceAmt_II)
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

