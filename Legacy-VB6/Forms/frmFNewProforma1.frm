VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmFNewProforma1 
   ClientHeight    =   9735
   ClientLeft      =   60
   ClientTop       =   60
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
   ScaleHeight     =   9735
   ScaleWidth      =   12195
   WindowState     =   2  'Maximized
   Begin VB.Frame FAll 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9675
      Left            =   45
      TabIndex        =   14
      Top             =   30
      Width           =   12120
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   90
         TabIndex        =   0
         Top             =   1035
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
         TabIndex        =   9
         Top             =   8640
         Width           =   1995
      End
      Begin VB.TextBox txtCurrency 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4590
         MaxLength       =   150
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1050
         Width           =   1425
      End
      Begin VB.TextBox txtPInvoice 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6045
         MaxLength       =   150
         TabIndex        =   3
         Top             =   1050
         Width           =   1710
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   0
         TabIndex        =   17
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
            TabIndex        =   18
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
            TabIndex        =   19
            Top             =   180
            Width           =   12000
         End
      End
      Begin MSComCtl2.DTPicker OrderDT 
         Height          =   300
         Left            =   7755
         TabIndex        =   4
         Top             =   1035
         Width           =   1875
         _ExtentX        =   3307
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
         Format          =   61079555
         CurrentDate     =   37250
      End
      Begin MSComCtl2.DTPicker DeliveryDT 
         Height          =   285
         Left            =   10245
         TabIndex        =   10
         Top             =   8640
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
         Format          =   61079555
         CurrentDate     =   37250
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6630
         Left            =   75
         TabIndex        =   26
         Top             =   1365
         Width           =   11940
         _ExtentX        =   21061
         _ExtentY        =   11695
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         TabCaption(0)   =   "Order Detail"
         TabPicture(0)   =   "frmFNewProforma1.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "cmbOrders"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label1(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdAddItem"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "LV"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cmdDelete"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "Other Info."
         TabPicture(1)   =   "frmFNewProforma1.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "LV2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame2"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Frame4"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Frame3"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).ControlCount=   4
         Begin VB.Frame Frame3 
            Height          =   2325
            Left            =   105
            TabIndex        =   43
            Top             =   345
            Width           =   5850
            Begin VB.TextBox txtFormE 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   60
               TabIndex        =   55
               Tag             =   "FormENo"
               Top             =   1950
               Width           =   1935
            End
            Begin VB.TextBox txtBL 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1950
               TabIndex        =   52
               Top             =   450
               Width           =   1710
            End
            Begin VB.TextBox txtBankAddress 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   1200
               Left            =   1980
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   44
               Top             =   1035
               Width           =   3765
            End
            Begin MSComCtl2.DTPicker BLDT 
               Height          =   270
               Left            =   3630
               TabIndex        =   54
               Tag             =   "BLDT"
               Top             =   450
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   476
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
               Format          =   61079555
               CurrentDate     =   37250
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Form E"
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
               Left            =   60
               TabIndex        =   56
               Top             =   1725
               Width           =   1935
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Date"
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
               Left            =   3645
               TabIndex        =   53
               Top             =   210
               Width           =   1710
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "AWB No."
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
               Left            =   1950
               TabIndex        =   51
               Top             =   210
               Width           =   1710
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
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
               Index           =   4
               Left            =   60
               TabIndex        =   49
               Top             =   210
               Width           =   1905
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
               Top             =   795
               Width           =   1920
            End
            Begin MSForms.ComboBox cmbThrough 
               Height          =   285
               Left            =   60
               TabIndex        =   46
               Top             =   1035
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
               Top             =   795
               Visible         =   0   'False
               Width           =   4050
            End
         End
         Begin VB.CommandButton cmdDelete 
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   -63510
            Picture         =   "frmFNewProforma1.frx":0038
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
            Left            =   6000
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
            Left            =   6000
            TabIndex        =   28
            Top             =   345
            Width           =   5835
            Begin MSForms.ComboBox cmbLoadPort 
               Height          =   285
               Left            =   3885
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
            Begin MSForms.ComboBox cmbLoadCity 
               Height          =   285
               Left            =   1995
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
            Begin MSForms.ComboBox cmbLoadCountry 
               Height          =   285
               Left            =   105
               TabIndex        =   30
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
               TabIndex        =   29
               Top             =   315
               Width           =   5640
            End
         End
         Begin MSComctlLib.ListView LV 
            Height          =   5550
            Left            =   -74700
            TabIndex        =   27
            Top             =   975
            Width           =   11130
            _ExtentX        =   19632
            _ExtentY        =   9790
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
            Height          =   3840
            Left            =   105
            TabIndex        =   33
            Top             =   2700
            Width           =   11730
            _ExtentX        =   20690
            _ExtentY        =   6773
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S No."
               Object.Width           =   1147
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Declarations"
               Object.Width           =   18918
            EndProperty
         End
         Begin MSForms.CommandButton cmdAddItem 
            Height          =   345
            Left            =   -72915
            TabIndex        =   42
            Top             =   615
            Width           =   1515
            Caption         =   "    Add"
            PicturePosition =   327683
            Size            =   "2672;609"
            Picture         =   "frmFNewProforma1.frx":056A
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
            Left            =   -74700
            TabIndex        =   41
            Top             =   450
            Width           =   1755
         End
         Begin MSForms.ComboBox cmbOrders 
            Height          =   285
            Left            =   -74700
            TabIndex        =   40
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
      End
      Begin VB.Frame Frame6 
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
         Height          =   660
         Left            =   6720
         TabIndex        =   15
         Top             =   8940
         Width           =   5280
         Begin MSForms.CommandButton cmdCancel 
            Height          =   360
            Left            =   3510
            TabIndex        =   13
            Top             =   195
            Width           =   1665
            Caption         =   "Cancel            "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmFNewProforma1.frx":088C
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
            TabIndex        =   11
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
            TabIndex        =   12
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
      Begin MSForms.ComboBox cmbPartial 
         Height          =   285
         Left            =   5520
         TabIndex        =   8
         Top             =   8640
         Width           =   2685
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4736;503"
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
         Top             =   8040
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
         TabIndex        =   25
         Top             =   8040
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
         TabIndex        =   24
         Top             =   8040
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
         TabIndex        =   23
         Top             =   8040
         Width           =   1800
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmFNewProforma1.frx":099E
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
         TabIndex        =   22
         Top             =   8400
         Width           =   11910
      End
      Begin MSForms.CheckBox chkPrint 
         Height          =   345
         Left            =   90
         TabIndex        =   21
         Top             =   9000
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
         TabIndex        =   6
         Top             =   8640
         Width           =   2685
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4736;503"
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
         Left            =   2805
         TabIndex        =   7
         Top             =   8640
         Width           =   2685
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4736;503"
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
         TabIndex        =   20
         Top             =   9255
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
         Left            =   9630
         TabIndex        =   5
         Top             =   1050
         Width           =   2400
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4233;503"
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
         Caption         =   $"frmFNewProforma1.frx":0A50
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
         TabIndex        =   16
         Top             =   795
         Width           =   11940
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   2415
         TabIndex        =   1
         Top             =   1050
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
Attribute VB_Name = "frmFNewProforma1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Orders As New Collection
Dim strProforma As String

Public Sub EditProforma(TempProforma As String)
    
    On Error GoTo ERR
    strProforma = TempProforma
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "Select * From VrptProformas Where PInvoice='" & strProforma & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCust.Text = ![CustCode] & ""
        cmbCountry = ![Country] & ""
        txtPInvoice = ![PInvoice] & ""
        OrderDT = ![DT] & ""
        cmbTradeTerms = ![TradeTerms] & ""
        
        cmbPayment = ![PaymentTerms] & ""
        cmbTrans = ![TransShipment] & ""
        cmbPartial = ![PartialShipment] & ""
        txtPackaging = ![Packaging] & ""
        DeliveryDT = ![DeliveryDT] & ""
        
        cmbSMethod = ![SMethod] & ""
        txtBL = ![BL] & ""
        BLDT = ![BLDT] & ""
        
        txtFormE = ![FormENo] & ""
        'If ![LoadPortID] <> 4 Then
        '    cmbLoadCountry = ![LoadCountry] & ""
        '    cmbLoadCity = ![LoadCity] & ""
        '    cmbLoadPort = ![LoadPort] & ""
        'Else
        '    cmbLoadCountry.Text = "Pakistan"
        'End If
        
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
        
        
        Do Until .EOF
            Set ITM = LV.ListItems.Add(, ![OrderNo] & ![ItemID] & "'", .AbsolutePosition)
            ITM.Tag = ![CompItemID] & "" 'cmbItems.List(cmbItems.ListIndex, 2)
            
            ITM.ListSubItems.Add , , ![OrderNo] '![CatID] & ""
            ITM.ListSubItems.Add , , ![ItemID] & ""
            ITM.ListSubItems.Add(, ![ItemName] & "", ![ItemName] & "").Tag = ![Description] & ""
            
            ITM.SubItems(4) = txtUnit
            
            ITM.SubItems(5) = Format(![Price] & "", "#0.00")
            ITM.SubItems(6) = Val(![Qty] & "")
            ITM.SubItems(7) = Val(![Qty] & "") * Val(![Price])
   
            .MoveNext
        Loop
        .Close
        
        'Now Load Orders In Orders Collection
        .Open "Select OrderNo from FProformaOrders Where PInvoice='" & strProforma & "'", con, adOpenForwardOnly, adLockReadOnly
        'ReDim Orders(.RecordCount)
        For i = 0 To .RecordCount - 1
            Orders.Add .Fields(0).value, .Fields(0).value
            .MoveNext
        Next i
        .Close
        
    End With
    Call CalculateTotals

    'txtQty = 0
    'cmbItems.SetFocus
    
    Set rs = Nothing
    Me.Show
    Exit Sub
ERR:
    MsgBox ERR.Description
    
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


Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()
    
    On Error GoTo ERR
    If cmbCountry.MatchFound = False Then Exit Sub
    Dim rs As New ADODB.Recordset
    txtCurrency = cmbCountry.List(cmbCountry.ListIndex, 1)
    Call GetOrders
    Call GetBanks
    
    Exit Sub
ERR:
    MsgBox ERR.Description
    
End Sub
Private Sub GetBanks()

    On Error GoTo ERR
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select BankID,BankName,Address1 From FCustBanks Where CustCode+Country='" & cmbCust.Text & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbThrough.Clear
        Do Until .EOF
            cmbThrough.AddItem ![BankName] & ""
            cmbThrough.List(cmbThrough.ListCount - 1, 1) = ![BankID]
            cmbThrough.List(cmbThrough.ListCount - 1, 2) = ![Address1] & ""
            .MoveNext
        Loop
        
        .Close
    End With
    Set rs = Nothing
    Exit Sub
ERR:
    MsgBox ERR.Description
    
End Sub
Private Sub cmbCust_matched()
    
    On Error GoTo ERR
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Country,Curr From ForeignCustomers Where CustCode='" & cmbCust.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            cmbCountry.List(cmbCountry.ListCount - 1, 1) = ![Curr] & ""
            cmbDisCountry.AddItem ![Country] & ""
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Exit Sub
ERR:
    MsgBox ERR.Description
    
End Sub

Private Sub cmbCust_UnMatched()
    'cmbCountry.ClearVals
End Sub

Private Sub cmbSMethod_Change()
    Call cmbSMethod_Click
End Sub

Private Sub cmbSMethod_Click()
    If cmbSMethod.ListIndex = 0 Then
        Label1(12).Caption = "AWB No."
    Else
        Label1(12).Caption = "BL No."
    End If
End Sub

Private Sub cmbTradeTerms_Change()
    Call cmbTradeTerms_Click
End Sub

Private Sub cmbTradeTerms_Click()

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

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub cmdAddItem_Click()
    
    On Error GoTo ERR
    If Not (cmbCust.MatchFound And cmbCountry.MatchFound And cmbOrders.MatchFound) Then Exit Sub
    If cmbCust.Text = "" Or cmbCountry.Text = "" Or cmbOrders.Text = "" Then Exit Sub
    Dim rs As New ADODB.Recordset
    
    Dim cmd As New ADODB.Command
    'Now Check If Same Item With The Same Size Already Exists Or Not.
    
    For i = 1 To LV.ListItems.Count
        If LV.ListItems(i).SubItems(1) = cmbOrders.Text Then
            MsgBox "Order Already Added.", vbInformation
            Exit Sub
        End If
    Next i
    
    Dim strMsg As String
    strMsg = ""
    With rs
        .Open "Select * from VFCustOrders Where OrderNo='" & cmbOrders.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        
        If LV.ListItems.Count > 0 Then
            
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
        
        cmbTrade = ![TradeTerms] & ""
        cmbPartial = ![PartialShipment] & ""
        cmbTrans = ![TransShipment] & ""
        cmbPayment = ![PaymentTerms] & ""
        txtPackaging = ![Packaging] & ""
        
        txtDeliveryDT = Format(![DeliveryDT], "dd-MM-yyyy")
        
        Dim ITM As ListItem
        Dim iSno As Integer
        
        
        Do Until .EOF
        
            If LV.ListItems.Count = 0 Then
                iSno = 1
            Else
                iSno = Val(LV.ListItems(LV.ListItems.Count).Text + 1)
            End If
            
            Set ITM = LV.ListItems.Add(, ![OrderNo] & ![ItemCode] & "'", iSno)
            ITM.SubItems(1) = ![OrderNo] & ""
            ITM.SubItems(2) = ![ItemCode] & ""
            ITM.SubItems(3) = ![Description] & ""
            ITM.SubItems(4) = ![Unit] & ""
            ITM.SubItems(5) = Val(![Price] & "")
            ITM.SubItems(6) = Val(![Qty] & "")
            ITM.SubItems(7) = Val(![Price] & "") * Val(![Qty] & "")
            .MoveNext
        Loop
        .Close
    End With
    
    'Dim iUp As Integer
    'iUp = UBound(Orders)
    'Orders(iUp) = cmbOrders
    'ReDim Preserve Orders(iUp + 1)
    
    Orders.Add cmbOrders, cmbOrders
    
    Set rs = Nothing
    
    Call CalculateTotals
    
    
    Exit Sub
    
ERR:
    MsgBox ERR.Description
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdDelete_Click()

    'LV.ListItems.Remove LV.SelectedItem.Index
    Dim strOrder As String
    strOrder = LV.SelectedItem.SubItems(1)
    LV.ListItems.Remove LV.SelectedItem.Index
    For i = 1 To LV.ListItems.Count
        If i > LV.ListItems.Count Then Exit For
        If LV.ListItems(i).SubItems(1) = strOrder Then
            LV.ListItems.Remove i
        End If
    Next i
    
    Orders.Remove strOrder
    cmdDelete.Visible = False
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
   
    On Error GoTo ERR
    
    SSTab1.Tab = 0
    
    
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode")
    OrderDT = ServerDate
    Me.KeyPreview = True
    cmbCust.ListHeight = 2500
    Call FillCmbs
           

    Exit Sub
ERR:
    MsgBox ERR.Description
   
End Sub
Public Sub AddNewProforma()

    Dim PNo As Long
    
    PNo = GetProformaNo
    
    If PNo = 0 Then
        Unload Me
        Exit Sub
    End If
    
    txtPInvoice = "HSPL-" & Format(PNo, "0000") & Format(ServerDate, "-MM-yy")
    
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    For i = 1 To Orders.Count
        Orders.Remove 1
    Next i
    Set Orders = Nothing
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
   FAll.Move ((Me.ScaleWidth - FAll.Width) / 2), ((Me.ScaleHeight - FAll.Height) / 2)
End Sub

Private Sub LV_Click()
    If LV.ListItems.Count > 0 Then
        Call LV_ItemClick(LV.SelectedItem)
    End If
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With cmdDelete
        .Move .Left, LV.Top + Item.Top
        .Visible = True
    End With
End Sub

Private Sub CalculateTotals()

    With LV
        For i = 1 To .ListItems.Count
            'dQty = dQty + Val(.TextMatrix(i, 4))
            DTotal = DTotal + Val(.ListItems(i).SubItems(7) & "")
        Next i
    End With
    'txttQty = dQty
    lblTotal = DTotal
    lblTotWords = ConvertInWords(CCur(DTotal), txtCurrency.Text)
    
End Sub
Private Function Saved() As Boolean
    
    On Error GoTo ERR
    Dim iBankID As Integer
    Dim iLoadID As Integer
    Dim iDisID As Integer
    
    Saved = False
    If txtPInvoice = "" Then
        MsgBox "Please Enter Proforma Invoice No.", vbInformation
        txtPInvoice.SetFocus
        Exit Function
    End If
    
    If LV.ListItems.Count = 0 Then
        MsgBox "No Orders Added To This Proforma.", vbInformation
        Exit Function
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
    
    
    Call StartTrans(con)
    '----------------------------------------------------------------
    If strProforma <> "" Then
        con.Execute "Delete from FPInvoice Where PInvoice='" & _
         strProforma & "'"
    End If
    
    'Save Order
    'con.Execute "Insert Into FCustomerOrders (CustCode,Country," & _
     "OrderNo,DT,TradeTerms,PartialShipment,PaymentTerms," & _
     "TransShipment,Packaging,DeliveryDT) values('" & _
     cmbCust.ID & "','" & cmbCountry & "','" & txtOrderNo & _
     "','" & OrderDT & "','" & cmbTradeTerms & "','" & cmbPartial & _
     "','" & cmbPayment & "','" & cmbTrans & "','" & txtPackaging & _
     "','" & DeliveryDT & "')"
     
     
    con.Execute "Insert into FPInvoice (PInvoice,DT,CustCode,Country,SMethod,BL,BLDT," & _
     "LoadPort,DischargePort,Shipment,TradeTerms,PaymentTerms,BankID,Packaging," & _
     "Carton,PartialShipment,TransShipment,DeliveryDT,FormENo) Values('" & txtPInvoice & "','" & _
     OrderDT & "','" & cmbCust.Text & "','" & _
     cmbCountry.Text & "','" & cmbSMethod.Text & "','" & txtBL & "','" & BLDT & "'," & iLoadID & _
     "," & iDisID & ",'" & _
     txtShipment & "','" & cmbTradeTerms.Text & "','" & cmbPayment & "'," & _
     iBankID & ",'" & txtPackaging & "','" & txtCarton & "','" & _
     cmbPartial & "','" & cmbTrans & "','" & _
     DeliveryDT & "','" & txtFormE & "')"
         
    '***********************************
    '*Now Save Orders With Proforma No.*
    '***********************************
    'Dim iUp As Integer
    'iUp = UBound(Orders)
    Dim UpLimit As Integer
    UpLimit = Orders.Count
    
    For i = 1 To UpLimit
        con.Execute "Insert into FProformaOrders(PInvoice,OrderNo)" & _
        "Values('" & txtPInvoice & "','" & Orders(i) & "')"
    Next i
          
    con.CommitTrans
    
    Saved = True
    Exit Function

ERR:
    MsgBox ERR.Description
End Function
Private Sub FillCmbs()

    With cmbTradeTerms
        .Clear
        .AddItem "FOB"
        .AddItem "CIF (Air)"
        .AddItem "CIF (Sea)"
        .AddItem "C&F (Air)"
        .AddItem "C&F (Sea)"
        .ListIndex = 0
    End With
    
    With cmbPartial
        .AddItem "Allowed"
        .AddItem "Not Allowed"
    End With
    
    With cmbTrans
        .AddItem "Allowed"
        .AddItem "Not Allowed"
    End With
    
    With cmbPayment
        .AddItem "CAD"
        .AddItem "30 DAYS AT SIGHT"
        .AddItem "90 DAYS AT SIGHT"
        .AddItem "60 DAYS AT SIGHT"
        .AddItem "Letter Of Credit"
        .AddItem "Advance Cheque"
        .AddItem "By Cheque"
    End With
    
    With cmbSMethod
        .AddItem "By Air"
        .AddItem "By Sea"
    End With
    
    cmbLoadCountry.AddItem "Pakistan"
    cmbLoadCountry.ListIndex = 0
    
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

    On Error GoTo ERR
    
    Dim strTTerms As String
    
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
        .Parameters("@CustCode").value = cmbCust.Text
        .Parameters("@Country").value = cmbCountry.Text
        Set rs = .Execute
    End With
    
    With rs
        cmbOrders.Clear
        Do Until .EOF
            cmbOrders.AddItem .Fields(0) & ""
            .MoveNext
        Loop
        .Close
    End With
    
    Exit Sub
ERR:
    MsgBox ERR.Description
    
End Sub


