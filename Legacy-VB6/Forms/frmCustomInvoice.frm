VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmCustomInvoice 
   ClientHeight    =   9495
   ClientLeft      =   1950
   ClientTop       =   435
   ClientWidth     =   12150
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
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   12150
   WindowState     =   2  'Maximized
   Begin VB.Frame FAll 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9420
      Left            =   45
      TabIndex        =   7
      Top             =   30
      Width           =   12120
      Begin ComboList.Usercontrol1 cmbCompany 
         Height          =   285
         Left            =   1740
         TabIndex        =   0
         Top             =   810
         Width           =   8055
         _ExtentX        =   14208
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
      Begin VB.TextBox txtCartons 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   10500
         MaxLength       =   150
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1470
         Width           =   825
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   285
         TabIndex        =   1
         Top             =   1455
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
      Begin VB.TextBox txtCurrency 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3810
         MaxLength       =   150
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1470
         Width           =   1005
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   0
         TabIndex        =   9
         Top             =   15
         Width           =   12120
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Custom Invoice"
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
            TabIndex        =   10
            Top             =   150
            Width           =   12000
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Custom Invoice"
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
            TabIndex        =   11
            Top             =   165
            Width           =   11985
         End
      End
      Begin MSComCtl2.DTPicker OrderDT 
         Height          =   300
         Left            =   6795
         TabIndex        =   4
         Top             =   1455
         Width           =   1755
         _ExtentX        =   3096
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
         Format          =   11730947
         CurrentDate     =   37250
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5940
         Left            =   135
         TabIndex        =   18
         Top             =   1785
         Width           =   11850
         _ExtentX        =   20902
         _ExtentY        =   10478
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         TabCaption(0)   =   "Order Detail"
         TabPicture(0)   =   "frmCustomInvoice.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "LV"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtEditRate"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Other Info."
         TabPicture(1)   =   "frmCustomInvoice.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame5"
         Tab(1).Control(1)=   "Frame3"
         Tab(1).Control(2)=   "Frame4"
         Tab(1).Control(3)=   "Frame2"
         Tab(1).ControlCount=   4
         Begin VB.Frame Frame5 
            Height          =   3870
            Left            =   -74895
            TabIndex        =   54
            Top             =   1965
            Width           =   11670
            Begin VB.TextBox txtVessel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   8775
               TabIndex        =   55
               Top             =   525
               Width           =   2820
            End
            Begin ComboList.Usercontrol1 cmbInsurAgent 
               Height          =   285
               Left            =   2955
               TabIndex        =   56
               Top             =   525
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
               Left            =   60
               TabIndex        =   65
               Top             =   525
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
            Begin MSForms.TextBox txtCommercialDeclarations 
               Height          =   1005
               Left            =   60
               TabIndex        =   68
               Top             =   2220
               Width           =   5730
               VariousPropertyBits=   -1400879077
               BorderStyle     =   1
               Size            =   "10107;1773"
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
               Caption         =   "Commercial Declarations"
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
               TabIndex        =   67
               Top             =   1980
               Width           =   5730
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
               Index           =   14
               Left            =   60
               TabIndex        =   66
               Top             =   285
               Width           =   2865
            End
            Begin MSForms.TextBox txtBankAddress 
               Height          =   2100
               Left            =   6060
               TabIndex        =   64
               Top             =   1140
               Width           =   5535
               VariousPropertyBits=   -1400879077
               BorderStyle     =   1
               Size            =   "9763;3704"
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
               Caption         =   "Bank Address"
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
               Left            =   6060
               TabIndex        =   63
               Top             =   900
               Width           =   5535
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Custom Declarations"
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
               Left            =   60
               TabIndex        =   62
               Top             =   900
               Width           =   5730
            End
            Begin MSForms.TextBox txtDeclarations 
               Height          =   855
               Left            =   60
               TabIndex        =   61
               Top             =   1140
               Width           =   5730
               VariousPropertyBits=   -1400879077
               BorderStyle     =   1
               Size            =   "10107;1508"
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
               Left            =   8775
               TabIndex        =   60
               Top             =   285
               Width           =   2820
            End
            Begin VB.Label Label1 
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
               Index           =   16
               Left            =   6060
               TabIndex        =   59
               Top             =   285
               Width           =   2655
            End
            Begin MSForms.ComboBox cmbThrough 
               Height          =   285
               Left            =   6060
               TabIndex        =   58
               Top             =   525
               Width           =   2655
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "4683;503"
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
               TabIndex        =   57
               Top             =   285
               Width           =   2835
            End
         End
         Begin VB.TextBox txtEditRate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4410
            TabIndex        =   53
            Top             =   1620
            Visible         =   0   'False
            Width           =   1230
         End
         Begin VB.Frame Frame3 
            Height          =   1590
            Left            =   -74895
            TabIndex        =   30
            Top             =   345
            Width           =   5805
            Begin VB.TextBox txtGWeight 
               Appearance      =   0  'Flat
               DataField       =   "RptTime"
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   3900
               MaxLength       =   150
               TabIndex        =   69
               Top             =   450
               Width           =   1830
            End
            Begin VB.TextBox txtFormE 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   75
               TabIndex        =   39
               Top             =   1125
               Width           =   2700
            End
            Begin VB.TextBox txtAWBNo 
               Appearance      =   0  'Flat
               DataField       =   "RptTime"
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   75
               MaxLength       =   150
               TabIndex        =   33
               Top             =   450
               Width           =   1830
            End
            Begin MSComCtl2.DTPicker AWBDT 
               Height          =   300
               Left            =   1935
               TabIndex        =   35
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
               Format          =   11730947
               CurrentDate     =   37250
            End
            Begin MSComCtl2.DTPicker FormEDT 
               Height          =   300
               Left            =   2820
               TabIndex        =   51
               Top             =   1125
               Width           =   2895
               _ExtentX        =   5106
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
               Format          =   11730947
               CurrentDate     =   37250
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
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
               TabIndex        =   70
               Top             =   195
               Width           =   1830
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
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
               Left            =   2835
               TabIndex        =   50
               Top             =   870
               Width           =   2895
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
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
               Index           =   0
               Left            =   75
               TabIndex        =   49
               Top             =   870
               Width           =   2700
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
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
               TabIndex        =   34
               Top             =   195
               Width           =   1935
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
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
               TabIndex        =   32
               Top             =   195
               Width           =   1830
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
            Height          =   795
            Left            =   -68865
            TabIndex        =   25
            Top             =   1140
            Width           =   5625
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Country                   City                          Port"
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
               TabIndex        =   29
               Top             =   210
               Width           =   5475
            End
            Begin MSForms.ComboBox cmbDisCountry 
               Height          =   285
               Left            =   75
               TabIndex        =   28
               Top             =   450
               Width           =   1815
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "3201;503"
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
               TabIndex        =   27
               Top             =   450
               Width           =   1815
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "3201;503"
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
               Left            =   3735
               TabIndex        =   26
               Top             =   450
               Width           =   1815
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "3201;503"
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
            Height          =   795
            Left            =   -68865
            TabIndex        =   20
            Top             =   345
            Width           =   5625
            Begin MSForms.ComboBox cmbLoadPort 
               Height          =   285
               Left            =   3735
               TabIndex        =   24
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
               TabIndex        =   23
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
            Begin MSForms.ComboBox cmbLoadCountry 
               Height          =   285
               Left            =   75
               TabIndex        =   22
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
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Country                  City                           Port"
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
               TabIndex        =   21
               Top             =   210
               Width           =   5460
            End
         End
         Begin MSComctlLib.ListView LV 
            Height          =   5355
            Left            =   300
            TabIndex        =   19
            Top             =   420
            Width           =   11070
            _ExtentX        =   19526
            _ExtentY        =   9446
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
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S No."
               Object.Width           =   1014
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Order #"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Item ID"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Description"
               Object.Width           =   6879
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
               SubItemIndex    =   6
               Text            =   "Ava. Qty"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "Amount"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Text            =   "Amount"
               Object.Width           =   0
            EndProperty
         End
      End
      Begin VB.Frame FraAdd 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   5085
         TabIndex        =   40
         Top             =   8700
         Width           =   6945
         Begin MSForms.CommandButton cmdPrint 
            Height          =   360
            Left            =   3480
            TabIndex        =   44
            Top             =   180
            Width           =   1665
            Caption         =   "Print             "
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
            Left            =   5175
            TabIndex        =   43
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
            TabIndex        =   42
            Top             =   180
            Width           =   1665
            Caption         =   "Add New     "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmCustomInvoice.frx":0038
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
            Left            =   1800
            TabIndex        =   41
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   8460
         TabIndex        =   45
         Top             =   8700
         Visible         =   0   'False
         Width           =   3555
         Begin MSForms.CommandButton cmdCancel 
            Height          =   360
            Left            =   1800
            TabIndex        =   47
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
            TabIndex        =   46
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
         Left            =   285
         TabIndex        =   73
         Top             =   825
         Width           =   1410
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
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
         Index           =   21
         Left            =   9495
         TabIndex        =   72
         Top             =   7800
         Width           =   1005
      End
      Begin VB.Label lblTotPcs 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   10485
         TabIndex        =   71
         Top             =   7800
         Width           =   1455
      End
      Begin MSForms.ComboBox cmbInvNoNew 
         Height          =   285
         Left            =   4845
         TabIndex        =   48
         Tag             =   "AutoCatID"
         Top             =   1470
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
      Begin MSForms.ComboBox cmbSMethod 
         Height          =   285
         Left            =   3120
         TabIndex        =   38
         Top             =   8400
         Width           =   2925
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5159;503"
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
         Left            =   3120
         TabIndex        =   37
         Top             =   8160
         Width           =   2910
      End
      Begin MSForms.ComboBox cmbInvNo 
         Height          =   285
         Left            =   4845
         TabIndex        =   36
         Top             =   1470
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
      Begin MSForms.TextBox lblTotWords 
         Height          =   270
         Left            =   2100
         TabIndex        =   31
         Top             =   7800
         Width           =   5055
         VariousPropertyBits=   746604575
         BackColor       =   15199215
         BorderStyle     =   1
         Size            =   "8916;476"
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
         Left            =   7140
         TabIndex        =   17
         Top             =   7800
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
         Left            =   285
         TabIndex        =   16
         Top             =   7800
         Width           =   1830
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   8010
         TabIndex        =   15
         Top             =   7800
         Width           =   1500
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
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
         Index           =   3
         Left            =   285
         TabIndex        =   14
         Top             =   8160
         Width           =   2805
      End
      Begin MSForms.CheckBox chkPrint 
         Height          =   345
         Left            =   270
         TabIndex        =   13
         Top             =   8730
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
         Left            =   285
         TabIndex        =   6
         Top             =   8400
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
      Begin MSForms.CheckBox chkCust 
         Height          =   360
         Left            =   270
         TabIndex        =   12
         Top             =   8985
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
         Left            =   8565
         TabIndex        =   5
         Top             =   1470
         Width           =   1920
         VariousPropertyBits=   746604571
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
         Caption         =   $"frmCustomInvoice.frx":014A
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
         TabIndex        =   8
         Top             =   1215
         Width           =   11040
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   2040
         TabIndex        =   2
         Top             =   1470
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
Attribute VB_Name = "frmCustomInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PInvoice As String

Private Sub cmbCompany_matched()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()
    
    On Error GoTo err
    If cmbCountry.MatchFound = False Then Exit Sub
    If cmbCompany.MatchFound = False Then
        MsgBox "Please Select Company.", vbInformation
        Exit Sub
    End If
    'cmbCust.Tag = cmbCust.Text & cmbCountry
    cmbCust.Tag = cmbCust.Text & cmbCountry
    txtCurrency = cmbCountry.List(cmbCountry.ListIndex, 1)
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select CustomInvoice,InvoiceL From VCustomInvoice1 Where CompanyRefID=" & cmbCompany.ID & " AND CustCode+Country='" & cmbCust.ID & cmbCountry.Text & "' And InvoiceH=0 Order By CustomInvoice", con, adOpenForwardOnly, adLockReadOnly
        cmbInvNo.Clear
        Do Until .EOF
            cmbInvNo.AddItem .Fields(0) & ""
            cmbInvNo.List(cmbInvNo.ListCount - 1, 1) = .Fields(1) & ""
            .MoveNEXT
        Loop
        .Close
        
        .Open "Select PInvoice From FPInvoice Where PInvoice Not In (Select CustomInvoice From CustomInvoice) and CompanyRefID=" & cmbCompany.ID & " AND CustCode+Country='" & cmbCust.ID & cmbCountry.Text & "' Order By PInvoice", con, adOpenForwardOnly, adLockReadOnly
        cmbInvNoNew.Clear
        Do Until .EOF
            cmbInvNoNew.AddItem .Fields(0) & ""
            
            .MoveNEXT
        Loop
        .Close
        
        '.Open "Select BankName,Address1 From FCustBanks Where CustCode+Country='" & cmbCust.ID & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        'cmbThrough.Clear
        'Do Until .EOF
        '    cmbThrough.AddItem .Fields(0) & ""
        '    cmbThrough.List(cmbThrough.ListCount - 1, 1) = .Fields(1) & ""
        '
        '    .MoveNext
        'Loop
        '.Close
        
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
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbinvno_Change()
    Call cmbinvno_click
End Sub

Private Sub cmbinvno_click()

    If cmbInvNo.MatchFound = False Then Exit Sub
    Call ShowDetail(cmbInvNo.Text, "VrptCustomInvoice")
    
    Dim Arr() As Variant
    
    Arr = rsExecute("Select TotalCartons From VTotalCartons Where PInvoice='" & cmbInvNo.Text & "'")
    txtCartons = Arr(0) & ""
    
End Sub

Private Sub lok(Locked As Boolean)
    
    If Locked Then PInvoice = ""
    
    FraAdd.Visible = Locked
    FraSave.Visible = Not Locked
    
    cmbInvNo.Visible = Locked
    cmbInvNoNew.Visible = Not Locked
    cmbCompany.Enabled = Locked
    
    cmbInvNo.Locked = False
    
End Sub

Private Sub cmbInvNoNew_Change()
    Call cmbInvNoNew_Click
End Sub

Private Sub cmbInvNoNew_Click()
    
    
    If cmbInvNoNew.MatchFound = False Then Exit Sub
    Call ShowDetail(cmbInvNoNew.Text, "VrptCustomInvoice")
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Declarations,CommercialDeclarations From CustomInvoice Where Cast(SubString(CustomInvoice,6,4) As int)=(Select Max(Cast(SubString(CustomInvoice,6,4) As int)) From CustomInvoice)", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            txtDeclarations = ![Declarations] & ""
            txtCommercialDeclarations = ![CommercialDeclarations] & ""
        End If
        .Close
        .Open "Select Min(FormENo) From GeneratedFormENos Where Used=0", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            txtFormE = "BAL-" & .Fields(0) & ""
        End If
        
        .Close
    End With
    Set rs = Nothing
    
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
    
End Sub

Private Sub cmdAdd_Click()
    
    If cmbCompany.MatchFound = False Then
        MsgBox "Please Select Company.", vbInformation
        Exit Sub
    End If
    
    Call lok(False)
    PInvoice = ""
    
    AWBDT.Value = Date
    AWBDT.Value = ""
    FormEDT.Value = Date
    
    LV.ListItems.Clear
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdedit_Click()
    
    If cmbInvNo.MatchFound = False Then Exit Sub
     
    If cmbInvNo.List(cmbInvNo.ListIndex, 1) Then
        MsgBox "This Record Has Been Locked.", vbInformation
        Exit Sub
    End If
    cmbCompany.Enabled = False
    PInvoice = cmbInvNo
    FraSave.Visible = True
    FraAdd.Visible = False
    cmbInvNo.Locked = True
    
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
    
    If cmbInvNo.MatchFound = False Then Exit Sub
    
    'frmNewCustomPListM.Show
    
    Dim retStr As String
    Dim rpt As CRAXDDRT.Report 'New rptCustomInvoice
    Set rpt = rptApp.OpenReport(RptPath & "\rptCustomInvoice.rpt") '
    'retStr = ShowPopUpMenu(FAll.Left + FraAdd.Left + cmdPrint.Left + 300, FraAdd.Top + FAll.Top + cmdPrint.Top - 200, "Print On &Blank Page", "-", "Print On &Pre-Printed Page")
    
    'Select Case retStr
    '    Case "Print On &Blank Page"
    '        Set rpt = New rptCustomInvoice
    '    Case "Print On &Pre-Printed Page"
    '        Set rpt = New rptCustomInvoicePre
    '    Case Else
    '        Exit Sub
    'End Select
    
    Screen.MousePointer = vbHourglass
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VrptProformas.PInvoice}='" & cmbInvNo.Text & "' ", rpt
    'frmPrevRpt.ShowReport
    Screen.MousePointer = vbDefault
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdSave_Click()

    On Error GoTo err
    Call CalculateTotals
    
    Dim TempInvNo As String
    Call StartTrans(con)
    
    If PInvoice <> "" Then
        con.Execute "Delete From CustomInvoice Where CustomInvoice='" & PInvoice & "'"
        TempInvNo = PInvoice
    Else
        TempInvNo = cmbInvNoNew
        
        con.Execute "Update GeneratedFormENos Set Used=1 Where FormENo=" & Val(txtFormE)
    End If
    
    con.Execute "Insert Into CustomInvoice(CustomInvoice,DT,ShippingMode,AWBNo,AWBNoDT" & _
     ",FormE,FormEDT,TradeTerms,Vessel,ComInvDT,TotalCustomAmt,TotalCustomAmtToRcv" & _
     ",TotalCommercialAmt,TotalCommercialAmtToRcv,Declarations,CommercialDeclarations,GrossWeight) Values('" & _
     TempInvNo & "','" & OrderDT & "','" & cmbSMethod & "','" & txtAWBNo & "','" & _
     AWBDT & "','" & txtFormE & "','" & FormEDT & "','" & _
     cmbTradeTerms & "','" & txtVessel & "','" & ComInvDT & "'," & Val(lblTotal.Caption) & _
     "," & Val(lblTotal.Caption) & "," & Val(lblTotal.Tag) & "," & Val(lblTotal.Tag) & ",'" & _
     txtDeclarations & "','" & txtCommercialDeclarations & "'," & Val(txtGWeight) & ")"
    
    
    'Also Add This Entry.
    
    If PInvoice = "" Then con.Execute "Insert Into ExportPackageLocknHide(CustomInvoice) Values('" & TempInvNo & "')"
    
    con.CommitTrans
    
    Call lok(True)
    Call cmbCountry_Click
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub cmdCancel_Click()
    Call cmbinvno_click
    Call lok(True)
End Sub

Private Sub Form_Load()
   
    On Error GoTo err
    
   
    SSTab1.Tab = 0
    
    Call SaveLV(LV)
    
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where custCode<>'Hico'")
    OrderDT = Date
    Me.KeyPreview = True
    cmbCust.ListHeight = 2500
    
    
    cmbShipAgent.ListHeight = 2000
    cmbInsurAgent.ListHeight = 2000
    
    Call FillCmbs
    
    
    Exit Sub
err:
    MsgBox err.Description
   
End Sub

Private Sub Form_Resize()
On Error Resume Next
   FAll.Move ((Me.ScaleWidth - FAll.Width) / 2), ((Me.ScaleHeight - FAll.Height) / 2)
End Sub

Private Sub CalculateTotals()

    Dim DCustomTotal As Double
    Dim DCommTotal As Double
    Dim LTotPcs As Long
    With LV
        For i = 1 To .ListItems.count Step 2
            'dQty = dQty + Val(.TextMatrix(i, 4))
            DCustomTotal = DCustomTotal + Val(.ListItems(i).SubItems(7) & "")
            DCommTotal = DCommTotal + (Val(.ListItems(i).ListSubItems(7).Tag & "") * Val(.ListItems(i).ListSubItems(6) & ""))
            LTotPcs = LTotPcs + Val(.ListItems(i).SubItems(6) & "")
        Next i
    End With
    'txttQty = dQty
    lblTotal = DCustomTotal
    lblTotal.Tag = DCommTotal
    lblTotWords = ConvertInWords(CCur(DCustomTotal), txtCurrency.Text)
    
    lblTotPcs = LTotPcs
    
End Sub

Private Sub FillCmbs()

    With cmbTradeTerms
        .Clear
        .AddItem "FOB"
        .List(0, 1) = "FOB"
        .AddItem "CIF (Air)"
        .List(1, 1) = "CIFAir"
        .AddItem "CIF (Sea)"
        .List(2, 1) = "CIFSea"
        .AddItem "C&F (Air)"
        .List(3, 1) = "CnFAir"
        .AddItem "C&F (Sea)"
        .List(4, 1) = "CnFSea"
        .ListIndex = 0
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
    
    cmbShipAgent.ClearVals
    Call cmbShipAgent.AddVals(con, "Name", "ShippingAgents", "EntryID", " Where Active=1")
    
    cmbInsurAgent.ClearVals
    Call cmbInsurAgent.AddVals(con, "Name", "InsuranceAgents", "EntryID", " Where Active=1")
    
    cmbCompany.ListHeight = 6000
    cmbCompany.AddVals con, "CompanyName", "Companies", "EntryID", " Order By CompanyName"
        
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


 
Private Sub GetOrders()

    On Error GoTo err
    
    Dim strTTerms As String
    
    With cmbCountry
        If Not .MatchFound Then Exit Sub
    End With
    If cmbProformas.MatchFound = False Then Exit Sub
    
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "SP_FOrdersForInvoice"
        Set .ActiveConnection = con
        
        .Parameters("@Customer").Value = cmbCust.Tag
        
        
        
        
        .Parameters("@PInvoice").Value = cmbProformas.Text
        
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

Private Sub cmbThrough_Change()
    'Call cmbThrough_Click
End Sub

Private Sub cmbThrough_Click()
    'If cmbThrough.MatchFound = False Then Exit Sub
    'txtBankAddress = cmbThrough.List(cmbThrough.ListIndex, 2)
End Sub
 
Private Sub ShowDetail(strProforma As String, TableName As String)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    LV.Visible = False
    'Debug.Print "Main Entry Start : " & Now
    With rs
        .Open "Select * From " & TableName & " Where PInvoice='" & strProforma & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCust.Text = ![CustCode] & ""
        cmbCountry = ![Country] & ""
        'txtPInvoice = ![PInvoice] & ""
        
        cmbTradeTerms = ![TradeTerms] & ""
        
        cmbPayment = ![PaymentTerms] & ""
        
        'txtLC = ![LCNo] & ""
        
        'cmbTrans = ![TransShipment] & ""
        'cmbPartial = ![PartialShipment] & ""
        'txtPackaging = ![Packaging] & ""
        'DeliveryDT = ![DeliveryDT] & ""
        
        cmbSMethod = ![SMethod] & ""
        
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
        
        cmbThrough = ![BankName] & ""
        txtBankAddress = ![Address2] & ""
        
                
        If cmbInvNo.Visible Then
            
            OrderDT = ![CustomInvoiceDT] & ""
            txtAWBNo = ![AWBNo] & ""
            AWBDT = CheckForNullDT(![AWBNoDT] & "")
            
            txtFormE = ![FormE] & ""
            FormEDT = ![FormEDT] & ""
            
            'txtCartons = ![TotalCartons] & ""
            txtVessel = ![Vessel] & ""
            
            txtDeclarations = ![CustomDeclarations] & ""
            txtCommercialDeclarations = ![CommercialDeclarations] & ""
            
            'txtGWeight = ![GrossWeight] & ""
        End If
        
        
        .Close
        
        'Debug.Print "Detail Entry Starts : " & Now
        
        
        .Open "Select * From VrptCustomInvoiceDetail Where PInvoice='" & strProforma & "' Order By ProformaQty Desc", con, adOpenForwardOnly, adLockReadOnly
        'Debug.Print "RS Opened : " & Now
        LV.ListItems.Clear
        
        Do Until .EOF
            Set ITM = LV.ListItems.add(, Val(![EntryID] & "") & "'", .AbsolutePosition)
            ITM.Tag = ![CompITemID] & "" 'cmbItems.List(cmbItems.ListIndex, 2)
            
            ITM.ListSubItems.add , , ![OrderNo] '![CatID] & ""
            ITM.ListSubItems.add , , ![ITemID] & ""
            ITM.ListSubItems.add(, ![ItemName] & "", ![ItemName] & "").Tag = ![Description] & ""
            
            ITM.ListSubItems.add , , txtUnit
            
            'ITM.ListSubItems.Add(, , Format(![CustomPrice] & "", "#0.000")).Tag = Format(![Price] & "")
            ITM.ListSubItems.add , , Format(![CustomPrice] & "", "#0.000")
            
            If cmbInvNo.Visible Then
                ITM.ListSubItems(5).Tag = Val(![ID] & "")
            End If
            
            ITM.ListSubItems.add , , Val(![ProformaQty] & "")
            
            
            
            ITM.ListSubItems(6).Tag = Val(![ProformaQty] & "") + Val(![InvQty] & "")
            
            ITM.SubItems(7) = Val(![ProformaQty] & "") * Val(![CustomPrice] & "")
            'ITM.SubItems(8) = Val(![ProformaQty] & "") * Val(![CustomPrice] & "")
               
            ITM.ListSubItems(7).Tag = Val(![Price] & "")
            
            Set ITM = LV.ListItems.add
            ITM.SubItems(3) = "Batch # " & ![BatchNo] & "   Lot # " & ![BatchNo] & "/" & ![ItemCode] & ""
            
            .MoveNEXT
        Loop
        .Close
        
        
    End With
    Debug.Print "Detail Entry Ends : " & Now
    LV.Visible = True
    Call CalculateTotals

    Set rs = Nothing
    Me.Show
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub LV_Click()
    txtEditRate.Visible = False
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    If FraAdd.Visible Then Exit Sub
    If LV.SelectedItem.Text = "" Then Exit Sub
    If PInvoice = "" Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select TotalCustomAmt,TotalCustomAmtToRcv From CustomInvoice Where CustomInvoice='" & cmbInvNo.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            If Val(.Fields(0) & "") <> Val(.Fields(1) & "") Then
                MsgBox "Cannot Edit Price." & vbNewLine & "Amount Has Been Receive Againgst This Invoice.", vbInformation
                Exit Sub
            End If
        End If
        .Close
    End With
    Set rs = Nothing
        
    With LV
        txtEditRate.Move .Left + .ColumnHeaders(6).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(6).Width
    End With
    
    With txtEditRate
        .Text = LV.SelectedItem.SubItems(5) & ""
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        
        .SetFocus
    End With
    
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtEditRate.Visible = False
End Sub

Private Sub txtEditRate_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo err
    
    If KeyCode = 13 Then
    
        With txtEditRate
        
            If Val(.Text) <= 0 Then
                MsgBox "Invalid Rate.", vbInformation
                Exit Sub
            End If
            
            Call StartTrans(con)
            
            Dim RecAff As Integer
            'con.Execute "Update FOrderItems Set CustomPrice=" & Val(.Text) & " Where ID=" & Val(LV.SelectedItem.ListSubItems(5).Tag), RecAff
            con.Execute "Update FProformaOrders Set CustomPrice=" & Val(.Text) & " Where EntryID=" & Val(LV.SelectedItem.key), RecAff
            
            If cmbTradeTerms.MatchFound Then
                con.Execute "Update Items Set " & cmbTradeTerms.List(cmbTradeTerms.ListIndex, 1) & "=" & Val(.Text) & " Where ItemID='" & LV.SelectedItem.Tag & "'"
            End If
            
            
            con.Execute "Update CustomInvoice Set TotalCustomAmt=" & Val(lblTotal.Caption) & ",TotalCustomAmtToRcv=" & Val(lblTotal.Caption) & _
             " Where CustomInvoice='" & cmbInvNo.Text & "'"
             
            con.CommitTrans
            
            'If RecAff > 0 Then
            LV.SelectedItem.SubItems(5) = Format(Val(.Text), "0.000")
            LV.SelectedItem.SubItems(7) = Val(.Text) * Val(LV.SelectedItem.SubItems(6))
                
            'End If
            Call CalculateTotals
            
            .Visible = False
            
        End With
        
    ElseIf KeyCode = 27 Then
        txtEditRate.Visible = False
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub txtEditRate_LostFocus()
    txtEditRate.Visible = False
End Sub
