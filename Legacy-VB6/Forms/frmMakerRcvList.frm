VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmMakerRcvList 
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13935
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
   ScaleHeight     =   8880
   ScaleMode       =   0  'User
   ScaleWidth      =   12481.77
   WindowState     =   2  'Maximized
   Begin VB.Frame FraManualPTCNo 
      BackColor       =   &H00E7EBEF&
      Caption         =   "Manual PTC No. :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1620
      Left            =   4050
      TabIndex        =   55
      Top             =   6435
      Visible         =   0   'False
      Width           =   4395
      Begin VB.CommandButton cmdUpdateManualPTCNo 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1410
         TabIndex        =   61
         Top             =   1035
         Width           =   1395
      End
      Begin VB.CommandButton cmdCancelManualPTC 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2865
         TabIndex        =   60
         Top             =   1035
         Width           =   1395
      End
      Begin VB.TextBox txtManualPTCNo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   2205
         TabIndex        =   58
         Top             =   540
         Width           =   1980
      End
      Begin VB.TextBox txtPTCNo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   540
         Width           =   1980
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Manual PTC No. :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   2205
         TabIndex        =   59
         Top             =   345
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PTC No. :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   57
         Top             =   345
         Width           =   690
      End
   End
   Begin VB.Frame FraSplitLot 
      BackColor       =   &H00E7EBEF&
      Caption         =   "Change / Split Lot"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4425
      Left            =   1680
      TabIndex        =   18
      Top             =   3390
      Visible         =   0   'False
      Width           =   8835
      Begin VB.Frame Fra 
         BackColor       =   &H00E7EBEF&
         Height          =   3015
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   780
         Visible         =   0   'False
         Width           =   8565
         Begin VB.Frame FraTSF_Frgz 
            BackColor       =   &H00E7EBEF&
            Height          =   2820
            Left            =   90
            TabIndex        =   40
            Top             =   120
            Width           =   8385
            Begin ComboList.Usercontrol1 cmbStore 
               Height          =   285
               Left            =   60
               TabIndex        =   44
               Top             =   360
               Width           =   2115
               _ExtentX        =   3731
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
            Begin ComboList.Usercontrol1 cmbLocation 
               Height          =   285
               Left            =   2190
               TabIndex        =   43
               Top             =   375
               Width           =   5400
               _ExtentX        =   9525
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
            Begin VB.TextBox txtRemarks 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   60
               TabIndex        =   42
               Top             =   885
               Width           =   8235
            End
            Begin VB.TextBox txtSFQty 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   7620
               TabIndex        =   41
               Top             =   375
               Width           =   660
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Location"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   2235
               TabIndex        =   48
               Top             =   180
               Width           =   720
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Store"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   47
               Top             =   165
               Width           =   465
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Remarks"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   60
               TabIndex        =   46
               Top             =   675
               Width           =   765
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Qty"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   7650
               TabIndex        =   45
               Top             =   150
               Width           =   300
            End
         End
         Begin VB.Frame FraTSF_QEL 
            BackColor       =   &H00E7EBEF&
            Height          =   2820
            Left            =   90
            TabIndex        =   49
            Top             =   120
            Width           =   8385
            Begin VB.TextBox txtQty_QEL 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3015
               TabIndex        =   50
               Top             =   390
               Width           =   660
            End
            Begin MSForms.ComboBox cmbLocation_QEL 
               Height          =   285
               Left            =   60
               TabIndex        =   53
               Tag             =   "PaymentTerms"
               Top             =   390
               Width           =   2940
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "5186;503"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   178
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qty"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   7
               Left            =   3030
               TabIndex        =   52
               Top             =   165
               Width           =   300
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Location"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   4
               Left            =   75
               TabIndex        =   51
               Top             =   180
               Width           =   720
            End
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E7EBEF&
         Height          =   3015
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   780
         Visible         =   0   'False
         Width           =   8565
         Begin ComboList.Usercontrol1 cmbSplitItemID 
            Height          =   285
            Left            =   3780
            TabIndex        =   35
            Top             =   330
            Width           =   3870
            _ExtentX        =   6826
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
         Begin ComboList.Usercontrol1 cmbSplitOrderNo 
            Height          =   285
            Left            =   1260
            TabIndex        =   31
            Top             =   330
            Width           =   2505
            _ExtentX        =   4419
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
         Begin ComboList.Usercontrol1 cmbSplitCust 
            Height          =   285
            Left            =   0
            TabIndex        =   30
            Top             =   330
            Width           =   1245
            _ExtentX        =   2196
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
         Begin VB.TextBox txtSplitLotRemarks 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   870
            Left            =   15
            MultiLine       =   -1  'True
            TabIndex        =   76
            Top             =   840
            Width           =   8475
         End
         Begin VB.TextBox txtSplitQty 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7665
            TabIndex        =   33
            Top             =   330
            Width           =   825
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Remarks"
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
            Left            =   15
            TabIndex        =   75
            Top             =   600
            Width           =   8475
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"frmMakerRcvList.frx":0000
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
            Left            =   0
            TabIndex        =   32
            Top             =   90
            Width           =   8490
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E7EBEF&
         Height          =   3015
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   780
         Visible         =   0   'False
         Width           =   8565
         Begin ComboList.Usercontrol1 cmbCOItemID 
            Height          =   285
            Left            =   3510
            TabIndex        =   34
            Top             =   330
            Width           =   4980
            _ExtentX        =   8784
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
         Begin ComboList.Usercontrol1 cmbCOOrderNo 
            Height          =   285
            Left            =   1545
            TabIndex        =   28
            Top             =   330
            Width           =   1965
            _ExtentX        =   3466
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
         Begin ComboList.Usercontrol1 cmbCOCust 
            Height          =   285
            Left            =   0
            TabIndex        =   27
            Top             =   330
            Width           =   1545
            _ExtentX        =   2725
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
         Begin VB.CheckBox chkCOStockOrder 
            BackColor       =   &H00E7EBEF&
            Caption         =   "Transfer to Stock Order"
            Height          =   315
            Left            =   45
            TabIndex        =   66
            Top             =   690
            Width           =   2235
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Customer               Order No.                           Item"
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
            Left            =   0
            TabIndex        =   29
            Top             =   90
            Width           =   8490
         End
      End
      Begin VB.OptionButton Opts 
         BackColor       =   &H00E7EBEF&
         Caption         =   "Change Customer Order No."
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
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   24
         Top             =   420
         Width           =   2805
      End
      Begin VB.OptionButton Opts 
         BackColor       =   &H00E7EBEF&
         Caption         =   "Split Lot"
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
         Height          =   285
         Index           =   1
         Left            =   3525
         TabIndex        =   23
         Top             =   420
         Width           =   1320
      End
      Begin VB.OptionButton Opts 
         BackColor       =   &H00E7EBEF&
         Caption         =   "Transfer to S/F Stock"
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
         Height          =   285
         Index           =   2
         Left            =   5400
         TabIndex        =   22
         Top             =   420
         Width           =   2625
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Post"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5565
         TabIndex        =   21
         Top             =   3840
         Width           =   1515
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cance&l"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7155
         TabIndex        =   20
         Top             =   3840
         Width           =   1515
      End
   End
   Begin VB.Frame FAcc 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   8895
      Left            =   60
      TabIndex        =   0
      Top             =   -45
      Width           =   13830
      Begin VB.TextBox txtMasterPONo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   75
         TabIndex        =   80
         Top             =   2070
         Width           =   960
      End
      Begin VB.CheckBox chkRegularLotsOnly 
         Caption         =   "Regular Lots Only"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11505
         TabIndex        =   77
         Top             =   1185
         Width           =   2145
      End
      Begin VB.CheckBox chkNonBilled 
         Caption         =   "Non-Billed Entries"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   9525
         TabIndex        =   74
         Top             =   1665
         Width           =   1935
      End
      Begin VB.CheckBox chkBilled 
         Caption         =   "Billed Entries"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   7830
         TabIndex        =   73
         Top             =   1665
         Width           =   1665
      End
      Begin VB.CheckBox chkRcvdButNotIssued 
         Caption         =   "Received but not Issued"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   11505
         TabIndex        =   72
         Top             =   1635
         Width           =   1785
      End
      Begin ComboList.Usercontrol1 cmbCustomer 
         Height          =   285
         Left            =   9840
         TabIndex        =   67
         Top             =   2070
         Width           =   3030
         _ExtentX        =   5345
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
      Begin ComboList.Usercontrol1 cmbMaker 
         Height          =   285
         Left            =   75
         TabIndex        =   9
         Top             =   840
         Width           =   4335
         _ExtentX        =   7646
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
      Begin ComboList.Usercontrol1 cmbProcess 
         Height          =   285
         Left            =   4410
         TabIndex        =   54
         Top             =   840
         Width           =   4560
         _ExtentX        =   8043
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
      Begin ComboList.Usercontrol1 cmbItemCat 
         Height          =   285
         Left            =   75
         TabIndex        =   36
         Top             =   1350
         Width           =   2325
         _ExtentX        =   4101
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
      Begin ComboList.Usercontrol1 cmbItemGroup 
         Height          =   285
         Left            =   2415
         TabIndex        =   65
         Top             =   1350
         Width           =   2490
         _ExtentX        =   4392
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
      Begin ComboList.Usercontrol1 cmbItems 
         Height          =   285
         Left            =   1050
         TabIndex        =   4
         Top             =   2070
         Width           =   8775
         _ExtentX        =   15478
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
      Begin VB.CheckBox chkRepair 
         Caption         =   "Repair Lots"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7830
         TabIndex        =   71
         Top             =   1425
         Width           =   1665
      End
      Begin VB.CheckBox chkRejection 
         Caption         =   "Rejection Lots"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9525
         TabIndex        =   70
         Top             =   1410
         Width           =   2190
      End
      Begin VB.CheckBox chkShowLastOnly 
         Caption         =   "Last Process Done"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9525
         TabIndex        =   64
         Top             =   1185
         Width           =   1920
      End
      Begin VB.TextBox txtOrderNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         TabIndex        =   63
         Top             =   1350
         Width           =   1050
      End
      Begin VB.CheckBox chkShowMasterPOOnly 
         Caption         =   "Master PO Only"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7830
         TabIndex        =   62
         Top             =   1200
         Width           =   1665
      End
      Begin VB.Frame fTitle 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Width           =   13815
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Receiving List"
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
            TabIndex        =   38
            Top             =   135
            Width           =   13785
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Receiving List"
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
            TabIndex        =   39
            Top             =   150
            Width           =   13785
         End
      End
      Begin VB.TextBox txtInActiveDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   6825
         TabIndex        =   17
         Top             =   1350
         Width           =   960
      End
      Begin VB.TextBox txtLotNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4920
         TabIndex        =   5
         Top             =   1350
         Width           =   825
      End
      Begin VB.Frame Frame1 
         Height          =   480
         Left            =   2910
         TabIndex        =   12
         Top             =   8295
         Width           =   5475
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   1965
            TabIndex        =   15
            Top             =   135
            Width           =   2430
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4395
            TabIndex        =   14
            Top             =   135
            Width           =   1020
         End
         Begin VB.ComboBox cmbfield 
            Height          =   315
            ItemData        =   "frmMakerRcvList.frx":009E
            Left            =   60
            List            =   "frmMakerRcvList.frx":00B7
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   135
            Width           =   1875
         End
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
         ItemData        =   "frmMakerRcvList.frx":00FE
         Left            =   8985
         List            =   "frmMakerRcvList.frx":0114
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   840
         Width           =   1770
      End
      Begin VB.CommandButton cmdPrintGP 
         Caption         =   "Print Rejection Gate-Pass"
         Height          =   345
         Left            =   8490
         TabIndex        =   3
         Top             =   8400
         Width           =   2130
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5835
         Left            =   60
         TabIndex        =   1
         Top             =   2385
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   10292
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
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
            Text            =   "S. #"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Maker"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "PO #"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Process"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Item"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Rcvd Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Rej. Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Re-Work"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Status"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "Lot #"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Order No."
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   12270
         TabIndex        =   7
         Top             =   840
         Width           =   1515
         _ExtentX        =   2672
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
         Format          =   151715843
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   10755
         TabIndex        =   8
         Top             =   825
         Width           =   1515
         _ExtentX        =   2672
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
         Format          =   151715843
         CurrentDate     =   37055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Master PO"
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
         Index           =   8
         Left            =   75
         TabIndex        =   81
         Top             =   1860
         Width           =   870
      End
      Begin MSForms.CommandButton cmdCheckAll 
         Height          =   375
         Index           =   0
         Left            =   60
         TabIndex        =   79
         Top             =   8325
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Check All"
         PicturePosition =   327683
         Size            =   "2355;661"
         Accelerator     =   104
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdCheckAll 
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   78
         Top             =   8325
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Un-Check All"
         PicturePosition =   327683
         Size            =   "2355;661"
         Accelerator     =   85
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Customer"
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
         Index           =   6
         Left            =   9840
         TabIndex        =   69
         Top             =   1875
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Article"
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
         Index           =   5
         Left            =   1050
         TabIndex        =   68
         Top             =   1860
         Width           =   555
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMakerRcvList.frx":0164
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
         Left            =   45
         TabIndex        =   16
         Top             =   1140
         Width           =   7710
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMakerRcvList.frx":01EB
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
         Left            =   45
         TabIndex        =   11
         Top             =   615
         Width           =   11130
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   345
         Left            =   12915
         TabIndex        =   10
         Top             =   2010
         Width           =   840
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "1482;609"
         Accelerator     =   82
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
         Left            =   12420
         TabIndex        =   2
         Top             =   8340
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
            Picture         =   "frmMakerRcvList.frx":02E5
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakerRcvList.frx":04A1
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PopMenu"
      Begin VB.Menu mnuPrintPurchaseOrder 
         Caption         =   "Print Slip"
      End
      Begin VB.Menu mnuPrintSlipMini 
         Caption         =   "Print Slip (Mini)"
      End
      Begin VB.Menu mnuPrintPTC 
         Caption         =   "Print PTC"
      End
      Begin VB.Menu mnuPrintPTCMini 
         Caption         =   "Print PTC Mini"
      End
      Begin VB.Menu mnuPrintPTCWithPrice 
         Caption         =   "Print PTC (Cost)"
      End
      Begin VB.Menu mnuDash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintRepairSlip 
         Caption         =   "Print Repair Slip"
      End
      Begin VB.Menu mnuDash4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteRcving 
         Caption         =   "Delete Receiving"
      End
      Begin VB.Menu mnuPrintDeletedLots 
         Caption         =   "Print Deleted Lots"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDashClose 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCloseLot 
         Caption         =   "Close Lot"
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSplitLot 
         Caption         =   "Split / Transfer Lot"
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuManualPTCNo 
         Caption         =   "Enter Manual PTC No."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuManualPTCNoDash 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintList 
         Caption         =   "Print List"
      End
      Begin VB.Menu mnuPrintProcessList 
         Caption         =   "Print Process List"
      End
      Begin VB.Menu mnuPrintListSummary 
         Caption         =   "Print List (Summary)"
      End
      Begin VB.Menu mnuRRReport 
         Caption         =   "Re-work/Rejection Report"
      End
      Begin VB.Menu mnuInspectionData 
         Caption         =   "Inspection Data"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintReWorkRejectionChart 
         Caption         =   "Re-wok/Rejection Chart"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuProcessInspection 
         Caption         =   "Process Inspection"
      End
      Begin VB.Menu mnuProcessInspectionReport 
         Caption         =   "Process Inspection Report"
      End
      Begin VB.Menu mnuPrintDashQCReport 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintQCReport 
         Caption         =   "Print QC Report"
      End
   End
End
Attribute VB_Name = "frmMakerRcvList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCourierInvoice As Boolean
Dim bShipping As Boolean
Dim bCustomInvoice As Boolean
Dim strReportSQL As String
Dim bShowForSplitTransferOnly As Boolean, strLotNo As String, lEntryID_Split As Long

Public Sub ShowMe(Optional p_bCustomInvoice As Boolean = True)
    
End Sub

Public Sub ShowForOrder(p_strOrderNo As String)

    txtOrderNo = p_strOrderNo
    DtFrom.Value = #1/1/2000#
    DtTo.Value = Date
    
    Call RefreshLV
    
    Me.Show
End Sub

Public Sub ShowForShipping()
    bShipping = True
    Me.Show
End Sub

Private Sub cmbCOCust_matched()

    cmbCOOrderNo.ClearVals
    'cmbCOOrderNo.AddVals con, "DISTINCT FOrderItems.OrderNo", "FOrderItems INNER JOIN FCustomerOrders ON FCustomerOrders.OrderNo=FOrderItems.OrderNo", "FCustomerOrders.OrderNo", "WHERE CustCode='" & cmbCOCust.ID & "' AND CompItemCode='" & LV.SelectedItem.ListSubItems(2).Tag & "'"
    cmbCOOrderNo.AddVals con, "DISTINCT FOrderItems.OrderNo", "FOrderItems INNER JOIN FCustomerOrders ON FCustomerOrders.OrderNo=FOrderItems.OrderNo", "FCustomerOrders.OrderNo", "WHERE CustCode='" & cmbCOCust.ID & "' AND FCustomerOrders.OrderNo IN(SELECT OrderNo FROM VUnshippedOrderList)"
    
    
End Sub

Private Sub cmbCOOrderNo_matched()

    cmbCOItemID.ClearVals
    If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
        'cmbCOItemID.AddVals con, "'{' + VFOrderItems.CompItemCode + '} ' + VFOrderItems.ItemName", "VFOrderItems INNER JOIN Items ON VFOrderItems.CompItemCode=Items.ItemID", "Items.ItemID", " WHERE OrderNo='" & cmbCOOrderNo.Text & "'"
        cmbCOItemID.AddVals con, "'{' + VrptOrders_ForProduction.CompItemID + '} ' + VrptOrders_ForProduction.ItemName", "VrptOrders_ForProduction INNER JOIN Items ON VrptOrders_ForProduction.CompItemID=Items.ItemID", "Items.ItemID", " WHERE OrderNo='" & cmbCOOrderNo.Text & "'"
    Else
        cmbCOItemID.AddVals con, "'{' + CompItemID + '} ' + ItemName", "VrptOrders", "CompItemID", " WHERE OrderNo='" & cmbCOOrderNo.Text & "'"
    End If
    
End Sub

Private Sub cmbLocation_matched()
    txtRemarks = GetSingleStringValue("Remarks", "StockOrderOpening", " WHERE ItemID='" & LV.SelectedItem.ListSubItems(2).Tag & "' AND ProcessID=" & LV.SelectedItem.ListSubItems(4).Tag & " AND Shelf_RefID=" & cmbLocation.ID)
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

Private Sub cmbSplitCust_matched()

    cmbSplitOrderNo.ClearVals
    cmbSplitOrderNo.AddVals con, "DISTINCT FOrderItems.OrderNo+' {'+InternalRefNo+'}'", "FOrderItems INNER JOIN FCustomerOrders ON FCustomerOrders.OrderNo=FOrderItems.OrderNo", "FCustomerOrders.OrderNo", "WHERE CustCode='" & cmbSplitCust.ID & "' AND FOrderItems.CompItemCode='" & LV.SelectedItem.ListSubItems(2).Tag & "' AND (FCustomerOrders.OrderNo IN(SELECT OrderNo FROM VUnshippedOrderList) OR FOrderItems.OrderNo='Stock-Order')"

End Sub

Private Sub cmbSplitOrderNo_matched()

    cmbSplitItemID.ClearVals
    If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
        'cmbSplitItemID.AddVals con, "'{' + VFOrderItems.CompItemCode + '} ' + VFOrderItems.ItemName", "VFOrderItems INNER JOIN Items ON VFOrderItems.CompItemCode=Items.ItemID", "Items.ItemID", " WHERE OrderNo='" & cmbSplitOrderNo.Text & "'"
        cmbSplitItemID.AddVals con, "'{' + VrptOrders_ForProduction.CompItemID + '} ' + VrptOrders_ForProduction.ItemName", "VrptOrders_ForProduction INNER JOIN Items ON VrptOrders_ForProduction.CompItemID=Items.ItemID", "Items.ItemID", " WHERE OrderNo='" & cmbSplitOrderNo.ID & "'"
    Else
        cmbSplitItemID.AddVals con, "'{' + CompItemID + '} ' + ItemName", "VrptOrders", "CompItemID", " WHERE OrderNo='" & cmbSplitOrderNo.ID & "'"
    End If
    
End Sub

Private Sub cmbStore_matched()
    
    cmbLocation.ClearVals
    cmbLocation.AddVals con, "RackNo + ' {' + ShelfNo + '}'", "VStoreShelfs", "EntryID", " WHERE Store_RefID=" & cmbStore.ID & " ORDER BY StoreName,RackNo,ShelfNo"
    
End Sub

Private Sub cmdCancel_Click()
    FraSplitLot.Visible = False
    FAcc.Enabled = True
    If bShowForSplitTransferOnly Then
        Unload Me
    End If
End Sub

Private Sub cmdCancelManualPTC_Click()
    FraManualPTCNo.Visible = False
    FAcc.Enabled = True
End Sub

Private Sub cmdCheckAll_Click(Index As Integer)

    Dim i As Integer
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = (Index = 0)
    Next

End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdCustomer_Click()

    If LVCustomer.Visible Then
    
        LVCustomer.Visible = False
        Dim i As Integer
        If LVCustomer.ListItems(1).Checked Then
            txtCustomer.Text = "<All Customers>"
            txtCustomer.Tag = "0"
        Else
            Dim strCustomers As String, strCustomerIDs As String
            For i = 1 To LVCustomer.ListItems.count
                If LVCustomer.ListItems(i).Checked Then
                    strCustomers = strCustomers & LVCustomer.ListItems(i).Text & ","
                    strCustomerIDs = strCustomerIDs & "'" & LVCustomer.ListItems(i).Tag & "',"
                End If
            Next
            If strCustomers <> "" Then
                strCustomers = Left(strCustomers, Len(strCustomers) - 1)
                strCustomerIDs = Left(strCustomerIDs, Len(strCustomerIDs) - 1)
            End If
            txtCustomer.Text = strCustomers
            txtCustomer.Tag = strCustomerIDs
        End If
        
    Else
        LVCustomer.Visible = True
    End If
    
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




    For i = StartIdx To LV.ListItems.count
       
       
       If cmbfield.ListIndex = 0 Then
          fldValue = Trim(Replace(LV.ListItems(i), "»", ""))
       Else
          fldValue = Trim(Replace(LV.ListItems(i).SubItems(cmbfield.ListIndex), "»", ""))
       End If
    
       If LCase(fldValue) Like "*" & LCase(Trim(txtFind)) & "*" Then
          LV.ListItems(i).Selected = True
          LV.ListItems(i).EnsureVisible
          If i <> StartIdx Then Exit For
       End If
    Next
    
    If i = LV.ListItems.count + 1 Then
        MsgBox "No Mathes Found.", vbInformation
        Set LV.SelectedItem = Nothing
    End If
End Sub

Private Sub cmdGroups_Click()
    
    If LVGroups.Visible Then
    
        LVGroups.Visible = False
        Dim i As Integer
        If LVGroups.ListItems(1).Checked Then
            txtGroups.Text = "<All Groups>"
            txtGroups.Tag = "0"
        Else
            Dim strGroups As String, strGroupIDs As String
            For i = 1 To LVGroups.ListItems.count
                If LVGroups.ListItems(i).Checked Then
                    strGroups = strGroups & LVGroups.ListItems(i).Text & ","
                    strGroupIDs = strGroupIDs & "'" & LVGroups.ListItems(i).Tag & "',"
                End If
            Next
            If strGroups <> "" Then
                strGroups = Left(strGroups, Len(strGroups) - 1)
                strGroupIDs = Left(strGroupIDs, Len(strGroupIDs) - 1)
            End If
            txtGroups.Text = strGroups
            txtGroups.Tag = strGroupIDs
        End If
        
    Else
        LVGroups.Visible = True
    End If
    

End Sub

Private Sub cmdPrintGP_Click()

    On Error GoTo err
    
    Dim strList As String
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                strList = strList & .Tag & ","
            End If
        End With
    Next
    If strList = "" Then Exit Sub
    strList = Left(strList, Len(strList) - 1)
    
    Dim lCount As Long
    lCount = GetSingleLongValue("SUM(VendRcvdDetailWastageDetail.Qty)", "VendRcvdDetail INNER JOIN VendRcvdDetailWastageDetail ON VendRcvdDetail.EntryID=VendRcvdDetailWastageDetail.VRD_RefID INNER JOIN VendReceived ON VendReceived.EntryID=VendRcvdDetail.RefID", "WHERE VendReceived.EntryID IN(" & strList & ")")
    
    If lCount > 0 Then
        Dim rpt As CRAXDDRT.Report
        Set rpt = rptApp.OpenReport(RptPath & "\RepairIssGP.rpt")
        Dim f As New frmPrevRpt
        f.ShowReport "{VendReceived.EntryID} IN[" & strList & "]", rpt
    End If
    
'    lCount = GetSingleLongValue("SUM(Qty)", "MakerRepair INNER JOIN VendRcvdDetail ON VendRcvdDetail.EntryID=MakerRepair.Rcvd_RefID" & _
'         " INNER JOIN VendReceived ON VendReceived.EntryID=VendRcvdDetail.RefID", " WHERE VendReceived.EntryID IN(" & strList & ") AND MakerRepair.Wastage_Originator=1")
'    If lCount > 0 Then
'        'Now Show Report for Wastage to Originator...
'        Dim rpt1 As CRAXDDRT.Report
'        Set rpt1 = rptApp.OpenReport(RptPath & "\RepairIssGPCh2Originator.rpt")
'        Dim f1 As New frmPrevRpt
'        f1.ShowReport "{VendReceived.EntryID} IN[" & strList & "] AND {MakerRepair.Wastage_Originator}=TRUE", rpt1
'    End If
'
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdProcess_Click()
    
    If LVProcess.Visible Then
    
        LVProcess.Visible = False
        Dim i As Integer
        If LVProcess.ListItems(1).Checked Then
            txtProcess.Text = "<All Processes>"
            txtProcess.Tag = "0"
        Else
            Dim strProcesses As String, strProcessIDs As String
            For i = 1 To LVProcess.ListItems.count
                If LVProcess.ListItems(i).Checked Then
                    strProcesses = strProcesses & LVProcess.ListItems(i).Text & ","
                    strProcessIDs = strProcessIDs & LVProcess.ListItems(i).Tag & ","
                End If
            Next
            If strProcesses <> "" Then
                strProcesses = Left(strProcesses, Len(strProcesses) - 1)
                strProcessIDs = Left(strProcessIDs, Len(strProcessIDs) - 1)
            End If
            txtProcess.Text = strProcesses
            txtProcess.Tag = strProcessIDs
        End If
        
    Else
        LVProcess.Visible = True
    End If

End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub


Private Sub Command1_Click()

End Sub

Private Sub cmdSave_Click()
        
    If Opts(0).Value Or Opts(1).Value Then
        If getDBPassword("Split_PTC") = False Then Exit Sub
    ElseIf Opts(2).Value Then
        If getDBPassword("Transfer_To_Semi_Finish") = False Then Exit Sub
    End If
    
    Dim strLotNo As String, lEntryID As Long, strFromOrderNo As String, lOriginalQty As Long, lIssdQty As Long
    strLotNo = LV.SelectedItem.SubItems(10)
    
    lEntryID = Val(LV.SelectedItem.ListSubItems(9).Tag)
    strFromOrderNo = GetSingleStringValue("OrderNo", "VendRcvdDetail", " WHERE EntryID=" & lEntryID)
    lOriginalQty = Val(LV.SelectedItem.SubItems(6)) - Val(LV.SelectedItem.SubItems(7)) - Val(LV.SelectedItem.SubItems(8))
    lIssdQty = Val(LV.SelectedItem.ListSubItems(6).Tag)
    
    Dim lVendID As Long, lRefID As Long
    lVendID = GetSingleLongValue("VendID", "VendIssued", " WHERE EntryID=" & Val(LV.SelectedItem.Tag))
    
    If Opts(0).Value Then
        
        Dim strToOrderNo As String, strToItemCode As String
        If chkCOStockOrder.Value = vbUnchecked Then
        
            If cmbCOOrderNo.MatchFound = False Then
                MsgBox "Please Select OrderNo.", vbInformation
                Exit Sub
            ElseIf cmbCOItemID.MatchFound = False Then
                MsgBox "Please Select Item.", vbInformation
                Exit Sub
            End If
            
            If strFromOrderNo = cmbCOOrderNo.ID And cmbCOItemID.ID = LV.SelectedItem.ListSubItems(2).Tag Then
                MsgBox "Same Order NO. and Item Code.", vbInformation
                Exit Sub
            End If
            strToOrderNo = cmbCOOrderNo.ID
            strToItemCode = cmbCOItemID.ID
        Else
            strToItemCode = LV.SelectedItem.ListSubItems(2).Tag
            strToOrderNo = "Stock-Order"
            'Now Check If Article is Assigned to Customer
            Dim lCount As Long
            lCount = 0
            lCount = GetSingleLongValue("COUNT(CompItemID)", "FCustomerCatalog", " WHERE Custcode='Stock' AND Country='PK' AND CompItemID='" & strToItemCode & "'")
            If lCount = 0 Then
                con.Execute "INSERT INTO FCustomerCatalog(Custcode,Country,ItemID,CompItemID) VALUES('Stock','PK','" & strToItemCode & "','" & strToItemCode & "')"
            End If
            'Now Check If Article is feeded against this Article
            lCount = 0
            lCount = GetSingleLongValue("COUNT(OrderNo)", "FOrderItems", " WHERE OrderNo='Stock-Order' AND CompItemCode='" & strToItemCode & "'")
            If lCount = 0 Then
                con.Execute "INSERT INTO FOrderItems(OrderNo,ItemCode,CompItemCode,Qty,DeliveryDT,Stamps,Quality) VALUES('Stock-Order','" & strToItemCode & _
                 "','" & strToItemCode & "'," & Val(txtQty) & ",'" & Date & "','','')"
            Else
                con.Execute "UPDATE FOrderItems SET Qty=Qty+" & lOriginalQty & " WHERE OrderNo='Stock-Order' AND ItemCode='" & strToItemCode & "'"
            End If
            
            
        End If
        
        Call StartTrans(con)
        
        con.Execute "UPDATE VendRcvdDetail SET OrderNo='" & strToOrderNo & "',ItemCode='" & strToItemCode & "' WHERE LotNo='" & strLotNo & "'"
        con.Execute "UPDATE VendIssdDetail SET OrderNo='" & strToOrderNo & "',ItemCode='" & strToItemCode & "' WHERE LotNo='" & strLotNo & "'"
        con.Execute "INSERT INTO LotTransferDetails(VRD_From_RefID,VRD_To_RefID,FromOrderNo,ToOrderNo,Qty,SplitQty,Type," & _
         "FromItemCode,ToItemCode) VALUES(" & lEntryID & "," & lEntryID & ",'" & strFromOrderNo & "','" & strToOrderNo & _
         "'," & lOriginalQty & "," & lOriginalQty & ",0,'" & strToItemCode & "','" & strToItemCode & "')"
         
        con.CommitTrans
        FraSplitLot.Visible = False
        FAcc.Enabled = True
        
    ElseIf Opts(1).Value Then                       'Split Lot...
    
        If lIssdQty > 0 Then
            MsgBox "Can't Split this Lot, It's been Issued.", vbInformation
            Exit Sub
        ElseIf cmbSplitOrderNo.MatchFound = False Then
            MsgBox "Please Select OrderNo.", vbInformation
            Exit Sub
        ElseIf cmbSplitItemID.MatchFound = False Then
            MsgBox "Please Select Article.", vbInformation
            Exit Sub
        ElseIf Val(txtSplitQty) > lOriginalQty Then
            MsgBox "Invalid Qty.", vbInformation
            Exit Sub
        End If
        
        Dim strNewLotNo As String, lNewEntryID As Long
        Dim strForgeBatchNo As String
        
        strNewLotNo = getNextLotNo()
        strForgeBatchNo = GetSingleStringValue("RcvID", "VLot_With_ForgeBatchNo", " WHERE LotNo='" & strLotNo & "'")
        
        'Check If Billing is Done for this process for this lot.
        Dim lBill_Count As Long, iNot_Available_For_Billing As Integer
        lBill_Count = GetSingleLongValue("COUNT(*)", "MakerPostedBillsDetail_Receivings", " WHERE VRD_RefID=" & lEntryID)
        If lBill_Count > 0 Then
            iNot_Available_For_Billing = 1
        Else
            iNot_Available_For_Billing = 0
        End If
        
        Call StartTrans(con)
           
        con.Execute "UPDATE VendRcvdDetail SET RcvdQty=RcvdQty-" & Val(txtSplitQty) & " WHERE EntryID=" & lEntryID
        
        con.Execute "INSERT INTO VendReceived(VendID,DT,RecieptID,UserID,ProcessID,Issuance_RefID,EmpID) " & _
         "SELECT VendID,DT,'',UserID,ProcessID,Issuance_RefID,EmpID FROM VendReceived WHERE EntryID=" & Val(LV.SelectedItem.Tag)
        
        lRefID = GetSingleLongValue("MAX(EntryID)", "VendReceived")
        
        con.Execute "INSERT INTO VendRcvdDetail(RefID,ItemCode,RecieptID,RcvdQty,Wastage,IssQty,Rate,LotNo,ReqAuth,NextProcessID," & _
         "LostQty,OrderNo,CountedBy,Issue_RefID,ProcessID,RcvdWeight,Opening_RefID,ReworkLot,Repair_RefID,Not_Available_For_Billing) SELECT " & lRefID & ",'" & cmbSplitItemID.ID & "',''," & _
         Val(txtSplitQty) & ",0,0,Rate,'" & strNewLotNo & "',ReqAuth,NextProcessID,0,'" & cmbSplitOrderNo.ID & "','',Issue_RefID," & _
         "ProcessID,0,Opening_RefID,ReWorkLot,Repair_RefID," & iNot_Available_For_Billing & " FROM VendRcvdDetail WHERE EntryID=" & lEntryID
         
        'con.Execute "INSERT INTO Lots_List(LotNo,ItemID,Lot_Type,Reference_LotNo,Batch_No,Mill_Certificate_No) VALUES('" & strNewLotNo & "','" & cmbSplitItemID.ID & "',4,'')"
        con.Execute "INSERT INTO Lots_List(LotNo,ItemID,Lot_Type,Reference_LotNo,Batch_No,Mill_Certificate_No,Forge_Batch_No) SELECT " & strNewLotNo & ",'" & cmbSplitItemID.ID & "',4,'" & strLotNo & "',Batch_No,Mill_Certificate_No,'" & strForgeBatchNo & "' FROM Lots_List WHERE LotNo='" & strLotNo & "'"
        
        lNewEntryID = GetSingleLongValue("MAX(EntryID)", "VendRcvdDetail")
        
        con.Execute "INSERT INTO LotTransferDetails(VRD_From_RefID,VRD_To_RefID,FromOrderNo,ToOrderNo,Qty,SplitQty,Type,LotTransferRemarks) VALUES(" & _
         lEntryID & "," & lNewEntryID & ",'" & strFromOrderNo & "','" & cmbSplitOrderNo.ID & "'," & lOriginalQty & "," & Val(txtSplitQty) & ",1,'" & txtSplitLotRemarks & "')"
        
        con.CommitTrans
                
        MsgBox "Lot No. " & strNewLotNo & " Generated."
        MsgBox "Please Print PTC Cards for Both Lots."
        
        FraSplitLot.Visible = False
        FAcc.Enabled = True
        
    ElseIf Opts(2).Value Then
        If lIssdQty > 0 Then
            MsgBox "Can't Transfer this Entry, It's been Issued.", vbInformation
            Exit Sub
        End If
        'If SplitQty is More than RcvdQty-Rework-Wastage...
        If Val(txtSFQty) > (Val(LV.SelectedItem.SubItems(6)) - Val(LV.SelectedItem.SubItems(7)) - Val(LV.SelectedItem.SubItems(8))) Then
            MsgBox "Invalid Qty.", vbInformation
            Exit Sub
        End If
        
        If strCompany = "QEL" Then
            Call TransferToSFStock_QEL(LV.SelectedItem.ListSubItems(2).Tag, Val(LV.SelectedItem.ListSubItems(4).Tag), strLotNo, lEntryID, strFromOrderNo, lOriginalQty)
        Else
            Call TransferToSFStock(LV.SelectedItem.ListSubItems(2).Tag, Val(LV.SelectedItem.ListSubItems(4).Tag), strLotNo, lEntryID, strFromOrderNo, lOriginalQty, txtRemarks)
        End If
        
    End If
    
    If bShowForSplitTransferOnly Then
        Unload Me
        Call frmProductionPlanningDashBoard.RefreshRunningLots("")
    End If
    
End Sub

Private Sub TransferToSFStock(strItemCode As String, lProcessID As Long, strLotNo As String, lFromEntryID As Long, strFromOrderNo As String, lOriginalQty As Long, strRemarks As String)

    'Check If Opening for Selected Article for Selected Process is feeded ?
'    Dim lCount As Long
'    lCount = GetSingleLongValue("COUNT(ItemID)", "StockOrderOpening", " WHERE ItemID='" & strItemCode & "' AND ProcessID=" & lProcessID & " AND Shelf_RefID=" & cmbLocation.ID)
    
    Dim lNewEntryID As Long
    Call StartTrans(con)
        
    Dim lSOO_EntryID As Long
'    If lCount > 0 Then
'
'        'strRemarks = GetSingleStringValue("Remarks", "StockOrderOpening", " WHERE ItemID='" & strItemCode & "' AND ProcessID=" & lProcessID & " AND Shelf_RefID=" & cmbLocation.ID)
''        If strRemarks <> txtRemarks Then
''            strRemarks = strRemarks & "," & strLotNo
''        End If
'        strRemarks = strRemarks & "," & strLotNo & "(" & Val(txtSFQty) & ")"
'        lSOO_EntryID = GetSingleLongValue("EntryID", "StockOrderOpening", "WHERE ItemID='" & strItemCode & "' AND ProcessID=" & lProcessID & " AND Shelf_RefID=" & cmbLocation.ID)
'        con.Execute "UPDATE StockOrderOpening SET Qty=Qty+" & Val(txtSFQty) & ",Remarks='" & strRemarks & "' WHERE EntryID=" & lSOO_EntryID
'
'    Else
'        strRemarks = strRemarks & "," & strLotNo & "(" & Val(txtSFQty) & ")"
'        con.Execute "INSERT INTO StockOrderOpening(ItemID,ProcessID,Qty,Location,UserName,MachineName,Shelf_RefID,Remarks) VALUES('" & strItemCode & _
'         "'," & lProcessID & "," & Val(txtSFQty) & ",'" & cmbLocation.Text & "','" & CurrentUserName & "','" & strComputerName & "'," & cmbLocation.ID & ",'" & strRemarks & "')"
'
'        lSOO_EntryID = GetSingleLongValue("MAX(EntryID)", "StockOrderOpening")
'    End If
    
    strRemarks = strRemarks & "," & strLotNo & "(" & Val(txtSFQty) & ")"
    
    con.Execute "INSERT INTO StockOrderOpening(ItemID,ProcessID,Qty,Location,UserName,MachineName,Shelf_RefID,Remarks,LotNo_Manual) VALUES('" & strItemCode & _
     "'," & lProcessID & "," & Val(txtSFQty) & ",'" & cmbLocation.Text & "','" & CurrentUserName & "','" & strComputerName & "'," & cmbLocation.ID & ",'" & strRemarks & "','" & strLotNo & "')"

    lSOO_EntryID = GetSingleLongValue("MAX(EntryID)", "StockOrderOpening")

    'Check Customer...
    lCount = 0
    lCount = GetSingleLongValue("COUNT(CustCode)", "ForeignCustomers", " WHERE CustCode='Stock' AND Country='PK'")
    If lCount = 0 Then
        con.Execute "INSERT INTO ForeignCustomers(Custcode,Country) VALUES('Stock','PK')"
    End If
    
    'Now Check to see if Stock Order is feeded.
    lCount = 0
    lCount = GetSingleLongValue("COUNT(OrderNo)", "FCustomerOrders", " WHERE OrderNo='Stock-Order'")
    If lCount = 0 Then
        con.Execute "INSERT INTO FCustomerOrders(CustCode,Country,OrderNo,DT) VALUES('Stock','PK','Stock-Order','" & Date & "')"
    End If
    
    'Now Check If Article is Assigned to Customer
    lCount = 0
    lCount = GetSingleLongValue("COUNT(CompItemID)", "FCustomerCatalog", " WHERE Custcode='Stock' AND Country='PK' AND CompItemID='" & strItemCode & "'")
    If lCount = 0 Then
        con.Execute "INSERT INTO FCustomerCatalog(Custcode,Country,ItemID,CompItemID) VALUES('Stock','PK','" & strItemCode & "','" & strItemCode & "')"
    End If
    'Now Check If Article is feeded against this Article
    lCount = 0
    lCount = GetSingleLongValue("COUNT(OrderNo)", "FOrderItems", " WHERE OrderNo='Stock-Order' AND CompItemCode='" & strItemCode & "'")
    If lCount = 0 Then
        con.Execute "INSERT INTO FOrderItems(OrderNo,ItemCode,CompItemCode,Qty,DeliveryDT,Stamps,Quality) VALUES('Stock-Order','" & strItemCode & _
         "','" & strItemCode & "'," & Val(txtSFQty) & ",'" & Date & "','','')"
    Else
        con.Execute "UPDATE FOrderItems SET Qty=Qty+" & Val(txtSFQty) & " WHERE OrderNo='Stock-Order' AND ItemCode='" & strItemCode & "'"
    End If
    
    Dim lEntryID As Long, lRcvEntryID As Long, lNextProcessID As Long
    lEntryID = GetSingleLongValue("MAX(EntryID)", "StockOrderOpening")
'    con.Execute "INSERT INTO VendReceived(VendID,DT,RecieptID,UserID,ProcessID,Issuance_RefID) VALUES(" & _
     "0,'" & Date & "','',0," & lProcessID & ",0)"
    
'    lRcvEntryID = GetSingleLongValue("MAX(EntryID)", "VendReceived")
    
    lNextProcessID = getNextProcessID(strItemCode, lProcessID)
    
    Dim t_strLotNo As String
    t_strLotNo = "0"
    If Val(txtSFQty) = (Val(LV.SelectedItem.SubItems(6)) - Val(LV.SelectedItem.SubItems(7))) Then
        t_strLotNo = strLotNo
    Else
        t_strLotNo = "0"
    End If
'    con.Execute "INSERT INTO VendRcvdDetail(RefID,ItemCode,RecieptID,RcvdQty,Rate,LotNo,NextProcessID,OrderNo" & _
     ",CountedBy,Issue_RefID,ProcessID,RcvdWeight,Opening_RefID) VALUES(" & lRcvEntryID & ",'" & strItemCode & _
     "',''," & Val(txtSFQty) & ",0,'" & t_strLotNo & "'," & lNextProcessID & ",'Stock-Order','',0," & lProcessID & _
     ",0," & lEntryID & ")"
            
'    lNewEntryID = GetSingleLongValue("MAX(EntryID)", "VendRcvdDetail")
    con.Execute "INSERT INTO LotTransferDetails(VRD_From_RefID,VRD_To_RefID,FromOrderNo,ToOrderNo,Qty,SplitQty,Type,SOO_RefID) VALUES(" & _
         lFromEntryID & "," & lNewEntryID & ",'" & strFromOrderNo & "','Stock-Order'," & lOriginalQty & "," & Val(txtSFQty) & ",2," & lSOO_EntryID & ")"
    
'    con.Execute "UPDATE VendRcvdDetail SET RcvdQty=RcvdQty-" & Val(txtSFQty) & " WHERE EntryID=" & lFromEntryID
    con.Execute "UPDATE VendRcvdDetail SET IssQty=IssQty+" & Val(txtSFQty) & " WHERE EntryID=" & lFromEntryID
    
    con.CommitTrans
            
    MsgBox "Please Print PTC Cards for Both Lots."
    
    FraSplitLot.Visible = False
    FAcc.Enabled = True
     
End Sub

Private Sub TransferToSFStock_QEL(strItemCode As String, lProcessID As Long, strLotNo As String, lFromEntryID As Long, strFromOrderNo As String, lOriginalQty As Long)

    'Check If Opening for Selected Article for Selected Process is feeded ?
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(ItemID)", "StockOrderOpening", " WHERE ItemID='" & strItemCode & "' AND ProcessID=" & lProcessID)
    
    Dim lNewEntryID As Long
    Call StartTrans(con)
    
    If lCount > 0 Then
        con.Execute "UPDATE StockOrderOpening SET Qty=Qty+" & Val(txtQty_QEL) & " WHERE ItemID='" & strItemCode & "' AND ProcessID=" & lProcessID
        'MsgBox "This Item has already Opening feeded at this Process.", vbInformation
        'Exit Sub
    Else
        con.Execute "INSERT INTO StockOrderOpening(ItemID,ProcessID,Qty,Location,UserName,MachineName) VALUES('" & strItemCode & _
         "'," & lProcessID & "," & Val(txtQty_QEL) & ",'" & cmbLocation_QEL & "','" & CurrentUserName & "','" & strComputerName & "')"
    End If
     
    'Check Customer...
    lCount = 0
    lCount = GetSingleLongValue("COUNT(CustCode)", "ForeignCustomers", " WHERE CustCode='Stock' AND Country='PK'")
    If lCount = 0 Then
        con.Execute "INSERT INTO ForeignCustomers(Custcode,Country) VALUES('Stock','PK')"
    End If
    
    'Now Check to see if Stock Order is feeded.
    lCount = 0
    lCount = GetSingleLongValue("COUNT(OrderNo)", "FCustomerOrders", " WHERE OrderNo='Stock-Order'")
    If lCount = 0 Then
        con.Execute "INSERT INTO FCustomerOrders(CustCode,Country,OrderNo,DT) VALUES('Stock','PK','Stock-Order','" & Date & "')"
    End If
    
    'Now Check If Article is Assigned to Customer
    lCount = 0
    lCount = GetSingleLongValue("COUNT(CompItemID)", "FCustomerCatalog", " WHERE Custcode='Stock' AND Country='PK' AND CompItemID='" & strItemCode & "'")
    If lCount = 0 Then
        con.Execute "INSERT INTO FCustomerCatalog(Custcode,Country,ItemID,CompItemID) VALUES('Stock','PK','" & strItemCode & "','" & strItemCode & "')"
    End If
    'Now Check If Article is feeded against this Article
    lCount = 0
    lCount = GetSingleLongValue("COUNT(OrderNo)", "FOrderItems", " WHERE OrderNo='Stock-Order' AND CompItemCode='" & strItemCode & "'")
    If lCount = 0 Then
        con.Execute "INSERT INTO FOrderItems(OrderNo,ItemCode,CompItemCode,Qty) VALUES('Stock-Order','" & strItemCode & _
         "','" & strItemCode & "'," & Val(txtQty_QEL) & ")"
    Else
        con.Execute "UPDATE FOrderItems SET Qty=Qty+" & Val(txtQty_QEL) & " WHERE OrderNo='Stock-Order' AND ItemCode='" & strItemCode & "'"
    End If
    
    Dim lEntryID As Long, lRcvEntryID As Long, lNextProcessID As Long
    lEntryID = GetSingleLongValue("MAX(EntryID)", "StockOrderOpening")
    con.Execute "INSERT INTO VendReceived(VendID,DT,RecieptID,UserID,ProcessID,Issuance_RefID) VALUES(" & _
     "0,'" & Date & "','',0," & lProcessID & ",0)"
    
    lRcvEntryID = GetSingleLongValue("MAX(EntryID)", "VendReceived")
    
    lNextProcessID = getNextProcessID(strItemCode, lProcessID)
    
    con.Execute "INSERT INTO VendRcvdDetail(RefID,ItemCode,RecieptID,RcvdQty,Rate,LotNo,NextProcessID,OrderNo" & _
     ",CountedBy,Issue_RefID,ProcessID,RcvdWeight,Opening_RefID) VALUES(" & lRcvEntryID & ",'" & strItemCode & _
     "',''," & Val(txtQty_QEL) & ",0,'0'," & lNextProcessID & ",'Stock-Order','',0," & lProcessID & _
     ",0," & lEntryID & ")"
            
    lNewEntryID = GetSingleLongValue("MAX(EntryID)", "VendRcvdDetail")
    con.Execute "INSERT INTO LotTransferDetails(VRD_From_RefID,VRD_To_RefID,FromOrderNo,ToOrderNo,Qty,SplitQty,Type) VALUES(" & _
         lFromEntryID & "," & lNewEntryID & ",'" & strFromOrderNo & "','Stock-Order'," & lOriginalQty & "," & Val(txtQty_QEL) & ",2)"
    
    con.Execute "UPDATE VendRcvdDetail SET RcvdQty=RcvdQty-" & Val(txtQty_QEL) & " WHERE EntryID=" & lFromEntryID
    
    con.CommitTrans
        
    FraSplitLot.Visible = False
    FAcc.Enabled = True
        
End Sub

Private Sub cmdUpdateManualPTCNo_Click()
    
    Dim iRecAffect As Integer
    con.Execute "UPDATE VendRcvdDetail_MoreDetails SET ManualPTCNo='" & txtManualPTCNo & "' WHERE LotNo='" & txtPTCNo & "'", iRecAffect
    If iRecAffect = 0 Then
        con.Execute "INSERT INTO VendRcvdDetail_MoreDetails(LotNo,ManualPTCNo) VALUES('" & txtPTCNo & "','" & txtManualPTCNo & "')"
    End If
    
    FraManualPTCNo.Visible = False
    FAcc.Enabled = True
    
End Sub

Private Sub DtFrom_Change()
    'Call setDateRange(DtFrom)
End Sub

Private Sub DtTo_Change()
    'Call setDateRange(DtTo)
End Sub


Private Sub Form_Load()
     
    bShowForSplitTransferOnly = False
    strLotNo = ""
    lEntryID_Split = 0
    
'    If strCompany = "IAA" OR strCompany = "Kami" Then
'        mnuPrintListSummary.Visible = True
'    Else
'        mnuPrintListSummary.Visible = False
'    End If
    
    mnuPop.Visible = False
    
'    If strCompany = "QEL" Then
'        mnuPrintPTCWithPrice.Visible = True
'    Else
'        mnuPrintPTCWithPrice.Visible = False
'    End If
    
    If Not (strCompany = "QEL" Or (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami") Or strCompany = "Towne" Then
        mnuPrintPTC.Visible = False
    End If
    
    cmbMaker.ListHeight = 3800
    cmbMaker.AddItem "<All Makers>", "0"
    Call cmbMaker.AddVals(con, "'{ ' + VendID1 + ' } ' + VenderName ", "Makers", "VendID")
    cmbMaker.ID = "0"
    
'    cmbCustomer.ListHeight = 4000
'    cmbCustomer.AddItem "<All Customers>", "0"
'    Call cmbCustomer.AddVals(con, "CustCode", "ForeignCustomers", "CustCode")
'    cmbCustomer.ID = "0"
    
    DtFrom = Date
    DtTo = Date
    cmbRange.ListIndex = 0

    cmbItemCat.ListHeight = 4000
    cmbItemCat.EnableFunctionality = True
    cmbItemCat.AddItem "<All Catagories>", "0"
    cmbItemCat.AddVals con, "Description", "ItemCatagories", "CatID"
    cmbItemCat.ID = "0"
     
    cmbItemGroup.EnableFunctionality = True
    cmbItemGroup.ListHeight = 4000
    cmbItemGroup.AddItem "<All Groups>", "0"
    cmbItemGroup.AddVals con, "Description", "ItemGroups", "ID"
    cmbItemGroup.ID = "0"
    
    cmbItems.DropDownOffLine = True
    cmbItems.AddItem "<All Articles>", "0"
    Call cmbItems.AddVals(con, "ItemID + ' ' + ItemName + ' { ' + ISNULL(Attributes,'') + ' }'", "Items", "ItemID", " WHERE InActive=0 ORDER BY ItemName")
    cmbItems.ID = "0"
     
    cmbProcess.EnableFunctionality = True
    cmbProcess.ListHeight = 4000
    cmbProcess.AddItem "<All Processes>", "0"
    cmbProcess.AddVals con, "Description", "Processes", "ProcessID", " ORDER BY Description"
    
    cmbProcess.ID = "0"
    
    cmbCustomer.EnableFunctionality = True
    cmbCustomer.ListHeight = 4000
    cmbCustomer.AddItem "<All Customers>", "0"
    cmbCustomer.AddVals con, "DISTINCT CustCode", "ForeignCustomers", "CustCode", " ORDER BY ForeignCustomers.CustCode"
    
    cmbCustomer.ID = "0"
    
    With cmbfield
        .Clear
        For i = 1 To LV.ColumnHeaders.count
            cmbfield.AddItem LV.ColumnHeaders(i).Text
        Next
    End With
        
    cmbCOCust.ListHeight = 2400
    cmbCOOrderNo.ListHeight = 2400
    cmbCOItemID.ListHeight = 2400
    
    cmbSplitCust.ListHeight = 2400
    cmbSplitOrderNo.ListHeight = 2400
    cmbSplitItemID.ListHeight = 2400
    
    cmbStore.ListHeight = 2000
    cmbLocation.ListHeight = 2400
    Dim strCondition As String
    If CurrentUserName <> "Administrator" Then
        strCondition = " WHERE EntryID IN(SELECT Store_RefID FROM Users_Stores WHERE UserID=" & UserID & ")"
    End If
    cmbStore.AddVals con, "StoreName", "Stores", "EntryID", strCondition & " ORDER BY StoreName"
    
    If strCompany = "QEL" Then
        Call AddToCombo(cmbLocation_QEL, "DISTINCT Location", "IssItemsSimpleDetail")
    End If
    
    If strCompany = "IAA" Or strCompany = "Kami" Then
        mnuPrintReWorkRejectionChart.Visible = True
    Else
        mnuPrintReWorkRejectionChart.Visible = False
    End If
    
End Sub

Private Sub RefreshLV()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim Cond As String
    Dim DTRange As String
    Dim ITM As ListItem
    
    Screen.MousePointer = vbHourglass
    
    If cmbRange.ListIndex = 0 Then
        DTRange = " OnlyDT BETWEEN '" & DateAdd("d", 0, GetServerDate) & "' and '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 1 Then
        DTRange = " OnlyDT BETWEEN '" & DateAdd("d", -15, GetServerDate) & "' and '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 2 Then
        DTRange = " OnlyDT BETWEEN '" & DateAdd("d", -30, GetServerDate) & "' and '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 3 Then
        DTRange = " OnlyDT BETWEEN '" & DateAdd("d", -60, GetServerDate) & "' and '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 4 Then
        DTRange = " OnlyDT BETWEEN '" & DateAdd("d", -90, GetServerDate) & "' and '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 5 Then
        DTRange = " OnlyDT BETWEEN '" & DtFrom & "' AND '" & DtTo & "'"
    End If
    
    If Trim(txtLotNo) <> "" Or Trim(txtOrderNo) <> "" Then
        Cond = " WHERE OnlyDT BETWEEN '1/1/1' AND '1/1/2090'"
    Else
        Cond = " WHERE " & DTRange
    End If
    
    Cond = Cond & " AND ISNULL(Opening_RefID,0)=0"
    If cmbMaker.ID <> "0" Then
        Cond = Cond & " AND VendID=" & cmbMaker.ID & " "
    End If
    If cmbItemCat.ID <> "0" Then
        If cmbItemCat.Functionality = 1 Then
            Cond = Cond & " AND CatID IN(" & cmbItemCat.ID & ")"
        Else
            Cond = Cond & " AND CatID='" & cmbItemCat.ID & "'"
        End If
    End If
'    If cmbItemGroup.ID <> "0" Then
'        Cond = Cond & " AND GroupID=" & cmbItemGroup.ID
'    End If

    If cmbItemGroup.ID <> "0" Then
        Cond = Cond & " AND GroupID IN(" & cmbItemGroup.ID & ")"
    End If
    
    If cmbItems.ID <> "0" Then
        Cond = Cond & " AND ItemCode='" & cmbItems.ID & "' "
    End If
    
    If cmbProcess.ID <> "0" Then
        Cond = Cond & " AND ProcessID IN(" & cmbProcess.ID & ")"
    End If
'    If txtProcess.Tag <> "0" Then
'        Cond = Cond & " AND ProcessID IN(" & txtProcess.Tag & ")"
'    End If
    
    If Trim(txtLotNo) <> "" Then
        Cond = Cond & " AND LotNo='" & txtLotNo & "' "
    End If
    
    If Trim(txtOrderNo) <> "" Then
        Cond = Cond & " AND OrderNo='" & txtOrderNo & "' "
    End If
    
    If Trim(txtMasterPONo) <> "" Then
        Cond = Cond & " AND MasterPONo='" & Trim(txtMasterPONo) & "' "
    End If
    
    If chkShowMasterPOOnly.Value = vbChecked Then
        Cond = Cond & " AND MasterPONo IS NOT NULL"
    End If
    
    If chkRepair.Value = vbChecked And chkRejection.Value = vbChecked Then
        Cond = Cond & " AND (ReWorkQty>0 OR Wastage>0)"
    ElseIf chkRepair.Value = vbChecked Then
        Cond = Cond & " AND (ReWorkQty>0 OR ReWorkLot=1)"
    ElseIf chkRejection.Value = vbChecked Then
        Cond = Cond & " AND Wastage>0"
    End If
    
    If chkRegularLotsOnly.Value = vbChecked Then
        Cond = Cond & " AND (ReWorkLot=0)"
    End If
    
    If Val(txtInActiveDays) > 0 Then
        Dim myDTInactive As Date
        myDTInactive = DateAdd("d", -Val(txtInActiveDays), Date)
        Cond = Cond & " AND (LotNo<>'0' AND OnlyDT<='" & myDTInactive & "' AND IssQty=0)"
    End If
    
    If cmbCustomer.ID <> "0" Then
        If cmbCustomer.Functionality = 0 Then
            Cond = Cond & " AND CustCode IN('" & cmbCustomer.ID & "')"
        Else
            Cond = Cond & " AND CustCode IN(" & cmbCustomer.ID & ")"
        End If
    End If
     
     If lEntryID_Split > 0 Then
        Cond = " WHERE VRD_EntryID=" & lEntryID_Split
    End If
     
    If chkRcvdButNotIssued.Value = vbChecked Then
        Cond = Cond & " AND (VVendReceivingList.VRD_EntryID NOT IN(SELECT Rcvd_RefID FROM VendIssdDetail) AND LotNo IN(SELECT LotNo FROM VRunningLots_Simple) AND LotNo NOT IN(SELECT LotNo FROM Lots_Closed))"
    End If
     
    If chkBilled.Value = vbChecked Then
        Cond = Cond & " AND (MPB_D_EntryID IS NOT NULL)"
    End If
    
    If chkNonBilled.Value = vbChecked Then
        Cond = Cond & " AND (MPB_D_EntryID IS NULL)"
    End If
    
    If chkShowLastOnly.Value = vbChecked Then
        'Cond = Cond & " AND VRD_EntryID IN(SELECT MAX(EntryID) FROM VendRcvdDetail GROUP BY LotNo)"
        strReportSQL = "SELECT VVendReceivingList.* FROM VVendReceivingList INNER JOIN (SELECT MAX(EntryID) AS LastEntryID FROM VendRcvdDetail WHERE RcvdQty>IssQty GROUP BY LotNo) T1 ON VVendReceivingList.VRD_EntryID=T1.LastEntryID " & Cond & " ORDER BY DT"
    Else
        strReportSQL = "SELECT VVendReceivingList.* FROM VVendReceivingList " & Cond & " ORDER BY DT"
    End If
      
    
    
    'Cond = Cond & " AND Left(CustomInvoice,4)='VTL '"
    Dim lForeColor As Long
    With rs
        
        .Open strReportSQL, con, adOpenForwardOnly, adLockReadOnly
        LV.Visible = False
        LV.ListItems.Clear
        'LV.Sorted = False
        Dim strStatus As String
        
        Do Until .EOF
            If IsNull(!CompItemCode) Then
                lForeColor = vbRed
            ElseIf !ReWorklot Then
                lForeColor = lForeColor_RepairLot
            ElseIf !ComplaintItemID & "" <> "" Then
                lForeColor = RGB(200, 200, 0)
            Else
                lForeColor = LV.ForeColor
            End If
            
            If strCompany = "Tecno" Or (strCompany = "Dr-Frgz" And strSubCompany = "Electro") Then  'Check If this is the Process Just Before Scanning If Yes, THen Show otherwise Skip.
                If IsNextProcessScanning(Val(!ProcessID & ""), !ItemCode) = False Then
                    GoTo SKIPEntry
                End If
            End If
            
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            ITM.Tag = Val(![EntryID] & "")
            ITM.ListSubItems.add(, , "{ " & !VendID1 & " } " & !VenderName).Tag = "" '![ImpMaterial] & ""
            ITM.ListSubItems.add , , ![RecieptID] & ""
            ITM.ListSubItems.add(, , Format(![DT] & "", "dd-MMM-yyyy")).Tag = ""
            ITM.ListSubItems.add(, , !Description & "").Tag = Val(!ProcessID & "")
            ITM.ListSubItems.add(, , ![ItemCode] & " " & !ItemName & " " & !TipSize & " " & !ItemSize & " " & !SizeUnit & "").Tag = Val(!Issuance_RefID & "")
            
            ITM.ListSubItems.add(, , Val(!RcvdQty & "")).Tag = Val(!IssQty & "")
            ITM.ListSubItems.add , , Val(!Wastage & "") + Val(!LostQty & "")
            ITM.ListSubItems.add(, , Val(!ReWorkQty & "")).Tag = !LotNo & ""
            If !ReqAuth Then
                strStatus = "Under Auth."
            Else
                strStatus = "Authorized"
            End If
            
            If Val(!VRD_RefID & "") <> 0 Then
                strStatus = "Shipped"
            End If
            
            ITM.ListSubItems.add(, , strStatus).Tag = Val(!VRD_EntryID & "")
            ITM.ListSubItems.add(, , !LotNo & "").Tag = Val(!VRD_EntryID & "")
            ITM.ListSubItems.add , , !OrderNo & ""
            ITM.ListSubItems(2).Tag = !ItemCode & ""
            ITM.ForeColor = lForeColor
            For i = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(i).ForeColor = lForeColor
            Next
SKIPEntry:
            .MoveNEXT
        Loop
        .Close
    End With
    LV.Visible = True
    Set rs = Nothing
    Call RefreshSnos
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Private Function IsNextProcessScanning(lProcessID As Long, strItemID As String) As Boolean
    Dim iSNo As Integer
    iSNo = GetSingleLongValue("SNo", "ItemProcesses", " WHERE ItemID='" & strItemID & "' AND ProcessID=" & lProcessID)
    'Now Get The Next Process
    Dim bScanning As Boolean
    bScanning = GetSingleBooleanValue("TOP 1 Scanning", "ItemProcesses", " WHERE ItemID='" & strItemID & "' AND SNo>" & iSNo & " ORDER BY SNo")
    IsNextProcessScanning = bScanning
End Function

Private Sub Form_Resize()

    On Error Resume Next
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    FAcc.Top = (Me.ScaleHeight - FAcc.Height) / 2
    
    FraSplitLot.Move (Me.ScaleWidth - FraSplitLot.Width) / 2, ((Me.ScaleHeight - FraSplitLot.Height + 800) / 2) + 500
    
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If LV.SortKey + 1 = ColumnHeader.Index Then
        If LV.SortOrder = lvwAscending Then
            LV.SortOrder = lvwDescending
            'ColumnHeader.Icon = "Down"
        Else
            LV.SortOrder = lvwAscending
            'ColumnHeader.Icon = "Up"
        End If
    Else
        'LV.ColumnHeaderIcons = Nothing
        'Set LV.ColumnHeaderIcons = ImageList2
        LV.SortKey = ColumnHeader.Index - 1
        LV.SortOrder = lvwAscending
        'ColumnHeader.Icon = "Down"
    End If
    Call RefreshSnos
    'If LV.Sorted = False Then LV.Sorted = True
End Sub

Private Sub LV_DblClick()
    Exit Sub
    If bShipping Then Exit Sub
    If LV.ListItems.count = 0 Then Exit Sub
    Dim f As New frmNewCustomInvoice
    Load f
    Call f.EditProforma(Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1), bCustomInvoice)
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 2 Then Exit Sub
    If LV.ListItems.count = 0 Then
        Call ShowMenus(False)
    Else
        Call ShowMenus(True)
    End If
'    If bShipping Then
'        ShowMenusForShipping
'    ElseIf bCourierInvoice Then
'        ShowMenusforCourier
'    Else
'        If LV.ListItems.Count = 0 Then
'            Call ShowMenus(False)
'        Else
'            Call ShowMenus(True)
'        End If
'    End If
    'If LV.ListItems.Count = 0 Or Button <> 2 Or OutOfFinancialYear Then Exit Sub

    'if a Voucher Entry Show Delete Menu
    'If InStr(LV.SelectedItem, "-") > 0 Then Me.PopupMenu mnuPop
    
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

Private Sub mnuAddShippInfo_Click()

    Load frmEditShippInfo
    'ShowForm f, MainForm, , "Shipp. Info"
    Call frmEditShippInfo.EditEntry(LV.SelectedItem.SubItems(3))
    
End Sub

Private Sub mnuBillOfExchange_Click()
    Call PrintBillOfExchange
End Sub

Private Sub PrintBillOfExchange(Optional Firstrpt As Boolean = True)

    On Error GoTo err
    Me.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptBillOfExchange.rpt")
    rpt.FormulaFields(3).Text = Firstrpt
    Dim frm As New frmPrevRpt
    frm.ShowReport "{CustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    Me.MousePointer = vbNormal
    
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description

End Sub

    
Private Sub mnuBillOfExchange2_Click()
    Call PrintBillOfExchange(False)
End Sub

Private Sub mnuCommercialPList_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3), False, True, LV.SelectedItem.SubItems(1) = "DECATHLON", LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2))
End Sub

Private Sub mnuCommercialPListOther_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3), True, True, LV.SelectedItem.SubItems(1) = "DECATHLON", LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2))
End Sub

Private Sub mnuCoveringLetter_Click()
    
    On Error GoTo err
    Me.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptComCoveringLetter.rpt")
    If LCase(CurrentUserName) = "mnaeem" Then
        rpt.FormulaFields(3).Text = "'Naeem Ahmed'"
    End If
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "{CustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Private Sub mnuCustomPList_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3))
End Sub

Private Sub mnuCustomPListOther_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3), True)
End Sub

Private Sub mnuDelPackingInfo_Click()
    
    If LV.ListItems.count = 0 Then Exit Sub
    Dim iAffect As Integer
    If MsgBox("Are you sure to Delete Packing List of Invoice # " & LV.SelectedItem.SubItems(3) & " ?", vbQuestion + vbYesNo) = vbYes Then
        Call StartTrans(con)
        con.Execute "Delete From CustomPList Where CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'", iAffect
        con.CommitTrans
        If iAffect > 0 Then
            LV.SelectedItem.ListSubItems(2).Tag = 0
        End If
    End If
    
End Sub

Private Sub mnuedit_Click()
    Call LV_DblClick
End Sub

Private Sub mnuEForm_Click()

    Load frmPrintFormE
    If LV.ListItems.count > 0 Then
        frmPrintFormE.ShowForm LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
    Else
        frmPrintFormE.Show , MainForm
    End If
    
End Sub

Private Sub mnuGSP_Click()

'    Load frmGSP
'    If LV.ListItems.Count > 0 Then
'        frmGSP.ShowForm LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
'    Else
'        frmGSP.Show , MainForm
'    End If
    
End Sub

Private Sub mnuPacking_Click()

'    If Left(mnuPacking.Caption, 3) = "Add" Then
'
'        Load frmNewPList
'        frmNewPList.ShowMe LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
'    Else
'        Load frmNewPList
'        Call frmNewPList.EditPList(LV.SelectedItem.SubItems(3))
'        ShowForm frmNewPList, MainForm, , "Packing List"
'    End If
    
End Sub

Private Sub mnuPrint_Click()
    Call PrintCustomInvoice
End Sub

Private Sub PrintCustomInvoice(Optional bImpMaterialDetail As Boolean = False)

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
    frmPrevRpt.ShowReport "{VrptCustomInvoiceDetail.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt
    
    Screen.MousePointer = vbDefault


    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub mnuPrint10_Click()
    Call PrintCommercialInvoice(True, False)
End Sub

Private Sub mnuPrint5_Click()
    Call PrintCommercialInvoice(True, True)
End Sub

Private Sub mnuPrintBillofExchange_Click()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptBillOfExchange.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VRptProformas.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintComInv_Click()
    Call PrintCommercialInvoice
End Sub

Private Sub PrintCommercialInvoice(Optional bPer As Boolean = False, Optional bFivePer As Boolean = True)
    
    On Error GoTo err

    Screen.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt
    Dim rpt As CRAXDDRT.Report
    
    If UCase(LV.SelectedItem.SubItems(1)) = "MIKASA" And LV.SelectedItem.SubItems(2) = "JAPAN" Then
        Set rpt = rptApp.OpenReport(rptexportpath & "\rptCommercialInvoiceMikasaJP.rpt")
    ElseIf bPer Then
        Set rpt = rptApp.OpenReport(rptexportpath & "\rptCommercialInvoiceFairTrade.rpt")
        rpt.FormulaFields(22).Text = bFivePer
    ElseIf LV.SelectedItem.SubItems(1) = "DECATHLON" Then
        Set rpt = rptApp.OpenReport(rptexportpath & "\rptDecathlonCommercialInvoice.rpt")
    Else
        Set rpt = rptApp.OpenReport(rptexportpath & "\rptCommercialInvoice.rpt")
    End If
    f.ShowReport "{VrptCustomInvoiceDetail.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt
    
    If Not bPer Then
        If LV.SelectedItem.SubItems(2) = "Canada" Then
            Dim rpt1 As CRAXDDRT.Report
            Set rpt1 = rptApp.OpenReport(rptexportpath & "\rptCommercialInvoice.rpt")
            Dim f1 As New frmPrevRpt
            Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
            Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
            
            Set FormulaFields = rpt1.FormulaFields
            
            For Each FormulaField In FormulaFields
                If FormulaField.Name = "{@forHeading}" Then
                    FormulaField.Text = "'C A N A D I A N   C U S T O M S   I N V O I C E'"
                    Exit For
                End If
            Next
            f1.ShowReport "{VrptCustomInvoiceDetail.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt1
        End If
    End If
    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub mnuPrintSummary_Click()

    On Error GoTo err
    Me.MousePointer = vbHourglass
    
    Call StartTrans(con)
    con.Execute "Delete From PrintFOrderList"
    
    For i = 1 To LV.ListItems.count
        con.Execute "Insert Into PrintFOrderList(Sno,Customer,Country," & _
         "OrderNo,OrderDT,OrderAmt,Status,Curr) Values(" & LV.ListItems(i) & _
         ",'" & LV.ListItems(i).SubItems(1) & "','" & LV.ListItems(i).SubItems(2) & _
         "','" & LV.ListItems(i).SubItems(3) & "','" & _
         LV.ListItems(i).SubItems(4) & "'," & Val(LV.ListItems(i).ListSubItems(5).Tag & "") & _
         ",'" & LV.ListItems(i).SubItems(6) & "','" & LV.ListItems(i).Tag & "')"
    Next i
    
    con.CommitTrans
    
    Load frmPrevRpt
    
    Dim strFromTo As String
    Dim iNumber As Integer
    If cmbRange.ListIndex = 0 Then
        iNumber = -15
    ElseIf cmbRange.ListIndex = 1 Then
        iNumber = -30
    ElseIf cmbRange.ListIndex = 2 Then
        iNumber = -60
    ElseIf cmbRange.ListIndex = 3 Then
        iNumber = -90
    End If
    
    If cmbRange.ListIndex = 4 Then
        strFromTo = Format(DtFrom, "dd-MMM-yyyy") & " To  " & Format(DtTo, "dd-MMM-yyyy")
    Else
        strFromTo = Format(DateAdd("d", iNumber, ServerDate), "dd-MMM-yyyy") & " To  " & Format(ServerDate, "dd-MMM-yyyy")
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptFProformaList.rpt")
    rpt.FormulaFields(3).Text = "'" & cmbMaker.Text & "'"
    rpt.FormulaFields(4).Text = "'" & strFromTo & "'"
    
    frmPrevRpt.ShowReport "", rpt
    
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Private Sub mnuPrintCoverLetter_Click()

    On Error GoTo err

    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptDecathlonCoveringLetter.rpt")
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VCustomInvoice.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt

    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub mnuPrintDecDocs_Click()
    Load frmPrintDecathlonDocs
    frmPrintDecathlonDocs.ShowMe (LV.SelectedItem.SubItems(3))
End Sub

Private Sub mnuPrintInvWDetail_Click()

    Call PrintCustomInvoice(True)
    
End Sub

Private Sub mnuPrintPackingLabels_Click()

    On Error GoTo err
    
    If LV.SelectedItem.SubItems(1) <> "DECATHLON" Then
        MsgBox "Packing Labels are available only for DECATHLON at the moment.", vbInformation
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report, f As New frmPrevRpt
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptDecathlonPackingLabels.rpt")
    f.ShowReport "{VCustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    
    Me.MousePointer = vbNormal
    
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub mnuShippingInstructions_Click()
'    Load frmShippingInstructions
'    If LV.ListItems.Count = 0 Then
'        frmShippingInstructions.Show , MainForm
'    Else
'        frmShippingInstructions.ShowForm LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
'    End If
End Sub

Private Sub mnuUpdateCustomPrice_Click()
    If LV.ListItems.count = 0 Then Exit Sub
    If MsgBox("Are you Sure, You want to Update Custom Prices?", vbInformation + vbYesNo) = vbNo Then Exit Sub
    
    Dim dCustomPrice As Double
    dCustomPrice = Val(con.Execute("Select SUM(Qty*Price)/SUM(Qty) From CustomInvoiceItems Where CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'").Fields(0).Value & "")
    con.Execute "Update CustomInvoiceItems SET CustomPrice=" & dCustomPrice & " WHERE CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'"
    
End Sub

Private Sub txtDesc_GotFocus()
    cmdRefresh.Default = False
    txtDesc.SelStart = 0
    txtDesc.SelLength = Len(txtDesc)
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
  
    If KeyAscii = 13 Then
    
        If Trim(txtDesc = "") Then
            MsgBox "Invalid Description", vbInformation
            Exit Sub
        End If
    
        con.Execute "Update Vouchers Set [Description]='" & txtDesc & "' Where Sno=" & Val(LV.SelectedItem.key)
        LV.SelectedItem.ListSubItems(3) = txtDesc & txtDesc.Tag
        txtDesc.Visible = False
        cmdRefresh.Default = False
    ElseIf KeyAscii = 27 Then
        txtDesc.Visible = False
        cmdRefresh.Default = False
    End If

End Sub

Private Sub txtDesc_LostFocus()
    txtDesc.Visible = False
End Sub

Private Sub RefreshSnos()
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Text = i
    Next i
End Sub

Private Sub LVCustomer_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer
    If Item.Checked Then
        If Item.Index = 1 Then
            For i = 2 To LVCustomer.ListItems.count
                LVCustomer.ListItems(i).Checked = False
            Next
        Else
            LVCustomer.ListItems(1).Checked = False
        End If
    Else
'        If Item.Index > 1 Then
'            lvcustomer.ListItems(1).Checked = False
'        End If
    End If
End Sub

Private Sub LVGroups_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    Dim i As Integer
    If Item.Checked Then
        If Item.Index = 1 Then
            For i = 2 To LVGroups.ListItems.count
                LVGroups.ListItems(i).Checked = False
            Next
        Else
            LVGroups.ListItems(1).Checked = False
        End If
    Else
'        If Item.Index > 1 Then
'            LVGroups.ListItems(1).Checked = False
'        End If
    End If

End Sub

Private Sub LVProcess_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    Dim i As Integer
    If Item.Checked Then
        If Item.Index = 1 Then
            For i = 2 To LVProcess.ListItems.count
                LVProcess.ListItems(i).Checked = False
            Next
        Else
            LVProcess.ListItems(1).Checked = False
        End If
    Else
'        If Item.Index > 1 Then
'            lvcustomer.ListItems(1).Checked = False
'        End If
    End If

End Sub

Private Sub mnuCloseLot_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    If getDBPassword("Close_Lot") = False Then Exit Sub
    
    Dim lCount As Long, strLotNo As String
    strLotNo = LV.SelectedItem.SubItems(10)
    lCount = GetSingleLongValue("COUNT(*)", "Lots_Closed", " WHERE LotNo='" & strLotNo & "'")
    
    If lCount > 0 Then
        MsgBox "Lot Already Closed.", vbInformation
        Exit Sub
    End If
    con.Execute "INSERT INTO Lots_Closed(LotNo,UserName,MachineName) VALUES('" & strLotNo & "','" & CurrentUserName & "','" & strComputerName & "')"
    
    MsgBox "Done.", vbInformation
    
End Sub

Private Sub mnuDeleteRcving_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    
    Dim lEntryID As Long, lLastEntryID As Long
    lEntryID = Val(LV.SelectedItem.ListSubItems(9).Tag) 'Val(LV.SelectedItem.Tag)
'    lLastEntryID = GetSingleLongValue("MAX(EntryID)", "VendReceived")
'    If lEntryID <> lLastEntryID Then
'        MsgBox "This is not the last entry.", vbInformation
'        Exit Sub
'    End If
    If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo) = vbYes Then
        Dim lIssCount As Long
        'lIssCount = GetSingleLongValue("Rcvd_RefID", "VendIssdDetail", " WHERE Rcvd_RefID IN(SELECT EntryID FROM VendRcvdDetail WHERE RefID=" & lEntryID & ")")
        lIssCount = GetSingleLongValue("Rcvd_RefID", "VendIssdDetail", " WHERE Rcvd_RefID=" & lEntryID)
        If lIssCount > 0 Then
            MsgBox "Can't Delete, Issuance Exist.", vbInformation
            Exit Sub
        End If
         
        Dim lSplitCount As Long
        lSplitCount = GetSingleLongValue("COUNT(VRD_From_RefID)", "LotTransferDetails", " WHERE VRD_From_RefID=" & lEntryID)
        If lSplitCount > 0 Then
            MsgBox "Can't Delete, Lot has been Split.", vbInformation
            Exit Sub
        End If
        
        Dim lBillingCount As Long
        lBillingCount = GetSingleLongValue("COUNT(VRD_RefID)", "MakerPostedBillsDetail_Receivings", " WHERE VRD_RefID=" & lEntryID)
        If lBillingCount > 0 Then
            MsgBox "Can't Delete, Lot has been Billed.", vbInformation
            Exit Sub
        End If
        
        If getDBPassword("DeleteProdRcv") Then
        
            Dim lQty As Long, lCurrQty As Long, lTempTot As Long, lTempQty As Long, strItemID As String, strLotNo As String
            strItemID = LV.SelectedItem.ListSubItems(2).Tag
            strLotNo = LV.SelectedItem.SubItems(10)
            lQty = GetSingleLongValue("SUM(RcvdQty)", "VendRcvdDetail", " WHERE EntryID=" & lEntryID)
            Dim rs As New ADODB.Recordset
            Call StartTrans(con)
            With rs
            
                'Check If It's Receiving from Stock Order Adjustment....
                Dim lSOACount As Long
                lSOACount = GetSingleLongValue("VID_RefID", "StockOrderAdjustments", " WHERE VID_RefID=" & lEntryID)
                If lSOACount > 0 Then
                    Dim lShelf_RefID As Long
                    Dim frmGetLoc As New frmGetLocation
                    Load frmGetLoc
                    If frmGetLoc.GetShelf(strItemID, lShelf_RefID) = False Then
                        Exit Sub
                    End If
                    
                   Call PostRFItemReceiving(strItemID, lShelf_RefID, CDbl(lQty), strLotNo)
                   
                   con.Execute "DELETE FROM StockOrderAdjustments WHERE VID_RefID=" & lEntryID
                End If
                
                .Open "SELECT VendIssdDetail_ReturnDTs.EntryID,RcvQty FROM VendIssdDetail_ReturnDTs INNER JOIN VendIssdDetail ON " & _
                 "VendIssdDetail_ReturnDTs.VIS_RefID=VendIssdDetail.EntryID INNER JOIN VendRcvdDetail ON VendIssdDetail.EntryID" & _
                 "=VendRcvdDetail.Issue_RefID INNER JOIN VendReceived ON VendRcvdDetail.RefID=VendReceived.EntryID WHERE RcvQty>0 " & _
                 "AND VendRcvdDetail.EntryID=" & lEntryID & " ORDER BY VendIssdDetail_ReturnDTs.ReturnDT DESC", con_ServerSide, adOpenForwardOnly, adLockReadOnly
                '
                Do Until .EOF
                    
                    
                    lCurrQty = Val(!RcvQty & "")
                    
                    If lCurrQty > (lQty - lTempTot) Then
                        lTempQty = (lQty - lTempTot)
                    Else
                        lTempQty = lCurrQty
                    End If
                    
                    con.Execute "UPDATE VendIssdDetail_ReturnDTs SET RcvQty=RcvQty-" & lTempQty & " WHERE EntryID=" & Val(!EntryID & "")
                    lTempTot = lTempTot + lTempQty
                    
                    If lTempTot = lQty Then
                        Exit Do
                    End If
                    .MoveNEXT
                Loop
                    
                .Close
            End With
            Set rs = Nothing
            
            con.Execute "UPDATE VendIssdDetail SET VendIssdDetail.RcvdQty=VendIssdDetail.RcvdQty-Tab1.RcvdQty " & _
             " FROM VendIssdDetail INNER JOIN (SELECT Issue_RefID,SUM(RcvdQty) AS RcvdQty FROM VendRcvdDetail" & _
             " WHERE VendRcvdDetail.EntryID=" & lEntryID & " GROUP BY " & _
             "Issue_RefID) Tab1 ON VendIssdDetail.EntryID=Tab1.Issue_RefID"
              
            con.Execute "INSERT INTO VendRcvdDetail_Deletions(LotNo,ProcessID,DTRcving,RcvdQty,WastageQty,ReWorkQty," & _
             "UserName,MachineName,VRD_RefID,VR_RefID,ItemID,MakerID) SELECT LotNo,VendRcvdDetail.ProcessID,DT,RcvdQty,Wastage,ReWorkQty,'" & _
             CurrentUserName & "','" & strComputerName & "',VendRcvdDetail.EntryID,VendReceived.EntryID,ItemCode,VendID FROM VendReceived " & _
             "INNER JOIN VendRcvdDetail ON VendReceived.EntryID=VendRcvdDetail.RefID WHERE VendRcvdDetail.EntryID=" & lEntryID
            
            con.Execute "DELETE FROM VendRcvdDetail WHERE EntryID=" & lEntryID
            con.Execute "DELETE FROM VendReceived WHERE EntryID NOT IN(SELECT RefID FROM VendRcvdDetail)"
            
            
            con.CommitTrans
        End If
        
    End If
    
End Sub

Private Sub mnuInspectionData_Click()
    
    If LV.ListItems.count = 0 Then Exit Sub
    
    Dim lEntryID As Long, lProcessID As Long
    lEntryID = Val(LV.SelectedItem.ListSubItems(9).Tag)
'    lProcessID = Val(LV.SelectedItem.ListSubItems(4).Tag)
'    Dim bInspectionProcess As Boolean
'    bInspectionProcess = GetSingleBooleanValue("InspectionProcess", "Processes", " WHERE ProcessID=" & lProcessID)
'
'    If bInspectionProcess = False Then
'        MsgBox "This is not Inspection Process.", vbInformation
'        Exit Sub
'    End If
    
    Dim f As New frmRcvdLotInspection
    Load f
    f.ShowMe (lEntryID)
    
End Sub

Private Sub mnuManualPTCNo_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    
    FAcc.Enabled = False
    txtPTCNo = LV.SelectedItem.SubItems(9)
    
    
    txtManualPTCNo = GetSingleStringValue("ManualPTCNo", "VendRcvdDetail_MoreDetails", " WHERE LotNo='" & txtPTCNo.Text & "'")
    
    FraManualPTCNo.Visible = True
    txtManualPTCNo.SetFocus
    
End Sub

Private Sub mnuPrintDeletedLots_Click()
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\DeletedLots_Receiving.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt
End Sub

Private Sub mnuPrintList_Click()
    Call PrintList(False)
End Sub

Private Sub PrintList(Optional bSummary As Boolean = False, Optional bProcessList As Boolean = False)

    Dim strFilters As String
    strFilters = GetFilters
   
    Dim strDateRange As String
    strDateRange = Format(DtFrom.Value, "dd-MMM-yyyy") & " to " & Format(DtTo.Value, "dd-MMM-yyyy")
    
    Dim rpt As CRAXDDRT.Report
    If bSummary Then
        If strCompany = "IAA" Or strCompany = "Kami" Then
            Set rpt = rptApp.OpenReport(RptPath & "\ReceivingListSummary.rpt")
        Else
            Set rpt = rptApp.OpenReport(RptPath & "\ReceivingListSummaryItemWise.rpt")
        End If
    Else
        If bProcessList Then
            Set rpt = rptApp.OpenReport(RptPath & "\ReceivingList_Process.rpt")
        Else
            Set rpt = rptApp.OpenReport(RptPath & "\ReceivingList.rpt")
        End If
    End If
    
    rpt.FormulaFields.GetItemByName("Filters").Text = "'" & Replace(strFilters, "'", "") & "'"
    rpt.FormulaFields.GetItemByName("DateRange").Text = "'" & strDateRange & "'"
    
    
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, strReportSQL
    
End Sub

Private Function GetFilters() As String

    strFilters = "Maker:" & cmbMaker.Text & " | Process:" & cmbProcess.Text & " | Catagory:" & cmbItemCat.Text & " | Article:" & cmbItems.Text & " | Customer:" & cmbCustomer.ID
    
    If Trim(txtLotNo) <> "" Then
        strFilters = strFilters & " | Lot No.:" & Trim(txtLotNo)
    End If
    
    If Trim(txtOrderNo) <> "" Then
        strFilters = strFilters & " | Order No.:" & Trim(txtOrderNo)
    End If
    
    If Trim(txtInActiveDays) <> "" Then
        strFilters = strFilters & " | In-Active Days:" & Trim(txtInActiveDays)
    End If
        
    If chkShowMasterPOOnly.Value = vbChecked Then
        strFilters = strFilters & " | Master POs Only"
    End If
    
    If chkShowLastOnly.Value = vbChecked Then
        strFilters = strFilters & " | Showing Only Last"
    End If
    
    If chkRepair.Value = vbChecked Then
        strFilters = strFilters & " | Repair Lots"
    End If
    
    If chkRejection.Value = vbChecked Then
        strFilters = strFilters & " | Rejection Lots"
    End If
        
    GetFilters = Replace(strFilters, "'", "")
    
End Function

Private Sub mnuPrintListSummary_Click()
    Call PrintList(True)
End Sub

Private Sub mnuPrintProcessList_Click()
    Call PrintList(False, True)
End Sub

Private Sub mnuPrintPTC_Click()
    Call PrintPTC(LV.SelectedItem.ListSubItems(8).Tag, False)
End Sub


Private Sub mnuPrintPTCMini_Click()
    Call PrintPTC(LV.SelectedItem.ListSubItems(8).Tag, False, 1)
End Sub

Private Sub mnuPrintPTCWithPrice_Click()

    Call PrintPTC(LV.SelectedItem.ListSubItems(8).Tag, True)
    
End Sub

Private Sub mnuPrintPurchaseOrder_Click()
    
    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\IssSlipWithRcving.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VendIssued.EntryID}=" & LV.SelectedItem.ListSubItems(5).Tag, rpt
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub mnuPrintQCReport_Click()
    
    Dim f As New frmPrintProcessQCReport
    Load f
    f.ShowMe (Val(LV.SelectedItem.ListSubItems(10).Tag))
    '-------------------------
    
    
End Sub

Private Sub mnuPrintRepairSlip_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\RepairSlip.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VendRcvdDetail.EntryID}=" & Val(LV.SelectedItem.ListSubItems(9).Tag), rpt
    
End Sub

Private Sub mnuPrintReWorkRejectionChart_Click()

    Dim strFilters As String
    strFilters = GetFilters
    
    
    Dim strDateRange As String
    strDateRange = Format(DtFrom.Value, "dd-MMM-yyyy") & " to " & Format(DtTo.Value, "dd-MMM-yyyy")
    
    Dim rpt As CRAXDDRT.Report, rptApp_Temp As New CRAXDDRT.Application
    Set rpt = rptApp_Temp.OpenReport(RptPath & "\RepairRejectionGraphs.rpt")
    rpt.FormulaFields.GetItemByName("Filters").Text = "'" & strFilters & "'"
    rpt.FormulaFields.GetItemByName("DateRange").Text = "'" & strDateRange & "'"
    
    Dim rptReWork As CRAXDDRT.Report
    Dim rptWastage As CRAXDDRT.Report
     
    Set rptReWork = rpt.OpenSubreport("ReWorkSummarySub")
    On Error Resume Next
    rptReWork.Database.Tables(1).ConnectBufferString = strConnectionStringrpt
    On Error GoTo 0
    rptReWork.SQLQueryString = Replace(strReportSQL, "VVendReceivingList", "VVendRcvdDetailReWorkDetail")
    
    Set rptWastage = rpt.OpenSubreport("WastageSummarySub")
    On Error Resume Next
    rptWastage.Database.Tables(1).ConnectBufferString = strConnectionStringrpt
    On Error GoTo 0
    rptWastage.SQLQueryString = Replace(strReportSQL, "VVendReceivingList", "VVendRcvdDetailWastageDetail")
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt
    
End Sub

Private Sub mnuPrintSlipMini_Click()
    
    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\IssSlipWithRcving_Mini.rpt")
    
    Dim rpt_Sub As CRAXDDRT.Report
    Set rpt_Sub = rpt.OpenSubreport("Rcvings")
    rpt_Sub.RecordSelectionFormula = "{VendRcvdDetail.EntryID}=" & Val(LV.SelectedItem.ListSubItems(9).Tag)
    
    Dim f As New frmPrevRpt
    f.ShowReport "{VendIssued.EntryID}=" & LV.SelectedItem.ListSubItems(5).Tag, rpt
    
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub mnuProcessInspection_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    Dim f As New frmProcesses_Inspection_Points_Entry
    Load f
    f.ShowMe (Val(LV.SelectedItem.ListSubItems(9).Tag))
    
End Sub

Private Sub mnuProcessInspectionReport_Click()
    If LV.ListItems.count = 0 Then Exit Sub
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Process_Inspection_Report.rpt")
    
    Dim f As New frmPrevRpt
    Load f
    f.ShowReport "{VRD_PIP_Details.VRD_RefID}=" & (Val(LV.SelectedItem.ListSubItems(9).Tag)), rpt
End Sub

Private Sub mnuRRReport_Click()

    Dim strFilters As String
    strFilters = GetFilters
    
    
    Dim strDateRange As String
    strDateRange = Format(DtFrom.Value, "dd-MMM-yyyy") & " to " & Format(DtTo.Value, "dd-MMM-yyyy")
    
    Dim rpt As CRAXDDRT.Report, rptApp_Temp As New CRAXDDRT.Application
    Set rpt = rptApp_Temp.OpenReport(RptPath & "\RepairRejectionProcessWiseSummary.rpt")
    rpt.FormulaFields.GetItemByName("Filters").Text = "'" & strFilters & "'"
    rpt.FormulaFields.GetItemByName("DateRange").Text = "'" & strDateRange & "'"
    
    Dim rptReWork As CRAXDDRT.Report
    Dim rptWastage As CRAXDDRT.Report
     
    Set rptReWork = rpt.OpenSubreport("ReWorkSummarySub")
    On Error Resume Next
    rptReWork.Database.Tables(1).ConnectBufferString = strConnectionStringrpt
    On Error GoTo 0
    rptReWork.SQLQueryString = Replace(strReportSQL, "VVendReceivingList", "VVendRcvdDetailReWorkDetail")
    
    Set rptWastage = rpt.OpenSubreport("WastageSummarySub")
    On Error Resume Next
    rptWastage.Database.Tables(1).ConnectBufferString = strConnectionStringrpt
    On Error GoTo 0
    rptWastage.SQLQueryString = Replace(strReportSQL, "VVendReceivingList", "VVendRcvdDetailWastageDetail")
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, strReportSQL
    
End Sub

Private Sub mnuSplitLot_Click()

    Dim i As Integer
    For i = 0 To Opts.UBound - 1
        Opts(i).Value = False
    Next
    
    Call ClearSplitData
    FAcc.Enabled = False
    FraSplitLot.Visible = True
    
End Sub

Private Sub ClearSplitData()

    Opts(0).Value = False
    Opts(1).Value = False
    Opts(2).Value = False
    
    cmbCOCust.ID = ""
    cmbCOOrderNo.ClearVals
    cmbCOItemID.ClearVals
    
    cmbSplitCust.ID = ""
    cmbSplitOrderNo.ClearVals
    cmbSplitItemID.ClearVals
    txtSplitQty = ""
    txtSplitLotRemarks = ""
    
    cmbStore.ID = ""
    cmbLocation.ID = ""
    txtSFQty = ""
    txtRemarks = ""
    
End Sub

Private Sub Opts_Click(Index As Integer)

    Dim i As Integer
    For i = 0 To Opts.UBound
        Fra(i).Visible = False
    Next
    
    Fra(Index).Visible = True
    If Index = 0 Then
        cmbCOCust.ClearVals
        'cmbCOCust.AddVals con, "DISTINCT CustCode", "FCustomerOrders", "CustCode", " WHERE CustCode<>'Stock'"
        cmbCOCust.AddVals con, "DISTINCT CustCode", "FCustomerOrders", "CustCode", ""
    ElseIf Index = 1 Then
        cmbSplitCust.ClearVals
        cmbSplitCust.AddVals con, "DISTINCT CustCode", "FCustomerOrders", "CustCode"
    End If
    
    If Index = 2 Then
        'txtSFQty.Locked = True
        txtSFQty = Val(LV.SelectedItem.SubItems(6))
        If strCompany = "QEL" Then
            FraTSF_QEL.Visible = True
            FraTSF_Frgz.Visible = False
        Else
            FraTSF_Frgz.Visible = True
            FraTSF_QEL.Visible = False
        End If
    Else
        txtSFQty.Locked = False
    End If
    
'    If Index = 2 Then
'        If strCompany = "QEL" Then
'            FraTSF_QEL.Visible = True
'            FraTSF_Frgz.Visible = False
'        Else
'            FraTSF_Frgz.Visible = True
'            FraTSF_QEL.Visible = False
'        End If
'    End If
    
End Sub

Private Sub txtFind_Change()
    cmdFind.Enabled = Len(Trim(txtFind)) > 0
End Sub


Private Sub txtLotNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call RefreshLV
    End If
End Sub


Private Sub PostRFItemReceiving(strMaterialID As String, lShelf_RefID As Long, dQty As Double, Optional p_strLotNo As String)
         
    Dim strRcvNo As String
    
    strRcvNo = GetRcvingNo(Date)
    'Now Receiving...
    con.Execute "INSERT INTO RcvItemsSimple(RcvNo,DT,OrderNo,RcvdBy,RcvdFrom,UserName,ComputerName) VALUES('" & _
     strRcvNo & "','" & Date & " " & ServerTime & "','Deleted from Stock Order Adjust.','','','" & CurrentUserName & _
     "','" & strComputerName & "')"
    
    lRefID = con.Execute("SELECT MAX(EntryID) FROM RcvItemsSimple").Fields(0).Value & ""
    
    Dim lRISD_EntryID As Long
    
    dBal = getItemBalance(strMaterialID, Date)
    dBal = dBal + dQty
            
    theDesc = "Material Movement"
            
    Dim strShelf As String
    strShelf = GetSingleStringValue("StoreName + ' - ' + RackNo + ' (' + ShelfNo + ')'", "VStoreShelfs", " WHERE EntryID=" & lShelf_RefID)
    
    con.Execute "INSERT INTO ItemLedgerSimple(ItemID,DT,Description,Qty,Balance,OrderNo,Location,Remarks) VALUES('" & _
     strMaterialID & "','" & Date & "','" & theDesc & "'," & dQty & "," & dBal & ",'','" & strShelf & "','')"
             
    lRefID = Val(con.Execute("Select Max(EntryID) From ItemLedgerSimple").Fields(0).Value & "")
            
    Call adjustItemLedger(strMaterialID, Date, dQty)

    con.Execute "INSERT INTO RcvItemsSimpleDetail(RcvNo,ItemID,Qty,InHand,LedgerRefID,OrderNo,Location,Remarks,LotNo_Manual) Values('" & _
     strRcvNo & "','" & strMaterialID & "'," & dQty & "," & getItemBalance(strMaterialID, Date) - dQty & "," & lRefID & _
     ",'','','','" & p_strLotNo & "')"
            
    lRISD_EntryID = GetSingleLongValue("MAX(EntryID)", "RcvItemsSimpleDetail")
            
    con.Execute "INSERT INTO RcvItemsSimpleDetail_Placement(RISD_RefID,Shelf_RefID,RcvdQty,Remarks) VALUES(" & _
     lRISD_EntryID & "," & lShelf_RefID & "," & dQty & ",'')"
            
End Sub

Public Function ShowForSplit(p_strLotNo As String, p_lEntryID As Long) As Boolean

    strLotNo = p_strLotNo
    lEntryID_Split = p_lEntryID
    bShowForSplitTransferOnly = True
    
    txtLotNo = strLotNo
    Call cmdRefresh_Click
     
    LV.ListItems(1).Selected = True
    
    Call mnuSplitLot_Click
    
    Me.Show
    
End Function
