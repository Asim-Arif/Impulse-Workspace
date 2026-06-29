VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmMakerIssuanceFromSF 
   ClientHeight    =   10125
   ClientLeft      =   1110
   ClientTop       =   360
   ClientWidth     =   11805
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10125
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkUseOriginalLotNo 
      Caption         =   "Use Original Lot No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Left            =   420
      TabIndex        =   81
      Top             =   7740
      Width           =   2115
   End
   Begin VB.CheckBox chkIssByLotNo 
      Caption         =   "Issuane by Lot No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   420
      TabIndex        =   80
      Top             =   690
      Width           =   1965
   End
   Begin ComboList.Usercontrol1 cmbProcess 
      Height          =   285
      Left            =   1290
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2355
      _ExtentX        =   4154
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
   Begin ComboList.Usercontrol1 cmbVender 
      Height          =   285
      Left            =   3285
      TabIndex        =   10
      Top             =   7410
      Width           =   5160
      _ExtentX        =   9102
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
   Begin ComboList.Usercontrol1 cmbStore 
      Height          =   285
      Left            =   420
      TabIndex        =   51
      Top             =   1695
      Width           =   2745
      _ExtentX        =   4842
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
      Left            =   3180
      TabIndex        =   50
      Top             =   1695
      Width           =   7020
      _ExtentX        =   12383
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
   Begin MSComctlLib.ListView LV 
      Height          =   3990
      Left            =   435
      TabIndex        =   4
      Top             =   2505
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   7038
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
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
         Text            =   "S #"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Location"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Remarks"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Lot No."
         Object.Width           =   2469
      EndProperty
   End
   Begin VB.TextBox txtQtyToAdd 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10380
      TabIndex        =   49
      Top             =   2190
      Width           =   825
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   420
      TabIndex        =   54
      Top             =   2190
      Width           =   9960
   End
   Begin VB.TextBox txtInHand 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10215
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   1695
      Width           =   990
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
      Height          =   255
      Left            =   405
      TabIndex        =   48
      Top             =   9780
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Frame Frame1 
      Caption         =   "Special Instructions :"
      Height          =   630
      Left            =   420
      TabIndex        =   25
      Top             =   6510
      Width           =   8550
      Begin VB.TextBox txtSpecialInstructions 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   60
         TabIndex        =   26
         Top             =   210
         Width           =   8415
      End
   End
   Begin VB.Frame FraChecker 
      Caption         =   "Checker :"
      Height          =   630
      Left            =   4785
      TabIndex        =   46
      Top             =   6510
      Width           =   4200
      Begin MSForms.ComboBox cmbChecker 
         Height          =   285
         Left            =   45
         TabIndex        =   47
         Top             =   240
         Width           =   4080
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "7197;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.TextBox txtLotNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   420
      TabIndex        =   45
      Top             =   1200
      Width           =   885
   End
   Begin VB.TextBox txtExcess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataField       =   "Fax2"
      DataSource      =   "datPrimaryRS"
      Height          =   330
      Left            =   10170
      MaxLength       =   100
      TabIndex        =   29
      Tag             =   "ExcessQtyPercentage"
      Text            =   "100"
      Top             =   6750
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.TextBox txtRcvings 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataField       =   "Fax2"
      DataSource      =   "datPrimaryRS"
      Height          =   330
      Left            =   9015
      MaxLength       =   100
      TabIndex        =   28
      Tag             =   "MaximumRcvingsAgainstPO"
      Text            =   "50"
      Top             =   6750
      Visible         =   0   'False
      Width           =   1110
   End
   Begin MSComCtl2.DTPicker DTReturn 
      Height          =   285
      Left            =   9960
      TabIndex        =   24
      Top             =   7410
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   129892355
      CurrentDate     =   37711
   End
   Begin ComboList.Usercontrol1 cmbReturnProcessOld 
      Height          =   285
      Left            =   120
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4845
      Visible         =   0   'False
      Width           =   5280
      _ExtentX        =   9313
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
   Begin VB.TextBox txtVendRate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8445
      TabIndex        =   22
      Top             =   7410
      Width           =   765
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Height          =   255
      Left            =   1110
      TabIndex        =   20
      Top             =   7050
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdCheckAll 
      Caption         =   "Check &All"
      Height          =   255
      Left            =   30
      TabIndex        =   19
      Top             =   7050
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtQty 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E3FBFB&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   9195
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   7410
      Width           =   765
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
      Height          =   600
      Left            =   6075
      TabIndex        =   7
      Top             =   9495
      Width           =   5280
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3480
         TabIndex        =   16
         Top             =   165
         Width           =   1665
      End
      Begin VB.CommandButton cmdSaveClose 
         Caption         =   "&Save && Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1770
         TabIndex        =   15
         Top             =   165
         Width           =   1665
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save && &New   "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   60
         TabIndex        =   14
         Top             =   165
         Width           =   1665
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   555
      Top             =   -195
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
            Picture         =   "frmMakerIssuanceFromSF.frx":0000
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakerIssuanceFromSF.frx":01BC
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -15
      Top             =   -195
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakerIssuanceFromSF.frx":0378
            Key             =   "Active"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakerIssuanceFromSF.frx":07CC
            Key             =   "Inactive"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   285
      Left            =   3645
      TabIndex        =   13
      Top             =   1200
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   129957891
      CurrentDate     =   37711
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   165
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3765
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   503
      _Version        =   393216
      CheckBox        =   -1  'True
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   129957891
      CurrentDate     =   37711
   End
   Begin VB.Frame FraLoad 
      Height          =   3045
      Left            =   420
      TabIndex        =   17
      Top             =   4020
      Visible         =   0   'False
      Width           =   10860
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Loading Data"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2295
         Left            =   90
         TabIndex        =   18
         Top             =   780
         Width           =   10605
      End
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
      Left            =   585
      TabIndex        =   0
      Top             =   0
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Semi Finish Issuance"
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
         Index           =   0
         Left            =   3780
         TabIndex        =   1
         Top             =   120
         Width           =   3195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Semi Finish Issuance"
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
         Index           =   1
         Left            =   3795
         TabIndex        =   2
         Top             =   135
         Width           =   3195
      End
   End
   Begin VB.Frame FraStockOrder 
      Caption         =   "Article for Stock Order Issuance :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3105
      Left            =   630
      TabIndex        =   38
      Top             =   3900
      Visible         =   0   'False
      Width           =   10485
      Begin VB.TextBox txtSOQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9615
         TabIndex        =   43
         Top             =   465
         Width           =   765
      End
      Begin ComboList.Usercontrol1 cmbSOItems 
         Height          =   285
         Left            =   75
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   465
         Width           =   9510
         _ExtentX        =   16775
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
      Begin VB.CommandButton cmdSOOK 
         Caption         =   "Ok"
         Height          =   435
         Left            =   7680
         TabIndex        =   41
         Top             =   840
         Width           =   1320
      End
      Begin VB.CommandButton cmdSOCancel 
         Caption         =   "Cancel"
         Height          =   435
         Left            =   9060
         TabIndex        =   40
         Top             =   840
         Width           =   1320
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Qty :"
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
         Left            =   9600
         TabIndex        =   44
         Top             =   255
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Article :"
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
         Left            =   75
         TabIndex        =   42
         Top             =   255
         Width           =   645
      End
   End
   Begin VB.Frame FraSkip 
      Caption         =   "Skip to :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3045
      Left            =   3195
      TabIndex        =   32
      Top             =   2265
      Visible         =   0   'False
      Width           =   5040
      Begin ComboList.Usercontrol1 cmbSkipProcess 
         Height          =   285
         Left            =   1005
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Width           =   3915
         _ExtentX        =   6906
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
      Begin VB.CommandButton cmdSkipCancel 
         Caption         =   "Cancel"
         Height          =   435
         Left            =   2535
         TabIndex        =   37
         Top             =   2760
         Width           =   1320
      End
      Begin VB.CommandButton cmdSkipOk 
         Caption         =   "Ok"
         Height          =   435
         Left            =   1170
         TabIndex        =   36
         Top             =   2760
         Width           =   1320
      End
      Begin VB.Label lblSkipDesc 
         Caption         =   "This will make the selected item to skip the current process && make it Issuable to the selected Process."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   645
         Left            =   210
         TabIndex        =   35
         Top             =   1365
         Width           =   4710
      End
      Begin VB.Label Label6A 
         AutoSize        =   -1  'True
         Caption         =   "Process :"
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
         Left            =   225
         TabIndex        =   33
         Top             =   270
         Width           =   750
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Hidden"
      Height          =   1305
      Left            =   480
      TabIndex        =   57
      Top             =   3315
      Visible         =   0   'False
      Width           =   11355
      Begin VB.CommandButton cmdUpdateRate 
         Cancel          =   -1  'True
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1185
         TabIndex        =   79
         Top             =   450
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "POST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   465
         Picture         =   "frmMakerIssuanceFromSF.frx":0C20
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   255
         UseMaskColor    =   -1  'True
         Width           =   600
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   750
         Left            =   240
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   210
         Visible         =   0   'False
         Width           =   10950
         _ExtentX        =   19315
         _ExtentY        =   1323
         _Version        =   393216
         Tabs            =   4
         Tab             =   2
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Vender Balances"
         TabPicture(0)   =   "frmMakerIssuanceFromSF.frx":0D22
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "LVBal"
         Tab(0).Control(1)=   "txtTotal"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Vender History"
         TabPicture(1)   =   "frmMakerIssuanceFromSF.frx":0D3E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label5(2)"
         Tab(1).Control(1)=   "Label5(3)"
         Tab(1).Control(2)=   "LVMaker"
         Tab(1).Control(3)=   "txtRate"
         Tab(1).Control(4)=   "txtTotalQtyHistory"
         Tab(1).Control(5)=   "txtTotalValueHistory"
         Tab(1).ControlCount=   6
         TabCaption(2)   =   "Issued Items"
         TabPicture(2)   =   "frmMakerIssuanceFromSF.frx":0D5A
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Label5(0)"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Label5(1)"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "LVIssued"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "txtRate1"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "txtTotalValue"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "txtTotalQty"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "chkSampleProvided"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "chkDrawingProvided"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "chkForgingProvided"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).Control(9)=   "chkSteelProvided"
         Tab(2).Control(9).Enabled=   0   'False
         Tab(2).ControlCount=   10
         TabCaption(3)   =   "Other Vender's History"
         TabPicture(3)   =   "frmMakerIssuanceFromSF.frx":0D76
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "LVMakerOther"
         Tab(3).ControlCount=   1
         Begin VB.TextBox txtTotalValueHistory 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   -67785
            Locked          =   -1  'True
            TabIndex        =   69
            Top             =   3525
            Width           =   1140
         End
         Begin VB.TextBox txtTotalQtyHistory 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   -69825
            Locked          =   -1  'True
            TabIndex        =   68
            Top             =   3525
            Width           =   1080
         End
         Begin VB.CheckBox chkSteelProvided 
            Appearance      =   0  'Flat
            Caption         =   "Steel Provided"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5250
            TabIndex        =   67
            Top             =   3480
            Width           =   1605
         End
         Begin VB.CheckBox chkForgingProvided 
            Appearance      =   0  'Flat
            Caption         =   "Forging Provided"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3525
            TabIndex        =   66
            Top             =   3480
            Width           =   1605
         End
         Begin VB.CheckBox chkDrawingProvided 
            Appearance      =   0  'Flat
            Caption         =   "Drawing Provided"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1800
            TabIndex        =   65
            Top             =   3480
            Width           =   1605
         End
         Begin VB.CheckBox chkSampleProvided 
            Appearance      =   0  'Flat
            Caption         =   "Sample Provided"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   75
            TabIndex        =   64
            Top             =   3480
            Width           =   1605
         End
         Begin VB.TextBox txtTotalQty 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7695
            Locked          =   -1  'True
            TabIndex        =   63
            Top             =   3435
            Width           =   1080
         End
         Begin VB.TextBox txtTotalValue 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   9735
            Locked          =   -1  'True
            TabIndex        =   62
            Top             =   3435
            Width           =   1140
         End
         Begin VB.TextBox txtTotal 
            Height          =   285
            Left            =   -65475
            TabIndex        =   61
            Top             =   3540
            Width           =   1590
         End
         Begin VB.TextBox txtRate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -70065
            TabIndex        =   60
            Top             =   630
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.TextBox txtRate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2145
            TabIndex        =   59
            Top             =   1575
            Visible         =   0   'False
            Width           =   1350
         End
         Begin MSComctlLib.ListView LVIssued 
            Height          =   3030
            Left            =   75
            TabIndex        =   70
            Top             =   360
            Width           =   10800
            _ExtentX        =   19050
            _ExtentY        =   5345
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
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
               Text            =   "S No."
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Vender"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Return Process"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Buyer"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Order No"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Item Code"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Return Date"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "Rate"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Text            =   "Issued Qty"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Value"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "Priority"
               Object.Width           =   1764
            EndProperty
         End
         Begin MSComctlLib.ListView LVMaker 
            Height          =   3120
            Left            =   -74925
            TabIndex        =   71
            Top             =   375
            Width           =   10770
            _ExtentX        =   18997
            _ExtentY        =   5503
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
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
            NumItems        =   10
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S No."
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Process"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Item Code"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Item Description"
               Object.Width           =   4586
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Date"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Qty"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Rate"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Value"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Del. Date"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Priority"
               Object.Width           =   1940
            EndProperty
         End
         Begin MSComctlLib.ListView LVBal 
            Height          =   3075
            Left            =   -74940
            TabIndex        =   72
            Top             =   360
            Width           =   10785
            _ExtentX        =   19024
            _ExtentY        =   5424
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
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
               Text            =   "S No."
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Process"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Item Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Item Description"
               Object.Width           =   7761
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Balance"
               Object.Width           =   2117
            EndProperty
         End
         Begin MSComctlLib.ListView LVMakerOther 
            Height          =   3450
            Left            =   -74925
            TabIndex        =   73
            Top             =   375
            Width           =   10770
            _ExtentX        =   18997
            _ExtentY        =   6085
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
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
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Maker"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Item Code"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Item Description"
               Object.Width           =   4586
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Date"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Qty"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Rate"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Value"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Del. Date"
               Object.Width           =   1940
            EndProperty
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Total Value :"
            Height          =   195
            Index           =   3
            Left            =   -68700
            TabIndex        =   77
            Top             =   3555
            Width           =   900
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Total Qty :"
            Height          =   195
            Index           =   2
            Left            =   -70635
            TabIndex        =   76
            Top             =   3555
            Width           =   780
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Total Qty :"
            Height          =   195
            Index           =   1
            Left            =   6885
            TabIndex        =   75
            Top             =   3465
            Width           =   780
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Total Value :"
            Height          =   195
            Index           =   0
            Left            =   8820
            TabIndex        =   74
            Top             =   3465
            Width           =   900
         End
      End
   End
   Begin MSForms.ComboBox cmbIssOrderNo 
      Height          =   285
      Left            =   9240
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1950
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3440;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label6 
      Height          =   240
      Left            =   420
      TabIndex        =   55
      Top             =   1965
      Width           =   10770
      BackColor       =   11517387
      Caption         =   $"frmMakerIssuanceFromSF.frx":0D92
      Size            =   "18997;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label4_1 
      Height          =   240
      Left            =   420
      TabIndex        =   53
      Top             =   1470
      Width           =   10770
      BackColor       =   11517387
      Caption         =   $"frmMakerIssuanceFromSF.frx":0E6F
      Size            =   "18997;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Excess Qty %"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   10185
      TabIndex        =   31
      Top             =   4830
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Max. Rcvings"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   4
      Left            =   9015
      TabIndex        =   30
      Top             =   4830
      Width           =   960
   End
   Begin MSForms.ComboBox cmbReturnProcess 
      Height          =   285
      Left            =   405
      TabIndex        =   27
      Top             =   7410
      Width           =   2880
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5080;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbOrderNo 
      Height          =   285
      Left            =   165
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4320
      Visible         =   0   'False
      Width           =   1965
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3466;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label3 
      Height          =   240
      Left            =   420
      TabIndex        =   11
      Top             =   7185
      Width           =   10860
      BackColor       =   11517387
      Caption         =   $"frmMakerIssuanceFromSF.frx":0F49
      Size            =   "19156;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cmbDesc 
      Height          =   285
      Left            =   165
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4035
      Visible         =   0   'False
      Width           =   3300
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5821;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbItems 
      Height          =   285
      Left            =   5055
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1200
      Width           =   4185
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "7382;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   420
      TabIndex        =   3
      Top             =   975
      Width           =   10770
      BackColor       =   11517387
      Caption         =   $"frmMakerIssuanceFromSF.frx":100F
      Size            =   "18997;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Menu mnuPop 
      Caption         =   "Pop up Menu"
      Begin VB.Menu mnuSkipProcess 
         Caption         =   "Skip Process"
      End
      Begin VB.Menu mnuStockOrderIssuance 
         Caption         =   "Stock Order Issuance"
      End
   End
End
Attribute VB_Name = "frmMakerIssuanceFromSF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyCol As New Collection
Dim MyCol1 As New Collection
Dim IsItemsFirst As Boolean
Dim IssIDs As New Collection
Dim rsMain As New ADODB.Recordset
Dim IsLotProcess As Boolean
Dim GotIssID As Long
Dim RcvingLot As Boolean
Dim UserReqAuth As Integer 'Whether Or Not User Required Authorization On Issue For This Process.

Dim lVendBal As Long
Dim iColNo As Integer
Dim bFromMaker As Boolean
Dim bFromProcess As Boolean
Dim bReIssuance As Boolean
Dim bDontUpdateFromProcess As Boolean
Dim bDontUpdateFromMaker As Boolean
Dim lCheckerProcessNo As Long
Dim bLotNoWiseEntry As Boolean
Dim bIssueAgainstOpening As Boolean
Private Sub AddItems()
    
    On Error GoTo err
    Exit Sub
    Dim OrCond As String
    'Create In List
    
    'For i = 1 To MyCol.Count
    '    OrCond = OrCond & "'" & MyCol(i).CompItemCode & "',"
    'Next i
    
    If Len(InList) > 0 Then
        InList = Mid$(InList, 1, Len(InList) - 1)
    End If
    'InList = InList & ")"
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select * From VVendAssItems", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            For i = 1 To MyCol.count
                If MyCol(i).CompItemCode = ![ITemID] Then
                    cmbItems.AddItem ![ITemID] 'MyCol(i).ItemCode
                    cmbItems.List(cmbItems.ListCount - 1, 1) = MyCol(i).ItemCode
                    cmbItems.List(cmbItems.ListCount - 1, 2) = MyCol(i).lQty
                    cmbItems.List(cmbItems.ListCount - 1, 3) = ![Rate]
                    cmbItemDesc.AddItem MyCol(i).ItemDescription
                    Exit For
                End If
            Next i
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub cmbItem_matched()
    cmbItemName.ID = cmbItem.ID
End Sub

Private Sub cmbItemName_matched()
    cmbItem.ID = cmbItemName.ID
End Sub

Private Sub chkIssByLotNo_Click()

    Dim bLock As Boolean
    If chkIssByLotNo.Value = vbChecked Then
        LV.ListItems.Clear
        bLock = False
        bLotNoWiseEntry = True
    Else
        bLock = True
        bLotNoWiseEntry = False
    End If
    
    txtLotNo.Enabled = Not bLock
    cmbProcess.Enabled = bLock
    cmbItems.Enabled = bLock
'    cmbStore.Enabled = bLock
'    cmbLocation.Enabled = bLock
    txtInHand.Enabled = bLock
    'txtQtyToAdd.Locked = Not bLock
    
    If chkIssByLotNo.Value = vbChecked Then
        txtQtyToAdd.SetFocus
    End If
    
End Sub

Private Sub cmbDesc_Change()
    Call cmbDesc_Click
End Sub

Private Sub cmbDesc_Click()

    If cmbDesc.MatchFound = False Then Exit Sub
    cmbItems.ListIndex = cmbDesc.ListIndex
    
    
    If RcvingLot Then
        LV.ListItems(cmbDesc.ListIndex + 1).Selected = True
    End If
    
End Sub

Private Sub cmbItems_Change()
    Call cmbItems_Click
End Sub

Private Sub cmbItems_Click()
    
    LV.ListItems.Clear
    If Not cmbItems.MatchFound Then Exit Sub
    
    If cmbProcess.MatchFound = False Then
    
        cmbProcess.ClearVals
        cmbProcess.AddVals con, "DISTINCT Description", "VStockOrderOpening", "ProcessID", " WHERE ItemID='" & cmbItems.Text & "' AND Qty>IssdQty"
       
    End If
    If cmbProcess.MatchFound = False Then Exit Sub
    'cmbDesc.ListIndex = cmbItems.ListIndex
    'txtVendRate = Val(cmbItems.List(cmbItems.ListIndex, 1))
'    If bReIssuance Then
'        Call FillList(cmbProcess.ID, cmbVender.ID, cmbItems.Text)
'    Else
'        Call FillList
'    End If
    
    '/\/\/\/\/\/\/\/\/\/
'    cmbIssOrderNo.Clear
'    If strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
'        Call AddToCombo(cmbIssOrderNo, "OrderNo + ' (' + CAST(OrderQty AS VARCHAR(50)) + ')'", "VFOrderItems_Grouped", " WHERE CompItemCode='" & cmbItems.Text & "' ORDER BY MinDeliveryDT", , "OrderNo")
'    Else
'        'Call AddToCombo(cmbIssOrderNo, "DISTINCT FOrderItems.OrderNo+ ' ('+InternalRefNo+')'", "FOrderItems INNER JOIN FCustomerOrders ON FOrderitems.OrderNo=FCustomerOrders.OrderNo", " WHERE CompItemCode='" & cmbItems.Text & "'", , "FOrderItems.OrderNo")
'        'Call AddToCombo(cmbIssOrderNo, "OrderNo+ ' ('+InternalRefNo+')'", "VrptOrders_ForProduction", " WHERE CompItemID='" & cmbItems.Text & "'", , "OrderNo")
'    End If
    
    Dim i As Integer, bStockOrderAdded As Boolean
    bStockOrderAdded = False
    For i = 0 To cmbIssOrderNo.ListCount - 1
        If cmbIssOrderNo.List(i, 1) = "Stock-Order" Then
            bStockOrderAdded = True
            cmbIssOrderNo.ListIndex = i
        End If
    Next
    
'    If bStockOrderAdded = False Then
'        cmbIssOrderNo.AddItem "Stock-Order"
'        cmbIssOrderNo.List(cmbIssOrderNo.ListCount - 1, 1) = "Stock-Order"
'        cmbIssOrderNo.ListIndex = cmbIssOrderNo.ListCount - 1
'    End If
        
    Dim rs As New ADODB.Recordset, lSNo As Long

'    If strCompany = "QEL" Then
'        txtInHand = GetSingleLongValue("RcvdQty-IssQty", "VendRcvdDetail INNER JOIN VStockOrderOpening ON VendRcvdDetail.Opening_RefID=VStockOrderOpening.EntryID" _
'         , " WHERE ItemID='" & cmbItems.Text & "' AND RcvdQty-IssQty>0 AND VStockOrderOpening.ProcessID=" & cmbProcess.ID & " AND LotNo='0'")
'
'        'txtRemarks = GetSingleStringValue("Remarks", "VendRcvdDetail INNER JOIN VStockOrderOpening ON VendRcvdDetail.Opening_RefID=VStockOrderOpening.EntryID" _
'         , " WHERE ItemID='" & cmbItems.Text & "' AND RcvdQty-IssQty>0 AND VStockOrderOpening.ProcessID=" & cmbProcess.ID)
'    Else
'        If bLotNoWiseEntry Then
'            cmbStore.ClearVals
'            cmbStore.AddVals con, "DISTINCT StoreName", "VendRcvdDetail INNER JOIN VStockOrderOpening ON VendRcvdDetail.Opening_RefID=VStockOrderOpening.EntryID" _
'             , "Store_RefID", "WHERE VStockOrderOpening.EntryID IN(SELECT SOO_RefID FROM LotTransferDetails WHERE EntryID=" & Val(txtLotNo.Tag) & ") AND Store_RefID IN(SELECT Store_RefID FROM Users_Stores WHERE UserID=" & UserID & ")"
'        Else
'            cmbStore.ClearVals
'            cmbStore.AddVals con, "DISTINCT StoreName", "VendRcvdDetail INNER JOIN VStockOrderOpening ON VendRcvdDetail.Opening_RefID=VStockOrderOpening.EntryID" _
'             , "Store_RefID", "WHERE ItemCode='" & cmbItems.Text & "' AND (Qty-IssdQty)>0 AND VStockOrderOpening.ProcessID=" & cmbProcess.ID & "  AND Store_RefID IN(SELECT Store_RefID FROM Users_Stores WHERE UserID=" & UserID & ")" '& " AND LotNo='0'"
'        End If
'    End If
    
    With rs
        .Open "SELECT SNO FROM ItemProcesses WHERE ProcessID=" & cmbProcess.ID & " AND ItemID='" & cmbItems.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            lSNo = Val(!SNo & "")
        End If
        .Close
    End With
    
'    If (cmbProcess.ID = iReadyFinishProcessID) Then
'        rs.Open "SELECT Rate,Unit FROM VendAssItems WHERE VendID=" & cmbVender.ID & " AND ProcessID=" & cmbProcess.ID & " AND ItemID='" & cmbItems.Text & "'", con, adOpenForwardOnly, adLockReadOnly
'        With cmbReturnProcess
'            .Clear
'            .AddItem cmbProcess.Text '"Ready Finish"
'            .List(0, 1) = cmbProcess.ID
'            .List(0, 2) = Val(rs!Rate & "")
'            .List(0, 3) = rs!Unit & ""
'        End With
'        cmbReturnProcess.ListIndex = 0
'    Else
'        Dim lNextProcessID As Long
'        lNextProcessID = getNextProcessID(cmbItems.Text, cmbProcess.ID)
'
'        cmbReturnProcess.Clear
'        'Now Check If Item is assigned to this process to selected vender.
'        With rs
'            .Open "SELECT VendAssItems.*,Description FROM VendAssItems INNER JOIN Processes ON VendAssItems.ProcessID=Processes.ProcessID WHERE Processes.ProcessID=" & lNextProcessID & " AND ItemID='" & cmbItems.Text & "' AND VendID=" & cmbVender.ID, con, adOpenForwardOnly, adLockReadOnly
'            If .EOF = False Then
'                cmbReturnProcess.AddItem rs!Description & ""
'                cmbReturnProcess.List(cmbReturnProcess.ListCount - 1, 1) = lNextProcessID 'Val(rs!ProcessID & "")
'                cmbReturnProcess.List(cmbReturnProcess.ListCount - 1, 2) = Val(rs!Rate & "")
'                cmbReturnProcess.List(cmbReturnProcess.ListCount - 1, 3) = rs!Unit & ""
'            End If
'            .Close
'        End With
'
'        If cmbReturnProcess.ListCount > 0 Then cmbReturnProcess.ListIndex = 0
'    End If
    Dim lNextProcessID As Long
    lNextProcessID = getNextProcessID(cmbItems.Text, cmbProcess.ID)

    cmbReturnProcess.Clear
    'Now Check If Item is assigned to this process to selected vender.
    With rs
        .Open "SELECT Processes.ProcessID,Description FROM Processes INNER JOIN ItemProcesses ON Processes.ProcessID=ItemProcesses.ProcessID WHERE ItemID='" & cmbItems.Text & "' ORDER BY ItemProcesses.SNo", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            cmbReturnProcess.AddItem rs!Description & ""
            cmbReturnProcess.List(cmbReturnProcess.ListCount - 1, 1) = Val(rs!ProcessID & "")
            'cmbReturnProcess.List(cmbReturnProcess.ListCount - 1, 2) = Val(rs!Rate & "")
            'cmbReturnProcess.List(cmbReturnProcess.ListCount - 1, 3) = rs!Unit & ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    For i = 0 To cmbReturnProcess.ListCount - 1
        If Val(cmbReturnProcess.List(i, 1)) = lNextProcessID Then
            cmbReturnProcess.ListIndex = i
            Exit For
        End If
    Next
    'If cmbReturnProcess.ListCount > 0 Then cmbReturnProcess.ListIndex = 0
    
    Set rs = Nothing
    
    'Call ShowOtherVendersHistory
    
End Sub

Private Sub ShowOtherVendersHistory()

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "Select Top 10 * From VMakerIssItems Where VendID<>" & cmbVender.ID & " AND ProcessID=" & cmbProcess.ID & " AND ItemCode='" & cmbItems & "' Order By DT Desc,EntryID Desc", con, adOpenForwardOnly, adLockReadOnly
        
        LVMakerOther.ListItems.Clear
        'LVBal.ListItems.Clear
        lVendBal = 0
        Do Until .EOF
            Set ITM = LVMakerOther.ListItems.add(, , LVMakerOther.ListItems.count + 1)
            ITM.SubItems(1) = ![VendID1] & ""
            ITM.SubItems(2) = ![ItemCode] & ""
            ITM.SubItems(3) = ![ItemName] & ""
            ITM.SubItems(4) = Format(![DT] & "", "dd-MM-yy")
            ITM.SubItems(5) = Val(![IssQty] & "")
            ITM.SubItems(6) = Val(![Rate] & "")
            ITM.SubItems(7) = Val(!IssValue & "") 'Round(Val(![IssQty] & "") * Val(![Rate] & ""))
            ITM.SubItems(8) = Format(![ReturnDT] & "", "dd-MM-yy")
            
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub FillList(Optional iProcessID As Integer = 0, Optional iMakerID As Integer = 0, Optional strItemID As String = "", Optional strLotNo As String = "")

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
    If bReIssuance = False Then
        If cmbItems.MatchFound = False Then Exit Sub
        
        With cmd
            Set .ActiveConnection = con
            .CommandType = adCmdStoredProc
            .CommandText = "IssueItems_SP"
            .Parameters("@ProcessID").Value = cmbProcess.ID
            .Parameters("@ItemID").Value = cmbItems.Text
            .Parameters("@SFIssuance").Value = 1
            Set rs = .Execute
        End With
    Else
        With cmd
            Set .ActiveConnection = con
            .CommandType = adCmdStoredProc
            .CommandText = "ReIssueItemsNew_SP"
            .Parameters("@ProcessID").Value = iProcessID
            .Parameters("@MakerID").Value = iMakerID
            .Parameters("@ItemID").Value = strItemID
            If strLotNo <> "" Then
                .Parameters("@LotNo").Value = strLotNo
            End If
            Set rs = .Execute
        End With
    End If
    
    Dim lQty As Long
    
    Dim bUpdateItemsCombo As Boolean
    If cmbVender.MatchFound And cmbItems.MatchFound = False Then
        bUpdateItemsCombo = True
    Else
        bUpdateItemsCombo = False
    End If
    
    
'    With rs
'        If bUpdateItemsCombo Then cmbItems.Clear
'
'        LV.ListItems.Clear
'        Do Until .EOF
'
'            .MoveNEXT
'        Loop
'        .Close
'    End With
'    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub



Private Sub cmbProcessOld_matched()
    
    On Error GoTo err
'    If cmbProcess.ID = 9 Then
'        cmdCheckAll.Visible = True
'        cmdClearAll.Visible = True
'    Else
'        cmdCheckAll.Visible = False
'        cmdClearAll.Visible = False
'    End If
    
    Dim cmd As New ADODB.Command
    
    LV.ListItems.Clear
    LVIssued.ListItems.Clear
    LVMaker.ListItems.Clear
    
    IsLotProcess = LotProcess(cmbProcess.ID)
    
    If IsLotProcess And RcvLot(cmbProcess.ID) Then
        RcvingLot = True
        cmbItems.Visible = False
        DTPicker1.Visible = True
        Label2.Caption = "  Process                                    Date                                      Lot No"
        LV.Checkboxes = True
        
        txtQty.Visible = False
        Label3.Caption = "  Venders                                                                Date"
        Label3.Width = 4830
        cmdPost.Left = DTPicker2.Left + DTPicker2.Width + 20
    Else
        RcvingLot = False
        cmbItems.Visible = True
        DTPicker1.Visible = False
        Label2.Caption = "  Process                                    Item Code                        Item Description"
        LV.Checkboxes = False
        
        txtQty.Visible = True
        Label3.Caption = "  Venders                                                                Date                      Quantity"
        Label3.Width = 5970
        cmdPost.Left = txtQty.Left + txtQty.Width + 20
        Me.Refresh
    End If
    
    Dim LColor As Long
    Dim ITM As ListItem
    
    With cmd
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        
        
        If IsLotProcess Then
            
            If RcvingLot Then
                LV.ColumnHeaders.add 2, , "Lot No.", 1200
                'LV.ColumnHeaders(2).Text = "Lot No/Order No"
                .CommandText = "SP_IssueItemsLots"
                txtQty.Locked = True
                txtRate1.Locked = True
            Else
                LV.ColumnHeaders(3).Text = "Order No"
                .CommandText = "SP_IssueItems"
                txtQty.Locked = False
                txtRate1.Locked = False
            End If
            cmbProcess.Tag = "Lot No"
        Else
            LV.ColumnHeaders(3).Text = "Order No"
            .CommandText = "SP_IssueItems"
            cmbProcess.Tag = ""
        End If
        
        If rsMain.State = 1 Then rsMain.Close
        rsMain.CursorType = adOpenForwardOnly
        rsMain.LockType = adLockReadOnly
        
        .Parameters("@ProcessID").Value = cmbProcess.ID
            
        Set rsMain = .Execute
    End With
    
    Dim isAdded As Boolean
    
    With rsMain
        cmbItems.Clear
        cmbDesc.Clear
        
        If RcvingLot Then
'            DTPicker1.value = ServerDate
            Do Until .EOF
                isAdded = False
                For i = 0 To cmbItems.ListCount - 1
                    If Trim(cmbItems.List(i)) = Trim(![RecieptID] & "") Then
                        isAdded = True
                        Exit For
                    End If
                Next i
                
                If Not isAdded Then
                    cmbItems.AddItem ![RecieptID] & ""
                    cmbDesc.AddItem ![RecieptID] & ""
                End If
                    
                
                Set ITM = LV.ListItems.add(, ![LotNo] & "''", LV.ListItems.count + 1)
                ITM.Tag = ![RecieptID] & ""
                'Debug.Print LV.ListItems.Count
                'ITM.Tag = ![EntryID]
                'itm.ListSubItems.Add , , ![CustCode] & ""
                ITM.ListSubItems.add(, , ![RecieptID] & "").Tag = ![LotNo] & ""
                
                ITM.ListSubItems.add(, , ![ItemCode] & "").Tag = ![ItemCode] & ""
                
                'itm.ListSubItems.Add , , ![ItemCode] & ""
                ITM.ListSubItems.add , , ![ItemDescription] & ""
                ITM.ListSubItems.add(, , Val(![IssQty] & "")).Tag = Val(![IssQty] & "")
                
                If IsNull(![LockedItemCode]) Then
                    LColor = LV.ForeColor
                Else
                    LColor = vbWhite
                End If
                
                ITM.ForeColor = LColor
                For i = 1 To ITM.ListSubItems.count
                    ITM.ListSubItems(i).ForeColor = LColor
                Next
                
                .MoveNEXT
            Loop
        Else
            Do Until .EOF
                isAdded = False
                For i = 0 To cmbItems.ListCount - 1
                    If Trim(cmbItems.List(i)) = Trim(![ItemCode] & "") Then
                        isAdded = True
                        Exit For
                    End If
                Next i
                
                If Not isAdded Then
                    cmbItems.AddItem ![ItemCode] & ""
                    cmbDesc.AddItem ![ItemDescription] & ""
                End If
                
                
            
                Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            
            
                'ITM.Tag = ![EntryID]
                
                ITM.ListSubItems.add , , ![ItemCode] & ""
                ITM.ListSubItems.add , , ![ItemDescription] & ""
                ITM.ListSubItems.add(, , Val(![IssQty] & "")).Tag = Val(![IssQty] & "")
                
                If IsNull(![LockedItemCode]) Then
                    LColor = LV.ForeColor
                Else
                    LColor = RGB(150, 55, 150)
                End If
                
                ITM.ForeColor = LColor
                For i = 1 To ITM.ListSubItems.count
                    ITM.ListSubItems(i).ForeColor = LColor
                Next
                .MoveNEXT
            Loop
        End If
    End With
    
    If LV.ListItems.count > 0 Then
        LV.ListItems(1).Selected = True
        Call LV_ItemClick(LV.ListItems(1))
    End If
    
    Me.Show
    cmbVender.SetFocus
    SendKeys "a"
    SendKeys Chr(8)
    
    FraLoad.Visible = False
    LV.Visible = True
    'Debug.Print Time
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbLocation_matched()
    
    If bLotNoWiseEntry Then
        txtInHand = GetSingleLongValue("RcvdQty-IssQty", "VendRcvdDetail INNER JOIN VStockOrderOpening ON VendRcvdDetail.Opening_RefID=VStockOrderOpening.EntryID" _
         , " WHERE VStockOrderOpening.EntryID IN(SELECT SOO_RefID FROM LotTransferDetails WHERE EntryID=" & Val(txtLotNo.Tag) & ")")
        
        txtQtyToAdd = Val(txtInHand)
        txtRemarks = GetSingleStringValue("Remarks", "VendRcvdDetail INNER JOIN VStockOrderOpening ON VendRcvdDetail.Opening_RefID=VStockOrderOpening.EntryID" _
        , " WHERE VStockOrderOpening.EntryID IN(SELECT SOO_RefID FROM LotTransferDetails WHERE EntryID=" & Val(txtLotNo.Tag) & ")")
     
    Else
        txtInHand = GetSingleLongValue("Qty-IssdQty", "VendRcvdDetail INNER JOIN VStockOrderOpening ON VendRcvdDetail.Opening_RefID=VStockOrderOpening.EntryID" _
         , " WHERE VStockOrderOpening.EntryID=" & cmbLocation.ID)
         
         txtRemarks = GetSingleStringValue("Remarks", "VendRcvdDetail INNER JOIN VStockOrderOpening ON VendRcvdDetail.Opening_RefID=VStockOrderOpening.EntryID" _
        , " WHERE VStockOrderOpening.EntryID=" & cmbLocation.ID)
    End If
    
End Sub

Private Sub cmbProcess_matched()

'    If bLotNoWiseEntry Then
'        Call AddToCombo(cmbItems, "ItemCode", "VendRcvdDetail INNER JOIN LotTransferDetails ON LotTransferDetails.VRD_To_RefID=VendRcvdDetail.EntryID", " WHERE LotTransferDetails.EntryID=" & Val(txtLotNo.Tag))
'        cmbItems.ListIndex = 0
'    Else
'        'Call AddToCombo(cmbItems, "DISTINCT ItemID", "VStockOrderOpening", " WHERE ProcessID=" & cmbProcess.ID & " AND (Qty-SplityQty)>IssdQty")
'        Dim strItemID As String
'        strItemID = cmbItems.Text
'
'        If strItemID = "" Then
'            Call AddToCombo(cmbItems, "DISTINCT ItemID", "VStockOrderOpening", " WHERE ProcessID=" & cmbProcess.ID & " AND Qty>IssdQty")
'        Else
'           cmbItems_Change
'        End If
''        If strItemID <> "" Then
''            cmbItems.Text = strItemID
''        End If
'
'    End If
End Sub

Private Sub cmbReturnProcess_Change()
    Call cmbReturnProcess_Click
End Sub


Private Sub cmbReturnProcess_Click()
    
    If cmbReturnProcess.MatchFound = False Then Exit Sub
    
    cmbVender.ClearVals
    cmbVender.AddVals con, "'{' + VendID1 + '} ' + VenderName", "VMakerAssItems", "VendID", " WHERE ProcessID=" & cmbReturnProcess.List(cmbReturnProcess.ListIndex, 1) & " AND ItemID='" & cmbItems.Text & "'"
    
    Exit Sub
    '-------------------------------------------------------------------------------
    
End Sub

Private Sub cmbStore_matched()

'    cmbLocation.ClearVals
'
'    If bLotNoWiseEntry Then
'        cmbLocation.AddVals con, "RackNo + ' {' + ShelfNo + '} ' + '{' + CAST((RcvdQty-IssQty) AS VARCHAR(50)) + '}'", "VendRcvdDetail INNER JOIN VStockOrderOpening ON VendRcvdDetail.Opening_RefID=VStockOrderOpening.EntryID" _
'         , "Shelf_RefID", " WHERE VStockOrderOpening.EntryID IN(SELECT SOO_RefID FROM LotTransferDetails WHERE EntryID=" & Val(txtLotNo.Tag) & ")"
'    Else
'        'cmbLocation.AddVals con, "RackNo + ' {' + ShelfNo + '} ' + '{' + CAST((RcvdQty-IssQty) AS VARCHAR(50)) + '}' + '(' + LotNo + ')'", "VendRcvdDetail INNER JOIN VStockOrderOpening ON VendRcvdDetail.Opening_RefID=VStockOrderOpening.EntryID" _
'         ,  "Shelf_RefID", " WHERE ItemID='" & cmbItems.Text & "' AND RcvdQty-IssQty>0 AND VStockOrderOpening.ProcessID=" & cmbProcess.ID & " AND Store_RefID=" & cmbStore.ID & ""
'         cmbLocation.AddVals con, "RackNo + ' {' + ShelfNo + '} ' + '{' + CAST((Qty-IssdQty) AS VARCHAR(50)) + '}' + '(' + LotNo + ')'", "VendRcvdDetail INNER JOIN VStockOrderOpening ON VendRcvdDetail.Opening_RefID=VStockOrderOpening.EntryID" _
'         , "VStockOrderOpening.EntryID", " WHERE ItemID='" & cmbItems.Text & "' AND Qty-IssdQty>0 AND VStockOrderOpening.ProcessID=" & cmbProcess.ID & " AND Store_RefID=" & cmbStore.ID & ""
'    End If
    
End Sub

Private Sub cmbVender_matched()
    
    Dim dRate As Double
    dRate = GetSingleDoubleValue("Rate", "VMakerAssItems", "WHERE VendID=" & cmbVender.ID & " AND ItemID='" & cmbItems.Text & "' AND ProcessID=" & cmbReturnProcess.List(cmbReturnProcess.ListIndex, 1))
    '------------------------------------------
    txtVendRate = dRate
        
End Sub
Private Sub ShowVenderHistory()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim lTotalQty As Long
    Dim dTotalValue As Double
    With rs
        .Open "Select Top 10 * From VMakerIssItems WHERE VendID=" & cmbVender.ID & " Order By DT Desc,EntryID Desc", con, adOpenForwardOnly, adLockReadOnly
        
        LVMaker.ListItems.Clear
        LVBal.ListItems.Clear
        lVendBal = 0
        lTotalQty = 0
        dTotalValue = 0
        Do Until .EOF
            Set ITM = LVMaker.ListItems.add(, , LVMaker.ListItems.count + 1)
            ITM.SubItems(1) = !ProcessName & ""
            ITM.SubItems(2) = ![ItemCode] & ""
            ITM.SubItems(3) = ![ItemName] & ""
            ITM.SubItems(4) = Format(![DT] & "", "dd-MM-yy")
            ITM.SubItems(5) = Val(![IssQty] & "")
            lTotalQty = lTotalQty + Val(![IssQty] & "")
            ITM.SubItems(6) = Val(![Rate] & "")
            ITM.SubItems(7) = Val(!IssValue & "") 'Round(Val(![IssQty] & "") * Val(![Rate] & ""))
            dTotalValue = dTotalValue + Val(!IssValue & "")
            ITM.SubItems(8) = Format(![ReturnDT] & "", "dd-MM-yy")
            ITM.SubItems(9) = Val(!Priority & "")
            .MoveNEXT
        Loop
        .Close
        txtTotalQtyHistory = lTotalQty
        txtTotalValueHistory = dTotalValue
        'If cmbVender.ID <> 129 Then
            .Open "Select ProcessName,ItemCode,ItemName,Sum(RemQty) As RemQty From VMakerIssItems Where VendID=" & cmbVender.ID & "  AND RemQty>0 Group By ProcessName,ItemCode,ItemName Having Sum(RemQty)>0", con, adOpenForwardOnly, adLockReadOnly
            
            Do Until .EOF
                Set ITM = LVBal.ListItems.add(, , LVBal.ListItems.count + 1)
                ITM.SubItems(1) = !ProcessName & ""
                ITM.SubItems(2) = ![ItemCode] & ""
                ITM.SubItems(3) = ![ItemName] & ""
                ITM.SubItems(4) = Val(![RemQty] & "")
                lVendBal = lVendBal + Val(![RemQty] & "")
                .MoveNEXT
            Loop
            .Close
            txtTotal = lVendBal
        'End If
        
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbVender_Validate(cancel As Boolean)
    'cmdPost.Default = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    con.Execute "Delete From PrintSession"
End Sub

Private Sub cmdCheckAll_Click()
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmdClearAll_Click()
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmdPost_Click()

    On Error GoTo err
'    If LV.SelectedItem.ForeColor = 9844630 Then
'        MsgBox "Can't Issue. Item Is Locked", vbInformation
'        Exit Sub
'    End If
        
    If Not cmbVender.MatchFound Then Exit Sub
    If cmbVender.ID = "" Or cmbVender.Text = "" Then Exit Sub
    If Val(txtVendRate) <= 0 And cmbVender.ID <> iFactoryMakerID Then
        MsgBox "Rate Can't be Zero!", vbInformation
        Exit Sub
    End If
    If cmbReturnProcess.MatchFound = False Then Exit Sub
    'cmbVender.Enabled = False
    
    'Check For Max Balance Limit For This Vender.......
    If cmbVender.ID <> 129 Then
        Dim rs As New ADODB.Recordset
        With rs
'            .Open "Select MaxLimit From Venders Where VendID=" & cmbVender.ID, con, adOpenForwardOnly, adLockReadOnly
'            If lVendBal >= Val(![MaxLimit] & "") Then
'                MsgBox "Sorry, The Selected Vender Has Reached His Max. Issuence Limit." & _
'                 vbNewLine & "Can't Post It For This Vender." & vbNewLine & "Vender Maximum Limit : " & Val(![MaxLimit] & "") & vbNewLine & "Vender Curr. Balance  : " & lVendBal, vbInformation
'
'                .Close
'                Exit Sub
'            End If
'            .Close
        End With
        Set rs = Nothing
    End If
    
    If RcvingLot Then
        PostForLots
        Exit Sub
    End If
    
    If Val(txtQty) < 1 Then Exit Sub
    
    If bReIssuance Then
        If Val(txtQty) > Val(LV.SelectedItem.SubItems(5)) Then
            MsgBox "Invalid Qty", vbInformation
            Exit Sub
        End If
    End If
    
    If GotIssID = 0 Then
        'Dim IssID As String
        'For i = 1 To LVIssued.ListItems.Count
        '    If Val(cmbVender.ID) = Val(LVIssued.ListItems(i).ListSubItems(2).Key) Then
        '        IssID = LVIssued.ListItems(i).ListSubItems(2)
        '        GoTo GotID:
        '    End If
        'Next i
    
        GotIssID = Val(con.Execute("Select Max(Cast(right(RecieptID,Len(RecieptID)-10) As Int)) From VMakerIssItems Where Convert(varchar,DT,6)='" & Format(DTPicker2, "dd MMM yy") & "' and Left(RecieptID,3)<>'HSR'").Fields(0).Value & "")
        GotIssID = GotIssID + 1
    Else
        GotIssID = GotIssID + 1
    End If
    'IssID = "ISS-" & Val(LVMaker.SelectedItem.ListSubItems(1).Key) & "-" & _
     Format(ServerDate, "ddMMyy") & "-" & Format(Val(IssID) + 1, "000")
    IssID = "ISU-" & _
     Format(DTPicker2, "ddMMyy") & GotIssID
    
    Dim strBuyerCode As String
    strBuyerCode = "" 'con.Execute("SELECT CustCode FROM FCUstomerOrders WHERE OrderNo='" & LV.SelectedItem.SubItems(2) & "'").Fields(0).Value & ""
    
    Dim strOrderNo As String
    If cmbIssOrderNo.MatchFound = False Then
        strOrderNo = ""
    Else
        strOrderNo = cmbIssOrderNo.List(cmbIssOrderNo.ListIndex, 1)
    End If
    
GotID:
        
    Dim ITM As ListItem
    With ITM
        'Set ITM = LVIssued.ListItems.Add(, Val(LV.SelectedItem.Key) & Val(LVMaker.SelectedItem.Key) & "'", LVIssued.ListItems.Count + 1)
        'Set ITM = LVIssued.ListItems.Add(, LV.SelectedItem.Key & "|" & LVIssued.ListItems.Count + 1, LVIssued.ListItems.Count + 1)
        Set ITM = LVIssued.ListItems.add(, , LVIssued.ListItems.count + 1)
        ITM.Tag = LV.SelectedItem.key
        'itm.ListSubItems.Add , cmbVender.ID & "'", cmbVender.Text
        ITM.ListSubItems.add(, , cmbVender.Tag).Tag = cmbVender.ID
        ITM.ListSubItems.add(, , cmbReturnProcess.Text).Tag = IssID & ""  'DTPicker2
        ITM.ListSubItems.add(, , strBuyerCode).Tag = cmbReturnProcess.List(cmbReturnProcess.ListIndex, 1) 'DTPicker2
        ITM.ListSubItems.add(, , strOrderNo).Tag = cmbReturnProcess.List(cmbReturnProcess.ListIndex, 1)
        ITM.ListSubItems.add(, , cmbItems.Text).Tag = txtSpecialInstructions.Text
        ITM.ListSubItems.add , , Format(DTReturn.Value, "dd-MMM-yyyy") 'LV.SelectedItem.SubItems(3)
        ITM.ListSubItems.add(, , Val(txtVendRate.Text)).Tag = cmbProcess.ID
        ITM.ListSubItems.add , , Val(txtQty) 'LV.SelectedItem.SubItems(3)
        ITM.ListSubItems.add , , GetValueofItem(cmbItems.Text)   'Val(txtQty) * Val(txtVendRate.Text)
        ITM.ListSubItems.add(, , LVIssued.ListItems.count).Tag = LV.SelectedItem.ListSubItems(2).Tag 'Assign From Rcvd_RefID to (10).tag
        'ITM.ListSubItems.Add , , Val(txtQty)
        If cmbProcess.Tag = "Lot No" Then
            ITM.ListSubItems(6).Tag = LV.SelectedItem.SubItems(3) & ""
        End If
        
        
        Dim bReWorkLot As Boolean, iRepairType As Integer
        
        With rs
            .Open "SELECT ReWorkLot,Repair_RefID FROM VendRcvdDetail WHERE LotNo=" & Val(txtLotNo) & " AND NextProcessID=" & cmbProcess.ID, con, adOpenForwardOnly, adLockReadOnly
            If .EOF = False Then
                bReWorkLot = !ReWorklot
                iRepairType = Val(!Repair_RefID & "")
            Else
                bReWorkLot = False
                iRepairType = 0
            End If
            .Close
        End With
        Set rs = Nothing
        ITM.ListSubItems(8).Tag = bReWorkLot
        ITM.ListSubItems(9).Tag = iRepairType
    End With
    Call getTotalValue
    
    'MyCol(LV.SelectedItem.Index).IssueableQty = MyCol(LV.SelectedItem.Index).IssueableQty - Val(txtQty)
    'LV.SelectedItem.SubItems(5) = Val(LV.SelectedItem.SubItems(5)) - Val(txtQty)
     
    txtRate.Visible = False
    Dim i As Integer
    For i = 1 To LVIssued.ListItems.count
        If LVIssued.ListItems(i).SubItems(5) = LV.SelectedItem.SubItems(3) Then
            LVIssued.ListItems(i).ListSubItems(5).Tag = txtSpecialInstructions.Text
        End If
    Next
    'Removing Item From Combos
'    If LV.SelectedItem.SubItems(4) = 0 Then
'        Dim ItemIndex As Long
'        ItemIndex = cmbItems.ListIndex
'        If ItemIndex <> -1 Then
'
'            cmbItems.RemoveItem ItemIndex
'            cmbDesc.RemoveItem ItemIndex
'            cmbItems.Text = ""
'            cmbDesc.Text = ""
'            If cmbItems.ListCount >= ItemIndex And ItemIndex > 0 Then
'                cmbItems.ListIndex = ItemIndex
'            ElseIf cmbItems.ListCount > 0 Then
'                cmbItems.ListIndex = 0
'            End If
'        End If
'    End If
    'if cmbitems.
    'Call LV_ItemClick(LV.SelectedItem)
    
    txtQty = ""
    
    'Select The Next Item Automaticaly....
    If LV.SelectedItem.Index < LV.ListItems.count Then
        If Val(LV.SelectedItem.SubItems(4)) = 0 Then
            LV.ListItems(LV.SelectedItem.Index + 1).Selected = True
        End If
        Call LV_ItemClick(LV.ListItems(LV.SelectedItem.Index))
    End If
    
    'cmbVender.SetFocus
    
    Exit Sub
err:
    MsgBox err.Description
        'End If
End Sub

Private Function GetValueofItem(strItemID As String) As Double

    On Error GoTo err
    If txtAssignedUnit = "Kgs" Then
        Dim dWeight As Double
        dWeight = GetSingleDoubleValue("ForgingWeight", "Items", " WHERE ItemID='" & strItemID & "'")
        GetValueofItem = ((dWeight * Val(txtQty)) / 1000) * Val(txtVendRate.Text)
    Else
        GetValueofItem = Val(txtQty) * Val(txtVendRate.Text)
    End If
    
    Exit Function
err:
    MsgBox err.Description
End Function

Private Sub getTotalValue()

    Dim i As Integer, lTotal As Long, lQty As Long
    For i = 1 To LVIssued.ListItems.count
        lTotal = lTotal + Val(LVIssued.ListItems(i).SubItems(9))
        lQty = lQty + Val(LVIssued.ListItems(i).SubItems(8))
    Next
    
    txtTotalValue = lTotal
    txtTotalQty = lQty
    
End Sub

Private Sub cmdSave_Click()

    If Saved Then
'        If chkPrint.Value = vbChecked Then
'            Call PrintIssuanceSlips
'            DoEvents
'        End If
        FormProcessID = cmbProcess.ID
        'Unload Me
        'Me.Show
        'cmbProcess.ID = cmbProcess.Tag
        'Call ShowIssForm(FormProcessID)
        'frmVendIssItemsByMaker.Show
        LVIssued.ListItems.Clear
        con.Execute "Delete From PrintSession"
        Call cmbItems_Click
    End If
    
End Sub

Private Sub cmdSaveClose_Click()
    
    If Saved Then
'        If chkPrint.Value = vbChecked Then
'            Call PrintIssuanceSlips
'            DoEvents
'        End If
'        If cmbProcess.ID <= 7 Then
'            If Not UserReqAuth Then Call PrintIssRcvRpts("Original", True)
'        End If
        Unload Me
        con.Execute "Delete From PrintSession"
    End If
End Sub

Private Sub PrintIssuanceSlips()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report, strSelection As String
    If bReIssuance Then
        Set rpt = rptApp.OpenReport(RptPath & "\IssSlipGP.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\IssSlip.rpt")
    End If
    strSelection = "ISNULL({PrintSession.RecieptNo})=False"
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdUnAssign_Click()

    On Error GoTo err
    If LV.ListItems.count = 0 Then Exit Sub
    Dim iAffect As Integer
    
    con.Execute "Delete From VendAssItems Where EntryID=" & Val(LV.SelectedItem.key), iAffect
    If iAffect > 0 Then
        'Call cmbNick_Click
        cmdUnAssign.Visible = False
    End If
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub Command1_Click()

End Sub

Private Sub cmdSkipCancel_Click()
    FraSkip.Visible = False
End Sub

Private Sub cmdSkipOk_Click()

    If cmbSkipProcess.MatchFound = False Then
        MsgBox "Please Select Process.", vbInformation
        Exit Sub
    End If
    
    Dim strItemID As String
    
    strItemID = LV.SelectedItem.SubItems(3)
    
    con.Execute "UPDATE VendRcvdDetail SET NextProcessID=" & cmbSkipProcess.ID & " WHERE ItemCode='" & _
     strItemID & "' AND RcvdQty>IssQty AND NextProcessID=" & cmbProcess.ID
     
    Call cmbItems_Click
    FraSkip.Visible = False
    
End Sub

Private Sub cmdSOCancel_Click()
    FraStockOrder.Visible = False
End Sub

Private Sub cmdSOOK_Click()

    If cmbSOItems.MatchFound = False Then
        MsgBox "Please Select Article.", vbInformation
        Exit Sub
    ElseIf Val(txtSOQty) <= 0 Then
        MsgBox "Invalid Qty.", vbInformation
        Exit Sub
    End If
    
    '-----------------------------------------
    'Check If Opening for Selected Article for Selected Process is feeded ?
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(ItemID)", "StockOrderOpening", " WHERE ItemID='" & cmbSOItems.ID & "' AND ProcessID=" & cmbProcess.ID)
'    If lCount > 0 Then
'        MsgBox "This Item has already Opening feeded at this Process.", vbInformation
'        Exit Sub
'    End If
    
    Call StartTrans(con)
    
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
    lCount = GetSingleLongValue("COUNT(CompItemID)", "FCustomerCatalog", " WHERE Custcode='Stock' AND Country='PK' AND CompItemID='" & cmbSOItems.ID & "'")
    If lCount = 0 Then
        con.Execute "INSERT INTO FCustomerCatalog(Custcode,Country,ItemID,CompItemID) VALUES('Stock','PK','" & cmbSOItems.ID & "','" & cmbSOItems.ID & "')"
    End If
    'Now Check If Article is feeded against this Article
    lCount = 0
    lCount = GetSingleLongValue("COUNT(OrderNo)", "FOrderItems", " WHERE OrderNo='Stock-Order' AND CompItemCode='" & cmbSOItems.ID & "'")
    If lCount = 0 Then
    
        If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "PMT" Or strCompany = "Kami" Or strCompany = "Towne" Then
            con.Execute "INSERT INTO FOrderItems(OrderNo,ItemCode,CompItemCode,Qty,DeliveryDT,Stamps,Quality) VALUES('Stock-Order','" & cmbSOItems.ID & _
             "','" & cmbSOItems.ID & "'," & Val(txtSOQty) & ",'" & Date & "','','')"
        Else
            con.Execute "INSERT INTO FOrderItems(OrderNo,ItemCode,CompItemCode,Qty) VALUES('Stock-Order','" & cmbSOItems.ID & _
             "','" & cmbSOItems.ID & "'," & Val(txtSOQty) & ")"
        End If
         
    Else
        con.Execute "UPDATE FOrderItems SET Qty=Qty+" & Val(txtSOQty) & " WHERE OrderNo='Stock-Order' AND CompItemCode='" & cmbSOItems.ID & "'"
    End If
    
    'Now Check If Article is feeded in Stock Order
    lCount = 0
    If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "PMT" Or strCompany = "Kami" Or strCompany = "Towne" Then
        lCount = GetSingleLongValue("MAX(EntryID)", "StockOrderOpening", " WHERE Shelf_RefID=0 AND ItemID='" & cmbSOItems.ID & "' AND ProcessID=" & Val(cmbProcess.ID))
    Else
        lCount = GetSingleLongValue("MAX(EntryID)", "StockOrderOpening", " WHERE ItemID='" & cmbSOItems.ID & "' AND ProcessID=" & cmbProcess.ID)
    End If
    
    If lCount = 0 Then
        Dim strLocation As String, lQty As Long, lEntryID As Long
        strLocation = "": lQty = Val(txtSOQty)
        
        If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "PMT" Or strCompany = "Kami" Or strCompany = "Towne" Then
            con.Execute "INSERT INTO StockOrderOpening(ItemID,ProcessID,Qty,Location,UserName,MachineName,Shelf_RefID,Remarks) VALUES('" & cmbSOItems.ID & _
             "'," & cmbProcess.ID & "," & lQty & ",'" & strLocation & "','" & CurrentUserName & "','" & strComputerName & "',0,'')"
        Else
            con.Execute "INSERT INTO StockOrderOpening(ItemID,ProcessID,Qty,Location,UserName,MachineName) VALUES('" & cmbSOItems.ID & _
             "'," & cmbProcess.ID & "," & lQty & ",'" & strLocation & "','" & CurrentUserName & "','" & strComputerName & "')"
        End If
        
        lEntryID = GetSingleLongValue("MAX(EntryID)", "StockOrderOpening")
    Else
        lEntryID = lCount
        
        con.Execute "UPDATE StockOrderOpening SET Qty=Qty+" & Val(txtSOQty) & " WHERE EntryID=" & lEntryID
    End If
    
    Dim lRcvEntryID As Long, lPrevProcessID As Long
    
    
    con.Execute "INSERT INTO VendReceived(VendID,DT,RecieptID,UserID,ProcessID,Issuance_RefID) VALUES(" & _
     cmbVender.ID & ",'" & Date & "','',0," & cmbProcess.ID & ",0)"
    
    lRcvEntryID = GetSingleLongValue("MAX(EntryID)", "VendReceived")
    
    lPrevProcessID = getPrevProcessID(cmbSOItems.ID, cmbProcess.ID)
    
    con.Execute "INSERT INTO VendRcvdDetail(RefID,ItemCode,RecieptID,RcvdQty,Rate,LotNo,NextProcessID,OrderNo" & _
     ",CountedBy,Issue_RefID,ProcessID,RcvdWeight,Opening_RefID) VALUES(" & lRcvEntryID & ",'" & cmbSOItems.ID & _
     "',''," & lQty & ",0,'0'," & cmbProcess.ID & ",'Stock-Order','',0," & lPrevProcessID & _
     ",0," & lEntryID & ")"
     
    con.CommitTrans
    '-----------------------------------------
    
    FraStockOrder.Visible = False
    Call cmbVender_matched
    
    Exit Sub
    
    
err:
    MsgBox err.Description
    
End Sub

Private Sub DTPicker1_Change()
    Call DTPicker1_Click
End Sub

Private Sub DTPicker1_Click()
    On Error GoTo err
    
    
    With rsMain
    
        If DTPicker1.Value & "" = "" Then
            .Filter = adFilterNone
        Else
            .Filter = "RecieptID Like 'HS-" & Format(DTPicker1, "ddMMyy") & "-*'"
            '.Filter = "Left(RecieptID,9) ='HS-" & Format(DTPicker1, "ddMMyy") & "'"
        End If
                
        Dim isAdded As Boolean
        
        cmbItems.Clear
        cmbDesc.Clear
        LV.ListItems.Clear
        Do Until .EOF
            isAdded = False
            For i = 0 To cmbItems.ListCount - 1
                If Trim(cmbItems.List(i)) = Trim(![RecieptID] & "") Then
                    isAdded = True
                    Exit For
                End If
            Next i
            
            If Not isAdded Then
                cmbItems.AddItem ![RecieptID] & ""
                cmbDesc.AddItem ![RecieptID] & ""
            End If
            .MoveNEXT
        Loop
    End With
    
    
        
    
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub DTPicker2_Change()
    DTReturn.Value = DateAdd("d", 30, DTPicker2.Value)
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Exit Sub 'For Now
    
    LV.Sorted = True
    If LV.SortKey = ColumnHeader.Index - 1 Then
        If LV.SortOrder = lvwAscending Then
            LV.SortOrder = lvwDescending
            ColumnHeader.Icon = "Up"
        Else
            LV.SortOrder = lvwAscending
            ColumnHeader.Icon = "Down"
        End If
    Else
        LV.ColumnHeaderIcons = Nothing
        Set LV.ColumnHeaderIcons = ImageList2
        LV.SortKey = ColumnHeader.Index - 1
        LV.SortOrder = lvwAscending
        ColumnHeader.Icon = "Down"
    End If

    On Error Resume Next
    LV.SelectedItem.EnsureVisible

End Sub


Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
    If LV.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure ?", vbYesNo + vbQuestion) = vbYes Then
            LV.ListItems.Remove LV.SelectedItem.Index
            Call GetIssQtyTotal
        End If
    End If
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Exit Sub
    If Button <> 2 Then Exit Sub
    
    If bReIssuance = True Then
        mnuSkipProcess.Visible = True
        mnuStockOrderIssuance.Visible = False
    Else
        mnuStockOrderIssuance.Visible = True
        mnuSkipProcess.Visible = False
    End If
    
    If strCompany = "Weldon Instruments" Or strCompany = "QEL" Or strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
        Me.PopUpMenu mnuPop
    End If
    
End Sub

Private Sub LVIssued_DblClick()
    
    
    If LVIssued.ListItems.count = 0 Then Exit Sub
    'If cmbCust.MatchFound = False Then Exit Sub
    If LV.ColumnHeaders(3).Text = "Lot No/Order No" And RcvingLot Then Exit Sub
     
    With txtRate1
        'cmdPost.Default = False
        .Move LVIssued.Left + LVIssued.ColumnHeaders(iColNo).Left, LVIssued.Top + LVIssued.SelectedItem.Top, LVIssued.ColumnHeaders(iColNo).Width
        .Text = LVIssued.SelectedItem.SubItems(iColNo - 1) 'LVMaker.SelectedItem.SubItems(3)
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
        
    End With

End Sub

Private Sub LVIssued_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo err
    
    If KeyCode = 46 Then
    
'        If RcvingLot Then
'            LVIssued.ListItems.Remove LVIssued.SelectedItem.Index
'            Exit Sub
'        End If
        
        'MyCol(LVIssued.SelectedItem.Tag).IssueableQty = MyCol(LVIssued.SelectedItem.Tag).IssueableQty + Val(LVIssued.SelectedItem.SubItems(7))
        LV.ListItems(LVIssued.SelectedItem.Tag).ListSubItems(4) = LV.ListItems(LVIssued.SelectedItem.Tag).ListSubItems(4).Tag
        'LV.ListItems(LVIssued.SelectedItem.Tag).SubItems(5) = MyCol(LVIssued.SelectedItem.Tag).IssueableQty
        LVIssued.ListItems.Remove LVIssued.SelectedItem.Index
        Call LV_ItemClick(LV.SelectedItem)
        
        'Reset The IssID's
        Dim TempID As Long
        TempID = 0
        If LVIssued.ListItems.count > 0 Then
            For i = LVIssued.SelectedItem.Index To LVIssued.ListItems.count
                TempID = Val(Right(LVIssued.ListItems(i).SubItems(2), 3)) - 1
                If TempID > 1 Then
                    LVIssued.ListItems(i).SubItems(2) = "ISU-" & _
                     Format(ServerDate, "ddMMyy") & "-" & Format(TempID, "000")
                End If
            Next i
        Else
            GotIssID = 0
            
        End If
        If GotIssID > 0 Then GotIssID = GotIssID - 1
        Call getTotalValue
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub LVIssued_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If x > LVIssued.ColumnHeaders(10).Left Then
        iColNo = 10
    ElseIf x > LVIssued.ColumnHeaders(9).Left Then
        iColNo = 9
    Else
        iColNo = 8
    End If
End Sub

Private Sub LvMaker_DblClick()
    
    Exit Sub
    If LVMaker.ListItems.count = 0 Then Exit Sub
    'If cmbCust.MatchFound = False Then Exit Sub

    
    With txtRate
        .Move LVMaker.Left + LVMaker.ColumnHeaders(4).Left, LVMaker.Top + LVMaker.SelectedItem.Top, LVMaker.ColumnHeaders(4).Width
        '.Text = MyCol(LV.SelectedItem.SubItems(2) & "|" & LV.SelectedItem.SubItems(3)).IssueableQty 'LVMaker.SelectedItem.Tag 'LVMaker.SelectedItem.SubItems(3)
        .Text = MyCol(LV.SelectedItem.key).IssueableQty 'LVMaker.SelectedItem.Tag 'LVMaker.SelectedItem.SubItems(3)
        .Visible = True
        
        If LV.ColumnHeaders(3).Text = "Lot No" Then
            If RcvingLot Then
                .Locked = True
            Else
                .Locked = False
            End If
        Else
            .Locked = False
        End If
        
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With

    

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    
    On Error GoTo err
    mnuPop.Visible = False
'    If strCompany <> "Dr-Frgz" Then
'        chkUseOriginalLotNo.Visible = False
'    End If
    'Call AddToCombo(cmbOrderNo, "OrderNo", "FCustomerOrders", " Order By OrderNo")
    
'    If strCompany = "Dr-Frgz" Then
'        chkIssByLotNo.Visible = True
'    Else
'        chkIssByLotNo.Visible = False
'    End If
    
    bReIssuance = False
    
    'cmbCust.ListHeight = 2600
    'cmbCustName.ListHeight = 2600
    'cmbOrders.ListHeight = 2600
    
    'cmbProcess.Enabled = True
    bFromMaker = False
    cmbProcess.ListHeight = 6000

    DTPicker2.Value = ServerDate
    Call DTPicker2_Change
    DTPicker1.Value = ServerDate
    DTPicker1.Value = ""
    DTReturn.Value = DateAdd("d", 30, DTPicker2.Value)
    SSTab1.Tab = 2
    IsLotProcess = False
    cmbVender.ListHeight = 2000
    GotIssID = 0

    cmbSkipProcess.ListHeight = 1800
 
    lCheckerProcessNo = Val(getGeneralDataValue("CheckerProcessNo"))
    bLotNoWiseEntry = False
    Me.Hide
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Public Sub LockProc(ProcID As Long)
    
    
    cmbProcess.Tag = ProcID
    
    'cmbProcess.Enabled = False
    'Me.Refresh
    
    cmbProcess.ID = ProcID
    Me.Show
    
End Sub
Public Sub RefreshLV()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "Select *  from FCustomerCatalog Where CustCode='" & cmbCust.ID & "' and Country='" & cmbCountry.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, ![CustCode] & ![Country] & ![ITemID], ![ITemID] & "")
            ITM.SubItems(1) = ![Description] & ""
            ITM.SubItems(2) = ![Unit] & ""
            ITM.SubItems(3) = ![FOB] & ""
            ITM.SubItems(4) = ![CIFSea] & ""
            ITM.SubItems(5) = ![CIFAir] & ""
            ITM.SubItems(6) = ![CnFSea] & ""
            ITM.SubItems(7) = ![CnFAir] & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Resize()
    
    Exit Sub
    On Error Resume Next
    
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
    
    Dim TitleBottom As Long

    TitleBottom = DTPicker1.Top + DTPicker1.Height + 50 '+ fTitle.Height
 
    
    LV.Move Label2.Left, TitleBottom
    
    Label3.Move Label2.Left, LV.Top + LV.Height + 50
'    cmbVender.Move Label3.Left, Label3.Top + Label3.Height - 15
    
    
    DTPicker2.Move cmbVender.Left + cmbVender.Width + 20, Label3.Top + Label3.Height - 15
    txtQty.Move DTPicker2.Left + DTPicker2.Width + 10, Label3.Top + Label3.Height - 15
    cmdPost.Move txtQty.Left + txtQty.Width + 10, Label3.Top
    SSTab1.Move LV.Left, cmbVender.Top + cmbVender.Height + 50
    
End Sub

Private Sub FillVendList()

    On Error GoTo err
    
    If LV.ListItems.count = 0 Then Exit Sub
    Dim VendIDList As String
    VendIDList = ""
    For i = 1 To LVIssued.ListItems.count
        With LVIssued.ListItems(i)
            'If LV.SelectedItem.key = .Tag Then
            '    VendIDList = VendIDList & Val(.ListSubItems(1).key) & ","
            'End If
            If LV.SelectedItem.ListSubItems(1) = .SubItems(3) Then
                VendIDList = VendIDList & Val(.ListSubItems(1).Tag) & ","
            End If
        End With
    Next i
    
    If VendIDList <> "" Then
        VendIDList = "(" & Left(VendIDList, Len(VendIDList) - 1) & ")"
    Else
        VendIDList = "(0)"
    End If
    
    'With rs
    '    .Open "Select VendID,Rate,VenderName From VVendAssItems Where ItemID='" & LV.SelectedItem.SubItems(3) & "' and PhaseID=" & cmbProcess.ID & " and VendID Not In" & VendIDList, con, adOpenForwardOnly, adLockReadOnly
    '    LVMaker.ListItems.Clear
    '    Do Until .EOF
    '
    '        'Set ITM = LVMaker.ListItems.Add(, ![VendID] & "'", LVMaker.ListItems.Count + 1)
    '        Set ITM = LVMaker.ListItems.Add(, , LVMaker.ListItems.Count + 1)
    '        ITM.Tag = Val(MyCol(Item.Index).IssueableQty & "")
    '        ITM.ListSubItems.Add , ![VendID] & "'", .Fields(2) & ""
    '        ITM.ListSubItems.Add , , Val(![Rate] & "")
    '
    '        ITM.ListSubItems.Add(, , 0).Tag = LV.SelectedItem.SubItems(2) & ""
    '        .MoveNext
    '    Loop
    '    .Close
    'End With
    Call cmbVender.ClearVals
    'If RcvingLot Then
    '    Call cmbVender.AddVals(con, "VenderName", "VVendAssItems", "VendID", " Where ItemID='" & LV.SelectedItem.SubItems(3) & "' and (PhaseID=" & cmbProcess.ID & " Or PhaseID=-1) ")
    'Else
    If HicoVisible And RcvingLot Then
        Call cmbVender.AddVals(con, "Distinct VenderName", "VMakerAssItems", "VendID", " Where (ItemID='" & LV.SelectedItem.SubItems(2) & "' and PhaseID=" & cmbProcess.ID & "  AND Active=1 and VendID Not In" & VendIDList & " and VendType=1 Or (PhaseID=-1))")
    Else
        Call cmbVender.AddVals(con, "Distinct VenderName", "VMakerAssItems", "VendID", " Where (ItemID='" & LV.SelectedItem.SubItems(1) & "' AND Active=1 and (PhaseID=" & cmbProcess.ID & ") and VendID Not In" & VendIDList & " ) OR VendID=129 ")
    End If
    'End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Exit Sub
    txtQty = LV.SelectedItem.ListSubItems(5)
        
    Dim i As Integer
    For i = 1 To LVIssued.ListItems.count
        If LV.SelectedItem.SubItems(3) = LVIssued.ListItems(i).SubItems(5) Then
            txtSpecialInstructions = LVIssued.ListItems(i).ListSubItems(5).Tag
            Exit For
        End If
    Next
    
    Dim bReWorkLot As Boolean, iRepairType As Integer
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT ReWorkLot,Repair_RefID FROM VendRcvdDetail WHERE LotNo=" & Val(txtLotNo) & " AND NextProcessID=" & cmbProcess.ID, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            bReWorkLot = !ReWorklot
            iRepairType = Val(!Repair_RefID & "")
        End If
        .Close
    End With
    
    
    If bReWorkLot = False Then
    
        Set rs = con.Execute("EXEC NextProcessesForIssuance_SP " & Val(cmbProcess.ID) & "," & Val(cmbVender.ID) & ",'" & Item.SubItems(3) & "'")
        
        cmbReturnProcess.Clear
        Do Until rs.EOF
            cmbReturnProcess.AddItem rs!Description & ""
            cmbReturnProcess.List(cmbReturnProcess.ListCount - 1, 1) = Val(rs!ProcessID & "")
            cmbReturnProcess.List(cmbReturnProcess.ListCount - 1, 2) = Val(rs!Rate & "")
            cmbReturnProcess.List(cmbReturnProcess.ListCount - 1, 3) = rs!Unit & ""
            rs.MoveNEXT
        Loop
    Else
        cmbReturnProcess.Clear
        cmbReturnProcess.AddItem cmbProcess.Text
        cmbReturnProcess.List(0, 1) = cmbProcess.ID
        Dim strUnit As String, dRate As Double
        
        With rs
            .Open "SELECT Unit,Rate FROM VendAssItems WHERE VendID=" & Val(cmbVender.ID) & " AND ItemID='" & Item.SubItems(3) & "' AND ProcessID=" & Val(cmbProcess.ID), con, adOpenForwardOnly, adLockReadOnly
            If .EOF = False Then
                strUnit = !Unit & ""
                dRate = Val(!Rate & "")
            Else
                strUnit = ""
                dRate = 0
            End If
            .Close
        End With
        cmbReturnProcess.List(0, 2) = dRate
'        cmbReturnProcess.List(0, 3) = strUnit
    End If
    If cmbReturnProcess.ListCount > 0 Then cmbReturnProcess.ListIndex = 0
    FraSkip.Visible = False
    FraStockOrder.Visible = False
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub



Private Sub mnuSkipProcess_Click()
    ''
    If LV.ListItems.count = 0 Then Exit Sub
    Dim lSNo As Long, strItemID As String
    strItemID = LV.SelectedItem.SubItems(3)
    lSNo = GetSingleLongValue("SNo", "ItemProcesses", " WHERE ItemID='" & strItemID & "' AND ProcessID=" & cmbProcess.ID)
    
    cmbSkipProcess.ClearVals
    cmbSkipProcess.AddVals con, "Description", "Processes", "ProcessID", " WHERE ProcessID IN(SELECT ProcessID FROM ItemProcesses WHERE ItemID='" & strItemID & "' AND SNo>" & lSNo & ") Order By SNo"
    
    FraSkip.Visible = True
    
End Sub

Private Sub mnuStockOrderIssuance_Click()
        ''
     
    cmbSOItems.ClearVals
    Dim strCondition As String
    strCondition = " WHERE InActive=0 AND ItemID IN(SELECT ItemID FROM VendAssItems WHERE VendID=" & Val(cmbVender.ID) & " AND ProcessID=" & Val(cmbProcess.ID) & ") ORDER BY ItemName"
    Call cmbSOItems.AddVals(con, "ItemID + ' ' + ItemName + ' { ' + ISNULL(Attributes,'') + ' }'", "Items", "ItemID", strCondition)
    
    FraStockOrder.Visible = True
    
End Sub

Private Sub txtLotNo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        'First Find out Process ID & Item Code of given LotNo.
        Dim rs As New ADODB.Recordset, bRecFound As Boolean
        Dim strItemCode As String, iProcessID As Integer
        
        With rs
            '.Open "SELECT TOP 1 LotTransferDetails.EntryID,LotNo,RcvdQty,NextProcessID,ItemCode,VendRcvdDetail.ProcessID,LotTransferDetails.FromOrderNo,T1.Shelf_RefID,LotTransferDetails.SOO_RefID,LotTransferDetails.SplitQty FROM LotTransferDetails INNER JOIN VendRcvdDetail ON LotTransferDetails.VRD_From_RefID=VendRcvdDetail.EntryID INNER JOIN StockOrderOpening T1 ON T1.EntryID=LotTransferDetails.SOO_RefID LEFT JOIN (SELECT SOO_RefID,SUM(QtyIssd) AS QtyIssd FROM StockOrderOpening_Issuance GROUP BY SOO_RefID) TSOOIss ON T1.EntryID=TSOOIss.SOO_RefID WHERE LotNo='" & txtLotNo & "' AND SplitQty>ISNULL(TSOOIss.QtyIssd,0)", con, adOpenForwardOnly, adLockReadOnly
            .Open "SELECT LotNo_Manual AS LotNo,EntryID,Qty-IssdQty AS RcvdQty,ProcessID,ItemID AS ItemCode,'' AS FromOrderNo,Shelf_RefID,EntryID AS SOO_RefID,Qty-IssdQty AS SplitQty,Remarks FROM VStockOrderOpening WHERE LotNo_Manual='" & txtLotNo & "' AND Qty-IssdQty>0", con, adOpenForwardOnly, adLockReadOnly
            If .EOF = False Then
                bRecFound = True
                txtLotNo.Tag = Val(!EntryID & "")
                iProcessID = Val(!ProcessID & "")
                cmbProcess.ID = iProcessID
                
                Call AddToCombo(cmbItems, "ItemID", "Items", " WHERE ItemID='" & !ItemCode & "'")
                cmbItems.ListIndex = 0
                
                Call AddToCombo(cmbIssOrderNo, "OrderNo+ ' ('+InternalRefNo+')'", "VrptOrders_ForProduction", " WHERE CompItemID='" & cmbItems.Text & "' AND OrderNo='" & !FromOrderNo & "'", , "OrderNo")
                
                For i = 0 To cmbIssOrderNo.ListCount - 1
                    If cmbIssOrderNo.List(i, 1) = !FromOrderNo & "" Then
                        cmbIssOrderNo.ListIndex = i
                        Exit For
                    End If
                Next
                If cmbIssOrderNo.MatchFound = False And cmbIssOrderNo.ListCount = 0 Then
                    'Call AddToCombo(cmbIssOrderNo, "OrderNo+ ' ('+ISNULL(InternalRefNo,'')+')'", "VrptOrders_ForProduction", " WHERE CompItemID='" & cmbItems.Text & "'", , "OrderNo")
                    Call AddToCombo(cmbIssOrderNo, "OrderNo+ ' ('+ISNULL(InternalRefNo,'')+')'", "VrptOrders_ForProduction", " WHERE CompItemID='" & cmbItems.Text & "' AND OrderNo IN(SELECT OrderNo FROM VUnshippedOrderList)", , "OrderNo")
                End If
                cmbStore.ClearVals
                cmbStore.AddVals con, "StoreName", "VStoreShelfs", "Store_RefID", " WHERE EntryID=" & Val(!Shelf_RefID & "")
                
                cmbStore.ID = GetSingleLongValue("Store_RefID", "VStoreShelfs", " WHERE EntryID=" & Val(!Shelf_RefID & ""))
                
                cmbLocation.ClearVals
                cmbLocation.AddVals con, "RackNo + ' {' + ShelfNo + '}'", "VStoreShelfs", "EntryID", " WHERE EntryID=" & Val(!Shelf_RefID & "")
                cmbLocation.ID = Val(!Shelf_RefID & "")
                cmbLocation.Tag = Val(!SOO_RefID & "")
                txtQtyToAdd.Text = Val(!SplitQty & "")
                txtRemarks = !Remarks & ""
                Call txtQtyToAdd_KeyPress(13)
            Else
                bRecFound = False
            End If
            .Close
        End With
        Set rs = Nothing
        
        
'        If bRecFound Then
'            cmbProcess.ID = iProcessID
'        Else
'            txtLotNo.Tag = ""
'        End If
        
    End If
    
End Sub

Private Sub txtQty_GotFocus()


    With txtQty
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    

End Sub

Private Sub txtQtyToAdd_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If cmbItems.MatchFound = False Then
            MsgBox "Item Not Selected.", vbInformation
            Exit Sub
        
'        ElseIf Val(txtQtyToAdd) > Val(txtInHand) Then
'            MsgBox "Invalid Qty.", vbInformation
'            Exit Sub
            
        ElseIf cmbLocation.MatchFound = False And strCompany <> "QEL" Then
            MsgBox "Invalid Location.", vbInformation
            Exit Sub
        
        ElseIf cmbIssOrderNo.MatchFound = False Then
            MsgBox "Invalid OrderNo.", vbInformation
            Exit Sub
        End If
        If LV.ListItems.count > 0 Then Exit Sub
        Dim i As Integer
'        For i = 1 To LV.ListItems.count
'            If Val(cmbLocation.ID) = Val(LV.ListItems(i).Tag) Then
'                MsgBox "Already Added.", vbInformation
'                Exit Sub
'            End If
'        Next
        
        Dim strLotNo As String
        'strLotNo = GetSingleStringValue("LotNo", "VendRcvdDetail INNER JOIN VStockOrderOpening ON VStockOrderOpening.EntryID=VendRcvdDetail.Opening_RefID", "WHERE VStockOrderOpening.EntryID=" & cmbLocation.ID)
        strLotNo = txtLotNo
        Dim ITM As ListItem
        Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
        ITM.Tag = Val(cmbLocation.ID)
        ITM.ListSubItems.add(, , cmbLocation.Text).Tag = cmbLocation.Tag 'cmbLocation.ID
        ITM.ListSubItems.add(, , txtRemarks).Tag = ""
        ITM.ListSubItems.add(, , Val(txtQtyToAdd)).Tag = ""
        ITM.ListSubItems.add(, , strLotNo).Tag = "" 'cmbLocation.ID
        
'        If bLotNoWiseEntry Then
'            ITM.ListSubItems.add(, , txtLotNo.Text).Tag = Val(txtLotNo.Tag)
'        End If
'        If LV.ListItems.count > 1 Then
'            chkUseOriginalLotNo.Value = vbUnchecked
'            chkUseOriginalLotNo.Enabled = False
'        Else
'            'Now Check If that specific LotNo is issd. using the Original lot no.
'            Dim lCount As Long
'            lCount = GetSingleLongValue("COUNT(*)", "StockOrderOpening_Issuance INNER JOIN VendIssdDetail ON StockOrderOpening_Issuance.VID_RefID=VendIssdDetail.EntryID", " WHERE StockOrderOpening_Issuance.SOO_RefID=" & cmbLocation.ID & " AND VendIssdDetail.LotNo='" & strLotNo & "'")
'
'            If lCount > 0 Then
'                chkUseOriginalLotNo.Value = vbUnchecked
'                chkUseOriginalLotNo.Enabled = False
'            Else
'                chkUseOriginalLotNo.Value = vbUnchecked
'                chkUseOriginalLotNo.Enabled = True
'            End If
'        End If
        
        Call GetIssQtyTotal
    End If
    
End Sub

Private Function GetIssQtyTotal() As Long

    Dim i As Integer
    Dim lTotal As Long
    For i = 1 To LV.ListItems.count
        lTotal = Val(LV.ListItems(i).SubItems(3)) + lTotal
    Next
    txtQty = lTotal
    
End Function

Private Sub txtRate_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    
    If KeyAscii = 13 Then
        'LVMaker.SelectedItem.SubItems(3) = Val(txtRate)
        
        Dim IssID As String
        For i = 1 To LVIssued.ListItems.count
            If Val(LVMaker.SelectedItem.ListSubItems(1).key) = Val(LVIssued.ListItems(i).ListSubItems(2).key) Then
                IssID = LVIssued.ListItems(i).ListSubItems(2)
                GoTo GotID:
            End If
        Next i
        
        IssID = Val(con.Execute("Select Max(right(RecieptID,3)) From VendIssued Where DT='" & Format(ServerDate, "dd-MMM-yyyy") & "'").Fields(0).Value & "")
        
        'IssID = "ISS-" & Val(LVMaker.SelectedItem.ListSubItems(1).Key) & "-" & _
         Format(ServerDate, "ddMMyy") & "-" & Format(Val(IssID) + 1, "000")
         IssID = "ISU-" & _
         Format(ServerDate, "ddMMyy") & "-" & Format(Val(IssID) + 1, "000")
GotID:
        
        Dim ITM As ListItem
        With ITM
            'Set ITM = LVIssued.ListItems.Add(, Val(LV.SelectedItem.Key) & Val(LVMaker.SelectedItem.Key) & "'", LVIssued.ListItems.Count + 1)
            'Set ITM = LVIssued.ListItems.Add(, LV.SelectedItem.Key & "|" & LVIssued.ListItems.Count + 1, LVIssued.ListItems.Count + 1)
            Set ITM = LVIssued.ListItems.add(, , LVIssued.ListItems.count + 1)
            ITM.Tag = LV.SelectedItem.key
            ITM.ListSubItems.add , LVMaker.SelectedItem.ListSubItems(1).key, LVMaker.SelectedItem.SubItems(1)
            ITM.ListSubItems.add , , IssID & ""
            ITM.ListSubItems.add , , LV.SelectedItem.SubItems(2)
            ITM.ListSubItems.add , , LV.SelectedItem.SubItems(3)
            ITM.ListSubItems.add , , LV.SelectedItem.SubItems(4)
            ITM.ListSubItems.add , , LVMaker.SelectedItem.ListSubItems(2)
            ITM.ListSubItems.add , , Val(txtRate)
            If cmbProcess.Tag = "Lot No" Then
                ITM.ListSubItems(6).Tag = LV.SelectedItem.SubItems(2) & ""
            End If
        End With
        
        'MyCol(LV.SelectedItem.Index).IssueableQty = MyCol(LV.SelectedItem.Index).IssueableQty - Val(txtRate)
        'LV.SelectedItem.ListSubItems(5) = Val(LV.SelectedItem.ListSubItems(5)) - Val(txtRate)
        LV.SelectedItem.SubItems(5) = Val(LV.SelectedItem.SubItems(5)) - Val(txtRate)
        txtRate.Visible = False
        Call LV_ItemClick(LV.SelectedItem)
        'End If
    ElseIf KeyAscii = 27 Then
        txtRate.Visible = False
    End If
    Call OnlyNums(KeyAscii)
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Function Saved() As Boolean

    On Error GoTo err
    If Val(txtRcvings) <= 0 Then
        MsgBox "Please Enter Max. Rcvings.", vbInformation
        Exit Function
    End If
    
    If cmbReturnProcess.MatchFound = False Then
        MsgBox "Please Select Return Process.", vbInformation
        Exit Function
    ElseIf cmbVender.MatchFound = False Then
        MsgBox "Please Select Maker.", vbInformation
        Exit Function
    ElseIf cmbItems.MatchFound = False Then
        MsgBox "Please Select Item.", vbInformation
        Exit Function
    ElseIf Val(txtQty) <= 0 Then
        MsgBox "Invalid Qty.", vbInformation
        Exit Function
    End If
    
    Dim iProcessID As Integer
    iProcessID = cmbReturnProcess.List(cmbReturnProcess.ListIndex, 1)
    
    If iProcessID = lCheckerProcessNo And cmbChecker.MatchFound = False Then
        MsgBox "Please Check Checker.", vbInformation
        Exit Function
    End If
    
    Dim lLotNo As String
    Dim SaveIt As Boolean, SaveItReturnDTs As Boolean
    
    Dim VendID As New Collection
    Dim VIssID As New Collection
     
'    Dim strVendnItem() As String
'    Dim iVendnItemCount As Integer
    
    Dim myPOSavedCol As New Collection, myPOSavedColWithOrder As New Collection
    
    Dim cmd As New ADODB.Command
    Dim TempIndex As Integer
    TempIndex = 1
    
    Dim lRefID As Long
    
    Dim IssLot As Boolean
    Dim iAuth As Integer 'Authorize
    
    IssLot = RcvingLot 'RcvLot(cmbProcess.ID)
    
'    With cmd
'        .CommandType = adCmdStoredProc
'        .CommandText = "SP_ReqAuthForIss"
'        Set .ActiveConnection = con
'        .Parameters("@UserID").value = UserID
'        .Parameters("@ProcessID").value = cmbProcess.ID
'        .Execute
'        UserReqAuth = Val(.Parameters("@ReqAuth").value & "")
'    End With
    UserReqAuth = 0
    
    Dim strMasterPONo As String
    If bReIssuance = False Then
        strMasterPONo = "NULL" '"'" & GetNextMasterPONo & "'"
    Else
        strMasterPONo = "NULL"
    End If
    
    Dim lRecAffect As Long
    Dim lVID_EntryID As Long
    Dim iAuthorized As Integer
    If iFactoryMakerID = cmbVender.ID Then
        iAuthorized = 1
    Else
        iAuthorized = 0
    End If
    Call StartTrans(con)
    
        SaveIt = True
        Dim myPOObjTemp As clsPOEntry
        For j = 1 To myPOSavedCol.count
            Set myPOObjTemp = myPOSavedCol(j)
            If myPOObjTemp.iIssueProcess = iProcessID And myPOObjTemp.iReturnProcess = iProcessID And myPOObjTemp.iVendID = cmbVender.ID And myPOObjTemp.strItemCode = cmbItems.Text Then
                SaveIt = False
                lRefID = myPOObjTemp.lRefID
                Set myPOObjTemp = Nothing
                Exit For
            End If
            Set myPOObjTemp = Nothing
        Next
            
    If SaveIt Then
        Dim lMakerIssNo As Long, strMakerIssNo As String
        lMakerIssNo = Val(con.Execute("SELECT MAX(CAST(RIGHT(RecieptID,Len(RecieptID)-12) As Int)) From VendIssued Where Convert(varchar,DT,6)='" & Format(DTPicker2, "dd MMM yy") & "'").Fields(0).Value & "")
        lMakerIssNo = lMakerIssNo + 1
        IssID = "ISU-" & Format(DTPicker2, "ddMMyy") & GotIssID
        strMakerIssNo = "M-ISU-" & Format(DTPicker2, "ddMMyy") & lMakerIssNo
        
        con.Execute "INSERT INTO VendIssued(VendID,DT,RecieptID,UserID,ProcessID,ItemID," & _
        "UserName,MachineName,SpecialInstructions,ExcessQtyPercentage," & _
        "MaximumRcvingsAgainstPO,MasterPONo,SampleProvided,DrawingProvided,ForgingProvided,SteelProvided,Authorized) VALUES(" & cmbVender.ID & ",'" & _
         DTPicker2.Value & "','" & strMakerIssNo & "'," & UserID & "," & _
         iProcessID & ",'" & cmbItems.Text & "','" & CurrentUserName & "','" & _
         strComputerName & "','" & txtSpecialInstructions & "'," & Val(txtExcess) & "," & _
         Val(txtRcvings) & "," & strMasterPONo & "," & chkSampleProvided.Value & "," & chkDrawingProvided.Value & "," & _
         chkForgingProvided.Value & "," & chkSteelProvided.Value & "," & iAuthorized & ")"
         
        'lRefID = Val(con.Execute("SELECT Max(EntryID) FROM VendIssued").Fields(0).Value & "")
        lRefID = GetSingleLongValue("MAX(EntryID)", "VendIssued", " WHERE MachineName='" & strComputerName & "'")
        If cmbVender.ID = iFactoryMakerID And (strCompany = "Weldon Instruments" Or strCompany = "QEL") Then
            con.Execute "UPDATE VendIssued SET Authorized=1 WHERE EntryID=" & lRefID
        End If
        
        'Add to Collection After Assigning to Object
        Dim myPOObj As New clsPOEntry
        myPOObj.iIssueProcess = iProcessID
        myPOObj.iReturnProcess = iProcessID
        myPOObj.iVendID = cmbVender.ID
        myPOObj.strItemCode = cmbItems.Text
        myPOObj.lRefID = lRefID
        myPOSavedCol.add myPOObj
        Set myPOObj = Nothing
                
        '1. Insert Into VendIssuedItems
            
        'If strCompany = "QEL" Or strCompany = "Dr-Frgz" Then
'            If bLotNoWiseEntry = False Then
'                lLotNo = getNextLotNo
'            Else
'                lLotNo = txtLotNo
'            End If
'            If chkUseOriginalLotNo.Value = vbChecked Then
'                lLotNo = LV.ListItems(1).SubItems(4)
'            Else
'                lLotNo = getNextLotNo
'            End If
'        Else
'            lLotNo = GetSingleLongValue("MAX(LotNo)", "VendIssdDetail") + 1
'        End If
        lLotNo = getNextSubLotNo(LV.ListItems(1).SubItems(4))
        'Also Get the Batch & Mil No.
        Dim strBatchNo As String, strMilNo As String
        strBatchNo = GetSingleStringValue("Batch_No", "Lots_List", " WHERE LotNo='" & LV.ListItems(1).SubItems(4) & "'")
        strMilNo = GetSingleStringValue("Mill_Certificate_No", "Lots_List", " WHERE LotNo='" & LV.ListItems(1).SubItems(4) & "'")
        
        con.Execute "INSERT INTO Lots_List(LotNo,ItemID,Lot_Type,Reference_LotNo,Batch_No,Mill_Certificate_No) VALUES('" & lLotNo & "','" & cmbItems.Text & "',3,'','" & strBatchNo & "','" & strMilNo & "')"
        
        SaveItReturnDTs = True
              
        Dim strOrderNo As String
        If cmbIssOrderNo.MatchFound = False Then
            strOrderNo = ""
        Else
            strOrderNo = cmbIssOrderNo.List(cmbIssOrderNo.ListIndex, 1)
        End If
    
        If SaveItReturnDTs = False Then
            con.Execute "UPDATE VendIssdDetail SET IssQty=IssQty+" & Val(txtQty) & " WHERE EntryID=" & lVID_EntryID
        Else
            
            GotIssID = Val(con.Execute("SELECT MAX(CAST(RIGHT(RecieptID,Len(RecieptID)-10) AS INT)) FROM VMakerIssItems WHERE Convert(varchar,DT,6)='" & Format(DTPicker2, "dd MMM yy") & "'").Fields(0).Value & "")
            GotIssID = GotIssID + 1
            IssID = "ISU-" & Format(DTPicker2, "ddMMyy") & GotIssID
     
            con.Execute "INSERT INTO VendIssdDetail(RefID,RecieptID,ItemCode" & _
             ",Rate,IssQty,ReqAuth,OrderNo,RcvProcessID,ReturnDT,Priority,Rcvd_RefID,LotNo,ReWorkLot,Repair_RefID) VALUES(" & lRefID & ",'" & _
             IssID & "','" & cmbItems.Text & "'," & Val(txtVendRate) & "," & Val(txtQty) & "," & UserReqAuth & ",'" & _
             strOrderNo & "'," & iProcessID & ",'" & DTReturn & "',1,0,'" & lLotNo & "',0,0)"
                    
            lVID_EntryID = GetSingleLongValue("MAX(EntryID)", "VendIssdDetail")
                    
            'Add to Collection After Assigning to Object
            Dim myPOObj1 As New clsPOEntryWithOrder
            myPOObj1.iIssueProcess = iProcessID
            myPOObj1.iReturnProcess = iProcessID
            myPOObj1.iVendID = cmbVender.ID
            myPOObj1.strItemCode = cmbItems.Text
            myPOObj1.strOrderNo = strOrderNo
            myPOObj1.lRefID = lVID_EntryID
            myPOSavedColWithOrder.add myPOObj1
            Set myPOObj1 = Nothing
        End If
                
        con.Execute "INSERT INTO VendIssdDetail_ReturnDTs(VIS_RefID,IssQty,RcvQty,ReturnDT) VALUES(" & _
         lVID_EntryID & "," & Val(txtQty) & ",0,'" & DTReturn.Value & "')"
                
        If iProcessID = lCheckerProcessNo Then
            con.Execute "INSERT INTO VendIssdDetail_MoreDetails(VID_RefID,EmpID) VALUES(" & lVID_EntryID & ",'" & cmbChecker.List(cmbChecker.ListIndex, 1) & "')"
        End If
                
        'Authentication Info...
        If UserReqAuth = 0 Then
            Dim DetailRefID As Long
            DetailRefID = Val(con.Execute("SELECT MAX(EntryID) FROM VendIssdDetail").Fields(0).Value & "")
            con.Execute "INSERT INTO VendIssAuthDetail(UserID,IssID,DT) Values(" & _
             UserID & "," & DetailRefID & ",'" & ServerDate & " " & ServerTime & "')"
        End If
    End If
    
    Dim lShelf_RefID As Long, lOpeningEntryID As Long, lShelfQty As Long
    
    For i = 1 To LV.ListItems.count
    
        With LV.ListItems(i)
        
            lOpeningEntryID = Val(.ListSubItems(1).Tag)
'            If bLotNoWiseEntry Then
'                lOpeningEntryID = GetSingleLongValue("SOO_RefID", "LotTransferDetails", " WHERE EntryID=" & Val(txtLotNo.Tag))
'            Else
''                If strCompany = "QEL" Then
''                    lOpeningEntryID = GetSingleLongValue("EntryID", "StockOrderOpening", " WHERE ItemID='" & cmbItems.Text & "' AND ProcessID=" & cmbProcess.ID)
''                Else
''                    lOpeningEntryID = GetSingleLongValue("EntryID", "StockOrderOpening", " WHERE ItemID='" & cmbItems.Text & "' AND ProcessID=" & cmbProcess.ID & " AND Shelf_RefID=" & lShelf_RefID)
''                End If
'            End If
            
            lShelfQty = Val(.SubItems(3))
            
            con.Execute "UPDATE VendRcvdDetail SET IssQty=IssQty+" & lShelfQty & " WHERE Opening_RefID=" & lOpeningEntryID
'            If bLotNoWiseEntry Then
'                con.Execute "UPDATE VendRcvdDetail SET IssQty=IssQty+" & lShelfQty & " FROM VendRcvdDetail INNER JOIN LotTransferDetails ON VendRcvdDetail.EntryID=LotTransferDetails.VRD_To_RefID WHERE LotTransferDetails.EntryID=" & Val(txtLotNo.Tag)
'            Else
'                con.Execute "UPDATE VendRcvdDetail SET IssQty=IssQty+" & lShelfQty & " WHERE Opening_RefID=" & lOpeningEntryID
'            End If
            If strCompany = "QEL" Then
                'con.Execute "UPDATE StockOrderOpening SET WHERE EntryID=" & lOpeningEntryID  Do Nothing...
            Else
                con.Execute "UPDATE StockOrderOpening SET Remarks='" & .SubItems(2) & "' WHERE EntryID=" & lOpeningEntryID
            End If
            If bLotNoWiseEntry Then
                con.Execute "INSERT INTO StockOrderOpening_Issuance(SOO_RefID,QtyIssd,VID_RefID,LTD_RefID) VALUES(" & lOpeningEntryID & _
                 "," & lShelfQty & "," & lVID_EntryID & "," & Val(txtLotNo.Tag) & ")"
            Else
                con.Execute "INSERT INTO StockOrderOpening_Issuance(SOO_RefID,QtyIssd,VID_RefID) VALUES(" & lOpeningEntryID & _
                 "," & lShelfQty & "," & lVID_EntryID & ")"
            End If
        End With
    Next
    
    '-----------------------------------------------------------------
    
    
    con.CommitTrans
    
'    If bLotNoWiseEntry = False Then
'        MsgBox "Lot #:" & lLotNo & " Generated for the Entry.", vbInformation
'    End If
    
    Saved = True
    Exit Function
err:
    MsgBox err.Description
    
End Function

Private Sub txtRate1_KeyDown(KeyCode As Integer, Shift As Integer)
'    MsgBox KeyCode
End Sub

Private Sub txtRate1_KeyPress(KeyAscii As Integer)

    On Error GoTo err
    Dim iQty  As Long
    If KeyAscii = 13 Then
        'LVMaker.SelectedItem.SubItems(3) = Val(txtRate)
        If Val(txtRate1) <= 0 Then
            MsgBox "Invalid Data.", vbInformation
            Exit Sub
        End If
        Dim ITM As ListItem
        If iColNo = 9 Then
            iQty = LV.ListItems(LVIssued.SelectedItem.Tag).ListSubItems(4).Tag - Val(txtRate1)
            If iQty < 0 Then
                MsgBox "Invalid Qty!", vbInformation
                Exit Sub
            End If
            LV.ListItems(LVIssued.SelectedItem.Tag).SubItems(4) = iQty
            LVIssued.SelectedItem.SubItems(iColNo - 1) = Val(txtRate1)
        Else
            LVIssued.SelectedItem.SubItems(iColNo - 1) = Val(txtRate1)
        End If
        
        txtRate1.Visible = False
        Call LV_ItemClick(LV.SelectedItem)
        'End If
'        If Val(txtRate1) = 0 Then
'            Call LVIssued_KeyDown(46, 0)
'        End If
        
    ElseIf KeyAscii = 27 Then
        txtRate1.Visible = False
    End If
    Call OnlyNums(KeyAscii)
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub txtRate1_LostFocus()
    txtRate1.Visible = False
    'Me.KeyPreview = True
    'cmdPost.Default = True
End Sub

Private Sub txtRate_LostFocus()
    txtRate.Visible = False
End Sub

Private Sub PostForLots()

    On Error GoTo err
    Dim ITM As ListItem

    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            IssID = LV.ListItems(i).Tag
            With ITM
                Set ITM = LVIssued.ListItems.add(, , LVIssued.ListItems.count + 1)
                ITM.Tag = LV.ListItems(i).key
                ITM.ListSubItems.add(, cmbVender.ID & "'", cmbVender.Text).Tag = cmbVender.ID
                ITM.ListSubItems.add(, , IssID & "").Tag = DTPicker2
                ITM.ListSubItems.add , , LV.ListItems(i).SubItems(2)
                ITM.ListSubItems.add , , LV.ListItems(i).SubItems(3)
                ITM.ListSubItems.add , , Val(cmbVender.Tag)
                ITM.ListSubItems.add , , LV.ListItems(i).SubItems(4)
                'itm.ListSubItems.Add , , Val(LV.ListItems(i).SubItems(5))
                If cmbProcess.Tag = "Lot No" Then
                    ITM.ListSubItems(6).Tag = LV.ListItems(i).ListSubItems(1).Tag
                End If
            End With
            LV.ListItems(i).SubItems(4) = 0
            
        End If
    Next
    
    
    'Remove Posted Entries From the List.
    i = 1
    Do While i <= LV.ListItems.count
        If LV.ListItems(i).Checked Then
            LV.ListItems.Remove i
            On Error Resume Next
            cmbDesc.RemoveItem i - 1
            If err.Number = -2147024809 Then Resume Next
                
            i = i - 1
        End If
        i = i + 1
    Loop
    
    'if cmbitems.
    'Call LV_ItemClick(LV.ListItems(i))
    
    Call FillVendList
    
    txtQty = ""
    Exit Sub
err:
    MsgBox err.Description

End Sub

Public Sub ShowMe(m_bReIssuance As Boolean, Optional p_bIssueAgainstOpening As Boolean = False)

    bReIssuance = m_bReIssuance
    bIssueAgainstOpening = p_bIssueAgainstOpening
    If bReIssuance = False Then
        'cmbProcess.AddItem "<All Processes>", "0"
        'txtLotNo.Enabled = False
    Else
        Label1(2).Visible = False
        txtExcess.Text = "0"
        txtExcess.Visible = False
        
        'txtLotNo.Enabled = False
       
    End If
    
    cmbProcess.AddVals con, "Description", "Processes", "ProcessID", " Order By SNo"
    
    Call AddToCombo(cmbItems, "DISTINCT ItemID", "VStockOrderOpening", " WHERE Qty>IssdQty")
    
    If bReIssuance = False Then cmbProcess.ID = "0"
    
    Me.Show
    
End Sub

'Private Function GetNextMasterPONo() As String
'
'    Dim Prefix As String
'    'ORD-DDMM-001
'    Dim myDT As Date
'    myDT = Date
'    Dim lOrderNo As Long
'    lOrderNo = GetSingleLongValue("MAX(CAST(LEFT(MasterPONo,CHARINDEX('/',MasterPONo)-1) AS INT))", "VendIssued", " WHERE RIGHT(MasterPONo,4)='" & Format(myDT, "MMyy") & "' AND MasterPONo IS NOT NULL")
'    lOrderNo = lOrderNo + 1
'    GetNextMasterPONo = lOrderNo & Format(myDT, "/MMyy")
'
'End Function

