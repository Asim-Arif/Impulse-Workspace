VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmRMIssuanceLocWise 
   ClientHeight    =   10095
   ClientLeft      =   1110
   ClientTop       =   60
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
   MDIChild        =   -1  'True
   ScaleHeight     =   10095
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicProduct 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2865
      Left            =   8385
      ScaleHeight     =   2835
      ScaleWidth      =   2760
      TabIndex        =   42
      Top             =   3765
      Visible         =   0   'False
      Width           =   2790
      Begin VB.Image ImgProduct 
         Height          =   2835
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2760
      End
   End
   Begin VB.CheckBox chkShowPicture 
      Caption         =   "Show Picture"
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
      Height          =   255
      Left            =   420
      TabIndex        =   41
      Top             =   9690
      Width           =   1560
   End
   Begin ComboList.Usercontrol1 cmbMaker 
      Height          =   285
      Left            =   405
      TabIndex        =   0
      Top             =   1410
      Width           =   4890
      _ExtentX        =   8625
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
   Begin ComboList.Usercontrol1 cmbIssuance 
      Height          =   285
      Left            =   8955
      TabIndex        =   3
      Top             =   1410
      Width           =   2190
      _ExtentX        =   3863
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
   Begin ComboList.Usercontrol1 cmbAuthBy 
      Height          =   285
      Left            =   5865
      TabIndex        =   2
      Top             =   1905
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
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   405
      TabIndex        =   1
      Top             =   1905
      Width           =   5460
      _ExtentX        =   9631
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
   Begin ComboList.Usercontrol1 cmbMaterial 
      Height          =   285
      Left            =   405
      TabIndex        =   4
      Top             =   2940
      Width           =   9930
      _ExtentX        =   17515
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
   Begin VB.TextBox txtMinStock 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   3465
      Width           =   960
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   405
      MultiLine       =   -1  'True
      TabIndex        =   38
      Top             =   8760
      Width           =   10755
   End
   Begin VB.Frame Frame1 
      Caption         =   "Last 2 Month's Issuance :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2190
      Left            =   405
      TabIndex        =   36
      Top             =   6330
      Width           =   10755
      Begin MSComctlLib.ListView LVHistory 
         Height          =   1860
         Left            =   60
         TabIndex        =   37
         Top             =   255
         Width           =   10590
         _ExtentX        =   18680
         _ExtentY        =   3281
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Material ID"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Material Name"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Qty"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   8085
      TabIndex        =   34
      Top             =   1425
      Width           =   900
   End
   Begin MSComCtl2.DTPicker DTFrom 
      Height          =   300
      Left            =   5310
      TabIndex        =   32
      Top             =   1410
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   159186947
      CurrentDate     =   40935
   End
   Begin VB.TextBox txtIssSheets 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4830
      TabIndex        =   31
      Top             =   9630
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox txtInHandSheets 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3645
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   9465
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CheckBox chkRMPOs 
      BackColor       =   &H00AFBDCB&
      Caption         =   "RM PO Nos."
      Height          =   195
      Left            =   9870
      TabIndex        =   29
      Top             =   1215
      Width           =   1185
   End
   Begin MSComctlLib.ListView LVRMIss 
      Height          =   2565
      Left            =   420
      TabIndex        =   26
      Top             =   3765
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   4524
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
         Text            =   "Material ID"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Material Name"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Store"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Unit"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Rate"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Location"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Batch #"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "Lot #"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.TextBox txtProcess 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D3E7EF&
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   2430
      Width           =   3420
   End
   Begin VB.TextBox txtIssdQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D3E7EF&
      Height          =   285
      Left            =   10380
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   2430
      Width           =   765
   End
   Begin VB.TextBox txtItemName 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D3E7EF&
      Height          =   285
      Left            =   1875
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   2430
      Width           =   5085
   End
   Begin VB.TextBox txtItemCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D3E7EF&
      Height          =   285
      Left            =   405
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   2430
      Width           =   1470
   End
   Begin VB.TextBox txtPrice 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10350
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2940
      Width           =   795
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5520
      TabIndex        =   20
      Top             =   5265
      Visible         =   0   'False
      Width           =   915
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
      Height          =   270
      Left            =   420
      TabIndex        =   12
      Top             =   9390
      Width           =   2070
   End
   Begin VB.TextBox txtInHand 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7785
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3465
      Width           =   960
   End
   Begin VB.TextBox txtRMQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8760
      TabIndex        =   7
      Top             =   3465
      Width           =   825
   End
   Begin VB.TextBox txtUnit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D3E7EF&
      Height          =   285
      Left            =   7155
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   3465
      Width           =   615
   End
   Begin VB.Frame Frame6 
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
      Left            =   5865
      TabIndex        =   16
      Top             =   9330
      Width           =   5280
      Begin MSForms.CommandButton cmdSave 
         Height          =   360
         Left            =   75
         TabIndex        =   9
         Top             =   195
         Width           =   1665
         Caption         =   "Save & New   "
         PicturePosition =   327683
         Size            =   "2937;635"
         Accelerator     =   110
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdCancel 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   3510
         TabIndex        =   11
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
      Begin MSForms.CommandButton cmdSaveClose 
         Height          =   360
         Left            =   1800
         TabIndex        =   10
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
      Left            =   720
      TabIndex        =   13
      Top             =   0
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Material"
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
         Left            =   3900
         TabIndex        =   14
         Top             =   135
         Width           =   2220
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Material"
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
         Left            =   3915
         TabIndex        =   15
         Top             =   150
         Width           =   2220
      End
   End
   Begin MSComCtl2.DTPicker DTTo 
      Height          =   300
      Left            =   6705
      TabIndex        =   33
      Top             =   1410
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   159186947
      CurrentDate     =   40935
   End
   Begin MSComCtl2.DTPicker DT 
      Height          =   300
      Left            =   405
      TabIndex        =   44
      Top             =   840
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   159186947
      CurrentDate     =   40935
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   4
      Left            =   405
      TabIndex        =   43
      Top             =   600
      Width           =   1410
      BackColor       =   11517387
      Caption         =   "Issue Date"
      Size            =   "2487;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label3 
      Height          =   240
      Left            =   405
      TabIndex        =   39
      Top             =   8535
      Width           =   10755
      BackColor       =   11517387
      Caption         =   " Remarks"
      Size            =   "18971;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   3
      Left            =   405
      TabIndex        =   35
      Top             =   1680
      Width           =   10740
      BackColor       =   11517387
      Caption         =   " Issued To                                                                                                        Authorized By"
      Size            =   "18944;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cmbLocation 
      Height          =   285
      Left            =   405
      TabIndex        =   6
      Top             =   3465
      Width           =   6735
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "11880;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label4 
      Height          =   240
      Left            =   405
      TabIndex        =   28
      Top             =   3240
      Width           =   10155
      BackColor       =   11517387
      Caption         =   $"frmRMIssuanceLocWise.frx":0000
      Size            =   "17912;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   2
      Left            =   405
      TabIndex        =   22
      Top             =   2190
      Width           =   10740
      BackColor       =   11517387
      Caption         =   $"frmRMIssuanceLocWise.frx":00C1
      Size            =   "18944;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   1
      Left            =   405
      TabIndex        =   21
      Top             =   1185
      Width           =   10740
      BackColor       =   11517387
      Caption         =   $"frmRMIssuanceLocWise.frx":018B
      Size            =   "18944;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   0
      Left            =   405
      TabIndex        =   18
      Top             =   2715
      Width           =   10740
      BackColor       =   11517387
      Caption         =   $"frmRMIssuanceLocWise.frx":023B
      Size            =   "18944;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdAdd 
      Height          =   525
      Left            =   10590
      TabIndex        =   8
      Top             =   3210
      Width           =   555
      Caption         =   "Add"
      PicturePosition =   2
      Size            =   "979;926"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmRMIssuanceLocWise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strIssNo As String
Dim strDelList As String
Dim strUserCondition As String

Private Sub chkRMPOs_Click()
    Call AddPOs
End Sub

Private Sub cmbEmp_matched()
    
    If strCompany = "Dr-Frgz" Then
    
        cmbMaterial.ClearVals
        cmbMaterial.AddVals con, "RMName + ' {' + RMID1 + '} ' + RMSize + ' ' + SizeUnit ", "RM", "RMID", " WHERE InActive=0 AND RMID IN(SELECT RMID FROM Employees_RM_Assigned WHERE EmpID='" & cmbEmp.ID & "') " & strUserCondition & " ORDER BY RMID"
    
    End If
    
    Call ShowHistory
    
End Sub

Private Sub cmbIssuance_matched()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM VVendIssued WHERE EntryID=" & cmbIssuance.ID, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            txtItemCode = !ITemID & ""
            'Filter RM as per assigned only to this item.
            cmbMaterial.ClearVals
            cmbMaterial.AddVals con, "RMName + ' {' + RMID1 + '} ' + CAST(RMSize AS VARCHAR(50)) + ' ' + SizeUnit ", "RM", "RMID", " WHERE InActive=0 AND RMID IN(SELECT RMID FROM ItemsRMComp WHERE ItemID='" & !ITemID & "' AND Functional_Status=0) ORDER BY RMID"
            txtItemName = !ItemName & ""
            txtProcess = !Description & ""
            txtIssdQty = Val(!TotalIssQty & "")
        Else
            txtItemCode = "-"
            txtItemName = "-"
            txtProcess = "-"
            txtIssdQty = "-"
        End If
        .Close
    End With
        
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmbmaker_matched()

    Call AddPOs
    
    If strCompany = "Dr-Frgz" Then
        cmbMaterial.ClearVals
    End If
    If cmbMaker.ID = iFactoryMakerID Then
        cmbEmp.Enabled = True
    Else
        cmbEmp.Enabled = False
        If strCompany = "Dr-Frgz" Then
            cmbMaterial.AddVals con, "RMName + ' {' + RMID1 + '} ' + RMSize + ' ' + SizeUnit ", "RM", "RMID", " WHERE InActive=0 AND RMID IN(SELECT RMID FROM Makers_RM_Assigned WHERE MakerID=" & cmbMaker.ID & ") " & strUserCondition & " ORDER BY RMID"
        End If
        Call ShowHistory
    End If
    
    
End Sub

Private Sub AddPOs(Optional bDateRange As Boolean = False)

    cmbIssuance.ClearVals
    If strCompany = "Tecno" Or (strCompany = "Dr-Frgz" And strSubCompany = "Electro") Then
        cmbIssuance.AddVals con, "VendIssued.RecieptID + ' (' + CAST(IPODL_RefID AS VARCHAR(50)) + ')'", "VendIssued INNER JOIN VendIssdDetail ON VendIssued.EntryID=VendIssdDetail.RefID", "VendIssued.EntryID", " WHERE VendID=" & cmbMaker.ID & " ORDER BY VendIssued.EntryID DESC"
    Else
        If chkRMPOs.Value = vbChecked Then
            Dim strAccNo As String
            strAccNo = GetSingleStringValue("AccNo", "VVendersANDMakers", " WHERE VendID=" & cmbMaker.ID)
            cmbIssuance.AddVals con, "OrderNo", "VendOrders", "OrderNo", " WHERE VendID='" & strAccNo & "' ORDER BY OrderNo DESC"
        Else
            If bDateRange Then
                cmbIssuance.AddVals con_ServerSide, "RecieptID", "VendIssued", "EntryID", " WHERE DT BETWEEN '" & DTFrom.Value & "' AND '" & DTTo.Value & "' AND VendID=" & cmbMaker.ID & " ORDER BY EntryID DESC"
            Else
                cmbIssuance.AddVals con_ServerSide, "TOP 30000 RecieptID", "VendIssued", "EntryID", " WHERE VendID=" & cmbMaker.ID & " ORDER BY EntryID DESC"
            End If
        End If
    End If
    
End Sub

Private Sub cmbMaterial_matched()
        
    Dim iRMGroup As Integer, iForgingCount As Integer
    iRMGroup = GetSingleLongValue("GroupID", "RM", " WHERE RMID=" & cmbMaterial.ID)
    iForgingCount = GetSingleLongValue("COUNT(Group_ID)", "RMGroupIDsForForging", " WHERE Group_ID=" & iRMGroup)
    
    Dim lTwoUnitsCount As Long, lBatchWiseIssuanceCount As Long
    lTwoUnitsCount = GetSingleLongValue("COUNT(*)", "RMGroupsForTwoUnits", " WHERE RMG_RefID=" & iRMGroup)
    
    If lTwoUnitsCount > 0 Then
        txtIssSheets.Enabled = True
    Else
        txtIssSheets.Enabled = False
    End If
    
    lBatchWiseIssuanceCount = GetSingleLongValue("COUNT(*)", "RMGroupsForBatchwiseIssuance", " WHERE RMG_RefID=" & iRMGroup)
    
    Dim rs As New ADODB.Recordset
'    If iForgingCount > 0 Then
'
'        With rs
'            .Open "SELECT Unit,QtyInStock,Rate,AnnealingStock,MachiningStock,MakerRate FROM RM WHERE RMID='" & cmbMaterial.ID & "'", con, adOpenForwardOnly, adLockReadOnly
'            txtPrice = Val(!MakerRate & "")
'            txtUnit = !Unit & ""
'            txtPrice.Tag = "1"  'Forging
'            txtInHand.Tag = Val(!MachiningStock & "")
'            txtInHand = Val(!MachiningStock & "")
'            txtUnit.Tag = Val(!MachiningStock & "")
'            .Close
'        End With
'    Else
        
        txtPrice.Tag = "0"
        
        With rs
            If lBatchWiseIssuanceCount > 0 Then
                .Open "SELECT * FROM VMaterialLocationANDPOWiseBalances WHERE RMID=" & cmbMaterial.ID, con, adOpenForwardOnly, adLockReadOnly
            Else
                .Open "SELECT * FROM VMaterialLocationWiseBalances WHERE RMID=" & cmbMaterial.ID & " AND ShelfQty>0", con, adOpenForwardOnly, adLockReadOnly
            End If
            
            cmbLocation.Clear
            Do Until .EOF
                If lBatchWiseIssuanceCount > 0 Then
                    cmbLocation.AddItem "Batch:" & !BatchNo & " Lot:" & !LotNo & " " & !StoreName & " [" & !RackNo & "] [" & !ShelfNo & "] (" & Val(!ShelfQty & "") & ")"
                Else
                    cmbLocation.AddItem !StoreName & " [" & !RackNo & "] [" & !ShelfNo & "] (" & Val(!ShelfQty & "") & ")"
                End If
                
                cmbLocation.List(.AbsolutePosition - 1, 1) = !Shelf_RefID
                cmbLocation.List(.AbsolutePosition - 1, 2) = Val(!ShelfQty & "")
                cmbLocation.List(.AbsolutePosition - 1, 3) = Val(!SheetsQty & "")
                If lBatchWiseIssuanceCount > 0 Then
                    cmbLocation.List(.AbsolutePosition - 1, 4) = !BatchNo & "" '!PORefNo & ""
                    cmbLocation.List(.AbsolutePosition - 1, 5) = !LotNo & ""
                Else
                    cmbLocation.List(.AbsolutePosition - 1, 4) = ""
                    cmbLocation.List(.AbsolutePosition - 1, 5) = ""
                End If
                .MoveNEXT
            Loop
            .Close
            
            .Open "SELECT Unit,QtyInStock,MakerRate,Rate,MinLevel FROM RM WHERE RMID='" & cmbMaterial.ID & "'", con, adOpenForwardOnly, adLockReadOnly
            txtMinStock = Val(!MinLevel & "")
            txtPrice = Val(!Rate & "")
            txtUnit = !Unit & ""
            '        txtUnit.Tag = Val(!QtyInStock & "")
            txtInHand.Tag = Val(!QtyInStock & "")   'Total Qty from All Stores....
        
            .Close
        
        End With
        Set rs = Nothing
        
'    End If
    
    Exit Sub
    
err:
    MsgBox err.Description
End Sub

Private Sub cmbPrinter_Change()
    txtAttn = cmbPrinter.Text
End Sub

Private Sub cmbStore_matched()

    Dim dStoreQty As Double
    dStoreQty = GetSingleDoubleValue("NetQty", "VRMStatusStorewise", " WHERE RMID=" & cmbMaterial.ID & " AND StoreName='" & cmbStore.Text & "'")
    txtUnit.Tag = dStoreQty
    txtInHand = dStoreQty
    
End Sub

Private Sub cmdAdd_Click()
    
    On Error GoTo err
    
    If cmbMaterial.MatchFound = False Then Exit Sub
    If cmbLocation.MatchFound = False Then
        MsgBox "Please Select Location.", vbInformation
        Exit Sub
    End If
    
    If Val(txtRMQty) <= 0 Then Exit Sub
    
    If txtIssSheets.Enabled And Val(txtIssSheets) <= 0 Then
        MsgBox "Invalid Sheets.", vbInformation
        Exit Sub
    End If
    
    Dim lSF_RefID As Long
    lSF_RefID = 0 'GetSingleLongValue("SF_RefID", "RM", " WHERE RMID=" & cmbMaterial.ID)
    
    If lSF_RefID = 0 Then   'If Regual Material then Apply check, No Check for SemiFinish materials...
        If Val(txtRMQty) > Val(txtInHand) Then
            MsgBox "Can't Issue This Quantity." & vbNewLine & "Only " & txtUnit.Tag & " " & txtUnit & " Is Available In The Store.", vbInformation
            Exit Sub
        End If
        If Val(txtIssSheets) > Val(txtInHandSheets) Then
            MsgBox "Can't Issue This Quantity." & vbNewLine & "Only " & txtInHandSheets.Text & " Sheets are Available In The Store.", vbInformation
            Exit Sub
        End If
    End If
    
    '...............................
    If strCompany <> "Dr-Frgz" Then GoTo AllowEntry
    Dim strSQLString As String, strSQLString1 As String
    Dim lQty As Long, lFrequencey As Long
    If cmbEmp.Enabled Then  'Employee
        strSQLString = "SELECT MaxQty,MaxFrequency FROM Employees_RM_Assigned WHERE EmpID='" & cmbEmp.ID & "' AND RMID=" & cmbMaterial.ID
    Else                    'Maker
        strSQLString = "SELECT MaxQty,MaxFrequency FROM Makers_RM_Assigned WHERE MakerID=" & cmbMaker.ID & " AND RMID=" & cmbMaterial.ID
    End If
    Dim rs As New ADODB.Recordset
    With rs
        .Open strSQLString, con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            lQty = -1
            lFrequency = -1
        Else
            lQty = Val(!MaxQty & "")
            lFrequency = Val(!MaxFrequency & "")
        End If
        .Close
        
        If lQty = -1 Then
            MsgBox "Invalid Entry.", vbInformation
            Exit Sub
        End If
        
        If lFrequency = 0 Or lQty = 0 Then GoTo AllowEntry
        
        Dim DTFrom As Date
        DTFrom = DateAdd("d", -lFrequency + 1, Date)
         
        Dim lFrequencyCount As Long, lIssdQty As Long
        If cmbEmp.Enabled Then
            .Open "SELECT COUNT(*) AS Frequency,SUM(Qty) AS Qty FROM VRawMaterialIssuance WHERE IssdToEmpID='" & cmbEmp.ID & "' AND DT BETWEEN '" & DTFrom & "' AND '" & Date & "'", con, adOpenForwardOnly, adLockReadOnly
        Else
            .Open "SELECT COUNT(*) AS Frequency,SUM(Qty) AS Qty FROM VRawMaterialIssuance WHERE VendID=" & cmbMaker.ID & " AND DT BETWEEN '" & DTFrom & "' AND '" & Date & "'", con, adOpenForwardOnly, adLockReadOnly
        End If
        lFrequencyCount = Val(!Frequency & "")
        lIssdQty = Val(!Qty & "")
        .Close
    End With
    If lFrequencyCount >= lFrequency Then
        MsgBox "Frequency Reacehd.", vbInformation
        Exit Sub
    End If
    If lIssdQty + Val(txtRMQty) > lQty Then
        MsgBox "Qty Reacehd.", vbInformation
        Exit Sub
    End If
    '...............................
AllowEntry:
    
    Dim ITM As ListItem
    Set ITM = LVRMIss.ListItems.add(, cmbMaterial.ID & "'", LVRMIss.ListItems.count + 1)
    ITM.Tag = Val(txtInHand.Tag)
    ITM.ListSubItems.add(, , cmbMaterial.ID).Tag = cmbMaterial.Tag ''Material Type
    ITM.ListSubItems.add(, , cmbMaterial.Text).Tag = ""
    ITM.ListSubItems.add(, , cmbLocation.Text).Tag = txtPrice.Tag
    ITM.ListSubItems.add(, , txtUnit).Tag = Val(cmbLocation.List(cmbLocation.ListIndex, 1))
    ITM.ListSubItems.add , , Val(txtRMQty)
    ITM.ListSubItems.add , , Val(txtPrice)
    ITM.ListSubItems.add , , ""
    ITM.ListSubItems.add(, , cmbLocation.List(cmbLocation.ListIndex, 4)).Tag = cmbLocation.List(cmbLocation.ListIndex, 4)   'Batch
    ITM.ListSubItems.add(, , cmbLocation.List(cmbLocation.ListIndex, 5)).Tag = cmbLocation.List(cmbLocation.ListIndex, 5)   'Lot
    
    cmbLocation.Clear
    cmbMaterial.ID = ""
    txtPrice = ""
    txtUnit = ""
    txtQty = ""
    txtInHand = ""
    txtIssSheets = ""
    txtInHandSheets = ""
    txtInHand.Tag = ""
    cmbMaterial.SetFocus
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    If cmbMaker.MatchFound = False Then
        MsgBox "Please Select Maker.", vbInformation
        Exit Sub
    End If
    Call AddPOs(True)
End Sub

Private Sub cmdSave_Click()
    If Saved Then
        Unload Me
        'LV.ListItems.Clear
        'Call DT_Change
        Me.Show
    End If
End Sub

Private Sub cmdSaveClose_Click()

    If Saved Then Unload Me
    
End Sub

Private Function Saved() As Boolean

    On Error GoTo err
'    If cmbMaker.MatchFound = False Then
'        MsgBox "Please Select Maker.", vbInformation
'        Exit Function
'    End If
    
'    If cmbMaker.ID <> iFactoryMakerID Then
'        If cmbIssuance.MatchFound = False Then
'            MsgBox "Please Select Issuance.", vbInformation
'            Exit Function
'        End If
'        If cmbEmp.Enabled And cmbEmp.MatchFound = False Then
'            MsgBox "Please Select Employee.", vbInformation
'            Exit Function
'        End If
'    End If
    
'    If cmbAuthBy.MatchFound = False Then
'        MsgBox "Please Select Auth. By.", vbInformation
'        Exit Function
'    End If
        
    If LVRMIss.ListItems.count = 0 Then
        MsgBox "No Entries Added.", vbInformation
        Exit Function
    End If
    
    Dim rs As New ADODB.Recordset, strIssNo As String, myDT As Date
    myDT = DT.Value
    With rs
        .Open "SELECT MAX(CAST(RIGHT(IssNo,3) AS INT)) FROM RawMaterialIssuance WHERE DT='" & myDT & "'", con, adOpenForwardOnly, adLockReadOnly
        
        strIssNo = "ISS-" & Format(myDT, "yyMMdd-") & Format(Val(.Fields(0) & "") + 1, "000")
        
        .Close
    End With
    Set rs = Nothing
    
    Dim strOrderNo As String, lVI_RefID As Long
    If chkRMPOs.Value = vbChecked Then
        strOrderNo = cmbIssuance.ID
        lVI_RefID = 0
    Else
        strOrderNo = ""
        lVI_RefID = Val(cmbIssuance.ID)
    End If
    
    Dim lDetailEntryID As Long
    Call StartTrans(con)
    con.Execute "INSERT INTO RawMaterialIssuance(IssNo,DT,UserName,ComputerName,VI_RefID,IssdToEmpID,AuthorizedByEmpID,VO_OrderNo,Remarks) VALUES('" & _
         strIssNo & "','" & myDT & "','" & CurrentUserName & "','" & strComputerName & "'," & _
         lVI_RefID & ",'" & cmbEmp.Text & "','" & cmbAuthBy.Text & "','" & strOrderNo & "','" & txtRemarks & "')"
      
    Dim strMaterialID As String, dRate As Double, dQty As Double, lSheets As Long, strPORefNo As String, strBatchNo As String, strLotNo As String
    For i = 1 To LVRMIss.ListItems.count
        With LVRMIss.ListItems(i)
            strMaterialID = GetSingleStringValue("RMID1", "RM", " WHERE RMID=" & Val(.key))
            strPORefNo = .ListSubItems(8).Tag
            strBatchNo = .ListSubItems(8).Tag
            strLotNo = .ListSubItems(9).Tag
            
            dQty = Val(.SubItems(5))
            lSheets = 0 'Val(.SubItems(8))
            'If Val(.ListSubItems(3).Tag) = 0 Then
                'dRate = GetSingleDoubleValue("Rate", "RM", " WHERE RMID=" & Val(.key))
                dRate = Val(.SubItems(6))
                '-------------------------------------------------------------------------
                con.Execute "INSERT INTO RawMaterialIssuanceDetail(IssNo,RMID1,Qty,InHand,StoreName,Price,Shelf_RefID,Sheets,PORefNo,BatchNo,LotNo) VALUES('" & _
                 strIssNo & "','" & strMaterialID & "'," & dQty & "," & Val(.Tag) - dQty & ",'" & .SubItems(3) & _
                 "'," & dRate & "," & Val(.ListSubItems(4).Tag) & "," & lSheets & ",'" & strPORefNo & "','" & strBatchNo & "','" & strLotNo & "')"
                 
                If Val(.ListSubItems(3).Tag) = 1 Then 'Forging
                    'con.Execute "UPDATE RM SET QtyInStock=QtyInStock-" & dQty & ",MachiningStock=MachiningStock-" & dQty & " WHERE RMID=" & Val(.key)
                    con.Execute "UPDATE RM SET QtyInStock=QtyInStock-" & dQty & ",AnnealingStock=ISNULL(AnnealingStock,0)-" & dQty & " WHERE RMID=" & Val(.key)
                Else
                    con.Execute "UPDATE RM SET QtyInStock=QtyInStock-" & dQty & " WHERE RMID=" & Val(.key)
                End If
                lDetailEntryID = GetSingleLongValue("MAX(EntryID)", "RawMaterialIssuanceDetail")
                Call UpdateMaterialLocation(strMaterialID, Val(.ListSubItems(4).Tag), dQty, lDetailEntryID, , lSheets, strPORefNo, strBatchNo, strLotNo)
                '
            'End If
        End With
    Next
    con.CommitTrans
    Saved = True
    
    If chkPrint.Value = vbChecked Then
        Dim f As New frmPrevRpt
        Dim rpt As CRAXDDRT.Report
        Set rpt = rptApp.OpenReport(RptPath & "\IssuanceNote.rpt")
        f.ShowReport "{IssItemsSimple.IssNo}='" & txtIssNo & "' ", rpt
    End If
    
    Exit Function
err:
    MsgBox err.Description
    
End Function

Private Sub DT_Change()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Max(Cast(Right(IssNo,3) As Int)) From IssItemsSimple Where Convert(Char(8),DT,1)='" & Format(DT, "MM/dd/yy") & "'", con, adOpenForwardOnly, adLockReadOnly
        
        txtIssNo = "ISS-" & Format(DT, "yyMMdd-") & Format(Val(.Fields(0) & "") + 1, "000")
        
        .Close
    End With
    Set rs = Nothing
    
    If cmbMaterial.MatchFound Then
        txtInHand = getItemBalance(cmbMaterialID.ID, Date)
        txtUnit.Tag = Val(txtInHand)
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()

    'DT = Date
    'Call DT_Change
    'Call RestrictIssRcvDates(DT)
'    DT = Date
'    Call DT_Change
'    If CurrentUserName = "Administrator" Then
'        DT.Enabled = True
'    Else
'        DT.Enabled = False
'    End If
    DT.Value = Date
    DTFrom.Value = Date
    DTTo.Value = Date
    
    cmbMaker.ListHeight = 4000
    If strCompany = "Dr-Frgz" Then
        cmbMaker.AddVals con, "VendID1 + ' ' + VenderName", "Makers", "VendID", " WHERE VendID IN(SELECT MakerID FROM Makers_RM_Assigned) OR VendID=" & iFactoryMakerID
    Else
        cmbMaker.AddVals con, "VendID1 + ' ' + VenderName", "Makers", "VendID"
    End If
    
    cmbIssuance.ListHeight = 4000
    cmbMaterial.ListHeight = 4000
        
    
    If CurrentUserName = "Administrator" Then
        strUserCondition = ""
    Else
        strUserCondition = " AND GroupID IN(SELECT RMGroupID FROM Users_RMGroups WHERE UserID=" & UserID & ")"
    End If
    
    If strCompany <> "Dr-Frgz" Then
        cmbMaterial.AddVals con, "RMName + ' {' + RMID1 + '} ' + CAST(RMSize AS VARCHAR(50)) + ' ' + SizeUnit ", "RM", "RMID", " WHERE InActive=0 " & strUserCondition & " ORDER BY RMID"
    End If
    
    cmbEmp.ListHeight = 4000
'    If strCompany = "Dr-Frgz" Then
'        cmbEmp.AddVals con, "'{' + EmpID + '} ' + Name", "Employees", "EmpID", " WHERE EmpID IN(SELECT EmpID FROM Employees_RM_Assigned) ORDER BY EmpID"
'    Else
        cmbEmp.AddVals con, "'{' + EmpID + '} ' + Name", "Employees", "EmpID", " ORDER BY EmpID"
'    End If
    'cmbEmp.AddVals con, "DISTINCT IssdToEmpID", "RawMaterialIssuance", "IssdToEmpID", " ORDER BY RawMaterialIssuance.IssdToEmpID"
    
    cmbAuthBy.ListHeight = 4000
    cmbAuthBy.AddVals con, "'{' + EmpID + '} ' + Name", "Employees", "EmpID", " ORDER BY EmpID"
    
    'AddToCombo cmbPrinter, "PrinterName", "Printers", , , "EntryID"
    strIssNo = ""
    strDelList = ""
        
    If strCompany = "QEL" Or strCompany = "Dr-Frgz" Then
        chkRMPOs.Visible = True
    Else
        chkRMPOs.Visible = False
    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
    
    Dim TitleBottom As Long

    TitleBottom = cmbMaterial.Top + cmbMaterial.Height + 50 '+ fTitle.Height
 
End Sub

Public Sub EditIssuance(p_strIssNo As String)

    On Error GoTo err
    strIssNo = p_strIssNo
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        .Open "SELECT * FROM IssItemsSimple WHERE IssNo='" & strIssNo & "'", con, adOpenForwardOnly, adLockReadOnly
        DT.Value = Format(!DT, "dd-MMM-yyyy")
        txtAttn = !AttnPerson & ""
        txtIssNo = !IssNo
        cmbRestaurant.ID = Val(!Restaurant_RefID & "")
        .Close
        .Open "SELECT * FROM VIssItemsSimpleDetail WHERE IssNo='" & strIssNo & "'", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, !RMID & "'", LV.ListItems.count + 1)
            ITM.Tag = Val(Val(!QtyInStock & ""))
            ITM.ListSubItems.add(, , !RMID).Tag = ""
            ITM.ListSubItems.add(, , !RMName + "{ " & !RMID1 & "}").Tag = ""
            ITM.ListSubItems.add(, , !StoreName & "").Tag = !EntryID
            ITM.SubItems(4) = !Unit & ""
            ITM.SubItems(5) = Val(!Qty & "")
            .MoveNEXT
        Loop
        .Close
    End With
    Me.Show
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub LV_DblClick()
    If LV.ListItems.count = 0 Then Exit Sub
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(6).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(6).Width
        .Text = Val(LV.SelectedItem.SubItems(5))
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)

    If LV.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If Val(LV.SelectedItem.ListSubItems(3).Tag) <> 0 Then
            strDelList = strDelList & Val(LV.SelectedItem.ListSubItems(3).Tag) & ","
        End If
        LV.ListItems.Remove LV.SelectedItem.Index
    End If
    
End Sub

Private Sub LVRMIss_KeyDown(KeyCode As Integer, Shift As Integer)
    If LVRMIss.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Do you want to Remove ?", vbQuestion + vbYesNo) = vbYes Then
            LVRMIss.ListItems.Remove LVRMIss.SelectedItem.Index
        End If
    End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Val(txtEdit) <= 0 Then
            MsgBox "Invalid Qty.", vbInformation
            Exit Sub
        End If
        LV.SelectedItem.SubItems(5) = Val(txtEdit)
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
    
End Sub

Private Sub cmbLocation_Change()
    Call cmbLocation_Click
End Sub

Private Sub cmbLocation_Click()

    If cmbLocation.MatchFound = False Then Exit Sub
    
    txtInHand = Val(cmbLocation.List(cmbLocation.ListIndex, 2))
    txtInHandSheets = Val(cmbLocation.List(cmbLocation.ListIndex, 3))
    
'    Dim lGroupID As Long
'    lGroupID = GetSingleLongValue("GroupID", "RM", " WHERE RMID=" & cmbMaterial.ID)
'
'    Dim lCount As Long
'    lCount = GetSingleLongValue("COUNT(*)", "RMGroupsForTwoUnits", " WHERE RMG_RefID=" & lGroupID)
    
End Sub

Private Sub ShowHistory()

    Dim bMaker As Boolean
    If cmbEmp.Enabled Then
         bMaker = False
    Else
        bMaker = True
    End If
    Dim myDTFrom As Date, myDTTo As Date
    myDTFrom = DateAdd("M", -2, Date)
    myDTTo = Date
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        If bMaker = False Then
            .Open "SELECT RMID1,RMName,SUM(Qty) AS Qty FROM VRawMaterialIssuance WHERE IssdToEmpID='" & cmbEmp.ID & "' AND DT BETWEEN '" & myDTFrom & "' AND '" & myDTTo & "' GROUP BY RMID1,RMName", con, adOpenForwardOnly, adLockReadOnly
        Else
            .Open "SELECT RMID1,RMName,SUM(Qty) AS Qty FROM VRawMaterialIssuance WHERE VendID=" & cmbMaker.ID & " AND DT BETWEEN '" & myDTFrom & "' AND '" & myDTTo & "' GROUP BY RMID1,RMName", con, adOpenForwardOnly, adLockReadOnly
        End If
        Do Until .EOF
            Set ITM = LVHistory.ListItems.add(, , !RMID1 & "")
            ITM.ListSubItems.add , , !RMName & ""
            ITM.ListSubItems.add , , Val(!Qty & "")
            .MoveNEXT
        Loop
        .Close
       
    End With
    Set rs = Nothing
    
End Sub

Private Sub chkShowPicture_Click()

    If chkShowPicture.Value = vbChecked Then
        Call LVRMIss_ItemClick(LVRMIss.SelectedItem)
    Else
        PicProduct.Visible = False
    End If
    
End Sub


Private Sub LVRMIss_ItemClick(ByVal Item As MSComctlLib.ListItem)

    If chkShowPicture.Value = vbChecked Then
        Dim rs As New ADODB.Recordset
        With rs
            .Open "SELECT Pic FROM RM WHERE RMID=" & Val(LVRMIss.SelectedItem.SubItems(1)), con, adOpenForwardOnly, adLockReadOnly
            If Not IsNull(![Pic]) Then
                Set ImgProduct.DataSource = rs
                ImgProduct.DataField = "Pic"
            Else
                ImgProduct.Picture = LoadPicture("")
            End If
            .Close
        End With
        Set rs = Nothing
        Dim lTop As Long
        lTop = LVRMIss.Top + LVRMIss.SelectedItem.Top
        If (lTop + PicProduct.Height) > LVRMIss.Top + LVRMIss.Height Then
            lTop = LVRMIss.Top + LVRMIss.Height - PicProduct.Height
        End If
        PicProduct.Move LVRMIss.Left + LVRMIss.Width - PicProduct.Width, lTop
        PicProduct.Visible = True
    Else
'        If PicProduct.Visible = True Then PicProduct.Visible = False
    End If
End Sub


