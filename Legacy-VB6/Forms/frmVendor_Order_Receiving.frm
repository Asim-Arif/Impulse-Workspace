VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmVendor_Order_Receiving 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendor Receiving"
   ClientHeight    =   7935
   ClientLeft      =   1095
   ClientTop       =   285
   ClientWidth     =   11910
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
   ScaleHeight     =   7935
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin VB.Frame FraRcv 
      BackColor       =   &H00E7EBEF&
      Caption         =   "Receiving :"
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
      Height          =   2835
      Left            =   1215
      TabIndex        =   20
      Top             =   1965
      Visible         =   0   'False
      Width           =   9360
      Begin ComboList.Usercontrol1 cmbLocation 
         Height          =   285
         Left            =   4845
         TabIndex        =   31
         Top             =   540
         Width           =   3540
         _ExtentX        =   6244
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
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1980
         TabIndex        =   28
         Top             =   540
         Width           =   1875
      End
      Begin VB.TextBox txtLotNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   90
         TabIndex        =   27
         Top             =   540
         Width           =   1875
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3870
         TabIndex        =   30
         Top             =   540
         Width           =   960
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   525
         Left            =   8430
         TabIndex        =   33
         Top             =   270
         Width           =   870
      End
      Begin VB.CommandButton cmdSaveRcv 
         Caption         =   "&Save"
         Height          =   345
         Left            =   105
         TabIndex        =   23
         Top             =   2340
         Width           =   870
      End
      Begin VB.TextBox txtTotalQtyWastage 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3870
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2385
         Width           =   960
      End
      Begin VB.CommandButton cmdCancelRcv 
         Caption         =   "Cancel"
         Height          =   345
         Left            =   1050
         TabIndex        =   21
         Top             =   2340
         Width           =   870
      End
      Begin MSComctlLib.ListView LVRcv_All 
         Height          =   825
         Left            =   1125
         TabIndex        =   24
         Top             =   1290
         Visible         =   0   'False
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   1455
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "LotNo"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "BatchNo"
            Object.Width           =   1693
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Qty"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView LVRcv_Current 
         Height          =   1440
         Left            =   90
         TabIndex        =   25
         Top             =   840
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   2540
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Lot No."
            Object.Width           =   3307
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Batch No."
            Object.Width           =   3307
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Location"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   0
         EndProperty
      End
      Begin MSForms.Label Label8 
         Height          =   240
         Index           =   3
         Left            =   4845
         TabIndex        =   34
         Top             =   285
         Width           =   3540
         BackColor       =   11517387
         Caption         =   "Location"
         Size            =   "6244;423"
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
      Begin MSForms.Label Label8 
         Height          =   240
         Index           =   2
         Left            =   3870
         TabIndex        =   32
         Top             =   285
         Width           =   960
         BackColor       =   11517387
         Caption         =   "Qty"
         Size            =   "1693;423"
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
      Begin MSForms.Label Label8 
         Height          =   240
         Index           =   1
         Left            =   1980
         TabIndex        =   29
         Top             =   285
         Width           =   1875
         BackColor       =   11517387
         Caption         =   "Batch No."
         Size            =   "3307;423"
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
      Begin MSForms.Label Label8 
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   26
         Top             =   285
         Width           =   1875
         BackColor       =   11517387
         Caption         =   "Lot No."
         Size            =   "3307;423"
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
   End
   Begin ComboList.Usercontrol1 cmbMaterial 
      Height          =   285
      Left            =   4575
      TabIndex        =   19
      Top             =   645
      Width           =   6885
      _ExtentX        =   12144
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
   Begin VB.TextBox txtBarcode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1335
      TabIndex        =   17
      Top             =   645
      Width           =   2445
   End
   Begin ComboList.Usercontrol1 cmbOrderNo 
      Height          =   285
      Left            =   8655
      TabIndex        =   13
      Top             =   1185
      Width           =   2820
      _ExtentX        =   4974
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
   Begin ComboList.Usercontrol1 cmbVend 
      Height          =   285
      Left            =   3795
      TabIndex        =   0
      Top             =   1185
      Width           =   4860
      _ExtentX        =   8573
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
   Begin VB.CheckBox chkPrint 
      Caption         =   "Print Receiving Note"
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
      Height          =   240
      Left            =   510
      TabIndex        =   5
      Top             =   5880
      Value           =   1  'Checked
      Width           =   4035
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
      Height          =   675
      Left            =   5070
      TabIndex        =   10
      Top             =   5610
      Width           =   5280
      Begin MSForms.CommandButton cmdSave 
         Height          =   360
         Left            =   75
         TabIndex        =   2
         Top             =   210
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
         TabIndex        =   4
         Top             =   210
         Width           =   1665
         Caption         =   "Close"
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
         TabIndex        =   3
         Top             =   210
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
      Begin VB.Label lbltotalBalls 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0 "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3945
         TabIndex        =   12
         Top             =   255
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Balls Issued :"
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
         Height          =   195
         Index           =   1
         Left            =   2340
         TabIndex        =   11
         Top             =   285
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4050
      Left            =   435
      TabIndex        =   1
      Top             =   1455
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   7144
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item ID"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Item Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Item Type"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty Ordered"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Already Rcvd."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Balance"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Qty Rcvd"
         Object.Width           =   2117
      EndProperty
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
      Left            =   15
      TabIndex        =   6
      Top             =   30
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Receiving"
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
         Left            =   4020
         TabIndex        =   7
         Top             =   120
         Width           =   2745
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Receiving"
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
         Left            =   4035
         TabIndex        =   8
         Top             =   135
         Width           =   2745
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   2175
      TabIndex        =   14
      Top             =   1170
      Width           =   1635
      _ExtentX        =   2884
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
      Format          =   148176899
      CurrentDate     =   -8461
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Material :"
      Height          =   195
      Index           =   1
      Left            =   3855
      TabIndex        =   18
      Top             =   690
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Barcode # :"
      Height          =   195
      Index           =   0
      Left            =   420
      TabIndex        =   16
      Top             =   690
      Width           =   855
   End
   Begin MSForms.TextBox txtRcvNo 
      Height          =   285
      Left            =   420
      TabIndex        =   15
      Top             =   1185
      Width           =   1755
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "3096;503"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label3 
      Height          =   240
      Left            =   420
      TabIndex        =   9
      Top             =   960
      Width           =   11055
      BackColor       =   11517387
      Caption         =   $"frmVendor_Order_Receiving.frx":0000
      Size            =   "19500;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmVendor_Order_Receiving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VendCapacity As Long
Dim IsItemsFirst As Boolean
Dim strRcvNo As String
Dim lShelfNoForAutoPlacement As Long
Dim bSample As Boolean

Private Sub chkOrderNo_Click()
'   Call cmbOrderNo_Click
   
      For i = 1 To LV.ListItems.count
         With LV.ListItems(i)
            .SubItems(1) = IIf(chkOrderNo = vbChecked, Replace(.ListSubItems(1).key, "'", ""), .ListSubItems(1).Tag)
         End With
      Next
   
End Sub


Private Sub cmbMaterial_matched()
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).SubItems(1) = cmbMaterial.ID Then
            LV.ListItems(i).Selected = True
            LV.ListItems(i).EnsureVisible
        End If
    Next
End Sub

Private Sub cmbStoreName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    If KeyCode.Value = 13 Then
        If cmbStoreName.MatchFound = False Then
            MsgBox "Please Select Store.", vbInformation
            Exit Sub
        End If
        If bSample Then
            LV.SelectedItem.SubItems(5) = Val(txtQty) & " /" & LV.SelectedItem.ListSubItems(5).Tag
            LV.SelectedItem.SubItems(6) = cmbStoreName.Text
            PicEdit.Visible = False
        End If
    ElseIf KeyCode.Value = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub

Private Sub cmbVend_matched()

    cmbOrderNo.ClearVals
    
    cmbOrderNo.AddVals con, "DISTINCT OrderNo", "VVendor_Orders_Items_Receiving", "OrderNo", " WHERE Vend_AccNo='" & cmbVend.ID & "' AND OrderQty-QtyRcvd>0"
   
End Sub

Private Sub cmbVend_UnMatched()
        'cmbVend.ClearVals
        LV.ListItems.Clear
End Sub


Private Sub cmbOrderNo_matched()

    On Error GoTo err
    Dim LastGroup As String
    Dim rs As New ADODB.Recordset
    
    With rs
        
        .Open "SELECT * FROM VVendor_Orders_Items_Receiving WHERE OrderNo='" & cmbOrderNo.ID & "' AND OrderQty-QtyRcvd>0 ORDER BY ItemID", con, adOpenForwardOnly, adLockReadOnly
        
        Dim ITM As ListItem
        LV.ListItems.Clear
        
        Do Until .EOF
            
            Set ITM = LV.ListItems.add(, , !ITemID & "")
            ITM.Tag = Val(!ID & "")
            
            ITM.ListSubItems.add , , ![ItemName] & ""
            ITM.ListSubItems.add(, , ![Item_Finishing_Type_Text] & "").Tag = Val(!Item_Finishing_Type & "")
            
            ITM.ListSubItems.add , , Val(!OrderQty & "")
            ITM.ListSubItems.add , , Val(!QtyRcvd & "")
            ITM.ListSubItems.add , , Val(!QtyBalance & "")
            ITM.ListSubItems.add , , ""
            
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    cmbMaterial.ClearVals
    cmbMaterial.AddVals con, "'{' + ItemID + '} ' + ItemName", "VVendor_Orders_Items_Receiving", "ItemID", " WHERE OrderNo='" & cmbOrderNo.ID & "'"
    
    
Exit Sub

err:
   MsgBox err.Description
End Sub



Private Sub cmdAdd_Click()

    If Trim(txtLotNo) = "" Then
        MsgBox "Please Enter Lot No.", vbInformation
        Exit Sub
    ElseIf Trim(txtBatchNo) = "" Then
        MsgBox "Please Enter Batch No.", vbInformation
        Exit Sub
    ElseIf Val(txtQty) <= 0 Then
        MsgBox "Invalid Qty.", vbInformation
        Exit Sub
    ElseIf cmbLocation.MatchFound = False Then
        MsgBox "Please Select Location.", vbInformation
        Exit Sub
    End If
    
    Dim ITM As ListItem
    Set ITM = LVRcv_Current.ListItems.add(, , Trim(txtLotNo))
    ITM.ListSubItems.add , , Trim(txtBatchNo)
    ITM.ListSubItems.add(, , Val(txtQty)).Tag = Val(LV.SelectedItem.Tag)
    ITM.ListSubItems.add(, , cmbLocation.Text).Tag = cmbLocation.ID
     
    txtLotNo = ""
    txtBatchNo = ""
    txtQty = ""
    cmbLocation.ID = ""
    cmbLocation.HideList
    txtLotNo.SetFocus
    
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub




Private Sub cmdCancelRcv_Click()
    FraRcv.Visible = False
    LV.Enabled = True
End Sub

Private Sub cmdSave_Click()
    
    If Saved Then
        Unload Me
        Me.Show
    End If
    
End Sub

Private Sub cmdSaveClose_Click()
    
    If Saved Then
        Unload Me
    End If
    
End Sub

Private Sub cmdSaveRcv_Click()
        
    Dim i As Integer, lQty As Long
    For i = LVRcv_All.ListItems.count To 1 Step -1
        If LVRcv_All.ListItems(i).ListSubItems(2).Tag = Val(LV.SelectedItem.Tag) Then
            LVRcv_All.ListItems.Remove i
        End If
    Next
    
    'Now Add these.
    Dim ITM As ListItem
    For i = 1 To LVRcv_Current.ListItems.count
        Set ITM = LVRcv_All.ListItems.add(, , LVRcv_Current.ListItems(i).Text)
        ITM.ListSubItems.add(, , LVRcv_Current.ListItems(i).SubItems(1)).Tag = LVRcv_Current.ListItems(i).ListSubItems(1).Tag
        ITM.ListSubItems.add(, , LVRcv_Current.ListItems(i).SubItems(2)).Tag = LVRcv_Current.ListItems(i).ListSubItems(2).Tag
        ITM.ListSubItems.add(, , LVRcv_Current.ListItems(i).SubItems(3)).Tag = LVRcv_Current.ListItems(i).ListSubItems(3).Tag
        lQty = lQty + Val(LVRcv_Current.ListItems(i).SubItems(2))
    Next
    LV.SelectedItem.SubItems(6) = lQty
    FraRcv.Visible = False
    LV.Enabled = True
    
End Sub

Private Sub cmdUnAssign_Click()

    On Error GoTo err
    If LV.ListItems.count = 0 Then Exit Sub
    Dim iAffect As Integer
    
    'con.Execute "Delete From VendAssItems Where EntryID=" & Val(LV.SelectedItem.Key), iAffect
    'If iAffect > 0 Then
    '    'Call cmbNick_Click
    '    cmdUnAssign.Visible = False
    'End If
    
    
    LV.ListItems.Remove LV.SelectedItem.Index
    cmbVend.Enabled = LV.ListItems.count = 0
    Call GetTotal
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub DTPicker1_Change()
    txtRcvNo = GetNextRcvID()
End Sub


Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    If cmbVend.MatchFound = False Then Exit Sub
    LV.Enabled = False
      
    Dim i As Integer, ITM As ListItem
    LVRcv_Current.ListItems.Clear
    For i = 1 To LVRcv_All.ListItems.count
        If Val(LV.SelectedItem.Tag) = Val(LVRcv_All.ListItems(i).ListSubItems(2).Tag) Then
            Set ITM = LVRcv_Current.ListItems.add(, , LVRcv_All.ListItems(i).Text)
            ITM.ListSubItems.add , , LVRcv_All.ListItems(i).SubItems(1)
            ITM.ListSubItems.add(, , LVRcv_All.ListItems(i).SubItems(2)).Tag = Val(LVRcv_All.ListItems(i).ListSubItems(2).Tag)
            ITM.ListSubItems.add(, , LVRcv_All.ListItems(i).SubItems(3)).Tag = Val(LVRcv_All.ListItems(i).ListSubItems(3).Tag)
        End If
    Next
    
    
    FraRcv.Visible = True
    txtLotNo.SetFocus
    
End Sub

Private Sub Form_Load()

    cmbVend.ListHeight = 3000

    cmbOrderNo.ListHeight = 2600

    'cmbVend.AddVals con, "AccTitle", "VVendersANDMakers", "AccNo"
    cmbVend.AddVals con, "AccTitle+' {'+AccNo+'}'", "VVenders", "AccNo"
    
    DTPicker1 = GetServerDate
    Call DTPicker1_Change
    
    For i = 1 To LV.ColumnHeaders.count
        LV.ColumnHeaders(i).Tag = LV.ColumnHeaders(i).Width
    Next
     
    'Call AddToCombo(cmbStoreName, "StoreName", "Stores", " WHERE EntryID=43")
    
'    With cmbVehicleType
'        .AddItem "Motor Cycle"
'        .AddItem "Suzuki Van"
'        .AddItem "Bicycle"
'        .AddItem "Other"
'        .ListIndex = 0
'    End With
    cmbMaterial.ListHeight = 4000
    strRcvNo = ""
   
        
    bSample = False
        
    cmbLocation.AddVals con, "StoreName+' '+RackNo+' '+ShelfNo", "VStoreShelfs", "EntryID"
    
End Sub

Private Sub Form_Resize()

 On Error Resume Next
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
 
    Dim TitleBottom As Long

    TitleBottom = cmbVend.Top + cmbVend.Height 'fTitle.Top + fTitle.Height
 
    LV.Move 50, TitleBottom, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + Frame6.Height + 100)
    FraRcv.Move (LV.Width - FraRcv.Width) / 2, (LV.Height - FraRcv.Height) / 2
    Frame6.Move Me.ScaleWidth - Frame6.Width - 50, Me.ScaleHeight - Frame6.Height - 50

    chkPrint.Move LV.Left, Frame6.Top + 50

 
End Sub


Private Sub TextBox1_Change()

End Sub

Private Sub txtBarcode_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Dim rs As New ADODB.Recordset
        With rs
            .Open "SELECT VendID,OrderNo FROM VendOrders WHERE RunningPONo=" & txtBarcode.Text, con, adOpenForwardOnly, adLockReadOnly
            If .EOF = False Then
                cmbVend.ID = !VendID & ""
                cmbOrderNo.ID = !OrderNo & ""
            End If
            .Close
        End With
        Set rs = Nothing
    End If
    
End Sub

Private Sub txtQty_GotFocus()
    txtQty.SelStart = 0
    txtQty.SelLength = Len(txtQty.Text)
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If bSample Then
            LV.SelectedItem.SubItems(5) = Val(txtQty) & " /" & LV.SelectedItem.ListSubItems(5).Tag
            PicEdit.Visible = False
        Else
            cmbStoreName.SetFocus
        End If
    End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    
    If KeyAscii = 13 Then
        
            
            If bSample Then
                LV.SelectedItem.SubItems(5) = Val(txtQty) & " /" & LV.SelectedItem.ListSubItems(5).Tag
                PicEdit.Visible = False
            Else
                cmbStoreName.SetFocus
            End If
            
       
         
    ElseIf KeyAscii = 27 Then
    
        PicEdit.Visible = False
        
    End If
    
    Call OnlyNums(KeyAscii)
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Function Saved() As Boolean

    On Error GoTo err
    Dim i As Integer, j As Integer, ITM As ListItem
    
    Dim lVOR_EntryID As Long, lVOR_D_RefID As Long, lVOR_D_L_RefID As Long
    Call StartTrans(con)
    con.Execute "INSERT INTO Vendor_Orders_Receiving(Vend_AccNo,DT,UserName,MachineName) VALUES('" & _
     cmbVend.ID & "','" & DTPicker1.Value & "','" & CurrentUserName & "','" & strComputerName & "')"
    
    lVOR_EntryID = GetSingleLongValue("MAX(EntryID)", "Vendor_Orders_Receiving", " WHERE MachineName='" & strComputerName & "'")
    
    Dim strLotNo As String, strBatchNo As String, lQty As Long, lShelf_RefID As Long, lProcessID As Long
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
        
            If Val(.SubItems(6)) > 0 Then
                con.Execute "INSERT INTO Vendor_Orders_Receiving_Detail(VOR_RefID,VOI_RefID,Qty) VALUES(" & lVOR_EntryID & "," & _
                 Val(.Tag) & "," & Val(.SubItems(6)) & ")"
                 
                lVOR_D_RefID = GetSingleLongValue("MAX(EntryID)", "Vendor_Orders_Receiving_Detail", " WHERE VOR_RefID=" & lVOR_EntryID)
                
                For j = 1 To LVRcv_All.ListItems.count
                    If Val(.Tag) = Val(LVRcv_All.ListItems(j).ListSubItems(2).Tag) Then
                        strLotNo = LVRcv_All.ListItems(j).Text
                        strBatchNo = LVRcv_All.ListItems(j).SubItems(1)
                        lQty = Val(LVRcv_All.ListItems(j).SubItems(2))
                        lShelf_RefID = Val(LVRcv_All.ListItems(j).ListSubItems(3).Tag)
                        
                        con.Execute "INSERT INTO Vendor_Orders_Receiving_Detail_Lots(VOR_D_RefID,LotNo,BatchNo,Qty,Shelf_RefID) VALUES(" & _
                         lVOR_D_RefID & ",'" & strLotNo & "','" & strBatchNo & "'," & lQty & "," & lShelf_RefID & ")"
                         
                        lVOR_D_L_RefID = GetSingleLongValue("MAX(EntryID)", "Vendor_Orders_Receiving_Detail_Lots", " WHERE VOR_D_RefID=" & lVOR_D_RefID)
                        
'                        If Val(.ListSubItems(2).Tag) = 0 Then   'Finished Stock
'                            Call ReceiveFinishedStock(lVOR_D_L_RefID, strLotNo, strBatchNo, lQty, lShelf_RefID, .Text)
'                        Else
'                            If Val(.ListSubItems(2).Tag) = 1 Then   'Temper
'                                lProcessID = iTemper_Inspection_Process_ID
'                            ElseIf Val(.ListSubItems(2).Tag) = 2 Then   'Inspection
'                                lProcessID = iFirst_Inspection_Process_ID
'                            End If
'                            Call ReceiveSemiFinishStock(lVOR_D_L_RefID, strLotNo, strBatchNo, lQty, lShelf_RefID, .Text, lProcessID)
'                        End If
                        
                    End If
                Next
                
            End If
            
        End With
    Next
        
    con.CommitTrans
    
    Unload Me
    Exit Function
err:
    MsgBox err.Description
    
End Function

Private Sub ReceiveSemiFinishStock(lVOR_D_L_RefID As Long, strLotNo As String, strBatchNo As String, lQty As Long, lShelf_RefID As Long, strItemID As String, lProcessID As Long)
    
    con.Execute "INSERT INTO StockOrderOpening(ItemID,ProcessID,Qty,Location,UserName,MachineName,Shelf_RefID,Remarks,LotNo_Manual,BatchNo_Manual,Mill_Certificate_No,VOR_D_L_RefID) VALUES('" & strItemID & _
         "'," & lProcessID & "," & lQty & ",'','" & CurrentUserName & "','" & strComputerName & "'," & _
         lShelf_RefID & ",'','" & strLotNo & "','" & strBatchNo & "',''," & lVOR_D_L_RefID & ")"
    
    con.Execute "INSERT INTO Lots_List(LotNo,ItemID,Lot_Type,Reference_LotNo,Batch_No,Mill_Certificate_No) VALUES('" & _
        strLotNo & "','" & strItemID & "',3,'','" & strBatchNo & "','')"
    
    Dim lEntryID As Long, lRcvEntryID As Long, lNextProcessID As Long
    lEntryID = GetSingleLongValue("MAX(EntryID)", "StockOrderOpening")
    
    con.Execute "INSERT INTO VendReceived(VendID,DT,RecieptID,UserID,ProcessID,Issuance_RefID,UserName,MachineName) VALUES(" & _
     "0,'" & DTPicker1.Value & "','',0," & lProcessID & ",0,'" & CurrentUserName & "','" & strComputerName & "')"
    
    lRcvEntryID = GetSingleLongValue("MAX(EntryID)", "VendReceived", " WHERE MachineName='" & strComputerName & "'")
    
    lNextProcessID = getNextProcessID(strItemID, lProcessID)
    
    con.Execute "INSERT INTO VendRcvdDetail(RefID,ItemCode,RecieptID,RcvdQty,Rate,LotNo,NextProcessID,OrderNo" & _
     ",CountedBy,Issue_RefID,ProcessID,RcvdWeight,Opening_RefID) VALUES(" & lRcvEntryID & ",'" & strItemID & _
     "',''," & lQty & ",0,'0'," & lNextProcessID & ",'Stock-Order','',0," & lProcessID & _
     ",0," & lEntryID & ")"
     
     
End Sub

Private Sub ReceiveFinishedStock(lVOR_D_L_RefID As Long, strLotNo As String, strBatchNo As String, lQty As Long, lShelf_RefID As Long, strItemID As String)
    
    Dim strRcvNo As String, lRISD_EntryID As Long
    strRcvNo = GetNextRFRcvID
    
    con.Execute "INSERT INTO RcvItemsSimple(RcvNo,DT,OrderNo,RcvdBy,RcvdFrom,UserName,ComputerName,IssNo,CustCode,VOR_D_L_RefID) VALUES('" & _
     strRcvNo & "','" & DTPicker1.Value & "','','','','" & CurrentUserName & "','" & strComputerName & "','','Stock'," & lVOR_D_L_RefID & ")"
    
    con.Execute "INSERT INTO RcvItemsSimpleDetail(RcvNo,ItemID,Qty,InHand,LedgerRefID,OrderNo,Location,Remarks,LotNo_Manual,BatchNo_Manual,Mill_Certificate_No) VALUES('" & _
     strRcvNo & "','" & strItemID & "'," & lQty & ",0,0,'','','','" & strLotNo & "','" & strBatchNo & "','')"
     
    con.Execute "INSERT INTO Lots_List(LotNo,ItemID,Lot_Type,Reference_LotNo,Batch_No,Mill_Certificate_No) VALUES('" & strLotNo & _
     "','" & strItemID & "',3,'','" & strBatchNo & "','')"
     
    lRISD_EntryID = GetSingleLongValue("MAX(EntryID)", "RcvItemsSimpleDetail")
    con.Execute "INSERT INTO RcvItemsSimpleDetail_Placement(RISD_RefID,Shelf_RefID,RcvdQty,Remarks) VALUES(" & _
     lRISD_EntryID & "," & lShelf_RefID & "," & lQty & ",'')"
             
End Sub

Private Sub AddForReadyFinishStockReceiving(strMaterialID As String, dQty As Double, strLocation As String, lVRDPO_RefID)
    
    
    If strRcvNo = "" Then
        strRcvNo = GetNextRFRcvID
        
        con.Execute "INSERT INTO RcvItemsSimple(RcvNo,DT,OrderNo,RcvdBy,RcvdFrom,UserName,ComputerName) VALUES('" & _
         strRcvNo & "','" & DTPicker1 & " " & ServerTime & "','Received against" & cmbOrderNo.ID & "','Purchased','" & _
         "','" & CurrentUserName & "','" & strComputerName & "')"
    
    End If
    
    Dim lRefID As Long
    lRefID = con.Execute("SELECT EntryID FROM RcvItemsSimple WHERE RcvNo='" & strRcvNo & "'").Fields(0).Value & ""
    Dim Dated As Date
    Dated = DTPicker1.Value
        
    dBal = getItemBalance(strMaterialID, Dated)
    dBal = dBal + dQty
    theDesc = "Item Received From " & cmbVend.Text & " against PO No.: " & cmbOrderNo.ID
            
    con.Execute "INSERT INTO ItemLedgerSimple(ItemID,DT,Description,Qty,Balance,OrderNo,Location) VALUES('" & _
     strMaterialID & "','" & Dated & "','" & theDesc & "'," & dQty & "," & dBal & ",'" & cmbOrderNo.ID & "','" & strLocation & "')"
     lRefID = Val(con.Execute("Select Max(EntryID) From ItemLedgerSimple").Fields(0).Value & "")
            
    Call adjustItemLedger(strMaterialID, Dated, dQty)

    con.Execute "INSERT INTO RcvItemsSimpleDetail(RcvNo,ItemID,Qty,InHand,LedgerRefID,OrderNo,Location,VRDPO_RefID) VALUES('" & _
     strRcvNo & "','" & strMaterialID & "'," & dQty & "," & dBal & "," & lRefID & ",'" & cmbOrderNo.Text & _
     "','" & strLocation & "'," & lVRDPO_RefID & ")"

End Sub

Private Function GetNextRFRcvID() As String

    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Max(Cast(Right(RcvNo,3) As Int)) From RcvItemsSimple Where Convert(Char(8),DT,1)='" & Format(DTPicker1, "MM/dd/yy") & "'", con, adOpenForwardOnly, adLockReadOnly
        
        GetNextRFRcvID = "RCV-" & Format(DTPicker1, "yyMMdd-") & Format(Val(.Fields(0) & "") + 1, "000")
        
        .Close
    End With
    Set rs = Nothing
    
End Function
Private Function GetNextEntryID() As String
    GetNextEntryID = GetMax("EntryID", "VendQuots") + 1
End Function


Private Sub PrintReport()

    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\GRN.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VendRcvd.RcvID}='" & txtRcvNo.Text & "'", rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub GetTotal()

   Exit Sub

   Dim total As Double
   For i = 1 To LV.ListItems.count
      total = total + Val(LV.ListItems(i).SubItems(7))
   Next
   lbltotalBalls = total
End Sub



Private Function GetNextRcvID() As String

    Dim Prefix As String
    'RCV-DDMM-001
    'Prefix = "RCV-" & Format(GetServerDate, "ddMM") & "-"
    Prefix = "RCV-" & Format(DTPicker1.Value, "ddMMyy") & "-"
    GetNextRcvID = Prefix & Format(GetMax("right(RTrim(RcvID),3)", "VendRcvd", "Where left(RcvID," & Len(Prefix) & ")='" & Prefix & "'") + 1, "000")
    
End Function

Private Sub txtRate_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        
        If Val(txtRate) <= 0 Then
            MsgBox "Invalid Rate.", vbInformation
            Exit Sub
        End If
        
        If Val(txtQty) > Val(LV.SelectedItem.SubItems(4)) Then
            If MsgBox("You should not Receive More Than Qty To Receive" & _
            "Do you want to contine?", vbInformation + vbYesNo) = vbNo Then
                KeyAscii = 0
                Exit Sub
            Else
                If getDBPassword("Access_Receiving") = False Then Exit Sub
            End If
        End If
        
        If txtSheets.Enabled = False Then
            LV.SelectedItem.SubItems(5) = Val(txtQty) & " /" & LV.SelectedItem.ListSubItems(5).Tag
            LV.SelectedItem.SubItems(6) = cmbStoreName.Text
            LV.SelectedItem.SubItems(7) = Val(txtRate)
            
            PicEdit.Visible = False
            
            Call GetTotal
        Else
            txtSheets.SetFocus
        End If
        
    ElseIf KeyAscii = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub

Private Sub txtSheets_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        If Val(txtSheets) <= 0 Then
            MsgBox "Invalid Sheet Qty.", vbInformation
            Exit Sub
        End If
        
        
        LV.SelectedItem.SubItems(5) = Val(txtQty) & " /" & LV.SelectedItem.ListSubItems(5).Tag
        LV.SelectedItem.SubItems(6) = cmbStoreName.Text
        LV.SelectedItem.SubItems(7) = Val(txtRate)
        LV.SelectedItem.SubItems(8) = Val(txtSheets)
        
        PicEdit.Visible = False
        
        Call GetTotal
        
    ElseIf KeyAscii = 27 Then
        PicEdit.Visible = False
    End If
End Sub


Public Sub ShowMe(Optional p_bSample As Boolean = False)
    bSample = p_bSample
    If bSample Then
        Label1(0).Caption = "Receive Samples"
        Label1(1).Caption = "Receive Samples"
    End If
    Me.Show
End Sub
