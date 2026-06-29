VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmVendGateRcvd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vender Gate Receiving"
   ClientHeight    =   6645
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
   ScaleHeight     =   6645
   ScaleWidth      =   11910
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
      Left            =   8685
      ScaleHeight     =   2835
      ScaleWidth      =   2760
      TabIndex        =   28
      Top             =   1455
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
      Left            =   3105
      TabIndex        =   27
      Top             =   5880
      Width           =   1560
   End
   Begin ComboList.Usercontrol1 cmbMaterial 
      Height          =   285
      Left            =   4575
      TabIndex        =   24
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
      TabIndex        =   22
      Top             =   645
      Width           =   2445
   End
   Begin VB.PictureBox PicEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4530
      ScaleHeight     =   285
      ScaleWidth      =   3405
      TabIndex        =   18
      Top             =   3375
      Visible         =   0   'False
      Width           =   3435
      Begin VB.TextBox txtSheets 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   25
         Top             =   0
         Width           =   1000
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   20
         Top             =   0
         Width           =   1200
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   1200
      End
   End
   Begin ComboList.Usercontrol1 cmbOrderNo 
      Height          =   285
      Left            =   6465
      TabIndex        =   14
      Top             =   1185
      Width           =   2010
      _ExtentX        =   3545
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
      Width           =   2685
      _ExtentX        =   4736
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
      TabIndex        =   6
      Top             =   5880
      Value           =   1  'Checked
      Width           =   2220
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
      TabIndex        =   11
      Top             =   5610
      Width           =   5280
      Begin MSForms.CommandButton cmdSave 
         Height          =   360
         Left            =   75
         TabIndex        =   3
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   285
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdUnAssign 
      Height          =   240
      Left            =   10410
      Picture         =   "frmVendGateRcvd.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "UnAssign Selected Item..."
      Top             =   2070
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   435
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4050
      Left            =   420
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "WO No."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Material ID"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Material Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Qty Ordered"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Qty To Rcv"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Qty Rcvd"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Store Name"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Rate"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Sheet/Bar"
         Object.Width           =   1764
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
      TabIndex        =   7
      Top             =   30
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vender Gate Receiving"
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
         TabIndex        =   8
         Top             =   120
         Width           =   3555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vender Gate Receiving"
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
         TabIndex        =   9
         Top             =   135
         Width           =   3555
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   2175
      TabIndex        =   15
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
      Format          =   129761283
      CurrentDate     =   -8461
   End
   Begin MSForms.ComboBox cmbStoreName 
      Height          =   285
      Left            =   2970
      TabIndex        =   26
      Top             =   5520
      Visible         =   0   'False
      Width           =   1395
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "2461;503"
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
      Caption         =   "Material :"
      Height          =   195
      Index           =   1
      Left            =   3855
      TabIndex        =   23
      Top             =   690
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Barcode # :"
      Height          =   195
      Index           =   0
      Left            =   420
      TabIndex        =   21
      Top             =   690
      Width           =   855
   End
   Begin MSForms.TextBox txtDesc 
      Height          =   285
      Left            =   8460
      TabIndex        =   17
      Top             =   1185
      Width           =   3015
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5318;503"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtRcvNo 
      Height          =   285
      Left            =   420
      TabIndex        =   16
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
      TabIndex        =   10
      Top             =   960
      Width           =   11055
      BackColor       =   11517387
      Caption         =   $"frmVendGateRcvd.frx":0102
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
Attribute VB_Name = "frmVendGateRcvd"
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
    If bSample Then
        cmbOrderNo.AddVals con, "Distinct PONo + ' (' + CAST(RunningPONo AS VARCHAR(50)) + ')'", "VVendOrdersToRcv", "PONo", " WHERE VendID='" & cmbVend.ID & "' AND SampleOrder=1"
    Else
        cmbOrderNo.AddVals con, "Distinct PONo + ' (' + CAST(RunningPONo AS VARCHAR(50)) + ')'", "VVendOrdersToRcv", "PONo", " WHERE VendID='" & cmbVend.ID & "'"
    End If
   
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
        
        .Open "SELECT * FROM VVendOrdersToRcv WHERE PONo='" & cmbOrderNo.ID & "' ORDER BY GroupID,MaterialName", con, adOpenForwardOnly, adLockReadOnly
        Dim ITM As ListItem
        LV.ListItems.Clear
        
        Do Until .EOF
            
            Set ITM = LV.ListItems.add(, , "")
            ITM.Tag = 0 'Val(![WORef] & "")
            
            'Display the Customer Order No if chkOrderNo is checked Else Show the Company Order No
            ITM.ListSubItems.add(, , ![MaterialID]).Tag = Val(!Rate & "")
            ITM.ListSubItems.add , , ![MaterialName] & ""
            ITM.ListSubItems.add(, , Val(![QtyOrdered]) & " " & ![Unit]).Tag = Abs(!CCItem)
            ITM.ListSubItems.add , , Val(![QtyToRcv]) & " " & ![Unit]
            ITM.ListSubItems.add().Tag = ![Unit] & ""
            ITM.ListSubItems.add().Tag = ""
            ITM.ListSubItems.add(, , Val(!Rate & "")).Tag = Val(!Rate & "")
            'Set Tool Tips
            
            ITM.ListSubItems(4).Tag = Val(!OtherUnitEntryID & "")
            
            For i = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(i).ToolTipText = ITM.ListSubItems(i).Text
            Next
            
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    cmbMaterial.ClearVals
    cmbMaterial.AddVals con, "'{' + MaterialID + '} ' + MaterialName", "VVendOrdersToRcv", "MaterialID", " WHERE PONo='" & cmbOrderNo.ID & "' ORDER BY GroupID,MaterialName"
    
    
Exit Sub

err:
   MsgBox err.Description
End Sub



Private Sub cmdAdd_Click()
    On Error GoTo err
    
    
    If Val(txtIssQty) > Val(txtQty) Then
        MsgBox "invalid quantity", vbInformation
        Exit Sub
    ElseIf Not cmbItems.MatchFound Then
        MsgBox "Select an Item first", vbInformation
        Exit Sub
    End If
    
    Dim ITM As ListItem
    
    Set ITM = LV.ListItems.add(, cmbVend.ID & ";" & cmbOrders.Text & ";" & cmbItems.Text, cmbVend.Text & " (" & cmbOrderNo.Text & ")")
    ITM.ListSubItems.add , , cmbOrders.Text
    ITM.ListSubItems.add , , cmbItems 'Format(DTPicker1.value, "dd-MMM-yyyy")
    'ITM.ListSubItems.Add(, , cmbProcess.Text).Tag = cmbProcess.List(cmbProcess.ListIndex, 1)
    ITM.ListSubItems.add(, , cmbVend.Text).Tag = cmbVend.ID
    ITM.ListSubItems.add , , txtQty1
    ITM.ListSubItems.add , , txtIssQty
    'itm.ListSubItems.Add , , txtRecNo
           
    cmbVend.Enabled = LV.ListItems.count = 0
    txtQty1 = ""
    txtQty = ""
    txtIssQty = ""
    Exit Sub
err:
If err.Number = 35602 Then
    MsgBox "This Item is Already Issued", vbInformation
Else
    MsgBox err.Description
End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
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

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    If cmbVend.MatchFound = False Then Exit Sub

    If Not LV.SelectedItem.Checked Then
        LV.SelectedItem.Checked = True
        'Call LV_ItemCheck(LV.SelectedItem)
        Exit Sub
    End If

    For i = 1 To LV.ColumnHeaders.count
        LV.ColumnHeaders(i).Width = Val(LV.ColumnHeaders(i).Tag)
    Next
    
    If Val(LV.SelectedItem.ListSubItems(4).Tag) > 0 Then
        txtSheets.Enabled = True
    Else
        txtSheets.Enabled = False
    End If
    
    With PicEdit
        .Move LV.Left + LV.ColumnHeaders(6).Left, LV.Top + LV.SelectedItem.Top
        .Visible = True
    End With
    
    If LV.SelectedItem.ListSubItems(3).Tag = "1" Then
        Call AddToCombo(cmbStoreName, "DISTINCT Location", "IssItemsSimpleDetail")
    Else
        Call AddToCombo(cmbStoreName, "StoreName", "Stores", " WHERE EntryID=43")
    End If
    
    With cmbStoreName
        .Text = LV.SelectedItem.SubItems(6)
    End With
    txtRate = LV.SelectedItem.SubItems(7)
    If txtSheets.Enabled Then
        txtSheets = LV.SelectedItem.SubItems(8)
    End If
    
    With txtQty
        '.Move LV.Left + LV.ColumnHeaders(6).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(6).Width
        '.Text = Val(LV.SelectedItem.SubItems(5))
        '.Visible = True
        '.Font.Size = 8
        .SetFocus
    End With

    

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   'If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()

    cmbVend.ListHeight = 3000

    cmbOrderNo.ListHeight = 2600


    'cmbVend.AddVals con, "AccTitle", "VVendersANDMakers", "AccNo"
    cmbVend.AddVals con, "AccTitle", "VVenders", "AccNo"
    
    DTPicker1 = GetServerDate
    Call DTPicker1_Change
    
    For i = 1 To LV.ColumnHeaders.count
        LV.ColumnHeaders(i).Tag = LV.ColumnHeaders(i).Width
    Next
     
    Call AddToCombo(cmbStoreName, "StoreName", "Stores", " WHERE EntryID=43")
    
'    With cmbVehicleType
'        .AddItem "Motor Cycle"
'        .AddItem "Suzuki Van"
'        .AddItem "Bicycle"
'        .AddItem "Other"
'        .ListIndex = 0
'    End With
    cmbMaterial.ListHeight = 4000
    strRcvNo = ""
            
    If strCompany <> "Tecno" Then
        lShelfNoForAutoPlacement = GetSingleLongValue("MAX(EntryID)", "StoreShelfs")
    End If
        
    bSample = False
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
 
    Dim TitleBottom As Long

    TitleBottom = cmbVend.Top + cmbVend.Height 'fTitle.Top + fTitle.Height
 
    LV.Move 50, TitleBottom, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + Frame6.Height + 100)
 
    Frame6.Move Me.ScaleWidth - Frame6.Width - 50, Me.ScaleHeight - Frame6.Height - 50

    chkPrint.Move LV.Left, Frame6.Top + 50
    chkShowPicture.Move chkPrint.Left + chkPrint.Width + 100, chkPrint.Top
 
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)

    If chkShowPicture.Value = vbChecked Then
        Dim rs As New ADODB.Recordset
        With rs
            .Open "SELECT Pic FROM RM WHERE RMID1='" & (LV.SelectedItem.SubItems(1)) & "'", con, adOpenForwardOnly, adLockReadOnly
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
        lTop = LV.Top + LV.SelectedItem.Top
        If (lTop + PicProduct.Height) > LV.Top + LV.Height Then
            lTop = LV.Top + LV.Height - PicProduct.Height
        End If
        PicProduct.Move LV.Left + LV.Width - PicProduct.Width, lTop
        PicProduct.Visible = True
    Else
'        If PicProduct.Visible = True Then PicProduct.Visible = False
    End If
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
            'cmbStoreName.SetFocus
            txtRate.SetFocus
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
                'cmbStoreName.SetFocus
                txtRate.SetFocus
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

    Saved = False
    On Error GoTo err
    Dim TempIndex As Integer
    TempIndex = 1
    
    For i = 1 To LV.ListItems.count
        If Val(LV.ListItems(i).SubItems(5)) > 0 Then Exit For
    Next
    
'    If i = LV.ListItems.count + 1 Then
'        MsgBox "Cannot Save, No Rates Defined For Items In Quotation.", vbInformation
'        Exit Function
'    End If
 
    Dim iIssuanceType As Integer
    iIssuanceType = GetSingleLongValue("IssuanceType", "VendOrders", " WHERE OrderNo='" & cmbOrderNo.ID & "'")
    
    Dim lRefID As Long, dExtraQty As Double, dQty As Double
    Call StartTrans(con)
    
    'Refresh The Next EntryID may be Changed On Network
    'txtRcvNo = GetNextRcvID
    Dim lRunningRcvNo As Long
    lRunningRcvNo = GetSingleLongValue("MAX(RunningRcvNo)", "VendRcvd") + 1
    
    con.Execute "INSERT INTO VendRcvd(RcvID,VendID,RcvDate,Description,StockUpdated,PostingDate,RunningRcvNo) Values ('" & txtRcvNo & _
    "','" & cmbVend.ID & "','" & Format(DTPicker1, "dd-MMM-yyyy") & "','" & txtDesc & "',1,'" & GetServerDate(True) & "'," & lRunningRcvNo & ")"
     
    Dim lRMID As Long
    Dim theDesc As String, dBal As Double, strMaterialID As String
    Dim lEntryIDDetail As Long
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            dQty = Val(.SubItems(5))
            If dQty > 0 Then
                
                If bSample Then
                    dExtraQty = 0
                    dQty = dQty
                ElseIf dQty > Val(.SubItems(4)) Then
                    dExtraQty = dQty - Val(.SubItems(4))
                    dQty = dQty - dExtraQty
                Else
                    dExtraQty = 0
                End If
                strMaterialID = .SubItems(1)
                
                con.Execute "INSERT INTO VendRcvdDetailPO(RcvID,PORefNo,MaterialID,QtyRcvd,WORef,QtyReturned,Rate,Checked,StoreName,ExtraQtyRcvd,Sheets) " & _
                " VALUES('" & txtRcvNo & "','" & cmbOrderNo.ID & "','" & .SubItems(1) & "'," & dQty & _
                "," & Val(.Tag) & ",0," & Val(.ListSubItems(7)) & ",1,'" & .SubItems(6) & "'," & dExtraQty & "," & Val(.SubItems(8)) & ")"
                
                lEntryIDDetail = GetSingleLongValue("MAX(EntryID)", "VendRcvdDetailPO")
                If strCompany = "Banzai" Then           'Auto Placement...
                    
                    con.Execute "INSERT INTO MaterialLocationwiseStatus(Rcvd_RefID,Shelf_RefID,QtyPlaced,UserName,MachineName,SheetsPlaced) VALUES(" & _
                     lEntryIDDetail & "," & lShelfNoForAutoPlacement & "," & dQty & ",'" & CurrentUserName & "','" & strComputerName & "'," & Val(.SubItems(8)) & ")"
                End If
'                dBal = getItemBalance(strMaterialID, DTPicker1)
'                dBal = dBal + dQty + dExtraQty
'
'                TheDesc = "Received from " & cmbVend.Text
'                con.Execute "INSERT INTO ItemLedgerSimple(ItemID,DT,Description,Qty,Balance) VALUES('" & _
'                 strMaterialID & "','" & DTPicker1 & "','" & TheDesc & "'," & (dQty + dExtraQty) & "," & dBal & ")"

'                lRefID = Val(con.Execute("SELECT MAX(EntryID) FROM ItemLedgerSimple").Fields(0).Value & "")
'
'                Call adjustItemLedger(strMaterialID, DTPicker1, (dQty + dExtraQty))
                
                If Val(.ListSubItems(3).Tag) = 0 Then
                    If iIssuanceType = 0 Then
                        con.Execute "UPDATE RM SET QtyInStock=ISNULL(QtyInStock,0)+" & Val(.SubItems(5)) & ",Rate=" & Val(.ListSubItems(7)) & " WHERE RMID1='" & .SubItems(1) & "'"
                    ElseIf iIssuanceType = 1 Then
                        con.Execute "UPDATE RM SET QtyInStock=ISNULL(QtyInStock,0)+" & Val(.SubItems(5)) & ",Rate=" & Val(.ListSubItems(7)) & ",AnnealingStock=ISNULL(AnnealingStock,0)+" & Val(.SubItems(5)) & " WHERE RMID1='" & .SubItems(1) & "'"
                    ElseIf iIssuanceType = 2 Then
                        con.Execute "UPDATE RM SET QtyInStock=ISNULL(QtyInStock,0)+" & Val(.SubItems(5)) & ",Rate=" & Val(.ListSubItems(7)) & ",MachiningStock=ISNULL(MachiningStock,0)+" & Val(.SubItems(5)) & " WHERE RMID1='" & .SubItems(1) & "'"
                    ElseIf iIssuanceType = 3 Then
                        con.Execute "UPDATE RM SET QtyInStock=ISNULL(QtyInStock,0)+" & Val(.SubItems(5)) & ",Rate=" & Val(.ListSubItems(7)) & ",MaleGrindingStock=ISNULL(MaleGrindingStock,0)+" & Val(.SubItems(5)) & " WHERE RMID1='" & .SubItems(1) & "'"
                    End If
                
                    lRMID = GetSingleLongValue("RMID", "RM", " WHERE RMID1='" & .SubItems(1) & "'")
                    con.Execute "UPDATE RM SET Rate=" & Val(.ListSubItems(7)) & " WHERE RMID=" & lRMID
                Else
                    If strCompany <> "Dr-Frgz" Then
                        Call AddForReadyFinishStockReceiving(.SubItems(1), dQty, .SubItems(6), lEntryIDDetail)
                    End If
                    'con.Execute "UPDATE Items SET InHand=ISNULL(InHand,0)+" & Val(.SubItems(5)) & " WHERE ItemID='" & .SubItems(1) & "'"
                End If

'                con.Execute "UPDATE VenderAssItems SET Rate=" & Val(.ListSubItems(7)) & " WHERE VendID=" & cmbVend.ID & " AND RMID=" & lRMID
                
'                If StockUpdated Then
'                    QtyDiff = Val(.ListSubItems(5).Tag) - Val(.ListSubItems(5))
'                    con.Execute "Update RM Set QtyInStock=QtyInStock+ (" & QtyDiff & "),Rate=" & Val(.SubItems(7)) & " WHERE RMID1='" & .SubItems(1) & "'"
'                Else
'                    con.Execute "Update RM Set QtyInStock=QtyInStock+" & Val(.SubItems(5)) & ",Rate=" & Val(.SubItems(7)) & " WHERE RMID1='" & .SubItems(1) & "'"
'                End If
                
            End If
        End With
    Next
    
    
    con.CommitTrans
    Saved = True
    
    If chkPrint = vbChecked Then Call PrintReport
    
    Screen.MousePointer = vbDefault
    
    Exit Function
    
err:
    MsgBox err.Description
    con.RollbackTrans
    Screen.MousePointer = vbDefault
    
End Function

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
            'LV.SelectedItem.SubItems(6) = cmbStoreName.Text
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


Private Sub chkShowPicture_Click()

    If chkShowPicture.Value = vbChecked Then
        Call LV_ItemClick(LV.SelectedItem)
    Else
        PicProduct.Visible = False
    End If
    
End Sub


