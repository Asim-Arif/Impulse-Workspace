VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form frmStockOrderAdjustment 
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   45
   ClientWidth     =   11970
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
   ScaleHeight     =   8910
   ScaleMode       =   0  'User
   ScaleWidth      =   10721.69
   WindowState     =   2  'Maximized
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
      Height          =   8880
      Left            =   60
      TabIndex        =   0
      Top             =   -45
      Width           =   11865
      Begin VB.PictureBox PicEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5355
         Left            =   6675
         ScaleHeight     =   5325
         ScaleWidth      =   5040
         TabIndex        =   4
         Top             =   2985
         Visible         =   0   'False
         Width           =   5070
         Begin VB.TextBox txtTotalAdj 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   3870
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   4920
            Width           =   1140
         End
         Begin MSComctlLib.ListView LVAdjAll 
            Height          =   3480
            Left            =   165
            TabIndex        =   18
            Top             =   1260
            Visible         =   0   'False
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   6138
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Location"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Qty"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Customer"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView LVAdj 
            Height          =   3885
            Left            =   15
            TabIndex        =   17
            Top             =   1005
            Width           =   5010
            _ExtentX        =   8837
            _ExtentY        =   6853
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Location"
               Object.Width           =   7302
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Qty"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Customer"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4175
            TabIndex        =   5
            Top             =   705
            Width           =   850
         End
         Begin MSForms.CommandButton cmdCancel 
            Height          =   375
            Left            =   1095
            TabIndex        =   16
            Top             =   4905
            Width           =   1050
            ForeColor       =   0
            BackColor       =   16777215
            Caption         =   "Cancel"
            PicturePosition =   327683
            Size            =   "1852;661"
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdOK 
            Height          =   375
            Left            =   15
            TabIndex        =   15
            Top             =   4890
            Width           =   1050
            ForeColor       =   0
            BackColor       =   16777215
            Caption         =   "OK"
            PicturePosition =   327683
            Size            =   "1852;661"
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Issuable to Process"
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
            Left            =   15
            TabIndex        =   14
            Top             =   0
            Width           =   1650
         End
         Begin VB.Label Label2 
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
            Index           =   1
            Left            =   4170
            TabIndex        =   13
            Top             =   510
            Width           =   300
         End
         Begin VB.Label Label2 
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
            Index           =   0
            Left            =   15
            TabIndex        =   12
            Top             =   510
            Width           =   720
         End
         Begin MSForms.ComboBox cmbLocation_QEL 
            Height          =   285
            Left            =   15
            TabIndex        =   11
            Tag             =   "PaymentTerms"
            Top             =   705
            Width           =   4140
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "7302;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbProcess 
            Height          =   285
            Left            =   15
            TabIndex        =   6
            Top             =   210
            Width           =   5010
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "8837;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
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
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   11865
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Stock Order Adjustment"
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
            TabIndex        =   21
            Top             =   135
            Width           =   11850
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Stock Order Adjustment"
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
            TabIndex        =   22
            Top             =   150
            Width           =   11850
         End
      End
      Begin MSComCtl2.DTPicker DT 
         Height          =   300
         Left            =   7560
         TabIndex        =   9
         Top             =   8460
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   116588547
         CurrentDate     =   40493
      End
      Begin VB.TextBox txtOrderNo 
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
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   45
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   810
         Width           =   3135
      End
      Begin MSComctlLib.ListView LV 
         Height          =   7140
         Left            =   60
         TabIndex        =   8
         Top             =   1200
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   12594
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item Code"
            Object.Width           =   2090
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item Name"
            Object.Width           =   6667
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Ord Qty"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Adj."
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Net"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Stock"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Adj. Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Process"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date :"
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
         Left            =   6990
         TabIndex        =   10
         Top             =   8475
         Width           =   495
      End
      Begin MSForms.CommandButton cmdSave 
         Height          =   375
         Left            =   8985
         TabIndex        =   7
         Top             =   8400
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Save"
         PicturePosition =   327683
         Size            =   "2355;661"
         Accelerator     =   83
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label lbl 
         Caption         =   "Order No. :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   45
         TabIndex        =   2
         Top             =   615
         Width           =   1155
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   375
         Left            =   10440
         TabIndex        =   1
         Top             =   8385
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
End
Attribute VB_Name = "frmStockOrderAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iColNo As Integer
Dim strOrderNo As String, strCustCode As String

Private Sub cmdCancel_Click()
    PicEdit.Visible = False
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
    
    If cmbProcess.MatchFound = False Then
        MsgBox "Invalid Process.", vbInformation
        Exit Sub
    End If
    
    Dim i As Integer
    For i = LVAdjAll.ListItems.count To 1 Step -1
        If Val(LVAdjAll.ListItems(i).Tag) = Val(LV.SelectedItem.Index) Then
            LVAdjAll.ListItems.Remove i
        End If
    Next
    
    'Now Add these.
    Dim ITM As ListItem
    For i = 1 To LVAdj.ListItems.count
        Set ITM = LVAdjAll.ListItems.add(, , LVAdj.ListItems(i).Text)
        ITM.Tag = LV.SelectedItem.Index
        ITM.ListSubItems.add(, , LVAdj.ListItems(i).SubItems(1)).Tag = LVAdj.ListItems(i).ListSubItems(1).Tag
        ITM.ListSubItems.add(, , "").Tag = LVAdj.ListItems(i).ListSubItems(2).Tag
        ITM.ListSubItems.add(, , "").Tag = LVAdj.ListItems(i).SubItems(2)
    Next
    
    LV.SelectedItem.SubItems(6) = Val(txtTotalAdj)
    LV.SelectedItem.SubItems(7) = cmbProcess.Text
    LV.SelectedItem.ListSubItems(7).Tag = cmbProcess.List(cmbProcess.ListIndex, 1)
    
    PicEdit.Visible = False
        
    
End Sub

Private Sub cmdSave_Click()
    If Saved Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()

    iColNo = 0
    DT.Value = Date
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    FAcc.Top = (Me.ScaleHeight - FAcc.Height) / 2
   
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
        LV.ColumnHeaderIcons = Nothing
        LV.SortKey = ColumnHeader.Index - 1
        LV.SortOrder = lvwAscending
        'ColumnHeader.Icon = "Down"
    End If
    If LV.Sorted = False Then LV.Sorted = True
End Sub

Private Sub LV_DblClick()
    
    Dim i As Integer
        
    If LV.ListItems.count = 0 Then Exit Sub
    Dim lSNo As Long, strItemID As String
    strItemID = LV.SelectedItem.Text
    
    cmbProcess.Clear
    Call AddToCombo(cmbProcess, "Description", "VItemProcesses", " WHERE IsExist='" & strItemID & "' ORDER BY ItemSNo", , "ProcessID")
    
    If Val(LV.SelectedItem.ListSubItems(7).Tag) > 0 Then
        For i = 0 To cmbProcess.ListCount - 1
            If Val(cmbProcess.List(i, 1)) = Val(LV.SelectedItem.ListSubItems(7).Tag) Then
                cmbProcess.ListIndex = i
                Exit For
            End If
        Next
    End If
    
    '----------------------------------------------
    LVAdj.ListItems.Clear
    txtEdit = ""
    
    'Call Already Added.
    Dim ITM As ListItem
    For i = 1 To LVAdjAll.ListItems.count
        If Val(LVAdjAll.ListItems(i).Tag) = Val(LV.SelectedItem.Index) Then
            'Add These...
            Set ITM = LVAdj.ListItems.add(, , LVAdjAll.ListItems(i).Text)
            ITM.Tag = LVAdjAll.ListItems(i).Tag
            ITM.ListSubItems.add(, , LVAdjAll.ListItems(i).SubItems(1)).Tag = LVAdjAll.ListItems(i).ListSubItems(1).Tag
            ITM.ListSubItems.add(, , LVAdjAll.ListItems(i).SubItems(3)).Tag = LVAdjAll.ListItems(i).ListSubItems(2).Tag
        End If
    Next
    
    Call GetAdjTotalQty
    
    cmbLocation_QEL.Clear
    If strCompany = "QEL" Or strCompany = "Banzai" Then
        Call AddToCombo(cmbLocation_QEL, "Location + ' (' + CAST(NetQty AS VARCHAR(50)) + ')'", "VItemsWithLocationwiseStock", " WHERE ItemID='" & strItemID & "' AND NetQty>0", , "Location,NetQty")
    Else
        'Call AddToCombo(cmbLocation_QEL, "StoreName + ' ' + RackNo + ' ' + ShelfNo + ' (' + CAST(NetQty AS VARCHAR(50)) + ') {' + Remarks + '}'", "VItemsWithShelfWiseStock", " WHERE ItemID='" & strItemID & "' AND NetQty>0", , "Shelf_RefID,NetQty")
        Call AddToCombo(cmbLocation_QEL, "StoreName + ' ' + RackNo + ' ' + ShelfNo + ' (' + CAST(NetQty AS VARCHAR(50)) + ') {' + Remarks + '}'", "VItemsWithShelfWiseStock_CustomerWise", " WHERE CustCode IN('" & strCustCode & "','Stock') AND ItemID='" & strItemID & "' AND NetQty>0", , "Shelf_RefID,NetQty,CustCode")
    End If
    
    With PicEdit
        .Move (LV.Width + LV.Left) - PicEdit.Width, LV.Top + LV.SelectedItem.Top
        If PicEdit.Top > 2520 Then PicEdit.Top = 2520
'            PicEdit.Top = LV.Top + LV.Height - PicEdit.Height
'        End If
        .Visible = True
        txtEdit.Text = Val(LV.SelectedItem.SubItems(6))
        'cmbLocation_QEL.SetFocus
        cmbProcess.SetFocus
    End With
    
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    PicEdit.Visible = False
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        If X > LV.ColumnHeaders(1).Left And X < LV.ColumnHeaders(i).Left + LV.ColumnHeaders(i).Width Then
            iColNo = i
            Exit For
        End If
    Next
    
End Sub

Public Sub ShowMe(p_strOrderNo As String)
    
    strOrderNo = p_strOrderNo
    txtOrderNo = strOrderNo
    On Error GoTo err
    Dim rs As New ADODB.Recordset, ITM As ListItem
    Dim lAlreadyAdjQty As Long
    Dim lInHand_Stock As Long
    With rs
        '.Open "SELECT VrptOrders.* FROM VrptOrders WHERE OrderNo='" & strOrderNo & "'", con, adOpenForwardOnly, adLockReadOnly
        .Open "SELECT VrptOrders_ForProduction.* FROM VrptOrders_ForProduction WHERE OrderNo='" & strOrderNo & "'", con, adOpenForwardOnly, adLockReadOnly
        
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , !CompITemID)
            ITM.ListSubItems.add , , !ItemName & ""
            ITM.ListSubItems.add(, , Val(!Qty & "")).Bold = True
            lAlreadyAdjQty = GetSingleLongValue("Qty", "VStockOrderAdjustments", " WHERE OrderNo='" & strOrderNo & "' AND ItemID='" & !CompITemID & "'")
            ITM.ListSubItems.add(, , lAlreadyAdjQty).Bold = True
            
            ITM.ListSubItems.add(, , Val(!Qty & "") - lAlreadyAdjQty).Bold = True
            
            lInHand_Stock = GetSingleLongValue("SUM(NetQty)", "VItemsWithShelfWiseStock_Customerwise", " WHERE ItemID='" & !CompITemID & "' AND CustCode='Stock'")
            ITM.ListSubItems.add(, , Val(!InHand_CustomerWise & "") + lInHand_Stock).Bold = True
            ITM.ListSubItems.add(, , "0").Bold = True
            ITM.ListSubItems.add(, , "-").Tag = "0"
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    strCustCode = GetSingleStringValue("CustCode", "FCustomerOrders", " WHERE OrderNo='" & strOrderNo & "'")
    
    Me.Show
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub LVAdj_KeyDown(KeyCode As Integer, Shift As Integer)
    If LVAdj.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure?", vbQuestion) = vbYes Then
            LVAdj.ListItems.Remove LVAdj.SelectedItem.Index
        End If
    End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        If cmbLocation_QEL.MatchFound = False Then
            MsgBox "Please Select Location.", vbInformation
            Exit Sub
        End If
        
        If Val(txtEdit) > Val(LV.SelectedItem.SubItems(5)) Or Val(txtEdit) <= 0 Or Val(txtEdit) > Val(cmbLocation_QEL.List(cmbLocation_QEL.ListIndex, 2)) Then
            MsgBox "Invalid Qty."
            Exit Sub
        End If
        
        Dim ITM As ListItem
        Set ITM = LVAdj.ListItems.add(, , cmbLocation_QEL.Text)
        ITM.ListSubItems.add , , Val(txtEdit)
        ITM.ListSubItems.add(, , cmbLocation_QEL.List(cmbLocation_QEL.ListIndex, 3)).Tag = cmbLocation_QEL.List(cmbLocation_QEL.ListIndex, 1)
        
        Call GetAdjTotalQty
        
    ElseIf KeyAscii = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub

Private Function GetAdjTotalQty() As Long
    Dim i As Integer, lTotal As Long
    For i = 1 To LVAdj.ListItems.count
        lTotal = lTotal + Val(LVAdj.ListItems(i).SubItems(1))
    Next
    GetAdjTotalQty = lTotal
    txtTotalAdj = GetAdjTotalQty
End Function

Private Function Saved() As Boolean
    
    On Error GoTo err
    
    Dim strIssNo As String, lIISD_EntryID As Long, lEntryID As Long, lRefID As Long
    Dim strMsgLots As String
                   
    Call StartTrans(con)
    
    Dim i As Integer, j As Integer, lProcessID As Long
    Dim lRcvEntryID As Long, lNextProcessID As Long, lVID_EntryID As Long
    Dim strMaterialID As String, dQty As Double, dBal As Double, theDesc As String
    Dim lLotNo As String, lSOA_EntryID As Long
    Dim lOriginalProcessID As Long
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If Val(.SubItems(6)) > 0 Then
                'Get Process-ID ...
                lOriginalProcessID = Val(.ListSubItems(7).Tag)
                lProcessID = getPrevProcessID(.Text, lOriginalProcessID)
                
                If lRcvEntryID = 0 Then
                    con.Execute "INSERT INTO VendReceived(VendID,DT,RecieptID,UserID,ProcessID,Issuance_RefID) VALUES(" & _
                     iFactoryMakerID & ",'" & DT & "','',0," & IIf(lProcessID = 0, lOriginalProcessID, lProcessID) & ",0)"
                 
                    lRcvEntryID = GetSingleLongValue("MAX(EntryID)", "VendReceived")
                End If
                If lProcessID = 0 Then
                    lNextProcessID = lOriginalProcessID
                Else
                    lNextProcessID = getNextProcessID(.Text, lProcessID)
                End If
                
'                lLotNo = GetSingleLongValue("MAX(CAST(LotNo AS INT))", "VendRcvdDetail")
'                lLotNo = lLotNo + 1
                lLotNo = getNextLotNo
                
                strMsgLots = strMsgLots & " Lot #:" & lLotNo & " Item:" & .Text & vbNewLine
                
                con.Execute "INSERT INTO VendRcvdDetail(RefID,ItemCode,RecieptID,RcvdQty,Rate,LotNo,NextProcessID,OrderNo" & _
                 ",CountedBy,Issue_RefID,ProcessID,RcvdWeight,Opening_RefID) VALUES(" & lRcvEntryID & ",'" & .Text & _
                 "',''," & Val(.SubItems(6)) & ",0,'" & lLotNo & "'," & lNextProcessID & ",'" & strOrderNo & "','',0," & lProcessID & _
                 ",0,0)"
                
                lVID_EntryID = GetSingleLongValue("MAX(EntryID)", "VendRcvdDetail")
                
                con.Execute "INSERT INTO StockOrderAdjustments(OrderNo,ItemID,Qty,UserName,MachineName,VID_RefID,Shelf_RefID) VALUES('" & _
                     strOrderNo & "','" & .Text & "'," & Val(.SubItems(6)) & ",'" & CurrentUserName & "','" & strComputerName & _
                     "'," & lVID_EntryID & "," & Val(.ListSubItems(6).Tag) & ")"
                
                lSOA_EntryID = GetSingleLongValue("MAX(EntryID)", "StockOrderAdjustments")
                
                strMaterialID = .Text
                dQty = Val(.SubItems(6))
                dBal = getItemBalance(strMaterialID, DT)
                dBal = dBal - dQty
                theDesc = "Stock Adjusted against Order # " & strOrderNo
                
                con.Execute "INSERT INTO ItemLedgerSimple(ItemID,DT,Description,Qty,Balance,OrderNo,Location) VALUES('" & _
                 strMaterialID & "','" & DT & "','" & theDesc & "'," & -dQty & "," & dBal & ",'" & strOrderNo & "','" & .ListSubItems(6).Tag & "')"
             
                Call adjustItemLedger(strMaterialID, DT, -dQty)
            
                lRefID = Val(con.Execute("SELECT MAX(EntryID) FROM ItemLedgerSimple").Fields(0).Value & "")
                
                If strIssNo = "" Then
                    strIssNo = getIssuanceNo(DT)
                    con.Execute "INSERT INTO IssItemsSimple(IssNo,DT,OrderNo,Department,AttnPerson,UserName,ComputerName,SampleIssuance,CustCode) VALUES('" & _
                     strIssNo & "','" & DT & " " & ServerTime & "','" & strOrderNo & "','','','" & CurrentUserName & _
                     "','" & strComputerName & "',0,'" & strCustCode & "')"
                End If
                
                con.Execute "INSERT INTO IssItemsSimpleDetail(IssNo,ItemID,Qty,InHand,LedgerRefID,OrderNo,Location,Remarks) Values('" & _
                 strIssNo & "','" & strMaterialID & "'," & dQty & "," & getItemBalance(strMaterialID, DT) - dQty & "," & lRefID & _
                 ",'" & strOrderNo & "','" & .ListSubItems(5).Tag & "','')"
                
                lIISD_EntryID = GetSingleLongValue("MAX(EntryID)", "IssItemsSimpleDetail")
                    
                For j = 1 To LVAdjAll.ListItems.count
                    With LVAdjAll.ListItems(j)
                    
                        If Val(.Tag) = Val(LV.ListItems(i).Index) Then
                            con.Execute "INSERT INTO StockOrderAdjustmentsDetail(SOA_RefID,Shelf_RefID,Qty) VALUES(" & lSOA_EntryID & _
                             "," & Val(.ListSubItems(2).Tag) & "," & Val(.SubItems(1)) & ")"
                             
                            con.Execute "INSERT INTO IssItemsSimpleDetail_More(IISD_RefID,IssdQty,Shelf_RefID,Remarks) VALUES(" & _
                             lIISD_EntryID & "," & Val(.SubItems(1)) & "," & Val(.ListSubItems(2).Tag) & ",'')"
                                    
                            Call UpdateItemLocation(strMaterialID, Val(.ListSubItems(2).Tag), Val(.SubItems(1)), , .SubItems(3))
                        End If
                                
                    End With
                Next
                
            End If
        End With
    Next
    
    con.CommitTrans
    Saved = True
    
    MsgBox "Following Lots Generated : " & vbNewLine & strMsgLots, vbInformation
    
    Exit Function
err:
    MsgBox err.Description
End Function
