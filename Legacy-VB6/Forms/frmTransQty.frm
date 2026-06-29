VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTransQty 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transfer Quantities"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   11865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   30
      Width           =   9015
      Begin VB.Label lblItemCode 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
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
         Height          =   285
         Left            =   3810
         TabIndex        =   11
         Top             =   300
         Width           =   1665
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Item Code :"
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
         Height          =   195
         Left            =   2760
         TabIndex        =   10
         Top             =   330
         Width           =   975
      End
      Begin VB.Label lblOrdNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
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
         Height          =   285
         Left            =   1020
         TabIndex        =   9
         Top             =   300
         Width           =   1665
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Order No :"
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
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   330
         Width           =   825
      End
      Begin VB.Label lblQty 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
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
         Height          =   285
         Left            =   7050
         TabIndex        =   7
         Top             =   300
         Width           =   1845
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Qty To Transfer :"
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
         Height          =   195
         Left            =   5580
         TabIndex        =   6
         Top             =   330
         Width           =   1410
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6480
      Left            =   0
      TabIndex        =   0
      Top             =   930
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   11430
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Free Stock"
      TabPicture(0)   =   "frmTransQty.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LVInHouse"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtRate"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Transferred"
      TabPicture(1)   =   "frmTransQty.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LVTrans"
      Tab(1).Control(1)=   "txtUpdate1"
      Tab(1).ControlCount=   2
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4950
         TabIndex        =   2
         Top             =   645
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.TextBox txtUpdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71670
         TabIndex        =   1
         Top             =   1650
         Visible         =   0   'False
         Width           =   1350
      End
      Begin MSComctlLib.ListView LVInHouse 
         Height          =   6015
         Left            =   90
         TabIndex        =   3
         Top             =   390
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   10610
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S No."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Production Phase"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Available Qty"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Qty To Transfer"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView LVTrans 
         Height          =   6015
         Left            =   -74910
         TabIndex        =   4
         Top             =   375
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   10610
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
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Production Phase"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item Code"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Description"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Transffered Qty"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin MSForms.CommandButton cmdSaveClose 
      Height          =   360
      Left            =   8430
      TabIndex        =   13
      Top             =   7500
      Width           =   1665
      Caption         =   "Save"
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
   Begin MSForms.CommandButton cmdCancel 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   10155
      TabIndex        =   12
      Top             =   7500
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
End
Attribute VB_Name = "frmTransQty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OrderNo As String
Dim ItemCode As String
Dim ItemName As String
Dim Qty As Long
Public Sub PrepareToShow(ArgOrderNo As String, ArgItemCode As String, ArgItemName As String, ArgQty As Long)

    On Error GoTo err
    
    OrderNo = ArgOrderNo
    ItemCode = ArgItemCode
    ItemName = ArgItemName
    Qty = ArgQty
    
    lblOrdNo = OrderNo
    lblItemCode = ItemCode
    lblQty = Qty
    
    Call FillFreeStock
    
    Me.Show 1
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub FillFreeStock()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim itm As ListItem
    Dim ITM1 As ListItem
    Dim lFreeStock As Long
    
    With rs
        .Open "Select * From VFFreeStockNew Where ItemCode='" & ItemCode & "'", con, adOpenForwardOnly, adLockReadOnly
        LVInHouse.ListItems.Clear
        
        Do Until .EOF
        
            Set itm = LVInHouse.ListItems.Add(, ![ProcessID] & "'", LVInHouse.ListItems.Count + 1)
            itm.Tag = Val(![InStock] & "")
            'If IsNull(![NextProcessID]) Then
            '    itm.ListSubItems.Add , , "Finished Pcs Available."
            'Else
            itm.ListSubItems.Add , , ![ProcDescription] & ""
            itm.ListSubItems(1).Tag = Val(![OutStock] & "")
            'End If
            
            lFreeStock = Val(![FreeStock] & "")
            
            For i = 1 To LVTrans.ListItems.Count
                Set ITM1 = LVTrans.ListItems(i)
                'If ITM1.SubItems(2) = LV1.SelectedItem.key And Val(ITM1.ListSubItems(1).key) = ![ProcessID] Then
                If ITM1.SubItems(2) = Left(LV1.SelectedItem.key, Len(LV1.SelectedItem.key) - 1) And Val(ITM1.ListSubItems(1).key) = ![ProcessID] Then
                    lFreeStock = lFreeStock - Val(LVTrans.ListItems(i).SubItems(4) & "")
                End If
            Next i
            
            itm.ListSubItems.Add , , lFreeStock
            itm.ListSubItems.Add , , "0"
            
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSaveClose_Click()
    If Saved Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    OrderNo = ""
    ItemCode = ""
    ItemName = ""
    Qty = 0
End Sub

Private Sub txtRate_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo err
    
    Select Case KeyCode
    
        Case 13
            
            
            If Val(txtRate) > Qty Then
                MsgBox "Invalid Qty.", vbInformation
                KeyCode = 0
                txtRate.SelStart = 0
                txtRate.SelLength = Len(txtQty)
                txtRate.SetFocus
                Exit Sub
            End If
            
            Dim itm As ListItem, SelItm As ListItem, SelITM1 As ListItem
            Set SelItm = LVInHouse.SelectedItem
            
            Set itm = LVTrans.ListItems.Add(, OrderNo & "|" & ItemCode, LVTrans.ListItems.Count + 1)
            itm.Tag = Val(LVInHouse.SelectedItem.key)  'ProcessID
            itm.ListSubItems.Add , SelItm.key, SelItm.SubItems(1)
            itm.ListSubItems.Add , , ItemCode
            itm.ListSubItems.Add , , ItemName
            itm.ListSubItems.Add , , Val(txtRate)
            
            itm.ListSubItems(1).Tag = SelItm.Tag
            itm.ListSubItems(2).Tag = SelItm.ListSubItems(1).Tag
            
            Qty = Qty - Val(txtRate)
            lblQty = Qty
            SelItm.SubItems(2) = SelItm.SubItems(2) - Val(txtRate)
            txtRate.Visible = False
            
        Case 27
            txtRate.Visible = False
            
    End Select
    
    Exit Sub
    
err:
    MsgBox err.Description
End Sub

Private Sub LVInHouse_DblClick()

    If LVInHouse.ListItems.Count = 0 Then Exit Sub
    'If cmbCust.MatchFound = False Then Exit Sub

    
    With txtRate
        .Move LVInHouse.Left + LVInHouse.ColumnHeaders(4).Left, LVInHouse.Top + LVInHouse.SelectedItem.Top, LVInHouse.ColumnHeaders(4).Width
        '.Text = MyCol(LV.SelectedItem.SubItems(2) & "|" & LV.SelectedItem.SubItems(3)).IssueableQty 'LVMaker.SelectedItem.Tag 'LVMaker.SelectedItem.SubItems(3)
        '.Text = MyCol(LV1.SelectedItem.Key).IssueableQty 'LVMaker.SelectedItem.Tag 'LVMaker.SelectedItem.SubItems(3)
        .Text = LVInHouse.SelectedItem.SubItems(2)
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With

End Sub


Private Sub LVInHouse_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call LVInHouse_DblClick
   End If
End Sub
Private Sub txtUpdate1_KeyPress(KeyAscii As Integer)

    On Error GoTo err
    
    
    If KeyAscii = 13 Then
        
        If Val(txtUpdate1) = 0 Then
            MsgBox "Invalid Qty!", vbInformation
            Exit Sub
        ElseIf Val(txtUpdate1) > Val(txtUpdate1.Tag) + Val(lblQty) Then
            MsgBox "Invalid Qty!", vbInformation
            Exit Sub
        End If
        
        Qty = Qty + Val(txtUpdate1.Tag) - Val(txtUpdate1)
        lblQty = Qty
        
        LVTrans.SelectedItem.SubItems(4) = Val(txtUpdate1)
        
        
        'Update The Qnty In LvInHouse
        
        
        LVInHouse.ListItems(LVTrans.SelectedItem.Tag & "'").SubItems(2) = LVInHouse.ListItems(LVTrans.SelectedItem.Tag & "'").SubItems(2) + Val(txtUpdate1.Tag) - Val(txtUpdate1)
        
        
        txtUpdate1.Visible = False
        
    ElseIf KeyAscii = 27 Then
        txtUpdate1.Visible = False
    End If
    
    Exit Sub
    
err:
    MsgBox err.Description
End Sub
Private Sub LVTrans_DblClick()

    On Error GoTo err
    
    If LVTrans.ListItems.Count = 0 Then Exit Sub
    With txtUpdate1
        .Text = Val(LVTrans.SelectedItem.SubItems(4))
        .Tag = Val(LVTrans.SelectedItem.SubItems(4))
        .Move LVTrans.Left + LVTrans.ColumnHeaders(5).Left + 10, LVTrans.SelectedItem.Top + LVTrans.Top, LVTrans.ColumnHeaders(5).Width
        
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub LVTrans_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo err
    
    If LVTrans.ListItems.Count = 0 Then Exit Sub
    
    If KeyCode = 46 Then
        If MsgBox("Are You Sure to Delete This Entry!.", vbQuestion + vbYesNo) = vbYes Then
            Qty = Qty + Val(LVTrans.SelectedItem.SubItems(4))
            lblQty = Qty
            LVInHouse.ListItems(LVTrans.SelectedItem.Tag & "'").SubItems(2) = LVInHouse.ListItems(LVTrans.SelectedItem.Tag & "'").SubItems(2) + Val(LVTrans.SelectedItem.SubItems(4))
            LVTrans.ListItems.Remove LVTrans.SelectedItem.Index
        End If
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Function Saved() As Boolean

    On Error GoTo err
    'Now Save Transferred Items...
    Dim cmd As New ADODB.Command
    Dim lTotInStock As Long
    Dim lTotOutStock As Long
    Dim lInStock As Long
    Dim lOutStock As Long
    Dim lQty As Long
    
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    For i = 1 To LVTrans.ListItems.Count
        
        lQty = Val(LVTrans.ListItems(i).SubItems(4) & "")
        lTotInStock = Val(LVTrans.ListItems(i).ListSubItems(1).Tag & "")
        lTotOutStock = Val(LVTrans.ListItems(i).ListSubItems(2).Tag & "")
        
        
        If lQty <= lTotInStock Then
            lInStock = lQty
            lOutStock = 0
        Else
            lInStock = lTotInStock
            lOutStock = lQty - lTotInStock
        End If
            
        With cmd
            If lInStock > 0 Then
        
                .CommandText = "SP_InsertTransferred"
                .Parameters("@OrderNo").value = lblOrdNo
                .Parameters("@ItemCode").value = lblItemCode 'LVTrans.ListItems(i).SubItems(2) & ""
                .Parameters("@ProcessID").value = Val(LVTrans.ListItems(i).ListSubItems(1).key)
                .Parameters("@Qty").value = lInStock 'Val(LVTrans.ListItems(i).SubItems(4) & "")
                .Parameters("@DT").value = Format(ServerDate, "dd-MMM-yyyy")
                .Execute
                
            End If
            If lOutStock > 0 Then
                'Do For Out Stock Tooooooooo.
                .CommandText = "SP_InsertTransferredIssd"
                .Parameters("@OrderNo").value = lblOrdNo
                .Parameters("@ItemCode").value = lblItemCode
                .Parameters("@ProcessID").value = Val(LVTrans.ListItems(i).ListSubItems(1).key)
                
                .Parameters("@Qty").value = lOutStock 'Val(LVTrans.ListItems(i).SubItems(4) & "")
                .Parameters("@DT").value = Format(ServerDate, "dd-MMM-yyyy")
                .Execute
            End If
        End With
    Next i
    Saved = True
    Exit Function
err:
    MsgBox err.Description
    
End Function
