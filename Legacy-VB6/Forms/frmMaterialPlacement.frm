VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmMaterialPlacement 
   Caption         =   "Material Placement"
   ClientHeight    =   9135
   ClientLeft      =   1170
   ClientTop       =   450
   ClientWidth     =   13800
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
   ScaleHeight     =   9135
   ScaleWidth      =   13800
   WindowState     =   2  'Maximized
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
      Left            =   75
      TabIndex        =   15
      Top             =   0
      Width           =   13695
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Material Placement"
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
         Left            =   5325
         TabIndex        =   16
         Top             =   135
         Width           =   3060
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Material Placement"
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
         Left            =   5340
         TabIndex        =   17
         Top             =   150
         Width           =   3060
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8550
      Left            =   75
      TabIndex        =   0
      Top             =   555
      Width           =   13695
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
         Left            =   10845
         ScaleHeight     =   2835
         ScaleWidth      =   2760
         TabIndex        =   20
         Top             =   690
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
         Left            =   45
         TabIndex        =   19
         Top             =   8145
         Width           =   1560
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   435
         Left            =   9675
         TabIndex        =   13
         Top             =   7965
         Width           =   1275
      End
      Begin VB.PictureBox PicEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4545
         ScaleHeight     =   285
         ScaleWidth      =   6375
         TabIndex        =   8
         Top             =   3015
         Visible         =   0   'False
         Width           =   6405
         Begin VB.TextBox txtLotNo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5385
            TabIndex        =   18
            Top             =   -15
            Width           =   1000
         End
         Begin VB.TextBox txtBatchNo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4395
            TabIndex        =   14
            Top             =   -15
            Width           =   1000
         End
         Begin VB.TextBox txtQty 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   -15
            TabIndex        =   9
            Top             =   -15
            Width           =   1000
         End
         Begin MSForms.ComboBox cmbShelf 
            Height          =   315
            Left            =   3405
            TabIndex        =   12
            Top             =   -15
            Width           =   1005
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "1764;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbRack 
            Height          =   315
            Left            =   2415
            TabIndex        =   11
            Top             =   -15
            Width           =   1000
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "1764;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbStore 
            Height          =   315
            Left            =   975
            TabIndex        =   10
            Top             =   -15
            Width           =   1455
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "2566;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   435
         Left            =   11025
         TabIndex        =   7
         Top             =   7965
         Width           =   1275
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   435
         Left            =   12360
         TabIndex        =   6
         Top             =   7965
         Width           =   1275
      End
      Begin ComboList.Usercontrol1 cmbVend 
         Height          =   285
         Left            =   6765
         TabIndex        =   3
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
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
         Left            =   60
         TabIndex        =   4
         Top             =   360
         Width           =   6690
         _ExtentX        =   11800
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
         Height          =   7245
         Left            =   45
         TabIndex        =   1
         Top             =   675
         Width           =   13590
         _ExtentX        =   23971
         _ExtentY        =   12779
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
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Vendor"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Rcv ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Material ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Material Name"
            Object.Width           =   3836
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Qty Rcvd"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Placed"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Balance"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Store"
            Object.Width           =   2566
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Rack"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "Shelf"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   11
            Text            =   "Batch #"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   12
            Text            =   "Lot #"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Material"
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
         Index           =   0
         Left            =   60
         TabIndex        =   5
         Top             =   150
         Width           =   705
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor"
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
         Index           =   2
         Left            =   6765
         TabIndex        =   2
         Top             =   150
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmMaterialPlacement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bDontStartRefreshing As Boolean
Dim iColNo As Integer
Dim bSample As Boolean

Private Sub RefreshList()
    
    If cmbVend.MatchFound = False Then Exit Sub
    If cmbMaterial.MatchFound = False Then Exit Sub
    
    If bDontStartRefreshing Then Exit Sub

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim strSQL As String
    Dim i As Integer
    Dim strCond As String
    
    strCond = " WHERE ROUND(QtyPassed,2)>ROUND(QtyPlaced,2)"
    If cmbMaterial.ID <> "0" Then
        strCond = strCond & " AND RMID1='" & cmbMaterial.ID & "'"
    End If
    If cmbVend.ID <> "0" Then
        strCond = strCond & " AND VendID='" & cmbVend.ID & "'"
    End If
    
    If bSample Then
        strCond = strCond & " AND Sample=" & Abs(bSample)
    End If
    
    With rs
        strSQL = "SELECT * FROM VMaterialsForPlacement " & strCond & " Order By RcvDate,EntryID"
        .Open strSQL, con, adOpenForwardOnly, adLockReadOnly
        
        LV.ListItems.Clear
        
        Do Until .EOF
            
            Set ITM = LV.ListItems.add(, , !AccTitle & "")
            ITM.Tag = !EntryID & ""
            ITM.ListSubItems.add , , ![RcvID] & ""
             
            ITM.ListSubItems.add , , !MaterialID & ""
            ITM.ListSubItems.add(, , !RMName & "").Tag = Abs(!CCItem)
            ITM.ListSubItems.add(, , Round(Val(![QtyPassed] & ""), 2) & " " & ![Unit]).Tag = Val(!RISD_EntryID & "")
            ITM.ListSubItems.add , , Round(Val(![QtyPlaced] & ""), 2)
            ITM.ListSubItems.add , , Round(Val(![QtyPassed] & "") - Val(![QtyPlaced] & ""), 2)
            
            ITM.ListSubItems.add , , ""
            ITM.ListSubItems.add , , ""
            ITM.ListSubItems.add , , ""
            ITM.ListSubItems.add , , ""
            ITM.ListSubItems.add , , ""
            ITM.ListSubItems.add , , ""
            ITM.ListSubItems.add , , ""
            'ITM.ListSubItems.add , , Round(Val(![Sheets] & "") - Val(![SheetsPlaced] & ""), 2)
            For i = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(i).ToolTipText = ITM.ListSubItems(i).Text
            Next
            
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub

err:
   MsgBox err.Description
   
End Sub

Private Sub cmbMaterial_matched()
    Call RefreshList
End Sub

Private Sub cmbRack_Change()
    Call cmbRack_Click
End Sub

Private Sub cmbRack_Click()

    On Error GoTo err
    
    If cmbRack.MatchFound = False Then Exit Sub
    Call AddToCombo(cmbShelf, "ShelfNo", "StoreShelfs", " WHERE Rack_RefID=" & cmbRack.List(cmbRack.ListIndex, 1), , "EntryID")
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbRack_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode.Value = 13 Then
        cmbShelf.SetFocus
    ElseIf KeyCode.Value = 27 Then
        PicEdit.Visible = False
    End If
End Sub

Private Sub cmbShelf_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    If KeyCode.Value = 13 Then
        txtBatchNo.SetFocus
    ElseIf KeyCode.Value = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub

Private Sub cmbStore_Change()
    Call cmbStore_Click
End Sub

Private Sub cmbStore_Click()

    On Error GoTo err
    
    If cmbStore.MatchFound = False Then Exit Sub
    Call AddToCombo(cmbRack, "RackNo", "StoreRacks", " WHERE Store_RefID=" & cmbStore.List(cmbStore.ListIndex, 1), , "EntryID")
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbStore_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    If KeyCode.Value = 13 Then
        cmbRack.SetFocus
    ElseIf KeyCode.Value = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub

Private Sub cmbVend_matched()
    Call RefreshList
End Sub

Private Sub cmbVend_UnMatched()
    LV.ListItems.Clear
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Dim strCondition As String
    Set rpt = rptApp.OpenReport(RptPath & "\MaterialPlacementPendencies.rpt")
    strCondition = "ROUND({VMaterialsForPlacement.QtyPassed},2)>ROUND({VMaterialsForPlacement.QtyPlaced},2)"
    If cmbMaterial.ID <> "0" Then
        strCondition = strCondition & " AND {VMaterialsForPlacement.MaterialID}='" & cmbMaterial.ID & "'"
    End If
    If cmbVend.ID <> "0" Then
        strCondition = strCondition & " AND {VMaterialsForPlacement.VendID}=" & cmbVend.ID
    End If
    
    Dim f As New frmPrevRpt
    f.ShowReport strCondition, rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdSave_Click()
    
    If Saved Then
        Call RefreshList
    End If
    
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
   'If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    
    bDontStartRefreshing = True
    cmbVend.AddItem "<All Venders>", "0"
    cmbVend.AddVals con, "AccTitle", "VVenders", "AccNo"
    cmbVend.ID = "0"
     
    Call AddToCombo(cmbStore, "StoreName", "Stores", " WHERE EntryID IN(SELECT Store_RefID FROM Users_Stores WHERE UserID=" & UserID & ") ORDER BY StoreName", , "EntryID")
    
    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        LV.ColumnHeaders(i).Tag = LV.ColumnHeaders(i).Width
    Next
    bDontStartRefreshing = False
    
    
End Sub

Private Sub LV_DblClick()

    On Error GoTo err
    If LV.ListItems.count = 0 Then Exit Sub
    
'    If Val(LV.SelectedItem.SubItems(11)) > 0 Then
'        txtSheets.Enabled = True
'    Else
'        txtSheets.Enabled = False
'    End If
    
    PicEdit.Move LV.Left + LV.ColumnHeaders(8).Left, LV.Top + LV.SelectedItem.Top
    PicEdit.Visible = True
    txtQty.Text = ""
    txtQty.SetFocus
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)

    PicEdit.Visible = False
    
    If chkShowPicture.Value = vbChecked Then
        Dim rs As New ADODB.Recordset
        With rs
            .Open "SELECT Pic FROM RM WHERE RMID1='" & (LV.SelectedItem.SubItems(2)) & "'", con, adOpenForwardOnly, adLockReadOnly
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

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    iColNo = 0
    'Dim i As Integer
    
End Sub

Private Function Saved() As Boolean

    Saved = False
    On Error GoTo err
    
    Call StartTrans(con)
    
    Dim i As Integer
    
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
        
            If Val(.SubItems(7)) > 0 Then
                
                con.Execute "INSERT INTO MaterialLocationwiseStatus(Rcvd_RefID,Shelf_RefID,QtyPlaced,UserName,MachineName,BatchNo,LotNo) VALUES(" & _
                 Val(.Tag) & "," & .ListSubItems(10).Tag & "," & Val(.SubItems(7)) & ",'" & CurrentUserName & "','" & strComputerName & "','" & (.SubItems(11)) & "','" & .SubItems(12) & "')"
                 
                If Val(.ListSubItems(3).Tag) = 1 Then
                    con.Execute "INSERT INTO RcvItemsSimpleDetail_Placement(RISD_RefID,Shelf_RefID,RcvdQty,Remarks) VALUES(" & _
                     Val(.ListSubItems(4).Tag) & "," & .ListSubItems(10).Tag & "," & Val(.SubItems(7)) & ",'')"
                End If
                
            End If
            
        End With
    Next
        
    con.CommitTrans
    Saved = True
     
    Screen.MousePointer = vbDefault
    
    Exit Function
    
err:
    MsgBox err.Description
    con.RollbackTrans
    Screen.MousePointer = vbDefault
    
End Function

Private Sub txtBatchNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtLotNo.SetFocus
    ElseIf KeyAscii = 27 Then
        PicEdit.Visible = False
    End If
End Sub

Private Sub txtLotNo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        Dim ITM As ListItem
        Set ITM = LV.SelectedItem
        
        If Val(txtQty) > Val(ITM.SubItems(6)) Or Val(txtQty) = 0 Then
            MsgBox "Invalid Qty.", vbInformation
            Exit Sub
        ElseIf cmbStore.MatchFound = False Then
            MsgBox "Invalid Store.", vbInformation
            Exit Sub
        ElseIf cmbRack.MatchFound = False Then
            MsgBox "Invalid Rack.", vbInformation
            Exit Sub
        ElseIf cmbShelf.MatchFound = False Then
            MsgBox "Invalid Shelf.", vbInformation
            Exit Sub
        End If
        
        ITM.SubItems(7) = Val(txtQty)
        ITM.SubItems(8) = cmbStore.Text
        ITM.SubItems(9) = cmbRack.Text
        ITM.SubItems(10) = cmbShelf.Text
        ITM.ListSubItems(10).Tag = cmbShelf.List(cmbShelf.ListIndex, 1)
        ITM.SubItems(11) = txtBatchNo.Text
        ITM.SubItems(12) = txtLotNo.Text
        LV.SetFocus
        PicEdit.Visible = False
        
    ElseIf KeyAscii = 27 Then
        PicEdit.Visible = False
    End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmbStore.SetFocus
    ElseIf KeyAscii = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub

Private Sub txtSheets_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        Dim ITM As ListItem
        Set ITM = LV.SelectedItem
        If Val(txtQty) > Val(ITM.SubItems(6)) Or Val(txtQty) = 0 Then
            MsgBox "Invalid Qty.", vbInformation
            Exit Sub
        ElseIf cmbStore.MatchFound = False Then
            MsgBox "Invalid Store.", vbInformation
            Exit Sub
        ElseIf cmbRack.MatchFound = False Then
            MsgBox "Invalid Rack.", vbInformation
            Exit Sub
        ElseIf cmbShelf.MatchFound = False Then
            MsgBox "Invalid Shelf.", vbInformation
            Exit Sub
'        ElseIf Val(txtSheets) > Val(LV.SelectedItem.SubItems(11)) Or Val(txtSheets) <= 0 Then
'            MsgBox "Invalid Sheets", vbInformation
'            Exit Sub
        End If
        
        ITM.SubItems(7) = Val(txtQty)
        ITM.SubItems(8) = cmbStore.Text
        ITM.SubItems(9) = cmbRack.Text
        ITM.SubItems(10) = cmbShelf.Text
        'ITM.SubItems(11) = Val(txtSheets)
        ITM.ListSubItems(10).Tag = cmbShelf.List(cmbShelf.ListIndex, 1)
        LV.SetFocus
        
        PicEdit.Visible = False
        
    ElseIf KeyAscii = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub

Public Sub ShowMe(Optional p_bSample As Boolean = False)

    bSample = p_bSample
        
        
    cmbMaterial.AddItem "<All Materials>", "0"
    If bSample Then
        cmbMaterial.AddVals con, "ItemName + '     {' + ItemID + '}'", "Items", "ItemID", " WHERE InActive=0 AND Sample=1 ORDER BY ItemName"
    Else
        cmbMaterial.AddVals con, "'{' + RMID1 + '} ' + RMName", "RM", "RMID1", " ORDER BY RMID1"
    End If
    cmbMaterial.ID = "0"
    
    Call RefreshList
    
    Me.Show
    
End Sub


Private Sub chkShowPicture_Click()

    If chkShowPicture.Value = vbChecked Then
        Call LV_ItemClick(LV.SelectedItem)
    Else
        PicProduct.Visible = False
    End If
    
End Sub

