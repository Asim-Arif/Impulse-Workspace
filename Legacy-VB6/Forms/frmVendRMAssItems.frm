VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmVendRMAssItems 
   ClientHeight    =   6570
   ClientLeft      =   1110
   ClientTop       =   60
   ClientWidth     =   11880
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
   ScaleHeight     =   6570
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2430
      TabIndex        =   19
      Top             =   2970
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.PictureBox PicEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4260
      ScaleHeight     =   285
      ScaleWidth      =   3990
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   3990
      Begin VB.TextBox txtSnaff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2985
         TabIndex        =   18
         Top             =   0
         Width           =   1000
      End
      Begin VB.TextBox txtStamp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1950
         TabIndex        =   17
         Top             =   0
         Width           =   1050
      End
      Begin VB.TextBox txtPlant 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   975
         TabIndex        =   16
         Top             =   0
         Width           =   1000
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   1000
      End
   End
   Begin ComboList.Usercontrol1 cmbCust 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   855
      Width           =   7320
      _ExtentX        =   12912
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
   Begin ComboList.Usercontrol1 cmbItem 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   1380
      Width           =   7320
      _ExtentX        =   12912
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
   Begin VB.Frame fupdate 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   390
      Left            =   120
      TabIndex        =   12
      Top             =   6045
      Visible         =   0   'False
      Width           =   3090
      Begin MSForms.CommandButton cmdCheckAll 
         Height          =   375
         Left            =   15
         TabIndex        =   3
         Top             =   0
         Width           =   1500
         ForeColor       =   0
         Caption         =   " Assign All"
         PicturePosition =   327683
         Size            =   "2646;661"
         Accelerator     =   104
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdClearAll 
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   0
         Width           =   1500
         ForeColor       =   0
         Caption         =   "Unassign All"
         PicturePosition =   327683
         Size            =   "2646;661"
         Accelerator     =   108
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin VB.CommandButton cmdUnAssign 
      Height          =   240
      Left            =   9585
      Picture         =   "frmVendRMAssItems.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "UnAssign Selected Item..."
      Top             =   1890
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   435
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
      TabIndex        =   7
      Top             =   0
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assigned Items"
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
         Left            =   3840
         TabIndex        =   8
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assigned Items"
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
         Left            =   3855
         TabIndex        =   9
         Top             =   135
         Width           =   2295
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4140
      Left            =   150
      TabIndex        =   6
      Top             =   1890
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   7303
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "S No."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "RM ID"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Group"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Remarks"
         Object.Width           =   4233
      EndProperty
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   1
      Left            =   720
      TabIndex        =   13
      Top             =   1155
      Width           =   7320
      BackColor       =   11517387
      Caption         =   " Raw Material"
      Size            =   "12912;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdAssign 
      Height          =   375
      Left            =   8055
      TabIndex        =   2
      Top             =   1275
      Width           =   1125
      ForeColor       =   6244673
      Caption         =   "Assign   "
      PicturePosition =   327683
      Size            =   "1984;661"
      Picture         =   "frmVendRMAssItems.frx":0102
      Accelerator     =   65
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   0
      Left            =   720
      TabIndex        =   10
      Top             =   630
      Width           =   7320
      BackColor       =   11517387
      Caption         =   " Vender"
      Size            =   "12912;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   375
      Left            =   8370
      TabIndex        =   5
      Top             =   6090
      Width           =   1665
      ForeColor       =   0
      Caption         =   "Close          "
      PicturePosition =   327683
      Size            =   "2937;661"
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
Attribute VB_Name = "frmVendRMAssItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iColNo As Integer

Private Sub cmbCust_matched()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    cmbItem.ClearVals
    
    Call cmbItem.AddVals(con, "RMName + ' {'+RMID1+'}'", "RM", "RMID", " Where RMID Not In(Select RMID From VenderAssItems WHERE VendID='" & cmbCust.ID & "')  ORDER BY RMName")
    
    With rs
        .Open "SELECT EntryID,RMID,RMID1,RMName,Rate,GroupName,Remarks FROM VVenderAssItems WHERE VendID='" & cmbCust.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, ![EntryID] & "'", LV.ListItems.count + 1)
            ITM.ListSubItems.add , , ![RMID1] & ""
            ITM.ListSubItems.add , , ![RMName] & ""
            ITM.ListSubItems.add , , ![GroupName] & ""
            ITM.ListSubItems.add , , Val(![Rate] & "")
            ITM.ListSubItems.add , , !Remarks & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmbCust_UnMatched()
    LV.ListItems.Clear
    'cmbCountry.ClearVals
End Sub

Private Sub cmbItemName_matched()
    cmbItem.ID = cmbItemName.ID
End Sub

Private Sub cmbProcess_matched()

'    cmbCust.ClearVals
'    If cmbProcess.ID = 7 Then
'        Call cmbCust.AddVals(con, "VenderName", "Venders", "VendID", " Where PhaseID=" & cmbProcess.ID)
'    Else
    cmbCust.ClearVals
    Call cmbCust.AddVals(con, "'{ ' + VendID1 + ' } ' + MakerName", "VMakers", "VendID", " Where PhaseID=" & cmbProcess.ID)
'    End If
    
End Sub

Private Sub cmdAssign_Click()

    On Error GoTo err
    Dim t1 As Long
    t1 = Timer
    If cmbCust.MatchFound = False Then Exit Sub
    
    If cmbItem.MatchFound = False Then Exit Sub
    
    Dim iSaved As Integer
    con.Execute "INSERT INTO VenderAssItems(VendID,RMID) VALUES('" & cmbCust.ID & "'," & cmbItem.ID & ")", iSaved
    
    If iSaved = 1 Then
         Call cmbCust_matched
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub



Private Sub cmdCheckAll_Click()
        
    
     
    If cmbCust.MatchFound = False Then Exit Sub
    
    'If cmbItem.MatchFound = False Then Exit Sub
    
    Dim Saved As Integer
    Dim cmd As New ADODB.Command
    
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_VendAssItem"
        .Parameters("@VendID").Value = cmbCust.ID 'cmbCust.ID
        .Parameters("@ItemID").Value = "0" 'cmbCountry.ID
        .Parameters("@ProcessID").Value = cmbCust.Tag
        .Execute Saved
    End With
    
    
    If Saved > 0 Then
        'Call RefreshLV
        'Call CmbCust's Matched Event To Remove Current Item
        'From The cmbItem...
        Call cmbCust_matched
        'cmdAssign.SetFocus
    End If

End Sub

Private Sub cmdClearAll_Click()

    On Error GoTo err
    Dim iAffect As Integer
    If MsgBox("Are You Sure To UnAssign All Items?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        con.Execute "Delete From VendAssItems Where VendID=" & cmbCust.ID, iAffect
        If iAffect > 0 Then
            Call cmbCust_matched
        End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdUnAssign_Click()

    On Error GoTo err
    If LV.ListItems.count = 0 Then Exit Sub
    Dim rs As New ADODB.Recordset
    With rs
        .Open "select Itemcode from Vvendissddetail where vendID=" & cmbCust.ID & " and Issqty>rcvdqty", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            MsgBox "Can't Unassign this item. It's Receivable.", vbInformation
            .Close
            Exit Sub
        End If
        .Close
    End With
    Set rs = Nothing
    
    Dim iAffect As Integer
    
    con.Execute "DELETE FROM VenderAssItems WHERE EntryID=" & Val(LV.SelectedItem.key), iAffect
    If iAffect > 0 Then
        Call cmbCust_matched
        cmdUnAssign.Visible = False
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
    End If
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

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
    If cmbCust.MatchFound = False Then Exit Sub

    If Not LV.SelectedItem.Checked Then
        LV.SelectedItem.Checked = True
        'Call LV_ItemCheck(LV.SelectedItem)
        Exit Sub
    End If
    
    If iColNo = 6 Then
        With txtRemarks
            .Move LV.Left + LV.ColumnHeaders(6).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(6).Width
            .Text = LV.SelectedItem.SubItems(5)
            .SelStart = 0
            .SelLength = Len(.Text)
            .Visible = True
            .SetFocus
        End With
        Exit Sub
    End If
    
    With txtRate
    
        '.Move LV.Left + LV.ColumnHeaders(5).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(5).Width
        PicEdit.Move LV.Left + LV.ColumnHeaders(5).Left, LV.Top + LV.SelectedItem.Top ', LV.ColumnHeaders(5).Width
        If LV.ColumnHeaders.count = 9 Then
            Dim rs As New ADODB.Recordset
            With rs
                .Open "SELECT IsNull(Planter,0),IsNull(Snaffer,0),IsNull(Stamper,0) From Venders Where VendID=" & cmbCust.ID, con, adOpenForwardOnly, adLockReadOnly
                
                'If .Fields(0) = 0 Then
                '    txtPlant.Enabled = False
                'Else
                '    txtPlant.Enabled = True
                'End If
                
                'If .Fields(1) = 0 Then
                '    txtSnaff.Enabled = False
                'Else
                '    txtSnaff.Enabled = True
                'End If
                
                'If .Fields(2) = 0 Then
                '    txtStamp.Enabled = False
                'Else
                '    txtStamp.Enabled = True
                'End If
                txtPlant.Enabled = False
                txtSnaff.Enabled = False
                txtStamp.Enabled = False
                .Close
                
            End With
            
            PicEdit.Width = 3990
            txtPlant.Visible = True
            txtSnaff.Visible = True
            txtStamp.Visible = True
            txtPlant.Text = LV.SelectedItem.SubItems(5)
            txtStamp.Text = LV.SelectedItem.SubItems(6)
            txtSnaff.Text = LV.SelectedItem.SubItems(7)
        Else
            PicEdit.Width = .Width
            txtPlant.Visible = False
            txtSnaff.Visible = False
            txtStamp.Visible = False
        End If
        PicEdit.Visible = True
        .Text = LV.SelectedItem.SubItems(4)
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With

    

End Sub





Private Sub Form_Load()

    cmbCust.ListHeight = 6000
    cmbItem.ListHeight = 6000
         
    cmbCust.AddVals con, "AccTitle", "VVendersANDMakers", "AccNo", "WHERE Active=1"
     
End Sub
Public Sub RefreshLV()

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
End Sub
Private Sub Form_Resize()

 On Error Resume Next
 fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
 
 Dim TitleBottom As Long

 TitleBottom = cmbItem.Top + cmbItem.Height 'fTitle.Top + fTitle.Height
 
 LV.Move 50, TitleBottom, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + cmdClose.Height + 100)
 
 cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 50
 fupdate.Move LV.Left, LV.Top + LV.Height + 50
'cmdCheckAll.Move LV.Left, LV.Top + LV.Height + 50
 
 
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With cmdUnAssign
        .Move LV.Left + LV.ColumnHeaders(LV.ColumnHeaders.count).Width + LV.ColumnHeaders(LV.ColumnHeaders.count).Left + 10, Item.Top + LV.Top
        .Visible = True
    End With
    PicEdit.Visible = False
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        With LV.ColumnHeaders(i)
            If x > Val(.Left) And x < (Val(.Left) + Val(.Width)) Then
                iColNo = i
                Exit For
            End If
        End With
    Next
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If LV.ListItems.count = 0 Then Exit Sub
    If Button <> 2 Then Exit Sub
    If bHideRate Then Exit Sub
    
    Dim strRet As String
    strRet = ShowPopUpMenu1("Show History")

    If strRet = "Show History" Then
        Dim f As New frmVendorAssItemHistory
        Load f
        f.ViewHistory (Val(LV.SelectedItem.key))
    End If
    
End Sub

Private Sub txtPlant_KeyPress(KeyAscii As Integer)
    On Error GoTo err
    
    If KeyAscii = 13 Then
        'txtStamp.SetFocus
        If Not (txtSnaff.Enabled Or txtStamp.Enabled) Then
            Call txtSnaff_KeyPress(13)
        End If
    ElseIf KeyAscii = 27 Then
        'txtRate.Visible = False
        PicEdit.Visible = False
    End If
    Call OnlyNums(KeyAscii)
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    
    If KeyAscii = 13 Then
        Dim iAffect As Integer
        con.Execute "UPDATE VenderAssItems SET Rate=" & Val(txtRate) & " Where EntryID=" & Val(LV.SelectedItem.key), iAffect
        If iAffect = 1 Then
            con.Execute "INSERT INTO VenderAssItems_Revisions(VAI_RefID,Rate,UserName,MachineName,Remarks) VALUES(" & Val(LV.SelectedItem.key) & "," & Val(LV.SelectedItem.SubItems(4)) & ",'" & CurrentUserName & "','" & strComputerName & "','" & LV.SelectedItem.SubItems(5) & "')"
            LV.SelectedItem.SubItems(4) = Val(txtRate)
            PicEdit.Visible = False
        End If
    ElseIf KeyAscii = 27 Then
        PicEdit.Visible = False
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If txtRemarks <> "" Then
            Dim iAffect As Integer
            con.Execute "UPDATE VenderAssItems SET Remarks='" & txtRemarks & "' WHERE EntryID=" & Val(LV.SelectedItem.key), iAffect
            If iAffect = 1 Then
                LV.SelectedItem.SubItems(5) = txtRemarks
                txtRemarks.Visible = False
            End If
        End If
    ElseIf KeyAscii = 27 Then
        txtRemarks.Visible = False
    End If
    
End Sub

Private Sub txtRemarks_LostFocus()
    txtRemarks.Visible = False
End Sub

Private Sub txtSnaff_KeyPress(KeyAscii As Integer)

    On Error GoTo err
    
    If KeyAscii = 13 Then
        Dim iAffect As Integer
        con.Execute "Update VendAssItems Set Rate=" & Val(txtRate) & ",PlantRate=" & Val(txtPlant) & ",StampRate=" & Val(txtStamp) & ",SnaffRate=" & Val(txtSnaff) & "  Where EntryID=" & Val(LV.SelectedItem.key), iAffect
        If iAffect = 1 Then
            LV.SelectedItem.SubItems(4) = Val(txtRate)
            LV.SelectedItem.SubItems(5) = Val(txtPlant)
            LV.SelectedItem.SubItems(6) = Val(txtStamp)
            LV.SelectedItem.SubItems(7) = Val(txtSnaff)
            LV.SelectedItem.SubItems(8) = Val(txtSnaff) + Val(txtStamp) + Val(txtPlant) + Val(txtRate)
            PicEdit.Visible = False
        End If
    ElseIf KeyAscii = 27 Then
        'txtRate.Visible = False
        PicEdit.Visible = False
    End If
    Call OnlyNums(KeyAscii)
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub txtStamp_KeyPress(KeyAscii As Integer)
    On Error GoTo err
    
    If KeyAscii = 13 Then
        'txtSnaff.SetFocus
        If Not (txtSnaff.Enabled) Then
            Call txtSnaff_KeyPress(13)
        End If
    ElseIf KeyAscii = 27 Then
        'txtRate.Visible = False
        PicEdit.Visible = False
    End If
    Call OnlyNums(KeyAscii)
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
