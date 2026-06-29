VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmVendAssHR 
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6540
   ScaleWidth      =   10170
   WindowState     =   2  'Maximized
   Begin ComboList.Usercontrol1 cmbCust 
      Height          =   285
      Left            =   135
      TabIndex        =   5
      Top             =   855
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
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assigned Resources"
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
         TabIndex        =   7
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assigned Resources"
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
         TabIndex        =   8
         Top             =   135
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdUnAssign 
      Height          =   240
      Left            =   9585
      Picture         =   "frmVendAssHR.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "UnAssign Selected Item..."
      Top             =   1230
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Frame fupdate 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   390
      Left            =   120
      TabIndex        =   1
      Top             =   6045
      Width           =   3090
      Begin MSForms.CommandButton cmdClearAll 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
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
      Begin MSForms.CommandButton cmdCheckAll 
         Height          =   375
         Left            =   15
         TabIndex        =   2
         Top             =   0
         Width           =   1500
         ForeColor       =   0
         Caption         =   " Assign All"
         PicturePosition =   327683
         Size            =   "2646;661"
         Picture         =   "frmVendAssHR.frx":0102
         Accelerator     =   104
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4290
      TabIndex        =   0
      Top             =   2055
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4785
      Left            =   150
      TabIndex        =   9
      Top             =   1230
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   8440
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
         Text            =   "Item ID"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   5997
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Group"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSForms.TextBox txtRateAdd 
      Height          =   285
      Left            =   8325
      TabIndex        =   15
      Top             =   855
      Width           =   870
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "1535;503"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbItem 
      Height          =   285
      Left            =   2985
      TabIndex        =   14
      Tag             =   "SubVendType"
      Top             =   855
      Width           =   2340
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "4128;503"
      ListWidth       =   8819
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbItemName 
      Height          =   285
      Left            =   5340
      TabIndex        =   13
      Tag             =   "SubVendType"
      Top             =   855
      Width           =   2970
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5239;503"
      ListWidth       =   8819
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   375
      Left            =   8370
      TabIndex        =   12
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
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   135
      TabIndex        =   11
      Top             =   630
      Width           =   9060
      BackColor       =   11517387
      Caption         =   $"frmVendAssHR.frx":025C
      Size            =   "15981;423"
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
      Left            =   9255
      TabIndex        =   10
      Top             =   765
      Width           =   1125
      ForeColor       =   6244673
      Caption         =   "Assign   "
      PicturePosition =   327683
      Size            =   "1984;661"
      Picture         =   "frmVendAssHR.frx":02F9
      Accelerator     =   65
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmVendAssHR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lSelRID As Long
Public Sub ShowForm(lRID As Long)
    lSelRID = lRID
    Me.Show
End Sub

Private Sub cmbCust_matched()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
'    cmbItem.ClearVals
'    Call cmbItem.AddVals(con, "RID1", "HICOResources", "RID", " Where RID Not In(Select RID From VendAssResources Where InUse=1)")
'
'    cmbItemName.ClearVals
'    Call cmbItemName.AddVals(con, "RName", "HICOResources", "RID", " Where RID Not In(Select RID From VendAssResources Where InUse=1)")
    
    'cmbCust.Tag = Val(con.Execute("Select PhaseID From Venders Where VendID=" & cmbCust.ID).Fields(0).value & "")
    With rs
        .Open "Select RID,RID1,RName From HICOResources Where RID Not In(Select RID From VendAssResources Where InUse=1)", con, adOpenForwardOnly, adLockReadOnly
        cmbItem.Clear
        cmbItemName.Clear
        Do Until .EOF
            cmbItem.AddItem ![RID1] & ""
            cmbItem.List(.AbsolutePosition - 1, 1) = ![RID] & ""
            cmbItemName.AddItem ![RName] & ""
            cmbItemName.List(.AbsolutePosition - 1, 1) = ![RID] & ""
            .MoveNext
        Loop
        .Close
        
        If lSelRID <> 0 Then
            For i = 0 To cmbItem.ListCount - 1
                If cmbItem.List(i, 1) = lSelRID Then
                    cmbItem.ListIndex = i
                    Exit For
                End If
            Next
        End If
                
        .Open "Select EntryID,RID,RID1,RName,RUsage,Rate From VVendAssResources Where VendID=" & cmbCust.ID, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.Add(, ![EntryID] & "'", ![RID1])
            ITM.SubItems(1) = ![RName] & ""
            'ITM.SubItems(2) = ![Type] & ""
            'ITM.SubItems(3) = ![Description] & ""
            ITM.SubItems(4) = ![Rate] & ""
            .MoveNext
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






Private Sub cmbItem_Change()
    Call cmbItem_Click
End Sub

Private Sub cmbItem_Click()
    cmbItemName.ListIndex = cmbItem.ListIndex
End Sub

Private Sub cmbItemName_Change()
    Call cmbItemName_Click
End Sub

Private Sub cmbItemName_Click()
    cmbItem.ListIndex = cmbItemName.ListIndex
End Sub

Private Sub cmdAssign_Click()
        
    If cmbCust.MatchFound = False Then Exit Sub
    
    If cmbItem.MatchFound = False Then Exit Sub
    Dim lRID As Long
    lRID = cmbItem.List(cmbItem.ListIndex, 1)
    Call StartTrans(con)
    con.Execute "Insert Into VendAssResources(VendID,RID,Rate,InUse,AssDT) Values(" & _
     cmbCust.ID & "," & lRID & "," & Val(txtRateAdd.Text) & _
     ",1,'" & ServerDate & "')"
    
    con.Execute "Update HICOResources Set InUse=1 Where RID=" & lRID
    con.CommitTrans
    
    
    Call cmbCust_matched
End Sub



Private Sub cmdCheckAll_Click()
        
    
     
    If cmbCust.MatchFound = False Then Exit Sub
    If cmbCust.MatchFound = False Then Exit Sub
    'If cmbItem.MatchFound = False Then Exit Sub
    
    Dim Saved As Integer
    Dim cmd As New ADODB.Command
    
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_VendAssHR"
        .Parameters("@VendID").value = cmbCust.ID 'cmbCust.ID
        .Parameters("@RID").value = "0" 'cmbCountry.ID
        '.Parameters("@ProcessID").value = cmbCust.Tag
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
    If MsgBox("Are You Sure To UnAssign All Resources?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        con.Execute "Delete From VendAssResources Where VendID=" & cmbCust.ID, iAffect
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
    If LV.ListItems.Count = 0 Then Exit Sub
    Dim iAffect As Integer
    
    con.Execute "Delete From VendAssResources Where EntryID=" & Val(LV.SelectedItem.Key), iAffect
    If iAffect > 0 Then
        Call cmbCust_matched
        cmdUnAssign.Visible = False
    End If
    Exit Sub
err:
    MsgBox err.Description
    
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

    If LV.ListItems.Count = 0 Then Exit Sub
    If cmbCust.MatchFound = False Then Exit Sub

    If Not LV.SelectedItem.Checked Then
        LV.SelectedItem.Checked = True
        'Call LV_ItemCheck(LV.SelectedItem)
        Exit Sub
    End If

    With txtRate
        .Move LV.Left + LV.ColumnHeaders(5).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(5).Width
        .Text = LV.SelectedItem.SubItems(4)
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With

    

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    
    lSelRID = 0
    cmbCust.ListHeight = 2600
    'cmbNick.ListHeight = 2600
    Call cmbCust.AddVals(con, "VenderName", "Venders", "VendID", " Where VendType In('0','1')")
    
    
End Sub
Public Sub RefreshLV()

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "Select *  from FCustomerCatalog Where CustCode='" & cmbCust.ID & "' and Country='" & cmbCountry.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.Add(, ![CustCode] & ![Country] & ![ITemID], ![ITemID] & "")
            ITM.SubItems(1) = ![Description] & ""
            ITM.SubItems(2) = ![Unit] & ""
            ITM.SubItems(3) = ![FOB] & ""
            ITM.SubItems(4) = ![CIFSea] & ""
            ITM.SubItems(5) = ![CIFAir] & ""
            ITM.SubItems(6) = ![CnFSea] & ""
            ITM.SubItems(7) = ![CnFAir] & ""
            .MoveNext
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
        .Move LV.Left + LV.ColumnHeaders(LV.ColumnHeaders.Count).Width + LV.ColumnHeaders(LV.ColumnHeaders.Count).Left + 10, Item.Top + LV.Top
        .Visible = True
    End With
    txtRate.Visible = False
End Sub

Private Sub txtAddQty_Change()

End Sub

Private Sub txtAddQty_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Call OnlyNums(KeyAscii.value, True)
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    
    If KeyAscii = 13 Then
        Dim iAffect As Integer
        con.Execute "Update VendAssResources Set Rate=" & Val(txtRate) & " Where EntryID=" & Val(LV.SelectedItem.Key), iAffect
        If iAffect = 1 Then
            LV.SelectedItem.SubItems(4) = Val(txtRate)
            txtRate.Visible = False
        End If
    ElseIf KeyAscii = 27 Then
        txtRate.Visible = False
    End If
    Call OnlyNums(KeyAscii)
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

