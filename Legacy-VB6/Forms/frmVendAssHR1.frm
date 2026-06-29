VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmVendAssHR1 
   ClientHeight    =   6570
   ClientLeft      =   1110
   ClientTop       =   60
   ClientWidth     =   10095
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
   ScaleWidth      =   10095
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7650
      TabIndex        =   13
      Top             =   2010
      Visible         =   0   'False
      Width           =   1000
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   6285
      TabIndex        =   12
      Top             =   1995
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   56557571
      CurrentDate     =   37679
   End
   Begin ComboList.Usercontrol1 cmbVend 
      Height          =   285
      Left            =   2235
      TabIndex        =   14
      Top             =   2010
      Visible         =   0   'False
      Width           =   4020
      _ExtentX        =   7091
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
   Begin VB.ComboBox cmb1 
      Height          =   315
      ItemData        =   "frmVendAssHR1.frx":0000
      Left            =   705
      List            =   "frmVendAssHR1.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   855
      Width           =   2550
   End
   Begin VB.Frame fupdate 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   390
      Left            =   120
      TabIndex        =   5
      Top             =   6045
      Visible         =   0   'False
      Width           =   3090
      Begin MSForms.CommandButton cmdCheckAll 
         Height          =   375
         Left            =   15
         TabIndex        =   7
         Top             =   0
         Width           =   1500
         ForeColor       =   0
         Caption         =   " Assign All"
         PicturePosition =   327683
         Size            =   "2646;661"
         Picture         =   "frmVendAssHR1.frx":005C
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
         TabIndex        =   6
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
      Picture         =   "frmVendAssHR1.frx":01B6
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "UnAssign Selected Item..."
      Top             =   1230
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   435
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   570
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
            Picture         =   "frmVendAssHR1.frx":02B8
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVendAssHR1.frx":0474
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmVendAssHR1.frx":0630
            Key             =   "Active"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVendAssHR1.frx":0A84
            Key             =   "Inactive"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fTitle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   0
      TabIndex        =   0
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
         TabIndex        =   1
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
         TabIndex        =   2
         Top             =   135
         Width           =   3015
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4785
      Left            =   150
      TabIndex        =   8
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Hico Resource"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   5997
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Vend ID"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Vend Name"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Ass Date"
         Object.Width           =   2408
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Rate"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   720
      TabIndex        =   10
      Top             =   630
      Width           =   2535
      BackColor       =   11517387
      Caption         =   "  Show"
      Size            =   "4471;423"
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
      Left            =   8925
      TabIndex        =   9
      Top             =   765
      Visible         =   0   'False
      Width           =   1125
      ForeColor       =   6244673
      Caption         =   "Assign   "
      PicturePosition =   327683
      Size            =   "1984;661"
      Picture         =   "frmVendAssHR1.frx":0ED8
      Accelerator     =   65
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
      Left            =   8370
      TabIndex        =   3
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
Attribute VB_Name = "frmVendAssHR1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rptSQL As String
Private Sub cmbCust_matched()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim itm As ListItem
    
    cmbItem.ClearVals
    Call cmbItem.AddVals(con, "RID1", "HICOResources", "RID", " Where RID Not In(Select RID From VendAssResources Where InUse=1)")
    
    cmbItemName.ClearVals
    Call cmbItemName.AddVals(con, "RName", "HICOResources", "RID", " Where RID Not In(Select RID From VendAssResources Where InUse=1)")
    
    'cmbCust.Tag = Val(con.Execute("Select PhaseID From Venders Where VendID=" & cmbCust.ID).Fields(0).value & "")
    With rs
        .Open "Select EntryID,RID,RID1,RName,RUsage,Rate From VVendAssResources Where VendID=" & cmbCust.ID, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set itm = LV.ListItems.Add(, ![EntryID] & "'", ![RID1])
            itm.SubItems(1) = ![RName] & ""
            'ITM.SubItems(2) = ![Type] & ""
            'ITM.SubItems(3) = ![Description] & ""
            itm.SubItems(4) = ![Rate] & ""
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





Private Sub cmbItem_matched()
    cmbItemName.ID = cmbItem.ID
End Sub

Private Sub cmbItemName_matched()
    cmbItem.ID = cmbItemName.ID
End Sub

Private Sub cmb1_Click()

    If cmb1.ListIndex = 0 Then
        Call RefreshLV
    ElseIf cmb1.ListIndex = 1 Then
        Call RefreshLV(" Where InUse=1")
    ElseIf cmb1.ListIndex = 2 Then
        Call RefreshLV(" Where InUse=0")
    End If
    
End Sub

Private Sub cmbVend_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmbVend.MatchFound Then
            DTPicker1.SetFocus
        Else
            'KeyAscii = 0
        End If
    ElseIf KeyAscii = 27 Then
        cmbVend.Visible = False
        DTPicker1.Visible = False
        txtRate.Visible = False
    End If
End Sub

Private Sub cmdAssign_Click()
        
    If cmbCust.MatchFound = False Then Exit Sub
    
    If cmbItem.MatchFound = False Then Exit Sub
    
    Dim Saved As Integer
    Dim cmd As New ADODB.Command
    
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_VendAssHR"
        .Parameters("@VendID").value = cmbCust.ID
        .Parameters("@RID").value = cmbItem.ID 'cmbNick.List(cmbNick.ListIndex, 1)
        '.Parameters("@ProcessID").value = cmbCust.Tag
        .Execute Saved
    End With
    
    If Saved = 1 Then
        
        Call cmbCust_matched
        
    End If
    
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



Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txtRate.SetFocus
    ElseIf KeyCode = 27 Then
        cmbVend.Visible = False
        DTPicker1.Visible = False
        txtRate.Visible = False
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
    
    Exit Sub
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
Private Sub Form_Load()

    'cmbCust.ListHeight = 2600
    'cmbNick.ListHeight = 2600
    'cmbItem.ListHeight = 2600
    cmbVend.ListHeight = 2600
    
    'Call cmbCust.AddVals(con, "VenderName", "Venders", "VendID", " Where VendType In('0','1')")
    cmb1.ListIndex = 0
    
End Sub
Public Sub RefreshLV(Optional Cond As String)

    Dim rs As New ADODB.Recordset
    Dim itm As ListItem
    
    With rs
        
        .Open "Select *  from VHicoResourceStatus " & Cond, con, adOpenForwardOnly, adLockReadOnly
        
        rptSQL = "Select *  from VHicoResourceStatus " & Cond
        
        LV.ListItems.Clear
        Dim LColor  As Long
        Do Until .EOF
            If ![InUse] Then
                LColor = vbRed
            Else
                LColor = vbBlack
            End If
            Set itm = LV.ListItems.Add(, ![RID] & "'", ![RID1])
            itm.ForeColor = LColor
            
            itm.ListSubItems.Add(, , ![RName] & "").ForeColor = LColor
            
            If LColor = vbRed Then
                itm.Tag = ![EntryID] & ""
                itm.ListSubItems.Add(, , ![VendID1] & "").ForeColor = LColor
                itm.ListSubItems.Add(, , ![VenderName] & "").ForeColor = LColor
                itm.ListSubItems.Add(, , Format(![AssDT] & "", "dd-MMM-yyyy")).ForeColor = LColor
                itm.ListSubItems.Add(, , Val(![Rate] & "")).ForeColor = LColor
            End If
            
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

 TitleBottom = cmb1.Top + cmb1.Height 'fTitle.Top + fTitle.Height
 
 LV.Move 50, TitleBottom, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + cmdClose.Height + 100)
 
 cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 50
 fupdate.Move LV.Left, LV.Top + LV.Height + 50
'cmdCheckAll.Move LV.Left, LV.Top + LV.Height + 50
 
 
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Exit Sub
    With cmdUnAssign
        .Move LV.Left + LV.ColumnHeaders(LV.ColumnHeaders.Count).Width + LV.ColumnHeaders(LV.ColumnHeaders.Count).Left + 10, Item.Top + LV.Top
        .Visible = True
    End With
    txtRate.Visible = False
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If LV.ListItems.Count = 0 Or Button <> 2 Then Exit Sub
    
    Dim strRet As String
    If LV.SelectedItem.ForeColor = vbRed Then
        strRet = ShowPopUpMenu("&UnAssign This Resource", "-", "&Show History For This Resource", "-", "&Print Report")
    Else
        strRet = ShowPopUpMenu("&Assign This Resource", "-", "&Show History For This Resource", "-", "&Print Report")
    End If
    
    Select Case strRet
    
        Case "&UnAssign This Resource"
            con.Execute "Update VendAssResources Set InUse=0,UnAssDT='" & ServerDate & "' Where EntryID=" & LV.SelectedItem.Tag
            Call cmb1_Click
        Case "&Assign This Resource"
            cmbVend.ClearVals
            Call cmbVend.AddVals(con, "VenderName", "Venders", "VendID", " Where VendType In('0','1') Or VendID=129")
            
            cmbVend.Move LV.ColumnHeaders(3).Left + LV.Left, LV.SelectedItem.Top + LV.Top
            DTPicker1.Move LV.ColumnHeaders(5).Left + LV.Left, LV.SelectedItem.Top + LV.Top
            DTPicker1.Visible = True
            txtRate.Move LV.ColumnHeaders(6).Left + LV.Left, LV.SelectedItem.Top + LV.Top
            txtRate.Visible = True
        
            cmbVend.Visible = True
            cmbVend.SetFocus
        Case "&Print Report"
            Call PrintHRrpt
    End Select
    
End Sub
Private Sub PrintHRrpt()
    Call ShowAssHRrpt(rptSQL)
End Sub
Private Sub txtRate_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    
    If KeyAscii = 13 Then
        Dim iAffect As Integer
        con.Execute "Insert Into VendAssResources(VendID,RID,Rate,InUse,AssDT)" & _
         " Values(" & cmbVend.ID & "," & Val(LV.SelectedItem.Key) & "," & _
          Val(txtRate) & ",1,'" & DTPicker1.value & "')", iAffect
        If iAffect = 1 Then
            cmbVend.Visible = False
            DTPicker1.Visible = False
            txtRate.Visible = False
            Call cmb1_Click
        End If
    ElseIf KeyAscii = 27 Then
        cmbVend.Visible = False
        DTPicker1.Visible = False
        txtRate.Visible = False
    End If
    Call OnlyNums(KeyAscii)
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
