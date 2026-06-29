VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "combolist.ocx"
Begin VB.Form frmRMAssign 
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
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9765
      TabIndex        =   21
      Top             =   630
      Width           =   900
   End
   Begin ComboList.Usercontrol1 cmbCust 
      Height          =   285
      Left            =   180
      TabIndex        =   0
      Top             =   855
      Width           =   9540
      _ExtentX        =   16828
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
   Begin ComboList.Usercontrol1 cmbRMGroup 
      Height          =   285
      Left            =   180
      TabIndex        =   20
      Top             =   1380
      Width           =   2430
      _ExtentX        =   4286
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
   Begin VB.CheckBox chkShowEmployees 
      BackColor       =   &H00AFBDCB&
      Caption         =   "Show Employees"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7920
      TabIndex        =   19
      Top             =   645
      Width           =   1770
   End
   Begin VB.TextBox txtFrequency 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8730
      TabIndex        =   3
      Top             =   1380
      Width           =   990
   End
   Begin VB.TextBox txtQty 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7740
      TabIndex        =   2
      Top             =   1380
      Width           =   1000
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   345
      TabIndex        =   18
      Top             =   2160
      Visible         =   0   'False
      Width           =   645
   End
   Begin ComboList.Usercontrol1 cmbItemName 
      Height          =   285
      Left            =   2610
      TabIndex        =   1
      Top             =   1380
      Width           =   5130
      _ExtentX        =   9049
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
      TabIndex        =   14
      Top             =   6045
      Width           =   5250
      Begin VB.CheckBox chkLoadCombo 
         Caption         =   "Load Combo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3120
         TabIndex        =   17
         Top             =   30
         Visible         =   0   'False
         Width           =   1470
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   375
         Left            =   15
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   1500
         ForeColor       =   0
         Caption         =   "Print"
         PicturePosition =   327683
         Size            =   "2646;661"
         Accelerator     =   80
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
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
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
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
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
      Picture         =   "frmRMAssign.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
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
      TabIndex        =   9
      Top             =   0
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assign Raw Material"
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
         Left            =   3705
         TabIndex        =   10
         Top             =   120
         Width           =   3240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assign Raw Material"
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
         Left            =   3720
         TabIndex        =   11
         Top             =   135
         Width           =   3240
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4290
      Left            =   180
      TabIndex        =   8
      Top             =   1725
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   7567
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
         Text            =   "RM"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Frequency"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   1
      Left            =   180
      TabIndex        =   15
      Top             =   1155
      Width           =   9540
      BackColor       =   11517387
      Caption         =   $"frmRMAssign.frx":0102
      Size            =   "16828;423"
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
      Height          =   525
      Left            =   9765
      TabIndex        =   4
      Top             =   1155
      Width           =   900
      ForeColor       =   6244673
      Caption         =   "Assign   "
      PicturePosition =   327683
      Size            =   "1587;926"
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
      Left            =   180
      TabIndex        =   12
      Top             =   630
      Width           =   9540
      BackColor       =   11517387
      Caption         =   " Maker"
      Size            =   "16828;423"
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
      Left            =   9060
      TabIndex        =   7
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
Attribute VB_Name = "frmRMAssign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iColNo As Integer

Private Sub chkLoadCombo_Click()

    If chkLoadCombo.Value = vbChecked Then
        If cmbCust.MatchFound Then
            '
            Call AddInCombo
        End If
    End If
    
End Sub

Private Sub AddInCombo()
    Call cmbItemName.AddVals(con, "'{' + ItemID + '} '  + IsNull(ItemName,'') + ' ' + IsNull(ItemSize,'') + ' ' + IsNull(SizeUnit,'') + ' ' + IsNull(Type1,'') + ' ' + IsNull(FinQuality,'') ", "VItems", "ItemID", " WHERE ItemID IN(SELECT ItemID FROM ItemProcesses WHERE ProcessID=" & cmbProcess.id & ") AND ItemID NOT IN(SELECT ItemID FROM VendAssItems WHERE VendID=" & cmbCust.id & " AND ProcessID=" & cmbProcess.id & ")  ORDER BY ItemName")
End Sub

Private Sub FillItems()

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
   
    With rs
        .Open "Select EntryID,ProcessID,ItemID,ItemName,Description,Rate,PlantRate,SnaffRate,StampRate,PhaseID,Unit,AssignedUnit,Remarks From VMakerAssItems Where VendID=" & cmbCust.id & " AND ProcessID=" & cmbProcess.id, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, ![EntryID] & "'", ![ITemID])
            ITM.Tag = Val(!ProcessID & "")
            ITM.ListSubItems.add , , ![ItemName] & ""
            ITM.ListSubItems.add , , "-" '![Type] & ""
            ITM.ListSubItems.add , , ![Description] & ""
            ITM.ListSubItems.add , , ![Rate] & ""
            ITM.ListSubItems.add(, , !AssignedUnit & "").Tag = !Unit & ""
            ITM.ListSubItems.add(, , Val(!Rate & "") * 12).Tag = ""
            
            ITM.ListSubItems.add , , !Remarks & "" & ""
'            If cmbCust.Tag = 7 Then
'                ITM.SubItems(5) = ![PlantRate] & ""
'                ITM.SubItems(6) = ![StampRate] & ""
'                ITM.SubItems(7) = ![SnaffRate] & ""
'                ITM.SubItems(8) = Val(ITM.SubItems(7) & "") + Val(ITM.SubItems(6) & "") + Val(ITM.SubItems(5) & "") + Val(ITM.SubItems(4) & "")
'            End If
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
End Sub

Private Sub chkShowEmployees_Click()
    
    cmbCust.ClearVals
    If chkShowEmployees.Value = vbChecked Then
        Call cmbCust.AddVals(con, "'{ ' + EmpID + ' } ' + Name", "Employees", "EmpID")
    Else
        Call cmbCust.AddVals(con, "'{ ' + VendID1 + ' } ' + MakerName", "VMakers", "VendID")
    End If
    
End Sub

Private Sub cmbCust_matched()

    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        If chkShowEmployees.Value = vbChecked Then
            .Open "SELECT * FROM VEmployees_RM_Assigned WHERE EmpID='" & cmbCust.id & "' ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        Else
            .Open "SELECT * FROM VMakers_RM_Assigned WHERE MakerID=" & cmbCust.id & " ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        End If
        
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , !RMID1 & "")
            ITM.Tag = Val(!EntryID & "")
            ITM.ListSubItems.add , , !RMName & ""
            ITM.ListSubItems.add , , Val(!MaxQty & "")
            ITM.ListSubItems.add , , Val(!MaxFrequency & "")
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    
    
    
End Sub

Private Sub cmbCust_UnMatched()
    LV.ListItems.Clear
    'cmbCountry.ClearVals
End Sub

Private Sub cmbItem_matched()
    cmbItemName.id = cmbItem.id
End Sub

Private Sub cmbItemName_matched()
    'cmbItem.id = cmbItemName.id
    'txtItemCode.Text = cmbItemName.id
End Sub

Private Sub cmbProcess_matched()

    cmbItemName.ClearVals
    
    If chkLoadCombo.Value = vbChecked Then
        Call cmbItemName.AddVals(con, "'{' + ItemID + '} '  + IsNull(ItemName,'') + ' ' + IsNull(ItemSize,'') + ' ' + IsNull(SizeUnit,'') + ' ' + IsNull(Type1,'') + ' ' + IsNull(FinQuality,'') ", "VItems", "ItemID", " Where ItemID In(Select ItemID From ItemProcesses Where ProcessID=" & cmbProcess.id & ") and ItemID Not In(Select ItemID From VendAssItems Where VendID=" & cmbCust.id & " AND ProcessID=" & cmbProcess.id & ")  Order By ItemName")
    End If
    
    cmbCust.Tag = cmbProcess.id
    
    Call FillItems
    
End Sub

Private Sub cmbUnit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        con.Execute "UPDATE VendAssItems SET Unit='" & cmbUnit.Text & "' Where EntryID=" & Val(LV.SelectedItem.key)
        LV.SelectedItem.SubItems(5) = cmbUnit
        cmbUnit.Visible = False
    ElseIf KeyAscii = 27 Then
        cmbUnit.Visible = False
    End If
End Sub


Private Sub cmbRMGroup_matched()

    cmbItemName.ClearVals
    cmbItemName.AddItem "<All Materials>", "0"
    
    If chkShowEmployees.Value = vbChecked Then
        cmbItemName.AddVals con, "'{'+RMID1+'} '+RMName", "RM", "RMID", " WHERE RMID NOT IN(SELECT RMID FROM Employees_RM_Assigned WHERE EmpID='" & cmbCust.id & "') AND GroupID=" & cmbRMGroup.id
    Else
        cmbItemName.AddVals con, "'{'+RMID1+'} '+RMName", "RM", "RMID", " WHERE RMID NOT IN(SELECT RMID FROM Makers_RM_Assigned WHERE MakerID=" & cmbCust.id & ") AND GroupID=" & cmbRMGroup.id
    End If
    
End Sub

Private Sub cmdAssign_Click()

    On Error GoTo err
    If cmbItemName.MatchFound = False Then Exit Sub
    If cmbCust.MatchFound = False Then Exit Sub
    
    Dim strCondition As String
    If cmbItemName.id = "0" Then
        If chkShowEmployees.Value = vbChecked Then
            strCondition = " WHERE GroupID=" & cmbRMGroup.id & " AND RMID NOT IN(SELECT RMID FROM Employees_RM_Assigned WHERE EmpID='" & cmbCust.id & "')"
        Else
            strCondition = " WHERE GroupID=" & cmbRMGroup.id & " AND RMID NOT IN(SELECT RMID FROM Makers_RM_Assigned WHERE MakerID=" & cmbCust.id & ")"
        End If
    Else
        strCondition = " WHERE RMID=" & cmbItemName.id
    End If
    
    Call StartTrans(con)
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT RMID FROM RM " & strCondition, con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            If chkShowEmployees.Value = vbChecked Then
                con.Execute "INSERT INTO Employees_RM_Assigned(EmpID,RMID,MaxQty,MaxFrequency) VALUES('" & cmbCust.id & _
                 "'," & !RMID & "," & Val(txtQty) & "," & Val(txtFrequency) & ")"
            Else
                con.Execute "INSERT INTO Makers_RM_Assigned(MakerID,RMID,MaxQty,MaxFrequency) VALUES(" & cmbCust.id & _
                 "," & !RMID & "," & Val(txtQty) & "," & Val(txtFrequency) & ")"
            End If
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    
    
    
    con.CommitTrans
    
    Call cmbCust_matched
    
    Exit Sub
err:
    MsgBox err.Description
End Sub



Private Sub cmdCheckAll_Click()
        
    Exit Sub
     
    If cmbCust.MatchFound = False Then Exit Sub
    
    'If cmbItem.MatchFound = False Then Exit Sub
    
    Dim Saved As Integer
    Dim cmd As New ADODB.Command
    
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_VendAssItem"
        .Parameters("@VendID").Value = cmbCust.id 'cmbCust.ID
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

    Exit Sub
    On Error GoTo err
    Dim iAffect As Integer
    If MsgBox("Are You Sure To UnAssign All Items?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        con.Execute "Delete From VendAssItems Where VendID=" & cmbCust.id, iAffect
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

Private Sub cmdHide_Click()
    PicSearch.Visible = False
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo err
    If cmbCust.MatchFound = False Then Exit Sub
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\MakerAssignedItems.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VendAssItems.VendID}=" & cmbCust.id, rpt
    Exit Sub
err:
    MsgBox err.Description
End Sub





Private Sub lstSearch_KeyPress(KeyAscii As MSForms.ReturnInteger)
'    MsgBox KeyAscii
End Sub

Private Sub cmdRefresh_Click()
    Call cmbCust_matched
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    LV.Sorted = True
    If LV.SortKey = ColumnHeader.Index - 1 Then
        If LV.SortOrder = lvwAscending Then
            LV.SortOrder = lvwDescending
            'ColumnHeader.Icon = "Up"
        Else
            LV.SortOrder = lvwAscending
            'ColumnHeader.Icon = "Down"
        End If
    Else
        LV.ColumnHeaderIcons = Nothing
        'Set LV.ColumnHeaderIcons = ImageList2
        LV.SortKey = ColumnHeader.Index - 1
        LV.SortOrder = lvwAscending
        'ColumnHeader.Icon = "Down"
    End If

    On Error Resume Next
    LV.SelectedItem.EnsureVisible

End Sub

Private Sub LV_DblClick()
    
    Exit Sub
    If LV.ListItems.count = 0 Then Exit Sub
    If cmbCust.MatchFound = False Then Exit Sub
    If bCanChangeRate = False Then
        MsgBox "You don't have the Access to Change Rate.", vbInformation
        Exit Sub
    End If
    
    If Not LV.SelectedItem.Checked Then
        LV.SelectedItem.Checked = True
        'Call LV_ItemCheck(LV.SelectedItem)
        Exit Sub
    End If
    
    If iColNo = 6 Then
        With cmbUnit
            .Clear
            If LV.SelectedItem.ListSubItems(5).Tag & "" <> "" Then .AddItem LV.SelectedItem.SubItems(5)
            If LV.SelectedItem.ListSubItems(5) & "" <> "" And LV.SelectedItem.ListSubItems(5) <> LV.SelectedItem.ListSubItems(5).Tag Then .AddItem LV.SelectedItem.SubItems(5)
            If LV.SelectedItem.ListSubItems(5) <> "Kgs" Then
                cmbUnit.AddItem "Kgs"
            End If
            .Move LV.Left + LV.ColumnHeaders(6).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(6).Width
            .Text = LV.SelectedItem.ListSubItems(5)
            .Visible = True
            .SetFocus
        End With
        Exit Sub
    End If
           
'    With txtRate
'        PicEdit.Move LV.Left + LV.ColumnHeaders(5).Left, LV.Top + LV.SelectedItem.Top ', LV.ColumnHeaders(5).Width
'        If LV.ColumnHeaders.count = 9 Then
'            Dim rs As New ADODB.Recordset
'            With rs
'                .Open "Select IsNull(Planter,0),IsNull(Snaffer,0),IsNull(Stamper,0) From Venders Where VendID=" & cmbCust.ID, con, adOpenForwardOnly, adLockReadOnly
'
'                txtPlant.Enabled = False
'                txtSnaff.Enabled = False
'                txtStamp.Enabled = False
'                .Close
'
'            End With
'
'            PicEdit.Width = 3990
'            txtPlant.Visible = True
'            txtSnaff.Visible = True
'            txtStamp.Visible = True
'            txtPlant.Text = LV.SelectedItem.SubItems(5)
'            txtStamp.Text = LV.SelectedItem.SubItems(6)
'            txtSnaff.Text = LV.SelectedItem.SubItems(7)
'        Else
'            PicEdit.Width = .Width
'            txtPlant.Visible = False
'            txtSnaff.Visible = False
'            txtStamp.Visible = False
'        End If
'        PicEdit.Visible = True
'        .Text = LV.SelectedItem.SubItems(4)
'        .Visible = True
'        .SelStart = 0
'        .SelLength = Len(.Text)
'        .SetFocus
'    End With
    
    If iColNo <> 8 Then iColNo = 5
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(iColNo).Width
        .Text = LV.SelectedItem.SubItems(iColNo - 1)
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With

End Sub





Private Sub Form_Load()

    cmbCust.ListHeight = 4000
    Call chkShowEmployees_Click
    
    cmbItemName.ListHeight = 2600

    iColNo = 0
     
    cmbRMGroup.ListHeight = 4000
    cmbRMGroup.AddVals con, "Description", "RMGroups", "ID", " ORDER BY Description"
    
End Sub

Public Sub RefreshLV()

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "Select *  from FCustomerCatalog Where CustCode='" & cmbCust.id & "' and Country='" & cmbCountry.id & "'", con, adOpenForwardOnly, adLockReadOnly
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

    TitleBottom = cmbItemName.Top + cmbItemName.Height 'fTitle.Top + fTitle.Height
 
    LV.Move 50, TitleBottom, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + cmdClose.Height + 100)
 
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 50
    fupdate.Move LV.Left, LV.Top + LV.Height + 50
    
    'cmdCheckAll.Move LV.Left, LV.Top + LV.Height + 50
 
 
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
'    With cmdUnAssign
'        .Move LV.Left + LV.ColumnHeaders(LV.ColumnHeaders.count).Width + LV.ColumnHeaders(LV.ColumnHeaders.count).Left + 10, Item.Top + LV.Top
'        .Visible = True
'    End With
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)


    On Error GoTo err
    If LV.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbYes Then
        
            If chkShowEmployees.Value = vbChecked Then
                con.Execute "DELETE FROM Employees_RM_Assigned WHERE EntryID=" & Val(LV.SelectedItem.Tag)
            Else
                con.Execute "DELETE FROM Makers_RM_Assigned WHERE EntryID=" & Val(LV.SelectedItem.Tag)
            End If
                
            Call cmbCust_matched
        End If
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button <> 1 Then Exit Sub
    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        If X > LV.ColumnHeaders(i).Left And X < LV.ColumnHeaders(i).Left + LV.ColumnHeaders(i).Width Then
            iColNo = i
            Exit For
        End If
    Next
    
End Sub



Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button <> 2 Then Exit Sub
'    Dim strRet As String
'    strRet = ShowPopUpMenu1("Show History")
'
'    If strRet = "Show History" Then
'        Dim f As New frmMakerAssItemHistory
'        Load f
'        f.ViewHistory (Val(LV.SelectedItem.key))
'    End If
    
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    
    If KeyAscii = 13 Then
        If iColNo = 5 Then
            Dim iAffect As Integer
            con.Execute "UPDATE VendAssItems SET Rate=" & Val(txtEdit) & " WHERE EntryID=" & Val(LV.SelectedItem.key), iAffect
            If iAffect = 1 Then
                con.Execute "INSERT INTO VendAssItems_Revisions(VAI_RefID,Rate,UserName,Remarks) VALUES(" & Val(LV.SelectedItem.key) & "," & Val(LV.SelectedItem.SubItems(4)) & ",'" & CurrentUserName & "','" & LV.SelectedItem.SubItems(7) & "')"
                LV.SelectedItem.SubItems(4) = Val(txtEdit)
                LV.SelectedItem.SubItems(6) = Val(txtEdit) * 12
            End If
        Else
            con.Execute "UPDATE VendAssItems SET Remarks='" & txtEdit & "' WHERE EntryID=" & Val(LV.SelectedItem.key)
            LV.SelectedItem.SubItems(iColNo - 1) = txtEdit
        End If
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        'txtedit.Visible = False
        txtEdit.Visible = False
    End If
    'Call OnlyNums(KeyAscii)
    
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
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
        If iColNo = 4 Then
            Dim iAffect As Integer
            con.Execute "UPDATE VendAssItems SET Rate=" & Val(txtRate) & " WHERE EntryID=" & Val(LV.SelectedItem.key), iAffect
            If iAffect = 1 Then
                con.Execute "INSERT INTO VendAssItems_Revisions(VAI_RefID,Rate,UserName) VALUES(" & Val(LV.SelectedItem.key) & "," & Val(txtRate) & ",'" & CurrentUserName & "')"
                LV.SelectedItem.SubItems(4) = Val(txtRate)
                LV.SelectedItem.SubItems(6) = Val(txtRate) * 12
                PicEdit.Visible = False
            End If
        Else
            con.Execute "UPDATE VendAssItems SET Remarks='" & txtRate & "' WHERE EntryID=" & Val(LV.SelectedItem.key)
        End If
    ElseIf KeyAscii = 27 Then
        'txtRate.Visible = False
        PicEdit.Visible = False
    End If
    'Call OnlyNums(KeyAscii)
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub txtSearch_Change()

    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT '{' + ItemID + '} '  + ItemName + ' ' + ItemSize + SizeUnit,ItemID FROM VItems WHERE ItemID LIKE '%" & txtSearch & "%' AND ItemID IN(SELECT ItemID FROM ItemProcesses WHERE ProcessID=" & cmbProcess.id & ") AND ItemID NOT IN(SELECT ItemID FROM VendAssItems WHERE VendID=" & cmbCust.id & " AND ProcessID=" & cmbProcess.id & ")", con, adOpenForwardOnly, adLockReadOnly
        lstSearch.Clear
        Do Until .EOF
            lstSearch.AddItem .Fields(0).Value & ""
            lstSearch.List(lstSearch.ListCount - 1, 1) = !ITemID & ""
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDown Then
        If lstSearch.ListCount > 0 Then
            If lstSearch.ListIndex < (lstSearch.ListCount - 1) Then
                lstSearch.ListIndex = lstSearch.ListIndex + 1
            End If
            KeyCode = 0
        End If
    ElseIf KeyCode = vbKeyUp Then
        If lstSearch.ListCount > 0 Then
            If lstSearch.ListIndex > 0 Then
                lstSearch.ListIndex = lstSearch.ListIndex - 1
            End If
            KeyCode = 0
        End If
    End If
    
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
