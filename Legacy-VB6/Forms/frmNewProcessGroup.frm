VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmNewProcessGroup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Process Groups"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8715
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   8715
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   6255
      Left            =   105
      TabIndex        =   7
      Top             =   195
      Width           =   8565
      Begin VB.TextBox txtHubName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6765
         TabIndex        =   3
         Top             =   1200
         Width           =   1680
      End
      Begin ComboList.Usercontrol1 cmbProcess 
         Height          =   285
         Left            =   150
         TabIndex        =   1
         Top             =   1200
         Width           =   5370
         _ExtentX        =   9472
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
         TextBackColor   =   16777215
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
      Begin VB.TextBox txtSortNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2775
         TabIndex        =   11
         Top             =   3540
         Visible         =   0   'False
         Width           =   1000
      End
      Begin MSComctlLib.ListView LV 
         Height          =   4515
         Left            =   150
         TabIndex        =   4
         Top             =   1590
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   7964
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDropMode     =   1
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   16777215
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
         OLEDropMode     =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Process"
            Object.Width           =   7055
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Scanning"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Hub"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Sort No"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Hub Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   0
         Left            =   6780
         TabIndex        =   12
         Top             =   990
         Width           =   840
      End
      Begin MSForms.ComboBox cmbScanning 
         Height          =   285
         Left            =   5535
         TabIndex        =   2
         Top             =   1200
         Width           =   1215
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2143;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Process :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   9
         Left            =   165
         TabIndex        =   10
         Top             =   990
         Width           =   660
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Scanning :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   7
         Left            =   5580
         TabIndex        =   9
         Top             =   990
         Width           =   750
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Group Description :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   8
         Left            =   165
         TabIndex        =   8
         Top             =   315
         Width           =   1380
      End
      Begin MSForms.TextBox TBox 
         Height          =   345
         Index           =   1
         Left            =   150
         TabIndex        =   0
         Tag             =   "ArticleNo"
         Top             =   540
         Width           =   8295
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "14631;609"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5910
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Please Select Employee's Picture"
      Filter          =   "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   375
      Left            =   5205
      TabIndex        =   5
      Top             =   6585
      Width           =   1695
      Caption         =   " Save            "
      PicturePosition =   327683
      Size            =   "2990;661"
      Picture         =   "frmNewProcessGroup.frx":0000
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
      Height          =   375
      Left            =   6945
      TabIndex        =   6
      Top             =   6585
      Width           =   1695
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "2990;661"
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
Attribute VB_Name = "frmNewProcessGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dFOB As Double, iEntryID As Integer
Dim strDesc As String
Dim strDelList As String
Dim iColNo As Integer
Dim bCopy As Boolean

Public Sub EditItem(p_iEntryID As Integer, Optional p_bCopy As Boolean = False)

    On Error GoTo err
    bCopy = p_bCopy
    Dim rs As New ADODB.Recordset, ITM As ListItem
    
    iEntryID = p_iEntryID

    With rs
        .Open "SELECT * FROM ProcessGroups WHERE EntryID=" & iEntryID, con, adOpenForwardOnly, adLockReadOnly
        TBox(1) = !GroupName
        If bCopy Then
            TBox(1) = TBox(1) & " (Copy)"
        End If
        .Close
        
        .Open "SELECT * FROM VProcessGroups WHERE EntryID=" & iEntryID & " ORDER BY SeqNo,ChildEntryID", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, Val(!Process_RefID & "") & "'", !code & "-" & !Description & "")
            If bCopy = False Then
                ITM.Tag = Val(!ChildEntryID & "")
            End If
            ITM.ListSubItems.add(, , IIf(!Scanning, "Yes", "No")).Tag = Abs(!Scanning)
            ITM.ListSubItems.add , , !Hub_Name & ""
            ITM.ListSubItems.add(, , Val(!SeqNo & "")).Tag = ""
            .MoveNEXT
        Loop
        .Close
         
    End With
    
    Set rs = Nothing
    
    If bCopy Then
        iEntryID = 0
    End If
    
    Me.Show 1
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub FillCmbs()

    cmbInnerMaterial.AddVals con, "'{ ' + MaterialID + ' } '  + MaterialName + ' (' + CAST(ISNULL(OuterLength,'') AS VARCHAR(50)) + ' X ' + CAST(ISNULL(OuterWidth,'') AS VARCHAR(50)) + ' X ' + CAST(ISNULL(OuterHeight,'') AS VARCHAR(50)) + ')'", "VMaterialCartons", "MaterialID", " WHERE InActive=0 AND LEFT(MaterialID,2)='07'"
    cmbMasterCarton.AddVals con, "'{ ' + MaterialID + ' } '  + MaterialName + ' (' + CAST(ISNULL(OuterLength,'') AS VARCHAR(50)) + ' X ' + CAST(ISNULL(OuterWidth,'') AS VARCHAR(50)) + ' X ' + CAST(ISNULL(OuterHeight,'') AS VARCHAR(50)) + ')'", "VMaterialCartons", "MaterialID", " WHERE InActive=0 AND LEFT(MaterialID,2)='07'"
     
End Sub

Private Sub cmbCust_matched()
    cmbItems.AddVals con, "'{ ' + ItemID + ' } ' + ArticleNo + ' ' + ItemName + ' ' + Color + ' (' + ItemSize + ')'", "VFCustomerCatalogActive", "ItemID", " WHERE CustCode='" & cmbCust.ID & "' AND EntryID IN(SELECT MAX(EntryID) FROM VFCustomerCatalogActive WHERE CustCode='" & cmbCust.ID & "' GROUP BY ItemID)"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If SaveItem Then
        Unload Me
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    
    On Error GoTo err
    
    Me.KeyPreview = True

    With cmbScanning
        .AddItem "No"
        .AddItem "Yes"
    End With
    
    With cmbProcess
        .ListHeight = 3000
        .AddVals con, "Code + '-' + Description", "Processes", "ProcessID", " ORDER BY ProcessID"
    End With
    
    iEntryID = 0
  
    strDelList = ""
    bCopy = False
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Function SaveItem() As Boolean

    On Error GoTo err
    
   
    If LV.ListItems.count = 0 Then
        MsgBox "Please Add Processes for this Group.", vbInformation
        Exit Function
    End If
    
    Dim SQL As String
    
    Call StartTrans(con)
    
    Dim lEditID As Long
    If iEntryID = 0 Then
        con.Execute "INSERT INTO ProcessGroups(GroupName) VALUES('" & TBox(1) & "')"
    Else
        con.Execute "UPDATE ProcessGroups SET GroupName='" & TBox(1) & "' WHERE EntryID=" & iEntryID
    End If
   
    If iEntryID = 0 Then
        lEditID = GetSingleLongValue("MAX(EntryID)", "ProcessGroups")
    Else
        lEditID = iEntryID
    End If
    
    If strDelList <> "" Then
        strDelList = Left(strDelList, Len(strDelList) - 1)
        con.Execute "DELETE FROM ProcessGroupsProcesses WHERE EntryID IN(" & strDelList & ")"
    End If
    
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If Val(.Tag) > 0 Then
                con.Execute "UPDATE ProcessGroupsProcesses SET Scanning=" & Val(.ListSubItems(1).Tag) & ",Hub_Name='" & .SubItems(2) & "',SeqNo=" & Val(.SubItems(3)) & " WHERE EntryID=" & Val(.Tag)
            Else
                con.Execute "INSERT INTO ProcessGroupsProcesses(Group_RefID,Process_RefID,Scanning,Hub_Name,SeqNo) VALUES(" & _
                 lEditID & "," & Val(.key) & "," & Val(.ListSubItems(1).Tag) & ",'" & .SubItems(2) & "'," & Val(.SubItems(3)) & ")"
            End If
        End With
    Next
    
    con.CommitTrans
    
    bCopy = False
    SaveItem = True
    
    Exit Function

err:
    MsgBox err.Description
    
End Function

Private Sub LV_DblClick()
    
    If iColNo <> 3 Then iColNo = 4
    With txtSortNo
        .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(iColNo).Width
        If iColNo = 3 Then
            .Text = (LV.SelectedItem.SubItems(iColNo - 1))
        Else
            .Text = Val(LV.SelectedItem.SubItems(iColNo - 1))
        End If
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
    

End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)

    If LV.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Do you want to Delete?", vbQuestion + vbYesNo) = vbYes Then
            If Val(LV.SelectedItem.Tag) > 0 Then
                strDelList = strDelList & Val(LV.SelectedItem.Tag) & ","
            End If
            LV.ListItems.Remove LV.SelectedItem.Index
        End If
    End If
    
End Sub

Private Sub NBox_Change(Index As Integer)
    Select Case Index
        Case 4, 5, 7
            'NBox(6) = Val(NBox(4)) - Val(NBox(5)) - Val(NBox(7))
    End Select
    
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


Private Sub Tbox_Change(Index As Integer)
    Select Case Index
    
    Case 8, 3
        'TBox(4) = TBox(8) & TBox(3)
    End Select
    
End Sub

Private Sub Tbox_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    Exit Sub
    Select Case Index
    Case 2, 3
        KeyAscii = OnlyNumber(Val(KeyAscii))
    Case 5
        KeyAscii = OnlyNumber(Val(KeyAscii), True)
    End Select
    
End Sub

Private Sub TBox_Validate(Index As Integer, cancel As Boolean)
    Exit Sub
    If Index = 5 Then
        TBox(Index) = Format(Val(TBox(Index)), "#0.00")
    ElseIf Index = 2 Or Index = 3 Then
        TBox(Index) = Format(Val(TBox(Index)), "#0")
    End If
End Sub

Private Sub txtHubName_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cmbProcess.MatchFound = False Then
            MsgBox "Why not Select Process.", vbInformation
            Exit Sub
        ElseIf cmbScanning.MatchFound = False Then
            MsgBox "I'm Sorry, But you need to Select Scanning.", vbInformation
            Exit Sub
        End If
        Dim iProcessID As Long
        iProcessID = Val(cmbProcess.ID)
        Dim i As Integer
        For i = 1 To LV.ListItems.count
            If Val(LV.ListItems(i).key) = iProcessID Then
                MsgBox "This Process Already Added in this Group.", vbInformation
                Exit Sub
            End If
        Next
        Dim ITM As ListItem
        Set ITM = LV.ListItems.add(, iProcessID & "'", cmbProcess.Text)
        ITM.Tag = "0" 'cmbProcess.List(cmbProcess.ListIndex, 1)
        ITM.ListSubItems.add(, , cmbScanning.Text).Tag = cmbScanning.ListIndex
        
        Dim iSortNo As Integer
        If LV.ListItems.count = 0 Then
            iSortNo = 1
        Else
            iSortNo = LV.ListItems.count
        End If
        ITM.SubItems(2) = txtHubName
        ITM.SubItems(3) = iSortNo
        cmbProcess.SetFocus
        
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub txtSortNo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        
        If iColNo = 4 Then
            If Val(txtSortNo) > LV.ListItems.count Then
                MsgBox "Invalid Sort No.", vbInformation
                Exit Sub
            End If
            Dim bPlus As Boolean
            If Val(txtSortNo) < Val(LV.SelectedItem.SubItems(3)) Then
                bPlus = True
            Else
                bPlus = False
            End If
            txtSortNo.Visible = False
            LV.SelectedItem.SubItems(3) = Val(txtSortNo)
                    
            Dim ITM As ListItem, myITM As ListItem
            Set ITM = LV.SelectedItem
            Dim iNewIndex As Integer
            If bPlus Then
                iNewIndex = Val(txtSortNo)
            Else
                iNewIndex = Val(txtSortNo) + 1
            End If
            Set myITM = LV.ListItems.add(iNewIndex, , ITM.Text)
            myITM.Checked = True
            myITM.Tag = ITM.Tag
            myITM.ListSubItems.add(, , ITM.SubItems(1)).Tag = ITM.ListSubItems(1).Tag
            myITM.ListSubItems.add(, , ITM.SubItems(2)).Tag = ITM.ListSubItems(2).Tag
            myITM.ListSubItems.add(, , ITM.SubItems(3)).Tag = ITM.ListSubItems(3).Tag
            
            LV.ListItems.Remove LV.SelectedItem.Index
            myITM.key = ITM.key
            
            Call ReIndexProcesses
    '        Dim i As Integer
    '        For i = 1 To LV.ListItems.count
    '            With LV.ListItems(i)
    '                con.Execute "UPDATE RepairTypeProcesses SET SeqNo=" & Val(.SubItems(2)) & " WHERE EntryID=" & Val(.ListSubItems(1).Tag)
    '            End With
    '        Next
        ElseIf iColNo = 3 Then
            LV.SelectedItem.SubItems(iColNo - 1) = txtSortNo
            txtSortNo.Visible = False
        End If
    ElseIf KeyAscii = 27 Then
        txtSortNo.Visible = False
    End If
    
End Sub

Private Sub ReIndexProcesses()

    'First Move All to LVHidden
    LV.Visible = False
    Dim i As Integer
    Dim FromITM As ListItem, ToITM As ListItem
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).SubItems(3) = i
        'Set FromITM = LV.ListItems(i)
        'set toitm=LVhidden.ListItems.Add(
    Next
    LV.Visible = True
End Sub


