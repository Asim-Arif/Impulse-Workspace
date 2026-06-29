VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "combolist.ocx"
Begin VB.Form frmRepairTypes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Repair Types"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9285
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraAdd 
      Height          =   630
      Left            =   2850
      TabIndex        =   3
      Top             =   7530
      Width           =   6315
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   345
         Left            =   3165
         TabIndex        =   17
         Top             =   180
         Width           =   1485
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   345
         Left            =   1635
         TabIndex        =   6
         Top             =   180
         Width           =   1485
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   345
         Left            =   4695
         TabIndex        =   5
         Top             =   180
         Width           =   1485
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   345
         Left            =   90
         TabIndex        =   4
         Top             =   180
         Width           =   1485
      End
   End
   Begin ComboList.Usercontrol1 cmbRepairs 
      Height          =   285
      Left            =   60
      TabIndex        =   1
      Top             =   300
      Width           =   9090
      _ExtentX        =   16034
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
   Begin VB.TextBox txtRepairType 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   60
      TabIndex        =   10
      Top             =   300
      Visible         =   0   'False
      Width           =   6555
   End
   Begin VB.TextBox txtDefect 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7860
      TabIndex        =   12
      Top             =   300
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.TextBox txtRepairCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      TabIndex        =   11
      Top             =   300
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.TextBox txtSortNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4050
      TabIndex        =   14
      Top             =   2355
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.CheckBox chkReturnRepairToOriginator 
      Caption         =   "Return to Originator"
      Height          =   285
      Left            =   7035
      TabIndex        =   13
      Top             =   585
      Width           =   2115
   End
   Begin VB.Frame FraSave 
      Height          =   630
      Left            =   5925
      TabIndex        =   7
      Top             =   7530
      Visible         =   0   'False
      Width           =   3240
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   345
         Left            =   90
         TabIndex        =   9
         Top             =   180
         Width           =   1485
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   345
         Left            =   1620
         TabIndex        =   8
         Top             =   180
         Width           =   1485
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   6615
      Left            =   75
      TabIndex        =   2
      Top             =   885
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   11668
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Process"
         Object.Width           =   11465
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Code"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "S No."
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Defect"
      Height          =   195
      Index           =   2
      Left            =   7890
      TabIndex        =   16
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Code"
      Height          =   195
      Index           =   1
      Left            =   6645
      TabIndex        =   15
      Top             =   90
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Repair Type"
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   75
      Width           =   870
   End
End
Attribute VB_Name = "frmRepairTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbRepairs_matched()
    
    Dim i As Integer, ITM As ListItem
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = False
    Next
    
    Dim lForeColor As Long
    Dim rs As New ADODB.Recordset
    With rs
        '.Open "SELECT * FROM RepairTypeProcesses WHERE Repair_RefID=" & cmbRepairs.ID, con, adOpenForwardOnly, adLockReadOnly
        .Open "SELECT * FROM VRepairTypeProcesses WHERE EntryID=" & cmbRepairs.ID & " ORDER BY SeqNo", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, !ProcessID & "'", !Description & "")
            
            ITM.ListSubItems.add(, , !code & "").Tag = Val(!RTP_EntryID & "")
            ITM.ListSubItems.add(, , Val(!SeqNo & "")).Tag = !Scanning
            
            If !Scanning Then
                lForeColor = vbRed
            Else
                lForeColor = LV.ForeColor
            End If
            
            ITM.ForeColor = lForeColor
            ITM.ListSubItems(1).ForeColor = lForeColor
            ITM.ListSubItems(2).ForeColor = lForeColor
            
            ITM.Checked = True
            .MoveNEXT
        Loop
        .Close
        .Open "SELECT ReturnToOriginator FROM RepairTypes WHERE EntryID=" & cmbRepairs.ID, con, adOpenForwardOnly, adLockReadOnly
        chkReturnRepairToOriginator.Value = Abs(!ReturnToOriginator)
        .Close
'        Do Until .EOF
'            LV.ListItems(!ProcessID & "'").Checked = True
'            .MoveNEXT
'        Loop
        
    End With
    Set rs = Nothing
    
    
    
'    For i = 1 To LV.ListItems.count
'        If LV.ListItems(i).Checked = False Then
'            LV.ListItems(i).Ghosted = True
'        End If
'    Next
'    LV.Refresh
    
End Sub

Private Sub cmdAdd_Click()
    Call FillProcesses
    Call ChangeState(True)
End Sub

Private Sub ChangeState(Optional bAdd As Boolean = False, Optional bEdit As Boolean = False)

    Dim bAddEdit As Boolean
    bAddEdit = bAdd Or bEdit
    
    FraAdd.Visible = Not bAddEdit
    
    Label1(1).Visible = bAddEdit
    Label1(2).Visible = bAddEdit
    txtRepairCode.Visible = bAddEdit
    txtDefect.Visible = bAddEdit
    
    FraSave.Visible = bAddEdit
    cmbRepairs.Visible = Not bAddEdit
    
    'chkReturnRepairToOriginator.Visible = bAddEdit
    
    txtRepairType.Visible = bAddEdit
    
    Dim i As Integer
    If bAdd Then
        For i = 1 To LV.ListItems.count
            LV.ListItems(i).Checked = False
        Next
        txtRepairType = ""
        txtRepairType.Tag = ""
        txtRepairCode = ""
        txtDefect = ""
    ElseIf bEdit Then
        'txtRepairType = cmbRepairs.Text
        txtRepairType.Tag = cmbRepairs.ID
    End If
    
    
'    For i = 1 To LV.ListItems.count
'        If LV.ListItems(i).Checked = False Then
'            LV.ListItems(i).Ghosted = Not bAddEdit
'        End If
'    Next
    
End Sub

Private Sub cmdCancel_Click()
    If cmbRepairs.MatchFound Then
        Call cmbRepairs_matched
    End If
    Call ChangeState
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    
    If cmbRepairs.MatchFound = False Then Exit Sub
    Call FillProcesses
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM RepairTypeProcesses WHERE Repair_RefID=" & cmbRepairs.ID, con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            LV.ListItems(!ProcessID & "'").Checked = True
            LV.ListItems(!ProcessID & "'").SubItems(2) = Val(!SeqNo & "")
            .MoveNEXT
        Loop
        .Close
        .Open "SELECT RepairType,RepairCode,DefectType FROM RepairTypes WHERE EntryID=" & cmbRepairs.ID, con, adOpenForwardOnly, adLockReadOnly
        txtRepairType = !RepairType & ""
        txtRepairCode = !RepairCode & ""
        txtDefect = !DefectType & ""
        .Close
    End With
    Set rs = Nothing
    
    Call ChangeState(False, True)
    
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
    
    Dim rpt As craxddrt.Report, strSelection As String
    Set rpt = rptApp.OpenReport(RptPath & "\RepairTypes.rpt")
    Dim f As New frmPrevRpt
    If cmbRepairs.MatchFound Then
        strSelection = "{RepairTypes.EntryID}=" & cmbRepairs.ID
    End If
    Me.Hide
    f.ShowReport strSelection, rpt, , , , Me
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdSave_Click()

    On Error GoTo err
        
    If txtRepairType = "" Then
        MsgBox "Please Enter Repair Type.", vbInformation
        Exit Sub
    ElseIf txtRepairCode = "" Then
        MsgBox "Please Enter Repair Code.", vbInformation
        Exit Sub
    ElseIf txtDefect = "" Then
        MsgBox "Please Enter Defect.", vbInformation
        Exit Sub
    End If
    
    Dim lEntryID As Long
    Call StartTrans(con)
    If Val(txtRepairType.Tag) > 0 Then
        con.Execute "UPDATE RepairTypes SET RepairType='" & txtRepairType & "',ReturnToOriginator=" & chkReturnRepairToOriginator & ",RepairCode='" & txtRepairCode & "',DefectType='" & txtDefect & "' WHERE EntryID=" & Val(txtRepairType.Tag)
        lEntryID = Val(txtRepairType.Tag)
    Else
        con.Execute "INSERT INTO RepairTypes(RepairType,ReturnToOriginator,RepairCode,DefectType) VALUES('" & txtRepairType & "'," & chkReturnRepairToOriginator & ",'" & txtRepairCode & "','" & txtDefect & "')"
        lEntryID = GetSingleLongValue("MAX(EntryID)", "RepairTypes")
    End If
    
    Dim iSelCount As Long, i As Integer
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            iSelCount = iSelCount + 1
        End If
    Next
    
    Dim lmySelProcesses() As Long, iIndex As Integer
    ReDim lmySelProcesses(iSelCount)
    iIndex = 0
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            lmySelProcesses(iIndex) = Val(LV.ListItems(i).key)
            iIndex = iIndex + 1
        End If
    Next
    
    con.Execute "DELETE FROM RepairTypeProcesses WHERE Repair_RefID=" & lEntryID
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            con.Execute "INSERT INTO RepairTypeProcesses(Repair_RefID,ProcessID,SeqNo) VALUES(" & _
             lEntryID & "," & Val(LV.ListItems(i).key) & "," & Val(LV.ListItems(i).SubItems(2)) & ")"
        End If
    Next
    
    con.CommitTrans
        
    Call FillRepairTypes
    cmbRepairs.ID = lEntryID
    
    Call ChangeState
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    
    Call FillProcesses
    
    Call FillRepairTypes
    
End Sub

Private Sub FillProcesses()
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        .Open "SELECT * FROM Processes", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, !ProcessID & "'", !Description & "")
            ITM.ListSubItems.add , , !code & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
End Sub

Private Sub FillRepairTypes()

    cmbRepairs.ClearVals
    cmbRepairs.AddVals con, "'{' + ISNULL(RepairCode,'') + '} ' + RepairType + ' (' + ISNULL(DefectType,'') + ')'", "RepairTypes", "EntryID", " ORDER BY RepairType"
        
End Sub

Private Sub LV_DblClick()

    With txtSortNo
        .Move LV.Left + LV.ColumnHeaders(3).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(3).Width
        .Text = Val(LV.SelectedItem.SubItems(2))
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
    
End Sub

Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    Dim lLastSeqNo As Long
    lLastSeqNo = 0
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If Val(.SubItems(2)) > lLastSeqNo Then
                lLastSeqNo = Val(.SubItems(2))
            End If
        End With
    Next
    Item.SubItems(2) = lLastSeqNo + 1
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If LV.ListItems.count = 0 Then Exit Sub
    If Button <> 2 Then Exit Sub
    If cmbRepairs.MatchFound = False Then Exit Sub
    
    Dim strRet As String
    If CBool(LV.SelectedItem.ListSubItems(2).Tag) Then  'Scanning
        strRet = ShowPopUpMenu1("Remove Scanning")
    Else
        strRet = ShowPopUpMenu1("Add Scanning")
    End If
    
    If strRet = "Add Scanning" Then
         con.Execute "UPDATE RepairTypeProcesses SET Scanning=1 WHERE EntryID=" & Val(LV.SelectedItem.ListSubItems(1).Tag)
         Call cmbRepairs_matched
    ElseIf strRet = "Remove Scanning" Then
        con.Execute "UPDATE RepairTypeProcesses SET Scanning=0 WHERE EntryID=" & Val(LV.SelectedItem.ListSubItems(1).Tag)
        Call cmbRepairs_matched
    End If
    
End Sub

Private Sub txtSortNo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        
        If Val(txtSortNo) > LV.ListItems.count Then
            MsgBox "Invalid Sort No.", vbInformation
            Exit Sub
        End If
        Dim bPlus As Boolean
        If Val(txtSortNo) < LV.SelectedItem.SubItems(2) Then
            bPlus = True
        Else
            bPlus = False
        End If
        txtSortNo.Visible = False
        LV.SelectedItem.SubItems(2) = Val(txtSortNo)
                
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

        LV.ListItems.Remove LV.SelectedItem.Index
        myITM.key = ITM.key
        
        Call ReIndexProcesses
        Dim i As Integer
        For i = 1 To LV.ListItems.count
            With LV.ListItems(i)
                con.Execute "UPDATE RepairTypeProcesses SET SeqNo=" & Val(.SubItems(2)) & " WHERE EntryID=" & Val(.ListSubItems(1).Tag)
            End With
        Next
        
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
        LV.ListItems(i).SubItems(2) = i
        'Set FromITM = LV.ListItems(i)
        'set toitm=LVhidden.ListItems.Add(
    Next
    LV.Visible = True
End Sub

