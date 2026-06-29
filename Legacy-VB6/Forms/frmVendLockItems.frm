VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmVendLockItems 
   ClientHeight    =   10125
   ClientLeft      =   1110
   ClientTop       =   60
   ClientWidth     =   11805
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
   ScaleHeight     =   10125
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check1 
      Caption         =   "Show Locked Items"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3345
      TabIndex        =   12
      Top             =   795
      Width           =   2160
   End
   Begin ComboList.Usercontrol1 cmbProcess 
      Height          =   285
      Left            =   795
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   810
      Width           =   2355
      _ExtentX        =   4154
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
   Begin VB.Frame FraLoad 
      Height          =   8250
      Left            =   420
      TabIndex        =   10
      Top             =   1140
      Width           =   10860
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Loading Data"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2295
         Left            =   90
         TabIndex        =   11
         Top             =   3090
         Width           =   10605
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5745
      TabIndex        =   6
      Top             =   9495
      Width           =   5280
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3480
         TabIndex        =   9
         Top             =   165
         Width           =   1665
      End
      Begin VB.CommandButton cmdSaveClose 
         Caption         =   "&Save && Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1770
         TabIndex        =   8
         Top             =   165
         Width           =   1665
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save && &New   "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   60
         TabIndex        =   7
         Top             =   165
         Width           =   1665
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   555
      Top             =   -195
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
            Picture         =   "frmVendLockItems.frx":0000
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVendLockItems.frx":01BC
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -15
      Top             =   -195
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
            Picture         =   "frmVendLockItems.frx":0378
            Key             =   "Active"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVendLockItems.frx":07CC
            Key             =   "Inactive"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV 
      Height          =   8190
      Left            =   420
      TabIndex        =   4
      Top             =   1140
      Visible         =   0   'False
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   14446
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
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
         Text            =   "Item Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   11827
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Issueable Qty"
         Object.Width           =   2822
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
      Left            =   795
      TabIndex        =   0
      Top             =   0
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lock Items"
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
         Left            =   4440
         TabIndex        =   1
         Top             =   120
         Width           =   1710
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lock Items"
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
         Left            =   4455
         TabIndex        =   2
         Top             =   135
         Width           =   1710
      End
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   795
      TabIndex        =   3
      Top             =   585
      Width           =   2340
      BackColor       =   11517387
      Caption         =   "  Process"
      Size            =   "4128;423"
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
Attribute VB_Name = "frmVendLockItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyCol As New Collection
Dim MyCol1 As New Collection
Dim IsItemsFirst As Boolean
Dim IssIDs As New Collection
Dim rsMain As New ADODB.Recordset
Dim IsLotProcess As Boolean
Dim GotIssID As Long
Dim RcvingLot As Boolean
Dim UserReqAuth As Integer 'Whether Or Not User Required Authorization On Issue For This Process.

Dim lVendBal As Long
Private Sub AddItems()
    
    On Error GoTo err
    Exit Sub
    Dim OrCond As String
    'Create In List
    
    'For i = 1 To MyCol.Count
    '    OrCond = OrCond & "'" & MyCol(i).CompItemCode & "',"
    'Next i
    
    If Len(InList) > 0 Then
        InList = Mid$(InList, 1, Len(InList) - 1)
    End If
    'InList = InList & ")"
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select * From VVendAssItems", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            For i = 1 To MyCol.Count
                If MyCol(i).CompItemCode = ![ITemID] Then
                    cmbItems.AddItem ![ITemID] 'MyCol(i).ItemCode
                    cmbItems.List(cmbItems.ListCount - 1, 1) = MyCol(i).ItemCode
                    cmbItems.List(cmbItems.ListCount - 1, 2) = MyCol(i).lQty
                    cmbItems.List(cmbItems.ListCount - 1, 3) = ![Rate]
                    cmbItemDesc.AddItem MyCol(i).ItemDescription
                    Exit For
                End If
            Next i
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub cmbItem_matched()
    cmbItemName.ID = cmbItem.ID
End Sub

Private Sub cmbItemName_matched()
    cmbItem.ID = cmbItemName.ID
End Sub

Private Sub cmbDesc_Change()
    Call cmbDesc_Click
End Sub

Private Sub cmbDesc_Click()

    If cmbDesc.MatchFound = False Then Exit Sub
    cmbItems.ListIndex = cmbDesc.ListIndex
    
    
    If RcvingLot Then
        LV.ListItems(cmbDesc.ListIndex + 1).Selected = True
    End If
    
End Sub

Private Sub FillList()

    On Error GoTo err
    
    LV.ListItems.Clear
    For i = 1 To MyCol.Count
        MyCol.Remove 1
    Next i
    
    Dim ITM As ListItem
    
    With rsMain
        .Filter = adFilterNone
        If RcvingLot Then
            .Filter = "RecieptID='" & cmbItems.Text & "'"
        Else
            .Filter = "ItemCode='" & cmbItems.Text & "'"
        End If
        LV.ListItems.Clear
            
        Do Until .EOF
            If RcvingLot Then
                Set ITM = LV.ListItems.Add(, ![LotNo] & "'", LV.ListItems.Count + 1)
                ITM.Tag = ![RecieptID] & ""
            Else
                Set ITM = LV.ListItems.Add(, ![OrderNo] & "|" & ![CompItemCode], LV.ListItems.Count + 1)
            End If
            
            'ITM.Tag = ![EntryID]
            ITM.ListSubItems.Add , , ![CustCode] & ""
            If RcvingLot Then
                ITM.ListSubItems.Add(, , ![LotNo] & "/" & ![OrderNo] & "").Tag = Val(![LotNo] & "")
            Else
                ITM.ListSubItems.Add , , ![OrderNo] & ""
            End If
            
            ITM.ListSubItems.Add , , ![ItemCode] & ""
            ITM.ListSubItems.Add , , ![ItemDescription] & ""
            ITM.ListSubItems.Add(, , Val(![IssQty] & "")).Tag = Val(![IssQty] & "")
            
            Dim Obj As New IssueItems
            
            Obj.LVIndex = LV.ListItems.Count + 1
                
            Obj.OrderQty = Val(![Qty] & "")
            Obj.FreeStock = Val(![FreeStock] & "")
            Obj.IssueableQty = Val(![IssQty] & "")
            If RcvingLot Then
                MyCol.Add Obj, ![LotNo] & "'"
            Else
                MyCol.Add Obj, ![OrderNo] & "|" & ![CompItemCode]
            End If
            Set Obj = Nothing
            
            .MoveNext
        Loop
        '.Close
    End With
    
    'Set rs = Nothing
    
    'Call AddItems
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Check1_Click()

    If Check1.value = vbChecked Then
        Dim rs As New ADODB.Recordset
        Dim ITM As ListItem
        With rs
            .Open "Select ItemCode From LockedItems Where ProcessID=" & cmbProcess.ID, con, adOpenForwardOnly, adLockReadOnly
            LV.ListItems.Clear
            Do Until .EOF
                Set ITM = LV.ListItems.Add(, , LV.ListItems.Count + 1)
                ITM.SubItems(1) = !ItemCode & ""
                ITM.Checked = True
                .MoveNext
            Loop
            .Close
        End With
        Set rs = Nothing
        FraLoad.Visible = False
        LV.Visible = True
    Else
        Call cmbProcess_matched
    End If
    
End Sub

Private Sub cmbProcess_matched()

    On Error GoTo err
    If Check1.value = vbChecked Then Exit Sub
    If cmbProcess.MatchFound = False Then Exit Sub
    Dim cmd As New ADODB.Command
    
    LV.ListItems.Clear

    IsLotProcess = LotProcess(cmbProcess.ID)
    
    
    With cmd
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        
        
        If IsLotProcess Then

            If RcvingLot Then
            
                'LV.ColumnHeaders(2).Text = "Lot No/Order No"
                .CommandText = "SP_IssueItemsLots"
            Else
            
                .CommandText = "SP_IssueItems"
            End If
            cmbProcess.Tag = "Lot No"
        Else
            
            .CommandText = "SP_IssueItems"
            cmbProcess.Tag = ""
        End If
        
        If rsMain.State = 1 Then rsMain.Close
        rsMain.CursorType = adOpenForwardOnly
        rsMain.LockType = adLockReadOnly
        
        .Parameters("@ProcessID").value = cmbProcess.ID
            
        Set rsMain = .Execute
    End With
    
    Dim isAdded As Boolean
    
    With rsMain
        For j = 0 To 1
            If j = 0 Then .Filter = "LockedItemCode=Null" Else .Filter = "LockedItemCode<>Null"
            If RcvingLot Then
    '            DTPicker1.value = ServerDate
                Do Until .EOF
                    isAdded = False
                    For i = 0 To cmbItems.ListCount - 1
                        If Trim(cmbItems.List(i)) = Trim(![RecieptID] & "") Then
                            isAdded = True
                            Exit For
                        End If
                    Next i
                    
                    If Not isAdded Then
                        cmbItems.AddItem ![RecieptID] & ""
                        cmbDesc.AddItem ![RecieptID] & ""
                    End If
                        
                    
                    Set ITM = LV.ListItems.Add(, ![LotNo] & "'", LV.ListItems.Count + 1)
                    ITM.Tag = ![RecieptID] & ""
                    
                    'ITM.Tag = ![EntryID]
                    'itm.ListSubItems.Add , , ![CustCode] & ""
                    ITM.ListSubItems.Add(, , ![RecieptID] & "").Tag = ![LotNo] & ""
                    
                    ITM.ListSubItems.Add(, , ![ItemCode] & "").Tag = ![ItemCode] & ""
                    
                    'itm.ListSubItems.Add , , ![ItemCode] & ""
                    ITM.ListSubItems.Add , , ![ItemDescription] & ""
                    ITM.ListSubItems.Add(, , Val(![IssQty] & "")).Tag = Val(![IssQty] & "")
                    If j = 1 Then ITM.Checked = True
                    .MoveNext
                Loop
            Else
                Do Until .EOF
                    isAdded = False
                    
                
                    Set ITM = LV.ListItems.Add(, , LV.ListItems.Count + 1)
                
                
                    'ITM.Tag = ![EntryID]
                    
                    ITM.ListSubItems.Add , , ![ItemCode] & ""
                    ITM.ListSubItems.Add , , ![ItemDescription] & ""
                    ITM.ListSubItems.Add(, , Val(![IssQty] & "")).Tag = Val(![IssQty] & "")
                    If j = 1 Then ITM.Checked = True
                    .MoveNext
                Loop
            End If
        Next
        
    End With
    
    
    Me.Show

    SendKeys "a"
    SendKeys Chr(8)
    
    FraLoad.Visible = False
    LV.Visible = True
    'Debug.Print Time
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbVender_GotFocus()

    cmdPost.Default = False
    SendKeys "a"
    SendKeys Chr(8)
    
End Sub

Private Sub cmbVender_matched()

    On Error GoTo err
    If cmbVender.Text = "HICO" Then Exit Sub
    If RcvingLot Then
        cmbVender.Tag = Val(con.Execute("Select Rate From VVendAssItems Where VendID=" & cmbVender.ID & " and ItemID='" & LV.SelectedItem.SubItems(2) & "'").Fields(0).value & "")
    Else
        cmbVender.Tag = Val(con.Execute("Select Rate From VVendAssItems Where VendID=" & cmbVender.ID & " and ItemID='" & LV.SelectedItem.SubItems(1) & "'").Fields(0).value & "")
    End If
    Call ShowVenderHistory
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub ShowVenderHistory()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "Select Top 10 * From VVendIssItems Where VendID=" & cmbVender.ID & " and ProcessID=" & cmbProcess.ID & " Order By DT Desc,EntryID Desc", con, adOpenForwardOnly, adLockReadOnly
        
        LVMaker.ListItems.Clear
        LVBal.ListItems.Clear
        lVendBal = 0
        Do Until .EOF
            Set ITM = LVMaker.ListItems.Add(, , LVMaker.ListItems.Count + 1)
            ITM.SubItems(1) = ![ItemCode] & ""
            ITM.SubItems(2) = ![ItemDescription] & ""
            ITM.SubItems(3) = Format(![DT] & "", "dd-MMM-yyyy")
            ITM.SubItems(4) = Val(![IssQty] & "")
            .MoveNext
        Loop
        .Close
        
        If cmbVender.ID <> 129 Then
            .Open "Select ItemCode,ItemDescription,Sum(RemQty) As RemQty From VVendIssItems Where VendID=" & cmbVender.ID & " And ProcessID=" & cmbProcess.ID & " AND RemQty>0 Group By ItemCode,ItemDescription Having Sum(RemQty)>0", con, adOpenForwardOnly, adLockReadOnly
            
            Do Until .EOF
                Set ITM = LVBal.ListItems.Add(, , LVBal.ListItems.Count + 1)
                ITM.SubItems(1) = ![ItemCode] & ""
                ITM.SubItems(2) = ![ItemDescription] & ""
                ITM.SubItems(3) = Val(![RemQty] & "")
                lVendBal = lVendBal + Val(![RemQty] & "")
                .MoveNext
            Loop
            .Close
            txtTotal = lVendBal
        End If
        
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbVender_Validate(Cancel As Boolean)
    cmdPost.Default = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    'con.Execute "Delete From PrintSession"
End Sub

Private Sub cmdPost_Click()

    On Error GoTo err
    
    If Not cmbVender.MatchFound Then Exit Sub
    If cmbVender.ID = "" Or cmbVender.Text = "" Then Exit Sub
    'cmbVender.Enabled = False
    
    
    'Check For Max Balance Limit For This Vender.......
    If cmbVender.ID <> 129 Then
        Dim rs As New ADODB.Recordset
        With rs
            .Open "Select MaxLimit From Venders Where VendID=" & cmbVender.ID, con, adOpenForwardOnly, adLockReadOnly
            If lVendBal >= Val(![MaxLimit] & "") Then
                MsgBox "Sorry, The Selected Vender Has Reached His Max. Issuence Limit." & _
                 vbNewLine & "Can't Post It For This Vender." & vbNewLine & "Vender Maximum Limit : " & Val(![MaxLimit] & "") & vbNewLine & "Vender Curr. Balance  : " & lVendBal, vbInformation
                
                .Close
                Exit Sub
            End If
            .Close
        End With
        Set rs = Nothing
    End If
    
    If RcvingLot Then
        PostForLots
        Exit Sub
    End If
    
    If Val(txtQty) < 1 Then Exit Sub
    
    If Val(txtQty) > Val(LV.SelectedItem.SubItems(3)) Then
        MsgBox "Invalid Qty", vbInformation
        Exit Sub
    End If
    
    If GotIssID = 0 Then
        'Dim IssID As String
        'For i = 1 To LVIssued.ListItems.Count
        '    If Val(cmbVender.ID) = Val(LVIssued.ListItems(i).ListSubItems(2).Key) Then
        '        IssID = LVIssued.ListItems(i).ListSubItems(2)
        '        GoTo GotID:
        '    End If
        'Next i
    
        GotIssID = Val(con.Execute("Select Max(Cast(right(RecieptID,Len(RecieptID)-10) As Int)) From VVendIssItems Where Convert(varchar,DT,6)='" & Format(DTPicker2, "dd MMM yy") & "' and Left(RecieptID,3)<>'HSR'").Fields(0).value & "")
        GotIssID = GotIssID + 1
    Else
        GotIssID = GotIssID + 1
    End If
    'IssID = "ISS-" & Val(LVMaker.SelectedItem.ListSubItems(1).Key) & "-" & _
     Format(ServerDate, "ddMMyy") & "-" & Format(Val(IssID) + 1, "000")
    IssID = "ISU-" & _
     Format(DTPicker2, "ddMMyy") & GotIssID
    
GotID:
        
    Dim ITM As ListItem
    With ITM
        'Set ITM = LVIssued.ListItems.Add(, Val(LV.SelectedItem.Key) & Val(LVMaker.SelectedItem.Key) & "'", LVIssued.ListItems.Count + 1)
        'Set ITM = LVIssued.ListItems.Add(, LV.SelectedItem.Key & "|" & LVIssued.ListItems.Count + 1, LVIssued.ListItems.Count + 1)
        Set ITM = LVIssued.ListItems.Add(, , LVIssued.ListItems.Count + 1)
        ITM.Tag = LV.SelectedItem.Key
        'itm.ListSubItems.Add , cmbVender.ID & "'", cmbVender.Text
        ITM.ListSubItems.Add(, , cmbVender.Text).Tag = cmbVender.ID
        ITM.ListSubItems.Add(, , IssID & "").Tag = DTPicker2
        ITM.ListSubItems.Add , , LV.SelectedItem.SubItems(1)
        ITM.ListSubItems.Add , , LV.SelectedItem.SubItems(2)
        ITM.ListSubItems.Add , , Val(cmbVender.Tag)
        ITM.ListSubItems.Add , , Val(txtQty) 'LV.SelectedItem.SubItems(3)
        'ITM.ListSubItems.Add , , Val(txtQty)
        If cmbProcess.Tag = "Lot No" Then
            ITM.ListSubItems(6).Tag = LV.SelectedItem.SubItems(2) & ""
        End If
        
    End With
    
    'MyCol(LV.SelectedItem.Index).IssueableQty = MyCol(LV.SelectedItem.Index).IssueableQty - Val(txtQty)
    LV.SelectedItem.SubItems(3) = Val(LV.SelectedItem.SubItems(3)) - Val(txtQty)
    
    
    txtRate.Visible = False
    
    'Removing Item From Combos
    If LV.SelectedItem.SubItems(3) = 0 Then
        Dim ItemIndex As Long
        ItemIndex = cmbItems.ListIndex
        If ItemIndex <> -1 Then
            
            cmbItems.RemoveItem ItemIndex
            cmbDesc.RemoveItem ItemIndex
            cmbItems.Text = ""
            cmbDesc.Text = ""
            If cmbItems.ListCount >= ItemIndex And ItemIndex > 0 Then
                cmbItems.ListIndex = ItemIndex
            ElseIf cmbItems.ListCount > 0 Then
                cmbItems.ListIndex = 0
            End If
        End If
    End If
    'if cmbitems.
    'Call LV_ItemClick(LV.SelectedItem)
    
    
    txtQty = ""
    
    'Select The Next Item Automaticaly....
    'cmbVender.SetFocus
    
    Exit Sub
err:
    MsgBox err.Description
        'End If
End Sub

Private Sub cmdSave_Click()

    If Saved Then
        'FormProcessID = cmbProcess.ID
        Unload Me
        'Me.Show
        'cmbProcess.ID = cmbProcess.Tag
        Me.Show
    End If
    
End Sub

Private Sub cmdSaveClose_Click()
    
    If Saved Then
        'If Not UserReqAuth Then Call PrintIssRcvRpts("Original", True)
        Unload Me
        'con.Execute "Delete From PrintSession"
    End If
End Sub

Private Sub cmdUnAssign_Click()

    On Error GoTo err
    If LV.ListItems.Count = 0 Then Exit Sub
    Dim iAffect As Integer
    
    con.Execute "Delete From VendAssItems Where EntryID=" & Val(LV.SelectedItem.Key), iAffect
    If iAffect > 0 Then
        'Call cmbNick_Click
        cmdUnAssign.Visible = False
    End If
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub Command1_Click()

End Sub

Private Sub DTPicker1_Change()
    Call DTPicker1_Click
End Sub

Private Sub DTPicker1_Click()
    On Error GoTo err
    
    
    With rsMain
    
        If DTPicker1.value & "" = "" Then
            .Filter = adFilterNone
        Else
            .Filter = "RecieptID Like 'HS-" & Format(DTPicker1, "ddMMyy") & "-*'"
            '.Filter = "Left(RecieptID,9) ='HS-" & Format(DTPicker1, "ddMMyy") & "'"
        End If
                
        Dim isAdded As Boolean
        
        cmbItems.Clear
        cmbDesc.Clear
        LV.ListItems.Clear
        Do Until .EOF
            isAdded = False
            For i = 0 To cmbItems.ListCount - 1
                If Trim(cmbItems.List(i)) = Trim(![RecieptID] & "") Then
                    isAdded = True
                    Exit For
                End If
            Next i
            
            If Not isAdded Then
                cmbItems.AddItem ![RecieptID] & ""
                cmbDesc.AddItem ![RecieptID] & ""
            End If
            .MoveNext
        Loop
    End With
    
    
        
    
    
    Exit Sub
err:
    MsgBox err.Description
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


Private Sub LVIssued_DblClick()
    
    
    If LVIssued.ListItems.Count = 0 Then Exit Sub
    'If cmbCust.MatchFound = False Then Exit Sub
    If LV.ColumnHeaders(3).Text = "Lot No/Order No" And RcvingLot Then Exit Sub
    
    
    With txtRate1
        .Move LVIssued.Left + LVIssued.ColumnHeaders(8).Left, LVIssued.Top + LVIssued.SelectedItem.Top, LVIssued.ColumnHeaders(8).Width
        .Text = LVIssued.SelectedItem.SubItems(7) 'LVMaker.SelectedItem.SubItems(3)
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
    
    On Error GoTo err
    
    'cmbCust.ListHeight = 2600
    'cmbCustName.ListHeight = 2600
    'cmbOrders.ListHeight = 2600
    
    'cmbProcess.Enabled = True
    
    
    
    cmbProcess.ListHeight = 2600
    cmbProcess.AddVals con, "Description", "Processes", "ProcessID", " Where ProcessID<>1 Order By SNo"



    IsLotProcess = False

    GotIssID = 0
    
    Me.Hide
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Public Sub LockProc(ProcID As Long)
    
    
    cmbProcess.Tag = ProcID
    
    'cmbProcess.Enabled = False
    'Me.Refresh
    
    cmbProcess.ID = ProcID
    'Me.Show
    
End Sub
Public Sub RefreshLV()
    
    On Error GoTo err
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
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub Form_Resize()

    On Error Resume Next
    
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
    
    Dim TitleBottom As Long
    TitleBottom = cmbProcess.Top + cmbProcess.Height + 50 '+ fTitle.Height
    LV.Move Label2.Left, TitleBottom
    
End Sub




Private Sub txtQty_GotFocus()


    With txtQty
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    

End Sub


Private Function Saved() As Boolean

    On Error GoTo err
    Call StartTrans(con)
    For i = 1 To LV.ListItems.Count
        con.Execute "Delete From LockedItems Where ItemCode='" & LV.ListItems(i).SubItems(1) & "' AND ProcessID=" & cmbProcess.ID
        If LV.ListItems(i).Checked Then
            con.Execute "Insert Into LockedItems(ItemCode,ProcessID) Values('" & LV.ListItems(i).SubItems(1) & "'," & cmbProcess.ID & ")"
        Else
            con.Execute "Delete From LockedItems Where ItemCode='" & LV.ListItems(i).SubItems(1) & "' AND ProcessID=" & cmbProcess.ID
        End If
    Next
    
    con.CommitTrans
    
    Saved = True
    Exit Function
err:
    MsgBox err.Description
    
End Function

Private Sub txtRate1_LostFocus()
    txtRate1.Visible = False
End Sub
Private Sub txtRate_LostFocus()
    txtRate.Visible = False
End Sub
Private Sub PostForLots()

    On Error GoTo err
    Dim ITM As ListItem

    For i = 1 To LV.ListItems.Count
        If LV.ListItems(i).Checked Then
            IssID = LV.ListItems(i).Tag
            With ITM
                Set ITM = LVIssued.ListItems.Add(, , LVIssued.ListItems.Count + 1)
                ITM.Tag = LV.ListItems(i).Key
                ITM.ListSubItems.Add(, cmbVender.ID & "'", cmbVender.Text).Tag = cmbVender.ID
                ITM.ListSubItems.Add(, , IssID & "").Tag = DTPicker2
                ITM.ListSubItems.Add , , LV.ListItems(i).SubItems(2)
                ITM.ListSubItems.Add , , LV.ListItems(i).SubItems(3)
                ITM.ListSubItems.Add , , Val(cmbVender.Tag)
                ITM.ListSubItems.Add , , LV.ListItems(i).SubItems(4)
                'itm.ListSubItems.Add , , Val(LV.ListItems(i).SubItems(5))
                If cmbProcess.Tag = "Lot No" Then
                    ITM.ListSubItems(6).Tag = LV.ListItems(i).ListSubItems(1).Tag
                End If
            End With
            LV.ListItems(i).SubItems(4) = 0
            
        End If
    Next
    
    
    'Remove Posted Entries From the List.
    i = 1
    Do While i <= LV.ListItems.Count
        If LV.ListItems(i).Checked Then
            LV.ListItems.Remove i
            On Error Resume Next
            cmbDesc.RemoveItem i - 1
            If err.Number = -2147024809 Then Resume Next
                
            i = i - 1
        End If
        i = i + 1
    Loop
    
    'if cmbitems.
    'Call LV_ItemClick(LV.ListItems(i))
    
    
    
    txtQty = ""
    Exit Sub
err:
    MsgBox err.Description
        
End Sub
