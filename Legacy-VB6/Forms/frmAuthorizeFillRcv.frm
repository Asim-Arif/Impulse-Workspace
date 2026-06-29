VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmAuthorizeFillRcv 
   ClientHeight    =   9510
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   9510
   ScaleMode       =   0  'User
   ScaleWidth      =   10667.95
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10560
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Authorization"
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
         Index           =   8
         Left            =   105
         TabIndex        =   2
         Top             =   165
         Width           =   10605
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Authorization"
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
         Height          =   465
         Index           =   9
         Left            =   120
         TabIndex        =   3
         Top             =   180
         Width           =   10590
      End
   End
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   8880
      Left            =   15
      TabIndex        =   0
      Top             =   615
      Width           =   11850
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         ScaleHeight     =   255
         ScaleWidth      =   945
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1170
         Visible         =   0   'False
         Width           =   945
         Begin VB.TextBox txtQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   950
         End
      End
      Begin MSComctlLib.ListView LVRcv 
         Height          =   8160
         Left            =   135
         TabIndex        =   8
         Top             =   180
         Width           =   11580
         _ExtentX        =   20426
         _ExtentY        =   14393
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Reciept ID"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Vend ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Vender"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Item Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Item Description"
            Object.Width           =   4207
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Process"
            Object.Width           =   1614
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Qty"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Repair"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Waste"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Lost"
            Object.Width           =   1411
         EndProperty
      End
      Begin MSForms.CommandButton cmdDelete 
         Height          =   375
         Left            =   9810
         TabIndex        =   5
         Top             =   8430
         Width           =   1905
         ForeColor       =   0
         Caption         =   "Close"
         PicturePosition =   327683
         Size            =   "3360;661"
         Picture         =   "frmAuthorizeFillRcv.frx":0000
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdUpdate 
         Height          =   375
         Left            =   7845
         TabIndex        =   4
         Top             =   8430
         Width           =   1905
         ForeColor       =   0
         Caption         =   "Save & Print"
         PicturePosition =   327683
         Size            =   "3360;661"
         Accelerator     =   83
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
Attribute VB_Name = "frmAuthorizeFillRcv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rptSQL As String
Dim TableNAme As String
Private Sub DeleteIssd()

    For i = 1 To LVIss.ListItems.Count
        If LVIss.ListItems(i).Checked Then
            con.Execute "Delete From VendIssdDetail Where EntryID=" & Val(LVIss.ListItems(i).Key)
        End If
    Next i
    'Now Clean VendIssued
    con.Execute "Delete From VendIssued Where EntryID Not In(Select RefID From VendIssdDetail)"
    
End Sub
Private Sub DeleteRcvd()

    For i = 1 To LVRcv.ListItems.Count
        If LVRcv.ListItems(i).Checked Then
            con.Execute "Delete From VendRcvdDetail Where EntryID=" & Val(LVRcv.ListItems(i).Key)
        End If
    Next i
    'Now Clean VendReceived
    con.Execute "Delete From VendReceived Where EntryID Not In(Select RefID From VendRcvdDetail)"
    
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub


Private Sub CmdDelete_Click()
    Unload Me
End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub



Private Sub cmdUnchkAll_Click()

    
    For i = 1 To LVRcv.ListItems.Count
        LVRcv.ListItems(i).Checked = False
    Next i
    

End Sub

Private Sub cmdUpdate_Click()

    On Error GoTo err
    
'    Dim strPassword As String
'    strPassword = InputBox("Please Enter Your Password.", "Identify Yourself")
'
'    If strPassword <> rsUserSettings.Fields("Password") & "" Then
'        MsgBox "Invalid Password.", vbInformation
'        Exit Sub
'    End If
    
    Call StartTrans(con)
    
    Call UpdateRcvEntries
    
    con.CommitTrans
    Call PrintRpts("Original")
    Call RefreshLV
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub PrintRpts(strOriginal As String)

    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    
    Dim rs As New ADODB.Recordset
    
    Call StartTrans(con)
    'Insert Into PrintIssRcvRptsTemp.
    
    con.Execute "Delete From PrintAuthorized"
    con.Execute "Delete From PrintIssRcvRptsTemp"
    
    For i = 1 To LVRcv.ListItems.Count
        With LVRcv.ListItems(i)
            If .Checked Then
                con.Execute "Insert Into PrintAuthorized(RecieptNo) Values('" & .Text & "')"
                con.Execute "Insert Into PrintIssRcvRptsTemp(RecieptID) Values('" & .Text & "')"
            End If
        End With
    Next
    
    con.Execute "Insert Into PrintIssRcvRpts Select CompItemDesc,DT,RecieptID,ItemCode,VenderName" & _
                ",VendID1,ProcessID,Process,RcvdQty,Wastage,LostQty,UserID,AuthUserID,AuthDT,0 As IssEntry " & _
                "From VVendRcvItemsrpt Where RecieptID In(Select RecieptNo From PrintAuthorized)"
                
'    con.Execute "Insert Into PrintIssRcvRpts Select CompItemDesc,DT,RecieptID,ItemCode,VenderName," & _
'                "VendID1,ProcessID,Process,IssQty,0 As Wastage,0 As LostQty,UserID,AuthUserID,AuthDT,1 As IssEntry " & _
'                "From VVendIssItemsrpt Where RecieptID In(Select RecieptNo From PrintAuthorized)"
'
    con.CommitTrans
    
    Dim rpt As craxddrt.Report  'New rptVendIssDetail
    Set rpt = rptApp.OpenReport(RptPath & "\rptVendIssDetail.rpt") '
    Dim ForTxt As String
    If TableNAme = "VVendIssItemsrpt" Then
        ForTxt = "GATEPASS ISSUED"
    Else
        ForTxt = "GOODS RECEIVED"
    End If
    
    Dim FFDefs As FormulaFieldDefinitions
    Dim FFDef As craxddrt.FormulaFieldDefinition
    Set FFDefs = rpt.FormulaFields
    
    For Each FFDef In FFDefs
'        If FFDef.Name = "{@Heading}" Then
'            FFDef.Text = "'" & ForTxt & "'"
        If FFDef.Name = "{@ForOriginal}" Then
            FFDef.Text = "'" & strOriginal & "'"
        End If
    Next
        
    Load frmPrevRpt
    frmPrevRpt.ShowReport " ", rpt
    
    Screen.MousePointer = vbDefault

    
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
End Sub

Private Sub UpdateRcvEntries()

    Dim cmd As New ADODB.Command
    Dim ProcID As Long
    
    For i = 1 To LVRcv.ListItems.Count
        If LVRcv.ListItems(i).Checked Then
            ProcID = LVRcv.ListItems(i).ListSubItems(3).Tag
            
            con.Execute "Update VendRcvdDetail Set ReqAuth=0 Where EntryID=" & Val(LVRcv.ListItems(i).Key)
            con.Execute "Insert Into VendRcvAuthDetail(UserID,RcvID,DT) Values(" & _
                UserID & "," & Val(LVRcv.ListItems(i).Key) & ",'" & ServerDate & " " & ServerTime & "')"
                
            With cmd
                .ActiveConnection = con
                .CommandType = adCmdStoredProc
                
                .CommandText = "SP_AuthorizeFilling"
                .Parameters("@ItemCode").value = LVRcv.ListItems(i).SubItems(3)
                .Parameters("@ProcessID").value = ProcID
                .Parameters("@Qty").value = Val(LVRcv.ListItems(i).SubItems(6))
                .Parameters("@Repair").value = Val(LVRcv.ListItems(i).SubItems(7))
                .Parameters("@WasteQty").value = 0 'Val(LVRcv.ListItems(i).SubItems(5) & "")
                .Parameters("@iRefID").value = LVRcv.ListItems(i).ListSubItems(5).Tag
                .Parameters("@VendID").value = Val(LVRcv.ListItems(i).ListSubItems(2).Tag & "")
                .Parameters("@LotNo").value = 0
                .Parameters("@RecID").value = ""
                .Parameters("@ReqAuth").value = 0   'Has To Pass Zero.
                .Parameters("@AuthorizeEntry").value = 1
                .Parameters("@UserID").value = UserID
                .Parameters("@UpdateEntryID").value = Val(LVRcv.ListItems(i).Key)
                .Parameters("@RcvDT").value = ServerDate 'This Won't be Used It's Just 4 Passing
                .Execute
                
            End With
            
        End If
        
    Next i
    
End Sub
Private Sub UpdateIssEntries()

    Dim cmd As New ADODB.Command
    Dim IssLot As Boolean
    
    With cmd
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_UpdateForIAR"
        For i = 1 To LVIss.ListItems.Count    'Update Rcv Quantities.
        
            If LVIss.ListItems(i).Checked Then
            
                IssLot = RcvLot(LVIss.ListItems(i).ListSubItems(3).Tag) 'Whether Or Not Issuing Lot.
                
                con.Execute "Update VendIssdDetail Set ReqAuth=0 Where EntryID=" & Val(LVIss.ListItems(i).Key)
                'Insert Into Iss Auth Detail Table
                con.Execute "Insert Into VendIssAuthDetail(UserID,IssID,DT) Values(" & _
                UserID & "," & Val(LVIss.ListItems(i).Key) & ",'" & ServerDate & " " & ServerTime & "')"
                
                If Not IssLot Then
                    '.Parameters("@OrderNo").value = LVIss.ListItems(i).ListSubItems(1).Tag
                    .Parameters("@ItemCode").value = LVIss.ListItems(i).SubItems(3)
                    .Parameters("@ProcessID").value = LVIss.ListItems(i).ListSubItems(3).Tag
                    .Parameters("@Qty").value = Val(LVIss.ListItems(i).SubItems(6))
                    .Execute
                Else
                    con.Execute "Update VendRcvdDetail Set IssQty=RcvdQty Where LotNo=" & LVIss.ListItems(i).ListSubItems(4).Tag & " and NextProcessID=" & LVIss.ListItems(i).ListSubItems(3).Tag
                End If
                
            End If
        Next i
    End With
    
End Sub
Private Sub DtFrom_Change()
    'Call setDateRange(DtFrom)
End Sub

Private Sub DtTo_Change()
    'Call setDateRange(DtTo)
End Sub


Private Sub Form_Load()
    Call RefreshLV
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Dim TheHeight As Long
    Frame1.Left = (Me.ScaleWidth - Frame1.Width) / 2
    
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    
End Sub


Public Sub ShowForIss(Optional ForRcv As Boolean)
    If ForRcv Then
        TableNAme = "VVendRcvItemsrpt"
    Else
        TableNAme = "VVendIssItemsrpt"
    End If
    Me.Show
End Sub
Private Sub RefreshLV()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim Cond As String
    Dim DTRange As String
    Dim ITM As ListItem
    
    Screen.MousePointer = vbHourglass
    
    
    With rs
        
        .Open "Select * From VVendRcvItemsrpt Where (ReqAuth=1) and (ProcessID=4) and VendID Not In(Select VendID From Venders Where VendType=0) Order By DT", con, adOpenForwardOnly, adLockReadOnly
        LVRcv.ListItems.Clear
        'LV.Sorted = False
        Do Until .EOF
            Set ITM = LVRcv.ListItems.Add(, ![EntryID] & "'", ![RecieptID] & "")
            ITM.ListSubItems.Add(, , ![VendID1] & "").ToolTipText = ![VendID1] & ""
            'itm.ListSubItems(1).Tag = ![OrderNo] & ""
            If ![VenderName] & "" = "" Then
                ITM.ListSubItems.Add(, , ![Supervisor] & "").ToolTipText = ![Supervisor] & ""
            Else
                ITM.ListSubItems.Add(, , ![VenderName] & "").ToolTipText = ![VenderName] & ""
                ITM.ListSubItems(2).Tag = ![VendID] & ""
            End If
            ITM.ListSubItems.Add(, , ![ItemCode] & "").ToolTipText = ![ItemCode] & ""
            ITM.ListSubItems(3).Tag = ![ProcessID] & ""
            ITM.ListSubItems.Add(, , ![ItemDescription] & "").ToolTipText = ![ItemDescription] & ""
            ITM.ListSubItems(4).Tag = ![LotNo] & ""
            ITM.ListSubItems.Add(, , ![Process] & "").ToolTipText = ![Process] & ""
            ITM.ListSubItems(5).Tag = ![RefID] & ""
            ITM.SubItems(6) = ![RcvdQty] & ""
            ITM.SubItems(7) = ![Wastage] & ""
            ITM.SubItems(8) = ![LostQty] & ""
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Private Sub RefreshSnos()
    For i = 1 To LV.ListItems.Count
        LV.ListItems(i).Text = i
    Next i
End Sub

Private Sub LVRcv_DblClick()
    
    If LVRcv.ListItems.Count = 0 Then Exit Sub
    If Val(LVRcv.SelectedItem.ListSubItems(3).Tag) <> 4 Then Exit Sub
    
    With Pic
        .Move LVRcv.Left + 10 + LVRcv.ColumnHeaders(8).Left, LVRcv.SelectedItem.Top + LVRcv.Top
        .Visible = True
    End With
    
    With txtQty
        
        .Text = Val(LVRcv.SelectedItem.SubItems(7))
        
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
        
    End With
    
End Sub


Private Sub LVRcv_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    If Val(LVRcv.SelectedItem.ListSubItems(3).Tag) <> 4 Then Exit Sub
    With Pic
        .Move LVRcv.Left + 10 + LVRcv.ColumnHeaders(8).Left, LVRcv.SelectedItem.Top + LVRcv.Top
        .Visible = True
    End With
    
    With txtQty
        
        .Text = Val(LVRcv.SelectedItem.SubItems(7))
        
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)

    On Error GoTo err
    If KeyAscii = 13 Then
        If Val(txtQty) > Val(LVRcv.SelectedItem.SubItems(6)) + Val(LVRcv.SelectedItem.SubItems(8)) + Val(LVRcv.SelectedItem.SubItems(9)) Or Val(txtQty) < 0 Then
            MsgBox "Invalid Quantities!", vbInformation
            Exit Sub
        End If
        
        
        LVRcv.SelectedItem.SubItems(7) = Val(txtQty)
            
        Pic.Visible = False
                
    ElseIf KeyAscii = 27 Then
        Pic.Visible = False
    End If
    
    Call OnlyNums(KeyAscii, False)
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
