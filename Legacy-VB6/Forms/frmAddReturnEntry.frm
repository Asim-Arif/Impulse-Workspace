VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmAddReturnEntry 
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   9945
   ScaleMode       =   0  'User
   ScaleWidth      =   10641.08
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
         Caption         =   "Return Entries From Venders"
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
         Caption         =   "Return Entries From Venders"
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
      Height          =   9285
      Left            =   15
      TabIndex        =   0
      Top             =   615
      Width           =   11850
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Return Selected"
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
         Left            =   3345
         TabIndex        =   15
         Top             =   8805
         Width           =   2040
      End
      Begin ComboList.Usercontrol1 cmbProcess 
         Height          =   285
         Left            =   90
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   435
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
      Begin VB.ComboBox cmbRange 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmAddReturnEntry.frx":0000
         Left            =   2475
         List            =   "frmAddReturnEntry.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   435
         Width           =   2010
      End
      Begin VB.CommandButton cmdUnchkAll 
         Caption         =   "&Uncheck All"
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
         Left            =   1695
         TabIndex        =   7
         Top             =   8790
         Width           =   1575
      End
      Begin VB.CommandButton cmdChkAll 
         Caption         =   "Check &All"
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
         TabIndex        =   6
         Top             =   8790
         Width           =   1575
      End
      Begin MSComctlLib.ListView LV 
         Height          =   7860
         Left            =   90
         TabIndex        =   8
         Top             =   810
         Width           =   11670
         _ExtentX        =   20585
         _ExtentY        =   13864
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Receipt No."
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Vender"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Process"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Item"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Qty"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   6225
         TabIndex        =   11
         Top             =   420
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   47906819
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   4500
         TabIndex        =   12
         Top             =   420
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   47906819
         CurrentDate     =   37055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         Caption         =   " Process                                      Range                                 From                            To"
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
         Height          =   225
         Index           =   0
         Left            =   75
         TabIndex        =   14
         Top             =   195
         Width           =   11130
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   360
         Left            =   7800
         TabIndex        =   13
         Top             =   420
         Width           =   1500
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "2646;635"
         Accelerator     =   82
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   375
         Left            =   9060
         TabIndex        =   5
         Top             =   8790
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Print        "
         PicturePosition =   327683
         Size            =   "2355;661"
         Accelerator     =   80
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
         Left            =   10425
         TabIndex        =   4
         Top             =   8790
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "2355;661"
         Picture         =   "frmAddReturnEntry.frx":0067
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
End
Attribute VB_Name = "frmAddReturnEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rptSQL As String
Dim TableNAme As String

Private Sub cmbRange_Click()

    DtTo = ServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    Select Case cmbRange.ListIndex
        
        Case 0
            DtFrom = ServerDate
        Case 1
            DtFrom = DateAdd("d", -15, ServerDate)
        Case 2
            DtFrom = DateAdd("d", -30, ServerDate)
        Case 3
            DtFrom = DateAdd("d", -60, ServerDate)
        Case 4
            DtFrom = DateAdd("d", -90, ServerDate)
        Case 5
            DtFrom.Enabled = True
            DtTo.Enabled = True
            DtFrom.SetFocus
            Exit Sub
    End Select
    
End Sub



Private Sub cmdChkAll_Click()

    
    For i = 1 To LV.ListItems.Count
        LV.ListItems(i).Checked = True
    Next i
    
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdDelete_Click()
    
    On Error GoTo err
    Dim cmd As New ADODB.Command
    With cmd
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_ReturnIssEntry"
    End With
    
    Call StartTrans(con)
    
    For i = 1 To LV.ListItems.Count
        With LV.ListItems(i)
            If .Checked Then
            
                cmd.Parameters("@RecieptID").value = .Text
                cmd.Parameters("@ItemCode").value = .ListSubItems(3).Tag
                cmd.Parameters("@ProcessID").value = .ListSubItems(2).Tag
                cmd.Parameters("@Qty").value = Val(.SubItems(4))
                cmd.Execute
                
            End If
        End With
    Next
    
    con.CommitTrans
    Call NewRefreshLV
    Exit Sub
    
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    
    Dim rpt As craxddrt.Report  'New rptDailyIssRcv
    Set rpt = rptApp.OpenReport(RptPath & "\rptDailyIssRcv.rpt") '
    Load frmPrevRpt
    frmPrevRpt.ShowReport " ", rpt
    
    'Call NewRefreshLV
    Screen.MousePointer = vbDefault

    
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Public Sub cmdRefresh_Click()
   
    Call NewRefreshLV
   
End Sub



Private Sub cmdUnchkAll_Click()
    For i = 1 To LV.ListItems.Count
        LV.ListItems(i).Checked = False
    Next i
End Sub

Private Sub DtFrom_Change()
    'Call setDateRange(DtFrom)
End Sub

Private Sub DtTo_Change()
    'Call setDateRange(DtTo)
End Sub




Private Sub Form_Load()

    cmbRange.ListIndex = 0
    cmbProcess.ListHeight = 2800
    cmbProcess.AddItem "<All Processes>", "0"
    cmbProcess.AddVals con, "Description", "Processes", "ProcessID", " Where ProcessID<=7 Order By SNo"
    
'    cmbIssRcv.AddItem "Issue Reports"
'    cmbIssRcv.AddItem "Receive Reports"
'    cmbIssRcv.ListIndex = 0
    
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
    
    If cmbRange.ListIndex = 0 Then
        DTRange = " Convert(Varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "'"
    ElseIf cmbRange.ListIndex = 1 Then
        DTRange = " DT Between '" & DateAdd("d", -15, ServerDate) & "' and '" & ServerDate & "' or " & " Convert(Varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "'"
    ElseIf cmbRange.ListIndex = 2 Then
        DTRange = " DT Between '" & DateAdd("d", -30, ServerDate) & "' and '" & ServerDate & "' or " & " Convert(Varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "'"
    ElseIf cmbRange.ListIndex = 3 Then
        DTRange = " DT Between '" & DateAdd("d", -60, ServerDate) & "' and '" & ServerDate & "' or " & " Convert(Varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "'"
    ElseIf cmbRange.ListIndex = 4 Then
        DTRange = " DT Between '" & DateAdd("d", -90, ServerDate) & "' and '" & ServerDate & "' or " & " Convert(Varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "'"
    ElseIf cmbRange.ListIndex = 5 Then
        DTRange = " DT Between '" & DtFrom & "' and '" & DtTo & "'"
    End If
    
    If Cond = "" Then
        Cond = " Where "
    Else
        Cond = Cond & " and "
    End If
    Cond = Cond & "(" & DTRange & ") and (ReqAuth=0)"
    
    Dim ExtraCond As String
    
    If TableNAme = "VVendIssItemsrpt" Then
        ExtraCond = " and (FromPlan Is Null) "
    Else
        ExtraCond = " and (Wastage=0)"
    End If
    
    Sql = "Select * From " & TableNAme & Cond & ExtraCond & "  Order By DT"
    
    With rs
        rptSQL = Sql
        .Open Sql, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        'LV.Sorted = False
        Do Until .EOF
            Set ITM = LV.ListItems.Add(, , .AbsolutePosition)
            ITM.SubItems(1) = ![VendID1] & ""
            ITM.SubItems(2) = ![VenderName] & ""
            'itm.SubItems(3) = ![ItemCode] & ""
            'itm.SubItems(4) = ![CompItemDesc] & ""
            ITM.SubItems(3) = ![RecieptID] & ""
            If TableNAme = "VVendIssItemsrpt" Then
                ITM.SubItems(4) = ![IssQty] & ""
            Else
                ITM.SubItems(4) = ![RcvdQty] & ""
            End If
            
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
Private Sub NewRefreshLV()
    
    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim FromDate As Date
    Dim ToDate As Date
    Dim Cond As String
    Dim DTRange As String
    Dim ITM As ListItem
    
    Screen.MousePointer = vbHourglass
    
    Dim strCond As String
    With rs
        If cmbProcess.ID <> "0" Then strCond = " AND ProcessID=" & cmbProcess.ID Else strCond = " AND ProcessID<=7"
        .Open "Select * From VVendIssdDetail1 Where RcvdQty=0 AND Convert(Varchar,DT,1) Between '" & _
         Format(DtFrom, "mm/dd/yy") & "' AND '" & Format(DtTo, "mm/dd/yy") & "' " & strCond & " Order By DT", con, adOpenForwardOnly, adLockReadOnly
         
        LV.ListItems.Clear
        
        Do Until .EOF
            Set ITM = LV.ListItems.Add(, , ![RecieptID] & "")
            
            
            ITM.ListSubItems.Add(, , "[" & ![VendID1] & "] " & ![VenderName] & "").Tag = !VendID
            ITM.ListSubItems.Add(, , ![ProcessName] & "").Tag = !ProcessID
            ITM.ListSubItems.Add(, , "[" & ![ItemCode] & "] " & ![ItemName] & "").Tag = !ItemCode & ""
            ITM.SubItems(4) = ![IssQty] & ""
            
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

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Exit Sub
    If LV.SortKey + 1 = ColumnHeader.Index Then
        If LV.SortOrder = lvwAscending Then
            LV.SortOrder = lvwDescending
            ColumnHeader.Icon = "Down"
        Else
            LV.SortOrder = lvwAscending
            ColumnHeader.Icon = "Up"
        End If
    Else
        LV.ColumnHeaderIcons = Nothing
        Set LV.ColumnHeaderIcons = ImageList2
        LV.SortKey = ColumnHeader.Index - 1
        LV.SortOrder = lvwAscending
        ColumnHeader.Icon = "Down"
    End If
    Call RefreshSnos
    'If LV.Sorted = False Then LV.Sorted = True
End Sub


Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Exit Sub
    If LV.ListItems.Count = 0 Or Button <> 2 Then Exit Sub
    
    Dim strRet As String
    strRet = ShowPopUpMenu("&Edit This Order", "-", "&Add Production Planning", "-", "History Of The Order (Summary)", "History Of The Order (Detail)", "-", "Print Order (&Summary)", "Print Order (&Detail)", "-", "&Print Order")
    
    Select Case strRet
        Case "&Edit This Order"
            
        Case "&Add Production Planning"
            
        Case "History Of The Order (Summary)"
            
        Case "History Of The Order (Detail)"
            
        Case "Print Order (&Summary)"
            
        Case "Print Order (&Detail)"
            
        Case "&Print Order"
            
    End Select
    'Me.PopupMenu mnuPop
    'If LV.ListItems.Count = 0 Or Button <> 2 Or OutOfFinancialYear Then Exit Sub

    'if a Voucher Entry Show Delete Menu
    'If InStr(LV.SelectedItem, "-") > 0 Then Me.PopupMenu mnuPop
    
End Sub
Private Sub RefreshSnos()
    For i = 1 To LV.ListItems.Count
        LV.ListItems(i).Text = i
    Next i
End Sub
