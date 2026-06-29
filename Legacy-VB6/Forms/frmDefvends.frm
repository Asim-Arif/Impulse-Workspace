VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmDefvends 
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
         Caption         =   "Defaulter Venders"
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
         Caption         =   "Defaulter Venders"
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
      Begin ComboList.Usercontrol1 cmbProcess 
         Height          =   285
         Left            =   90
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   435
         Width           =   4755
         _ExtentX        =   8387
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
      Begin VB.CommandButton cmdUnchkAll 
         Caption         =   "Uncheck All"
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
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdChkAll 
         Caption         =   "Check All"
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
         Visible         =   0   'False
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   1005
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Vend ID"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Vender"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Last Rcvd Date"
            Object.Width           =   2822
         EndProperty
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         Caption         =   " Process"
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
         TabIndex        =   11
         Top             =   195
         Width           =   4785
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   360
         Left            =   4875
         TabIndex        =   10
         Top             =   360
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
         Picture         =   "frmDefvends.frx":0000
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
         Picture         =   "frmDefvends.frx":0112
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
Attribute VB_Name = "frmDefvends"
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
    If SSTab1.Tab = 0 Then
        For i = 1 To LV.ListItems.Count
            LV.ListItems(i).Checked = True
        Next i
    Else
        For i = 1 To LVPrinted.ListItems.Count
            LVPrinted.ListItems(i).Checked = True
        Next i
    End If
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    Call StartTrans(con)
    con.Execute "Delete From PrintDefVendList"
    For i = 1 To LV.ListItems.Count
        With LV.ListItems(i)
            con.Execute "Insert Into PrintDefVendList(VendID,VendName,DT) Values('" & _
             .SubItems(1) & "','" & .SubItems(2) & "','" & .SubItems(3) & "')"
        End With
    Next
    con.CommitTrans
    
    Dim rpt As craxddrt.Report  'New rptDailyIssRcv
    Set rpt = rptApp.OpenReport(RptPath & "\rptDefVends.rpt") '
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
    If SSTab1.Tab = 0 Then
        For i = 1 To LV.ListItems.Count
            LV.ListItems(i).Checked = False
        Next i
    Else
        For i = 1 To LVPrinted.ListItems.Count
            LVPrinted.ListItems(i).Checked = False
        Next i
    End If
End Sub

Private Sub DtFrom_Change()
    'Call setDateRange(DtFrom)
End Sub

Private Sub DtTo_Change()
    'Call setDateRange(DtTo)
End Sub




Private Sub Form_Load()

    
    cmbProcess.ListHeight = 2800
    
    cmbProcess.AddItem "<All Processes>", "0"
    cmbProcess.AddVals con, "Description", "Processes", "ProcessID", " Order By SNo"
    
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
    'FromDate = ServerDate
    'ToDate = ServerDate
    
    FromDate = DateAdd("d", -7, ServerDate)
    
    Dim strSql As String
    If cmbProcess.ID = "0" Then
        strSql = "Select venders.*,t1.lastentrydt From Venders left outer JOIN (select VendID,Max(DT) as lastEntrydt From VVendRcvdDetail " & _
        "where wastage=0 and lostqty=0 Group By vendID) T1 ON t1.VendID = venders.VendID Where venders.Vendid  In(Select VendID From VVendIssdDetail Where Issqty>rcvdqty) " & _
        "and venders.Vendid   Not In(select VendID From VendReceived Where DT > '" & FromDate & "') And Venders.VendID<>129"
    Else
        strSql = "Select venders.*,t1.lastentrydt From Venders left OUTER JOIN (select VendID,Max(DT) as LastEntrydt From VVendRcvdDetail " & _
        "where wastage=0 and lostqty=0 Group By vendID) T1 ON t1.VendID = venders.VendID Where venders.Vendid  In(Select VendID From VVendIssdDetail Where Issqty>rcvdqty) " & _
        "and venders.Vendid Not In(select VendID From VendReceived Where DT > '" & FromDate & "') And Venders.VendID<>129 And Venders.PhaseID=" & cmbProcess.ID
    End If
    
    With rs
        .Open strSql, con, adOpenForwardOnly, adLockReadOnly
        
        LV.ListItems.Clear
        'LV.Sorted = False
        Do Until .EOF
            Set ITM = LV.ListItems.Add(, , .AbsolutePosition)
            
            ITM.SubItems(1) = ![VendID1] & ""
            ITM.SubItems(2) = ![VenderName] & ""
            ITM.SubItems(3) = Format(!LastEntryDT & "", "dd-MMM-yyyy")
            
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
