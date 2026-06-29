VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmVenderList 
   ClientHeight    =   6795
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   10365
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   10365
   WindowState     =   2  'Maximized
   Begin ComboList.Usercontrol1 cmbVendName 
      Height          =   285
      Left            =   2280
      TabIndex        =   6
      Top             =   750
      Width           =   5040
      _ExtentX        =   8890
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
   Begin ComboList.Usercontrol1 cmbVendID 
      Height          =   285
      Left            =   0
      TabIndex        =   7
      Top             =   750
      Width           =   2265
      _ExtentX        =   3995
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
      Left            =   15
      TabIndex        =   0
      Top             =   -90
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vender List"
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
         Left            =   4260
         TabIndex        =   1
         Top             =   120
         Width           =   1830
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vender List"
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
         Left            =   4275
         TabIndex        =   2
         Top             =   135
         Width           =   1830
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   5145
      Left            =   15
      TabIndex        =   3
      Top             =   1215
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   9075
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   15199215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Vend ID"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Acc No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Vender Name"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Phone 1"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Phone 2"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Fax"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Balance"
         Object.Width           =   0
      EndProperty
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   525
      Width           =   7320
      BackColor       =   11517387
      Caption         =   " Vender ID                                 Vender"
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
      Left            =   8685
      TabIndex        =   5
      Top             =   6420
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
   Begin MSForms.CommandButton cmdOpt 
      Height          =   375
      Left            =   7005
      TabIndex        =   4
      Top             =   6420
      Width           =   1680
      Caption         =   "Options         "
      PicturePosition =   327683
      Size            =   "2963;661"
      Accelerator     =   79
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Menu mnupop 
      Caption         =   "mnupop"
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuShowAll 
         Caption         =   "Show &Inactive Venders"
      End
      Begin VB.Menu mnuLedger 
         Caption         =   "View &Ledger For Last 30 Days"
      End
      Begin VB.Menu mnuduePymts 
         Caption         =   "View Due Payments"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "Add &New Vender"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit Vender Information"
      End
      Begin VB.Menu mnuInActive 
         Caption         =   "Make This Vender Ina&ctive"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyVender 
         Caption         =   "Copy Vender"
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrchs 
         Caption         =   "Enter Purc&hase From This Vender"
      End
      Begin VB.Menu mnuPymt 
         Caption         =   "Enter Pa&yment To This Vender"
      End
      Begin VB.Menu dash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintDetail 
         Caption         =   "Print Selected Vender's Detail"
      End
      Begin VB.Menu mnuPrintList 
         Caption         =   "&Print Vender List"
      End
      Begin VB.Menu mnuPrintLabels 
         Caption         =   "Print Mailing Labels"
      End
      Begin VB.Menu mnuPrintVendCards 
         Caption         =   "Print Vender Cards"
      End
      Begin VB.Menu mnuPrintDetailList 
         Caption         =   "Print Vender &Detailed List"
      End
   End
End
Attribute VB_Name = "frmVenderList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Add As Boolean, Edit As Boolean
Dim PermissionDenied As Boolean

Private Sub cmbVendID_matched()
    cmbVendName.ID = cmbVendID.ID
    Call GoToSelected
End Sub

Private Sub cmbVendName_matched()
    cmbVendID.ID = cmbVendName.ID
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOpt_Click()

    With cmdOpt
        Call CheckMenu
        Me.PopUpMenu mnupop, 255, .Left, .Top + .Height
    End With
    
End Sub

Private Sub Form_Activate()

    'Check For User Access
    If PermissionDenied Then
        Me.Hide
        MsgBox "You Do not Have Permissions To Use This Option. Please Contact Your Administrator", vbInformation
        Unload Me
        Exit Sub
    End If
    mnupop.Visible = False
    
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    Exit Sub
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
    
    Call mnuEdit_Click
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()

    'PermissionDenied = Not UserHasAccess("VenderList")
    Call RefreshList
    mnuAdd.Enabled = Not OutOfFinancialYear
    mnuEdit.Enabled = Not OutOfFinancialYear
    '  mnuDelete.Enabled = Not OutOfFinancialYear
    cmbVendID.ListHeight = 2800
    cmbVendName.ListHeight = 2800
    
    cmbVendID.AddVals con, "VendID1", "Venders", "VendID", " Where VendID<>129"
    cmbVendName.AddVals con, "VenderName", "Venders", "VendID", " Where VendID<>129"
    
End Sub

Private Sub GoToSelected()

    On Error Resume Next
    LV.ListItems(cmbVendID.ID & "'").EnsureVisible
    LV.ListItems(cmbVendID.ID & "'").Selected = True
    LV.SetFocus
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
 
    Dim TitleBottom As Long

    TitleBottom = cmbVendName.Top + cmbVendName.Height
 
    LV.Move 50, TitleBottom, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + cmdClose.Height + 100)
 
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 50
    cmdOpt.Move cmdClose.Left - cmdOpt.Width - 75, cmdClose.Top
 
 
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button <> 2 Then Exit Sub
    Call CheckMenu
    Me.PopUpMenu mnupop
    
End Sub

Public Sub mnuAdd_Click()

    Dim F As New frmVender
    With F
        .Add = True
        .Show 1
    End With
    mnuRefresh_Click
    
End Sub


Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub mnuCopyVender_Click()
    If Val(LV.SelectedItem.Key) = 129 Then Exit Sub
    Load frmVender
    Call frmVender.CopyVender(Val(LV.SelectedItem.Key))
    
    mnuRefresh_Click
End Sub

Private Sub mnuduePymts_Click()

    With frmDueRcptsPymts
        Call .FillCmbs(True)
        .Show
        .cmbAccNo = LV.SelectedItem
        .SetFocus
    End With
    
End Sub

Private Sub mnuInActive_Click()
    
    'Exit Sub   'FOr now...
    If LV.ListItems.Count = 0 Then Exit Sub

    If mnuInActive.Tag = "Active" Then
        If Val(LV.SelectedItem.SubItems(6)) <> 0 Then
            MsgBox "This Vender Cannot Be Inactive Unless Balance Becomes Zero.", vbInformation
            Exit Sub
        Else
            'Check The Balance
            If MsgBox("Are You Sure To Make '" & LV.SelectedItem.SubItems(2) & "' Inactive", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            
            Dim VendBal As Double
            VendBal = GetVendBalance
            If VendBal = 0 Then
                con.Execute "Update Venders Set Active=0 Where VendID=" & Val(LV.SelectedItem.Key)
            Else
                con.Execute "Update Venders Set Active=Null Where VendID=" & Val(LV.SelectedItem.Key)
            End If
            
        End If
    Else
        con.Execute "Update Venders Set Active=1 Where VendID=" & Val(LV.SelectedItem.Key)
    End If
  
    Call RefreshList
    
End Sub

Private Function GetVendBalance() As Double
    
    Dim cmd As New ADODB.Command
    Dim rsLedger As New ADODB.Recordset
    
    Dim TotalSalDays As Integer
    TotalSalDays = Day(ServerDate)
    
    Dim IsComputed As Boolean
    Dim STLoan As Double, LongLoan As Double
    Dim VendWorkDone As Double
    Dim NetPayable As Long
    
    Dim DT As String
    Dim IsPayed As Boolean
    Dim IsTransfered As Boolean
        
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "SP_VendAdvSalary"
        Set .ActiveConnection = con
        .Parameters("@ParDate").value = ServerDate
        .Parameters("@VendID").value = Val(LV.SelectedItem.Key)
        .Parameters("@ProcessID").value = 0
        .Parameters("@TotMonDays").value = TotalSalDays
        
        If rsLedger.State = 1 Then rsLedger.Close
        rsLedger.CursorType = adOpenKeyset
        rsLedger.LockType = adLockReadOnly
        Set rsLedger = cmd.Execute
        
    End With
    
    
    With rsLedger
    
        
        If .EOF Then
            GetVendBalance = 0
        Else
            Call Computed(![VendID], IsComputed, STLoan, LongLoan, DT, IsPayed, IsTransfered, ServerDate, VendWorkDone)
            
            If Not IsComputed Then
                    
                'Long Term Loan Deduction
                Dim lDedAmt As Long, lClsAmt As Long, lTotAmt As Long
                If IsNull(![DedAmt]) Then lDedAmt = 0 Else lDedAmt = ![DedAmt]
                If IsNull(![totamt]) Then lTotAmt = 0 Else lTotAmt = ![totamt]
                If IsNull(![ClsAmt]) Then lClsAmt = 0 Else lClsAmt = ![ClsAmt]
                If IsNull(![total]) Then VendWorkDone = 0 Else VendWorkDone = Val(![total] & "")
                
                If lDedAmt > (lTotAmt - lClsAmt) Then
                    LongLoan = (lTotAmt - lClsAmt)
                Else
                    LongLoan = lDedAmt
                End If
                
            End If
        
        
            NetPayable = Round((VendWorkDone - Val(![Payed] & "") - LongLoan - Val(![TotalResDeduction] & "") - Val(![AdvSalpaid] & "")) / 10, 0) * 10
        
            GetVendBalance = NetPayable
        End If
    End With
    
End Function

Private Sub Computed(VendID As Long, ByRef IsComputed As Boolean, ByRef STLoan As Double, ByRef LongLoan As Double, ByRef DT As String, ByRef Payed As Boolean, ByRef Transfered As Boolean, ToDate As Date, ByRef VendWorkDone As Double)

    Dim rsCheck As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    With cmd
        '.CommandText = "Select * from VSalarySettings where empid='" & Emp & "'"
        .CommandType = adCmdStoredProc
        .CommandText = "SP_VendSalarySettings"
        Set .ActiveConnection = con
        .Parameters.Refresh
        .Parameters("@ParDate").value = ToDate
        
        .Parameters("@VendID").value = VendID
        
        .Parameters("@AdvSalary").value = 0
        
        rsCheck.CursorType = adOpenForwardOnly
        rsCheck.LockType = adLockReadOnly
        Set rsCheck = .Execute
    End With
    
    With rsCheck
        
        If .RecordCount > 0 Then
            IsComputed = True
            
            
            DT = Format(![DT], "dd-MMM-yyyy") & ""
            ', AbsentDays, OverTimeHours, NTime, LTime)
            LV.Checkboxes = True
            'cmdUpdate.Visible = UserHasAccess("FinalAdvSal")
            
            Payed = ![Payed] & ""
            Transfered = ![Transfered] & ""
            
            If IsNull(![LongLoan]) Then LongLoan = 0 Else LongLoan = ![LongLoan]
            STLoan = Val(![STAmt] & "")
            
            'Get The Salary Depending On Salary Sheet Type
            'VendWorkDone = IIf(chkExternal = vbUnchecked, Val(![Total] & ""), Val(![ExtSalary] & ""))
            VendWorkDone = Val(![WorkDone] & "")
            
            'EmpSal = Val(![StartingSalary] & "")
            
            cmdUpdate.Caption = "Save Changes"
            
        Else
            
            IsComputed = False
            
            
            
        End If
        
        
    End With
    
End Sub

Private Sub mnuEdit_Click()
    If Val(LV.SelectedItem.Key) = 129 Then Exit Sub
    Load frmVender
    Call frmVender.EditVender(Val(LV.SelectedItem.Key))
    
    mnuRefresh_Click
End Sub



Private Sub mnuLedger_Click()

    With frmLedger
        .cmbAccNo = LV.SelectedItem
        .cmdRefresh_Click
        .Show
    End With
    
End Sub

Private Sub mnuPrintDetail_Click()
    
    On Error GoTo err
    Screen.MousePointer = vbHourglass
    With Cr1
        .ReportFileName = RptPath & "\VendDetailList.rpt"
        .Connect = ConnectStr
        .Formulas(0) = "Company='" & company & "'"
        .Formulas(1) = "Address='" & COMPANYADDRESS & "'"
      
        .SelectionFormula = IIf(mnuShowAll.Checked, "", "{VVenders.Active}=True AND {VVenders.AccNo}='" & LV.SelectedItem & "'")
        .Action = 1
        .PageZoomNext
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub mnuPrintDetailList_Click()
    
    On Error GoTo err
    Screen.MousePointer = vbHourglass
    With Cr1
        .ReportFileName = RptPath & "\VendDetailList.rpt"
        .Connect = ConnectStr
        .Formulas(0) = "Company='" & company & "'"
        .Formulas(1) = "Address='" & COMPANYADDRESS & "'"
        .SelectionFormula = IIf(mnuShowAll.Checked, "", "{VVenders.Active}=True")
        .Action = 1
        .PageZoomNext
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintLabels_Click()
    
    On Error GoTo err
    
    Call ShowVendCardsrpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintList_Click()
    
    On Error GoTo err
    
    Call ShowVendListrpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintVendCards_Click()

     On Error GoTo err
    Screen.MousePointer = vbHourglass
    
    Dim rpt As craxddrt.Report 'New rptVendCards
    Set rpt = rptApp.OpenReport(RptPath & "\VenderCards.rpt") '
    Dim FFDefs As FormulaFieldDefinitions
    Dim FFDef As craxddrt.FormulaFieldDefinition
    Set FFDefs = rpt.FormulaFields
    
    For Each FFDef In FFDefs
        If FFDef.Name = "{@Valid Till}" Then
            FFDef.Text = "#" & Format(DateAdd("yyyy", 1, ServerDate), "dd-MMM-yyyy") & "#"
            Exit For
        End If
    Next
    
    Dim F As New frmPrevRpt
    
    F.ShowReport " ", rpt
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
End Sub

Private Sub mnuPymt_Click()

    'Dim f As New frmCashPV
    With F
        .cmbAccNo = LV.SelectedItem
        .Show
    End With
    
End Sub

Private Sub mnuRefresh_Click()
    Call RefreshList
End Sub


Private Sub RefreshList()
    
    On Error GoTo err
    Screen.MousePointer = vbHourglass

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim PrevIndex As Double
    Dim TheColor As Long
    Dim ActiveCondition As String
    ActiveCondition = " Where VendID<>129"
    Dim TempID As Integer
    If LV.ListItems.Count > 0 Then PrevIndex = LV.SelectedItem.Index

    With rs
        'Check For Show All Option
        ActiveCondition = " Where VendID<>129 "
        ActiveCondition = IIf(mnuShowAll.Checked, "", ActiveCondition & " AND IsNull(Active,1)=1 ")
        .CursorLocation = adUseClient
        .Open "Select * From Venders " & ActiveCondition & " Order By PhaseID", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        
        'For i = 0 To .RecordCount - 1
        Dim SubVendClr As Long
        SubVendClr = vbBlue
        Do Until .EOF
            'Set Color For Inactive
            'If ![Active] = 0 Then
            '    TheColor = &H808080
            '    TheTag = "Inactive"
            'Else
            '    TheColor = &H800000
            '    TheTag = "Active"
            'End If
            If ![VendType] = "2" Then
                TheColor = SubVendClr
            ElseIf IsNull(![Active]) Then
                TheColor = RGB(100, 0, 100)
            ElseIf Abs(![Active]) = 0 Then
                TheColor = vbRed
            Else
                TheColor = 0
            End If
            
            TempID = ![VendID]
            Set ITM = LV.ListItems.Add(, TempID & "'", ![VendID1])
            'Set ITM = LV.ListItems.Add(, ![VendID] & .AbsolutePosition & "'", ![AccNo], , TheTag)
            If IsNull(![Active]) Then
                ITM.Tag = "Null"
            ElseIf ![Active] Then
                ITM.Tag = "Active"
            Else
                ITM.Tag = "InActive"
            End If
            
            
            ITM.ForeColor = TheColor
            'ITM.Tag = TheTag
            ITM.Bold = True
            
            ITM.ListSubItems.Add , , ![AccNo] & ""
            ITM.ListSubItems(1).ForeColor = TheColor
            ITM.ListSubItems.Add(, , ![VenderName] & "").ToolTipText = ![VenderName] & ""
            ITM.ListSubItems(2).ForeColor = TheColor
            
            'ITM.ListSubItems.Add(, , ![ContactPerson] & "").ToolTipText = ![ContactPerson] & ""
            'ITM.ListSubItems(3).ForeColor = TheColor
            ITM.ListSubItems.Add(, , ![Phone1] & " ").ForeColor = TheColor
            ITM.ListSubItems.Add(, , ![Phone2] & " ").ForeColor = TheColor
            ITM.ListSubItems.Add(, , ![Fax1] & " ").ForeColor = TheColor
            ITM.ListSubItems.Add().ForeColor = TheColor
            'ITM.ListSubItems.Add(, , Format(Val(![Balance] & ""), "0.00 DR;0.00 CR;0.00     ")).ForeColor = TheColor
            .MoveNext
            
            'If Not .EOF Then
            '    Do While (![VendID] = TempID)
            '        Set ITM = LV.ListItems.Add(, ![EntryID] & "'")
            '
            '        ITM.ListSubItems.Add(, , ![AccNo] & "").ToolTipText = ![AccNo] & ""
            '        ITM.ListSubItems(1).ForeColor = TheColor
            '
            '        ITM.ListSubItems.Add(, , ![AccTitle] & "").ToolTipText = ![AccTitle] & ""
            '        ITM.ListSubItems(2).ForeColor = TheColor
            '        'ITM.ListSubItems.Add(, , ![ContactPerson] & "").ToolTipText = ![ContactPerson] & ""
            '        'ITM.ListSubItems(2).ForeColor = TheColor
            '        'ITM.ListSubItems.Add(, , ![Phone1] & " ").ForeColor = TheColor
            '        'ITM.ListSubItems.Add(, , ![Phone2] & " ").ForeColor = TheColor
            '        'ITM.ListSubItems.Add(, , ![Fax1] & " ").ForeColor = TheColor
            '        'ITM.ListSubItems.Add
            '        ITM.ListSubItems.Add
            '        ITM.ListSubItems.Add
            '        ITM.ListSubItems.Add
            '        ITM.ListSubItems.Add
            '        ITM.ListSubItems.Add(, , Format(Val(![Balance] & ""), "0.00 DR;0.00 CR;0.00     ")).ForeColor = TheColor
            '        .MoveNext
            '    Loop
            'End If
            
        Loop
        .Close
    End With
    Set rs = Nothing

    If PrevIndex > 0 And PrevIndex <= LV.ListItems.Count Then LV.ListItems(PrevIndex).Selected = True
   
    On Error Resume Next
    'LV.SelectedItem.EnsureVisible
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrchs_Click()

    'Dim f As New frmPurchaseJV
    With F
        .cmbVendAccno = LV.SelectedItem
        .Show
    End With
    
End Sub

Private Sub mnuShowAll_Click()

    mnuShowAll.Checked = Not mnuShowAll.Checked
    Call RefreshList
    
End Sub
 
 
 
Private Sub CheckMenu()

    If OutOfFinancialYear Then
        mnuEdit.Enabled = False
        mnuLedger.Enabled = False
        mnuPrintList.Enabled = True
        mnuPrintDetail.Enabled = True
        mnuPrintDetailList.Enabled = True
        mnuAdd.Enabled = False
        mnuInActive.Enabled = False
        mnuPrchs.Enabled = False
        mnuPymt.Enabled = False
        mnuduePymts.Enabled = True
    ElseIf LV.ListItems.Count = 0 Then
        mnuEdit.Enabled = False
        mnuLedger.Enabled = False
        mnuPrintList.Enabled = False
        mnuPrintDetail.Enabled = False
        mnuPrintDetailList.Enabled = False
        mnuAdd.Enabled = True
        mnuInActive.Enabled = True
        mnuPrchs.Enabled = False
        mnuPymt.Enabled = False
        mnuduePymts.Enabled = False
    Else
        mnuEdit.Enabled = True
        mnuLedger.Enabled = True
        mnuPrintList.Enabled = True
        mnuPrintDetail.Enabled = True
        mnuPrintDetailList.Enabled = True
        mnuAdd.Enabled = True
        mnuInActive.Caption = IIf(LV.SelectedItem.Tag = "Active", "Make This Vender Ina&ctive", "Make This Vender A&ctive")
        mnuInActive.Tag = LV.SelectedItem.Tag
        mnuInActive.Enabled = True
        
        mnuPrchs.Enabled = (mnuInActive.Tag = "Active")
        mnuPymt.Enabled = (mnuInActive.Tag = "Active")
        mnuduePymts.Caption = "View Due Payments From " & LV.SelectedItem.SubItems(1)
        mnuduePymts.Enabled = True
    End If
    mnuPrintLabels.Enabled = mnuPrintList.Enabled
    
End Sub


