VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRMVenderList 
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
   Begin Crystal.CrystalReport cr1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   480
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
            Picture         =   "frmRMVenderList.frx":0000
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRMVenderList.frx":01BC
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
            Picture         =   "frmRMVenderList.frx":0378
            Key             =   "Active"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRMVenderList.frx":07CC
            Key             =   "Inactive"
         EndProperty
      EndProperty
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
         Caption         =   "Raw Material Vender List"
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
         Left            =   3435
         TabIndex        =   1
         Top             =   120
         Width           =   4065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Raw Material Vender List"
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
         Left            =   3450
         TabIndex        =   2
         Top             =   135
         Width           =   4065
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   5295
      Left            =   15
      TabIndex        =   3
      Top             =   495
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   9340
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
         Object.Width           =   2293
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
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Phone 2"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Fax"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Balance"
         Object.Width           =   2293
      EndProperty
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   375
      Left            =   8685
      TabIndex        =   5
      Top             =   6135
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
      Top             =   6135
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
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintList 
         Caption         =   "&Print Vender List"
      End
      Begin VB.Menu mnuPrintLabels 
         Caption         =   "Print Mailing Labels"
      End
      Begin VB.Menu mnuPrintDetailList 
         Caption         =   "Print Vender &Detailed List"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmRMVenderList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Add As Boolean, Edit As Boolean
Dim PermissionDenied As Boolean
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
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()

    'PermissionDenied = Not UserHasAccess("VenderList")
    Call RefreshList
    mnuAdd.Enabled = Not OutOfFinancialYear
    mnuEdit.Enabled = Not OutOfFinancialYear
    '  mnuDelete.Enabled = Not OutOfFinancialYear
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
 
    Dim TitleBottom As Long

    TitleBottom = fTitle.Top + fTitle.Height
 
    LV.Move 50, TitleBottom, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + cmdClose.Height + 100)
 
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 50
    cmdOpt.Move cmdClose.Left - cmdOpt.Width - 75, cmdClose.Top
 
 
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button <> 2 Then Exit Sub
    Call CheckMenu
    Me.PopUpMenu mnupop
    
End Sub

Public Sub mnuAdd_Click()

    Dim f As New frmRMVender
    With f
        .Add = True
        .Show 1
    End With
    mnuRefresh_Click
    
End Sub


Private Sub mnuClose_Click()
    Unload Me
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

    If LV.ListItems.Count = 0 Then Exit Sub

    If mnuInActive.Tag = "Active" Then
        If Val(LV.SelectedItem.SubItems(6)) <> 0 Then
            MsgBox "This Vender Cannot Be Inactive Unless Balance Becomes Zero.", vbInformation
            Exit Sub
        Else
            If MsgBox("Are You Sure To Make '" & LV.SelectedItem.SubItems(1) & "' Inactive", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            con.Execute "Update Accounts Set Active=False Where Accno='" & LV.SelectedItem & "'"
        End If
    Else
        con.Execute "Update Accounts Set Active=True Where Accno='" & LV.SelectedItem & "'"
    End If
  
    Call RefreshList
    
End Sub


Private Sub mnuEdit_Click()

    Load frmRMVender
    Call frmRMVender.EditVender(Val(LV.SelectedItem.Key))
    
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

    Screen.MousePointer = vbHourglass
    With cr1
        .ReportFileName = RptPath & "\VendDetailList.rpt"
        .Connect = ConnectStr
        .Formulas(0) = "Company='" & company & "'"
        .Formulas(1) = "Address='" & COMPANYADDRESS & "'"
      
        .SelectionFormula = IIf(mnuShowAll.Checked, "", "{VVenders.Active}=True AND {VVenders.AccNo}='" & LV.SelectedItem & "'")
        .Action = 1
        .PageZoomNext
    End With
    Screen.MousePointer = vbDefault
   
End Sub


Private Sub mnuPrintDetailList_Click()

    Screen.MousePointer = vbHourglass
    With cr1
        .ReportFileName = RptPath & "\VendDetailList.rpt"
        .Connect = ConnectStr
        .Formulas(0) = "Company='" & company & "'"
        .Formulas(1) = "Address='" & COMPANYADDRESS & "'"
        .SelectionFormula = IIf(mnuShowAll.Checked, "", "{VVenders.Active}=True")
        .Action = 1
        .PageZoomNext
    End With
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub mnuPrintLabels_Click()
    Call ShowRMVendCardsrpt
End Sub

Private Sub mnuPrintList_Click()
    Call ShowRMVendListrpt
End Sub

Private Sub mnuPymt_Click()

    'Dim f As New frmCashPV
    With f
        .cmbAccNo = LV.SelectedItem
        .Show
    End With
    
End Sub

Private Sub mnuRefresh_Click()
    Call RefreshList
End Sub


Private Sub RefreshList()

    Screen.MousePointer = vbHourglass

    Dim rs As New ADODB.Recordset
    Dim itm As ListItem
    Dim PrevIndex As Double
    Dim TheColor As Long
    Dim ActiveCondition As String
    Dim TempID As Integer
    If LV.ListItems.Count > 0 Then PrevIndex = LV.SelectedItem.Index
    Set LV.SmallIcons = ImageList1
    With rs
        'Check For Show All Option
        'ActiveCondition = IIf(mnuShowAll.Checked, "", " Where Active=1")
        .CursorLocation = adUseClient
        .Open "Select * From RawMaterialVenders " & ActiveCondition & " Order By VendID", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        
        'For i = 0 To .RecordCount - 1
        Do Until .EOF
            'Set Color For Inactive
            'If ![Active] = 0 Then
            '    TheColor = &H808080
            '    TheTag = "Inactive"
            'Else
            '    TheColor = &H800000
            '    TheTag = "Active"
            'End If
                    
            TempID = ![VendID]
            Set itm = LV.ListItems.Add(, TempID & "'", ![VendID1])
            'Set ITM = LV.ListItems.Add(, ![VendID] & .AbsolutePosition & "'", ![AccNo], , TheTag)
            itm.ForeColor = TheColor
            'ITM.Tag = TheTag
            itm.Bold = True
            
            itm.ListSubItems.Add , , ![AccNo] & ""
            itm.ListSubItems(1).ForeColor = TheColor
            itm.ListSubItems.Add(, , ![VenderName] & "").ToolTipText = ![VenderName] & ""
            itm.ListSubItems(2).ForeColor = TheColor
            
            'ITM.ListSubItems.Add(, , ![ContactPerson] & "").ToolTipText = ![ContactPerson] & ""
            'ITM.ListSubItems(3).ForeColor = TheColor
            itm.ListSubItems.Add(, , ![Phone1] & " ").ForeColor = TheColor
            itm.ListSubItems.Add(, , ![Phone2] & " ").ForeColor = TheColor
            itm.ListSubItems.Add(, , ![Fax1] & " ").ForeColor = TheColor
            itm.ListSubItems.Add().ForeColor = TheColor
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
    
End Sub

Private Sub mnuPrchs_Click()

    'Dim f As New frmPurchaseJV
    With f
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


