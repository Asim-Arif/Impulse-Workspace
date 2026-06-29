VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmChartOfAccounts 
   ClientHeight    =   8625
   ClientLeft      =   165
   ClientTop       =   60
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicEditRG 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   2310
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   2340
      Begin MSForms.ComboBox cmbReportingGroup 
         Height          =   255
         Left            =   -15
         TabIndex        =   15
         Top             =   -15
         Width           =   2340
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4128;450"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
   End
   Begin VB.PictureBox PicStatus 
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   105
      ScaleHeight     =   420
      ScaleWidth      =   2625
      TabIndex        =   7
      Top             =   8085
      Visible         =   0   'False
      Width           =   2625
      Begin MSComctlLib.ProgressBar PBar 
         Height          =   210
         Left            =   30
         TabIndex        =   8
         Top             =   210
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   370
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Refreshing Chart Of Accounts ..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   9
         Top             =   0
         Width           =   2385
      End
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   390
      Top             =   630
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowBorderStyle=   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.PictureBox PicEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1275
      ScaleHeight     =   225
      ScaleWidth      =   5805
      TabIndex        =   6
      Top             =   1770
      Visible         =   0   'False
      Width           =   5835
      Begin VB.TextBox txtBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "RptTime"
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
         Height          =   225
         Left            =   3540
         MaxLength       =   12
         TabIndex        =   1
         Top             =   0
         Width           =   1680
      End
      Begin VB.TextBox txtAccTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "RptTime"
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
         Height          =   225
         Left            =   -15
         MaxLength       =   150
         TabIndex        =   0
         Top             =   0
         Width           =   3510
      End
      Begin VB.ComboBox cmbCrDr 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         ItemData        =   "frmChartOfAccounts.frx":0000
         Left            =   5250
         List            =   "frmChartOfAccounts.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   -45
         Width           =   600
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   105
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChartOfAccounts.frx":0016
            Key             =   "ACC"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChartOfAccounts.frx":046A
            Key             =   "PARIN"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChartOfAccounts.frx":08BE
            Key             =   "ACCIN"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChartOfAccounts.frx":0D12
            Key             =   "PAR"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChartOfAccounts.frx":1166
            Key             =   "TYPE"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV 
      Height          =   7845
      Left            =   60
      TabIndex        =   3
      Top             =   15
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   13838
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   15199215
      BorderStyle     =   1
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Account #"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Account Title"
         Object.Width           =   10583
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Opening Bal."
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Current Bal."
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Reporting Group"
         Object.Width           =   4128
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   480
      Left            =   105
      TabIndex        =   10
      Top             =   8070
      Visible         =   0   'False
      Width           =   5475
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   1965
         TabIndex        =   13
         Top             =   135
         Width           =   2430
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Enabled         =   0   'False
         Height          =   300
         Left            =   4395
         TabIndex        =   12
         Top             =   135
         Width           =   1020
      End
      Begin VB.ComboBox cmbfield 
         Height          =   315
         ItemData        =   "frmChartOfAccounts.frx":15BA
         Left            =   60
         List            =   "frmChartOfAccounts.frx":15D3
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   135
         Width           =   1875
      End
   End
   Begin MSForms.CommandButton cmdOpt 
      Height          =   390
      Left            =   8670
      TabIndex        =   4
      Top             =   8025
      Width           =   1560
      Caption         =   "Options         "
      PicturePosition =   327683
      Size            =   "2752;688"
      Accelerator     =   79
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   390
      Left            =   10275
      TabIndex        =   5
      Top             =   8025
      Width           =   1545
      ForeColor       =   64
      Caption         =   "Close          "
      PicturePosition =   327683
      Size            =   "2725;688"
      Accelerator     =   67
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
      Begin VB.Menu mnurefresh 
         Caption         =   "Re&fresh"
      End
      Begin VB.Menu mnuShowAll 
         Caption         =   "Show &Inactive Accounts"
      End
      Begin VB.Menu mnuLedger 
         Caption         =   "View &Ledger For Last 30 Days"
      End
      Begin VB.Menu dash0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewAcc 
         Caption         =   "&New Account"
      End
      Begin VB.Menu mnuDelAcc 
         Caption         =   "Delete Account"
      End
      Begin VB.Menu mnuEditAcc 
         Caption         =   "&Edit Account"
      End
      Begin VB.Menu dashh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRelatedAcc 
         Caption         =   "Add Related Account Info"
      End
      Begin VB.Menu dash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActive 
         Caption         =   "Make This Account Ina&ctive"
      End
      Begin VB.Menu mnuParent 
         Caption         =   "Make This Account Paren&t"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTransfer 
         Caption         =   "&Transfer Funds"
      End
      Begin VB.Menu mnuJV 
         Caption         =   "Make &Journal Entry"
      End
      Begin VB.Menu dash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print Chart of Accounts"
      End
      Begin VB.Menu mnuPrintWithOpening 
         Caption         =   "Print Chart of Accounts With &Opening"
      End
      Begin VB.Menu mnuPrintBal 
         Caption         =   "Print Current &Balances Of Accounts"
      End
   End
End
Attribute VB_Name = "frmChartOfAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iColNo As Integer

Dim Loaded As Boolean
Dim Rec As New ADODB.Recordset
Dim rsAccounts As New ADODB.Recordset
Dim i As Integer
Dim Bal As Double
Dim AccList As String
Dim SubAccList As String
Dim TotalBalance As Double
Dim TotalOpenBalance As Double
Dim ShowAll As Boolean

Dim PermissionDenied As Boolean

Private Sub Command2_Click()


    Dim AccNo As String
    Dim AccTitle As String
    Dim Bal As Double

    For i = 1 To LV.ListItems.count
        AccNo = LV.ListItems(i).Text & ""
        AccTitle = LV.ListItems(i).ListSubItems(1).Text
        Bal = Val(LV.ListItems(i).ListSubItems(2).Tag)
        con.Execute "Delete * From TempChaccount"
        con.Execute "Insert into TempChAccount Values('" & AccNo & "','" & AccTitle & "'," & Bal & ")"
    Next i
    
    
End Sub




Private Sub RefreshList()
Screen.MousePointer = vbHourglass
LV.Sorted = False

Dim PrevIndex As Double

If LV.ListItems.count > 0 Then PrevIndex = LV.SelectedItem.Index


    'LV.Visible = False
    LV.ListItems.Clear
    Dim i As Integer
    Screen.MousePointer = vbHourglass


With rsAccounts
    .MoveFirst
    
    ''''''''''Show Pregress Bar'''''''''''''
      pbar.Value = 0
      Frame1.Visible = False
      Me.Refresh
      PicStatus.Visible = True
      
      pbar.Max = .RecordCount + 25
      pbar.Refresh
      PicStatus.Refresh
    ''''''''''''''''''''''''''''''''''''''''
    
    
    For i = 0 To .RecordCount - 1
        TotalBalance = 0
        TotalOpenBalance = 0
        AccList = ""
        
        If ![Parent] Then
            Call CalculateBalance(.Fields("AccNo"))
        Else
            TotalBalance = Val(![Balance] & "")
            TotalOpenBalance = Val(![OpenBal] & "")
        End If
        
        Call AddToList(i, .Fields("SubaccOf") & "", .Fields("AccNo"), .Fields("AccTitle"), TotalOpenBalance, TotalBalance, .Fields("AccNo"), ![Parent], ![Active])
        .MoveNEXT
      
        pbar.Value = i
        
    Next i
End With

    Call AddHeads
   
    Frame1.Visible = True
    Me.Refresh
    PicStatus.Visible = False
    
   
    'LV.Visible = True
    LV.Sorted = True
    If PrevIndex > 0 And PrevIndex <= LV.ListItems.count Then LV.ListItems(PrevIndex).Selected = True
    
    On Error Resume Next
        LV.SelectedItem.EnsureVisible
        
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmbCrDr_Click()
 If cmbCrDr.Text = "CR" Then
  cmbCrDr.Tag = -Val(txtBal.Text)
 Else
  cmbCrDr.Tag = Val(txtBal.Text)
 End If
End Sub

Private Sub cmbCrDr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call UpdateAccount
End If
End Sub

Private Sub cmbReportingGroup_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    If KeyCode.Value = 13 Then
    
        If cmbReportingGroup.MatchFound = False Then
            MsgBox "Please Select from the List.", vbInformation
            Exit Sub
        End If
        
        Dim iAffect As Integer, iGroupNo As Integer
        
        If cmbReportingGroup.ListIndex = 0 Then ' Means --------
            con.Execute "DELETE FROM AccountGroupsForExpenseReport WHERE AccNo='" & LV.SelectedItem.Text & "'"
            LV.SelectedItem.SubItems(iColNo - 1) = ""
            PicEditRG.Visible = False
            Exit Sub
        End If
        
        iGroupNo = cmbReportingGroup.List(cmbReportingGroup.ListIndex, 1)
        con.Execute "UPDATE AccountGroupsForExpenseReport SET AccGroup_RefID=" & _
         iGroupNo & " WHERE AccNo='" & LV.SelectedItem.Text & "'", iAffect
        
        If iAffect = 0 Then
            con.Execute "INSERT INTO AccountGroupsForExpenseReport(AccGroup_RefID,AccNo) VALUES(" & _
             iGroupNo & ",'" & LV.SelectedItem.Text & "')"
        End If
        LV.SelectedItem.SubItems(iColNo - 1) = cmbReportingGroup.Text
        PicEditRG.Visible = False
        
    ElseIf KeyCode.Value = 27 Then
        PicEditRG.Visible = False
    End If
    
End Sub


Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdFind_Click()

    Dim StartIdx As Integer

    If LV.SelectedItem Is Nothing Then
       StartIdx = 1
    Else
       If LV.SelectedItem.Index = LV.ListItems.count Then
          StartIdx = 1
       Else
          StartIdx = LV.SelectedItem.Index
       End If
    End If

    Dim fldValue As String




    For i = StartIdx To LV.ListItems.count
       
       
       If cmbfield.ListIndex = 0 Then
          fldValue = Trim(Replace(LV.ListItems(i), "»", ""))
       Else
          fldValue = Trim(Replace(LV.ListItems(i).SubItems(cmbfield.ListIndex), "»", ""))
       End If
    
       If LCase(fldValue) Like "*" & LCase(Trim(txtFind)) & "*" Then
          LV.ListItems(i).Selected = True
          LV.ListItems(i).EnsureVisible
          If i <> StartIdx Then Exit For
       End If
    Next
    
    If i = LV.ListItems.count + 1 Then
        MsgBox "No Mathes Found.", vbInformation
        Set LV.SelectedItem = Nothing
    End If

    
    
End Sub



Private Sub cmdOpt_Click()
With cmdOpt
 Call CheckMenu
 Me.PopUpMenu mnuPop, 255, .Left, .Top + .Height
End With
End Sub

Private Sub Form_Activate()

    'Check For User Access
    If PermissionDenied Then
        Me.Hide
        MsgBox "You Donot Have Permissions To Use This Option. Please Contact Your Administrator", vbInformation
        Unload Me
        Exit Sub
    End If

    If Not Loaded Then
        Call RefreshList
        Loaded = True
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 And PicEdit.Visible Then
        PicEdit.Visible = False
    ElseIf KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()

    'PermissionDenied = Not UserHasAccess("ChartofAccounts")
    mnuPop.Visible = False

    Me.KeyPreview = True

    Dim Condition As String
    
    If ShowAll Then
       Condition = ""
    Else
       Condition = " Where Active=1 "
    End If

    cmbfield.Clear
    For i = 1 To LV.ColumnHeaders.count
        cmbfield.AddItem LV.ColumnHeaders(i).Text
    Next
    cmbfield.ListIndex = 1
    
    cmbReportingGroup.AddItem "------------"
    Call AddToCombo(cmbReportingGroup, "GroupName", "GroupsForExpenseReport", " ORDER BY SortNo", True, "EntryID")
    
   Set LV.SmallIcons = ImageList1
   rsAccounts.Open "SELECT * FROM Accounts " & Condition & " ORDER BY Accno", con, adOpenForwardOnly
   If Loaded Then Call RefreshList
   
End Sub
Private Sub AddToList(No As Integer, Parent As String, code As String, AccTitle As String, OpenBal, Bal As Double, key As String, ParentAccount As Boolean, Active As Boolean)

    Dim LColor As Long
    Dim NumSpaces As Integer
    Dim TheTag As String
    
    ' Format Opening Balance
    
    Dim Openbalance As String


    Openbalance = Format(OpenBal, "#,##0.00 DR;#,##0.00 CR;#,##0.00      ")


    'Format current Balance


    Dim CurrBalance As String

    CurrBalance = Format(Bal, "#,##0.00 DR;#,##0.00 CR;#,##0.00      ")


    If Not Active Then
        LColor = &H808080
        TheTag = "Inactive"
    Else
        LColor = &H800000
        TheTag = "Active"
    End If

    'For Image Selection
    Dim img As String
    If ParentAccount Then
        If Active Then
            img = "PAR"
        Else
            img = "PARIN"
        End If
    Else
        If Active Then
            img = "ACC"
        Else
            img = "ACCIN"
        End If
    End If
   
    Dim i As Integer

    For i = 1 To LV.ListItems.count
        If Parent = Replace(LV.ListItems(i).key, "[", "") Then
            NumSpaces = Len(LV.ListItems(i).ListSubItems(1).Text) - Len(Trim(LV.ListItems(i).ListSubItems(1).Text))
            LV.ListItems.add i + 1, "[" & key, code, , img
            LV.ListItems(i + 1).Bold = ParentAccount
            LV.ListItems(i + 1).Tag = TheTag
            LV.ListItems(i + 1).ForeColor = LColor
            LV.ListItems(i + 1).ListSubItems.add , , Space$(NumSpaces + 8) & "»  " & AccTitle
            LV.ListItems(i + 1).ListSubItems(1).ForeColor = LColor
            LV.ListItems(i + 1).ListSubItems(1).Bold = ParentAccount
            '        LV.ListItems(i + 1).ListSubItems(1).ReportIcon = img
        
            LV.ListItems(i + 1).ListSubItems.add , , Openbalance
            LV.ListItems(i + 1).ListSubItems(2).Tag = OpenBal
            LV.ListItems(i + 1).ListSubItems(2).ForeColor = LColor
            LV.ListItems(i + 1).ListSubItems(2).Bold = ParentAccount
            LV.ListItems(i + 1).ListSubItems.add , , "" 'CurrBalance
            LV.ListItems(i + 1).ListSubItems(3).ForeColor = LColor
            LV.ListItems(i + 1).ListSubItems(3).Tag = Bal
            LV.ListItems(i + 1).ListSubItems(3).Bold = ParentAccount
        
            LV.ListItems(i + 1).SubItems(4) = GetSingleStringValue("GroupName", "VAccountGroupsForExpenseReport", " WHERE AccNo='" & code & "'")
            Exit Sub
        End If
    Next i


    LV.ListItems.add , "[" & key, code, , img
    LV.ListItems(No + 1).Bold = ParentAccount
    LV.ListItems(No + 1).Tag = TheTag
    LV.ListItems(No + 1).ForeColor = LColor
    LV.ListItems(No + 1).ListSubItems.add , , "      »  " & AccTitle
    LV.ListItems(No + 1).ListSubItems(1).ForeColor = LColor
    LV.ListItems(No + 1).ListSubItems(1).Bold = ParentAccount
    'LV.ListItems(No + 1).ListSubItems(1).ReportIcon = img
    
    LV.ListItems(No + 1).ListSubItems.add , , Openbalance
    LV.ListItems(No + 1).ListSubItems(2).Bold = ParentAccount
    LV.ListItems(No + 1).ListSubItems(2).ForeColor = LColor
    LV.ListItems(No + 1).ListSubItems(2).Tag = OpenBal
    
    LV.ListItems(No + 1).ListSubItems.add , , "" 'CurrBalance
    LV.ListItems(No + 1).ListSubItems(3).Bold = ParentAccount
    LV.ListItems(No + 1).ListSubItems(3).ForeColor = LColor
    LV.ListItems(No + 1).ListSubItems(3).Tag = Bal
    
    LV.ListItems(No + 1).SubItems(4) = GetSingleStringValue("GroupName", "VAccountGroupsForExpenseReport", " WHERE AccNo='" & code & "'")

End Sub

Private Sub CreateLists(PRec As ADODB.Recordset, AccNo As String)

'If Len(AccList) > 0 Then AccList = AccList & "'" & AccNo & "',"
AccList = AccList & "'" & AccNo & "',"
SubAccList = ""
Dim i As Integer
With PRec
    For i = 0 To .RecordCount - 1
        AccList = AccList & "'" & .Fields("AccNo") & "',"
        SubAccList = SubAccList & "'" & .Fields("AccNo") & "',"
        .MoveNEXT
    Next i
    
End With
If Right$(SubAccList, 1) = "," Then SubAccList = Left$(SubAccList, Len(SubAccList) - 1)

End Sub

Private Sub CalculateBalance(AccNo As String)

Dim i As Integer
With Rec

    If .State = 1 Then .Close
    .CursorLocation = adUseClient
    '.Open "Select AccNo from Accounts Where SubAccOf='" & AccNo & "'", con, adOpenForwardOnly, adLockReadOnly
    'Do While .RecordCount > 0
    '    Call CreateLists(rec, AccNo)
    '    If .State = 1 Then .Close
    '.CursorLocation = adUseClient
    'If Right$(SubAccList, 1) = "," Then
    '    Useable = Left$(SubAccList, Len(SubAccList) - 1)
    'Else
    '    Useable = SubAccList
    'End If
    '.Open "Select AccNo from Accounts Where SubAccOf in (" & SubAccList & ")", con, adOpenForwardOnly, adLockReadOnly
    'Loop
    'If Len(AccList) > 0 Then
    '    AccList = Left$(AccList, Len(AccList) - 1)
    'Else
    '    AccList = "'" & AccNo & "'"
    'End If
    'If .State = 1 Then .Close
    '.CursorLocation = adUseClient
    '.Open "Select Sum(Balance),Sum(OpenBal) from Accounts Where AccNo in (" & AccList & ")", con, adOpenForwardOnly, adLockReadOnly
   
    .Open "Select Sum(Balance),Sum(OpenBal) from Accounts Where Ltrim(AccNo) Like '" & AccNo & "%'", con, adOpenForwardOnly, adLockReadOnly
    
    TotalBalance = Val(.Fields(0) & "")
    TotalOpenBalance = Val(.Fields(1) & "")
    '.Close
End With
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next

    LV.Move 50, 50, Me.ScaleWidth - 100, Me.ScaleHeight - (cmdClose.Height + 100)
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, LV.Top + LV.Height + 50
    'cmdPrint.Move cmdClose.Left - cmdPrint.Width - 50, cmdClose.Top
    cmdOpt.Move cmdClose.Left - cmdOpt.Width - 50, cmdClose.Top

    PicStatus.Move LV.Left, LV.Top + LV.Height + 25
    Frame1.Move LV.Left, LV.Top + LV.Height - 25
    
End Sub

Private Sub Form_Unload(cancel As Integer)
rsAccounts.Close
Set rsAccounts = Nothing
Set frmChartOfAccounts = Nothing
End Sub

Private Sub LV_DblClick()
    
    If iColNo = 5 Then
        If LV.SelectedItem.Bold Then Exit Sub
        PicEditRG.Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top
        cmbReportingGroup.Text = LV.SelectedItem.SubItems(iColNo - 1)
        PicEditRG.Visible = True
        cmbReportingGroup.SetFocus
    Else
        Dim lCount As Long
        lCount = GetSingleLongValue("COUNT(*)", "FinancialClosings")
        If lCount > 0 Then
            Exit Sub
        End If
        If OutOfFinancialYear Then Exit Sub
        
        'Check For User Access
        If Not UserHasAccess("EditAccount") Then
           MsgBox "You Don't Have Permissions To Edit Account", vbInformation
           Exit Sub
        End If
        
        
        With LV
        
           If .ListItems.count = 0 Then Exit Sub
           
           If .SelectedItem.Tag = "Inactive" Then
              MsgBox "Inactive Account Cannot Be Edited.", vbInformation
              Exit Sub
           ElseIf .SelectedItem.SmallIcon = "TYPE" Then
              Exit Sub
           End If
           
           txtBal.Enabled = .SelectedItem.SmallIcon <> "PAR"
           cmbCrDr.Enabled = .SelectedItem.SmallIcon <> "PAR"
           
           
           PicEdit.Move .Left + .ColumnHeaders(2).Left + 25, LV.Top + .SelectedItem.Top + 40, .ColumnHeaders(2).Width + .ColumnHeaders(3).Width
           txtAccTitle.Move 0, 0, .ColumnHeaders(2).Width
           txtBal.Move txtAccTitle.Width + 25, 0, .ColumnHeaders(3).Width - cmbCrDr.Width - 15
           cmbCrDr.Move PicEdit.Width - cmbCrDr.Width
           
           txtAccTitle = Trim(Replace(.SelectedItem.SubItems(1), "»", ""))
           
           txtAccTitle.FontBold = .SelectedItem.Bold
           
           txtBal = Val(Replace(.SelectedItem.SubItems(2), ",", ""))
           'FORMAT AND STORE IN TAG
           Call txtBal_LostFocus
           Call FindInCombo(cmbCrDr, Right(Trim(.SelectedItem.SubItems(2)), 2))
           'STORE PREVIOUS OPENING BALANCE
           PicEdit.Tag = Val(.SelectedItem.ListSubItems(2).Tag)
        End With
        PicEdit.Visible = True
        PicEdit.ZOrder
        txtAccTitle.SetFocus
    End If
End Sub

Private Sub LV_GotFocus()
   If PicEdit.Visible Then Call Form_KeyPress(27)
End Sub

Private Sub LV_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If LV.ListItems.count = 0 Then Exit Sub
   If LV.SelectedItem.SmallIcon = "TYPE" Then Exit Sub
   Call CheckMenu
   Me.PopUpMenu mnuPop, , (LV.Left + (LV.Width / 2)), (LV.Top + LV.SelectedItem.Top)
End If
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        If x > LV.ColumnHeaders(i).Left And x < LV.ColumnHeaders(i).Left + LV.ColumnHeaders(i).Width Then
            iColNo = i
            Exit For
        End If
    Next
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 2 Or LV.ListItems.count = 0 Then Exit Sub
'If LV.SelectedItem.SmallIcon = "TYPE" Then Exit Sub

 Call CheckMenu
 Me.PopUpMenu mnuPop
 
End Sub



Private Sub mnuActive_Click()

   If mnuActive.Tag = "Active" Then
      If Not MakeAccountInActive Then Exit Sub
   ElseIf mnuActive.Tag = "Inactive" Then
      If Not MakeAccountActive Then Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass

   Call mnuRefresh_Click
   Screen.MousePointer = vbDefault
End Sub


Private Sub mnuDelAcc_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    
    With LV.SelectedItem
    
    
        With con.Execute("Select Count(*) From FixedAccounts Where AccNo = '" & .Text & "'")
            If Val(.Fields(0) & "") > 0 Then
                MsgBox "Cannot Delete This Account, This Account Has A Fixed Account Number.", vbCritical
                .Close
                Exit Sub
            End If
            .Close
        End With
        
        
        
        If .Bold Then
            With con.Execute("Select Count(*) From Accounts Where Subaccof='" & .Text & "'")
                If Val(.Fields(0) & "") > 0 Then
                    MsgBox "Cannot Delete This Account, This Account Has Some Child Accounts", vbCritical
                    .Close
                    Exit Sub
                End If
                .Close
            End With
        End If
        
        With con.Execute("Select Count(*) From VLedger Where AccNo Like'" & .Text & "%'")
            If Val(.Fields(0) & "") > 0 Then
                MsgBox "Cannot Delete This Account, There is Some Transaction Related To This Account", vbCritical
                .Close
                Exit Sub
            End If
            .Close
        End With
        
        If MsgBox("Are You Sure To Remove This Account ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
        'Check If It's Maker Account
        Dim lCount As Long, lIssCount As Long
        lCount = GetSingleLongValue("COUNT(*)", "Makers", " WHERE AccNo='" & .Text & "'")
        If lCount > 0 Then
            'Now Check If there is any Issuance against this Maker...
            lIssCount = GetSingleLongValue("COUNT(*)", "VendIssued INNER JOIN Makers ON VendIssued.VendID=Makers.VendID", " WHERE AccNo='" & .Text & "'")
            If lIssCount > 0 Then
                MsgBox "Can't Delete, This Account is Linked with Maker Account Who has Some PO(s).", vbInformation
                Exit Sub
            End If
        End If
        con.Execute "Delete from Accounts Where ACCNo='" & .Text & "'", a
        If lCount > 0 Then
            con.Execute "DELETE FROM Makers WHERE AccNo='" & .Text & "'"
        End If
        
        If a = 1 Then
            LV.ListItems.Remove (LV.SelectedItem.Index)
            MsgBox "Account Deleted Successfully", vbInformation
        End If
        
    End With
    mnuRefresh_Click
    
End Sub

Private Sub mnuEditAcc_Click()
   Call LV_DblClick
End Sub

Private Sub mnuJV_Click()

    If LV.SelectedItem.Tag = "Inactive" Then
       MsgBox "No Transaction Can Be Made In An Inactive Account.", vbInformation
       Exit Sub
    End If

Dim f As Form
Select Case mnuJV.Tag
 Case "JV"
   Set f = New frmJV
 Case "RV"
  Set f = New frmCashRV
 Case "PV"
  Set f = New frmCashPV
 Case "EXP"
  Set f = New frmExpenses
 Case Else
  Set f = New frmJV
End Select
   
 
   With f
      .Show
      Call FindInCombo(.cmbAccNo, LV.SelectedItem)
   End With
   
'End If
End Sub

Private Sub mnuLedger_Click()
    Dim f As New frmLedger
    With f
        f.Show
        Call FindInCombo(.cmbAccNo, LV.SelectedItem)
        Call .cmdRefresh_Click
    End With
End Sub

Private Sub mnuNewAcc_Click()
    With frmNewAccount
        .ShowWithDetail (LV.SelectedItem.Text)
    End With
   Call mnuRefresh_Click
End Sub

Private Sub mnuParent_Click()
    If LV.SelectedItem.SmallIcon Like "PAR*" Then
        If MsgBox("Are You Sure To Make This Account Sub Account", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        con.Execute "update accounts set Parent=0 Where AccNo='" & LV.SelectedItem & "'"
        Call mnuRefresh_Click
    Else
        If MsgBox("Are You Sure To Make This Account A Master Account", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        con.Execute "update accounts set Parent=1 Where AccNo='" & LV.SelectedItem & "'"
        Call mnuRefresh_Click
    End If
End Sub

Private Sub mnuPrint_Click()
    Call PrintChartOfAccounts
End Sub

Private Sub PrintChartOfAccounts(Optional bOpening As Boolean = False)

    Dim Inactive As Boolean
    'Screen.MousePointer = vbHourglass
    con.BeginTrans
    con.Execute "Delete From TempAccounts"
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .SmallIcon = "PARIN" Or .SmallIcon = "ACCIN" Then Inactive = True Else Inactive = False
            con.Execute "Insert into TempAccounts(AccNo,AccTitle,Type,OpenBal,Balance,Bold,Inactive) " & _
            "Values('" & .Text & "','" & .SubItems(1) & "','" & Left(.Text, 2) & "','" & _
            .ListSubItems(2) & "','" & .ListSubItems(3) & "'," & Abs(.Bold) & "," & Abs(Inactive) & ")"
        End With
     Next
    
    con.CommitTrans
    
    Dim rpt As CRAXDDRT.Report, strSelection As String
    If bOpening Then
        Set rpt = rptApp.OpenReport(RptPath & "\AccountsWithOpening.rpt")
        strSelection = "{VTempAccounts.OpenBal}<>'0.00'"
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\Accounts.rpt")
    End If
    rpt.FormulaFields.GetItemByName("Company").Text = "'" & strCompany & "'"
    rpt.FormulaFields.GetItemByName("Caption").Text = "'Chart Of Accounts'"
    rpt.FormulaFields.GetItemByName("PrintBalance").Text = False
    rpt.FormulaFields.GetItemByName("Address").Text = "'" & COMPANYADDRESS & "'"
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
'    With cr1
'        .Connect = ConnectStr
'        If bOpening Then
'            .ReportFileName = RptPath & "\AccountsWithOpening.rpt"
'            .ReplaceSelectionFormula "{VTempAccounts.OpenBal}<>'0.00'"
'        Else
'            .ReportFileName = RptPath & "\Accounts.rpt"
'        End If
'        .Formulas(0) = "Company='" & strCompany & "'"
'        .Formulas(1) = "Caption='Chart Of Accounts'"
'        .Formulas(2) = "PrintBalance=False"
'        .Formulas(3) = "Address='" & COMPANYADDRESS & "'"
'        .Action = 1
'        .PageZoomNext
'    End With
'
'    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuPrintBal_Click()

    Dim Inactive As Boolean
    'Screen.MousePointer = vbHourglass
    con.BeginTrans
    con.Execute "Delete From TempAccounts"
     For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .SmallIcon = "PARIN" Or .SmallIcon = "ACCIN" Then Inactive = True Else Inactive = False
            con.Execute "Insert into TempAccounts Values('" & .Text & "','" & .SubItems(1) & _
                        "','" & Left(.Text, 2) & "','" & .ListSubItems(2) & "','" & .ListSubItems(3) & "'," & Abs(.Bold) & "," & Abs(Inactive) & ")"
        End With
     Next
    
    con.CommitTrans
    
    Dim rpt As CRAXDDRT.Report, strSelection As String
    
    Set rpt = rptApp.OpenReport(RptPath & "\Accounts.rpt")
    
    rpt.FormulaFields.GetItemByName("Company").Text = "'" & strCompany & "'"
    rpt.FormulaFields.GetItemByName("Caption").Text = "'Current Balances of Accounts'"
    rpt.FormulaFields.GetItemByName("PrintBalance").Text = False
    rpt.FormulaFields.GetItemByName("Address").Text = "'" & COMPANYADDRESS & "'"
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
'    With cr1
'        .Connect = ConnectStr
'        .ReportFileName = RptPath & "\Accounts.rpt"
'        .Formulas(0) = "Company='" & strCompany & "'"
'        .Formulas(1) = "Caption='Current Balances of Accounts'"
'        .Formulas(2) = "PrintBalance=True"
'        .Formulas(3) = "Address='" & COMPANYADDRESS & "'"
'        .Action = 1
'        .PageZoomNext
'    End With
'
'    Screen.MousePointer = vbDefault
    
End Sub

Private Sub mnuPrintWithOpening_Click()
    Call PrintChartOfAccounts(True)
End Sub

Private Sub mnuRefresh_Click()
   Screen.MousePointer = vbHourglass
   rsAccounts.Close
   Set rsAccounts = Nothing
   Call Form_Load
   Screen.MousePointer = vbDefault
End Sub


Private Sub CheckMenu()

If OutOfFinancialYear Then

     mnuEditAcc.Enabled = False
     mnuTransfer.Enabled = False
     mnuJV.Enabled = False
     mnuLedger.Enabled = True
     mnuPrint.Enabled = True
     mnuPrintBal.Enabled = True
     mnuParent.Enabled = False
     mnuNewAcc.Enabled = False
     mnuEditAcc.Enabled = False
     mnuActive.Enabled = False
     mnuRelatedAcc.Enabled = False

   Exit Sub
End If



If LV.SelectedItem.SmallIcon = "TYPE" Then
     mnuEditAcc.Enabled = False
     mnuTransfer.Enabled = False
     mnuJV.Enabled = False
     mnuLedger.Enabled = False
     mnuPrint.Enabled = True
     mnuParent.Enabled = False
     mnuNewAcc.Enabled = True
     mnuEditAcc.Enabled = False
     mnuActive.Enabled = False
     mnuRelatedAcc.Enabled = False
     Exit Sub
End If






   If LV.ListItems.count = 0 Then
     mnuEditAcc.Enabled = False
     mnuTransfer.Enabled = False
     mnuJV.Enabled = False
     mnuLedger.Enabled = False
     mnuPrint.Enabled = True
     mnuPrintBal.Enabled = True
     mnuParent.Enabled = False
     
   Else
     mnuEditAcc.Enabled = LV.SelectedItem.SmallIcon <> "TYPE"
     mnuTransfer.Enabled = True
     mnuJV.Enabled = True
     
     If LV.SelectedItem Like "15-003-*" Then
      mnuJV.Caption = "Make &Receipt Entry"
      mnuJV.Tag = "RV"
     ElseIf LV.SelectedItem Like "24-004-*" Then
      mnuJV.Caption = "Make &Payment Entry"
      mnuJV.Tag = "PV"
     ElseIf LV.SelectedItem Like "4*" Then
      mnuJV.Caption = "Make &Expense Entry"
      mnuJV.Tag = "EXP"
     Else
      mnuJV.Caption = "Make &Journal Entry"
      mnuJV.Tag = ""
     End If
     
     mnuLedger.Enabled = Not (LV.SelectedItem.SmallIcon = "PAR" Or LV.SelectedItem.SmallIcon = "PARIN" Or LV.SelectedItem.SmallIcon = "TYPE")
    
     
     Dim RecCount As Long
     If LV.SelectedItem.SmallIcon Like "PAR*" Then
         mnuParent.Caption = "Make This Account A Sub Account"
         RecCount = Val(con.Execute("Select Count(*) From Accounts Where Ltrim(SubAccof) like '" & LV.SelectedItem & "%'").Fields(0) & "")
     Else
         mnuParent.Caption = "Make This Account A Master Account"
         RecCount = Val(con.Execute("Select Count(*) From Vouchers Where Accno='" & LV.SelectedItem & "'").Fields(0) & "")
     End If
     mnuParent.Enabled = (RecCount = 0) And (Val(LV.SelectedItem.SubItems(3)) = 0) And (Val(LV.SelectedItem.SubItems(3)) = 0)
     
     If LV.SelectedItem.Tag = "Inactive" Then
         mnuEditAcc.Enabled = False
         mnuTransfer.Enabled = False
         mnuJV.Enabled = False
         mnuLedger.Enabled = False
         mnuParent.Enabled = False
       mnuActive.Caption = "Make This Account A&ctive."
     Else
       mnuActive.Caption = "Make This Account Ina&ctive."
     End If
     mnuActive.Tag = LV.SelectedItem.Tag
     mnuActive.Enabled = LV.SelectedItem.SmallIcon <> "TYPE"
     mnuPrint.Enabled = True
     mnuPrintBal.Enabled = True
     
   End If
   
'Related Accounts Check

If Not LV.SelectedItem.SmallIcon Like "ACC*" Then Exit Sub
Select Case Left(LV.SelectedItem, 2)
Case "11"
      mnuRelatedAcc.Caption = "Edit Related Asset Accounts"
      mnuRelatedAcc.Tag = "Asset"
      mnuRelatedAcc.Enabled = True
Case "41", "42", "43", "44"

If Left(LV.SelectedItem, 2) = "41" And Not LV.SelectedItem Like "41-003" Then GoTo NoRelatedAcc

   If Val(con.Execute("Select count(*) from ExpenseAccounts Where ExpAccno='" & LV.SelectedItem & "'").Fields(0) & "") = 0 Then
      mnuRelatedAcc.Caption = "Add Related Expense Accounts"
      mnuRelatedAcc.Tag = "ExpAdd"
   Else
      mnuRelatedAcc.Caption = "Edit Related Expense Accounts"
      mnuRelatedAcc.Tag = "ExpEdit"
   End If
   mnuRelatedAcc.Enabled = True
   
Case Else
NoRelatedAcc:
   mnuRelatedAcc.Enabled = False
      
End Select








End Sub

Private Sub mnuRelatedAcc_Click()
'Check For User Access
If Not UserHasAccess("EditVoucherDesc") Then
   MsgBox "You Don't Have Permissions To Edit Account", vbInformation
   Exit Sub
End If


Load frmRelatedAcc
With frmRelatedAcc
   Select Case mnuRelatedAcc.Tag
      Case "ExpEdit"
         .RelatedExpenseAccount = True
         .RelatedAssetAccount = False
         .AddNew = False
         .edit = True
         .Tag = LV.SelectedItem
      Case "ExpAdd"
         .RelatedExpenseAccount = True
         .RelatedAssetAccount = False
         .AddNew = True
         .edit = False
         .Tag = LV.SelectedItem
      Case "Asset"
         .RelatedExpenseAccount = False
         .RelatedAssetAccount = True
         .AddNew = False
         .edit = True
         .Tag = LV.SelectedItem

   End Select
   
   .Show 1
   Call mnuRefresh_Click
End With

End Sub

Private Sub mnuShowAll_Click()
mnuShowAll.Checked = Not mnuShowAll.Checked
ShowAll = mnuShowAll.Checked
Call mnuRefresh_Click
End Sub

Private Sub mnuTransfer_Click()
   frmTransferFunds.Show 1
End Sub



Private Sub txtAccTitle_GotFocus()
txtAccTitle.SelStart = 0
txtAccTitle.SelLength = Len(txtAccTitle)
End Sub

Private Sub txtAccTitle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then If Not txtBal.Enabled Then Call UpdateAccount

End Sub


Private Sub txtBal_GotFocus()
txtBal.SelStart = 0
txtBal.SelLength = Len(Bal)
End Sub

Private Sub txtBal_KeyPress(KeyAscii As Integer)
   KeyAscii = OnlyNumber(KeyAscii, True)
End Sub

Private Sub txtBal_LostFocus()
 txtBal = Format(Val(txtBal), "###0.00")
 txtBal.Tag = LV.SelectedItem.ListSubItems(2).Tag
 Call cmbCrDr_Click
End Sub



Private Function MakeAccountInActive() As Boolean
MakeAccountInActive = False
 With LV.SelectedItem
        
        With con.Execute("Select Count(*) From FixedAccounts Where AccNo = '" & .Text & "'")
            If Val(.Fields(0) & "") > 0 Then
                MsgBox "Cannot Make This Account Inactive, This Account Has A Fixed Account Number.", vbCritical
                .Close
                Exit Function
            End If
            .Close
        End With
        
   If .ListSubItems(3).Tag <> 0 Then
        MsgBox "This Account Cannot Be Made Inactive." & vbNewLine & _
           "To Make An Account Inactive Balance Must Be Zero.", vbInformation
         Exit Function
   ElseIf .SmallIcon = "PAR" Then
       If MsgBox("This Is A Master Account, All Sub Accounts Will Also Become Inactive." & vbNewLine & _
           "Are You Sure to Make '" & Trim(Replace(LV.SelectedItem.SubItems(1), "»", "")) & "' Inactive", vbYesNo + vbQuestion) = vbNo Then Exit Function
   Else
      If MsgBox("Are You Sure to Make '" & Trim(Replace(LV.SelectedItem.SubItems(1), "»", "")) & "' Inactive", vbYesNo + vbQuestion) = vbNo Then Exit Function
   End If
   
   
   
   
   con.Execute "Update Accounts Set Active=0 Where Accno='" & Trim(.Text) & _
               "' OR SubAccOf Like '" & Trim(.Text) & "%'"
  
   MakeAccountInActive = True
End With
Call mnuRefresh_Click
End Function


Private Function MakeAccountActive() As Boolean

MakeAccountActive = False
Dim AlsoChild As String

 With LV.SelectedItem
   
   If .SmallIcon = "PARIN" Then
       If MsgBox("This Is A Parent Account, Do You Want All Child Accounts To Become Active." _
          , vbYesNo + vbQuestion) = vbYes Then
          AlsoChild = " OR SubAccOf Like '" & Trim(.Text) & "'"
       Else
           AlsoChild = ""
       End If
   End If
   
   
   
   
   con.Execute "Update Accounts Set Active=1 Where Accno='" & Trim(.Text) & "'" & AlsoChild
  
MakeAccountActive = True
End With




End Function



Private Function UpdateAccount() As Boolean
 If txtAccTitle.Text = "" Then
   MsgBox "Invalid Account Title", vbInformation
   Exit Function
 End If
 
Dim Diff As Double

Screen.MousePointer = vbHourglass
Diff = Val(cmbCrDr.Tag) - Val(txtBal.Tag)
   con.Execute "Update Accounts Set Acctitle='" & txtAccTitle & "',OpenBal=" & Val(cmbCrDr.Tag) & _
               ", Balance=Balance+" & Diff & " Where AccNo='" & LV.SelectedItem.Text & "'"
               
   'Adjust The Ledger
   Call adjustLedger(LV.SelectedItem.Text, YearStartDate, 0, Diff, True)
   
   'Update The Opening Blaance Entry if Vender Or Customer
   If LV.SelectedItem.Text Like "24-004*" Or LV.SelectedItem.Text Like "15-003*" Then
      Call UpdateAdvEntry(LV.SelectedItem.Text, Val(cmbCrDr.Tag), YearStartDate)
   End If
   
   
   'Sleep (500)
   Call mnuRefresh_Click
Screen.MousePointer = vbDefault
   
End Function



Private Sub AddHeads()
Dim rs As New ADODB.Recordset
With rs
   .CursorLocation = adUseClient
   If ShowAll Then
      .Open "Select * from VHeads Order By Code", con, adOpenForwardOnly
   Else
      .Open "Select * from VActiveHeads Order By Code", con, adOpenForwardOnly
   End If
Dim ITM As ListItem, OpenBal As String, Balance As String

   For i = 0 To .RecordCount - 1
            
     
         OpenBal = Format(Val(![OpenBal] & ""), "#,##0.00 DR;#,##0.00 CR;#,##0.00      ")
         Balance = Format(Val(![Balance] & ""), "#,##0.00 DR;#,##0.00 CR;#,##0.00      ")
      
      
      Set ITM = LV.ListItems.add(, ![code] & "'", ![code] & "", , "TYPE")
      ITM.Bold = True
      ITM.ListSubItems.add(, , ![Head]).Bold = True
      'itm.ListSubItems(1).ReportIcon = "TYPE"
      ITM.ListSubItems.add(, , OpenBal).Bold = True
      ITM.ListSubItems.add(, , "").Bold = True
      .MoveNEXT
      
      pbar.Value = pbar.Value + 1
      
   Next
.Close
End With
Set rs = Nothing

End Sub

Private Sub txtFind_Change()
    cmdFind.Enabled = Len(Trim(txtFind)) > 0
End Sub
