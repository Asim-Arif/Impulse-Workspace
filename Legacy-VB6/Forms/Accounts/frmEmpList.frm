VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEmpList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee List"
   ClientHeight    =   6795
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   10365
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
            Picture         =   "frmEmpList.frx":0000
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpList.frx":01BC
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
            Picture         =   "frmEmpList.frx":0378
            Key             =   "Active"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpList.frx":07CC
            Key             =   "Inactive"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   15
      TabIndex        =   2
      Top             =   -90
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee List"
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
         TabIndex        =   3
         Top             =   120
         Width           =   2205
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee List"
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
         TabIndex        =   4
         Top             =   135
         Width           =   2205
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   5295
      Left            =   15
      TabIndex        =   0
      Top             =   495
      Width           =   10320
      _ExtentX        =   18203
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
         Text            =   "Employee Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Emp. Type"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Job"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Phone 1"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Phone 2"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Joining Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Salary"
         Object.Width           =   2293
      EndProperty
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   375
      Left            =   8640
      TabIndex        =   5
      Top             =   6135
      Width           =   1665
      ForeColor       =   64
      Caption         =   "Close          "
      PicturePosition =   327683
      Size            =   "2937;661"
      Picture         =   "frmEmpList.frx":0C20
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
      Left            =   6915
      TabIndex        =   1
      Top             =   6135
      Width           =   1680
      Caption         =   "Options         "
      PicturePosition =   327683
      Size            =   "2963;661"
      Picture         =   "frmEmpList.frx":0D32
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
      Visible         =   0   'False
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuShowAll 
         Caption         =   "Show &Inactive Employees"
      End
      Begin VB.Menu mnuLedger 
         Caption         =   "View Short Term Adv's For &Last 30 Days"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "Add &New Employee"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit Employee Information"
      End
      Begin VB.Menu mnuInActive 
         Caption         =   "Make This Employee Ina&ctive"
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSAdv 
         Caption         =   "Enter Short Term Advance Entry"
      End
      Begin VB.Menu mnuLAdv 
         Caption         =   "Enter Long Term Advance Entry"
      End
      Begin VB.Menu dash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintList 
         Caption         =   "&Print Employee List"
      End
      Begin VB.Menu mnuPrintDetail 
         Caption         =   "Print Employee &Detailed List"
      End
   End
End
Attribute VB_Name = "frmEmpList"
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
 Me.PopupMenu mnupop, 255, .Left, .Top + .Height
End With
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Lv.Sorted = True
Set Lv.ColumnHeaderIcons = ImageList2
If Lv.SortKey = ColumnHeader.Index - 1 Then
   If Lv.SortOrder = lvwAscending Then
      Lv.SortOrder = lvwDescending
      ColumnHeader.Icon = "Up"
   Else
      Lv.SortOrder = lvwAscending
      ColumnHeader.Icon = "Down"
   End If
Else
 Lv.ColumnHeaderIcons = Nothing
 Set Lv.ColumnHeaderIcons = ImageList2
 Lv.SortKey = ColumnHeader.Index - 1
 Lv.SortOrder = lvwAscending
 ColumnHeader.Icon = "Down"
End If

On Error Resume Next
Lv.SelectedItem.EnsureVisible
End Sub

Private Sub Lv_DblClick()
If Lv.ListItems.Count = 0 Then Exit Sub
Call mnuedit_Click
    
End Sub




Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
PermissionDenied = Not UserHasAccess("EmployeeList")
  Call RefreshList
  mnuAdd.Enabled = Not OutOfFinancialYear And UserHasAccess("NewEmployee")
  mnuEdit.Enabled = Not OutOfFinancialYear And UserHasAccess("EditEmployee")
' mnuDelete.Enabled = Not OutOfFinancialYear
End Sub

Private Sub Form_Resize()
 On Error Resume Next
 fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
 
 Dim TitleBottom As Long

 TitleBottom = fTitle.Top + fTitle.Height
 
 Lv.Move 50, TitleBottom, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + cmdClose.Height + 100)
 
 cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 50
 cmdOpt.Move cmdClose.Left - cmdOpt.Width - 75, cmdClose.Top
 
 
End Sub






Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 2 Then Exit Sub
 Call CheckMenu
 Me.PopupMenu mnupop
End Sub

Public Sub mnuAdd_Click()
Dim f As New frmEmp
With f
 .Add = True
 .Show 1
 
End With
mnuRefresh_Click
End Sub


Private Sub mnuClose_Click()
Unload Me
End Sub

Private Sub mnuInActive_Click()
If Lv.ListItems.Count = 0 Then Exit Sub

If mnuInActive.Tag = "Active" Then
    'If Val(LV.SelectedItem.SubItems(6)) <> 0 Then
    '    MsgBox "This Employee Cannot Be Inactive Unless Balance Becomes Zero.", vbInformation
    '    Exit Sub
    'Else
        If MsgBox("Are You Sure To Make '" & Lv.SelectedItem & "' Inactive", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        Con.Execute "Update Accounts Set Active=False Where Accno in (" & Lv.SelectedItem.Tag & ")"
    'End If
Else
    Con.Execute "Update Accounts Set Active=True Where Accno in (" & Lv.SelectedItem.Tag & ")"
End If
  
  Call RefreshList
End Sub


Private Sub mnuedit_Click()
   Dim f As New frmEmp
   With f
      .Edit = True
      .Tag = Val(Lv.SelectedItem.key)
      .Show 1
      mnuRefresh_Click
   End With
End Sub



Private Sub mnuLedger_Click()
Dim EmpAccs() As String
With frmLedger
         EmpAccs = Split(Lv.SelectedItem.Tag, "','")
        .cmbAccNo = Replace(EmpAccs(0), "'", "")
    .cmdRefresh_Click
    .Show
End With
End Sub

Private Sub mnuPrintList_Click()

With cr1
   .ReportFileName = RptPath & "\EmpList.rpt"
   .Connect = ConnectStr
   .Formulas(0) = "Company='" & company & "'"
   .Formulas(1) = "Address='" & COMPANYADDRESS & "'"
   .SelectionFormula = IIf(mnuShowAll.Checked, "", "{VEmps.Active}=True")
   .Action = 1
   .PageZoomNext
End With

End Sub

Private Sub mnuSadv_Click()
    Dim f As New frmEmpAdv
    With f
      .LongTermAdvances = False
        Dim EmpAccs() As String
        Call .Form_Activate      'Add Accounts To The Combo's
        
        'Get Accounts in a List
        EmpAccs = Split(Lv.SelectedItem.Tag, "','")
        .cmbAccNo = Replace(EmpAccs(0), "'", "")
        .Show
        .TBox(1).SetFocus
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


If Lv.ListItems.Count > 0 Then PrevIndex = Lv.SelectedItem.Index
Set Lv.SmallIcons = ImageList1
With rs
    'Check For Show All Option
    ActiveCondition = IIf(mnuShowAll.Checked, "", " Where Active=1")
    
    .CursorLocation = adUseClient
    .Open "Select * From VEmps " & ActiveCondition, Con, adOpenForwardOnly, adLockReadOnly
    
    Lv.ListItems.Clear
    
    For i = 0 To .RecordCount - 1
         'Set Color For Inactive
        If Not ![Active] Then
            TheColor = &H808080
            TheTag = "Inactive"
        Else
            TheColor = &H800000
            TheTag = "Active"
        End If
        
        Set itm = Lv.ListItems.Add(, ![EmpID] & "'", ![EmpName] & "", , TheTag)
        
        itm.ForeColor = TheColor
        
        ' Save Employees Related Account List in Tag To Make Active Or Inactive
        itm.Tag = "'" & ![SAdvAccno] & "','" & ![LAdvAccno] & "','" & ![AccNo] & _
                  "','" & ![AcrAccNo] & "','" & ![OTAcc] & "','" & ![AcrOTAcc] & "'"
        
        
        itm.ListSubItems.Add(, , ![EmpType] & "").ToolTipText = ![EmpType] & ""
        itm.ListSubItems(1).ForeColor = TheColor
        itm.ListSubItems.Add(, , ![Job] & "").ToolTipText = ![Job] & ""
        itm.ListSubItems(2).ForeColor = TheColor
        itm.ListSubItems.Add(, , ![Phone1] & " ").ForeColor = TheColor
        itm.ListSubItems.Add(, , ![Phone2] & " ").ForeColor = TheColor
        itm.ListSubItems.Add(, , Format(![J Date], "dd/MM/yyyy")).ForeColor = TheColor
        itm.ListSubItems.Add(, , Val(![Salary] & "")).ForeColor = TheColor
        .MoveNext
    Next
    .Close
End With
Set rs = Nothing


If PrevIndex > 0 And PrevIndex <= Lv.ListItems.Count Then Lv.ListItems(PrevIndex).Selected = True
   
    On Error Resume Next
    Lv.SelectedItem.EnsureVisible
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuLadv_Click()

Dim f As New frmEmpAdv
With f
    .LongTermAdvances = True
    Dim EmpAccs() As String
    EmpAccs = Split(Lv.SelectedItem.Tag, "','")
    Call .Form_Activate
    .cmbAccNo = EmpAccs(1)
    .Show
    .TBox(1).SetFocus
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
     mnuAdd.Enabled = False
     mnuInActive.Enabled = False
     mnuLAdv.Enabled = False
     mnuSAdv.Enabled = False
   
ElseIf Lv.ListItems.Count = 0 Then
     mnuEdit.Enabled = False
     mnuLedger.Enabled = False
     mnuPrintList.Enabled = False
     mnuPrintDetail.Enabled = False
     mnuAdd.Enabled = True
     mnuInActive.Enabled = True
     mnuLAdv.Enabled = False
     mnuSAdv.Enabled = False
Else
     mnuEdit.Enabled = True
     mnuLedger.Enabled = True
     mnuPrintList.Enabled = True
     mnuPrintDetail.Enabled = True
     mnuAdd.Enabled = True
     
     mnuInActive.Caption = IIf(Lv.SelectedItem.SmallIcon = "Active", "Make This Employee Ina&ctive", "Make This Employee A&ctive")
     mnuInActive.Tag = Lv.SelectedItem.SmallIcon
     mnuInActive.Enabled = True
     
     mnuLAdv.Enabled = (mnuInActive.Tag = "Active")
     mnuSAdv.Enabled = (mnuInActive.Tag = "Active")

End If

 End Sub
