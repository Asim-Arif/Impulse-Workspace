VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDCEmpList 
   ClientHeight    =   6795
   ClientLeft      =   1170
   ClientTop       =   15
   ClientWidth     =   11895
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   11895
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   795
      Left            =   30
      TabIndex        =   3
      Top             =   5880
      Width           =   7155
      Begin VB.TextBox txtValue 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2910
         TabIndex        =   5
         Top             =   390
         Width           =   2745
      End
      Begin VB.ComboBox cmbField 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmDCEmpList.frx":0000
         Left            =   60
         List            =   "frmDCEmpList.frx":001F
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   390
         Width           =   2835
      End
      Begin MSForms.CommandButton cmdSearch 
         Height          =   360
         Left            =   5700
         TabIndex        =   7
         Top             =   360
         Width           =   1380
         Caption         =   "Search      "
         PicturePosition =   327683
         Size            =   "2434;635"
         Accelerator     =   83
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Search By :                                           Search For:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   6
         Top             =   150
         Width           =   5235
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   435
      Top             =   930
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
            Picture         =   "frmDCEmpList.frx":0078
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDCEmpList.frx":0234
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2580
      Top             =   1590
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
            Picture         =   "frmDCEmpList.frx":03F0
            Key             =   "Active"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDCEmpList.frx":0844
            Key             =   "Inactive"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV 
      Height          =   5745
      Left            =   15
      TabIndex        =   0
      Top             =   60
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   10134
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   16777215
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Emp ID"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   4145
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Father's Name"
         Object.Width           =   3464
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Designation"
         Object.Width           =   2671
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Religion"
         Object.Width           =   1639
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Sex"
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Phone 1"
         Object.Width           =   1826
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Phone 2"
         Object.Width           =   1905
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Salary"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   1215
      Left            =   60
      TabIndex        =   8
      Top             =   1440
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   2143
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
      NumItems        =   28
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "Emp ID"
         Text            =   "Emp ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "Emp Name"
         Text            =   "Emp Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Tag             =   "B. Sal"
         Text            =   "B. Sal"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "Rate/Hr"
         Text            =   "Rate/Day"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "A. Hrs"
         Text            =   "A.Days"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "A. Amt"
         Text            =   "A. Amt"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "A.Allow"
         Text            =   "A.Allow"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "Allow.Amt"
         Text            =   "Allow.Amt"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Object.Tag             =   "S. Hrs"
         Text            =   "S.Days"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Object.Tag             =   "S. Amt"
         Text            =   "S. Amt"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Tag             =   "Leave Hrs"
         Text            =   "Leaves"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Object.Tag             =   "Lv. Amt."
         Text            =   "Lv. Amt."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Object.Tag             =   "OT Hrs"
         Text            =   "OT Hrs"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Object.Tag             =   "OT Amt"
         Text            =   "OT Amt"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Object.Tag             =   "L.Hrs"
         Text            =   "L.Hrs"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Object.Tag             =   "L.Amt"
         Text            =   "L.Amt"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Object.Tag             =   "Total"
         Text            =   "Total"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   17
         Object.Tag             =   "T. Amt"
         Text            =   "T. Amt"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   18
         Object.Tag             =   "Net Ttl"
         Text            =   "Net Ttl"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   19
         Object.Tag             =   "ST Loan"
         Text            =   "ST Loan"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   20
         Object.Tag             =   "Loan Ded."
         Text            =   "Loan Ded."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   21
         Object.Tag             =   "Adv Sal"
         Text            =   "Adv Sal"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   22
         Object.Tag             =   "U. Fund"
         Text            =   "U. Fund"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   23
         Object.Tag             =   "Fine"
         Text            =   "Fine"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   24
         Object.Tag             =   "Scholorship"
         Text            =   "Scholorship"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   25
         Object.Tag             =   "Lunch Ded."
         Text            =   "Lunch Ded."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   26
         Object.Tag             =   "EOBI"
         Text            =   "EOBI"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   27
         Object.Tag             =   "Balance"
         Text            =   "Balance"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   375
      Left            =   10170
      TabIndex        =   2
      Top             =   5895
      Width           =   1665
      ForeColor       =   64
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
      Left            =   8490
      TabIndex        =   1
      Top             =   5895
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
   Begin VB.Menu mnuPop 
      Caption         =   "PopMenuEmp"
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu d 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print Settlement Form"
      End
   End
End
Attribute VB_Name = "frmDCEmpList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PrevIndex As Long

Public add As Boolean, edit As Boolean

Private Sub cmbField_Change()
    txtValue = ""
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub ShowInactive(m As Menu)

    Dim Cond As String
    If m.Checked Then
        Cond = " Where Active=1"
    Else
        Cond = ""
    End If
    
    m.Checked = Not m.Checked
    Call RefreshList(Cond)
    
End Sub

Private Sub cmdOpt_Click()

    With cmdOpt
        Me.PopUpMenu mnuPop, , .Left, .Top
    End With
    
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
    'MainForm.mnuShowInactive.Checked = False
End Sub

Private Sub LV_ColumnClick1(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

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
    If LV.ListItems.count = 0 Then Exit Sub
     If UserHasAccess("EditEmployee") Then Call EditItem(IIf(Left$(LV.SelectedItem.key, 3) = "DPT", "DPT", "EMP"))
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    mnuPop.Visible = False
    cmbField.ListIndex = 0
    Call RefreshList
    
End Sub

Private Sub Form_Resize()
 On Error Resume Next
 'fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
 
 Dim TitleBottom As Long

 TitleBottom = 20 'fTitle.Top + fTitle.Height
 
 LV.Move 50, TitleBottom, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + cmdClose.Height + 450)
 
 cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 50
 cmdOpt.Move cmdClose.Left - cmdOpt.Width - 75, cmdClose.Top
 Frame1.Top = Me.ScaleHeight - Frame1.Height - 50
 
End Sub




Private Sub cmdSearch_Click()


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
       
       
       If cmbField.ListIndex = 0 Then
          fldValue = Trim(Replace(LV.ListItems(i), "»", ""))
       Else
          fldValue = Trim(Replace(LV.ListItems(i).SubItems(cmbField.ListIndex), "»", ""))
       End If
    
    
       If LCase(fldValue) Like "*" & LCase(Trim(txtValue)) & "*" Then
          LV.ListItems(i).Selected = True
          LV.ListItems(i).EnsureVisible
          If i <> StartIdx Then Exit For
       End If
       
    Next

    If i = LV.ListItems.count + 1 Then
        MsgBox "Item Not Found.", vbInformation
        Set LV.SelectedItem = Nothing
    End If
    
Exit Sub



On Error Resume Next

If cmbField.ListIndex = 0 Then
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).key Like "*" & txtValue & "*" Then
            LV.ListItems(i).Selected = True
            Exit For
        End If
    Next
Else
    For i = 1 To LV.ListItems.count
        If LCase(LV.ListItems(i).SubItems(cmbField.ListIndex) & "") Like "*" & LCase(txtValue) & "*" Then
            LV.ListItems(i).Selected = True
            Exit For
        End If
    Next
End If
        

End Sub



Public Sub EditItem(Cond As String)
    Dim f As Form
    If Cond = "DPT" Then
        Set f = New NewDepartment
        f.TableName = "Departments"
        f.Tag = Right$(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 3)
        f.add = False
        f.Editrec
        f.Show 1
        
    Else
        Set f = New NewEmployee
        'Load NewEmployee
        f.TableName = "Employees"
        f.Tag = LV.SelectedItem.key
        f.add = False
        f.Editrec
        'Me.Show
        'f.EmpTab.Tab = 0
        'Me.Hide
        'Me.Show 1
        'f.Show 1
        'f.TBox(0).SetFocus
    End If
    Call RefreshList
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   
   If Button <> 2 Then Exit Sub
   
    
    PopUpMenu mnuPop
    
End Sub

Private Sub mnuappForm_Click()
    Call PrintAppForm
End Sub

Private Sub mnuedit_Click()
    Call EditItem(mnuEdit.Tag)
End Sub

Private Sub mnuEmpLeaveBalance_Click()

    Screen.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpLeaveBalance.rpt")
    SelFormula = "{VEmpLeaveBalance.DeptID}<>'VDRS' and {VEmpLeaveBalance.EmpID}<>''"
    SelFormula = SelFormula & IIf(mnuShowInactive.Checked, "", " and {VEmpLeaveBalance.Active}=True ")
    
    f.ShowReport SelFormula & " ", rpt
    
    Screen.MousePointer = vbDefault


End Sub

Public Sub mnuNewDept_Click()
    Dim f As New NewDepartment
    f.TableName = "Departments"
    f.lblCatID = GetNextDeptID
    f.add = True
    f.Show 1
    Call RefreshList
End Sub
Private Function GetNextDeptID() As String
    Dim NextVal As String
    NextVal = GetMax("cast(DeptID as int)", "Departments")
    GetNextDeptID = Format(Val(NextVal) + 1, "00")
End Function
Public Sub mnunewEmp_Click()
    Dim f As New NewEmployee
    f.TableName = "Employees"
    f.add = True
    
    If LV.ListItems.count > 0 Then
        f.cmbDept.Text = LV.SelectedItem.Tag
    End If
    
    f.Show 1
    Call RefreshList
End Sub
Public Sub RefreshIt()

    'Dim Cond As String
    'If MainForm.mnuShowInactive.Checked Then
    '    Cond = ""
    'Else
    '    Cond = " Where Active=1"
    'End If
   '
   ' Call RefreshList(Cond)
    
End Sub
Public Sub RefreshList(Optional Cond)

Dim rs As New ADODB.Recordset
Dim rsDept As New ADODB.Recordset
Dim ITM As ListItem
Dim PrevIndex As Double
Dim TheColor As Long
Dim ActiveCondition As String

If IsMissing(Cond) Then
    Cond = " Where Active=0"
End If
    
If LV.ListItems.count > 0 Then PrevIndex = LV.SelectedItem.Index
rsDept.Open "Select Deptid,Name,Active,Color from Departments ", con, adOpenForwardOnly, adLockReadOnly


LV.ListItems.Clear
Dim lDeptColor As Long
Do Until rsDept.EOF
    With rs
        TheColor = Val(rsDept![Color] & "")
        lDeptColor = TheColor
        .Open "Select EmpID,Name,Active,FName,Designation,Religion,Sex,Phone1,Phone2,StartingSalary From VEmp where DeptID='" & rsDept.Fields(0) & "' AND Active=0", con, adOpenForwardOnly, adLockReadOnly
        Set ITM = LV.ListItems.add(, "DPT" & rsDept.Fields(0), rsDept.Fields(0))
        ITM.Bold = True
        ITM.Tag = rsDept!Name
        
        ITM.ForeColor = TheColor
        ITM.ListSubItems.add(, , rsDept.Fields(1) & "").Bold = True
        ITM.ListSubItems(1).ForeColor = TheColor
        ITM.ListSubItems(1).Tag = rsDept![Active] & ""
        
        If Not rsDept![Active] Then
            ITM.ForeColor = vbBlack
        End If
        For i = 0 To .RecordCount - 1
            If IsNull(![EmpId]) Then Exit For
            Set ITM = LV.ListItems.add(, ![EmpId], "  » " & ![EmpId])
            ITM.Tag = rsDept!Name
            TheColor = lDeptColor
            ITM.ForeColor = TheColor
            
            'itm.ForeColor = TheColor
            'itm.Tag = TheTag
            If Not ![Active] Then
                ITM.ForeColor = vbRed
                TheColor = vbRed
            End If
            ITM.ListSubItems.add(, , ![Name] & "").ToolTipText = ![Name] & ""
            ITM.ListSubItems(1).ForeColor = TheColor
            ITM.ListSubItems(1).Tag = ![Active] & ""
            
            'itm.ListSubItems(1).ForeColor = TheColor
            ITM.ListSubItems.add(, , ![FName] & "").ToolTipText = ![FName] & ""
            ITM.ListSubItems.add(, , ![Designation] & "").ForeColor = TheColor
            ITM.ListSubItems(2).ForeColor = TheColor
            
            'itm.ListSubItems(2).ForeColor = TheColor
            ITM.ListSubItems.add(, , ![Religion] & "").ForeColor = TheColor
            ITM.ListSubItems.add(, , ![Sex] & "").ForeColor = TheColor
            ITM.ListSubItems.add(, , ![Phone1] & " ").ForeColor = TheColor
            ITM.ListSubItems.add(, , ![Phone2] & "").ForeColor = TheColor
            ITM.ListSubItems.add(, , ![StartingSalary] & "").ForeColor = TheColor
            'itm.ListSubItems.Add , , ![Phone1] & " "
            'itm.ListSubItems.Add(, , Format(Val(![Balance] & ""), "0.00 DR;0.00 CR;0.00     ")).ForeColor = TheColor
            .MoveNEXT
        Next
        .Close
        rsDept.MoveNEXT
    End With
Loop
Set rs = Nothing


If PrevIndex > 0 And PrevIndex <= LV.ListItems.count Then LV.ListItems(PrevIndex).Selected = True
   
On Error Resume Next
LV.SelectedItem.EnsureVisible
Me.MousePointer = vbNormal
End Sub
 
Public Sub MakeInactive(Optional NoMessage As Boolean = False)

    Dim Act As Boolean
    Dim SQL As String
    Dim theKey As String
    Dim DCDate As Date
    theKey = LV.SelectedItem.key
    
    If LV.SelectedItem.ForeColor = vbBlack Then Act = True Else Act = False
    
    Dim DisContOn As String
    If Act Then DisContOn = "NULL" Else DisContOn = "'" & Format(Date, "dd-MMM-yyyy") & "'"
        
        
    If Left$(theKey, 3) = "DPT" Then
    
        Dim Childs As Integer
        Dim rs As New ADODB.Recordset
        With rs
            .Open "Select Count(EmpID) From Employees Where DeptID='" & Right(theKey, Len(theKey) - 3) & "'", con, adOpenForwardOnly, adLockReadOnly
            If IsNull(.Fields(0)) Then Childs = 0 Else Childs = .Fields(0)
            .Close
        End With
        Set rs = Nothing
        
        If MsgBox("This Will Affect All The Employees Of This Department." & _
         vbNewLine & "There Are " & Childs & _
         " Employees In This Department." & vbNewLine & "Do You Want To Continue?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
        con.BeginTrans
        
        
        con.Execute "Update Departments set Active=" & Abs(Act) & " Where DeptID='" & Right(theKey, Len(theKey) - 3) & "'"
        con.Execute "Update Employees Set Discontinuedon=" & DisContOn & ",Active=" & Abs(Act) & " Where DeptID='" & Right(theKey, Len(theKey) - 3) & "'"
        con.CommitTrans
    Else
        If Not Act Then
            If MsgBox("This Is Usually Done When An Employee Discontinues. Do You Want To Continue ?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
            
            DCDate = InputBox("Plz Enter Date[mm/dd/yy] To Discontinue.")
            If Not IsDate(DCDate) Then
                MsgBox "Invalid Date.", vbCritical
                Exit Sub
            End If
            con.Execute "Update Employees Set DiscontinuedOn='" & DCDate & "',Active=" & Abs(Act) & " Where EmpID='" & theKey & "'"
        Else

            DCDate = InputBox("Plz Enter Date[mm/dd/yy] To ReJoin.", "Re Joining After Discontinuation.")
            If Not IsDate(DCDate) Then
                MsgBox "Invalid Date.", vbCritical
                Exit Sub
            End If
            con.Execute "Update Employees Set JoinDate='" & DCDate & "',DiscontinuedOn=Null,Active=" & Abs(Act) & " Where EmpID='" & theKey & "'"
            
        End If
        
        
'       con.Execute "INSERT INTO Employees(EmpID,DisjoinDT) VALUES()"
'        If Not Act Then
'            With Payslip
'                .cmbEmp.Text = LV.SelectedItem.SubItems(1)
'                .cmbDept.Enabled = False
'                .cmbEmp.Enabled = False
'                .Show
'            End With
'        End If
        
        
    End If
    
    Call RefreshList
    Dim Cond As String

End Sub



Public Sub PrintAppForm()
'With cr1
'    .ReportFileName = RptPath & "\ApplicationForm.rpt"
'    .Connect = ""
'    .SelectionFormula = ""
'    .Formulas(0) = ""
'    .Formulas(1) = "Company='" & company & "'"
'    .Action = 1
'End With

Screen.MousePointer = vbHourglass
Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptAppForm.rpt")
rpt.FormulaFields(1).Text = "'" & company & "'"
f.ShowReport " ", rpt
Screen.MousePointer = vbDefault

End Sub

Public Sub PrintEmpList(Optional WithExternalSalary As Boolean = False, Optional Transport As Boolean = False)

'    With cr1
'        .ReportFileName = RptPath & "\EmpList.rpt"
'        .Connect = ConnectStr
'
'        Dim SelFormula As String
        If mnuShowInactive.Checked Then
            SelFormula = IIf(Transport, " {VEmp.Transport}=true", "")
        Else
            SelFormula = "{VEmp.Active}=true " & IIf(Transport, " and {VEmp.Transport}=true", "")
        End If
    
'        .SelectionFormula = SelFormula
'
'        .Formulas(0) = "External=" & WithExternalSalary
'
'        .Action = 1
    
'    End With

Screen.MousePointer = vbHourglass

Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpList.rpt")
Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

Set FormulaFields = rpt.FormulaFields
For Each FormulaField In FormulaFields
    If FormulaField.Name = "{@Company}" Then
        FormulaField.Text = "'" & company & "'"
    End If
Next


f.ShowReport SelFormula & " ", rpt

Screen.MousePointer = vbDefault


End Sub

Public Sub PrintEmpListWithoutSal(Optional WithDOB As Boolean = False, Optional SelFormula As String, Optional ReportCaption As String = "Employee List")

'    With cr1
'        .ReportFileName = RptPath & "\EmpListWithoutSal.rpt"
'        .Connect = ConnectStr
'
        Dim TransCondition As String

        If mnuShowInactive.Checked Then
            TransCondition = IIf(Transport, " {VEmp.Transport}=true", "")
        Else
            TransCondition = "{VEmp.Active}=true " & IIf(Transport, " and {VEmp.Transport}=true", "")
        End If
        
        If TransCondition <> "" Then
            If SelFormula <> "" Then
                SelFormula = TransCondition & " and " & SelFormula
            Else
                SelFormula = TransCondition
            End If
        End If
        
        
'
'        .SelectionFormula = SelFormula
'
'        .Formulas(0) = "Company='" & company & "'"
'
'        .Action = 1
'
'    End With

Screen.MousePointer = vbHourglass

Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpListWithoutSal.rpt")
Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

Set FormulaFields = rpt.FormulaFields
For Each FormulaField In FormulaFields
    If FormulaField.Name = "{@Company}" Then
        FormulaField.Text = "'" & company & "'"
    ElseIf FormulaField.Name = "{@BirthDate}" Then
        FormulaField.Text = WithDOB
    ElseIf FormulaField.Name = "{@RptCaption}" Then
        FormulaField.Text = "'" & ReportCaption & "'"
    End If
Next


f.ShowReport SelFormula & " ", rpt

Screen.MousePointer = vbDefault

End Sub


Public Sub PrintEmpCard(Optional EmpId As String)

Dim Cond As String
'    With cr1
'        .ReportFileName = RptPath & "\EmpCardsNew.rpt"
'        .Connect = ConnectStr
'        If EmpId = "" Then
'            .SelectionFormula = "{Employees.Active}=True"
'        Else
'            .SelectionFormula = "{Employees.EmpID}='" & Trim(Replace(EmpId, "»", "")) & "' and {Employees.Active}=True"
'        End If
'        '.Formulas(0) = "Valid Till=#" & Format(DateAdd("yyyy", 1, GetServerDate), "dd-MMM-yyyy") & "#"
'        .Formulas(1) = ""
'        .Action = 1
'    End With

If EmpId = "" Then
    Cond = "{Employees.Active}=True"
Else
    Cond = "{Employees.EmpID}='" & Trim(Replace(EmpId, "»", "")) & "' and {Employees.Active}=True"
End If


Screen.MousePointer = vbHourglass

Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpCards.rpt")
f.ShowReport Cond & " ", rpt

Screen.MousePointer = vbDefault
    
    
End Sub


Private Sub mnuempRefresh_Click()
    Call RefreshList
End Sub

Private Sub mnuMake_Click()
    If getPassword Then
        Call MakeInactive
    End If
End Sub

Private Sub mnuPrint_Click()
    
    On Error GoTo err
    
    If Left(LV.SelectedItem.key, 3) = "DPT" Then Exit Sub
    
'    Dim rs As New ADODB.Recordset
'
'    With rs
'        .Open "Select EmpID From EmpSettlements Where EmpID='" & LV.SelectedItem.key & "'", con, adOpenForwardOnly, adLockReadOnly
'        If .EOF Then
'            Call PrintSettlementOld
'        Else
'            Call PrintSettlementNew
'        End If
'        .Close
'    End With
'    Set rs = Nothing
    Call PrintSettlementNew
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub PrintSettlementNew()
        
    
    Call FinalizedSalarySheetLoaded
    Dim rpt As CRAXDDRT.Report
    Call InsertRecords
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptFinalSett.rpt")
    SelFor = "{Employees.EmpID}='" & LV.SelectedItem.key & "'"
    
    Dim frm As New frmPrevRpt
    frm.ShowReport SelFor & " ", rpt
    
End Sub
Private Sub InsertRecords()

    Dim DT As Date
    DT = Date 'DateSerial(DTPicker1.year, DTPicker1.Month, 1)

    Dim i As Integer
    Dim j As Integer

    con.Execute "Delete from PrintSalary"

    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select * from PrintSalary", con, adOpenStatic, adLockOptimistic
        For i = 1 To LV1.ListItems.count
            .AddNew
            ![EmpId] = LV1.ListItems(i).Text & ""
        
            For j = 1 To 25
                If j = 25 Then
                    If .Fields(j).Type = adDouble Then
                        .Fields(j) = Val(LV1.ListItems(i).SubItems(j + 2) & "")
                    Else
                        .Fields(j) = LV1.ListItems(i).SubItems(j + 2) & ""
                    End If
                ElseIf j = 8 And chkExternal Then                           '''' For Saving in Decimal Format instead of Time
                    .Fields(j) = LV1.ListItems(i).ListSubItems(4).Tag
                Else
                    If .Fields(j).Type = adDouble Then
                        .Fields(j) = Val(LV1.ListItems(i).SubItems(j) & "")
                    Else
                        .Fields(j) = LV1.ListItems(i).SubItems(j) & ""
                    End If
                End If
            Next j
        
            ![DT] = DT & ""
            ![PrevLTLoan] = Val(LV1.ListItems(i).ListSubItems(1).Tag)
                
            Dim ArrLeaves() As String
        
            ArrLeaves = Split(LV1.ListItems(i).ListSubItems(2).Tag, ",")
        
            ![CasualLeaves] = Val(ArrLeaves(0))
            ![SickLeaves] = Val(ArrLeaves(1))
            ![AnnualLeaves] = Val(ArrLeaves(2))
            ![CompensatoryLeaves] = Val(ArrLeaves(3))
            ![WPLeaves] = Val(ArrLeaves(4))
            ![MaternityLeaves] = Val(ArrLeaves(5))
            ![HrsPerDay] = Val(LV1.ListItems(i).ListSubItems(3).Tag & "")
        
            .Update
        Next i
    
        .Close
    
    End With

    Set rs = Nothing

    SQL = "Insert into EmpLeaveBalance SELECT  EmpID,Count(CL) as CasualLeaves,Count(SL) as SickLeaves," & _
     "count(AL) as AnnualLeaves" & _
     ",count(ML) as MaternityLeaves From VEmpLeaves Group By empID"
    
    con.Execute "Delete From EmpLeaveBalance"
    con.Execute SQL, a

End Sub

Private Sub PrintSettlementOld()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptFinalSett1.rpt")
    Dim frm As New frmPrevRpt
    frm.ShowReport "{Employees.EmpID}='" & LV.SelectedItem.key & "' ", rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuRefresh_Click()

    RefreshList

End Sub

Private Sub txtValue_GotFocus()

    txtValue.SelStart = 0
    txtValue.SelLength = Len(txtValue)

End Sub

Private Function FinalizedSalarySheetLoaded() As Boolean

    Dim rs As New ADODB.Recordset
    Dim Cond As String
    Dim IsExernalSalary As Byte
    Dim GrandTotal As Double
    
    IsExernalSalary = True
    
    Cond = "Where EmpID='" & LV.SelectedItem.key & "'"
    
    If Cond = "" Then
        Cond = "Where"
    Else
        Cond = Cond & " And "
    End If
    
    'Cond = Cond & " Month(DT)=" & DTPicker1.Month & _
    " and Year(DT)=" & DTPicker1.year
    
    With rs
        .Open "Select * From VEmpSettlements Where EmpID='" & LV.SelectedItem.key & "'", con, adOpenForwardOnly, adLockReadOnly
        
        If .RecordCount = 0 Then
            FinalizedSalarySheetLoaded = False
            GoTo END_OF_FUNCTION
        End If
            
            LV1.ListItems.Clear
            Do Until .EOF
        
                Set ITM = LV1.ListItems.add(, ![EmpId], ![EmpId] & "")
                
'                If IsComputed Then
'                    ITM.Checked = IsPayed
'                    ITM.Bold = IsTransfered
'                End If
                
'                If Val(txtTotalHrs) <> Val(![TotalMonthHrs] & "") Then
'                    txtTotalHrs = ![TotalMonthHrs] & ""
'                    MsgBox "Salary Sheet Is Finalized According To " & txtTotalHrs & " Working Hours.", vbInformation
'                End If
                
            
                ITM.Tag = "" 'Val(![Emptype] & "")
                
                ITM.ListSubItems.add(, , ![Name] & "").Tag = Val(![PrevLTLoan] & "")
                ITM.ListSubItems.add , , ![BSal] & ""
                
                ITM.ListSubItems(2).Tag = ![CasualLeaves] & "," & ![SickLeaves] & "," & ![CompensatoryLeaves] & _
                "," & ![AnnualLeaves] & "," & ![WPLeaves] & "," & ![MaternityLeaves]
                
                ITM.ListSubItems.add(, , ![Rate] & "").Tag = ![HrsPerDay]
                ITM.ListSubItems.add(, , ![ADays] & "").Tag = Val(![SDays] & "")
                ITM.ListSubItems.add , , ![AAmt] & ""
                ITM.ListSubItems.add , , ![AAllow] & ""
                ITM.ListSubItems.add , , ![AAllowAmt] & ""
                ITM.ListSubItems.add , , IIf(chkExternal = vbChecked, getHrsMin(Val(!SDays & "")), ![SDays] & "")
                ITM.ListSubItems.add , , ![SAmt] & ""
                ITM.ListSubItems.add , , ![Leaves] & ""
                ITM.ListSubItems.add , , ![LeaveAmt] & ""
                ITM.ListSubItems.add , , ![OHrs] & ""
                ITM.ListSubItems.add , , ![OAmt] & ""
                ITM.ListSubItems.add , , ![lHrs] & ""
                ITM.ListSubItems.add , , ![lAmt] & ""
                ITM.ListSubItems.add , , ![total] & ""
                ITM.ListSubItems.add , , ![Tax] & ""
                ITM.ListSubItems.add , , ![NetTtl] & ""
                ITM.ListSubItems.add , , ![ShortTerm] & ""
                ITM.ListSubItems.add , , ![LongTerm] & ""
                ITM.ListSubItems.add , , ![AdvSal] & ""
                ITM.ListSubItems.add , , ![UnionFund] & ""
                ITM.ListSubItems.add , , ![Fine] & ""
                ITM.ListSubItems.add , , ![Bonus] & ""
                ITM.ListSubItems.add , , ![Lunch] & ""
                ITM.ListSubItems.add , , ![EOBI] & ""
                ITM.ListSubItems.add , , ![Balance] & ""
                
                'GrandTotal = GrandTotal + Val(ITM.SubItems(27))
                
                .MoveNEXT
            Loop
        
        
    End With
    
    'lblTotal = Format(GrandTotal, "#,##0.00")
    
    FinalizedSalarySheetLoaded = True
    'cmdUpdate.Visible = False
END_OF_FUNCTION:

    If rs.State = 1 Then rs.Close
    Set rs = Nothing
    
End Function


