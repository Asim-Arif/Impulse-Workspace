VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form DeptList 
   ClientHeight    =   6795
   ClientLeft      =   1110
   ClientTop       =   240
   ClientWidth     =   10425
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   10425
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   540
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
            Picture         =   "DeptList.frx":0000
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DeptList.frx":01BC
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
            Picture         =   "DeptList.frx":0378
            Key             =   "Active"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DeptList.frx":07CC
            Key             =   "Inactive"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV 
      Height          =   5940
      Left            =   15
      TabIndex        =   0
      Top             =   105
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   10478
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Dept ID"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   6350
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Salary Acc"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Accrued Acc"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   375
      Left            =   8685
      TabIndex        =   2
      Top             =   6135
      Width           =   1665
      ForeColor       =   64
      Caption         =   "Close          "
      PicturePosition =   327683
      Size            =   "2937;661"
      Picture         =   "DeptList.frx":0C20
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
      TabIndex        =   1
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
         Caption         =   "Refresh"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnushowDept 
         Caption         =   "Show Inactive Departements"
      End
      Begin VB.Menu mnuInActive 
         Caption         =   "Make This Departement Inactive"
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewDept 
         Caption         =   "Add New Departement"
      End
      Begin VB.Menu mnuEditDept 
         Caption         =   "Edit Departement Info."
      End
      Begin VB.Menu dash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintdeptList 
         Caption         =   "Print Departement List"
      End
   End
End
Attribute VB_Name = "DeptList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public add As Boolean, edit As Boolean

Private Sub cmdClose_Click()
    Unload Me
End Sub
Public Sub ShowInactive(m As Menu)

    Dim Cond As String
    If m.Checked Then
        Cond = " WHERE Active=1"
    Else
        Cond = ""
    End If
    m.Checked = Not m.Checked
    Call RefreshList(Cond)
    
End Sub

Private Sub cmdOpt_Click()

    Call CheckMenu
    
    Me.PopUpMenu mnuPop
    
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
    'MainForm.mnuShowDept.Checked = False
End Sub


Private Sub LV_DblClick()
    If LV.ListItems.count = 0 Then Exit Sub
    Call EditDept
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()

    Call RefreshList
    mnuPop.Visible = False
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    ' fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
 
    Dim TitleBottom As Long

    TitleBottom = 20 'fTitle.Top + fTitle.Height
 
    LV.Move 50, TitleBottom, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + cmdClose.Height + 100)
 
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 50
    cmdOpt.Move cmdClose.Left - cmdOpt.Width - 75, cmdClose.Top
 
 
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    
    Call CheckMenu
    
    Me.PopUpMenu mnuPop
End Sub


Public Sub EditDept()
    

    Dim f As New NewDepartment
    f.TableName = "Departments"
    f.Tag = Right$(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 3)
    f.add = False
    f.edit = True
    f.Editrec
    f.Show 1
    
    Call RefreshList
End Sub
Public Sub mnuAddDept()
    Dim f As New NewDepartment
    f.TableName = "Departments"
    f.lblCatID = GetNextDeptID(f.TableName)
    f.add = True
    f.Show 1
    Call RefreshList
End Sub

Public Sub AddNewDept()
    Dim f As New NewDepartment
    f.add = True
    f.edit = False
    f.Show 1
    Call RefreshList
End Sub
Public Sub RefreshIt()

    'Dim Cond As String
    'If MainForm.mnuShowDept.Checked Then
    '    Cond = ""
    'Else
    '    Cond = " Where Active=1"
    'End If
   '
   ' Call RefreshList(Cond)
    
End Sub
Public Sub RefreshList(Optional Cond)

'Dim rs As New ADODB.Recordset
Dim rsDept As New ADODB.Recordset
Dim ITM As ListItem
Dim PrevIndex As Double
Dim TheColor As Long
Dim ActiveCondition As String


If IsMissing(Cond) Then
    Cond = " Where Active=1"
End If
    
'If LV.ListItems.Count > 0 Then PrevIndex = LV.SelectedItem.Index
'Set LV.SmallIcons = ImageList1

If mnushowDept.Checked Then
    Cond = ""
Else
    Cond = " Where Active=1"
End If

rsDept.Open "Select deptid,Name,Active,Description,Color,SalaryAccNo,AccruedAccNo from Departments " & Cond, con, adOpenForwardOnly, adLockReadOnly

LV.ListItems.Clear
Do Until rsDept.EOF
    With rsDept
        'Check For Show All Option
        'ActiveCondition = IIf(mnuShowAll.Checked, "Where Left(ID,3)='CUS'", " Where not InActive and Left(ID,3)='CUS'")
        
        TheColor = Val(![Color] & "")

        Set ITM = LV.ListItems.add(, "DPT" & rsDept.Fields(0), rsDept.Fields(0))
        ITM.ForeColor = TheColor
        'itm.Bold = True
        ITM.ListSubItems.add(, , rsDept.Fields(1) & "").ForeColor = TheColor
        ITM.ListSubItems.add(, , rsDept.Fields(3) & "").ForeColor = TheColor
        ITM.ListSubItems(2).Tag = Abs(![Active]) & ""
        ITM.ListSubItems.add(, , !SalaryAccNo & "").ForeColor = TheColor
        ITM.ListSubItems.add(, , !AccruedAccNo & "").ForeColor = TheColor
        If Not rsDept![Active] Then
            ITM.ForeColor = vbBlack
        End If
        
        .MoveNEXT
    End With
Loop
rsDept.Close
Set rsDept = Nothing


'If PrevIndex > 0 And PrevIndex <= LV.ListItems.Count Then LV.ListItems(PrevIndex).Selected = True
   
'On Error Resume Next
'LV.SelectedItem.EnsureVisible
'Me.MousePointer = vbNormal
End Sub




 
 
 
Public Sub MakeInactive()
    Dim Act As Boolean
    Dim SQL As String
    Dim theKey As String
    theKey = LV.SelectedItem.key
    If LV.SelectedItem.ForeColor = vbBlack Then Act = True Else Act = False
        Dim DisContOn As String
    If Act Then DisContOn = "NULL" Else DisContOn = "'" & Format(Date, "dd-MMM-yyyy") & "'"

    If Left$(theKey, 3) = "DPT" Then
    
        Dim Childs As Integer
        Dim rs As New ADODB.Recordset
        With rs
            .Open "Select Count(EmpID) from Employees Where DeptID='" & Right(theKey, Len(theKey) - 3) & "'", con, adOpenForwardOnly, adLockReadOnly
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
        con.Execute "Update Departments Set Active=" & Abs(Act) & " Where DeptID='" & Right(theKey, Len(theKey) - 3) & "'"
        con.Execute "Update Employees Set Active=" & Abs(Act) & ",DiscontinuedOn=" & DisContOn & " Where DeptID='" & Right(theKey, Len(theKey) - 3) & "'"
        con.CommitTrans
    Else
        con.Execute "Update Employees Set Active=" & Abs(Act) & " Where EmpID='" & theKey & "'"
    End If
    
    Call CheckMenu
    RefreshList
    
End Sub
Private Sub CheckMenu()
    
    With Me
        If LV.ListItems.count = 0 Then
        
            .mnuInActive.Enabled = False
            .mnuEditDept.Enabled = False
            .mnuPrintdeptList.Enabled = False
            .mnuNewDept.Enabled = True
        Else
            Dim DeptEmp As String
            Dim ActInAct As String

            DeptEmp = "Department"
            If LV.SelectedItem.ListSubItems(2).Tag = "0" Then ActInAct = "Active" Else ActInAct = "InActive"
            .mnuInActive.Caption = "&Make This " & DeptEmp & " " & ActInAct

            .mnuEditDept.Caption = IIf(Left$(LV.SelectedItem.key, 3) = "DPT", "&Edit This Department", "&Edit This Employee")
            .mnuEditDept.Tag = IIf(Left$(LV.SelectedItem.key, 3) = "DPT", "DPT", "EMP")
            .mnuEditDept.Enabled = True
            .mnuPrintdeptList.Enabled = True
            .mnushowDept.Enabled = True

            .mnuInActive.Enabled = UserHasAccess("EmpActive")
            .mnuNewDept.Enabled = UserHasAccess("NewDept")

        End If
    End With ''

 End Sub
 
 
Public Sub PrintDeptList()

    Screen.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptDeptList.rpt")
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
    
    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & strCompany & "'"
        End If
    Next
    
    
    f.ShowReport " ", rpt
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub mnuEditDept_Click()
    Call EditDept
End Sub

Private Sub mnuInActive_Click()
    Call MakeInactive
End Sub

Private Sub mnuNewDept_Click()
    Call AddNewDept
End Sub

Private Sub mnuPrintdeptList_Click()
    Call PrintDeptList
End Sub

Private Sub mnuRefresh_Click()
    Call RefreshList
End Sub

Private Sub mnuShowDept_Click()
    'mnushowDept.Checked = Not mnushowDept.Checked
    Call ShowInactive(mnushowDept)
End Sub
