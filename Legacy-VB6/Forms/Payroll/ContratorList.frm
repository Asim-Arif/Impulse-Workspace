VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form ContractorList 
   ClientHeight    =   6570
   ClientLeft      =   1110
   ClientTop       =   60
   ClientWidth     =   10365
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6570
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
            Picture         =   "ContratorList.frx":0000
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ContratorList.frx":01BC
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
            Picture         =   "ContratorList.frx":0378
            Key             =   "Active"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ContratorList.frx":07CC
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
         Caption         =   "Contractor List"
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
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contractor List"
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
         Width           =   2415
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   5565
      Left            =   15
      TabIndex        =   3
      Top             =   495
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   9816
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
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
         Text            =   "Emp ID"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Father's Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Religion"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Sex"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Phone 1"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Phone 2"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   375
      Left            =   8685
      TabIndex        =   5
      Top             =   6135
      Width           =   1665
      ForeColor       =   64
      Caption         =   "Close          "
      PicturePosition =   327683
      Size            =   "2937;661"
      Picture         =   "ContratorList.frx":0C20
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
      Picture         =   "ContratorList.frx":0D32
      Accelerator     =   79
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "ContractorList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Add As Boolean, Edit As Boolean

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
        Call CheckMenu
        MainForm.mnupopemp.Tag = "From Cont"
        Me.PopUpMenu MainForm.mnupopemp, , .Left, .Top
    End With
    
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
    MainForm.mnuShowInactive.Checked = False
End Sub

Private Sub LV1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

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
Call mnuEdit(IIf(Left$(LV.SelectedItem.key, 3) = "DPT", "DPT", "EMP"))
    
End Sub




Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()

    Call RefreshList
    'mnuAdd.Enabled = Not OutOfFinancialYear
    'mnuEdit.Enabled = Not OutOfFinancialYear
    'mnuDelete.Enabled = Not OutOfFinancialYear
    
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






Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    Call CheckMenu
    MainForm.mnupopemp.Tag = "From Cont"
    Me.PopUpMenu MainForm.mnupopemp
End Sub








Public Sub mnuEdit(Cond As String)
    Dim f As Form
    If Cond = "DPT" Then
        Set f = New NewDepartment
        f.TableName = "ContDepartments"
        f.Tag = Right$(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 3)
        f.Add = False
        f.Editrec
        f.Show 1
        
    Else
        Set f = New NewEmployee
        'Load NewEmployee
        f.TableName = "Contractors"
        f.Tag = LV.SelectedItem.key
        f.Add = False
        f.EmpTab.TabVisible(3) = False
        f.Editrec
        
        'Load NewEmployee
        'NewEmployee.TableName = "Contractors"
        'NewEmployee.Tag = LV.SelectedItem.Key
        'NewEmployee.Add = False
        'NewEmployee.Editrec
        'NewEmployee.Show 1
        
    End If
    Call RefreshList
End Sub
Public Sub mnuAddDept()
    Dim f As New NewDepartment
    f.TableName = "ContDepartments"
    f.lblCatID = GetNextDeptID
    f.Add = True
    f.Show 1
    Call RefreshList
End Sub
Private Function GetNextDeptID() As String
    Dim NextVal As String
    NextVal = GetMax("cast(DeptID as int)", "contdepartments")
    GetNextDeptID = Format(Val(NextVal) + 1, "00")
End Function
Public Sub mnuAdd()
    Dim f As New NewEmployee
    f.TableName = "Contractors"
    f.Add = True
    f.EmpTab.TabVisible(3) = False
    
    f.Show 1
    Call RefreshList
End Sub
Public Sub RefreshIt()

    Dim Cond As String
    If MainForm.mnuShowInactive.Checked Then
        Cond = ""
    Else
        Cond = " Where Active=1"
    End If
    
    Call RefreshList(Cond)
    
End Sub
Public Sub RefreshList(Optional Cond)

Dim rs As New ADODB.Recordset
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
rsDept.Open "Select deptid,Name,Active from contdepartments " & Cond, con, adOpenForwardOnly, adLockReadOnly

LV.ListItems.Clear
Do Until rsDept.EOF
    With rs
        'Check For Show All Option
        'ActiveCondition = IIf(mnuShowAll.Checked, "Where Left(ID,3)='CUS'", " Where not InActive and Left(ID,3)='CUS'")
        If Cond = "" Then
        Else
            Cond = " and Active=1"
        End If
        .Open "Select * From VContractors where DeptID='" & rsDept.Fields(0) & "'" & Cond, con, adOpenForwardOnly, adLockReadOnly
        Set ITM = LV.ListItems.Add(, "DPT" & rsDept.Fields(0), rsDept.Fields(0))
        ITM.Bold = True
        ITM.ListSubItems.Add(, , rsDept.Fields(1) & "").Bold = True
        If Not rsDept![Active] Then
            ITM.ForeColor = vbBlack
        End If
        For i = 0 To .RecordCount - 1
            'Set Color For Inactive
            'If ![Inactive] Then
            '    TheColor = &H808080
            '    TheTag = "Inactive"
            'Else
            '    TheColor = &H800000
            '    TheTag = "Active"
            'End If
            If IsNull(![EmpId]) Then Exit For
            Set ITM = LV.ListItems.Add(, ![EmpId], "  » " & ![EmpId])
            'itm.ForeColor = TheColor
            'itm.Tag = TheTag
            If Not ![Active] Then
                ITM.ForeColor = vbBlack
            End If
            ITM.ListSubItems.Add(, , ![Contractors.Name] & "").ToolTipText = ![Contractors.Name] & ""
            'itm.ListSubItems(1).ForeColor = TheColor
            ITM.ListSubItems.Add(, , ![FName] & "").ToolTipText = ![FName] & ""
            'itm.ListSubItems(2).ForeColor = TheColor
            ITM.ListSubItems.Add , , ![Religion] & ""
            ITM.ListSubItems.Add , , ![Sex] & ""
            ITM.ListSubItems.Add , , ![Phone1] & " "
            ITM.ListSubItems.Add , , ![Phone2] & ""
            'itm.ListSubItems.Add , , ![Phone1] & " "
            'itm.ListSubItems.Add(, , Format(Val(![Balance] & ""), "0.00 DR;0.00 CR;0.00     ")).ForeColor = TheColor
            .MoveNext
        Next
        .Close
        rsDept.MoveNext
    End With
Loop
Set rs = Nothing


'If PrevIndex > 0 And PrevIndex <= LV.ListItems.Count Then LV.ListItems(PrevIndex).Selected = True
   
'On Error Resume Next
'LV.SelectedItem.EnsureVisible
'call hidewait
End Sub




 
 
 
Public Sub MakeInactive()
    Dim Act As Boolean
    Dim Sql As String
    Dim theKey As String
    theKey = LV.SelectedItem.key
    If LV.SelectedItem.ForeColor = vbBlack Then Act = True Else Act = False
    
    If Left$(theKey, 3) = "DPT" Then
    
        Dim Childs As Integer
        Dim rs As New ADODB.Recordset
        With rs
            .Open "Select Count(EmpID) from Contractors Where DeptID='" & Right(theKey, Len(theKey) - 3) & "'", con, adOpenForwardOnly, adLockReadOnly
            If IsNull(.Fields(0)) Then Childs = 0 Else Childs = .Fields(0)
            .Close
        End With
        Set rs = Nothing
        
        If MsgBox("This Will Affect All The Contractors Of This Department." & _
         vbNewLine & "There Are " & Childs & _
         " Contractors In This Department." & vbNewLine & "Do You Want To Continue?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
        con.BeginTrans
        con.Execute "Update contdepartments Set Active=" & Act & " Where DeptID='" & Right(theKey, Len(theKey) - 3) & "'"
        con.Execute "Update Contractors Set Active=" & Act & " Where DeptID='" & Right(theKey, Len(theKey) - 3) & "'"
        con.CommitTrans
    Else
        con.Execute "Update Contractors Set Active=" & Act & " Where EmpID='" & theKey & "'"
    End If
    
    
    
    Dim Cond As String
    If MainForm.mnuShowInactive.Checked Then
        Cond = ""
    Else
        Cond = " Where Active=1"
    End If
    
    Call RefreshList(Cond)
    
End Sub
Private Sub CheckMenu()
    
    With MainForm
        
        .mnuEmpNewEmp.Caption = "Add New &Contractors"
        .mnuShowInactive.Caption = "Show Inactive Contractors/Departments"
        .mnuMake.Caption = "Make This Contractor Inactive"
        .mnuEmpPrint.Caption = "&Print Contractor List"
        If LV.ListItems.Count = 0 Then
            .mnuEmpEdit.Enabled = False
            'mnuLedger.Enabled = False
            .mnuEmpPrint.Enabled = False
            .mnuMake.Enabled = False
            
            'mnuPrintDetail.Enabled = False
            '.mnuAdd.Enabled = True
            .mnuempNewDept.Enabled = True
            .mnuEmpNewEmp.Enabled = True
            'mnuInActive.Enabled = True
            'mnuPrchs.Enabled = False
            'mnuPymt.Enabled = False
            .mnuAssItems.Visible = False
        Else
            Dim DeptEmp As String
            Dim ActInAct As String
            
            If Left$(LV.SelectedItem.key, 3) = "DPT" Then DeptEmp = "Department" Else DeptEmp = "Contractor"
            
            If Left$(LV.SelectedItem.key, 3) = "DPT" Then .mnuAssItems.Visible = False Else .mnuAssItems.Visible = True
            
            If LV.SelectedItem.ForeColor = vbBlack Then ActInAct = "Active" Else ActInAct = "InActive"
            .mnuMake.Caption = "&Make This " & DeptEmp & " " & ActInAct
            
            .mnuEmpEdit.Caption = IIf(Left$(LV.SelectedItem.key, 3) = "DPT", "&Edit This Department", "&Edit This Contractor")
            .mnuEmpEdit.Tag = IIf(Left$(LV.SelectedItem.key, 3) = "DPT", "DPT", "EMP")
            .mnuEmpEdit.Enabled = True
            .mnuEmpPrint.Enabled = True
            .mnuMake.Enabled = True
        End If
    End With

 End Sub
Public Sub showItems()

    Load AssItems
    Call AssItems.showAssigned(LV.SelectedItem.Text, LV.SelectedItem.SubItems(1))
    
End Sub
