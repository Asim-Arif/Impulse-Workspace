VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form AssItems 
   ClientHeight    =   6795
   ClientLeft      =   1110
   ClientTop       =   60
   ClientWidth     =   10365
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   10365
   WindowState     =   2  'Maximized
   Begin ComboList.Usercontrol1 cmbItem 
      Height          =   330
      Left            =   3480
      TabIndex        =   8
      Top             =   810
      Width           =   2220
      _ExtentX        =   3916
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin ComboList.Usercontrol1 cmbCont 
      Height          =   330
      Left            =   1290
      TabIndex        =   6
      Top             =   810
      Width           =   2220
      _ExtentX        =   3916
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
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
      Left            =   570
      Top             =   0
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
            Picture         =   "AssItems.frx":0000
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AssItems.frx":01BC
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
            Picture         =   "AssItems.frx":0378
            Key             =   "Active"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AssItems.frx":07CC
            Key             =   "Inactive"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fTitle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   15
      TabIndex        =   0
      Top             =   -90
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assigned Items"
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
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assigned Items"
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
         Width           =   2295
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4785
      Left            =   15
      TabIndex        =   3
      Top             =   1185
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   8440
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item ID"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   6350
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   9172
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Rate"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSForms.CommandButton cmdAssign 
      Height          =   375
      Left            =   7140
      TabIndex        =   10
      Top             =   750
      Width           =   1680
      Caption         =   "Assign       "
      PicturePosition =   327683
      Size            =   "2963;661"
      Accelerator     =   65
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cmbRate 
      Height          =   285
      Left            =   5700
      TabIndex        =   9
      Top             =   840
      Width           =   1365
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2408;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   1320
      TabIndex        =   7
      Top             =   570
      Width           =   5745
      BackColor       =   11517387
      Caption         =   "  Contractor                            Item                                        Rate"
      Size            =   "10134;423"
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
      Top             =   6135
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
      Left            =   7005
      TabIndex        =   4
      Top             =   6135
      Visible         =   0   'False
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
End
Attribute VB_Name = "AssItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Add As Boolean, Edit As Boolean

Private Sub cmbCont_matched()
    Call RefreshList
End Sub

Private Sub cmbItem_matched()
On Error GoTo ERR
    Dim rs As New ADODB.Recordset
    cmbRate.Clear
    With rs
        .Open "select Rate1,Rate2,Rate3 from Items Where ItemID='" & cmbItem.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbRate.AddItem .Fields(0) & ""
        cmbRate.AddItem .Fields(1) & ""
        cmbRate.AddItem .Fields(2) & ""
        .Close
    End With
    Set rs = Nothing
    Exit Sub
ERR:
    MsgBox ERR.Description
End Sub

Private Sub cmdAssign_Click()
    If cmbCont.MatchFound = False Then Exit Sub
    If cmbItem.MatchFound = False Then Exit Sub
    If cmbRate.MatchFound = False Then Exit Sub
    Dim Saved As Integer
    con.Execute "Insert into AssignedItems(EmpID,ItemID,Rate) Values('" & _
                 cmbCont.ID & "','" & cmbItem.ID & "'," & cmbRate & ")", Saved
    If Saved = 1 Then
        Call RefreshList
    End If
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
        Call CheckMenu
        Me.PopUpMenu MainForm.mnupopdept, , .Left, .Top
    End With
    
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
    MainForm.mnuShowDept.Checked = False
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
'Call mnuedit_Click
    
End Sub




Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    cmbCont.ListHeight = 1600
    cmbItem.ListHeight = 1600
    cmbCont.ClearVals
    Call cmbCont.AddVals(con, "Name", "Contractors", "EmpID")
    'Call RefreshList
    'mnuAdd.Enabled = Not OutOfFinancialYear
    'mnuEdit.Enabled = Not OutOfFinancialYear
    'mnuDelete.Enabled = Not OutOfFinancialYear
    
End Sub

Private Sub Form_Resize()
 On Error Resume Next
 fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
 
 Dim TitleBottom As Long

 TitleBottom = cmbCont.Top + cmbCont.Height 'fTitle.Top + fTitle.Height
 
 LV.Move 50, TitleBottom, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + cmdClose.Height + 100)
 
 cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 50
 cmdOpt.Move cmdClose.Left - cmdOpt.Width - 75, cmdClose.Top
 
 
End Sub






Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Exit Sub
    If Button <> 2 Then Exit Sub
    Call CheckMenu
    MainForm.mnupopdept.Tag = "From Dept"
    Me.PopUpMenu MainForm.mnupopdept
End Sub








Public Sub mnuEdit(Cond As String)
    Dim f As Form
    If Cond = "DPT" Then
        Set f = New NewDepartment
        f.Tag = Right$(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 3)
        f.Add = False
        f.Editrec
        f.Show 1
        
    Else
        Load NewEmployee
        NewEmployee.Tag = LV.SelectedItem.key
        NewEmployee.Add = False
        NewEmployee.Editrec
        NewEmployee.Show 1
    End If
    Call RefreshList
End Sub
Public Sub mnuAddDept()
    Dim f As New NewDepartment
    f.TableName = "Departments"
    f.lblCatID = GetNextDeptID
    f.Add = True
    f.Show 1
    Call RefreshList
End Sub
Private Function GetNextDeptID() As String
    Dim NextVal As String
    NextVal = GetMax("cast(DeptID as int)", "Departments")
    GetNextDeptID = Format(Val(NextVal) + 1, "000")
End Function
Public Sub mnuAdd()
    Dim f As New NewEmployee
    f.Add = True
    f.Show 1
    Call RefreshList
End Sub
Public Sub RefreshIt()

    Dim Cond As String
    If MainForm.mnuShowDept.Checked Then
        Cond = ""
    Else
        Cond = " Where Active=1"
    End If
    
    Call RefreshList(Cond)
    
End Sub
Public Sub RefreshList(Optional Cond)

On Error GoTo ERR
    'Dim rs As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim PrevIndex As Double
    Dim TheColor As Long
    Dim ActiveCondition As String
    
    If IsMissing(Cond) Then
        Cond = " Where Active=1"
    End If
    'If LV.ListItems.Count > 0 Then PrevIndex = LV.SelectedItem.Index
    'Set LV.SmallIcons = ImageList1
    With rs
        .Open "Select *  from VAssignedItems Where EmpID='" & cmbCont.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            'Check For Show All Option
            'ActiveCondition = IIf(mnuShowAll.Checked, "Where Left(ID,3)='CUS'", " Where not InActive and Left(ID,3)='CUS'")
            Set ITM = LV.ListItems.Add(, , ![ItemID] & "")
            ITM.ListSubItems.Add , , ![ItemName] & ""
            ITM.ListSubItems.Add , , ![Description] & ""
            ITM.ListSubItems.Add , , ![Rate] & ""
            
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    cmbItem.ClearVals
    cmbRate.Clear
    Call cmbItem.AddVals(con, "ItemName", "Items", "ItemID", " Where ItemID not in (Select ItemID from AssignedItems Where EmpID='" & cmbCont.ID & "')")
    
    
    Exit Sub
ERR:
    MsgBox ERR.Description

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
        con.Execute "Update Departments Set Active=" & Act & " Where DeptID='" & Right(theKey, Len(theKey) - 3) & "'"
        con.Execute "Update Employees Set Active=" & Act & " Where DeptID='" & Right(theKey, Len(theKey) - 3) & "'"
        con.CommitTrans
    Else
        con.Execute "Update Employees Set Active=" & Act & " Where EmpID='" & theKey & "'"
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
        If LV.ListItems.Count = 0 Then
            .mnuDeptInact.Enabled = False
            .mnuDeptEdit.Enabled = False
            'mnuLedger.Enabled = False
            .mnuDeptPrint.Enabled = False
            'mnuPrintDetail.Enabled = False
            '.mnuAdd.Enabled = True
            .mnuDeptnew.Enabled = True
            '.mnuEmpNewEmp.Enabled = True
            'mnuInActive.Enabled = True
            'mnuPrchs.Enabled = False
            'mnuPymt.Enabled = False
        Else
            Dim DeptEmp As String
            Dim ActInAct As String
            
            DeptEmp = "Department"
            If LV.SelectedItem.ForeColor = vbBlack Then ActInAct = "Active" Else ActInAct = "InActive"
            .mnuDeptInact.Caption = "&Make This " & DeptEmp & " " & ActInAct
            
            .mnuDeptEdit.Caption = IIf(Left$(LV.SelectedItem.key, 3) = "DPT", "&Edit This Department", "&Edit This Employee")
            .mnuDeptEdit.Tag = IIf(Left$(LV.SelectedItem.key, 3) = "DPT", "DPT", "EMP")
            .mnuDeptEdit.Enabled = True
            .mnuDeptPrint.Enabled = True
            .mnuShowDept.Enabled = True
        End If
    End With

 End Sub
Public Sub showAssigned(ContID As String, ContName As String)
    
    cmbCont.Text = ContName
    Me.Show
End Sub
