VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form EmployeeList 
   ClientHeight    =   6810
   ClientLeft      =   1110
   ClientTop       =   -45
   ClientWidth     =   12990
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
   ScaleHeight     =   6810
   ScaleWidth      =   12990
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicProduct 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2865
      Left            =   9030
      ScaleHeight     =   2835
      ScaleWidth      =   2760
      TabIndex        =   16
      Top             =   780
      Visible         =   0   'False
      Width           =   2790
      Begin VB.Image ImgProduct 
         Height          =   2835
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2760
      End
   End
   Begin VB.CheckBox chkShowPicture 
      Caption         =   "Show Picture"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7290
      TabIndex        =   15
      Top             =   6405
      Width           =   1560
   End
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   315
      Left            =   2295
      TabIndex        =   12
      Top             =   285
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   556
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   315
      Left            =   75
      TabIndex        =   11
      Top             =   285
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   556
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9270
      TabIndex        =   14
      Top             =   60
      Width           =   1620
   End
   Begin VB.Frame FraSummary 
      Height          =   795
      Left            =   10965
      TabIndex        =   8
      Top             =   -90
      Width           =   1620
      Begin VB.TextBox txtTotalEmps 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   435
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total Employees"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   9
         Top             =   180
         Width           =   1395
      End
   End
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
         ItemData        =   "frmEmployeeList.frx":0000
         Left            =   60
         List            =   "frmEmployeeList.frx":001F
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
   Begin MSComctlLib.ListView LV 
      Height          =   5010
      Left            =   15
      TabIndex        =   0
      Top             =   780
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   8837
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Emp ID"
         Object.Width           =   2469
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
         Text            =   "Phone 1"
         Object.Width           =   1826
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Salary"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Employee Since"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Departement                      Employee"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   0
      Left            =   75
      TabIndex        =   13
      Top             =   45
      Width           =   9135
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
      Begin VB.Menu mnuempdash0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowInactive 
         Caption         =   "Show Inactive Employees/Departments"
      End
      Begin VB.Menu mnuempdash1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMake 
         Caption         =   "Make This Employee Inactive"
      End
      Begin VB.Menu dd1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewDept 
         Caption         =   "Add New &Department"
      End
      Begin VB.Menu mnuNewEmp 
         Caption         =   "Add New &Employee"
      End
      Begin VB.Menu mnuempdash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuEmpdash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print Employee List (With Salary)"
      End
      Begin VB.Menu mnuPrintEmployeeListDeptSelection 
         Caption         =   "Print Employee List (Department Selection)"
      End
      Begin VB.Menu mnuPrintEmpListSal 
         Caption         =   "Print Employee List (Without Salary)"
      End
      Begin VB.Menu mnuPrintEmpExt 
         Caption         =   "Print Employee List (E&xternal Salary)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintEmpSalRange 
         Caption         =   "Print Employee List (Salary Range)"
      End
      Begin VB.Menu mnuPrintEmpListSalWDOB 
         Caption         =   "Print Employee List (With DOB)"
      End
      Begin VB.Menu mnuPrintUnionMembers 
         Caption         =   "Print Employee List (Union Members)"
      End
      Begin VB.Menu mnuPrintEmpWithLunch 
         Caption         =   "Print Employee List (Mess Members)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintEmpListCF 
         Caption         =   "Print Employee List (Card Fields)"
      End
      Begin VB.Menu mnuPrintEmpListWCardID 
         Caption         =   "Print Employee List (Card ID)"
      End
      Begin VB.Menu mnuPrintEmpListSalaryPercentage 
         Caption         =   "Print Employee List (Salary Percentage)"
      End
      Begin VB.Menu PrintEmpListOldID 
         Caption         =   "Print Employee List (Old ID)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintEmpListDesignationwise 
         Caption         =   "Print Employee List (&Designationwise)"
      End
      Begin VB.Menu mnuDashPrintTemp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintBloodGroupWise 
         Caption         =   "Print Blood Group Wise"
         Begin VB.Menu mnuBloodGroupwiseSub 
            Caption         =   "mnuBloodGroupwiseSub"
            Index           =   0
         End
      End
      Begin VB.Menu mnuPrintTempEmpList 
         Caption         =   "Print Temporary Employee List"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDashLeaveBal 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEmpLeaveBalance 
         Caption         =   "Print Employees Leave Balance"
         Visible         =   0   'False
      End
      Begin VB.Menu ln1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintEmpTrans 
         Caption         =   "Print Employee List (Company Transport)"
         Visible         =   0   'False
      End
      Begin VB.Menu dashemps 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuappForm 
         Caption         =   "Print Application Form"
      End
      Begin VB.Menu mnuPrintCard 
         Caption         =   "Print Employee Card"
         Begin VB.Menu mnuPrintEmpCard 
            Caption         =   "This Employee Only"
         End
         Begin VB.Menu mnuPrintEmpBarcode 
            Caption         =   "This Employee Only (Barcode)"
            Visible         =   0   'False
         End
         Begin VB.Menu dashprintcardsdept 
            Caption         =   "-"
         End
         Begin VB.Menu mnuForAllEmployeesDept 
            Caption         =   "For All Employees of <Deptartment>"
         End
         Begin VB.Menu ln2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPrintAllCard 
            Caption         =   "For All Employees"
         End
         Begin VB.Menu mnuPrintAllBarcode 
            Caption         =   "For All Employees (Barcode)"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuTransferPermanent 
         Caption         =   "Transfer to Permanent"
      End
      Begin VB.Menu mnuPrintEmpDetail 
         Caption         =   "Print Employee Detail"
      End
   End
End
Attribute VB_Name = "EmployeeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PrevIndex As Long

Public add As Boolean, edit As Boolean
Dim m_bContractorsOnly As Boolean
Dim lTotalEmployees As Long

Private Sub chkShowPicture_Click()
    If chkShowPicture.Value = vbChecked Then
        Call LV_ItemClick(LV.SelectedItem)
    Else
        PicProduct.Visible = False
    End If
End Sub

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
        Call CheckMenu
        Me.PopUpMenu mnuPop, , .Left, .Top
    End With
    
End Sub

Private Sub cmdRefresh_Click()
    Call RefreshList
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
    If UserHasAccess("EditEmployee") Or CurrentUserName = "Administrator" Then Call EditItem(IIf(Left$(LV.SelectedItem.key, 3) = "DPT", "DPT", "EMP"))
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    
    
    With cmbDept
        .ListHeight = 4000
        .ClearVals
        .AddItem "<All Departments>", "0"
        .AddVals con, "Name", "Departments", "DeptID"
        .ID = "0"
    End With
    cmbEmp.ListHeight = 4000
    
    
    mnuPop.Visible = False
    Dim rs As New ADODB.Recordset
    Dim iIndex As Integer
    iIndex = -1
    With rs
        .Open "SELECT BloodGroup FROM BloodGroups ORDER BY BloodGroup", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            iIndex = iIndex + 1
            If .AbsolutePosition > 1 Then
                Load mnuBloodGroupwiseSub(iIndex)
                mnuBloodGroupwiseSub(iIndex).Visible = True
            End If
            mnuBloodGroupwiseSub(iIndex).Caption = !BloodGroup & ""
            mnuBloodGroupwiseSub(iIndex).Tag = !BloodGroup & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    cmbfield.ListIndex = 0
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    'fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
 
    Dim TitleBottom As Long

    'TitleBottom = 20 'fTitle.Top + fTitle.Height
    TitleBottom = FraSummary.Top + FraSummary.Height + 20
    
    LV.Move 50, TitleBottom, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + cmdClose.Height + 450)
 
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 50
    cmdOpt.Move cmdClose.Left - cmdOpt.Width - 75, cmdClose.Top
    Frame1.Top = Me.ScaleHeight - Frame1.Height - 50
    
    chkShowPicture.Move (Frame1.Left + Frame1.Width + 50), Frame1.Top + Frame1.Height - chkShowPicture.Height
    'FraSummary.Top = Me.ScaleHeight - Frame1.Height - 50
    
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
       
       
       If cmbfield.ListIndex = 0 Then
          fldValue = Trim(Replace(LV.ListItems(i), "»", ""))
       Else
          fldValue = Trim(Replace(LV.ListItems(i).SubItems(cmbfield.ListIndex), "»", ""))
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

If cmbfield.ListIndex = 0 Then
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).key Like "*" & txtValue & "*" Then
            LV.ListItems(i).Selected = True
            Exit For
        End If
    Next
Else
    For i = 1 To LV.ListItems.count
        If LCase(LV.ListItems(i).SubItems(cmbfield.ListIndex) & "") Like "*" & LCase(txtValue) & "*" Then
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

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)

    If LV.SelectedItem.Bold Then
        PicProduct.Visible = False
        Exit Sub
    End If
    
    If chkShowPicture.Value = vbChecked Then
        Dim rs As New ADODB.Recordset
        With rs
            .Open "SELECT EmpPic FROM Employees WHERE EmpID='" & Item.key & "'", con, adOpenForwardOnly, adLockReadOnly
            If Not IsNull(![EmpPic]) Then
                Set ImgProduct.DataSource = rs
                ImgProduct.DataField = "EmpPic"
            Else
                ImgProduct.Picture = LoadPicture("")
            End If
            .Close
        End With
        Set rs = Nothing
        Dim lTop As Long
        lTop = LV.Top + LV.SelectedItem.Top
        If (lTop + PicProduct.Height) > LV.Top + LV.Height Then
            lTop = LV.Top + LV.Height - PicProduct.Height
        End If
        PicProduct.Move LV.Left + LV.Width - PicProduct.Width, lTop
        PicProduct.Visible = True
    Else
'        If PicProduct.Visible = True Then PicProduct.Visible = False
    End If
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   
   If Button <> 2 Then Exit Sub
   
    Call CheckMenu
    PopUpMenu mnuPop
    
End Sub

Private Sub mnuappForm_Click()
    Call PrintAppForm
End Sub

Private Sub mnuBloodGroupwiseSub_Click(Index As Integer)
    Call PrintBloodGroupReport(mnuBloodGroupwiseSub(Index).Tag)
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

Private Sub mnuForAllEmployeesDept_Click()
    Call PrintEmpCard("", LV.SelectedItem.ListSubItems(2).Tag)
End Sub

Public Sub mnuNewDept_Click()
    Dim f As New NewDepartment
    f.TableName = "Departments"
    'f.lblCatID = GetNextDeptID
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
    Load f
'    If LV.ListItems.count > 0 Then
'        f.cmbDept.Text = LV.SelectedItem.Tag
'    End If
    
    Dim strDeptID As String
    If LV.ListItems.count > 0 Then
        strDeptID = LV.SelectedItem.Tag
    End If
    f.AddNewEmployee strDeptID, m_bContractorsOnly
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
        Cond = " WHERE Active=1"
    End If
    lTotalEmployees = 0
    If LV.ListItems.count > 0 Then PrevIndex = LV.SelectedItem.Index
'    If m_bContractorsOnly Then
'        rsDept.Open "SELECT Deptid,Name,Active,Color from Departments Where TempDept=1", con, adOpenForwardOnly, adLockReadOnly
'    Else
    If Cond = "" Then
        Cond = " WHERE IsNull(TempDept,0)=0"
    Else
        Cond = Cond & " AND IsNull(TempDept,0)=0"
    End If
    
    If cmbDept.ID <> "0" Then
        Cond = Cond & " AND DeptID='" & cmbDept.ID & "'"
    End If
    
    rsDept.Open "SELECT Deptid,Name,Active,Color FROM Departments " & Cond, con, adOpenForwardOnly, adLockReadOnly
'    End If

    
    LV.ListItems.Clear
    Dim lDeptColor As Long
    Do Until rsDept.EOF
        With rs
            TheColor = Val(rsDept![Color] & "")
            lDeptColor = TheColor
            If mnuShowInactive.Checked = True Then
                Cond = ""
            Else
                Cond = " AND Active=1"
            End If
            
            If m_bContractorsOnly Then
                Cond = Cond & " AND TempEmp=1"
            Else
                Cond = Cond & " AND TempEmp=0"
            End If
            If cmbEmp.ID <> "0" Then
                Cond = Cond & " AND EmpID='" & cmbEmp.ID & "'"
            End If
            
            .Open "SELECT EmpID,Name,Active,FName,Designation,Religion,Sex,StartingSalary,Phone1,Phone2,DeptID,JoinDate FROM VEmp WHERE DeptID='" & rsDept.Fields(0) & "'" & Cond & " ORDER BY EmpID", con, adOpenForwardOnly, adLockReadOnly
            Set ITM = LV.ListItems.add(, "DPT" & rsDept.Fields(0), rsDept.Fields(0))
            ITM.Bold = True
            ITM.Tag = rsDept!Name
        
            ITM.ForeColor = TheColor
            ITM.ListSubItems.add(, , rsDept.Fields(1) & "").Bold = True
            ITM.ListSubItems(1).ForeColor = TheColor
            ITM.ListSubItems(1).Tag = rsDept![Active] & ""
            ITM.ListSubItems.add(, , "").Tag = rsDept!DeptID
            If Not rsDept![Active] Then
                ITM.ForeColor = vbBlack
            End If
            
            For i = 0 To .RecordCount - 1
                If IsNull(![EmpId]) Then Exit For
                Set ITM = LV.ListItems.add(, ![EmpId] & "", "  » " & ![EmpId])
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
                ITM.ListSubItems(2).Tag = !DeptID
                'itm.ListSubItems(2).ForeColor = TheColor
                
                ITM.ListSubItems.add(, , ![Phone1] & " ").ForeColor = TheColor
                
                ITM.ListSubItems.add(, , Format(Val(!StartingSalary & ""), "#,##")).ForeColor = TheColor
                ITM.ListSubItems.add(, , Format(![JoinDate], "dd-MMM-yyyy")).ForeColor = TheColor
                'itm.ListSubItems.Add , , ![Phone1] & " "
                'itm.ListSubItems.Add(, , Format(Val(![Balance] & ""), "0.00 DR;0.00 CR;0.00     ")).ForeColor = TheColor
                lTotalEmployees = lTotalEmployees + 1
                .MoveNEXT
            Next
            .Close
            rsDept.MoveNEXT
        End With
    Loop
    Set rs = Nothing
    
    txtTotalEmps = GetSingleLongValue("COUNT(EmpID)", "Employees", "WHERE Active=1") 'lTotalEmployees
    
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
    
    If CBool(LV.SelectedItem.ListSubItems(1).Tag) = False Then Act = True Else Act = False
    
    Dim DisContOn As String
    If Act Then DisContOn = "NULL" Else DisContOn = "'" & Format(Date, "dd-MMM-yyyy") & "'"
        
        
'    If Left$(theKey, 3) = "DPT" Then
'
'        Dim Childs As Integer
'        Dim rs As New ADODB.Recordset
'        With rs
'            .Open "Select Count(EmpID) From Employees Where DeptID='" & Right(theKey, Len(theKey) - 3) & "'", con, adOpenForwardOnly, adLockReadOnly
'            If IsNull(.Fields(0)) Then Childs = 0 Else Childs = .Fields(0)
'            .Close
'        End With
'        Set rs = Nothing
'
'        If MsgBox("This Will Affect All The Employees Of This Department." & _
'         vbNewLine & "There Are " & Childs & _
'         " Employees In This Department." & vbNewLine & "Do You Want To Continue?", vbQuestion + vbYesNo) = vbNo Then
'            Exit Sub
'        End If
'        con.BeginTrans
'
'        con.Execute "Update Departments set Active=" & Abs(Act) & " Where DeptID='" & Right(theKey, Len(theKey) - 3) & "'"
'        con.Execute "Update Employees Set Discontinuedon=" & DisContOn & ",Active=" & Abs(Act) & " Where DeptID='" & Right(theKey, Len(theKey) - 3) & "'"
'        con.CommitTrans
'    Else
        If Not Act Then
            If MsgBox("This Is Usually Done When An Employee Discontinues. Do You Want To Continue ?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
            Dim f As Form
            If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                Set f = New frmMakeEmployeeInActive
                Dim strRemarks As String
                If f.getDate(theKey, DCDate, strRemarks) = False Then Exit Sub
                Call StartTrans(con)
                
                con.Execute "UPDATE Employees SET Active=0 WHERE EmpID='" & theKey & "'"
                
                con.Execute "INSERT INTO EmployeeInActiveDetails(EmpID,DT,InAciveRemarks" & _
                ",UserName,MachineName) VALUES('" & theKey & "','" & DCDate & "','" & _
                strRemarks & "','" & CurrentUserName & "','" & strComputerName & "')"
                 
                con.CommitTrans
            Else
                Set f = New frmGetDateSingle
                If f.getDate(DCDate) = False Then
                    MsgBox "Invalid Date.", vbCritical
                    Exit Sub
                End If
                con.Execute "UPDATE Employees SET DiscontinuedOn='" & DCDate & "',Active=" & Abs(Act) & " Where EmpID='" & theKey & "'"
            End If
        Else
            con.Execute "UPDATE Employees SET Active=1 WHERE EmpID='" & theKey & "'"
            'If strCompany = "Dr-Frgz" Then
                
'            Else
'                DCDate = InputBox("Plz Enter Date[mm/dd/yy] To ReJoin.", "Re Joining After Discontinuation.")
'                If Not IsDate(DCDate) Then
'                    MsgBox "Invalid Date.", vbCritical
'                    Exit Sub
'                End If
'                con.Execute "Update Employees Set JoinDate='" & DCDate & "',DiscontinuedOn=Null,Active=" & Abs(Act) & " Where EmpID='" & theKey & "'"
            
            'End If
        'End If
        
        
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

Private Sub CheckMenu()
    
    
    
   With Me
        
        .mnuNewEmp.Caption = "Add New &Employee"
        .mnuShowInactive.Caption = "Show Inactive Employees/Departments"
        .mnuMake.Caption = "Make This Employee Inactive"
        .mnuPrint.Caption = "&Print Employee List (With Salary)"
        
        If LV.ListItems.count = 0 Then
            .mnuEdit.Enabled = False
            .mnuPrint.Enabled = False
            .mnuMake.Enabled = False
            .mnuNewDept.Enabled = True
            .mnuNewEmp.Enabled = True
        Else
            Dim DeptEmp As String
            Dim ActInAct As String
            
            If Left$(LV.SelectedItem.key, 3) = "DPT" Then DeptEmp = "Department" Else DeptEmp = "Employee"
            If LV.SelectedItem.ForeColor = vbRed Then ActInAct = "Active" Else ActInAct = "InActive"
            .mnuMake.Caption = "&Make This " & DeptEmp & " " & ActInAct
            
            .mnuEdit.Caption = IIf(Left$(LV.SelectedItem.key, 3) = "DPT", "&Edit This Department", "&Edit This Employee")
            .mnuEdit.Tag = IIf(Left$(LV.SelectedItem.key, 3) = "DPT", "DPT", "EMP")
            .mnuEdit.Enabled = UserHasAccess("EditEmployee")
            
            .mnuMake.Enabled = UserHasAccess("EmpActive")
            .mnuNewEmp.Enabled = UserHasAccess("NewEmployee")
            .mnuNewDept.Enabled = UserHasAccess("NewDept")
            .mnuPrint.Enabled = UserHasAccess("PrintEmpList")
            .mnuPrintCard.Enabled = UserHasAccess("PrintEmpCards")
            .mnuPrintEmpExt.Enabled = UserHasAccess("PrintEmpList")
            .mnuappForm.Enabled = UserHasAccess("ApplicationForm")
            .mnuForAllEmployeesDept.Caption = "For All Employees of " & LV.SelectedItem.Tag
          End If
    End With
    
    If m_bContractorsOnly And Left$(LV.SelectedItem.key, 3) <> "DPT" Then
        mnuTransferPermanent.Visible = True
    Else
        mnuTransferPermanent.Visible = False
    End If
    
End Sub



Public Sub PrintEmpList(Optional WithExternalSalary As Boolean = False, Optional Transport As Boolean = False, Optional bOldID As Boolean = False, Optional bDesignationwise As Boolean = False)
    
    Dim SelFormula As String
    If mnuShowInactive.Checked Then
        SelFormula = IIf(Transport, " {VEmp.Transport}=true", "")
    Else
        SelFormula = "{VEmp.Active}=True " & IIf(Transport, " and {VEmp.Transport}=true", "")
    End If
    
    If m_bContractorsOnly Then
        SelFormula = SelFormula & " AND {VEmp.TempDept}=True"
    Else
        SelFormula = SelFormula & " AND {VEmp.TempDept}=False"
    End If
    
    Screen.MousePointer = vbHourglass

    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    If bOldID Then
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpListWithOldID.rpt")
    ElseIf bDesignationwise Then
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpListDesignationwise.rpt")
    Else
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpList.rpt")
    End If
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & strCompany & "'"
        End If
    Next

    f.ShowReport SelFormula, rpt

    Screen.MousePointer = vbDefault

End Sub

Public Sub PrintEmpListWithoutSal(Optional WithDOB As Boolean = False, Optional SelFormula As String, Optional ReportCaption As String = "Employee List", Optional TempEmpList As Boolean = False)

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
    
    If m_bContractorsOnly Then
        SelFormula = SelFormula & " AND {VEmp.TempDept}=True"
    Else
        SelFormula = SelFormula & " AND {VEmp.TempDept}=False"
    End If
    
    If TempEmpList Then
        SelFormula = SelFormula & " AND {@TempEmployee}=True"
    End If
    
    Screen.MousePointer = vbHourglass

    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpListWithoutSal.rpt")
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & strCompany & "'"
        ElseIf FormulaField.Name = "{@BirthDate}" Then
            FormulaField.Text = WithDOB
        ElseIf FormulaField.Name = "{@RptCaption}" Then
            FormulaField.Text = "'" & ReportCaption & "'"
        End If
    Next

    f.ShowReport SelFormula, rpt

    Screen.MousePointer = vbDefault

End Sub


Public Sub PrintEmpCard(Optional EmpId As String, Optional DeptID As String)

    Dim Cond As String

    If EmpId = "" Then
        Cond = "{Employees.Active}=True"
        If m_bContractorsOnly Then
            Cond = Cond & " AND {Employees.TempDept}=True"
        Else
            Cond = Cond & " AND {Employees.TempDept}=False"
        End If
    Else
        If EmpId = "" Then
            Cond = "{Employees.EmpID}='" & Trim(Replace(EmpId, "»", "")) & "' and {Employees.Active}=True"
        Else
            Cond = "{Employees.EmpID}='" & Trim(Replace(EmpId, "»", "")) & "' "
        End If
    End If
     
    If DeptID <> "" Then
        Cond = "{Employees.DeptID}='" & DeptID & "' AND {Employees.Active}=TRUE"
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
    
'    Dim rs As New ADODB.Recordset
'    Dim TempEmp As Boolean
'    With rs
'        .Open "Select TempDept  From Departments Where DeptID=(Select DeptID From Employees Where EmpID='" & LV.SelectedItem.key & "')", con, adOpenForwardOnly, adLockReadOnly
'
'        TempEmp = .Fields(0)
'
'        .Close
'    End With
'    Set rs = Nothing
'
'    If TempEmp = False Then
'        MsgBox "This Option Is Not Available." & vbNewLine & "Use 'Discontinue Employee' From Employees.", vbInformation
'        Exit Sub
'    End If
    
    'If getPassword Then
    Call MakeInactive
    'End If
    
End Sub

Private Sub mnuPrint_Click()
    Call PrintEmpList
End Sub

Private Sub mnuPrintAllBarcode_Click()
    
    If LV.ListItems.count = 0 Then Exit Sub
    frmPrint.Show 1

End Sub

Private Sub mnuPrintAllCard_Click()
    Call PrintEmpCard
End Sub

Private Sub mnuPrintEmpBarcode_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    If LV.SelectedItem.Bold Then Exit Sub
    Call PrintBarcode("WHERE Active=1 AND EmpID='" & LV.SelectedItem.key & "' ORDER BY EmpID", 1, 1)

End Sub

Private Sub mnuPrintEmpCard_Click()
    If LV.ListItems.count = 0 Then Exit Sub
    If LV.SelectedItem.Bold Then Exit Sub
    Call PrintEmpCard(LV.SelectedItem)
End Sub

Private Sub mnuPrintEmpDetail_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpDetail.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{Employees.empid}='" & LV.SelectedItem.key & "'", rpt

End Sub

Private Sub mnuPrintEmpExt_Click()
    Call PrintEmpList(True)
End Sub

Private Sub mnuPrintEmpListCF_Click()
    
    Dim SelFormula As String
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpCardFields.rpt")
    Dim f As New frmPrevRpt
    
    SelFormula = IIf(mnuShowInactive.Checked, " ", "{VEmp.Active}=True ")
    
    Dim bFormula As Boolean
    If m_bContractorsOnly Then
        bFormula = True
    Else
        bFormula = False
    End If
    
    If SelFormula = "" Then
        SelFormula = "{VEmp.TempDept}=" & bFormula
    Else
        SelFormula = SelFormula & " AND {VEmp.TempDept}=" & bFormula
    End If
    
    f.ShowReport SelFormula, rpt
    
End Sub

Private Sub mnuPrintEmpListDesignationwise_Click()
    Call PrintEmpList(, , , True)
End Sub

Private Sub mnuPrintEmpListSal_Click()

    PrintEmpListWithoutSal

End Sub

Private Sub mnuPrintEmpListSalaryPercentage_Click()
        
    Dim SelFormula As String, iPercentage As Integer
    iPercentage = InputBox("Please Enter Percentage", "Salarey Percentage", "75")
    
    SelFormula = "{VEmp.Active}=True AND {VEmp.TempDept}=False"
     
    Screen.MousePointer = vbHourglass

    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpListPercentage.rpt")
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & strCompany & "'"
        ElseIf FormulaField.Name = "{@SalaryPercentageHead}" Then
            FormulaField.Text = "'Salary @ " & iPercentage & "%'"
        ElseIf FormulaField.Name = "{@iPercentage}" Then
            FormulaField.Text = iPercentage
        End If
    Next

    f.ShowReport SelFormula, rpt

    Screen.MousePointer = vbDefault

End Sub

Private Sub mnuPrintEmpListSalWDOB_Click()
    PrintEmpListWithoutSal True, , "Employee List (With DOB)"
End Sub

Private Sub mnuPrintEmpListWCardID_Click()
        
    Dim SelFormula As String
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpListWCardID.rpt")
    Dim f As New frmPrevRpt
    
    SelFormula = IIf(mnuShowInactive.Checked, "", "{VEmp.Active}=True ")
    
    Dim bFormula As Boolean
    If m_bContractorsOnly Then
        bFormula = True
    Else
        bFormula = False
    End If
    
    If Trim(SelFormula) = "" Then
        SelFormula = "{VEmp.TempDept}=" & bFormula
    Else
        SelFormula = SelFormula & " AND {VEmp.TempDept}=" & bFormula
    End If
    
    f.ShowReport SelFormula, rpt

End Sub

Private Sub mnuPrintEmployeeListDeptSelection_Click()

    Dim strDepts As String
    Dim f As New frmGetDepartments
    Load f
    If f.getData(strDepts) = False Then Exit Sub
    
    Dim strSelection As String
    strSelection = "{VEmp.Active}=TRUE"
    If strDepts <> "All" Then
        strSelection = strSelection & " AND {VEmp.DeptID} IN [" & strDepts & "]"
    End If
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpList.rpt")
    Dim frm As New frmPrevRpt
    frm.ShowReport strSelection, rpt
    
End Sub

Private Sub mnuPrintEmpSalRange_Click()
    frmPrintEmpListWithSalRange.Show
End Sub

Private Sub mnuPrintEmpTrans_Click()
    Call PrintEmpList(False, True)
End Sub

Private Sub mnuPrintEmpWithLunch_Click()
    PrintEmpListWithoutSal False, "{Vemp.Lunch}=True", "Mess Members"
End Sub

Private Sub mnuPrintTempEmpList_Click()
    PrintEmpListWithoutSal , , "Temporary Employee List", True
End Sub

Private Sub mnuPrintUnionMembers_Click()
    PrintEmpListWithoutSal False, "{Vemp.UnionFund}=True", "Union Members"
End Sub

Private Sub mnuRefresh_Click()

    RefreshList

End Sub

Private Sub mnuShowInactive_Click()

    Call ShowInactive(mnuShowInactive)

End Sub

Private Sub mnuTransferPermanent_Click()
    If MsgBox("Are You Sure ?", vbQuestion + vbYesNo) = vbYes Then
        Dim strTemp As String
        If strCompany = "PAKSMITH GROUP" Then
            strTemp = "PSS"
        ElseIf strCompany = "Weldon Instruments" Then
            strTemp = "WIS"
        ElseIf strCompany = "Sunlike" Then
            strTemp = "SLI"
        ElseIf strCompany = "QEL" Then
            strTemp = "QEL"
        ElseIf strCompany = "Dr-Frgz" Then
            If strSubCompany = "Intl" Then
                strTemp = "DF"
            Else
                strTemp = "DFG"
            End If
        ElseIf strCompany = "Tecno" Then
            strTemp = "TCN"
        ElseIf strCompany = "Towne" Then
            strEmp = "TWN"
        End If
        
        Call StartTrans(con)
        If strCompany = "Tecno" Then
            con.Execute "UPDATE Employees SET TempEmp=0,EmpID=REPLACE(EmpID,'TMP','" & strTemp & "') WHERE EmpID='" & LV.SelectedItem.key & "'"
        Else
            con.Execute "UPDATE Employees SET TempEmp=0,EmpID=REPLACE(EmpID,'T','" & strTemp & "') WHERE EmpID='" & LV.SelectedItem.key & "'"
        End If
        con.CommitTrans
        RefreshList
    End If
End Sub

Private Sub PrintEmpListOldID_Click()
    Call PrintEmpList(, , True)
End Sub

Private Sub txtValue_GotFocus()

    txtValue.SelStart = 0
    txtValue.SelLength = Len(txtValue)

End Sub

Public Sub ShowMe(p_bContractorsOnly As Boolean)

    Dim strTitle As String
    m_bContractorsOnly = p_bContractorsOnly
    If m_bContractorsOnly Then
        strTitle = "Temp. Employee List"
    Else
        strTitle = "Employee List"
    End If
    Call RefreshList
    Me.Show
    
End Sub

Private Sub cmbdept_matched()

    cmbEmp.ClearVals
    cmbEmp.AddItem "<All Employees>", "0"
    
    Dim strCondition As String
    strCondition = " WHERE 1=1"
    If mnuShowInactive.Checked = False Then
        strCondition = strCondition & " AND Active=1"
    End If
    
    If cmbDept.ID = "0" Then
        
    Else
        strCondition = strCondition & " AND DeptID='" & cmbDept.ID & "'"
    End If
    cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' - ' + FName + ' {' + Designation + '}'", "Employees", "EmpID", strCondition
    cmbEmp.ID = "0"
    
End Sub

