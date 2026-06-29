VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmShortTermLedger 
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   11850
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleMode       =   0  'User
   ScaleWidth      =   10614.2
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   10440
      TabIndex        =   15
      Top             =   60
      Width           =   1365
      Begin VB.TextBox txtTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total"
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
         Left            =   420
         TabIndex        =   16
         Top             =   150
         Width           =   435
      End
   End
   Begin Crystal.CrystalReport CR 
      Left            =   705
      Top             =   1230
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
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   15
      TabIndex        =   5
      Top             =   405
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
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   2235
      TabIndex        =   6
      Top             =   405
      Width           =   3750
      _ExtentX        =   6615
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
   Begin VB.Frame FLV 
      Height          =   6045
      Left            =   30
      TabIndex        =   0
      Top             =   780
      Width           =   11715
      Begin VB.Frame fButs2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   510
         Left            =   165
         TabIndex        =   11
         Top             =   5460
         Width           =   5535
         Begin MSForms.CommandButton cmdNewLT 
            Height          =   360
            Left            =   2760
            TabIndex        =   13
            Top             =   90
            Width           =   2700
            ForeColor       =   4194304
            Caption         =   "New Long Term Advance          "
            PicturePosition =   327683
            Size            =   "4762;635"
            Accelerator     =   78
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdNewST 
            Height          =   360
            Left            =   30
            TabIndex        =   12
            Top             =   90
            Width           =   2700
            ForeColor       =   4194304
            Caption         =   "New Short Term Advance"
            PicturePosition =   327683
            Size            =   "4762;635"
            Picture         =   "frmShortTermLedger.frx":0000
            Accelerator     =   78
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.Frame fButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   465
         Left            =   8595
         TabIndex        =   8
         Top             =   5475
         Width           =   3030
         Begin MSForms.CommandButton cmdClose 
            Height          =   360
            Left            =   1530
            TabIndex        =   10
            Top             =   75
            Width           =   1425
            ForeColor       =   0
            Caption         =   " Close      "
            PicturePosition =   327683
            Size            =   "2514;635"
            Accelerator     =   67
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdPrint 
            Height          =   360
            Left            =   60
            TabIndex        =   9
            Top             =   75
            Width           =   1425
            ForeColor       =   0
            Caption         =   "Print         "
            PicturePosition =   327683
            Size            =   "2514;635"
            Picture         =   "frmShortTermLedger.frx":0112
            Accelerator     =   80
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin Crystal.CrystalReport cr1 
         Left            =   3840
         Top             =   195
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
      Begin MSComctlLib.ListView LV 
         Height          =   5265
         Left            =   105
         TabIndex        =   1
         Top             =   195
         Width           =   11520
         _ExtentX        =   20320
         _ExtentY        =   9287
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Emp ID"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Emp Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Loan Type"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Vchr #"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Balance"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Ded. Month"
            Object.Width           =   2399
         EndProperty
      End
   End
   Begin MSComCtl2.DTPicker DtTo 
      Height          =   300
      Left            =   7440
      TabIndex        =   2
      Top             =   405
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   167968771
      CurrentDate     =   37081
   End
   Begin MSComCtl2.DTPicker DtFrom 
      Height          =   300
      Left            =   6000
      TabIndex        =   3
      Top             =   405
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   167968771
      CurrentDate     =   37055
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortTermLedger.frx":0224
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortTermLedger.frx":0676
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmShortTermLedger.frx":0788
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
      Index           =   4
      Left            =   30
      TabIndex        =   14
      Top             =   165
      Width           =   8775
   End
   Begin MSForms.ComboBox cmbShift 
      Height          =   285
      Left            =   -570
      TabIndex        =   7
      Top             =   495
      Visible         =   0   'False
      Width           =   1500
      VariousPropertyBits=   679495705
      DisplayStyle    =   7
      Size            =   "2646;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   8895
      TabIndex        =   4
      Top             =   405
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
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Begin VB.Menu mnuLoansList 
         Caption         =   "Loans List"
      End
      Begin VB.Menu mnuPayslip 
         Caption         =   "Loan Payslip"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnupop1 
      Caption         =   "mnupop1"
      Begin VB.Menu mnuConvertToLT 
         Caption         =   "Convert to Long Term"
      End
   End
End
Attribute VB_Name = "frmShortTermLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdNewLT_Click()
    With NewAdvanceShort
      .LongTerm = True
      .cmbDept.ID = cmbDept.ID
      .cmbEmp.ID = cmbEmp.ID
      .Show 1
      Call cmdRefresh_Click
   End With
    
End Sub

Private Sub cmdNewST_Click()
    With NewAdvanceShort
        .LongTerm = False
        .cmbDept.ID = cmbDept.ID
        .cmbEmp.ID = cmbEmp.ID
        .Show 1
        Call cmdRefresh_Click
    End With
End Sub

Private Sub cmdPrint_Click()
            
    If strCompany = "Banzai" Then
        Call PrintEmpShortTermLedger
        Exit Sub
    End If
   
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\ShortTermLoanLedger.rpt")
    Dim f As New frmPrevRpt

    Dim strSelection As String
    strSelection = "{VAdvances.DT}=#" & DtFrom & "# TO #" & DtTo & "# AND {VAdvances.Type}=0"

    If cmbEmp.ID = "0" Then
        If cmbDept.ID <> "0" Then
            strSelection = strSelection & " AND {VAdvances.DeptID}='" & cmbDept.ID & "'"
        End If
    Else
        strSelection = strSelection & " AND {VAdvances.EmpID}='" & cmbEmp.ID & "'"
    End If

    f.ShowReport strSelection, rpt
    
    'PopUpMenu mnupopup

End Sub

Private Sub PrintEmpShortTermLedger()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\EmpShortTermLoanLedger.rpt")
    Dim f As New frmPrevRpt

    Dim strSelection As String
    strSelection = "{VEmp.EmpID}='" & cmbEmp.ID & "'"

    f.ShowReport strSelection, rpt
    
End Sub

Public Sub cmdRefresh_Click()
    
    Call RefreshLedger

End Sub

Private Sub Form_Load()

    mnuPopup.Visible = False
    mnupop1.Visible = False

    cmdNewST.Enabled = UserHasAccess("ShortTermLoan")
    cmdNewLT.Enabled = UserHasAccess("LongTermLoan")

    DtTo = GetServerDate
    DtFrom = DateAdd("M", -1, GetServerDate)
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    
    'cmbShift.AddItem "<All>"
    cmbShift.AddItem "Short Term"
    cmbShift.AddItem "Long Term"
    cmbShift.ListIndex = 0
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID"
    cmbDept.Text = "<All Departments>"
    cmbEmp.Text = "<All Employees>"
    'Set LV.SmallIcons = ImageList1
    
End Sub
Private Sub cmbdept_matched()

    cmbEmp.ClearVals
    cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.Text = "<All Departments>" Then
        cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' - ' + FName + ' {' + Designation + '}'", "Employees", "EmpID", "Where Active=1"
    Else
        cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' - ' + FName + ' {' + Designation + '}'", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "' and Active=1"
    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    Dim TopMargin As Integer
    TopMargin = cmbEmp.Top + cmbEmp.Height + 30
    
    FLV.Move 50, TopMargin, Me.ScaleWidth - 100, Me.ScaleHeight - TopMargin
    
    LV.Move 100, 200, Me.Width - 600, FLV.Height - fButs1.Height - 300
    
    fButs1.Move LV.Left + LV.Width - fButs1.Width, LV.Top + LV.Height
    
    fButs2.Move LV.Left, LV.Top + LV.Height

 
End Sub



Private Sub RefreshLedger()

    If cmbDept.MatchFound = False Or cmbEmp.MatchFound = False Then Exit Sub
    Dim rsLedger As New ADODB.Recordset
    Dim ITM As ListItem
    Dim Cond As String
    Dim lTotal As Long
    Dim ICO As String
    
    Cond = " WHERE Type=" & Val(cmbShift.ListIndex)
     
    If cmbEmp.ID = "0" Then
        If cmbDept.ID = "0" Then
            Cond = Cond & " and DT between '" & DtFrom & "' and '" & DtTo & "' "
        Else
            Cond = Cond & " and DT between '" & DtFrom & "' and '" & DtTo & "' and DeptID='" & cmbDept.ID & "'"
        End If
    Else
        Cond = Cond & " and empid='" & cmbEmp.ID & "' and DT between '" & DtFrom & "' and '" & DtTo & "' "
    End If
    
    
    
    SQL = "SELECT * FROM VAdvances " & Cond & " ORDER BY EmpID,DT,Type ASC"
    
    Dim dLastSavedSalaryDT As Date
    dLastSavedSalaryDT = GetSingleDateValue("MAX(DT)", "MonthlySalaries")
    '--------------------------------------------------------------------
    Dim LColor As Long
    With rsLedger
        
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            'If ![Attendance] Then
            '    ICO = "Present"
            'Else
            '    ICO = "Absent"
            'End If
            Set ITM = LV.ListItems.add(, ![EntryID] & "'", ![EmpId] & "")
            ITM.Tag = !EntryID & ""
            ITM.ListSubItems.add(, , ![Name] & "").Tag = Format(![DT], "dd-MMM-yyyy")
            ITM.ListSubItems.add(, , Format(![DT], "dd-MMM-yyyy")).Tag = Val(!Cleared & "")
            ITM.SubItems(3) = IIf(![Type] = 0, "Short Term", "Long Term")
            ITM.SubItems(4) = ![Amount] & ""
            ITM.SubItems(5) = ![AccVoucherNo] & ""
            ITM.SubItems(6) = Val(![Amount] & "") - Val(![AmountCleared] & "")
            ITM.SubItems(7) = Format(!DeductionStartDT, "MMM-yyyy")
            If Val(!Cleared & "") = 0 Then
                LColor = LV.ForeColor
            Else
                LColor = vbBlue
            End If
            ITM.ForeColor = LColor
            For j = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(j).ForeColor = LColor
            Next
            lTotal = lTotal + Val(!Amount & "")
            .MoveNEXT
        Next
        .Close
    End With
    txtTotal = Format(lTotal, "#,##")
    
    If cmbShift.ListIndex = 0 Then
        LV.ColumnHeaders(7).Width = 0
    Else
        LV.ColumnHeaders(7).Width = 1200
    End If
    
    Set rsLedger = Nothing
    
    Cond = ""
    ICO = ""
    
End Sub






Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopUpMenu mnupop1
    End If
End Sub

Private Sub mnuDelEntry_Click()
        
    MsgBox "Are you kidding?", vbInformation
    Exit Sub
    
    If LV.SelectedItem Is Nothing Then Exit Sub
        
    Dim DT As Date
    
    
    With LV.SelectedItem
        
        DT = DateSerial(year(CDate(.ListSubItems(1).Tag)), Month(CDate(.ListSubItems(1).Tag)), 1)
        DT = DateSerial(year(DT), Month(DT), DateDiff("d", DT, DateAdd("M", 1, DT)))
    
    
        With con.Execute("Select count(*) from MonthlySalaries Where DT>='" & Format(DT, "dd-MMM-yyyy") & "'")
            If Val(.Fields(0) & "") > 0 Then
                MsgBox "Cannot Remove The Entry Related To The Saved Salary", vbInformation
                .Close
                Exit Sub
            End If
            .Close
        End With
    
        If MsgBox("Are you sure to Remove this Entry ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        con.Execute "Delete from Advances Where EntryID=" & Val(.Tag)
        Call cmdRefresh_Click
    End With
End Sub

Private Sub mnuConvertToLT_Click()

'    If Val(LV.SelectedItem.Tag) <> 0 Then
'        MsgBox "Can't Do !!!", vbInformation
'        Exit Sub
'    End If
    If Val(LV.SelectedItem.ListSubItems(2).Tag) <> 0 Then
        MsgBox "Can't Do !!!", vbInformation
        Exit Sub
    End If
    
    Dim f As New frmConvertShortToLongTerm
    Load f
    f.ShowMe Val(LV.SelectedItem.key)
    
End Sub

Private Sub mnuLoansList_Click()
    Dim Cond As String
    
    
    
'    If cmbEmp.ID = "0" Then
'        If cmbDept.ID = "0" Then
'            Cond = "({VAdvances.DT}=date(" & Format(DTFrom, "yyyy,MM,dd") & ") TO date(" & Format(DTTo, "yyyy,MM,dd") & "))"
'        Else
'            Cond = "({VAdvances.DT}=date(" & Format(DTFrom, "yyyy,MM,dd") & ") TO date(" & Format(DTTo, "yyyy,MM,dd") & ")) and {VAdvances.DeptID}='" & cmbDept.ID & "'"
'        End If
'    Else
'        Cond = "{VAdvances.empid}='" & cmbEmp.ID & "' AND ({VAdvances.DT}=date(" & Format(DTFrom, "yyyy,MM,dd") & ") TO date(" & Format(DTTo, "yyyy,MM,dd") & "))"
'    End If
'    Cond = Cond & "AND {VAdvances.Type}=0"
'
'    If cmbShift.ListIndex > 0 Then
'      Cond = Cond & " and {VAdvances.Type}=" & cmbShift.ListIndex
'
'    End If
'With cr1
'    .ReportFileName = RptPath & "\LoanLedger.rpt"
'    '.DataFiles(0) = rptpath & "\Payroll.mdb"
'    .Connect = ConnectStr
'    .SelectionFormula = Cond
'
'    .Formulas(0) = "FromTo='From " & Format(DTFrom, "dd-MMM-yyyy") & " To " & Format(DTTo, "dd-MMM-yyyy") & "'"
'    .Formulas(1) = "Company='" & strCompany & "'"
'
'    .Action = 1
'    .PageZoomNext
'End With


    Screen.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptLoanLedger.rpt")
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
    
    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & strCompany & "'"
        ElseIf FormulaField.Name = "{@FromTo}" Then
            FormulaField.Text = "'From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
        End If
    Next
    If cmbEmp.ID = "0" Then
        Cond = ""
    Else
        Cond = "{VAdvances.EmpID}='" & cmbEmp.ID & "'"
    End If
'    On Error Resume Next
'    SrcReport.Database.Tables(1).ConnectBufferString = "DSN=DSN_Vision;;User ID=sa;;Database=Vision;;Password=visionasim"
'    rpt.SQLQueryString = SQL
    f.ShowReport Cond & " ", rpt
    
    
    Screen.MousePointer = vbDefault


End Sub

Private Sub mnuPaySlip_Click()

If Val(LV.SelectedItem.Tag) = 0 Then Exit Sub

'With Cr
'    .ReportFileName = RptPath & "\AdvancesPayslip.rpt"
'    '.DataFiles(0) = rptpath & "\Payroll.mdb"
'    .Connect = ConnectStr
'    .SelectionFormula = "{VAdvances.EntryID}=" & LV.SelectedItem.Tag & ""
'
''    .Formulas(0) = "FromTo='From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
'    .Formulas(0) = "Company='" & strCompany & "'"
'
'    .Action = 1
'    .PageZoomNext
'End With

Screen.MousePointer = vbHourglass

Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptAdvancesPayslip.rpt")
Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

Set FormulaFields = rpt.FormulaFields
For Each FormulaField In FormulaFields
    If FormulaField.Name = "{@Company}" Then
        FormulaField.Text = "'" & strCompany & "'"
    ElseIf FormulaField.Name = "{@FromTo}" Then
        FormulaField.Text = "'From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
    End If
Next

f.ShowReport "{VAdvances.EntryID}=" & LV.SelectedItem.Tag & " ", rpt


Screen.MousePointer = vbDefault


End Sub
