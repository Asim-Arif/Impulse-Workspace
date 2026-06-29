VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{47AA402D-55E1-44AC-B1E8-CCCB5C56AC58}#1.0#0"; "ComboList.ocx"
Begin VB.Form AdvanceSalary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advance Salary..."
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleMode       =   0  'User
   ScaleWidth      =   10667.95
   WindowState     =   2  'Maximized
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   330
      Left            =   855
      TabIndex        =   6
      Top             =   405
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   582
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   330
      Left            =   3075
      TabIndex        =   7
      Top             =   405
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   582
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin VB.Frame FLV 
      Height          =   6135
      Left            =   45
      TabIndex        =   0
      Top             =   810
      Width           =   11835
      Begin Crystal.CrystalReport cr1 
         Left            =   120
         Top             =   2865
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
         Height          =   5310
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   11685
         _ExtentX        =   20611
         _ExtentY        =   9366
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
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
         NumItems        =   17
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Emp ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Emp Name"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "B. Sal"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Rate"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "A.Days"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "A. Amt"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "S.days"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "S. Amt"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "O.Hrs"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "O.Amt"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "L.Hrs"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Text            =   "L.Amt"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   12
            Text            =   "Total"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   13
            Text            =   "Tax %"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   14
            Text            =   "T. Amt"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   15
            Text            =   "Net Ttl"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   16
            Text            =   "Paid"
            Object.Width           =   1411
         EndProperty
      End
      Begin MSForms.CommandButton cmdNew 
         Height          =   360
         Left            =   8820
         TabIndex        =   8
         Top             =   5670
         Visible         =   0   'False
         Width           =   1425
         ForeColor       =   0
         Caption         =   "New Entry   "
         PicturePosition =   327683
         Size            =   "2514;635"
         Picture         =   "AdvanceSalary.frx":0000
         Accelerator     =   78
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   360
         Left            =   10320
         TabIndex        =   3
         Top             =   5670
         Width           =   1425
         ForeColor       =   0
         Caption         =   " Close      "
         PicturePosition =   327683
         Size            =   "2514;635"
         Picture         =   "AdvanceSalary.frx":0112
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
         Left            =   7185
         TabIndex        =   2
         Top             =   5670
         Visible         =   0   'False
         Width           =   1425
         ForeColor       =   0
         Caption         =   "Print         "
         PicturePosition =   327683
         Size            =   "2514;635"
         Picture         =   "AdvanceSalary.frx":0224
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
            Picture         =   "AdvanceSalary.frx":0336
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AdvanceSalary.frx":0788
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin MSForms.ComboBox cmbMonth 
      Height          =   300
      Left            =   6495
      TabIndex        =   10
      Top             =   420
      Width           =   1320
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2328;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbYear 
      Height          =   300
      Left            =   5340
      TabIndex        =   9
      Top             =   420
      Width           =   1110
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "1958;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   885
      TabIndex        =   5
      Top             =   165
      Width           =   6915
      BackColor       =   11517387
      Caption         =   "  Department                         Employee                              Year                  Month"
      Size            =   "12197;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   7920
      TabIndex        =   4
      Top             =   375
      Width           =   1500
      Caption         =   "Refresh        "
      PicturePosition =   327683
      Size            =   "2646;635"
      Picture         =   "AdvanceSalary.frx":089A
      Accelerator     =   82
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "AdvanceSalary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdNew_Click()
    'NewAdvSalary.Show
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
    Screen.MousePointer = vbHourglass
    Dim Cond As String
    
    If cmbShift.ListIndex = 0 Then
        'Cond = " {VAttendance.Shift} in('First','Second') "
        Cond = " {VAttendance.Shift}='First' Or {VAttendance.Shift}='Second'"
    Else
        Cond = " {VAttendance.Shift}='" & cmbShift & "' "
    End If
    
    If cmbEmp.ID = "0" Then
        If cmbDept.ID = "0" Then
            Cond = Cond & " AND ({VAttendance.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & "))"
        Else
            Cond = Cond & " AND ({VAttendance.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & ")) and {VAttendance.DeptID}='" & cmbDept.ID & "'"
        End If
    Else
        Cond = Cond & " and {VAttendance.empid}='" & cmbEmp.ID & "' AND ({VAttendance.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & "))"
    End If
    
With cr1
    .ReportFileName = App.Path & "\AttSheet.rpt"
    .DataFiles(0) = App.Path & "\Payroll.mdb"
    .SelectionFormula = Cond
    '.SelectionFormula = "{VAttendance.VendID}='" & cmbCustID & "' AND ({VOrdItems.Dated}=date(" & Format(DtFrom.Tag, "yyyy,MM,dd") & ") TO date(" & Format(DtTo.Tag, "yyyy,MM,dd") & "))"
    '.Formulas(0) = "CompanyName='" & getCompanyName & "'"
    .Formulas(0) = "FromTo='From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
    '.Formulas(2) = "Ledger='Ledger of " & cmbCustName & "'"
    '.Formulas(2) = "SubReport=" & LVItems.Visible
    'if Cheques Are Present then Show Cheque Detail Also
    'If LVItems.Visible Then
    '    .SubreportToChange = "ChqLedger"
    '    .DataFiles(0) = DBName
    '    .SelectionFormula = "{VChqLedger.AccNo}='" & cmbCustID & "' AND ({VChqLedger.CDate}=date(" & Format(DtFrom.Tag, "yyyy,MM,dd") & ") TO date(" & Format(DtTo.Tag, "yyyy,MM,dd") & "))"
    'End If
    .Action = 1
    .PageZoomNext
End With

Screen.MousePointer = vbDefault
Exit Sub
err:
    MsgBox err.Description
    Screen.MousePointer = vbDefault
End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLedger
    'Call Form_Resize
End Sub

Private Sub Form_Load()

    DtTo = Date
    DtFrom = DateAdd("M", -1, Date)
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    
    
    cmbMonth.AddItem "Jan"
    cmbMonth.List(0, 1) = 1
    cmbMonth.AddItem "Feb"
    cmbMonth.List(1, 1) = 2
    cmbMonth.AddItem "Mar"
    cmbMonth.List(2, 1) = 3
    cmbMonth.AddItem "Apr"
    cmbMonth.List(3, 1) = 4
    cmbMonth.AddItem "May"
    cmbMonth.List(4, 1) = 5
    cmbMonth.AddItem "Jun"
    cmbMonth.List(5, 1) = 6
    cmbMonth.AddItem "Jul"
    cmbMonth.List(6, 1) = 7
    cmbMonth.AddItem "Aug"
    cmbMonth.List(7, 1) = 8
    cmbMonth.AddItem "Sep"
    cmbMonth.List(8, 1) = 9
    cmbMonth.AddItem "Oct"
    cmbMonth.List(9, 1) = "10"
    cmbMonth.AddItem "Nov"
    cmbMonth.List(10, 1) = 11
    cmbMonth.AddItem "Dec"
    cmbMonth.List(11, 1) = 12
    
    Dim CYear As Long
    CYear = Format(Date, "yyyy")
    For i = 2002 To CYear
        cmbYear.AddItem i
    Next i
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals Con, "Name", "Departments", "DeptID"
    cmbDept.Text = "<All Departments>"
    cmbEmp.Text = "<All Employees>"
    cmbMonth.ListIndex = Val(Format$(Date, "MM")) - 1
    cmbYear = Format$(Date, "yyyy")
    'Set LV.SmallIcons = ImageList1
    Call cmdRefresh_Click
    
End Sub
Private Sub cmbdept_matched()

    cmbEmp.ClearVals
    cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.Text = "<All Departments>" Then
        cmbEmp.AddVals Con, "Name", "Employees", "EmpID"
    Else
        cmbEmp.AddVals Con, "Name", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "'"
    End If
    
    
End Sub

Private Sub Form_Resize()

On Error Resume Next
Dim TheHeight As Long


' FAcc.Move (Me.ScaleWidth - FAcc.Width) / 2
 FLV.Move (Me.ScaleWidth - FLV.Width) / 2
'cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 100
'cmdPrint.Move cmdClose.Left - cmdPrint.Width - 50, cmdClose.Top

'FDrag.Move 50, Me.ScaleHeight * 3 / 5, Me.ScaleWidth - 100
 
 'If LVItems.Visible Then
 '  TheHeight = FDrag.Top
 '  FDrag.Visible = True
 'Else
 '  TheHeight = Me.ScaleHeight - cmdClose.Height - 200
 '  FDrag.Visible = False
 'End If
 
 'LV.Move 50, FAcc.Top + FAcc.Height, Me.ScaleWidth - 100, TheHeight - (FAcc.Top + FAcc.Height)
 'LVItems.Move 50, FDrag.Top + FDrag.Height, Me.ScaleWidth - 100, Me.ScaleHeight - (FAcc.Height + cmdClose.Height + LV.Height + 250)
 'FDrag.ZOrder 1
 
End Sub


Private Sub RefreshLedger()


    If cmbDept.MatchFound = False Or cmbEmp.MatchFound = False Then Exit Sub
    If cmbYear.MatchFound = False Or cmbMonth.MatchFound = False Then Exit Sub
    Dim rsLedger As New ADODB.Recordset
    Dim itm As ListItem
    Dim Cond As String
    Dim Sql As String
    Dim cmd As New ADODB.Command
    
    
    
    
    If cmbEmp.ID = "0" Then
        If cmbDept.ID = "0" Then
            Cond = " Where "
        Else
            Cond = " Where DeptID='" & cmbDept.ID & "' and"
        End If
    Else
        Cond = " Where empid='" & cmbEmp.ID & "' and"
    End If
    Dim rsCheck As New ADODB.Recordset
    Dim TDate As Date
    
    TDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), Day(Date))
    
    Sql = "select * from VAdvanceSalary " & Cond & " JoinDate<=#" & Format(TDate, "dd-MMM-yyyy") & "# Order By DeptID,EmpID"
    cmd.CommandText = Sql
    Set cmd.ActiveConnection = Con
    
    cmd.Parameters.Refresh
    
    cmd.Parameters(0).Type = adVarChar
    cmd.Parameters(0).value = cmbMonth.Text
    cmd.Parameters(1).Type = adVarChar
    cmd.Parameters(1).value = cmbYear.Text
    rsLedger.CursorType = adOpenForwardOnly
    rsLedger.LockType = adLockReadOnly
    Set rsLedger = cmd.Execute
    With rsLedger
        
        '.Open Sql, Con, adOpenForwardOnly, adLockReadOnly
        Dim TotalSuns As Integer
        Dim SalDays As Double
        Dim Rate As Double
        Dim AbsAllowed As Integer
        Dim AbsentDays As Double
        Dim TotalSalDays As Double
        
        Dim Total As Double
        Dim OverTimeHours As Double
        Dim LateHours As Double
        Dim TaxPer As Double
        Dim DedAmt As Double
        Dim Payed As Double
        Dim TaxSal As Double
        Dim IsComputed As Boolean
        Dim LastDay As Integer
        Dim NTime As Integer
        Dim LTime As Double
        Dim DT As String
        Dim IsPayed As Boolean
        Dim Holidays As Integer
        Dim rs As New ADODB.Recordset
        
        TotalSalDays = DateAdd("m", 1, TDate) - TDate
        LastDay = TotalSalDays
        
        If TDate = Date Then
            TotalSalDays = Day(Date)
        ElseIf Date < TDate Then
            MsgBox "Can't Show Entries Of Future Months.", vbInformation
            Exit Sub
        End If
        
        With rs
            .Open "Select count(DT) from Holidays Where Format(DT,'MMM-yyyy')='" & cmbMonth.Text & "-" & cmbYear & "' and day(DT)<=" & TotalSalDays, Con, adOpenForwardOnly, adLockReadOnly
            If .RecordCount > 0 Then
                Holidays = .Fields(0) & ""
            End If
            .Close
        End With
        Set rs = Nothing
        
        
        TotalSuns = TotalSundays(cmbMonth.List(cmbMonth.ListIndex, 1), cmbYear, TotalSalDays)
        Holidays = TotalSuns + Holidays
        
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
                    
            Call Computed(![EmpID], IsComputed, TaxPer, AbsAllowed, OverTimeHours, NTime, LTime, DT, IsPayed)
            
            If Not IsComputed Then
                'Tax %
                If IsNull(![PerDeduct]) Then TaxPer = 0 Else TaxPer = Val(![PerDeduct]) & ""
                'Absent Days....
                AbsAllowed = ![AbsAllowed]
                'OverTime Hours
                If IsNull(![OverTimeHours]) Then OverTimeHours = 0 Else OverTimeHours = ![OverTimeHours]
                'Late Hours
                NTime = ![NTimes]
                LTime = ![TimeAllowed]
                
            End If
            
            If IsNull(![AbsentDays]) Then AbsentDays = 0 Else AbsentDays = Val(![AbsentDays]) / 2
            AbsentDays = IIf(AbsentDays > 0, AbsentDays - AbsAllowed, AbsentDays)
            Dim theKey As String
            If IsComputed Then
                theKey = ![EmpID] & DT
            Else
                theKey = ![EmpID] & ""
            End If
            Set itm = LV.ListItems.Add(, theKey, ![EmpID] & "")
            If IsComputed Then
                itm.Checked = IsPayed
            End If
            itm.SubItems(1) = ![Name] & ""
            itm.SubItems(2) = ![StartingSalary] & ""
            Rate = Round(Val(![StartingSalary]) / LastDay)
            itm.SubItems(3) = Rate
            
            
            itm.SubItems(4) = AbsentDays
            itm.SubItems(5) = AbsentDays * Rate
            
            
            If IsNull(![PresentDays]) Then
                SalDays = 0
            Else
                SalDays = (Val(![PresentDays]) / 2) + Holidays
            End If
            If SalDays = 0 Then SalDays = 0 Else SalDays = (SalDays - AbsentDays)
            itm.SubItems(6) = SalDays
            itm.SubItems(7) = Round(SalDays * Rate)
            
            
            
            itm.SubItems(8) = OverTimeHours
            itm.SubItems(9) = (OverTimeHours * ((Rate / 8) * Val(![OverTime])))
            
            LateHours = GetLateHours(![EmpID], NTime, LTime)
            
            itm.SubItems(10) = LateHours
            itm.SubItems(11) = Round((LateHours * (Rate / 8)))
            
            Total = Round(Val(itm.SubItems(7)) + Val(itm.SubItems(9)) - Val(itm.SubItems(11)))
            itm.SubItems(12) = Total
            'TaxPer = Val(![PerDeduct])
            itm.SubItems(13) = TaxPer
            'DedAmt = Round((Val(![Salary]) / TotalSalDays) * (TaxPer / 100))
            TaxSal = (Val(![Salary]) / LastDay) * TotalSalDays
            DedAmt = Round(TaxSal * (TaxPer / 100))
            itm.SubItems(14) = DedAmt
            itm.SubItems(15) = Total - DedAmt
            If IsNull(![Payed]) Then Payed = 0 Else Payed = ![Payed]
            itm.SubItems(16) = Payed
            'itm.SubItems(17) = Val(itm.SubItems(15)) - Payed
            Dim LastDate As Date
            LastDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), LastDay)
            'Save Tax,Settings for Selected Month...
            'If Not (IsComputed) And (Date > LastDate) Then Call SaveIt(rsLedger, TaxPer)
            
            
            .MoveNext
        Next
        .Close
    End With
    cmbMonth.Tag = cmbMonth.Text
    cmbYear.Tag = cmbYear.Text
    Set rsLedger = Nothing
    Sql = ""
    Cond = ""
    ICO = ""
End Sub
Private Sub OldRefreshLedger()
    'New Is Copied From SalarySheet
    If cmbDept.MatchFound = False Or cmbEmp.MatchFound = False Then Exit Sub
    If cmbYear.MatchFound = False Or cmbMonth.MatchFound = False Then Exit Sub
    Dim rsLedger As New ADODB.Recordset
    Dim itm As ListItem
    Dim Cond As String
    Dim Sql As String
    Dim cmd As New ADODB.Command
    
    If cmbEmp.ID = "0" Then
        If cmbDept.ID = "0" Then
            Cond = " "
        Else
            Cond = " Where DeptID='" & cmbDept.ID & "'"
        End If
    Else
        Cond = " Where empid='" & cmbEmp.ID & "'"
    End If
    Dim rsCheck As New ADODB.Recordset
    
    Sql = "select * from VAdvanceSalary " & Cond & " Order By DeptID,EmpID"
    cmd.CommandText = Sql
    Set cmd.ActiveConnection = Con
    
    cmd.Parameters.Refresh
    
    cmd.Parameters(0).Type = adVarChar
    cmd.Parameters(0).value = cmbMonth.Text
    cmd.Parameters(1).Type = adVarChar
    cmd.Parameters(1).value = cmbYear.Text
    rsLedger.CursorType = adOpenForwardOnly
    rsLedger.LockType = adLockReadOnly
    Set rsLedger = cmd.Execute
    With rsLedger
        
        '.Open Sql, Con, adOpenForwardOnly, adLockReadOnly
        Dim SalDayCheck As Integer
        Dim Rate As Double
        Dim AbsAllowed As Integer
        Dim AbsentDays As Long
        Dim TotalSalDays As Long
        Dim TDate As Date
        Dim Total As Double
        Dim OverTimeHours As Double
        Dim LateHours As Double
        Dim TaxPer As Double
        Dim DedAmt As Double
        Dim Payed As Double
        Dim TaxSal As Double
        Dim IsComputed As Boolean
        Dim LastDay As Integer
        Dim NTime As Integer
        Dim LTime As Double
        
        TDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), Day(Date))
        TotalSalDays = DateAdd("m", 1, TDate) - TDate
        LastDay = TotalSalDays
        
        If TDate = Date Then
            TotalSalDays = Day(Date)
        ElseIf Date < TDate Then
            MsgBox "Can't Show Entries Of Future Months.", vbInformation
            Exit Sub
        End If
        
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
             'This is to avoid Exe Error...
            'Call Computed(![EmpID], IsComputed, TaxPer, AbsAllowed, OverTimeHours, NTime, LTime)
            
            If Not IsComputed Then
                'Tax %
                If IsNull(![PerDeduct]) Then TaxPer = 0 Else TaxPer = Val(![PerDeduct]) & ""
                'Absent Days....
                AbsAllowed = ![AbsAllowed]
                'OverTime Hours
                If IsNull(![OverTimeHours]) Then OverTimeHours = 0 Else OverTimeHours = ![OverTimeHours]
                'Late Hours
                NTime = ![NTimes]
                LTime = ![TimeAllowed]
                
            End If
            
            If IsNull(![AbsentDays]) Then AbsentDays = 0 Else AbsentDays = Val(![AbsentDays]) / 2
            AbsentDays = IIf(AbsentDays > 0, AbsentDays - AbsAllowed, AbsentDays)
            
            Set itm = LV.ListItems.Add(, ![EmpID] & "", ![EmpID] & "")
            itm.SubItems(1) = ![Name] & ""
            itm.SubItems(2) = ![StartingSalary] & ""
            Rate = Round(Val(![StartingSalary]) / 30)
            itm.SubItems(3) = Rate
            
            
            itm.SubItems(4) = AbsentDays
            itm.SubItems(5) = AbsentDays * Rate
            
            If IsNull(![PresentDays]) Then SalDayCheck = 0 Else SalDayCheck = 1
            
            itm.SubItems(6) = (TotalSalDays - AbsentDays) * SalDayCheck
            itm.SubItems(7) = (TotalSalDays * Rate) * SalDayCheck
            
            
            
            itm.SubItems(8) = OverTimeHours
            itm.SubItems(9) = (OverTimeHours * (Rate * Val(![OverTime]))) * SalDayCheck
            
            LateHours = GetLateHours(![EmpID], NTime, LTime)
            
            itm.SubItems(10) = LateHours
            itm.SubItems(11) = Round((LateHours * Rate)) * SalDayCheck
            
            Total = Round(Val(itm.SubItems(7)) + Val(itm.SubItems(9)) - Val(itm.SubItems(11)))
            itm.SubItems(12) = Total
            'TaxPer = Val(![PerDeduct])
            itm.SubItems(13) = TaxPer
            'DedAmt = Round((Val(![Salary]) / TotalSalDays) * (TaxPer / 100))
            TaxSal = (Val(![Salary]) / 30) * TotalSalDays
            DedAmt = Round(TaxSal * (TaxPer / 100))
            itm.SubItems(14) = DedAmt
            itm.SubItems(15) = Total - DedAmt
            If IsNull(![Payed]) Then Payed = 0 Else Payed = ![Payed]
            itm.SubItems(16) = Payed
            'itm.SubItems(17) = Val(itm.SubItems(15)) - Payed
            Dim LastDate As Date
            LastDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), LastDay)
            'Save Tax,Settings for Selected Month...
            'If Not (IsComputed) And (Date > LastDate) Then Call SaveIt(rsLedger, TaxPer)
            .MoveNext
        Next
        .Close
    End With
    Set rsLedger = Nothing
    Sql = ""
    Cond = ""
    ICO = ""
End Sub
Private Function GetLateHours(EmpID As String, NTime As Integer, LTime As Double) As Double

    Dim LHours As Double
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
    With cmd
        .CommandText = "Select LHrs from VLHours Where EmpID='" & EmpID & "'"
        Set .ActiveConnection = Con
        .Parameters.Refresh
        .Parameters(0).Type = adDouble
        .Parameters(0).value = LTime
        .Parameters(1).Type = adVarChar
        .Parameters(1).value = cmbMonth
        .Parameters(2).Type = adVarChar
        .Parameters(2).value = cmbYear
        .Parameters(3).Type = adDouble
        .Parameters(3).value = NTime
        
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        Set rs = .Execute
        
    End With
    
    If rs.RecordCount > 0 Then
        If IsNull(rs.Fields(0)) Then LHours = 0 Else LHours = rs.Fields(0)
    Else
        LHours = 0
    End If
    GetLateHours = LHours
    
End Function
Private Sub Computed(Emp As String, ByRef IsComputed As Boolean, ByRef TaxPer As Double, ByRef AbsentDays As Integer, ByRef OverTimeHours As Double, ByRef NTime As Integer, ByRef LTime As Double, ByRef DT As String, ByRef Payed As Boolean)

    Dim rsCheck As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    With cmd
        .CommandText = "Select * from VSalarySettings where empid='" & Emp & "'"
        Set .ActiveConnection = Con
        .Parameters.Refresh
        .Parameters(0).Type = adVarChar
        .Parameters(0).value = cmbMonth.Text
        .Parameters(1).Type = adVarChar
        .Parameters(1).value = cmbYear.Text
        rsCheck.CursorType = adOpenForwardOnly
        rsCheck.LockType = adLockReadOnly
        Set rsCheck = .Execute()
    End With
    
    With rsCheck
        
        If .RecordCount > 0 Then
            IsComputed = True
            TaxPer = ![PerDeduct]
            AbsentDays = ![absentallowed]
            OverTimeHours = ![OverTime]
            NTime = ![NTimes]
            LTime = ![TimeAllowed]
            DT = Format(![DT], "dd-MMM-yyyy") & ""
            ', AbsentDays, OverTimeHours, NTime, LTime)
            LV.Checkboxes = True
            'cmdUpdate.Visible = True
            Payed = ![Payed]
        Else
            IsComputed = False
            LV.Checkboxes = False
            'cmdUpdate.Visible = False
        End If
    End With
    
End Sub

Private Sub OldComputed(Emp As String, ByRef IsComputed As Boolean, ByRef TaxPer As Double, ByRef AbsentDays As Integer, ByRef OverTimeHours As Double, ByRef NTime As Integer, LTime As Double)

    Dim rsCheck As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    With cmd
        .CommandText = "Select * from VSalarySettings where empid='" & Emp & "'"
        Set .ActiveConnection = Con
        .Parameters.Refresh
        .Parameters(0).Type = adVarChar
        .Parameters(0).value = cmbMonth.Text
        .Parameters(1).Type = adVarChar
        .Parameters(1).value = cmbYear.Text
        rsCheck.CursorType = adOpenForwardOnly
        rsCheck.LockType = adLockReadOnly
        Set rsCheck = .Execute()
    End With
    
    With rsCheck
        
        If .RecordCount > 0 Then
            IsComputed = True
            TaxPer = ![PerDeduct]
            AbsentDays = ![absentallowed]
            OverTimeHours = ![OverTime]
            NTime = ![NTimes]
            LTime = ![TimeAllowed]
            ', AbsentDays, OverTimeHours, NTime, LTime)
        Else
            IsComputed = False
        End If
    End With
    
End Sub

