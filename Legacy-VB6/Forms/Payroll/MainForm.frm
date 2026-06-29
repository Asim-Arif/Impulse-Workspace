VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.MDIForm MainForm 
   BackColor       =   &H8000000C&
   Caption         =   "Payroll Package"
   ClientHeight    =   7035
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9990
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2550
      Top             =   2760
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4215
      Top             =   1770
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":030A
            Key             =   "Cont"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":080C
            Key             =   "AdvancesLedger"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":0D0E
            Key             =   "OverTime"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":1150
            Key             =   "SalSheet"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":1732
            Key             =   "Emp"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":1C34
            Key             =   "NewShort"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":2136
            Key             =   "NewLong"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":2638
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":2B3A
            Key             =   "EmpLedger"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":2F8C
            Key             =   "ContLedger"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":348E
            Key             =   "Settings"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":38E0
            Key             =   "TakeAtt"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   1005
      ButtonWidth     =   1826
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Employees"
            Key             =   "Emp"
            Object.ToolTipText     =   "Employees List"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Departments"
            Key             =   "Dept"
            Object.ToolTipText     =   "Departments List"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Attendence"
            Key             =   "TakeAtt"
            Object.ToolTipText     =   "Take Attendance"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Over Time"
            Key             =   "OverTime"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "S.T. Loan"
            Key             =   "NewShort"
            Object.ToolTipText     =   "New Short Term Loan"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "L.T. Loan"
            Key             =   "NewLong"
            Object.ToolTipText     =   "New Long Term Loan"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Advances"
            Key             =   "AdvancesLedger"
            Object.ToolTipText     =   "Advances Ledger"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Emp Ledger"
            Key             =   "EmpLedger"
            Object.ToolTipText     =   "Employee Ledger"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salary Sheet"
            Key             =   "SalSheet"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Contractors"
            Key             =   "Cont"
            Object.ToolTipText     =   "Contractors List"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cont Ledger"
            Key             =   "ContLedger"
            Object.ToolTipText     =   "Contractor Ledger"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3705
      Top             =   1785
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Please Select Payroll Database to Connect."
      Filter          =   """Payroll Database|Payroll.mdb"""
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuUserManager 
         Caption         =   "User Manager"
      End
      Begin VB.Menu mnuCalculater 
         Caption         =   "Use Calculater"
      End
      Begin VB.Menu mnuReminders 
         Caption         =   "Reminders"
      End
      Begin VB.Menu dashcalc 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangePass 
         Caption         =   "Change Password"
      End
      Begin VB.Menu mnuLogOff 
         Caption         =   "Log Off..."
      End
      Begin VB.Menu DashLogOff 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnugeneral 
      Caption         =   "Settings"
      Begin VB.Menu mnuChange 
         Caption         =   "Depatemental Settings"
      End
      Begin VB.Menu mnuHolidays 
         Caption         =   "Define Holidays"
      End
      Begin VB.Menu mnugendash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTaxRanges 
         Caption         =   "Tax Ranges"
      End
   End
   Begin VB.Menu mnuEmp 
      Caption         =   "Employees"
      Begin VB.Menu mnuDeptList 
         Caption         =   "&Department List"
      End
      Begin VB.Menu mnuNewDept 
         Caption         =   "&Add New Department"
      End
      Begin VB.Menu empdash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmpList 
         Caption         =   "Employee &List"
      End
      Begin VB.Menu mnunewEmp 
         Caption         =   "Add &New Employee"
      End
      Begin VB.Menu dashempsal 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmpLedger 
         Caption         =   "Employee Ledger"
      End
      Begin VB.Menu ln4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalaryLedger 
         Caption         =   "Employee &Salary Ledger"
      End
      Begin VB.Menu mnuSalHistory 
         Caption         =   "Employees Salary History"
      End
   End
   Begin VB.Menu mnuAttSheet 
      Caption         =   "Attendance"
      Begin VB.Menu mnutake 
         Caption         =   "Take Attendance"
      End
      Begin VB.Menu mnuSheet 
         Caption         =   "Attendance Sheet"
      End
      Begin VB.Menu mnuAttdash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOverTime 
         Caption         =   "Over Time Record"
      End
      Begin VB.Menu dashpenalty 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPenaltyDays 
         Caption         =   "Employee Penalty Days"
      End
   End
   Begin VB.Menu mnuAdvHistory 
      Caption         =   "Advances"
      Begin VB.Menu mnuShort 
         Caption         =   "New &Short Term Advance"
      End
      Begin VB.Menu mnuLong 
         Caption         =   "New &Long Term Loan"
      End
      Begin VB.Menu advdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShortAdv 
         Caption         =   "Short Term Advance Ledger"
      End
      Begin VB.Menu mnuLTLedger 
         Caption         =   "Long Term Loan Ledger"
      End
   End
   Begin VB.Menu mnuSalary 
      Caption         =   "Salary"
      Begin VB.Menu mnuAdvSalary 
         Caption         =   "Advance Salary"
      End
      Begin VB.Menu mnuAttdash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalSheet 
         Caption         =   "Salary Sheet"
      End
      Begin VB.Menu mnuExtSalarySheet 
         Caption         =   "External Salary Sheet"
      End
      Begin VB.Menu mnuPaySlip 
         Caption         =   "Payslip"
      End
      Begin VB.Menu dashSalary 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOldAge 
         Caption         =   "Old Age"
      End
      Begin VB.Menu mnuSecurity 
         Caption         =   "Social Security"
      End
      Begin VB.Menu dashbonus 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBonusSheet 
         Caption         =   "Bonus Sheet"
      End
   End
   Begin VB.Menu mnucont 
      Caption         =   "Contractor's Info"
      Begin VB.Menu mnuContDept 
         Caption         =   "Contractor's Departments"
      End
      Begin VB.Menu cont4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuItems 
         Caption         =   "Items Info"
      End
      Begin VB.Menu cont1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContInfo 
         Caption         =   "Contractor's Info"
      End
      Begin VB.Menu mnuAss 
         Caption         =   "Assigned Items"
      End
      Begin VB.Menu cont2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuloan 
         Caption         =   "Loan Deduction"
         Visible         =   0   'False
      End
      Begin VB.Menu cont3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuContLedger 
         Caption         =   "Contractor's Ledger"
      End
      Begin VB.Menu contdash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewEntry 
         Caption         =   "&New Entry"
      End
      Begin VB.Menu mnuPayLedger 
         Caption         =   "Payment Ledger"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPayEntry 
         Caption         =   "New Payment Entry"
      End
      Begin VB.Menu otherdash1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Begin VB.Menu mnuApplicationForm 
         Caption         =   "Application Form"
      End
      Begin VB.Menu mnurptemplist 
         Caption         =   "Employee List "
         Begin VB.Menu mnurotEmplistInt 
            Caption         =   "With Internal Salary"
         End
         Begin VB.Menu mnurotEmplistExt 
            Caption         =   "With External Salary"
         End
      End
      Begin VB.Menu mnurptEmpCards 
         Caption         =   "Employee Cards"
      End
      Begin VB.Menu mnuAbsentSheet 
         Caption         =   "Absent Sheet"
      End
      Begin VB.Menu mnuPrintJoinLeave 
         Caption         =   "Employees Joined and Discontinued"
      End
      Begin VB.Menu dashRpts1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAttRegister 
         Caption         =   "Attendance Register"
      End
      Begin VB.Menu mnuExtSalsheet 
         Caption         =   "External Salary Sheet"
      End
      Begin VB.Menu mnuSalSummary 
         Caption         =   "Monthly Salary Sheet Summary"
      End
      Begin VB.Menu dashsal 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalaryComparison 
         Caption         =   "Monthly Salary Comparison"
      End
      Begin VB.Menu mnuOutStanding 
         Caption         =   "Out Standing Loans Report"
      End
   End
   Begin VB.Menu mnupopemp 
      Caption         =   "PopMenuEmp"
      Begin VB.Menu mnuempRefresh 
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
      End
      Begin VB.Menu mnuMake 
         Caption         =   "Make This Employee Inactive"
      End
      Begin VB.Menu dd1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuempNewDept 
         Caption         =   "Add New &Department"
      End
      Begin VB.Menu mnuEmpNewEmp 
         Caption         =   "Add New &Employee"
      End
      Begin VB.Menu mnuempdash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmpEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuEmpdash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAssItems 
         Caption         =   "Show &Assigned Items"
      End
      Begin VB.Menu mnuEmpPrint 
         Caption         =   "&Print Employee List (Internal Salary)"
      End
      Begin VB.Menu mnuPrintEmpExt 
         Caption         =   "Print Employee List (E&xternal Salary)"
      End
      Begin VB.Menu mnuPrintEmpTrans 
         Caption         =   "Print Employee List (Company Transport)"
      End
      Begin VB.Menu dashemps 
         Caption         =   "-"
      End
      Begin VB.Menu mnuappForm 
         Caption         =   "Print Application Form"
      End
      Begin VB.Menu mnuPrintCard 
         Caption         =   "Print Employee Card"
         Begin VB.Menu mnuPrintEmpCard 
            Caption         =   "This Employee Only"
         End
         Begin VB.Menu mnuPrintAllCard 
            Caption         =   "For All Employees"
         End
      End
   End
   Begin VB.Menu mnupopdept 
      Caption         =   "popMenuDept"
      Begin VB.Menu mnuDeptRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu popdeptdash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowDept 
         Caption         =   "Show Inactive Departments"
      End
      Begin VB.Menu mnuDeptdash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeptInact 
         Caption         =   "Make This Department Inactive"
      End
      Begin VB.Menu mnuDeptdash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeptnew 
         Caption         =   "Add &New Department"
      End
      Begin VB.Menu mnuDepdash4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeptEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuDeptdash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeptPrint 
         Caption         =   "&Print Department List"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()
    With Toolbar1
        Set Toolbar1.ImageList = ImageList1
        For i = 1 To .Buttons.Count
            If Not .Buttons(i).Style = tbrSeparator Then
                .Buttons(i).Image = ImageList1.ListImages(.Buttons(i).Key).Key
            End If
        Next i
    End With
    
    mnupopemp.Visible = False
    mnupopdept.Visible = False

    mnucont.Visible = ContractorsEnabled
    Toolbar1.Buttons("Cont").Visible = ContractorsEnabled
    Toolbar1.Buttons("ContLedger").Visible = ContractorsEnabled
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    On Error Resume Next
    If Not LoggingOff Then Set Con = Nothing
End Sub

Private Sub mnuAbsentSheet_Click()
AbsentSheet.Show
End Sub

Public Sub mnuAdvSalary_Click()
    AdvSalary.Show
End Sub

Private Sub mnuappForm_Click()
    EmployeeList.PrintAppForm
End Sub

Private Sub mnuApplicationForm_Click()

If Not IsLoaded("EmployeeList") Then
    EmployeeList.Visible = False
End If

EmployeeList.PrintAppForm

End Sub

Private Sub mnuAss_Click()
    AssItems.Show
End Sub
Private Sub mnuAssItems_Click()
    Call ContractorList.showItems
End Sub

Private Sub mnuAttRegister_Click()
PrintAttReg.Show 1
End Sub

Public Sub mnuBonusSheet_Click()
BonusSheet.Show
End Sub

Private Sub mnuCalculater_Click()
   On Error GoTo ERR
   Call Shell("calc.exe")
   Exit Sub
ERR:
   MsgBox "Cannot Run Calculater Due To Error :" & vbNewLine & ERR.Description, vbCritical
End Sub

Private Sub mnuChange_Click()

    Dim f As New frmSettings
    Load f
    'f.lblEmpID = "GENERAL"
    'f.Editrec
    f.Show
    
End Sub

Private Sub mnuChangePass_Click()
 With frmUserInfo
   .Show_Option = Opt_ChangeUserPassword
   .Tag = CurrentUserName
   .Show 1
 End With
End Sub

Private Sub mnuContDept_Click()
    ContDeptList.Show
End Sub
Private Sub mnuContInfo_Click()
    ContractorList.Show
End Sub
Private Sub mnuContLedger_Click()
    ContLedger.Show
End Sub
Private Sub mnuDeptEdit_Click()
    If mnupopdept.Tag = "From Dept" Then
        Call DeptList.mnuEdit(mnuDeptEdit.Tag)
    Else
        Call ContDeptList.mnuEdit(mnuDeptEdit.Tag)
    End If
End Sub
Private Sub mnuDeptInact_Click()
    If mnupopdept.Tag = "From Dept" Then
        Call DeptList.MakeInactive
    Else
        Call ContDeptList.MakeInactive
    End If
End Sub
Private Sub mnuDeptList_Click()
    DeptList.Show
End Sub

Private Sub mnuDeptnew_Click()
    If mnupopdept.Tag = "From Dept" Then
        Call DeptList.mnuAddDept
    Else
        Call ContDeptList.mnuAddDept
    End If
End Sub

Private Sub mnuDeptPrint_Click()
Call DeptList.PrintDeptList
End Sub

Private Sub mnuDeptRefresh_Click()
    If mnupopdept.Tag = "From Dept" Then
        Call DeptList.RefreshIt
    Else
        Call ContDeptList.RefreshIt
    End If
End Sub

Private Sub mnuEmpEdit_Click()
    If mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.mnuEdit(mnuEmpEdit.Tag)
    Else
        Call ContractorList.mnuEdit(mnuEmpEdit.Tag)
    End If
End Sub

Private Sub mnuEmpLedger_Click()
    EmpLedger.Show
End Sub

Private Sub mnuEmpList_Click()
    EmployeeList.Show
End Sub
Private Sub mnuempNewDept_Click()
    If Me.mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.mnuAddDept
    Else
        Call ContractorList.mnuAddDept
    End If
End Sub
Private Sub mnuEmpNewEmp_Click()
    If Me.mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.mnuAdd
    Else
        Call ContractorList.mnuAdd
    End If
End Sub

Private Sub mnuLateHours_Click()
    Call Attend.enterLateHours
End Sub

Private Sub mnuEmpPrint_Click()
Call EmployeeList.PrintEmpList
End Sub

Private Sub mnuempRefresh_Click()
    If mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.RefreshIt
    Else
        Call ContractorList.RefreshIt
    End If
End Sub
Private Sub mnuExit_Click()
    If Con.State = 1 Then Con.Close
    Set Con = Nothing
    End
End Sub

Private Sub mnuExtSalarySheet_Click()

SalarySheetExt.Show

End Sub

Private Sub mnuExtSalsheet_Click()
SalarySheetExt.Show

End Sub

Private Sub mnuHolidays_Click()
    Load Holidays
    'Holidays.Move 3000, 500
    With Holidays
        .Move (Me.Width - .Width) / 2, Me.Top + 500
        .Show
    End With
End Sub


'Private Sub mnuLongADv_Click()
'    AdvanceLong.Show
'End Sub
Private Sub mnuItems_Click()
    frmItems.Show
End Sub

Private Sub mnuLogOff_Click()
If MsgBox("Are You Sure To Log Off ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
   LoggingOff = True
   Unload MainForm
   Call Main
End Sub

Public Sub mnuLong_Click()
      With NewAdvanceShort
      .LongTerm = True
      .Show 1
   End With
End Sub

Private Sub mnuLTLedger_Click()
LoanLedger.Show
End Sub

Private Sub mnuMake_Click()
    If mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.MakeInactive
    Else
        Call ContractorList.MakeInactive
    End If
    
End Sub

Private Sub mnuNewDept_Click()
    Dim f As New NewDepartment
    f.TableName = "Departments"
    f.lblCatID = GetNextDeptID("Departments")
    f.Add = True
    f.Show 1
End Sub

Private Sub mnunewEmp_Click()
    Dim f As New NewEmployee
    f.TableName = "Employees"
    f.Add = True
    'f.Move (Me.Width - f.Width) / 2, Me.Top + 1000
    f.Show 1
End Sub

Private Sub mnuNewEntry_Click()
    NewContEntry.Show 1
End Sub

Private Sub mnuOldAge_Click()
    OldAge.Show
End Sub

Private Sub mnuOutStanding_Click()
PrintOutStanding.Show 1
End Sub

Private Sub mnuOverTime_Click()
    OverTime.Show
End Sub

Private Sub mnuPayEntry_Click()
    NewPayEntry.Show 1
End Sub

Private Sub mnuPayLedger_Click()
    PaymentLedger.Show
End Sub

Private Sub mnuPaySlip_Click()
    Load Payslip
    'Holidays.Move 3000, 500
    With Payslip
        .Move (Me.Width - .Width) / 2, Me.Top + 500
        .Show
    End With
End Sub

Private Sub mnuPenaltyDays_Click()
    EmpAbsents.Show
End Sub

Private Sub mnuPrintAllCard_Click()
Call EmployeeList.PrintEmpCard
End Sub

Private Sub mnuPrintEmpCard_Click()
If EmployeeList.LV.ListItems.Count = 0 Then Exit Sub
If EmployeeList.LV.SelectedItem.Bold Then Exit Sub
Call EmployeeList.PrintEmpCard(EmployeeList.LV.SelectedItem)
End Sub

Private Sub mnuPrintEmpExt_Click()
Call EmployeeList.PrintEmpList(True)
End Sub

Private Sub mnuPrintEmpTrans_Click()
 Call EmployeeList.PrintEmpList(False, True)
End Sub

Private Sub mnuPrintJoinLeave_Click()
   PrintEmpJoinLeave.Show 1
End Sub

Private Sub mnuReminders_Click()
frmReminders.Show 1
End Sub

Private Sub mnurotEmplistExt_Click()
If Not IsLoaded("EmployeeList") Then
    EmployeeList.Visible = False
End If
EmployeeList.PrintEmpList (True)
End Sub

Private Sub mnurotEmplistInt_Click()
If Not IsLoaded("EmployeeList") Then
    EmployeeList.Visible = False
End If

EmployeeList.PrintEmpList
End Sub

Private Sub mnurptEmpCards_Click()
If Not IsLoaded("EmployeeList") Then
    EmployeeList.Visible = False
End If

EmployeeList.PrintEmpCard
End Sub

Private Sub mnuSalaryComparison_Click()
PrintSalaryComparison.Show 1
End Sub

Private Sub mnuSalaryLedger_Click()
SalaryLedger.Show
End Sub

Private Sub mnuSalHistory_Click()
SalaryHistory.Show

End Sub

Public Sub mnuSalSheet_Click()
    SalarySheet.Show
End Sub

Private Sub mnuSalSummary_Click()
PrintSalarySummary.Show 1
End Sub

Private Sub mnuSecurity_Click()
    SocialSecurity.Show
End Sub

Private Sub mnuSheet_Click()
    AttSheet.Show
End Sub

Public Sub mnuShort_Click()
    'NewAdvShort.Show
    With NewAdvanceShort
      .LongTerm = False
      .Show 1
   End With
End Sub

Private Sub mnuShortAdv_Click()
   AdvanceShort.Show
End Sub

Private Sub mnuShowDept_Click()
    If mnupopdept.Tag = "From Dept" Then
        Call DeptList.ShowInactive(mnuShowDept)
    Else
        Call ContDeptList.ShowInactive(mnuShowDept)
    End If
End Sub

Private Sub mnuShowInactive_Click()
    If mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.ShowInactive(mnuShowInactive)
    Else
        Call ContractorList.ShowInactive(mnuShowInactive)
    End If
End Sub
Private Sub mnutake_Click()
    Load Attend
    'Attend.Move (Me.Width - Attend.Width) / 2, 100
    Attend.Show
End Sub

Private Sub mnuTaxRanges_Click()
    Load TaxRanges
    'Holidays.Move 3000, 500
    With TaxRanges
        .Move (Me.Width - .Width) / 2, Me.Top + 500
        .Show
    End With
End Sub

Private Sub mnuUserManager_Click()
frmUsers.Show
End Sub

Private Sub Timer1_Timer()
Timer1.Interval = 1000
    Dim SeverTime As Date
    ServerTime = Con.Execute("Select GetDate()").Fields(0)
    
    MainForm.Caption = "Payroll Package." & Space$(5) & " SERVER TIME : " & Format(ServerTime, "dddd dd-MMM-yyyy hh:mm:ss AMPM")
    
    If Format(ServerDate, "dd-MMM-yyyy") <> Format(ServerTime, "dd-MMM-yyyy") Then ServerDate = Format(ServerTime, "dd-MMM-yyyy")
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Emp"
            EmployeeList.Show
        Case "Dept"
            DeptList.Show
        Case "TakeAtt"
            Call mnutake_Click
        Case "NewShort"
            Call mnuShort_Click
        Case "NewLong"
            Call mnuLong_Click
        Case "AdvancesLedger"
            Call AdvanceShort.Show
        Case "EmpLedger"
            Call mnuEmpLedger_Click
        Case "Cont"
            Call mnuContInfo_Click
        Case "ContLedger"
            Call ContLedger.Show
        Case "SalSheet"
            Call mnuSalSheet_Click
         Case "OverTime"
            Call mnuOverTime_Click
        Case "Settings"
            Dim f As New frmSettings
            Load f
            'f.lblEmpID = "GENERAL"
            'f.Editrec
            f.Show
    End Select
End Sub
Public Sub ApplySecuritySettings()
      
   mnuAdvSalary.Enabled = UserHasAccess("CalculateAdvSal")
   mnuSalSheet.Enabled = UserHasAccess("CalculateSalSheet")
   
   mnuShort.Enabled = UserHasAccess("ShortTermLoan")
   mnuShortAdv.Enabled = UserHasAccess("ShortTermLoan")
   mnuLong.Enabled = UserHasAccess("LongTermLoan")
   
   
   mnuEmpLedger.Enabled = UserHasAccess("EmployeeLedger")
   mnurptEmpCards.Enabled = UserHasAccess("PrintEmpCards")
   mnuApplicationForm.Enabled = UserHasAccess("ApplicationForm")
   mnurptemplist.Enabled = UserHasAccess("PrintEmpList")
   mnuAttRegister.Enabled = UserHasAccess("AttendanceRegister")
   mnuPaySlip.Enabled = UserHasAccess("EmpPaySlip")
   
   mnutake.Enabled = UserHasAccess("TakeAttendance")
   mnunewEmp.Enabled = UserHasAccess("NewEmployee")
   mnuEmpList.Enabled = UserHasAccess("EmployeeList")
   
   mnuOldAge.Enabled = UserHasAccess("OldAge")
   mnuSecurity.Enabled = UserHasAccess("SocialSecurity")
   
   mnuDeptList.Enabled = UserHasAccess("DeptList")
   mnuNewDept.Enabled = UserHasAccess("NewDept")
   
   mnuChange.Enabled = UserHasAccess("DeptSettings")
   
   mnuHolidays.Enabled = UserHasAccess("DefineHolidays")
   mnuDeptList.Enabled = UserHasAccess("DeptList")
   mnuSalaryLedger.Enabled = UserHasAccess("SalaryLedger")
   mnuSheet.Enabled = UserHasAccess("TakeAttendance")
   mnuOverTime.Enabled = UserHasAccess("AddOverTime")
   mnuLTLedger.Enabled = UserHasAccess("LongTermLoan")
   mnuTaxRanges.Enabled = UserHasAccess("TaxRanges")
   
   'Users
   mnuUserManager.Enabled = UserHasAccess("UserManagement")
   mnuReminders.Visible = UserHasAccess("Reminders")
   With Toolbar1.Buttons
      .Item("Emp").Enabled = mnuEmpList.Enabled
      .Item("AdvancesLedger").Enabled = mnuShortAdv.Enabled
      .Item("TakeAtt").Enabled = mnutake.Enabled
      .Item("NewShort").Enabled = mnuShort.Enabled
      .Item("NewLong").Enabled = mnuLong.Enabled
      .Item("EmpLedger").Enabled = mnuEmpLedger.Enabled
      .Item("SalSheet").Enabled = mnuSalSheet.Enabled
      .Item("OverTime").Enabled = mnuOverTime.Enabled
      .Item("Dept").Enabled = mnuDeptList.Enabled
   End With

End Sub
