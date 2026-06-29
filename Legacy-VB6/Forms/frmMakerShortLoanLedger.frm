VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmMakerShortLoanLedger 
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   11550
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   7425
   ScaleMode       =   0  'User
   ScaleWidth      =   10345.49
   WindowState     =   2  'Maximized
   Begin VB.Frame fTitle 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   450
      TabIndex        =   11
      Top             =   0
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Short Term Loan Ledger"
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
         Left            =   2985
         TabIndex        =   12
         Top             =   120
         Width           =   5115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Short Term Loan Ledger"
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
         Left            =   3000
         TabIndex        =   13
         Top             =   135
         Width           =   5115
      End
   End
   Begin ComboList.Usercontrol1 cmbMaker 
      Height          =   315
      Left            =   690
      TabIndex        =   5
      Top             =   915
      Width           =   6390
      _ExtentX        =   11271
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
   Begin MSComCtl2.DTPicker DtTo 
      Height          =   300
      Left            =   8460
      TabIndex        =   2
      Top             =   915
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
      Format          =   89522179
      CurrentDate     =   37081
   End
   Begin MSComCtl2.DTPicker DtFrom 
      Height          =   300
      Left            =   7110
      TabIndex        =   3
      Top             =   915
      Width           =   1335
      _ExtentX        =   2355
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
      Format          =   89522179
      CurrentDate     =   37055
   End
   Begin VB.Frame FLV 
      Height          =   6045
      Left            =   30
      TabIndex        =   0
      Top             =   1260
      Width           =   11385
      Begin VB.Frame fButs2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   525
         Left            =   60
         TabIndex        =   9
         Top             =   5475
         Width           =   5505
      End
      Begin VB.Frame fButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   7080
         TabIndex        =   6
         Top             =   5445
         Width           =   3000
         Begin MSForms.CommandButton cmdClose 
            Height          =   360
            Left            =   1455
            TabIndex        =   8
            Top             =   135
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
            Left            =   -15
            TabIndex        =   7
            Top             =   135
            Width           =   1425
            ForeColor       =   0
            Caption         =   "Print         "
            PicturePosition =   327683
            Size            =   "2514;635"
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
      Begin MSComctlLib.ListView LV 
         Height          =   5265
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   9287
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   4322
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Voucher #"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Balance"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Ded/Month"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmMakerShortLoanLedger.frx":0000
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
      Left            =   690
      TabIndex        =   10
      Top             =   675
      Width           =   9135
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   10200
      Picture         =   "frmMakerShortLoanLedger.frx":00A6
      Top             =   1395
      Width           =   465
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   9900
      TabIndex        =   4
      Top             =   840
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
      Begin VB.Menu mnuLoanList 
         Caption         =   "Loan List"
      End
      Begin VB.Menu mnuPayslip 
         Caption         =   "Loan Payslip"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnupop1 
      Caption         =   "mnupop1"
      Begin VB.Menu mnuDelEntry 
         Caption         =   "Remove This Entry"
      End
   End
End
Attribute VB_Name = "frmMakerShortLoanLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String
Dim Openbalance As Double

Private Sub cmbmaker_matched()
    cmbMaker.Tag = GetSingleStringValue("AccNo", "Makers", " WHERE VendID=" & cmbMaker.ID)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    'PopUpMenu mnuPopup
    Dim lOpenBalance As Long
    lOpenBalance = GetSingleLongValue("SUM(Amount)", "VMakerAdvancesLedger_Short", " WHERE AccNo='" & cmbMaker.ID & "' AND DT<'" & DtFrom.Value & "'")
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\MakerAdvancesLedger_Short.rpt")
    rpt.FormulaFields.GetItemByName("OpeningBalance").Text = lOpenBalance
    Dim strSelection As String
    strSelection = "{VMakerAdvancesLedger_Short.AccNo}='" & cmbMaker.ID & "' AND {VMakerAdvancesLedger_Short.DT}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#"
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLedger
End Sub

Private Sub Form_Load()

    mnuPopup.Visible = False
    mnupop1.Visible = False
    
    DtTo = GetServerDate
    DtFrom = DateAdd("M", -1, GetServerDate)
    
    'Call cmbMaker.AddVals(con, "'{ ' + VendID1 + ' } ' + VenderName ", "Makers", "VendID", " WHERE VendID<>" & iFactoryMakerID)
    Call cmbMaker.AddVals(con, "'{ ' + MakerNo + ' } ' + AccTitle ", "VVendersAndMakers", "AccNo")
    
End Sub
Private Sub cmbdept_matched()

    cmbEmp.ClearVals
    'cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.Text = "<All Departments>" Then
        cmbEmp.AddVals con, "'{' + EmpID + '} ' + Name", "Employees", "EmpID", " "
    Else
        cmbEmp.AddVals con, "'{' + EmpID + '} ' + Name", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "'"
    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
    
    Dim TopMargin As Integer
    TopMargin = cmbMaker.Top + cmbMaker.Height + 30
    
    FLV.Move 50, TopMargin, Me.ScaleWidth - 100, Me.ScaleHeight - TopMargin
    
    LV.Move 100, 200, Me.Width - 400, FLV.Height - fButs1.Height - 300
    
    fButs1.Move LV.Left + LV.Width - fButs1.Width, LV.Top + LV.Height
    
    fButs2.Move LV.Left, LV.Top + LV.Height
End Sub



Private Sub RefreshLedger()

    Dim rsLedger As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rsLedger
        
        .Open "SELECT * FROM VMakerAdvancesLedger_Short WHERE AccNo='" & cmbMaker.ID & "' AND DT BETWEEN '" & DtFrom.Value & "' AND '" & DtTo.Value & "' ORDER BY DT,Amount DESC", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        
        'SQL = "(SELECT ISNULL(Sum(AmtClrd),0) Balance FROM MakerAmtCleared WHERE Dt<'" & Format(DtFrom, "dd-MMM-yyyy") & "' and AccNo='" & cmbMaker.Tag & "')"
        
        'Balance = Val(con.Execute("SELECT SUM(Amount)-" & SQL & "  FROM MakerAdvances WHERE AccNo='" & cmbMaker.Tag & "' AND Dt<'" & Format(DtFrom, "dd-MMM-yyyy") & "'").Fields(0) & "")
        Balance = GetSingleLongValue("SUM(Amount)", "VMakerAdvancesLedger_Short", " WHERE AccNo='" & cmbMaker.ID & "' AND DT<'" & DtFrom.Value & "'")
        Openbalance = Balance
        Set ITM = LV.ListItems.add()
        ITM.SubItems(1) = "Previous Balance "
        ITM.SubItems(4) = Balance
        
        For i = 0 To .RecordCount - 1
                    
            Set ITM = LV.ListItems.add(, , Format(![DT], "dddd dd MMMM ,yyyy"))
            ITM.Tag = !EntryID & ""
            ITM.ListSubItems.add(, , ![Description] & "").Tag = ![DT] & ""
            ITM.SubItems(2) = Val(![Amount] & "")
            ITM.SubItems(3) = ![AccVoucherNo] & ""
            Balance = Balance + Val(![Amount] & "")
            ITM.SubItems(4) = Balance
                  
            .MoveNEXT
        Next
        .Close
    End With
'    If cmbShift.ListIndex = 0 Then
'        LV.ColumnHeaders(7).Width = 0
'    Else
'            LV.ColumnHeaders(7).Width = 1200
'    End If
    Set rsLedger = Nothing
    SQL = ""
    Cond = ""
    ICO = ""
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If LV.SelectedItem Is Nothing Then Exit Sub
        If Val(LV.SelectedItem.Tag) = 0 Then Exit Sub              'If It is a Deduction Entry Exit
        PopUpMenu mnupop1
    End If
End Sub

Private Sub mnuLoanList_Click()
    
    Dim Cond As String
    Cond = "{VAdvances.empid}='" & cmbMaker.ID & "' AND ({VAdvances.DT}>=date(" & Format(DtFrom, "yyyy,MM,dd") & ") AND {VAdvances.DT}<= date(" & Format(DtTo, "yyyy,MM,dd") & ")) AND {VAdvances.Type}=1"
    
    'With cr1
    '    .ReportFileName = RptPath & "\LTLoanLedger.rpt"
    '    '.DataFiles(0) = rptpath & "\Payroll.mdb"
    '    .Connect = ConnectStr
    '    .SelectionFormula = Cond
    '
    '    .Formulas(0) = "FromTo='From " & Format(DTFrom, "dd-MMM-yyyy") & " To " & Format(DTTo, "dd-MMM-yyyy") & "'"
    '    .Formulas(1) = "PrevBal=" & Openbalance
    '    .Formulas(2) = "Company='" & strCompany & "'"
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
Set rpt = rpt.OpenSubreport(rptPayrollPath & "\rptAdvancesPayslip.rpt")
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

