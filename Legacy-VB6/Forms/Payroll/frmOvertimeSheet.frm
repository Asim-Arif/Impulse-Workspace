VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOvertimeSheet 
   BorderStyle     =   0  'None
   ClientHeight    =   7035
   ClientLeft      =   15
   ClientTop       =   195
   ClientWidth     =   11505
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleMode       =   0  'User
   ScaleWidth      =   10305.19
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   150
      Top             =   420
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
            Picture         =   "frmOvertimeSheet.frx":0000
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOvertimeSheet.frx":0452
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker DT 
      Height          =   315
      Left            =   330
      TabIndex        =   5
      Top             =   360
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
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
      CustomFormat    =   "ddd dd-MMM-yyyy"
      Format          =   20709379
      CurrentDate     =   37886
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   180
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   10
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOvertimeSheet.frx":0564
            Key             =   "Paid"
         EndProperty
      EndProperty
   End
   Begin VB.Frame FLV 
      Height          =   6225
      Left            =   15
      TabIndex        =   0
      Top             =   705
      Width           =   11385
      Begin VB.Frame FButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   930
         Left            =   7935
         TabIndex        =   8
         Top             =   5250
         Width           =   3330
         Begin MSForms.CommandButton cmdPrint 
            Height          =   360
            Left            =   105
            TabIndex        =   12
            Top             =   480
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
         Begin MSForms.CommandButton cmdClose 
            Height          =   360
            Left            =   1605
            TabIndex        =   11
            Top             =   465
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
         Begin VB.Label lblBalance 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   270
            Left            =   1455
            TabIndex        =   10
            Top             =   30
            Width           =   1590
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Total Amount :"
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
            Height          =   195
            Left            =   180
            TabIndex        =   9
            Top             =   75
            Width           =   1245
         End
      End
      Begin VB.TextBox txtHours 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   5130
         TabIndex        =   4
         Top             =   1245
         Visible         =   0   'False
         Width           =   915
      End
      Begin MSComctlLib.ListView LV 
         Height          =   4950
         Left            =   105
         TabIndex        =   1
         Top             =   180
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   8731
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Emp ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Employee Name"
            Object.Width           =   5644
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Departement"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "O.T. Hours"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Rate"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "O.T. Amt"
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin MSComCtl2.DTPicker DtTo 
      Height          =   300
      Left            =   3840
      TabIndex        =   6
      Top             =   360
      Width           =   1755
      _ExtentX        =   3096
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
      CustomFormat    =   "ddd dd-MMM-yyyy"
      Format          =   20709379
      CurrentDate     =   37081
   End
   Begin MSComCtl2.DTPicker DtFrom 
      Height          =   300
      Left            =   2085
      TabIndex        =   7
      Top             =   360
      Width           =   1755
      _ExtentX        =   3096
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
      CustomFormat    =   "ddd dd-MMM-yyyy"
      Format          =   20709379
      CurrentDate     =   37055
   End
   Begin MSForms.Label Label2 
      Height          =   270
      Left            =   330
      TabIndex        =   3
      Top             =   90
      Width           =   5235
      BackColor       =   11517387
      VariousPropertyBits=   8388627
      Caption         =   " Payment Date           From                               To"
      Size            =   "9234;476"
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
      Height          =   375
      Left            =   6570
      TabIndex        =   2
      Top             =   285
      Width           =   1395
      Caption         =   "Refresh        "
      PicturePosition =   327683
      Size            =   "2461;661"
      Accelerator     =   82
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
      Begin VB.Menu mnuDetail 
         Caption         =   "Current Employees Detail"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSummary 
         Caption         =   "Print Over Time Payment Sheet"
      End
      Begin VB.Menu dashsumm 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCurrency 
         Caption         =   "Print Currency Denomination"
      End
      Begin VB.Menu dash 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOTPayslip 
         Caption         =   "Over Time Payment Slip (Current Employee)"
      End
      Begin VB.Menu mnuOTPayslips 
         Caption         =   "Over Time Payment Slip (All Employees)"
      End
   End
End
Attribute VB_Name = "frmOvertimeSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Thousands As Double, FHundreds As Double, Hundreds As Double, Fifties As Double, Tens As Double, Fives As Double, TotalSal As Double, Twos As Double, Ones As Double



Private Sub cmbPymtDate_Click()
    If cmbPymtDate.MatchFound And cmbPymtDate <> "" Then Call cmdRefresh_Click
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub PrintDetail()

    On Error GoTo err
      
      If Not cmbPymtDate.MatchFound Or LV.ListItems.Count = 0 Then Exit Sub
      
      
    Screen.MousePointer = vbHourglass
    Dim Cond As String
    
    
    
    
    Cond = "{VOTPayments.empid}='" & LV.SelectedItem.Text & "' AND {VOTPayments.PaidOn}=date(" & Format(cmbPymtDate, "yyyy,MM,dd") & ")"
    
    
With Cr
    .ReportFileName = App.Path & "\PayOverTime.rpt"
    .DataFiles(0) = DatabasePath
    .ReplaceSelectionFormula Cond
    .Formulas(0) = "FromTo='From " & Format(DTFrom, "dd-MMM-yyyy") & " To " & Format(DTTo, "dd-MMM-yyyy") & "'"
    .Action = 1
    .PageZoomNext
    .Formulas(0) = ""
End With
Screen.MousePointer = vbDefault
Exit Sub

err:
    MsgBox err.Description
    Screen.MousePointer = vbDefault
End Sub

Private Sub PrintSummary()

    Dim Cond As String
    
    Cond = "{VOverTime.PaidOn}=date(" & Format(DT, "yyyy,MM,dd") & ")"
    
    
'    With cr1
'        .ReportFileName = RptPath & "\OverTimePaymt.rpt"
'        '.DataFiles(0) = DatabasePath
'        .SelectionFormula = Cond
'
'
'        .Formulas(0) = "FromTo='" & Format(DT, "ddd dd-mmm-yyyy") & "'"
'        .Formulas(1) = "Company='" & company & "'"
'        .Action = 1
'        .PageZoomNext
'        .Formulas(0) = ""
'    End With
    
Screen.MousePointer = vbHourglass

Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
Set rpt = rpt.OpenSubreport(rptPayrollPath & "\rptOverTimePayment.rpt")
Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

Set FormulaFields = rpt.FormulaFields

For Each FormulaField In FormulaFields
    If FormulaField.Name = "{@Company}" Then
        FormulaField.Text = "'" & company & "'"
    ElseIf FormulaField.Name = "{@FromTo}" Then
        FormulaField.Text = "'" & Format(DT, "ddd dd-mmm-yyyy") & "'"
    End If
Next


f.ShowReport Cond & " ", rpt

Screen.MousePointer = vbDefault


End Sub

Private Sub PrintPaySlip(Optional EmpId As String)

    Dim Cond As String
    
    If EmpId = "" Then
        Cond = "{VEmp.EmpID}='" & LV.ListItems(1) & "'"
        
        For i = 2 To LV.ListItems.Count
            Cond = Cond & " OR {VEmp.EmpID}='" & LV.ListItems(i) & "'"
        Next
    Else
        Cond = "{VEmp.EmpID}='" & EmpId & "'"
    End If
'With cr1
'    .ReportFileName = RptPath & "\OverTimePayslip.rpt"
'    '.DataFiles(0) = DatabasePath
'    .SelectionFormula = Cond
'    .Formulas(0) = "Rate=" & LV.SelectedItem.SubItems(4) & ""
'    .Formulas(1) = "Amount=" & LV.SelectedItem.SubItems(5) & ""
'    .Formulas(2) = "Hrs='" & LV.SelectedItem.SubItems(3) & "'"
'
'    .Action = 1
'    .PageZoomNext
'End With

Screen.MousePointer = vbHourglass

Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
Set rpt = rpt.OpenSubreport(rptPayrollPath & "\rptOverTimePayslip.rpt")
Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

Set FormulaFields = rpt.FormulaFields

For Each FormulaField In FormulaFields
    If FormulaField.Name = "{@Company}" Then
        FormulaField.Text = "'" & company & "'"
    ElseIf FormulaField.Name = "{@Rate}" Then
        FormulaField.Text = LV.SelectedItem.SubItems(4)
    ElseIf FormulaField.Name = "{@Amount}" Then
        FormulaField.Text = LV.SelectedItem.SubItems(5)
    ElseIf FormulaField.Name = "{@Hrs}" Then
        FormulaField.Text = "'" & LV.SelectedItem.SubItems(3) & "'"
    ElseIf FormulaField.Name = "{@ForMonth}" Then
        FormulaField.Text = "'From " & Format(DTFrom, "dd-MMM-yyyy") & " To " & Format(DTTo, "dd-MMM-yyyy") & "'"
    End If
Next


f.ShowReport Cond & " ", rpt

Screen.MousePointer = vbDefault

End Sub



Private Sub cmdPrint_Click()
PopUpMenu mnuPop
End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLedger
End Sub



Private Sub DT_Change()

Dim DFrom As String, DTo As String

DFrom = con.Execute("Select MIN(DT) FROM AttendanceSheet WHERE PaidOn='" & DT & "'").Fields(0) & ""
DTo = con.Execute("Select MAX(DT) FROM AttendanceSheet WHERE PaidOn='" & DT & "'").Fields(0) & ""

DTFrom = IIf(DFrom = "", Date, DFrom)
DTTo = IIf(DTo = "", Date, DTo)

Call cmdRefresh_Click

End Sub

Private Sub Form_Load()

    mnuPop.Visible = False
    DT = Date
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Dim TopMargin As Integer
    TopMargin = DT.Top + DT.Height + 30
    
    FLV.Move 50, TopMargin, Me.ScaleWidth - 100, Me.ScaleHeight - TopMargin
    
    LV.Move 100, 200, Me.Width - 400, FLV.Height - fButs1.Height - 300
    
    fButs1.Move LV.Left + LV.Width - fButs1.Width, LV.Top + LV.Height

'cmdRefresh.Move FLV.Left + FLV.Width - cmdRefresh.Width
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

    Dim TotalAmt As Double
    
    Dim rsLedger As New ADODB.Recordset
    
    Dim ITM As ListItem
    Dim Cond As String
    Dim SQL As String, OTAmount As Double, OverTime As Double
    
    LV.Checkboxes = chkOnlyPayable
    
    SQL = "SELECT DISTINCT(EmpID),Name,DeptName,Rate,ExtraHrs FROM VOverTime  Where PaidOn='" & DT & "' Order By EmpID"
    
    
    With rsLedger
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        LV.Visible = False
        
        LV.ListItems.Clear
        
        For i = 0 To .RecordCount - 1
            OverTime = con.Execute("Select SUM(ExtraHrs) FROM AttendanceSheet WHERE EmpID='" & !EmpId & "' AND PaidOn='" & DT & "'").Fields(0)
            OTAmount = Val(!Rate & "") * Val(OverTime & "")
            Set ITM = LV.ListItems.add(, , ![EmpId] & "")
               ITM.SubItems(1) = ![Name] & ""
               ITM.SubItems(2) = ![Deptname] & ""
               ITM.SubItems(3) = MinsToTime([OverTime] * 60)
               ITM.SubItems(4) = Format(![Rate] & "", "0.00")
               ITM.SubItems(5) = Format(OTAmount, "0.00")
               
               'Get Total Amount
               TotalAmt = TotalAmt + OTAmount
               
               .MoveNext
               
        Next
        
        LV.Visible = True
        .Close
    End With
    
    Set rsLedger = Nothing
    
    SQL = ""
    Cond = ""
    
    lblBalance = Format(TotalAmt, "0.00")
    
End Sub




Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)
   If Item.Bold Then
      For i = 1 To LV.ListItems.Count
         LV.ListItems(i).Checked = True
      Next
   End If
   'Call GetTotals
End Sub

Private Sub mnuDetail_Click()
Call PrintDetail
End Sub

Private Sub mnuOTPayslip_Click()


Call PrintPaySlip(LV.SelectedItem)

End Sub

Private Sub mnuOTPayslips_Click()
    Call PrintPaySlip
End Sub

Private Sub mnuSummary_Click()
    Call PrintSummary
End Sub

Private Sub mnuSummaryPaid_Click()
    'Call PrintSummary(True)
End Sub

Private Sub txtHours_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       con.Execute "Update OverTime Set Hours=" & Val(txtHours) & " Where Empid='" & LV.SelectedItem & "' and dt=#" & CDate(LV.SelectedItem.Tag) & "#"
       LV.SelectedItem.SubItems(4) = Val(txtHours)
       LV.SetFocus
    ElseIf KeyAscii = 27 Then
       LV.SetFocus
    Else
       KeyAscii = OnlyNumber(KeyAscii, True)
    End If
   
End Sub

Private Sub txtHours_LostFocus()
   txtHours.Visible = False
End Sub



'Private Sub mnuCurrency_Click()
'
'    Thousands = 0
'    FHundreds = 0
'    Hundreds = 0
'    Fifties = 0
'    Tens = 0
'    Fives = 0
'    Twos = 0
'    Ones = 0
'    TotalSal = 0
'
'
'    For i = 1 To LV.ListItems.Count
'        CountRupee Val(LV.ListItems(i).ListSubItems(5))
'        TotalSal = TotalSal + Val(LV.ListItems(i).ListSubItems(5))
'    Next
'
'
'
'    With Cr
'        .ReportFileName = RptPath & "\RupeeStructure.rpt"
'
'            .Formulas(0) = "Thousands=" & Thousands
'            .Formulas(1) = "FHundreds=" & FHundreds
'            .Formulas(2) = "Hundreds=" & Hundreds
'            .Formulas(3) = "Fifties=" & Fifties
'            .Formulas(4) = "Tens=" & Tens
'            .Formulas(5) = "Fives=" & Fives
'            .Formulas(6) = "Tows=" & Twos
'            .Formulas(7) = "Ones=" & Ones
'            .Formulas(8) = "TotalSal=" & Round(TotalSal)
'            .Formulas(9) = "Company='" & company & "'"
'
'        .WindowState = crptMaximized
'        .WindowTitle = "Print Rupee Structure"
'        .Action = 1
'
'    End With
'
'End Sub
'
Private Sub mnuCurrency_Click()

    Thousands = 0
    FHundreds = 0
    Hundreds = 0
    Fifties = 0
    Tens = 0
    Fives = 0
    Twos = 0
    Ones = 0
    TotalSal = 0


    For i = 1 To LV.ListItems.Count
        CountRupee Val(LV.ListItems(i).ListSubItems(5))
        TotalSal = TotalSal + Val(LV.ListItems(i).ListSubItems(5))
    Next
    
'    With CR2
'        .ReportFileName = RptPath & "\RupeeStructure.rpt"
'
'            .Formulas(0) = "Thousands=" & Thousands
'            .Formulas(1) = "FHundreds=" & FHundreds
'            .Formulas(2) = "Hundreds=" & Hundreds
'            .Formulas(3) = "Fifties=" & Fifties
'            .Formulas(4) = "Tens=" & Tens
'            .Formulas(5) = "Fives=" & Fives
'            .Formulas(6) = "Tows=" & Twos
'            .Formulas(7) = "Ones=" & Ones
'            .Formulas(8) = "TotalSal='" & lblTotal & "'"
'            .Formulas(9) = "Company='" & company & "'"
'
'        .WindowState = crptMaximized
'        .WindowTitle = "Print Rupee Structure"
'        .Action = 1
'
'    End With

Screen.MousePointer = vbHourglass

Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
Set rpt = rpt.OpenSubreport(rptPayrollPath & "\rptRupeeStructure.rpt")
Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

Set FormulaFields = rpt.FormulaFields
For Each FormulaField In FormulaFields
    If FormulaField.Name = "{@Company}" Then
        FormulaField.Text = "'" & company & "'"
    ElseIf FormulaField.Name = "{@Thousands}" Then
        FormulaField.Text = Thousands
    ElseIf FormulaField.Name = "{@FHundreds}" Then
        FormulaField.Text = FHundreds
    ElseIf FormulaField.Name = "{@Hundreds}" Then
        FormulaField.Text = Hundreds
    ElseIf FormulaField.Name = "{@Fifties}" Then
        FormulaField.Text = Fifties
    ElseIf FormulaField.Name = "{@Tens}" Then
        FormulaField.Text = Tens
    ElseIf FormulaField.Name = "{@Fives}" Then
        FormulaField.Text = Fives
    ElseIf FormulaField.Name = "{@Tows}" Then
        FormulaField.Text = Twos
    ElseIf FormulaField.Name = "{@Ones}" Then
        FormulaField.Text = Ones
    ElseIf FormulaField.Name = "{@TotalSal}" Then
        FormulaField.Text = Round(TotalSal)
    End If
Next

f.ShowReport " ", rpt


Screen.MousePointer = vbDefault


End Sub

Private Sub CountRupee(Salary As Double)

    Thousands = Thousands + Int(Salary / 1000)
    If Salary >= 500 Then FHundreds = FHundreds + Int(Val(Right(Salary, 3)) / 500)
    Hundreds = Hundreds + Int(Val(Right(Salary, 3) Mod 500) / 100)
    Fifties = Fifties + Int(Val(Right(Salary, 2)) / 50)
    Tens = Tens + Int(Val(Right(Salary, 3) Mod 50) / 10)
    Fives = Fives + Int(Val(Right(Salary, 1)) / 5)
    Twos = Twos + Int(Val(Right(Salary, 1) Mod 5) / 2)
    Ones = Ones + Int(Val(Right(Salary, 1) Mod 5) Mod 2)


End Sub



