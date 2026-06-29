VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmNewEmpFine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Fine"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   315
      Left            =   1755
      TabIndex        =   1
      Top             =   375
      Width           =   6630
      _ExtentX        =   11695
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
      Left            =   1755
      TabIndex        =   0
      Top             =   30
      Width           =   6630
      _ExtentX        =   11695
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
   Begin VB.TextBox txtDescription 
      Height          =   3150
      Left            =   1755
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   1530
      Width           =   6630
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "Print When Saved"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1620
      TabIndex        =   10
      Top             =   4935
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1800
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1755
      TabIndex        =   2
      Top             =   765
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dddd dd-MMM-yyyy"
      Format          =   76414979
      CurrentDate     =   37384
   End
   Begin MSComCtl2.DTPicker DTDedStartMonth 
      Height          =   315
      Left            =   6135
      TabIndex        =   12
      Top             =   765
      Visible         =   0   'False
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MMMM, yyyy"
      Format          =   76414979
      CurrentDate     =   40229
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Description :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   3
      Left            =   810
      TabIndex        =   14
      Top             =   1515
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ded. Month :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   16
      Left            =   5145
      TabIndex        =   13
      Top             =   795
      Visible         =   0   'False
      Width           =   945
   End
   Begin MSForms.TextBox txtBal 
      Height          =   315
      Left            =   1740
      TabIndex        =   11
      Top             =   1155
      Width           =   1050
      VariousPropertyBits=   746604571
      BackColor       =   16761024
      ForeColor       =   255
      Size            =   "1852;556"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   3630
      X2              =   8330
      Y1              =   4830
      Y2              =   4830
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Amount Paid :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   705
      TabIndex        =   9
      Top             =   1185
      Width           =   1005
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   360
      Left            =   5190
      TabIndex        =   4
      Top             =   4905
      Width           =   1560
      Caption         =   "Save & Close"
      PicturePosition =   327683
      Size            =   "2752;635"
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdNext 
      Height          =   360
      Left            =   3570
      TabIndex        =   3
      Top             =   4890
      Width           =   1560
      Caption         =   "Save & New    "
      PicturePosition =   327683
      Size            =   "2752;635"
      Accelerator     =   78
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   360
      Left            =   6810
      TabIndex        =   5
      Top             =   4890
      Width           =   1560
      Caption         =   "Cancel    "
      PicturePosition =   327683
      Size            =   "2752;635"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fine Date :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   915
      TabIndex        =   8
      Top             =   780
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Employee        :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   420
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Department    :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   4
      Left            =   135
      TabIndex        =   6
      Top             =   90
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   3630
      X2              =   8330
      Y1              =   4845
      Y2              =   4845
   End
End
Attribute VB_Name = "frmNewEmpFine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LongTerm As Boolean
Dim m_bContractors As Boolean

Private Sub PrintReport(Optional bDirectPrint As Boolean = False)

    Screen.MousePointer = vbHourglass
    Dim lID As Long
    lID = Val(con.Execute("Select Max(EntryID) From Advances").Fields(0).Value & "")
    
    con.Execute "Delete From PrintTempAdv"
    
    con.Execute "Insert Into PrintTempAdv(EmpID,BSal,Rate,SDays,SAmt,Total,Paid,LongTerm,Balance,DT,LunchAmt,AmtPaid,AdvancesID) Values('" & _
     cmbEmp.ID & "'," & Val(txtSalary) & "," & Val(txtRate) & "," & Val(txtHrsWorked) & "," & Val(txtAmount) & "," & _
     Val(txtAmt) & "," & Val(txtPrevST) & "," & Val(txtLongTerm) & "," & Val(txtBalance) & ",'" & DTPicker1 & "'," & Val(txtLunchAmt) & "," & Val(txtBal) & "," & lID & ")"
    
    Dim rpt As New CRAXDDRT.Report
    
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptTempAdv.rpt")
    
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
    
    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & strCompany & "'"
        ElseIf FormulaField.Name = "{@FromTo}" Then
            FormulaField.Text = "'" & Format(DTPicker1, "dd-MMM-yyyy") & "'"
        ElseIf FormulaField.Name = "{@MonthDays}" Then
            FormulaField.Text = Val(txtWorkingHrs) 'lWHrs * 8 'Val(txtTotalHrs)
        End If
    Next
   
    
    If bDirectPrint Then
        PrintReportDirectly " ", rpt
    Else
        Me.Hide
        Dim f As New frmPrevRpt
        f.ShowReport " ", rpt
    End If
        
    Screen.MousePointer = vbDefault

End Sub

Private Sub chkOtherAccount_Click()
    If chkOtherAccount.Value = vbChecked Then
        cmbAccNo.Visible = True
        cmbCashInHand.Visible = False
        cmbAccNo.SetFocus
    Else
        cmbAccNo.Visible = False
        cmbCashInHand.Visible = True
        cmbCashInHand.SetFocus
    End If
    Call DTPicker1_Change
End Sub

Private Sub cmbdept_matched()

    cmbEmp.ClearVals
    'cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.Text = "<All Departments>" Then
        If m_bContractors Then
            cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' - ' + FName + ' {' + Designation + '}'", "VTempEmps", "EmpID", "Where Active=1 "
        Else
            cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' - ' + FName + ' {' + Designation + '}'", "VSalariedEmps", "EmpID", "Where Active=1 "
        End If
    Else
        cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' - ' + FName + ' {' + Designation + '}'", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "' and Active=1"
    End If
    
End Sub

Private Function Saved() As Boolean

    Saved = False
     
    Dim SQL As String
    
    If Not cmbEmp.MatchFound Then
        MsgBox "Please Select Valid Employee From The List.", vbInformation
        Exit Function
    ElseIf Val(txtBal) = 0 Then
        MsgBox "Invalid Amount.", vbInformation
        Exit Function
    End If
   
    Dim strVchrNo As String
    Call StartTrans(con)
    
'    If chkPost.Value = vbChecked Then
'        strVchrNo = AddVoucher
'    End If
    
    'Sql = "INSERT INTO Advances(EmpID,DT,Type,Description," & _
     "Amount,DAmount,AmountCleared,AccVoucherNo,DeductionStartDT,ActualAmountTaken) VALUES('" & cmbEmp.ID & _
     "','" & DTPicker1 & "',0,'Short Term'," & Val(txtBal) & ",0,0,'" & _
     strVchrNo & "','" & DTDedStartMonth.Value & "'," & Val(txtBal) & ")"
    
    'con.Execute Sql
    con.Execute "INSERT INTO EmpFine(EmpID,FineDate,Amount,Description) VALUES('" & cmbEmp.ID & "','" & _
     DTPicker1.Value & "'," & Val(txtBal) & ",'" & txtDescription & "')"
    con.CommitTrans
        
    MsgBox "Saved Successfully.", vbInformation
    
    Saved = True
    'If chkPrint = 1 Then Call PrintVoucher(MainForm.cr1, strVchrNo)
    SQL = ""
    Exit Function
err:
    MsgBox err.Description


End Function

Private Function AddVoucher() As String

    Dim i As Long, INextSno As Double, Bal As Double, Amt As Double
    Dim VchrNo As String, theDesc As String, AccNo As String
    
    Dim Dated As Date
    Dated = DTPicker1
    'Charge All Amount to Cash
    Amt = Val(txtBal)
    
    If Amt = 0 Then Exit Function
    
    'ConAcc.BeginTrans
     
    'Get The Required Values
    Dim strVchrType As String
    If chkOtherAccount.Value = vbChecked Then
        strVchrType = "JV"
    Else
        strVchrType = "CPV"
    End If
    
    Dim bWaitForAuth As Boolean
    bWaitForAuth = VoucherRequiresAuthorization(strVchrType)
    
    INextSno = getNextSno(Dated)
    VchrNo = getNextVchrNo(Dated, strVchrType)
     
    'Give The Description
    theDesc = "Temp. Advance Paid To " & cmbEmp.Text
     
    ''''''''''''''''Charge To Cash Entry'''''''''''''
    
    If chkOtherAccount.Value = vbChecked Then
        AccNo = cmbAccNo.ID
    Else
        AccNo = cmbCashInHand.List(cmbCashInHand.ListIndex, 1) 'CashInHand
    End If
    
    If bWaitForAuth = False Then
        'AccNo = cmbCashInHand.List(cmbCashInHand.ListIndex, 1) 'CashInHand
        Bal = getBalance(AccNo, Dated)
        Bal = Bal - Amt
        
        'Credit Entry
        con.Execute "INSERT INTO Vouchers VALUES(" & INextSno & ",'" & _
         Dated & "','" & VchrNo & "','" & AccNo & _
         "','" & theDesc & "',0," & Amt & "," & Bal & ",'',0)"
    
        Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
     
        INextSno = INextSno + 1
     
        '''''''''''''''''' D E B I T   E N T R I E S '''''''''''''
'        If (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Then
'            AccNo = "15-004-01001"
'        ElseIf strCompany = "Banzai" Then
'            AccNo = "15-004-19001"
'        Else
'            AccNo = "13-003-19001" 'IIf(LongTerm, LongAdvAccNo, ShortAdvAccNo)
'        End If
        AccNo = strEmpShortTermAccNo
        Bal = getBalance(AccNo, Dated)
        Amt = Amt
        Bal = Bal + Amt
     
        con.Execute "INSERT INTO Vouchers VALUES(" & INextSno & ",'" & _
         Dated & "','" & VchrNo & "','" & AccNo & _
         "','" & theDesc & "'," & Amt & ",0," & Bal & ",'',0)"
        
        Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
    
        con.Execute "INSERT INTO VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & VchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
    Else
        con.Execute "INSERT INTO VouchersForAuthorization(VDate,VchrNo,Accno,Description,Debit,Credit,balance,DpstSlip,CSNo) VALUES('" & Dated & "','" & _
             VchrNo & "','" & AccNo & "','" & theDesc & "',0," & Amt & ",0,'',0)"
        
        con.Execute "INSERT INTO VouchersForAuthorization(VDate,VchrNo,Accno,Description,Debit,Credit,balance,DpstSlip,CSNo) VALUES('" & Dated & "','" & _
             VchrNo & "','" & strEmpShortTermAccNo & "','" & theDesc & "'," & Amt & ",0,0,'',0)"
             
        con.Execute "INSERT INTO VouchersForAuthorizationInfo(VchrNo,UserName,MachineName,DT,PostedThroughJVForm,Notes,Payee) VALUES('" & VchrNo & "','" & _
         CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "',0,'','')"
    End If
        
    If VchrNo <> txtVchrNo Then
        MsgBox "Vchr No. " & txtVchrNo & " had been used, " & VchrNo & " is used Instead.", vbInformation
    End If
    
    AddVoucher = VchrNo
    
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()

    If Saved Then
        Call DTPicker1_Change
        txtBal = ""
        cmbEmp.ID = ""
        cmbEmp.SetFocus
        'cmbDept.SetFocus
    End If
    
End Sub

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub

Private Sub DTDedStartMonth_Change()
    DTDedStartMonth = DateSerial(DTDedStartMonth.year, DTDedStartMonth.Month + 1, 0)
End Sub

Private Sub DTPicker1_Change()

'    If cmbEmp.MatchFound Then
'        Call cmbEmp_matched
'    End If
    DTDedStartMonth.Value = DTPicker1.Value
    
'    Dim strVchrType As String
'    If chkOtherAccount.Value = vbChecked Then
'        strVchrType = "JV"
'    Else
'        strVchrType = "CPV"
'    End If
'
'    txtVchrNo = getNextVchrNo(DTPicker1.Value, strVchrType)
    
    
End Sub

Private Sub Form_Load()

    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
      
    DTPicker1 = GetServerDate
    Call DTPicker1_Change
    m_bContractors = False
    chkPrint = Abs(bPrintWhenSaved)
     
'    cmbAccNo.ListHeight = 1600
'    cmbAccNo.AddVals con, "'{' + AccNo  + '} ' + AccTitle", "VActiveAccounts", "AccNo"
    
'    If strCompany = "Tecno" Or strCompany = "IAA" OR strCompany = "Kami" Then
'        chkPost.Visible = True
'        If strCompany = "IAA" OR strCompany = "Kami" Then
'            chkPost.Value = vbUnchecked
'        End If
'    Else
'        chkPost.Visible = False
'    End If
    
    
    
End Sub

Private Sub txtAmt_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyNumber(KeyAscii.Value, True)
End Sub

Private Sub txtDeduct_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyNumber(KeyAscii.Value, True)
End Sub

'To Check The Long Term Settings That Are Made For Departments
Private Function IsNotLongApplicable() As Boolean

    Dim Rec As New ADODB.Recordset
    'dim Item as New
    
    With Rec
        .CursorLocation = adUseClient
        .Open "SELECT JoinDate,LongTermTime,LongTermLimit FROM VOnlyEmps WHERE EmpID='" & cmbEmp.ID & "'", con, adOpenKeyset, adLockReadOnly
        
        If DateDiff("m", !JoinDate, DTPicker1) < !longtermtime Then
            MsgBox "Could Not Give Long Term Loan To Employee." & vbNewLine & "It Can Be Given After " & !longtermtime & " Months", vbInformation
            IsNotLongApplicable = True
        ElseIf Val(txtAmt) > !LongTermLimit Then
            MsgBox "Could Not Give Long Term Loan To Employee. Because Amount " & vbNewLine & " Is Greater Than Long Term Limit.", vbInformation
            IsNotLongApplicable = True
        End If
        
    End With


End Function

Public Sub ShowMe(p_bContractors)
    
    m_bContractors = p_bContractors
    Dim strCond As String
    
    If m_bContractors Then
        strCond = "Where IsNull(TempDept,0)=1"
    Else
        strCond = "Where IsNull(TempDept,0)=0"
    End If
    
    With cmbDept
        .ClearVals
        .AddItem "<All Departments>", "0"
        .AddVals con, "Name", "Departments", "DeptID", strCond
        .Text = "<All Departments>"
    End With
'    Dim rs As New ADODB.Recordset
'    With rs
'        .Open "SELECT AccNo,AccTitle FROM VActiveAccounts WHERE SubAccOf='15-001'", con, adOpenForwardOnly, adLockReadOnly
'        cmbCashInHand.Clear
'        Do Until .EOF
'            cmbCashInHand.AddItem !AccTitle & ""
'            cmbCashInHand.List(.AbsolutePosition - 1, 1) = !AccNo & ""
'            .MoveNEXT
'        Loop
'        .Close
'    End With
'    Set rs = Nothing
'    If cmbCashInHand.ListCount > 0 Then cmbCashInHand.ListIndex = i
   
'    If strCompany = "Tecno" Then
'        Label2.Visible = False
'        cmbCashInHand.Visible = False
'        chkPrint.Value = vbUnchecked
'        chkPrint.Visible = False
'    End If
    
    Me.Show 1
    
End Sub

