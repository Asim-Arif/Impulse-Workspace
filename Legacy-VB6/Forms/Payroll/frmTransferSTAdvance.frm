VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmTransferSTAdvance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transfer Short Term Loan..."
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
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
   ScaleHeight     =   5430
   ScaleWidth      =   5100
   StartUpPosition =   2  'CenterScreen
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
      Left            =   225
      TabIndex        =   13
      Top             =   4635
      Visible         =   0   'False
      Width           =   1800
   End
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   1755
      TabIndex        =   0
      Top             =   75
      Width           =   3120
      _ExtentX        =   5503
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1755
      TabIndex        =   1
      Top             =   405
      Width           =   3120
      _ExtentX        =   5503
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
      Format          =   121372675
      CurrentDate     =   37384
   End
   Begin MSComCtl2.DTPicker DTDedStartMonth 
      Height          =   315
      Left            =   1755
      TabIndex        =   18
      Top             =   735
      Width           =   3120
      _ExtentX        =   5503
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
      Format          =   121372675
      CurrentDate     =   40229
   End
   Begin ComboList.Usercontrol1 cmbEmpTo 
      Height          =   285
      Left            =   1755
      TabIndex        =   20
      Top             =   2640
      Width           =   3120
      _ExtentX        =   5503
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Transfer To :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   4
      Left            =   780
      TabIndex        =   21
      Top             =   2685
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ded. Start Month :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   8
      Left            =   375
      TabIndex        =   19
      Top             =   750
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Deduction Amount      :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   1650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "New Ded. :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   3
      Left            =   2970
      TabIndex        =   16
      Top             =   1575
      Width           =   810
   End
   Begin MSForms.TextBox txtBal 
      Height          =   315
      Left            =   1770
      TabIndex        =   15
      Top             =   1125
      Width           =   1050
      VariousPropertyBits=   746604575
      Size            =   "1852;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox txtPrevDed 
      Height          =   315
      Left            =   1770
      TabIndex        =   14
      Top             =   1515
      Width           =   1050
      VariousPropertyBits=   746604575
      Size            =   "1852;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Description                :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   6
      Left            =   150
      TabIndex        =   12
      Top             =   1935
      Width           =   1575
   End
   Begin MSForms.TextBox txtDesc 
      Height          =   690
      Left            =   1770
      TabIndex        =   4
      Top             =   1920
      Width           =   3090
      VariousPropertyBits=   -1467987941
      ScrollBars      =   3
      Size            =   "5450;1217"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   225
      X2              =   4925
      Y1              =   4935
      Y2              =   4935
   End
   Begin MSForms.TextBox txtAmt 
      Height          =   315
      Left            =   3810
      TabIndex        =   2
      Top             =   1125
      Width           =   1050
      VariousPropertyBits=   746604571
      Size            =   "1852;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox txtDeduct 
      Height          =   315
      Left            =   3810
      TabIndex        =   3
      Top             =   1485
      Width           =   1050
      VariousPropertyBits=   746604571
      Size            =   "1852;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "New Loan :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   5
      Left            =   2940
      TabIndex        =   11
      Top             =   1185
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Previous Balance        :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   1185
      Width           =   1635
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   360
      Left            =   1845
      TabIndex        =   6
      Top             =   4995
      Width           =   1710
      Caption         =   "Save & Close"
      PicturePosition =   327683
      Size            =   "3016;635"
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
      Left            =   195
      TabIndex        =   5
      Top             =   4995
      Width           =   1590
      Caption         =   "Save & New    "
      PicturePosition =   327683
      Size            =   "2805;635"
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
      Left            =   3600
      TabIndex        =   7
      Top             =   4995
      Width           =   1380
      Caption         =   "Cancel    "
      PicturePosition =   327683
      Size            =   "2434;635"
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
      Caption         =   "Date :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   1275
      TabIndex        =   9
      Top             =   420
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Employee :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   930
      TabIndex        =   8
      Top             =   120
      Width           =   795
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   225
      X2              =   4925
      Y1              =   4950
      Y2              =   4950
   End
End
Attribute VB_Name = "frmTransferSTAdvance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_bContractors As Boolean
Public LongTerm As Boolean

Private Sub PrintReport()

'Screen.MousePointer = vbHourglass
'
'Dim f As New frmPrevRpt, rpt As New craxddrt.Report, ID As Integer
'
'Set rpt = rptApp.OpenReport
'ID = Val(con.Execute("Select MAX(EntryID) FROM Advances").Fields(0) & "")
'
'Dim FormulaFields As craxddrt.FormulaFieldDefinitions
'Dim FormulaField As craxddrt.FormulaFieldDefinition
'
'Set FormulaFields = rpt.FormulaFields
'For Each FormulaField In FormulaFields
'    If FormulaField.Name = "{@Company}" Then
'        FormulaField.Text = "'" & strCompany & "'"
'    ElseIf FormulaField.Name = "{@FromTo}" Then
'        FormulaField.Text = "'From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
'    End If
'Next
'
'Me.Hide
'f.ShowReport "{VAdvances.EntryID}=" & ID & " ", rpt
'
'
'Screen.MousePointer = vbDefault


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
    
End Sub

Private Sub cmbdept_matched()

    cmbEmp.ClearVals
    'cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.Text = "<All Departments>" Then
        If m_bContractors Then
            cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name", "VTempEmps", "EmpID", "Where Active=1 "
        Else
            cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name", "VSalariedEmps", "EmpID", "Where Active=1 "
        End If
    Else
        cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "' and Active=1"
    End If
    
End Sub

Private Function Saved() As Boolean

    Saved = False
    'On Error GoTo ERR
    Dim SQL As String
    
    If Not cmbEmp.MatchFound Then
        MsgBox "Please Select Valid Employee From The List.", vbInformation
        Exit Function
    ElseIf cmbEmpTo.MatchFound = False Then
        MsgBox "Please Select To Employee.", vbInformation
        Exit Function
    ElseIf Val(txtAmt) = 0 Then
        MsgBox "Invalid Amount.", vbInformation
        Exit Function
    ElseIf Val(txtDeduct) > Val(txtAmt) And txtDeduct.Visible Then
        MsgBox "Deduction Cannot Be More Than Loan Amount.", vbInformation
        Exit Function
    ElseIf Val(txtAmt) > Val(txtBal) Then
        MsgBox "Transfer Amount Can't be More than Balance.", vbInformation
        Exit Function
    End If
    
    Dim strVchrNo As String
    
    strVchrNo = "Transferred from " & cmbEmp.ID
    
    Call StartTrans(con)
    
    SQL = "INSERT INTO Advances(EmpID,DT,Type,Description," & _
     "Amount,DAmount,AmountCleared,AccVoucherNo,DeductionStartDT) VALUES('" & cmbEmpTo.ID & _
     "','" & DTPicker1 & "',1,'" & txtDesc & "'," & Val(txtAmt) & "," & Val(txtDeduct) & _
     ",0,'" & strVchrNo & "','" & DTDedStartMonth.Value & "')"
     
     
    con.Execute SQL
    
    con.Execute "INSERT INTO AmtCleared(EmpID,DT,AmtClrd,VchrNo) VALUES('" & cmbEmp.ID & "','" & DTPicker1 & _
     "'," & Val(txtAmt) & ",'" & strVchrNo & "')"
     
    con.CommitTrans
    
    Saved = True
    'If chkPrint = 1 Then Call PrintVoucher(MainForm.Cr1, strVchrNo)
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
    Amt = Val(txtAmt)
    
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
    theDesc = "Long Term Advance Paid To " & cmbEmp.Text & " Total Amount : " & Amt & " Deduction : " & Val(txtDeduct)
     
    ''''''''''''''''Charge To  Cash Entry'''''''''''''
    If chkOtherAccount.Value = vbChecked Then
        AccNo = cmbAccNo.ID
    Else
        AccNo = cmbCashInHand.List(cmbCashInHand.ListIndex, 1) 'CashInHand
    End If
    
    If bWaitForAuth = False Then
        
        Bal = getBalance(AccNo, Dated)
        Bal = Bal - Amt
        
        'Credit Entry
        con.Execute "INSERT INTO Vouchers VALUES(" & INextSno & ",'" & _
         Dated & "','" & VchrNo & "','" & AccNo & _
         "','" & theDesc & "',0," & Amt & "," & Bal & ",'',0)"
        
        Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
     
        INextSno = INextSno + 1
        If (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Banzai" Then
            AccNo = "15-004-12001"
        Else
            AccNo = "13-003-12001" 'IIf(LongTerm, LongAdvAccNo, ShortAdvAccNo)
        End If
        Bal = getBalance(AccNo, Dated)
      
        Amt = Amt
        Bal = Bal + Amt
       
        con.Execute "INSERT INTO Vouchers VALUES(" & INextSno & ",'" & _
         Dated & "','" & VchrNo & "','" & AccNo & _
         "','" & theDesc & "'," & Amt & ",0," & Bal & ",'',0)"
        
        Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
        con.Execute "Insert Into VoucherInfo(VchrNo,UserName,MachineName,DT) VALUES('" & VchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
    Else
        con.Execute "INSERT INTO VouchersForAuthorization(VDate,VchrNo,Accno,Description,Debit,Credit,balance,DpstSlip,CSNo) VALUES('" & Dated & "','" & _
             VchrNo & "','" & AccNo & "','" & theDesc & "',0," & Amt & ",0,'',0)"
        
        con.Execute "INSERT INTO VouchersForAuthorization(VDate,VchrNo,Accno,Description,Debit,Credit,balance,DpstSlip,CSNo) VALUES('" & Dated & "','" & _
             VchrNo & "','" & "13-003-12001" & "','" & theDesc & "'," & Amt & ",0,0,'',0)"
             
        con.Execute "INSERT INTO VouchersForAuthorizationInfo(VchrNo,UserName,MachineName,DT,PostedThroughJVForm,Notes,Payee) VALUES('" & VchrNo & "','" & _
         CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "',0,'','')"
    End If
    AddVoucher = VchrNo
    
End Function


Private Sub cmbEmp_matched()

    On Error GoTo err
    
    SQL = "(Select isnull(Sum(AmtClrd),0) balance From AmtCleared Where empid='" & cmbEmp.ID & "')"
        
    Balance = Val(con.Execute("Select Sum(Amount)-" & SQL & "  from Advances Where empid='" & cmbEmp.ID & "' and Type=1 ").Fields(0) & "")
    txtBal = Balance
    txtAmt = txtBal
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select EntryID,DAmount From Advances Where EntryID=(Select Max(EntryID) From Advances Where Type=1 AND EmpID='" & cmbEmp.ID & "')", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = True Then
            cmbEmp.Tag = "N.A"
        Else
            cmbEmp.Tag = !EntryID
            txtPrevDed = Val(!dAmount & "")
        End If
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If Saved Then
         'Unload Me
         cmbEmp.ID = ""
         txtAmt = ""
         txtDeduct = ""
         cmbEmp.SetFocus
    End If
End Sub

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub


Private Sub Form_Activate()

'    Label1(5).Visible = LongTerm
'    txtDeduct.Visible = LongTerm
   
    txtDesc = "Long Term Loan"
    Me.Caption = "Long Term Loan"
    DTDedStartMonth.Value = DateSerial(year(Date), Month(Date), 1)

End Sub

Private Sub Form_Load()

    cmbEmp.ListHeight = 1250
    
    'cmbType.AddItem "Short Term"
    'cmbType.AddItem "Long Term"
    'cmbType.ListIndex = 0
    m_bContractors = False
      
    DTPicker1 = GetServerDate
    chkPrint = Abs(bPrintWhenSaved)
    
    
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
    
'    If m_bContractors Then
'        strCond = "Where IsNull(TempDept,0)=1"
'    Else
'        strCond = "Where IsNull(TempDept,0)=0"
'    End If
    cmbEmp.ListHeight = 2400
    cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' - ' + FName + ' {' + Designation + '}'", "VSalariedEmps", "EmpID", ""
    cmbEmpTo.ListHeight = 2400
    cmbEmpTo.AddVals con, "'{' + EmpID + '} ' +  Name + ' - ' + FName + ' {' + Designation + '}'", "VSalariedEmps", "EmpID", "Where Active=1 "
    
    Me.Show 1
    
End Sub

