VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form NewAdvanceLong 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Long Term Loan..."
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10305
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
   ScaleWidth      =   10305
   StartUpPosition =   2  'CenterScreen
   Begin ComboList.Usercontrol1 cmbAccNo 
      Height          =   285
      Left            =   1755
      TabIndex        =   24
      Top             =   2895
      Visible         =   0   'False
      Width           =   6630
      _ExtentX        =   11695
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
      TextBackColor   =   12640511
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
      Appearance      =   0
   End
   Begin VB.CheckBox chkPost 
      Caption         =   "Post in Financials"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6429
      TabIndex        =   28
      Top             =   3195
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1956
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
   Begin VB.CheckBox chkOtherAccount 
      Caption         =   "Show All Accounts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1770
      TabIndex        =   25
      Top             =   3195
      Width           =   2235
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
      Left            =   1770
      TabIndex        =   15
      Top             =   5055
      Width           =   1800
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1755
      TabIndex        =   2
      Top             =   765
      Width           =   1560
      _ExtentX        =   2752
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
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   118423555
      CurrentDate     =   37384
   End
   Begin MSComCtl2.DTPicker DTDedStartMonth 
      Height          =   315
      Left            =   6465
      TabIndex        =   22
      Top             =   765
      Width           =   1920
      _ExtentX        =   3387
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
      Format          =   125763587
      CurrentDate     =   40229
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1755
      Left            =   8460
      Stretch         =   -1  'True
      Top             =   30
      Width           =   1785
   End
   Begin MSForms.TextBox txtVchrNo 
      Height          =   315
      Left            =   3945
      TabIndex        =   27
      Top             =   750
      Width           =   1455
      VariousPropertyBits=   746604575
      Size            =   "2566;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Vchr # :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   9
      Left            =   3345
      TabIndex        =   26
      Top             =   810
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ded. Month :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   8
      Left            =   5490
      TabIndex        =   23
      Top             =   795
      Width           =   945
   End
   Begin MSForms.ComboBox cmbCashInHand 
      Height          =   285
      Left            =   1755
      TabIndex        =   21
      Top             =   2895
      Width           =   6630
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "11695;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Cash In Hand :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   300
      TabIndex        =   20
      Top             =   2910
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Deduction Amount :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   7
      Left            =   300
      TabIndex        =   19
      Top             =   1485
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "New Ded. :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   3
      Left            =   4905
      TabIndex        =   18
      Top             =   1485
      Width           =   810
   End
   Begin MSForms.TextBox txtBal 
      Height          =   315
      Left            =   1755
      TabIndex        =   17
      Top             =   1095
      Width           =   2565
      VariousPropertyBits=   746604575
      Size            =   "4524;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox txtPrevDed 
      Height          =   315
      Left            =   1755
      TabIndex        =   16
      Top             =   1455
      Width           =   2565
      VariousPropertyBits=   746604575
      Size            =   "4524;556"
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
      TabIndex        =   14
      Top             =   1860
      Width           =   1575
   End
   Begin MSForms.TextBox txtDesc 
      Height          =   1005
      Left            =   1755
      TabIndex        =   5
      Top             =   1845
      Width           =   6630
      VariousPropertyBits=   -1467987941
      ScrollBars      =   3
      Size            =   "11695;1773"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   3630
      X2              =   8330
      Y1              =   4935
      Y2              =   4935
   End
   Begin MSForms.TextBox txtAmt 
      Height          =   315
      Left            =   5745
      TabIndex        =   3
      Top             =   1095
      Width           =   2640
      VariousPropertyBits=   746604571
      BackColor       =   16761024
      ForeColor       =   255
      Size            =   "4657;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox txtDeduct 
      Height          =   315
      Left            =   5745
      TabIndex        =   4
      Top             =   1455
      Width           =   2640
      VariousPropertyBits=   746604571
      Size            =   "4657;556"
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
      Left            =   4875
      TabIndex        =   13
      Top             =   1110
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Previous Balance        :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   12
      Top             =   1125
      Width           =   1635
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   360
      Left            =   5250
      TabIndex        =   7
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
      Left            =   3600
      TabIndex        =   6
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
      Left            =   7005
      TabIndex        =   8
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
      Caption         =   "Issue Date                 :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   810
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Employee        :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   10
      Top             =   420
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Department    :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   4
      Left            =   150
      TabIndex        =   9
      Top             =   90
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   3630
      X2              =   8330
      Y1              =   4950
      Y2              =   4950
   End
End
Attribute VB_Name = "NewAdvanceLong"
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
    'On Error GoTo ERR
    Dim SQL As String
    
    If Not cmbEmp.MatchFound Then
        MsgBox "Please Select Valid Employee From The List.", vbInformation
        Exit Function
    ElseIf Val(txtAmt) = 0 Then
        MsgBox "Invalid Amount.", vbInformation
        Exit Function
    ElseIf chkOtherAccount.Value = vbUnchecked And cmbCashInHand.MatchFound = False And chkPost.Value = vbChecked Then
        MsgBox "Please Select Cash In Hand.", vbInformation
        Exit Function
    ElseIf chkPost.Value = vbChecked And chkOtherAccount.Value = vbChecked And cmbAccNo.MatchFound = False Then
        MsgBox "Please Select Account.", vbInformation
        Exit Function
    ElseIf chkPost.Value = vbChecked And strEmpLongTermAccNo = "" Then
        MsgBox "Can't Post, Emp. Long Term Loan Account Not Found.", vbInformation
        Exit Function
    'ElseIf Val(txtDeduct) = 0 And txtDeduct.Visible Then
    '    MsgBox "Invalid Deduction Amount.", vbInformation
    '    Exit Function
        
    ElseIf Val(txtDeduct) > Val(txtAmt) And txtDeduct.Visible Then
        MsgBox "Deduction Cannot Be More Than Loan Amount.", vbInformation
        Exit Function
    End If
    
    If chkPost.Value = vbChecked And cmbCashInHand.Visible Then
        Dim dCashBalance As Double
        dCashBalance = getBalance(cmbCashInHand.List(cmbCashInHand.ListIndex, 1), DTPicker1)
        
        If Val(txtAmt) > dCashBalance Then
            MsgBox "Invalid Balance in Cash In Hand.", vbInformation
            Exit Function
        End If
        
    End If
    '''''''''''''''''''''''''''''''''
    ' Remove This Check at the moment
    
    
'    If LongTerm = True Then
'        If IsNotLongApplicable Then
'            Exit Function
'        End If
'    End If
'
    
    
    '''''''''''''''''''''''''''''''''
    ' Remove This Check at the moment
    
'    Dim AmtAllowed As Double
'    AmtAllowed = GetShortAmt(cmbEmp.ID, DTPicker1.value)
'    If Val(txtAmt) > AmtAllowed Then
'      MsgBox "Cannot Pay More Than Rs. " & Format(AmtAllowed, "0.00"), vbInformation
'      Exit Function
'    End If
'
    Dim strVchrNo As String
         
    Call StartTrans(con)
    
    If chkPost.Value = vbChecked Then
        strVchrNo = AddVoucher
    End If
    
    SQL = "INSERT INTO Advances(EmpID,DT,Type,Description," & _
     "Amount,DAmount,AmountCleared,AccVoucherNo,DeductionStartDT) VALUES('" & cmbEmp.ID & _
     "','" & DTPicker1 & "',1,'" & txtDesc & "'," & Val(txtAmt) & "," & Val(txtDeduct) & _
     ",0,'" & strVchrNo & "','" & DTDedStartMonth.Value & "')"
     
    con.Execute SQL
    
    con.CommitTrans
        
    MsgBox "Saved Successfully.", vbInformation
    
    Saved = True
    If chkPrint = 1 Then Call PrintVoucher(MainForm.Cr1, strVchrNo)
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
'        If (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Banzai" Then
'            AccNo = "15-004-12001"
'        Else
'            AccNo = "13-003-12001" 'IIf(LongTerm, LongAdvAccNo, ShortAdvAccNo)
'        End If
        AccNo = strEmpLongTermAccNo
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
             VchrNo & "','" & strEmpLongTermAccNo & "','" & theDesc & "'," & Amt & ",0,0,'',0)"
             
        con.Execute "INSERT INTO VouchersForAuthorizationInfo(VchrNo,UserName,MachineName,DT,PostedThroughJVForm,Notes,Payee) VALUES('" & VchrNo & "','" & _
         CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "',0,'','')"
    End If
    AddVoucher = VchrNo
    
    If VchrNo <> txtVchrNo Then
        MsgBox "Vchr No. " & txtVchrNo & " had been used, " & VchrNo & " is used Instead.", vbInformation
    End If
    
End Function


Private Sub cmbEmp_matched()

    On Error GoTo err
    
    SQL = "(SELECT ISNULL(SUM(AmtClrd),0) Balance FROM AmtCleared WHERE EmpID='" & cmbEmp.ID & "')"
        
    Balance = Val(con.Execute("SELECT SUM(Amount)-" & SQL & "  FROM Advances WHERE EmpID='" & cmbEmp.ID & "' AND Type=1 ").Fields(0) & "")
    txtBal = Balance
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "SELECT EntryID,DAmount FROM Advances WHERE EntryID=(SELECT MAX(EntryID) FROM Advances WHERE Type=1 AND EmpID='" & cmbEmp.ID & "')", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = True Then
            cmbEmp.Tag = "N.A"
        Else
            cmbEmp.Tag = !EntryID
            txtPrevDed = Val(!dAmount & "")
        End If
        .Close
        
        .Open "SELECT EmpID,EmpPic FROM Employees WHERE EmpID='" & cmbEmp.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        If Not IsNull(![EmpPic]) Then
            Set Image1.DataSource = rs
            Image1.DataField = "EmpPic"
        Else
            Image1.Picture = LoadPicture("")
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
         Call DTPicker1_Change
         cmbEmp.ID = ""
         txtAmt = ""
         txtDeduct = ""
         cmbEmp.SetFocus
    End If
End Sub

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub


Private Sub DTPicker1_Change()

    Dim strVchrType As String
    If chkOtherAccount.Value = vbChecked Then
        strVchrType = "JV"
    Else
        strVchrType = "CPV"
    End If
    txtVchrNo = getNextVchrNo(DTPicker1.Value, strVchrType)
    
End Sub

Private Sub Form_Activate()

'    Label1(5).Visible = LongTerm
'    txtDeduct.Visible = LongTerm
   
    txtDesc = "Long Term Loan"
    Me.Caption = "Long Term Loan"
    DTDedStartMonth.Value = DateSerial(year(Date), Month(Date), 1)

End Sub

Private Sub Form_Load()

    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    
    'cmbType.AddItem "Short Term"
    'cmbType.AddItem "Long Term"
    'cmbType.ListIndex = 0
    m_bContractors = False
      
    DTPicker1 = GetServerDate
    Call DTPicker1_Change
    
    chkPrint = Abs(bPrintWhenSaved)
    
    cmbAccNo.ListHeight = 1600
    cmbAccNo.AddVals con, "'{' + AccNo  + '} ' + AccTitle", "VActiveAccounts", "AccNo"
    
    
    If strCompany = "Tecno" Or strCompany = "IAA" Or strCompany = "Kami" Then
        chkPost.Visible = True
        If strCompany = "IAA" Or strCompany = "Kami" Then
            chkPost.Value = vbUnchecked
        End If
    Else
        chkPost.Visible = False
    End If
    
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
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT AccNo,AccTitle FROM VActiveAccounts WHERE SubAccOf='" & CashInHand & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCashInHand.Clear
        Do Until .EOF
            cmbCashInHand.AddItem !AccTitle & ""
            cmbCashInHand.List(.AbsolutePosition - 1, 1) = !AccNo & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    If cmbCashInHand.ListCount > 0 Then cmbCashInHand.ListIndex = cmbCashInHand.ListCount - 1
   
'    If strCompany = "Tecno" Then
'        Label2.Visible = False
'        chkOtherAccount.Visible = False
'        chkPrint.Value = vbUnchecked
'        chkPrint.Visible = False
'        cmbCashInHand.Visible = False
'        cmbAccNo.Visible = False
'    End If
    Me.Show 1
    
End Sub

