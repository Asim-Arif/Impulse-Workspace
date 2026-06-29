VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmEmpHoldSalary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hold Salary"
   ClientHeight    =   4605
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
   ScaleHeight     =   4605
   ScaleWidth      =   5100
   StartUpPosition =   2  'CenterScreen
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   315
      Left            =   1305
      TabIndex        =   0
      Top             =   75
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
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   315
      Left            =   1305
      TabIndex        =   1
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
   Begin MSComCtl2.DTPicker DTMonth 
      Height          =   315
      Left            =   1305
      TabIndex        =   12
      Top             =   765
      Width           =   3750
      _ExtentX        =   6615
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
      Format          =   108068867
      CurrentDate     =   40229
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Salary :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   8
      Left            =   705
      TabIndex        =   13
      Top             =   780
      Width           =   555
   End
   Begin MSForms.TextBox txtBal 
      Height          =   315
      Left            =   1305
      TabIndex        =   11
      Top             =   1155
      Width           =   1680
      VariousPropertyBits=   746604575
      ForeColor       =   12582912
      Size            =   "2963;556"
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
      Caption         =   "Description :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   6
      Left            =   360
      TabIndex        =   10
      Top             =   1935
      Width           =   900
   End
   Begin MSForms.TextBox txtDesc 
      Height          =   2070
      Left            =   1305
      TabIndex        =   3
      Top             =   1920
      Width           =   3720
      VariousPropertyBits=   -1467987941
      ScrollBars      =   3
      Size            =   "6562;3651"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   1305
      X2              =   4930
      Y1              =   4065
      Y2              =   4065
   End
   Begin MSForms.TextBox txtAmt 
      Height          =   315
      Left            =   1305
      TabIndex        =   2
      Top             =   1500
      Width           =   1680
      VariousPropertyBits=   746604571
      ForeColor       =   12582912
      Size            =   "2963;556"
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
      Caption         =   "Hold Amount :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Saalry Amount :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   105
      TabIndex        =   8
      Top             =   1200
      Width           =   1155
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   360
      Left            =   1335
      TabIndex        =   4
      Top             =   4155
      Width           =   1830
      Caption         =   "Save & Close"
      PicturePosition =   327683
      Size            =   "3228;635"
      Accelerator     =   83
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
      Left            =   3195
      TabIndex        =   5
      Top             =   4155
      Width           =   1830
      Caption         =   "Cancel    "
      PicturePosition =   327683
      Size            =   "3228;635"
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
      Caption         =   "Employee :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   465
      TabIndex        =   7
      Top             =   450
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Department :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   4
      Left            =   300
      TabIndex        =   6
      Top             =   120
      Width           =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   1305
      X2              =   4930
      Y1              =   4080
      Y2              =   4080
   End
End
Attribute VB_Name = "frmEmpHoldSalary"
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
    ElseIf Val(txtAmt) <= 0 Then
        MsgBox "Invalid Amount.", vbInformation
        Exit Function
    ElseIf Val(txtAmt) > Val(txtBal) Then
        MsgBox "Hold Amount Cannot Be More Than Salary.", vbInformation
        Exit Function
    End If
    
    'Dim strVchrNo As String
         
    Call StartTrans(con)
    
'    If Not m_bContractors Then
'        strVchrNo = AddVoucher
'    End If
    
    Dim myDT As Date
    myDT = DateSerial(DTMonth.year, DTMonth.Month, 1)
    
    con.Execute "INSERT INTO EmpHoldSalary(EmpID,DT,HoldAmt,Description,UserName,MachineName) VALUES('" & _
     cmbEmp.ID & "','" & myDT & "'," & Val(txtAmt) & ",'" & txtDesc & "','" & CurrentUserName & _
     "','" & strComputerName & "')"
     
    
    con.CommitTrans
    
    Saved = True
    'If chkPrint = 1 Then Call PrintReport
    
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
     
    INextSno = getNextSno(Dated)
    VchrNo = getNextVchrNo(Dated, "CPV")
     
    'Give The Description
    theDesc = "Long Term Advance Paid To " & cmbEmp.Text
     
    ''''''''''''''''Charge To  Cash Entry'''''''''''''
    
    AccNo = cmbCashInHand.List(cmbCashInHand.ListIndex, 1) 'CashInHand
    
    Bal = getBalance(AccNo, Dated)
    
    Bal = Bal - Amt
        
    'Credit Entry
    con.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & theDesc & "',0," & Amt & "," & Bal & ",'',0)"
           
    
    Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
     
     
     INextSno = INextSno + 1
     
    '''''''''''''''''' D E B I T   E N T R I E S '''''''''''''
      
      'Determine the Account no for Staff Member
      'Long / Short Tem Advance
      
      AccNo = "13-003-12001" 'IIf(LongTerm, LongAdvAccNo, ShortAdvAccNo)
     
      Bal = getBalance(AccNo, Dated)
      
      Amt = Amt
      Bal = Bal + Amt
    
    
      'Debit To Staff Long/Short Advance Account
      
      con.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & theDesc & "'," & Amt & ",0," & Bal & ",'',0)"
        
     
    Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
    con.Execute "Insert Into VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & VchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
        
    AddVoucher = VchrNo
    
End Function


Private Sub cmbEmp_matched()
    Call RefreshData
End Sub

Private Sub RefreshData()

    On Error GoTo err
    
    Dim myDT As Date
    myDT = DateSerial(DTMonth.year, DTMonth.Month + 1, 0)
    txtBal = GetSingleLongValue("Balance+ISNULL(HoldSalaryAmt,0)", "MonthlySalaries", " WHERE EmpID='" & cmbEmp.ID & "' AND DT='" & myDT & "'")
      
    myDT = DateSerial(DTMonth.year, DTMonth.Month, 1)
    txtAmt = GetSingleLongValue("HoldAmt", "EmpHoldSalary", " WHERE EmpID='" & cmbEmp.ID & "' AND DT='" & myDT & "'")
    
    If Val(txtAmt) > 0 Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If Saved Then
         Unload Me
    End If
End Sub

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub


Private Sub DTMonth_Change()
    RefreshData
End Sub

Private Sub Form_Activate()

'    Label1(5).Visible = LongTerm
'    txtDeduct.Visible = LongTerm
 
    DTMonth.Value = DateSerial(year(Date), Month(Date), 1)

End Sub

Private Sub Form_Load()

    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    
    'cmbType.AddItem "Short Term"
    'cmbType.AddItem "Long Term"
    'cmbType.ListIndex = 0
    m_bContractors = False
      
    DTPicker1 = GetServerDate
    chkPrint = Abs(bPrintWhenSaved)
    
    With cmbDept
        .ClearVals
        .AddItem "<All Departments>", "0"
        .AddVals con, "Name", "Departments", "DeptID"
        .Text = "<All Departments>"
    End With
    
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
