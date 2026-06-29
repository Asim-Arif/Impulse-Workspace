VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmNewShortTermMaker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Short Term Advance For Makers"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
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
   ScaleHeight     =   4035
   ScaleWidth      =   5010
   StartUpPosition =   2  'CenterScreen
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   315
      Left            =   1755
      TabIndex        =   0
      Top             =   90
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
   Begin VB.Frame Frame1 
      Caption         =   "Deductions :"
      Height          =   1155
      Left            =   75
      TabIndex        =   10
      Top             =   1095
      Width           =   4800
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fixed Advance Inst. : "
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   9
         Left            =   2085
         TabIndex        =   14
         Top             =   315
         Width           =   1605
      End
      Begin MSForms.TextBox txtLongTerm 
         Height          =   315
         Left            =   3690
         TabIndex        =   13
         Top             =   270
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
         Caption         =   "Previouse Short Term :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   14
         Left            =   2025
         TabIndex        =   12
         Top             =   675
         Width           =   1650
      End
      Begin MSForms.TextBox txtPrevST 
         Height          =   315
         Left            =   3690
         TabIndex        =   11
         Top             =   645
         Width           =   1050
         VariousPropertyBits=   746604575
         Size            =   "1852;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
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
      Left            =   120
      TabIndex        =   8
      Top             =   3045
      Visible         =   0   'False
      Width           =   1800
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1785
      TabIndex        =   1
      Top             =   495
      Width           =   3075
      _ExtentX        =   5424
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
      Format          =   224591875
      CurrentDate     =   37384
   End
   Begin MSForms.TextBox txtBal 
      Height          =   315
      Left            =   3810
      TabIndex        =   9
      Top             =   2490
      Width           =   1050
      VariousPropertyBits=   746604571
      Size            =   "1852;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   120
      X2              =   4820
      Y1              =   3435
      Y2              =   3435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Amount Paid :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   2760
      TabIndex        =   7
      Top             =   2550
      Width           =   1005
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   360
      Left            =   1680
      TabIndex        =   3
      Top             =   3495
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
      Left            =   60
      TabIndex        =   2
      Top             =   3495
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
      Left            =   3300
      TabIndex        =   4
      Top             =   3495
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
      Caption         =   "Issue Date                 :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   105
      TabIndex        =   6
      Top             =   510
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Maker :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   135
      Width           =   1020
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   4820
      Y1              =   3450
      Y2              =   3450
   End
End
Attribute VB_Name = "frmNewShortTermMaker"
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
    
    con.Execute "DELETE FROM PrintTempAdv"
    
    con.Execute "INSERT INTO PrintTempAdv(EmpID,BSal,Rate,SDays,SAmt,Total,Paid,LongTerm,Balance,DT,LunchAmt,AmtPaid,AdvancesID) Values('" & _
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

Private Function Saved() As Boolean

    Saved = False
    'On Error GoTo ERR
    
'    If Not m_bContractors Then
'        If Month(DTPicker1) <> Month(GetServerDate) Then
'            MsgBox "Invalid Date.", vbInformation
'            Exit Function
'        ElseIf year(DTPicker1) <> year(GetServerDate) Then
'            MsgBox "Invalid Date.", vbInformation
'            Exit Function
'        End If
'    End If
    
    Dim SQL As String
    
    If Not cmbEmp.MatchFound Then
        MsgBox "Please Select Valid Maker From The List.", vbInformation
        Exit Function
    ElseIf Val(txtBal) = 0 Then
        MsgBox "Invalid Amount.", vbInformation
        Exit Function
    End If
    
    Dim strVchrNo As String
    Call StartTrans(con)
    'strVchrNo = AddVoucher
    
    SQL = "INSERT INTO MakerAdvances(MakerID,DT,Type,Description," & _
     "Amount,DAmount,AmtCleared,AccVoucherNo) VALUES(" & cmbEmp.ID & _
     ",'" & DTPicker1 & "',0,'Short Term'," & Val(txtBal) & ",0,0,'" & _
     strVchrNo & "')"
    
    con.Execute SQL
    
    con.CommitTrans
    
    Saved = True
    'If chkPrint = 1 Then Call PrintReport(True)
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
     
    INextSno = getNextSno(Dated)
    VchrNo = getNextVchrNo(Dated, "CPV")
     
    'Give The Description
    theDesc = "Temp. Advance Paid To " & cmbEmp.Text
     
    ''''''''''''''''Charge To Cash Entry'''''''''''''
    
    AccNo = cmbCashInHand.List(cmbCashInHand.ListIndex, 1) 'CashInHand
   
    Bal = getBalance(AccNo, Dated)
    
    Bal = Bal - Amt
        
    'Credit Entry
    con.Execute "INSERT INTO Vouchers VALUES(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & theDesc & "',0," & Amt & "," & Bal & ",'',0)"
    
    Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
     
     
    INextSno = INextSno + 1
     
    '''''''''''''''''' D E B I T   E N T R I E S '''''''''''''
      
      'Determine the Account no for Staff Member
      'Long / Short Tem Advance
      
      AccNo = "13-003-19001" 'IIf(LongTerm, LongAdvAccNo, ShortAdvAccNo)
     
      Bal = getBalance(AccNo, Dated)
      
      Amt = Amt
      Bal = Bal + Amt
    
    
      'Debit To Staff Long/Short Advance Account
      
      con.Execute "INSERT INTO Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & theDesc & "'," & Amt & ",0," & Bal & ",'',0)"
        
     
    Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
    
    con.Execute "INSERT INTO VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & VchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
    
    AddVoucher = VchrNo
    
End Function

Private Sub cmbEmp_matched()

    On Error GoTo err
    txtLongTerm = GetSingleLongValue("ActualDeductionAmount", "VMakerLongTermDeductions", " WHERE VendID=" & cmbEmp.ID)
    txtPrevST = GetSingleLongValue("SUM(Amount)-SUM(AmtCleared)", "MakerAdvances", " WHERE MakerID=" & cmbEmp.ID & " AND Type=0")
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()

    If Saved Then
        txtBal = ""
        cmbEmp.ID = ""
        cmbEmp.SetFocus
        'cmbDept.SetFocus
    End If
    
End Sub

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub

Private Sub DTPicker1_Change()
    If cmbEmp.MatchFound Then
        Call cmbEmp_matched
    End If
End Sub

Private Sub Form_Load()

    cmbEmp.ListHeight = 1250
      
    DTPicker1 = GetServerDate
    m_bContractors = False
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

Public Sub ShowMe()

    cmbEmp.AddVals con, "'{' + VendID1 + '} ' +  VenderName", "Makers", "VendID", "WHERE Active=1 "
    Me.Show 1
    
End Sub

