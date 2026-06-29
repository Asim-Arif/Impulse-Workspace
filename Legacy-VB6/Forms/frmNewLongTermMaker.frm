VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmNewLongTermMaker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Long Term Advance For Makers"
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
      Top             =   3795
      Visible         =   0   'False
      Width           =   1800
   End
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   315
      Left            =   1755
      TabIndex        =   0
      Top             =   105
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
      Top             =   525
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
      Format          =   108199939
      CurrentDate     =   37384
   End
   Begin MSComCtl2.DTPicker DTDedStartMonth 
      Height          =   315
      Left            =   1800
      TabIndex        =   18
      Top             =   3390
      Visible         =   0   'False
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
      Format          =   108199939
      CurrentDate     =   40229
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ded. Start Month :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   8
      Left            =   420
      TabIndex        =   19
      Top             =   3405
      Visible         =   0   'False
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
      Top             =   1380
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
      Top             =   1395
      Width           =   810
   End
   Begin MSForms.TextBox txtBal 
      Height          =   315
      Left            =   1770
      TabIndex        =   15
      Top             =   945
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
      Top             =   1335
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
      Top             =   1785
      Width           =   1575
   End
   Begin MSForms.TextBox txtDesc 
      Height          =   1005
      Left            =   1770
      TabIndex        =   4
      Top             =   1770
      Width           =   3090
      VariousPropertyBits=   -1467987941
      ScrollBars      =   3
      Size            =   "5450;1773"
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
      Y1              =   4095
      Y2              =   4095
   End
   Begin MSForms.TextBox txtAmt 
      Height          =   315
      Left            =   3810
      TabIndex        =   2
      Top             =   945
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
      Top             =   1335
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
      Top             =   1005
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
      Top             =   1005
      Width           =   1635
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   360
      Left            =   1845
      TabIndex        =   6
      Top             =   4155
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
      Top             =   4155
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
      Top             =   4155
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
      TabIndex        =   9
      Top             =   540
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Maker :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   8
      Top             =   150
      Width           =   1020
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   225
      X2              =   4925
      Y1              =   4110
      Y2              =   4110
   End
End
Attribute VB_Name = "frmNewLongTermMaker"
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


Private Function Saved() As Boolean

    Saved = False
    'On Error GoTo ERR
    Dim SQL As String
    
    If Not cmbEmp.MatchFound Then
        MsgBox "Please Select Valid Maker From The List.", vbInformation
        Exit Function
    ElseIf Val(txtAmt) = 0 Then
        MsgBox "Invalid Amount.", vbInformation
        Exit Function
'    ElseIf Val(txtDeduct) > Val(txtAmt) And txtDeduct.Visible Then
'        MsgBox "Deduction Cannot Be More Than Loan Amount.", vbInformation
'        Exit Function
    End If
    
    Dim strVchrNo As String
         
    Call StartTrans(con)
    
    
    'strVchrNo = AddVoucher
   
    
    SQL = "INSERT INTO MakerAdvances(MakerID,DT,Type,Description," & _
     "Amount,DAmount,AmtCleared,AccVoucherNo,DeductionStartDT) VALUES('" & cmbEmp.ID & _
     "','" & DTPicker1 & "',1,'" & txtDesc & "'," & Val(txtAmt) & "," & Val(txtDeduct) & _
     ",0,'" & strVchrNo & "','" & DTDedStartMonth.Value & "')"
     
    con.Execute SQL
    
    con.CommitTrans
    
    Saved = True
    If chkPrint = 1 Then Call PrintReport
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

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "SELECT * FROM VMakerLongTermDeductions WHERE VendID=" & cmbEmp.ID, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            txtBal = Val(!TotalAmountTaken & "") - Val(!TotalAmountCleared & "")
           txtPrevDed = Val(!DeductionAmount & "")
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

Public Sub ShowMe()
     
    cmbEmp.AddVals con, "'{' + VendID1 + '} ' +  VenderName", "Makers", "VendID", "WHERE Active=1 "
    Me.Show 1
    
End Sub

