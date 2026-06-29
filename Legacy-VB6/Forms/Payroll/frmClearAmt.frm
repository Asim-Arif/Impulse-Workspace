VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmClearAmt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clear long term loan"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9180
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
   ScaleHeight     =   4470
   ScaleWidth      =   9180
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDontPostVoucher 
      Caption         =   "Don't Post Voucher"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   5370
      TabIndex        =   26
      Top             =   3300
      Width           =   1950
   End
   Begin ComboList.Usercontrol1 cmbAccNo 
      Height          =   285
      Left            =   1770
      TabIndex        =   22
      Top             =   2970
      Visible         =   0   'False
      Width           =   5520
      _ExtentX        =   9737
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
      TabIndex        =   23
      Top             =   3270
      Width           =   2235
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   315
      Left            =   1770
      TabIndex        =   0
      Top             =   30
      Width           =   5520
      _ExtentX        =   9737
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
      Left            =   1770
      TabIndex        =   1
      Top             =   375
      Width           =   5520
      _ExtentX        =   9737
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
      Height          =   345
      Left            =   480
      TabIndex        =   15
      Top             =   4080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1770
      TabIndex        =   2
      Top             =   765
      Width           =   5520
      _ExtentX        =   9737
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
      Format          =   157155331
      CurrentDate     =   37384
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1755
      Left            =   7320
      Stretch         =   -1  'True
      Top             =   30
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Vchr # :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   9
      Left            =   3030
      TabIndex        =   25
      Top             =   1200
      Width           =   585
   End
   Begin MSForms.TextBox txtVchrNo 
      Height          =   315
      Left            =   3630
      TabIndex        =   24
      Top             =   1155
      Width           =   1455
      VariousPropertyBits=   746604575
      Size            =   "2566;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
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
      TabIndex        =   21
      Top             =   2985
      Width           =   1425
   End
   Begin MSForms.ComboBox cmbCashInHand 
      Height          =   300
      Left            =   1770
      TabIndex        =   20
      Top             =   2970
      Width           =   5520
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "9737;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Deduction Amount    :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   7
      Left            =   150
      TabIndex        =   19
      Top             =   1575
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "New Ded. :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   3
      Left            =   5355
      TabIndex        =   18
      Top             =   1590
      Width           =   810
   End
   Begin MSForms.TextBox txtNewDed 
      Height          =   315
      Left            =   6240
      TabIndex        =   17
      Top             =   1530
      Width           =   1050
      VariousPropertyBits=   746604571
      Size            =   "1852;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox txtDed 
      Height          =   315
      Left            =   1770
      TabIndex        =   16
      Top             =   1530
      Width           =   1050
      VariousPropertyBits=   746604573
      Size            =   "1852;556"
      FontName        =   "Tahoma"
      FontEffects     =   1073750016
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
      Top             =   1920
      Width           =   1575
   End
   Begin MSForms.TextBox txtDesc 
      Height          =   1005
      Left            =   1770
      TabIndex        =   13
      Top             =   1905
      Width           =   5520
      VariousPropertyBits=   -1467987941
      ScrollBars      =   3
      Size            =   "9737;1773"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtAmt 
      Height          =   315
      Left            =   1770
      TabIndex        =   12
      Top             =   1155
      Width           =   1050
      VariousPropertyBits=   746604573
      Size            =   "1852;556"
      FontName        =   "Tahoma"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox txtAmtRec 
      Height          =   315
      Left            =   6240
      TabIndex        =   11
      Top             =   1155
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
      Caption         =   "Amt Rec. :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   5
      Left            =   5340
      TabIndex        =   10
      Top             =   1215
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Amount Remaing       :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   9
      Top             =   1200
      Width           =   1590
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   360
      Left            =   4110
      TabIndex        =   4
      Top             =   4065
      Width           =   1530
      Caption         =   "Save & Close"
      PicturePosition =   327683
      Size            =   "2699;635"
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
      Left            =   2520
      TabIndex        =   3
      Top             =   4065
      Width           =   1530
      Caption         =   "Save & New    "
      PicturePosition =   327683
      Size            =   "2699;635"
      Picture         =   "frmClearAmt.frx":0000
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
      Left            =   5700
      TabIndex        =   5
      Top             =   4065
      Width           =   1530
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "2699;635"
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
      Caption         =   "Clear Date                 :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   8
      Top             =   810
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Employee        :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   135
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
      Left            =   150
      TabIndex        =   6
      Top             =   90
      Width           =   1575
   End
End
Attribute VB_Name = "frmClearAmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LongTerm As Boolean
'Dim con As New ADODB.Connection

Private Sub PrintReport()

    Screen.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report, ID As Integer
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptClearLongTerm.rpt")
    'ID = Val(con.Execute("Select MAX(EntryID) FROM Advances").Fields(0) & "")
    
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
    
    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & strCompany & "'"
        ElseIf FormulaField.Name = "{@Amount}" Then
            FormulaField.Text = txtAmtRec
        ElseIf FormulaField.Name = "{@Balance}" Then
            FormulaField.Text = Val(txtAmt) - Val(txtAmtRec)
        End If
    Next
    
    Me.Hide
    f.ShowReport "{VEmp.EmpID}='" & cmbEmp.ID & "' ", rpt
    
    
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
        cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' - ' + FName + ' {' + Designation + '}'", "Employees", "EmpID"
    Else
        cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' - ' + FName + ' {' + Designation + '}'", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "'"
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
        MsgBox "No Amount To Receive !", vbInformation
        Exit Function
    ElseIf Val(txtAmtRec) = 0 Then
        MsgBox "Receiving Amount Could Not be Zero !", vbInformation
        Exit Function
    ElseIf Val(txtAmtRec) > Val(txtAmt) Then
        MsgBox "Receive Amount Could Not Be Greater Than Total Amount.", vbInformation
        Exit Function
    ElseIf chkDontPostVoucher.Value = vbUnchecked And chkOtherAccount.Value = vbUnchecked And cmbCashInHand.MatchFound = False Then
        MsgBox "Please Select Cash In Hand.", vbInformation
        Exit Function
    ElseIf chkDontPostVoucher.Value = vbUnchecked And chkOtherAccount.Value = vbChecked And cmbAccNo.MatchFound = False Then
        MsgBox "Please Select Account.", vbInformation
        Exit Function
    End If
    
    Dim strVchrNo As String
         
    Call StartTrans(con)
    
    'If Not m_bContractors Then
'    If strCompany <> "Tecno" Then
'        strVchrNo = AddVoucher
'    End If
    'End If
    If chkDontPostVoucher.Value = vbUnchecked Then
        strVchrNo = AddVoucher
    End If
    
    con.Execute "INSERT INTO AmtCleared(EmpID,DT,AmtClrd,VchrNo) VALUES('" & cmbEmp.ID & "','" & DTPicker1 & "'," & Val(txtAmtRec) & ",'" & strVchrNo & "')"
    
    If cmbEmp.Tag <> "N.A" And Val(txtNewDed) > 0 Then
        con.Execute "UPDATE Advances SET DAmount=" & Val(txtNewDed) & " WHERE EntryID=" & cmbEmp.Tag
    End If
    
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
    Amt = Val(txtAmtRec)
    
    If Amt = 0 Then Exit Function
    
    'ConAcc.BeginTrans
     
    'Get The Required Values
    Dim strVchrType As String
    If chkOtherAccount.Value = vbChecked Then
        strVchrType = "JV"
    Else
        strVchrType = "CRV"
    End If
    
    Dim bWaitForAuth As Boolean
    bWaitForAuth = VoucherRequiresAuthorization(strVchrType)
    
    INextSno = getNextSno(Dated)
    VchrNo = getNextVchrNo(Dated, strVchrType)
     
    'Give The Description
    theDesc = "Permanent Advance Received From [ " & cmbEmp.ID & " ] " & cmbEmp.Text
     
    ''''''''''''''''Charge To  Cash Entry'''''''''''''
    
    Dim strDebitAccNo As String
    If chkOtherAccount.Value = vbChecked Then
        strDebitAccNo = cmbAccNo.ID
    Else
        strDebitAccNo = cmbCashInHand.List(cmbCashInHand.ListIndex, 1) 'CashInHand
    End If
    
    If bWaitForAuth = False Then
'        If (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Banzai" Or strCompany = "IAA" OR strCompany = "Kami" Then
'            AccNo = "15-004-12001"
'        Else
'            AccNo = "13-003-12001" 'LongTerm
'        End If
        AccNo = strEmpLongTermAccNo
        Bal = getBalance(AccNo, Dated)
        Bal = Bal - Amt
        'Credit Entry
        con.Execute "INSERT INTO Vouchers VALUES(" & INextSno & ",'" & _
         Dated & "','" & VchrNo & "','" & AccNo & _
         "','" & theDesc & "',0," & Amt & "," & Bal & ",'',0)"
    
        Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
        INextSno = INextSno + 1
     
        '''''''''''''''''' D E B I T   E N T R I E S '''''''''''''
      
        'AccNo = cmbCashInHand.List(cmbCashInHand.ListIndex, 1) 'CashInHand 'IIf(LongTerm, LongAdvAccNo, ShortAdvAccNo)
        
        
        Bal = getBalance(strDebitAccNo, Dated)
      
        Amt = Amt
        Bal = Bal + Amt
      
        con.Execute "INSERT INTO Vouchers VALUES(" & INextSno & ",'" & _
         Dated & "','" & VchrNo & "','" & strDebitAccNo & _
         "','" & theDesc & "'," & Amt & ",0," & Bal & ",'',0)"
        
        Call adjustLedger(strDebitAccNo, Dated, INextSno, Val(Amt))
    
        con.Execute "INSERT INTO VoucherInfo(VchrNo,UserName,MachineName,DT) VALUES('" & VchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
    Else
        con.Execute "INSERT INTO VouchersForAuthorization(VDate,VchrNo,Accno,Description,Debit,Credit,balance,DpstSlip,CSNo) VALUES('" & Dated & "','" & _
             VchrNo & "','" & strDebitAccNo & "','" & theDesc & "'," & Amt & ",0,0,'',0)"
             
        con.Execute "INSERT INTO VouchersForAuthorization(VDate,VchrNo,Accno,Description,Debit,Credit,balance,DpstSlip,CSNo) VALUES('" & Dated & "','" & _
             VchrNo & "','" & strEmpLongTermAccNo & "','" & theDesc & "',0," & Amt & ",0,'',0)"
             
        con.Execute "INSERT INTO VouchersForAuthorizationInfo(VchrNo,UserName,MachineName,DT,PostedThroughJVForm,Notes,Payee) VALUES('" & VchrNo & "','" & _
         CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "',0,'','')"
    End If
    AddVoucher = VchrNo
        
    If VchrNo <> txtVchrNo Then
        MsgBox "Vchr No. " & txtVchrNo & " had been used, " & VchrNo & " is used Instead.", vbInformation
    End If
    
End Function


Private Sub cmbType_Change()

End Sub

Private Sub cmbEmp_matched()

    On Error GoTo err
    'and  Dt<'" & Format(DTPicker1, "dd-MMM-yyyy") & "'
    
    SQL = "(Select isnull(Sum(AmtClrd),0) balance From AmtCleared Where empid='" & cmbEmp.ID & "')"
        
    Balance = Val(con.Execute("Select Sum(Amount)-" & SQL & "  from Advances Where empid='" & cmbEmp.ID & "' and Type=1 ").Fields(0) & "")
    txtAmt = Balance
    
    
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select EntryID,DAmount From Advances Where EntryID=(Select Max(EntryID) From Advances Where Type=1 AND EmpID='" & cmbEmp.ID & "')", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = True Then
            cmbEmp.Tag = "N.A"
        Else
            cmbEmp.Tag = !EntryID
            txtDed = Val(!dAmount & "")
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
         Unload Me
         frmClearAmt.Show 1
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
        strVchrType = "CRV"
    End If
    txtVchrNo = getNextVchrNo(DTPicker1.Value, strVchrType)
    
End Sub

Private Sub Form_Load()
    
'    con.CursorLocation = adUseClient

    
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    
    'cmbType.AddItem "Short Term"
    'cmbType.AddItem "Long Term"
    'cmbType.ListIndex = 0
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID"
    cmbDept.Text = "<All Departments>"
      
    DTPicker1 = GetServerDate
    Call DTPicker1_Change
    
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
     
    cmbAccNo.ListHeight = 1600
    cmbAccNo.AddVals con, "'{' + AccNo  + '} ' + AccTitle", "VActiveAccounts", "AccNo"
    
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
    .Open "SELECT * FROM VOnlyEmps WHERE EmpID='" & cmbEmp.ID & "'", con, adOpenKeyset, adLockReadOnly
    
    If DateDiff("m", !JoinDate, DTPicker1) < !longtermtime Then
        MsgBox "Could Not Give Long Term Loan To Employee." & vbNewLine & "It Can Be Given After " & !longtermtime & " Months", vbInformation
        IsNotLongApplicable = True
    
    ElseIf Val(txtAmt) > !LongTermLimit Then
        MsgBox "Could Not Give Long Term Loan To Employee. Because Amount " & vbNewLine & " Is Greater Than Long Term Limit.", vbInformation
        IsNotLongApplicable = True
    End If
    
End With


End Function

