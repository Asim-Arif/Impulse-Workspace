VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmSalariedEmps 
   BorderStyle     =   0  'None
   ClientHeight    =   7695
   ClientLeft      =   390
   ClientTop       =   5550
   ClientWidth     =   12780
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleMode       =   0  'User
   ScaleWidth      =   11447.22
   ShowInTaskbar   =   0   'False
   Begin ComboList.Usercontrol1 cmbAccNo 
      Height          =   285
      Left            =   8910
      TabIndex        =   10
      Top             =   900
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin ComboList.Usercontrol1 cmbBank 
      Height          =   285
      Left            =   2610
      TabIndex        =   9
      Top             =   885
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   11670
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salaried Employees"
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
         Index           =   1
         Left            =   4560
         TabIndex        =   7
         Top             =   120
         Width           =   3030
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salaried Employees"
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
         Index           =   2
         Left            =   4575
         TabIndex        =   8
         Top             =   135
         Width           =   3030
      End
   End
   Begin VB.Frame FLV 
      Height          =   6375
      Left            =   75
      TabIndex        =   0
      Top             =   1200
      Width           =   12630
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10920
         TabIndex        =   14
         Top             =   5430
         Width           =   1650
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5190
         Left            =   90
         TabIndex        =   1
         Top             =   195
         Width           =   12465
         _ExtentX        =   21987
         _ExtentY        =   9155
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15199215
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Emp ID"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Emp Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "F Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Account #"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Total"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   10410
         TabIndex        =   15
         Top             =   5475
         Width           =   465
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   375
         Left            =   11055
         TabIndex        =   13
         Top             =   5880
         Width           =   1500
         ForeColor       =   0
         Caption         =   " Close    "
         PicturePosition =   327683
         Size            =   "2646;661"
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
         Height          =   375
         Left            =   9480
         TabIndex        =   12
         Top             =   5880
         Width           =   1500
         ForeColor       =   0
         Caption         =   "Print         "
         PicturePosition =   327683
         Size            =   "2646;661"
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdSave 
         Height          =   375
         Left            =   7935
         TabIndex        =   11
         Top             =   5880
         Width           =   1500
         ForeColor       =   0
         Caption         =   "Save"
         PicturePosition =   327683
         Size            =   "2646;661"
         Accelerator     =   83
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   11205
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   1500
      Caption         =   "Refresh       "
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
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   630
      Width           =   11010
      BackColor       =   11517387
      Caption         =   $"frmSalariedEmps.frx":0000
      Size            =   "19420;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cmbYear 
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   885
      Width           =   1110
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "1958;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbMonth 
      Height          =   300
      Left            =   1245
      TabIndex        =   2
      Top             =   885
      Width           =   1365
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2408;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmSalariedEmps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Loading As Boolean

Private Sub cmbShift_Change()
    Call FillList
End Sub

Private Sub cmdCheckAll_Click()
    With LV
        For i = 1 To .ListItems.Count
            .ListItems(i).Checked = True
        Next i
    End With
End Sub

Private Sub cmdClearAll_Click()
    With LV
        For i = 1 To .ListItems.Count
            .ListItems(i).Checked = False
        Next i
    End With
End Sub

Private Sub cmbdept_matched()
    If Loading Then Exit Sub
    Call cmdRefresh_Click
End Sub

Private Sub cmbBank_matched()
    cmbAccNo.ClearVals
    cmbAccNo.AddVals con, "AccTitle", "VBankAccounts", "AccNo", " Where BankID=" & cmbBank.ID
End Sub

Private Sub cmbBank_UnMatched()
    MsgBox "HELLO"
End Sub

Private Sub cmbMonth_Change()
    If Loading Or cmbMonth.MatchFound = False Then Exit Sub
    Dim iDays As Integer
    Dim DT As Date
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    iDays = DateAdd("m", 1, DT) - DT
    cmbMonth.Tag = iDays
    txtMax = iDays
    txtMin = "1"
End Sub

Private Sub cmbMonth_Click()

    Call cmdRefresh_Click
End Sub

Private Sub cmbYear_Change()
    If Loading Then Exit Sub
    Dim iDays As Integer
    Dim DT As Date
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    iDays = DateAdd("m", 1, DT) - DT
    txtMax = iDays
    txtMin = "1"
End Sub

Private Sub cmdCancel_Click()
'    cmdSave.Visible = False
    cmdCancel.Visible = False
    cmdRefresh.Visible = True
    cmdClose.Visible = True
    'cmdPrint.Visible = True
End Sub

Private Sub cmbYear_Click()
Call cmdRefresh_Click
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
    
    If cmbYear.ListIndex = -1 Or cmbMonth.ListIndex = -1 Or cmbBank.MatchFound = False Or cmbAccNo.MatchFound = False Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptSalariedEmps.rpt")
    Dim f As New frmPrevRpt
    Dim SelFormula As String
    
    SelFormula = "Month({MonthlyBankSalaries.DT})=" & cmbMonth.List(cmbMonth.ListIndex, 1) & _
    " and Year({MonthlyBankSalaries.DT})=" & Val(cmbYear) & " AND {VMonthlySalaries.BankPymt}=True"
    
    f.ShowReport SelFormula & " ", rpt
    
    Screen.MousePointer = vbDefault
    
    
Exit Sub

err:

    MsgBox err.Description


End Sub

Public Sub cmdRefresh_Click()
    
    If Not cmbYear.MatchFound Or Not cmbMonth.MatchFound Then Exit Sub
    Call FillList
    
End Sub

Private Sub cmdSave_Click()

    On Error GoTo err
    If cmbMonth.ListIndex = -1 Or cmbYear.ListIndex = -1 Then
        Exit Sub
    End If
    Dim strVchrNo As String
    Call StartTrans(con)
    
    If cmbBank.Tag = "Saved" Then
        If cmbAccNo.Tag <> "" Then
            MsgBox "Delete Voucher # " & cmbAccNo.Tag & " and Try Again.", vbInformation
            Exit Sub
        End If
    End If
    
    con.Execute "Delete From MonthlyBankSalaries Where DT='" & DateSerial(cmbYear, cmbMonth.ListIndex + 2, 0) & "'"
    
    strVchrNo = AddVoucher
    con.Execute "Insert Into MonthlyBankSalaries(DT,BankAccNo,VchrNo) Values('" & DateSerial(cmbYear, cmbMonth.ListIndex + 2, 0) & "','" & cmbAccNo.ID & "','" & strVchrNo & "')"
        
    con.CommitTrans
    MsgBox "Saved Successfully", vbInformation
    
    Call cmdRefresh_Click
    
'    cmdSave.Visible = False
    
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Function AddVoucher() As String

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim i As Long, INextSno As Double, Bal As Double, Amt As Double, Amt1 As Double, Amt2 As Double
    Dim VchrNo As String, TheDesc As String, AccNo As String, AccNo1 As String, AccNo2 As String
    Dim Dated As Date
    
    With rs
        .Open "Select Sum(Balance) As TotalAmt From MonthlySalaries Where Month(DT)=" & _
         cmbMonth.ListIndex + 1 & " and year(DT)=" & cmbYear & " AND EmpID In(Select EmpID " & _
         "From Employees Where BankPymt=1 AND DeptID='VDRS')", con, adOpenForwardOnly, adLockReadOnly
        Amt1 = Val(!TotalAmt & "")
        
        .Close
        
        .Open "Select AccruedAccNo From Departments Where DeptID='VDRS'", con, adOpenForwardOnly, adLockReadOnly
        AccNo1 = !AccruedAccNo & ""
        .Close
        
        .Open "Select Sum(Balance) As TotalAmt From MonthlySalaries Where Month(DT)=" & _
         cmbMonth.ListIndex + 1 & " and year(DT)=" & cmbYear & " AND EmpID In(Select EmpID " & _
         "From Employees Where BankPymt=1 AND DeptID<>'VDRS')", con, adOpenForwardOnly, adLockReadOnly
        Amt2 = Val(!TotalAmt & "")
        
        .Close
        
        .Open "Select AccruedAccNo From Departments Where DeptID='VFNC'", con, adOpenForwardOnly, adLockReadOnly
        AccNo2 = !AccruedAccNo & ""
                
        .Close
    End With
    Dated = Date
            
    INextSno = getNextSno(Dated)
    VchrNo = getNextVchrNo(Dated, "JV")
            
    Amt = Amt1
    
    If Amt = 0 Then Exit Function
    'Give The Description
    TheDesc = "Salary Of  [VDRS] Directors Salary."
            
      '''''''''''''''''' D E B I T   E N T R I E S '''''''''''''
      '''''''''''''''''' Accrued Salary Account '''''''''''''
      AccNo = AccNo1
      Amt = Amt1
      Bal = getBalance(AccNo, Dated)
    
      Bal = Bal + Amt
          
      con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
       Dated & "','" & VchrNo & "','" & AccNo & _
       "','" & TheDesc & "'," & Amt & ",0," & Bal & ",'',0)"
      
      Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
      INextSno = INextSno + 1
        '''''''''''''''''' C R E D I T   E N T R I E S '''''''''''''
        
        TheDesc = "Salary Of Other Employees"
            
        '''''''''''''''''' D E B I T   E N T R I E S '''''''''''''
        '''''''''''''''''' Accrued Salary Account '''''''''''''
        AccNo = AccNo2
        Amt = Amt2
        Bal = getBalance(AccNo, Dated)
      
        Bal = Bal + Amt
            
        con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
         Dated & "','" & VchrNo & "','" & AccNo & _
         "','" & TheDesc & "'," & Amt & ",0," & Bal & ",'',0)"
        
        Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
        INextSno = INextSno + 1
        
        '''''''''''''''''' C R E D I T   E N T R I E S '''''''''''''
        
        AccNo = cmbAccNo.ID
        Amt = Amt1 + Amt2
        
        Bal = getBalance(AccNo, Dated)
                
        Bal = Bal - Amt
            
        con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
            Dated & "','" & VchrNo & "','" & AccNo & _
            "','" & TheDesc & "',0," & Amt & "," & Bal & ",'',0)"
        
        Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
         
        INextSno = INextSno + 1
               
        con.Execute "Insert Into VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & VchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
        
        
        Set rs = Nothing
        AddVoucher = VchrNo
    
    Exit Function
err:
    MsgBox err.Description
    
End Function

Private Sub DTPicker1_Change()
    Call FillList
End Sub

Private Sub Form_Activate()
    Loading = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()

    For i = 1 To LV.ColumnHeaders.Count
        LV.ColumnHeaders(i).Tag = LV.ColumnHeaders(i).Width
    Next
    
    Loading = True
    
    cmbMonth.AddItem "Jan"
    cmbMonth.List(0, 1) = 1
    cmbMonth.AddItem "Feb"
    cmbMonth.List(1, 1) = 2
    cmbMonth.AddItem "Mar"
    cmbMonth.List(2, 1) = 3
    cmbMonth.AddItem "Apr"
    cmbMonth.List(3, 1) = 4
    cmbMonth.AddItem "May"
    cmbMonth.List(4, 1) = 5
    cmbMonth.AddItem "Jun"
    cmbMonth.List(5, 1) = 6
    cmbMonth.AddItem "Jul"
    cmbMonth.List(6, 1) = 7
    cmbMonth.AddItem "Aug"
    cmbMonth.List(7, 1) = 8
    cmbMonth.AddItem "Sep"
    cmbMonth.List(8, 1) = 9
    cmbMonth.AddItem "Oct"
    cmbMonth.List(9, 1) = "10"
    cmbMonth.AddItem "Nov"
    cmbMonth.List(10, 1) = 11
    cmbMonth.AddItem "Dec"
    cmbMonth.List(11, 1) = 12
    
    
    
    Dim CYear As Long
    CYear = Format(GetServerDate, "yyyy")
    
    For i = 2002 To CYear
        cmbYear.AddItem i
    Next i
    
    cmbMonth = Format(GetServerDate, "MMM")
    cmbYear = CYear
    
    Call cmbMonth_Change
    cmbBank.ListHeight = 2600

    cmbAccNo.ListHeight = 2600
    
    cmbBank.ClearVals
    cmbBank.AddVals con, "Bank + ' {' + Branch + '}'", "BankList", "BankID"
    
'    cmdSave.Visible = False
    
End Sub

Private Sub FillList()
    
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.command
    Dim Total As Long, Cond As String
    
    With rs
    
        .Open "Select * From VMonthlyBankSalaries Where Year(DT)=" & cmbYear & " AND Month(DT)=" & cmbMonth.ListIndex + 1, con, adOpenForwardOnly, adLockReadOnly
        cmbBank.Tag = ""
        If Not .EOF Then
            cmbBank.ID = !BankID
            cmbBank.Tag = "Saved"
            cmbAccNo.ID = !BankAccNo & ""
            cmbAccNo.Tag = !VchrNo & ""
        End If
        .Close
        
        .Open "Select EmpID,Name,FName,EmpAccNo,Balance From VMonthlySalaries Where Year(DT)=" & cmbYear & " AND Month(DT)=" & cmbMonth.ListIndex + 1 & " AND BankPymt=1", con, adOpenForwardOnly, adLockReadOnly
        
        LV.ListItems.Clear
            
        For i = 0 To .RecordCount - 1
            Set ITM = LV.ListItems.add(, ![EmpId], ![EmpId])
            ITM.SubItems(1) = ![Name] & ""
            ITM.SubItems(2) = ![FName] & ""
            
            ITM.SubItems(3) = ![EmpAccNo] & ""
            ITM.SubItems(4) = Val(![Balance] & "")
            
            .MoveNext
        Next
            
        .Close
    End With
    
    Set rs = Nothing
    Call GetTotal
    
End Sub


Private Sub Form_Resize()
  
    On Error Resume Next
    Exit Sub
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
    
    FLV.Move 50, 1230, Me.ScaleWidth - 100, Me.ScaleHeight - 1400
     
    LV.Move 100, 200, Me.Width - 400, FLV.Height - FButs1.Height - 300
     
    FButs1.Move LV.Left + LV.Width - FButs1.Width, LV.Top + LV.Height
 
End Sub

Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    If Item.Checked Then
        Item.Checked = False
        Item.Selected = True
        Item.EnsureVisible
    Else
        Item.SubItems(6) = ""
        Item.SubItems(7) = ""
        Item.SubItems(8) = ""
        Call GetTotal
    End If
'    cmdSave.Enabled = True
    
End Sub

Private Sub mnuPrintSSForm_Click()

    Screen.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptOldAge.rpt")
    Dim f As New frmPrevRpt
    Dim SelFormula As String
    SelFormula = "Month({VOldAge.DT})=" & cmbMonth.List(cmbMonth.ListIndex, 1) & _
    " and Year({VOldAge.DT})=" & Val(cmbYear)
    
    f.ShowReport SelFormula, rpt
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub muuPrintSSSheet_Click()
    
    Screen.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptSocialSecurity.rpt")
    Dim f As New frmPrevRpt
    Dim SelFormula As String
    SelFormula = "Month({VSocialSecurity.DT})=" & cmbMonth.List(cmbMonth.ListIndex, 1) & _
    " and Year({VSocialSecurity.DT})=" & Val(cmbYear)
    
    f.ShowReport SelFormula & " ", rpt
    
    Screen.MousePointer = vbDefault
    
End Sub




Private Sub txtMax_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumber(KeyAscii)
End Sub

Private Sub txtMax_Validate(cancel As Boolean)

    Dim iDays As Integer
    Dim DT As Date
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    iDays = DateAdd("m", 1, DT) - DT
    With txtMax
        If Val(.Text) > iDays Then
            MsgBox "Value Can't Be More Than The Days Of Selected Month.", vbInformation
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            cancel = True
        ElseIf Val(.Text) < 2 Then
            MsgBox "Value Can't Be Less Than The 2.", vbInformation
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            cancel = True
        ElseIf Val(.Text) < Val(txtMin) Then
            MsgBox "Value Can't Be Less Than Min Limit.", vbInformation
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            cancel = True
        End If
    End With
    
End Sub

Private Sub txtMin_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumber(KeyAscii)
End Sub

Private Sub txtMin_Validate(cancel As Boolean)

    Dim iDays As Integer
    Dim DT As Date
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    iDays = DateAdd("m", 1, DT) - DT
    
    With txtMin
    
        If Val(.Text) < 1 Then
            MsgBox "Value Can't Be Less Than 1.", vbInformation
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            cancel = True
        ElseIf Val(.Text) > (iDays - 1) Then
            MsgBox "Value Can't Be More Than The Days Of Selected Month.", vbInformation
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            cancel = True
        ElseIf Val(.Text) > Val(txtMax.Text) Then
            MsgBox "Value Can't Be More Than Max Limit.", vbInformation
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            cancel = True
        End If
        
    End With
    
    
End Sub



Private Sub GetTotal()
    
    Dim Total As Double

    For i = 1 To LV.ListItems.Count
        Total = Total + Val(LV.ListItems(i).SubItems(4))
    Next
    txtTotal = Format(Total, "0.00")
    
End Sub
