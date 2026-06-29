VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "combolist.ocx"
Begin VB.Form frmConvertShortToLongTerm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Long Term Loan..."
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   285
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
   ScaleHeight     =   3750
   ScaleWidth      =   5100
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraButts 
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   15
      TabIndex        =   16
      Top             =   3075
      Width           =   5070
      Begin MSForms.CommandButton cmdCancel 
         Height          =   360
         Left            =   3555
         TabIndex        =   19
         Top             =   225
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
      Begin MSForms.CommandButton cmdNext 
         Height          =   360
         Left            =   150
         TabIndex        =   18
         Top             =   225
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
      Begin MSForms.CommandButton cmdOk 
         Height          =   360
         Left            =   1800
         TabIndex        =   17
         Top             =   225
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
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   1
         X1              =   180
         X2              =   4880
         Y1              =   165
         Y2              =   165
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   180
         X2              =   4880
         Y1              =   180
         Y2              =   180
      End
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   1755
      TabIndex        =   0
      Top             =   30
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
      Enabled         =   0   'False
   End
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   1755
      TabIndex        =   1
      Top             =   375
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
      Enabled         =   0   'False
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1785
      TabIndex        =   2
      Top             =   765
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
      Format          =   103153667
      CurrentDate     =   37384
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Deduction Amount      :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   15
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
      TabIndex        =   14
      Top             =   1575
      Width           =   810
   End
   Begin MSForms.TextBox txtBal 
      Height          =   315
      Left            =   1770
      TabIndex        =   13
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
      TabIndex        =   12
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
      Caption         =   "Reason of Conv. :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   6
      Left            =   450
      TabIndex        =   11
      Top             =   1920
      Width           =   1320
   End
   Begin MSForms.TextBox txtDesc 
      Height          =   1005
      Left            =   1770
      TabIndex        =   5
      Top             =   1920
      Width           =   3090
      VariousPropertyBits=   -1467987941
      ScrollBars      =   3
      Size            =   "5450;1773"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtAmt 
      Height          =   315
      Left            =   3810
      TabIndex        =   3
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
   Begin MSForms.TextBox txtDeduct 
      Height          =   315
      Left            =   3810
      TabIndex        =   4
      Top             =   1515
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
      TabIndex        =   10
      Top             =   1185
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Previous Balance        :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   105
      TabIndex        =   9
      Top             =   1185
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   1275
      TabIndex        =   8
      Top             =   780
      Width           =   450
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
Attribute VB_Name = "frmConvertShortToLongTerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lRequestEntryID As Long, lAdvanceEntryID As Long

Const iHeightForRequest As Integer = 4800
Const iHeightForApproval As Integer = 5880
Const iButtonsFrameTopForRequest As Integer = 3645
Const iButtonsFrameTopForApproval As Integer = 4710

Private Sub PrintReport()

    Screen.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report, ID As Integer
    
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\LoanConvertRequest.rpt")
    ID = Val(con.Execute("Select MAX(EntryID) FROM EmpLoanConvertRequest").Fields(0) & "")
    
    Unload Me
    f.ShowReport "{EmpAdvanceRequest.EntryID}=" & ID, rpt
     
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmbdept_matched()
    
'    If lRequestEntryID <> 0 Then Exit Sub
'    cmbEmp.ClearVals
'    'cmbEmp.AddItem "<All Employees>", "0"
'    If cmbDept.Text = "<All Departments>" Then
'        cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name", "VSalariedEmps", "EmpID", "Where Active=1 "
'    Else
'        cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "' and Active=1"
'    End If
    
End Sub

Private Function Saved() As Boolean

    Saved = False
    'On Error GoTo ERR
    Dim Sql As String
    
    If lRequestEntryID = 0 Then
        If Not cmbEmp.MatchFound Then
            MsgBox "Please Select Valid Employee From The List.", vbInformation
            Exit Function
        ElseIf Val(txtAmt) = 0 Then
            MsgBox "Invalid Amount.", vbInformation
            Exit Function
        ElseIf Val(txtDeduct) > Val(txtAmt) Or Val(txtDeduct) <= 0 Then
            MsgBox "Invalid Deduction Amount.", vbInformation
            Exit Function
        End If
    End If
    
    Call StartTrans(con)
    Dim strVchrNo As String
    If lRequestEntryID = 0 Then
        
        con.Execute "INSERT INTO Advances(EmpID,DT,Type,Description,Amount,DAmount,AmountCleared,AccVoucherNo) VALUES('" & _
         cmbEmp.ID & "','" & DTPicker1.Value & "',1,'Converted From Short Term, EntryID:" & lAdvanceEntryID & "'," & Val(txtAmt) & _
         "," & Val(txtDeduct) & ",0,'" & strVchrNo & "')"
     
        con.Execute "INSERT INTO ShortTermToLongTermConversions(Advances_RefID,UserName,MachineName,DT) VALUES(" & _
         lAdvanceEntryID & ",'" & CurrentUserName & "','" & strComputerName & "','" & DTPicker1.Value & "')"
         
        con.Execute "UPDATE Advances SET Cleared=2 WHERE EntryID=" & lAdvanceEntryID
    End If
    
    con.CommitTrans
    
    Saved = True
    'If chkPrint = 1 Then Call PrintReport
    Sql = ""
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
    
    AccNo = CashInHand
    
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
    Dim Sql As String, Balance As Double
    Sql = "(Select isnull(Sum(AmtClrd),0) balance From AmtCleared Where empid='" & cmbEmp.ID & "')"
        
    Balance = Val(con.Execute("Select Sum(Amount)-" & Sql & "  from Advances Where empid='" & cmbEmp.ID & "' and Type=1 ").Fields(0) & "")
    txtBal = Balance
    
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
         Unload Me
    End If
End Sub

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub




Private Sub Form_Load()
    
    
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    
    'cmbType.AddItem "Short Term"
    'cmbType.AddItem "Long Term"
    'cmbType.ListIndex = 0
       
    DTPicker1 = GetServerDate
    
   
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

Public Sub ShowMe(p_lAdvanceEntryID As Long)
        
    lRequestEntryID = 0
    lAdvanceEntryID = p_lAdvanceEntryID
    Dim strDeptID As String, strEmpID As String, dAmt As Double
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT Advances.Amount,Employees.EmpID,Employees.DeptID FROM Advances INNER JOIN Employees ON Advances.EmpID=Employees.EmpID WHERE Advances.EntryID=" & lAdvanceEntryID, con, adOpenForwardOnly, adLockReadOnly
        strDeptID = !DeptID & ""
        strEmpID = !EmpId & ""
        dAmt = Val(!Amount & "")
        .Close
    End With
    If lRequestEntryID = 0 Then
        With cmbDept
            .ClearVals
            .AddVals con, "Name", "Departments", "DeptID", " WHERE DeptID='" & strDeptID & "'"
            .ID = strDeptID
        End With
        With cmbEmp
            .ClearVals
            .AddVals con, "'{' + EmpID + '} ' +  Name", "Employees", "EmpID", "WHERE EmpID='" & strEmpID & "'"
            .ID = strEmpID
        End With
        txtAmt = dAmt
        
        Me.Height = iHeightForRequest
        fraButts.Top = iButtonsFrameTopForRequest
        Me.Show 1
    End If
    Set rs = Nothing
    
End Sub

