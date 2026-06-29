VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmClearShortTerm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clear / Transfer Short Term Loan"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10590
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
   ScaleHeight     =   6195
   ScaleWidth      =   10590
   StartUpPosition =   1  'CenterOwner
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
      Left            =   1515
      TabIndex        =   17
      Top             =   5220
      Width           =   2235
   End
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
      Left            =   6750
      TabIndex        =   14
      Top             =   5235
      Width           =   1950
   End
   Begin ComboList.Usercontrol1 cmbMainDept 
      Height          =   315
      Left            =   1515
      TabIndex        =   0
      Top             =   45
      Width           =   7185
      _ExtentX        =   12674
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
      Left            =   1515
      TabIndex        =   1
      Top             =   375
      Width           =   7185
      _ExtentX        =   12674
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
   Begin ComboList.Usercontrol1 cmbEmpTo 
      Height          =   315
      Left            =   1305
      TabIndex        =   12
      Top             =   5700
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
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
   Begin MSComctlLib.ListView LV 
      Height          =   3825
      Left            =   1515
      TabIndex        =   10
      Top             =   1080
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   6747
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "S #"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Description"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1515
      TabIndex        =   2
      Top             =   720
      Width           =   1530
      _ExtentX        =   2699
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
      Format          =   158597123
      CurrentDate     =   37384
   End
   Begin ComboList.Usercontrol1 cmbAccNo 
      Height          =   285
      Left            =   1515
      TabIndex        =   18
      Top             =   4920
      Visible         =   0   'False
      Width           =   7185
      _ExtentX        =   12674
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
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1755
      Left            =   8730
      Stretch         =   -1  'True
      Top             =   45
      Width           =   1785
   End
   Begin MSForms.ComboBox cmbCashInHand 
      Height          =   300
      Left            =   1515
      TabIndex        =   16
      Top             =   4920
      Width           =   7185
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "12674;529"
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
      Left            =   30
      TabIndex        =   15
      Top             =   4920
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Transfer To :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   330
      TabIndex        =   13
      Top             =   5745
      Visible         =   0   'False
      Width           =   945
   End
   Begin MSForms.CommandButton cmdTransfer 
      Height          =   465
      Left            =   3885
      TabIndex        =   11
      Top             =   5685
      Visible         =   0   'False
      Width           =   1530
      Caption         =   "Transfer"
      PicturePosition =   327683
      Size            =   "2699;820"
      Accelerator     =   84
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
      Caption         =   "Department :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   8
      Left            =   525
      TabIndex        =   9
      Top             =   90
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Remarks :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   6
      Left            =   3075
      TabIndex        =   8
      Top             =   750
      Width           =   720
   End
   Begin MSForms.TextBox txtDesc 
      Height          =   315
      Left            =   3810
      TabIndex        =   7
      Top             =   720
      Width           =   4890
      VariousPropertyBits=   679495707
      ScrollBars      =   3
      Size            =   "8625;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   465
      Left            =   5505
      TabIndex        =   3
      Top             =   5685
      Width           =   1530
      Caption         =   "Clear"
      PicturePosition =   327683
      Size            =   "2699;820"
      Accelerator     =   108
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   465
      Left            =   7170
      TabIndex        =   4
      Top             =   5685
      Width           =   1530
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "2699;820"
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
      Left            =   1035
      TabIndex        =   6
      Top             =   750
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Employee :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   690
      TabIndex        =   5
      Top             =   375
      Width           =   795
   End
End
Attribute VB_Name = "frmClearShortTerm"
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
    'ID = Val(con.Execute("SELECT MAX(EntryID) FROM Advances").Fields(0) & "")
    
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
    
    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & company & "'"
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


Private Sub cmbdept_matched()

    Dim strCondition As String
    strCondition = " WHERE Active=1"
    If cmbDept.ID = "0" Then
        If cmbMainDept.ID <> "0" Then
            strCondition = strCondition & " AND MainUnit='" & cmbMainDept.ID & "'"
        End If
    Else
        strCondition = strCondition & " AND DeptID='" & cmbDept.ID & "'"
    End If
            
    cmbEmp.ClearVals
    cmbEmp.AddVals con, "'{'+EmpID+'} ' + Name + ' ' + Rel + ' ' + FName", "Employees", "EmpID", strCondition
     
End Sub

Private Function Saved(Optional bTransfer As Boolean = False) As Boolean

    Saved = False
    'On Error GoTo ERR
    Dim SQL As String
    
    If Not cmbEmp.MatchFound Then
        MsgBox "Please Select Valid Employee From The List.", vbInformation
        Exit Function
    End If
    
    If bTransfer Then
        If cmbEmpTo.MatchFound = False Then
            MsgBox "Please Select Emp to.", vbInformation
            Exit Function
        ElseIf cmbEmp.ID = cmbEmpTo.ID Then
            MsgBox "Can't Transfer to the Same Employee.", vbInformation
            Exit Function
        End If
        
    End If
    
    Dim i As Integer, bFound As Boolean
    bFound = False
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            bFound = True
            Exit For
        End If
    Next
    
    If bFound = False Then
        MsgBox "No Entry to Clear.", vbInformation
        Exit Function
    End If
    
    If bTransfer = False Then
        If chkDontPostVoucher.Value = vbUnchecked Then
            If chkOtherAccount.Value = vbChecked And cmbAccNo.MatchFound = False Then
                MsgBox "Please Select Account.", vbInformation
                Exit Function
            ElseIf chkOtherAccount.Value = vbUnchecked And cmbCashInHand.MatchFound = False Then
                MsgBox "Please Select Cash Account.", vbInformation
                Exit Function
            End If
        End If
    End If
    
    Dim lEntryID As Long, strVchrNo As String
    Call StartTrans(con)
    
    If bTransfer Then
        con.Execute "INSERT INTO ShortTermAdvanceClearance(EmpID,DT,Remarks,UserName,MachineName,Type,EmpIDTo) VALUES('" & cmbEmp.ID & _
         "','" & DTPicker1.Value & "','" & txtDesc & "','" & CurrentUserName & "','" & strComputerName & "',1,'" & cmbEmpTo.ID & "')"
    Else
        If chkDontPostVoucher.Value = vbUnchecked Then
            strVchrNo = AddVoucher
        End If
        con.Execute "INSERT INTO ShortTermAdvanceClearance(EmpID,DT,Remarks,UserName,MachineName,Type,VchrNo) VALUES('" & cmbEmp.ID & _
         "','" & DTPicker1.Value & "','" & txtDesc & "','" & CurrentUserName & "','" & strComputerName & "',0,'" & strVchrNo & "')"
    End If
         
    lEntryID = GetSingleLongValue("MAX(EntryID)", "ShortTermAdvanceClearance")
    Dim lNewEntryID As Long
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                
                If bTransfer Then
                    
                    con.Execute "UPDATE Advances SET Cleared=4 WHERE EntryID=" & Val(.key)
                    
                    con.Execute "INSERT INTO Advances(EmpID,DT,Type,Description,Amount,DAmount,AmountCleared,AccVoucherNo" & _
                     ") SELECT '" & cmbEmpTo.ID & "',DT,Type,'" & txtDesc & "',Amount," & _
                     "DAmount,AmountCleared,AccVoucherNo FROM Advances WHERE EntryID=" & Val(.key)
                          
                    lNewEntryID = GetSingleLongValue("MAX(EntryID)", "Advances")
                    
                    con.Execute "INSERT INTO ShortTermAdvanceClearanceDetail(SAC_RefID,Advances_RefID,Advances_To_RefID) VALUES(" & lEntryID & "," & Val(.key) & "," & lNewEntryID & ")"
                Else
                    con.Execute "UPDATE Advances SET Cleared=2 WHERE EntryID=" & Val(.key)
                    con.Execute "INSERT INTO ShortTermAdvanceClearanceDetail(SAC_RefID,Advances_RefID) VALUES(" & lEntryID & "," & Val(.key) & ")"
                End If
            End If
        End With
    Next
    
    con.CommitTrans
    Saved = True
    
    SQL = ""
    Exit Function
err:
    MsgBox err.Description

End Function

Private Sub cmbType_Change()

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

Private Sub cmbEmp_matched()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "SELECT * FROM Advances WHERE EmpID='" & cmbEmp.ID & "' AND Type=0 AND Cleared=0 ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, !EntryID & "'", LV.ListItems.count + 1)
            ITM.ListSubItems.add , , Format(!DT, "dd-MMM-yyyy")
            ITM.ListSubItems.add , , Val(!Amount & "")
            ITM.ListSubItems.add , , !Description & ""
            .MoveNEXT
        Loop
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

Private Sub cmbEmpCode_matched()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select MainUnit,DeptID,Name From Employees Where EmpID='" & cmbEmpCode.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbMainDept.ID = !mainUnit & ""
        cmbDept.ID = !DeptID & ""
        cmbEmp.ID = cmbEmpCode.ID
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
     
End Sub

Private Sub cmbMainDept_matched()

    
    cmbEmp.ClearVals
    'cmbEmp.AddItem "<All Employees>", "0"
    If cmbMainDept.Text = "<All Departments>" Then
        If m_bContractors Then
            cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' - ' + FName + ' {' + Designation + '}'", "VTempEmps", "EmpID", "Where Active=1 "
        Else
            cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' - ' + FName + ' {' + Designation + '}'", "VSalariedEmps", "EmpID", "Where Active=1 "
        End If
    Else
        cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' - ' + FName + ' {' + Designation + '}'", "Employees", "EmpID", "Where DeptID='" & cmbMainDept.ID & "' and Active=1"
    End If
    
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()

    If Saved Then
         Unload Me
         frmClearShortTerm.Show 1
    End If
End Sub

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub


Private Sub cmdTransfer_Click()
    If Saved(True) Then Unload Me
End Sub

Private Sub Form_Load()
    
    cmbMainDept.ListHeight = 2200

    cmbEmp.ListHeight = 1250

    
    cmbMainDept.AddItem "<All Departments>", "0"
    cmbMainDept.AddVals con, "Name", "Departments", "DeptID"
    cmbMainDept.ID = "0"
   
    DTPicker1 = GetServerDate
    
    cmbEmpTo.AddVals con, "'{'+EmpID+'} ' + Name + ' ' + Rel + ' ' + FName", "VSalariedEmps", "EmpID", "WHERE Active=1 "
      
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

Private Function AddVoucher() As String

    Dim i As Long, INextSno As Double, Bal As Double, Amt As Double
    Dim VchrNo As String, theDesc As String, AccNo As String
    
    Dim Dated As Date
    Dated = DTPicker1
    'Charge All Amount to Cash
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            Amt = Amt + Val(LV.ListItems(i).SubItems(2))
        End If
    Next
    
    
    INextSno = getNextSno(Dated)
    VchrNo = getNextVchrNo(Dated, "CRV")
     
    'Give The Description
    theDesc = "Temp. Advance Received from " & cmbEmp.Text
     
    ''''''''''''''''Charge To  Cash Entry'''''''''''''
    If chkOtherAccount.Value = vbChecked Then
        AccNo = cmbAccNo.ID
    Else
        AccNo = cmbCashInHand.List(cmbCashInHand.ListIndex, 1) 'CashInHand
    End If
     
    Bal = getBalance(AccNo, Dated)
    
    Bal = Bal + Amt
    
    'Credit Entry
    con.Execute "INSERT INTO Vouchers VALUES(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & theDesc & "'," & Amt & ",0," & Bal & ",'',0)"
    
    
    Call adjustLedger(AccNo, Dated, INextSno, Amt)
     
     
     INextSno = INextSno + 1
     
    '''''''''''''''''' D E B I T   E N T R I E S '''''''''''''
      
      'Determine the Account no for Staff Member
      'Long / Short Tem Advance
      
      AccNo = strEmpShortTermAccNo '"15-007-001-19001" '"13-003-19001" 'IIf(LongTerm, LongAdvAccNo, ShortAdvAccNo)
     
      Bal = getBalance(AccNo, Dated)
      
      Amt = Amt
      Bal = Bal - Amt
    
    
      'Debit To Staff Long/Short Advance Account
      
      con.Execute "INSERT INTO Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & theDesc & "',0," & Amt & "," & Bal & ",'',0)"
        
     
    Call adjustLedger(AccNo, Dated, INextSno, -Amt)
    
    con.Execute "Insert Into VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & VchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
    
    AddVoucher = VchrNo
    
End Function


