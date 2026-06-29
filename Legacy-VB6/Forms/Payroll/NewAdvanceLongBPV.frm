VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form NewAdvanceLongBPV 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Long Term Loan..."
   ClientHeight    =   6165
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   8475
   StartUpPosition =   1  'CenterOwner
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
   Begin VB.Frame FraBank 
      Caption         =   "Cheque Details :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   75
      TabIndex        =   20
      Top             =   3345
      Width           =   8325
      Begin VB.CheckBox ChkClear 
         Caption         =   "Clear Cheque "
         Enabled         =   0   'False
         Height          =   225
         Left            =   6855
         TabIndex        =   23
         ToolTipText     =   "Clear This Cheque Automatically"
         Top             =   1335
         Value           =   1  'Checked
         Width           =   1350
      End
      Begin VB.TextBox Tbox 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   60
         TabIndex        =   22
         Top             =   1275
         Width           =   1725
      End
      Begin VB.CommandButton cmdBank 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7845
         Picture         =   "NewAdvanceLongBPV.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   555
         Width           =   360
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   5385
         TabIndex        =   24
         Top             =   1275
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   150929411
         CurrentDate     =   37250
      End
      Begin MSForms.ComboBox cmbCheckType 
         Height          =   315
         Left            =   1815
         TabIndex        =   30
         Top             =   1275
         Width           =   3510
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "6191;556"
         ListWidth       =   2822
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblBankAcc 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   45
         TabIndex        =   29
         Top             =   570
         Width           =   7755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bank Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   11
         Left            =   60
         TabIndex        =   28
         Top             =   330
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cheque No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   12
         Left            =   75
         TabIndex        =   27
         Top             =   1035
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cheque Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   13
         Left            =   1830
         TabIndex        =   26
         Top             =   1020
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   14
         Left            =   5370
         TabIndex        =   25
         Top             =   1065
         Width           =   390
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
      Left            =   75
      TabIndex        =   15
      Top             =   5745
      Width           =   1800
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1755
      TabIndex        =   2
      Top             =   735
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
      Format          =   151126019
      CurrentDate     =   37384
   End
   Begin MSComCtl2.DTPicker DTDedStartMonth 
      Height          =   315
      Left            =   6555
      TabIndex        =   31
      Top             =   735
      Width           =   1830
      _ExtentX        =   3228
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
      Format          =   151126019
      CurrentDate     =   40229
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Vchr # :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   9
      Left            =   3330
      TabIndex        =   34
      Top             =   780
      Width           =   585
   End
   Begin MSForms.TextBox txtVchrNo 
      Height          =   315
      Left            =   3930
      TabIndex        =   33
      Top             =   735
      Width           =   1590
      VariousPropertyBits=   746604575
      Size            =   "2805;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ded. Month :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   8
      Left            =   5580
      TabIndex        =   32
      Top             =   780
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Deduction Amount      :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   19
      Top             =   1500
      Width           =   1650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "New Ded. :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   3
      Left            =   4950
      TabIndex        =   18
      Top             =   1515
      Width           =   810
   End
   Begin MSForms.TextBox txtBal 
      Height          =   315
      Left            =   1755
      TabIndex        =   17
      Top             =   1110
      Width           =   2595
      VariousPropertyBits=   746604575
      Size            =   "4577;556"
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
      Width           =   2595
      VariousPropertyBits=   746604575
      Size            =   "4577;556"
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
      Top             =   1815
      Width           =   1575
   End
   Begin MSForms.TextBox txtDesc 
      Height          =   1515
      Left            =   1755
      TabIndex        =   5
      Top             =   1800
      Width           =   6630
      VariousPropertyBits=   -1467987941
      ScrollBars      =   3
      Size            =   "11695;2672"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   1950
      X2              =   8360
      Y1              =   5580
      Y2              =   5580
   End
   Begin MSForms.TextBox txtAmt 
      Height          =   315
      Left            =   5790
      TabIndex        =   3
      Top             =   1110
      Width           =   2595
      VariousPropertyBits=   746604571
      BackColor       =   16761024
      ForeColor       =   255
      Size            =   "4577;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox txtDeduct 
      Height          =   315
      Left            =   5790
      TabIndex        =   4
      Top             =   1455
      Width           =   2595
      VariousPropertyBits=   746604571
      Size            =   "4577;556"
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
      Left            =   4950
      TabIndex        =   13
      Top             =   1170
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Previous Balance        :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   1170
      Width           =   1635
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   405
      Left            =   4095
      TabIndex        =   7
      Top             =   5640
      Width           =   2085
      Caption         =   "Save & Close"
      PicturePosition =   327683
      Size            =   "3678;714"
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
      Height          =   405
      Left            =   1920
      TabIndex        =   6
      Top             =   5640
      Width           =   2085
      Caption         =   "Save & New    "
      PicturePosition =   327683
      Size            =   "3678;714"
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
      Height          =   405
      Left            =   6270
      TabIndex        =   8
      Top             =   5640
      Width           =   2085
      Caption         =   "Cancel    "
      PicturePosition =   327683
      Size            =   "3678;714"
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
      Top             =   780
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
      X1              =   1950
      X2              =   8360
      Y1              =   5595
      Y2              =   5595
   End
End
Attribute VB_Name = "NewAdvanceLongBPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_bContractors As Boolean
Public LongTerm As Boolean

Dim lRequestEntryID As Long

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
'        FormulaField.Text = "'" & company & "'"
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
    'ElseIf Val(txtDeduct) = 0 And txtDeduct.Visible Then
    '    MsgBox "Invalid Deduction Amount.", vbInformation
    '    Exit Function
        
    ElseIf Val(txtDeduct) > Val(txtAmt) And txtDeduct.Visible Then
        MsgBox "Deduction Cannot Be More Than Loan Amount.", vbInformation
        Exit Function
    ElseIf strEmpLongTermAccNo = "" Then
        MsgBox "Can't Post, Emp. Long Term Loan Account Not Found.", vbInformation
        Exit Function
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
    Dim Issued As Boolean, ChqBookNo As Double
    
    If lblBankAcc = "" Then
        MsgBox "Invalid Bank Selected.", vbInformation
        Exit Function
    ElseIf Not ValidChqNo(GetID(lblBankAcc), Tbox(2), False, ChqBookNo) Then
        If Issued Then
            MsgBox "The Cheque With Number '" & txtChqNo & "' Has Been Issued.", vbInformation
        Else
            MsgBox "Invalid Cheque No.", vbInformation
        End If
        Tbox(2).SetFocus
        Exit Function
    End If
    Tbox(2).Tag = ChqBookNo
    
    Dim strVchrNo As String
         
    Call StartTrans(con)
    
    If Not m_bContractors Then
        strVchrNo = SaveTransForCheque
    End If
    
    SQL = "INSERT INTO Advances(EmpID,DT,Type,Description," & _
     "Amount,DAmount,AmountCleared,AccVoucherNo,DeductionStartDT) VALUES('" & cmbEmp.ID & _
     "','" & DTPicker1 & "',1,'" & txtDesc & "'," & Val(txtAmt) & "," & _
     Val(txtDeduct) & ",0,'" & strVchrNo & "','" & DTDedStartMonth & "')"
     
    con.Execute SQL
'    If lRequestEntryID <> 0 Then
'        con.Execute "UPDATE EmpAdvanceRequest SET PostedVchrNo='" & strVchrNo & "' WHERE EntryID=" & lRequestEntryID
'    End If
    
    con.CommitTrans
        
    MsgBox "Saved Successfully.", vbInformation
    
    Saved = True
    If chkPrint = 1 Then Call PrintVoucher(MainForm.Cr1, strVchrNo)
    SQL = ""
    Exit Function
err:
    MsgBox err.Description

End Function

Private Function SaveTransForCheque() As String

    On Error GoTo err
    'If CheckData = False Then Exit Function

    Dim i As Integer
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim Bal As Double
    Dim RcptNo As String
    
    RcptNo = "" 'IIf(Len(Trim(Tbox(3))) = 0, "", "Rcpt # :" & Tbox(3))
    
    'Get Serial # and Balances
    INextSno = getNextSno(DTPicker1.Value, "Cheque")
    SnoToUse = INextSno
    
    AccNo = strEmpLongTermAccNo 'Employee Permanent Loan Account
    
    theDesc = "Long Term Advance Paid To " & cmbEmp.Text
    'Start New Transaction
    'con.BeginTrans
    bankDesc = Right(lblBankAcc, Len(lblBankAcc) - InStr(lblBankAcc, "]")) & " Chq # :" & Tbox(2)
    
    con.Execute "INSERT INTO Cheque (SNo,CDate,AccNo,Description,Amount,BankID,chequeno,ChqBookNo,chequeType,chequeDate,Posted,Payment,Bounced,BankAccNo)  Values(" & INextSno & ",'" & Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & _
    AccNo & "','" & theDesc & " (" & bankDesc & ") '," & txtAmt & "," & lblBankAcc.Tag & ",'" & _
    Tbox(2) & "'," & Tbox(2).Tag & ",'" & cmbCheckType & "','" & DTPicker2 & "',0,1,0,'" & GetID(lblBankAcc) & "')"
    'Remove the Cheque From Cheque book
    Call IssueChq(Val(Tbox(2).Tag), Tbox(2))
    
'    With LVPymt.ListItems
'        For i = 1 To .Count
'            If .Item(i).Checked Then con.Execute "Insert into ChequeDetail Values('" & SnoToUse & "','" & Val(.Item(i).key) & "'," & Val(LVPymt.ListItems(i).SubItems(4)) & ")"
'        Next
'    End With
    
    'Commit Transaction
    'con.CommitTrans
    
    SaveTransForCheque = SavePaymentTrans(SnoToUse)

    Call PlaySaveSound
    'MsgBox "Transaction Saved Successfuly.", vbInformation
   
    Exit Function
err:
    MsgBox "Transaction Terminated Due to ERROR:" & err.Number & vbNewLine & err.Description, vbCritical
    SaveTransForCheque = ""
    con.RollbackTrans
    
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
    theDesc = "Long Term Advance Paid To " & cmbEmp.Text & " Total Amount : " & Amt & " Deduction : " & Val(txtDeduct)
     
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
'    If strCompany = "Dr-Frgz" And strSubCompany = "Intl" Or strCompany = "Banzai" Then
'        AccNo = "15-004-12001"
'    Else
'        AccNo = "13-003-12001" 'IIf(LongTerm, LongAdvAccNo, ShortAdvAccNo)
'    End If
    AccNo = strEmpLongTermAccNo
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
    
    SQL = "(Select isnull(Sum(AmtClrd),0) balance From AmtCleared Where empid='" & cmbEmp.ID & "')"
        
    Balance = Val(con.Execute("Select Sum(Amount)-" & SQL & "  from Advances Where empid='" & cmbEmp.ID & "' and Type=1 ").Fields(0) & "")
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

Private Sub cmdBank_Click()

    Dim BankID As Double
    lblBankAcc = SelectBankAcc(BankID)
    lblBankAcc.Tag = BankID
    
    'Get ChqNo
         
    Dim ChqBookNo As Double
    Tbox(2) = GetNextChqNo(GetID(lblBankAcc), ChqBookNo)
    Tbox(2).Tag = ChqBookNo
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If Saved Then
        Call DTPicker1_Change
        cmbEmp.ID = ""
        txtAmt = ""
        txtDeduct = ""
        cmbEmp.SetFocus
        'Unload Me
    End If
End Sub

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub


Private Sub DTPicker1_Change()
    txtVchrNo = getNextVchrNo(DTPicker1, "BPV")
End Sub

Private Sub Form_Activate()

'    Label1(5).Visible = LongTerm
'    txtDeduct.Visible = LongTerm
   
    txtDesc = "Long Term Loan"
    Me.Caption = "Long Term Loan"
    
    DTPicker2 = DTPicker1.Value
    cmbCheckType.AddItem "Open"
    cmbCheckType.AddItem "Payees A/C Only"
    cmbCheckType.ListIndex = 0
    chkPrint = Abs(bPrintWhenSaved)
    DTDedStartMonth.Value = DateSerial(year(Date), Month(Date), 1)
    
End Sub

Private Sub Form_Load()

    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    
    'cmbType.AddItem "Short Term"
    'cmbType.AddItem "Long Term"
    'cmbType.ListIndex = 0
    m_bContractors = False
    lRequestEntryID = 0
    
    DTPicker1 = GetServerDate
    Call DTPicker1_Change
    
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

Public Sub ShowMe(p_bContractors, Optional p_lRequestEntryID = 0)
    
    lRequestEntryID = p_lRequestEntryID
    m_bContractors = p_bContractors
    Dim strCond As String
    
    If m_bContractors Then
        strCond = "Where IsNull(TempDept,0)=1"
    Else
        strCond = "Where IsNull(TempDept,0)=0"
    End If
    
    If lRequestEntryID = 0 Then
        With cmbDept
            .ClearVals
            .AddItem "<All Departments>", "0"
            .AddVals con, "Name", "Departments", "DeptID", strCond
            .Text = "<All Departments>"
        End With
    Else
        Dim rs As New ADODB.Recordset
        With rs
            .Open "SELECT * FROM VEmpAdvanceRequest WHERE EntryID=" & lRequestEntryID, con, adOpenForwardOnly, adLockReadOnly
            cmbDept.AddVals con, "Name", "Departments", "DeptID", " WHERE DeptID='" & !DeptID & "'"
            cmbDept.ID = !DeptID & ""
            cmbDept.Enabled = False
            cmbEmp.AddVals con, "Name", "Employees", "EmpID", " WHERE EmpID='" & !EmpId & "'"
            cmbEmp.ID = !EmpId & ""
            cmbEmp.Enabled = False
'            DTPicker1.Value = !DT
            DTPicker1.Enabled = False
            txtAmt = Val(!ApprovedAmt & "")
            txtAmt.Locked = True
            txtDeduct = Val(!ApprovedDed & "")
            txtDeduct.Locked = True
            .Close
        End With
    End If
    Me.Show 1
End Sub

Private Function SavePaymentTrans(ChqSNo As Double) As String

    Dim i As Integer
    Dim NextVchrNo As String
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim dAmount As Double
    Dim sScNo As String
    Dim sCIH As String
    Dim Bal As Double
    Dim Desc As String
    Dim AccNo As String
    Dim Dated As Date
    'Get Serial # and Balances
    Dim dLongTerm As Double
    'dLongTerm = Val(txtFixLoan.Text)
'    If (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Banzai" Then
'        AccNo = "15-004-12001"
'    Else
'        AccNo = "13-003-12001"
'    End If
    AccNo = strEmpLongTermAccNo
    
    Dated = DTPicker2
    
    NextVchrNo = getNextVchrNo(Dated, "BPV")
    INextSno = getNextSno(Dated)
    
    SnoToUse = INextSno
    
    'Party Account's Balance
    Bal = getBalance(AccNo, Dated)
    dAmount = Val(txtAmt)
    
    sScNo = ChqSNo
    
    theDesc = "Long Term Advance Paid To " & cmbEmp.Text
    
    'Bank Account No.
    sCIH = GetID(lblBankAcc)
    
    Bal = Bal + dAmount
    Dim bWaitForAuth As Boolean
    bWaitForAuth = VoucherRequiresAuthorization("BPV")
    
    'Debit Entry
    bankDesc = Right(lblBankAcc, Len(lblBankAcc) - InStr(lblBankAcc, "]")) & " Chq # :" & Tbox(2)
    theDesc = theDesc & " " & bankDesc
     
    If bWaitForAuth = False Then
        con.Execute "INSERT INTO Vouchers Values(" & INextSno & ",'" & _
         Dated & "','" & NextVchrNo & "','" & AccNo & _
         "','" & theDesc & "'," & dAmount & ",0," & Bal & ",''," & sScNo & ")"
    
        Call adjustLedger(AccNo, Dated, INextSno, Val(dAmount))
    
        Dim ChqNo As String
        INextSno = INextSno + 1
        Bal = getBalance(sCIH, Dated)
        dAmount = Val(txtAmt)
        Bal = Bal - dAmount
        'Description
    
        Desc = "Cheque Paid To " & cmbEmp.Text & " Chq# : " & Tbox(2)
        ChqNo = Tbox(2)
           
        'Credit Entry
    
        con.Execute "INSERT INTO Vouchers Values(" & INextSno & ",'" & _
         Dated & "','" & NextVchrNo & "','" & sCIH & _
         "','" & theDesc & "',0," & dAmount & "," & Bal & ",''," & sScNo & ")"
    
        Call adjustLedger(sCIH, Dated, INextSno, -Val(dAmount))
    Else
    
        con.Execute "INSERT INTO VouchersForAuthorization(VDate,VchrNo,Accno,Description,Debit,Credit,balance,DpstSlip,CSNo) VALUES('" & _
         Dated & "','" & NextVchrNo & "','" & AccNo & "','" & theDesc & "'," & dAmount & ",0," & Bal & ",''," & sScNo & ")"
       
        ChqNo = Tbox(2)
        
        con.Execute "INSERT INTO VouchersForAuthorization(VDate,VchrNo,Accno,Description,Debit,Credit,balance,DpstSlip,CSNo) VALUES('" & _
         Dated & "','" & NextVchrNo & "','" & sCIH & _
         "','" & theDesc & "',0," & dAmount & "," & Bal & ",''," & sScNo & ")"
         
    End If
    Dim rsCopy As New ADODB.Recordset
    With rsCopy
        .CursorLocation = adUseClient
        .Open "Select * from ChequeDetail Where (Sno=" & sScNo & ")", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            con.Execute "Insert into VoucherDetail Values('" & ![BalancedSno] & "','" & SnoToUse & "'," & ![Amount] & ")"
            .MoveNEXT
        Loop
        .Close
    End With
    Set rsCopy = Nothing
    con.Execute "Update Cheque Set Posted=1 Where Sno=" & sScNo
    
    Dim myDT As Date, dpAmount As Double, DTLastReconciliation As Date
    
    If bWaitForAuth = False Then
        con.Execute "Insert Into VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & NextVchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
    Else
        con.Execute "INSERT INTO VouchersForAuthorizationInfo(VchrNo,UserName,MachineName,DT,ApprovedBy,Notes,Payee) VALUES('" & NextVchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "','','','')"
    End If
        
    If NextVchrNo <> txtVchrNo Then
        MsgBox "Vchr No. " & txtVchrNo & " had been used, " & NextVchrNo & " is used Instead.", vbInformation
    End If
    SavePaymentTrans = NextVchrNo
    Exit Function
err:
    'MsgBox "Transaction Terminated Due To Following Error;" & vbNewLine & err.Description
    SavePaymentTrans = ""
End Function



