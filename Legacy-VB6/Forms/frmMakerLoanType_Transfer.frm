VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmMakerLoanType_Transfer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maker Loan Transfer"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
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
   ScaleHeight     =   5160
   ScaleWidth      =   6705
   StartUpPosition =   1  'CenterOwner
   Begin ComboList.Usercontrol1 cmbAccName 
      Height          =   360
      Left            =   60
      TabIndex        =   0
      Top             =   450
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   635
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin VB.CommandButton cmdUpdateDeductionAmt 
      Caption         =   "Update Ded. Amount"
      Height          =   285
      Left            =   2340
      TabIndex        =   22
      Top             =   2280
      Visible         =   0   'False
      Width           =   1755
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
      Left            =   30
      TabIndex        =   6
      Top             =   3870
      Visible         =   0   'False
      Width           =   1956
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   360
      Left            =   4815
      TabIndex        =   1
      Top             =   450
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   167968771
      CurrentDate     =   37384
   End
   Begin MSForms.TextBox txtLTDed_New 
      Height          =   360
      Left            =   2190
      TabIndex        =   3
      Top             =   1920
      Width           =   2100
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "3704;635"
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "New Deduction Amount"
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
      Index           =   8
      Left            =   2205
      TabIndex        =   21
      Top             =   1665
      Width           =   1995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "S.T Amount to L.T"
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
      Index           =   7
      Left            =   4335
      TabIndex        =   20
      Top             =   1665
      Width           =   1575
   End
   Begin MSForms.TextBox txtSTToClear 
      Height          =   360
      Left            =   4320
      TabIndex        =   4
      Top             =   1920
      Width           =   2100
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "3704;635"
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "L.T Amount to S.T"
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
      Index           =   4
      Left            =   75
      TabIndex        =   19
      Top             =   1665
      Width           =   1575
   End
   Begin MSForms.TextBox txtLTToClear 
      Height          =   360
      Left            =   60
      TabIndex        =   2
      Top             =   1920
      Width           =   2100
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "3704;635"
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "L.T Deduction Amount"
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
      Index           =   3
      Left            =   2190
      TabIndex        =   18
      Top             =   930
      Width           =   1890
   End
   Begin MSForms.TextBox txtLTDed 
      Height          =   360
      Left            =   2190
      TabIndex        =   17
      Top             =   1200
      Width           =   2100
      VariousPropertyBits=   746604575
      BackColor       =   12648447
      Size            =   "3704;635"
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   555
      X2              =   6460
      Y1              =   4575
      Y2              =   4575
   End
   Begin MSForms.TextBox txtSTBalance 
      Height          =   360
      Left            =   4320
      TabIndex        =   16
      Top             =   1200
      Width           =   2100
      VariousPropertyBits=   746604575
      BackColor       =   12648447
      Size            =   "3704;635"
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txtRemarks 
      Height          =   1200
      Left            =   30
      TabIndex        =   5
      Top             =   2580
      Width           =   6465
      VariousPropertyBits=   -1467987941
      ScrollBars      =   3
      Size            =   "11404;2117"
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtLTBalance 
      Height          =   360
      Left            =   60
      TabIndex        =   15
      Top             =   1200
      Width           =   2100
      VariousPropertyBits=   746604575
      BackColor       =   12648447
      Size            =   "3704;635"
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Maker"
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
      Index           =   0
      Left            =   90
      TabIndex        =   14
      Top             =   225
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Remarks"
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
      Index           =   6
      Left            =   60
      TabIndex        =   13
      Top             =   2310
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Short Term Balance"
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
      Index           =   5
      Left            =   4335
      TabIndex        =   12
      Top             =   930
      Width           =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Long Term Balance"
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
      Index           =   2
      Left            =   75
      TabIndex        =   11
      Top             =   930
      Width           =   1650
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   435
      Left            =   2580
      TabIndex        =   8
      Top             =   4650
      Width           =   2100
      Caption         =   "Save & Close"
      PicturePosition =   327683
      Size            =   "3704;767"
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdNext 
      Height          =   435
      Left            =   540
      TabIndex        =   7
      Top             =   4650
      Width           =   1980
      Caption         =   "Save & New    "
      PicturePosition =   327683
      Size            =   "3492;767"
      Accelerator     =   78
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   435
      Left            =   4740
      TabIndex        =   9
      Top             =   4650
      Width           =   1770
      Caption         =   "Cancel    "
      PicturePosition =   327683
      Size            =   "3122;767"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
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
      Index           =   1
      Left            =   4845
      TabIndex        =   10
      Top             =   210
      Width           =   390
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   555
      X2              =   6460
      Y1              =   4590
      Y2              =   4590
   End
End
Attribute VB_Name = "frmMakerLoanType_Transfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lVenderID As Long
Dim lRequestEntryID  As Long

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


Private Function Saved() As Boolean

    Saved = False
    'On Error GoTo ERR
    Dim SQL As String
    
    If Not cmbAccName.MatchFound Then
        MsgBox "Please Select Valid Maker From The List.", vbInformation
        Exit Function
    ElseIf Val(txtLTToClear) = 0 And Val(txtSTToClear) = 0 Then
        MsgBox "Invalid Amounts to Transfer.", vbInformation
        Exit Function
    End If
    
    Dim lSTEntryID As Long, lLTEntryID As Long, lMLT_RefID As Long
    lSTEntryID = 0
    lLTEntryID = 0
    
    Dim strVchrNo As String
    
    Call StartTrans(con)
    
'    If chkPost.Value = vbChecked Then
'        Dim objVchr As New clsVouchers, strLoanClearance_AccNo As String, dTotalAmt As Double
'        strLoanClearance_AccNo = getGeneralDataValue("LoanClearance_AccNo")
'        dTotalAmt = Val(txtLTToClear) + Val(txtSTToClear)
'        objVchr.VoucherType = "JV"
'        objVchr.VoucherDate = DTPicker1.Value
'        objVchr.AddVoucherHead strLoanClearance_AccNo, dTotalAmt, txtRemarks
'        objVchr.AddVoucherHead cmbAccName.ID, -dTotalAmt, txtRemarks
'
'        strVchrNo = objVchr.PostVoucher(False, True)
'    End If
    
    con.Execute "INSERT INTO MakerLoanTransfer(AccNo_From,AccNo_To,ST_Amt,LT_Amt,UserName,MachineName,Description,DT) VALUES('" & cmbAccName.ID & "','" & cmbAccName.ID & "'," & _
     Val(txtSTToClear) & "," & Val(txtLTToClear) & ",'" & CurrentUserName & "','" & strComputerName & "','" & txtRemarks & "','" & DTPicker1.Value & "')"
    lMLT_RefID = GetSingleLongValue("MAX(EntryID)", "MakerLoanTransfer")
    If Val(txtSTToClear) > 0 Then
        con.Execute "INSERT INTO MakerPostedBills_STDeductions(MPB_RefID,Amount,AccNo,MLT_RefID) VALUES(0," & Val(txtSTToClear) & ",'" & cmbAccName.ID & "'," & lMLT_RefID & ")"
         
        con.Execute "INSERT INTO MakerAdvances(AccNo,DT,Type,Description,Amount,DAmount,AmountCleared,AccVoucherNo,MLT_RefID) VALUES('" & cmbAccName.ID & _
         "','" & DTPicker1 & "',1,'" & txtDesc & "'," & Val(txtSTToClear) & "," & Val(txtLTDed_New) & ",0,'" & strVchrNo & "'," & lMLT_RefID & ")"
         
        lSTEntryID = GetSingleLongValue("MAX(EntryID)", "MakerPostedBills_STDeductions")
    End If
    
    If Val(txtLTToClear) > 0 Then
        con.Execute "INSERT INTO MakerAmtCleared(AccNo,DT,AmtClrd,MLT_RefID) VALUES('" & cmbAccName.ID & "','" & DTPicker1.Value & "'," & Val(txtLTToClear) & "," & lMLT_RefID & ")"
        
        con.Execute "INSERT INTO MakerAdvancesShort(AccNo,DT,Description,Amount,AmountCleared,AccVoucherNo,MLT_RefID) VALUES('" & _
         cmbAccName.ID & "','" & DTPicker1 & "','" & txtDesc & "'," & Val(txtLTToClear) & ",0,'" & strVchrNo & "'," & lMLT_RefID & ")"
         
        lLTEntryID = GetSingleLongValue("MAX(EntryID)", "MakerAmtCleared")
    End If
    
    'con.Execute "INSERT INTO MakerLoanClearance_Manual(DT,AccNo,MPB_STD_EntryID,MAC_EntryID,UserName,MachineName,Remarks,VchrNo) " & _
     "VALUES('" & DTPicker1.Value & "','" & cmbAccName.ID & "'," & lSTEntryID & "," & lLTEntryID & ",'" & _
     CurrentUserName & "','" & strComputerName & "','" & txtRemarks & "','" & strVchrNo & "')"
    
    
    con.CommitTrans
    
    Saved = True
    
    
    SQL = ""
    Exit Function
err:
    MsgBox err.Description

End Function
Private Sub cmbAccName_matched()
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT TotalAmountTaken,TotalAmountCleared,ActualDeductionAmount FROM VMakerAdvancesDeduction WHERE AccNo='" & cmbAccName.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            txtLTBalance = "0"
            txtLTDed = "0"
            txtLTDed_New = "0"
        Else
            txtLTDed = Val(!ActualDeductionAmount & "")
            txtLTDed_New = Val(txtLTDed)
            txtLTBalance = Val(!TotalAmountTaken & "") - Val(!TotalAmountCleared & "") '- Val(!ActualDeductionAmount & "")
        End If
        .Close
        
        .Open "SELECT * FROM VMakerAdvancesShortDeduction WHERE AccNo='" & cmbAccName.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            txtSTBalance = Val(!Amount & "") - Val(!AmountCleared & "")
        Else
            txtSTBalance = 0
        End If
        .Close
    End With
    Set rs = Nothing
    
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

Private Sub cmdUpdateDeductionAmt_Click()
    If cmbAccName.MatchFound = False Then
        MsgBox "Please Select Maker.", vbInformation
        Exit Sub
    ElseIf Val(txtLTDed_New) < 0 Then
        MsgBox "Invalid Ded. Amount.", vbInformation
        Exit Sub
    End If
    
    If MsgBox("Are you sure?", vbYesNo + vbQuestion) = vbYes Then
        con.Execute "UPDATE MakerAdvances SET DAmount=" & Val(txtLTDed_New) & " WHERE EntryID IN(SELECT MAX(EntryID) FROM MakerAdvances WHERE AccNo='" & cmbAccName.ID & "')"
    End If
    Call cmbAccName_matched
End Sub

Private Sub Form_Load()

    cmbAccName.ListHeight = 3000
    DTPicker1 = GetServerDate
    
End Sub

Private Sub txtAmt_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyNumber(KeyAscii.Value, True)
End Sub

Private Sub txtDeduct_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyNumber(KeyAscii.Value, True)
End Sub

Public Sub ShowMe(Optional str_MakerAccNo As String)
        
    With cmbAccName
        .ClearVals
        .AddVals con, "'{' + VendID1 + '} ' + AccTitle", "VMakers", "AccNo", " ORDER BY AccNo"
    End With
    
    If str_MakerAccNo <> "" Then
        cmbAccName.ID = str_MakerAccNo
    End If
    
    Me.Show 1
    
End Sub

Private Sub txtLTToClear_Validate(cancel As Boolean)

    If Val(txtLTToClear) > Val(txtLTBalance) Then
        MsgBox "Invalid L.T Amount.", vbInformation
        cancel = True
    End If
    
End Sub

Private Sub txtSTToClear_Validate(cancel As Boolean)
    If Val(txtSTToClear) > Val(txtSTBalance) Then
        MsgBox "Invalid S.T Amount.", vbInformation
        cancel = True
    End If
End Sub
