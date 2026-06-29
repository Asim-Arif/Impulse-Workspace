VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmFCustRcvPymnts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Receive Payments From Customers"
   ClientHeight    =   3375
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBankCharges 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2265
      MaxLength       =   150
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2025
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.TextBox txtExchRate 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   5220
      MaxLength       =   150
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2040
      Width           =   810
   End
   Begin VB.TextBox txtFDBC 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   60
      MaxLength       =   150
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1710
      Width           =   2145
   End
   Begin VB.TextBox txtAccNo 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   4290
      MaxLength       =   150
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1035
      Visible         =   0   'False
      Width           =   1740
   End
   Begin MSComCtl2.DTPicker DT1 
      Height          =   300
      Left            =   4245
      TabIndex        =   2
      Top             =   360
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   101908483
      CurrentDate     =   37928
   End
   Begin ComboList.Usercontrol1 cmbCust 
      Height          =   285
      Left            =   45
      TabIndex        =   0
      Top             =   375
      Width           =   2220
      _ExtentX        =   3916
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
   Begin VB.TextBox txtAmt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   4950
      MaxLength       =   150
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1710
      Width           =   1080
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   705
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2610
      Width           =   4170
   End
   Begin VB.TextBox txtCurrency 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2220
      MaxLength       =   150
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1710
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C5D1DC&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Exchange Rate :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   3
      Left            =   3570
      TabIndex        =   19
      Top             =   2040
      Width           =   1620
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C5D1DC&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Bank"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   16
      Top             =   765
      Width           =   5970
   End
   Begin MSForms.ComboBox cmbBankName 
      Height          =   285
      Left            =   60
      TabIndex        =   15
      Top             =   1035
      Width           =   5970
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "10530;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbBranch 
      Height          =   285
      Left            =   2175
      TabIndex        =   14
      Top             =   1035
      Visible         =   0   'False
      Width           =   2085
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3678;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbCustomInvoice 
      Height          =   285
      Left            =   3090
      TabIndex        =   4
      Top             =   1710
      Width           =   1845
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3254;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbCountry 
      Height          =   285
      Left            =   2295
      TabIndex        =   1
      Top             =   375
      Width           =   1920
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3387;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C5D1DC&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   1
      Left            =   45
      TabIndex        =   11
      Top             =   2340
      Width           =   4170
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C5D1DC&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FDBC No.                                Curr           Invoice No                      Amount "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   10
      Top             =   1440
      Width           =   5970
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C5D1DC&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Customer                               Country                            Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   7
      Left            =   45
      TabIndex        =   9
      Top             =   105
      Width           =   5970
   End
   Begin MSForms.CommandButton cmdClose 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   4275
      TabIndex        =   8
      Top             =   2955
      Width           =   1755
      ForeColor       =   0
      Caption         =   "Cancel          "
      PicturePosition =   327683
      Size            =   "3096;635"
      Picture         =   "frmFCustRcvPymnts.frx":0000
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdGenerate 
      Height          =   360
      Left            =   4275
      TabIndex        =   7
      Top             =   2565
      Width           =   1755
      ForeColor       =   0
      Caption         =   "Receive     "
      PicturePosition =   327683
      Size            =   "3096;635"
      Accelerator     =   82
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmFCustRcvPymnts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RcvCustom As Boolean

Private Sub cmbCust_matched()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Country,Curr From ForeignCustomers Where CustCode='" & cmbCust.id & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            cmbCountry.List(cmbCountry.ListCount - 1, 1) = ![Curr] & ""
            .MoveNext
        Loop
        .Close
    End With
    
    If cmbCountry.ListCount > 0 Then cmbCountry.ListIndex = 0
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub cmbCustomInvoice_Change()
    Call cmbCustomInvoice_Click
End Sub

Private Sub cmbCustomInvoice_Click()
    
    txtAmt.Tag = cmbCustomInvoice.List(cmbCustomInvoice.ListIndex, 1)
    txtAmt.Text = txtAmt.Tag
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
    
    On Error GoTo err
    
    Dim RecAffect As Long
    
    If cmbCustomInvoice.MatchFound = False Then
        MsgBox "Select Valid Invoice.", vbInformation
        
        Exit Sub
    End If
    
    If cmbBankName.MatchFound = False Then
        MsgBox "Please Select Bank.", vbInformation
        Exit Sub
    End If
    
    If Val(txtAmt) + Val(txtBankCharges) > Val(txtAmt.Tag) Or Val(txtAmt) + Val(txtBankCharges) <= 0 Then
        MsgBox "Invalid Amount Entered.", vbInformation
        cancel = True
        Exit Sub
    ElseIf Val(txtExchRate) <= 0 Then
        MsgBox "Invalid Exchange Rate.", vbInformation
        Exit Sub
    End If
    
    Call StartTrans(con)
    
    Dim TableName As String
    If RcvCustom Then
        TableName = "FCustPayments" '"CustomPayments"
    Else
        TableName = "FCustPayments"
    End If
    
    'cmbBranch.List(cmbBranch.ListIndex, 2)
    
    con.Execute "Insert Into " & TableName & "(CustomInvoice,DT,AmtRcvd,Description" & _
     ",CustomPayment,FDBCNo,BankID,ExchRate,AmountEE) VALUES('" & _
     cmbCustomInvoice.Text & "','" & DT1.Value & "'," & Val(txtAmt) & ",'" & _
     txtDescription & "'," & Abs(RcvCustom) & ",'" & txtFDBC & "'," & _
     Val(cmbBankName.List(cmbBankName.ListIndex, 1)) & "," & Val(txtExchRate) & "," & Val(txtAmt) & ")"
    
'    If RcvCustom Then
'        con.Execute "Update CustomInvoice Set TotalCustomAmtToRcv=TotalCustomAmtToRcv-" & _
'         Val(txtAmt) + Val(txtBankCharges) & " Where CustomInvoice='" & cmbCustomInvoice & "'"
'    Else
'        con.Execute "Update CustomInvoice Set TotalCommercialAmtToRcv=TotalCommercialAmtToRcv-" & _
'         Val(txtAmt) + Val(txtBankCharges) & " Where CustomInvoice='" & cmbCustomInvoice & "'"
'    End If
    
    con.CommitTrans
    
    Unload Me
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub Form_Load()
        
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode")
'    Call AddToCombo(cmbBankName, "Distinct Bank", "BankList", "Order By Bank")
    
    cmbCust.ListHeight = 2500
    
    RcvCustom = False
    DT1 = ServerDate
    
End Sub
Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()
    
    On Error GoTo err
    If cmbCountry.MatchFound = False Then Exit Sub
    
    cmbCountry.Tag = cmbCust.Text & cmbCountry
    txtCurrency = cmbCountry.List(cmbCountry.ListIndex, 1)
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    If RcvCustom Then
        SQL = "Select CustomInvoice,TotalCustomAmtToRcv From VrptCustomInvoice Where CustCode='" & cmbCust.Text & "' And Country='" & cmbCountry.Text & "' And TotalCustomAmtToRcv>0 Order By CustomInvoice"
    Else
        SQL = "Select CustomInvoice,TotalCommercialAmtToRcv From VrptCustomInvoice Where CustCode='" & cmbCust.Text & "' And Country='" & cmbCountry.Text & "' AND TotalCommercialAmtToRcv>0 Order By CustomInvoice"
    End If
    
    With rs
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        cmbCustomInvoice.Clear
        Do Until .EOF
            cmbCustomInvoice.AddItem .Fields(0) & ""
            cmbCustomInvoice.List(cmbCustomInvoice.ListCount - 1, 1) = .Fields(1) & ""
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub txtAmt_Validate(cancel As Boolean)

    If Val(txtAmt) > Val(txtAmt.Tag) Or Val(txtAmt) < 0 Then
        MsgBox "Invalid Amount Entered.", vbInformation
        cancel = True
        Exit Sub
    ElseIf Val(txtBankCharges) < 0 Then
        MsgBox "Invalid Amount Entered.", vbInformation
        cancel = True
        Exit Sub
    End If
    
End Sub

Private Sub txtDescription_Validate(cancel As Boolean)
    txtDescription = Replace(txtDescription, "|", "")
End Sub

Public Sub ShowForm(CustCode As String, InvoiceNo As String, Optional TempRcvCustom = False)

    On Error GoTo err
    RcvCustom = TempRcvCustom
    If RcvCustom Then Me.Caption = "Receive Custom Payments"
    
    cmbCust.id = CustCode
    cmbCustomInvoice = InvoiceNo
    
    txtAmt.TabIndex = 0
     
    Dim lBankID As Long, iIndex As Integer
    lBankID = GetSingleLongValue("OurBankID", "CustomInvoice", " WHERE CustomInvoice='" & InvoiceNo & "'")
    Dim rs As New ADODB.Recordset
    iIndex = 0
    With rs
        '.Open "SELECT Bank,Branch,AccNo,BankAccNo FROM VBankAccounts WHERE BankID=" & lBankID, con, adOpenForwardOnly, adLockReadOnly
        .Open "SELECT Bank,Branch,AccNo,BankAccNo,BankID FROM VBankAccounts", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
        
            cmbBankName.AddItem !Bank & " (" & !Branch & ") (" & !BankAccNo & ")"
            cmbBankName.List(iIndex, 1) = Val(!BankID & "")
            cmbBankName.List(iIndex, 2) = !Branch & ""
            cmbBankName.List(iIndex, 3) = !BankAccNo & ""
            
            iIndex = iIndex + 1
            .MoveNext
        Loop
        
        .Close
    End With
    Set rs = Nothing
    
    Dim i As Integer
    
    For i = 0 To cmbBankName.ListCount - 1
        If cmbBankName.List(i, 1) = lBankID Then
            cmbBankName.ListIndex = i
            Exit For
        End If
    Next
    
    Me.Show 1
    
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

'Private Sub cmbBankName_Change()
'    Call cmbBankName_Click
'End Sub
'
'Private Sub cmbBankName_Click()
'
'    On Error GoTo err
'
'    If cmbBankName.MatchFound = False Then Exit Sub
'
'    Dim rs As New ADODB.Recordset
'    With rs
'        .Open "Select BankID,Branch,AccNo From BankList Where Bank='" & cmbBankName & "'", con, adOpenForwardOnly, adLockReadOnly
'        cmbBranch.Clear
'        Do Until .EOF
'            cmbBranch.AddItem ![Branch] & ""
'
'            cmbBranch.List(cmbBranch.ListCount - 1, 1) = ![AccNo] & ""
'            cmbBranch.List(cmbBranch.ListCount - 1, 2) = ![BankID] & ""
'            .MoveNext
'        Loop
'        .Close
'    End With
'
'    Set rs = Nothing
'
'    Exit Sub
'err:
'    MsgBox err.Description
'
'End Sub

'Private Sub cmbBranch_Change()
'    Call cmbBranch_Click
'End Sub
'
'Private Sub cmbBranch_Click()
'    txtAccNo.Text = cmbBranch.List(cmbBranch.ListIndex, 1)
'End Sub
'
