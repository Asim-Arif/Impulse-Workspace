VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrintCheque 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cheque Printing"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6195
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHandedOverTo 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   165
      TabIndex        =   2
      Top             =   2265
      Width           =   5895
   End
   Begin VB.TextBox txtCNICNo 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   165
      TabIndex        =   1
      Top             =   1485
      Width           =   5895
   End
   Begin VB.ComboBox cmbCompany 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "frmPrintCheque.frx":0000
      Left            =   1830
      List            =   "frmPrintCheque.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3120
      Width           =   2640
   End
   Begin MSComCtl2.DTPicker DT 
      Height          =   420
      Left            =   4470
      TabIndex        =   9
      Top             =   3120
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   157089795
      CurrentDate     =   39701
   End
   Begin VB.ComboBox cmbDesignation 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "frmPrintCheque.frx":0029
      Left            =   165
      List            =   "frmPrintCheque.frx":0033
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3120
      Width           =   1680
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print Cheque"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2835
      TabIndex        =   5
      Top             =   3930
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4485
      TabIndex        =   4
      Top             =   3930
      Width           =   1575
   End
   Begin VB.TextBox txtPay 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   165
      TabIndex        =   0
      Top             =   720
      Width           =   5895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Handed Over To"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   195
      TabIndex        =   13
      Top             =   1980
      Width           =   1515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CNIC No."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   195
      TabIndex        =   12
      Top             =   1200
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Company"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1860
      TabIndex        =   11
      Top             =   2865
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cheque Date"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4515
      TabIndex        =   8
      Top             =   2865
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Designation"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   195
      TabIndex        =   6
      Top             =   2865
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Payee"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   195
      TabIndex        =   3
      Top             =   435
      Width           =   540
   End
End
Attribute VB_Name = "frmPrintCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_dSNo As Double
Dim iBankID As Integer

Dim rptName As String, bDuplicate As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    On Error Resume Next
    
     
    Dim rpt As CRAXDDRT.Report
    If Format(DT.Value, "dddd") = "Sunday" Then
        MsgBox "Invalid Date, Can't Accept Sunday.", vbInformation
        Me.Show
        Exit Sub
    End If
    
    If txtPay = "" Then
        MsgBox "Please Enter Payee", vbInformation
        Exit Sub
    End If
    If txtHandedOverTo = "" Then
        MsgBox "Please Enter Handed Over To", vbInformation
        Exit Sub
    End If
'    If iBankID = 7 Then
'        Set rpt = rptApp.OpenReport(rptAccountsPath & "\rptABNAmroChq.rpt")
'    ElseIf iBankID = 8 Or iBankID = 9 Then
'        Set rpt = rptApp.OpenReport(rptAccountsPath & "\rptUBLChq.rpt")
'    Else
'        MsgBox "Bank Not Supported.", vbInformation
'        Exit Sub
'    End If

    If rptName = "" Then
        rptName = "rptChq.rpt"
    Else
        rptName = Replace(rptName, ".rpt", "") & ".rpt"
    End If
    
    Set rpt = rptApp.OpenReport(RptPath & "\" & rptName)
    
    Dim f As New frmPrevRpt
    rpt.FormulaFields.GetItemByName("Pay").Text = "'" & txtPay.Text & "'"
    rpt.FormulaFields.GetItemByName("Designation").Text = "'" & IIf(cmbDesignation.Text = "-", "", cmbDesignation.Text) & "'"
    rpt.FormulaFields.GetItemByName("Company").Text = "'" & IIf(cmbCompany.Text = "-", "", cmbCompany.Text) & "'"
    rpt.FormulaFields.GetItemByName("PrintDT").Text = "#" & DT.Value & "#"
    
    On Error Resume Next
    rpt.FormulaFields.GetItemByName("Duplicate").Text = bDuplicate
    On Error GoTo 0
    Me.Hide
         
         
'    If rptName = "rptChqCiti.rpt" Then
'        f.ShowReport "{VChqLedger.SNo}=" & m_dSNo & " AND {VLedger.Debit}>0 AND {VLedger.VchrNo}='" & m_strVchrNo & "'", rpt, , , "ChqPaperCITI"
'    Else
'        f.ShowReport "{VChqLedger.SNo}=" & m_dSNo & " AND {VLedger.Debit}>0 AND {VLedger.VchrNo}='" & m_strVchrNo & "'", rpt, , , "ChqPaper"
'    End If
    If strCompany = "Banzai" Then
        f.ShowReport "{VChqLedger.SNo}=" & m_dSNo, rpt
    Else
        f.ShowReport "{VChqLedger.SNo}=" & m_dSNo & " AND {VLedger.Debit}>0", rpt
    End If
    
    con.Execute "UPDATE VoucherInfo SET Handed_Over_To='" & txtHandedOverTo & "' WHERE VchrNo='" & txtHandedOverTo.Tag & "'"
    con.Execute "UPDATE Cheque SET ChqPrintingDone=1 WHERE SNo=" & m_dSNo
    
    Unload Me
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Public Sub ShowMe(p_dSNo As Double, Optional p_bUnAuthorized As Boolean = False)

    m_dSNo = p_dSNo
    bUnAuthorized = p_bUnAuthorized
    
    Dim rs As New ADODB.Recordset
    Dim strMakerAccTitle As String
    Dim strChqType As String
    Dim strAccTitle As String
    Dim bPrintingDone As Boolean
    With rs
        .Open "SELECT * FROM VChqLedger WHERE SNo=" & m_dSNo, con, adOpenForwardOnly, adLockReadOnly
'        strMakerAccTitle = GetSingleStringValue("BankAccTitle + ' ' + BankAccNo", "Makers", " WHERE AccNo='" & !AccNo & "'")
'        strMakerAccTitle = Trim(strMakerAccTitle)
        bPrintingDone = !ChqPrintingDone
        iBankID = !BankID
        strChqType = !ChequeType & ""
        If strMakerAccTitle = "" Then
            strAccTitle = !Payee & ""
        Else
            strAccTitle = strMakerAccTitle
        End If
        DT.Value = !ChequeDate
        rptName = !ChqFormat & ""
        txtHandedOverTo.Text = !Handed_Over_To & ""
        txtHandedOverTo.Tag = !VchrNo & ""
        txtCNICNo = !NICNo & ""
        .Close
    End With
    Set rs = Nothing
    
    DT.MinDate = DT.Value
    DT.MaxDate = DateAdd("d", 7, DT.Value)
    
'    If strChqType = "Open" Then
'        txtPay = "Self"
'    Else
'        txtPay = strAccTitle
'    End If
    txtPay = strAccTitle
    'cmbDesignation.ListIndex = 0
    cmbDesignation.Clear
    cmbDesignation.AddItem "-"
    Call AddToCombo(cmbDesignation, "Designations", "DesignationsForChqPrinting", , True)
    
    cmbCompany.Clear
    cmbCompany.AddItem "-"
    Call AddToCombo(cmbCompany, "CompanyName", "CompaniesForChqPrinting", , True)
    
    cmbCompany.ListIndex = 0
    cmbDesignation.ListIndex = 0
    
    bDuplicate = bPrintingDone
    
    If bPrintingDone Then
        If MsgBox("Printing already done for this cheque, do you want to Print Duplicate?", vbQuestion + vbYesNo) = vbYes Then
            If getDBPassword("DuplicateChqPrinting") = False Then
                Unload Me
                Exit Sub
            End If
        Else
            Unload Me
            Exit Sub
        End If
    End If
    
    Me.Show 1
    
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_Load()
    bUnAuthorized = False
End Sub
