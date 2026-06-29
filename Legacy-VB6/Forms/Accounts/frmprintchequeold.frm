VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrintChequeOld 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cheque Printing"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6315
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DT 
      Height          =   315
      Left            =   2655
      TabIndex        =   7
      Top             =   1365
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   53149699
      CurrentDate     =   39701
   End
   Begin VB.ComboBox cmbDesignation 
      Height          =   315
      ItemData        =   "frmprintchequeold.frx":0000
      Left            =   165
      List            =   "frmprintchequeold.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1365
      Width           =   2460
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print Cheque"
      Height          =   390
      Left            =   3585
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   4860
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtPay 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   165
      TabIndex        =   0
      Top             =   720
      Width           =   5895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cheque Date :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   2700
      TabIndex        =   6
      Top             =   1125
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Designation :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   195
      TabIndex        =   4
      Top             =   1125
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Pay :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   195
      TabIndex        =   1
      Top             =   465
      Width           =   975
   End
End
Attribute VB_Name = "frmPrintChequeOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_dSNo As Double
Dim iBankID As Integer
Dim bUnAuthorized As Boolean
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
    
'    If iBankID = 7 Then
'        Set rpt = rptApp.OpenReport(rptAccountsPath & "\rptABNAmroChq.rpt")
'    ElseIf iBankID = 8 Or iBankID = 9 Then
'        Set rpt = rptApp.OpenReport(rptAccountsPath & "\rptUBLChq.rpt")
'    Else
'        MsgBox "Bank Not Supported.", vbInformation
'        Exit Sub
'    End If
    If bUnAuthorized = False Then
        Set rpt = rptApp.OpenReport(RptPath & "\rptChq.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\rptChqUA.rpt")
    End If
    
    Dim f As New frmPrevRpt
    rpt.FormulaFields.GetItemByName("Pay").Text = "'" & txtPay.Text & "'"
    rpt.FormulaFields.GetItemByName("Designation").Text = "'" & cmbDesignation.Text & "'"
    rpt.FormulaFields.GetItemByName("PrintDT").Text = "#" & DT.Value & "#"
    Me.Hide
    f.ShowReport "{VChqLedger.SNo}=" & m_dSNo & " AND {VLedger.Debit}>0", rpt, , , "ChqPaper"
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
    With rs
        .Open "SELECT * FROM VChqLedger Where SNo=" & m_dSNo, con, adOpenForwardOnly, adLockReadOnly
'        strMakerAccTitle = GetSingleStringValue("BankAccTitle + ' ' + BankAccNo", "Makers", " WHERE AccNo='" & !AccNo & "'")
'        strMakerAccTitle = Trim(strMakerAccTitle)
        iBankID = !BankID
        strChqType = !ChequeType & ""
        If strMakerAccTitle = "" Then
            strAccTitle = !AccTitle & ""
        Else
            strAccTitle = strMakerAccTitle
        End If
        DT.Value = !ChequeDate
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
    cmbDesignation.ListIndex = 0
    
    Me.Show 1
    
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_Load()
    bUnAuthorized = False
End Sub
