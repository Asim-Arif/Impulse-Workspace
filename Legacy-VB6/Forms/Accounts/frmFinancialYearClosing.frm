VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFinancialYearClosing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Financial Year Closing...."
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10305
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   10305
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtClosingStockWorkInProcess 
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
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   6855
      TabIndex        =   3
      Top             =   1125
      Width           =   3315
   End
   Begin VB.TextBox txtClosingStockRawMaterial 
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
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   3510
      TabIndex        =   2
      Top             =   1125
      Width           =   3315
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   480
      Left            =   8355
      TabIndex        =   15
      Top             =   9045
      Width           =   1920
   End
   Begin VB.Frame Frame3 
      Caption         =   "Closing Stock :"
      Height          =   4290
      Left            =   135
      TabIndex        =   12
      Top             =   4695
      Width           =   10125
      Begin MSComctlLib.ListView LVYC 
         Height          =   3960
         Left            =   90
         TabIndex        =   13
         Top             =   240
         Width           =   9930
         _ExtentX        =   17515
         _ExtentY        =   6985
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   15199215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Account"
            Object.Width           =   7408
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Debit"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Credit"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Balance"
            Object.Width           =   3175
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Closing Stock :"
      Height          =   1215
      Left            =   135
      TabIndex        =   10
      Top             =   3405
      Width           =   10125
      Begin MSComctlLib.ListView LVSE 
         Height          =   900
         Left            =   90
         TabIndex        =   11
         Top             =   225
         Width           =   9930
         _ExtentX        =   17515
         _ExtentY        =   1588
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   15199215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Account"
            Object.Width           =   7699
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Debit"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Credit"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Balance"
            Object.Width           =   3175
         EndProperty
      End
   End
   Begin VB.CommandButton cmdViewOnly 
      Caption         =   "Generate Vouchers"
      Height          =   480
      Left            =   4395
      TabIndex        =   9
      Top             =   9045
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Frame Frame1 
      Caption         =   "Stock Adjustment Entry :"
      Height          =   1785
      Left            =   135
      TabIndex        =   7
      Top             =   1605
      Width           =   10125
      Begin MSComctlLib.ListView LVSAE 
         Height          =   1395
         Left            =   90
         TabIndex        =   8
         Top             =   285
         Width           =   9930
         _ExtentX        =   17515
         _ExtentY        =   2461
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   15199215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Account"
            Object.Width           =   7699
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Debit"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Credit"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Balance"
            Object.Width           =   3175
         EndProperty
      End
   End
   Begin VB.CommandButton cmdCloseYear 
      Caption         =   "Post Vouchers"
      Height          =   480
      Left            =   6375
      TabIndex        =   6
      Top             =   9045
      Width           =   1920
   End
   Begin VB.TextBox txtClosingStock 
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
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   165
      TabIndex        =   1
      Top             =   1125
      Width           =   3315
   End
   Begin VB.ComboBox cmbFinYear 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   330
      Width           =   3810
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   330
      Left            =   120
      TabIndex        =   14
      Top             =   9075
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Closing Stock (Work In Process)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   6855
      TabIndex        =   17
      Top             =   900
      Width           =   2700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Closing Stock (Raw Material)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   3510
      TabIndex        =   16
      Top             =   885
      Width           =   2430
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Closing Stock (Finished)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   5
      Top             =   870
      Width           =   2010
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Financial Year :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   4
      Top             =   375
      Width           =   1260
   End
End
Attribute VB_Name = "frmFinancialYearClosing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DTStart As Date
Dim DTFinish As Date
Dim m_dCostofSalesBal As Double
Dim dCostofSalesBal As Double

Dim strCostOfSalesAccNo As String, strProfitANDLossAccNo As String
Dim strStockAccNoFinished As String, strStockAccNoRawMaterial As String, strStockAccNoWorkInProcess As String

Private Sub cmbFinYear_Change()
    If cmbFinYear.ListIndex <> -1 Then
        If cmdCloseYear.Enabled = False Then
            cmdCloseYear.Enabled = True
            cmdViewOnly.Enabled = True
        End If
    End If
    LVSAE.ListItems.Clear
    LVSE.ListItems.Clear
    LVYC.ListItems.Clear
    
End Sub

Private Sub cmbFinYear_Click()
    Call cmbFinYear_Change
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCloseYear_Click()
    
    On Error GoTo err
    If cmbFinYear.ListIndex = -1 Then Exit Sub
     If Val(txtClosingStock) <= 0 Then
        MsgBox "Invalid Closing Stock Amount.", vbInformation
        Exit Sub
    End If
    
    If MsgBox("Are you Sure?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Dim iYear As Integer
    iYear = cmbFinYear.ItemData(cmbFinYear.ListIndex)
    DTStart = DateSerial(iYear, 7, 1)
    DTFinish = DateSerial(iYear + 1, 6, 30)
    'First Add Stock Adjustment Voucher....
    
    LVSAE.ListItems.Clear
    LVSE.ListItems.Clear
    LVYC.ListItems.Clear
    
    Dim strStockAdjVchrNo As String, strClosingStockVchrNo As String, strYearClosingVchrNo As String
    
    strStockAdjVchrNo = AddClosingStockAdjustmentVoucher(False)
    strClosingStockVchrNo = AddClosingStockVoucher(False)
    strYearClosingVchrNo = AddYearClosingVoucher(False)
    
    con.Execute "INSERT INTO FinancialClosings(DTStart,DTFinish,UserName,MachineName,StockAdjustmentVchr," & _
     "StockClosingVchr,YearClosingVchr,ClosingStockAmtFinished,ClosingStockAmtRawMaterial,ClosingStockAmtWIP) VALUES('" & DTStart & "','" & DTFinish & "','" & _
     CurrentUserName & "','" & strComputerName & "','" & strStockAdjVchrNo & "','" & strClosingStockVchrNo & _
     "','" & strYearClosingVchrNo & "'," & Val(txtClosingStock) & "," & Val(txtClosingStockRawMaterial) & "," & Val(txtClosingStockWorkInProcess) & ")"
    
    MsgBox "Year Closed.", vbInformation
   
    Call FillCmbs
      
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdViewOnly_Click()
    
    If Val(txtClosingStock) <= 0 Then
        MsgBox "Invalid Closing Stock Amount.", vbInformation
        Exit Sub
    End If
    
    Dim iYear As Integer
    iYear = cmbFinYear.ItemData(cmbFinYear.ListIndex)
    DTStart = DateSerial(iYear, 7, 1)
    DTFinish = DateSerial(iYear + 1, 6, 30)
    LVSAE.ListItems.Clear
    LVSE.ListItems.Clear
    LVYC.ListItems.Clear
    'First Add Stock Adjustment Voucher....
    Dim strClosingStockVchrNo As String
    AddClosingStockAdjustmentVoucher (True)
    AddClosingStockVoucher (True)
    AddYearClosingVoucher (True)
    
    Call Command1_Click 'To Sort
End Sub

Private Sub Command1_Click()
    
    Dim j As Integer, ITM As ListItem, ITMtemp As ListItem
    Dim strAccTitle As String, dDebit As String, dCredit As String, dBalance As String
    For i = 1 To LVYC.ListItems.count
        Set ITM = LVYC.ListItems(i)
        For j = i + 1 To LVYC.ListItems.count
            Set ITMtemp = LVYC.ListItems(j)
            If Val(ITM.SubItems(1)) > 0 Then
                If Val(ITM.SubItems(1)) < Val(ITMtemp.SubItems(1)) Then
                    strAccTitle = ITM.Text
                    dDebit = ITM.SubItems(1)
                    dCredit = ITM.SubItems(2)
                    dBalance = ITM.SubItems(3)
                    ITM.Text = ITMtemp.Text
                    ITM.SubItems(1) = ITMtemp.SubItems(1)
                    ITM.SubItems(2) = ITMtemp.SubItems(2)
                    ITM.SubItems(3) = ITMtemp.SubItems(3)
                    ITMtemp.Text = strAccTitle
                    ITMtemp.SubItems(1) = dDebit
                    ITMtemp.SubItems(2) = dCredit
                    ITMtemp.SubItems(3) = dBalance
                End If
            Else
                If Val(ITM.SubItems(2)) > Val(ITMtemp.SubItems(2)) Then
                    strAccTitle = ITM.Text
                    dDebit = ITM.SubItems(1)
                    dCredit = ITM.SubItems(2)
                    dBalance = ITM.SubItems(3)
                    ITM.Text = ITMtemp.Text
                    ITM.SubItems(1) = ITMtemp.SubItems(1)
                    ITM.SubItems(2) = ITMtemp.SubItems(2)
                    ITM.SubItems(3) = ITMtemp.SubItems(3)
                    ITMtemp.Text = strAccTitle
                    ITMtemp.SubItems(1) = dDebit
                    ITMtemp.SubItems(2) = dCredit
                    ITMtemp.SubItems(3) = dBalance
                End If
            End If
        Next
    Next
    
    Dim lColor As Long
    For i = 1 To LVYC.ListItems.count
        With LVYC.ListItems(i)
            If Val(.SubItems(1)) <= 0 Then lColor = vbRed Else lColor = vbBlue
'            .ForeColor = lColor
            For j = 1 To .ListSubItems.count
                .ListSubItems(j).ForeColor = lColor
            Next
        End With
    Next
    LVYC.Refresh
End Sub

Private Sub Form_Load()

    Call FillCmbs
    
    strCostOfSalesAccNo = getGeneralDataValue("Closing-CostOfSalesAccNo")
    
    strStockAccNoFinished = getGeneralDataValue("Closing-StockAccNo-Finished")
    strStockAccNoRawMaterial = getGeneralDataValue("Closing-StockAccNo-RawMaterial")
    strStockAccNoWorkInProcess = getGeneralDataValue("Closing-StockAccNo-WorkInProcess")
    
    strProfitANDLossAccNo = getGeneralDataValue("Closing-P&L")
    
    If strCostOfSalesAccNo = "" Or strStockAccNoFinished = "" Or strProfitANDLossAccNo = "" Or strStockAccNoRawMaterial = "" Or strStockAccNoWorkInProcess = "" Then
        MsgBox "Can't Find Closing Account(s) in General Data.", vbCritical
        cmdCloseYear.Enabled = False
    End If
    
End Sub

Private Sub FillCmbs()

    cmbFinYear.Clear
    Dim iStartYear As Integer
    Dim rs As New ADODB.Recordset
        
    With rs
        .Open "SELECT MAX(DTFinish) FROM FinancialClosings", con, adOpenForwardOnly, adLockReadOnly
        If IsNull(.Fields(0)) Then
            iStartYear = 2005
        Else
            iStartYear = year(.Fields(0))
        End If
    End With
    
    For i = iStartYear To year(Date)
        cmbFinYear.AddItem i & " x " & i + 1 & " [ " & Format(DateSerial(i, 7, 1), "dd-MMM-yyyy") & " to " & Format(DateSerial(i + 1, 6, 30), "dd-MMM-yyyy") & " ]"
        cmbFinYear.ItemData(cmbFinYear.ListCount - 1) = i
    Next i
    
End Sub

Private Function AddClosingStockAdjustmentVoucher(Optional bViewOnly As Boolean = True) As String

    'Store Null To Return on an Unsuccessful Entry
    Dim VNo As String, m_DTVchr As Date
    m_DTVchr = DTFinish
    
    AddVoucher = ""
    Dim i As Integer
    Dim dAmt As Double, dTotalAmt As Double
    dTotalAmt = 0
   
    'Account Nos to be Defined....
    'Credit Entries
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim Bal As Double, Desc As String
    
    Desc = "Closing Stock Adjustment"
    
    VNo = getNextVchrNo(m_DTVchr, "CL")
    INextSno = getNextSno(m_DTVchr)
    
    Dim rs As New ADODB.Recordset
    Dim strAccNo As String
    Dim dTotalAmtP As Double
    Dim dAmtDX As Double, dTotalAmtDX As Double
    Dim dAmtFO As Double, dTotalAmtFO As Double
    
    With rs
        'Purchases
        .Open "Select AccNo FROM Accounts WHERE LEFT(AccNo,7) IN('41-001-') AND Parent=0", con
        Do Until .EOF
            strAccNo = !AccNo
            Bal = getBalance(strAccNo, m_DTVchr)
            dAmt = -Bal
            dTotalAmtP = dTotalAmtP + Abs(dAmt)
            Bal = 0
            
            If bViewOnly = False And dAmt <> 0 Then
                con.Execute "INSERT INTO Vouchers VALUES(" & INextSno & ",'" & m_DTVchr & "','" & _
                 VNo & "','" & strAccNo & "','" & Desc & "'," & _
                 "0," & Abs(dAmt) & "," & Bal & ",'',0)"
                    
                Call adjustLedger(strAccNo, m_DTVchr, INextSno, dAmt)
            End If
             INextSno = INextSno + 1
            .MoveNext
        Loop
        .Close
        
        'Dirext Expenses; DX
'        .Open "Select AccNo FROM Accounts WHERE LEFT(AccNo,6)='41-003' AND Parent=0", con
'        Do Until .EOF
'            strAccNo = !AccNo
'            Bal = getBalance(strAccNo, m_DTVchr)
''            If Bal < 0 Then Stop
'            dAmt = -Bal
'            dTotalAmtDX = dTotalAmtDX - dAmt
'            'dTotalAmtDX = dTotalAmtDX + Abs(dAmt)
'            Bal = 0
'
'            If bViewOnly = False Then
'                If dAmt < 0 Then
'                    con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & m_DTVchr & "','" & _
'                     VNo & "','" & strAccNo & "','" & Desc & "'," & _
'                     "0," & Abs(dAmt) & "," & Bal & ",'',0)"
'                Else
'                    con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & m_DTVchr & "','" & _
'                     VNo & "','" & strAccNo & "','" & Desc & "'," & _
'                     dAmt & ",0," & Bal & ",'',0)"
'                End If
'
'                Call adjustLedger(strAccNo, m_DTVchr, INextSno, dAmt)
'            End If
'             INextSno = INextSno + 1
'            .MoveNEXT
'        Loop
'        .Close
'
'         'Factory Overheads; FO
'        .Open "Select AccNo FROM Accounts WHERE LEFT(AccNo,6)='41-004' AND Parent=0", con
'        Do Until .EOF
'            strAccNo = !AccNo
'            Bal = getBalance(strAccNo, m_DTVchr)
'            dAmt = -Bal
'            dTotalAmtFO = dTotalAmtFO + Abs(dAmt)
'            Bal = 0
'
'            If bViewOnly = False Then
'                con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & m_DTVchr & "','" & _
'                 VNo & "','" & strAccNo & "','" & Desc & "'," & _
'                 "0," & Abs(dAmt) & "," & Bal & ",'',0)"
'
'                Call adjustLedger(strAccNo, m_DTVchr, INextSno, dAmt)
'            End If
'             INextSno = INextSno + 1
'            .MoveNEXT
'        Loop
'        .Close
    End With
     
    ' Debit Entries
    
    Dim dStockBal As Double
    dStockBal = getBalance(strCostOfSalesAccNo, m_DTVchr)
     
    Dim dNetAmt As Double
    'dNetAmt = dTotalAmt + dStockBal
    dTotalAmt = dTotalAmtP + dTotalAmtDX + dTotalAmtFO
    dStockBal = dTotalAmt + dStockBal
    'Start New Transaction
    'con.BeginTrans
    
    'Credit Entry
    If bViewOnly = False Then
        con.Execute "INSERT INTO Vouchers VALUES(" & INextSno & ",'" & _
         Format(m_DTVchr, "dd-MMM-yyyy") & "','" & VNo & "','" & strCostOfSalesAccNo & _
        "','" & Desc & "'," & dTotalAmt & ",0," & dStockBal & ",'',0)"
          
        Call adjustLedger(strCostOfSalesAccNo, m_DTVchr, INextSno, dTotalAmt)
        con.Execute "INSERT INTO VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & VNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
    End If
    
    
    Dim ITM As ListItem
    Set ITM = LVSAE.ListItems.add(, , "Cost of Sales")
    ITM.ListSubItems.add(, , Format(dTotalAmt, "0.00")).ForeColor = vbBlue
    ITM.ListSubItems.add(, , Format(0, "0.00")).ForeColor = vbBlue
    ITM.ListSubItems.add(, , Format(dStockBal, "0.00")).ForeColor = vbBlue
    
    Set ITM = LVSAE.ListItems.add(, , "Purchases")
    ITM.ListSubItems.add(, , Format(0, "0.00")).ForeColor = vbRed
    ITM.ListSubItems.add(, , Format(dTotalAmtP, "0.00")).ForeColor = vbRed
    ITM.ListSubItems.add(, , Format(0, "0.00")).ForeColor = vbRed
    
    Set ITM = LVSAE.ListItems.add(, , "Direct Expenses")
    ITM.ListSubItems.add(, , Format(0, "0.00")).ForeColor = vbRed
    ITM.ListSubItems.add(, , Format(dTotalAmtDX, "0.00")).ForeColor = vbRed
    ITM.ListSubItems.add(, , Format(0, "0.00")).ForeColor = vbRed
    
     Set ITM = LVSAE.ListItems.add(, , "Factory Overheads")
    ITM.ListSubItems.add(, , Format(0, "0.00")).ForeColor = vbRed
    ITM.ListSubItems.add(, , Format(dTotalAmtFO, "0.00")).ForeColor = vbRed
    ITM.ListSubItems.add(, , Format(0, "0.00")).ForeColor = vbRed
   
    
    m_dCostofSalesBal = dTotalAmt
    AddClosingStockAdjustmentVoucher = VNo
    
    Exit Function
err:
    AddVoucher = False
    MsgBox "Error In AddVoucher :" & err.Number & vbNewLine & err.Description
    

End Function

Private Function AddClosingStockVoucher(Optional bViewOnly As Boolean = True) As String

    'Store Null To Return on an Unsuccessful Entry
    Dim VNo As String, m_DTVchr As Date
    m_DTVchr = DTFinish
    VNo = getNextVchrNo(m_DTVchr, "CL")
    AddVoucher = ""
    Dim i As Integer
    
    Dim dAmt As Double, dAmtRawMaterial As Double, dAmtWIP As Double
    
    dAmt = Val(txtClosingStock)
    dAmtRawMaterial = Val(txtClosingStockRawMaterial)
    dAmtWIP = Val(txtClosingStockWorkInProcess)
    
    Dim strDebitAccNo As String, strCreditAccNo As String
    
    Dim strDebitAccTitle As String, strCreditAccTitle As String
    Dim dDebitAmt As Double, dCreditAmt As Double
    Dim dDebitBal As Double, dCreditBal As Double
   
    'Account Nos to be Defined....
     
    Dim dStockBal As Double, dStockBalRawMaterial As Double, dStockBalWIP As Double
    dStockBal = getBalance(strStockAccNoFinished, m_DTVchr)
    dStockBalRawMaterial = getBalance(strStockAccNoRawMaterial, m_DTVchr)
    dStockBalWIP = getBalance(strStockAccNoWorkInProcess, m_DTVchr)
     
    Dim dNetAmt As Double, dNetAmtFinished As Double, dNetAmtRawMaterial As Double, dNetAmtWIP As Double
    Dim dDebit As Double, dCredit As Double
    
    'Finished//////////////////////////////////////////
    If dAmt > dStockBal Then    'Finished...
        dNetAmtFinished = dAmt - dStockBal
        dDebit = dNetAmtFinished
        dCredit = 0
    Else
        dNetAmtFinished = dAmt - dStockBal
        dDebit = 0
        dCredit = Abs(dNetAmtFinished)
    End If
    
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim Bal As Double, Desc As String
    INextSno = getNextSno(m_DTVchr)
    SnoToUse = INextSno
    Bal = getBalance(strStockAccNoFinished, m_DTVchr)
    Desc = "Closing Stock"
    Bal = Bal + dDebit - dCredit
    
    If bViewOnly = False Then
        con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
          Format(m_DTVchr, "dd-MMM-yyyy") & "','" & VNo & "','" & strStockAccNoFinished & _
         "','" & Desc & "'," & dDebit & "," & Abs(dCredit) & "," & Bal & ",'',0)"
           
        Call adjustLedger(strStockAccNoFinished, m_DTVchr, INextSno, dDebit - dCredit)
    End If
    'Finished//////////////////////////////////////////
    
    'Raw Material//////////////////////////////////////////
    If dAmtRawMaterial > dStockBalRawMaterial Then    'Raw Material...
        dNetAmtRawMaterial = dAmtRawMaterial - dStockBalRawMaterial
        dDebit = dNetAmtRawMaterial
        dCredit = 0
    Else
        dNetAmtRawMaterial = dAmtRawMaterial - dStockBalRawMaterial
        dDebit = 0
        dCredit = Abs(dNetAmtRawMaterial)
    End If
   
    INextSno = INextSno + 1
    Bal = getBalance(strStockAccNoRawMaterial, m_DTVchr)
    Bal = Bal + dDebit - dCredit
    
    If bViewOnly = False Then
        con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
          Format(m_DTVchr, "dd-MMM-yyyy") & "','" & VNo & "','" & strStockAccNoRawMaterial & _
         "','" & Desc & "'," & dDebit & "," & Abs(dCredit) & "," & Bal & ",'',0)"
           
        Call adjustLedger(strStockAccNoRawMaterial, m_DTVchr, INextSno, dDebit - dCredit)
    End If
    'Raw Material//////////////////////////////////////////
     
    'Work In Process//////////////////////////////////////////
    If dAmtWIP > dStockBalWIP Then    'Work In Process...
        dNetAmtWIP = dAmtWIP - dStockBalWIP
        dDebit = dNetAmtWIP
        dCredit = 0
    Else
        dNetAmtWIP = dAmtWIP - dStockBalWIP
        dDebit = 0
        dCredit = Abs(dNetAmtWIP)
    End If
   
    INextSno = INextSno + 1
    Bal = getBalance(strStockAccNoWorkInProcess, m_DTVchr)
    Bal = Bal + dDebit - dCredit
    
    If bViewOnly = False Then
        con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
          Format(m_DTVchr, "dd-MMM-yyyy") & "','" & VNo & "','" & strStockAccNoWorkInProcess & _
         "','" & Desc & "'," & dDebit & "," & Abs(dCredit) & "," & Bal & ",'',0)"
           
        Call adjustLedger(strStockAccNoWorkInProcess, m_DTVchr, INextSno, dDebit - dCredit)
    End If
    'Work In Process//////////////////////////////////////////
    
'    If dAmt > dStockBal Then
'        dNetAmt = dAmt - dStockBal
'        strDebitAccNo = strStockAccNo
'        strCreditAccNo = strCostOfSalesAccNo
'        strCreditAccTitle = "Cost of Sales"
'        strDebitAccTitle = "Stock"
'    Else
'        dNetAmt = dAmt - dStockBal
'        strDebitAccNo = strCostOfSalesAccNo
'        strCreditAccNo = strStockAccNo
'        strCreditAccTitle = "Stock"
'        strDebitAccTitle = "Cost of Sales"
'    End If
    
    
    
    
    
    
    'Start New Transaction
    'con.BeginTrans
    
    'Credit Entry
    
    
    INextSno = INextSno + 1
    ' Debit Entries
    
    Bal = getBalance(strCostOfSalesAccNo, m_DTVchr)
   
    dNetAmt = dNetAmtFinished + dNetAmtRawMaterial + dNetAmtWIP
    
    
    If dNetAmt > 0 Then
        dCredit = dNetAmt
        dDebit = 0
    Else
        dDebit = Abs(dNetAmt)
        dCredit = 0
    End If
    Bal = Bal + dDebit - dCredit
    
    If bViewOnly = False Then
        con.Execute "INSERT INTO  Vouchers VALUES(" & INextSno & ",'" & m_DTVchr & "','" & _
         VNo & "','" & strCostOfSalesAccNo & "','" & Desc & "'," & _
         dDebit & "," & dCredit & "," & Bal & ",'',0)"
        
        Call adjustLedger(strCostOfSalesAccNo, m_DTVchr, INextSno, dDebit - dCredit)
    
        con.Execute "INSERT INTO VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & VNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
    End If
    

'    Dim ITM As ListItem
'    Set ITM = LVSE.ListItems.add(, , strDebitAccTitle)
'    ITM.ListSubItems.add(, , Format(dDebitAmt, "0.00")).ForeColor = vbBlue
'    ITM.ListSubItems.add(, , Format(0, "0.00")).ForeColor = vbBlue
'    ITM.ListSubItems.add(, , Format(dDebitBal, "0.00")).ForeColor = vbBlue
'
'    Set ITM = LVSE.ListItems.add(, , strCreditAccTitle)
'    ITM.ListSubItems.add(, , Format(0, "0.00")).ForeColor = vbRed
'    ITM.ListSubItems.add(, , Format(dCreditAmt, "0.00")).ForeColor = vbRed
'    ITM.ListSubItems.add(, , Format(dCreditBal, "0.00")).ForeColor = vbRed
    
    
    AddClosingStockVoucher = VNo
    
    Exit Function
err:
    AddVoucher = False
    MsgBox "Error In AddVoucher :" & err.Number & vbNewLine & err.Description
    

End Function

Private Function AddYearClosingVoucher(Optional bViewOnly As Boolean = True) As String
    
    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim AccList() As String
    Dim AccBalances() As Double
    Dim i As Integer
    Dim strVchrNo As String
    Dim strTableName As String
    Dim ITM As ListItem
    If bViewOnly Then
        strTableName = "VouchersYC"
        con.Execute "DELETE FROM VouchersYC"
    Else
        strTableName = "Vouchers"
    End If
    Dim strCondition As String
    strCondition = " WHERE Cast(Left(AccNo,2) AS INT)>=31 AND AccNo<>'" & strProfitANDLossAccNo & "'"
    With rs
'        If dtFinish = #6/30/2011# Then
'            .Open "SELECT Accounts.AccNo,ISNULL(T1.Balance,0)-ISNULL(OpenBal,0) AS Balance FROM Accounts LEFT OUTER JOIN (SELECT AccNo,SUM(Credit)-SUM(Debit) As Balance FROM Vouchers Accounts " & strCondition & " AND VDate BETWEEN '" & DTStart & "' AND '" & dtFinish & "' GROUP BY AccNo HAVING SUM(Credit)-SUM(Debit)<>0) T1 ON Accounts.AccNo=T1.AccNo " & strCondition & " ORDER BY Accounts.AccNo", con, adOpenForwardOnly, adLockReadOnly
'        Else
            .Open "SELECT AccNo,SUM(Credit)-SUM(Debit) As Balance FROM Vouchers " & strCondition & " AND VDate BETWEEN '" & DTStart & "' AND '" & DTFinish & "' GROUP BY AccNo HAVING SUM(Credit)-SUM(Debit)<>0  ORDER BY AccNo", con, adOpenForwardOnly, adLockReadOnly
'        End If
        
        ReDim AccList(.RecordCount)
        ReDim AccBalances(.RecordCount)
        i = 0
        Do Until .EOF
            AccList(i) = !AccNo & ""
            
            AccBalances(i) = !Balance
            i = i + 1
            .MoveNext
        Loop

        .Close
        strVchrNo = getNextVchrNo(DTFinish, "CL")
    End With
    
    If bViewOnly Then
        Dim iNo As Integer
        iNo = UBound(AccList)
        ReDim Preserve AccList(iNo + 1)
        ReDim Preserve AccBalances(iNo + 1)
        AccList(iNo) = strCostOfSalesAccNo
        AccBalances(iNo) = -dCostofSalesBal
    End If
    
    Dim INextSno As Double
    INextSno = getNextSno(DTFinish)
    
    If bViewOnly = False Then
        Call StartTrans(con)
    End If
    
    Dim dCredit As Double, dDebit As Double
    Dim dTotalCredit As Double, dTotalDebit As Double
    
    For i = 0 To UBound(AccList) - 1
        If AccBalances(i) < 0 Then  'Debit Is Larget
            dCredit = Abs(AccBalances(i))
            dDebit = 0
        Else
            dDebit = Abs(AccBalances(i))
            dCredit = 0
        End If
        
        If dDebit <> 0 Or dCredit <> 0 Then
            con.Execute "INSERT INTO  " & strTableName & " Values(" & INextSno & ",'" & Format(DTFinish, "dd-MMM-yyyy") & "','" & _
             strVchrNo & "','" & AccList(i) & "','Year Closing Voucher'," & _
             dDebit & "," & dCredit & ",0,'',0)"
        End If
       If bViewOnly = False Then
            If dCredit > 0 Then
                Call adjustLedger(AccList(i), DTFinish, INextSno, -dCredit)
            Else
                Call adjustLedger(AccList(i), DTFinish, INextSno, dDebit)
            End If
        End If
        
        INextSno = INextSno + 1
        dTotalCredit = dTotalCredit + dCredit
        dTotalDebit = dTotalDebit + dDebit
        Debug.Print i
    Next
    
    Dim strAccNo As String, dAmt As Double
    Dim dCreditAmt As Double, dDebitAmt As Double
    
    If dTotalCredit > dTotalDebit Then
        dDebitAmt = dTotalCredit - dTotalDebit
        dCreditAmt = 0
    Else
        dCreditAmt = dTotalDebit - dTotalCredit
        dDebitAmt = 0
    End If
     
    Dim dPLBal As Double
    dPLBal = getBalance(strProfitANDLossAccNo, DTFinish)
    
    con.Execute "Insert Into " & strTableName & " Values(" & INextSno & ",'" & Format(DTFinish, "dd-MMM-yyyy") & "','" & _
     strVchrNo & "','" & strProfitANDLossAccNo & "','Year Closing Voucher'," & _
     dDebitAmt & "," & dCreditAmt & "," & dPLBal + (dDebitAmt - dCreditAmt) & ",'',0)"
         
    If bViewOnly = False Then
        Call adjustLedger(strAccNo, DTFinish, INextSno, dDebitAmt - dCreditAmt)
     
        con.Execute "Insert Into VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & strVchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
        con.CommitTrans
    End If
    

        Dim lColor As Long
        Dim myArr() As String
        Dim iIndex As Integer
        iIndex = 0
        With rs
            .Open "SELECT Accounts.AccTitle,Tab1.* FROM Accounts INNER JOIN (SELECT Accounts.SubAccOf,SUM(Credit) AS Credit,SUM(Debit) AS Debit,SUM(VouchersYC.Balance) AS Balance FROM VOUCHERSYC " & _
                "INNER JOIN Accounts ON VouchersYC.AccNo=Accounts.AccNo " & _
                "WHERE Accounts.SubAccOf<>'' GROUP BY Accounts.SubAccOf) Tab1 ON " & _
                "Tab1.SubAccOf = Accounts.AccNo ", con, adOpenForwardOnly, adLockReadOnly
            Do Until .EOF
                Set ITM = LVYC.ListItems.add(, , !AccTitle & "")
                
                If Val(!Credit) > 0 Then lColor = vbRed Else lColor = vbBlue
                ITM.ListSubItems.add(, , Format(Round(Val(!Debit & ""), 2), "0.00")).ForeColor = lColor
                ITM.ListSubItems.add(, , Format(Round(Val(!Credit & ""), 2), "0.00")).ForeColor = lColor
                ITM.ListSubItems.add(, , Format(Round(Val(!Balance & ""), 2), "0.00")).ForeColor = lColor
                .MoveNext
            Loop
            .Close
            
            .Open "SELECT Accounts.Type,Heads.Head,SUM(Credit) AS Credit,SUM(Debit) AS Debit,SUM(VouchersYC.Balance) AS Balance FROM VOUCHERSYC " & _
                "INNER JOIN Accounts ON VouchersYC.AccNo=Accounts.AccNo INNER JOIN Heads ON Accounts.TYpe=Heads.Code " & _
                "WHERE ISNULL(Accounts.SubAccOf,'')='' and Accounts.AccNo<>'46-16001' GROUP BY Accounts.Type,Heads.Head", con, adOpenForwardOnly, adLockReadOnly
            ReDim Preserve myArr(iIndex + .RecordCount + 1, 3)
            Do Until .EOF
                Set ITM = LVYC.ListItems.add(, , !Head & "")
                If Val(!Credit) > 0 Then lColor = vbRed Else lColor = vbBlue
                ITM.ListSubItems.add(, , Format(Round(Val(!Debit & ""), 2), "0.00")).ForeColor = lColor
                ITM.ListSubItems.add(, , Format(Round(Val(!Credit & ""), 2), "0.00")).ForeColor = lColor
                ITM.ListSubItems.add(, , Format(Round(Val(!Balance & ""), 2), "0.00")).ForeColor = lColor
                .MoveNext
            Loop
            .Close
            
'            .Open "SELECT Credit,Debit,Balance FROM VouchersYC WHERE AccNo='" & strProfitANDLossAccNo & "'", con, adOpenForwardOnly, adLockReadOnly
'            Set ITM = LVYC.ListItems.add(, , "Profit & Loss")
'            If Val(!Credit) > 0 Then lColor = vbRed Else lColor = vbBlue
'            ITM.ListSubItems.add(, , Format(Round(Val(!Debit & ""), 2), "0.00")).ForeColor = lColor
'            ITM.ListSubItems.add(, , Format(Round(Val(!Credit & ""), 2), "0.00")).ForeColor = lColor
'            ITM.ListSubItems.add(, , Format(Round(Abs(Val(!Balance & "")), 2), "0.00")).ForeColor = lColor
'            .Close
        End With

    Set rs = Nothing
    
    AddYearClosingVoucher = strVchrNo
    Exit Function
    
err:
    MsgBox err.Description

End Function
 
