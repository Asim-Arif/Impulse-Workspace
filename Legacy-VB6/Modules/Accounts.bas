Attribute VB_Name = "Accounts"
Public Const ACCRECTITLE = "Accounts Receivable"
Public Const ACCRECNO = "15-003"
Public Const FSALESACC = "31-001-06001"
Public Const FSAMPLESALESACC = "31-004-19001"
'Foreign Customers AccNo
Public Const FCUSTACCNO = "15-003"

'Local Customers Accno
'Public Const LCUSTACCNO = "15-003-002"

Public Const VENDACCNO = "24-004-001"
Public Const FBCONTACCNO = "24-005-003-001"


Public Const ACCPAYTITLE = "Accounts Payable"
Public Const ACCPAYNO = "24-004"

Public Const ACCBANKSTITLE = "Bank Balances"
Public Const ACCBANKSNo = "15-002"

Public CashInHand As String

Public Const SALEACC = "31-001"

Public Const PAPERACC = "41-001"
Public Const ACCPCHNO = "41-001"

Public Const PCHACC = ACCPCHNO

Public Const ACCEMPNO = "42-002"
Public Const ACCDIRNO = "42-001"

'Discount Account
Public Const DISCOUNTACC = "41-014"

'Sale Return & Purchase Return Account #
Public Const SRTACC = "31-003"
Public Const PRTACC = "41-002"

'Stock Adjustement Account
Public Const ADJACC = "41-015"

Public Const STOCKASSETACC = "15-004"

Public Const CONSUMABLESTOCKASSETACC = "15-006"

Public Const FOOTBALLSWAGESACC = "41-003-001-001"
Public Const GARMENTSWAGESACC = "41-003-002-001"
Public Const GLOVESWAGESACC = "41-003-003-001"


Public Const FOOTBALLSACCRUEDWAGESACC = "24-005-030-001"
Public Const GARMENTSACCRUEDWAGESACC = "24-005-030-002"
Public Const GLOVESACCRUEDWAGESACC = "24-005-030-003"

'Public CashInHand As String
'Public AccRecTITLE As String

'Public AccPayTITLE As String
'Public AccBanksTITLE As String

'Public AccEmpTitle As String
'Public AccRecNo As String

'Public AccPayNo As String
'Public AccBanksNo As String

'Public AccEmpNo As String
'Public AccDirNo As String

'Public AccPchNo As String
'Public SaleAcc As String
'Public PchAcc As String

Public YearStartDate As Date
Public YearLastDate As Date

'Discount Account
'Public DiscountAcc As String

'Sales & Purchase Return Accounts
'Public SRTAcc As String
'Public PRTAcc As String

'Stock Adjustement Account
'Public AdjAcc As String
'Public StockAssetAcc As String


'Public PaperAcc As String
Public OutOfFinancialYear As Boolean
Public Const StitchingAccNo = "24-005-003-001-01005"
Public RecentAccNo As String

Public Sub PrintPaymentSheet()

    Dim rpt As CRAXDDRT.Report, f As New frmPrevRpt
    Set rpt = rptApp.OpenReport(rptAccountsPath & "\rptPaymentList.rpt")
    f.ShowReport "", rpt
    
End Sub

Public Function getBalance(AccountNo As String, OnDate As Date, Optional OpeningBalance As Boolean) As Double

    On Error GoTo err

    Dim rsBal As New ADODB.Recordset, SubQry As String
    Dim MaxDate As Date
    
    'Get The date Of Last Entry
    rsBal.Open "Select Max(VDate) from Vouchers Where Accno='" & AccountNo & "'", con, adOpenForwardOnly, adLockOptimistic
    
    If Not IsNull(rsBal.Fields(0)) Then
        MaxDate = rsBal.Fields(0)
    Else
        'If No Entry found  Then Current Date is Maximum Date
        MaxDate = Date
    End If
    'Close Recordset
    rsBal.Close

    'If This Entry is In Previous Dates then
    'Get Balance From Previous entries

    Dim TransAmt As Double
    TransAmt = 0
    If OnDate < MaxDate Then
        SubQry = "Select max(SNo) From Vouchers where VDate<='" & Format(OnDate, "dd-MMM-yyyy") & "' and accno='" & AccountNo & "'"
        rsBal.Open "Select Balance,Credit-Debit as TransAmt From Vouchers where Sno=(" & SubQry & ")", con, adOpenForwardOnly, adLockOptimistic
        If rsBal.RecordCount > 0 Then TransAmt = rsBal![TransAmt]
    Else
        'Else Get Current Balance
        rsBal.Open "Select Balance from Accounts Where Accno='" & AccountNo & "'", con, adOpenForwardOnly, adLockReadOnly
    End If


    If rsBal.RecordCount = 0 Then
        rsBal.Close
        rsBal.Open "Select OpenBal from Accounts Where AccNo='" & AccountNo & "'", con, adOpenForwardOnly, adLockReadOnly
    End If

    If OpeningBalance Then getBalance = Val(rsBal.Fields(0) & "") + TransAmt Else getBalance = Val(rsBal.Fields(0) & "")

    Set rsBal = Nothing

    Exit Function
err:
    MsgBox "Followinfg error Occured in getBalance" & vbNewLine & err.Description & vbNewLine & " Acc No.:" & AccountNo
    
End Function

Function getAccountNo(AccountHead As String, Optional Condition As String) As String

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This Function Finds an account no against the given Account Head.'                                       '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim rsAccount As New ADODB.Recordset
  
    If Len(Trim(Condition)) > 0 Then Condition = " AND " & Condition
  
    'First Find In Second Level
    With rsAccount
        .CursorLocation = adUseClient
        .Open "SELECT AccNo From Accounts Where Upper(AccTitle)='" & UCase(AccountHead) & "' " & Condition, con, adOpenForwardOnly, adLockReadOnly
 
        If .RecordCount > 0 Then
            getAccountNo = .Fields(0) & ""
        Else
            getAccountNo = ""
        End If
 
        .Close
    End With
    Set rsAccount = Nothing
 
End Function

Public Function NewAccRcv(AccountName As String, OpenDate As Date, Openbalance As Double, Active As Boolean, Optional SubAccOf As String, Optional dontSave As Boolean = False, Optional IsParent As Boolean = False) As String

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '            This Function Adds new Accounts Receivable.           '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim rsAcc As New ADODB.Recordset
    Dim AccRec As String, Acno As String
    'Generate Customer Code
    Dim NewVal As String, Prefix As String

    NewAccRcv = ""
    If SubAccOf = "" Then AccRec = ACCRECNO Else AccRec = SubAccOf
 
    'Generate The Account # Sorted By Name
    'Prefix = Asc(UCase(Left(Trim(AccountName), 1)))
    'If Prefix >= 65 And Prefix <= 90 Then       'Starting From Alphabet
        '  Prefix = Format(Prefix - 64, "00")
    'Else
        '  Prefix = "00"
    'End If
 
    '   NewVal = GetMax("right(AccNo,3)", "Accounts", "Where SubAccOf='" & AccRec & "' and substring(AccNo,len(accno)-4,2)='" & Prefix & "'")
    '   AcNo = AccRec & "-" & Prefix & Format(Val(NewVal) + 1, "000")
   
    Acno = GetNextAccno(AccountName, Left(AccRec, 2), AccRec, IsParent)
   
    'If Just To Get Next Account No Then Exit Function
   
    If dontSave Then
        NewAccRcv = Acno
        Exit Function
    End If
   
    If AccRec = "" Then Exit Function

    con.Execute "Insert Into Accounts(AccNo,SubAccOf,Type,AccTitle,OpenDate,OpenBal,Balance,Parent)" & _
            " Values('" & Acno & "','" & AccRec & "','" & Left(AccRec, 2) & "','" & AccountName & "','" & Format(OpenDate, "dd-MMM-yyyy") & "'," & _
            Openbalance & "," & Openbalance & "," & Abs(IsParent) & ")"
         
    ''''''''''''To Do''''''''''''''''''
    'Call UpdateAdvEntry(Acno, OpenBalance, OpenDate)
 
    NewAccRcv = Acno
    If rsAcc.State = 1 Then rsAcc.Close
    Set rsAcc = Nothing
    
    Exit Function
err:
    MsgBox "Cannot Add Account Due to ERROR:" & err.Number & vbNewLine & err.Description, vbInformation
    NewAccRcv = ""
    
End Function

Public Function NewAccEmp(AccountName As String, OpenDate As Date, Openbalance As Double, Active As Boolean) As String


      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '                 This Function Adds new Employee Account          '
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim rsAcc As New ADODB.Recordset
    Dim AccEmp As String, Acno As String
    'Generate Employee Account #
    Dim NewVal As String
    NewAccEmp = ""
    AccEmp = ACCEMPNO
    NewVal = GetMax("right(AccNo,3)", "Accounts", "Where SubAccOf='" & AccEmp & "'")
    Acno = AccEmp & "-" & Format(Val(NewVal) + 1, "000")
 
    Dim SAdvAcc As String, LAdvAcc As String, AcrAcc As String
    
    SAdvAcc = getAccountNo("Short Term Advances")
    LAdvAcc = getAccountNo("Long Term Advances")
    AcrAcc = getAccountNo("Accrued Salaries")
 
    'Short Term Advance Account
    NewVal = GetMax("right(AccNo,3)", "Accounts", "Where SubAccOf='" & SAdvAcc & "'")
    SAdvAcc = SAdvAcc & "-" & Format(Val(NewVal) + 1, "000")
    'Long Term Advance Account
    NewVal = GetMax("right(AccNo,3)", "Accounts", "Where SubAccOf='" & LAdvAcc & "'")
    LAdvAcc = LAdvAcc & "-" & Format(Val(NewVal) + 1, "000")
    'Accrued salaries
    NewVal = GetMax("right(AccNo,3)", "Accounts", "Where SubAccOf='" & AcrAcc & "'")
    AcrAcc = AcrAcc & "-" & Format(Val(NewVal) + 1, "000")
 
    If AccEmp = "" Then Exit Function
    
    With rsAcc
    
        .Open "Select * from Accounts where AccNo='fake'", con, adOpenDynamic, adLockOptimistic
        
        .AddNew
 
        ![Type] = Left(AccEmp, 2)
        ![AccNo] = Acno
        ![SubAccOf] = AccEmp
        ![AccTitle] = AccountName
        ![OpenDate] = OpenDate
        ![OpenBal] = Openbalance
        ![Balance] = Openbalance
        
        .Update
 
        'If OpenBal <> 0 Then Call UpdateAdvEntry(![AccNo], ![OpenBal])
    End With
 
    NewAccEmp = Acno
 
 
    Exit Function
err:
    MsgBox "Cannot Add Account Due to ERROR:" & err.Number & vbNewLine & err.Description, vbInformation
    NewAccEmp = ""
    
End Function

Public Function NewAccPay(AccountName As String, OpenDate As Date, Openbalance As Double, Active As Boolean, Optional SubAccOf As String, Optional dontSave As Boolean = False) As String

      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '                 This Function Adds new Vender Account
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
    Dim accPay As String, Acno As String

    'Generate Employee Account #

    Dim NewVal As String, Prefix As String
    NewAccPay = ""
 
    If SubAccOf = "" Then accPay = ACCPAYNO Else accPay = SubAccOf
 
    If accPay = "" Then Exit Function
 
    'Generate The Account # Sorted By Name
    Prefix = Asc(UCase(Left(Trim(AccountName), 1)))
    If Prefix >= 65 And Prefix <= 90 Then       'Starting From Alphabet
        Prefix = Format(Prefix - 64, "00")
    Else
        Prefix = "00"
    End If

    NewVal = GetMax("right(AccNo,3)", "Accounts", "Where SubAccOf='" & accPay & "' and substring(AccNo,len(accno)-4,2)='" & Prefix & "'")
    Acno = accPay & "-" & Prefix & Format(Val(NewVal) + 1, "000")
   
    If dontSave Then
        NewAccPay = Acno
        Exit Function
    End If

    con.Execute "Insert Into Accounts(AccNo,SubAccOf,Type,AccTitle,OpenDate,OpenBal,Balance)" & _
            " Values('" & Acno & "','" & accPay & "','" & Left(accPay, 2) & "','" & AccountName & "','" & Format(OpenDate, "dd-MMM-yyyy") & "'," & _
             Openbalance & "," & Openbalance & ")"
            
    If Openbalance <> 0 Then Call UpdateAdvEntry(Acno, Openbalance, OpenDate)
 
    NewAccPay = Acno

    Exit Function
err:
    MsgBox "Cannot Add Account Due to ERROR:" & err.Number & vbNewLine & err.Description, vbInformation
    NewAccPay = ""
    
End Function

Public Function GetNextAccno(AccountName, AccType, Optional ParentAccount, Optional IsParent As Boolean = False)

    'Generate Customer Code
    Dim NewVal As String, Prefix As String
    Dim ParentCondition As String
 
    ParentCondition = IIf(Len(ParentAccount) = 2, " (len(isnull(SubAccof,''))=0) and type ='" & ParentAccount & "'", " subaccof='" & ParentAccount & "'")
 
    If IsParent Then
        NewVal = GetMax("right(AccNo,3)", "Accounts", "Where " & ParentCondition & " and Parent=1")
        GetNextAccno = ParentAccount & "-" & Format(Val(NewVal) + 1, "000")
    Else
 
        'Generate The Account # Sorted By Name
        Prefix = Asc(UCase(Left(Trim(AccountName), 1)))
        If Prefix >= 65 And Prefix <= 90 Then       'Starting From Alphabet
           Prefix = Format(Prefix - 64, "00")
        Else
           Prefix = "00"
        End If
 
        NewVal = GetMax("right(AccNo,3)", "Accounts", "Where " & ParentCondition & " and substring(AccNo,len(accno)-4,2)='" & Prefix & "'")
        GetNextAccno = ParentAccount & "-" & Prefix & Format(Val(NewVal) + 1, "000")
   
    End If

End Function

Public Function AddNewAccount(AccountName As String, OpenDate As Date, Optional ParentAccount As String, Optional Openbalance As Double, Optional IsParent As Boolean = False) As String

    Dim AccNo As String
    
    AccNo = GetNextAccno(AccountName, Left(ParentAccount, 2), ParentAccount, IsParent)
    
    con.Execute "Insert Into Accounts(AccNo,SubAccOf,Type,AccTitle,OpenDate,OpenBal,Balance)" & _
        " Values('" & AccNo & "','" & ParentAccount & "','" & Left(ParentAccount, 2) & "','" & AccountName & "','" & Format(OpenDate, "dd-MMM-yyyy") & "'," & _
        Openbalance & "," & Openbalance & ")", a

    If Openbalance <> 0 Then Call UpdateAdvEntry(AccNo, Openbalance, OpenDate)
 
    If a = 1 Then AddNewAccount = AccNo
    
End Function

Public Sub PrintIncomeStatement(Optional bShowDetail As Boolean = True)

    Dim myDTFrom As Date, myDTTo As Date
    Dim f As New frmGetDateDouble
    Load f
    If f.getDate(myDTFrom, myDTTo) = False Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\IncomeStatement.rpt")
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue myDTFrom
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue myDTTo
    
    rpt.EnableParameterPrompting = False
    
    rpt.FormulaFields.GetItemByName("DateRange").Text = "'" & Format(myDTFrom, "dd-MMM-yyyy") & " to " & Format(myDTTo, "dd-MMM-yyyy") & "'"
'    rpt.FormulaFields.GetItemByName("ShowDetail").Text = bShowDetail
    
    Dim frm As New frmPrevRpt
    Load frm
    frm.ShowReport "", rpt, "", False
    
End Sub

Public Sub PrintBalanceSheet(Optional bShowDetail As Boolean = True)

    Dim myDTFrom As Date, myDTTo As Date
    Dim f As New frmGetDateSingle
    Load f
    If f.getDate(myDTTo) = False Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\BalanceSheet.rpt")
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue myDTTo
    rpt.EnableParameterPrompting = False
    
    rpt.FormulaFields.GetItemByName("DateRange").Text = "'" & Format(myDTTo, "dd-MMM-yyyy") & "'"
    rpt.FormulaFields.GetItemByName("ShowDetail").Text = bShowDetail
    
    Dim frm As New frmPrevRpt
    Load frm
    frm.ShowReport "", rpt, "", False
    
End Sub


Public Sub PrintCashBookReport(DtFrom As Date, DtTo As Date)

    Dim rpt As CRAXDDRT.Report, rptApp_Temp As New CRAXDDRT.Application
    Set rpt = rptApp.OpenReport(RptPath & "\CashBookReport.rpt")
    
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DtFrom
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DtTo
     
    Dim rptSubCredit As CRAXDDRT.Report, rptSubDebit As CRAXDDRT.Report
        
    Set rptSubCredit = rpt.OpenSubreport("CashCreditEntries")
    rptSubCredit.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DtFrom
    rptSubCredit.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DtTo
    rptSubCredit.RecordSelectionFormula = "{VLedger.Credit}>0"
     
    Set rptSubDebit = rpt.OpenSubreport("CashDebitEntries")
    rptSubDebit.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DtFrom
    rptSubDebit.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DtTo
    rptSubDebit.RecordSelectionFormula = "{VLedger.Debit}>0"
    
    rpt.EnableParameterPrompting = False
    rpt.FormulaFields.GetItemByName("FromTo").Text = "'" & Format(DtFrom, "dd-MMM-yyyy") & " to " & Format(DtTo, "dd-MMM-yyyy") & "'"
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, , False
    
End Sub


Public Sub PrintMonthlyBudgetReport()
    Dim myDT As Date
    Dim frm As New frmGetDateSingle
    
    Load frm
    frm.DTPicker1.CustomFormat = "MMM, yyyy"
    If frm.getDate(myDT) = False Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\FinancialBudgeting_GroupWise.rpt")
    rpt.ParameterFields.GetItemByName("@DT").SetCurrentValue myDT
    
    rpt.EnableParameterPrompting = False
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, , False
    
End Sub
