Attribute VB_Name = "Vouchers"
Public AutoDepriciation As Boolean

Public Function VoucherRequiresAuthorization(strVchrType As String) As Boolean
    
    If strCompany <> "Sunlike" Then
        VoucherRequiresAuthorization = False
        Exit Function
    End If
    Dim iCount As Integer, strTableName As String
    iCount = GetSingleLongValue("COUNT(*)", "VoucherTypesToAuthorize", " WHERE WaitForAuthorization=1 AND VoucherType='" & strVchrType & "'")
    If iCount = 0 Then
        VoucherRequiresAuthorization = False
    Else
        VoucherRequiresAuthorization = True
    End If
    
End Function

Function getNextVchrNo(OnDate As Date, VoucherType As String) As String
 
    Dim rsVchr As New ADODB.Recordset
      
    Dim MonthlyVF As Boolean
    If strCompany = "Banzai" Then
        MonthlyVF = True
    Else
        MonthlyVF = False
    End If
    
    Dim lNewVchrNoFromVouchers As Long, lNewVchrNoFromUnAuth As Long, lNewVchrNo As Long
    lNewVchrNoFromVouchers = 0
    lNewVchrNoFromUnAuth = 0
    
    
    Dim strCondition As String
    
    If MonthlyVF Then
        If Len(VoucherType) = 2 Then
            strCondition = " WHERE LEFT(VchrNo,8)='" & VoucherType & "-" & Format(OnDate, "MMYY") & "-'"
        Else
            strCondition = " WHERE LEFT(VchrNo,9)='" & VoucherType & "-" & Format(OnDate, "MMYY") & "-'"
        End If
    Else
        If Len(VoucherType) = 2 Then
            strCondition = " WHERE LEFT(VchrNo,10)='" & VoucherType & "-" & Format(OnDate, "YYMMdd") & "-'"
        Else
            strCondition = " WHERE LEFT(VchrNo,11)='" & VoucherType & "-" & Format(OnDate, "YYMMdd") & "-'"
        End If
    End If
    
    lNewVchrNoFromVouchers = GetSingleLongValue("MAX(CAST(RIGHT(VchrNo,CHARINDEX('-',REVERSE(VchrNo))-1) AS INT))", "Vouchers", strCondition)
    lNewVchrNoFromUnAuth = GetSingleLongValue("MAX(CAST(RIGHT(VchrNo,CHARINDEX('-',REVERSE(VchrNo))-1) AS INT))", "VouchersForAuthorization", strCondition)
     
    If lNewVchrNoFromVouchers > lNewVchrNoFromUnAuth Then
        lNewVchrNo = lNewVchrNoFromVouchers + 1
    Else
        lNewVchrNo = lNewVchrNoFromUnAuth + 1
    End If
    
    If MonthlyVF Then
        getNextVchrNo = VoucherType & "-" & Format(OnDate, "MMYY") & "-" & Format(lNewVchrNo, "000")
    Else
        getNextVchrNo = VoucherType & "-" & Format(OnDate, "YYMMdd") & "-" & Format(lNewVchrNo, "000")
    End If
    

 End Function
 
 

Public Function getNextSno(ForDate As Date, Optional TableName) As Double

Dim rsSno As New ADODB.Recordset
If IsMissing(TableName) Then
    rsSno.Open "Select Max(Sno) from Vouchers Where VDate='" & Format(ForDate, "dd-MMM-yyyy") & "'", con, adOpenForwardOnly, adLockReadOnly
ElseIf UCase(TableName) = "VINOUTDETAIL" Then
    rsSno.Open "Select Max(Sno) from " & TableName & " Where Dated='" & Format(ForDate, "dd-MMM-yyyy") & "'", con, adOpenForwardOnly, adLockReadOnly
Else
    rsSno.Open "Select Max(Sno) from " & TableName & " Where CDate='" & Format(ForDate, "dd-MMM-yyyy") & "'", con, adOpenForwardOnly, adLockReadOnly
End If
 
If Val(rsSno.Fields(0) & "") = 0 Then
    getNextSno = Val(Format(ForDate, "yyMMdd") & "0001")
Else
    getNextSno = Val(rsSno.Fields(0) & "") + 1
End If

Set rsSno = Nothing

End Function




Public Sub adjustLedger(AccountNo As String, StartDate As Date, SNo As Double, AmountToAdd As Double, Optional DonotUpdateCurrentBalance As Boolean)

 
    If Not DonotUpdateCurrentBalance Then
        'Update Current Balance
        con.Execute "Update Accounts Set Balance=Balance+" & AmountToAdd & " Where AccNo='" & AccountNo & "'"
    End If
 
    Dim SQL As String
    'Update Ledger Balances
    SQL = "Update Vouchers set Balance=Balance+(" & AmountToAdd & ") Where AccNo='" & AccountNo & _
       "' and (VDate>='" & Format(StartDate, "dd-MMM-yyyy") & "' and Sno>" & SNo & ")"
 
    con.Execute SQL
 
End Sub

Public Function DeleteVoucher(VchrNo As String, Optional DontStartNewTransaction As Boolean = False, Optional strReason As String) As Boolean

    ''''''''    Deletes A Voucher And Returns All its Effects Back      ''''''''''''''''

    Dim rs As New ADODB.Recordset
    With rs
        .Open "select * From Vouchers Where VchrNo='" & VchrNo & "'", con, adOpenForwardOnly, adLockReadOnly
   
        If Not DontStartNewTransaction Then con.BeginTrans
      
        For i = 0 To .RecordCount - 1
            Call adjustLedger(![AccNo], ![Vdate], ![SNo], ![Credit] - ![Debit])
            If ![CSNo] > 0 Then con.Execute "Update Cheque Set [Description]='Canceled Due To Voucher Deletetion', chequeType='Canceled',Amount=0,Bounced=0,Posted=0  Where Sno=" & ![CSNo]
            .MoveNEXT
        Next
        con.Execute "INSERT INTO VouchersDeleted Select * From Vouchers Where VchrNo='" & VchrNo & "'"
        con.Execute "INSERT INTO VouchersDeletedInfo Select VchrNo,'" & CurrentUserName & "','" & _
         strComputerName & "','" & Date & "','" & strReason & "',UserName,MachineName,DT From " & _
         " VoucherInfo Where VchrNo='" & VchrNo & "'"
        
        con.Execute "DELETE FROM Vouchers Where VchrNo='" & VchrNo & "'"
        con.Execute "DELETE FROM VoucherInfo Where VchrNo='" & VchrNo & "'"
        .Close
    End With
    
    '/\\/\/\\/\\/\\/\/\/\/\/\\/\/\/\/\/\/\/\/\/\/\///////////////////////////
    'Delete From Advances..If It's An Advance Payment Entry To Employee.
    con.Execute "Delete From Advances WHERE AccVoucherNo='" & VchrNo & "'"
    con.Execute "DELETE FROM AmtCleared WHERE VchrNo='" & VchrNo & "'"
    Dim lMPB_EntryID As Long
    lMPB_EntryID = GetSingleLongValue("EntryID", "MakerPostedBills", " WHERE VchrNo='" & VchrNo & "'")
    con.Execute "DELETE FROM MakerPostedBills WHERE VchrNo='" & VchrNo & "'"
    con.Execute "DELETE FROM MakerPostedBills_STDeductions WHERE MPB_RefID=" & lMPB_EntryID
    con.Execute "DELETE FROM MakerAmtCleared WHERE VchrNo='" & VchrNo & "'"
    con.Execute "UPDATE VendIssued SET VchrNo=NULL WHERE VchrNo='" & VchrNo & "'"
    
    con.Execute "DELETE FROM MakerAdvancesShort WHERE AccVoucherNo='" & VchrNo & "'"
    con.Execute "DELETE FROM MakerAdvances WHERE AccVoucherNo='" & VchrNo & "'"
     
    con.Execute "DELETE FROM VenderPostedBills WHERE VchrNo='" & VchrNo & "'"
    con.Execute "DELETE FROM CustomInvoiceAuthorization WHERE VchrNo='" & VchrNo & "'"
   
    With rs
        .Open "SELECT Advances.EntryID FROM Advances INNER JOIN ShortTermAdvanceClearanceDetail ON ShortTermAdvanceClearanceDetail.Advances_RefID=Advances.EntryID INNER JOIN ShortTermAdvanceClearance ON ShortTermAdvanceClearance.EntryID=ShortTermAdvanceClearanceDetail.SAC_RefID WHERE ShortTermAdvanceClearance.VchrNo='" & VchrNo & "'"
        Do Until .EOF
            con.Execute "UPDATE Advances SET Cleared=0 WHERE EntryID=" & Val(!EntryID & "")
            .MoveNEXT
        Loop
        .Close
    End With
    con.Execute "DELETE FROM ShortTermAdvanceClearance WHERE VchrNo='" & VchrNo & "'"
    con.Execute "UPDATE FCustAdvancePayments SET VchrNo=NULL WHERE VchrNo='" & VchrNo & "'" 'Advance Payment
    con.Execute "DELETE FROM FCustAdvancePaymentsVouchers WHERE VchrNo='" & VchrNo & "'"
    con.Execute "DELETE FROM PRC WHERE VchrNo='" & VchrNo & "'"
    '/\\/\/\\/\\/\\/\/\/\/\/\\/\/\/\/\/\/\/\/\/\/\///////////////////////////
    
    If Not DontStartNewTransaction Then con.CommitTrans

    Set rs = Nothing

    DeleteVoucher = True

    Exit Function

err:
    MsgBox "Voucher Cannot Be Deleted Due To Error:" & err.Number & vbNewLine & err.Description
    DeleteVoucher = False

End Function

Public Function DeleteVoucherUnAuthorized(VchrNo As String, Optional DontStartNewTransaction As Boolean = False, Optional strReason As String) As Boolean

    ''''''''    Deletes A Voucher And Returns All its Effects Back      ''''''''''''''''

'    Dim rs As New ADODB.Recordset
'    With rs
'        .Open "SELECT * From Vouchers Where VchrNo='" & VchrNo & "'", con, adOpenForwardOnly, adLockReadOnly
'
        If Not DontStartNewTransaction Then con.BeginTrans
'
'        For i = 0 To .RecordCount - 1
'            Call adjustLedger(![AccNo], ![VDate], ![Sno], ![Credit] - ![Debit])
'            If ![CSNo] > 0 Then con.Execute "UPDATE Cheque Set [Description]='Canceled Due To Voucher Deletetion', chequeType='Canceled',Amount=0,Bounced=0,Posted=0  WHERE Sno=" & ![CSNo]
'            .MoveNext
'        Next
'        'con.Execute "INSERT INTO VouchersDeleted Select * From Vouchers Where VchrNo='" & VchrNo & "'"
'        'con.Execute "Insert Into VouchersDeletedInfo Select VchrNo,'" & CurrentUserName & "','" & _
'         strComputerName & "','" & Date & "','" & strReason & "',UserName,MachineName,DT From " & _
'         " VoucherInfo Where VchrNo='" & VchrNo & "'"
'
'        con.Execute "DELETE FROM VouchersForAuthorization WHERE VchrNo='" & VchrNo & "'"
'        con.Execute "DELETE FROM VouchersForAuthorizationInfo WHERE VchrNo='" & VchrNo & "'"
'        .Close
'    End With
        
    con.Execute "UPDATE Cheque SET Description='Cancelled Due to Voucher Deletion',ChequeType='Canceled',Amount=0," & _
     "Bounced=0,Posted=0 WHERE SNo IN(SELECT CSNo FROM VouchersForAuthorization WHERE VchrNo='" & VchrNo & "')"
     
    con.Execute "DELETE FROM VouchersForAuthorization WHERE VchrNo='" & VchrNo & "'"
    con.Execute "DELETE FROM VouchersForAuthorizationInfo WHERE VchrNo='" & VchrNo & "'"
    
    '/\\/\/\\/\\/\\/\/\/\/\/\\/\/\/\/\/\/\/\/\/\/\///////////////////////////
    'Delete From Advances..If It's An Advance Payment Entry To Employee.
    con.Execute "DELETE FROM Advances WHERE AccVoucherNo='" & VchrNo & "'"
    '/\\/\/\\/\\/\\/\/\/\/\/\\/\/\/\/\/\/\/\/\/\/\///////////////////////////
    
    If Not DontStartNewTransaction Then con.CommitTrans

    Set rs = Nothing

    DeleteVoucherUnAuthorized = True

    Exit Function

err:
    MsgBox "Voucher Cannot Be Deleted Due To Error:" & err.Number & vbNewLine & err.Description
    DeleteVoucherUnAuthorized = False

End Function

Public Sub UpdateAdvEntry(AccNo As String, OpenBal As Double, Optional OpenDate As Date)

    On Error GoTo err
    Exit Sub            'This Is Done To Solve the Problem Of InEqual Credit Debit Totals In Trial Balance.
    Dim add As Boolean, SNo As Double
    Dim rs  As New ADODB.Recordset
    With rs
        .CursorLocation = adUseClient
        .Open "Select * From Vouchers Where Accno='" & AccNo & "' AND Sno=(Select Sno From AdvSno Where AccNo='" & AccNo & "')", con, adOpenForwardOnly, adLockOptimistic
        ' If Previous Entry Doesnot Exist Make New One
        
        If .RecordCount = 0 Then
            add = True
            .AddNew
            ![SNo] = getNextSno(OpenDate)
            SNo = ![SNo]
            ![AccNo] = AccNo
            ![VchrNo] = "OpeningBalance"
            ![Vdate] = OpenDate
        End If
 
        'If Exists Then Adjust Update it
 
        If OpenBal < 0 Then
            ![Credit] = Abs(OpenBal)
            ![Debit] = 0
        ElseIf OpenBal > 0 Then
            ![Credit] = 0
            ![Debit] = Abs(OpenBal)
        Else
            ![Credit] = 0
            ![Debit] = 0
        End If
        ![Balance] = OpenBal
        .Update
    End With
    'If Added then Update The AdvanceList
    If add Then con.Execute "Insert into AdvSno(Accno,Sno) Values ('" & AccNo & "'," & SNo & ")"
 
TheEnd:
    rs.Close
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
End Sub




Public Function IsEmpSalDue() As Boolean
   IsEmpSalDue = False

   Dim NumberOfEmployees As Long

   NumberOfEmployees = Val(con.Execute("Select Count(*) From VEmps Where Active=1 ").Fields(0) & "")
   
   If NumberOfEmployees = 0 Then Exit Function

   With con.Execute("Select Count(*) From Vouchers Where Month(Vdate)=" & Month(Date) & " And Accno in (Select AcrAccno From Employees)")
       IsEmpSalDue = (Val(.Fields(0) & "") < NumberOfEmployees)
   End With

End Function

Public Sub CheckForAutoVouchers()
   
   Dim ExpensesAreDue As Boolean
   
   'Check For Accured Expense Voucher
   
   Dim NumberOfExpenses As Long
      NumberOfExpenses = Val(con.Execute("Select Count(*) From ExpenseAccounts Where Amount>0").Fields(0) & "")
      
      ExpensesAreDue = Val(con.Execute("Select Count(*) From Vouchers Where Month(Vdate)=" & Month(Date) & " And Accno in (Select AcrExpAccno From ExpenseAccounts)").Fields(0) & "") < NumberOfExpenses
         
         
      If Not (ExpensesAreDue Or IsEmpSalDue) Then Exit Sub
         If MsgBox("Accured / Prepaid Expense For This Month Are Not Added." & vbNewLine & "Do You Want To Add Them Now ?", vbQuestion + vbYesNo) = vbYes Then
         frmMonthlyAccured.Show
      End If

End Sub

'Check That Depriciation Vouchers are Updated ....
Public Sub CheckDepriciation()



Dim idt As Date
Dim LastDepVchrDate As String

   If Val(con.Execute("Select Count(*) from VAssets Where PercentRate>0 and Balance>0").Fields(0) & "") = 0 Then Exit Sub
   
   With con.Execute("Select Max(VDate),Min(VDate)From Vouchers Where Accno in (Select DepAccno From VAssets Where PercentRate >0 And Balance>0) ")
      LastDepVchrDate = Format(.Fields(0) & "", "dd-MMM-yyyy")
      If LastDepVchrDate = "" Then
         LastDepVchrDate = Format(IIf(IsNull(.Fields(1)), YearStartDate, .Fields(1)), "dd-MMM-yyyy")
      End If
   End With
   
If CDate(LastDepVchrDate) < Date Then
   
   If MsgBox("Depriciation Of Assets Was Updated On " & LastDepVchrDate & vbNewLine & _
             "Do You Want To Update Depriciation Now ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            
   For idt = CDate(LastDepVchrDate) + 1 To Date Step 1
      Call AddDepriciationVoucher(idt)
   Next
   PlaySaveSound
   MsgBox "Depriciation Vouchers Added Successfully", vbInformation
End If
End Sub

Public Sub AddDepriciationVoucher(ByVal Dated As Date)
   Dim rs As New ADODB.Recordset
   
   'Select The Assets Which Are Not Depriciated on Given Date
   
   With rs
      .CursorLocation = adUseClient
      .Open "Select * From VAssets Where AccDepAccNo Not in (Select Distinct AccNo From Vouchers Where VDATE='" & _
       Format(Dated, "dd-MMM-yyyy") & "') and (PercentRate>0) and (OpenDate<='" & _
       Format(Dated, "dd-MMM-yyyy") & "') And (Balance>0)", con, adOpenForwardOnly, adLockReadOnly
       
      If .RecordCount = 0 Then Exit Sub
      
      Dim VNo As String
      Dim INextSno As Double
      Dim DepAmt As Double
      Dim AccNo As String
      Dim Bal As Double
      INextSno = getNextSno(Dated)
      VNo = getNextVchrNo(Dated, "JV")
      con.BeginTrans
      For i = 0 To .RecordCount - 1
      
         DepAmt = (![Balance]) * (![PercentRate] / 100 / 365)
         
         If DepAmt = 0 Then GoTo EndOfLoop
         AccNo = ![DepAccNo]
         Bal = getBalance(AccNo, Dated)
         Bal = Bal + DepAmt
        
         con.Execute "insert into Vouchers Values(" & INextSno & ",'" & Format(Dated, "dd-MMM-yyyy") & "','" & _
         VNo & "','" & AccNo & "','Depriciation',0," & DepAmt & "," & Bal & ",'',0)"
         
         Call adjustLedger(AccNo, Dated, INextSno, DepAmt)
         
         AccNo = ![AccDepAccno]
         Bal = getBalance(AccNo, Dated)
         
         Bal = Bal - DepAmt
         INextSno = INextSno + 1
         con.Execute "insert into Vouchers Values(" & INextSno & ",'" & Format(Dated, "dd-MMM-yyyy") & "','" & _
         VNo & "','" & AccNo & "','Depriciation'," & DepAmt & ",0," & Bal & ",'',0)"
         
         
         Call adjustLedger(AccNo, Dated, INextSno, -DepAmt)
         
         
         INextSno = INextSno + 1
         .MoveNEXT
EndOfLoop:
    Next i

   .Close
   
End With
Set rs = Nothing
con.CommitTrans
      
      
End Sub

Sub PrintVoucherOld(CR As CrystalReport, VchrNos As String, Optional DontPrintOnBlank As Boolean, Optional PrintDirect As Boolean = True)
 
    'Print On Blank Paper for all Other Than Helicon Sports
 
    If Not LCase(strCompany) Like "helicon*" Then DontPrintOnBlank = False
    
    With CR
        If PrintDirect Then
            .Destination = crptToPrinter
        Else
            .Destination = crptToWindow
        End If
        .ReportFileName = RptPath & IIf(DontPrintOnBlank, "\Voucher1.rpt", "\Voucher.rpt")
        .SubreportToChange = ""
        .Connect = ConnectStr
        .Formulas(0) = "Company='" & strCompany & "'"
        .Formulas(1) = IIf(DontPrintOnBlank, "", "Address='" & COMPANYADDRESS & "'")
        
        .SqlQuery = "Select * From VLedger Where VLedger.VchrNo in ('" & Replace(VchrNos, ",", "','") & "') Order By Debit Desc"
        If DontPrintOnBlank Then
            .SubreportToChange = "SubRptVchr.rpt"
            .Connect = ConnectStr
            
            .SubreportToChange = "SubRptVchr.rpt - 01"
            .Connect = ConnectStr
        End If
        
       .Action = 1
       .PageZoomNext
    End With
    Set CR = Nothing
    
End Sub

Sub PrintVoucher(CR As CrystalReport, VchrNos As String, Optional DontPrintOnBlank As Boolean, Optional PrintDirect As Boolean = True, Optional bUnAuthorized As Boolean = False, Optional iVoucherType As Integer = 0)
 
    'Print On Blank Paper for all Other Than Helicon Sports
 
    If Not LCase(strCompany) Like "helicon*" Then DontPrintOnBlank = False
    Dim rpt As CRAXDDRT.Report
    'Dim strTableName As String
    If bUnAuthorized = False Then
        If strCompany = "Instrumed" And (Left(VchrNos, 3) = "CPV" Or Left(VchrNos, 3) = "BPV" Or Left(VchrNos, 3) = "CRV") Then
            Set rpt = rptApp.OpenReport(RptPath & "\PV.rpt")
        Else
            If iVoucherType = 0 Then
                Set rpt = rptApp.OpenReport(RptPath & "\Voucher.rpt")
            ElseIf iVoucherType = 1 Then
                Set rpt = rptApp.OpenReport(RptPath & "\VoucherWO.rpt")
            End If
        End If
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\VoucherUA.rpt")
    End If
    Dim strQueryString As String
    With rpt
        .FormulaFields.GetItemByName("Company").Text = "'" & strCompany & "'"
        .FormulaFields.GetItemByName("Address").Text = "'" & COMPANYADDRESS & "'"
        
        'strQueryString = "SELECT * FROM VLedger Where VLedger.VchrNo in ('" & Replace(VchrNos, ",", "','") & "') Order By Debit Desc"
        If PrintDirect Then
            Call PrintReport("{VLedger.VchrNo}='" & VchrNos & "'", rpt)
        Else
            Dim frm As New frmPrevRpt
            frm.ShowReport "{VLedger.VchrNo}='" & VchrNos & "'", rpt
        End If
    End With
    Set rpt = Nothing
    
End Sub

Sub PrintChq(CR As CrystalReport, SerialNo)
 
    With CR
        .ReportFileName = RptPath & "\Cheque.rpt"
        .Destination = crptToWindow
        .Connect = ConnectStr
        .Formulas(0) = "Company='" & strCompany & "'"
        .Formulas(1) = IIf(DontPrintOnBlank, "", "Address='" & COMPANYADDRESS & "'")
        .SqlQuery = "Select * From VChqLedger Where Sno=" & SerialNo
        .Action = 1
        .PageZoomNext
    End With
    Set CR = Nothing
    
End Sub

Public Function VoucherExists(strVchrNo As String) As Boolean

    Dim iVchrCount As Integer
    iVchrCount = GetSingleStringValue("COUNT(VchrNo)", "Vouchers", " WHERE VchrNo='" & strVchrNo & "'")
    If iVchrCount = 0 Then
        VoucherExists = False
    Else
        VoucherExists = True
    End If
    
End Function
