Attribute VB_Name = "PayrollVouchers"

'AccNos as defined in c:\Accounts 2002-2003.ctw


'Const CashInHand = "15-001"


'Director's AccNos
Const DLongAdvAccNo = "13-002"
Const DShortAdvAccNo = "15-010-002"
'Const DSalaryAccNo = "42-001"
Const DSalaryAccNo = "42-001-001"
'Staff Members's AccNos
   Const LongAdvAccNo = "13-003-001"
   Const ShortAdvAccNo = "15-010-003-001"
   'Const SalaryAccNo = "42-002"
Const SalaryAccNo = "42-002-001"
'Production Staff Members's AccNos
   Const PLongAdvAccNo = "13-003-002"
   Const PShortAdvAccNo = "15-010-003-002"
  ' Const PSalaryAccNo = "41-003-001"
 Const PSalaryAccNo = "41-003-006"
   Const TaxAccNo = "45-003"

'Over Time AccNos
'Const OTimeAccNo = "41-003-002"
Const OTimeAccNo = "41-003-002-001"



Public Function addAdvanceVoucher(StaffAmt As Double, DirctorAmt As Double, ProductionAmt As Double, Dated As Date, Optional LongTerm As Boolean) As Boolean
    
    Dim i As Long, INextSno As Double, Bal As Double, Amt As Double
    Dim VchrNo As String, theDesc As String, AccNo As String
    
      
    'Charge All Amount to Cash
    Amt = StaffAmt + ProductionAmt + DirectorAmt
    
    If Amt = 0 Then Exit Function
      
    
    
     'ConAcc.BeginTrans
     
     'Get The Required Values
     
     INextSno = getNextSno(Dated)
     VchrNo = getNextVchrNo(Dated, "PV")
     
     
    'Give The Description
     theDesc = IIf(LongTerm, "Long Term Advance To Employees", "Short Term Advance To Employees")
     
     
     
    ''''''''''''''''Charge To  Cash Entry'''''''''''''
    
    AccNo = CashInHand
    
    Bal = getBalance(AccNo, Dated)
    
    Bal = Bal - Amt
        
    'Credit Entry
    ConAcc.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & theDesc & "',0," & Amt & "," & Bal & ",'',0)"
           
    
    Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
     
     
     INextSno = INextSno + 1
     
     
     
     
    '''''''''''''''''' D E B I T   E N T R I E S '''''''''''''
     
     
     
     
     
If StaffAmt = 0 Then GoTo DIRECTOR_ENTRY
      
      'Determine the Account no for Staff Member
      'Long / Short Tem Advance
      
      AccNo = IIf(LongTerm, LongAdvAccNo, ShortAdvAccNo)
     
      Bal = getBalance(AccNo, Dated)
      
      Amt = StaffAmt
      Bal = Bal + Amt
    
    
      'Debit To Staff Long/Short Advance Account
      
      ConAcc.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & theDesc & "'," & Amt & ",0," & Bal & ",'',0)"
        
     
        Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
     
         INextSno = INextSno + 1
     
DIRECTOR_ENTRY:

If DirectorAmt = 0 Then GoTo PRODUCTION_ENTRY
      'Determine the Account no for Directors
      'Long Term Or Short Tem Advance
      
      AccNo = IIf(LongTerm, DLongAdvAccNo, DShortAdvAccNo)
     
      Bal = getBalance(AccNo, Dated)
      
      
      Amt = DirectorAmt
      Bal = Bal + Amt
    
    
      'Debit To Staff Long/Short Advance Account
      
      ConAcc.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & theDesc & "'," & Amt & ",0," & Bal & ",'',0)"
        
     
        Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
        
         INextSno = INextSno + 1
         
         
PRODUCTION_ENTRY:
         'Determine the Account no for Directors
      'Long Term Or Short Tem Advance
If ProductionAmt = 0 Then GoTo END_OF_VOUCHER

      AccNo = IIf(LongTerm, PLongAdvAccNo, PShortAdvAccNo)
     
      Bal = getBalance(AccNo, Dated)
      
      
      Amt = ProductionAmt
      Bal = Bal + Amt
    
    
      'Debit To Staff Long/Short Advance Account
      
      ConAcc.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & theDesc & "'," & Amt & ",0," & Bal & ",'',0)"
        
     
        Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
        
         
END_OF_VOUCHER:

    'ConAcc.CommitTrans
    addAdvanceVoucher = True
    Exit Function
err:
   ' ConAcc.RollbackTrans
    addAdvanceVoucher = False
    MsgBox "Voucher For Employee Advance Cannot Be Added Due To Error:" & _
            vbNewLine & err.Description
            
End Function


Public Function AddAdvClearVoucher(StaffAmt As Double, DirctorAmt As Double, ProductionAmt As Double, Dated As Date, Optional LongTerm As Boolean, Optional AdvFor As String) As Boolean
    Dim i As Long, INextSno As Double, Bal As Double, Amt As Double
    Dim VchrNo As String, theDesc As String, AccNo As String
    
      
    'All Back Amount to Cash
    Amt = StaffAmt + ProductionAmt + DirectorAmt
    
    If Amt = 0 Then Exit Function
      
    
    
     'ConAcc.BeginTrans
     
     'Get The Required Values
     
     INextSno = getNextSno(Dated)
     VchrNo = getNextVchrNo(Dated, "RV")
     
     
    'Give The Description
     theDesc = IIf(LongTerm, "Long Term Loan Deductions", "Short Term Loan Deductions") & " for " & AdvFor
     
     
     
    ''''''''''''''''Cash Entry'''''''''''''
    
    AccNo = CashInHand
    
    Bal = getBalance(AccNo, Dated)
    
    Bal = Bal + Amt
    
    
    'Credit Entry
    ConAcc.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & theDesc & "'," & Amt & ",0," & Bal & ",'',0)"
        
    
    
    
    Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
     
     
     INextSno = INextSno + 1
     
     
     
     
    '''''''''''''''''' C R E D I T   E N T R I E S '''''''''''''
     
     
     
     
     
If StaffAmt = 0 Then GoTo DIRECTOR_ENTRY
      
      'Determine the Account no for Staff Member
      'Long / Short Tem Advance
      
      AccNo = IIf(LongTerm, LongAdvAccNo, ShortAdvAccNo)
     
      Bal = getBalance(AccNo, Dated)
      
      Amt = StaffAmt
      Bal = Bal - Amt
    
    theDesc = "Deductions for " & AdvFor

      'Debit To Staff Long/Short Advance Account
      
      ConAcc.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & theDesc & "',0," & Amt & "," & Bal & ",'',0)"
        
     
        Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
     
         INextSno = INextSno + 1
     
DIRECTOR_ENTRY:

If DirectorAmt = 0 Then GoTo PRODUCTION_ENTRY
      'Determine the Account no for Directors
      'Long Term Or Short Tem Advance
      
      AccNo = IIf(LongTerm, DLongAdvAccNo, DShortAdvAccNo)
     
      Bal = getBalance(AccNo, Dated)
      
      
      Amt = DirectorAmt
      Bal = Bal - Amt
    
    
      'Credit To Staff Long/Short Advance Account
      
      ConAcc.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & theDesc & "',0," & Amt & "," & Bal & ",'',0)"
        
     
        Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
        
         INextSno = INextSno + 1
         
         
PRODUCTION_ENTRY:
         'Determine the Account no for Directors
      'Long Term Or Short Tem Advance
If ProductionAmt = 0 Then GoTo END_OF_VOUCHER

      AccNo = IIf(LongTerm, PLongAdvAccNo, PShortAdvAccNo)
     
      Bal = getBalance(AccNo, Dated)
      
      
      Amt = ProductionAmt
      Bal = Bal - Amt
    
    
      'Credit To Staff Long/Short Advance Account
      
      ConAcc.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & theDesc & "',0," & Amt & "," & Bal & ",'',0)"
        
     
        Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
        
         
END_OF_VOUCHER:

    'ConAcc.CommitTrans
    AddAdvClearVoucher = True
    Exit Function
        
End Function

Public Function addSalaryVoucher(StaffAmt As Double, DirectorAmt As Double, ProductionAmt As Double, Dated As Date, STax As Double, DTax As Double, PTax As Double, Optional SalFor As String) As Boolean
    
    Dim i As Long, INextSno As Double, Bal As Double, Amt As Double
    Dim VchrNo As String, theDesc As String, AccNo As String
    
      
    'Charge All Amount to Cash
    Amt = StaffAmt + ProductionAmt + DirectorAmt
    
    If Amt = 0 Then Exit Function
    
     'ConAcc.BeginTrans
     
     'Get The Required Values
     
     INextSno = getNextSno(Dated)
     VchrNo = getNextVchrNo(Dated, "PV")
     
     

     
     
     
     
    ''''''''''''''''Charge To  Cash Entry'''''''''''''
    
    AccNo = CashInHand
    
    Bal = getBalance(AccNo, Dated)
    
    Bal = Bal - Amt
    
    theDesc = "Payment For " & SalFor
    
    'Credit Entry
    ConAcc.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & theDesc & "',0," & Amt & "," & Bal & ",'',0)"
        
    
    
    
    Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
     
     
     INextSno = INextSno + 1
     
     
     
     
    '''''''''''''''''' D E B I T   E N T R I E S '''''''''''''
     
     
     
     
     
If StaffAmt = 0 Then GoTo DIRECTOR_ENTRY
      
      'Determine the Account no for Staff Member
      'Long / Short Tem Advance
      
      AccNo = SalaryAccNo
     
      Bal = getBalance(AccNo, Dated)
      
      Amt = StaffAmt - STax
      Bal = Bal + Amt
      
      
    
      'Debit To Staff Long/Short Advance Account
      
      ConAcc.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & theDesc & "'," & Amt & ",0," & Bal & ",'',0)"
        
     
        Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
     
         INextSno = INextSno + 1
     
DIRECTOR_ENTRY:

If DirectorAmt = 0 Then GoTo PRODUCTION_ENTRY
      'Determine the Account no for Directors
      
      
      
      AccNo = DSalaryAccNo
     
      Bal = getBalance(AccNo, Dated)
      
      
      Amt = DirectorAmt - DTax
      Bal = Bal + Amt
    
    
      'Debit To Staff Long/Short Advance Account
      
      ConAcc.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & theDesc & "'," & Amt & ",0," & Bal & ",'',0)"
        
     
        Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
        
         INextSno = INextSno + 1
         
         
PRODUCTION_ENTRY:
         'Determine the Account no for Directors
      
If ProductionAmt = 0 Then GoTo TAX_ENTRY

      
      
      AccNo = PSalaryAccNo
     
      Bal = getBalance(AccNo, Dated)
      
      
      Amt = ProductionAmt - PTax
      Bal = Bal + Amt
    
    
      'Debit To Staff Long/Short Advance Account
      
      ConAcc.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & theDesc & "'," & Amt & ",0," & Bal & ",'',0)"
        
     
        Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
         INextSno = INextSno + 1
         
TAX_ENTRY:

   'Calculate Total Tax
   Amt = STax + DTax + PTax
   
   If Amt = 0 Then GoTo END_OF_VOUCHER

      theDesc = "Tax Deduction For " & SalFor
      
      AccNo = TaxAccNo
     
      Bal = getBalance(AccNo, Dated)
      
      
      
      Bal = Bal + Amt
    
    
      'TAX ACCOUNT ENTRY
      
      ConAcc.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
                     Dated & "','" & VchrNo & "','" & AccNo & _
                     "','" & theDesc & "'," & Amt & ",0," & Bal & ",'',0)"
        
     
        Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
        

END_OF_VOUCHER:

    'ConAcc.CommitTrans
    addSalaryVoucher = True
    Exit Function
        
End Function

