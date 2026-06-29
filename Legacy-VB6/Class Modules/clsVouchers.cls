VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVouchers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'I'm trying to create a class for voucher posting....
'an object will be created from this class
'which will have an AddHead Function to Add AccNo & Transaction Amount
'Class will have a property for Voucher Type.
'I'm trying to figure out about how to add a custom collection to this class which will have
'AccNo & Transaction Amout Properties...
'Also I'll Add functionality which will show Expected Values of Properties if any....

Option Explicit

Private VoucherHeads As New Collection
Private m_strVoucherType As String
Private m_DTVoucherDate As Date

Public Function AddVoucherHead(p_strAccNo As String, p_dTransAmt As Double, p_strDescription As String, Optional p_cSNo As Long) As Boolean
    
    'Check If Account No. Is Valid
    Dim rs As New ADODB.Recordset
    Dim bInvalid As Boolean
    With rs
        .Open "Select AccNo From Accounts Where AccNo='" & p_strAccNo & "' AND Parent=0 AND Active=1", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            bInvalid = True
        End If
        .Close
    End With
    Set rs = Nothing
    
    If bInvalid Then
        'MsgBox "Invalid Account No.", vbInformation
        'Exit Function
        err.Raise vbObjectError, , "Invalid Acc No.:" & p_strAccNo
    End If
     
    If p_dTransAmt = 0 Then
        'MsgBox "Invalid Amount.", vbInformation
        'Exit Function
        err.Raise vbObjectError, , "Invalid Amount"
    End If
    
    Dim p_VoucherHead As New clsVouchersSub
    p_VoucherHead.strAccNo = p_strAccNo
    p_VoucherHead.dTransAmt = p_dTransAmt
    p_VoucherHead.StrDescription = p_strDescription
    p_VoucherHead.lCSNo = p_cSNo
    
    VoucherHeads.add p_VoucherHead
    
    Set p_VoucherHead = Nothing
    
    AddVoucherHead = True
    
End Function


Public Property Get VoucherType() As String
    VoucherType = m_strVoucherType
End Property

Public Property Let VoucherType(ByVal p_strVoucherType As String)
    m_strVoucherType = p_strVoucherType
End Property

Public Property Get VoucherDate() As Date
    
    VoucherDate = m_DTVoucherDate
   
End Property

Public Property Let VoucherDate(ByVal p_DTVoucherDate As Date)

   
    
    Dim rs As New ADODB.Recordset
    Dim bPeriodOpen As Boolean
    bPeriodOpen = True
'    With rs
'        .Open "Select EntryID From FinancialPeriods Where '" & p_DTVoucherDate & "' Between StartDate AND FinishDate AND PeriodOpen=1", con, adOpenForwardOnly, adLockReadOnly
'        If .EOF Then
'            bPeriodOpen = False
'        Else
'            bPeriodOpen = True
'        End If
'        .Close
'    End With
    
    If bPeriodOpen = False Then
        'MsgBox "Invalid Date!!!" & vbNewLine & "Period Closed or Doesn't Exist.", vbInformation
        'Exit Property
        Call err.Raise(vbObjectError, , "Invalid Date")
        
    End If
    
    m_DTVoucherDate = p_DTVoucherDate
    
    
End Property

Public Function PostVoucher(Optional bPostedThroughJVForm As Boolean = False, Optional bDontStartTransaction As Boolean = False, Optional strVchrNo As String = "", Optional bAdjustDecimalDiff As Boolean = False) As String

    If ValidateVoucher() = False Then
        If bAdjustDecimalDiff Then
            Call AdjustDecimalDiff
            If ValidateVoucher() = False Then
                Call err.Raise(vbObjectError, , "Can't Validate Voucher")
            End If
        Else
            Call err.Raise(vbObjectError, , "Can't Validate Voucher")
        End If
    End If
    
    'Ready to Post Voucher Now....
    'Get New VoucherNo....
    'Dim strVchrNo As String
    Dim INextSno As Double
    If strVchrNo = "" Then
        strVchrNo = getVoucherNo
    End If
    INextSno = getNextSno(m_DTVoucherDate)
    
    Dim objVchrs As clsVouchersSub
    
    Dim i As Integer
    If bDontStartTransaction = False Then Call StartTrans(con)
    
    For i = 1 To VoucherHeads.count
        'Post Debit Heads First....
        Set objVchrs = VoucherHeads(i)
        If objVchrs.dTransAmt > 0 Then
            Call PostDebitSide(objVchrs.strAccNo, objVchrs.dTransAmt, objVchrs.StrDescription, INextSno, strVchrNo, objVchrs.lCSNo)
            INextSno = INextSno + 1
        End If
    Next
     
    For i = 1 To VoucherHeads.count
        'Post Credit Heads Now....
        Set objVchrs = VoucherHeads(i)
        If objVchrs.dTransAmt < 0 Then
            Call PostCreditSide(objVchrs.strAccNo, Abs(objVchrs.dTransAmt), objVchrs.StrDescription, INextSno, strVchrNo, objVchrs.lCSNo)
            INextSno = INextSno + 1
        End If
    Next
    con.Execute "INSERT INTO VoucherInfo(VchrNo,UserName,MachineName,DT,PostedThroughJVForm) Values('" & strVchrNo & _
     "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "'," & Abs(bPostedThroughJVForm) & ")"
     
    If bDontStartTransaction = False Then con.CommitTrans
    'con.Execute "Commit Transaction"
    PostVoucher = strVchrNo
    
End Function

Private Sub PostDebitSide(strAccNo As String, dAmt As Double, strDesc As String, SNo As Double, strVchrNo As String, lCSNo As Long)

    Dim Bal As Double
   
'    SnoToUse = INextSno
      
    Bal = getBalance(strAccNo, m_DTVoucherDate)
    Bal = Bal + dAmt
   
    con.Execute "Insert Into Vouchers(SNo,VDate,VchrNo,AccNo,Description,Debit,Credit,Balance,DpstSlip,CSNo) " & _
     "Values(" & SNo & ",'" & Format(m_DTVoucherDate, "dd-MMM-yyyy") & "','" & strVchrNo & "','" & strAccNo & _
     "','" & strDesc & "'," & dAmt & ",0," & Bal & ",''," & lCSNo & ")"
          
   Call adjustLedger(strAccNo, m_DTVoucherDate, SNo, dAmt)
   
End Sub

Private Sub PostCreditSide(strAccNo As String, dAmt As Double, strDesc As String, SNo As Double, strVchrNo As String, lCSNo As Long)

    Dim Bal As Double
   
'    SnoToUse = INextSno
      
    Bal = getBalance(strAccNo, m_DTVoucherDate)
    Bal = Bal - dAmt
   
    con.Execute "INSERT INTO Vouchers(SNo,VDate,VchrNo,AccNo,Description,Debit,Credit,Balance,DpstSlip,CSNo) " & _
     "Values(" & SNo & ",'" & Format(m_DTVoucherDate, "dd-MMM-yyyy") & "','" & strVchrNo & "','" & strAccNo & _
     "','" & strDesc & "',0," & dAmt & "," & Bal & ",''," & lCSNo & ")"
          
   Call adjustLedger(strAccNo, m_DTVoucherDate, SNo, -dAmt)
   
End Sub



Private Function ValidateVoucher(Optional bAdjustDecimalDiff As Boolean = False) As Boolean

    Dim i As Integer
    Dim dDebitTotal As Double, dCreditTotal As Double
    For i = 1 To VoucherHeads.count
        If VoucherHeads(i).dTransAmt < 0 Then
            dCreditTotal = dCreditTotal + Abs(VoucherHeads(i).dTransAmt)
        Else
            dDebitTotal = dDebitTotal + VoucherHeads(i).dTransAmt
        End If
    Next
    If Val(dCreditTotal) = Val(dDebitTotal) Then
        ValidateVoucher = True
    Else
        ValidateVoucher = False
    End If
    
End Function


Private Function getVoucherNo() As String
 
    Dim rsVchr As New ADODB.Recordset

    Dim MonthlyVF As Boolean
    
    If Len(m_strVoucherType) = 2 Then
        rsVchr.Open "SELECT MAX(RIGHT(VchrNo,3)) FROM Vouchers WHERE LEFT(VchrNo,10) = '" & _
         VoucherType & "-" & Format(m_DTVoucherDate, "YYMMdd") & "-'", con, adOpenForwardOnly, adLockReadOnly
    Else
        rsVchr.Open "SELECT max(right(VchrNo,3)) FROM Vouchers WHERE LEFT(VchrNo,11) = '" & _
         VoucherType & "-" & Format(m_DTVoucherDate, "YYMMdd") & "-'", con, adOpenForwardOnly, adLockReadOnly
    End If
    getVoucherNo = VoucherType & "-" & Format(m_DTVoucherDate, "YYMMdd") & "-" & Format((Val(rsVchr.Fields(0) & "") + 1), "000")

    rsVchr.Close

    Set rsVchr = Nothing

End Function

Private Function AdjustDecimalDiff() As Boolean
    
    Dim i As Integer
    Dim dDebitTotal As Double, dCreditTotal As Double
    For i = 1 To VoucherHeads.count
        If VoucherHeads(i).dTransAmt < 0 Then
            dCreditTotal = dCreditTotal + Abs(VoucherHeads(i).dTransAmt)
        Else
            dDebitTotal = dDebitTotal + VoucherHeads(i).dTransAmt
        End If
    Next
    
    Dim dDebitDiff As Double, dLargestValue As Double, iLargestHead As Integer
    dLargestValue = 0
    dDebitDiff = Val(dDebitTotal) - Val(dCreditTotal)
    If Round(dDebitDiff) = 0 Then 'Diff is in Decimals, Adjust it if bAdjustDecmalDiff=True....
        If dDebitDiff > 0 Then  'Debit is More....Adjust Diff. by Decreasing or Increasing the Largest Value
            For i = 1 To VoucherHeads.count
                If VoucherHeads(i).dTransAmt > 0 Then
                    If VoucherHeads(i).dTransAmt > dLargestValue Then
                        dLargestValue = VoucherHeads(i).dTransAmt
                        iLargestHead = i
                    End If
                End If
            Next
            VoucherHeads(iLargestHead).dTransAmt = VoucherHeads(iLargestHead).dTransAmt - dDebitDiff
        Else
            For i = 1 To VoucherHeads.count
                If VoucherHeads(i).dTransAmt < 0 Then
                    If Abs(VoucherHeads(i).dTransAmt) > dLargestValue Then
                        dLargestValue = Abs(VoucherHeads(i).dTransAmt)
                        iLargestHead = i
                    End If
                End If
            Next
            VoucherHeads(iLargestHead).dTransAmt = VoucherHeads(iLargestHead).dTransAmt - dDebitDiff
        End If
    End If
        
End Function
