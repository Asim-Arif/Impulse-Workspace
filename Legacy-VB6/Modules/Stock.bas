Attribute VB_Name = "Stock"
Public Function getStockBalance(ITemID As String, OnDate As Date) As Double

On Error GoTo err

 Dim rsBal As New ADODB.Recordset, SubQry As String
 Dim MaxDate As Date

' Get The date Of Last Entry
rsBal.Open "Select Max(Dated) from VInOutDetail Where ItemID='" & ITemID & "'", con, adOpenForwardOnly, adLockOptimistic

If Not IsNull(rsBal.Fields(0)) Then
 MaxDate = rsBal.Fields(0)
Else
' If No Entry found  Then Current Date is Maximum Date
 MaxDate = ServerDate
End If
' Close Recordset
rsBal.Close

' If This Entry is In Previous Dates then
' Get Stock Balance From Previous entries

If OnDate < MaxDate Then
  SubQry = "Select max(SNo) From VInOutDetail where Dated<='" & OnDate & "' and ItemID='" & ITemID & "'"
  rsBal.Open "Select inhand From VInOutDetail where Sno=(" & SubQry & ")", con, adOpenForwardOnly, adLockOptimistic
Else
'Else Get Cuurent Stock
  rsBal.Open "Select InHand from Items Where ItemID='" & ITemID & "'", con, adOpenForwardOnly, adLockReadOnly
End If


If rsBal.RecordCount = 0 Then
   rsBal.Close

   rsBal.Open "Select InHand from Items Where ItemID='" & ITemID & "'", con, adOpenForwardOnly, adLockReadOnly
 
  
End If

getStockBalance = Val(rsBal.Fields(0) & "")

Set rsBal = Nothing


Exit Function
err:
MsgBox "Followinfg error Occured in getStockBalance" & vbNewLine & err.Description
End Function



Public Function GetExistingSnosStr(ITemID As String) As String
GetExistingSnosStr = ""
   Dim rs As New ADODB.Recordset
   rs.CursorLocation = adUseClient
   rs.Open "Select Sno From ItemSno Where ItemID='" & ITemID & "' and not out", con, adOpenForwardOnly
   If rs.RecordCount = 0 Then Exit Function
   GetExistingSnosStr = "¤" & rs.GetString(adClipString, , , "¤", "")
   rs.Close
   Set rs = Nothing
End Function


Public Function GetNextStockID(Dated As Date, IDPrefix As String)
    'e.g. SAL-0701-001
    GetNextStockID = IDPrefix & "-" & Format(Dated, "MMdd") & "-" & Format(GetMax("Cast(Right(EntryID,3) as numeric)", "StockInOut", "Where left(EntryID,8)='" & IDPrefix & "-" & Format(Dated, "MMdd") & "'") + 1, "000")
End Function


Public Function getItemBalance(ITemID As String, OnDate As Date) As Double

    On Error GoTo err

    Dim rsBal As New ADODB.Recordset, SubQry As String
    Dim MaxDate As Date
    
    'Get The date Of Last Entry
    rsBal.Open "Select Max(DT) from ItemLedgerSimple Where ItemID='" & ITemID & "'", con, adOpenForwardOnly, adLockReadOnly
    
    If Not IsNull(rsBal.Fields(0)) Then
        MaxDate = rsBal.Fields(0)
    Else
        'If No Entry found Then Current Date is Maximum Date
        MaxDate = Date
    End If
    'Close Recordset
    rsBal.Close

    'If This Entry is In Previous Dates then
    'Get Balance From Previous entries

    Dim TransAmt As Double
    TransAmt = 0
    If OnDate < MaxDate Then
        'SubQry = "Select Max(EntryID) From ItemLedgerSimple Where DT<='" & Format(OnDate, "dd-MMM-yyyy") & "' and ItemID='" & ItemID & "'"
        SubQry = "Select Max(EntryID) From ItemLedgerSimple Where DT<='" & Format(OnDate, "dd-MMM-yyyy") & "' and ItemID='" & ITemID & "' AND DT=(Select Max(DT) From ItemLedgerSimple Where DT<='" & Format(OnDate, "dd-MMM-yyyy") & "' and ItemID='" & ITemID & "')"
        rsBal.Open "Select Balance From ItemLedgerSimple where EntryID=(" & SubQry & ")", con, adOpenForwardOnly, adLockOptimistic
        If rsBal.RecordCount > 0 Then TransAmt = 0 'rsBal![TransAmt]
    Else
        'Else Get Current Balance
        rsBal.Open "Select InHand from Items Where ItemID='" & ITemID & "'", con, adOpenForwardOnly, adLockReadOnly
    End If

    If rsBal.RecordCount = 0 Then
        rsBal.Close
        rsBal.Open "Select OpenBal from Items Where ItemID='" & ITemID & "'", con, adOpenForwardOnly, adLockReadOnly
    End If

    getItemBalance = Val(rsBal.Fields(0) & "")

    Set rsBal = Nothing

    Exit Function
err:
    MsgBox "Following error Occured in getBalance" & vbNewLine & err.Description
    
End Function

Public Sub adjustItemLedger(ITemID As String, StartDate As Date, QtyToAdd As Double)

    con.Execute "UPDATE Items SET InHand=InHand+" & QtyToAdd & " WHERE ItemID='" & ITemID & "'"

    Dim SQL As String
    'Update Ledger Balances
    SQL = "UPDATE ItemLedgerSimple SET Balance=Balance+(" & QtyToAdd & ") WHERE ItemID='" & ITemID & _
       "' AND (DT>'" & Format(StartDate, "dd-MMM-yyyy") & "')"
 
    con.Execute SQL
 
End Sub

Public Function UpdateMaterialLocation(strMaterialID As String, lLocationID As Long, dQty As Double, lDetailEntryID As Long, Optional bReturn As Boolean = False, Optional lSheets As Long, Optional strPORefNo As String, Optional strBatchNo As String, Optional strLotNo As String)

    Dim cmd As New ADODB.Command
    With cmd
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        If bReturn Then
            .CommandText = "UpdateMaterialLocationFromReturn"
        Else
            .CommandText = "UpdateMaterialLocation"
        End If
        .Parameters("@MaterialID").Value = strMaterialID
        .Parameters("@LocationID").Value = lLocationID
        .Parameters("@Qty").Value = dQty
        .Parameters("@Sheets").Value = lSheets
        .Parameters("@PORefNo").Value = strPORefNo
        .Parameters("@DetailEntryID").Value = lDetailEntryID
        
        .Parameters("@BatchNo").Value = strBatchNo
        .Parameters("@LotNo").Value = strLotNo
        .Execute
    End With
    
End Function

Public Function UpdateItemLocation(strItemID As String, lLocationID As Long, dQty As Double, Optional iSample_Type As Integer = -1, Optional strCustCode As String)

    Dim cmd As New ADODB.Command
    With cmd
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        
        .CommandText = "UpdateItemLocation_SP"
        
        .Parameters("@MaterialID").Value = strItemID
        .Parameters("@LocationID").Value = lLocationID
        .Parameters("@Qty").Value = dQty
        .Parameters("@CustCode").Value = strCustCode
        
        Dim ir As Integer
        .Execute ir
    End With
    
End Function


Public Function DeleteIssuance(strIssNo As String) As Boolean
    'Ist Update the Locations...
    '2nd Update the Rcvings...
    '3rd Update RM Table
    '4th Delete from RawMaterialIssuanceDetail
    '5th Delete from RawMaterialIssuance
    
    DeleteIssuance = False
    
    If getDBPassword("DeleteRMIssuance") = False Then
        Exit Function
    End If
    
    Dim cmd As New ADODB.Command
    With cmd
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        
        .CommandText = "UpdateMaterialLocation_Deletion_SP"
        
        .Parameters("@IssNo").Value = strIssNo
        Dim ir As Integer
        .Execute ir
        
        If ir > 0 Then DeleteIssuance = True
    End With
    
End Function

Public Function getIssuanceNo(myDT) As String
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Max(Cast(Right(IssNo,3) As Int)) From IssItemsSimple Where CAST(CONVERT(Varchar(50),DT,6) AS DATETIME)='" & myDT & "'", con, adOpenForwardOnly, adLockReadOnly
        
        getIssuanceNo = "ISS-" & Format(myDT, "yyMMdd-") & Format(Val(.Fields(0) & "") + 1, "000")
        
        .Close
    End With
    Set rs = Nothing
End Function

Public Function GetRcvingNo(myDT) As String
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Max(Cast(Right(RcvNo,3) As Int)) From RcvItemsSimple Where CAST(CONVERT(Varchar(50),DT,6) AS DATETIME)='" & myDT & "'", con, adOpenForwardOnly, adLockReadOnly
        
        GetRcvingNo = "RCV-" & Format(myDT, "yyMMdd-") & Format(Val(.Fields(0) & "") + 1, "000")
        
        .Close
    End With
    Set rs = Nothing
End Function

