Attribute VB_Name = "Banks"
Public Function AnyBankAccountDefined() As Boolean

If con.Execute("Select Count(*) From VBankAccounts").Fields(0) > 0 Then
   AnyBankAccountDefined = True
Else
   AnyBankAccountDefined = False
End If
End Function


 Private Function GetBankAccNo(ChqBookNo As Double) As String
   With con.Execute("Select Accno From ChqBooks Where ChqBookNo=" & ChqBookNo)
      If .RecordCount > 0 Then GetBankAccNo = .Fields(0) & "" Else GetBankAccNo = ""
   End With
End Function

Public Function GetNextChqNo(AccNo As String, Optional ChqBookNo As Double, Optional LastChqNo As Double) As String

    Dim rsChq As New ADODB.Recordset
    rsChq.Open "Select ChqBookNo,ChqNo From VChqList Where AccNo='" & AccNo & "' and  cast(ChqNo as numeric) >" & _
     LastChqNo & " AND Issued=0 Order by ChqNo Asc", con, adOpenForwardOnly
               
    If rsChq.RecordCount > 0 Then
        GetNextChqNo = rsChq![ChqNo]
        ChqBookNo = rsChq![ChqBookNo]
    Else
        GetNextChqNo = 0
    End If
   
    rsChq.Close
    Set rsChq = Nothing
    
End Function

Public Function GetDefaultBankAcc(Optional BankID As Double) As String

    Dim rs As New ADODB.Recordset
    rs.Open "Select * from VBankAccounts Where DefaultBank=1", con, adOpenForwardOnly
    If rs.RecordCount > 0 Then
        GetDefaultBankAcc = "[" & rs![AccNo] & "] " & rs![AccTitle]
        BankID = rs![BankID]
    Else
        GetDefaultBankAcc = ""
    End If
    rs.Close
    Set rs = Nothing
    
End Function


Public Function SelectBankAcc(Optional BankID As Double) As String
 With frmSelectBank
   
   If .TV.Nodes.count = 0 Then
      MsgBox "No Bank Account Has Been Defined.", vbInformation
      GoTo TheEnd
   End If
   .Show 1
   SelectBankAcc = .BankAcc
   BankID = .BankID
 End With
 
TheEnd:
Unload frmSelectBank
End Function


Public Function ValidChqNo(AccNo As String, ChqNo As String, Optional Issued As Boolean, Optional ChqBookNo As Double) As Boolean
   Dim rs As New ADODB.Recordset
   rs.CursorLocation = adUseClient
  
  rs.Open "Select * From VChqList Where (AccNo='" & AccNo & _
  "') and  (ChqNo='" & ChqNo & "')", con, adOpenForwardOnly
  
  If rs.RecordCount > 0 Then
      If rs![Issued] Then
         ValidChqNo = False
         Issued = True
      Else
         ValidChqNo = True
         ChqBookNo = rs![ChqBookNo]
         Issued = False
      End If
  Else
   ValidChqNo = False
   Issued = False
  End If
  
  rs.Close
  Set rs = Nothing
  
End Function



Public Function HasRunnigChqBook(AccNo As String) As Boolean

    Dim rs As New ADODB.Recordset
    rs.Open "select count(*) from VChqList Where AccNo='" & AccNo & "'", con, adOpenForwardOnly
   
    If rs.Fields(0) > 0 Then
        HasRunnigChqBook = True
    Else
        HasRunnigChqBook = False
    End If
  
    rs.Close
    Set rs = Nothing

End Function

Public Sub IssueChq(ChqBookNo As Double, ChqNo As String)
    con.Execute "Update ChqList Set Issued=1 Where ChqBookNo=" & ChqBookNo & " and chqNo='" & ChqNo & "'"
End Sub

Public Function AllowOD(AccNo As String) As Boolean

    With con.Execute("Select AllowOD From BankAccounts Where AccNo='" & AccNo & "'")
        If .RecordCount = 0 Then
            AllowOD = False
            Exit Function
        Else
            AllowOD = IIf(IsNull(![AllowOD]), False, ![AllowOD])
        End If
   
    End With
    
End Function

Public Function AllowODAmt(AccNo As String) As Double

    With con.Execute("Select AllowOD,ODAmt From BankAccounts Where AccNo='" & AccNo & "'")
        If .RecordCount = 0 Then
            AllowODAmt = 0
            Exit Function
        Else
            If !AllowOD Then
                AllowODAmt = Val(!ODAmt & "")
            Else
                AllowODAmt = 0
            End If
        End If
   
    End With
    
End Function

