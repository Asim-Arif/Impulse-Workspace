Attribute VB_Name = "mod_Encryption"
Option Explicit


Public Function CryptRC4(sText As String, sKey As String) As String

'    Dim baS(0 To 255) As Byte
'    Dim baK(0 To 255) As Byte
'    Dim bytSwap     As Byte
'    Dim lI          As Long
'    Dim lJ          As Long
'    Dim lIdx        As Long
'
'    For lIdx = 0 To 255
'        baS(lIdx) = lIdx
'        baK(lIdx) = Asc(Mid$(sKey, 1 + (lIdx Mod Len(sKey)), 1))
'    Next
'    For lI = 0 To 255
'        lJ = (lJ + baS(lI) + baK(lI)) Mod 256
'        bytSwap = baS(lI)
'        baS(lI) = baS(lJ)
'        baS(lJ) = bytSwap
'    Next
'    lI = 0
'    lJ = 0
'    For lIdx = 1 To Len(sText)
'        lI = (lI + 1) Mod 256
'        lJ = (lJ + baS(lI)) Mod 256
'        bytSwap = baS(lI)
'        baS(lI) = baS(lJ)
'        baS(lJ) = bytSwap
'        CryptRC4 = CryptRC4 & Chr$((pvCryptXor(baS((CLng(baS(lI)) + baS(lJ)) Mod 256), Asc(Mid$(sText, lIdx, 1)))))
'    Next
    CryptRC4 = sText
End Function

Private Function pvCryptXor(ByVal lI As Long, ByVal lJ As Long) As Long
    If lI = lJ Then
        pvCryptXor = lJ
    Else
        pvCryptXor = lI Xor lJ
    End If
End Function

Public Function ToHexDump(sText As String) As String
    Dim lIdx            As Long

'    For lIdx = 1 To Len(sText)
'        ToHexDump = ToHexDump & Right$("0" & Hex(Asc(Mid(sText, lIdx, 1))), 2)
'    Next
    ToHexDump = sText
End Function

Public Function FromHexDump(sText As String) As String
    Dim lIdx            As Long

'    For lIdx = 1 To Len(sText) Step 2
'        FromHexDump = FromHexDump & Chr$(CLng("&H" & Mid(sText, lIdx, 2)))
'    Next
    FromHexDump = sText
End Function

Public Sub EncryptSensitiveData()
    
    Dim strCondition As String
    
    If CBool(getGeneralDataValue("Encryption")) = False Then
        strCondition = ""
    Else
        strCondition = " WHERE PasswordEncrypted=0"
    End If
    
    Dim rs As New ADODB.Recordset
    Dim strOriginal As String, strEncrypted As String
    With rs
        .Open "SELECT * FROM Security" & strCondition, con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            strOriginal = !Password & ""
            strEncrypted = ToHexDump(CryptRC4(strOriginal, "AwmWrangler"))
            
            con.Execute "UPDATE Security SET Password='" & strEncrypted & "',PasswordEncrypted=1 WHERE EntryID=" & !EntryID
            .MoveNEXT
        Loop
        .Close
        .Open "SELECT UserName,Password FROM Users" & strCondition, con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            strOriginal = !Password & ""
            strEncrypted = ToHexDump(CryptRC4(strOriginal, "AwmWrangler"))
            con.Execute "UPDATE Users SET Password='" & strEncrypted & "',PasswordEncrypted=1 WHERE UserName='" & !UserName & "'"
            .MoveNEXT
        Loop
        .Close
        con.Execute "UPDATE GeneralData SET DataValue='True' WHERE DataName='Encryption'"
    End With
    Set rs = Nothing

End Sub
