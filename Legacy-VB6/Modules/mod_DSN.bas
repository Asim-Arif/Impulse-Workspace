Attribute VB_Name = "mod_DSN"
'Constant Declaration
Private Const ODBC_ADD_DSN = 1        ' Add data source
Private Const ODBC_CONFIG_DSN = 2     ' Configure (edit) data source
Private Const ODBC_REMOVE_DSN = 3     ' Remove data source
Private Const vbAPINull As Long = 0& 'amp;  ' NULL Pointer

'Function Declare

Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hwndParent As Long, ByVal fRequest As Long, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long

Public Function CreateDSN(Optional strSvrName) As Boolean
    
    Dim intRet As Long
    
    Dim strDriver As String
    Dim strAttributes As String
    If IsMissing(strSvrName) Then
        strSvrName = "SERVER"
    End If
    
    'Set the driver to SQL Server because it is most common.
    strDriver = "SQL Server"
    'Set the attributes delimited by null.
    'See driver documentation for a complete
    'list of supported attributes.
    strAttributes = "SERVER=" & strSvrName & Chr$(0)
    strAttributes = strAttributes & "DESCRIPTION=Dynamically Created For AWM" & Chr$(0)
    If strCompany = "PAKSMITH GROUP" Then
        strAttributes = strAttributes & "DSN=DSN_PakSmith_AWM" & Chr$(0)
        strAttributes = strAttributes & "DATABASE=PakSmith_AWM" & Chr$(0)
    ElseIf strCompany = "Weldon Instruments" Then
        strAttributes = strAttributes & "DSN=DSN_WeldonInstr_AWM" & Chr$(0)
        strAttributes = strAttributes & "DATABASE=WeldonInstr_AWM" & Chr$(0)
    ElseIf strCompany = "Weldon Industries" Then
        strAttributes = strAttributes & "DSN=DSN_WeldonInd_AWM" & Chr$(0)
        strAttributes = strAttributes & "DATABASE=WeldonInd_AWM" & Chr$(0)
    'ElseIf strCompany = "All Seas" Then
    Else
        strAttributes = strAttributes & "DSN=DSN_SMBI_AWM" & Chr$(0)
        strAttributes = strAttributes & "DATABASE=SMBI_AWM" & Chr$(0)
    End If
   
    strAttributes = strAttributes & "Trusted_Connection=yes" & Chr$(0)
    'To show dialog, use Form1.Hwnd instead of vbAPINull.
    intRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_DSN, _
    strDriver, strAttributes)
    
    If intRet Then
        CreateDSN = True
    Else
        CreateDSN = False
    End If
End Function


Public Function ConfigDSN(Optional strSvrName) As Boolean
    
    Dim intRet As Long
    
    Dim strDriver As String
    Dim strAttributes As String
    If IsMissing(strSvrName) Then
        strSvrName = "SERVER"
    End If
    
    'Set the driver to SQL Server because it is most common.
    strDriver = "SQL Server"
    'Set the attributes delimited by null.
    'See driver documentation for a complete
    'list of supported attributes.
    strAttributes = "SERVER=" & strSvrName & Chr$(0)
    strAttributes = strAttributes & "DESCRIPTION=Dynamically Created For Star" & Chr$(0)
    If strCompany = "PAKSMITH GROUP" Then
        strAttributes = strAttributes & "DSN=DSN_PakSmith_AWM" & Chr$(0)
        strAttributes = strAttributes & "DATABASE=PakSmith_AWM" & Chr$(0)
    ElseIf strCompany = "Weldon Instruments" Then
        strAttributes = strAttributes & "DSN=DSN_WeldonInstr_AWM" & Chr$(0)
        strAttributes = strAttributes & "DATABASE=WeldonInstr_AWM" & Chr$(0)
    ElseIf strCompany = "Weldon Industries" Then
        strAttributes = strAttributes & "DSN=DSN_WeldonInd_AWM" & Chr$(0)
        strAttributes = strAttributes & "DATABASE=WeldonInd_AWM" & Chr$(0)
    'ElseIf strCompany = "All Seas" Then
    Else
        strAttributes = strAttributes & "DSN=DSN_SMBI_AWM" & Chr$(0)
        strAttributes = strAttributes & "DATABASE=SMBI_AWM" & Chr$(0)
    End If
    strAttributes = strAttributes & "Trusted_Connection=yes" & Chr$(0)
    'To show dialog, use Form1.Hwnd instead of vbAPINull.
    intRet = SQLConfigDataSource(vbAPINull, ODBC_CONFIG_DSN, strDriver, strAttributes)
      
    If intRet Then
        ConfigDSN = True
    Else
        ConfigDSN = False
    End If
    
End Function
