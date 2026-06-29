Attribute VB_Name = "mod_General"
Public Sub OnlyNums(ByRef KeyAscii As Integer, Optional bPoint As Boolean = True)

    If Not (IsNumeric(Chr$(KeyAscii)) Or KeyAscii = 8) Then
        If Not (bPoint And Chr$(KeyAscii) = ".") Then
            
            KeyAscii = 0
        End If
    End If
    
End Sub

Public Function OnlyNums1(KeyAscii As Integer, Optional bPoint As Boolean = True) As Integer

    If Not (IsNumeric(Chr$(KeyAscii))) Then
        If Not (bPoint And Chr$(KeyAscii) = ".") Then
            KeyAscii = 0
        End If
    End If
    OnlyNums1 = KeyAscii
    
End Function

Public Sub AddDecForProforma(strProforma As String, FGrid As MSFlexGrid)

    On Error GoTo err
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "VPDeclarations"
        .ActiveConnection = con
        
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        Set rs = .Execute(, strProforma)
    End With
    
    Dim iRow As Integer
    
    With FGrid
        Do Until rs.EOF
            iRow = rs.AbsolutePosition
            If iRow > 1 Then .AddItem ""
            .TextMatrix(iRow, 0) = rs.Fields(1)
            '.TextMatrix(iRow, 1) = rs.Fields(1)
            rs.MoveNext
            Call SetBackColor(FGrid, iRow)
        Loop
        rs.Close
        'If .Rows = 2 And .TextMatrix(1, 0) = "" Then .AddItem
        
    End With
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
End Sub

Public Sub AddDecForCustom(strProforma As String, FGrid As MSFlexGrid, FGrid1 As MSFlexGrid)

    On Error GoTo err
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "VCustomDeclarations"
        .ActiveConnection = con
        
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        Set rs = .Execute(, strProforma)
    End With
    
    Dim iRow As Integer
    
    With FGrid
        Do Until rs.EOF
            iRow = rs.AbsolutePosition
            If iRow > 1 Then .AddItem ""
            
            .TextMatrix(iRow, 0) = rs.Fields(1) & ""
            Call SetBackColor(FGrid, iRow)
            rs.MoveNext
        Loop
        rs.Close
    End With
    
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "VCustomSROText"
        .ActiveConnection = con
        
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        Set rs = .Execute(, strProforma)
    End With
    
    
    
    With FGrid1
        Do Until rs.EOF
            iRow = rs.AbsolutePosition
            If iRow > 1 Then .AddItem ""
            
            .TextMatrix(iRow, 0) = rs.Fields(1) & ""
            Call SetBackColor(FGrid1, iRow)
            rs.MoveNext
        Loop
        rs.Close
    End With
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
End Sub

Public Sub ResetLV(LV As ListView)
    For i = 1 To LV.ColumnHeaders.Count
        LV.ColumnHeaders(i).Width = LV.ColumnHeaders(i).Tag
    Next i
End Sub
Public Sub SaveLV(LV As ListView)
    For i = 1 To LV.ColumnHeaders.Count
        LV.ColumnHeaders(i).Tag = LV.ColumnHeaders(i).Width
    Next i
End Sub
Public Function GetColNo(LV As ListView, x As Single) As Integer

    Dim iCol As Integer
    iCol = 0
    For i = 1 To LV.ColumnHeaders.Count
        If x > LV.ColumnHeaders(i).Left And x < (LV.ColumnHeaders(i).Width + LV.ColumnHeaders(i).Left) Then
            iCol = i
            Exit For
        End If
    Next i
    
    GetColNo = iCol
    
End Function

Public Function IsLoaded(Str As String) As Boolean

    Dim f As Form
    IsLoaded = False
    For Each f In Forms
        If UCase(f.Name) = UCase(Str) Then
            IsLoaded = True
            Exit For
        End If
    Next
    
End Function

Public Function rsExecute(StrSQL As String) As Variant()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim Arr() As Variant
    
    With rs
        .Open StrSQL, con, adOpenForwardOnly, adLockReadOnly
        ReDim Arr(.Fields.Count - 1)
        If Not .EOF Then
            For i = 0 To .Fields.Count - 1
                Arr(i) = .Fields(i).value & ""
            Next i
        End If
        .Close
    End With
    
    Set rs = Nothing
    
    rsExecute = Arr
    
    Exit Function
err:
    MsgBox err.Description
End Function

Public Sub PrintReportDirectly(strSel As String, ArgSourceRpt As CRAXDDRT.Report, Optional strQueryString As String, Optional bDiscardSaved As Boolean = True)
    
    Dim SrcReport As CRAXDDRT.Report
    Set SrcReport = ArgSourceRpt
    
'    On Error Resume Next
'    SrcReport.Database.Tables(1).ConnectBufferString = "DSN=DSN_Vision;;User ID=sa;;Database=vision;;Password=visionasim"
'    On Error GoTo 0
    strSel = strSel & " "
    Screen.MousePointer = vbHourglass
    
    If bDiscardSaved Then
        SrcReport.DiscardSavedData
    End If
    If strSel <> "" Then SrcReport.RecordSelectionFormula = Left(strSel, Len(strSel) - 1)
    SrcReport.RecordSelectionFormula = strSel '& "'"
    If strQueryString <> "" Then
        SrcReport.SQLQueryString = strQueryString
    End If
    SrcReport.PrintOut False
    
    
End Sub

