Attribute VB_Name = "mod_FGrid"
Public Const GRID_BK_COLOR_ODD As Long = &HD7DDD5       '&HF8F3FE       '&HE3ECF4         ' &HE9F9FA   '&HE3F4E3
Public Const GRID_BK_COLOR_EVEN As Long = vbWhite '&HE3F4E3

Public Sub SetBackColor(FGrid As MSFlexGrid, RowNo As Integer)

    Dim lColor As Long
    
    If RowNo Mod 2 = 0 Then
        lColor = GRID_BK_COLOR_EVEN
    Else
        lColor = GRID_BK_COLOR_ODD
    End If
    
    With FGrid
        .Row = RowNo
        For i = 0 To .Cols - 1
            .Col = i
            .CellBackColor = lColor
        Next i
    End With
    
End Sub

Public Sub ClearRows(FGrid As MSFlexGrid)
    FGrid.Rows = 2
    
    For i = 0 To FGrid.Cols - 1
        FGrid.TextMatrix(1, i) = ""
    Next i
    
End Sub
