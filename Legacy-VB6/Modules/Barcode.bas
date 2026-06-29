Attribute VB_Name = "Barcode"
Public Sub PrintBarcode(Cond As String, Optional Pages As Integer, Optional Copies As Integer, Optional FromPageNo As Integer)

'Dim y(8) As Double
'Pic.Cls
Dim Rec As New ADODB.Recordset
With Rec
    .Open "SELECT EmpID FROM Employees " & Cond, con, adOpenStatic, adLockReadOnly
    
    
     
    
    
    
    
    If .RecordCount = 0 Then Exit Sub
    For k = 1 To Copies
        .MoveFirst
        For i = 1 To Pages
            Printer.CurrentY = 450
            For j = 1 To 4 '.RecordCount / 2
            
                If i < FromPageNo Then GoTo Skip_The_Record
            
                Printer.CurrentX = 3650
                
                Printer.CurrentY = Printer.CurrentY + 1900
                Printer.Print ![EmpId] & ""
                Printer.CurrentY = Printer.CurrentY + 75
                Call DrawBarCode(!EmpId & "", 900, Printer.CurrentY)
                Printer.CurrentY = Printer.CurrentY + 125
                Printer.Print ""
                
Skip_The_Record:
                .MoveNext
                
                If .EOF Then Exit For

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            Next
        Printer.EndDoc

        Next i
    Next k
End With

End Sub

Private Sub DrawBarCode(EmpId As String, x As Double, y As Double)

    Dim strCode As String
    
    Dim barscale As Integer, barHeight As Single, StartX As Single, startY As Single
    
    'Dim pic As PictureBox
    strCode = EmpId
    
    'Pic.Cls
'    Set pic = PicBarCode
    'Barcode Text1, PicBarCode, 2, 60, 20, 20
    barscale = 2 * 10
    barHeight = 30 * 15
    StartX = x
    startY = y
        
    Printer.CurrentX = 1800
    
    
    
    
    Dim varBar1, varBar2
    varBar1 = Split("0,1,2,3,4,5,6,7,8,9,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,-,., ,$,/,+,%,*", ",")
    varBar2 = Split("111221211,211211112,112211112,212211111,111221112,211221111,112221111,111211212,211211211,112211211,211112112,112112112,212112111,111122112,211122111,112122111,111112212,211112211,112112211,111122211,211111122,112111122,212111121,111121122,211121121,112121121,111111222,211111221,112111221,111121221,221111112,122111112,222111111,121121112,221121111,122121111,121111212,221111211,122111211,121212111,121211121,121112121,111212121,121121211", ",")
    sBar = "121121211" & "1"
    'sBar = ""
    For i0 = 1 To Len(strCode)
        For i1 = 0 To UBound(varBar1)
            If Mid(strCode, i0, 1) = varBar1(i1) Then sBar = sBar & varBar2(i1) & "1"
        Next
    Next
    sBar = sBar & "121121211"
    
    
    Dim barWidth As Single, barStart As Single
    
    'Case "39": strCode = UCase(strCode): Code39 strCode
    
    barStart = StartX + 100
    For i0 = 1 To Len(sBar)
'        barWidth = Mid(sBar, i0, 1) * barscale
'
'        If i0 Mod 2 > 0 Then
'            pic.Line (barStart, startY)-Step(barWidth, barHeight), vbBlack, BF
'        End If
'        barStart = barStart + IIf(i0 Mod 2 > 0, barWidth, barWidth * 1.3)
        barWidth = Mid(sBar, i0, 1) * barscale
        
        If i0 Mod 2 > 0 Then
            'pic.Line (barStart, startY)-Step(barWidth, barHeight), vbBlack, BF
            Printer.Line (barStart, startY)-Step(barWidth, barHeight), vbBlack, BF
        End If
        
        barStart = barStart + IIf(i0 Mod 2 > 0, barWidth * 1.1, barWidth * 1.3 * 1.1)
        
    Next
    'Printer.EndDoc
'    Pic.DrawWidth = 1
'    Pic.FontSize = 16
'    Pic.CurrentX = StartX
'    Pic.CurrentY = (startY * 1.2) + barHeight
'    Pic.Print strCode
    
    
End Sub



