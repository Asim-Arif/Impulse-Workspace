Attribute VB_Name = "Salaries"
Public Function GetShortAmt(EmpId As String, TDate As Date) As Double

    Dim rsLedger As New ADODB.Recordset
    Dim Cond As String
    Dim SQL As String
    Dim cmd As New ADODB.Command
    Dim rsCheck As New ADODB.Recordset
    
    
    'Cond = " Where empid='" & EmpID & "' and"
    'TDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), Day(Date))
    
    'Sql = "select * from VAdvanceSalary " & Cond & " JoinDate<='" & Format(TDate, "dd-MMM-yyyy") & "' "
    'cmd.CommandText = Sql
    
    
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "VAdvanceSalary"
    
    Set cmd.ActiveConnection = con
    
    cmd.Parameters.Refresh
    
    
    
    cmd.Parameters(1).Type = adDBTimeStamp
    cmd.Parameters(1).Value = TDate
    
    cmd.Parameters(2).Type = adVarChar
    cmd.Parameters(2).Value = EmpId
    
    'cmd.Parameters(4).Type = adDBTimeStamp
    'cmd.Parameters(4).value = EmpID
    
    rsLedger.CursorType = adOpenForwardOnly
    rsLedger.LockType = adLockReadOnly
    Set rsLedger = cmd.Execute
    With rsLedger
        
        '.Open Sql, Con, adOpenForwardOnly, adLockReadOnly
        Dim TotalSuns As Integer
        Dim SalDays As Double
        Dim Rate As Double
        Dim AbsAllowed As Integer
        Dim AbsentDays As Double
        Dim TotalSalDays As Double
        Dim TempTotalSaldays As Double
        
        Dim total As Double
        Dim OverTimeHours As Double
        Dim LateHours As Double
        Dim TaxPer As Double
        Dim ShortTermPer As Double
        Dim DedAmt As Double
        Dim Payed As Double
        Dim TaxSal As Double
        Dim IsComputed As Boolean
        Dim LastDay As Integer
        Dim NTime As Integer
        Dim LTime As Double
        Dim DT As String
        Dim IsPayed As Boolean
        Dim Holidays As Integer
        Dim rs As New ADODB.Recordset
        
        Dim SundayAbsents As Double
        
        TotalSalDays = DateAdd("m", 1, TDate) - TDate
        
        
        LastDay = TotalSalDays
        
        If TDate = ServerDate Then
            TotalSalDays = Day(ServerDate)
        ElseIf ServerDate < TDate Then
            MsgBox "Can't Show Entries Of Future Months.", vbInformation
            Exit Function
        End If
        
        TempTotalSaldays = TotalSalDays
        
        'Check that Employee Discontinued In This Month
                    
               If Not IsNull(![DiscontinuedOn]) Then
                   If Month(![DiscontinuedOn]) = Month(TDate) And ![DiscontinuedOn] <= TDate Then
                     TotalSalDays = Day(![DiscontinuedOn])
                   End If
               Else
                  TotalSalDays = TempTotalSaldays
               End If
                    
                    
           
           'if Joined This Month Calculate From Join Date
           
           If Format(![JoinDate], "MMM-yyyy") = Format(TDate, "MMM-yyyy") Then
                rs.Open "Select count(DT) from Holidays Where DT Between '" & ![JoinDate] & "' and '" & JDate & "'", con, adOpenForwardOnly, adLockReadOnly
                  If rs.RecordCount > 0 Then
                     Holidays = rs.Fields(0) & ""
                  End If
                rs.Close
               Set rs = Nothing
               
               SundayAbsents = getSundayAbsents(![EmpId] & "", Month(TDate), year(TDate), ![WeeklyAbs], Int(TotalSalDays))
               
               TotalSuns = TotalSundays(Month(TDate), year(TDate), TotalSalDays, Day(![JoinDate]))
               
               Holidays = TotalSuns + Holidays - SundayAbsents
            Else
                 rs.Open "Select count(DT) from Holidays Where month(dt)=" & Month(TDate) & " and Year(DT)=" & year(TDate) & " and day(dt)<" & TotalSalDays, con, adOpenForwardOnly, adLockReadOnly
                  If rs.RecordCount > 0 Then
                     Holidays = rs.Fields(0) & ""
                  End If
                 rs.Close
               Set rs = Nothing
               
               
               SundayAbsents = getSundayAbsents(![EmpId], Month(TDate), year(TDate), Int(TotalSalDays), ![WeeklyAbs])
               
               TotalSuns = TotalSundays(Month(TDate), year(TDate), TotalSalDays)
                                              
               Holidays = TotalSuns + Holidays - SundayAbsents
               
               
               
            End If
                    
                    
                    
                    
                    
                    
        'LV.ListItems.Clear
        If .RecordCount > 0 Then
                    
            Call Computed(![EmpId], IsComputed, TaxPer, AbsAllowed, OverTimeHours, NTime, LTime, DT, IsPayed, ShortTermPer, TDate)
            
            If Not IsComputed Then
                'Tax %
                If IsNull(![PerDeduct]) Then TaxPer = 0 Else TaxPer = Val(![PerDeduct]) & ""
                'ShortTerm %
                If IsNull(![ShortTermPer]) Then ShortTermPer = 0 Else ShortTermPer = Val(![ShortTermPer]) & ""
                'Absent Days....
                AbsAllowed = Val(![AbsAllowed] & "")
                'OverTime Hours
                If IsNull(![OverTimeHours]) Then OverTimeHours = 0 Else OverTimeHours = ![OverTimeHours]
                'Late Hours
                NTime = Val(![NTimes] & "")
                LTime = Val(![TimeAllowed] & "")
                
            End If
            
            ShortTermPer = ShortTermPer / 100
            
            Dim theKey As String
            If IsComputed Then
                theKey = ![EmpId] & DT
            Else
                theKey = ![EmpId] & ""
            End If
            
            Rate = Round(Val(![StartingSalary]) / LastDay)
            
            
            If IsNull(![presentDays]) Then
                SalDays = 0
            Else
                SalDays = ((Val(![presentDays]) / 2) + Holidays) + AbsAllowed
                
                'if Salary Days Exceed Total Salary Days Set Them To Salary Days
                If SalDays > TotalSalDays Then SalDays = TotalSalDays
            End If
            
            AbsentDays = TotalSalDays - SalDays
            If AbsentDays < 0 Then AbsentDays = 0
            'If IsNull(![AbsentDays]) Then AbsentDays = 0 Else AbsentDays = Val(![AbsentDays]) / 2
            'AbsentDays = IIf(AbsentDays > 0, AbsentDays - AbsAllowed, AbsentDays)
            'itm.SubItems(4) = AbsentDays
            'itm.SubItems(5) = Round(AbsentDays * Rate)
            AbsentDays = Round(AbsentDays * Rate)
            
            
            ''If SalDays = 0 Then SalDays = 0 Else SalDays = (SalDays - AbsentDays)
            'itm.SubItems(6) = SalDays
            'itm.SubItems(7) = Round(SalDays * Rate)
            SalDays = Round(SalDays * Rate)
            
            
            'itm.SubItems(8) = OverTimeHours
            'itm.SubItems(9) = (OverTimeHours * (Rate * Val(![OverTime])))
            OverTimeHours = OverTimeHours * (Rate / Val(![OverTime] & ""))
            
            LateHours = GetLateHours(![EmpId], NTime, LTime, TDate)
            
            'itm.SubItems(10) = LateHours
            'itm.SubItems(11) = Round((LateHours * (Rate / 8)))
            LateHours = Round((LateHours * (Rate / 8)))
            
            total = Round(SalDays + OverTimeHours - LateHours)
            'itm.SubItems(12) = Total
            'TaxPer = Val(![PerDeduct])
            'itm.SubItems(13) = TaxPer
            'DedAmt = Round((Val(![Salary]) / TotalSalDays) * (TaxPer / 100))
            TaxSal = (Val(![StartingSalary]) / LastDay) * TotalSalDays
            DedAmt = Round(TaxSal * (TaxPer / 100))
            'itm.SubItems(14) = DedAmt
            'itm.SubItems(15) = Total - DedAmt
            total = total - DedAmt
            
            If IsNull(![Payed]) Then Payed = 0 Else Payed = ![Payed]
            'itm.SubItems(16) = Payed
            'itm.SubItems(17) = Val(itm.SubItems(15)) - Payed
            total = total - Payed - Val(![AdvSalpaid] & "")
            
            total = total * ShortTermPer
            'LastDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), LastDay)
            'Save Tax,Settings for Selected Month...
            'If Not (IsComputed) And (ServerDate > LastDate) Then Call SaveIt(rsLedger, TaxPer)
            
            '.MoveNext
        End If
        .Close
    End With
    Set rsLedger = Nothing
    SQL = ""
    Cond = ""
    ICO = ""
    GetShortAmt = total
End Function

Public Function TotalSundays(iMonth As Integer, iYear As Integer, Optional iDay, Optional StartDay As Integer = 1) As Integer
    Dim FirstDay As Integer
    Dim FirstSunday As Integer
    Dim TotSuns As Integer
    Dim DT As Date
    DT = DateSerial(iYear, iMonth, StartDay)
    If IsMissing(iDay) Then
        iDay = DateAdd("m", 1, DT) - DT
    End If
    FirstDay = Weekday(DT)
    If StartDay <> FirstDay Then
        FirstSunday = StartDay + (8 - FirstDay)
    Else
        FirstSunday = StartDay
    End If
    'TotSuns = (Int((iDay - FirstSunday) / 7))
    TotSuns = 0
    For i = FirstSunday To iDay Step 7
        TotSuns = TotSuns + 1
    Next
    TotalSundays = TotSuns
End Function

Public Function SundayDates(iMonth As Integer, iYear As Integer, Optional iDay, Optional StartDay As Integer = 1) As Date()
    Dim FirstDay As Integer
    Dim FirstSunday As Integer
    Dim TotSuns As Integer
    Dim SunDates() As Date
    Dim DT As Date
    DT = DateSerial(iYear, iMonth, StartDay)
    If IsMissing(iDay) Then
        iDay = DateAdd("m", 1, DT) - DT
    End If
    FirstDay = Weekday(DT, vbSunday)
    'FirstSunday = StartDay + (8 - FirstDay)
    
    
    If StartDay <> FirstDay Then
        FirstSunday = StartDay + (8 - FirstDay)
    Else
        FirstSunday = StartDay
    End If
    'TotSuns = (Int((iDay - FirstSunday) / 7))
    TotSuns = 0
    For i = FirstSunday To iDay Step 7
        TotSuns = TotSuns + 1
    Next
    
    
    
    'TotSuns = (Int((iDay - FirstSunday) / 7)) + 1
    'TotalSundays = TotSuns
    DT = DateSerial(iYear, iMonth, FirstSunday)
    
    ReDim SunDates(TotSuns + 1)
    SunDates(0) = DateSerial(iYear, iMonth, StartDay)
    For i = 1 To TotSuns
        SunDates(i) = DT + ((i - 1) * 7)
    Next
    SunDates(i) = DateSerial(iYear, iMonth, iDay)
    SundayDates = SunDates
    
End Function

Private Function GetLateHours(EmpId As String, NTime As Integer, LTime As Double, TDate As Date) As Double

    Dim LHours As Double
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
    With cmd
        '.CommandText = "Select LHrs from VLHours Where EmpID='" & EmpID & "'"
        .CommandType = adCmdStoredProc
        .CommandText = "VLHours"
        Set .ActiveConnection = con
        .Parameters.Refresh
        
        
        .Parameters(1).Type = adNumeric
        .Parameters(1).Value = LTime
        
        
        .Parameters(2).Type = adDBTimeStamp
        .Parameters(2).Value = TDate
        
        .Parameters(3).Type = adNumeric
        .Parameters(3).Value = NTime
        
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        Set rs = .Execute
        
    End With
    
    If rs.RecordCount > 0 Then
        If IsNull(rs.Fields(1)) Then LHours = 0 Else LHours = rs.Fields(1)
    Else
        LHours = 0
    End If
    GetLateHours = LHours
    
End Function
Private Sub Computed(Emp As String, ByRef IsComputed As Boolean, ByRef TaxPer As Double, ByRef AbsentDays As Integer, ByRef OverTimeHours As Double, ByRef NTime As Integer, ByRef LTime As Double, ByRef DT As String, ByRef Payed As Boolean, ShortTermPer, ToDate As Date)

    Dim rsCheck As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    With cmd
        '.CommandText = "Select * from VSalarySettings where empid='" & Emp & "'"
        
        .CommandType = adCmdStoredProc
        .CommandText = "VSalarySettings"
        
        Set .ActiveConnection = con
        .Parameters.Refresh
        
        .Parameters(1).Type = adDBTimeStamp
        .Parameters(1).Value = ToDate
        
        
        
        .Parameters(2).Type = adVarChar
        .Parameters(2).Value = Emp
        
        .Parameters(3).Type = adBoolean
        .Parameters(3).Value = 0
        
        rsCheck.CursorType = adOpenForwardOnly
        rsCheck.LockType = adLockReadOnly
        Set rsCheck = .Execute()
    End With
    
    With rsCheck
        
        If .RecordCount > 0 Then
            IsComputed = True
            TaxPer = ![PerDeduct]
            AbsentDays = ![absentallowed]
            OverTimeHours = ![OverTime]
            NTime = ![NTimes]
            LTime = ![TimeAllowed]
            DT = Format(![DT], "dd-MMM-yyyy") & ""
            ', AbsentDays, OverTimeHours, NTime, LTime)
            'LV.Checkboxes = True
            'cmdUpdate.Visible = True
            Payed = ![Payed]
            ShortTermPer = ![ShortTermPer]
        Else
            IsComputed = False
            'LV.Checkboxes = False
            'cmdUpdate.Visible = False
        End If
    End With
    
End Sub


Public Function getSundayAbsents(EmpId As String, iMonth As Integer, iYear As Integer, iDay As Integer, DedAbs As Integer, Optional StartDay As Integer = 1) As Integer
Dim SunArr() As Date
Dim SunAbsents As Integer
Dim SUNABS As Double
Dim WeeklyPresents As Double

If DedAbs = 0 Then DedAbs = 7

SunAbsents = 0

SunArr = SundayDates(iMonth, iYear, iDay, StartDay)

For i = LBound(SunArr) To UBound(SunArr) - 1

    'Get Teh Absents From One Monday to Next Monday
    

    
    WeeklyPresents = con.Execute("Select Count(ATtendance)/2  from AttendanceSheet Where Attendance=1 and Dt Between '" & _
                 Format(SunArr(i), "dd-MMM-yyyy") & "' and '" & Format(SunArr(i + 1), "dd-MMM-yyyy") & "' and EmpID='" & EmpId & "'").Fields(0)
                 
                 'Calculate Weekly  Absents From 1st Holiday To 2nd Holiday
                 
                 If LCase(Format(SunArr(i + 1), "ddd")) = "sun" Then
                    SUNABS = (SunArr(i + 1) - SunArr(i)) - WeeklyPresents - 1
                 Else
                    SUNABS = (SunArr(i + 1) - SunArr(i)) - WeeklyPresents
                 End If
                 
    If SUNABS >= DedAbs Then
        SunAbsents = SunAbsents + 1
        With con.Execute("Select count(*) From Holidays Where  Dt Between '" & Format(SunArr(i), "dd-MMM-yyyy") & "' and '" & Format(SunArr(i + 1), "dd-MMM-yyyy") & "'")
            If .Fields(0) > 0 Then
                SunAbsents = SunAbsents + 1
            End If
        
        End With
    End If
    
Next

getSundayAbsents = SunAbsents

'If SunAbsents > 0 Then Debug.Print EmpId, SunAbsents

End Function

