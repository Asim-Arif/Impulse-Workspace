Attribute VB_Name = "PayrollFunctions"
 Public Function IsWorkingDay(DT As Date) As Boolean

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select DT From Holidays Where DT='" & DT & "'", con, adOpenForwardOnly, adLockReadOnly
        IsWorkingDay = .EOF
        .Close
    End With
    Set rs = Nothing
    
    Exit Function
err:
    MsgBox err.Description
    
End Function

Public Function TimeToMin(RTime As String) As Integer
    
    If RTime = "" Then RTime = "00:00"
    TimeToMin = Hour(RTime) * 60 + Minute(RTime)

End Function

Public Function MinsToTime(Minutes As Long) As String

    MinsToTime = Format(Int((Minutes / 60)), "00") & ":" & Format((Minutes Mod 60), "00")

End Function

Public Function addMinToTime(Min As Integer, RTime As String) As String

    Dim THrs As Integer
    Dim Tmin As Integer
    THrs = Int(Min / 60) + Hour(RTime)
    Tmin = Min Mod 60 + Minute(RTime)
    
    If Tmin > 60 Then
        THrs = THrs + Int(Tmin / 60)
        Tmin = Tmin Mod 60
    End If
    
    If THrs = 0 And Tmin = 0 Then
        addMinToTime = "- -"
    Else
        addMinToTime = Format(THrs, "00") & ":" & Format(Tmin, "00")
    End If

End Function

Public Function getHrsMin(ByVal Hrs As Double, Optional ShowDashInsteadOfZero As Boolean) As String

    Dim THrs As Integer
    Dim Tmin As Integer
    
    Hrs = Hrs * 60
    THrs = Int(Hrs / 60)
    If Round(Hrs - (THrs * 60)) = 60 Then
        THrs = THrs + 1
    Else
        Tmin = Hrs Mod 60
    End If
    
    If THrs = 0 And Tmin = 0 Then
        If ShowDashInsteadOfZero Then
            getHrsMin = "-"
        Else
            getHrsMin = "0"
        End If
    Else
        getHrsMin = Format(THrs, "00") & ":" & Format(Tmin, "00")
    End If
End Function




Public Function CalcSpecialAbs(Employee As String, DTDate As Date, _
    BGHoliday As Boolean, AGHoliday As Boolean, BPHoliday As Boolean, APHoliday As Boolean) As Integer

Dim Absents As Integer
Dim DTDay As DayConstants

Dim Rec As New ADODB.Recordset, RecHol As New ADODB.Recordset

Dim DT As Date, TotalCombHol As Integer, LastHol As Date, OnceCalc As Boolean
Dim PrevAbsent As Date

With Rec


    .CursorLocation = adUseClient
    
    
    '.Open "SELECT DT,Count(Attendance) as TotalAtt FROM AttendanceSheet Where EmpID='" & Employee & "' AND MONTH(DT)=" & DTMonth & "" & _
          " AND Year(DT)=" & DTYear & " AND Attendance=false  group by DT having  count(Attendance)=2", con, adOpenForwardOnly, adLockReadOnly
          
          
          
    .Open "SELECT DT  FROM AttendanceSheet Where EmpID='" & Employee & "' AND month(dt) = Month('" & DTDate & "') AND year(DT) = Year('" & DTDate & "') AND DT <='" & DTDate & "'" & _
          " AND Attendance=0 AND DATENAME(WEEKDAY,DT)<>'Sunday' ORDER BY DT", con, adOpenForwardOnly, adLockReadOnly
          
          
          
    Absents = 0
    If .RecordCount > 0 Then
    
        Do While .EOF = False
            
            
            If Format(!DT, "ddd") = "Sat" And BGHoliday Then
                Absents = Absents + 1
                DT = !DT
            ElseIf Format(!DT, "ddd") = "Mon" And DT <> !DT - 2 And AGHoliday Then
                Absents = Absents + 1
            End If
            '----------------------------
            RecHol.CursorLocation = adUseClient
            RecHol.Open "SELECT * FROM Holidays WHERE MONTH(DT)=" & Month(DTDate) & " ORDER BY DT", con, adOpenKeyset, adLockReadOnly
            
            'For Absent Before Any Public Holiday/Holidays
            If RecHol.RecordCount > 0 Then
            
                Do While RecHol.EOF = False
                
          
                    If LastHol + 1 = CDate(RecHol!DT) Then
                        TotalCombHol = TotalCombHol + 1
                    Else
                        TotalCombHol = 1
                    End If
          
          
                    If PrevHol = CDate(RecHol!DT) - 1 Then
                        Absents = Absents + 1
                        PrevHol = RecHol!DT
                        'PrevAbsent = CDate(!dt)
                    End If
                              
                    If Day(!DT) + 1 = Day(RecHol!DT) And Not (CDate(!DT) = PrevHol + 1) And (CDate(!DT) - TotalCombHol) <> PrevAbsent And BPHoliday Then
                        PrevAbsent = CDate(!DT)
                        'OnceCalc = True
                        Absents = Absents + 1
                        PrevHol = RecHol!DT
                    End If
                    

                    
                    If Day(!DT) - 1 = Day(RecHol!DT) And (CDate(!DT) - (TotalCombHol + 1)) <> PrevAbsent And APHoliday Then
                        PrevAbsent = CDate(!DT)
                        Absents = Absents + TotalCombHol
                        OnceCalc = True
                    End If
                    
                    LastHol = CDate(RecHol!DT)
                
                    RecHol.MoveNEXT
                Loop
            End If
            
            '--------------------------
            .MoveNEXT
            Set RecHol = Nothing
            'OnceCalc = False
        Loop
        'OnceCalc = False

    End If
    
End With

CalcSpecialAbs = Absents

End Function

Public Sub ShowShortTermForm(bContractor As Boolean, Optional bBPV As Boolean = False)
    Dim f As Form
    If bBPV Then
        Set f = New NewAdvanceShortBPV
    Else
        Set f = New NewAdvanceShort
    End If
    Load f
    Call f.ShowMe(bContractor)
End Sub

Public Sub ShowLongTermForm(bContractor As Boolean, Optional bBPV As Boolean = False)
    Dim f As Form
    If bBPV Then
        Set f = New NewAdvanceLongBPV
    Else
        Set f = New NewAdvanceLong
    End If
    Load f
    Call f.ShowMe(bContractor)
End Sub

Public Function IsEmpPresent(strEmpID As String, DT As Date) As Boolean

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim bPresent As Boolean
    With rs
        .Open "Select Attendance From AttendanceSheet Where EmpID='" & strEmpID & "' AND DT='" & DT & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            bPresent = False
        Else
            If !Attendance = 1 Then
                bPresent = True
            Else
                bPresent = False
            End If
        End If
        .Close
    End With
    Set rs = Nothing
    IsEmpPresent = bPresent
    Exit Function
err:
    MsgBox err.Description
End Function


Public Function IsProbation(strEmpID As String, DtTo As Date) As Boolean

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim DTJoin As Date
    Dim iMonths As Integer, iDays As Integer
    Dim iProbationPeriod As Integer
    With rs
        .Open "SELECT ProbationPeriod,JoinDate FROM Employees WHERE EmpID='" & strEmpID & "'", con, adOpenForwardOnly, adLockReadOnly
        DTJoin = Format(!JoinDate, "dd-MMM-yyyy")
        iProbationPeriod = Val(!ProbationPeriod & "")
        .Close
    End With
    
    iDays = DateDiff("d", DTJoin, DtTo)
    'Call DateDiffInMonthsNDays(dtJoin, DtTo, iMonths, iDays)
'    If iMonths >= 3 Then
'        IsProbation = False
'    Else
'        IsProbation = True
'    End If
    If iProbationPeriod > iDays Then
        IsProbation = True
    Else
        IsProbation = False
    End If
    Set rs = Nothing
    
    Exit Function
err:
    MsgBox err.Description
End Function
Public Function GetRate(strEmpID As String, iMonth As Integer, iYear As Integer, Optional iTotalMonthDays As Integer = 0, Optional iTotalSundays As Integer = 0, Optional iHDays As Integer, Optional iWorkingHrs As Integer, Optional iSalary As Integer = 0)

    On Error GoTo err
    Dim iTotal As Integer, iTotalWorkingHrs As Integer
    Dim MonthDays As Integer
    Dim dRate As Double
    
    If iWorkingHrs = 0 Then
        If iTotalMonthDays = 0 Then
            iTotal = TotalMonthDays(iMonth, iYear)
        Else
            iTotal = iTotalMonthDays
        End If
         
        If iTotalSundays = 0 Then
            iTotalSundays = TotalSundays(iMonth, iYear, iTotal)
        End If
        
        MonthDays = iTotal - iTotalSundays
          
        If iHDays = 0 Then
            iHDays = Val(con.Execute("Select count(DT) from Holidays Where month(DT)=" & _
            iMonth & " and Year(DT)=" & iYear).Fields(0))
        End If
         
        MonthDays = MonthDays - iHDays
        iTotalWorkingHrs = MonthDays * 8
    Else
        iTotalWorkingHrs = iWorkingHrs
    End If
    
    If iSalary = 0 Then
        Dim rs As New ADODB.Recordset
        With rs
            .Open "Select StartingSalary From Employees Where EmpID='" & strEmpID & "'", con, adOpenForwardOnly, adLockReadOnly
            iSalary = Val(!StartingSalary & "")
            .Close
        End With
        Set rs = Nothing
    End If
    
    GetRate = iSalary / iTotalWorkingHrs
    
    Exit Function
err:
    MsgBox err.Description
    
End Function

Public Function TotalMonthDays(iMonth As Integer, iYear As Integer) As Integer

    Dim TDate As Date
    Dim iTot As Integer

    TDate = DateSerial(iYear, iMonth, 1)
    iTot = DateAdd("m", 1, TDate) - TDate

    TotalMonthDays = iTot

End Function

Public Function TotalMonthWorkingHrs(iMonth As Integer, iYear As Integer, Optional iTotalSundays As Integer = 0, Optional iHDays As Integer = 0) As Integer

    Dim iMonthDays As Integer
    iMonthDays = TotalMonthDays(iMonth, iYear)
    If iTotalSundays = 0 Then
        iTotalSundays = TotalSundays(iMonth, iYear, iMonthDays)
    End If
    iMonthDays = iMonthDays - iTotalSundays
    
    If iHDays = 0 Then
        iHDays = Val(con.Execute("Select count(DT) from Holidays Where month(DT)=" & _
         iMonth & " and Year(DT)=" & iYear).Fields(0))
    End If
         
    iMonthDays = iMonthDays - iHDays
    TotalMonthWorkingHrs = iMonthDays * 8
    
End Function


Public Sub getLunchTimingsByEmpID(strEmpID As String, ByRef strBrkOutTime As String, ByRef strBrkInTime As String)
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim bExempt As Boolean, strDeptID As String
    Dim strSQL As String
    With rs
        .Open "Select ExemptSettings,DeptID From Employees Where EmpID='" & strEmpID & "'", con, adOpenForwardOnly, adLockReadOnly
        bExempt = !ExemptSettings
        strDeptID = !DeptID & ""
        .Close
        If bExempt Then
            strSQL = "Select LunchInTime,LunchOutTime From EmpSettings Where EmpID='" & strEmpID & "'"
        Else
            strSQL = "Select LunchInTime,LunchOutTime From GeneralSettings Where DeptID='" & strDeptID & "'"
        End If
        .Open strSQL, con, adOpenForwardOnly, adLockReadOnly
        strBrkOutTime = !LunchInTime & ""
        strBrkInTime = !LunchOutTime & ""
        .Close
    End With
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
End Sub

Public Sub getInOutTimingsByEmpID(strEmpID As String, ByRef strInTime As String, ByRef strOutTime As String)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim bExempt As Boolean, strDeptID As String
    Dim strSQL As String
    With rs
        .Open "Select ExemptSettings,DeptID From Employees Where EmpID='" & strEmpID & "'", con, adOpenForwardOnly, adLockReadOnly
        bExempt = !ExemptSettings
        strDeptID = !DeptID & ""
        .Close
        If bExempt Then
            strSQL = "Select EnterTime,ExitTime From EmpSettings Where EmpID='" & strEmpID & "'"
        Else
            strSQL = "Select EnterTime,ExitTime From GeneralSettings Where DeptID='" & strDeptID & "'"
        End If
        .Open strSQL, con, adOpenForwardOnly, adLockReadOnly
        strInTime = !EnterTime & ""
        strOutTime = !ExitTime & ""
        .Close
    End With
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
End Sub

Public Function TakeAttendance(EmpId As String, AttDate As Date, AttTime As Date, Optional bAuto As Boolean = False, Optional bSecondTimeOnly As Boolean = False) As Boolean

    On Error GoTo err

    Dim ExtraDeds As Double, ExtraHrs As Double
    ExtraDeds = 0
    ExtraHrs = 0
    Dim MultipleAllowed As Boolean
    Dim bDeductLHrs As Boolean
    Dim bExempt As Boolean
    Dim DeptID As String
    Dim dLateHrs As Double
    Dim TotalEntries As Integer
    Dim AttRecs As Integer
    Dim bOverTime As Boolean
    Dim strEmpInTime As String, strEmpOutTime As String
    Dim strEmpLunchIn As String, strEmpLunchOut As String
    bOverTime = False
    Dim iAttCat As Integer, bLateDeduction As Boolean, strLateDeductionTime As String, bLunchTimeScanning As Boolean, iSaturday_Early_Out_Mins As Integer, dSaturday_Early_Out_Hrs As Double
    
    Dim iMinutes As Integer
    Dim dTotalHrs As Double
    Dim bOverTimePaid As Boolean
    Dim rs As New ADODB.Recordset
    '''''''''''   This Will Is The Employee Exists Or Not    '''''''''''
    
    With con.Execute("SELECT DeptID,EmpID,EmpType,ExemptSettings FROM Employees WHERE EmpID='" & EmpId & "'")
    
        If .EOF Then
            Exit Function
        Else
            EmpId = ![EmpId] & ""
            Emptype = ![Emptype] & ""
            bExempt = ![ExemptSettings] & ""
            DeptID = ![DeptID] & ""
        End If
        .Close
        
        If bExempt Then
            .Open "SELECT EmpAttCat_RefID,EnterTime,ExitTime,LunchInTime,LunchOutTime,DeductLHrs,FridayBreakStartTime,FridayBreakEndTime,FridayCheckOut FROM EmpSettings WHERE EmpID='" & EmpId & "'", con, adOpenForwardOnly, adLockReadOnly
        Else
            .Open "SELECT EmpAttCat_RefID,EnterTime,ExitTime,LunchInTime,LunchOutTime,DeductLHrs,FridayBreakStartTime,FridayBreakEndTime,FridayCheckOut FROM GeneralSettings WHERE DeptID='" & DeptID & "'", con, adOpenForwardOnly, adLockReadOnly
        End If
        
        Dim bNextDayOT As Boolean, myDTToUse As Date
        bNextDayOT = False
        If strCompany = "QEL" Or strCompany = "Dr-Frgz" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
        Dim lCount As Long
        lCount = GetSingleLongValue("COUNT(*)", "EmpTimes", " WHERE EmpID='" & EmpId & "' AND DT='" & AttDate & "'")
        'If AttTime <= #4:00:00 AM# And lCount > 0 Then
        If AttTime <= #4:00:00 AM# Then
            'Check If Out Entry of Previous day pending...
            myDTToUse = DateAdd("d", -1, AttDate)
            With rs
                .Open "SELECT TotalEntries,SecondOutTime FROM VEmpTimes2 WHERE EmpID='" & EmpId & "' AND DT='" & myDTToUse & "'", con, adOpenForwardOnly, adLockReadOnly
                If .EOF = False Then
                    If IsNull(!SecondOutTime) Then
                        AttDate = myDTToUse
                        bNextDayOT = True
                    End If
                End If
                .Close
            End With
        End If
    End If
    
        If .EOF Then
            bDeductLHrs = True
        Else
            bDeductLHrs = ![DeductLHrs]
            strEmpInTime = Format(!EnterTime, "HH:NN") '"08:00" 'Format(!EnterTime, "HH:NN")
            If Format(AttDate, "ddd") = "Fri" And (strCompany = "Sunlike" Or strCompany = "Dr-Frgz" Or strCompany = "Tecno" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne") Then
                strEmpOutTime = Format(!FridayCheckOut, "HH:NN")
                strEmpLunchIn = Format(!FridayBreakStartTime, "HH:NN")
                strEmpLunchOut = Format(!FridayBreakEndTime, "HH:NN")
                Dim dFridayExtraHrs As Double
                dFridayExtraHrs = GetSingleDoubleValue("DateDiff(n,EnterTime,FridayCheckOut)-DateDiff(n,FridayBreakStartTime,FridayBreakEndTime)", "VEmpSettings", " WHERE EmpID='" & EmpId & "'")
                dFridayExtraHrs = dFridayExtraHrs / 60
                If dFridayExtraHrs < 8 Then
                    dFridayExtraHrs = (8 - dFridayExtraHrs)
                Else
                    dFridayExtraHrs = 0
                End If
            Else
                strEmpOutTime = Format(!ExitTime, "HH:NN") '"17:00" 'Format(!ExitTime, "HH:NN")
                strEmpLunchIn = Format(!LunchInTime, "HH:NN") '"12:30" 'Format(!LunchInTime, "HH:NN")
                strEmpLunchOut = Format(!LunchOutTime, "HH:NN") '"13:30" 'Format(!LunchOutTime, "HH:NN")
            End If
            iAttCat = Val(!EmpAttCat_RefID & "")
            'bOverTimePaid = !OverTimePaid
        End If
        .Close
        
        .Open "SELECT LunchTimeScanning,LateDeduction,LateDeductionAfter,OverTimePaid,Saturday_Early_Out_Mins FROM EmployeeAttendanceCatagories WHERE EntryID=" & iAttCat, con, adOpenForwardOnly, adLockReadOnly
       
        bLateDeduction = !LateDeduction
        strLateDeductionTime = Format(DateAdd("n", Val(!LateDeductionAfter & ""), strEmpInTime), "HH:NN")
        bLunchTimeScanning = !LunchTimeScanning
        bOverTimePaid = !OverTimePaid
        iSaturday_Early_Out_Mins = 0
        dSaturday_Early_Out_Hrs = 0
        If Format(AttDate, "ddd") = "Sat" Then
            If Val(!Saturday_Early_Out_Mins & "") > 0 Then
                 iSaturday_Early_Out_Mins = Val(!Saturday_Early_Out_Mins & "")
                 dSaturday_Early_Out_Hrs = iSaturday_Early_Out_Mins / 60
            End If
        End If
        
        .Close
    End With
    
    ''''''''''--------------------''''''''''''''''''''
    'This code is for QEL to Allow Overtime untill 7:30 of next day...
    ''''''''''--------------------''''''''''''''''''''
'    Dim bNextDayOT As Boolean, myDTToUse As Date
'    bNextDayOT = False
'    If strCompany = "QEL" Or strCompany = "Dr-Frgz" Or strCompany = "IAA" OR strCompany = "Kami" Then
'        Dim lCount As Long
'         lCount = GetSingleLongValue("COUNT(*)", "EmpTimes", " WHERE EmpID='" & EmpId & "' AND DT='" & AttDate & "'")
'        If AttTime < #7:30:00 AM# And lCount > 0 Then
'            'Check If Out Entry of Previous day pending...
'            myDTToUse = DateAdd("d", -1, AttDate)
'            With rs
'                .Open "SELECT TotalEntries,SecondOutTime FROM VEmpTimes2 WHERE EmpID='" & EmpId & "' AND DT='" & myDTToUse & "'", con, adOpenForwardOnly, adLockReadOnly
'                If .EOF = False Then
'                    If IsNull(!SecondOutTime) Then
'                        AttDate = myDTToUse
'                        bNextDayOT = True
'                    End If
'                End If
'                .Close
'            End With
'        End If
'    End If
    ''''''''''--------------------''''''''''''''''''''
    'Closing
    ''''''''''--------------------''''''''''''''''''''
    
    Dim Rec As New ADODB.Recordset, CurrTime As String, Hrs As Double, SQL As String, Cond As String
    Dim TimeDiff As Integer

    Cond = "WHERE EmpID='" & EmpId & "' AND DT='" & AttDate & "'"

    CurrTime = Format(AttTime, "HH:NN")

    TimeDiff = Val(con.Execute("SELECT TimeDiff FROM TimeSettings").Fields(0) & "")
    
    Dim lSpecialDeductionHours As Long
    With Rec
        .Open "SELECT TotalEntries FROM VEmpDailyEntries WHERE EmpID='" & EmpId & "' and DT='" & AttDate & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            TotalEntries = 0
        Else
            TotalEntries = .Fields(0) & ""
        End If
        .Close
        
        iMinutes = DateDiff("n", strEmpInTime, strLateDeductionTime) + 1
        .Open "SELECT * FROM EmpTimes WHERE EntryID=(SELECT MAX(EntryID) FROM EmpTimes " & Cond & ")", con, adOpenForwardOnly, adLockReadOnly
            
        If TotalEntries > 0 Then
            
            If TotalEntries = 1 And !OutTime & "" = "" And bSecondTimeOnly = False And strEmpLunchIn <> strEmpLunchOut Then 'Means It's Lunch In Entry...
                '12:50 to 1:10
'                If CurrTime < #12:25:00 PM# Then
                If strEmpLunchIn <> strEmpLunchOut Then
                    If CurrTime < CDate(DateAdd("N", -5, strEmpLunchIn)) Then
                        ExtraDeds = 0 'DateDiff("n", CurrTime, "12:55") / 60
                    ElseIf CurrTime > CDate(DateAdd("N", 5, strEmpLunchIn)) Then
                        ExtraDeds = 0 'DateDiff("n", "13:00", CurrTime) / 60
                    End If
                End If
                'If DateDiff("n", strEmpInTime, CDate(!InTime)) > 0 And DateDiff("n", strEmpInTime, CDate(!InTime)) < 11 Then this was original
                
                If strEmpLunchIn = strEmpLunchOut Then
                    dTotalHrs = DateDiff("n", strEmpInTime, strEmpOutTime) / 60
                Else
                    dTotalHrs = DateDiff("n", strEmpInTime, strEmpLunchIn) / 60
                End If
                If DateDiff("n", strEmpInTime, CDate(!InTime)) > 0 And DateDiff("n", strEmpInTime, CDate(!InTime)) < iMinutes Then
                    ExtraHrs = DateDiff("n", strEmpInTime, CDate(!InTime)) / 60
                ElseIf DateDiff("n", strEmpInTime, CDate(!InTime)) < 0 Then
                    ExtraDeds = ExtraDeds + DateDiff("n", CDate(!InTime), strEmpInTime) / 60
                End If
                
                If DateDiff("n", strEmpInTime, CDate(!InTime)) > iMinutes Then  'Give Back the Relaxion Minutes....
                    ExtraHrs = Round((iMinutes - 1) / 60)
                End If
                '/\/\/\/\/\/\/Friday Timing Working/\/\/\//\/\/\/\/\
'                If Format(AttDate, "dddd") = "Friday" Then
'                    ExtraHrs = ExtraHrs '+ 0.25
'                    If CurrTime > #1:00:00 PM# Then
'                        Hrs = DateDiff("n", CDate(!InTime), #1:00:00 PM#)
'                    Else
'                        Hrs = DateDiff("n", CDate(!InTime), CurrTime)
'                    End If
'                Else
'                    If CurrTime > CDate(DateAdd("N", -6, strEmpLunchIn)) Then '#12:24:00 PM#
'                        Hrs = DateDiff("n", CDate(!InTime), strEmpLunchIn) '#12:30:00 PM#
'                    Else
'                        Hrs = DateDiff("n", CDate(!InTime), CurrTime)
'                    End If
'                End If

                If strEmpLunchIn = strEmpLunchOut Then
                    If CurrTime > CDate(DateAdd("N", -6, strEmpOutTime)) Then '#12:24:00 PM#
                        Hrs = DateDiff("n", CDate(!InTime), strEmpOutTime) '#12:30:00 PM#
                    Else
                        Hrs = DateDiff("n", CDate(!InTime), CurrTime)
                    End If
                Else
                    Dim strOutTime_Check As String
                    If CurrTime > strEmpOutTime Then
                        strOutTime_Check = strEmpOutTime
                    Else
                        strOutTime_Check = CurrTime
                    End If
                    If CurrTime > CDate(DateAdd("N", -6, strEmpLunchIn)) Then '#12:24:00 PM#
'                        If strCompany = "Tecno" Then
'                            Hrs = DateDiff("n", strEmpInTime, strEmpLunchIn) '#12:30:00 PM#
'                        Else
'                            Hrs = DateDiff("n", CDate(!InTime), strEmpLunchIn) '#12:30:00 PM#
'                        End If
                        If CDate(!InTime) > strEmpLunchIn Then
                            If CDate(!InTime) > strEmpLunchOut Then
                                Hrs = DateDiff("n", CDate(!InTime), strOutTime_Check)
                            Else
                                Hrs = DateDiff("n", strEmpLunchOut, strOutTime_Check)
                            End If
                        Else
                            Hrs = DateDiff("n", CDate(!InTime), strEmpLunchIn) '#12:30:00 PM#
                        End If
                    Else
                        Hrs = DateDiff("n", CDate(!InTime), CurrTime)
                    End If
                   
                End If
                '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
            ElseIf (TotalEntries = 2 And !OutTime & "" = "") Or bSecondTimeOnly Or (TotalEntries = 1 And !OutTime & "" = "" And (strEmpLunchIn = strEmpLunchOut)) Then 'Means It's Lunch Out Entry
                'Here we will check for over time, if overtime employee then first add over time
                Dim p_dOTHrs As Double
                If CurrTime > strEmpOutTime And bOverTimePaid Then
                    'Remember to Add OverTime Entries
                    p_dOTHrs = DateDiff("n", strEmpOutTime, CurrTime) / 60
                    con.Execute "INSERT INTO EmpTimes(EmpID,DT,InTime,OutTime,Hrs,PayableHrs,UserName,ComputerName,OverTime) VALUES('" & EmpId & "','" & AttDate & "','" & strEmpOutTime & "','" & CurrTime & "'," & p_dOTHrs & "," & p_dOTHrs & ",'" & CurrentUserName & "','" & strComputerName & "',1)"
                    CurrTime = strEmpOutTime
                ElseIf ((strCompany = "QEL" Or strCompany = "Dr-Frgz" Or strCompany = "Towne") And CurrTime >= Format(DateAdd("n", 30, strEmpOutTime), "hh:nn")) Or (bNextDayOT And bOverTimePaid) Then
                    'If Compan QEL & Out Time>=30 Mins.
                    If bNextDayOT Then
                        p_dOTHrs = DateDiff("n", AttDate & " " & strEmpOutTime, DateAdd("d", 1, AttDate) & " " & CurrTime) / 60
                    Else
                        p_dOTHrs = DateDiff("n", strEmpOutTime, CurrTime) / 60
                    End If
                    con.Execute "INSERT INTO EmpTimes(EmpID,DT,InTime,OutTime,Hrs,PayableHrs,UserName,ComputerName,OverTime) VALUES('" & EmpId & "','" & AttDate & "','" & strEmpOutTime & "','" & CurrTime & "'," & p_dOTHrs & "," & p_dOTHrs & ",'" & CurrentUserName & "','" & strComputerName & "',1)"
                    CurrTime = strEmpOutTime
                End If
                'dTotalHrs = DateDiff("n", strEmpLunchOut, "16:30") / 60
                If strEmpLunchIn = strEmpLunchOut Then
                    dTotalHrs = DateDiff("n", strEmpInTime, strEmpOutTime) / 60
                Else
                    dTotalHrs = DateDiff("n", strEmpLunchOut, strEmpOutTime) / 60
                End If
'                If Format(AttDate, "dddd") = "Friday" Then
'                    ExtraHrs = ExtraHrs + 0.5
'                    If CDate(!InTime) > #2:05:00 PM# Then
'                        ExtraDeds = 0 'DateDiff("n", "14:30", CDate(!InTime)) / 60
'                    ElseIf CDate(!InTime) < #2:00:00 PM# Then
'                        ExtraDeds = DateDiff("n", CDate(!InTime), "14:00") / 60
'                    End If
'                    If DateDiff("n", "14:00", CDate(!InTime)) > 0 And DateDiff("n", "14:00", CDate(!InTime)) < 6 Then
'                        ExtraHrs = ExtraHrs + DateDiff("n", "13:30", CDate(!InTime)) / 60
'                    End If
'                    Hrs = 0.5
'                Else
                    
                    If Format(AttDate, "dddd") = "Friday" Then
                        ExtraHrs = dFridayExtraHrs  'FridayExtraHrs Adding
                    ElseIf Format(AttDate, "dddd") = "Saturday" Then
                        ExtraHrs = dSaturday_Early_Out_Hrs
                    End If
                    
                    If strEmpLunchIn = strEmpLunchOut Then
                        If DateDiff("n", strEmpInTime, CDate(!InTime)) > 0 And DateDiff("n", strEmpInTime, CDate(!InTime)) < iMinutes Then
                            ExtraHrs = DateDiff("n", strEmpInTime, CDate(!InTime)) / 60
                        End If
                    Else
                        If DateDiff("n", strEmpLunchOut, CDate(!InTime)) > 0 And DateDiff("n", strEmpLunchOut, CDate(!InTime)) < 6 Then '"13:30"
                            ExtraHrs = ExtraHrs + DateDiff("n", strEmpLunchOut, CDate(!InTime)) / 60 '"13:30"
                        End If
                    End If
                    
                    If strEmpLunchIn = strEmpLunchOut Then
                        
'                        If CDate(!InTime) > CDate(DateAdd("n", 5, strEmpOutTime)) Then  '#1:35:00 PM#
'                            ExtraDeds = 0 'DateDiff("n", "14:00", CDate(!InTime)) / 60
'                        ElseIf !InTime < strEmpOutTime Then '#1:30:00 PM#
'                            ExtraDeds = DateDiff("n", CDate(!InTime), strEmpOutTime) / 60 '"13:30"
'                        End If
                        ExtraDeds = 0
                    Else
                        If CDate(!InTime) > CDate(DateAdd("n", 5, strEmpLunchOut)) Then  '#1:35:00 PM#
                            ExtraDeds = 0 'DateDiff("n", "14:00", CDate(!InTime)) / 60
                        ElseIf !InTime < strEmpLunchOut Then '#1:30:00 PM#
                            ExtraDeds = DateDiff("n", CDate(!InTime), strEmpLunchOut) / 60 '"13:30"
                        End If
                    End If
                'End If
                
                If CurrTime > strEmpOutTime Then '#5:00:00 PM#
                    ExtraDeds = ExtraDeds + 0 'DateDiff("n", "17:00", CurrTime) / 60
                End If
                
                If CurrTime > strEmpOutTime Then '#5:00:00 PM#
                    Hrs = (Hrs * 60) + DateDiff("n", CDate(!InTime), strEmpOutTime) '#5:00:00 PM#
                Else
                    Hrs = (Hrs * 60) + DateDiff("n", CDate(!InTime), CurrTime)
                End If
                
            End If
            'Check ... Is This Out Time Entry ?
            
            If TotalEntries < 3 Then
                If !OutTime & "" = "" Then
                    'Hrs = DateDiff("n", CDate(!InTime), CDate(CurrTime))
                
                    'Check for the Minimum Time Difference
                    If Hrs < TimeDiff Then GoTo END_OF_PROCEDURE
                
                    Hrs = Round(Hrs / 60, 2)
                            
                    'Update Out Time and Do Nothing Else
                    
                    SQL = "UPDATE EmpTimes SET OutTime='" & CurrTime & "',Hrs=" & Hrs & ",PayableHrs=" & IIf(bLateDeduction, Round((Hrs + ExtraHrs - ExtraDeds), 2), dTotalHrs) & " WHERE EntryID=" & Val(!EntryID & "") & ""
                     
                    con.Execute SQL
                    'con.Execute "Update AttendanceSheet Set Attendance=1 Where EmpID='" & EmpId & "' AND DT='" & AttDate & "'"
                    GoTo END_OF_PROCEDURE
                Else
                    'Check for the Minimum Time Difference
                
                    Hrs = DateDiff("n", CDate(!OutTime), CDate(CurrTime))
                    If TotalEntries <> 2 Then
                        If Hrs < TimeDiff Then GoTo END_OF_PROCEDURE
                    End If
                End If
            End If
        Else
            iMinutes = DateDiff("n", strEmpInTime, strLateDeductionTime)
            If bDeductLHrs And DateDiff("n", strEmpInTime, CurrTime) > iMinutes Then '"08:00"
                dLateHrs = DateDiff("n", strEmpInTime, CurrTime) / 60 '"08:00"
            Else
                dLateHrs = 0
            End If
            AttRecs = Val(con.Execute("SELECT COUNT(*) FROM AttendanceSheet WHERE EmpID='" & EmpId & "' AND DT='" & AttDate & "'").Fields(0).Value & "")
            If AttRecs <> 0 Then con.Execute "DELETE FROM AttendanceSheet WHERE EmpID='" & EmpId & "' AND DT='" & AttDate & "'"
            
            lSpecialDeductionHours = 0
'            If strCompany = "Tecno" Then
'                Dim iLateMins As Integer
'                iLateMins = DateDiff("n", strEmpInTime, CurrTime)
'                If iLateMins > 5 And iLateMins < 16 Then
'                    lSpecialDeductionHours = 1
'                ElseIf iLateMins > 15 Then
'                    lSpecialDeductionHours = 4
'                Else
'                    lSpecialDeductionHours = 0
'                End If
'            End If
            
            con.Execute "INSERT INTO AttendanceSheet(EmpID,DT,Attendance,LateHours,ExtraHrsToDeduct) VALUES('" & EmpId & "','" & AttDate & "',1," & _
             dLateHrs & "," & lSpecialDeductionHours & ")", a
            
        End If

        '.Close
        
        If Emptype = 2 Then
            
            ' If No Previous Entry Found Accept The Entry
            If TotalEntries = 0 Then GoTo ACCEPT_ENTRY
        
            'Remember The Total Entries
        
            ' If this is Second Time Entry Accept It
            If TotalEntries < 2 Then GoTo ACCEPT_ENTRY

        End If
            
        ' If Two Entries Have Been Made ... Check For Over Time Authorization
        If TotalEntries >= 2 Then
            Dim rsOT As New ADODB.Recordset
            Dim bAuthorized As Boolean
            Dim dTotalOTHrs As Double
            Dim dExcessHrs As Double, dValidHrs As Double
            With rsOT
                bAuthorized = True
                If bAuthorized = False Or bOverTimePaid = False Then
'                    MsgBox "Not Authorized.", vbInformation
                    GoTo END_OF_PROCEDURE
                End If
            End With
            'If Val(![Authority] & "") <> 1 Then GoTo END_OF_PROCEDURE
            
            ' Other Wise Accept The Entry
            If !OutTime & "" <> "" Then
                If bAuto Then
                    con.Execute "INSERT INTO EmpTimes(EmpID,DT,InTime,ManualEntry,OverTime) VALUES('" & EmpId _
                     & "','" & AttDate & "','" & CurrTime & "',0,1)"
                Else
                    con.Execute "INSERT INTO EmpTimes(EmpID,DT,InTime,ManualEntry,OverTime,UserName,ComputerName) VALUES('" & EmpId _
                     & "','" & AttDate & "','" & CurrTime & "',1,1,'" & CurrentUserName & "','" & strComputerName & "')"
                End If
            Else
                Hrs = DateDiff("n", CDate(!InTime), CurrTime)
                If Hrs < TimeDiff Then GoTo END_OF_PROCEDURE
                Hrs = Round(Hrs / 60, 2)
'                dTotalOTHrs = OverTimeHrsUsedInAWeek(EmpId, AttDate)
'                If Hrs + dTotalOTHrs > 12 Then
'                    dValidHrs = 12 - dTotalOTHrs
'                    dExcessHrs = Hrs - dValidHrs
'                Else
'                    dValidHrs = Hrs
'                    dExcessHrs = 0
'                End If
'                If Hrs > 5 Then
'                    Hrs = 6
'                End If
                dValidHrs = Hrs
                dExcessHrs = 0
                con.Execute "UPDATE EmpTimes SET OutTime='" & CurrTime & "',Hrs=" & Hrs & ",PayableHrs=" & (dValidHrs) & " WHERE EntryID=" & Val(!EntryID & "") & ""
                If dExcessHrs > 0 Then
                    con.Execute "INSERT INTO EmpExcessOverTime(EmpID,DT,InTime,OutTime,HrsInSalarySheet,ExcessHrs) Values('" & _
                    EmpId & "','" & AttDate & "','" & !InTime & "','" & CurrTime & "'," & dValidHrs & "," & dExcessHrs & ")"
                End If
            End If
            GoTo END_OF_PROCEDURE
            bOverTime = True
        End If
        
    End With

    ' Enter The New Entry With In Time

ACCEPT_ENTRY:
    If bAuto Then
        SQL = "INSERT INTO EmpTimes(EmpID,DT,InTime,ManualEntry) VALUES('" & EmpId _
         & "','" & AttDate & "','" & CurrTime & "',0)"
    Else
        SQL = "INSERT INTO EmpTimes(EmpID,DT,InTime,ManualEntry,UserName,ComputerName) VALUES('" & EmpId _
         & "','" & AttDate & "','" & CurrTime & "',1,'" & CurrentUserName & "','" & strComputerName & "')"
    End If
    
    con.Execute SQL
'    TotalEntries = Val(con.Execute("Select TotalEntries From VEmpDailyEntries Where EmpID='" & EmpID & "' and DT='" & AttDate & "'").Fields(0).Value & "")
'
'    If TotalEntries = 1 Then    'That Means that it's first entry of the day. So Calculate LateHours...
'        If DateDiff("n", "08:00", CurrTime) > 15 Then
'            con.Execute "Update AttendanceSheet Set LateHours=" & DateDiff("n", "08:00", CurrTime) / 60 & " Where EmpID='" & EmpID & "' and DT='" & AttDate & "'"
'        End If
'    End If
    
END_OF_PROCEDURE:
    
    If Rec.State = 1 Then Rec.Close
    Set Rec = Nothing
    TakeAttendance = True
    Exit Function
err:
    Debug.Print err.Description
    'So That Next Entries Should Not Be Skipped
    Resume Next
    
End Function

Public Function OverTimeHrsUsedInAWeek(strEmpID As String, DT As Date) As Double

    On Error GoTo err
    Dim cmd As New ADODB.Command
    Dim rs As ADODB.Recordset
    Dim dTotalOTHrs As Double
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "EmpOverTime_SP"
        .Parameters("@EmpID").Value = strEmpID
        .Parameters("@AttDate").Value = DT
        Set rs = .Execute()
    End With
    If rs.EOF Then
        dTotalOTHrs = 0
    Else
        dTotalOTHrs = Val(rs!TotalOTHrs & "")
    End If
    OverTimeHrsUsedInAWeek = dTotalOTHrs
    Exit Function
err:
    MsgBox err.Description
    
End Function

Public Sub PrintOTSheet()

    On Error GoTo err
     
    Dim frmDT As New frmGetDateSingle, myDT As Date, strSelection As String
    If frmDT.getDate(myDT) = False Then
        Exit Sub
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpOverTimeSheet.rpt")
    Dim frm As New frmPrevRpt
    strSelection = "MONTH({VEmpOTHrsValidnExcess.DT})=" & Month(myDT) & " AND YEAR({VEmpOTHrsValidnExcess.DT})=" & year(myDT) & " AND {VEmpOTHrsValidnExcess.ValidOTHrs}+{VEmpOTHrsValidnExcess.ExcessOTHrs}>0"
    
    frm.ShowReport strSelection, rpt
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Public Sub PrintDailyOTSheet()

    On Error GoTo err
     
    Dim frmDT As New frmGetDateDouble, myDT As Date, myDT1 As Date, strSelection As String
    If frmDT.getDate(myDT, myDT1) = False Then
        Exit Sub
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptDailyOverTime.rpt")
    Dim frm As New frmPrevRpt
    rpt.FormulaFields.GetItemByName("DTFrom").Text = "#" & myDT & "#"
    rpt.FormulaFields.GetItemByName("DTTo").Text = "#" & myDT1 & "#"
    strSelection = "{VEmpOTHrsValidnExcess.DT}={@DTFrom} to {@DTTo}"
    frm.ShowReport strSelection, rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Public Sub PrintDailyLeaves()

    On Error GoTo err
     
    Dim frmDT As New frmGetDateDouble, myDT As Date, myDT1 As Date, strSelection As String
    If frmDT.getDate(myDT, myDT1) = False Then
        Exit Sub
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptDailyLeaves.rpt")
    Dim frm As New frmPrevRpt
    rpt.FormulaFields.GetItemByName("DTFrom").Text = "#" & myDT & "#"
    rpt.FormulaFields.GetItemByName("DTTo").Text = "#" & myDT1 & "#"
    strSelection = "{Leaves.DT}={@DTFrom} to {@DTTo}"
    frm.ShowReport strSelection, rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Public Sub ReCalculateAttendance()
    
    On Error GoTo err
    Dim myDTFrom As Date, myDTTo As Date, p_strEmpID As String, p_strDeptID As String
    Dim f As New frmGetEmpIDAndDateRange
    If f.getEmpIDAndDateRange(myDTFrom, myDTTo, p_strDeptID, p_strEmpID) = False Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    Dim strFirstInTime As String, strFirstOutTime As String
    Dim strSecondInTime As String, strSecondOutTime As String
    Dim strOTInTime As String, strOTOutTime As String
    Dim strEmpID As String, myDT As Date
    Dim strCondition As String
    strCondition = "WHERE DT BETWEEN '" & myDTFrom & "' AND '" & myDTTo & "'"
    
    If p_strDeptID <> "0" Then
        strCondition = strCondition & " AND DeptID='" & p_strDeptID & "'"
    End If
    
    If p_strEmpID <> "0" Then
        strCondition = strCondition & " AND EmpID='" & p_strEmpID & "'"
    End If
    Dim dOTHrs As Double
    With rs
        .Open "SELECT * FROM VEmpTimes2 " & strCondition & " ORDER BY EmpID,DT", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            strEmpID = !EmpId & ""
              
            strFirstInTime = ""
            strFirstOutTime = ""
            strSecondInTime = ""
            strSecondOutTime = ""
            strOTInTime = ""
            strOTOutTime = ""
            
            myDT = !DT
            strFirstInTime = !FirstInTime & ""
            strFirstOutTime = !FirstOutTime & ""
            
            If !LunchInTime & "" = !LunchOutTime & "" Then
                If !SecondOutTime <> "" Then
                    strFirstOutTime = !SecondOutTime & ""
                End If
            Else
                strSecondInTime = !SecondInTime & ""
                strSecondOutTime = !SecondOutTime & ""
                strOTInTime = !OTInTime & ""
                strOTOutTime = !OTOutTime & ""
                dOTHrs = Val(!OTHrs & "")
            End If
           
            
            Call StartTrans(con)
            
            con.Execute "DELETE FROM EmpTimes WHERE EmpID='" & strEmpID & "' AND DT='" & myDT & "'"
            If strFirstInTime <> "" Then
                Call TakeAttendance(strEmpID, myDT, Format(strFirstInTime, "HH:NN"))
            End If
            If strFirstOutTime <> "" Then
                Call TakeAttendance(strEmpID, myDT, Format(strFirstOutTime, "HH:NN"))
            End If
            If strSecondInTime <> "" Then
                Call TakeAttendance(strEmpID, myDT, Format(strSecondInTime, "HH:NN"))
            End If
            
            If strOTOutTime <> "" Then
                If strOTOutTime < #7:00:00 AM# And dOTHrs > 3 Then
                    Call TakeAttendance(strEmpID, DateAdd("d", 1, myDT), Format(strOTOutTime, "HH:NN"))
                Else
                    Call TakeAttendance(strEmpID, myDT, Format(strOTOutTime, "HH:NN"))
                End If
            Else
                If strSecondOutTime <> "" Then
                    Call TakeAttendance(strEmpID, myDT, Format(strSecondOutTime, "HH:NN"))
                End If
                If strOTInTime <> "" Then
                    Call TakeAttendance(strEmpID, myDT, Format(strOTInTime, "HH:NN"))
                End If
                If strOTOutTime <> "" Then
                    Call TakeAttendance(strEmpID, myDT, Format(strOTOutTime, "HH:NN"))
                End If
            End If
            
            
            
'            If strOTOutTime <> "" Then
'                Call TakeAttendance(strEmpID, myDT, Format(strOTOutTime, "HH:NN"))
'                If strOTOutTime < #7:00:00 AM# And dOTHrs > 3 Then
'                    Call TakeAttendance(strEmpID, DateAdd("d", 1, myDT), Format(strOTOutTime, "HH:NN"))
'                Else
'                    Call TakeAttendance(strEmpID, myDT, Format(strOTOutTime, "HH:NN"))
'                End If
'            End If
            con.CommitTrans
'            Debug.Print strEmpID & ";" & Format(myDT, "dd-MMM-yyyy")
            
            .MoveNEXT
        Loop
        .Close
    End With
    MsgBox "Done!!!", vbInformation
     
    Exit Sub
err:
    MsgBox err.Description
End Sub

Public Sub PrintBloodGroupReport(strBloodGroup As String)
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\EmployeeListBloodGroupWise.rpt")
    Dim f As New frmPrevRpt
    rpt.FormulaFields.GetItemByName("MainHeading").Text = "'" & strBloodGroup & "'"
    f.ShowReport "{VEmp.BloodGroup}='" & strBloodGroup & "' AND {VEmp.Active}=TRUE AND {VEmp.TempDept}=FALSE", rpt
    
End Sub

Public Sub PrintLeaveList()
    'PrintMissingAttendance
    On Error GoTo err
     
    Dim frmGet As New frmGetEmpIDAndDateRange, myDTFrom As Date, myDTTo As Date, strDeptID As String, strEmpID As String
    If frmGet.getEmpIDAndDateRange(myDTFrom, myDTTo, strDeptID, strEmpID) = False Then
        Exit Sub
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\LeavesList.rpt")
    Dim frm As New frmPrevRpt
    Dim strSelection As String
    strSelection = "{Employees.TempDept}=FALSE AND {Leaves.DT}=#" & myDTFrom & "# to #" & myDTTo & "#"
    rpt.FormulaFields.GetItemByName("MainHeading").Text = "'" & Format(myDTFrom, "dd-MMM-yyyy") & " to " & Format(myDTTo, "dd-MMM-yyyy") & "'"
    If strEmpID = "0" Then
        If strDeptID <> "0" Then
            strSelection = strSelection & " AND {Employees.DeptID}='" & strDeptID & "'"
        End If
    Else
        strSelection = strSelection & " AND {Employees.EmpID}='" & strEmpID & "'"
    End If
    
    frm.ShowReport strSelection, rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Public Sub PrintMissingAttendance()
    
    On Error GoTo err
     
    Dim frmGet As New frmGetEmpIDAndDateRange, myDTFrom As Date, myDTTo As Date
    Dim strDeptID As String, strEmpID As String
    If frmGet.getEmpIDAndDateRange(myDTFrom, myDTTo, strDeptID, strEmpID) = False Then
        Exit Sub
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\MissingAttendance.rpt")
    rpt.GroupSelectionFormula = "COUNT({EmpTimes.EmpID},{Employees.empid})=1"
    Dim frm As New frmPrevRpt
    Dim strSelection As String
    strSelection = "{EmpTimes.DT}=#" & myDTFrom & "#"
    If strDeptID <> "0" Then
        strSelection = strSelection & " AND {Employees.DeptID}='" & strDeptID & "'"
    End If
    rpt.FormulaFields.GetItemByName("MainHeading").Text = "'" & Format(myDTFrom, "dd-MMM-yyyy") & "'"
      
    frm.ShowReport strSelection, rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Public Sub PrintTerminatedEmpReport()
    '
    On Error GoTo err
     
    Dim frmGet As New frmGetDateDouble, myDTFrom As Date, myDTTo As Date
    If frmGet.getDate(myDTFrom, myDTTo) = False Then
        Exit Sub
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\TerminatedEmployees.rpt")
    
    Dim frm As New frmPrevRpt
    Dim strSelection As String
    strSelection = "{VEmpLastInActiveDate.InActiveDT}=#" & myDTFrom & "# TO # " & myDTTo & "#"
    rpt.FormulaFields.GetItemByName("MainHeading").Text = "'" & Format(myDTFrom, "dd-MMM-yyyy") & " To " & Format(myDTTo, "dd-MMM-yyyy") & "'"
      
    frm.ShowReport strSelection, rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Public Sub PrintNewHiredReport()
    'PrintLateAttendanceReport
    On Error GoTo err
     
    Dim frmGet As New frmGetDateDouble, myDTFrom As Date, myDTTo As Date
    If frmGet.getDate(myDTFrom, myDTTo) = False Then
        Exit Sub
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\NewHiredEmployees.rpt")
    
    Dim frm As New frmPrevRpt
    Dim strSelection As String
    strSelection = "{Employees.JoinDate}=#" & myDTFrom & "# TO # " & myDTTo & "#"
    rpt.FormulaFields.GetItemByName("MainHeading").Text = "'" & Format(myDTFrom, "dd-MMM-yyyy") & " To " & Format(myDTTo, "dd-MMM-yyyy") & "'"
      
    frm.ShowReport strSelection, rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Public Sub PrintLateAttendanceReport()
    '
    On Error GoTo err
     
    Dim frmGet As New frmGetEmpIDAndDateRange, myDTFrom As Date, myDTTo As Date
    Dim strEmpID As String, strDeptID As String
    If frmGet.getEmpIDAndDateRange(myDTFrom, myDTTo, strDeptID, strEmpID) = False Then
        Exit Sub
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\LateAttReportDeptwise.rpt")
    
    Dim frm As New frmPrevRpt
    Dim strSelection As String
    
    strSelection = "{AttendanceSheet.DT}=#" & myDTFrom & "# TO # " & myDTTo & "#"
    If strDeptID <> "0" Then
        strSelection = strSelection & " AND {Employees.DeptID}='" & strDeptID & "'"
    End If
    
    rpt.FormulaFields.GetItemByName("MainHeading").Text = "'" & Format(myDTFrom, "dd-MMM-yyyy") & " To " & Format(myDTTo, "dd-MMM-yyyy") & "'"
      
    frm.ShowReport strSelection, rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Public Sub PrintOverTimeSheet(iOption As Integer)

    Dim f As New frmGetEmpIDAndDateRange
    Dim DtFrom As Date, DtTo As Date, strEmpID As String, strDeptID As String
    If f.getEmpIDAndDateRange(DtFrom, DtTo, strDeptID, strEmpID) = False Then Exit Sub
    Dim rpt As CRAXDDRT.Report, strCond As String
    If iOption = 0 Or iOption = 1 Then
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpOverTimeSheet.rpt")
        strCond = ""
'    ElseIf iOption = 1 Then
'        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpOverTimePaymentSheetValid.rpt")
'        strCond = "{VEmp.TotalOTHrs}>0"
    ElseIf iOption = 2 Then
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpOverTimePaymentSheetExcess.rpt")
        strCond = "{VEmp.TotalExcessHrs}>0"
    ElseIf iOption = 3 Then
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\OverTimeSupvervisoryStaff.rpt")
        'strCond = "{VEmp.TotalExcessHrs}>0"
    End If
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DtFrom
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DtTo
    
    rpt.ParameterFields.GetItemByName("@DeptID").SetCurrentValue IIf(strDeptID = "0", "", strDeptID)
    rpt.ParameterFields.GetItemByName("@EmpID").SetCurrentValue IIf(strEmpID = "0", "", strEmpID)
    
    rpt.EnableParameterPrompting = False
        
    rpt.FormulaFields.GetItemByName("Heading").Text = "'" & Format(DtFrom, "dd-MMM-yyyy") & " to " & Format(DtTo, "dd-MMM-yyyy") & "'"
    
    Dim iMonthDays As Integer, lHolidays As Long, iSundays As Integer
    Dim myDTFrom As Date, myDTTo As Date
    
    myDTFrom = DateSerial(year(DtFrom), Month(DtFrom), 1)
    myDTTo = DateSerial(year(DtFrom), Month(DtFrom) + 1, 0)
    iMonthDays = Day(myDTTo)
    'iMonthDays = DateDiff("d", myDTFrom, myDTto)
'    lHolidays = GetSingleLongValue("COUNT(DT)", "Holidays", " WHERE DT BETWEEN '" & myDTFrom & "' AND '" & myDTto & "'")
'    iSundays = TotalSundays(Month(myDTFrom), year(myDTFrom))
'    iMonthDays = iMonthDays - iholidays - iSundays
    
    rpt.FormulaFields.GetItemByName("MonthDays").Text = iMonthDays
    
    Dim frm As New frmPrevRpt
    frm.ShowReport strCond, rpt, , False
    
End Sub

Public Sub PrintOverTimeIndividualReport()

    Dim frm As New frmGetEmpIDnFinYear
    Dim strEmpID As String, strYear As String
    Dim DTStart As Date, DTFinish As Date, iYear As Integer
    Load frm
    
    If frm.getValues(strEmpID, strYear) = False Then Exit Sub
    iYear = Val(strYear) 'cmbFinYear.ItemData(cmbFinYear.ListIndex)
    DTStart = DateSerial(iYear, 7, 1)
    DTFinish = DateSerial(iYear + 1, 6, 30)
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\EmployeeOTRecord.rpt")
    
    rpt.ParameterFields.GetItemByName("@EmpID").SetCurrentValue strEmpID
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DTStart
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DTFinish
    
    rpt.FormulaFields.GetItemByName("Year").Text = "'" & Left(strYear, 11) & "'"
    
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, , False
    
End Sub

Public Sub PrintOverTimeReportMonthwise(bHours As Boolean)

    Dim frm As New frmGetFinYear
    Dim strYear As String
    Dim DTStart As Date, DTFinish As Date, iYear As Integer
    Load frm
    
    If frm.getValues(strYear) = False Then Exit Sub
    iYear = Val(strYear) 'cmbFinYear.ItemData(cmbFinYear.ListIndex)
    DTStart = DateSerial(iYear, 7, 1)
    DTFinish = DateSerial(iYear + 1, 6, 30)
    
    Dim rpt As CRAXDDRT.Report
    If bHours Then
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\OverTimeHoursYearwise.rpt")
    Else
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\OverTimeAmountsYearwise.rpt")
    End If
   
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DTStart
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DTFinish
    
    rpt.FormulaFields.GetItemByName("Year").Text = "'" & Left(strYear, 11) & "'"
    
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, , False
    
End Sub

Public Sub PrintEmpListSalaryRevisions()
    'PrintEmpListProbationEnd
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\EmpListSalaryRevisions.rpt")
    Dim f As New frmPrevRpt
    Dim myDT As Date
    myDT = DateAdd("yyyy", -1, Date)
    
    rpt.GroupSelectionFormula = "MAXIMUM({VSalaryLedger.DT},{VSalaryLedger.EmpID})<#" & myDT & "#"
    f.ShowReport "{VSalaryLedger.Active}=TRUE", rpt
    
End Sub

Public Function DateDiffInMonths(myDT As Date, Optional myDTTo) As Integer
    
    Dim iMonths As Integer, iDays As Integer
    Dim p_myDTTo As Date
    If IsMissing(myDTTo) Then
        p_myDTTo = Date
    Else
        p_myDTTo = myDTTo
    End If
    
    iMonths = DateDiff("m", myDT, p_myDTTo)
    
    If DateAdd("M", iMonths, myDT) > p_myDTTo Then
        iMonths = iMonths - 1
    End If
    
    DateDiffInMonths = iMonths
    
End Function


Public Sub PrintOTRegister()
    'PrintEmpProductivityReport
    Dim f As New frmGetEmpIDAndDateRange
    Dim DtFrom As Date, DtTo As Date, strCond As String, strDeptID As String, strEmpID As String
    
    If f.getEmpIDAndDateRange(DtFrom, DtTo, strDeptID, strEmpID) = False Then Exit Sub
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\OverTimeMonthlyRegister.rpt")
'    strCond = "{VEmp.TotalExcessHrs}>0"
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DtFrom
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DtTo
    rpt.ParameterFields.GetItemByName("@AttGate").SetCurrentValue ""
    rpt.ParameterFields.GetItemByName("@ReportType").SetCurrentValue 0
    rpt.ParameterFields.GetItemByName("@MainUnit").SetCurrentValue ""
    rpt.ParameterFields.GetItemByName("@DeptID").SetCurrentValue IIf(strDeptID = "0", "", strDeptID)
    rpt.ParameterFields.GetItemByName("@EmpID").SetCurrentValue IIf(strEmpID = "0", "", strEmpID)
    
    rpt.EnableParameterPrompting = False
        
    rpt.FormulaFields.GetItemByName("ForMonth").Text = "'" & Format(DtFrom, "MMM,yyyy") & "'"
    
    
    Dim frm As New frmPrevRpt
    frm.ShowReport strCond, rpt, , False
    
End Sub


Public Sub PrintEmpProductivityReport(Optional bDaily As Boolean = False, Optional bLineChart As Boolean, Optional bOverTime As Boolean = False)

    Dim f As Form
    
    Dim DtFrom As Date, DtTo As Date, strCond As String, strEmpID As String, strDeptID As String, strProcesses As String
    If bDaily Then
        Set f = New frmGetDateSingle
        If f.getDate(DtFrom) = False Then Exit Sub
        Dim strDepts As String
        Set f = New frmGetDepartments
        If f.getData(strDepts) = False Then Exit Sub
    Else
        Set f = New frmGetEmpIDAndDateRangeAndProcesses
        If f.getEmpIDAndDateRangeAndProcess(DtFrom, DtTo, strDeptID, strEmpID, strProcesses) = False Then Exit Sub
    End If
    
    Dim rpt As CRAXDDRT.Report, strSelection As String
    Dim strSubSelectionRegular As String, strSubSelectionOT As String
    strSubSelectionRegular = "{VEmpProductivityReport.DTOnly}={?Pm-@DTFrom} TO {?Pm-@DTTo} and {VEmpProductivityReport.OverTime}=FALSE and {VEmpProductivityReport.empid} = {?Pm-VEmpProductivityReport.empid}"
    strSubSelectionOT = "{VEmpProductivityReport.DTOnly}={?Pm-@DTFrom} TO {?Pm-@DTTo} and {VEmpProductivityReport.OverTime}=TRUE AND {VEmpProductivityReport.empid} = {?Pm-VEmpProductivityReport.empid}"
    If bDaily Then
        Set rpt = rptApp.OpenReport(RptPath & "\EmployeeProductivityReportDatewise.rpt")
        rpt.FormulaFields.GetItemByName("SubHeading").Text = "'" & Format(DtFrom, "dd-MMM-yyyy") & "'"
        rpt.FormulaFields.GetItemByName("ReportType").Text = "'" & IIf(bOverTime, "Overtime", "Regular") & "'"
        strSelection = "{VEmpProductivityReport.DTOnly}=#" & DtFrom & "# AND {VEmpProductivityReport.OverTime}=" & bOverTime
        If strDepts <> "All" Then
            strSelection = strSelection & " AND {VEmpProductivityReport.DeptID} IN[" & strDepts & "]"
        End If
    ElseIf bLineChart Then
        Set rpt = rptApp.OpenReport(RptPath & "\EmployeeProductivityReportLineChart.rpt")
        rpt.FormulaFields.GetItemByName("DateRange").Text = "'" & Format(DtFrom, "dd-MMM-yyyy") & " to " & Format(DtTo, "dd-MMM-yyyy") & "'"
        rpt.FormulaFields.GetItemByName("ReportType").Text = "'" & IIf(bOverTime, "Overtime", "Regular") & "'"
        strSelection = "{VEmpProductivityReport_Production.DTOnly}=#" & DtFrom & "# TO #" & DtTo & "# AND {VEmpProductivityReport_Production.OverTime}=" & bOverTime
    ElseIf bOverTime Then
        Set rpt = rptApp.OpenReport(RptPath & "\EmployeeProductivityReport.rpt")
        rpt.FormulaFields.GetItemByName("SubHeading").Text = "'" & Format(DtFrom, "dd-MMM-yyyy") & " to " & Format(DtTo, "dd-MMM-yyyy") & "'"
        rpt.FormulaFields.GetItemByName("ReportType").Text = "'" & IIf(bOverTime, "Overtime", "Regular") & "'"
        rpt.FormulaFields.GetItemByName("OverTime").Text = bOverTime
        strSelection = "{VEmpProductivityReport.DTOnly}=#" & DtFrom & "# TO #" & DtTo & "# AND {VEmpProductivityReport.OverTime}=" & bOverTime
        If strEmpID <> "0" Then
            strSelection = strSelection & " AND {VEmpProductivityReport.EmpID}='" & strEmpID & "'"
        End If
    Else
        Dim rptApp_Temp As New CRAXDDRT.Application
        Set rpt = rptApp_Temp.OpenReport(RptPath & "\EmployeeProductivityReportMerged.rpt")
        rpt.FormulaFields.GetItemByName("SubHeading").Text = "'" & Format(DtFrom, "dd-MMM-yyyy") & " to " & Format(DtTo, "dd-MMM-yyyy") & "'"
        rpt.FormulaFields.GetItemByName("ReportType").Text = "'" & IIf(bOverTime, "Overtime", "Regular") & "'"
        
        strSelection = "{VEmpProductivityReport.DTOnly}=#" & DtFrom & "# TO #" & DtTo & "#"
        If strEmpID <> "0" Then
            strSelection = strSelection & " AND {VEmpProductivityReport.EmpID}='" & strEmpID & "'"
        End If
        
        If strDeptID <> "0" Then
            strSelection = strSelection & " AND {VEmpProductivityReport.DeptID}='" & strDeptID & "'"
        End If
        
        If strProcesses <> "0" Then
            strSelection = strSelection & " AND {VEmpProductivityReport.ProcessID} IN[" & Replace(strProcesses, "'", "") & "]"
            strSubSelectionRegular = strSubSelectionRegular & " AND {VEmpProductivityReport.ProcessID} IN[" & Replace(strProcesses, "'", "") & "]"
            strSubSelectionOT = strSubSelectionRegular & " AND {VEmpProductivityReport.ProcessID} IN[" & Replace(strProcesses, "'", "") & "]"
        End If
         'AND {VEmpProductivityReport.OverTime}=" & bOverTime
        Dim rptSub1 As CRAXDDRT.Report, rptSub2 As CRAXDDRT.Report
        Set rptSub1 = rpt.OpenSubreport("Regular")
        rptSub1.RecordSelectionFormula = strSubSelectionRegular 'strSelection & " AND {VEmpProductivityReport.OverTime}=FALSE"
        
        Set rptSub2 = rpt.OpenSubreport("OverTime")
        rptSub2.RecordSelectionFormula = strSubSelectionOT
         
    End If
    rpt.FormulaFields.GetItemByName("DTFrom").Text = "#" & DtFrom & "#"
    rpt.FormulaFields.GetItemByName("DTTo").Text = "#" & DtTo & "#"
    
'    If strProcesses = "0" Then
'        rpt.FormulaFields.GetItemByName("DTTo").Text = "#" & DtTo & "#"
'    Else
'    End If
    
    rpt.FormulaFields.GetItemByName("DTTo").Text = "#" & DtTo & "#"
    Dim frm As New frmPrevRpt
    frm.ShowReport strSelection, rpt
    
End Sub

Public Function DateLocked(myDT As Date, bBlockFutureDates As Boolean, Optional bMachineLocked As Boolean = False) As Boolean
     
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim bDateLocked As Boolean
    
    With rs
        .Open "SELECT COUNT(*) FROM MonthlySalaries WHERE Month(DT)=" & Month(myDT) & " AND YEAR(DT)=" & year(myDT), con, adOpenForwardOnly, adLockReadOnly
        If .Fields(0).Value = 0 Then
            bDateLocked = False
        Else
            bDateLocked = True
        End If
        .Close
        If bDateLocked = False Then
            .Open "SELECT DateOpen,MachineDataLocked FROM PayrollDateLocks WHERE DT='" & myDT & "'", con, adOpenForwardOnly, adLockReadOnly
            If .EOF Then
                bDateLocked = False
            Else
                If !DateOpen Then
                    If bMachineLocked Then  'If Machine Lock is Required
                        If !MachineDataLocked = True Then
                            bDateLocked = True
                        Else
                            bDateLocked = False
                        End If
                    Else
                        bDateLocked = False
                    End If
                    
                Else
                    bDateLocked = True
                End If
            End If
            .Close
        End If
    End With
    Set rs = Nothing
     
    If bBlockFutureDates Then
        If myDT > Date Then
            bDateLocked = True
        End If
    End If
    
    DateLocked = bDateLocked
    
    Exit Function
err:
    MsgBox err.Description
    
End Function

Public Sub PrintEmpListGuarantorWise()
    'PrintManualAttendance
    Dim strGuarantor As String, strSelection As String
    
    Dim f As New frmGetSingleComboValue
    f.cmbCombo.AddItem "<All Guarantors>", "0"
    f.cmbCombo.AddVals con, "strFullName + ' ' + strFatherName", "GuarantorList_Imported", "intGuarantorID"
    
    If f.getComboValue(strGuarantor, "Guarantor", "Guarantors") = False Then Exit Sub
    
    If Val(strGuarantor) <> 0 Then
        strSelection = "{VEmp.intGuarantorId}=" & Val(strGuarantor)
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\EmployeeListGuarantorWise.rpt")
    
    Dim frm As New frmPrevRpt
    
    frm.ShowReport strSelection, rpt
    
End Sub

Public Sub PrintEmpListProbationEnd()
    '
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\EmployeeListProbation.rpt")
    Dim f As New frmPrevRpt
    Dim myDT As Date
    myDT = Date
    
    
    f.ShowReport "MONTH({@ProbationEndDate})=" & Month(myDT) & " AND YEAR({@ProbationEndDate})=" & year(myDT) & " AND {VEmp.Active}=TRUE", rpt
    
End Sub

Public Sub PrintEmpProductivityReportMonthly()

    Dim myDTFrom As Date, myDTTo As Date, strEmpID As String, strDeptID As String
    Dim f As New frmGetEmpIDAndDateRange 'frmgetdatesingle
    Load f
    'f.DTPicker1.CustomFormat = "MMM, yyyy"
    
    If f.getEmpIDAndDateRange(myDTFrom, myDTTo, strDeptID, strEmpID) = False Then Exit Sub
    
'    Dim myDTFrom As Date, myDTTo As Date
'    myDTFrom = DateSerial(year(myDT), Month(myDT), 1)
'    myDTTo = DateSerial(year(myDT), Month(myDT) + 1, 0)
     
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\EmployeeMonthlyProductivityReport.rpt")
    rpt.FormulaFields.GetItemByName("Range").Text = "'" & Format(myDTFrom, "dd-MMM-yyyy") & " to " & Format(myDTTo, "dd-MMM-yyyy") & "'"
     
    Dim strSelection As String
    strSelection = "{VEmpProductivityReport_Production.DT}=#" & myDTFrom & "# TO #" & myDTTo & "#"
    
    If strEmpID <> "0" Then
        strSelection = strSelection & " AND {VEmpProductivityReport_Production.EmpID}='" & strEmpID & "'"
    ElseIf strDeptID <> "0" Then
        strSelection = strSelection & " AND {VEmpProductivityReport_Production.DeptiD}='" & strDeptID & "'"
    End If
    
    Dim frm As New frmPrevRpt
    frm.ShowReport strSelection, rpt
    
End Sub

Public Sub PrintManualAttendance()

    Dim myDTFrom As Date, myDTTo As Date, strEmpID As String, strDeptID As String
    Dim f As New frmGetEmpIDAndDateRange
    Load f
    
    If f.getEmpIDAndDateRange(myDTFrom, myDTTo, strDeptID, strEmpID) = False Then Exit Sub
     
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\ManualAttendanceDetail.rpt")
    rpt.FormulaFields.GetItemByName("Range").Text = "'" & Format(myDTFrom, "dd-MMM-yyyy") & " to " & Format(myDTTo, "dd-MMM-yyyy") & "'"
     
    Dim strSelection As String
    strSelection = "DATEDIFF('d',{EmpTimes.DT},{EmpTimes.EntryDTOnly})>1"
    strSelection = strSelection & " AND {EmpTimes.DT}=#" & myDTFrom & "# TO #" & myDTTo & "#"
    
    If strEmpID <> "0" Then
        strSelection = strSelection & " AND {EmpTimes.EmpID}='" & strEmpID & "'"
    ElseIf strDeptID <> "0" Then
        strSelection = strSelection & " AND {EmpTimes.DeptiD}='" & strDeptID & "'"
    End If
    
    Dim frm As New frmPrevRpt
    frm.ShowReport strSelection, rpt
    
End Sub


Public Sub PrintSecuirtyGuardsTimings()

    Dim myDTFrom As Date, myDTTo As Date, strEmpID As String, strDeptID As String
    Dim f As New frmGetDateDouble, strSelection As String
    Load f
    
    If f.getDate(myDTFrom, myDTTo) = False Then Exit Sub
     
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\Employee_Timings_All.rpt")
    rpt.FormulaFields.GetItemByName("Range").Text = "'" & Format(myDTFrom, "dd-MMM-yyyy") & " to " & Format(myDTTo, "dd-MMM-yyyy") & "'"
     
    
    strSelection = "{VEmp.deptid}='25' AND {EmployeeTimings_All.DT}=#" & myDTFrom & "# TO #" & myDTTo & "#"
    
    
    Dim frm As New frmPrevRpt
    frm.ShowReport strSelection, rpt
    
End Sub


Public Sub PrintEmpYearlyDaysReport()
    'PrintEmployeeTargets
    Dim frm As New frmGetDateSingle
    Dim strYear As String
    Dim DTStart As Date, DTFinish As Date, iYear As Integer, myDT As Date
    Load frm
    frm.DTPicker1.CustomFormat = "yyyy"
    If frm.getDate(myDT) = False Then Exit Sub
    iYear = year(myDT) 'cmbFinYear.ItemData(cmbFinYear.ListIndex)
    DTStart = DateSerial(iYear, 1, 1)
    DTFinish = DateSerial(iYear, 12, 31)
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\EmpYearlyDaysReport.rpt")
   
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DTStart
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DTFinish
    
    rpt.FormulaFields.GetItemByName("Year").Text = "'" & iYear & "'"
    
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, , False
    
End Sub

Public Sub PrintEmployeeTargets()
    '
    Dim frm As New frmGetEmpIDAndDateRange
    Dim myDTFrom As Date, myDTTo As Date
    Load frm
    Dim strDeptID As String, strEmpID As String
    If frm.getEmpIDAndDateRange(myDTFrom, myDTTo, strDeptID, strEmpID) = False Then Exit Sub
     
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\Employees_Targets_Performance.rpt")
   
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue myDTFrom
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue myDTTo
     
    rpt.EnableParameterPrompting = False
    Dim strSelection As String
    If strEmpID <> "0" Then
        strSelection = "{Employees_Targets_Report_SP;1.EmpID}='" & strEmpID & "'"
    End If
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt, "", False
    
End Sub

