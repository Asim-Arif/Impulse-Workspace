Attribute VB_Name = "mod_Rpts"
Public Sub ShowRMVendListrpt()
    
    On Error GoTo err
    Screen.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report  ' New rptRMVenderList
    Set rpt = rptApp.OpenReport(RptPath & "\rptRMVenderList.rpt") '
    Dim f As New frmPrevRpt
    f.ShowReport " ", rpt
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Public Sub ShowHRrpt()
    
    On Error GoTo err
    Screen.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report 'New rptHRList
    Set rpt = rptApp.OpenReport(RptPath & "\rptHRList.rpt") '
    Dim f As New frmPrevRpt
    f.ShowReport " ", rpt
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Public Sub ShowAssHRrpt(Optional SqlQuery As String)
    
    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    
    If SqlQuery = "" Then
        SqlQuery = " Select * From VHicoResourceStatus"
    End If
    
    Dim rpt As CRAXDDRT.Report 'New rptAssHRList
    Set rpt = rptApp.OpenReport(RptPath & "\rptAssHRList.rpt") '
    Dim f As New frmPrevRpt
    
    rpt.SQLQueryString = SqlQuery
    
    
    f.ShowReport " ", rpt
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub
Public Sub ShowVendListrpt()
    
    On Error GoTo err
    Screen.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report 'New rptVenderList
    
    Dim f As New frmPrevRpt
    f.ShowReport " ", rpt
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Public Sub ShowRMListrpt(Optional strSelection As String = "")
    
    On Error GoTo err
    Screen.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report 'New rptRawMaterialList
    Set rpt = rptApp.OpenReport(RptPath & "\RMList.rpt") '
    Dim f As New frmPrevRpt
    
    f.ShowReport strSelection, rpt
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub
Public Sub ShowVendCardsrpt()
    
    On Error GoTo err
    Screen.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report 'New rptVendCards
    Set rpt = rptApp.OpenReport(RptPath & "\rptVendCards.rpt") '
    Dim FFDefs As FormulaFieldDefinitions
    Dim FFDef As CRAXDDRT.FormulaFieldDefinition
    Set FFDefs = rpt.FormulaFields
    
    For Each FFDef In FFDefs
        If FFDef.Name = "{@Valid Till}" Then
            FFDef.Text = "#" & Format(DateAdd("yyyy", 1, ServerDate), "dd-MMM-yyyy") & "#"
            Exit For
        End If
    Next
    
    Dim f As New frmPrevRpt
    
    f.ShowReport " ", rpt
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub
Public Sub ShowRMVendCardsrpt()
    
    On Error GoTo err
    Screen.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report 'New rptRMVendCards
    Set rpt = rptApp.OpenReport(RptPath & "\rptRMVendCards.rpt") '
    Dim FFDefs As FormulaFieldDefinitions
    Dim FFDef As CRAXDDRT.FormulaFieldDefinition
    Set FFDefs = rpt.FormulaFields
    
    For Each FFDef In FFDefs
        If FFDef.Name = "{@Valid Till}" Then
            FFDef.Text = "#" & Format(DateAdd("yyyy", 1, ServerDate), "dd-MMM-yyyy") & "#"
            Exit For
        End If
    Next
    
    Dim f As New frmPrevRpt
    
    f.ShowReport " ", rpt
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub


Public Sub PrintMakerPostedBill(strVchrNo As String, Optional Bold As Boolean, Optional lMPB_EntryID As Long = 0)

    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    
    If Bold Then
        Set rpt = rptApp.OpenReport(RptPath & "\MakerPostedBills.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\MakerPostedBillsRcvWise.rpt")
    End If
    
    Dim f As New frmPrevRpt
    Dim strSelection As String
    If lMPB_EntryID > 0 Then
        strSelection = "{MakerPostedBills.EntryID}=" & lMPB_EntryID
    Else
        strSelection = "{MakerPostedBills.VchrNo}='" & strVchrNo & "'"
    End If
    f.ShowReport strSelection, rpt, , False
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Public Sub PrintVenderPostedBill(strVchrNo As String)

    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\VenderPostedBill.rpt")
    
    Dim f As New frmPrevRpt
    f.ShowReport "{VenderPostedBills.VchrNo}='" & strVchrNo & "'", rpt
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Public Sub PrintMakerPostedBillTecno(lEntryID As Long)

    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\MakerPostedBillTecno.rpt")
    
'    Dim rptSub As CRAXDDRT.Report
'    Set rptSub = rpt.OpenSubreport("MakerUnPostedBills.rpt")
    Dim lMakerID As Long
    lMakerID = GetSingleLongValue("VendID", "MakerPostedBills", " WHERE VchrNo='" & strVchrNo & "'")
'    rptSub.ParameterFields.GetItemByName("@MakerID").SetCurrentValue lMakerID
'    rptSub.FormulaFields.GetItemByName("ShowMaker").Text = False
'    rptSub.EnableParameterPrompting = False
    'rptSub.DiscardSavedData
    Dim f As New frmPrevRpt
    f.ShowReport "{MakerPostedBills.EntryID}=" & lEntryID, rpt, , False
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Public Sub PrintMakerPaymentSheet()

    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\MakersPaymentSheet.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, "", False
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Public Sub PrintEmpSalaryLedger()
    'PrintMonthlyAttendanceStatus
    Dim frmGet As New frmGetEmpIDAndDateRange, myDTFrom As Date, myDTTo As Date, strDeptID As String, strEmpID As String
    If frmGet.getEmpIDAndDateRange(myDTFrom, myDTTo, strDeptID, strEmpID) = False Then
        Exit Sub
    End If
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\EmpSalaryLedger.rpt")
    Dim f As New frmPrevRpt
    
    Dim strCondition As String
    strCondition = "{MonthlySalaries.DT}=#" & myDTFrom & "# TO #" & myDTTo & "#"
    If strEmpID <> "0" Then
        strCondition = strCondition & " AND {VEmp.EmpID}='" & strEmpID & "'"
    ElseIf strDeptID <> "0" Then
        strCondition = strCondition & " AND {VEmp.DeptID}='" & strDeptID & "'"
    End If
    
    f.ShowReport strCondition, rpt
    
End Sub

Public Sub PrintMonthlyRejectionReport()

    Dim myDT As Date, frm As New frmGetDateSingle
    If frm.getDate(myDT) = False Then Exit Sub
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\MonthlyRejectionReportLotwise.rpt")
    rpt.ParameterFields.GetItemByName("@DT").SetCurrentValue myDT
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, "", False
    
End Sub

Public Sub PrintMonthlyAttendanceStatus(Optional bYearly As Boolean = False)
    '
    Dim frmGet As New frmGetEmpIDAndDateRange, myDTFrom As Date, myDTTo As Date, strDeptID As String, strEmpID As String
    If frmGet.getEmpIDAndDateRange(myDTFrom, myDTTo, strDeptID, strEmpID) = False Then
        Exit Sub
    End If
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\MonthlyAttendanceStatus.rpt")
    Dim f As New frmPrevRpt
    
    If bYearly = False Then
        myDTFrom = DateSerial(year(myDTFrom), Month(myDTFrom), 1)
        myDTTo = DateSerial(year(myDTFrom), Month(myDTFrom) + 1, 0)
    End If
    
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue myDTFrom
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue myDTTo
    rpt.ParameterFields.GetItemByName("@DeptID").SetCurrentValue strDeptID
    
    rpt.FormulaFields.GetItemByName("DateRange").Text = "'" & Format(myDTFrom, "dd-MMM-yyyy") & " to " & Format(myDTTo, "dd-MMM-yyyy") & "'"
    rpt.EnableParameterPrompting = False
    
    f.ShowReport "", rpt, , False
    
End Sub

Public Sub PrintOrderBalanceReport(strOrderNo As String)


    On Error GoTo err
    
    
    Dim rpt As CRAXDDRT.Report, f As New frmPrevRpt
    Set rpt = rptApp.OpenReport(RptPath & "\rptPartiallyShippedOrders.rpt")
    f.ShowReport "{FCustomerOrders.OrderNo}='" & strOrderNo & "'", rpt
    
    
    Exit Sub
err:
    
    MsgBox err.Description
    
End Sub


Public Sub PrintOrderPlanningReportR1(strOrderNo As String, Optional bR1 As Boolean = True)

    On Error GoTo err
    'Prepare Data for the Report...
    Dim cmd As New ADODB.Command
    With cmd
        Set .ActiveConnection = con_ServerSide
        .CommandType = adCmdStoredProc
        .CommandTimeout = 300
        .CommandText = "PreparePlanningReport_SP"
        .Parameters("@OrderNo").Value = strOrderNo
        .Parameters("@R1").Value = bR1
        .Execute
    End With
     
    Dim rpt As CRAXDDRT.Report
    If bR1 Then
        Set rpt = rptApp.OpenReport(RptPath & "\OrderPlanningR1.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\OrderPlanningR2.rpt")
    End If
    
    rpt.FormulaFields.GetItemByName("OrderWise").Text = True
    rpt.FormulaFields.GetItemByName("OrderNo").Text = "'" & strOrderNo & "'"
    
    Dim f As New frmPrevRpt
    f.ShowReport "{FCustomerOrders.OrderNo}='" & strOrderNo & "'", rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Public Sub PrintOrderPlanningReportR2(strOrderNo As String)

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\OrderPlanningR2.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{FCustomerOrders.OrderNo}='" & strOrderNo & "'", rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Public Sub PrintContractReview(strOrderNo As String)
    
    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ContractReview.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{FCustomerOrders.OrderNo}='" & strOrderNo & "'", rpt
    
    Exit Sub
err:
    MsgBox err.Description

End Sub

Public Sub PrintLoanBalance(Optional bInActive As Boolean = False, Optional bDateRange As Boolean)
    'PrintDailyPackingReport
    Dim SvrDt As Date
    If bDateRange Then
        Dim frmDT As New frmGetDateSingle
        Load frmDT
        If frmDT.getDate(SvrDt) = False Then
            Exit Sub
        End If
    Else
        SvrDt = GetServerDate
    End If
    
     
    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    If bInActive Then
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpLoanBalanceTerminated.rpt")
    Else
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpLoanBalance.rpt")
    End If
    
    rpt.ParameterFields.GetItemByName("@DT").SetCurrentValue SvrDt
    rpt.ParameterFields.GetItemByName("@InActive").SetCurrentValue bInActive
    
    rpt.EnableParameterPrompting = False
    
    f.ShowReport "", rpt, , False
   
End Sub

Public Sub PrintMasterPOStatus(strMasterPONo As String)

    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\MasterPOStatus.rpt")
    Dim f As New frmPrevRpt
    
    Dim strSelection As String

    strSelection = "{VendIssued.MasterPONo}='" & strMasterPONo & "'"
    
    f.ShowReport strSelection, rpt
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Public Sub PrintDailyPackingReport()
    'PrintPlanningReport
    Dim myDTFrom1 As Date, myDTTo1 As Date
    Dim myDTFrom2 As Date, myDTTo2 As Date
    Dim strCustCode As String, strItemGroup As String
    
    Dim frmDT As New frmGetCustomerANDItemGroupsANDDoubleDateRange
    Load frmDT
    If frmDT.getData(strCustCode, strItemGroup, myDTFrom1, myDTTo1, myDTFrom2, myDTTo2) = False Then Exit Sub
        
    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\DailyPackingReport.rpt")
    
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue myDTFrom1
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue myDTTo1
    
    rpt.ParameterFields.GetItemByName("@CustCode").SetCurrentValue IIf(strCustCode = "0", "", strCustCode)
    rpt.ParameterFields.GetItemByName("@ItemGroups").SetCurrentValue IIf(strItemGroup = "0", "", strItemGroup)
    
    rpt.ParameterFields.GetItemByName("@DTFrom2").SetCurrentValue myDTFrom2
    rpt.ParameterFields.GetItemByName("@DTTo2").SetCurrentValue myDTTo2
    
    rpt.EnableParameterPrompting = False
    
    f.ShowReport "", rpt, , False
   
End Sub

Public Sub PrintPlanningReport()
    '
    Dim myDTFrom As Date, myDTTo As Date
    
    Dim strCustCode As String, strItemGroup As String
    
    Dim frmDT As New frmGetItemGroupsANDDateRange
    Load frmDT
    If frmDT.getData(strItemGroup, myDTFrom, myDTTo) = False Then Exit Sub
        
    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\PlanningReport.rpt")
    
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue myDTFrom
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue myDTTo
    
    rpt.ParameterFields.GetItemByName("@ItemGroups").SetCurrentValue IIf(strItemGroup = "0", "", strItemGroup)
     
    rpt.EnableParameterPrompting = False
    
    f.ShowReport "", rpt, , False
   
End Sub

Public Sub PrintMonthlyProductionTimeLineReport()
    '
    Dim myDT As Date
    Dim frmDT As New frmGetDateSingle
    Load frmDT
    frmDT.DTPicker1.CustomFormat = "MMM, yyyy"
    If frmDT.getDate(myDT) = False Then Exit Sub
         
    Dim myDTFrom As Date, myDTTo As Date
    myDTFrom = DateSerial(year(myDT), Month(myDT), 1)
    myDTTo = DateSerial(year(myDT), Month(myDT) + 1, 0)
    
    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ProductionTimeLineReport_Monthly.rpt")
     
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue myDTFrom
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue myDTTo
    
    rpt.EnableParameterPrompting = False
    
    f.ShowReport "", rpt, , False
   
End Sub


Public Sub PrintYearlyDemandReport(Optional bItemGroupWise As Boolean = False, Optional bAmtWise As Boolean = False)
    'PrintYearlyDemandSupply
    Dim myDT As Date
    Dim f As New frmGetDateSingle
    Load f
    f.DTPicker1.CustomFormat = "yyyy"
    If f.getDate(myDT) = False Then Exit Sub
        
    Dim rpt As CRAXDDRT.Report
    If bItemGroupWise Then
        Set rpt = rptApp.OpenReport(RptPath & "\YearlyMainGroupwiseDemandReport.rpt")
    Else
        If bAmtWise Then
            Set rpt = rptApp.OpenReport(RptPath & "\YearlyCustomerwiseDemandReportAmtWise.rpt")
        Else
            Set rpt = rptApp.OpenReport(RptPath & "\YearlyCustomerwiseDemandReport.rpt")
        End If
    End If
    
    rpt.ParameterFields.GetItemByName("@DT").SetCurrentValue myDT
    rpt.EnableParameterPrompting = False
    rpt.FormulaFields.GetItemByName("Year").Text = "'" & Format(myDT, "yyyy") & "'"
    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, "", False
     
End Sub


Public Sub ItemWiseDemandSupply(Optional bCustomerwise As Boolean = False, Optional iReportType As Integer = 0, Optional bWithoutValue As Boolean = False)
    'PrintPeriodWiseDemand
    Dim myDTFrom As Date, myDTTo As Date
    Dim f As New frmGetDateDouble
    Load f
    
    If f.getDate(myDTFrom, myDTTo) = False Then Exit Sub
        
    Dim rpt As CRAXDDRT.Report
        
    If bCustomerwise Then
        'Set rpt = rptApp.OpenReport(RptPath & "\CustomerYearwiseDemandSupply.rpt")
        If iReportType = 3 Then
            Set rpt = rptApp.OpenReport(RptPath & "\TotalOrderValueShare.rpt")
        ElseIf iReportType = 6 Then
            Set rpt = rptApp.OpenReport(RptPath & "\TotalOrderValueShare_Demand.rpt")
        Else
            Set rpt = rptApp.OpenReport(RptPath & "\CustomerYearwiseDemands.rpt")
        End If
        
    Else
        If iReportType = 4 Then
            Set rpt = rptApp.OpenReport(RptPath & "\ExportDeliveryPerformanceReport.rpt")
        ElseIf iReportType = 5 Then
            Set rpt = rptApp.OpenReport(RptPath & "\ExportDeliveryPerformanceDetailReport.rpt")
            rpt.FormulaFields.GetItemByName("HideValue").Text = bWithoutValue
        Else
            Set rpt = rptApp.OpenReport(RptPath & "\ItemwiseDemandSupply.rpt")
            If iReportType = 0 Then         'Show upto Itemlevel.
                rpt.FormulaFields.GetItemByName("ShowItems").Text = True
                rpt.FormulaFields.GetItemByName("ShowSubGroup").Text = True
            ElseIf iReportType = 1 Then         'Show upto Subgruop.
                rpt.FormulaFields.GetItemByName("ShowItems").Text = False
                rpt.FormulaFields.GetItemByName("ShowSubGroup").Text = True
            ElseIf iReportType = 2 Then
                rpt.FormulaFields.GetItemByName("ShowItems").Text = False
                rpt.FormulaFields.GetItemByName("ShowSubGroup").Text = False
            End If
        End If
    End If
    
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue myDTFrom
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue myDTTo
    
    rpt.EnableParameterPrompting = False
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, "", False
     
End Sub

Public Sub PrintPeriodWiseDemand()
    '
    Dim DTStart As Date, DTFinish As Date
    Dim strCustCode As String, strGroups As String, strYear As String
    Dim f As New frmGetCustomerANDItemGroupsANDFinYear
    Load f
    
    If f.getData(strCustCode, strGroups, strYear) = False Then Exit Sub
     
    Dim iYear As Integer
    iYear = Val(strYear) 'cmbFinYear.ItemData(cmbFinYear.ListIndex)
    DTStart = DateSerial(iYear, 7, 1)
    DTFinish = DateSerial(iYear + 1, 6, 30)
    
    Dim rpt As CRAXDDRT.Report
         
    Set rpt = rptApp.OpenReport(RptPath & "\DemandsReceivedReport_FourMonthWise_SP.rpt")
    
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DTStart
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DTFinish
    rpt.ParameterFields.GetItemByName("@CustCode").SetCurrentValue IIf(strCustCode = "0", "", strCustCode)
    rpt.ParameterFields.GetItemByName("@ItemGroup").SetCurrentValue IIf(strGroups = "0", "", strGroups)
     
    rpt.FormulaFields.GetItemByName("Period").Text = "'" & year(DTStart) & "-" & year(DTFinish) & "'"
    rpt.EnableParameterPrompting = False
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, "", False
     
End Sub


Public Sub PrintDeadStockReport()

    Dim f As New frmGetSingleComboANDNumericValue, strStore As String, dDays As Double
    Load f
    
    
    f.cmbCombo.Functionality = 1
    f.cmbCombo.AddItem "<All Stores>", "0"
    f.cmbCombo.AddVals con, "StoreName", "Stores", "EntryID"
    f.cmbCombo.ID = "0"
    
    If f.getComboValue(strStore, dDays, "Stores", "Days", "Dead Stock") = False Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\DeadStockReport.rpt")
    rpt.EnableParameterPrompting = False
    rpt.ParameterFields.GetItemByName("@Stores").SetCurrentValue IIf(strStore = "0", "", strStore)
    rpt.ParameterFields.GetItemByName("@iDays").SetCurrentValue CLng(dDays)
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, "", False
    
End Sub

Public Sub PrintStockOutReport(strOrderNo As String)
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\StockOutReport.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{FCustomerOrders.OrderNo}='" & strOrderNo & "'", rpt
End Sub

Public Sub PrintYearlyDemandSupply(Optional bGraph As Boolean = False, Optional bAmtWise As Boolean = False)
    '
    Dim myDT As Date
    Dim f As Form
    
    
    Dim myDTFrom As Date, myDTTo As Date
    Dim strYear As String
    Dim iYear As Integer
       
    Dim rpt As CRAXDDRT.Report
    If bGraph Then
        If bAmtWise = False Then
            Set rpt = rptApp.OpenReport(RptPath & "\DemandAndSupplyComparisonGraph.rpt")
        Else
            Set rpt = rptApp.OpenReport(RptPath & "\DemandAndSupplyComparisonGraphAmtWise.rpt")
        End If
        Set f = New frmGetDateSingle
        
        f.DTPicker1.CustomFormat = "yyyy"
        Load f
        If f.getDate(myDT) = False Then Exit Sub
        iYear = year(myDT) 'cmbFinYear.ItemData(cmbFinYear.ListIndex)
        myDTFrom = DateSerial(iYear, 1, 1)
        myDTTo = DateSerial(iYear, 12, 31)
        
        strYear = "'" & iYear & "'"
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\DemandAndSupplyReport.rpt")
        Set f = New frmGetDateSingle
        Load f
        f.DTPicker1.CustomFormat = "yyyy"
        If f.getDate(myDT) = False Then Exit Sub
        
        myDTFrom = DateSerial(year(myDT), 1, 1)
        myDTTo = DateSerial(year(myDT), 12, 31)
        strYear = "'" & Format(myDT, "yyyy") & "'"
    End If
    
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue myDTFrom
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue myDTTo
    
    rpt.EnableParameterPrompting = False
    
    rpt.FormulaFields.GetItemByName("DateRange").Text = strYear
    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, "", False
     
End Sub

Public Sub CustomerPurchaseHistory_Orderwise()

    Dim f As New frmGetCustomerANDDateRange
    Dim strCustCode As String, myDTFrom As Date, myDTTo As Date
    
    Load f
    f.cmbCust.Functionality = 0
    If f.getData(strCustCode, myDTFrom, myDTTo) = False Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\CustomerPurchaseHistory_Orders.rpt")
    rpt.FormulaFields.GetItemByName("DateRange").Text = "'" & Format(myDTFrom, "dd-MMM-yyyy") & " to " & Format(myDTTo, "dd-MMM-yyyy") & "'"
    
    Dim strSelection As String
    strSelection = "{FCustomerOrders.DT}=#" & myDTFrom & "# TO #" & myDTTo & "#"
    If strCustCode <> "0" Then
        strSelection = strSelection & " AND {FCustomerOrders.CustCode}='" & strCustCode & "'"
    End If
    
    Dim frm As New frmPrevRpt
    frm.ShowReport strSelection, rpt
    
End Sub

Public Sub PrintYearlyDemandReport_ItemWise(Optional bItemGroupWise As Boolean = False, Optional bAmtWise As Boolean = False, Optional bYearWise As Boolean = False, Optional bX2Only As Boolean = False)
    'PrintYearlyDemandSupply
    Dim myDT As Date
    Dim f As New frmGetDateSingle
    Load f
    f.DTPicker1.CustomFormat = "yyyy"
    If f.getDate(myDT) = False Then Exit Sub
        
    Dim rpt As CRAXDDRT.Report
        
    If bYearWise Then
        Set rpt = rptApp.OpenReport(RptPath & "\YearlyItemwiseDemandReport_Yearwise.rpt")
    Else
        If bAmtWise Then
            Set rpt = rptApp.OpenReport(RptPath & "\YearlyCustomerwiseDemandReportAmtWise.rpt")
        Else
            Set rpt = rptApp.OpenReport(RptPath & "\YearlyItemwiseDemandReport.rpt")
        End If
    End If
    
    
    rpt.ParameterFields.GetItemByName("@DT").SetCurrentValue myDT
    rpt.ParameterFields.GetItemByName("@YearlyReport").SetCurrentValue bYearWise
        
    If bX2Only Then
        rpt.ParameterFields.GetItemByName("@X2Only").SetCurrentValue True
    Else
        rpt.ParameterFields.GetItemByName("@X2Only").SetCurrentValue False
    End If
    
    rpt.EnableParameterPrompting = False
    rpt.FormulaFields.GetItemByName("Year").Text = "'" & Format(myDT, "yyyy") & "'"
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, "", False
     
End Sub

Public Sub PrintCustomerWiseDemands_Yearwise()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\CustomerWiseDemands_Yearwise.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VFOrderList.OrderType}='Customer Order' AND {VFOrderList.Cancelled}=0 AND (NOT ({VFOrderList.CustCode} IN['Stock','Wrangler']))", rpt
    
End Sub

Public Sub PrintProductLineDemand(Optional bGroupwise As Boolean = False)
    'PrintProductLineSupply
    Dim frm As New frmGetDateDouble
    Dim myDTFrom As Date, myDTTo As Date
    Load frm
    If frm.getDate(myDTFrom, myDTTo) = False Then Exit Sub
    
    Dim strSelection As String
    strSelection = "{FCustomerOrders.DT}=#" & myDTFrom & "# TO #" & myDTTo & "#"
    strSelection = strSelection & " AND (NOT ({FCustomerOrders.CustCode} IN['Stock','Wrangler','SAL'])) "
    strSelection = strSelection & " AND {FCustomerOrders.OrderType}='Customer Order'"
    strSelection = strSelection & " AND ({FCustomerFinalOrders.Cancelled}=0 OR ISNULL({FCustomerFinalOrders.Cancelled}))"
     
    Dim f As New frmPrevRpt
    Dim rpt As CRAXDDRT.Report
    If bGroupwise Then
        Set rpt = rptApp.OpenReport(RptPath & "\ProductGroupwiseDemand.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\ProductLineDemand.rpt")
    End If
    rpt.FormulaFields.GetItemByName("DateRange").Text = "'" & Format(myDTFrom, "dd-MMM-yyyy") & " to " & Format(myDTTo, "dd-MMM-yyyy") & "'"
    
    f.ShowReport strSelection, rpt
    
End Sub

Public Sub PrintProductLineSupply(Optional bGroupwise As Boolean = False)
    '
    Dim frm As New frmGetDateDouble
    Dim myDTFrom As Date, myDTTo As Date
    Load frm
    If frm.getDate(myDTFrom, myDTTo) = False Then Exit Sub
    
    Dim strSelection As String
    strSelection = "{CustomInvoice.DT}=#" & myDTFrom & "# TO #" & myDTTo & "#"
    strSelection = strSelection & " AND {CustomInvoice.GatePassNo}<>''"
      
    Dim f As New frmPrevRpt
    Dim rpt As CRAXDDRT.Report
    If bGroupwise Then
        Set rpt = rptApp.OpenReport(RptPath & "\ProductGroupwiseSupply.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\ProductLineSupply.rpt")
    End If
    rpt.FormulaFields.GetItemByName("DateRange").Text = "'" & Format(myDTFrom, "dd-MMM-yyyy") & " to " & Format(myDTTo, "dd-MMM-yyyy") & "'"
    
    f.ShowReport strSelection, rpt
    
End Sub


Public Sub PrintAdvancePlanceReport()

    Dim frm As New frmGetItemGroupsANDAdditionalGroupsANDDateRange
    Load frm
    frm.DTFrom1.CustomFormat = "yyyy"
    frm.DTTo1.CustomFormat = "yyyy"
    
    Dim myDTFrom As Date, myDTTo As Date, strItemGroup As String, strAdditionalGroup As String
    If frm.getData(strItemGroup, strAdditionalGroup, myDTFrom, myDTTo) = False Then Exit Sub
    
    myDTFrom = DateSerial(year(myDTFrom), 1, 1)
    myDTTo = DateSerial(year(myDTTo), 12, 31)
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\YearlyItemwiseDemandReport_EachYear.rpt")
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue myDTFrom
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue myDTTo
    rpt.ParameterFields.GetItemByName("@ItemGroups").SetCurrentValue IIf(strItemGroup = "0", "", strItemGroup)
    rpt.ParameterFields.GetItemByName("@AdditionalGroups").SetCurrentValue IIf(strAdditionalGroup = "0", "", strAdditionalGroup)
    
    rpt.EnableParameterPrompting = False
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, , False
    
End Sub

Public Sub PrintSamplingStockReport()
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Items_Sampling_Stock_Detail.rpt")

    Dim frm As New frmPrevRpt
    frm.ShowReport "{VItemSFStock.NetQty}>0", rpt
End Sub

Public Sub Print_SF_Opening_Slip(lSOO_EntryID As Long)

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\SF_Opening_Slip.rpt")

    Dim frm As New frmPrevRpt
    frm.ShowReport "{VStockOrderOpening.EntryID}=" & lSOO_EntryID, rpt
    
End Sub

Public Sub Print_Packed_Opening_Slip(strRcvNo As String)

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Finished_Stock_Receiving_Slip.rpt")

    Dim frm As New frmPrevRpt
    frm.ShowReport "{VStockTransactions.EntryNo}='" & strRcvNo & "'", rpt
    
End Sub

Public Sub PrintHubWiseOrderStatusReport(strOrderNo As String)
    'PrintUnPlannedItems
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Customer_Order_Status_Hubwise.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VFOrderItems.OrderNo}='" & strOrderNo & "'", rpt
End Sub

Public Sub PrintUnPlannedItems(strOrderNo As String)
    'PrintRawMaterialPlanning
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Customer_Order_Unplanned_Items.rpt")
    Dim f As New frmPrevRpt
    Dim strSQL As String
    strSQL = "SELECT VFOrderItems.* FROM VFOrderItems WHERE OrderNo='" & strOrderNo & "' AND OrderNo+CompItemCode NOT IN(SELECT OrderNo+ItemCode FROM VendRcvdDetail) AND OrderNo+CompItemCode NOT IN(SELECT OrderNo+ItemCode FROM VendIssdDetail)"
    f.ShowReport "", rpt, strSQL
End Sub

Public Sub PrintVendorSummaryReport()
    'PrintMakerPayableReport
    Dim f As New frmGetDateSingle, myDT As Date
    Load f
    f.Label1.Caption = "Starting Month"
    f.DTPicker1.CustomFormat = "MMM, yyyy"
    If f.getDate(myDT) = False Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\VendorSummaryReport.rpt")
    rpt.ParameterFields.GetItemByName("@DT").SetCurrentValue myDT
    rpt.EnableParameterPrompting = False
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, , False
     
End Sub

Public Sub PrintMakerPayableReport()

    'PrintDailyCashANDBankReport
'    Dim f As New frmGetSingleComboValue, strAccNo As String
'    Load
'    If f.getComboValue(strAccNo, "Account", "Account") = False Then Exit Sub
            
    Dim f As New frmGetDateDoubleANDComboValue
    Load f
    
    f.lblCombo.Caption = "Catagory"
    
    f.cmb.AddItem "<All>", "0"
    'f.cmb.AddVals con, "AccTitle", "Accounts", "AccNo", "WHERE AccNo IN('15-005','15-006','15-007')"
    f.cmb.ID = "0"
    
    Dim myDTFrom As Date, myDTTo As Date, strCatagory As String
    
    If f.getDate(myDTFrom, myDTTo, strCatagory) = False Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\MakersPayableReport.rpt")
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue myDTFrom '#7/1/2013#
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue myDTTo
    rpt.ParameterFields.GetItemByName("@Catagory").SetCurrentValue strCatagory
    
    rpt.EnableParameterPrompting = False
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, , False
     
End Sub

Public Sub PrintQMSGraph(iType As Integer)

    Dim myDTFrom As Date, myDTTo As Date
    Dim f As New frmGetDateDouble
    Load f
    If f.getDate(myDTFrom, myDTTo) = False Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Dim strSelection As String
    
    If iType = 0 Then
        Set rpt = rptApp.OpenReport(RptPath & "\Customer_Complaints_Graph.rpt")
        strSelection = "{CustomerComplaints.DT}=#" & myDTFrom & "# TO #" & myDTTo & "#"
    ElseIf iType = 1 Then
        Set rpt = rptApp.OpenReport(RptPath & "\Customer_Complaints_ItemGroupwise_Graph.rpt")
        strSelection = "{CustomerComplaints.DT}=#" & myDTFrom & "# TO #" & myDTTo & "#"
    ElseIf iType = 2 Then
        Set rpt = rptApp.OpenReport(RptPath & "\Claims_Statistics_Graph.rpt")
        strSelection = "{VClaims_Statistics.DT}=#" & myDTFrom & "# TO #" & myDTTo & "#"
    End If
    rpt.FormulaFields.GetItemByName("DateRange").Text = "'" & Format(myDTFrom, "dd-MMM-yyyy") & " to " & Format(myDTTo, "dd-MMM-yyyy") & "'"
    Dim frm As New frmPrevRpt
    frm.ShowReport strSelection, rpt
    
End Sub


Public Sub PrintRunningLots_Hubwise()

    Dim frm As New frmGetDateDouble
    Load frm
    
    Dim myDTFrom As Date, myDTTo As Date
    If frm.getDate(myDTFrom, myDTTo) = False Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Running_Lots_Summary_Hubwise.rpt")
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue myDTFrom
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue myDTTo
    rpt.ParameterFields.GetItemByName("@OrderNo").SetCurrentValue "0"
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, "", False
        
End Sub


Public Sub Print_User_Activity_Report()

    Dim f As New frmGetDateDoubleANDComboValue
    Load f
    f.lblCombo.Caption = "User Name"
    f.cmb.AddVals con, "DISTINCT UserName", "Users", "UserName", ""
    
    Dim myDTFrom As Date, myDTTo As Date, strUserName As String
    
    If f.getDate(myDTFrom, myDTTo, strUserName) = False Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\User_Activity.rpt")
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue myDTFrom
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue myDTTo
    rpt.ParameterFields.GetItemByName("@UserName").SetCurrentValue strUserName
    rpt.EnableParameterPrompting = False
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, "", False
    
End Sub

Public Sub PrintRunningLots(Optional bWithValue As Boolean = False, Optional iReportType As Integer = 0)

    Dim rpt As CRAXDDRT.Report
    Dim strSelection As String
    If bWithValue Then
        If getDBPassword("Running_Lots_Value") = False Then Exit Sub
        Set rpt = rptApp.OpenReport(RptPath & "\RunningLots_Value.rpt")
    ElseIf iReportType = 1 Then 'Stock Order
        Set rpt = rptApp.OpenReport(RptPath & "\RunningLotsStock.rpt")
        strSelection = "{VRunningLots.OrderNo} IN['Stock-Order','Stock-OrderIss']"
    ElseIf iReportType = 2 Then 'Lots Packed
        Set rpt = rptApp.OpenReport(RptPath & "\LotsPacked.rpt")
        strSelection = "{VRunningLots.ProcessID} IN[163]"
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\RunningLots.rpt")
    End If
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
End Sub

Public Sub PrintVendorBalances()
    'PrintVendorBalances_LongTerm
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\VendorLoanBalance.rpt")
    rpt.ParameterFields.GetItemByName("@DT").SetCurrentValue Date
    rpt.ParameterFields.GetItemByName("@InActive").SetCurrentValue False
    
    rpt.EnableParameterPrompting = False
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, , False
End Sub

Public Sub PrintOpenMakerOrders()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Maker_Open_POs.rpt")

    
    Dim strSQL As String
    strSQL = "SELECT VVendIssued.* FROM VVendIssued LEFT OUTER JOIN VItemOpenPOs ON VVendIssued.DetailEntryID=VItemOpenPOs.VID_EntryID  WHERE  DetailEntryID IN(SELECT VID_EntryID FROM VItemOpenPOs) ORDER BY DT"
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, strSQL
    
End Sub

Public Sub PrintOpenVendorOrders()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Vendor_Open_POs.rpt")

    
    Dim strSelection As String
    strSelection = "{@qtybalance}>0 AND {VVendOrders.Final}=FALSE"
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
End Sub


Public Sub PrintOpenLots()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Open_Lots.rpt")

    
'    Dim strSQL As String
'    strSQL = "SELECT VVendIssued.* FROM VVendIssued LEFT OUTER JOIN VItemOpenPOs ON VVendIssued.DetailEntryID=VItemOpenPOs.VID_EntryID  WHERE  DetailEntryID IN(SELECT VID_EntryID FROM VItemOpenPOs) ORDER BY DT"
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt
    
End Sub

Public Sub PrintAppForm()

    Screen.MousePointer = vbHourglass
    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptAppForm.rpt")
    rpt.FormulaFields(1).Text = "'" & strCompany & "'"
    f.ShowReport " ", rpt
    Screen.MousePointer = vbDefault

End Sub

Public Sub PrintEmpDetails()
    
    Dim frm As New frmGetSingleComboValue, strEmpID As String
    Load frm
    frm.cmbCombo.AddVals con, "'{'+EmpID+'} '+Name", "VEmp", "EmpID", " ORDER BY VEmp.EmpID"
    frm.Label1.Caption = "Employee"
    If frm.getComboValue(strEmpID, "Employee", "Employees") = False Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpDetail.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{Employees.empid}='" & strEmpID & "'", rpt

End Sub

Public Sub PrintFinishStockBelowMinLevelReport()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\CompanyCatalog_MinLevel.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VItems.InActive}=FALSE AND {VItems.InHand}<{VItems.MinLevel}", rpt
    
End Sub

Public Sub PrintFinishStockAboveMinLevelReport()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\CompanyCatalog_MaxLevel.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VItems.InActive}=FALSE AND {VItems.InHand}>{VItems.MaxLevel} AND {VItems.MaxLevel}>0", rpt

End Sub


Public Sub PrintRawMaterialPlanning(strOrderNo As String)
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Customer_Order_RM.rpt")
    
    Dim strSelection As String
    strSelection = "{VrptOrders_ForProduction_Reports.OrderNo}='" & strOrderNo & "'"
    Dim f As New frmPrevRpt
    
    f.ShowReport strSelection, rpt
    
End Sub

