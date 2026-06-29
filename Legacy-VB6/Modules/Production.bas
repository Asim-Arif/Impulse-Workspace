Attribute VB_Name = "Production"
Public Sub ShowStockStatus(ProcID As Long)
    
'    Dim f As Form
'    If ProcID = 13 Then
'        Set f = New frmStockStatusFinal
'    Else
'        Set f = New frmStockStatus
'    End If
'
'    f.ShowStockStatus (ProcID)
    
End Sub
Public Sub ShowIssRcvReports(ProcID As Integer)

'    Dim f As New frmRptIssRcv
'    Load f
'    Call f.ShowForIss(False, ProcID)
    
End Sub

Public Sub ShowIssForm(ProcID As Long)
    
    Dim f As Form
    Set f = New frmVendIssItemsByItem
    Load f
    Call f.LockProc(ProcID)

End Sub

Public Sub ShowIssPOForm(ProcID As Long)
    
    Dim f As New frmVendNewPO
    Load f
    Call f.LockProc(ProcID)

End Sub
Public Sub ShowRcvPOForm(ProcID As Long)

    Dim f As New frmVendRcvPO
    Load f
    Call f.LockProc(ProcID)
    
End Sub
Public Sub ShowRcvDirectForm(ProcID As Long)

'    Dim f As New frmVendRcvItemsDirect1
'    Load f
'    Call f.LockProc(ProcID)
    
End Sub

Public Sub ShowRcvVForm(ProcID As Long, Optional bManual As Boolean = False)

    Dim f As Form 'New frmVendRcvItems
    
'    If ProcID = 1 Then
'        Set f = New frmVendRcvItemsDirect
'    ElseIf ProcID = 12 Then
'        Set f = New frmVendRcvItemsP
'    Else
'        Set f = New frmVendRcvItems
'    End If
    If bManual Then
        Set f = New frmVendRcvItemsManual
    Else
        Set f = New frmVendRcvItems_Lots
    End If
    Load f
    Call f.LockProc(ProcID)

End Sub
Public Sub ShowWastRcvForm(ProcID As Long)

'    Dim f As New frmVendRcvWastage
'    Load f
'    Call f.LockProc(ProcID)

End Sub

Public Sub ShowRcvSForm(ProcID As Long)

'    Dim f As New frmVendRcvItems1
'    Load f
'    Call f.LockProc(ProcID)

End Sub

Public Sub ShowLostRcvForm(ProcID As Long)

'    Dim f As New frmVendRcvLost
'    Load f
'    Call f.LockProc(ProcID)

End Sub

Public Sub PrintIssRcvRpts(strOriginal As String, Optional IssEntry As Boolean)

    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    
    Dim rs As New ADODB.Recordset
    
    Call StartTrans(con)
    'Insert Into PrintIssRcvRptsTemp.
    
    con.Execute "Delete From PrintIssRcvRptsTemp"
    con.Execute "Insert Into PrintIssRcvRptsTemp Select RecieptNo From PrintSession"
    
'    If IssEntry = False Then
'        con.Execute "INSERT INTO PrintIssRcvRpts Select CompItemDesc,DT,RecieptID,ItemCode,VenderName" & _
'         ",VendID1,ProcessID,Process,RcvdQty,Wastage,LostQty,UserID,AuthUserID,AuthDT,0 As IssEntry " & _
'         "From VVendRcvItemsrpt Where ReqAuth=0 AND RecieptID In(Select RecieptNo From PrintSession)"
'    Else
'        con.Execute "INSERT INTO PrintIssRcvRpts Select CompItemDesc,DT,RecieptID,ItemCode,VenderName," & _
'         "VendID1,ProcessID,Process,IssQty,0 As Wastage,0 As LostQty,UserID,AuthUserID,AuthDT,1 As IssEntry " & _
'         "From VVendIssItemsrpt Where ReqAuth=0 AND RecieptID In(Select RecieptNo From PrintSession)"
'    End If
    con.CommitTrans
    
    Dim rpt As CRAXDDRT.Report  'New rptVendIssDetail
    Set rpt = rptApp.OpenReport(RptPath & "\ProcessIssSlip.rpt") '
    
    Dim FFDefs As FormulaFieldDefinitions
    Dim FFDef As CRAXDDRT.FormulaFieldDefinition
    Set FFDefs = rpt.FormulaFields
    
    For Each FFDef In FFDefs
        If FFDef.Name = "{@ForOriginal}" Then
            FFDef.Text = "'" & strOriginal & "'"
        End If
    Next
        
    Load frmPrevRpt
    frmPrevRpt.ShowReport " ", rpt
    
    Screen.MousePointer = vbDefault

    
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Public Sub PrintMaterialConsumptionSheet(strOrderNo As String, Optional bArticlewise As Boolean = False)

    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    If bArticlewise Then
        Set rpt = rptApp.OpenReport(RptPath & "\MaterialConsumptionSheetArticlewise.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\MaterialConsumptionSheet.rpt")
    End If
    
    Dim f As New frmPrevRpt
    f.ShowReport "{FOrderItems.OrderNo}='" & strOrderNo & "'", rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Public Function getNextProcessID(strItemID As String, lProcessID As Long) As Long

    Dim lNextProcessID As Long
    Dim cmd As New ADODB.Command
    With cmd
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_GetNextProcID"
        .Parameters("@ItemCode").Value = strItemID
        .Parameters("@ProcID").Value = lProcessID
        .Parameters("@ReWorkLot").Value = 0
        .Execute
        lNextProcessID = Val(.Parameters("@NextProcID").Value & "")
    End With
    getNextProcessID = lNextProcessID
    
End Function

Public Function getPrevProcessID(strItemID As String, lProcessID As Long) As Long

    Dim lPrevProcessID As Long
    Dim cmd As New ADODB.Command
    With cmd
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_GetPrevProcID"
        .Parameters("@ItemCode").Value = strItemID
        .Parameters("@ProcID").Value = lProcessID
        .Execute
        lPrevProcessID = Val(.Parameters("@PrevProcID").Value & "")
    End With
    getPrevProcessID = lPrevProcessID
    
End Function

Public Function GetIssuanceIDDetail(myDT As Date) As String
    
    Dim lID As Long
    lID = GetSingleLongValue("Max(Cast(right(RecieptID,Len(RecieptID)-10) As Int))", "VMakerIssItems", " WHERE DT='" & myDT & "'")
    lID = lID + 1
    Dim strIssID As String
    strIssID = "ISU-" & Format(myDT, "ddMMyy") & lID
    
    GetIssuanceIDDetail = strIssID
    
End Function

Public Function getNextLotNo() As String

    Dim lRcvLotNo As Long, lIssLotNo As Long, lSOO_LotNo As Long, lRcv_LotNo As Long
    Dim lLotNo As Long
    Dim lLotList_LotNo As Long
    
    lLotList_LotNo = GetSingleLongValue("MAX(CAST(SUBSTRING(LotNo,3,CASE WHEN CHARINDEX('-',LotNo)>0 THEN CHARINDEX('-',LotNo)-3 ELSE LEN(LotNo) END) AS INT))", "Lots_List", " WHERE LEFT(LotNo,2)=" & Format(Date, "yy") & " AND LEN(LotNo)=7")
    
    lRcvLotNo = GetSingleLongValue("MAX(CAST(SUBSTRING(LotNo,7,CASE WHEN CHARINDEX('-',LotNo)>0 THEN CHARINDEX('-',LotNo)-3 ELSE LEN(LotNo) END) AS INT))", "VendRcvdDetail", " WHERE LEFT(LotNo,2)=" & Format(Date, "yy") & " AND LEN(LotNo)=7")
    lIssLotNo = GetSingleLongValue("MAX(CAST(SUBSTRING(LotNo,7,CASE WHEN CHARINDEX('-',LotNo)>0 THEN CHARINDEX('-',LotNo)-3 ELSE LEN(LotNo) END) AS INT))", "VendIssdDetail", " WHERE LEFT(LotNo,2)=" & Format(Date, "yy") & " AND LEN(LotNo)=7")
    lSOO_LotNo = GetSingleLongValue("MAX(CAST(SUBSTRING(LotNo_Manual,7,CASE WHEN CHARINDEX('-',LotNo_Manual)>0 THEN CHARINDEX('-',LotNo_Manual)-3 ELSE LEN(LotNo_Manual) END) AS INT))", "StockOrderOpening", " WHERE LEFT(LotNo_Manual,2)=" & Format(Date, "yy") & " AND LEN(LotNo_Manual)=7 AND ISNUMERIC(LotNo_Manual)=1")
    lRcv_LotNo = GetSingleLongValue("MAX(CAST(SUBSTRING(LotNo_Manual,7,CASE WHEN CHARINDEX('-',LotNo_Manual)>0 THEN CHARINDEX('-',LotNo_Manual)-3 ELSE LEN(LotNo_Manual) END) AS INT))", "RcvItemsSimpleDetail", " WHERE LEFT(LotNo_Manual,2)=" & Format(Date, "yy") & " AND LEN(LotNo_Manual)=7 AND ISNUMERIC(LotNo_Manual)=1")
    
'    lLotList_LotNo = GetSingleLongValue("MAX(CAST(SUBSTRING(LotNo,7,CASE WHEN CHARINDEX('-',LotNo)>0 THEN CHARINDEX('-',LotNo)-7 ELSE LEN(LotNo) END) AS INT))", "Lots_List", " WHERE LEFT(LotNo,6)=" & Format(Date, "DDMMyy") & " AND LEN(LotNo)>8")
'
'    lRcvLotNo = GetSingleLongValue("MAX(CAST(SUBSTRING(LotNo,7,CASE WHEN CHARINDEX('-',LotNo)>0 THEN CHARINDEX('-',LotNo)-7 ELSE LEN(LotNo) END) AS INT))", "VendRcvdDetail", " WHERE LEFT(LotNo,6)=" & Format(Date, "DDMMyy") & " AND LEN(LotNo)>8")
'    lIssLotNo = GetSingleLongValue("MAX(CAST(SUBSTRING(LotNo,7,CASE WHEN CHARINDEX('-',LotNo)>0 THEN CHARINDEX('-',LotNo)-7 ELSE LEN(LotNo) END) AS INT))", "VendIssdDetail", " WHERE LEFT(LotNo,6)=" & Format(Date, "DDMMyy") & " AND LEN(LotNo)>8")
'    lSOO_LotNo = GetSingleLongValue("MAX(CAST(SUBSTRING(LotNo_Manual,7,CASE WHEN CHARINDEX('-',LotNo_Manual)>0 THEN CHARINDEX('-',LotNo_Manual)-7 ELSE LEN(LotNo_Manual) END) AS INT))", "StockOrderOpening", " WHERE LEFT(LotNo_Manual,6)=" & Format(Date, "DDMMyy") & " AND LEN(LotNo_Manual)>8 AND ISNUMERIC(LotNo_Manual)=1")
'    lRcv_LotNo = GetSingleLongValue("MAX(CAST(SUBSTRING(LotNo_Manual,7,CASE WHEN CHARINDEX('-',LotNo_Manual)>0 THEN CHARINDEX('-',LotNo_Manual)-7 ELSE LEN(LotNo_Manual) END) AS INT))", "RcvItemsSimpleDetail", " WHERE LEFT(LotNo_Manual,6)=" & Format(Date, "DDMMyy") & " AND LEN(LotNo_Manual)>8 AND ISNUMERIC(LotNo_Manual)=1")
    
    If lLotList_LotNo = 0 And lRcvLotNo = 0 And lIssLotNo = 0 And lSOO_LotNo = 0 And lRcv_LotNo = 0 Then
        lLotNo = 53000
    Else
        If lRcvLotNo > lIssLotNo Then
            If lRcvLotNo > lsso_lotno Then
                lLotNo = lRcvLotNo
            Else
                lLotNo = lSOO_LotNo
            End If
        Else
            If lIssLotNo > lSOO_LotNo Then
                lLotNo = lIssLotNo
            Else
                lLotNo = lSOO_LotNo
            End If
        End If
        If lRcv_LotNo > lLotNo Then
            lLotNo = lRcv_LotNo
        End If
        If lLotList_LotNo > lLotNo Then
            lLotNo = lLotList_LotNo
        End If
        If Left(lLotNo, 6) = Format(Date, "DDMMyy") Then
            lLotNo = Mid(lLotNo, 7)
        End If
    End If
    lLotNo = lLotNo + 1
    
    getNextLotNo = Format(Date, "yy") & Format(lLotNo, "00000")
    'getNextLotNo = Format(Date, "DDMMyy") & Format(lLotNo, "000")
    
End Function

Public Sub PrintOrderPlanningReport(strOrderNo As String, Optional bQuotationPlanning As Boolean = False)
    'PrintSFLedger
    Dim rpt As CRAXDDRT.Report
    
    If bQuotationPlanning Then
        Set rpt = rptApp.OpenReport(RptPath & "\QuotationPlanningReport.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\OrderPlanningReport.rpt")
    End If
    
    rpt.ParameterFields.GetItemByName("@OrderNo").SetCurrentValue strOrderNo
    rpt.ParameterFields.GetItemByName("@FromItems").SetCurrentValue False
    rpt.ParameterFields.GetItemByName("@ItemCode").SetCurrentValue "0"
    rpt.ParameterFields.GetItemByName("@TempList_RefID").SetCurrentValue 0
    
    rpt.EnableParameterPrompting = False
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, , False
    
End Sub

Public Sub PrintSFLedger(strItemID As String)
    '
    Dim rpt As CRAXDDRT.Report
    
    Set rpt = rptApp.OpenReport(RptPath & "\SFLedger.rpt")
    
    rpt.ParameterFields.GetItemByName("@ItemID").SetCurrentValue strItemID
    
    rpt.EnableParameterPrompting = False
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, , False
    
End Sub

Public Sub PrintPTC(strLotNo As String, Optional bWithPrice As Boolean = False, Optional iReportType As Integer = 0)

    On Error GoTo err
    
    If bWithPrice Then
        If getDBPassword("PTC_Cost") = False Then Exit Sub
    End If
    
    Dim rpt As CRAXDDRT.Report
    Dim f As New frmPrevRpt, strSelection As String
    
    Dim bRepairLot As Boolean
    'strLotNo = LV.SelectedItem.ListSubItems(8).Tag
    'bRepairLot = GetSingleBooleanValue("ReWorkLot", "VendRcvdDetail", " WHERE LotNo='" & strLotNo & "'")
    Dim lIssuanceCount As Long, lRcvCount As Long
    lIssuanceCount = GetSingleLongValue("COUNT(*)", "VendIssdDetail", " WHERE LotNo='" & strLotNo & "'")
    lRcvCount = GetSingleLongValue("COUNT(*)", "VendRcvdDetail", " WHERE LotNo='" & strLotNo & "'")
    If lRcvCount > 0 Then
        bRepairLot = GetSingleBooleanValue("ReWorkLot", "VendRcvdDetail", " WHERE LotNo='" & strLotNo & "'")
    Else
        bRepairLot = GetSingleBooleanValue("ReWorkLot", "VendIssdDetail", " WHERE LotNo='" & strLotNo & "'")
    End If
    
    If bRepairLot Then
        If lRcvCount = 0 Then
            Call PrintIPOFromLotNo(strLotNo)
            Exit Sub
        End If
        Set rpt = rptApp.OpenReport(RptPath & "\PTC_Repair.rpt")
    Else
        If bWithPrice = False Then
            If iReportType = 1 Then
                Set rpt = rptApp.OpenReport(RptPath & "\PTCQEL_Mini.rpt")
            Else
                Set rpt = rptApp.OpenReport(RptPath & "\PTCQEL.rpt")
            End If
        Else
            Set rpt = rptApp.OpenReport(RptPath & "\PTCQELWithPrice.rpt")
        End If
    End If
    rpt.FormulaFields.GetItemByName("ComputerName").Text = "'" & strComputerName & "'"
    rpt.FormulaFields.GetItemByName("UserName").Text = "'" & CurrentUserName & "'"
    
    Dim bAllowDuplicate As Boolean
    bAllowDuplicate = GetSingleBooleanValue("PrintDuplicate_PTC", "Users", " WHERE UserName='" & CurrentUserName & "'")
    
    rpt.ParameterFields.GetItemByName("@LotNo").SetCurrentValue strLotNo
    'rpt.ParameterFields.GetItemByName("@Repair").SetCurrentValue bRepairLot
    
    rpt.EnableParameterPrompting = False
    f.ShowReport "", rpt, "", False, , , True, "TrackPrinting_PTC", "LotNo", strLotNo, bAllowDuplicate
    

    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub PrintIPOFromLotNo(strLotNo As String)

    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    
    Set rpt = rptApp.OpenReport(RptPath & "\IssSlip.rpt")
    
     
    Dim f As New frmPrevRpt
'    If strList = "" Then
'        f.ShowReport "{VendIssued.EntryID} IN[" & LV.SelectedItem.Tag & "]", rpt
'    Else
'        f.ShowReport "{VendIssued.EntryID} IN[" & Left(strList, Len(strList) - 1) & "]", rpt
'    End If
    f.ShowReport "{VendIssdDetail.LotNo}='" & strLotNo & "'", rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
Public Sub PrintPTCForElectro(strLotNo As String)
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ProductionSpecification.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{InternalProductionOrderDetailLots.EntryID}=" & strLotNo, rpt
End Sub

Public Sub PrintAllPTC(strOrderNo As String)

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Dim f As New frmPrevRpt, strSelection As String
    
    'Dim bRepairLot As Boolean
    
    'bRepairLot = GetSingleBooleanValue("ReWorkLot", "VendRcvdDetail", " WHERE LotNo='" & strLotNo & "'")
    
'    If bRepairLot Then
'        Set rpt = rptApp.OpenReport(RptPath & "\PTC_Repair.rpt")
'    Else
'        If bWithPrice = False Then
'            Set rpt = rptApp.OpenReport(RptPath & "\PTCQEL.rpt")
'        Else
'            Set rpt = rptApp.OpenReport(RptPath & "\PTCQELWithPrice.rpt")
'        End If
'    End If
    Set rpt = rptApp.OpenReport(RptPath & "\PTCQEL.rpt")
    strSelection = "{VFOrderItemPTC.OrderNo}='" & strOrderNo & "'"
    
    f.ShowReport strSelection, rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
    

Public Sub PrintOrderStatusReportProcessWise(Optional bWranglerFormat As Boolean = False, Optional bUrgentOnly As Boolean = False)

    Dim f As New frmGetCustomerANDOrders
    Load f
    Dim strOrders As String, strCustCode As String, strCountry As String
    If f.getData(strCustCode, strCountry, strOrders) = False Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    If bWranglerFormat Then
        Set rpt = rptApp.OpenReport(RptPath & "\OrderStatusReport_II.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\OrderStatusReport.rpt")
    End If
    
    rpt.ParameterFields.GetItemByName("@OrderNo").SetCurrentValue strOrders
    rpt.ParameterFields.GetItemByName("@CustCode").SetCurrentValue strCustCode
    rpt.ParameterFields.GetItemByName("@Country").SetCurrentValue strCountry
    rpt.ParameterFields.GetItemByName("@FOI_IDs").SetCurrentValue "0"
    
    rpt.EnableParameterPrompting = False
    
    Dim frm As New frmPrevRpt
    Dim strSelection As String
    If bUrgentOnly Then
        strSelection = "{OrderStatusReport_SP;1.DeliveryStatus}=2"
    Else
        strSelection = ""
    End If
     
    frm.ShowReport strSelection, rpt, "", False
    
End Sub


Public Sub PrintCTC(lEntryID As Long)

    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\IssSlip.rpt")
    
    Dim f As New frmPrevRpt
    f.ShowReport "{VendIssued.EntryID}=" & lEntryID, rpt
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Public Function PrintIssSlipWithReceiving(lEntryID As Long)

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\IssSlipWithRcving.rpt")
    
    Dim f As New frmPrevRpt
    f.ShowReport "{VendIssued.EntryID}=" & lEntryID, rpt
    
    Exit Function
err:
    MsgBox err.Description
    
End Function

Public Sub PrintProcessWiseProductionReport()
    'PrintMonthlyDateWiseProductionReport
    Dim f As New frmGetNumericDataValues
    Load f
    
    Dim strIDs As String, myDTFrom As Date, myDTTo As Date
    
    If f.getData("ItemGroups", "Description", "ID", strIDs, myDTTo) = False Then Exit Sub
'    strIDs = "125,124,58,123,122"
'    myDTFrom = #8/12/2011#
'    myDTTo = #8/18/2011#
    
    myDTFrom = DateAdd("d", -7, myDTTo)
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ProcessWiseProductionReport.rpt")
    
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue myDTFrom
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue myDTTo
    rpt.ParameterFields.GetItemByName("@GroupIDs").SetCurrentValue IIf(strIDs = "All", "0", strIDs)
    
    rpt.EnableParameterPrompting = False
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, "", False
    
End Sub

Public Sub PrintMonthlyDateWiseProductionReport(Optional bAskForDates As Boolean = True, Optional myDTFrom As Date, Optional myDTTo As Date)
    'PrintMonthlyDateWiseProductionReport
     
    If bAskForDates Then
        Dim f As New frmGetDateSingle, myDT As Date
        f.DTPicker1.CustomFormat = "MMMM, yyyy"
        Load f
        If f.getDate(myDT) = False Then Exit Sub
        myDTFrom = DateSerial(year(myDT), Month(myDT), 1)
        myDTTo = DateSerial(year(myDT), Month(myDT) + 1, 0)
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\MonthWiseProduction.rpt")
    
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue myDTFrom
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue myDTTo
    
    rpt.EnableParameterPrompting = False
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, "", False
    
End Sub

Public Sub PrintPieciesReport()
    'printCustomerStatus_Valuewise
    Dim f As New frmGetCustomerANDItemGroups
    Load f
    
    Dim strItemGroups As String, strCustomers As String
    If f.getData(strCustomers, strItemGroups) = False Then Exit Sub
     
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\PiecesReport.rpt")
    
    rpt.ParameterFields.GetItemByName("@CustCode").SetCurrentValue strCustomers
    rpt.ParameterFields.GetItemByName("@ItemGroups").SetCurrentValue strItemGroups
    
    rpt.EnableParameterPrompting = False

    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, "", False
    
End Sub

Public Sub ShowInitialPlanningMemo(p_lFOI_ID As Long, lRunningStatus As Long, lPOStatus As Long, lSemiStatus As Long, lFinishStatus As Long, lForgingStatus As Long)

    Dim f As New frmInitialPlanningMemo
    
    Load f
    Call f.ViewInitialPlanningMemo(p_lFOI_ID, lRunningStatus, lPOStatus, lSemiStatus, lFinishStatus, lForgingStatus)
    
End Sub

Public Sub PrintCustomerStatus_Valuewise()
    '
    Dim f As New frmGetCustomerANDDateRange
    Load f
    
    Dim strItemGroups As String, strCustomers As String, myDTFrom As Date, myDTTo As Date
    If f.getData(strCustomers, myDTFrom, myDTTo) = False Then Exit Sub
     
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\CustomerStats_ValueWise_SP.rpt")
    
    rpt.ParameterFields.GetItemByName("@CustCode").SetCurrentValue strCustomers
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue myDTFrom
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue myDTTo
    
    rpt.EnableParameterPrompting = False

    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, "", False
    
End Sub

Public Sub PrintDemandAndSupplyGraph(myDT As Date)
    
    Dim myDTFrom As Date, myDTTo As Date
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\DemandAndSupplyGraph.rpt")
    
    myDTFrom = DateSerial(year(myDT), Month(myDT), 1)
    myDTTo = DateSerial(year(myDT), Month(myDT) + 1, 0)
    
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue myDTFrom
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue myDTTo
    rpt.ParameterFields.GetItemByName("@Summary").SetCurrentValue True
    
    rpt.FormulaFields.GetItemByName("DateRange").Text = "'" & Format(myDT, "MMM, yyyy") & "'"
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, , False
    
End Sub

Public Sub PrepareDemandAndSupplyGraph()

    Dim myDT As Date
    Dim f As New frmGetDateSingle
    f.DTPicker1.CustomFormat = "MMM, yyyy"
    Load f
    If f.getDate(myDT) = False Then Exit Sub
    
    Call PrintDemandAndSupplyGraph(myDT)
    
End Sub

Public Sub PrintDeliveryPerformanceGraph()
    
    Dim myDT As Date
    Dim f As New frmGetDateSingle
    f.DTPicker1.CustomFormat = "MMM, yyyy"
    Load f
    If f.getDate(myDT) = False Then Exit Sub
    
    Dim myDTFrom As Date, myDTTo As Date
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ExportDeliveryPerformanceReport_Graph.rpt")
    
    myDTFrom = DateSerial(year(myDT), Month(myDT), 1)
    myDTTo = DateSerial(year(myDT), Month(myDT) + 1, 0)
    
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue myDTFrom
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue myDTTo
     
    rpt.FormulaFields.GetItemByName("DateRange").Text = "'" & Format(myDT, "MMM, yyyy") & "'"
    rpt.EnableParameterPrompting = False
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, , False
    
End Sub

Public Sub PrintMakerPerformanceReportWrangler()

    Dim myDTFrom As Date, myDTTo As Date, lMakerID As Long
    Dim f As New frmGetMakerANDDateRange
   
    Load f
    If f.getData(lMakerID, myDTFrom, myDTTo) = False Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\MakerPerformanceReport.rpt")
    rpt.FormulaFields.GetItemByName("UserName").Text = "'" & CurrentUserName & "'"
     
    Dim strSelection As String
    strSelection = "{VMakerPerformance.VendID}<>" & iFactoryMakerID
    
    strSelection = strSelection & " AND {VMakerPerformance.DT}=#" & myDTFrom & "# TO #" & myDTTo & "#"
    
    If lMakerID <> 0 Then
        strSelection = strSelection & " AND {VMakerPerformance.VendID}=" & lMakerID
    End If
    
    Dim frm As New frmPrevRpt
    frm.ShowReport strSelection, rpt
    
End Sub

Public Function getNextSubLotNo(strLotNo As String) As String

    Dim lRcvLotNo As Long, lIssLotNo As Long, lLotList_LotNo As Long, lSOO_LotNo As Long, lRcv_LotNo As Long
    Dim lLotNo As Long
    Dim iLen As Integer
    strLotNo = strLotNo & "-"
    iLen = Len(strLotNo)
        
     
    lLotList_LotNo = GetSingleLongValue("MAX(CAST(REPLACE(LotNo,'" & strLotNo & "','') AS INT))", "Lots_List", " WHERE LEFT(LotNo," & iLen & ")='" & strLotNo & "' AND CHARINDEX('-',REPLACE(LotNo,'" & strLotNo & "',''))=0")
    lSOO_LotNo = GetSingleLongValue("MAX(CAST(REPLACE(LotNo_Manual,'" & strLotNo & "','') AS INT))", "StockOrderOpening", " WHERE LEFT(LotNo_Manual," & iLen & ")='" & strLotNo & "' AND CHARINDEX('-',REPLACE(LotNo_Manual,'" & strLotNo & "',''))=0")
    lRcv_LotNo = GetSingleLongValue("MAX(CAST(REPLACE(LotNo_Manual,'" & strLotNo & "','') AS INT))", "RcvItemsSimpleDetail", " WHERE LEFT(LotNo_Manual," & iLen & ")='" & strLotNo & "' AND CHARINDEX('-',REPLACE(LotNo_Manual,'" & strLotNo & "',''))=0")
    
    
    lRcvLotNo = GetSingleLongValue("MAX(CAST(REPLACE(LotNo,'" & strLotNo & "','') AS INT))", "VendRcvdDetail", " WHERE LEFT(LotNo," & iLen & ")='" & strLotNo & "' AND CHARINDEX('-',REPLACE(LotNo,'" & strLotNo & "',''))=0")
    lIssLotNo = GetSingleLongValue("MAX(CAST(REPLACE(LotNo,'" & strLotNo & "','') AS INT))", "VendIssdDetail", " WHERE LEFT(LotNo," & iLen & ")='" & strLotNo & "' AND CHARINDEX('-',REPLACE(LotNo,'" & strLotNo & "',''))=0")
    
    If lRcvLotNo > lIssLotNo Then
        If lRcvLotNo > lsso_lotno Then
            lLotNo = lRcvLotNo
        Else
            lLotNo = lSOO_LotNo
        End If
    Else
        If lIssLotNo > lSOO_LotNo Then
            lLotNo = lIssLotNo
        Else
            lLotNo = lSOO_LotNo
        End If
    End If
    If lRcv_LotNo > lLotNo Then
        lLotNo = lRcv_LotNo
    End If
    If lLotList_LotNo > lLotNo Then
        lLotNo = lLotList_LotNo
    End If
    
'    If lRcvLotNo = 0 Then
'        lRcvLotNo = GetSingleLongValue("MAX(CAST(REPLACE(LotNo,'" & strLotNo & "','') AS INT))", "VendIssdDetail", " WHERE LEFT(LotNo," & iLen & ")='" & strLotNo & "'")
'    End If
    
    'lLotNo = lRcvLotNo
    
    lLotNo = lLotNo + 1
    'getNextLotNo = lLotNo & "" & Format(Date, "yy")
    getNextSubLotNo = strLotNo & lLotNo
    
End Function

Public Function getNextBatchNo() As String

    Dim lBatchNo As Long
    lBatchNo = GetSingleLongValue("MAX(CAST(RIGHT(Batch_No,LEN(Batch_No)-3) AS INT))", "VendIssdDetail", " WHERE LEFT(Batch_No,3)='B17' AND ISNUMERIC(RIGHT(Batch_No,LEN(Batch_No)-3))=1")
    lBatchNo = lBatchNo + 1
    
    getNextBatchNo = "B17" & lBatchNo
    
End Function
