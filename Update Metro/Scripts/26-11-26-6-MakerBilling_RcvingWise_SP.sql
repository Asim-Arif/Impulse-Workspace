
CREATE OR ALTER PROCEDURE [dbo].[MakerBilling_RcvingWise_SP] (
    @MakerID AS INT = 0,
    @DTFrom AS DATETIME,
    @DTTo AS DATETIME,
    @CatID AS VARCHAR(255),
    @ProcessID AS INT = 0,
    @MasterPONo VARCHAR(50) = ''
)
AS
BEGIN

DECLARE @IncludeRejected BIT = 0;
SELECT TOP 1 @IncludeRejected =
    CASE
        WHEN UPPER(LTRIM(RTRIM(ISNULL(DataValue, '')))) IN ('TRUE', '1', 'YES') THEN 1
        ELSE 0
    END
FROM dbo.GeneralData
WHERE DataName = 'MakerBilling_IncludeRejectedQty';

CREATE TABLE #myTBL (CatID VARCHAR(50))
CREATE TABLE #Processes(ProcessID INT)
IF @ProcessID=0 
	INSERT INTO #Processes SELECT ProcessID FROM Processes
ELSE
	INSERT INTO #Processes SELECT ProcessID FROM Processes WHERE ProcessID=@ProcessID
	
DECLARE @myQuery AS VARCHAR(1000)

IF @CatID='0' 
	INSERT INTO #myTBL SELECT CatID FROM ItemCatagories
ELSE
	BEGIN	
		SET @myQuery='INSERT INTO #myTBL(CatID) SELECT CatID FROM ItemCatagories WHERE CatID IN(' + @CatID +')'
		EXEC(@myQuery)
	END
	
SELECT Makers.VendID,Makers.VendID1,Makers.VenderName,VendIssued.EntryID,VendIssued.RecieptID,Vendissued.DT,VendIssued.ItemID,TabIssd.IssQty,TabIssd.Rate,TempAdvance,TabRcvd.RcvdQty ,RcvdQtyUnAuth,RcvdWeight,
RcvdWeightUnAuth,RcvdValue,RcvdValueUnAuth
--,(TabRcvd.RcvdQty+RcvdQtyUnAuth)/CAST(TabIssd.IssQty AS FLOAT)*100 AS PercentRcvd
,100 AS PercentRcvd
,TabRcvd.RepairAmt,ItemName,ItemNameUrdu,Wastage,Accounts.Balance As AccBalance,TabRcvd.OrderNo,
Issue_RefID,TabRcvd.EntryID AS RcvDetailEntryID,TabRcvd.RecieptID AS RcvRecieptID,TabRcvd.LotNo,
BillingProcRcvdQty - CASE WHEN @IncludeRejected = 1 THEN 0 ELSE ISNULL(WastageQtyFromRepair,0) END AS BillingProcRcvdQty,
BillingProcAmount - (CASE WHEN @IncludeRejected = 1 THEN 0 ELSE ISNULL(WastageQtyFromRepair,0) END * TabIssd.Rate) AS BillingProcAmount
--BillingProcRcvdQty-ISNULL(WastageQtyFromRepair,0) AS BillingProcRcvdQty,BillingProcAmount-ISNULL(WastageQtyFromRepair*Rate,0) AS BillingProcAmount
--TabRcvd.RcvdQty-ISNULL(ReWorkQty,0)-ISNULL(WastageQtyFromRepair,0) AS BillingProcRcvdQty,((TabRcvd.RcvdQty-ISNULL(ReWorkQty,0)-ISNULL(WastageQtyFromRepair,0))*Rate) AS BillingProcAmount
,OnlyDT,TabRcvd.ProcessID,Processes.Description,
CatID,TabRcvd.MasterPONo,MasterRecieptID,MainIssuanceEntryID,LotClearDT
,WastageQtyFromRepair,WastageQtyFromRepair*VendAssItems.Rate AS WastageValueFromRepair,RcvDT,ItemGroups.Description AS ItemGroupName,Audited,BillQtyDirect,BillQtyInDirect,FCustomerOrders.CustCode
,VendIssued.BillingDays,TRunningLots.Lot_Current_Process
,FCustomerOrders.InternalRefNo
FROM VendIssued
LEFT OUTER JOIN VendIssdDetail TabIssd ON VendIssued.EntryID=TabIssd.RefID
LEFT OUTER JOIN VendAssItems ON VendIssued.VendID=VendAssItems.VendID AND VendIssued.ProcessID=VendAssItems.ProcessID AND TabIssd.ItemCode=VendAssItems.ItemID
LEFT OUTER JOIN (SELECT VVendRcvdDetail.EntryID,Issuance_RefID,Issue_RefID,OrderNo,0 AS Not_Available_For_Billing
					--,CASE WHEN ReqAuth=0 THEN RcvdQty ELSE 0 END +ISNULL(SplitQtyForRcvdQty,0) AS RcvdQty
					,CASE WHEN ReqAuth=0 THEN RcvdQty ELSE 0 END AS RcvdQty
					,CASE WHEN ReqAuth=1 THEN RcvdQty ELSE 0 END AS RcvdQtyUnAuth,
					CASE WHEN ReqAuth=0 THEN RcvdWeight ELSE 0 END AS RcvdWeight,CASE WHEN ReqAuth=1 THEN RcvdWeight ELSE 0 END AS RcvdWeightUnAuth,
					CASE WHEN ReqAuth=0 THEN RcvdValue_For_Billing ELSE 0 END AS RcvdValue,CASE WHEN ReqAuth=1 THEN RcvdValue_For_Billing ELSE 0 END AS RcvdValueUnAuth,RepairAmt AS RepairAmt
,Wastage,VVendRcvdDetail.LotNo,VVendRcvdDetail.RecieptID,ISNULL(BillQtyDirect,0)+ISNULL(BillQtyInDirect,0) AS BillingProcRcvdQty,(ISNULL(BillQtyDirect,0)+ISNULL(BillQtyInDirect,0))*IssRate AS BillingProcAmount,OnlyDT,VVendRcvdDetail.ProcessID,MasterPONo,TMasters.RecieptID AS MasterRecieptID,MainIssuanceEntryID,LotClearDT,DT AS RcvDT,ReWorkQty,Audited,BillQtyDirect,BillQtyInDirect
FROM VVendRcvdDetail 
LEFT OUTER JOIN(SELECT VRD_From_RefID,SUM(SplitQty) AS SplitQtyForRcvdQty FROM LotTransferDetails GROUP BY VRD_From_RefID) TSplitQtyForRcvdQty ON VVendRcvdDetail.EntryID=TSplitQtyForRcvdQty.VRD_From_RefID

INNER JOIN(
	SELECT LotNo,
	       VendRcvdDetail.ProcessID,
	       MAX(DT) AS LotClearDT,
	       SUM(RcvdQty - ISNULL(ReWorkQty,0) - CASE WHEN @IncludeRejected = 1 THEN 0 ELSE ISNULL(Wastage,0) END) AS BillQtyDirect
	FROM VendRcvdDetail
	INNER JOIN VendReceived ON VendReceived.EntryID=VendRcvdDetail.RefID
	GROUP BY LotNo,VendRcvdDetail.ProcessID
) TBillQtyDirect ON VVendRcvdDetail.LotNo=TBillQtyDirect.LotNo AND VVendRcvdDetail.BillingProcessID=TBillQtyDirect.ProcessID
LEFT OUTER JOIN(SELECT VRD_From.LotNo,VRD_Lots.ProcessID,SNo,SUM(VRD_Lots.RcvdQty-ISNULL(VRD_Lots.ReWorkQty,0)) AS BillQtyInDirect FROM LotTransferDetails
	INNER JOIN VendRcvdDetail VRD_From ON LotTransferDetails.VRD_From_RefID=VRD_From.EntryID
	INNER JOIN VendRcvdDetail VRD_To ON LotTransferDetails.VRD_To_RefID=VRD_To.EntryID
	INNER JOIN VendRcvdDetail VRD_Lots ON VRD_To.LotNo=VRD_Lots.LotNo
	INNER JOIN VendReceived ON VRD_Lots.RefID=VendReceived.EntryID
	INNER JOIN ItemProcesses ON VRD_From.ProcessID=ItemProcesses.ProcessID AND ItemProcesses.ItemID=VRD_From.ItemCode 
	WHERE VRD_Lots.ReqAuth=0 AND Type=1 GROUP BY VRD_From.LotNo,VRD_Lots.ProcessID,SNo) TBillQtyInDirect ON VVendRcvdDetail.LotNo=TBillQtyInDirect.LotNo AND VVendRcvdDetail.BillingProcessID=TBillQtyInDirect.ProcessID AND (TBillQtyInDirect.SNo>VVendRcvdDetail.ItemSNo)

LEFT OUTER JOIN (SELECT LotNo,MIN(EntryID) AS MinEntryID FROM VendRcvdDetail GROUP BY LotNo) TMinLot ON VVendRcvdDetail.LotNo=TMinLot.LotNo
LEFT OUTER JOIN (SELECT VendIssued.EntryID AS MainIssuanceEntryID,VendRcvdDetail.EntryID,MasterPONo,VendIssued.RecieptID FROM VendIssued INNER JOIN VendIssdDetail ON VendIssdDetail.RefID=VendIssued.EntryID INNER JOIN VendRcvdDetail ON VendRcvdDetail.Issue_RefID=VendIssdDetail.EntryID) TMasters
	ON TMinLot.MinEntryID=TMasters.EntryID

WHERE (VendID=@MakerID OR @MakerID=0) AND OnlyDT BETWEEN @DTFrom AND @DTTo  --AND ReWorkLot=0
			) TabRcvd ON VendIssued.EntryID=TabRcvd.Issuance_RefID AND TabIssd.EntryID=TabRcvd.Issue_RefID
LEFT OUTER JOIN (SELECT Issuance_RefID,SUM(Amount) AS TempAdvance FROM MakerTemporaryLoans WHERE ISNULL(GroupID,0)=0 OR GroupID IN(SELECT EntryID FROM MakerTemporaryLoansGroup WHERE Active=1) 
																	GROUP BY Issuance_RefID) TabTemp ON VendIssued.EntryID=TabTemp.Issuance_RefID

LEFT OUTER JOIN (SELECT VRD_RefID_ChargeTo,SUM(Qty) AS WastageQtyFromRepair FROM VMakerRepair_BIlling WHERE ShowInBilling=0 AND ItemSNo>BillingProcessSNo AND BillingProcessSNo IS NOT NULL GROUP BY VRD_RefID_ChargeTo) TMakerWastage ON TabRcvd.EntryID=TMakerWastage.VRD_RefID_ChargeTo
INNER JOIN Items ON VendIssued.ItemID=Items.ItemID
INNER JOIN ItemGroups ON Items.GroupID=ItemGroups.ID
INNER JOIN Makers ON VendIssued.VendID=Makers.VendID
INNER JOIN Accounts ON Makers.AccNo=Accounts.AccNo
INNER JOIN Processes ON TabRcvd.ProcessID=Processes.ProcessID
INNER JOIN #Processes ON #Processes.ProcessID=TabRcvd.ProcessID
INNER JOIN FCustomerOrders ON TabIssd.OrderNo=FCustomerOrders.OrderNo
LEFT JOIN (SELECT LotNo,MAX(Description) AS Lot_Current_Process FROM VRunningLots_Simple GROUP BY LotNo) TRunningLots ON TabRcvd.LotNo=TRunningLots.LotNo
WHERE (VendIssued.VendID=@MakerID OR @MakerID=0) AND Issue_RefID IS NOT NULL
AND TabRcvd.EntryID NOT IN(SELECT VRD_RefID FROM MakerPostedBillsDetail_Receivings WHERE VRD_RefID IS NOT NULL)
AND Items.CatID IN(SELECT CatID FROM #myTBL)
AND TabRcvd.EntryID NOT IN(SELECT VRD_RefID FROM MakerBilling_Receivings_Hidden WHERE VRD_RefID IS NOT NULL)
AND TabRcvd.Not_Available_For_Billing=0
AND ((TabRcvd.MasterPONo LIKE '%'+@MasterPONo+'%') OR @MasterPONo='')
AND OnlyDT>='12/1/2019'
--ORDER BY RcvDT
ORDER BY FCustomerOrders.InternalRefNo
--ORDER BY (RcvdQty+RcvdQtyUnAuth)/CAST(TabIssd.IssQty AS FLOAT) *100 DESC
END
GO

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO