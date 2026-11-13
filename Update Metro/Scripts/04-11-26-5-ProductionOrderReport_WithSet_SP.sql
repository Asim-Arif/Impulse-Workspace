
/****** Object:  StoredProcedure [dbo].[ProductionOrderReport_WithSet_SP]    Script Date: 11/4/2026 8:41:31 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


/****** Object:  StoredProcedure [dbo].[ProductionOrderReportPHF_SP]    Script Date: 03/10/2014 09:45:10 ******/
/*SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO*/

CREATE PROCEDURE [dbo].[ProductionOrderReport_WithSet_SP] (@OrderNo AS VARCHAR(255)) AS

BEGIN
	SET ANSI_NULLS ON
	
	SET ARITHABORT ON


	SELECT     dbo.VrptOrders_ForProduction.OrderNo, dbo.VrptOrders_ForProduction.DT, dbo.VrptOrders_ForProduction.TradeTerms, dbo.VrptOrders_ForProduction.CustCode, dbo.VrptOrders_ForProduction.Country, dbo.VrptOrders_ForProduction.PartialShipment, 
                      dbo.VrptOrders_ForProduction.PaymentTerms, dbo.VrptOrders_ForProduction.TransShipment, dbo.VrptOrders_ForProduction.Packaging, dbo.VrptOrders_ForProduction.DeliveryDT, dbo.VrptOrders_ForProduction.BatchNo, 
                      dbo.VrptOrders_ForProduction.CompanyRefID, dbo.VrptOrders_ForProduction.StampDT, dbo.VrptOrders_ForProduction.Quality, dbo.VrptOrders_ForProduction.InternalRefNo, dbo.VrptOrders_ForProduction.OrderRcvdVia, 0 AS ID, 
                      '' AS ItemCode, 0 AS Price, dbo.VrptOrders_ForProduction.Qty, 0 AS InvQty, dbo.VrptOrders_ForProduction.ItemName, dbo.VrptOrders_ForProduction.Description, 
                      dbo.VrptOrders_ForProduction.ItemID, dbo.VrptOrders_ForProduction.CompItemID, dbo.VrptOrders_ForProduction.Curr, dbo.VrptOrders_ForProduction.Unit, dbo.VrptOrders_ForProduction.Name, dbo.VrptOrders_ForProduction.Address, 
                      0 AS CustomPrice, dbo.VrptOrders_ForProduction.CustomDescription, dbo.VrptOrders_ForProduction.ItemSize, dbo.VrptOrders_ForProduction.SizeUnit, dbo.VrptOrders_ForProduction.SteelType, dbo.VrptOrders_ForProduction.Gage, 
                      dbo.VrptOrders_ForProduction.ItemColor, dbo.VrptOrders_ForProduction.BarcodeNo, dbo.VrptOrders_ForProduction.PackingMode, dbo.VrptOrders_ForProduction.MasterCartonL, dbo.VrptOrders_ForProduction.MasterCartonW, 
                      dbo.VrptOrders_ForProduction.MasterCartonH, dbo.VrptOrders_ForProduction.SmallCartonL, dbo.VrptOrders_ForProduction.SmallCartonW, dbo.VrptOrders_ForProduction.SmallCartonH, dbo.VrptOrders_ForProduction.PolyBag, 
                      dbo.VrptOrders_ForProduction.SpecialInstructions, dbo.VrptOrders_ForProduction.StampInstructions, dbo.VrptOrders_ForProduction.PackingInstructions, dbo.VrptOrders_ForProduction.UnitItems, dbo.VrptOrders_ForProduction.CustomerColor, 
                      dbo.VrptOrders_ForProduction.ShippedQty, dbo.VrptOrders_ForProduction.ItemNameUrdu, dbo.VrptOrders_ForProduction.InHand, 0 AS SortNo, dbo.VrptOrders_ForProduction.MaxLotSize, dbo.VrptOrders_ForProduction.FinQuality
                      ,dbo.VrptOrders_ForProduction.Type, TabPO.RecieptID, TabPO.IssQty, TabAdj.AdjQty
                      ,TPTC.PTC
                      ,TProcesses.Processes
                      ,TPOQty.POQty
                      ,TPONo.PONos
                      ,TabPacked.PackQty
                      ,TCartons.Cartons
                      ,TabQC.RcvdQtyQC,TabQC.WastageQC,TabQC.ReWorkQtyQC,TabInsp.RcvdQtyInsp,TabInsp.WastageInsp,TabInsp.ReWorkQtyInsp,ReWorkQtyQCBalance,TotalQtyRcvd,TotalTransferred,ReWorkQtyInspBalance
                      ,TIMR.InitialMakerRcvings
                      ,'' AS DeliveryDTItem,'' AS DeliveryStatus,'' AS StampsItem,'' AS QualityItem,'' AS Remarks
					  ,TPOMN.POMakerName
					  ,VenderName,ReturnDT,'' AS GroupDescription,0 AS GroupID,CompanyName,0 AS FinishedWeight,0 AS SmallBoxPcs,0 AS MasterCartonSmallBoxes
					  ,0 AS MainItem,ItemsSets.Set_ItemID
FROM         dbo.VrptOrders_ForProduction WITH (NOLOCK) 
			LEFT OUTER JOIN Companies WITH (NOLOCK)ON VrptOrders_ForProduction.CompanyRefID=Companies.EntryID 
			LEFT OUTER JOIN ItemsSets WITH (NOLOCK)ON VrptOrders_ForProduction.ItemID=ItemsSets.ItemID
			AND VrptOrders_ForProduction.CompItemID=ItemsSets.Set_ItemID
			LEFT OUTER JOIN
                          (SELECT     OrderNo, ItemCode, RecieptID, IssQty,VenderName,ReturnDT
                            FROM          dbo.VVendIssdDetail WITH (NOLOCK) INNER JOIN
                            (SELECT     MIN(EntryID) AS MINEntryID FROM dbo.VVendIssdDetail WITH (NOLOCK) WHERE OrderNo=@OrderNo AND VendID<>653 GROUP BY OrderNo, ItemCode) TI1 ON dbo.VVendIssdDetail.EntryID=TI1.MinEntryID) TabPO ON dbo.VrptOrders_ForProduction.OrderNo = TabPO.OrderNo AND 
                      dbo.VrptOrders_ForProduction.CompItemID = TabPO.ItemCode 
					  
					  LEFT OUTER JOIN
                      (SELECT     OrderNo, ItemID, SUM(Qty) AS AdjQty
                      FROM          StockOrderAdjustments WITH (NOLOCK) WHERE OrderNo=@OrderNo
                      GROUP BY OrderNo, ItemID) TabAdj ON dbo.VrptOrders_ForProduction.OrderNo = TabAdj.OrderNo AND dbo.VrptOrders_ForProduction.CompItemID = TabAdj.ItemID LEFT OUTER JOIN
                      (SELECT     OrderNo, ItemCode, SUM(Qty) AS PackQty
                      FROM          VDispatchListDetail  WITH (NOLOCK) WHERE OrderNo=@OrderNo
                      GROUP BY OrderNo, ItemCode) TabPacked ON dbo.VrptOrders_ForProduction.OrderNo = TabPacked.OrderNo AND 
                      dbo.VrptOrders_ForProduction.CompItemID = TabPacked.ItemCode 
					  
					  LEFT OUTER JOIN
                          (SELECT     OrderNo, ItemCode, SUM(RcvdQty) AS RcvdQtyQC, SUM(Wastage + LostQty) AS WastageQC, SUM(VendRcvdDetailReWorkDetail.Qty) AS ReWorkQtyQC, 
                                                   SUM(VendRcvdDetailReWorkDetail.Qty - VendRcvdDetailReWorkDetail.IssQty) AS ReWorkQtyQCBalance
                            FROM          VendRcvdDetail  WITH (NOLOCK) LEFT OUTER JOIN
                                                   VendRcvdDetailReWorkDetail  WITH (NOLOCK) ON VendRcvdDetail.EntryID = VendRcvdDetailReWorkDetail.VRD_RefID
                            WHERE      ProcessID = 69 AND ReqAuth = 0 AND ReWorkLot=0 AND OrderNo=@OrderNo
                            GROUP BY OrderNo, ItemCode) TabQC ON dbo.VrptOrders_ForProduction.OrderNo = TabQC.OrderNo AND dbo.VrptOrders_ForProduction.CompItemID = TabQC.ItemCode LEFT OUTER JOIN
                          (SELECT     OrderNo, ItemCode, SUM(RcvdQty) AS RcvdQtyInsp, SUM(Wastage + LostQty) AS WastageInsp, SUM(VendRcvdDetailReWorkDetail.Qty) AS ReWorkQtyInsp,
			SUM(VendRcvdDetailReWorkDetail.Qty - VendRcvdDetailReWorkDetail.IssQty) AS ReWorkQtyInspBalance
                            FROM          VendRcvdDetail  WITH (NOLOCK) LEFT OUTER JOIN
                                                   VendRcvdDetailReWorkDetail  WITH (NOLOCK) ON VendRcvdDetail.EntryID = VendRcvdDetailReWorkDetail.VRD_RefID
                            WHERE      ProcessID = 65 AND ReqAuth = 0 AND ReWorkLot=0 AND OrderNo=@OrderNo
                            GROUP BY OrderNo, ItemCode) TabInsp ON dbo.VrptOrders_ForProduction.OrderNo = TabInsp.OrderNo AND dbo.VrptOrders_ForProduction.CompItemID = TabInsp.ItemCode LEFT OUTER JOIN
		(SELECT VendRcvdDetail.OrderNo,VendRcvdDetail.ItemCode,SUM(VendRcvdDetail.RcvdQty) AS TotalQtyRcvd FROM VendRcvdDetail  WITH (NOLOCK) LEFT OUTER JOIN VendIssdDetail  WITH (NOLOCK) ON VendRcvdDetail.Issue_RefID=VendIssdDetail.EntryID 
			WHERE ISNULL(VendIssdDetail.LotNo,'0')='0' AND VendRcvdDetail.OrderNo=@OrderNo--VendRcvdDetail.EntryID IN(SELECT MIN(EntryID) FROM VendRcvdDetail WHERE ReWorkLot=0)  
			GROUP BY VendRcvdDetail.OrderNo,VendRcvdDetail.ItemCode) TTotalRcvd ON dbo.VrptOrders_ForProduction.OrderNo = TTotalRcvd.OrderNo AND dbo.VrptOrders_ForProduction.CompItemID = TTotalRcvd.ItemCode LEFT OUTER JOIN
		(SELECT OrderNo,ItemCode,SUM(Qty) AS TotalTransferred FROM TransferredToReadyFinishLotsDetail  WITH (NOLOCK) INNER JOIN VendRcvdDetail  WITH (NOLOCK) ON TransferredToReadyFinishLotsDetail.VRD_RefID=VendRcvdDetail.EntryID WHERE OrderNo=@OrderNo
			GROUP BY OrderNo,ItemCode) TTTRF ON dbo.VrptOrders_ForProduction.OrderNo = TTTRF.OrderNo AND dbo.VrptOrders_ForProduction.CompItemID = TTTRF.ItemCode 
			
			
		LEFT OUTER JOIN (SELECT DISTINCT OrderNo,ItemCode,
(STUFF((SELECT ','+LotNo + ' ('+CAST(RcvdQty AS VARCHAR(5)) +')' FROM VendRcvdDetail WITH (NOLOCK) 
	INNER JOIN (SELECT MAX(EntryID) AS MAXEntryID FROM VendRcvdDetail WITH (NOLOCK) WHERE OrderNo=@OrderNo GROUP BY LotNo) T1 ON VendRcvdDetail.EntryID=T1.MaxEntryID
	WHERE OrderNo=@OrderNo AND VendRcvdDetail.OrderNo=FOrderItems.OrderNo AND VendRcvdDetail.ItemCode=FOrderItems.CompItemCode  ORDER BY LotNo
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS PTC
	FROM FOrderItems WHERE OrderNo=@OrderNo) TPTC ON VrptOrders_ForProduction.OrderNo=TPTC.OrderNo AND VrptOrders_ForProduction.ItemID=TPTC.ItemCode
		

LEFT OUTER JOIN (
SELECT DISTINCT OrderNo,ItemCode,
(STUFF((SELECT CHAR(13)+Description + + ' (' + CONVERT(VARCHAR(50),DT,5) + ') ('  + CAST(RcvdQty-ISNULL(Wastage,0)-ISNULL(ReWorkQty,0) AS VARCHAR(10)) + ')' FROM VendRcvdDetail WITH (NOLOCK) INNER JOIN VendReceived WITH (NOLOCK) ON VendReceived.EntryID=VendRcvdDetail.RefID 
	INNER JOIN Processes ON VendReceived.ProcessID=Processes.ProcessID
	INNER JOIN (SELECT MAX(EntryID) AS MAXEntryID FROM VendRcvdDetail WITH (NOLOCK) WHERE OrderNo=@OrderNo AND VendRcvdDetail.OrderNo=FOrderItems.OrderNo AND VendRcvdDetail.ItemCode=FOrderItems.CompItemCode GROUP BY LotNo) T1 ON VendRcvdDetail.EntryID=T1.MAXEntryID
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS Processes
	FROM FOrderItems WHERE OrderNo=@OrderNo) TProcesses ON VrptOrders_ForProduction.OrderNo=TProcesses.OrderNo AND VrptOrders_ForProduction.ItemID=TProcesses.ItemCode

LEFT OUTER JOIN (
SELECT DISTINCT OrderNo,ItemCode,	
(STUFF((SELECT CHAR(13) + CHAR(10)+CAST(IssQty AS VARCHAR(5)) FROM VVendIssdDetail_Simple WITH (NOLOCK) INNER JOIN 
	(SELECT MIN(EntryID) AS MINEntryID FROM dbo.VVendIssdDetail_Simple WITH (NOLOCK) WHERE OrderNo=@OrderNo AND ProcessID IN(172,212) 
	AND VVendIssdDetail_Simple.OrderNo=FOrderItems.OrderNo AND VVendIssdDetail_Simple.ItemCode=FOrderItems.CompItemCode GROUP BY OrderNo,ItemCode,ProcessID) T1 ON VVendIssdDetail_Simple.EntryID=T1.MINEntryID
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS POQty
	FROM FOrderItems WHERE OrderNo=@OrderNo) TPOQty ON VrptOrders_ForProduction.OrderNo=TPOQty.OrderNo AND VrptOrders_ForProduction.ItemID=TPOQty.ItemCode
	
LEFT OUTER JOIN (
SELECT DISTINCT OrderNo,ItemCode,	
(STUFF((SELECT CHAR(13) + CHAR(10)+RecieptID FROM VVendIssdDetail_Simple WITH (NOLOCK) INNER JOIN 
	(SELECT MIN(EntryID) AS MINEntryID FROM dbo.VVendIssdDetail_Simple WITH (NOLOCK) WHERE OrderNo=@OrderNo AND ProcessID IN(172,212) 
	AND VVendIssdDetail_Simple.OrderNo=FOrderItems.OrderNo AND VVendIssdDetail_Simple.ItemCode=FOrderItems.CompItemCode GROUP BY OrderNo,ItemCode,ProcessID) T1 ON VVendIssdDetail_Simple.EntryID=T1.MINEntryID
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS PONos
	FROM FOrderItems WHERE OrderNo=@OrderNo) TPONo ON VrptOrders_ForProduction.OrderNo=TPONo.OrderNo AND VrptOrders_ForProduction.ItemID=TPONo.ItemCode
	
LEFT OUTER JOIN (
SELECT DISTINCT OrderNo,FOrderItems.CompItemCode,
(STUFF((SELECT CHAR(13)+CHAR(10)+CAST(CartonNo AS VARCHAR(10)) + '(' + CAST(SUM(Qty) AS VARCHAR(10)) + ')' FROM VDispatchListDetail_Adv WITH (NOLOCK) WHERE OrderNo=@OrderNo
	AND VDispatchListDetail_Adv.OrderNo=FOrderItems.OrderNo AND VDispatchListDetail_Adv.ItemCode=FOrderItems.CompItemCode GROUP BY OrderNo,ItemCode,CartonNo
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS Cartons
	FROM FOrderItems WHERE OrderNo=@OrderNo) TCartons ON VrptOrders_ForProduction.OrderNo=TCartons.OrderNo AND VrptOrders_ForProduction.CompItemID=TCartons.CompItemCode

LEFT OUTER JOIN (
SELECT DISTINCT OrderNo,ItemCode,
(STUFF((SELECT CHAR(13)+VendRcvdDetail.LotNo+ ' (' + CAST(SUM(VendRcvdDetail.RcvdQty) AS VARCHAR(10)) + ')' FROM VendRcvdDetail WITH (NOLOCK) 
	INNER JOIN VendIssdDetail ON VendRcvdDetail.Issue_RefID=VendIssdDetail.EntryID WHERE VendIssdDetail.LotNo='0' AND VendRcvdDetail.ReWorkLot=0 AND VendRcvdDetail.EntryID NOT IN(SELECT EntryID FROM LotTransferDetails)
		AND VendRcvdDetail.OrderNo=@OrderNo	
		AND VendRcvdDetail.OrderNo=FOrderItems.OrderNo AND VendRcvdDetail.ItemCode=ForderItems.CompItemCode
		GROUP BY VendRcvdDetail.LotNo
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS InitialMakerRcvings
	FROM FOrderItems WHERE OrderNo=@OrderNo) TIMR ON VrptOrders_ForProduction.OrderNo=TIMR.OrderNo AND VrptOrders_ForProduction.ItemID=TIMR.ItemCode
	
LEFT OUTER JOIN (SELECT DISTINCT OrderNo,ItemCode,
(STUFF((SELECT CHAR(13)+CHAR(10)+VenderName FROM VendIssdDetail WITH (NOLOCK) INNER JOIN VendIssued WITH (NOLOCK) ON VendIssued.EntryID=VendIssdDetail.RefID INNER JOIN Makers WITH (NOLOCK) ON Makers.VendID=VendIssued.VendID INNER JOIN
	(SELECT MIN(EntryID) MINEntryID FROM dbo.VVendIssdDetail_Simple WITH (NOLOCK) WHERE ProcessID IN(172,212) AND VVendIssdDetail_Simple.OrderNo=@OrderNo AND VVendIssdDetail_Simple.OrderNo=FOrderItems.OrderNo AND VVendIssdDetail_Simple.ItemCode=ForderItems.CompItemCode GROUP BY OrderNo,ItemCode,ProcessID) T1 ON VendIssdDetail.EntryID=T1.MinEntryID
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS POMakerName
	FROM FOrderItems WHERE OrderNo=@OrderNo) TPOMN ON VrptOrders_ForProduction.OrderNo=TPOMN.OrderNo AND VrptOrders_ForProduction.ItemID=TPOMN.ItemCode
	WHERE VrptOrders_ForProduction.OrderNo=@OrderNo

END
GO

