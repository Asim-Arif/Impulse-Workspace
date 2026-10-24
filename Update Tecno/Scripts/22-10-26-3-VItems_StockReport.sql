USE [Impulse_AWM]
GO

/****** Object:  View [dbo].[VItems_StockReport]    Script Date: 10/22/2026 7:12:15 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO




ALTER VIEW [dbo].[VItems_StockReport]
AS
SELECT        dbo.VItems.CatID, dbo.VItems.ItemID, dbo.VItems.ItemName, dbo.VItems.Unit, dbo.VItems.Type, dbo.VItems.ItemSize, dbo.VItems.SizeUnit, dbo.VItems.ItemUsage, dbo.VItems.FinQuality, dbo.VItems.ItemGroup, 
                         dbo.VItems.InActive, TSF.SFStock, TForging.ForgingStock, dbo.VItems.MinLevel, dbo.VItems.MaxLevel, dbo.VItems.ReOrderLevel, TFinished.InHand, TRunning.InProductionQty, dbo.VStoreShelfs.ShelfNo, TFinished.Shelf_RefID, 
dbo.VStoreShelfs.RackNo, dbo.VStoreShelfs.StoreName, dbo.VItems.PriceForCost, dbo.VItems.FillingPrice,
ISNULL(TMakerPO.MaxRate,0) AS ItemMaxRate
,VItems.Parts
FROM            dbo.VStoreShelfs RIGHT OUTER JOIN
                             (SELECT        ItemID, SUM(NetQty) AS InHand, MAX(Shelf_RefID) AS Shelf_RefID
                                FROM            dbo.VItemsWithShelfWiseStock_Simple
                                GROUP BY ItemID) AS TFinished ON dbo.VStoreShelfs.EntryID = TFinished.Shelf_RefID RIGHT OUTER JOIN
                         dbo.VItems LEFT OUTER JOIN
                             (SELECT        ItemID, SUM(Qty) AS SFStock
                                FROM            dbo.VItemSFStock
                                GROUP BY ItemID) AS TSF ON dbo.VItems.ItemID = TSF.ItemID LEFT OUTER JOIN
                             (SELECT        dbo.ItemsRMComp.ItemID, SUM(dbo.VRM.QtyInStock) AS ForgingStock
                                FROM            dbo.ItemsRMComp INNER JOIN
                                                         dbo.VRM ON dbo.VRM.RMID = dbo.ItemsRMComp.RMID
                                GROUP BY dbo.ItemsRMComp.ItemID) AS TForging ON dbo.VItems.ItemID = TForging.ItemID ON TFinished.ItemID = dbo.VItems.ItemID LEFT OUTER JOIN
                             (SELECT        dbo.VRunningLots_Simple_PHF.ItemCode, SUM(dbo.VRunningLots_Simple_PHF.Qty) AS InProductionQty
                                FROM            dbo.VRunningLots_Simple_PHF INNER JOIN
                                                         dbo.FCustomerOrders ON dbo.VRunningLots_Simple_PHF.OrderNo = dbo.FCustomerOrders.OrderNo
                                WHERE        (dbo.FCustomerOrders.CustCode IN ('Stock'))
                                GROUP BY dbo.VRunningLots_Simple_PHF.ItemCode) AS TRunning ON dbo.VItems.ItemID = TRunning.ItemCode
LEFT JOIN (SELECT ItemCode,MAX(IssRate) As MaxRate FROM VVendRcvdDetail GROUP BY ItemCode) TMakerPO ON 
VItems.ItemID=TMakerPO.ItemCode
GO


