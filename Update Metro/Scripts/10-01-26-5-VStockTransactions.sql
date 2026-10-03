USE [Impulse_AWM]
GO

/****** Object:  View [dbo].[VStockTransactions]    Script Date: 10/1/2026 6:51:41 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


ALTER VIEW [dbo].[VStockTransactions]
AS
SELECT     StockTransactions.EntryNo, StockTransactions.DT, StockTransactions.OrderNo, StockTransactions.RcvdBy, StockTransactions.RcvdFrom, StockTransactions.ItemID, StockTransactions.Qty, StockTransactions.Location, StockTransactions.ChildEntryID, 
                  StockTransactions.EntryType, StockTransactions.Shelf_RefID, dbo.Items.ItemName, dbo.VStoreShelfs.Store_RefID, StockTransactions.SampleEntry, StockTransactions.LotNo_Manual, dbo.Items.ItemPic, StockTransactions.Remarks, dbo.VStoreShelfs.StoreName, 
                  dbo.VStoreShelfs.RackNo, dbo.VStoreShelfs.ShelfNo, dbo.Items.ItemNameUrdu, dbo.Items.Unit,StockTransactions.TransferRemarks
FROM        dbo.Items INNER JOIN
                      (SELECT     dbo.RcvItemsSimple.RcvNo AS EntryNo, CAST(CONVERT(VARCHAR(50), dbo.RcvItemsSimple.DT, 6) AS DATETIME) AS DT, dbo.RcvItemsSimple.OrderNo, dbo.RcvItemsSimple.RcvdBy, dbo.RcvItemsSimple.RcvdFrom, dbo.RcvItemsSimpleDetail.ItemID, 
                                         dbo.RcvItemsSimpleDetail.Qty, dbo.RcvItemsSimpleDetail.Location, dbo.RcvItemsSimpleDetail.EntryID AS ChildEntryID, 1 AS EntryType, dbo.RcvItemsSimpleDetail_Placement.Shelf_RefID, 0 AS SampleEntry, dbo.RcvItemsSimpleDetail.LotNo_Manual, 
                                         dbo.RcvItemsSimpleDetail.Remarks,
										 dbo.RcvItemsSimpleDetail_Placement.Remarks TransferRemarks
                       FROM        dbo.RcvItemsSimple INNER JOIN
                                         dbo.RcvItemsSimpleDetail ON dbo.RcvItemsSimple.RcvNo = dbo.RcvItemsSimpleDetail.RcvNo INNER JOIN
                                         dbo.RcvItemsSimpleDetail_Placement ON dbo.RcvItemsSimpleDetail.EntryID = dbo.RcvItemsSimpleDetail_Placement.RISD_RefID
                       UNION
                       SELECT     dbo.IssItemsSimple.IssNo AS EntryNo, CAST(CONVERT(VARCHAR(50), dbo.IssItemsSimple.DT, 6) AS DATETIME) AS DT, dbo.IssItemsSimple.OrderNo, dbo.IssItemsSimple.Department, dbo.IssItemsSimple.AttnPerson, dbo.IssItemsSimpleDetail.ItemID, 
                                         dbo.IssItemsSimpleDetail_More.IssdQty AS Qty, dbo.IssItemsSimpleDetail.Location, dbo.IssItemsSimpleDetail.EntryID AS ChildEntryID, 0 AS EntryType, dbo.IssItemsSimpleDetail_More.Shelf_RefID, 
                                         dbo.IssItemsSimple.SampleIssuance AS SampleEntry, dbo.IssItemsSimpleDetail.LotNo_Manual, 
										 dbo.IssItemsSimpleDetail.Remarks,
										 dbo.IssItemsSimpleDetail_More.Remarks  TransferRemarks
                       FROM        dbo.IssItemsSimple INNER JOIN
                                         dbo.IssItemsSimpleDetail ON dbo.IssItemsSimple.IssNo = dbo.IssItemsSimpleDetail.IssNo INNER JOIN
                                         dbo.IssItemsSimpleDetail_More ON dbo.IssItemsSimpleDetail.EntryID = dbo.IssItemsSimpleDetail_More.IISD_RefID
				  ) AS StockTransactions ON dbo.Items.ItemID = StockTransactions.ItemID LEFT OUTER JOIN
                  dbo.VStoreShelfs ON StockTransactions.Shelf_RefID = dbo.VStoreShelfs.EntryID
GO


