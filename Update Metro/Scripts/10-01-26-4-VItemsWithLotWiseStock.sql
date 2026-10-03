USE [Impulse_AWM]
GO

/****** Object:  View [dbo].[VItemsWithLotWiseStock]    Script Date: 10/1/2026 6:36:19 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


ALTER VIEW [dbo].[VItemsWithLotWiseStock]
AS
SELECT dbo.VStoreShelfs.StoreName, dbo.VStoreShelfs.RackNo, dbo.VStoreShelfs.ShelfNo, T1.ItemID, T1.LotNo, T1.RcvdQty - ISNULL(T1.IssdQty, 0) AS NetQty, T1.Shelf_RefID, T1.Remarks, dbo.VStoreShelfs.Store_RefID, T1.EntryID
    ,CASE WHEN ISNULL(Lots_List.Batch_No,'')='' THEN 'No-Batch' ELSE ISNULL(Lots_List.Batch_No,'') END  AS Batch_No,ISNULL(Lots_List.Mill_Certificate_No,'') AS Mill_Certificate_No
FROM  dbo.VStoreShelfs INNER JOIN
             (SELECT dbo.RcvItemsSimpleDetail.EntryID, dbo.RcvItemsSimpleDetail.ItemID, ISNULL(dbo.vendRcvdDetail.LotNo, dbo.RcvItemsSimpleDetail.LotNo_Manual) AS LotNo, dbo.RcvItemsSimpleDetail_Placement.Shelf_RefID, dbo.RcvItemsSimpleDetail_Placement.RcvdQty, ISNULL(T1_1.IssdQty, 0) AS IssdQty, dbo.RcvItemsSimpleDetail_Placement.RcvdQty - ISNULL(T1_1.IssdQty, 0) 
                     AS NetQty, dbo.RcvItemsSimpleDetail.Remarks
            FROM  dbo.RcvItemsSimpleDetail INNER JOIN
                     dbo.RcvItemsSimpleDetail_Placement ON dbo.RcvItemsSimpleDetail.EntryID = dbo.RcvItemsSimpleDetail_Placement.RISD_RefID LEFT OUTER JOIN
                     dbo.TransferredToReadyFinishLotsDetail ON dbo.RcvItemsSimpleDetail.TTRFLD_RefID = dbo.TransferredToReadyFinishLotsDetail.EntryID LEFT OUTER JOIN
                     dbo.vendRcvdDetail ON dbo.TransferredToReadyFinishLotsDetail.VRD_RefID = dbo.vendRcvdDetail.EntryID LEFT OUTER JOIN
                         (SELECT RCV_ISD_RefID, SUM(Qty) AS IssdQty
                        FROM  dbo.IssItemsSimpleDetail
                        GROUP BY RCV_ISD_RefID) AS T1_1 ON dbo.RcvItemsSimpleDetail.EntryID = T1_1.RCV_ISD_RefID) AS T1 ON T1.Shelf_RefID = dbo.VStoreShelfs.EntryID
                        LEFT OUTER JOIN Lots_List ON T1.LotNo=Lots_List.LotNo
GO


