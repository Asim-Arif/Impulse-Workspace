USE [Impulse_AWM]
GO

/****** Object:  View [dbo].[VMaterialLocationWiseStatus]    Script Date: 10/5/2026 2:30:23 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


ALTER VIEW [dbo].[VMaterialLocationWiseStatus]
AS
SELECT     dbo.MaterialLocationwiseStatus.EntryID, dbo.MaterialLocationwiseStatus.Rcvd_RefID, dbo.MaterialLocationwiseStatus.Shelf_RefID, dbo.MaterialLocationwiseStatus.QtyPlaced, dbo.MaterialLocationwiseStatus.QtyIssued, dbo.VendRcvdDetailPO.MaterialID, 
                  dbo.MaterialLocationwiseStatus.SheetsPlaced, dbo.MaterialLocationwiseStatus.SheetsIssued, dbo.RM.RMName, dbo.VVenders.AccTitle, dbo.VendRcvdDetailPO.QtyRcvd, dbo.MaterialLocationwiseStatus.BatchNo, dbo.MaterialLocationwiseStatus.LotNo, 
                  dbo.VStoreShelfs.StoreName, dbo.VStoreShelfs.RackNo, dbo.VStoreShelfs.ShelfNo, dbo.VendRcvdDetailPO.RcvID, dbo.VendRcvdDetailPO.PORefNo, dbo.VendRcvd.RcvDate, dbo.MaterialLocationwiseStatus.Mill_Certificate_No, dbo.VVenders.AccNo, dbo.RM.Pic, 
                  dbo.RM.Unit, dbo.RM.GroupID, dbo.RM.RMID1, dbo.RM.TechnicalDrawing, dbo.MaterialLocationwiseStatus.UserName, dbo.MaterialLocationwiseStatus.MachineName, dbo.VVenders.VendID, dbo.VVenders.VenderDescription, dbo.VVenders.ContactPerson, 
                  dbo.VVenders.CPhone
FROM        dbo.MaterialLocationwiseStatus INNER JOIN
                  dbo.VendRcvdDetailPO ON dbo.MaterialLocationwiseStatus.Rcvd_RefID = dbo.VendRcvdDetailPO.EntryID INNER JOIN
                  dbo.RM ON dbo.VendRcvdDetailPO.MaterialID = dbo.RM.RMID1 INNER JOIN
                  dbo.VendRcvd ON dbo.VendRcvdDetailPO.RcvID = dbo.VendRcvd.RcvID INNER JOIN
                  dbo.VStoreShelfs ON dbo.MaterialLocationwiseStatus.Shelf_RefID = dbo.VStoreShelfs.EntryID LEFT OUTER JOIN
                  dbo.VVenders ON dbo.VendRcvd.VendID = dbo.VVenders.AccNo
GO


