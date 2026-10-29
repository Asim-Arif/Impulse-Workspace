
-- ============================================================================
-- Fix VSFStockTransaction View: Ensure LotNo is populated for manual SF open receiving
-- Fallback to StockOrderOpening.LotNo_Manual when VendRcvdDetail.LotNo is null
-- ============================================================================
CREATE OR ALTER VIEW [dbo].[VSFStockTransaction]
AS
SELECT        dbo.StockOrderOpening.EntryID, dbo.VendIssued.UserName, dbo.VendIssued.DT, dbo.VendIssdDetail.ItemCode, dbo.Items.ItemName, dbo.StockOrderOpening.ProcessID, dbo.Processes.Description, dbo.StockOrderOpening_Issuance.QtyIssd, 1 AS Issuance, 
                         Location, StoreName, Store_RefID, ISNULL(NULLIF(VendIssdDetail.LotNo, ''), StockOrderOpening.LotNo_Manual) AS LotNo
FROM            dbo.StockOrderOpening_Issuance INNER JOIN
                         dbo.StockOrderOpening ON dbo.StockOrderOpening_Issuance.SOO_RefID = dbo.StockOrderOpening.EntryID INNER JOIN
                         dbo.VendIssdDetail ON dbo.VendIssdDetail.EntryID = dbo.StockOrderOpening_Issuance.VID_RefID INNER JOIN
                         dbo.VendIssued ON dbo.VendIssued.EntryID = dbo.VendIssdDetail.RefID INNER JOIN
                         dbo.Items ON dbo.VendIssdDetail.ItemCode = dbo.Items.ItemID INNER JOIN
                         dbo.Processes ON dbo.StockOrderOpening.ProcessID = dbo.Processes.ProcessID INNER JOIN
                         VStoreShelfs ON VStoreShelfs.EntryID = StockOrderOpening.Shelf_RefID
UNION
SELECT        StockOrderOpening.EntryID, UserName, CAST(CONVERT(VARCHAR(50), DTEntry, 6) AS DATETIME), Items.ItemID, ItemName, StockOrderOpening.ProcessID, Processes.Description, StockOrderOpening.Qty, 0 AS Issuance, Location, StoreName, Store_RefID, 
                         ISNULL(NULLIF(VendRcvdDetail.LotNo, ''), StockOrderOpening.LotNo_Manual) AS LotNo
FROM            StockOrderOpening INNER JOIN
                         Items ON StockOrderOpening.ItemID = Items.ItemID INNER JOIN
                         Processes ON StockOrderOpening.ProcessID = Processes.ProcessID INNER JOIN
                         VStoreShelfs ON VStoreShelfs.EntryID = StockOrderOpening.Shelf_RefID LEFT OUTER JOIN
                         LotTransferDetails ON StockOrderOpening.EntryID = LotTransferDetails.SOO_RefID LEFT OUTER JOIN
                         VendRcvdDetail ON LotTransferDetails.VRD_From_RefID = VendRcvdDetail.EntryID
GO