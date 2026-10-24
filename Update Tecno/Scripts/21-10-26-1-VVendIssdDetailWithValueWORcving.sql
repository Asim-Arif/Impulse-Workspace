USE [Impulse_AWM]
GO

/****** Object:  View [dbo].[VVendIssdDetailWithValueWORcving]    Script Date: 10/21/2026 10:23:19 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


ALTER   VIEW [dbo].[VVendIssdDetailWithValueWORcving]
AS
SELECT        dbo.VendIssued.EntryID AS MainEntryID, dbo.VendIssued.RecieptID AS MainRecieptID, dbo.VendIssued.DT, dbo.VendIssdDetail.EntryID, dbo.VendIssued.VendID, 
                         dbo.VendIssdDetail.RefID, dbo.VendIssdDetail.RecieptID, dbo.VendIssdDetail.ItemCode, dbo.VendIssdDetail.Rate, dbo.VendIssdDetail.IssQty, 
                         dbo.VendIssdDetail.RcvdQty, dbo.VendIssdDetail.LotNo, dbo.VendIssdDetail.ReqAuth, dbo.VendIssdDetail.OrderNo, dbo.VendIssdDetail.RcvProcessID, 
                         dbo.VendIssdDetail.ReturnDT, dbo.VendIssdDetail.Priority, dbo.VendIssdDetail.SpecialInstructions, dbo.VendIssued.VchrNo, dbo.Items.ForgingWeight, 
                         VendIssdDetail.IssQty * VendIssdDetail.Rate AS IssValue, 
                         (VendIssdDetail.IssQty - VendIssdDetail.RcvdQty) * VendIssdDetail.Rate AS BalanceValue
                         , dbo.VendIssdDetail.ReWorkLot, 
                         dbo.VendIssdDetail.Repair_RefID, dbo.FCustomerOrders.CustCode, dbo.FCustomerOrders.InternalRefNo, dbo.VID_Bookmarks.EntryID AS BookMarkEntryID, 
                         T1.ItemCode AS CustItemCode, 
			T1.Quality, dbo.VIssuanceEntryWithWastage.WastageQty,T1.OrderQty,VendIssdDetail.Batch_No
FROM            dbo.VendIssdDetail INNER JOIN
                         dbo.VendIssued ON dbo.VendIssdDetail.RefID = dbo.VendIssued.EntryID INNER JOIN
                         dbo.Items ON dbo.VendIssued.ItemID = dbo.Items.ItemID INNER JOIN
                         dbo.FCustomerOrders ON dbo.VendIssdDetail.OrderNo = dbo.FCustomerOrders.OrderNo LEFT OUTER JOIN
                         dbo.VID_Bookmarks ON dbo.VendIssdDetail.EntryID = dbo.VID_Bookmarks.VID_RefID LEFT OUTER JOIN
                             (SELECT        OrderNo, ItemCode,CompItemCode, MAX(Quality) AS Quality,SUM(Qty) AS OrderQty
                                FROM            dbo.FOrderItems
                                GROUP BY OrderNo, ItemCode,CompItemCode) AS T1 ON dbo.VendIssdDetail.OrderNo = T1.OrderNo AND 
                         dbo.VendIssdDetail.ItemCode = T1.CompItemCode LEFT OUTER JOIN
                         dbo.VIssuanceEntryWithWastage ON dbo.VendIssdDetail.EntryID = dbo.VIssuanceEntryWithWastage.EntryID

GO


