

/****** Object:  View [dbo].[VMakerBilling_Receivings_Hidden]    Script Date: 12/25/2024 9:35:28 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE VIEW [dbo].[VMakerBilling_Receivings_Hidden]
AS
SELECT        dbo.MakerBilling_Receivings_Hidden.EntryID, dbo.MakerBilling_Receivings_Hidden.VRD_RefID, dbo.VVendReceivingList_Simple.DT, dbo.VVendReceivingList_Simple.ProcessID, dbo.VVendReceivingList_Simple.ItemCode, 
                         dbo.VVendReceivingList_Simple.RcvdQty, dbo.VVendReceivingList_Simple.VenderName, dbo.VVendReceivingList_Simple.VendID1, dbo.VVendReceivingList_Simple.OnlyDT, dbo.VVendReceivingList_Simple.Description, 
                         dbo.VVendReceivingList_Simple.LotNo, dbo.VVendReceivingList_Simple.OrderNo, dbo.VVendReceivingList_Simple.CompItemCode, dbo.VVendReceivingList_Simple.ReWorkLot, dbo.VVendReceivingList_Simple.ItemName, 
                         dbo.VVendReceivingList_Simple.ItemSize, dbo.VVendReceivingList_Simple.SizeUnit, dbo.VVendReceivingList_Simple.Type, dbo.VVendReceivingList_Simple.CustCode, dbo.VVendReceivingList_Simple.VendID
FROM            dbo.MakerBilling_Receivings_Hidden INNER JOIN
                         dbo.VVendReceivingList_Simple ON dbo.MakerBilling_Receivings_Hidden.VRD_RefID = dbo.VVendReceivingList_Simple.VRD_EntryID
GO

