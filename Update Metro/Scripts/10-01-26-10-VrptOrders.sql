USE [Impulse_AWM]
GO

/****** Object:  View [dbo].[VrptOrders]    Script Date: 10/2/2026 4:13:28 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


ALTER VIEW [dbo].[VrptOrders]
AS
SELECT dbo.FOrderItems.ID, dbo.FOrderItems.ItemCode, dbo.FOrderItems.Price, dbo.FOrderItems.Qty, dbo.FOrderItems.Qty - ISNULL(TProforma.ProformaUsedQty, 0) AS InvQty, dbo.Items.ItemName, dbo.FCustomerCatalog.Description, dbo.FCustomerCatalog.ItemID, dbo.FCustomerCatalog.CompItemID, dbo.ForeignCustomers.Curr, dbo.FCustomerCatalog.Unit, 
         dbo.ForeignCustomers.Name, dbo.ForeignCustomers.Address, dbo.FOrderItems.CustomPrice, dbo.Items.CustomDescription, dbo.Items.ItemSize, dbo.Items.SizeUnit, dbo.SteelTypes.SteelType, dbo.SteelGages.Gage, dbo.Items.ItemColor, dbo.FCustomerCatalog.BarcodeNo, dbo.FCustomerCatalog.PackingMode, dbo.Items.MasterCartonL, dbo.Items.MasterCartonW, 
         dbo.Items.MasterCartonH, dbo.Items.SmallCartonL, dbo.Items.SmallCartonW, dbo.Items.SmallCartonH, dbo.Items.PolyBag, dbo.FCustomerCatalog.SpecialInstructions, dbo.FCustomerCatalog.StampInstructions, dbo.FCustomerCatalog.PackingInstructions, dbo.Items.Unit AS UnitItems, dbo.FCustomerCatalog.ItemColor AS CustomerColor, ISNULL(Shipped.TotalShipped, 0) 
         AS ShippedQty, dbo.Items.ItemNameUrdu, dbo.Items.InHand, dbo.FOrderItems.SortNo, dbo.ItemCatagories.MaxLotSize, dbo.FOrderItems.DeliveryDT AS DeliveryDTItem, dbo.FOrderItems.Stamps AS StampsItem, dbo.FOrderItems.Quality AS QualityItem, dbo.Items.FinQuality, dbo.Items.Type, dbo.FOrderItems.DeliveryStatus, dbo.FCustomerCatalog.OCR, 
         dbo.FOrderItems.CompItemCode, dbo.ItemGroups.Description AS GroupDescription, dbo.FCustomerOrders.OrderNo, dbo.FCustomerOrders.DT, dbo.FCustomerOrders.TradeTerms, dbo.FCustomerOrders.CustCode, dbo.FCustomerOrders.Country, dbo.FCustomerOrders.PartialShipment, dbo.FCustomerOrders.PaymentTerms, dbo.FCustomerOrders.TransShipment, 
         dbo.FCustomerOrders.Packaging, dbo.FCustomerOrders.DeliveryDT, dbo.FCustomerOrders.BatchNo, dbo.FCustomerOrders.CompanyRefID, dbo.FCustomerOrders.StampDT, dbo.FCustomerOrders.Quality, dbo.FCustomerOrders.InternalRefNo, dbo.FCustomerOrders.OrderRcvdVia, dbo.FOrderItems.Remarks, dbo.FCustomerOrders.OrderType, dbo.Items.GroupID, T1.FOI_RefID, 
         dbo.FCustomerFinalOrders.Cancelled, dbo.Items.MasterCartonSmallBoxes, dbo.Items.SmallBoxPcs, dbo.Items.FinishedWeight, dbo.FCustomerOrders.OrderRevisionNo, dbo.FCustomerOrders.OrderRevisionDT, dbo.FOrderItems.Item_Edited, dbo.FCustomerOrders.ShippingMode, dbo.FOrderItems.Weight, dbo.FOrderItems.Item_Finishing_Type, 
         dbo.Item_Finishing_Type.Item_Finishing_Type_Text
         , dbo.FOrderItems.IW_OrderNo, dbo.FOrderItems.IW_BatchNo, dbo.Items.ItemPic
				  ,dbo.FCustomerOrders.Packaging_Weight,ReadyFinishPrice,Fillingprice,PriceForCost
FROM  dbo.FCustomerOrders INNER JOIN
         dbo.FOrderItems ON dbo.FCustomerOrders.OrderNo = dbo.FOrderItems.OrderNo INNER JOIN
         dbo.FCustomerCatalog ON dbo.FCustomerOrders.CustCode = dbo.FCustomerCatalog.CustCode AND dbo.FCustomerOrders.Country = dbo.FCustomerCatalog.Country AND dbo.FOrderItems.CompItemCode = dbo.FCustomerCatalog.CompItemID INNER JOIN
         dbo.Items ON dbo.FCustomerCatalog.CompItemID = dbo.Items.ItemID INNER JOIN
         dbo.ForeignCustomers ON dbo.FCustomerOrders.CustCode = dbo.ForeignCustomers.CustCode AND dbo.FCustomerOrders.Country = dbo.ForeignCustomers.Country INNER JOIN
         dbo.ItemCatagories ON dbo.Items.CatID = dbo.ItemCatagories.CatID INNER JOIN
         dbo.ItemGroups ON dbo.Items.GroupID = dbo.ItemGroups.ID LEFT OUTER JOIN
         dbo.Item_Finishing_Type ON dbo.FOrderItems.Item_Finishing_Type = dbo.Item_Finishing_Type.EntryID LEFT OUTER JOIN
         dbo.FCustomerFinalOrders ON dbo.FCustomerOrders.OrderNo = dbo.FCustomerFinalOrders.OrderNo LEFT OUTER JOIN
         dbo.SteelGages ON dbo.Items.Gage = dbo.SteelGages.GageID LEFT OUTER JOIN
         dbo.SteelTypes ON dbo.Items.SteelUsed = dbo.SteelTypes.SteelID LEFT OUTER JOIN
             (SELECT FOI_RefID
            FROM  dbo.FOrderItems_PDF
            GROUP BY FOI_RefID) AS T1 ON dbo.FOrderItems.ID = T1.FOI_RefID LEFT OUTER JOIN
             (SELECT OrderEntryID, SUM(Qty) AS ProformaUsedQty
            FROM  dbo.FProformaOrders
            GROUP BY OrderEntryID) AS TProforma ON dbo.FOrderItems.ID = TProforma.OrderEntryID LEFT OUTER JOIN
             (SELECT FProformaOrders_1.OrderEntryID, SUM(ISNULL(dbo.CustomInvoiceItems.Qty, 0)) AS TotalShipped
            FROM  dbo.CustomInvoiceItems INNER JOIN
                     dbo.CustomInvoice ON dbo.CustomInvoiceItems.CustomInvoice = dbo.CustomInvoice.CustomInvoice INNER JOIN
                     dbo.FProformaOrders AS FProformaOrders_1 ON dbo.CustomInvoiceItems.RefID = FProformaOrders_1.EntryID
            WHERE (dbo.CustomInvoice.GatePassDT IS NOT NULL)
            GROUP BY FProformaOrders_1.OrderEntryID) AS Shipped ON dbo.FOrderItems.ID = Shipped.OrderEntryID
GO


