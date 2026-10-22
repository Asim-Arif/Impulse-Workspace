USE [Impulse_AWM]
GO

/****** Object:  View [dbo].[VItems]    Script Date: 10/22/2026 7:13:17 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


ALTER   VIEW [dbo].[VItems]
AS
SELECT dbo.Items.CreateDT, dbo.Items.CatID, dbo.Items.ItemID, dbo.Items.ItemName, dbo.Items.Unit, dbo.Items.Type, dbo.Items.ItemSize, dbo.Items.SizeUnit, dbo.Items.ItemUsage, dbo.Items.Attributes, dbo.Items.FinQuality, dbo.Items.SteelUsed, dbo.Items.Gage, dbo.Items.PacknLabel, dbo.Items.AssetAccNo, dbo.Items.SaleAccNo, dbo.Items.SRTAccNo, dbo.Items.FOB, dbo.Items.CIFSea, 
         dbo.Items.CIFAir, dbo.Items.CnFSea, dbo.Items.CnFAir, dbo.Items.ExWorks, dbo.Items.CnISea, dbo.Items.CnIAir, dbo.Items.ItemCurr, dbo.Items.UnitWeight, dbo.Items.ForgingWeight, dbo.Items.WasteVisible, dbo.Items.FinishedWeight, dbo.Items.InActive, dbo.Items.ItemPic, dbo.Items.RevID, dbo.Items.GroupID, dbo.Items.PlantRate, dbo.Items.SnaffRate, dbo.Items.StampRate, 
         dbo.Items.ReorderPoint, dbo.Items.FinishDescription1, dbo.Items.FinishDescription2, dbo.Items.MakerDescription1, dbo.Items.MakerDescription2, dbo.Items.Tagging, dbo.Items.EAN128, dbo.Items.CustomDescription, dbo.Items.ItemColor, dbo.Items.PackingInstructions, dbo.Items.MasterCartonL, dbo.Items.MasterCartonW, dbo.Items.MasterCartonH, dbo.Items.SmallCartonL, 
         dbo.Items.SmallCartonW, dbo.Items.SmallCartonH, dbo.Items.PolyBag, dbo.Items.FixedPackingUnit, dbo.VItemWithStocks.TotalFinishedStock AS InHand, dbo.Items.OpenBal, dbo.Items.ItemNameUrdu, dbo.Items.PriceForCost, dbo.Items.TechnicalDrawing, dbo.Items.TipSize, dbo.Items.MinLevel, dbo.Items.MaxLevel, dbo.Items.ReOrderLevel, dbo.Items.ItemType, dbo.Items.FOBTop, 
         dbo.Items.CIFSeaTop, dbo.Items.CIFAirTop, dbo.Items.CnFSeaTop, dbo.Items.CnFAirTop, dbo.Items.ExWorksTop, dbo.Items.CnISeaTop, dbo.Items.CnIAirTop, dbo.Items.GTINBarcodeNo, dbo.Items.Description, dbo.Items.UMDNSCode, dbo.Items.FDAListingNo, dbo.Items.FDAProductCode, dbo.Items.EuropeanRegNo, dbo.Items.FDA510K, dbo.Items.ItemMaxLotSize, 
         dbo.Items.ItemLotSizeBuffer, dbo.Items.ReadyFinishPrice, dbo.Items.FillingPrice, dbo.Items.MainGroupID, dbo.Items.Sample, dbo.Items.Type AS Typ1, Accounts_3.AccTitle AS AssetAccNoText, Accounts_1.AccTitle AS SaleAccNoText, Accounts_2.AccTitle AS SRTAccNoText, dbo.SteelGages.Gage AS SteelGage, dbo.SteelTypes.SteelType, dbo.ItemGroups.Description AS ItemGroup, 
         dbo.ItemGroups.GrpColor, dbo.ItemCatagories.Description AS Category, dbo.ItemGroupsMain.MainGroupName, dbo.Items.SmallBoxPcs, dbo.Items.MasterCartonSmallBoxes, dbo.Items.AvailableForECommerce, dbo.Items.PicForECommerce, dbo.Items.DescriptionForECommerce, dbo.Items.PriceForECommerce, dbo.Items.GroupIDForECommerce, 
         dbo.ItemGroups_ECommerce.ECommerceGroupName, dbo.Items.POInustructions, dbo.Items.SFDA_Name, dbo.Items.SFDA_No, dbo.Items.SFDA_Listing, dbo.Items.HRC_From, dbo.Items.HRC_To,
         dbo.Items.Fig_No, dbo.Items.Scale, dbo.Items.Additional_Detail, dbo.SteelGages.GageUnit, dbo.Items.MakerDescription
         ,0 AS Parts
FROM  dbo.SteelTypes RIGHT OUTER JOIN
         dbo.Items INNER JOIN
         dbo.ItemCatagories ON dbo.Items.CatID = dbo.ItemCatagories.CatID LEFT OUTER JOIN
         dbo.ItemGroups_ECommerce ON dbo.Items.GroupIDForECommerce = dbo.ItemGroups_ECommerce.ID LEFT OUTER JOIN
         dbo.ItemGroupsMain ON dbo.Items.MainGroupID = dbo.ItemGroupsMain.MainGroupID LEFT OUTER JOIN
         dbo.ItemGroups ON dbo.Items.GroupID = dbo.ItemGroups.ID ON dbo.SteelTypes.SteelID = dbo.Items.SteelUsed LEFT OUTER JOIN
         dbo.SteelGages ON dbo.Items.Gage = dbo.SteelGages.GageID LEFT OUTER JOIN
         dbo.Accounts AS Accounts_2 ON dbo.Items.SRTAccNo = Accounts_2.AccNo LEFT OUTER JOIN
         dbo.Accounts AS Accounts_1 ON dbo.Items.SaleAccNo = Accounts_1.AccNo LEFT OUTER JOIN
         dbo.Accounts AS Accounts_3 ON dbo.Items.AssetAccNo = Accounts_3.AccNo LEFT OUTER JOIN
         dbo.VItemWithStocks ON dbo.Items.ItemID = dbo.VItemWithStocks.ItemID

GO


