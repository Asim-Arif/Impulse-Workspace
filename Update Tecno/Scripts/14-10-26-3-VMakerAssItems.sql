USE [Impulse_AWM]
GO

/****** Object:  View [dbo].[VMakerAssItems]    Script Date: 10/14/2026 4:25:46 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


ALTER VIEW [dbo].[VMakerAssItems]
AS
SELECT        dbo.ItemGroups.Description, dbo.Makers.VenderName, dbo.Makers.PhaseID, dbo.Makers.VendType, dbo.Makers.RepairDedRate, dbo.VItems1.ItemName, dbo.VItems1.Unit, dbo.VItems1.Type, dbo.Makers.Active, 
                         dbo.Makers.VendID1, dbo.VendAssItems.EntryID, dbo.VendAssItems.VendID, dbo.VendAssItems.ProcessID, dbo.VendAssItems.ItemID, dbo.VendAssItems.Rate, dbo.VendAssItems.PlantRate, dbo.VendAssItems.StampRate, 
                         dbo.VendAssItems.SnaffRate, dbo.VendAssItems.Unit AS AssignedUnit, dbo.VendAssItems.Remarks, dbo.VItems1.ItemSize, dbo.VItems1.SizeUnit, dbo.VendAssItems.TopRate, dbo.VendAssItems.CustCode
FROM            dbo.VendAssItems INNER JOIN
                         dbo.Makers ON dbo.VendAssItems.VendID = dbo.Makers.VendID INNER JOIN
                         dbo.VItems1 ON dbo.VendAssItems.ItemID = dbo.VItems1.ItemID LEFT OUTER JOIN
                         dbo.ItemGroups ON dbo.VItems1.GroupID = dbo.ItemGroups.ID
GO


