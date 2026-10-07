USE [Impulse_AWM]
GO

/****** Object:  View [dbo].[VItemsRMComp]    Script Date: 10/7/2026 7:17:11 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

ALTER VIEW [dbo].[VItemsRMComp]
AS
SELECT        dbo.ItemsRMComp.EntryID, dbo.ItemsRMComp.ItemID, dbo.ItemsRMComp.RMID, dbo.ItemsRMComp.Qty, dbo.ItemsRMComp.ProcessID, dbo.VRM.RMName AS Description, dbo.Processes.Description AS ProcDesc, 
                         dbo.VRM.Description AS GroupDescription, dbo.Items.ItemName, dbo.VRM.RMID1, dbo.VRM.GroupID, dbo.VRM.QtyInStock,dbo.ItemsRMComp.Functional_Status
                         ,VRM.GroupName,VRM.Unit,VRM.InActive
FROM            dbo.ItemsRMComp INNER JOIN
                         dbo.VRM ON dbo.ItemsRMComp.RMID = dbo.VRM.RMID INNER JOIN
                         dbo.Items ON dbo.ItemsRMComp.ItemID = dbo.Items.ItemID LEFT OUTER JOIN
                         dbo.Processes ON dbo.ItemsRMComp.ProcessID = dbo.Processes.ProcessID
GO


