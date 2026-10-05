USE [Impulse_AWM]
GO

/****** Object:  View [dbo].[VVenderAssItems]    Script Date: 10/5/2026 7:36:53 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


ALTER VIEW [dbo].[VVenderAssItems]
AS
SELECT     dbo.VenderAssItems.VendID, dbo.VenderAssItems.RMID, dbo.VenderAssItems.Rate, dbo.RM.RMID1, dbo.RM.RMName, dbo.RM.Unit, 
                      dbo.RMGroups.Description AS GroupName, dbo.VenderAssItems.EntryID
                      ,dbo.VenderAssItems.Remarks
                      ,Accounts.AccTitle
FROM         dbo.VenderAssItems INNER JOIN
                      dbo.RM ON dbo.VenderAssItems.RMID = dbo.RM.RMID INNER JOIN
                      dbo.RMGroups ON dbo.RM.GroupID = dbo.RMGroups.ID
                      INNER JOIN Accounts ON dbo.VenderAssItems.VendID=Accounts.AccNo
GO


