USE [Impulse_AWM]
GO

/****** Object:  View [dbo].[VGatePass_Return_Details]    Script Date: 11/1/2026 8:01:12 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO


CREATE OR ALTER   VIEW [dbo].[VGatePass_Return_Details]
AS
SELECT     dbo.GatePass_Return_Details.EntryID, dbo.GatePass_Return_Details.MEGO_RefID, dbo.GatePass_Return_Details.Detail, 
                      dbo.GatePass_Return_Details.Unit, dbo.GatePass_Return_Details.Packing, dbo.GatePass_Return_Details.Qty, ISNULL(T1.QtyRcvd, 0) 
                      AS TotalQtyRcvd
FROM         dbo.GatePass_Return_Details LEFT OUTER JOIN
                          (SELECT     ME_GO_D_RefID, SUM(QtyRcvd) AS QtyRcvd
                             FROM         dbo.GatePass_Return_Receiving_Detail
                             GROUP BY ME_GO_D_RefID) AS T1 ON dbo.GatePass_Return_Details.EntryID = T1.ME_GO_D_RefID
GO


