USE [Impulse_AWM]
GO

/****** Object:  View [dbo].[VFProformaList]    Script Date: 11/24/2026 11:57:42 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


ALTER VIEW [dbo].[VFProformaList]
AS
SELECT     dbo.ForeignCustomers.Curr, dbo.FPInvoice.PInvoice, dbo.FPInvoice.DT, dbo.ForeignCustomers.CustCode, dbo.ForeignCustomers.Country, 
                      dbo.F_ProformaAmt(dbo.FPInvoice.PInvoice) AS ProformaAmt, dbo.ForeignCustomers.Name, dbo.FPInvoice.CompanyRefID, dbo.Companies.CompanyName, 
                      dbo.F_ProformaOrders(dbo.FPInvoice.PInvoice) AS Orders, TabOCs.ChargesAmt, dbo.F_ProformaInternalOrders(dbo.FPInvoice.PInvoice) AS InternalOrders, 
                      dbo.FPInvoice.AdvanceDescription, dbo.FPInvoice.AdvanceAmount
                      ,dbo.FPInvoice.DiscountAmount,dbo.FPInvoice.Insurance_Freight_Amount,dbo.FPInvoice.Insurance_Freight_Amount_II,dbo.FPInvoice.DiscountPer
FROM         dbo.FPInvoice INNER JOIN
                      dbo.ForeignCustomers ON dbo.FPInvoice.CustCode = dbo.ForeignCustomers.CustCode AND dbo.FPInvoice.Country = dbo.ForeignCustomers.Country INNER JOIN
                      dbo.Companies ON dbo.FPInvoice.CompanyRefID = dbo.Companies.EntryID LEFT OUTER JOIN
                          (SELECT     PInvoice, SUM(ChargesAmt) AS ChargesAmt
                             FROM         FPInvoiceOtherCharges
                             GROUP BY PInvoice) TabOCs ON dbo.FPInvoice.PInvoice = TabOCs.PInvoice

GO


