-- ====================================================================================================
-- Script Name: Prerequisite_Schema_Updates.sql
-- Description: Idempotent prerequisite schema updates required before deploying features and running
--              the sequential update scripts. Adds missing columns and synchronizes prerequisite views.
-- Target DB:   Works on SMBI_AWM, Impulse_AWM, or any production instance (database-agnostic).
-- ====================================================================================================

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

PRINT '>>> Starting Prerequisite Schema Updates on ' + DB_NAME() + '...';
GO

-- ----------------------------------------------------------------------------------------------------
-- 1. FCustomerCatalog: Missing columns for extended customer catalog & regulatory info
-- ----------------------------------------------------------------------------------------------------
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'FCustomerCatalog' AND COLUMN_NAME = 'MDMA')
BEGIN
    ALTER TABLE [dbo].[FCustomerCatalog] ADD [MDMA] VARCHAR(255) NULL;
    PRINT '  + Added FCustomerCatalog.MDMA';
END
GO

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'FCustomerCatalog' AND COLUMN_NAME = 'SFDA_Listing_No')
BEGIN
    ALTER TABLE [dbo].[FCustomerCatalog] ADD [SFDA_Listing_No] VARCHAR(255) NULL;
    PRINT '  + Added FCustomerCatalog.SFDA_Listing_No';
END
GO

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'FCustomerCatalog' AND COLUMN_NAME = 'MD_Group')
BEGIN
    ALTER TABLE [dbo].[FCustomerCatalog] ADD [MD_Group] VARCHAR(50) NULL;
    PRINT '  + Added FCustomerCatalog.MD_Group';
END
GO

-- ----------------------------------------------------------------------------------------------------
-- 2. VoucherInfo: Missing handover tracking column for vouchers
-- ----------------------------------------------------------------------------------------------------
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'VoucherInfo' AND COLUMN_NAME = 'Handed_Over_To')
BEGIN
    ALTER TABLE [dbo].[VoucherInfo] ADD [Handed_Over_To] VARCHAR(255) NULL;
    PRINT '  + Added VoucherInfo.Handed_Over_To';
END
GO

-- ----------------------------------------------------------------------------------------------------
-- 3. Makers: Extended company, secondary contact, and CNIC document columns
-- ----------------------------------------------------------------------------------------------------
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'Makers' AND COLUMN_NAME = 'CompanyName')
BEGIN
    ALTER TABLE [dbo].[Makers] ADD [CompanyName] VARCHAR(255) NULL;
    PRINT '  + Added Makers.CompanyName';
END
GO

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'Makers' AND COLUMN_NAME = 'Maker_Second_Name')
BEGIN
    ALTER TABLE [dbo].[Makers] ADD [Maker_Second_Name] VARCHAR(255) NULL;
    PRINT '  + Added Makers.Maker_Second_Name';
END
GO

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'Makers' AND COLUMN_NAME = 'CNIC_PDF')
BEGIN
    ALTER TABLE [dbo].[Makers] ADD [CNIC_PDF] IMAGE NULL;
    PRINT '  + Added Makers.CNIC_PDF';
END
GO

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'Makers' AND COLUMN_NAME = 'CNIC_PDF_FileName')
BEGIN
    ALTER TABLE [dbo].[Makers] ADD [CNIC_PDF_FileName] VARCHAR(1000) NULL;
    PRINT '  + Added Makers.CNIC_PDF_FileName';
END
GO

-- ----------------------------------------------------------------------------------------------------
-- 4. FCustomerOrders: Order authorization & packaging metadata
-- ----------------------------------------------------------------------------------------------------
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'FCustomerOrders' AND COLUMN_NAME = 'Packaging_Weight')
BEGIN
    ALTER TABLE [dbo].[FCustomerOrders] ADD [Packaging_Weight] VARCHAR(4000) NULL;
    PRINT '  + Added FCustomerOrders.Packaging_Weight';
END
GO

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'FCustomerOrders' AND COLUMN_NAME = 'Authorized')
BEGIN
    ALTER TABLE [dbo].[FCustomerOrders] ADD [Authorized] BIT NOT NULL CONSTRAINT [DF_FCustomerOrders_Authorized] DEFAULT 0;
    PRINT '  + Added FCustomerOrders.Authorized';
END
GO

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'FCustomerOrders' AND COLUMN_NAME = 'AuthorizedBy')
BEGIN
    ALTER TABLE [dbo].[FCustomerOrders] ADD [AuthorizedBy] VARCHAR(50) NULL;
    PRINT '  + Added FCustomerOrders.AuthorizedBy';
END
GO

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'FCustomerOrders' AND COLUMN_NAME = 'AuthorizedDT')
BEGIN
    ALTER TABLE [dbo].[FCustomerOrders] ADD [AuthorizedDT] DATETIME NULL;
    PRINT '  + Added FCustomerOrders.AuthorizedDT';
END
GO

-- ----------------------------------------------------------------------------------------------------
-- 5. VendReceived & VendIssued: Entry timestamp audit columns
-- ----------------------------------------------------------------------------------------------------
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'VendReceived' AND COLUMN_NAME = 'EntryDT')
BEGIN
    ALTER TABLE [dbo].[VendReceived] ADD [EntryDT] DATETIME NULL;
    PRINT '  + Added VendReceived.EntryDT';
END
GO

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'VendIssued' AND COLUMN_NAME = 'DTEntry')
BEGIN
    ALTER TABLE [dbo].[VendIssued] ADD [DTEntry] DATETIME NULL;
    PRINT '  + Added VendIssued.DTEntry';
END
GO

-- ----------------------------------------------------------------------------------------------------
-- 6. View: VVendRcvdDetail_Simple (must expose NextProcessID for VRunningLots_Simple_Main)
-- ----------------------------------------------------------------------------------------------------
PRINT '  * Synchronizing view [VVendRcvdDetail_Simple]...';
GO

CREATE OR ALTER VIEW [dbo].[VVendRcvdDetail_Simple]
AS
SELECT dbo.vendRcvdDetail.ItemCode, dbo.vendRcvdDetail.LotNo, dbo.vendRcvdDetail.OrderNo, dbo.Processes.Description, dbo.vendRcvdDetail.RcvdQty, dbo.vendRcvdDetail.IssQty, dbo.vendRcvdDetail.ProcessID, dbo.Processes.Code, dbo.ItemProcesses.SNO, dbo.vendRcvdDetail.Issue_RefID, dbo.vendRcvdDetail.EntryID, dbo.vendRcvdDetail.Wastage, dbo.vendRcvdDetail.ReWorkQty, 
         dbo.vendRcvdDetail.Opening_RefID, dbo.Items.GroupID, CAST(CONVERT(Varchar(10), dbo.VendReceived.DT, 1) AS DATETIME) AS DT, dbo.VendReceived.VendID, dbo.vendRcvdDetail.ReWorkLot, dbo.FCustomerOrders.InternalRefNo, dbo.FCustomerOrders.CustCode, dbo.VendReceived.OverTime, dbo.vendRcvdDetail.RefID, dbo.VendReceived.Issuance_RefID,
         dbo.vendRcvdDetail.NextProcessID
FROM  dbo.vendRcvdDetail INNER JOIN
         dbo.Processes ON dbo.vendRcvdDetail.ProcessID = dbo.Processes.ProcessID LEFT OUTER JOIN
         dbo.ItemProcesses ON dbo.vendRcvdDetail.ProcessID = dbo.ItemProcesses.ProcessID AND dbo.vendRcvdDetail.ItemCode = dbo.ItemProcesses.ItemID INNER JOIN
         dbo.Items ON dbo.vendRcvdDetail.ItemCode = dbo.Items.ItemID INNER JOIN
         dbo.VendReceived ON dbo.vendRcvdDetail.RefID = dbo.VendReceived.EntryID INNER JOIN
         dbo.FCustomerOrders ON dbo.vendRcvdDetail.OrderNo = dbo.FCustomerOrders.OrderNo
GO

-- ----------------------------------------------------------------------------------------------------
-- 7. View: VVendIssdDetailWithValueWORcving (must expose Batch_No for VVendIssued)
-- ----------------------------------------------------------------------------------------------------
PRINT '  * Synchronizing view [VVendIssdDetailWithValueWORcving]...';
GO

CREATE OR ALTER VIEW [dbo].[VVendIssdDetailWithValueWORcving]
AS
SELECT        dbo.VendIssued.EntryID AS MainEntryID, dbo.VendIssued.RecieptID AS MainRecieptID, dbo.VendIssued.DT, dbo.VendIssdDetail.EntryID, dbo.VendIssued.VendID, 
                         dbo.VendIssdDetail.RefID, dbo.VendIssdDetail.RecieptID, dbo.VendIssdDetail.ItemCode, dbo.VendIssdDetail.Rate, dbo.VendIssdDetail.IssQty, 
                         dbo.VendIssdDetail.RcvdQty, dbo.VendIssdDetail.LotNo, dbo.VendIssdDetail.ReqAuth, dbo.VendIssdDetail.OrderNo, dbo.VendIssdDetail.RcvProcessID, 
                         dbo.VendIssdDetail.ReturnDT, dbo.VendIssdDetail.Priority, dbo.VendIssdDetail.SpecialInstructions, dbo.VendIssued.VchrNo, dbo.Items.ForgingWeight, 
                         CASE WHEN VendAssItems.Unit = 'Kgs' THEN ((ForgingWeight * VendIssdDetail.IssQty) / 1000) 
                         * VendIssdDetail.Rate ELSE VendIssdDetail.IssQty * VendIssdDetail.Rate END AS IssValue, 
                         CASE WHEN VendAssItems.Unit = 'Kgs' THEN ((ForgingWeight * (VendIssdDetail.IssQty - VendIssdDetail.RcvdQty)) / 1000) 
                         * VendIssdDetail.Rate ELSE (VendIssdDetail.IssQty - VendIssdDetail.RcvdQty) * VendIssdDetail.Rate END AS BalanceValue, dbo.VendIssdDetail.ReWorkLot, 
                         dbo.VendIssdDetail.Repair_RefID, dbo.FCustomerOrders.CustCode, dbo.FCustomerOrders.InternalRefNo, dbo.VID_Bookmarks.EntryID AS BookMarkEntryID, 
                         T1.ItemCode AS CustItemCode, 
			T1.Quality, dbo.VIssuanceEntryWithWastage.WastageQty,T1.OrderQty,VendIssdDetail.Batch_No
FROM            dbo.VendIssdDetail INNER JOIN
                         dbo.VendIssued ON dbo.VendIssdDetail.RefID = dbo.VendIssued.EntryID LEFT OUTER JOIN
                         dbo.VendAssItems ON dbo.VendIssued.VendID = dbo.VendAssItems.VendID AND dbo.VendIssdDetail.ItemCode = dbo.VendAssItems.ItemID AND 
                         dbo.VendIssued.ProcessID = dbo.VendAssItems.ProcessID INNER JOIN
                         dbo.Items ON dbo.VendIssued.ItemID = dbo.Items.ItemID INNER JOIN
                         dbo.FCustomerOrders ON dbo.VendIssdDetail.OrderNo = dbo.FCustomerOrders.OrderNo LEFT OUTER JOIN
                         dbo.VID_Bookmarks ON dbo.VendIssdDetail.EntryID = dbo.VID_Bookmarks.VID_RefID LEFT OUTER JOIN
                             (SELECT        OrderNo, ItemCode,CompItemCode, MAX(Quality) AS Quality,SUM(Qty) AS OrderQty
                                FROM            dbo.FOrderItems
                                GROUP BY OrderNo, ItemCode,CompItemCode) AS T1 ON dbo.VendIssdDetail.OrderNo = T1.OrderNo AND 
                         dbo.VendIssdDetail.ItemCode = T1.CompItemCode LEFT OUTER JOIN
                         dbo.VIssuanceEntryWithWastage ON dbo.VendIssdDetail.EntryID = dbo.VIssuanceEntryWithWastage.EntryID
GO

PRINT '>>> Prerequisite Schema Updates Completed Successfully!';
GO
