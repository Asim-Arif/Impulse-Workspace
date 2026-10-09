-- ====================================================================================================
-- MASTER MIGRATION SCRIPT FOR IMPULSE ERP DATABASE
-- Generated: 2026-09-29 14:17:02
-- Total Scripts Combined: 54
-- Instructions:
--   1. Connect to your SQL Server (e.g. Asim-PC\SQL2017STD, WIN-KEJVO9CLD80, or production server).
--   2. Ensure the target database is selected (e.g. USE [Impulse_AWM] or USE [SMBI_AWM]).
--   3. Execute this entire script (F5 in SSMS). It is 100% idempotent and safe to re-run.
-- ====================================================================================================

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- ====================================================================================================
-- [01/49] SCRIPT: 00-Prerequisite_Schema_Updates.sql
-- ====================================================================================================
-- ====================================================================================================
-- Script Name: 00-Prerequisite_Schema_Updates.sql
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

GO

-- ====================================================================================================
-- [02/49] SCRIPT: 04-07-26-1-VFCustomerCatalog_Ex.sql
-- ====================================================================================================
-- [USE statement stripped for database independence]
GO

/****** Object:  View [dbo].[VFCustomerCatalog_Ex]    Script Date: 7/4/2026 1:27:07 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO






CREATE OR ALTER VIEW [dbo].[VFCustomerCatalog_Ex]
AS
SELECT     dbo.FCustomerCatalog.EntryID, dbo.FCustomerCatalog.CustCode, dbo.FCustomerCatalog.Country, dbo.FCustomerCatalog.ItemID, dbo.FCustomerCatalog.Unit, 
                      dbo.FCustomerCatalog.CompItemID, dbo.FCustomerCatalog.Description, dbo.FCustomerCatalog.FOB, dbo.FCustomerCatalog.CnFAir, dbo.FCustomerCatalog.CnFSea, 
                      dbo.FCustomerCatalog.CIFAir, dbo.FCustomerCatalog.CIFSea, dbo.FCustomerCatalog.PackingMode, dbo.VItems.ItemName AS CompDesc, dbo.VItems.CatID, 
                      dbo.ItemCatagories.Description AS CatDesc, dbo.VItems.FOB AS CustomFOB, dbo.VItems.CIFSea AS CustomCIFSea, dbo.VItems.CIFAir AS CustomCIFAir, 
                      dbo.VItems.CnFSea AS CustomCnFSea, dbo.VItems.CnFAir AS CustomCnFAir, dbo.FCustomerCatalog.ExWorks, dbo.FCustomerCatalog.CnIAir, 
                      dbo.FCustomerCatalog.CnISea, dbo.FCustomerCatalog.ExWorks AS CustomExWorks, dbo.FCustomerCatalog.CnIAir AS CustomCnIAir, 
                      dbo.FCustomerCatalog.CnISea AS CustomCnISea, dbo.FCustomerCatalog.FOBTop, dbo.FCustomerCatalog.ExWorksTop, dbo.FCustomerCatalog.CnFAirTop, 
                      dbo.FCustomerCatalog.CnFSeaTop, dbo.FCustomerCatalog.CIFAirTop, dbo.FCustomerCatalog.CIFSeaTop, dbo.FCustomerCatalog.CnIAirTop, 
                      dbo.FCustomerCatalog.CnISeaTop, dbo.FCustomerCatalog.ILO, dbo.FCustomerCatalog.OCR, dbo.FCustomerCatalog.OCR2, dbo.FCustomerCatalog.FinQuality, 
                      dbo.FCustomerCatalog.StampInstructions, dbo.FCustomerCatalog.SpecialInstructions, dbo.VItems.InActive
					  ,dbo.VItems.ItemSize,dbo.VItems.SizeUnit,dbo.VItems.TipSize
					  ,TRL.Running_Lots_No,TRL.Running_Lots_Qty,TForging.QtyInStock AS Forging_Stock
					  ,dbo.VItems.ItemGroup,dbo.VItems.GroupID
                      ,dbo.FCustomerCatalog.BarcodeNo,dbo.FCustomerCatalog.Temper_Rate,dbo.FCustomerCatalog.PackingInstructions,dbo.FCustomerCatalog.ItemColor
                      ,dbo.FCustomerCatalog.BarcodeFile,dbo.FCustomerCatalog.First_Inspection_Rate,dbo.FCustomerCatalog.MDMA,dbo.FCustomerCatalog.SFDA_Listing_No
                      ,dbo.FCustomerCatalog.MD_Group
FROM         dbo.FCustomerCatalog INNER JOIN
                      dbo.VItems ON dbo.FCustomerCatalog.CompItemID = dbo.VItems.ItemID INNER JOIN
                      dbo.ItemCatagories ON dbo.VItems.CatID = dbo.ItemCatagories.CatID
					  LEFT JOIN (SELECT ItemCode,COUNT(LotNo) AS Running_Lots_No,SUM(Qty) AS Running_Lots_Qty FROM VRunningLots_Simple GROUP BY ItemCode) TRL ON dbo.VItems.ItemID=TRL.ItemCode
					  LEFT JOIN (SELECT ItemID,SUM(VRM.QtyInStock) AS QtyInStock FROM VRM
									INNER JOIN ItemsRMComp ON VRM.RMID=ItemsRMComp.RMID
									INNER JOIN RMGroupIDsForForging ON VRM.GroupID=RMGroupIDsForForging.Group_ID GROUP BY ItemID) TForging ON dbo.VItems.ItemID=TForging.ItemID
GO

GO

-- ====================================================================================================
-- [03/49] SCRIPT: 04-07-26-2-Cities.sql
-- ====================================================================================================
-- Replaced by 04-07-26-3-Cities.sql which adds IDENTITY(1,1)
PRINT '  * 04-07-26-2-Cities.sql supersceded by 04-07-26-3-Cities.sql';
GO

GO

-- ====================================================================================================
-- [04/49] SCRIPT: 04-07-26-3-Cities.sql
-- ====================================================================================================
IF COLUMNPROPERTY(OBJECT_ID('dbo.Cities'), 'CityID', 'IsIdentity') = 1
BEGIN
    PRINT '  * Cities.CityID is already an identity column. Skipping table rebuild.';
END
ELSE
BEGIN
    EXEC sp_executesql N'/*
   Saturday, July 4, 20262:45:29 PM
   User: sa
   Server: Asim-PC\SQL2014STD
   Database: SMBI_AWM
   Application: 
*/

/* To prevent any potential data loss issues, you should review this script in detail before running it outside the context of the database designer.*/
BEGIN TRANSACTION
SET QUOTED_IDENTIFIER ON
SET ARITHABORT ON
SET NUMERIC_ROUNDABORT OFF
SET CONCAT_NULL_YIELDS_NULL ON
SET ANSI_NULLS ON
SET ANSI_PADDING ON
SET ANSI_WARNINGS ON
COMMIT
BEGIN TRANSACTION
GO
ALTER TABLE dbo.Cities
	DROP CONSTRAINT FK_Cities_Countries
GO
ALTER TABLE dbo.Countries SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
BEGIN TRANSACTION
GO
CREATE TABLE dbo.Tmp_Cities
	(
	CityID int NOT NULL IDENTITY (1, 1),
	CountryName nvarchar(50) NOT NULL,
	City nvarchar(50) NULL
	)  ON [PRIMARY]
GO
ALTER TABLE dbo.Tmp_Cities SET (LOCK_ESCALATION = TABLE)
GO
SET IDENTITY_INSERT dbo.Tmp_Cities ON
GO
IF EXISTS(SELECT * FROM dbo.Cities)
	 EXEC(''INSERT INTO dbo.Tmp_Cities (CityID, CountryName, City)
		SELECT CityID, CountryName, City FROM dbo.Cities WITH (HOLDLOCK TABLOCKX)'')
GO
SET IDENTITY_INSERT dbo.Tmp_Cities OFF
GO
ALTER TABLE dbo.Ports
	DROP CONSTRAINT FK_Ports_Cities
GO
DROP TABLE dbo.Cities
GO
EXECUTE sp_rename N''dbo.Tmp_Cities'', N''Cities'', ''OBJECT'' 
GO
ALTER TABLE dbo.Cities ADD CONSTRAINT
	PK_Cities PRIMARY KEY CLUSTERED 
	(
	CityID
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
ALTER TABLE dbo.Cities ADD CONSTRAINT
	FK_Cities_Countries FOREIGN KEY
	(
	CountryName
	) REFERENCES dbo.Countries
	(
	CountryName
	) ON UPDATE  CASCADE 
	 ON DELETE  CASCADE 
	
GO
COMMIT
BEGIN TRANSACTION
GO
ALTER TABLE dbo.Ports ADD CONSTRAINT
	FK_Ports_Cities FOREIGN KEY
	(
	CityID
	) REFERENCES dbo.Cities
	(
	CityID
	) ON UPDATE  CASCADE 
	 ON DELETE  CASCADE 
	
GO
ALTER TABLE dbo.Ports SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
';
END
GO

GO

-- ====================================================================================================
-- [05/49] SCRIPT: 19-07-26-1-InsertVoucherHead_SP.sql
-- ====================================================================================================
-- [USE statement stripped for database independence]
GO
/****** Object:  StoredProcedure [dbo].[InsertVoucherHead_SP]    Script Date: 2/5/2026 7:57:15 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[InsertVoucherHead_SP] 
				(

					@VchrNo [varchar](50) ,
					@UserName [varchar](50) NULL,
					@MachineName [varchar](50) NULL,
					@DT [datetime] NULL,
					@Notes [varchar](8000) NULL,
					@Payee [varchar](255) NULL,
					@ChequeNo [varchar] (50) NULL,
					@BankAccNo [varchar] (50) NULL,
					@ChqBookNo int NULL,
					@DueDate [datetime] NULL,
					@Handed_Over_To [varchar] (255) NULL
				)
AS
BEGIN
	 
	SET NOCOUNT ON;

	DECLARE @VchrType VARCHAR(5)
	SET @VchrType=SUBSTRING(@VchrNo,1,CHARINDEX('-',@VchrNo,1)-1)
	---
	DECLARE @NewVchrSequence INT; -- Holds the sequential part (e.g., 1, 2, 3)
    DECLARE @VchrNoPrefix VARCHAR(255); -- Holds the fixed prefix (e.g., 'PAY-251126-')
    DECLARE @FinalVchrNo VARCHAR(255); -- Holds the final, complete VchrNo string

    -- 1. CALCULATE THE PREFIX
    SET @VchrNoPrefix = @VchrType + '-' + FORMAT(@DT, 'yyMMdd') + '-';

    -- START THE DATABASE TRANSACTION AND APPLY LOCKING
    BEGIN TRANSACTION;
    
    SELECT 
        @NewVchrSequence = ISNULL(MAX(CAST(RIGHT(VchrNo,CHARINDEX('-',REVERSE(VchrNo))-1) AS INT)), 0) + 1
    FROM 
        dbo.Vouchers WITH (UPDLOCK, HOLDLOCK) -- LOCKING IS APPLIED HERE
    WHERE 
        VDate = @DT 
        AND LEFT(VchrNo, 8 + LEN(@VchrType)) = @VchrNoPrefix;

    -- 3. ASSEMBLE THE FINAL UNIQUE VOUCHER NUMBER
    SET @FinalVchrNo = @VchrNoPrefix + FORMAT(@NewVchrSequence, '000');


	IF @ChequeNo IS NOT NULL AND @BankAccNo IS NOT NULL
	BEGIN
		DECLARE @IsAvailable BIT
		SELECT @IsAvailable=Issued FROM dbo.VChqList WITH (UPDLOCK, HOLDLOCK)		
			WHERE AccNo=@BankAccNo
			AND ChqNo=@ChequeNo
		IF @IsAvailable IS NULL OR @IsAvailable=1
		BEGIN
			ROLLBACK TRANSACTION;
			THROW 50001, N'The Cheque Number is already used.',1;
		END
	END

	INSERT INTO [dbo].[VoucherInfo]
           (
			   [VchrNo]
			   ,[UserName]
			   ,[MachineName]
			   ,[DT]           
			   ,[Notes]           
			   ,[Payee]			   
			   ,[DueDate]
			   ,[Handed_Over_To]
		   )
		OUTPUT INSERTED.VchrNo  --Return the Final VchrNo
     VALUES
		(
			@FinalVchrNo
			,@UserName
			,@MachineName
			,@DT           
			,@Notes           
			,@Payee			
			,@DueDate
			,@Handed_Over_To
		)

	IF @ChequeNo IS NOT NULL AND @BankAccNo IS NOT NULL
		BEGIN
			UPDATE ChqList SET ChqList.Issued=1
				WHERE ChqList.ChqBookNo=@ChqBookNo AND ChqList.ChqNo=@ChequeNo
		END

	COMMIT TRANSACTION;

END

GO

-- ====================================================================================================
-- [06/49] SCRIPT: 19-07-26-2-EnumValues.sql
-- ====================================================================================================
-- [USE statement stripped for database independence]
GO

/****** Object:  Table [dbo].[EnumValues]    Script Date: 11/23/2025 1:48:29 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'EnumValues')
CREATE TABLE [dbo].[EnumValues](
	[EntryID] [int] IDENTITY(1,1) NOT NULL,
	[EnumName] [varchar](50) NOT NULL,
	[EnumValue] [int] NOT NULL,
	[EnumDescription] [varchar](50) NOT NULL,
 CONSTRAINT [PK_EnumValues] PRIMARY KEY CLUSTERED 
(
	[EnumName] ASC,
	[EnumValue] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

GO

-- ====================================================================================================
-- [07/49] SCRIPT: 19-07-26-3-INSERT INTO EnumValues.sql
-- ====================================================================================================
IF NOT EXISTS (SELECT 1 FROM EnumValues WHERE EnumName = 'BPV_Cheque_Type')
INSERT INTO EnumValues(EnumName,EnumValue,EnumDescription)
VALUES('BPV_Cheque_Type',0,'Open')
,('BPV_Cheque_Type',1,'Payees A/C Only')

GO

-- ====================================================================================================
-- [08/49] SCRIPT: 19-07-26-4-CREATE TYPE dbo.Vouchers_TT.sql
-- ====================================================================================================
IF TYPE_ID(N'dbo.Vouchers_TT') IS NULL
BEGIN
    EXEC('CREATE TYPE dbo.Vouchers_TT AS TABLE
    (
        [SNo] [float] NULL,
        [VDate] [datetime] NULL,
        [VchrNo] [varchar](50) NOT NULL,
        [Accno] [nvarchar](255) NOT NULL,
        [Description] [varchar](8000) NULL,
        [Debit] [float] NULL,
        [Credit] [float] NULL,
        [balance] [float] NULL,
        [DpstSlip] [varchar](50) NULL,
        [CSNo] [bigint] NULL
    )');
END
GO

GO

-- ====================================================================================================
-- [09/49] SCRIPT: 19-07-26-5-GetNextVchrNo.sql
-- ====================================================================================================
-- [USE statement stripped for database independence]
GO
/****** Object:  UserDefinedFunction [dbo].[GetNextAccno]    Script Date: 11/26/2025 7:27:00 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE OR ALTER FUNCTION [dbo].[GetNextVchrNo](@DT DATETIME,@VchrType VARCHAR(50))

RETURNS VARCHAR(255) AS  
	BEGIN
		DECLARE @NewVchrNo INT
		SELECT @NewVchrNo=MAX(CAST(RIGHT(VchrNo,CHARINDEX('-',REVERSE(VchrNo))-1) AS INT)) FROM Vouchers 
			WHERE VDate=@DT
			AND LEFT(VchrNo,8+LEN(@VchrType))=@VchrType+'-'+FORMAT(@DT,'yyMMdd')+'-'

		SET @NewVchrNo=ISNULL(@NewVchrNo,0)+1

	
	RETURN @VchrType+'-'+FORMAT(@DT,'yyMMdd')+'-'+FORMAT(@NewVchrNo,'000')
END

GO

-- ====================================================================================================
-- [10/49] SCRIPT: 19-07-26-6-InsertVoucherLineItems_SP.sql
-- ====================================================================================================
-- [USE statement stripped for database independence]
GO
/****** Object:  StoredProcedure [dbo].[InsertVoucherLineItems_SP]    Script Date: 11/28/2025 9:34:11 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE OR ALTER PROCEDURE [dbo].[InsertVoucherLineItems_SP]
    @Vouchers_TT dbo.Vouchers_TT READONLY
AS
BEGIN
    
    SET NOCOUNT ON;

    INSERT INTO [dbo].[Vouchers] 
    (
        SNo,
		VDate,
		VchrNo,
        AccNo,
        Description,
		Debit,
		Credit,
		Balance,
		DpstSlip,
		CSNo
    )
    SELECT
		VL.SNo,
		VL.VDate,
        VL.VchrNo,
        VL.AccNo,
        VL.Description,
		VL.Debit,
        VL.Credit,
		--VL.Balance,		
		0,
		VL.DpstSlip,		
        VL.CSNo
        
    FROM @Vouchers_TT AS VL; 
	-- 1. Calculate and store the running balance ONE TIME
	SELECT
		V.SNo,
		V.AccNo,
		NewBalance = Accounts.OpenBal+SUM(V.Debit - V.Credit) OVER (PARTITION BY V.AccNo ORDER BY V.SNo)
	INTO #RecalcResults -- Stores the results in a temporary table
	FROM Vouchers V
	INNER JOIN Accounts ON V.AccNo=Accounts.AccNo
	WHERE V.AccNo IN (SELECT DISTINCT AccNo FROM @Vouchers_TT);
	--Update Running Balance in Vouchers
    UPDATE V
		SET V.Balance = R.NewBalance
		FROM Vouchers V
		JOIN #RecalcResults R ON V.SNo = R.SNo AND V.AccNo = R.AccNo;
	--Update Final Balance in Accounts Table
	UPDATE A
		SET A.Balance = X.FinalBalance
		FROM Accounts A
		INNER JOIN (
			SELECT r.AccNo, r.NewBalance AS FinalBalance
			FROM #RecalcResults r
			INNER JOIN (
				SELECT AccNo, MAX(SNo) AS LastSNo
				FROM #RecalcResults
				GROUP BY AccNo
			) L ON r.AccNo = L.AccNo AND r.SNo = L.LastSNo
		) X ON A.AccNo = X.AccNo;

	DROP TABLE #RecalcResults

END

GO

-- ====================================================================================================
-- [11/49] SCRIPT: 19-07-26-7-InsertChequeDetails_SP.sql
-- ====================================================================================================
-- [USE statement stripped for database independence]
GO
/****** Object:  StoredProcedure [dbo].[InsertChequeDetails_SP]    Script Date: 12/5/2025 9:54:06 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE OR ALTER PROCEDURE [dbo].[InsertChequeDetails_SP]
    @SNo [bigint] ,
	@CDate [smalldatetime] NULL,
	@AccNo [nvarchar](255) NULL,
	@Description [nvarchar](4000) NULL,
	@Amount [money] NULL,
	@BankID [int] NULL,
	@chequeno [nvarchar](50) NULL,
	@ChqBookNo [int] NULL,
	@chequeType [nvarchar](50) NULL,
	@chequeDate [smalldatetime] NULL,
	@Posted [bit] ,
	@Payment [bit] ,
	@Bounced [bit] ,
	@ChqIsDue [bit] NULL,
	@Payee [varchar](4000) NULL,
	@BankAccNo [varchar](50) NULL,
	@ClearanceDT [datetime] NULL,
	@ChqPrintingDone [bit] NULL
AS
BEGIN
    SET NOCOUNT ON;

    INSERT INTO [dbo].[Cheque]
           ([SNo]
           ,[CDate]
           ,[AccNo]
           ,[Description]
           ,[Amount]
           ,[BankID]
           ,[chequeno]
           ,[ChqBookNo]
           ,[chequeType]
           ,[chequeDate]
           ,[Posted]
           ,[Payment]
           ,[Bounced]
           ,[ChqIsDue]
           ,[Payee]
           ,[BankAccNo]
           ,[ClearanceDT]
           ,[ChqPrintingDone]
           )                
    OUTPUT INSERTED.SNo  -- ***CRITICAL: Returns the new ID***
    VALUES 
    (
        @SNo,
		@CDate,
		@AccNo,
		@Description,
		@Amount,
		@BankID,
		@chequeno,
		@ChqBookNo,
		@chequeType,
		@chequeDate,
		@Posted,
		@Payment,
		@Bounced,
		@ChqIsDue,
		@Payee,
		@BankAccNo,
		@ClearanceDT,
		@ChqPrintingDone
    );

	--Commented the following line as it's being handled in InsertVoucherHead_SP SP.
	--UPDATE ChqList SET Issued=1 WHERE ChqBookNo=@ChqBookNo AND ChqNo=@chequeno

END

GO

-- ====================================================================================================
-- [12/49] SCRIPT: 19-07-26-8-CREATE NONCLUSTERED INDEX IX_Vouchers_AccNo_SNo.sql
-- ====================================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Vouchers_AccNo_SNo' AND object_id = OBJECT_ID('dbo.Vouchers'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_Vouchers_AccNo_SNo 
    ON dbo.Vouchers (AccNo, SNo) 
    INCLUDE (Debit, Credit) 
    WITH (FILLFACTOR = 90);
END
GO

GO

-- ====================================================================================================
-- [13/49] SCRIPT: 19-07-26-9-INSERT INTO EnumValues.sql
-- ====================================================================================================
IF NOT EXISTS (SELECT 1 FROM EnumValues WHERE EnumName = 'BPV_Payment_Types')
INSERT INTO EnumValues(EnumName,EnumValue,EnumDescription)
VALUES('BPV_Payment_Types',0,'Cheque')

GO

-- ====================================================================================================
-- [14/49] SCRIPT: 19-07-26-10-VLedger_Original.sql
-- ====================================================================================================
-- [USE statement stripped for database independence]
GO

/****** Object:  View [dbo].[VLedger]    Script Date: 7/20/2026 11:26:14 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE OR ALTER VIEW [dbo].[VLedger_Original]
AS
SELECT     dbo.Vouchers.SNo, dbo.Vouchers.VDate, dbo.Vouchers.VchrNo, dbo.Vouchers.Accno, dbo.Accounts.AccTitle, dbo.Vouchers.Description, dbo.Vouchers.Debit, dbo.Vouchers.Credit, dbo.Vouchers.balance, dbo.Vouchers.CSNo, dbo.Vouchers.DpstSlip, 
                  dbo.VoucherInfo.UserName, dbo.VoucherInfo.MachineName, dbo.VoucherInfo.DT, dbo.Cheque.chequeno, dbo.Cheque.chequeDate, dbo.Accounts.SubAccOf, dbo.Cheque.chequeType, dbo.VoucherInfo.Handed_Over_To, dbo.Cheque.Payee
FROM        dbo.Accounts INNER JOIN
                  dbo.Vouchers ON dbo.Accounts.AccNo = dbo.Vouchers.Accno LEFT OUTER JOIN
                  dbo.Cheque ON dbo.Vouchers.CSNo = dbo.Cheque.SNo LEFT OUTER JOIN
                  dbo.VoucherInfo ON dbo.Vouchers.VchrNo = dbo.VoucherInfo.VchrNo
WHERE     (dbo.Vouchers.SNo NOT IN
                      (SELECT     Sno
                       FROM        dbo.ADVSno
                       WHERE     (AccNo = dbo.Vouchers.Accno)))
GO

GO

-- ====================================================================================================
-- [15/49] SCRIPT: 19-07-26-11-VLedger.sql
-- ====================================================================================================
CREATE OR ALTER VIEW [dbo].[VLedger] AS
WITH ChequeMapping AS (
    -- This isolates the VchrNo and finds the single populated CSNo, ignoring the NULLs
    SELECT 
        VchrNo, 
        MAX(CSNo) AS ValidCSNo
    FROM dbo.Vouchers
    WHERE CSNo IS NOT NULL
    GROUP BY VchrNo
)
SELECT 
    v.SNo, 
    v.VDate, 
    v.VchrNo, 
    v.Accno, 
    a.AccTitle, 
    v.Description, 
    v.Debit, 
    v.Credit, 
    v.balance, 
    v.CSNo, 
    v.DpstSlip, 
    vi.UserName, 
    vi.MachineName, 
    vi.DT, 
    c.chequeno, 
    c.chequeDate, 
    a.SubAccOf, 
    c.chequeType, 
    vi.Handed_Over_To, 
    c.Payee
FROM dbo.Vouchers v
INNER JOIN dbo.Accounts a 
    ON v.Accno = a.AccNo
-- Join to our new map to get the correct CSNo for every row of the voucher
LEFT OUTER JOIN ChequeMapping cm 
    ON v.VchrNo = cm.VchrNo
-- Join to the Cheque table using the mapped CSNo, not the row-level CSNo
LEFT OUTER JOIN dbo.Cheque c 
    ON cm.ValidCSNo = c.SNo
LEFT OUTER JOIN dbo.VoucherInfo vi 
    ON v.VchrNo = vi.VchrNo
WHERE v.SNo NOT IN (
    SELECT Sno 
    FROM dbo.ADVSno 
    WHERE AccNo = v.Accno
);

GO

-- ====================================================================================================
-- [16/49] SCRIPT: 19-07-26-12-VChqLedger.sql
-- ====================================================================================================
-- [USE statement stripped for database independence]
GO

/****** Object:  View [dbo].[VChqLedger]    Script Date: 7/20/2026 12:22:27 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



CREATE OR ALTER VIEW [dbo].[VChqLedger]
AS

SELECT        dbo.Cheque.SNo, dbo.Cheque.CDate, dbo.Cheque.AccNo, dbo.Cheque.Description, dbo.Cheque.Amount, dbo.Cheque.BankID, 
dbo.Cheque.chequeno, dbo.Cheque.ChqBookNo, dbo.Cheque.chequeType, dbo.Cheque.chequeDate, 
dbo.Cheque.Posted, dbo.Cheque.Payment, dbo.Cheque.Bounced, dbo.Cheque.ChqIsDue, dbo.Cheque.Payee AS ChqPayee, dbo.Cheque.BankAccNo, dbo.Cheque.ClearanceDT, dbo.Cheque.ChqPrintingDone, dbo.BankList.Bank, 
dbo.BankList.Branch, dbo.BankAccounts.AccNo AS BankAccNo_old, Accounts_1.AccTitle AS BankAccountTitle, dbo.BankAccounts.Chqformat, Accounts_2.AccTitle, dbo.Makers.NICNo
,dbo.VoucherInfo.Handed_Over_To,dbo.VoucherInfo.VchrNo,Makers.CompanyName AS MakerPayeeName,
ISNULL(dbo.Cheque.Payee,Makers.CompanyName) AS Payee
,dbo.Vouchers.VDate,dbo.Vouchers.SNo AS Vouchers_SNo

FROM dbo.Cheque
LEFT JOIN dbo.ChqBooks ON dbo.Cheque.ChqBookNo = dbo.ChqBooks.ChqBookNo
LEFT JOIN dbo.Accounts AS Accounts_2 ON dbo.Cheque.AccNo=Accounts_2.AccNo
LEFT JOIN dbo.Makers ON dbo.Cheque.AccNo=dbo.Makers.AccNo
LEFT JOIN dbo.BankAccounts ON dbo.Cheque.BankAccNo = dbo.BankAccounts.AccNo
LEFT JOIN dbo.BankList ON dbo.BankList.BankID = dbo.BankAccounts.BankID
LEFT JOIN dbo.Accounts AS Accounts_1 ON dbo.BankAccounts.AccNo = Accounts_1.AccNo
LEFT JOIN dbo.Vouchers ON dbo.Cheque.BankAccNo=Vouchers.AccNo AND dbo.Cheque.SNo=dbo.Vouchers.CSNo
LEFT JOIN dbo.VoucherInfo ON Vouchers.VchrNO=VoucherInfo.VchrNo
                         
GO

GO

-- ====================================================================================================
-- [17/49] SCRIPT: 19-07-26-13-BalanceTags.sql
-- ====================================================================================================
/*
   Monday, July 20, 202612:58:25 PM
   User: sa
   Server: Asim-PC\SQL2014STD
   Database: SMBI_AWM
   Application: 
*/

/* To prevent any potential data loss issues, you should review this script in detail before running it outside the context of the database designer.*/
BEGIN TRANSACTION
SET QUOTED_IDENTIFIER ON
SET ARITHABORT ON
SET NUMERIC_ROUNDABORT OFF
SET CONCAT_NULL_YIELDS_NULL ON
SET ANSI_NULLS ON
SET ANSI_PADDING ON
SET ANSI_WARNINGS ON
COMMIT
BEGIN TRANSACTION
GO
ALTER TABLE dbo.BalanceTags
	DROP CONSTRAINT DF_BalanceTags_DTEntry
GO
CREATE TABLE dbo.Tmp_BalanceTags
	(
	Sno float(53) NOT NULL,
	UserName varchar(50) NULL,
	DTEntry datetime NULL
	)  ON [PRIMARY]
GO
ALTER TABLE dbo.Tmp_BalanceTags SET (LOCK_ESCALATION = TABLE)
GO
ALTER TABLE dbo.Tmp_BalanceTags ADD CONSTRAINT
	DF_BalanceTags_DTEntry DEFAULT (getdate()) FOR DTEntry
GO
IF EXISTS(SELECT * FROM dbo.BalanceTags)
	 EXEC('INSERT INTO dbo.Tmp_BalanceTags (Sno, UserName, DTEntry)
		SELECT CONVERT(float(53), Sno), UserName, DTEntry FROM dbo.BalanceTags WITH (HOLDLOCK TABLOCKX)')
GO
DROP TABLE dbo.BalanceTags
GO
EXECUTE sp_rename N'dbo.Tmp_BalanceTags', N'BalanceTags', 'OBJECT' 
GO
ALTER TABLE dbo.BalanceTags ADD CONSTRAINT
	PK_BalanceTags PRIMARY KEY CLUSTERED 
	(
	Sno
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
COMMIT

GO

-- ====================================================================================================
-- [18/49] SCRIPT: 24-07-26-1-Trial_Balance_SP.sql
-- ====================================================================================================
-- [USE statement stripped for database independence]
GO
/****** Object:  StoredProcedure [dbo].[Trial_Balance_SP]    Script Date: 2/21/2026 4:32:34 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		Asim Arif
-- Create date: Feb-2017
-- Description:	Trial Balance (Fastest, Simplest)
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Trial_Balance_SP] (@DTFrom DATETIME,@DTTo DATETIME,@Code VARCHAR(50)='0')
	
AS
BEGIN
	
	SET NOCOUNT ON;
	
	DECLARE @CodeLength AS INT
	IF ISNULL(@Code,'0')='0'
		SET @CodeLength=0
	ELSE
		SET @CodeLength=LEN(@Code)

	CREATE TABLE #Accs(AccNo VARCHAR(50),Opening_Balance BIGINT,Opening_Debit BIGINT,Opening_Credit BIGINT,Debit BIGINT,Credit BIGINT,Closing_Balance BIGINT,Balance_Debit BIGINT,Balance_Credit BIGINT,EntryType TINYINT,AccType VARCHAR(50))
	CREATE TABLE #Accs_Parent(AccNo VARCHAR(50),Opening_Balance BIGINT,Opening_Debit BIGINT,Opening_Credit BIGINT,Debit BIGINT,Credit BIGINT,Closing_Balance BIGINT,Balance_Debit BIGINT,Balance_Credit BIGINT,EntryType TINYINT,AccType VARCHAR(50))
	CREATE TABLE #Accs_Heads(AccNo VARCHAR(50),Opening_Balance BIGINT,Opening_Debit BIGINT,Opening_Credit BIGINT,Debit BIGINT,Credit BIGINT,Closing_Balance BIGINT,Balance_Debit BIGINT,Balance_Credit BIGINT,EntryType TINYINT,AccType VARCHAR(50))

	INSERT INTO #Accs(AccNo,Opening_Balance,Closing_Balance,Debit,Credit,EntryType,AccType)
		SELECT Accounts.AccNo,ISNULL(TOpening.Balance,Accounts.openbal),ISNULL(TClosing.Balance,Accounts.openbal),TTotal.TotalDebit,TTotal.TotalCredit,0,Accounts.Type FROM Accounts
			LEFT JOIN (SELECT AccNo,Balance FROM Vouchers INNER JOIN (SELECT MAX(SNo) AS MaxSNo FROM Vouchers WHERE VDate<@DTFrom GROUP BY AccNo) T1 ON Vouchers.SNo=T1.MaxSNo) TOpening ON Accounts.AccNo=TOpening.Accno
			LEFT JOIN (SELECT AccNo,Balance FROM Vouchers INNER JOIN (SELECT MAX(SNo) AS MaxSNo FROM Vouchers WHERE VDate<=@DTTo GROUP BY AccNo) T1 ON Vouchers.SNo=T1.MaxSNo) TClosing ON Accounts.AccNo=TClosing.Accno
			LEFT JOIN (SELECT AccNo,SUM(Debit) AS TotalDebit,SUM(Credit) AS TotalCredit FROM Vouchers WHERE VDate BETWEEN @DTFrom AND @DTTo GROUP BY Accno) TTotal ON Accounts.AccNo=TTotal.AccNo
			WHERE Parent=0
			AND (LEFT(Accounts.AccNo,@CodeLength)=@Code OR ISNULL(@Code,'0')='0')
			--AND (Accounts.Account_Type=@Account_Type OR @Account_Type=-1)

	--INSERT INTO #Accs(AccNo,Opening_Debit,Opening_Credit,Debit,Credit,Balance_Debit,Balance_Credit)
	UPDATE #Accs SET	Opening_Debit=CASE WHEN Opening_Balance>=0 THEN Opening_Balance ELSE 0 END,
						Opening_Credit=CASE WHEN Opening_Balance<0 THEN ABS(Opening_Balance) ELSE 0 END,
						Balance_Debit=CASE WHEN Closing_Balance>=0 THEN Closing_Balance ELSE 0 END,
						Balance_Credit=CASE WHEN Closing_Balance<0 THEN ABS(Closing_Balance) ELSE 0 END

	INSERT INTO #Accs_Parent(AccNo,EntryType,AccType)
		SELECT AccNo,1,Accounts.Type FROM Accounts
			WHERE Parent=1
			AND (LEFT(Accounts.AccNo,@CodeLength)=@Code OR ISNULL(@Code,'0')='0')

	UPDATE T1 SET	T1.Opening_Debit=(SELECT SUM(Opening_Debit) FROM #Accs T2 WHERE LEFT(T2.AccNo,LEN(T1.AccNo))=T1.AccNo),
					T1.Opening_Credit=(SELECT SUM(Opening_Credit) FROM #Accs T2 WHERE LEFT(T2.AccNo,LEN(T1.AccNo))=T1.AccNo),
					T1.Balance_Debit=(SELECT SUM(Balance_Debit) FROM #Accs T2 WHERE LEFT(T2.AccNo,LEN(T1.AccNo))=T1.AccNo),
					T1.Balance_Credit=(SELECT SUM(Balance_Credit) FROM #Accs T2 WHERE LEFT(T2.AccNo,LEN(T1.AccNo))=T1.AccNo),
					T1.Debit=(SELECT SUM(Debit) FROM #Accs T2 WHERE LEFT(T2.AccNo,LEN(T1.AccNo))=T1.AccNo),
					T1.Credit=(SELECT SUM(Credit) FROM #Accs T2 WHERE LEFT(T2.AccNo,LEN(T1.AccNo))=T1.AccNo)
		FROM #Accs_Parent T1

	INSERT INTO #Accs_Heads(AccNo,EntryType,AccType)
		SELECT Code,2,Code FROM Heads
		
	UPDATE T1 SET	T1.Opening_Debit=(SELECT SUM(Opening_Debit) FROM #Accs T2 WHERE LEFT(T2.AccNo,2)=T1.AccNo),
					T1.Opening_Credit=(SELECT SUM(Opening_Credit) FROM #Accs T2 WHERE LEFT(T2.AccNo,2)=T1.AccNo),
					T1.Balance_Debit=(SELECT SUM(Balance_Debit) FROM #Accs T2 WHERE LEFT(T2.AccNo,2)=T1.AccNo),
					T1.Balance_Credit=(SELECT SUM(Balance_Credit) FROM #Accs T2 WHERE LEFT(T2.AccNo,2)=T1.AccNo),
					T1.Debit=(SELECT SUM(Debit) FROM #Accs T2 WHERE LEFT(T2.AccNo,2)=T1.AccNo),
					T1.Credit=(SELECT SUM(Credit) FROM #Accs T2 WHERE LEFT(T2.AccNo,2)=T1.AccNo)
		FROM #Accs_Heads T1
		
		
	SELECT * FROM 
	(SELECT #Accs.*,Accounts.AccTitle FROM #Accs INNER JOIN Accounts ON #Accs.AccNo=Accounts.AccNo
	UNION ALL
	SELECT #Accs_Parent.*,Accounts.AccTitle FROM #Accs_Parent INNER JOIN Accounts ON #Accs_Parent.AccNo=Accounts.AccNo
	UNION ALL
	SELECT #Accs_Heads.*,Heads.Head FROM #Accs_Heads INNER JOIN Heads ON #Accs_Heads.AccNo=Heads.Code) TUnioned 
	--WHERE (AccType=@Code OR @Code='0')
	--WHERE (LEFT(TUnioned.AccNo,@CodeLength)=@Code OR @Code IS NULL)

	ORDER BY AccNo
	
	DROP TABLE #Accs
	DROP TABLE #Accs_Parent
	DROP TABLE #Accs_Heads
END

GO

-- ====================================================================================================
-- [19/49] SCRIPT: 24-07-26-2-Trial_Balance_SP_Ex.sql
-- ====================================================================================================
-- [USE statement stripped for database independence]
GO

/****** Object:  StoredProcedure [dbo].[Trial_Balance_SP] ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		Asim Arif
-- Refactored:  Performance Optimized (No Resultset Changes)
-- Create date: Feb-2017
-- Description:	Trial Balance (Fastest, Simplest)
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Trial_Balance_SP_Ex] (
    @DTFrom DATETIME,
    @DTTo DATETIME,
    @Code VARCHAR(50)='0'
)
AS
BEGIN
    SET NOCOUNT ON;
	
    DECLARE @CodeLength AS INT
    IF ISNULL(@Code,'0') = '0'
        SET @CodeLength = 0
    ELSE
        SET @CodeLength = LEN(@Code)

    -- Single temp table for base accounts (Parent = 0)
    CREATE TABLE #Accs (
        AccNo VARCHAR(50),
        Opening_Balance BIGINT,
        Opening_Debit BIGINT,
        Opening_Credit BIGINT,
        Debit BIGINT,
        Credit BIGINT,
        Closing_Balance BIGINT,
        Balance_Debit BIGINT,
        Balance_Credit BIGINT,
        EntryType TINYINT,
        AccType VARCHAR(50)
    )

    -- 1. Insert and calculate child accounts directly without UPDATE passes
    INSERT INTO #Accs (
        AccNo, Opening_Balance, Closing_Balance,
        Opening_Debit, Opening_Credit,
        Balance_Debit, Balance_Credit,
        Debit, Credit, EntryType, AccType
    )
    SELECT 
        A.AccNo, 
        Base.OpBal, 
        Base.ClBal,
        CASE WHEN Base.OpBal >= 0 THEN Base.OpBal ELSE 0 END,
        CASE WHEN Base.OpBal < 0 THEN ABS(Base.OpBal) ELSE 0 END,
        CASE WHEN Base.ClBal >= 0 THEN Base.ClBal ELSE 0 END,
        CASE WHEN Base.ClBal < 0 THEN ABS(Base.ClBal) ELSE 0 END,
        Tot.TotalDebit,
        Tot.TotalCredit,
        0,
        A.Type
    FROM Accounts A
    CROSS APPLY (
        -- Replaces expensive MAX(SNo) groupings with highly targeted Top 1 index seeks
        SELECT 
            OpBal = ISNULL((SELECT TOP 1 Balance FROM Vouchers V WHERE V.AccNo = A.AccNo AND V.VDate < @DTFrom ORDER BY SNo DESC), A.openbal),
            ClBal = ISNULL((SELECT TOP 1 Balance FROM Vouchers V WHERE V.AccNo = A.AccNo AND V.VDate <= @DTTo ORDER BY SNo DESC), A.openbal)
    ) Base
    OUTER APPLY (
        SELECT SUM(Debit) AS TotalDebit, SUM(Credit) AS TotalCredit 
        FROM Vouchers V 
        WHERE V.AccNo = A.AccNo AND V.VDate BETWEEN @DTFrom AND @DTTo
    ) Tot
    WHERE A.Parent = 0
      AND (LEFT(A.AccNo, @CodeLength) = @Code OR @CodeLength = 0)


    -- 2. Output combined resultset directly (Base + Parents + Heads)
    SELECT 
        A.AccNo, A.Opening_Balance, A.Opening_Debit, A.Opening_Credit, 
        A.Debit, A.Credit, A.Closing_Balance, A.Balance_Debit, A.Balance_Credit, 
        A.EntryType, A.AccType, Acc.AccTitle
    FROM #Accs A
    INNER JOIN Accounts Acc ON A.AccNo = Acc.AccNo

    UNION ALL

    -- Aggregate Parent Accounts (Parent = 1)
    SELECT 
        P.AccNo, 
        CAST(NULL AS BIGINT) AS Opening_Balance,
        SUM(C.Opening_Debit) AS Opening_Debit,
        SUM(C.Opening_Credit) AS Opening_Credit,
        SUM(C.Debit) AS Debit,
        SUM(C.Credit) AS Credit,
        CAST(NULL AS BIGINT) AS Closing_Balance,
        SUM(C.Balance_Debit) AS Balance_Debit,
        SUM(C.Balance_Credit) AS Balance_Credit,
        CAST(1 AS TINYINT) AS EntryType, 
        P.Type AS AccType,
        P.AccTitle
    FROM Accounts P
    LEFT JOIN #Accs C ON LEFT(C.AccNo, LEN(P.AccNo)) = P.AccNo
    WHERE P.Parent = 1
      AND (LEFT(P.AccNo, @CodeLength) = @Code OR @CodeLength = 0)
    GROUP BY P.AccNo, P.Type, P.AccTitle

    UNION ALL

    -- Aggregate Head Accounts
    SELECT 
        H.Code AS AccNo, 
        CAST(NULL AS BIGINT) AS Opening_Balance,
        SUM(C.Opening_Debit) AS Opening_Debit,
        SUM(C.Opening_Credit) AS Opening_Credit,
        SUM(C.Debit) AS Debit,
        SUM(C.Credit) AS Credit,
        CAST(NULL AS BIGINT) AS Closing_Balance,
        SUM(C.Balance_Debit) AS Balance_Debit,
        SUM(C.Balance_Credit) AS Balance_Credit,
        CAST(2 AS TINYINT) AS EntryType, 
        H.Code AS AccType,
        H.Head AS AccTitle
    FROM Heads H
    LEFT JOIN #Accs C ON LEFT(C.AccNo, 2) = H.Code
    GROUP BY H.Code, H.Head

    ORDER BY AccNo
	
    DROP TABLE #Accs
END

GO

-- ====================================================================================================
-- [20/49] SCRIPT: 24-07-26-3-CREATE NONCLUSTERED INDEX [IX_Vouchers_AccNo_SNo].sql
-- ====================================================================================================
CREATE NONCLUSTERED INDEX [IX_Vouchers_AccNo_SNo] 
ON [dbo].[Vouchers] (AccNo, VDate) 
INCLUDE (Debit, Credit, Balance) 
WITH (DROP_EXISTING = ON);

GO

-- ====================================================================================================
-- [21/49] SCRIPT: 24-07-26-4-ChequePrintHistory.sql
-- ====================================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'ChequePrintHistory')
BEGIN
    CREATE TABLE [dbo].[ChequePrintHistory](
        [Sno] [bigint] NULL,
        [Payee] [varchar](500) NULL,
        [PrintDT] [datetime] NULL,
        [Designation] [varchar](255) NULL,
        [Company] [varchar](255) NULL,
        [EntryDT] [datetime] NULL CONSTRAINT [DF_ChequePrintHistory_EntryDT] DEFAULT (getdate()),
        [UserName] [varchar](50) NULL,
        [MachineName] [varchar](50) NULL,
        [EntryID] [int] IDENTITY(1,1) NOT NULL PRIMARY KEY CLUSTERED ([EntryID])
    );
END
GO

GO

-- ====================================================================================================
-- [22/49] SCRIPT: 25-07-26-1-UPDATE VendOrderDetail.sql
-- ====================================================================================================
UPDATE VendOrderDetail SET CCItem=0 WHERE CCItem IS NULL

GO

-- ====================================================================================================
-- [23/49] SCRIPT: 31-07-26-1-INSERT INTO GeneralData.sql
-- ====================================================================================================
INSERT INTO GeneralData(DataName,DataValue)
	VALUES('Company','IAA')

GO

-- ====================================================================================================
-- [24/49] SCRIPT: 08-08-26-1-WastageTypes.sql
-- ====================================================================================================
-- Replaced by 08-08-26-2-WastageTypes.sql which adds IDENTITY(1,1)
PRINT '  * 08-08-26-1-WastageTypes.sql supersceded by 08-08-26-2-WastageTypes.sql';
GO

GO

-- ====================================================================================================
-- [25/49] SCRIPT: 08-08-26-2-WastageTypes.sql
-- ====================================================================================================
IF COLUMNPROPERTY(OBJECT_ID('dbo.WastageTypes'), 'EntryID', 'IsIdentity') = 1
BEGIN
    PRINT '  * WastageTypes.EntryID is already an identity column. Skipping table rebuild.';
END
ELSE
BEGIN
    EXEC sp_executesql N'/*
   Saturday, August 8, 20267:05:49 PM
   User: sa
   Server: Asim-PC\SQL2014STD
   Database: SMBI_AWM
   Application: 
*/

/* To prevent any potential data loss issues, you should review this script in detail before running it outside the context of the database designer.*/
BEGIN TRANSACTION
SET QUOTED_IDENTIFIER ON
SET ARITHABORT ON
SET NUMERIC_ROUNDABORT OFF
SET CONCAT_NULL_YIELDS_NULL ON
SET ANSI_NULLS ON
SET ANSI_PADDING ON
SET ANSI_WARNINGS ON
COMMIT
BEGIN TRANSACTION
GO
ALTER TABLE dbo.WastageTypes
	DROP CONSTRAINT DF_WastageTypes_Closed
GO
CREATE TABLE dbo.Tmp_WastageTypes
	(
	EntryID int NOT NULL IDENTITY (1, 1),
	WastageName varchar(50) NULL,
	WastageType tinyint NULL,
	Closed bit NULL
	)  ON [PRIMARY]
GO
ALTER TABLE dbo.Tmp_WastageTypes SET (LOCK_ESCALATION = TABLE)
GO
ALTER TABLE dbo.Tmp_WastageTypes ADD CONSTRAINT
	DF_WastageTypes_Closed DEFAULT ((0)) FOR Closed
GO
SET IDENTITY_INSERT dbo.Tmp_WastageTypes ON
GO
IF EXISTS(SELECT * FROM dbo.WastageTypes)
	 EXEC(''INSERT INTO dbo.Tmp_WastageTypes (EntryID, WastageName, WastageType, Closed)
		SELECT EntryID, WastageName, WastageType, Closed FROM dbo.WastageTypes WITH (HOLDLOCK TABLOCKX)'')
GO
SET IDENTITY_INSERT dbo.Tmp_WastageTypes OFF
GO
ALTER TABLE dbo.VendRcvdDetailWastageDetail
	DROP CONSTRAINT FK_VendRcvdDetailWastageDetail_WastageTypes
GO
DROP TABLE dbo.WastageTypes
GO
EXECUTE sp_rename N''dbo.Tmp_WastageTypes'', N''WastageTypes'', ''OBJECT'' 
GO
ALTER TABLE dbo.WastageTypes ADD CONSTRAINT
	PK_WastageTypes PRIMARY KEY CLUSTERED 
	(
	EntryID
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
COMMIT
BEGIN TRANSACTION
GO
ALTER TABLE dbo.VendRcvdDetailWastageDetail WITH NOCHECK ADD CONSTRAINT
	FK_VendRcvdDetailWastageDetail_WastageTypes FOREIGN KEY
	(
	Wastage_RefID
	) REFERENCES dbo.WastageTypes
	(
	EntryID
	) ON UPDATE  CASCADE 
	 ON DELETE  CASCADE 
	
GO
ALTER TABLE dbo.VendRcvdDetailWastageDetail SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
';
END
GO

GO

-- ====================================================================================================
-- [26/49] SCRIPT: 24-08-26-1-Users.sql
-- ====================================================================================================
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='Users' AND COLUMN_NAME='EmpID')
BEGIN
    ALTER TABLE dbo.Users ADD EmpID varchar(50) NULL;
END
GO

GO

-- ====================================================================================================
-- [27/49] SCRIPT: 24-08-26-2-Intraoffice Tables.sql
-- ====================================================================================================
-- ==============================================================================================
-- IntraOffice Communication Suite - Clean Table Creation Script for SMBI_AWM
-- Target: Microsoft SQL Server 2012 / 2014 / 2016 / 2019 / 2022
-- Integrates directly with existing [Departments] (deptid varchar(50)) and [Users] (UserName / EmpID)
-- ==============================================================================================
-- [USE statement stripped for database independence]
GO

-- 1. CHANNELS (Departmental & Custom Group Chat)
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Channels')
BEGIN
    CREATE TABLE [dbo].[Channels] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [Name] NVARCHAR(200) NOT NULL,
        [Description] NVARCHAR(1000) NULL,
        [DepartmentId] VARCHAR(50) NULL,
        [IsPrivate] BIT NOT NULL DEFAULT 0,
        [IsActive] BIT NOT NULL DEFAULT 1,
        [CreatedBy] NVARCHAR(100) NOT NULL,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE()
    );
    CREATE NONCLUSTERED INDEX [IX_Channels_DepartmentId] ON [dbo].[Channels]([DepartmentId]);
    CREATE NONCLUSTERED INDEX [IX_Channels_CreatedBy] ON [dbo].[Channels]([CreatedBy]);
    PRINT 'Created table: Channels';
END
GO

-- 2. CHANNEL MEMBERS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ChannelMembers')
BEGIN
    CREATE TABLE [dbo].[ChannelMembers] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [ChannelId] INT NOT NULL,
        [UserId] NVARCHAR(100) NOT NULL,
        [IsAdmin] BIT NOT NULL DEFAULT 0,
        [JoinedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        CONSTRAINT [FK_ChannelMembers_Channels] FOREIGN KEY ([ChannelId]) REFERENCES [dbo].[Channels]([Id]) ON DELETE CASCADE
    );
    CREATE UNIQUE NONCLUSTERED INDEX [IX_ChannelMembers_ChannelId_UserId] ON [dbo].[ChannelMembers]([ChannelId], [UserId]);
    CREATE NONCLUSTERED INDEX [IX_ChannelMembers_UserId] ON [dbo].[ChannelMembers]([UserId]);
    PRINT 'Created table: ChannelMembers';
END
GO

-- 3. MESSAGES (Direct 1-on-1 & Channel Messages)
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Messages')
BEGIN
    CREATE TABLE [dbo].[Messages] (
        [Id] BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [ChannelId] INT NULL,
        [SenderId] NVARCHAR(100) NOT NULL,
        [ReceiverId] NVARCHAR(100) NULL,
        [Content] NVARCHAR(MAX) NOT NULL,
        [MessageType] INT NOT NULL DEFAULT 0, -- 0=Text, 1=File, 2=Audio/Voice, 3=System
        [IsRead] BIT NOT NULL DEFAULT 0,
        [IsDeleted] BIT NOT NULL DEFAULT 0,
        [ParentMessageId] BIGINT NULL,
        [SentAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [EditedAt] DATETIME2 NULL,
        CONSTRAINT [FK_Messages_Channels] FOREIGN KEY ([ChannelId]) REFERENCES [dbo].[Channels]([Id]) ON DELETE SET NULL,
        CONSTRAINT [FK_Messages_ParentMessage] FOREIGN KEY ([ParentMessageId]) REFERENCES [dbo].[Messages]([Id])
    );
    CREATE NONCLUSTERED INDEX [IX_Messages_ChannelId_SentAt] ON [dbo].[Messages]([ChannelId], [SentAt] DESC);
    CREATE NONCLUSTERED INDEX [IX_Messages_Sender_Receiver] ON [dbo].[Messages]([SenderId], [ReceiverId], [SentAt] DESC);
    CREATE NONCLUSTERED INDEX [IX_Messages_ParentMessageId] ON [dbo].[Messages]([ParentMessageId]);
    PRINT 'Created table: Messages';
END
GO

-- 4. MESSAGE ATTACHMENTS (Images, Documents, Voice Notes)
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'MessageAttachments')
BEGIN
    CREATE TABLE [dbo].[MessageAttachments] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [MessageId] BIGINT NOT NULL,
        [FileName] NVARCHAR(510) NOT NULL,
        [FilePath] NVARCHAR(1000) NOT NULL,
        [FileSize] BIGINT NOT NULL DEFAULT 0,
        [ContentType] NVARCHAR(200) NULL,
        [UploadedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        CONSTRAINT [FK_MessageAttachments_Messages] FOREIGN KEY ([MessageId]) REFERENCES [dbo].[Messages]([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_MessageAttachments_MessageId] ON [dbo].[MessageAttachments]([MessageId]);
    PRINT 'Created table: MessageAttachments';
END
GO

-- 5. ANNOUNCEMENTS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Announcements')
BEGIN
    CREATE TABLE [dbo].[Announcements] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [Title] NVARCHAR(400) NOT NULL,
        [Content] NVARCHAR(MAX) NOT NULL,
        [Priority] INT NOT NULL DEFAULT 0, -- 0=Normal, 1=Important, 2=Critical
        [CreatedBy] NVARCHAR(100) NOT NULL,
        [DepartmentId] VARCHAR(50) NULL,
        [IsPinned] BIT NOT NULL DEFAULT 0,
        [IsActive] BIT NOT NULL DEFAULT 1,
        [IsAcknowledged] BIT NOT NULL DEFAULT 0,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [UpdatedAt] DATETIME2 NULL,
        [ExpiresAt] DATETIME2 NULL
    );
    CREATE NONCLUSTERED INDEX [IX_Announcements_CreatedBy] ON [dbo].[Announcements]([CreatedBy]);
    CREATE NONCLUSTERED INDEX [IX_Announcements_DepartmentId] ON [dbo].[Announcements]([DepartmentId]);
    CREATE NONCLUSTERED INDEX [IX_Announcements_CreatedAt] ON [dbo].[Announcements]([CreatedAt] DESC);
    PRINT 'Created table: Announcements';
END
GO

-- 6. ANNOUNCEMENT ATTACHMENTS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'AnnouncementAttachments')
BEGIN
    CREATE TABLE [dbo].[AnnouncementAttachments] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [AnnouncementId] INT NOT NULL,
        [FileName] NVARCHAR(MAX) NOT NULL,
        [FilePath] NVARCHAR(MAX) NOT NULL,
        [FileSize] BIGINT NOT NULL,
        [ContentType] NVARCHAR(MAX) NULL,
        [UploadedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        CONSTRAINT [FK_AnnouncementAttachments_Announcements] FOREIGN KEY ([AnnouncementId]) REFERENCES [dbo].[Announcements]([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_AnnouncementAttachments_AnnouncementId] ON [dbo].[AnnouncementAttachments]([AnnouncementId]);
    PRINT 'Created table: AnnouncementAttachments';
END
GO

-- 7. ANNOUNCEMENT ACKNOWLEDGMENTS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'AnnouncementAcknowledgments')
BEGIN
    CREATE TABLE [dbo].[AnnouncementAcknowledgments] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [AnnouncementId] INT NOT NULL,
        [UserId] NVARCHAR(100) NOT NULL,
        [AcknowledgedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        CONSTRAINT [FK_AnnouncementAcknowledgments_Announcements] FOREIGN KEY ([AnnouncementId]) REFERENCES [dbo].[Announcements]([Id]) ON DELETE CASCADE
    );
    CREATE UNIQUE NONCLUSTERED INDEX [IX_AnnouncementAcknowledgments_Announcement_User] ON [dbo].[AnnouncementAcknowledgments]([AnnouncementId], [UserId]);
    PRINT 'Created table: AnnouncementAcknowledgments';
END
GO

-- 8. TASK ITEMS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'TaskItems')
BEGIN
    CREATE TABLE [dbo].[TaskItems] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [Title] NVARCHAR(400) NOT NULL,
        [Description] NVARCHAR(MAX) NULL,
        [AssignedTo] NVARCHAR(100) NULL,
        [AssignedBy] NVARCHAR(100) NOT NULL,
        [DepartmentId] VARCHAR(50) NULL,
        [Priority] INT NOT NULL DEFAULT 1, -- 0=Low, 1=Normal, 2=High, 3=Urgent
        [Status] INT NOT NULL DEFAULT 0,   -- 0=Pending, 1=InProgress, 2=Completed, 3=Cancelled
        [DueDate] DATETIME2 NULL,
        [WhatsAppMessageSent] BIT NOT NULL DEFAULT 0,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [UpdatedAt] DATETIME2 NULL,
        [CompletedAt] DATETIME2 NULL
    );
    CREATE NONCLUSTERED INDEX [IX_TaskItems_AssignedTo] ON [dbo].[TaskItems]([AssignedTo]);
    CREATE NONCLUSTERED INDEX [IX_TaskItems_AssignedBy] ON [dbo].[TaskItems]([AssignedBy]);
    CREATE NONCLUSTERED INDEX [IX_TaskItems_Status] ON [dbo].[TaskItems]([Status]);
    CREATE NONCLUSTERED INDEX [IX_TaskItems_DepartmentId] ON [dbo].[TaskItems]([DepartmentId]);
    PRINT 'Created table: TaskItems';
END
GO

-- 9. TASK COMMENTS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'TaskComments')
BEGIN
    CREATE TABLE [dbo].[TaskComments] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [TaskId] INT NOT NULL,
        [UserId] NVARCHAR(100) NOT NULL,
        [Content] NVARCHAR(MAX) NOT NULL,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        CONSTRAINT [FK_TaskComments_TaskItems] FOREIGN KEY ([TaskId]) REFERENCES [dbo].[TaskItems]([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_TaskComments_TaskId] ON [dbo].[TaskComments]([TaskId]);
    PRINT 'Created table: TaskComments';
END
GO

-- 10. TASK ATTACHMENTS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'TaskAttachments')
BEGIN
    CREATE TABLE [dbo].[TaskAttachments] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [TaskId] INT NOT NULL,
        [FileName] NVARCHAR(510) NOT NULL,
        [FilePath] NVARCHAR(1000) NOT NULL,
        [FileSize] BIGINT NOT NULL DEFAULT 0,
        [ContentType] NVARCHAR(200) NULL,
        [UploadedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        CONSTRAINT [FK_TaskAttachments_TaskItems] FOREIGN KEY ([TaskId]) REFERENCES [dbo].[TaskItems]([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_TaskAttachments_TaskId] ON [dbo].[TaskAttachments]([TaskId]);
    PRINT 'Created table: TaskAttachments';
END
GO

-- 11. USER PRESENCE (Real-time Online Status via SignalR)
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'UserPresences')
BEGIN
    CREATE TABLE [dbo].[UserPresences] (
        [UserId] NVARCHAR(100) NOT NULL PRIMARY KEY,
        [Status] INT NOT NULL DEFAULT 0, -- 0=Offline, 1=Online, 2=Away, 3=Busy
        [LastSeen] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [ConnectionId] NVARCHAR(400) NULL
    );
    PRINT 'Created table: UserPresences';
END
GO

-- 12. MEETINGS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Meetings')
BEGIN
    CREATE TABLE [dbo].[Meetings] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [Title] NVARCHAR(400) NOT NULL,
        [OrganizerId] NVARCHAR(100) NOT NULL,
        [MeetingUrl] NVARCHAR(1000) NULL,
        [ScheduledStartTime] DATETIME2 NOT NULL,
        [ScheduledEndTime] DATETIME2 NULL,
        [Status] INT NOT NULL DEFAULT 0, -- 0=Scheduled, 1=InProgress, 2=Completed, 3=Cancelled
        [MeetingMinutes] NVARCHAR(MAX) NULL,
        [IsReminderSent] BIT NOT NULL DEFAULT 0,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE()
    );
    CREATE NONCLUSTERED INDEX [IX_Meetings_OrganizerId] ON [dbo].[Meetings]([OrganizerId]);
    CREATE NONCLUSTERED INDEX [IX_Meetings_ScheduledStartTime] ON [dbo].[Meetings]([ScheduledStartTime]);
    PRINT 'Created table: Meetings';
END
GO

-- 13. MEETING PARTICIPANTS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'MeetingParticipants')
BEGIN
    CREATE TABLE [dbo].[MeetingParticipants] (
        [MeetingId] INT NOT NULL,
        [UserId] NVARCHAR(100) NOT NULL,
        [HasAttended] BIT NOT NULL DEFAULT 0,
        CONSTRAINT [PK_MeetingParticipants] PRIMARY KEY ([MeetingId], [UserId]),
        CONSTRAINT [FK_MeetingParticipants_Meetings] FOREIGN KEY ([MeetingId]) REFERENCES [dbo].[Meetings]([Id]) ON DELETE CASCADE
    );
    PRINT 'Created table: MeetingParticipants';
END
GO

-- 14. MINUTE TYPES
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'MinuteTypes')
BEGIN
    CREATE TABLE [dbo].[MinuteTypes] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [Name] NVARCHAR(200) NOT NULL,
        [IsActive] BIT NOT NULL DEFAULT 1,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE()
    );
    PRINT 'Created table: MinuteTypes';
END
GO

-- 15. MINUTE APPROVALS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'MinuteApprovals')
BEGIN
    CREATE TABLE [dbo].[MinuteApprovals] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [Date] DATETIME2 NOT NULL,
        [No] NVARCHAR(100) NOT NULL,
        [Type] NVARCHAR(100) NOT NULL,
        [Subject] NVARCHAR(MAX) NOT NULL,
        [Points] NVARCHAR(MAX) NOT NULL,
        [ForwardToUserId] NVARCHAR(100) NULL,
        [Currency] NVARCHAR(20) NULL,
        [TotalAmount] DECIMAL(18,2) NULL,
        [AdvancePercentage] DECIMAL(5,2) NULL,
        [AdvanceAmount] DECIMAL(18,2) NULL,
        [IsUrgent] BIT NOT NULL DEFAULT 0,
        [CloseByInitiator] BIT NOT NULL DEFAULT 0,
        [CreatedByUserId] NVARCHAR(100) NOT NULL,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [Status] NVARCHAR(100) NOT NULL DEFAULT 'Pending',
        [SignaturePath] NVARCHAR(MAX) NULL,
        [PurchaseOrderValue] DECIMAL(18,2) NULL,
        [PurchaseAdvanceRecommend] DECIMAL(18,2) NULL,
        [PurchaseApprovedAmount] DECIMAL(18,2) NULL,
        [RequestedStockQty] DECIMAL(18,2) NULL,
        [CurrentStockQty] DECIMAL(18,2) NULL,
        [ApprovedStockQty] DECIMAL(18,2) NULL,
        [HRLeaveType] NVARCHAR(510) NULL,
        [FinancialType] NVARCHAR(510) NULL
    );
    CREATE NONCLUSTERED INDEX [IX_MinuteApprovals_CreatedByUserId] ON [dbo].[MinuteApprovals]([CreatedByUserId]);
    CREATE NONCLUSTERED INDEX [IX_MinuteApprovals_ForwardToUserId] ON [dbo].[MinuteApprovals]([ForwardToUserId]);
    CREATE NONCLUSTERED INDEX [IX_MinuteApprovals_Date] ON [dbo].[MinuteApprovals]([Date] DESC);
    PRINT 'Created table: MinuteApprovals';
END
GO

-- 16. MINUTE ATTACHMENTS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'MinuteAttachments')
BEGIN
    CREATE TABLE [dbo].[MinuteAttachments] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [MinuteApprovalId] INT NOT NULL,
        [FileName] NVARCHAR(510) NOT NULL,
        [FilePath] NVARCHAR(1000) NOT NULL,
        [FileSize] BIGINT NOT NULL DEFAULT 0,
        [ContentType] NVARCHAR(200) NULL,
        [UploadedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        CONSTRAINT [FK_MinuteAttachments_MinuteApprovals] FOREIGN KEY ([MinuteApprovalId]) REFERENCES [dbo].[MinuteApprovals]([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_MinuteAttachments_MinuteApprovalId] ON [dbo].[MinuteAttachments]([MinuteApprovalId]);
    PRINT 'Created table: MinuteAttachments';
END
GO

-- 17. MINUTE WORKFLOW HISTORIES
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'MinuteWorkflowHistories')
BEGIN
    CREATE TABLE [dbo].[MinuteWorkflowHistories] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [MinuteApprovalId] INT NOT NULL,
        [UserId] NVARCHAR(100) NOT NULL,
        [ActionTaken] NVARCHAR(100) NOT NULL,
        [Remarks] NVARCHAR(MAX) NULL,
        [SignaturePath] NVARCHAR(MAX) NULL,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        CONSTRAINT [FK_MinuteWorkflowHistories_MinuteApprovals] FOREIGN KEY ([MinuteApprovalId]) REFERENCES [dbo].[MinuteApprovals]([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_MinuteWorkflowHistories_MinuteApprovalId] ON [dbo].[MinuteWorkflowHistories]([MinuteApprovalId]);
    PRINT 'Created table: MinuteWorkflowHistories';
END
GO

-- 18. WORKFLOW TEMPLATES
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'WorkflowTemplates')
BEGIN
    CREATE TABLE [dbo].[WorkflowTemplates] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [Name] NVARCHAR(400) NOT NULL,
        [Description] NVARCHAR(MAX) NULL,
        [RequestType] NVARCHAR(200) NOT NULL,
        [IsActive] BIT NOT NULL DEFAULT 1,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [UpdatedAt] DATETIME2 NULL
    );
    PRINT 'Created table: WorkflowTemplates';
END
GO

-- 19. WORKFLOW STEPS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'WorkflowSteps')
BEGIN
    CREATE TABLE [dbo].[WorkflowSteps] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [WorkflowTemplateId] INT NOT NULL,
        [StepOrder] INT NOT NULL,
        [StepName] NVARCHAR(MAX) NOT NULL,
        [ApproverType] INT NOT NULL, -- 0=Manager, 1=Role, 2=SpecificUser
        [SpecificUserId] NVARCHAR(100) NULL,
        [RoleName] NVARCHAR(MAX) NULL,
        [RequiredAction] INT NOT NULL,
        [SLAHours] INT NOT NULL DEFAULT 24,
        CONSTRAINT [FK_WorkflowSteps_WorkflowTemplates] FOREIGN KEY ([WorkflowTemplateId]) REFERENCES [dbo].[WorkflowTemplates]([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_WorkflowSteps_WorkflowTemplateId] ON [dbo].[WorkflowSteps]([WorkflowTemplateId]);
    PRINT 'Created table: WorkflowSteps';
END
GO

-- 20. APPROVAL REQUESTS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ApprovalRequests')
BEGIN
    CREATE TABLE [dbo].[ApprovalRequests] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [RequestNumber] NVARCHAR(100) NOT NULL,
        [Title] NVARCHAR(MAX) NOT NULL,
        [Description] NVARCHAR(MAX) NULL,
        [WorkflowTemplateId] INT NOT NULL,
        [RequesterId] NVARCHAR(100) NOT NULL,
        [DepartmentId] VARCHAR(50) NULL,
        [Priority] INT NOT NULL DEFAULT 1,
        [Status] INT NOT NULL DEFAULT 0,
        [RequestedDate] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [EffectiveDate] DATETIME2 NULL,
        [Amount] DECIMAL(18,2) NULL,
        [CostCenter] NVARCHAR(MAX) NULL,
        [Project] NVARCHAR(MAX) NULL,
        [CurrentStepOrder] INT NOT NULL DEFAULT 1,
        CONSTRAINT [FK_ApprovalRequests_WorkflowTemplates] FOREIGN KEY ([WorkflowTemplateId]) REFERENCES [dbo].[WorkflowTemplates]([Id])
    );
    CREATE NONCLUSTERED INDEX [IX_ApprovalRequests_RequesterId] ON [dbo].[ApprovalRequests]([RequesterId]);
    CREATE NONCLUSTERED INDEX [IX_ApprovalRequests_WorkflowTemplateId] ON [dbo].[ApprovalRequests]([WorkflowTemplateId]);
    CREATE NONCLUSTERED INDEX [IX_ApprovalRequests_DepartmentId] ON [dbo].[ApprovalRequests]([DepartmentId]);
    PRINT 'Created table: ApprovalRequests';
END
GO

-- 21. APPROVAL REQUEST ATTACHMENTS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ApprovalRequestAttachments')
BEGIN
    CREATE TABLE [dbo].[ApprovalRequestAttachments] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [ApprovalRequestId] INT NOT NULL,
        [FileName] NVARCHAR(MAX) NOT NULL,
        [FilePath] NVARCHAR(MAX) NOT NULL,
        [FileSize] BIGINT NOT NULL,
        [ContentType] NVARCHAR(MAX) NULL,
        [UploadedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        CONSTRAINT [FK_ApprovalRequestAttachments_ApprovalRequests] FOREIGN KEY ([ApprovalRequestId]) REFERENCES [dbo].[ApprovalRequests]([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_ApprovalRequestAttachments_ApprovalRequestId] ON [dbo].[ApprovalRequestAttachments]([ApprovalRequestId]);
    PRINT 'Created table: ApprovalRequestAttachments';
END
GO

-- 22. APPROVAL REQUEST HISTORIES
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ApprovalRequestHistories')
BEGIN
    CREATE TABLE [dbo].[ApprovalRequestHistories] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [ApprovalRequestId] INT NOT NULL,
        [StepOrder] INT NOT NULL,
        [ActionTaken] INT NOT NULL,
        [ActionById] NVARCHAR(100) NOT NULL,
        [ActionDate] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [Comments] NVARCHAR(MAX) NULL,
        CONSTRAINT [FK_ApprovalRequestHistories_ApprovalRequests] FOREIGN KEY ([ApprovalRequestId]) REFERENCES [dbo].[ApprovalRequests]([Id]) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX [IX_ApprovalRequestHistories_ApprovalRequestId] ON [dbo].[ApprovalRequestHistories]([ApprovalRequestId]);
    PRINT 'Created table: ApprovalRequestHistories';
END
GO

-- 23. APPROVAL DELEGATIONS
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ApprovalDelegations')
BEGIN
    CREATE TABLE [dbo].[ApprovalDelegations] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [DelegatorId] NVARCHAR(100) NOT NULL,
        [DelegateeId] NVARCHAR(100) NOT NULL,
        [StartDate] DATETIME2 NOT NULL,
        [EndDate] DATETIME2 NOT NULL,
        [IsActive] BIT NOT NULL DEFAULT 1,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE()
    );
    CREATE NONCLUSTERED INDEX [IX_ApprovalDelegations_DelegatorId] ON [dbo].[ApprovalDelegations]([DelegatorId]);
    CREATE NONCLUSTERED INDEX [IX_ApprovalDelegations_DelegateeId] ON [dbo].[ApprovalDelegations]([DelegateeId]);
    PRINT 'Created table: ApprovalDelegations';
END
GO

-- 24. STICKY NOTES
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'StickyNotes')
BEGIN
    CREATE TABLE [dbo].[StickyNotes] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [UserId] NVARCHAR(100) NOT NULL,
        [Content] NVARCHAR(MAX) NOT NULL,
        [Color] NVARCHAR(100) NOT NULL DEFAULT '#fffa65',
        [XPos] INT NOT NULL DEFAULT 100,
        [YPos] INT NOT NULL DEFAULT 100,
        [ReminderTime] DATETIME2 NULL,
        [IsReminderSent] BIT NOT NULL DEFAULT 0,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE()
    );
    CREATE NONCLUSTERED INDEX [IX_StickyNotes_UserId] ON [dbo].[StickyNotes]([UserId]);
    PRINT 'Created table: StickyNotes';
END
GO

-- 25. GAME SCORES
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'GameScores')
BEGIN
    CREATE TABLE [dbo].[GameScores] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [UserId] NVARCHAR(100) NOT NULL,
        [GameName] NVARCHAR(200) NOT NULL,
        [Score] INT NOT NULL,
        [PlayedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE()
    );
    CREATE NONCLUSTERED INDEX [IX_GameScores_UserId] ON [dbo].[GameScores]([UserId]);
    PRINT 'Created table: GameScores';
END
GO

-- 26. DEFAULT SEED DATA (Minute Types)
IF NOT EXISTS (SELECT * FROM [dbo].[MinuteTypes] WHERE [Name] = 'General Memo')
BEGIN
    INSERT INTO [dbo].[MinuteTypes] ([Name]) VALUES 
    ('General Memo'),
    ('Purchase Approval'),
    ('Leave Request'),
    ('Financial Clearance'),
    ('Store & Inventory Request');
    PRINT 'Seeded default MinuteTypes';
END
GO

PRINT '==============================================================================================';
PRINT 'IntraOffice clean tables created successfully on SMBI_AWM!';
PRINT '==============================================================================================';
GO

GO

-- ====================================================================================================
-- [28/49] SCRIPT: 25-09-26-1-User_Roles.sql
-- ====================================================================================================
-- [USE statement stripped for database independence]
GO

/****** Object:  Table [dbo].[User_Roles]    Script Date: 9/25/2026 1:17:09 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'User_Roles')
CREATE TABLE [dbo].[User_Roles](
	[User_Role] [varchar](50) NOT NULL,
 CONSTRAINT [PK_User_Roles] PRIMARY KEY CLUSTERED 
(
	[User_Role] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

GO

-- ====================================================================================================
-- [29/49] SCRIPT: 25-09-26-2-INSERT INTO User_Roles.sql
-- ====================================================================================================
IF NOT EXISTS (SELECT 1 FROM User_Roles WHERE User_Role = 'Stock')
INSERT INTO User_Roles(User_Role)
VALUES ('Stock')
	,('PPC')
	,('Purchaser')

GO

-- ====================================================================================================
-- [30/49] SCRIPT: 25-09-26-3-Users_User_Roles.sql
-- ====================================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'Users_User_Roles')
BEGIN
    CREATE TABLE [dbo].[Users_User_Roles](
        [EntryID] [int] IDENTITY(1,1) NOT NULL PRIMARY KEY CLUSTERED ([EntryID]),
        [UserID] [int] NOT NULL,
        [User_Role] [varchar](50) NOT NULL,
        CONSTRAINT [FK_Users_User_Roles_User_Roles] FOREIGN KEY([User_Role]) REFERENCES [dbo].[User_Roles] ([User_Role]) ON UPDATE CASCADE ON DELETE CASCADE,
        CONSTRAINT [FK_Users_User_Roles_Users] FOREIGN KEY([UserID]) REFERENCES [dbo].[Users] ([UserID]) ON UPDATE CASCADE ON DELETE CASCADE
    );
END
GO

GO

-- ====================================================================================================
-- [31/49] SCRIPT: 26-09-26-1-Users.sql
-- ====================================================================================================
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='Users' AND COLUMN_NAME='IntraOfficeMainLink')
BEGIN
    ALTER TABLE dbo.Users ADD IntraOfficeMainLink bit NOT NULL CONSTRAINT DF_Users_IntraOfficeMainLink DEFAULT 0;
END
GO

GO

-- ====================================================================================================
-- [32/49] SCRIPT: 26-09-26-2-Hub_Names.sql
-- ====================================================================================================
-- [USE statement stripped for database independence]
GO

/****** Object:  Table [dbo].[Hub_Names]    Script Date: 9/26/2026 11:19:56 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'Hub_Names')
CREATE TABLE [dbo].[Hub_Names](
	[Hub_Name] [varchar](50) NOT NULL,
 CONSTRAINT [PK_Hub_Names] PRIMARY KEY CLUSTERED 
(
	[Hub_Name] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

GO

-- ====================================================================================================
-- [33/49] SCRIPT: 26-09-26-3-INSERT INTO Hub_Names.sql
-- ====================================================================================================
INSERT INTO Hub_Names(Hub_Name)
SELECT DISTINCT Hub_Name FROM ProcessGroupsProcesses WHERE Hub_Name IS NOT NULL AND Hub_Name <> '' AND Hub_Name NOT IN (SELECT Hub_Name FROM Hub_Names);
GO

GO

-- ====================================================================================================
-- [34/49] SCRIPT: 26-09-26-4-PPC_Order_Planning_Master.sql
-- ====================================================================================================
-- 1. PPC Order Planning Master
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'PPC_Order_Planning_Master')
BEGIN
    CREATE TABLE [dbo].[PPC_Order_Planning_Master] (
        [OrderNo] VARCHAR(50) NOT NULL PRIMARY KEY,
        [PlannedBy] NVARCHAR(100) NOT NULL,
        [PlannedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [Status] VARCHAR(20) NOT NULL DEFAULT 'Planned', -- 'Draft', 'Planned', 'InExecution'
        [Notes] NVARCHAR(500) NULL,
        CONSTRAINT [FK_PpcPlanningMaster_OrderNo] FOREIGN KEY ([OrderNo]) REFERENCES [dbo].[FCustomerOrders] ([OrderNo]) ON DELETE CASCADE
    );
END

-- 2. PPC Item Planning (Overall Item Summary: Order Qty, Stock Qty, Total Purchase Qty, Production Qty)
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'PPC_Order_Item_Planning')
BEGIN
    CREATE TABLE [dbo].[PPC_Order_Item_Planning] (
        [EntryID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [OrderNo] VARCHAR(50) NOT NULL,
        [ItemID] VARCHAR(50) NOT NULL,
        [OrderQty] INT NOT NULL,
        [StockQty] INT NOT NULL DEFAULT 0,              -- Planned stock quantity to issue
        [TotalPurchaseQty] INT NOT NULL DEFAULT 0,      -- Sum of all purchase splits
        [ProductionQty] INT NOT NULL DEFAULT 0,         -- Remaining quantity to manufacture in hubs
        [Remarks] NVARCHAR(250) NULL,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [CreatedBy] NVARCHAR(100) NOT NULL,
        CONSTRAINT [FK_PpcItemPlanning_OrderNo] FOREIGN KEY ([OrderNo]) REFERENCES [dbo].[FCustomerOrders] ([OrderNo]) ON DELETE CASCADE
    );

    CREATE NONCLUSTERED INDEX [IX_PpcItemPlanning_Order_Item] 
    ON [dbo].[PPC_Order_Item_Planning] ([OrderNo], [ItemID]);
END

-- 3. PPC Item Purchases (Supports MULTIPLE purchase splits per item)
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'PPC_Order_Item_Purchases')
BEGIN
    CREATE TABLE [dbo].[PPC_Order_Item_Purchases] (
        [EntryID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [OrderNo] VARCHAR(50) NOT NULL,
        [ItemID] VARCHAR(50) NOT NULL,
        [ProcessID] INT NOT NULL,                       -- From Processes_Purchase / ItemProcesses
        [VendID] INT NULL,                              -- Maker from VendAssItems (optional)
        [PurchaseQty] INT NOT NULL,
        [PurchaseRate] FLOAT NULL DEFAULT 0,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [CreatedBy] NVARCHAR(100) NOT NULL,
        CONSTRAINT [FK_PpcItemPurchases_OrderNo] FOREIGN KEY ([OrderNo]) REFERENCES [dbo].[FCustomerOrders] ([OrderNo]) ON DELETE CASCADE
    );

    CREATE NONCLUSTERED INDEX [IX_PpcItemPurchases_Order_Item] 
    ON [dbo].[PPC_Order_Item_Purchases] ([OrderNo], [ItemID]);
END

-- 4. PPC Item Hub Schedules (Stores Start Date and End Date per Item Hub)
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'PPC_Order_Item_Hub_Schedules')
BEGIN
    CREATE TABLE [dbo].[PPC_Order_Item_Hub_Schedules] (
        [EntryID] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [OrderNo] VARCHAR(50) NOT NULL,
        [ItemID] VARCHAR(50) NOT NULL,
        [Hub_Name] VARCHAR(50) NOT NULL,
        [PlannedQty] INT NOT NULL DEFAULT 0,
        [StartDate] DATETIME NOT NULL,
        [EndDate] DATETIME NOT NULL,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [CreatedBy] NVARCHAR(100) NOT NULL,
        CONSTRAINT [FK_PpcHubSchedules_OrderNo] FOREIGN KEY ([OrderNo]) REFERENCES [dbo].[FCustomerOrders] ([OrderNo]) ON DELETE CASCADE
    );

    CREATE NONCLUSTERED INDEX [IX_PpcHubSchedules_Order_Item] 
    ON [dbo].[PPC_Order_Item_Hub_Schedules] ([OrderNo], [ItemID]);
END

GO

-- ====================================================================================================
-- [35/49] SCRIPT: 26-09-26-5-ProcessGroup_Hub_Supervisors.sql
-- ====================================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'ProcessGroup_Hub_Supervisors')
BEGIN
    CREATE TABLE [dbo].[ProcessGroup_Hub_Supervisors](
        [GroupID] [int] NOT NULL,
        [Hub_Name] [varchar](50) NOT NULL,
        [UserID] [int] NOT NULL,
        [UserName] [nvarchar](100) NOT NULL,
        [AssignedAt] [datetime2](7) NOT NULL DEFAULT (getdate()),
        CONSTRAINT [PK_ProcessGroup_Hub_Supervisors] PRIMARY KEY CLUSTERED ([GroupID] ASC, [Hub_Name] ASC, [UserID] ASC),
        CONSTRAINT [FK_PG_HubSupervisors_HubNames] FOREIGN KEY([Hub_Name]) REFERENCES [dbo].[Hub_Names] ([Hub_Name]) ON UPDATE CASCADE,
        CONSTRAINT [FK_PG_HubSupervisors_ProcessGroups] FOREIGN KEY([GroupID]) REFERENCES [dbo].[ProcessGroups] ([EntryID]) ON DELETE CASCADE,
        CONSTRAINT [FK_PG_HubSupervisors_Users] FOREIGN KEY([UserID]) REFERENCES [dbo].[Users] ([UserID])
    );
END
GO

GO

-- ====================================================================================================
-- [36/49] SCRIPT: 26-09-26-6-TaskItems.sql
-- ====================================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TaskItems') AND name = 'SourceEntityType')
BEGIN
    ALTER TABLE [dbo].[TaskItems] ADD [SourceEntityType] NVARCHAR(50) NULL;
    ALTER TABLE [dbo].[TaskItems] ADD [SourceEntityRefId] NVARCHAR(100) NULL;
    ALTER TABLE [dbo].[TaskItems] ADD [TargetRole] NVARCHAR(50) NULL;
    ALTER TABLE [dbo].[TaskItems] ADD [CompletedBy] NVARCHAR(100) NULL;
    ALTER TABLE [dbo].[TaskItems] ADD [ActionUrl] NVARCHAR(300) NULL;
    CREATE NONCLUSTERED INDEX [IX_TaskItems_SourceEntity] ON [dbo].[TaskItems]([SourceEntityType], [SourceEntityRefId]);
    CREATE NONCLUSTERED INDEX [IX_TaskItems_TargetRole_Status] ON [dbo].[TaskItems]([TargetRole], [Status]);
    PRINT 'Added workflow columns to TaskItems';
END
GO

GO

-- ====================================================================================================
-- [37/49] SCRIPT: 26-09-26-7-VFOrderList.sql
-- ====================================================================================================
-- [USE statement stripped for database independence]
GO

CREATE OR ALTER VIEW [dbo].[VFOrderList]
AS
SELECT     dbo.FCustomerOrders.OrderNo, dbo.FCustomerOrders.DT, dbo.FCustomerOrders.CustCode, dbo.FCustomerOrders.Country, dbo.F_OrderAmt(dbo.FCustomerOrders.OrderNo, 0) AS OrderAmt, 
                      dbo.ForeignCustomers.Curr, dbo.VFOrderTotalQty.TotalInvQty, dbo.VFOrderTotalQty.TotalOrderQty, dbo.FCustomerOrders.DeliveryDT, dbo.FCustomerOrders.CompanyRefID, 
                      dbo.Companies.CompanyName, dbo.VUnshippedOrderList.OrderNo AS UnshippedOrderNo, TotalShipped.TotalShippedQty, dbo.VCurrencyExchangeRates.ExchRate, 
                      dbo.FCustomerOrders.InternalRefNo, ROUND(dbo.VFOrderTotalQty.TotalWeight, 2) AS TotalWeight, TotalShipped.TotalShippedAmt, ISNULL(dbo.FCustomerFinalOrders.Cancelled, 0) AS Cancelled, 
                      TotalQuantities.TotalArticles, dbo.FCustomerFinalOrders.Remarks, TotalShipped.ShippedItemCount, TotalShipped.ShippedItemQty, TotalShipped.PartiallyShippedItemCount, 
                      TotalShipped.PartiallyShippedItemQty, TotalShipped.TotalBalanceQty, dbo.FCustomerOrders.OrderType, T1.TotalPlannedQty, dbo.FCustomerOrders.OrderPlanApproved
                      ,dbo.ForeignCustomers.LateOrderAlerts
                      ,ISNULL(dbo.FCustomerOrders.Authorized, 0) AS Authorized
                      ,dbo.FCustomerOrders.AuthorizedBy
                      ,dbo.FCustomerOrders.AuthorizedDT
FROM         dbo.FCustomerOrders INNER JOIN
                      dbo.ForeignCustomers ON dbo.FCustomerOrders.CustCode = dbo.ForeignCustomers.CustCode AND dbo.FCustomerOrders.Country = dbo.ForeignCustomers.Country INNER JOIN
                      dbo.VFOrderTotalQty ON dbo.FCustomerOrders.OrderNo = dbo.VFOrderTotalQty.OrderNo INNER JOIN
                      dbo.Companies ON dbo.FCustomerOrders.CompanyRefID = dbo.Companies.EntryID LEFT OUTER JOIN
                      dbo.VUnshippedOrderList ON dbo.FCustomerOrders.OrderNo = dbo.VUnshippedOrderList.OrderNo INNER JOIN
                          (SELECT     OrderNo, SUM(Qty) AS TotalOrderQty, COUNT(*) AS TotalArticles
                            FROM          dbo.FOrderItems
                            GROUP BY OrderNo) AS TotalQuantities ON dbo.FCustomerOrders.OrderNo = TotalQuantities.OrderNo LEFT OUTER JOIN
                      dbo.FCustomerFinalOrders ON dbo.FCustomerOrders.OrderNo = dbo.FCustomerFinalOrders.OrderNo LEFT OUTER JOIN
                          (SELECT     OrderNo, SUM(ShippedQty) AS TotalShippedQty, SUM(ShippedQty * Price) AS TotalShippedAmt, SUM(CASE WHEN ShippedQty >= Qty THEN 1 ELSE 0 END) AS ShippedItemCount, 
                                                   SUM(CASE WHEN ShippedQty >= Qty THEN ShippedQty ELSE 0 END) AS ShippedItemQty, SUM(CASE WHEN ShippedQty > 0 AND ShippedQty < Qty THEN 1 ELSE 0 END) 
                                                   AS PartiallyShippedItemCount, SUM(CASE WHEN ShippedQty > 0 AND ShippedQty < Qty THEN ShippedQty ELSE 0 END) AS PartiallyShippedItemQty, 
                                                   SUM(CASE WHEN ShippedQty >= Qty THEN 0 ELSE Qty - ShippedQty END) AS TotalBalanceQty
                            FROM          dbo.VFOrderItemswithShippedQty
                            WHERE      (CompItemCode IN
                                                       (SELECT     ItemID
                                                         FROM          dbo.Items))
                            GROUP BY OrderNo) AS TotalShipped ON dbo.FCustomerOrders.OrderNo = TotalShipped.OrderNo LEFT OUTER JOIN
                      dbo.VCurrencyExchangeRates ON dbo.ForeignCustomers.Curr = dbo.VCurrencyExchangeRates.Currency LEFT OUTER JOIN
                          (SELECT     OrderNo, SUM(Qty) AS TotalPlannedQty
                            FROM          dbo.OrderPlanningDetails
                            GROUP BY OrderNo) AS T1 ON dbo.FCustomerOrders.OrderNo = T1.OrderNo
GO

GO

-- ====================================================================================================
-- [38/49] SCRIPT: 26-09-26-8-Task_Roles.sql
-- ====================================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'Task_Roles')
BEGIN
    CREATE TABLE [dbo].[Task_Roles](
        [TaskID] [int] NOT NULL,
        [RoleName] [nvarchar](50) NOT NULL,
        [AssignedAt] [datetime2](7) NOT NULL DEFAULT (getdate()),
        CONSTRAINT [PK_Task_Roles] PRIMARY KEY CLUSTERED ([TaskID] ASC, [RoleName] ASC),
        CONSTRAINT [FK_Task_Roles_TaskItems] FOREIGN KEY([TaskID]) REFERENCES [dbo].[TaskItems] ([Id]) ON DELETE CASCADE
    );
END
GO

GO

-- ====================================================================================================
-- [39/49] SCRIPT: 26-09-26-9-Task_Assignees.sql
-- ====================================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'Task_Assignees')
BEGIN
    CREATE TABLE [dbo].[Task_Assignees](
        [TaskID] [int] NOT NULL,
        [UserID] [int] NOT NULL,
        [UserName] [nvarchar](100) NOT NULL,
        [AssignedAt] [datetime2](7) NOT NULL DEFAULT (getdate()),
        CONSTRAINT [PK_Task_Assignees] PRIMARY KEY CLUSTERED ([TaskID] ASC, [UserID] ASC),
        CONSTRAINT [FK_Task_Assignees_TaskItems] FOREIGN KEY([TaskID]) REFERENCES [dbo].[TaskItems] ([Id]) ON DELETE CASCADE,
        CONSTRAINT [FK_Task_Assignees_Users] FOREIGN KEY([UserID]) REFERENCES [dbo].[Users] ([UserID])
    );
END
GO

GO

-- ====================================================================================================
-- [40/49] SCRIPT: 26-09-26-10-Processes.sql
-- ====================================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('Processes') AND name = 'EstimatedMinutes')
BEGIN
    ALTER TABLE [dbo].[Processes] ADD [EstimatedMinutes] INT NOT NULL CONSTRAINT [DF_Processes_EstimatedMinutes] DEFAULT 0;
    PRINT 'Added column EstimatedMinutes to Processes';
END
GO

IF EXISTS (SELECT 1 FROM sys.views WHERE name = 'VProcesses')
BEGIN
    EXEC('
    CREATE OR ALTER VIEW dbo.VProcesses
    AS
    SELECT     dbo.Processes.ProcessID, dbo.Processes.SNO, dbo.Processes.Description, dbo.Processes.Supervisor, dbo.Processes.Operation, dbo.Processes.AuthRequired, 
                          dbo.Processes.Code, dbo.Processes.ProcessNameUrdu, dbo.Processes.Insp_RefID, dbo.Processes.Fix_Maker_RefID, dbo.Makers.VenderName, 
                          dbo.InspectionProcesses.Code AS Insp_Code, dbo.InspectionProcesses.ProcessName AS Insp_ProcessName, 
                          dbo.InspectionProcesses.ProcessNameUrdu AS Insp_ProcessNameUrdu, dbo.Processes.InspectionProcess, dbo.Processes.ProcessNameUrduOther, 
                          dbo.Processes.BillingProcessID,
                          ISNULL(dbo.Processes.EstimatedMinutes, 0) AS EstimatedMinutes
    FROM         dbo.Processes LEFT OUTER JOIN
                          dbo.InspectionProcesses ON dbo.Processes.Insp_RefID = dbo.InspectionProcesses.EntryID LEFT OUTER JOIN
                          dbo.Makers ON dbo.Processes.Fix_Maker_RefID = dbo.Makers.VendID
    ');
    PRINT 'Updated view: VProcesses with EstimatedMinutes';
END
GO

GO

-- ====================================================================================================
-- [41/49] SCRIPT: 26-09-26-11-VVendIssdDetail_ForRunningLots.sql
-- ====================================================================================================
-- [USE statement stripped for database independence]
GO

/****** Object:  View [dbo].[VVendIssdDetail_ForRunningLots]    Script Date: 9/28/2026 12:20:51 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE OR ALTER VIEW [dbo].[VVendIssdDetail_ForRunningLots]
AS
SELECT        dbo.VendIssued.VendID, dbo.VendIssued.DT, dbo.VendIssued.UserID, dbo.VendIssued.ProcessID, dbo.VendIssdDetail.EntryID, dbo.VendIssdDetail.RefID, 
                         dbo.VendIssdDetail.RecieptID, dbo.VendIssdDetail.ItemCode, dbo.VendIssdDetail.Rate, dbo.VendIssdDetail.IssQty, dbo.VendIssdDetail.RcvdQty, 
                         dbo.VendIssdDetail.LotNo, dbo.VendIssdDetail.ReqAuth, dbo.VendIssdDetail.OrderNo, dbo.VendIssdDetail.RcvProcessID, dbo.VendIssdDetail.Rcvd_RefID, 
                         dbo.VendIssdDetail.ReWorkLot, dbo.VendIssdDetail.Repair_RefID, dbo.VendIssued.RecieptID AS MainRecieptID, dbo.Processes.Description, dbo.Processes.Code, 
                         dbo.ItemProcesses.SNO, dbo.VendIssued.Closed, dbo.VendIssdDetail.ReturnDT, dbo.VendIssued.MasterPONo, dbo.FCustomerOrders.InternalRefNo
                         ,dbo.VendIssued.IssEmpID
FROM            dbo.VendIssued INNER JOIN
                         dbo.VendIssdDetail ON dbo.VendIssued.EntryID = dbo.VendIssdDetail.RefID INNER JOIN
                         dbo.Processes ON dbo.VendIssued.ProcessID = dbo.Processes.ProcessID INNER JOIN
                         dbo.FCustomerOrders ON dbo.VendIssdDetail.OrderNo = dbo.FCustomerOrders.OrderNo LEFT OUTER JOIN
                         dbo.ItemProcesses ON dbo.VendIssued.ProcessID = dbo.ItemProcesses.ProcessID AND dbo.VendIssdDetail.ItemCode = dbo.ItemProcesses.ItemID

GO

GO

-- ====================================================================================================
-- [42/49] SCRIPT: 26-09-26-12-VRunningLots_Simple_Main.sql
-- ====================================================================================================
-- [USE statement stripped for database independence]
GO

/****** Object:  View [dbo].[VRunningLots_Simple_Main]    Script Date: 9/28/2026 12:16:31 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO





CREATE OR ALTER VIEW [dbo].[VRunningLots_Simple_Main]
AS
SELECT        TOP (100) PERCENT T1.ItemCode, T1.LotNo, T1.OrderNo, T1.Description, T1.Qty, T1.ProcessID, T1.VendID, T1.ReWorkLot, dbo.VLotWithR2InDate.DT AS R2InDT, dbo.VItemPolishingProcessSNo.PolishingItemSNo, 
                         dbo.ItemProcesses.SNO, T1.DT, T1.LotType,T1.VRD_EntryID,T1.Lot_Available_On_ProcessID
                         ,T1.IssEmpID
FROM            (SELECT        ItemCode, LotNo, OrderNo, Description, IssQty - RcvdQty AS Qty, ProcessID, VendID, ReWorkLot, DT, 0 AS LotType,0 AS VRD_EntryID,RcvProcessID AS Lot_Available_On_ProcessID
                                ,dbo.VVendIssdDetail_ForRunningLots.IssEmpID
                          FROM            dbo.VVendIssdDetail_ForRunningLots
                          WHERE        (LotNo <> '0') AND (IssQty - RcvdQty > 0) AND (EntryID NOT IN
                                                        (SELECT        Issue_RefID
                                                          FROM            dbo.VendRcvdDetail))
                          UNION
                          SELECT        ItemCode, LotNo, OrderNo, Description, RcvdQty - IssQty - ISNULL(Wastage, 0) - ISNULL(ReWorkQty, 0) AS Qty, ProcessID, VendID, ReWorkLot, DT, 1 AS LotType,dbo.VVendRcvdDetail_Simple.EntryID AS VRD_EntryID,NextProcessID  AS Lot_Available_On_ProcessID
                                ,'' AS IssEmpID
                          FROM            dbo.VVendRcvdDetail_Simple
                          WHERE        (LotNo <> '0') AND (RcvdQty - IssQty - ISNULL(Wastage, 0) - ISNULL(ReWorkQty, 0) > 0) AND (EntryID NOT IN
                                                       (SELECT        Rcvd_RefID
                                                         FROM            dbo.VendIssdDetail)) AND (ISNULL(Opening_RefID, 0) = 0)) AS T1 LEFT OUTER JOIN
                         dbo.VLotWithR2InDate ON T1.LotNo = dbo.VLotWithR2InDate.LotNo LEFT OUTER JOIN
                         dbo.VItemPolishingProcessSNo ON T1.ItemCode = dbo.VItemPolishingProcessSNo.ItemCode LEFT OUTER JOIN
                         dbo.ItemProcesses ON T1.ItemCode = dbo.ItemProcesses.ItemID AND dbo.ItemProcesses.ProcessID = T1.ProcessID
WHERE        (T1.LotNo NOT IN
                             (SELECT        LotNo
                               FROM            dbo.Lots_Closed))
ORDER BY T1.ProcessID
GO

GO

-- ====================================================================================================
-- [43/49] SCRIPT: 26-09-26-13-VRunningLots_Simple.sql
-- ====================================================================================================
-- [USE statement stripped for database independence]
GO

/****** Object:  View [dbo].[VRunningLots_Simple]    Script Date: 9/28/2026 12:16:23 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



CREATE OR ALTER VIEW [dbo].[VRunningLots_Simple]
AS
SELECT        dbo.VRunningLots_Simple_Main.ItemCode, dbo.VRunningLots_Simple_Main.LotNo, dbo.VRunningLots_Simple_Main.OrderNo, dbo.VRunningLots_Simple_Main.Description, dbo.VRunningLots_Simple_Main.Qty, 
                         dbo.VRunningLots_Simple_Main.ProcessID, dbo.VRunningLots_Simple_Main.VendID, dbo.VRunningLots_Simple_Main.ReWorkLot, dbo.VRunningLots_Simple_Main.R2InDT, dbo.VRunningLots_Simple_Main.PolishingItemSNo, 
                         dbo.VRunningLots_Simple_Main.SNO, dbo.VRunningLots_Simple_Main.DT, T1.LotNo AS Expr1, T1.MaxSno, dbo.VRunningLots_Simple_Main.LotType,dbo.VRunningLots_Simple_Main.VRD_EntryID
                         ,dbo.VRunningLots_Simple_Main.IssEmpID
FROM            dbo.VRunningLots_Simple_Main LEFT OUTER JOIN
                             (SELECT        LotNo, MAX(SNO) AS MaxSno
                               FROM            dbo.VRunningLots_Simple_Main AS VRunningLots_Simple_Main_1
                               GROUP BY LotNo) AS T1 ON dbo.VRunningLots_Simple_Main.LotNo = T1.LotNo AND dbo.VRunningLots_Simple_Main.SNO = T1.MaxSno
GO

GO

-- ====================================================================================================
-- [44/49] SCRIPT: 26-09-26-14-DispatchListDetail_VRD.sql
-- ====================================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('DispatchListDetail_VRD') AND name = 'EntryDT')
BEGIN
    ALTER TABLE DispatchListDetail_VRD ADD EntryDT DATETIME NULL CONSTRAINT DF_DispatchListDetail_VRD_EntryDT DEFAULT GETDATE();
    PRINT 'Added EntryDT to DispatchListDetail_VRD';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('DispatchListDetail_VRD') AND name = 'AddedBy')
BEGIN
    ALTER TABLE DispatchListDetail_VRD ADD AddedBy VARCHAR(50) NULL;
    PRINT 'Added AddedBy to DispatchListDetail_VRD';
END
GO

GO

-- ====================================================================================================
-- [45/49] SCRIPT: 26-09-26-15-VHubSupervisors.sql
-- ====================================================================================================
CREATE OR ALTER VIEW [dbo].[VHubSupervisors]
AS
SELECT 
    s1.GroupID,
    s1.Hub_Name,
    STUFF((
        SELECT ', ' + s2.UserName
        FROM dbo.ProcessGroup_Hub_Supervisors s2
        WHERE s2.GroupID = s1.GroupID AND s2.Hub_Name = s1.Hub_Name
        ORDER BY s2.UserName
        FOR XML PATH(''), TYPE
    ).value('.', 'NVARCHAR(MAX)'), 1, 2, '') AS Supervisors
FROM dbo.ProcessGroup_Hub_Supervisors s1
GROUP BY s1.GroupID, s1.Hub_Name
GO

GO

-- ====================================================================================================
-- [46/49] SCRIPT: 26-09-26-16-VVendReceivingList.sql
-- ====================================================================================================
-- [USE statement stripped for database independence]
GO

/****** Object:  View [dbo].[VVendReceivingList]    Script Date: 9/29/2026 10:27:13 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE OR ALTER VIEW [dbo].[VVendReceivingList]
AS
SELECT     dbo.VendReceived.EntryID, dbo.VendReceived.VendID, dbo.VendReceived.DT, dbo.VendReceived.UserID, dbo.VendReceived.ProcessID, dbo.vendRcvdDetail.RecieptID, dbo.vendRcvdDetail.ItemCode, dbo.vendRcvdDetail.RcvdQty, dbo.vendRcvdDetail.Wastage, 
                  dbo.vendRcvdDetail.LostQty, dbo.Makers.VenderName, dbo.Makers.VendID1, CAST(CONVERT(Varchar(10), dbo.VendReceived.DT, 1) AS DATETIME) AS OnlyDT, dbo.vendRcvdDetail.ReqAuth, dbo.VendReceived.Issuance_RefID, dbo.Processes.Description, 
                  dbo.ItemProcesses.Scanning, dbo.vendRcvdDetail.LotNo, dbo.vendRcvdDetail.NextProcessID, dbo.vendRcvdDetail.IssQty, dbo.vendRcvdDetail.EntryID AS VRD_EntryID, dbo.vendRcvdDetail.Opening_RefID, dbo.Items.CatID, dbo.vendRcvdDetail.OrderNo, 
                  dbo.VendIssued.MasterPONo, dbo.vendRcvdDetail.ReWorkLot, FOrderItems.CompItemCode, dbo.vendRcvdDetail.ReWorkQty, dbo.FCustomerOrders.CustCode, dbo.Items.ItemName, dbo.Items.ItemSize, dbo.Items.SizeUnit, dbo.Items.Type, dbo.Items.GroupID, 
                  dbo.ItemGroups.Description AS ItemGroup, dbo.VendReceived.EmpID, T1.VRD_From_RefID, dbo.VendIssued.DT AS IssDT, dbo.VendIssdDetail.IssQty AS Issuance_IssQty, dbo.VendIssdDetail.RcvdQty AS Issuance_RcvdQty, 
                  dbo.VendReceived_Employees_F(dbo.VendReceived.EntryID) AS Empoloyees, TDL.VRD_RefID, dbo.VendIssdDetail.ReturnDT, dbo.MakerPostedBillsDetail_Receivings.EntryID AS MPB_D_EntryID, dbo.MakerPostedBills.BillNo, dbo.FCustomerOrders.InternalRefNo, 
                  FOrderItems.OrderQty, dbo.VendIssued.Closed, dbo.VItems_Complaints.ItemID AS ComplaintItemID, dbo.Lots_List.Batch_No, dbo.Lots_List.Mill_Certificate_No, dbo.Items.TipSize, dbo.VendReceived.EntryDT, dbo.Employees.Phone1, 
                VHubSupervisors.Hub_Name,VHubSupervisors.Supervisors
FROM        dbo.Employees RIGHT OUTER JOIN
                  dbo.VendReceived INNER JOIN
                  dbo.vendRcvdDetail ON dbo.VendReceived.EntryID = dbo.vendRcvdDetail.RefID INNER JOIN
                  dbo.Makers ON dbo.VendReceived.VendID = dbo.Makers.VendID INNER JOIN
                  dbo.Processes ON dbo.VendReceived.ProcessID = dbo.Processes.ProcessID INNER JOIN
                  dbo.Items ON dbo.vendRcvdDetail.ItemCode = dbo.Items.ItemID LEFT OUTER JOIN
                  dbo.Lots_List ON dbo.vendRcvdDetail.LotNo = dbo.Lots_List.LotNo LEFT OUTER JOIN
                  dbo.FCustomerOrders ON dbo.vendRcvdDetail.OrderNo = dbo.FCustomerOrders.OrderNo LEFT OUTER JOIN
                  dbo.VItems_Complaints ON dbo.Items.ItemID = dbo.VItems_Complaints.ItemID LEFT OUTER JOIN
                  dbo.MakerPostedBillsDetail_Receivings INNER JOIN
                  dbo.MakerPostedBills ON dbo.MakerPostedBillsDetail_Receivings.MPB_RefID = dbo.MakerPostedBills.EntryID ON dbo.vendRcvdDetail.EntryID = dbo.MakerPostedBillsDetail_Receivings.VRD_RefID LEFT OUTER JOIN
                  dbo.VendIssued INNER JOIN
                  dbo.VendIssdDetail ON dbo.VendIssued.EntryID = dbo.VendIssdDetail.RefID ON dbo.vendRcvdDetail.Issue_RefID = dbo.VendIssdDetail.EntryID LEFT OUTER JOIN
                  dbo.ItemGroups ON dbo.Items.GroupID = dbo.ItemGroups.ID ON dbo.Employees.empid = dbo.VendReceived.EmpID LEFT OUTER JOIN
                  dbo.VFOrderItems_Grouped AS FOrderItems ON dbo.vendRcvdDetail.OrderNo = FOrderItems.OrderNo AND dbo.vendRcvdDetail.ItemCode = FOrderItems.CompItemCode LEFT OUTER JOIN
                  dbo.ItemProcesses ON dbo.Processes.ProcessID = dbo.ItemProcesses.ProcessID AND dbo.vendRcvdDetail.ItemCode = dbo.ItemProcesses.ItemID AND dbo.VendReceived.ProcessID = dbo.ItemProcesses.ProcessID LEFT OUTER JOIN
                      (SELECT     VRD_From_RefID
                       FROM        dbo.LotTransferDetails
                       WHERE     (Type = 1)
                       GROUP BY VRD_From_RefID) AS T1 ON dbo.vendRcvdDetail.EntryID = T1.VRD_From_RefID LEFT OUTER JOIN
                      (SELECT     VRD_RefID
                       FROM        dbo.DispatchListDetail_VRD
                       GROUP BY VRD_RefID) AS TDL ON dbo.vendRcvdDetail.EntryID = TDL.VRD_RefID
                LEFT OUTER JOIN dbo.ItemProcessGroups ON dbo.vendRcvdDetail.ItemCode = dbo.ItemProcessGroups.ItemID 
                LEFT OUTER JOIN dbo.ProcessGroupsProcesses ON dbo.ItemProcessGroups.PG_RefID = dbo.ProcessGroupsProcesses.Group_RefID 
                                                          AND dbo.VendReceived.ProcessID = dbo.ProcessGroupsProcesses.Process_RefID 
                LEFT OUTER JOIN dbo.VHubSupervisors ON dbo.ProcessGroupsProcesses.Group_RefID = dbo.VHubSupervisors.GroupID 
                                                  AND dbo.ProcessGroupsProcesses.Hub_Name = dbo.VHubSupervisors.Hub_Name

GO

GO

-- ====================================================================================================
-- [47/49] SCRIPT: 26-09-26-17-VVendIssued.sql
-- ====================================================================================================
-- [USE statement stripped for database independence]
GO

/****** Object:  View [dbo].[VVendIssued]    Script Date: 9/29/2026 9:11:49 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO





CREATE OR ALTER VIEW [dbo].[VVendIssued]
AS
SELECT     dbo.VendIssued.VendID, dbo.VendIssued.DT, dbo.VendIssued.RecieptID, dbo.VendIssued.ItemID, dbo.Makers.VendID1, dbo.Makers.VenderName, CAST(CONVERT(Varchar(10), dbo.VendIssued.DT, 1) AS DATETIME) AS OnlyDT, TabValue.TotalValue, 
                  dbo.VendIssued.Authorized, dbo.MakerPOPromises_F(dbo.VendIssued.EntryID) AS MakerPOReturnDTs, dbo.Processes.Description, dbo.Items.ItemName, TabValue.TotalIssQty, TabValue.LotNo, dbo.VendIssued.MasterPONo, TabValue.RcvdQty, dbo.Items.CatID, 
                  TabValue.ReWorkLot, dbo.VendIssued.ProcessID, dbo.Items.GroupID, TabValue.OrderNo, dbo.VendIssued.Closed, dbo.FCustomerOrders.CustCode, TabValue.LastRcvDT, TabValue.EntryID AS DetailEntryID, dbo.FCustomerOrders.InternalRefNo, 
                  TabValue.BookMarkEntryID, dbo.VendIssued.EntryID, dbo.Items.TipSize, dbo.Items.ItemSize, dbo.Items.SizeUnit, dbo.Makers.Phone1, dbo.Makers.Phone2, dbo.Makers.CPhone, dbo.Makers.ContactPerson, dbo.Makers.Address, dbo.Makers.Mobile, 
                  dbo.Makers.CompanyName, dbo.Makers.Maker_Second_Name, dbo.Makers.MakerNameUrdu, dbo.Makers.Fax1,
				  dbo.VendIssued.IssEmpID,dbo.VEmp.Name AS Emp_Name,VendIssued.DTEntry
                  ,TabValue.Batch_No
                  ,VHubSupervisors.Hub_Name,VHubSupervisors.Supervisors
FROM        dbo.VendIssued INNER JOIN
                  dbo.Items ON dbo.VendIssued.ItemID = dbo.Items.ItemID INNER JOIN
                  dbo.Processes ON dbo.VendIssued.ProcessID = dbo.Processes.ProcessID LEFT OUTER JOIN
                  dbo.Makers ON dbo.VendIssued.VendID = dbo.Makers.VendID LEFT OUTER JOIN
                  dbo.FCustomerOrders RIGHT OUTER JOIN
                      (SELECT     dbo.VVendIssdDetailWithValueWORcving.OrderNo, dbo.VVendIssdDetailWithValueWORcving.EntryID, dbo.VVendIssdDetailWithValueWORcving.RefID, dbo.VVendIssdDetailWithValueWORcving.ReWorkLot, 
                                         dbo.VVendIssdDetailWithValueWORcving.BookMarkEntryID, SUM(dbo.VVendIssdDetailWithValueWORcving.IssValue) AS TotalValue, SUM(dbo.VVendIssdDetailWithValueWORcving.IssQty) AS TotalIssQty, 
                                         MAX(dbo.VVendIssdDetailWithValueWORcving.LotNo) AS LotNo, SUM(TRcv.RcvRcvdQty) AS RcvdQty, MAX(TRcv.RcvDT) AS LastRcvDT
                                         ,dbo.VVendIssdDetailWithValueWORcving.Batch_No
                       FROM        dbo.VVendIssdDetailWithValueWORcving LEFT OUTER JOIN
                                             (SELECT     dbo.vendRcvdDetail.Issue_RefID, MAX(dbo.VendReceived.DT) AS RcvDT, SUM(dbo.vendRcvdDetail.RcvdQty) AS RcvRcvdQty
                                              FROM        dbo.VendReceived INNER JOIN
                                                                dbo.vendRcvdDetail ON dbo.VendReceived.EntryID = dbo.vendRcvdDetail.RefID
                                              GROUP BY dbo.vendRcvdDetail.Issue_RefID) AS TRcv ON dbo.VVendIssdDetailWithValueWORcving.EntryID = TRcv.Issue_RefID
                       GROUP BY dbo.VVendIssdDetailWithValueWORcving.OrderNo, dbo.VVendIssdDetailWithValueWORcving.EntryID, dbo.VVendIssdDetailWithValueWORcving.RefID, dbo.VVendIssdDetailWithValueWORcving.ReWorkLot, 
                                         dbo.VVendIssdDetailWithValueWORcving.BookMarkEntryID,dbo.VVendIssdDetailWithValueWORcving.Batch_No) AS TabValue ON dbo.FCustomerOrders.OrderNo = TabValue.OrderNo ON dbo.VendIssued.EntryID = TabValue.RefID
				LEFT JOIN VEmp ON dbo.VendIssued.IssEmpID=VEmp.EmpID
                LEFT OUTER JOIN dbo.ItemProcessGroups ON VendIssued.ItemID = dbo.ItemProcessGroups.ItemID 
                LEFT OUTER JOIN dbo.ProcessGroupsProcesses ON dbo.ItemProcessGroups.PG_RefID = dbo.ProcessGroupsProcesses.Group_RefID 
                                                          AND dbo.VendIssued.ProcessID = dbo.ProcessGroupsProcesses.Process_RefID 
                LEFT OUTER JOIN dbo.VHubSupervisors ON dbo.ProcessGroupsProcesses.Group_RefID = dbo.VHubSupervisors.GroupID 
                                                  AND dbo.ProcessGroupsProcesses.Hub_Name = dbo.VHubSupervisors.Hub_Name
                                                  
GO

GO

-- ====================================================================================================
-- [48/49] SCRIPT: 26-09-26-18-AppNotifications.sql
-- ====================================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'AppNotifications')
BEGIN
    CREATE TABLE [dbo].[AppNotifications](
        [Id] [nvarchar](50) NOT NULL PRIMARY KEY CLUSTERED ([Id] ASC),
        [Category] [int] NOT NULL DEFAULT ((5)),
        [Title] [nvarchar](250) NOT NULL,
        [Message] [nvarchar](max) NOT NULL,
        [SenderName] [nvarchar](100) NULL,
        [TargetUserId] [nvarchar](100) NULL,
        [ActionUrl] [nvarchar](500) NOT NULL DEFAULT ('/'),
        [CreatedAt] [datetime2](7) NOT NULL DEFAULT (getutcdate()),
        [IsRead] [bit] NOT NULL DEFAULT ((0)),
        [IsReadReceipt] [bit] NOT NULL DEFAULT ((0))
    );
    CREATE NONCLUSTERED INDEX [IX_AppNotifications_TargetUser_IsRead] 
        ON [dbo].[AppNotifications] ([TargetUserId], [IsRead], [CreatedAt] DESC);
END
GO

GO

-- ====================================================================================================
-- [49/49] SCRIPT: 26-09-26-19-TaskItems.sql
-- ====================================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TaskItems') AND name = 'DueWarningSent')
BEGIN
    ALTER TABLE [dbo].[TaskItems] ADD [DueWarningSent] BIT NOT NULL CONSTRAINT [DF_TaskItems_DueWarningSent] DEFAULT 0;
    PRINT 'Added column DueWarningSent to TaskItems';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TaskItems') AND name = 'OverdueWarningSent')
BEGIN
    ALTER TABLE [dbo].[TaskItems] ADD [OverdueWarningSent] BIT NOT NULL CONSTRAINT [DF_TaskItems_OverdueWarningSent] DEFAULT 0;
    PRINT 'Added column OverdueWarningSent to TaskItems';
END
GO

GO


-- ====================================================================================================
-- [50/50] SCRIPT: 26-09-29-20-Items-VItems-Sync.sql
-- ====================================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Items') AND name = 'Fig_No')
BEGIN
    ALTER TABLE dbo.Items ADD [Fig_No] [varchar](255) NULL;
    PRINT 'Added Fig_No to Items';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Items') AND name = 'Scale')
BEGIN
    ALTER TABLE dbo.Items ADD [Scale] [varchar](255) NULL;
    PRINT 'Added Scale to Items';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Items') AND name = 'Additional_Detail')
BEGIN
    ALTER TABLE dbo.Items ADD [Additional_Detail] [varchar](4000) NULL;
    PRINT 'Added Additional_Detail to Items';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Items') AND name = 'MakerDescription')
BEGIN
    ALTER TABLE dbo.Items ADD [MakerDescription] [varchar](4000) NULL;
    PRINT 'Added MakerDescription to Items';
END
GO

CREATE OR ALTER VIEW [dbo].[VItems]
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


-- ====================================================================================================
-- [51/51] SCRIPT: 26-09-29-21-TaskItems-IntraOffice-Columns.sql
-- ====================================================================================================
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TaskItems') AND name = 'AdditionalAssigneeIds')
BEGIN
    ALTER TABLE [dbo].[TaskItems] ADD [AdditionalAssigneeIds] NVARCHAR(MAX) NULL;
    PRINT 'Added column AdditionalAssigneeIds to TaskItems';
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TaskItems') AND name = 'AssignedToNames')
BEGIN
    ALTER TABLE [dbo].[TaskItems] ADD [AssignedToNames] NVARCHAR(MAX) NULL;
    PRINT 'Added column AssignedToNames to TaskItems';
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TaskItems') AND name = 'EmailMessageSent')
BEGIN
    ALTER TABLE [dbo].[TaskItems] ADD [EmailMessageSent] BIT NOT NULL CONSTRAINT [DF_TaskItems_EmailSent] DEFAULT 0;
    PRINT 'Added column EmailMessageSent to TaskItems';
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TaskItems') AND name = 'IsRead')
BEGIN
    ALTER TABLE [dbo].[TaskItems] ADD [IsRead] BIT NOT NULL CONSTRAINT [DF_TaskItems_IsRead] DEFAULT 0;
    PRINT 'Added column IsRead to TaskItems';
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TaskItems') AND name = 'StartedAt')
BEGIN
    ALTER TABLE [dbo].[TaskItems] ADD [StartedAt] DATETIME2 NULL;
    PRINT 'Added column StartedAt to TaskItems';
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Meetings') AND name = 'VoiceNotePath')
BEGIN
    ALTER TABLE [dbo].[Meetings] ADD [VoiceNotePath] NVARCHAR(500) NULL;
    PRINT 'Added column VoiceNotePath to Meetings';
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('dbo.MinuteApprovals') AND name = 'IsRead')
BEGIN
    ALTER TABLE [dbo].[MinuteApprovals] ADD [IsRead] BIT NOT NULL CONSTRAINT [DF_MinuteApprovals_IsRead] DEFAULT 0;
    PRINT 'Added column IsRead to MinuteApprovals';
END
GO

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('dbo.MinuteApprovals') AND name = 'ReadAt')
BEGIN
    ALTER TABLE [dbo].[MinuteApprovals] ADD [ReadAt] DATETIME2 NULL;
    PRINT 'Added column ReadAt to MinuteApprovals';
END
GO


-- ====================================================================================================
-- [52/52] SCRIPT: 26-09-29-22-Processes_Purchase-Table.sql
-- ====================================================================================================
IF OBJECT_ID('dbo.Processes_Purchase', 'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[Processes_Purchase](
        [ProcessID] [int] NOT NULL,
        CONSTRAINT [PK_Processes_Purchase] PRIMARY KEY CLUSTERED ([ProcessID] ASC)
    );
    PRINT 'Created table dbo.Processes_Purchase';
END
GO

IF OBJECT_ID('dbo.Processes_Purchase', 'U') IS NOT NULL
BEGIN
    IF NOT EXISTS (SELECT 1 FROM dbo.Processes_Purchase)
    BEGIN
        INSERT INTO dbo.Processes_Purchase (ProcessID)
        VALUES 
            (1), (26), (79), (112), (119), (153), (210), (270), (271), (272),
            (273), (274), (275), (309), (310), (311), (312), (313), (314), (315),
            (316), (361);
        PRINT 'Seeded 22 standard processes into dbo.Processes_Purchase';
    END
END
GO


-- ====================================================================================================
-- [53/53] SCRIPT: 26-09-29-23-MonthlySalaries-SalaryPaid-Columns.sql
-- ====================================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.MonthlySalaries') AND name = 'OTHrs_Original')
BEGIN
    ALTER TABLE [dbo].[MonthlySalaries] ADD [OTHrs_Original] [real] NULL;
    PRINT 'Added column OTHrs_Original to MonthlySalaries';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.MonthlySalaries') AND name = 'LateHrs_Original')
BEGIN
    ALTER TABLE [dbo].[MonthlySalaries] ADD [LateHrs_Original] [real] NULL;
    PRINT 'Added column LateHrs_Original to MonthlySalaries';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.MonthlySalaries') AND name = 'OTHrs_Net')
BEGIN
    ALTER TABLE [dbo].[MonthlySalaries] ADD [OTHrs_Net] [real] NULL;
    PRINT 'Added column OTHrs_Net to MonthlySalaries';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.MonthlySalaries') AND name = 'LateHrs_Net')
BEGIN
    ALTER TABLE [dbo].[MonthlySalaries] ADD [LateHrs_Net] [real] NULL;
    PRINT 'Added column LateHrs_Net to MonthlySalaries';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.MonthlySalaries') AND name = 'Salary_Paid')
BEGIN
    ALTER TABLE [dbo].[MonthlySalaries] ADD [Salary_Paid] [bit] NOT NULL CONSTRAINT [DF_MonthlySalaries_Salary_Paid] DEFAULT ((0));
    PRINT 'Added column Salary_Paid to MonthlySalaries';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.MonthlySalaries') AND name = 'Salary_Paid_UserName')
BEGIN
    ALTER TABLE [dbo].[MonthlySalaries] ADD [Salary_Paid_UserName] [varchar](50) NULL;
    PRINT 'Added column Salary_Paid_UserName to MonthlySalaries';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.MonthlySalaries') AND name = 'Salary_Paid_MachineName')
BEGIN
    ALTER TABLE [dbo].[MonthlySalaries] ADD [Salary_Paid_MachineName] [varchar](50) NULL;
    PRINT 'Added column Salary_Paid_MachineName to MonthlySalaries';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.MonthlySalaries') AND name = 'Salary_Paid_DTEntry')
BEGIN
    ALTER TABLE [dbo].[MonthlySalaries] ADD [Salary_Paid_DTEntry] [datetime] NULL;
    PRINT 'Added column Salary_Paid_DTEntry to MonthlySalaries';
END
GO

CREATE OR ALTER VIEW [dbo].[VMonthlySalaries]
AS
SELECT        dbo.MonthlySalaries.EntryID, dbo.MonthlySalaries.EmpID, dbo.MonthlySalaries.DT, dbo.MonthlySalaries.IsForSA, dbo.MonthlySalaries.BSal, dbo.MonthlySalaries.Rate, dbo.MonthlySalaries.ADays, dbo.MonthlySalaries.AAmt, 
                         dbo.MonthlySalaries.AAllow, dbo.MonthlySalaries.AAllowAmt, dbo.MonthlySalaries.SDays, dbo.MonthlySalaries.SAmt, dbo.MonthlySalaries.Leaves, dbo.MonthlySalaries.LeaveAmt, dbo.MonthlySalaries.OHrs, 
                         dbo.MonthlySalaries.OAmt, dbo.MonthlySalaries.LHrs, dbo.MonthlySalaries.LAmt, dbo.MonthlySalaries.Total, dbo.MonthlySalaries.Tax, dbo.MonthlySalaries.NetTtl, dbo.MonthlySalaries.ShortTerm, 
                         dbo.MonthlySalaries.LongTerm, dbo.MonthlySalaries.AdvSal, dbo.MonthlySalaries.Unionfund, dbo.MonthlySalaries.Fine, dbo.MonthlySalaries.Bonus, dbo.MonthlySalaries.Lunch, dbo.MonthlySalaries.EOBI, 
                         dbo.MonthlySalaries.Balance, dbo.MonthlySalaries.PrevLTLoan, dbo.MonthlySalaries.CasualLeaves, dbo.MonthlySalaries.SickLeaves, dbo.MonthlySalaries.AnnualLeaves, dbo.MonthlySalaries.CompensatoryLeaves, 
                         dbo.MonthlySalaries.WPLeaves, dbo.MonthlySalaries.MaternityLeaves, dbo.MonthlySalaries.HrsPerDay, dbo.MonthlySalaries.TotalMonthHrs, dbo.MonthlySalaries.DTFinal, dbo.MonthlySalaries.FakeWorkingHrs, 
                         dbo.MonthlySalaries.FakeRate, dbo.MonthlySalaries.FakeSalary, dbo.MonthlySalaries.SundayOTHrs, dbo.MonthlySalaries.SundayOTRate, dbo.MonthlySalaries.FixAllowance, dbo.MonthlySalaries.HoldSalaryAmt, 
                         dbo.MonthlySalaries.PresentDays, dbo.MonthlySalaries.LeaveDays, dbo.MonthlySalaries.LateComingHrs, dbo.MonthlySalaries.ShortHrs, dbo.MonthlySalaries.AmtPaid, dbo.MonthlySalaries.GPHrs, 
                         dbo.MonthlySalaries.GPHrsAmt, dbo.MonthlySalaries.DeptID, dbo.MonthlySalaries.OTDinnerCount, dbo.MonthlySalaries.OTDinnerAmount, dbo.MonthlySalaries.DedOnePercent, dbo.MonthlySalaries.PerformanceDedAmt, 
                         dbo.MonthlySalaries.RejectionDedAmt, dbo.VEmp.name, dbo.VEmp.fname, dbo.VEmp.EmpType, dbo.VEmp.EmpAccNo, dbo.VEmp.BankPymt, dbo.VEmp.Rel, dbo.Departments.name AS DeptName, 
                         dbo.MonthlySalaries.ShortHrsAmt, dbo.MonthlySalaries.ZeroAbsentBonus, dbo.MonthlySalaries.OTHrs_Original, dbo.MonthlySalaries.LateHrs_Original, dbo.MonthlySalaries.OTHrs_Net, dbo.MonthlySalaries.LateHrs_Net
						 ,dbo.SalaryAlreadyPaid.AmtPaid AS Salary_Already_Paid_Amt
						 ,dbo.MonthlySalaries.Salary_Paid,dbo.MonthlySalaries.Salary_Paid_UserName,dbo.MonthlySalaries.Salary_Paid_MachineName,dbo.MonthlySalaries.Salary_Paid_DTEntry
FROM            dbo.MonthlySalaries INNER JOIN
                         dbo.VEmp ON dbo.MonthlySalaries.EmpID = dbo.VEmp.empid INNER JOIN
                         dbo.Departments ON dbo.MonthlySalaries.DeptID = dbo.Departments.deptid
						 LEFT JOIN dbo.SalaryAlreadyPaid ON dbo.MonthlySalaries.EmpID=dbo.SalaryAlreadyPaid.EmpID AND dbo.MonthlySalaries.DT=dbo.SalaryAlreadyPaid.ForMonth
GO

-- ====================================================================================================
-- [54/54] SCRIPT: 26-09-29-24-SMBI-Impulse-Schema-Sync.sql
-- ====================================================================================================
-- ====================================================================================================
-- Script Name: 26-09-29-24-SMBI-Impulse-Schema-Sync.sql
-- Description: Complete one-way schema synchronization from SMBI_AWM (SQL2014STD) to Impulse_AWM (SQL2017STD).
--              Synchronizes missing tables, missing columns, prerequisite views, user-defined functions,
--              views, and stored procedures.
-- Target DB:   Impulse_AWM or any target ERP database (100% Idempotent & Database-Agnostic).
-- ====================================================================================================

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

PRINT '>>> Starting Schema Synchronization from SMBI_AWM on ' + DB_NAME() + '...';
GO

-- ====================================================================================================
-- PART 1: MISSING TABLES (10)
-- ====================================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'ForeignCustomers_InvoiceTo' AND schema_id = SCHEMA_ID('dbo'))
BEGIN
    CREATE TABLE [dbo].[ForeignCustomers_InvoiceTo](
    [EntryID] [int] IDENTITY(1,1) NOT NULL,
    [CustCode] [varchar](50) NULL,
    [Country] [varchar](50) NULL,
    [InvoiceTo] [varchar](1000) NULL
    );
    PRINT 'Created table [dbo].[ForeignCustomers_InvoiceTo]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'RawMaterialIssuanceDetail_Hidden_From_Maker_Billing' AND schema_id = SCHEMA_ID('dbo'))
BEGIN
    CREATE TABLE [dbo].[RawMaterialIssuanceDetail_Hidden_From_Maker_Billing](
    [EntryID] [int] IDENTITY(1,1) NOT NULL,
    [RMID_RefID] [int] NOT NULL,
    [UserName] [varchar](50) NULL,
    [MachineName] [varchar](50) NULL,
    [DTEntry] [datetime] NULL DEFAULT (getdate())
    );
    PRINT 'Created table [dbo].[RawMaterialIssuanceDetail_Hidden_From_Maker_Billing]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'VenderAssItems_Revisions' AND schema_id = SCHEMA_ID('dbo'))
BEGIN
    CREATE TABLE [dbo].[VenderAssItems_Revisions](
    [EntryID] [int] IDENTITY(1,1) NOT NULL,
    [VAI_RefID] [int] NOT NULL,
    [Rate] [float] NULL,
    [Remarks] [varchar](255) NULL,
    [UserName] [varchar](50) NULL,
    [MachineName] [varchar](50) NULL,
    [DTEntry] [datetime] NULL DEFAULT (getdate())
    );
    PRINT 'Created table [dbo].[VenderAssItems_Revisions]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'License_Keys' AND schema_id = SCHEMA_ID('dbo'))
BEGIN
    CREATE TABLE [dbo].[License_Keys](
    [EntryID] [int] IDENTITY(1,1) NOT NULL,
    [LicenseKey] [varchar](255) NULL,
    [MachineName] [varchar](50) NULL
    );
    PRINT 'Created table [dbo].[License_Keys]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'EmailTemplates' AND schema_id = SCHEMA_ID('dbo'))
BEGIN
    CREATE TABLE [dbo].[EmailTemplates](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [TemplateCode] [nvarchar](50) NOT NULL,
    [Name] [nvarchar](150) NOT NULL,
    [Category] [nvarchar](50) NOT NULL DEFAULT ('QuotationFollowUp'),
    [SubjectTemplate] [nvarchar](255) NOT NULL,
    [BodyTemplate] [nvarchar](max) NOT NULL,
    [AvailablePlaceholders] [nvarchar](500) NULL,
    [IsActive] [bit] NOT NULL DEFAULT ((1)),
    [CreatedAt] [datetime2] NOT NULL DEFAULT (getutcdate()),
    CONSTRAINT [PK_EmailTemplates] PRIMARY KEY CLUSTERED ([Id])
    );
    PRINT 'Created table [dbo].[EmailTemplates]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'FCO_Qualities' AND schema_id = SCHEMA_ID('dbo'))
BEGIN
    CREATE TABLE [dbo].[FCO_Qualities](
    [EntryID] [int] IDENTITY(1,1) NOT NULL,
    [Quality] [varchar](50) NOT NULL,
    CONSTRAINT [PK_FCO_Qualities] PRIMARY KEY CLUSTERED ([Quality])
    );
    PRINT 'Created table [dbo].[FCO_Qualities]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'VendRcvdDetail_QC_Report' AND schema_id = SCHEMA_ID('dbo'))
BEGIN
    CREATE TABLE [dbo].[VendRcvdDetail_QC_Report](
    [VRD_RefID] [int] NOT NULL,
    [TrayNo] [varchar](50) NULL,
    [UserName] [varchar](50) NULL,
    [MachineName] [varchar](50) NULL,
    [DTEntry] [datetime] NULL DEFAULT (getdate()),
    [PassQty] [int] NULL,
    CONSTRAINT [PK_VendRcvdDetail_QC_Report] PRIMARY KEY CLUSTERED ([VRD_RefID])
    );
    PRINT 'Created table [dbo].[VendRcvdDetail_QC_Report]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'VendRcvdDetail_QC_Report_PIP' AND schema_id = SCHEMA_ID('dbo'))
BEGIN
    CREATE TABLE [dbo].[VendRcvdDetail_QC_Report_PIP](
    [VRD_RefID] [int] NOT NULL,
    [PIP_RefID] [int] NOT NULL,
    [DT] [datetime] NULL,
    [EmpID] [varchar](50) NOT NULL,
    [Pieces] [int] NULL,
    [Remarks] [varchar](255) NULL,
    [Action] [tinyint] NULL,
    CONSTRAINT [PK_VendRcvdDetail_QC_Report_PIP] PRIMARY KEY CLUSTERED ([VRD_RefID], [PIP_RefID])
    );
    PRINT 'Created table [dbo].[VendRcvdDetail_QC_Report_PIP]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'VendRcvdDetailPO_Hidden_Billing' AND schema_id = SCHEMA_ID('dbo'))
BEGIN
    CREATE TABLE [dbo].[VendRcvdDetailPO_Hidden_Billing](
    [EntryID] [int] IDENTITY(1,1) NOT NULL,
    [VRDPO_RefID] [int] NOT NULL,
    [UserName] [varchar](50) NULL,
    [MachineName] [varchar](50) NULL,
    [DTEntry] [datetime] NULL DEFAULT (getdate())
    );
    PRINT 'Created table [dbo].[VendRcvdDetailPO_Hidden_Billing]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'PrintEmpAbsentList_Summary' AND schema_id = SCHEMA_ID('dbo'))
BEGIN
    CREATE TABLE [dbo].[PrintEmpAbsentList_Summary](
    [EmpID] [varchar](50) NULL,
    [Total_Absents] [int] NULL
    );
    PRINT 'Created table [dbo].[PrintEmpAbsentList_Summary]';
END
GO
-- ====================================================================================================
-- PART 2: MISSING COLUMNS IN EXISTING TABLES
-- ====================================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[VendIssdDetail_ReturnDTs_Revisions]') AND name = 'Qty_New')
BEGIN
    ALTER TABLE [dbo].[VendIssdDetail_ReturnDTs_Revisions] ADD [Qty_New] [int] NULL;
    PRINT 'Added column [Qty_New] to [dbo].[VendIssdDetail_ReturnDTs_Revisions]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[VendIssdDetail_ReturnDTs_Revisions]') AND name = 'Qty_Old')
BEGIN
    ALTER TABLE [dbo].[VendIssdDetail_ReturnDTs_Revisions] ADD [Qty_Old] [int] NULL;
    PRINT 'Added column [Qty_Old] to [dbo].[VendIssdDetail_ReturnDTs_Revisions]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[VendIssdDetail_ReturnDTs_Revisions]') AND name = 'Rate_New')
BEGIN
    ALTER TABLE [dbo].[VendIssdDetail_ReturnDTs_Revisions] ADD [Rate_New] [real] NULL;
    PRINT 'Added column [Rate_New] to [dbo].[VendIssdDetail_ReturnDTs_Revisions]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[VendIssdDetail_ReturnDTs_Revisions]') AND name = 'Rate_Old')
BEGIN
    ALTER TABLE [dbo].[VendIssdDetail_ReturnDTs_Revisions] ADD [Rate_Old] [real] NULL;
    PRINT 'Added column [Rate_Old] to [dbo].[VendIssdDetail_ReturnDTs_Revisions]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[Employees]') AND name = 'Caste')
BEGIN
    ALTER TABLE [dbo].[Employees] ADD [Caste] [varchar](50) NULL;
    PRINT 'Added column [Caste] to [dbo].[Employees]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[Employees]') AND name = 'CNIC_PDF')
BEGIN
    ALTER TABLE [dbo].[Employees] ADD [CNIC_PDF] [image] NULL;
    PRINT 'Added column [CNIC_PDF] to [dbo].[Employees]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[Employees]') AND name = 'CNIC_PDF_FileName')
BEGIN
    ALTER TABLE [dbo].[Employees] ADD [CNIC_PDF_FileName] [varchar](1000) NULL;
    PRINT 'Added column [CNIC_PDF_FileName] to [dbo].[Employees]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[Employees]') AND name = 'Maslak')
BEGIN
    ALTER TABLE [dbo].[Employees] ADD [Maslak] [varchar](50) NULL;
    PRINT 'Added column [Maslak] to [dbo].[Employees]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[Employees]') AND name = 'Pay_Full_Salary')
BEGIN
    ALTER TABLE [dbo].[Employees] ADD [Pay_Full_Salary] [bit] NOT NULL CONSTRAINT [DF_Employees_Pay_Full_Salary] DEFAULT ((0));
    PRINT 'Added column [Pay_Full_Salary] to [dbo].[Employees]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[RM]') AND name = 'RM_Maker_Description')
BEGIN
    ALTER TABLE [dbo].[RM] ADD [RM_Maker_Description] [varchar](1000) NULL;
    PRINT 'Added column [RM_Maker_Description] to [dbo].[RM]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[RM]') AND name = 'Weight_Length')
BEGIN
    ALTER TABLE [dbo].[RM] ADD [Weight_Length] [varchar](255) NULL;
    PRINT 'Added column [Weight_Length] to [dbo].[RM]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[ForeignCustomers]') AND name = 'FC_Note_I')
BEGIN
    ALTER TABLE [dbo].[ForeignCustomers] ADD [FC_Note_I] [varchar](4000) NULL;
    PRINT 'Added column [FC_Note_I] to [dbo].[ForeignCustomers]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[ForeignCustomers]') AND name = 'FC_Note_II')
BEGIN
    ALTER TABLE [dbo].[ForeignCustomers] ADD [FC_Note_II] [varchar](4000) NULL;
    PRINT 'Added column [FC_Note_II] to [dbo].[ForeignCustomers]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[ForeignCustomers]') AND name = 'Inner_Label_Manual_I')
BEGIN
    ALTER TABLE [dbo].[ForeignCustomers] ADD [Inner_Label_Manual_I] [varchar](255) NULL;
    PRINT 'Added column [Inner_Label_Manual_I] to [dbo].[ForeignCustomers]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[ForeignCustomers]') AND name = 'Inner_Label_Manual_II')
BEGIN
    ALTER TABLE [dbo].[ForeignCustomers] ADD [Inner_Label_Manual_II] [varchar](255) NULL;
    PRINT 'Added column [Inner_Label_Manual_II] to [dbo].[ForeignCustomers]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[VendOrders]') AND name = 'Order_Collection_By')
BEGIN
    ALTER TABLE [dbo].[VendOrders] ADD [Order_Collection_By] [varchar](255) NULL;
    PRINT 'Added column [Order_Collection_By] to [dbo].[VendOrders]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[EmployeeAttendanceCatagories]') AND name = 'Saturday_Early_Out_Mins')
BEGIN
    ALTER TABLE [dbo].[EmployeeAttendanceCatagories] ADD [Saturday_Early_Out_Mins] [int] NOT NULL CONSTRAINT [DF_EmployeeAttendanceCatagories_Saturday_Early_Out_Mins] DEFAULT ((0));
    PRINT 'Added column [Saturday_Early_Out_Mins] to [dbo].[EmployeeAttendanceCatagories]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[ItemsRMComp]') AND name = 'Functional_Status')
BEGIN
    ALTER TABLE [dbo].[ItemsRMComp] ADD [Functional_Status] [tinyint] NOT NULL CONSTRAINT [DF_ItemsRMComp_Functional_Status] DEFAULT ((0));
    PRINT 'Added column [Functional_Status] to [dbo].[ItemsRMComp]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[CustomerQuotationsItems]') AND name = 'ItemRemarks')
BEGIN
    ALTER TABLE [dbo].[CustomerQuotationsItems] ADD [ItemRemarks] [varchar](max) NULL;
    PRINT 'Added column [ItemRemarks] to [dbo].[CustomerQuotationsItems]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[VenderAssItems]') AND name = 'Remarks')
BEGIN
    ALTER TABLE [dbo].[VenderAssItems] ADD [Remarks] [varchar](255) NULL;
    PRINT 'Added column [Remarks] to [dbo].[VenderAssItems]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[Lots_List]') AND name = 'Forge_Batch_No')
BEGIN
    ALTER TABLE [dbo].[Lots_List] ADD [Forge_Batch_No] [varchar](50) NULL;
    PRINT 'Added column [Forge_Batch_No] to [dbo].[Lots_List]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[Lots_List]') AND name = 'Lot_Remarks')
BEGIN
    ALTER TABLE [dbo].[Lots_List] ADD [Lot_Remarks] [varchar](4000) NULL;
    PRINT 'Added column [Lot_Remarks] to [dbo].[Lots_List]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[FPInvoice]') AND name = 'InvoiceTo')
BEGIN
    ALTER TABLE [dbo].[FPInvoice] ADD [InvoiceTo] [varchar](255) NULL;
    PRINT 'Added column [InvoiceTo] to [dbo].[FPInvoice]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[PrintSalary]') AND name = 'LateHrs_Net')
BEGIN
    ALTER TABLE [dbo].[PrintSalary] ADD [LateHrs_Net] [real] NULL;
    PRINT 'Added column [LateHrs_Net] to [dbo].[PrintSalary]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[PrintSalary]') AND name = 'LateHrs_Original')
BEGIN
    ALTER TABLE [dbo].[PrintSalary] ADD [LateHrs_Original] [real] NULL;
    PRINT 'Added column [LateHrs_Original] to [dbo].[PrintSalary]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[PrintSalary]') AND name = 'OTHrs_Net')
BEGIN
    ALTER TABLE [dbo].[PrintSalary] ADD [OTHrs_Net] [real] NULL;
    PRINT 'Added column [OTHrs_Net] to [dbo].[PrintSalary]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[PrintSalary]') AND name = 'OTHrs_Original')
BEGIN
    ALTER TABLE [dbo].[PrintSalary] ADD [OTHrs_Original] [real] NULL;
    PRINT 'Added column [OTHrs_Original] to [dbo].[PrintSalary]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[VendOrderDetail]') AND name = 'ItemRemarks')
BEGIN
    ALTER TABLE [dbo].[VendOrderDetail] ADD [ItemRemarks] [varchar](4000) NULL;
    PRINT 'Added column [ItemRemarks] to [dbo].[VendOrderDetail]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[MaterialMovements]') AND name = 'QtyMoved')
BEGIN
    ALTER TABLE [dbo].[MaterialMovements] ADD [QtyMoved] [real] NULL;
    PRINT 'Added column [QtyMoved] to [dbo].[MaterialMovements]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[PPC_Order_Item_Purchases]') AND name = 'IssuedAt')
BEGIN
    ALTER TABLE [dbo].[PPC_Order_Item_Purchases] ADD [IssuedAt] [datetime2] NULL;
    PRINT 'Added column [IssuedAt] to [dbo].[PPC_Order_Item_Purchases]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[PPC_Order_Item_Purchases]') AND name = 'IssuedBy')
BEGIN
    ALTER TABLE [dbo].[PPC_Order_Item_Purchases] ADD [IssuedBy] [nvarchar](100) NULL;
    PRINT 'Added column [IssuedBy] to [dbo].[PPC_Order_Item_Purchases]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[PPC_Order_Item_Purchases]') AND name = 'MasterPONo')
BEGIN
    ALTER TABLE [dbo].[PPC_Order_Item_Purchases] ADD [MasterPONo] [varchar](50) NULL;
    PRINT 'Added column [MasterPONo] to [dbo].[PPC_Order_Item_Purchases]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[PPC_Order_Item_Purchases]') AND name = 'VendIssued_RefID')
BEGIN
    ALTER TABLE [dbo].[PPC_Order_Item_Purchases] ADD [VendIssued_RefID] [bigint] NULL;
    PRINT 'Added column [VendIssued_RefID] to [dbo].[PPC_Order_Item_Purchases]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[Users]') AND name = 'MIL_Print_Maker_Issuance_Report_Valuewise')
BEGIN
    ALTER TABLE [dbo].[Users] ADD [MIL_Print_Maker_Issuance_Report_Valuewise] [bit] NOT NULL CONSTRAINT [DF_Users_MIL_Print_Maker_Issuance_Report_Valuewise] DEFAULT ((0));
    PRINT 'Added column [MIL_Print_Maker_Issuance_Report_Valuewise] to [dbo].[Users]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[Users]') AND name = 'MIL_Print_Master_PO')
BEGIN
    ALTER TABLE [dbo].[Users] ADD [MIL_Print_Master_PO] [bit] NOT NULL CONSTRAINT [DF_Users_MIL_Print_Master_PO] DEFAULT ((0));
    PRINT 'Added column [MIL_Print_Master_PO] to [dbo].[Users]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[Users]') AND name = 'Show_Balance_Vendor_List')
BEGIN
    ALTER TABLE [dbo].[Users] ADD [Show_Balance_Vendor_List] [bit] NOT NULL CONSTRAINT [DF_Users_Show_Balance_Vendor_List] DEFAULT ((0));
    PRINT 'Added column [Show_Balance_Vendor_List] to [dbo].[Users]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[Users]') AND name = 'Show_Values_Company_Catalog')
BEGIN
    ALTER TABLE [dbo].[Users] ADD [Show_Values_Company_Catalog] [bit] NOT NULL CONSTRAINT [DF_Users_Show_Values_Company_Catalog] DEFAULT ((0));
    PRINT 'Added column [Show_Values_Company_Catalog] to [dbo].[Users]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[CustomInvoice]') AND name = 'GDNo')
BEGIN
    ALTER TABLE [dbo].[CustomInvoice] ADD [GDNo] [varchar](50) NULL;
    PRINT 'Added column [GDNo] to [dbo].[CustomInvoice]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[CustomInvoice]') AND name = 'InvoiceTo')
BEGIN
    ALTER TABLE [dbo].[CustomInvoice] ADD [InvoiceTo] [varchar](255) NULL;
    PRINT 'Added column [InvoiceTo] to [dbo].[CustomInvoice]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[Employees_ST_Sheet_Posting]') AND name = 'Posting_Type')
BEGIN
    ALTER TABLE [dbo].[Employees_ST_Sheet_Posting] ADD [Posting_Type] [tinyint] NOT NULL CONSTRAINT [DF_Employees_ST_Sheet_Posting_Posting_Type] DEFAULT ((0));
    PRINT 'Added column [Posting_Type] to [dbo].[Employees_ST_Sheet_Posting]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[VendIssdDetail]') AND name = 'CountedBy')
BEGIN
    ALTER TABLE [dbo].[VendIssdDetail] ADD [CountedBy] [varchar](255) NULL;
    PRINT 'Added column [CountedBy] to [dbo].[VendIssdDetail]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[RunningLots_Temp]') AND name = 'TrayNo')
BEGIN
    ALTER TABLE [dbo].[RunningLots_Temp] ADD [TrayNo] [varchar](50) NULL;
    PRINT 'Added column [TrayNo] to [dbo].[RunningLots_Temp]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[vendRcvdDetail]') AND name = 'Not_Available_For_Billing')
BEGIN
    ALTER TABLE [dbo].[vendRcvdDetail] ADD [Not_Available_For_Billing] [bit] NOT NULL CONSTRAINT [DF_vendRcvdDetail_Not_Available_For_Billing] DEFAULT ((0));
    PRINT 'Added column [Not_Available_For_Billing] to [dbo].[vendRcvdDetail]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[FOrderItems]') AND name = 'IW_BatchNo')
BEGIN
    ALTER TABLE [dbo].[FOrderItems] ADD [IW_BatchNo] [varchar](50) NULL;
    PRINT 'Added column [IW_BatchNo] to [dbo].[FOrderItems]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[FOrderItems]') AND name = 'IW_OrderNo')
BEGIN
    ALTER TABLE [dbo].[FOrderItems] ADD [IW_OrderNo] [varchar](50) NULL;
    PRINT 'Added column [IW_OrderNo] to [dbo].[FOrderItems]';
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[EmpFine]') AND name = 'FineID')
BEGIN
    IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[EmpFine]') AND name = 'EntryID' AND is_identity = 1)
    BEGIN
        ALTER TABLE [dbo].[EmpFine] ADD [FineID] AS [EntryID];
        PRINT 'Added computed column [FineID] to [dbo].[EmpFine]';
    END
    ELSE
    BEGIN
        ALTER TABLE [dbo].[EmpFine] ADD [FineID] [int] IDENTITY(1,1) NOT NULL;
        PRINT 'Added identity column [FineID] to [dbo].[EmpFine]';
    END
END
GO
-- ====================================================================================================
-- PART 3: PREREQUISITE SYNCHRONIZED VIEWS
-- ====================================================================================================

CREATE OR ALTER VIEW [dbo].[VRM]
AS
SELECT        ISNULL(dbo.VRMAvgRate.AvgRate, 0) AS AvgRate, dbo.RMGroups.Description AS GroupName, dbo.RMGroups.GrpColor AS GroupColor, dbo.RM.RMID, dbo.RM.GroupID, dbo.RM.RMID1, dbo.RM.RMName, dbo.RM.Unit, 
                         dbo.RM.RMSize, dbo.RM.SizeUnit, dbo.RM.Rate, dbo.RM.RMUsage, dbo.RM.AssetAccNo, dbo.RM.SaleAccNo, dbo.RM.SRTAccNo, dbo.RM.InActive, dbo.RM.ReorderPoint, dbo.RM.MinLevel, dbo.RM.MaxLevel, 
                         dbo.RM.OpeningStock, dbo.RM.Description, dbo.RM.MaterialCode, dbo.RM.TechnicalDrawing, dbo.RM.Pic, dbo.RM.AnnealingStock, dbo.RM.MachiningStock, dbo.RM.MakerRate, dbo.RM.UrduName, Tab1.QtyInStock, 
                         dbo.RM.RequiresMaleGrinding, dbo.RM.MaleGrindingStock, dbo.RMGroupIDsForForging.Group_ID AS ForgingGroupID, Tab1.SheetsInStock, dbo.RM.Returnable, dbo.RM.Sampling, dbo.RM.RM_Maker_Description, 
                         dbo.RM.Weight_Length
FROM            dbo.RM LEFT OUTER JOIN
                         dbo.RMGroupIDsForForging ON dbo.RM.GroupID = dbo.RMGroupIDsForForging.Group_ID LEFT OUTER JOIN
                         dbo.RMGroups ON dbo.RM.GroupID = dbo.RMGroups.ID LEFT OUTER JOIN
                         dbo.VRMAvgRate ON dbo.RM.RMID = dbo.VRMAvgRate.RMID LEFT OUTER JOIN
                             (SELECT        MaterialID, SUM(QtyPlaced - QtyIssued) AS QtyInStock, SUM(ISNULL(SheetsPlaced, 0) - ISNULL(SheetsIssued, 0)) AS SheetsInStock
                               FROM            dbo.VMaterialLocationWiseStatus
                               GROUP BY MaterialID) AS Tab1 ON dbo.RM.RMID1 = Tab1.MaterialID
GO

CREATE OR ALTER VIEW [dbo].[VRMWithRMOpenPOs]
AS
SELECT        dbo.VRM.AvgRate, dbo.VRM.GroupName, dbo.VRM.GroupColor, dbo.VRM.RMID, dbo.VRM.GroupID, dbo.VRM.RMID1, dbo.VRM.RMName, dbo.VRM.RMSize, dbo.VRM.SizeUnit, dbo.VRM.Unit, dbo.VRM.Rate, 
                         dbo.VRM.RMUsage, dbo.VRM.AssetAccNo, dbo.VRM.SaleAccNo, dbo.VRM.SRTAccNo, dbo.VRM.InActive, dbo.VRM.ReorderPoint, dbo.VRM.MinLevel, dbo.VRM.MaxLevel, dbo.VRM.OpeningStock, 
                         dbo.VRM.Description, dbo.VRM.MaterialCode, dbo.VRM.TechnicalDrawing, dbo.VRM.Pic, dbo.VRM.AnnealingStock, dbo.VRM.MakerRate, dbo.VRM.MachiningStock, dbo.VRM.UrduName, dbo.VRM.QtyInStock, 
                         dbo.VRM.RequiresMaleGrinding, dbo.VRM.MaleGrindingStock, dbo.VRM.ForgingGroupID, dbo.VRM.SheetsInStock, TOpenPos.OpenPOsQty, dbo.VRM.Sampling
						 ,dbo.VRM.Weight_Length
FROM            dbo.VRM LEFT OUTER JOIN
                             (SELECT        RMID1, SUM(QtyToRcv) AS OpenPOsQty
                               FROM            dbo.VVendOrdersToRcv
                               GROUP BY RMID1) AS TOpenPos ON dbo.VRM.RMID1 = TOpenPos.RMID1
GO
CREATE OR ALTER VIEW [dbo].[VrptCustomInvoiceDetail]
AS
SELECT dbo.FOrderItems.ID, dbo.FOrderItems.ItemCode, dbo.FCustomerCatalog.Description, dbo.FCustomerCatalog.ItemID, dbo.FCustomerCatalog.CompItemID, dbo.ForeignCustomers.Curr, dbo.FCustomerCatalog.Unit, dbo.ForeignCustomers.Name, dbo.ForeignCustomers.Address, dbo.FOrderItems.OrderNo, dbo.ForeignCustomers.Cont1name, dbo.ForeignCustomers.DTFormat, 
         dbo.FCustomerCatalog.PackingMode, dbo.VItems1.CompleteItemName, dbo.FProformaOrders.Qty AS ProformaQty, dbo.ForeignCustomers.FakeAddress, dbo.ItemCatagories.Description AS ItemCatagory, dbo.ItemCatagories.SRONo, dbo.FCustomerOrders.DeliveryDT, dbo.FProformaOrders.EntryID, dbo.FCustomerOrders.DT, dbo.VItems1.ItemGroup, 
         dbo.CustomInvoice.CustomInvoice, dbo.CustomInvoiceItems.EntryID AS CustomInvoiceItemsEntryID, dbo.CustomInvoiceItems.Qty, dbo.CustomInvoiceItems.Price, dbo.CustomInvoiceItems.CustomPrice, dbo.VItems1.CatID, dbo.VItems1.ItemName, dbo.VItems1.CustomDescription, dbo.VItems1.FinishedWeight, dbo.VItems1.ItemSize, dbo.VItems1.SizeUnit, 
         dbo.VItems1.Unit AS ItemUnit, dbo.VItems1.ItemColor, dbo.CustomInvoice.Consignee, dbo.FCustomerCatalog.ItemColor AS CustomerItemColor, dbo.VItems1.Type, dbo.CustomInvoiceItems.DTRENo, dbo.CustomInvoice.DTREDescription, dbo.CustomInvoiceItems.BatchNo, dbo.VItems1.FDAListingNo, dbo.VItems1.FDAProductCode, dbo.CustomInvoice.GrossWeight, 
         dbo.BatchNosFromPackingListForInvoice_F(dbo.CustomInvoiceItems.EntryID) AS Batches, dbo.CustomInvoice.ComDrawnUnder, dbo.CustomInvoice.ComSpecial, dbo.FCustomerOrders.InternalRefNo, dbo.FProformaOrders.ItemDescription, dbo.FOrderItems.Item_Finishing_Type
		 ,dbo.FCustomerCatalog.MDMA,dbo.FCustomerCatalog.SFDA_Listing_No,dbo.FCustomerCatalog.MD_Group,
		 LEFT(ItemCatagories.HSCode,4) AS HSCode,FOrderItems.IW_BatchNo,FOrderItems.IW_OrderNo,FOrderItems.Remarks,FOrderItems.Stamps,
		 FOrderItems.Quality,VItems1.SteelType,VItems1.SteelUsed,VItems1.SteelGage,CustomInvoice.DT AS InvoiceDT
FROM  dbo.VItems1 INNER JOIN
         dbo.CustomInvoiceItems INNER JOIN
         dbo.ForeignCustomers INNER JOIN
         dbo.FProformaOrders INNER JOIN
         dbo.FOrderItems ON dbo.FProformaOrders.OrderEntryID = dbo.FOrderItems.ID INNER JOIN
         dbo.FCustomerOrders ON dbo.FOrderItems.OrderNo = dbo.FCustomerOrders.OrderNo ON dbo.ForeignCustomers.CustCode = dbo.FCustomerOrders.CustCode AND dbo.ForeignCustomers.Country = dbo.FCustomerOrders.Country INNER JOIN
         dbo.CustomInvoice INNER JOIN
         dbo.FCustomerCatalog ON dbo.CustomInvoice.CustCode = dbo.FCustomerCatalog.CustCode AND dbo.CustomInvoice.Country = dbo.FCustomerCatalog.Country ON dbo.ForeignCustomers.CustCode = dbo.CustomInvoice.CustCode AND dbo.ForeignCustomers.Country = dbo.CustomInvoice.Country ON dbo.CustomInvoiceItems.RefID = dbo.FProformaOrders.EntryID AND 
         dbo.CustomInvoiceItems.CustomInvoice = dbo.CustomInvoice.CustomInvoice ON dbo.VItems1.ItemID = dbo.FOrderItems.CompItemCode AND dbo.VItems1.ItemID = dbo.FCustomerCatalog.CompItemID INNER JOIN
         dbo.ItemCatagories ON dbo.VItems1.CatID = dbo.ItemCatagories.CatID
GO

CREATE OR ALTER VIEW [dbo].[VVenders]
AS
SELECT        dbo.Venders.VendID, dbo.Venders.AccNo, dbo.Venders.Phone1, dbo.Venders.Phone2, dbo.Venders.Phone3, dbo.Venders.Fax1, dbo.Venders.Fax2, dbo.Venders.Address, dbo.Venders.ContactPerson, dbo.Venders.CPhone, 
                         dbo.Venders.CEmail, dbo.Venders.Mobile, dbo.Venders.ImportVender, dbo.Venders.MakerNo, dbo.Venders.BankAccNo, dbo.Venders.ProcessID, dbo.Venders.VenderNameUrdu, dbo.Venders.VenderDescription, 
                         dbo.Venders.VenderPic, dbo.Venders.VenderSig, dbo.Accounts.Balance, dbo.Accounts.openbal, dbo.Accounts.opendate, dbo.Accounts.AccTitle, dbo.Accounts.Active, dbo.Venders.MakerNo AS MKNo, dbo.Accounts.SubAccOf
FROM            dbo.Venders INNER JOIN
                         dbo.Accounts ON dbo.Venders.AccNo = dbo.Accounts.AccNo
GO

CREATE OR ALTER VIEW [dbo].[VMakers]
AS
SELECT        dbo.Makers.VendID, dbo.Makers.VendID1, CASE WHEN dbo.Accounts.AccTitle IS NULL THEN VenderName ELSE Accounts.AccTitle END AS VenderName, dbo.Makers.Phone1, dbo.Makers.Phone2, dbo.Makers.Fax1, 
                         dbo.Makers.Address, dbo.Makers.ContactPerson, dbo.Makers.CPhone, dbo.Makers.CEmail, dbo.Makers.Mobile, dbo.Makers.NICNo, dbo.Makers.NTNNo, dbo.Makers.RefBy, dbo.Makers.RefByFName, dbo.Makers.RefByPhone1, 
                         dbo.Makers.RefByPhone2, dbo.Makers.RefByAddress, dbo.Makers.TimeLimit, dbo.Makers.MaxLimit, dbo.Makers.AccNo, dbo.Makers.VendType, dbo.Makers.PhaseID, dbo.Makers.SubVendType, dbo.Makers.Planter, 
                         dbo.Makers.Snaffer, dbo.Makers.Stamper, dbo.Makers.Experience, dbo.Makers.OtherCompany1, dbo.Makers.OtherCompany2, dbo.Makers.OpenBal, dbo.Makers.Active, dbo.Makers.OwnRepair, dbo.Makers.RepairVend, 
                         dbo.Makers.RepairDedRate, dbo.Makers.MakerNameUrdu, dbo.Makers.BankAccNo, dbo.Accounts.AccTitle, dbo.Processes.Description, dbo.Makers.VenderName AS MakerName, dbo.Makers.MakerType, dbo.Accounts.Balance, 
                         dbo.Makers.VendPic, dbo.Makers.VendThumb, dbo.Makers.AuthRequired, dbo.Makers.ExcessQtyPercentage, dbo.Makers.MaximumRcvingsAgainstPO, dbo.Makers.ShowRateOnPO, dbo.Makers.PaymentTerms, 
                         dbo.Accounts.SubAccOf, dbo.Makers.MakerCapacity, dbo.Makers.CompanyName, dbo.Makers.Maker_Second_Name
						 ,dbo.Makers.CNIC_PDF_FileName
FROM            dbo.Makers LEFT OUTER JOIN
                         dbo.Accounts ON dbo.Makers.AccNo = dbo.Accounts.AccNo LEFT OUTER JOIN
                         dbo.Processes ON dbo.Makers.PhaseID = dbo.Processes.ProcessID
GO

CREATE OR ALTER VIEW [dbo].[VVendersAndMakers]
AS
SELECT        T1.VendID, T1.AccNo, T1.AccTitle, T1.Phone1, T1.Phone2, T1.Fax1, T1.Address, T1.Active, T1.Vender, T1.MakerNo, T1.SubAccOf, dbo.Accounts.AccTitle AS AccTitle_Parent
FROM            (SELECT        VendID, AccNo, AccTitle, Phone1, Phone2, Fax1, Address, Active, 1 AS Vender, MakerNo, SubAccOf
                          FROM            dbo.VVenders
                          UNION
                          SELECT        VendID, AccNo, AccTitle, Phone1, Phone2, Fax1, Address, Active, 0 AS Vender, VendID1, SubAccOf
                          FROM            dbo.VMakers) AS T1 INNER JOIN
                         dbo.Accounts ON T1.SubAccOf = dbo.Accounts.AccNo
GO

CREATE OR ALTER VIEW [dbo].[VrptOrders_ForProduction_Simple]
AS
SELECT TFOrderItems.ItemCode, TFOrderItems.Qty, dbo.FCustomerOrders.OrderNo, dbo.FCustomerOrders.DT, dbo.FCustomerOrders.TradeTerms, dbo.FCustomerOrders.CustCode, dbo.FCustomerOrders.Country, dbo.FCustomerOrders.PartialShipment, dbo.FCustomerOrders.PaymentTerms, dbo.FCustomerOrders.TransShipment, dbo.FCustomerOrders.Packaging, 
         dbo.FCustomerOrders.DeliveryDT, dbo.FCustomerOrders.BatchNo, dbo.FCustomerOrders.CompanyRefID, dbo.FCustomerOrders.StampDT, 
		 dbo.FCustomerOrders.Quality, dbo.FCustomerOrders.InternalRefNo, dbo.FCustomerOrders.OrderRcvdVia, dbo.FCustomerOrders.OrderType, TFOrderItems.SetItem, TFOrderItems.OrderItemID, TFOrderItems.CustomerItemCode, 
         TFOrderItems.CompItemCode, TFOrderItems.DeliveryDT AS DeliveryDTItem,TFOrderItems.Stamps AS FOI_Stamps,TFOrderItems.IW_BatchNo,
		 TFOrderItems.Finishing,TFOrderItems.ItemOrderNo,TFOrderItems.DeliveryDate,TFOrderItems.ItemRemarks
FROM  dbo.FCustomerOrders INNER JOIN
             (SELECT dbo.FOrderItems.OrderNo, dbo.FOrderItems.ItemCode, dbo.FOrderItems.CompItemCode, 0 AS SetItem, 
			 dbo.FOrderItems.CompItemCode AS OrderItemID, SUM(dbo.FOrderItems.Qty) AS Qty, dbo.FOrderItems.ItemCode AS CustomerItemCode, 
			 dbo.FOrderItems.DeliveryDT,dbo.FOrderItems.Stamps,IW_BatchNo AS IW_BatchNo,
			 Quality as Finishing,IW_OrderNo As ItemOrderNo,DeliveryDT AS DeliveryDate,Remarks as ItemRemarks
            FROM  dbo.FOrderItems INNER JOIN
                     dbo.Items ON dbo.Items.ItemID = dbo.FOrderItems.CompItemCode
            WHERE (dbo.Items.ItemType <> 2)
            GROUP BY dbo.FOrderItems.OrderNo, dbo.FOrderItems.ItemCode, dbo.FOrderItems.CompItemCode, dbo.FOrderItems.DeliveryDT,dbo.FOrderItems.Stamps,IW_BatchNo,
			Quality,IW_OrderNo,DeliveryDT,Remarks
            UNION ALL
            SELECT FOrderItems_1.OrderNo, dbo.ItemsSets.Set_ItemID, dbo.ItemsSets.Set_ItemID AS Expr1, 1 AS SetItem, FOrderItems_1.CompItemCode AS OrderItemID, 
			SUM(FOrderItems_1.Qty * dbo.ItemsSets.Qty) AS Qty, FOrderItems_1.ItemCode AS CustomerItemCode, FOrderItems_1.DeliveryDT,FOrderItems_1.Stamps
			,IW_BatchNo AS IW_BatchNo,Quality AS Finishing,IW_OrderNo As ItemOrderNo,DeliveryDT AS DeliveryDate,Remarks as ItemRemarks
            FROM  dbo.FOrderItems AS FOrderItems_1 INNER JOIN
                     dbo.Items AS Items_1 ON FOrderItems_1.CompItemCode = Items_1.ItemID INNER JOIN
                     dbo.ItemsSets ON dbo.ItemsSets.ItemID = Items_1.ItemID
            WHERE (Items_1.ItemType = 2)
            GROUP BY FOrderItems_1.OrderNo, dbo.ItemsSets.Set_ItemID, FOrderItems_1.CompItemCode, FOrderItems_1.ItemCode, FOrderItems_1.DeliveryDT,
			FOrderItems_1.Stamps,FOrderItems_1.IW_BatchNo,Quality,IW_OrderNo,DeliveryDT,Remarks
			) AS TFOrderItems ON dbo.FCustomerOrders.OrderNo = TFOrderItems.OrderNo
GO

-- ====================================================================================================
-- PART 4: USER-DEFINED FUNCTIONS
-- ====================================================================================================
CREATE OR ALTER FUNCTION [dbo].[FSundayOTHrs_BZeeshan](@ParDate datetime)
RETURNS TABLE AS  
RETURN(
SELECT EmpID--, SUM(CASE WHEN PayableHrs>8 THEN 8 ELSE ROUND(PayableHrs,1) END) AS SundayOTHrs
	, SUM(ROUND(PayableHrs,2)+ROUND(OTHrs,2)) AS SundayOTHrs
FROM VEmpTimes2
WHERE (MONTH(dt) = MONTH(@ParDate) AND YEAR(DT) = YEAR(@ParDate) AND DT <=@ParDATE )  AND DATENAME(WEEKDAY,DT)='Sunday'
GROUP BY EmpID)
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE OR ALTER FUNCTION GetItem_Previous_Next_ProcessID_TblFn(@ProcessID INT)

RETURNS TABLE 
AS
RETURN 
(
	SELECT T1.* FROM (
		SELECT ItemID
			,LAG(ProcessID) OVER(PARTITION BY ItemID ORDER BY SNo) AS Previous_ProcessID
			,ProcessID AS Current_ProcessID 
			,LEAD(ProcessID) OVER(PARTITION BY ItemID ORDER BY SNo) AS Next_ProcessID
		FROM ItemProcesses --WHERE ProcessID=@ProcessID
		) T1 WHERE Current_ProcessID=@ProcessID
)
GO

CREATE OR ALTER FUNCTION [dbo].[RunningLots] (@OrderNo AS VARCHAR (255),@ItemID AS VARCHAR (255))
RETURNS VARCHAR(1000) AS

BEGIN
DECLARE @ReturnStr AS VARCHAR(1000)
SET @ReturnStr=''
DECLARE @LotNo AS VARCHAR(255),@Description AS VARCHAR(255),@Qty INT
--DECLARE Abc CURSOR FOR SELECT LotNo,Description,Qty FROM VRunningLots_Simple WHERE OrderNo=@OrderNo AND ItemCode=@ItemID
DECLARE Abc CURSOR FOR SELECT DISTINCT Description FROM VRunningLots_Simple WHERE OrderNo=@OrderNo AND ItemCode=@ItemID
OPEN Abc

--FETCH NEXT FROM Abc INTO @LotNo,@Description,@Qty
FETCH NEXT FROM Abc INTO @Description
WHILE @@Fetch_Status=0
	BEGIN
		SET @ReturnStr=@ReturnStr + @Description + ',' 
		--FETCH NEXT FROM Abc INTO @LotNo,@Description,@Qty
		FETCH NEXT FROM Abc INTO @Description
	END
	IF LEN(@ReturnStr)>0 
		SET @ReturnStr = LEFT(@ReturnStr,LEN(@ReturnStr)-1)
CLOSE Abc
DEALLOCATE Abc

RETURN @ReturnStr

END
GO

CREATE OR ALTER FUNCTION [dbo].[Invoice_DispatchListNos_Fn] (@InvoiceNo VARCHAR(50))
RETURNS VARCHAR(1000) AS

BEGIN
DECLARE @ReturnStr AS VARCHAR(1000)
SET @ReturnStr=''
DECLARE @Temp AS VARCHAR(255)
DECLARE Abc CURSOR FOR SELECT DISTINCT DispatchListNo FROM CustomPList_DispatchListDetail
						INNER JOIN CustomPList ON CustomPList_DispatchListDetail.CustomPList_RefID=CustomPList.ID
						INNER JOIN DispatchList ON CustomPList_DispatchListDetail.DP_RefID=DispatchList.EntryID
						WHERE CustomPList.CustomInvoice=@InvoiceNo
OPEN Abc

FETCH NEXT FROM Abc into @Temp
WHILE @@Fetch_Status=0
	BEGIN
		SET @ReturnStr=@ReturnStr + @Temp + ',' 
		FETCH NEXT FROM Abc Into @Temp
	END
	IF LEN(@ReturnStr)>0 
		SET @ReturnStr = LEFT(@ReturnStr,LEN(@ReturnStr)-1)
CLOSE Abc
DEALLOCATE Abc

RETURN @ReturnStr

END
GO

CREATE OR ALTER FUNCTION [dbo].[Lot_State_For_Inner_Labels_F] (@EntryID AS INT,@ShowAll AS BIT=0)
RETURNS VARCHAR(1000) AS

BEGIN
DECLARE @Desc AS VARCHAR(4000)
SET @Desc=''
DECLARE  @Process AS VARCHAR(255)
/*DECLARE @Temper_Process_ID VARCHAR(50),@Inspection_Process_ID VARCHAR(50)
SET @Temper_Process_ID=dbo.GetGeneralDataValue_F('Temper_Inspection_ProcessID')
SET @Inspection_Process_ID=dbo.GetGeneralDataValue_F('First_Inspection_ProcessID')*/

IF @ShowAll=0

	DECLARE myCursor CURSOR FOR SELECT TOP 1 Processes.Description FROM DispatchListDetail_VRD 
	INNER JOIN VendRcvdDetail ON DispatchListDetail_VRD.VRD_RefID=VendRcvdDetail.EntryID 
	INNER JOIN DispatchListDetails_Adv ON DispatchListDetail_VRD.DLD_refID=DispatchListDetails_Adv.EntryID
	INNER JOIN DispatchList ON DispatchList.EntryID=DispatchListDetails_Adv.RefID
	INNER JOIN DispatchListDetail_Inners ON DispatchListDetails_Adv.EntryID=DispatchListDetail_Inners.RefID	
	INNER JOIN Processes ON VendRcvdDetail.ProcessID=Processes.ProcessID
	WHERE DLDC_RefID=@EntryID AND DispatchList.NewFormat=1 ORDER BY DispatchListDetail_Inners.Qty DESC

ELSE

	DECLARE myCursor CURSOR FOR SELECT Processes.Description FROM DispatchListDetail_VRD 
	INNER JOIN VendRcvdDetail ON DispatchListDetail_VRD.VRD_RefID=VendRcvdDetail.EntryID 
	INNER JOIN DispatchListDetails_Adv ON DispatchListDetail_VRD.DLD_refID=DispatchListDetails_Adv.EntryID
	INNER JOIN DispatchList ON DispatchList.EntryID=DispatchListDetails_Adv.RefID
	INNER JOIN DispatchListDetail_Inners ON DispatchListDetails_Adv.EntryID=DispatchListDetail_Inners.RefID	
	INNER JOIN Processes ON VendRcvdDetail.ProcessID=Processes.ProcessID
	WHERE DLDC_RefID=@EntryID AND DispatchList.NewFormat=1 ORDER BY DispatchListDetail_Inners.Qty DESC

	
OPEN myCursor

FETCH NEXT FROM myCursor INTO @Process
WHILE @@Fetch_Status=0
	BEGIN
		SET @Desc=@Desc +  @Process + ','
		FETCH NEXT FROM myCursor INTO @Process
	END
	IF LEN(@Desc)>0
		SET @Desc = LEFT(@Desc,LEN(@Desc)-1)
CLOSE myCursor
DEALLOCATE myCursor

RETURN @Desc

END
GO

CREATE OR ALTER FUNCTION [dbo].[F_CustomInvoice_HSCodes] (@CustomInvoice as Varchar(50))  
RETURNS Varchar(1000) AS

Begin
Declare @HSCodes As Varchar(1000)
Set @HSCodes=''
Declare @Temp As Varchar(255)
Declare Abc Cursor For Select Distinct HSCode From VrptCustomInvoiceDetail Where CustomInvoice=@CustomInvoice
Open Abc

Fetch Next From Abc Into @Temp
While @@Fetch_Status=0
	Begin
		Set @HSCodes=@HSCodes + ',' + @Temp
		Fetch Next From Abc Into @Temp
	End
	IF LEN(@HSCodes)>0 
		Set @HSCodes = Right(@HSCodes,Len(@HSCodes)-1)
Close Abc
Deallocate Abc

Return @HSCodes

End
GO

-- ====================================================================================================
-- PART 5: VIEWS
-- ====================================================================================================
CREATE OR ALTER VIEW dbo.VFOrderItems_PTC
AS
SELECT DISTINCT dbo.FOrderItems.OrderNo, dbo.FOrderItems.ItemCode, T2.PTC
FROM        dbo.FOrderItems LEFT OUTER JOIN
                      (SELECT     dbo.vendRcvdDetail.OrderNo, dbo.vendRcvdDetail.ItemCode, dbo.vendRcvdDetail.LotNo + ' (' + CAST(dbo.vendRcvdDetail.RcvdQty AS VARCHAR(5)) + ')' AS PTC
                       FROM        dbo.vendRcvdDetail WITH (NOLOCK) INNER JOIN
                                             (SELECT     MAX(EntryID) AS MAXEntryID
                                              FROM        dbo.vendRcvdDetail WITH (NOLOCK)
                                              GROUP BY LotNo) AS T1 ON dbo.vendRcvdDetail.EntryID = T1.MAXEntryID) AS T2 ON dbo.FOrderItems.OrderNo = T2.OrderNo AND dbo.FOrderItems.ItemCode = T2.ItemCode
GO

CREATE OR ALTER VIEW [dbo].[VEmpOTHrs_New]
AS
SELECT     EmpID, DT, SUM(PayableHrs) AS ValidOTHrs
FROM         dbo.EmpTimes
WHERE     (OverTime = 1) AND  (DateName(WeekDay,DT)<>'Sunday')
AND DT NOT IN (select DT FROM Holidays)
GROUP BY EmpID, DT
UNION ALL
SELECT     EmpID, DT, SUM(PayableHrs) AS ValidOTHrs
FROM         dbo.EmpTimes
WHERE (DateName(WeekDay,DT)<>'Sunday')
AND DT IN (select DT FROM Holidays)
GROUP BY EmpID, DT
GO

CREATE OR ALTER VIEW [dbo].[VClaims_Statistics]
AS
SELECT DispatchListNo,CustCode,DT,TotalQty AS DispatchQty,0 AS ClaimQty FROM VDispatchList
UNION ALL
SELECT CustomerComplaints.ComplaintNo,CustomerComplaints.CustCode,CustomerComplaints.DT,0,CustomerComplaints_Detail.Qty FROM CustomerComplaints 
	INNER JOIN CustomerComplaints_Detail ON CustomerComplaints.EntryID=CustomerComplaints_Detail.RefID
GO

CREATE OR ALTER VIEW [dbo].[VFOrderItems_Lots]
AS
SELECT        CustCode, Country, OrderNo, InternalRefNo, CompItemCode, DeliveryDT, Qty, ShippedQty, DeliveryStatus, ItemName, ID, DT, ItemCode, Description, Packaging, OrderDeliveryDT, Quality, ItemQuality, GroupID, Remarks, SortNo, 
                         Price, Item_Finishing_Type, dbo.RunningLots(OrderNo, CompItemCode) AS Lot_Processes
FROM            dbo.VFOrderItems
GO

CREATE OR ALTER VIEW [dbo].[VCustomPList_Itemwise]
AS
SELECT     TOP (100) PERCENT MIN(CartonFrom) AS FromCartonNo, MAX(CartonTo) AS ToCartonNo, CustomInvoice, OrderItemID, SUM(Qty) AS Expr1, COUNT(*) AS Tots
FROM        dbo.CustomPList
--WHERE     (CustomInvoice = '2322311')
GROUP BY CustomInvoice, OrderItemID
ORDER BY FromCartonNo
GO

CREATE OR ALTER VIEW [dbo].[VItemsRMCompWithLocationWiseStatus]
AS
SELECT        dbo.ItemsRMComp.EntryID, dbo.ItemsRMComp.ItemID, dbo.ItemsRMComp.RMID, dbo.ItemsRMComp.Qty, dbo.ItemsRMComp.ProcessID, dbo.VRM.RMName AS Description, dbo.Processes.Description AS ProcDesc, 
                         dbo.VRM.Description AS GroupDescription, dbo.Items.ItemName, dbo.VRM.RMID1, dbo.VRM.GroupID, 
						 dbo.VRM.QtyInStock,dbo.ItemsRMComp.Functional_Status,
						 VMaterialLocationWiseStatus.QtyRcvd,VMaterialLocationWiseStatus.QtyIssued,VMaterialLocationWiseStatus.QtyPlaced,
						 VMaterialLocationWiseStatus.StoreName,VMaterialLocationWiseStatus.RackNo,VMaterialLocationWiseStatus.ShelfNo

FROM            dbo.ItemsRMComp INNER JOIN
                         dbo.VRM ON dbo.ItemsRMComp.RMID = dbo.VRM.RMID INNER JOIN
                         dbo.Items ON dbo.ItemsRMComp.ItemID = dbo.Items.ItemID LEFT OUTER JOIN
                         dbo.Processes ON dbo.ItemsRMComp.ProcessID = dbo.Processes.ProcessID
						 LEFT OUTER JOIN VMaterialLocationWiseStatus ON VRM.RMID1=VMaterialLocationWiseStatus.MaterialID
GO

CREATE OR ALTER VIEW dbo.VItem_RunningLots
AS
SELECT     ItemCode, COUNT(*) AS RunningLots, SUM(Qty) AS Qty
FROM        dbo.VRunningLots
GROUP BY ItemCode
GO

CREATE OR ALTER VIEW dbo.VCustomInvoiceDetail_From_PList_Total
AS
SELECT     CustomInvoice, OrderNo, CompItemID, SUM(Qty) AS TotalQty
FROM        dbo.VCustomInvoiceDetail_From_PList
GROUP BY CustomInvoice, OrderNo, CompItemID
GO

CREATE OR ALTER VIEW dbo.VMakerBalance_Long
AS
SELECT     AccNo, SUM(Amount) AS Balance
FROM        dbo.VMakerAdvancesLedger
GROUP BY AccNo
GO

CREATE OR ALTER VIEW [dbo].[VItemOpenPOs_itemwise]
AS
SELECT ItemCode,SUM(Qty) AS OpenPOQty FROM VItemOpenPOs GROUP BY ItemCode
GO

CREATE OR ALTER VIEW dbo.VRunningLots_StockOrder
AS
SELECT     OrderNo, ItemCode, RunningQty
FROM        dbo.VRunningLots_Summary
WHERE     (OrderNo IN ('Stock-Order', 'Stock-OrderIss'))
GO

CREATE OR ALTER VIEW dbo.VVendIssued_rpt
AS
SELECT     dbo.VendIssued.EntryID, dbo.VendIssued.VendID, dbo.VendIssued.DT, dbo.VendIssued.RecieptID, dbo.VendIssued.UserID, dbo.VendIssued.ProcessID, dbo.VendIssued.ItemID, dbo.VendIssued.UserName, dbo.VendIssued.MachineName, 
                  dbo.VendIssued.SpecialInstructions, dbo.VendIssued.VchrNo, dbo.VendIssued.Authorized, dbo.VendIssued.ExcessQtyPercentage, dbo.VendIssued.MaximumRcvingsAgainstPO, dbo.VendIssued.ProcessOrderNo, dbo.VendIssued.MasterPONo, 
                  dbo.VendIssued.SampleProvided, dbo.VendIssued.DrawingProvided, dbo.VendIssued.ForgingProvided, dbo.VendIssued.SteelProvided, dbo.VendIssued.Closed, dbo.VendIssued.IssEmpID, dbo.VendIssued.SteelType_RefID, dbo.VendIssued.AuthUserName, 
                  dbo.VendIssued.AuthMachineName, dbo.VendIssued.AuthEntryDT, ISNULL(T1.Posted, 0) AS Posted
FROM        dbo.VendIssued LEFT OUTER JOIN
                      (SELECT     VI_RefID, SUM(CASE WHEN ISNULL(VI_RefID, 0) > 0 THEN 1 ELSE 0 END) AS Posted
                       FROM        dbo.RawMaterialIssuance
                       GROUP BY VI_RefID) AS T1 ON dbo.VendIssued.EntryID = T1.VI_RefID
GO

CREATE OR ALTER VIEW [dbo].[VClosedLots_Simple_Main]
AS
SELECT        DISTINCT T1.ItemCode, T1.OrderNo, T1.LotNo
FROM            (SELECT        ItemCode, LotNo, OrderNo, Description, IssQty - RcvdQty AS Qty, ProcessID, VendID, ReWorkLot, DT, 0 AS LotType,0 AS VRD_EntryID
                          FROM            dbo.VVendIssdDetail_ForRunningLots
                          WHERE        (LotNo <> '0') AND (IssQty - RcvdQty > 0) AND (EntryID NOT IN
                                                        (SELECT        Issue_RefID
                                                          FROM            dbo.VendRcvdDetail))
                          UNION
                          SELECT        ItemCode, LotNo, OrderNo, Description, RcvdQty - IssQty - ISNULL(Wastage, 0) - ISNULL(ReWorkQty, 0) AS Qty, ProcessID, VendID, ReWorkLot, DT, 1 AS LotType,dbo.VVendRcvdDetail_Simple.EntryID AS VRD_EntryID
                          FROM            dbo.VVendRcvdDetail_Simple
                          WHERE        (LotNo <> '0') AND (RcvdQty - IssQty - ISNULL(Wastage, 0) - ISNULL(ReWorkQty, 0) > 0) AND (EntryID NOT IN
                                                       (SELECT        Rcvd_RefID
                                                         FROM            dbo.VendIssdDetail)) AND (ISNULL(Opening_RefID, 0) = 0)) AS T1 LEFT OUTER JOIN
                         dbo.VLotWithR2InDate ON T1.LotNo = dbo.VLotWithR2InDate.LotNo LEFT OUTER JOIN
                         dbo.VItemPolishingProcessSNo ON T1.ItemCode = dbo.VItemPolishingProcessSNo.ItemCode LEFT OUTER JOIN
                         dbo.ItemProcesses ON T1.ItemCode = dbo.ItemProcesses.ItemID AND dbo.ItemProcesses.ProcessID = T1.ProcessID
WHERE        (T1.LotNo IN
                             (SELECT        LotNo
                               FROM            dbo.Lots_Closed))
--ORDER BY T1.ProcessID
GO

CREATE OR ALTER VIEW [dbo].[VItemsWithShelfWiseStock_Old]
AS
SELECT dbo.VStoreShelfs.StoreName, dbo.VStoreShelfs.RackNo, dbo.VStoreShelfs.ShelfNo, T1.ItemID, T1.RcvdQty - ISNULL(T1.IssdQty, 0) AS NetQty, T1.Shelf_RefID, T2.Remarks, dbo.VStoreShelfs.Store_RefID
FROM  dbo.VStoreShelfs INNER JOIN
             (SELECT ItemID, Shelf_RefID, SUM(RcvdQty) AS RcvdQty, SUM(IssdQty) AS IssdQty
            FROM  (SELECT dbo.RcvItemsSimpleDetail.ItemID, dbo.RcvItemsSimpleDetail_Placement.Shelf_RefID, SUM(dbo.RcvItemsSimpleDetail_Placement.RcvdQty) AS RcvdQty, 0 AS IssdQty
                     FROM  dbo.RcvItemsSimpleDetail_Placement INNER JOIN
                              dbo.RcvItemsSimpleDetail ON dbo.RcvItemsSimpleDetail_Placement.RISD_RefID = dbo.RcvItemsSimpleDetail.EntryID
                     GROUP BY dbo.RcvItemsSimpleDetail.ItemID, dbo.RcvItemsSimpleDetail_Placement.Shelf_RefID
                     UNION ALL
                     SELECT dbo.IssItemsSimpleDetail.ItemID, dbo.IssItemsSimpleDetail_More.Shelf_RefID, 0 AS RcvdQty, SUM(dbo.IssItemsSimpleDetail_More.IssdQty) AS IssdQty
                     FROM  dbo.IssItemsSimpleDetail INNER JOIN
                              dbo.IssItemsSimpleDetail_More ON dbo.IssItemsSimpleDetail.EntryID = dbo.IssItemsSimpleDetail_More.IISD_RefID
                     GROUP BY dbo.IssItemsSimpleDetail.ItemID, dbo.IssItemsSimpleDetail_More.Shelf_RefID) AS TInner1
            GROUP BY ItemID, Shelf_RefID) AS T1 ON dbo.VStoreShelfs.EntryID = T1.Shelf_RefID LEFT OUTER JOIN
             (SELECT dbo.VRcvitemsSimpleDetail_Placement.ItemID, dbo.VRcvitemsSimpleDetail_Placement.Shelf_RefID, dbo.VRcvitemsSimpleDetail_Placement.Remarks
            FROM  dbo.VRcvitemsSimpleDetail_Placement INNER JOIN
                         (SELECT MAX(EntryID) AS MaxEntryID
                        FROM  dbo.VRcvitemsSimpleDetail_Placement AS VRcvitemsSimpleDetail_Placement_1
                        GROUP BY ItemID, Shelf_RefID) AS TMax ON dbo.VRcvitemsSimpleDetail_Placement.EntryID = TMax.MaxEntryID) AS T2 ON T1.ItemID = T2.ItemID AND T1.Shelf_RefID = T2.Shelf_RefID
WHERE (T1.RcvdQty - ISNULL(T1.IssdQty, 0) > 0)
GO

CREATE OR ALTER VIEW [dbo].[VLots_With_SOADetails]
AS
SELECT        TLots.LotNo, dbo.RcvItemsSimpleDetail.LotNo_Manual AS LotNo_SOA, dbo.Lots_List.Batch_No AS BatchNo_SOA, dbo.Lots_List.Mill_Certificate_No AS Mil_Certificate_No_SOA, 
                         dbo.StockOrderAdjustmentsDetail.Qty AS Qty_SOA
FROM            dbo.StockOrderAdjustments INNER JOIN
                             (SELECT        LotNo, MIN(EntryID) AS MinEntryID
                               FROM            dbo.VendRcvdDetail
                               GROUP BY LotNo) AS TLots ON dbo.StockOrderAdjustments.VID_RefID = TLots.MinEntryID INNER JOIN
                         dbo.StockOrderAdjustmentsDetail ON dbo.StockOrderAdjustments.EntryID = dbo.StockOrderAdjustmentsDetail.SOA_RefID INNER JOIN
                         dbo.RcvItemsSimpleDetail ON dbo.StockOrderAdjustmentsDetail.RCV_ISD_RefID = dbo.RcvItemsSimpleDetail.EntryID INNER JOIN
                         dbo.Lots_List ON dbo.RcvItemsSimpleDetail.LotNo_Manual = dbo.Lots_List.LotNo
GO

CREATE OR ALTER VIEW [dbo].[VVendRcvdDetail_QC_Report_PIP]
AS
SELECT dbo.VendRcvdDetail_QC_Report_PIP.VRD_RefID, dbo.VendRcvdDetail_QC_Report_PIP.PIP_RefID, dbo.VendRcvdDetail_QC_Report_PIP.DT, dbo.VendRcvdDetail_QC_Report_PIP.EmpID, dbo.VendRcvdDetail_QC_Report_PIP.Pieces, dbo.VendRcvdDetail_QC_Report_PIP.Remarks, dbo.VendRcvdDetail_QC_Report_PIP.Action, dbo.Process_Inspection_Points.Point_Description, 
         dbo.VEmp.name, dbo.VEmp.Designation, dbo.VEmp.DeptName
FROM  dbo.VendRcvdDetail_QC_Report_PIP INNER JOIN
         dbo.Process_Inspection_Points ON dbo.VendRcvdDetail_QC_Report_PIP.PIP_RefID = dbo.Process_Inspection_Points.EntryID LEFT OUTER JOIN
         dbo.VEmp ON dbo.VendRcvdDetail_QC_Report_PIP.EmpID = dbo.VEmp.empid
GO

CREATE OR ALTER VIEW [dbo].[VRunningLots_Simple_WPP]
AS
SELECT        dbo.VRunningLots_Simple_Main.ItemCode, dbo.VRunningLots_Simple_Main.LotNo, dbo.VRunningLots_Simple_Main.OrderNo, dbo.VRunningLots_Simple_Main.Description
						,dbo.VRunningLots_Simple_Main.Qty,dbo.VRunningLots_Simple_Main.ProcessID, dbo.VRunningLots_Simple_Main.VendID, dbo.VRunningLots_Simple_Main.ReWorkLot
						,dbo.VRunningLots_Simple_Main.R2InDT, dbo.VRunningLots_Simple_Main.PolishingItemSNo,dbo.VRunningLots_Simple_Main.SNO, dbo.VRunningLots_Simple_Main.DT
						,T1.MaxSno, dbo.VRunningLots_Simple_Main.LotType,dbo.VRunningLots_Simple_Main.VRD_EntryID
						,dbo.VRunningLots_Simple_Main.Lot_Available_On_ProcessID
FROM            dbo.VRunningLots_Simple_Main LEFT OUTER JOIN
                             (SELECT        LotNo, MAX(SNO) AS MaxSno
                               FROM            dbo.VRunningLots_Simple_Main AS VRunningLots_Simple_Main_1
                               GROUP BY LotNo) AS T1 ON dbo.VRunningLots_Simple_Main.LotNo = T1.LotNo AND dbo.VRunningLots_Simple_Main.SNO = T1.MaxSno
GO

CREATE OR ALTER VIEW [dbo].[VLot_With_TrayNo]
AS
SELECT dbo.vendRcvdDetail.LotNo, dbo.VendRcvdDetail_QC_Report.TrayNo
FROM  dbo.VendRcvdDetail_QC_Report INNER JOIN
         dbo.vendRcvdDetail ON dbo.VendRcvdDetail_QC_Report.VRD_RefID = dbo.vendRcvdDetail.EntryID
GO

CREATE OR ALTER VIEW [dbo].[VLot_With_ForgeBatchNo]
AS
SELECT dbo.vendRcvdDetail.LotNo, dbo.VendRcvdDetailPO.RcvID
FROM  dbo.VendIssdDetail INNER JOIN
         dbo.VendIssued ON dbo.VendIssued.EntryID = dbo.VendIssdDetail.RefID INNER JOIN
         dbo.RawMaterialIssuance ON dbo.VendIssued.EntryID = dbo.RawMaterialIssuance.VI_RefID INNER JOIN
         dbo.RawMaterialIssuanceDetail ON dbo.RawMaterialIssuance.IssNo = dbo.RawMaterialIssuanceDetail.IssNo INNER JOIN
         dbo.RMID_MLS_Details ON dbo.RawMaterialIssuanceDetail.EntryID = dbo.RMID_MLS_Details.RMID_RefID INNER JOIN
         dbo.MaterialLocationwiseStatus ON dbo.RMID_MLS_Details.MLS_RefID = dbo.MaterialLocationwiseStatus.EntryID INNER JOIN
         dbo.VendRcvdDetailPO ON dbo.MaterialLocationwiseStatus.Rcvd_RefID = dbo.VendRcvdDetailPO.EntryID INNER JOIN
         dbo.vendRcvdDetail ON dbo.VendIssdDetail.EntryID = dbo.vendRcvdDetail.Issue_RefID
GO

CREATE OR ALTER VIEW [dbo].[VItems_StockReport_New]
AS
SELECT     dbo.VItems.CatID, dbo.VItems.ItemID, dbo.VItems.ItemName, dbo.VItems.Unit, dbo.VItems.Type, dbo.VItems.ItemSize, dbo.VItems.SizeUnit, dbo.VItems.ItemUsage, dbo.VItems.FinQuality, dbo.VItems.ItemGroup, dbo.VItems.InActive, TSF.SFStock, 
                  TForging.ForgingStock, dbo.VItems.MinLevel, dbo.VItems.MaxLevel, dbo.VItems.ReOrderLevel, TFinished.InHand, TRunning.InProductionQty, dbo.VStoreShelfs.ShelfNo, TFinished.Shelf_RefID, dbo.VStoreShelfs.RackNo, dbo.VStoreShelfs.StoreName, 
                  dbo.VItems.PriceForCost, dbo.VItems.FillingPrice, dbo.VItems.GroupID
FROM        dbo.VStoreShelfs RIGHT OUTER JOIN
                      (SELECT     ItemID, SUM(NetQty) AS InHand, MAX(Shelf_RefID) AS Shelf_RefID
                       FROM        dbo.VItemsWithShelfWiseStock_Simple
                       GROUP BY ItemID) AS TFinished ON dbo.VStoreShelfs.EntryID = TFinished.Shelf_RefID RIGHT OUTER JOIN
                  dbo.VItems LEFT OUTER JOIN
                      (SELECT     ItemID, SUM(Qty) AS SFStock
                       FROM        dbo.VItemSFStock
                       GROUP BY ItemID) AS TSF ON dbo.VItems.ItemID = TSF.ItemID LEFT OUTER JOIN
                      (SELECT     dbo.ItemsRMComp.ItemID, SUM(dbo.VRM.QtyInStock) AS ForgingStock
                       FROM        dbo.ItemsRMComp INNER JOIN
                                         dbo.VRM ON dbo.VRM.RMID = dbo.ItemsRMComp.RMID
                       GROUP BY dbo.ItemsRMComp.ItemID) AS TForging ON dbo.VItems.ItemID = TForging.ItemID ON TFinished.ItemID = dbo.VItems.ItemID LEFT OUTER JOIN
                      (SELECT     dbo.VRunningLots_Simple_PHF.ItemCode, SUM(dbo.VRunningLots_Simple_PHF.Qty) AS InProductionQty
                       FROM        dbo.VRunningLots_Simple_PHF INNER JOIN
                                         dbo.FCustomerOrders ON dbo.VRunningLots_Simple_PHF.OrderNo = dbo.FCustomerOrders.OrderNo
                       WHERE     (dbo.FCustomerOrders.CustCode IN ('Stock'))
                       GROUP BY dbo.VRunningLots_Simple_PHF.ItemCode) AS TRunning ON dbo.VItems.ItemID = TRunning.ItemCode

WHERE TFinished.InHand>0
GO

CREATE OR ALTER VIEW [dbo].[VEmpProductivityReport_Production_New]
AS
SELECT dbo.Employees.empid, dbo.Employees.name, dbo.vendRcvdDetail.ItemCode, dbo.Items.ItemName, VEmpTimes2.DT, dbo.Processes.Description, dbo.VendReceived.ProcessID, ROUND(dbo.vendRcvdDetail.RcvdQty / TCount.EmpCount, 0) AS RcvdQty, 0 AS Type, CAST(CONVERT(VARCHAR(50), dbo.VendReceived.DT, 6) AS DATETIME) AS DTOnly, dbo.vendRcvdDetail.LotNo, 
         dbo.Departments.name AS Dname, TEmpTarget.Qty, dbo.VendReceived.OverTime, dbo.Employees.deptid, TEmpTargetOT.OTQty, TRegularGroupTargets.Qty AS GroupTargetQtyRegular, TOTGroupTargets.Qty AS GroupTargetQtyOverTime, TOTGroupTargets.GroupAvgOT, TRegularGroupTargets.GroupAvgRegular,
		 VEmpTimes2.FirstInTime,VEmpTimes2.SecondOutTime,ItemGroups.Description AS ItemGroupName
FROM VEmpTimes2 LEFT JOIN 
         dbo.VendReceived_Employees ON VEmpTimes2.empid=VendReceived_Employees.EmpID 
		 --dbo.VendReceived.EntryID = dbo.VendReceived_Employees.VR_RefID  LEFT JOIN--INNER JOIN
		 LEFT JOIN dbo.VendReceived ON dbo.VendReceived.EntryID = dbo.VendReceived_Employees.VR_RefID --AND VEmpTimes2.DT=VendReceived.DT  
		 LEFT JOIN --INNER JOIN

         dbo.Employees ON VEmpTimes2.EmpID = dbo.Employees.empid LEFT JOIN --INNER JOIN
         dbo.vendRcvdDetail ON dbo.VendReceived.EntryID = dbo.vendRcvdDetail.RefID LEFT JOIN --INNER JOIN
         dbo.Items ON dbo.vendRcvdDetail.ItemCode = dbo.Items.ItemID LEFT JOIN --INNER JOIN
         dbo.ItemGroups ON dbo.Items.GroupID = dbo.ItemGroups.ID LEFT JOIN --INNER JOIN
         dbo.Processes ON dbo.VendReceived.ProcessID = dbo.Processes.ProcessID LEFT JOIN --INNER JOIN
         dbo.Departments ON dbo.Employees.deptid = dbo.Departments.deptid LEFT OUTER JOIN
             (SELECT EmpID, ProcessID, GroupID, Qty
            FROM  dbo.EmpDailyTargets
            WHERE (OverTime = 0)) AS TEmpTarget 
			ON dbo.Processes.ProcessID = TEmpTarget.ProcessID AND dbo.Items.GroupID = TEmpTarget.GroupID AND dbo.Employees.empid = TEmpTarget.EmpID LEFT OUTER JOIN

             (SELECT EmpID, ProcessID, GroupID, Qty AS OTQty
            FROM  dbo.EmpDailyTargets AS EmpDailyTargets_1
            WHERE (OverTime = 1)) AS TEmpTargetOT ON dbo.Processes.ProcessID = TEmpTargetOT.ProcessID AND dbo.Items.GroupID = TEmpTargetOT.GroupID AND dbo.Employees.empid = TEmpTargetOT.EmpID LEFT OUTER JOIN
             (SELECT VR_RefID, COUNT(VR_RefID) AS EmpCount
            FROM  dbo.VendReceived_Employees AS VendReceived_Employees_1
            GROUP BY VR_RefID) AS TCount ON dbo.VendReceived.EntryID = TCount.VR_RefID LEFT OUTER JOIN
             (SELECT EmpID, SUM(Qty) AS Qty, AVG(Qty) AS GroupAvgRegular
            FROM  dbo.EmpDailyTargets AS EmpDailyTargets_Group
            WHERE (OverTime = 0)
            GROUP BY EmpID) AS TRegularGroupTargets ON dbo.Employees.empid = TRegularGroupTargets.EmpID LEFT OUTER JOIN
             (SELECT EmpID, SUM(Qty) AS Qty, AVG(Qty) AS GroupAvgOT
            FROM  dbo.EmpDailyTargets AS EmpDailyTargets_GroupOT
            WHERE (OverTime = 1)
            GROUP BY EmpID) AS TOTGroupTargets ON dbo.Employees.empid = TOTGroupTargets.EmpID
--WHERE VEmpTimes2.DT=VendReceived.DT
GO

-- ====================================================================================================
-- PART 6: STORED PROCEDURES
-- ====================================================================================================
CREATE OR ALTER PROCEDURE Running_Lots_Summary_Hubwise(@DTFrom DATETIME,@DTTo DATETIME,@OrderNo VARCHAR(50)='0')

AS
BEGIN

	SET NOCOUNT ON;

    SELECT 
	COUNT(DISTINCT (CASE WHEN Hub_Name='H1' THEN ItemCode ELSE NULL END)) AS H1_Items
	,COUNT(DISTINCT (CASE WHEN Hub_Name='H2' THEN ItemCode ELSE NULL END)) AS H2_Items
	,COUNT(DISTINCT (CASE WHEN Hub_Name='H3' THEN ItemCode ELSE NULL END)) AS H3_Items
	,COUNT(DISTINCT (CASE WHEN Hub_Name='H4' THEN ItemCode ELSE NULL END)) AS H4_Items
	,COUNT(DISTINCT (CASE WHEN Hub_Name='H5' THEN ItemCode ELSE NULL END)) AS H5_Items
	,COUNT(DISTINCT (CASE WHEN Hub_Name='H6' THEN ItemCode ELSE NULL END)) AS H6_Items
	,SUM(CASE WHEN Hub_Name='H1' THEN Qty ELSE 0 END) AS H1_Qty
	,SUM(CASE WHEN Hub_Name='H2' THEN Qty ELSE 0 END) AS H2_Qty
	,SUM(CASE WHEN Hub_Name='H3' THEN Qty ELSE 0 END) AS H3_Qty
	,SUM(CASE WHEN Hub_Name='H4' THEN Qty ELSE 0 END) AS H4_Qty
	,SUM(CASE WHEN Hub_Name='H5' THEN Qty ELSE 0 END) AS H5_Qty
	,SUM(CASE WHEN Hub_Name='H6' THEN Qty ELSE 0 END) AS H6_Qty
	FROM VRunningLots
	INNER JOIN ItemProcessGroups ON VRunningLots.ItemCode=ItemProcessGroups.ItemID
	INNER JOIN ProcessGroupsProcesses ON ItemProcessGroups.PG_RefID=ProcessGroupsProcesses.Group_RefID AND VRunningLots.ProcessID=ProcessGroupsProcesses.Process_RefID
	WHERE VRunningLots.DT BETWEEN @DTFrom AND @DTTo
	AND (VRunningLots.OrderNo=@OrderNo OR @OrderNo IN('0',''))
	GROUP BY Hub_Name

END
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[User_Activity_Report_SP] (@UserName VARCHAR(50),@DTFrom DATETIME,@DTTO DATETIME)

AS
BEGIN

	SET @DTFrom=CAST(CONVERT(VARCHAR(50),@DTFrom,6) AS DATETIME)
	SET @DTTo=CAST(CONVERT(VARCHAR(50),@DTTo,6) AS DATETIME)
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

	SELECT T1.UserName,T1.DT,T1.ItemCode,T1.Qty,T1.ProcessID,T1.EntryType,Items.ItemName,
		   Processes.Description AS Process_Description,DTEntry,MachineName
	 FROM (
		SELECT VI.UserName,VI.DT,VID.ItemCode,VID.IssQty AS Qty,VI.ProcessID,
		CASE WHEN MasterPONo IS NULL THEN 'Process Issuance' ELSE 'Purcahse Order' END AS EntryType ,
		VI.DTEntry,VI.MachineName
		FROM VendIssued VI
		INNER JOIN VendIssdDetail VID ON VI.EntryID=VID.RefID 	
		INNER JOIN Makers ON VI.VendID=Makers.VendID
		WHERE VI.DT BETWEEN @DTFrom AND @DTTo
		AND VI.UserName=@UserName

	UNION ALL


	SELECT VR.UserName,VR.DT,VRD.ItemCode,VRD.RcvdQty,VRD.ProcessID,CASE WHEN VI.MasterPONo IS NULL THEN 'Process Receiving' ELSE 'Receiving' END AS EntryType,
		 VR.EntryDT,VR.MachineName
	FROM VendReceived VR 
		INNER JOIN VendRcvdDetail VRD ON VR.EntryID=VRD.RefID
		INNER JOIN VendIssued VI ON VR.Issuance_RefID=VI.EntryID
		WHERE VR.DT BETWEEN @DTFrom AND @DTTo
		AND VR.UserName=@UserName
		) T1

		INNER JOIN Items ON T1.ItemCode=Items.ItemID
		INNER JOIN Processes ON T1.ProcessID=Processes.ProcessID
		
END


SELECT * FROM VendIssued
GO

CREATE OR ALTER PROCEDURE [dbo].[PrintItemLabelsCustomerWise_SP] (@CustCode VARCHAR(50),@ItemID AS VARCHAR(255),@No AS INT,@PONo AS VARCHAR(255),@Qty AS INT) AS

BEGIN

	DECLARE @TempTB TABLE(AutoEntryID INT IDENTITY (1, 1),ItemNo INT,ItemID VARCHAR(255))

	DECLARE @ItemNo AS INT


	DECLARE @iCounter AS INT
	SET @iCounter=1
	WHILE @iCounter<=@No
	BEGIN
		INSERT INTO @TempTB(ItemNo,ItemID) VALUES(@iCounter,@ItemID)
		SET @iCounter=@iCounter+1
	END
			

END

SELECT TempTB.*,ItemName,@PONo AS PONo,@Qty AS Qty
	,FCustomerCatalog.ItemID AS Customer_ItemID	
	FROM @TempTB TempTB
	INNER JOIN Items ON Items.ItemID=TempTB.ItemID
	INNER JOIN FCustomerCatalog ON Items.ItemID=FCustomerCatalog.CompItemID AND FCustomerCatalog.CustCode=@CustCode
GO

CREATE OR ALTER PROCEDURE [dbo].[VendorLoanBalanceReport_SP] (@DT AS DATETIME,@InActive AS BIT) AS
SET @DT=CAST(CONVERT(VARCHAR(50),@DT,6) AS DATETIME)

SELECT     TotalAmountTaken-TotalAmountCleared AS LongTermBalance, VVendersAndMakers.SubAccOf, VVendersAndMakers.AccTitle_Parent, VVendersAndMakers.AccNo, VVendersAndMakers.AccTitle,VVendersAndMakers.MakerNo,
                      Amount-AmountCleared AS STLoan,ActualDeductionAmount AS LTLoanDed,Active
FROM         VVendersAndMakers
		LEFT OUTER JOIN VMakerAdvancesShortDeduction ON VVendersANDMakers.AccNo=VMakerAdvancesShortDeduction.AccNo
		LEFT OUTER JOIN VMakerAdvancesDeduction ON VVendersAndMakers.AccNo=VMakerAdvancesDeduction.AccNo
	WHERE (Active=0 OR @InActive=0)
GO

CREATE OR ALTER PROCEDURE [dbo].[ProductionOrderReportPHF_SP_Old] (@OrderNo AS VARCHAR(255)) AS

BEGIN
	SET ANSI_NULLS ON
	
	SET ARITHABORT ON
	

	SELECT     dbo.VrptOrders.OrderNo, dbo.VrptOrders.DT, dbo.VrptOrders.TradeTerms, dbo.VrptOrders.CustCode, dbo.VrptOrders.Country, dbo.VrptOrders.PartialShipment, 
                      dbo.VrptOrders.PaymentTerms, dbo.VrptOrders.TransShipment, dbo.VrptOrders.Packaging, dbo.VrptOrders.DeliveryDT, dbo.VrptOrders.BatchNo, 
                      dbo.VrptOrders.CompanyRefID, dbo.VrptOrders.StampDT, dbo.VrptOrders.Quality, dbo.VrptOrders.InternalRefNo, dbo.VrptOrders.OrderRcvdVia, dbo.VrptOrders.ID, 
                      dbo.VrptOrders.ItemCode, dbo.VrptOrders.Price, dbo.VrptOrders.Qty, dbo.VrptOrders.InvQty, dbo.VrptOrders.ItemName, dbo.VrptOrders.Description, 
                      dbo.VrptOrders.ItemID, dbo.VrptOrders.CompItemID, dbo.VrptOrders.Curr, dbo.VrptOrders.Unit, dbo.VrptOrders.Name, dbo.VrptOrders.Address, 
                      dbo.VrptOrders.CustomPrice, dbo.VrptOrders.CustomDescription, dbo.VrptOrders.ItemSize, dbo.VrptOrders.SizeUnit, dbo.VrptOrders.SteelType, dbo.VrptOrders.Gage, 
                      dbo.VrptOrders.ItemColor, dbo.VrptOrders.BarcodeNo, dbo.VrptOrders.PackingMode, dbo.VrptOrders.MasterCartonL, dbo.VrptOrders.MasterCartonW, 
                      dbo.VrptOrders.MasterCartonH, dbo.VrptOrders.SmallCartonL, dbo.VrptOrders.SmallCartonW, dbo.VrptOrders.SmallCartonH, dbo.VrptOrders.PolyBag, 
                      dbo.VrptOrders.SpecialInstructions, dbo.VrptOrders.StampInstructions, dbo.VrptOrders.PackingInstructions, dbo.VrptOrders.UnitItems, dbo.VrptOrders.CustomerColor, 
                      dbo.VrptOrders.ShippedQty, dbo.VrptOrders.ItemNameUrdu, dbo.VrptOrders.InHand, dbo.VrptOrders.SortNo, dbo.VrptOrders.MaxLotSize, dbo.VrptOrders.FinQuality
                      ,dbo.VrptOrders.Type, TabPO.RecieptID, TabPO.IssQty, TabAdj.AdjQty
                      ,TPTC.PTC
                      ,TProcesses.Processes
                      ,TPOQty.POQty
                      ,TPONo.PONos
                      ,TabPacked.PackQty
                      ,TCartons.Cartons
                      ,TabQC.RcvdQtyQC,TabQC.WastageQC,TabQC.ReWorkQtyQC,TabInsp.RcvdQtyInsp,TabInsp.WastageInsp,TabInsp.ReWorkQtyInsp,ReWorkQtyQCBalance,TotalQtyRcvd,TotalTransferred,ReWorkQtyInspBalance
                      ,TIMR.InitialMakerRcvings
                      ,DeliveryDTItem,DeliveryStatus,StampsItem,QualityItem,Remarks
					  ,TPOMN.POMakerName
					  ,VenderName,ReturnDT,GroupDescription,GroupID,CompanyName,FinishedWeight,SmallBoxPcs,MasterCartonSmallBoxes
					  --,dbo.VrptOrders.EmpID,dbo.VrptOrders.EmpName
					  ,VItemSFStock.Qty AS SFQty
					  ,VrptOrders.Item_Edited
					  ,VrptOrders.OrderRevisionNo
					  ,VrptOrders.ShippingMode
					  ,VrptOrders.Item_Finishing_Type_Text
FROM         dbo.VrptOrders WITH (NOLOCK) 
LEFT OUTER JOIN VItemSFStock ON dbo.VrptOrders.CompItemCode=VItemSFStock.ItemID
LEFT OUTER JOIN Companies WITH (NOLOCK)ON VrptOrders.CompanyRefID=Companies.EntryID LEFT OUTER JOIN
                          (SELECT     OrderNo, ItemCode, RecieptID, IssQty,VenderName,ReturnDT
                            FROM          dbo.VVendIssdDetail WITH (NOLOCK) INNER JOIN
                            (SELECT     MIN(EntryID) AS MINEntryID FROM dbo.VVendIssdDetail WITH (NOLOCK) WHERE OrderNo=@OrderNo AND VendID<>653 GROUP BY OrderNo, ItemCode) TI1 ON dbo.VVendIssdDetail.EntryID=TI1.MinEntryID) TabPO ON dbo.VrptOrders.OrderNo = TabPO.OrderNo AND 
                      dbo.VrptOrders.CompItemID = TabPO.ItemCode LEFT OUTER JOIN
                          (SELECT     OrderNo, ItemID, SUM(Qty) AS AdjQty
                            FROM          StockOrderAdjustments WITH (NOLOCK) WHERE OrderNo=@OrderNo
                            GROUP BY OrderNo, ItemID) TabAdj ON dbo.VrptOrders.OrderNo = TabAdj.OrderNo AND dbo.VrptOrders.CompItemID = TabAdj.ItemID LEFT OUTER JOIN
                          (SELECT     OrderNo, ItemCode, SUM(Qty) AS PackQty
                            FROM          VDispatchListDetail  WITH (NOLOCK) WHERE OrderNo=@OrderNo
                            GROUP BY OrderNo, ItemCode) TabPacked ON dbo.VrptOrders.OrderNo = TabPacked.OrderNo AND 
                      dbo.VrptOrders.CompItemID = TabPacked.ItemCode LEFT OUTER JOIN
                          (SELECT     OrderNo, ItemCode, SUM(RcvdQty) AS RcvdQtyQC, SUM(Wastage + LostQty) AS WastageQC, SUM(VendRcvdDetailReWorkDetail.Qty) AS ReWorkQtyQC, 
                                                   SUM(VendRcvdDetailReWorkDetail.Qty - VendRcvdDetailReWorkDetail.IssQty) AS ReWorkQtyQCBalance
                            FROM          VendRcvdDetail  WITH (NOLOCK) LEFT OUTER JOIN
                                                   VendRcvdDetailReWorkDetail  WITH (NOLOCK) ON VendRcvdDetail.EntryID = VendRcvdDetailReWorkDetail.VRD_RefID
                            WHERE      ProcessID = 69 AND ReqAuth = 0 AND ReWorkLot=0 AND OrderNo=@OrderNo
                            GROUP BY OrderNo, ItemCode) TabQC ON dbo.VrptOrders.OrderNo = TabQC.OrderNo AND dbo.VrptOrders.CompItemID = TabQC.ItemCode LEFT OUTER JOIN
                          (SELECT     OrderNo, ItemCode, SUM(RcvdQty) AS RcvdQtyInsp, SUM(Wastage + LostQty) AS WastageInsp, SUM(VendRcvdDetailReWorkDetail.Qty) AS ReWorkQtyInsp,
			SUM(VendRcvdDetailReWorkDetail.Qty - VendRcvdDetailReWorkDetail.IssQty) AS ReWorkQtyInspBalance
                            FROM          VendRcvdDetail  WITH (NOLOCK) LEFT OUTER JOIN
                                                   VendRcvdDetailReWorkDetail  WITH (NOLOCK) ON VendRcvdDetail.EntryID = VendRcvdDetailReWorkDetail.VRD_RefID
                            WHERE      ProcessID = 65 AND ReqAuth = 0 AND ReWorkLot=0 AND OrderNo=@OrderNo
                            GROUP BY OrderNo, ItemCode) TabInsp ON dbo.VrptOrders.OrderNo = TabInsp.OrderNo AND dbo.VrptOrders.CompItemID = TabInsp.ItemCode LEFT OUTER JOIN
		(SELECT VendRcvdDetail.OrderNo,VendRcvdDetail.ItemCode,SUM(VendRcvdDetail.RcvdQty) AS TotalQtyRcvd FROM VendRcvdDetail  WITH (NOLOCK) LEFT OUTER JOIN VendIssdDetail  WITH (NOLOCK) ON VendRcvdDetail.Issue_RefID=VendIssdDetail.EntryID 
			WHERE ISNULL(VendIssdDetail.LotNo,'0')='0' AND VendRcvdDetail.OrderNo=@OrderNo--VendRcvdDetail.EntryID IN(SELECT MIN(EntryID) FROM VendRcvdDetail WHERE ReWorkLot=0)  
			GROUP BY VendRcvdDetail.OrderNo,VendRcvdDetail.ItemCode) TTotalRcvd ON dbo.VrptOrders.OrderNo = TTotalRcvd.OrderNo AND dbo.VrptOrders.CompItemID = TTotalRcvd.ItemCode LEFT OUTER JOIN
		(SELECT OrderNo,ItemCode,SUM(Qty) AS TotalTransferred FROM TransferredToReadyFinishLotsDetail  WITH (NOLOCK) INNER JOIN VendRcvdDetail  WITH (NOLOCK) ON TransferredToReadyFinishLotsDetail.VRD_RefID=VendRcvdDetail.EntryID WHERE OrderNo=@OrderNo
			GROUP BY OrderNo,ItemCode) TTTRF ON dbo.VrptOrders.OrderNo = TTTRF.OrderNo AND dbo.VrptOrders.CompItemID = TTTRF.ItemCode 
			
			
		LEFT OUTER JOIN (SELECT DISTINCT OrderNo,ItemCode,
(STUFF((SELECT ','+LotNo + ' ('+CAST(RcvdQty AS VARCHAR(5)) +')' FROM VendRcvdDetail WITH (NOLOCK) 
	INNER JOIN (SELECT MAX(EntryID) AS MAXEntryID FROM VendRcvdDetail WITH (NOLOCK) WHERE OrderNo=@OrderNo GROUP BY LotNo) T1 ON VendRcvdDetail.EntryID=T1.MaxEntryID
	WHERE OrderNo=@OrderNo AND VendRcvdDetail.OrderNo=FOrderItems.OrderNo AND VendRcvdDetail.ItemCode=FOrderItems.CompItemCode  ORDER BY LotNo
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS PTC
	FROM FOrderItems WHERE OrderNo=@OrderNo) TPTC ON VrptOrders.OrderNo=TPTC.OrderNo AND VrptOrders.ItemID=TPTC.ItemCode
		

LEFT OUTER JOIN (
SELECT DISTINCT OrderNo,ItemCode,
(STUFF((SELECT CHAR(13)+Description + + ' (' + CONVERT(VARCHAR(50),DT,5) + ') ('  + CAST(RcvdQty-ISNULL(Wastage,0)-ISNULL(ReWorkQty,0) AS VARCHAR(10)) + ')' FROM VendRcvdDetail WITH (NOLOCK) INNER JOIN VendReceived WITH (NOLOCK) ON VendReceived.EntryID=VendRcvdDetail.RefID 
	INNER JOIN Processes ON VendReceived.ProcessID=Processes.ProcessID
	INNER JOIN (SELECT MAX(EntryID) AS MAXEntryID FROM VendRcvdDetail WITH (NOLOCK) WHERE OrderNo=@OrderNo AND VendRcvdDetail.OrderNo=FOrderItems.OrderNo AND VendRcvdDetail.ItemCode=FOrderItems.CompItemCode GROUP BY LotNo) T1 ON VendRcvdDetail.EntryID=T1.MAXEntryID
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS Processes
	FROM FOrderItems WHERE OrderNo=@OrderNo) TProcesses ON VrptOrders.OrderNo=TProcesses.OrderNo AND VrptOrders.ItemID=TProcesses.ItemCode

LEFT OUTER JOIN (
SELECT DISTINCT OrderNo,ItemCode,	
(STUFF((SELECT CHAR(13) + CHAR(10)+CAST(IssQty AS VARCHAR(5)) FROM VVendIssdDetail_Simple WITH (NOLOCK) INNER JOIN 
	(SELECT MIN(EntryID) AS MINEntryID FROM dbo.VVendIssdDetail_Simple WITH (NOLOCK) WHERE OrderNo=@OrderNo AND MasterPONo<>''
	AND VVendIssdDetail_Simple.OrderNo=FOrderItems.OrderNo AND VVendIssdDetail_Simple.ItemCode=FOrderItems.CompItemCode GROUP BY OrderNo,ItemCode,ProcessID) T1 ON VVendIssdDetail_Simple.EntryID=T1.MINEntryID
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS POQty
	FROM FOrderItems WHERE OrderNo=@OrderNo) TPOQty ON VrptOrders.OrderNo=TPOQty.OrderNo AND VrptOrders.ItemID=TPOQty.ItemCode
	
	
	
LEFT OUTER JOIN (
SELECT DISTINCT OrderNo,ItemCode,	
(STUFF((SELECT CHAR(13) + CHAR(10)+RecieptID FROM VVendIssdDetail_Simple WITH (NOLOCK) INNER JOIN 
	(SELECT MIN(EntryID) AS MINEntryID FROM dbo.VVendIssdDetail_Simple WITH (NOLOCK) WHERE OrderNo=@OrderNo AND ProcessID IN(172,212) 
	AND VVendIssdDetail_Simple.OrderNo=FOrderItems.OrderNo AND VVendIssdDetail_Simple.ItemCode=FOrderItems.CompItemCode GROUP BY OrderNo,ItemCode,ProcessID) T1 ON VVendIssdDetail_Simple.EntryID=T1.MINEntryID
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS PONos
	FROM FOrderItems WHERE OrderNo=@OrderNo) TPONo ON VrptOrders.OrderNo=TPONo.OrderNo AND VrptOrders.ItemID=TPONo.ItemCode
	
LEFT OUTER JOIN (
SELECT DISTINCT OrderNo,ItemCode,
(STUFF((SELECT CHAR(13)+CHAR(10)+CAST(CartonNo AS VARCHAR(10)) + '(' + CAST(SUM(Qty) AS VARCHAR(10)) + ')' FROM VDispatchListDetail_Adv WITH (NOLOCK) WHERE OrderNo=@OrderNo
	AND VDispatchListDetail_Adv.OrderNo=FOrderItems.OrderNo AND VDispatchListDetail_Adv.ItemCode=FOrderItems.CompItemCode GROUP BY OrderNo,ItemCode,CartonNo
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS Cartons
	FROM FOrderItems WHERE OrderNo=@OrderNo) TCartons ON VrptOrders.OrderNo=TCartons.OrderNo AND VrptOrders.ItemID=TCartons.ItemCode

LEFT OUTER JOIN (
SELECT DISTINCT OrderNo,ItemCode,
(STUFF((SELECT CHAR(13)+VendRcvdDetail.LotNo+ ' (' + CAST(SUM(VendRcvdDetail.RcvdQty) AS VARCHAR(10)) + ')' FROM VendRcvdDetail WITH (NOLOCK) 
	INNER JOIN VendIssdDetail ON VendRcvdDetail.Issue_RefID=VendIssdDetail.EntryID WHERE VendIssdDetail.LotNo='0' AND VendRcvdDetail.ReWorkLot=0 AND VendRcvdDetail.EntryID NOT IN(SELECT EntryID FROM LotTransferDetails)
		AND VendRcvdDetail.OrderNo=@OrderNo	
		AND VendRcvdDetail.OrderNo=FOrderItems.OrderNo AND VendRcvdDetail.ItemCode=ForderItems.CompItemCode
		GROUP BY VendRcvdDetail.LotNo
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS InitialMakerRcvings
	FROM FOrderItems WHERE OrderNo=@OrderNo) TIMR ON VrptOrders.OrderNo=TIMR.OrderNo AND VrptOrders.ItemID=TIMR.ItemCode
	
LEFT OUTER JOIN (SELECT DISTINCT OrderNo,ItemCode,
(STUFF((SELECT CHAR(13)+CHAR(10)+VenderName FROM VendIssdDetail WITH (NOLOCK) INNER JOIN VendIssued WITH (NOLOCK) ON VendIssued.EntryID=VendIssdDetail.RefID INNER JOIN Makers WITH (NOLOCK) ON Makers.VendID=VendIssued.VendID INNER JOIN
	(SELECT MIN(EntryID) MINEntryID FROM dbo.VVendIssdDetail_Simple WITH (NOLOCK) WHERE MasterPONo<>'' AND VVendIssdDetail_Simple.OrderNo=@OrderNo AND VVendIssdDetail_Simple.OrderNo=FOrderItems.OrderNo AND VVendIssdDetail_Simple.ItemCode=ForderItems.CompItemCode GROUP BY OrderNo,ItemCode,ProcessID) T1 ON VendIssdDetail.EntryID=T1.MinEntryID
	FOR XML PATH(''), TYPE, ROOT).value('root[1]','nvarchar(max)'),1,1,'')) AS POMakerName
	FROM FOrderItems WHERE OrderNo=@OrderNo) TPOMN ON VrptOrders.OrderNo=TPOMN.OrderNo AND VrptOrders.ItemID=TPOMN.ItemCode
	WHERE VrptOrders.OrderNo=@OrderNo

END
GO

CREATE OR ALTER PROCEDURE 
	[dbo].[InsertChqBookData_SP]
    @StartingFrom INT,
    @Chqs INT,
    @ManualNo NVARCHAR(50),
    @ChqBookDetail NVARCHAR(200),
    @AccNo NVARCHAR(50)	
AS
BEGIN

	SET NOCOUNT ON;
	
	INSERT INTO ChqBooks(AccNo,StartingFrom,Chqs,ManualNo,ChqBookDetail)
	VALUES 
	(@AccNo,@StartingFrom,@Chqs,@ManualNo,@ChqBookDetail)

	SELECT MAX(ChqBookNo) FROM ChqBooks;
END
GO

CREATE OR ALTER PROCEDURE [dbo].[InsertChqBookDetailData_SP] (@StartingFrom VARCHAR(50),@Chqs INT,@ChqBookNo AS INT)
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @iSrNo INT,@ChqLen INT,@EndNo INT

    --SET @iSrNo = CAST(@StartingFrom AS INT);
    SET @iSrNo = 0;
    SET @ChqLen = LEN(@StartingFrom);
    SET @EndNo = @StartingFrom + @Chqs - 1;

    --CREATE TABLE #Temp(ChqBookNo VARCHAR(50), ChqNo VARCHAR(50),Issued tinyint)
    WHILE @iSrNo <= (@Chqs-1)
    BEGIN
        INSERT INTO ChqList (ChqBookNo, ChqNo,Issued)
        --INSERT INTO #Temp(ChqBookNo, ChqNo,Issued)
        VALUES (@ChqBookNo,FORMAT(@StartingFrom + @iSrNo, REPLICATE('0', @ChqLen)),0);

        SET @iSrNo = @iSrNo + 1;
    END

    SELECT @EndNo AS LastChequeNo;
    --SELECT * FROM #Temp
    --DROP TABLE #Temp;

END;
GO

CREATE OR ALTER PROCEDURE [dbo].[PTC_SP_2_9_2023] (@LotNo AS VARCHAR(50)) AS

BEGIN
SET ANSI_WARNINGS OFF

DECLARE @RepairGeneratedFrom AS VARCHAR(4000),@RepairType AS VARCHAR(255)


--IF @Repair=0
	BEGIN
	SET @RepairGeneratedFrom=''
	SET @RepairType=''
SELECT     TFOrderItems.OrderNo, TFOrderItems.ItemCode, TFOrderItems.Qty, dbo.FCustomerOrders.CustCode, dbo.FCustomerOrders.Country, dbo.FCustomerOrders.DT, 
                      dbo.Items.ItemName, dbo.Items.ItemSize, dbo.Items.SizeUnit, dbo.Items.Type, dbo.FCustomerOrders.Packaging
					  ,CASE WHEN DeliveryDTItem='1/1/1900' THEN dbo.FCustomerOrders.DeliveryDT ELSE DeliveryDTItem END AS DeliveryDT
                      ,dbo.ItemProcesses.SNO, dbo.ItemProcesses.ProcessID, dbo.Processes.Description, Tab1.RcvdQty, Tab2.LotNo, Tab1.VendID, dbo.Makers.VenderName, Tab1.RcvDT, 
                      TabIss.IssDT, TabMin.MinProcessID, Tab1.Wastage, Tab1.PassQty, dbo.Employees.empid, dbo.Employees.name, Tab1.ReWorkQty, TFOrderItems.CompItemCode, 
                      dbo.Processes.Code, dbo.Processes.ProcessNameUrdu, dbo.Items.ItemNameUrdu, EmpRcv.name AS RcvEmpName, Tab1.Rate, TMasterPO.MasterPONo
                      ,TSteelBatch.SteelBatchNo
					  --,TForgeBatch.ForgeBatchNo
					  ,TSteelBatch.ForgeBatchNo
					  , TSFRemarks.SFRemarks, TSplit.SplitQty, ISNULL(TSplit.Type, TSplitFrom.SplitTypeFrom) AS SplitType, 
                      TSplit.ToOrderNo, TSplit.ToLotNo, TSplitFrom.FromOrderNo, TSplitFrom.FromLotNo, TSplitFrom.SplitQtyFrom, TSplit.FromItemCode, TSplit.ToItemCode, 
                      TMasterPO.RecieptID, TIssFromSF.LotNo AS IssSFLotNo, TIssFromSF.IssSFRemarks, TIssFromSF.IssSFQty, TLastRcv.LotLastEntryID, 
                      dbo.VFOrderItemsWithShippedQtyItemCodeWise.ShippedQty, ISNULL(dbo.VLotWithPolishingMaker.VenderName, '') AS FromLotPolisher, 
                      ISNULL(dbo.VLotWithFirstProcessOfIssuance.VenderName, '') AS FromLotOriginator, Tab1.RcvTemperValue, TabIss.IssSpecialInstructions, 
                      dbo.FCustomerOrders.InternalRefNo, dbo.ProcessGroups.PGReference, Tab1.ReceivedEntryID, dbo.VendReceived_Employees_F(Tab1.ReceivedEntryID) 
                      AS Employees, TFOrderItems.Quality, dbo.GetAllRepairLots_F(Tab2.LotNo) AS RepairLots, TSplit.LotTransferRemarks, 
                      TSplitFrom.LotTransferRemarks AS LotTransferRemarksFrom, Tab1.UserName, Tab1.MachineName,ProcessGroupsProcesses.Hub_Name,VI_DT,VLotsWithMasterPONo.Batch_No AS VID_Batch_No,VLotsWithMasterPONo.MasterPODT
					  ,TIssFromSF.StoreName,TIssFromSF.Location
					  ,Lots_List.Batch_No AS Lots_List_Batch_No,Lots_List.Reference_LotNo,Lots_List.Mill_Certificate_No,Mill_Certificate_No_From_RM
					  ,Items.HRC_From,Items.HRC_To
					  ,Lots_List.Forge_Batch_No AS Lots_List_Forge_Batch_No
					  ,ISNULL(TFCC.FinQuality,'') AS FinQuality,ISNULL(TFCC.SpecialInstructions,'') AS SpecialInstructions_FCC,
					  ISNULL(TFCC.StampInstructions,'') AS StampInstructions,ISNULL(TFCC.PackingInstructions,'') AS PackingInstructions,
					  TFCC.MDMA,TFCC.SFDA_Listing_No,TFCC.MD_Group
					  ,dbo.Items.ItemPic,TFOrderItems.FOI_Stamps,ISNULL(FCustomerOrders.Packaging,'') AS SpecialInstruction,
					  TFOrderItems.BatchNo,ISNULL(Makers.Maker_Second_Name,'') AS Maker_Second_Name

FROM         (SELECT     OrderNo, ItemCode, CompItemCode, Quality, SUM(Qty) AS Qty,DeliveryDTItem,FOI_Stamps,IW_BatchNo AS BatchNo
                       FROM          dbo.VrptOrders_ForProduction_Simple 
                       GROUP BY OrderNo, ItemCode, CompItemCode, Quality,DeliveryDTItem,FOI_Stamps,IW_BatchNo) AS TFOrderItems LEFT OUTER JOIN
                      dbo.VFOrderItemsWithShippedQtyItemCodeWise ON TFOrderItems.OrderNo = dbo.VFOrderItemsWithShippedQtyItemCodeWise.OrderNo AND 
                      TFOrderItems.CompItemCode = dbo.VFOrderItemsWithShippedQtyItemCodeWise.CompItemCode INNER JOIN
                      dbo.Items ON TFOrderItems.CompItemCode = dbo.Items.ItemID INNER JOIN
                      dbo.ItemProcesses ON dbo.Items.ItemID = dbo.ItemProcesses.ItemID INNER JOIN
                      dbo.FCustomerOrders ON TFOrderItems.OrderNo = dbo.FCustomerOrders.OrderNo INNER JOIN
                      dbo.Processes ON dbo.ItemProcesses.ProcessID = dbo.Processes.ProcessID LEFT OUTER JOIN
                      dbo.ItemProcessGroups ON dbo.Items.ItemID = dbo.ItemProcessGroups.ItemID LEFT OUTER JOIN
                      dbo.ProcessGroups ON dbo.ItemProcessGroups.PG_RefID = dbo.ProcessGroups.EntryID LEFT OUTER JOIN
					  ProcessGroupsProcesses ON dbo.ProcessGroups.EntryID=ProcessGroupsProcesses.Group_RefID AND ItemProcesses.ProcessID=ProcessGroupsProcesses.Process_RefID
					  LEFT OUTER JOIN
                          (SELECT DISTINCT OrderNo, ItemCode, LotNo
                             FROM dbo.VendRcvdDetail WHERE LotNo=@LotNo) AS Tab2 ON TFOrderItems.OrderNo = Tab2.OrderNo AND TFOrderItems.CompItemCode = Tab2.ItemCode LEFT OUTER JOIN
                          (SELECT     VendRcvdDetail_2.ProcessID,VendRcvdDetail_2.LotNo,VendRcvdDetail_2.Rate,dbo.VendReceived.EmpID, 
                                                   dbo.VendReceived.TemperValue AS RcvTemperValue,VendRcvdDetail_2.RcvdQty,CASE WHEN VendRcvdDetail_2.ReqAuth = 1 THEN 0 ELSE VendRcvdDetail_2.RcvdQty END AS PassQty, 
                                                   dbo.VendReceived.VendID,dbo.VendReceived.DT AS RcvDT,VendRcvdDetail_2.Wastage,VendRcvdDetail_2.ReWorkQty, 
                                                   dbo.VendReceived.EntryID AS ReceivedEntryID, dbo.VendReceived.UserName, dbo.VendReceived.MachineName,VendRcvdDetail_2.EntryID AS VRD_EntryID,VendIssued.DT AS VI_DT,VendIssdDetail.Batch_No AS VID_Batch_No
                             FROM dbo.VendRcvdDetail AS VendRcvdDetail_2 
							 INNER JOIN dbo.VendReceived ON dbo.VendReceived.EntryID = VendRcvdDetail_2.RefID
							 LEFT JOIN VendIssdDetail ON VendRcvdDetail_2.Issue_RefID=VendIssdDetail.EntryID
							 LEFT JOIN VendIssued ON VendIssdDetail.RefID=VendIssued.EntryID
								WHERE VendRcvdDetail_2.LotNo=@LotNo) AS Tab1 ON Tab2.LotNo = Tab1.LotNo AND 
                      dbo.ItemProcesses.ProcessID = Tab1.ProcessID LEFT OUTER JOIN
                          (SELECT     dbo.VendIssued.ProcessID, dbo.VendIssdDetail.LotNo, dbo.VendIssdDetail_MoreDetails.EmpID,
                                                   dbo.VendIssued.SpecialInstructions AS IssSpecialInstructions, MIN(dbo.VendIssued.DT) AS IssDT
                             FROM         dbo.VendIssued INNER JOIN
                                                   dbo.VendIssdDetail ON dbo.VendIssued.EntryID = dbo.VendIssdDetail.RefID LEFT OUTER JOIN
                                                   dbo.VendIssdDetail_MoreDetails ON dbo.VendIssdDetail.EntryID = dbo.VendIssdDetail_MoreDetails.VID_RefID
                             GROUP BY dbo.VendIssued.ProcessID, dbo.VendIssdDetail.LotNo, dbo.VendIssdDetail_MoreDetails.EmpID, dbo.VendIssued.SpecialInstructions) AS TabIss ON 
                      Tab2.LotNo = TabIss.LotNo AND dbo.ItemProcesses.ProcessID = TabIss.ProcessID LEFT OUTER JOIN
                          (SELECT     VendRcvdDetail_1.LotNo, MIN(VendReceived_1.ProcessID) AS MinProcessID
                             FROM         dbo.VendReceived AS VendReceived_1 INNER JOIN
                                                   dbo.VendRcvdDetail AS VendRcvdDetail_1 ON VendReceived_1.EntryID = VendRcvdDetail_1.RefID  WHERE VendRcvdDetail_1.LotNo=@LotNo
                             GROUP BY VendRcvdDetail_1.LotNo) AS TabMin ON Tab2.LotNo = TabMin.LotNo LEFT OUTER JOIN
                      dbo.Makers ON Tab1.VendID = dbo.Makers.VendID LEFT OUTER JOIN
                      dbo.Employees ON TabIss.EmpID = dbo.Employees.empid LEFT OUTER JOIN
                      dbo.Employees AS EmpRcv ON Tab1.EmpID = EmpRcv.empid LEFT OUTER JOIN
                          (SELECT     TNest1.LotNo, VendIssued_2.EntryID, VendIssued_2.MasterPONo, VendIssued_2.RecieptID
                             FROM         dbo.VendIssued AS VendIssued_2 INNER JOIN
                                                   dbo.VendIssdDetail AS VendIssdDetail_2 ON VendIssued_2.EntryID = VendIssdDetail_2.RefID INNER JOIN
                                                       (SELECT     LotNo, MIN(Issue_RefID) AS Issue_RefID
                                                          FROM         dbo.VendRcvdDetail AS VendRcvdDetail_7 WHERE LotNo=@LotNo
                                                          GROUP BY LotNo) AS TNest1 ON TNest1.Issue_RefID = VendIssdDetail_2.EntryID) AS TMasterPO ON 
                      Tab2.LotNo = TMasterPO.LotNo LEFT OUTER JOIN
                          (SELECT     dbo.RawMaterialIssuance.VI_RefID, MAX(dbo.RawMaterialIssuanceDetail.PORefNo) AS SteelBatchNo,MAX(VendRcvdDetailPO.RcvID) AS ForgeBatchNo
								,MAX(MaterialLocationwiseStatus.Mill_Certificate_No) AS Mill_Certificate_No_From_RM
                             FROM         dbo.RawMaterialIssuance INNER JOIN
                                                   dbo.RawMaterialIssuanceDetail ON dbo.RawMaterialIssuance.IssNo = dbo.RawMaterialIssuanceDetail.IssNo INNER JOIN
                                                   dbo.RM ON dbo.RawMaterialIssuanceDetail.RMID1 = dbo.RM.RMID1
												   INNER JOIN RMID_MLS_Details ON RawMaterialIssuanceDetail.EntryID=RMID_MLS_Details.RMID_RefID
												   INNER JOIN MaterialLocationwiseStatus ON RMID_MLS_Details.MLS_RefID=MaterialLocationwiseStatus.EntryID
												   INNER JOIN VendRcvdDetailPO ON MaterialLocationwiseStatus.Rcvd_RefID=VendRcvdDetailPO.EntryID
                             WHERE     (dbo.RM.GroupID = 1)
                             GROUP BY dbo.RawMaterialIssuance.VI_RefID) AS TSteelBatch ON TMasterPO.EntryID = TSteelBatch.VI_RefID LEFT OUTER JOIN
                          /*(SELECT     RawMaterialIssuance_1.VI_RefID, MAX(RawMaterialIssuanceDetail_1.PORefNo) AS ForgeBatchNo
                             FROM         dbo.RawMaterialIssuance AS RawMaterialIssuance_1 INNER JOIN
                                                   dbo.RawMaterialIssuanceDetail AS RawMaterialIssuanceDetail_1 ON RawMaterialIssuance_1.IssNo = RawMaterialIssuanceDetail_1.IssNo INNER JOIN
                                                   dbo.RM AS RM_1 ON RawMaterialIssuanceDetail_1.RMID1 = RM_1.RMID1
                             WHERE     (RM_1.GroupID IN (25, 27, 28, 29, 30))
                             GROUP BY RawMaterialIssuance_1.VI_RefID) AS TForgeBatch ON TMasterPO.EntryID = TForgeBatch.VI_RefID LEFT OUTER JOIN*/
						
                          (SELECT     VendRcvdDetail_6.ProcessID, VendRcvdDetail_6.LotNo, '' AS SFRemarks
                             FROM         dbo.VendRcvdDetail AS VendRcvdDetail_6 INNER JOIN
                                                       (SELECT     VID_RefID
                                                          FROM         dbo.StockOrderOpening_Issuance
                                                          GROUP BY VID_RefID) AS TSOO_I ON TSOO_I.VID_RefID = VendRcvdDetail_6.Issue_RefID) AS TSFRemarks ON Tab2.LotNo = TSFRemarks.LotNo AND
                       Tab1.ProcessID = TSFRemarks.ProcessID LEFT OUTER JOIN
                          (SELECT     VendRcvdDetail_5.LotNo, VendRcvdDetail_5.ProcessID, dbo.LotTransferDetails.SplitQty, dbo.LotTransferDetails.Type, 
                                                   dbo.LotTransferDetails.ToOrderNo, VendRcvdDetail_1.LotNo AS ToLotNo, dbo.LotTransferDetails.FromItemCode, dbo.LotTransferDetails.ToItemCode, 
                                                   dbo.LotTransferDetails.LotTransferRemarks
                             FROM         dbo.LotTransferDetails INNER JOIN
                                                   dbo.VendRcvdDetail AS VendRcvdDetail_5 ON dbo.LotTransferDetails.VRD_From_RefID = VendRcvdDetail_5.EntryID LEFT OUTER JOIN
                                                   dbo.VendRcvdDetail AS VendRcvdDetail_1 ON dbo.LotTransferDetails.VRD_To_RefID = VendRcvdDetail_1.EntryID) AS TSplit ON 
                      Tab2.LotNo = TSplit.LotNo AND Tab1.ProcessID = TSplit.ProcessID LEFT OUTER JOIN
                          (SELECT     VendRcvdDetail_4.LotNo AS FromLotNo, VendRcvdDetail_4.ProcessID, LotTransferDetails_1.SplitQty AS SplitQtyFrom, 
                                                   LotTransferDetails_1.Type AS SplitTypeFrom, LotTransferDetails_1.FromOrderNo, VendRcvdDetail_1.LotNo, 
                                                   LotTransferDetails_1.LotTransferRemarks
                             FROM         dbo.LotTransferDetails AS LotTransferDetails_1 INNER JOIN
                                                   dbo.VendRcvdDetail AS VendRcvdDetail_4 ON LotTransferDetails_1.VRD_From_RefID = VendRcvdDetail_4.EntryID INNER JOIN
                                                   dbo.VendRcvdDetail AS VendRcvdDetail_1 ON LotTransferDetails_1.VRD_To_RefID = VendRcvdDetail_1.EntryID) AS TSplitFrom ON 
                      Tab2.LotNo = TSplitFrom.LotNo AND Tab1.ProcessID = TSplitFrom.ProcessID LEFT OUTER JOIN
                          (SELECT     VendIssued_1.ProcessID, VendIssdDetail_1.LotNo, '' AS IssSFRemarks, VendIssdDetail_1.IssQty AS IssSFQty,TSOO_I_1.StoreName,TSOO_I_1.Location
                             FROM         dbo.VendIssdDetail AS VendIssdDetail_1 INNER JOIN
                                                       (SELECT     VID_RefID,StoreName,Location
                                                          FROM         dbo.StockOrderOpening_Issuance AS StockOrderOpening_Issuance_1
														  LEFT JOIN StockOrderOpening ON StockOrderOpening_Issuance_1.SOO_RefID=StockOrderOpening.EntryID
														  LEFT JOIN VStoreShelfs ON StockOrderOpening.Shelf_RefID=VStoreShelfs.EntryID
                                                          GROUP BY VID_RefID,StoreName,Location) AS TSOO_I_1 ON VendIssdDetail_1.EntryID = TSOO_I_1.VID_RefID INNER JOIN
                                                   dbo.VendIssued AS VendIssued_1 ON VendIssdDetail_1.RefID = VendIssued_1.EntryID) AS TIssFromSF ON Tab2.LotNo = TIssFromSF.LotNo AND 
                      Tab1.ProcessID = TIssFromSF.ProcessID LEFT OUTER JOIN
                          (SELECT     LotNo, MAX(EntryID) AS LotLastEntryID
                             FROM         dbo.VendRcvdDetail AS VendRcvdDetail_3 WHERE LotNo=@LotNo
                             GROUP BY LotNo) AS TLastRcv ON Tab2.LotNo = TLastRcv.LotNo LEFT OUTER JOIN
                      dbo.VLotWithPolishingMaker ON TSplitFrom.FromLotNo = dbo.VLotWithPolishingMaker.LotNo LEFT OUTER JOIN
                      dbo.VLotWithFirstProcessOfIssuance ON TSplitFrom.FromLotNo = dbo.VLotWithFirstProcessOfIssuance.LotNo
					  LEFT OUTER JOIN VLotsWithMasterPONo ON Tab2.LotNo=VLotsWithMasterPONo.LotNO
					  LEFT OUTER JOIN Lots_List ON Tab2.LotNo=Lots_List.LotNo
					  LEFT OUTER JOIN FCustomerCatalog TFCC ON FCustomerOrders.CustCode=TFCC.CustCode AND FCustomerOrders.Country=TFCC.Country AND TFOrderItems.CompItemCode=TFCC.CompItemID
	WHERE Tab2.LotNo=@LotNo
	
	END
/*ELSE
	BEGIN
	
	SET @RepairGeneratedFrom=dbo.RepairLot_OriginalLots_F(@LotNo)
	SELECT        TFOrderItems.OrderNo, TFOrderItems.ItemCode, TFOrderItems.Qty, dbo.FCustomerOrders.CustCode, dbo.FCustomerOrders.Country, dbo.FCustomerOrders.DT, 
                         dbo.Items.ItemName, dbo.Items.ItemSize, dbo.Items.SizeUnit, dbo.Items.Type, dbo.FCustomerOrders.Packaging, dbo.FCustomerOrders.DeliveryDT, 
                         dbo.RepairTypeProcesses.SeqNo AS SNO, dbo.RepairTypeProcesses.ProcessID, dbo.Processes.Description, Tab1.RcvdQty, Tab2.LotNo, Tab1.VendID, 
                         dbo.Makers.VenderName, Tab1.RcvDT, TabIss.IssDT, TabMin.MinProcessID, Tab1.Wastage, Tab1.PassQty, dbo.Employees.empid, dbo.Employees.name, 
                         Tab1.ReWorkQty, TFOrderItems.CompItemCode, dbo.Processes.Code, dbo.Processes.ProcessNameUrdu, dbo.Items.ItemNameUrdu, 
                         EmpRcv.name AS RcvEmpName, Tab1.Rate, TMasterPO.MasterPONo, TSteelBatch.SteelBatchNo, TForgeBatch.ForgeBatchNo, TSFRemarks.SFRemarks, 
                         TSplit.SplitQty, ISNULL(TSplit.Type, TSplitFrom.SplitTypeFrom) AS SplitType, TSplit.ToOrderNo, TSplit.ToLotNo, TSplitFrom.FromOrderNo, TSplitFrom.FromLotNo, 
                         TSplitFrom.SplitQtyFrom, @RepairGeneratedFrom AS RepairGeneratedFrom, dbo.RepairTypes.RepairType, TLastRcv.LotLastEntryID, 
                         Tab_TTRF.TTRFQty, Tab1.ReceivedEntryID, dbo.VendReceived_Employees_F(Tab1.ReceivedEntryID) AS Employees,'' AS IssEmpID,'' AS IssEmpName,dbo.GetMainLots_Repair_F(@LotNo) AS RepairLotsGenerated
FROM            (SELECT        OrderNo, ItemCode, CompItemCode, SUM(Qty) AS Qty
                           FROM            dbo.FOrderItems
                           GROUP BY OrderNo, ItemCode, CompItemCode) AS TFOrderItems INNER JOIN
                         dbo.Items ON TFOrderItems.CompItemCode = dbo.Items.ItemID INNER JOIN
                         dbo.FCustomerOrders ON TFOrderItems.OrderNo = dbo.FCustomerOrders.OrderNo LEFT OUTER JOIN
                             (SELECT DISTINCT OrderNo, ItemCode, LotNo, Repair_RefID
                                FROM            dbo.VendRcvdDetail WHERE LotNo=@LotNo) AS Tab2 ON TFOrderItems.OrderNo = Tab2.OrderNo AND TFOrderItems.CompItemCode = Tab2.ItemCode INNER JOIN
                         dbo.RepairTypeProcesses ON dbo.RepairTypeProcesses.Repair_RefID = Tab2.Repair_RefID INNER JOIN
                         dbo.Processes ON dbo.RepairTypeProcesses.ProcessID = dbo.Processes.ProcessID INNER JOIN
                         dbo.RepairTypes ON Tab2.Repair_RefID = dbo.RepairTypes.EntryID LEFT OUTER JOIN
                             (SELECT        VendRcvdDetail_2.ProcessID, VendRcvdDetail_2.LotNo, VendRcvdDetail_2.Rate, dbo.VendReceived.EmpID, VendRcvdDetail_2.RcvdQty, 
                                                         CASE WHEN ReqAuth = 1 THEN 0 ELSE RcvdQty END AS PassQty, dbo.VendReceived.VendID, dbo.VendReceived.DT AS RcvDT, 
                                                         VendRcvdDetail_2.Wastage, VendRcvdDetail_2.ReWorkQty, dbo.VendReceived.EntryID AS ReceivedEntryID
                                FROM            dbo.VendRcvdDetail AS VendRcvdDetail_2 INNER JOIN
                                                         dbo.VendReceived ON dbo.VendReceived.EntryID = VendRcvdDetail_2.RefID) AS Tab1 ON Tab2.LotNo = Tab1.LotNo AND 
                         dbo.RepairTypeProcesses.ProcessID = Tab1.ProcessID LEFT OUTER JOIN
                             (SELECT        dbo.VendIssued.ProcessID, dbo.VendIssdDetail.LotNo, dbo.VendIssdDetail_MoreDetails.EmpID, dbo.VendIssdDetail.Repair_RefID, 
                                                         MIN(dbo.VendIssued.DT) AS IssDT
                                FROM            dbo.VendIssued INNER JOIN
                                                         dbo.VendIssdDetail ON dbo.VendIssued.EntryID = dbo.VendIssdDetail.RefID LEFT OUTER JOIN
                                                         dbo.VendIssdDetail_MoreDetails ON dbo.VendIssdDetail.EntryID = dbo.VendIssdDetail_MoreDetails.VID_RefID WHERE dbo.VendIssdDetail.LotNo=@LotNo
                                GROUP BY dbo.VendIssued.ProcessID, dbo.VendIssdDetail.LotNo, dbo.VendIssdDetail_MoreDetails.EmpID, dbo.VendIssdDetail.Repair_RefID) AS TabIss ON 
                         Tab2.LotNo = TabIss.LotNo AND dbo.RepairTypeProcesses.ProcessID = TabIss.ProcessID LEFT OUTER JOIN
                             (SELECT        VendRcvdDetail_1.LotNo, MIN(VendReceived_1.ProcessID) AS MinProcessID
                                FROM            dbo.VendReceived AS VendReceived_1 INNER JOIN
                                                         dbo.VendRcvdDetail AS VendRcvdDetail_1 ON VendReceived_1.EntryID = VendRcvdDetail_1.RefID WHERE VendRcvdDetail_1.LotNo=@LotNo
                                GROUP BY VendRcvdDetail_1.LotNo) AS TabMin ON Tab2.LotNo = TabMin.LotNo LEFT OUTER JOIN
                         dbo.Makers ON Tab1.VendID = dbo.Makers.VendID LEFT OUTER JOIN
                         dbo.Employees ON TabIss.EmpID = dbo.Employees.empid LEFT OUTER JOIN
                         dbo.Employees AS EmpRcv ON Tab1.EmpID = EmpRcv.empid LEFT OUTER JOIN
                             (SELECT        TNest1.LotNo, dbo.VendIssued.EntryID, dbo.VendIssued.MasterPONo
                                FROM            dbo.VendIssued INNER JOIN
                                                         dbo.VendIssdDetail ON dbo.VendIssued.EntryID = dbo.VendIssdDetail.RefID INNER JOIN
                                                             (SELECT        LotNo, MIN(Issue_RefID) AS Issue_RefID
                                                                FROM            dbo.VendRcvdDetail WHERE LotNo=@LotNo
                                                                GROUP BY LotNo) AS TNest1 ON TNest1.Issue_RefID = dbo.VendIssdDetail.EntryID) AS TMasterPO ON 
                         Tab1.LotNo = TMasterPO.LotNo LEFT OUTER JOIN
                             (SELECT        dbo.RawMaterialIssuance.VI_RefID, MAX(dbo.RawMaterialIssuanceDetail.PORefNo) AS SteelBatchNo
                                FROM            dbo.RawMaterialIssuance INNER JOIN
                                                         dbo.RawMaterialIssuanceDetail ON dbo.RawMaterialIssuance.IssNo = dbo.RawMaterialIssuanceDetail.IssNo INNER JOIN
                                                         dbo.RM ON dbo.RawMaterialIssuanceDetail.RMID1 = dbo.RM.RMID1
                                WHERE        (dbo.RM.GroupID = 20)
                                GROUP BY dbo.RawMaterialIssuance.VI_RefID) AS TSteelBatch ON TMasterPO.EntryID = TSteelBatch.VI_RefID LEFT OUTER JOIN
                             (SELECT        dbo.RawMaterialIssuance.VI_RefID, MAX(dbo.RawMaterialIssuanceDetail.PORefNo) AS ForgeBatchNo
                                FROM            dbo.RawMaterialIssuance INNER JOIN
                                                         dbo.RawMaterialIssuanceDetail ON dbo.RawMaterialIssuance.IssNo = dbo.RawMaterialIssuanceDetail.IssNo INNER JOIN
                                                         dbo.RM ON dbo.RawMaterialIssuanceDetail.RMID1 = dbo.RM.RMID1
                                WHERE        (dbo.RM.GroupID IN (25, 27, 28, 29, 30))
                                GROUP BY dbo.RawMaterialIssuance.VI_RefID) AS TForgeBatch ON TMasterPO.EntryID = TForgeBatch.VI_RefID LEFT OUTER JOIN
                             (SELECT        dbo.VendRcvdDetail.ProcessID, dbo.VendRcvdDetail.LotNo, dbo.StockOrderOpening.Remarks AS SFRemarks
                                FROM            dbo.StockOrderOpening INNER JOIN
                                                         dbo.StockOrderOpening_Issuance ON dbo.StockOrderOpening.EntryID = dbo.StockOrderOpening_Issuance.SOO_RefID INNER JOIN
                                                         dbo.VendRcvdDetail ON dbo.StockOrderOpening_Issuance.VID_RefID = dbo.VendRcvdDetail.Issue_RefID) AS TSFRemarks ON 
                         Tab2.LotNo = TSFRemarks.LotNo AND Tab1.ProcessID = TSFRemarks.ProcessID LEFT OUTER JOIN
                             (SELECT        dbo.VendRcvdDetail.LotNo, dbo.VendRcvdDetail.ProcessID, dbo.LotTransferDetails.SplitQty, dbo.LotTransferDetails.Type, 
                                                         dbo.LotTransferDetails.ToOrderNo, VendRcvdDetail_1.LotNo AS ToLotNo
                                FROM            dbo.LotTransferDetails INNER JOIN
                                                         dbo.VendRcvdDetail ON dbo.LotTransferDetails.VRD_From_RefID = dbo.VendRcvdDetail.EntryID INNER JOIN
                                                         dbo.VendRcvdDetail AS VendRcvdDetail_1 ON dbo.LotTransferDetails.VRD_To_RefID = VendRcvdDetail_1.EntryID) AS TSplit ON 
                         Tab2.LotNo = TSplit.LotNo AND Tab1.ProcessID = TSplit.ProcessID LEFT OUTER JOIN
                             (SELECT        dbo.VendRcvdDetail.LotNo AS FromLotNo, dbo.VendRcvdDetail.ProcessID, dbo.LotTransferDetails.SplitQty AS SplitQtyFrom, 
                                                         dbo.LotTransferDetails.Type AS SplitTypeFrom, dbo.LotTransferDetails.FromOrderNo, VendRcvdDetail_1.LotNo
                                FROM            dbo.LotTransferDetails INNER JOIN
                                                         dbo.VendRcvdDetail ON dbo.LotTransferDetails.VRD_From_RefID = dbo.VendRcvdDetail.EntryID INNER JOIN
                                                         dbo.VendRcvdDetail AS VendRcvdDetail_1 ON dbo.LotTransferDetails.VRD_To_RefID = VendRcvdDetail_1.EntryID) AS TSplitFrom ON 
                         Tab2.LotNo = TSplitFrom.LotNo AND Tab1.ProcessID = TSplitFrom.ProcessID LEFT OUTER JOIN
                             (SELECT        LotNo, MAX(EntryID) AS LotLastEntryID
                                FROM            dbo.VendRcvdDetail WHERE LotNo=@LotNo
                                GROUP BY LotNo) AS TLastRcv ON Tab2.LotNo = TLastRcv.LotNo LEFT OUTER JOIN
                             (SELECT        dbo.VendRcvdDetail.ProcessID, dbo.VendRcvdDetail.LotNo, SUM(dbo.TransferredToReadyFinishLotsDetail.Qty) AS TTRFQty
                                FROM            dbo.TransferredToReadyFinishLotsDetail INNER JOIN
                                                         dbo.VendRcvdDetail ON dbo.TransferredToReadyFinishLotsDetail.VRD_RefID = dbo.VendRcvdDetail.EntryID
                                WHERE VendRcvdDetail.LotNo=@LotNo GROUP BY dbo.VendRcvdDetail.ProcessID, dbo.VendRcvdDetail.LotNo ) AS Tab_TTRF ON Tab1.ProcessID = Tab_TTRF.ProcessID AND 
                         Tab2.LotNo = Tab_TTRF.LotNo
							WHERE Tab2.LotNo=@LotNo
		END	*/
SET ANSI_WARNINGS ON
END
GO

CREATE OR ALTER PROCEDURE [dbo].[PTC_SP_1] (@LotNo AS VARCHAR(50)) AS

BEGIN
SET ANSI_WARNINGS OFF

DECLARE @RepairGeneratedFrom AS VARCHAR(4000),@RepairType AS VARCHAR(255)


--IF @Repair=0
	BEGIN
	SET @RepairGeneratedFrom=''
	SET @RepairType=''
SELECT     TFOrderItems.OrderNo, TFOrderItems.ItemCode, TFOrderItems.Qty, dbo.FCustomerOrders.CustCode, dbo.FCustomerOrders.Country, dbo.FCustomerOrders.DT, 
                      dbo.Items.ItemName, dbo.Items.ItemSize, dbo.Items.SizeUnit, dbo.Items.Type, dbo.FCustomerOrders.Packaging
					  ,CASE WHEN DeliveryDTItem='1/1/1900' THEN dbo.FCustomerOrders.DeliveryDT ELSE DeliveryDTItem END AS DeliveryDT
                      ,dbo.ItemProcesses.SNO, dbo.ItemProcesses.ProcessID, dbo.Processes.Description, Tab1.RcvdQty, Tab2.LotNo, Tab1.VendID, dbo.Makers.VenderName, Tab1.RcvDT, 
                      TabIss.IssDT, TabMin.MinProcessID, Tab1.Wastage, Tab1.PassQty, dbo.Employees.empid, dbo.Employees.name, Tab1.ReWorkQty, TFOrderItems.CompItemCode, 
                      dbo.Processes.Code, dbo.Processes.ProcessNameUrdu, dbo.Items.ItemNameUrdu, EmpRcv.name AS RcvEmpName, Tab1.Rate, TMasterPO.MasterPONo
                      ,TSteelBatch.SteelBatchNo
					  --,TForgeBatch.ForgeBatchNo
					  ,TSteelBatch.ForgeBatchNo
					  , TSFRemarks.SFRemarks, TSplit.SplitQty, ISNULL(TSplit.Type, TSplitFrom.SplitTypeFrom) AS SplitType, 
                      TSplit.ToOrderNo, TSplit.ToLotNo, TSplitFrom.FromOrderNo, TSplitFrom.FromLotNo, TSplitFrom.SplitQtyFrom, TSplit.FromItemCode, TSplit.ToItemCode, 
                      TMasterPO.RecieptID, TIssFromSF.LotNo AS IssSFLotNo, TIssFromSF.IssSFRemarks, TIssFromSF.IssSFQty, TLastRcv.LotLastEntryID, 
                      dbo.VFOrderItemsWithShippedQtyItemCodeWise.ShippedQty, ISNULL(dbo.VLotWithPolishingMaker.VenderName, '') AS FromLotPolisher, 
                      ISNULL(dbo.VLotWithFirstProcessOfIssuance.VenderName, '') AS FromLotOriginator, Tab1.RcvTemperValue, TabIss.IssSpecialInstructions, 
                      dbo.FCustomerOrders.InternalRefNo, dbo.ProcessGroups.PGReference, Tab1.ReceivedEntryID, dbo.VendReceived_Employees_F(Tab1.ReceivedEntryID) 
                      AS Employees, TFOrderItems.Quality, dbo.GetAllRepairLots_F(Tab2.LotNo) AS RepairLots, TSplit.LotTransferRemarks, 
                      TSplitFrom.LotTransferRemarks AS LotTransferRemarksFrom, Tab1.UserName, Tab1.MachineName,ProcessGroupsProcesses.Hub_Name,VI_DT,VLotsWithMasterPONo.Batch_No AS VID_Batch_No,VLotsWithMasterPONo.MasterPODT
					  ,TIssFromSF.StoreName,TIssFromSF.Location
					  ,Lots_List.Batch_No AS Lots_List_Batch_No,Lots_List.Reference_LotNo,Lots_List.Mill_Certificate_No,Mill_Certificate_No_From_RM
					  ,Items.HRC_From,Items.HRC_To
					  ,Lots_List.Forge_Batch_No AS Lots_List_Forge_Batch_No
					  ,ISNULL(TFCC.FinQuality,'') AS FinQuality,ISNULL(TFCC.SpecialInstructions,'') AS SpecialInstructions_FCC,
					  ISNULL(TFCC.StampInstructions,'') AS StampInstructions,ISNULL(TFCC.PackingInstructions,'') AS PackingInstructions,
					  TFCC.MDMA,TFCC.SFDA_Listing_No,TFCC.MD_Group
					  ,dbo.Items.ItemPic,TFOrderItems.FOI_Stamps,ISNULL(FCustomerOrders.Packaging,'') AS SpecialInstruction,
					  TFOrderItems.BatchNo,ISNULL(Makers.Maker_Second_Name,'') AS Maker_Second_Name,
					  ISNULL(Finishing,'') AS Finishing,ISNULL(ItemOrderNo,'') AS ItemOrderNo,
					  DeliveryDate,ISNULL(ItemRemarks,'') AS ItemRemarks
					  ,TabIss.VI_UserName,TabIss.VI_MachineName
					  ,TabIss.IssEmpID,TabIss.Iss_Emp_Name,Items.Additional_Detail
FROM         (SELECT     OrderNo, ItemCode, CompItemCode, Quality, SUM(Qty) AS Qty,DeliveryDTItem,FOI_Stamps,IW_BatchNo AS BatchNo,
				Finishing AS Finishing,ItemOrderNo,DeliveryDate,ItemRemarks
                       FROM          dbo.VrptOrders_ForProduction_Simple 
                       GROUP BY OrderNo, ItemCode, CompItemCode, Quality,DeliveryDTItem,FOI_Stamps,IW_BatchNo,
					   Finishing,ItemOrderNo,DeliveryDate,ItemRemarks) AS TFOrderItems LEFT OUTER JOIN
                      dbo.VFOrderItemsWithShippedQtyItemCodeWise ON TFOrderItems.OrderNo = dbo.VFOrderItemsWithShippedQtyItemCodeWise.OrderNo AND 
                      TFOrderItems.CompItemCode = dbo.VFOrderItemsWithShippedQtyItemCodeWise.CompItemCode INNER JOIN
                      dbo.Items ON TFOrderItems.CompItemCode = dbo.Items.ItemID INNER JOIN
                      dbo.ItemProcesses ON dbo.Items.ItemID = dbo.ItemProcesses.ItemID INNER JOIN
                      dbo.FCustomerOrders ON TFOrderItems.OrderNo = dbo.FCustomerOrders.OrderNo INNER JOIN
                      dbo.Processes ON dbo.ItemProcesses.ProcessID = dbo.Processes.ProcessID LEFT OUTER JOIN
                      dbo.ItemProcessGroups ON dbo.Items.ItemID = dbo.ItemProcessGroups.ItemID LEFT OUTER JOIN
                      dbo.ProcessGroups ON dbo.ItemProcessGroups.PG_RefID = dbo.ProcessGroups.EntryID LEFT OUTER JOIN
					  ProcessGroupsProcesses ON dbo.ProcessGroups.EntryID=ProcessGroupsProcesses.Group_RefID AND ItemProcesses.ProcessID=ProcessGroupsProcesses.Process_RefID
					  LEFT OUTER JOIN
                          (SELECT DISTINCT OrderNo, ItemCode, LotNo
                             FROM dbo.VendRcvdDetail WHERE LotNo=@LotNo) AS Tab2 ON TFOrderItems.OrderNo = Tab2.OrderNo AND TFOrderItems.CompItemCode = Tab2.ItemCode LEFT OUTER JOIN
                          (SELECT     VendRcvdDetail_2.ProcessID,VendRcvdDetail_2.LotNo,VendRcvdDetail_2.Rate,dbo.VendReceived.EmpID, 
                                                   dbo.VendReceived.TemperValue AS RcvTemperValue,VendRcvdDetail_2.RcvdQty,CASE WHEN VendRcvdDetail_2.ReqAuth = 1 THEN 0 ELSE VendRcvdDetail_2.RcvdQty END AS PassQty, 
                                                   dbo.VendReceived.VendID,dbo.VendReceived.DT AS RcvDT,VendRcvdDetail_2.Wastage,VendRcvdDetail_2.ReWorkQty, 
                                                   dbo.VendReceived.EntryID AS ReceivedEntryID, dbo.VendReceived.UserName, dbo.VendReceived.MachineName,VendRcvdDetail_2.EntryID AS VRD_EntryID,VendIssued.DT AS VI_DT
												   ,VendIssdDetail.Batch_No AS VID_Batch_No--,VendIssued.UserName AS VI_UserName,VendIssued.MachineName AS VI_MachineName
                             FROM dbo.VendRcvdDetail AS VendRcvdDetail_2 
							 INNER JOIN dbo.VendReceived ON dbo.VendReceived.EntryID = VendRcvdDetail_2.RefID
							 LEFT JOIN VendIssdDetail ON VendRcvdDetail_2.Issue_RefID=VendIssdDetail.EntryID
							 LEFT JOIN VendIssued ON VendIssdDetail.RefID=VendIssued.EntryID
								WHERE VendRcvdDetail_2.LotNo=@LotNo) AS Tab1 ON Tab2.LotNo = Tab1.LotNo AND 
                      dbo.ItemProcesses.ProcessID = Tab1.ProcessID LEFT OUTER JOIN
                          (SELECT     dbo.VendIssued.ProcessID, dbo.VendIssdDetail.LotNo, dbo.VendIssdDetail_MoreDetails.EmpID,
                                                   dbo.VendIssued.SpecialInstructions AS IssSpecialInstructions, MIN(dbo.VendIssued.DT) AS IssDT
												   ,MIN(VendIssued.UserName) AS VI_UserName,MIN(VendIssued.MachineName) AS VI_MachineName
												   ,VendIssued.IssEmpID,VEmp.Name AS Iss_Emp_Name
                             FROM         dbo.VendIssued INNER JOIN
                                                   dbo.VendIssdDetail ON dbo.VendIssued.EntryID = dbo.VendIssdDetail.RefID LEFT OUTER JOIN
                                                   dbo.VendIssdDetail_MoreDetails ON dbo.VendIssdDetail.EntryID = dbo.VendIssdDetail_MoreDetails.VID_RefID
												   LEFT JOIN VEmp ON dbo.VendIssued.IssEmpID=VEmp.EmpID
                             GROUP BY dbo.VendIssued.ProcessID, dbo.VendIssdDetail.LotNo, dbo.VendIssdDetail_MoreDetails.EmpID, dbo.VendIssued.SpecialInstructions,VendIssued.IssEmpID,VEmp.Name) AS TabIss ON 
                      Tab2.LotNo = TabIss.LotNo AND dbo.ItemProcesses.ProcessID = TabIss.ProcessID LEFT OUTER JOIN
                          (SELECT     VendRcvdDetail_1.LotNo, MIN(VendReceived_1.ProcessID) AS MinProcessID
                             FROM         dbo.VendReceived AS VendReceived_1 INNER JOIN
                                                   dbo.VendRcvdDetail AS VendRcvdDetail_1 ON VendReceived_1.EntryID = VendRcvdDetail_1.RefID  WHERE VendRcvdDetail_1.LotNo=@LotNo
                             GROUP BY VendRcvdDetail_1.LotNo) AS TabMin ON Tab2.LotNo = TabMin.LotNo LEFT OUTER JOIN
                      dbo.Makers ON Tab1.VendID = dbo.Makers.VendID LEFT OUTER JOIN
                      dbo.Employees ON TabIss.EmpID = dbo.Employees.empid LEFT OUTER JOIN
                      dbo.Employees AS EmpRcv ON Tab1.EmpID = EmpRcv.empid LEFT OUTER JOIN
                          (SELECT     TNest1.LotNo, VendIssued_2.EntryID, VendIssued_2.MasterPONo, VendIssued_2.RecieptID
                             FROM         dbo.VendIssued AS VendIssued_2 INNER JOIN
                                                   dbo.VendIssdDetail AS VendIssdDetail_2 ON VendIssued_2.EntryID = VendIssdDetail_2.RefID INNER JOIN
                                                       (SELECT     LotNo, MIN(Issue_RefID) AS Issue_RefID
                                                          FROM         dbo.VendRcvdDetail AS VendRcvdDetail_7 WHERE LotNo=@LotNo
                                                          GROUP BY LotNo) AS TNest1 ON TNest1.Issue_RefID = VendIssdDetail_2.EntryID) AS TMasterPO ON 
                      Tab2.LotNo = TMasterPO.LotNo LEFT OUTER JOIN
                          (SELECT     dbo.RawMaterialIssuance.VI_RefID, MAX(dbo.RawMaterialIssuanceDetail.PORefNo) AS SteelBatchNo,MAX(VendRcvdDetailPO.RcvID) AS ForgeBatchNo
								,MAX(MaterialLocationwiseStatus.Mill_Certificate_No) AS Mill_Certificate_No_From_RM
                             FROM         dbo.RawMaterialIssuance INNER JOIN
                                                   dbo.RawMaterialIssuanceDetail ON dbo.RawMaterialIssuance.IssNo = dbo.RawMaterialIssuanceDetail.IssNo INNER JOIN
                                                   dbo.RM ON dbo.RawMaterialIssuanceDetail.RMID1 = dbo.RM.RMID1
												   INNER JOIN RMID_MLS_Details ON RawMaterialIssuanceDetail.EntryID=RMID_MLS_Details.RMID_RefID
												   INNER JOIN MaterialLocationwiseStatus ON RMID_MLS_Details.MLS_RefID=MaterialLocationwiseStatus.EntryID
												   INNER JOIN VendRcvdDetailPO ON MaterialLocationwiseStatus.Rcvd_RefID=VendRcvdDetailPO.EntryID
                             WHERE     (dbo.RM.GroupID = 1)
                             GROUP BY dbo.RawMaterialIssuance.VI_RefID) AS TSteelBatch ON TMasterPO.EntryID = TSteelBatch.VI_RefID LEFT OUTER JOIN
                          /*(SELECT     RawMaterialIssuance_1.VI_RefID, MAX(RawMaterialIssuanceDetail_1.PORefNo) AS ForgeBatchNo
                             FROM         dbo.RawMaterialIssuance AS RawMaterialIssuance_1 INNER JOIN
                                                   dbo.RawMaterialIssuanceDetail AS RawMaterialIssuanceDetail_1 ON RawMaterialIssuance_1.IssNo = RawMaterialIssuanceDetail_1.IssNo INNER JOIN
                                                   dbo.RM AS RM_1 ON RawMaterialIssuanceDetail_1.RMID1 = RM_1.RMID1
                             WHERE     (RM_1.GroupID IN (25, 27, 28, 29, 30))
                             GROUP BY RawMaterialIssuance_1.VI_RefID) AS TForgeBatch ON TMasterPO.EntryID = TForgeBatch.VI_RefID LEFT OUTER JOIN*/
						
                          (SELECT     VendRcvdDetail_6.ProcessID, VendRcvdDetail_6.LotNo, '' AS SFRemarks
                             FROM         dbo.VendRcvdDetail AS VendRcvdDetail_6 INNER JOIN
                                                       (SELECT     VID_RefID
                                                          FROM         dbo.StockOrderOpening_Issuance
                                                          GROUP BY VID_RefID) AS TSOO_I ON TSOO_I.VID_RefID = VendRcvdDetail_6.Issue_RefID) AS TSFRemarks ON Tab2.LotNo = TSFRemarks.LotNo AND
                       Tab1.ProcessID = TSFRemarks.ProcessID LEFT OUTER JOIN
                          (SELECT     VendRcvdDetail_5.LotNo, VendRcvdDetail_5.ProcessID, dbo.LotTransferDetails.SplitQty, dbo.LotTransferDetails.Type, 
                                                   dbo.LotTransferDetails.ToOrderNo, VendRcvdDetail_1.LotNo AS ToLotNo, dbo.LotTransferDetails.FromItemCode, dbo.LotTransferDetails.ToItemCode, 
                                                   dbo.LotTransferDetails.LotTransferRemarks
                             FROM         dbo.LotTransferDetails INNER JOIN
                                                   dbo.VendRcvdDetail AS VendRcvdDetail_5 ON dbo.LotTransferDetails.VRD_From_RefID = VendRcvdDetail_5.EntryID LEFT OUTER JOIN
                                                   dbo.VendRcvdDetail AS VendRcvdDetail_1 ON dbo.LotTransferDetails.VRD_To_RefID = VendRcvdDetail_1.EntryID) AS TSplit ON 
                      Tab2.LotNo = TSplit.LotNo AND Tab1.ProcessID = TSplit.ProcessID LEFT OUTER JOIN
                          (SELECT     VendRcvdDetail_4.LotNo AS FromLotNo, VendRcvdDetail_4.ProcessID, LotTransferDetails_1.SplitQty AS SplitQtyFrom, 
                                                   LotTransferDetails_1.Type AS SplitTypeFrom, LotTransferDetails_1.FromOrderNo, VendRcvdDetail_1.LotNo, 
                                                   LotTransferDetails_1.LotTransferRemarks
                             FROM         dbo.LotTransferDetails AS LotTransferDetails_1 INNER JOIN
                                                   dbo.VendRcvdDetail AS VendRcvdDetail_4 ON LotTransferDetails_1.VRD_From_RefID = VendRcvdDetail_4.EntryID INNER JOIN
                                                   dbo.VendRcvdDetail AS VendRcvdDetail_1 ON LotTransferDetails_1.VRD_To_RefID = VendRcvdDetail_1.EntryID) AS TSplitFrom ON 
                      Tab2.LotNo = TSplitFrom.LotNo AND Tab1.ProcessID = TSplitFrom.ProcessID LEFT OUTER JOIN
                          (SELECT     VendIssued_1.ProcessID, VendIssdDetail_1.LotNo, '' AS IssSFRemarks, VendIssdDetail_1.IssQty AS IssSFQty,TSOO_I_1.StoreName,TSOO_I_1.Location
                             FROM         dbo.VendIssdDetail AS VendIssdDetail_1 INNER JOIN
                                                       (SELECT     VID_RefID,StoreName,Location
                                                          FROM         dbo.StockOrderOpening_Issuance AS StockOrderOpening_Issuance_1
														  LEFT JOIN StockOrderOpening ON StockOrderOpening_Issuance_1.SOO_RefID=StockOrderOpening.EntryID
														  LEFT JOIN VStoreShelfs ON StockOrderOpening.Shelf_RefID=VStoreShelfs.EntryID
                                                          GROUP BY VID_RefID,StoreName,Location) AS TSOO_I_1 ON VendIssdDetail_1.EntryID = TSOO_I_1.VID_RefID INNER JOIN
                                                   dbo.VendIssued AS VendIssued_1 ON VendIssdDetail_1.RefID = VendIssued_1.EntryID) AS TIssFromSF ON Tab2.LotNo = TIssFromSF.LotNo AND 
                      Tab1.ProcessID = TIssFromSF.ProcessID LEFT OUTER JOIN
                          (SELECT     LotNo, MAX(EntryID) AS LotLastEntryID
                             FROM         dbo.VendRcvdDetail AS VendRcvdDetail_3 WHERE LotNo=@LotNo
                             GROUP BY LotNo) AS TLastRcv ON Tab2.LotNo = TLastRcv.LotNo LEFT OUTER JOIN
                      dbo.VLotWithPolishingMaker ON TSplitFrom.FromLotNo = dbo.VLotWithPolishingMaker.LotNo LEFT OUTER JOIN
                      dbo.VLotWithFirstProcessOfIssuance ON TSplitFrom.FromLotNo = dbo.VLotWithFirstProcessOfIssuance.LotNo
					  LEFT OUTER JOIN VLotsWithMasterPONo ON Tab2.LotNo=VLotsWithMasterPONo.LotNO
					  LEFT OUTER JOIN Lots_List ON Tab2.LotNo=Lots_List.LotNo
					  LEFT OUTER JOIN FCustomerCatalog TFCC ON FCustomerOrders.CustCode=TFCC.CustCode AND FCustomerOrders.Country=TFCC.Country AND TFOrderItems.CompItemCode=TFCC.CompItemID
	WHERE Tab2.LotNo=@LotNo
	
	END
/*ELSE
	BEGIN
	
	SET @RepairGeneratedFrom=dbo.RepairLot_OriginalLots_F(@LotNo)
	SELECT        TFOrderItems.OrderNo, TFOrderItems.ItemCode, TFOrderItems.Qty, dbo.FCustomerOrders.CustCode, dbo.FCustomerOrders.Country, dbo.FCustomerOrders.DT, 
                         dbo.Items.ItemName, dbo.Items.ItemSize, dbo.Items.SizeUnit, dbo.Items.Type, dbo.FCustomerOrders.Packaging, dbo.FCustomerOrders.DeliveryDT, 
                         dbo.RepairTypeProcesses.SeqNo AS SNO, dbo.RepairTypeProcesses.ProcessID, dbo.Processes.Description, Tab1.RcvdQty, Tab2.LotNo, Tab1.VendID, 
                         dbo.Makers.VenderName, Tab1.RcvDT, TabIss.IssDT, TabMin.MinProcessID, Tab1.Wastage, Tab1.PassQty, dbo.Employees.empid, dbo.Employees.name, 
                         Tab1.ReWorkQty, TFOrderItems.CompItemCode, dbo.Processes.Code, dbo.Processes.ProcessNameUrdu, dbo.Items.ItemNameUrdu, 
                         EmpRcv.name AS RcvEmpName, Tab1.Rate, TMasterPO.MasterPONo, TSteelBatch.SteelBatchNo, TForgeBatch.ForgeBatchNo, TSFRemarks.SFRemarks, 
                         TSplit.SplitQty, ISNULL(TSplit.Type, TSplitFrom.SplitTypeFrom) AS SplitType, TSplit.ToOrderNo, TSplit.ToLotNo, TSplitFrom.FromOrderNo, TSplitFrom.FromLotNo, 
                         TSplitFrom.SplitQtyFrom, @RepairGeneratedFrom AS RepairGeneratedFrom, dbo.RepairTypes.RepairType, TLastRcv.LotLastEntryID, 
                         Tab_TTRF.TTRFQty, Tab1.ReceivedEntryID, dbo.VendReceived_Employees_F(Tab1.ReceivedEntryID) AS Employees,'' AS IssEmpID,'' AS IssEmpName,dbo.GetMainLots_Repair_F(@LotNo) AS RepairLotsGenerated
FROM            (SELECT        OrderNo, ItemCode, CompItemCode, SUM(Qty) AS Qty
                           FROM            dbo.FOrderItems
                           GROUP BY OrderNo, ItemCode, CompItemCode) AS TFOrderItems INNER JOIN
                         dbo.Items ON TFOrderItems.CompItemCode = dbo.Items.ItemID INNER JOIN
                         dbo.FCustomerOrders ON TFOrderItems.OrderNo = dbo.FCustomerOrders.OrderNo LEFT OUTER JOIN
                             (SELECT DISTINCT OrderNo, ItemCode, LotNo, Repair_RefID
                                FROM            dbo.VendRcvdDetail WHERE LotNo=@LotNo) AS Tab2 ON TFOrderItems.OrderNo = Tab2.OrderNo AND TFOrderItems.CompItemCode = Tab2.ItemCode INNER JOIN
                         dbo.RepairTypeProcesses ON dbo.RepairTypeProcesses.Repair_RefID = Tab2.Repair_RefID INNER JOIN
                         dbo.Processes ON dbo.RepairTypeProcesses.ProcessID = dbo.Processes.ProcessID INNER JOIN
                         dbo.RepairTypes ON Tab2.Repair_RefID = dbo.RepairTypes.EntryID LEFT OUTER JOIN
                             (SELECT        VendRcvdDetail_2.ProcessID, VendRcvdDetail_2.LotNo, VendRcvdDetail_2.Rate, dbo.VendReceived.EmpID, VendRcvdDetail_2.RcvdQty, 
                                                         CASE WHEN ReqAuth = 1 THEN 0 ELSE RcvdQty END AS PassQty, dbo.VendReceived.VendID, dbo.VendReceived.DT AS RcvDT, 
                                                         VendRcvdDetail_2.Wastage, VendRcvdDetail_2.ReWorkQty, dbo.VendReceived.EntryID AS ReceivedEntryID
                                FROM            dbo.VendRcvdDetail AS VendRcvdDetail_2 INNER JOIN
                                                         dbo.VendReceived ON dbo.VendReceived.EntryID = VendRcvdDetail_2.RefID) AS Tab1 ON Tab2.LotNo = Tab1.LotNo AND 
                         dbo.RepairTypeProcesses.ProcessID = Tab1.ProcessID LEFT OUTER JOIN
                             (SELECT        dbo.VendIssued.ProcessID, dbo.VendIssdDetail.LotNo, dbo.VendIssdDetail_MoreDetails.EmpID, dbo.VendIssdDetail.Repair_RefID, 
                                                         MIN(dbo.VendIssued.DT) AS IssDT
                                FROM            dbo.VendIssued INNER JOIN
                                                         dbo.VendIssdDetail ON dbo.VendIssued.EntryID = dbo.VendIssdDetail.RefID LEFT OUTER JOIN
                                                         dbo.VendIssdDetail_MoreDetails ON dbo.VendIssdDetail.EntryID = dbo.VendIssdDetail_MoreDetails.VID_RefID WHERE dbo.VendIssdDetail.LotNo=@LotNo
                                GROUP BY dbo.VendIssued.ProcessID, dbo.VendIssdDetail.LotNo, dbo.VendIssdDetail_MoreDetails.EmpID, dbo.VendIssdDetail.Repair_RefID) AS TabIss ON 
                         Tab2.LotNo = TabIss.LotNo AND dbo.RepairTypeProcesses.ProcessID = TabIss.ProcessID LEFT OUTER JOIN
                             (SELECT        VendRcvdDetail_1.LotNo, MIN(VendReceived_1.ProcessID) AS MinProcessID
                                FROM            dbo.VendReceived AS VendReceived_1 INNER JOIN
                                                         dbo.VendRcvdDetail AS VendRcvdDetail_1 ON VendReceived_1.EntryID = VendRcvdDetail_1.RefID WHERE VendRcvdDetail_1.LotNo=@LotNo
                                GROUP BY VendRcvdDetail_1.LotNo) AS TabMin ON Tab2.LotNo = TabMin.LotNo LEFT OUTER JOIN
                         dbo.Makers ON Tab1.VendID = dbo.Makers.VendID LEFT OUTER JOIN
                         dbo.Employees ON TabIss.EmpID = dbo.Employees.empid LEFT OUTER JOIN
                         dbo.Employees AS EmpRcv ON Tab1.EmpID = EmpRcv.empid LEFT OUTER JOIN
                             (SELECT        TNest1.LotNo, dbo.VendIssued.EntryID, dbo.VendIssued.MasterPONo
                                FROM            dbo.VendIssued INNER JOIN
                                                         dbo.VendIssdDetail ON dbo.VendIssued.EntryID = dbo.VendIssdDetail.RefID INNER JOIN
                                                             (SELECT        LotNo, MIN(Issue_RefID) AS Issue_RefID
                                                                FROM            dbo.VendRcvdDetail WHERE LotNo=@LotNo
                                                                GROUP BY LotNo) AS TNest1 ON TNest1.Issue_RefID = dbo.VendIssdDetail.EntryID) AS TMasterPO ON 
                         Tab1.LotNo = TMasterPO.LotNo LEFT OUTER JOIN
                             (SELECT        dbo.RawMaterialIssuance.VI_RefID, MAX(dbo.RawMaterialIssuanceDetail.PORefNo) AS SteelBatchNo
                                FROM            dbo.RawMaterialIssuance INNER JOIN
                                                         dbo.RawMaterialIssuanceDetail ON dbo.RawMaterialIssuance.IssNo = dbo.RawMaterialIssuanceDetail.IssNo INNER JOIN
                                                         dbo.RM ON dbo.RawMaterialIssuanceDetail.RMID1 = dbo.RM.RMID1
                                WHERE        (dbo.RM.GroupID = 20)
                                GROUP BY dbo.RawMaterialIssuance.VI_RefID) AS TSteelBatch ON TMasterPO.EntryID = TSteelBatch.VI_RefID LEFT OUTER JOIN
                             (SELECT        dbo.RawMaterialIssuance.VI_RefID, MAX(dbo.RawMaterialIssuanceDetail.PORefNo) AS ForgeBatchNo
                                FROM            dbo.RawMaterialIssuance INNER JOIN
                                                         dbo.RawMaterialIssuanceDetail ON dbo.RawMaterialIssuance.IssNo = dbo.RawMaterialIssuanceDetail.IssNo INNER JOIN
                                                         dbo.RM ON dbo.RawMaterialIssuanceDetail.RMID1 = dbo.RM.RMID1
                                WHERE        (dbo.RM.GroupID IN (25, 27, 28, 29, 30))
                                GROUP BY dbo.RawMaterialIssuance.VI_RefID) AS TForgeBatch ON TMasterPO.EntryID = TForgeBatch.VI_RefID LEFT OUTER JOIN
                             (SELECT        dbo.VendRcvdDetail.ProcessID, dbo.VendRcvdDetail.LotNo, dbo.StockOrderOpening.Remarks AS SFRemarks
                                FROM            dbo.StockOrderOpening INNER JOIN
                                                         dbo.StockOrderOpening_Issuance ON dbo.StockOrderOpening.EntryID = dbo.StockOrderOpening_Issuance.SOO_RefID INNER JOIN
                                                         dbo.VendRcvdDetail ON dbo.StockOrderOpening_Issuance.VID_RefID = dbo.VendRcvdDetail.Issue_RefID) AS TSFRemarks ON 
                         Tab2.LotNo = TSFRemarks.LotNo AND Tab1.ProcessID = TSFRemarks.ProcessID LEFT OUTER JOIN
                             (SELECT        dbo.VendRcvdDetail.LotNo, dbo.VendRcvdDetail.ProcessID, dbo.LotTransferDetails.SplitQty, dbo.LotTransferDetails.Type, 
                                                         dbo.LotTransferDetails.ToOrderNo, VendRcvdDetail_1.LotNo AS ToLotNo
                                FROM            dbo.LotTransferDetails INNER JOIN
                                                         dbo.VendRcvdDetail ON dbo.LotTransferDetails.VRD_From_RefID = dbo.VendRcvdDetail.EntryID INNER JOIN
                                                         dbo.VendRcvdDetail AS VendRcvdDetail_1 ON dbo.LotTransferDetails.VRD_To_RefID = VendRcvdDetail_1.EntryID) AS TSplit ON 
                         Tab2.LotNo = TSplit.LotNo AND Tab1.ProcessID = TSplit.ProcessID LEFT OUTER JOIN
                             (SELECT        dbo.VendRcvdDetail.LotNo AS FromLotNo, dbo.VendRcvdDetail.ProcessID, dbo.LotTransferDetails.SplitQty AS SplitQtyFrom, 
                                                         dbo.LotTransferDetails.Type AS SplitTypeFrom, dbo.LotTransferDetails.FromOrderNo, VendRcvdDetail_1.LotNo
                                FROM            dbo.LotTransferDetails INNER JOIN
                                                         dbo.VendRcvdDetail ON dbo.LotTransferDetails.VRD_From_RefID = dbo.VendRcvdDetail.EntryID INNER JOIN
                                                         dbo.VendRcvdDetail AS VendRcvdDetail_1 ON dbo.LotTransferDetails.VRD_To_RefID = VendRcvdDetail_1.EntryID) AS TSplitFrom ON 
                         Tab2.LotNo = TSplitFrom.LotNo AND Tab1.ProcessID = TSplitFrom.ProcessID LEFT OUTER JOIN
                             (SELECT        LotNo, MAX(EntryID) AS LotLastEntryID
                                FROM            dbo.VendRcvdDetail WHERE LotNo=@LotNo
                                GROUP BY LotNo) AS TLastRcv ON Tab2.LotNo = TLastRcv.LotNo LEFT OUTER JOIN
                             (SELECT        dbo.VendRcvdDetail.ProcessID, dbo.VendRcvdDetail.LotNo, SUM(dbo.TransferredToReadyFinishLotsDetail.Qty) AS TTRFQty
                                FROM            dbo.TransferredToReadyFinishLotsDetail INNER JOIN
                                                         dbo.VendRcvdDetail ON dbo.TransferredToReadyFinishLotsDetail.VRD_RefID = dbo.VendRcvdDetail.EntryID
                                WHERE VendRcvdDetail.LotNo=@LotNo GROUP BY dbo.VendRcvdDetail.ProcessID, dbo.VendRcvdDetail.LotNo ) AS Tab_TTRF ON Tab1.ProcessID = Tab_TTRF.ProcessID AND 
                         Tab2.LotNo = Tab_TTRF.LotNo
							WHERE Tab2.LotNo=@LotNo
		END	*/
SET ANSI_WARNINGS ON
END
GO

-- -------------------------------------------------------------
-- Ensure VMakerAssItems contains MakerDescription cleanly
-- -------------------------------------------------------------
CREATE OR ALTER VIEW [dbo].[VMakerAssItems]
AS
SELECT     dbo.ItemGroups.Description, dbo.Makers.VenderName, dbo.Makers.PhaseID, dbo.Makers.VendType, dbo.Makers.RepairDedRate, 
           dbo.VItems1.ItemName AS ItemName, dbo.VItems1.Unit, dbo.VItems1.Type, dbo.Makers.Active, dbo.Makers.VendID1, dbo.VendAssItems.EntryID, 
           dbo.VendAssItems.VendID, dbo.VendAssItems.ProcessID, dbo.VendAssItems.ItemID, dbo.VendAssItems.Rate, dbo.VendAssItems.PlantRate, 
           dbo.VendAssItems.StampRate, dbo.VendAssItems.SnaffRate, dbo.VendAssItems.Unit AS AssignedUnit, dbo.VendAssItems.Remarks,
           ISNULL(dbo.VItems1.MakerDescription, '') AS MakerDescription
FROM       dbo.VendAssItems 
INNER JOIN dbo.Makers ON dbo.VendAssItems.VendID = dbo.Makers.VendID 
INNER JOIN dbo.VItems1 ON dbo.VendAssItems.ItemID = dbo.VItems1.ItemID 
LEFT OUTER JOIN dbo.ItemGroups ON dbo.VItems1.GroupID = dbo.ItemGroups.ID
GO
END
GO



-- ====================================================================================================
-- SCRIPT: Users - Show_Customer_Order_No
-- ====================================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('Users') AND name = 'Show_Customer_Order_No')
BEGIN
    ALTER TABLE [dbo].[Users] ADD [Show_Customer_Order_No] BIT NOT NULL CONSTRAINT [DF_Users_Show_Customer_Order_No] DEFAULT (0);
    PRINT 'Added column Show_Customer_Order_No to Users';
END
GO

-- ====================================================================================================
-- SCRIPT: FOrderItems - Authorized, AuthorizedQty, AuthorizedBy, AuthorizedDT
-- ====================================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('FOrderItems') AND name = 'Authorized')
BEGIN
    ALTER TABLE [dbo].[FOrderItems] ADD [Authorized] BIT NOT NULL CONSTRAINT [DF_FOrderItems_Authorized] DEFAULT (0);
    PRINT 'Added column Authorized to FOrderItems';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('FOrderItems') AND name = 'AuthorizedQty')
BEGIN
    ALTER TABLE [dbo].[FOrderItems] ADD [AuthorizedQty] INT NOT NULL CONSTRAINT [DF_FOrderItems_AuthorizedQty] DEFAULT (0);
    PRINT 'Added column AuthorizedQty to FOrderItems';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('FOrderItems') AND name = 'AuthorizedBy')
BEGIN
    ALTER TABLE [dbo].[FOrderItems] ADD [AuthorizedBy] VARCHAR(50) NULL;
    PRINT 'Added column AuthorizedBy to FOrderItems';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('FOrderItems') AND name = 'AuthorizedDT')
BEGIN
    ALTER TABLE [dbo].[FOrderItems] ADD [AuthorizedDT] DATETIME NULL;
    PRINT 'Added column AuthorizedDT to FOrderItems';
END
GO

-- Backfill existing authorized orders so their items match current authorization
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('FOrderItems') AND name = 'Authorized')
BEGIN
    UPDATE oi
    SET oi.Authorized = 1,
        oi.AuthorizedQty = ISNULL(oi.Qty, 0),
        oi.AuthorizedBy = ISNULL(co.AuthorizedBy, 'System'),
        oi.AuthorizedDT = ISNULL(co.AuthorizedDT, co.DT)
    FROM [dbo].[FOrderItems] oi
    INNER JOIN [dbo].[FCustomerOrders] co ON oi.OrderNo = co.OrderNo
    WHERE ISNULL(co.Authorized, 0) = 1 AND ISNULL(oi.Authorized, 0) = 0 AND ISNULL(oi.AuthorizedQty, 0) = 0;
    PRINT 'Backfilled FOrderItems authorization for existing authorized orders';
END
GO

-- -------------------------------------------------------------
-- Update VFOrderList to include TotalAuthorizedArticles & TotalPendingArticles
-- -------------------------------------------------------------
CREATE OR ALTER VIEW [dbo].[VFOrderList]
AS
SELECT     dbo.FCustomerOrders.OrderNo, dbo.FCustomerOrders.DT, dbo.FCustomerOrders.CustCode, dbo.FCustomerOrders.Country, dbo.F_OrderAmt(dbo.FCustomerOrders.OrderNo, 0) AS OrderAmt, 
                      dbo.ForeignCustomers.Curr, dbo.VFOrderTotalQty.TotalInvQty, dbo.VFOrderTotalQty.TotalOrderQty, dbo.FCustomerOrders.DeliveryDT, dbo.FCustomerOrders.CompanyRefID, 
                      dbo.Companies.CompanyName, dbo.VUnshippedOrderList.OrderNo AS UnshippedOrderNo, TotalShipped.TotalShippedQty, dbo.VCurrencyExchangeRates.ExchRate, 
                      dbo.FCustomerOrders.InternalRefNo, ROUND(dbo.VFOrderTotalQty.TotalWeight, 2) AS TotalWeight, TotalShipped.TotalShippedAmt, ISNULL(dbo.FCustomerFinalOrders.Cancelled, 0) AS Cancelled, 
                      TotalQuantities.TotalArticles, dbo.FCustomerFinalOrders.Remarks, TotalShipped.ShippedItemCount, TotalShipped.ShippedItemQty, TotalShipped.PartiallyShippedItemCount, 
                      TotalShipped.PartiallyShippedItemQty, TotalShipped.TotalBalanceQty, dbo.FCustomerOrders.OrderType, T1.TotalPlannedQty, dbo.FCustomerOrders.OrderPlanApproved
                      ,dbo.ForeignCustomers.LateOrderAlerts
                      ,ISNULL(dbo.FCustomerOrders.Authorized, 0) AS Authorized
                      ,dbo.FCustomerOrders.AuthorizedBy
                      ,dbo.FCustomerOrders.AuthorizedDT
                      ,ISNULL(TotalQuantities.TotalAuthorizedArticles, 0) AS TotalAuthorizedArticles
                      ,ISNULL(TotalQuantities.TotalPendingArticles, 0) AS TotalPendingArticles
FROM         dbo.FCustomerOrders INNER JOIN
                      dbo.ForeignCustomers ON dbo.FCustomerOrders.CustCode = dbo.ForeignCustomers.CustCode AND dbo.FCustomerOrders.Country = dbo.ForeignCustomers.Country INNER JOIN
                      dbo.VFOrderTotalQty ON dbo.FCustomerOrders.OrderNo = dbo.VFOrderTotalQty.OrderNo INNER JOIN
                      dbo.Companies ON dbo.FCustomerOrders.CompanyRefID = dbo.Companies.EntryID LEFT OUTER JOIN
                      dbo.VUnshippedOrderList ON dbo.FCustomerOrders.OrderNo = dbo.VUnshippedOrderList.OrderNo INNER JOIN
                          (SELECT     OrderNo, 
                                      SUM(Qty) AS TotalOrderQty, 
                                      COUNT(*) AS TotalArticles,
                                      SUM(CASE WHEN ISNULL(Authorized, 0) = 1 AND ISNULL(AuthorizedQty, 0) >= ISNULL(Qty, 0) THEN 1 ELSE 0 END) AS TotalAuthorizedArticles,
                                      SUM(CASE WHEN ISNULL(Authorized, 0) = 0 OR ISNULL(AuthorizedQty, 0) < ISNULL(Qty, 0) THEN 1 ELSE 0 END) AS TotalPendingArticles
                            FROM          dbo.FOrderItems
                            GROUP BY OrderNo) AS TotalQuantities ON dbo.FCustomerOrders.OrderNo = TotalQuantities.OrderNo LEFT OUTER JOIN
                      dbo.FCustomerFinalOrders ON dbo.FCustomerOrders.OrderNo = dbo.FCustomerFinalOrders.OrderNo LEFT OUTER JOIN
                          (SELECT     OrderNo, SUM(ShippedQty) AS TotalShippedQty, SUM(ShippedQty * Price) AS TotalShippedAmt, SUM(CASE WHEN ShippedQty >= Qty THEN 1 ELSE 0 END) AS ShippedItemCount, 
                                                   SUM(CASE WHEN ShippedQty >= Qty THEN ShippedQty ELSE 0 END) AS ShippedItemQty, SUM(CASE WHEN ShippedQty > 0 AND ShippedQty < Qty THEN 1 ELSE 0 END) 
                                                   AS PartiallyShippedItemCount, SUM(CASE WHEN ShippedQty > 0 AND ShippedQty < Qty THEN ShippedQty ELSE 0 END) AS PartiallyShippedItemQty, 
                                                   SUM(CASE WHEN ShippedQty >= Qty THEN 0 ELSE Qty - ShippedQty END) AS TotalBalanceQty
                            FROM          dbo.VFOrderItemswithShippedQty
                            WHERE      (CompItemCode IN
                                                       (SELECT     ItemID
                                                         FROM          dbo.Items))
                            GROUP BY OrderNo) AS TotalShipped ON dbo.FCustomerOrders.OrderNo = TotalShipped.OrderNo LEFT OUTER JOIN
                      dbo.VCurrencyExchangeRates ON dbo.ForeignCustomers.Curr = dbo.VCurrencyExchangeRates.Currency LEFT OUTER JOIN
                          (SELECT     OrderNo, SUM(Qty) AS TotalPlannedQty
                            FROM          dbo.OrderPlanningDetails
                            GROUP BY OrderNo) AS T1 ON dbo.FCustomerOrders.OrderNo = T1.OrderNo
GO

-- ============================================================================
-- Maker Purchase Orders: Attach Master PO Maker Signed Scanned Copy
-- ============================================================================
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('VendIssued') AND name = 'MakerSignedCopyPath')
BEGIN
    ALTER TABLE VendIssued ADD MakerSignedCopyPath NVARCHAR(500) NULL;
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('VendIssued') AND name = 'MakerSignedCopyFileName')
BEGIN
    ALTER TABLE VendIssued ADD MakerSignedCopyFileName NVARCHAR(255) NULL;
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('VendIssued') AND name = 'MakerSignedCopyUploadedAt')
BEGIN
    ALTER TABLE VendIssued ADD MakerSignedCopyUploadedAt DATETIME NULL;
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('VendIssued') AND name = 'MakerSignedCopyUploadedBy')
BEGIN
    ALTER TABLE VendIssued ADD MakerSignedCopyUploadedBy NVARCHAR(100) NULL;
END
GO

-- ============================================================================
-- Setup Module Security & Screen Rights: SetupMainLink and Setup MenuOptions
-- ============================================================================
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='Users' AND COLUMN_NAME='SetupMainLink')
BEGIN
    ALTER TABLE dbo.Users ADD SetupMainLink BIT NOT NULL CONSTRAINT DF_Users_SetupMainLink DEFAULT (0);
    PRINT 'Added SetupMainLink to Users table';
END
GO

IF EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='Users' AND COLUMN_NAME='SetupMainLink')
BEGIN
    EXEC sp_executesql N'UPDATE dbo.Users SET SetupMainLink = 1 WHERE UserManagement = 1;';
    PRINT 'Updated SetupMainLink = 1 for users with UserManagement = 1';
END
GO

-- Ensure Setup Menu Options exist in MenuOptions table
IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'SetupHub')
BEGIN
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName)
    VALUES ('SetupHub', 'Setups Hub', 'Setup', 'SetupDashboard');
END
ELSE
BEGIN
    UPDATE MenuOptions SET ModuleName = 'Setup', OptionName = 'Setups Hub', FormName = 'SetupDashboard' WHERE OptionID = 'SetupHub';
END
GO

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'SetupUsers')
BEGIN
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName)
    VALUES ('SetupUsers', 'User Management', 'Setup', 'UsersList');
END
ELSE
BEGIN
    UPDATE MenuOptions SET ModuleName = 'Setup', OptionName = 'User Management', FormName = 'UsersList' WHERE OptionID = 'SetupUsers';
END
GO

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'OfficeMinuteTypes')
BEGIN
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName)
    VALUES ('OfficeMinuteTypes', 'Minute Types', 'Setup', 'MinuteTypesAdmin');
END
ELSE
BEGIN
    UPDATE MenuOptions SET ModuleName = 'Setup', OptionName = 'Minute Types', FormName = 'MinuteTypesAdmin' WHERE OptionID = 'OfficeMinuteTypes';
END
GO

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'OfficeEmailSettings')
BEGIN
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName)
    VALUES ('OfficeEmailSettings', 'Email & SMTP Settings', 'Setup', 'EmailSettingsAdmin');
END
ELSE
BEGIN
    UPDATE MenuOptions SET ModuleName = 'Setup', OptionName = 'Email & SMTP Settings', FormName = 'EmailSettingsAdmin' WHERE OptionID = 'OfficeEmailSettings';
END
GO

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'IntraOfficeHealth')
BEGIN
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName)
    VALUES ('IntraOfficeHealth', 'System Diagnostics', 'Setup', 'SystemHealth');
END
ELSE
BEGIN
    UPDATE MenuOptions SET ModuleName = 'Setup', OptionName = 'System Diagnostics', FormName = 'SystemHealth' WHERE OptionID = 'IntraOfficeHealth';
END
GO

-- Backfill UserMenuOptions for users who already have legacy 'UserManagement' right
IF EXISTS (SELECT 1 FROM UserMenuOptions WHERE OptionID = 'UserManagement')
BEGIN
    INSERT INTO UserMenuOptions (UserID, OptionID)
    SELECT DISTINCT umo.UserID, 'SetupUsers'
    FROM UserMenuOptions umo
    WHERE umo.OptionID = 'UserManagement'
      AND NOT EXISTS (SELECT 1 FROM UserMenuOptions existing WHERE existing.UserID = umo.UserID AND existing.OptionID = 'SetupUsers');
    PRINT 'Backfilled SetupUsers permission for users having legacy UserManagement';
END
GO

-- ==============================================================================
-- Deduplicate and Harmonize MenuOptions with UserMenuOptions
-- Realigns Blazor OptionIDs with canonical legacy OptionIDs where user rights reside
-- ==============================================================================

-- 1. Ensure any permissions granted under duplicate OptionIDs are preserved in canonical OptionIDs
INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'CompanyCatalog'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'CmpCompanyCatalog'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'CompanyCatalog');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'CustomerCatalog'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'CmpCustomerCatalog'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'CustomerCatalog');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'OrderItemList'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'ExpOrderItemList'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'OrderItemList');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'CustomInvoice'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'ExpCustomInvoice'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'CustomInvoice');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'NewCustomInvoice'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'ExpNewCustomInvoice'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'NewCustomInvoice');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'PrintValuationForm'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'ExpValuationForm'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'PrintValuationForm');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'StkMaterialGroup'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'StkRMGroups'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'StkMaterialGroup');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'StkVenderBilling'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'StkVendorBilling'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'StkVenderBilling');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'StkVenderBillingList'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'StkVendorBillingList'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'StkVenderBillingList');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'PrlTakeAttendanceEx'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'PayAttendanceManual'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'PrlTakeAttendanceEx');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'PrlAbsentSheet'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'PayAbsentSheet'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'PrlAbsentSheet');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'PrlSocialSecurity'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'PaySocialSecurity'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'PrlSocialSecurity');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'PrlEOBI'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'PayEOBI'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'PrlEOBI');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'PrlPayrollPolicies'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'PayPolicies'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'PrlPayrollPolicies');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'PrdReceivingAgainstPO'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'PrdReceivePO'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'PrdReceivingAgainstPO');

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'AccMakerList'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'PrdMakerList'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'AccMakerList');
GO

-- 2. Remove orphaned duplicate assignments from UserMenuOptions
DELETE FROM UserMenuOptions 
WHERE OptionID IN (
    'CmpCompanyCatalog', 'CmpCustomerCatalog', 'ExpOrderItemList', 'ExpCustomInvoice',
    'ExpNewCustomInvoice', 'ExpValuationForm', 'StkRMGroups', 'StkVendorBilling',
    'StkVendorBillingList', 'PayAttendanceManual', 'PayAbsentSheet', 'PaySocialSecurity',
    'PayEOBI', 'PayPolicies', 'PrdReceivePO', 'PrdMakerList'
);
GO

-- 3. Remove orphaned duplicate rows from MenuOptions
DELETE FROM MenuOptions 
WHERE OptionID IN (
    'CmpCompanyCatalog', 'CmpCustomerCatalog', 'ExpOrderItemList', 'ExpCustomInvoice',
    'ExpNewCustomInvoice', 'ExpValuationForm', 'StkRMGroups', 'StkVendorBilling',
    'StkVendorBillingList', 'PayAttendanceManual', 'PayAbsentSheet', 'PaySocialSecurity',
    'PayEOBI', 'PayPolicies', 'PrdReceivePO', 'PrdMakerList'
);
GO

-- 4. Disambiguate identical display names that represent different screens
UPDATE MenuOptions SET OptionName = 'Commercial Packing List' WHERE OptionID = 'ComPackingList' AND OptionName = 'Packing List';
UPDATE MenuOptions SET OptionName = 'Custom Packing List' WHERE OptionID = 'CustomPackingList' AND OptionName = 'Packing List';
UPDATE MenuOptions SET OptionName = 'Change Locations (Raw Material)' WHERE OptionID = 'StkChangeLocations' AND OptionName = 'Change Locations';
UPDATE MenuOptions SET OptionName = 'Change Locations (Semi-Finished)' WHERE OptionID = 'StkChangeLocationsSF' AND OptionName = 'Change Locations';
GO

-- 5. Add Employees_Prefix to Company table if not exists
IF NOT EXISTS (
    SELECT 1 FROM sys.columns 
    WHERE object_id = OBJECT_ID('dbo.Company') 
      AND name = 'Employees_Prefix'
)
BEGIN
    ALTER TABLE [dbo].[Company] ADD [Employees_Prefix] VARCHAR(10) NULL;
    PRINT 'Added Employees_Prefix column to Company table.';
END
GO

-- 6. Clean up invalid/empty department records
DELETE FROM Departments WHERE LTRIM(RTRIM(ISNULL(deptid, ''))) = '';
GO

-- 7. Ensure GetNextAccno function handles alphabetical and numeric account prefixes safely
CREATE OR ALTER FUNCTION [dbo].[GetNextAccno](@AccountName AS VARCHAR(255),@AccType AS VARCHAR(50),@ParentAccount As Varchar(255)=NULL,@IsParent As BIT=0)
RETURNS VARCHAR(255) AS  
BEGIN
DECLARE @AccNo AS VARCHAR(255)
DECLARE @NewVal AS VARCHAR(255), @Prefix As VARCHAR(50)

IF @IsParent=1
	BEGIN
		SET @NewVal=(SELECT MAX(CAST(RIGHT(AccNo,3) AS INT)) from Accounts WHERE SubAccOf=@ParentAccount)
		SET @NewVal=CAST(ISNULL(@NewVal,'') AS INT)+1
		SET @AccNo=@ParentAccount + '-' + REPLICATE('0',3-LEN(@NewVal)) + @NewVal
	END
ELSE
	BEGIN
		SET @Prefix=ASCII(UPPER(LEFT(LTRIM(RTRIM(ISNULL(@AccountName,''))),1)))
		IF @Prefix>=65 AND @Prefix<=90
			SET @Prefix=REPLICATE('0',2-LEN(@Prefix-64)) + CAST(@Prefix-64 AS VARCHAR)
		ELSE 
			SET @Prefix='00'
		
		SET @NewVal=(SELECT MAX(CAST(RIGHT(AccNo,3) AS INT)) FROM Accounts WHERE SubAccOf=@ParentAccount AND SUBSTRING(AccNo,LEN(Accno)-4,2)=@Prefix)
		SET @NewVal=CAST(ISNULL(@NewVal,'') AS INT)+1
		SET @AccNo=@ParentAccount + '-' + @Prefix + REPLICATE('0',3-LEN(@NewVal)) + @NewVal
	END
RETURN @AccNo
END
GO

-- ============================================================================
-- Maker Receiving List: VRD_Mark_Move_To_Store Table
-- Tracks receiving records marked for transfer / move to store
-- ============================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'VRD_Mark_Move_To_Store')
BEGIN
    CREATE TABLE dbo.VRD_Mark_Move_To_Store (
        VRD_EntryID INT NOT NULL PRIMARY KEY,
        UserName VARCHAR(50) NULL,
        MachineName VARCHAR(50) NULL,
        DTEntry DATETIME NULL DEFAULT (GETDATE())
    );
END
GO

-- ============================================================================
-- Web Push Notifications: UserPushSubscriptions Table
-- Stores browser and mobile push notification device subscriptions
-- ============================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'UserPushSubscriptions')
BEGIN
    CREATE TABLE dbo.UserPushSubscriptions (
        Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        UserID INT NULL,
        UserName VARCHAR(50) NOT NULL,
        Endpoint VARCHAR(MAX) NOT NULL,
        P256dh VARCHAR(255) NOT NULL,
        Auth VARCHAR(255) NOT NULL,
        DeviceName VARCHAR(100) NULL,
        CreatedAt DATETIME NOT NULL DEFAULT (GETDATE()),
        LastUsedAt DATETIME NULL
    );
    CREATE NONCLUSTERED INDEX IX_UserPushSubscriptions_UserName ON dbo.UserPushSubscriptions (UserName);
END
GO

-- ============================================================================
-- Workflow Tasks & Governance Engine Tables
-- Used by: Voucher Deletion, Attendance Modifications/Approvals, Order Authorizations
-- ============================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'TaskItems')
BEGIN
    CREATE TABLE dbo.TaskItems (
        Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        Title NVARCHAR(400) NOT NULL,
        Description NVARCHAR(MAX) NULL,
        AssignedTo NVARCHAR(100) NULL,
        AssignedBy NVARCHAR(100) NOT NULL,
        DepartmentId VARCHAR(50) NULL,
        Priority INT NOT NULL DEFAULT (1),
        Status INT NOT NULL DEFAULT (0),
        DueDate DATETIME2(7) NULL,
        WhatsAppMessageSent BIT NOT NULL DEFAULT (0),
        CreatedAt DATETIME2(7) NOT NULL DEFAULT (GETUTCDATE()),
        UpdatedAt DATETIME2(7) NULL,
        CompletedAt DATETIME2(7) NULL,
        AdditionalAssigneeIds NVARCHAR(MAX) NULL,
        AssignedToNames NVARCHAR(MAX) NULL,
        EmailMessageSent BIT NOT NULL DEFAULT (0),
        IsRead BIT NOT NULL DEFAULT (0),
        StartedAt DATETIME2(7) NULL,
        SourceEntityType NVARCHAR(50) NULL,
        SourceEntityRefId NVARCHAR(100) NULL,
        TargetRole NVARCHAR(50) NULL,
        CompletedBy NVARCHAR(100) NULL,
        ActionUrl NVARCHAR(300) NULL,
        DueWarningSent BIT NOT NULL DEFAULT (0),
        OverdueWarningSent BIT NOT NULL DEFAULT (0)
    );
    CREATE NONCLUSTERED INDEX IX_TaskItems_AssignedTo ON dbo.TaskItems (AssignedTo ASC);
    CREATE NONCLUSTERED INDEX IX_TaskItems_AssignedBy ON dbo.TaskItems (AssignedBy ASC);
    CREATE NONCLUSTERED INDEX IX_TaskItems_Status ON dbo.TaskItems (Status ASC);
    CREATE NONCLUSTERED INDEX IX_TaskItems_DepartmentId ON dbo.TaskItems (DepartmentId ASC);
    CREATE NONCLUSTERED INDEX IX_TaskItems_SourceEntity ON dbo.TaskItems (SourceEntityType ASC, SourceEntityRefId ASC);
    CREATE NONCLUSTERED INDEX IX_TaskItems_TargetRole_Status ON dbo.TaskItems (TargetRole ASC, Status ASC);
END
GO

-- Ensure workflow governance columns exist if TaskItems was created earlier
IF EXISTS (SELECT 1 FROM sys.tables WHERE name = 'TaskItems')
BEGIN
    IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TaskItems') AND name = 'SourceEntityType')
        ALTER TABLE dbo.TaskItems ADD SourceEntityType NVARCHAR(50) NULL;

    IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TaskItems') AND name = 'SourceEntityRefId')
        ALTER TABLE dbo.TaskItems ADD SourceEntityRefId NVARCHAR(100) NULL;

    IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TaskItems') AND name = 'TargetRole')
        ALTER TABLE dbo.TaskItems ADD TargetRole NVARCHAR(50) NULL;

    IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TaskItems') AND name = 'ActionUrl')
        ALTER TABLE dbo.TaskItems ADD ActionUrl NVARCHAR(300) NULL;

    IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.TaskItems') AND name = 'CompletedBy')
        ALTER TABLE dbo.TaskItems ADD CompletedBy NVARCHAR(100) NULL;
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'Task_Assignees')
BEGIN
    CREATE TABLE dbo.Task_Assignees (
        TaskID INT NOT NULL,
        UserID INT NOT NULL,
        UserName NVARCHAR(100) NOT NULL,
        AssignedAt DATETIME2(7) NOT NULL DEFAULT (GETDATE()),
        CONSTRAINT PK_Task_Assignees PRIMARY KEY (TaskID ASC, UserID ASC),
        CONSTRAINT FK_Task_Assignees_TaskItems FOREIGN KEY (TaskID) REFERENCES dbo.TaskItems (Id) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX IX_Task_Assignees_User ON dbo.Task_Assignees (UserID ASC, UserName ASC);
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'Task_Roles')
BEGIN
    CREATE TABLE dbo.Task_Roles (
        TaskID INT NOT NULL,
        RoleName NVARCHAR(50) NOT NULL,
        AssignedAt DATETIME2(7) NOT NULL DEFAULT (GETDATE()),
        CONSTRAINT PK_Task_Roles PRIMARY KEY (TaskID ASC, RoleName ASC),
        CONSTRAINT FK_Task_Roles_TaskItems FOREIGN KEY (TaskID) REFERENCES dbo.TaskItems (Id) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX IX_Task_Roles_Role ON dbo.Task_Roles (RoleName ASC);
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'TaskComments')
BEGIN
    CREATE TABLE dbo.TaskComments (
        Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        TaskId INT NOT NULL,
        UserId NVARCHAR(100) NOT NULL,
        Content NVARCHAR(MAX) NOT NULL,
        CreatedAt DATETIME2(7) NOT NULL DEFAULT (GETUTCDATE()),
        CONSTRAINT FK_TaskComments_TaskItems FOREIGN KEY (TaskId) REFERENCES dbo.TaskItems (Id) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX IX_TaskComments_TaskId ON dbo.TaskComments (TaskId ASC);
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'TaskAttachments')
BEGIN
    CREATE TABLE dbo.TaskAttachments (
        Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        TaskId INT NOT NULL,
        FileName NVARCHAR(510) NOT NULL,
        FilePath NVARCHAR(1000) NOT NULL,
        FileSize BIGINT NOT NULL DEFAULT (0),
        ContentType NVARCHAR(200) NULL,
        UploadedAt DATETIME2(7) NOT NULL DEFAULT (GETUTCDATE()),
        CONSTRAINT FK_TaskAttachments_TaskItems FOREIGN KEY (TaskId) REFERENCES dbo.TaskItems (Id) ON DELETE CASCADE
    );
    CREATE NONCLUSTERED INDEX IX_TaskAttachments_TaskId ON dbo.TaskAttachments (TaskId ASC);
END
GO

-- ============================================================================
-- MonthlySalaries: Index for ultra-fast finalized salary month-locking checks
-- ============================================================================
IF EXISTS (SELECT 1 FROM sys.tables WHERE name = 'MonthlySalaries')
BEGIN
    IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_MonthlySalaries_DT' AND object_id = OBJECT_ID('dbo.MonthlySalaries'))
    BEGIN
        CREATE NONCLUSTERED INDEX IX_MonthlySalaries_DT ON dbo.MonthlySalaries (DT);
    END
END
GO

-- ============================================================================
-- Workflow Configurations & Governance Framework
-- ============================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'WorkflowConfigurations')
BEGIN
    CREATE TABLE [dbo].[WorkflowConfigurations] (
        [WorkflowCode]   VARCHAR(50)   NOT NULL,
        [WorkflowName]   VARCHAR(100)  NOT NULL,
        [Module]         VARCHAR(50)   NOT NULL,
        [Description]    VARCHAR(255)  NULL,
        [IsEnabled]      BIT           NOT NULL CONSTRAINT [DF_WFConfig_IsEnabled] DEFAULT (1),
        [ApproverRole]   VARCHAR(50)   NOT NULL CONSTRAINT [DF_WFConfig_ApproverRole] DEFAULT ('Director'),
        [UpdatedDate]    DATETIME      NOT NULL CONSTRAINT [DF_WFConfig_UpdatedDate] DEFAULT (GETDATE()),
        [UpdatedBy]      VARCHAR(50)   NULL,

        CONSTRAINT [PK_WorkflowConfigurations] PRIMARY KEY CLUSTERED ([WorkflowCode] ASC)
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'WorkflowExemptRoles')
BEGIN
    CREATE TABLE [dbo].[WorkflowExemptRoles] (
        [ID]           INT          IDENTITY(1,1) NOT NULL,
        [WorkflowCode] VARCHAR(50)  NOT NULL,
        [RoleName]     VARCHAR(50)  NOT NULL,

        CONSTRAINT [PK_WorkflowExemptRoles] PRIMARY KEY CLUSTERED ([ID] ASC),
        CONSTRAINT [FK_WorkflowExemptRoles_Config] FOREIGN KEY ([WorkflowCode]) 
            REFERENCES [dbo].[WorkflowConfigurations] ([WorkflowCode]) ON DELETE CASCADE,
        CONSTRAINT [UQ_WorkflowExemptRoles_Code_Role] UNIQUE ([WorkflowCode], [RoleName])
    );
END
GO

-- Seed Master Workflow Configurations
IF NOT EXISTS (SELECT 1 FROM [dbo].[WorkflowConfigurations] WHERE [WorkflowCode] = 'AttendanceApproval')
    INSERT INTO [dbo].[WorkflowConfigurations] ([WorkflowCode], [WorkflowName], [Module], [Description], [IsEnabled], [ApproverRole])
    VALUES ('AttendanceApproval', 'Attendance Edit / Delete Approval', 'Payroll', 'Requires Director approval when non-exempt users modify or delete attendance records.', 1, 'Director');

IF NOT EXISTS (SELECT 1 FROM [dbo].[WorkflowConfigurations] WHERE [WorkflowCode] = 'MakerRateChange')
    INSERT INTO [dbo].[WorkflowConfigurations] ([WorkflowCode], [WorkflowName], [Module], [Description], [IsEnabled], [ApproverRole])
    VALUES ('MakerRateChange', 'Maker Item Rate Change Approval', 'Production', 'Requires Director approval when non-exempt users edit maker assignment rates.', 1, 'Director');

IF NOT EXISTS (SELECT 1 FROM [dbo].[WorkflowConfigurations] WHERE [WorkflowCode] = 'DeleteLotAuthorization')
    INSERT INTO [dbo].[WorkflowConfigurations] ([WorkflowCode], [WorkflowName], [Module], [Description], [IsEnabled], [ApproverRole])
    VALUES ('DeleteLotAuthorization', 'Delete Lot Authorization Approval', 'Production', 'Requires Director approval when non-exempt users request to un-authorize or delete a lot authorization.', 1, 'Director');

IF NOT EXISTS (SELECT 1 FROM [dbo].[WorkflowConfigurations] WHERE [WorkflowCode] = 'DeleteProductionIssuance')
    INSERT INTO [dbo].[WorkflowConfigurations] ([WorkflowCode], [WorkflowName], [Module], [Description], [IsEnabled], [ApproverRole])
    VALUES ('DeleteProductionIssuance', 'Delete Production Issuance Approval', 'Production', 'Requires Director approval when non-exempt users request deletion of an issued Lot or Master PO.', 1, 'Director');

IF NOT EXISTS (SELECT 1 FROM [dbo].[WorkflowConfigurations] WHERE [WorkflowCode] = 'SkipProcess')
    INSERT INTO [dbo].[WorkflowConfigurations] ([WorkflowCode], [WorkflowName], [Module], [Description], [IsEnabled], [ApproverRole])
    VALUES ('SkipProcess', 'Skip Production Process Approval', 'Production', 'Requires Director approval when non-exempt users request to skip a sequential production process on a lot item.', 1, 'Director');
GO

-- Seed Default Exempt Roles (Direct-Save without approval)
DECLARE @ExemptList TABLE (Code VARCHAR(50), RoleName VARCHAR(50));
INSERT INTO @ExemptList (Code, RoleName) VALUES
    ('AttendanceApproval', 'Director'),
    ('AttendanceApproval', 'Admin'),
    ('AttendanceApproval', 'Administrator'),

    ('MakerRateChange', 'Director'),
    ('MakerRateChange', 'Admin'),
    ('MakerRateChange', 'Administrator'),

    ('DeleteLotAuthorization', 'Director'),
    ('DeleteLotAuthorization', 'Admin'),
    ('DeleteLotAuthorization', 'Administrator'),

    ('DeleteProductionIssuance', 'Director'),
    ('DeleteProductionIssuance', 'Admin'),
    ('DeleteProductionIssuance', 'Administrator'),

    ('SkipProcess', 'Director'),
    ('SkipProcess', 'Admin'),
    ('SkipProcess', 'Administrator');

INSERT INTO [dbo].[WorkflowExemptRoles] ([WorkflowCode], [RoleName])
SELECT e.Code, e.RoleName
FROM @ExemptList e
WHERE NOT EXISTS (
    SELECT 1 FROM [dbo].[WorkflowExemptRoles] r 
    WHERE r.WorkflowCode = e.Code AND r.RoleName = e.RoleName
);
GO

-- ====================================================================================================
-- SCRIPT: Production Deletion Requests Table & Lock Check Function
-- ====================================================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'ProductionDeletionRequests')
BEGIN
    CREATE TABLE [dbo].[ProductionDeletionRequests] (
        [Id] INT IDENTITY(1,1) PRIMARY KEY,
        [RequestType] VARCHAR(50) NOT NULL, -- 'LotReceiving', 'LotIssuance', 'MasterPOIssuance', 'SkipProcess'
        [EntityRefID] BIGINT NOT NULL,       -- VRD_EntryID, Issuance EntryID, etc.
        [LotNo] VARCHAR(50) NOT NULL,
        [OrderNo] VARCHAR(50) NULL,
        [ItemCode] VARCHAR(50) NULL,
        [ItemName] NVARCHAR(200) NULL,
        [ProcessID] INT NULL,
        [ProcessName] NVARCHAR(100) NULL,
        [MakerID] BIGINT NULL,
        [MakerName] NVARCHAR(150) NULL,
        [Qty] NUMERIC(18,2) NOT NULL DEFAULT 0,
        [RequestedBy] VARCHAR(50) NOT NULL,
        [RequestedDT] DATETIME NOT NULL DEFAULT GETDATE(),
        [Reason] NVARCHAR(500) NOT NULL,
        [MachineName] VARCHAR(100) NULL,
        [Status] VARCHAR(20) NOT NULL DEFAULT 'Pending', -- 'Pending', 'Approved', 'Rejected', 'Cancelled'
        [ReviewedBy] VARCHAR(50) NULL,
        [ReviewedDT] DATETIME NULL,
        [DirectorRemarks] NVARCHAR(500) NULL,
        [TaskId] INT NULL
    );

    CREATE INDEX IX_ProdDelReq_Active ON [dbo].[ProductionDeletionRequests] ([EntityRefID], [Status]);
    CREATE INDEX IX_ProdDelReq_LotNo ON [dbo].[ProductionDeletionRequests] ([LotNo], [Status]);
    CREATE INDEX IX_ProdDelReq_TypeStatus ON [dbo].[ProductionDeletionRequests] ([RequestType], [Status]);
END
GO

CREATE OR ALTER FUNCTION [dbo].[fn_IsLotPendingDeletion] (@LotNo VARCHAR(50), @VRD_EntryID BIGINT = 0)
RETURNS BIT
AS
BEGIN
    IF EXISTS (
        SELECT 1 FROM dbo.ProductionDeletionRequests WITH (NOLOCK)
        WHERE Status = 'Pending'
          AND (
              (@VRD_EntryID > 0 AND EntityRefID = @VRD_EntryID)
              OR (@VRD_EntryID = 0 AND LotNo = @LotNo)
          )
    )
        RETURN 1;
    RETURN 0;
END
GO

