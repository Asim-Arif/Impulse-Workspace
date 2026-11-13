-- ==============================================================================
-- PPC Order Planning: Exception Handling, Rejection Tracking & PO Short-Close
-- ==============================================================================
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PPC_Order_Item_Planning') AND name = 'RejectedStockQty')
BEGIN
    ALTER TABLE PPC_Order_Item_Planning ADD RejectedStockQty INT NOT NULL CONSTRAINT DF_PPC_ItemPlan_RejectedStockQty DEFAULT 0;
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PPC_Order_Item_Planning') AND name = 'RejectedSFQty')
BEGIN
    ALTER TABLE PPC_Order_Item_Planning ADD RejectedSFQty INT NOT NULL CONSTRAINT DF_PPC_ItemPlan_RejectedSFQty DEFAULT 0;
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PPC_Order_Item_Purchases') AND name = 'Status')
BEGIN
    ALTER TABLE PPC_Order_Item_Purchases ADD Status VARCHAR(30) NOT NULL CONSTRAINT DF_PPC_ItemPurch_Status DEFAULT 'Active';
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PPC_Order_Item_Purchases') AND name = 'OriginalPlannedQty')
BEGIN
    ALTER TABLE PPC_Order_Item_Purchases ADD OriginalPlannedQty INT NULL;
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PPC_Order_Item_Purchases') AND name = 'ShortClosedQty')
BEGIN
    ALTER TABLE PPC_Order_Item_Purchases ADD ShortClosedQty INT NOT NULL CONSTRAINT DF_PPC_ItemPurch_ShortClosedQty DEFAULT 0;
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PPC_Order_Item_Purchases') AND name = 'ExceptionNotes')
BEGIN
    ALTER TABLE PPC_Order_Item_Purchases ADD ExceptionNotes NVARCHAR(500) NULL;
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PPC_Order_Item_Purchases') AND name = 'UpdatedAt')
BEGIN
    ALTER TABLE PPC_Order_Item_Purchases ADD UpdatedAt DATETIME2 NULL;
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PPC_Order_Item_Purchases') AND name = 'UpdatedBy')
BEGIN
    ALTER TABLE PPC_Order_Item_Purchases ADD UpdatedBy NVARCHAR(100) NULL;
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'PPC_Order_Planning_Exceptions')
BEGIN
    CREATE TABLE dbo.PPC_Order_Planning_Exceptions (
        EntryID INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        OrderNo VARCHAR(100) NOT NULL,
        ItemID VARCHAR(50) NOT NULL,
        ExceptionType VARCHAR(50) NOT NULL, -- 'FinishStockRelease', 'FinishStockRejection', 'SFStockRelease', 'SFStockRejection', 'CancelMakerPO', 'ShortCloseMakerPO', 'UnlockPlan'
        AffectedQty INT NOT NULL DEFAULT 0,
        SourceRef VARCHAR(100) NULL,
        Reason NVARCHAR(500) NOT NULL,
        CreatedBy NVARCHAR(100) NOT NULL,
        CreatedAt DATETIME2 NOT NULL DEFAULT GETUTCDATE()
    );

    CREATE INDEX IX_PPC_Exceptions_Order_Item ON dbo.PPC_Order_Planning_Exceptions(OrderNo, ItemID, CreatedAt DESC);
END
GO
