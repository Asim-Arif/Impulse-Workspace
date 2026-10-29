
-- ==============================================================================
-- PPC Order Planning: Semi-Finished (SF) Stock Support
-- ==============================================================================

IF NOT EXISTS (
    SELECT 1 FROM sys.columns 
    WHERE object_id = OBJECT_ID('PPC_Order_Item_Planning') AND name = 'SFStockQty'
)
BEGIN
    ALTER TABLE PPC_Order_Item_Planning ADD SFStockQty INT NOT NULL CONSTRAINT DF_PPC_SFStockQty DEFAULT 0;
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'PPC_Order_Item_SF_Allocations')
BEGIN
    CREATE TABLE PPC_Order_Item_SF_Allocations (
        EntryID INT IDENTITY(1,1) PRIMARY KEY,
        OrderNo VARCHAR(100) NOT NULL,
        ItemID VARCHAR(50) NOT NULL,
        SOO_RefID INT NOT NULL,
        ProcessID INT NOT NULL,
        ProcessName NVARCHAR(100) NULL,
        LotNo VARCHAR(100) NULL,
        BatchNo VARCHAR(100) NULL,
        Shelf_RefID INT NULL,
        StoreLocation NVARCHAR(200) NULL,
        AllocatedQty INT NOT NULL,
        CreatedAt DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        CreatedBy NVARCHAR(100) NULL
    );

    CREATE INDEX IX_PPC_SF_Alloc_Order_Item ON PPC_Order_Item_SF_Allocations(OrderNo, ItemID);
END
GO