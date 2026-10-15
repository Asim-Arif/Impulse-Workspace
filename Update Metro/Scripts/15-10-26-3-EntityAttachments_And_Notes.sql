-- ====================================================================================================
-- SCRIPT: Universal Entity Attachments & Notes System (Customers, Orders, Invoices, Employees)
-- ====================================================================================================

-- 1. EntityAttachments Table
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'EntityAttachments')
BEGIN
    CREATE TABLE [dbo].[EntityAttachments] (
        [Id]                BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [EntityType]        VARCHAR(50) NOT NULL,          -- 'Customer', 'Order', 'Invoice', 'Employee'
        [EntityId]          VARCHAR(100) NOT NULL,         -- e.g. CustCode, OrderNo, InvoiceNo, EmpID
        [EntitySecondaryId] VARCHAR(100) NULL,             -- e.g. Country (for ForeignCustomers composite key)
        [OriginalFileName]  NVARCHAR(500) NOT NULL,        -- e.g. "Tax_Exemption_Certificate_2026.pdf"
        [StoredFileName]    NVARCHAR(500) NOT NULL,        -- e.g. "guid_Tax_Exemption_Certificate_2026.pdf"
        [FilePath]          NVARCHAR(1000) NOT NULL,       -- e.g. "/uploads/attachments/customer/..."
        [FileSize]          BIGINT NOT NULL DEFAULT 0,     -- Size in bytes
        [FileExtension]     VARCHAR(20) NOT NULL,          -- e.g. ".pdf", ".xlsx", ".dwg"
        [ContentType]       VARCHAR(150) NULL,             -- MIME type
        [Remarks]           NVARCHAR(500) NULL,            -- Optional user description / tag
        [UploadedBy]        NVARCHAR(100) NOT NULL,        -- Current user name
        [UploadedAt]        DATETIME2 NOT NULL DEFAULT GETDATE(),
        [IsDeleted]         BIT NOT NULL DEFAULT 0,
        [DeletedBy]         NVARCHAR(100) NULL,
        [DeletedAt]         DATETIME2 NULL
    );

    CREATE NONCLUSTERED INDEX [IX_EntityAttachments_Lookup] 
        ON [dbo].[EntityAttachments]([EntityType], [EntityId], [EntitySecondaryId], [IsDeleted])
        INCLUDE ([OriginalFileName], [FileSize], [FileExtension], [UploadedAt], [UploadedBy]);
        
    PRINT 'Created table: EntityAttachments';
END
GO

-- 2. EntityNotes Table
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'EntityNotes')
BEGIN
    CREATE TABLE [dbo].[EntityNotes] (
        [Id]                BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [EntityType]        VARCHAR(50) NOT NULL,          -- 'Customer', 'Order', 'Invoice', 'Employee'
        [EntityId]          VARCHAR(100) NOT NULL,         -- e.g. CustCode, OrderNo, InvoiceNo, EmpID
        [EntitySecondaryId] VARCHAR(100) NULL,             -- e.g. Country (for ForeignCustomers composite key)
        [NoteText]          NVARCHAR(MAX) NOT NULL DEFAULT '',
        [LastModifiedBy]    NVARCHAR(100) NOT NULL,
        [LastModifiedAt]    DATETIME2 NOT NULL DEFAULT GETDATE()
    );

    CREATE NONCLUSTERED INDEX [IX_EntityNotes_Lookup] 
        ON [dbo].[EntityNotes]([EntityType], [EntityId], [EntitySecondaryId]);
        
    PRINT 'Created table: EntityNotes';
END
GO
