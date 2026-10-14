IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'License_Keys' AND schema_id = SCHEMA_ID('dbo'))
BEGIN
    CREATE TABLE [dbo].[License_Keys](
        [EntryID] [int] IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [LicenseKey] [varchar](max) NULL,
        [MachineName] [varchar](100) NULL,
        [CreatedAt] [datetime] DEFAULT GETUTCDATE()
    );
    PRINT 'Created table [dbo].[License_Keys]';
END
ELSE
BEGIN
    -- Ensure LicenseKey column is varchar(max)
    IF EXISTS (
        SELECT 1 FROM sys.columns 
        WHERE object_id = OBJECT_ID('dbo.License_Keys') 
          AND name = 'LicenseKey' 
          AND max_length <> -1
    )
    BEGIN
        ALTER TABLE [dbo].[License_Keys] ALTER COLUMN [LicenseKey] [varchar](max) NULL;
        PRINT 'Altered [dbo].[License_Keys].[LicenseKey] to varchar(max)';
    END
END