IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('VendAssItems') AND name = 'CustCode')
BEGIN
    ALTER TABLE [dbo].[VendAssItems] ADD [CustCode] VARCHAR(50) NULL;
END
GO