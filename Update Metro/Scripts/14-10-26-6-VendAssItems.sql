IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('VendAssItems') AND name = 'CustCode')
   AND NOT EXISTS (SELECT 1 FROM sys.default_constraints WHERE parent_object_id = OBJECT_ID('VendAssItems') AND col_name(parent_object_id, parent_column_id) = 'CustCode')
BEGIN
    ALTER TABLE [dbo].[VendAssItems] ADD CONSTRAINT [DF_VendAssItems_CustCode] DEFAULT ('Stock') FOR [CustCode];
END
GO