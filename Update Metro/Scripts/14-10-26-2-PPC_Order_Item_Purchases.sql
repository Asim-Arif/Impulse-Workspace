IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.[PPC_Order_Item_Purchases]') AND name = 'ExcessQtyPercentage')
BEGIN
    ALTER TABLE [dbo].[PPC_Order_Item_Purchases] ADD [ExcessQtyPercentage] [real] NOT NULL DEFAULT ((50));
    PRINT 'Added column [ExcessQtyPercentage] to [dbo].[PPC_Order_Item_Purchases]';
END