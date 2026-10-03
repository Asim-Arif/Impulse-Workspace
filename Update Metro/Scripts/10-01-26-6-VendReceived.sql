-- Index on VendReceived Date and foreign keys
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_VendReceived_DT' AND object_id = OBJECT_ID('dbo.VendReceived'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_VendReceived_DT]
    ON [dbo].[VendReceived] ([DT])
    INCLUDE ([EntryID], [VendID], [ProcessID], [Issuance_RefID], [EmpID], [UserID]);
END
GO

-- Index on vendRcvdDetail RefID and ItemCode
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_vendRcvdDetail_RefID_ItemCode' AND object_id = OBJECT_ID('dbo.vendRcvdDetail'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_vendRcvdDetail_RefID_ItemCode]
    ON [dbo].[vendRcvdDetail] ([RefID], [ItemCode])
    INCLUDE ([EntryID], [RecieptID], [LotNo], [OrderNo], [RcvdQty], [Wastage], [LostQty], [ReWorkQty], [ReWorkLot], [IssQty], [Opening_RefID], [Issue_RefID]);
END
GO

-- Index on VendReceived_Employees
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_VendReceived_Employees_VR_RefID' AND object_id = OBJECT_ID('dbo.VendReceived_Employees'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_VendReceived_Employees_VR_RefID]
    ON [dbo].[VendReceived_Employees] ([VR_RefID], [EmpID]);
END
GO
