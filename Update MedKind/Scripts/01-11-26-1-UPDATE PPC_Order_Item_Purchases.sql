IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('PPC_Order_Item_Purchases') AND name = 'MasterPONo')
BEGIN
    UPDATE PPC_Order_Item_Purchases 
    SET VendIssued_RefID = NULL, MasterPONo = NULL, IssuedAt = NULL, IssuedBy = NULL 
    WHERE MasterPONo LIKE 'M-ISU%' OR MasterPONo LIKE 'ISU%';
END
GO