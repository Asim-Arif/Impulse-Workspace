IF NOT EXISTS (SELECT 1 FROM dbo.GeneralData WHERE DataName = 'MakerBilling_IncludeRejectedQty')
BEGIN
    INSERT INTO dbo.GeneralData (DataName, DataValue, PasswordEncrypted)
    VALUES ('MakerBilling_IncludeRejectedQty', 'False', 0);
END