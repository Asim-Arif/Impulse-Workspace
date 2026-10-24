-- ====================================================================================================
-- SCRIPT: Production Approval Options in MenuOptions Catalog
-- ====================================================================================================
IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'PrdIssuanceDeletionApproval')
BEGIN
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName)
    VALUES ('PrdIssuanceDeletionApproval', 'Issuance Deletion Approvals', 'Production', 'IssuanceDeletionApproval');
END
ELSE
BEGIN
    UPDATE MenuOptions SET OptionName = 'Issuance Deletion Approvals', ModuleName = 'Production', FormName = 'IssuanceDeletionApproval' WHERE OptionID = 'PrdIssuanceDeletionApproval';
END
GO

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'PrdSkipProcessApproval')
BEGIN
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName)
    VALUES ('PrdSkipProcessApproval', 'Skip Process Approvals', 'Production', 'SkipProcessApproval');
END
ELSE
BEGIN
    UPDATE MenuOptions SET OptionName = 'Skip Process Approvals', ModuleName = 'Production', FormName = 'SkipProcessApproval' WHERE OptionID = 'PrdSkipProcessApproval';
END
GO

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'PrdLotDeletionApproval')
BEGIN
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName)
    VALUES ('PrdLotDeletionApproval', 'Lot Deletion Approvals', 'Production', 'LotDeletionApproval');
END
ELSE
BEGIN
    UPDATE MenuOptions SET OptionName = 'Lot Deletion Approvals', ModuleName = 'Production', FormName = 'LotDeletionApproval' WHERE OptionID = 'PrdLotDeletionApproval';
END
GO
