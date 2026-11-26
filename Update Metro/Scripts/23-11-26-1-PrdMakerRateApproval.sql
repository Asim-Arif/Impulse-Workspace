IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'PrdMakerRateApproval')
BEGIN
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName)
    VALUES ('PrdMakerRateApproval', 'Maker Rate Approvals', 'Production', 'MakerRateApproval');
END
ELSE
BEGIN
    UPDATE MenuOptions
    SET OptionName = 'Maker Rate Approvals', ModuleName = 'Production', FormName = 'MakerRateApproval'
    WHERE OptionID = 'PrdMakerRateApproval';
END