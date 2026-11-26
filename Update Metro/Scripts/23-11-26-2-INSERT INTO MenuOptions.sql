IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'PrdMakerPOList')
BEGIN
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName)
    VALUES ('PrdMakerPOList', 'Issuance List', 'Production', 'MakerPOList');
END
ELSE
BEGIN
    UPDATE MenuOptions SET OptionName = 'Issuance List', ModuleName = 'Production' WHERE OptionID = 'PrdMakerPOList';
END
GO

INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT umo.UserID, 'PrdMakerPOList'
FROM UserMenuOptions umo
WHERE umo.OptionID = 'PrdIssuanceList'
  AND NOT EXISTS (SELECT 1 FROM UserMenuOptions e WHERE e.UserID = umo.UserID AND e.OptionID = 'PrdMakerPOList');

DELETE FROM UserMenuOptions WHERE OptionID = 'PrdIssuanceList';
DELETE FROM MenuOptions WHERE OptionID = 'PrdIssuanceList';
GO