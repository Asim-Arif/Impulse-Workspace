-- Synchronize Default Menu Options for Accounts user (UserID where UserName = 'accounts')
INSERT INTO UserMenuOptions (UserID, OptionID)
SELECT u.UserID, mo.OptionID
FROM Users u
CROSS JOIN MenuOptions mo
WHERE u.UserName = 'accounts'
  AND mo.ModuleName IN ('Budgeting', 'FixedAssets', 'GateManagement')
  AND NOT EXISTS (
      SELECT 1 FROM UserMenuOptions umo WHERE umo.UserID = u.UserID AND umo.OptionID = mo.OptionID
  );
GO
