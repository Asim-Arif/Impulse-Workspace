-- ============================================================================
-- Setup Module Security & Screen Rights: SetupMainLink and Setup MenuOptions
-- ============================================================================
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='Users' AND COLUMN_NAME='SetupMainLink')
BEGIN
    ALTER TABLE dbo.Users ADD SetupMainLink BIT NOT NULL CONSTRAINT DF_Users_SetupMainLink DEFAULT (0);
    PRINT 'Added SetupMainLink to Users table';
END
GO

IF EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='Users' AND COLUMN_NAME='SetupMainLink')
BEGIN
    EXEC sp_executesql N'UPDATE dbo.Users SET SetupMainLink = 1 WHERE UserManagement = 1;';
    PRINT 'Updated SetupMainLink = 1 for users with UserManagement = 1';
END
GO

-- Ensure Setup Menu Options exist in MenuOptions table
IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'SetupHub')
BEGIN
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName)
    VALUES ('SetupHub', 'Setups Hub', 'Setup', 'SetupDashboard');
END
ELSE
BEGIN
    UPDATE MenuOptions SET ModuleName = 'Setup', OptionName = 'Setups Hub', FormName = 'SetupDashboard' WHERE OptionID = 'SetupHub';
END
GO

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'SetupUsers')
BEGIN
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName)
    VALUES ('SetupUsers', 'User Management', 'Setup', 'UsersList');
END
ELSE
BEGIN
    UPDATE MenuOptions SET ModuleName = 'Setup', OptionName = 'User Management', FormName = 'UsersList' WHERE OptionID = 'SetupUsers';
END
GO

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'OfficeMinuteTypes')
BEGIN
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName)
    VALUES ('OfficeMinuteTypes', 'Minute Types', 'Setup', 'MinuteTypesAdmin');
END
ELSE
BEGIN
    UPDATE MenuOptions SET ModuleName = 'Setup', OptionName = 'Minute Types', FormName = 'MinuteTypesAdmin' WHERE OptionID = 'OfficeMinuteTypes';
END
GO

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'OfficeEmailSettings')
BEGIN
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName)
    VALUES ('OfficeEmailSettings', 'Email & SMTP Settings', 'Setup', 'EmailSettingsAdmin');
END
ELSE
BEGIN
    UPDATE MenuOptions SET ModuleName = 'Setup', OptionName = 'Email & SMTP Settings', FormName = 'EmailSettingsAdmin' WHERE OptionID = 'OfficeEmailSettings';
END
GO

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'IntraOfficeHealth')
BEGIN
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName)
    VALUES ('IntraOfficeHealth', 'System Diagnostics', 'Setup', 'SystemHealth');
END
ELSE
BEGIN
    UPDATE MenuOptions SET ModuleName = 'Setup', OptionName = 'System Diagnostics', FormName = 'SystemHealth' WHERE OptionID = 'IntraOfficeHealth';
END
GO

-- Backfill UserMenuOptions for users who already have legacy 'UserManagement' right
IF EXISTS (SELECT 1 FROM UserMenuOptions WHERE OptionID = 'UserManagement')
BEGIN
    INSERT INTO UserMenuOptions (UserID, OptionID)
    SELECT DISTINCT umo.UserID, 'SetupUsers'
    FROM UserMenuOptions umo
    WHERE umo.OptionID = 'UserManagement'
      AND NOT EXISTS (SELECT 1 FROM UserMenuOptions existing WHERE existing.UserID = umo.UserID AND existing.OptionID = 'SetupUsers');
    PRINT 'Backfilled SetupUsers permission for users having legacy UserManagement';
END