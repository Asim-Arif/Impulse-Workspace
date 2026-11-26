IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'PrdAuthRejection')
BEGIN
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName)
    VALUES ('PrdAuthRejection', 'Auth. Rejection', 'Production', 'RejectionAuthorization');
END
ELSE
BEGIN
    UPDATE MenuOptions
    SET OptionName = 'Auth. Rejection',
        ModuleName = 'Production',
        FormName = 'RejectionAuthorization'
    WHERE OptionID = 'PrdAuthRejection';
END
GO