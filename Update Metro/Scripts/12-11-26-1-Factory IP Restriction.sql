-- -------------------------------------------------------------------------------------------------
-- Factory IP Restriction & Remote Access Management
-- -------------------------------------------------------------------------------------------------
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'Users' AND COLUMN_NAME = 'AllowOutsideAccess')
BEGIN
    ALTER TABLE dbo.Users ADD AllowOutsideAccess BIT NOT NULL CONSTRAINT DF_Users_AllowOutsideAccess DEFAULT 0;
END
GO

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'RemoteAccessRequests')
BEGIN
    CREATE TABLE dbo.RemoteAccessRequests (
        EntryID INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        UserID INT NOT NULL,
        UserName VARCHAR(50) NOT NULL,
        ClientIP VARCHAR(50) NOT NULL,
        RequestDT DATETIME NOT NULL DEFAULT GETDATE(),
        Reason NVARCHAR(255) NULL,
        Status VARCHAR(20) NOT NULL DEFAULT 'Pending', -- 'Pending', 'Approved', 'Rejected', 'Expired'
        DurationMinutes INT NULL,
        ExpiresAt DATETIME NULL,
        ActionDT DATETIME NULL,
        ActionByUserID INT NULL,
        ActionByUserName VARCHAR(50) NULL,
        Remarks NVARCHAR(255) NULL
    );

    CREATE INDEX IX_RemoteAccessRequests_User_Status 
        ON dbo.RemoteAccessRequests (UserID, Status, RequestDT DESC);
    CREATE INDEX IX_RemoteAccessRequests_Expires 
        ON dbo.RemoteAccessRequests (Status, ExpiresAt);
END