IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'VRD_Mark_Move_To_Store')
BEGIN
    CREATE TABLE dbo.VRD_Mark_Move_To_Store (
        VRD_EntryID INT NOT NULL PRIMARY KEY,
        UserName VARCHAR(50) NULL,
        MachineName VARCHAR(50) NULL,
        DTEntry DATETIME NULL DEFAULT (GETDATE())
    );
END
GO