
-- ============================================================================
-- Gate Management: Courier_Issuance Table
-- ============================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'Courier_Issuance')
BEGIN
    CREATE TABLE dbo.Courier_Issuance (
        EntryID int IDENTITY(1,1) NOT NULL CONSTRAINT PK_Courier_Issuance PRIMARY KEY,
        DT datetime NULL,
        TrackingNo varchar(50) NULL,
        Description varchar(255) NULL,
        VendorName varchar(50) NULL,
        VendorAddress varchar(255) NULL,
        IssuedBy varchar(50) NULL,
        ConcernedPerson varchar(50) NULL,
        ReceivedBy varchar(50) NULL,
        Self bit NULL CONSTRAINT DF_Courier_Issuance_Self DEFAULT ((1)),
        Paid bit NULL CONSTRAINT DF_Courier_Issuance_Paid DEFAULT ((0)),
        VendID varchar(50) NULL,
        Weight varchar(50) NULL,
        Dimension varchar(50) NULL,
        CourierStatus bit NULL CONSTRAINT DF_Courier_Issuance_CourierStatus DEFAULT ((0)),
        MachineName varchar(50) NULL,
        UserName varchar(50) NULL,
        Parcel_Qty real NULL,
        Vehicle_No varchar(50) NULL
    );
END
GO

IF EXISTS (
    SELECT 1 FROM sys.columns c
    JOIN sys.types t ON c.user_type_id = t.user_type_id
    WHERE c.object_id = OBJECT_ID('Courier_Issuance')
      AND c.name = 'VendID'
      AND t.name <> 'varchar'
)
BEGIN
    ALTER TABLE dbo.Courier_Issuance ALTER COLUMN VendID varchar(50) NULL;
END
GO
