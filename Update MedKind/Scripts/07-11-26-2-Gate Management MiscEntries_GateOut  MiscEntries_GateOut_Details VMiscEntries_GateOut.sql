-- ============================================================================
-- Gate Management: MiscEntries_GateOut & MiscEntries_GateOut_Details & VMiscEntries_GateOut
-- ============================================================================
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'MiscEntries_GateOut')
BEGIN
    CREATE TABLE dbo.MiscEntries_GateOut (
        EntryID int IDENTITY(1,1) NOT NULL CONSTRAINT PK_MiscEntries_GateOut PRIMARY KEY,
        GatePassNo varchar(50) NULL,
        DT datetime NULL,
        Detail varchar(255) NULL,
        Unit varchar(50) NULL,
        Packing varchar(50) NULL,
        PersonName varchar(255) NULL,
        VehicleNo varchar(50) NULL,
        Description varchar(255) NULL,
        Qty int NULL,
        UserName varchar(50) NULL,
        MachineName varchar(50) NULL,
        DTEntry datetime NULL,
        EmpID varchar(50) NULL,
        GatePassType int NULL,
        DeptID varchar(50) NULL,
        Approved_HOD bit NULL CONSTRAINT DF_MiscEntries_GateOut_Approved_HOD DEFAULT (0),
        Approved bit NULL CONSTRAINT DF_MiscEntries_GateOut_Approved DEFAULT (0),
        Approved_Finance bit NULL CONSTRAINT DF_MiscEntries_GateOut_Approved_Finance DEFAULT (0),
        VendID int NULL,
        VendNameManual varchar(50) NULL,
        VM_AccNo varchar(50) NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.MiscEntries_GateOut') AND name = 'EmpID')
BEGIN
    ALTER TABLE dbo.MiscEntries_GateOut ADD EmpID varchar(50) NULL;
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.MiscEntries_GateOut') AND name = 'GatePassType')
BEGIN
    ALTER TABLE dbo.MiscEntries_GateOut ADD GatePassType int NULL;
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.MiscEntries_GateOut') AND name = 'DeptID')
BEGIN
    ALTER TABLE dbo.MiscEntries_GateOut ADD DeptID varchar(50) NULL;
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.MiscEntries_GateOut') AND name = 'Approved_HOD')
BEGIN
    ALTER TABLE dbo.MiscEntries_GateOut ADD Approved_HOD bit NULL CONSTRAINT DF_MiscEntries_GateOut_Approved_HOD DEFAULT (0);
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.MiscEntries_GateOut') AND name = 'Approved')
BEGIN
    ALTER TABLE dbo.MiscEntries_GateOut ADD Approved bit NULL CONSTRAINT DF_MiscEntries_GateOut_Approved DEFAULT (0);
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.MiscEntries_GateOut') AND name = 'Approved_Finance')
BEGIN
    ALTER TABLE dbo.MiscEntries_GateOut ADD Approved_Finance bit NULL CONSTRAINT DF_MiscEntries_GateOut_Approved_Finance DEFAULT (0);
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.MiscEntries_GateOut') AND name = 'VendID')
BEGIN
    ALTER TABLE dbo.MiscEntries_GateOut ADD VendID int NULL;
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.MiscEntries_GateOut') AND name = 'VendNameManual')
BEGIN
    ALTER TABLE dbo.MiscEntries_GateOut ADD VendNameManual varchar(50) NULL;
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.MiscEntries_GateOut') AND name = 'VM_AccNo')
BEGIN
    ALTER TABLE dbo.MiscEntries_GateOut ADD VM_AccNo varchar(50) NULL;
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'MiscEntries_GateOut_Details')
BEGIN
    CREATE TABLE dbo.MiscEntries_GateOut_Details (
        EntryID int IDENTITY(1,1) NOT NULL CONSTRAINT PK_MiscEntries_GateOut_Details PRIMARY KEY,
        MEGO_RefID int NULL,
        Detail varchar(255) NULL,
        Unit varchar(50) NULL,
        Packing varchar(255) NULL,
        Qty real NULL,
        Rate real NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.MiscEntries_GateOut_Details') AND name = 'Rate')
BEGIN
    ALTER TABLE dbo.MiscEntries_GateOut_Details ADD Rate real NULL;
END
GO

IF EXISTS (
    SELECT 1 FROM sys.columns c
    JOIN sys.types t ON c.user_type_id = t.user_type_id
    WHERE c.object_id = OBJECT_ID('dbo.MiscEntries_GateOut_Details')
      AND c.name = 'Qty'
      AND t.name <> 'real'
)
BEGIN
    ALTER TABLE dbo.MiscEntries_GateOut_Details ALTER COLUMN Qty real NULL;
END
GO

CREATE OR ALTER VIEW dbo.VMiscEntries_GateOut
AS
SELECT 
    m.EntryID, 
    m.GatePassNo, 
    m.DT, 
    m.PersonName, 
    m.VehicleNo, 
    m.Description, 
    m.UserName, 
    m.MachineName, 
    m.DTEntry, 
    CAST(CONVERT(VARCHAR(50), m.DT, 6) AS DATETIME) AS OnlyDT, 
    m.EmpID, 
    m.GatePassType, 
    m.DeptID, 
    m.VendID, 
    vm.AccTitle AS Name, 
    m.VendNameManual,
    m.VM_AccNo,
    ISNULL(m.Approved_HOD, 0) AS Approved_HOD,
    ISNULL(m.Approved, 0) AS Approved,
    ISNULL(m.Approved_Finance, 0) AS Approved_Finance
FROM dbo.MiscEntries_GateOut m
LEFT OUTER JOIN dbo.VVendersAndMakers vm ON m.VM_AccNo = vm.AccNo;
GO
